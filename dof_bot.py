import asyncio
import sqlite3
import json
import os
import io
import logging
from datetime import datetime, date, timedelta
from typing import Optional

from dotenv import load_dotenv
import xlrd
import openpyxl

import aiohttp
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message, CallbackQuery, Document,
    InlineKeyboardMarkup, InlineKeyboardButton,
    ReplyKeyboardMarkup, KeyboardButton,
)

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

# ════════════════════════════════════════════════════════
#  НАСТРОЙКИ
# ════════════════════════════════════════════════════════
BOT_TOKEN         = os.getenv("BOT_TOKEN")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
DB_PATH           = "dof_balance.db"

# Маппинг имен конвейеров из XLS в БД ключи
CONV_MAP = {
    "Конвейер 4": "kv4",      "Конвейер 4Д": "kv4d",
    "Конвейер 14": "kv14",    "Конвейер 32": "kv32",
    "Конвейер 34": "kv34",    "Конвейер 34А": "kv34a",
    "Конвейер 102": "kv102",  "Конвейер 24ПП": "kv24p",
    "Конвейер 24ХВ": "kv24hv","Конвейер 28А.1": "kv28a1",
    
    "Конвейер 3": "kv3",      "Конвейer 3Д": "kv3d", "Конвейер 3Д": "kv3d",
    "Конвейер 15": "kv15",    "Конвейер 19": "kv19",
    "Конвейер 31": "kv31",    "Конвейер 33": "kv33",
    "Конвейер 101": "kv101",  "Конвейер 28А.2": "kv28a2",
    
    "Конвейер 44": "kv44",    "Конвейер 46": "kv46",
    "Конвейер 74Д": "kv74",   "Конвейер 74": "kv74"
}

# ════════════════════════════════════════════════════════
#  БАЗА ДАННЫХ
# ════════════════════════════════════════════════════════
def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS daily_data (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                report_date TEXT,
                day_num  INTEGER,
                year     INTEGER,
                month    INTEGER,
                kv4 REAL DEFAULT 0, kv4d REAL DEFAULT 0, kv14 REAL DEFAULT 0, kv32 REAL DEFAULT 0,
                kv34 REAL DEFAULT 0, kv34a REAL DEFAULT 0, kv102 REAL DEFAULT 0, kv24p REAL DEFAULT 0,
                kv24hv REAL DEFAULT 0, kv28a1 REAL DEFAULT 0, kv3 REAL DEFAULT 0, kv3d REAL DEFAULT 0,
                kv15 REAL DEFAULT 0, kv19 REAL DEFAULT 0, kv31 REAL DEFAULT 0, kv33 REAL DEFAULT 0,
                kv101 REAL DEFAULT 0, kv28a2 REAL DEFAULT 0, kv44 REAL DEFAULT 0, kv46 REAL DEFAULT 0,
                kv74 REAL DEFAULT 0,
                source   TEXT,
                is_complete INTEGER DEFAULT 1,  -- 1 = полные сутки, 0 = только 1 смена (неполный день)
                uploaded TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(year, month, day_num)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS report_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, period TEXT, rows_saved INTEGER, user_id INTEGER, uploaded TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

def db_save_daily(parsed: dict, user_id: int, filename: str, year: int, month: int):
    """Сохраняет суточные данные. Последний день файла гарантированно помечается как неполный."""
    saved = 0
    FIELDS = [
        "kv4","kv4d","kv14","kv32","kv34","kv34a","kv102","kv24p",
        "kv24hv","kv28a1","kv3","kv3d","kv15","kv19","kv31","kv33",
        "kv101","kv28a2","kv44","kv46","kv74"
    ]
    
    base_fields = ["year", "month", "day_num", "report_date", "source", "is_complete"] + FIELDS
    cols = ",".join(base_fields)
    qs   = ",".join(["?"] * len(base_fields))
    upd  = ",".join(f"{f}=excluded.{f}" for f in FIELDS) + ", is_complete=excluded.is_complete"
    sql_query = f"INSERT INTO daily_data ({cols}) VALUES ({qs}) ON CONFLICT(year,month,day_num) DO UPDATE SET {upd}"

    all_days = list(parsed["daily"].keys())
    max_day = max(all_days) if all_days else 0

    with sqlite3.connect(DB_PATH) as conn:
        for day_num, vals in parsed["daily"].items():
            if not vals: 
                continue
            
            # Последний день в документе — всегда неполные сутки (is_complete = 0)
            is_complete = 0 if day_num == max_day else 1
            
            rec = {
                "year": year, "month": month, "day_num": day_num,
                "report_date": f"{year:04d}-{month:02d}-{day_num:02d}",
                "source": filename,
                "is_complete": is_complete
            }
            for f in FIELDS:
                rec[f] = vals.get(f, 0.0)

            try:
                conn.execute(sql_query, list(rec.values()))
                saved += 1
            except Exception as e:
                logging.warning(f"Ошибка сохранения дня {day_num}: {e}")

        conn.execute(
            "INSERT INTO report_log (filename, period, rows_saved, user_id) VALUES (?, ?, ?, ?)",
            (filename, parsed.get("period", ""), saved, user_id)
        )
        conn.commit()
    return saved, max_day

def db_get_month_data(year: int, month: int) -> list:
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cursor = conn.execute("SELECT * FROM daily_data WHERE year=? AND month=? ORDER BY day_num ASC", (year, month))
        return cursor.fetchall()

# ════════════════════════════════════════════════════════
#  ПАРСЕР EXCEL
# ════════════════════════════════════════════════════════
def _safe(v):
    try: return float(v) if v is not None else 0.0
    except (ValueError, TypeError): return 0.0

def _read_xls_rows(file_bytes: bytes) -> tuple:
    wb_xls = xlrd.open_workbook(file_contents=file_bytes)
    sheet_names = wb_xls.sheet_names()
    used_sheet = sheet_names[0]
    for sh in ["ДетТекМесяц", "ТекМесяц", "ДетСменТекМесяц"]:
        if sh in sheet_names:
            used_sheet = sh
            break
    ws_xls = wb_xls.sheet_by_name(used_sheet)
    rows = []
    for r_idx in range(ws_xls.nrows):
        rows.append([ws_xls.cell_value(r_idx, c_idx) for c_idx in range(ws_xls.ncols)])
    totals_sheet = "ТекМесяц" if "ТекМесяц" in sheet_names else used_sheet
    ws2_xls = wb_xls.sheet_by_name(totals_sheet)
    rows2 = []
    for r_idx in range(ws2_xls.nrows):
        rows2.append([ws2_xls.cell_value(r_idx, c_idx) for c_idx in range(ws2_xls.ncols)])
    return rows, rows2, sheet_names, used_sheet, totals_sheet

def _read_xlsx_rows(file_bytes: bytes) -> tuple:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    used_sheet = ""
    ws = None
    for sh in ["ДетТекМесяц", "ТекМесяц", "ДетСменТекМесяц"]:
        if sh in sheet_names:
            ws = wb[sh]
            used_sheet = sh
            break
    if ws is None:
        ws = wb[sheet_names[0]]
        used_sheet = sheet_names[0]
    rows = list(ws.iter_rows(values_only=True))
    totals_sheet = "ТекМесяц" if "ТекМесяц" in sheet_names else used_sheet
    ws2 = wb[totals_sheet]
    rows2 = list(ws2.iter_rows(values_only=True))
    return rows, rows2, sheet_names, used_sheet, totals_sheet

def parse_report(file_bytes: bytes, filename: str) -> dict:
    result = {"period": "", "month_label": "", "daily": {}, "totals": {}, "sheets_found": []}
    try:
        if filename.lower().endswith(".xls"):
            rows, rows2, sheet_names, used_sheet, totals_sheet = _read_xls_rows(file_bytes)
        else:
            rows, rows2, sheet_names, used_sheet, totals_sheet = _read_xlsx_rows(file_bytes)
        result["sheets_found"] = sheet_names
        result["used_sheet"] = used_sheet
        result["month_label"] = "Текущий месяц" if "Тек" in totals_sheet else "Предыдущий месяц"
    except Exception as err:
        logging.error(f"Ошибка чтения Excel: {err}")
        return result

    date_row_idx = None
    for i, row in enumerate(rows):
        if row and str(row[0]).strip() in ("Дата:", "Дата"):
            date_row_idx = i
            break
    if date_row_idx is None: return result

    date_row = rows[date_row_idx]
    days = []
    for v in date_row[1:]:
        try:
            if v is not None: days.append(int(float(v)))
        except (ValueError, TypeError): pass

    daily = {d: {} for d in days}
    for row in rows[date_row_idx+1:]:
        if not row or not row[0]: continue
        name = str(row[0]).strip()
        key = CONV_MAP.get(name) or CONV_MAP.get(name.rstrip())
        if not key: continue
        for col_i, day in enumerate(days, start=1):
            if col_i < len(row): daily[day][key] = _safe(row[col_i])

    result["daily"] = daily
    totals = {}
    period_str = ""
    for row in rows2:
        if not row: continue
        name = str(row[0]).strip()
        if name == "Период:":
            period_str = str(row[1]).strip() if len(row) > 1 else ""
            continue
        key = CONV_MAP.get(name) or CONV_MAP.get(name.rstrip())
        if key and len(row) > 1:
            v = _safe(row[1])
            if v > 0: totals[key] = v

    result["totals"] = totals
    result["period"] = period_str
    return result

# ════════════════════════════════════════════════════════
#  АНАЛИТИКА ТЕХНОЛОГИИ
# ════════════════════════════════════════════════════════
def analyze_day(r: sqlite3.Row) -> dict:
    res = {"day": r["day_num"], "alerts": [], "b1_diff": 0.0, "b2_diff": 0.0, "ok": True}
    
    if r["kv4"] == 0 and r["kv4d"] == 0:
        res["alerts"].append("🔧 Секция 1: Остановка конвейеров (ремонт)")
    else:
        if r["kv4"] > 0:
            diff_p = abs(r["kv4"] - r["kv4d"]) / r["kv4"] * 100
            res["b1_diff"] = r["kv4"] - r["kv4d"]
            if diff_p > 1.5:
                res["alerts"].append(f"⚠️ Расхождение секции 1 (Кв4/4Д): {diff_p:.1f}% (Разница: {res['b1_diff']:.1f} т)")
        if r["kv4"] > 0 and r["kv14"] > 0:
            p = (r["kv14"] / r["kv4"]) * 100
            if p < 60 or p > 80: res["alerts"].append(f"🛑 Дозирование Секция 1: {p:.1f}% (Норма 60-80%)")
        if r["kv4"] > 0 and r["kv102"] > 0:
            p = (r["kv102"] / r["kv4"]) * 100
            if p < 6 or p > 20: res["alerts"].append(f"📉 Нагрузка Секция 1: {p:.1f}% (Норма 6-20%)")

    if r["kv3"] == 0 and r["kv3d"] == 0:
        res["alerts"].append("🔧 Секция 2: Остановка конвейеров (ремонт)")
    else:
        if r["kv3"] > 0:
            diff_p = abs(r["kv3"] - r["kv3d"]) / r["kv3"] * 100
            res["b2_diff"] = r["kv3"] - r["kv3d"]
            if diff_p > 1.5:
                res["alerts"].append(f"⚠️ Расхождение секции 2 (Кв3/3Д): {diff_p:.1f}% (Разница: {res['b2_diff']:.1f} т)")
        if r["kv3"] > 0 and r["kv15"] > 0:
            p = (r["kv15"] / r["kv3"]) * 100
            if p < 60 or p > 80: res["alerts"].append(f"🛑 Дозирование Секция 2: {p:.1f}% (Норма 60-80%)")
        if r["kv3"] > 0 and r["kv101"] > 0:
            p = (r["kv101"] / r["kv3"]) * 100
            if p < 6 or p > 20: res["alerts"].append(f"📉 Нагрузка Секция 2: {p:.1f}% (Норма 6-20%)")

    if any(sym in "".join(res["alerts"]) for sym in ["⚠️", "🛑", "📉"]):
        res["ok"] = False
    return res

# ════════════════════════════════════════════════════════
#  AI АССИСТЕНТ
# ════════════════════════════════════════════════════════
async def ask_ai(prompt: str, context: str) -> str:
    if not ANTHROPIC_API_KEY: return "Ошибка: API ключ не задан."
    url = "https://api.anthropic.com/v1/messages"
    headers = {"x-api-key": ANTHROPIC_API_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"}
    payload = {
        "model": "claude-3-5-sonnet-20241022", "max_tokens": 1000,
        "system": "Ты — ведущий инженер-технолог ДОФ. Анализируй логи весов фабрики и давай рекомендации. Отвечай кратко на русском.",
        "messages": [{"role": "user", "content": f"Контекст:\n{context}\n\nВопрос: {prompt}"}]
    }
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(url, headers=headers, json=payload) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    return data["content"][0]["text"]
                return f"Ошибка AI (Код {resp.status})"
    except Exception as e: return f"Сбой связи с AI: {e}"

def make_ai_context(rows: list) -> str:
    lines = ["Логи нарушений ДОФ за полные сутки:"]
    for r in rows:
        if r["is_complete"] == 1:
            an = analyze_day(r)
            if not an["ok"]:
                lines.append(f"- День {r['day_num']}: " + ", ".join(an["alerts"]))
    return "\n".join(lines)

# ════════════════════════════════════════════════════════
#  TELEGRAM ОБРАБОТЧИКИ
# ════════════════════════════════════════════════════════
storage = MemoryStorage()
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=storage)

class AIState(StatesGroup): waiting_for_question = State()

def main_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📅 Суточный баланс"), KeyboardButton(text="⚡ Отчёт за 1 смену")],
            [KeyboardButton(text="🗓 Месячный итог"), KeyboardButton(text="🔔 Алерты фабрики")],
            [KeyboardButton(text="📆 Недельная сводка"), KeyboardButton(text="🤖 Задать вопрос AI")],
            [KeyboardButton(text="❓ Помощь")]
        ], resize_keyboard=True
    )

@dp.message(Command("start"))
async def cmd_start(msg: Message):
    init_db()
    await msg.answer("🏭 *Система весового контроля ДОФ готова!* \nЗагрузите актуальный файл отчёта весов.", parse_mode="Markdown", reply_markup=main_keyboard())

@dp.message(F.document)
async def handle_report(msg: Message):
    doc = msg.document
    if not (doc.file_name.lower().endswith(".xls") or doc.file_name.lower().endswith(".xlsx")):
        return await msg.answer("❌ Формат должен быть .xls или .xlsx")

    wait_msg = await msg.answer("⏳ _Обработка данных..._", parse_mode="Markdown")
    try:
        file_obj = await bot.get_file(doc.file_id)
        file_io = io.BytesIO()
        await bot.download_file(file_obj.file_path, file_io)
        parsed = parse_report(file_io.getvalue(), doc.file_name)
        
        if not parsed or not parsed["daily"]:
            return await wait_msg.edit_text("❌ Не удалось считать таблицы отчёта.")
            
        now = datetime.now()
        saved_days, current_incomplete_day = db_save_daily(parsed, msg.from_user.id, doc.file_name, now.year, now.month)
        
        # Точное число полных дней из файла
        total_full_days = current_incomplete_day - 1 if current_incomplete_day > 0 else 0
        
        await wait_msg.edit_text(
            f"✅ *Отчёт успешно загружен!*\n\n"
            f"📊 Анализ по листу: `{parsed['used_sheet']}`\n"
            f"📆 Записано полных суток в базу: *{total_full_days}* (с 1 по {total_full_days} число)\n"
            f"⚡ Выделена текущая неполная смена: *День {current_incomplete_day}*\n\n"
            f"Используйте меню для получения раздельных отчетов.",
            parse_mode="Markdown"
        )
    except Exception as e:
        logging.error(f"Сбой: {e}")
        await wait_msg.edit_text(f"❌ Сбой обработки: {e}")

# ── КНОПКА: СУТОЧНЫЙ БАЛАНС (СТРОГО ЗА ПОСЛЕДНИЕ ПОЛНЫЕ СУТКИ) ──
@dp.message(F.text == "📅 Суточный баланс")
async def report_daily(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    complete_rows = [r for r in rows if r["is_complete"] == 1]
    
    if not complete_rows:
        return await msg.answer("❌ В базе нет завершённых суток.")
        
    last_full_day = complete_rows[-1]
    diff = last_full_day["kv4"] - last_full_day["kv4d"]
    diff2 = last_full_day["kv3"] - last_full_day["kv3d"]
    
    text = (
        f"📅 *Суточный отчет за {last_full_day['day_num']:02d}.{now.month:02d}.{now.year} (Полные закрытые сутки):*\n\n"
        f"🔹 *Секция 1 (Измельчение):*\n"
        f"  • Конв. 4 (Питание): `{last_full_day['kv4']:.1f}` т\n"
        f"  • Конв. 4Д (Дублир): `{last_full_day['kv4d']:.1f}` т\n"
        f"  • Погрешность весов: `{diff:+.1f}` т\n"
        f"  • ... 14 (Дозирование): `{last_full_day['kv14']:.1f}` т\n\n"
        f"🔸 *Секция 2 (Измельчение):*\n"
        f"  • Конв. 3 (Питание): `{last_full_day['kv3']:.1f}` т\n"
        f"  • Конв. 3Д (Дублир): `{last_full_day['kv3d']:.1f}` т\n"
        f"  • Погрешность весов: `{diff2:+.1f}` т\n"
        f"  • ... 15 (Дозирование): `{last_full_day['kv15']:.1f}` т"
    )
    await msg.answer(text, parse_mode="Markdown")

# ── КНОПКА: ОТЧЕТ ЗА 1 СМЕНУ ТЕКУЩЕГО НЕПОЛНОГО ДНЯ ──
@dp.message(F.text == "⚡ Отчёт за 1 смену")
async def report_incomplete_shift(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    incomplete_rows = [r for r in rows if r["is_complete"] == 0]
    
    if not incomplete_rows:
        return await msg.answer("📋 Все суточные данные закрыты. Неполных промежуточных смен не найдено.")
        
    inc = incomplete_rows[0]
    text = (
        f"⚡ *Промежуточный отчёт по 1-й смене за {inc['day_num']:02d}.{now.month:02d}:*\n"
        f"_(Данные за 12 часов выгрузки, исключены из месячных итогов)_\n\n"
        f"🏭 *Секция 1:* Кв4: `{inc['kv4']:.1f}` т | Кв4Д: `{inc['kv4d']:.1f}` т\n"
        f"💊 Дозирование Кв14: `{inc['kv14']:.1f}` т\n"
        f"🔄 Циркуляция Кв102: `{inc['kv102']:.1f}` т\n\n"
        f"🏭 *Секция 2:* Кв3: `{inc['kv3']:.1f}` т | Кв3Д: `{inc['kv3d']:.1f}` т\n"
        f"💊 Дозирование Кв15: `{inc['kv15']:.1f}` т\n"
        f"🔄 Циркуляция Кв101: `{inc['kv101']:.1f}` т"
    )
    await msg.answer(text, parse_mode="Markdown")

# ── КНОПКА: МЕСЯЧНЫЙ ИТОГ (СТРОГО ЗА ПОЛНЫЕ СУТКИ) ──
@dp.message(F.text == "🗓 Месячный итог")
async def report_monthly(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    full_days = [r for r in rows if r["is_complete"] == 1]
    
    if not full_days:
        return await msg.answer("❌ Нет полных данных для расчета месячного итога.")
        
    t_kv4 = sum(r["kv4"] for r in full_days)
    t_kv3 = sum(r["kv3"] for r in full_days)
    t_kv14 = sum(r["kv14"] for r in full_days)
    t_kv15 = sum(r["kv15"] for r in full_days)
    
    text = (
        f"🗓 *Накопленный итог за месяц (Всего полных суток: {len(full_days)}):*\n"
        f"_(Текущий незавершённый день полностью исключён из подсчёта)_\n\n"
        f"🔹 *Секция 1 (Измельчение):*\n"
        f"  • Всего питания (Кв4): `{t_kv4:,.1f}` т\n"
        f"  • Всего дозирования (Кв14): `{t_kv14:,.1f}` т\n\n"
        f"🔸 *Секция 2 (Измельчение):*\n"
        f"  • Всего питания (Кв3): `{t_kv3:,.1f}` т\n"
        f"  • Всего дозирования (Кв15): `{t_kv15:,.1f}` т"
    )
    await msg.answer(text, parse_mode="Markdown")

@dp.message(F.text == "📆 Недельная сводка")
async def report_weekly(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    full_days = [r for r in rows if r["is_complete"] == 1]
    if not full_days: return await msg.answer("❌ Данные отсутствуют.")
    
    text = "📆 *Сводка по полным суткам (последние 7 дней):*\n\n"
    for r in full_days[-7:]:
        text += f"▪️ *День {r['day_num']}:* Секц1: `{r['kv4']:.1f}` т | Секц2: `{r['kv3']:.1f}` т\n"
    await msg.answer(text, parse_mode="Markdown")

@dp.message(F.text == "🔔 Алерты фабрики")
async def report_alerts(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    full_days = [r for r in rows if r["is_complete"] == 1]
    if not full_days: return await msg.answer("❌ Нет полных данных.")
        
    text = "🔔 *Технологические нарушения (только по закрытым суткам):*\n\n"
    found = False
    for r in full_days:
        analysis = analyze_day(r)
        if analysis["alerts"] and not analysis["ok"]:
            found = True
            text += f"📅 *День {r['day_num']}:*\n" + "\n".join(f"  {a}" for a in analysis["alerts"]) + "\n\n"
    if not found: text += "✅ По всем полным суткам нарушений норм не обнаружено!"
    await msg.answer(text[:4000], parse_mode="Markdown")

@dp.message(F.text == "🤖 Задать вопрос AI")
async def ai_request(msg: Message, state: FSMContext):
    await state.set_state(AIState.waiting_for_question)
    await msg.answer("🤖 *Режим AI-Технолога ДОФ*\nВведите ваш вопрос по работе фабрики:")

@dp.message(AIState.waiting_for_question)
async def ai_processing(msg: Message, state: FSMContext):
    await state.clear()
    wait = await msg.answer("🧠 _ИИ строит отчет..._")
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    ctx = make_ai_context(rows) if rows else "Нет полных данных."
    answer = await ask_ai(msg.text, ctx)
    await wait.edit_text(f"🤖 *AI-Агент:*\n\n{answer[:4000]}")

@dp.message(F.text == "❓ Помощь")
async def h_help(msg: Message):
    await msg.answer("📖 *Инструкция:* \nСкиньте файл весов. Последний день отчёта автоматически помечается как неполный (1-я смена) и не включается в месячную статистику.")

async def to_start():
    init_db()
    logging.info("ДОФ Баланс Bot запущен.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(to_start())