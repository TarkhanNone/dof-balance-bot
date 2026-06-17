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

# Восстановленные оригинальные константы норм ДОФ из dof_bot_old
NORMS = {
    "kv14":  (60, 80,  "Конв.14",      1),
    "kv15":  (60, 80,  "Конв.15",      2),
    "kv32":  (40, 60,  "Конв.32",      1),
    "kv31":  (40, 60,  "Конв.31",      2),
    "kv34":  (83, 87,  "Конв.34→ММС",  1),
    "kv33":  (83, 87,  "Конв.33→ММС",  2),
    "kv102": (6,  20,  "Конв.102 хв.", 1),
    "kv101": (6,  20,  "Конв.101 хв.", 2),
    "kv19":  (40, 60,  "Конв.19",      2),
    "kv34a": (15, 45,  "Конв.10/34А",  1),
}
WARN_PCT = 2.0
CRIT_PCT = 5.0

# Маппинг индивидуальных конвейеров из XLS в БД ключи
CONV_MAP = {
    "Конвейер 4":    "kv4",   "Конвейер 4Д":   "kv4d",
    "Конвейер 14":   "kv14",  "Конвейер 32":   "kv32",
    "Конвейер 18":   "kv18",  "Конвейер 102":  "kv102",
    "Конвейер 34 ":  "kv34",  "Конвейер 34":   "kv34",
    "Конвейер 24ПП": "kv24p", "Конвейер 24ХВ": "kv24hv",
    "Конвейер 28А.1":"kv28a1","Конвейер 34A":  "kv34a",
    "Конвейер 3":    "kv3",   "Конвейер 3Д":   "kv3d",
    "Конвейер 15":   "kv15",  "Конвейер 31":   "kv31",
    "Конвейер 19":   "kv19",  "Конвейер 101":  "kv101",
    "Конвейер 33":   "kv33",  "Конвейер 28А.2":"kv28a2",
    "Конвейер 44":   "kv44",  "Конвейер 44Д":  "kv44d",
    "Конвейер 46":   "kv46",  "Конвейер 46Д":  "kv46d",
    "Конвейер 74":   "kv74",  "Конвейер 74Д":  "kv74d",
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
    except: return 0.0

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
#  ОРИГИНАЛЬНЫЕ МАТЕМАТИЧЕСКИЕ ФОРМУЛЫ ИЗ DOF_BOT_OLD
# ════════════════════════════════════════════════════════
def bal1(d):
    b = d.get("kv4", 0)
    if not b: return None
    return (d.get("kv102",0) + d.get("kv34",0) + d.get("kv24p",0) + d.get("kv24hv",0) - d.get("kv28a1",0) - b) / b * 100

def bal2(d):
    b = d.get("kv3", 0)
    if not b: return None
    return (d.get("kv101",0) + d.get("kv33",0) - d.get("kv28a2",0) - b) / b * 100

def balc1(d):
    b = d.get("kv4", 0)
    if not b: return None
    return (d.get("kv102",0) + d.get("kv24hv",0) + d.get("kv24p",0) + d.get("kv32",0) - d.get("kv28a1",0) - d.get("kv14",0)) / b * 100

def balc2(d):
    b = d.get("kv3", 0)
    if not b: return None
    return (d.get("kv101",0) + d.get("kv31",0) - d.get("kv28a2",0) - d.get("kv15",0)) / b * 100

def pct(v, base):
    return v / base * 100 if base else 0

def check_norm(val, base, key):
    if not base or key not in NORMS: return "none", 0
    p = pct(val, base)
    mn, mx = NORMS[key][0], NORMS[key][1]
    m = (mx - mn) * 0.5
    if mn <= p <= mx: return "ok", p
    if mn - m <= p <= mx + m: return "warn", p
    return "crit", p

def em_bal(v):
    if v is None: return "⬜"
    return "✅" if abs(v) <= WARN_PCT else "⚠️" if abs(v) <= CRIT_PCT else "🚨"

def em_norm(st):
    return {"ok": "✅", "warn": "⚠️", "crit": "🚨", "none": "⬜"}.get(st, "⬜")

def sign(v, d=2):
    if v is None: return "—"
    return f"+{v:.{d}f}%" if v >= 0 else f"{v:.{d}f}%"

def fmt(v):
    if not v and v != 0: return "—"
    return f"{int(v):,}".replace(",", "_")

def build_alerts(d: dict, label="") -> list:
    alerts = []
    base4, base3 = d.get("kv4", 0), d.get("kv3", 0)
    prefix = f"[{label}] " if label else ""
    
    # Сверка Балансов
    b1, b2 = bal1(d), bal2(d)
    if b1 is not None and abs(b1) > WARN_PCT:
        alerts.append(f"{em_bal(b1)} {prefix}Техн. Баланс Секции 1: `{sign(b1)}` (Питание Кв4 `{fmt(base4)}` т)")
    if b2 is not None and abs(b2) > WARN_PCT:
        alerts.append(f"{em_bal(b2)} {prefix}Техн. Баланс Секции 2: `{sign(b2)}` (Питание Кв3 `{fmt(base3)}` т)")
        
    bc1, bc2 = balc1(d), balc2(d)
    if bc1 is not None and abs(bc1) > WARN_PCT:
        alerts.append(f"{em_bal(bc1)} {prefix}Баланс Цикла Секции 1: `{sign(bc1)}`")
    if bc2 is not None and abs(bc2) > WARN_PCT:
        alerts.append(f"{em_bal(bc2)} {prefix}Баланс Цикла Секции 2: `{sign(bc2)}`")

    # Сверка Технологических норм конвейеров питания дозирования
    for k, base in [("kv14", base4), ("kv32", base4), ("kv34", base4), ("kv34a", base4),
                    ("kv102", base4), ("kv15", base3), ("kv31", base3), ("kv33", base3),
                    ("kv101", base3), ("kv19", base3)]:
        val = d.get(k, 0)
        if val > 0 and base > 0:
            st, p = check_norm(val, base, k)
            if st in ("warn", "crit"):
                alerts.append(f"{em_norm(st)} {prefix}{NORMS[k][2]}: `{p:.1f}%` (Норма {NORMS[k][0]}-{NORMS[k][1]}%)")
    return alerts

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
    lines = ["Логи нарушений технологического процесса ДОФ за полные сутки:"]
    for r in rows:
        if r["is_complete"] == 1:
            al = build_alerts(dict(r), label=f"День {r['day_num']}")
            if al: lines.extend(al)
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

# ── КНОПКА: СУТОЧНЫЙ БАЛАНС (ПОЛНОСТЬЮ ВОССТАНОВЛЕН ИЗ DOF_BOT_OLD) ──
@dp.message(F.text == "📅 Суточный баланс")
async def report_daily(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    complete_rows = [r for r in rows if r["is_complete"] == 1]
    
    if not complete_rows:
        return await msg.answer("❌ В базе нет завершённых суток.")
        
    # СТРОГАЯ ОРИГИНАЛЬНАЯ ТАБЛИЦА С ВЫЧИСЛЕНИЕМ ТЕХНОЛОГИЧЕСКИХ ОТКЛОНЕНИЙ % ИЗ DOF_BOT_OLD
    text = f"📊 *Суточный технологический отчет ({now.month:02d}/{now.year}):*\n"
    text += "_(Учтены только полные закрытые сутки)_\n\n"
    text += "`Дн | Секц1 (Кв4/4Д) | Секц2 (Кв3/3Д) `\n"
    text += "────────────────────────────────\n"
    
    for r in complete_rows:
        d = dict(r)
        
        # Расчет погрешности Секции 1
        diff1 = d.get("kv4", 0) - d.get("kv4d", 0)
        pct1 = (diff1 / d.get("kv4", 1) * 100) if d.get("kv4", 0) > 0 else 0.0
        em1 = "✅" if abs(pct1) <= 1.5 else "⚠️" if d.get("kv4", 0) > 0 else "⬜"
        
        # Расчет погрешности Секции 2
        diff2 = d.get("kv3", 0) - d.get("kv3d", 0)
        pct2 = (diff2 / d.get("kv3", 1) * 100) if d.get("kv3", 0) > 0 else 0.0
        em2 = "✅" if abs(pct2) <= 1.5 else "⚠️" if d.get("kv3", 0) > 0 else "⬜"
        
        text += f"`{d['day_num']:02d} | {em1} {pct1:<+5.1f}% ({diff1:<+4.0f}т) | {em2} {pct2:<+5.1f}% ({diff2:<+4.0f}т)`\n"
        
    await msg.answer(text[:4000], parse_mode="Markdown")

# ── КНОПКА: ОТЧЕТ ЗА 1 СМЕНУ ТЕКУЩЕГО НЕПОЛНОГО ДНЯ ──
@dp.message(F.text == "⚡ Отчёт за 1 смену")
async def report_incomplete_shift(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    incomplete_rows = [r for r in rows if r["is_complete"] == 0]
    
    if not incomplete_rows:
        return await msg.answer("📋 Все суточные данные закрыты. Неполных промежуточных смен не найдено.")
        
    inc = dict(incomplete_rows[0])
    
    # Считаем промежуточные балансы 1-й смены по оригинальным формулам
    b1, b2 = bal1(inc), bal2(inc)
    bc1, bc2 = balc1(inc), balc2(inc)
    
    text = (
        f"⚡ *Промежуточный отчёт по 1-й смене за {inc['day_num']:02d}.{now.month:02d}:*\n"
        f"_(Данные за 12 часов выгрузки, исключены из месячных итогов)_\n\n"
        f"🏭 *Секция 1 (Питание Кв4: `{fmt(inc['kv4'])}` т):*\n"
        f"  • Расхождение Кв4/4Д: `{sign(b1)}` {em_bal(b1)}\n"
        f"  • Баланс цикла: `{sign(bc1)}` {em_bal(bc1)}\n"
        f"  • Дозирование Кв14: `{inc['kv14']:.1f}` т\n"
        f"  • Циркуляция Кв102: `{inc['kv102']:.1f}` т\n\n"
        f"🏭 *Секция 2 (Питание Кв3: `{fmt(inc['kv3'])}` т):*\n"
        f"  • Расхождение Кв3/3Д: `{sign(b2)}` {em_bal(b2)}\n"
        f"  • Баланс цикла: `{sign(bc2)}` {em_bal(bc2)}\n"
        f"  • Дозирование Кв15: `{inc['kv15']:.1f}` т\n"
        f"  • Циркуляция Кв101: `{inc['kv101']:.1f}` т"
    )
    await msg.answer(text, parse_mode="Markdown")

# ── КНОПКА: МЕСЯЧНЫЙ ИТОГ (ОРИГИНАЛЬНЫЕ ИТОГОВЫЕ ФОРМУЛЫ ИЗ DOF_BOT_OLD) ──
@dp.message(F.text == "🗓 Месячный итог")
async def report_monthly(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    full_days = [r for r in rows if r["is_complete"] == 1]
    
    if not full_days:
        return await msg.answer("❌ Нет полных данных для расчета месячного итога.")
        
    s = {k: sum(r[k] for r in full_days) for k in ["kv4","kv4d","kv14","kv32","kv34","kv34a","kv102","kv24p",
                                                   "kv24hv","kv28a1","kv3","kv3d","kv15","kv19","kv31","kv33",
                                                   "kv101","kv28a2","kv44","kv46","kv74"]}
    
    mb1, mb2 = bal1(s), bal2(s)
    mbc1, mbc2 = balc1(s), balc2(s)
    
    text = (
        f"🗓 *Накопленный баланс за месяц (Учтено суток: {len(full_days)}):*\n"
        f"_(Промежуточный день выгрузки полностью исключён из подсчёта)_\n\n"
        f"🔹 *Секция 1 (Питание Кв4: `{fmt(s['kv4'])}` т):*\n"
        f"  • Расхождение весов Кв4/4Д: `{sign(mb1)}` {em_bal(mb1)}\n"
        f"  • Баланс технологического цикла: `{sign(mbc1)}` {em_bal(mbc1)}\n"
        f"  • Итого дозирование Кв14: `{fmt(s['kv14'])}` т\n\n"
        f"🔸 *Секция 2 (Питание Кв3: `{fmt(s['kv3'])}` т):*\n"
        f"  • Расхождение весов Кв3/3Д: `{sign(mb2)}` {em_bal(mb2)}\n"
        f"  • Баланс технологического цикла: `{sign(mbc2)}` {em_bal(mbc2)}\n"
        f"  • Итого дозирование Кв15: `{fmt(s['kv15'])}` т"
    )
    await msg.answer(text, parse_mode="Markdown")

@dp.message(F.text == "%📆 Недельная сводка" or F.text == "📆 Недельная сводка")
async def report_weekly(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    full_days = [r for r in rows if r["is_complete"] == 1]
    if not full_days: return await msg.answer("❌ Данные отсутствуют.")
    
    text = "📆 *Сводка питания по закрытым суткам (последние 7 дней):*\n\n"
    for r in full_days[-7:]:
        text += f"▪️ *День {r['day_num']}:* Секция 1: `{fmt(r['kv4'])}` т | Секция 2: `{fmt(r['kv3'])}` т\n"
    await msg.answer(text, parse_mode="Markdown")

# ── КНОПКА: АЛЕРТЫ ФАБРИКИ (ОРИГИНАЛЬНАЯ ПРОВЕРКА ИЗ DOF_BOT_OLD) ──
@dp.message(F.text == "🔔 Алерты фабрики")
async def report_alerts(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    full_days = [r for r in rows if r["is_complete"] == 1]
    if not full_days: return await msg.answer("❌ Нет закрытых суток в базе.")
        
    text = "🔔 *Технологические нарушения норм ДОФ за текущий месяц:*\n\n"
    found = False
    for r in full_days:
        al = build_alerts(dict(r), label=f"День {r['day_num']}")
        if al:
            found = True
            text += "\n".join(al) + "\n\n"
            
    if not found: 
        text += "✅ Нарушений технологических балансов и норм дозирования не обнаружено!"
    await msg.answer(text[:4000], parse_mode="Markdown")

@dp.message(F.text == "🤖 Задать вопрос AI")
async def ai_request(msg: Message, state: FSMContext):
    await state.set_state(AIState.waiting_for_question)
    await msg.answer("🤖 *Режим AI-Технолога ДОФ*\nВведите ваш вопрос по работе фабрики:")

@dp.message(AIState.waiting_for_question)
async def ai_processing(msg: Message, state: FSMContext):
    await state.clear()
    wait = await msg.answer("🧠 _ИИ строит технологический анализ..._")
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