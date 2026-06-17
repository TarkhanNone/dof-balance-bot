import asyncio
import sqlite3
import json
import os
import io
import logging
from datetime import datetime, date, timedelta
from typing import Optional

# Подключаем библиотеки для секретов (.env) и Excel
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
#  НАСТРОЙКИ — ЗАПОЛНИ ПЕРЕД ЗАПУСКОМ
# ════════════════════════════════════════════════════════
BOT_TOKEN         = os.getenv("BOT_TOKEN")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
DB_PATH           = "dof_balance.db"

# Пользователи {telegram_id: роль}
# Роли: "operator" / "master" / "chief"
USERS = {
    123456789: "chief",     # ← замени на свой реальный Telegram ID
}

# Маппинг имен конвейеров из XLS в БД ключи
CONV_MAP = {
    "Конвейер 4": "kv4",      "Конвейер 4Д": "kv4d",
    "Конвейер 14": "kv14",    "Конвейер 32": "kv32",
    "Конвейер 34": "kv34",    "Конвейер 34А": "kv34a",
    "Конвейер 102": "kv102",  "Конвейер 24ПП": "kv24p",
    "Конвейер 24ХВ": "kv24hv","Конвейер 28А.1": "kv28a1",
    
    "Конвейер 3": "kv3",      "Конвейер 3Д": "kv3d",
    "Конвейер 15": "kv15",    "Конвейер 19": "kv19",
    "Конвейер 31": "kv31",    "Конвейer 33": "kv33",  # на случай опечаток в исходниках
    "Конвейер 33": "kv33",
    "Конвейер 101": "kv101",  "Конвейер 28А.2": "kv28a2",
    
    "Конвейер 44": "kv44",    "Конвейер 46": "kv46",
    "Конвейер 74Д": "kv74",   "Конвейер 74": "kv74"
}

# ════════════════════════════════════════════════════════
#  БАЗА ДАННЫХ (БЕЗОПАСНЫЙ РЕФАКТОРИНГ С WITH)
# ════════════════════════════════════════════════════════
def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS daily_data (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                report_date TEXT,        -- дата отчёта (YYYY-MM-DD)
                day_num  INTEGER,        -- день месяца (1-31)
                year     INTEGER,
                month    INTEGER,
                kv4      REAL DEFAULT 0, kv4d  REAL DEFAULT 0,
                kv14     REAL DEFAULT 0, kv32  REAL DEFAULT 0,
                kv34     REAL DEFAULT 0, kv34a REAL DEFAULT 0,
                kv102    REAL DEFAULT 0, kv24p REAL DEFAULT 0,
                kv24hv   REAL DEFAULT 0, kv28a1 REAL DEFAULT 0,
                kv3      REAL DEFAULT 0, kv3d  REAL DEFAULT 0,
                kv15     REAL DEFAULT 0, kv19  REAL DEFAULT 0,
                kv31     REAL DEFAULT 0, kv33  REAL DEFAULT 0,
                kv101    REAL DEFAULT 0, kv28a2 REAL DEFAULT 0,
                kv44     REAL DEFAULT 0, kv46  REAL DEFAULT 0,
                kv74     REAL DEFAULT 0,
                source   TEXT,
                uploaded TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(year, month, day_num)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS report_log (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT,
                period   TEXT,
                rows_saved INTEGER,
                user_id  INTEGER,
                uploaded TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

def db_save_daily(parsed: dict, user_id: int, filename: str, year: int, month: int):
    """Сохраняет суточные данные, игнорируя неполные текущие сутки"""
    saved = 0
    FIELDS = [
        "kv4","kv4d","kv14","kv32","kv34","kv34a","kv102","kv24p",
        "kv24hv","kv28a1","kv3","kv3d","kv15","kv19","kv31","kv33",
        "kv101","kv28a2","kv44","kv46","kv74"
    ]
    
    base_fields = ["year", "month", "day_num", "report_date", "source"] + FIELDS
    cols = ",".join(base_fields)
    qs   = ",".join(["?"] * len(base_fields))
    upd  = ",".join(f"{f}=excluded.{f}" for f in FIELDS)
    sql_query = f"INSERT INTO daily_data ({cols}) VALUES ({qs}) ON CONFLICT(year,month,day_num) DO UPDATE SET {upd}"

    # Текущая дата для отсечения незаконченных суток
    today = datetime.now()
    
    with sqlite3.connect(DB_PATH) as conn:
        for day_num, vals in parsed["daily"].items():
            if not vals: 
                continue
            
            # 🛡️ ОТСЕЧКА НЕПОЛНОГО ДНЯ: Если день в отчете совпадает с сегодня — не пишем его
            if year == today.year and month == today.month and day_num == today.day:
                logging.info(f"День {day_num} пропущен при записи: текущие сутки ещё не завершены.")
                continue
            
            rec = {
                "year": year, "month": month, "day_num": day_num,
                "report_date": f"{year:04d}-{month:02d}-{day_num:02d}",
                "source": filename,
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
    return saved

def db_get_month_data(year: int, month: int) -> list:
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cursor = conn.execute(
            "SELECT * FROM daily_data WHERE year=? AND month=? ORDER BY day_num ASC", 
            (year, month)
        )
        return cursor.fetchall()

# ════════════════════════════════════════════════════════
#  ПАРСЕР EXCEL (ОПТИМИЗИРОВАННЫЙ И БЕЗОПАСНЫЙ)
# ════════════════════════════════════════════════════════
def _safe(v):
    try: 
        return float(v) if v is not None else 0.0
    except (ValueError, TypeError): 
        return 0.0

def _read_xls_rows(file_bytes: bytes) -> tuple:
    """Вспомогательная функция: читает строки из старого формата .xls в ОЗУ"""
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
    """Вспомогательная функция: читает строки из нового формата .xlsx в ОЗУ"""
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
    """Парсит XLS/XLSX отчёт конвейерных весов строго в ОЗУ хостинга"""
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
        logging.error(f"Ошибка чтения структуры Excel {filename}: {err}")
        return result

    # Найти строку с датами
    date_row_idx = None
    for i, row in enumerate(rows):
        if row and str(row[0]).strip() in ("Дата:", "Дата"):
            date_row_idx = i
            break

    if date_row_idx is None:
        return result

    date_row = rows[date_row_idx]
    days = []
    for v in date_row[1:]:
        try:
            if v is not None:
                days.append(int(float(v)))
        except (ValueError, TypeError):
            pass

    daily = {d: {} for d in days}

    for row in rows[date_row_idx+1:]:
        if not row or not row[0]: 
            continue
        name = str(row[0]).strip()
        key = CONV_MAP.get(name) or CONV_MAP.get(name.rstrip())
        if not key: 
            continue
        for col_i, day in enumerate(days, start=1):
            if col_i < len(row):
                daily[day][key] = _safe(row[col_i])

    result["daily"] = daily

    # Итоги за период
    totals = {}
    period_str = ""
    for row in rows2:
        if not row: 
            continue
        name = str(row[0]).strip()
        if name == "Период:":
            period_str = str(row[1]).strip() if len(row) > 1 else ""
            continue
        key = CONV_MAP.get(name) or CONV_MAP.get(name.rstrip())
        if key and len(row) > 1:
            v = _safe(row[1])
            if v > 0:
                totals[key] = v

    result["totals"] = totals
    result["period"] = period_str
    return result

# ════════════════════════════════════════════════════════
#  БИЗНЕС-ЛОГИКА: БАЛАНСЫ, АЛЕРТЫ (ЗАЩИЩЕНО ОТ НУЛЕЙ)
# ════════════════════════════════════════════════════════
def analyze_day(r: sqlite3.Row) -> dict:
    """Вычисляет балансы и отклонения. Полностью защищена от дней ремонта конвейеров."""
    res = {
        "day": r["day_num"],
        "alerts": [],
        "b1_diff": 0.0, "b2_diff": 0.0, "b3_diff": 0.0,
        "ok": True
    }
    
    # ─── СЕКЦИЯ 1 (Ремонт или работа) ───
    if r["kv4"] == 0 and r["kv4d"] == 0:
        res["alerts"].append("🔧 Секция 1: Остановка конвейеров (плановый ремонт / простой)")
    else:
        if r["kv4"] > 0:
            diff_p = abs(r["kv4"] - r["kv4d"]) / r["kv4"] * 100
            res["b1_diff"] = r["kv4"] - r["kv4d"]
            if diff_p > 1.5:
                res["alerts"].append(f"⚠️ Расхождение секции 1 (Кв4 и Кв4Д): {diff_p:.1f}% (Разница: {res['b1_diff']:.1f} т)")
                
        if r["kv4"] > 0 and r["kv14"] > 0:
            p = (r["kv14"] / r["kv4"]) * 100
            if p < 60 or p > 80:
                res["alerts"].append(f"🛑 Дозирование Секция 1 (Кв14 от Кв4): {p:.1f}% (Норма 60-80%)")

        if r["kv4"] > 0 and r["kv102"] > 0:
            p = (r["kv102"] / r["kv4"]) * 100
            if p < 6 or p > 20:
                res["alerts"].append(f"📉 Нагрузка Секция 1 (Кв102 от Кв4): {p:.1f}% (Норма 6-20%)")

    # ─── СЕКЦИЯ 2 (Ремонт или работа) ───
    if r["kv3"] == 0 and r["kv3d"] == 0:
        res["alerts"].append("🔧 Секция 2: Остановка конвейеров (плановый ремонт / простой)")
    else:
        if r["kv3"] > 0:
            diff_p = abs(r["kv3"] - r["kv3d"]) / r["kv3"] * 100
            res["b2_diff"] = r["kv3"] - r["kv3d"]
            if diff_p > 1.5:
                res["alerts"].append(f"⚠️ Расхождение секции 2 (Кв3 и Кв3Д): {diff_p:.1f}% (Разница: {res['b2_diff']:.1f} т)")

        if r["kv3"] > 0 and r["kv15"] > 0:
            p = (r["kv15"] / r["kv3"]) * 100
            if p < 60 or p > 80:
                res["alerts"].append(f"🛑 Дозирование Секция 2 (Кв15 от Кв3): {p:.1f}% (Норма 60-80%)")

        if r["kv3"] > 0 and r["kv101"] > 0:
            p = (r["kv101"] / r["kv3"]) * 100
            if p < 6 or p > 20:
                res["alerts"].append(f"📉 Нагрузка Секция 2 (Кв101 от Кв3): {p:.1f}% (Норма 6-20%)")

    # Фильтруем: Дни с чистым ремонтом не считаются критической ошибкой технологии
    has_real_violations = any(sym in "".join(res["alerts"]) for sym in ["⚠️", "🛑", "📉"])
    if has_real_violations:
        res["ok"] = False
        
    return res

# ════════════════════════════════════════════════════════
#  ИНТЕГРАЦИЯ С AI (ANTHROPIC CLAUDE)
# ════════════════════════════════════════════════════════
async def ask_ai(prompt: str, context: str) -> str:
    if not ANTHROPIC_API_KEY:
        return "Ошибка: На сервере не задан API ключ `ANTHROPIC_API_KEY`."
        
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    payload = {
        "model": "claude-3-5-sonnet-20241022",
        "max_tokens": 1000,
        "system": (
            "Ты — ведущий инженер-технолог и эксперт АСУТП обогатительной фабрики ДОФ.\n"
            "Твоя задача — анализировать логи конвейерных весов, выявлять нарушения норм дозирования (норма 60-80%),\n"
            "проблемы циркуляции (норма 6-20%) и расхождения датчиков (более 1.5%).\n"
            "Давай краткие, профессиональные технические рекомендации обогатителям. Отвечай строго на русском."
        ),
        "messages": [
            {"role": "user", "content": f"Контекст (Данные весов фабрики):\n{context}\n\nВопрос пользователя: {prompt}"}
        ]
    }
    
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(url, headers=headers, json=payload) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    return data["content"][0]["text"]
                else:
                    err_body = await resp.text()
                    return f"Ошибка AI-сервера (Код {resp.status}): {err_body}"
    except Exception as e:
        return f"Не удалось связаться с ИИ-моделью: {e}"

def make_ai_context(rows: list) -> str:
    lines = ["Дни с технологическими нарушениями на ДОФ:"]
    for r in rows:
        analysis = analyze_day(r)
        if not analysis["ok"]:
            lines.append(f"- День {r['day_num']}:")
            for a in analysis["alerts"]:
                lines.append(f"  {a}")
    return "\n".join(lines)

# ════════════════════════════════════════════════════════
#  TELEGRAM BOT - КЛАВИАТУРА И ОБРАБОТЧИКИ (AIOGRAM v3)
# ════════════════════════════════════════════════════════
storage = MemoryStorage()
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=storage)

class AIState(StatesGroup):
    waiting_for_question = State()

def main_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📅 Суточный баланс"), KeyboardButton(text="📆 Недельная сводка")],
            [KeyboardButton(text="🗓 Месячный итог"), KeyboardButton(text="🔔 Алерты фабрики")],
            [KeyboardButton(text="🔍 Дублирование весов"), KeyboardButton(text="🤖 Задать вопрос AI")],
            [KeyboardButton(text="❓ Помощь")]
        ],
        resize_keyboard=True
    )

@dp.message(Command("start"))
async def cmd_start(msg: Message):
    init_db()
    await msg.answer(
        "🏭 *Добро пожаловать в систему весового контроля ДОФ!*\\n\\n"
        "Отправьте мне файл отчёта весов (`report.xls` или `.xlsx`), и я автоматически сохраню "
        "все данные, рассчитаю балансы, циркуляцию и выведу нарушения.",
        parse_mode="Markdown",
        reply_markup=main_keyboard()
    )

@dp.message(F.document)
async def handle_report(msg: Message):
    doc = msg.document
    if not (doc.file_name.lower().endswith(".xls") or doc.file_name.lower().endswith(".xlsx")):
        return await msg.answer("❌ Ошибка: Поддерживаются только файлы весовых отчётов `.xls` или `.xlsx`.")

    wait_msg = await msg.answer("⏳ _Считываю данные отчёта фабрики в ОЗУ..._", parse_mode="Markdown")
    
    try:
        file_obj = await bot.get_file(doc.file_id)
        file_io = io.BytesIO()
        await bot.download_file(file_obj.file_path, file_io)
        file_bytes = file_io.getvalue()
        
        parsed = parse_report(file_bytes, doc.file_name)
        
        if not parsed or not parsed["daily"]:
            return await wait_msg.edit_text("❌ Не удалось считать таблицы. Проверьте структуру файла отчета весов.")
            
        now = datetime.now()
        saved_days = db_save_daily(parsed, msg.from_user.id, doc.file_name, now.year, now.month)
        
        await wait_msg.edit_text(
            f"✅ *Файл успешно обработан на сервере!*\\n\\n"
            f"📊 Лист отчёта: `{parsed['used_sheet']}` ({parsed['month_label']})\\n"
            f"📅 Период данных: *{parsed['period']}*\\n"
            f"💾 Записано/обновлено завершённых суток в БД: `{saved_days}`\\n"
            f"_(Текущие незавершенные сутки автоматически пропущены до завтрашнего дня)_",
            parse_mode="Markdown"
        )
    except Exception as e:
        logging.error(f"Ошибка в handle_report: {e}")
        await wait_msg.edit_text(f"❌ Критический сбой обработки: {e}")

@dp.message(F.text == "📅 Суточный баланс")
async def report_daily(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    if not rows:
        return await msg.answer("❌ В базе данных нет данных за текущий месяц. Загрузите файл весов.")
        
    text = f"📊 *Суточный баланс за {now.month}/{now.year} (тонны):*\\n\\n"
    text += "`День | Конв.4  | Конв.4Д | Разница`\\n"
    text += "─────────────────────────────\\n"
    
    for r in rows:
        diff = r["kv4"] - r["kv4d"]
        text += f"`{r['day_num']:02d}   | {r['kv4']:<7.1f} | {r['kv4d']:<7.1f} | {diff:<+6.1f}`\\n"
        
    await msg.answer(text[:4000], parse_mode="Markdown")

@dp.message(F.text == "📆 Недельная сводка")
async def report_weekly(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    if not rows:
        return await msg.answer("❌ Данные отсутствуют.")
        
    last_7 = rows[-7:]
    text = "📆 *Сводка за последние 7 завершённых дней работы фабрики (т):*\\n\\n"
    for r in last_7:
        text += f"▪️ *День {r['day_num']}:* Питание Секц1: `{r['kv4']:.1f}` т, Секц2: `{r['kv3']:.1f}` т\\n"
    await msg.answer(text, parse_mode="Markdown")

@dp.message(F.text == "🗓 Месячный итог")
async def report_monthly(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    if not rows:
        return await msg.answer("❌ Нет данных.")
        
    t_kv4 = sum(r["kv4"] for r in rows)
    t_kv3 = sum(r["kv3"] for r in rows)
    t_kv14 = sum(r["kv14"] for r in rows)
    t_kv15 = sum(r["kv15"] for r in rows)
    
    text = (
        f"🗓 *Накопленный итог за месяц ({now.month}/{now.year}):*\\n\\n"
        f"🔹 *Секция 1 (Измельчение):*\\n"
        f"  • Всего переработано (Кв4): `{t_kv4:,.1f}` тонн\\n"
        f"  • Дозирование (Кв14): `{t_kv14:,.1f}` тонн\\n\\n"
        f"🔸 *Секция 2 (Измельчение):*\\n"
        f"  • Всего переработано (Кв3): `{t_kv3:,.1f}` тонн\\n"
        f"  • Дозирование (Кв15): `{t_kv15:,.1f}` тонн"
    )
    await msg.answer(text, parse_mode="Markdown")

@dp.message(F.text == "🔔 Алерты фабрики")
async def report_alerts(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    if not rows:
        return await msg.answer("❌ Нет данных.")
        
    text = "🔔 *Технологические сводки и нарушения норм ДОФ за месяц:*\\n\\n"
    found = False
    
    for r in rows:
        analysis = analyze_day(r)
        if analysis["alerts"]:
            found = True
            text += f"📅 *День {r['day_num']}:*\\n"
            for alert in analysis["alerts"]:
                text += f"  {alert}\\n"
            text += "\\n"
            
    if not found:
        text += "✅ Все показатели дозирования, циркуляции и весов в пределах технологических норм!"
        
    await msg.answer(text[:4000], parse_mode="Markdown")

@dp.message(F.text == "🔍 Дублирование весов")
async def report_duplication(msg: Message):
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    if not rows:
        return await msg.answer("❌ Нет данных.")
        
    text = "🔍 *Анализ дублирующих весов (Допустимо < 1.5%):*\\n\\n"
    for r in rows:
        an = analyze_day(r)
        if r["kv4"] > 0 or r["kv3"] > 0:
            text += f"▪️ *День {r['day_num']}:* Разница Секц1: `{an['b1_diff']:.1f}` т | Секц2: `{an['b2_diff']:.1f}` т\\n"
    await msg.answer(text[:4000], parse_mode="Markdown")

@dp.message(F.text == "🤖 Задать вопрос AI")
async def ai_request(msg: Message, state: FSMContext):
    await state.set_state(AIState.waiting_for_question)
    await msg.answer(
        "🤖 *Режим AI-Ассистента АСУТП/Технолога*\\n\\n"
        "Я проанализирую все нарушения и суточные балансы фабрики за текущий месяц.\\n"
        "Введите ваш вопрос. Например:\\n"
        "_— Сделай краткую диагностику работы секций и выдели критические дни_"\\n"
        "_— Какая секция дозирования питания работала стабильнее в этом месяце?_",
        parse_mode="Markdown"
    )

@dp.message(AIState.waiting_for_question)
async def ai_processing(msg: Message, state: FSMContext):
    await state.clear()
    wait = await msg.answer("🧠 _ИИ-Агент изучает логи базы данных daily_data и строит отчет..._", parse_mode="Markdown")
    
    now = datetime.now()
    rows = db_get_month_data(now.year, now.month)
    ctx = make_ai_context(rows) if rows else "Нет данных за этот месяц."
    
    answer = await ask_ai(msg.text, ctx)
    await wait.edit_text(f"🤖 *AI-Агент:*\\n\\n{answer[:4000]}", parse_mode="Markdown")

@dp.message(F.text == "❓ Помощь")
async def h_help(msg: Message):
    await msg.answer(
        "📖 *Как пользоваться ботом:*\\n\\n"
        "1️⃣ Скиньте файл `report.xls` или `.xlsx` в чат.\\n"
        "2️⃣ Бот прочитает его в ОЗУ и обновит сохранённые дни.\\n"
        "3️⃣ Нажимайте встроенные кнопки меню для выгрузки аналитики.\\n\\n"
        "*Установленные нормы ДОФ:*\\n"
        "• Конв.14/15 (дозирование): `60–80%` от Конв.4/3\\n"
        "• Конв.101/102 (циркулирующая нагрузка): `6–20%`\\n"
        "• Погрешность дублирования датчиков: `<= 1.5%`\\n"
        "• _Текущие незавершенные сутки отсекаются автоматически, чтобы не искажать итоги._",
        parse_mode="Markdown"
    )

# ════════════════════════════════════════════════════════
#  ТОЧКА ВХОДА В ПРИЛОЖЕНИЕ
# ════════════════════════════════════════════════════════
async def to_start():
    init_db()
    logging.info("ДОФ Баланс Bot запущен. Ожидаю файлы весов...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(to_start())