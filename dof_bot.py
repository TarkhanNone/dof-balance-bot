import asyncio
import sqlite3
import json
import os
import io
import logging
import tempfile
from datetime import datetime, date, timedelta
from typing import Optional
import xlrd
import openpyxl

# Подключаем библиотеки для секретов (.env) и Excel
from dotenv import load_dotenv
from openpyxl import load_workbook

import aiohttp
import openpyxl
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
# Узнать свой ID: написать @userinfobot
USERS = {
    123456789: "chief",     # ← замени на свой ID
    987654321: "master",
    111222333: "operator",
}

# ════════════════════════════════════════════════════════
#  НОРМЫ БАЛАНСА (% от Конв.3 или Конв.4)
# ════════════════════════════════════════════════════════
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

# ════════════════════════════════════════════════════════
#  ПАРСЕР XLS/XLSX ОТЧЁТА
# ════════════════════════════════════════════════════════
# Структура листа ТекМесяц / ПредМесяц:
#   row0: заголовок
#   row1: "Дата:" + числа дней
#   row2+: строки с названием группы и данными по дням
#   Далее: итоги "Конвейер X: NNNNN"

GROUP_ROWS = {
    # название в файле → ключ нашей БД (группа)
    "3+4":                    "kv3_4",
    "101+102+24ХВ":           "kv_hv",
    "33+34":                  "kv33_34",
    "24ПП":                   "kv24p_gr",
    "28А.1+28А.2":            "kv28a_gr",
    "(44+46+74+44Д+46Д+74Д)/2": "kv_mmc",
    "65ЦПО+66ЦПО+84ЦПО":     "kv_cpo",
    "65МПС+66МПС+84МПС":     "kv_mps",
}

# Названия индивидуальных конвейеров → ключи БД
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

def _safe(v):
    try: return float(v) if v is not None else 0.0
    except: return 0.0

def parse_report(file_bytes: bytes, filename: str) -> dict:
    """
    Парсит XLS/XLSX отчёт конвейерных весов.
    Возвращает:
      {
        'period': 'c 1 по 15',
        'month_label': 'Текущий/Предыдущий',
        'daily': {день: {kv4:.., kv3:.., ...}},   # суточные
        'totals': {kv4:.., kv3:.., ...},           # итоги за период
        'sheets_found': [...]
      }
    """
    # Конвертируем .xls → .xlsx через LibreOffice если нужно
    if filename.lower().endswith(".xls"):
        import subprocess, shutil
        tmp = tempfile.mkdtemp()
        src = os.path.join(tmp, "report.xls")
        with open(src, "wb") as f: f.write(file_bytes)
        subprocess.run(
            ["soffice","--headless","--convert-to","xlsx",src,"--outdir",tmp],
            capture_output=True
        )
        out = os.path.join(tmp, "report.xlsx")
        if os.path.exists(out):
            with open(out,"rb") as f: file_bytes = f.read()
        shutil.rmtree(tmp, ignore_errors=True)

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)

    result = {"period":"", "month_label":"", "daily":{}, "totals":{}, "sheets_found": wb.sheetnames}

    # Читаем лист ДетТекМесяц (детально по дням, текущий)
    # или ТекМесяц если детального нет
    target_sheets = ["ДетТекМесяц", "ТекМесяц", "ДетСменТекМесяц"]
    ws = None
    used_sheet = ""
    for sh in target_sheets:
        if sh in wb.sheetnames:
            ws = wb[sh]
            used_sheet = sh
            break

    if ws is None:
        # Попробуем первый лист
        ws = wb[wb.sheetnames[0]]
        used_sheet = wb.sheetnames[0]

    rows = list(ws.iter_rows(values_only=True))

    # Найти строку с датами (row где row[0]=="Дата:")
    date_row_idx = None
    for i, row in enumerate(rows):
        if row and str(row[0]).strip() in ("Дата:", "Дата"):
            date_row_idx = i
            break

    if date_row_idx is None:
        return result

    # Дни — числа начиная с col 1
    date_row = rows[date_row_idx]
    days = []
    for v in date_row[1:]:
        try:
            d = int(v)
            days.append(d)
        except:
            pass

    # Данные конвейеров по дням
    daily = {d: {} for d in days}

    for row in rows[date_row_idx+1:]:
        if not row or not row[0]: continue
        name = str(row[0]).strip()
        key = CONV_MAP.get(name) or CONV_MAP.get(name.rstrip())
        if not key: continue
        for col_i, day in enumerate(days, start=1):
            if col_i < len(row):
                daily[day][key] = _safe(row[col_i])

    result["daily"] = daily

    # Итоги за период — ищем в листе ТекМесяц
    totals_sheet = "ТекМесяц" if "ТекМесяц" in wb.sheetnames else used_sheet
    ws2 = wb[totals_sheet]
    rows2 = list(ws2.iter_rows(values_only=True))

    totals = {}
    period_str = ""
    for row in rows2:
        if not row: continue
        # Период
        if str(row[0]).strip() == "Период:":
            period_str = str(row[1]).strip() if len(row)>1 else ""
        # Конвейер X: значение
        name = str(row[0]).strip()
        key = CONV_MAP.get(name) or CONV_MAP.get(name.rstrip())
        if key and len(row) > 1:
            v = _safe(row[1])
            if v > 0:
                totals[key] = v

    result["totals"] = totals
    result["period"] = period_str
    result["month_label"] = "Текущий месяц" if "Тек" in totals_sheet else "Предыдущий месяц"
    result["used_sheet"] = used_sheet

    return result

# ════════════════════════════════════════════════════════
#  БАЗА ДАННЫХ
# ════════════════════════════════════════════════════════
def init_db():
    conn = sqlite3.connect(DB_PATH)
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
    conn.close()

def db_save_daily(parsed: dict, user_id: int, filename: str, year: int, month: int):
    """Сохраняет суточные данные из распарсенного отчёта."""
    conn = sqlite3.connect(DB_PATH)
    saved = 0
    for day_num, vals in parsed["daily"].items():
        if not vals: continue
        # Собираем запись
        rec = {
            "year": year, "month": month, "day_num": day_num,
            "report_date": f"{year:04d}-{month:02d}-{day_num:02d}",
            "source": filename,
        }
        FIELDS = ["kv4","kv4d","kv14","kv32","kv34","kv34a","kv102","kv24p",
                  "kv24hv","kv28a1","kv3","kv3d","kv15","kv19","kv31","kv33",
                  "kv101","kv28a2","kv44","kv46","kv74"]
        for f in FIELDS:
            rec[f] = vals.get(f, 0)

        cols = ",".join(rec.keys())
        qs   = ",".join(["?"]*len(rec))
        upd  = ",".join(f"{k}=excluded.{k}" for k in rec if k not in ("year","month","day_num"))
        try:
            conn.execute(
                f"INSERT INTO daily_data ({cols}) VALUES ({qs}) "
                f"ON CONFLICT(year,month,day_num) DO UPDATE SET {upd}",
                list(rec.values())
            )
            saved += 1
        except Exception as e:
            logging.warning(f"Ошибка сохранения дня {day_num}: {e}")

    conn.execute(
        "INSERT INTO report_log (filename,period,rows_saved,user_id) VALUES (?,?,?,?)",
        (filename, parsed.get("period",""), saved, user_id)
    )
    conn.commit()
    conn.close()
    return saved

def db_get_days(year: int, month: int) -> list:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT * FROM daily_data WHERE year=? AND month=? ORDER BY day_num",
        (year, month)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def db_get_range(days_back: int) -> list:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    since = (date.today() - timedelta(days=days_back)).isoformat()
    rows = conn.execute(
        "SELECT * FROM daily_data WHERE report_date>=? ORDER BY report_date",
        (since,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def db_get_months() -> list:
    """Список доступных месяцев в БД."""
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT DISTINCT year,month FROM daily_data ORDER BY year,month"
    ).fetchall()
    conn.close()
    return [{"year":r[0],"month":r[1]} for r in rows]

# ════════════════════════════════════════════════════════
#  РАСЧЁТЫ БАЛАНСА
# ════════════════════════════════════════════════════════
MONTH_RU = {1:"Январь",2:"Февраль",3:"Март",4:"Апрель",5:"Май",6:"Июнь",
            7:"Июль",8:"Август",9:"Сентябрь",10:"Октябрь",11:"Ноябрь",12:"Декабрь"}

def sum_days(rows: list) -> dict:
    """Суммирует суточные данные."""
    if not rows: return {}
    keys = ["kv4","kv4d","kv14","kv32","kv34","kv34a","kv102","kv24p",
            "kv24hv","kv28a1","kv3","kv3d","kv15","kv19","kv31","kv33",
            "kv101","kv28a2","kv44","kv46","kv74"]
    return {k: sum(r.get(k,0) for r in rows) for k in keys}

def bal1(d):
    b = d.get("kv4",0)
    if not b: return None
    return (d.get("kv102",0)+d.get("kv34",0)+d.get("kv24p",0)+d.get("kv24hv",0)-d.get("kv28a1",0)-b)/b*100

def bal2(d):
    b = d.get("kv3",0)
    if not b: return None
    return (d.get("kv101",0)+d.get("kv33",0)-d.get("kv28a2",0)-b)/b*100

def balc1(d):
    b = d.get("kv4",0)
    if not b: return None
    return (d.get("kv102",0)+d.get("kv24hv",0)+d.get("kv24p",0)+d.get("kv32",0)-d.get("kv28a1",0)-d.get("kv14",0))/b*100

def balc2(d):
    b = d.get("kv3",0)
    if not b: return None
    return (d.get("kv101",0)+d.get("kv31",0)-d.get("kv28a2",0)-d.get("kv15",0))/b*100

def pct(v, base):
    return v/base*100 if base else 0

def check_norm(val, base, key):
    if not base or key not in NORMS: return "none", 0
    p = pct(val, base)
    mn, mx = NORMS[key][0], NORMS[key][1]
    m = (mx-mn)*0.5
    if mn<=p<=mx: return "ok", p
    if mn-m<=p<=mx+m: return "warn", p
    return "crit", p

def em_bal(v):
    if v is None: return "⬜"
    return "✅" if abs(v)<=WARN_PCT else "⚠️" if abs(v)<=CRIT_PCT else "🚨"

def em_norm(st):
    return {"ok":"✅","warn":"⚠️","crit":"🚨","none":"⬜"}.get(st,"⬜")

def sign(v, d=2):
    if v is None: return "—"
    return f"+{v:.{d}f}%" if v>=0 else f"{v:.{d}f}%"

def fmt(v):
    if not v and v!=0: return "—"
    return f"{int(v):,}".replace(",","_")

def fmt2(v):
    """Форматирование в тысячах."""
    if not v and v!=0: return "—"
    return f"{v/1000:.1f}k"

def build_alerts(d: dict, label="") -> list:
    alerts = []
    base4, base3 = d.get("kv4",0), d.get("kv3",0)
    prefix = f"[{label}] " if label else ""

    for name, val in [("Баланс1",bal1(d)),("Баланс2",bal2(d)),("БалС1",balc1(d)),("БалС2",balc2(d))]:
        if val is None: continue
        if abs(val)>CRIT_PCT: alerts.append(("crit",f"🚨 {prefix}{name}={sign(val)} крит."))
        elif abs(val)>WARN_PCT: alerts.append(("warn",f"⚠️ {prefix}{name}={sign(val)} предупр."))

    checks = [("kv14",base4),("kv32",base4),("kv34",base4),("kv34a",base4),("kv102",base4),
              ("kv15",base3),("kv19",base3),("kv31",base3),("kv33",base3),("kv101",base3)]
    for key, base in checks:
        if not base: continue
        st, p = check_norm(d.get(key,0), base, key)
        if st in ("crit","warn"):
            mn,mx,desc,_ = NORMS[key]
            em = "🚨" if st=="crit" else "⚠️"
            alerts.append((st, f"{em} {prefix}{desc}: {p:.0f}% (норма {mn}–{mx}%)"))

    if base4 and d.get("kv4d"):
        diff = abs(base4-d["kv4d"])/base4*100
        if diff>2: alerts.append(("crit",f"🚨 {prefix}Кв4 vs 4Д: {diff:.1f}%"))
        elif diff>0.5: alerts.append(("warn",f"⚠️ {prefix}Кв4 vs 4Д: {diff:.1f}%"))
    if base3 and d.get("kv3d"):
        diff = abs(base3-d["kv3d"])/base3*100
        if diff>2: alerts.append(("crit",f"🚨 {prefix}Кв3 vs 3Д: {diff:.1f}%"))
        elif diff>0.5: alerts.append(("warn",f"⚠️ {prefix}Кв3 vs 3Д: {diff:.1f}%"))

    return alerts

# ════════════════════════════════════════════════════════
#  ФОРМАТИРОВАНИЕ ОТЧЁТОВ
# ════════════════════════════════════════════════════════
def fmt_summary(d: dict, title: str) -> str:
    """Сводка по набору данных (суточному, недельному, месячному)."""
    b4 = d.get("kv4",0); b3 = d.get("kv3",0)
    b1 = bal1(d); b2 = bal2(d)
    bc1 = balc1(d); bc2 = balc2(d)

    lines = [f"📊 *{title}*\n"]

    lines += [
        "━━━ БАЛАНСЫ ━━━",
        f"{em_bal(b1)} Баланс 1   : {sign(b1)}",
        f"{em_bal(bc1)} Баланс С.1 : {sign(bc1)}",
        f"{em_bal(b2)} Баланс 2   : {sign(b2)}",
        f"{em_bal(bc2)} Баланс С.2 : {sign(bc2)}",
        "",
        "━━━ 1 ОЧЕРЕДЬ ━━━",
        f"Конв.4  : {fmt2(b4)} т  (осн.)",
        f"Конв.4Д : {fmt2(d.get('kv4d'))} т  (дубл.)",
        f"Конв.14 : {fmt2(d.get('kv14'))} т  {pct(d.get('kv14',0),b4):.0f}% {'✅' if 60<=pct(d.get('kv14',0),b4)<=80 else '⚠️'}",
        f"Конв.32 : {fmt2(d.get('kv32'))} т  {pct(d.get('kv32',0),b4):.0f}% {'✅' if 40<=pct(d.get('kv32',0),b4)<=60 else '⚠️'}",
        f"Конв.34 : {fmt2(d.get('kv34'))} т  {pct(d.get('kv34',0),b4):.0f}% {'✅' if 83<=pct(d.get('kv34',0),b4)<=87 else '⚠️'}",
        f"Конв.34А: {fmt2(d.get('kv34a'))} т  {pct(d.get('kv34a',0),b4):.0f}% {'✅' if 15<=pct(d.get('kv34a',0),b4)<=45 else '⚠️'}",
        f"Конв.102: {fmt2(d.get('kv102'))} т  {pct(d.get('kv102',0),b4):.0f}% {'✅' if 6<=pct(d.get('kv102',0),b4)<=20 else '⚠️'}",
        "",
        "━━━ 2 ОЧЕРЕДЬ ━━━",
        f"Конв.3  : {fmt2(b3)} т  (осн.)",
        f"Конв.3Д : {fmt2(d.get('kv3d'))} т  (дубл.)",
        f"Конв.15 : {fmt2(d.get('kv15'))} т  {pct(d.get('kv15',0),b3):.0f}% {'✅' if 60<=pct(d.get('kv15',0),b3)<=80 else '⚠️'}",
        f"Конв.19 : {fmt2(d.get('kv19'))} т  {pct(d.get('kv19',0),b3):.0f}% {'✅' if 40<=pct(d.get('kv19',0),b3)<=60 else '⚠️'}",
        f"Конв.31 : {fmt2(d.get('kv31'))} т  {pct(d.get('kv31',0),b3):.0f}% {'✅' if 40<=pct(d.get('kv31',0),b3)<=60 else '⚠️'}",
        f"Конв.33 : {fmt2(d.get('kv33'))} т  {pct(d.get('kv33',0),b3):.0f}% {'✅' if 83<=pct(d.get('kv33',0),b3)<=87 else '⚠️'}",
        f"Конв.101: {fmt2(d.get('kv101'))} т  {pct(d.get('kv101',0),b3):.0f}% {'✅' if 6<=pct(d.get('kv101',0),b3)<=20 else '⚠️'}",
    ]
    return "\n".join(lines)

def fmt_daily_table(rows: list) -> str:
    """Таблица по дням: день | Б1 | Б2 | Кв4 | Кв3."""
    if not rows: return "📭 Нет данных."
    lines = ["📅 *Суточные балансы*\n`День  Б1      Б2      Кв4     Кв3`"]
    for r in rows:
        b1 = bal1(r); b2 = bal2(r)
        lines.append(
            f"`{r['day_num']:>2d}   "
            f"{em_bal(b1)}{sign(b1,1):>6s}  "
            f"{em_bal(b2)}{sign(b2,1):>6s}  "
            f"{fmt2(r.get('kv4')):>6s}  "
            f"{fmt2(r.get('kv3')):>6s}`"
        )
    return "\n".join(lines)

def fmt_alerts_list(rows: list) -> str:
    all_a = []
    for r in rows:
        label = f"д.{r['day_num']}"
        for lvl, txt in build_alerts(r, label):
            all_a.append((lvl, txt))
    if not all_a:
        return "✅ Алертов нет."
    crits = [a[1] for a in all_a if a[0]=="crit"]
    warns = [a[1] for a in all_a if a[0]=="warn"]
    lines = [f"⚡ *Алерты: {len(crits)} крит. / {len(warns)} предупр.*\n"]
    if crits:
        lines.append("🚨 *Критичные:*")
        lines += crits
    if warns:
        lines.append("\n⚠️ *Предупреждения:*")
        lines += warns
    return "\n".join(lines)

def fmt_doubles(rows: list) -> str:
    """Расхождение 4/4Д и 3/3Д по дням."""
    lines = ["🔍 *Дублирующие весы (расхождение)*\n`День  Кв4 vs 4Д   Кв3 vs 3Д`"]
    for r in rows:
        b4 = r.get("kv4",0); b4d = r.get("kv4d",0)
        b3 = r.get("kv3",0); b3d = r.get("kv3d",0)
        d4 = abs(b4-b4d)/b4*100 if b4 else 0
        d3 = abs(b3-b3d)/b3*100 if b3 else 0
        e4 = "✅" if d4<=0.5 else "⚠️" if d4<=2 else "🚨"
        e3 = "✅" if d3<=0.5 else "⚠️" if d3<=2 else "🚨"
        lines.append(f"`{r['day_num']:>2d}    {e4}{d4:>5.2f}%       {e3}{d3:>5.2f}%`")
    return "\n".join(lines)

# ════════════════════════════════════════════════════════
#  AI АГЕНТ
# ════════════════════════════════════════════════════════
SYS = """Ты — AI-агент метролога горно-обогатительной фабрики (ДОФ, Казахстан).

ТЕХНОЛОГИЧЕСКАЯ СХЕМА:
1оч: Руда→Конв.4(6400т/ч)→6 бункеров(6000т)→дробление/грохочение→Конв.14(2500т/ч)→5 бункеров(1000т)→Сепарация→[ПП:18(усл)→32→34→ММС][Хвосты:20(усл)→24→скл.хв.25 ИЛИ 102→скл.хв.105][Мелкая фракция:34А→34][Скл.ПП:24→склад→28А.I→32]
2оч: Руда→Конв.3(6400т/ч)→6 бункеров→дробление→Конв.15(2500т/ч)→5 бункеров→Сепарация→[ПП:19→31→33→ММС][Хвосты:101→скл.хв.105][Скл.ПП:28А.II→31]
НЕ в балансе: Конв.18, 20

НОРМЫ: Конв.14/15: 70%±10% | 101/102: 13%±7% | 19: 50%±10% | 33/34: 83-87% | 31/32: 50%±10% | 34А: 30%±15%
БАЛАНСЫ: Б1=102+34+24П+24Хв-28А.I-4 | БС1=102+24Хв+24П+32-28А.I-14 | Б2=101+33-28А.II-3 | БС2=101+31-28А.II-15
ПОРОГИ: ±2% предупреждение, ±5% критично
ВАЖНО: задержки руды в бункерах — НОРМАЛЬНОЕ явление, объясняет почему 14/15 < 4/3

Давай конкретные метрологические рекомендации. Отвечай по-русски, кратко. Используй ✅⚠️🚨🔧📊🔮."""

async def ask_ai(question: str, data_summary: str) -> str:
    try:
        async with aiohttp.ClientSession() as s:
            r = await s.post(
                "https://api.anthropic.com/v1/messages",
                headers={"x-api-key":ANTHROPIC_API_KEY,"anthropic-version":"2023-06-01","content-type":"application/json"},
                json={"model":"claude-sonnet-4-6","max_tokens":900,"system":SYS,
                      "messages":[{"role":"user","content":question+"\n\nДанные:\n"+data_summary}]}
            )
            if r.status!=200: return f"⚠️ Ошибка API ({r.status})"
            d = await r.json()
            return d["content"][0]["text"]
    except Exception as e:
        return f"⚠️ Ошибка: {e}"

def make_ai_context(rows: list) -> str:
    """Формирует краткий контекст для AI из набора суточных данных."""
    if not rows: return "Нет данных."
    s = sum_days(rows)
    lines = [
        f"Период: {rows[0]['report_date']} — {rows[-1]['report_date']} ({len(rows)} дней)",
        f"Кв4={fmt2(s.get('kv4'))} Кв3={fmt2(s.get('kv3'))}",
        f"14={fmt2(s.get('kv14'))}({pct(s.get('kv14',0),s.get('kv4',1)):.0f}%) "
        f"15={fmt2(s.get('kv15'))}({pct(s.get('kv15',0),s.get('kv3',1)):.0f}%)",
        f"32={fmt2(s.get('kv32'))}({pct(s.get('kv32',0),s.get('kv4',1)):.0f}%) "
        f"34={fmt2(s.get('kv34'))}({pct(s.get('kv34',0),s.get('kv4',1)):.0f}%)",
        f"33={fmt2(s.get('kv33'))}({pct(s.get('kv33',0),s.get('kv3',1)):.0f}%) "
        f"101={fmt2(s.get('kv101'))}({pct(s.get('kv101',0),s.get('kv3',1)):.0f}%) "
        f"102={fmt2(s.get('kv102'))}({pct(s.get('kv102',0),s.get('kv4',1)):.0f}%)",
        f"Б1={sign(bal1(s))} БС1={sign(balc1(s))} Б2={sign(bal2(s))} БС2={sign(balc2(s))}",
    ]
    # Аномальные дни
    bad = [r for r in rows if bal1(r) and abs(bal1(r))>CRIT_PCT]
    if bad:
        lines.append(f"Критичные дни: {[r['day_num'] for r in bad]}")
    return "\n".join(lines)

# ════════════════════════════════════════════════════════
#  FSM — выбор периода для AI
# ════════════════════════════════════════════════════════
class AskAIState(StatesGroup):
    custom = State()

# ════════════════════════════════════════════════════════
#  КЛАВИАТУРЫ
# ════════════════════════════════════════════════════════
def main_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="📁 Загрузить отчёт XLS")],
        [KeyboardButton(text="📅 Суточный"), KeyboardButton(text="📆 Недельный")],
        [KeyboardButton(text="🗓 Месячный"), KeyboardButton(text="🔔 Алерты")],
        [KeyboardButton(text="🔍 Дублирование"), KeyboardButton(text="🤖 AI-анализ(временно не работает)")],
        [KeyboardButton(text="❓ Помощь")],
    ], resize_keyboard=True)

def period_kb(prefix):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📅 Сегодня",    callback_data=f"{prefix}_today"),
         InlineKeyboardButton(text="📆 7 дней",     callback_data=f"{prefix}_7d")],
        [InlineKeyboardButton(text="🗓 30 дней",    callback_data=f"{prefix}_30d"),
         InlineKeyboardButton(text="📊 Этот месяц", callback_data=f"{prefix}_month")],
    ])

def ai_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📊 Полный анализ",          callback_data="ai_full")],
        [InlineKeyboardButton(text="🔍 Диагностика отклонений", callback_data="ai_diag")],
        [InlineKeyboardButton(text="🔮 Прогноз",                callback_data="ai_forecast")],
        [InlineKeyboardButton(text="🔧 Нужна ли поверка?",      callback_data="ai_calib")],
        [InlineKeyboardButton(text="💬 Свой вопрос",            callback_data="ai_custom")],
    ])

# ════════════════════════════════════════════════════════
#  ХЭНДЛЕРЫ
# ════════════════════════════════════════════════════════
bot = Bot(token=BOT_TOKEN)
dp  = Dispatcher(storage=MemoryStorage())

def get_role(uid): return USERS.get(uid)

def require_auth(uid) -> bool:
    return True # Временно разрешаем всем. убрать после тестов и юзать реальную проверку ролей.
    #return uid in USERS Доступ по айди телеграм-аккаунта. В реальной системе нужно проверять роль и разрешения.

@dp.message(Command("start"))
async def cmd_start(msg: Message):
    if not require_auth(msg.from_user.id):
        await msg.answer("🚫 Нет доступа. Обратитесь к администратору.")
        return
    await msg.answer(
        "⚖️ *ДОФ Баланс*\n"
        "Мониторинг конвейерных весов\n\n"
        "📁 Просто скиньте файл *report.xls* — бот сам всё посчитает.\n\n"
        "Доступные отчёты:\n"
        "• 📅 Суточный — данные по дням\n"
        "• 📆 Недельный — за 7 дней\n"
        "• 🗓 Месячный — за текущий месяц\n"
        "• 🔔 Алерты — нарушения норм\n"
        "• 🤖 AI — анализ и рекомендации",
        reply_markup=main_kb(), parse_mode="Markdown"
    )

# ── ЗАГРУЗКА ФАЙЛА ────────────────────────────────────
@dp.message(F.document)
async def handle_file(msg: Message):
    if not require_auth(msg.from_user.id): return
    doc: Document = msg.document
    fn = doc.file_name or ""
    if not (fn.lower().endswith(".xls") or fn.lower().endswith(".xlsx")):
        await msg.answer("❌ Нужен файл .xls или .xlsx (отчёт конвейерных весов).")
        return

    status = await msg.answer("⏳ Читаю файл...")

    # Скачать
    file = await bot.get_file(doc.file_id)
    buf = io.BytesIO()
    await bot.download_file(file.file_path, buf)
    file_bytes = buf.getvalue()

    # Парсинг
    try:
        parsed = parse_report(file_bytes, fn)
    except Exception as e:
        await status.edit_text(f"❌ Ошибка парсинга: {e}")
        return

    if not parsed["daily"]:
        await status.edit_text(
            f"⚠️ Не удалось распознать данные.\n"
            f"Листы в файле: {', '.join(parsed['sheets_found'])}\n"
            f"Использован лист: {parsed.get('used_sheet','?')}\n\n"
            f"Ожидаемая структура: строка 'Дата:' с числами дней, затем строки конвейеров."
        )
        return

    # Определить год/месяц (спросить у пользователя или взять текущий)
    now = date.today()
    year, month = now.year, now.month
    # Если в имени файла есть дата — можно парсить, иначе текущий месяц
    saved = db_save_daily(parsed, msg.from_user.id, fn, year, month)

    # Считаем итог за загруженный период
    rows = db_get_days(year, month)
    total = sum_days(rows)
    alerts_all = []
    for r in rows:
        alerts_all.extend(build_alerts(r, f"д.{r['day_num']}"))

    crits = [a for a in alerts_all if a[0]=="crit"]
    warns = [a for a in alerts_all if a[0]=="warn"]

    text = (
        f"✅ *Файл загружен!* `{fn}`\n"
        f"Период: {parsed['period'] or 'текущий месяц'}\n"
        f"Дней сохранено: {saved}\n\n"
        f"📊 *Итог за период:*\n"
        f"Конв.4: {fmt2(total.get('kv4'))} т\n"
        f"Конв.3: {fmt2(total.get('kv3'))} т\n"
        f"{em_bal(bal1(total))} Баланс 1: {sign(bal1(total))}\n"
        f"{em_bal(bal2(total))} Баланс 2: {sign(bal2(total))}\n\n"
    )
    if crits:
        text += f"🚨 *{len(crits)} критичных алерта!*\n" + "\n".join(a[1] for a in crits[:5])
        if len(crits)>5: text += f"\n... и ещё {len(crits)-5}"
    elif warns:
        text += f"⚠️ *{len(warns)} предупреждений*\n" + "\n".join(a[1] for a in warns[:3])
    else:
        text += "✅ Все показатели в норме."

    await status.edit_text(text, parse_mode="Markdown")

# ── ПОДСКАЗКА ДЛЯ КНОПКИ ─────────────────────────────
@dp.message(F.text == "📁 Загрузить отчёт XLS")
async def h_upload_hint(msg: Message):
    await msg.answer(
        "📁 Просто *прикрепите файл* report.xls к сообщению и отправьте.\n\n"
        "Бот автоматически распознает данные из листов:\n"
        "• ДетТекМесяц — детально по дням\n"
        "• ТекМесяц — итоги за период\n\n"
        "После загрузки используйте кнопки для получения отчётов.",
        parse_mode="Markdown"
    )

# ── СУТОЧНЫЙ ─────────────────────────────────────────
@dp.message(F.text == "📅 Суточный")
async def h_daily(msg: Message):
    if not require_auth(msg.from_user.id): return
    now = date.today()
    rows = db_get_days(now.year, now.month)
    if not rows:
        await msg.answer("📭 Нет данных. Загрузите файл отчёта."); return
    await msg.answer(fmt_daily_table(rows), parse_mode="Markdown")

# ── НЕДЕЛЬНЫЙ ─────────────────────────────────────────
@dp.message(F.text == "📆 Недельный")
async def h_weekly(msg: Message):
    if not require_auth(msg.from_user.id): return
    rows = db_get_range(7)
    if not rows:
        await msg.answer("📭 Нет данных за 7 дней."); return
    total = sum_days(rows)
    title = f"Неделя ({rows[0]['report_date']} — {rows[-1]['report_date']})"
    text = fmt_summary(total, title)
    text += "\n\n" + fmt_daily_table(rows)
    await msg.answer(text[:4000], parse_mode="Markdown")

# ── МЕСЯЧНЫЙ ─────────────────────────────────────────
@dp.message(F.text == "🗓 Месячный")
async def h_monthly(msg: Message):
    if not require_auth(msg.from_user.id): return
    now = date.today()
    rows = db_get_days(now.year, now.month)
    if not rows:
        await msg.answer("📭 Нет данных за текущий месяц."); return
    total = sum_days(rows)
    title = f"{MONTH_RU[now.month]} {now.year} (дней: {len(rows)})"
    text = fmt_summary(total, title)
    await msg.answer(text, parse_mode="Markdown")

# ── АЛЕРТЫ ───────────────────────────────────────────
@dp.message(F.text == "🔔 Алерты")
async def h_alerts(msg: Message):
    if not require_auth(msg.from_user.id): return
    now = date.today()
    rows = db_get_days(now.year, now.month)
    if not rows:
        await msg.answer("📭 Нет данных."); return
    await msg.answer(fmt_alerts_list(rows)[:4000], parse_mode="Markdown")

# ── ДУБЛИРОВАНИЕ ─────────────────────────────────────
@dp.message(F.text == "🔍 Дублирование")
async def h_doubles(msg: Message):
    if not require_auth(msg.from_user.id): return
    now = date.today()
    rows = db_get_days(now.year, now.month)
    if not rows:
        await msg.answer("📭 Нет данных."); return
    await msg.answer(fmt_doubles(rows)[:4000], parse_mode="Markdown")

# ── AI-АНАЛИЗ ────────────────────────────────────────
@dp.message(F.text == "🤖 AI-анализ")
async def h_ai(msg: Message):
    if not require_auth(msg.from_user.id): return
    await msg.answer("🤖 *AI-Агент метролога*", reply_markup=ai_kb(), parse_mode="Markdown")

AI_Q = {
    "ai_full":     "Проведи полный анализ: балансы, нормы конвейеров, отклонения, рекомендации.",
    "ai_diag":     "Найди и объясни причины отклонений от норм. Что аномально?",
    "ai_forecast": "На основе тренда спрогнозируй баланс на следующий период.",
    "ai_calib":    "Оцени состояние весового оборудования. Какие весы требуют поверки?",
}

@dp.callback_query(F.data.startswith("ai_"))
async def h_ai_action(cb: CallbackQuery, state: FSMContext):
    if not require_auth(cb.from_user.id): return
    action = cb.data
    if action == "ai_custom":
        await cb.message.edit_text("💬 Введите вопрос:")
        await state.set_state(AskAIState.custom)
        return
    await cb.message.edit_text("⏳ AI анализирует...")
    now = date.today()
    rows = db_get_days(now.year, now.month)
    if not rows:
        await cb.message.edit_text("📭 Нет данных."); return
    ctx = make_ai_context(rows)
    answer = await ask_ai(AI_Q.get(action,"Проанализируй данные."), ctx)
    await cb.message.edit_text(f"🤖 *AI-Агент:*\n\n{answer[:4000]}", parse_mode="Markdown")

@dp.message(AskAIState.custom)
async def h_ai_custom(msg: Message, state: FSMContext):
    await state.clear()
    if not require_auth(msg.from_user.id): return
    wait = await msg.answer("⏳ Думаю...")
    now = date.today()
    rows = db_get_days(now.year, now.month)
    ctx = make_ai_context(rows) if rows else "Нет данных."
    answer = await ask_ai(msg.text, ctx)
    await wait.edit_text(f"🤖 *AI-Агент:*\n\n{answer[:4000]}", parse_mode="Markdown")

# ── ПОМОЩЬ ───────────────────────────────────────────
@dp.message(F.text == "❓ Помощь")
async def h_help(msg: Message):
    await msg.answer(
        "📖 *Как пользоваться:*\n\n"
        "1️⃣ Скиньте файл `report.xls` в чат\n"
        "2️⃣ Бот прочитает и сохранит все данные\n"
        "3️⃣ Нажимайте кнопки для отчётов\n\n"
        "*Кнопки:*\n"
        "📅 Суточный — таблица по дням с балансами\n"
        "📆 Недельный — сводка за 7 дней\n"
        "🗓 Месячный — итог за текущий месяц\n"
        "🔔 Алерты — все нарушения норм\n"
        "🔍 Дублирование — расхождение Кв4/4Д и Кв3/3Д\n"
        "🤖 AI — анализ, диагностика, прогноз\n\n"
        "*Нормы конвейеров:*\n"
        "Конв.14/15: `60–80%` от Конв.4/3\n"
        "Конв.101/102: `6–20%`\n"
        "Конв.19/31/32: `40–60%`\n"
        "Конв.33/34: `83–87%`\n"
        "Конв.10/34А: `15–45%`\n\n"
        "*Баланс:*\n"
        "✅ Норма: ±2% | ⚠️ Внимание: ±5% | 🚨 Критично: >5%",
        parse_mode="Markdown"
    )

# ════════════════════════════════════════════════════════
#  ЗАПУСК
# ════════════════════════════════════════════════════════
async def main():
    init_db()
    logging.info("ДОФ Баланс Bot запущен. Ожидаю файлы report.xls...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
