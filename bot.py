"""
Telegram-бот для отчётов по таблице GREEN TASHKENT (только чтение).

Функции:
  /start, /menu      — главное меню (кнопки)
  📊 Сводка           — итоги по месяцам и категориям
  💵 Курс USD         — последний курс из листа «Курс валют»
  📦 Остаток          — остаток арматуры по диаметрам
  📞 Контакты         — поиск контакта по имени/должности
  🔄 Обновить таблицу — прислать боту новый .xlsx, бот заменит файл

Установка:
  pip install aiogram openpyxl

Запуск:
  1) Положите GREEN_TASHKENT.xlsx рядом с этим файлом (или обновите через бота).
  2) Вставьте токен из @BotFather в BOT_TOKEN.
  3) python bot.py
"""

import asyncio
import os
from datetime import datetime, timedelta

import openpyxl
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
)

# ──────────────────────────── НАСТРОЙКИ ────────────────────────────
# Токен берётся из переменной окружения BOT_TOKEN (так настроено на хостинге).
# Для локального запуска можно вписать токен прямо в кавычки ниже.
BOT_TOKEN = os.getenv("BOT_TOKEN", "ВАШ_ТОКЕН_ЗДЕСЬ")

# Папка для данных. На хостинге сюда подключается постоянный диск (volume),
# чтобы обновлённый через Telegram файл не пропадал при перезапуске.
# Задаётся переменной окружения DATA_DIR (например, /data на Railway).
# Локально по умолчанию — папка рядом с bot.py.
DATA_DIR = os.getenv("DATA_DIR", os.path.dirname(__file__))
EXCEL_FILE = os.path.join(DATA_DIR, "GREEN_TASHKENT.xlsx")

# Если файла ещё нет в DATA_DIR (первый запуск на чистом volume),
# но он лежит рядом с кодом — копируем его в DATA_DIR.
_bundled = os.path.join(os.path.dirname(__file__), "GREEN_TASHKENT.xlsx")
if not os.path.exists(EXCEL_FILE) and os.path.exists(_bundled) and _bundled != EXCEL_FILE:
    import shutil
    os.makedirs(DATA_DIR, exist_ok=True)
    shutil.copy(_bundled, EXCEL_FILE)

# Ограничение доступа: впишите Telegram ID через переменную ALLOWED_USERS
# (числа через запятую, узнать ID: @userinfobot). Пусто = доступ всем.
_allowed_env = os.getenv("ALLOWED_USERS", "").strip()
ALLOWED_USERS: list[int] = (
    [int(x) for x in _allowed_env.replace(" ", "").split(",") if x]
    if _allowed_env else []
)

# ──────────────────────────── БОТ ────────────────────────────
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())


class Form(StatesGroup):
    contact_query = State()
    waiting_file = State()


def allowed(user_id: int) -> bool:
    return not ALLOWED_USERS or user_id in ALLOWED_USERS


def main_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📊 Сводка"), KeyboardButton(text="💵 Курс USD")],
            [KeyboardButton(text="📦 Остаток"), KeyboardButton(text="📞 Контакты")],
            [KeyboardButton(text="🔄 Обновить таблицу")],
        ],
        resize_keyboard=True,
    )


def fmt_num(value) -> str:
    """Число с пробелами-разделителями тысяч."""
    try:
        n = float(value)
    except (TypeError, ValueError):
        return str(value)
    if n == int(n):
        return f"{int(n):,}".replace(",", " ")
    return f"{n:,.2f}".replace(",", " ")


def load_wb():
    if not os.path.exists(EXCEL_FILE):
        return None
    return openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)


# ──────────────────────────── ОТЧЁТЫ ────────────────────────────
SKIP_CATS = {"Арматура тоннаж"}  # это объём (тонны), не деньги


def _num(v):
    """Число или None (отсекаем текст, ошибки вроде #DIV/0!)."""
    return float(v) if isinstance(v, (int, float)) else None


def _read_summary_data():
    """Читает лист «Итого Месяцев»: список месяцев, суммы и доллары по категориям.

    Возвращает (months, sums, usd):
      months — список названий месяцев
      sums   — {категория: [значения по месяцам в сумах]}
      usd    — {категория: [значения по месяцам в долларах]}
    """
    wb = load_wb()
    if wb is None:
        return None
    ws = wb["Итого Месяцев"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    months = [c for c in rows[1][1:] if c]
    n = len(months)

    def parse_block(start_idx):
        """Читает категории от строки start_idx до строки «Итого»."""
        data = {}
        for row in rows[start_idx:]:
            cat = row[0]
            if cat is None or str(cat).strip() == "":
                continue
            if str(cat).strip().lower() == "итого":
                break
            if cat in SKIP_CATS:
                continue
            data[cat] = [_num(v) for v in row[1:1 + n]]
        return data

    # Блок сумов начинается со строки 2 (индекс 2 = третья строка листа).
    sums = parse_block(2)
    # Блок долларов — второй такой же блок ниже. Ищем вторую шапку «Категория».
    usd = {}
    for i, row in enumerate(rows):
        if i > 1 and row and str(row[0]).strip() == "Категория":
            usd = parse_block(i + 1)
            break

    return months, sums, usd


def get_month_list():
    """Список месяцев для кнопок (только те, где есть хоть какие-то расходы)."""
    data = _read_summary_data()
    if not data:
        return []
    months, sums, _ = data
    active = []
    for idx, m in enumerate(months):
        total = 0.0
        for vals in sums.values():
            v = vals[idx] if idx < len(vals) else None
            if v is not None:
                total += v
        if total:
            active.append((idx, m))
    return active


def report_summary(month_idx=None) -> str:
    """Сводка. month_idx=None — весь период; иначе конкретный месяц."""
    data = _read_summary_data()
    if not data:
        return "⚠️ Файл таблицы не найден. Обновите его через «🔄 Обновить таблицу»."
    months, sums, usd = data

    def agg(store, idx):
        """Сумма по категории: за месяц idx или за весь период (idx=None)."""
        out = {}
        for cat, vals in store.items():
            if idx is None:
                s = sum(v for v in vals if v is not None)
            else:
                s = vals[idx] if idx < len(vals) and vals[idx] is not None else 0
            out[cat] = s
        return out

    sum_tot = agg(sums, month_idx)
    usd_tot = agg(usd, month_idx)
    grand_sum = sum(sum_tot.values())
    grand_usd = sum(usd_tot.values())

    if month_idx is None:
        header = f"📊 *СВОДКА — ВЕСЬ ПЕРИОД*\n{months[0]} — {months[-1]}"
    else:
        header = f"📊 *СВОДКА — {months[month_idx]}*"

    lines = [header, ""]
    lines.append("*Расходы по категориям:*")
    for cat, s in sorted(sum_tot.items(), key=lambda x: -x[1]):
        if s == 0:
            continue
        share = (s / grand_sum * 100) if grand_sum else 0
        d = usd_tot.get(cat, 0)
        d_str = f" / ${fmt_num(d)}" if d else ""
        lines.append(f"• {cat}: {fmt_num(s)} сум{d_str}  ({share:.1f}%)")

    if grand_sum == 0:
        lines.append("_За этот месяц расходов нет._")

    lines.append("")
    lines.append(f"*ИТОГО: {fmt_num(grand_sum)} сум*")
    if grand_usd:
        lines.append(f"*≈ ${fmt_num(grand_usd)}*")
    return "\n".join(lines)


def report_rate() -> str:
    wb = load_wb()
    if wb is None:
        return "⚠️ Файл таблицы не найден."
    ws = wb["Курс валют"]
    last = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is not None and row[3] is not None:
            last = row
    wb.close()
    if not last:
        return "Курс не найден."

    date_raw, cur, qty, rate = last[0], last[1], last[2], last[3]
    # Excel-дата (число) → реальная дата
    if isinstance(date_raw, (int, float)):
        date_str = (datetime(1899, 12, 30) + timedelta(days=int(date_raw))).strftime("%d.%m.%Y")
    elif isinstance(date_raw, datetime):
        date_str = date_raw.strftime("%d.%m.%Y")
    else:
        date_str = str(date_raw)

    return (f"💵 *Курс {cur}*\n\n"
            f"Дата: {date_str}\n"
            f"1 {cur} = *{fmt_num(rate)} сум*")


def report_stock() -> str:
    wb = load_wb()
    if wb is None:
        return "⚠️ Файл таблицы не найден."
    ws = wb["Остаток"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Находим строку заголовков (где первая ячейка — «Диаметр (D)»).
    head_idx = None
    for i, row in enumerate(rows):
        if row and str(row[0]).strip().startswith("Диаметр"):
            head_idx = i
            break
    if head_idx is None:
        return "Данные об остатке не найдены."

    headers = rows[head_idx]
    # Названия колонок (берём как есть из файла).
    col1 = str(headers[1]).strip() if len(headers) > 1 and headers[1] else "Приход"
    col2 = str(headers[2]).strip() if len(headers) > 2 and headers[2] else None
    col3 = str(headers[3]).strip() if len(headers) > 3 and headers[3] else None

    # Решаем, тонны это или штуки: если значения дробные и небольшие — тонны.
    sample = []
    for row in rows[head_idx + 1:]:
        if row and isinstance(row[1], (int, float)):
            sample.append(row[1])
    is_tonnes = bool(sample) and all(v < 100000 for v in sample) and \
        any(v != int(v) for v in sample)
    unit = " т" if is_tonnes else ""

    def f(v):
        if not isinstance(v, (int, float)):
            return "—"
        return f"{fmt_num(v)}{unit}"

    lines = ["📦 *ОСТАТОК АРМАТУРЫ*", ""]
    total_row = None
    for row in rows[head_idx + 1:]:
        if not row or row[0] in (None, ""):
            continue
        d = str(row[0]).strip()
        if d.lower() in ("итого", "всего"):
            total_row = row
            continue
        parts = [f"• Ø {d}: {f(row[1])}"]
        # вторую/третью колонку добавляем, только если в них есть ненулевое значение
        if col2 and isinstance(row[2], (int, float)) and row[2]:
            parts.append(f"{col2}: {f(row[2])}")
        if col3 and isinstance(row[3], (int, float)) and row[3]:
            parts.append(f"{col3}: {f(row[3])}")
        lines.append("  |  ".join(parts))

    if total_row is not None:
        lines.append("")
        lines.append(f"*ИТОГО: {f(total_row[1])}*")

    if len(lines) <= 2:
        return "Данные об остатке не найдены."
    return "\n".join(lines)


def search_contacts(query: str) -> str:
    wb = load_wb()
    if wb is None:
        return "⚠️ Файл таблицы не найден."
    ws = wb["Контакты"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    q = query.strip().lower()
    found = []
    for row in rows:
        if not row or row[0] in (None, "№", "") or str(row[0]).startswith("📞"):
            continue
        name = str(row[1] or "")
        role = str(row[2] or "")
        phone = str(row[3] or "")
        org = str(row[4] or "")
        haystack = f"{name} {role} {org}".lower()
        if q in haystack:
            found.append(f"👤 *{name.strip()}*\n💼 {role}\n📱 {phone}\n🏢 {org}")
    if not found:
        return f"По запросу «{query}» ничего не найдено."
    return "📞 *Найдено:*\n\n" + "\n\n".join(found[:15])


# ──────────────────────────── ХЕНДЛЕРЫ ────────────────────────────
@dp.message(Command("start"))
@dp.message(Command("menu"))
async def start_handler(message: types.Message, state: FSMContext):
    await state.clear()
    if not allowed(message.from_user.id):
        await message.answer("⛔ Доступ ограничен.")
        return
    await message.answer(
        "👷 *Бот отчётов — GREEN TASHKENT*\n\nВыберите раздел:",
        reply_markup=main_kb(),
        parse_mode="Markdown",
    )


def summary_months_kb() -> InlineKeyboardMarkup:
    """Кнопки выбора месяца + «Весь период»."""
    rows = []
    months = get_month_list()
    row = []
    for idx, name in months:
        row.append(InlineKeyboardButton(text=name, callback_data=f"sum:{idx}"))
        if len(row) == 2:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    rows.append([InlineKeyboardButton(text="📈 Весь период (итого)",
                                      callback_data="sum:all")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@dp.message(F.text == "📊 Сводка")
async def h_summary(message: types.Message):
    if not allowed(message.from_user.id):
        return
    kb = summary_months_kb()
    if not kb.inline_keyboard or len(kb.inline_keyboard) == 1:
        await message.answer(report_summary(), parse_mode="Markdown")
        return
    await message.answer("Выберите месяц или весь период:", reply_markup=kb)


@dp.callback_query(F.data.startswith("sum:"))
async def cb_summary(callback: types.CallbackQuery):
    if not allowed(callback.from_user.id):
        await callback.answer()
        return
    key = callback.data.split(":", 1)[1]
    idx = None if key == "all" else int(key)
    await callback.message.answer(report_summary(idx), parse_mode="Markdown")
    await callback.answer()


@dp.message(F.text == "💵 Курс USD")
async def h_rate(message: types.Message):
    if not allowed(message.from_user.id):
        return
    await message.answer(report_rate(), parse_mode="Markdown")


@dp.message(F.text == "📦 Остаток")
async def h_stock(message: types.Message):
    if not allowed(message.from_user.id):
        return
    await message.answer(report_stock(), parse_mode="Markdown")


@dp.message(F.text == "📞 Контакты")
async def h_contacts(message: types.Message, state: FSMContext):
    if not allowed(message.from_user.id):
        return
    await state.set_state(Form.contact_query)
    await message.answer("Введите имя, должность или организацию для поиска:")


@dp.message(Form.contact_query)
async def h_contacts_query(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(search_contacts(message.text), parse_mode="Markdown",
                         reply_markup=main_kb())


@dp.message(F.text == "🔄 Обновить таблицу")
async def h_update(message: types.Message, state: FSMContext):
    if not allowed(message.from_user.id):
        return
    await state.set_state(Form.waiting_file)
    await message.answer("Пришлите новый файл *GREEN_TASHKENT.xlsx* как документ.",
                         parse_mode="Markdown")


@dp.message(Form.waiting_file, F.document)
async def h_receive_file(message: types.Message, state: FSMContext):
    await state.clear()
    doc = message.document
    if not doc.file_name.lower().endswith(".xlsx"):
        await message.answer("❌ Это не .xlsx файл.", reply_markup=main_kb())
        return
    file = await bot.get_file(doc.file_id)
    await bot.download_file(file.file_path, EXCEL_FILE)
    await message.answer("✅ Таблица обновлена.", reply_markup=main_kb())


@dp.message(Form.waiting_file)
async def h_receive_file_wrong(message: types.Message):
    await message.answer("Жду файл .xlsx как документ. Или /menu для отмены.")


@dp.message()
async def fallback(message: types.Message):
    if not allowed(message.from_user.id):
        return
    await message.answer("Выберите раздел в меню 👇", reply_markup=main_kb())


# ──────────────────────────── ЗАПУСК ────────────────────────────
async def main():
    print("Бот запущен.")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
