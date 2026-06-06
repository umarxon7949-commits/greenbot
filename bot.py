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
def report_summary() -> str:
    wb = load_wb()
    if wb is None:
        return "⚠️ Файл таблицы не найден. Обновите его через «🔄 Обновить таблицу»."
    ws = wb["Итого Месяцев"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    months = [c for c in rows[1][1:] if c]  # шапка месяцев
    n_months = len(months)

    # На листе два блока с одинаковыми категориями: «СУММЫ» и «ТОННАЖИ».
    # Берём только первый (деньги) — от шапки до строки «Итого».
    # «Арматура тоннаж» — это объём, не деньги, его пропускаем.
    skip = {"Арматура тоннаж"}
    totals = {}
    grand_total = 0.0
    for row in rows[2:]:
        cat = row[0]
        if cat is None or str(cat).strip() == "":
            continue
        if str(cat).strip().lower() == "итого":
            break  # конец денежного блока
        if cat in skip:
            continue
        s = sum(float(v) for v in row[1:1 + n_months]
                if isinstance(v, (int, float)))
        totals[cat] = s
        grand_total += s

    lines = ["📊 *СВОДКА ПО МЕСЯЦАМ* (сум)", ""]
    lines.append(f"Период: {months[0]} — {months[-1]}")
    lines.append("")
    lines.append("*Расходы по категориям (итого):*")
    for cat, s in sorted(totals.items(), key=lambda x: -x[1]):
        share = (s / grand_total * 100) if grand_total else 0
        lines.append(f"• {cat}: {fmt_num(s)}  ({share:.1f}%)")
    lines.append("")
    lines.append(f"*ИТОГО: {fmt_num(grand_total)} сум*")
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

    lines = ["📦 *ОСТАТОК АРМАТУРЫ*", ""]
    started = False
    for row in rows:
        if row[0] == "Диаметр (D)":
            started = True
            continue
        if started and row[0] is not None:
            d, qty = row[0], row[1]
            if d in (None, "") or str(d).strip().lower() in ("итого", "всего"):
                continue
            lines.append(f"• Ø {d}: {fmt_num(qty)} шт")
    if len(lines) == 2:
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


@dp.message(F.text == "📊 Сводка")
async def h_summary(message: types.Message):
    if not allowed(message.from_user.id):
        return
    await message.answer(report_summary(), parse_mode="Markdown")


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
