import asyncio
import csv
import hmac
import io
import logging
import os
import random
import re
import sqlite3
from contextlib import asynccontextmanager
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
from aiogram import Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    BufferedInputFile, InlineKeyboardButton, InlineKeyboardMarkup,
    FSInputFile, CallbackQuery, Message, PollAnswer, Update
)
from fastapi import FastAPI, Request, HTTPException

# ============================ SETTINGS ============================

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "0"))
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "theStudentSupportZone")
TELEGRAM_WEBHOOK_SECRET = os.environ.get("TELEGRAM_WEBHOOK_SECRET", "hartron_deo_secret_99")

MONTHLY_PLAN_PRICE = 59
MONTHLY_PLAN_DAYS = 30
TWO_MONTH_PLAN_PRICE = 99
TWO_MONTH_PLAN_DAYS = 60
THREE_MONTH_PLAN_PRICE = 129
THREE_MONTH_PLAN_DAYS = 90

QUESTIONS_PER_TEST = 30
MIX_QUESTIONS_PER_TEST = 50
DEFAULT_TIMER_SECONDS = 30
EXCEL_FILE = "questions.xlsx"
QR_IMAGE_FILE = "qr.jpg"
DATABASE_FILE = "quiz_bot.db"

SHUFFLE_QUESTIONS = True

# ==================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s"
)
log = logging.getLogger("hartron_deo")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, EXCEL_FILE)
QR_PATH = os.path.join(BASE_DIR, QR_IMAGE_FILE)
DB_PATH = os.path.join(BASE_DIR, DATABASE_FILE)

bot = Bot(
    token=BOT_TOKEN,
    default=DefaultBotProperties(parse_mode="HTML")
)
dp = Dispatcher()

TOPIC_ALIASES = {
    "Windows": ["windows", "win"],
    "Internet": ["internet", "web", "browser"],
    "HTML/DHTML": ["html", "dhtml", "html/asp"],
    "MS Office": ["ms-office", "ms office", "office"],
    "Computer Fundamental": ["cmpfund", "computer fundamental", "fundamental", "basics", "cpu", "memory", "hardware"],
    "Word/Excel": ["word/xcl", "word", "excel", "spreadsheet"],
    "Networking": ["networking", "lan", "wan", "protocol", "network"],
}
MIX_TOPIC_NAME = "Mix Questions"
TOPICS = list(TOPIC_ALIASES.keys()) + [MIX_TOPIC_NAME]

user_sessions = {}
poll_to_user = {}
daily_poll_to_user = {}
last_wrong_answers = {}
last_bot_menu_msg = {}
test_catalog = {}


# ------------------------------- HELPERS -------------------------------

async def safe_delete(chat_id: int, message_id: Optional[int]):
    if not message_id:
        return
    try:
        await bot.delete_message(chat_id=chat_id, message_id=message_id)
    except Exception:
        pass


async def send_clean_menu(chat_id: int, user_id: int, text: str, reply_markup=None, photo_path: Optional[str] = None):
    """Deletes previous menu message for the user before sending a new one."""
    old_msg_id = last_bot_menu_msg.pop(user_id, None)
    await safe_delete(chat_id, old_msg_id)

    if photo_path and os.path.isfile(photo_path):
        sent = await bot.send_photo(chat_id=chat_id, photo=FSInputFile(photo_path), caption=text, reply_markup=reply_markup)
    else:
        sent = await bot.send_message(chat_id=chat_id, text=text, reply_markup=reply_markup)
    last_bot_menu_msg[user_id] = sent.message_id
    return sent


# ------------------------------- DATABASE -------------------------------

def connect_db():
    return sqlite3.connect(DB_PATH)


def init_db():
    with connect_db() as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                username TEXT,
                full_name TEXT,
                expiry_date TEXT
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS payment_requests (
                request_id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                username TEXT,
                full_name TEXT,
                plan_days INTEGER NOT NULL,
                plan_price INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT NOT NULL
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS results (
                result_id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                topic TEXT NOT NULL,
                test_no INTEGER NOT NULL,
                score INTEGER NOT NULL,
                total INTEGER NOT NULL,
                percentage REAL NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        con.execute("""
            CREATE TABLE IF NOT EXISTS daily_attempts (
                user_id INTEGER NOT NULL,
                attempt_date TEXT NOT NULL,
                PRIMARY KEY (user_id, attempt_date)
            )
        """)
        con.execute("CREATE INDEX IF NOT EXISTS idx_results_user ON results(user_id)")

        user_columns = {row[1] for row in con.execute("PRAGMA table_info(users)").fetchall()}
        if "full_name" not in user_columns:
            con.execute("ALTER TABLE users ADD COLUMN full_name TEXT")


def register_user(user_id: int, username: str, full_name: str):
    with connect_db() as con:
        con.execute("""
            INSERT INTO users(user_id, username, full_name, expiry_date)
            VALUES(?,?,?,NULL)
            ON CONFLICT(user_id) DO UPDATE SET
                username=excluded.username,
                full_name=excluded.full_name
        """, (user_id, username, full_name))


def get_expiry(user_id: int) -> Optional[datetime]:
    if user_id == ADMIN_ID:
        return datetime.max
    with connect_db() as con:
        row = con.execute(
            "SELECT expiry_date FROM users WHERE user_id=?", (user_id,)
        ).fetchone()
    if not row or not row[0]:
        return None
    try:
        return datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def is_subscribed(user_id: int) -> bool:
    expiry = get_expiry(user_id)
    return bool(expiry and datetime.now() < expiry)


def can_access_test(user_id: int, test_no: int) -> bool:
    return test_no == 1 or is_subscribed(user_id)


def activate_subscription(user_id: int, username: str, full_name: str, days: int):
    old_expiry = get_expiry(user_id)
    start_date = old_expiry if old_expiry and old_expiry > datetime.now() else datetime.now()
    expiry = start_date + timedelta(days=days)
    expiry_text = expiry.strftime("%Y-%m-%d %H:%M:%S")
    with connect_db() as con:
        con.execute("""
            INSERT INTO users(user_id, username, full_name, expiry_date)
            VALUES(?,?,?,?)
            ON CONFLICT(user_id) DO UPDATE SET
                username=excluded.username,
                full_name=excluded.full_name,
                expiry_date=excluded.expiry_date
        """, (user_id, username, full_name, expiry_text))
    return expiry


def get_best_score(user_id: int, topic: str, test_no: int) -> float:
    with connect_db() as con:
        row = con.execute("""
            SELECT MAX(percentage) FROM results
            WHERE user_id=? AND topic=? AND test_no=?
        """, (user_id, topic, test_no)).fetchone()
    return float(row[0] or 0)


def save_result(user_id: int, topic: str, test_no: int, score: int, total: int):
    percentage = round((score / total) * 100, 1) if total else 0
    with connect_db() as con:
        con.execute("""
            INSERT INTO results(user_id, topic, test_no, score, total, percentage, created_at)
            VALUES(?,?,?,?,?,?,?)
        """, (
            user_id, topic, test_no, score, total, percentage,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))
    return percentage


# ------------------------------- EXCEL QUESTION BANK -------------------------------

def normalize_columns(df):
    df.columns = [
        re.sub(r"\s+", "_", str(col).strip().lower())
        for col in df.columns
    ]
    aliases = {
        "topic": "questype",
        "subject": "questype",
        "option_1": "opt_1",
        "option_2": "opt_2",
        "option_3": "opt_3",
        "option_4": "opt_4",
        "answer": "ans",
        "correct_answer": "ans",
        "test": "test_no",
        "test_number": "test_no",
    }
    return df.rename(columns={k: v for k, v in aliases.items() if k in df.columns})


def load_question_bank():
    if not os.path.isfile(EXCEL_PATH):
        log.error("Excel question bank not found: %s", EXCEL_PATH)
        return pd.DataFrame()

    try:
        df = normalize_columns(pd.read_excel(EXCEL_PATH))
        required = {"questype", "question", "opt_1", "opt_2", "opt_3", "opt_4", "ans"}
        missing = required - set(df.columns)
        if missing:
            log.error("Missing Excel columns: %s", ", ".join(sorted(missing)))
            return pd.DataFrame()

        if "test_no" not in df.columns:
            df["test_no"] = None

        df = df.dropna(subset=["questype", "question", "opt_1", "opt_2", "opt_3", "opt_4", "ans"]).copy()

        answer_text = df["ans"].astype(str).str.strip().str.upper()
        numeric_answers = pd.to_numeric(answer_text, errors="coerce")
        valid_answer = numeric_answers.isin([1, 2, 3, 4]) | answer_text.isin(["A", "B", "C", "D"])
        removed = int((~valid_answer).sum())
        if removed:
            log.warning("Excluded %s row(s) with invalid answer keys.", removed)
        df = df[valid_answer]
        return df.reset_index(drop=True)
    except Exception:
        log.exception("Failed to read Excel question bank")
        return pd.DataFrame()


def build_test_catalog(df):
    catalog = {topic: {} for topic in TOPICS}
    if df.empty:
        return catalog

    for topic, aliases in TOPIC_ALIASES.items():
        pattern = "|".join(re.escape(x) for x in aliases)
        topic_df = df[
            df["questype"].astype(str).str.contains(pattern, case=False, na=False)
        ].copy()

        if topic_df.empty:
            continue

        test_values = pd.to_numeric(topic_df["test_no"], errors="coerce")
        if test_values.notna().any():
            topic_df["_test_no"] = test_values
            for number, group in topic_df.dropna(subset=["_test_no"]).groupby("_test_no"):
                catalog[topic][int(number)] = group.drop(columns=["_test_no"]).to_dict("records")
        else:
            records = topic_df.to_dict("records")
            for offset in range(0, len(records), QUESTIONS_PER_TEST):
                number = offset // QUESTIONS_PER_TEST + 1
                catalog[topic][number] = records[offset:offset + QUESTIONS_PER_TEST]

    all_records = df.to_dict("records")
    rng = random.Random(42)
    rng.shuffle(all_records)
    for offset in range(0, len(all_records), MIX_QUESTIONS_PER_TEST):
        number = offset // MIX_QUESTIONS_PER_TEST + 1
        catalog[MIX_TOPIC_NAME][number] = all_records[offset:offset + MIX_QUESTIONS_PER_TEST]

    return catalog


QUESTION_DF = load_question_bank()
test_catalog = build_test_catalog(QUESTION_DF)


def reload_question_bank():
    global QUESTION_DF, test_catalog
    QUESTION_DF = load_question_bank()
    test_catalog = build_test_catalog(QUESTION_DF)
    return sum(len(tests) for tests in test_catalog.values())


# ------------------------------- KEYBOARD HELPERS -------------------------------

def button(label, callback_data):
    return InlineKeyboardButton(text=label, callback_data=callback_data)


def home_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [button("👨‍💻  DATA ENTRY OPERATOR (DEO)", "home:deo")],
        [
            button("📅  DAILY QUIZ", "home:daily"),
            button("🏆  TOP RANKERS", "home:rankers"),
        ],
        [button("💎  SUBSCRIPTION PLANS", "home:plans")],
        [
            button("👤  MY ACCOUNT", "home:account"),
            button("📞  SUPPORT", "home:contact"),
        ],
    ])


def topics_keyboard():
    rows = []
    for index, topic in enumerate(TOPICS):
        count = len(test_catalog.get(topic, {}))
        icon = "🔀" if topic == MIX_TOPIC_NAME else "📚"
        rows.append([button(f"{icon}  {topic}  •  {count} TESTS", f"topic:{index}")])
    rows.append([button("⬅️  MAIN MENU", "home:main")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def tests_keyboard(topic_index: int, user_id: int):
    topic = TOPICS[topic_index]
    tests = test_catalog.get(topic, {})
    subscribed = is_subscribed(user_id)
    rows = []
    for test_no in sorted(tests):
        score = get_best_score(user_id, topic, test_no)
        q_count = len(tests[test_no])
        if subscribed:
            label = f"📝  TEST SERIES {test_no} ({q_count} Qs)  •  BEST {score:.0f}%"
        elif test_no == 1:
            label = f"🆓  TEST SERIES 1 ({q_count} Qs)  •  FREE DEMO"
        else:
            label = f"🔒  TEST SERIES {test_no} ({q_count} Qs)  •  PREMIUM"
        rows.append([button(label, f"test:{topic_index}:{test_no}")])
    if not rows:
        rows.append([button("⚠️  TESTS NOT AVAILABLE", "noop")])
    rows.append([button("⬅️  BACK TO SUBJECTS", "home:deo")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def test_detail_keyboard(topic_index: int, test_no: int):
    return InlineKeyboardMarkup(inline_keyboard=[
        [button("📘  PRACTICE MODE", f"practice:{topic_index}:{test_no}")],
        [button("🚀  START CBT EXAM", f"timer:{topic_index}:{test_no}")],
        [button("⬅️  BACK TO TEST SERIES", f"topic:{topic_index}")],
    ])


def timer_keyboard(topic_index: int, test_no: int, practice=False):
    prefix = "practicego" if practice else "quizgo"
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            button("15 SEC", f"{prefix}:{topic_index}:{test_no}:15"),
            button("20 SEC", f"{prefix}:{topic_index}:{test_no}:20"),
            button("30 SEC", f"{prefix}:{topic_index}:{test_no}:30"),
        ],
        [button("⬅️  BACK", f"test:{topic_index}:{test_no}")],
    ])


def plans_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [button(f"🥉  1 MONTH — ₹{MONTHLY_PLAN_PRICE}", f"plan:{MONTHLY_PLAN_DAYS}:{MONTHLY_PLAN_PRICE}")],
        [button(f"🥈  2 MONTHS — ₹{TWO_MONTH_PLAN_PRICE} (SAVE ₹19)", f"plan:{TWO_MONTH_PLAN_DAYS}:{TWO_MONTH_PLAN_PRICE}")],
        [button(f"🥇  3 MONTHS — ₹{THREE_MONTH_PLAN_PRICE} (BEST VALUE)", f"plan:{THREE_MONTH_PLAN_DAYS}:{THREE_MONTH_PLAN_PRICE}")],
        [button("⬅️  MAIN MENU", "home:main")],
    ])


def payment_review_keyboard(request_id: int, user_id: int):
    return InlineKeyboardMarkup(inline_keyboard=[[
        button("✅ APPROVE", f"approve:{request_id}:{user_id}"),
        button("❌ REJECT", f"reject:{request_id}:{user_id}"),
    ]])


# ------------------------------- SCREEN TEXT -------------------------------

def premium_benefits():
    return (
        "💎 <b>PREMIUM MEMBERSHIP BENEFITS</b>\n\n"
        "✅ Unlock every DEO subject & 50-question Mix Test Series\n"
        "✅ Access complete syllabus question bank\n"
        "✅ Unlimited test attempts during membership\n"
        "✅ Topic-wise practice and CBT exam mode\n"
        "✅ Choose 15, 20 or 30 seconds per question\n"
        "✅ Instant score, accuracy & wrong answer review\n"
        "✅ Personal best-score & leaderboard ranking\n"
        "✅ 30, 60, or 90-day access options"
    )


async def show_plans(chat_id: int, user_id: int):
    text = (
        "💎 <b>CHOOSE YOUR PREMIUM PLAN</b>\n\n"
        f"🥉 <b>1 MONTH — ₹{MONTHLY_PLAN_PRICE}</b> (30 Days)\n"
        f"🥈 <b>2 MONTHS — ₹{TWO_MONTH_PLAN_PRICE}</b> (60 Days • Save ₹19)\n"
        f"🥇 <b>3 MONTHS — ₹{THREE_MONTH_PLAN_PRICE}</b> (90 Days • Save ₹48)\n\n"
        + premium_benefits()
        + "\n\nSelect a plan below to continue with payment."
    )
    await send_clean_menu(chat_id, user_id, text, reply_markup=plans_keyboard(), photo_path=QR_PATH)


def main_menu_text():
    return (
        "🎓 <b>WELCOME TO HARTRON DEO PREMIUM CBT PORTAL</b>\n\n"
        "Your focused preparation platform for the Data Entry Operator exam.\n\n"
        "🆓 <b>Test Series 1 is FREE</b> in every subject!\n"
        "📅 Free Daily Question & Live Rankers Leaderboard\n"
        "📚 Subject-wise & 50-question Mix test series\n"
        "⏱ Timed CBT examination & wrong answer review\n\n"
        "<b>Select an option to get started:</b>"
    )


# ------------------------------- START & HOME MENU -------------------------------

@dp.message(CommandStart())
async def start_handler(message: Message):
    user = message.from_user
    register_user(user.id, user.username or "", user.full_name)
    await safe_delete(message.chat.id, message.message_id)
    await send_clean_menu(message.chat.id, user.id, main_menu_text(), reply_markup=home_keyboard())


@dp.callback_query(F.data == "home:main")
async def home_main(call: CallbackQuery):
    await safe_delete(call.message.chat.id, call.message.message_id)
    await send_clean_menu(call.message.chat.id, call.from_user.id, main_menu_text(), reply_markup=home_keyboard())
    await call.answer()


@dp.callback_query(F.data == "home:deo")
async def home_deo(call: CallbackQuery):
    text = (
        "📚 <b>DATA ENTRY OPERATOR — SUBJECTS</b>\n\n"
        "Choose a subject (or Mix Questions) to start.\n"
        "🆓 <b>Test Series 1 is FREE</b> for all students!"
    )
    if call.message.photo:
        await safe_delete(call.message.chat.id, call.message.message_id)
        await send_clean_menu(call.message.chat.id, call.from_user.id, text, reply_markup=topics_keyboard())
    else:
        await call.message.edit_text(text, reply_markup=topics_keyboard())
        last_bot_menu_msg[call.from_user.id] = call.message.message_id
    await call.answer()


@dp.callback_query(F.data == "home:plans")
async def home_plans(call: CallbackQuery):
    await safe_delete(call.message.chat.id, call.message.message_id)
    await show_plans(call.message.chat.id, call.from_user.id)
    await call.answer()


@dp.callback_query(F.data == "home:rankers")
async def home_rankers(call: CallbackQuery):
    with connect_db() as con:
        rows = con.execute("""
            SELECT u.full_name, SUM(r.score) as total_score, COUNT(r.result_id) as tests, ROUND(AVG(r.percentage), 1) as avg_acc
            FROM results r
            LEFT JOIN users u ON r.user_id = u.user_id
            GROUP BY r.user_id
            ORDER BY total_score DESC, avg_acc DESC
            LIMIT 10
        """).fetchall()

    if not rows:
        text = "🏆 <b>HARTRON TOP RANKERS</b>\n\nAbhi tak kisi ne test complete nahi kiya hai. Pehla test dekar #1 Rank haasil karein!"
    else:
        medals = ["🥇", "🥈", "🥉"]
        lines = ["🏆 <b>HARTRON TOP 10 RANKERS</b>\n"]
        for idx, (fname, t_score, t_count, avg_acc) in enumerate(rows, start=1):
            badge = medals[idx - 1] if idx <= 3 else f"#{idx}"
            name_display = (fname or "Student")[:22]
            lines.append(f"{badge} <b>{name_display}</b> — {t_score} pts ({t_count} tests • {avg_acc}%)")
        lines.append("\n<i>Zyada tests attempt karein aur Top 3 mein jagah banayein!</i>")
        text = "\n".join(lines)

    await call.message.edit_text(
        text,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [button("🚀  START A TEST", "home:deo")],
            [button("⬅️  MAIN MENU", "home:main")],
        ])
    )
    last_bot_menu_msg[call.from_user.id] = call.message.message_id
    await call.answer()


@dp.callback_query(F.data == "home:daily")
async def home_daily_quiz(call: CallbackQuery):
    user = call.from_user
    register_user(user.id, user.username or "", user.full_name)
    today_str = datetime.now().strftime("%Y-%m-%d")

    with connect_db() as con:
        already = con.execute(
            "SELECT 1 FROM daily_attempts WHERE user_id=? AND attempt_date=?",
            (user.id, today_str)
        ).fetchone()

    if already:
        await call.answer("Aap aaj ka Daily Quiz pehle hi attempt kar chuke hain! Kal phir aayein.", show_alert=True)
        return

    if QUESTION_DF.empty:
        await call.answer("Questions available nahi hain.", show_alert=True)
        return

    seed_val = int(datetime.now().strftime("%Y%m%d"))
    row_idx = seed_val % len(QUESTION_DF)
    row = QUESTION_DF.iloc[row_idx]

    try:
        correct_index = get_correct_index(row["ans"])
    except Exception:
        await call.answer("Technical error in daily question.", show_alert=True)
        return

    options = [str(row[f"opt_{n}"]).strip()[:100] for n in range(1, 5)]
    question = str(row["question"]).strip()
    subject_tag = str(row.get("questype", "Computer")).strip()

    with connect_db() as con:
        con.execute(
            "INSERT OR IGNORE INTO daily_attempts(user_id, attempt_date) VALUES(?,?)",
            (user.id, today_str)
        )

    poll_msg = await call.message.answer_poll(
        question=f"📅 DAILY QUIZ ({subject_tag}): {question}"[:300],
        options=options,
        type="quiz",
        correct_option_id=correct_index,
        is_anonymous=False,
        open_period=30,
    )
    daily_poll_to_user[poll_msg.poll.id] = (user.id, correct_index, poll_msg.message_id)
    await call.answer("Aaj ka Daily Question bhej diya gaya hai!")


@dp.callback_query(F.data == "home:account")
async def home_account(call: CallbackQuery):
    user = call.from_user
    register_user(user.id, user.username or "", user.full_name)
    expiry = get_expiry(user.id)
    active = bool(expiry and expiry > datetime.now())
    with connect_db() as con:
        count, best = con.execute(
            "SELECT COUNT(*), COALESCE(MAX(percentage),0) FROM results WHERE user_id=?",
            (user.id,)
        ).fetchone()
    if active:
        status = f"🟢 ACTIVE — until <b>{expiry.strftime('%d %b %Y, %I:%M %p')}</b>"
    else:
        status = "🆓 FREE DEMO (Test Series 1 Unlocked)"
    await call.message.edit_text(
        "👤 <b>MY ACCOUNT</b>\n\n"
        f"Name: <b>{user.full_name}</b>\n"
        f"Telegram ID: <code>{user.id}</code>\n"
        f"Membership: {status}\n"
        f"Tests attempted: <b>{count}</b>\n"
        f"Best score: <b>{float(best):.1f}%</b>",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [button("💎  VIEW SUBSCRIPTION PLANS", "home:plans")],
            [button("⬅️  MAIN MENU", "home:main")],
        ])
    )
    last_bot_menu_msg[call.from_user.id] = call.message.message_id
    await call.answer()


@dp.callback_query(F.data == "home:contact")
async def home_contact(call: CallbackQuery):
    username = ADMIN_USERNAME.strip().lstrip("@")
    rows = []
    if username and username != "your_admin_username":
        rows.append([InlineKeyboardButton(
            text="💬  MESSAGE ADMIN",
            url=f"https://t.me/{username}"
        )])
    rows.append([button("⬅️  MAIN MENU", "home:main")])
    await call.message.edit_text(
        "📞 <b>HARTRON DEO SUPPORT</b>\n\n"
        "For subscription, payment verification or technical assistance, contact the administrator.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows)
    )
    last_bot_menu_msg[call.from_user.id] = call.message.message_id
    await call.answer()


@dp.callback_query(F.data == "noop")
async def noop(call: CallbackQuery):
    await call.answer("No test series are available in this subject yet.", show_alert=True)


# ------------------------------- SUBJECT / TEST SERIES -------------------------------

@dp.callback_query(F.data.startswith("topic:"))
async def select_topic(call: CallbackQuery):
    try:
        topic_index = int(call.data.split(":")[1])
        topic = TOPICS[topic_index]
    except (ValueError, IndexError):
        await call.answer("Invalid subject.", show_alert=True)
        return

    count = len(test_catalog.get(topic, {}))
    await call.message.edit_text(
        f"📂 <b>{topic.upper()}</b>\n\n"
        f"Available test series: <b>{count}</b>\n"
        "Select a test series below.\n\n"
        + ("🟢 Subscription active — all tests unlocked." if is_subscribed(call.from_user.id)
           else "🆓 Test Series 1 is FREE! Subscribe to unlock Test Series 2 onwards."),
        reply_markup=tests_keyboard(topic_index, call.from_user.id)
    )
    last_bot_menu_msg[call.from_user.id] = call.message.message_id
    await call.answer()


@dp.callback_query(F.data.startswith("test:"))
async def select_test(call: CallbackQuery):
    try:
        _, topic_index_text, test_no_text = call.data.split(":")
        topic_index, test_no = int(topic_index_text), int(test_no_text)
        topic = TOPICS[topic_index]
        questions = test_catalog[topic][test_no]
    except (ValueError, KeyError, IndexError):
        await call.answer("Test series not found.", show_alert=True)
        return

    if not can_access_test(call.from_user.id, test_no):
        await call.answer("🔒 Test Series 1 is free. Subscribe to unlock all remaining tests!", show_alert=True)
        await safe_delete(call.message.chat.id, call.message.message_id)
        await show_plans(call.message.chat.id, call.from_user.id)
        return

    score = get_best_score(call.from_user.id, topic, test_no)
    badge = " (FREE DEMO)" if (test_no == 1 and not is_subscribed(call.from_user.id)) else ""
    await call.message.edit_text(
        f"📝 <b>{topic.upper()} — TEST SERIES {test_no}{badge}</b>\n\n"
        f"Questions in this test: <b>{len(questions)}</b>\n"
        f"Best score: <b>{score:.1f}%</b>\n"
        "Mode: Computer Based Test (CBT)\n\n"
        "Choose your test mode:",
        reply_markup=test_detail_keyboard(topic_index, test_no)
    )
    last_bot_menu_msg[call.from_user.id] = call.message.message_id
    await call.answer()


@dp.callback_query(F.data.startswith("practice:"))
async def practice_mode(call: CallbackQuery):
    _, topic_index, test_no = call.data.split(":")
    if not can_access_test(call.from_user.id, int(test_no)):
        await call.answer("🔒 Please subscribe to unlock this test.", show_alert=True)
        await safe_delete(call.message.chat.id, call.message.message_id)
        await show_plans(call.message.chat.id, call.from_user.id)
        return
    await call.message.edit_text(
        "📘 <b>PRACTICE MODE SETTINGS</b>\n\nChoose time allowed per question:",
        reply_markup=timer_keyboard(int(topic_index), int(test_no), practice=True)
    )
    last_bot_menu_msg[call.from_user.id] = call.message.message_id
    await call.answer()


@dp.callback_query(F.data.startswith("timer:"))
async def exam_timer(call: CallbackQuery):
    _, topic_index, test_no = call.data.split(":")
    if not can_access_test(call.from_user.id, int(test_no)):
        await call.answer("🔒 Please subscribe to unlock this test.", show_alert=True)
        await safe_delete(call.message.chat.id, call.message.message_id)
        await show_plans(call.message.chat.id, call.from_user.id)
        return
    await call.message.edit_text(
        "⏱ <b>CBT EXAM SETTINGS</b>\n\nSelect time allowed for each question:",
        reply_markup=timer_keyboard(int(topic_index), int(test_no), practice=False)
    )
    last_bot_menu_msg[call.from_user.id] = call.message.message_id
    await call.answer()


# ------------------------------- QUIZ ENGINE -------------------------------

async def begin_test(chat_id: int, user_id: int, topic_index: int,
                     test_no: int, seconds: int, practice: bool):
    if not can_access_test(user_id, test_no):
        await show_plans(chat_id, user_id)
        return

    topic = TOPICS[topic_index]
    questions = test_catalog.get(topic, {}).get(test_no, [])
    if not questions:
        await bot.send_message(chat_id, "This test has no questions yet. Please contact admin.")
        return

    old = user_sessions.get(user_id)
    if old and old.get("timeout_task"):
        old["timeout_task"].cancel()
    if old and old.get("poll_id"):
        poll_to_user.pop(old["poll_id"], None)
    if old and old.get("last_poll_msg_id"):
        await safe_delete(chat_id, old["last_poll_msg_id"])

    chosen = list(questions)
    if SHUFFLE_QUESTIONS:
        random.shuffle(chosen)

    intro_msg = await bot.send_message(
        chat_id,
        f"🚀 <b>TEST STARTED</b>\n\n"
        f"Subject: <b>{topic}</b>\n"
        f"Test Series: <b>{test_no}</b>\n"
        f"Questions: <b>{len(chosen)}</b>\n"
        f"Time: <b>{seconds} seconds/question</b>\n"
        f"Mode: <b>{'Practice' if practice else 'CBT Exam'}</b>\n\n"
        "First question loading..."
    )

    user_sessions[user_id] = {
        "chat_id": chat_id,
        "topic": topic,
        "test_no": test_no,
        "questions": chosen,
        "index": 0,
        "timer": seconds,
        "score": 0,
        "wrong_list": [],
        "poll_id": None,
        "last_poll_msg_id": intro_msg.message_id,
        "correct_option": None,
        "timeout_task": None,
        "practice": practice,
    }
    await asyncio.sleep(1.5)
    await send_next_question(user_id)


@dp.callback_query(F.data.startswith("quizgo:"))
@dp.callback_query(F.data.startswith("practicego:"))
async def start_quiz_callback(call: CallbackQuery):
    try:
        mode, topic_index, test_no, seconds = call.data.split(":")
        if not can_access_test(call.from_user.id, int(test_no)):
            await call.answer("Subscription required.", show_alert=True)
            await safe_delete(call.message.chat.id, call.message.message_id)
            await show_plans(call.message.chat.id, call.from_user.id)
            return
        await call.answer("Your test is starting...")
        await safe_delete(call.message.chat.id, call.message.message_id)
        last_bot_menu_msg.pop(call.from_user.id, None)
        await begin_test(
            call.message.chat.id, call.from_user.id,
            int(topic_index), int(test_no), int(seconds),
            mode == "practicego"
        )
    except (ValueError, IndexError):
        await call.answer("Invalid test settings.", show_alert=True)


def get_correct_index(value):
    raw = str(value).strip().upper()
    if raw in {"A", "B", "C", "D"}:
        return ord(raw) - ord("A")
    match = re.search(r"[1-4]", raw)
    if match:
        number = int(match.group())
        return number - 1
    raise ValueError(f"Invalid answer value: {value}")


async def send_next_question(user_id: int):
    state = user_sessions.get(user_id)
    if not state:
        return

    # Delete previous question poll or intro message to keep chat clean
    if state.get("last_poll_msg_id"):
        await safe_delete(state["chat_id"], state["last_poll_msg_id"])
        state["last_poll_msg_id"] = None

    index = state["index"]
    questions = state["questions"]
    if index >= len(questions):
        await finish_test(user_id)
        return

    row = questions[index]
    try:
        correct_index = get_correct_index(row["ans"])
    except Exception:
        log.exception("Invalid answer key in row %s", index)
        state["index"] += 1
        await send_next_question(user_id)
        return

    options = [str(row[f"opt_{n}"]).strip()[:100] for n in range(1, 5)]
    question = str(row["question"]).strip()
    try:
        poll_message = await bot.send_poll(
            chat_id=state["chat_id"],
            question=f"Q{index + 1}/{len(questions)}: {question}"[:300],
            options=options,
            type="quiz",
            correct_option_id=correct_index,
            is_anonymous=False,
            open_period=state["timer"],
        )
        poll_id = poll_message.poll.id
        state["poll_id"] = poll_id
        state["last_poll_msg_id"] = poll_message.message_id
        state["correct_option"] = correct_index
        poll_to_user[poll_id] = user_id
        state["timeout_task"] = asyncio.create_task(
            question_timeout(user_id, poll_id, state["timer"])
        )
    except Exception:
        log.exception("Failed to send question poll")
        state["index"] += 1
        await send_next_question(user_id)


def record_wrong_question(state: dict, user_choice_text: str):
    row = state["questions"][state["index"]]
    correct_idx = state["correct_option"]
    correct_text = str(row[f"opt_{correct_idx + 1}"]).strip()
    state["wrong_list"].append({
        "q_no": state["index"] + 1,
        "question": str(row["question"]).strip(),
        "your_ans": user_choice_text,
        "correct_ans": correct_text,
    })


async def question_timeout(user_id: int, poll_id: str, seconds: int):
    await asyncio.sleep(seconds + 1)
    state = user_sessions.get(user_id)
    if state and state.get("poll_id") == poll_id:
        record_wrong_question(state, "Skipped (Time Out)")
        poll_to_user.pop(poll_id, None)
        state["poll_id"] = None
        state["index"] += 1
        await send_next_question(user_id)


@dp.poll_answer()
async def process_poll_answer(answer: PollAnswer):
    if answer.poll_id in daily_poll_to_user:
        uid, correct_idx, poll_msg_id = daily_poll_to_user.pop(answer.poll_id)
        await asyncio.sleep(1.2)
        await safe_delete(uid, poll_msg_id)
        selected = answer.option_ids[0] if answer.option_ids else None
        if selected == correct_idx:
            await bot.send_message(uid, "🎉 <b>Sahi Jawab!</b> Kal phir naye Daily Question ke liye aayein.")
        else:
            await bot.send_message(uid, "❌ <b>Galat Jawab!</b> Koi baat nahi, kal phir try karein ya Full Test Series practice karein.")
        return

    user_id = poll_to_user.get(answer.poll_id)
    state = user_sessions.get(user_id) if user_id else None
    if not state or state.get("poll_id") != answer.poll_id:
        return

    task = state.get("timeout_task")
    if task and not task.done():
        task.cancel()

    selected = answer.option_ids[0] if answer.option_ids else None
    if selected == state["correct_option"]:
        state["score"] += 1
    else:
        row = state["questions"][state["index"]]
        chosen_text = str(row[f"opt_{selected + 1}"]).strip() if selected is not None else "Skipped"
        record_wrong_question(state, chosen_text)

    poll_to_user.pop(answer.poll_id, None)
    state["poll_id"] = None
    state["index"] += 1
    await asyncio.sleep(1)
    await send_next_question(user_id)


async def finish_test(user_id: int):
    state = user_sessions.pop(user_id, None)
    if not state:
        return

    total = len(state["questions"])
    score = state["score"]
    wrong_or_skipped = total - score
    percentage = save_result(user_id, state["topic"], state["test_no"], score, total)
    last_wrong_answers[user_id] = state["wrong_list"]

    performance = (
        "EXCELLENT 🏆" if percentage >= 75
        else "GOOD 👍" if percentage >= 50
        else "KEEP PRACTICING 💪"
    )

    extra_note = ""
    buttons = [
        [button("🔄  RETAKE TEST", f"test:{TOPICS.index(state['topic'])}:{state['test_no']}")],
        [button("📚  ALL SUBJECTS", "home:deo")],
        [button("🏠  MAIN MENU", "home:main")],
    ]
    if state["wrong_list"]:
        buttons.insert(0, [button(f"❌  REVIEW WRONG ANSWERS ({len(state['wrong_list'])})", "review:wrong")])

    if not is_subscribed(user_id):
        extra_note = "\n\n🔓 <i>Want to unlock Test Series 2, 3, 4 & all Mix Tests? Upgrade to Premium!</i>"
        buttons.insert(1, [button("💎  UNLOCK ALL TESTS (₹59)", "home:plans")])

    res_msg = await bot.send_message(
        state["chat_id"],
        "🏁 <b>YOUR TEST RESULT</b>\n\n"
        f"Subject: <b>{state['topic']}</b>\n"
        f"Test Series: <b>{state['test_no']}</b>\n"
        f"Total Questions: <b>{total}</b>\n"
        f"Correct Answers: <b>{score}</b>\n"
        f"Wrong / Skipped: <b>{wrong_or_skipped}</b>\n"
        f"Accuracy: <b>{percentage:.1f}%</b>\n"
        f"Performance: <b>{performance}</b>\n\n"
        "Your result has been saved to My Account." + extra_note,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )
    last_bot_menu_msg[user_id] = res_msg.message_id


@dp.callback_query(F.data == "review:wrong")
async def review_wrong_answers(call: CallbackQuery):
    wrong_items = last_wrong_answers.get(call.from_user.id, [])
    if not wrong_items:
        await call.answer("Is test mein aapka koi bhi sawal galat nahi hua!", show_alert=True)
        return

    chunks = []
    current_msg = "❌ <b>WRONG / SKIPPED QUESTIONS REVIEW</b>\n\n"
    for item in wrong_items:
        block = (
            f"<b>Q{item['q_no']}.</b> {item['question']}\n"
            f"   🔸 Your Answer: <i>{item['your_ans']}</i>\n"
            f"   ✅ Correct Answer: <b>{item['correct_ans']}</b>\n\n"
        )
        if len(current_msg) + len(block) > 3800:
            chunks.append(current_msg)
            current_msg = block
        else:
            current_msg += block
    chunks.append(current_msg)

    for text_part in chunks:
        await call.message.answer(text_part)
    await call.answer()


@dp.message(Command("stop"))
async def stop_test(message: Message):
    await safe_delete(message.chat.id, message.message_id)
    state = user_sessions.pop(message.from_user.id, None)
    if state:
        task = state.get("timeout_task")
        if task and not task.done():
            task.cancel()
        if state.get("poll_id"):
            poll_to_user.pop(state["poll_id"], None)
        if state.get("last_poll_msg_id"):
            await safe_delete(state["chat_id"], state["last_poll_msg_id"])
        await send_clean_menu(
            message.chat.id,
            message.from_user.id,
            "⏹ Test stopped. Your partial score was not saved.",
            reply_markup=home_keyboard()
        )
    else:
        await send_clean_menu(
            message.chat.id,
            message.from_user.id,
            "You do not have an active test.",
            reply_markup=home_keyboard()
        )


# ------------------------------- SUBSCRIPTION / PAYMENT -------------------------------

@dp.callback_query(F.data.startswith("plan:"))
async def choose_plan(call: CallbackQuery):
    try:
        _, days_text, price_text = call.data.split(":")
        days, price = int(days_text), int(price_text)
    except ValueError:
        await call.answer("Invalid plan.", show_alert=True)
        return

    valid_plans = {
        (MONTHLY_PLAN_DAYS, MONTHLY_PLAN_PRICE),
        (TWO_MONTH_PLAN_DAYS, TWO_MONTH_PLAN_PRICE),
        (THREE_MONTH_PLAN_DAYS, THREE_MONTH_PLAN_PRICE),
    }
    if (days, price) not in valid_plans:
        await call.answer("Plan not available.", show_alert=True)
        return

    user = call.from_user
    register_user(user.id, user.username or "", user.full_name)
    with connect_db() as con:
        cursor = con.execute("""
            INSERT INTO payment_requests(user_id, username, full_name, plan_days, plan_price, status, created_at)
            VALUES(?,?,?,?,?,'pending',?)
        """, (
            user.id, user.username or "", user.full_name,
            days, price, datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))
        request_id = cursor.lastrowid

    await call.message.answer(
        "💳 <b>PAYMENT INSTRUCTIONS</b>\n\n"
        f"Selected plan: <b>{days} days</b>\n"
        f"Amount: <b>₹{price}</b>\n"
        f"Payment Request: <code>#{request_id}</code>\n\n"
        "1. Scan the QR code shown above or use the UPI details provided by admin.\n"
        f"2. Pay exactly ₹{price}.\n"
        "3. Send your payment screenshot here in this chat.\n"
        "4. Admin will verify the payment and activate your membership.\n\n"
        "Do not share your UPI PIN, OTP or banking password."
    )
    await call.answer("Plan selected.")


@dp.message(F.photo)
async def receive_payment_screenshot(message: Message):
    user = message.from_user
    with connect_db() as con:
        request = con.execute("""
            SELECT request_id, plan_days, plan_price
            FROM payment_requests
            WHERE user_id=? AND status='pending'
            ORDER BY request_id DESC LIMIT 1
        """, (user.id,)).fetchone()

    if not request:
        await message.answer(
            "No pending payment request was found. Please select a subscription plan first.",
            reply_markup=plans_keyboard()
        )
        return

    request_id, days, price = request
    caption = (
        "🔔 <b>NEW SUBSCRIPTION PAYMENT</b>\n\n"
        f"Request ID: <code>#{request_id}</code>\n"
        f"Name: {user.full_name}\n"
        f"User ID: <code>{user.id}</code>\n"
        f"Username: @{user.username or 'N/A'}\n"
        f"Plan: <b>{days} days</b>\n"
        f"Amount: <b>₹{price}</b>\n"
        "Status: PENDING"
    )
    await bot.send_photo(
        chat_id=ADMIN_ID,
        photo=message.photo[-1].file_id,
        caption=caption,
        reply_markup=payment_review_keyboard(request_id, user.id)
    )
    await message.answer(
        "✅ Payment screenshot submitted successfully.\n"
        "Your membership will be activated after admin verification."
    )


@dp.callback_query(F.data.startswith("approve:"))
async def approve_payment(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        await call.answer("Admin access only.", show_alert=True)
        return

    try:
        _, request_text, user_text = call.data.split(":")
        request_id, user_id = int(request_text), int(user_text)
    except ValueError:
        await call.answer("Invalid payment request.", show_alert=True)
        return

    with connect_db() as con:
        row = con.execute("""
            SELECT username, full_name, plan_days, plan_price, status
            FROM payment_requests WHERE request_id=? AND user_id=?
        """, (request_id, user_id)).fetchone()

    if not row:
        await call.answer("Payment request not found.", show_alert=True)
        return
    username, full_name, days, price, status = row
    if status != "pending":
        await call.answer("This request has already been processed.", show_alert=True)
        return

    expiry = activate_subscription(user_id, username or "", full_name or "", days)
    with connect_db() as con:
        con.execute(
            "UPDATE payment_requests SET status='approved' WHERE request_id=?",
            (request_id,)
        )

    try:
        await bot.send_message(
            user_id,
            "🎉 <b>PAYMENT APPROVED — PREMIUM ACTIVATED!</b>\n\n"
            f"Plan: <b>{days} days</b>\n"
            f"Paid: <b>₹{price}</b>\n"
            f"Valid until: <b>{expiry.strftime('%d %b %Y, %I:%M %p')}</b>\n\n"
            "All available DEO test series are now unlocked.\n"
            "Send /start to begin your preparation."
        )
    except Exception:
        log.warning("Could not DM user %s after approval", user_id)

    await call.message.edit_caption(
        caption=(call.message.caption or "") + "\n\n🟢 <b>APPROVED — ACCESS ACTIVATED</b>",
        reply_markup=None
    )
    await call.answer("Subscription activated successfully.")


@dp.callback_query(F.data.startswith("reject:"))
async def reject_payment(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        await call.answer("Admin access only.", show_alert=True)
        return

    try:
        _, request_text, user_text = call.data.split(":")
        request_id, user_id = int(request_text), int(user_text)
    except ValueError:
        await call.answer("Invalid payment request.", show_alert=True)
        return

    with connect_db() as con:
        con.execute(
            "UPDATE payment_requests SET status='rejected' WHERE request_id=? AND user_id=? AND status='pending'",
            (request_id, user_id)
        )

    await call.message.edit_caption(
        caption=(call.message.caption or "") + "\n\n🔴 <b>REJECTED</b>",
        reply_markup=None
    )
    try:
        await bot.send_message(
            user_id,
            "❌ Your payment request was not approved.\n"
            "Please check the payment details or contact admin for help."
        )
    except Exception:
        pass
    await call.answer("Payment request rejected.")


# ------------------------------- ADMIN COMMANDS -------------------------------

@dp.message(Command("reload"))
async def admin_reload(message: Message):
    if message.from_user.id != ADMIN_ID:
        return
    count = reload_question_bank()
    question_count = sum(len(v) for v in test_catalog.values())
    await message.answer(
        f"✅ Question bank reloaded.\n"
        f"Subjects with test sets: {count}\n"
        f"Total test series: {question_count}\n"
        f"Total question rows: {len(QUESTION_DF)}"
    )


@dp.message(Command("stats"))
async def admin_stats(message: Message):
    if message.from_user.id != ADMIN_ID:
        return
    with connect_db() as con:
        users = con.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        pending = con.execute(
            "SELECT COUNT(*) FROM payment_requests WHERE status='pending'"
        ).fetchone()[0]
        active = con.execute(
            "SELECT COUNT(*) FROM users WHERE expiry_date > ?",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),)
        ).fetchone()[0]
    await message.answer(
        "📊 <b>BOT ADMIN STATS</b>\n\n"
        f"Total registered users: {users}\n"
        f"Active paid subscriptions: {active}\n"
        f"Pending payments: {pending}\n"
        f"Question bank rows: {len(QUESTION_DF)}"
    )


@dp.message(Command("users"))
async def admin_list_users(message: Message):
    if message.from_user.id != ADMIN_ID:
        return
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with connect_db() as con:
        rows = con.execute(
            "SELECT user_id, username, full_name, expiry_date FROM users WHERE expiry_date > ? ORDER BY expiry_date DESC",
            (now_str,)
        ).fetchall()

    if not rows:
        await message.answer("ℹ️ Abhi kisi bhi student ka active subscription nahi hai.")
        return

    lines = ["👥 <b>ACTIVE PAID SUBSCRIBERS</b>\n"]
    for idx, (uid, uname, fname, exp) in enumerate(rows, start=1):
        u_tag = f"@{uname}" if uname else "No username"
        lines.append(
            f"{idx}. <b>{fname or 'Student'}</b> ({u_tag})\n"
            f"   ID: <code>{uid}</code> | Valid till: <b>{exp[:10]}</b>"
        )
    await message.answer("\n".join(lines))


@dp.message(Command("approve"))
async def admin_manual_approve(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    parts = (message.text or "").split()
    if len(parts) < 3 or not parts[1].isdigit() or not parts[2].isdigit():
        await message.answer(
            "⚠️ Sahi format use karein:\n"
            "<code>/approve USER_ID DAYS</code>\n\n"
            "Example (1 month ke liye):\n"
            "<code>/approve 123456789 30</code>"
        )
        return

    target_user_id = int(parts[1])
    days = int(parts[2])

    with connect_db() as con:
        req = con.execute(
            "SELECT username, full_name FROM payment_requests WHERE user_id=? ORDER BY request_id DESC LIMIT 1",
            (target_user_id,)
        ).fetchone()
        usr = con.execute(
            "SELECT username, full_name FROM users WHERE user_id=?",
            (target_user_id,)
        ).fetchone()

    username = (req[0] if req else None) or (usr[0] if usr else "") or ""
    full_name = (req[1] if req else None) or (usr[1] if usr else "Student") or "Student"

    expiry = activate_subscription(target_user_id, username, full_name, days)
    await message.answer(
        f"✅ <b>MANUAL APPROVAL SUCCESSFUL</b>\n\n"
        f"Name: <b>{full_name}</b>\n"
        f"User ID: <code>{target_user_id}</code>\n"
        f"Plan: <b>{days} days</b>\n"
        f"Valid until: <b>{expiry.strftime('%d %b %Y, %I:%M %p')}</b>"
    )
    try:
        await bot.send_message(
            target_user_id,
            "🎉 <b>PREMIUM SUBSCRIPTION ACTIVATED!</b>\n\n"
            f"Plan: <b>{days} days</b>\n"
            f"Valid until: <b>{expiry.strftime('%d %b %Y, %I:%M %p')}</b>\n\n"
            "All available DEO test series are now unlocked.\n"
            "Send /start to begin your preparation."
        )
    except Exception:
        log.warning("Could not DM user %s after manual approval", target_user_id)


@dp.message(Command("cancel"))
async def admin_cancel_subscription(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    parts = (message.text or "").split()
    if len(parts) < 2 or not parts[1].isdigit():
        await message.answer(
            "⚠️ Sahi format use karein:\n"
            "<code>/cancel USER_ID</code>\n\n"
            "Example: <code>/cancel 123456789</code>"
        )
        return

    target_user_id = int(parts[1])
    with connect_db() as con:
        row = con.execute(
            "SELECT full_name FROM users WHERE user_id=?",
            (target_user_id,)
        ).fetchone()
        con.execute("UPDATE users SET expiry_date=NULL WHERE user_id=?", (target_user_id,))

    if row:
        await message.answer(
            f"🚫 <b>SUBSCRIPTION CANCELLED</b>\n\n"
            f"Name: {row[0] or 'N/A'}\n"
            f"User ID: <code>{target_user_id}</code>\n"
            "Inka premium access ab band kar diya gaya hai."
        )
        try:
            await bot.send_message(
                target_user_id,
                "⚠️ Aapka premium subscription cancel kar diya gaya hai. "
                "Kisi bhi jaankari ke liye admin se sampark karein."
            )
        except Exception:
            pass
    else:
        await message.answer(f"ℹ️ User ID <code>{target_user_id}</code> ka koi active subscription nahi mila.")


@dp.message(Command("broadcast"))
async def admin_broadcast(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    text_to_send = (message.text or "").partition(" ")[2].strip()
    if not text_to_send:
        await message.answer(
            "⚠️ Sahi format use karein:\n"
            "<code>/broadcast Aapka message yahan likhein</code>"
        )
        return

    with connect_db() as con:
        user_ids = [r[0] for r in con.execute("SELECT user_id FROM users").fetchall()]

    sent = 0
    failed = 0
    status_msg = await message.answer(f"📢 {len(user_ids)} students ko message bheja ja raha hai...")

    for uid in user_ids:
        try:
            await bot.send_message(uid, f"📢 <b>HARTRON RANKERS UPDATE</b>\n\n{text_to_send}")
            sent += 1
            await asyncio.sleep(0.05)
        except Exception:
            failed += 1

    await status_msg.edit_text(
        f"✅ <b>BROADCAST COMPLETE</b>\n\n"
        f"Delivered: <b>{sent}</b>\n"
        f"Failed (Blocked bot): <b>{failed}</b>"
    )


@dp.message(Command("remind"))
async def admin_expiry_reminders(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    now = datetime.now()
    three_days_later = (now + timedelta(days=3)).strftime("%Y-%m-%d %H:%M:%S")
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")

    with connect_db() as con:
        rows = con.execute(
            "SELECT user_id, full_name, expiry_date FROM users WHERE expiry_date > ? AND expiry_date <= ?",
            (now_str, three_days_later)
        ).fetchall()

    if not rows:
        await message.answer("ℹ️ Agle 3 din mein kisi bhi student ka subscription expire nahi ho raha hai.")
        return

    sent = 0
    for uid, fname, exp in rows:
        try:
            await bot.send_message(
                uid,
                f"⏰ <b>SUBSCRIPTION EXPIRY REMINDER</b>\n\n"
                f"Namaste <b>{fname or 'Student'}</b>,\n"
                f"Aapka Hartron Rankers Premium Plan <b>{exp[:10]}</b> ko expire hone wala hai.\n\n"
                "Bina kisi rukawat ke apni practice jaari rakhne ke liye abhi renew karein!",
                reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                    [button("💎  RENEW NOW", "home:plans")]
                ])
            )
            sent += 1
        except Exception:
            pass

    await message.answer(f"✅ {sent} students ko renewal reminder bhej diya gaya hai.")


@dp.message(Command("backup"))
async def admin_backup_csv(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    with connect_db() as con:
        rows = con.execute(
            "SELECT user_id, username, full_name, expiry_date FROM users ORDER BY expiry_date DESC"
        ).fetchall()

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["user_id", "username", "full_name", "expiry_date"])
    writer.writerows(rows)

    csv_bytes = output.getvalue().encode("utf-8")
    filename = f"hartron_users_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
    await message.answer_document(
        BufferedInputFile(csv_bytes, filename=filename),
        caption=f"💾 <b>DATABASE BACKUP</b>\nTotal Users: {len(rows)}\n\nIs file ko safe rakhein!"
    )


# ------------------------------- WEBHOOK & FASTAPI -------------------------------

@asynccontextmanager
async def lifespan(app: FastAPI):
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN environment variable is missing.")
    init_db()
    total_tests = sum(len(items) for items in test_catalog.values())
    log.info("Hartron DEO bot starting | question rows=%s | test series=%s", len(QUESTION_DF), total_tests)

    render_url = os.environ.get("RENDER_EXTERNAL_URL", "").rstrip("/")
    if render_url:
        await bot.set_webhook(
            url=f"{render_url}/telegram/webhook",
            secret_token=TELEGRAM_WEBHOOK_SECRET,
            drop_pending_updates=True
        )
        log.info("Webhook set to %s/telegram/webhook", render_url)
    yield
    await bot.session.close()


app = FastAPI(lifespan=lifespan)


@app.get("/")
async def health_check():
    return {"status": "Hartron DEO Bot is live!"}


@app.post("/telegram/webhook")
async def telegram_webhook(request: Request):
    token = request.headers.get("X-Telegram-Bot-Api-Secret-Token", "")
    if not hmac.compare_digest(token, TELEGRAM_WEBHOOK_SECRET):
        raise HTTPException(status_code=403, detail="Invalid secret token")

    update = Update.model_validate(await request.json())
    await dp.feed_update(bot, update)
    return {"ok": True}