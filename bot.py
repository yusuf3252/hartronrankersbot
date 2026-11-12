import asyncio
import hmac
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
    InlineKeyboardButton, InlineKeyboardMarkup, FSInputFile,
    CallbackQuery, Message, PollAnswer, Update
)
from fastapi import FastAPI, Request, HTTPException

# ============================ SETTINGS ============================

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
ADMIN_ID = int(os.environ.get("ADMIN_ID", "0"))
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "theStudentSupportZone")
TELEGRAM_WEBHOOK_SECRET = os.environ.get("TELEGRAM_WEBHOOK_SECRET", "hartron_deo_secret_99")

MONTHLY_PLAN_PRICE = 149
MONTHLY_PLAN_DAYS = 30
TWO_MONTH_PLAN_PRICE = 239
TWO_MONTH_PLAN_DAYS = 60

QUESTIONS_PER_TEST = 30
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
TOPICS = list(TOPIC_ALIASES.keys())

user_sessions = {}
poll_to_user = {}
test_catalog = {}


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
        con.execute("CREATE INDEX IF NOT EXISTS idx_results_user ON results(user_id)")

        user_columns = {row[1] for row in con.execute("PRAGMA table_info(users)").fetchall()}
        if "full_name" not in user_columns:
            con.execute("ALTER TABLE users ADD COLUMN full_name TEXT")


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
        [button("💎  SUBSCRIPTION PLANS", "home:plans")],
        [button("👤  MY ACCOUNT", "home:account")],
        [button("📞  CONTACT ADMIN", "home:contact")],
    ])


def topics_keyboard():
    rows = []
    for index, topic in enumerate(TOPICS):
        count = len(test_catalog.get(topic, {}))
        rows.append([button(f"📚  {topic}  •  {count} TESTS", f"topic:{index}")])
    rows.append([button("⬅️  MAIN MENU", "home:main")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def tests_keyboard(topic_index: int, user_id: int):
    topic = TOPICS[topic_index]
    tests = test_catalog.get(topic, {})
    rows = []
    for test_no in sorted(tests):
        score = get_best_score(user_id, topic, test_no)
        if is_subscribed(user_id):
            label = f"📝  TEST SERIES {test_no}  •  BEST {score:.0f}%"
        else:
            label = f"🔒  TEST SERIES {test_no}  •  PREMIUM"
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
        [button(f"🥈  2 MONTHS — ₹{TWO_MONTH_PLAN_PRICE} (SAVE ₹59)", f"plan:{TWO_MONTH_PLAN_DAYS}:{TWO_MONTH_PLAN_PRICE}")],
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
        "✅ Unlock every DEO subject and test series\n"
        "✅ Access 1,000+ questions (as your question bank grows)\n"
        "✅ Unlimited test attempts during membership\n"
        "✅ Topic-wise practice and CBT exam mode\n"
        "✅ Choose 15, 20 or 30 seconds per question\n"
        "✅ Instant score and accuracy report\n"
        "✅ Personal best-score tracking\n"
        "✅ Retake tests to improve performance\n"
        "✅ New tests and questions can be added by admin\n"
        "✅ 30-day or 60-day access options\n\n"
        "<i>Question count depends on the questions uploaded by the admin.</i>"
    )


async def show_plans(message: Message):
    text = (
        "💎 <b>CHOOSE YOUR PREMIUM PLAN</b>\n\n"
        f"🥉 <b>1 MONTH — ₹{MONTHLY_PLAN_PRICE}</b>\n"
        "Access validity: 30 days\n\n"
        f"🥈 <b>2 MONTHS — ₹{TWO_MONTH_PLAN_PRICE}</b>\n"
        "Access validity: 60 days\n"
        "Save ₹59 compared with two monthly renewals.\n\n"
        + premium_benefits()
        + "\n\nSelect a plan below to continue with payment."
    )
    if os.path.isfile(QR_PATH):
        await message.answer_photo(FSInputFile(QR_PATH), caption=text, reply_markup=plans_keyboard())
    else:
        await message.answer(text, reply_markup=plans_keyboard())


# ------------------------------- START & HOME MENU -------------------------------

@dp.message(CommandStart())
async def start_handler(message: Message):
    await message.answer(
        "🎓 <b>WELCOME TO HARTRON DEO PREMIUM CBT PORTAL</b>\n\n"
        "Your focused preparation platform for the Data Entry Operator exam.\n\n"
        "📚 Subject-wise test series\n"
        "🧠 1,000+ question learning bank as content is uploaded\n"
        "⏱ Timed CBT examination\n"
        "📊 Performance and score tracking\n"
        "🔐 Premium access for all test series\n\n"
        "<b>Select an option to get started:</b>",
        reply_markup=home_keyboard()
    )


@dp.callback_query(F.data == "home:main")
async def home_main(call: CallbackQuery):
    await call.message.answer(
        "🎓 <b>HARTRON DEO PREMIUM CBT PORTAL</b>\n\nSelect an option:",
        reply_markup=home_keyboard()
    )
    await call.answer()


@dp.callback_query(F.data == "home:deo")
async def home_deo(call: CallbackQuery):
    await call.message.edit_text(
        "📚 <b>DATA ENTRY OPERATOR — SUBJECTS</b>\n\n"
        "Choose a subject to view its available test series.\n"
        "🔒 All test series require an active subscription.",
        reply_markup=topics_keyboard()
    )
    await call.answer()


@dp.callback_query(F.data == "home:plans")
async def home_plans(call: CallbackQuery):
    await show_plans(call.message)
    await call.answer()


@dp.callback_query(F.data == "home:account")
async def home_account(call: CallbackQuery):
    user = call.from_user
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
        status = "🔴 INACTIVE"
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
        + ("🟢 Subscription active — tests unlocked." if is_subscribed(call.from_user.id)
           else "🔒 Premium membership is required to start any test."),
        reply_markup=tests_keyboard(topic_index, call.from_user.id)
    )
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

    if not is_subscribed(call.from_user.id):
        await call.answer("🔒 Premium subscription required to unlock all tests.", show_alert=True)
        await show_plans(call.message)
        return

    score = get_best_score(call.from_user.id, topic, test_no)
    await call.message.edit_text(
        f"📝 <b>{topic.upper()} — TEST SERIES {test_no}</b>\n\n"
        f"Questions in this test: <b>{len(questions)}</b>\n"
        f"Best score: <b>{score:.1f}%</b>\n"
        "Mode: Computer Based Test (CBT)\n\n"
        "Choose your test mode:",
        reply_markup=test_detail_keyboard(topic_index, test_no)
    )
    await call.answer()


@dp.callback_query(F.data.startswith("practice:"))
async def practice_mode(call: CallbackQuery):
    if not is_subscribed(call.from_user.id):
        await call.answer("🔒 Please subscribe to unlock practice tests.", show_alert=True)
        await show_plans(call.message)
        return
    _, topic_index, test_no = call.data.split(":")
    await call.message.edit_text(
        "📘 <b>PRACTICE MODE SETTINGS</b>\n\nChoose time allowed per question:",
        reply_markup=timer_keyboard(int(topic_index), int(test_no), practice=True)
    )
    await call.answer()


@dp.callback_query(F.data.startswith("timer:"))
async def exam_timer(call: CallbackQuery):
    if not is_subscribed(call.from_user.id):
        await call.answer("🔒 Please subscribe to unlock CBT exams.", show_alert=True)
        await show_plans(call.message)
        return
    _, topic_index, test_no = call.data.split(":")
    await call.message.edit_text(
        "⏱ <b>CBT EXAM SETTINGS</b>\n\nSelect time allowed for each question:",
        reply_markup=timer_keyboard(int(topic_index), int(test_no), practice=False)
    )
    await call.answer()


# ------------------------------- QUIZ ENGINE -------------------------------

async def begin_test(message: Message, user_id: int, topic_index: int,
                     test_no: int, seconds: int, practice: bool):
    if not is_subscribed(user_id):
        await message.answer("🔒 An active premium subscription is required.")
        await show_plans(message)
        return

    topic = TOPICS[topic_index]
    questions = test_catalog.get(topic, {}).get(test_no, [])
    if not questions:
        await message.answer("This test has no questions yet. Please contact admin.")
        return

    old = user_sessions.get(user_id)
    if old and old.get("timeout_task"):
        old["timeout_task"].cancel()
    if old and old.get("poll_id"):
        poll_to_user.pop(old["poll_id"], None)

    chosen = list(questions)
    if SHUFFLE_QUESTIONS:
        random.shuffle(chosen)

    user_sessions[user_id] = {
        "chat_id": message.chat.id,
        "topic": topic,
        "test_no": test_no,
        "questions": chosen,
        "index": 0,
        "timer": seconds,
        "score": 0,
        "poll_id": None,
        "correct_option": None,
        "timeout_task": None,
        "practice": practice,
    }
    await message.answer(
        f"🚀 <b>TEST STARTED</b>\n\n"
        f"Subject: <b>{topic}</b>\n"
        f"Test Series: <b>{test_no}</b>\n"
        f"Questions: <b>{len(chosen)}</b>\n"
        f"Time: <b>{seconds} seconds/question</b>\n"
        f"Mode: <b>{'Practice' if practice else 'CBT Exam'}</b>\n\n"
        "Answer each question before the timer ends. Good luck!"
    )
    await send_next_question(user_id)


@dp.callback_query(F.data.startswith("quizgo:"))
@dp.callback_query(F.data.startswith("practicego:"))
async def start_quiz_callback(call: CallbackQuery):
    if not is_subscribed(call.from_user.id):
        await call.answer("Subscription required.", show_alert=True)
        await show_plans(call.message)
        return
    try:
        mode, topic_index, test_no, seconds = call.data.split(":")
        await call.answer("Your test is starting...")
        try:
            await call.message.delete()
        except Exception:
            pass
        await begin_test(
            call.message, call.from_user.id,
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
        state["correct_option"] = correct_index
        poll_to_user[poll_id] = user_id
        state["timeout_task"] = asyncio.create_task(
            question_timeout(user_id, poll_id, state["timer"])
        )
    except Exception:
        log.exception("Failed to send question poll")
        state["index"] += 1
        await send_next_question(user_id)


async def question_timeout(user_id: int, poll_id: str, seconds: int):
    await asyncio.sleep(seconds + 1)
    state = user_sessions.get(user_id)
    if state and state.get("poll_id") == poll_id:
        poll_to_user.pop(poll_id, None)
        state["poll_id"] = None
        state["index"] += 1
        await send_next_question(user_id)


@dp.poll_answer()
async def process_poll_answer(answer: PollAnswer):
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
    performance = (
        "EXCELLENT 🏆" if percentage >= 75
        else "GOOD 👍" if percentage >= 50
        else "KEEP PRACTICING 💪"
    )

    await bot.send_message(
        state["chat_id"],
        "🏁 <b>YOUR TEST RESULT</b>\n\n"
        f"Subject: <b>{state['topic']}</b>\n"
        f"Test Series: <b>{state['test_no']}</b>\n"
        f"Total Questions: <b>{total}</b>\n"
        f"Correct Answers: <b>{score}</b>\n"
        f"Wrong / Skipped: <b>{wrong_or_skipped}</b>\n"
        f"Accuracy: <b>{percentage:.1f}%</b>\n"
        f"Performance: <b>{performance}</b>\n\n"
        "Your result has been saved to My Account.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [button("🔄  RETAKE TEST", f"test:{TOPICS.index(state['topic'])}:{state['test_no']}")],
            [button("📚  ALL SUBJECTS", "home:deo")],
            [button("🏠  MAIN MENU", "home:main")],
        ])
    )


@dp.message(Command("stop"))
async def stop_test(message: Message):
    state = user_sessions.pop(message.from_user.id, None)
    if state:
        task = state.get("timeout_task")
        if task and not task.done():
            task.cancel()
        if state.get("poll_id"):
            poll_to_user.pop(state["poll_id"], None)
        await message.answer("⏹ Test stopped. Your partial score was not saved.")
    else:
        await message.answer("You do not have an active test.")


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
    }
    if (days, price) not in valid_plans:
        await call.answer("Plan not available.", show_alert=True)
        return

    user = call.from_user
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
        con.execute("DELETE FROM users WHERE user_id=?", (target_user_id,))

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
        f"Registered members: {users}\n"
        f"Active subscriptions: {active}\n"
        f"Pending payments: {pending}\n"
        f"Question bank rows: {len(QUESTION_DF)}"
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