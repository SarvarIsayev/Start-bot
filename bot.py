import logging
import os
import json
from datetime import datetime, time
import pytz
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)
import gspread
from google.oauth2.service_account import Credentials

# ─── SOZLAMALAR ───────────────────────────────────────────────────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN", "BU_YERGA_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID", "123456789"))
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1L4wpKTkFghanh55c2_tNVD7O3CvmphhDTeJ7cKcKke8")
CREDENTIALS_FILE = "credentials.json"
UZ_TZ = pytz.timezone("Asia/Tashkent")

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ─── STATES ───────────────────────────────────────────────────────────────────
(REG_ISM, REG_BIZNES_TURI, REG_A, REG_B,
 REG_M1, REG_M2, REG_M3, REG_M4, REG_M5, REG_M6,
 DAILY_SANA, DAILY_METRIKA, DAILY_REJA, DAILY_PLAN,
 EVE_FAKT, EVE_DAROMAD) = range(16)

CHAKANA_METRIKALARI = [
    "Mehmonlar soni",
    "Xaridorlar soni",
    "O'rtacha chek",
    "Marja",
    "Qayta sotuv",
]
SERVICE_METRIKALARI = [
    "Ko'rishlar soni",
    "Qo'ng'iroqlar soni",
    "Uchrashuvlar soni",
    "Xaridlar soni",
    "O'rtacha chek",
    "Marja",
]
DISTRIBUTSIYA_METRIKALARI = [
    "OKB",
    "AKB",
    "Xaridlar soni",
    "O'rtacha chek",
    "Marja",
]

def get_metrikalari(biznes_turi: str):
    if biznes_turi == "Chakana":
        return CHAKANA_METRIKALARI
    elif biznes_turi == "Service":
        return SERVICE_METRIKALARI
    else:
        return DISTRIBUTSIYA_METRIKALARI

# ─── GOOGLE SHEETS ────────────────────────────────────────────────────────────
def get_sheet():
    creds_json = os.getenv("GOOGLE_CREDENTIALS")
    if creds_json:
        creds_dict = json.loads(creds_json)
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"]
        )
    else:
        creds = Credentials.from_service_account_file(
            CREDENTIALS_FILE,
            scopes=["https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"]
        )
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

def get_or_create_tab(spreadsheet, tab_name: str, metrikalari: list):
    try:
        ws = spreadsheet.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=tab_name, rows=200, cols=15)
        # Sarlavhalar
        ws.update("A1:H1", [["Sana", "Metrika", "B nuqta uchun reja", "Plan", "Fakt", "Daromad/Sof foyda", "", ""]])
        ws.format("A1:H1", {"textFormat": {"bold": True}})
    return ws

def save_registration(tab_name: str, profile: dict):
    spreadsheet = get_sheet()
    # "Ro'yxat" nomli umumiy sheet
    try:
        ws = spreadsheet.worksheet("Ro'yxat")
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="Ro'yxat", rows=200, cols=20)
        headers = ["Ism Familiya", "Biznes turi", "A nuqta", "B nuqta"]
        metrikalari_all = CHAKANA_METRIKALARI + [m for m in SERVICE_METRIKALARI if m not in CHAKANA_METRIKALARI] + [m for m in DISTRIBUTSIYA_METRIKALARI if m not in SERVICE_METRIKALARI]
        for m in metrikalari_all:
            headers += [f"{m} (A)", f"{m} (B)"]
        ws.update("A1", [headers])
        ws.format("A1:Z1", {"textFormat": {"bold": True}})

    metrikalari = get_metrikalari(profile["biznes_turi"])
    row = [
        profile["ism"],
        profile["biznes_turi"],
        profile["a_nuqta"],
        profile["b_nuqta"],
    ]
    for i, m in enumerate(metrikalari):
        key_a = f"m{i+1}_a"
        key_b = f"m{i+1}_b"
        row += [profile.get(key_a, ""), profile.get(key_b, "")]

    all_vals = ws.get_all_values()
    ws.update(f"A{len(all_vals)+1}", [row])

    # Alohida tab yaratish
    get_or_create_tab(spreadsheet, tab_name, metrikalari)

def save_morning(tab_name: str, data: dict):
    spreadsheet = get_sheet()
    ws = spreadsheet.worksheet(tab_name)
    all_vals = ws.get_all_values()
    # Oxirgi qatorni topamiz — agar bugungi sana bor bo'lsa o'sha qatorga, aks holda yangi
    today = data["sana"]
    row_idx = None
    for i, row in enumerate(all_vals):
        if row and row[0] == today:
            row_idx = i + 1
            break
    if row_idx is None:
        row_idx = len(all_vals) + 1
    ws.update(f"A{row_idx}:D{row_idx}", [[data["sana"], data["metrika"], data["reja"], data["plan"]]])

def save_evening(tab_name: str, data: dict):
    spreadsheet = get_sheet()
    ws = spreadsheet.worksheet(tab_name)
    all_vals = ws.get_all_values()
    today = data["sana"]
    row_idx = None
    for i, row in enumerate(all_vals):
        if row and row[0] == today:
            row_idx = i + 1
            break
    if row_idx is None:
        row_idx = len(all_vals) + 1
    ws.update(f"E{row_idx}:F{row_idx}", [[data["fakt"], data["daromad"]]])

# ─── /start ───────────────────────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = context.user_data.get("profile", {})
    if profile.get("ism"):
        await update.message.reply_text(
            f"👋 Salom, {profile['ism']}!\n\n"
            "/kunlik — kunlik hisobot\n"
            "/kechki — kechki hisobot\n"
            "/profil — profilingiz"
        )
        return ConversationHandler.END
    await update.message.reply_text(
        "👋 Assalomu alaykum!\n\n"
        "30 kunlik biznes o'sish botiga xush kelibsiz. 📈\n\n"
        "Ro'yxatdan o'tamiz.\n\n"
        "1️⃣ Ism Familiyangizni kiriting:"
    )
    return REG_ISM

async def reg_ism(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("profile", {})["ism"] = update.message.text.strip()
    kb = [["Chakana", "Service", "Distributsiya"]]
    await update.message.reply_text(
        "2️⃣ Biznesingiz turi:",
        reply_markup=ReplyKeyboardMarkup(kb, one_time_keyboard=True, resize_keyboard=True)
    )
    return REG_BIZNES_TURI

async def reg_biznes_turi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    turi = update.message.text.strip()
    if turi not in ["Chakana", "Service", "Distributsiya"]:
        await update.message.reply_text("Iltimos, quyidagilardan birini tanlang:", 
            reply_markup=ReplyKeyboardMarkup([["Chakana", "Service", "Distributsiya"]], one_time_keyboard=True, resize_keyboard=True))
        return REG_BIZNES_TURI
    context.user_data["profile"]["biznes_turi"] = turi
    await update.message.reply_text(
        "3️⃣ A nuqtangiz qancha? (hozirgi oylik daromadingiz):",
        reply_markup=ReplyKeyboardRemove()
    )
    return REG_A

async def reg_a(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["profile"]["a_nuqta"] = update.message.text.strip()
    await update.message.reply_text("4️⃣ B nuqtangiz qancha? (30 kundan keyin bo'lishi kerak bo'lgan daromad):")
    return REG_B

async def reg_b(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["profile"]["b_nuqta"] = update.message.text.strip()
    biznes_turi = context.user_data["profile"]["biznes_turi"]
    metrikalari = get_metrikalari(biznes_turi)
    context.user_data["metrikalari"] = metrikalari
    context.user_data["metrika_step"] = 0

    m = metrikalari[0]
    await update.message.reply_text(
        f"Endi {biznes_turi} biznesining hozirgi holatini aniqlaymiz.\n\n"
        f"5️⃣ {m} — A nuqtada qancha?"
    )
    return REG_M1

async def reg_metrika(update: Update, context: ContextTypes.DEFAULT_TYPE, step: int, next_state):
    metrikalari = context.user_data["metrikalari"]
    idx = (step - 1) // 2
    is_a = (step - 1) % 2 == 0

    key = f"m{idx+1}_{'a' if is_a else 'b'}"
    context.user_data["profile"][key] = update.message.text.strip()

    # Keyingi savol
    if is_a:
        m = metrikalari[idx]
        await update.message.reply_text(f"{m} — B nuqtada qancha?")
    else:
        next_idx = idx + 1
        if next_idx < len(metrikalari):
            m = metrikalari[next_idx]
            await update.message.reply_text(f"{m} — A nuqtada qancha?")
        else:
            # Tugadi — saqlash
            profile = context.user_data["profile"]
            tab_name = profile["ism"][:25]
            context.user_data["profile"]["tab_name"] = tab_name
            try:
                save_registration(tab_name, profile)
                msg = "✅ Google Sheets ga saqlandi!"
            except Exception as e:
                logger.error(e)
                msg = "⚠️ Sheets ga saqlashda xato. Admin bilan bog'laning."

            await update.message.reply_text(
                f"🎉 Ro'yxatdan o'tdingiz!\n\n"
                f"👤 {profile['ism']}\n"
                f"🏪 {profile['biznes_turi']} biznes\n"
                f"📊 A nuqta: {profile['a_nuqta']}\n"
                f"🎯 B nuqta: {profile['b_nuqta']}\n\n"
                f"{msg}\n\n"
                f"Har kuni ertalab soat 9:00 da savol keladi! 📅"
            )
            return ConversationHandler.END
    return next_state

# Har bir metrika uchun handler
async def reg_m1(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await reg_metrika(update, context, 1, REG_M2)
async def reg_m2(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await reg_metrika(update, context, 2, REG_M3)
async def reg_m3(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await reg_metrika(update, context, 3, REG_M4)
async def reg_m4(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await reg_metrika(update, context, 4, REG_M5)
async def reg_m5(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await reg_metrika(update, context, 5, REG_M6)
async def reg_m6(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await reg_metrika(update, context, 6, ConversationHandler.END)

# ─── ERTALABKI SAVOL (soat 9:00) ─────────────────────────────────────────────
async def kunlik_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = context.user_data.get("profile", {})
    if not profile.get("ism"):
        await update.message.reply_text("❌ Avval /start orqali ro'yxatdan o'ting.")
        return ConversationHandler.END
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    await update.message.reply_text(
        f"🌅 Ertalabki hisobot\n\n"
        f"📅 Bugungi sana: {today}\n\n"
        f"Qaysi metrikaga bugun ta'sir qilmoqchisiz?\n"
        f"(Masalan: Mehmonlar soni, O'rtacha chek...)"
    )
    context.user_data["daily_sana"] = today
    return DAILY_METRIKA

async def daily_metrika(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["daily_metrika"] = update.message.text.strip()
    await update.message.reply_text("B nuqtaga yetish uchun bugun nima qilasiz?")
    return DAILY_REJA

async def daily_reja(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["daily_reja"] = update.message.text.strip()
    await update.message.reply_text("🎯 Bugungi PLAN (raqamda):")
    return DAILY_PLAN

async def daily_plan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = context.user_data.get("profile", {})
    data = {
        "sana": context.user_data.get("daily_sana", datetime.now(UZ_TZ).strftime("%d.%m.%Y")),
        "metrika": context.user_data.get("daily_metrika", ""),
        "reja": context.user_data.get("daily_reja", ""),
        "plan": update.message.text.strip(),
    }
    try:
        save_morning(profile["tab_name"], data)
        msg = "✅ Saqlandi!"
    except Exception as e:
        logger.error(e)
        msg = "⚠️ Xato yuz berdi."

    await update.message.reply_text(
        f"✅ Ertalabki hisobot qabul qilindi!\n\n"
        f"Kechqurun soat 9:00 da fakt va daromadni kiritishni unutmang. 💪\n\n{msg}"
    )
    return ConversationHandler.END

# ─── KECHKI SAVOL (soat 21:00) ───────────────────────────────────────────────
async def kechki_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = context.user_data.get("profile", {})
    if not profile.get("ism"):
        await update.message.reply_text("❌ Avval /start orqali ro'yxatdan o'ting.")
        return ConversationHandler.END
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    context.user_data["daily_sana"] = today
    await update.message.reply_text(
        f"🌆 Kechki hisobot — {today}\n\n"
        f"✅ Bugungi FAKT (haqiqatda nima bo'ldi, raqamda):"
    )
    return EVE_FAKT

async def eve_fakt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["eve_fakt"] = update.message.text.strip()
    await update.message.reply_text("💵 Bugungi kunlik aylanma yoki sof foyda (so'mda):")
    return EVE_DAROMAD

async def eve_daromad(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = context.user_data.get("profile", {})
    data = {
        "sana": context.user_data.get("daily_sana", datetime.now(UZ_TZ).strftime("%d.%m.%Y")),
        "fakt": context.user_data.get("eve_fakt", ""),
        "daromad": update.message.text.strip(),
    }
    try:
        save_evening(profile["tab_name"], data)
        msg = "✅ Saqlandi!"
    except Exception as e:
        logger.error(e)
        msg = "⚠️ Xato yuz berdi."

    await update.message.reply_text(
        f"🎉 Bugungi hisobot to'liq saqlandi!\n\n"
        f"Ertaga ham /kunlik buyrug'ini kutamiz. 💪\n\n{msg}"
    )
    return ConversationHandler.END

# ─── AVTOMATIK XABAR YUBORISH ─────────────────────────────────────────────────
async def send_morning_reminder(context: ContextTypes.DEFAULT_TYPE):
    """Ertalab 9:00 da barcha userlarga yuboriladi"""
    job_data = context.job.data
    user_id = job_data["user_id"]
    ism = job_data["ism"]
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    weekday = datetime.now(UZ_TZ).weekday()
    if weekday == 6:  # Yakshanba
        return
    await context.bot.send_message(
        chat_id=user_id,
        text=(
            f"🌅 Xayrli tong, {ism}!\n\n"
            f"📅 Bugun: {today}\n\n"
            f"Bugungi ertalabki hisobotni to'ldiring 👇\n"
            f"/kunlik"
        )
    )

async def send_evening_reminder(context: ContextTypes.DEFAULT_TYPE):
    """Kechqurun 21:00 da barcha userlarga yuboriladi"""
    job_data = context.job.data
    user_id = job_data["user_id"]
    ism = job_data["ism"]
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    weekday = datetime.now(UZ_TZ).weekday()
    if weekday == 6:  # Yakshanba
        return
    await context.bot.send_message(
        chat_id=user_id,
        text=(
            f"🌆 Kechqurun eslatma, {ism}!\n\n"
            f"📅 Bugun: {today}\n\n"
            f"Kechki hisobotni to'ldiring 👇\n"
            f"/kechki"
        )
    )

def schedule_reminders(app, user_id: int, ism: str):
    """Userga kunlik eslatmalar qo'shish"""
    job_data = {"user_id": user_id, "ism": ism}
    morning_time = time(hour=9, minute=0, tzinfo=UZ_TZ)
    evening_time = time(hour=21, minute=0, tzinfo=UZ_TZ)
    app.job_queue.run_daily(
        send_morning_reminder,
        time=morning_time,
        name=f"morning_{user_id}",
        data=job_data
    )
    app.job_queue.run_daily(
        send_evening_reminder,
        time=evening_time,
        name=f"evening_{user_id}",
        data=job_data
    )

# ─── /profil ──────────────────────────────────────────────────────────────────
async def profil(update: Update, context: ContextTypes.DEFAULT_TYPE):
    p = context.user_data.get("profile", {})
    if not p.get("ism"):
        await update.message.reply_text("❌ Avval /start orqali ro'yxatdan o'ting.")
        return
    await update.message.reply_text(
        f"👤 {p.get('ism')}\n"
        f"🏪 {p.get('biznes_turi')} biznes\n"
        f"📊 A nuqta: {p.get('a_nuqta')}\n"
        f"🎯 B nuqta: {p.get('b_nuqta')}"
    )

# ─── /cancel ──────────────────────────────────────────────────────────────────
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ Bekor qilindi.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    reg_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            REG_ISM: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_ism)],
            REG_BIZNES_TURI: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_biznes_turi)],
            REG_A: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_a)],
            REG_B: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_b)],
            REG_M1: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_m1)],
            REG_M2: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_m2)],
            REG_M3: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_m3)],
            REG_M4: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_m4)],
            REG_M5: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_m5)],
            REG_M6: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_m6)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    kunlik_handler = ConversationHandler(
        entry_points=[CommandHandler("kunlik", kunlik_start)],
        states={
            DAILY_METRIKA: [MessageHandler(filters.TEXT & ~filters.COMMAND, daily_metrika)],
            DAILY_REJA: [MessageHandler(filters.TEXT & ~filters.COMMAND, daily_reja)],
            DAILY_PLAN: [MessageHandler(filters.TEXT & ~filters.COMMAND, daily_plan)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    kechki_handler = ConversationHandler(
        entry_points=[CommandHandler("kechki", kechki_start)],
        states={
            EVE_FAKT: [MessageHandler(filters.TEXT & ~filters.COMMAND, eve_fakt)],
            EVE_DAROMAD: [MessageHandler(filters.TEXT & ~filters.COMMAND, eve_daromad)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(reg_handler)
    app.add_handler(kunlik_handler)
    app.add_handler(kechki_handler)
    app.add_handler(CommandHandler("profil", profil))

    logger.info("Bot ishga tushdi!")
    app.run_polling()

if __name__ == "__main__":
    main()
