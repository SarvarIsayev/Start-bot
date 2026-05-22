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
(REG_ISM, REG_BIZNES_TURI, REG_A, REG_B, REG_METRIKALARI,
 DAILY_METRIKA, DAILY_REJA, DAILY_PLAN,
 EVE_FAKT, EVE_DAROMAD) = range(10)

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

def get_or_create_tab(spreadsheet, tab_name: str):
    try:
        ws = spreadsheet.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=tab_name, rows=200, cols=15)
        ws.update("A1:F1", [["Sana", "Metrika", "B nuqta uchun reja", "Plan", "Fakt", "Daromad/Sof foyda"]])
        ws.format("A1:F1", {"textFormat": {"bold": True}})
    return ws

def save_registration(tab_name: str, profile: dict):
    spreadsheet = get_sheet()
    try:
        ws = spreadsheet.worksheet("Ro'yxat")
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title="Ro'yxat", rows=200, cols=30)
        ws.update("A1:D1", [["Ism Familiya", "Biznes turi", "A nuqta", "B nuqta"]])
        ws.format("A1:Z1", {"textFormat": {"bold": True}})

    metrikalari = get_metrikalari(profile["biznes_turi"])
    row = [profile["ism"], profile["biznes_turi"], profile["a_nuqta"], profile["b_nuqta"]]

    saved_metrikalari = profile.get("saved_metrikalari", {})
    for m in metrikalari:
        row.append(saved_metrikalari.get(f"{m}_a", ""))
        row.append(saved_metrikalari.get(f"{m}_b", ""))

    all_vals = ws.get_all_values()
    ws.update(f"A{len(all_vals)+1}", [row])
    get_or_create_tab(spreadsheet, tab_name)

def save_morning(tab_name: str, data: dict):
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
            "/kunlik — ertalabki hisobot\n"
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
    context.user_data["profile"] = {"ism": update.message.text.strip()}
    kb = [["Chakana", "Service", "Distributsiya"]]
    await update.message.reply_text(
        "2️⃣ Biznesingiz turi:",
        reply_markup=ReplyKeyboardMarkup(kb, one_time_keyboard=True, resize_keyboard=True)
    )
    return REG_BIZNES_TURI

async def reg_biznes_turi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    turi = update.message.text.strip()
    if turi not in ["Chakana", "Service", "Distributsiya"]:
        kb = [["Chakana", "Service", "Distributsiya"]]
        await update.message.reply_text(
            "Iltimos, tugmalardan birini tanlang:",
            reply_markup=ReplyKeyboardMarkup(kb, one_time_keyboard=True, resize_keyboard=True)
        )
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
    context.user_data["metrika_idx"] = 0
    context.user_data["metrika_is_a"] = True
    context.user_data["profile"]["saved_metrikalari"] = {}

    m = metrikalari[0]
    await update.message.reply_text(
        f"Endi {biznes_turi} biznesining hozirgi holatini aniqlaymiz.\n\n"
        f"5️⃣ {m} — A nuqtada qancha?"
    )
    return REG_METRIKALARI

async def reg_metrikalari(update: Update, context: ContextTypes.DEFAULT_TYPE):
    metrikalari = context.user_data["metrikalari"]
    idx = context.user_data["metrika_idx"]
    is_a = context.user_data["metrika_is_a"]
    m = metrikalari[idx]

    key = f"{m}_{'a' if is_a else 'b'}"
    context.user_data["profile"]["saved_metrikalari"][key] = update.message.text.strip()

    if is_a:
        # B nuqtani so'rash
        context.user_data["metrika_is_a"] = False
        await update.message.reply_text(f"{m} — B nuqtada qancha?")
        return REG_METRIKALARI
    else:
        next_idx = idx + 1
        if next_idx < len(metrikalari):
            # Keyingi metrika
            context.user_data["metrika_idx"] = next_idx
            context.user_data["metrika_is_a"] = True
            next_m = metrikalari[next_idx]
            await update.message.reply_text(f"{next_m} — A nuqtada qancha?")
            return REG_METRIKALARI
        else:
            # Hammasi tugadi — saqlash
            profile = context.user_data["profile"]
            tab_name = profile["ism"][:25]
            context.user_data["profile"]["tab_name"] = tab_name
            try:
                save_registration(tab_name, profile)
                msg = "✅ Google Sheets ga saqlandi!"
            except Exception as e:
                logger.error(e)
                msg = "⚠️ Sheets ga saqlashda xato. Admin bilan bog'laning."

            # Eslatmalarni rejalashtirish
            user_id = update.effective_user.id
            schedule_reminders(context.application, user_id, profile["ism"])

            await update.message.reply_text(
                f"🎉 Ro'yxatdan muvaffaqiyatli o'tdingiz!\n\n"
                f"👤 {profile['ism']}\n"
                f"🏪 {profile['biznes_turi']} biznes\n"
                f"📊 A nuqta: {profile['a_nuqta']}\n"
                f"🎯 B nuqta: {profile['b_nuqta']}\n\n"
                f"{msg}\n\n"
                f"Har kuni ertalab soat 9:00 da savol keladi! 📅"
            )
            return ConversationHandler.END

# ─── ERTALABKI HISOBOT ────────────────────────────────────────────────────────
async def kunlik_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = context.user_data.get("profile", {})
    if not profile.get("ism"):
        await update.message.reply_text("❌ Avval /start orqali ro'yxatdan o'ting.")
        return ConversationHandler.END
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    context.user_data["daily_sana"] = today
    await update.message.reply_text(
        f"🌅 Ertalabki hisobot\n\n"
        f"📅 Bugungi sana: {today}\n\n"
        f"Qaysi metrikaga bugun ta'sir qilmoqchisiz?\n"
        f"(Masalan: Mehmonlar soni, O'rtacha chek...)"
    )
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
        f"Kechqurun soat 9:00 da /kechki buyrug'ini yuboring. 💪\n\n{msg}"
    )
    return ConversationHandler.END

# ─── KECHKI HISOBOT ───────────────────────────────────────────────────────────
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

# ─── AVTOMATIK ESLATMALAR ─────────────────────────────────────────────────────
async def send_morning_reminder(context: ContextTypes.DEFAULT_TYPE):
    job_data = context.job.data
    user_id = job_data["user_id"]
    ism = job_data["ism"]
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    if datetime.now(UZ_TZ).weekday() == 6:
        return
    await context.bot.send_message(
        chat_id=user_id,
        text=f"🌅 Xayrli tong, {ism}!\n\n📅 Bugun: {today}\n\nErtalabki hisobotni to'ldiring 👇\n/kunlik"
    )

async def send_evening_reminder(context: ContextTypes.DEFAULT_TYPE):
    job_data = context.job.data
    user_id = job_data["user_id"]
    ism = job_data["ism"]
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    if datetime.now(UZ_TZ).weekday() == 6:
        return
    await context.bot.send_message(
        chat_id=user_id,
        text=f"🌆 Kechqurun eslatma, {ism}!\n\n📅 Bugun: {today}\n\nKechki hisobotni to'ldiring 👇\n/kechki"
    )

def schedule_reminders(app, user_id: int, ism: str):
    job_data = {"user_id": user_id, "ism": ism}
    morning_time = time(hour=9, minute=0, tzinfo=UZ_TZ)
    evening_time = time(hour=21, minute=0, tzinfo=UZ_TZ)
    # Eski joblarni o'chirish
    current_jobs = app.job_queue.get_jobs_by_name(f"morning_{user_id}")
    for job in current_jobs:
        job.schedule_removal()
    current_jobs = app.job_queue.get_jobs_by_name(f"evening_{user_id}")
    for job in current_jobs:
        job.schedule_removal()
    app.job_queue.run_daily(send_morning_reminder, time=morning_time, name=f"morning_{user_id}", data=job_data)
    app.job_queue.run_daily(send_evening_reminder, time=evening_time, name=f"evening_{user_id}", data=job_data)

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
            REG_METRIKALARI: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_metrikalari)],
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

