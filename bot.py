import logging
import os
from datetime import datetime
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
BOT_TOKEN = "BU_YERGA_BOT_TOKENINGIZNI_YOZING"
SPREADSHEET_ID = "1L4wpKTkFghanh55c2_tNVD7O3CvmphhDTeJ7cKcKke8"
CREDENTIALS_FILE = "credentials.json"  # Service account JSON fayli
ADMIN_ID = 123456789  # @userinfobot dan o'z Telegram ID ingizni oling

# ─── LOGGING ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ─── CONVERSATION STATES ──────────────────────────────────────────────────────
# Ro'yxatdan o'tish
REG_NAME, REG_BIZNES = range(2)

# Kunlik to'ldirish
DAILY_METRIKA, DAILY_REJA, DAILY_PLAN, DAILY_FAKT, DAILY_DAROMAD = range(5)

# ─── GOOGLE SHEETS ULANISH ────────────────────────────────────────────────────
def get_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

def get_or_create_tab(spreadsheet, tab_name: str):
    """Userning tab'ini topadi yoki yangi yaratadi."""
    try:
        worksheet = spreadsheet.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        # Yangi tab yaratish
        worksheet = spreadsheet.add_worksheet(title=tab_name, rows=100, cols=10)
        # Sarlavha qatorlarini qo'shish
        worksheet.update("A1:H2", [
            ["30 KUNLIK BIZNES O'SISH REJASI", "", "", "", "", "", "", ""],
            ["Sana", "Qaysi metrikaga ta'sir", "B nuqta uchun reja", "Plan", "Fakt", "Kunlik daromad/sof foyda", "", ""],
        ])
        worksheet.format("A1:H1", {"textFormat": {"bold": True}})
        worksheet.format("A2:H2", {"textFormat": {"bold": True}})
    return worksheet

def append_daily_row(tab_name: str, data: dict):
    """Kunlik ma'lumotni sheetsga qo'shadi."""
    spreadsheet = get_sheet()
    worksheet = get_or_create_tab(spreadsheet, tab_name)
    all_values = worksheet.get_all_values()
    next_row = len(all_values) + 1
    worksheet.update(
        f"A{next_row}:F{next_row}",
        [[
            data["sana"],
            data["metrika"],
            data["reja"],
            data["plan"],
            data["fakt"],
            data["daromad"],
        ]]
    )

# ─── YORDAMCHI FUNKSIYALAR ────────────────────────────────────────────────────
def get_user_data(context: ContextTypes.DEFAULT_TYPE) -> dict:
    return context.user_data.get("profile", {})

def save_user_data(context: ContextTypes.DEFAULT_TYPE, key: str, value):
    if "profile" not in context.user_data:
        context.user_data["profile"] = {}
    context.user_data["profile"][key] = value

# ─── /start KOMANDASI ─────────────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = get_user_data(context)

    if profile.get("ism") and profile.get("biznes"):
        # Allaqachon ro'yxatdan o'tgan
        await update.message.reply_text(
            f"👋 Salom, {profile['ism']}!\n\n"
            f"📋 Bugungi kunlik hisobotni to'ldirish uchun /kunlik buyrug'ini yuboring.\n"
            f"ℹ️ Profilingizni ko'rish uchun /profil buyrug'ini yuboring."
        )
        return ConversationHandler.END

    await update.message.reply_text(
        "👋 Assalomu alaykum! \n\n"
        "Bu bot sizga 30 kunlik biznes o'sish rejangizni kuzatib borishga yordam beradi. 📈\n\n"
        "Avval ro'yxatdan o'tamiz.\n\n"
        "👤 Ismingizni kiriting (to'liq ism):"
    )
    return REG_NAME

# ─── RO'YXATDAN O'TISH ────────────────────────────────────────────────────────
async def reg_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ism = update.message.text.strip()
    save_user_data(context, "ism", ism)
    await update.message.reply_text(
        f"✅ Rahmat, {ism}!\n\n"
        f"🏪 Biznesingizning nomini kiriting:"
    )
    return REG_BIZNES

async def reg_biznes(update: Update, context: ContextTypes.DEFAULT_TYPE):
    biznes = update.message.text.strip()
    save_user_data(context, "biznes", biznes)
    profile = get_user_data(context)

    # Tab nomini saqlash (Ism_Familiya formatida)
    tab_name = profile["ism"][:30]  # Max 30 belgi
    save_user_data(context, "tab_name", tab_name)

    # Google Sheets da tab yaratish
    try:
        spreadsheet = get_sheet()
        get_or_create_tab(spreadsheet, tab_name)
        sheets_ok = True
    except Exception as e:
        logger.error(f"Sheets xatosi: {e}")
        sheets_ok = False

    await update.message.reply_text(
        f"🎉 Ro'yxatdan muvaffaqiyatli o'tdingiz!\n\n"
        f"👤 Ism: {profile['ism']}\n"
        f"🏪 Biznes: {biznes}\n\n"
        f"{'✅ Google Sheets da joyingiz tayyorlandi!' if sheets_ok else '⚠️ Sheets ga ulanishda xato. Admin bilan bog'laning.'}\n\n"
        f"📅 Har kuni /kunlik buyrug'i orqali hisobotingizni to'ldiring!"
    )
    return ConversationHandler.END

# ─── KUNLIK TO'LDIRISH ────────────────────────────────────────────────────────
async def kunlik_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = get_user_data(context)

    if not profile.get("ism"):
        await update.message.reply_text(
            "❌ Avval ro'yxatdan o'ting! /start buyrug'ini yuboring."
        )
        return ConversationHandler.END

    keyboard = [
        ["📊 Mehmonlar soni", "🛒 Xaridorlar soni"],
        ["💰 O'rtacha chek", "📈 Marja"],
        ["🔄 Qayta sotuv"],
    ]
    await update.message.reply_text(
        f"📅 Bugun: {datetime.now().strftime('%d.%m.%Y')}\n\n"
        f"Qaysi metrikaga ta'sir qilmoqchisiz?",
        reply_markup=ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
    )
    return DAILY_METRIKA

async def daily_metrika(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["daily"] = {"metrika": update.message.text.strip()}
    await update.message.reply_text(
        "📝 B nuqtaga yetish uchun bugun nima qilasiz?\n\n"
        "Masalan: Yangi mahsulot qo'ydim, chegirma e'lon qildim...",
        reply_markup=ReplyKeyboardRemove()
    )
    return DAILY_REJA

async def daily_reja(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["daily"]["reja"] = update.message.text.strip()
    await update.message.reply_text(
        "🎯 Bugungi PLAN (maqsad):\n\n"
        "Masalan: 120 ta mehmon, 75,000 so'm chek..."
    )
    return DAILY_PLAN

async def daily_plan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["daily"]["plan"] = update.message.text.strip()
    await update.message.reply_text(
        "✅ Bugungi FAKT (haqiqatda bo'lgan):\n\n"
        "Masalan: 115 ta mehmon, 72,000 so'm chek..."
    )
    return DAILY_FAKT

async def daily_fakt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["daily"]["fakt"] = update.message.text.strip()
    await update.message.reply_text(
        "💵 Bugungi kunlik aylanma / sof foyda (so'mda):\n\n"
        "Masalan: 1,200,000 so'm"
    )
    return DAILY_DAROMAD

async def daily_daromad(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["daily"]["daromad"] = update.message.text.strip()
    daily = context.user_data["daily"]
    profile = get_user_data(context)

    data = {
        "sana": datetime.now().strftime("%d.%m.%Y"),
        "metrika": daily["metrika"],
        "reja": daily["reja"],
        "plan": daily["plan"],
        "fakt": daily["fakt"],
        "daromad": daily["daromad"],
    }

    # Sheetsga yozish
    try:
        append_daily_row(profile["tab_name"], data)
        sheets_msg = "✅ Ma'lumotlar Google Sheets ga saqlandi!"
    except Exception as e:
        logger.error(f"Sheets xatosi: {e}")
        sheets_msg = "⚠️ Sheets ga saqlashda xato yuz berdi. Admin bilan bog'laning."

    await update.message.reply_text(
        f"🎉 Bugungi hisobot saqlandi!\n\n"
        f"📅 Sana: {data['sana']}\n"
        f"📊 Metrika: {data['metrika']}\n"
        f"📝 Reja: {data['reja']}\n"
        f"🎯 Plan: {data['plan']}\n"
        f"✅ Fakt: {data['fakt']}\n"
        f"💵 Daromad: {data['daromad']}\n\n"
        f"{sheets_msg}\n\n"
        f"Ertaga ham /kunlik buyrug'ini unutmang! 💪"
    )
    return ConversationHandler.END

# ─── /profil KOMANDASI ────────────────────────────────────────────────────────
async def profil(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = get_user_data(context)
    if not profile.get("ism"):
        await update.message.reply_text("❌ Avval /start orqali ro'yxatdan o'ting.")
        return
    await update.message.reply_text(
        f"👤 Profilingiz:\n\n"
        f"Ism: {profile.get('ism', '-')}\n"
        f"Biznes: {profile.get('biznes', '-')}\n"
        f"Sheets tab: {profile.get('tab_name', '-')}"
    )

# ─── /admin KOMANDASI (faqat siz uchun) ──────────────────────────────────────
async def admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Sizda bu buyruqqa ruxsat yo'q.")
        return
    await update.message.reply_text(
        f"👑 Admin panel\n\n"
        f"Google Sheets: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}\n\n"
        f"Barcha foydalanuvchilar ma'lumotlari shu linkda ko'rinadi."
    )

# ─── CANCEL ───────────────────────────────────────────────────────────────────
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "❌ Bekor qilindi. /kunlik yoki /start buyrug'ini yuboring.",
        reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    # Ro'yxatdan o'tish conversation
    reg_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            REG_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_name)],
            REG_BIZNES: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_biznes)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    # Kunlik to'ldirish conversation
    kunlik_handler = ConversationHandler(
        entry_points=[CommandHandler("kunlik", kunlik_start)],
        states={
            DAILY_METRIKA: [MessageHandler(filters.TEXT & ~filters.COMMAND, daily_metrika)],
            DAILY_REJA: [MessageHandler(filters.TEXT & ~filters.COMMAND, daily_reja)],
            DAILY_PLAN: [MessageHandler(filters.TEXT & ~filters.COMMAND, daily_plan)],
            DAILY_FAKT: [MessageHandler(filters.TEXT & ~filters.COMMAND, daily_fakt)],
            DAILY_DAROMAD: [MessageHandler(filters.TEXT & ~filters.COMMAND, daily_daromad)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(reg_handler)
    app.add_handler(kunlik_handler)
    app.add_handler(CommandHandler("profil", profil))
    app.add_handler(CommandHandler("admin", admin))

    logger.info("Bot ishga tushdi...")
    app.run_polling()

if __name__ == "__main__":
    main() 