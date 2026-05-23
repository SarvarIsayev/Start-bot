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
from gspread.utils import rowcol_to_a1

UZ_TZ = pytz.timezone("Asia/Tashkent")
BOT_TOKEN = os.getenv("BOT_TOKEN", "BU_YERGA_TOKEN")
ADMIN_ID = int(os.getenv("ADMIN_ID", "123456789"))
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1L4wpKTkFghanh55c2_tNVD7O3CvmphhDTeJ7cKcKke8")

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ─── STATES ───────────────────────────────────────────────────────────────────
(REG_ISM, REG_BIZNES_NOMI, REG_BIZNES_TURI, REG_A, REG_B, REG_METRIKALARI,
 DAILY_METRIKA, DAILY_REJA, DAILY_PLAN,
 EVE_FAKT, EVE_DAROMAD) = range(11)

# ─── METRIKALAR ───────────────────────────────────────────────────────────────
CHAKANA = ["Mehmonlar", "Xaridorlar", "O'rt.Chek", "Marja", "Qayta sotuv"]
SERVICE  = ["SMM ko'rish", "Qo'ng'iroq", "Uchrashuv", "Xaridor", "O'rt.Chek", "Marja"]
DISTRIB  = ["OKB", "AKB", "O'rt.Chek", "Qayta sotuv", "Marja"]

BIZNES_EMOJI = {"Chakana": "🛒", "Service": "🎓", "Distributsiya": "🚚"}

# Ranglar (hex -> RGB tuple)
COLORS = {
    "Chakana": {
        "header_bg":  (0.02, 0.27, 0.36),   # teal dark
        "header_fg":  (1, 1, 1),
        "label_bg":   (0.02, 0.27, 0.36),
        "meta_bg":    (0.68, 0.85, 0.90),
        "meta_fg":    (0.02, 0.18, 0.27),
        "row_odd":    (0.85, 0.94, 0.97),
        "row_even":   (0.93, 0.97, 0.99),
        "daily_odd":  (0.85, 0.94, 0.97),
        "daily_even": (0.93, 1.00, 0.95),
    },
    "Service": {
        "header_bg":  (0.20, 0.08, 0.35),
        "header_fg":  (1, 1, 1),
        "label_bg":   (0.20, 0.08, 0.35),
        "meta_bg":    (0.80, 0.72, 0.92),
        "meta_fg":    (0.15, 0.05, 0.28),
        "row_odd":    (0.87, 0.82, 0.95),
        "row_even":   (0.93, 0.90, 0.98),
        "daily_odd":  (0.87, 0.82, 0.95),
        "daily_even": (0.93, 1.00, 0.95),
    },
    "Distributsiya": {
        "header_bg":  (0.40, 0.20, 0.04),
        "header_fg":  (1, 1, 1),
        "label_bg":   (0.40, 0.20, 0.04),
        "meta_bg":    (0.95, 0.80, 0.60),
        "meta_fg":    (0.30, 0.13, 0.02),
        "row_odd":    (0.97, 0.87, 0.72),
        "row_even":   (0.99, 0.93, 0.85),
        "daily_odd":  (0.97, 0.87, 0.72),
        "daily_even": (0.93, 1.00, 0.95),
    },
}

def rgb(r, g, b):
    return {"red": r, "green": g, "blue": b}

def get_metrikalari(turi):
    return {"Chakana": CHAKANA, "Service": SERVICE, "Distributsiya": DISTRIB}.get(turi, CHAKANA)

# ─── GOOGLE SHEETS ────────────────────────────────────────────────────────────
def get_sheet():
    creds_json = os.getenv("GOOGLE_CREDENTIALS")
    if creds_json:
        info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(info, scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"])
    else:
        creds = Credentials.from_service_account_file(
            "credentials.json", scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)

def create_user_tab(spreadsheet, profile: dict):
    turi = profile["biznes_turi"]
    metrikalari = get_metrikalari(turi)
    c = COLORS[turi]
    emoji = BIZNES_EMOJI[turi]
    tab_name = profile["tab_name"]
    saved = profile.get("saved_metrikalari", {})
    ism = profile["ism"]
    biznes_nomi = profile.get("biznes_nomi", "")
    a_nuqta = profile["a_nuqta"]
    b_nuqta = profile["b_nuqta"]
    n = len(metrikalari)  # metrikalar soni

    # Tab yaratish
    try:
        ws = spreadsheet.worksheet(tab_name)
        spreadsheet.del_worksheet(ws)
    except:
        pass
    ws = spreadsheet.add_worksheet(title=tab_name, rows=200, cols=max(n+2, 8))

    # ── 1-qator: Sarlavha ──
    title_text = f"30 KUNLIK BIZNES O'SISH REJASI  |  {emoji} {turi.upper()}"
    ws.update("A1", [[title_text]])

    # ── 2-qator: Tadbirkor / Biznes nomi / A nuqta / B nuqta ──
    ws.update("A2:H2", [[
        "Tadbirkor:", ism,
        "Biznes nomi:", biznes_nomi,
        "A nuqtasi (hozir):", a_nuqta,
        "B nuqtasi (maqsad):", b_nuqta
    ]])

    # ── 3-qator: Metrika nomlari ──
    row3 = ["HOZIRGI HOLAT (A)"] + [f"{m}" for m in metrikalari]
    ws.update(f"A3", [row3[:n+1]])

    # ── 4-qator: A nuqta qiymatlari ──
    row4 = ["A nuqta"] + [saved.get(f"{m}_a", "") for m in metrikalari]
    ws.update(f"A4", [row4[:n+1]])

    # ── 5-qator: B nuqta qiymatlari ──
    row5 = ["B nuqta"] + [saved.get(f"{m}_b", "") for m in metrikalari]
    ws.update(f"A5", [row5[:n+1]])

    # ── 7-qator: Kunlik sarlavhalar ──
    ws.update("A7:F7", [["Sana", "Qaysi metrikaga ta'sir qilmoqchi", "B nuqta uchun nima qiladi", "Plan", "Fakt", "Kunlik abarot/sof foyda"]])

    # ── FORMATLAR ──
    last_col = max(n + 1, 6)
    last_col_letter = chr(64 + last_col) if last_col <= 26 else "H"

    requests = [
        # 1-qator merge va rang
        {"mergeCells": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
            "startColumnIndex": 0, "endColumnIndex": last_col}, "mergeType": "MERGE_ALL"}},
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
            "startColumnIndex": 0, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb(*c["header_bg"]),
                "textFormat": {"foregroundColor": rgb(*c["header_fg"]), "bold": True, "fontSize": 14},
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat"}},

        # 2-qator: label bg
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2,
            "startColumnIndex": 0, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb(*c["meta_bg"]),
                "textFormat": {"foregroundColor": rgb(*c["meta_fg"]), "bold": False}}},
            "fields": "userEnteredFormat"}},
        # 2-qator label ustunlari bold
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2,
            "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
            "fields": "userEnteredFormat.textFormat"}},

        # 3-qator: header
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 2, "endRowIndex": 3,
            "startColumnIndex": 0, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb(*c["label_bg"]),
                "textFormat": {"foregroundColor": rgb(*c["header_fg"]), "bold": True},
                "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat"}},

        # 4-qator: A nuqta
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 3, "endRowIndex": 4,
            "startColumnIndex": 0, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb(*c["row_odd"]),
                "textFormat": {"bold": False},
                "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat"}},
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 3, "endRowIndex": 4,
            "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
            "fields": "userEnteredFormat.textFormat"}},

        # 5-qator: B nuqta (yashil matn)
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 4, "endRowIndex": 5,
            "startColumnIndex": 0, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb(*c["row_even"]),
                "textFormat": {"foregroundColor": rgb(0.0, 0.39, 0.0), "bold": True},
                "horizontalAlignment": "CENTER"}},
            "fields": "userEnteredFormat"}},
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 4, "endRowIndex": 5,
            "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"foregroundColor": rgb(0,0,0), "bold": True}}},
            "fields": "userEnteredFormat.textFormat"}},

        # 6-qator: bo'sh separator
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 5, "endRowIndex": 6,
            "startColumnIndex": 0, "endColumnIndex": last_col},
            "cell": {"userEnteredFormat": {"backgroundColor": rgb(*c["header_bg"])}},
            "fields": "userEnteredFormat"}},

        # 7-qator: kunlik sarlavha
        {"repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": 6, "endRowIndex": 7,
            "startColumnIndex": 0, "endColumnIndex": 6},
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb(*c["header_bg"]),
                "textFormat": {"foregroundColor": rgb(*c["header_fg"]), "bold": True},
                "horizontalAlignment": "CENTER", "wrapStrategy": "WRAP"}},
            "fields": "userEnteredFormat"}},

        # 1-qator balandligi
        {"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "ROWS",
            "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 50}, "fields": "pixelSize"}},
        # 7-qator balandligi
        {"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "ROWS",
            "startIndex": 6, "endIndex": 7},
            "properties": {"pixelSize": 50}, "fields": "pixelSize"}},

        # A ustun kengligi
        {"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS",
            "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 140}, "fields": "pixelSize"}},

        # Chegara chiziqlar (7-qator)
        {"updateBorders": {"range": {"sheetId": ws.id, "startRowIndex": 6, "endRowIndex": 7,
            "startColumnIndex": 0, "endColumnIndex": 6},
            "innerVertical": {"style": "SOLID", "color": rgb(1,1,1), "width": 1},
            "bottom": {"style": "SOLID", "color": rgb(1,1,1), "width": 1}}},
    ]

    # Kunlik qatorlar uchun ranglar (8-37 qatorlar)
    for i in range(30):
        row_i = 7 + i
        bg = c["daily_odd"] if i % 2 == 0 else c["daily_even"]
        requests.append({
            "repeatCell": {"range": {"sheetId": ws.id, "startRowIndex": row_i, "endRowIndex": row_i+1,
                "startColumnIndex": 0, "endColumnIndex": 6},
                "cell": {"userEnteredFormat": {"backgroundColor": rgb(*bg)}},
                "fields": "userEnteredFormat"}})

    spreadsheet.batch_update({"requests": requests})
    return ws

def save_registration(profile: dict):
    spreadsheet = get_sheet()

    # Ro'yxat tab
    try:
        ws_reg = spreadsheet.worksheet("Ro'yxat")
    except:
        ws_reg = spreadsheet.add_worksheet(title="Ro'yxat", rows=200, cols=10)
        ws_reg.update("A1:E1", [["Ism Familiya", "Biznes nomi", "Biznes turi", "A nuqta", "B nuqta"]])
        ws_reg.format("A1:E1", {"textFormat": {"bold": True}})

    all_vals = ws_reg.get_all_values()
    ws_reg.update(f"A{len(all_vals)+1}:E{len(all_vals)+1}", [[
        profile["ism"], profile.get("biznes_nomi",""),
        profile["biznes_turi"], profile["a_nuqta"], profile["b_nuqta"]
    ]])

    # Foydalanuvchi tab yaratish
    create_user_tab(spreadsheet, profile)

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
    await update.message.reply_text("2️⃣ Biznesingizning nomini kiriting (masalan: Oziq-ovqat do'koni):")
    return REG_BIZNES_NOMI

async def reg_biznes_nomi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["profile"]["biznes_nomi"] = update.message.text.strip()
    kb = [["Chakana", "Service", "Distributsiya"]]
    await update.message.reply_text(
        "3️⃣ Biznesingiz turi:",
        reply_markup=ReplyKeyboardMarkup(kb, one_time_keyboard=True, resize_keyboard=True)
    )
    return REG_BIZNES_TURI

async def reg_biznes_turi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    turi = update.message.text.strip()
    if turi not in ["Chakana", "Service", "Distributsiya"]:
        kb = [["Chakana", "Service", "Distributsiya"]]
        await update.message.reply_text("Iltimos, tugmalardan birini tanlang:",
            reply_markup=ReplyKeyboardMarkup(kb, one_time_keyboard=True, resize_keyboard=True))
        return REG_BIZNES_TURI
    context.user_data["profile"]["biznes_turi"] = turi
    await update.message.reply_text(
        "4️⃣ A nuqtangiz qancha? (hozirgi oylik daromadingiz):",
        reply_markup=ReplyKeyboardRemove()
    )
    return REG_A

async def reg_a(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["profile"]["a_nuqta"] = update.message.text.strip()
    await update.message.reply_text("5️⃣ B nuqtangiz qancha? (30 kundan keyin bo'lishi kerak bo'lgan daromad):")
    return REG_B

async def reg_b(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["profile"]["b_nuqta"] = update.message.text.strip()
    turi = context.user_data["profile"]["biznes_turi"]
    metrikalari = get_metrikalari(turi)
    context.user_data["metrikalari"] = metrikalari
    context.user_data["metrika_idx"] = 0
    context.user_data["metrika_is_a"] = True
    context.user_data["profile"]["saved_metrikalari"] = {}
    await update.message.reply_text(
        f"Endi {turi} biznesining hozirgi holatini aniqlaymiz.\n\n"
        f"6️⃣ {metrikalari[0]} — A nuqtada qancha?"
    )
    return REG_METRIKALARI

async def reg_metrikalari(update: Update, context: ContextTypes.DEFAULT_TYPE):
    metrikalari = context.user_data["metrikalari"]
    idx = context.user_data["metrika_idx"]
    is_a = context.user_data["metrika_is_a"]
    m = metrikalari[idx]
    context.user_data["profile"]["saved_metrikalari"][f"{m}_{'a' if is_a else 'b'}"] = update.message.text.strip()

    if is_a:
        context.user_data["metrika_is_a"] = False
        await update.message.reply_text(f"{m} — B nuqtada qancha?")
        return REG_METRIKALARI
    else:
        next_idx = idx + 1
        if next_idx < len(metrikalari):
            context.user_data["metrika_idx"] = next_idx
            context.user_data["metrika_is_a"] = True
            await update.message.reply_text(f"{metrikalari[next_idx]} — A nuqtada qancha?")
            return REG_METRIKALARI
        else:
            profile = context.user_data["profile"]
            tab_name = profile["ism"][:25]
            context.user_data["profile"]["tab_name"] = tab_name
            try:
                save_registration(profile)
                msg = "✅ Google Sheets ga chiroyli formatda saqlandi!"
            except Exception as e:
                logger.error(e)
                msg = f"⚠️ Sheets xatosi: {e}"

            schedule_reminders(context.application, update.effective_user.id, profile["ism"])

            # bot_data ga saqlash (admin ko'rishi uchun)
            if "all_users" not in context.bot_data:
                context.bot_data["all_users"] = {}
            context.bot_data["all_users"][update.effective_user.id] = {
                "ism": profile["ism"],
                "biznes_nomi": profile.get("biznes_nomi", ""),
                "biznes_turi": profile["biznes_turi"],
                "a_nuqta": profile["a_nuqta"],
                "b_nuqta": profile["b_nuqta"],
                "bugun": {},
            }

            await update.message.reply_text(
                f"🎉 Ro'yxatdan muvaffaqiyatli o'tdingiz!\n\n"
                f"👤 {profile['ism']}\n"
                f"🏪 {profile.get('biznes_nomi','')} ({profile['biznes_turi']})\n"
                f"📊 A nuqta: {profile['a_nuqta']}\n"
                f"🎯 B nuqta: {profile['b_nuqta']}\n\n"
                f"{msg}\n\n"
                f"Har kuni ertalab soat 9:00 da savol keladi! 📅"
            )
            return ConversationHandler.END

# ─── KUNLIK ───────────────────────────────────────────────────────────────────
async def kunlik_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = context.user_data.get("profile", {})
    if not profile.get("ism"):
        await update.message.reply_text("❌ Avval /start orqali ro'yxatdan o'ting.")
        return ConversationHandler.END
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    context.user_data["daily_sana"] = today
    await update.message.reply_text(
        f"🌅 Ertalabki hisobot\n📅 Bugun: {today}\n\n"
        f"Qaysi metrikaga bugun ta'sir qilmoqchisiz?"
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
    # bot_data ga bugungi vazifani saqlash (admin ko'rishi uchun)
    uid = update.effective_user.id
    if "all_users" not in context.bot_data:
        context.bot_data["all_users"] = {}
    if uid not in context.bot_data["all_users"]:
        context.bot_data["all_users"][uid] = {"ism": profile.get("ism",""), "bugun": {}}
    context.bot_data["all_users"][uid]["bugun"] = {
        "sana": data["sana"],
        "metrika": data["metrika"],
        "reja": data["reja"],
        "plan": data["plan"],
    }

    try:
        save_morning(profile["tab_name"], data)
        msg = "✅ Saqlandi!"
    except Exception as e:
        logger.error(e)
        msg = "⚠️ Xato yuz berdi."
    await update.message.reply_text(f"✅ Ertalabki hisobot qabul qilindi!\n\nKechqurun soat 9:00 da /kechki yuboring. 💪\n{msg}")
    return ConversationHandler.END

# ─── KECHKI ───────────────────────────────────────────────────────────────────
async def kechki_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    profile = context.user_data.get("profile", {})
    if not profile.get("ism"):
        await update.message.reply_text("❌ Avval /start orqali ro'yxatdan o'ting.")
        return ConversationHandler.END
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    context.user_data["daily_sana"] = today
    await update.message.reply_text(f"🌆 Kechki hisobot — {today}\n\n✅ Bugungi FAKT (raqamda):")
    return EVE_FAKT

async def eve_fakt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["eve_fakt"] = update.message.text.strip()
    await update.message.reply_text("💵 Bugungi kunlik aylanma yoki sof foyda:")
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
    await update.message.reply_text(f"🎉 Bugungi hisobot to'liq saqlandi! Ertaga ham /kunlik kutamiz. 💪\n{msg}")
    return ConversationHandler.END

# ─── ESLATMALAR ───────────────────────────────────────────────────────────────
async def send_morning_reminder(context: ContextTypes.DEFAULT_TYPE):
    d = context.job.data
    if datetime.now(UZ_TZ).weekday() == 6:
        return
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    await context.bot.send_message(chat_id=d["user_id"],
        text=f"🌅 Xayrli tong, {d['ism']}!\n\n📅 Bugun: {today}\n\nErtalabki hisobotni to'ldiring 👇\n/kunlik")

async def send_evening_reminder(context: ContextTypes.DEFAULT_TYPE):
    d = context.job.data
    if datetime.now(UZ_TZ).weekday() == 6:
        return
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    await context.bot.send_message(chat_id=d["user_id"],
        text=f"🌆 Kechqurun eslatma, {d['ism']}!\n\n📅 Bugun: {today}\n\nKechki hisobotni to'ldiring 👇\n/kechki")

def schedule_reminders(app, user_id: int, ism: str):
    data = {"user_id": user_id, "ism": ism}
    for name in [f"morning_{user_id}", f"evening_{user_id}"]:
        for job in app.job_queue.get_jobs_by_name(name):
            job.schedule_removal()
    app.job_queue.run_daily(send_morning_reminder, time=time(9, 0, tzinfo=UZ_TZ), name=f"morning_{user_id}", data=data)
    app.job_queue.run_daily(send_evening_reminder, time=time(21, 0, tzinfo=UZ_TZ), name=f"evening_{user_id}", data=data)

# ─── /profil ──────────────────────────────────────────────────────────────────
async def profil(update: Update, context: ContextTypes.DEFAULT_TYPE):
    p = context.user_data.get("profile", {})
    if not p.get("ism"):
        await update.message.reply_text("❌ Avval /start orqali ro'yxatdan o'ting.")
        return
    await update.message.reply_text(
        f"👤 {p.get('ism')}\n"
        f"🏪 {p.get('biznes_nomi','')} ({p.get('biznes_turi','')})\n"
        f"📊 A nuqta: {p.get('a_nuqta')}\n"
        f"🎯 B nuqta: {p.get('b_nuqta')}"
    )

async def users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin: barcha ishtirokchilar ro'yxati — Sheets dan o'qiladi"""
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Sizda bu buyruqqa ruxsat yo'q.")
        return
    try:
        spreadsheet = get_sheet()
        ws = spreadsheet.worksheet("Ro'yxat")
        all_vals = ws.get_all_values()
        if len(all_vals) <= 1:
            await update.message.reply_text("📋 Hali hech kim ro'yxatdan o'tmagan.")
            return
        rows = all_vals[1:]
        text = f"👥 Ishtirokchilar ro'yxati — {len(rows)} ta:\n\n"
        for i, row in enumerate(rows, 1):
            ism         = row[0] if len(row) > 0 else "—"
            biznes_nomi = row[1] if len(row) > 1 else "—"
            biznes_turi = row[2] if len(row) > 2 else "—"
            a           = row[3] if len(row) > 3 else "—"
            b           = row[4] if len(row) > 4 else "—"
            text += (
                f"{i}. 👤 {ism}\n"
                f"   🏪 {biznes_nomi} ({biznes_turi})\n"
                f"   📊 A: {a}  →  🎯 B: {b}\n\n"
            )
        await update.message.reply_text(text)
    except Exception as e:
        logger.error(e)
        await update.message.reply_text(f"⚠️ Xato: {e}")

async def bugun(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin: bugungi barcha ishtirokchilarning vazifalari"""
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Sizda bu buyruqqa ruxsat yo'q.")
        return
    today = datetime.now(UZ_TZ).strftime("%d.%m.%Y")
    all_users = context.bot_data.get("all_users", {})
    if not all_users:
        await update.message.reply_text("📋 Hali hech kim bugun hisobot to'ldirmagan.")
        return
    text = f"📅 {today} — Bugungi hisobotlar:\n\n"
    topshirdi = 0
    topshirmadi_list = []
    for uid, udata in all_users.items():
        ism = udata.get("ism", "—")
        bugun_data = udata.get("bugun", {})
        if bugun_data.get("sana") == today:
            topshirdi += 1
            text += (
                f"✅ {ism}\n"
                f"   📊 Metrika: {bugun_data.get('metrika','—')}\n"
                f"   📝 Reja: {bugun_data.get('reja','—')}\n"
                f"   🎯 Plan: {bugun_data.get('plan','—')}\n\n"
            )
        else:
            topshirmadi_list.append(ism)
    if topshirmadi_list:
        text += f"❌ Hali to'ldirmagan ({len(topshirmadi_list)} ta):\n"
        for ism in topshirmadi_list:
            text += f"   • {ism}\n"
    await update.message.reply_text(text)

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
            REG_BIZNES_NOMI: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_biznes_nomi)],
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
    app.add_handler(CommandHandler("users", users))
    app.add_handler(CommandHandler("bugun", bugun))

    logger.info("Bot ishga tushdi!")
    app.run_polling()

if __name__ == "__main__":
    main()
