"""
ربات پیام‌رسان بله (Bale) مدیریت مستندات بارنامه - پی‌بار
نسخه ۴ — با منوی Reply ثابت پایین صفحه

توجه‌های مهم برای اجرا روی بله:
──────────────────────────────
1. توکن ربات را از @Bot_Father داخل خودِ پیام‌رسان بله بگیرید (نه تلگرام)
   و در متغیر محیطی BOT_TOKEN قرار دهید.
2. کتابخانه‌ی python-telegram-bot به‌صورت پیش‌فرض به سرور تلگرام وصل می‌شود؛
   در تابع main() آدرس پایه به سرور بله (tapi.bale.ai) تغییر داده شده است
   — این روش دقیقاً همان چیزی است که خودِ تیم بله در نمونه‌کدهای رسمی‌شان
   (bale-bot-samples) برای استفاده از python-telegram-bot پیشنهاد داده‌اند.
3. API بله زیرمجموعه‌ای از API تلگرام و «تا حد زیادی» سازگار با آن است، اما
   ۱۰۰٪ یکسان نیست. مواردی که ممکن است نیاز به تست/تنظیم داشته باشند:
   - برخی حالت‌های parse_mode (Markdown/HTML) ممکن است در بله رفتار
     کمی متفاوتی داشته باشند.
   - لینک‌های ویژه‌ی تلگرام مثل tg://user?id=... در بله کار نمی‌کنند؛
     به همین دلیل در این نسخه به‌جای لینک کلیک‌پذیر، آیدی راننده به‌صورت
     متن قابل‌کپی (`کد`) نمایش داده می‌شود.
   - محدودیت حجم فایل/عکس و نرخ ارسال پیام ممکن است با تلگرام فرق داشته باشد.
4. پیشنهاد می‌شود قبل از استفاده‌ی نهایی، کل سناریوها (آپلود، تایید،
   عدم تایید، کسری بار، پاسخ راننده) را یک‌بار کامل روی بله تست کنید.
"""

import os
import json
import uuid
import hashlib
import logging
from datetime import datetime, timedelta
from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup,
    ReplyKeyboardMarkup, ReplyKeyboardRemove, KeyboardButton
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ConversationHandler, filters, ContextTypes, TypeHandler, ApplicationHandlerStop
)
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    handlers=[
        logging.FileHandler("bot.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ─── محافظ ضد ارسال تکراری ───
# اگر پلتفرم بله (یا مکانیزم polling) یک آپدیت را بیش از یک‌بار تحویل بدهد،
# این محافظ با شناسه‌ی یکتای هر آپدیت (update_id) جلوی پردازش دوباره‌اش را می‌گیرد.
_seen_update_ids: set = set()
_seen_update_ids_order: list = []
_MAX_SEEN_UPDATE_IDS = 5000

async def dedupe_update_guard(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.update_id
    if uid in _seen_update_ids:
        logger.warning(f"⚠️ آپدیت تکراری نادیده گرفته شد — update_id={uid} (این آپدیت قبلاً پردازش شده بود)")
        raise ApplicationHandlerStop
    _seen_update_ids.add(uid)
    _seen_update_ids_order.append(uid)
    if len(_seen_update_ids_order) > _MAX_SEEN_UPDATE_IDS:
        oldest = _seen_update_ids_order.pop(0)
        _seen_update_ids.discard(oldest)

# ─── تنظیمات ───
# توکن را از @Bot_Father داخل پیام‌رسان بله بگیرید
BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BALE_BOT_TOKEN_HERE")
ADMIN_IDS = list(map(int, os.getenv("ADMIN_IDS", "").split(","))) if os.getenv("ADMIN_IDS") else []
STORAGE_CHANNEL_ID = os.getenv("STORAGE_CHANNEL_ID", "")
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "@koroshrice")
DB_FILE = "database.json"

# آدرس پایه‌ی API بله (به‌جای api.telegram.org)
BALE_API_BASE_URL = "https://tapi.bale.ai/"
BALE_API_FILE_URL = "https://tapi.bale.ai/file/"

# ─── مراحل مکالمه ───
WAIT_BARNAME, WAIT_PHOTO, WAIT_CONFIRM = range(3)

# ─── ترتیب و توضیح مستندات ───
DOC_STEPS = [
    (
        "bill",
        "📄 اصل بارنامه",
        "📄 *مرحله ۱ از ۴ — اصل بارنامه*\n\n"
        "لطفاً عکس واضح از *اصل بارنامه* بگیرید و ارسال کنید.\n\n"
        "⚠️ نکات:\n"
        "• کل برگه در کادر باشد\n"
        "• نوشته‌ها خوانا باشند\n"
        "• عکس تار نباشد"
    ),
    (
        "origin",
        "🔵 حواله بار مبدأ",
        "🔵 *مرحله ۲ از ۴ — حواله بار مبدأ*\n\n"
        "لطفاً عکس *حواله بار مبدأ* (محل بارگیری) را ارسال کنید.\n\n"
        "⚠️ نکات:\n"
        "• مهر و امضا باید واضح باشد\n"
        "• تاریخ حواله مشخص باشد"
    ),
    (
        "dest",
        "🔴 رسید بار مقصد",
        "🔴 *مرحله ۳ از ۴ — رسید بار مقصد*\n\n"
        "لطفاً عکس *رسید بار مقصد* (محل تحویل) را ارسال کنید.\n\n"
        "⚠️ نکات:\n"
        "• مهر تحویل‌گیرنده باشد\n"
        "• امضای تحویل‌گیرنده واضح باشد"
    ),
    (
        "account",
        "💳 شماره حساب/شبا",
        "💳 *مرحله ۴ از ۴ — شماره حساب بانکی*\n\n"
        "لطفاً *شماره حساب بانک تجارت* یا *شماره شبا* سایر بانک‌ها را به نام "
        "*راننده اول یا دوم* که در بارنامه ذکر شده، ارسال کنید.\n\n"
        "می‌توانید شماره را به‌صورت *متن* تایپ کنید یا عکس کارت/برگه حاوی شماره را ارسال نمایید.\n\n"
        "⚠️ نکات:\n"
        "• نام صاحب حساب باید با نام راننده در بارنامه مطابقت داشته باشد\n"
        "• شماره شبا با IR شروع می‌شود"
    ),
]

DOC_KEYS = [s[0] for s in DOC_STEPS]
DOC_NAMES = {s[0]: s[1] for s in DOC_STEPS}
DOC_GUIDES = {s[0]: s[2] for s in DOC_STEPS}

# ─────────── پایگاه داده ───────────

def load_db() -> dict:
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_db(db: dict):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)

def get_barname_data(barname: str) -> dict:
    db = load_db()
    return db.get(barname, {
        "barname": barname,
        "created_at": "",
        "driver_id": 0,
        "driver_name": "",
        "documents": {},
        "message_ids": {}
    })

def save_barname_data(barname: str, data: dict):
    db = load_db()
    db[barname] = data
    save_db(db)

def make_review_id() -> str:
    """شناسه یکتا برای هر درخواست بررسی (هر بار ارسال نهایی، شناسه جدید می‌گیرد)"""
    return uuid.uuid4().hex[:10]

def find_barname_by_review_id(rid: str):
    """پیدا کردن بارنامه بر اساس شناسه بررسی فعال آن"""
    db = load_db()
    for bn, data in db.items():
        if data.get("review", {}).get("id") == rid:
            return bn, data
    return None, None

def make_barname_token(barname: str) -> str:
    """شناسه کوتاه و پایدار برای هر شماره بارنامه (برای استفاده در callback_data)"""
    return hashlib.md5(barname.encode("utf-8")).hexdigest()[:12]

def find_barname_by_token(token: str):
    db = load_db()
    for bn in db.keys():
        if make_barname_token(bn) == token:
            return bn
    return None

def review_status_label(db_data: dict) -> str:
    """برچسب وضعیت آخرین بررسی یک بارنامه"""
    review = db_data.get("review", {})
    status = review.get("status")
    if status == "approved":
        return "✅ تایید شده"
    if status == "rejected":
        return "❌ تایید نشده"
    if status == "partial":
        return "⚠️ تایید با کسری بار"
    if status == "pending":
        return "🕐 در انتظار بررسی"
    return "⏳ ثبت نشده / ناقص"

def get_driver_barnames(driver_id: int):
    """همه‌ی بارنامه‌های ثبت‌شده توسط یک راننده، از دیتابیس (نه حافظه‌ی موقت)"""
    db = load_db()
    rows = []
    for barname, data in db.items():
        if data.get("driver_id") == driver_id:
            rows.append((barname, data))
    rows.sort(key=lambda x: x[1].get("created_at", ""), reverse=True)
    return rows

# ─────────── کیبوردها ───────────

def main_menu_keyboard() -> InlineKeyboardMarkup:
    """منوی اصلی راننده"""
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🚀 شروع آپلود مستندات", callback_data="start_upload")],
        [InlineKeyboardButton("📋 راهنما", callback_data="show_help")],
    ])

def admin_keyboard() -> InlineKeyboardMarkup:
    """منوی اصلی ادمین"""
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🔍 دریافت مستندات بارنامه", callback_data="admin_get")],
        [InlineKeyboardButton("📊 لیست بارنامه‌ها", callback_data="admin_list")],
        [InlineKeyboardButton("✅ بارنامه‌های تایید شده (۵ روز اخیر)", callback_data="admin_approved_list")],
        [InlineKeyboardButton("🕐 نیازمند بررسی / تایید نشده", callback_data="admin_pending_list")],
        [InlineKeyboardButton("🗑 حذف بارنامه", callback_data="admin_delete")],
        [InlineKeyboardButton("📈 آمار کلی", callback_data="admin_stats")],
    ])

def driver_reply_keyboard() -> ReplyKeyboardMarkup:
    """منوی ثابت پایین صفحه برای راننده"""
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("🚀 آپلود مستندات"), KeyboardButton("📋 راهنما")],
            [KeyboardButton("📦 وضعیت بارنامه"), KeyboardButton("❌ لغو عملیات")],
        ],
        resize_keyboard=True,
        input_field_placeholder="از منوی پایین انتخاب کنید..."
    )

def admin_reply_keyboard() -> ReplyKeyboardMarkup:
    """منوی ثابت پایین صفحه برای ادمین"""
    return ReplyKeyboardMarkup(
        [
            [KeyboardButton("🔍 دریافت مستندات"), KeyboardButton("📊 لیست بارنامه‌ها")],
            [KeyboardButton("✅ بارنامه‌های تایید شده"), KeyboardButton("🕐 نیازمند بررسی")],
            [KeyboardButton("📈 آمار کلی"), KeyboardButton("🗑 حذف بارنامه")],
            [KeyboardButton("🏠 منوی اصلی ادمین")],
        ],
        resize_keyboard=True,
        input_field_placeholder="پنل مدیریت پی‌بار"
    )


def confirm_keyboard(has_prev: bool = False) -> InlineKeyboardMarkup:
    """کیبورد تأیید عکس — با دکمه بازگشت به مرحله قبل"""
    rows = [
        [InlineKeyboardButton("✅ تأیید — برو مرحله بعد", callback_data="confirm_photo")],
        [InlineKeyboardButton("🔄 عکس رو عوض کن", callback_data="retake_photo")],
    ]
    if has_prev:
        rows.append([InlineKeyboardButton("⬅️ برگشت به مرحله قبل", callback_data="go_prev_step")])
    rows.append([InlineKeyboardButton("🏁 اتمام و ثبت همین‌ها", callback_data="finish_early")])
    rows.append([InlineKeyboardButton("🏠 بازگشت به منوی اصلی", callback_data="back_to_main")])
    return InlineKeyboardMarkup(rows)

def step_waiting_keyboard(has_prev: bool = False) -> InlineKeyboardMarkup:
    """دکمه‌های زیر راهنمای هر مرحله"""
    rows = [
        [InlineKeyboardButton("⏭ این مرحله رو رد کن", callback_data="skip_step")],
    ]
    if has_prev:
        rows.append([InlineKeyboardButton("⬅️ برگشت به مرحله قبل", callback_data="go_prev_step")])
    rows.append([InlineKeyboardButton("🏁 اتمام و ثبت همین‌ها", callback_data="finish_early")])
    rows.append([InlineKeyboardButton("🏠 بازگشت به منوی اصلی", callback_data="back_to_main")])
    return InlineKeyboardMarkup(rows)

def skip_keyboard(has_prev: bool = False) -> InlineKeyboardMarkup:
    """کیبورد مرحله اختیاری (سایر مستندات)"""
    rows = [
        [InlineKeyboardButton("⏭ ندارم، برو بعدی", callback_data="skip_step")],
    ]
    if has_prev:
        rows.append([InlineKeyboardButton("⬅️ برگشت به مرحله قبل", callback_data="go_prev_step")])
    rows.append([InlineKeyboardButton("🏁 اتمام و ثبت همین‌ها", callback_data="finish_early")])
    rows.append([InlineKeyboardButton("🏠 بازگشت به منوی اصلی", callback_data="back_to_main")])
    return InlineKeyboardMarkup(rows)

def final_confirm_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ بله، تأیید نهایی", callback_data="final_confirm")],
        [InlineKeyboardButton("✏️ ویرایش مستندات", callback_data="edit_docs")],
        [InlineKeyboardButton("🏠 بازگشت به منوی اصلی", callback_data="back_to_main")],
    ])

def review_keyboard(rid: str, barname: str = "") -> InlineKeyboardMarkup:
    """کیبورد بررسی مستندات برای ادمین"""
    rows = [
        [InlineKeyboardButton("✅ تأیید", callback_data=f"radm_appr_{rid}")],
        [InlineKeyboardButton("❌ عدم تأیید", callback_data=f"radm_rej_{rid}")],
        [InlineKeyboardButton("⚠️ تأیید با کسری بار", callback_data=f"radm_part_{rid}")],
    ]
    if barname:
        token = make_barname_token(barname)
        rows.append([InlineKeyboardButton(f"📦 نمایش مجدد مستندات {barname}", callback_data=f"admin_open_{token}")])
    return InlineKeyboardMarkup(rows)

def driver_partial_response_keyboard(rid: str) -> InlineKeyboardMarkup:
    """کیبورد پاسخ راننده به تایید با کسری بار"""
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ مورد تایید است", callback_data=f"drv_ok_{rid}")],
        [InlineKeyboardButton("❌ مورد تایید نیست", callback_data=f"drv_no_{rid}")],
    ])

def barname_entry_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🏠 بازگشت به منوی اصلی", callback_data="back_to_main")],
    ])

def admin_back_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🔙 بازگشت به پنل ادمین", callback_data="admin_back")],
    ])

def edit_docs_keyboard(docs: dict) -> InlineKeyboardMarkup:
    """کیبورد ویرایش مستندات آپلودشده"""
    rows = []
    for key in DOC_KEYS:
        if key in docs:
            name = DOC_NAMES.get(key, key)
            rows.append([InlineKeyboardButton(f"🔄 {name}", callback_data=f"edit_doc_{key}")])
    rows.append([InlineKeyboardButton("🔙 بازگشت به خلاصه", callback_data="back_to_summary")])
    rows.append([InlineKeyboardButton("🏠 بازگشت به منوی اصلی", callback_data="back_to_main")])
    return InlineKeyboardMarkup(rows)

# ─────────── کمکی ───────────

def current_step_index(context) -> int:
    return context.user_data.get("step", 0)

def current_step_key(context) -> str:
    idx = current_step_index(context)
    if idx < len(DOC_STEPS):
        return DOC_STEPS[idx][0]
    return None

def progress_bar(current: int, total: int) -> str:
    filled = '▓' * current
    empty = '░' * (total - current)
    return f"{filled}{empty} ({current}/{total})"

async def send_step_guide(target, context, is_callback=False):
    """ارسال راهنمای مرحله فعلی با دکمه بازگشت"""
    idx = current_step_index(context)
    if idx >= len(DOC_STEPS):
        return
    key, name, guide = DOC_STEPS[idx]
    barname = context.user_data.get("barname", "")
    has_prev = idx > 0

    if key == "other":
        markup = skip_keyboard(has_prev=has_prev)
    else:
        markup = step_waiting_keyboard(has_prev=has_prev)

    text = (
        f"📦 بارنامه: *{barname}*\n"
        f"📊 {progress_bar(idx, len(DOC_STEPS))}\n\n"
        f"{guide}"
    )

    if is_callback:
        await target.edit_message_text(text, parse_mode="Markdown", reply_markup=markup)
    else:
        await target.message.reply_text(text, parse_mode="Markdown", reply_markup=markup)

async def go_to_main_menu(query, context, is_admin=False):
    """بازگشت به منوی اصلی"""
    context.user_data.clear()
    user_name = query.from_user.first_name

    if is_admin:
        await query.edit_message_text(
            f"🔐 *پنل مدیریت پی‌بار*\n\nسلام {user_name} عزیز! از منوی زیر انتخاب کنید:",
            parse_mode="Markdown",
            reply_markup=admin_keyboard()
        )
    else:
        await query.edit_message_text(
            f"🏠 *منوی اصلی پی‌بار*\n\nسلام {user_name} عزیز!\nبرای شروع آپلود مستندات دکمه زیر را بزنید:",
            parse_mode="Markdown",
            reply_markup=main_menu_keyboard()
        )

# ─────────── هندلرهای اصلی ───────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user = update.effective_user
    context.user_data.clear()

    if user.id in ADMIN_IDS:
        await update.message.reply_text(
            f"👋 سلام {user.first_name} عزیز!\n\n"
            "🔐 *پنل مدیریت پی‌بار*\n"
            "از منوی پایین صفحه انتخاب کنید:",
            parse_mode="Markdown",
            reply_markup=admin_reply_keyboard()
        )
        await update.message.reply_text(
            "یا از دکمه‌های زیر استفاده کنید:",
            reply_markup=admin_keyboard()
        )
        return ConversationHandler.END

    await update.message.reply_text(
        f"👋 سلام {user.first_name} عزیز!\n\n"
        "🚛 به ربات مستندات *پی‌بار* خوش آمدید.\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "📋 *نحوه کار:*\n"
        "۱. دکمه شروع را بزنید\n"
        "۲. شماره بارنامه را وارد کنید\n"
        "۳. ربات یک‌به‌یک هر مستند را درخواست می‌کند\n"
        "۴. عکس بگیرید و ارسال کنید\n"
        "━━━━━━━━━━━━━━━━━━━━",
        parse_mode="Markdown",
        reply_markup=driver_reply_keyboard()
    )
    await update.message.reply_text(
        "یا از دکمه‌های زیر شروع کنید:",
        reply_markup=main_menu_keyboard()
    )
    return ConversationHandler.END


async def handle_start_upload_button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """دکمه‌ی «🚀 آپلود مستندات» از منوی پایین صفحه — باید entry_point مکالمه باشد تا وضعیت درست ثبت شود"""
    await update.message.reply_text(
        "📝 لطفاً *شماره بارنامه* را وارد کنید:",
        parse_mode="Markdown",
        reply_markup=barname_entry_keyboard()
    )
    return WAIT_BARNAME


async def handle_start_upload(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """شروع فرایند آپلود از منوی اصلی"""
    query = update.callback_query
    await query.answer()

    await query.edit_message_text(
        "📝 لطفاً *شماره بارنامه* را وارد کنید:\n\n"
        "_(حداقل ۳ کاراکتر)_",
        parse_mode="Markdown",
        reply_markup=barname_entry_keyboard()
    )
    return WAIT_BARNAME


async def resubmit_barname(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """بارگذاری مجدد مستندات یک بارنامه رد شده، از طریق دکمه پیام عدم تایید"""
    query = update.callback_query
    await query.answer()

    token = query.data.replace("resubmit_", "")
    barname = find_barname_by_token(token)
    if not barname:
        await query.message.reply_text("⚠️ بارنامه یافت نشد.")
        return ConversationHandler.END

    context.user_data.clear()
    context.user_data["barname"] = barname
    context.user_data["step"] = 0
    context.user_data["session_documents"] = {}

    await query.message.reply_text(
        f"🔄 بارگذاری مجدد مستندات بارنامه *{barname}*\n\n"
        "مستندات را دوباره یک‌به‌یک ارسال کنید 👇",
        parse_mode="Markdown"
    )
    await send_step_guide(query, context, is_callback=False)
    return WAIT_PHOTO


async def handle_show_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """نمایش راهنما"""
    query = update.callback_query
    await query.answer()

    await query.edit_message_text(
        "📋 *راهنمای استفاده از ربات پی‌بار*\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "🔹 *مستندات مورد نیاز:*\n"
        "۱. 📄 اصل بارنامه\n"
        "۲. 🔵 حواله بار مبدأ\n"
        "۳. 🔴 رسید بار مقصد\n"
        "۴. 💳 شماره حساب بانک تجارت یا شماره شبا (به نام راننده اول یا دوم بارنامه)\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "💡 *نکات مهم:*\n"
        "• عکس‌ها باید واضح و خوانا باشند\n"
        "• در هر مرحله می‌توانید برگردید\n"
        "• در صورت خطا دکمه بازگشت را بزنید\n"
        "• /cancel برای لغو کامل عملیات",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🚀 شروع آپلود", callback_data="start_upload")],
            [InlineKeyboardButton("🏠 منوی اصلی", callback_data="back_to_main")],
        ])
    )


async def back_to_main(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """بازگشت به منوی اصلی — برای راننده و ادمین"""
    query = update.callback_query
    await query.answer()
    user = query.from_user
    is_admin = user.id in ADMIN_IDS
    await go_to_main_menu(query, context, is_admin=is_admin)
    return ConversationHandler.END


async def receive_barname(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    barname = update.message.text.strip()

    if not barname or len(barname) < 3:
        await update.message.reply_text(
            "⚠️ شماره بارنامه معتبر نیست.\nلطفاً دوباره وارد کنید:",
            reply_markup=barname_entry_keyboard()
        )
        return WAIT_BARNAME

    context.user_data["barname"] = barname
    context.user_data["step"] = 0
    context.user_data["session_documents"] = {}

    await update.message.reply_text(
        f"✅ بارنامه *{barname}* ثبت شد.\n\n"
        "الان شروع می‌کنیم — مستندات را یک‌به‌یک آپلود کنید 👇",
        parse_mode="Markdown"
    )

    await send_step_guide(update, context, is_callback=False)
    return WAIT_PHOTO


async def receive_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """دریافت عکس یا متن (برای مرحله شماره حساب/شبا) از راننده"""
    idx = current_step_index(context)
    key, name, _ = DOC_STEPS[idx]
    has_prev = idx > 0

    photo = update.message.photo[-1] if update.message.photo else None
    doc = update.message.document
    text = update.message.text

    # مرحله شماره حساب/شبا می‌تواند به‌صورت متن هم ارسال شود
    if key == "account" and text and not photo and not doc:
        context.user_data["pending_file"] = {"file_id": None, "type": "text", "text": text.strip()}
        await update.message.reply_text(
            f"👆 اطلاعات *{name}* دریافت شد:\n\n`{text.strip()}`\n\n"
            "آیا این اطلاعات صحیح است؟",
            parse_mode="Markdown",
            reply_markup=confirm_keyboard(has_prev=has_prev)
        )
        return WAIT_CONFIRM

    if not photo and not doc:
        warn = (
            "⚠️ لطفاً شماره حساب/شبا را به‌صورت متن تایپ کنید یا عکس آن را ارسال نمایید."
            if key == "account" else "⚠️ لطفاً عکس ارسال کنید."
        )
        await update.message.reply_text(
            warn,
            reply_markup=step_waiting_keyboard(has_prev=has_prev)
        )
        return WAIT_PHOTO

    if photo:
        file_id = photo.file_id
        file_type = "photo"
    else:
        file_id = doc.file_id
        file_type = "document"

    context.user_data["pending_file"] = {"file_id": file_id, "type": file_type}

    await update.message.reply_text(
        f"👆 عکس *{name}* دریافت شد!\n\n"
        "آیا این عکس واضح و کامل است؟",
        parse_mode="Markdown",
        reply_markup=confirm_keyboard(has_prev=has_prev)
    )
    return WAIT_CONFIRM


async def send_single_doc(bot, chat_id, doc_key: str, doc_info: dict, barname: str, extra_caption: str = "", parse_mode: str = None):
    """ارسال یک مستند (عکس/فایل/متن) به یک چت مشخص"""
    label = DOC_NAMES.get(doc_key, doc_key)
    caption = f"{label}\n📦 بارنامه: {barname}"
    if extra_caption:
        caption += f"\n{extra_caption}"

    ftype = doc_info.get("file_type")
    if ftype == "photo":
        await bot.send_photo(chat_id=chat_id, photo=doc_info["file_id"], caption=caption, parse_mode=parse_mode)
    elif ftype == "document":
        await bot.send_document(chat_id=chat_id, document=doc_info["file_id"], caption=caption, parse_mode=parse_mode)
    else:
        await bot.send_message(chat_id=chat_id, text=f"{caption}\n\n📝 {doc_info.get('text', '-')}", parse_mode=parse_mode)


async def confirm_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """تأیید عکس/متن و رفتن به مرحله بعد"""
    query = update.callback_query
    await query.answer()

    pending = context.user_data.get("pending_file")
    if not pending:
        await query.edit_message_text(
            "⚠️ خطا — اطلاعاتی برای ذخیره یافت نشد.\nلطفاً دوباره ارسال کنید.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔄 تلاش مجدد", callback_data="retake_photo")],
                [InlineKeyboardButton("🏠 بازگشت به منو", callback_data="back_to_main")],
            ])
        )
        return WAIT_PHOTO

    barname = context.user_data["barname"]
    idx = current_step_index(context)
    key, name, _ = DOC_STEPS[idx]

    session_docs = context.user_data.setdefault("session_documents", {})
    session_docs[key] = {
        "file_id": pending["file_id"],
        "file_type": pending["type"],
        "text": pending.get("text", ""),
        "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "driver_id": update.effective_user.id,
    }

    context.user_data["pending_file"] = None

    next_idx = idx + 1
    context.user_data["step"] = next_idx

    if next_idx >= len(DOC_STEPS):
        return await show_final_summary(query, context)

    next_key, next_name, next_guide = DOC_STEPS[next_idx]

    await query.edit_message_text(
        f"✅ *{name}* با موفقیت ذخیره شد!\n\n"
        f"📊 پیشرفت: {progress_bar(next_idx, len(DOC_STEPS))}\n\n"
        f"حالا نوبت *{next_name}* است 👇",
        parse_mode="Markdown"
    )

    has_prev = next_idx > 0
    markup = step_waiting_keyboard(has_prev=has_prev)

    barname = context.user_data.get("barname", "")
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=f"📦 بارنامه: *{barname}*\n\n{next_guide}",
        parse_mode="Markdown",
        reply_markup=markup
    )

    return WAIT_PHOTO


async def retake_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """عکس رو عوض کن"""
    query = update.callback_query
    await query.answer()

    idx = current_step_index(context)
    key, name, _ = DOC_STEPS[idx]
    context.user_data["pending_file"] = None
    has_prev = idx > 0

    await query.edit_message_text(
        f"🔄 باشه! دوباره عکس *{name}* را ارسال کنید:\n\n"
        "_(عکس جدید را در چت ارسال کنید)_",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("⬅️ برگشت به مرحله قبل", callback_data="go_prev_step")] if has_prev else [],
            [InlineKeyboardButton("🏠 بازگشت به منوی اصلی", callback_data="back_to_main")],
        ])
    )
    return WAIT_PHOTO


async def go_prev_step(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """بازگشت به مرحله قبل"""
    query = update.callback_query
    await query.answer()

    idx = current_step_index(context)

    if idx <= 0:
        # اگر در اولین مرحله‌ایم، برگشت به ورود بارنامه
        await query.edit_message_text(
            "📝 لطفاً *شماره بارنامه* را دوباره وارد کنید\n"
            "یا اگر می‌خواهید بارنامه فعلی را ادامه دهید شماره همان را وارد کنید:",
            parse_mode="Markdown",
            reply_markup=barname_entry_keyboard()
        )
        context.user_data["step"] = 0
        return WAIT_BARNAME

    prev_idx = idx - 1
    context.user_data["step"] = prev_idx
    context.user_data["pending_file"] = None

    # حذف مستند قبلی اگه ذخیره شده (فقط از حافظه‌ی موقت جلسه، نه دیتابیس)
    session_docs = context.user_data.get("session_documents", {})
    prev_key = DOC_STEPS[prev_idx][0]
    if prev_key in session_docs:
        del session_docs[prev_key]

    await send_step_guide(query, context, is_callback=True)
    return WAIT_PHOTO


async def skip_step(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """رد کردن مرحله"""
    query = update.callback_query
    await query.answer()

    idx = current_step_index(context)
    key, name, _ = DOC_STEPS[idx]
    next_idx = idx + 1
    context.user_data["step"] = next_idx

    if next_idx >= len(DOC_STEPS):
        return await show_final_summary(query, context)

    await query.edit_message_text(
        f"⏭ *{name}* رد شد.\n\n"
        f"📊 پیشرفت: {progress_bar(next_idx, len(DOC_STEPS))}",
        parse_mode="Markdown"
    )

    context.user_data["step"] = next_idx
    next_key, next_name, next_guide = DOC_STEPS[next_idx]
    barname = context.user_data.get("barname", "")
    has_prev = next_idx > 0

    if next_key == "other":
        markup = skip_keyboard(has_prev=has_prev)
    else:
        markup = step_waiting_keyboard(has_prev=has_prev)

    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=f"📦 بارنامه: *{barname}*\n\n{next_guide}",
        parse_mode="Markdown",
        reply_markup=markup
    )
    return WAIT_PHOTO


async def finish_early(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """اتمام زودهنگام"""
    query = update.callback_query
    await query.answer()

    barname = context.user_data.get("barname", "")
    docs = context.user_data.get("session_documents", {})

    if not docs:
        await query.edit_message_text(
            "⚠️ *هنوز هیچ مستندی آپلود نشده!*\n\n"
            "لطفاً حداقل یک مستند ارسال کنید تا بتوانید ثبت کنید.",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔙 برگشت به مرحله جاری", callback_data="go_back_to_step")],
                [InlineKeyboardButton("🏠 بازگشت به منوی اصلی", callback_data="back_to_main")],
            ])
        )
        return WAIT_PHOTO

    return await show_final_summary(query, context)


async def go_back_to_step(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """بازگشت به مرحله جاری بدون تغییر"""
    query = update.callback_query
    await query.answer()
    await send_step_guide(query, context, is_callback=True)
    return WAIT_PHOTO


async def show_final_summary(query, context) -> int:
    """نمایش خلاصه نهایی"""
    barname = context.user_data.get("barname", "")
    docs = context.user_data.get("session_documents", {})

    done_list = "\n".join([f"  ✅ {DOC_NAMES.get(k, k)}" for k in docs.keys()]) or "  (هیچ‌کدام)"
    missing = [k for k in DOC_KEYS if k not in docs]
    missing_list = "\n".join([f"  ⬜ {DOC_NAMES.get(k, k)}" for k in missing]) if missing else ""

    summary = (
        f"🎯 *خلاصه بارنامه {barname}*\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        f"✅ *آپلود شده ({len(docs)}/{len(DOC_KEYS)}):*\n{done_list}\n"
    )
    if missing_list:
        summary += f"\n⬜ *آپلود نشده:*\n{missing_list}\n"
    summary += "\n━━━━━━━━━━━━━━━━━━━━\n\nآیا همه چیز درست است؟"

    await query.edit_message_text(
        summary,
        parse_mode="Markdown",
        reply_markup=final_confirm_keyboard()
    )
    return WAIT_CONFIRM


async def edit_docs(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """نمایش منوی ویرایش مستندات"""
    query = update.callback_query
    await query.answer()

    barname = context.user_data.get("barname", "")
    docs = context.user_data.get("session_documents", {})

    if not docs:
        await query.edit_message_text(
            "⚠️ هیچ مستندی برای ویرایش وجود ندارد.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔙 بازگشت", callback_data="back_to_summary")],
            ])
        )
        return WAIT_CONFIRM

    await query.edit_message_text(
        f"✏️ *ویرایش مستندات بارنامه {barname}*\n\n"
        "کدام مستند را می‌خواهید دوباره آپلود کنید؟",
        parse_mode="Markdown",
        reply_markup=edit_docs_keyboard(docs)
    )
    return WAIT_CONFIRM


async def edit_specific_doc(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """شروع دوباره یک مستند خاص"""
    query = update.callback_query
    await query.answer()

    doc_key = query.data.replace("edit_doc_", "")
    idx = DOC_KEYS.index(doc_key) if doc_key in DOC_KEYS else 0

    context.user_data["step"] = idx
    context.user_data["pending_file"] = None

    barname = context.user_data.get("barname", "")
    session_docs = context.user_data.get("session_documents", {})
    if doc_key in session_docs:
        del session_docs[doc_key]

    name = DOC_NAMES.get(doc_key, doc_key)
    guide = DOC_GUIDES.get(doc_key, "")
    has_prev = idx > 0

    if doc_key == "other":
        markup = skip_keyboard(has_prev=has_prev)
    else:
        markup = step_waiting_keyboard(has_prev=has_prev)

    await query.edit_message_text(
        f"🔄 *ویرایش {name}*\n\n"
        f"📦 بارنامه: *{barname}*\n\n"
        f"{guide}",
        parse_mode="Markdown",
        reply_markup=markup
    )
    return WAIT_PHOTO


async def back_to_summary(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """بازگشت به صفحه خلاصه"""
    query = update.callback_query
    await query.answer()
    return await show_final_summary(query, context)


async def dispatch_review_package(context, barname: str, db_data: dict) -> str:
    """ارسال پکیج مستندات و دکمه‌های بررسی برای همه‌ی ادمین‌ها؛ یک شناسه بررسی جدید تولید می‌کند"""
    rid = make_review_id()
    db_data["review"] = {
        "id": rid,
        "status": "pending",
        "reason": "",
        "deduction_note": "",
        "reviewed_by": None,
        "reviewed_at": None,
        "admin_messages": {},
    }
    save_barname_data(barname, db_data)

    docs = db_data.get("documents", {})
    driver_id = db_data.get("driver_id")
    driver_name = db_data.get("driver_name", "-")
    driver_link = f"{driver_name} (آیدی: `{driver_id}`)" if driver_id else driver_name

    if ADMIN_IDS:
        for admin_id in ADMIN_IDS:
            try:
                # ارسال مستندات به صورت پکیج برای ادمین
                for doc_key, doc_info in docs.items():
                    await send_single_doc(
                        context.bot, admin_id, doc_key, doc_info, barname,
                        extra_caption=f"👤 فرستنده: {driver_link}",
                        parse_mode="Markdown"
                    )

                package_text = (
                    f"📦 مستندات بارنامه شماره `{barname}` توسط آیدی {driver_link} ارسال شد\n\n"
                    f"📄 تعداد مستندات: {len(docs)}\n"
                    f"🕐 زمان: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n"
                    "برای مشاهده‌ی مجدد مستندات این بارنامه، دکمه‌ی «📦 نمایش مجدد مستندات» زیر همین پیام را بزنید.\n\n"
                    "لطفاً بررسی و اقدام کنید:"
                )
                sent = await context.bot.send_message(
                    chat_id=admin_id,
                    text=package_text,
                    parse_mode="Markdown",
                    reply_markup=review_keyboard(rid, barname)
                )
                db_data = get_barname_data(barname)
                db_data["review"]["admin_messages"][str(admin_id)] = sent.message_id
                save_barname_data(barname, db_data)
            except Exception as e:
                logger.error(f"خطا در ارسال پکیج به ادمین {admin_id}: {e}")

    return rid


async def final_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """تایید نهایی — تنها از همین نقطه به بعد، مستندات در دیتابیس ذخیره و برای ادمین ارسال می‌شوند"""
    query = update.callback_query
    await query.answer()

    barname = context.user_data.get("barname", "")
    session_docs = context.user_data.get("session_documents", {})

    if not barname or not session_docs:
        await query.edit_message_text(
            "⚠️ هیچ مستندی برای ثبت وجود ندارد.\nلطفاً از نو شروع کنید.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🆕 شروع مجدد", callback_data="start_upload")],
                [InlineKeyboardButton("🏠 منوی اصلی", callback_data="back_to_main")],
            ])
        )
        context.user_data.clear()
        return ConversationHandler.END

    # از همین‌جا به بعد، ثبت رسمی در دیتابیس انجام می‌شود
    db_data = get_barname_data(barname)
    if not db_data.get("created_at"):
        db_data["created_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    db_data["driver_id"] = update.effective_user.id
    db_data["driver_name"] = update.effective_user.full_name
    db_data["documents"] = session_docs
    db_data["completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    db_data["status"] = "completed"
    save_barname_data(barname, db_data)

    # فوروارد به کانال آرشیو (در صورت تنظیم) — فقط بعد از تایید نهایی
    if STORAGE_CHANNEL_ID and str(STORAGE_CHANNEL_ID) not in [str(a) for a in ADMIN_IDS]:
        for doc_key, doc_info in session_docs.items():
            try:
                await send_single_doc(
                    context.bot, STORAGE_CHANNEL_ID, doc_key, doc_info, barname,
                    extra_caption=f"👤 راننده: {update.effective_user.full_name}"
                )
            except Exception as e:
                logger.error(f"خطا در فوروارد به کانال آرشیو: {e}")

    await dispatch_review_package(context, barname, db_data)

    await query.edit_message_text(
        f"✅ *مستندات ارسال شد!*\n\n"
        f"📦 بارنامه: {barname}\n\n"
        f"🚛 *راننده گرامی،*\n"
        f"مستندات شما برای بررسی به ادمین ارسال شد. نتیجه بررسی به‌زودی از طریق همین ربات به شما اطلاع داده خواهد شد.\n\n"
        f"برای بارنامه جدید دکمه زیر را بزنید:",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🆕 بارنامه جدید", callback_data="start_upload")],
            [InlineKeyboardButton("🏠 منوی اصلی", callback_data="back_to_main")],
        ])
    )
    context.user_data.clear()
    return ConversationHandler.END


# ─────────── بررسی مستندات توسط ادمین ───────────

async def _finalize_review(context, barname: str, db_data: dict, status_label: str):
    """حذف دکمه‌های بررسی از پیام تمام ادمین‌ها و اطلاع‌رسانی تصمیم نهایی به آن‌ها"""
    admin_messages = db_data.get("review", {}).get("admin_messages", {})
    for admin_id_str, msg_id in admin_messages.items():
        try:
            await context.bot.edit_message_reply_markup(
                chat_id=int(admin_id_str),
                message_id=msg_id,
                reply_markup=None
            )
        except Exception as e:
            logger.error(f"خطا در حذف دکمه‌های ادمین {admin_id_str}: {e}")
        try:
            await context.bot.send_message(
                chat_id=int(admin_id_str),
                text=f"📦 بارنامه {barname}: {status_label}"
            )
        except Exception as e:
            logger.error(f"خطا در اطلاع‌رسانی به ادمین {admin_id_str}: {e}")


async def review_approve(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """تایید مستندات توسط ادمین"""
    query = update.callback_query
    user = query.from_user
    if user.id not in ADMIN_IDS:
        logger.warning(f"⛔️ تلاش دسترسی غیرمجاز به دکمه ادمین توسط آیدی {user.id} (لیست ادمین‌های مجاز: {ADMIN_IDS})")
        await query.answer("⛔️ دسترسی ندارید.", show_alert=True)
        return

    rid = query.data.replace("radm_appr_", "")
    barname, db_data = find_barname_by_review_id(rid)
    if not barname or db_data.get("review", {}).get("status") != "pending":
        logger.warning(f"⚠️ درخواست بررسی نامعتبر — rid={rid} یافت‌نشد یا قبلاً بررسی شده (دیتابیس ممکن است ری‌ست شده باشد)")
        await query.answer("⚠️ این درخواست دیگر معتبر نیست یا قبلاً بررسی شده.", show_alert=True)
        return

    await query.answer("✅ تأیید شد")

    db_data["review"]["status"] = "approved"
    db_data["review"]["reviewed_by"] = user.id
    db_data["review"]["reviewed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    save_barname_data(barname, db_data)

    driver_id = db_data.get("driver_id")
    if driver_id:
        try:
            await context.bot.send_message(
                chat_id=driver_id,
                text=(
                    f"✅ راننده محترم مستندات ارسالی بارنامه شماره {barname} شما تایید شد "
                    "و ظرف ۲ روز کاری هزینه بارنامه به شماره حساب اعلامی شما واریز خواهد شد."
                ),
                reply_markup=driver_reply_keyboard()
            )
        except Exception as e:
            logger.error(f"خطا در اطلاع‌رسانی به راننده: {e}")

    await _finalize_review(context, barname, db_data, f"✅ تأیید شد توسط {user.first_name}")


async def review_reject_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """شروع فرایند ثبت علت عدم تایید"""
    query = update.callback_query
    user = query.from_user
    if user.id not in ADMIN_IDS:
        logger.warning(f"⛔️ تلاش دسترسی غیرمجاز به دکمه ادمین توسط آیدی {user.id} (لیست ادمین‌های مجاز: {ADMIN_IDS})")
        await query.answer("⛔️ دسترسی ندارید.", show_alert=True)
        return

    rid = query.data.replace("radm_rej_", "")
    barname, db_data = find_barname_by_review_id(rid)
    if not barname or db_data.get("review", {}).get("status") != "pending":
        logger.warning(f"⚠️ درخواست بررسی نامعتبر — rid={rid} یافت‌نشد یا قبلاً بررسی شده (دیتابیس ممکن است ری‌ست شده باشد)")
        await query.answer("⚠️ این درخواست دیگر معتبر نیست یا قبلاً بررسی شده.", show_alert=True)
        return

    await query.answer()
    context.user_data["review_pending"] = {"action": "reject", "rid": rid, "barname": barname}

    await context.bot.send_message(
        chat_id=user.id,
        text=f"❌ لطفاً *علت عدم تأیید* مستندات بارنامه {barname} را بنویسید:",
        parse_mode="Markdown"
    )


async def review_partial_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """شروع فرایند ثبت توضیح تایید با کسری بار"""
    query = update.callback_query
    user = query.from_user
    if user.id not in ADMIN_IDS:
        logger.warning(f"⛔️ تلاش دسترسی غیرمجاز به دکمه ادمین توسط آیدی {user.id} (لیست ادمین‌های مجاز: {ADMIN_IDS})")
        await query.answer("⛔️ دسترسی ندارید.", show_alert=True)
        return

    rid = query.data.replace("radm_part_", "")
    barname, db_data = find_barname_by_review_id(rid)
    if not barname or db_data.get("review", {}).get("status") != "pending":
        logger.warning(f"⚠️ درخواست بررسی نامعتبر — rid={rid} یافت‌نشد یا قبلاً بررسی شده (دیتابیس ممکن است ری‌ست شده باشد)")
        await query.answer("⚠️ این درخواست دیگر معتبر نیست یا قبلاً بررسی شده.", show_alert=True)
        return

    await query.answer()
    context.user_data["review_pending"] = {"action": "partial", "rid": rid, "barname": barname}

    await context.bot.send_message(
        chat_id=user.id,
        text=(
            f"⚠️ لطفاً *مقدار کسری بار* مربوط به بارنامه {barname} را بنویسید:\n\n"
            "_(مثال: ۲ تن، یا ۳ کیسه)_"
        ),
        parse_mode="Markdown"
    )


# ─────────── پاسخ راننده به تایید با کسری بار ───────────

async def driver_partial_accept(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """راننده کسری بار را قبول می‌کند"""
    query = update.callback_query
    rid = query.data.replace("drv_ok_", "")
    barname, db_data = find_barname_by_review_id(rid)

    if not barname or db_data.get("review", {}).get("status") != "partial" or db_data["review"].get("driver_response"):
        await query.answer("⚠️ این درخواست دیگر معتبر نیست یا قبلاً پاسخ داده شده.", show_alert=True)
        return

    await query.answer("✅ ثبت شد")
    await query.edit_message_reply_markup(reply_markup=None)

    db_data["review"]["driver_response"] = "accepted"
    db_data["review"]["status"] = "approved"
    save_barname_data(barname, db_data)

    await query.message.reply_text(
        "✅ تایید شما ثبت شد. با تشکر از همکاری شما.",
        reply_markup=driver_reply_keyboard()
    )

    for admin_id in ADMIN_IDS:
        try:
            await context.bot.send_message(
                chat_id=admin_id,
                text=(
                    f"✅ بارنامه شماره {barname} با کسری بار مورد تایید راننده قرار گرفت.\n"
                    "📌 وضعیت بارنامه به «تایید شده» تغییر یافت."
                )
            )
        except Exception as e:
            logger.error(f"خطا در اطلاع‌رسانی به ادمین {admin_id}: {e}")


async def driver_partial_reject_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """راننده کسری بار را قبول ندارد — شروع دریافت علت"""
    query = update.callback_query
    rid = query.data.replace("drv_no_", "")
    barname, db_data = find_barname_by_review_id(rid)

    if not barname or db_data.get("review", {}).get("status") != "partial" or db_data["review"].get("driver_response"):
        await query.answer("⚠️ این درخواست دیگر معتبر نیست یا قبلاً پاسخ داده شده.", show_alert=True)
        return

    await query.answer()
    await query.edit_message_reply_markup(reply_markup=None)

    context.user_data["driver_dispute_pending"] = {"rid": rid, "barname": barname}

    await query.message.reply_text("❌ لطفاً *علت عدم تایید* کسری بار را بنویسید:", parse_mode="Markdown")


# ─────────── پنل ادمین ───────────

async def admin_back(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """بازگشت به پنل ادمین"""
    query = update.callback_query
    await query.answer()
    context.user_data.pop("admin_action", None)

    await query.edit_message_text(
        "🔐 *پنل مدیریت پی‌بار*\n\nاز منوی زیر انتخاب کنید:",
        parse_mode="Markdown",
        reply_markup=admin_keyboard()
    )


async def admin_get_docs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["admin_action"] = "get_docs"

    await query.edit_message_text(
        "🔍 *دریافت مستندات*\n\nشماره بارنامه مورد نظر را وارد کنید:",
        parse_mode="Markdown",
        reply_markup=admin_back_keyboard()
    )


async def admin_delete_barname(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["admin_action"] = "delete_barname"

    await query.edit_message_text(
        "🗑 *حذف بارنامه*\n\nشماره بارنامه‌ای که می‌خواهید حذف کنید را وارد کنید:",
        parse_mode="Markdown",
        reply_markup=admin_back_keyboard()
    )


async def admin_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    db = load_db()
    total = len(db)
    completed = sum(1 for d in db.values() if d.get("status") == "completed")
    pending = total - completed
    total_docs = sum(len(d.get("documents", {})) for d in db.values())

    await query.edit_message_text(
        f"📈 *آمار کلی پی‌بار*\n\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"📦 کل بارنامه‌ها: *{total}*\n"
        f"✅ تکمیل‌شده: *{completed}*\n"
        f"⏳ در انتظار: *{pending}*\n"
        f"📄 کل مستندات: *{total_docs}*\n"
        f"━━━━━━━━━━━━━━━━━━━━",
        parse_mode="Markdown",
        reply_markup=admin_back_keyboard()
    )


async def admin_list_barnames(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    db = load_db()
    if not db:
        await query.edit_message_text(
            "📭 هیچ بارنامه‌ای ثبت نشده.",
            reply_markup=admin_keyboard()
        )
        return

    lines = []
    for barname, data in sorted(db.items(), key=lambda x: x[1].get("created_at", ""), reverse=True):
        status = review_status_label(data)
        doc_count = len(data.get("documents", {}))
        created = data.get("created_at", "-")
        lines.append(f"{status} | `{barname}` | {doc_count} مستند | {created}")

    text = "📊 *لیست بارنامه‌ها:*\n\n" + "\n".join(lines[:20])
    if len(db) > 20:
        text += f"\n\n... و {len(db)-20} بارنامه دیگر"

    await query.edit_message_text(
        text,
        parse_mode="Markdown",
        reply_markup=admin_back_keyboard()
    )


def build_approved_list_content():
    """متن و کیبورد لیست بارنامه‌های تایید شده در ۵ روز اخیر"""
    db = load_db()
    cutoff = datetime.now() - timedelta(days=5)
    rows = []
    for barname, data in db.items():
        review = data.get("review", {})
        if review.get("status") != "approved":
            continue
        reviewed_at = review.get("reviewed_at", "")
        try:
            reviewed_dt = datetime.strptime(reviewed_at, "%Y-%m-%d %H:%M")
        except (ValueError, TypeError):
            continue
        if reviewed_dt >= cutoff:
            rows.append((barname, data, reviewed_dt))

    if not rows:
        return "📭 در ۵ روز گذشته بارنامه تایید شده‌ای ثبت نشده.", admin_back_keyboard()

    rows.sort(key=lambda x: x[2], reverse=True)
    lines = []
    for barname, data, reviewed_dt in rows:
        lines.append(
            f"✅ `{barname}` | راننده: {data.get('driver_name', '-')} | "
            f"{reviewed_dt.strftime('%Y-%m-%d %H:%M')}"
        )

    text = "✅ *بارنامه‌های تایید شده (۵ روز اخیر):*\n\n" + "\n".join(lines[:30])
    if len(rows) > 30:
        text += f"\n\n... و {len(rows)-30} بارنامه دیگر"

    return text, admin_back_keyboard()


PENDING_PAGE_SIZE = 8

def build_pending_list_content(page: int = 0):
    """متن و کیبورد لیست بارنامه‌های نیازمند بررسی / تایید نشده (صفحه‌بندی‌شده تا در یک پیام جا بشه)"""
    db = load_db()
    rows = []
    for barname, data in db.items():
        review = data.get("review", {})
        status = review.get("status")
        if status == "approved":
            continue
        if not data.get("documents"):
            continue
        rows.append((barname, data))

    if not rows:
        return "📭 در حال حاضر بارنامه‌ای در انتظار بررسی یا تایید‌نشده وجود ندارد.", admin_back_keyboard()

    rows.sort(key=lambda x: x[1].get("created_at", ""), reverse=True)

    total_pages = max(1, (len(rows) + PENDING_PAGE_SIZE - 1) // PENDING_PAGE_SIZE)
    page = max(0, min(page, total_pages - 1))
    start = page * PENDING_PAGE_SIZE
    page_rows = rows[start:start + PENDING_PAGE_SIZE]

    buttons = []
    for barname, data in page_rows:
        label = f"{review_status_label(data)} — {barname}"
        token = make_barname_token(barname)
        buttons.append([InlineKeyboardButton(label, callback_data=f"admin_open_{token}")])

    nav_row = []
    if page > 0:
        nav_row.append(InlineKeyboardButton("◀️ قبلی", callback_data=f"admin_pending_page_{page-1}"))
    if page < total_pages - 1:
        nav_row.append(InlineKeyboardButton("بعدی ▶️", callback_data=f"admin_pending_page_{page+1}"))
    if nav_row:
        buttons.append(nav_row)

    buttons.append([InlineKeyboardButton("🔙 بازگشت به پنل ادمین", callback_data="admin_back")])

    text = (
        "🕐 *بارنامه‌های نیازمند بررسی / تایید نشده:*\n\n"
        "برای مشاهده‌ی مجدد مستندات و اقدام روی هرکدام، آن را انتخاب کنید 👇\n\n"
        f"_(صفحه {page+1} از {total_pages} — مجموعاً {len(rows)} بارنامه)_"
    )

    return text, InlineKeyboardMarkup(buttons)


async def admin_approved_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """لیست بارنامه‌های تایید شده تا ۵ روز گذشته"""
    query = update.callback_query
    await query.answer()
    text, markup = build_approved_list_content()
    await query.edit_message_text(text, parse_mode="Markdown", reply_markup=markup)


async def admin_pending_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """لیست بارنامه‌هایی که واکنش نشان داده نشده یا تایید نشده‌اند - قابل انتخاب برای بررسی مجدد"""
    query = update.callback_query
    await query.answer()
    text, markup = build_pending_list_content(page=0)
    await query.edit_message_text(text, parse_mode="Markdown", reply_markup=markup)


async def admin_pending_page(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """جابه‌جایی صفحه در لیست نیازمند بررسی (بدون ارسال پیام جدید)"""
    query = update.callback_query
    await query.answer()
    try:
        page = int(query.data.replace("admin_pending_page_", ""))
    except ValueError:
        page = 0
    text, markup = build_pending_list_content(page=page)
    await query.edit_message_text(text, parse_mode="Markdown", reply_markup=markup)


async def admin_open_barname(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """ارسال مجدد مستندات یک بارنامه خاص به همراه دکمه‌های بررسی"""
    query = update.callback_query
    user = query.from_user
    if user.id not in ADMIN_IDS:
        logger.warning(f"⛔️ تلاش دسترسی غیرمجاز به دکمه ادمین توسط آیدی {user.id} (لیست ادمین‌های مجاز: {ADMIN_IDS})")
        await query.answer("⛔️ دسترسی ندارید.", show_alert=True)
        return

    token = query.data.replace("admin_open_", "")
    barname = find_barname_by_token(token)
    if not barname:
        await query.answer("⚠️ بارنامه یافت نشد.", show_alert=True)
        return

    await query.answer("در حال ارسال مستندات...")

    db_data = get_barname_data(barname)
    if not db_data.get("documents"):
        await context.bot.send_message(chat_id=user.id, text=f"⚠️ بارنامه {barname} مستندی ندارد.")
        return

    await dispatch_review_package(context, barname, db_data)


async def admin_text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    # ─── پاسخ راننده به «مورد تایید نیست» برای کسری بار (برای همه کاربران، نه فقط ادمین) ───
    dispute = context.user_data.get("driver_dispute_pending")
    if dispute:
        rid = dispute["rid"]
        barname = dispute["barname"]
        note_text = update.message.text.strip()

        _, db_data = find_barname_by_review_id(rid)
        if not db_data or db_data.get("review", {}).get("status") != "partial" or db_data["review"].get("driver_response"):
            await update.message.reply_text(
                "⚠️ این درخواست دیگر معتبر نیست یا قبلاً پاسخ داده شده.",
                reply_markup=driver_reply_keyboard()
            )
            context.user_data.pop("driver_dispute_pending", None)
            return

        db_data["review"]["driver_response"] = "rejected"
        db_data["review"]["driver_reason"] = note_text
        save_barname_data(barname, db_data)

        driver_link = f"{user.full_name} (آیدی: `{user.id}`)"
        for admin_id in ADMIN_IDS:
            try:
                await context.bot.send_message(
                    chat_id=admin_id,
                    text=(
                        f"❌ راننده کسری بار بارنامه شماره {barname} را قبول ندارد.\n\n"
                        f"📝 علت: {note_text}\n"
                        f"👤 راننده: {driver_link}\n\n"
                        "لطفاً جهت هماهنگی، مستقیماً با راننده ارتباط بگیرید."
                    ),
                    parse_mode="Markdown"
                )
            except Exception as e:
                logger.error(f"خطا در اطلاع‌رسانی به ادمین {admin_id}: {e}")

        await update.message.reply_text(
            f"راننده محترم با توجه به عدم تایید کسری بار توسط شما، به آیدی ادمین {ADMIN_USERNAME} پیام دهید.",
            reply_markup=driver_reply_keyboard()
        )
        context.user_data.pop("driver_dispute_pending", None)
        return

    if user.id not in ADMIN_IDS:
        return

    # ─── پاسخ ادمین به عدم تایید / تایید با کسری بار ───
    review_pending = context.user_data.get("review_pending")
    if review_pending:
        rid = review_pending["rid"]
        r_action = review_pending["action"]
        note_text = update.message.text.strip()

        barname, db_data = find_barname_by_review_id(rid)
        if not barname or db_data.get("review", {}).get("status") != "pending":
            await update.message.reply_text(
                "⚠️ این درخواست دیگر معتبر نیست یا قبلاً بررسی شده.",
                reply_markup=admin_keyboard()
            )
            context.user_data.pop("review_pending", None)
            return

        db_data["review"]["reviewed_by"] = user.id
        db_data["review"]["reviewed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        driver_id = db_data.get("driver_id")

        if r_action == "reject":
            db_data["review"]["status"] = "rejected"
            db_data["review"]["reason"] = note_text
            save_barname_data(barname, db_data)

            if driver_id:
                try:
                    resubmit_token = make_barname_token(barname)
                    await context.bot.send_message(
                        chat_id=driver_id,
                        text=(
                            f"❌ راننده محترم مستندات ارسالی مربوط به بارنامه شماره {barname} "
                            f"به دلیل «{note_text}» مورد تایید قرار نگرفت، "
                            "لطفا مجددا مستندات را به صورت واضح و کامل ارسال نمایید."
                        ),
                        reply_markup=InlineKeyboardMarkup([
                            [InlineKeyboardButton("🔄 بارگذاری مجدد مستندات", callback_data=f"resubmit_{resubmit_token}")]
                        ])
                    )
                    await context.bot.send_message(
                        chat_id=driver_id,
                        text="👇 همچنین می‌توانید از منوی پایین صفحه استفاده کنید.",
                        reply_markup=driver_reply_keyboard()
                    )
                except Exception as e:
                    logger.error(f"خطا در اطلاع‌رسانی به راننده: {e}")

            await update.message.reply_text(
                f"❌ عدم تأیید بارنامه {barname} ثبت و به راننده اطلاع داده شد.",
                reply_markup=admin_keyboard()
            )
            await _finalize_review(context, barname, db_data, f"❌ عدم تأیید (علت: {note_text})")

        elif r_action == "partial":
            db_data["review"]["status"] = "partial"
            db_data["review"]["deduction_note"] = note_text
            db_data["review"]["driver_response"] = None
            save_barname_data(barname, db_data)

            if driver_id:
                try:
                    await context.bot.send_message(
                        chat_id=driver_id,
                        text=(
                            f"⚠️ راننده محترم طبق بررسی حواله و رسید ارسالی مقدار {note_text} "
                            "کسری بار دارید که هزینه آن می‌بایست از مبلغ بارنامه کسر شود.\n\n"
                            "آیا این کسری بار مورد تایید شماست؟"
                        ),
                        reply_markup=driver_partial_response_keyboard(rid)
                    )
                    await context.bot.send_message(
                        chat_id=driver_id,
                        text="👇 همچنین می‌توانید از منوی پایین صفحه استفاده کنید.",
                        reply_markup=driver_reply_keyboard()
                    )
                except Exception as e:
                    logger.error(f"خطا در اطلاع‌رسانی به راننده: {e}")

            await update.message.reply_text(
                f"⚠️ تأیید با کسری بار برای بارنامه {barname} ثبت و به راننده اطلاع داده شد. "
                "پاسخ راننده به‌زودی برای شما ارسال خواهد شد.",
                reply_markup=admin_keyboard()
            )
            await _finalize_review(context, barname, db_data, f"⚠️ تأیید با کسری بار ({note_text}) — در انتظار پاسخ راننده")

        context.user_data.pop("review_pending", None)
        return

    action = context.user_data.get("admin_action")
    if not action:
        return

    barname = update.message.text.strip()

    # ─── حذف بارنامه ───
    if action == "delete_barname":
        db = load_db()
        if barname not in db:
            await update.message.reply_text(
                f"❌ بارنامه *{barname}* یافت نشد.",
                parse_mode="Markdown",
                reply_markup=admin_keyboard()
            )
        else:
            del db[barname]
            save_db(db)
            await update.message.reply_text(
                f"🗑 بارنامه *{barname}* با موفقیت حذف شد.",
                parse_mode="Markdown",
                reply_markup=admin_keyboard()
            )
        context.user_data.pop("admin_action", None)
        return

    # ─── دریافت مستندات ───
    if action == "get_docs":
        db_data = get_barname_data(barname)

        if not db_data.get("documents"):
            await update.message.reply_text(
                f"❌ بارنامه *{barname}* یافت نشد یا مستندی ندارد.",
                parse_mode="Markdown",
                reply_markup=admin_keyboard()
            )
            context.user_data.pop("admin_action", None)
            return

        docs = db_data["documents"]
        driver_id = db_data.get("driver_id")
        driver_link = f"{db_data.get('driver_name', '-')} (آیدی: `{driver_id}`)" if driver_id else db_data.get('driver_name', '-')
        await update.message.reply_text(
            f"📦 *بارنامه {barname}*\n"
            f"👤 راننده: {driver_link}\n"
            f"📅 تاریخ: {db_data.get('created_at', '-')}\n"
            f"📄 تعداد: {len(docs)} مستند\n"
            f"📌 وضعیت: {review_status_label(db_data)}\n\n"
            "در حال ارسال...",
            parse_mode="Markdown"
        )

        for doc_key, doc_info in docs.items():
            try:
                await send_single_doc(
                    context.bot, user.id, doc_key, doc_info, barname,
                    extra_caption=f"👤 فرستنده: {driver_link}\n🕐 {doc_info.get('uploaded_at', '-')}",
                    parse_mode="Markdown"
                )
            except Exception as e:
                label = DOC_NAMES.get(doc_key, doc_key)
                await update.message.reply_text(f"⚠️ خطا در ارسال {label}: {e}")

        await update.message.reply_text(
            f"✅ همه مستندات بارنامه *{barname}* ارسال شد.",
            parse_mode="Markdown",
            reply_markup=admin_keyboard()
        )
        context.user_data.pop("admin_action", None)


async def reply_keyboard_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """هندلر دکمه‌های منوی پایین صفحه"""
    text = update.message.text
    user = update.effective_user
    is_admin = user.id in ADMIN_IDS

    # ─── دکمه‌های راننده ───
    if text == "🚀 آپلود مستندات":
        await update.message.reply_text(
            "📝 لطفاً *شماره بارنامه* را وارد کنید:",
            parse_mode="Markdown",
            reply_markup=barname_entry_keyboard()
        )
        return WAIT_BARNAME

    elif text == "📋 راهنما":
        await update.message.reply_text(
            "📋 *راهنمای استفاده از ربات پی‌بار*\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "🔹 *مستندات مورد نیاز:*\n"
            "۱. 📄 اصل بارنامه\n"
            "۲. 🔵 حواله بار مبدأ\n"
            "۳. 🔴 رسید بار مقصد\n"
            "۴. 💳 شماره حساب بانک تجارت یا شماره شبا (به نام راننده اول یا دوم بارنامه)\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "💡 *نکات مهم:*\n"
            "• عکس‌ها باید واضح و خوانا باشند\n"
            "• در هر مرحله می‌توانید برگردید\n"
            "• /cancel برای لغو کامل عملیات",
            parse_mode="Markdown",
            reply_markup=driver_reply_keyboard()
        )

    elif text == "📦 وضعیت بارنامه":
        rows = get_driver_barnames(user.id)

        if not rows:
            await update.message.reply_text(
                "📭 شما تا الان هیچ بارنامه‌ای ثبت نکرده‌اید.\n"
                "برای شروع دکمه 🚀 آپلود مستندات را بزنید.",
                reply_markup=driver_reply_keyboard()
            )
        else:
            lines = []
            for bn, data in rows[:20]:
                doc_count = len(data.get("documents", {}))
                status = review_status_label(data)
                lines.append(f"{status} | `{bn}` | {doc_count}/{len(DOC_STEPS)} مستند | {data.get('created_at', '-')}")

            text_msg = "📦 *وضعیت بارنامه‌های شما:*\n\n" + "\n".join(lines)
            if len(rows) > 20:
                text_msg += f"\n\n... و {len(rows)-20} بارنامه دیگر"

            # اگه راننده در یک بارنامه‌ی رد شده گیر کرده، دکمه‌ی بارگذاری مجدد رو هم نشون بده
            last_rejected = next((bn for bn, d in rows if d.get("review", {}).get("status") == "rejected"), None)
            markup = None
            if last_rejected:
                token = make_barname_token(last_rejected)
                markup = InlineKeyboardMarkup([
                    [InlineKeyboardButton(f"🔄 بارگذاری مجدد {last_rejected}", callback_data=f"resubmit_{token}")]
                ])

            await update.message.reply_text(text_msg, parse_mode="Markdown", reply_markup=driver_reply_keyboard())
            if markup:
                await update.message.reply_text("برای بارگذاری مجدد مستندات رد شده 👇", reply_markup=markup)

    elif text == "❌ لغو عملیات":
        context.user_data.clear()
        await update.message.reply_text(
            "❌ عملیات لغو شد.\nهر زمان آماده بودید دوباره شروع کنید.",
            reply_markup=driver_reply_keyboard()
        )
        return ConversationHandler.END

    # ─── دکمه‌های ادمین ───
    elif is_admin and text == "🔍 دریافت مستندات":
        context.user_data["admin_action"] = "get_docs"
        await update.message.reply_text(
            "🔍 شماره بارنامه مورد نظر را وارد کنید:",
            reply_markup=admin_reply_keyboard()
        )

    elif is_admin and text == "📊 لیست بارنامه‌ها":
        db = load_db()
        if not db:
            await update.message.reply_text("📭 هیچ بارنامه‌ای ثبت نشده.", reply_markup=admin_reply_keyboard())
            return
        lines = []
        for bn, data in sorted(db.items(), key=lambda x: x[1].get("created_at",""), reverse=True):
            status = review_status_label(data)
            lines.append(f"{status} | `{bn}` | {len(data.get('documents',{}))} مستند | {data.get('created_at','-')}")
        text_msg = "📊 *لیست بارنامه‌ها:*\n\n" + "\n".join(lines[:20])
        if len(db) > 20:
            text_msg += f"\n\n... و {len(db)-20} بارنامه دیگر"
        await update.message.reply_text(text_msg, parse_mode="Markdown", reply_markup=admin_reply_keyboard())

    elif is_admin and text == "✅ بارنامه‌های تایید شده":
        list_text, _ = build_approved_list_content()
        await update.message.reply_text(list_text, parse_mode="Markdown", reply_markup=admin_reply_keyboard())

    elif is_admin and text == "🕐 نیازمند بررسی":
        list_text, list_markup = build_pending_list_content()
        await update.message.reply_text(list_text, parse_mode="Markdown", reply_markup=list_markup)

    elif is_admin and text == "📈 آمار کلی":
        db = load_db()
        total = len(db)
        completed = sum(1 for d in db.values() if d.get("status") == "completed")
        total_docs = sum(len(d.get("documents", {})) for d in db.values())
        await update.message.reply_text(
            f"📈 *آمار کلی پی‌بار*\n\n"
            f"📦 کل بارنامه‌ها: *{total}*\n"
            f"✅ تکمیل‌شده: *{completed}*\n"
            f"⏳ در انتظار: *{total - completed}*\n"
            f"📄 کل مستندات: *{total_docs}*",
            parse_mode="Markdown",
            reply_markup=admin_reply_keyboard()
        )

    elif is_admin and text == "🗑 حذف بارنامه":
        context.user_data["admin_action"] = "delete_barname"
        await update.message.reply_text(
            "🗑 شماره بارنامه‌ای که می‌خواهید حذف کنید را وارد کنید:",
            reply_markup=admin_reply_keyboard()
        )

    elif is_admin and text == "🏠 منوی اصلی ادمین":
        context.user_data.clear()
        await update.message.reply_text(
            "🔐 *پنل مدیریت پی‌بار*\nاز منوی پایین یا دکمه‌های زیر انتخاب کنید:",
            parse_mode="Markdown",
            reply_markup=admin_reply_keyboard()
        )
        await update.message.reply_text("👇", reply_markup=admin_keyboard())

    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text(
        "❌ عملیات لغو شد.",
        reply_markup=ReplyKeyboardRemove()
    )
    # نمایش منوی اصلی بعد از لغو
    user = update.effective_user
    if user.id in ADMIN_IDS:
        await update.message.reply_text(
            "🔐 *پنل مدیریت پی‌بار*",
            parse_mode="Markdown",
            reply_markup=admin_keyboard()
        )
    else:
        await update.message.reply_text(
            "🏠 *منوی اصلی*\nبرای شروع مجدد دکمه زیر را بزنید:",
            parse_mode="Markdown",
            reply_markup=main_menu_keyboard()
        )
    return ConversationHandler.END


# ─────────── راه‌اندازی ───────────

async def _on_startup(application):
    """قبل از شروع polling، هر webhook فعال روی این ربات را حذف می‌کند تا هرگز polling و webhook هم‌زمان فعال نباشند
    (فعال بودن هم‌زمان هر دو، دلیل رایج ارسال دوبار هر پیام است)."""
    try:
        info = await application.bot.get_webhook_info()
        if info and info.url:
            logger.warning(f"⚠️ یک webhook فعال روی این ربات پیدا شد ({info.url}) — در حال حذف آن تا فقط polling فعال باشد...")
        deleted = await application.bot.delete_webhook(drop_pending_updates=True)
        logger.info(f"✅ بررسی/حذف webhook انجام شد (نتیجه: {deleted}).")
    except Exception as e:
        logger.error(f"⚠️ خطا در بررسی/حذف webhook: {e}")


def main():
    # شناسه‌ی یکتای این نمونه از پروسه — اگر در لاگ‌ها دو INSTANCE_ID متفاوت هم‌زمان دیدید،
    # یعنی دو نسخه از ربات هم‌زمان در حال اجرا هستند و همین باعث ارسال دوبار هر پیام می‌شود.
    instance_id = uuid.uuid4().hex[:8]
    logger.info(f"🆔 INSTANCE_ID این اجرا: {instance_id} — اگر جای دیگری هم‌زمان یک INSTANCE_ID دیگر دیدید، دو نمونه از ربات هم‌زمان اجرا شده‌اند.")

    # تشخیص مشکلات رایج پیکربندی قبل از شروع
    if not ADMIN_IDS:
        logger.warning("⚠️ متغیر ADMIN_IDS خالی است — هیچ ادمینی به پنل مدیریت دسترسی نخواهد داشت!")
    else:
        logger.info(f"✅ آیدی ادمین‌های تعریف‌شده: {ADMIN_IDS}")

    if os.path.exists(DB_FILE):
        try:
            existing = load_db()
            logger.info(f"📦 دیتابیس موجود بارگذاری شد — {len(existing)} بارنامه ثبت‌شده.")
        except Exception as e:
            logger.error(f"⚠️ خطا در خواندن دیتابیس موجود: {e}")
    else:
        logger.warning(
            "⚠️ فایل دیتابیس (database.json) پیدا نشد — یک دیتابیس خالی جدید ساخته می‌شود. "
            "اگر این پیام را بعد از هر ری‌استارت می‌بینید، یعنی فضای ذخیره‌سازی سرور شما پایدار (persistent) نیست "
            "و تمام بارنامه‌ها/تاریخچه‌ها با هر دیپلوی یا ری‌استارت پاک می‌شوند — "
            "برای رفع این مشکل باید یک Volume دائمی در Railway به مسیر پروژه متصل کنید."
        )

    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .base_url(BALE_API_BASE_URL)
        .base_file_url(BALE_API_FILE_URL)
        .post_init(_on_startup)
        .build()
    )

    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            CallbackQueryHandler(handle_start_upload, pattern="^start_upload$"),
            CallbackQueryHandler(resubmit_barname, pattern="^resubmit_"),
            MessageHandler(filters.Regex("^🚀 آپلود مستندات$"), handle_start_upload_button),
        ],
        states={
            WAIT_BARNAME: [
                MessageHandler(filters.Regex("^(🚀 آپلود مستندات|📋 راهنما|📦 وضعیت بارنامه|❌ لغو عملیات|🔍 دریافت مستندات|📊 لیست بارنامه‌ها|📈 آمار کلی|🗑 حذف بارنامه|✅ بارنامه‌های تایید شده|🕐 نیازمند بررسی|🏠 منوی اصلی ادمین)$"), reply_keyboard_handler),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_barname),
                CallbackQueryHandler(back_to_main, pattern="^back_to_main$"),
            ],
            WAIT_PHOTO: [
                MessageHandler(filters.Regex("^(🚀 آپلود مستندات|📋 راهنما|📦 وضعیت بارنامه|❌ لغو عملیات|🔍 دریافت مستندات|📊 لیست بارنامه‌ها|📈 آمار کلی|🗑 حذف بارنامه|✅ بارنامه‌های تایید شده|🕐 نیازمند بررسی|🏠 منوی اصلی ادمین)$"), reply_keyboard_handler),
                MessageHandler(filters.PHOTO | filters.Document.ALL, receive_photo),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_photo),
                CallbackQueryHandler(skip_step, pattern="^skip_step$"),
                CallbackQueryHandler(finish_early, pattern="^finish_early$"),
                CallbackQueryHandler(go_prev_step, pattern="^go_prev_step$"),
                CallbackQueryHandler(go_back_to_step, pattern="^go_back_to_step$"),
                CallbackQueryHandler(back_to_main, pattern="^back_to_main$"),
                CallbackQueryHandler(edit_specific_doc, pattern="^edit_doc_"),
            ],
            WAIT_CONFIRM: [
                MessageHandler(filters.Regex("^(🚀 آپلود مستندات|📋 راهنما|📦 وضعیت بارنامه|❌ لغو عملیات|🔍 دریافت مستندات|📊 لیست بارنامه‌ها|📈 آمار کلی|🗑 حذف بارنامه|✅ بارنامه‌های تایید شده|🕐 نیازمند بررسی|🏠 منوی اصلی ادمین)$"), reply_keyboard_handler),
                CallbackQueryHandler(confirm_photo, pattern="^confirm_photo$"),
                CallbackQueryHandler(retake_photo, pattern="^retake_photo$"),
                CallbackQueryHandler(finish_early, pattern="^finish_early$"),
                CallbackQueryHandler(go_prev_step, pattern="^go_prev_step$"),
                CallbackQueryHandler(final_confirm, pattern="^final_confirm$"),
                CallbackQueryHandler(edit_docs, pattern="^edit_docs$"),
                CallbackQueryHandler(edit_specific_doc, pattern="^edit_doc_"),
                CallbackQueryHandler(back_to_summary, pattern="^back_to_summary$"),
                CallbackQueryHandler(back_to_main, pattern="^back_to_main$"),
            ],
        },
        fallbacks=[
            CommandHandler("cancel", cancel),
            CallbackQueryHandler(back_to_main, pattern="^back_to_main$"),
            MessageHandler(filters.Regex("^❌ لغو عملیات$"), cancel),
        ],
        allow_reentry=True,
    )

    # محافظ ضد آپدیت تکراری — باید زودتر از همه‌ی هندلرهای دیگر اجرا شود (group=-1)
    app.add_handler(TypeHandler(Update, dedupe_update_guard), group=-1)

    app.add_handler(conv_handler)

    # هندلرهای خارج از conversation (ادمین + منو)
    app.add_handler(CallbackQueryHandler(handle_start_upload, pattern="^start_upload$"))
    app.add_handler(CallbackQueryHandler(handle_show_help, pattern="^show_help$"))
    app.add_handler(CallbackQueryHandler(admin_get_docs, pattern="^admin_get$"))
    app.add_handler(CallbackQueryHandler(admin_list_barnames, pattern="^admin_list$"))
    app.add_handler(CallbackQueryHandler(admin_delete_barname, pattern="^admin_delete$"))
    app.add_handler(CallbackQueryHandler(admin_stats, pattern="^admin_stats$"))
    app.add_handler(CallbackQueryHandler(admin_approved_list, pattern="^admin_approved_list$"))
    app.add_handler(CallbackQueryHandler(admin_pending_list, pattern="^admin_pending_list$"))
    app.add_handler(CallbackQueryHandler(admin_pending_page, pattern="^admin_pending_page_"))
    app.add_handler(CallbackQueryHandler(admin_open_barname, pattern="^admin_open_"))
    app.add_handler(CallbackQueryHandler(admin_back, pattern="^admin_back$"))
    app.add_handler(CallbackQueryHandler(back_to_main, pattern="^back_to_main$"))
    # هندلرهای بررسی مستندات توسط ادمین (تایید / عدم تایید / تایید با کسری بار)
    app.add_handler(CallbackQueryHandler(review_approve, pattern="^radm_appr_"))
    app.add_handler(CallbackQueryHandler(review_reject_start, pattern="^radm_rej_"))
    app.add_handler(CallbackQueryHandler(review_partial_start, pattern="^radm_part_"))
    app.add_handler(CallbackQueryHandler(driver_partial_accept, pattern="^drv_ok_"))
    app.add_handler(CallbackQueryHandler(driver_partial_reject_start, pattern="^drv_no_"))
    # هندلر دکمه‌های منوی Reply (باید قبل از admin_text_handler باشه)
    reply_kb_filter = filters.Regex("^(🚀 آپلود مستندات|📋 راهنما|📦 وضعیت بارنامه|❌ لغو عملیات|🔍 دریافت مستندات|📊 لیست بارنامه‌ها|📈 آمار کلی|🗑 حذف بارنامه|✅ بارنامه‌های تایید شده|🕐 نیازمند بررسی|🏠 منوی اصلی ادمین)$")
    app.add_handler(MessageHandler(reply_kb_filter, reply_keyboard_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, admin_text_handler))

    logger.info("🤖 ربات پی‌بار v3 شروع به کار کرد...")
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()
