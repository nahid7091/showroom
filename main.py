import os
import random
import requests
from datetime import datetime
from typing import Optional

import pandas as pd
from fastapi import FastAPI, Request, Form, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from filelock import FileLock
from itsdangerous import URLSafeSerializer, BadSignature

import jdatetime
from zoneinfo import ZoneInfo

IRAN_TZ = ZoneInfo("Asia/Tehran")
    
# -----------------------------
# Config
# -----------------------------
APP_SECRET = os.getenv("APP_SECRET", "CHANGE_ME_SUPER_SECRET")
DB_PATH = os.getenv("EXCEL_DB_PATH", "/data/showroom/database.xlsx")
LOCK_PATH = DB_PATH + ".lock"

# تنظیمات پیام‌رسان بله
BALE_BOT_TOKEN = os.getenv("APP_BALE_TOKEN", "").strip()
BALE_API_URL = f"https://tapi.bale.ai/bot{BALE_BOT_TOKEN}"
BALE_BOT_USERNAME = os.getenv("BALE_BOT_USERNAME", "sammotor_bot")
BALE_BOT_LINK = f"https://ble.ir/{BALE_BOT_USERNAME}"

serializer = URLSafeSerializer(APP_SECRET, salt="session")
captcha_serializer = URLSafeSerializer(APP_SECRET, salt="captcha-salt")

ROLES_CAN_ADD_CAR = {"admin", "showroom_manager"}
ROLES_CAN_DELETE_CAR = {"admin", "showroom_manager"}

USERS_SHEET = "users"
CARS_SHEET = "cars"
HISTORY_SHEET = "car_logs"

ALLOWED_STATUSES = ["free", "reserved(4 hours)", "بیعانه", "فروخته شده", "غیر قابل فروش"]
ALLOWED_PLACES = ["شوروم تهران", "انزلی", "ساری"]

COLUMN_TITLES_FA = {
    # شیت خودروها
    "car_code": "کد خودرو",
    "VIN": "شماره شاسی",
    "title": "مشخصات خودرو",
    "model_year": "سال مدل",
    "production_year": "سال ساخت",
    "color": "رنگ",
    "owner": "مالک",
    "place": "محل خودرو",
    "status": "وضعیت",
    "reserved_until": "مهلت رزرو",
    "last_updated_at": "آخرین به‌روزرسانی",
    "last_updated_by": "ثبت‌کننده آخرین تغییر",
    "options": "متعلقات و آپشن ها",
    "No_Guarantee": "خارج از گارانتی",
    "Due": "موعد وکالت تعویض پلاک",
    "technical_issue": "مشکل فنی",
    "physical_issue": "مشکل ظاهری",

    # شیت تاریخچه لاگ‌ها (Modal History)
    "id": "شناسه",
    "car_code": "کد خودرو",
    "changed_at": "تاریخ و ساعت تغییر",
    "changed_by": "تغییردهنده",
    "from_status": "از وضعیت",
    "to_status": "به وضعیت",
    "from_place": "از محل",
    "to_place": "به محل",

    # شیت کاربران
    "user_code": "کد کاربر",
    "username": "نام کاربری",
    "password": "رمز عبور",
    "bale_chat_id": "شناسه کاربری بله",
    "user_has_bale": "اتصال بله",
    "name": "نام",
    "lastname": "نام خانوادگی",
    "role": "نقش",
    "phone": "شماره موبایل",
}

REQUIRED_USER_COLS = [
    "user_code",
    "username",
    "password",
    "name",
    "lastname",
    "role",
    "phone",
    "bale_chat_id",
    "user_has_bale",
]

REQUIRED_CAR_COLS = [
    "car_code",
    "VIN",
    "title",
    "model_year",
    "production_year",
    "color",
    "owner",
    "place",
    "status",
    "reserved_until",
    "last_updated_at",
    "last_updated_by",
    "options",
    "No_Guarantee",
    "Due",
    "technical_issue",
    "physical_issue",
]

REQUIRED_HISTORY_COLS = [
    "id",
    "car_code",
    "changed_at",
    "changed_by",
    "from_status",
    "to_status",
    "from_place",
    "to_place",
]

PENDING_VERIFICATIONS = {}

# -----------------------------
# App
# -----------------------------
app = FastAPI()
templates = Jinja2Templates(directory="templates")

if os.path.isdir("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")

# -----------------------------
# Helper Functions
# -----------------------------
def get_current_jalali_now() -> str:
    """دریافت تاریخ و ساعت جاری ایران به صورت شمسی و رشته‌ای"""
    now = jdatetime.datetime.now(IRAN_TZ)
    return now.strftime("%Y/%m/%d %H:%M:%S")

def to_jalali_display(dt_val) -> str:
    """تبدیل هر نوع تاریخی (چه رشته میلادی قدیمی، چه آبجکت تاریخ) به رشته شمسی زیبا"""
    if dt_val is None or pd.isna(dt_val):
        return ""
    val_str = str(dt_val).strip()
    if not val_str or val_str.lower() in ["nan", "none", "nat"]:
        return ""
    
    # اگر از قبل شمسی است (با 14 شروع شده)
    if val_str.startswith("14"):
        return val_str.replace("-", "/")

    # اگر میلادی ثبت شده
    try:
        dt = pd.to_datetime(val_str)
        j_dt = jdatetime.datetime.fromgregorian(datetime=dt)
        return j_dt.strftime("%Y/%m/%d %H:%M:%S")
    except Exception:
        return val_str

def get_jalali_reserve_expiry(hours=4) -> str:
    expiry = jdatetime.datetime.now(IRAN_TZ) + jdatetime.timedelta(hours=hours)
    return expiry.strftime("%Y/%m/%d %H:%M:%S")

def clean_phone(phone) -> str:
    """استانداردسازی شماره موبایل و حذف قطعی .0 اکسل"""
    if phone is None or pd.isna(phone):
        return ""
    p = str(phone).strip()
    if "." in p:
        p = p.split(".")[0]
    p = "".join(filter(str.isdigit, p))
    if p.startswith("98") and len(p) == 12:
        p = "0" + p[2:]
    elif p.startswith("9") and len(p) == 10:
        p = "0" + p
    return p

def generate_captcha():
    n1 = random.randint(1, 9)
    n2 = random.randint(1, 9)
    # مستقیماً خود عدد (یا رشته‌ی عدد) را محافظت کنید، نه دیکشنری
    captcha_token = captcha_serializer.dumps(n1 + n2) 
    return n1, n2, captcha_token

def send_bale_message(chat_id: int, text: str, reply_markup: dict = None):
    url = f"{BALE_API_URL}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text
    }
    if reply_markup:
        payload["reply_markup"] = reply_markup
    try:
        requests.post(url, json=payload, timeout=5)
    except Exception as ex:
        print(f"Error sending Bale message: {ex}")

# -----------------------------
# Excel Helpers
# -----------------------------
PERSIAN_ARABIC_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

def normalize_digits_value(val):
    """تبدیل کاراکترهای عددی فارسی و عربی به انگلیسی در مقادیر خوانده‌شده"""
    if isinstance(val, str):
        return val.translate(PERSIAN_ARABIC_DIGITS)
    return val

def read_sheet(sheet_name: str) -> pd.DataFrame:
    with FileLock(LOCK_PATH):
        try:
            df = pd.read_excel(DB_PATH, sheet_name=sheet_name, dtype=str)
        except Exception:
            return pd.DataFrame()

    df = df.map(normalize_digits_value)
    return df

def write_sheets(users_df: pd.DataFrame, cars_df: pd.DataFrame):
    with FileLock(LOCK_PATH):
        try:
            all_sheets = pd.read_excel(DB_PATH, sheet_name=None, dtype=str)
        except Exception:
            all_sheets = {}
        all_sheets[USERS_SHEET] = users_df
        all_sheets[CARS_SHEET] = cars_df
        with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
            for s_name, s_df in all_sheets.items():
                s_df.to_excel(writer, sheet_name=s_name, index=False)

def save_cars(cars_df: pd.DataFrame):
    users_df = read_sheet(USERS_SHEET)
    write_sheets(users_df, cars_df)

def save_users(users_df: pd.DataFrame):
    cars_df = read_sheet(CARS_SHEET)
    write_sheets(users_df, cars_df)

def append_history(car_code: str, changed_by: str, from_status: str, to_status: str, from_place: str, to_place: str):
    history_df = read_sheet(HISTORY_SHEET)
    next_id = 1
    if not history_df.empty and "id" in history_df.columns:
        try:
            next_id = int(pd.to_numeric(history_df["id"], errors="coerce").max() or 0) + 1
        except Exception:
            next_id = len(history_df) + 1

    now_jalali = get_current_jalali_now()
    row = {
        "id": next_id,
        "car_code": str(car_code),
        "changed_at": now_jalali,
        "changed_by": changed_by,
        "from_status": from_status,
        "to_status": to_status,
        "from_place": from_place,
        "to_place": to_place,
    }
    history_df = pd.concat([history_df, pd.DataFrame([row])], ignore_index=True)

    with FileLock(LOCK_PATH):
        try:
            all_sheets = pd.read_excel(DB_PATH, sheet_name=None, dtype=str)
        except Exception:
            all_sheets = {}
        all_sheets[HISTORY_SHEET] = history_df
        with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
            for s_name, s_df in all_sheets.items():
                s_df.to_excel(writer, sheet_name=s_name, index=False)

def expire_reserved_if_needed(system_user_label: str = "SYSTEM"):
    cars_df = read_sheet(CARS_SHEET)
    if cars_df.empty or "reserved_until" not in cars_df.columns:
        return

    now_jalali = get_current_jalali_now()

    status_series = cars_df["status"].fillna("").astype(str).str.strip()
    reserved_until_series = cars_df["reserved_until"].fillna("").astype(str).str.strip()

    expired_mask = (
        (status_series == "reserved(4 hours)")
        & (reserved_until_series != "")
        & (reserved_until_series <= now_jalali)
    )

    expired_idx = cars_df.index[expired_mask].tolist()
    if not expired_idx:
        return

    for i in expired_idx:
        car_code = str(cars_df.at[i, "car_code"])
        old_status = str(cars_df.at[i, "status"])
        old_place = "" if pd.isna(cars_df.at[i, "place"]) else str(cars_df.at[i, "place"])

        cars_df.at[i, "status"] = "free"
        cars_df.at[i, "reserved_until"] = ""
        cars_df.at[i, "last_updated_by"] = system_user_label
        cars_df.at[i, "last_updated_at"] = now_jalali

        append_history(
            car_code=car_code,
            changed_by=system_user_label,
            from_status=old_status,
            to_status="free",
            from_place=old_place,
            to_place=old_place,
        )

    save_cars(cars_df)

# -----------------------------
# Auth & Session Helpers
# -----------------------------
def set_session(response: RedirectResponse, user: dict):
    session_data = serializer.dumps({
        "username": user["username"],
        "role": user["role"],
        "user_code": user["user_code"],
        "name": user["name"],
        "lastname": user["lastname"],
    })
    response.set_cookie("session", session_data, httponly=True, samesite="lax")

def clear_session(response: RedirectResponse):
    response.delete_cookie("session")

def get_current_user(request: Request) -> Optional[dict]:
    token = request.cookies.get("session_token")
    if not token:
        return None
    try:
        username = captcha_serializer.loads(token, max_age=86400 * 7) # اعتبار ۷ روزه
        df_users = read_sheet("users")
        df_users["username"] = df_users["username"].astype(str).str.strip()
        matched = df_users[df_users["username"] == str(username).strip()]
        if not matched.empty:
            user = matched.iloc[0].to_dict()
            # نرمال‌سازی فیلدها برای جلوگیری از باگ
            user["role"] = str(user.get("role", "")).strip().lower()
            user["user_has_bale"] = str(user.get("user_has_bale", "")).strip() in ["1", "true", "True"]
            return user
    except Exception:
        return None
    return None

def require_user(request: Request) -> dict:
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    return user

def require_role(*allowed_roles):
    def dep(user=Depends(require_user)):
        if user["role"] not in allowed_roles:
            raise HTTPException(status_code=403, detail="Forbidden")
        return user
    return dep

def user_has_bale(request: Request) -> bool:
    user = get_current_user(request)
    if not user:
        return False
    try:
        users_df = read_sheet(USERS_SHEET)
        if users_df.empty or "username" not in users_df.columns:
            return False
        
        users_df["username"] = users_df["username"].astype(str).str.strip()
        row = users_df[users_df["username"] == str(user.get("username", "")).strip()]
        if row.empty:
            return False
        
        target = row.iloc[0]
        cid = str(target.get("bale_chat_id", "")).strip()
        has_bale_flag = target.get("user_has_bale", False)

        valid_cid = bool(cid and cid.lower() not in ["nan", "none", "", "-", "0"])
        valid_flag = str(has_bale_flag).strip().lower() in ["true", "1", "yes"]

        return valid_cid and valid_flag
    except Exception as e:
        print(f"user_has_bale check error: {e}")
        return False

def broadcast_car_update(car_code: str, title: str, from_status: str, to_status: str, from_place: str, to_place: str, changed_by: str, changed_at: str):
    try:
        users_df = read_sheet(USERS_SHEET)
        if "bale_chat_id" not in users_df.columns:
            return

        active_chat_ids = users_df["bale_chat_id"].dropna().astype(str).tolist()
        active_chat_ids = [cid.strip() for cid in active_chat_ids if cid.strip() and cid.strip() not in ["nan", "none", "-", "0"]]

        if not active_chat_ids:
            return

        status_emoji = {
            "free": "🟢 آزاد شد",
            "reserved(4 hours)": "⏳ رزرو ۴ ساعته",
            "بیعانه": "🟡 بیعانه دریافت شد",
            "فروخته شده": "🔴 فروخته شد",
            "غیر قابل فروش": "⛔️ غیر قابل فروش"
        }.get(to_status, to_status)

        message = (
            f"🚗 **تغییر وضعیت خودرو**\n\n"
            f"▫️ **کد خودرو:** {car_code}\n"
            f"▫️ **خودرو:** {title}\n"
            f"▫️ **وضعیت جدید:** {status_emoji}\n"
            f"▫️ **محل جدید:** {to_place}\n"
            f"▫️ **وضعیت قبلی:** {from_status} ({from_place})\n"
            f"👤 **ثبت‌کننده:** {changed_by}\n"
            f"🕒 **زمان:** {changed_at}"
        )

        for cid in set(active_chat_ids):
            send_bale_message(cid, message)

    except Exception as e:
        print(f"Error in broadcast_car_update: {e}")

# -----------------------------
# Bale Webhook Route
# -----------------------------
@app.post("/bale/webhook")
async def bale_webhook(request: Request):
    try:
        body = await request.json()
        print(f"BALE WEBHOOK RECEIVED: {body}")
        
        message = body.get("message", {})
        chat_id = message.get("chat", {}).get("id")
        text = (message.get("text") or "").strip()
        contact = message.get("contact")
        
        if not chat_id:
            return {"status": "ok"}

        if text.startswith("/start"):
            parts = text.split()
            if len(parts) > 1:
                target_username = parts[1].strip()
                
                with FileLock(LOCK_PATH):
                    users_df = pd.read_excel(DB_PATH, sheet_name=USERS_SHEET, dtype=str)
                    users_df["username"] = users_df["username"].astype(str).str.strip()
                    
                    if target_username in users_df["username"].values:
                        PENDING_VERIFICATIONS[str(chat_id)] = target_username
                        
                        keyboard = {
                            "keyboard": [[{"text": "📱 ارسال و تأیید شماره موبایل", "request_contact": True}]],
                            "resize_keyboard": True,
                            "one_time_keyboard": True
                        }
                        send_bale_message(
                            chat_id,
                            f"سلام! شما در حال فعال‌سازی دسترسی برای نام کاربری «{target_username}» هستید.\n\n"
                            f"⚠️ جهت راستی‌آزمایی و امنیت سامانه، لطفاً دکمه زیر را لمس نمایید تا شماره همراه اکانت بله شما با شماره تلفن ثبت‌شده در سیستم مقایسه شود:",
                            reply_markup=keyboard
                        )
                    else:
                        send_bale_message(chat_id, f"❌ کاربری با نام «{target_username}» در سامانه یافت نشد.")
            else:
                keyboard = {
                    "keyboard": [[{"text": "📱 ارسال و تأیید شماره موبایل", "request_contact": True}]],
                    "resize_keyboard": True,
                    "one_time_keyboard": True
                }
                send_bale_message(
                    chat_id,
                    "سلام! جهت احراز هویت و فعال‌سازی حساب، لطفاً دکمه زیر را لمس کرده و شماره تلفن خود را ارسال کنید:",
                    reply_markup=keyboard
                )

        elif contact:
            user_phone = clean_phone(str(contact.get("phone_number", "")))
            target_user = PENDING_VERIFICATIONS.get(str(chat_id))

            with FileLock(LOCK_PATH):
                all_sheets = pd.read_excel(DB_PATH, sheet_name=None, dtype=str)
                users_df = all_sheets.get(USERS_SHEET)
                
                if users_df is not None:
                    users_df["username"] = users_df["username"].astype(str).str.strip()
                    
                    if "bale_chat_id" not in users_df.columns:
                        users_df["bale_chat_id"] = ""
                    users_df["bale_chat_id"] = users_df["bale_chat_id"].astype("object")
                    
                    if "user_has_bale" not in users_df.columns:
                        users_df["user_has_bale"] = False
                    users_df["user_has_bale"] = users_df["user_has_bale"].astype("object")
                    
                    users_df["norm_phone"] = users_df["phone"].apply(clean_phone)

                    if target_user:
                        match_row = users_df[users_df["username"] == target_user]
                        if not match_row.empty:
                            expected_phone = match_row.iloc[0]["norm_phone"]
                            
                            if expected_phone and expected_phone == user_phone:
                                users_df.loc[users_df["username"] == target_user, "bale_chat_id"] = str(chat_id)
                                users_df.loc[users_df["username"] == target_user, "user_has_bale"] = True
                                users_df.drop(columns=["norm_phone"], inplace=True)
                                all_sheets[USERS_SHEET] = users_df
                                
                                with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
                                    for s_name, s_df in all_sheets.items():
                                        s_df.to_excel(writer, sheet_name=s_name, index=False)
                                
                                PENDING_VERIFICATIONS.pop(str(chat_id), None)
                                send_bale_message(
                                    chat_id,
                                    f"✅ احراز هویت با موفقیت تأیید شد!\nحساب کاربری «{target_user}» در سیستم فعال شد. اکنون می‌توانید بدون محدودیت از سایت استفاده کنید."
                                )
                            else:
                                users_df.drop(columns=["norm_phone"], inplace=True)
                                send_bale_message(
                                    chat_id,
                                    f"⛔️ عدم تطابق شماره همراه!\n\n"
                                    f"شماره اکانت بله شما ({user_phone}) با شماره تلفن ثبت‌شده برای کاربر «{target_user}» یکسان نیست."
                                )
                        else:
                            users_df.drop(columns=["norm_phone"], inplace=True)
                            send_bale_message(chat_id, "❌ کاربری برای تکمیل این درخواست یافت نشد.")

                    else:
                        match_users = users_df[users_df["norm_phone"] == user_phone]
                        if not match_users.empty:
                            matched_username = match_users.iloc[0]["username"]
                            users_df.loc[users_df["username"] == matched_username, "bale_chat_id"] = str(chat_id)
                            users_df.loc[users_df["username"] == matched_username, "user_has_bale"] = True
                            users_df.drop(columns=["norm_phone"], inplace=True)
                            all_sheets[USERS_SHEET] = users_df
                            
                            with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
                                for s_name, s_df in all_sheets.items():
                                    s_df.to_excel(writer, sheet_name=s_name, index=False)
                                    
                            send_bale_message(chat_id, f"✅ حساب کاربری «{matched_username}» با موفقیت فعال شد.")
                        else:
                            users_df.drop(columns=["norm_phone"], inplace=True)
                            send_bale_message(chat_id, "❌ شماره موبایل شما در لیست کاربران سیستم یافت نشد.")

    except Exception as e:
        print(f"WEBHOOK PROCESSING ERROR: {e}")

    return {"status": "ok"}

# -----------------------------
# UI Pages
# -----------------------------
@app.get("/activate-bale", response_class=HTMLResponse)
async def activate_bale_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    
    if user_has_bale(request):
        return RedirectResponse(url="/cars", status_code=303)

    bale_link = f"{BALE_BOT_LINK}?start={user['username']}"
    
    return templates.TemplateResponse(
        request=request,
        name="activate_bale.html",
        context={"request": request, "user": user, "bale_link": bale_link}
    )

@app.get("/")
async def read_root(request: Request):
    """
    مسیر ریشه: فقط هدایت‌کننده است و نیازی به رندر مستقیم فرم لاگین ندارد.
    """
    user = get_current_user(request)
    if user:
        # اگر لاگین کرده بود، وضعیت بله را بررسی و به مقصد درست هدایت می‌کنیم
        if not user_has_bale(request):
            return RedirectResponse("/activate-bale", status_code=303)
        else:
            return RedirectResponse("/cars", status_code=303)
    
    # اگر لاگین نبود، می‌رود به صفحه لاگین
    else:
        return RedirectResponse("/login", status_code=303)


@app.get("/login", response_class=HTMLResponse)
async def login_get(request: Request):
    """
    مسیر نمایش فرم لاگین:
    اگر قبلاً لاگین کرده باشد، اجازه نمی‌دهد دوباره فرم لاگین را ببیند.
    """
    user = get_current_user(request)
    if user:
        if not user_has_bale(request):
            return RedirectResponse("/activate-bale", status_code=303)
        else:
            return RedirectResponse("/cars", status_code=303)

    # کاربر مهمان است؛ کپچا ساخته شده و فرم لاگین رندر می‌شود
    n1, n2, captcha_token = generate_captcha()
    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={
            "request": request,
            "error": None,
            "num1": n1,
            "num2": n2,
            "captcha_token": captcha_token
        }
    )
    
@app.post("/login")
async def login_post(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    captcha_answer: str = Form(...),
    captcha_token: str = Form(...),
):
    captcha_valid = False
    try:
        # ۱. استخراج پاسخ واقعی از دل توکن با loads
        expected_ans = captcha_serializer.loads(captcha_token)
        if int(str(captcha_answer).strip()) == int(expected_ans):
            captcha_valid = True
    except (BadSignature, ValueError, TypeError):
        captcha_valid = False

    # ۲. اگر کپچا اشتباه بود، خطا را در همان صفحه لاگین به کاربر نشان بده
    if not captcha_valid:
        n1, n2, new_token = generate_captcha()
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "request": request,
                "error": "پاسخ سوال جمع ریاضی اشتباه است. لطفاً دوباره تلاش کنید.",
                "num1": n1,
                "num2": n2,
                "captcha_token": new_token
            },
            status_code=400
        )

    # ۳. خواندن کاربران از اکسل
    users_df = read_sheet(USERS_SHEET)
    username = str(username).strip()
    password = str(password)

    # ۴. بررسی وجود نام کاربری
    if users_df.empty or username not in users_df["username"].astype(str).str.strip().values:
        n1, n2, new_token = generate_captcha()
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "request": request,
                "error": "نام کاربری یا رمز اشتباه است.",
                "num1": n1,
                "num2": n2,
                "captcha_token": new_token
            },
            status_code=400
        )

    # ۵. استخراج ردیف کاربر و تطابق رمز عبور
    row = users_df[users_df["username"].astype(str).str.strip() == username].iloc[0].to_dict()
    saved_password = "" if pd.isna(row.get("password")) else str(row.get("password"))

    if saved_password != password:
        n1, n2, new_token = generate_captcha()
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "request": request,
                "error": "نام کاربری یا رمز اشتباه است.",
                "num1": n1,
                "num2": n2,
                "captcha_token": new_token
            },
            status_code=400
        )

    # ۶. بررسی فعال بودن بله برای مقصد ریدایرکت
    raw_bale = str(row.get("bale_chat_id", "")).strip().lower()
    has_bale = raw_bale not in ["", "nan", "none", "-", "0"]
    
    target_url = "/cars" if has_bale else "/activate-bale"
    
    resp = RedirectResponse(target_url, status_code=303)
    set_session(resp, row)
    return resp


@app.get("/logout")
def logout():
    resp = RedirectResponse("/login", status_code=302)
    clear_session(resp)
    return resp

@app.get("/cars", response_class=HTMLResponse)
async def cars_page(request: Request, q: str = "", status: str = "", place: str = ""):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    if not user_has_bale(request):
        return RedirectResponse(url="/activate-bale", status_code=303)
        
    expire_reserved_if_needed(system_user_label="SYSTEM (auto)")
    cars_df = read_sheet(CARS_SHEET)
    if q:
        q_lower = q.lower()
        mask = (
            cars_df.get("car_code", pd.Series(dtype=str)).astype(str).str.lower().str.contains(q_lower, na=False) |
            cars_df.get("title", pd.Series(dtype=str)).astype(str).str.lower().str.contains(q_lower, na=False) |
            cars_df.get("owner", pd.Series(dtype=str)).astype(str).str.lower().str.contains(q_lower, na=False)
        )
        cars_df = cars_df[mask]

    if status:
        cars_df = cars_df[cars_df["status"].astype(str) == str(status)]

    if place:
        cars_df = cars_df[cars_df["place"].astype(str) == str(place)]

    all_cars_df = read_sheet(CARS_SHEET)
    status_options = ALLOWED_STATUSES
    place_options = ALLOWED_PLACES

    can_add = user["role"] in ROLES_CAN_ADD_CAR

    # تبدیل تاریخ‌ها به شمسی برای نمایش
    if "last_updated_at" in cars_df.columns:
        cars_df["last_updated_at"] = cars_df["last_updated_at"].apply(to_jalali_display)
    if "reserved_until" in cars_df.columns:
        cars_df["reserved_until"] = cars_df["reserved_until"].apply(to_jalali_display)

    cars = cars_df.fillna("").to_dict(orient="records")    
    return templates.TemplateResponse(
        request=request,
        name="cars.html",
        context={
            "request": request,
            "user": user,
            "cars": cars,
            "col_names": COLUMN_TITLES_FA,
            "q": q or "",
            "status": status or "",
            "place": place or "",
            "status_options": status_options,
            "place_options": place_options,
            "can_add": can_add,
        }
    )

# -----------------------------
# روت فرم افزودن خودرو
# -----------------------------
@app.get("/cars/add", response_class=HTMLResponse)
async def add_car_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=302)
    if not user.get("user_has_bale"):
        return RedirectResponse(url="/activate-bale", status_code=302)
    if user.get("role") not in ["admin", "showroom_manager"]:
        raise HTTPException(status_code=403, detail="شما دسترسی لازم برای افزودن خودرو را ندارید.")
    
    return templates.TemplateResponse(
        request=request,
        name="add_car.html",
        context={"user": user}
    )

# -----------------------------
# ذخیره خودرو جدید
# -----------------------------
@app.post("/cars/add")
async def add_car(
    request: Request,
    car_code: str = Form(...),
    VIN: str = Form(""),
    title: str = Form(""),
    model_year: str = Form(""),
    production_year: str = Form(""),
    color: str = Form(""),
    owner: str = Form(""),
    place: str = Form(""),
    status: str = Form("free"),
    options: str = Form(""),
    No_Guarantee: str = Form(""),
    Due: str = Form(""),
    technical_issue: str = Form(""),
    physical_issue: str = Form(""),
    user=Depends(require_role("admin", "showroom_manager"))
):
    if not user_has_bale(request):
        return RedirectResponse(url="/activate-bale", status_code=303)
            
    cars_df = read_sheet(CARS_SHEET)
    car_code = str(car_code).strip()

    if not cars_df.empty and car_code in cars_df["car_code"].astype(str).str.strip().values:
        raise HTTPException(status_code=400, detail="این کد ماشین قبلاً ثبت شده است.")

    if status not in ALLOWED_STATUSES:
        raise HTTPException(status_code=400, detail="وضعیت نامعتبر است.")

    now_jalali = get_current_jalali_now()
    reserved_until = ""
    if status == "reserved(4 hours)":
        reserved_until = get_jalali_reserve_expiry(4)
    
    user_fullname = f'{user.get("name", "")} {user.get("lastname", "")} ({user.get("user_code", "")})'.strip()

    new_row = {
        "car_code": car_code,
        "VIN": VIN.strip(),
        "title": title.strip(),
        "model_year": model_year.strip(),
        "production_year": production_year.strip(),
        "color": color.strip(),
        "owner": owner.strip(),
        "place": place.strip(),
        "status": status,
        "reserved_until": reserved_until,
        "last_updated_at": now_jalali,
        "last_updated_by": user_fullname,
        "options": options.strip(),
        "No_Guarantee": No_Guarantee.strip(),
        "Due": Due.strip(),
        "technical_issue": technical_issue.strip(),
        "physical_issue": physical_issue.strip(),
    }

    cars_df = pd.concat([cars_df, pd.DataFrame([new_row])], ignore_index=True)
    save_cars(cars_df)

    append_history(
        car_code=car_code,
        changed_by=user_fullname,
        from_status="ثبت اولیه",
        to_status=status,
        from_place="-",
        to_place=place.strip(),
    )
    return RedirectResponse("/cars", status_code=303)

@app.post("/cars/add")
async def add_car_submit(
    request: Request,
    title: str = Form(...),
    VIN: str = Form(...),
    model_year: Optional[str] = Form(""),
    production_year: Optional[str] = Form(""),
    color: Optional[str] = Form(""),
    specs: Optional[str] = Form(""),
    place: Optional[str] = Form(""),
    owner: Optional[str] = Form(""),
    status: Optional[str] = Form("free"),
    options: Optional[str] = Form(""),
    No_Guarantee: Optional[str] = Form("خیر"),
    Due: Optional[str] = Form(""),
    technical_issue: Optional[str] = Form("ندارد"),
    physical_issue: Optional[str] = Form("ندارد")
):
    user = get_current_user(request)
    if not user or user.get("role") not in ["admin", "showroom_manager"]:
        raise HTTPException(status_code=403, detail="دسترسی غیرمجاز")

    now_shamsi = jdatetime.datetime.now(IRAN_TZ).strftime("%Y/%m/%d %H:%M")

    with FileLock(LOCK_FILE):
        df_cars = read_sheet("cars")
        
        # تولید کد جدید خودرو
        if not df_cars.empty and "car_code" in df_cars.columns:
            try:
                new_code = int(df_cars["car_code"].max()) + 1
            except:
                new_code = len(df_cars) + 1
        else:
            new_code = 1001

        new_car = {
            "car_code": new_code,
            "VIN": VIN.strip(),
            "title": title.strip(),
            "specs": specs.strip() if specs else "",
            "model_year": model_year.strip() if model_year else "",
            "production_year": production_year.strip() if production_year else "",
            "color": color.strip() if color else "",
            "owner": owner.strip() if owner else "",
            "place": place.strip() if place else "",
            "status": status if status in ALLOWED_STATUSES else "free",
            "reserved_until": "",
            "last_updated_at": now_shamsi,
            "last_updated_by": user.get("username", "system"),
            "options": options.strip() if options else "",
            "No_Guarantee": No_Guarantee,
            "Due": Due.strip() if Due else "",
            "technical_issue": technical_issue.strip() if technical_issue else "ندارد",
            "physical_issue": physical_issue.strip() if physical_issue else "ندارد",
        }

        df_cars = pd.concat([df_cars, pd.DataFrame([new_car])], ignore_index=True)
        write_sheet("cars", df_cars)

        # ثبت لاگ اولیه
        df_logs = read_sheet("car_logs")
        log_entry = {
            "car_code": new_code,
            "action": "افزودن خودرو جدید",
            "user": user.get("username"),
            "timestamp": now_shamsi,
            "details": f"ثبت اولیه خودرو با وضعیت {status} و پلاک/محل {place}"
        }
        df_logs = pd.concat([df_logs, pd.DataFrame([log_entry])], ignore_index=True)
        write_sheet("car_logs", df_logs)

    return RedirectResponse(url="/cars", status_code=303)


@app.get("/cars/history/{car_code}")
def car_history(car_code: str, user=Depends(require_user)):
    history_df = read_sheet(HISTORY_SHEET)
    if history_df.empty:
        return JSONResponse({"car_code": car_code, "items": []})

    df = history_df[history_df["car_code"].astype(str) == str(car_code)].copy()
    df = df.sort_values(by="id", ascending=False) if "id" in df.columns else df
    if "changed_at" in df.columns:
        df["changed_at"] = df["changed_at"].apply(to_jalali_display)
    items = df.fillna("").to_dict(orient="records")
    return JSONResponse({"car_code": car_code, "items": items})

@app.post("/cars/update_status")
async def update_status(
    request: Request,
    car_code: str = Form(...),
    status: str = Form(...),
    place: str = Form(None)
):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    if not user_has_bale(request):
        return RedirectResponse(url="/activate-bale", status_code=303)

    user_fullname = f"{user.get('name', '')} {user.get('lastname', '')}".strip() or user.get("username", "ناشناس")
    
    title = ""
    from_status = ""
    from_place = ""
    final_place = place
    current_jalali_time = get_current_jalali_now()

    with FileLock(LOCK_PATH):
        try:
            all_sheets = pd.read_excel(DB_PATH, sheet_name=None, dtype=str)
            cars_df = all_sheets.get(CARS_SHEET, pd.DataFrame(columns=REQUIRED_CAR_COLS))
            cars_df["car_code"] = cars_df["car_code"].astype(str)
            car_code = str(car_code).strip()
            row_idx = cars_df[cars_df["car_code"] == car_code].index

            if not row_idx.empty:
                idx = row_idx[0]
                title = str(cars_df.loc[idx, "title"]) if "title" in cars_df.columns else "نامشخص"
                from_status = str(cars_df.loc[idx, "status"]) if "status" in cars_df.columns else "نامشخص"
                from_place = str(cars_df.loc[idx, "place"]) if "place" in cars_df.columns else ""

                if not final_place:
                    final_place = from_place

                cars_df.loc[idx, "status"] = status
                if place:
                    cars_df.loc[idx, "place"] = place

                cars_df.loc[idx, "last_updated_at"] = current_jalali_time
                cars_df.loc[idx, "last_updated_by"] = user_fullname

                if status == "reserved(4 hours)":
                    cars_df.loc[idx, "reserved_until"] = get_jalali_reserve_expiry(4)
                else:
                    cars_df.loc[idx, "reserved_until"] = ""
                
                all_sheets[CARS_SHEET] = cars_df

                logs_df = all_sheets.get(HISTORY_SHEET, pd.DataFrame(columns=REQUIRED_HISTORY_COLS))
                new_id = 1
                if not logs_df.empty and "id" in logs_df.columns:
                    try:
                        new_id = int(pd.to_numeric(logs_df["id"], errors="coerce").max() or 0) + 1
                    except Exception:
                        new_id = len(logs_df) + 1
                
                new_log = {
                    "id": new_id,
                    "car_code": car_code,
                    "changed_at": current_jalali_time,
                    "changed_by": user_fullname,
                    "from_status": from_status,
                    "to_status": status,
                    "from_place": from_place,
                    "to_place": final_place
                }
                logs_df = pd.concat([logs_df, pd.DataFrame([new_log])], ignore_index=True)
                all_sheets[HISTORY_SHEET] = logs_df

                with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
                    for s_name, s_df in all_sheets.items():
                        s_df.to_excel(writer, sheet_name=s_name, index=False)

        except Exception as e:
            print(f"Error in update_status: {e}")
            raise e

    try:
        broadcast_car_update(
            car_code=car_code,
            title=title,
            from_status=from_status,
            to_status=status,
            from_place=from_place,
            to_place=final_place,
            changed_by=user_fullname,
            changed_at=current_jalali_time
        )
    except Exception as e:
        print(f"Bale broadcast failed: {e}")

    return RedirectResponse(url="/cars", status_code=303)

# -----------------------------
# Admin helper: create users
# -----------------------------
@app.post("/admin/create_user")
def create_user(
    request: Request,
    user_code: str = Form(...),
    username: str = Form(...),
    password: str = Form(...),
    name: str = Form(...),
    lastname: str = Form(...),
    role: str = Form(...),
    phone: str = Form(""),
    admin=Depends(require_role("admin"))
):
    users_df = read_sheet(USERS_SHEET)
    username = str(username).strip()

    if username in users_df["username"].astype(str).str.strip().values:
        raise HTTPException(status_code=400, detail="این username قبلاً ثبت شده است.")
    if user_code in users_df["user_code"].astype(str).values:
        raise HTTPException(status_code=400, detail="این کد کاربر قبلاً ثبت شده است.")

    row = {
        "user_code": str(user_code),
        "username": username,
        "password": str(password),
        "name": name,
        "lastname": lastname,
        "role": role,
        "phone": clean_phone(phone),
        "bale_chat_id": "",
        "user_has_bale": False
    }
    users_df = pd.concat([users_df, pd.DataFrame([row])], ignore_index=True)
    save_users(users_df)
    return JSONResponse({"ok": True})

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "80"))
    uvicorn.run(app, host="0.0.0.0", port=port)
