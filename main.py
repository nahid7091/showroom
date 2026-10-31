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

ROLES_CAN_MANAGE = {"admin", "manager"}

USERS_SHEET = "users"
CARS_SHEET = "cars"
HISTORY_SHEET = "car_logs"

ALLOWED_STATUSES = ["free", "reserved(4 hours)", "بیعانه", "فروخته شده", "غیر قابل فروش"]
ALLOWED_PLACES = ["شوروم فلاحی", "انزلی", "ساری", "تنکابن","سالار","جمهوری","شهرک غرب","اتو شیخ","یافت آباد","...."]

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
    "last_updated_by": "ویرایشگر",
    "options": "متعلقات و آپشن ها",
    "No_Guarantee": "شروع گارانتی",
    "Due": "موعد وکالت تعویض پلاک",
    "technical_issue": "مشکل فنی",
    "physical_issue": "مشکل ظاهری",
    "doc":"سند",
    "plate":"پلاک",
    "price(Rials)":"قیمت(ریال)",
    "buyer":"خریدار",
    "bime":"انقضای بیمه",

    # شیت تاریخچه لاگ‌ها
    "id": "شناسه",
    "car_code": "کد خودرو",
    "changed_at": "زمان تغییر",
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
    "doc",
    "plate",
    "price(Rials)",
    "buyer",
    "bime",
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
    
SHOWN_COLS = [
    "car_code",
    "title",
    "VIN",
    "model_year",
    "color",
    "place",
    "status",
    "reserved_until",
    "options"
]   
    
OTHER_CAR_COLS = [
    "production_year",
    "owner",
    "No_Guarantee",
    "last_updated_at",
    "last_updated_by",
    "Due",
    "physical_issue",
    "technical_issue",
    "doc",
    "plate",
    "price(Rials)",
    "buyer",
    "bime"
]
PENDING_VERIFICATIONS = {}

# -----------------------------
# App
# -----------------------------
app = FastAPI(title="Showroom Management System")
templates = Jinja2Templates(directory="templates")

if os.path.isdir("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")

# -----------------------------
# Helper Functions
# -----------------------------
def get_current_jalali_now() -> str:
    """دریافت تاریخ و ساعت جاری ایران به صورت شمسی و رشته‌ای"""
    now = jdatetime.datetime.now(IRAN_TZ)
    return now.strftime("%Y/%m/%d %H:%M")

def to_jalali_display(dt_val) -> str:
    """تبدیل هر نوع تاریخی به رشته شمسی زیبا"""
    if dt_val is None or pd.isna(dt_val):
        return ""
    val_str = str(dt_val).strip()
    if not val_str or val_str.lower() in ["nan", "none", "nat"]:
        return ""
    
    if val_str.startswith("14"):
        return val_str.replace("-", "/")

    try:
        dt = pd.to_datetime(val_str)
        j_dt = jdatetime.datetime.fromgregorian(datetime=dt)
        return j_dt.strftime("%Y/%m/%d %H:%M:%S")
    except Exception:
        return val_str

def get_jalali_reserve_expiry(hours=4) -> str:
    expiry = jdatetime.datetime.now(IRAN_TZ) + jdatetime.timedelta(hours=hours)
    return expiry.strftime("%Y/%m/%d %H:%M")

def clean_phone(phone) -> str:
    """استانداردسازی شماره موبایل و حذف اعشار و کدهای متفرقه"""
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
            df = pd.DataFrame()
        if not df.empty:
            df = df.map(normalize_digits_value)
            df = df.fillna("")
        return df

def write_sheets(users_df: pd.DataFrame, cars_df: pd.DataFrame, logs_df: Optional[pd.DataFrame] = None):
    with FileLock(LOCK_PATH):
        try:
            all_sheets = pd.read_excel(DB_PATH, sheet_name=None, dtype=str)
        except Exception:
            all_sheets = {}
        all_sheets[USERS_SHEET] = users_df.fillna("")
        all_sheets[CARS_SHEET] = cars_df.fillna("")
        if logs_df is not None:
            all_sheets[HISTORY_SHEET] = logs_df.fillna("")
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
        next_id = int(pd.to_numeric(history_df["id"], errors="coerce").max() or 0) + 1
    elif not history_df.empty:
        next_id = len(history_df) + 1

    new_log = {
        "id": str(next_id),
        "car_code": str(car_code),
        "changed_at": get_current_jalali_now(),
        "changed_by": str(changed_by),
        "from_status": str(from_status),
        "to_status": str(to_status),
        "from_place": str(from_place),
        "to_place": str(to_place),
    }
    history_df = pd.concat([history_df, pd.DataFrame([new_log])], ignore_index=True)
    users_df = read_sheet(USERS_SHEET)
    cars_df = read_sheet(CARS_SHEET)
    write_sheets(users_df, cars_df, history_df)
    
def check_auto_release_reservations(cars_df: pd.DataFrame) -> pd.DataFrame:
    """بررسی منقضی شدن رزروهای ۴ ساعته و تبدیل وضعیت به free"""
    if cars_df.empty or "reserved_until" not in cars_df.columns:
        return cars_df
    
    now_dt = jdatetime.datetime.now(IRAN_TZ)
    changed = False

    for idx, row in cars_df.iterrows():
        if row.get("status") == "reserved(4 hours)":
            r_until = str(row.get("reserved_until", "")).strip()
            if r_until:
                try:
                    exp_dt = jdatetime.datetime.strptime(r_until, "%Y/%m/%d %H:%M").replace(tzinfo=IRAN_TZ)
                    if now_dt >= exp_dt:
                        old_s = cars_df.at[idx, "status"]
                        cars_df.at[idx, "status"] = "free"
                        cars_df.at[idx, "reserved_until"] = ""
                        cars_df.at[idx, "last_updated_at"] = get_current_jalali_now()
                        cars_df.at[idx, "last_updated_by"] = "انقضای سیستمی"
                        changed = True
                        append_history(row["car_code"], "سیستم", old_s, "free", row.get("place", ""), row.get("place", ""))
                except Exception:
                    pass
    if changed:
        users_df = read_sheet(USERS_SHEET)
        write_sheets(users_df, cars_df)
    return cars_df

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
    token = request.cookies.get("session")
    if not token:
        return None
    try:
        data = serializer.loads(token, max_age=60 * 60 * 24 * 7) # تطابق تایم با max_age
        username = data.get("username")
        if not username:
            return None
        
        # خواندن مشخصات کاربر از اکسل
        user = get_user_by_username(username)
        return user
    except (BadSignature, Exception):
        return None
        
    users_df = read_sheet(USERS_SHEET)
    if users_df.empty or "username" not in users_df.columns:
        return None

    user_match = users_df[users_df["username"].astype(str).str.strip() == str(username).strip()]
    if user_match.empty:
        return None

    return user_match.iloc[0].to_dict()

def require_user(request: Request) -> dict:
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=303, headers={"Location": "/login"})
    return user

def require_role(*allowed_roles):
    def dep(user=Depends(require_user)):
        if user["role"] not in allowed_roles:
            raise HTTPException(status_code=403, detail="Forbidden")
        return user
    return dep

def user_has_bale(request: Request, user: dict = Depends(get_current_user)) -> bool:
    
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
        has_bale_flag = str(target.get("user_has_bale", "false")).strip().lower()
        
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
                    if "user_has_bale" not in users_df.columns:
                        users_df["user_has_bale"] = "false"                    
                                    
                    users_df["norm_phone"] = users_df["phone"].apply(clean_phone)

                    if target_user:
                        match_row = users_df[users_df["username"] == target_user]
                        if not match_row.empty:
                            expected_phone = match_row.iloc[0]["norm_phone"]
                            
                            if expected_phone and expected_phone == user_phone:
                                users_df.loc[users_df["username"] == target_user, "bale_chat_id"] = str(chat_id)
                                users_df.loc[users_df["username"] == target_user, "user_has_bale"] = "true"
                                users_df.drop(columns=["norm_phone"], inplace=True)
                                all_sheets[USERS_SHEET] = users_df
                                
                                with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
                                    for s_name, s_df in all_sheets.items():
                                        s_df.to_excel(writer, sheet_name=s_name, index=False)
                                
                                PENDING_VERIFICATIONS.pop(str(chat_id), None)
                                login_token = serializer.dumps({"auto_login": matched_username})
                                send_bale_message(
                                    chat_id,
                                    f"✅ احراز هویت با موفقیت انجام شد!\nحساب کاربری «{matched_username}» فعال شد.\n\nبرای ورود مستقیم به پنل روی دکمه زیر کلیک کنید:",
                                    reply_markup={
                                        "inline_keyboard": [
                                            [{"text": "🌐 ورود به سامانه شوروم", "url": f"https://iranianmoqim.ir/auth/bale-login?token={login_token}"}]
                                        ]
                                    }
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
                            users_df.loc[users_df["username"] == matched_username, "user_has_bale"] = "true"
                            users_df.drop(columns=["norm_phone"], inplace=True, errors="ignore")
                            all_sheets[USERS_SHEET] = users_df
                            
                            with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
                                for s_name, s_df in all_sheets.items():
                                    s_df.to_excel(writer, sheet_name=s_name, index=False)
                                    
                            send_bale_message(chat_id, f"✅ حساب کاربری «{matched_username}» با موفقیت فعال شد.")
                        else:
                            users_df.drop(columns=["norm_phone"], inplace=True, errors="ignore")
                            send_bale_message(chat_id, "❌ شماره موبایل شما در لیست کاربران سیستم یافت نشد.")

    except Exception as e:
        print(f"WEBHOOK PROCESSING ERROR: {e}")

# -----------------------------
# UI Pages
# -----------------------------
@app.get("/auth/bale-login")
def bale_auto_login(token: str):
    try:
        # توکن فقط ۱۰ دقیقه برای کلیک کردن وقت دارد
        data = serializer.loads(token, max_age=600)
        username = data.get("auto_login")
        if not username:
            return RedirectResponse("/login")
        
        # لاگین کردن کاربر و ریدایرکت مستقیم به صفحه خودروها
        response = RedirectResponse(url="/cars", status_code=303)
        return set_session(response, username)
    except Exception:
        return RedirectResponse("/login")

@app.get("/activate-bale", response_class=HTMLResponse)
def activate_bale_page(request: Request, user: dict = Depends(get_current_user)):
    
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
async def read_root(request: Request,user: dict = Depends(get_current_user)):
    
    if user:
        if not user_has_bale(request):
            return RedirectResponse("/activate-bale", status_code=303)
        return RedirectResponse("/cars", status_code=303)
    return RedirectResponse("/login", status_code=303)

@app.get("/login", response_class=HTMLResponse)
async def login_get(request: Request, error: Optional[str] = None):
    user = get_current_user(request)
    if user:
        if not user_has_bale(request):
            return RedirectResponse("/activate-bale", status_code=303)
        return RedirectResponse("/cars", status_code=303)

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
        expected_ans = captcha_serializer.loads(captcha_token)
        if int(str(captcha_answer).strip()) == int(expected_ans):
            captcha_valid = True
    except (BadSignature, ValueError, TypeError):
        captcha_valid = False

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

    users_df = read_sheet(USERS_SHEET)
    matched = users_df[
        (users_df["username"].astype(str).str.strip() == username.strip()) &
        (users_df["password"].astype(str).str.strip() == password.strip())
    ]
    if matched.empty:
        return RedirectResponse(url="/login?error=نام کاربری یا کلمه عبور اشتباه است", status_code=303)

    # حتماً row را تعریف کنید:
    row = matched.iloc[0].to_dict()

    raw_bale = str(row.get("bale_chat_id", "")).strip().lower()
    has_bale_id = raw_bale not in ["", "nan", "none", "-", "0"]
    has_bale_flag = str(row.get("user_has_bale", "")).strip().lower() == "true"
    
    target_url = "/cars" if (has_bale_id and has_bale_flag) else "/activate-bale"
    
    resp = RedirectResponse(target_url, status_code=303)
    set_session(resp, row)
    return resp
    
@app.get("/logout")
def logout():
    resp = RedirectResponse("/login", status_code=302)
    clear_session(resp)
    return resp

@app.get("/cars", response_class=HTMLResponse)
async def cars_page(request: Request, q: str = "", status: str = "", place: str = "",user: dict = Depends(get_current_user)):
    
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    if not user_has_bale(request):
        return RedirectResponse(url="/activate-bale", status_code=303)
        
    cars_df = read_sheet(CARS_SHEET)
    cars_df = check_auto_release_reservations(cars_df)

    if not cars_df.empty:
        if q:
            q_clean = q.strip().lower()
            mask = (
                cars_df["title"].astype(str).str.lower().str.contains(q_clean) |
                cars_df["car_code"].astype(str).str.lower().str.contains(q_clean) |
                cars_df["VIN"].astype(str).str.lower().str.contains(q_clean)
            )
            cars_df = cars_df[mask]

        if status:
            cars_df = cars_df[cars_df["status"].astype(str) == status.strip()]

        if place:
            cars_df = cars_df[cars_df["place"].astype(str) == place.strip()]

    cars_list = cars_df.to_dict(orient="records") if not cars_df.empty else []

    return templates.TemplateResponse(
        request=request,
        name="cars.html",
        context={
            "request": request,
            "user": user,
            "cars": cars_list,
            "q": q,
            "status": status,
            "place": place,
            "shown_cols": SHOWN_COLS,
            "other_cols": OTHER_CAR_COLS,
            "column_titles": COLUMN_TITLES_FA,
            "allowed_statuses": ALLOWED_STATUSES,
            "allowed_places": ALLOWED_PLACES,
            "roles_can_manage": ROLES_CAN_MANAGE
        }
    )


# -----------------------------
# روت فرم افزودن خودرو
# -----------------------------
@app.get("/cars/add", response_class=HTMLResponse)
async def add_car_page(request: Request, user: dict = Depends(get_current_user)):
    
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    if not user_has_bale(request):
        return RedirectResponse(url="/activate-bale", status_code=303)
    if user.get("role") not in ["admin", "manager"]:
        raise HTTPException(status_code=403, detail="شما دسترسی لازم برای افزودن خودرو را ندارید.")
    return templates.TemplateResponse(
        request=request,
        name="add_car.html",
        context={"request": request, "user": user}
    )

# -----------------------------
# ذخیره خودرو جدید
# -----------------------------
@app.post("/cars/add")
async def add_car(
    request: Request,
    car_code: str = Form(...),
    VIN: str = Form(...),
    title: str = Form(...),
    model_year: str = Form(""),
    production_year: str = Form(""),
    color: str = Form(""),
    owner: str = Form(""),
    place: str = Form("نمایشگاه"),
    status: str = Form("free"),
    options: str = Form(""),
    No_Guarantee: str = Form("خیر"),
    Due: str = Form(""),
    technical_issue: str = Form(""),
    physical_issue: str = Form(""),
    user: dict = Depends(get_current_user)
):
    # بررسی دسترسی: فقط ادمین و مدیر نمایشگاه
    if user.get("role") not in ROLES_CAN_MANAGE:
        return JSONResponse(status_code=403, content={"success": False, "message": "فقط مدیران اجازه ثبت خودرو دارند."})

    now_jalali = get_current_jalali_now()
    reserved_until = get_jalali_reserve_expiry(4) if status == "reserved(4 hours)" else ""


    new_car = {
        "car_code": car_code.strip(),
        "VIN": VIN.strip(),
        "title": title.strip(),
        "model_year": model_year.strip(),
        "production_year": production_year.strip(),
        "color": color.strip(),
        "owner": owner.strip(),
        "place": place.strip(),
        "status": status.strip(),
        "reserved_until": reserved_until,
        "last_updated_at": now_jalali,
        "last_updated_by": user.get("username", "system"),
        "options": options.strip(),
        "No_Guarantee": No_Guarantee.strip(),
        "Due": Due.strip(),
        "technical_issue": technical_issue.strip(),
        "physical_issue": physical_issue.strip()
    }

    # قفل فایل اکسل و ثبت رکورد جدید
    with FileLock(LOCK_FILE, timeout=10):
        cars_df = read_sheet(CARS_SHEET)
        
        # بررسی عدم تکراری بودن کد خودرو یا VIN
        if not cars_df.empty and "car_code" in cars_df.columns:
            if str(car_code.strip()) in cars_df["car_code"].astype(str).values:
                return JSONResponse(status_code=400, content={"success": False, "message": "این کد خودرو قبلاً ثبت شده است."})

        # افزودن به دیتافریم
        cars_df = pd.concat([cars_df, pd.DataFrame([new_car])], ignore_index=True)
        users_df = read_sheet(USERS_SHEET)
        write_sheets(users_df, cars_df)

    append_history(new_car["car_code"], user.get("username"), "-", new_car["status"], "-", new_car["place"])

    return JSONResponse(content={"success": True, "message": "خودرو با موفقیت اضافه شد.", "car": new_car})

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
    place: str = Form(None),
    user: dict = Depends(get_current_user)
):
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
    append_history(car_code, user.get("username"), from_status, status, from_place, final_place)

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
@app.post("/users/add")
async def add_user(
    username: str = Form(...),
    password: str = Form(...),
    name: str = Form(...),
    lastname: str = Form(...),
    role: str = Form("expert"),
    phone: str = Form(""),
    user: dict = Depends(get_current_user)
):
    if user.get("role") not in ROLES_CAN_MANAGE:
        return JSONResponse(status_code=403, content={"success": False, "message": "فقط مدیران امکان تعریف کاربر دارند."})

    cleaned_phone = clean_phone(phone)
    with FileLock(LOCK_PATH):
        users_df = read_sheet(USERS_SHEET)
        if not users_df.empty and "username" in users_df.columns:
            if username.strip() in users_df["username"].astype(str).values:
                return JSONResponse(status_code=400, content={"success": False, "message": "این نام کاربری قبلاً استفاده شده است."})

        next_user_code = 101
        if not users_df.empty and "user_code" in users_df.columns:
            next_user_code = int(pd.to_numeric(users_df["user_code"], errors="coerce").max() or 100) + 1

        new_user_data = {
            "user_code": str(next_user_code),
            "username": username.strip(),
            "password": password.strip(),
            "name": name.strip(),
            "lastname": lastname.strip(),
            "role": role.strip(),
            "phone": cleaned_phone,
            "bale_chat_id": "",
            "user_has_bale": "false"
        }
        users_df = pd.concat([users_df, pd.DataFrame([new_user_data])], ignore_index=True)
        cars_df = read_sheet(CARS_SHEET)
        write_sheets(users_df, cars_df)

    return JSONResponse(content={"success": True, "message": f"کاربر {name} {lastname} با موفقیت ثبت شد."})


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "80"))
    uvicorn.run(app, host="0.0.0.0", port=port)
