import os
import random
import requests
from datetime import datetime
from typing import Optional
import re
import pandas as pd
from fastapi import FastAPI, Request, Query, Form, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from filelock import FileLock
from itsdangerous import URLSafeSerializer, BadSignature
from openpyxl import load_workbook

import jdatetime
from zoneinfo import ZoneInfo

IRAN_TZ = ZoneInfo("Asia/Tehran")

HISTORY_DB_PATH = "/data/showroom/database.xlsx"
CAR_KEY_COL = "car_code"           # نام ستون شناسهٔ یکتای ماشین در cars و تاریخچه
HISTORY_PAGE_SIZE = 50

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

ALLOWED_STATUSES = ["free", "reserved(4 hours)", "بیعانه", "فروخته شده", "غیرقابل فروش","نامشخص"]
ALLOWED_PLACES = ["روبی کار","پارکینگ ۱۳ ولنجک", "اتو منصوری", "شوروم فلاحی", "انزلی-بومی", "انزلی-علوی", "ساری"
                  , "تنکابن","اتو سالار","جمهوری","شهرک غرب","اتو شیخ","یافت آباد","تحویل خریدار",
                 "پشت فنس امیرآباد","پشت فنس گمرک انزلی","منزل آرش سماواتی","کار سنتر محبت"]

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
    "keys":"تعداد سوئیچ",

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
    "user_has_bale"
]

REQUIRED_CAR_COLS = [
    "car_code", "VIN", "title", "model_year", "production_year", 
    "color", "owner", "place", "last_updated_at", "last_updated_by", 
    "expert", "reserved_until", "status", "options", "No_Guarantee", 
    "Due", "technical_issue", "physical_issue", "doc", "plate", 
    "price(Rials)", "buyer", "bime", "keys"
]

REQUIRED_HISTORY_COLS = [
    "id",
    "car_code",
    "changed_at",
    "changed_by",
    "from_status",
    "to_status",
    "from_place",
    "to_place"
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
    "bime",
    "keys"
]

# فیلدهایی که کاربر در فرم «افزودن خودرو» می‌بیند (به ترتیب نمایش)
ADD_CAR_FORM_COLS = [
    "VIN",
    "title",
    "model_year",
    "production_year",
    "color",
    "owner",
    "place",
    "status",
    "No_Guarantee",
    "Due",
    "options",
    "physical_issue",
    "technical_issue",
    "doc",
    "plate",
    "price(Rials)",
    "buyer",
    "bime",
    "keys",
]

# فیلدهای اجباری فرم (ستاره قرمز + required)
REQUIRED_ADD_CAR_FORM_COLS = [
    "VIN",
    "title",
    "model_year",
    "place",
    "status",
]

# فیلدهای چندخطی / تمام‌عرض
FULL_WIDTH_CAR_FORM_COLS = [
    "options",
    "physical_issue",
    "technical_issue",
]

# فیلدهای عددی که باید LTR نمایش داده شوند
LTR_CAR_FORM_COLS = [
    "VIN",
    "model_year",
    "production_year",
    "price(Rials)",
    "bime",
    "keys",
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
    next_id = get_next_id("/data/showroom/database.xlsx", "car_logs", "id")

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
def normalize_shamsi_date(date_val: str) -> str:
    """اعداد فارسی را انگلیسی کرده و فرمت تاریخ را به YYYY/MM/DD استاندارد تبدیل می‌کند"""
    if not date_val:
        return ""
    val = str(date_val).strip()
    
    # اگر مقادیر متنی خاص بود کاری باهاش نداشته باشیم
    if val in ["فاقد گارانتی", "فوری", "-", "nan", "None"]:
        return val if val not in ["nan", "None"] else ""
        
    # تبدیل ارقام فارسی و عربی به انگلیسی
    digits_map = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
    val = val.translate(digits_map).replace("-", "/")
    
    # استانداردسازی پد کردن صفر (مثلاً تبدیل 1403/5/9 به 1403/05/09)
    match = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})$", val)
    if match:
        y, m, d = match.groups()
        return f"{y}/{int(m):02d}/{int(d):02d}"
        
    return val

def set_session(response: RedirectResponse, username: str)-> RedirectResponse:
    token = serializer.dumps({"username": str(username)})
    response.set_cookie(
        key="session",
        value=token,
        httponly=True,
        path="/",
        samesite="lax",
    )
    return response
    
def clear_session(response: RedirectResponse):
    response.delete_cookie("session")

def get_current_user(request: Request) -> Optional[dict]:
    token = request.cookies.get("session")
    if not token:
        return None
    try:
        data = serializer.loads(token)
        username = data.get("username")
        if not username:
            return None
        
        # پیدا کردن کاربر از اکسل (همان کدی که در پروژه داشتید)
        df_users = read_sheet(USERS_SHEET)
        if df_users.empty or "username" not in df_users.columns:
            return None
        df_users["username"] = df_users["username"].astype(str).str.strip()
        matched = df_users[df_users["username"] == str(username).strip()]
        
        if matched.empty:
            return None
            
        return matched.iloc[0].to_dict()
    except Exception as e:
        print(f"Error in get_current_user: {e}")
        return None

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

def user_has_bale(request: Request, user: Optional[dict]=None) -> bool:
    
    if user is None:
        user = get_current_user(request)
        
    if not user:
        return False
        
    try:
        users_df = read_sheet(USERS_SHEET)
        if users_df.empty or "username" not in users_df.columns:
            return False
        
        users_df["username"] = users_df["username"].astype(str).str.strip()
        current_username = str(user.get("username", "")).strip()
        row = users_df[users_df["username"] == current_username]
        if row.empty:
            return False
        
        target = row.iloc[0]
        cid = str(target.get("bale_chat_id", "")).strip()
        has_bale_flag = str(target.get("user_has_bale", "false")).strip().lower()
        
        valid_cid = bool(cid and cid.lower() not in ["nan", "none", "", "-", "0"])
        valid_flag = has_bale_flag in ["true", "1", "yes"]

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


@app.get("/api/cars/{car_code}/history")
def car_status_history(
    request: Request,
    car_code: str,
    offset: int = Query(0, ge=0),
):
    # ۱. احراز هویت با استفاده از سشن سیستم خودتان
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="ابتدا وارد حساب کاربری شوید.")

    # ۲. بررسی تأیید اتصال بله
    if not user_has_bale(request, user):
        raise HTTPException(status_code=403, detail="اتصال به پیام‌رسان بله فعال نیست.")

    # ۳. خواندن شیت لاگ‌ها به کمک تابع امن سیستم (read_sheet)
    history_df = read_sheet(HISTORY_SHEET)
    if history_df.empty or "car_code" not in history_df.columns:
        return {"items": [], "next_offset": offset, "has_more": False}

    # ۴. فیلتر کردن لاگ‌های مربوط به همین خودرو
    rows = history_df[history_df["car_code"].astype(str).str.strip() == str(car_code).strip()].copy()

    # مرتب‌سازی: جدیدترین‌ها در ابتدا (بر اساس شناسه id)
    if "id" in rows.columns:
        rows["num_id"] = pd.to_numeric(rows["id"], errors="coerce").fillna(0)
        rows = rows.sort_values(by="num_id", ascending=False)

    total = len(rows)
    page = rows.iloc[offset:offset + HISTORY_PAGE_SIZE]

    # ۵. قالب‌بندی خروجی با ستون‌های واقعی اکسل شما
    items = []
    for _, r in page.iterrows():
        items.append({
            "changed_at": to_jalali_display(r.get("changed_at", "")),
            "changed_by": r.get("changed_by", "—"),
            "from_status": r.get("from_status", "—"),
            "to_status": r.get("to_status", "—"),
            "from_place": r.get("from_place", "—"),
            "to_place": r.get("to_place", "—"),
        })

    next_offset = offset + len(items)
    return {
        "items": items,
        "next_offset": next_offset,
        "has_more": next_offset < total,
    }

@app.get("/api/users/list")
async def list_users(request: Request, user: dict = Depends(get_current_user)):
    # بررسی دسترسی: فقط ادمین و مدیر
    if not user or user.get("role") not in ROLES_CAN_MANAGE:
        raise HTTPException(status_code=403, detail="دسترسی غیرمجاز. فقط مدیران مجاز هستند.")

    users_df = read_sheet(USERS_SHEET)
    if users_df.empty:
        return {"users": []}

    # اطمینان از وجود ستون‌های استاندارد و جایگزینی مقادیر خالی
    users_df = users_df.fillna("")
    
    # ساخت لیست دیکشنری از کاربران
    users_list = users_df.to_dict(orient="records")
    return {"users": users_list}


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
                                login_token = serializer.dumps({"auto_login": target_user})
                                send_bale_message(
                                    chat_id,
                                    f"✅ احراز هویت با موفقیت انجام شد!\nحساب کاربری «{target_user}» فعال شد.\n\nبرای ورود مستقیم به پنل روی دکمه زیر کلیک کنید:",
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
        data = serializer.loads(token)
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
    set_session(resp, row["username"])
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

    # -------------------------------------------------------------
    # اصلاح دائمی در اکسل: اگر فروخته شده و محل خالی است -> تحویل خریدار
    # -------------------------------------------------------------
    if not cars_df.empty and "status" in cars_df.columns and "place" in cars_df.columns:
        # پیدا کردن خودروهای فروخته شده‌ای که محل ندارند یا nan هستند
        is_sold = cars_df["status"].astype(str).str.strip() == "فروخته شده"
        is_empty_place = cars_df["place"].astype(str).str.strip().isin(["", "nan", "None"])
        
        needs_update_mask = is_sold & is_empty_place
        
        # اگر موردی نیاز به اصلاح داشت، یک‌بار در اکسل ذخیره می‌کنیم
        if needs_update_mask.any():
            cars_df.loc[needs_update_mask, "place"] = "تحویل خریدار"
            # استفاده از تابع امن خودتان که شیت‌های users و car_logs را حفظ می‌کند:
            save_cars(cars_df)
            
    # -------------------------------------------------------------
    # حالا فیلترهای جستجو را برای نمایش اعمال می‌کنیم:
    # -------------------------------------------------------------
    filtered_df = cars_df.copy()
    if not filtered_df.empty:
        if q and q.strip():
            q_clean = q.strip().lower()
            # در تمام ستون‌ها می‌گردد؛ اگر در هر ستونی پیدا شد، آن ردیف را برمی‌دارد
            mask = filtered_df.astype(str).apply(
                lambda row: row.str.lower().str.contains(q_clean, na=False).any(), 
                axis=1
            )
            filtered_df = filtered_df[mask]

        if status:
            filtered_df = filtered_df[filtered_df["status"].astype(str).str.strip() == status.strip()]

        if place:
            filtered_df = filtered_df[filtered_df["place"].astype(str).str.strip() == place.strip()]

    # ۱. تبدیل دیتافریم فیلتر شده به لیست دیکشنری
    cars_list = filtered_df.to_dict(orient="records") if not filtered_df.empty else []

    # ۲. تابع مرتب‌سازی
    def car_display_sort_key(car):
        place = str(car.get("place") or "")
        place = place.replace("ي", "ی").replace("ك", "ک").replace("\u200c", " ")
        normalized_place = re.sub(r"\s+", "", place)
    
        is_not_fallahi = normalized_place != "شورومفلاحی"
    
        digits_map = str.maketrans(
            "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
            "01234567890123456789",
        )
        code = str(car.get("car_code", "")).strip().translate(digits_map)
    
        # کدهای عددی، از جمله کدهایی که از اکسل به شکل 123.0 خوانده شده‌اند
        match = re.fullmatch(r"(\d+)(?:\.0+)?", code)
        if match:
            code_key = (0, int(match.group(1)), "")
        elif code.lower() in ("", "nan", "none"):
            code_key = (2, 0, "")
        else:
            code_key = (1, 0, code.casefold())
    
        return (is_not_fallahi, *code_key)

    # ۳. مرتب‌سازی لیست خودروها (بدون نیاز به hasattr)
    cars = sorted(cars_list, key=car_display_sort_key)

    # ۴. ارسال به قالب
    return templates.TemplateResponse(
        request=request,
        name="cars.html",
        context={
            "request": request,
            "user": user,
            "cars": cars,
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

def get_next_id(file_path: str, sheet_name: str, id_column: str = "id") -> int:
    """
    بیشترین مقدار ستون شناسه را پیدا کرده و مقدار بعدی (+1) را برمی‌گرداند.
    اگر شیت خالی باشد یا هیچ رکوردی نباشد، عدد 1 را برمی‌گرداند.
    
    مثال استفاده:
        next_car_id = get_next_id("/data/showroom/database.xlsx", "cars", "car_code")
        next_log_id = get_next_id("/data/showroom/database.xlsx", "car_logs", "id")
    """
    try:
        # فقط همان شیت مورد نظر خوانده می‌شود
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # اگر ستون وجود نداشت یا دیتایی نبود
        if df.empty or id_column not in df.columns:
            return 1
            
        # تبدیل ستون به رشته، نرمال‌سازی ارقام فارسی/عربی و تبدیل امن به عدد
        s = df[id_column].astype(str)
        
        # تبدیل اعداد فارسی/عربی به انگلیسی در صورت وجود
        persian_digits = "۰۱۲۳۴۵۶۷۸۹"
        arabic_digits = "٠١٢٣٤٥٦٧٨٩"
        for i in range(10):
            s = s.str.replace(persian_digits[i], str(i))
            s = s.str.replace(arabic_digits[i], str(i))
            
        # استخراج مقادیر عددی (خطاها تبدیل به NaN می‌شوند)
        numeric_ids = pd.to_numeric(s, errors="coerce").dropna()
        
        if numeric_ids.empty:
            return 1
            
        # بزرگترین مقدار را گرفته و یک واحد اضافه می‌کند
        return int(numeric_ids.max()) + 1

    except Exception:
        # در صورت نبود شیت یا فایل، پیش‌فرض از 1 شروع می‌شود
        return 1

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
    VIN: str = Form(...),
    title: str = Form(...),
    model_year: str = Form(""),
    production_year: str = Form(""),
    color: str = Form(""),
    owner: str = Form(""),
    place: str = Form("نمایشگاه"),
    status: str = Form("free"),
    No_Guarantee: str = Form(""),
    Due: str = Form(""),
    options: str = Form(""),
    physical_issue: str = Form(""),
    technical_issue: str = Form(""),
    doc: str = Form(""),           # <--- ۱. افزودن پارامتر سند
    keys: str = Form(""),          # <--- ۲. افزودن پارامتر سوئیچ
    price_rials: str = Form("", alias="price(Rials)"), # <--- ۳. دریافت قیمت با alias دقیق هدر اکسل
    user: dict = Depends(get_current_user)
):
    # بررسی دسترسی: فقط ادمین و مدیر نمایشگاه
    if not user or user.get("role") not in ROLES_CAN_MANAGE:
        return JSONResponse(status_code=403, content={"success": False, "message": "فقط مدیران اجازه ثبت خودرو دارند."})

    clean_vin = VIN.strip()
    clean_place = place.strip()
    clean_status = status.strip()
    now_jalali = get_current_jalali_now()
    reserved_until = get_jalali_reserve_expiry(4) if status == "reserved(4 hours)" else ""
    final_place="تحویل خریدار" if status=="فروخته شده" else clean_place

    try:
        with FileLock(LOCK_PATH, timeout=10):
            all_sheets = pd.read_excel(DB_PATH, sheet_name=None, dtype=str)
            cars_df = all_sheets.get(CARS_SHEET, pd.DataFrame(columns=REQUIRED_CAR_COLS)).fillna("")
            logs_df = all_sheets.get(HISTORY_SHEET, pd.DataFrame(columns=REQUIRED_HISTORY_COLS)).fillna("")

            # بررسی تکراری نبودن شاسی
            if not cars_df.empty and "VIN" in cars_df.columns:
                if clean_vin and clean_vin in cars_df["VIN"].astype(str).str.strip().values:
                    return JSONResponse(status_code=400, content={"success": False, "message": "این شماره شاسی (VIN) قبلاً ثبت شده است."})
            next_car_code = get_next_id("/data/showroom/database.xlsx", "cars", "car_code")     
            new_car = {
                "car_code": next_car_code,
                "VIN": clean_vin,
                "title": title.strip(),
                "model_year": model_year.strip(),
                "production_year": production_year.strip(),
                "color": color.strip(),
                "owner": owner.strip(),
                "place": final_place,
                "status": clean_status,
                "last_updated_at": now_jalali,
                "last_updated_by": user.get("username", "system"),
                "expert": "", # ستون اکسل
                "reserved_until": "",
                "options": options.strip(),
                "No_Guarantee": normalize_shamsi_date(No_Guarantee),
                "Due": normalize_shamsi_date(Due),
                "technical_issue": technical_issue.strip(),
                "physical_issue": physical_issue.strip(),
                "doc": doc.strip(),                 # <--- ذخیره در اکسل
                "plate": "",                         # اگر فیلد پلاک در فرم هست اضافه شود
                "price(Rials)": price_rials.strip(), # <--- ذخیره با نام دقیق اکسل
                "buyer": "",
                "bime": "",
                "keys": keys.strip()                 # <--- ذخیره در اکسل
            }
            cars_df = pd.concat([cars_df, pd.DataFrame([new_car])], ignore_index=True)
            all_sheets[CARS_SHEET] = cars_df
            new_log = {
                "id": get_next_id("/data/showroom/database.xlsx", "car_logs", "id"),
                "car_code": str(next_car_code),
                "changed_at": now_jalali,
                "changed_by": user.get("username", "سیستم"),
                "from_status": "-",
                "to_status": clean_status,
                "from_place": "-",
                "to_place": final_place
            }
            logs_df = pd.concat([logs_df, pd.DataFrame([new_log])], ignore_index=True)
            all_sheets[HISTORY_SHEET] = logs_df

            # ذخیره نهایی اکسل فقط با ۱ بار باز شدن
            with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
                for s_name, s_df in all_sheets.items():
                    s_df.to_excel(writer, sheet_name=s_name, index=False)
            
        return JSONResponse(content={"success": True, "message": "خودرو با موفقیت اضافه شد.", "car": new_car})

    except Exception as e:
      print(f"Error in add_car: {e}")
      return JSONResponse(status_code=500, content={"success": False, "message": f"خطا در ثبت خودرو: {str(e)}"})


    

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
    place: str = Form(...),
    user: dict = Depends(get_current_user)
):
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    if not user_has_bale(request):
        return RedirectResponse(url="/activate-bale", status_code=303)

    status = (status or "").strip()
    place = (place or "").strip()
    car_code = (car_code or "").strip()

    if not status or not place:
        raise HTTPException(
            status_code=400,
            detail="وضعیت و محل خودرو هر دو الزامی هستند."
        )

    user_fullname = f"{user.get('name', '')} {user.get('lastname', '')}".strip() or user.get("username", "ناشناس")
    
    title = ""
    from_status = ""
    from_place = ""
    current_jalali_time = get_current_jalali_now()

    with FileLock(LOCK_PATH, timeout=10):
        try:
            all_sheets = pd.read_excel(DB_PATH, sheet_name=None, dtype=str)
            cars_df = all_sheets.get(CARS_SHEET, pd.DataFrame(columns=REQUIRED_CAR_COLS))
            if cars_df.empty or "car_code" not in cars_df.columns:
                raise HTTPException(status_code=404, detail="اطلاعات خودروها یافت نشد.")
            cars_df["car_code"] = cars_df["car_code"].astype(str)
            car_code = str(car_code).strip()
            row_idx = cars_df[cars_df["car_code"] == car_code].index

            if row_idx.empty:
                raise HTTPException(status_code=404, detail=f"خودرو با کد {car_code} یافت نشد.")
            idx = row_idx[0]
            title = str(cars_df.loc[idx, "title"]) if "title" in cars_df.columns else "نامشخص"
            from_status = str(cars_df.loc[idx, "status"]) if "status" in cars_df.columns else "نامشخص"
            if from_status == "فروخته شده" and user.get("role") not in ROLES_CAN_MANAGE:
                raise HTTPException(status_code=403, detail="خودروی فروخته شده فقط توسط مدیر قابل ویرایش است.")

            from_place = str(cars_df.loc[idx, "place"]) if "place" in cars_df.columns else ""

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

            status = status.strip()
            place = place.strip()
            if not status or not place:
                raise HTTPException(
                    status_code=400,
                    detail="وضعیت و محل خودرو هر دو الزامی هستند."
                )
                   
            logs_df = all_sheets.get(HISTORY_SHEET, pd.DataFrame(columns=REQUIRED_HISTORY_COLS))
            new_id = get_next_id("/data/showroom/database.xlsx", "car_logs", "id")
                
            new_log = {
                "id": new_id,
                "car_code": car_code,
                "changed_at": current_jalali_time,
                "changed_by": user_fullname,
                "from_status": from_status,
                "to_status": status,
                "from_place": from_place,
                "to_place": place
            }
            logs_df = pd.concat([logs_df, pd.DataFrame([new_log])], ignore_index=True)
            all_sheets[HISTORY_SHEET] = logs_df

            with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
                for s_name, s_df in all_sheets.items():
                    s_df.to_excel(writer, sheet_name=s_name, index=False)

        except Exception as e:
            print(f"Error in update_status: {e}")
            raise HTTPException(status_code=500, detail=f"خطا در ثبت وضعیت: {str(e)}")

    try:
        broadcast_car_update(
            car_code=car_code,
            title=title,
            from_status=from_status,
            to_status=status,
            from_place=from_place,
            to_place=place,
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
    if not user or user.get("role") not in ROLES_CAN_MANAGE:
        return JSONResponse(status_code=403, content={"success": False, "message": "فقط مدیران امکان تعریف کاربر دارند."})

    clean_uname = username.strip()
    cleaned_phone = clean_phone(phone)

    try:
        # ۲. فقط یک بار قفل و یک بار باز کردن کل اکسل
        with FileLock(LOCK_PATH, timeout=10):
            try:
                all_sheets = pd.read_excel(DB_PATH, sheet_name=None, dtype=str)
            except Exception:
                all_sheets = {}

            users_df = all_sheets.get(USERS_SHEET, pd.DataFrame(columns=REQUIRED_USER_COLS)).fillna("")

            # بررسی تکراری نبودن نام کاربری
            if not users_df.empty and "username" in users_df.columns:
                existing_users = users_df["username"].astype(str).str.strip().values
                if clean_uname in existing_users:
                    return JSONResponse(
                        status_code=400, 
                        content={"success": False, "message": "این نام کاربری قبلاً استفاده شده است."}
                    )
            next_user_code = get_next_id("/data/showroom/database.xlsx", "users", "user_code")

            new_user_data = {
                "user_code": str(next_user_code),
                "username": clean_uname,
                "password": password.strip(),
                "name": name.strip(),
                "lastname": lastname.strip(),
                "role": role.strip(),
                "phone": cleaned_phone,
                "bale_chat_id": "",
                "user_has_bale": "false"
            }
            users_df = pd.concat([users_df, pd.DataFrame([new_user_data])], ignore_index=True)
            all_sheets[USERS_SHEET] = users_df

            # ذخیره تمیز و سریع
            with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
                for s_name, s_df in all_sheets.items():
                    s_df.to_excel(writer, sheet_name=s_name, index=False)

        return JSONResponse(
            content={"success": True, "message": f"کاربر {name} {lastname} با موفقیت ثبت شد."}
        )

    except Exception as e:
        print(f"Error in add_user: {e}")
        return JSONResponse(
            status_code=500, 
            content={"success": False, "message": f"خطا در ثبت کاربر: {str(e)}"}
        )

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "80"))
    uvicorn.run(app, host="0.0.0.0", port=port)
