import os
import random
from fastapi import Body

from datetime import datetime
from typing import Optional

import pandas as pd
from fastapi import FastAPI, Request, Form, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from filelock import FileLock
from itsdangerous import URLSafeSerializer, BadSignature

# -----------------------------
# Config
# -----------------------------
APP_SECRET = os.getenv("APP_SECRET", "CHANGE_ME_SUPER_SECRET")
DB_PATH = os.getenv("EXCEL_DB_PATH", "/data/showroom/database.xlsx")
LOCK_PATH = DB_PATH + ".lock"

# توکن ربات بله که از بازوی BotFather بله گرفتید
BALE_BOT_TOKEN = os.getenv("APP_BALE_TOKEN", "").strip()
BALE_API_URL = f"https://tapi.bale.ai/bot{BALE_BOT_TOKEN}"

BALE_BOT_USERNAME = os.getenv("BALE_BOT_USERNAME", "sammotor_bot")  # بدون @
BALE_BOT_LINK = f"https://ble.ir/{BALE_BOT_USERNAME}"

serializer = URLSafeSerializer(APP_SECRET, salt="session")
captcha_serializer = URLSafeSerializer(APP_SECRET, salt="captcha-salt")

def send_bale_message(chat_id: int | str, text: str, reply_markup: dict = None) -> bool:
    """ارسال متن ساده یا همراه با کیبورد در بله"""
    if not BALE_BOT_TOKEN:
        print("BALE: Token not configured!")
        return False
    payload = {"chat_id": chat_id, "text": text}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    try:
        resp = requests.post(f"{BALE_API_URL}/sendMessage", json=payload, timeout=8)
        return resp.status_code == 200
    except Exception as e:
        print(f"BALE send error: {e}")
        return False
@app.get("/activate-bale", response_class=HTMLResponse)
async def activate_bale_page(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    # اگر قبلاً فعال کرده باشد، به /cars بفرست
    try:
        users_df = read_sheet(USERS_SHEET)
        users_df["username"] = users_df["username"].astype(str)
        row = users_df[users_df["username"] == str(user["username"])]
        has_chat = False
        if not row.empty:
            cid = str(row.iloc[0].get("bale_chat_id", "")).strip()
            has_chat = bool(cid and cid.lower() != "nan")
        if has_chat:
            return RedirectResponse(url="/cars", status_code=303)
    except Exception as e:
        print(f"activate_bale_page check error: {e}")

    # صفحه راهنما به فارسی
    html = f"""
    <!doctype html>
    <html lang="fa" dir="rtl">
    <head>
      <meta charset="utf-8">
      <title>فعالسازی اعلان‌های بله</title>
      <style>
        body {{ font-family: Tahoma, sans-serif; background:#fafafa; margin:0; }}
        .card {{ max-width:680px; margin:40px auto; background:#fff; border:1px solid #eee; border-radius:8px; padding:24px; line-height:1.9; }}
        .btn {{ display:inline-block; background:#0b72e7; color:#fff; padding:10px 16px; border-radius:6px; text-decoration:none; }}
        .muted {{ color:#666; font-size:13px; }}
        ol {{ padding-right:18px; }}
        code {{ background:#f5f5f5; padding:2px 6px; border-radius:4px; }}
      </style>
    </head>
    <body>
      <div class="card">
        <h2>فعالسازی اعلان‌های بله الزامی است</h2>
        <p>برای استفاده از سامانه نمایشگاه، ابتدا اعلان‌های «بله» را فعال کنید تا از تغییر وضعیت خودروها مطلع شوید.</p>
        <ol>
          <li>روی دکمه زیر بزنید تا وارد ربات شوید.</li>
          <li>در بله دکمه <strong>Start</strong> را بزنید.</li>
          <li>دکمه «📱 ارسال و تایید شماره موبایل» را لمس کنید تا <strong>Share Contact</strong> انجام شود.</li>
        </ol>
        <p><a class="btn" href="{BALE_BOT_LINK}" target="_blank">ورود به ربات بله</a></p>
        <p class="muted">پس از انجام مراحل بالا، به این صفحه برگردید و آن را رفرش کنید.</p>
        <hr>
        <p class="muted">اگر دکمه ارسال شماره را نمی‌بینید، یک‌بار دستور <code>/start</code> را بفرستید.</p>
      </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html)

    def user_has_bale(request: Request) -> bool:
    user = get_current_user(request)
    if not user:
        return False
    try:
        users_df = read_sheet(USERS_SHEET)
        users_df["username"] = users_df["username"].astype(str)
        row = users_df[users_df["username"] == str(user["username"])]
        if row.empty:
            return False
        cid = str(row.iloc[0].get("bale_chat_id", "")).strip()
        return bool(cid and cid.lower() != "nan")
    except Exception as e:
        print(f"user_has_bale error: {e}")
        return False
        
def clean_phone(phone: str) -> str:
    """استانداردسازی شماره موبایل برای جلوگیری از مغایرت"""
    if not phone:
        return ""
    p = str(phone).replace(" ", "").replace("-", "").strip()
    if p.startswith("+98"):
        p = "0" + p[3:]
    elif p.startswith("98"):
        p = "0" + p[2:]
    elif not p.startswith("0") and len(p) == 10:
        p = "0" + p
    return p

        
def generate_captcha():
    """تولید دو عدد تصادفی و توکن رمزنگاری‌شده از حاصل جمع"""
    n1 = random.randint(1, 9)
    n2 = random.randint(1, 9)
    token = captcha_serializer.dumps(n1 + n2)
    return n1, n2, token
    
ROLES_CAN_ADD_CAR = {"admin", "showroom_manager"}
ROLES_CAN_DELETE_CAR = {"admin", "showroom_manager"}

USERS_SHEET = "users"
CARS_SHEET = "cars"
HISTORY_SHEET = "car_logs"

ALLOWED_STATUSES = ["free", "reserved(4 hours)", "بیعانه", "فروخته شده"]

REQUIRED_USER_COLS = ["user_code", "username", "password", "name", "lastname", "role"]

REQUIRED_CAR_COLS = [
    "car_code",
    "VIN",
    "title",
    "specs",              # اضافه شد: مشخصات
    "model_year",
    "production_year",
    "color",
    "owner",
    "place",
    "status",
    "reserved_until",     # اضافه شد: زمان پایان رزرو
    "last_updated_at",
    "last_updated_by",
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
# -----------------------------
# App
# -----------------------------
app = FastAPI()
templates = Jinja2Templates(directory="templates")

if os.path.isdir("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")

# -----------------------------
# Excel helpers (with lock)
# -----------------------------
def ensure_db():
    """Create Excel with required sheets/columns if missing."""
    if os.path.exists(DB_PATH):
        return
    with FileLock(LOCK_PATH):
        users_df = pd.DataFrame(columns=REQUIRED_USER_COLS)
        cars_df = pd.DataFrame(columns=REQUIRED_CAR_COLS)
        history_df = pd.DataFrame(columns=REQUIRED_HISTORY_COLS)
        with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
            users_df.to_excel(writer, sheet_name=USERS_SHEET, index=False)
            cars_df.to_excel(writer, sheet_name=CARS_SHEET, index=False)
            history_df.to_excel(writer, sheet_name=HISTORY_SHEET, index=False)

def read_sheet(sheet_name: str) -> pd.DataFrame:
    with FileLock(LOCK_PATH):
        try:
            return pd.read_excel(DB_PATH, sheet_name=sheet_name)
        except ValueError:
            if sheet_name == HISTORY_SHEET:
                columns = ["id", "car_code", "changed_at", "changed_by", "from_status", "to_status", "from_place", "to_place"]
                return pd.DataFrame(columns=columns)
            elif sheet_name == CARS_SHEET:
                columns = ["car_code", "VIN", "title", "specs", "model_year", "production_year", "color", "owner", "place", "status", "reserved_until", "last_updated_at", "last_updated_by"]
                return pd.DataFrame(columns=columns)
            return pd.DataFrame()


def write_sheets(users_df: pd.DataFrame, cars_df: pd.DataFrame):
    with FileLock(LOCK_PATH):
        with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
            users_df.to_excel(writer, sheet_name=USERS_SHEET, index=False)
            cars_df.to_excel(writer, sheet_name=CARS_SHEET, index=False)

def save_cars(cars_df: pd.DataFrame):
    users_df = read_sheet(USERS_SHEET)
    write_sheets(users_df, cars_df)

def save_users(users_df: pd.DataFrame):
    cars_df = read_sheet(CARS_SHEET)
    write_sheets(users_df, cars_df)

@app.post("/bale/webhook")
async def bale_webhook(update: dict = Body(...)):
    try:
        message = update.get("message") or {}
        chat = message.get("chat") or {}
        chat_id = chat.get("id")
        text = (message.get("text") or "").strip()
        contact = message.get("contact")  # وقتی کاربر دکمه ارسال شماره را می‌زند

        if not chat_id:
            return {"ok": True}

        # ۱. اگر کاربر /start زد -> بهش دکمه ارسال شماره تماس نشان بده
        if text.startswith("/start"):
            keyboard = {
                "keyboard": [
                    [
                        {
                            "text": "📱 ارسال و تایید شماره موبایل",
                            "request_contact": True  # کلید جادویی بله برای گرفتن شماره واقعی کاربر
                        }
                    ]
                ],
                "resize_keyboard": True,
                "one_time_keyboard": True
            }
            send_bale_message(
                chat_id,
                "سلام! به سامانه اعلان‌های نمایشگاه خودرو خوش آمدید.\n\n"
                "برای فعال‌سازی و دریافت لحظه‌ای وضعیت خودروها، لطفاً روی دکمه زیر کلیک کنید تا شماره موبایل شما تایید شود:",
                reply_markup=keyboard
            )
            return {"ok": True}

        # ۲. وقتی کاربر شماره موبایل را فرستاد
        if contact and "phone_number" in contact:
            user_phone = clean_phone(contact["phone_number"])
            
            with FileLock(LOCK_PATH):
                users_df = read_sheet(USERS_SHEET)
                
                # استانداردسازی ستون phone دیتابیس برای مقایسه
                if "phone" not in users_df.columns:
                    users_df["phone"] = ""
                if "bale_chat_id" not in users_df.columns:
                    users_df["bale_chat_id"] = ""

                users_df["clean_phone"] = users_df["phone"].astype(str).apply(clean_phone)
                
                # پیدا کردن کاربر با تطابق شماره موبایل
                matched_user = users_df[users_df["clean_phone"] == user_phone]

                if not matched_user.empty:
                    idx = matched_user.index[0]
                    u_code = users_df.loc[idx, "user_code"]
                    u_name = f"{users_df.loc[idx, 'name']} {users_df.loc[idx, 'lastname']}"

                    # ثبت chat_id کاربر در دیتابیس اکسل
                    users_df.loc[idx, "bale_chat_id"] = str(chat_id)
                    users_df = users_df.drop(columns=["clean_phone"])

                    with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                        users_df.to_excel(writer, sheet_name=USERS_SHEET, index=False)

                    # بستن کیبورد و ارسال پیام تبریک
                    remove_keyboard = {"remove_keyboard": True}
                    send_bale_message(
                        chat_id,
                        f"✅ هویت شما تایید شد!\n"
                        f"همکار گرامی: {u_name}\n"
                        f"کد پرسنلی: {u_code}\n\n"
                        f"از این لحظه کلیه تغییر وضعیت‌های خودروها به صورت آنی برای شما ارسال خواهد شد.",
                        reply_markup=remove_keyboard
                    )
                else:
                    send_bale_message(
                        chat_id,
                        f"❌ شماره موبایل شما ({user_phone}) در لیست کارشناسان و پرسنل نمایشگاه تعریف نشده است.\n"
                        f"لطفاً به مدیریت اطلاع دهید تا شماره شما را ثبت کند."
                    )
            return {"ok": True}

    except Exception as e:
        print(f"Webhook processing error: {e}")

    return {"ok": True}

def broadcast_car_update(car_code: str, title: str, from_status: str, to_status: str, from_place: str, to_place: str, changed_by: str, changed_at: str):
    """ارسال پیام به پی‌وی تمام کارشناسانی که بله خود را متصل کرده‌اند"""
    try:
        users_df = read_sheet(USERS_SHEET)
        if "bale_chat_id" not in users_df.columns:
            return

        # گرفتن تمام چت‌آیدی‌های معتبر
        active_chat_ids = users_df["bale_chat_id"].dropna().astype(str).tolist()
        active_chat_ids = [cid.strip() for cid in active_chat_ids if cid.strip() and cid.strip() != "nan"]

        if not active_chat_ids:
            return

        status_emoji = {
            "free": "🟢 آزاد شد",
            "reserved(4 hours)": "⏳ رزرو ۴ ساعته",
            "بیعانه": "🟡 بیعانه دریافت شد",
            "فروخته شده": "🔴 فروخته شد"
        }.get(to_status, to_status)

        message = (
            f"🚗 **تغییر وضعیت خودرو**\n\n"
            f"▫️ **کد خودرو:** {car_code}\n"
            f"▫️ **خودرو:** {title}\n"
            f"▫️ **وضعیت جدید:** {status_emoji}\n"
            f"▫️ **محل:** {to_place}\n"
            f"▫️ **وضعیت قبلی:** {from_status} ({from_place})\n"
            f"👤 **ثبت کننده:** {changed_by}\n"
            f"🕒 **زمان:** {changed_at}"
        )

        for cid in set(active_chat_ids):
            send_bale_message(cid, message)

    except Exception as e:
        print(f"Error in broadcast_car_update: {e}")
        
def expire_reserved_if_needed(system_user_label: str = "SYSTEM"):
    cars_df = read_sheet(CARS_SHEET)
    if cars_df.empty:
        return

    if "reserved_until" not in cars_df.columns:
        return

    now = datetime.now()

    # فقط reserved(4 hours) هایی که reserved_until دارند و گذشته
    reserved_mask = cars_df["status"].astype(str) == "reserved(4 hours)"
    ru = pd.to_datetime(cars_df["reserved_until"], errors="coerce")

    expired_idx = cars_df.index[reserved_mask & ru.notna() & (ru < now)].tolist()
    if not expired_idx:
        return

    for i in expired_idx:
        car_code = str(cars_df.at[i, "car_code"])
        old_status = str(cars_df.at[i, "status"])
        old_place = "" if pd.isna(cars_df.at[i, "place"]) else str(cars_df.at[i, "place"])

        cars_df.at[i, "status"] = "free"
        cars_df.at[i, "reserved_until"] = ""
        cars_df.at[i, "last_updated_by"] = system_user_label
        cars_df.at[i, "last_updated_at"] = now.strftime("%Y-%m-%d %H:%M:%S")

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
# Auth helpers (cookie session)
# -----------------------------
def append_history(car_code: str, changed_by: str, from_status: str, to_status: str, from_place: str, to_place: str):
    history_df = read_sheet(HISTORY_SHEET)

    next_id = 1
    if not history_df.empty and "id" in history_df.columns:
        try:
            next_id = int(pd.to_numeric(history_df["id"], errors="coerce").max() or 0) + 1
        except Exception:
            next_id = len(history_df) + 1

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row = {
        "id": next_id,
        "car_code": str(car_code),
        "changed_at": now,
        "changed_by": changed_by,
        "from_status": from_status,
        "to_status": to_status,
        "from_place": from_place,
        "to_place": to_place,
    }
    history_df = pd.concat([history_df, pd.DataFrame([row])], ignore_index=True)

    # ذخیره سه شیت با هم
    users_df = read_sheet(USERS_SHEET)
    cars_df = read_sheet(CARS_SHEET)
    with FileLock(LOCK_PATH):
        with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
            users_df.to_excel(writer, sheet_name=USERS_SHEET, index=False)
            cars_df.to_excel(writer, sheet_name=CARS_SHEET, index=False)
            history_df.to_excel(writer, sheet_name=HISTORY_SHEET, index=False)


def set_session(response: RedirectResponse, user: dict):
    token = serializer.dumps({
        "username": user["username"],
        "role": user["role"],
        "user_code": user["user_code"],
        "name": user["name"],
        "lastname": user["lastname"],
    })
    response.set_cookie("session", token, httponly=True, samesite="lax")

def clear_session(response: RedirectResponse):
    response.delete_cookie("session")

def get_current_user(request: Request) -> Optional[dict]:
    token = request.cookies.get("session")
    if not token:
        return None
    try:
        return serializer.loads(token)
    except BadSignature:
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

# -----------------------------
# UI pages
# -----------------------------
@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    user = get_current_user(request)
    if user:
        return RedirectResponse("/cars", status_code=302)
    n1, n2, token = generate_captcha()
    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={"error": None, "num1": n1, "num2": n2, "captcha_token": token}
    )

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    n1, n2, token = generate_captcha()
    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={"error": None, "num1": n1, "num2": n2, "captcha_token": token}
    )

@app.post("/login")
async def login_post(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    captcha_answer: str = Form(...),
    captcha_token: str = Form(...)
):
    # ۱. بررسی صحت حاصل‌جمع کپچا
    captcha_valid = False
    try:
        expected_sum = captcha_serializer.loads(captcha_token)
        if int(str(captcha_answer).strip()) == expected_sum:
            captcha_valid = True
    except (BadSignature, ValueError, TypeError):
        captcha_valid = False

    if not captcha_valid:
        n1, n2, token = generate_captcha()
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "error": "پاسخ سوال جمع ریاضی اشتباه است. لطفاً دوباره تلاش کنید.",
                "num1": n1,
                "num2": n2,
                "captcha_token": token
            }
        )

    # ۲. بررسی نام کاربری و پسورد
    users_df = read_sheet(USERS_SHEET)
    username = str(username).strip()
    password = str(password)

    if users_df.empty or username not in users_df["username"].astype(str).str.strip().values:
        n1, n2, token = generate_captcha()
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "error": "نام کاربری یا رمز اشتباه است.",
                "num1": n1,
                "num2": n2,
                "captcha_token": token
            }
        )
    row = users_df[users_df["username"].astype(str).str.strip() == username].iloc[0].to_dict()
    saved_password = "" if pd.isna(row.get("password")) else str(row.get("password"))

    if saved_password != password:
        n1, n2, token = generate_captcha()
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "error": "نام کاربری یا رمز اشتباه است.",
                "num1": n1,
                "num2": n2,
                "captcha_token": token
            }
        )

    resp = RedirectResponse("/cars", status_code=303)
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
            cars_df.get("plate", pd.Series(dtype=str)).astype(str).str.lower().str.contains(q_lower, na=False) |
            cars_df.get("owner", pd.Series(dtype=str)).astype(str).str.lower().str.contains(q_lower, na=False)
        )
        cars_df = cars_df[mask]

    if status:
        cars_df = cars_df[cars_df["status"].astype(str) == str(status)]

    if place:
        cars_df = cars_df[cars_df["place"].astype(str) == str(place)]

    all_cars_df = read_sheet(CARS_SHEET)
    status_options = sorted(set(all_cars_df["status"].dropna().astype(str).tolist())) if "status" in all_cars_df.columns else []
    place_options = sorted(set(all_cars_df["place"].dropna().astype(str).tolist())) if "place" in all_cars_df.columns else []

    can_add = user["role"] in ROLES_CAN_ADD_CAR

    cars = cars_df.fillna("").to_dict(orient="records")
    return templates.TemplateResponse(
        name="cars.html",
        context={
            "request": request,
            "user": user,
            "cars": cars,
            "q": q or "",
            "status": status or "",
            "place": place or "",
            "status_options": status_options,
            "place_options": place_options,
            "can_add": can_add,
        },
        request=request,
    )

@app.post("/cars/add")
async def add_car(
    request: Request,
    car_code: str = Form(...),
    title: str = Form(""),
    specs: str = Form(""),
    color: str = Form(""),
    status: str = Form("free"),
    model_year: str = Form(""),
    production_year: str = Form(""),
    owner: str = Form(""),
    place: str = Form(""),
    user=Depends(require_role("admin", "showroom_manager"))
):
    user = get_current_user(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    if not user_has_bale(request):
        return RedirectResponse(url="/activate-bale", status_code=303)
            
    cars_df = read_sheet(CARS_SHEET)

    if car_code in cars_df["car_code"].astype(str).values:
        raise HTTPException(status_code=400, detail="این کد ماشین قبلاً ثبت شده است.")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if status not in ALLOWED_STATUSES:
        raise HTTPException(status_code=400, detail="وضعیت نامعتبر است.")

    reserved_until = ""
    if status == "reserved(4 hours)":
        reserved_until = (datetime.now() + pd.Timedelta(hours=4)).strftime("%Y-%m-%d %H:%M:%S")
    
    new_row = {
        "car_code": str(car_code),
        "title": title,
        "specs": specs,
        "color": color,
        "status": status,
        "reserved_until": reserved_until,
        "model_year": model_year,
        "production_year": production_year,
        "owner": owner,
        "place": place,
        "last_updated_by": f'{user["name"]} {user["lastname"]} ({user["user_code"]})',
        "last_updated_at": now,
    }

    cars_df = pd.concat([cars_df, pd.DataFrame([new_row])], ignore_index=True)
    save_cars(cars_df)
    append_history(
        car_code=str(car_code),
        changed_by=f'{user["name"]} {user["lastname"]} ({user["user_code"]})',
        from_status="",
        to_status=status,
        from_place="",
        to_place=place,
    )
    return RedirectResponse("/cars", status_code=302)

@app.get("/cars/history/{car_code}")
def car_history(car_code: str, user=Depends(require_user)):
    history_df = read_sheet(HISTORY_SHEET)
    if history_df.empty:
        return JSONResponse({"car_code": car_code, "items": []})

    df = history_df[history_df["car_code"].astype(str) == str(car_code)].copy()
    df = df.sort_values(by="id", ascending=False) if "id" in df.columns else df
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
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # این متغیرها را قبل از بلاک تعریف می‌کنیم تا بیرون از lock هم در دسترس باشند
    title = ""
    from_status = ""
    from_place = ""
    final_place = place

    with FileLock(LOCK_PATH):
        try:
            cars_df = read_sheet(CARS_SHEET)
            
            # تبدیل car_code به رشته برای تطابق دقیق
            cars_df["car_code"] = cars_df["car_code"].astype(str)
            car_code = str(car_code).strip()

            row_idx = cars_df[cars_df["car_code"] == car_code].index

            if not row_idx.empty:
                idx = row_idx[0]
                
                # خواندن مقادیر فعلی خودرو از اکسل
                title = str(cars_df.loc[idx, "title"]) if "title" in cars_df.columns else "نامشخص"
                from_status = str(cars_df.loc[idx, "status"]) if "status" in cars_df.columns else "نامشخص"
                from_place = str(cars_df.loc[idx, "place"]) if "place" in cars_df.columns else ""

                if not final_place:
                    final_place = from_place

                # آپدیت مشخصات خودرو
                cars_df.loc[idx, "status"] = status
                if place:
                    cars_df.loc[idx, "place"] = place
                cars_df.loc[idx, "last_updated_at"] = now_str
                cars_df.loc[idx, "last_updated_by"] = user_fullname

                # ذخیره در شیت cars
                with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    cars_df.to_excel(writer, sheet_name=CARS_SHEET, index=False)

                # ثبت در لاگ (شیت car_logs)
                logs_df = read_sheet("car_logs")
                new_log = {
                    "id": len(logs_df) + 1,
                    "car_code": car_code,
                    "changed_at": now_str,
                    "changed_by": user_fullname,
                    "from_status": from_status,
                    "to_status": status,
                    "from_place": from_place,
                    "to_place": final_place
                }
                logs_df = pd.concat([logs_df, pd.DataFrame([new_log])], ignore_index=True)
                with pd.ExcelWriter(DB_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    logs_df.to_excel(writer, sheet_name="car_logs", index=False)

        except Exception as e:
            print(f"Error in update_status: {e}")
            raise e

    # ارسال اعلان بله (حالا title مقدار دارد و ارور نمی‌دهد)
    try:
        broadcast_car_update(
            car_code=car_code,
            title=title,
            from_status=from_status,
            to_status=status,
            from_place=from_place,
            to_place=final_place,
            changed_by=user_fullname,
            changed_at=now_str
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
        "password": str(password),  # متن ساده ذخیره می‌شود
        "name": name,
        "lastname": lastname,
        "role": role,
    }
    users_df = pd.concat([users_df, pd.DataFrame([row])], ignore_index=True)
    save_users(users_df)
    return JSONResponse({"ok": True})

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "80"))
    uvicorn.run(app, host="0.0.0.0", port=port)
