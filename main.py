import os
import random

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

serializer = URLSafeSerializer(APP_SECRET, salt="session")
captcha_serializer = URLSafeSerializer(APP_SECRET, salt="captcha-salt")

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
HISTORY_SHEET = "history"

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
    ensure_db()
    with FileLock(LOCK_PATH):
        return pd.read_excel(DB_PATH, sheet_name=sheet_name, engine="openpyxl")

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
def login(
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

    resp = RedirectResponse("/cars", status_code=302)
    set_session(resp, row)
    return resp

@app.get("/logout")
def logout():
    resp = RedirectResponse("/login", status_code=302)
    clear_session(resp)
    return resp

@app.get("/cars", response_class=HTMLResponse)
def cars_page(
    expire_reserved_if_needed(system_user_label="SYSTEM (auto)")
    cars_df = read_sheet(CARS_SHEET)
    request: Request,
    q: Optional[str] = None,
    status: Optional[str] = None,
    place: Optional[str] = None,
    user=Depends(require_user),
):
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
def add_car(
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
def update_status(
    request: Request,
    car_code: str = Form(...),
    status: str = Form(...),
    place: str = Form(""),
    user=Depends(require_user)
):
    cars_df = read_sheet(CARS_SHEET)
    idx = cars_df.index[cars_df["car_code"].astype(str) == str(car_code)].tolist()
    if not idx:
        raise HTTPException(status_code=404, detail="ماشین پیدا نشد.")
    i = idx[0]

    if status not in ALLOWED_STATUSES:
        raise HTTPException(status_code=400, detail="وضعیت نامعتبر است.")

    old_status = "" if pd.isna(cars_df.at[i, "status"]) else str(cars_df.at[i, "status"])
    old_place  = "" if pd.isna(cars_df.at[i, "place"])  else str(cars_df.at[i, "place"])
    
    cars_df.at[i, "status"] = status
    
    new_place = old_place
    if place != "":
        cars_df.at[i, "place"] = place
        new_place = place
    
    # reserved_until logic
    if status == "reserved(4 hours)":
        cars_df.at[i, "reserved_until"] = (datetime.now() + pd.Timedelta(hours=4)).strftime("%Y-%m-%d %H:%M:%S")
    else:
        cars_df.at[i, "reserved_until"] = ""
    

    if place != "":
        cars_df.at[i, "place"] = place

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cars_df.at[i, "last_updated_by"] = f'{user["name"]} {user["lastname"]} ({user["user_code"]})'
    cars_df.at[i, "last_updated_at"] = now
    append_history(
        car_code=str(car_code),
        changed_by=f'{user["name"]} {user["lastname"]} ({user["user_code"]})',
        from_status=old_status,
        to_status=status,
        from_place=old_place,
        to_place=new_place,
    )
    save_cars(cars_df)
    return RedirectResponse("/cars", status_code=302)

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
