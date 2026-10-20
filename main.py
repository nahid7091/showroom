import os
from datetime import datetime
from typing import Optional

import pandas as pd
from fastapi import FastAPI, Request, Form, Depends, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from filelock import FileLock
from passlib.context import CryptContext
from itsdangerous import URLSafeSerializer, BadSignature

# -----------------------------
# Config
# -----------------------------
APP_SECRET = os.getenv("APP_SECRET", "CHANGE_ME_SUPER_SECRET")
DB_PATH = os.getenv("EXCEL_DB_PATH", "/data/showroom/database.xlsx")
LOCK_PATH = DB_PATH + ".lock"

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
serializer = URLSafeSerializer(APP_SECRET, salt="session")

ROLES_CAN_ADD_CAR = {"admin", "showroom_manager"}
ROLES_CAN_DELETE_CAR = {"admin", "showroom_manager"}  # شما گفتی فقط این دو نفر
# NOTE: فعلاً API حذف نداریم، ولی اگر اضافه شد از این استفاده کن.

USERS_SHEET = "users"
CARS_SHEET = "cars"

REQUIRED_USER_COLS = ["user_code", "username", "password", "name", "lastname", "role"]
REQUIRED_CAR_COLS = [
    "car_code", "VIN", "title", "model_year","production_year", "color", "owner", "place","last_updated_at", "last_updated_by", "status"
]

# -----------------------------
# App
# -----------------------------
app = FastAPI()
templates = Jinja2Templates(directory="templates")

# اگر خواستی فایل‌های css/js بذاری:
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
        with pd.ExcelWriter(DB_PATH, engine="openpyxl") as writer:
            users_df.to_excel(writer, sheet_name=USERS_SHEET, index=False)
            cars_df.to_excel(writer, sheet_name=CARS_SHEET, index=False)

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


# -----------------------------
# Auth helpers (cookie session)
# -----------------------------
def set_session(response: RedirectResponse, user: dict):
    token = serializer.dumps({"username": user["username"], "role": user["role"], "user_code": user["user_code"],
                              "name": user["name"], "lastname": user["lastname"]})
    response.set_cookie("session", token, httponly=True, samesite="lax")

def clear_session(response: RedirectResponse):
    response.delete_cookie("session")

def get_current_user(request: Request) -> Optional[dict]:
    token = request.cookies.get("session")
    if not token:
        return None
    try:
        data = serializer.loads(token)
        return data
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
# UI pages (simple HTML)
# -----------------------------
@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
  # مطمئن شوید فایلی به اسم login.html یا cars.html داخل پوشه templates دارید
  return templates.TemplateResponse("login.html", {"request": request})
    
def home(request: Request):
    user = get_current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=302)
    return RedirectResponse("/cars", status_code=302)

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    return templates.TemplateResponse("login.html", {"request": request, "error": None})

@app.post("/login")
def login(request: Request, username: str = Form(...), password: str = Form(...)):
    users_df = read_sheet(USERS_SHEET)
    if users_df.empty or username not in users_df["username"].astype(str).values:
        return templates.TemplateResponse("login.html", {"request": request, "error": "نام کاربری یا رمز اشتباه است."})

    row = users_df[users_df["username"].astype(str) == str(username)].iloc[0].to_dict()
    if not pwd_context.verify(password, str(row["password"])):
        return templates.TemplateResponse("login.html", {"request": request, "error": "نام کاربری یا رمز اشتباه است."})

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
    request: Request,
    q: Optional[str] = None,
    status: Optional[str] = None,
    place: Optional[str] = None,
    user=Depends(require_user),
):
    cars_df = read_sheet(CARS_SHEET)

    # فیلترها
    if q:
        q_lower = q.lower()
        # سرچ روی چند ستون مهم
        for col in ["car_code", "title", "plate", "owner", "color"]:
            if col not in cars_df.columns:
                continue
        mask = (
            cars_df["car_code"].astype(str).str.lower().str.contains(q_lower, na=False) |
            cars_df["title"].astype(str).str.lower().str.contains(q_lower, na=False) |
            cars_df["plate"].astype(str).str.lower().str.contains(q_lower, na=False) |
            cars_df["owner"].astype(str).str.lower().str.contains(q_lower, na=False)
        )
        cars_df = cars_df[mask]

    if status:
        cars_df = cars_df[cars_df["status"].astype(str) == str(status)]

    if place:
        cars_df = cars_df[cars_df["place"].astype(str) == str(place)]

    # برای dropdownها
    all_cars_df = read_sheet(CARS_SHEET)
    status_options = sorted(set(all_cars_df["status"].dropna().astype(str).tolist()))
    place_options = sorted(set(all_cars_df["place"].dropna().astype(str).tolist()))

    can_add = user["role"] in ROLES_CAN_ADD_CAR

    # نمایش
    cars = cars_df.fillna("").to_dict(orient="records")
    return templates.TemplateResponse("cars.html", {
        "request": request,
        "user": user,
        "cars": cars,
        "q": q or "",
        "status": status or "",
        "place": place or "",
        "status_options": status_options,
        "place_options": place_options,
        "can_add": can_add,
    })


@app.post("/cars/add")
def add_car(
    request: Request,
    car_code: str = Form(...),
    title: str = Form(""),
    color: str = Form(""),
    status: str = Form("موجود"),
    model_year: str = Form(""),
    production_year: str = Form(""),
    agency_end_date: str = Form(""),
    owner: str = Form(""),
    plate: str = Form(""),
    place: str = Form(""),
    user=Depends(require_role("admin", "showroom_manager"))
):
    cars_df = read_sheet(CARS_SHEET)

    if car_code in cars_df["car_code"].astype(str).values:
        raise HTTPException(status_code=400, detail="این کد ماشین قبلاً ثبت شده است.")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_row = {
        "car_code": str(car_code),
        "title": title,
        "color": color,
        "status": status,
        "model_year": model_year,
        "production_year": production_year,
        "agency_end_date": agency_end_date,
        "owner": owner,
        "plate": plate,
        "last_updated_by": f'{user["name"]} {user["lastname"]} ({user["user_code"]})',
        "last_updated_at": now,
        "place": place,
    }

    cars_df = pd.concat([cars_df, pd.DataFrame([new_row])], ignore_index=True)
    save_cars(cars_df)
    return RedirectResponse("/cars", status_code=302)


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

    cars_df.at[i, "status"] = status
    if place != "":
        cars_df.at[i, "place"] = place

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cars_df.at[i, "last_updated_by"] = f'{user["name"]} {user["lastname"]} ({user["user_code"]})'
    cars_df.at[i, "last_updated_at"] = now

    save_cars(cars_df)
    return RedirectResponse("/cars", status_code=302)


# -----------------------------
# Admin helper: create users (optional endpoint)
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

    if username in users_df["username"].astype(str).values:
        raise HTTPException(status_code=400, detail="این username قبلاً ثبت شده است.")
    if user_code in users_df["user_code"].astype(str).values:
        raise HTTPException(status_code=400, detail="این کد کاربر قبلاً ثبت شده است.")

    row = {
        "user_code": str(user_code),
        "username": str(username),
        #"password": pwd_context.hash(password),
        "password": str(password),
        "name": name,
        "lastname": lastname,
        "role": role,
    }
    users_df = pd.concat([users_df, pd.DataFrame([row])], ignore_index=True)
    save_users(users_df)
    return JSONResponse({"ok": True})
    
if __name__ == "__main__":
    import os
    import uvicorn

    port = int(os.environ.get("PORT", "80"))
    uvicorn.run(app, host="0.0.0.0", port=port)
