# -*- coding: utf-8 -*-
"""
سیستم مدیریت اطلاعات پرسنل — نسخه وب با FastAPI
چندکاربره | تقویم شمسی | عکس پرسنلی | ایمپورت/اکسپورت اکسل | مدیریت ردیف‌های نامعتبر ایمپورت
"""
import io
import sqlite3
import zipfile
import json
import os
import secrets
import shutil
import time
from datetime import datetime, date, timedelta
from typing import Optional
import re

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image
import jdatetime

from database import (
    ALL_PERMISSION_KEYS,
    DATE_FIELDS,
    NORMALIZED_MAPPING,
    PERMISSION_DEFS,
    PERSONAL_FIELDS,
    DB_PATH,
    PHOTOS_DIR,
    ROLE_ADMIN,
    ROLE_CUSTOM,
    ROLE_USER,
    VALID_ROLES,
    WORK_FIELDS,
    create_session,
    delete_session,
    get_db,
    get_user_by_token,
    hash_password,
    init_db,
    log_access,
    normalize_for_search,
    normalize_text,
    normalize_active_status,
    parse_permissions,
    serialize_permissions,
    sql_normalize_col,
    user_has_permission,
    user_to_public,
    verify_password,
)
from card_generator import (
    generate_card_pdf,
    generate_addiction_letter_pdf,
    generate_group_cards_pdf,
    resolve_issue_date,
    resolve_validity_date,
)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS_DIR = os.path.join(APP_DIR, "downloads_tmp")
DOCS_DIR = os.path.join(APP_DIR, "docs")  # پوشه مدارک پرسنل (بر اساس شماره اندیکاتور)
DOWNLOAD_TTL_SEC = 15 * 60  # لینک موقت ۱۵ دقیقه معتبر است

app = FastAPI(title="سیستم مدیریت پرسنل", version="2.3")
app.add_middleware(GZipMiddleware, minimum_size=500)
app.mount("/static", StaticFiles(directory=os.path.join(APP_DIR, "static")), name="static")
app.mount("/photos", StaticFiles(directory=PHOTOS_DIR), name="photos")
templates = Jinja2Templates(directory=os.path.join(APP_DIR, "templates"))


@app.middleware("http")
async def audit_log_middleware(request: Request, call_next):
    """ثبت خودکار همه عملیات موفق کاربر در لاگ دسترسی (با تاریخ و ساعت)."""
    response = await call_next(request)
    try:
        _audit_log_request(request, response)
    except Exception:
        pass
    return response


def _ensure_downloads_dir():
    os.makedirs(DOWNLOADS_DIR, exist_ok=True)


def _cleanup_old_downloads(max_age_sec: int = DOWNLOAD_TTL_SEC):
    """حذف فایل‌های موقت منقضی‌شده"""
    try:
        _ensure_downloads_dir()
        now = time.time()
        for name in os.listdir(DOWNLOADS_DIR):
            path = os.path.join(DOWNLOADS_DIR, name)
            try:
                if os.path.isfile(path) and (now - os.path.getmtime(path)) > max_age_sec:
                    os.remove(path)
            except OSError:
                pass
    except OSError:
        pass


def create_temp_download(data: bytes, filename: str, media_type: str = "application/pdf") -> dict:
    """ذخیره فایل موقت و برگرداندن لینک دانلود یک‌بارمصرف/کوتاه‌مدت.

    فرانت فقط این لینک را باز می‌کند تا مرورگر (و در صورت نیاز IDM)
    بدون هدر احراز هویت فایل را مستقیم دانلود کند.
    """
    _ensure_downloads_dir()
    _cleanup_old_downloads()
    token = secrets.token_urlsafe(24)
    # نام فایل روی دیسک فقط توکن؛ نام واقعی در متادیتا
    safe_name = re.sub(r"[^A-Za-z0-9_.\-]+", "_", filename) or "file.bin"
    meta = {
        "filename": safe_name,
        "media_type": media_type,
        "created": time.time(),
    }
    bin_path = os.path.join(DOWNLOADS_DIR, f"{token}.bin")
    meta_path = os.path.join(DOWNLOADS_DIR, f"{token}.json")
    with open(bin_path, "wb") as f:
        f.write(data)
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)
    return {
        "token": token,
        "download_url": f"/api/downloads/{token}",
        "filename": safe_name,
        "size": len(data),
        "expires_in": DOWNLOAD_TTL_SEC,
    }


def create_temp_download_from_path(
    src_path: str,
    filename: str = None,
    media_type: str = None,
    disposition: str = "attachment",
) -> dict:
    """کپی فایل موجود به پوشه موقت و لینک دانلود بدون نیاز به هدر احراز هویت.

    disposition: attachment (دانلود) یا inline (باز کردن در مرورگر)
    برای مدارک docs و فایل‌های بزرگ (بدون بارگذاری کامل در حافظه).
    """
    if not os.path.isfile(src_path):
        raise HTTPException(404, "فایل یافت نشد")
    _ensure_downloads_dir()
    _cleanup_old_downloads()
    token = secrets.token_urlsafe(24)
    raw_name = filename or os.path.basename(src_path)
    display_name = os.path.basename(str(raw_name).replace("\\", "/").replace("/", "_")).strip()
    display_name = display_name.replace("\0", "").replace("\r", "").replace("\n", "")
    if not display_name or display_name in (".", ".."):
        display_name = "file.bin"
    ascii_name = re.sub(r"[^A-Za-z0-9_.\-]+", "_", display_name).strip("._") or "file.bin"
    if not media_type:
        ext = os.path.splitext(display_name)[1].lower()
        media_type = {
            ".pdf": "application/pdf",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".gif": "image/gif",
            ".webp": "image/webp",
            ".txt": "text/plain; charset=utf-8",
            ".csv": "text/csv; charset=utf-8",
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xls": "application/vnd.ms-excel",
            ".doc": "application/msword",
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".zip": "application/zip",
        }.get(ext, "application/octet-stream")
    disp = "inline" if str(disposition).lower() == "inline" else "attachment"
    meta = {
        "filename": display_name,
        "ascii_filename": ascii_name,
        "media_type": media_type,
        "disposition": disp,
        "created": time.time(),
    }
    bin_path = os.path.join(DOWNLOADS_DIR, f"{token}.bin")
    meta_path = os.path.join(DOWNLOADS_DIR, f"{token}.json")
    shutil.copy2(src_path, bin_path)
    size = os.path.getsize(bin_path)
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)
    return {
        "token": token,
        "download_url": f"/api/downloads/{token}",
        "filename": display_name,
        "size": size,
        "expires_in": DOWNLOAD_TTL_SEC,
    }


@app.on_event("startup")
def on_startup():
    init_db()
    _ensure_downloads_dir()
    os.makedirs(DOCS_DIR, exist_ok=True)
    _cleanup_old_downloads()


@app.get("/api/downloads/{token}")
def download_temp_file(token: str):
    """دانلود فایل موقت با لینک — بدون نیاز به هدر احراز هویت.

    توکن تصادفی و کوتاه‌مدت است؛ پس از دانلود یا انقضا حذف می‌شود.
    """
    # فقط کاراکترهای امن
    if not re.fullmatch(r"[A-Za-z0-9_\-]{16,64}", token or ""):
        raise HTTPException(404, "لینک نامعتبر است")
    bin_path = os.path.join(DOWNLOADS_DIR, f"{token}.bin")
    meta_path = os.path.join(DOWNLOADS_DIR, f"{token}.json")
    if not os.path.isfile(bin_path) or not os.path.isfile(meta_path):
        raise HTTPException(404, "لینک منقضی شده یا یافت نشد")
    try:
        with open(meta_path, "r", encoding="utf-8") as f:
            meta = json.load(f)
    except Exception:
        raise HTTPException(404, "لینک نامعتبر است")
    created = float(meta.get("created") or 0)
    if time.time() - created > DOWNLOAD_TTL_SEC:
        try:
            os.remove(bin_path)
            os.remove(meta_path)
        except OSError:
            pass
        raise HTTPException(410, "لینک منقضی شده است")
    filename = meta.get("filename") or f"{token}.bin"
    ascii_filename = meta.get("ascii_filename") or re.sub(r"[^A-Za-z0-9_.\-]+", "_", str(filename)) or "file.bin"
    media_type = meta.get("media_type") or "application/octet-stream"
    size = os.path.getsize(bin_path)

    from urllib.parse import quote
    disp = "inline" if str(meta.get("disposition") or "").lower() == "inline" else "attachment"
    cd = "%s; filename=\"%s\"; filename*=UTF-8''%s" % (disp, ascii_filename, quote(str(filename)))

    # فایل تا انقضای TTL روی دیسک می‌ماند (برای سازگاری با دانلود‌منیجر که ممکن است چندبار درخواست بزند)
    return FileResponse(
        bin_path,
        media_type=media_type,
        filename=ascii_filename,
        headers={
            "Content-Disposition": cd,
            "Content-Length": str(size),
            "Cache-Control": "no-store",
            "X-Content-Type-Options": "nosniff",
        },
    )


# ──────────────────────────────────────────────
# احراز هویت
# ──────────────────────────────────────────────
def _extract_token(authorization: Optional[str] = None, x_token: Optional[str] = None) -> str:
    if x_token and x_token.strip():
        return x_token.strip()
    if authorization and authorization.lower().startswith("bearer "):
        return authorization[7:].strip()
    return ""


def get_current_user(
    authorization: Optional[str] = Header(None),
    x_auth_token: Optional[str] = Header(None, alias="X-Auth-Token"),
):
    token = _extract_token(authorization, x_auth_token)
    if not token:
        raise HTTPException(401, "لطفاً وارد سیستم شوید")
    with get_db() as conn:
        user = get_user_by_token(conn, token)
    if not user:
        raise HTTPException(401, "نشست منقضی شده یا نامعتبر است — دوباره وارد شوید")
    return user


def get_optional_user(
    authorization: Optional[str] = Header(None),
    x_auth_token: Optional[str] = Header(None, alias="X-Auth-Token"),
):
    token = _extract_token(authorization, x_auth_token)
    if not token:
        return None
    with get_db() as conn:
        return get_user_by_token(conn, token)


def require_admin(user=Depends(get_current_user)):
    if user.get("role") != ROLE_ADMIN:
        raise HTTPException(403, "فقط مدیر سیستم به این بخش دسترسی دارد")
    return user


def require_write(user=Depends(get_current_user)):
    """دسترسی نوشتن کلی — مدیر یا سفارشی با حداقل یکی از دسترسی‌های نوشتن"""
    if user.get("role") == ROLE_ADMIN:
        return user
    if user.get("role") == ROLE_CUSTOM:
        write_perms = (
            "personnel_create", "personnel_edit", "personnel_delete",
            "photo_manage", "work_write", "import_excel", "invalid_imports",
        )
        if any(user_has_permission(user, p) for p in write_perms):
            return user
    raise HTTPException(
        403,
        "حساب شما فقط خواندنی است. ذخیره و تغییر اطلاعات فقط توسط مدیر یا کاربر دارای دسترسی انجام می‌شود.",
    )
    return user


def require_perm(perm: str):
    """Dependency factory برای یک دسترسی مشخص"""
    def _checker(user=Depends(get_current_user)):
        if not user_has_permission(user, perm):
            raise HTTPException(403, "شما به این بخش دسترسی ندارید")
        return user
    return _checker


def _client_ip(request: Request) -> str:
    forwarded = (request.headers.get("x-forwarded-for") or "").split(",")[0].strip()
    if forwarded:
        return forwarded
    if request.client:
        return request.client.host or ""
    return ""


def _client_ua(request: Request) -> str:
    return request.headers.get("user-agent") or ""


def _map_api_action(method: str, path: str) -> tuple:
    """برگرداندن (کد رویداد، توضیح فارسی) بر اساس مسیر API."""
    m = (method or "").upper()
    p = path or ""

    rules = [
        (r"^/api/users/?$", "POST", "user_create", "ایجاد کاربر"),
        (r"^/api/users/\d+$", "PUT", "user_update", "ویرایش کاربر"),
        (r"^/api/users/\d+$", "DELETE", "user_delete", "حذف کاربر"),
        (r"^/api/personnel/?$", "POST", "personnel_create", "ایجاد پرسنل"),
        (r"^/api/personnel/\d+$", "PUT", "personnel_update", "ویرایش پرسنل"),
        (r"^/api/personnel/\d+$", "DELETE", "personnel_delete", "حذف پرسنل"),
        (r"^/api/personnel/\d+/photo$", "POST", "photo_upload", "آپلود عکس"),
        (r"^/api/personnel/\d+/photo$", "DELETE", "photo_delete", "حذف عکس"),
        (r"^/api/personnel/\d+/work$", "POST", "work_create", "افزودن سابقه کاری"),
        (r"^/api/work/\d+$", "PUT", "work_update", "ویرایش سابقه کاری"),
        (r"^/api/work/\d+$", "DELETE", "work_delete", "حذف سابقه کاری"),
        (r"^/api/work/\d+/promote$", "POST", "work_promote", "ارتقا سابقه به رکورد اصلی"),
        (r"^/api/personnel/\d+/card$", "GET", "card_issue", "صدور کارت تکی"),
        (r"^/api/personnel/cards/group$", "POST", "card_issue", "صدور کارت گروهی"),
        (r"^/api/personnel/\d+/card/renew$", "POST", "card_renew", "تمدید کارت موقت"),
        (r"^/api/personnel/\d+/card-issued/reset$", "POST", "card_reset", "بازنشانی وضعیت صدور کارت"),
        (r"^/api/import/excel$", "POST", "import_excel", "ایمپورت اکسل"),
        (r"^/api/export/excel$", "GET", "export_excel", "خروجی اکسل"),
        (r"^/api/export/excel/filtered$", "POST", "export_excel", "خروجی اکسل فیلترشده"),
        (r"^/api/invalid-imports/\d+/resolve$", "POST", "invalid_resolve", "رفع ردیف نامعتبر"),
        (r"^/api/invalid-imports/\d+$", "DELETE", "invalid_delete", "حذف ردیف نامعتبر"),
        (r"^/api/invalid-imports$", "DELETE", "invalid_clear", "پاک‌سازی ردیف‌های نامعتبر"),
        (r"^/api/admin/merge-histories$", "POST", "merge_histories", "ادغام سوابق"),
        (r"^/api/access-logs$", "DELETE", "logs_clear", "پاک‌سازی لاگ دسترسی"),
    ]
    for pattern, meth, action, label in rules:
        if m == meth and re.match(pattern, p):
            return action, label
    # رویداد عمومی برای سایر APIهای تغییردهنده
    if m in ("POST", "PUT", "DELETE", "PATCH"):
        return "api_action", f"{m} {p}"
    return "", ""


def _person_label_from_row(row) -> str:
    """نام کامل + کد ملی از رکورد پرسنل."""
    if not row:
        return ""
    try:
        first = (row["first_name"] or "").strip()
        last = (row["last_name"] or "").strip()
        nid = (row["national_id"] or "").strip()
    except Exception:
        return ""
    name = f"{first} {last}".strip()
    parts = []
    if name:
        parts.append(name)
    if nid:
        parts.append(f"کدملی {nid}")
    return " | ".join(parts)


def _resolve_target_label(conn, path: str) -> str:
    """استخراج نام مرتبط با مسیر API (پرسنل / سابقه / کاربر سیستم)."""
    # پرسنل: /api/personnel/{id}/...
    m = re.search(r"^/api/personnel/(\d+)", path or "")
    if m:
        pid = int(m.group(1))
        row = conn.execute(
            "SELECT first_name, last_name, national_id FROM personnel WHERE id=?",
            (pid,),
        ).fetchone()
        label = _person_label_from_row(row)
        return f"شناسه {pid}" + (f" — {label}" if label else "")

    # سابقه کاری: /api/work/{id}
    m = re.search(r"^/api/work/(\d+)", path or "")
    if m:
        wid = int(m.group(1))
        row = conn.execute(
            """
            SELECT p.id AS pid, p.first_name, p.last_name, p.national_id
            FROM work_history w
            JOIN personnel p ON p.id = w.personnel_id
            WHERE w.id=?
            """,
            (wid,),
        ).fetchone()
        if row:
            label = _person_label_from_row(row)
            return f"سابقه {wid} / پرسنل {row['pid']}" + (f" — {label}" if label else "")
        return f"سابقه {wid}"

    # کاربر سیستم: /api/users/{id}
    m = re.search(r"^/api/users/(\d+)", path or "")
    if m:
        uid = int(m.group(1))
        row = conn.execute(
            "SELECT username, full_name FROM app_users WHERE id=?",
            (uid,),
        ).fetchone()
        if row:
            un = (row["username"] or "").strip()
            fn = (row["full_name"] or "").strip()
            return f"کاربر {uid}" + (f" — {un}" if un else "") + (f" ({fn})" if fn else "")
        return f"کاربر {uid}"

    # ردیف نامعتبر
    m = re.search(r"^/api/invalid-imports/(\d+)", path or "")
    if m:
        iid = int(m.group(1))
        row = conn.execute(
            "SELECT first_name, last_name, national_id FROM invalid_imports WHERE id=?",
            (iid,),
        ).fetchone()
        label = _person_label_from_row(row)
        return f"ردیف نامعتبر {iid}" + (f" — {label}" if label else "")

    return ""


def _audit_log_request(request: Request, response) -> None:
    """ثبت لاگ برای درخواست‌های موفق تغییردهنده (یا عملیات مهم مثل صدور کارت/اکسل)."""
    if response is None or getattr(response, "status_code", 500) >= 400:
        return
    method = (request.method or "").upper()
    path = request.url.path or ""
    if not path.startswith("/api/"):
        return
    # ورود/خروج جداگانه و با جزئیات بیشتر ثبت می‌شوند
    if path in ("/api/auth/login", "/api/auth/logout", "/api/auth/me"):
        return
    # دانلود فایل موقت نیاز به لاگ عملیاتی ندارد
    if path.startswith("/api/downloads/"):
        return
    # فقط متدهای تغییردهنده + چند GET مهم
    action, label = _map_api_action(method, path)
    if not action:
        return
    token = _extract_token(
        request.headers.get("authorization"),
        request.headers.get("x-auth-token"),
    )
    username = ""
    user_id = None
    try:
        with get_db() as conn:
            if token:
                user = get_user_by_token(conn, token)
                if user:
                    username = user.get("username") or ""
                    user_id = user.get("id")
            target = _resolve_target_label(conn, path)
            detail = label
            if target:
                detail = f"{label} — {target}"
            log_access(
                conn,
                action=action,
                username=username or "—",
                user_id=user_id,
                detail=detail,
                ip=_client_ip(request),
                user_agent=_client_ua(request),
            )
    except Exception:
        pass


@app.post("/api/auth/login")
async def auth_login(request: Request):
    try:
        data = await request.json()
    except Exception:
        raise HTTPException(400, "داده نامعتبر است")
    username = str(data.get("username") or "").strip()
    password = str(data.get("password") or "")
    if not username or not password:
        raise HTTPException(400, "نام کاربری و رمز عبور الزامی است")

    ip = _client_ip(request)
    ua = _client_ua(request)

    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM app_users WHERE username=?", (username,)
        ).fetchone()
        if not row or not verify_password(password, row["password_hash"]):
            log_access(
                conn,
                action="login_failed",
                username=username,
                user_id=row["id"] if row else None,
                detail="نام کاربری یا رمز عبور اشتباه",
                ip=ip,
                user_agent=ua,
            )
            raise HTTPException(401, "نام کاربری یا رمز عبور اشتباه است")
        if not row["is_active"]:
            log_access(
                conn,
                action="login_denied",
                username=username,
                user_id=row["id"],
                detail="حساب غیرفعال",
                ip=ip,
                user_agent=ua,
            )
            raise HTTPException(403, "این حساب غیرفعال است")
        token = create_session(conn, row["id"])
        log_access(
            conn,
            action="login",
            username=row["username"],
            user_id=row["id"],
            detail=f"ورود موفق — نقش: {row['role']}",
            ip=ip,
            user_agent=ua,
        )
        perms_raw = row["permissions"] if "permissions" in row.keys() else ""
        user = {
            "id": row["id"],
            "username": row["username"],
            "role": row["role"],
            "full_name": row["full_name"] or "",
            "is_active": bool(row["is_active"]),
            "permissions": parse_permissions(perms_raw),
        }
    return {"token": token, "user": user_to_public(user)}


@app.post("/api/auth/logout")
def auth_logout(
    request: Request,
    authorization: Optional[str] = Header(None),
    x_auth_token: Optional[str] = Header(None, alias="X-Auth-Token"),
):
    token = _extract_token(authorization, x_auth_token)
    if token:
        with get_db() as conn:
            user = get_user_by_token(conn, token)
            if user:
                log_access(
                    conn,
                    action="logout",
                    username=user.get("username") or "",
                    user_id=user.get("id"),
                    detail="خروج از سامانه",
                    ip=_client_ip(request),
                    user_agent=_client_ua(request),
                )
            delete_session(conn, token)
    return {"message": "خروج انجام شد"}


@app.get("/api/auth/me")
def auth_me(user=Depends(get_current_user)):
    return user_to_public(user)


# ──────────────────────────────────────────────
# مدیریت کاربران (فقط مدیر)
# ──────────────────────────────────────────────
@app.get("/api/users")
def list_users(_admin=Depends(require_admin)):
    with get_db() as conn:
        rows = conn.execute(
            """
            SELECT id, username, role, full_name, is_active, permissions, created_at, updated_at
            FROM app_users
            ORDER BY id
            """
        ).fetchall()
        result = []
        for r in rows:
            d = dict(r)
            d["permissions"] = parse_permissions(d.get("permissions"))
            pub = user_to_public(d)
            pub["created_at"] = r["created_at"]
            pub["updated_at"] = r["updated_at"]
            result.append(pub)
        return result


@app.get("/api/meta/permissions")
def get_permission_defs(_admin=Depends(require_admin)):
    """لیست دسترسی‌های قابل انتخاب برای کاربر سفارشی"""
    return [{"key": k, "label": lab} for k, lab in PERMISSION_DEFS]


@app.post("/api/users")
async def create_user(request: Request, _admin=Depends(require_admin)):
    data = await request.json()
    username = str(data.get("username") or "").strip()
    password = str(data.get("password") or "")
    role = str(data.get("role") or ROLE_USER).strip()
    full_name = str(data.get("full_name") or "").strip()
    if not username or len(username) < 3:
        raise HTTPException(400, "نام کاربری حداقل ۳ کاراکتر باشد")
    if not password or len(password) < 6:
        raise HTTPException(400, "رمز عبور حداقل ۶ کاراکتر باشد")
    if role not in VALID_ROLES:
        raise HTTPException(400, "نقش نامعتبر است")
    perms_json = ""
    if role == ROLE_CUSTOM:
        perms_json = serialize_permissions(data.get("permissions") or [])
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_db() as conn:
        exists = conn.execute(
            "SELECT id FROM app_users WHERE username=?", (username,)
        ).fetchone()
        if exists:
            raise HTTPException(400, "این نام کاربری قبلاً ثبت شده است")
        cur = conn.execute(
            """
            INSERT INTO app_users (username, password_hash, role, full_name, is_active, permissions, created_at, updated_at)
            VALUES (?, ?, ?, ?, 1, ?, ?, ?)
            """,
            (username, hash_password(password), role, full_name, perms_json, now, now),
        )
        return {"id": cur.lastrowid, "message": "کاربر ایجاد شد"}


@app.put("/api/users/{uid}")
async def update_user(uid: int, request: Request, admin=Depends(require_admin)):
    data = await request.json()
    with get_db() as conn:
        row = conn.execute("SELECT * FROM app_users WHERE id=?", (uid,)).fetchone()
        if not row:
            raise HTTPException(404, "کاربر یافت نشد")

        username = str(data.get("username") or row["username"]).strip()
        role = str(data.get("role") or row["role"]).strip()
        full_name = str(data.get("full_name") if "full_name" in data else (row["full_name"] or "")).strip()
        is_active = data.get("is_active")
        if is_active is None:
            is_active = row["is_active"]
        else:
            is_active = 1 if is_active else 0

        if role not in VALID_ROLES:
            raise HTTPException(400, "نقش نامعتبر است")
        if not username or len(username) < 3:
            raise HTTPException(400, "نام کاربری حداقل ۳ کاراکتر باشد")

        # جلوگیری از حذف آخرین مدیر فعال
        if row["role"] == ROLE_ADMIN and (role != ROLE_ADMIN or not is_active):
            other_admins = conn.execute(
                "SELECT COUNT(*) FROM app_users WHERE role=? AND is_active=1 AND id!=?",
                (ROLE_ADMIN, uid),
            ).fetchone()[0]
            if other_admins == 0:
                raise HTTPException(400, "حداقل یک مدیر فعال باید در سیستم باقی بماند")

        dup = conn.execute(
            "SELECT id FROM app_users WHERE username=? AND id!=?", (username, uid)
        ).fetchone()
        if dup:
            raise HTTPException(400, "این نام کاربری قبلاً ثبت شده است")

        if role == ROLE_CUSTOM:
            perms_json = serialize_permissions(data.get("permissions") or [])
        else:
            perms_json = ""

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            """
            UPDATE app_users
            SET username=?, role=?, full_name=?, is_active=?, permissions=?, updated_at=?
            WHERE id=?
            """,
            (username, role, full_name, is_active, perms_json, now, uid),
        )

        new_password = str(data.get("password") or "")
        if new_password:
            if len(new_password) < 6:
                raise HTTPException(400, "رمز عبور حداقل ۶ کاراکتر باشد")
            conn.execute(
                "UPDATE app_users SET password_hash=?, updated_at=? WHERE id=?",
                (hash_password(new_password), now, uid),
            )
            # پس از تغییر رمز، نشست‌های قبلی باطل شود
            conn.execute("DELETE FROM app_sessions WHERE user_id=?", (uid,))

    return {"message": "کاربر بروزرسانی شد"}


@app.delete("/api/users/{uid}")
def delete_user(uid: int, admin=Depends(require_admin)):
    with get_db() as conn:
        row = conn.execute("SELECT * FROM app_users WHERE id=?", (uid,)).fetchone()
        if not row:
            raise HTTPException(404, "کاربر یافت نشد")
        if row["id"] == admin["id"]:
            raise HTTPException(400, "نمی‌توانید حساب خودتان را حذف کنید")
        if row["role"] == ROLE_ADMIN:
            other_admins = conn.execute(
                "SELECT COUNT(*) FROM app_users WHERE role=? AND is_active=1 AND id!=?",
                (ROLE_ADMIN, uid),
            ).fetchone()[0]
            if other_admins == 0:
                raise HTTPException(400, "حداقل یک مدیر فعال باید در سیستم باقی بماند")
        conn.execute("DELETE FROM app_sessions WHERE user_id=?", (uid,))
        conn.execute("DELETE FROM app_users WHERE id=?", (uid,))
    return {"message": "کاربر حذف شد"}


# ──────────────────────────────────────────────
# لاگ دسترسی کاربران (فقط مدیر)
# ──────────────────────────────────────────────
ACCESS_LOG_ACTIONS = {
    "login": "ورود موفق",
    "login_failed": "ورود ناموفق",
    "login_denied": "ورود رد شده",
    "logout": "خروج",
    "user_create": "ایجاد کاربر",
    "user_update": "ویرایش کاربر",
    "user_delete": "حذف کاربر",
    "personnel_create": "ایجاد پرسنل",
    "personnel_update": "ویرایش پرسنل",
    "personnel_delete": "حذف پرسنل",
    "photo_upload": "آپلود عکس",
    "photo_delete": "حذف عکس",
    "work_create": "افزودن سابقه کاری",
    "work_update": "ویرایش سابقه کاری",
    "work_delete": "حذف سابقه کاری",
    "work_promote": "ارتقا سابقه",
    "card_issue": "صدور کارت",
    "card_renew": "تمدید کارت",
    "card_reset": "بازنشانی وضعیت کارت",
    "export_excel": "خروجی اکسل",
    "import_excel": "ایمپورت اکسل",
    "invalid_resolve": "رفع ردیف نامعتبر",
    "invalid_delete": "حذف ردیف نامعتبر",
    "invalid_clear": "پاک‌سازی ردیف‌های نامعتبر",
    "merge_histories": "ادغام سوابق",
    "logs_clear": "پاک‌سازی لاگ",
    "api_action": "سایر عملیات",
}


@app.get("/api/access-logs")
def list_access_logs(
    q: Optional[str] = None,
    action: Optional[str] = None,
    limit: int = 100,
    offset: int = 0,
    _admin=Depends(require_admin),
):
    """لیست لاگ‌های دسترسی با فیلتر نام کاربری/جزئیات و نوع رویداد."""
    limit = max(1, min(int(limit or 100), 500))
    offset = max(0, int(offset or 0))
    clauses = []
    params = []
    if q and q.strip():
        like = f"%{q.strip()}%"
        clauses.append("(username LIKE ? OR detail LIKE ? OR ip LIKE ?)")
        params.extend([like, like, like])
    if action and action.strip():
        clauses.append("action = ?")
        params.append(action.strip())
    where = (" WHERE " + " AND ".join(clauses)) if clauses else ""
    with get_db() as conn:
        total = conn.execute(
            f"SELECT COUNT(*) FROM access_logs{where}", params
        ).fetchone()[0]
        rows = conn.execute(
            f"""
            SELECT id, user_id, username, action, detail, ip, user_agent, created_at
            FROM access_logs
            {where}
            ORDER BY id DESC
            LIMIT ? OFFSET ?
            """,
            params + [limit, offset],
        ).fetchall()
        items = []
        for r in rows:
            d = dict(r)
            d["action_label"] = ACCESS_LOG_ACTIONS.get(d.get("action") or "", d.get("action") or "")
            items.append(d)
    return {
        "items": items,
        "total": total,
        "limit": limit,
        "offset": offset,
        "actions": [{"key": k, "label": v} for k, v in ACCESS_LOG_ACTIONS.items()],
    }


@app.delete("/api/access-logs")
def clear_access_logs(
    older_than_days: Optional[int] = None,
    admin=Depends(require_admin),
):
    """پاک‌سازی لاگ‌ها. اگر older_than_days داده شود فقط قدیمی‌تر از آن روز حذف می‌شود."""
    with get_db() as conn:
        if older_than_days is not None:
            days = max(1, int(older_than_days))
            cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
            conn.execute("DELETE FROM access_logs WHERE created_at < ?", (cutoff,))
            msg = f"لاگ‌های قدیمی‌تر از {days} روز حذف شد"
        else:
            conn.execute("DELETE FROM access_logs")
            msg = "همه لاگ‌های دسترسی حذف شد"
        deleted = conn.execute("SELECT changes()").fetchone()[0]
    return {"message": msg, "deleted": deleted}


# ──────────────────────────────────────────────
# صفحات HTML
# ──────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


# ──────────────────────────────────────────────
# توابع کمکی
# ──────────────────────────────────────────────
def row_to_dict(row):
    if row is None:
        return None
    return dict(row)

def _norm_field(val: str, field: str = "") -> str:
    """یکسان‌سازی حروف عربی/فارسی و ارقام هنگام ذخیره."""
    if field == "active_status":
        return normalize_active_status(val)
    if field in ("national_id",):
        return normalize_text(str(val or "")).replace(" ", "").replace("-", "")
    if field in ("indicator_num", "row_num", "phone", "children_count"):
        return normalize_text(str(val or "")).replace(" ", "")
    return normalize_text(str(val or ""))

def normalize_jalali_value(value, *, field_label="تاریخ", strict=True):
    """تبدیل/اعتبارسنجی تاریخ. رشته‌ها فقط باید تاریخ جلالی معتبر YYYY/MM/DD باشند."""
    if value is None or str(value).strip() == "":
        return ""
    if isinstance(value, datetime):
        return jdatetime.date.fromgregorian(date=value.date()).strftime("%Y/%m/%d")
    if isinstance(value, date):
        return jdatetime.date.fromgregorian(date=value).strftime("%Y/%m/%d")

    text = str(value).strip()
    digits = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
    text = text.translate(digits).replace("-", "/").replace(".", "/")
    parts = text.split("/")
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        y, m, d = map(int, parts)
        if 0 <= y <= 99:
            y = 1400 + y if y <= 95 else 1300 + y
        if not (1200 <= y <= 1600):
            raise ValueError(f"{field_label}: تاریخ باید جلالی باشد (مثال 1405/06/05)، مقدار واردشده: {text}")
        try:
            return jdatetime.date(y, m, d).strftime("%Y/%m/%d")
        except ValueError:
            raise ValueError(f"{field_label}: تاریخ جلالی نامعتبر است: {text}")
    if strict:
        raise ValueError(f"{field_label}: فرمت تاریخ جلالی نامعتبر است؛ فرمت صحیح YYYY/MM/DD است. مقدار: {text}")
    return text

def jalali_to_tuple(value):
    """تبدیل تاریخ جلالی نرمال‌شده به (y, m, d) برای مقایسه؛ در صورت نامعتبر بودن None."""
    if not value:
        return None
    try:
        text = normalize_jalali_value(value, field_label="تاریخ", strict=True)
    except ValueError:
        return None
    parts = text.split("/")
    if len(parts) != 3:
        return None
    try:
        return tuple(int(p) for p in parts)
    except ValueError:
        return None


def ensure_jalali_after(earlier, later, *, earlier_label="تاریخ اول", later_label="تاریخ دوم"):
    """اگر later <= earlier باشد، ValueError با پیام فارسی پرتاب می‌کند."""
    if not earlier or not later:
        return
    t1 = jalali_to_tuple(earlier)
    t2 = jalali_to_tuple(later)
    if t1 is None:
        raise ValueError(f"{earlier_label} نامعتبر است")
    if t2 is None:
        raise ValueError(f"{later_label} نامعتبر است")
    if t2 <= t1:
        raise ValueError(
            f"{later_label} باید بعد از {earlier_label} باشد "
            f"(«{later}» بعد از «{earlier}» نیست). لطفاً تاریخ را اصلاح کنید."
        )


def normalize_date_fields(data: dict, *, labels=None) -> dict:
    result = dict(data)
    labels = labels or {name: label for name, label in PERSONAL_FIELDS + WORK_FIELDS}
    for field in DATE_FIELDS:
        if field in result and result[field] not in (None, ""):
            try:
                result[field] = normalize_jalali_value(
                    result[field], field_label=labels.get(field, field), strict=True
                )
            except ValueError as exc:
                raise HTTPException(400, str(exc))
    return result


# ──────────────────────────────────────────────
# API پرسنل
# ──────────────────────────────────────────────
@app.get("/api/personnel")
def list_personnel(
    q: Optional[str] = None,
    limit: int = 500,
    offset: int = 0,
    _user=Depends(get_current_user),
):
    """لیست پرسنل با صفحه‌بندی و جستجوی بهینه — یک کوئری به‌جای دو کوئری"""
    limit = max(1, min(limit, 5000))
    offset = max(0, offset)
    # COUNT(*) OVER() در همان SELECT تا رفت‌وبرگشت دوم حذف شود
    select_cols = """
        id, row_num, first_name, last_name, father_name,
        national_id, indicator_num, company, active_status, phone,
        CASE WHEN photo IS NOT NULL AND TRIM(photo) != '' THEN 1 ELSE 0 END AS has_photo,
        CASE WHEN card_issued IS NOT NULL AND TRIM(card_issued) != '' THEN 1 ELSE 0 END AS card_issued,
        COALESCE(card_type, '') AS card_type,
        (SELECT COUNT(*) FROM work_history wh WHERE wh.personnel_id = personnel.id) AS work_count,
        (SELECT COUNT(*) FROM side_notes sn WHERE sn.personnel_id = personnel.id) AS side_notes_count,
        COUNT(*) OVER() AS _total
    """

    with get_db() as conn:
        if q and q.strip():
            raw = q.strip()
            q_norm = normalize_for_search(raw)
            digits_only = q_norm.replace(" ", "").replace("-", "")
            if digits_only.isdigit() and len(digits_only) >= 4:
                like = f"%{digits_only}%"
                rows = conn.execute(
                    f"""
                    SELECT {select_cols}
                    FROM personnel
                    WHERE national_id LIKE ? OR phone LIKE ?
                    ORDER BY id DESC
                    LIMIT ? OFFSET ?
                    """,
                    (like, like, limit, offset),
                ).fetchall()
            else:
                like = f"%{q_norm}%"
                # جستجوی متنی سبک‌تر: بدون REPLACE زنجیره‌ای روی کل جدول
                # نرمال‌سازی سمت کلاینت/ورودی کافی است؛ برای ی/ك رایج OR ساده
                rows = conn.execute(
                    f"""
                    SELECT {select_cols}
                    FROM personnel
                    WHERE first_name LIKE ? OR last_name LIKE ?
                       OR father_name LIKE ? OR company LIKE ?
                       OR national_id LIKE ? OR phone LIKE ?
                       OR first_name LIKE ? OR last_name LIKE ? OR company LIKE ?
                    ORDER BY id DESC
                    LIMIT ? OFFSET ?
                    """,
                    (
                        like, like, like, like, like, like,
                        like.replace("ی", "ي").replace("ک", "ك"),
                        like.replace("ی", "ي").replace("ک", "ك"),
                        like.replace("ی", "ي").replace("ک", "ك"),
                        limit, offset,
                    ),
                ).fetchall()
        else:
            rows = conn.execute(
                f"""
                SELECT {select_cols}
                FROM personnel
                ORDER BY id DESC
                LIMIT ? OFFSET ?
                """,
                (limit, offset),
            ).fetchall()

        total = int(rows[0]["_total"]) if rows else (
            conn.execute("SELECT COUNT(*) FROM personnel").fetchone()[0] if not (q and q.strip()) else 0
        )
        result = []
        for r in rows:
            d = row_to_dict(r)
            d.pop("_total", None)
            d["has_photo"] = bool(d.get("has_photo"))
            d["card_issued"] = bool(d.get("card_issued"))
            ct = (d.get("card_type") or "").strip()
            d["card_type"] = ct if ct in ("temp", "one_year") else ("" if not d["card_issued"] else "temp")
            try:
                d["work_count"] = int(d.get("work_count") or 0)
            except (TypeError, ValueError):
                d["work_count"] = 0
            try:
                d["side_notes_count"] = int(d.get("side_notes_count") or 0)
            except (TypeError, ValueError):
                d["side_notes_count"] = 0
            result.append(d)
        return {"items": result, "total": total, "limit": limit, "offset": offset}


@app.get("/api/personnel/by-national-id/{national_id}")
def get_personnel_by_national_id(national_id: str, _user=Depends(get_current_user)):
    """جستجوی رکورد اصلی بر اساس کد ملی — برای پر کردن فرم فرد جدید"""
    nid = _normalize_national_id(national_id)
    if not nid or not nid.isdigit() or len(nid) != 10:
        raise HTTPException(400, "کد ملی نامعتبر است")
    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM personnel WHERE national_id=?",
            (nid,),
        ).fetchone()
        if not row:
            return {"found": False, "national_id": nid}
        d = row_to_dict(row)
        d["found"] = True
        # تعداد سوابق
        try:
            d["work_count"] = conn.execute(
                "SELECT COUNT(*) FROM work_history WHERE personnel_id=?",
                (d["id"],),
            ).fetchone()[0]
        except Exception:
            d["work_count"] = 0
        try:
            d["side_notes_count"] = conn.execute(
                "SELECT COUNT(*) FROM side_notes WHERE personnel_id=?",
                (d["id"],),
            ).fetchone()[0]
        except Exception:
            d["side_notes_count"] = 0
        return d


@app.get("/api/personnel/{pid}")
def get_personnel(pid: int, _user=Depends(get_current_user)):
    with get_db() as conn:
        row = conn.execute("SELECT * FROM personnel WHERE id=?", (pid,)).fetchone()
        if not row:
            raise HTTPException(404, "رکورد یافت نشد")
        return row_to_dict(row)


def _normalize_national_id(raw) -> str:
    """نرمال‌سازی کد ملی: ارقام فارسی/عربی → انگلیسی، حذف فاصله و خط تیره"""
    if raw is None:
        return ""
    text = str(raw).strip()
    digits = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
    text = text.translate(digits).replace(" ", "").replace("-", "")
    return text


def _normalize_indicator(raw) -> str:
    """نرمال‌سازی شماره اندیکاتور: ارقام فارسی/عربی → انگلیسی، حذف فاصله"""
    if raw is None:
        return ""
    text = str(raw).strip()
    digits = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
    text = text.translate(digits).replace(" ", "").replace("-", "")
    if text.endswith(".0"):
        text = text[:-2]
    try:
        if "." in text:
            text = str(int(float(text)))
    except Exception:
        pass
    return text


def _max_numeric_in_tables(conn, column: str) -> int:
    """بیشترین مقدار عددی یک ستون در personnel و work_history"""
    mx = 0
    for table in ("personnel", "work_history"):
        try:
            row = conn.execute(
                f"""
                SELECT MAX(CAST({column} AS INTEGER))
                FROM {table}
                WHERE {column} IS NOT NULL
                  AND TRIM({column}) != ''
                  AND {column} GLOB '[0-9]*'
                """
            ).fetchone()
            if row and row[0] is not None:
                mx = max(mx, int(row[0]))
        except Exception:
            pass
    return mx


def _next_indicator_num(conn) -> str:
    """آخرین شماره اندیکاتور + ۱"""
    return str(_max_numeric_in_tables(conn, "indicator_num") + 1)


def _next_row_num(conn) -> str:
    """آخرین شماره ردیف + ۱"""
    return str(_max_numeric_in_tables(conn, "row_num") + 1)


def _today_jalali_str() -> str:
    """تاریخ جاری سیستم به‌صورت جلالی YYYY/MM/DD"""
    try:
        return jdatetime.date.today().strftime("%Y/%m/%d")
    except Exception:
        return datetime.now().strftime("%Y/%m/%d")


def _auto_fill_row_and_indicator(conn, data_dict: dict) -> dict:
    """
    اگر ردیف یا اندیکاتور خالی باشد → آخرین + ۱
    تاریخ صدور موقت همیشه از تاریخ جاری سیستم
    """
    data_dict = dict(data_dict)
    ind = _normalize_indicator(data_dict.get("indicator_num", ""))
    row = _normalize_indicator(data_dict.get("row_num", ""))
    if not ind:
        ind = _next_indicator_num(conn)
    if not row:
        row = _next_row_num(conn)
    data_dict["indicator_num"] = ind
    data_dict["row_num"] = row
    data_dict["temp_issue_date"] = _today_jalali_str()
    return data_dict


def _indicator_sort_key(ind: str):
    """کلید مرتب‌سازی اندیکاتور — عدد کمتر = اصلی‌تر؛ خالی در انتها"""
    s = (ind or "").strip()
    if not s:
        return (1, 999999999)
    try:
        return (0, int(s))
    except ValueError:
        return (0, s)


def _is_indicator_smaller(a: str, b: str) -> bool:
    """آیا اندیکاتور a کوچکتر از b است؟"""
    return _indicator_sort_key(a) < _indicator_sort_key(b)


def _personnel_row_to_work_vals(row, personnel_id: int) -> tuple:
    """ساخت مقادیر INSERT برای work_history از یک رکورد پرسنل"""
    work_fields = [f for f, _ in WORK_FIELDS]
    vals = [personnel_id]
    for fname in work_fields:
        if fname in ("start_date", "end_date", "work_description"):
            vals.append("")
        else:
            try:
                v = row[fname] if row[fname] is not None else ""
            except (KeyError, IndexError, TypeError):
                v = ""
            vals.append(v if v is not None else "")
    cols = ["personnel_id"] + work_fields
    return cols, vals


def _move_personnel_to_history(conn, personnel_row, target_personnel_id: int):
    """انتقال یک رکورد پرسنل به سوابق کاری فرد هدف"""
    cols, vals = _personnel_row_to_work_vals(personnel_row, target_personnel_id)
    placeholders = ", ".join(["?"] * len(vals))
    conn.execute(
        f"INSERT INTO work_history ({', '.join(cols)}) VALUES ({placeholders})",
        vals,
    )


def _insert_data_as_history(conn, data_dict: dict, target_personnel_id: int):
    """ثبت داده‌های دیکشنری به‌عنوان سابقه کاری — شناسه ردیف را برمی‌گرداند"""
    work_fields = [f for f, _ in WORK_FIELDS]
    vals = [target_personnel_id]
    for fname in work_fields:
        if fname in ("start_date", "end_date", "work_description"):
            vals.append(data_dict.get(fname, "") or "")
        else:
            vals.append(data_dict.get(fname, "") or "")
    cols = ["personnel_id"] + work_fields
    placeholders = ", ".join(["?"] * len(vals))
    cur = conn.execute(
        f"INSERT INTO work_history ({', '.join(cols)}) VALUES ({placeholders})",
        vals,
    )
    return cur.lastrowid


def _promote_history_to_main(conn, wid: int, pid: int):
    """جابجایی یک سابقه با رکورد اصلی (بدون چک دسترسی)"""
    personal_fields = [f for f, _ in PERSONAL_FIELDS]
    work_fields = [f for f, _ in WORK_FIELDS]

    hist = conn.execute("SELECT * FROM work_history WHERE id=?", (wid,)).fetchone()
    if not hist:
        raise HTTPException(404, "سابقه یافت نشد")
    main = conn.execute("SELECT * FROM personnel WHERE id=?", (pid,)).fetchone()
    if not main:
        raise HTTPException(404, "رکورد اصلی پرسنل یافت نشد")

    hist_dict = row_to_dict(hist)
    main_dict = row_to_dict(main)

    # رکورد اصلی فعلی → داخل همین ردیف سابقه
    hist_update_vals = []
    for fname in work_fields:
        if fname in ("start_date", "end_date", "work_description"):
            hist_update_vals.append("")
        else:
            hist_update_vals.append(main_dict.get(fname) or "")
    hist_set = ", ".join([f"{f}=?" for f in work_fields])
    conn.execute(
        f"UPDATE work_history SET {hist_set} WHERE id=?",
        hist_update_vals + [wid],
    )

    # داده سابقه → رکورد اصلی (عکس دست نخورده)
    main_update_vals = [hist_dict.get(fname) or "" for fname in personal_fields]
    main_set = ", ".join([f"{f}=?" for f in personal_fields])
    conn.execute(
        f"UPDATE personnel SET {main_set} WHERE id=?",
        main_update_vals + [pid],
    )


def _ensure_indicator(conn, row_dict: dict) -> dict:
    """اگر اندیکاتور خالی است شماره خودکار بده"""
    row_dict = dict(row_dict)
    if not (row_dict.get("indicator_num") or "").strip():
        row_dict["indicator_num"] = _next_indicator_num(conn)
    if not (row_dict.get("row_num") or "").strip():
        row_dict["row_num"] = _next_row_num(conn)
    return row_dict


def _merge_duplicates_for_national_id(conn, national_id: str, prefer_main_id: int | None = None):
    """
    تمام رکوردهای پرسنل با این کد ملی را ادغام می‌کند.
    prefer_main_id در صورت وجود به‌عنوان رکورد اصلی نگه داشته می‌شود.
    """
    nid = _normalize_national_id(national_id)
    if not nid:
        return 0
    rows = [
        row_to_dict(r)
        for r in conn.execute(
            "SELECT * FROM personnel WHERE national_id=? ORDER BY id",
            (nid,),
        ).fetchall()
    ]
    if len(rows) < 2:
        return 0

    def score(row):
        if prefer_main_id and int(row.get("id") or 0) == int(prefer_main_id):
            return (2, 0, 0)  # بالاترین اولویت
        ind = (row.get("indicator_num") or "").strip()
        try:
            ind_n = int(ind) if ind else -1
        except ValueError:
            ind_n = 0
        has_ind = 1 if ind else 0
        return (has_ind, ind_n, int(row.get("id") or 0))

    rows_sorted = sorted(rows, key=score, reverse=True)
    main = rows_sorted[0]
    others = rows_sorted[1:]
    main_id = main["id"]

    # اصلی بدون اندیکاتور → پر شود
    if not (main.get("indicator_num") or "").strip():
        main["indicator_num"] = _next_indicator_num(conn)
        conn.execute(
            "UPDATE personnel SET indicator_num=? WHERE id=?",
            (main["indicator_num"], main_id),
        )

    moved = 0
    for other in others:
        other = _ensure_indicator(conn, other)
        while conn.execute(
            "SELECT id FROM work_history WHERE personnel_id=? AND indicator_num=?",
            (main_id, other.get("indicator_num") or ""),
        ).fetchone():
            other["indicator_num"] = _next_indicator_num(conn)

        _insert_data_as_history(conn, other, main_id)
        conn.execute(
            "UPDATE work_history SET personnel_id=? WHERE personnel_id=?",
            (main_id, other["id"]),
        )
        if (other.get("photo") or "").strip() and not (main.get("photo") or "").strip():
            conn.execute(
                "UPDATE personnel SET photo=? WHERE id=?",
                (other.get("photo"), main_id),
            )
            main["photo"] = other.get("photo")
        conn.execute("DELETE FROM personnel WHERE id=?", (other["id"],))
        moved += 1
    return moved


def _register_person_record(conn, data_dict: dict, *, source_file: str = "web"):
    """
    ثبت فرد جدید:
    - اگر کد ملی وجود نداشت → رکورد اصلی جدید
    - اگر کد ملی وجود داشت → ابتدا به‌عنوان سابقه ثبت، سپس با رکورد اصلی جابجا می‌شود
      (رکورد جدید = اصلی، رکورد قبلی = سابقه)
    - پس از ثبت، ادغام سوابق برای همان کد ملی اعمال می‌شود
    برمی‌گرداند: (personnel_id, message, was_promoted)
    """
    data_dict = _auto_fill_row_and_indicator(conn, data_dict)
    fields = [f for f, _ in PERSONAL_FIELDS]
    national_id = _normalize_national_id(data_dict.get("national_id") or "")
    data_dict["national_id"] = national_id
    indicator_num = data_dict.get("indicator_num") or ""
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    main = None
    if national_id:
        main = conn.execute(
            "SELECT * FROM personnel WHERE national_id=?",
            (national_id,),
        ).fetchone()

    if main:
        main_ind = (main["indicator_num"] or "").strip()
        # همان اندیکاتور اصلی → تکراری
        if main_ind == (indicator_num or "").strip():
            name = f"{main['first_name'] or ''} {main['last_name'] or ''}".strip()
            raise HTTPException(
                400,
                f"این ترکیب کد ملی و شماره اندیکاتور قبلاً ثبت شده است (شناسه {main['id']}"
                + (f": {name}" if name else "")
                + f" — اندیکاتور: {main_ind or '—'})",
            )
        # همان اندیکاتور در سوابق
        dup_hist = conn.execute(
            "SELECT id FROM work_history WHERE personnel_id=? AND indicator_num=?",
            (main["id"], indicator_num),
        ).fetchone()
        if dup_hist:
            raise HTTPException(
                400,
                f"این شماره اندیکاتور ({indicator_num}) قبلاً در سوابق این فرد ثبت شده است",
            )

        # اگر اصلی اندیکاتور خالی دارد، قبل از رفتن به سابقه شماره بده
        if not main_ind:
            assigned = _next_indicator_num(conn)
            conn.execute(
                "UPDATE personnel SET indicator_num=? WHERE id=?",
                (assigned, main["id"]),
            )
            main_ind = assigned
            main = conn.execute(
                "SELECT * FROM personnel WHERE id=?", (main["id"],)
            ).fetchone()

        # ۱) رکورد جدید → سابقه
        wid = _insert_data_as_history(conn, data_dict, main["id"])
        # ۲) جابجایی: سابقه جدید = اصلی، قبلی = سابقه
        _promote_history_to_main(conn, wid, main["id"])
        # ۳) ادغام کامل هر رکورد تکراری احتمالی
        _merge_duplicates_for_national_id(conn, national_id, prefer_main_id=main["id"])
        return (
            main["id"],
            (
                f"ادغام انجام شد. رکورد جدید (اندیکاتور {indicator_num or '—'}) "
                f"به‌عنوان رکورد اصلی ثبت شد و رکورد قبلی (اندیکاتور {main_ind or '—'}) "
                f"به سوابق منتقل شد."
            ),
            True,
        )

    # کد ملی جدید → رکورد اصلی
    values = [data_dict.get(f, "") or "" for f in fields]
    values_full = values + [now_str, source_file]
    cols = fields + ["created_at", "source_file"]
    placeholders = ", ".join(["?"] * len(values_full))
    try:
        cur = conn.execute(
            f"INSERT INTO personnel ({', '.join(cols)}) VALUES ({placeholders})",
            values_full,
        )
        pid = cur.lastrowid
        # ادغام احتمالی (اگر به‌خاطر ناسازگاری ایندکس تکراری مانده باشد)
        moved = _merge_duplicates_for_national_id(conn, national_id, prefer_main_id=pid)
        if moved:
            return (
                pid,
                f"ذخیره شد و {moved} رکورد تکراری با همین کد ملی به سوابق ادغام شد.",
                True,
            )
        return pid, "ذخیره شد", False
    except Exception as e:
        if "UNIQUE" in str(e).upper():
            # رقابت هم‌زمان: دوباره به‌عنوان ادغام با رکورد موجود
            main = conn.execute(
                "SELECT * FROM personnel WHERE national_id=?",
                (national_id,),
            ).fetchone()
            if main:
                wid = _insert_data_as_history(conn, data_dict, main["id"])
                _promote_history_to_main(conn, wid, main["id"])
                _merge_duplicates_for_national_id(conn, national_id, prefer_main_id=main["id"])
                return (
                    main["id"],
                    f"کد ملی موجود بود؛ رکورد جدید ادغام و به‌عنوان اصلی ثبت شد.",
                    True,
                )
            raise HTTPException(400, "این کد ملی قبلاً ثبت شده است")
        raise HTTPException(500, str(e))


def _validate_national_id(nid: str, *, required: bool = True) -> str:
    """اعتبارسنجی کد ملی؛ در صورت نامعتبر بودن HTTPException پرتاب می‌کند"""
    nid = _normalize_national_id(nid)
    if not nid:
        if required:
            raise HTTPException(400, "کد ملی الزامی است")
        return ""
    if not nid.isdigit() or len(nid) != 10:
        raise HTTPException(400, "کد ملی باید دقیقاً ۱۰ رقم باشد")
    return nid

@app.post("/api/personnel")
async def create_personnel(request: Request, _admin=Depends(require_perm("personnel_create"))):
    data = normalize_date_fields(await request.json())
    fields = [f for f, _ in PERSONAL_FIELDS]
    values = [_norm_field(data.get(f, ""), f) for f in fields]
    if not values[1] and not values[2]:  # first_name, last_name
        raise HTTPException(400, "نام و نام خانوادگی الزامی است")

    national_id = _validate_national_id(data.get("national_id", ""))
    indicator_num = _normalize_indicator(data.get("indicator_num", ""))
    try:
        values[fields.index("national_id")] = national_id
    except ValueError:
        pass
    try:
        values[fields.index("indicator_num")] = indicator_num
    except ValueError:
        pass

    data_dict = {f: values[i] for i, f in enumerate(fields)}

    with get_db() as conn:
        # کد ملی جدید → اصلی؛ کد ملی تکراری → سابقه سپس جابجایی با اصلی
        pid, message, promoted = _register_person_record(conn, data_dict, source_file="web")
        # وضعیت نهایی رکورد اصلی پس از ادغام
        row = conn.execute(
            "SELECT indicator_num, row_num, first_name, last_name, father_name, "
            "national_id, active_status, "
            "(SELECT COUNT(*) FROM work_history wh WHERE wh.personnel_id = personnel.id) AS work_count, "
            "(SELECT COUNT(*) FROM side_notes sn WHERE sn.personnel_id = personnel.id) AS side_notes_count "
            "FROM personnel WHERE id=?",
            (pid,),
        ).fetchone()
        out = {
            "id": pid,
            "message": message,
            "promoted": bool(promoted),
            "merged": bool(promoted),
        }
        if row:
            d = dict(row)
            out.update({
                "indicator_num": d.get("indicator_num") or "",
                "row_num": d.get("row_num") or "",
                "first_name": d.get("first_name") or "",
                "last_name": d.get("last_name") or "",
                "father_name": d.get("father_name") or "",
                "national_id": d.get("national_id") or "",
                "active_status": d.get("active_status") or "",
                "work_count": int(d.get("work_count") or 0),
                "side_notes_count": int(d.get("side_notes_count") or 0),
            })
        return out


@app.put("/api/personnel/{pid}")
async def update_personnel(pid: int, request: Request, user=Depends(require_perm("personnel_edit"))):
    data = normalize_date_fields(await request.json())
    fields = [f for f, _ in PERSONAL_FIELDS]
    values = [_norm_field(data.get(f, ""), f) for f in fields]

    national_id = _validate_national_id(data.get("national_id", ""))
    indicator_num = _normalize_indicator(data.get("indicator_num", ""))
    try:
        nid_idx = fields.index("national_id")
        values[nid_idx] = national_id
    except ValueError:
        pass
    try:
        ind_idx = fields.index("indicator_num")
        values[ind_idx] = indicator_num
    except ValueError:
        pass

    set_clause = ", ".join([f"{f}=?" for f in fields])
    values_with_id = values + [pid]

    with get_db() as conn:
        exists = conn.execute("SELECT id, active_status FROM personnel WHERE id=?", (pid,)).fetchone()
        if not exists:
            raise HTTPException(404, "رکورد یافت نشد")

        # فقط مدیر می‌تواند پرسنل «مسدود» را از مسدودی خارج کند (تغییر به هر وضعیت دیگر)
        old_status = normalize_active_status(exists["active_status"])
        try:
            st_idx = fields.index("active_status")
            new_status = normalize_active_status(values[st_idx])
        except ValueError:
            new_status = old_status
        if old_status == "مسدود" and new_status != "مسدود":
            if user.get("role") != ROLE_ADMIN:
                raise HTTPException(
                    403,
                    "این پرسنل مسدود است. فقط مدیر سیستم می‌تواند وضعیت را از «مسدود» به حالت دیگر تغییر دهد.",
                )

        # اعتبار ۲ فقط وقتی اعتبار ۱ وجود داشته باشد قابل ثبت است
        try:
            v1 = (values[fields.index("validity1")] or "").strip()
            v2 = (values[fields.index("validity2")] or "").strip()
        except ValueError:
            v1, v2 = "", ""
        if v2 and not v1:
            raise HTTPException(
                400,
                "برای ثبت «اعتبار ۲» ابتدا باید «اعتبار ۱» وارد شده باشد.",
            )

        # هر کد ملی فقط یک رکورد اصلی دارد
        dup = conn.execute(
            "SELECT id, first_name, last_name, indicator_num FROM personnel "
            "WHERE national_id=? AND id!=?",
            (national_id, pid),
        ).fetchone()
        if dup:
            name = f"{dup['first_name'] or ''} {dup['last_name'] or ''}".strip()
            raise HTTPException(
                400,
                f"این کد ملی متعلق به پرسنل دیگری است (شناسه {dup['id']}"
                + (f": {name}" if name else "")
                + f" — اندیکاتور اصلی: {dup['indicator_num'] or '—'})",
            )
        try:
            conn.execute(f"UPDATE personnel SET {set_clause} WHERE id=?", values_with_id)
        except Exception as e:
            if "UNIQUE" in str(e).upper():
                raise HTTPException(400, "این کد ملی قبلاً ثبت شده است")
            raise HTTPException(500, str(e))
    return {"message": "بروزرسانی شد"}


@app.delete("/api/personnel/{pid}")
def delete_personnel(pid: int, _admin=Depends(require_perm("personnel_delete"))):
    with get_db() as conn:
        row = conn.execute("SELECT photo FROM personnel WHERE id=?", (pid,)).fetchone()
        if not row:
            raise HTTPException(404, "رکورد یافت نشد")
        photo = row["photo"]
        conn.execute("DELETE FROM personnel WHERE id=?", (pid,))
        if photo:
            path = os.path.join(APP_DIR, photo) if not os.path.isabs(photo) else photo
            if not os.path.exists(path):
                path = os.path.join(PHOTOS_DIR, os.path.basename(photo))
            if os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    pass
    return {"message": "حذف شد"}


# ──────────────────────────────────────────────
# عکس پرسنلی (۳×۴)
# ──────────────────────────────────────────────
PHOTO_W, PHOTO_H = 120, 160


def to_personnel_photo(img: Image.Image) -> Image.Image:
    if img.mode in ("RGBA", "P", "LA"):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        if img.mode == "P":
            img = img.convert("RGBA")
        if img.mode in ("RGBA", "LA"):
            bg.paste(img, mask=img.split()[-1])
        else:
            bg.paste(img)
        img = bg
    elif img.mode != "RGB":
        img = img.convert("RGB")

    src_w, src_h = img.size
    target = 3.0 / 4.0
    if src_w / src_h > target:
        new_w = int(src_h * target)
        left = (src_w - new_w) // 2
        img = img.crop((left, 0, left + new_w, src_h))
    elif src_w / src_h < target:
        new_h = int(src_w / target)
        top = (src_h - new_h) // 2
        img = img.crop((0, top, src_w, top + new_h))

    return img.resize((PHOTO_W, PHOTO_H), Image.Resampling.LANCZOS)


@app.post("/api/personnel/{pid}/photo")
async def upload_photo(pid: int, file: UploadFile = File(...), _admin=Depends(require_perm("photo_manage"))):
    with get_db() as conn:
        exists = conn.execute("SELECT id, photo FROM personnel WHERE id=?", (pid,)).fetchone()
        if not exists:
            raise HTTPException(404, "رکورد یافت نشد")

    try:
        content = await file.read()
        img = Image.open(io.BytesIO(content))
        img = to_personnel_photo(img)
        os.makedirs(PHOTOS_DIR, exist_ok=True)
        name = f"personnel_{pid}.jpg"
        dest = os.path.join(PHOTOS_DIR, name)
        img.save(dest, "JPEG", quality=92)
        rel = f"photos/{name}"

        old = exists["photo"]
        if old and os.path.basename(old) != name:
            old_path = os.path.join(PHOTOS_DIR, os.path.basename(old))
            if os.path.exists(old_path):
                try:
                    os.remove(old_path)
                except OSError:
                    pass

        with get_db() as conn:
            conn.execute("UPDATE personnel SET photo=? WHERE id=?", (rel, pid))
        return {"photo": rel, "url": f"/{rel}", "message": "عکس ذخیره شد"}
    except Exception as e:
        raise HTTPException(400, f"خطا در پردازش عکس: {e}")


@app.delete("/api/personnel/{pid}/photo")
def delete_photo(pid: int, _admin=Depends(require_perm("photo_manage"))):
    with get_db() as conn:
        row = conn.execute("SELECT photo FROM personnel WHERE id=?", (pid,)).fetchone()
        if not row:
            raise HTTPException(404, "رکورد یافت نشد")
        photo = row["photo"]
        conn.execute("UPDATE personnel SET photo='' WHERE id=?", (pid,))
        if photo:
            path = os.path.join(PHOTOS_DIR, os.path.basename(photo))
            if os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    pass
    return {"message": "عکس حذف شد"}


def _find_jpg_candidates_for_indicator(indicator: str) -> list:
    """در docs همه پوشه‌های منطبق با اندیکاتور را بگرد و فایل‌های jpg/jpeg را برگردان."""
    key = _normalize_indicator_key(indicator)
    if not key or not os.path.isdir(DOCS_DIR):
        return []
    matches = _scan_docs_for_indicators({key})
    candidates = []
    for m in matches:
        folder = os.path.join(DOCS_DIR, m["relative_path"])
        if not os.path.isdir(folder):
            continue
        # فقط فایل‌های مستقیم داخل همان پوشه + یک سطح زیرپوشه (برای انعطاف)
        search_roots = [folder]
        try:
            for name in os.listdir(folder):
                sub = os.path.join(folder, name)
                if os.path.isdir(sub):
                    search_roots.append(sub)
        except OSError:
            pass
        for root in search_roots:
            try:
                for fname in os.listdir(root):
                    low = fname.lower()
                    if not (low.endswith(".jpg") or low.endswith(".jpeg")):
                        continue
                    fpath = os.path.join(root, fname)
                    if not os.path.isfile(fpath):
                        continue
                    try:
                        size = os.path.getsize(fpath)
                    except OSError:
                        size = 0
                    score = 0
                    base = os.path.splitext(fname)[0]
                    base_key = _normalize_indicator_key(base)
                    if base_key == key:
                        score += 100
                    for token in ("photo", "pic", "image", "img", "3x4", "۳x۴", "عکس", "پرسنل"):
                        if token in low or token in fname:
                            score += 20
                    # فایل‌های خیلی کوچک (آیکون) را عقب بینداز
                    if size > 20_000:
                        score += 10
                    if size > 50_000:
                        score += 5
                    candidates.append({
                        "path": fpath,
                        "name": fname,
                        "size": size,
                        "score": score,
                        "folder": m["relative_path"],
                    })
            except OSError:
                continue
    candidates.sort(key=lambda c: (-c["score"], -c["size"], c["name"].lower()))
    return candidates


def _import_photo_from_file(pid: int, src_path: str) -> dict:
    """خواندن jpg از مسیر، استاندارد ۳×۴، ذخیره در photos و به‌روزرسانی دیتابیس."""
    with open(src_path, "rb") as f:
        content = f.read()
    img = Image.open(io.BytesIO(content))
    img = to_personnel_photo(img)
    os.makedirs(PHOTOS_DIR, exist_ok=True)
    name = f"personnel_{pid}.jpg"
    dest = os.path.join(PHOTOS_DIR, name)
    img.save(dest, "JPEG", quality=92)
    rel = f"photos/{name}"
    with get_db() as conn:
        row = conn.execute("SELECT photo FROM personnel WHERE id=?", (pid,)).fetchone()
        old = (row["photo"] if row else "") or ""
        if old and os.path.basename(old) != name:
            old_path = os.path.join(PHOTOS_DIR, os.path.basename(old))
            if os.path.exists(old_path):
                try:
                    os.remove(old_path)
                except OSError:
                    pass
        conn.execute("UPDATE personnel SET photo=? WHERE id=?", (rel, pid))
    return {"photo": rel, "url": f"/{rel}", "source_file": os.path.basename(src_path)}


@app.post("/api/personnel/photos/from-docs")
async def transfer_photos_from_docs(request: Request, _user=Depends(require_perm("photo_manage"))):
    """انتقال عکس JPG از پوشه docs (بر اساس اندیکاتور اصلی) به عکس پرسنلی.

    بدنه:
      ids: لیست شناسه پرسنل (الزامی)
      overwrite: اگر true باشد روی عکس موجود هم بازنویسی می‌کند (پیش‌فرض false)
    """
    try:
        payload = await request.json()
    except Exception:
        raise HTTPException(400, "داده ارسالی نامعتبر است")

    raw_ids = payload.get("ids") or []
    if not isinstance(raw_ids, list) or not raw_ids:
        raise HTTPException(400, "حداقل یک پرسنل باید انتخاب شود")
    try:
        ids = [int(x) for x in raw_ids]
    except (TypeError, ValueError):
        raise HTTPException(400, "شناسه‌های پرسنل نامعتبر است")
    overwrite = bool(payload.get("overwrite", False))

    results = []
    success = 0
    skipped = 0
    failed = 0

    with get_db() as conn:
        for pid in ids:
            row = conn.execute(
                "SELECT id, first_name, last_name, indicator_num, photo FROM personnel WHERE id=?",
                (pid,),
            ).fetchone()
            if not row:
                results.append({"id": pid, "status": "error", "message": "پرسنل یافت نشد"})
                failed += 1
                continue
            person = dict(row)
            name = f"{(person.get('first_name') or '').strip()} {(person.get('last_name') or '').strip()}".strip()
            ind = _normalize_indicator_key(person.get("indicator_num"))
            has_photo = bool((person.get("photo") or "").strip())
            if has_photo and not overwrite:
                results.append({
                    "id": pid,
                    "name": name,
                    "indicator": ind,
                    "status": "skipped",
                    "message": "عکس پرسنلی از قبل موجود است",
                })
                skipped += 1
                continue
            if not ind:
                results.append({
                    "id": pid,
                    "name": name,
                    "indicator": "",
                    "status": "error",
                    "message": "شماره اندیکاتور اصلی خالی است",
                })
                failed += 1
                continue
            candidates = _find_jpg_candidates_for_indicator(ind)
            if not candidates:
                results.append({
                    "id": pid,
                    "name": name,
                    "indicator": ind,
                    "status": "error",
                    "message": f"فایل JPG در پوشه docs برای اندیکاتور {ind} یافت نشد",
                })
                failed += 1
                continue
            chosen = candidates[0]
            try:
                info = _import_photo_from_file(pid, chosen["path"])
                results.append({
                    "id": pid,
                    "name": name,
                    "indicator": ind,
                    "status": "ok",
                    "message": f"عکس از {chosen['folder']}/{chosen['name']} منتقل شد",
                    "photo": info["photo"],
                    "url": info["url"],
                    "source_file": chosen["name"],
                })
                success += 1
            except Exception as e:
                results.append({
                    "id": pid,
                    "name": name,
                    "indicator": ind,
                    "status": "error",
                    "message": f"خطا در پردازش عکس: {e}",
                })
                failed += 1

    return {
        "message": f"انتقال عکس: {success} موفق، {skipped} ردشده، {failed} ناموفق از {len(ids)} نفر",
        "success": success,
        "skipped": skipped,
        "failed": failed,
        "total": len(ids),
        "results": results,
    }


# ──────────────────────────────────────────────
# سوابق کاری
# ──────────────────────────────────────────────
@app.get("/api/personnel/{pid}/work")
def list_work(pid: int, _user=Depends(get_current_user)):
    with get_db() as conn:
        exists = conn.execute("SELECT id FROM personnel WHERE id=?", (pid,)).fetchone()
        if not exists:
            raise HTTPException(404, "پرسنل یافت نشد")
        fields = ", ".join([f for f, _ in WORK_FIELDS])
        rows = conn.execute(
            f"SELECT id, {fields} FROM work_history WHERE personnel_id=? ORDER BY id DESC",
            (pid,),
        ).fetchall()
        return [row_to_dict(r) for r in rows]


@app.post("/api/personnel/{pid}/work")
async def create_work(pid: int, request: Request, _admin=Depends(require_perm("work_write"))):
    data = normalize_date_fields(await request.json())
    company = str(data.get("company", "") or "").strip()
    if not company:
        raise HTTPException(400, "نام شرکت الزامی است")

    with get_db() as conn:
        exists = conn.execute("SELECT id FROM personnel WHERE id=?", (pid,)).fetchone()
        if not exists:
            raise HTTPException(404, "پرسنل یافت نشد")
        data = _auto_fill_row_and_indicator(conn, data)
        fields = [f for f, _ in WORK_FIELDS]
        values = [str(data.get(f, "") or "").strip() for f in fields]
        cols = ["personnel_id"] + fields
        placeholders = ", ".join(["?"] * len(cols))
        cur = conn.execute(
            f"INSERT INTO work_history ({', '.join(cols)}) VALUES ({placeholders})",
            [pid] + values,
        )
        return {
            "id": cur.lastrowid,
            "message": "سابقه کاری اضافه شد",
            "indicator_num": data.get("indicator_num", ""),
            "row_num": data.get("row_num", ""),
        }


@app.put("/api/work/{wid}")
async def update_work(wid: int, request: Request, _admin=Depends(require_perm("work_write"))):
    data = normalize_date_fields(await request.json())
    company = str(data.get("company", "") or "").strip()
    if not company:
        raise HTTPException(400, "نام شرکت الزامی است")

    fields = [f for f, _ in WORK_FIELDS]
    set_clause = ", ".join([f"{f}=?" for f in fields])
    values = [str(data.get(f, "") or "").strip() for f in fields] + [wid]

    with get_db() as conn:
        exists = conn.execute("SELECT id FROM work_history WHERE id=?", (wid,)).fetchone()
        if not exists:
            raise HTTPException(404, "سابقه یافت نشد")
        conn.execute(f"UPDATE work_history SET {set_clause} WHERE id=?", values)
    return {"message": "بروزرسانی شد"}


@app.delete("/api/work/{wid}")
def delete_work(wid: int, _admin=Depends(require_perm("work_write"))):
    with get_db() as conn:
        exists = conn.execute("SELECT id FROM work_history WHERE id=?", (wid,)).fetchone()
        if not exists:
            raise HTTPException(404, "سابقه یافت نشد")
        conn.execute("DELETE FROM work_history WHERE id=?", (wid,))
    return {"message": "حذف شد"}


# ─── سو سابقه (مستقل از سوابق کاری) ───
@app.get("/api/personnel/{pid}/side-notes")
def list_side_notes(pid: int, _user=Depends(get_current_user)):
    """لیست سو سابقه یک پرسنل — بدون ارتباط با work_history"""
    with get_db() as conn:
        exists = conn.execute("SELECT id FROM personnel WHERE id=?", (pid,)).fetchone()
        if not exists:
            raise HTTPException(404, "پرسنل یافت نشد")
        rows = conn.execute(
            """
            SELECT id, personnel_id, description, created_at, updated_at
            FROM side_notes
            WHERE personnel_id=?
            ORDER BY id DESC
            """,
            (pid,),
        ).fetchall()
        return [dict(r) for r in rows]


@app.post("/api/personnel/{pid}/side-notes")
async def create_side_note(pid: int, request: Request, _user=Depends(require_perm("personnel_edit"))):
    """افزودن سو سابقه — فقط توضیحات؛ در سوابق کاری ثبت نمی‌شود"""
    data = await request.json()
    description = str(data.get("description", "") or "").strip()
    if not description:
        raise HTTPException(400, "توضیحات الزامی است")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_db() as conn:
        exists = conn.execute("SELECT id FROM personnel WHERE id=?", (pid,)).fetchone()
        if not exists:
            raise HTTPException(404, "پرسنل یافت نشد")
        cur = conn.execute(
            """
            INSERT INTO side_notes (personnel_id, description, created_at, updated_at)
            VALUES (?, ?, ?, ?)
            """,
            (pid, description, now, now),
        )
        return {"id": cur.lastrowid, "message": "سو سابقه ثبت شد"}


@app.put("/api/side-notes/{nid}")
async def update_side_note(nid: int, request: Request, _user=Depends(require_perm("personnel_edit"))):
    data = await request.json()
    description = str(data.get("description", "") or "").strip()
    if not description:
        raise HTTPException(400, "توضیحات الزامی است")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_db() as conn:
        exists = conn.execute("SELECT id FROM side_notes WHERE id=?", (nid,)).fetchone()
        if not exists:
            raise HTTPException(404, "رکورد یافت نشد")
        conn.execute(
            "UPDATE side_notes SET description=?, updated_at=? WHERE id=?",
            (description, now, nid),
        )
    return {"message": "بروزرسانی شد"}


@app.delete("/api/side-notes/{nid}")
def delete_side_note(nid: int, _user=Depends(require_perm("personnel_edit"))):
    with get_db() as conn:
        exists = conn.execute("SELECT id FROM side_notes WHERE id=?", (nid,)).fetchone()
        if not exists:
            raise HTTPException(404, "رکورد یافت نشد")
        conn.execute("DELETE FROM side_notes WHERE id=?", (nid,))
    return {"message": "حذف شد"}



@app.post("/api/work/{wid}/promote")
def promote_work_to_main(wid: int, _user=Depends(get_current_user)):
    """
    جابجایی سابقه با رکورد اصلی:
    - این سابقه به رکورد اصلی تبدیل می‌شود
    - رکورد اصلی فعلی به‌عنوان سابقه (جای همان ردیف) ذخیره می‌شود
    نیاز به دسترسی ویرایش پرسنل یا مدیریت سوابق دارد.
    """
    if not (
        user_has_permission(_user, "personnel_edit")
        or user_has_permission(_user, "work_write")
    ):
        raise HTTPException(403, "شما به جابجایی سابقه با رکورد اصلی دسترسی ندارید")

    personal_fields = [f for f, _ in PERSONAL_FIELDS]
    work_fields = [f for f, _ in WORK_FIELDS]

    with get_db() as conn:
        hist = conn.execute("SELECT * FROM work_history WHERE id=?", (wid,)).fetchone()
        if not hist:
            raise HTTPException(404, "سابقه یافت نشد")

        pid = hist["personnel_id"]
        main = conn.execute("SELECT * FROM personnel WHERE id=?", (pid,)).fetchone()
        if not main:
            raise HTTPException(404, "رکورد اصلی پرسنل یافت نشد")

        hist_dict = row_to_dict(hist)
        main_dict = row_to_dict(main)

        old_ind = main_dict.get("indicator_num") or "—"
        new_ind = hist_dict.get("indicator_num") or "—"

        # ۱) رکورد اصلی فعلی → داخل همین ردیف سابقه نوشته می‌شود
        hist_update_vals = []
        for fname in work_fields:
            if fname in ("start_date", "end_date", "work_description"):
                # فیلدهای اختصاصی سابقه روی رکورد اصلی نبودند → خالی
                hist_update_vals.append("")
            else:
                hist_update_vals.append(main_dict.get(fname) or "")
        hist_set = ", ".join([f"{f}=?" for f in work_fields])
        conn.execute(
            f"UPDATE work_history SET {hist_set} WHERE id=?",
            hist_update_vals + [wid],
        )

        # ۲) داده سابقه → رکورد اصلی (عکس دست نخورده می‌ماند)
        main_update_vals = []
        for fname in personal_fields:
            main_update_vals.append(hist_dict.get(fname) or "")
        main_set = ", ".join([f"{f}=?" for f in personal_fields])
        conn.execute(
            f"UPDATE personnel SET {main_set} WHERE id=?",
            main_update_vals + [pid],
        )

    return {
        "id": pid,
        "message": (
            f"جابجایی انجام شد. رکورد اصلی اکنون اندیکاتور {new_ind} است "
            f"و رکورد قبلی (اندیکاتور {old_ind}) به سوابق منتقل شد."
        ),
    }



@app.post("/api/admin/merge-histories")
def merge_histories_by_national_id(_admin=Depends(require_admin)):
    """
    ادغام سابقه با رکورد اصلی:
    - کد ملی نرمال می‌شود
    - اگر چند رکورد با یک کد ملی باشد:
      * رکورد دارای اندیکاتور (بزرگ‌تر) = اصلی
      * رکورد بدون اندیکاتور = سابقه (با اندیکاتور خودکار قدیمی)
    """
    assigned_empty = 0
    normalized_nids = 0
    merged_groups = 0
    moved_to_history = 0
    samples = []

    with get_db() as conn:
        try:
            conn.execute("DROP INDEX IF EXISTS idx_personnel_national_id_unique")
        except Exception:
            pass

        # ۱) نرمال‌سازی کد ملی
        for table in ("personnel", "work_history"):
            for r in conn.execute(f"SELECT id, national_id FROM {table}").fetchall():
                nid = _normalize_national_id(r["national_id"])
                if nid != (r["national_id"] or ""):
                    conn.execute(
                        f"UPDATE {table} SET national_id=? WHERE id=?",
                        (nid, r["id"]),
                    )
                    normalized_nids += 1

        # ۲) ادغام تکراری‌ها — اول اصلی را انتخاب کن، بعد به خالی‌ها اندیکاتور بده
        dup_groups = conn.execute(
            """
            SELECT national_id, COUNT(*) AS cnt
            FROM personnel
            WHERE national_id IS NOT NULL AND TRIM(national_id) != ''
            GROUP BY national_id
            HAVING cnt > 1
            """
        ).fetchall()

        for g in dup_groups:
            nid = g["national_id"]
            rows = [
                row_to_dict(r)
                for r in conn.execute(
                    "SELECT * FROM personnel WHERE national_id=? ORDER BY id",
                    (nid,),
                ).fetchall()
            ]
            if len(rows) < 2:
                continue

            def score(row):
                ind = (row.get("indicator_num") or "").strip()
                try:
                    ind_n = int(ind) if ind else -1
                except ValueError:
                    ind_n = 0
                # اولویت با داشتن اندیکاتور، سپس بزرگ‌تر بودن
                has_ind = 1 if ind else 0
                return (has_ind, ind_n, int(row.get("id") or 0))

            rows_sorted = sorted(rows, key=score, reverse=True)
            main = rows_sorted[0]
            others = rows_sorted[1:]
            main_id = main["id"]

            # اگر خود اصلی هم اندیکاتور نداشت، یکی بده
            if not (main.get("indicator_num") or "").strip():
                main["indicator_num"] = _next_indicator_num(conn)
                conn.execute(
                    "UPDATE personnel SET indicator_num=? WHERE id=?",
                    (main["indicator_num"], main_id),
                )
                assigned_empty += 1

            for other in others:
                # اندیکاتور خالی → خودکار قبل از انتقال به سابقه
                if not (other.get("indicator_num") or "").strip():
                    other["indicator_num"] = _next_indicator_num(conn)
                    assigned_empty += 1
                if not (other.get("row_num") or "").strip():
                    other["row_num"] = _next_row_num(conn)

                # جلوگیری از برخورد اندیکاتور در سوابق
                while conn.execute(
                    "SELECT id FROM work_history WHERE personnel_id=? AND indicator_num=?",
                    (main_id, other.get("indicator_num") or ""),
                ).fetchone():
                    other["indicator_num"] = _next_indicator_num(conn)

                _insert_data_as_history(conn, other, main_id)
                conn.execute(
                    "UPDATE work_history SET personnel_id=? WHERE personnel_id=?",
                    (main_id, other["id"]),
                )
                if (other.get("photo") or "").strip() and not (main.get("photo") or "").strip():
                    conn.execute(
                        "UPDATE personnel SET photo=? WHERE id=?",
                        (other.get("photo"), main_id),
                    )
                    main["photo"] = other.get("photo")

                conn.execute("DELETE FROM personnel WHERE id=?", (other["id"],))
                moved_to_history += 1

            merged_groups += 1
            if len(samples) < 20:
                samples.append(
                    {
                        "national_id": nid,
                        "main_id": main_id,
                        "main_indicator": main.get("indicator_num") or "",
                        "merged_count": len(others),
                    }
                )

        # ۳) اندیکاتور خالی باقی‌مانده (تک‌رکوردها)
        for table in ("personnel", "work_history"):
            for r in conn.execute(
                f"""
                SELECT id FROM {table}
                WHERE indicator_num IS NULL OR TRIM(COALESCE(indicator_num,'')) = ''
                ORDER BY id
                """
            ).fetchall():
                new_ind = _next_indicator_num(conn)
                cur = conn.execute(
                    f"SELECT row_num FROM {table} WHERE id=?", (r["id"],)
                ).fetchone()
                row_num = (cur["row_num"] or "").strip() if cur else ""
                if not row_num:
                    conn.execute(
                        f"UPDATE {table} SET indicator_num=?, row_num=? WHERE id=?",
                        (new_ind, _next_row_num(conn), r["id"]),
                    )
                else:
                    conn.execute(
                        f"UPDATE {table} SET indicator_num=? WHERE id=?",
                        (new_ind, r["id"]),
                    )
                assigned_empty += 1

        # ۴) ایندکس یکتا
        try:
            conn.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_personnel_national_id_unique "
                "ON personnel(national_id)"
            )
        except Exception as e:
            left = conn.execute(
                """
                SELECT national_id, COUNT(*) c FROM personnel
                WHERE national_id IS NOT NULL AND TRIM(national_id)!=''
                GROUP BY national_id HAVING c>1
                """
            ).fetchall()
            raise HTTPException(
                500,
                f"ادغام ناقص؛ هنوز {len(left)} کد ملی تکراری: {e}",
            )

    return {
        "message": (
            f"ادغام انجام شد: {merged_groups} گروه، "
            f"{moved_to_history} رکورد → سابقه، "
            f"{assigned_empty} اندیکاتور خالی تکمیل، "
            f"{normalized_nids} کد ملی نرمال شد."
        ),
        "merged_groups": merged_groups,
        "moved_to_history": moved_to_history,
        "assigned_empty_indicators": assigned_empty,
        "normalized_national_ids": normalized_nids,
        "samples": samples,
    }


# ──────────────────────────────────────────────
# مدیریت ردیف‌های نامعتبر ایمپورت (Invalid Imports)
# ──────────────────────────────────────────────
@app.get("/api/invalid-imports")
def list_invalid_imports(_user=Depends(get_current_user)):
    with get_db() as conn:
        rows = conn.execute(
            """
            SELECT id, excel_row_num, source_file, validation_errors, created_at,
                   row_num, first_name, last_name, father_name, national_id, active_status
            FROM invalid_imports
            ORDER BY id DESC
            """
        ).fetchall()
        result = []
        for r in rows:
            d = row_to_dict(r)
            try:
                d["errors_list"] = json.loads(d.get("validation_errors") or "[]")
            except Exception:
                d["errors_list"] = [d.get("validation_errors")] if d.get("validation_errors") else []
            result.append(d)
        return result


@app.get("/api/invalid-imports/{iid}")
def get_invalid_import(iid: int, _user=Depends(get_current_user)):
    with get_db() as conn:
        row = conn.execute("SELECT * FROM invalid_imports WHERE id=?", (iid,)).fetchone()
        if not row:
            raise HTTPException(404, "ردیف نامعتبر یافت نشد")
        d = row_to_dict(row)
        try:
            d["errors_list"] = json.loads(d.get("validation_errors") or "[]")
        except Exception:
            d["errors_list"] = [d.get("validation_errors")] if d.get("validation_errors") else []
        return d


@app.post("/api/invalid-imports/{iid}/resolve")
async def resolve_invalid_import(iid: int, request: Request, _admin=Depends(require_perm("invalid_imports"))):
    """
    اعتبارسنجی مجدد و انتقال رکورد اصلاح‌شده از invalid_imports به personnel
    """
    data = await request.json()
    labels = {name: label for name, label in PERSONAL_FIELDS}

    # اعتبارسنجی تاریخ‌ها
    norm_data = dict(data)
    for field in DATE_FIELDS:
        if field in norm_data and norm_data[field] not in (None, ""):
            try:
                norm_data[field] = normalize_jalali_value(
                    norm_data[field], field_label=labels.get(field, field), strict=True
                )
            except ValueError as exc:
                raise HTTPException(400, str(exc))

    fields = [f for f, _ in PERSONAL_FIELDS]
    values = [_norm_field(norm_data.get(f, ""), f) for f in fields]
    first_name = values[1]
    last_name = values[2]
    if not first_name and not last_name:
        raise HTTPException(400, "نام و نام خانوادگی الزامی است")

    national_id = _validate_national_id(norm_data.get("national_id", ""))
    indicator_num = _normalize_indicator(norm_data.get("indicator_num", ""))
    try:
        nid_idx = fields.index("national_id")
        values[nid_idx] = national_id
    except ValueError:
        pass
    try:
        ind_idx = fields.index("indicator_num")
        values[ind_idx] = indicator_num
    except ValueError:
        pass

    data_dict = {f: values[i] for i, f in enumerate(fields)}

    with get_db() as conn:
        inv = conn.execute("SELECT * FROM invalid_imports WHERE id=?", (iid,)).fetchone()
        if not inv:
            raise HTTPException(404, "رکورد نامعتبر یافت نشد")

        source_file = inv["source_file"] or "import-resolved"
        personnel_id, msg, _ = _register_person_record(
            conn, data_dict, source_file=source_file
        )
        conn.execute("DELETE FROM invalid_imports WHERE id=?", (iid,))

    return {"message": msg, "id": personnel_id}


@app.delete("/api/invalid-imports/{iid}")
def delete_invalid_import(iid: int, _admin=Depends(require_perm("invalid_imports"))):
    with get_db() as conn:
        exists = conn.execute("SELECT id FROM invalid_imports WHERE id=?", (iid,)).fetchone()
        if not exists:
            raise HTTPException(404, "ردیف یافت نشد")
        conn.execute("DELETE FROM invalid_imports WHERE id=?", (iid,))
    return {"message": "ردیف نامعتبر حذف شد"}


@app.delete("/api/invalid-imports")
def clear_all_invalid_imports(_admin=Depends(require_perm("invalid_imports"))):
    with get_db() as conn:
        conn.execute("DELETE FROM invalid_imports")
    return {"message": "تمام ردیف‌های نامعتبر پاکسازی شدند"}


# ──────────────────────────────────────────────
# ایمپورت اکسل (مستقل برای هر ردیف)
# ──────────────────────────────────────────────
EXPECTED_IMPORT_FIELDS = {f: label for f, label in PERSONAL_FIELDS}

def _excel_header_map(headers):
    column_map = {}
    duplicates = []
    for idx, header in enumerate(headers):
        if not header:
            continue
        norm = normalize_text(header)
        field_name = NORMALIZED_MAPPING.get(norm)
        if not field_name or field_name not in EXPECTED_IMPORT_FIELDS:
            continue
        if field_name in column_map.values():
            duplicates.append(header)
            continue
        column_map[idx] = field_name
    return column_map, duplicates

@app.post("/api/import/excel")
async def import_excel(file: UploadFile = File(...), _admin=Depends(require_perm("import_excel"))):
    content = await file.read()
    try:
        wb = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        ws = wb.active
        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
        column_map, duplicates = _excel_header_map(headers)

        if not column_map:
            raise HTTPException(400, "هیچ‌کدام از ستون‌های فایل با ساختار پرسنل مطابقت ندارد")

        missing = [label for field, label in PERSONAL_FIELDS if field not in column_map.values()]
        if missing:
            raise HTTPException(400, "ستون‌های زیر در فایل وجود ندارند: " + "، ".join(missing))

        valid_records = []
        invalid_records = []

        for excel_row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            record = {}
            row_errors = []
            has_any_content = False

            for col_idx, field_name in column_map.items():
                raw_val = row[col_idx] if col_idx < len(row) else None
                if raw_val is None or str(raw_val).strip() == "":
                    record[field_name] = ""
                    continue

                has_any_content = True
                if field_name in DATE_FIELDS:
                    try:
                        record[field_name] = normalize_jalali_value(
                            raw_val,
                            field_label=EXPECTED_IMPORT_FIELDS[field_name],
                            strict=True,
                        )
                    except ValueError as exc:
                        record[field_name] = str(raw_val).strip()
                        row_errors.append({
                            "field": EXPECTED_IMPORT_FIELDS[field_name],
                            "value": str(raw_val),
                            "message": str(exc),
                        })
                elif field_name == "national_id":
                    record[field_name] = _normalize_national_id(raw_val)
                else:
                    record[field_name] = str(raw_val).strip()

            if not has_any_content:
                continue

            # اعتبارسنجی نام و نام خانوادگی
            fn = str(record.get("first_name", "") or "").strip()
            ln = str(record.get("last_name", "") or "").strip()
            if not fn and not ln:
                row_errors.append({
                    "field": "نام و نام خانوادگی",
                    "value": "",
                    "message": "نام یا نام خانوادگی باید پر باشد",
                })

            # اعتبارسنجی کد ملی
            nid = _normalize_national_id(record.get("national_id", ""))
            record["national_id"] = nid
            if not nid:
                row_errors.append({
                    "field": "کد ملی",
                    "value": "",
                    "message": "کد ملی الزامی است",
                })
            elif not nid.isdigit() or len(nid) != 10:
                row_errors.append({
                    "field": "کد ملی",
                    "value": nid,
                    "message": "کد ملی باید دقیقاً ۱۰ رقم باشد",
                })

            record["_excel_row"] = excel_row_num
            if row_errors:
                record["_errors"] = row_errors
                invalid_records.append(record)
            else:
                valid_records.append(record)

        wb.close()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, f"خطا در خواندن فایل: {e}")

    if not valid_records and not invalid_records:
        return {"success": 0, "invalid": 0, "message": "هیچ رکورد پرسنلی یافت نشد"}

    success_count = 0
    invalid_count = 0
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    src_file_name = file.filename or "import"

    history_count = 0
    with get_db() as conn:
        # کلیدهای دیده‌شده در همین فایل: (کد ملی, اندیکاتور)
        seen_keys_in_file = set()

        # ۱. ذخیره ردیف‌های معتبر
        # کد ملی جدید → اصلی؛ کد ملی تکراری → سابقه سپس جابجایی با اصلی
        for record in valid_records:
            excel_row = record.pop("_excel_row", None)
            for field, _ in PERSONAL_FIELDS:
                if field in record and field not in DATE_FIELDS:
                    if field == "national_id":
                        record[field] = _normalize_national_id(record[field]) or ""
                    elif field == "indicator_num":
                        record[field] = _normalize_indicator(record[field]) or ""
                    else:
                        record[field] = _norm_field(record[field], field)

            national_id = record.get("national_id") or ""
            indicator_num = record.get("indicator_num") or ""
            key = (national_id, indicator_num)

            if national_id and key in seen_keys_in_file:
                inv_err = [{
                    "field": "کد ملی + اندیکاتور",
                    "value": f"{national_id} / {indicator_num}",
                    "message": "این ترکیب کد ملی و شماره اندیکاتور در همین فایل تکراری است — ثبت نشد",
                }]
                record["_excel_row"] = excel_row
                record["_errors"] = inv_err
                invalid_records.append(record)
                continue

            if national_id:
                seen_keys_in_file.add(key)

            try:
                _pid, _msg, promoted = _register_person_record(
                    conn, record, source_file=src_file_name
                )
                success_count += 1
                if promoted:
                    history_count += 1
            except HTTPException as he:
                inv_err = [{
                    "field": "کد ملی + اندیکاتور",
                    "value": f"{national_id} / {indicator_num}",
                    "message": str(he.detail),
                }]
                record["_excel_row"] = excel_row
                record["_errors"] = inv_err
                invalid_records.append(record)
            except Exception as exc:
                inv_err = [{
                    "field": "کد ملی + اندیکاتور",
                    "value": f"{national_id} / {indicator_num}",
                    "message": f"خطا در ثبت: {exc}",
                }]
                record["_excel_row"] = excel_row
                record["_errors"] = inv_err
                invalid_records.append(record)

        # ۲. ذخیره ردیف‌های نامعتبر در SQLite
        for record in invalid_records:
            excel_row = record.pop("_excel_row", 0)
            errors_json = json.dumps(record.pop("_errors", []), ensure_ascii=False)
            inv_vals = [excel_row, src_file_name, errors_json, now_str, ""]
            for f, _ in PERSONAL_FIELDS:
                inv_vals.append(str(record.get(f, "") or "").strip())

            cols = ["excel_row_num", "source_file", "validation_errors", "created_at", "photo"] + [f for f, _ in PERSONAL_FIELDS]
            placeholders = ", ".join(["?"] * len(inv_vals))
            conn.execute(
                f"INSERT INTO invalid_imports ({', '.join(cols)}) VALUES ({placeholders})",
                inv_vals,
            )
            invalid_count += 1

    msg = f"ایمپورت انجام شد: {success_count} ردیف ذخیره شد"
    if history_count:
        msg += f" (از این تعداد {history_count} مورد با کد ملی تکراری ثبت و به‌عنوان رکورد اصلی جابجا شد)"
    msg += "."
    if invalid_count > 0:
        msg += f" {invalid_count} ردیف دارای خطای اعتبارسنجی شناسایی و در جدول ردیف‌های نامعتبر ذخیره شدند."

    return {
        "success": success_count,
        "invalid": invalid_count,
        "message": msg,
    }


# ──────────────────────────────────────────────
# اکسپورت اکسل
# ──────────────────────────────────────────────
def _build_personnel_excel(rows) -> bytes:
    """ساخت فایل اکسل از ردیف‌های پرسنل."""
    wb = Workbook()
    ws = wb.active
    ws.title = "پرسنل"
    headers = [label for _, label in PERSONAL_FIELDS] + ["عكس"]
    ws.append(headers)
    for rec in rows:
        values = list(rec)
        ws.append(values)
    for i in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@app.get("/api/export/excel")
def export_excel(_user=Depends(require_perm("export_excel"))):
    with get_db() as conn:
        personal_fields = [f for f, _ in PERSONAL_FIELDS]
        rows = conn.execute(
            f"SELECT {', '.join(personal_fields)}, photo FROM personnel ORDER BY id"
        ).fetchall()
        if not rows:
            raise HTTPException(400, "دیتابیس خالی است")

        data = _build_personnel_excel(rows)
        filename = f"personnel_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return create_temp_download(
            data,
            filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


@app.post("/api/export/excel/filtered")
async def export_excel_filtered(request: Request, _user=Depends(require_perm("export_excel"))):
    """خروجی اکسل فقط برای شناسه‌های فیلترشده/انتخاب‌شده.

    بدنه: { "ids": [1, 2, 3] }
    """
    try:
        payload = await request.json()
    except Exception:
        raise HTTPException(400, "داده ارسالی نامعتبر است")
    raw_ids = payload.get("ids") if isinstance(payload, dict) else None
    if not isinstance(raw_ids, list) or not raw_ids:
        raise HTTPException(400, "لیست فیلترشده خالی است")
    try:
        ids = [int(x) for x in raw_ids]
    except (TypeError, ValueError):
        raise HTTPException(400, "شناسه‌های پرسنل نامعتبر است")
    # حذف تکراری با حفظ ترتیب
    seen = set()
    ordered = []
    for i in ids:
        if i not in seen:
            seen.add(i)
            ordered.append(i)
    if len(ordered) > 20000:
        raise HTTPException(400, "تعداد ردیف‌ها بیش از حد مجاز است (حداکثر ۲۰۰۰۰)")

    with get_db() as conn:
        personal_fields = [f for f, _ in PERSONAL_FIELDS]
        placeholders = ",".join("?" * len(ordered))
        # حفظ ترتیب فیلتر فرانت
        rows_by_id = {}
        for rec in conn.execute(
            f"SELECT id, {', '.join(personal_fields)}, photo FROM personnel WHERE id IN ({placeholders})",
            ordered,
        ).fetchall():
            rows_by_id[int(rec[0])] = rec[1:]  # بدون id برای سازگاری با هدر

        rows = []
        for pid in ordered:
            if pid in rows_by_id:
                rows.append(rows_by_id[pid])
        if not rows:
            raise HTTPException(400, "هیچ رکوردی از لیست فیلترشده در دیتابیس یافت نشد")

        data = _build_personnel_excel(rows)
        filename = f"personnel_filtered_{len(rows)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return create_temp_download(
            data,
            filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )




@app.post("/api/personnel/{pid}/addiction-letter")
async def issue_addiction_letter(pid: int, request: Request, user=Depends(require_perm("card_issue"))):
    """صدور برگه معرفی عدم اعتیاد (PDF).

    بدنه:
      recipient: خطاب نامه / سازمان مقصد (الزامی)
      letter_no: شماره نامه سال/اندیکاتور (الزامی)
      letter_date: تاریخ جلالی نامه (اختیاری)
      validity_days: مدت اعتبار به روز (پیش‌فرض ۳)
    """
    try:
        payload = await request.json()
    except Exception:
        payload = {}
    if not isinstance(payload, dict):
        payload = {}

    letter_date = (payload.get("letter_date") or "").strip() or None
    letter_no = (payload.get("letter_no") or "").strip() or None
    recipient = (payload.get("recipient") or "").strip() or None
    try:
        validity_days = int(payload.get("validity_days") or 3)
    except (TypeError, ValueError):
        validity_days = 3
    if validity_days < 1 or validity_days > 365:
        raise HTTPException(400, "مدت اعتبار نامه باید بین ۱ تا ۳۶۵ روز باشد")

    if not recipient:
        raise HTTPException(400, "خطاب نامه (سازمان مقصد) الزامی است")
    if not letter_no:
        raise HTTPException(400, "شماره نامه الزامی است")

    with get_db() as conn:
        row = conn.execute("SELECT * FROM personnel WHERE id=?", (pid,)).fetchone()
        if not row:
            raise HTTPException(404, "رکورد یافت نشد")
        person = row_to_dict(row)

    # مسیر عکس پرسنلی
    photo_path = None
    photo = (person.get("photo") or "").strip()
    if photo:
        candidate = os.path.join(APP_DIR, photo) if not os.path.isabs(photo) else photo
        if not os.path.isfile(candidate):
            candidate = os.path.join(PHOTOS_DIR, os.path.basename(photo))
        if os.path.isfile(candidate):
            photo_path = candidate

    try:
        pdf_bytes = generate_addiction_letter_pdf(
            person,
            letter_date=letter_date,
            letter_no=letter_no,
            recipient=recipient,
            photo_path=photo_path,
            validity_days=validity_days,
        )
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(500, f"خطا در تولید برگه عدم اعتیاد: {e}")

    first = (person.get("first_name") or "").strip()
    last = (person.get("last_name") or "").strip()
    nid = (person.get("national_id") or "").strip()
    ascii_name = re.sub(r"[^A-Za-z0-9_\-]+", "_", f"{nid or pid}_{first}_{last}".strip("_")) or str(pid)
    filename = f"addiction_letter_{ascii_name}.pdf"
    return create_temp_download(pdf_bytes, filename, "application/pdf")


# ──────────────────────────────────────────────
# صدور گروهی کارت تردد موقت — ۴ نفر در هر برگه A4 ایستاده (روی و پشت روبه‌روی هم)
# ──────────────────────────────────────────────
@app.post("/api/personnel/cards/group")
async def issue_group_cards(request: Request, user=Depends(require_perm("card_issue"))):
    """تولید گروهی کارت در A4 ایستاده؛ هر صفحه ۴ نفر، روی و پشت هر نفر روبه‌روی هم.

    بدنه درخواست:
      ids: لیست شناسه پرسنل
      issue_date: تاریخ صدور جلالی (اختیاری؛ پیش‌فرض تاریخ جاری سیستم)
      validity_date: تاریخ جلالی اعتبار (اختیاری)
      months: تعداد ماه پس از تاریخ صدور (اختیاری؛ جایگزین validity_date)
      is_one_year: اگر true باشد کادر تمدید اعتبار پشت کارت حذف می‌شود
    """
    try:
        payload = await request.json()
    except Exception:
        raise HTTPException(400, "داده ارسالی نامعتبر است")

    ids = payload.get("ids") if isinstance(payload, dict) else None
    if not isinstance(ids, list) or not ids:
        raise HTTPException(400, "حداقل یک پرسنل را انتخاب کنید")

    try:
        ids = list(dict.fromkeys(int(x) for x in ids))
    except (TypeError, ValueError):
        raise HTTPException(400, "شناسه پرسنل نامعتبر است")

    if len(ids) > 1000:
        raise HTTPException(400, "تعداد پرسنل انتخابی بیش از حد مجاز است")

    issue_date = (payload.get("issue_date") or "").strip() or None
    validity_date = (payload.get("validity_date") or "").strip() or None
    months_raw = payload.get("months")
    months = None
    if months_raw is not None and str(months_raw).strip() != "":
        try:
            months = int(months_raw)
        except (TypeError, ValueError):
            raise HTTPException(400, "تعداد ماه باید عدد صحیح باشد")
    is_one_year = bool(payload.get("is_one_year"))

    try:
        if issue_date:
            resolve_issue_date(issue_date)
        if validity_date or months is not None:
            resolve_validity_date(
                validity_date=validity_date,
                months=months,
                issue_date=issue_date,
            )
    except ValueError as exc:
        raise HTTPException(400, str(exc))

    with get_db() as conn:
        placeholders = ",".join("?" for _ in ids)
        rows = conn.execute(
            f"SELECT * FROM personnel WHERE id IN ({placeholders})", ids
        ).fetchall()
        by_id = {int(r["id"]): r for r in rows}

    missing = [pid for pid in ids if pid not in by_id]
    if missing:
        raise HTTPException(404, f"پرسنل با شناسه‌های {missing} یافت نشد")

    already_issued = [
        pid for pid in ids
        if (row_to_dict(by_id[pid]).get("card_issued") or "").strip()
    ]
    if already_issued:
        raise HTTPException(
            400,
            "برای این افراد قبلاً کارت صادر شده و صدور مجدد مجاز نیست "
            f"(شناسه‌ها: {already_issued}). در صورت نیاز مدیر وضعیت صدور را بازنشانی کند؛ "
            "برای کارت موقت از تمدید اعتبار استفاده کنید.",
        )

    items = []
    for pid in ids:
        person = row_to_dict(by_id[pid])
        photo_path = None
        photo = (person.get("photo") or "").strip()
        if photo:
            candidate = os.path.join(APP_DIR, photo) if not os.path.isabs(photo) else photo
            if not os.path.exists(candidate):
                candidate = os.path.join(PHOTOS_DIR, os.path.basename(photo))
            if os.path.exists(candidate):
                photo_path = candidate
        items.append((person, photo_path))

    try:
        pdf_bytes = generate_group_cards_pdf(
            items,
            validity_date=validity_date,
            months=months,
            issue_date=issue_date,
            is_one_year=is_one_year,
        )
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(500, f"خطا در تولید کارت‌های گروهی: {e}")

    filename = f"group_cards_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    # علامت‌گذاری: حداقل یک‌بار کارت صادر شده + نوع کارت
    issued_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    card_type = "one_year" if is_one_year else "temp"
    with get_db() as conn:
        placeholders = ",".join("?" for _ in ids)
        conn.execute(
            f"UPDATE personnel SET card_issued=?, card_type=? WHERE id IN ({placeholders})",
            [issued_at, card_type, *ids],
        )
    # به‌جای ارسال مستقیم بایت‌ها، لینک موقت دانلود برمی‌گردانیم
    info = create_temp_download(pdf_bytes, filename, "application/pdf")
    info["card_issued"] = True
    info["card_type"] = card_type
    info["issued_ids"] = ids
    return info


# ──────────────────────────────────────────────
# صدور کارت تردد موقت
# ──────────────────────────────────────────────
@app.get("/api/personnel/{pid}/card")
def issue_temp_card(
    request: Request,
    pid: int,
    validity_date: Optional[str] = None,
    months: Optional[int] = None,
    issue_date: Optional[str] = None,
    is_one_year: bool = False,
    user=Depends(require_perm("card_issue")),
):
    """تولید PDF کارت تردد موقت (روی و پشت) بر اساس اطلاعات پرسنل.

    پارامترهای اختیاری:
      issue_date — تاریخ صدور جلالی (پیش‌فرض: تاریخ جاری سیستم)
      validity_date — تاریخ جلالی اعتبار (YYYY/MM/DD)
      months — تعداد ماه پس از تاریخ صدور (جایگزین validity_date)
      is_one_year — اگر true باشد کادر تمدید اعتبار پشت کارت حذف می‌شود
    """
    with get_db() as conn:
        row = conn.execute("SELECT * FROM personnel WHERE id=?", (pid,)).fetchone()
        if not row:
            raise HTTPException(404, "رکورد یافت نشد")
        person = row_to_dict(row)
        already = (person.get("card_issued") or "").strip()
        if already:
            raise HTTPException(
                400,
                "کارت این پرسنل قبلاً صادر شده است. برای صدور مجدد ابتدا مدیر باید وضعیت صدور را بازنشانی کند. "
                "برای کارت موقت از دکمه «تمدید اعتبار» استفاده کنید.",
            )

    photo_path = None
    photo = (person.get("photo") or "").strip()
    if photo:
        candidate = os.path.join(APP_DIR, photo) if not os.path.isabs(photo) else photo
        if not os.path.exists(candidate):
            candidate = os.path.join(PHOTOS_DIR, os.path.basename(photo))
        if os.path.exists(candidate):
            photo_path = candidate

    try:
        pdf_bytes = generate_card_pdf(
            person,
            photo_path=photo_path,
            validity_date=validity_date,
            months=months,
            issue_date=issue_date,
            is_one_year=is_one_year,
        )
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(500, f"خطا در تولید کارت: {e}")

    first = (person.get("first_name") or "").strip()
    last = (person.get("last_name") or "").strip()
    nid = (person.get("national_id") or "").strip()
    # نام فایل فقط ASCII تا مرورگر PDF را درست دانلود کند
    ascii_name = re.sub(r"[^A-Za-z0-9_\-]+", "_", f"{nid or pid}_{first}_{last}".strip("_")) or str(pid)
    filename = f"card_{ascii_name}.pdf"

    # علامت‌گذاری: حداقل یک‌بار کارت صادر شده + نوع کارت
    issued_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    card_type = "one_year" if is_one_year else "temp"
    with get_db() as conn:
        if is_one_year and validity_date:
            conn.execute(
                "UPDATE personnel SET card_issued=?, card_type=?, one_year_validity=? WHERE id=?",
                (issued_at, card_type, validity_date, pid),
            )
        else:
            conn.execute(
                "UPDATE personnel SET card_issued=?, card_type=? WHERE id=?",
                (issued_at, card_type, pid),
            )

    # به‌جای ارسال مستقیم بایت‌ها، لینک موقت دانلود برمی‌گردانیم
    # تا مرورگر/IDM بدون هدر احراز هویت فایل را مستقیم بگیرند
    info = create_temp_download(pdf_bytes, filename, "application/pdf")
    info["card_issued"] = True
    info["card_type"] = card_type
    info["issued_ids"] = [pid]
    return info


@app.post("/api/personnel/{pid}/card-issued/reset")
def reset_card_issued(pid: int, _admin=Depends(require_admin)):
    """بازنشانی فلگ صدور کارت — فقط مدیر (رنگ آیکون به حالت اولیه برمی‌گردد)."""
    with get_db() as conn:
        row = conn.execute("SELECT id, card_issued FROM personnel WHERE id=?", (pid,)).fetchone()
        if not row:
            raise HTTPException(404, "رکورد یافت نشد")
        conn.execute(
            "UPDATE personnel SET card_issued='', card_type='' WHERE id=?",
            (pid,),
        )
    return {"message": "وضعیت صدور کارت بازنشانی شد", "card_issued": False, "card_type": ""}


@app.post("/api/personnel/{pid}/card/renew")
async def renew_temp_card(pid: int, request: Request, user=Depends(require_perm("card_issue"))):
    """تمدید اعتبار کارت صادرشده قبلی.

    بدنه:
      extension1: تاریخ جلالی تمدید اعتبار ۱ (الزامی برای کارت موقت)
      extension2: تاریخ جلالی تمدید اعتبار ۲ (اختیاری؛ فقط اگر extension1 باشد)
      validity_date: برای کارت یکساله — تاریخ اعتبار جدید
      issue_date: اختیاری
    """
    try:
        payload = await request.json()
    except Exception:
        raise HTTPException(400, "داده ارسالی نامعتبر است")

    with get_db() as conn:
        row = conn.execute("SELECT * FROM personnel WHERE id=?", (pid,)).fetchone()
        if not row:
            raise HTTPException(404, "رکورد یافت نشد")
        person = row_to_dict(row)
        issued = (person.get("card_issued") or "").strip()
        card_type = (person.get("card_type") or "").strip()
        if not issued:
            raise HTTPException(400, "ابتدا باید حداقل یک‌بار کارت صادر شده باشد")
        if card_type == "one_year":
            raise HTTPException(400, "تمدید اعتبار برای کارت یکساله مجاز نیست")
        if card_type != "temp":
            card_type = "temp"

    issue_date = (payload.get("issue_date") or "").strip() or None
    is_one_year = False

    extension1 = (payload.get("extension1") or "").strip() or None
    extension2 = (payload.get("extension2") or "").strip() or None

    # اگر قبلاً تاریخ اعتبار ۱ ذخیره شده باشد، همان را در کادر بالای پشت کارت نگه می‌داریم
    # تا مشخص باشد کارت یک‌بار تمدید شده؛ تاریخ جدید در کادر پایین (تمدید ۲) قرار می‌گیرد.
    prev_v1 = (person.get("validity1") or "").strip()
    if prev_v1 and prev_v1 != "—":
        # تمدید دوم (یا بیشتر): اعتبار ۱ قبلی حفظ می‌شود
        card_ext1 = prev_v1
        # تاریخ جدید از extension2 یا در صورت نبود از extension1 گرفته می‌شود
        new_date = extension2 or extension1
        if not new_date:
            raise HTTPException(400, "تاریخ تمدید اعتبار جدید (تمدید ۲) الزامی است")
        try:
            card_ext1 = normalize_jalali_value(card_ext1, field_label="تمدید اعتبار ۱", strict=True)
            new_date = normalize_jalali_value(new_date, field_label="تمدید اعتبار ۲", strict=True)
            ensure_jalali_after(
                card_ext1,
                new_date,
                earlier_label="تمدید اعتبار ۱",
                later_label="تمدید اعتبار ۲",
            )
        except ValueError as e:
            raise HTTPException(400, str(e))
        card_ext2 = new_date
        front_validity = new_date
        save_v1 = prev_v1
        save_v2 = new_date
    else:
        # اولین تمدید: رفتار قبلی
        if not extension1:
            raise HTTPException(400, "تاریخ تمدید اعتبار ۱ الزامی است")
        try:
            extension1 = normalize_jalali_value(extension1, field_label="تمدید اعتبار ۱", strict=True)
            if extension2:
                extension2 = normalize_jalali_value(extension2, field_label="تمدید اعتبار ۲", strict=True)
                ensure_jalali_after(
                    extension1,
                    extension2,
                    earlier_label="تمدید اعتبار ۱",
                    later_label="تمدید اعتبار ۲",
                )
        except ValueError as e:
            raise HTTPException(400, str(e))
        card_ext1 = extension1
        card_ext2 = extension2 if extension1 else None
        front_validity = extension2 or extension1
        save_v1 = extension1 or ""
        save_v2 = (extension2 if extension1 else "") or ""

    photo_path = None
    photo = (person.get("photo") or "").strip()
    if photo:
        candidate = os.path.join(APP_DIR, photo) if not os.path.isabs(photo) else photo
        if not os.path.exists(candidate):
            candidate = os.path.join(PHOTOS_DIR, os.path.basename(photo))
        if os.path.exists(candidate):
            photo_path = candidate

    try:
        pdf_bytes = generate_card_pdf(
            person,
            photo_path=photo_path,
            validity_date=front_validity,
            issue_date=issue_date,
            is_one_year=is_one_year,
            extension1=card_ext1,
            extension2=card_ext2,
        )
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(500, f"خطا در تولید کارت تمدیدی: {e}")

    # ذخیره تاریخ‌های تمدید روی رکورد پرسنل (اعتبار ۱ قبلی در صورت وجود حفظ می‌شود)
    issued_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_db() as conn:
        conn.execute(
            "UPDATE personnel SET card_issued=?, card_type=?, validity1=?, validity2=? WHERE id=?",
            (
                issued_at,
                "temp",
                save_v1,
                save_v2,
                pid,
            ),
        )

    first = (person.get("first_name") or "").strip()
    last = (person.get("last_name") or "").strip()
    nid = (person.get("national_id") or "").strip()
    ascii_name = re.sub(r"[^A-Za-z0-9_\-]+", "_", f"{nid or pid}_{first}_{last}".strip("_")) or str(pid)
    filename = f"card_renew_{ascii_name}.pdf"
    info = create_temp_download(pdf_bytes, filename, "application/pdf")
    info["card_issued"] = True
    info["card_type"] = card_type
    info["issued_ids"] = [pid]
    return info


# ──────────────────────────────────────────────
# دسترسی به مدارک (docs/) بر اساس شماره اندیکاتور
# ──────────────────────────────────────────────
def _normalize_indicator_key(val) -> str:
    """نرمال‌سازی شماره اندیکاتور برای مقایسه با نام پوشه."""
    if val is None:
        return ""
    text = str(val).strip()
    digits = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
    text = text.translate(digits).replace(" ", "").replace("-", "").replace("_", "")
    return text


def _collect_person_indicators(conn, pid: int) -> list:
    """اندیکاتور اصلی + اندیکاتورهای سوابق کاری."""
    indicators = []
    seen = set()
    row = conn.execute(
        "SELECT indicator_num FROM personnel WHERE id=?", (pid,)
    ).fetchone()
    if not row:
        return []
    main_ind = _normalize_indicator_key(row["indicator_num"] if "indicator_num" in row.keys() else row[0])
    if main_ind:
        indicators.append({"value": main_ind, "source": "main"})
        seen.add(main_ind)
    try:
        wh_rows = conn.execute(
            "SELECT DISTINCT indicator_num FROM work_history WHERE personnel_id=?",
            (pid,),
        ).fetchall()
        for r in wh_rows:
            ind = _normalize_indicator_key(r["indicator_num"] if hasattr(r, "keys") else r[0])
            if ind and ind not in seen:
                indicators.append({"value": ind, "source": "work_history"})
                seen.add(ind)
    except Exception:
        pass
    return indicators


def _safe_docs_resolve(rel_path: str) -> str:
    """مسیر نسبی داخل docs را به مسیر مطلق امن تبدیل می‌کند؛ خارج از docs ممنوع."""
    rel = (rel_path or "").replace("\\", "/").strip().lstrip("/")
    if not rel or ".." in rel.split("/"):
        raise HTTPException(400, "مسیر نامعتبر است")
    root = os.path.realpath(DOCS_DIR)
    full = os.path.realpath(os.path.join(DOCS_DIR, rel))
    if full != root and not full.startswith(root + os.sep):
        raise HTTPException(403, "دسترسی به این مسیر مجاز نیست")
    return full


def _scan_docs_for_indicators(indicator_set: set) -> list:
    """جستجوی بازگشتی همه پوشه‌ها و زیرپوشه‌های docs برای نام منطبق با اندیکاتورها."""
    matches = []
    if not os.path.isdir(DOCS_DIR):
        return matches
    root = os.path.realpath(DOCS_DIR)
    for dirpath, dirnames, filenames in os.walk(root):
        # نام خود پوشه جاری
        base = os.path.basename(dirpath)
        key = _normalize_indicator_key(base)
        if key and key in indicator_set:
            rel = os.path.relpath(dirpath, root).replace("\\", "/")
            if rel == ".":
                continue  # خود ریشه docs را نادیده بگیر
            file_count = 0
            subfolder_count = 0
            try:
                for _r, _ds, _fs in os.walk(dirpath):
                    file_count += len(_fs)
                    subfolder_count += len(_ds)
            except OSError:
                pass
            matches.append({
                "indicator": key,
                "folder_name": base,
                "relative_path": rel,
                "access_url": f"/api/docs/browse?path={rel}",
                "list_url": f"/api/docs/list?path={rel}",
                "file_count": file_count,
                "subfolder_count": subfolder_count,
            })
    # یکتا بر اساس relative_path
    seen_paths = set()
    unique = []
    for m in matches:
        if m["relative_path"] in seen_paths:
            continue
        seen_paths.add(m["relative_path"])
        unique.append(m)
    unique.sort(key=lambda x: (x["indicator"], x["relative_path"]))
    return unique


@app.get("/api/personnel/{pid}/docs-access")
def personnel_docs_access(pid: int, _user=Depends(get_current_user)):
    """جستجوی پوشه docs برای اندیکاتور اصلی و سوابق؛ برگرداندن لینک‌های دسترسی."""
    with get_db() as conn:
        exists = conn.execute("SELECT id, first_name, last_name, indicator_num FROM personnel WHERE id=?", (pid,)).fetchone()
        if not exists:
            raise HTTPException(404, "پرسنل یافت نشد")
        person = dict(exists)
        indicators = _collect_person_indicators(conn, pid)

    ind_set = {i["value"] for i in indicators if i.get("value")}
    matches = _scan_docs_for_indicators(ind_set) if ind_set else []

    # برچسب منبع برای هر match
    source_map = {i["value"]: i["source"] for i in indicators}
    for m in matches:
        m["source"] = source_map.get(m["indicator"], "unknown")

    return {
        "personnel_id": pid,
        "name": f"{(person.get('first_name') or '').strip()} {(person.get('last_name') or '').strip()}".strip(),
        "main_indicator": _normalize_indicator_key(person.get("indicator_num")),
        "indicators": indicators,
        "docs_root": "docs",
        "docs_exists": os.path.isdir(DOCS_DIR),
        "matches": matches,
        "match_count": len(matches),
    }


@app.get("/api/docs/list")
def docs_list_folder(path: str = "", _user=Depends(get_current_user)):
    """لیست فایل‌ها و زیرپوشه‌های یک مسیر داخل docs."""
    full = _safe_docs_resolve(path)
    if not os.path.exists(full):
        raise HTTPException(404, "مسیر یافت نشد")
    if not os.path.isdir(full):
        raise HTTPException(400, "مسیر یک پوشه نیست")
    items = []
    try:
        for name in sorted(os.listdir(full), key=lambda n: n.lower()):
            item_path = os.path.join(full, name)
            rel = os.path.relpath(item_path, os.path.realpath(DOCS_DIR)).replace("\\", "/")
            is_dir = os.path.isdir(item_path)
            size = None
            if not is_dir:
                try:
                    size = os.path.getsize(item_path)
                except OSError:
                    size = None
            items.append({
                "name": name,
                "relative_path": rel,
                "is_dir": is_dir,
                "size": size,
                "browse_url": f"/api/docs/list?path={rel}" if is_dir else None,
                "download_url": None if is_dir else f"/api/docs/file?path={rel}",
            })
    except OSError as e:
        raise HTTPException(500, f"خطا در خواندن پوشه: {e}")
    return {
        "path": path.replace("\\", "/").strip().lstrip("/"),
        "items": items,
        "count": len(items),
    }


@app.get("/api/docs/file")
def docs_download_file(
    path: str,
    inline: bool = False,
    _user=Depends(get_current_user),
):
    """آماده‌سازی دانلود/نمایش فایل docs با لینک موقت (الگوی کارت/اکسل).

    inline=true → Content-Disposition: inline (باز شدن در تب مرورگر)
    مرحله ۱ با توکن نشست: JSON شامل download_url
    مرحله ۲ بدون هدر احراز هویت: GET /api/downloads/{token}
    """
    full = _safe_docs_resolve(path)
    if not os.path.isfile(full):
        raise HTTPException(404, "فایل یافت نشد")
    filename = os.path.basename(full)
    return create_temp_download_from_path(
        full,
        filename=filename,
        disposition="inline" if inline else "attachment",
    )


@app.get("/api/docs/browse")
def docs_browse_info(path: str = "", _user=Depends(get_current_user)):
    """اطلاعات خلاصه یک پوشه تطبیق‌یافته + لیست سطح اول (همان list)."""
    return docs_list_folder(path=path, _user=_user)


# ──────────────────────────────────────────────
# پشتیبان‌گیری و بازگردانی دیتابیس (فقط مدیر)
# ──────────────────────────────────────────────
def _sqlite_consistent_copy(dest_path: str) -> None:
    """کپی یکپارچه از دیتابیس زنده با API پشتیبان SQLite."""
    if not os.path.isfile(DB_PATH):
        # ساخت دیتابیس خالی در صورت نبود
        init_db()
    dest_conn = None
    src_conn = None
    try:
        src_conn = sqlite3.connect(DB_PATH, timeout=60)
        dest_conn = sqlite3.connect(dest_path, timeout=60)
        src_conn.backup(dest_conn)
        dest_conn.commit()
    finally:
        if dest_conn is not None:
            try:
                dest_conn.close()
            except Exception:
                pass
        if src_conn is not None:
            try:
                src_conn.close()
            except Exception:
                pass


def _is_sqlite_file(path: str) -> bool:
    try:
        with open(path, "rb") as f:
            header = f.read(16)
        return header.startswith(b"SQLite format 3")
    except OSError:
        return False


@app.get("/api/admin/backup")
def admin_backup_database(_admin=Depends(require_admin)):
    """پشتیبان کامل: دیتابیس + عکس‌های پرسنلی در یک فایل ZIP."""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    tmp_dir = os.path.join(DOWNLOADS_DIR, f"backup_build_{secrets.token_urlsafe(8)}")
    os.makedirs(tmp_dir, exist_ok=True)
    try:
        db_copy = os.path.join(tmp_dir, "personnel.db")
        _sqlite_consistent_copy(db_copy)
        if not _is_sqlite_file(db_copy):
            raise HTTPException(500, "کپی دیتابیس نامعتبر است")

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.write(db_copy, arcname="personnel.db")
            # عکس‌های پرسنلی
            if os.path.isdir(PHOTOS_DIR):
                for root, _dirs, files in os.walk(PHOTOS_DIR):
                    for name in files:
                        full = os.path.join(root, name)
                        rel = os.path.relpath(full, APP_DIR).replace("\\", "/")
                        try:
                            zf.write(full, arcname=rel)
                        except OSError:
                            pass
            # متادیتای پشتیبان
            meta = {
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "app": "tasisat_web",
                "includes": ["personnel.db", "photos/"],
            }
            zf.writestr("backup_meta.json", json.dumps(meta, ensure_ascii=False, indent=2))

        data = zip_buf.getvalue()
        filename = f"tasisat_backup_{stamp}.zip"
        return create_temp_download(data, filename, "application/zip")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"خطا در تهیه پشتیبان: {e}")
    finally:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass


@app.post("/api/admin/restore")
async def admin_restore_database(
    file: UploadFile = File(...),
    _admin=Depends(require_admin),
):
    """بازگردانی از فایل ZIP پشتیبان یا فایل personnel.db.

    هشدار: داده‌های فعلی دیتابیس جایگزین می‌شوند.
    """
    raw_name = (file.filename or "").lower()
    content = await file.read()
    if not content:
        raise HTTPException(400, "فایل خالی است")

    tmp_dir = os.path.join(DOWNLOADS_DIR, f"restore_{secrets.token_urlsafe(8)}")
    os.makedirs(tmp_dir, exist_ok=True)
    restored_photos = 0
    try:
        candidate_db = None

        if raw_name.endswith(".zip") or content[:2] == b"PK":
            zip_path = os.path.join(tmp_dir, "upload.zip")
            with open(zip_path, "wb") as f:
                f.write(content)
            try:
                with zipfile.ZipFile(zip_path, "r") as zf:
                    # امنیت: جلوگیری از path traversal
                    for info in zf.infolist():
                        name = info.filename.replace("\\", "/")
                        if name.startswith("/") or ".." in name.split("/"):
                            raise HTTPException(400, f"مسیر نامعتبر در ZIP: {info.filename}")
                    zf.extractall(tmp_dir)
            except zipfile.BadZipFile:
                raise HTTPException(400, "فایل ZIP نامعتبر است")

            # پیدا کردن personnel.db داخل ZIP
            for root, _dirs, files in os.walk(tmp_dir):
                for n in files:
                    if n == "personnel.db" or n.endswith(".db"):
                        p = os.path.join(root, n)
                        if _is_sqlite_file(p):
                            candidate_db = p
                            break
                if candidate_db:
                    break
            if not candidate_db:
                raise HTTPException(400, "داخل ZIP فایل دیتابیس SQLite یافت نشد")

            # بازگردانی عکس‌ها در صورت وجود
            photos_src = None
            for root, dirs, _files in os.walk(tmp_dir):
                if "photos" in dirs:
                    photos_src = os.path.join(root, "photos")
                    break
            if photos_src and os.path.isdir(photos_src):
                os.makedirs(PHOTOS_DIR, exist_ok=True)
                for root, _dirs, files in os.walk(photos_src):
                    for n in files:
                        src = os.path.join(root, n)
                        rel = os.path.relpath(src, photos_src)
                        dest = os.path.join(PHOTOS_DIR, rel)
                        os.makedirs(os.path.dirname(dest), exist_ok=True)
                        try:
                            shutil.copy2(src, dest)
                            restored_photos += 1
                        except OSError:
                            pass
        else:
            # فرض: فایل خام sqlite
            candidate_db = os.path.join(tmp_dir, "personnel.db")
            with open(candidate_db, "wb") as f:
                f.write(content)
            if not _is_sqlite_file(candidate_db):
                raise HTTPException(400, "فایل دیتابیس SQLite معتبر نیست")

        # پشتیبان اضطراری از دیتابیس فعلی قبل از جایگزینی
        emergency = os.path.join(
            DOWNLOADS_DIR,
            f"pre_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
        )
        try:
            if os.path.isfile(DB_PATH):
                _sqlite_consistent_copy(emergency)
        except Exception:
            emergency = None

        # جایگزینی دیتابیس
        tmp_live = DB_PATH + ".restoring"
        shutil.copy2(candidate_db, tmp_live)
        # حذف wal/shm قدیمی تا با فایل جدید تداخل نکنند
        for suffix in ("-wal", "-shm"):
            side = DB_PATH + suffix
            if os.path.exists(side):
                try:
                    os.remove(side)
                except OSError:
                    pass
        os.replace(tmp_live, DB_PATH)

        # مهاجرت/ساخت جداول جدید در صورت نیاز
        try:
            init_db()
        except Exception as e:
            raise HTTPException(500, f"دیتابیس جایگزین شد ولی init_db خطا داد: {e}")

        return {
            "message": "بازگردانی با موفقیت انجام شد. صفحه را تازه کنید.",
            "restored_photos": restored_photos,
            "emergency_backup": os.path.basename(emergency) if emergency else None,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"خطا در بازگردانی: {e}")
    finally:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass


# ──────────────────────────────────────────────
# متادیتا فیلدها (برای فرانت)
# ──────────────────────────────────────────────
@app.get("/api/meta/fields")
def get_fields(_user=Depends(get_current_user)):
    return {
        "personal": [
            {"name": name, "label": label, "is_date": name in DATE_FIELDS}
            for name, label in PERSONAL_FIELDS
        ],
        "work": [
            {"name": name, "label": label, "is_date": name in DATE_FIELDS}
            for name, label in WORK_FIELDS
        ],
    }


@app.get("/api/meta/next-numbers")
def get_next_numbers(_user=Depends(get_current_user)):
    """آخرین شماره ردیف و اندیکاتور + ۱ و تاریخ صدور جاری"""
    with get_db() as conn:
        return {
            "row_num": _next_row_num(conn),
            "indicator_num": _next_indicator_num(conn),
            "temp_issue_date": _today_jalali_str(),
        }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
