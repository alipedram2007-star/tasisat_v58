# -*- coding: utf-8 -*-
"""مدیریت دیتابیس SQLite — سازگار با نسخه دسکتاپ"""
import os
import sqlite3
from contextlib import contextmanager

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "personnel.db")
PHOTOS_DIR = os.path.join(APP_DIR, "photos")

PERSONAL_FIELDS = [
    ("row_num", "ردیف"),
    ("first_name", "نام"),
    ("last_name", "نام خانوادگي"),
    ("father_name", "نام پدر"),
    ("national_id", "كد ملي"),
    ("indicator_num", "شماره انديكاتور"),
    ("company", "شركت"),
    ("position", "سمت"),
    ("issued_from", "صادره"),
    ("marital_status", "وضعيت تاهل"),
    ("children_count", "تعداد فرزندان"),
    ("native_status", "بومي/غير بومي"),
    ("yard", "يارد"),
    ("addiction_intro", "معرفي به عدم اعتياد"),
    ("result_date", "تاريخ دريافت جواب"),
    ("addiction_result", "نتيجه  عدم اعتياد"),
    ("three_signatures", "تاييد سه امضاء"),
    ("one_year_issue", "صدور يكساله"),
    ("one_year_validity", "اعتبار يكساله"),
    ("active_status", "فعال/غيرفعال"),
    ("temp_issue_date", "تاريخ صدور موقت"),
    ("validity1", "اعتبار 1"),
    ("validity2", "اعتبار 2"),
    ("validity3", "اعتبار 3"),
    ("validity4", "اعتبار4"),
    ("validity5", "اعتبار5"),
    ("validity6", "اعتبار 6"),
    ("validity7", "اعتبار7"),
    ("validity8", "اعتبار8"),
    ("validity9", "اعتبار9"),
    ("validity10", "اعتبار10"),
    ("address", "آدرس"),
    ("phone", "شماره تماس"),
    ("introducer", "معرف"),
    ("description", "توضيحات"),
]

# سابقه کاری: تمام فیلدهای پرسنل + تاریخ شروع/پایان و توضیحات کار
# هر رکورد پرسنل با شماره اندیکاتور متفاوت می‌تواند به عنوان یک سابقه ذخیره شود
WORK_FIELDS = [
    ("row_num", "ردیف"),
    ("first_name", "نام"),
    ("last_name", "نام خانوادگي"),
    ("father_name", "نام پدر"),
    ("national_id", "كد ملي"),
    ("indicator_num", "شماره انديكاتور"),
    ("company", "شركت"),
    ("position", "سمت"),
    ("issued_from", "صادره"),
    ("marital_status", "وضعيت تاهل"),
    ("children_count", "تعداد فرزندان"),
    ("native_status", "بومي/غير بومي"),
    ("yard", "يارد"),
    ("addiction_intro", "معرفي به عدم اعتياد"),
    ("result_date", "تاريخ دريافت جواب"),
    ("addiction_result", "نتيجه  عدم اعتياد"),
    ("three_signatures", "تاييد سه امضاء"),
    ("one_year_issue", "صدور يكساله"),
    ("one_year_validity", "اعتبار يكساله"),
    ("active_status", "فعال/غيرفعال"),
    ("temp_issue_date", "تاريخ صدور موقت"),
    ("validity1", "اعتبار 1"),
    ("validity2", "اعتبار 2"),
    ("validity3", "اعتبار 3"),
    ("validity4", "اعتبار4"),
    ("validity5", "اعتبار5"),
    ("validity6", "اعتبار 6"),
    ("validity7", "اعتبار7"),
    ("validity8", "اعتبار8"),
    ("validity9", "اعتبار9"),
    ("validity10", "اعتبار10"),
    ("address", "آدرس"),
    ("phone", "شماره تماس"),
    ("introducer", "معرف"),
    ("description", "توضيحات"),
    ("start_date", "تاريخ شروع"),
    ("end_date", "تاريخ پايان"),
    ("work_description", "توضيحات کار"),
]

DATE_FIELDS = {
    "result_date",
    "one_year_issue",
    "one_year_validity",
    "temp_issue_date",
    "validity1",
    "validity2",
    "validity3",
    "validity4",
    "validity5",
    "validity6",
    "validity7",
    "validity8",
    "validity9",
    "validity10",
    "start_date",
    "end_date",
}

COLUMN_MAPPING = {
    "ردیف": "row_num",
    "نام": "first_name",
    "نام خانوادگی": "last_name",
    "نام خانوادگي": "last_name",
    "نام پدر": "father_name",
    "کد ملی": "national_id",
    "كد ملي": "national_id",
    "شماره اندیکاتور": "indicator_num",
    "شماره انديكاتور": "indicator_num",
    "شرکت": "company",
    "شركت": "company",
    "سمت": "position",
    "صادره": "issued_from",
    "وضعیت تاهل": "marital_status",
    "وضعيت تاهل": "marital_status",
    "وضعیت تأهل": "marital_status",
    "وضعيت تأهل": "marital_status",
    "تعداد فرزندان": "children_count",
    "بومی/غیر بومی": "native_status",
    "بومی/غیربومی": "native_status",
    "بومي/غير بومي": "native_status",
    "یارد": "yard",
    "يارد": "yard",
    "معرفی به عدم اعتیاد": "addiction_intro",
    "معرفي به عدم اعتياد": "addiction_intro",
    "معرفی عدم اعتیاد": "addiction_intro",
    "تاریخ دریافت جواب": "result_date",
    "تاريخ دريافت جواب": "result_date",
    "نتیجه عدم اعتیاد": "addiction_result",
    "نتيجه  عدم اعتياد": "addiction_result",
    "تایید سه امضاء": "three_signatures",
    "تأیید سه امضاء": "three_signatures",
    "تاييد سه امضاء": "three_signatures",
    "صدور یکساله": "one_year_issue",
    "صدور يكساله": "one_year_issue",
    "اعتبار یکساله": "one_year_validity",
    "اعتبار يكساله": "one_year_validity",
    "فعال/غیرفعال": "active_status",
    "فعال/غيرفعال": "active_status",
    "تاریخ صدور موقت": "temp_issue_date",
    "تاريخ صدور موقت": "temp_issue_date",
    "اعتبار 1": "validity1", "اعتبار ۱": "validity1",
    "اعتبار2": "validity2", "اعتبار 2": "validity2", "اعتبار ۲": "validity2",
    "اعتبار3": "validity3", "اعتبار 3": "validity3", "اعتبار ۳": "validity3",
    "اعتبار4": "validity4", "اعتبار 4": "validity4", "اعتبار ۴": "validity4",
    "اعتبار5": "validity5", "اعتبار 5": "validity5", "اعتبار ۵": "validity5",
    "اعتبار6": "validity6", "اعتبار 6": "validity6", "اعتبار ۶": "validity6",
    "اعتبار7": "validity7", "اعتبار 7": "validity7", "اعتبار ۷": "validity7",
    "اعتبار8": "validity8", "اعتبار 8": "validity8", "اعتبار ۸": "validity8",
    "اعتبار9": "validity9", "اعتبار 9": "validity9", "اعتبار ۹": "validity9",
    "اعتبار10": "validity10", "اعتبار 10": "validity10", "اعتبار ۱۰": "validity10",
    "آدرس": "address", "ادرس": "address",
    "شماره تماس": "phone",
    "معرف": "introducer",
    "توضیحات": "description",
    "عكس": "photo", "عکس": "photo",
}


def normalize_text(text):
    """یکسان‌سازی نویسه‌های عربی/فارسی، ارقام و فاصله‌های نامرئی برای مقایسه و ذخیره.

    ي/ى → ی | ك → ک | ة/ۀ/ھ → ه | أ/إ/آ → ا | ارقام فارسی/عربی → انگلیسی
    """
    if text is None:
        return ""
    text = str(text).strip()
    if not text:
        return ""
    # ارقام فارسی و عربی → انگلیسی
    text = text.translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789"))
    # ی فارسی ← ي عربی، ى (ألف مقصورة)، ۍ
    text = text.replace("ي", "ی").replace("ى", "ی").replace("ۍ", "ی").replace("ے", "ی")
    # ک فارسی ← ك عربی
    text = text.replace("ك", "ک")
    # ه
    text = text.replace("ة", "ه").replace("ۀ", "ه").replace("ھ", "ه")
    # همزه و الف
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا").replace("ٱ", "ا")
    text = text.replace("ؤ", "و").replace("ئ", "ی")
    # فاصله‌های نامرئی
    text = (
        text.replace("‌", "")  # ZWNJ
        .replace("‍", "")  # ZWJ
        .replace("‎", "")  # LRM
        .replace("‏", "")  # RLM
        .replace("﻿", "")  # BOM
        .replace(" ", " ")  # NBSP
    )
    import re
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_for_search(text):
    """نرمال‌سازی عبارت جستجو (همان قوانین normalize_text)."""
    return normalize_text(text)


def normalize_active_status(value) -> str:
    """نرمال‌سازی وضعیت به یکی از: فعال | غیرفعال | مسدود | ''."""
    s = normalize_text(value)
    if not s or s in ("—", "-"):
        return ""
    compact = s.replace(" ", "")
    if compact == "مسدود":
        return "مسدود"
    if compact == "فعال":
        return "فعال"
    if "غیر" in compact and "فعال" in compact:
        return "غیرفعال"
    if compact in ("نافعال", "inactive", "disabled"):
        return "غیرفعال"
    if compact in ("active",):
        return "فعال"
    if compact in ("blocked", "block"):
        return "مسدود"
    return s


# عبارت SQL برای نرمال‌سازی ستون در جستجو (SQLite)
def sql_normalize_col(col: str) -> str:
    """تولید عبارت SQL که حروف عربی ستون را به فارسی تبدیل می‌کند."""
    expr = f"COALESCE({col}, '')"
    replacements = [
        ("ي", "ی"),
        ("ى", "ی"),
        ("ۍ", "ی"),
        ("ے", "ی"),
        ("ك", "ک"),
        ("ة", "ه"),
        ("ۀ", "ه"),
        ("ھ", "ه"),
        ("أ", "ا"),
        ("إ", "ا"),
        ("آ", "ا"),
        ("ٱ", "ا"),
        ("ؤ", "و"),
        ("ئ", "ی"),
        # ارقام فارسی
        ("۰", "0"), ("۱", "1"), ("۲", "2"), ("۳", "3"), ("۴", "4"),
        ("۵", "5"), ("۶", "6"), ("۷", "7"), ("۸", "8"), ("۹", "9"),
        # ارقام عربی
        ("٠", "0"), ("١", "1"), ("٢", "2"), ("٣", "3"), ("٤", "4"),
        ("٥", "5"), ("٦", "6"), ("٧", "7"), ("٨", "8"), ("٩", "9"),
    ]
    for src, dst in replacements:
        expr = f"REPLACE({expr}, '{src}', '{dst}')"
    expr = f"REPLACE({expr}, char(8204), '')"  # ZWNJ
    expr = f"REPLACE({expr}, char(8205), '')"  # ZWJ
    expr = f"REPLACE({expr}, char(65279), '')"  # BOM
    return expr


def migrate_arabic_persian_chars(conn) -> int:
    """یکسان‌سازی نویسه‌های عربی/فارسی در داده‌های موجود (یک‌بار در startup)."""
    updated = 0
    text_cols_personnel = [name for name, _ in PERSONAL_FIELDS]
    text_cols_work = [name for name, _ in WORK_FIELDS]

    def fix_table(table, cols, id_col="id"):
        nonlocal updated
        try:
            rows = conn.execute(f"SELECT {id_col}, {', '.join(cols)} FROM {table}").fetchall()
        except Exception:
            return
        for row in rows:
            rid = row[0]
            changes = {}
            for i, col in enumerate(cols):
                raw = row[i + 1]
                if raw is None or raw == "":
                    continue
                if col == "active_status":
                    fixed = normalize_active_status(raw)
                elif col in ("national_id",):
                    fixed = normalize_text(raw).replace(" ", "").replace("-", "")
                elif col in ("indicator_num", "row_num", "phone", "children_count"):
                    fixed = normalize_text(raw).replace(" ", "")
                else:
                    fixed = normalize_text(raw)
                if fixed != str(raw):
                    changes[col] = fixed
            if changes:
                sets = ", ".join(f"{c}=?" for c in changes)
                conn.execute(
                    f"UPDATE {table} SET {sets} WHERE {id_col}=?",
                    list(changes.values()) + [rid],
                )
                updated += 1

    fix_table("personnel", text_cols_personnel)
    try:
        fix_table("work_history", text_cols_work)
    except Exception:
        pass
    try:
        # side_notes description
        rows = conn.execute("SELECT id, description FROM side_notes").fetchall()
        for rid, desc in rows:
            if not desc:
                continue
            fixed = normalize_text(desc)
            if fixed != str(desc):
                conn.execute("UPDATE side_notes SET description=? WHERE id=?", (fixed, rid))
                updated += 1
    except Exception:
        pass
    return updated


NORMALIZED_MAPPING = {normalize_text(k): v for k, v in COLUMN_MAPPING.items()}


def get_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=30)
    conn.row_factory = sqlite3.Row
    # بهینه‌سازی سرعت SQLite
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    conn.execute("PRAGMA cache_size = -64000")  # ~64MB
    conn.execute("PRAGMA temp_store = MEMORY")
    conn.execute("PRAGMA mmap_size = 268435456")  # 256MB
    return conn


@contextmanager
def get_db():
    conn = get_connection()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    os.makedirs(PHOTOS_DIR, exist_ok=True)
    with get_db() as conn:
        c = conn.cursor()
        personal_sql = ", ".join([f"{name} TEXT" for name, _ in PERSONAL_FIELDS])
        c.execute(
            f"""
            CREATE TABLE IF NOT EXISTS personnel (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {personal_sql},
                photo TEXT DEFAULT '',
                created_at TEXT,
                source_file TEXT
            )
            """
        )
        c.execute("PRAGMA table_info(personnel)")
        cols = [r[1] for r in c.fetchall()]
        if "photo" not in cols:
            c.execute("ALTER TABLE personnel ADD COLUMN photo TEXT DEFAULT ''")
        # فلگ صدور کارت: خالی = هنوز صادر نشده؛ مقدار غیرخالی = حداقل یک‌بار صادر شده
        if "card_issued" not in cols:
            c.execute("ALTER TABLE personnel ADD COLUMN card_issued TEXT DEFAULT ''")
        # نوع کارت: temp | one_year
        if "card_type" not in cols:
            c.execute("ALTER TABLE personnel ADD COLUMN card_type TEXT DEFAULT ''")
        # مهاجرت از نسخه‌های قبلی: فیلدهای جدید مطابق فایل مرجع
        existing_personal = {r[1] for r in c.execute("PRAGMA table_info(personnel)").fetchall()}
        for name, _ in PERSONAL_FIELDS:
            if name not in existing_personal:
                c.execute(f"ALTER TABLE personnel ADD COLUMN {name} TEXT")
        # اطمینان مجدد بعد از مهاجرت فیلدهای شخصی
        existing_all = {r[1] for r in c.execute("PRAGMA table_info(personnel)").fetchall()}
        if "card_issued" not in existing_all:
            c.execute("ALTER TABLE personnel ADD COLUMN card_issued TEXT DEFAULT ''")
        if "card_type" not in existing_all:
            c.execute("ALTER TABLE personnel ADD COLUMN card_type TEXT DEFAULT ''")

        # هر کد ملی فقط یک رکورد اصلی دارد (اندیکاتور کوچکتر = اصلی)
        # رکوردهای با اندیکاتور بزرگتر در جدول work_history به‌عنوان سابقه ذخیره می‌شوند
        try:
            indexes = c.execute("PRAGMA index_list(personnel)").fetchall()
            for idx in indexes:
                idx_name = idx[1]
                if idx[2]:  # unique
                    idx_info = c.execute(f"PRAGMA index_info({idx_name})").fetchall()
                    cols_in_idx = [r[2] for r in idx_info]
                    # حذف یونیک ترکیبی قدیمی (national_id, indicator_num)
                    if cols_in_idx == ["national_id", "indicator_num"] or idx_name == "idx_personnel_nid_indicator":
                        c.execute(f"DROP INDEX IF EXISTS {idx_name}")
        except Exception:
            pass
        c.execute("DROP INDEX IF EXISTS idx_personnel_nid_indicator")
        c.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_personnel_national_id_unique "
            "ON personnel(national_id)"
        )

        work_sql = ", ".join([f"{name} TEXT" for name, _ in WORK_FIELDS])
        c.execute(
            f"""
            CREATE TABLE IF NOT EXISTS work_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                personnel_id INTEGER,
                {work_sql},
                FOREIGN KEY (personnel_id) REFERENCES personnel(id) ON DELETE CASCADE
            )
            """
        )
        # مهاجرت فیلدهای جدید سوابق کاری
        existing_work = {r[1] for r in c.execute("PRAGMA table_info(work_history)").fetchall()}
        for name, _ in WORK_FIELDS:
            if name not in existing_work:
                c.execute(f"ALTER TABLE work_history ADD COLUMN {name} TEXT")

        # جدول ردیف‌های نامعتبر ایمپورت
        inv_sql = ", ".join([f"{name} TEXT" for name, _ in PERSONAL_FIELDS])
        c.execute(
            f"""
            CREATE TABLE IF NOT EXISTS invalid_imports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                excel_row_num INTEGER,
                source_file TEXT,
                validation_errors TEXT,
                created_at TEXT,
                photo TEXT DEFAULT '',
                {inv_sql}
            )
            """
        )

        # ایندکس‌ها برای سرعت جستجو و لیست
        c.execute("CREATE INDEX IF NOT EXISTS idx_personnel_national_id ON personnel(national_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_personnel_indicator ON personnel(indicator_num)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_personnel_last_name ON personnel(last_name)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_personnel_first_name ON personnel(first_name)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_personnel_phone ON personnel(phone)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_personnel_active ON personnel(active_status)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_work_personnel ON work_history(personnel_id)")

        # سو سابقه — یادداشت‌های مستقل از سوابق کاری (هیچ ارتباطی با work_history ندارد)
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS side_notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                personnel_id INTEGER NOT NULL,
                description TEXT DEFAULT '',
                created_at TEXT,
                updated_at TEXT,
                FOREIGN KEY (personnel_id) REFERENCES personnel(id) ON DELETE CASCADE
            )
            """
        )
        c.execute("CREATE INDEX IF NOT EXISTS idx_side_notes_personnel ON side_notes(personnel_id)")

        # ─── کاربران سیستم ───
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS app_users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                full_name TEXT DEFAULT '',
                is_active INTEGER NOT NULL DEFAULT 1,
                permissions TEXT DEFAULT '',
                created_at TEXT,
                updated_at TEXT
            )
            """
        )
        # مهاجرت ستون دسترسی‌های سفارشی
        user_cols = {r[1] for r in c.execute("PRAGMA table_info(app_users)").fetchall()}
        if "permissions" not in user_cols:
            c.execute("ALTER TABLE app_users ADD COLUMN permissions TEXT DEFAULT ''")
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS app_sessions (
                token TEXT PRIMARY KEY,
                user_id INTEGER NOT NULL,
                created_at TEXT,
                expires_at TEXT,
                FOREIGN KEY (user_id) REFERENCES app_users(id) ON DELETE CASCADE
            )
            """
        )
        c.execute("CREATE INDEX IF NOT EXISTS idx_sessions_user ON app_sessions(user_id)")

        # لاگ دسترسی کاربران
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS access_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                username TEXT,
                action TEXT NOT NULL,
                detail TEXT DEFAULT '',
                ip TEXT DEFAULT '',
                user_agent TEXT DEFAULT '',
                created_at TEXT NOT NULL
            )
            """
        )
        c.execute("CREATE INDEX IF NOT EXISTS idx_access_logs_created ON access_logs(created_at)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_access_logs_user ON access_logs(username)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_access_logs_action ON access_logs(action)")

        # یوزر پیش‌فرض مدیر
        try:
            n = migrate_arabic_persian_chars(conn)
            if n:
                print(f"[init_db] normalized arabic/persian chars in {n} rows")
        except Exception as e:
            print(f"[init_db] migrate_arabic_persian_chars skipped: {e}")

        _ensure_default_admin(conn)


# ──────────────────────────────────────────────
# احراز هویت و کاربران
# ──────────────────────────────────────────────
import hashlib
import secrets
from datetime import datetime, timedelta

ROLE_ADMIN = "admin"
ROLE_USER = "user"
ROLE_CUSTOM = "custom"
DEFAULT_ADMIN_USERNAME = "TasisatAdmin"
DEFAULT_ADMIN_PASSWORD = "Tasisat@1405"
SESSION_DAYS = 7

# دسترسی‌های قابل انتخاب برای کاربر سفارشی
PERMISSION_DEFS = [
    ("personnel_create", "ایجاد پرسنل جدید"),
    ("personnel_edit", "ویرایش پرسنل"),
    ("personnel_delete", "حذف پرسنل"),
    ("photo_manage", "آپلود و حذف عکس"),
    ("work_write", "مدیریت سوابق کاری (افزودن/ویرایش/حذف)"),
    ("import_excel", "ایمپورت اکسل"),
    ("export_excel", "خروجی اکسل"),
    ("card_issue", "صدور کارت (تکی و گروهی)"),
    ("invalid_imports", "مدیریت ردیف‌های نامعتبر"),
]
ALL_PERMISSION_KEYS = [k for k, _ in PERMISSION_DEFS]
VALID_ROLES = (ROLE_ADMIN, ROLE_USER, ROLE_CUSTOM)

ROLE_LABELS = {
    ROLE_ADMIN: "مدیر",
    ROLE_USER: "عادی",
    ROLE_CUSTOM: "سفارشی",
}


def parse_permissions(raw) -> list:
    """تبدیل رشته JSON دسترسی‌ها به لیست کلیدها"""
    if not raw:
        return []
    if isinstance(raw, list):
        return [p for p in raw if p in ALL_PERMISSION_KEYS]
    try:
        import json
        data = json.loads(raw) if isinstance(raw, str) else raw
        if isinstance(data, list):
            return [p for p in data if p in ALL_PERMISSION_KEYS]
    except Exception:
        pass
    return []


def serialize_permissions(perms) -> str:
    import json
    cleaned = [p for p in (perms or []) if p in ALL_PERMISSION_KEYS]
    return json.dumps(cleaned, ensure_ascii=False)


def user_has_permission(user: dict, perm: str) -> bool:
    """بررسی دسترسی: مدیر همه چیز دارد؛ کاربر عادی فقط card_issue؛ سفارشی طبق لیست"""
    if not user:
        return False
    role = user.get("role")
    if role == ROLE_ADMIN:
        return True
    if role == ROLE_USER:
        # کاربر عادی: فقط خواندن + صدور کارت
        return perm in ("card_issue", "export_excel")
    if role == ROLE_CUSTOM:
        perms = user.get("permissions") or []
        if isinstance(perms, str):
            perms = parse_permissions(perms)
        return perm in perms
    return False


def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120000)
    return f"{salt}${dk.hex()}"


def verify_password(password: str, password_hash: str) -> bool:
    try:
        salt, stored = password_hash.split("$", 1)
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120000)
        return secrets.compare_digest(dk.hex(), stored)
    except Exception:
        return False


def _ensure_default_admin(conn):
    row = conn.execute(
        "SELECT id FROM app_users WHERE username=?", (DEFAULT_ADMIN_USERNAME,)
    ).fetchone()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not row:
        conn.execute(
            """
            INSERT INTO app_users (username, password_hash, role, full_name, is_active, permissions, created_at, updated_at)
            VALUES (?, ?, ?, ?, 1, '', ?, ?)
            """,
            (
                DEFAULT_ADMIN_USERNAME,
                hash_password(DEFAULT_ADMIN_PASSWORD),
                ROLE_ADMIN,
                "مدیر سیستم",
                now,
                now,
            ),
        )
    # کاربران نمونه برای تست
    _ensure_user(
        conn,
        username="test",
        password="test",
        role=ROLE_USER,
        full_name="کاربر عادی تست",
        permissions=[],
    )
    _ensure_user(
        conn,
        username="test1",
        password="test1",
        role=ROLE_CUSTOM,
        full_name="کاربر سفارشی تست",
        permissions=list(ALL_PERMISSION_KEYS),
    )


def _ensure_user(conn, *, username: str, password: str, role: str, full_name: str, permissions: list):
    """ایجاد یا به‌روزرسانی کاربر نمونه (رمز و نقش و دسترسی‌ها)."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    perms_raw = serialize_permissions(permissions) if role == ROLE_CUSTOM else ""
    row = conn.execute(
        "SELECT id FROM app_users WHERE username=?", (username,)
    ).fetchone()
    if not row:
        conn.execute(
            """
            INSERT INTO app_users (username, password_hash, role, full_name, is_active, permissions, created_at, updated_at)
            VALUES (?, ?, ?, ?, 1, ?, ?, ?)
            """,
            (username, hash_password(password), role, full_name, perms_raw, now, now),
        )
    else:
        conn.execute(
            """
            UPDATE app_users
            SET password_hash=?, role=?, full_name=?, is_active=1, permissions=?, updated_at=?
            WHERE username=?
            """,
            (hash_password(password), role, full_name, perms_raw, now, username),
        )


def log_access(
    conn,
    *,
    action: str,
    username: str = "",
    user_id=None,
    detail: str = "",
    ip: str = "",
    user_agent: str = "",
):
    """ثبت یک رویداد دسترسی در جدول access_logs."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        conn.execute(
            """
            INSERT INTO access_logs (user_id, username, action, detail, ip, user_agent, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                user_id,
                (username or "")[:120],
                (action or "unknown")[:80],
                (detail or "")[:500],
                (ip or "")[:80],
                (user_agent or "")[:300],
                now,
            ),
        )
    except Exception:
        # لاگ نباید منطق اصلی را مختل کند
        pass


def create_session(conn, user_id: int) -> str:
    token = secrets.token_urlsafe(32)
    now = datetime.now()
    expires = now + timedelta(days=SESSION_DAYS)
    conn.execute(
        "INSERT INTO app_sessions (token, user_id, created_at, expires_at) VALUES (?, ?, ?, ?)",
        (
            token,
            user_id,
            now.strftime("%Y-%m-%d %H:%M:%S"),
            expires.strftime("%Y-%m-%d %H:%M:%S"),
        ),
    )
    return token


def delete_session(conn, token: str):
    conn.execute("DELETE FROM app_sessions WHERE token=?", (token,))


def get_user_by_token(conn, token: str):
    if not token:
        return None
    row = conn.execute(
        """
        SELECT u.id, u.username, u.role, u.full_name, u.is_active, u.permissions, s.expires_at
        FROM app_sessions s
        JOIN app_users u ON u.id = s.user_id
        WHERE s.token=?
        """,
        (token,),
    ).fetchone()
    if not row:
        return None
    if not row["is_active"]:
        return None
    try:
        exp = datetime.strptime(row["expires_at"], "%Y-%m-%d %H:%M:%S")
        if exp < datetime.now():
            conn.execute("DELETE FROM app_sessions WHERE token=?", (token,))
            return None
    except Exception:
        return None
    perms = parse_permissions(row["permissions"] if "permissions" in row.keys() else "")
    return {
        "id": row["id"],
        "username": row["username"],
        "role": row["role"],
        "full_name": row["full_name"] or "",
        "is_active": bool(row["is_active"]),
        "permissions": perms,
    }


def user_to_public(user: dict) -> dict:
    role = user.get("role") or ROLE_USER
    perms = user.get("permissions") or []
    if isinstance(perms, str):
        perms = parse_permissions(perms)
    return {
        "id": user["id"],
        "username": user["username"],
        "role": role,
        "role_label": ROLE_LABELS.get(role, role),
        "full_name": user.get("full_name") or "",
        "is_active": bool(user.get("is_active", 1)),
        "permissions": perms if role == ROLE_CUSTOM else [],
    }
