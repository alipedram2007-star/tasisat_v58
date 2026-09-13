# -*- coding: utf-8 -*-
"""تولید کارت تردد موقت ویژه پیمانکاران (روی و پشت) — PDF"""
import os
from io import BytesIO

from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4, A5
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.utils import ImageReader

import arabic_reshaper
from bidi.algorithm import get_display
import qrcode
import jdatetime

APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(APP_DIR, "static", "img", "ioec_logo.png")
if not os.path.exists(LOGO_PATH):
    LOGO_PATH = os.path.join(APP_DIR, "static", "img", "company_logo.png")

FONTS_DIR = os.path.join(APP_DIR, "static", "fonts")

def _find_font(*candidates):
    """جستجوی فونت در مسیرهای پروژه، ویندوز و لینوکس"""
    for p in candidates:
        if p and os.path.isfile(p):
            return p
    return None

# اولویت: فونت‌های همراه پروژه (کار روی ویندوز و لینوکس)
FONT_REG = _find_font(
    os.path.join(FONTS_DIR, "NotoSansArabic-Regular.ttf"),
    r"C:\Windows\Fonts\arial.ttf",
    r"C:\Windows\Fonts\tahoma.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/SlidesCarnival/google/Noto Sans Arabic/static/NotoSansArabic-Regular.ttf",
)
FONT_BOLD = _find_font(
    os.path.join(FONTS_DIR, "NotoSansArabic-Bold.ttf"),
    r"C:\Windows\Fonts\arialbd.ttf",
    r"C:\Windows\Fonts\tahomabd.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/SlidesCarnival/google/Noto Sans Arabic/static/NotoSansArabic-Bold.ttf",
    FONT_REG,
)
FONT_SEMI = _find_font(
    os.path.join(FONTS_DIR, "NotoSansArabic-SemiBold.ttf"),
    FONT_BOLD,
    FONT_REG,
)

# اندازه کارت
CARD_W = 105 * mm
CARD_H = 70 * mm

BLUE = HexColor("#1a3a6b")
BLUE_LIGHT = HexColor("#2b5aa0")
RED = HexColor("#c0392b")
GRAY_LINE = HexColor("#b0b0b0")
GRAY_TEXT = HexColor("#7f8c8d")


def _register_fonts():
    names = set(pdfmetrics.getRegisteredFontNames())
    if "NotoAr" in names:
        return
    if not FONT_REG:
        raise RuntimeError(
            "فونت فارسی یافت نشد. پوشه static/fonts را از پروژه کپی کنید "
            "یا فونت NotoSansArabic را در Windows\\Fonts نصب کنید."
        )
    pdfmetrics.registerFont(TTFont("NotoAr", FONT_REG))
    pdfmetrics.registerFont(TTFont("NotoArBold", FONT_BOLD or FONT_REG))
    pdfmetrics.registerFont(TTFont("NotoArSemi", FONT_SEMI or FONT_BOLD or FONT_REG))


def _to_persian_digits(text: str) -> str:
    """تبدیل ارقام انگلیسی به فارسی برای نمایش در PDF."""
    if text is None:
        return ""
    return str(text).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))


def rtl(text: str) -> str:
    if not text:
        return ""
    text = str(text).strip()
    try:
        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text



def _qr_image_reader(data: str, box_size: int = 6, border: int = 1) -> ImageReader | None:
    """ساخت QR Code از داده و برگرداندن ImageReader برای ReportLab."""
    data = (data or "").strip()
    if not data or data == "—":
        return None
    try:
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=box_size,
            border=border,
        )
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return ImageReader(buf)
    except Exception:
        return None


def _pick_date(*candidates):
    for v in candidates:
        if v and str(v).strip():
            return str(v).strip()
    return "—"


def _today_jalali() -> str:
    """تاریخ جاری سیستم به صورت جلالی YYYY/MM/DD"""
    return jdatetime.date.today().strftime("%Y/%m/%d")


def _add_months_jalali(base: str, months: int) -> str:
    """افزودن تعداد ماه به تاریخ جلالی YYYY/MM/DD و برگرداندن تاریخ جدید."""
    parts = str(base).strip().replace("-", "/").split("/")
    if len(parts) != 3:
        raise ValueError(f"تاریخ نامعتبر: {base}")
    y, m, d = map(int, parts)
    m += int(months)
    while m > 12:
        m -= 12
        y += 1
    while m < 1:
        m += 12
        y -= 1
    # محدود کردن روز به آخرین روز ماه مقصد
    for day in range(d, 0, -1):
        try:
            return jdatetime.date(y, m, day).strftime("%Y/%m/%d")
        except ValueError:
            continue
    return jdatetime.date(y, m, 1).strftime("%Y/%m/%d")


def _normalize_jalali_date(text: str, *, field_label: str = "تاریخ") -> str:
    """نرمال‌سازی و اعتبارسنجی تاریخ جلالی YYYY/MM/DD."""
    text = str(text or "").strip().replace("-", "/").replace(".", "/")
    if not text or text == "—":
        raise ValueError(f"{field_label} الزامی است")
    parts = text.split("/")
    if len(parts) != 3 or not all(p.isdigit() for p in parts):
        raise ValueError(f"فرمت {field_label} نامعتبر است (YYYY/MM/DD): {text}")
    y, m, d = map(int, parts)
    if 0 <= y <= 99:
        y = 1400 + y if y <= 95 else 1300 + y
    try:
        return jdatetime.date(y, m, d).strftime("%Y/%m/%d")
    except ValueError as exc:
        raise ValueError(f"{field_label} نامعتبر است: {text}") from exc


def resolve_issue_date(issue_date: str | None = None) -> str:
    """تاریخ صدور: اگر مشخص نشده، تاریخ جاری سیستم."""
    if issue_date and str(issue_date).strip() and str(issue_date).strip() != "—":
        return _normalize_jalali_date(issue_date, field_label="تاریخ صدور")
    return _today_jalali()


def resolve_validity_date(
    *,
    validity_date: str | None = None,
    months: int | None = None,
    issue_date: str | None = None,
    person: dict | None = None,
) -> str:
    """تعیین تاریخ اعتبار کارت.

    اولویت:
      1) validity_date صریح (جلالی)
      2) months → تاریخ صدور + ماه
      3) فیلدهای پرسنل (one_year_validity و ...)
    """
    if validity_date and str(validity_date).strip() and str(validity_date).strip() != "—":
        return _normalize_jalali_date(validity_date, field_label="تاریخ اعتبار")

    if months is not None:
        try:
            months_int = int(months)
        except (TypeError, ValueError) as exc:
            raise ValueError("تعداد ماه باید عدد صحیح باشد") from exc
        if months_int < 1 or months_int > 120:
            raise ValueError("تعداد ماه باید بین ۱ تا ۱۲۰ باشد")
        base = resolve_issue_date(issue_date)
        return _add_months_jalali(base, months_int)

    if person:
        return _pick_date(
            person.get("one_year_validity"),
            person.get("validity1"),
            person.get("validity2"),
            person.get("validity3"),
        )
    return "—"


def generate_card_pdf(
    person: dict,
    photo_path: str | None = None,
    *,
    validity_date: str | None = None,
    months: int | None = None,
    issue_date: str | None = None,
    is_one_year: bool = False,
    extension1: str | None = None,
    extension2: str | None = None,
) -> bytes:
    _register_fonts()
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(CARD_W, CARD_H))

    first_name = person.get("first_name") or ""
    last_name = person.get("last_name") or ""
    company = person.get("company") or ""
    position = person.get("position") or ""
    national_id = person.get("national_id") or ""
    indicator_num = person.get("indicator_num") or ""

    resolved_issue = resolve_issue_date(issue_date)
    resolved_validity = resolve_validity_date(
        validity_date=validity_date,
        months=months,
        issue_date=resolved_issue,
        person=person,
    )

    # نرمال‌سازی تاریخ‌های تمدید (در صورت ارسال)
    ext1 = None
    ext2 = None
    if extension1 and str(extension1).strip() and str(extension1).strip() != "—":
        ext1 = _normalize_jalali_date(extension1, field_label="تمدید اعتبار ۱")
    if extension2 and str(extension2).strip() and str(extension2).strip() != "—":
        # تمدید ۲ فقط وقتی تمدید ۱ وجود دارد معنی دارد
        if ext1:
            ext2 = _normalize_jalali_date(extension2, field_label="تمدید اعتبار ۲")
        else:
            raise ValueError("برای ثبت تمدید اعتبار ۲ ابتدا تمدید اعتبار ۱ را وارد کنید")

    _draw_front(
        c,
        first_name=first_name,
        last_name=last_name,
        company=company,
        position=position,
        national_id=national_id,
        indicator_num=indicator_num,
        validity_date=resolved_validity,
        photo_path=photo_path,
    )
    c.showPage()
    _draw_back(
        c,
        issue_date=resolved_issue,
        show_extension=not is_one_year,
        extension1=ext1,
        extension2=ext2,
    )
    c.save()
    return buf.getvalue()


def generate_group_cards_pdf(
    items: list[tuple[dict, str | None]],
    *,
    validity_date: str | None = None,
    months: int | None = None,
    issue_date: str | None = None,
    is_one_year: bool = False,
    extension1: str | None = None,
    extension2: str | None = None,
) -> bytes:
    """تولید گروهی کارت‌ها در برگه A4 ایستاده (portrait).

    در هر صفحه حداکثر ۴ نفر چاپ می‌شوند. برای هر نفر، روی کارت در سمت چپ
    و پشت همان کارت دقیقاً روبه‌روی آن در سمت راست قرار می‌گیرد
    (۴ ردیف × ۱ زوج روی+پشت).

    چیدمان هر صفحه:
        [روی ۱] [پشت ۱]
        [روی ۲] [پشت ۲]
        [روی ۳] [پشت ۳]
        [روی ۴] [پشت ۴]

    حاشیه ۲ سانتی‌متری از لبه کاغذ رعایت می‌شود؛ کارت‌ها متناسب
    کوچک می‌شوند. عکس و بارکد در طراحی کارت بزرگ‌تر شده‌اند.
    """
    _register_fonts()
    buf = BytesIO()

    # A4 ایستاده (عمودی)
    page_w, page_h = A4  # 210mm × 297mm
    c = canvas.Canvas(buf, pagesize=A4)

    # فاصله تا لبه کاغذ: ۲ سانتی‌متر از هر طرف → کارت کمی کوچک‌تر چاپ می‌شود
    edge_margin = 20 * mm
    per_page = 4
    row_gap = 3 * mm

    avail_w = page_w - 2 * edge_margin
    avail_h = page_h - 2 * edge_margin
    pair_natural_w = 2 * CARD_W  # روی + پشت در مقیاس ۱
    scale_w = avail_w / pair_natural_w
    scale_h = (avail_h - (per_page - 1) * row_gap) / (per_page * CARD_H)
    scale = min(scale_w, scale_h, 1.0)

    print_card_w = CARD_W * scale
    print_card_h = CARD_H * scale
    pair_w = print_card_w * 2

    total_h = per_page * print_card_h + (per_page - 1) * row_gap
    left_margin = (page_w - pair_w) / 2
    top_margin = (page_h - total_h) / 2

    # تاریخ صدور مشترک
    resolved_issue = resolve_issue_date(issue_date)

    # تاریخ اعتبار مشترک برای همه کارت‌های این نوبت
    shared_validity = None
    if (validity_date and str(validity_date).strip()) or months is not None:
        shared_validity = resolve_validity_date(
            validity_date=validity_date,
            months=months,
            issue_date=resolved_issue,
            person=None,
        )

    for start in range(0, len(items), per_page):
        batch = items[start:start + per_page]

        for slot, (person, photo_path) in enumerate(batch):
            y = page_h - top_margin - print_card_h - slot * (print_card_h + row_gap)
            pair_x = left_margin

            first_name = person.get("first_name") or ""
            last_name = person.get("last_name") or ""
            company = person.get("company") or ""
            position = person.get("position") or ""
            national_id = person.get("national_id") or ""
            indicator_num = person.get("indicator_num") or ""
            if shared_validity is not None:
                card_validity = shared_validity
            else:
                card_validity = resolve_validity_date(
                    issue_date=resolved_issue, person=person
                )

            # روی کارت — سمت چپ
            c.saveState()
            c.translate(pair_x, y)
            c.scale(scale, scale)
            _draw_front(
                c,
                first_name=first_name,
                last_name=last_name,
                company=company,
                position=position,
                national_id=national_id,
                indicator_num=indicator_num,
                validity_date=card_validity,
                photo_path=photo_path,
            )
            c.restoreState()

            # پشت همان کارت — روبه‌روی روی، سمت راست
            c.saveState()
            c.translate(pair_x + print_card_w, y)
            c.scale(scale, scale)
            _draw_back(
                c,
                issue_date=resolved_issue,
                show_extension=not is_one_year,
                extension1=extension1,
                extension2=extension2 if extension1 else None,
            )
            c.restoreState()

        # هر ۴ نفر در یک برگ A4 ایستاده؛ نفرات بعدی در برگ بعد
        c.showPage()

    c.save()
    return buf.getvalue()


def _draw_border(c):
    c.setFillColor(white)
    c.rect(0, 0, CARD_W, CARD_H, fill=1, stroke=0)
    c.setStrokeColor(BLUE)
    c.setLineWidth(2)
    c.rect(1.5 * mm, 1.5 * mm, CARD_W - 3 * mm, CARD_H - 3 * mm, fill=0, stroke=1)
    c.setLineWidth(0.5)
    c.setStrokeColor(BLUE_LIGHT)
    c.rect(2.5 * mm, 2.5 * mm, CARD_W - 5 * mm, CARD_H - 5 * mm, fill=0, stroke=1)


def _draw_front(c, *, first_name, last_name, company, position, national_id, indicator_num, validity_date, photo_path):
    _draw_border(c)

    # ── لوگو (گوشه راست بالا) ──
    logo_size = 13 * mm
    logo_x = CARD_W - 5 * mm - logo_size
    logo_y = CARD_H - 5 * mm - logo_size
    if os.path.exists(LOGO_PATH):
        try:
            c.drawImage(
                LOGO_PATH, logo_x, logo_y,
                width=logo_size, height=logo_size,
                mask="auto", preserveAspectRatio=True, anchor="c",
            )
        except Exception:
            pass

    # نام شرکت
    c.setFillColor(BLUE)
    c.setFont("NotoArBold", 7.5)
    c.drawRightString(logo_x - 2 * mm, CARD_H - 8 * mm, rtl("شرکت مهندسی و ساخت تاسیسات دریایی ایران"))
    c.setFont("NotoAr", 5)
    c.setFillColor(HexColor("#34495e"))
    c.drawRightString(logo_x - 2 * mm, CARD_H - 11.2 * mm, "IRANIAN OFFSHORE ENGINEERING AND")
    c.drawRightString(logo_x - 2 * mm, CARD_H - 13.5 * mm, "CONSTRUCTION COMPANY")

    # خط هدر
    c.setStrokeColor(BLUE)
    c.setLineWidth(0.9)
    c.line(4 * mm, CARD_H - 18 * mm, CARD_W - 4 * mm, CARD_H - 18 * mm)

    # عنوان
    c.setFillColor(BLUE_LIGHT)
    c.setFont("NotoArBold", 10.5)
    c.drawCentredString(CARD_W / 2 + 8 * mm, CARD_H - 22.5 * mm, rtl("کارت تردد موقت ویژه پیمانکاران"))
    c.setStrokeColor(BLUE_LIGHT)
    c.setLineWidth(0.5)
    c.line(CARD_W / 2 - 18 * mm, CARD_H - 24 * mm, CARD_W / 2 + 34 * mm, CARD_H - 24 * mm)

    # ── عکس (چپ) — بزرگ‌تر؛ عرض QR برابر عرض عکس ──
    # از فضای خالی سمت راست ستون عکس/بارکد استفاده می‌شود
    photo_w, photo_h = 22 * mm, 24 * mm
    photo_x = 4.5 * mm
    qr_size = photo_w  # عرض بارکد = عرض عکس
    # از پایین: نوار آبی → شماره اندیکاتور → QR → فاصله → عکس
    ind_y = 4.2 * mm
    qr_y = ind_y + 1.6 * mm
    photo_y = qr_y + qr_size + 0.5 * mm

    c.setStrokeColor(GRAY_LINE)
    c.setLineWidth(0.8)
    c.setFillColor(HexColor("#ecf0f1"))
    c.rect(photo_x, photo_y, photo_w, photo_h, fill=1, stroke=1)

    if photo_path and os.path.exists(photo_path):
        try:
            c.drawImage(
                photo_path,
                photo_x + 0.4 * mm, photo_y + 0.4 * mm,
                width=photo_w - 0.8 * mm, height=photo_h - 0.8 * mm,
                preserveAspectRatio=True, anchor="c",
            )
        except Exception:
            c.setFillColor(GRAY_TEXT)
            c.setFont("NotoAr", 6)
            c.drawCentredString(photo_x + photo_w / 2, photo_y + photo_h / 2 - 1.5, rtl("بدون عکس"))
    else:
        c.setFillColor(GRAY_TEXT)
        c.setFont("NotoAr", 6)
        c.drawCentredString(photo_x + photo_w / 2, photo_y + photo_h / 2 - 1.5, rtl("بدون عکس"))

    # ── جدول فیلدها (راست عکس) ──
    # ستون برچسب در سمت راست، مقدار در سمت چپ برچسب
    label_x = CARD_W - 5 * mm          # لبه راست برچسب‌ها
    line_left = photo_x + photo_w + 2.5 * mm
    line_right = CARD_W - 5 * mm
    value_center_x = (line_left + line_right) / 2

    fields = [
        ("نام", first_name or "—"),
        ("شهرت", last_name or "—"),
        ("شرکت مرتبط", company or "—"),
        ("سمت", position or "—"),
        ("کد ملی", national_id or "—"),
        ("تاریخ اعتبار", validity_date or "—"),
    ]

    row_h = 5.6 * mm
    start_y = CARD_H - 29 * mm

    for i, (label, value) in enumerate(fields):
        y = start_y - i * row_h

        # برچسب
        c.setFillColor(BLUE)
        c.setFont("NotoArSemi", 8)
        c.drawRightString(label_x, y, rtl(f"{label} :"))

        # مقدار — همه مقادیر دقیقاً وسط ستون مقدار قرار می‌گیرند
        c.setFillColor(RED if label == "تاریخ اعتبار" else black)
        c.setFont("NotoArBold" if label == "تاریخ اعتبار" else "NotoAr", 9)
        value_text = str(value) if label == "کد ملی" else rtl(str(value))
        c.drawCentredString(value_center_x, y, value_text)

    # ── QR Code زیر عکس (نام + شرکت + کد ملی + تاریخ اعتبار) — بدون شماره اندیکاتور ──
    name = " ".join(
        p for p in (str(first_name or "").strip(), str(last_name or "").strip()) if p
    ).strip()
    co = str(company or "").strip()
    if co == "—":
        co = ""
    nid = str(national_id or "").strip()
    exp = str(validity_date or "").strip()
    if exp in ("", "—"):
        exp = ""
    parts = []
    if name:
        parts.append(f"NAME:{name}")
    if co:
        parts.append(f"CO:{co}")
    if nid:
        parts.append(f"NID:{nid}")
    if exp:
        parts.append(f"EXP:{exp}")
    qr_data = "\n".join(parts) if parts else ""
    qr_reader = _qr_image_reader(qr_data)
    qr_x = photo_x  # هم‌تراز با لبه عکس (عرض یکسان)
    if qr_reader is not None:
        try:
            c.drawImage(
                qr_reader,
                qr_x, qr_y,
                width=qr_size, height=qr_size,
                mask="auto", preserveAspectRatio=True, anchor="c",
            )
        except Exception:
            pass

    # فقط مقدار شماره اندیکاتور (بدون عنوان) زیر QR
    c.setFillColor(black)
    c.setFont("NotoArBold", 7.5)
    c.drawCentredString(photo_x + photo_w / 2, ind_y, str(indicator_num or "—"))

    # نوار آبی پایین
    c.setFillColor(BLUE)
    c.rect(1.5 * mm, 1.5 * mm, CARD_W - 3 * mm, 2 * mm, fill=1, stroke=0)


def _draw_back(
    c,
    *,
    issue_date,
    show_extension: bool = True,
    extension1: str | None = None,
    extension2: str | None = None,
):
    _draw_border(c)

    # متن قانونی
    c.setFillColor(RED)
    c.setFont("NotoAr", 7)
    y = CARD_H - 11 * mm
    line1 = rtl("این کارت به منظور شناسایی جهت تردد در شرکت صادر گردیده و ارزش قانونی دیگری ندارد.")
    line2 = rtl("از یابنده تقاضا می‌شود کارت را به حراست شرکت IOEC واقع در اداره بندر خرمشهر تحویل نماید.")
    c.drawCentredString(CARD_W / 2, y, line1)
    c.drawCentredString(CARD_W / 2, y - 4.5 * mm, line2)

    # خط جدا
    c.setStrokeColor(GRAY_LINE)
    c.setLineWidth(0.5)
    c.line(10 * mm, y - 8 * mm, CARD_W - 10 * mm, y - 8 * mm)

    # تاریخ صدور — برچسب و تاریخ با فاصله مشخص (بدون همپوشانی)
    y_date = y - 15 * mm
    c.setFillColor(BLUE)
    c.setFont("NotoArSemi", 8)
    c.drawRightString(CARD_W / 2 + 22 * mm, y_date, rtl("تاریخ صدور"))
    c.setFillColor(black)
    c.setFont("NotoArBold", 7.5)
    # تاریخ (اعداد لاتین) در سمت چپ برچسب
    c.drawCentredString(CARD_W / 2 - 12 * mm, y_date, issue_date)

    # یک کادر تمدید اعتبار با برچسب بالای کادر و خط افقی میانی
    # (در کارت یکساله این بخش حذف می‌شود)
    if show_extension:
        box_w, box_h = 55 * mm, 18 * mm
        box_x = (CARD_W - box_w) / 2
        box_y = 18 * mm

        # عنوان «تمدید اعتبار» بالای کادر
        c.setFillColor(BLUE)
        c.setFont("NotoArSemi", 8)
        c.drawCentredString(CARD_W / 2, box_y + box_h + 2.2 * mm, rtl("تمدید اعتبار"))

        # کادر اصلی
        c.setStrokeColor(GRAY_LINE)
        c.setLineWidth(0.8)
        c.setFillColor(HexColor("#fafafa"))
        c.rect(box_x, box_y, box_w, box_h, fill=1, stroke=1)

        # خط افقی وسط — تقسیم کادر به دو بخش بالا و پایین
        mid_y = box_y + box_h / 2
        c.setStrokeColor(GRAY_LINE)
        c.setLineWidth(0.6)
        c.line(box_x, mid_y, box_x + box_w, mid_y)

        # تاریخ تمدید ۱ در نیمه بالا، تمدید ۲ در نیمه پایین (فقط اگر تمدید ۱ وجود داشته باشد)
        top_text_y = mid_y + (box_h / 4) - 1.2 * mm
        bot_text_y = box_y + (box_h / 4) - 1.2 * mm
        c.setFillColor(black)
        c.setFont("NotoArBold", 9)
        if extension1:
            c.drawCentredString(CARD_W / 2, top_text_y, str(extension1))
        if extension1 and extension2:
            c.drawCentredString(CARD_W / 2, bot_text_y, str(extension2))

        sig_y1, sig_y2 = 11 * mm, 6.5 * mm
    else:
        # بدون کادر تمدید: امضا کمی بالاتر قرار می‌گیرد
        sig_y1, sig_y2 = 22 * mm, 16 * mm

    # رئیس حراست
    c.setFillColor(BLUE)
    c.setFont("NotoArSemi", 8)
    c.drawCentredString(CARD_W / 2, sig_y1, rtl("رئیس حراست IOEC خرمشهر"))
    c.setFont("NotoArBold", 7.5)
    c.drawCentredString(CARD_W / 2, sig_y2, rtl("علی مبارکی"))

    c.setFillColor(BLUE)
    c.rect(1.5 * mm, 1.5 * mm, CARD_W - 3 * mm, 2 * mm, fill=1, stroke=0)



# ──────────────────────────────────────────────
# برگه معرفی عدم اعتیاد (نامه رسمی A5)
# ──────────────────────────────────────────────

def _fa_space_normalize(text: str) -> str:
    """فاصله‌گذاری استاندارد فارسی: نیم‌فاصله، علائم، فاصله‌های تکراری."""
    if not text:
        return ""
    import re
    s = str(text)
    # یکسان‌سازی فاصله و نیم‌فاصله
    s = s.replace(" ", " ").replace("‌", "‌")
    s = re.sub(r"[ \t\u2000-\u200b\u202f]+", " ", s)
    # فاصله قبل از علائم حذف، بعد از علائم یک فاصله
    s = re.sub(r"\s+([،؛:!؟\?\.])", r"\1", s)
    s = re.sub(r"([،؛:!؟\?\.])(\S)", r"\1 \2", s)
    # فاصله دور پرانتز
    s = re.sub(r"\(\s+", "(", s)
    s = re.sub(r"\s+\)", ")", s)
    return s.strip()


def _draw_rtl_paragraph(
    c,
    text: str,
    *,
    x_left: float,
    x_right: float,
    y: float,
    font_name: str = "NotoAr",
    font_bold: str = "NotoArBold",
    font_size: float = 11,
    leading: float = None,
    justify: bool = True,
    first_line_indent: float = 0,
    bold_phrases: list | None = None,
) -> float:
    """رسم پاراگراف فارسی با قالب‌بندی پیشرفته.

    - تراز دوطرفه (justify) برای خطوط میانی
    - تورفتگی خط اول (first_line_indent)
    - برجسته‌سازی عبارات مهم (bold_phrases)
    - فاصله‌گذاری یکنواخت واژه‌ها با سقف فاصله
    """
    if leading is None:
        leading = font_size * 1.9
    max_w = x_right - x_left
    if max_w <= 0:
        return y

    text = _fa_space_normalize(text)
    if not text:
        return y

    bold_set = []
    for p in (bold_phrases or []):
        p = _fa_space_normalize(str(p or ""))
        if p:
            bold_set.append(p)

    # توکن‌سازی: واژه‌ها + حفظ عبارات bold به‌صورت واحد در صورت امکان
    # ساده‌سازی: هر واژه؛ اگر واژه داخل یک bold phrase باشد bold می‌شود
    def is_bold_word(w: str) -> bool:
        for phrase in bold_set:
            if w in phrase.split() or w == phrase:
                return True
            # کد ملی / اعداد داخل عبارت
            if w.isdigit() and any(w in ph for ph in bold_set):
                return True
        return False

    words = text.split()
    if not words:
        return y

    def word_width(w: str, bold: bool) -> float:
        fn = font_bold if bold else font_name
        return c.stringWidth(rtl(w), fn, font_size)

    def line_width(ws: list) -> float:
        if not ws:
            return 0
        total = 0
        for idx, w in enumerate(ws):
            total += word_width(w, is_bold_word(w))
            if idx < len(ws) - 1:
                total += c.stringWidth(" ", font_name, font_size)
        return total

    # شکستن خطوط (عرض خط اول با تورفتگی کمتر)
    lines = []
    current = []
    for w in words:
        trial = current + [w]
        line_idx = len(lines)
        avail = max_w - (first_line_indent if line_idx == 0 and not current else 0)
        # وقتی current خالی است و خط اول است avail کم شده
        if not current and line_idx == 0:
            avail = max_w - first_line_indent
        elif not current:
            avail = max_w
        else:
            avail = max_w - (first_line_indent if line_idx == 0 else 0)

        # recalculate avail properly for trial on current line
        on_first = len(lines) == 0
        avail = max_w - (first_line_indent if on_first else 0)

        if line_width(trial) <= avail or not current:
            current = trial
        else:
            lines.append(current)
            current = [w]
    if current:
        lines.append(current)

    min_gap = c.stringWidth(" ", font_name, font_size)
    max_gap = min_gap * 3.2  # سقف فاصله برای جلوگیری از پخش بیش‌ازحد

    for i, line_words in enumerate(lines):
        is_last = i == len(lines) - 1
        on_first = i == 0
        indent = first_line_indent if on_first else 0
        avail = max_w - indent
        right_edge = x_right - indent

        bold_flags = [is_bold_word(w) for w in line_words]
        widths = [word_width(w, b) for w, b in zip(line_words, bold_flags)]
        total_words = sum(widths)
        gaps = len(line_words) - 1

        if justify and (not is_last) and gaps > 0:
            raw_gap = (avail - total_words) / gaps
            gap_w = max(min_gap, min(max_gap, raw_gap))
            # اگر فاصله سقف خورد، خط کمی جمع می‌شود (راست‌چین با gap محدود)
            used = total_words + gap_w * gaps
            start_right = right_edge if used >= avail - 0.5 else right_edge
            x = start_right
            for j, w in enumerate(line_words):
                sw = widths[j]
                x -= sw
                fn = font_bold if bold_flags[j] else font_name
                c.setFont(fn, font_size)
                c.drawString(x, y, rtl(w))
                if j < gaps:
                    x -= gap_w
        else:
            # راست‌چین ساده با فاصله استاندارد
            x = right_edge
            for j, w in enumerate(line_words):
                sw = widths[j]
                x -= sw
                fn = font_bold if bold_flags[j] else font_name
                c.setFont(fn, font_size)
                c.drawString(x, y, rtl(w))
                if j < gaps:
                    x -= min_gap
        y -= leading
    return y


def generate_addiction_letter_pdf(
    person: dict,
    *,
    letter_date: str | None = None,
    letter_no: str | None = None,
    recipient: str | None = None,
    photo_path: str | None = None,
    validity_days: int = 3,
) -> bytes:
    """تولید PDF برگه معرفی به آزمایش عدم اعتیاد (مطابق نمونه اداری IOEC)."""
    _register_fonts()

    first = (person.get("first_name") or "").strip()
    last = (person.get("last_name") or "").strip()
    father = (person.get("father_name") or "").strip()
    nid = (person.get("national_id") or "").strip()
    full_name = f"{first} {last}".strip() or "—"
    father_disp = father or "—"
    nid_disp = nid or "—"

    date_str = _normalize_jalali_date(letter_date, field_label="تاریخ نامه") if letter_date else _today_jalali()
    # شماره نامه: الزامی از کاربر؛ در غیر این صورت از اندیکاتور
    if letter_no and str(letter_no).strip():
        no_str = str(letter_no).strip()
    else:
        ind = (person.get("indicator_num") or "").strip()
        pid = person.get("id") or ""
        year = date_str.split("/")[0] if date_str else ""
        if year and ind:
            no_str = f"{year}/{ind}"
        else:
            no_str = f"{year}/{ind or pid}".strip("/")
    no_str = _to_persian_digits(no_str)

    recipient_text = (recipient or "").strip() or "بهداری محترم نیروی انتظامی شهرستان خرمشهر"
    try:
        validity_days_n = int(validity_days)
    except (TypeError, ValueError):
        validity_days_n = 3
    if validity_days_n < 1:
        validity_days_n = 1
    if validity_days_n > 365:
        validity_days_n = 365


    buf = BytesIO()
    # پیش‌فرض A5 ایستاده (148 × 210 میلی‌متر)
    page_w, page_h = A5
    c = canvas.Canvas(buf, pagesize=A5)

    margin_r = 12 * mm  # راست (شروع متن RTL)
    margin_l = 12 * mm
    usable_w = page_w - margin_l - margin_r

    # --- هدر لوگو و نام شرکت ---
    y = page_h - 12 * mm
    logo_size = 16 * mm
    if os.path.exists(LOGO_PATH):
        try:
            c.drawImage(
                LOGO_PATH,
                (page_w - logo_size) / 2,
                y - logo_size,
                width=logo_size,
                height=logo_size,
                mask="auto",
                preserveAspectRatio=True,
                anchor="c",
            )
        except Exception:
            pass
    y = y - logo_size - 4 * mm

    c.setFillColor(BLUE)
    c.setFont("NotoArBold", 9)
    c.drawCentredString(page_w / 2, y, rtl("شرکت مهندسی و ساخت تأسیسات دریایی ایران (سهامی خاص)"))
    y -= 4 * mm
    c.setFont("NotoAr", 7)
    c.setFillColor(HexColor("#333333"))
    c.drawCentredString(page_w / 2, y, "IRANIAN OFFSHORE ENGINEERING")
    y -= 3.5 * mm
    c.drawCentredString(page_w / 2, y, "AND CONSTRUCTION COMPANY")
    y -= 4 * mm

    # خط زیر هدر
    c.setStrokeColor(BLUE)
    c.setLineWidth(1.2)
    c.line(margin_l, y, page_w - margin_r, y)
    y -= 10 * mm

    # --- تاریخ و شماره (راست‌چین) + عکس پرسنل (سمت چپ) ---
    photo_w, photo_h = 22 * mm, 28 * mm  # نسبت تقریبی 3×4 — مناسب A5
    photo_x = margin_l
    photo_top = y
    photo_drawn = False
    if photo_path and os.path.isfile(photo_path):
        try:
            c.setStrokeColor(GRAY_LINE)
            c.setLineWidth(0.8)
            c.setFillColor(white)
            c.rect(photo_x, photo_top - photo_h, photo_w, photo_h, fill=1, stroke=1)
            c.drawImage(
                photo_path,
                photo_x + 0.6 * mm,
                photo_top - photo_h + 0.6 * mm,
                width=photo_w - 1.2 * mm,
                height=photo_h - 1.2 * mm,
                preserveAspectRatio=True,
                anchor="c",
                mask="auto",
            )
            photo_drawn = True
        except Exception:
            photo_drawn = False
    if not photo_drawn:
        # کادر خالی با متن
        c.setStrokeColor(GRAY_LINE)
        c.setLineWidth(0.8)
        c.setFillColor(HexColor("#f5f5f5"))
        c.rect(photo_x, photo_top - photo_h, photo_w, photo_h, fill=1, stroke=1)
        c.setFillColor(GRAY_TEXT)
        c.setFont("NotoAr", 8)
        c.drawCentredString(photo_x + photo_w / 2, photo_top - photo_h / 2 - 2, rtl("بدون عکس"))

    c.setFillColor(black)
    c.setFont("NotoAr", 9)
    # تاریخ
    date_label = rtl("تاریخ:")
    date_val = _to_persian_digits(date_str)  # ارقام فارسی
    c.drawRightString(page_w - margin_r, y, f"{date_val}  {date_label}")
    y -= 6 * mm
    # شماره
    no_label = rtl("شماره:")
    c.drawRightString(page_w - margin_r, y, f"{no_str}  {no_label}")
    # بعد از بلوک عکس ادامه دهیم
    y = min(y - 6 * mm, photo_top - photo_h - 8 * mm)

    # --- مخاطب ---
    c.setFont("NotoArBold", 10)
    c.setFillColor(black)
    c.drawRightString(page_w - margin_r, y, rtl(recipient_text))
    y -= 9 * mm

    # --- متن نامه ---
    c.setFont("NotoArBold", 10)
    c.drawRightString(page_w - margin_r, y, rtl("با سلام"))
    y -= 8 * mm

    c.setFillColor(black)
    body = (
        f"احتراماً خواهشمند است نسبت به انجام آزمایشات عدم اعتیاد از آقای {full_name} "
        f"نام پدر {father_disp} با کد ملی {nid_disp} "
        f"متقاضی اشتغال در این شرکت می‌باشد، لذا اقدامات لازم را مبذول نموده "
        f"و نتیجه را کتباً به این شرکت اعلام فرمایید. قبلاً از همکاری شما سپاسگزاریم."
    )
    # قالب‌بندی پیشرفته: justify + تورفتگی خط اول + برجسته کردن نام و کد ملی
    bold_parts = [p for p in (full_name, father_disp, nid_disp) if p and p != "—"]
    y = _draw_rtl_paragraph(
        c,
        body,
        x_left=margin_l,
        x_right=page_w - margin_r,
        y=y,
        font_name="NotoAr",
        font_bold="NotoArBold",
        font_size=10,
        leading=6.5 * mm,
        justify=True,
        first_line_indent=6 * mm,
        bold_phrases=bold_parts,
    )
    y -= 10 * mm

    # --- امضا ---
    sig_x = margin_l + 28 * mm
    c.setFont("NotoAr", 10)
    c.drawCentredString(sig_x, y, rtl("با تشکر"))
    y -= 6 * mm
    c.setFont("NotoArBold", 10)
    c.drawCentredString(sig_x, y, rtl("علی مبارکی"))
    y -= 5 * mm
    c.setFont("NotoArSemi", 9)
    c.setFillColor(BLUE)
    c.drawCentredString(sig_x, y, rtl("رئیس حراست IOEC خرمشهر"))

    # --- پاورقی اعتبار ---
    c.setFillColor(HexColor("#444444"))
    c.setFont("NotoAr", 7)
    # نمایش تعداد روز به فارسی
    _fa_digits = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
    days_fa = str(validity_days_n).translate(_fa_digits)
    footer = rtl(f"این برگه از تاریخ صدور فقط {days_fa} روز اعتبار دارد.")
    c.drawCentredString(page_w / 2, 10 * mm, footer)
    c.setStrokeColor(GRAY_LINE)
    c.setLineWidth(0.5)
    c.line(margin_l, 14 * mm, page_w - margin_r, 14 * mm)

    c.save()
    return buf.getvalue()
