"""
تقویم شمسی (جلالی) — تعیین‌جدولِ قطعی، بدون وابستگی به تنظیمات ویندوز/Excel.

چرا جدول و نه فرمولِ تبدیل؟
  • Excel تابعِ بومیِ JALALI ندارد و «قالبِ تاریخِ ویندوز» `[$-960429]` که فایل
    فعلی استفاده می‌کند، نمایش را به تنظیماتِ دستگاه وابسته می‌کند
    (یافتهٔ D6 فاز ۱؛ روی سیستمی با تقویمِ دیگر، گزارشِ مدیریتی جابه‌جا می‌شود).
  • الگوریتم را همین‌جا اجرا و **با jdatetime (تقویم رسمی ایران) روی همهٔ روزهای
    ۱۹۰۰–۲۲۰۰ سنجیده‌ایم** (tools/test_jalali.py، T-J1: ۱۰۹٬۹۳۸ روز، صفر اختلاف).
  • خروجی به‌شکل جدول در شیت `Helper` نوشته می‌شود؛ پس در فایل، تاریخ
    «داده» است نه محاسبه — هیچ فرمولِ شکننده‌ای برای بازکردن فایل لازم نیست.

قاعدهٔ کبیسه (چرخهٔ ۳۳ساله، مطابق تقویم رسمی):
    ۱ فروردینِ سال jy = 18/03/622م + 365*(jy-1) + floor((jy-1+16)*8/33) روز
توجه: روزِ کبیسه در **پایان** همان سال می‌افتد (اسفندِ ۳۰ روزه → ۲۹/۳۰ روزه)،
پس طولِ ماه از مرزهای ۱ فروردین محاسبه می‌شود، نه از قاعدهٔ ثابت.

ردیف‌های جدولِ Helper:
    G_DATE | J_TEXT | J_YEAR | J_MONTH | J_DAY | J_MONTH_NAME | J_MONTH_KEY |
    J_WEEK_KEY | WEEKDAY_FA | IS_HOLIDAY
"""
from __future__ import annotations

import datetime

_J0 = datetime.date(622, 3, 18)     # ۱ فروردین سال ۱ جلالی (میلادیِ پرولهپتیک)
LEAP_SHIFT = 16

# ترتیبِ روزهای هفته در ایران: شنبه=۰ … جمعه=۶.
# Python: Mon=0 … Fri=4, Sat=5, Sun=6  ⇒  اندیس = (weekday() + 2) % 7
#   Sat: 5+2=7 → 0=شنبه ✓  Sun: 6+2=8 → 1=یکشنبه ✓  Fri: 4+2=6 → 6=جمعه ✓
_WEEKDAY_FA = ('شنبه', 'یکشنبه', 'دوشنبه', 'سه‌شنبه', 'چهارشنبه', 'پنجشنبه', 'جمعه')
_MONTH_FA = ('فروردین', 'اردیبهشت', 'خرداد', 'تیر', 'مرداد', 'شهریور',
             'مهر', 'آبان', 'آذر', 'دی', 'بهمن', 'اسفند')


def _leaps_before(jy: int) -> int:
    """تعدادِ روزهایِ کبیسهٔ جمع‌شده پیش از ۱ فروردینِ سال jy."""
    return ((jy - 1 + LEAP_SHIFT) * 8) // 33


def _farvardin_1(jy: int) -> datetime.date:
    return _J0 + datetime.timedelta(days=365 * (jy - 1) + _leaps_before(jy))


def g_to_j(date: datetime.date) -> tuple[int, int, int]:
    """میلادی → جلالی (jy, jm, jd)."""
    for jy in (date.year - 622, date.year - 621, date.year - 620):
        f1 = _farvardin_1(jy)
        if f1 <= date < _farvardin_1(jy + 1):
            doy = (date - f1).days + 1
            if doy <= 186:
                return jy, (doy - 1) // 31 + 1, (doy - 1) % 31 + 1
            doy -= 186
            return jy, (doy - 1) // 30 + 7, (doy - 1) % 30 + 1
    raise ValueError(f'jalali lookup failed for {date}')


def j_days_in_year(jy: int) -> int:
    return (_farvardin_1(jy + 1) - _farvardin_1(jy)).days


def is_leap_jy(jy: int) -> bool:
    return j_days_in_year(jy) == 366


def j_days_in_month(jy: int, jm: int) -> int:
    if jm <= 6:
        return 31
    if jm <= 11:
        return 30
    return j_days_in_year(jy) - 186 - 150


def j_to_g(jy: int, jm: int, jd: int) -> datetime.date:
    """جلالی → میلادی؛ تاریخِ نامعتبر را با پیامِ فارسی رد می‌کند (نه خطای خام)."""
    dim = j_days_in_month(jy, jm)
    if not (1 <= jm <= 12):
        raise ValueError(f'ماه جلالی نامعتبر: {jy}/{jm}/{jd}')
    if not (1 <= jd <= dim):
        raise ValueError(f'روز {jd} در ماه {jm} سال {jy} معتبر نیست (بیشترین {dim} روز)')
    off = (31 * (jm - 1) + jd - 1) if jm <= 7 else (186 + 30 * (jm - 7) + jd - 1)
    d = _farvardin_1(jy) + datetime.timedelta(days=off)
    if g_to_j(d) != (jy, jm, jd):                      # گاردِ درونی: round-trip
        raise AssertionError(f'jalali round-trip broken for {jy}/{jm}/{jd}')
    return d


def parse_j_text(text) -> tuple[int, int, int] | None:
    """'1405/06/28' → (1405, 6, 28)؛ اگر رشته/قالب درست نبود None."""
    if not isinstance(text, str):
        return None
    parts = text.strip().replace('-', '/').split('/')
    if len(parts) != 3 or any(not p.isdigit() for p in parts):
        return None
    jy, jm, jd = (int(p) for p in parts)
    try:
        j_to_g(jy, jm, jd)
    except ValueError:
        return None
    return jy, jm, jd


def j_text(jy: int, jm: int, jd: int) -> str:
    return f'{jy:04d}/{jm:02d}/{jd:02d}'


def j_month_name(jm: int) -> str:
    return _MONTH_FA[jm - 1]


def weekday_fa(date: datetime.date) -> str:
    return _WEEKDAY_FA[(date.weekday() + 2) % 7]


def week_index(jy: int, jm: int, jd: int) -> int:
    """شمارهٔ هفتهٔ جلالی، شنبه‌مبدأ؛ هفتهٔ ۱ همان هفته‌ای است که ۱ فروردین در آن است."""
    d = j_to_g(jy, jm, jd)
    f1 = _farvardin_1(jy)
    lead = (f1.weekday() + 2) % 7            # تعدادِ روزِ «شنبه» تا ۱ فروردین
    return ((d - f1).days + lead) // 7 + 1


def week_key(jy: int, jm: int, jd: int) -> str:
    return f'{jy:04d}-W{week_index(jy, jm, jd):02d}'


# تعطیلات رسمی — تنها مناسبت‌های *ثابت* (جمعه + تاریخ‌های قطعی). تعطیلاتِ متحرک
# (عیدِ فطر/قربان و…) در این جدول نیست و در Help صریحاً ذکر شده؛ برای نسخهٔ
# بعدی می‌توان ستونِ IS_HOLIDAY را از فایلِ رسمی جایگزین کرد.
_FIXED_JALALI_HOLIDAYS = {(1, 1), (1, 2), (1, 3), (1, 4), (1, 12), (1, 13),
                          (3, 14), (3, 15), (6, 14), (10, 24), (11, 1), (12, 29)}
_FIXED_GREG_HOLIDAYS = {(y, 2, 11) for y in range(2018, 2036)}


def is_holiday(date: datetime.date) -> bool:
    if date.weekday() == 4:                      # Friday == 4 در Python
        return True
    if (date.year, date.month, date.day) in _FIXED_GREG_HOLIDAYS:
        return True
    _, jm, jd = g_to_j(date)
    return (jm, jd) in _FIXED_JALALI_HOLIDAYS


CAL_START = datetime.date(2018, 1, 1)
CAL_END = datetime.date(2035, 12, 31)

CAL_COLUMNS = ('G_DATE', 'J_TEXT', 'J_YEAR', 'J_MONTH', 'J_DAY', 'J_MONTH_NAME',
               'J_MONTH_KEY', 'J_WEEK_KEY', 'WEEKDAY_FA', 'IS_HOLIDAY')


def rows(start: datetime.date = CAL_START, end: datetime.date = CAL_END):
    d = start
    while d <= end:
        jy, jm, jd = g_to_j(d)
        yield {
            'G_DATE': d,
            'J_TEXT': j_text(jy, jm, jd),
            'J_YEAR': jy,
            'J_MONTH': jm,
            'J_DAY': jd,
            'J_MONTH_NAME': j_month_name(jm),
            'J_MONTH_KEY': jy * 100 + jm,
            'J_WEEK_KEY': week_key(jy, jm, jd),
            'WEEKDAY_FA': weekday_fa(d),
            'IS_HOLIDAY': 'بله' if is_holiday(d) else 'خیر',
        }
        d += datetime.timedelta(days=1)
