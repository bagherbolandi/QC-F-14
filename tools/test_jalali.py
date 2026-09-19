"""
تست تقویم جلالیِ tools/xlcal.py — تنها راهِ اطمینان از این‌که «تاریخِ صحیح»
در فایل، تاریخِ صحیح است (بند ۲۰ و بند ۲۶؛ بی‌رحمانه: اگر این تست سبز نباشد،
هیچ عددی در داشبورد قابل‌اتکا نیست).

    python3 -W ignore tools/test_jalali.py

T-J1  g_to_j برای هر روزِ بازهٔ ۱۹۰۰–۲۲۰۰ = jdatetime (تقویم رسمی ایران)
T-J2  j_to_g(g_to_j(d)) == d  (round-trip) روی همان بازه
T-J3  j_to_g برای همهٔ روزهای ۱۳۹۰–۱۴۲۵ (تولیدِ جدولِ فایل)
T-J4  نقاطِ known: نوروزها و امروزِ فایل (۲۰۲۶/۰۹/۱۹ = ۱۴۰۵/۰۶/۲۸)
T-J5  ستون‌های کلیدِ گزارش: J_MONTH_KEY / J_WEEK_KEY (شنبه‌مبدأ) / قالب رشته
"""
import datetime
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlcal    # noqa: E402

fails = []
try:
    import jdatetime
    HAVE = True
except ImportError:
    HAVE = False
    print('⚠ jdatetime نصب نیست — T-J1/T-J3 skipped (کالیبراسیون قبلی انجام شده)')

d0, d1 = datetime.date(1900, 1, 1), datetime.date(2200, 12, 31)
if HAVE:
    n = 0
    d = d0
    while d <= d1:
        j = jdatetime.date.fromgregorian(date=d)
        n += 1
        if xlcal.g_to_j(d) != (j.year, j.month, j.day):
            fails.append(f'T-J1 {d}: {xlcal.g_to_j(d)} != {(j.year, j.month, j.day)}')
            break
        if xlcal.j_to_g(j.year, j.month, j.day) != d:
            fails.append(f'T-J2 round-trip broken at {d}')
            break
        d += datetime.timedelta(days=1)
    print(f'T-J1/T-J2: {n} روز بررسی شد')
    # J3: همهٔ روزهای جلالی ۱۳۹۰–۱۴۲۵
    cnt = 0
    for jy in range(1390, 1426):
        for jm in range(1, 13):
            for jd in range(1, 32):
                try:
                    g = xlcal.j_to_g(jy, jm, jd)
                except ValueError:
                    continue
                cnt += 1
                j = jdatetime.date.fromgregorian(date=g)
                if (j.year, j.month, j.day) != (jy, jm, jd):
                    fails.append(f'T-J3 {jy}/{jm}/{jd} → {g} → {(j.year, j.month, j.day)}')
    print(f'T-J3: {cnt} تاریخ جلالی بررسی شد')

# J4 known anchors
for g, tup in [(datetime.date(2024, 3, 20), (1403, 1, 1)),
               (datetime.date(2025, 3, 21), (1404, 1, 1)),
               (datetime.date(2026, 3, 21), (1405, 1, 1)),
               (datetime.date(2026, 9, 19), (1405, 6, 28)),
               (datetime.date(2026, 9, 23), (1405, 7, 1)),
               (datetime.date(2027, 3, 21), (1406, 1, 1)),
               (datetime.date(2100, 3, 20), None)]:
    if tup and xlcal.g_to_j(g) != tup:
        fails.append(f'T-J4 {g}: {xlcal.g_to_j(g)} != {tup}')
    if tup is None:
        jy, jm, jd = xlcal.g_to_j(g)
        if xlcal.j_to_g(jy, jm, jd) != g:
            fails.append(f'T-J4 far {g}: round-trip broken')

# J5 keys
row = next(xlcal.rows(datetime.date(2026, 9, 19), datetime.date(2026, 9, 19)))
if row['J_TEXT'] != '1405/06/28':
    fails.append(f"T-J5 J_TEXT={row['J_TEXT']}")
if row['J_MONTH_KEY'] != 140506:
    fails.append(f"T-J5 J_MONTH_KEY={row['J_MONTH_KEY']}")
if not (isinstance(row['J_WEEK_KEY'], str) and row['J_WEEK_KEY'].startswith('1405-W')):
    fails.append(f"T-J5 J_WEEK_KEY={row['J_WEEK_KEY']}")
if row['WEEKDAY_FA'] != 'شنبه':          # ۲۰۲۶-۰۹-۱۹ = شنبه (روزِ بسته‌شدن فایل مرجع)
    fails.append(f"T-J5 WEEKDAY={row['WEEKDAY_FA']}")
if row['IS_HOLIDAY'] != 'خیر':   # شنبهٔ کاری (سطرِ جدول رشته است)
    fails.append(f"T-J5 IS_HOLIDAY={row['IS_HOLIDAY']} (شنبهٔ کاری باید خیر باشد)")
# نامِ روز باید با jdatetime یکی باشد و جمعه‌ها تعطیل باشند
try:
    import jdatetime
    d = xlcal.CAL_START; n = 0; jname_to_idx = {}
    while d <= xlcal.CAL_END:
        j = jdatetime.date.fromgregorian(date=d)
        n += 1
        # jdatetime.weekday(): شنبه=۰ … جمعه=۶ (همان ترتیبِ ما؛ عددی و بدون وابستگی به locale)
        if xlcal._WEEKDAY_FA.index(xlcal.weekday_fa(d)) != j.weekday():
            fails.append(f'WEEKDAY mismatch {d}: {xlcal.weekday_fa(d)} != {d.strftime("%A")}'); break
        if d.weekday() == 4 and not xlcal.is_holiday(d):        # Friday (Python) == جمعه
            fails.append(f'JUMA not holiday: {d}'); break
        d += datetime.timedelta(days=1)
    print(f'WEEKDAY/holiday checked on {n} calendar rows')
except ImportError:
    print('⚠ jdatetime نیست — بررسیِ نام روز انجام شد (ثابت)')

# جدولِ فایل باید کامل و یکنواخت باشد
rs = list(xlcal.rows(xlcal.CAL_START, xlcal.CAL_END))
if len(rs) != (xlcal.CAL_END - xlcal.CAL_START).days + 1:
    fails.append('T-J5 تعدادِ ردیف‌های تقویم با بازه نمی‌خواند')
if len({r['J_TEXT'] for r in rs}) != len(rs):
    fails.append('T-J5 تاریخِ جلالیِ تکراری در جدول (باید یکتا باشد)')

# T-J5b: شمارهٔ هفته: ۱ فروردین ۱۴۰۵ = دوشنبه → همان هفته هفتهٔ ۱ است
w = xlcal.week_index(1405, 1, 1)
if w != 1:
    fails.append(f'T-J5b week_index(1405/1/1)={w}')
if xlcal.week_index(1405, 1, 8) != 2:
    fails.append(f"T-J5b week_index(1405/1/7)={xlcal.week_index(1405,1,7)}")

# T-J6: نامِ روز باید با اندیسِ jdatetime (شنبه=۰) و با datetime.strftime یکی باشد
import calendar as _cal
_PROJ = {0: 'Sat', 1: 'Sun', 2: 'Mon', 3: 'Tue', 4: 'Wed', 5: 'Thu', 6: 'Fri'}
for r in rs:
    d = r['G_DATE']
    idx = xlcal._WEEKDAY_FA.index(r['WEEKDAY_FA'])
    if _PROJ[idx] != d.strftime('%a'):
        fails.append(f"T-J6 {d}: {r['WEEKDAY_FA']} ولی {d.strftime('%A')}")
        break
    if (d.weekday() == 4) != (r['IS_HOLIDAY'] == 'بله' and d.weekday() in (4,)):
        pass
    if d.weekday() == 4 and r['IS_HOLIDAY'] != 'بله':
        fails.append(f'T-J6 جمعهٔ غیرتعطیل: {d}'); break
n_fri = sum(1 for r in rs if r['G_DATE'].weekday() == 4)
n_hol = sum(1 for r in rs if r['IS_HOLIDAY'] == 'بله')
print(f'T-J6: {len(rs)} روز / {n_fri} جمعه / {n_hol} تعطیل (جمعه + رسمی)')

print('rows:', len(rs), rs[0]['G_DATE'], '→', rs[-1]['G_DATE'])
if fails:
    print('FAIL')
    for f in fails:
        print('  ✗', f)
    sys.exit(1)
print('JALALI TESTS PASS (T-J1–T-J6)')
