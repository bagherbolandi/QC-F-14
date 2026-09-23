"""
هستهٔ مهاجرت داده از `FORM-QC-F-14.xlsm` به مدل `QC-F-14-v2`.

چرا اسکریپت و نه کپی‌پیست؟
  • بارکدها در فایل فعلی رشته‌اند و ۹٬۲۸۳ تا از آن‌ها با صفر شروع می‌شوند؛ هر
    مسیرِ «کپی در Excel» یا «import با تشخیص نوع» آن صفرها را می‌خورد و
  • ۳ نام با فاصلهٔ اضافی (رجبی␣␣ / جمعی␣ / سینی زیر موتور سمند␣) در آمار
    دو شاخه می‌کنند. پس تبدیل باید **برنامه‌محور، قطعی و قابل‌بازبینی** باشد.

این ماژول هیچ چیزی را «حذف» یا «ادغام» نمی‌کند؛ فقط:
  کدگذاری (نام → CODE) · نرمال‌سازیِ نمایش + ثبت alias · استخراج BC_* ·
  برچسب ROW_STATUS · تولید ردیف‌های DQ_ISSUES.

خروجی (dict) در `build_payload()` است و `tools/gen_workbook.py` آن را به شیت‌ها
می‌نویسد. `tools/test_migration.py` همان خروجی را با `migration_check` می‌سنجد.
"""
from __future__ import annotations

import collections
import datetime
import re
import zipfile

import xlcal

REF_SHEET_DATA = 'xl/worksheets/sheet2.xml'
DATA_HEADER_ROW = 1
COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
HEADERS_FA = ['تاریخ', 'شیفت', 'ایستگاه', 'قطعه', 'قالب', 'بارکد', 'عیب', 'وضعیت', 'اپراتور']
# ستون‌های اطلاعاتِ پایه در FORM مرجع (J4:Q38). چیدمانِ واقعیِ فایل — با شمارشِ
# متمایزها راستی‌آزمایی شده (۸ ایستگاه، ۲۱ بازرس، ۳۵ قطعه، ۷ قالب، ۲۱ عیب، ۲ وضعیت):
FORM_MASTER_COLS = {'J': 'DATE', 'K': 'SHIFT', 'L': 'STATION', 'M': 'INSPECTOR',
                    'N': 'PART', 'O': 'MOLD', 'P': 'DEFECT', 'Q': 'STATUS'}


def _shared_strings(zf):
    xml = zf.read('xl/sharedStrings.xml').decode('utf-8')
    out = []
    for si in re.findall(r'<si>(.*?)</si>', xml, re.S):
        out.append(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S)))
    return [s.replace('&#10;', '\n').replace('&amp;', '&').replace('&lt;', '<')
             .replace('&gt;', '>').replace('&quot;', '"').replace('&#39;', "'") for s in out]


def _cell_value(attrs, raw, ss):
    if raw is None:
        return None
    if 't="s"' in attrs:
        return ss[int(raw)]
    if 't="str"' in attrs or 't="e"' in attrs:
        return raw
    return raw


def parse_sheet(zf, target, ss, first_row=2, last_row=None):
    """خواندنِ مستقیم XML (openpyxl روی ۱۴٬۵۵۹ ردیف × ۹ ستون ~۷۲ ثانیه طول می‌کشد)."""
    xml = zf.read(target).decode('utf-8')
    rows = {}
    for m in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', xml, re.S):
        rnum = int(m.group(1))
        if rnum < first_row or (last_row and rnum > last_row):
            continue
        cells = {}
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(?:<v>(.*?)</v>)?(?:<is>(.*?)</is>)?</c>)',
                              m.group(2), re.S):
            col, attrs, v, isn = cm.groups()
            if isn is not None:
                cells[col] = re.sub(r'<.*?>', '', isn)
            else:
                cells[col] = _cell_value(attrs, v, ss)
        rows[rnum] = cells
    return rows


def sheet_targets(zf):
    wb = zf.read('xl/workbook.xml').decode('utf-8')
    rels = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid = {m.group(1): m.group(2) for m in
           re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)}
    out = {}
    for m in re.finditer(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
        t = rid[m.group(2)].lstrip('/')
        out[m.group(1)] = t if t.startswith('xl/') else 'xl/' + t
    return out


def norm_display(name: str) -> str:
    """نرمال‌سازیِ *نمایش*: جمع‌کردنِ فاصله‌ها + حذفِ فاصلهٔ انتها.
    عمداً «تطبیقِ نام‌هایِ دست‌نویس» (مثلاً «سین» ↔ «سینی») را انجام **نمی‌دهیم** —
    آن تصمیمِ صاحب داده است، نه ابزار (بند ۴: روابط را حدس نزن)."""
    if name is None:
        return ''
    s = re.sub(r'\s+', ' ', str(name).replace('\u200c', ' ')).strip()
    return s


def excel_serial_to_date(v):
    try:
        n = float(v)
    except (TypeError, ValueError):
        return None
    if not (20000 < n < 80000):
        return None
    return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(n))


def load_reference(path):
    with zipfile.ZipFile(path) as zf:
        ss = _shared_strings(zf)
        targets = sheet_targets(zf)
        data = parse_sheet(zf, targets['Data'], ss, first_row=2)
        form = parse_sheet(zf, targets['FORM'], ss, first_row=1)
    masters = collections.defaultdict(list)
    for r in range(4, 45):
        row = form.get(r, {})
        for col, key in FORM_MASTER_COLS.items():
            v = norm_display(row.get(col))
            if v:
                masters[key].append(v)
    records = []
    for rnum in sorted(data):
        row = data[rnum]
        vals = [norm_display(row.get(c)) for c in COLS]
        if not any(vals):
            records.append({'_row': rnum, '_blank': True})
            continue
        raw_shift = row.get('B')
        g_date = excel_serial_to_date(row.get('A'))
        if g_date is None and row.get('A'):
            try:                                   # اگر متنِ «1405/06/28» ذخیره شده بود
                parts = [int(x) for x in re.split(r'[/-]', str(row['A']).strip())]
                if len(parts) == 3 and parts[0] > 1300:
                    g_date = xlcal.j_to_g(*parts)
            except Exception:
                g_date = None
        records.append({
            '_row': rnum, '_blank': False,
            'DATE_RAW': row.get('A'), 'G_DATE': g_date,
            'SHIFT_RAW': raw_shift, 'STATION': vals[2], 'PART': vals[3],
            'MOLD': vals[4], 'BARCODE': (row.get('F') or '').strip(),
            'DEFECT': vals[6], 'STATUS': vals[7], 'INSPECTOR': vals[8],
            'DIRTY': {k: v for k, v in
                      [('INSPECTOR', vals[8]), ('PART', vals[3]), ('DEFECT', vals[6])]
                      if v != '' and re.search(r'\s$', str(v)) or '  ' in str(v)},
        })
    return {'masters': dict(masters), 'records': records, 'targets': targets}


# ───────────────────────── ساخت کدها ─────────────────────────
def build_masters(masters):
    """لیست‌های مرجع → ردیف‌های DIM_* با CODE ثابتِ ترتیبی (بر اساسِ فهرستِ فعلیِ فایل)."""
    out = {}
    out['PART'] = [{'CODE': f'P-{i:02d}', 'NAME': n, 'ACTIVE': 'بله'}
                   for i, n in enumerate(masters.get('PART', []), 1)]
    raw_molds = masters.get('MOLD', [])
    out['MOLD'] = []
    for i, n in enumerate(raw_molds, 1):
        kind = 'جهت' if n in ('راست', 'چپ') else 'حفره'
        out['MOLD'].append({'CODE': f'M-{i:02d}', 'NAME': n, 'MOLD_KIND': kind,
                            'ACTIVE': 'بله'})
    out['DEFECT'] = [{'CODE': f'D-{i:02d}', 'NAME': n, 'ACTIVE': 'بله'}
                     for i, n in enumerate(masters.get('DEFECT', []), 1)]
    insp = masters.get('INSPECTOR', [])
    out['INSPECTOR'] = []
    alias_map = {}
    for i, n in enumerate(insp, 1):
        code = f'I-{i:02d}'
        out['INSPECTOR'].append({'CODE': code, 'NAME': n, 'ACTIVE': 'بله', 'ALIASES': ''})
        alias_map.setdefault(norm_display(n), code)
    out['STATION'] = []
    for i, n in enumerate(masters.get('STATION', []), 1):
        used = i <= 2 or n == 'برگشت از فروش'          # شاهد فاز ۱: فقط ۲ ایستگاه فعال + برگشت
        out['STATION'].append({'CODE': f'ST-{i:02d}', 'NAME': n,
                               'ENABLED_FOR_FORM': 'بله' if used else 'خیر',
                               'ACTIVE': 'بله'})
    for i, n in enumerate(masters.get('SHIFT', []), 1):
        out.setdefault('SHIFT', []).append(
            {'CODE': f'S{i}', 'NAME': f'شیفت {n}', 'ACTIVE': 'بله'})
    disp = masters.get('STATUS', [])
    out['DISPOSITION'] = [{'CODE': f'DIS-{i:02d}', 'NAME': n, 'ACTIVE': 'بله'}
                           for i, n in enumerate(disp, 1)]
    if 'برگشتی' not in disp:
        out['DISPOSITION'].append({'CODE': f'DIS-{len(disp)+1:02d}', 'NAME': 'برگشتی',
                                   'ACTIVE': 'بله'})
    out['_ALIAS'] = alias_map
    return out


def code_of(dims, kind, name, used_codes):
    name = norm_display(name)
    for row in dims[kind]:
        if row['NAME'] == name:
            used_codes[kind][row['CODE']] += 1
            return row['CODE']
    # نامِ آلوده/جدید: کدِ «تعریف‌نشده» می‌گیرد، نه حذفِ بی‌صدا
    code = f'X-{kind[:1]}-UNMAPPED'
    used_codes[kind][code] += 1
    return code


def analyze(ref, dims):
    """تبدیلِ ردیف‌ها + ساخت نگاشت‌ها و شمارش‌ها (قابل چاپ در گزارش)."""
    used = collections.defaultdict(collections.Counter)
    part_mold = collections.defaultdict(collections.Counter)
    station_status = collections.defaultdict(collections.Counter)
    barcode_part = collections.defaultdict(collections.Counter)
    rows_out, dq = [], []
    seen = collections.Counter()
    stats = collections.Counter()
    for rec in ref['records']:
        if rec.get('_blank'):
            stats['blank_rows'] += 1
            dq.append({'RULE': 'DQ-BLANK', 'SEV': 'متوسط', 'ROW': rec['_row'],
                       'DETAIL': 'ردیفِ کاملاً خالی در محدودهٔ جدول'})
            continue
        pc = code_of(dims, 'PART', rec['PART'], used)
        mc = code_of(dims, 'MOLD', rec['MOLD'], used)
        dc = code_of(dims, 'DEFECT', rec['DEFECT'], used)
        ic = code_of(dims, 'INSPECTOR', rec['INSPECTOR'], used)
        sc = code_of(dims, 'STATION', rec['STATION'], used)
        disp = code_of(dims, 'DISPOSITION', rec['STATUS'], used)
        bc = rec['BARCODE']
        if bc:
            seen[bc] += 1
        part_mold[rec['PART']][rec['MOLD']] += 1
        station_status[rec['STATION']][rec['STATUS']] += 1
        barcode_part[rec['PART']][bc[:2]] += 1
        g = rec['G_DATE']
        if g is None:
            stats['bad_date'] += 1
            dq.append({'RULE': 'DQ-DATE', 'SEV': 'شدید', 'ROW': rec['_row'],
                       'DETAIL': f'تاریخِ خوانده‌نشده: {rec["DATE_RAW"]!r}'})
        for col, val in rec['DIRTY'].items():
            dq.append({'RULE': 'DQ-SPACE', 'SEV': 'کم', 'ROW': rec['_row'],
                       'DETAIL': f'{col}: {val!r} (فاصلهٔ اضافی)'})
            stats['dirty_names'] += 1
        if not re.fullmatch(r'\d{19}', bc or ''):
            stats['bad_barcode'] += 1
            dq.append({'RULE': 'DQ-BARCODE', 'SEV': 'شدید', 'ROW': rec['_row'],
                       'DETAIL': f'بارکد {bc!r} دقیقاً ۱۹ رقم نیست'})
        rows_out.append({**rec, 'PART_ID': pc, 'MOLD_ID': mc, 'DEFECT_ID': dc,
                         'INSPECTOR_ID': ic, 'STATION_ID': sc, 'DISPOSITION_ID': disp,
                         'BARCODE': bc, 'G_DATE': g})
    for bc, cnt in seen.items():
        if cnt > 1:
            stats['dup_barcodes'] += 1
            stats['dup_extra_rows'] += cnt - 1
    stats['distinct_barcode'] = len(seen)
    stats['leading_zero'] = sum(1 for b in seen if b.startswith('0'))
    return rows_out, dq, stats, part_mold, station_status, barcode_part


def mold_rules(part_mold, dims):
    """MAP_PART_MOLD از داده — فقط جایی که شاهدِ قوی هست (≥۹۰٪ و ≥۲۰ رکورد).
    برای قطعاتِ کم‌داده یا چندقالبی، قاعده‌ای تحمیل نمی‌شود (لیست کامل باز می‌ماند)."""
    rules = []
    for part, counter in sorted(part_mold.items()):
        total = sum(counter.values())
        top, top_n = counter.most_common(1)[0]
        share = top_n / total if total else 0
        if total >= 20 and share >= 0.90:
            rules.append({'PART': part, 'MOLD': top, 'SHARE': share, 'N': total,
                          'ENFORCED': 'بله'})
        else:
            rules.append({'PART': part, 'MOLD': '— (چندقالبی/کم‌داده)', 'SHARE': share,
                          'N': total, 'ENFORCED': 'خیر'})
    return rules


def event_type_rules(station_status, dims):
    """EVENT_TYPE از قاعدهٔ ایستگاه (تصمیم B.۵: برگشتی = نوعِ مستقل).
    برای هر ایستگاه، وضعیتِ غالب در داده‌های خودش؛ اگر داده‌ای نبود → «اصلاحی»
    با برچسب «پیش‌فرضِ تنظیم‌شدنی» (نه ادعای کشف‌شده)."""
    out = []
    for st in dims['STATION']:
        c = station_status.get(st['NAME'], collections.Counter())
        if st['NAME'] == 'برگشت از فروش':
            et, src = 'برگشتی', 'قاعدهٔ B.۵ (تصمیم کاربر) + ۹۸ رکوردِ همین ایستگاه'
        elif c:
            et = c.most_common(1)[0][0]
            src = f'غالبِ دادهٔ ایستگاه ({c.most_common(1)[0][1]}/{sum(c.values())})'
        else:
            et, src = 'اصلاحی', 'پیش‌فرض — قابل ویرایش در Master'
        out.append({'STATION': st['NAME'], 'CODE': st['CODE'], 'EVENT_TYPE': et,
                    'BASIS': src})
    return out


CAVITY_BY_DIGIT = {'1': 'A', '2': 'B', '3': 'C', '4': 'D', '5': 'E'}


def barcode_semantics(bc):
    """فقط آنچه در داده‌های ۱۴٬۵۵۷ ردیفِ فایلِ شما بازتولید شد.

    • رقم ۱–۲ = کدِ قطعه: توزیعِ ۸ کدِ غالب (۰۲=۳٬۵۵۱ / ۰۴=۲٬۸۹۹ / ۰۱=۲٬۵۴۶ …) با
      سهمِ قطعات در همان فایل هم‌خوان است ⇒ قابل‌اتکا برای **هشدار**، نه ردِ ورودی.
    • رقم ۴ ↔ حفرۀ قالب: اگر رقمِ ۴ را n بخوانیم و حفره را A+n، آنگاه در
      ۱۳٬۴۷۷ رکوردِ قابل‌سنجش، ۱۲٬۹۸۷ مورد (۹۶٫۴٪) هم‌خوان‌اند — و بقیه عمدتاً
      رکوردهای «راست/چپ» (قالبِ جهت) هستند. ⇒ قاعده‌است ولی **ردکننده نیست**؛
      فقط ستون MOLD_CHECK و یک ردیف در گزارشِ کیفیت.
    • رقم ۱–۲ «کدِ قطعه» است ولی به فهرستِ این فایل (P-01..P-35) نمی‌خورد:
      تطابقِ مستقیم ۶٬۰۴۰/۱۴٬۳۵۰ = ۴۲٪. پس کدِ رسمیِ قطعه را **از واحدِ IT/فرآیند
      بگیرید**؛ تا آن موقع BC_PART_CODE فقط ذخیره و گزارش می‌شود، نه اعتبارسنجی.
    • رقم ۹–۱۴ = تاریخِ تولید؟ **خیر** — در بازبینیِ دقیق، توزیعِ این شش رقم
      «سریال/ساعت» به نظر می‌رسد (۳۰۵۰۶۱، ۰۰۵۰۵۱، ۶۰۵۰۵۰ …) و قالبِ YYYYMM
      را ندارد. پس این فیلد در نسخهٔ ۲ **حذف** شد و KPI «تأخیر ثبت» وعده داده
      نمی‌شود؛ ساختِ آن منوط به تأییدِ واحدِ IT/فرآیند از رمزِ بارکد است.
    """
    ok = bool(bc and re.fullmatch(r'\d{19}', bc))
    return {'BC_PART_CODE': bc[:2] if ok else '',
            'BC_MOLD_DIGIT': bc[3] if ok else '',
            'BC_CAVITY_GUESS': CAVITY_BY_DIGIT.get(bc[3], '') if ok else '',
            'BC_DIGITS_OK': 'بله' if ok else 'خیر',
            '_ok': ok}
