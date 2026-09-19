"""
ساختِ «بارِ نهایی» فایل نسخهٔ ۲ — مهاجرت + اطلاعاتِ پایه + قواعد + گزارش‌ها.

این فایل تنها «مغز» است: هیچ Excel/فرمولی نمی‌سازد؛ فقط دادهٔ نرمال‌شده و
متادیتایِ گزارش‌ها را به‌صورت JSON می‌نویسد تا `gen_workbook.py` بخواند و
`test_workbook.py` بسنجد. جدا‌کردنِ این دو، عمداً انجام شده: محاسباتِ مهاجرت
(۷۰ ثانیه روی ۱۴٬۵۵۹ ردیف) نباید با هر تغییرِ ظاهریِ فایل تکرار شود.

    python3 -W ignore tools/build_payload.py [FORM-QC-F-14.xlsm] [-o /tmp/payload.json] [--force]

خروجیِ مهم:
  facts          ردیف‌های tblQC (۲۰ ستون، مقادیرِ خالص — بدون فرمول روی داده)
  dims           DIM_* با CODE ثابت (کد، نام، فعال/غیرفعال، alias)
  code_name      کد → نام (برای ستون‌های snapshot و Drill-down)
  map_part_mold  قاعدهٔ قطعه→قالب، فقط با شاهدِ ≥۹۰٪ و ≥۲۰ رکورد
  rule_station   ایستگاه → نوعِ رخداد (با ستون BASIS: از داده / پیش‌فرض)
  dq_issues      ایرادهای داده (هرکدام با RULE + SEVERITY + ردیفِ مبدأ)
  reports        مشخصهٔ ۱۵ گزارش (بعدِ تحلیل، فیلترِ گروه‌بندی، نوع نمودار)
  stats          شمارش‌های تراز (برای migration_check و بخش «چرا این‌طور؟»)
"""
from __future__ import annotations

import argparse
import collections
import datetime
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlcal           # noqa: E402
import xlmigrate as M   # noqa: E402

FACT_COLUMNS = [
    ('EVENT_ID', 'E-000001'), ('EVENT_DATE_J', '1405/06/28'), ('EVENT_DATE_G', 'تاریخ میلادی'),
    ('J_MONTH_KEY', '140506'), ('J_WEEK_KEY', '1405-W26'), ('SHIFT_ID', 'S1'),
    ('STATION_ID', 'ST-07'), ('STATION_SNAP', 'نام ایستگاه در لحظه ثبت'),
    ('PART_ID', 'P-01'), ('PART_NAME_SNAP', 'نام قطعه در لحظه ثبت'),
    ('MOLD_ID', 'M-01'), ('MOLD_NAME_SNAP', 'نام قالب/حفره'),
    ('DEFECT_ID', 'D-01'), ('DEFECT_NAME_SNAP', 'نام عیب در لحظه ثبت'),
    ('DISPOSITION_ID', 'DIS-01'), ('EVENT_TYPE', 'برگشتی/ضایعات/اصلاحی'),
    ('BARCODE', 'رشته ۱۹ رقمی'), ('BC_PART_CODE', 'دو رقم اول بارکد'),
    ('MOLD_CHECK', 'هم‌خوان/ناهم‌خوان/نامعلوم'), ('ENTRY_USER', 'کاربر ثبت'),
    ('INSPECTOR_ID', 'I-01'), ('INSPECTOR_SNAP', 'نام بازرس در لحظه ثبت'),
]
MAX_FACT_ROWS = 20000          # جای خالیِ از پیش‌ساخته‌شده تا Table اتساع پیدا کند
# مقدارهایِ مجازِ MOLD_CHECK — در Helper به‌صورتِ گریدِ CODE_MOLD_CHECK نوشته می‌شود تا
# ماکرو آن‌ها را از فایل بخواند (نه از literal). T-W12 تطابقِ این دو را می‌سنجد.
MOLD_CHECK_OPTIONS = ['هم‌خوان', 'ناهم‌خوان', 'قالبِ جهت', 'رقمِ خارج از بازه', 'بارکد نامعتبر']
ROW_STATUS_OPTIONS = ['تأییدشده', 'پیش‌نویس', 'ابطال‌شده', 'مشکوک-تکراری']
REPORTS = [
    dict(id='R01', title='روند روزانه (۳۰ روز اخیر)', group='J_DAY', axis='EVENT_DATE_J',
         chart='line', block='روند', note='بر اساس تاریخ وقوع، نه تاریخ ثبت'),
    dict(id='R02', title='روند ماهانه — تفکیک نوع رخداد', group='J_MONTH_KEY', axis='J_MONTH_KEY',
         chart='line', block='روند'),
    dict(id='R03', title='روند هفتگی', group='J_WEEK_KEY', axis='J_WEEK_KEY', chart='line', block='روند'),
    dict(id='R04', title='شیفت (۱/۲/۳) × نوع رخداد', group='SHIFT_ID', axis='SHIFT_ID',
         chart='stacked', block='تولید'),
    dict(id='R05', title='ایستگاه × نوع رخداد', group='STATION_ID', axis='STATION_ID',
         chart='stacked', block='تولید'),
    dict(id='R06', title='۱۵ قطعه برتر (تعداد و سهم)', group='PART_ID', axis='PART_ID',
         chart='bar', block='قطعه', top=15),
    dict(id='R07', title='Pareto عیب (تعداد، ٪، تجمعی)', group='DEFECT_ID', axis='DEFECT_ID',
         chart='pareto', block='عیب', top=15),
    dict(id='R08', title='قالب/حفره × نوع رخداد', group='MOLD_ID', axis='MOLD_ID',
         chart='stacked', block='قطعه'),
    dict(id='R09', title='بازرس × تعداد (کنترل یکنواختی ثبت)', group='INSPECTOR_ID',
         axis='INSPECTOR_ID', chart='bar', block='بازرس', top=21,
         note='اپراتورِ فایلِ قبلی؛ نام‌ها TRIM و alias نرمال شده‌اند (ادغام نشدند)'),
    dict(id='R10', title='وضعیت/محل صدور (DISPOSITION)', group='DISPOSITION_ID', axis='DISPOSITION_ID',
         chart='bar', block='تولید'),
    dict(id='R11', title='ردیابی بارکد (Drill-down)', group='BARCODE', axis='BARCODE',
         chart=None, block='بارکد', special='barcode'),
    dict(id='R12', title='عیب بر حسب قطعه (ماتریس ۱۰×۱۰)', group='PART_X_DEFECT', axis=None,
         chart='matrix', block='عیب', special='matrix'),
    dict(id='R13', title='رکوردهای اصلاح‌شده/عقب‌گرد ثبت‌شده', group='ENTRY_VS_EVENT', axis=None,
         chart=None, block='کیفیت', special='backdate'),
    dict(id='R14', title='کیفیت داده — خلاصهٔ ایرادها', group='DQ_RULE', axis=None,
         chart='bar', block='کیفیت', special='dq'),
    dict(id='R15', title='خلاصه مدیریتی ماهانه (ضایعات/اصلاحی/برگشتی)', group='J_MONTH_KEY',
         axis='J_MONTH_KEY', chart='stacked', block='مدیریتی', special='mgmt'),
]
KPI = [
    dict(id='K1', name='کل موارد ثبت‌شده', unit='عدد', agg='count', etype=None),
    dict(id='K2', name='کل ضایعات', unit='عدد', agg='count', etype='ضایعات'),
    dict(id='K3', name='کل اصلاحی', unit='عدد', agg='count', etype='اصلاحی'),
    dict(id='K4', name='کل برگشتی', unit='عدد', agg='count', etype='برگشتی'),
    dict(id='K5', name='قطعات دارای عیب', unit='عدد', agg='distinct_part', etype=None),
    dict(id='K6', name='انواع عیب ثبت‌شده', unit='عدد', agg='distinct_defect', etype=None),
]
FILTERS = ['J_YEAR', 'J_MONTH_KEY', 'EVENT_DATE_J', 'J_WEEK_KEY', 'SHIFT_ID', 'STATION_ID',
           'PART_ID', 'MOLD_ID', 'DEFECT_ID', 'ROW_STATUS']
FILTER_LABELS = {'J_YEAR': 'سال جلالی', 'J_MONTH_KEY': 'ماه', 'EVENT_DATE_J': 'روز (تاریخ وقوع)',
                 'J_WEEK_KEY': 'هفته', 'SHIFT_ID': 'شیفت', 'STATION_ID': 'ایستگاه',
                 'PART_ID': 'قطعه', 'MOLD_ID': 'قالب/حفره', 'DEFECT_ID': 'عیب',
                 'ROW_STATUS': 'وضعیت رکورد'}


def name_of(dims, kind, code):
    for row in dims[kind]:
        if row['CODE'] == code:
            return row['NAME']
    return '— تعریف‌نشده —'


def build(ref_path, cache=None, verbose=True):
    if cache:
        os.makedirs(os.path.dirname(cache), exist_ok=True)
    if cache and os.path.exists(cache):
        ref = pickle_load(cache)
    else:
        ref = M.load_reference(ref_path)
        if cache:
            pickle_dump(cache, ref)
    dims = M.build_masters(ref['masters'])
    rows, dq, stats_raw, part_mold, station_status, _ = M.analyze(ref, dims)
    stats = dict(stats_raw)
    mold_rules = M.mold_rules(part_mold, dims)
    enforced = {r['PART']: r['MOLD'] for r in mold_rules if r['ENFORCED'] == 'بله'}
    rule_station = M.event_type_rules(station_status, dims)
    etype_by_station = {r['STATION']: r['EVENT_TYPE'] for r in rule_station}
    # قاعدهٔ EVENT_TYPE در v2 (صادقانه و با شاهدِ خودِ داده):
    #   • اگر محل‌صدور = ضایعات  → ضایعات        (بازرس صریحاً گفته)
    #   • اگر ایستگاه = برگشت از فروش → برگشتی    (تصمیم B.۵: نوعِ مستقل)
    #   • در غیر این صورت → اصلاحی
    # چرا «محل صدور» بر «قاعدهٔ ایستگاه» مقدم است: در فایلِ شما ۱٬۶۴۵ رکورد
    # DISPOSITION=ضایعات دارد ولی ایستگاهشان «کنترل نهایی» است؛ اگر فقط قاعدهٔ
    # ایستگاه را اجرا می‌کردیم، کارت «کل ضایعات» صفر می‌شد و آمار تحریف.
    disp_name = {r['CODE']: r['NAME'] for r in dims['DISPOSITION']}
    # «نوعِ رخدادِ قطعی» را روی خودِ DIM_DISPOSITION می‌نویسیم و «ردیفِ پیش‌فرض» را روی
    # RULE_STATION علامت می‌زنیم: به همین دلیل ماکرو (که فقط ASCII است) هیچ رشتهٔ فارسیِ
    # منطقی در کد ندارد و مهاجرت و ماکرو دقیقاً یک قاعده را می‌خوانند (آزمون T-W12).
    for row in dims['DISPOSITION']:
        row['EVENT_TYPE_IF_PICKED'] = row['NAME'] if row['NAME'] in ('ضایعات', 'برگشتی') else ''
    pick_rule = {r['CODE']: r['EVENT_TYPE_IF_PICKED'] for r in dims['DISPOSITION']
                 if r['EVENT_TYPE_IF_PICKED']}
    for rec in rule_station:
        rec['IS_DEFAULT'] = 1 if str(rec['BASIS']).startswith('پیش‌فرض') else 0
    default_type = next((r['EVENT_TYPE'] for r in rule_station if r['IS_DEFAULT']), 'اصلاحی')
    n_overlap = 0
    # حرفِ حفره برای هر کدِ قالب (A..E). «جهت» (راست/چپ) حفره ندارد ⇒ MOLD_CHECK
    # روی آن‌ها «قالبِ جهت» می‌شود، نه خطا — ۴۹۳ رکوردِ امروز دقیقاً این‌اند.
    cavity_of_mold = {r['CODE']: r['NAME'] for r in dims['MOLD'] if r['MOLD_KIND'] == 'حفره'}

    facts = []
    backdate = 0
    prev_g = None
    for i, r in enumerate(rows, start=1):
        bc = r['BARCODE']
        sem = M.barcode_semantics(bc)
        g = r['G_DATE']
        if g is None:
            jy = jm = jd = None
            j_text_v = ''
            mkey = wkey = ''
        else:
            jy, jm, jd = xlcal.g_to_j(g)
            j_text_v = xlcal.j_text(jy, jm, jd)
            mkey = jy * 100 + jm
            wkey = xlcal.week_key(jy, jm, jd)
        mold_name = name_of(dims, 'MOLD', r['MOLD_ID'])
        guess = sem.get('BC_CAVITY_GUESS', '')
        cavity = cavity_of_mold.get(r['MOLD_ID'], '')
        if not sem['_ok']:
            mold_check = 'بارکد نامعتبر'
        elif not cavity:
            mold_check = 'قالبِ جهت'          # راست/چپ — قابل سنجش با رقم ۴ نیست
        elif not guess:
            mold_check = 'رقمِ خارج از بازه'   # رقم ۴ بیرون از ۱..۵
        else:
            mold_check = 'هم‌خوان' if cavity == guess else 'ناهم‌خوان'
        # — ردیفِ ایرادهای کیفیت داده (همان چیزی که در شیت DQ دیده می‌شود)
        if r['MOLD_ID'] in [x['CODE'] for x in dims['MOLD']] and enforced.get(r['PART']) and \
                name_of(dims, 'MOLD', enforced[r['PART']]) != mold_name:
            dq.append({'RULE': 'DQ-MOLD-RULE', 'SEV': 'کم', 'ROW': r['_row'],
                       'DETAIL': f"قطعه {r['PART']} ↔ قالب {mold_name} (قاعده: {enforced[r['PART']]})"})
        if prev_g is not None and g is not None and g < prev_g:
            backdate += 1
            dq.append({'RULE': 'DQ-BACKDATE', 'SEV': 'کم', 'ROW': r['_row'],
                       'DETAIL': f'تاریخِ وقوع از ردیفِ پیشین کوچک‌تر است ({g} < {prev_g})'})
        if g is not None:
            prev_g = g
        facts.append({
            'EVENT_ID': f'E-{i:06d}',
            'EVENT_DATE_J': j_text_v,
            'EVENT_DATE_G': g.isoformat() if g else '',
            'J_MONTH_KEY': mkey,
            'J_WEEK_KEY': wkey,
            'SHIFT_ID': r.get('SHIFT_CODE') or f"S{r['SHIFT_RAW']}" if str(r.get('SHIFT_RAW') or '').isdigit() else 'S?',
            'STATION_ID': r['STATION_ID'],
            'STATION_SNAP': name_of(dims, 'STATION', r['STATION_ID']),
            'PART_ID': r['PART_ID'],
            'PART_NAME_SNAP': name_of(dims, 'PART', r['PART_ID']),
            'MOLD_ID': r['MOLD_ID'],
            'MOLD_NAME_SNAP': mold_name,
            'DEFECT_ID': r['DEFECT_ID'],
            'DEFECT_NAME_SNAP': name_of(dims, 'DEFECT', r['DEFECT_ID']),
            'DISPOSITION_ID': r['DISPOSITION_ID'],
            'EVENT_TYPE': (pick_rule.get(r['DISPOSITION_ID'])
                           or etype_by_station.get(name_of(dims, 'STATION', r['STATION_ID']))
                           or default_type),
            'BARCODE': bc,
            'BC_PART_CODE': sem['BC_PART_CODE'],
            'MOLD_CHECK': mold_check,
            'ENTRY_USER': 'MIGRATION',
            # بازرس از ستون «اپراتور» فایلِ مرجع می‌آید (TRIM + alias، بدونِ ادغام):
            # بدونِ این دو ستون، R09 فقط یک مقدار «MIGRATION» داشت و تاریخِ ۲۱ بازرس دور ریخته می‌شد.
            'INSPECTOR_ID': r.get('INSPECTOR_ID') or '—',
            'INSPECTOR_SNAP': name_of(dims, 'INSPECTOR', r.get('INSPECTOR_ID')),
            # فیلدهای افزوده (خارج از ۲۰ ستونِ نمایشیِ جدول، برای DQ و ردیابی):
            '_J_YEAR': jy, '_J_MONTH': jm, '_J_DAY': jd,
            '_ENTRY_CHANNEL': 'مهاجرت', '_ROW_STATUS': 'تأییدشده', '_NOTE': '',
            '_SRC_ROW': r['_row'],
        })
    # ردیف‌های تکراریِ بارکد → ROW_STATUS = مشکوک-تکراری (حذف نمی‌شوند)
    by_bc = collections.defaultdict(list)
    for idx, f in enumerate(facts):
        if f['BARCODE']:
            by_bc[f['BARCODE']].append(idx)
    dup_groups = 0
    for bc, idxs in by_bc.items():
        if len(idxs) > 1:
            dup_groups += 1
            for k, idx in enumerate(idxs):
                if k:
                    facts[idx]['_ROW_STATUS'] = 'مشکوک-تکراری'
            exact = sum(1 for k, idx in enumerate(idxs)
                        if facts[idx]['EVENT_DATE_J'] == facts[idxs[0]]['EVENT_DATE_J']
                        and facts[idx]['STATION_ID'] == facts[idxs[0]]['STATION_ID'])
            dq.append({'RULE': 'DQ-DUP', 'SEV': 'شدید' if exact == len(idxs) else 'متوسط',
                       'ROW': facts[idxs[0]]['_SRC_ROW'],
                       'DETAIL': f'بارکد {bc} در {len(idxs)} ردیف'
                                 + (f' — {exact} ردیف کاملاً یکسان' if exact else '')})
    # رکوردهایی که نامِ آلوده (فاصلهٔ اضافی) در ورودی داشته‌اند، «بازبینی‌شدنی» علامت
    # می‌خورند: ما نام را نرمال کردیم و اثرِ آن را در DQ ثبت کردیم، ولی مقدارِ
    # اصلیِ بازرس را جعل نمی‌کنیم.
    # رکوردهایی که نامِ آلوده (فاصلهٔ اضافی) داشتند، در ROW_STATUS علامت *نمی‌خورند*:
    # مجموعهٔ مقادیرِ ROW_STATUS باید بسته بماند (۴ مقدارِ شناخته‌شده) تا شمارش‌ها
    # و لیستِ کشوییِ Data گمراه‌کننده نشوند. اثرِ آن‌ها فقط در DQ-SPACE ثبت شده.
    stats['space_flagged_rows'] = len({d['ROW'] for d in dq if d['RULE'] == 'DQ-SPACE'})
    stats['event_type_vs_disposition_conflict'] = sum(
        1 for f in facts
        if disp_name.get(f['DISPOSITION_ID']) and disp_name[f['DISPOSITION_ID']] != f['EVENT_TYPE'])
    stats.update(dict(
        fact_rows=len(facts), dq_issues=len(dq), dup_groups=dup_groups,
        backdate_rows=backdate, unmapped=sum(1 for f in facts if f['PART_ID'].startswith('X-')),
        mold_mismatch=sum(1 for f in facts if f['MOLD_CHECK'] == 'ناهم‌خوان'),
        event_types=dict(collections.Counter(f['EVENT_TYPE'] for f in facts)),
        stations_used=dict(collections.Counter(f['STATION_SNAP'] for f in facts)),
    ))
    code_name = {}
    for kind, key in (('PART', 'PART'), ('MOLD', 'MOLD'), ('DEFECT', 'DEFECT'),
                      ('INSPECTOR', 'INSPECTOR'), ('STATION', 'STATION'),
                      ('DISPOSITION', 'DISPOSITION'), ('SHIFT', 'SHIFT')):
        code_name[kind] = {r['CODE']: r['NAME'] for r in dims[key]}
    payload = dict(
        generated=datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        source=os.path.basename(ref_path),
        fact_columns=[c for c, _ in FACT_COLUMNS],
        fact_headers=dict(FACT_COLUMNS),
        facts=facts, dims={k: v for k, v in dims.items() if k != '_ALIAS'},
        inspector_alias=dims['_ALIAS'],
        map_part_mold=mold_rules, rule_station=rule_station, dq_issues=dq,
        reports=REPORTS, kpi=KPI, filters=FILTERS, filter_labels=FILTER_LABELS,
        mold_check_options=MOLD_CHECK_OPTIONS, row_status_options=ROW_STATUS_OPTIONS,
        code_name=code_name, stats=stats, max_fact_rows=MAX_FACT_ROWS,
        calendar=dict(start=xlcal.CAL_START.isoformat(), end=xlcal.CAL_END.isoformat()),
    )
    if verbose:
        print('facts:', len(facts), '| dq issues:', len(dq), '| dup groups:', dup_groups,
              '| backdate:', backdate)
        print('event types:', stats['event_types'])
        print('mold check ناهم‌خوان:', stats['mold_mismatch'])
    return payload


def pickle_dump(path, obj):
    import pickle
    with open(path, 'wb') as fh:
        pickle.dump(obj, fh, protocol=4)


def pickle_load(path):
    import pickle
    with open(path, 'rb') as fh:
        return pickle.load(fh)


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('src', nargs='?', default='FORM-QC-F-14.xlsm')
    ap.add_argument('-o', '--out', default='analysis/payload.json')
    ap.add_argument('--cache', default='analysis/cache/ref.pkl')
    ap.add_argument('--force', action='store_true')
    a = ap.parse_args()
    if a.force and os.path.exists(a.cache):
        os.remove(a.cache)
    pl = build(a.src, cache=a.cache)
    with open(a.out, 'w', encoding='utf-8') as fh:
        json.dump(pl, fh, ensure_ascii=False)
    print('written:', a.out, os.path.getsize(a.out), 'bytes')
