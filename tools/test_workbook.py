"""
آزمونِ ساختاریِ فایلِ ساخته‌شده (QC-F-14-v2.xlsx) — قبل از هر حرفه‌ای‌تر از این،
«این فایل را باز نکنید» را می‌گوید.

    python3 -W ignore tools/test_workbook.py QC-F-14-v2.xlsx [--payload analysis/payload.json]

T-W1  فایل با openpyxl (read_only و normal) بدونِ خطا باز می‌شود؛ شیت‌های لازم موجودند
T-W2  جدول‌های tblQC / DIM_* / MAP_PART_MOLD / RULE_STATION موجود و refِ معتبر دارند
T-W3  ردیف‌های داده با payload «بایت‌به‌بایت» می‌خوانند (بارکد، تاریخ، کدها، نوع)
T-W4  ستونِ بارکد در همۀ ردیف‌هایِ پر، رشته و ۱۹ رقمی است (صفرِ ابتدا زنده است)
T-W5  هیچ فرمولی روی شیت Data نیست (قاعدهٔ «داده بدون محاسبه»)
T-W6  هر کدِ مصرف‌شده در Data در Master وجود دارد (مگر «X-UNMAPPED» که باید صفر باشد)
T-W7  DVها به نام‌هایِ تعریف‌شده اشاره می‌کنند؛ نام‌ها به محدودهٔ معتبر می‌رسند
T-W8  فرمول‌های COUNTIFS کرانه‌دارند (تا CAP_ROWS) و به ستونِ درست ارجاع می‌دهند
T-W9  مدلِ شمارش (ارزیابِ COUNTIFS همین تست) = شمارشِ مستقیمِ payload، صفر اختلاف
T-W10 داشبورد: ۱۰ فیلتر، ۶ کارت، ۸ نمودار؛ گزارش‌ها: ۱۵ بلوک با سطرِ «جمعِ بلوک»
T-W11 تقویمِ Helper: هر J_TEXT یکتاست و با ستونِ میلادی می‌خواند (بدون locale)
"""
import argparse
import collections
import json
import os
import re
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlcal                                    # noqa: E402

# انتظاراتِ مستقلِ آزمون (از gen_workbook نمی‌خوانیم؛ اگر فایلِ دیگری این ترتیب را
# ندهد، T-W3 باید قرمز شود). ترتیبِ واقعی از سرستونِ خودِ فایل خوانده و مقایسه می‌شود.
MUST_COLS = ['EVENT_ID', 'EVENT_DATE_J', 'EVENT_DATE_G', 'J_MONTH_KEY', 'J_WEEK_KEY',
             'SHIFT_ID', 'STATION_ID', 'STATION_SNAP', 'PART_ID', 'PART_NAME_SNAP',
             'MOLD_ID', 'MOLD_NAME_SNAP', 'DEFECT_ID', 'DEFECT_NAME_SNAP',
             'DISPOSITION_ID', 'EVENT_TYPE', 'BARCODE', 'BC_PART_CODE', 'MOLD_CHECK',
             'ENTRY_STAMP', 'ENTRY_USER', 'INSPECTOR_ID', 'INSPECTOR_SNAP',
             'ENTRY_CHANNEL', 'ROW_STATUS', 'NOTE']
FACT_COLS = MUST_COLS
COL = {c: i + 1 for i, c in enumerate(FACT_COLS)}
EXPECTED_SHEETS = ['FORM فرم ثبت', 'Data داده', 'STAGING ورود اضطراری', 'Master اطلاعات پایه',
                   'Helper محاسبات', 'Reports گزارش‌ها', 'Dashboard داشبورد', 'DQ کیفیت داده',
                   'Settings تنظیمات', 'Help راهنما', 'LOG_AUDIT', 'Verify']
fails, notes = [], []


def check(name, cond, detail=''):
    (notes if cond else fails).append(f'{"✓" if cond else "✗"} {name}' + (f' — {detail}' if detail else ''))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('file', nargs='?', default='QC-F-14-v2.xlsx')
    ap.add_argument('--payload', default='analysis/payload.json')
    a = ap.parse_args()
    pl = json.load(open(a.payload, encoding='utf-8'))
    path = a.file
    check('T-W0 فایل وجود دارد', os.path.exists(path), path)
    if not os.path.exists(path):
        report()
        return
    wb = openpyxl.load_workbook(path)
    wbr = openpyxl.load_workbook(path, read_only=True)
    names = wb.sheetnames
    missing = [s for s in EXPECTED_SHEETS if s not in names]
    check('T-W1 شیت‌ها', not missing, 'ناقص: ' + ', '.join(missing) if missing else f'{len(names)} شیت')
    ws = wb[names[names.index('Data داده')] if 'Data داده' in names else 1]
    tables = {t: wb[ws.title].tables[t] for t in wb[ws.title].tables}
    check('T-W2 tblQC', 'tblQC' in tables, tables.get('tblQC').ref if 'tblQC' in tables else '')
    ws_m = wb['Master اطلاعات پایه']
    mt = set(ws_m.tables)
    need = {'DIM_PART', 'DIM_MOLD', 'DIM_DEFECT', 'DIM_INSPECTOR', 'DIM_STATION', 'DIM_SHIFT',
            'DIM_DISPOSITION', 'MAP_PART_MOLD', 'RULE_STATION'}
    check('T-W2 جداولِ Master', need <= mt, 'ناقص: ' + ', '.join(sorted(need - mt)) if need - mt else f'{len(mt)} جدول')

    # T-W3/T-W4/T-W5/T-W6
    n = len(pl['facts'])
    header = [c.value for c in ws[1]]
    global FACT_COLS, COL
    file_cols = [(str(h.value) if h.value else '').split('\n')[-1].strip() for h in ws[1]]
    file_cols = [c for c in file_cols if c]
    check(f'T-W3 سرستونِ {len(MUST_COLS)}تایی', len(file_cols) == len(MUST_COLS), f'{len(file_cols)}')
    check('T-W3 ترتیبِ ستون‌ها = انتظارِ آزمون', file_cols == MUST_COLS,
          f'اختلاف در: {[i for i, (a, b) in enumerate(zip(file_cols, MUST_COLS)) if a != b][:4]}')
    FACT_COLS = file_cols or MUST_COLS
    COL = {c: i + 1 for i, c in enumerate(FACT_COLS)}
    bad_row = 0
    bad_bc = 0
    formula_rows = 0
    unmapped = 0
    codes = {k: {r['CODE'] for r in v} for k, v in pl['dims'].items() if isinstance(v, list)}
    grid = list(ws.iter_rows(min_row=2, max_row=1 + n, max_col=len(FACT_COLS), values_only=True))
    def norm(x):
        import datetime as _dt
        if isinstance(x, _dt.datetime):
            return x.date()
        return x

    for i, f in enumerate(pl['facts'], start=2):
        r = [norm(x) for x in grid[i - 2]]
        exp = [f['EVENT_ID'], f['EVENT_DATE_J'],
               __import__('datetime').date.fromisoformat(f['EVENT_DATE_G']) if f['EVENT_DATE_G'] else '',
               f['J_MONTH_KEY'], f['J_WEEK_KEY'], f['SHIFT_ID'], f['STATION_ID'], f['STATION_SNAP'],
               f['PART_ID'], f['PART_NAME_SNAP'], f['MOLD_ID'], f['MOLD_NAME_SNAP'], f['DEFECT_ID'],
               f['DEFECT_NAME_SNAP'], f['DISPOSITION_ID'], f['EVENT_TYPE'], f['BARCODE'],
               f['BC_PART_CODE'], f['MOLD_CHECK'], '', f['ENTRY_USER'],
               f['INSPECTOR_ID'], f['INSPECTOR_SNAP'], f['_ENTRY_CHANNEL'],
               f['_ROW_STATUS'], f['_NOTE']]
        got = [('' if v is None else v) for v in r]
        exp = [('' if v is None else v) for v in exp]
        if [str(x) for x in got] != [str(x) for x in exp]:
            bad_row += 1
            if bad_row < 3:
                fails.append(f'   سطر {i}: {[(g, e) for g, e in zip(got, exp) if str(g) != str(e)][:3]}')
        bc = r[COL['BARCODE'] - 1]
        if not isinstance(bc, str) or len(bc) != 19 or not bc.isdigit():
            bad_bc += 1
        for key, dim in (('PART_ID', 'PART'), ('MOLD_ID', 'MOLD'), ('DEFECT_ID', 'DEFECT'),
                        ('STATION_ID', 'STATION'), ('DISPOSITION_ID', 'DISPOSITION'),
                        ('SHIFT_ID', 'SHIFT'), ('INSPECTOR_ID', 'INSPECTOR')):
            if r[COL[key] - 1] not in codes[dim]:
                unmapped += 1
        if any(isinstance(x, str) and x.startswith('=') for x in r):
            formula_rows += 1
    check(f'T-W3 خوانشِ {n} ردیف با payload', bad_row == 0, f'{bad_row} ردیف نخوانده')
    check('T-W4 بارکد = متنِ ۱۹ رقمی', bad_bc == 0, f'{bad_bc} بد')
    check('T-W5 بدون فرمول روی Data', formula_rows == 0, f'{formula_rows} ردیف فرمولی')
    check('T-W6 ارجاعِ سالم به Master', unmapped == 0, f'{unmapped} کدِ بیرونِ Master')
    # بازرس (اپراتورِ فایلِ مرجع) باید در tblQC نشسته باشد — اگر نبود، R09 فقط یک سطر می‌شد
    ins_col = COL['INSPECTOR_ID'] - 1
    ins_file = {row[ins_col] for row in grid if row[ins_col] not in (None, '')}
    ins_model = {f['INSPECTOR_ID'] for f in pl['facts'] if f['INSPECTOR_ID'] != '—'}
    check('T-W6 بازرس مهاجرت کرد (R09/CH8 زنده‌اند)',
          ins_file == ins_model and len(ins_model) >= 10,
          f'فایل={len(ins_file)} مدل={len(ins_model)}')
    lead_rows = sum(1 for row in grid if str(row[COL['BARCODE'] - 1]).startswith('0'))
    lead_dist = len({row[COL['BARCODE'] - 1] for row in grid
                     if str(row[COL['BARCODE'] - 1]).startswith('0')})
    check('T-W4 صفرهای ابتدایی حفظ شده', lead_dist == pl['stats']['leading_zero'],
          f'متمایز={lead_dist} (منتظر {pl["stats"]["leading_zero"]}) / ردیف={lead_rows}')

    # T-W7 defined names
    dn = set(wb.defined_names)
    want = {'shift_list', 'station_list', 'part_list', 'mold_list', 'defect_list',
            'inspector_list', 'disposition_list', 'date_list', 'N_ROWS', 'CAP_ROWS'}
    check('T-W7 نام‌هایِ تعریف‌شده', want <= dn, 'ناقص: ' + ', '.join(sorted(want - dn)))
    ws_h = wb['Helper محاسبات']
    for nm in sorted(want & dn):
        ref = wb.defined_names[nm].value or ''
        m = re.match(r"'?([^'!]+)'?!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?", ref)
        ok = bool(m) and m.group(1) in wb.sheetnames
        if ok and m.group(5):
            ok = int(m.group(5)) <= ws_h.max_row
        if ok and not m.group(5) and m.group(1) == ws_h.title:
            ok = int(m.group(3)) <= ws_h.max_row
        if not ok:
            fails.append(f'✗ T-W7 محدودهٔ نام {nm} نامعتبر: {ref}')

    # ── T-W8/T-W9 : کرانه‌ها + ارزیابِ مدلِ COUNTIFS ──
    cap = nrow_h = None
    for row in ws_h.iter_rows(min_row=2, max_row=20, min_col=1, max_col=45):
        for c in row:
            if c.value == 'CAP_ROWS':
                cap = ws_h.cell(row=c.row, column=c.column + 1).value
            if c.value == 'N_ROWS':
                nrow_h = ws_h.cell(row=c.row, column=c.column + 1).value
    over = []
    used_cols = set()
    for sh in ('Reports گزارش‌ها', 'Dashboard داشبورد', 'DQ کیفیت داده', 'FORM فرم ثبت',
               'STAGING ورود اضطراری'):
        for row in wb[sh].iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith('='):
                    if '#REF' in v:
                        over.append(f'{sh}!{c.coordinate}: #REF')
                    for m in re.finditer(r'Data!\$([A-Z]+)\$?(\d+):\$[A-Z]+\$?(\d+)', v):
                        used_cols.add(m.group(1))
                        if cap and int(m.group(3)) > cap:
                            over.append(f'{sh}!{c.coordinate}: سطر {m.group(3)} > CAP {cap}')
    check('T-W8 کرانه‌دار و بدون #REF', not over, '; '.join(over[:3]))
    check('T-W8 ارجاع فقط به ستون‌های معتبر', used_cols <= set(
        openpyxl.utils.get_column_letter(i + 1) for i in range(len(FACT_COLS))),
        'بیرونِ جدول: ' + ', '.join(sorted(used_cols - set(
            openpyxl.utils.get_column_letter(i + 1) for i in range(len(FACT_COLS))))))

    wd = wb['Dashboard داشبورد']
    grid_all = [list(r) for r in ws.iter_rows(min_row=2, max_row=1 + n, max_col=len(FACT_COLS),
                                              values_only=True)]

    def split_args(text):
        out, depth, cur, q = [], 0, '', False
        for ch in text:
            if ch == '"':
                q = not q
                cur += ch
            elif ch == '(' and not q:
                depth += 1; cur += ch
            elif ch == ')' and not q:
                depth -= 1; cur += ch
            elif ch == ',' and depth == 0 and not q:
                out.append(cur.strip()); cur = ''
            else:
                cur += ch
        if cur.strip():
            out.append(cur.strip())
        return out

    def evaluate(formula):
        """ارزیابِ COUNTIFSِ همون ساختاری که این فایل تولید می‌کند:
        جفت‌های (محدودهٔ Data، معیار). معیارهای IF(...="همه",...) از Dashboard خوانده می‌شود."""
        f = (formula or '').strip()
        m = re.fullmatch(r'=COUNTIFS\((.+)\)', f, re.S)
        if not m:
            return None
        args = split_args(m.group(1))
        if len(args) % 2:
            return None
        conds = []
        for rng, crit in zip(args[0::2], args[1::2]):
            mm = re.fullmatch(r'Data!\$([A-Z]+)\$?(\d+):\$[A-Z]+\$?(\d+)', rng)
            if not mm:
                return None
            idx = openpyxl.utils.column_index_from_string(mm.group(1)) - 1
            r0, r1 = int(mm.group(2)), int(mm.group(3))
            if crit.startswith('IF(Dashboard!'):
                g = re.search(r'\$G\$(\d+)', crit)
                fv = wd[f'G{int(g.group(1))}'].value
                if fv is None or str(fv) == 'همه':
                    continue
                v = str(fv)[:4] if crit.count('LEFT(') else str(fv)
                conds.append((idx, r0, r1, v))
            elif crit.startswith('"'):
                conds.append((idx, r0, r1, crit.strip('"')))
            elif re.fullmatch(r'-?\d+(\.\d+)?', crit):
                conds.append((idx, r0, r1, crit))
            else:
                return None
        cnt = 0
        for ri, row in enumerate(grid_all, start=2):
            ok = True
            for idx, r0, r1, v in conds:
                if not (r0 <= ri <= r1):
                    continue
                val = row[idx]
                if val is None or str(val) != v:
                    ok = False
                    break
            if ok:
                cnt += 1
        return cnt

    # فیلترِ پیش‌فرضِ داشبورد: ۹ فیلتر «همه»، فقط ROW_STATUS = تأییدشده
    def model_match(f, key=None, val=None):
        if f['_ROW_STATUS'] != 'تأییدشده':
            return False
        if key is not None and f[key] != val:
            return False
        return True

    ok1 = evaluate(wd['C7'].value)
    direct = sum(1 for f in pl['facts'] if model_match(f))
    check('T-W9 کارت K1 (COUNTIFS) = مدل', ok1 == direct, f'فرمول→{ok1} مدل→{direct}')
    # چند کارت/سطرِ گزارش را هم می‌سنجد
    tested = 1
    for row in wd.iter_rows(min_row=7, max_row=12, min_col=3, max_col=3):
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('=COUNTIFS('):
                e = evaluate(c.value)
                key = {'C8': 'ضایعات', 'C9': 'اصلاحی', 'C10': 'برگشتی'}.get(c.coordinate)
                if key and e is not None:
                    d = sum(1 for f in pl['facts'] if model_match(f, 'EVENT_TYPE', key))
                    if e != d:
                        fails.append(f'✗ T-W9 {c.coordinate} ({key}): فرمول→{e} مدل→{d}')
                    tested += 1
    check(f'T-W9 ارزیابی {tested} عددِ کلیدی با مدل', not any('T-W9' in f for f in fails))
    # T-W10
    wd = wb['Dashboard داشبورد']
    n_filters = sum(1 for r in range(4, 14) if wd[f'F{r}'].value and wd[f'G{r}'].value)
    n_cards = sum(1 for r in range(7, 13) if wd[f'B{r}'].value and str(wd[f'C{r}'].value).startswith('='))
    n_charts = len(wd._charts)
    check('T-W10 داشبورد', n_filters == 10 and n_cards == 6 and n_charts == 8,
          f'فیلتر={n_filters} کارت={n_cards} نمودار={n_charts}')
    wr = wb['Reports گزارش‌ها']
    n_blocks = sum(1 for row in wr.iter_rows(min_col=2, max_col=2) for c in row
                   if isinstance(c.value, str) and re.match(r'^R\d\d — ', c.value))
    n_tot = sum(1 for row in wr.iter_rows(min_col=2, max_col=2) for c in row
                if isinstance(c.value, str) and c.value.startswith('جمعِ بلوک'))
    check('T-W10 ۱۵ گزارش با سطرِ تراز', n_blocks == 15 and n_tot == 10,
          f'بلوک={n_blocks} سطرِ تراز={n_tot} (۱۰ گزارشِ جدولی + ۵ گزارشِ خاص)')
    # T-W11
    cal = list(ws_h.iter_rows(min_row=2, max_row=ws_h.max_row, min_col=1, max_col=10, values_only=True))
    cal = [r for r in cal if r[0] is not None]
    jt = [r[1] for r in cal]
    ok = len(set(jt)) == len(jt)
    bad = 0
    for r in cal[:2000]:
        g, txt = r[0].date() if hasattr(r[0], 'date') else r[0], r[1]
        if xlcal.j_text(*xlcal.g_to_j(g)) != txt:
            bad += 1
    check('T-W11 تقویمِ Helper', ok and bad == 0, f'{len(cal)} روز، یکتا={ok}، بد={bad}')

    # ── T-W12: قاعده‌ها باید «در جدول» باشند، نه در کدِ ماکرو/اسکریپت ──────────
    def table_values(wsM, tname):
        t = wsM.tables.get(tname)
        if t is None:
            return []
        m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', t.ref)
        ci = openpyxl.utils.column_index_from_string
        c0, r0, c1, r1 = ci(m.group(1)), int(m.group(2)), ci(m.group(3)), int(m.group(4))
        return [[wsM.cell(row=r, column=c0 + k).value for k in range(c1 - c0 + 1)]
                for r in range(r0 + 1, r1 + 1)]

    disp = table_values(ws_m, 'DIM_DISPOSITION')
    check('T-W12 DIM_DISPOSITION ستونِ قاعده را دارد',
          all(len(r) >= 4 for r in disp) and len(disp) == 3,
          f'{len(disp)} ردیف، عرضِ {max((len(r) for r in disp), default=0)}')
    pick = {r[0]: r[3] for r in disp if len(r) > 3 and r[3]}
    check('T-W12 فقط ضایعات/برگشتی قطعی‌اند',
          pick.get('DIS-02') == 'ضایعات' and set(pick.values()) <= {'ضایعات', 'برگشتی'},
          str(pick))
    rs = table_values(ws_m, 'RULE_STATION')
    rule = {r[0]: r[2] for r in rs if r and r[0]}
    st_all = {r[0] for r in table_values(ws_m, 'DIM_STATION') if r and r[0]}
    # «پیش‌فرض؟» یعنی «این ایستگاه در دادهٔ تاریخی شاهد ندارد» — نه «ردیفِ پیش‌فرضِ سراسری».
    # شش ایستگاهِ غیرفعال دقیقاً همین‌اند و همه به اصلاحی می‌روند (قابل ویرایش در Master).
    n_def = sum(1 for r in rs if str(r[4] if len(r) > 4 else '').startswith('1'))
    check('T-W12 RULE_STATION همهٔ ایستگاه‌ها را پوشش می‌دهد',
          st_all <= set(rule) and n_def == 6 and len(rs) == 8,
          f'پوشش={len(st_all & set(rule))}/{len(st_all)} بی‌شاهد={n_def} ردیف‌ها={len(rs)}')
    # بازسازیِ EVENT_TYPE فقط از جدول‌های Master و مقایسه با Data — همان مسیری که QC_Append می‌رود
    et_default = 'اصلاحی'   # فقط برای ردیف‌هایِ مهاجرت‌یافتهٔ بی‌ایستگاه؛ ماکرو در این حالت رد می‌کند
    bad_et = 0
    di, si, ei = COL['DISPOSITION_ID'] - 1, COL['STATION_ID'] - 1, COL['EVENT_TYPE'] - 1
    for row in grid:
        want = pick.get(row[di]) or rule.get(row[si]) or et_default
        if str(row[ei]) != str(want):
            bad_et += 1
    check('T-W12 EVENT_TYPE = قاعدهٔ جدول‌ها (برای هر ردیف)', bad_et == 0,
          f'{bad_et} ردیفِ ناسازگار از {len(grid)}')
    # CODE_MOLD_CHECK باید دقیقاً همان مقدارهایِ Data باشد (ماکرو از همین گرید می‌خواند)
    mc_col = None
    for c in ws_h[1]:
        if c.value == 'CODE_MOLD_CHECK':
            mc_col = c.column
    opts = []
    if mc_col:
        for r in range(2, 20):
            v = ws_h.cell(row=r, column=mc_col).value
            if v in (None, ''):
                break
            opts.append(str(v))
    used = {str(row[COL['MOLD_CHECK'] - 1]) for row in grid if row[COL['MOLD_CHECK'] - 1]}
    check('T-W12 گریدِ CODE_MOLD_CHECK = مقدارهایِ Data', used <= set(opts) and bool(opts),
          f'گرید={len(opts)} استفاده‌شده={len(used)} نااشنا={sorted(used - set(opts))[:2]}')
    # QC_SHEETS / QC_MSG — پیکربندیِ ماکرو
    dn = {d.name: d for d in wb.defined_names.values()} if hasattr(wb.defined_names, 'values') \
        else {d.name: d for d in wb.defined_names}
    ws_s = wb['Settings تنظیمات']
    okcfg, cfgdet = True, []
    for nm in ('QC_SHEETS', 'QC_MSG'):
        if nm not in dn:
            okcfg, cfgdet = False, cfgdet + [f'{nm} نیست']
            continue
        m = re.match(r".*!\$B\$(\d+):\$C\$(\d+)", str(dn[nm].attr_text))
        if not m:
            okcfg, cfgdet = False, cfgdet + [f'{nm} بازهٔ عجیب']
            continue
        rows = [[ws_s.cell(row=r, column=2).value, ws_s.cell(row=r, column=3).value]
                for r in range(int(m.group(1)), int(m.group(2)) + 1)]
        if nm == 'QC_SHEETS':
            names = set(wb.sheetnames)
            miss = [k for k, v in rows if str(v) not in names]
            if miss:
                okcfg, cfgdet = False, cfgdet + [f'نامِ شیتِ ناموجود: {miss}']
            cfgdet.append(f'sheet={len(rows)}')
        else:
            need = {'E_BARCODE', 'E_INCOMPLETE', 'E_CAPACITY', 'OK_APPEND', 'E_STAGE_NONE',
                    'E_CONFIG', 'E_SHEET', 'E_VOID', 'OK_VOID', 'E_DUP', 'E_NORULE',
                    'E_DATE', 'E_DUP_ACT', 'OK_IMPORT', 'ABOUT'}
            have = {str(k) for k, _v in rows}
            if need - have:
                okcfg, cfgdet = False, cfgdet + [f'پیامِ ناقص: {sorted(need - have)}']
            blank = [k for k, v in rows if not v or not str(v).strip()]
            if blank:
                okcfg, cfgdet = False, cfgdet + [f'متنِ خالی: {blank}']
            cfgdet.append(f'msg={len(rows)}')
    check('T-W12 پیکربندیِ ماکرو (QC_SHEETS/QC_MSG)', okcfg, ' ، '.join(cfgdet))
    report()


def report():
    for n in notes:
        print(n)
    if fails:
        print('FAIL')
        for f in fails:
            print(' ', f)
        sys.exit(1)
    print('WORKBOOK TESTS PASS (T-W1–T-W12)')


if __name__ == '__main__':
    main()
