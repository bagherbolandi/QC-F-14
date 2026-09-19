#!/usr/bin/env python3
"""
مهندسی معکوس / ممیزی فایل ثبت برگشتی-ضایعات-اصلاحی (Phase 1 + Phase 7 QA)
ابزار بدون وابستگی: فقط zipfile + re (استخراج مستقیم OOXML و VBA).

کاربرد:
  1) فاز ۱: استخراج واقعیات ساختاری و آماری فایل فعلی (read-only).
  2) فاز ۸ (تأیید مهاجرت): مقایسه فایل جدید با فایل مرجع → شمارش رکورد،
     یکپارچگی بارکد (رشته، ۱۹ رقم، صفر ابتدایی)، کنترل عدم‌تغییر مقادیر.
     خروجی با کد ۰ = مهاجرت صحیح؛ غیرصفر = اختلاف (جلوگیری از خرابی بی‌صدا).

用法: python3 analysis/reverse_engineer.py [مسیر فایل]
"""
import zipfile, re, sys, json, collections, datetime

CELL = re.compile(r'<c r="([A-Z]+)(\d+)"((?:[^>"]|"[^"]*")*?)(?:/>|>(.*?)</c>)', re.S)
ROW = re.compile(r'<row r="(\d+)"((?:[^>"]|"[^"]*")*?)>(.*?)</row>', re.S)
ROW_EMPTY = re.compile(r'<row r="(\d+)"((?:[^>"]|"[^"]*")*?)/>')

COLMAP = {'A': 'تاریخ', 'B': 'شیفت', 'C': 'ایستگاه', 'D': 'قطعه', 'E': 'قالب',
          'F': 'بارکد', 'G': 'عیب', 'H': 'وضعیت', 'I': 'اپراتور'}


def sst(z):
    try:
        raw = z.read('xl/sharedStrings.xml').decode('utf-8')
    except KeyError:
        return []
    return [''.join(re.findall(r'<t[^>]*>(.*?)</t>', s, re.S))
            for s in re.findall(r'<si>(.*?)</si>', raw, re.S)]


def serial_to_gregorian(n):
    return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(n))).isoformat()


def audit(path):
    z = zipfile.ZipFile(path)
    out = {'file': path, 'bytes': None, 'parts': {}, 'structure': {}, 'data': {}}
    out['bytes'] = z.fp.seek(0, 2)
    names = z.namelist()
    out['parts'] = {p: z.getinfo(p).file_size for p in names}
    wb = z.read('xl/workbook.xml').decode('utf-8')
    out['structure']['sheets'] = re.findall(r'<sheet name="([^"]+)"[^>]*?(?:state="(\w+)")?[^>]*/>', wb)
    out['structure']['defined_names'] = re.findall(r'<definedName[^>]*name="([^"]+)"[^>]*>([^<]*)</definedName>', wb)
    out['structure']['workbook_protection'] = bool(re.search(r'<workbookProtection', wb))
    out['structure']['has_tables'] = any('xl/tables/' in n for n in names)
    out['structure']['has_pivots'] = any('pivotTable' in n for n in names)
    out['structure']['has_powerquery'] = any('queries.xml' in n or ' connections.xml' in n for n in names)
    out['structure']['has_model'] = any('xl/model' in n or 'DataModel' in n for n in names)
    out['structure']['has_vba'] = 'xl/vbaProject.bin' in names
    out['structure']['external_links'] = [n for n in names if n.startswith('xl/externalLinks/')]
    try:
        out['structure']['vba_unprotected'] = (b'DPX' not in z.read('xl/vbaProject.bin')
                                               and b'DPx' not in z.read('xl/vbaProject.bin'))
    except KeyError:
        out['structure']['vba_unprotected'] = None
    for i in (1, 2, 3):
        try:
            z.read(f'xl/worksheets/sheet{i}.xml')
        except KeyError:
            continue
        d = z.read(f'xl/worksheets/sheet{i}.xml').decode('utf-8')
        info = {'xml_bytes': len(d.encode()),
                'content_rows': len(ROW.findall(d)),
                'phantom_styled_rows': len(ROW_EMPTY.findall(d)),
                'custom_format_rows': len(re.findall(r'customFormat="1"', d)),
                'hidden_rows': len(re.findall(r'hidden="1"', d)),
                'merged': len(re.findall(r'<mergeCell ', d)),
                'formulas': re.findall(r'<f>(.*?)</f>', d, re.S),
                'volatile_count': len(re.findall(r'<f[^>]*>(?:(?!</f>).)*?\b(TODAY|NOW|OFFSET|INDIRECT|RAND|CELL)\b', d, re.S)),
                'rtl': 'rightToLeft="1"' in d,
                'gridlines_off': 'showGridLines="0"' in d,
                'headers_off': 'showRowColHeaders="0"' in d,
                'protection': (re.search(r'<sheetProtection[^>]*/>', d).group(0)
                               if re.search(r'<sheetProtection', d) else None),
                'print_scale': (lambda m: m.group(1) if m else None)(
                    re.search(r'<pageSetup[^>]*scale="(\d+)"', d)),
                'dimension': (re.search(r'<dimension ref="([^"]+)"', d).group(1)
                              if re.search(r'<dimension', d) else None)}
        out['structure'][f'sheet{i}'] = info
    out['structure']['sheet1_dv_rules'] = re.findall(
        r'<dataValidation [^>]*?/>|<dataValidation .*?</dataValidation>',
        z.read('xl/worksheets/sheet1.xml').decode('utf-8'), re.S)
    return out


def data_stats(path, sheet='xl/worksheets/sheet2.xml', first_data_row=2):
    z = zipfile.ZipFile(path)
    S = sst(z)
    d = z.read(sheet).decode('utf-8')
    rows = collections.defaultdict(dict)
    for r, attrs, body in ROW.findall(d):
        rr = int(r)
        if rr < first_data_row:
            continue
        for m in CELL.finditer(body):
            col, ref, attrs_c, inner = m.group(1), m.group(2), m.group(3), m.group(4) or ''
            t = re.search(r'\bt="(\w+)"', attrs_c)
            t = t.group(1) if t else 'n'
            v = re.search(r'<v>(.*?)</v>', inner, re.S)
            if not v:
                continue
            val = S[int(v.group(1))] if t == 's' else v.group(1)
            rows[rr][col] = val
    recs = [{'row': rr, **{name: rows[rr].get(col, '') for col, name in COLMAP.items()}}
            for rr in sorted(rows) if rows[rr]]
    st = {'records': len(recs)}
    st['blank_rows'] = [r['row'] for r in recs
                        if all(str(r[k]).strip() == '' for k in COLMAP.values())]
    for col in COLMAP.values():
        st[f'distinct_{col}'] = len({str(r[col]).strip() for r in recs if str(r[col]).strip() != ''})
    st['whitespace_pollution'] = {col: collections.Counter(
        r[col] for r in recs if r[col] and r[col] != r[col].strip())
        for col in ['اپراتور', 'قطعه', 'قالب', 'عیب', 'وضعیت', 'ایستگاه']}
    st['whitespace_pollution'] = {k: dict(v) for k, v in st['whitespace_pollution'].items() if v}
    bc = collections.Counter(str(r['بارکد']).strip() for r in recs if str(r['بارکد']).strip())
    st['barcode'] = {'n': sum(bc.values()), 'distinct': len(bc),
                     'len_hist': dict(collections.Counter(len(k) for k in bc.elements())),
                     'non_digit': sum(1 for k in bc if not k.isdigit()),
                     'leading_zero': sum(1 for k in bc if k.startswith('0')),
                     'dup_groups': sum(1 for v in bc.values() if v > 1),
                     'dup_extra_rows': sum(v - 1 for v in bc.values() if v > 1),
                     'max_repeat': max(bc.values()) if bc else 0}
    bykey = collections.defaultdict(list)
    for r in recs:
        bykey[tuple(str(r[c]).strip() for c in COLMAP.values())].append(r['row'])
    groups = [v for v in bykey.values() if len(v) > 1]
    st['exact_duplicates'] = {'groups': len(groups), 'extra_rows': sum(len(v) - 1 for v in groups),
                              'consecutive_groups': sum(1 for v in groups if max(v) - min(v) == len(v) - 1)}
    ds = collections.Counter()
    for r in recs:
        try:
            ds[serial_to_gregorian(float(r['تاریخ']))] += 1
        except Exception:
            pass
    st['dates'] = {'days': len(ds), 'min': min(ds) if ds else None, 'max': max(ds) if ds else None,
                   'per_day_avg': round(sum(ds.values()) / len(ds), 1) if ds else None,
                   'per_day_max': max(ds.values()) if ds else None,
                   'projection_300_workdays': round(sum(ds.values()) / len(ds) * 300) if ds else None}
    def pair(k1, k2):
        c = collections.Counter((r[k1].strip(), r[k2].strip()) for r in recs)
        return {f"{a} | {b}": n for (a, b), n in c.most_common(12)}
    st['crosstab'] = {'وضعیت×ایستگاه': pair('وضعیت', 'ایستگاه'), 'وضعیت×شیفت': pair('وضعیت', 'شیفت')}
    pref = collections.defaultdict(collections.Counter)
    for r in recs:
        b = str(r['بارکد']).strip()
        if len(b) == 19:
            pref[r['قطعه'].strip()][b[:2]] += 1
    st['barcode_prefix_part_purity'] = {
        p: {'n': sum(c.values()), 'codes': dict(c.most_common(3)),
            'top_share': round(c.most_common(1)[0][1] / max(1, sum(c.values())) * 100, 1)}
        for p, c in sorted(pref.items(), key=lambda x: -sum(x[1].values()))[:12]}
    return st, recs


def migration_check(reference_path, new_path, new_sheet='xl/worksheets/sheet2.xml',
                    new_first_row=2):
    """مقایسه رکورد-به-رکورد فایل مرجع با فایل جدید؛ خروجی ۰ یعنی مهاجرت بدون افت."""
    ref, ref_recs = data_stats(reference_path)
    new, new_recs = data_stats(new_path, sheet=new_sheet, first_data_row=new_first_row)
    fails = []
    if ref['records'] != new['records']:
        fails.append(f"شمار رکورد: مرجع {ref['records']} ≠ جدید {new['records']}")
    k = lambda r, c: str(r.get(c, '')).strip()
    if ref['barcode']['n'] != new['barcode']['n']:
        fails.append("تعداد بارکدهای ثبت‌شده متفاوت است")
    if ref['barcode']['leading_zero'] != new['barcode']['leading_zero']:
        fails.append("تعداد بارکدهای دارای صفر ابتدایی تغییر کرده ← ذخیره عددی رخ داده")
    if ref['barcode']['len_hist'] != new['barcode']['len_hist']:
        fails.append(f"توزیع طول بارکد تغییر کرده: {ref['barcode']['len_hist']} → {new['barcode']['len_hist']}")
    bref = collections.Counter(k(r, 'بارکد') for r in ref_recs if k(r, 'بارکد'))
    bnew = collections.Counter(k(r, 'بارکد') for r in new_recs if k(r, 'بارکد'))
    missing = set(bref) - set(bnew)
    if missing:
        fails.append(f"{len(missing)} بارکد در فایل جدید یافت نشد (نمونه: {sorted(missing)[:3]})")
    return fails, ref, new


if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else 'FORM-QC-F-14.xlsm'
    a = audit(path)
    st, recs = data_stats(path)
    rep = {
        'audit': {kk: vv for kk, vv in a['structure'].items() if not kk.startswith('sheet')},
        'sheet1_protection': a['structure']['sheet1']['protection'],
        'sheet1_print_scale': a['structure']['sheet1']['print_scale'],
        'sheet1_formulas': a['structure']['sheet1']['formulas'],
        'sheet1_volatile': a['structure']['sheet1']['volatile_count'],
        'sheet2': {kk: a['structure']['sheet2'][kk] for kk in
                   ['xml_bytes', 'content_rows', 'phantom_styled_rows', 'custom_format_rows',
                    'merged', 'dimension', 'protection', 'rtl', 'print_scale']},
        'data': st,
        'parts': {p: s for p, s in a['parts'].items() if s > 1000},
    }
    print(json.dumps(rep, ensure_ascii=False, indent=1, default=dict))
