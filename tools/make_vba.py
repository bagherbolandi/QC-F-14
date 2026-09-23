"""دروازهٔ ماژولِ VBA — «آزمونِ» کدی که در Excel اجرا می‌شود، بدونِ Excel.

چرا این‌قدر سخت‌گیر: تنها چیزی که می‌توانیم راستی‌آزمایی کنیم، *قراردادها* است —
۱) منبع باید خالصاً ASCII باشد (VBE یک ویرایشگرِ ANSI است؛ فارسیِ درونِ .bas هنگامِ
   Import با cp1256 خراب می‌شود و ZWNJ/U+06CC اصلاً قابلِ ذخیره نیستند).
۲) هر رشتهٔ فارسیِ قابلِ مشاهده باید از فایل خوانده شود، پس هر کلیدِ Msg()/Label()
   باید در Settings!QC_MSG وجود داشته باشد.
۳) هر نامِ ستون/نامِ جدول/نامِ تعریف‌شده که کد به آن تکیه می‌کند باید در فایلِ
   ساخته‌شده پیدا شود (وگرنه ماکرو در Excel «subscript out of range» می‌گیرد).
۴) ثابت‌هایِ اندیسِ گرید (ST_OK/MC_MISS/…) باید با ترتیبِ گریدها در فایل یکی باشند.
۵) مسیرِ نوشتن: ممنوعیتِ حذف/ویرایشِ تاریخچه، Unprotect، SaveAs، Shell، DisplayAlerts.

    python3 -W ignore tools/make_vba.py vba/QC_WritePath.bas --workbook /tmp/QC-F-14-v2.xlsx

خروجی: صفرِ خطا = قابلِ Import. در غیر این صورت exit code = 1 و فهرستِ نقض‌ها.
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sys
import zipfile

import openpyxl

# VBA هر دو شکل را دارد: PutText wsD, r, "TOKEN", v   و   QcCol(wsD, "TOKEN")
TOKEN_RE = re.compile(r'(?:PutText|PutNum|QcText|QcCol|ColIdx)\s*\(?[^"\n]*?"([A-Z][A-Z0-9_]{2,})"')
KEY_RE = re.compile(r'(?:Msg|Label)\s*\(\s*"([A-Z][A-Z0-9_]{2,})"')
NAME_RE = re.compile(r'(?:GridRange|GridVal|GridAt)\s*\(\s*"?([A-Za-z_][A-Za-z0-9_]*)"?\s*[,)]')
CONST_NAME_RE = re.compile(r'Private Const (GRID_\w+|CFG_\w+) As String = "([^"]+)"')
TBL_RE = re.compile(r'ListObjects\((?:"([A-Za-z0-9_]+)"|([A-Z_]+))\)')

BLOCK_OPEN = re.compile(r'^\s*(?:Public |Private )?(Sub|Function|Property)\s+(\w+)', re.I)
FOR_RE = re.compile(r'^\s*For\b', re.I)
NEXT_RE = re.compile(r'^\s*Next\b', re.I)
DO_RE = re.compile(r'^\s*Do\b', re.I)
LOOP_RE = re.compile(r'^\s*Loop\b', re.I)
WITH_RE = re.compile(r'^\s*With\b', re.I)
ENDWITH_RE = re.compile(r'^\s*End With\b', re.I)
SEL_RE = re.compile(r'^\s*Select Case\b', re.I)
ENDSEL_RE = re.compile(r'^\s*End Select\b', re.I)
IF_BLOCK_RE = re.compile(r'^\s*(?:If|ElseIf|For Each)\b.*\bThen\s*(?:_.*)?$', re.I)


def lines_of(src: str):
    return src.replace('\r\n', '\n').split('\n')


def join_cont(src: str) -> str:
    """سطرهایِ ادامه‌دار (« _» در انتهای سطر) را به سطرِ منطقیِ خودشان می‌چسباند —
    بدونِ این کار، شمارشِ If/Do و برچسب‌ها روی خط‌هایِ فیزیکی معنی ندارد."""
    out, buf = [], ''
    for ln in lines_of(src):
        if buf:
            ln = buf + ' ' + ln.strip()
            buf = ''
        if re.search(r'\s_\s*$', ln):
            buf = re.sub(r'\s+_\s*$', '', ln)
            continue
        out.append(ln)
    if buf:
        out.append(buf)
    return '\n'.join(out)


def strip_comments(src: str) -> str:
    """حذفِ کامنت‌ها (برای lint رفتاری) — با احترام به رشته‌هایِ «"…»»."""
    out = []
    for ln in lines_of(src):
        res, q, i = [], False, 0
        while i < len(ln):
            ch = ln[i]
            if ch == '"':
                q = not q
            elif ch == "'" and not q:
                break
            res.append(ch)
            i += 1
        out.append(''.join(res))
    return '\n'.join(out)


def check(bad: list, name: str, cond, detail: str = '') -> None:
    print(('  ✓ ' if cond else '  ✗ ') + name + (f' — {detail}' if detail else ''))
    if not cond:
        bad.append(name)


def lint(src: str, bad: list) -> None:
    print('— قراردادِ منبع (ASCII / CRLF / تعادلِ بلوک‌ها)')
    non_ascii = sorted({c for c in src if ord(c) > 127})
    check(bad, 'T-V1 منبع خالصاً ASCII است (بدونِ فارسیِ درونِ کد)', not non_ascii,
          f'{len(non_ascii)} نویسهٔ غیرASCII: ' + ' '.join(f'U+{ord(c):04X}' for c in non_ascii[:6]))
    check(bad, 'T-V1 خط‌پایانِ CRLF (VBE با این انتظار می‌خواند)',
          '\r\n' in src and src.count('\n') == src.count('\r\n'))
    code = strip_comments(join_cont(src))
    for tag, cre, ere in (('Sub/End Sub', r'^\s*(?:Public |Private )?Sub\s+\w+', r'^\s*End Sub\b'),
                          ('Function/End Function', r'^\s*(?:Public |Private )?Function\s+\w+',
                           r'^\s*End Function\b'),
                          ('For/Next', r'^\s*For\s', r'^\s*Next\b'),
                          ('Do/Loop', r'^\s*Do\b', r'^\s*Loop\b'),
                          ('With/End With', r'^\s*With\s', r'^\s*End With\b'),
                          ('Select/End Select', r'^\s*Select Case\b', r'^\s*End Select\b')):
        o = sum(1 for ln in lines_of(code) if re.match(cre, ln, re.I))
        c = sum(1 for ln in lines_of(code) if re.match(ere, ln, re.I))
        check(bad, f'T-V1 تعادلِ {tag}', o == c, f'باز={o} بسته={c}')
    # ElseIf یک «Ifِ جدید» نیست؛ سطرِ یک‌خطی (If x Then y) هم End If لازم ندارد
    ifs = sum(1 for ln in lines_of(code)
              if re.match(r'^\s*If\b.*\bThen\s*$', ln, re.I))
    eifs = sum(1 for ln in lines_of(code) if re.match(r'^\s*End If\b', ln, re.I))
    check(bad, 'T-V1 تعادلِ If بلوکی/End If', ifs == eifs, f'بلوک={ifs} EndIf={eifs}')
    # continuation lines: "If cond _" - not used; keep it simple
    # پرش‌ها: On Error GoTo <برچسب> باید برچسبش در همان رویه باشد
    ok_lbl, det = True, []
    for fn in re.finditer(r'(?:Public|Private) Sub (\w+)[\s\S]*?^End Sub', code, re.M | re.I):
        body = fn.group(0)
        labels = {m.lower() for m in re.findall(r'^\s*(\w+):\s*$', body, re.M)}
        # Application.GoTo یک متدِ ناوبری است، نه پرشِ برچسبی
        for tgt in re.findall(r'(?<!Application\.)\bGoTo\s+(?!0\b)(\w+)', body, re.I):
            if tgt.lower() not in labels:
                ok_lbl, _ = False, det.append(f'{fn.group(1)}→{tgt}')
    check(bad, 'T-V1 هر GoTo به برچسبِ همان رویه می‌رود', ok_lbl, ', '.join(det))
    # رویهٔ خصوصیِ بی‌استفاده = کدِ مرده در فایلِ تحویلی
    # Public = نقطهٔ ورودِ دکمه/میان‌بر (از کد صدا زده نمی‌شود)؛ Private باید صدا زده شود
    priv = re.findall(r'^Private (?:Sub|Function) (\w+)', code, re.M | re.I)
    orphans = [d for d in priv
               if len(re.findall(rf'\b{re.escape(d)}\b', code, re.I)) < 2]
    check(bad, 'T-V1 هیچ رویهٔ خصوصیِ بی‌استفاده‌ای نمانده', not orphans, str(orphans))
    pub = re.findall(r'^Public (?:Sub) (\w+)', code, re.M | re.I)
    check(bad, 'T-V1 نقاطِ ورودِ عمومی دقیقاً همینی است که README وعده می‌دهد',
          sorted(pub) == sorted(['QC_Append', 'QC_Import_Staging', 'QC_VoidRow',
                                 'QC_ClearForm', 'QC_About']), str(sorted(pub)))


def policy(code: str, bad: list) -> None:
    print('— قراردادِ مسیرِ نوشتن (چیزی که VBA *نباید* بکند)')
    banned = [
        (r'\\.Rows\\([^)]*\\)\\.Delete', 'حذفِ ردیف'),
        (r'\.EntireRow\.Delete|\.EntireColumn\.Delete', 'Deleteِ سطر/ستون'),
        (r'\.Unprotect\b', 'Unprotect (قفل را نمی‌شکند)'),
        (r'\.SaveAs\b|ActiveWorkbook\.Save\b', 'ذخیرهٔ خودکار/SaveAs'),
        (r'\bShell\b|\bKill\b|\bOpen\s+"', 'دسترسیِ فایل/سیستمی'),
        (r'DisplayAlerts', 'خاموش‌کردنِ هشدارهایِ Excel'),
        (r'\.Visible\s*=\s*False|xlSheetHidden|xlSheetVeryHidden', 'پنهان/آشکارکردنِ شیت'),
        (r'\.Names\([^)]*\)\.Delete|ListObjects\([^)]*\)\.Delete', 'حذفِ نام/جدول'),
        (r'Application\.Run|Evaluate\(', 'اجرایِ ماکرو/فرمولِ دلخواه'),
        (r'\.Formula\s*=', 'نوشتنِ فرمول از VBA (فرمول‌ها محاسبه می‌کنند، نه کد)'),
        (r'Worksheets\([^\)]*\)\.Range\([^)]*\)\.Value\s*=\s*"', 'literalِ فارسیِ مستقیمِ متنی'),
    ]
    for pat, why in banned:
        hits = re.findall(pat, code)
        check(bad, f'T-V2 ممنوع: {why}', not hits, f'{len(hits)} مورد')
    # تنها .ClearContents مجاز: STAGING و FORM (نه Data)
    for i, ln in enumerate(lines_of(code), 1):
        if '.ClearContents' in ln:
            ok = re.search(r'wsS\.|wsF\.|STAGING|FORM', ln)
            check(bad, f'T-V2 ClearContents فقط روی فرم/STAGING (سطر {i})', bool(ok), ln.strip()[:60])
    check(bad, 'T-V2 هیچ رشتهٔ فارسیِ منطقی در کد نیست (همه از فایل)',
          not re.search(r'[\u0600-\u06FF]', code))


def contract(src: str, xlsx: str, bad: list) -> None:
    print('— قراردادِ ماژول ↔ فایل (نام‌هایی که کد به آن‌ها تکیه می‌کند)')
    wb = openpyxl.load_workbook(xlsx)
    # 1) ستون‌های tblQC
    wsD = wb['Data داده']
    toks = {str(c.value).split('\n')[-1].strip() for c in wsD[1] if c.value}
    used = set(TOKEN_RE.findall(src))
    check(bad, 'T-V3 هر توکنِ ستونِ استفاده‌شده در سرستونِ tblQC هست', used <= toks,
          f'ناشناخته: {sorted(used - toks)}  (از {len(used)})')
    # «نیمه‌پرشدنِ سطر» بیماریِ فایلِ قبلی بود: مسیرِ نوشتن باید هر ۲۶ ستون را بنویسد
    written = set(re.findall(r'^\s*Put(?:Text|Num) wsD, r, "([A-Z][A-Z0-9_]+)"', src, re.M))
    check(bad, 'T-V3 WriteRow هر ستونِ tblQC را پر می‌کند', toks <= written,
          f'جاافتاده: {sorted(toks - written)} (نوشته‌شده={len(written & toks)}/{len(toks)})')
    tabs = set(TBL_RE.findall(src))
    tab_names = {t for t in [x for x in re.findall(r'ListObjects\("([^"]+)"\)', src)]}
    have_tabs = set(wsD.tables)
    for nm in wb.sheetnames:
        have_tabs |= set(wb[nm].tables)
    # نام‌هایِ ثابتِ جدول در کد (با جایگزینیِ Const)
    consts = dict(re.findall(r'Const (TBL_\w+|GRID_\w+|CFG_\w+) As String = "([^"]+)"', src))
    want = {consts.get(m, m) for m in tab_names}
    check(bad, 'T-V3 جدول‌هایِ مرجع در فایل هستند', want <= have_tabs,
          f'ناشناخته: {sorted(want - have_tabs)}')
    # 2) نام‌هایِ تعریف‌شده
    dn = set(wb.defined_names)
    raw_refs = set(re.findall(r'(?:GridRange|GridVal|GridAt)\s*\(\s*"([A-Z][A-Z0-9_]*)"', src))
    raw_refs |= {consts[m] for m in re.findall(
        r'(?:GridRange|GridVal|GridAt)\s*\(\s*(GRID_[A-Z]+|CFG_[A-Z]+)\b', src) if m in consts}
    missing = sorted(g for g in raw_refs if g not in dn)
    check(bad, 'T-V3 گریدها/بلوک‌ها به‌صورتِ نامِ تعریف‌شده پیدا می‌شوند', not missing, f'نیست: {missing}')
    # 3) کلیدهایِ Msg()/Label() در Settings!QC_MSG
    wsS = wb['Settings تنظیمات']
    rng = wb.defined_names.get('QC_MSG')
    keys = set()
    if rng is not None:
        m = re.search(r'\$B\$(\d+):\$C\$(\d+)', str(rng.attr_text))
        for r in range(int(m.group(1)), int(m.group(2)) + 1):
            keys.add(str(wsS.cell(row=r, column=2).value))
    need = set(KEY_RE.findall(src))
    check(bad, 'T-V3 هر کلیدِ پیام در QC_MSG تعریف شده', need <= keys,
          f'ناقص: {sorted(need - keys)}')
    # 4) کلیدهایِ QC_SHEETS
    keys2 = set()
    r2 = wb.defined_names.get('QC_SHEETS')
    if r2 is not None:
        m = re.search(r'\$B\$(\d+):\$C\$(\d+)', str(r2.attr_text))
        for r in range(int(m.group(1)), int(m.group(2)) + 1):
            keys2.add(str(wsS.cell(row=r, column=2).value))
    used_sheets = set(re.findall(r'SheetAt\("([A-Z]+)"\)', src))
    check(bad, 'T-V3 هر کلیدِ شیت در QC_SHEETS تعریف شده', used_sheets <= keys2,
          f'ناقص: {sorted(used_sheets - keys2)}')
    # 5) اندیس‌هایِ گرید = ترتیبِ واقعیِ فایل
    def grid_vals(nm):
        out = []
        rg = wb.defined_names.get(nm)
        if rg is None:
            return out
        m = re.search(r"'([^']+)'!\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+)", str(rg.attr_text))
        ws = wb[m.group(1)]
        for r in range(int(m.group(3)), int(m.group(4)) + 1):
            v = ws[f'{m.group(2)}{r}'].value
            if v in (None, ''):
                break
            out.append(str(v))
        return out
    st = grid_vals(consts.get('GRID_STAT', 'CODE_ROW_STATUS'))
    want_st = {'ST_OK': (1, 'تأییدشده'), 'ST_VOID': (3, 'ابطال‌شده'), 'ST_DUP': (4, 'مشکوک-تکراری')}
    ok, det = True, []
    for k, (i, v) in want_st.items():
        ci = int(re.search(rf'Const {k} As Long = (\d+)', src).group(1))
        got = st[ci - 1] if 0 < ci <= len(st) else None
        if got != v:
            ok = False
            det.append(f'{k}={ci}→{got!r}')
    check(bad, 'T-V3 اندیس‌هایِ CODE_ROW_STATUS با فایل می‌خواند', ok and len(st) == 4,
          f'گرید={st}' + (' | ' + ', '.join(det) if det else ''))
    mc = grid_vals(consts.get('GRID_MOLD', 'CODE_MOLD_CHECK'))
    for k in ('MC_MATCH', 'MC_MISS', 'MC_DIRECTIONAL', 'MC_DIGIT'):
        ci = int(re.search(rf'Const {k} As Long = (\d+)', src).group(1))
        check(bad, f'T-V3 اندیسِ {k} در بازهٔ CODE_MOLD_CHECK است', 1 <= ci <= len(mc), f'{ci}/{len(mc)}')
    # 6) آدرس‌هایِ فرم و STAGING که کد کورکورانه استفاده می‌کند
    wsF = wb['FORM فرم ثبت']
    cells = set(re.findall(r'(?:CellTxt|RawText)\(wsF, "([A-Z]+\d+)"\)', src)) | \
        set(re.findall(r'wsF\.Range\("([A-Z]+)\d?', src))
    inputs = {'C6', 'C8', 'C10', 'C12', 'F6', 'F8', 'F10', 'F12', 'C14', 'C28'}
    dv_addr = set()
    for d in wsF.data_validations.dataValidation:
        for part in re.split(r'[ ,:]+', str(d.sqref or '')):
            if re.fullmatch(r'[A-Z]+\d+', part):
                dv_addr.add(part)
    check(bad, 'T-V3 هر سلولِ ورودی که کد می‌خواند روی فرم DV دارد',
          {c for c in cells if c in inputs} == (cells & inputs) and
          all(c in dv_addr or c == 'C28' for c in cells & inputs),
          f'خوانده‌شده={sorted(cells & inputs)} بدونِ DV={sorted(c for c in cells & inputs if c not in dv_addr)}')
    wsG = wb['STAGING ورود اضطراری']
    flag = wsG['K8'].value
    check(bad, 'T-V3 ستونِ FLAG در STAGING همان‌جاست که کد می‌خواند',
          isinstance(flag, str) and flag.startswith('='), str(flag)[:40])
    # 7) LOG_AUDIT: سطرِ ۳ باز است و ۸ ستونِ سرستون دارد
    wsL = wb['LOG_AUDIT']
    hdr = [c.value for c in wsL[2]][:8]
    check(bad, 'T-V3 LOG_AUDIT سرستونِ ۸تایی در سطرِ ۲ دارد', len([h for h in hdr if h]) == 8, str(hdr))
    check(bad, 'T-V3 LOG_AUDIT از سطرِ ۳ پر می‌شود (ثابتِ LOG_FIRST=3)',
          all(wsL.cell(row=3, column=j).value in (None, '') for j in range(1, 9)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('bas', nargs='?', default='vba/QC_WritePath.bas')
    ap.add_argument('--workbook', default='/tmp/QC-F-14-v2.xlsx')
    a = ap.parse_args()
    raw = io.open(a.bas, 'rb').read()
    src = raw.decode('utf-8')
    print(f'make_vba: {a.bas} ({len(raw)} B, {len(lines_of(src))} سطر) ↔ {a.workbook}')
    bad = []
    lint(src, bad)
    policy(strip_comments(src), bad)
    if os.path.exists(a.workbook):
        contract(src, a.workbook, bad)
    else:
        print('  (فایل را پیدا نکردم: قراردادِ ماژول↔فایل سنجیده نشد)')
    print('VBA MODULE CHECKS: ' + ('PASS' if not bad else f'FAIL ({len(bad)})'))
    for b in bad:
        print('   ✗', b)
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
