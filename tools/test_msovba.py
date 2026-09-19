"""
تست‌های ابزار MS-OVBA (vbaProject.bin) — شرطِ لازمِ build فایل جدید.

    python3 -W ignore tools/test_msovba.py [FORM-QC-F-14.xlsm]

T1  استخراج منبع هر ماژول = خروجی oletools
T2  جایگزینی منبع (حتی با متنِ یکسان) → round-trip سالم و حفظ طول جریان
T3  منابع مصنوعی (کوتاه/تکراری/بلند/تصادفی) → pack/unpack دقیق
T4  جریان dir با decoder ما == oletools
T5  منبعِ واقعی هر ماژول، بعد از فشرده‌سازیِ مجدد، «بایت‌به‌بایت» با oletools یکی باشد
    (این تست، فرمولِ موقعیتِ FlagByteها را می‌سنجد؛ بدون آن encoder «تقریباً درست» است)
T6  تزریق کدِ جدیدِ واقعی (ModQC ~ 6KB) در Module1 و خوانشِ مجدد با olevba
"""
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import msovba                                    # noqa: E402
import olefile                                   # noqa: E402
from oletools.olevba import VBA_Parser, decompress_stream   # noqa: E402

SRC = sys.argv[1] if len(sys.argv) > 1 else 'FORM-QC-F-14.xlsm'
if not os.path.exists(SRC):
    print('فایل مرجع پیدا نشد:', SRC)
    sys.exit(2)

binpath = '/tmp/_msovba_test.bin'
with zipfile.ZipFile(SRC) as z:
    open(binpath, 'wb').write(z.read('xl/vbaProject.bin'))
ole = olefile.OleFileIO(binpath)
ref = {n.strip().rsplit('.', 1)[0]: c for (f, sp, n, c) in VBA_Parser(SRC).extract_macros()}
MODULES = [e[1].split('/')[-1] for e in ole.listdir()
           if e[0] == 'VBA' and e[1].split('/')[-1] not in ('dir', '_VBA_PROJECT')
           and '__SRP' not in e[1]]
fails = []

streams = {}
for st in MODULES:
    streams[st] = ole.openstream('VBA/' + st).read()
SPACE = {st: len(raw) - msovba.find_container_start(raw) for st, raw in streams.items()}
print('SPACE (بایتِ قابل‌بازنویسیِ هر ماژول):',
      ', '.join(f'{k}={v}' for k, v in sorted(SPACE.items())))

# T1
for st, raw in streams.items():
    try:
        src = msovba.extract_module_source(raw).decode('cp1252', 'replace')
    except Exception as e:
        fails.append(f'T1 {st}: extract failed: {e}')
        continue
    base = st.replace('.bas', '').replace('.cls', '')
    if base in ref and src.strip() != ref[base].strip():
        fails.append(f'T1 {st}: != olevba ({len(src)} vs {len(ref[base])})')

# T2
for st, raw in streams.items():
    try:
        src = msovba.extract_module_source(raw)
        new = msovba.replace_module_source(raw, src)
        back = msovba.extract_module_source(new)
        if back != src:
            fails.append(f'T2 {st}: round-trip mismatch')
        if len(new) != len(raw):
            fails.append(f'T2 {st}: stream length changed')
    except Exception as e:
        fails.append(f'T2 {st}: {e}')

# T3
synthetic = {
    'short': b'Attribute VB_Name = "X"\r\nSub A()\r\nEnd Sub\r\n',
    'repeat': b'Attribute VB_Name = "X"\r\n' + b'    x = 1\r\n' * 400,
    'structured': b'Attribute VB_Name = "X"\r\n' +
                  (b'Option Explicit\r\nPublic Sub A()\r\n    Dim q As Long\r\nEnd Sub\r\n') * 300,
    'random': b'Attribute VB_Name = "X"\r\n' +
              bytes([(k * 7919 + 13) % 26 + 65 for k in range(9000)]),
    'two-byte-rows': b'Attribute VB_Name = "X"\r\n' + b'\r\n' * 3000,
}
for name, data in synthetic.items():
    try:
        blob = msovba.pack_source(data)
        if msovba.decompress(blob, len(blob)) != data:
            fails.append(f'T3 {name}: my round-trip failed')
        if decompress_stream(bytearray(blob)) != data:
            fails.append(f'T3 {name}: oletools cannot decode my output')
    except Exception as e:
        fails.append(f'T3 {name}: {e}')

# T4
try:
    d = ole.openstream('VBA/dir').read()
    if msovba.decompress(d) != decompress_stream(bytearray(d)):
        fails.append('T4 dir: decompress != oletools')
except Exception as e:
    fails.append(f'T4 dir: {e}')

# T5 — بایت‌به‌بایت با دادهٔ واقعی
for st, raw in streams.items():
    try:
        src = msovba.extract_module_source(raw)
        blob = msovba.pack_source(src)
        if decompress_stream(bytearray(blob)) != src:
            fails.append(f'T5 {st}: recompressed source is not byte-identical')
    except Exception as e:
        fails.append(f'T5 {st}: {e}')

# T6 — تزریق کدِ جدید + ظرفیتِ واقعیِ هر ماژول (چون جریان‌ها بازنویسی‌اند و
# stream جدید نمی‌سازیم، ظرفیت = فضایِ موجود بعد از PerformanceCache)
def inject(cap_bytes):
    line = "Public Function V{n}(ByVal s As String) As Boolean\r\n    V{n} = Len(s) > 0\r\nEnd Function\r\n"
    k, src = 0, b'Attribute VB_Name = "ModQC"\r\nOption Explicit\r\n'
    while len(src) < cap_bytes:
        k += 1
        src += line.format(n=k).encode()
    return src[:cap_bytes], k


def fits(st, cap):
    try:
        msovba.replace_module_source(streams[st], inject(cap)[0])
        return True
    except ValueError:
        return False


CAP = {}
for st in sorted(streams):
    lo, hi = 0, 16384
    while lo < hi:
        mid = (lo + hi + 1) // 2
        if fits(st, mid):
            lo = mid
        else:
            hi = mid - 1
    CAP[st] = lo
    print(f'  ظرفیتِ Module stream {st}: {SPACE[st]}B فشرده → ≈{lo} بایت منبع')


src, _ = inject(CAP['Module1'])
try:
    raw = streams['Module1']
    new_stream = msovba.replace_module_source(raw, src)
    if msovba.extract_module_source(new_stream) != src:
        fails.append('T6: re-read of injected source differs')
    st = msovba.find_container_start(new_stream)
    end, _ = msovba._walk_container(new_stream, st)
    if decompress_stream(bytearray(new_stream[st:end]))[:len(src)] != src:
        fails.append('T6: oletools cannot read the injected stream')
    if new_stream[:st] != raw[:st]:
        fails.append('T6: stream prefix (PerformanceCache) was modified')
    if len(new_stream) != len(raw):
        fails.append('T6: stream length changed')
except Exception as e:
    fails.append(f'T6: {e}')

print('MODULES:', ', '.join(sorted(MODULES)))
print('ظرفیتِ منبعِ Module1 (بازنویسی‌پذیر) ≈', CAP['Module1'], 'بایت')
if fails:
    print('FAIL')
    for f in fails:
        print('  ✗', f)
    sys.exit(1)
print('ALL TESTS PASS (T1–T6)')
