"""
MS-OVBA codece — خواندن/جایگزینی متن منبع ماژول‌ها در `xl/vbaProject.bin`

چرا این ابزار لازم است؟
  openpyxl (و هیچ ابزارِ متنیِ دیگری) `vbaProject.bin` را نمی‌سازد؛ بدون آن،
  فایل `.xlsm` ماکرو ندارد. استراتژی تحویل: «پروژهٔ VBA فایل مرجع را بازیافت
  کن و فقط متن منبع ماژول‌ها را جایگزین کن». پس تنها لایهٔ فشرده‌سازی پیاده
  شده و به ساختار OLE، جریان `dir`، `PROJECT`، `PROJECTwm`، `_VBA_PROJECT` و
  PerformanceCache دست نمی‌زنیم؛ طول هر جریان هم با پدِ صفر حفظ می‌شود تا هیچ
  زنجیرهٔ سکتوری عوض نشود.

الگوریتم (MS-OVBA 2.4.1، همان‌که Excel می‌نویسد و oletools می‌خواند)
  • جریانِ ماژول = [پیشوند/PerformanceCache] + کانتینرِ منبع؛ کانتینر با بایت
    `0x01` شروع می‌شود و Excel آن را در «آخرین» موقعیتِ معتبر جریان می‌گذارد.
  • هر chunk: هدر ۲ بایتی
        CompressedChunkSize = (hdr & 0xFFF) + 3
        CompressedChunkFlag = (hdr >> 15) & 1        0 = RawChunk
        CompressedChunkSignature = (hdr >> 12) & 0b011 == 0b011
  • داخل chunk (per oletools `decompress_stream`):
        pos_in_out = len(decompressed_buffer) - len(at_chunk_start)
        BitCount   = max(4, ceil(log2(max(pos_in_out, 2))))
        LengthMask = 0xFFFF >> BitCount ; OffsetMask = ~LengthMask
        enc  = (token & LengthMask) + 3
        off  = (token >> (16 - BitCount)) + 1
    و چون `pos_in_out` در decoder، **شاملِ FlagByteها** است، encoder هم باید
    موقعیتِ هر توکن را «آفستِ آن توکن در جریانِ فشرده نسبت به ابتدای chunk»
    بگیرد (نه شمارشِ بایت‌هایِ محتوایی). همین یک نکته، سه بار این ابزار را
    می‌شکست؛ تستِ تطبیقی (`tools/test_msovba.py`، T5) آن را می‌گیرد.
  • قاعدهٔ 2.4.1.3.21: اگر CopyOffset == 1 باشد، CopySize باید ۱ (۳ بایت) باشد.

تست‌ها (اجرا از ریشهٔ ریپازیتوری):
    python3 -W ignore tools/test_msovba.py FORM-QC-F-14.xlsm
هر تستی که شکست بخورد، build فایل جدید را متوقف می‌کند؛ ما فایلِ «تقریباً درست»
تحویل نمی‌دهیم.
"""
import math
import struct

from oletools.olevba import decompress_stream

DECOMP_BLOCK = 4096
CHUNK_SIG = 0b011


def bitcount_for(pos_in_out: int) -> int:
    """pos_in_out = بایت‌هایِ موجودِ انتهای DecompressedBuffer در همین chunk
    (FlagByteها هم شمرده می‌شوند) — همان چیزی که Excel/oletools به‌کار می‌برد."""
    return max(4, int(math.ceil(math.log(max(pos_in_out, 2), 2))))


def lmask(bit_count: int) -> int:
    return 0xFFFF >> bit_count


def omask(bit_count: int) -> int:
    return 0xFFFF ^ lmask(bit_count)


def max_copy_len(bit_count: int) -> int:
    return lmask(bit_count) + 2      # enc = len-2  ≤ LengthMask


def max_copy_off(bit_count: int) -> int:
    lm = lmask(bit_count)
    return ((0xFFFF ^ lm) >> (16 - bit_count)) + 1 if bit_count < 16 else 4095


# ─────────────────────────────── decode ───────────────────────────────
def decompress(blob: bytes, end: int | None = None, read_end: int | None = None) -> bytes:
    """رمزگشایی یک CompressedContainer.

    end = پایانِ دقیقِ کانتینر. اگر داده نشود، تا انتهای blob خوانده می‌شود؛
    داخلِ جریان ماژول، کانتینر با صفرها پد شده، پس فراخوان‌ها باید endِ درست
    را بدهند (وگرنه padding به‌عنوان chunk جدید تفسیر می‌شود).
    """
    if blob[:1] != b'\x01':
        raise ValueError('bad container signature byte')
    out = bytearray()
    i, n = 1, len(blob) if end is None else end
    lim = len(blob) if read_end is None else read_end      # حدِ خواندنِ فیزیکی
    while i + 2 <= n:
        (hdr,) = struct.unpack_from('<H', blob, i)
        size = (hdr & 0x0FFF) + 3
        flag = (hdr >> 15) & 1
        dstart, dend = i + 2, min(lim, i + size)
        if flag == 0:                                   # RawChunk
            out += blob[dstart:dstart + DECOMP_BLOCK]
            i = dstart + DECOMP_BLOCK
            continue
        if (hdr >> 12) & 0b111 != CHUNK_SIG:
            break
        chunk_start_out = len(out)
        pos = dstart
        while pos < dend:
            if len(out) - chunk_start_out >= DECOMP_BLOCK:
                break
            flags = blob[pos]
            pos += 1
            for bit in range(8):
                if pos >= dend:
                    break
                if not (flags >> bit) & 1:
                    out.append(blob[pos])
                    pos += 1
                else:
                    token = blob[pos] | (blob[pos + 1] << 8)
                    pos += 2
                    # دقیقاً مثل oletools: BitCount از (DecompressedCurrent -
                    # DecompressedChunkStart) لحظهٔ خواندنِ CopyToken.
                    bc = bitcount_for(len(out) - chunk_start_out)
                    lm = lmask(bc)
                    ln = (token & lm) + 3
                    off = (token >> (16 - bc)) + 1
                    # oletools فقط چک می‌کند offset از len(decompressed_container)
                    # بزرگ‌تر نباشد؛ یعنی ارجاعِ CopyToken به chunkهای قبلی مجاز است
                    # (در جریان dirِ فایل مرجع هم دیده می‌شود).
                    if off < 1 or off > len(out):
                        raise ValueError(f'bad copy offset {off} at {len(out)}')
                    src = len(out) - off
                    for k in range(ln):
                        if len(out) - chunk_start_out >= DECOMP_BLOCK:
                            break
                        out.append(out[src + k])
        i += size
    return bytes(out)


# ─────────────────────────────── encode ───────────────────────────────
def _encode_chunk(seg: bytes) -> bytes:
    """seg ≤ ۴۰۹۶ بایتِ باز → محتوایِ یک chunk فشرده (بدون هدر).

    دقیقاً معادلِ oletools/Excel:
        BitCount از (DecompressedCurrent - DecompressedChunkStart) در **لحظهٔ
        خواندنِ هر CopyToken** حساب می‌شود؛ پس داخلِ یک گروهِ ۸تایی هم عوض
        می‌شود (بایت‌های literalِ همان گروه هم شمرده می‌شوند، ولی بایت‌های
        فشردهٔ خودِ توکن نه — آن‌ها درِ DecompressedBuffer نمی‌شوند).
        LengthMask = 0xFFFF>>BC ، CopySize = (tok & LM)+3 ،
        CopyOffset = (tok & ~LM)>>(16-BC) + 1  ⇒  tok = ((off-1)<<(16-BC)) | (len-3)
    قاعدهٔ 2.4.1.3.21: CopyOffset==1 ⇒ CopySize باید ۱ (۳ بایت) باشد.
    """
    res = bytearray()
    dlen = 0                       # بایت‌هایِ بازِ تولیدشده در این chunk
    i, n = 0, len(seg)
    while i < n:
        flag = len(res)
        res.append(0)
        for bit in range(8):
            if i >= n or dlen >= DECOMP_BLOCK:
                break
            # مبنایِ BitCount همان شمارشِ درونِ chunk است (oletools:
            # DecompressedCurrent - DecompressedChunkStart). چون هر chunk را کامل
            # ۴۰۹۶ بایتی می‌نویسیم، این با شمارشِ سراسری Excel یکی است.
            bc = bitcount_for(dlen)
            lm = lmask(bc)
            cap_len = min(lm + 2, DECOMP_BLOCK - dlen, n - i)
            cap_off = min(dlen, max_copy_off(bc))
            best_len = best_off = 0
            if cap_len >= 3 and cap_off >= 1:
                first = seg[i]
                for off in range(1, cap_off + 1):
                    j = i - off
                    if j < 0 or seg[j] != first:
                        continue
                    ln = 1
                    while ln < cap_len and seg[j + ln] == seg[i + ln]:
                        ln += 1
                    if ln >= 3 and not (off == 1 and ln > 3) and ln > best_len:
                        best_len, best_off = ln, off
                        if ln >= cap_len:
                            break
            if best_len >= 3:
                enc = best_len - 3
                assert enc <= lm, f'enc {enc} > lmask {lm} (dlen={dlen}, bc={bc})'
                assert (best_off - 1) <= (0xFFFF ^ lm) >> (16 - bc), 'offset beyond OffsetMask'
                res[flag] |= 1 << bit
                token = ((best_off - 1) << (16 - bc)) | enc
                res.append(token & 0xFF)
                res.append(token >> 8)
                dlen += best_len
                i += best_len
            else:
                res.append(seg[i])
                dlen += 1
                i += 1
    assert dlen == len(seg), f'chunk produced {dlen} of {len(seg)} decompressed bytes'
    return bytes(res)


def compress(data: bytes) -> bytes:
    """chunkهای فشرده (بدون بایتِ امضا).

    هر chunk با شمارشِ *درون‌chunkی* (مثل oletools) رمزگذاری می‌شود، پس chunkِ
    پایانیِ کوتاه هم مشکلی ندارد و لازم نیست منبع به ۴۰۹۶ پد شود.
    """
    out = bytearray()
    for start in range(0, max(len(data), 1), DECOMP_BLOCK):
        seg = data[start:start + DECOMP_BLOCK]
        enc = _encode_chunk(seg)
        # CompressedChunkSize = ۲ بایت هدر + دادهٔ فشرده.
        # عمداً از RawChunk استفاده نمی‌کنیم: آن همیشه ۴۰۹۶ بایتِ کامل است و
        # در segmentِ آخر، بایت‌های اضافی به خروجیِ باز اضافه می‌کرد (خطای
        # «دیکدِ درست ولی طولِ زیاد»). chunkِ تمام‑literal با همان طولِ واقعی،
        # خروجی را بایت‌به‌بایت دقیق نگه می‌دارد و با Excel/oletools سازگار است.
        assert len(enc) + 2 <= 0xFFB, (
            f'chunk too large for the 12-bit size field ({len(enc)}B) — '
            'منبع خیلی بزرگ یا غیرقابل‌فشرده است؛ کد را به چند ماژول تقسیم کنید')
        out += struct.pack('<H', (len(enc) + 2 - 3) | (CHUNK_SIG << 12) | (1 << 15)) + enc
    return bytes(out)


def pack_source(source: bytes) -> bytes:
    """جریانِ کاملِ قابل‌جایگذاری = 0x01 + chunkها، با verifyِ سخت."""
    blob = b'\x01' + compress(source)
    back = decompress(blob, len(blob))
    if back[:len(source)] != source:
        for k in range(min(len(back), len(source))):
            if back[k] != source[k]:
                raise ValueError(f'VBA encoder/decoder mismatch در بایت {k}: '
                                 f'{source[max(0,k-16):k+16]!r} ≠ {back[max(0,k-16):k+16]!r}')
        raise ValueError(f'VBA encoder/decoder mismatch: len {len(source)} → {len(back)}')
    if decompress_stream(bytearray(blob))[:len(source)] != source:
        raise ValueError('oletools cannot decode the encoder output')
    return blob


# ─────────────────── پیدا کردن و جایگزینی منبع در جریانِ ماژول ───────────────────
def _candidates(module_stream: bytes):
    n = len(module_stream)
    for cand in range(0, n - 3):
        if module_stream[cand] != 0x01:
            continue
        (hdr,) = struct.unpack_from('<H', module_stream, cand + 1)
        size = (hdr & 0x0FFF) + 3
        flag = (hdr >> 15) & 1
        if flag == 0 and size != 4098:
            continue
        if flag == 1 and (hdr >> 12) & 0b011 != CHUNK_SIG:
            continue
        if cand + 1 + size > n:
            continue
        yield cand


def _walk_container(module_stream: bytes, start: int, stop: int | None = None):
    """پایانِ کانتینر = انتهایِ آخرین chunkِ اعلام‌شده (padding صفر/0xFF شمرده نمی‌شود)."""
    """(پایانِ اعلام‌شدهٔ کانتینر, پایانِ فیزیکیِ آخرین chunk) — padding صفر را
    نمی‌شمارد، اما اجازه می‌دهد chunkهای آخر تا انتهای جریان خوانده شوند."""
    n = len(module_stream) if stop is None else min(stop, len(module_stream))
    i = start + 1
    end = i
    while i + 2 <= n:
        (hdr,) = struct.unpack_from('<H', module_stream, i)
        if (hdr >> 15) & 1 == 0 and (hdr & 0x0FFF) + 3 != 4098:
            break                       # هدرِ معتبر نیست → انتهای کانتینر
        size = (hdr & 0x0FFF) + 3
        if size < 3 or i + size > n:
            break
        end = i + size
        i = end
    return end, min(max(n, start + 1), len(module_stream))


def find_container_start(module_stream: bytes) -> int:
    """آخرین کاندیدی که منبعِ «Attribute …» می‌دهد (Excel منبع را در انتها می‌گذارد؛
    کاندیدهای زودتر می‌توانند از دلِ PerformanceCache تصادفی بیرون بزنند)."""
    best = -1
    n = len(module_stream)
    for cand in _candidates(module_stream):
        try:
            end, _ = _walk_container(module_stream, cand)
            src = decompress(module_stream[cand:end], end - cand, n)
        except Exception:
            continue
        if src.startswith(b'Attribute ') or src.startswith(b'option ') or src.strip() == b'':
            best = cand
    return best


def extract_module_source(module_stream: bytes) -> bytes:
    start = find_container_start(module_stream)
    if start < 0:
        raise ValueError('VBA source container not found in module stream')
    end, phys = _walk_container(module_stream, start)
    return decompress(module_stream[start:phys], end - start, phys - start)


FILLER = b'\xff'      # 0xFF هدرِ نامعتبر است → مرزِ کانتینر را برای خواننده حفظ می‌کند


def replace_module_source(module_stream: bytes, new_source: bytes,
                         pad: bytes = FILLER) -> bytes:
    """جایگزینی منبع با حفظِ طولِ کلیِ جریان (ساختارِ OLE دست‌نخورده می‌ماند)."""
    start = find_container_start(module_stream)
    if start < 0:
        raise ValueError('VBA source container not found in module stream')
    blob = pack_source(new_source)
    avail = len(module_stream) - start
    if len(blob) > avail:
        raise ValueError(f'منبع فشرده {len(blob)}B در فضای {avail}B جا نمی‌شود؛ '
                         f'کد را کوتاه کنید یا در VBE جایگذاری کنید.')
    out = module_stream[:start] + blob + pad * (avail - len(blob))
    if extract_module_source(out) != new_source:
        raise ValueError('verify failed: re-read of replaced module source differs')
    end, _ = _walk_container(out, start)
    if decompress_stream(bytearray(out[start:end]))[:len(new_source)] != new_source:
        raise ValueError('verify failed: oletools cannot read the replaced module stream')
    return out
