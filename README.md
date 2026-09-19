# QC-F-14 — سامانه QC مبتنی بر Excel (نسخه ۳ آماده اجرا)

این مخزن شامل **فایل‌های آماده اجرا** و **ابزارهای بازتولید و آزمون** است. هیچ سلولی دستی ویرایش نشده؛ هر فایل خروجی `tools/` است.

## فایل‌های تحویل

| مسیر | چیست | وضعیت |
|---|---|---|
| `QC-F-14-v3.xlsx` | **نسخه ۳ آماده اجرا** — 12 شیت، `tblQC` با **32 ستون** و **14,557 رکورد** (RECORD_ID QC-00000001..), 15 گزارش با Barcode History کامل, Dashboard با Record vs Unique, DQ, Help. ماکرو-آزاد | **جدید — مطابق دستور جدید** |
| `QC-F-14-v2.xlsx` | نسخه ۲ (26 ستون, EVENT_ID) — 2.0 MB, 14,557 رکورد | قبلی، حفظ شد |
| `FORM-QC-F-14.xlsm` | فایل مرجع اصلی (684KB) — فقط خواندنی | مرجع |
| `vba/QC_WritePath_v3.bas` | VBA مسیر نوشتن v3 (24KB, ASCII): QC_Append (11 فیلد), QC_Import_Staging, QC_VoidRow (با دلیل), QC_ClearForm, QC_About | جدید |
| `vba/QC_WritePath.bas` | VBA v2 (23KB) | قبلی |
| `docs/` | Data Dictionary, Master Data Guide, VBA Doc, Migration Report, Test Report (30 تست), User Guide, Known Limitations — همه برای v3 | جدید |
| `analysis/05-phase1-re-audit-2026-09-19.md` | حسابرسی مجدد فاز ۱ با اعداد واقعی از فایل | جدید |
| `analysis/00..04-*.md` | گزارش‌های فاز ۱-۲ قدیمی | قبلی |
| `tools/` | `build_payload_v3.py`, `gen_workbook_v3.py`, `build_payload.py`, `gen_workbook.py`, `xlcal.py`, `xlmigrate.py`, `msovba.py`, تست‌ها | |

## معماری داده v3 (مطابق دستور جدید)

- **Record_ID** = `QC-00000001` یکتا، خودکار MAX+1، غیرقابل تغییر، ثابت
- **Barcode** = شناسه قطعه، نه رکورد — تکرار خطا نیست (463 گروه تکراری → REINSPECTION)
- **Record_Type**: INSPECTION 13,986 / REINSPECTION 473 / RETURN 98 (از ایستگاه برگشت + تاریخ بارکد)
- **Inspection_Result**: OK/NOK (مهاجرت: NOK 14,557 چون همه دارای عیب)
- **Disposition**: ACCEPT/REWORK (12,814)/SCRAP (1,645)/RETURN_TO_PROCESS (98)/HOLD — 5 مقدار
- **Current_Status**: OPEN (478)/UNDER_REWORK/WAITING_REINSPECTION/CLOSED (14,079)/VOID — 5 مقدار
- **Void**: به‌جای Delete → VOID_FLAG TRUE + VOID_REASON + CURRENT_STATUS=VOID
- **Master Data**: 11 جدول CODE|NAME|ACTIVE (Shifts, Stations, Inspectors, Parts, Molds, Defects, RecordTypes, Results, Dispositions, Statuses) + MAP_PART_MOLD + RULE_STATION + RULE_LOGIC (7 قانون ناسازگاری)
- **Barcode History**: Reports R11 — با وارد کردن بارکد، تمام رخدادها به ترتیب زمانی (Record_ID, Date, Shift, Station, Defect, Result, Disposition, Status, Inspector)

## ساختن دوباره و آزمون

```bash
# v3
python3 -W ignore tools/build_payload_v3.py FORM-QC-F-14.xlsm -o analysis/payload_v3.json
python3 -W ignore tools/gen_workbook_v3.py --payload analysis/payload_v3.json -o QC-F-14-v3.xlsx

# v2 (قدیمی)
python3 -W ignore tools/build_payload.py FORM-QC-F-14.xlsm -o analysis/payload.json
python3 -W ignore tools/gen_workbook.py --payload analysis/payload.json -o QC-F-14-v2.xlsx

# تست‌ها
python3 -W ignore tools/test_jalali.py
python3 -W ignore tools/test_workbook.py QC-F-14-v2.xlsx
```

**آخرین اجرا v3:**
- records 14,557 → 14,557, distinct barcode 14,079, leading_zero 9,283, len 19=14,557, dup_groups 463, dup_extra 478, max_repeat 4, dates 2026-04-20→2026-09-19 123 روز
- Record_Type: INSPECTION 13,986 / REINSPECTION 473 / RETURN 98
- Disposition: REWORK 12,814 / SCRAP 1,645 / RETURN_TO_PROCESS 98
- MIGRATION FAILS: NONE

## نصب ماکرو v3 (۲ دقیقه)

1. QC-F-14-v3.xlsx را باز کنید
2. Alt+F11 → File → Import File → `vba/QC_WritePath_v3.bas`
3. Save As `.xlsm`
4. FORM: Shape → Assign Macro → QC_Append
5. STAGING: دکمه → QC_Import_Staging, Data: دکمه → QC_VoidRow
6. Trusted Location

## سه هشدار صریح (در فایل هم نوشته)

1. **چندکاربره همزمان روی یک فایل پشتیبانی نمی‌شود** — الگو: یک فایل به‌ازای ایستگاه + تجمیع مرکزی Power Query
2. Excel Database/سیستم امنیتی نیست — LOG_AUDIT best-effort
3. نرخ ضایعات بر تولید و تأخیر ثبت به دلیل نبود داده ساخته نشدند (❌ در Settings)

## محدودیت‌های صادقانه

- بدون Excel در sandbox → رفتار ماکرو، رندر نمودار، RTL در Excel Desktop باید با چک‌لیست 10 موردی (analysis/04) توسط کاربر تأیید شود
- VBA 24KB > ظرفیت بازنویسی Module1 قدیمی (4,142 بایت) → تزریق باینری مستقیم انجام نشد، ماژول import است
- GRID 40 ردیف ثابت — اگر Master >40 شد، range(40) را در build_helper بزرگ کنید

## مستندات

- `docs/DataDictionary_v3.md` — 32 ستون tblQC
- `docs/MasterDataGuide_v3.md` — 11 جدول + MAP + RULE
- `docs/VBADoc_v3.md` — 5 نقطه ورود
- `docs/MigrationReport_v3.md` — تطبیق قبل/بعد
- `docs/TestReport_v3.md` — 30 تست با Evidence
- `docs/UserGuide_v3.md` — راهنمای کاربر
- `docs/KnownLimitations_v3.md` — مرزهای صریح
- `analysis/05-phase1-re-audit-2026-09-19.md` — حسابرسی مجدد فاز ۱
