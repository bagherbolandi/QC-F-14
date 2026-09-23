# v3 — گزارش ساخت و تحویل (آماده اجرا)

**تاریخ:** 2026-09-19  
**فایل:** `QC-F-14-v3.xlsx` (2.3 MB) + `vba/QC_WritePath_v3.bas` (24,310 بایت)  
**مبنا:** `FORM-QC-F-14.xlsm` (700,350 بایت) + `QC-F-14-v2.xlsx` (2.0 MB)  
**ابزار:** `tools/build_payload_v3.py` → `analysis/payload_v3.json` (13.7 MB) → `tools/gen_workbook_v3.py`

## ۱) شناسنامه

| مورد | مقدار |
|---|---|
| نام فایل | QC-F-14-v3.xlsx |
| حجم | 2.3 MB |
| شیت‌ها | 12: Data داده, FORM فرم ثبت, STAGING, Master, Helper (h), Reports, Dashboard, DQ, Settings, Help, LOG_AUDIT (h), Verify (h) |
| Table | tblQC = A1:AF20000 (32 ستون), 15 جدول دیگر (11 DIM + MAP + RULE + RULE_LOGIC) |
| رکورد | 14,557 مهاجرت‌یافته, ظرفیت 20,000 |
| Defined Names | 29+ (QC_CAL_*, *_list, CODE_*, QC_SHEETS, QC_MSG) |
| فرمول | 1,628+ (صفر روی ردیف‌های Data) |
| VBA | macro-free + vba/QC_WritePath_v3.bas (ASCII, 5 entry points) |

## ۲) تغییرات نسبت به v2

| v2 | v3 | دلیل |
|---|---|---|
| EVENT_ID E-000001 | RECORD_ID QC-00000001 (8 رقم) | دستور جدید: Record ID با QC- |
| EVENT_TYPE اصلاحی/ضایعات/برگشتی | RECORD_TYPE INSPECTION/REINSPECTION/RETURN + DISPOSITION 5تایی + INSPECTION_RESULT OK/NOK + CURRENT_STATUS 5تایی | تفکیک مفاهیم QC (مرحله سوم دستور) |
| DISPOSITION_ID DIS-xx | DISPOSITION ACCEPT/REWORK/SCRAP/RETURN_TO_PROCESS/HOLD + DISPOSITION_ID legacy | 5 مقدار جدید |
| ROW_STATUS 4 مقدار | CURRENT_STATUS 5 مقدار + VOID_FLAG + VOID_REASON + ROW_STATUS legacy | Void به‌جای Delete |
| 26 ستون | 32 ستون | حداقل فیلدهای پیشنهادی دستور (21) + traceability |
| 7 جدول Master | 11 جدول (افزودن RecordType, Result, Disposition_new, Status_new) | Master Data Guide |
| Barcode History ساده (اولین رکورد) | Barcode History کامل (تا 10 ردیف با Record_ID, Date, Type, Shift, Station, Part, Defect, Result, Disposition, Status, Inspector) | مرحله پانزدهم |
| RULE_STATION فقط | RULE_STATION + RULE_LOGIC (7 قانون ناسازگاری) | مرحله هشتم |
| Dashboard 6 KPI | Dashboard 8 KPI (کل رکورد, بارکد یکتا, NOK, OK, Rework, Scrap, Return, نرخ) | حفظ تفاوت Record vs Unique |

## ۳) مهاجرت

- records 14,557 → 14,557, distinct 14,079, leading_zero 9,283, len 19=100%, non_digit 0, dup_groups 463, dup_extra 478, max 4, dates 2026-04-20→2026-09-19 123 روز
- Record_Type: INSPECTION 13,986 (اولین هر بارکد), REINSPECTION 473 (تکرار), RETURN 98 (برگشت از فروش)
- Disposition: REWORK 12,814 (اصلاحی), SCRAP 1,645 (ضایعات), RETURN_TO_PROCESS 98
- Inspection_Result: NOK 14,557 (همه دارای عیب در تاریخ), OK 0 (آینده)
- Current_Status: CLOSED 14,079 (اولین), OPEN 478 (تکراری)
- Name drift: چکمه ای→چکمه‌ای 95, رجبی  + جمعی  trim → گزارش شد، نه Fail

## ۴) تست‌ها

- 30 تست طراحی شد (docs/TestReport_v3.md) — 30/30 PASS (ساختار, بدون Excel runtime)
- Jalali: PASS (109,938 روز)
- v2 tests still PASS (T-W1..T-W12)

## ۵) چک‌لیست پذیرش دستی در Excel (باید کاربر اجرا کند)

1. FORM: 11 فیلد + C28 پیام + بارکد 19 رقمی
2. ثبت OK (بدون عیب) → RECORD_ID QC-... + CLOSED
3. ثبت NOK + عیب → REWORK/SCRAP
4. بارکد تکراری → REINSPECTION + OPEN + مشکوک-تکراری
5. بارکد 18/20 رقمی / حروف → رد
6. Void: انتخاب ردیف Data → QC_VoidRow → VOID_FLAG TRUE + دلیل الزامی
7. STAGING: B..K پر → FLAG ✓ آماده Import → QC_Import_Staging → پاک شدن STAGING
8. Reports: جمع هر بلوک = Dashboard K1 (تراز)
9. Dashboard: فیلتر تاریخ/شیفت/قطعه → KPIها به‌روز
10. Barcode History: بارکد تکراری → 2+ ردیف با Record_IDهای مختلف

## ۶) محدودیت‌ها

- یک فایل به‌ازای ایستگاه (Option A)
- Excel DB نیست
- LOG_AUDIT best-effort
- VBA 24KB > 4,142 ظرفیت بازنویسی → import
- بدون Excel در sandbox → ماکرو runtime تأیید نشده

## ۷) تحویل

1. QC-F-14-v3.xlsx (آماده, بدون ماکرو)
2. vba/QC_WritePath_v3.bas (ماکرو)
3. docs/ (7 فایل)
4. analysis/05 + 06 گزارش‌ها
5. tools/build_payload_v3.py + gen_workbook_v3.py
6. FORM-QC-F-14.xlsm مرجع

**گام بعدی:** کاربر چک‌لیست §5 را در Excel Desktop اجرا کند. هر مورد قرمز → fix در build_payload_v3.py → gen_workbook_v3.py → rebuild.
