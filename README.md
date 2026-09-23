# QC-F-14 — سامانه QC مبتنی بر Excel (نسخه ۴ با لاگین نقش‌محور)

این مخزن شامل **فایل‌های آماده اجرا** و **ابزارهای بازتولید و آزمون** است. هیچ سلولی دستی ویرایش نشده؛ هر فایل خروجی `tools/` است.

## فایل‌های تحویل

| مسیر | چیست | وضعیت |
|---|---|---|
| `QC-F-14-v4.xlsx` / `.xlsm` | **نسخه ۴ آماده اجرا — با لاگین نقش‌محور** — 14 شیت (LOGIN + 13), `tblQC` با **32 ستون** و **14,557 رکورد** (RECORD_ID QC-00000001..), 15 گزارش با Barcode History کامل, Dashboard با Record vs Unique, DQ, Help, USERS. فقط LOGIN در شروع Visible (بقیه VeryHidden). ماکرو QC_Login/QC_Logout + WritePath | **جدید — مطابق دستور لاگین** |
| `QC-F-14-v3.xlsx` | نسخه ۳ — 12 شیت، 32 ستون, 14,557 رکورد | قبلی، حفظ شد |
| `QC-F-14-v2.xlsx` | نسخه ۲ (26 ستون, EVENT_ID) — 2.0 MB | قبلی |
| `FORM-QC-F-14.xlsm` | فایل مرجع اصلی (684KB) — فقط خواندنی | مرجع |
| `vba/QC_Full_v4.bas` | VBA v4 کامل (45KB, ASCII): QC_Login, QC_Logout, Auto_Open/HideAllExceptLogin, ShowSheetsByRole + QC_Append, QC_Import_Staging, QC_VoidRow, QC_ClearForm, QC_About | جدید |
| `vba/ThisWorkbook_v4.bas` | Workbook_Open + BeforeClose — برای ThisWorkbook | جدید |
| `vba/QC_WritePath_v3.bas` | VBA v3 (24KB) | قبلی |
| `docs/` | UserGuide_v4 (لاگین), VBADoc_v4, + v3 docs | جدید |
| `analysis/payload_v4.json` | Payload v4 با DIM_USER (8 کاربر پیش‌فرض) | جدید |
| `tools/` | `build_payload_v4.py`, `gen_workbook_v4.py` + v3/v2 tools | |

## معماری داده v4 (لاگین + v3)

- **Record_ID** = `QC-00000001` یکتا، خودکار MAX+1، غیرقابل تغییر
- **Barcode** = شناسه قطعه، نه رکورد — تکرار خطا نیست (463 گروه تکراری → REINSPECTION)
- **Record_Type**: INSPECTION 13,986 / REINSPECTION 473 / RETURN 98
- **Inspection_Result**: OK/NOK (مهاجرت: NOK 14,557)
- **Disposition**: ACCEPT/REWORK (12,814)/SCRAP (1,645)/RETURN_TO_PROCESS (98)/HOLD — 5 مقدار
- **Current_Status**: OPEN (478)/UNDER_REWORK/WAITING_REINSPECTION/CLOSED (14,079)/VOID — 5 مقدار
- **Void**: VOID_FLAG TRUE + VOID_REASON + CURRENT_STATUS=VOID
- **Master Data**: 12 جدول (11 قبلی + DIM_USER) + MAP_PART_MOLD + RULE_STATION + RULE_LOGIC (7 قانون)
- **DIM_USER**: CODE, USERNAME, PASSWORD, ROLE, DISPLAY_NAME, ACTIVE — 8 کاربر پیش‌فرض (admin/admin123 ADMIN, operator/1234 OPERATOR, viewer/viewer123 VIEWER, i-01..i-05)
- **LOGIN**: شیت LOGIN ورود با C6=user, C8=pass, C10=msg, C12=current user, C13=role — فقط LOGIN در شروع Visible (بقیه VeryHidden)
- **نقش‌ها**: ADMIN=همه شیت‌ها, OPERATOR=FORM+STAGING+Help, VIEWER=Dashboard+Reports+DQ+Help
- **Barcode History**: Reports R11

## ساختن دوباره و آزمون

```bash
# v4 (جدید با لاگین)
python3 -W ignore tools/build_payload_v4.py FORM-QC-F-14.xlsm -o analysis/payload_v4.json
python3 -W ignore tools/gen_workbook_v4.py --payload analysis/payload_v4.json -o QC-F-14-v4.xlsx
cp QC-F-14-v4.xlsx QC-F-14-v4.xlsm

# v3
python3 -W ignore tools/build_payload_v3.py FORM-QC-F-14.xlsm -o analysis/payload_v3.json
python3 -W ignore tools/gen_workbook_v3.py --payload analysis/payload_v3.json -o QC-F-14-v3.xlsx

# تست‌ها
python3 -W ignore tools/test_jalali.py
```

**آخرین اجرا v4:**
- records 14,557 → 14,557, distinct barcode 14,079, leading_zero 9,283, len 19=14,557, dup_groups 463, dup_extra 478, max_repeat 4
- Record_Type: INSPECTION 13,986 / REINSPECTION 473 / RETURN 98
- Disposition: REWORK 12,814 / SCRAP 1,645 / RETURN_TO_PROCESS 98
- Users: 8 (admin, operator, viewer, i-01..i-05)
- Sheets: 14 (LOGIN visible, rest VeryHidden)
- MIGRATION FAILS: NONE

## نصب ماکرو v4 (۳ دقیقه — لاگین الزامی)

1. QC-F-14-v4.xlsx را باز کنید (Excel Desktop)
2. Alt+F11 → File → Import File → `vba/QC_Full_v4.bas`
3. ThisWorkbook: دابل‌کلیک ThisWorkbook → کد `vba/ThisWorkbook_v4.bas` را Paste کنید (Workbook_Open → HideAllExceptLogin)
4. Save As `.xlsm` (QC-F-14-v4.xlsm)
5. LOGIN: Developer → Insert → Button → Assign QC_Login ("ورود"), دوم QC_Logout ("خروج")
6. FORM: Shape → Assign QC_Append, STAGING → QC_Import_Staging, Data → QC_VoidRow (ADMIN)
7. Trusted Location (پوشه را Trusted کنید)
8. ببندید و دوباره باز کنید — فقط LOGIN باید دیده شود. تست:
   - admin / admin123 → همه شیت‌ها
   - operator / 1234 → FORM+STAGING+Help
   - viewer / viewer123 → Dashboard+Reports+DQ+Help
   - Logout → فقط LOGIN

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
