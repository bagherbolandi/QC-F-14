# User Guide — QC-F-14 v4 (Login + WritePath)

## نصب و فعال‌سازی ماکرو (الزامی برای لاگین)

1. فایل `QC-F-14-v4.xlsx` را باز کنید (Excel Desktop 365/2021 — Web Excel ماکرو ندارد)
2. Alt+F11 → File → Import File → `vba/QC_Full_v4.bas` → Import
3. همچنین ThisWorkbook: در Project Explorer روی ThisWorkbook دابل‌کلیک → کد `vba/ThisWorkbook_v4.bas` را کپی کنید (Workbook_Open)
4. Save As → `QC-F-14-v4.xlsm` (Macro-Enabled)
5. پوشه را Trusted Location کنید: File → Options → Trust Center → Trusted Locations → Add new location → پوشه فایل را اضافه کنید (Subfolders تیک)
6. فایل xlsm را ببندید و دوباره باز کنید — فقط شیت LOGIN باید دیده شود. اگر همه شیت‌ها دیده شد، ماکرو فعال نیست.

## لاگین نقش‌محور (v4 جدید)

### صفحه LOGIN

- `LOGIN ورود` تنها شیت قابل مشاهده در شروع است (بقیه VeryHidden)
- C6 = نام کاربری، C8 = رمز عبور (Plain Text — Excel امنیت واقعی نیست)
- C10 = پیام ورود
- C12 = کاربر فعلی، C13 = نقش

### کاربران پیش‌فرض (از DIM_USER)

| USERNAME | PASSWORD | ROLE | DISPLAY_NAME |
|----------|----------|------|--------------|
| admin | admin123 | ADMIN | مدیر سیستم |
| operator | 1234 | OPERATOR | اپراتور پیش‌فرض |
| viewer | viewer123 | VIEWER | نمایش |
| i-01 | 1234 | OPERATOR | بازرس 01 |
| i-02 | 1234 | OPERATOR | بازرس 02 |
| i-03 | 1234 | OPERATOR | بازرس 03 |
| i-04 | 1234 | OPERATOR | بازرس 04 |
| i-05 | 1234 | OPERATOR | بازرس 05 |

- ADMIN: همه شیت‌ها (Data, FORM, STAGING, Master, Reports, Dashboard, DQ, Settings, Help, LOG_AUDIT, USERS, Verify, Helper)
- OPERATOR: فقط FORM + STAGING + Help + LOGIN
- VIEWER: Dashboard + Reports + DQ + Help + LOGIN

### دکمه‌های لاگین

1. در LOGIN، Developer → Insert → Button
2. یک دکمه بکشید → Assign Macro → `QC_Login` → نام دکمه = "ورود"
3. دکمه دوم → `QC_Logout` → "خروج"
4. دکمه‌ها را کنار C6/C8 قرار دهید

### جریان کار

- ورود: نام کاربری و رمز را وارد → دکمه ورود → پیام موفقیت + شیت‌های مجاز Visible
- خروج: دکمه خروج → همه شیت‌ها VeryHidden به‌جز LOGIN + LOG_AUDIT ثبت خروج
- ورود ناموفق: پیام E_LOGIN + ثبت در LOG_AUDIT (LOGIN_FAIL)
- Workbook_Open: هر بار فایل باز می‌شود، همه شیت‌ها VeryHidden می‌شوند — بدون لاگین هیچ داده‌ای قابل دسترسی نیست

### امنیت

- Excel امنیت واقعی نیست: VeryHidden و Sheet Protection با چند کلیک دور زده می‌شود
- رمزها Plain Text در USERS (VeryHidden) — برای امنیت واقعی → SQL Server / Web App
- این لاگین فقط بازدارندگی برای اپراتور خط است، نه جایگزین کنترل دسترسی سازمانی
- پیشنهاد: فایل را در SharePoint با دسترسی AD قرار دهید + این لاگین به‌عنوان لایه دوم

## ثبت روزانه (پس از ورود OPERATOR/ADMIN)

1. FORM فرم ثبت:
   - تاریخ وقوع (شمسی)* — از لیست 180 روز اخیر
   - شیفت* — S1/S2/S3
   - ایستگاه* — فقط فعال‌ها
   - بازرس* — از لیست
   - قطعه* — کد | نام
   - قالب* — A-E یا راست/چپ
   - بارکد* — اسکن 19 رقمی (صفر ابتدایی حفظ)
   - نتیجه بازرسی* — OK (بدون عیب) یا NOK (دارای عیب)
   - عیب — اگر NOK الزامی (RULE-01/02)
   - تعیین تکلیف* — ACCEPT, REWORK, SCRAP, RETURN_TO_PROCESS, HOLD
   - وضعیت فعلی — اگر خالی، خودکار: REWORK→UNDER_REWORK, HOLD→OPEN, بقیه CLOSED
   - توضیح — اختیاری
2. پیام C28 باید "✓ آماده ثبت" باشد
3. دکمه ثبت (QC_Append) → RECORD_ID مثل QC-00014558 ساخته می‌شود
4. اگر بارکد تکراری: نوع رکورد خودکار REINSPECTION + وضعیت OPEN + ROW_STATUS مشکوک-تکراری (رد نمی‌شود، فقط پرچم)

## Barcode History (ADMIN/VIEWER)

- Reports → R11: بارکد را در D وارد کنید
- تعداد رکورد + تا 10 ردیف تاریخچه با Record_ID, تاریخ, نوع, شیفت, ایستگاه, قطعه, عیب, نتیجه, تعیین تکلیف, وضعیت, بازرس

## STAGING (بدون VBA یا برای OPERATOR)

- اگر ماکرو بلاک شد: STAGING ورود اضطراری
- B..K را پر کنید
- ستون L FLAG باید "✓ آماده Import" باشد
- سپس QC_Import_Staging → به tblQC اضافه و ردیف STAGING پاک

## Void (ابطال) — ADMIN

- در Data، یک سلول از ردیف موردنظر را انتخاب → QC_VoidRow
- دلیل الزامی (RULE-07)
- رکورد حذف نمی‌شود: CURRENT_STATUS=VOID, VOID_FLAG=TRUE, ROW_STATUS=ابطال‌شده

## گزارش‌ها و داشبورد — ADMIN/VIEWER

- 15 گزارش در Reports
- فیلترها در Dashboard G4:G16 — "همه" = بدون فیلتر
- K1 کل رکوردها (14557 مهاجرت), K2 بارکد یکتا (14079)
- نمودار روند ماهانه Disposition

## Master Data — ADMIN

- فقط در Master اطلاعات پایه، داخل Table
- افزودن: ردیف جدید در Table, CODE یکتا, ACTIVE=بله
- غیرفعال: ACTIVE=خیر
- USERS: در USERS کاربران (VeryHidden) — فقط ADMIN پس از لاگین می‌بیند. افزودن کاربر جدید = ردیف جدید در TBL_USERS با USERNAME یکتا, ROLE=ADMIN/OPERATOR/VIEWER, ACTIVE=بله

## LOG_AUDIT

- هر APPEND, IMPORT, VOID, LOGIN_OK, LOGIN_FAIL, LOGOUT ثبت می‌شود
- best-effort — اگر کاربر ماکرو را غیرفعال کند، لاگ نمی‌نویسد
- 200 ردیف اول — برای آرشیو، شیت را کپی کنید

## محدودیت‌ها

- یک فایل به‌ازای ایستگاه — چندکاربره همزمان روی یک فایل = از دست رفتن داده
- Excel دیتابیس نیست
- LOG_AUDIT best-effort
- سقف عملیاتی ~25k ردیف/سال
- لاگین Excel امنیت سازمانی نیست — فقط UX

## پشتیبان

- روزانه کپی در OneDrive/SharePoint با Version History
- پس از هر روز کاری، xlsm را ببندید تا HideAllExceptLogin اجرا و فایل با فقط LOGIN ذخیره شود
