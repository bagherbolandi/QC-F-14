# QC-F-14 v5 — نمونه‌مانند (Sample-Like) Login + Form

این نسخه دقیقاً مانند فایل نمونه شما (`sampel file.xlsm`) ساخته شده:

## ویژگی‌های مشابه نمونه

| ویژگی | نمونه (`sampel file.xlsm`) | QC-F-14 v5 |
|------|---------------------------|------------|
| **شیت قابل مشاهده در شروع** | فقط `Menu` Visible، بقیه VeryHidden | فقط `LOGIN ورود` Visible، بقیه VeryHidden |
| **لاگین با UserForm** | `Login.Show` در `Workbook_Open` | `Login.Show` در `Workbook_Open` |
| **مخفی کردن ریبون و تب‌ها** | `CloseMenu` → Ribbon=False, FormulaBar=False, WorkbookTabs=False, Headings=False | `CloseMenu` دقیقاً مشابه |
| **تمام‌صفحه** | `Application.DisplayFullScreen = True` | `Application.DisplayFullScreen = True` |
| **اعتبارسنجی** | `Information1!N3 = IF(MATCH(M3,J:J,0),"YES","NO")` | `Information1!N3` و `LOG_AUDIT!N3` مشابه |
| **ذخیره کاربران** | `HelpRamz!A23:G42` master list | `HelpRamz!A23:G42` + `USERS کاربران` (CODE/USERNAME/PASSWORD/ROLE) |
| **لاگ ورود** | `Information1!B:F` (System/User/Date/TimeIn/TimeOut) | `LOG_AUDIT` + `Information1` B:F |
| **نقش‌ها** | admin = همه شیت‌ها، بقیه = لیست دسترسی در `Information1!L5:L23` | ADMIN=همه، OPERATOR=FORM+STAGING+Data+Master+Help+Dashboard+Reports، VIEWER=Dashboard+Reports+Help+Data |
| **فرم‌ها** | `Form1..Form9`, `Login`, `warning`, `Empety`, `UserPassword` | همان `vbaProject.bin` نمونه (1.5MB) با 17 UserForm حفظ شده، فقط کد `Login`, `Module1`, `ThisWorkbook` جایگزین شده |
| **فرمت فایل** | `xlsm` با `vbaProject.bin` و ContentType `macroEnabled` | `xlsm` با `macroEnabled` + `vbaProject` rel — بدون خطای `format or extension is not valid` |

## نحوه لاگین

- فایل را باز کنید → Excel تمام‌صفحه می‌شود، فقط LOGIN دیده می‌شود، فرم لاگین وسط صفحه ظاهر می‌شود (مانند نمونه).
- اگر فرم را با X ببندید، فایل بسته می‌شود (مانند نمونه: `UserForm_QueryClose` → `ActiveWorkbook.Close` + `Application.Quit`).
- نام کاربری و رمز را وارد کنید، Enter یا دکمه لاگین:

| کاربر | رمز | نقش |
|------|-----|-----|
| admin | admin123 | ADMIN — همه شیت‌ها |
| operator | 1234 | OPERATOR — FORM, STAGING, Data, Master, Help, Dashboard, Reports |
| viewer | viewer123 | VIEWER — Dashboard, Reports, Help, Data |
| i-01 .. i-10 | 1234 | OPERATOR |
| kianloo_e | 1234 | OPERATOR (از HelpRamz) |

- پس از لاگین موفق، شیت‌های مجاز Visible می‌شوند، `OpenMenu` ریبون و تب‌ها را برمی‌گرداند (مانند نمونه).
- لاگ ورود در `LOG_AUDIT` و `Information1` ثبت می‌شود (ستون‌های B:F).

## شیت‌های QC-F-14

- `LOGIN ورود` — مانند `Menu` نمونه، فقط Visible در شروع
- `Data داده` — جدول `tblQC` با 32 ستون، 14557 رکورد مهاجرتی، `RECORD_ID = QC-00000001` immutable
- `FORM فرم ثبت` — فرم ثبت با اعتبارسنجی (تاریخ شمسی، شیفت، ایستگاه، بازرس، قطعه، قالب، بارکد 19 رقمی، نتیجه OK/NOK، تعیین تکلیف 5 مقدار، وضعیت 5 مقدار)
- `STAGING ورود اضطراری` — ورود بدون ماکرو
- `Master اطلاعات پایه` — اطلاعات پایه
- `Dashboard داشبورد` + `Reports گزارش‌ها` — 15 گزارش
- `LOG_AUDIT` + `Information1` — لاگ + اعتبارسنجی
- `USERS کاربران` + `HelpRamz` — کاربران
- `Menu`, `Data`, `List`, `Index`, `Home` — برای سازگاری با VBA نمونه (VeryHidden)

## تفاوت با v4

- v4 از `FORM-QC-F-14.xlsm` با `vbaProject.bin` 26KB استفاده می‌کرد (فضای Module1 فقط 643B) → فقط لاگین ساده sheet-based.
- v5 از `sampel file.xlsm` با `vbaProject.bin` 1.5MB استفاده می‌کند → فضای Module1 5961B، Login 7022B → امکان UserForm حرفه‌ای + `CloseMenu/OpenMenu` کامل مانند نمونه.
- ContentType قبلاً `sheet.main+xml` (xlsx) بود → خطای `format or extension is not valid` — در v5 به `macroEnabled.main+xml` اصلاح شد.

## فایل‌ها

- `QC-F-14-v5-READY-LOGIN-FORM.xlsm` (2.9M) — نسخه اصلی، آماده استفاده
- `QC-F-14-v5-SAMPLE-LIKE.xlsm` — کپی مشابه
- `QC-F-14-v4-READY-LOGIN-EMBEDDED.xlsm` (2.3M) — نسخه قبلی با فیکس ContentType

## نحوه ساخت (برای توسعه)

```bash
python3 tools/gen_workbook_v5.py --payload analysis/payload_v4.json -o /tmp/QC-F-14-v5-BASE.xlsx
python3 tools/build_v5_xlsm.py
```

- `gen_workbook_v5.py` — ساخت base xlsx با 21 شیت، LOGIN Visible، بقیه VeryHidden
- `build_v5_xlsm.py` — لود `sampel file.xlsm` → `vbaProject.bin` → جایگزینی `VBA/Module1`, `VBA/ThisWorkbook`, `VBA/Login` با `msovba.replace_module_source` → embed در xlsx + fix ContentTypes → xlsm معتبر

## نکات امنیتی

مانند نمونه، رمزها Plain Text در شیت VeryHidden هستند — Excel امنیت واقعی ندارد، فقط بازدارندگی. برای امنیت واقعی → SQL/Web.

