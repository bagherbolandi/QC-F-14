# Data Dictionary — QC-F-14 v3

**File:** `QC-F-14-v3.xlsx` / `QC-F-14-v3.xlsm`  
**Table:** `tblQC` = `Data داده!A1:AF20000` (32 columns, 14,557 migrated rows, capacity 20,000)

## Columns

| # | Column | FA Header | Type | Source | Required | Description |
|---|---|---|---|---|---|---|
| 1 | RECORD_ID | شناسه رکورد | Text QC-00000001 | VBA auto MAX+1 | Yes | یکتا، خودکار، غیرقابل ویرایش، ثابت پس از ثبت |
| 2 | RECORD_TYPE | نوع رکورد | INSPECTION/REINSPECTION/RETURN | Derived: station=برگشت→RETURN, barcode exists→REINSPECTION else INSPECTION | Yes | ماهیت رخداد QC |
| 3 | EVENT_DATE_J | تاریخ وقوع (شمسی) | Text 1405/06/28 | Form list (180 days) + Helper calendar | Yes | شمسی متنی 10 کاراکتر |
| 4 | EVENT_DATE_G | تاریخ میلادی | Date | Lookup QC_CAL_G from Helper | Yes | برای فیلتر بازه |
| 5 | J_MONTH_KEY | کلید ماه | Number 140506 | Calc from G | Auto | گروه‌بندی ماهانه |
| 6 | J_WEEK_KEY | کلید هفته | Text 1405-W26 | Calc شنبه‌مبدأ | Auto | گزارش هفتگی |
| 7 | SHIFT_ID | شیفت | S1/S2/S3 | Master DIM_SHIFT | Yes | Code متنی |
| 8 | STATION_ID | کد ایستگاه | ST-xx | Master DIM_STATION | Yes | Code |
| 9 | STATION_SNAP | ایستگاه (اسنپ‌شات) | Text | Snapshot from Master at registration | Yes | نام در لحظه ثبت |
| 10 | INSPECTOR_ID | کد بازرس | I-xx | Master DIM_INSPECTOR | Yes | Code |
| 11 | INSPECTOR_SNAP | بازرس (اسنپ‌شات) | Text | Snapshot | Yes | نام در لحظه ثبت |
| 12 | PART_ID | کد قطعه | P-xx | Master DIM_PART | Yes | Code |
| 13 | PART_NAME_SNAP | قطعه (اسنپ‌شات) | Text | Snapshot | Yes | نام قطعه |
| 14 | MOLD_ID | کد قالب | M-xx | Master DIM_MOLD | Yes | Code |
| 15 | MOLD_NAME_SNAP | قالب (اسنپ‌شات) | Text | Snapshot | Yes | نام قالب/حفره |
| 16 | BARCODE | بارکد | Text 19 digits | Barcode scanner, Text format @ | Yes | شناسه قطعه، نه رکورد — تکرار مجاز |
| 17 | BC_PART_CODE | کد قطعه بارکد | Text 2 digits | Left(BARCODE,2) | Auto | اعتبار متقابل |
| 18 | MOLD_CHECK | بررسی حفره | هم‌خوان/ناهم‌خوان/قالبِ جهت/... | Derived: mold letter vs barcode digit 4 | Auto | هشدار DQ، نه بلوکه |
| 19 | INSPECTION_RESULT | نتیجه بازرسی | OK/NOK | Form | Yes | OK=بدون عیب, NOK=دارای عیب |
| 20 | DEFECT_ID | کد عیب | D-xx | Master DIM_DEFECT | Conditional | اگر NOK الزامی |
| 21 | DEFECT_NAME_SNAP | عیب (اسنپ‌شات) | Text | Snapshot | Conditional | |
| 22 | DISPOSITION | تعیین تکلیف | ACCEPT/REWORK/SCRAP/RETURN_TO_PROCESS/HOLD | Form | Yes | |
| 23 | CURRENT_STATUS | وضعیت فعلی | OPEN/UNDER_REWORK/WAITING_REINSPECTION/CLOSED/VOID | Form or derived | Yes | گردش رکورد |
| 24 | ENTRY_USER | کاربر ثبت | Text | Environ(USERNAME) | Auto | |
| 25 | ENTRY_STAMP | زمان ثبت | DateTime | Now() | Auto | |
| 26 | ENTRY_CHANNEL | کانال ثبت | ماکرو/فرم, ماکرو/STAGING, مهاجرت | VBA | Auto | |
| 27 | ROW_STATUS | وضعیت رکورد قدیمی | تأییدشده/پیش‌نویس/ابطال‌شده/مشکوک-تکراری | Legacy | Auto | برای سازگاری |
| 28 | NOTE | توضیح | Text | Optional | No | |
| 29 | VOID_FLAG | پرچم ابطال | TRUE/FALSE | Void logic | Auto | به‌جای Delete |
| 30 | VOID_REASON | دلیل ابطال | Text | InputBox on Void | Conditional | |
| 31 | EVENT_TYPE | نوع رخداد قدیمی | اصلاحی/ضایعات/برگشتی | Legacy from RULE_STATION | Auto | برای سازگاری گزارش‌های قدیمی |
| 32 | DISPOSITION_ID | کد محل صدور قدیمی | DIS-xx | Legacy | Auto | |

**No merged cells in Database. No formulas on data rows.**

## Rules

- Barcode = Text, exactly 19 digits, `AND(LEN=19, ISNUMBER(x+0))`, TRIM
- Shift/Station/Inspector/Part/Mold/Defect/RecordType/Result/Disposition/Status only from Master
- Part→Mold: if MAP_PART_MOLD enforced, only allowed molds (warning, not block if low data)
- Logic inconsistencies in RULE_LOGIC table (e.g., OK + defect → error)

## Barcode History

- Table `Reports!R11` shows for a given barcode: count, plus up to 10 rows with Record_ID, Date, Record_Type, Shift, Station, Part, Defect, Result, Disposition, Current_Status, Inspector sorted by date.

## Record Count vs Unique Barcode

- Record Count = COUNT(tblQC)
- Unique Barcode = DISTINCT(BARCODE) via SUMPRODUCT/COUNTIF
- Dashboard shows both: K1=14,557 records, K2=14,079 unique (migration). Never mix.
