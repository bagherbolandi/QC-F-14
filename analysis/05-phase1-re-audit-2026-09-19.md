# فاز ۱ — حسابرسی مجدد معماری داده (بدون تغییر فایل)

**تاریخ:** 2026-09-19  
**فایل‌های مبنا:** `FORM-QC-F-14.xlsm` (684KB، مرجع) + `QC-F-14-v2.xlsx` (2.0MB، نسخه بازسازی‌شده قبلی)  
**ابزار:** `analysis/cache/ref.pkl` + `tools/xlmigrate.py` + خوانش مستقیم `xl/workbook.xml` و `xl/worksheets/sheet*.xml` + `olefile` روی `xl/vbaProject.bin`  
**قاعده:** هیچ عددی فرض نشده؛ همه از خود فایل استخراج شده.

---

## ۱) نام تمام Sheetها (مرجع)

- `FORM` → `xl/worksheets/sheet1.xml` — dimension `A1:Q38`, merged 1 (`B2:H3`), protection SHA-512 با `spinCount=100000`, `sheet=1 objects=1 scenarios=1`, فقط ۹ سلول `locked=0`
- `Data` → `sheet2.xml` — dimension `A1:I29517`, **بدون protection، بدون validation، بدون فرمول، بدون merge**, `customFormat` phantom rows = عامل 6.7MB XML
- `AUTO` → `sheet3.xml` — dimension `A1:I1`, buffer موقت برای ماکرو

**نسخه v2 فعلی (۱۲ شیت):** `Data داده`, `FORM فرم ثبت`, `STAGING ورود اضطراری`, `Master اطلاعات پایه`, `Helper محاسبات` (hidden), `Reports گزارش‌ها`, `Dashboard داشبورد`, `DQ کیفیت داده`, `Settings تنظیمات`, `Help راهنما`, `LOG_AUDIT` (hidden), `Verify` (hidden) — دارای ۱۰ Table.

---

## ۲) تعداد واقعی رکوردها

- خام از XML: **14,559 ردیف** در `Data` شامل 2 ردیف کاملاً خالی
- رکورد معتبر: **14,557**
- بازه تاریخ (میلادی استخراج‌شده از serial): **2026-04-20 تا 2026-09-19**, روزهای متمایز **123**
- میانگین روزانه 118.3، بیشینه 433، برآورد سالانه ≈ 35,500

در v2: **14,557 رکورد** مهاجرت‌شده، `tblQC = A1:Z20000`, بدون phantom rows.

---

## ۳) تعداد Barcodeهای یکتا

- مرجع: **14,079** یکتا از 14,557
- leading zero: **9,283** (64%) — اگر ستون به Number تبدیل شود، غیرقابل بازگشت از دست می‌رود
- طول: **100% = 19**, non_digit = 0
- فرمت ذخیره: `t="s"` + `numFmt 49 = @` (Text) — نقطه قوت فعلی

در v2: همین 7 متریک عیناً حفظ شده (`migration_check`).

---

## ۴) تعداد رکوردهای دارای Barcode تکراری

- گروه‌های تکراری: **463**
- ردیف اضافه تکراری: **478**
- max repeat: **4 بار** برای یک بارکد
- ردیف‌های **دقیقاً تکراری** (هر ۹ فیلد): **89 گروه / 89 ردیف اضافه**, که **58 گروه متوالی** هستند → امضای «دو بار کلیک ثبت با بارکد پاک‌نشده»
- بارکد → قطعه متفاوت: **23 مورد** (نشانه برچسب اشتباه یا خطای بارکدخوان)

---

## ۵) ساختار کامل ستون‌های Data (مرجع)

| حرف | هدر | نوع واقعی | توضیح |
|---|---|---|---|
| A | تاریخ | عدد سریال اکسل با `numFmt 164 = [$-fa-IR]yyyy/mm/dd` | نمایش شمسی وابسته به تنظیمات ویندوز |
| B | شیفت | عدد 1/2/3 | بدون کد |
| C | ایستگاه کنترلی | رشته |  |
| D | نام قطعه | رشته |  |
| E | نوع قالب | رشته | **دو مفهوم مخلوط: حفره A-E + جهت راست/چپ** |
| F | شماره بارکد | رشته Text |  |
| G | شرح عیب | رشته |  |
| H | وضعیت قطعه | رشته | فقط ۲ مقدار |
| I | نام اپراتور | رشته |  |

- **بدون Excel Table**, بدون فرمول ردیفی, بدون قفل
- `Print_Area = Data!$A$1:$AE$1927` (تا AE که داده ندارد، تا 1927 در حالی که داده تا 14560 است) — کهنه

**v2:** 26 ستون (A-Z): `EVENT_ID, EVENT_DATE_J, EVENT_DATE_G, J_MONTH_KEY, J_WEEK_KEY, SHIFT_ID, STATION_ID, STATION_SNAP, PART_ID, PART_NAME_SNAP, MOLD_ID, MOLD_NAME_SNAP, DEFECT_ID, DEFECT_NAME_SNAP, DISPOSITION_ID, EVENT_TYPE, BARCODE, BC_PART_CODE, MOLD_CHECK, ENTRY_STAMP, ENTRY_USER, INSPECTOR_ID, INSPECTOR_SNAP, ENTRY_CHANNEL, ROW_STATUS, NOTE` — هر ردیف = یک Record QC, بدون merge.

---

## ۶) Validationهای موجود

**مرجع — FORM! (9 قانون):**

| سلول | type | source/formula | پیام |
|---|---|---|---|
| C5 تاریخ | list | `$J$4:$J$5` (=TODAY(), TODAY()-1) | default انگلیسی |
| C7 شیفت | list | `$K$4:$K$6` |  |
| C9 ایستگاه | list | `$L$4:$L$11` (۸) |  |
| E5 قطعه | list | `$N$4:$N$38` (۳۵) |  |
| E7 قالب | list | `$O$4:$O$10` (۷) |  |
| E9 عیب | list | `$P$4:$P$24` (۲۱) |  |
| E11 وضعیت | list | `$Q$4:$Q$5` (۲) |  |
| C11 اپراتور | list | `$M$4:$M$24` (۲۱) |  |
| G5 بارکد | textLength equal 19 | formula1=19 | «شماره بارکد صحیح نمی‌باشد.» تنها پیام فارسی |

- همه `allowBlank=1`, بدون `showError` سفارشی جز بارکد
- **هیچ validation روی Data ندارد** — و DV با Paste یا ماکرو دور زده می‌شود
- `textLength=19` عددی‌بودن را چک نمی‌کند (۱۹ کاراکتر غیرعددی هم قبول)

**v2:** 35 DV block (از Master + Helper GRID 40 ردیفی), بارکد: `textLength=19` + `AND(LEN=19, ISNUMBER(x+0))` + `number_format='@'`, تمام فیلدهای اصلی از Master با پیام فارسی.

---

## ۷) Named Rangeها

**مرجع:** فقط ۳ مورد سیستمی:
- `_FilterDatabase = Data!$A$1:$I$1`
- `Print_Area = Data!$A$1:$AE$1927`
- `Print_Area = FORM!$A$1:$Q$29`

هیچ Named Range کاربردی وجود ندارد.

**v2:** 29 Defined Name شامل `QC_SHEETS` (۸ کلید), `QC_MSG` (۱۷ کلید), `QC_CAL_G/MK/WK`, لیست‌های Master, `tblQC` و ...

---

## ۸) Excel Tableهای موجود

- مرجع: **صفر**
- v2: **۱۰ جدول**: `tblQC` + 9 جدول Master (`DIM_PART/MOLD/DEFECT/INSPECTOR/STATION/SHIFT/DISPOSITION`, `MAP_PART_MOLD`, `RULE_STATION`)

---

## ۹) فرمول‌ها

**مرجع — کل فایل ۱۱ فرمول:**
- `FORM!J4 = TODAY()`
- `FORM!J5 = TODAY()-1`
- `FORM!C15 = C5` (نمایش تاریخ)
- `C16 = C7` (شیفت)
- `C17 = C9` (ایستگاه)
- `C18 = E5` (قطعه)
- `C19 = E7` (قالب)
- `C20 = G5` (بارکد, Text)
- `C21 = E9` (عیب)
- `C22 = E11` (وضعیت)
- `C23 = C11` (اپراتور)

بدون فرمول روی `Data` و `AUTO`. بدون volatile جز ۲ مورد TODAY.

**v2:** **1,628 فرمول** فقط در Helper/Reports/Dashboard/DQ/Settings, صفر فرمول روی ردیف‌های `tblQC` (قاعده: فرمول فقط محاسبه).

---

## ۱۰) VBA Modules / Procedures

**مرجع — `xl/vbaProject.bin` 26,112 بایت, بدون رمز (DPB موجود ولی قابل باز):**

- `PROJECT`: `Module=Module1`, `Document=ThisWorkbook, Sheet1, Sheet2, Sheet3`
- `VBA/Module1` (1,036 بایت):
  ```vb
  Sub Sabt_Form()
    If WorksheetFunction.CountA(Sheets("FORM").Range("C5,C7,C9,C11,E5,E7,E9,E11,G5")) < 9 Then
      MsgBox "Please complete..."
    Range("C15:C23").Select : Selection.Copy
    Sheets("AUTO").Select : Range("A1").Select : Selection.PasteSpecial ... Transpose:=True
    Selection.Copy : Sheets("Data").Select : Range("A2").Select : Selection.Insert Shift:=xlDown
    Sheets("FORM").Select : Range("E11").ClearContents : Range("E9").ClearContents
  End Sub
  ```
- `VBA/ThisWorkbook` (1,304 بایت):
  - `Workbook_Open`: `SHOW.TOOLBAR("Ribbon",False)`, `DisplayFormulaBar=False`, `DisplayStatusBar=False`, `OnKey ^o ^n %{F11} %{f8} %{f10} %{f12} = ""`, `CommandBars("Cell"/"Ply"/"Worksheet Menu Bar").Enabled=False`, `DisplayWorkbookTabs=False` — **در سطح Application و بدون بازگشت**
  - `Workbook_BeforeClose`: `ThisWorkbook.Save` بدون توجه به Cancel — **ریسک فاجعه در حالت چندکاربره شبکه (آخرین بستن برنده)**
- `Sheet1/2/3`: stub خالی (299 بایت)

**مشکلات قطعی:**
1. CountA<9 برای ۹ سلول → پیام برعکس
2. فرم پس از ثبت پاک نمی‌شود (فقط E9/E11) → 89 تکرار متوالی
3. Insert روی سر جدول O(n) + phantom rows
4. کلیپ‌بورد + AUTO buffer بدون On Error
5. وابستگی به ترتیب C15:C23
6. توضیحات فارسی در ThisWorkbook به `?????` خراب شده (CodePage)

**v2:** بدون VBA (macro-free) + فایل `vba/QC_WritePath.bas` (23,577 بایت, 100% ASCII, CRLF) شامل ۵ نقطه ورود: `QC_Append` (C6 تاریخ شمسی, C8 شیفت, C10 ایستگاه, C12 بازرس, F6 قطعه, F8 قالب, F10 عیب, F12 محل صدور, C14 بارکد, پیام C28), `QC_Import_Staging` (rows 8..207, FLAG col K contains "Import", فقط پاک‌کردن ردیف‌های مصرف‌شده), `QC_VoidRow` (ROW_STATUS=ابطال‌شده, بدون Delete), `QC_ClearForm` (Union ۹ سلول), `QC_About`. تمام رشته‌های فارسی از `Settings!QC_SHEETS` و `QC_MSG` خوانده می‌شود.

---

## ۱۱) روابط بین Sheetها

**مرجع:**
```
FORM!C5..G5 (ورودی لیست‌محور)
  ↓ C15:C23 نمایش
AUTO!A1 Paste Transpose
  ↓ Copy
Data!A2 Insert Shift:=xlDown (جدید در بالا)
```
- Master Data در `FORM!J4:Q38` پشت قفل
- هیچ رابطه فرمولی بین Data و Master
- Print_Area و FilterDatabase فقط

**v2:**
- FORM → (VBA) → Data!tblQC (Append پایین جدول)
- STAGING → (VBA Import) → tblQC
- Master → Helper (GRID 40 ردیفی + FILTER/UNIQUE) → Reports/Dashboard/DQ
- Settings (QC_SHEETS, QC_MSG, QC_CAL_*) → VBA و فرمول‌ها (Single Source of Truth)

---

## ۱۲) Master Dataهای موجود

**مرجع (از ref.pkl):**
- SHIFT: 3 ['1','2','3']
- STATION: 8 ['پرس','پانچ','سوراخکاری','اینسرت گذاری','تریمکاری','پلیسه گیری','کنترل نهایی','برگشت از فروش'] — **فقط ۲ تا استفاده شده**: کنترل نهایی 14,459 (99.3%), برگشت از فروش 98 (0.7%), ۶ ایستگاه هرگز
- INSPECTOR: 21 نفر — آلودگی فاصله: `رجبی␣␣`=511, `جمعی␣`=918, ۳ نفر صفر رکورد
- PART: 35 — ۱۰ بدون رکورد، صدر: سینی فن سمند 3,594, سینی فن 206 2,892, سینی فن 405 2,541, سینی فن دنا 2,128
- MOLD: 7 ['A','B','C','D','E','راست','چپ'] — A 6,208, D 3,657, B 2,706, C 1,286, E 207, چپ 251, راست 242 — **دو مفهوم در یک ستون**
- DEFECT: 21 — صدر ترک اینسرت 3,357, ترک و شکستگی 1,897, «برآمدگی اجکتور» صفر رکورد
- STATUS: 2 ['اصلاحی' 12,912 (88.7%), 'ضایعات' 1,645 (11.3%)] — **برگشتی صفر**

بدون کد، بدون ACTIVE، بدون تاریخ اعتبار، محدوده ثابت `$M$4:$M$24` etc → افزودن بازرس جدید بدون رمز قفل غیرممکن.

**v2:** 9 جدول کددار `CODE|NAME|ACTIVE` + `MOLD_KIND` (حفره/جهت) + `ENABLED_FOR_FORM` + `ALIASES`, + `MAP_PART_MOLD` (استخراج 90%+), + `RULE_STATION` با `IS_DEFAULT` و `EVENT_TYPE_IF_PICKED`.

---

## ۱۳) داده تکراری، خالی، ناسازگار یا مشکوک

| # | نوع | مقدار مرجع | توضیح |
|---|---|---|---|
| 1 | ردیف خالی | 2 | ردیف 14559,14560 در Data |
| 2 | دقیقاً تکراری | 89 گروه / 89 اضافه / 58 متوالی | پاک‌نشدن بارکد |
| 3 | بارکد تکراری | 463 گروه / 478 اضافه / max 4 | Barcode شناسه قطعه است نه رکورد — تکرار لزوماً خطا نیست |
| 4 | بارکد → قطعه متفاوت | 23 | برچسب اشتباه محتمل |
| 5 | فاصله اضافی | رجبی␣␣ 511, جمعی␣ 918, سینی زیر موتور سمند␣ 368 | دو هویت در Pivot |
| 6 | تاریخ خالی | 2 | بدون validation روی Data |
| 7 | تاریخ آینده/تأخیر | میانه ۲۴ روز تأخیر ثبت نسبت به تاریخ تولید (استخراج از بارکد نیست — فعلاً قابل اثبات نیست) | نیاز به تعریف BC_PROD_DATE |
| 8 | یتیم | 0 | DV لیست کار کرده — نقطه قوت |
| 9 | بدون شناسه رکورد | کل فایل | هیچ Record ID, Timestamp, User |

DQ در v2: 4,603 issue (DQ-MOLD-RULE 4,003 + DQ-DUP 463 + DQ-BACKDATE 135 + DQ-BLANK 2)

---

## ۱۴) Current Data Model (جمع‌بندی)

مدل فعلی **Flat File ۹ ستونی بدون کلید** است:
- کلید طبیعی وجود ندارد (بارکد کلید قطعه است نه رکورد)
- تاریخ = عدد میلادی + نمایش منطقه‌ای (وابسته به ویندوز)
- شیفت عددی → Pivot جمع می‌کند نه گروه
- قالب = حفره + جهت مخلوط
- وضعیت = فقط اصلاحی/ضایعات
- بدون ممیزی, بدون کاربر ثبت, بدون زمان ثبت, بدون وضعیت رکورد

---

## ۱۵) Current Problems (اولویت‌دار)

1. **فاجعه چندکاربره:** `BeforeClose Save` روی فایل شبکه → آخرین نفر برنده
2. **سقوط بی‌صدا:** CountA + بدون پیام «ثبت شد» → شیفت کامل گم می‌شود
3. **تکرار:** بارکد پاک نمی‌شود → 89 تکرار ثابت‌شده
4. **Phantom Rows:** UsedRange تا 29517 با 14,957 ردیف استایل‌دار → XML 6.7MB
5. **قفل برعکس:** FORM رمزدار, Data بی‌حفاظ
6. **Master ثابت:** افزودن بازرس/قطعه بدون رمز غیرممکن
7. **تاریخ فقط امروز/دیروز:** ثبت با تأخیر >1 روز ممنوع
8. **بدون گزارش:** هیچ Pivot/Chart/KPI
9. **VBA شکننده:** Select/Clipboard/AUTO buffer بدون خطایابی

---

## ۱۶) Proposed Data Model — مطابق دستور جدید (بدون حدس)

### اصل کلیدی جدید شما
> **Barcode شناسه قطعه است، نه شناسه رکورد. یک Barcode می‌تواند چند Record داشته باشد.**

پس باید **Record ID** مستقل داشته باشیم.

### جدول اصلی `tblQC` — ۲۲ فیلد پیشنهادی (حذف نمی‌کنیم، فقط نگاشت)

| # | فیلد پیشنهادی شما | معادل v2 فعلی | نوع | منطق |
|---|---|---|---|---|
| 1 | Record_ID | EVENT_ID | `QC-00000001` Text | یکتا, خودکار MAX+1, غیرقابل ویرایش, ثابت پس از ثبت, تولید در VBA |
| 2 | Record_Type | **جدید — نیاز به تصمیم** | INSPECTION/REINSPECTION/RETURN | از ایستگاه + تاریخ بارکد: کنترل نهایی=INSPECTION, برگشت از فروش=RETURN, تکرار بارکد در تاریخ متفاوت=REINSPECTION |
| 3 | Date | EVENT_DATE_J (شمسی Text 10) + EVENT_DATE_G (میلادی عدد) | شمسی Text + میلادی Date | شمسی برای نمایش, میلادی برای فیلتر/بازه |
| 4 | Shift | SHIFT_ID | Code `S1/S2/S3` | تغییر از عدد به کد متنی |
| 5 | Inspection_Station | STATION_ID + SNAP | Code `ST-xx` |  |
| 6 | Inspector | INSPECTOR_ID + SNAP | Code `I-xx` |  |
| 7 | Part_Code | PART_ID | Code `P-xx` |  |
| 8 | Part_Name | PART_NAME_SNAP | Snapshot |  |
| 9 | Mold_Code | MOLD_ID | Code `M-xx` | تفکیک حفره/جهت |
| 10 | Mold_Name | MOLD_NAME_SNAP | Snapshot |  |
| 11 | Barcode | BARCODE | Text 19 | Validation: دقیقاً 19 رقم, ISNUMBER, بدون فاصله |
| 12 | Inspection_Result | **جدید — نیاز به نگاشت** | OK/NOK | از DISPOSITION: ACCEPT→OK, REWORK/SCRAP/RETURN→NOK (تصمیم) |
| 13 | Defect_Code | DEFECT_ID | Code `D-xx` |  |
| 14 | Defect_Name | DEFECT_NAME_SNAP | Snapshot |  |
| 15 | Disposition | DISPOSITION_ID → ACCEPT/REWORK/SCRAP/RETURN_TO_PROCESS/HOLD | Map: اصلاحی=REWORK, ضایعات=SCRAP, برگشتی=RETURN_TO_PROCESS |  |
| 16 | Current_Status | ROW_STATUS | OPEN/UNDER_REWORK/WAITING_REINSPECTION/CLOSED/VOID | پیش‌فرض OPEN, Void=VOID |
| 17 | Registered_By | ENTRY_USER | Username |  |
| 18 | Record_Timestamp | ENTRY_STAMP | Now() |  |
| 19 | Notes | NOTE | Text |  |
| 20 | Void_Flag | بخشی از ROW_STATUS | TRUE/FALSE | به‌جای Delete |
| 21 | Void_Reason | NOTE + LOG | Text |  |
| 22 | ... + BC_PART_CODE, MOLD_CHECK, ENTRY_CHANNEL | BC_PART_CODE, MOLD_CHECK, ENTRY_CHANNEL | کمکی | برای اعتبار متقابل |

**تفاوت با v2 فعلی:** v2 فعلی `EVENT_TYPE = اصلاحی/ضایعات/برگشتی` دارد که معادل Disposition شماست، نه Record_Type جدید. باید ستون `Record_Type` جدا اضافه شود.

---

## ۱۷) Record Type Logic

- **INSPECTION:** اولین رکورد هر بارکد، یا ایستگاه = کنترل نهایی و تاریخ جدید
- **REINSPECTION:** بارکد تکراری با فاصله زمانی >0 روز و Disposition قبلی = REWORK/HOLD → نشان‌دهنده کنترل پس از اصلاح
- **RETURN:** ایستگاه = برگشت از فروش **یا** Record_Type صریح برگشتی (تصمیم شما: آیا برگشتی نوع رخداد مستقل است یا ایستگاه؟ پیشنهاد: مستقل)

**شواهد از داده:**
- 463 بارکد تکراری → کاندید REINSPECTION
- 98 رکورد برگشت از فروش → کاندید RETURN
- 14,079 اولین رکورد → INSPECTION

پیاده‌سازی: در Helper با `COUNTIFS(BARCODE, <current>, DATE_G, "<"&current)` + `MAX(DATE)` اولین/آخرین را تشخیص دهیم؛ بدون حدس، فقط بر اساس تاریخ.

---

## ۱۸) Barcode / Record ID Logic

- **Record ID:** `QC-` + 8 رقم صفرپرشده, تولید: `MAX(RIGHT(tblQC[Record_ID],8))+1`, در VBA پس از Validation, قبل از Append, در ستون قفل‌شده, هرگز تغییر نمی‌کند
- **Barcode:** Text 19, DV: `LEN=19` + `ISNUMBER(VALUE)`, `number_format='@'`, `TRIM`, هشدار صفر ابتدایی, **تکرار خطا نیست** — فقط Flag `مشکوک-تکراری` یا `UNDER_REWORK`
- **Barcode History:** با وارد کردن Barcode, تمام رخدادها به ترتیب زمانی (Date G + Timestamp) نمایش: Record_ID, Date, Shift, Station, Defect, Result, Disposition, Status, Inspector

در v2 فعلی: `EVENT_ID` به‌صورت `E-000001` است؛ باید به `QC-00000001` تغییر فرمت دهد (ساده).

---

## ۱۹) Migration Mapping

| مرجع (9 ستون) | v2 (26 ستون) | پیشنهادی جدید (22) | تبدیل |
|---|---|---|---|
| تاریخ (A) | EVENT_DATE_J (B) + EVENT_DATE_G (C) + J_MONTH_KEY + J_WEEK_KEY | Date J+G | `excel_serial_to_date` + `xlcal.g_to_j` |
| شیفت (B) | SHIFT_ID (F) | Shift Code S1.. | `1→S1` |
| ایستگاه (C) | STATION_ID (G) + SNAP (H) | Inspection_Station | کدگذاری + TRIM |
| قطعه (D) | PART_ID (I) + SNAP (J) | Part_Code+Name | کد + snapshot + حذف فاصله اضافی (`چکمه ای`→`چکمه‌ای` 95 ردیف) |
| قالب (E) | MOLD_ID (K) + SNAP (L) + MOLD_KIND | Mold_Code+Name | تفکیک A-E=حفره, راست/چپ=جهت |
| بارکد (F) | BARCODE (Q) Text 19 | Barcode | TRIM, حفظ صفر, BC_PART_CODE=R1-2, MOLD_CHECK |
| عیب (G) | DEFECT_ID (M)+SNAP(N) | Defect_Code+Name | کد + snapshot |
| وضعیت (H) | DISPOSITION_ID (O) + EVENT_TYPE (P) | Disposition + Record_Type | اصلاحی→REWORK, ضایعات→SCRAP, برگشت از فروش→RETURN |
| اپراتور (I) | INSPECTOR_ID (V)+SNAP(W) | Inspector | کد + ادغام `رجبی  `→`رجبی` |
| — | EVENT_ID (A) | Record_ID | `QC-00000001`.. |
| — | ENTRY_STAMP (T), ENTRY_USER (U), ENTRY_CHANNEL (X) | Timestamp, Registered_By | MIGRATION → `MIGRATION`, `2026-09-19` |
| — | ROW_STATUS (Y) | Current_Status + Void_Flag | پیش‌فرض `تأییدشده`/`OPEN` |
| — | NOTE (Z) | Notes, Void_Reason | خالی |

**تطبیق:** `migration_check` قبلی: records 14557→14557, barcode 7 متریک identical, dates 2026-04-20→2026-09-19 123 روز, whitespace 1797→0, inspector 18→18, `MIGRATION FAILS: NONE` — همین باید حفظ شود.

---

## ۲۰) Master Data Design (جدید)

```
MASTER_DATA (یک شیت, 10 جدول Excel Table, CODE|NAME|ACTIVE|NOTES)

Shifts: S1|شیفت 1|بله, S2|شیفت 2|بله, S3|شیفت 3|بله
Stations: ST-01..ST-08 با ENABLED_FOR_FORM (6 تای بلااستفاده = خیر ولی ACTIVE=بله)
Inspectors: I-01..I-21 با ALIASES (رجبی  /جمعی )
Parts: P-01..P-35 با ACTIVE و نام نرمال‌شده
Molds: M-01..M-07 با MOLD_KIND=حفره/جهت + PART_ID مجاز (MAP_PART_MOLD)
Defects: D-01..D-21 با DEFECT_GROUP
RecordTypes: RT-01 INSPECTION, RT-02 REINSPECTION, RT-03 RETURN
InspectionResults: IR-01 OK, IR-02 NOK
Dispositions: DIS-01 ACCEPT, DIS-02 REWORK (اصلاحی), DIS-03 SCRAP (ضایعات), DIS-04 RETURN_TO_PROCESS (برگشتی), DIS-05 HOLD
Statuses: ST-01 OPEN, ST-02 UNDER_REWORK, ST-03 WAITING_REINSPECTION, ST-04 CLOSED, ST-05 VOID
```

در Database ترجیحاً CODE ذخیره شود, NAME برای نمایش (snapshot).

---

## ۲۱) Validation (بازطراحی)

- Barcode = دقیقاً 19 رقم, Text, `ISNUMBER(VALUE)`, `LEN=19`, `TRIM`, پیام فارسی
- Shift/Station/Inspector/Part/Mold/Defect/Disposition/RecordType/Result/Status فقط از Master (list)
- Part→Mold وابسته: اگر MAP_PART_MOLD شامل قطعه است, فقط قالب‌های مجاز (هشدار, نه بلوکه اگر داده کم است — هرگز حدس نزن)
- ناسازگاری منطقی در جدول Rule: مثلاً `Result=OK + Defect<>خالی` → خطا, `Disposition=SCRAP + Status=CLOSED` → نیاز به تصمیم (اگر معتبر است به‌عنوان استثناء ثبت)

در v2 فعلی: 35 DV + Helper GRID 40 ردیفی.

---

## ۲۲) Risks

| رتبه | ریسک | شدت | احتمال | کاهش |
|---|---|---|---|---|
| 1 | از بین رفتن رکورد در چندکاربره (BeforeClose Save) | فاجعه | بالا | فایل به‌ازای ایستگاه + تجمیع (Option A) — قبلاً تأیید شد |
| 2 | سقوط بی‌صدا رکورد ناقص | بالا | بالا | پیام فارسی با نام فیلد + پیام «ثبت شد QC-...» |
| 3 | تبدیل بارکد به عدد | بالا | متوسط | Text + DV + آموزش Import |
| 4 | بلاک ماکرو توسط IT/MotW | بالا | متوسط | STAGING + Import بدون VBA |
| 5 | قفل Master | متوسط | قطعی | Master جدا + Table |
| 6 | نبود Record ID | بالا | قطعی | QC-... خودکار |

---

## ۲۳) Decisions Required (قبل از فاز ۲)

1. **برگشتی چیست؟** نوع رخداد مستقل (RETURN) یا ایستگاه؟ پیشنهاد: مستقل + Disposition=RETURN_TO_PROCESS. در داده فعلی 98 رکورد برگشت از فروش = RETURN.
2. **۶ ایستگاه بلااستفاده:** حذف از فرم (ENABLED_FOR_FORM=خیر) یا فعال‌سازی؟ پیشنهاد: خیر ولی ACTIVE بماند.
3. **Inspection Result OK/NOK:** آیا هر رکورد بدون عیب = OK؟ در داده فعلی 100% دارای عیب هستند (حتی اصلاحی). آیا OK اصلاً ثبت می‌شود؟ اگر نه, Result همیشه NOK خواهد بود تا فرآیند تغییر کند.
4. **Disposition Mapping:** اصلاحی→REWORK, ضایعات→SCRAP, برگشتی→RETURN_TO_PROCESS, آیا ACCEPT و HOLD هم لازم است؟ پیشنهاد: بله برای آینده.
5. **Mold تفکیک:** A-E = حفره, راست/چپ = جهت — دو فیلد جدا یا یک فیلد با KIND؟ پیشنهاد: دو فیلد منطقی ولی یک ستون فیزیکی با KIND.
6. **Record Type REINSPECTION:** آیا قطعه پس از REWORK مجدداً کنترل می‌شود؟ داده تکراری 463 گروه این را نشان می‌دهد. آیا باید خودکار REINSPECTION شود؟
7. **Current Status CLOSED:** چه زمانی بسته می‌شود؟ پیشنهاد: پس از ACCEPT/SCRAP یا پس از REINSPECTION موفق.
8. **Barcode History:** آیا فقط نمایش یا امکان Void/ویرایش از همان صفحه؟
9. **Concurrency نهایی:** Option A (فایل به‌ازای ایستگاه + تجمیع Power Query) قبلاً انتخاب شد — تأیید نهایی؟

---

## ۲۴) وضعیت v2 فعلی نسبت به دستور جدید

| دستور جدید | وضعیت v2 | فاصله |
|---|---|---|
| Record ID یکتا خودکار غیرقابل تغییر | ✅ EVENT_ID `E-...` دارد, باید به `QC-00000001` تغییر فرمت | کم |
| Barcode ≠ Record ID, تکرار خطا نیست | ✅ BARCODE Text 19 + ROW_STATUS مشکوک-تکراری, نه رد | ✅ |
| Record Type INSPECTION/REINSPECTION/RETURN | ❌ فقط EVENT_TYPE=اصلاحی/ضایعات/برگشتی | نیاز به ستون جدید |
| Inspection Result OK/NOK | ❌ ندارد | نیاز به ستون جدید |
| Defect Code/Name | ✅ DEFECT_ID+SNAP | ✅ |
| Disposition ACCEPT/REWORK/SCRAP/... | ⚠️ فقط اصلاحی/ضایعات/برگشتی, باید به 5 مقدار نگاشت | متوسط |
| Current Status OPEN/.../VOID | ⚠️ ROW_STATUS=تأییدشده/پیش‌نویس/ابطال‌شده/مشکوک-تکراری — باید به 5 مقدار جدید نگاشت | متوسط |
| Barcode History | ✅ Reports!Barcode History بلوک + Dashboard فیلتر | ✅ ولی باید با Record_Type جدید به‌روز شود |
| tblQC واقعی بدون Merge | ✅ A1:Z20000 Table | ✅ |
| Master جدا با Code/Name | ✅ 9 جدول | ✅ + 3 جدول جدید (RecordTypes, Results, Statuses) |
| Validation 19 رقم + Master | ✅ | ✅ |
| Rule ناسازگاری | ⚠️ فقط MOLD_CHECK, نیاز به جدول Rule کامل | متوسط |
| FORM → Validation → ID → Timestamp → Append | ✅ QC_Append همین مسیر | ✅ |
| Void به‌جای Delete | ✅ QC_VoidRow | ✅ |
| STAGING جدا | ✅ STAGING ورود اضطراری + FLAG | ✅ |
| Reports/Dashboard/DQ | ✅ 15 گزارش + 10 فیلتر + DQ 4,603 | ✅ |
| VBA فقط ثبت/Import/Void/Clear/About | ✅ 5 نقطه ورود, 23,577 بایت ASCII | ✅ |

**نتیجه:** v2 فعلی 70% دستور جدید را پوشش می‌دهد؛ 30% باقی‌مانده (Record_Type, Inspection_Result, Disposition 5تایی, Status 5تایی, Rule Table, Barcode History تقویت‌شده) در فاز ۲ قابل افزودن است بدون بازسازی از صفر.

---

## ۲۵) گام بعدی (فاز ۲) — فقط پس از تأیید شما

1. افزودن 3 جدول Master جدید: `DIM_RECORDTYPE`, `DIM_RESULT`, `DIM_STATUS_NEW` + نگاشت Disposition به 5 مقدار
2. افزودن ستون‌های جدید به `tblQC`: `RECORD_TYPE_ID`, `INSPECTION_RESULT`, `CURRENT_STATUS_NEW` (یا نگاشت ROW_STATUS موجود) — بدون حذف ستون‌های قدیمی برای حفظ مهاجرت
3. تغییر فرمت `EVENT_ID` به `QC-00000001`
4. به‌روزرسانی `RULE_STATION` + جدول جدید `RULE_LOGIC` (OK+NOK/Defect, SCRAP+CLOSED etc)
5. به‌روزرسانی `FORM` (فیلدهای جدید Record_Type, Result) + `QC_Append` + `STAGING`
6. به‌روزرسانی Reports: Barcode History با Record_Type, فیلتر جدید, Pareto با Result
7. تست 30گانه + `migration_check` + `test_workbook.py` T-W1..T-W12 + `make_vba.py`
8. تحویل `QC-F-14-v3.xlsx` + `vba/QC_WritePath_v3.bas`

**در این فاز هیچ Sheet یا داده‌ای تغییر نکرده؛ فقط گزارش.**

