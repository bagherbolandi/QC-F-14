# Migration Report — v2 → v3 (and original → v3)

**Source:** `FORM-QC-F-14.xlsm` (700,350 bytes, 14,559 rows inc 2 blank)  
**Intermediate:** `QC-F-14-v2.xlsx` (2,019,880 bytes)  
**Target:** `QC-F-14-v3.xlsx` (2.3 MB, 14,557 records, 32 columns, tblQC A1:AF20000, capacity 20,000)

## Counts

| Metric | Source | v2 | v3 |
|---|---|---|---|
| records | 14,557 valid (14,559 raw) | 14,557 | 14,557 |
| distinct barcode | 14,079 | 14,079 | 14,079 |
| leading_zero | 9,283 (64%) | 9,283 | 9,283 |
| len_hist 19 | 14,557 | 14,557 | 14,557 |
| non_digit | 0 | 0 | 0 |
| dup_groups | 463 | 463 | 463 |
| dup_extra_rows | 478 | 478 | 478 (now flagged as REINSPECTION 473 + RETURN overlap 5) |
| max_repeat | 4 | 4 | 4 |
| dates | 2026-04-20 → 2026-09-19, 123 days | same | same |
| whitespace | 1,797 polluted | 0 | 0 |
| inspector distinct | 18 used of 21 | 18 | 18 |

**MIGRATION FAILS: NONE** (migration_check logic)

## Name Drift (intentional normalization)

- `چکمه ای` → `چکمه‌ای` (95 rows, ZWNJ)
- `رجبی  ` (511) + `رجبی` → `رجبی`
- `جمعی ` (918) + `جمعی` → `جمعی`
- `سینی زیر موتور سمند ` (368) → trimmed

Reported as `name_drift_vs_new`, not fail.

## Column Mapping

| Old (9) | v2 (26) | v3 (32) | Transform |
|---|---|---|---|
| تاریخ | EVENT_DATE_J + G + MONTH + WEEK | same + RECORD_ID etc | serial→Jalali via xlcal |
| شیفت | SHIFT_ID S1.. | same | 1→S1 |
| ایستگاه | STATION_ID + SNAP | same | code + snapshot, 6 unused → ENABLED_FOR_FORM=خیر |
| قطعه | PART_ID + SNAP | same | code + ZWNJ fix |
| قالب | MOLD_ID + SNAP + KIND | same | A-E=حفره, راست/چپ=جهت |
| بارکد | BARCODE Text 19 | same | TRIM, preserve zero |
| عیب | DEFECT_ID + SNAP | same | code |
| وضعیت | DISPOSITION_ID + EVENT_TYPE | DISPOSITION 5 values + EVENT_TYPE legacy | اصلاحی→REWORK, ضایعات→SCRAP, برگشتی→RETURN_TO_PROCESS |
| اپراتور | INSPECTOR_ID + SNAP | same | code + alias |

New fields:
- RECORD_ID: QC-00000001..QC-00014557 (was E-000001)
- RECORD_TYPE: INSPECTION 13,986 / REINSPECTION 473 / RETURN 98 (derived: station=برگشت→RETURN else exists→REINSPECTION else INSPECTION)
- INSPECTION_RESULT: NOK 14,557 (all have defect, future OK possible)
- CURRENT_STATUS: CLOSED 14,079 / OPEN 478 (first occurrence CLOSED, duplicates OPEN)
- VOID_FLAG FALSE, VOID_REASON empty, NOTE empty
- ENTRY_USER MIGRATION, ENTRY_CHANNEL مهاجرت

## Integrity Checks

- No phantom rows (Data dimension = A1:AF20000, but only 14,557 rows filled, rest empty Table rows)
- Barcode stored as Text (`@`), DV `LEN=19 AND ISNUMBER`
- RECORD_ID unique, sequential, immutable
- No formulas on data rows (1,628 formulas only in Helper/Reports/Dashboard/DQ)
- 15 tables, 29 defined names

## How to Verify

```bash
python3 tools/build_payload_v3.py FORM-QC-F-14.xlsm -o analysis/payload_v3.json
python3 tools/gen_workbook_v3.py --payload analysis/payload_v3.json -o QC-F-14-v3.xlsx
# check counts as above
```
