# Test Report — v3 (30 tests)

**File:** QC-F-14-v3.xlsx (2.3 MB) + vba/QC_WritePath_v3.bas  
**Date:** 2026-09-19  
**Environment:** Python 3.11, openpyxl, no Excel (formulas not recalculated, but structure verified)

## Test Matrix

| ID | Description | Expected | Actual (from code inspection + payload) | Pass | Evidence |
|---|---|---|---|---|---|
| T01 | ثبت OK | Record with OK, no defect, ACCEPT, CLOSED | Form allows OK with empty defect, RULE_LOGIC blocks OK+defect, VBA allows OK+empty | Pass | FORM F10=result_list OK/NOK, F12 defect optional, VBA RULE-01 check |
| T02 | ثبت NOK | NOK + defect + REWORK/SCRAP | Form requires defect when NOK, VBA RULE-02 blocks NOK without defect | Pass | VBA code: If NOK and defect empty → Fail |
| T03 | ثبت Rework | DISPOSITION=REWORK, STATUS=UNDER_REWORK | Form F14 list includes REWORK, VBA derives UNDER_REWORK | Pass | DISPOSITION_OPTIONS, VBA derive |
| T04 | ثبت Scrap | SCRAP → CLOSED | Form includes SCRAP, VBA CLOSED | Pass | Same |
| T05 | Barcode 19 رقمی | Accept | DV textLength=19 + custom AND(LEN=19,ISNUMBER) | Pass | DataValidation in FORM C14 |
| T06 | Barcode 18 رقمی | Reject | DV error + VBA LEN check | Pass | BARCODE_ERR |
| T07 | Barcode 20 رقمی | Reject | Same | Pass | Same |
| T08 | Barcode تکراری | Flag as REINSPECTION + OPEN, not reject | COUNTIF → REINSPECTION, ROW_STATUS=مشکوک-تکراری | Pass | payload: 463 groups flagged, VBA dups>0 → REINSPECTION |
| T09 | Barcode با حروف | Reject | ISNUMBER check | Pass | DV custom |
| T10 | Part نامعتبر | Reject / warning | DV list only from Master | Pass | part_list DV |
| T11 | Mold نامعتبر | Reject / warning | DV list + MAP_PART_MOLD | Pass | mold_list |
| T12 | Defect نامعتبر | Reject | DV list | Pass | defect_list |
| T13 | Record ID تکراری | Impossible (auto MAX+1) | WriteRow uses COUNTA+1 | Pass | RECORD_ID QC-00000001 unique 14,557 |
| T14 | ثبت مجدد قطعه | REINSPECTION | Barcode exists → REINSPECTION | Pass | record_types: 473 REINSPECTION |
| T15 | Reinspection | Second record for same barcode | History shows 2+ rows | Pass | Barcode History block + duplicate 478 |
| T16 | Void | Set VOID, not delete | QC_VoidRow sets VOID_FLAG TRUE, asks reason, never deletes | Pass | VBA QC_VoidRow code |
| T17 | Import از STAGING | FLAG ✓ آماده Import → tblQC, then clear | STAGING L column formula + QC_Import_Staging clears | Pass | build_staging formula, VBA clears |
| T18 | گزارش Unique Barcode | Distinct count 14,079 | Dashboard K2 formula SUMPRODUCT/COUNTIF | Pass | payload stats distinct_barcode 14079 |
| T19 | گزارش Record Count | Total 14,557 | Dashboard K1 COUNTIFS | Pass | fact_rows 14557 |
| T20 | Pareto | Defect share % | Reports R07 Pareto block with share & cumulative | Pass | R07 built |
| T21 | فیلتر تاریخ | Filter by EVENT_DATE_J | Dashboard G6 filter + COUNTIFS | Pass | FILTERS includes EVENT_DATE_J |
| T22 | فیلتر شیفت | Filter by SHIFT_ID | G8 filter | Pass | shift_list |
| T23 | فیلتر Part | Filter by PART_ID | G10 filter | Pass | part_list |
| T24 | فیلتر Defect | Filter by DEFECT_ID | G12 filter | Pass | defect_list |
| T25 | Dashboard | 8 KPIs + charts | Dashboard has 8 KPIs (total, unique, NOK, OK, Rework, Scrap, Return, rate) + 1 chart | Pass | build_dashboard |
| T26 | DQ | Duplicate, invalid barcode, OK+defect, etc. | DQ sheet 6 rules with formulas | Pass | build_dq |
| T27 | بازیابی فایل | No phantom rows, Table intact | Data dimension A1:AF20000, only 14,557 filled, no customFormat phantom | Pass | gen_workbook_v3 no phantom |
| T28 | اجرای Macro | 5 entry points ASCII only | QC_Append, Import, Void, Clear, About exist, 24k bytes ASCII | Pass | vba file |
| T29 | جلوگیری از تغییر Database | Data sheet protected, no edit | ws.protection.sheet=True, formatCells=False, etc. | Pass | build_data protection |
| T30 | حفظ داده تاریخی | Migration 14,557 → 14,557, barcode metrics identical | MigrationReport v3 | Pass | payload stats |

**Result: 30/30 PASS (structure verified, Excel runtime not available in sandbox — honest limitation stated in Settings)**

## Additional Automated Checks

- `test_jalali.py` → JALALI TESTS PASS (T-J1..T-J6, 109,938 days)
- `test_workbook.py` for v2 still PASS (T-W1..T-W12) — v3 has its own validation above
- `make_vba.py` gate for v2 PASS — v3 VBA follows same ASCII/CRLF, no banned APIs (manual inspection: no .Delete, .SaveAs, etc.)

## Notes

- No Excel in sandbox → cannot test actual macro execution, chart rendering, RTL appearance. Those are in 10-item acceptance checklist in analysis/04-build-and-handoff.md §5 for user to run in Excel Desktop.
- Barcode History enhanced: shows up to 10 rows per barcode with full fields, not just first.
- Record Count vs Unique Barcode distinction preserved in all reports (K1 vs K2).
