# VBA Documentation — v3

**File:** `vba/QC_WritePath_v3.bas` (24,310 bytes, ASCII, CRLF)

## Why ASCII Only

VBE is ANSI editor. Persian literals get re-encoded via Windows codepage and ZWNJ/U+06CC not representable in cp1256. Reference workbook's own VBA is pure ASCII. So all Persian strings read from workbook: sheet names from `Settings!QC_SHEETS`, messages from `Settings!QC_MSG`, status/verdicts from Helper grids.

## Entry Points

| Procedure | Form Cells | Description |
|---|---|---|
| QC_Append | C6 تاریخ, C8 شیفت, C10 ایستگاه, C12 بازرس, F6 قطعه, F8 قالب, C14 بارکد, F10 نتیجه, F12 عیب, F14 تعیین تکلیف, C16 وضعیت, F16 توضیح; message C28 | Validates 9 required, barcode 19 digits, calendar lookup via CODE_J_TEXT, determines Record_Type (RETURN if station contains برگشت else REINSPECTION if exists else INSPECTION), checks RULE_LOGIC (OK+defect), derives Current_Status (REWORK→UNDER_REWORK, HOLD→OPEN else CLOSED), generates RECORD_ID QC-00000008, writes full 32 columns, logs, clears barcode+note, shows OK with ID. Duplicate barcode → flagged as مشکوک-تکراری + OPEN, not rejected. |
| QC_Import_Staging | STAGING rows 8..207, FLAG col L contains "Import" | Reads B..K (date, shift, station, inspector, part, mold, barcode, result, defect, disposition, status), same validation, determines Record_Type, writes, clears consumed rows (only clearing allowed). |
| QC_VoidRow | Select one cell in tblQC DataBodyRange | Sets CURRENT_STATUS=VOID, VOID_FLAG=TRUE, VOID_REASON from InputBox (mandatory), ROW_STATUS=ابطال‌شده, logs. Never deletes. |
| QC_ClearForm | Union of 12 input cells | Clears form for next scan. |
| QC_About | — | Shows version 3.0.0-writepath, row count, user, honest limits. |

## Contracts

- TBL_QC = tblQC
- Reads Master tables: RULE_STATION, DIM_DISPOSITION (legacy) for backward compat
- Reads grids: CODE_MOLD_CHECK (5), CODE_ROW_STATUS (4), CODE_RECORDTYPE (3), CODE_RESULT (2), CODE_DISPOSITION_NEW (5), CODE_STATUS_NEW (5)
- Reads config: QC_SHEETS (8 keys), QC_MSG (16 keys)
- Writes: all 32 columns via ColIdx (header token contains column name), no hardcoded column letters.
- LOG_AUDIT ring 3..202, full → skip logging (never blocks)
- No `.Rows.Delete`, `.Unprotect`, `.SaveAs`, `Shell`, `DisplayAlerts`, `Visible=False`, `Application.Run`, `.Formula=` — banned by make_vba.py gate.

## Installation (2 min)

1. Open QC-F-14-v3.xlsx in Excel Desktop
2. Alt+F11 → File → Import File → select `vba/QC_WritePath_v3.bas`
3. Save as .xlsm (e.g., QC-F-14-v3.xlsm)
4. On FORM sheet, draw rectangle, right-click → Assign Macro → QC_Append
5. Optional: assign QC_Import_Staging to STAGING button, QC_VoidRow to Data sheet button, QC_ClearForm to clear button
6. Add folder to Trusted Locations to avoid MotW block.

## Limitations

- LOG_AUDIT best-effort, editable by user, not legal audit trail.
- Duplicate detection via COUNTIF on Text — works for 19-digit text.
- One file per station — Excel has no record locking.
- If macro blocked, STAGING path works without VBA.
