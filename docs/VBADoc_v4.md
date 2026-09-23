# VBA Documentation — v4 (Login + WritePath)

**Files:** 
- `vba/QC_Full_v4.bas` (standard module, ASCII, ~45KB) — main logic
- `vba/ThisWorkbook_v4.bas` (ThisWorkbook event) — Workbook_Open → HideAllExceptLogin

## Why ASCII Only (still valid)

VBE is ANSI editor. Persian literals get re-encoded via Windows codepage and ZWNJ/U+06CC not representable in cp1256. Reference workbook's own VBA is pure ASCII. So all Persian strings read from workbook: sheet names from `Settings!QC_SHEETS`, messages from `Settings!QC_MSG`, status/verdicts from Helper grids.

## Entry Points — v4

| Procedure | Sheet/Cells | Description |
|---|---|---|
| Auto_Open | — | Called automatically on open (if in standard module). Calls HideAllExceptLogin — only LOGIN visible. Also called by Workbook_Open_Handler for ThisWorkbook event. |
| Workbook_Open_Handler | — | Wrapper for ThisWorkbook.Workbook_Open. |
| HideAllExceptLogin | — | Loops ThisWorkbook.Worksheets, sets Visible=xlSheetVeryHidden except LOGIN. Clears C12/C13/C10/C6/C8 on LOGIN. |
| QC_Login | LOGIN C6=user, C8=pass, C10=msg, C12=current user, C13=role | Reads username/password, searches TBL_USERS (USERS sheet) then DIM_USER (Master fallback). Checks ACTIVE (بله/Yes/TRUE = active, خیر/No/FALSE/0 = inactive). Password case-sensitive plain compare. On success: sets C12=user (display name), C13=ROLE, C10=message per role (OK_LOGIN_ADMIN/OPERATOR/VIEWER), calls ShowSheetsByRole, logs LOGIN_OK. On fail: sets C10=E_LOGIN or E_LOGIN_INACTIVE, logs LOGIN_FAIL to LOG_AUDIT, MsgBox. |
| QC_Logout | LOGIN C12 | Checks if logged in, logs LOGOUT, calls HideAllExceptLogin, MsgBox OK_LOGOUT. |
| ShowSheetsByRole | — | ADMIN → all sheets Visible. OPERATOR → FORM, STAGING, Help, LOGIN Visible, rest VeryHidden. VIEWER → Dashboard, Reports, DQ, Help, LOGIN Visible, rest VeryHidden. Uses GetSheetName() to resolve localized sheet names from Settings!QC_SHEETS with hardcoded fallback. |
| QC_Append | FORM C6..F16, C28 | v3 logic unchanged: validates 9 required, barcode 19 digits, calendar lookup, determines Record_Type (RETURN if station برگشت else REINSPECTION if exists else INSPECTION), checks RULE-01/02 (OK+defect), derives Current_Status, generates RECORD_ID QC-00000008, writes 32 cols via ColIdx, logs, clears barcode+note. Uses UserName() which now returns logged-in user from LOGIN!C12 if present, else ENV USERNAME. |
| QC_Import_Staging | STAGING 8..207, FLAG L contains "Import" | Same as v3. |
| QC_VoidRow | Select one cell in tblQC | Same as v3: VOID with reason mandatory. |
| QC_ClearForm | — | Clears 12 inputs. |
| QC_About | — | Shows 4.0.0-login-writepath, row count, limits. |

## Users Table

- Sheet `USERS کاربران` VeryHidden
- Table `TBL_USERS` (and `DIM_USER` in Master as duplicate for admin view)
- Columns: CODE, USERNAME, PASSWORD, ROLE, DISPLAY_NAME, ACTIVE
- ROLE ∈ {ADMIN, OPERATOR, VIEWER}
- ACTIVE = بله / خیر (or Yes/No/TRUE/FALSE)
- Passwords Plain Text — documented as limitation

## Config

- QC_SHEETS now 14 keys: LOGIN, DATA, FORM, STAGING, MASTER, HELPER, REPORTS, DASH, DQ, SETTINGS, LOG, HELP, USERS, VERIFY
- QC_MSG now 28 keys: 17 from v3 + 11 login-related (E_LOGIN, E_LOGIN_INACTIVE, OK_LOGIN_ADMIN, OK_LOGIN_OPERATOR, OK_LOGIN_VIEWER, OK_LOGOUT, E_LOGOUT, LOGIN_TITLE, ROLE_ADMIN, ROLE_OPERATOR, ROLE_VIEWER)

## Security Notes

- VeryHidden ≠ security. User can open VBA and set Visible=-1. But it prevents casual unhide via Format → Sheet → Unhide.
- Passwords in sheet are readable if user unhides USERS via VBA. This is by design for Excel — real security needs SQL/Web.
- LOG_AUDIT logs LOGIN_FAIL with username attempted — helps detect brute force attempts (manual review).
- Workbook_BeforeClose (optional in ThisWorkbook_v4.bas) re-hides sheets and saves, so file always reopens to LOGIN-only state even if user left it open with ADMIN sheets visible.

## Installation v4 (3 min)

1. Open QC-F-14-v4.xlsx in Excel Desktop
2. Alt+F11 → File → Import File → `vba/QC_Full_v4.bas`
3. Double-click ThisWorkbook in Project Explorer → paste code from `vba/ThisWorkbook_v4.bas` (Workbook_Open + Workbook_BeforeClose)
4. Save as .xlsm (QC-F-14-v4.xlsm)
5. On LOGIN sheet, Insert → Shape → Button → Assign QC_Login ("ورود"), second button QC_Logout ("خروج")
6. Add folder to Trusted Locations
7. Close and reopen xlsm — only LOGIN should be visible. Test logins:
   - admin / admin123 → should see all sheets
   - operator / 1234 → FORM+STAGING+Help
   - viewer / viewer123 → Dashboard+Reports+DQ+Help
   - Logout → only LOGIN

## Limitations (same as v3 plus login)

- LOG_AUDIT best-effort, editable, not legal trail
- One file per station, no record locking
- Login is UX deterrent, not security boundary
- If macros disabled, file shows only LOGIN — no data leakage (good), but also no work possible — user must enable macros via Trusted Location
