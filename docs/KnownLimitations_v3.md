# Known Limitations — v3

## Excel is not a Database

- No foreign key enforcement, no transaction, no rollback. RECORD_ID = COUNTA+1 can duplicate if two users append simultaneously. **Mitigation:** one file per station + central aggregation (Option A, confirmed).
- Sheet protection password is deterrence, not security — recoverable in minutes with free tools. Real security → SQL/PostgreSQL.

## Concurrency

- Excel has no record locking. `Workbook_BeforeClose Save` in original file caused data loss (last closer wins). v3 removes that event, but still single-writer per file.
- Readers unlimited if file opened read-only.

## Audit Trail

- LOG_AUDIT is best-effort ring 200 rows (3..202). User can edit. Not legal audit trail. For ISO/IATF → DB with trigger.

## Barcode

- Stored as Text `@` with DV `LEN=19 AND ISNUMBER`. If opened via CSV without Text import, leading zeros lost irreversibly (9,283 barcodes affected). Always import with column as Text.
- Duplicate barcode is NOT error — it's history. Flagged as REINSPECTION + OPEN + مشکوک-تکراری.

## Formulas & Performance

- 1,628 formulas only in Helper/Reports/Dashboard/DQ, zero on data rows. COUNTIFS bounded to N_ROWS (14,557) for ~1-2 sec refresh.
- GRID blocks fixed at 40 rows — if Master exceeds 40, increase range(40) in build_helper.
- Calendar 2018-01-01 → 2035-12-31 (6,574 days). Beyond that, extend xlcal.CAL_END.

## VBA

- ASCII only (VBE ANSI). Persian strings from workbook.
- Module size 24k — cannot inject length-preserving into old vbaProject.bin (capacity 4,142). So deliver .xlsx + .bas import.
- No Excel in sandbox → macro runtime (Assign Macro, Application.Match, COUNTIF on 19-digit text) not verified in Excel. User must run 10-item checklist in Excel Desktop.

## Missing KPIs (by design)

- نرخ ضایعات بر تولید: no production denominator in file → marked ❌
- تأخیر ثبت: barcode digit 9-14 not production date (checked) → removed.

## Data Quality

- MOLD_CHECK: 13,087 هم‌خوان / 822 ناهم‌خوان / 493 قالبِ جهت / 155 رقم خارج از بازه — warning only, never blocks.
- 6 unused stations kept ACTIVE but ENABLED_FOR_FORM=خیر — decision required if to activate.

## When to Leave Excel

- >3 concurrent writers, or >25k rows/year, or need signature/NCR/photo → PostgreSQL + Web App + Power BI.
