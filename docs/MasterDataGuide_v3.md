# Master Data Guide — v3

**Sheet:** `Master اطلاعات پایه` — 11 Tables

## Structure

Each DIM_* has CODE | NAME | ACTIVE (+ extra)

| Table | Fields | Count (migrated) | Example |
|---|---|---|---|
| DIM_PART | CODE, NAME, ACTIVE | 35 | P-01 \| سینی فن 405 \| بله |
| DIM_MOLD | CODE, NAME, MOLD_KIND, ACTIVE | 7 | M-01 \| A \| حفره \| بله ; M-06 \| راست \| جهت \| بله |
| DIM_DEFECT | CODE, NAME, ACTIVE | 21 | D-01 \| کمی مواد \| بله |
| DIM_INSPECTOR | CODE, NAME, ALIASES, ACTIVE | 21 | I-01 \| هادی پور \|  \| بله (رجبی  → رجبی) |
| DIM_STATION | CODE, NAME, ENABLED_FOR_FORM, ACTIVE | 8 | ST-07 \| کنترل نهایی \| بله \| بله ; ST-01 پرس \| خیر \| بله |
| DIM_SHIFT | CODE, NAME, ACTIVE | 3 | S1 \| شیفت 1 \| بله |
| DIM_DISPOSITION (legacy) | CODE, NAME, ACTIVE | 3 | DIS-01 اصلاحی |
| DIM_RECORDTYPE | CODE, NAME, ACTIVE, DESC | 3 | RT-01 INSPECTION, RT-02 REINSPECTION, RT-03 RETURN |
| DIM_RESULT | CODE, NAME, ACTIVE | 2 | RES-01 OK, RES-02 NOK |
| DIM_DISPOSITION_NEW | CODE, NAME, ACTIVE, LEGACY | 5 | DP-01 ACCEPT, DP-02 REWORK (اصلاحی), DP-03 SCRAP (ضایعات), DP-04 RETURN_TO_PROCESS (برگشتی), DP-05 HOLD |
| DIM_STATUS_NEW | CODE, NAME, ACTIVE | 5 | STT-01 OPEN, STT-02 UNDER_REWORK, STT-03 WAITING_REINSPECTION, STT-04 CLOSED, STT-05 VOID |

Plus:
- MAP_PART_MOLD: PART → dominant MOLD with SHARE% and ENFORCED flag (only if >=90% and >=20 records). Example: محافظ شیشه آریسان → A 99.7%
- RULE_STATION: STATION → EVENT_TYPE (legacy) + IS_DEFAULT (1 if پیش‌فرض)
- RULE_LOGIC: 7 rules for inconsistency (OK+defect, NOK without defect, SCRAP+CLOSED, etc.)

## Code vs Name

- Database stores CODE, snapshot stores NAME at registration time.
- If NAME changes later, historical data stays intact (snapshot), but new registrations use new NAME.
- ACTIVE=خیر → filtered out from Helper lists (not shown in form), but historical rows keep CODE.
- ENABLED_FOR_FORM=خیر → 6 unused stations (پرس, پانچ, ...) hidden from form but ACTIVE remains for future.

## How to Add

1. Go to Master sheet, find Table, add row at bottom (Table auto-expands)
2. CODE must be unique, format as above (P-xx, M-xx, etc.)
3. NAME trimmed, no trailing spaces
4. ACTIVE=بله
5. For MOLD: set MOLD_KIND=حفره or جهت
6. Helper lists auto-update (40-row GRID), Dashboard/Reports use new values immediately.

## Normalization Applied

- `چکمه ای` (95 rows) → `چکمه‌ای` (ZWNJ)
- `رجبی  ` (511) + `رجبی` → `رجبی` with ALIAS
- `جمعی ` (918) + `جمعی` → `جمعی`
- `سینی زیر موتور سمند ` (368) → trimmed

All reported in migration_check as name_drift, not fail.
