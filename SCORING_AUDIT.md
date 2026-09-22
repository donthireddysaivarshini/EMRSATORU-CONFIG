# Scoring audit — DDST-II, DST, MoCA, VSMS

First audit 2026-09-17 from the form files. **Updated the same day after
running every scale in CHT 5.3's real form engine** (`tests/forms/`). Testing
changed some findings: one was wrong, and it found a worse bug.

**Important:** scores are calculated when a form is submitted and stored in
the report. Fixing a form **does not fix reports already saved.**

## Summary

| # | Scale | Problem | Status |
|---|---|---|---|
| X4 | DDST-II | **Every report said overall "Suspect"**, even when the child passed everything | ✅ fixed in form · old reports: `scripts/rescore_ddst_overall.py` |
| M1 | MoCA | Typed fluency count wiped when a saved report is opened for **edit** | ✅ fixed · old edited reports can't be recovered |
| M2 | MoCA | +1 education point almost never applied | ✅ fixed (new question: years of education) · old reports can't be re-scored |
| T1 | all | Every saved date/time in India says **+06:30** instead of +05:30 (1 hour off) | ✅ fixed in cht-core webapp |
| M3 | MoCA | Assessment time always "now", overwritten on edit | ✅ fixed |
| A1 | VSMS, DST | Age text wrong around birthdays (display only) | ✅ fixed |
| F1 | 3 forms | The .xlsx files did not match the live forms | ✅ synced |
| V1 | VSMS | Social age = raw score, not converted — SQ too low, worse with age | ❌ **open — needs the VSMS table from NIMS** |
| X1 | DDST-II | One real delay gives overall **Normal** | ❓ needs clinician decision |
| X2 | DDST-II | Not standard Denver II rules; no "caution" zone | ❓ needs clinician decision |
| V2 | VSMS | No above-average bands | ❓ needs clinician decision |
| X3 | DDST-II | No "Untestable" result | ❓ low |
| D1 | DST | Arithmetic | ✅ verified by tests |
| E1 | MoCA | No consent saved total "NaN" and interpretation **"impaired"** | ✅ fixed: no score, "not assessed (no consent)" |
| E2 | DDST-II | Birth date after test date → no items → false **"Normal"** | ✅ fixed: refused |
| E3 | DST, VSMS | Newborn (assessed on birth day): DQ "NaN", VSMS "Social Developmental Delay" | ✅ fixed |
| E4 | DST, VSMS | Birth date checked against today, not the assessment date (a child exactly 15 on the test day was refused) | ✅ fixed |
| E5 | 10 forms | "Default today / now" never fills in CHT, dates were blank | ✅ fixed in all 10 app forms |
| E6 | 3 scales | Assessment date could be in the future | ✅ fixed: refused |

All scales now have scoring tests: `cd tests/forms && CHROMIUM_PATH=/usr/bin/chromium npm test`
(64 passing, 1 pending = V1).

---

## DDST-II

### X4 — Every report said "Suspect" ✅ fixed

```
any_suspect = d_ps_status = 'Suspect' or ...     -> saved as the text "false"
overall     = if(any_suspect or caution_domains >= 2, 'Suspect', 'Normal')
```

A saved field is text. In a formula, any non-empty field counts as **true** —
including the text "false". So `overall` was "Suspect" for every child.

Found by the test "all passed: every domain Normal, overall Normal".

**Fix:** `overall` now checks the four domain results directly.
**Old reports:** the four domain results were saved correctly, so the overall
result can be recomputed. `scripts/rescore_ddst_overall.py` shows the changes;
`--apply` backs up the reports first, then fixes them. A static test now
blocks this mistake in any form.

### X1 — A single real delay is reported Normal ❓

```
per item : failed/refused AND child is past the item's age  -> "delayed"
domain   : 0 delayed = Normal · 1 = Caution · 2+ = Suspect
overall  : Suspect if any domain Suspect, or 2+ domains on Caution; else Normal
```

In standard Denver II any Delay makes the test Suspect. **Decide:** is that
NIMS's intended rule? The tests currently lock in the rule above.

### X2 — Not Denver II rules ❓

One age cut-off per item, and items only appear once the child is past it —
so the caution zone (75th–90th percentile) is never tested. Fine if intended.

### X3 — No "Untestable" (low)

---

## MoCA (checked against `scales/moca/moca-instructions-english_2010.pdf`)

### M1 — Fluency count wiped on edit ✅ fixed

*First audit said "the point can never be awarded". That was wrong:* in the
real engine a **new** form keeps the typed count. But the field also had a
formula (`true()`), and CHT re-runs formulas when a saved report is opened
for **edit** — which blanks the count and removes the point. The field was
also not required, so a blank count silently scored 0.

**Fix:** formula removed, count required (0–100).

### M2 — Education point ✅ fixed

Manual: *"Add one point for 12 years or fewer of formal education."* The form
only gave the point for the exact text `illiterate` or `primary` in a free-text
field.

**Fix:** new required question **"Years of formal education"** (shows what
registration says as a hint); `+1 if years <= 12`. Old reports have no years
recorded, so they can't be re-scored automatically.

### M3 — Assessment time ✅ fixed

Now fills in "now" only while empty, so the assessor can enter when the paper
test was given, and editing keeps it. (A plain "default" does not work in CHT:
CHT always pre-fills forms, so first-load defaults never fire.)

### Verified by tests

Section maxima 5+3+6+3+2+5+6 = 30 · fluency 11 = point, 10 = no point ·
serial 7s 0–3 · only uncued recall counts · total capped at 30 · 26+ normal.

---

## VSMS

### V1 — Social age not converted ❌ open

```
raw_score         = count of "yes"                 correct
social_age_months = raw_score  (1 : 1)             wrong
SQ                = social_age_months / age x 100
```

A 15-year-old passing every item gets social age 89 months → SQ 49 →
"Social Developmental Delay". **Needs the raw score → social age table from
the VSMS manual NIMS uses. Do not guess it.** Until then, treat VSMS SQ and
interpretation as unreliable, especially for older children.

The test `raw score -> social age uses the VSMS manual table` is pending
until the table arrives.

### V2 — No above-average bands ❓

---

## DST — verified ✅

18 age bands, 88 items, each passed item earns band width ÷ items in band.
Tests: 24 months all passed → DA 24, DQ 100; two failed 18–24 month items
→ DA 21, DQ 88; exactly 2 years shows "2y 0m"; any "not assessed" → no DQ.

Still worth one check against the DST manual: that each item sits in the
right age band.

---

## Cross-cutting

### T1 — Timezone one hour off ✅ fixed (cht-core)

CHT's `getTimezoneOffsetAsTime` does `Math.abs(Math.floor(-330 / 60))` = 6, so
India is written as **+06:30**. Every `now()` / date value from a form was
stored an hour off. Still unfixed in upstream CHT. Fixed in
`webapp/src/js/enketo/medic-xpath-extensions.js` with a unit test
(India, Nepal, Newfoundland).

### A1 — Age text around birthdays ✅ fixed

Years used 365.25 days, leftover months used 365. Now both come from the same
decimal age.

### F1 — Spreadsheets did not match the live forms ✅ synced

`ddst.xlsx` had lost the ages of 4 language items, `vineland.xlsx` had broken
SQ formulas and no participant picker, `moca_assessment.xlsx` had different
upload fields. Uploading those spreadsheets would have broken the live forms.
All four now rebuild the live forms exactly, then the fixes above were applied.

## Needs from NIMS clinicians

1. VSMS raw score → social age table (V1) — **most important**
2. VSMS interpretation bands (V2)
3. DDST-II: confirm what one delay should mean (X1) and the single cut-off method (X2)
