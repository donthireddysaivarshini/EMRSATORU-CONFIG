# Forms panel + print all scales — design

Status: **built** · 2026-09-17 — see SETUP.md "Forms panel". Differences from
this draft: the panel lives at `/exports/forms/` (no new nginx location);
cht-conf is used to convert and validate, but the publish writes the form
document itself (same fields as `cht upload-app-forms`); the test harness runs
CHT 5.3's engine built from our cht-core, not the harness's bundled 4.11; the
harness spike found the MoCA fluency bug only bites on **edit** (SCORING_AUDIT
M1) and found a new DDST-II bug (X4). Uploading media with a form is not
supported yet: existing images are kept.

Two features:

- **A. Forms panel** — a non-technical admin uploads an XLSForm in the
  browser and it goes live safely, with scoring checked.
- **B. Print all scales** — one PDF with every scale a participant has done.

---

## A. Forms panel

### Goal

A CHT admin uploads a `.xlsx` form (plus any images) and publishes it
without a terminal. **Scoring must not silently break**, and anything
published must be recoverable.

### Decisions already made

| Question | Decision |
|---|---|
| Where does the live version of a form live? | **The database.** The panel publishes; the repo is no longer the source of truth for app forms. |
| Who can publish? | **Any CHT admin, one step.** No second-person approval. Because of that, the automatic checks below are the only safety net: every check blocks, and the four known scales cannot be published without passing tests. Only a brand-new form with no tests yet can go live, behind a red warning. |

### Facts this design rests on (all checked on 2026-09-17)

- CHT authors forms in XLSForm, but the **server only accepts XML**. The
  conversion is done by `cht-conf`, which bundles pyxform
  (`xls2xform-medic`). The panel runs that same tool — no new converter.
- **cht-conf's own validation passed the MoCA form that can never award the
  fluency point.** CHT's checks cover structure, not scoring. The panel must
  test scoring itself.
- An app form is stored as doc `form:<id>` with attachments `xml`,
  `form.html`, `model.xml` and any media. CHT regenerates `form.html` /
  `model.xml` from the XML when it changes.
- Each report stores `form_version`, so we can tell which form version
  produced its scores.
- CHT's login cookie (`AuthSession`, path `/`) reaches pages under the same
  host. From inside the Docker network, `GET http://haproxy:5984/_session`
  with that cookie returns the user and roles; a forged cookie is rejected.
- `cht-conf-test-harness` 5.0.4 fills forms in the real CHT form engine
  running in headless Chromium. Debian packages Chromium for arm64 (Oracle
  Ampere) and amd64.

### Architecture

The existing `exports` container becomes **satoru-tools**, still one
container in the same compose stack. It gains Node 22, cht-conf, the test
harness and Chromium alongside the current Python app.

```
browser ──► nginx ──► /exports/*  ─┐
                      /forms/*    ─┴─► satoru-tools ──► CouchDB (medic)
                      /*          ───► api               └─► CouchDB (satoru-forms-history)
```

Nothing new to host: one more nginx `location`, same container count.

### Login — for both tabs

Every request checks the `AuthSession` cookie against `/_session` and
requires the `_admin` or `admin` role. No cookie or the wrong role → redirect
to `/medic/login?redirect=<page>`.

This also fixes an existing problem: **today `/exports/` has no login**, so
anyone who can reach the URL can download every participant's data. That fix
ships first and does not wait for the panel.

### Publish pipeline

Every step must pass. If any step fails, **nothing is published** and the
admin sees what failed in plain words.

1. **Upload.** `.xlsx` required; images as extra files or a `.zip`. The form
   id comes from the file name. Re-uploading an existing form keeps its title,
   icon and "who can fill it" settings; a new form asks for a title and who it
   is for (participants / camps) — no JSON editing.
2. **Convert** with cht-conf `convert-app-forms` in a throwaway project
   folder.
3. **cht-conf validation** (`validate-app-forms`, with the instance, so all
   9 checks run).
4. **Scoring safety checks** (ours):
   - a field the assessor types into must not also have a `calculation`
     (this is exactly the MoCA fluency bug)
   - every formula refers only to fields that exist
   - known scales (DDST-II, DST, MoCA, VSMS) still contain their result
     fields (`moca_total_score`, `social_quotient_rounded`, …)
   - fields that exist in **already-stored reports** but disappear from the
     new version are listed with how many reports they affect
5. **Test cases.** Each known scale has a file of answer sets and the
   results they must produce, run through the test harness. **Any failure
   blocks.** A brand-new form with no test cases may be published, but with a
   red "scoring untested" banner on the confirm step.
6. **Review screen.** Plain list of scoring formulas added, removed and
   changed compared with the live version, plus test results. The admin
   presses **Publish**.
7. **Publish.**
   1. save the current live version to history (below) — if this fails, stop
   2. upload the form with cht-conf `upload-app-forms`
   3. regenerate report layout + labels (`gen_report_config.py`) and upload
      settings and translations
   4. exports reload the form from the database
8. **Verify.** Re-read `form:<id>`, confirm the new version is live, record a
   history entry: who, when, file hash, test results.

### Test cases

One file per scale, kept with the forms, written from the scale's manual.

```yaml
# tests/moca_assessment.yaml  (education_years exists only after audit fix M2)
- name: everything correct, 12 years schooling
  answers: { all_items: correct, fluency_count: 14, education_years: 12 }
  expect:  { moca_raw_total: 29, moca_education_adj: 1, moca_total_score: 30, moca_interpretation: normal }

- name: fluency 11 words earns the point
  answers: { fluency_count: 11 }
  expect:  { moca_language_total: 1 }
```

MoCA can be written now from `scales/moca/moca-instructions-english_2010.pdf`.
VSMS and DDST-II wait for the NIMS answers listed in `SCORING_AUDIT.md`.

### History and rollback

- Stored in a **separate CouchDB database, `satoru-forms-history`** — not in
  `medic`, which is copied to every offline phone.
- Each entry keeps the `.xlsx`, generated XML, properties, media, publisher,
  time and test results.
- **Rollback** re-publishes a chosen entry through the same pipeline, so its
  tests run again.

### Keeping the database the single source of truth

- Exports and the report-layout generator read form XML **from CouchDB**
  instead of files baked into the image.
- `scripts/pull_live_forms.py` downloads every live form (with its `.xlsx`)
  into the repo, for developers.
- Docs change: **app forms are no longer uploaded with
  `cht upload-app-forms`.** A developer running it from an old copy would
  silently undo a published fix.
- The panel shows **"changed outside the panel"** when a live form's hash
  differs from its last panel publish.

### Existing reports

Never modified by this feature. The forms page shows how many stored reports
were produced by each version. Re-scoring old reports (needed after the
MoCA and VSMS fixes) is a separate, explicit, backed-up job — **out of scope
here**.

### Out of scope

- Re-scoring stored reports
- Second-person approval (declined)
- Readable report layout for **brand-new** scales: layout rules exist only for
  the four known scales, so a new scale renders like stock CHT until rules are
  added. Later option: mark result fields in the spreadsheet itself.

### Risks

| Risk | Handling |
|---|---|
| Harness 5.0.4 may not score exactly like CHT 5.3.0 | **Spike first:** run the harness on today's MoCA form and confirm it reproduces the fluency bug. If it can't catch a known bug, the approach is wrong. |
| Chromium on ARM | Use Debian's `chromium` package, point the harness at it |
| Bigger container (Node + Chromium, several hundred MB) | Acceptable on the target 12 GB box; Chromium runs only during a test |
| One-step publish | Blocking checks + mandatory tests for known scales + one-click rollback |

---

## B. Print all scales

### Problem

Ravi has two scales. The Reports list shows both, but **Print / Save as PDF
only prints the report that is open.**

### Design

- On the expanded report, next to the current button:
  **"Print all scales for <name>"**, shown only when the participant has more
  than one report.
- Loads all reports whose subject is that participant, formats each one
  exactly as the screen does (existing `FormatDataRecordService`), and builds
  one print document:
  - a cover block with participant details
  - one section per scale, oldest first, each starting on a new page
- Same print path as today (`window.print()` → Save as PDF), same translated
  labels.
- cht-core change → webapp rebuild and redeploy.

---

## Build order

| # | Piece | Why here |
|---|---|---|
| 0 | Harness fidelity spike | Everything in A depends on it |
| 1 | Login on `/exports/` | Live data exposure; independent |
| 2 | Exports + report layout read forms from CouchDB; `pull_live_forms.py` | Makes the database the truth before anything writes to it |
| 3 | Upload → convert → validate → scoring checks → review screen | |
| 4 | MoCA test cases | First real tests |
| 5 | Publish + history + rollback | |
| 6 | **Fix MoCA fluency (M1) through the panel** | The panel's first real use |
| 7 | Print all scales (B) | Independent; can run in parallel |
