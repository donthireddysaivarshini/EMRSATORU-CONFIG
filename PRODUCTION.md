# Will it break in production? — checklist

Checked 2026-09-17 on the local stack. **Must fix** = don't put real patients
on it until done.

## Must fix before real patients

| # | Risk | What happens | Fix |
|---|---|---|---|
| 1 | **TLS certificate** is `local-ip.medicmobile.org` | its private key is public: anyone on the network can read traffic | real domain + Let's Encrypt certificate in nginx |
| 2 | **Default passwords** (`medic` / `password`) | anyone can log in as admin | new `COUCHDB_PASSWORD` in `satoruemr.env`, strong passwords for every user |
| 3 | **No scheduled backups** | a dead disk loses every record | nightly `scripts/backup_raw.py` to another machine / bucket, and **test a restore once** |
| 4 | **VSMS social quotient is wrong** (SCORING_AUDIT V1) | older children look delayed | get the VSMS table from NIMS; until then don't use VSMS SQ for decisions |
| 5 | **Old DDST-II reports all say "Suspect"** (X4) | wrong referrals from past data | on the real server: `scripts/rescore_ddst_overall.py --url ...` then `--apply` |
| 6 | **"Upgrade" button in CHT Admin** | replaces the api image → our reports/print/search/timezone changes disappear | never upgrade from Admin; upgrades = rebuild cht-core on the new version and re-run the tests |

## Fixed today

- `/exports/` had **no login** → now CHT admins only, every download logged.
- App changes were copied into the running container by hand and **vanished if
  it was recreated** → now baked into `satoru-cht-api:5.3.0` (compose overlay).
- Every saved date/time was **1 hour off** (+06:30 for India) — CHT bug.
- DDST-II always "Suspect", MoCA fluency/education/time — see SCORING_AUDIT.md.
- Spreadsheets didn't match live forms — uploading them would have broken 3 scales.
- Forms can now only go live through checks + scoring tests, with history.
- A publish during an Excel export could mix form versions → locked.

## Checked, OK

| Thing | Result |
|---|---|
| Unauthenticated / forged / non-admin access to exports and forms panel | redirected / 403, tested |
| Cross-site form posts, oversized uploads | blocked (403 / 413), tested |
| Publishing a form keeps its title, icon, "who can fill it" and images | yes, tested |
| Publish a bad MoCA (old bugs) | blocked by checks, tested |
| Roll back to an earlier version | works, tested; double publish refused |
| CHT's own unit tests for changed files | 51 passing |
| Memory: exports container idle | ~35 MB; Chromium (~300 MB) only while checking a form |
| Oracle A1 (4 vCPU / 12 GB) | fine; whole stack ~1 GB RAM |

## Good to know

- Intern privacy (only own participants for 24 h, others by exact search) is
  enforced in the app screens, not in sync: phones still hold their camp's
  data. Use device screen locks and only give interns their own camps.

- The exports image is **2.3 GB** (Node + Chromium for the scoring tests). Disk
  only; fine on a 50 GB+ boot volume.
- After anyone publishes from the forms panel, the **database** has the real
  forms. Run `scripts/pull_live_forms.py` before editing forms in the repo.
- A login can only fill forms if it is linked to a person. Link it with a
  **contact only, no home place** — a home place makes the People tab list
  that camp's people instead of the camps.
- Staff (screeners) are stored as "participants" inside a camp, so they show
  up in that camp's list. Exports leave out anyone who submitted reports, but
  a separate staff contact type would be cleaner.
- Existing MoCA reports can't get the education point retroactively (years
  were never recorded) and edited ones may have lost the fluency count.
- Offline users get form and app updates on their next sync / reload.
