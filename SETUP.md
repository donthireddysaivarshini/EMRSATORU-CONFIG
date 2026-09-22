# Satoru EMR — setup

Two repos, side by side:

```
EMRCHT/
  cht-core/    the CHT platform, patched: readable reports, Print / Print all scales, exact search, timezone fix
  satoruemr/   this repo: forms, settings, translations, Excel exports, forms panel, scoring tests, deploy overlay
```

Neither works alone. The cht-core patch is **inert** unless this repo's
`report_display` config is uploaded — that config is what tells it which
fields to hide and how to turn `P` into `Pass`.

Patched cht-core: https://github.com/defaultname-ayan/cht-core-satoru-reports

---

## 1. Prerequisites

| Tool | Version | Note |
|---|---|---|
| Docker | any recent | user must be in the `docker` group |
| Node | **22.x** | cht-core needs >=22.15.0; Node 25 is untested |
| Python | 3.10+ | for the export scripts |
| cht-conf | ^6.x | `npm i -g cht-conf` (the `cht` command) |

```bash
sudo usermod -aG docker $USER && sudo systemctl enable --now docker   # then log out/in
python -m venv .venv && .venv/bin/pip install -r scripts/requirements.txt
```

## 2. Start CHT — **use 5.3.0 images, not 5.2.0**

```bash
V=5.3.0
D=~/.medic/cht-docker/satoruemr-dir
mkdir -p "$D/compose" "$D/couch"
BASE="https://staging.dev.medicmobile.org/_couch/builds_4/medic:medic:${V}"
curl -sS -o "$D/compose/cht-core.yml"    "$BASE/docker-compose/cht-core.yml"
curl -sS -o "$D/compose/cht-couchdb.yml" "$BASE/docker-compose/cht-couchdb.yml"

cat > ../cht-core/satoruemr.env <<EOF
NGINX_HTTP_PORT=10080
NGINX_HTTPS_PORT=10443
COUCHDB_USER=medic
COUCHDB_PASSWORD=password
COUCHDB_SECRET=$(openssl rand -hex 16)
COUCHDB_UUID=$(openssl rand -hex 16)
COUCHDB_DATA=$D/couch
CHT_COMPOSE_PATH=$D/compose
CHT_NETWORK=cht-net
SATORU_REPO=$(pwd)
SATORU_EXPORTS_NGINX_TEMPLATE=$(pwd)/deploy/nginx-server.conf.template
CHT_CORE_REPO=$(cd ../cht-core && pwd)
EOF
```

> `COUCHDB_PASSWORD=password` is the CHT docker-helper default and is fine
> for a local dev box only. Change it for anything reachable by anyone else.
> `satoruemr.env` is gitignored in cht-core — never commit it.

> **Why 5.3.0.** The patch is cut from upstream master after the 5.2.0
> release, and the webapp calls `POST /api/v1/report/summary`. The 5.2.0
> API does not have that route, so it 404s and the Reports view hangs on
> a spinner forever. 5.3.0 has it.

**Build the webapp first** (step 4) — starting the stack builds an api image
from that build output. Then start everything (CHT, exports, forms panel)
with one command:

```bash
C=~/.medic/cht-docker/satoruemr-dir/compose
alias satoru='docker compose --env-file ../cht-core/satoruemr.env \
  -f '"$C"'/cht-core.yml -f '"$C"'/cht-couchdb.yml -f '"$(pwd)"'/deploy/cht-exports.yml'
satoru up -d
```

Trust the TLS cert (optional, removes browser warnings):

```bash
curl -s -o /tmp/c https://local-ip.medicmobile.org/fullchain
curl -s -o /tmp/k https://local-ip.medicmobile.org/key
docker cp /tmp/c <nginx-container>:/etc/nginx/private/cert.pem
docker cp /tmp/k <nginx-container>:/etc/nginx/private/key.pem
docker exec <nginx-container> nginx -s reload
```

## 3. Upload this config

```bash
python scripts/gen_report_config.py        # regenerate report_display + labels
cht --url=https://medic:PASSWORD@<HOST>:10443 --accept-self-signed-certs --force \
    upload-app-settings upload-app-forms upload-contact-forms upload-custom-translations
```

`--force` matters: without it cht-conf prompts before overwriting and dies
in any non-interactive shell. Hard-refresh the browser afterwards —
translations are cached client-side.

Run `gen_report_config.py` after **any** change to `forms/app/*.xml`.

> **Once anyone publishes from the forms panel, the database has the real
> forms.** Run `python scripts/pull_live_forms.py --url ...` before
> `upload-app-forms`, or you will overwrite live forms with old copies. For
> day-to-day form changes use the panel (section 5).

## 4. Build the webapp (first time, and whenever cht-core changes)

```bash
cd ../cht-core
npx sass webapp/src/css/enketo/enketo.scss api/build/static/webapp/enketo.less --no-source-map
cd webapp && npm ci && npm run build -- --configuration=production
cd ../../satoruemr
satoru build api && satoru up -d api      # bakes the build into satoru-cht-api:5.3.0
```

> **`--configuration=production` is mandatory.** `angular.json` sets
> `defaultConfiguration: None`, so a bare `ng build` inlines "critical CSS"
> into `index.html`. That inlined copy of `.bootstrap-layer{display:flex}`
> is unlayered and overrides `.bootstrapped .bootstrap-layer{display:none}`,
> leaving the app stuck behind a loading spinner with the real UI rendered
> underneath. Production sets `inlineCritical:false`. Sanity check:
> `index.html` should be **~850 bytes**, not ~56 KB.
>
> The enketo.less step is also required — `ng build` fails without it.

## 5. Excel exports

Browser: **https://\<host\>:10443/exports/** — same stack, same URL as CHT.

Login: you must be logged in to CHT as an **admin**. Not logged in → sent to
the CHT login page; logged in but not admin → "only administrators". Every
download and every forms-panel action is logged with the username
(`docker logs compose-exports-1`).
Command line and format details: [scripts/EXPORTS.md](scripts/EXPORTS.md).

### Forms panel — https://\<host\>:10443/exports/forms/

Upload a form as an Excel XLSForm; it goes live only if every check passes:

1. it is an .xlsx and converts to a CHT form (cht-conf)
2. the form id inside matches, and CHT's own validation passes
3. no answer is overwritten by a formula; no yes/no text field is used as a condition
4. DDST-II, DST, MoCA, VSMS: result fields still exist and **the scoring tests pass**
   (other forms can be published with a "scoring untested" warning)

The review page shows every score formula that changed. **Publish** saves the
old version to the `satoru-forms-history` database first, keeps the form's
title, icon and images, then checks the upload landed. **History** lets you
download any version or re-check an old one to go back.

**Manage / history** (link next to each form):

- **Archive** — nobody can start new reports with the form; its saved reports
  still show and export. **Unarchive** undoes it. Uploading a new version of
  an archived form keeps it archived.
- **Delete** — only when the form has **no saved reports**, and you must type
  the form id. A copy is kept in history; to bring it back, use "Deleted
  form?" on the Forms page → "Check the version it replaced" → Publish.

Forms published from the panel get readable report labels automatically
(question text, or the field name made readable).

Exports always read the forms that are live in the database.

## 5b. Users and participant privacy (interns vs supervisors)

Roles are `intern` and `supervisor` (both work offline). Permissions live in
`app_settings/base_settings.json` → `permissions`.

| | Intern | Supervisor |
|---|---|---|
| Participants they registered, for `participant_visibility_hours` (default 24) | ✅ | ✅ |
| Everyone else's participants and reports | only after an **exact** search | ✅ always |
| Excel exports / forms panel | ❌ (CHT admins only) | ❌ (CHT admins only) |

- After registration an intern sees that participant (and their reports) for
  24 hours, then they disappear from the lists.
- To open anyone again, the intern uses **Search Participant** at the top of
  People: exact CR number, 10-digit phone, or exact full name. Matches are
  unlocked for that browser session only.
- New participants record `created_by_person_uuid`, `created_by_name`,
  `created_by_user`. Participants created **before** this feature have no
  creator, so interns only reach them by exact search.
- Change the window with `participant_visibility_hours` in
  `app_settings/base_settings.json`, then `cht compile-app-settings`,
  `python scripts/gen_report_config.py`, and upload settings.
- A login must be linked to a person (contact) to fill forms. Give interns a
  home camp; don't give the `medic` admin one (see PRODUCTION.md).
- **Limit:** this hides data in the app screens. An intern's phone still stores
  their camp's data, so a technical user could read it from the device.

## 6. Test data / self-checks

```bash
.venv/bin/python scripts/load_sample_data.py --url https://medic:PASSWORD@<HOST>:10443/medic
.venv/bin/python scripts/validate_exports.py        # 58 checks, no server needed
.venv/bin/python scripts/validate_export_login.py   # 44 checks, login + POST protection on /exports/
.venv/bin/python scripts/validate_forms_panel.py    # forms panel helpers
node scripts/validate_contact_summary.js            # participant access gate in the compiled contact summary
node scripts/validate_report_patch.js               # 50 checks, needs ../cht-core/node_modules/typescript
node scripts/validate_exact_search.js               # 27 checks

# scoring tests in CHT 5.3's real form engine (needs Chromium)
cd tests/forms && npm ci && CHROMIUM_PATH=/usr/bin/chromium npm test   # 64 passing, 1 pending (VSMS table)
```

Scoring problems found and fixed: [SCORING_AUDIT.md](SCORING_AUDIT.md).
DDST-II reports saved before 2026-09-17 all say "Suspect" — fix them with
`scripts/rescore_ddst_overall.py` (dry run first, `--apply` backs up then fixes).

`tests/forms/engine/cht-core-5.3.0/` is CHT's form engine built from our
cht-core (`cd webapp && npx ng build cht-form`, then copy
`build/cht-form/{main,polyfills,runtime,scripts}.js` and `styles.css`).
Rebuild it whenever cht-core's form code changes.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Reports show `report.dst.g_child.assessment_date` | config/translations not uploaded | step 3, then hard-refresh |
| App stuck on spinner, Reports never load | built without `--configuration=production`, or running 5.2.0 images | step 4 / step 2 |
| `permission denied ... docker.sock` | user not in `docker` group | `sudo usermod -aG docker $USER`, re-login |
| cht-conf exits on "overwrite?" | no TTY | add `--force` |
| Export sheet missing a scale | nobody completed it in that camp | expected — empty scales are skipped |
