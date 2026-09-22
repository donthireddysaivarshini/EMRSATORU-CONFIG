"""forms_panel.py - upload an XLSForm in the browser and publish it safely.

Lives inside the exports service at /exports/forms/ (same login: CHT admins).

    upload .xlsx  ->  convert (cht-conf)  ->  cht-conf validation
                  ->  form safety checks  ->  scoring tests (known scales)
                  ->  review screen        ->  Publish
                  ->  save previous version to history  ->  upload  ->  verify

Nothing is published unless every blocking check passed. Every publish keeps
the version it replaced in the `satoru-forms-history` database, and any
history entry can be re-published (it goes through the same checks).

The live database is the source of truth for app forms: after a publish the
repo's forms/app/*.xml is out of date until someone runs
scripts/pull_live_forms.py.
"""

import base64
import hashlib
import hmac
import html
import json
import os
import re
import secrets
import shutil
import subprocess
import threading
import time
import urllib.parse
from datetime import datetime, timezone
from email.parser import BytesParser
from email.policy import default as email_policy

import requests

import scale_meta

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
TESTS_DIR = os.environ.get('FORM_TESTS_DIR', os.path.join(REPO, 'tests', 'forms'))
CHT_BIN = os.environ.get('CHT_BIN', 'cht')
STAGING_ROOT = os.environ.get('FORMS_STAGING_DIR', '/tmp/satoru-forms-staging')
HISTORY_DB = 'satoru-forms-history'

SCALE_SPECS = {
    'ddst': 'ddst.spec.js',
    'dst': 'dst.spec.js',
    'moca_assessment': 'moca.spec.js',
    'vineland': 'vineland.spec.js',
}
FORM_ID_RE = re.compile(r'^[a-z][a-z0-9_]{1,62}$')
TOKEN_RE = re.compile(r'^[A-Za-z0-9_-]{20,64}$')
HISTORY_ID_RE = re.compile(r'^[a-z0-9_]+:[0-9TZ:.-]+:[a-f0-9]{8}$')
MAX_UPLOAD_BYTES = 15 * 1024 * 1024
STAGING_TTL_SECONDS = 2 * 60 * 60
GENERATED_ATTACHMENTS = {'xml', 'form.html', 'model.xml'}
PROPERTY_KEYS = ('title', 'context', 'icon', 'subject_key', 'hidden_fields', 'xml2sms')

# one pipeline at a time: the tests start a browser, and publishes must not race
PIPELINE_LOCK = threading.Lock()


class PanelError(Exception):
    """A problem to show the admin in plain words."""


# ---------------------------------------------------------------------------
# CouchDB
# ---------------------------------------------------------------------------

class Couch:
    """Service-account access to CouchDB (COUCH_URL = .../medic)."""

    def __init__(self, medic_url):
        parts = urllib.parse.urlsplit(medic_url)
        self.session = requests.Session()
        self.session.verify = False
        if parts.username:
            self.session.auth = (urllib.parse.unquote(parts.username),
                                 urllib.parse.unquote(parts.password or ''))
        host = parts.hostname + (f':{parts.port}' if parts.port else '')
        self.server = f'{parts.scheme}://{host}'
        self.medic = f'{self.server}{parts.path.rstrip("/")}'

    def _check(self, resp, what):
        if resp.status_code >= 400:
            raise PanelError(f'{what} failed (HTTP {resp.status_code}): {resp.text[:200]}')
        return resp

    def get_doc(self, doc_id, db=None, **params):
        url = f'{db or self.medic}/{urllib.parse.quote(doc_id, safe="")}'
        resp = self.session.get(url, params=params, timeout=60)
        if resp.status_code == 404:
            return None
        return self._check(resp, f'Reading {doc_id}').json()

    def get_attachment(self, doc_id, name):
        url = (f'{self.medic}/{urllib.parse.quote(doc_id, safe="")}/'
               f'{urllib.parse.quote(name, safe="/")}')
        return self._check(self.session.get(url, timeout=60), f'Reading {doc_id}/{name}').content

    def put_doc(self, doc, db=None):
        url = f'{db or self.medic}/{urllib.parse.quote(doc["_id"], safe="")}'
        return self._check(self.session.put(url, json=doc, timeout=120), f'Saving {doc["_id"]}').json()

    def app_forms(self):
        resp = self.session.get(f'{self.medic}/_all_docs', timeout=60, params={
            'startkey': json.dumps('form:'), 'endkey': json.dumps('form:￰'), 'include_docs': 'true'})
        docs = [row['doc'] for row in self._check(resp, 'Listing forms').json()['rows'] if row.get('doc')]
        return [d for d in docs if not d['_id'].startswith('form:contact:')]

    def report_counts(self):
        resp = self.session.get(f'{self.medic}/_design/medic-client/_view/reports_by_form',
                                params={'group': 'true'}, timeout=120)
        rows = self._check(resp, 'Counting reports').json()['rows']
        return {row['key'][0]: row['value']['count'] if isinstance(row['value'], dict) else row['value']
                for row in rows}

    def history_db(self):
        url = f'{self.server}/{HISTORY_DB}'
        if self.session.head(url, timeout=30).status_code == 404:
            self._check(self.session.put(url, timeout=30), 'Creating the history database')
        return url

    def history(self, form_id):
        url = self.history_db()
        resp = self.session.get(f'{url}/_all_docs', timeout=60, params={
            'startkey': json.dumps(f'{form_id}:￰'), 'endkey': json.dumps(f'{form_id}:'),
            'descending': 'true', 'include_docs': 'true'})
        return [row['doc'] for row in self._check(resp, 'Reading history').json()['rows']]


# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------

def now_iso():
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.%fZ')


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def run(cmd, cwd, env=None, timeout=300):
    try:
        proc = subprocess.run(cmd, cwd=cwd, env={**os.environ, **(env or {})}, capture_output=True,
                              text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        return 124, f'timed out after {timeout}s'
    except FileNotFoundError as exc:
        return 127, f'not installed: {exc}'
    output = (proc.stdout + proc.stderr)
    output = re.sub(r'\x1b\[[0-9;]*m', '', output)
    return proc.returncode, output


def cht_conf_errors(output):
    """The ERROR lines of cht-conf output, or its last lines."""
    errors = [line for line in output.splitlines() if line.startswith(('ERROR', 'WARN'))]
    return '\n'.join(errors)[-3000:] if errors else output[-3000:]


def test_summary(output):
    """mocha output without stack traces: passes, failures and why."""
    keep = re.compile(r'✔|^\s+\d+\) |passing|failing|pending|AssertionError|expected|^\s+[-+]\s|Error:')
    lines = [line for line in output.splitlines() if keep.search(line) and not line.strip().startswith('at ')]
    return '\n'.join(lines)[-8000:]


def form_xml_id(xml_text):
    match = re.search(r'<instance>\s*<([A-Za-z_][\w.-]*)[^>]*\sid="([^"]+)"', xml_text)
    return match.group(2) if match else None


def form_version(xml_text):
    match = re.search(r'<instance>\s*<[^>]*\sversion="([^"]*)"', xml_text)
    return match.group(1) if match else ''


def calculations(xml_text):
    return {m.group(1): html.unescape(m.group(2))
            for m in re.finditer(r'<bind nodeset="([^"]+)"[^>]*\scalculate="([^"]*)"', xml_text)}


def fields(xml_text):
    return set(re.findall(r'<bind nodeset="([^"]+)"', xml_text))


MEDIA_FOLDERS = {'images': 'images/', 'audio': 'audio/', 'video': 'video/', 'file': '', 'file-csv': ''}


def missing_media(xml_text, media_dir):
    """jr:// media the form references that would not be attached."""
    missing = []
    for kind, name in sorted(set(re.findall(r'jr://([a-z-]+)/([^"<\s]+)', xml_text))):
        rel = MEDIA_FOLDERS.get(kind, '') + html.unescape(name)
        if not os.path.isfile(os.path.join(media_dir, rel)):
            missing.append(rel)
    return missing


def csrf_token(secret, auth_cookie):
    return hmac.new(secret, (auth_cookie or '').encode(), hashlib.sha256).hexdigest()


def parse_multipart(content_type, body):
    """{name: str} and {name: (filename, bytes)} from a multipart/form-data body."""
    message = BytesParser(policy=email_policy).parsebytes(
        b'Content-Type: ' + content_type.encode('latin-1') + b'\r\n\r\n' + body)
    if not message.is_multipart():
        raise PanelError('The upload was not sent as a file form.')
    values, files = {}, {}
    for part in message.iter_parts():
        name = part.get_param('name', header='content-disposition')
        if not name:
            continue
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b''
        if filename is not None:
            files[name] = (os.path.basename(filename), payload)
        else:
            values[name] = payload.decode('utf-8', 'replace')
    return values, files


# ---------------------------------------------------------------------------
# staging
# ---------------------------------------------------------------------------

def clean_staging():
    if not os.path.isdir(STAGING_ROOT):
        return
    for name in os.listdir(STAGING_ROOT):
        path = os.path.join(STAGING_ROOT, name)
        if time.time() - os.path.getmtime(path) > STAGING_TTL_SECONDS:
            shutil.rmtree(path, ignore_errors=True)


def new_staging():
    clean_staging()
    token = secrets.token_urlsafe(24)
    path = os.path.join(STAGING_ROOT, token)
    os.makedirs(STAGING_ROOT, mode=0o700, exist_ok=True)
    os.mkdir(path, mode=0o700)
    os.makedirs(os.path.join(path, 'forms', 'app'))
    return token, path


def staging_path(token):
    if not TOKEN_RE.match(token or ''):
        raise PanelError('That check has expired. Upload the file again.')
    path = os.path.join(STAGING_ROOT, token)
    if not os.path.isfile(os.path.join(path, 'result.json')):
        raise PanelError('That check has expired. Upload the file again.')
    return path


def load_result(path):
    with open(os.path.join(path, 'result.json'), encoding='utf-8') as f:
        return json.load(f)


def save_result(path, result):
    with open(os.path.join(path, 'result.json'), 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=1)


# ---------------------------------------------------------------------------
# checks
# ---------------------------------------------------------------------------

def stage_live_media(couch, live_doc, media_dir):
    """Keep the live form's images unless new ones were uploaded."""
    for name, info in (live_doc or {}).get('_attachments', {}).items():
        if name in GENERATED_ATTACHMENTS:
            continue
        target = os.path.normpath(os.path.join(media_dir, name))
        if not target.startswith(os.path.normpath(media_dir) + os.sep):
            continue
        os.makedirs(os.path.dirname(target), exist_ok=True)
        with open(target, 'wb') as f:
            f.write(couch.get_attachment(live_doc['_id'], name))


def check_form(couch, path, form_id, user, source, xlsx=None, xml=None, new_properties=None):
    """Run every check on a staged form. Writes and returns result.json."""
    app_dir = os.path.join(path, 'forms', 'app')
    checks = []

    def add(name, ok, detail='', blocking=True):
        checks.append({'name': name, 'ok': bool(ok), 'detail': '' if ok else detail, 'blocking': blocking})
        return ok

    live_doc = couch.get_doc(f'form:{form_id}', attachments='false')
    live_xml = couch.get_attachment(live_doc['_id'], 'xml').decode('utf-8') if live_doc else ''

    if xlsx is not None:
        add('File is an Excel spreadsheet (.xlsx)', xlsx[:2] == b'PK',
            'XLSForms must be saved as .xlsx')
        with open(os.path.join(app_dir, f'{form_id}.xlsx'), 'wb') as f:
            f.write(xlsx)
        if checks[-1]['ok']:
            code, out = run([CHT_BIN, '--skip-dependency-check', 'convert-app-forms', '--', form_id], path)
            add('Spreadsheet converts to a CHT form', code == 0 and os.path.isfile(os.path.join(app_dir, f'{form_id}.xml')),
                cht_conf_errors(out))
    else:
        with open(os.path.join(app_dir, f'{form_id}.xml'), 'wb') as f:
            f.write(xml)

    xml_path = os.path.join(app_dir, f'{form_id}.xml')
    new_xml = ''
    if os.path.isfile(xml_path) and all(c['ok'] for c in checks):
        with open(xml_path, encoding='utf-8') as f:
            new_xml = f.read()
        found_id = form_xml_id(new_xml)
        add(f'Form id inside the file is "{form_id}"', found_id == form_id,
            f'The spreadsheet settings sheet says "{found_id}". Fix form_id or rename the file.')

    properties = {k: live_doc[k] for k in PROPERTY_KEYS if live_doc and k in live_doc}
    properties.update(new_properties or {})
    with open(os.path.join(app_dir, f'{form_id}.properties.json'), 'w', encoding='utf-8') as f:
        json.dump(properties, f)
    if live_doc:
        stage_live_media(couch, live_doc, os.path.join(app_dir, f'{form_id}-media'))

    if new_xml:
        missing = missing_media(new_xml, os.path.join(app_dir, f'{form_id}-media'))
        add('Every picture / file the form uses is on the server', not missing,
            'Not found: ' + ', '.join(missing) + '. Uploading pictures from this page is not supported yet: '
            'use pictures the live form already has, or ask the admin to upload them.')

    blocked_early = not all(c['ok'] for c in checks if c['blocking'])
    test_output = ''
    if not blocked_early:
        code, out = run([CHT_BIN, '--skip-dependency-check', 'validate-app-forms', '--', form_id], path)
        add('CHT form validation', code == 0 and 'ERROR' not in out, cht_conf_errors(out))

        mocha = os.path.join(TESTS_DIR, 'node_modules', '.bin', 'mocha')
        env = {'FORMS_DIR': app_dir, 'FORM_IDS': form_id}
        code, out = run([mocha, 'static.spec.js', '--timeout', '60000'], TESTS_DIR, env, timeout=300)
        add('Answers are never overwritten by formulas; yes/no flags used safely', code == 0, test_summary(out))

        spec = SCALE_SPECS.get(form_id)
        if spec:
            for field in scale_meta.RESULT_FIELDS.get(form_id, []):
                if not re.search(r'nodeset="[^"]*/%s"' % re.escape(field), new_xml):
                    add(f'Result field "{field}" still exists', False,
                        'Reports, exports and the printout need this field.')
            code, out = run([mocha, spec, '--timeout', '120000'], TESTS_DIR, env, timeout=600)
            test_output = test_summary(out)
            add(f'Scoring tests ({spec})', code == 0, test_output)
        else:
            add('Scoring tests', False, 'This form has no scoring tests yet. Its scores have NOT been checked.',
                blocking=False)

    old_calc, new_calc = calculations(live_xml), calculations(new_xml)
    result = {
        'form_id': form_id,
        'source': source,
        'user': user,
        'created': now_iso(),
        'is_new_form': live_doc is None,
        'live_rev': live_doc['_rev'] if live_doc else None,
        'live_version': form_version(live_xml),
        'new_version': form_version(new_xml),
        'xlsx_sha256': sha256(xlsx) if xlsx is not None else None,
        'xml_sha256': sha256(new_xml.encode('utf-8')) if new_xml else None,
        'checks': checks,
        'blocked': not new_xml or any(not c['ok'] and c['blocking'] for c in checks),
        'untested': form_id not in SCALE_SPECS,
        'formulas': {
            'added': sorted(set(new_calc) - set(old_calc)),
            'removed': sorted(set(old_calc) - set(new_calc)),
            'changed': [{'field': k, 'old': old_calc[k], 'new': new_calc[k]}
                        for k in sorted(set(old_calc) & set(new_calc)) if old_calc[k] != new_calc[k]],
        },
        'removed_fields': sorted(fields(live_xml) - fields(new_xml)) if live_xml and new_xml else [],
        'reports_for_form': couch.report_counts().get(form_id, 0),
        'test_output': test_output,
        'properties': properties,
    }
    if live_doc and result['live_version'] == result['new_version']:
        checks.append({'name': 'Version number changed', 'ok': False, 'blocking': False,
                       'detail': 'The version in the settings sheet is the same as the live form. '
                                 'Change it so reports show which version they came from.'})
    save_result(path, result)
    return result


# ---------------------------------------------------------------------------
# publish
# ---------------------------------------------------------------------------

COMMON_FIELD_LABELS = {'patient_uuid': 'Participant', 'patient_id': 'Participant ID', 'patient_name': 'Participant name',
                       'screened_by': 'Submitted by (user ID)', 'screened_by_name': 'Submitted by'}


def generic_labels(form_id, xml_path):
    """report.<form>.<field> labels for any form: its question text, or the
    field name made readable (patient_name -> Patient name)."""
    with open(xml_path, encoding='utf-8') as f:
        xml_text = f.read()
    try:
        flat = scale_meta.flatten_fields(scale_meta.parse_form_xml(xml_path))
        question_labels = {item['path']: ' '.join((item['label'] or '').split()) for item in flat}
    except Exception:                                 # noqa: BLE001 - fall back to field names
        question_labels = {}
    wanted = []
    for nodeset in sorted(fields(xml_text)):
        parts = nodeset.strip('/').split('/')[1:]    # drop the /data root
        if not parts or parts[0] == 'inputs' or parts[-1] in ('meta', 'instanceID', 'deprecatedID'):
            continue
        for depth in range(1, len(parts) + 1):        # groups too
            key_path = '.'.join(parts[:depth])
            name = parts[depth - 1]
            label = (question_labels.get(key_path) or COMMON_FIELD_LABELS.get(name)
                     or name.replace('_', ' ').strip().capitalize())
            wanted.append((f'report.{form_id}.{key_path}', label))
    seen, unique = set(), []
    for key, label in wanted:
        if key not in seen:
            seen.add(key)
            unique.append((key, label))
    return unique


def add_missing_labels(couch, wanted):
    translations = couch.get_doc('messages-en')
    existing = set((translations.get('custom') or {}).keys()) | set((translations.get('generic') or {}).keys())
    missing = {key: label for key, label in wanted if key not in existing}
    if missing:
        translations.setdefault('custom', {}).update(missing)
        couch.put_doc(translations)
    return len(missing)


def update_report_config(couch, path, form_id):
    """Labels for the new version; known scales also get their report layout refreshed."""
    if form_id not in scale_meta.SCALES:
        added = add_missing_labels(couch, generic_labels(form_id, os.path.join(path, 'forms', 'app', f'{form_id}.xml')))
        return [f'{added} new label(s) added'] if added else []
    merged = os.path.join(path, 'all-scales')
    os.makedirs(merged, exist_ok=True)
    for form, info in scale_meta.SCALES.items():
        target = os.path.join(merged, info['file'])
        if form == form_id:
            shutil.copyfile(os.path.join(path, 'forms', 'app', f'{form_id}.xml'), target)
        else:
            with open(target, 'wb') as f:
                f.write(couch.get_attachment(f'form:{form}', 'xml'))

    import gen_report_config
    with scale_meta.LOCK:
        previous_dir = scale_meta.FORMS_DIR
        try:
            scale_meta.set_forms_dir(merged)
            report_display = scale_meta.build_report_display_config()
            wanted = gen_report_config.desired_keys_for(form_id)
        finally:
            scale_meta.set_forms_dir(previous_dir)

    notes = []
    settings = couch.get_doc('settings')
    if settings['settings'].get('report_display') != report_display:
        settings['settings']['report_display'] = report_display
        couch.put_doc(settings)
        notes.append('report layout updated')

    added = add_missing_labels(couch, wanted)
    if added:
        notes.append(f'{added} new label(s) added')
    return notes


def publish(couch, token, user):
    path = staging_path(token)
    with PIPELINE_LOCK:
        result = load_result(path)
        if result.get('published'):
            raise PanelError('This version was already published.')
        if result['blocked']:
            raise PanelError('Some checks failed. Nothing was published.')
        if result['user'] != user:
            raise PanelError('Only the person who ran the checks can publish them.')
        form_id = result['form_id']
        app_dir = os.path.join(path, 'forms', 'app')
        with open(os.path.join(app_dir, f'{form_id}.xml'), 'rb') as f:
            new_xml = f.read()
        if sha256(new_xml) != result['xml_sha256']:
            raise PanelError('The checked file changed. Upload it again.')

        live_doc = couch.get_doc(f'form:{form_id}')
        if (live_doc['_rev'] if live_doc else None) != result['live_rev']:
            raise PanelError('Someone changed this form after the checks ran. Upload it again so the checks re-run.')

        # 1. history first - if this fails, nothing is published
        history_url = couch.history_db()
        entry_id = f'{form_id}:{now_iso()}:{secrets.token_hex(4)}'
        attachments = {'new.xml': {'content_type': 'application/xml', 'data': base64.b64encode(new_xml).decode()}}
        xlsx_path = os.path.join(app_dir, f'{form_id}.xlsx')
        if os.path.isfile(xlsx_path):
            with open(xlsx_path, 'rb') as f:
                attachments['new.xlsx'] = {
                    'content_type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'data': base64.b64encode(f.read()).decode()}
        if live_doc:
            attachments['previous.xml'] = {'content_type': 'application/xml', 'data': base64.b64encode(
                couch.get_attachment(live_doc['_id'], 'xml')).decode()}
        entry = {
            '_id': entry_id, 'type': 'form_publish', 'form_id': form_id, 'status': 'saving',
            'published_by': user, 'started_at': now_iso(), 'source': result['source'],
            'previous_version': result['live_version'], 'new_version': result['new_version'],
            'xml_sha256': result['xml_sha256'], 'xlsx_sha256': result['xlsx_sha256'],
            'checks': result['checks'], 'test_output': result['test_output'],
            'previous_properties': {k: live_doc[k] for k in PROPERTY_KEYS if live_doc and k in live_doc},
            '_attachments': attachments,
        }
        saved = couch.put_doc(entry, db=history_url)

        # 2. upload the form, the same document cht-conf upload-app-forms writes
        doc = {
            '_id': f'form:{form_id}', 'type': 'form', 'internalId': form_id,
            'xmlVersion': {'time': int(time.time() * 1000), 'sha256': result['xml_sha256']},
            '_attachments': {},
        }
        doc.update(result['properties'])
        if live_doc:
            doc['_rev'] = live_doc['_rev']
            if is_archived(live_doc):       # a new version of an archived form stays archived
                doc['satoru_archived'] = live_doc['satoru_archived']
                doc['context'] = ARCHIVED_CONTEXT
        if 'title' not in doc:
            match = re.search(rb'<h:title>([^<]*)</h:title>', new_xml)
            doc['title'] = html.unescape(match.group(1).decode()) if match else form_id
        media_dir = os.path.join(app_dir, f'{form_id}-media')
        for root, _, names in os.walk(media_dir):
            for name in names:
                full = os.path.join(root, name)
                rel = os.path.relpath(full, media_dir).replace(os.sep, '/')
                with open(full, 'rb') as f:
                    content = f.read()
                content_type = live_doc['_attachments'][rel]['content_type'] if live_doc and rel in live_doc.get(
                    '_attachments', {}) else 'application/octet-stream'
                doc['_attachments'][rel] = {'content_type': content_type, 'data': base64.b64encode(content).decode()}
        doc['_attachments']['xml'] = {'content_type': 'application/xml', 'data': base64.b64encode(new_xml).decode()}
        try:
            couch.put_doc(doc)
        except PanelError as exc:
            entry.update({'_rev': saved['rev'], 'status': 'failed', 'error': str(exc), 'finished_at': now_iso()})
            entry.pop('_attachments')
            couch.put_doc(entry, db=history_url)
            raise

        # 3. verify it is live, then layout + labels
        live = couch.get_doc(f'form:{form_id}')
        live_ok = live and sha256(couch.get_attachment(live['_id'], 'xml')) == result['xml_sha256']
        notes = []
        if live_ok:
            try:
                notes = update_report_config(couch, path, form_id)
            except Exception as exc:                          # noqa: BLE001
                notes = [f'FORM IS LIVE but report layout/labels were not updated: {exc}']

        entry = couch.get_doc(entry_id, db=history_url)
        entry.update({'status': 'published' if live_ok else 'uploaded_but_not_verified',
                      'finished_at': now_iso(), 'notes': notes})
        couch.put_doc(entry, db=history_url)
        result['published'] = entry_id
        save_result(path, result)
        if not live_ok:
            raise PanelError('The form was uploaded but could not be read back. Check it in CHT before using it.')
        return form_id, notes


# ---------------------------------------------------------------------------
# archive / delete
# ---------------------------------------------------------------------------

ARCHIVED_CONTEXT = {'person': False, 'place': False, 'expression': 'false'}


def is_archived(doc):
    return 'satoru_archived' in doc


def save_history_copy(couch, live_doc, user, action):
    """Keep the form as it is now, so archive/delete can be undone."""
    form_id = live_doc['internalId']
    entry = {
        '_id': f'{form_id}:{now_iso()}:{secrets.token_hex(4)}', 'type': 'form_publish', 'form_id': form_id,
        'status': action, 'published_by': user, 'started_at': now_iso(), 'finished_at': now_iso(),
        'source': action, 'previous_version': form_version(couch.get_attachment(live_doc['_id'], 'xml').decode('utf-8')),
        'new_version': '', 'checks': [], 'test_output': '',
        'previous_properties': {k: live_doc[k] for k in PROPERTY_KEYS if k in live_doc},
        'previous_media': sorted(n for n in live_doc.get('_attachments', {}) if n not in GENERATED_ATTACHMENTS),
        '_attachments': {'previous.xml': {'content_type': 'application/xml', 'data': base64.b64encode(
            couch.get_attachment(live_doc['_id'], 'xml')).decode()}},
    }
    if 'satoru_archived' in live_doc:
        entry['previous_properties']['context'] = live_doc['satoru_archived'].get('context')
    couch.put_doc(entry, db=couch.history_db())


def set_archived(couch, form_id, user, archive):
    with PIPELINE_LOCK:
        doc = couch.get_doc(f'form:{form_id}')
        if not doc:
            raise PanelError(f'There is no live form "{form_id}".')
        if archive == is_archived(doc):
            raise PanelError(f'"{form_id}" is already {"archived" if archive else "active"}.')
        save_history_copy(couch, doc, user, 'archived' if archive else 'unarchived')
        if archive:
            doc['satoru_archived'] = {'context': doc.get('context'), 'by': user, 'at': now_iso()}
            doc['context'] = ARCHIVED_CONTEXT
        else:
            original = doc.pop('satoru_archived').get('context')
            if original is None:
                doc.pop('context', None)
            else:
                doc['context'] = original
        couch.put_doc(doc)


def delete_form(couch, form_id, user, confirm):
    with PIPELINE_LOCK:
        doc = couch.get_doc(f'form:{form_id}')
        if not doc:
            raise PanelError(f'There is no live form "{form_id}".')
        if confirm.strip() != form_id:
            raise PanelError(f'To delete, type the form id exactly: {form_id}')
        reports = couch.report_counts().get(form_id, 0)
        if reports:
            raise PanelError(f'"{form_id}" has {reports} saved report(s). Deleting it would leave them without their form. '
                             f'Archive it instead: nobody can start new ones and the old reports still show.')
        save_history_copy(couch, doc, user, 'deleted')
        resp = couch.session.delete(f'{couch.medic}/{urllib.parse.quote(doc["_id"], safe="")}',
                                    params={'rev': doc['_rev']}, timeout=60)
        couch._check(resp, f'Deleting {doc["_id"]}')


# ---------------------------------------------------------------------------
# pages
# ---------------------------------------------------------------------------

STYLE = """<style>
 body{font:15px/1.5 system-ui,sans-serif;margin:0;background:#f2f2f3;color:#1a1a1a}
 .wrap{max-width:900px;margin:0 auto;padding:32px 20px 64px}
 h1{font-size:22px;margin:0 0 4px} h2{font-size:16px;margin:0 0 8px}
 .sub{color:#666;font-size:13px;margin-bottom:24px}
 .nav a{margin-right:16px}
 .card{background:#fff;border:1px solid #e2e2e5;border-radius:8px;padding:20px;margin-bottom:16px;overflow-x:auto}
 table{border-collapse:collapse;width:100%;font-size:14px} td,th{padding:6px 8px;border-bottom:1px solid #eee;text-align:left;vertical-align:top}
 button{font:inherit;padding:9px 16px;border:0;border-radius:5px;background:#0a5ca8;color:#fff;cursor:pointer}
 button[disabled]{background:#999;cursor:not-allowed}
 input,select{font:inherit;padding:6px;margin:4px 0 10px;max-width:100%}
 .ok{color:#0a7a2f;font-weight:600} .bad{color:#b00;font-weight:600} .warn{color:#a15c00;font-weight:600}
 .banner{padding:12px 16px;border-radius:6px;margin-bottom:16px}
 .banner.bad{background:#fde8e8} .banner.warn{background:#fff4dd} .banner.ok{background:#e6f5ea}
 pre{white-space:pre-wrap;background:#f6f6f6;padding:8px;font-size:12px;max-height:320px;overflow:auto}
 code{background:#eee;padding:1px 4px;border-radius:3px;font-size:12px;word-break:break-all}
 label{display:block;font-weight:600;font-size:14px}
</style>"""


def page(title, body):
    return (f'<!doctype html><html><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width,initial-scale=1">'
            f'<title>{html.escape(title)}</title>{STYLE}</head><body><div class="wrap">'
            f'<div class="nav sub"><a href="../">Excel exports</a><a href="./">Forms</a></div>'
            f'{body}</div></body></html>')


def e(value):
    return html.escape(str(value if value is not None else ''))


def title_text(doc):
    title = doc.get('title')
    if isinstance(title, list):
        title = next((t.get('content') for t in title if t.get('locale') == 'en'), None) or \
            (title[0].get('content') if title else '')
    return title or doc.get('internalId') or doc['_id']


def forms_page(couch, csrf, message=''):
    docs = sorted(couch.app_forms(), key=lambda d: d['_id'])
    counts = couch.report_counts()
    rows = ''.join(
        f'<tr><td><b>{e(title_text(d))}</b><br><code>{e(d.get("internalId"))}</code></td>'
        f'<td>{e(counts.get(d.get("internalId"), 0))}</td>'
        f'<td>{"<span class=ok>yes</span>" if d.get("internalId") in SCALE_SPECS else "<span class=warn>no</span>"}</td>'
        f'<td>{"<span class=warn>archived</span>" if is_archived(d) else "active"}</td>'
        f'<td><a href="history?form={urllib.parse.quote(d.get("internalId", ""))}">manage / history</a></td></tr>'
        for d in docs)
    options = ''.join(f'<option value="{e(d.get("internalId"))}">{e(title_text(d))} ({e(d.get("internalId"))})</option>'
                      for d in docs)
    body = f"""
<h1>Forms</h1>
<div class="sub">Upload a new version of a form as an Excel XLSForm. It is checked before anything goes live.</div>
{message}
<div class="card"><h2>Upload a form</h2>
<form method="post" action="check" enctype="multipart/form-data">
 <input type="hidden" name="csrf" value="{e(csrf)}">
 <label>Which form</label>
 <select name="form_id"><option value="">A new form (id from the file name)</option>{options}</select>
 <label>Spreadsheet (.xlsx)</label>
 <input type="file" name="xlsx" accept=".xlsx" required>
 <label>New form only: title</label>
 <input type="text" name="title" maxlength="120" placeholder="e.g. Hearing screening">
 <label>New form only: filled in for</label>
 <select name="for"><option value="person">a participant</option><option value="place">a camp</option></select>
 <div><button type="submit">Check this file</button></div>
 <div class="sub">Checks take up to a minute for forms with scoring tests.</div>
</form></div>
<div class="card"><h2>Deleted form?</h2>
<form method="get" action="history"><input type="text" name="form" placeholder="form id, e.g. test_form" required>
 <button type="submit">Open its history</button></form></div>
<div class="card"><h2>Live forms</h2>
<table><tr><th>Form</th><th>Reports</th><th>Scoring tests</th><th>Status</th><th></th></tr>{rows}</table></div>"""
    return page('Forms', body)


def check_row(check):
    if check['ok']:
        status = '<span class="ok">PASS</span>'
    elif check['blocking']:
        status = '<span class="bad">FAIL</span>'
    else:
        status = '<span class="warn">WARNING</span>'
    detail = f'<pre>{e(check["detail"])}</pre>' if check['detail'] else ''
    return f'<tr><td>{status}</td><td>{e(check["name"])}{detail}</td></tr>'


def review_page(result, token, csrf):
    checks = ''.join(check_row(c) for c in result['checks'])
    formulas = result['formulas']
    changed = ''.join(f'<tr><td><code>{e(c["field"])}</code></td><td><code>{e(c["old"])}</code></td>'
                      f'<td><code>{e(c["new"])}</code></td></tr>' for c in formulas['changed'])
    listing = lambda items: ', '.join(f'<code>{e(i)}</code>' for i in items) or 'none'  # noqa: E731
    if result.get('published'):
        banner = '<div class="banner ok">Published.</div>'
    elif result['blocked']:
        banner = '<div class="banner bad"><b>Cannot publish.</b> Fix the failed checks and upload again.</div>'
    elif result['untested']:
        banner = ('<div class="banner warn"><b>Scoring untested.</b> This form has no scoring tests, '
                  'so nobody has checked that its scores are right.</div>')
    else:
        banner = '<div class="banner ok">All checks passed.</div>'
    removed = ''
    if result['removed_fields']:
        removed = (f'<div class="banner warn">Fields in the live form that are gone from this version: '
                   f'{listing(result["removed_fields"])}. {e(result["reports_for_form"])} stored report(s) use this form; '
                   f'they keep their data but new reports will not have these fields.</div>')
    button = '' if result['blocked'] or result.get('published') else f"""
<form method="post" action="publish">
 <input type="hidden" name="csrf" value="{e(csrf)}"><input type="hidden" name="token" value="{e(token)}">
 <button type="submit">Publish {e(result["form_id"])} now</button>
</form>"""
    body = f"""
<h1>Check: {e(result['form_id'])}</h1>
<div class="sub">{"New form" if result["is_new_form"] else f"Live version {e(result['live_version'])}"} &rarr;
 new version {e(result['new_version'])} &middot; checked by {e(result['user'])}</div>
{banner}{removed}
<div class="card"><h2>Checks</h2><table>{checks}</table></div>
<div class="card"><h2>Score formulas compared with the live form</h2>
<p>Added: {listing(formulas['added'])}</p><p>Removed: {listing(formulas['removed'])}</p>
<table><tr><th>Changed field</th><th>Live</th><th>New</th></tr>{changed or '<tr><td colspan=3>none</td></tr>'}</table></div>
{f"<div class=card><h2>Scoring test results</h2><pre>{e(result['test_output'])}</pre></div>" if result['test_output'] else ''}
{button}"""
    return page(f'Check {result["form_id"]}', body)


def history_row(entry, csrf):
    entry_id = entry['_id']
    files = sorted(entry.get('_attachments', {}))
    links = ' '.join(f'<a href="history-file?id={urllib.parse.quote(entry_id)}&amp;name={e(name)}">{e(name)}</a>'
                     for name in files)

    def restore(which, label):
        return (f'<form method="post" action="restore"><input type="hidden" name="csrf" value="{e(csrf)}">'
                f'<input type="hidden" name="id" value="{e(entry_id)}">'
                f'<input type="hidden" name="which" value="{which}"><button>{label}</button></form>')

    buttons = ''
    if 'new.xml' in files:
        buttons += restore('new', 'Check this version again')
    if 'previous.xml' in files:
        buttons += restore('previous', 'Check the version it replaced')
    when = entry.get('started_at', '')[:19].replace('T', ' ')
    return (f'<tr><td>{e(when)} UTC</td><td>{e(entry.get("published_by"))}</td>'
            f'<td>{e(entry.get("previous_version"))} &rarr; {e(entry.get("new_version"))}</td>'
            f'<td>{e(entry.get("status"))}</td><td>{links}</td><td>{buttons}</td></tr>')


def manage_card(couch, form_id, csrf):
    doc = couch.get_doc(f'form:{form_id}', attachments='false')
    if not doc:
        return ('<div class="card"><h2>Not live</h2><p>This form is deleted. To bring it back, '
                'use "Check the version it replaced" on the latest history row, then publish.</p></div>')
    reports = couch.report_counts().get(form_id, 0)
    hidden = f'<input type="hidden" name="csrf" value="{e(csrf)}"><input type="hidden" name="id" value="{e(form_id)}">'
    if is_archived(doc):
        archive = (f'<p><b>Archived</b> by {e(doc["satoru_archived"].get("by"))}. Nobody can start new reports. '
                   f'Old reports still show.</p><form method="post" action="archive">{hidden}'
                   f'<input type="hidden" name="action" value="unarchive"><button>Unarchive</button></form>')
    else:
        archive = (f'<p><b>Archive</b>: nobody can start new reports with this form. The {e(reports)} saved report(s) '
                   f'still show and export. You can unarchive any time.</p><form method="post" action="archive">{hidden}'
                   f'<input type="hidden" name="action" value="archive"><button>Archive</button></form>')
    if reports:
        delete = (f'<p><b>Delete</b> is not possible: {e(reports)} saved report(s) use this form. Archive it instead.</p>')
    else:
        delete = (f'<p><b>Delete</b>: removes the form. A copy is kept in the history below. '
                  f'Type <code>{e(form_id)}</code> to confirm.</p><form method="post" action="delete">{hidden}'
                  f'<input type="text" name="confirm" autocomplete="off" placeholder="{e(form_id)}"> '
                  f'<button style="background:#b00">Delete form</button></form>')
    return f'<div class="card"><h2>{e(title_text(doc))}</h2>{archive}<hr>{delete}</div>'


def history_page(couch, form_id, csrf, message=''):
    rows = ''.join(history_row(h, csrf) for h in couch.history(form_id))
    body = f"""<h1>Manage: {e(form_id)}</h1>{message}{manage_card(couch, form_id, csrf)}
<h2>History</h2><div class="sub">Every publish, archive and delete from this panel. Going back to a version runs the same checks before it can be published.</div>
<div class="card"><table><tr><th>When</th><th>Who</th><th>Version</th><th>Status</th><th>Files</th><th></th></tr>
{rows or '<tr><td colspan=6>Nothing published from the panel yet.</td></tr>'}</table></div>"""
    return page(f'History {form_id}', body)


# ---------------------------------------------------------------------------
# request handling (called by export_server's handler)
# ---------------------------------------------------------------------------

class Panel:
    def __init__(self, couch_url):
        self.couch = Couch(couch_url)
        self.secret = secrets.token_bytes(32)

    def csrf(self, auth_cookie):
        return csrf_token(self.secret, auth_cookie)

    def handle_get(self, path, query, auth_cookie):
        csrf = self.csrf(auth_cookie)
        if path in ('', '/'):
            message = ''
            if query.get('deleted'):
                message = f'<div class="banner ok">Deleted <b>{e(query["deleted"][0])}</b>. A copy is in its history.</div>'
            if query.get('published'):
                message = f'<div class="banner ok">Published <b>{e(query["published"][0])}</b>. ' \
                          f'Users get it the next time their app syncs.</div>'
            return 200, 'text/html', forms_page(self.couch, csrf, message)
        if path == '/review':
            token = (query.get('token') or [''])[0]
            return 200, 'text/html', review_page(load_result(staging_path(token)), token, csrf)
        if path == '/history':
            form_id = (query.get('form') or [''])[0]
            if not FORM_ID_RE.match(form_id):
                raise PanelError('Unknown form.')
            done = (query.get('done') or [''])[0]
            message = f'<div class="banner ok">{e(done.capitalize())}.</div>' if done in ('archived', 'unarchived') else ''
            return 200, 'text/html', history_page(self.couch, form_id, csrf, message)
        if path == '/history-file':
            entry_id, name = (query.get('id') or [''])[0], (query.get('name') or [''])[0]
            if not HISTORY_ID_RE.match(entry_id) or name not in ('new.xml', 'new.xlsx', 'previous.xml'):
                raise PanelError('Unknown file.')
            url = f'{self.couch.history_db()}/{urllib.parse.quote(entry_id, safe="")}/{name}'
            resp = self.couch._check(self.couch.session.get(url, timeout=60), 'Reading history file')
            return 200, resp.headers.get('Content-Type', 'application/octet-stream'), resp.content, \
                f'{entry_id.split(":")[0]}-{name}'
        return 404, 'text/plain', 'Not found'

    def handle_post(self, path, content_type, body, auth_cookie, user):
        if path == '/check':
            values, files = parse_multipart(content_type, body)
        else:
            values = {k: v[0] for k, v in urllib.parse.parse_qs(body.decode('utf-8', 'replace')).items()}
            files = {}
        if not hmac.compare_digest(values.get('csrf', ''), self.csrf(auth_cookie)):
            raise PanelError('This page is out of date. Go back, reload it and try again.')

        if path == '/check':
            filename, data = files.get('xlsx') or ('', b'')
            if not data:
                raise PanelError('Choose a .xlsx file to upload.')
            form_id = values.get('form_id') or os.path.splitext(filename)[0].strip().lower()
            if not FORM_ID_RE.match(form_id):
                raise PanelError(f'"{form_id}" is not a valid form id: use lowercase letters, numbers and _ '
                                 f'(rename the file, e.g. hearing_screening.xlsx).')
            new_properties = None
            if not values.get('form_id'):
                if self.couch.get_doc(f'form:{form_id}', attachments='false'):
                    raise PanelError(f'A form "{form_id}" already exists. Choose it in "Which form" to replace it.')
                if not values.get('title', '').strip():
                    raise PanelError('A new form needs a title.')
                person = values.get('for') != 'place'
                new_properties = {'title': values['title'].strip(), 'icon': 'icon-healthcare-assessment',
                                  'context': {'person': person, 'place': not person}}
            if not PIPELINE_LOCK.acquire(timeout=600):
                raise PanelError('Another form is being checked. Try again in a minute.')
            try:
                token, staged = new_staging()
                check_form(self.couch, staged, form_id, user, 'upload', xlsx=data, new_properties=new_properties)
            finally:
                PIPELINE_LOCK.release()
            return 303, f'review?token={token}'

        if path == '/publish':
            form_id, notes = publish(self.couch, values.get('token', ''), user)
            return 303, f'./?published={urllib.parse.quote(form_id)}'

        if path in ('/archive', '/delete'):
            form_id = values.get('id', '')
            if not FORM_ID_RE.match(form_id):
                raise PanelError('Unknown form.')
            if path == '/delete':
                delete_form(self.couch, form_id, user, values.get('confirm', ''))
                return 303, f'./?deleted={urllib.parse.quote(form_id)}'
            archive = values.get('action') == 'archive'
            set_archived(self.couch, form_id, user, archive)
            return 303, f'history?form={urllib.parse.quote(form_id)}&done={"archived" if archive else "unarchived"}'

        if path == '/restore':
            entry_id = values.get('id', '')
            if not HISTORY_ID_RE.match(entry_id):
                raise PanelError('Unknown history entry.')
            name = 'previous.xml' if values.get('which') == 'previous' else 'new.xml'
            url = f'{self.couch.history_db()}/{urllib.parse.quote(entry_id, safe="")}/{name}'
            xml = self.couch._check(self.couch.session.get(url, timeout=60), 'Reading history').content
            form_id = entry_id.split(':')[0]
            entry = self.couch.get_doc(entry_id, db=self.couch.history_db())
            # a deleted form has no live document to copy title / context from
            restore_properties = None if self.couch.get_doc(f'form:{form_id}', attachments='false') \
                else (entry or {}).get('previous_properties')
            if not PIPELINE_LOCK.acquire(timeout=600):
                raise PanelError('Another form is being checked. Try again in a minute.')
            try:
                token, staged = new_staging()
                check_form(self.couch, staged, form_id, user, f'history {entry_id} {name}', xml=xml,
                           new_properties=restore_properties)
            finally:
                PIPELINE_LOCK.release()
            return 303, f'review?token={token}'

        return 404, 'Not found'
