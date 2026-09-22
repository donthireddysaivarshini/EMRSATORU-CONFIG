"""export_server.py - point-and-click UI for the Excel exports.

A tiny local web page so nobody has to type a command. Same exporter as
scripts/export_emr.py (identical output), just driven from a browser:

    python scripts/export_server.py --url "https://medic:PASSWORD@HOST:10443/medic"

then open http://127.0.0.1:8095

  * "All participants"        -> all-workbook
  * pick a camp               -> camp-workbook=<id>
  * pick a participant        -> participant-report=<id>

Stdlib only (plus openpyxl/requests, already required by export_emr).
Read-only: only GET requests are made against CouchDB.

The CouchDB URL (with its password) stays on the server side - it is never
sent to the browser.

Login: every request must carry the browser's CHT login cookie (AuthSession)
belonging to a CHT admin. The cookie is checked against CouchDB's /_session
on every request, WITHOUT the service password, so only a real logged-in
admin gets through. No/expired cookie -> sent to the CHT login page; logged in
but not an admin -> 403; CouchDB unreachable -> 503 (never let through).
--no-login exists only for running on your own machine and refuses to start
unless bound to localhost.
"""

import argparse
import hashlib
import html
import io
import os
import sys
import tempfile
import urllib.parse
from http.cookies import CookieError, SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import requests

from emr_common import EmrDb, classify_contacts, is_emr_report
import export_emr
import forms_panel
import scale_meta

XLSX_MIME = ('application/vnd.openxmlformats-officedocument'
             '.spreadsheetml.sheet')

PAGE = """<!doctype html>
<html><head><meta charset="utf-8"><title>Satoru EMR exports</title>
<style>
 body{{font:15px/1.5 system-ui,sans-serif;margin:0;background:#f2f2f3;color:#1a1a1a}}
 .wrap{{max-width:760px;margin:0 auto;padding:32px 20px 64px}}
 h1{{font-size:22px;margin:0 0 4px}}
 .sub{{color:#666;font-size:13px;margin-bottom:28px}}
 .card{{background:#fff;border:1px solid #e2e2e5;border-radius:8px;padding:20px;margin-bottom:16px}}
 .card h2{{font-size:16px;margin:0 0 6px}}
 .card p{{color:#666;font-size:13px;margin:0 0 14px}}
 select{{font:inherit;padding:8px;border:1px solid #ccc;border-radius:5px;min-width:280px;background:#fff}}
 button{{font:inherit;padding:9px 16px;border:0;border-radius:5px;background:#0a5ca8;color:#fff;cursor:pointer}}
 button:hover{{background:#08498a}}
 .row{{display:flex;gap:10px;align-items:center;flex-wrap:wrap}}
 .empty{{color:#999;font-style:italic}}
 code{{background:#eee;padding:1px 5px;border-radius:3px;font-size:12px}}
</style></head><body><div class="wrap">
<h1>Satoru EMR &mdash; Excel exports</h1>
<div class="sub">Connected to <code>{host}</code> &middot; {ncamps} camps, {nparts} participants, {nreports} assessments &middot; <a href="forms/">Upload / manage forms</a></div>

<div class="card">
  <h2>Everyone</h2>
  <p>All participants across every camp. One sheet per scale, plus a Camp column.</p>
  <form action="export" method="get"><input type="hidden" name="scope" value="all-workbook">
  <button type="submit">Download all participants</button></form>
</div>

<div class="card">
  <h2>One camp</h2>
  <p>Participants sheet plus one sheet per scale completed in that camp.</p>
  <form action="export" method="get" class="row">
    <select name="scope">{camp_options}</select>
    <button type="submit">Download camp workbook</button>
  </form>
</div>

<div class="card">
  <h2>One participant</h2>
  <p>Report-style workbook: a sectioned sheet per completed assessment.</p>
  <form action="export" method="get" class="row">
    <select name="scope">{part_options}</select>
    <button type="submit">Download participant report</button>
  </form>
</div>
</div></body></html>"""


ADMIN_ROLES = frozenset({'_admin', 'admin'})
LOCAL_HOSTS = ('127.0.0.1', 'localhost', '::1')


class SessionCheckError(Exception):
    """CouchDB could not tell us who the user is. Always fail closed."""


def session_url(couch_url):
    """http://user:pw@haproxy:5984/medic -> http://haproxy:5984/_session
    (credentials dropped on purpose: with them, everyone is 'the service')."""
    parts = urllib.parse.urlsplit(couch_url)
    host = parts.hostname or ''
    if ':' in host:
        host = f'[{host}]'
    netloc = f'{host}:{parts.port}' if parts.port else host
    return urllib.parse.urlunsplit((parts.scheme, netloc, '/_session', '', ''))


def auth_cookie(cookie_header):
    """The AuthSession value from a Cookie header, or None."""
    if not cookie_header:
        return None
    jar = SimpleCookie()
    try:
        jar.load(cookie_header)
    except CookieError:
        return None
    morsel = jar.get('AuthSession')
    return morsel.value if morsel and morsel.value else None


def lookup_session(couch_url, token, timeout=10):
    """(name, roles) for a CHT login cookie; (None, []) if not logged in.
    Raises SessionCheckError when CouchDB can't answer."""
    try:
        resp = requests.get(session_url(couch_url),
                            headers={'Cookie': f'AuthSession={token}',
                                     'Accept': 'application/json'},
                            timeout=timeout, verify=False, allow_redirects=False)
    except requests.RequestException as exc:
        raise SessionCheckError(str(exc)) from exc
    if resp.status_code in (400, 401, 403):       # malformed / bad cookie
        return None, []
    if resp.status_code != 200:
        raise SessionCheckError(f'/_session returned HTTP {resp.status_code}')
    try:
        ctx = resp.json().get('userCtx') or {}
    except ValueError as exc:
        raise SessionCheckError('/_session did not return JSON') from exc
    name = ctx.get('name')
    roles = ctx.get('roles') if isinstance(ctx.get('roles'), list) else []
    if not isinstance(name, str) or not name:   # expired cookie -> name null
        return None, []
    return name, roles


def authorize(cookie_header, check):
    """Decide one request. check(token) -> (name, roles).
    Returns ('ok', name) | ('login', None) | ('forbidden', name) | ('error', msg)."""
    token = auth_cookie(cookie_header)
    if not token:
        return 'login', None
    try:
        name, roles = check(token)
    except SessionCheckError as exc:
        return 'error', str(exc)
    except Exception as exc:                      # noqa: BLE001 - fail closed
        return 'error', f'session check failed: {exc}'
    if not name:
        return 'login', None
    if not ADMIN_ROLES.intersection(roles):
        return 'forbidden', name
    return 'ok', name


def login_redirect(request_path):
    """CHT's login page sends the user back to `redirect` after login."""
    return '/medic/login?redirect=' + urllib.parse.quote(request_path, safe='/')


def _options(pairs, prefix):
    if not pairs:
        return '<option value="">(none found)</option>'
    return ''.join(
        f'<option value="{html.escape(prefix + i)}">{html.escape(label)}</option>'
        for i, label in pairs)


def _snapshot(db):
    """Camps and participants for the dropdowns (one read)."""
    docs = db.fetch_all_docs()
    camps, participants = classify_contacts(docs)
    reports = [d for d in docs if is_emr_report(d)]
    staff = export_emr.report_submitter_ids(reports)

    camp_pairs = sorted(
        ((cid, c.get('name') or cid) for cid, c in camps.items()),
        key=lambda p: p[1].lower())
    part_pairs = sorted(
        ((pid, ' — '.join(x for x in (p.get('cr_no'), p.get('name')) if x) or pid)
         for pid, p in participants.items() if pid not in staff),
        key=lambda p: p[1].lower())
    return camp_pairs, part_pairs, len(reports)


LIVE_FORMS_DIR = os.path.join(tempfile.gettempdir(), 'satoru-live-forms')
_live_forms_hash = None


def sync_live_forms(couch):
    """Exports read the scale forms as they are live in CouchDB (the forms
    panel can publish new versions), not the copies baked into the image.
    Call while holding scale_meta.LOCK."""
    global _live_forms_hash
    contents = {info['file']: couch.get_attachment(f'form:{form}', 'xml')
                for form, info in scale_meta.SCALES.items()}
    digest = hashlib.sha256(b''.join(contents[k] for k in sorted(contents))).hexdigest()
    if digest != _live_forms_hash:
        os.makedirs(LIVE_FORMS_DIR, exist_ok=True)
        for name, data in contents.items():
            with open(os.path.join(LIVE_FORMS_DIR, name), 'wb') as f:
                f.write(data)
        scale_meta.set_forms_dir(LIVE_FORMS_DIR)
        scale_meta._cache.clear()
        _live_forms_hash = digest


def same_origin(headers):
    """POSTs must come from our own pages (Origin/Referer host = Host)."""
    host = (headers.get('Host') or '').split(':')[0].lower()
    source = headers.get('Origin') or headers.get('Referer')
    if not host or not source:
        return False
    return (urllib.parse.urlsplit(source).hostname or '').lower() == host


def make_handler(url, base_path='', require_login=True, check=None, live_forms=True):
    check = check or (lambda token: lookup_session(url, token))
    panel = forms_panel.Panel(url)

    class Handler(BaseHTTPRequestHandler):
        user = None

        def log_message(self, fmt, *args):
            sys.stderr.write('  %s\n' % (fmt % args))

        def end_headers(self):
            # patient data: never cache, never sniff, never frame
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('X-Frame-Options', 'DENY')
            # same-origin, not no-referrer: with no-referrer browsers send
            # "Origin: null" on form posts and the same-site check rejects them
            self.send_header('Referrer-Policy', 'same-origin')
            super().end_headers()

        def _allowed(self):
            if not require_login:
                self.user = '(no-login mode)'
                return True
            decision, detail = authorize(self.headers.get('Cookie'), check)
            if decision == 'ok':
                self.user = detail
                return True
            if decision == 'login':
                self.send_response(302)
                self.send_header('Location', login_redirect(self.path))
                self.send_header('Content-Length', '0')
                self.end_headers()
            elif decision == 'forbidden':
                sys.stderr.write(f'  DENIED user={detail} path={self.path}\n')
                self._fail('Only CHT administrators can download exports.',
                           status=403)
            else:
                sys.stderr.write(f'  LOGIN CHECK FAILED: {detail}\n')
                self._fail('Could not check your login right now. '
                           'Try again in a minute.', status=503)
            return False

        def _db(self):
            return EmrDb(url, verify=False)

        def _route(self):
            parsed = urllib.parse.urlparse(self.path)
            path = parsed.path
            # strip the reverse-proxy prefix, if any
            if base_path and path.startswith(base_path):
                path = path[len(base_path):] or '/'
            return parsed, path

        def _send(self, status, content_type, body, filename=None):
            data = body.encode('utf-8') if isinstance(body, str) else body
            self.send_response(status)
            self.send_header('Content-Type', content_type + ('; charset=utf-8' if content_type.startswith('text/') else ''))
            if filename:
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def _redirect(self, location):
            self.send_response(303)
            self.send_header('Location', location)
            self.send_header('Content-Length', '0')
            self.end_headers()

        def _forms_get(self, parsed, path):
            if path == '/forms':
                return self._redirect(f'{base_path}/forms/')
            try:
                out = panel.handle_get(path[len('/forms'):], urllib.parse.parse_qs(parsed.query),
                                       auth_cookie(self.headers.get('Cookie')))
            except forms_panel.PanelError as exc:
                return self._fail(str(exc), status=400)
            except Exception as exc:                      # noqa: BLE001
                sys.stderr.write(f'  FORMS PAGE FAILED: {exc!r}\n')
                return self._fail('The forms page failed. Ask the admin to check the exports log.')
            self._send(*out)

        def do_POST(self):
            if not self._allowed():
                return
            parsed, path = self._route()
            if not path.startswith('/forms/'):
                return self.send_error(405)
            if not same_origin(self.headers):
                return self._fail('Request blocked: it did not come from this site.', status=403)
            try:
                length = int(self.headers.get('Content-Length') or 0)
            except ValueError:
                length = -1
            if length < 0 or length > forms_panel.MAX_UPLOAD_BYTES:
                return self._fail('The upload is too big (limit 15 MB).', status=413)
            body = self.rfile.read(length)
            sys.stderr.write(f'  FORMS user={self.user} action={path}\n')
            try:
                status, location = panel.handle_post(
                    path[len('/forms'):], self.headers.get('Content-Type') or '', body,
                    auth_cookie(self.headers.get('Cookie')), self.user)
            except forms_panel.PanelError as exc:
                sys.stderr.write(f'  FORMS REFUSED user={self.user}: {exc}\n')
                return self._fail(str(exc), status=400)
            except Exception as exc:                      # noqa: BLE001
                sys.stderr.write(f'  FORMS FAILED user={self.user}: {exc!r}\n')
                return self._fail('Something went wrong. Nothing was published unless the history page says so. '
                                  'Ask the admin to check the exports log.')
            if status == 303:
                sys.stderr.write(f'  FORMS DONE user={self.user} -> {location}\n')
                return self._redirect(location)
            self.send_error(status)

        def do_GET(self):
            if not self._allowed():
                return
            parsed, path = self._route()
            if path == '/forms' or path.startswith('/forms/'):
                return self._forms_get(parsed, path)
            if path in ('/', ''):
                return self._page()
            if path == '/export':
                qs = urllib.parse.parse_qs(parsed.query)
                return self._export((qs.get('scope') or [''])[0])
            self.send_error(404)

        def _page(self):
            try:
                with scale_meta.LOCK:
                    if live_forms:
                        sync_live_forms(panel.couch)
                    camps, parts, nreports = _snapshot(self._db())
            except Exception as exc:                      # noqa: BLE001
                sys.stderr.write(f'  PAGE FAILED: {exc}\n')
                return self._fail('Could not read the database. Try again.')
            host = urllib.parse.urlsplit(url).hostname or url
            body = PAGE.format(
                host=html.escape(host), ncamps=len(camps),
                nparts=len(parts), nreports=nreports,
                camp_options=_options(camps, 'camp-workbook='),
                part_options=_options(parts, 'participant-report='),
            ).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _export(self, scope):
            if not scope:
                return self._fail('No export selected.')
            sys.stderr.write(f'  EXPORT user={self.user} scope={scope}\n')
            try:
                with scale_meta.LOCK:
                    if live_forms:
                        sync_live_forms(panel.couch)
                    db = self._db()
                    wb = export_emr.build_workbook(scope, db)
                    name = export_emr.default_output_name(scope, wb, db)
                buf = io.BytesIO()
                wb.save(buf)
            except SystemExit as exc:                     # build_workbook's own errors
                return self._fail(str(exc))
            except Exception as exc:                      # noqa: BLE001
                sys.stderr.write(f'  EXPORT FAILED scope={scope}: {exc}\n')
                return self._fail('Export failed. Ask the admin to check '
                                  'the exports log.')
            data = buf.getvalue()
            self.send_response(200)
            self.send_header('Content-Type', XLSX_MIME)
            self.send_header('Content-Disposition',
                             f'attachment; filename="{name}"')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            sys.stderr.write(f'  -> sent {name} ({len(data):,} bytes) '
                             f'to user={self.user}\n')

        def _fail(self, message, status=500):
            body = (f'<!doctype html><meta charset="utf-8">'
                    f'<body style="font:15px system-ui;padding:32px">'
                    f'<p style="color:#b00">{html.escape(message)}</p>'
                    f'<p><a href=".">Back</a></p>').encode('utf-8')
            self.send_response(status)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    return Handler


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--url', required=True,
                    help='CouchDB medic URL with credentials')
    ap.add_argument('--port', type=int, default=8095)
    ap.add_argument('--host', default='127.0.0.1',
                    help='bind address (default localhost only)')
    ap.add_argument('--base-path', default='',
                    help='path prefix when served behind a reverse proxy, '
                         'e.g. /exports')
    ap.add_argument('--no-login', action='store_true',
                    help='skip the CHT admin login check (localhost only)')
    args = ap.parse_args()
    if args.no_login and args.host not in LOCAL_HOSTS:
        ap.error('--no-login is only allowed with --host 127.0.0.1')

    base_path = '/' + args.base_path.strip('/') if args.base_path.strip('/') else ''
    server = ThreadingHTTPServer((args.host, args.port),
                        make_handler(args.url, base_path,
                                     require_login=not args.no_login))
    print(f'Satoru EMR exports  ->  http://{args.host}:{args.port}')
    print('Ctrl+C to stop.')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nstopped')


if __name__ == '__main__':
    main()
