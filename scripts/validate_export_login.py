"""validate_export_login.py - checks the login on the exports page.

    python scripts/validate_export_login.py

Offline: runs the real HTTP handler with a fake session lookup, plus the
parsing of CouchDB /_session answers. Needs no CHT instance.
"""

import os
import sys
import threading
import urllib.request
from http.server import HTTPServer
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import export_server as es

FAILS = []


def check(name, cond, extra=''):
    print(('  ok   ' if cond else '  FAIL ') + name + (f'  [{extra}]' if extra and not cond else ''))
    if not cond:
        FAILS.append(name)


class FakeResp:
    def __init__(self, status, payload=None, bad_json=False):
        self.status_code = status
        self._payload = payload
        self._bad = bad_json

    def json(self):
        if self._bad:
            raise ValueError('nope')
        return self._payload


def test_helpers():
    print('helpers')
    check('session url drops credentials',
          es.session_url('http://medic:secret@haproxy:5984/medic') == 'http://haproxy:5984/_session')
    check('session url keeps https + port',
          es.session_url('https://a:b@x.org:10443/medic') == 'https://x.org:10443/_session')
    check('cookie found among others',
          es.auth_cookie('foo=1; AuthSession=abc123; bar=2') == 'abc123')
    check('no cookie header', es.auth_cookie(None) is None)
    check('no AuthSession', es.auth_cookie('foo=1') is None)
    check('empty AuthSession', es.auth_cookie('AuthSession=') is None)
    check('garbage header', es.auth_cookie('";;;=\x00') is None)
    check('redirect keeps path + query',
          es.login_redirect('/exports/export?scope=all-workbook')
          == '/medic/login?redirect=/exports/export%3Fscope%3Dall-workbook')


def test_lookup():
    print('lookup_session (CouchDB answers)')
    admin = {'ok': True, 'userCtx': {'name': 'medic', 'roles': ['_admin', 'admin', 'mm-online']}}
    cases = [
        ('admin', FakeResp(200, admin), ('medic', ['_admin', 'admin', 'mm-online'])),
        ('expired cookie (name null)', FakeResp(200, {'ok': True, 'userCtx': {'name': None, 'roles': []}}), (None, [])),
        ('forged cookie (400)', FakeResp(400, {'error': 'bad_request'}), (None, [])),
        ('401', FakeResp(401, {}), (None, [])),
        ('roles not a list', FakeResp(200, {'userCtx': {'name': 'x', 'roles': 'admin'}}), ('x', [])),
    ]
    for label, resp, want in cases:
        with mock.patch.object(es.requests, 'get', return_value=resp) as g:
            got = es.lookup_session('http://medic:secret@haproxy:5984/medic', 'tok')
        check(label, got == want, got)
    sent = g.call_args
    check('request has no basic auth', 'auth' not in sent.kwargs and '@' not in sent.args[0], sent)
    check('request sends only the cookie', sent.kwargs['headers'].get('Cookie') == 'AuthSession=tok')

    for label, kw in [('500 -> error', {'return_value': FakeResp(500)}),
                      ('not JSON -> error', {'return_value': FakeResp(200, bad_json=True)}),
                      ('unreachable -> error', {'side_effect': es.requests.ConnectionError('down')})]:
        with mock.patch.object(es.requests, 'get', **kw):
            try:
                es.lookup_session('http://h:5984/medic', 'tok')
                check(label, False, 'no exception')
            except es.SessionCheckError:
                check(label, True)


def serve(check_fn, require_login=True):
    handler = es.make_handler('http://medic:secret@haproxy:5984/medic', '/exports',
                              require_login=require_login, check=check_fn, live_forms=False)
    srv = HTTPServer(('127.0.0.1', 0), handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k):
        return None


def get(srv, path, cookie=None):
    req = urllib.request.Request(f'http://127.0.0.1:{srv.server_port}{path}')
    if cookie:
        req.add_header('Cookie', cookie)
    opener = urllib.request.build_opener(NoRedirect)
    try:
        r = opener.open(req, timeout=10)
        return r.status, r.headers, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.headers, e.read()


def test_http():
    print('HTTP handler')
    sessions = {'admintok': ('medic', ['_admin']),
                'chwtok': ('chw1', ['chw', 'mm-online']),
                'expired': (None, [])}

    def fake(token):
        if token == 'boom':
            raise es.SessionCheckError('couch down')
        if token == 'crash':
            raise RuntimeError('unexpected')
        return sessions.get(token, (None, []))

    built = []
    with mock.patch.object(es, '_snapshot', return_value=([], [], 0)), \
         mock.patch.object(es.export_emr, 'build_workbook', side_effect=lambda *a: built.append(a) or SystemExit('stub')):
        srv = serve(fake)
        for path in ('/exports/', '/exports/export?scope=all-workbook', '/exports/nope'):
            st, h, _ = get(srv, path)
            check(f'no cookie {path} -> 302 login', st == 302 and h['Location'].startswith('/medic/login?redirect=/exports/'), (st, h.get('Location')))
            check(f'no cookie {path} -> no-store', h.get('Cache-Control') == 'no-store')
        st, h, _ = get(srv, '/exports/', 'AuthSession=forged')
        check('unknown cookie -> 302', st == 302)
        st, h, _ = get(srv, '/exports/', 'AuthSession=expired')
        check('expired cookie -> 302', st == 302)
        st, _, body = get(srv, '/exports/export?scope=all-workbook', 'AuthSession=chwtok')
        check('non-admin -> 403', st == 403, st)
        check('non-admin export never built', not built)
        st, _, _ = get(srv, '/exports/export?scope=all-workbook', 'AuthSession=boom')
        check('couch down -> 503', st == 503, st)
        st, _, _ = get(srv, '/exports/export?scope=all-workbook', 'AuthSession=crash')
        check('unexpected error -> 503', st == 503, st)
        check('nothing built for denied requests', not built)
        st, h, body = get(srv, '/exports/', 'AuthSession=admintok')
        check('admin -> page 200', st == 200 and b'Excel exports' in body, st)
        check('admin page no-store + nosniff', h.get('Cache-Control') == 'no-store' and h.get('X-Content-Type-Options') == 'nosniff')
        get(srv, '/exports/export?scope=all-workbook', 'AuthSession=admintok')
        check('admin export reaches the exporter', len(built) == 1)
        check('HEAD not served', get_method(srv, 'HEAD') == 501)
        check('POST without login -> sent to login', get_method(srv, 'POST') == 302)
        check('POST to exports page -> 405', post(srv, '/exports/', 'AuthSession=admintok') == 405)
        check('POST without Origin -> 403', post(srv, '/exports/forms/publish', 'AuthSession=admintok', origin=None) == 403)
        check('POST with Origin: null -> 403', post(srv, '/exports/forms/publish', 'AuthSession=admintok', origin='null') == 403)
        check('pages let the browser send our Origin on posts', h.get('Referrer-Policy') == 'same-origin', h.get('Referrer-Policy'))
        check('POST from another site -> 403', post(srv, '/exports/forms/publish', 'AuthSession=admintok', origin='https://evil.example') == 403)
        check('POST with bad CSRF token -> 400', post(srv, '/exports/forms/publish', 'AuthSession=admintok', body=b'csrf=nope&token=x') == 400)
        check('POST by non-admin -> 403', post(srv, '/exports/forms/publish', 'AuthSession=chwtok') == 403)
        check('POST too big -> 413', post(srv, '/exports/forms/check', 'AuthSession=admintok', length=10**9) == 413)
        srv.shutdown()

        srv = serve(lambda t: (_ for _ in ()).throw(AssertionError('must not be called')), require_login=False)
        st, _, _ = get(srv, '/exports/')
        check('--no-login mode serves without cookie', st == 200, st)
        srv.shutdown()


def get_method(srv, method):
    req = urllib.request.Request(f'http://127.0.0.1:{srv.server_port}/exports/', method=method)
    try:
        return urllib.request.build_opener(NoRedirect).open(req, timeout=10).status
    except urllib.error.HTTPError as e:
        return e.code


def post(srv, path, cookie, origin='same', body=b'csrf=x', length=None):
    import http.client
    conn = http.client.HTTPConnection('127.0.0.1', srv.server_port, timeout=10)
    headers = {'Cookie': cookie, 'Content-Type': 'application/x-www-form-urlencoded', 'Host': '127.0.0.1',
               'Content-Length': str(length if length is not None else len(body))}
    if origin == 'same':
        headers['Origin'] = f'http://127.0.0.1:{srv.server_port}'
    elif origin:
        headers['Origin'] = origin
    conn.putrequest('POST', path, skip_host=True, skip_accept_encoding=True)
    for k, v in headers.items():
        conn.putheader(k, v)
    conn.endheaders()
    if length is None:
        conn.send(body)
    status = conn.getresponse().status
    conn.close()
    return status


def test_cli():
    print('command line')
    with mock.patch.object(sys, 'argv', ['x', '--url', 'http://h/medic', '--host', '0.0.0.0', '--no-login']), \
         mock.patch.object(sys, 'stderr'):
        try:
            es.main()
            check('--no-login refused on 0.0.0.0', False)
        except SystemExit as e:
            check('--no-login refused on 0.0.0.0', e.code == 2)


if __name__ == '__main__':
    test_helpers()
    test_lookup()
    test_http()
    test_cli()
    print(f'\n{"ALL PASSED" if not FAILS else f"{len(FAILS)} FAILED"}')
    sys.exit(1 if FAILS else 0)
