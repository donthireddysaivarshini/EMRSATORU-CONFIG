"""validate_forms_panel.py - offline checks for the forms panel helpers.

    python scripts/validate_forms_panel.py

The full pipeline (convert, scoring tests, publish, rollback) needs the
container; see SETUP.md "Forms panel" for the manual check.
"""

import os
import sys
import tempfile
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import forms_panel as fp

FAILS = []


def check(name, cond, extra=''):
    print(('  ok   ' if cond else '  FAIL ') + name + (f'  [{extra}]' if extra and not cond else ''))
    if not cond:
        FAILS.append(name)


def raises(fn):
    try:
        fn()
    except fp.PanelError:
        return True
    return False


XML = '''<h:html><h:head><h:title>T</h:title><model><instance>
<data id="moca_assessment" version="2026-09-17 00-00" prefix="J1!x!">
</data></instance>
<bind nodeset="/data/a" type="string" calculate="if( /data/b  &gt;= 1, 1, 0)"/>
<bind nodeset="/data/b" type="int"/>
</model></h:head></h:html>'''


def main():
    print('parsing forms')
    check('form id', fp.form_xml_id(XML) == 'moca_assessment')
    check('version', fp.form_version(XML) == '2026-09-17 00-00')
    check('calculations unescaped', fp.calculations(XML) == {'/data/a': 'if( /data/b  >= 1, 1, 0)'})
    check('fields', fp.fields(XML) == {'/data/a', '/data/b'})

    print('input validation')
    for bad in ['', 'A', 'x', '../etc', 'form:x', 'a b', 'a' * 70]:
        check(f'form id rejected: {bad[:12]!r}', not fp.FORM_ID_RE.match(bad))
    check('form id accepted', bool(fp.FORM_ID_RE.match('hearing_test2')))
    check('bad token rejected', raises(lambda: fp.staging_path('../../etc')))
    check('unknown token rejected', raises(lambda: fp.staging_path('a' * 32)))
    check('history id shape', bool(fp.HISTORY_ID_RE.match('moca_assessment:2026-09-17T08:31:39.037537Z:f978948d')))
    check('history id rejects traversal', not fp.HISTORY_ID_RE.match('x:../../_users:abcdef12'))

    print('media')
    with tempfile.TemporaryDirectory() as media:
        os.makedirs(os.path.join(media, 'images'))
        open(os.path.join(media, 'images', 'cube.png'), 'wb').close()
        xml = '<value form="image">jr://images/cube.png</value><value form="image">jr://images/lion.png</value>'
        check('missing picture found', fp.missing_media(xml, media) == ['images/lion.png'], fp.missing_media(xml, media))
        check('form without media is fine', fp.missing_media('<x/>', media) == [])

    print('labels for any form')
    with tempfile.NamedTemporaryFile('w', suffix='.xml', delete=False) as f:
        f.write('''<h:html xmlns:h="http://www.w3.org/1999/xhtml"><h:head><h:title>T</h:title><model><instance>
<data id="t"><inputs><contact><_id/></contact></inputs><patient_name/><g_one><hears_clap/></g_one><meta><instanceID/></meta></data>
</instance><bind nodeset="/data/inputs/contact/_id" type="string"/><bind nodeset="/data/patient_name" type="string"/>
<bind nodeset="/data/g_one/hears_clap" type="string"/><bind nodeset="/data/meta/instanceID" type="string"/></model></h:head>
<h:body><group ref="/data/g_one"><label>Hearing</label><select1 ref="/data/g_one/hears_clap"><label>Turns to a clap?</label></select1></group></h:body></h:html>''')
    labels = dict(fp.generic_labels('t', f.name))
    os.unlink(f.name)
    check('question text used', labels.get('report.t.g_one.hears_clap') == 'Turns to a clap?', labels)
    check('group label used', labels.get('report.t.g_one') == 'Hearing', labels)
    check('calculated field name made readable', labels.get('report.t.patient_name') == 'Participant name', labels)
    check('inputs and meta skipped', not any('inputs' in k or 'instanceID' in k for k in labels), labels)

    print('csrf')
    secret = b's' * 32
    check('token depends on login cookie', fp.csrf_token(secret, 'a') != fp.csrf_token(secret, 'b'))
    check('token stable', fp.csrf_token(secret, 'a') == fp.csrf_token(secret, 'a'))

    print('multipart')
    body = (b'--B\r\nContent-Disposition: form-data; name="csrf"\r\n\r\nabc\r\n'
            b'--B\r\nContent-Disposition: form-data; name="xlsx"; filename="../../moca.xlsx"\r\n'
            b'Content-Type: application/octet-stream\r\n\r\nPK\x03\x04data\r\n--B--\r\n')
    values, files = fp.parse_multipart('multipart/form-data; boundary=B', body)
    check('field value', values.get('csrf') == 'abc')
    check('file name stripped of folders', files['xlsx'][0] == 'moca.xlsx')
    check('file bytes kept', files['xlsx'][1] == b'PK\x03\x04data')
    check('non-multipart rejected', raises(lambda: fp.parse_multipart('text/plain', b'x')))

    print('output summaries')
    out = 'INFO a\nWARN careful\nERROR broken\n  at x (y)\n'
    check('cht-conf errors only', fp.cht_conf_errors(out) == 'WARN careful\nERROR broken')
    mocha = '  ✔ one\n  1) two\n     AssertionError: expected 1\n      at Context.<anonymous> (a.js:1)\n  1 passing\n'
    check('mocha summary drops stack', 'at Context' not in fp.test_summary(mocha) and 'AssertionError' in fp.test_summary(mocha))

    print('staging cleanup')
    with tempfile.TemporaryDirectory() as root, mock.patch.object(fp, 'STAGING_ROOT', root):
        token, path = fp.new_staging()
        check('staging dir private', oct(os.stat(path).st_mode)[-3:] == '700')
        os.utime(path, (0, 0))
        fp.clean_staging()
        check('old staging removed', not os.path.exists(path))

    print(f'\n{"ALL PASSED" if not FAILS else f"{len(FAILS)} FAILED"}')
    sys.exit(1 if FAILS else 0)


if __name__ == '__main__':
    main()
