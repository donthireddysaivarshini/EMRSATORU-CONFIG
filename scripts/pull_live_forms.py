"""pull_live_forms.py - copy the app forms that are live in CHT into this repo.

Forms published from the forms panel (/exports/forms/) only exist in the
database. Run this before editing forms or running `cht upload-app-forms`
from a checkout, or you will overwrite the live versions with old copies.

    python scripts/pull_live_forms.py --url "https://medic:PASSWORD@HOST:10443/medic"

Writes forms/app/<id>.xml, and <id>.xlsx when the last panel publish kept one
(needs direct CouchDB access for the history database; skipped otherwise).
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import forms_panel

APP_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'forms', 'app')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--url', required=True, help='CouchDB medic URL with credentials')
    args = ap.parse_args()

    couch = forms_panel.Couch(args.url)
    for doc in couch.app_forms():
        form_id = doc.get('internalId')
        if not form_id or not forms_panel.FORM_ID_RE.match(form_id) or 'xml' not in doc.get('_attachments', {}):
            continue
        xml = couch.get_attachment(doc['_id'], 'xml')
        path = os.path.join(APP_DIR, f'{form_id}.xml')
        old = open(path, 'rb').read() if os.path.exists(path) else None
        if old != xml:
            with open(path, 'wb') as f:
                f.write(xml)
            print(f'updated  forms/app/{form_id}.xml')
        try:
            published = [h for h in couch.history(form_id) if h.get('status') == 'published']
        except forms_panel.PanelError:
            published = []
        if published and 'new.xlsx' in published[0].get('_attachments', {}) \
                and published[0].get('xml_sha256') == forms_panel.sha256(xml):
            url = f'{couch.history_db()}/{published[0]["_id"]}/new.xlsx'
            data = couch._check(couch.session.get(url, timeout=60), 'Reading xlsx').content
            with open(os.path.join(APP_DIR, f'{form_id}.xlsx'), 'wb') as f:
                f.write(data)
            print(f'updated  forms/app/{form_id}.xlsx (from panel history)')
    print('done - review with `git diff forms/app` and commit')


if __name__ == '__main__':
    main()
