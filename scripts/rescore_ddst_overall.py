"""rescore_ddst_overall.py - fix the overall result of saved DDST-II reports.

Until form version 2026-09-17 every DDST-II report was saved as overall
"Suspect", even when every domain was Normal (SCORING_AUDIT.md X4). The four
domain results were saved correctly, so the right overall result can be
recomputed from them with the form's own rule:

    Suspect  if any domain is Suspect, or 2+ domains are Caution
    Normal   otherwise

    python scripts/rescore_ddst_overall.py --url URL              # show changes only
    python scripts/rescore_ddst_overall.py --url URL --apply      # back up, then fix

--apply first writes every report it will change to
ddst-rescore-backup-<time>.json, then saves each report with the corrected
`overall` and a `scoring_fix` note (old value, new value, date, reason).
"""

import argparse
import json
import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import forms_panel

DOMAINS = ('d_ps_status', 'd_fm_status', 'd_lang_status', 'd_gm_status')
REASON = 'overall was always Suspect before form version 2026-09-17 (any_suspect text flag bug)'


def correct_overall(fields):
    statuses = [fields.get(d) for d in DOMAINS]
    if any(s not in ('Normal', 'Caution', 'Suspect') for s in statuses):
        return None                                   # incomplete report: leave alone
    if 'Suspect' in statuses or statuses.count('Caution') >= 2:
        return 'Suspect'
    return 'Normal'


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--url', required=True, help='CouchDB medic URL with credentials')
    ap.add_argument('--apply', action='store_true', help='write the fixes (default: only show them)')
    args = ap.parse_args()

    couch = forms_panel.Couch(args.url)
    resp = couch.session.get(f'{couch.medic}/_design/medic-client/_view/reports_by_form',
                             params={'key': json.dumps(['ddst']), 'include_docs': 'true', 'reduce': 'false'},
                             timeout=300)
    docs = [row['doc'] for row in couch._check(resp, 'Reading DDST reports').json()['rows']]

    changes = []
    for doc in docs:
        fields = doc.get('fields') or {}
        if 'scoring_fix' in fields:
            continue
        new = correct_overall(fields)
        if new and new != fields.get('overall'):
            changes.append((doc, fields.get('overall'), new))

    print(f'{len(docs)} DDST-II report(s), {len(changes)} to fix')
    for doc, old, new in changes:
        f = doc['fields']
        print(f'  {doc["_id"]}  {f.get("patient_name", "")}: {old} -> {new}  '
              f'({", ".join(f"{d[2:-7]}={f.get(d)}" for d in DOMAINS)})')
    if not changes or not args.apply:
        if changes:
            print('\nnothing written - run again with --apply to fix')
        return

    stamp = datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')
    backup = f'ddst-rescore-backup-{stamp}.json'
    with open(backup, 'w', encoding='utf-8') as f:
        json.dump([doc for doc, _, _ in changes], f, indent=1)
    print(f'\nbackup written: {backup}')

    for doc, old, new in changes:
        doc['fields']['overall'] = new
        doc['fields']['scoring_fix'] = {'field': 'overall', 'old': old, 'new': new,
                                        'date': forms_panel.now_iso(), 'reason': REASON}
        couch.put_doc(doc)
        print(f'  fixed {doc["_id"]}')


if __name__ == '__main__':
    main()
