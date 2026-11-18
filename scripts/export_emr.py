"""Read-only Satoru EMR -> Excel exporter.

Produces one .xlsx per run. The workbook contents depend on --scope:

  all                      Sheets: Camps | Participants | Reports
  camp=<campId>            Sheets: Camp | Participants | Reports
  participants_in_camp=..  Sheets: Participants | Reports
  participant=<id>         Sheets: Participant | Reports
  all-workbook             EVERY participant across every camp, organised by
                           scale: Participants | DDST-II | DST | MoCA | VSMS
                           (adds a 'Camp' column to every sheet)
  camp-workbook=<campId>   One workbook per camp, organised by scale:
                           Participants | DDST-II | DST | MoCA | VSMS
                           (only scales actually completed in that camp;
                           one row per participant per scale sheet)
  participant-report=<id>  Human-readable workbook for one participant's
                           completed assessments: Participant sheet + one
                           sheet per completed scale (sectioned rows)

Relationships are by exact _id only:
  participant.parent._id  == campId
  report.contact._id      == participantId
  report.contact.parent._id == campId

NOTE: on a CHT doc, `report.contact` is the SUBMITTER's contact; the
subject of a scale report is `report.fields.patient_uuid` (== the
participant contact _id). The camp of a report is therefore derived from
the SUBJECT participant's parent, not from report.contact.parent.

Repeat-group choice (documented in requirements): arrays are stored as JSON
in a single cell (lossless, schema-generic). Attachments are NOT embedded in
Excel; the `attachments` column lists file names, referenced as
`attachment/<report-id>/<file-name>`.

Scale sheets (camp-workbook / participant-report) are built through
scripts/scale_export.py, which reads the form XMLs and the English
translation file - the same sources as the Reports UI - so labels can never
drift between UI and exports. Scoring fields are consumed as stored; no
score is recalculated.

Read-only: only GET requests are made. No data is ever modified.

Usage:
  python scripts/export_emr.py --url https://user:pass@host:port/medic [--scope ...] [--out DIR]
"""

import argparse
import json
import os
import sys

try:
    import openpyxl
except ModuleNotFoundError:
    sys.exit('missing dependencies - run:  pip install -r scripts/requirements.txt')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from emr_common import (EmrDb, attachment_names, classify_contacts, flatten,
                        iso_datetime, is_emr_report, sanitize_filename,
                        timestamp_stamp)
import scale_export
import scale_meta

PREFERRED_SCALE_ORDER = [
    'assessment_visit',
    'ddst',
    'dst',
    'moca',
    'moca_assessment',
    'vineland',
    'cog5',
    'mse',
    'phq9',
    'ad8',
    'easi',
    'gds15',
    'follow_up_contact',
    'longitudinal_review',
    'referral_outcome',
]


def get_scale_order(extra_forms=None):
    """Dynamic list of all known forms in stable order.
    Starts with PREFERRED_SCALE_ORDER, then appends any other forms discovered
    in database reports or on disk (excluding contacts and training)."""
    known = list(PREFERRED_SCALE_ORDER)
    discovered = set(scale_meta.SCALES.keys())
    if extra_forms:
        discovered.update(extra_forms)
    for form in sorted(discovered):
        if form and form not in known and not form.startswith('contact:') and not form.startswith('training:'):
            known.append(form)
    return known


SCALE_ORDER = get_scale_order()
GENDER_LABELS = {'male': 'Male', 'female': 'Female'}


def fixed_columns_first(columns, priority):
    """Order columns: priority ones first (in given order), rest sorted."""
    present = [c for c in priority if c in columns]
    rest = sorted(c for c in columns if c not in priority)
    return present + rest


def collect_report_rows(reports, participants, camps):
    """Map each EMR report to its participant and camp by exact _id.

    The assessed participant is the report SUBJECT (fields.patient_uuid), not
    report.contact (which is the submitter). The camp is that participant's
    parent.
    """
    rows = []
    for doc in reports:
        participant_id = subject_id_of(doc)
        participant = participants.get(participant_id) if participant_id else None
        camp_id = ((participant.get('parent') or {}).get('_id')
                   if participant else None)
        camp = camps.get(camp_id) if camp_id else None

        fields = flatten(doc.get('fields') or {})
        row = {
            '_id': doc.get('_id'),
            'form': doc.get('form'),
            'content_type': doc.get('content_type'),
            'reported_date': iso_datetime(doc.get('reported_date')),
            'reported_date_ms': doc.get('reported_date'),
            'from': doc.get('from'),
            'submitted_by': submitter_label(doc),
            'participant_id': participant_id,
            'camp_id': camp_id,
            'participant_name': participant.get('name') if participant else None,
            'camp_name': camp.get('name') if camp else None,
            'form_version': doc.get('form_version'),
            'hidden_fields': doc.get('hidden_fields'),
            'geolocation': doc.get('geolocation'),
            'geolocation_log': doc.get('geolocation_log'),
            'attachments': attachment_names(doc),
        }
        row.update(fields)
        rows.append(row)
    return rows


def flatten_contact_columns(doc, link_cols):
    """Flatten a contact doc, carrying exact-id link columns forward."""
    fields = flatten({k: v for k, v in doc.items()
                      if k not in ('_id', '_rev', 'parent', '_attachments')})
    row = {'_id': doc.get('_id')}
    row.update(link_cols)
    row.update(fields)
    row['attachments'] = attachment_names(doc)
    return row


def subject_id_of(doc):
    """The assessed participant's contact _id (NOT the submitter)."""
    sid = (doc.get('fields') or {}).get('patient_uuid')
    return sid if isinstance(sid, str) and sid else None


def submitter_label(doc):
    """Human name (+phone) of the submitting user's contact."""
    contact = doc.get('contact') or {}
    if not isinstance(contact, dict):
        return None
    name = contact.get('name')
    phone = contact.get('phone')
    if name and phone:
        return f'{name} ({phone})'
    return name


def scale_reports_by_participant(reports, participant_ids):
    """form -> list of doc for reports whose SUBJECT is one of
    participant_ids, ordered by reported_date (then _id for stability).
    Dynamically includes ALL forms present in the reports."""
    forms_in_reports = {
        doc.get('form') for doc in reports
        if doc.get('form') and subject_id_of(doc) in participant_ids
    }
    order = get_scale_order(forms_in_reports)
    by_form = {form: [] for form in order}
    for doc in reports:
        sid = subject_id_of(doc)
        if sid not in participant_ids:
            continue
        form = doc.get('form')
        if form:
            if form not in by_form:
                by_form[form] = []
            by_form[form].append(doc)
    for docs in by_form.values():
        docs.sort(key=lambda d: (d.get('reported_date') or 0, d.get('_id') or ''))
    return by_form


def write_generic_scale_sheet(ws, form, reports, participants, camps=None):
    """Fallback sheet writer for forms without XML definitions on disk."""
    keys = []
    for doc in reports:
        for k in flatten(doc.get('fields') or {}).keys():
            if k not in keys and not k.startswith('inputs.') and not k.startswith('meta.'):
                keys.append(k)
    keys.sort()

    headers = ['Participant ID', 'Participant Name']
    if camps is not None:
        headers.append('Camp')
    headers.extend(['Report Date', 'Submitted By'])
    headers.extend([scale_meta.clean_human_label(k) for k in keys])
    ws.append(headers)

    for doc in reports:
        participant = participants.get(subject_id_of(doc)) or {}
        camp_name = ''
        if camps is not None:
            camp = camps.get((participant.get('parent') or {}).get('_id'))
            camp_name = (camp or {}).get('name') or ''
        flat_fields = flatten(doc.get('fields') or {})
        row = [
            participant.get('cr_no') or participant.get('patient_id') or '',
            participant.get('name') or '',
        ]
        if camps is not None:
            row.append(camp_name)
        row.extend([
            iso_datetime(doc.get('reported_date')),
            submitter_label(doc),
        ])
        for k in keys:
            v = flat_fields.get(k)
            row.append(json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v)
        ws.append(row)
    ws.freeze_panes = 'C2'


def write_scale_sheet(wb, form, reports, participants, camps=None):
    """One row per completed assessment of `form`, human-readable columns.

    Passing `camps` adds a 'Camp' column (used by the all-camps workbook).
    """
    ws = wb.create_sheet(scale_export.sheet_title(form))
    meta = scale_meta.get_scale_meta(form)
    if meta.get('flat'):
        try:
            ws.append(scale_export.scale_headers(form, include_camp=camps is not None))
            for doc in reports:
                participant = participants.get(subject_id_of(doc)) or {}
                camp_name = None
                if camps is not None:
                    camp = camps.get((participant.get('parent') or {}).get('_id'))
                    camp_name = (camp or {}).get('name') or ''
                ws.append(scale_export.scale_row(
                    form,
                    doc.get('fields') or {},
                    participant,
                    submitter_label(doc),
                    attachment_names(doc),
                    camp_name=camp_name,
                    report_id=doc.get('_id'),
                ))
            ws.freeze_panes = 'F2'  # keep ID + name + dates visible while scrolling
            return ws
        except Exception:
            pass
    write_generic_scale_sheet(ws, form, reports, participants, camps=camps)
    return ws


def report_submitter_ids(reports):
    """Contact _ids that appear as a report submitter (screeners/staff) -
    these are never screening subjects, so they are kept out of participant
    listings even if their contact sits under the camp."""
    ids = set()
    for doc in reports:
        contact = doc.get('contact')
        if isinstance(contact, dict) and contact.get('_id'):
            ids.add(contact['_id'])
    return ids


REPORT_BASE_URL = 'https://cht.kobosatorufoundation.org/#/reports/'


def extract_participant_columns(p, camp_doc=None, all_camps=False, scales_done=None):
    row = {
        'CHT ID': p.get('_id'),
        'Date registered (CHT)': iso_datetime(p.get('reported_date')),
        'Date': iso_datetime(p.get('registration_date')),
        'Ward / Dept No': p.get('ward_dept_no'),
        'Hospital Name': p.get('hospital_name'),
        'Patient Name': p.get('name'),
        'CR (Case Record) No.': p.get('cr_no'),
        'Aadhaar (last 4 digits)': p.get('aadhaar_last4'),
        'Age': p.get('age'),
        'Sex': GENDER_LABELS.get(p.get('gender'), p.get('gender')),
        'Education': p.get('education_level'),
        'Occupation': p.get('occupation'),
        'Phone (patient / caregiver / attendant)': p.get('phone'),
        'Address / Locality - City / Town / Village / District': p.get('address'),
        'Department (referral)': p.get('referral_department'),
        'Referred by (doctor name)': p.get('referred_by'),
        'Health complaint & reason for referral': p.get('referral_complaint'),
        'Notes': p.get('notes'),
    }
    if all_camps or camp_doc:
        row['Camp'] = (camp_doc or {}).get('name') or ''
    row.update({
        'Registered by': p.get('created_by_name') or p.get('created_by_user'),
        'Scales completed': ', '.join(scales_done or []),
    })
    return row


def extract_assessment_visit_columns(r=None, visit_idx=None):
    if not r:
        return {
            'Visit #': None,
            'Date & time of assessment': None,
            'Administered by (staff name/s)': None,
            'Uploaded by': None,
            'Patient Name (Confirmed)': None,
            'CR (Case Record) No. (Confirmed)': None,
            'Type of consent obtained': None,
            'Consent given for screening': None,
            'Respondent name & relation': None,
            'Informant phone number': None,
            'Assessment scale used': None,
            'Score / result': None,
            'Interpretation': None,
            'Provisional diagnosis (post-assessment) / Clinical impression': None,
            'Follow-up / next visit date': None,
            'Notes (Assessment)': None,
            'Assessed by (staff - name & signature)': None,
            'Reviewed by (staff - name & signature)': None,
            'Entered in Kobo / CHT': None,
            'Quality control - checked by (name)': None,
            'Participant kobo / Record ID': None,
            'Quality control checking date': None,
            'Supporting Docs Caption': None,
            'Assessment Attachments': None,
            'Assessment Report ID': None,
        }
    fields = flatten(r.get('fields') or {})
    consent_mode = fields.get('g_consent.consent_mode')
    consent_given = fields.get('g_consent.consent_given')
    att = attachment_names(r)
    rep_id = r.get('_id')
    att_cell = f'=HYPERLINK("{REPORT_BASE_URL}{rep_id}", "View in CHT")' if (att and rep_id) else (', '.join(att) if att else None)
    return {
        'Visit #': visit_idx or 1,
        'Date & time of assessment': iso_datetime(fields.get('g_assessment.assessment_datetime') or r.get('reported_date')),
        'Administered by (staff name/s)': fields.get('g_assessment.administered_by'),
        'Uploaded by': fields.get('g_assessment.uploaded_by'),
        'Patient Name (Confirmed)': fields.get('g_patient.patient_name_confirm'),
        'CR (Case Record) No. (Confirmed)': fields.get('g_patient.patient_cr_confirm'),
        'Type of consent obtained': {'written': 'Written', 'verbal': 'Verbal'}.get(consent_mode, consent_mode),
        'Consent given for screening': {'yes': 'Yes', 'no': 'No'}.get(consent_given, consent_given),
        'Respondent name & relation': fields.get('g_informant.informant_name'),
        'Informant phone number': fields.get('g_informant.informant_phone'),
        'Assessment scale used': fields.get('g_assessment.g_scale.scale_name'),
        'Score / result': fields.get('g_assessment.g_scale.score'),
        'Interpretation': fields.get('g_assessment.g_scale.interpretation'),
        'Provisional diagnosis (post-assessment) / Clinical impression': fields.get('g_clinical.clinical_impression'),
        'Follow-up / next visit date': iso_datetime(fields.get('g_followup.next_visit_date')),
        'Notes (Assessment)': fields.get('g_followup.notes'),
        'Assessed by (staff - name & signature)': fields.get('g_followup.assessed_by'),
        'Reviewed by (staff - name & signature)': fields.get('g_followup.reviewed_by'),
        'Entered in Kobo / CHT': fields.get('g_office.entered_in_cht'),
        'Quality control - checked by (name)': fields.get('g_office.qc_checked_by'),
        'Participant kobo / Record ID': fields.get('g_office.record_id'),
        'Quality control checking date': iso_datetime(fields.get('g_office.qc_check_date')),
        'Supporting Docs Caption': fields.get('g_supporting_docs.upload_caption'),
        'Assessment Attachments': att_cell,
        'Assessment Report ID': rep_id,
    }


def build_camp_workbook(wb, camp_id, camp, participants, reports, camps=None):
    """Participants sheet + one sheet per scale completed.

    camp_id=None exports EVERY participant across every camp (all-workbook
    scope) and adds a 'Camp' column so rows stay attributable.
    """
    all_camps = camp_id is None
    staff = report_submitter_ids(reports)
    members = {pid: p for pid, p in participants.items()
               if pid not in staff
               and (all_camps
                    or (p.get('parent') or {}).get('_id') == camp_id)}
    by_form = scale_reports_by_participant(reports, set(members))
    active_forms = [f for f, docs in by_form.items() if docs]
    order = get_scale_order(active_forms)

    # Collect assessment visits by participant
    assessment_visits_by_pid = {}
    for r in reports:
        if r.get('form') == 'assessment_visit':
            pid = subject_id_of(r)
            if pid and pid in members:
                assessment_visits_by_pid.setdefault(pid, []).append(r)
    for visits in assessment_visits_by_pid.values():
        visits.sort(key=lambda d: (d.get('reported_date') or 0, d.get('_id') or ''))

    rows = []
    for pid, p in members.items():
        scales_done = [scale_export.sheet_title(f) for f in order
                       if any(subject_id_of(d) == pid for d in by_form.get(f, []))]
        camp_doc = (camps or {}).get((p.get('parent') or {}).get('_id')) if all_camps else None
        base_part = extract_participant_columns(p, camp_doc=camp_doc, all_camps=all_camps, scales_done=scales_done)

        p_visits = assessment_visits_by_pid.get(pid, [])
        if not p_visits:
            combined = dict(base_part)
            combined.update(extract_assessment_visit_columns(None))
            rows.append(combined)
        else:
            for v_idx, visit_doc in enumerate(p_visits, 1):
                combined = dict(base_part)
                combined.update(extract_assessment_visit_columns(visit_doc, visit_idx=v_idx))
                rows.append(combined)

    if all_camps:
        rows.sort(key=lambda r: (r.get('Camp') or '',
                                 r.get('CR (Case Record) No.') or r.get('Participant ID') or '',
                                 r.get('Patient Name') or r.get('Name') or '',
                                 r.get('Visit #') or 0))
    else:
        rows.sort(key=lambda r: (r.get('CR (Case Record) No.') or r.get('Participant ID') or '',
                                 r.get('Patient Name') or r.get('Name') or '',
                                 r.get('Visit #') or 0))
    write_sheet(wb, 'Participants', rows)

    for form in order:
        if by_form.get(form):
            write_scale_sheet(wb, form, by_form[form], participants,
                              camps=camps if all_camps else None)
    return by_form


def write_participant_sheet(wb, participant, camp, scales_done, assessment_visits=None):
    """Field | Value summary sheet for one participant."""
    rows = [
        {'Field': 'CHT ID', 'Value': participant.get('_id')},
        {'Field': 'Date registered (CHT)', 'Value': iso_datetime(participant.get('reported_date'))},
        {'Field': 'Date', 'Value': iso_datetime(participant.get('registration_date'))},
        {'Field': 'Ward / Dept No', 'Value': participant.get('ward_dept_no')},
        {'Field': 'Hospital Name', 'Value': participant.get('hospital_name')},
        {'Field': 'Patient Name', 'Value': participant.get('name')},
        {'Field': 'CR (Case Record) No.', 'Value': participant.get('cr_no')},
        {'Field': 'Aadhaar (last 4 digits)', 'Value': participant.get('aadhaar_last4')},
        {'Field': 'Age', 'Value': participant.get('age')},
        {'Field': 'Sex', 'Value': GENDER_LABELS.get(participant.get('gender'),
                                                    participant.get('gender'))},
        {'Field': 'Education', 'Value': participant.get('education_level')},
        {'Field': 'Occupation', 'Value': participant.get('occupation')},
        {'Field': 'Phone (patient / caregiver / attendant)', 'Value': participant.get('phone')},
        {'Field': 'Address / Locality - City / Town / Village / District', 'Value': participant.get('address')},
        {'Field': 'Department (referral)', 'Value': participant.get('referral_department')},
        {'Field': 'Referred by (doctor name)', 'Value': participant.get('referred_by')},
        {'Field': 'Health complaint & reason for referral', 'Value': participant.get('referral_complaint')},
        {'Field': 'Notes', 'Value': participant.get('notes')},
        {'Field': 'Camp', 'Value': camp.get('name') if camp else None},
        {'Field': 'Registered by',
         'Value': participant.get('created_by_name') or participant.get('created_by_user')},
        {'Field': 'Scales completed', 'Value': ', '.join(scales_done)},
    ]
    if assessment_visits:
        for idx, r in enumerate(assessment_visits, 1):
            fields = flatten(r.get('fields') or {})
            rows.append({'Field': f'--- Assessment Visit #{idx} ---', 'Value': ''})
            rows.append({'Field': 'Date & time of assessment',
                         'Value': iso_datetime(fields.get('g_assessment.assessment_datetime') or r.get('reported_date'))})
            if fields.get('g_assessment.administered_by'):
                rows.append({'Field': 'Administered by (staff name/s)', 'Value': fields.get('g_assessment.administered_by')})
            if fields.get('g_assessment.uploaded_by'):
                rows.append({'Field': 'Uploaded by', 'Value': fields.get('g_assessment.uploaded_by')})
            if fields.get('g_patient.patient_name_confirm'):
                rows.append({'Field': 'Patient Name (Confirmed)', 'Value': fields.get('g_patient.patient_name_confirm')})
            if fields.get('g_patient.patient_cr_confirm'):
                rows.append({'Field': 'CR (Case Record) No. (Confirmed)', 'Value': fields.get('g_patient.patient_cr_confirm')})
            if fields.get('g_consent.consent_mode'):
                rows.append({'Field': 'Type of consent obtained',
                             'Value': {'written': 'Written', 'verbal': 'Verbal'}.get(fields.get('g_consent.consent_mode'), fields.get('g_consent.consent_mode'))})
            if fields.get('g_consent.consent_given'):
                rows.append({'Field': 'Consent given for screening',
                             'Value': {'yes': 'Yes', 'no': 'No'}.get(fields.get('g_consent.consent_given'), fields.get('g_consent.consent_given'))})
            if fields.get('g_informant.informant_name'):
                rows.append({'Field': 'Respondent name & relation', 'Value': fields.get('g_informant.informant_name')})
            if fields.get('g_informant.informant_phone'):
                rows.append({'Field': 'Informant phone number', 'Value': fields.get('g_informant.informant_phone')})
            if fields.get('g_assessment.g_scale.scale_name'):
                rows.append({'Field': 'Assessment scale used', 'Value': fields.get('g_assessment.g_scale.scale_name')})
            if fields.get('g_assessment.g_scale.score'):
                rows.append({'Field': 'Score / result', 'Value': fields.get('g_assessment.g_scale.score')})
            if fields.get('g_assessment.g_scale.interpretation'):
                rows.append({'Field': 'Interpretation', 'Value': fields.get('g_assessment.g_scale.interpretation')})
            if fields.get('g_clinical.clinical_impression'):
                rows.append({'Field': 'Provisional diagnosis (post-assessment) / Clinical impression', 'Value': fields.get('g_clinical.clinical_impression')})
            if fields.get('g_followup.next_visit_date'):
                rows.append({'Field': 'Follow-up / next visit date', 'Value': iso_datetime(fields.get('g_followup.next_visit_date'))})
            if fields.get('g_followup.notes'):
                rows.append({'Field': 'Notes (Assessment)', 'Value': fields.get('g_followup.notes')})
            if fields.get('g_followup.assessed_by'):
                rows.append({'Field': 'Assessed by (staff - name & signature)', 'Value': fields.get('g_followup.assessed_by')})
            if fields.get('g_followup.reviewed_by'):
                rows.append({'Field': 'Reviewed by (staff - name & signature)', 'Value': fields.get('g_followup.reviewed_by')})
            if fields.get('g_office.entered_in_cht'):
                rows.append({'Field': 'Entered in Kobo / CHT', 'Value': fields.get('g_office.entered_in_cht')})
            if fields.get('g_office.qc_checked_by'):
                rows.append({'Field': 'Quality control - checked by (name)', 'Value': fields.get('g_office.qc_checked_by')})
            if fields.get('g_office.record_id'):
                rows.append({'Field': 'Participant kobo / Record ID', 'Value': fields.get('g_office.record_id')})
            if fields.get('g_office.qc_check_date'):
                rows.append({'Field': 'Quality control checking date', 'Value': iso_datetime(fields.get('g_office.qc_check_date'))})
            if fields.get('g_supporting_docs.upload_caption'):
                rows.append({'Field': 'Supporting Docs Caption', 'Value': fields.get('g_supporting_docs.upload_caption')})
            att = attachment_names(r)
            if att:
                rep_id = r.get('_id')
                rows.append({'Field': 'Assessment Attachments', 'Value': f'=HYPERLINK("{REPORT_BASE_URL}{rep_id}", "View in CHT")' if rep_id else ', '.join(att)})
            if r.get('_id'):
                rows.append({'Field': 'Assessment Report ID', 'Value': r.get('_id')})
    rows = [r for r in rows if r['Value'] not in (None, '') or r['Field'].startswith('---')]
    return write_sheet(wb, 'Participant', rows)


def write_participant_scale_sheet(wb, form, doc, participant):
    """Sectioned human-readable sheet for one participant's completed scale."""
    ws = wb.create_sheet(scale_export.sheet_title(form))
    ws.append([scale_export.scale_full_title(form)])
    reported = iso_datetime(doc.get('reported_date'))
    if reported:
        ws.append([f'Report submitted: {reported}'])
    ws.append([])
    sections = None
    try:
        sections = scale_export.scale_sections(
            form,
            doc.get('fields') or {},
            participant,
            submitter_label(doc),
            attachment_names(doc),
            report_id=doc.get('_id'),
        )
    except Exception:
        pass
    if sections:
        for section, rows in sections:
            ws.append([section])
            ws.append(['Question / Field', 'Answer / Value', 'Score'])
            for label, value, score in rows:
                ws.append([label, value, score if score is not None else ''])
            ws.append([])
    else:
        ws.append(['Field', 'Value'])
        flat_fields = flatten(doc.get('fields') or {})
        for k, v in sorted(flat_fields.items()):
            if not k.startswith('inputs.') and not k.startswith('meta.'):
                ws.append([k, json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v])
    return ws


def build_participant_report_workbook(wb, participant_id, participants, camps, reports):
    """Participant sheet + one sectioned sheet per completed scale."""
    participant = participants.get(participant_id)
    if not participant:
        raise SystemExit(f'participant not found: {participant_id}')
    camp = camps.get((participant.get('parent') or {}).get('_id'))
    by_form = scale_reports_by_participant(reports, {participant_id})
    active_forms = [f for f, docs in by_form.items() if docs]
    order = get_scale_order(active_forms)
    scales_done = [scale_export.sheet_title(f) for f in order if by_form.get(f)]

    p_assessment_visits = [d for d in reports
                           if d.get('form') == 'assessment_visit' and subject_id_of(d) == participant_id]
    p_assessment_visits.sort(key=lambda d: (d.get('reported_date') or 0, d.get('_id') or ''))

    write_participant_sheet(wb, participant, camp, scales_done, assessment_visits=p_assessment_visits)
    for form in order:
        for doc in by_form.get(form, []):  # a re-completed scale gets its own sheet copy
            write_participant_scale_sheet(wb, form, doc, participant)
    return by_form


def write_sheet(wb, name, rows):
    """Schema-generic sheet writer: columns = union of all row keys."""
    ws = wb.create_sheet(name)
    if not rows:
        ws.append(['_id'])
        return ws
    columns = list(rows[0].keys())
    for row in rows[1:]:
        for key in row.keys():
            if key not in columns:
                columns.append(key)
    ws.append(columns)
    for row in rows:
        ws.append([json.dumps(row.get(col), ensure_ascii=False)
                   if isinstance(row.get(col), (dict, list))
                   else row.get(col) for col in columns])
    ws.freeze_panes = 'A2'
    return ws


def build_workbook(scope, db):
    docs = db.fetch_all_docs()
    camps, participants = classify_contacts(docs)
    reports = [d for d in docs if is_emr_report(d)]

    # Ensure form XMLs exist locally for any form present in reports
    if hasattr(db, 'fetch_attachment'):
        for doc in reports:
            f = doc.get('form')
            if f and f not in scale_meta.SCALES:
                xml_path = os.path.join(scale_meta.FORMS_DIR, f'{f}.xml')
                if not os.path.isfile(xml_path):
                    xml_bytes = db.fetch_attachment(f'form:{f}', 'xml')
                    if xml_bytes:
                        try:
                            os.makedirs(scale_meta.FORMS_DIR, exist_ok=True)
                            with open(xml_path, 'wb') as fp:
                                fp.write(xml_bytes)
                        except Exception:
                            pass

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    if scope == 'all':
        camp_rows = [flatten_contact_columns(d, {}) for d in camps.values()]
        part_rows = [flatten_contact_columns(d, {'camp_id': (d.get('parent') or {}).get('_id')})
                     for d in participants.values()]
        report_rows = collect_report_rows(reports, participants, camps)
        write_sheet(wb, 'Camps', camp_rows)
        write_sheet(wb, 'Participants', part_rows)
        write_sheet(wb, 'Reports', report_rows)

    elif scope.startswith('camp='):
        camp_id = scope.split('=', 1)[1]
        camp = camps.get(camp_id)
        if not camp:
            raise SystemExit(f'camp not found: {camp_id}')
        camp_rows = [flatten_contact_columns(camp, {})]
        part_rows = [flatten_contact_columns(d, {'camp_id': camp_id})
                     for d in participants.values()
                     if (d.get('parent') or {}).get('_id') == camp_id]
        camp_participant_ids = {r['_id'] for r in part_rows}
        report_rows = collect_report_rows(reports, participants, camps)
        report_rows = [r for r in report_rows if r['camp_id'] == camp_id]
        write_sheet(wb, 'Camp', camp_rows)
        write_sheet(wb, 'Participants', part_rows)
        write_sheet(wb, 'Reports', report_rows)

    elif scope.startswith('participants_in_camp='):
        camp_id = scope.split('=', 1)[1]
        part_rows = [flatten_contact_columns(d, {'camp_id': camp_id})
                     for d in participants.values()
                     if (d.get('parent') or {}).get('_id') == camp_id]
        report_rows = collect_report_rows(reports, participants, camps)
        report_rows = [r for r in report_rows if r['camp_id'] == camp_id]
        write_sheet(wb, 'Participants', part_rows)
        write_sheet(wb, 'Reports', report_rows)

    elif scope.startswith('participant='):
        participant_id = scope.split('=', 1)[1]
        participant = participants.get(participant_id)
        if not participant:
            raise SystemExit(f'participant not found: {participant_id}')
        part_rows = [flatten_contact_columns(
            participant, {'camp_id': (participant.get('parent') or {}).get('_id')})]
        report_rows = collect_report_rows(reports, participants, camps)
        report_rows = [r for r in report_rows if r['participant_id'] == participant_id]
        write_sheet(wb, 'Participant', part_rows)
        write_sheet(wb, 'Reports', report_rows)

    elif scope == 'all-workbook':
        build_camp_workbook(wb, None, None, participants, reports, camps=camps)

    elif scope.startswith('camp-workbook='):
        camp_id = scope.split('=', 1)[1]
        camp = camps.get(camp_id)
        if not camp:
            raise SystemExit(f'camp not found: {camp_id}')
        build_camp_workbook(wb, camp_id, camp, participants, reports)

    elif scope.startswith('participant-report='):
        build_participant_report_workbook(
            wb, scope.split('=', 1)[1], participants, camps, reports)

    else:
        raise SystemExit(f'unknown scope: {scope}')

    for ws in wb.worksheets:
        autosize_columns(ws)
    return wb


def autosize_columns(ws, min_width=10, max_width=70):
    """Set each column width to fit its longest cell (clamped)."""
    from openpyxl.utils import get_column_letter
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            longest = max((len(line) for line in str(cell.value).splitlines()),
                          default=0)
            widths[cell.column] = max(widths.get(cell.column, 0), longest)
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = \
            max(min_width, min(max_width, width + 2))


def default_output_name(scope, wb, db=None):
    stamp = timestamp_stamp()
    if scope == 'all':
        return f'satoru-emr-all-data-{stamp}.xlsx'
    if scope.startswith('camp='):
        camp_sheet = wb['Camp']
        name = None
        header = [c.value for c in camp_sheet[1]]
        if 'name' in header:
            name = camp_sheet[2][header.index('name')].value
        return f'satoru-emr-camp-{sanitize_filename(name, "camp")}-{stamp}.xlsx'
    if scope.startswith('participants_in_camp='):
        camp_id = scope.split('=', 1)[1]
        camp_name = None
        if db:
            camp_doc = db.fetch_doc(camp_id)
            camp_name = camp_doc.get('name') if camp_doc else None
        return (f'satoru-emr-camp-{sanitize_filename(camp_name or camp_id, "camp")}'
                f'-participants-{stamp}.xlsx')
    if scope.startswith('participant='):
        part_sheet = wb['Participant']
        header = [c.value for c in part_sheet[1]]
        cr = part_sheet[2][header.index('cr_no')].value if 'cr_no' in header else None
        name = part_sheet[2][header.index('name')].value if 'name' in header else None
        label = cr or name or scope.split('=', 1)[1]
        return f'satoru-emr-participant-{sanitize_filename(label, "participant")}-{stamp}.xlsx'
    if scope == 'all-workbook':
        return f'satoru-emr-all-participants-scales-{stamp}.xlsx'
    if scope.startswith('camp-workbook='):
        camp_id = scope.split('=', 1)[1]
        camp_name = None
        if db:
            camp_doc = db.fetch_doc(camp_id)
            camp_name = camp_doc.get('name') if camp_doc else None
        return (f'satoru-emr-camp-{sanitize_filename(camp_name or camp_id, "camp")}'
                f'-scales-{stamp}.xlsx')
    if scope.startswith('participant-report='):
        part_id = scope.split('=', 1)[1]
        label = part_id
        if db:
            part_doc = db.fetch_doc(part_id)
            if part_doc:
                label = part_doc.get('cr_no') or part_doc.get('name') or part_id
        return (f'satoru-emr-participant-{sanitize_filename(label, "participant")}'
                f'-report-{stamp}.xlsx')
    raise SystemExit(f'unknown scope: {scope}')


def main():
    parser = argparse.ArgumentParser(description='Satoru EMR -> Excel exporter (read-only)')
    parser.add_argument('--url', required=True, help='CouchDB URL, e.g. https://user:pass@host:port/medic')
    parser.add_argument('--scope', default='all',
                        help='all | all-workbook | camp=<campId> | '
                             'participants_in_camp=<campId> | '
                             'participant=<participantId> | camp-workbook=<campId> | '
                             'participant-report=<participantId>')
    parser.add_argument('--out', default='.', help='output directory (default: current directory)')
    parser.add_argument('--verify', action='store_true', help='verify TLS certificates (default: off for self-signed)')
    args = parser.parse_args()

    db = EmrDb(args.url, verify=args.verify)
    wb = build_workbook(args.scope, db)
    filename = default_output_name(args.scope, wb, db)
    os.makedirs(args.out, exist_ok=True)
    path = os.path.join(args.out, filename)
    wb.save(path)
    for ws in wb.worksheets:
        print(f'  {ws.title}: {ws.max_row - 1} rows x {ws.max_column} cols')
    print(f'wrote {path}')


if __name__ == '__main__':
    main()
