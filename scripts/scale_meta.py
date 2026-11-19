"""scale_meta.py - single source of truth for scale report metadata.

Parses the actual form definitions (forms/app/*.xml) - the same files cht-conf
uploads - and derives, for each scale (ddst, dst, moca_assessment, vineland):

  * field labels        (from XForm body <label> / itext English translations)
  * group labels        (age bands / domains / sections, in form order)
  * choice maps         (per field: stored value -> human label, from the
                         form's own <instance> lists and itext blocks)
  * results             (calculated scoring/result fields, in display order)
  * report_display config (consumed by the cht-core report patch:
                         hide internal fields, join CA display, note values,
                         per-item score joins, empty suppression)

Nothing here invents question text or choice meanings - everything is read
from the form XML. Python 3.8+, stdlib only.
"""

import json
import os
import threading
import re
import xml.etree.ElementTree as ET

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FORMS_DIR = os.path.join(REPO_ROOT, 'forms', 'app')

KNOWN_TITLES = {
    'ddst': 'Denver Developmental Screening Test II (DDST-II)',
    'dst': 'Developmental Screening Test (DST)',
    'moca': 'MOCA (Integrated)',
    'moca_assessment': 'Montreal Cognitive Assessment (MoCA)',
    'moca_b': 'MoCA-Basic (MoCA-B)',
    'moca_telugu': 'ICMR-NCTB Telugu MoCA (v7.1)',
    'vineland': 'Vineland-style Social Maturity Scale (VSMS)',
    'cog5': 'CoG5 - Memory Screening Battery',
    'mse': 'Mental State Examination (MSE)',
    'phq9': 'PHQ-9 Depression Screening',
    'ad8': 'AD8 Dementia Screening Interview',
    'easi': 'Everyday Abilities Scale for India (EASI)',
    'gds15': 'Geriatric Depression Scale (GDS-15)',
    'assessment_visit': 'Assessment Visit',
}


def _extract_xml_title(path):
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        for head in root:
            if head.tag.endswith('head'):
                for child in head:
                    if child.tag.endswith('title') and child.text:
                        return child.text.strip()
    except Exception:
        pass
    return None


class ScalesRegistry(dict):
    """Dynamic scale registry that auto-discovers all app forms in FORMS_DIR
    while preserving human titles for known scales."""

    def __init__(self):
        super().__init__()
        for k, v in KNOWN_TITLES.items():
            self[k] = {'file': f'{k}.xml', 'title': v}

    def _refresh(self):
        if not os.path.isdir(FORMS_DIR):
            return
        for fname in os.listdir(FORMS_DIR):
            if fname.endswith('.xml') and not fname.startswith('.'):
                form_id = fname[:-4]
                if not dict.__contains__(self, form_id):
                    path = os.path.join(FORMS_DIR, fname)
                    title = _extract_xml_title(path) or KNOWN_TITLES.get(form_id) or form_id.replace('_', ' ').title()
                    dict.__setitem__(self, form_id, {'file': fname, 'title': title})

    def __contains__(self, key):
        self._refresh()
        return dict.__contains__(self, key) or key in KNOWN_TITLES

    def __getitem__(self, key):
        self._refresh()
        if not dict.__contains__(self, key):
            path = os.path.join(FORMS_DIR, f'{key}.xml')
            title = None
            if os.path.isfile(path):
                title = _extract_xml_title(path)
            title = title or KNOWN_TITLES.get(key) or key.replace('_', ' ').title()
            dict.__setitem__(self, key, {'file': f'{key}.xml', 'title': title})
        return dict.__getitem__(self, key)

    def get(self, key, default=None):
        try:
            return self[key]
        except Exception:
            return default

    def __iter__(self):
        self._refresh()
        return dict.__iter__(self)

    def keys(self):
        self._refresh()
        return dict.keys(self)

    def values(self):
        self._refresh()
        return dict.values(self)

    def items(self):
        self._refresh()
        return dict.items(self)

    def get(self, key, default=None):
        self._refresh()
        return dict.get(self, key, default)


SCALES = ScalesRegistry()

# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------


def _local(tag):
    """Strip any xml namespace: '{ns}name' -> 'name'."""
    return tag.rsplit('}', 1)[-1]


def _children(elem, name):
    return [c for c in elem if _local(c.tag) == name]


def _child(elem, name):
    for c in elem:
        if _local(c.tag) == name:
            return c
    return None


def _text(elem):
    return ''.join(elem.itertext()).strip() if elem is not None else ''


def parse_form_xml(path):
    """Parse a form XML into raw pieces needed downstream."""
    tree = ET.parse(path)
    html = tree.getroot()
    head = _child(html, 'head')
    body = _child(html, 'body')
    model = None
    for cand in _children(head, 'model'):
        model = cand
        break

    # -- itext: first translation (en) block, id -> text ------------------
    itext = {}
    itext_elem = _child(model, 'itext') if model is not None else None
    if itext_elem is not None:
        en = None
        for tr in _children(itext_elem, 'translation'):
            if tr.get('lang') == 'en':
                en = tr
                break
        if en is None and _children(itext_elem, 'translation'):
            en = _children(itext_elem, 'translation')[0]
        if en is not None:
            for t in _children(en, 'text'):
                itext[t.get('id')] = _text(t)

    # -- named choice instances: instance id -> [(name, label)] -----------
    choice_lists = {}
    for inst in _children(model, 'instance'):
        inst_id = inst.get('id')
        if not inst_id or inst_id in ('contact-summary', 'user-contact-summary'):
            continue
        root = _child(inst, 'root')
        if root is None:
            continue
        items = []
        for item in _children(root, 'item'):
            name = _text(_child(item, 'name'))
            if not name:
                continue
            label_elem = _child(item, 'label')
            itext_id_elem = _child(item, 'itextId')
            if label_elem is not None:
                label = _text(label_elem)
            elif itext_id_elem is not None:
                label = itext.get(_text(itext_id_elem), name)
            else:
                label = name
            items.append((name, label))
        if items:
            choice_lists[inst_id] = items

    # -- instance field order (model/instance/data) -----------------------
    field_order = []

    def walk_instance(elem, prefix):
        for c in elem:
            name = _local(c.tag)
            if name in ('meta', 'inputs'):
                continue
            path = prefix + name
            if len(list(c)):
                # has children: a group
                field_order.append(('group', path))
                walk_instance(c, path + '.')
            else:
                field_order.append(('leaf', path))

    data = None
    for inst in _children(model, 'instance'):
        if not inst.get('id'):
            inner = _child(inst, 'data')
            data = inner if inner is not None else inst
            break
    if data is not None:
        walk_instance(data, '')

    # -- body: groups and fields, in display order ------------------------
    body_groups = []   # [(ref, group_label, [field dicts])]
    body_loose = []    # top-level fields (ref, label, hint, itemset)

    def _resolve_label(elem):
        lbl_elem = _child(elem, 'label')
        label = _text(lbl_elem)
        if not label and lbl_elem is not None:
            lref = lbl_elem.get('ref') or ''
            m = re.search(r"jr:itext\('([^']+)'\)", lref)
            if m:
                label = itext.get(m.group(1), '')
        if label:
            label = re.sub(r'\*\*|__', '', label).strip().rstrip(':').strip()
        return label

    def _resolve_hint(elem):
        hint_elem = _child(elem, 'hint')
        hint = _text(hint_elem)
        if not hint and hint_elem is not None:
            href = hint_elem.get('ref') or ''
            m = re.search(r"jr:itext\('([^']+)'\)", href)
            if m:
                hint = itext.get(m.group(1), '')
        if hint:
            hint = re.sub(r'\*\*|__', '', hint).strip()
        return hint

    def parse_field(elem):
        ref = elem.get('ref') or ''
        path = ref.split('/')[-1] if ref else ''
        # full path relative to /data: replace '/' separators
        parts = [p for p in ref.split('/') if p and p != 'data']
        fpath = '.'.join(parts)
        label = _resolve_label(elem)
        hint = _resolve_hint(elem)
        itemset = None
        iset = _child(elem, 'itemset')
        if iset is not None:
            nodeset = iset.get('nodeset') or ''
            m = re.search(r"instance\('([^']+)'\)", nodeset)
            itemset = m.group(1) if m else None
        return {
            'ref': ref,
            'path': fpath,
            'name': path,
            'label': label,
            'hint': hint,
            'itemset': itemset,
        }

    def parse_group(group_elem, depth=0):
        ref = group_elem.get('ref') or ''
        parts = [p for p in ref.split('/') if p and p != 'data']
        gpath = '.'.join(parts)
        glabel = _resolve_label(group_elem)
        children = []
        for sub in group_elem:
            stag = _local(sub.tag)
            if stag in ('select1', 'select', 'input', 'upload', 'repeat', 'trigger'):
                children.append(parse_field(sub))
            elif stag == 'group':
                sub_g = parse_group(sub, depth + 1)
                children.append({'__group__': sub_g})
        return {
            'ref': ref,
            'path': gpath,
            'label': glabel,
            'depth': depth,
            'fields': children,
        }

    all_groups = [parse_group(elem, 0) for elem in body if _local(elem.tag) == 'group']

    # fields that sit directly in <h:body> outside any group
    direct_children = []
    for elem in body:
        tag = _local(elem.tag)
        if tag in ('select1', 'select', 'input', 'upload'):
            direct_children.append(parse_field(elem))

    title = _text(_child(head, 'title'))

    return {
        'title': title,
        'itext': itext,
        'choice_lists': choice_lists,
        'field_order': field_order,
        'groups': all_groups,
        'body_loose': direct_children,
    }


# ---------------------------------------------------------------------------
# Flatten body structure into ordered field list with group membership
# ---------------------------------------------------------------------------


def flatten_fields(parsed):
    """Return ordered list of dicts:
       {'type': 'group'|'field', 'path':..., 'label':..., 'itemset':...}
    Groups carry their nested order; fields keep form order."""
    out = []

    def emit_group(g):
        if g['path']:
            out.append({'type': 'group', 'path': g['path'],
                        'label': g['label']})
        for f in g['fields']:
            if '__group__' in f:
                emit_group(f['__group__'])
            else:
                out.append({'type': 'field', 'path': f['path'],
                            'label': f['label'], 'hint': f['hint'],
                            'itemset': f['itemset']})

    for g in parsed['groups']:
        emit_group(g)
    for f in parsed['body_loose']:
        out.append({'type': 'field', 'path': f['path'], 'label': f['label'],
                    'hint': f['hint'], 'itemset': f['itemset']})
    return out


def build_choice_maps(parsed, flat):
    """Per-field choice maps from the form's own lists.

    Resolution order: explicit <item> children (none in these forms),
    then <itemset instance('x')> -> model instance x (labels may come from
    itext). Only fields that reference a choice list get a map.
    """
    maps = {}
    for f in flat:
        if f['type'] != 'field' or not f.get('itemset'):
            continue
        items = parsed['choice_lists'].get(f['itemset'])
        if not items:
            continue
        maps[f['path']] = {name: label for name, label in items}
    return maps


# ---------------------------------------------------------------------------
# Per-scale report configuration
# ---------------------------------------------------------------------------


def _all_field_paths(parsed):
    """Every leaf path stored in the doc (from the model instance), covering
    calculated fields that never appear in the form body."""
    paths = []
    for kind, path in parsed.get('field_order', []):
        if kind == 'leaf':
            paths.append(path)
    return paths


def _hidden_scale_fields(flat, form, parsed=None):
    """Internal/calculation field paths to hide, derived programmatically.

    Scans both body fields and model-instance leaves so calculated fields
    (item_N_score, band_N_*, *_delayed) are covered too.
    """
    hidden = {
        'patient_uuid',
        'screened_by',
        'scoring_version',
        'result_complete',
        'result_incomplete',
        'result_invalid',
        # DOB warning is surfaced conditionally by the report layer, not as a
        # generic field (see DOB_WARNING); keep it out of the raw field list.
        'g_child.dob_warning',
        'g_meta.dob_warning',
    }
    every = [item['path'] for item in flat if item['type'] == 'field']
    if parsed is not None:
        every = every + [p for p in _all_field_paths(parsed) if p not in every]
    if form in ('vineland', 'dst'):
        hidden |= {
            'g_child.ca_days',
            'g_child.ca_months_dec',
            'g_child.ca_band',
            'g_child.age_valid',
            'g_child.show_all',
            'show_all',
            'ca_band',
            'score_status',
            # readonly "**y m** ( days)." display string; the joined
            # "Chronological age" row already carries the readable value
            'g_child.ca_display',
        }
    if form == 'vineland':
        hidden |= {
            'g_child.ca_months_floor',
            'g_child.ca_years_display',
            'g_child.ca_rem_months',
        }
        # per-item score fields are joined into their question rows instead
        for path in every:
            if re.match(r'item_\d+_score$', path):
                hidden.add(path)
    if form == 'dst':
        # 18 intermediate band calculation blocks (band_N_*)
        for path in every:
            if re.match(r'band_\d+_(passed|failed|not_assessed|earned)', path):
                hidden.add(path)
    if form == 'ddst':
        hidden |= {
            'g_meta.ca_days',
            'g_meta.ca_months_dec',
            'g_meta.ca_band',
            'g_meta.show_all',
            'g_meta.ca_note',
            'caution_domains',
            'any_suspect',
            'result_note',  # composed note; duplicated by the results rows
        }
        # per-item and per-domain delayed flags
        for path in every:
            if path.endswith('_delayed'):
                hidden.add(path)
    if form == 'moca_assessment':
        hidden |= {
            'patient_sex',
            'patient_age',
            'patient_phone',
            'patient_education',
            'moca_delayed_total',   # duplicate of moca_memory_total
            'memory_words',
        }
        for path in every:
            if path.endswith('_inst'):
                hidden.add(path)
    if form == 'cog5':
        hidden |= {
            'patient_sex',
            'patient_age',
            'patient_phone',
            'patient_education',
            'patient_id',
            'patient_name',
            'screened_by_name',
            'grp_cog5_scale.calc_free_score',
            'grp_cog5_scale.calc_cued_score',
            'grp_cog5_scale.score_pmis_total',
            'grp_cog5_scale.pmis_flag',
            'grp_cog5_scale.grp_mcr.mcr_cog_flag',
            'grp_cog5_scale.grp_timed_walk.mcr_motor_flag',
            'g_participant.score_education',
            'grp_cog5_scale.grp_picture_based_screening_test.score_pbst',
            'grp_cog5_scale.grp_risk_factors.score_dementia_risk',
            'grp_cog5_scale.grp_picture_based_memory_impairment.pmis_image',
            'grp_cog5_scale.grp_picture_based_screening_test.triangle_thing',
            'grp_cog5_scale.grp_timed_walk.walk_note',
            'grp_cog5_scale.grp_timed_walk.walk_intro',
            'grp_cog5_scale.results_summary',
            'grp_cog5_scale.display_pmis',
            'grp_cog5_scale.display_pbst',
            'grp_cog5_scale.display_mcr',
            'grp_cog5_scale.display_risk',
            'grp_cog5_scale.display_triage',
            'grp_cog5_scale.display_overall',
        }
        for path in every:
            if any(path.endswith('.' + s) or path == s for s in (
                'pmis_image', 'triangle_thing', 'walk_note', 'walk_intro',
                'results_summary', 'display_pmis', 'display_pbst', 'display_mcr',
                'display_risk', 'display_triage', 'display_overall',
                'calc_free_score', 'calc_cued_score', 'score_pmis_total', 'pmis_flag',
                'mcr_cog_flag', 'mcr_motor_flag', 'score_education', 'score_pbst',
                'score_dementia_risk', 'patient_id', 'patient_name', 'screened_by_name'
            )):
                hidden.add(path)
    if form in ('moca', 'moca_assessment', 'moca_b', 'moca_telugu'):
        hidden |= {
            'patient_sex',
            'patient_age',
            'patient_phone',
            'patient_education',
            'g_admin.score_summary_note',
            'g_edu_lang_informant.education_band_note',
            'g_edu_lang_informant.band_a_rebook_note',
            'te_g_memory.te_memory_words',
            'mb_g_memory.mb_memory_words',
            'm8_g_memory.m8_memory_words',
            'g_mandatory_history.mh_inst',
            'te_raw_total',
            'te_total_score',
            'mb_raw_total',
            'm8_raw_total',
            'mb_visuoperception_count',
            'mb_attention_white_score',
            'mb_attention_black_score',
            'moca_delayed_total',
            'memory_words',
            'g_admin.moca_raw_total',
            'g_admin.moca_total_score',
            'g_admin.te_raw_total',
            'g_admin.te_total_score',
            'g_admin.mb_raw_total',
            'g_admin.mb_total_score',
        }
        for path in every:
            if path.endswith('_inst') or '_inst_' in path or '.banner_' in path:
                hidden.add(path)
            if any(path.endswith('.' + s) or path == s for s in (
                'te_raw_total', 'te_total_score', 'mb_raw_total', 'm8_raw_total',
                'mb_visuoperception_count', 'mb_attention_white_score', 'mb_attention_black_score'
            )):
                hidden.add(path)
    return sorted(hidden)


# Static note fields: stored value is empty; the meaningful text is the
# form label. build_report_display resolves the actual label text from the
# parsed form so no wording is duplicated or invented here.
#
# DOB-warning fields are deliberately NOT here: their relevance in the form
# is "DOB missing or age invalid", so forcing their text onto every report
# printed a false "WARNING" on normal assessments. They are surfaced only
# when actually triggered - see scale_export._dob_warning_active and the
# cht-core patch's equivalent check.
NOTE_VALUES = {
    'vineland': {
        'g_intro.vineland_disclaimer': None,
    },
    'dst': {
        'g_intro.dst_disclaimer': None,
    },
    'ddst': {},
    'moca': {},
    'moca_assessment': {},
}

# Path of the DOB-warning field per form, and the condition fields that make
# it relevant. Empty DOB or age_valid != 'yes' => the warning applies.
DOB_WARNING = {
    'vineland': {'field': 'g_child.dob_warning',
                 'dob': 'g_child.child_dob', 'age_valid': 'g_child.age_valid'},
    'dst': {'field': 'g_child.dob_warning',
            'dob': 'g_child.child_dob', 'age_valid': 'g_child.age_valid'},
    'ddst': {'field': 'g_meta.dob_warning',
             'dob': 'g_meta.child_dob', 'age_valid': 'g_meta.age_valid'},
    'moca': None,
    'moca_assessment': None,
}

# Short display labels for note-value fields (their VALUE is the full form
# text, e.g. the disclaimer sentence). Without this, the label and value
# would both show the same long text. Purely presentational UI wording.
NOTE_LABELS = {
    'vineland': {
        'g_intro.vineland_disclaimer': 'Disclaimer',
        'g_child.dob_warning': 'Date of birth warning',
    },
    'dst': {
        'g_intro.dst_disclaimer': 'Disclaimer',
        'g_child.dob_warning': 'Date of birth warning',
    },
    'ddst': {
        'g_meta.dob_warning': 'Date of birth warning',
    },
    'moca': {},
    'moca_assessment': {},
}

# Nothing is force-shown when empty any more (see DOB_WARNING note above).
KEEP_EMPTY = {
    'vineland': [],
    'dst': [],
    'ddst': [],
    'moca': [],
    'moca_assessment': [],
}

# Joined human-readable chronological-age rows.
# target: field path whose value/label is replaced.
#   'years_months_parts': parts = [yearsField, remainingMonthsField]
#   'days_to_years_months': parts = [daysField]
# label: translation key shown for the joined row (shared across scales;
# defined once in messages-en.properties).
CA_LABEL_KEY = 'report.scale.chronological_age'
JOIN_FIELDS = {
    'vineland': [
        {'target': 'g_child.ca_display', 'format': 'years_months_parts',
         'parts': ['g_child.ca_years_display', 'g_child.ca_rem_months'],
         'label': CA_LABEL_KEY},
    ],
    'dst': [
        {'target': 'g_child.ca_days', 'format': 'days_to_years_months',
         'parts': ['g_child.ca_days'], 'label': CA_LABEL_KEY},
    ],
    'ddst': [
        {'target': 'g_meta.ca_days', 'format': 'days_to_years_months',
         'parts': ['g_meta.ca_days'], 'label': CA_LABEL_KEY},
    ],
    'moca': [],
    'moca_assessment': [],
}

# Date fields to format as readable dates in the report.
DATE_FIELDS = {
    'vineland': ['g_child.assessment_date', 'g_child.child_dob'],
    'dst': ['g_child.assessment_date', 'g_child.child_dob'],
    'ddst': ['g_meta.test_date', 'g_meta.child_dob'],
    'moca': ['g_admin.assessment_datetime'],
    'moca_assessment': ['g_admin.assessment_datetime'],
}

# Calculated result fields shown in the results summary, in order.
RESULT_FIELDS = {
    'cog5': [
        'score_pmis_total', 'score_pbst', 'score_dementia_risk',
        'PMIS_FINAL_RESULT', 'PBST_FINAL_RESULT', 'MCR_FINAL_RESULT',
        'FINAL_TRIAGE', 'calc_overall_conclusion',
    ],
    'vineland': [
        'raw_score', 'total_yes_items', 'total_no_items',
        'total_not_assessed_items', 'social_age_months', 'social_age_years',
        'social_age_remaining_months', 'social_quotient_raw',
        'social_quotient_rounded', 'sq_interpretation',
    ],
    'dst': [
        'total_passed_items', 'total_failed_items',
        'total_not_assessed_items', 'developmental_age_months',
        'developmental_age_years', 'developmental_age_remaining_months',
        'developmental_quotient_raw', 'developmental_quotient_rounded',
    ],
    'ddst': [
        'd_ps_status', 'd_fm_status', 'd_lang_status', 'd_gm_status',
        'overall',
    ],
    'moca_assessment': [
        'moca_visuospatial_total', 'moca_naming_total', 'moca_memory_total',
        'moca_attention_total', 'moca_language_total',
        'moca_abstraction_total', 'moca_orientation_total',
        'moca_raw_total', 'moca_education_adj', 'moca_total_score',
        'moca_interpretation',
    ],
    'moca': [
        'selected_instrument', 'education_band', 'education_confounded',
        'moca_total_score', 'moca_mis_score', 'moca_interpretation',
        'te_visuospatial_total', 'te_naming_total', 'te_memory_total',
        'te_attention_total', 'te_language_total', 'te_abstraction_total',
        'te_orientation_total', 'te_education_adj',
        'mb_executive_total', 'mb_fluency_total', 'mb_orientation_total',
        'mb_calculation_total', 'mb_abstraction_total', 'mb_memory_total',
        'mb_visuoperception_total', 'mb_naming_total', 'mb_attention_total',
        'mb_education_adj', 'mb_literacy_adj',
        'm8_visuospatial_total', 'm8_naming_total', 'm8_memory_total',
        'm8_attention_total', 'm8_language_total', 'm8_abstraction_total',
        'm8_orientation_total', 'm8_education_adj',
    ],
}

# Question -> item score field. Per-item scores are joined into the answer
# row. Derived from the form itself (every `item_<n>_score` leaf that has a
# matching `item_<n>` question), never a hard-coded count - if the form
# gains or loses items the mapping follows automatically.
SCORE_LEAF_RE = re.compile(r'^item_\d+_score$')


def _derive_score_fields(flat, parsed):
    if parsed is None:
        return {}
    leaves = set(_all_field_paths(parsed))
    field_paths = [i['path'] for i in flat if i['type'] == 'field']
    out = {}
    for leaf in sorted(leaves):
        if not SCORE_LEAF_RE.match(leaf):
            continue
        question = leaf[:-len('_score')]              # item_5_score -> item_5
        match = next((p for p in field_paths
                      if p == question or p.endswith('.' + question)), None)
        out[match or question] = leaf
    return out

# Free-text/notes fields (kept visible, suppressed when empty).
NOTES_FIELDS = {
    'cog5': ['Survey_Status_cog5', 'Reason_for_Dropout_cog5', 'screener_notes', 'clinical_notes'],
    'vineland': ['na_reason', 'clinical_notes', 'scoring_notes'],
    'dst': ['na_reason', 'clinical_notes'],
    'ddst': [],
    'moca': [],
    'moca_assessment': [],
}

# Curated human-readable labels for calculated, score, and system fields
KNOWN_FIELD_LABELS = {
    # CoG5 result & score fields
    'score_pmis_total': 'PMIS Total Score',
    'score_pbst': 'PBST Total Score',
    'score_dementia_risk': 'Dementia Risk Score',
    'PMIS_FINAL_RESULT': 'PMIS Final Result',
    'PBST_FINAL_RESULT': 'PBST Final Result',
    'MCR_FINAL_RESULT': 'MCR Final Result',
    'FINAL_TRIAGE': 'Overall Triage',
    'calc_overall_conclusion': 'Overall Conclusion',
    'Survey_Status_cog5': 'Survey Status',
    'Reason_for_Dropout_cog5': 'Reason for Dropout / Refusal',
    'g_meta.test_date': 'Assessment Date',
    'g_meta.child_dob': 'Date of Birth',
    'g_child.child_dob': 'Date of Birth',
    'g_child.assessment_date': 'Assessment Date',
    'g_admin.assessment_datetime': 'Assessment Date & Time',
    'upload_scale_image': 'Scale Upload (Image)',
    'upload_scale_pdf': 'Scale Upload (PDF)',
    'upload_caption': 'Upload Caption',

    # DDST results
    'd_ps_status': 'Personal-Social Status',
    'd_fm_status': 'Fine Motor Status',
    'd_lang_status': 'Language Status',
    'd_gm_status': 'Gross Motor Status',
    'overall': 'Overall Developmental Status',

    # Vineland / VSMS results
    'raw_score': 'Raw Score',
    'total_yes_items': 'Total Passed (Yes)',
    'total_no_items': 'Total Failed (No)',
    'total_not_assessed_items': 'Total Not Assessed',
    'social_age_months': 'Social Age (Months)',
    'social_age_years': 'Social Age (Years)',
    'social_age_remaining_months': 'Social Age (Remaining Months)',
    'social_quotient_raw': 'Social Quotient (Raw)',
    'social_quotient_rounded': 'Social Quotient (Rounded)',
    'sq_interpretation': 'Social Quotient Interpretation',

    # DST results
    'total_passed_items': 'Total Passed Items',
    'total_failed_items': 'Total Failed Items',
    'developmental_age_months': 'Developmental Age (Months)',
    'developmental_age_years': 'Developmental Age (Years)',
    'developmental_age_remaining_months': 'Developmental Age (Remaining Months)',
    'developmental_quotient_raw': 'Developmental Quotient (Raw)',
    'developmental_quotient_rounded': 'Developmental Quotient (Rounded)',

    # MoCA results
    'selected_instrument': 'Selected Instrument',
    'education_band': 'Education Band',
    'education_confounded': 'Education Confounded',
    'moca_total_score': 'MoCA Total Score',
    'moca_mis_score': 'Memory Index Score (MIS)',
    'moca_interpretation': 'MoCA Interpretation',
    'te_visuospatial_total': 'Telugu MoCA Visuospatial Total',
    'te_naming_total': 'Telugu MoCA Naming Total',
    'te_memory_total': 'Telugu MoCA Memory Total',
    'te_attention_total': 'Telugu MoCA Attention Total',
    'te_language_total': 'Telugu MoCA Language Total',
    'te_abstraction_total': 'Telugu MoCA Abstraction Total',
    'te_orientation_total': 'Telugu MoCA Orientation Total',
    'te_education_adj': 'Telugu MoCA Education Adjustment',
    'mb_executive_total': 'MoCA-B Executive Total',
    'mb_fluency_total': 'MoCA-B Fluency Total',
    'mb_orientation_total': 'MoCA-B Orientation Total',
    'mb_calculation_total': 'MoCA-B Calculation Total',
    'mb_abstraction_total': 'MoCA-B Abstraction Total',
    'mb_memory_total': 'MoCA-B Memory Total',
    'mb_visuoperception_total': 'MoCA-B Visuoperception Total',
    'mb_naming_total': 'MoCA-B Naming Total',
    'mb_attention_total': 'MoCA-B Attention Total',
    'mb_education_adj': 'MoCA-B Education Adjustment',
    'mb_literacy_adj': 'MoCA-B Literacy Adjustment',
    'm8_visuospatial_total': 'MoCA 8.1 Visuospatial Total',
    'm8_naming_total': 'MoCA 8.1 Naming Total',
    'm8_memory_total': 'MoCA 8.1 Memory Total',
    'm8_attention_total': 'MoCA 8.1 Attention Total',
    'm8_language_total': 'MoCA 8.1 Language Total',
    'm8_abstraction_total': 'MoCA 8.1 Abstraction Total',
    'm8_orientation_total': 'MoCA 8.1 Orientation Total',
    'm8_education_adj': 'MoCA 8.1 Education Adjustment',
    'moca_visuospatial_total': 'MoCA Visuospatial Total',
    'moca_naming_total': 'MoCA Naming Total',
    'moca_memory_total': 'MoCA Memory Total',
    'moca_attention_total': 'MoCA Attention Total',
    'moca_language_total': 'MoCA Language Total',
    'moca_abstraction_total': 'MoCA Abstraction Total',
    'moca_orientation_total': 'MoCA Orientation Total',
    'moca_raw_total': 'MoCA Raw Total',
    'moca_education_adj': 'MoCA Education Adjustment',

    # MOCA groups & fields
    'g_consent': 'Consent',
    'consent_reconfirmed': 'Consent Reconfirmed',
    'no_consent_note': 'Consent Note',
    'te_g_visuospatial': 'Visuospatial / Executive',
    'te_trail_making': 'Trail Making',
    'te_cube_copy': 'Cube Copy',
    'te_clock_contour': 'Clock Contour',
    'te_clock_numbers': 'Clock Numbers',
    'te_clock_hands': 'Clock Hands',
    'te_g_naming': 'Naming',
    'te_naming_1': 'Lion',
    'te_naming_2': 'Rhinoceros',
    'te_naming_3': 'Camel',
    'te_g_memory': 'Memory',
    'te_g_attention': 'Attention',
    'te_digit_span_fwd': 'Digit Span Forward',
    'te_digit_span_bwd': 'Digit Span Backward',
    'te_vigilance': 'Vigilance',
    'te_serial_7s': 'Serial 7s',
    'te_g_language': 'Language',
    'te_sentence_rep_1': 'Sentence Repetition 1',
    'te_sentence_rep_2': 'Sentence Repetition 2',
    'te_fluency_count': 'Fluency Count',
    'te_g_abstraction': 'Abstraction',
    'te_abstraction_1': 'Abstraction 1',
    'te_abstraction_2': 'Abstraction 2',
    'te_g_delayed_recall': 'Delayed Recall',
    'te_delayed_kannu': 'Eye (Kannu)',
    'te_delayed_kannu_cat': 'Eye (Kannu) Category Cue',
    'te_delayed_kannu_mc': 'Eye (Kannu) Multiple Choice',
    'te_delayed_cheera': 'Saree (Cheera)',
    'te_delayed_cheera_cat': 'Saree (Cheera) Category Cue',
    'te_delayed_cheera_mc': 'Saree (Cheera) Multiple Choice',
    'te_delayed_gudi': 'Temple (Gudi)',
    'te_delayed_gudi_cat': 'Temple (Gudi) Category Cue',
    'te_delayed_gudi_mc': 'Temple (Gudi) Multiple Choice',
    'te_delayed_gulabi': 'Rose (Gulabi)',
    'te_delayed_gulabi_cat': 'Rose (Gulabi) Category Cue',
    'te_delayed_gulabi_mc': 'Rose (Gulabi) Multiple Choice',
    'te_delayed_neelam': 'Blue (Neelam)',
    'te_delayed_neelam_cat': 'Blue (Neelam) Category Cue',
    'te_delayed_neelam_mc': 'Blue (Neelam) Multiple Choice',
    'te_g_orientation': 'Orientation',
    'te_orientation_date': 'Date',
    'te_orientation_month': 'Month',
    'te_orientation_year': 'Year',
    'te_orientation_day': 'Day',
    'te_orientation_place': 'Place',
    'te_orientation_city': 'City',
    'g_mandatory_history': 'Mandatory History',
    'mh_asked_to': 'Information Asked To',
    'mh_answered_by': 'Information Answered By',
    'mh_q1': '1. Loss of Memory',
    'mh_q2': '2. Difficulty in Everyday Activities',
    'mh_q3': '3. Getting Lost in Familiar Places',
    'mh_q4': '4. Difficulty Recognizing Familiar People',
    'mh_q5': '5. Personality / Behavior Changes',
    'mh_q6': '6. Depressed Mood / Loss of Interest',
    'mh_q7': '7. Physical Illness / Head Injury',
    'mh_q8': '8. Alcohol / Substance Use',
    'mh_q9': '9. Family History of Dementia',
    'mh_q10': '10. Sensory Impairment',
    'mb_g_timing': 'Timing',
    'mb_start_time': 'Start Time',
    'mb_g_executive': 'Executive Function',
    'mb_trail_making': 'Trail Making',
    'mb_g_fluency': 'Fluency',
    'mb_fluency_words': 'Fluency Words',
    'mb_fluency_count': 'Fluency Count',
    'mb_g_orientation': 'Orientation',
    'mb_orientation_time': 'Orientation Time',
    'mb_orientation_day': 'Orientation Day',
    'mb_orientation_month': 'Orientation Month',
    'mb_orientation_year': 'Orientation Year',
    'mb_orientation_place': 'Orientation Place',
    'mb_orientation_city': 'Orientation City',
    'mb_g_calculation': 'Calculation',
    'mb_calc_combinations': 'Calculation Combinations',
    'mb_calculation_score_select': 'Calculation Score',
    'mb_g_abstraction': 'Abstraction',
    'mb_abstraction_1': 'Abstraction 1',
    'mb_abstraction_2': 'Abstraction 2',
    'mb_abstraction_3': 'Abstraction 3',
    'mb_g_delayed_recall': 'Delayed Recall',
    'mb_g_visuoperception': 'Visuoperception',
    'mb_g_naming': 'Naming',
    'mb_attention_black': 'Attention - Black',
    'mb_attention_white': 'Attention - White',
    'assessment_image': 'Upload assessment sheet (image)',
    'assessment_pdf': 'Upload assessment sheet (PDF)',
    'g_assessment_sheets': 'Assessment Sheets',
    'assessment_sheets': 'Assessment Sheet',

    # CoG-5 results & triage
    'PMIS_FINAL_RESULT': 'PMIS Memory Result',
    'PBST_FINAL_RESULT': 'PBST Symbol Pattern Result',
    'MCR_FINAL_RESULT': 'MCR Motoric Cognitive Result',
    'FINAL_TRIAGE': 'Overall Clinical Triage',
    'calc_overall_conclusion': 'Overall Clinical Conclusion',
    'calc_pmis_meaning': 'Memory Screening Interpretation',
    'calc_pbst_meaning': 'Pattern Test Interpretation',
    'calc_mcr_meaning': 'Motoric Risk Interpretation',
    'calc_risk_breakdown': 'Identified Risk Factors',
    'Survey_Status_cog5': 'Survey Status',
    'Reason_for_Dropout_cog5': 'Reason for Dropout / Refusal',

    # Notes & general fields
    'na_reason': 'Reason for Not Assessing',
    'clinical_notes': 'Clinical Notes',
    'scoring_notes': 'Scoring Notes',
    'screener_notes': 'Screener Notes',
}


def clean_human_label(path, fallback=None):
    """Convert raw field keys or markdown question labels into clean, human-readable column headers."""
    if not path:
        return fallback or ''
    # Strip any markdown bold/italics and trailing colons
    cleaned = re.sub(r'\*\*|__', '', str(path)).strip().rstrip(':').strip()
    if cleaned in KNOWN_FIELD_LABELS:
        return KNOWN_FIELD_LABELS[cleaned]
    # Check last dotted path segment if it's a technical path
    last_seg = cleaned.split('.')[-1]
    if last_seg in KNOWN_FIELD_LABELS:
        return KNOWN_FIELD_LABELS[last_seg]
    # If cleaned is already a human phrase with spaces and words (e.g. "Q10. Timed Walk (2.5m)" or "1. Cries, laughs"), return it directly!
    if ' ' in cleaned and not cleaned.startswith('g_') and not cleaned.startswith('grp_'):
        return cleaned
    # If the string contains underscores or technical prefixes or looks like a variable name:
    if '.' in cleaned or '_' in cleaned or cleaned.startswith('g_') or cleaned.startswith('grp_') or cleaned.islower():
        s = cleaned
        for pfx in ('g_supporting_docs.', 'g_participant.', 'grp_cog5_scale.', 'g_meta.',
                    'g_child.', 'g_intro.', 'g_admin.', 'g_consent.', 'te_g_', 'mb_g_', 'm8_g_', 'grp_'):
            if s.startswith(pfx):
                s = s[len(pfx):]
        if '.' in s and ' ' not in s:
            s = s.split('.')[-1]
        for pfx in ('g_', 'calc_', 'grp_'):
            if s.startswith(pfx):
                s = s[len(pfx):]
        words = s.replace('_', ' ').split()
        return ' '.join(w.capitalize() for w in words)
    return cleaned

# Supporting-document fields (attachment leaves; suppressed when empty).
SUPPORTING_PREFIX = 'g_supporting_docs.'


# Human-readable value mappings for calculated fields (not select_one/select)
EXTRA_CHOICE_MAPS = {
    'moca': {
        'selected_instrument': {
            'moca_telugu': 'ICMR-NCTB Telugu MoCA (v7.1)',
            'moca_b': 'MoCA-Basic (MoCA-B)',
            'moca_8_1': 'MoCA Full, English, Version 8.1',
            'moca_8_2': 'MoCA Full, English, Version 8.2',
            'none': 'None',
        },
        'education_confounded': {
            'yes': 'Yes (Band A Telugu)',
            'no': 'No',
        },
    },
}


def build_report_display(form, flat, choice_maps, parsed=None):
    """The report_display entry for one scale (consumed by cht-core patch)."""
    hidden = _hidden_scale_fields(flat, form, parsed)
    # join targets must stay visible
    join_fields = JOIN_FIELDS.get(form, [])
    join_targets = {j['target'] for j in join_fields}
    hidden = [h for h in hidden if h not in join_targets]

    labels = {item['path']: item['label'] for item in flat}

    def _note_text(path):
        text = ' '.join((labels.get(path) or '').split())
        # strip the form's emphasis markers; Enketo renders them, the
        # report does not
        return text.replace('**', '')

    # score join map: question path (with group prefix) -> score field,
    # derived from the form (Vineland items live in band groups but their
    # *_score leaves are top-level - _derive_score_fields resolves both).
    score_fields = _derive_score_fields(flat, parsed)
    # per-item score leaves are internal - hide them from the raw field list
    for leaf in score_fields.values():
        if leaf not in hidden:
            hidden.append(leaf)
    # questions with score joins get an "Answer:" prefix in the report
    return {
        'title_key': f'report.{form}._title',
        'hide_fields': hidden,
        'suppress_empty': True,
        'keep_empty': KEEP_EMPTY.get(form, []),
        'choice_map': {**{p: m for p, m in choice_maps.items()}, **EXTRA_CHOICE_MAPS.get(form, {})},
        'score_fields': score_fields,
        'join_fields': join_fields,
        'note_values': {
            p: _note_text(p) for p in NOTE_VALUES.get(form, {})
        },
        'date_fields': DATE_FIELDS.get(form, []),
        'notes_fields': NOTES_FIELDS.get(form, []),
        'result_fields': RESULT_FIELDS.get(form, []),
    }


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

_cache = {}
# hold while reading metas when another thread may switch FORMS_DIR
# (the exports server: live-form sync, exports and forms-panel publishes)
LOCK = threading.RLock()


def set_forms_dir(path):
    """Read form XMLs from another folder (e.g. copies of the live forms)."""
    global FORMS_DIR
    if path != FORMS_DIR:
        FORMS_DIR = path
        _cache.clear()


def get_scale_meta(form):
    """Parsed + derived metadata for one scale (cached)."""
    meta = _cache.get(form)
    if meta is not None:
        return meta
    info = SCALES.get(form) or {'file': f'{form}.xml', 'title': form.replace('_', ' ').title()}
    path = os.path.join(FORMS_DIR, info['file'])
    if not os.path.isfile(path):
        title = info.get('title') or KNOWN_TITLES.get(form) or form.replace('_', ' ').title()
        meta = _cache[form] = {
            'form': form,
            'title': title,
            'xml_title': title,
            'flat': [],
            'labels': {},
            'choice_maps': {},
            'report_display': {'hide_fields': [], 'score_fields': {}, 'result_fields': [], 'notes_fields': []},
            'result_fields': [],
            'notes_fields': [],
        }
        return meta
    parsed = parse_form_xml(path)
    flat = flatten_fields(parsed)
    choice_maps = build_choice_maps(parsed, flat)
    labels = {}
    for item in flat:
        if item['type'] == 'group':
            labels[item['path']] = item['label']
        else:
            labels[item['path']] = item['label']
    meta = _cache[form] = {
        'form': form,
        'title': SCALES[form]['title'],
        'xml_title': parsed['title'],
        'flat': flat,
        'labels': labels,
        'choice_maps': choice_maps,
        'report_display': build_report_display(form, flat, choice_maps,
                                               parsed),
        'result_fields': RESULT_FIELDS.get(form, []),
        'notes_fields': NOTES_FIELDS.get(form, []),
    }
    return meta


def all_scale_metas():
    return {form: get_scale_meta(form) for form in SCALES}


def build_message_keys(form):
    """report.<form>.<path> -> label for every group and field that should
    be displayable. Hidden fields are included too (harmless, keeps the
    file complete and future-proof)."""
    meta = get_scale_meta(form)
    prefix = f'report.{form}.'
    out = {f'{prefix}_title': meta['title']}
    for path, label in meta['labels'].items():
        if label:
            out[prefix + path] = label
    # note-value fields get short display labels (value keeps the full text)
    for path, short in NOTE_LABELS.get(form, {}).items():
        if prefix + path in out or short:
            out[prefix + path] = short
    return out


def build_report_display_config():
    """Full report_display settings object for app_settings.json."""
    return {form: get_scale_meta(form)['report_display']
            for form in SCALES}


if __name__ == '__main__':
    meta = all_scale_metas()
    for form, m in meta.items():
        n_fields = sum(1 for i in m['flat'] if i['type'] == 'field')
        n_groups = sum(1 for i in m['flat'] if i['type'] == 'group')
        print(f"{form}: {n_fields} fields, {n_groups} groups, "
              f"{len(m['choice_maps'])} choice-mapped fields, "
              f"{len(m['report_display']['hide_fields'])} hidden, "
              f"title={m['title']!r}")
    print(json.dumps(meta['vineland']['choice_maps']['g_y0_1.item_1'],
                     indent=2))
