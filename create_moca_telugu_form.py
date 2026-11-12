"""create_moca_telugu_form.py - Generate the MoCA Version 7.1 Telugu XLSForm and properties.

Source of truth (exact text & scoring only):
  scales/moca_telugu/moca-7.1-telugu-icmr-nctb.pdf (pages 24-30)
"""

import json
import os
import openpyxl

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
FORMS_APP_DIR = os.path.join(REPO_ROOT, 'forms', 'app')


def build_moca_telugu_xlsx():
    wb = openpyxl.Workbook()

    # ============================================================
    # SURVEY SHEET
    # ============================================================
    ws = wb.active
    ws.title = 'survey'

    headers = [
        'type', 'name', 'label::en', 'label::te', 'hint::en', 'appearance',
        'relevant', 'constraint', 'constraint_message::en', 'required',
        'calculation', 'default', 'read_only', 'media::image::en',
    ]
    ws.append(headers)

    consent_yes = "${consent_reconfirmed} = 'yes'"

    rows = [
        # Contact & user metadata inputs (CHT standard)
        ['begin group', 'inputs', 'Participant', 'పాల్గొనేవారు', '', 'field-list', "./source = 'user'", '', '', '', '', '', '', ''],
        ['hidden', 'source', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'source_id', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'user', 'User', 'వినియోగదారు', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['hidden', 'contact_id', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'name', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'contact', 'Participant', 'పాల్గొనేవారు', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['db:person', '_id', 'Select participant', 'పాల్గొనేవారిని ఎంచుకోండి', '', 'select-contact type-participant', '', '', '', '', '', '', '', ''],
        ['hidden', 'patient_id', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'name', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'age', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'gender', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'phone', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'education_level', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'consent_given', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'informant_name', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'informant_phone', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['calculate', 'patient_uuid', '', '', '', '', '', '', '', '', '../inputs/contact/_id', '', '', ''],
        ['calculate', 'patient_id', '', '', '', '', '', '', '', '', 'coalesce(../inputs/contact/patient_id, ../inputs/contact/_id)', '', '', ''],
        ['calculate', 'patient_name', '', '', '', '', '', '', '', '', '../inputs/contact/name', '', '', ''],
        ['calculate', 'patient_sex', '', '', '', '', '', '', '', '', '../inputs/contact/gender', '', '', ''],
        ['calculate', 'patient_age', '', '', '', '', '', '', '', '', '../inputs/contact/age', '', '', ''],
        ['calculate', 'patient_phone', '', '', '', '', '', '', '', '', '../inputs/contact/phone', '', '', ''],
        ['calculate', 'patient_education', '', '', '', '', '', '', '', '', '../inputs/contact/education_level', '', '', ''],
        ['calculate', 'screened_by', '', '', '', '', '', '', '', '', '../inputs/user/contact_id', '', '', ''],
        ['calculate', 'screened_by_name', '', '', '', '', '', '', '', '', '../inputs/user/name', '', '', ''],

        # Education & consent
        ['begin group', 'g_consent', 'Education', 'విద్య', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'consent_reconfirmed', 'Consent re-confirmed?', 'సమ్మతి నిర్ధారించబడిందా?', '', '', '', '', '', 'yes', '', '', '', ''],
        ['integer', 'education_years', 'Education (years)', 'విద్య (సంవత్సరాలు)', 'Add one point for an individual who has 12 years or fewer of formal education, for a possible maximum of 30 points.', '', consent_yes, '. >= 0 and . <= 30', 'Enter years between 0 and 30.', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 1. VISUOSPATIAL / EXECUTIVE (5 pts)
        ['begin group', 'g_visuospatial', 'VISUOSPATIAL / EXECUTIVE', 'విజువోస్పేషియల్ / ఎగ్జిక్యూటివ్', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'trail_making_inst', '1 − క - 2 - ఖ - 3 - గ - 4 - ఘ - 5 - ఙ (మొదలు పెట్టండి -> ఆపండి)', '1 − క - 2 - ఖ - 3 - గ - 4 - ఘ - 5 - ఙ (మొదలు పెట్టండి -> ఆపండి)', '', '', '', '', '', '', '', '', '', 'trail_making_telugu.png'],
        ['select_one yesno', 'trail_making', 'Alternating Trail Making (1 - క - 2 - ఖ - 3 - గ - 4 - ఘ - 5 - ఙ)', 'ట్రయల్ మేకింగ్ (1 - క - 2 - ఖ - 3 - గ - 4 - ఘ - 5 - ఙ)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'cube_inst', 'చిత్రాన్ని గీయండి.', 'చిత్రాన్ని గీయండి.', '', '', '', '', '', '', '', '', '', 'cube.png'],
        ['select_one yesno', 'cube_copy', 'Cube Copy (చిత్రాన్ని గీయండి.)', 'ఘనం (చిత్రాన్ని గీయండి.)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'clock_inst', 'గడియారం గీయండి. (పదకొండు గంటల పది నిమిషాలు) (3 Points)', 'గడియారం గీయండి. (పదకొండు గంటల పది నిమిషాలు) (3 Points)', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'clock_contour', 'Contour', 'గడియార ఆకారం (Contour)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'clock_numbers', 'Numbers', 'సంఖ్యలు (Numbers)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'clock_hands', 'Hands', 'ముళ్ళు (Hands)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 2. NAMING (3 pts)
        ['begin group', 'g_naming', 'NAMING', 'పేర్లు చెప్పడం (NAMING)', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'naming_inst', 'ఈ జంతువు పేరు చెప్పండి.', 'ఈ జంతువు పేరు చెప్పండి.', '', '', '', '', '', '', '', '', '', 'naming.png'],
        ['select_one yesno', 'naming_1', '1. Lion', '1. సింహం (Lion)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'naming_2', '2. Rhinoceros / Rhino', '2. ఖడ్గమృగం (Rhinoceros)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'naming_3', '3. Camel / Dromedary', '3. ఒంటె (Camel)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 3. MEMORY (0 pts - 2 trials)
        ['begin group', 'g_memory', 'MEMORY (No points)', 'జ్ఞాపకశక్తి (No points)', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'memory_inst_trial1', 'పదాల జాబితా చదవండి, పేషెంట్/సబ్జెక్ట్ ఆ పదాలను తిరిగి చెప్పాలి. మొదటి ప్రయత్నం సరి అయినా కూడా రెండు ప్రయత్నాలు చేయండి. 5 నిమిషాల తరువాత మళ్ళీ అడగండి.', 'పదాల జాబితా చదవండి, పేషెంట్/సబ్జెక్ట్ ఆ పదాలను తిరిగి చెప్పాలి. మొదటి ప్రయత్నం సరి అయినా కూడా రెండు ప్రయత్నాలు చేయండి. 5 నిమిషాల తరువాత మళ్ళీ అడగండి.', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'memory_words', 'కన్ను | చీర | గుడి | గులాబీ | నీలం', 'కన్ను | చీర | గుడి | గులాబీ | నీలం', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'mem_trial1', '1వ ప్రయత్నం (1st trial)', '1వ ప్రయత్నం', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mem_t1_kannu', 'కన్ను (1వ ప్రయత్నం)', 'కన్ను (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_cheera', 'చీర (1వ ప్రయత్నం)', 'చీర (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_gudi', 'గుడి (1వ ప్రయత్నం)', 'గుడి (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_gulabi', 'గులాబీ (1వ ప్రయత్నం)', 'గులాబీ (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_neelam', 'నీలం (1వ ప్రయత్నం)', 'నీలం (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'memory_inst_trial2', '2వ ప్రయత్నం', '2వ ప్రయత్నం', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'mem_trial2', '2వ ప్రయత్నం (2nd trial)', '2వ ప్రయత్నం', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mem_t2_kannu', 'కన్ను (2వ ప్రయత్నం)', 'కన్ను (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_cheera', 'చీర (2వ ప్రయత్నం)', 'చీర (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_gudi', 'గుడి (2వ ప్రయత్నం)', 'గుడి (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_gulabi', 'గులాబీ (2వ ప్రయత్నం)', 'గులాబీ (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_neelam', 'నీలం (2వ ప్రయత్నం)', 'నీలం (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 4. ATTENTION (6 pts)
        ['begin group', 'g_attention', 'ATTENTION', 'అవధానం (ATTENTION)', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'digits_inst', 'ఈ అంకెలు చదవండి. [ 1 Digit / sec. ]', 'ఈ అంకెలు చదవండి. [ 1 Digit / sec. ]', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'digit_span_fwd', 'పేషెంట్/సబ్జెక్ట్ అంకెలను వరుస క్రమంలో చెప్పాలి: 2 1 8 5 4', 'పేషెంట్/సబ్జెక్ట్ అంకెలను వరుస క్రమంలో చెప్పాలి: 2 1 8 5 4', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'digit_span_bwd', 'పేషెంట్/సబ్జెక్ట్ అంకెలను వెనుక క్రమంలో చెప్పాలి: 7 4 2', 'పేషెంట్/సబ్జెక్ట్ అంకెలను వెనుక క్రమంలో చెప్పాలి: 7 4 2', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'vigilance_inst', "అక్షరాల జాబితా చదవండి. 'అ' అనే అక్షరం విన్నప్పుడల్లా పేషెంట్/వ్యక్తి చేయి తో ఒకసారి తట్టాలి. No points if >= 2 errors: ప బ అ క మ న అ అ జ క ల బ అ ఫ అ క డ ఈ అ అ అ జ అ మా ఓ ఫ అ అ బ", "అక్షరాల జాబితా చదవండి. 'అ' అనే అక్షరం విన్నప్పుడల్లా పేషెంట్/వ్యక్తి చేయి తో ఒకసారి తట్టాలి. No points if >= 2 errors: ప బ అ క మ న అ అ జ క ల బ అ ఫ అ క డ ఈ అ అ అ జ అ మా ఓ ఫ అ అ బ", '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'vigilance', "Vigilance ('అ' — No points if >= 2 errors)", "అక్షరాల జాబితా ('అ' — No points if >= 2 errors)", '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'serial7_inst', '100 నుండి క్రమంగా 7 ను తీసివేయండి. [ ] 93  [ ] 86  [ ] 79  [ ] 72  [ ] 65', '100 నుండి క్రమంగా 7 ను తీసివేయండి. [ ] 93  [ ] 86  [ ] 79  [ ] 72  [ ] 65', '', '', '', '', '', '', '', '', '', ''],
        ['select_one serial7_score', 'serial_7s', 'Serial 7s (100 నుండి క్రమంగా 7 ను తీసివేయండి)', '100 నుండి 7 తీసివేయడం', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 5. LANGUAGE (3 pts)
        ['begin group', 'g_language', 'LANGUAGE', 'భాష (LANGUAGE)', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['select_one yesno', 'sentence_rep_1', 'తిరిగి చెప్పండి: రాజు ఒక్కడే ఈ రోజు సహాయం చేస్తాడని నేను అనుకుంటున్నాను.', 'తిరిగి చెప్పండి: రాజు ఒక్కడే ఈ రోజు సహాయం చేస్తాడని నేను అనుకుంటున్నాను.', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'sentence_rep_2', 'కుక్కను గదిలో చూసినప్పుడల్లా పిల్లి మంచం కింద దాక్కుంటుంది.', 'కుక్కను గదిలో చూసినప్పుడల్లా పిల్లి మంచం కింద దాక్కుంటుంది.', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['integer', 'fluency_count', "ఒక నిమిషంలో 'ప' అక్షరంతో మొదలయ్యే పదాలు ఎన్ని ఎక్కువ చెప్పగలిగితే అన్ని చెప్పండి (N >= 11 words)", "ఒక నిమిషంలో 'ప' అక్షరంతో మొదలయ్యే పదాలు (N >= 11 words)", '', '', '', '. >= 0 and . <= 100', 'Enter a count between 0 and 100.', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 6. ABSTRACTION (2 pts)
        ['begin group', 'g_abstraction', 'ABSTRACTION', 'పోలికలు (ABSTRACTION)', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'abstraction_inst_practice', 'పోలికలు : టమాట - బంగాళాదుంప = కూరగాయలు', 'పోలికలు : టమాట - బంగాళాదుంప = కూరగాయలు', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'abstraction_1', 'రైలు - సైకిల్', 'రైలు - సైకిల్', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'abstraction_2', 'గడియారం - స్కేలు', 'గడియారం - స్కేలు', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 7. DELAYED RECALL (5 pts - uncued only)
        ['begin group', 'g_delayed_recall', 'DELAYED RECALL', 'ఆలస్య జ్ఞాపకశక్తి (DELAYED RECALL)', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'delayed_inst', 'ఇంతకుముందు చెప్పిన పదాలను చెప్పండి. (ఎటువంటి సూచన లేకుండా) — Points for UNCUED recall only', 'ఇంతకుముందు చెప్పిన పదాలను చెప్పండి. (ఎటువంటి సూచన లేకుండా) — Points for UNCUED recall only', '', '', '', '', '', '', '', '', '', ''],

        ['select_one yesno', 'delayed_kannu', 'కన్ను (ఎటువంటి సూచన లేకుండా)', 'కన్ను (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_kannu_cat', 'కన్ను — Category cue: ముఖంలో ఒక భాగం.', 'కన్ను — Category cue: ముఖంలో ఒక భాగం.', '', '', "${delayed_kannu} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_kannu_mc', 'కన్ను — Multiple choice: కన్ను, ముక్కు, చెవి', 'కన్ను — Multiple choice: కన్ను, ముక్కు, చెవి', '', '', "${delayed_kannu} = 'no' and ${delayed_kannu_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_cheera', 'చీర (ఎటువంటి సూచన లేకుండా)', 'చీర (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_cheera_cat', 'చీర — Category cue: దుస్తులు/వస్త్రాలలో ఒక రకం.', 'చీర — Category cue: దుస్తులు/వస్త్రాలలో ఒక రకం.', '', '', "${delayed_cheera} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_cheera_mc', 'చీర — Multiple choice: వోణి, పంచ, చీర', 'చీర — Multiple choice: వోణి, పంచ, చీర', '', '', "${delayed_cheera} = 'no' and ${delayed_cheera_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_gudi', 'గుడి (ఎటువంటి సూచన లేకుండా)', 'గుడి (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_gudi_cat', 'గుడి — Category cue: భవనాలలో ఒక రకం.', 'గుడి — Category cue: భవనాలలో ఒక రకం.', '', '', "${delayed_gudi} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_gudi_mc', 'గుడి — Multiple choice: గుడి, బడి, ఆసుపత్రి', 'గుడి — Multiple choice: గుడి, బడి, ఆసుపత్రి', '', '', "${delayed_gudi} = 'no' and ${delayed_gudi_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_gulabi', 'గులాబీ (ఎటువంటి సూచన లేకుండా)', 'గులాబీ (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_gulabi_cat', 'గులాబీ — Category cue: పూలలో ఒక రకం.', 'గులాబీ — Category cue: పూలలో ఒక రకం.', '', '', "${delayed_gulabi} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_gulabi_mc', 'గులాబీ — Multiple choice: మల్లె, గులాబీ, చామంతి', 'గులాబీ — Multiple choice: మల్లె, గులాబీ, చామంతి', '', '', "${delayed_gulabi} = 'no' and ${delayed_gulabi_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_neelam', 'నీలం (ఎటువంటి సూచన లేకుండా)', 'నీలం (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_neelam_cat', 'నీలం — Category cue: ఒక రంగు.', 'నీలం — Category cue: ఒక రంగు.', '', '', "${delayed_neelam} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_neelam_mc', 'నీలం — Multiple choice: తెలుపు, ఎరుపు, నీలం', 'నీలం — Multiple choice: తెలుపు, ఎరుపు, నీలం', '', '', "${delayed_neelam} = 'no' and ${delayed_neelam_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 8. ORIENTATION (6 pts)
        ['begin group', 'g_orientation', 'ORIENTATION', 'దిక్సూచి (ORIENTATION)', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['select_one yesno', 'orientation_date', 'తేది (Date)', 'తేది', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_month', 'నెల (Month)', 'నెల', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_year', 'సంవత్సరం (Year)', 'సంవత్సరం', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_day', 'రోజు (Day)', 'రోజు', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_place', 'స్థలం (Place)', 'స్థలం', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_city', 'నగరం (City)', 'నగరం', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # Scoring calculations (exact PDF scoring rules)
        ['calculate', 'moca_te_visuospatial_total', '', '', '', '', '', '', '', '', 'if(${trail_making} = "yes", 1, 0) + if(${cube_copy} = "yes", 1, 0) + if(${clock_contour} = "yes", 1, 0) + if(${clock_numbers} = "yes", 1, 0) + if(${clock_hands} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'moca_te_naming_total', '', '', '', '', '', '', '', '', 'if(${naming_1} = "yes", 1, 0) + if(${naming_2} = "yes", 1, 0) + if(${naming_3} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'moca_te_attention_total', '', '', '', '', '', '', '', '', 'if(${digit_span_fwd} = "yes", 1, 0) + if(${digit_span_bwd} = "yes", 1, 0) + if(${vigilance} = "yes", 1, 0) + if(${serial_7s} != "", number(${serial_7s}), 0)', '', '', ''],
        ['calculate', 'moca_te_language_total', '', '', '', '', '', '', '', '', 'if(${sentence_rep_1} = "yes", 1, 0) + if(${sentence_rep_2} = "yes", 1, 0) + if(${fluency_count} != "" and ${fluency_count} >= 11, 1, 0)', '', '', ''],
        ['calculate', 'moca_te_abstraction_total', '', '', '', '', '', '', '', '', 'if(${abstraction_1} = "yes", 1, 0) + if(${abstraction_2} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'moca_te_memory_total', '', '', '', '', '', '', '', '', 'if(${delayed_kannu} = "yes", 1, 0) + if(${delayed_cheera} = "yes", 1, 0) + if(${delayed_gudi} = "yes", 1, 0) + if(${delayed_gulabi} = "yes", 1, 0) + if(${delayed_neelam} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'moca_te_orientation_total', '', '', '', '', '', '', '', '', 'if(${orientation_date} = "yes", 1, 0) + if(${orientation_month} = "yes", 1, 0) + if(${orientation_year} = "yes", 1, 0) + if(${orientation_day} = "yes", 1, 0) + if(${orientation_place} = "yes", 1, 0) + if(${orientation_city} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'moca_te_raw_total', '', '', '', '', '', '', '', '', '${moca_te_visuospatial_total} + ${moca_te_naming_total} + ${moca_te_attention_total} + ${moca_te_language_total} + ${moca_te_abstraction_total} + ${moca_te_memory_total} + ${moca_te_orientation_total}', '', '', ''],
        ['calculate', 'moca_te_education_adj', '', '', '', '', '', '', '', '', 'if(${education_years} != "" and ${education_years} <= 12, 1, 0)', '', '', ''],
        ['calculate', 'moca_te_total_score', '', '', '', '', '', '', '', '', 'if(${consent_reconfirmed} = "yes", min(${moca_te_raw_total} + ${moca_te_education_adj}, 30), "")', '', '', ''],
        ['calculate', 'moca_te_interpretation', '', '', '', '', '', '', '', '', 'if(${consent_reconfirmed} != "yes", "not assessed (no consent)", if(${moca_te_total_score} >= 26, "normal", "impaired"))', '', '', ''],

        # Admin / Date
        ['begin group', 'g_admin', 'TOTAL', 'మొత్తం స్కోరు', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['string', 'administered_by', 'Administered by', 'పరీక్ష నిర్వహించినవారు', '', '', '', '', '', 'yes', '', '', '', ''],
        ['dateTime', 'assessment_datetime', 'DATE', 'తేదీ', '', '', '', '', '', 'yes', "if(. = '', now(), .)", '', '', ''],
        ['note', 'score_summary_note', 'TOTAL: ${moca_te_total_score} / 30 (A final total score of 26 and above is considered normal)', 'మొత్తం: ${moca_te_total_score} / 30 (26 మరియు అంతకంటే ఎక్కువ స్కోరు సాధారణంగా పరిగణించబడుతుంది)', '', '', consent_yes, '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ]

    for r in rows:
        ws.append(r)

    # ============================================================
    # CHOICES SHEET
    # ============================================================
    ws_choices = wb.create_sheet('choices')
    ws_choices.append(['list_name', 'name', 'label::en', 'label::te'])

    choices = [
        ['yesno', 'yes', 'Yes', 'అవును'],
        ['yesno', 'no', 'No', 'కాదు'],
        ['serial7_score', '0', '0 correct (0 pt)', '0 సరిగ్గా (0 pt)'],
        ['serial7_score', '1', '1 correct (1 pt)', '1 సరిగ్గా (1 pt)'],
        ['serial7_score', '2', '2 or 3 correct (2 pts)', '2 లేదా 3 సరిగ్గా (2 pts)'],
        ['serial7_score', '3', '4 or 5 correct subtractions (3 pts)', '4 లేదా 5 సరిగ్గా (3 pts)'],
        ['delayed_cue', 'yes', 'Recalled with cue', 'సూచనతో గుర్తుకొచ్చింది'],
        ['delayed_cue', 'no', 'Not recalled even with cue', 'సూచనతో కూడా గుర్తురాలేదు'],
    ]

    for c in choices:
        ws_choices.append(c)

    # ============================================================
    # SETTINGS SHEET
    # ============================================================
    ws_settings = wb.create_sheet('settings')
    ws_settings.append(['form_title', 'form_id', 'version', 'default_language'])
    ws_settings.append(['MoCA Version 7.1 Telugu', 'moca_telugu', '2026-11-12 01-00', 'en'])

    os.makedirs(FORMS_APP_DIR, exist_ok=True)
    xlsx_path = os.path.join(FORMS_APP_DIR, 'moca_telugu.xlsx')
    wb.save(xlsx_path)
    print(f'Saved: {xlsx_path}')

    props_path = os.path.join(FORMS_APP_DIR, 'moca_telugu.properties.json')
    props = {
        'title': [
            {
                'locale': 'en',
                'content': 'Montreal Cognitive Assessment (MoCA) - Version 7.1 Telugu',
            },
            {
                'locale': 'te',
                'content': 'మాంట్రియల్ కాగ్నిటివ్ అసెస్‌మెంట్ (MoCA) - తెలుగు వెర్షన్ 7.1',
            },
        ],
        'icon': 'icon-healthcare-assessment',
        'context': {
            'person': True,
            'place': False,
            'expression': '!contact || summary.is_participant',
        },
    }
    with open(props_path, 'w', encoding='utf-8', newline='\n') as f:
        json.dump(props, f, ensure_ascii=False, indent=2)
        f.write('\n')
    print(f'Saved: {props_path}')


if __name__ == '__main__':
    build_moca_telugu_xlsx()

