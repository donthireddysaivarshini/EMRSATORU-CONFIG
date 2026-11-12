"""create_moca_form.py - Generate the unified MOCA XLSForm and properties.

Integrates:
  1. Consent
  2. Education, Language and Informant (Annex A Part 2)
  3. Sensory Status and Testing Conditions (Annex A Part 3)
  4. Conditional MoCA Instrument Selection (Table 8.2):
     - Telugu (Band A/B/C) -> ICMR-NCTB Telugu MoCA v7.1 (Band A flagged education-confounded)
     - English + Band A    -> MoCA-Basic (MoCA-B)
     - English + Band B/C  -> MoCA Full, English, Version 8.1
  5. Mandatory History Questions (Annex A Part 5)
"""

import json
import os
import openpyxl

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
FORMS_APP_DIR = os.path.join(REPO_ROOT, 'forms', 'app')


def build_moca_xlsx():
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = 'survey'

    headers = [
        'type', 'name', 'label::en', 'label::te', 'hint::en', 'appearance',
        'relevant', 'constraint', 'constraint_message::en', 'required',
        'calculation', 'default', 'read_only', 'media::image::en',
    ]
    ws.append(headers)

    consent_yes = "${consent_reconfirmed} = 'yes'"
    is_te = "${selected_instrument} = 'moca_telugu'"
    is_mb = "${selected_instrument} = 'moca_b'"
    is_m8 = "${selected_instrument} = 'moca_8_1'"

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

        # ------------------------------------------------------------
        # 1. CONSENT
        # ------------------------------------------------------------
        ['begin group', 'g_consent', '1. CONSENT', '1. సమ్మతి', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'consent_reconfirmed', 'Consent re-confirmed?', 'సమ్మతి నిర్ధారించబడిందా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'no_consent_note', 'Consent not given. Assessment cannot proceed.', 'సమ్మతి ఇవ్వబడలేదు. పరీక్ష కొనసాగదు.', '', '', "${consent_reconfirmed} = 'no'", '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ------------------------------------------------------------
        # 2. EDUCATION, LANGUAGE AND INFORMANT
        # ------------------------------------------------------------
        ['begin group', 'g_edu_lang_informant', '2. EDUCATION, LANGUAGE AND INFORMANT', '2. విద్య, భాష మరియు సమాచారకర్త', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['integer', 'years_schooling', 'Years of formal schooling', 'పాఠశాల విద్య సంవత్సరాలు', '', '', '', '. >= 0 and . <= 30', 'Enter years between 0 and 30.', 'yes', '', '', '', ''],
        ['select_one yesno', 'can_read_write', 'Can read and write a simple sentence?', 'సాధారణ వాక్యాన్ని చదవగలరా మరియు వ్రాయగలరా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['calculate', 'education_band', '', '', '', '', '', '', '', '', "if(${years_schooling} != '' and (number(${years_schooling}) <= 4 or ${can_read_write} = 'no'), 'A', if(${years_schooling} != '' and number(${years_schooling}) <= 10, 'B', if(${years_schooling} != '', 'C', '')))", '', '', ''],
        ['note', 'education_band_note', 'Education band: ${education_band} (A: <=4 yrs or illiterate | B: 5-10 yrs | C: >10 yrs)', 'విద్యా స్థాయి: ${education_band}', '', '', "${education_band} != ''", '', '', '', '', '', '', ''],
        ['text', 'languages_spoken', 'Languages spoken', 'మాట్లాడే భాషలు', '', '', '', '', '', 'no', '', '', '', ''],
        ['select_one luq_lang', 'luq_proficiency', 'LUQ >=70% proficiency in', 'LUQ >=70% ప్రావీణ్యం ఉన్న భాష', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'is_strongest_language', "Is this the patient's strongest language?", 'ఇది రోగికి అత్యంత బలమైన భాషా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one test_lang', 'test_language', 'Test language used', 'పరీక్షకు ఉపయోగించిన భాష', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'informant_present', 'Informant/attendant present today?', 'ఈ రోజు సమాచారకర్త/సహాయకుడు ఉన్నారా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'band_a_rebook_note', 'Band A with no informant present — protocol recommends rebooking with an informant/attendant (EASI is mandatory for Band A).', 'బ్యాండ్ A — సమాచారకర్త లేరు; దయచేసి సమాచారకర్తతో మళ్ళీ రీబుక్ చేయండి.', '', '', "${education_band} = 'A' and ${informant_present} = 'no'", '', '', '', '', '', '', ''],
        ['text', 'informant_name_entry', 'Informant name', 'సమాచారకర్త పేరు', '', '', "${informant_present} = 'yes'", '', '', 'no', '', '', '', ''],
        ['select_one contact_freq', 'informant_relationship_contact', 'Relationship / contact with patient', 'రోగితో సంబంధం / పరిచయం', '', 'horizontal-compact', "${informant_present} = 'yes'", '', '', 'no', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ------------------------------------------------------------
        # 3. SENSORY STATUS AND TESTING CONDITIONS
        # ------------------------------------------------------------
        ['begin group', 'g_sensory_conditions', '3. SENSORY STATUS AND TESTING CONDITIONS', '3. ఇంద్రియ స్థితి మరియు పరీక్ష పరిస్థితులు', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['select_one yesno', 'glasses_needed_worn', 'Glasses needed — and worn today?', 'కళ్ళద్దాలు అవసరమా — మరియు ఈ రోజు ధరించారా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'hearing_aid_working', 'Hearing aid needed — and worn and working?', 'వినికిడి యంత్రం అవసరమా — మరియు ధరించి పనిచేస్తుందా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'whisper_test_passed', "Whisper test at arm's length passed?", 'చేయి దూరంలో గుసగుస పరీక్షలో ఉత్తీర్ణత సాధించారా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'near_vision_passed', 'Near-vision card passed?', 'సమీప దృష్టి కార్డ్ పరీక్షలో ఉత్తీర్ణత సాధించారా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'room_conditions_ok', 'Room quiet, door closed, phones away, informant outside?', 'గది నిశ్శబ్దంగా, తలుపు మూసి, ఫోన్లు దూరంగా, సమాచారకర్త బయట ఉన్నారా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['integer', 'mood_screen_score', 'Mood screen score (record only, do not interpret)', 'మూడ్ స్క్రీన్ స్కోరు (నమోదు మాత్రమే)', '', '', '', '. >= 0 and . <= 100', 'Enter a non-negative integer.', 'no', '', '', '', ''],
        ['select_multiple other_factor', 'other_factors_today', 'Other factors today', 'ఈ రోజు ఇతర అంశాలు', '', '', '', '', '', 'no', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ------------------------------------------------------------
        # INSTRUMENT SELECTION (Table 8.2)
        # ------------------------------------------------------------
        ['calculate', 'selected_instrument', '', '', '', '', '', '', '', '', "if(${consent_reconfirmed} != 'yes', '', if(${test_language} = 'telugu', 'moca_telugu', if(${test_language} = 'english' and ${education_band} = 'A', 'moca_b', if(${test_language} = 'english' and (${education_band} = 'B' or ${education_band} = 'C'), 'moca_8_1', 'none'))))", '', '', ''],
        ['calculate', 'education_confounded', '', '', '', '', '', '', '', '', "if(${selected_instrument} = 'moca_telugu' and ${education_band} = 'A', 'yes', 'no')", '', '', ''],

        ['begin group', 'g_instrument_banner', '4. INSTRUMENT SELECTION', '4. పరీక్ష ఎంపిక', '', 'field-list', "${selected_instrument} != ''", '', '', '', '', '', '', ''],
        ['note', 'banner_telugu_a', 'Primary instrument: ICMR-NCTB Telugu MoCA (administer and record; flag as education-confounded). Second instrument: RUDAS-Telugu. Informant: EASI-Telugu Mandatory.', '', '', '', "${selected_instrument} = 'moca_telugu' and ${education_band} = 'A'", '', '', '', '', '', '', ''],
        ['note', 'banner_telugu_b', 'Primary instrument: ICMR-NCTB Telugu MoCA. Informant: EASI-Telugu strongly recommended (mandatory if referral concerns daily functioning).', '', '', '', "${selected_instrument} = 'moca_telugu' and ${education_band} = 'B'", '', '', '', '', '', '', ''],
        ['note', 'banner_telugu_c', 'Primary instrument: ICMR-NCTB Telugu MoCA. Informant: EASI-Telugu if an informant is present.', '', '', '', "${selected_instrument} = 'moca_telugu' and ${education_band} = 'C'", '', '', '', '', '', '', ''],
        ['note', 'banner_moca_b', 'Primary instrument: MoCA-Basic (English Band A: illiterate or <=4 yrs). Informant: EASI Mandatory.', '', '', '', is_mb, '', '', '', '', '', '', ''],
        ['note', 'banner_moca_8_1', 'Primary instrument: MoCA Full, English, Version 8.1 (English Band B or C). Informant: EASI if an informant is present.', '', '', '', is_m8, '', '', '', '', '', '', ''],
        ['note', 'banner_none', 'Neither Telugu nor English selected — no MoCA instrument is opened. Continue to Mandatory History Questions.', '', '', '', "${selected_instrument} = 'none'", '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ============================================================
        # BRANCH 4A: ICMR-NCTB TELUGU MoCA (Version 7.1)
        # ============================================================
        ['begin group', 'te_g_visuospatial', 'ICMR-NCTB Telugu MoCA — VISUOSPATIAL / EXECUTIVE', 'విజువోస్పేషియల్ / ఎగ్జిక్యూటివ్', '', 'field-list', is_te, '', '', '', '', '', '', ''],
        ['note', 'te_trail_making_inst', '1 − క - 2 - ఖ - 3 - గ - 4 - ఘ - 5 - ఙ (మొదలు పెట్టండి -> ఆపండి)', '1 − క - 2 - ఖ - 3 - గ - 4 - ఘ - 5 - ఙ (మొదలు పెట్టండి -> ఆపండి)', '', '', '', '', '', '', '', '', '', 'trail_making_telugu.png'],
        ['select_one yesno', 'te_trail_making', 'Alternating Trail Making (1 - క - 2 - ఖ - 3 - గ - 4 - ఘ - 5 - ఙ)', 'ట్రయల్ మేకింగ్ (1 - క - 2 - ఖ - 3 - గ - 4 - ఘ - 5 - ఙ)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'te_cube_inst', 'చిత్రాన్ని గీయండి.', 'చిత్రాన్ని గీయండి.', '', '', '', '', '', '', '', '', '', 'cube_telugu.png'],
        ['select_one yesno', 'te_cube_copy', 'Cube Copy (చిత్రాన్ని గీయండి.)', 'ఘనం (చిత్రాన్ని గీయండి.)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'te_clock_inst', 'గడియారం గీయండి. (పదకొండు గంటల పది నిమిషాలు) (3 Points)', 'గడియారం గీయండి. (పదకొండు గంటల పది నిమిషాలు) (3 Points)', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_clock_contour', 'Contour', 'గడియార ఆకారం (Contour)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_clock_numbers', 'Numbers', 'సంఖ్యలు (Numbers)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_clock_hands', 'Hands', 'ముళ్ళు (Hands)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'te_g_naming', 'ICMR-NCTB Telugu MoCA — NAMING', 'పేర్లు చెప్పడం (NAMING)', '', 'field-list', is_te, '', '', '', '', '', '', ''],
        ['note', 'te_naming_inst', 'ఈ జంతువు పేరు చెప్పండి.', 'ఈ జంతువు పేరు చెప్పండి.', '', '', '', '', '', '', '', '', '', 'naming_telugu.png'],
        ['select_one yesno', 'te_naming_1', '1. Lion', '1. సింహం (Lion)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_naming_2', '2. Rhinoceros / Rhino', '2. ఖడ్గమృగం (Rhinoceros)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_naming_3', '3. Camel / Dromedary', '3. ఒంటె (Camel)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'te_g_memory', 'ICMR-NCTB Telugu MoCA — MEMORY (No points)', 'జ్ఞాపకశక్తి (No points)', '', 'field-list', is_te, '', '', '', '', '', '', ''],
        ['note', 'te_memory_inst_trial1', 'పదాల జాబితా చదవండి, పేషెంట్/సబ్జెక్ట్ ఆ పదాలను తిరిగి చెప్పాలి. మొదటి ప్రయత్నం సరి అయినా కూడా రెండు ప్రయత్నాలు చేయండి. 5 నిమిషాల తరువాత మళ్ళీ అడగండి.', 'పదాల జాబితా చదవండి, పేషెంట్/సబ్జెక్ట్ ఆ పదాలను తిరిగి చెప్పాలి. మొదటి ప్రయత్నం సరి అయినా కూడా రెండు ప్రయత్నాలు చేయండి. 5 నిమిషాల తరువాత మళ్ళీ అడగండి.', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'te_memory_words', 'కన్ను | చీర | గుడి | గులాబీ | నీలం', 'కన్ను | చీర | గుడి | గులాబీ | నీలం', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'te_mem_trial1', '1వ ప్రయత్నం (1st trial)', '1వ ప్రయత్నం', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_mem_t1_kannu', 'కన్ను (1వ ప్రయత్నం)', 'కన్ను (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_mem_t1_cheera', 'చీర (1వ ప్రయత్నం)', 'చీర (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_mem_t1_gudi', 'గుడి (1వ ప్రయత్నం)', 'గుడి (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_mem_t1_gulabi', 'గులాబీ (1వ ప్రయత్నం)', 'గులాబీ (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_mem_t1_neelam', 'నీలం (1వ ప్రయత్నం)', 'నీలం (1వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'te_mem_trial2', '2వ ప్రయత్నం (2nd trial)', '2వ ప్రయత్నం', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_mem_t2_kannu', 'కన్ను (2వ ప్రయత్నం)', 'కన్ను (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_mem_t2_cheera', 'చీర (2వ ప్రయత్నం)', 'చీర (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_mem_t2_gudi', 'గుడి (2వ ప్రయత్నం)', 'గుడి (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_mem_t2_gulabi', 'గులాబీ (2వ ప్రయత్నం)', 'గులాబీ (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_mem_t2_neelam', 'నీలం (2వ ప్రయత్నం)', 'నీలం (2వ ప్రయత్నం)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'te_g_attention', 'ICMR-NCTB Telugu MoCA — ATTENTION', 'అవధానం (ATTENTION)', '', 'field-list', is_te, '', '', '', '', '', '', ''],
        ['note', 'te_digits_inst', 'ఈ అంకెలు చదవండి. [ 1 Digit / sec. ]', 'ఈ అంకెలు చదవండి. [ 1 Digit / sec. ]', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_digit_span_fwd', 'పేషెంట్/సబ్జెక్ట్ అంకెలను వరుస క్రమంలో చెప్పాలి: 2 1 8 5 4', 'పేషెంట్/సబ్జెక్ట్ అంకెలను వరుస క్రమంలో చెప్పాలి: 2 1 8 5 4', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_digit_span_bwd', 'పేషెంట్/సబ్జెక్ట్ అంకెలను వెనుక క్రమంలో చెప్పాలి: 7 4 2', 'పేషెంట్/సబ్జెక్ట్ అంకెలను వెనుక క్రమంలో చెప్పాలి: 7 4 2', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'te_vigilance_inst', "అక్షరాల జాబితా చదవండి. 'అ' అనే అక్షరం విన్నప్పుడల్లా పేషెంట్/వ్యక్తి చేయి తో ఒకసారి తట్టాలి. No points if >= 2 errors: ప బ అ క మ న అ అ జ క ల బ అ ఫ అ క డ ఈ అ అ అ జ అ మా ఓ ఫ అ అ బ", "అక్షరాల జాబితా చదవండి. 'అ' అనే అక్షరం విన్నప్పుడల్లా పేషెంట్/వ్యక్తి చేయి తో ఒకసారి తట్టాలి. No points if >= 2 errors: ప బ అ క మ న అ అ జ క ల బ అ ఫ అ క డ ఈ అ అ అ జ అ మా ఓ ఫ అ అ బ", '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_vigilance', "Vigilance ('అ' — No points if >= 2 errors)", "అక్షరాల జాబితా ('అ' — No points if >= 2 errors)", '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'te_serial7_inst', '100 నుండి క్రమంగా 7 ను తీసివేయండి. [ ] 93  [ ] 86  [ ] 79  [ ] 72  [ ] 65', '100 నుండి క్రమంగా 7 ను తీసివేయండి. [ ] 93  [ ] 86  [ ] 79  [ ] 72  [ ] 65', '', '', '', '', '', '', '', '', '', ''],
        ['select_one serial7_score', 'te_serial_7s', 'Serial 7s (100 నుండి క్రమంగా 7 ను తీసివేయండి)', '100 నుండి 7 తీసివేయడం', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'te_g_language', 'ICMR-NCTB Telugu MoCA — LANGUAGE', 'భాష (LANGUAGE)', '', 'field-list', is_te, '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_sentence_rep_1', 'తిరిగి చెప్పండి: రాజు ఒక్కడే ఈ రోజు సహాయం చేస్తాడని నేను అనుకుంటున్నాను.', 'తిరిగి చెప్పండి: రాజు ఒక్కడే ఈ రోజు సహాయం చేస్తాడని నేను అనుకుంటున్నాను.', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_sentence_rep_2', 'కుక్కను గదిలో చూసినప్పుడల్లా పిల్లి మంచం కింద దాక్కుంటుంది.', 'కుక్కను గదిలో చూసినప్పుడల్లా పిల్లి మంచం కింద దాక్కుంటుంది.', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['integer', 'te_fluency_count', "ఒక నిమిషంలో 'ప' అక్షరంతో మొదలయ్యే పదాలు ఎన్ని ఎక్కువ చెప్పగలిగితే అన్ని చెప్పండి (N >= 11 words)", "ఒక నిమిషంలో 'ప' అక్షరంతో మొదలయ్యే పదాలు (N >= 11 words)", '', '', '', '. >= 0 and . <= 100', 'Enter a count between 0 and 100.', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'te_g_abstraction', 'ICMR-NCTB Telugu MoCA — ABSTRACTION', 'పోలికలు (ABSTRACTION)', '', 'field-list', is_te, '', '', '', '', '', '', ''],
        ['note', 'te_abstraction_inst_practice', 'పోలికలు : టమాట - బంగాళాదుంప = కూరగాయలు', 'పోలికలు : టమాట - బంగాళాదుంప = కూరగాయలు', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_abstraction_1', 'రైలు - సైకిల్', 'రైలు - సైకిల్', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_abstraction_2', 'గడియారం - స్కేలు', 'గడియారం - స్కేలు', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'te_g_delayed_recall', 'ICMR-NCTB Telugu MoCA — DELAYED RECALL', 'ఆలస్య జ్ఞాపకశక్తి (DELAYED RECALL)', '', 'field-list', is_te, '', '', '', '', '', '', ''],
        ['note', 'te_delayed_inst', 'ఇంతకుముందు చెప్పిన పదాలను చెప్పండి. (ఎటువంటి సూచన లేకుండా) — Points for UNCUED recall only', 'ఇంతకుముందు చెప్పిన పదాలను చెప్పండి. (ఎటువంటి సూచన లేకుండా) — Points for UNCUED recall only', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_delayed_kannu', 'కన్ను (ఎటువంటి సూచన లేకుండా)', 'కన్ను (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_kannu_cat', 'కన్ను — Category cue: ముఖంలో ఒక భాగం.', 'కన్ను — Category cue: ముఖంలో ఒక భాగం.', '', '', "${te_delayed_kannu} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_kannu_mc', 'కన్ను — Multiple choice: కన్ను, ముక్కు, చెవి', 'కన్ను — Multiple choice: కన్ను, ముక్కు, చెవి', '', '', "${te_delayed_kannu} = 'no' and ${te_delayed_kannu_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_delayed_cheera', 'చీర (ఎటువంటి సూచన లేకుండా)', 'చీర (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_cheera_cat', 'చీర — Category cue: దుస్తులు/వస్త్రాలలో ఒక రకం.', 'చీర — Category cue: దుస్తులు/వస్త్రాలలో ఒక రకం.', '', '', "${te_delayed_cheera} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_cheera_mc', 'చీర — Multiple choice: వోణి, పంచ, చీర', 'చీర — Multiple choice: వోణి, పంచ, చీర', '', '', "${te_delayed_cheera} = 'no' and ${te_delayed_cheera_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_delayed_gudi', 'గుడి (ఎటువంటి సూచన లేకుండా)', 'గుడి (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_gudi_cat', 'గుడి — Category cue: భవనాలలో ఒక రకం.', 'గుడి — Category cue: భవనాలలో ఒక రకం.', '', '', "${te_delayed_gudi} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_gudi_mc', 'గుడి — Multiple choice: గుడి, బడి, ఆసుపత్రి', 'గుడి — Multiple choice: గుడి, బడి, ఆసుపత్రి', '', '', "${te_delayed_gudi} = 'no' and ${te_delayed_gudi_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_delayed_gulabi', 'గులాబీ (ఎటువంటి సూచన లేకుండా)', 'గులాబీ (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_gulabi_cat', 'గులాబీ — Category cue: పూలలో ఒక రకం.', 'గులాబీ — Category cue: పూలలో ఒక రకం.', '', '', "${te_delayed_gulabi} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_gulabi_mc', 'గులాబీ — Multiple choice: మల్లె, గులాబీ, చామంతి', 'గులాబీ — Multiple choice: మల్లె, గులాబీ, చామంతి', '', '', "${te_delayed_gulabi} = 'no' and ${te_delayed_gulabi_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_delayed_neelam', 'నీలం (ఎటువంటి సూచన లేకుండా)', 'నీలం (ఎటువంటి సూచన లేకుండా)', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_neelam_cat', 'నీలం — Category cue: ఒక రంగు.', 'నీలం — Category cue: ఒక రంగు.', '', '', "${te_delayed_neelam} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'te_delayed_neelam_mc', 'నీలం — Multiple choice: తెలుపు, ఎరుపు, నీలం', 'నీలం — Multiple choice: తెలుపు, ఎరుపు, నీలం', '', '', "${te_delayed_neelam} = 'no' and ${te_delayed_neelam_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'te_g_orientation', 'ICMR-NCTB Telugu MoCA — ORIENTATION', 'దిక్సూచి (ORIENTATION)', '', 'field-list', is_te, '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_orientation_date', 'తేది (Date)', 'తేది', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_orientation_month', 'నెల (Month)', 'నెల', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_orientation_year', 'సంవత్సరం (Year)', 'సంవత్సరం', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_orientation_day', 'రోజు (Day)', 'రోజు', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_orientation_place', 'స్థలం (Place)', 'స్థలం', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_orientation_city', 'నగరం (City)', 'నగరం', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ============================================================
        # BRANCH 4B: MoCA-BASIC (MoCA-B)
        # ============================================================
        ['begin group', 'mb_g_timing', 'MoCA-Basic (MoCA-B) — Timing', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['time', 'mb_start_time', 'START TIME', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_executive', 'MoCA-B — EXECUTIVE FUNCTION', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_trail_making_inst', 'Please draw a line alternating between a square with a number and a square with dots in increasing order. Begin here [1] and draw a line to [1 dot], then [2], then [2 dots], and so on. End here [6 dots]. 1 point if completed on first attempt without any errors or self-corrections.', '', '', '', '', '', '', '', '', '', '', 'trail_making.png'],
        ['select_one yesno', 'mb_trail_making', 'Alternating Trail Making (SCORE /1)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_memory', 'MoCA-B — IMMEDIATE RECALL (No point)', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_memory_inst_trial1', 'Perform 2 trials even if 1st trial is successful. No point.', '', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'mb_memory_words', 'ROSE | CHAIR | HAND | BLUE | SPOON', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'mb_mem_trial1', '1st trial', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t1_rose', 'ROSE (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t1_chair', 'CHAIR (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t1_hand', 'HAND (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t1_blue', 'BLUE (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t1_spoon', 'SPOON (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'mb_mem_trial2', '2nd trial', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t2_rose', 'ROSE (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t2_chair', 'CHAIR (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t2_hand', 'HAND (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t2_blue', 'BLUE (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_mem_t2_spoon', 'SPOON (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_fluency', 'MoCA-B — FLUENCY', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_fluency_inst', 'Name maximum numbers of FRUITS in 1 minute. 2 points if N = 13 or more; 1 point if N = 8-12; 0 point if N = 7 or less.', '', '', '', '', '', '', '', '', '', '', ''],
        ['text', 'mb_fluency_words', 'Fruits named (1-18)', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['integer', 'mb_fluency_count', 'N items', '', '', '', '', '. >= 0 and . <= 100', 'Enter a count between 0 and 100.', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_orientation', 'MoCA-B — ORIENTATION', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_orientation_inst', '1 point is allocated for each correct answer.', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mb_orientation_time', 'time (± 2 hr)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_orientation_day', 'day', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_orientation_month', 'month', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_orientation_year', 'year', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_orientation_place', 'place', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_orientation_city', 'city', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_calculation', 'MoCA-B — CALCULATION', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_calculation_inst', 'Provide 3 ways to pay using 1 dollar coins, 5 dollar and 10 dollar bills for an object that costs exactly "13 Dollars". (3 points if 3 ways, 2 points if 2 ways, 1 point if 1 way, 0 point if no correct way)', '', '', '', '', '', '', '', '', '', '', ''],
        ['text', 'mb_calc_combinations', '1. / 2. / 3.', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['select_one calc_score', 'mb_calculation_score_select', 'SCORE ( /3)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_abstraction', 'MoCA-B — ABSTRACTION', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_abstraction_inst_practice', 'To what category these objects belong to? (e.g. orange - banana = fruit)', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mb_abstraction_1', 'train - boat', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_abstraction_2', 'north - south', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_abstraction_3', 'drum - flute', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_delayed_recall', 'MoCA-B — DELAYED RECALL', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_delayed_inst', 'Points are awarded for recall with No cue (1 point for each item).', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mb_delayed_rose', 'ROSE - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_rose_cat', 'ROSE - Recall with category cue (type of flower)', '', '', '', "${mb_delayed_rose} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_rose_mc', 'ROSE - Recall with multiple choice cue (rose / daisy / tulip)', '', '', '', "${mb_delayed_rose} = 'no' and ${mb_delayed_rose_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_delayed_chair', 'CHAIR - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_chair_cat', 'CHAIR - Recall with category cue (type of furniture)', '', '', '', "${mb_delayed_chair} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_chair_mc', 'CHAIR - Recall with multiple choice cue (table / chair / bed)', '', '', '', "${mb_delayed_chair} = 'no' and ${mb_delayed_chair_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_delayed_hand', 'HAND - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_hand_cat', 'HAND - Recall with category cue (body part)', '', '', '', "${mb_delayed_hand} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_hand_mc', 'HAND - Recall with multiple choice cue (foot / hand / knee)', '', '', '', "${mb_delayed_hand} = 'no' and ${mb_delayed_hand_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_delayed_blue', 'BLUE - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_blue_cat', 'BLUE - Recall with category cue (colour)', '', '', '', "${mb_delayed_blue} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_blue_mc', 'BLUE - Recall with multiple choice cue (blue / brown / red)', '', '', '', "${mb_delayed_blue} = 'no' and ${mb_delayed_blue_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_delayed_spoon', 'SPOON - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_spoon_cat', 'SPOON - Recall with category cue (kitchen instrument)', '', '', '', "${mb_delayed_spoon} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'mb_delayed_spoon_mc', 'SPOON - Recall with multiple choice cue (fork / spoon / knife)', '', '', '', "${mb_delayed_spoon} = 'no' and ${mb_delayed_spoon_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_visuoperception', 'MoCA-B — VISUOPERCEPTION', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_visuoperception_inst', 'Identify drawings. No more than 60 seconds. See complementary sheet. 3 points if N = 9-10; 2 points if N = 6-8; 1 point if N = 4-5; 0 point if N = 0-3.', '', '', '', '', '', '', '', '', '', '', 'visuoperception.png'],
        ['select_one yesno', 'mb_vp_scissors', 'scissors', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_vp_tshirt', 'T-shirt', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_vp_banana', 'banana', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_vp_lamp', 'lamp', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_vp_candle', 'candle', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_vp_watch', 'watch', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_vp_cup', 'cup', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_vp_leaf', 'leaf', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_vp_key', 'key', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_vp_spoon', 'spoon', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_naming', 'MoCA-B — NAMING', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_naming_inst', 'Identify animals. See complementary sheet.', '', '', '', '', '', '', '', '', '', '', 'naming_b.png'],
        ['select_one yesno', 'mb_naming_zebra', 'zebra', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_naming_peacock', 'peacock', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_naming_tiger', 'tiger', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mb_naming_butterfly', 'butterfly', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'mb_g_attention', 'MoCA-B — ATTENTION', '', '', 'field-list', is_mb, '', '', '', '', '', '', ''],
        ['note', 'mb_attention_white_inst', 'Name the numbers in circles. See complementary sheet. No point if 2 errors or more.', '', '', '', '', '', '', '', '', '', '', 'attention_white.png'],
        ['integer', 'mb_attention_white_errors', 'ERROR N (White background)', '', '', '', '', '. >= 0 and . <= 50', 'Enter errors between 0 and 50.', 'yes', '', '', '', ''],
        ['note', 'mb_attention_black_inst', 'Name the numbers in circles & squares. See complementary sheet. 2 points if 2 errors or less; 1 point if 3 errors; 0 point if 4 errors or more.', '', '', '', '', '', '', '', '', '', '', 'attention_black.png'],
        ['integer', 'mb_attention_black_errors', 'ERROR N (Black background)', '', '', '', '', '. >= 0 and . <= 50', 'Enter errors between 0 and 50.', 'yes', '', '', '', ''],
        ['time', 'mb_end_time', 'END TIME', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['integer', 'mb_total_time_min', 'TOTAL TIME (min)', '', '', '', '', '. >= 0 and . <= 180', 'Enter minutes between 0 and 180.', 'no', '', '', '', ''],
        ['integer', 'mb_total_time_sec', 'TOTAL TIME (sec)', '', '', '', '', '. >= 0 and . <= 59', 'Enter seconds between 0 and 59.', 'no', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ============================================================
        # BRANCH 4C: MoCA FULL, ENGLISH, VERSION 8.1
        # ============================================================
        ['begin group', 'm8_g_visuospatial', 'MoCA 8.1 — VISUOSPATIAL / EXECUTIVE', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_trail_inst', 'Please draw a line going from a number to a letter in ascending order (1-A-2-B-3-C-4-D-5-E). 1 point if completed without any lines crossing and without connecting E to 1.', '', '', '', '', '', '', '', '', '', '', 'trail_making_8_1.png'],
        ['select_one yesno', 'm8_trail_making', 'Alternating Trail Making (1-A-2-B-3-C-4-D-5-E)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'm8_cube_inst', 'Copy cube: 3D, all lines drawn, meet with little/no space, no line added, parallel/similar length, orientation preserved.', '', '', '', '', '', '', '', '', '', '', 'cube_8_1.png'],
        ['select_one yesno', 'm8_cube_copy', 'Copy cube', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'm8_clock_inst', 'Draw CLOCK (Ten past eleven) (3 points).', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_clock_contour', 'Contour (1 pt)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_clock_numbers', 'Numbers (1 pt)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_clock_hands', 'Hands (1 pt)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_naming', 'MoCA 8.1 — NAMING', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_naming_inst', 'Beginning on the left, point to each figure and say: "Tell me the name of this animal."', '', '', '', '', '', '', '', '', '', '', 'naming_8_1.png'],
        ['select_one yesno', 'm8_naming_1', '1. Lion', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_naming_2', '2. Rhinoceros or rhino', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_naming_3', '3. Camel or dromedary', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_memory', 'MoCA 8.1 — MEMORY (No points)', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_memory_inst', 'Read list of words (1/sec), subject must repeat them. Do 2 trials, even if 1st trial is successful. Do a recall after 5 minutes.', '', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'm8_memory_words', 'FACE | VELVET | CHURCH | DAISY | RED', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'm8_mem_trial1', '1st TRIAL', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_face', 'FACE (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_velvet', 'VELVET (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_church', 'CHURCH (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_daisy', 'DAISY (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_red', 'RED (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'm8_mem_trial2', '2nd TRIAL', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_face', 'FACE (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_velvet', 'VELVET (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_church', 'CHURCH (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_daisy', 'DAISY (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_red', 'RED (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_attention', 'MoCA 8.1 — ATTENTION', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_digits_inst', 'Read list of digits (1 digit/sec.).', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_digit_span_fwd', 'Forward order: 2 1 8 5 4', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_digit_span_bwd', 'Backward order: 7 4 2', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'm8_vigilance_inst', 'Read list of letters: F B A C M N A A J K L B A F A K D E A A A J A M O F A A B. Tap hand at each letter A. No points if >= 2 errors.', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_vigilance', 'Vigilance (0-1 error = 1 pt; >= 2 errors = 0 pt)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'm8_serial7_inst', 'Serial 7 subtraction starting at 100: [ ] 93  [ ] 86  [ ] 79  [ ] 72  [ ] 65', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one serial7_score', 'm8_serial_7s', 'Serial 7s (100 - 7)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_language', 'MoCA 8.1 — LANGUAGE', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_sentence_rep_1', 'Repeat: I only know that John is the one to help today.', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_sentence_rep_2', 'Repeat: The cat always hid under the couch when dogs were in the room.', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['integer', 'm8_fluency_count', 'Fluency: Number of words in one minute beginning with letter F (N >= 11 words)', '', '', '', '', '. >= 0 and . <= 100', 'Enter a count between 0 and 100.', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_abstraction', 'MoCA 8.1 — ABSTRACTION', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_abstraction_inst', 'Similarity between e.g. banana - orange = fruit', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_abstraction_1', 'train - bicycle', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_abstraction_2', 'watch - ruler', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_delayed_recall', 'MoCA 8.1 — DELAYED RECALL & MEMORY INDEX SCORE (MIS)', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_delayed_inst', 'Points for UNCUED recall only (5 pts max). MIS = 3x uncued + 2x category cue + 1x multiple-choice cue (/15).', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_face', 'FACE', 'Category cue: body part | MC: nose, face, hand (shoulder, leg)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_velvet', 'VELVET', 'Category cue: type of fabric | MC: denim, velvet, cotton (nylon, silk)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_church', 'CHURCH', 'Category cue: type of building | MC: church, school, hospital (library, store)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_daisy', 'DAISY', 'Category cue: type of flower | MC: rose, daisy, tulip (lily, daffodil)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_red', 'RED', 'Category cue: color | MC: red, blue, green (yellow, purple)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_orientation', 'MoCA 8.1 — ORIENTATION', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_orientation_date', 'Date', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_orientation_month', 'Month', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_orientation_year', 'Year', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_orientation_day', 'Day', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_orientation_place', 'Place', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_orientation_city', 'City', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ============================================================
        # 5. MANDATORY HISTORY QUESTIONS (Annex A Part 5)
        # ============================================================
        ['begin group', 'g_mandatory_history', '5. MANDATORY HISTORY QUESTIONS', '5. తప్పనిసరి చరిత్ర ప్రశ్నలు', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'mh_inst', 'Ask the patient and the informant separately. Do not prompt or suggest answers.', 'రోగిని మరియు సమాచారకర్తను విడివిడిగా అడగండి. సమాధానాలను సూచించవద్దు.', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q1', '1. Is there any complaint of memory or thinking difficulty?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['text', 'mh_q1_duration', 'If Q1 is Yes — difficulty first began (months / years ago)', '', '', '', "${mh_q1} = 'yes'", '', '', 'yes', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q2', '2. Did it begin gradually rather than suddenly?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q3', '3. Has it worsened over the last six months?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q4', '4. Does the patient repeat the same questions or stories?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q5', '5. Any difficulty managing money, medicines or the phone?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q6', '6. Has the patient become lost in a familiar place?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q7', '7. Any change in personality, mood or behaviour?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q8', '8. Any stroke, head injury, fits or loss of consciousness?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q9', '9. Is the patient currently on any sedating or psychiatric medication?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yes_no_unsure', 'mh_q10', '10. Has the patient been tested for memory before?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['text', 'mh_q10_details', 'If Q10 is Yes — where and when', '', '', '', "${mh_q10} = 'yes'", '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ============================================================
        # SCORING CALCULATIONS (all 3 branches + unified totals)
        # ============================================================
        # Telugu 7.1
        ['calculate', 'te_visuospatial_total', '', '', '', '', '', '', '', '', 'if(${te_trail_making} = "yes", 1, 0) + if(${te_cube_copy} = "yes", 1, 0) + if(${te_clock_contour} = "yes", 1, 0) + if(${te_clock_numbers} = "yes", 1, 0) + if(${te_clock_hands} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'te_naming_total', '', '', '', '', '', '', '', '', 'if(${te_naming_1} = "yes", 1, 0) + if(${te_naming_2} = "yes", 1, 0) + if(${te_naming_3} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'te_attention_total', '', '', '', '', '', '', '', '', 'if(${te_digit_span_fwd} = "yes", 1, 0) + if(${te_digit_span_bwd} = "yes", 1, 0) + if(${te_vigilance} = "yes", 1, 0) + if(${te_serial_7s} != "", number(${te_serial_7s}), 0)', '', '', ''],
        ['calculate', 'te_language_total', '', '', '', '', '', '', '', '', 'if(${te_sentence_rep_1} = "yes", 1, 0) + if(${te_sentence_rep_2} = "yes", 1, 0) + if(${te_fluency_count} != "" and ${te_fluency_count} >= 11, 1, 0)', '', '', ''],
        ['calculate', 'te_abstraction_total', '', '', '', '', '', '', '', '', 'if(${te_abstraction_1} = "yes", 1, 0) + if(${te_abstraction_2} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'te_memory_total', '', '', '', '', '', '', '', '', 'if(${te_delayed_kannu} = "yes", 1, 0) + if(${te_delayed_cheera} = "yes", 1, 0) + if(${te_delayed_gudi} = "yes", 1, 0) + if(${te_delayed_gulabi} = "yes", 1, 0) + if(${te_delayed_neelam} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'te_orientation_total', '', '', '', '', '', '', '', '', 'if(${te_orientation_date} = "yes", 1, 0) + if(${te_orientation_month} = "yes", 1, 0) + if(${te_orientation_year} = "yes", 1, 0) + if(${te_orientation_day} = "yes", 1, 0) + if(${te_orientation_place} = "yes", 1, 0) + if(${te_orientation_city} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'te_raw_total', '', '', '', '', '', '', '', '', '${te_visuospatial_total} + ${te_naming_total} + ${te_attention_total} + ${te_language_total} + ${te_abstraction_total} + ${te_memory_total} + ${te_orientation_total}', '', '', ''],
        ['calculate', 'te_education_adj', '', '', '', '', '', '', '', '', 'if(${years_schooling} != "" and number(${years_schooling}) <= 12, 1, 0)', '', '', ''],
        ['calculate', 'te_total_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_telugu", min(${te_raw_total} + ${te_education_adj}, 30), "")', '', '', ''],

        # MoCA-B
        ['calculate', 'mb_executive_total', '', '', '', '', '', '', '', '', 'if(${mb_trail_making} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mb_fluency_total', '', '', '', '', '', '', '', '', 'if(${mb_fluency_count} != "" and ${mb_fluency_count} >= 13, 2, if(${mb_fluency_count} != "" and ${mb_fluency_count} >= 8, 1, 0))', '', '', ''],
        ['calculate', 'mb_orientation_total', '', '', '', '', '', '', '', '', 'if(${mb_orientation_time} = "yes", 1, 0) + if(${mb_orientation_day} = "yes", 1, 0) + if(${mb_orientation_month} = "yes", 1, 0) + if(${mb_orientation_year} = "yes", 1, 0) + if(${mb_orientation_place} = "yes", 1, 0) + if(${mb_orientation_city} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mb_calculation_total', '', '', '', '', '', '', '', '', 'if(${mb_calculation_score_select} != "", number(${mb_calculation_score_select}), 0)', '', '', ''],
        ['calculate', 'mb_abstraction_total', '', '', '', '', '', '', '', '', 'if(${mb_abstraction_1} = "yes", 1, 0) + if(${mb_abstraction_2} = "yes", 1, 0) + if(${mb_abstraction_3} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mb_memory_total', '', '', '', '', '', '', '', '', 'if(${mb_delayed_rose} = "yes", 1, 0) + if(${mb_delayed_chair} = "yes", 1, 0) + if(${mb_delayed_hand} = "yes", 1, 0) + if(${mb_delayed_blue} = "yes", 1, 0) + if(${mb_delayed_spoon} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mb_visuoperception_count', '', '', '', '', '', '', '', '', 'if(${mb_vp_scissors} = "yes", 1, 0) + if(${mb_vp_tshirt} = "yes", 1, 0) + if(${mb_vp_banana} = "yes", 1, 0) + if(${mb_vp_lamp} = "yes", 1, 0) + if(${mb_vp_candle} = "yes", 1, 0) + if(${mb_vp_watch} = "yes", 1, 0) + if(${mb_vp_cup} = "yes", 1, 0) + if(${mb_vp_leaf} = "yes", 1, 0) + if(${mb_vp_key} = "yes", 1, 0) + if(${mb_vp_spoon} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mb_visuoperception_total', '', '', '', '', '', '', '', '', 'if(${mb_visuoperception_count} >= 9, 3, if(${mb_visuoperception_count} >= 6, 2, if(${mb_visuoperception_count} >= 4, 1, 0)))', '', '', ''],
        ['calculate', 'mb_naming_total', '', '', '', '', '', '', '', '', 'if(${mb_naming_zebra} = "yes", 1, 0) + if(${mb_naming_peacock} = "yes", 1, 0) + if(${mb_naming_tiger} = "yes", 1, 0) + if(${mb_naming_butterfly} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mb_attention_white_score', '', '', '', '', '', '', '', '', 'if(${mb_attention_white_errors} != "" and ${mb_attention_white_errors} <= 1, 1, 0)', '', '', ''],
        ['calculate', 'mb_attention_black_score', '', '', '', '', '', '', '', '', 'if(${mb_attention_black_errors} != "" and ${mb_attention_black_errors} <= 2, 2, if(${mb_attention_black_errors} = 3, 1, 0))', '', '', ''],
        ['calculate', 'mb_attention_total', '', '', '', '', '', '', '', '', '${mb_attention_white_score} + ${mb_attention_black_score}', '', '', ''],
        ['calculate', 'mb_raw_total', '', '', '', '', '', '', '', '', '${mb_executive_total} + ${mb_fluency_total} + ${mb_orientation_total} + ${mb_calculation_total} + ${mb_abstraction_total} + ${mb_memory_total} + ${mb_visuoperception_total} + ${mb_naming_total} + ${mb_attention_total}', '', '', ''],
        ['calculate', 'mb_education_adj', '', '', '', '', '', '', '', '', 'if(${years_schooling} != "" and number(${years_schooling}) < 4, 1, 0)', '', '', ''],
        ['calculate', 'mb_literacy_adj', '', '', '', '', '', '', '', '', 'if(${can_read_write} = "no", 1, 0)', '', '', ''],
        ['calculate', 'mb_total_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_b", min(${mb_raw_total} + ${mb_education_adj} + ${mb_literacy_adj}, 30), "")', '', '', ''],

        # MoCA 8.1
        ['calculate', 'm8_visuospatial_total', '', '', '', '', '', '', '', '', 'if(${m8_trail_making} = "yes", 1, 0) + if(${m8_cube_copy} = "yes", 1, 0) + if(${m8_clock_contour} = "yes", 1, 0) + if(${m8_clock_numbers} = "yes", 1, 0) + if(${m8_clock_hands} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'm8_naming_total', '', '', '', '', '', '', '', '', 'if(${m8_naming_1} = "yes", 1, 0) + if(${m8_naming_2} = "yes", 1, 0) + if(${m8_naming_3} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'm8_attention_total', '', '', '', '', '', '', '', '', 'if(${m8_digit_span_fwd} = "yes", 1, 0) + if(${m8_digit_span_bwd} = "yes", 1, 0) + if(${m8_vigilance} = "yes", 1, 0) + if(${m8_serial_7s} != "", number(${m8_serial_7s}), 0)', '', '', ''],
        ['calculate', 'm8_language_total', '', '', '', '', '', '', '', '', 'if(${m8_sentence_rep_1} = "yes", 1, 0) + if(${m8_sentence_rep_2} = "yes", 1, 0) + if(${m8_fluency_count} != "" and ${m8_fluency_count} >= 11, 1, 0)', '', '', ''],
        ['calculate', 'm8_abstraction_total', '', '', '', '', '', '', '', '', 'if(${m8_abstraction_1} = "yes", 1, 0) + if(${m8_abstraction_2} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'm8_memory_total', '', '', '', '', '', '', '', '', 'if(${m8_delayed_face} = "uncued", 1, 0) + if(${m8_delayed_velvet} = "uncued", 1, 0) + if(${m8_delayed_church} = "uncued", 1, 0) + if(${m8_delayed_daisy} = "uncued", 1, 0) + if(${m8_delayed_red} = "uncued", 1, 0)', '', '', ''],
        ['calculate', 'm8_mis_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_8_1", (if(${m8_delayed_face} = "uncued", 3, if(${m8_delayed_face} = "category", 2, if(${m8_delayed_face} = "multiple_choice", 1, 0))) + if(${m8_delayed_velvet} = "uncued", 3, if(${m8_delayed_velvet} = "category", 2, if(${m8_delayed_velvet} = "multiple_choice", 1, 0))) + if(${m8_delayed_church} = "uncued", 3, if(${m8_delayed_church} = "category", 2, if(${m8_delayed_church} = "multiple_choice", 1, 0))) + if(${m8_delayed_daisy} = "uncued", 3, if(${m8_delayed_daisy} = "category", 2, if(${m8_delayed_daisy} = "multiple_choice", 1, 0))) + if(${m8_delayed_red} = "uncued", 3, if(${m8_delayed_red} = "category", 2, if(${m8_delayed_red} = "multiple_choice", 1, 0)))), "")', '', '', ''],
        ['calculate', 'm8_orientation_total', '', '', '', '', '', '', '', '', 'if(${m8_orientation_date} = "yes", 1, 0) + if(${m8_orientation_month} = "yes", 1, 0) + if(${m8_orientation_year} = "yes", 1, 0) + if(${m8_orientation_day} = "yes", 1, 0) + if(${m8_orientation_place} = "yes", 1, 0) + if(${m8_orientation_city} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'm8_raw_total', '', '', '', '', '', '', '', '', '${m8_visuospatial_total} + ${m8_naming_total} + ${m8_attention_total} + ${m8_language_total} + ${m8_abstraction_total} + ${m8_memory_total} + ${m8_orientation_total}', '', '', ''],
        ['calculate', 'm8_education_adj', '', '', '', '', '', '', '', '', 'if(${years_schooling} != "" and number(${years_schooling}) <= 12, 1, 0)', '', '', ''],
        ['calculate', 'm8_total_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_8_1", min(${m8_raw_total} + ${m8_education_adj}, 30), "")', '', '', ''],

        # Unified totals
        ['calculate', 'moca_total_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_telugu", ${te_total_score}, if(${selected_instrument} = "moca_b", ${mb_total_score}, if(${selected_instrument} = "moca_8_1", ${m8_total_score}, "")))', '', '', ''],
        ['calculate', 'moca_mis_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_8_1", ${m8_mis_score}, "")', '', '', ''],
        ['calculate', 'moca_interpretation', '', '', '', '', '', '', '', '', 'if(${consent_reconfirmed} != "yes", "not assessed (no consent)", if(${selected_instrument} = "moca_telugu" or ${selected_instrument} = "moca_8_1", if(${moca_total_score} >= 26, "normal", "impaired"), ""))', '', '', ''],

        # Admin / Summary
        ['begin group', 'g_admin', 'TOTAL & SIGN-OFF', 'మొత్తం స్కోరు', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['string', 'administered_by', 'Administered by', 'పరీక్ష నిర్వహించినవారు', '', '', '', '', '', 'yes', '', '', '', ''],
        ['dateTime', 'assessment_datetime', 'Date of exam', 'తేదీ', '', '', '', '', '', 'yes', "if(. = '', now(), .)", '', '', ''],
        ['note', 'score_summary_note', 'Instrument: ${selected_instrument} | TOTAL SCORE: ${moca_total_score} / 30', '', '', '', "${selected_instrument} = 'moca_telugu' or ${selected_instrument} = 'moca_b' or ${selected_instrument} = 'moca_8_1'", '', '', '', '', '', '', ''],
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
        ['yes_no_unsure', 'yes', 'Yes', 'అవును'],
        ['yes_no_unsure', 'no', 'No', 'కాదు'],
        ['yes_no_unsure', 'unsure', 'Unsure', 'తెలియదు'],
        ['luq_lang', 'telugu', 'Telugu', 'తెలుగు'],
        ['luq_lang', 'english', 'English', 'ఇంగ్లీష్'],
        ['luq_lang', 'neither', 'Neither', 'రెండూ కాదు'],
        ['test_lang', 'telugu', 'Telugu', 'తెలుగు'],
        ['test_lang', 'english', 'English', 'ఇంగ్లీష్'],
        ['test_lang', 'neither', 'Neither / Other', 'ఇతర భాష'],
        ['contact_freq', 'daily', 'Daily', 'రోజువారీ'],
        ['contact_freq', 'weekly', 'Weekly', 'వారానికొకసారి'],
        ['contact_freq', 'less_often', 'Less often', 'తక్కువసార్లు'],
        ['other_factor', 'poor_sleep', 'Poor sleep', 'నిద్ర లేకపోవడం'],
        ['other_factor', 'fasting', 'Fasting', 'ఉపవాసం'],
        ['other_factor', 'long_travel', 'Long travel', 'సుదీర్ఘ ప్రయాణం'],
        ['other_factor', 'recent_bereavement', 'Recent bereavement', 'ఇటీవలి విషాదం'],
        ['other_factor', 'sedating_medication', 'Sedating medication', 'మత్తుమందు'],
        ['other_factor', 'none', 'None', 'ఏవీ లేవు'],
        ['serial7_score', '0', '0 correct (0 pt)', '0 సరిగ్గా (0 pt)'],
        ['serial7_score', '1', '1 correct (1 pt)', '1 సరిగ్గా (1 pt)'],
        ['serial7_score', '2', '2 or 3 correct (2 pts)', '2 లేదా 3 సరిగ్గా (2 pts)'],
        ['serial7_score', '3', '4 or 5 correct subtractions (3 pts)', '4 లేదా 5 సరిగ్గా (3 pts)'],
        ['calc_score', '0', '0 point if no correct way', '0 పాయింట్లు'],
        ['calc_score', '1', '1 point if 1 way', '1 పాయింట్'],
        ['calc_score', '2', '2 points if 2 ways', '2 పాయింట్లు'],
        ['calc_score', '3', '3 points if 3 ways', '3 పాయింట్లు'],
        ['delayed_cue', 'yes', 'Recalled with cue', 'సూచనతో గుర్తుకొచ్చింది'],
        ['delayed_cue', 'no', 'Not recalled even with cue', 'సూచనతో కూడా గుర్తురాలేదు'],
        ['m8_recall_type', 'uncued', 'Recalled with NO CUE (1 pt, MIS x3)', 'సూచన లేకుండా గుర్తుకు వచ్చింది'],
        ['m8_recall_type', 'category', 'Recalled with Category cue (0 pt, MIS x2)', 'వర్గ సూచనతో గుర్తుకు వచ్చింది'],
        ['m8_recall_type', 'multiple_choice', 'Recalled with Multiple choice cue (0 pt, MIS x1)', 'బహుళ ఎంపిక సూచనతో గుర్తుకు వచ్చింది'],
        ['m8_recall_type', 'not_recalled', 'Not recalled (0 pt, MIS 0)', 'గుర్తుకు రాలేదు'],
    ]

    for c in choices:
        ws_choices.append(c)

    # ============================================================
    # SETTINGS SHEET
    # ============================================================
    ws_settings = wb.create_sheet('settings')
    ws_settings.append(['form_title', 'form_id', 'version', 'default_language'])
    ws_settings.append(['MOCA (Integrated)', 'moca', '2026-11-12 02-00', 'en'])

    os.makedirs(FORMS_APP_DIR, exist_ok=True)
    xlsx_path = os.path.join(FORMS_APP_DIR, 'moca.xlsx')
    wb.save(xlsx_path)
    print(f'Saved: {xlsx_path}')

    props_path = os.path.join(FORMS_APP_DIR, 'moca.properties.json')
    props = {
        'title': [
            {
                'locale': 'en',
                'content': 'MOCA (Integrated)',
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
    build_moca_xlsx()
