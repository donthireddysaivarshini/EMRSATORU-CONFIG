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
    is_m8 = "${selected_instrument} = 'moca_8_2'"

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
        ['select_one yesno', 'can_read_write', 'Can read and write a simple sentence in your own language?', 'మీ స్వంత భాషలో ఒక సాధారణ వాక్యాన్ని చదవగలరా మరియు వ్రాయగలరా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['calculate', 'education_band', '', '', '', '', '', '', '', '', "if(${years_schooling} != '' and (number(${years_schooling}) <= 4 or ${can_read_write} = 'no'), 'A', if(${years_schooling} != '' and number(${years_schooling}) <= 10, 'B', if(${years_schooling} != '', 'C', '')))", '', '', ''],
        ['note', 'education_band_note', 'Education band: ${education_band} (A: <=4 yrs or illiterate | B: 5-10 yrs | C: >10 yrs)', 'విద్యా స్థాయి: ${education_band}', '', '', "${education_band} != ''", '', '', '', '', '', '', ''],
        ['select_one spoken_lang', 'languages_spoken', 'Language spoken', 'మాట్లాడే భాష', 'Select the participant’s spoken language.', 'minimal', '', '', '', 'yes', '', '', '', ''],
        ['text', 'other_language_spoken', 'What language? (specify other language)', 'ఏ భాష? (ఇతర భాషను పేర్కొనండి)', '', '', "${languages_spoken} = 'other'", '', '', 'yes', '', '', '', ''],
        ['select_one luq_lang', 'luq_proficiency', 'Language Usage Questionnaire (LUQ) — >=70% daily proficiency in:', 'భాషా వినియోగ ప్రశ్నావళి (LUQ) — >=70% ప్రావీణ్యం ఉన్న భాష:', 'Select the language the patient uses at least 70% of the time in daily life based on LUQ.', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
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
        ['note', 'other_factors_heading', 'Other factors today (potential confounding factors):', 'ఈ రోజు ఇతర అంశాలు:', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'factor_poor_sleep', 'Poor sleep?', 'నిద్ర లేకపోవడం?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'factor_fasting', 'Fasting?', 'ఉపవాసం?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'factor_long_travel', 'Long travel before test?', 'సుదీర్ఘ ప్రయాణం?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'factor_recent_bereavement', 'Recent bereavement / grief?', 'ఇటీవలి విషాదం?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'factor_sedating_medication', 'Sedating / psychiatric medication taken today?', 'మత్తుమందు తీసుకున్నారా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'factor_other', 'Any other confounding factor today?', 'ఇతర అంశాలు ఏవైనా ఉన్నాయా?', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['text', 'factor_other_details', 'Details of other factor', 'ఇతర వివరాలు', '', '', "${factor_other} = 'yes'", '', '', 'no', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ------------------------------------------------------------
        # 4. INSTRUMENT SELECTION (Table 8.2) & CLINICAL START
        # ------------------------------------------------------------
        ['calculate', 'selected_instrument', '', '', '', '', '', '', '', '', "if(${consent_reconfirmed} != 'yes', '', if(${test_language} = 'telugu', 'moca_telugu', if(${test_language} = 'english' and ${education_band} = 'A', 'moca_b', if(${test_language} = 'english' and (${education_band} = 'B' or ${education_band} = 'C'), 'moca_8_2', 'none'))))", '', '', ''],
        ['calculate', 'education_confounded', '', '', '', '', '', '', '', '', "if(${selected_instrument} = 'moca_telugu' and ${education_band} = 'A', 'yes', 'no')", '', '', ''],

        ['note', 'preassessment_divider', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n📋 PRE-ASSESSMENT COMPLETE  ➔  CLINICAL TEST STARTS HERE\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n📋 ముందస్తు అంచనా పూర్తయింది  ➔  క్లినికల్ పరీక్ష ఇక్కడ ప్రారంభమవుతుంది\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', '', '', consent_yes, '', '', '', '', '', '', ''],

        ['begin group', 'g_instrument_banner', '4. INSTRUMENT SELECTION & START TIME', '4. పరీక్ష ఎంపిక మరియు ప్రారంభ సమయం', '', 'field-list', "${selected_instrument} != ''", '', '', '', '', '', '', ''],
        ['note', 'banner_telugu_a', 'Primary instrument: ICMR-NCTB Telugu MoCA (administer and record; flag as education-confounded). Second instrument: RUDAS-Telugu. Informant: EASI-Telugu Mandatory.', '', '', '', "${selected_instrument} = 'moca_telugu' and ${education_band} = 'A'", '', '', '', '', '', '', ''],
        ['note', 'banner_telugu_b', 'Primary instrument: ICMR-NCTB Telugu MoCA. Informant: EASI-Telugu strongly recommended (mandatory if referral concerns daily functioning).', '', '', '', "${selected_instrument} = 'moca_telugu' and ${education_band} = 'B'", '', '', '', '', '', '', ''],
        ['note', 'banner_telugu_c', 'Primary instrument: ICMR-NCTB Telugu MoCA. Informant: EASI-Telugu if an informant is present.', '', '', '', "${selected_instrument} = 'moca_telugu' and ${education_band} = 'C'", '', '', '', '', '', '', ''],
        ['note', 'banner_moca_b', 'Primary instrument: MoCA-Basic (English Band A: illiterate or <=4 yrs). Informant: EASI Mandatory.', '', '', '', is_mb, '', '', '', '', '', '', ''],
        ['note', 'banner_moca_8_2', 'Primary instrument: MoCA Full, English, Version 8.2 (English Band B or C). Informant: EASI if an informant is present.', '', '', '', is_m8, '', '', '', '', '', '', ''],
        ['note', 'banner_none', 'Neither Telugu nor English selected — no MoCA instrument is opened. Continue to Mandatory History Questions.', '', '', '', "${selected_instrument} = 'none'", '', '', '', '', '', '', ''],
        ['time', 'assessment_start_time', 'Clinical Assessment Start Time (HH:MM)', 'పరీక్ష ప్రారంభ సమయం', 'Record the time when cognitive testing begins.', '', "${selected_instrument} != 'none'", '', '', 'yes', '', '', '', ''],
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
        ['note', 'te_vigilance_inst', "📋 అవధానం (Vigilance) — స్పష్టమైన సూచనలు:\n1. రోగికి చెప్పండి: 'నేను కొన్ని అక్షరాలు చదువుతాను. నేను \"అ\" అనే అక్షరం పలికిన ప్రతిసారీ, మీరు మీ చేతితో బల్లపై ఒకసారి తట్టాలి. నేను వేరే అక్షరం పలికినప్పుడు మీరు తట్టకూడదు.'\n2. పరీక్షకుడు సెకనుకు 1 అక్షరం చొప్పున చదవాలి.\n3. అక్షరాల జాబితా (కేవలం [అ] వచ్చినప్పుడు మాత్రమే తట్టాలి):\n   ప  బ  [అ]  క  మ  న  [అ]  [అ]  జ  క  ల  బ  [అ]  ఫ  [అ]  క  డ  ఈ  [అ]  [అ]  [అ]  జ  [అ]  మా  ఓ  ఫ  [అ]  [అ]  బ\n4. స్కోరింగ్ నిబంధన: 0 లేదా 1 తప్పు అయితే 1 పాయింట్. 2 లేదా అంతకంటే ఎక్కువ తప్పులు ఉంటే 0 పాయింట్లు.", "📋 అవధానం (Vigilance) — స్పష్టమైన సూచనలు:\n1. రోగికి చెప్పండి: 'నేను కొన్ని అక్షరాలు చదువుతాను. నేను \"అ\" అనే అక్షరం పలికిన ప్రతిసారీ, మీరు మీ చేతితో బల్లపై ఒకసారి తట్టాలి. నేను వేరే అక్షరం పలికినప్పుడు మీరు తట్టకూడదు.'\n2. పరీక్షకుడు సెకనుకు 1 అక్షరం చొప్పున చదవాలి.\n3. అక్షరాల జాబితా (కేవలం [అ] వచ్చినప్పుడు మాత్రమే తట్టాలి):\n   ప  బ  [అ]  క  మ  న  [అ]  [అ]  జ  క  ల  బ  [అ]  ఫ  [అ]  క  డ  ఈ  [అ]  [అ]  [అ]  జ  [అ]  మా  ఓ  ఫ  [అ]  [అ]  బ\n4. స్కోరింగ్ నిబంధన: 0 లేదా 1 తప్పు అయితే 1 పాయింట్. 2 లేదా అంతకంటే ఎక్కువ తప్పులు ఉంటే 0 పాయింట్లు.", '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_vigilance', "Vigilance Score ('అ' — 1 pt if 0-1 error, 0 pt if >= 2 errors)", "అక్షరాల జాబితా ('అ' — 0-1 తప్పులకు 1 పాయింట్, >=2 తప్పులకు 0)", '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'te_serial7_inst', "📋 7 తీసివేయడం (Serial 7s) — స్పష్టమైన సూచనలు:\n1. రోగికి చెప్పండి: 'ఇప్పుడు, నేను మిమ్మల్ని 100 నుండి 7 తీసివేయమని అడుగుతాను, ఆపై నేను ఆపమని చెప్పే వరకు మీ సమాధానం నుండి 7 తీసివేస్తూ ఉండండి.'\n2. అవసరమైతే ఈ సూచనను రెండుసార్లు ఇవ్వవచ్చు. రోగికి ఎటువంటి సూచనలు ఇవ్వవద్దు లేదా లెక్కను సరిచేయవద్దు.\n3. తీసివేత దశలు:\n   [1] 100 - 7 = 93\n   [2] 93 - 7 = 86\n   [3] 86 - 7 = 79\n   [4] 79 - 7 = 72\n   [5] 72 - 7 = 65\n4. స్వతంత్ర తీసివేత నియమం: ప్రతి తీసివేత స్వతంత్రంగా లెక్కించబడుతుంది.\n5. స్కోరింగ్: 4-5 సరైనవి = 3 పాయింట్లు | 2-3 సరైనవి = 2 పాయింట్లు | 1 సరైనది = 1 పాయింట్ | 0 సరైనవి = 0 పాయింట్లు.", "📋 7 తీసివేయడం (Serial 7s) — స్పష్టమైన సూచనలు:\n1. రోగికి చెప్పండి: 'ఇప్పుడు, నేను మిమ్మల్ని 100 నుండి 7 తీసివేయమని అడుగుతాను, ఆపై నేను ఆపమని చెప్పే వరకు మీ సమాధానం నుండి 7 తీసివేస్తూ ఉండండి.'\n2. అవసరమైతే ఈ సూచనను రెండుసార్లు ఇవ్వవచ్చు. రోగికి ఎటువంటి సూచనలు ఇవ్వవద్దు లేదా లెక్కను సరిచేయవద్దు.\n3. తీసివేత దశలు:\n   [1] 100 - 7 = 93\n   [2] 93 - 7 = 86\n   [3] 86 - 7 = 79\n   [4] 79 - 7 = 72\n   [5] 72 - 7 = 65\n4. స్వతంత్ర తీసివేత నియమం: ప్రతి తీసివేత స్వతంత్రంగా లెక్కించబడుతుంది.\n5. స్కోరింగ్: 4-5 సరైనవి = 3 పాయింట్లు | 2-3 సరైనవి = 2 పాయింట్లు | 1 సరైనది = 1 పాయింట్ | 0 సరైనవి = 0 పాయింట్లు.", '', '', '', '', '', '', '', '', '', ''],
        ['select_one serial7_score', 'te_serial_7s', 'Serial 7s (100 నుండి క్రమంగా 7 ను తీసివేయండి)', '100 నుండి 7 తీసివేయడం', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'te_g_language', 'ICMR-NCTB Telugu MoCA — LANGUAGE', 'భాష (LANGUAGE)', '', 'field-list', is_te, '', '', '', '', '', '', ''],
        ['select_one yesno', 'te_sentence_rep_1', 'తిరిగి చెప్పండి: రాజు ఒక్కడే ఈ రోజు సహాయం చేస్తాడని నేను అనుకుంటున్నాను.', 'తిరిగి చెప్పండి: రాజు ఒక్కడే ఈ రోజు సహాయం చేస్తాడని నేను అనుకుంటున్నాను.', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'te_sentence_rep_2', 'కుక్కను గదిలో చూసినప్పుడల్లా పిల్లి మంచం కింద దాక్కుంటుంది.', 'కుక్కను గదిలో చూసినప్పుడల్లా పిల్లి మంచం కింద దాక్కుంటుంది.', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'te_fluency_timer_guidance', '⏱️ 1-నిమిషం టైమర్: మీ స్టాప్‌వాచ్ లేదా మొబైల్ టైమర్‌ను ప్రారంభించండి. రోగి ఒక నిమిషంలో వీలైనన్ని పదాలు చెప్పాలి. సరిగ్గా 60 సెకన్ల తర్వాత ఆపండి.', '⏱️ 1-నిమిషం టైమర్: మీ స్టాప్‌వాచ్ లేదా మొబైల్ టైమర్‌ను ప్రారంభించండి. రోగి ఒక నిమిషంలో వీలైనన్ని పదాలు చెప్పాలి. సరిగ్గా 60 సెకన్ల తర్వాత ఆపండి.', '', '', '', '', '', '', '', '', '', ''],
        ['integer', 'te_fluency_count', "ఒక నిమిషంలో 'ప' అక్షరంతో మొదలయ్యే పదాలు ఎన్ని ఎక్కువ చెప్పగలిగితే అన్ని చెప్పండి (N >= 11 words = 1 pt)", "ఒక నిమిషంలో 'ప' అక్షరంతో మొదలయ్యే పదాలు (N >= 11 words = 1 pt)", '', '', '', '. >= 0 and . <= 100', 'Enter a count between 0 and 100.', 'yes', '', '', '', ''],
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
        # BRANCH 4C: MoCA FULL, ENGLISH, VERSION 8.2 (exact page 12 wording)
        # ============================================================
        ['note', 'm8_top_header', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n# MONTREAL COGNITIVE ASSESSMENT (MoCA 8.2)\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', '', '', '', is_m8, '', '', '', '', '', '', ''],
        ['begin group', 'm8_g_visuospatial', 'VISUOSPATIAL / EXECUTIVE', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_trail_inst', 'Alternating Trail Making (Begin -> End)', '', '', '', '', '', '', '', '', '', '', 'trail_making_8_2.png'],
        ['select_one yesno', 'm8_trail_making', 'Alternating Trail Making', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'm8_chair_inst', 'Copy chair', '', '', '', '', '', '', '', '', '', '', 'chair_8_2.png'],
        ['select_one yesno', 'm8_chair_copy', 'Copy chair', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'm8_clock_inst', 'Draw CLOCK ( Ten past nine ) ( 3 points )', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_clock_contour', 'Contour', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_clock_numbers', 'Numbers', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_clock_hands', 'Hands', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_naming', 'NAMING', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_naming_inst', 'NAMING', '', '', '', '', '', '', '', '', '', '', 'naming_8_2.png'],
        ['select_one yesno', 'm8_naming_1', '1. Snake (or a type of snake like Boa or Cobra)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_naming_2', '2. Elephant', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_naming_3', '3. Crocodile or Alligator', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_memory', 'MEMORY (NO POINTS)', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_memory_inst', 'Read list of words, subject must repeat them. Do 2 trials, even if 1st trial is successful. Do a recall after 5 minutes.', '', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'm8_memory_words', 'HAND | NYLON | PARK | CARROT | YELLOW', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'm8_mem_trial1', '1ST TRIAL', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_hand', 'HAND (1ST TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_nylon', 'NYLON (1ST TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_park', 'PARK (1ST TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_carrot', 'CARROT (1ST TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t1_yellow', 'YELLOW (1ST TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'm8_mem_trial2', '2ND TRIAL', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_hand', 'HAND (2ND TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_nylon', 'NYLON (2ND TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_park', 'PARK (2ND TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_carrot', 'CARROT (2ND TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_mem_t2_yellow', 'YELLOW (2ND TRIAL)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_attention', 'ATTENTION', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_digits_inst', 'Read list of digits ( 1 digit/ sec. ).', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_digit_span_fwd', 'Subject has to repeat them in the forward order. 8 1 5 2 4', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_digit_span_bwd', 'Subject has to repeat them in the backward order. 2 4 7', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'm8_vigilance_inst', '📋 VIGILANCE — CLEAR INSTRUCTIONS:\n1. Instruct participant: "I am going to read a sequence of letters. Every time I say the letter A, tap your hand once on the table. If I say a different letter, do NOT tap your hand."\n2. Assessor guideline: Read letters at exactly 1 letter per second.\n3. Sequence (tap ONLY on [A]):\n   F  B  [A]  C  M  N  [A]  [A]  J  K  L  B  [A]  F  [A]  K  D  E  [A]  [A]  [A]  J  [A]  M  O  F  [A]  [A]  B\n4. Scoring: 1 point if 0 or 1 error (tapping on non-A letter or failing to tap on A). 0 points if >= 2 errors.', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_vigilance', 'Vigilance Score (1 pt if 0-1 error, 0 pt if >= 2 errors)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'm8_serial7_inst', '📋 SERIAL 7 SUBTRACTION — CLEAR INSTRUCTIONS:\n1. Instruct participant: "Now, I will ask you to count by subtracting seven from seventy, and keep subtracting seven from your answer until I tell you to stop."\n2. Give instruction twice if needed. Do NOT interrupt, give hints, or correct math.\n3. Correct Subtraction Steps:\n   [1] 70 - 7 = 63\n   [2] 63 - 7 = 56\n   [3] 56 - 7 = 49\n   [4] 49 - 7 = 42\n   [5] 42 - 7 = 35\n4. Independent Subtraction Rule: Each subtraction is scored independently. (e.g. if subject says "62, 55, 48, 41, 34" — 62 is wrong, but 55, 48, 41, 34 are all correct subtractions from the previous number, so award 4 correct = 3 points).\n5. Score criteria: 4-5 correct = 3 pts | 2-3 correct = 2 pts | 1 correct = 1 pt | 0 correct = 0 pts.', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one serial7_score', 'm8_serial_7s', 'Serial 7s Score (subtraction starting at 70)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_language', 'LANGUAGE', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_sentence_rep_1', 'Repeat: The robber of the gray car was stopped by the police.', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_sentence_rep_2', 'The student went back to school without his books and pencils.', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['note', 'm8_fluency_timer_guidance', '⏱️ 1-MINUTE TIMER GUIDANCE:\n• Instruct: "Tell me as many words you can think of that begin with the letter S in one minute."\n• Start your stopwatch as soon as you finish instructions.\n• Stop the participant immediately after exactly 60 seconds (1 minute).', '', '', '', '', '', '', '', '', '', '', ''],
        ['integer', 'm8_fluency_count', 'Language Fluency (N words in 1 minute, N >= 11 words = 1 pt)', '', 'Do not count proper nouns (like Sam, Seattle), numbers, or words with different endings of the same word (e.g. swim, swimming).', '', '', '. >= 0 and . <= 100', 'Enter a count between 0 and 100.', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_abstraction', 'ABSTRACTION', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_abstraction_inst', 'Similarity between e.g. banana - orange = fruit', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'm8_abstraction_1', 'bed - table', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'm8_abstraction_2', 'letter - telephone', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_delayed_recall', 'DELAYED RECALL', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
        ['note', 'm8_delayed_inst', 'Has to recall words WITH NO CUE. Points for UNCUED recall only.', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_hand', 'HAND', 'Category cue: body part | Multiple choice: hand, shoulder, face (nose, leg)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_nylon', 'NYLON', 'Category cue: type of fabric | Multiple choice: cotton, nylon, silk (velvet, denim)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_park', 'PARK', 'Category cue: public place | Multiple choice: park, library, school (church, hospital)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_carrot', 'CARROT', 'Category cue: type of food | Multiple choice: tomato, carrot, lettuce (cucumber, celery)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one m8_recall_type', 'm8_delayed_yellow', 'YELLOW', 'Category cue: color | Multiple choice: yellow, purple, green (red, blue)', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        ['begin group', 'm8_g_orientation', 'ORIENTATION', '', '', 'field-list', is_m8, '', '', '', '', '', '', ''],
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
        ['select_one patient_informant', 'mh_asked_to', 'Questions are asked to', 'ప్రశ్నలు ఎవరిని అడిగారు', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
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
        ['select_one patient_informant', 'mh_answered_by', 'Answered by', 'సమాధానం ఇచ్చినవారు', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
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

        # MoCA 8.2
        ['calculate', 'm8_visuospatial_total', '', '', '', '', '', '', '', '', 'if(${m8_trail_making} = "yes", 1, 0) + if(${m8_chair_copy} = "yes", 1, 0) + if(${m8_clock_contour} = "yes", 1, 0) + if(${m8_clock_numbers} = "yes", 1, 0) + if(${m8_clock_hands} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'm8_naming_total', '', '', '', '', '', '', '', '', 'if(${m8_naming_1} = "yes", 1, 0) + if(${m8_naming_2} = "yes", 1, 0) + if(${m8_naming_3} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'm8_attention_total', '', '', '', '', '', '', '', '', 'if(${m8_digit_span_fwd} = "yes", 1, 0) + if(${m8_digit_span_bwd} = "yes", 1, 0) + if(${m8_vigilance} = "yes", 1, 0) + if(${m8_serial_7s} != "", number(${m8_serial_7s}), 0)', '', '', ''],
        ['calculate', 'm8_language_total', '', '', '', '', '', '', '', '', 'if(${m8_sentence_rep_1} = "yes", 1, 0) + if(${m8_sentence_rep_2} = "yes", 1, 0) + if(${m8_fluency_count} != "" and ${m8_fluency_count} >= 11, 1, 0)', '', '', ''],
        ['calculate', 'm8_abstraction_total', '', '', '', '', '', '', '', '', 'if(${m8_abstraction_1} = "yes", 1, 0) + if(${m8_abstraction_2} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'm8_memory_total', '', '', '', '', '', '', '', '', 'if(${m8_delayed_hand} = "uncued", 1, 0) + if(${m8_delayed_nylon} = "uncued", 1, 0) + if(${m8_delayed_park} = "uncued", 1, 0) + if(${m8_delayed_carrot} = "uncued", 1, 0) + if(${m8_delayed_yellow} = "uncued", 1, 0)', '', '', ''],
        ['calculate', 'm8_mis_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_8_2", (if(${m8_delayed_hand} = "uncued", 3, if(${m8_delayed_hand} = "category", 2, if(${m8_delayed_hand} = "multiple_choice", 1, 0))) + if(${m8_delayed_nylon} = "uncued", 3, if(${m8_delayed_nylon} = "category", 2, if(${m8_delayed_nylon} = "multiple_choice", 1, 0))) + if(${m8_delayed_park} = "uncued", 3, if(${m8_delayed_park} = "category", 2, if(${m8_delayed_park} = "multiple_choice", 1, 0))) + if(${m8_delayed_carrot} = "uncued", 3, if(${m8_delayed_carrot} = "category", 2, if(${m8_delayed_carrot} = "multiple_choice", 1, 0))) + if(${m8_delayed_yellow} = "uncued", 3, if(${m8_delayed_yellow} = "category", 2, if(${m8_delayed_yellow} = "multiple_choice", 1, 0)))), "")', '', '', ''],
        ['calculate', 'm8_orientation_total', '', '', '', '', '', '', '', '', 'if(${m8_orientation_date} = "yes", 1, 0) + if(${m8_orientation_month} = "yes", 1, 0) + if(${m8_orientation_year} = "yes", 1, 0) + if(${m8_orientation_day} = "yes", 1, 0) + if(${m8_orientation_place} = "yes", 1, 0) + if(${m8_orientation_city} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'm8_raw_total', '', '', '', '', '', '', '', '', '${m8_visuospatial_total} + ${m8_naming_total} + ${m8_attention_total} + ${m8_language_total} + ${m8_abstraction_total} + ${m8_memory_total} + ${m8_orientation_total}', '', '', ''],
        ['calculate', 'm8_education_adj', '', '', '', '', '', '', '', '', 'if(${years_schooling} != "" and number(${years_schooling}) <= 12, 1, 0)', '', '', ''],
        ['calculate', 'm8_total_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_8_2", min(${m8_raw_total} + ${m8_education_adj}, 30), "")', '', '', ''],

        # Unified totals
        ['calculate', 'moca_total_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_telugu", ${te_total_score}, if(${selected_instrument} = "moca_b", ${mb_total_score}, if(${selected_instrument} = "moca_8_2", ${m8_total_score}, "")))', '', '', ''],
        ['calculate', 'moca_mis_score', '', '', '', '', '', '', '', '', 'if(${selected_instrument} = "moca_8_2", ${m8_mis_score}, "")', '', '', ''],
        ['calculate', 'moca_interpretation', '', '', '', '', '', '', '', '', 'if(${consent_reconfirmed} != "yes", "Not assessed (no consent)", if(${selected_instrument} = "moca_telugu" or ${selected_instrument} = "moca_8_2", if(${moca_total_score} >= 26, "Normal (>= 26/30)", "Below normal (< 26/30)"), ""))', '', '', ''],

        # Admin / Summary
        ['begin group', 'g_admin', 'TOTAL & SIGN-OFF', 'మొత్తం స్కోరు', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['string', 'administered_by', 'Administered by', 'పరీక్ష నిర్వహించినవారు', '', '', '', '', '', 'yes', '', '', '', ''],
        ['dateTime', 'assessment_datetime', 'Date of exam', 'తేదీ', '', '', '', '', '', 'yes', "if(. = '', now(), .)", '', '', ''],
        ['time', 'assessment_end_time', 'Clinical Assessment End Time (HH:MM)', 'పరీక్ష ముగింపు సమయం', 'Record the time when cognitive testing finishes.', '', "${selected_instrument} != 'none'", '', '', 'yes', '', '', '', ''],
        ['note', 'score_summary_note', 'Instrument: ${selected_instrument} | TOTAL SCORE: ${moca_total_score} / 30', '', '', '', "${selected_instrument} = 'moca_telugu' or ${selected_instrument} = 'moca_b' or ${selected_instrument} = 'moca_8_2'", '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # ============================================================
        # SUPPORTING DOCUMENTS / ASSESSMENT SHEET UPLOAD
        # ============================================================
        ['begin group', 'g_supporting_docs', 'Supporting Documents / Scale Upload', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['image', 'upload_scale_image', 'Upload completed scale / assessment sheet (image)', '', 'Upload clear photos of the completed paper scale, consent form, or supporting assessment document. Accepts JPG, JPEG, PNG. For PDF use next field.', '', '', '', '', 'no', '', '', '', ''],
        ['file', 'upload_scale_pdf', 'Upload completed scale / assessment sheet (PDF)', '', 'Upload a PDF of the completed paper scale if not already uploaded as image. Accepts PDF.', '', '', '', '', 'no', '', '', '', ''],
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
        ['patient_informant', 'patient', 'Patient', 'రోగి'],
        ['patient_informant', 'informant', 'Informant', 'సమాచారకర్త'],
        ['spoken_lang', 'telugu', 'Telugu', 'తెలుగు'],
        ['spoken_lang', 'english', 'English', 'ఇంగ్లీష్'],
        ['spoken_lang', 'hindi', 'Hindi', 'హిందీ'],
        ['spoken_lang', 'lambada', 'Lambada', 'లంబాడా'],
        ['spoken_lang', 'marathi', 'Marathi', 'మరాఠీ'],
        ['spoken_lang', 'kannada', 'Kannada', 'కన్నడ'],
        ['spoken_lang', 'tamil', 'Tamil', 'తమిళం'],
        ['spoken_lang', 'odia', 'Odia', 'ఒడియా'],
        ['spoken_lang', 'other', 'Other', 'ఇతర భాష'],
        ['luq_lang', 'telugu', 'Telugu', 'తెలుగు'],
        ['luq_lang', 'english', 'English', 'ఇంగ్లీష్'],
        ['luq_lang', 'neither', 'Neither', 'రెండూ కాదు'],
        ['test_lang', 'telugu', 'Telugu', 'తెలుగు'],
        ['test_lang', 'english', 'English', 'ఇంగ్లీష్'],
        ['test_lang', 'neither', 'Neither / Other', 'ఇతర భాష'],
        ['contact_freq', 'daily', 'Daily', 'రోజువారీ'],
        ['contact_freq', 'weekly', 'Weekly', 'వారానికొకసారి'],
        ['contact_freq', 'less_often', 'Less often', 'తక్కువసార్లు'],
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
        ['m8_recall_type', 'uncued', 'Has to recall words WITH NO CUE (MIS X3)', 'సూచన లేకుండా గుర్తుకు వచ్చింది'],
        ['m8_recall_type', 'category', 'Category cue (MIS X2)', 'వర్గ సూచనతో గుర్తుకు వచ్చింది'],
        ['m8_recall_type', 'multiple_choice', 'Multiple choice cue (MIS X1)', 'బహుళ ఎంపిక సూచనతో గుర్తుకు వచ్చింది'],
        ['m8_recall_type', 'not_recalled', 'Not recalled', 'గుర్తుకు రాలేదు'],
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
