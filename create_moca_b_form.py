"""create_moca_b_form.py - Generate the MoCA-B (Basic) XLSForm and properties.

Source of truth (exact text & scoring only):
  scales/moca_b/moca-b-scoring-sheet-june-2014.pdf
  scales/moca_b/moca-b-instructions-june-2014.pdf
"""

import json
import os
import openpyxl

REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
FORMS_APP_DIR = os.path.join(REPO_ROOT, 'forms', 'app')


def build_moca_b_xlsx():
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
        ['begin group', 'inputs', 'Participant', '', '', 'field-list', "./source = 'user'", '', '', '', '', '', '', ''],
        ['hidden', 'source', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'source_id', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'user', 'User', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['hidden', 'contact_id', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['hidden', 'name', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'contact', 'Participant', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['db:person', '_id', 'Select participant', '', '', 'select-contact type-participant', '', '', '', '', '', '', '', ''],
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

        # Education, Illiteracy & Start Time
        ['begin group', 'g_consent', 'Education & Timing', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'consent_reconfirmed', 'Consent re-confirmed?', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['integer', 'education_years', 'Education (years)', '', 'Add 1 point if education < 4 year', '', consent_yes, '. >= 0 and . <= 30', 'Enter years between 0 and 30.', 'yes', '', '', '', ''],
        ['select_one yesno', 'is_illiterate', 'Illiterate', '', 'AND add 1 point if illiterate (inability to read or write fluently in daily living)', '', consent_yes, '', '', 'yes', '', '', '', ''],
        ['time', 'start_time', 'START TIME', '', '', '', consent_yes, '', '', 'no', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 1. Executive Function (1 pt)
        ['begin group', 'g_executive', 'EXECUTIVE FUNCTION', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'trail_making_inst', 'Please draw a line alternating between a square with a number and a square with dots in increasing order. Begin here [1] and draw a line to [1 dot], then [2], then [2 dots], and so on. End here [6 dots]. 1 point if completed on first attempt without any errors or self-corrections.', '', '', '', '', '', '', '', '', '', '', 'trail_making.png'],
        ['select_one yesno', 'trail_making', 'Alternating Trail Making (SCORE /1)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 2. Immediate Recall (No point - 2 trials)
        ['begin group', 'g_memory', 'IMMEDIATE RECALL (No point)', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'memory_inst_trial1', 'Perform 2 trials even if 1st trial is successful. No point.', '', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'memory_words', 'ROSE | CHAIR | HAND | BLUE | SPOON', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'mem_trial1', '1st trial', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mem_t1_rose', 'ROSE (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_chair', 'CHAIR (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_hand', 'HAND (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_blue', 'BLUE (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_spoon', 'SPOON (1st trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'memory_inst_trial2', '2nd trial', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'mem_trial2', '2nd trial', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mem_t2_rose', 'ROSE (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_chair', 'CHAIR (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_hand', 'HAND (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_blue', 'BLUE (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_spoon', 'SPOON (2nd trial)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 3. Fluency (2 pts)
        ['begin group', 'g_fluency', 'FLUENCY', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'fluency_inst', 'Name maximum numbers of FRUITS in 1 minute. 2 points if N = 13 or more; 1 point if N = 8-12; 0 point if N = 7 or less.', '', '', '', '', '', '', '', '', '', '', ''],
        ['text', 'fluency_words', 'Fruits named (1-18)', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['integer', 'fluency_count', 'N items', '', '', '', '', '. >= 0 and . <= 100', 'Enter a count between 0 and 100.', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 4. Orientation (6 pts)
        ['begin group', 'g_orientation', 'ORIENTATION', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'orientation_inst', '1 point is allocated for each correct answer.', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'orientation_time', 'time (± 2 hr)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_day', 'day', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_month', 'month', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_year', 'year', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_place', 'place', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_city', 'city', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 5. Calculation (3 pts)
        ['begin group', 'g_calculation', 'CALCULATION', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'calculation_inst', 'Provide 3 ways to pay using 1 dollar coins, 5 dollar and 10 dollar bills for an object that costs exactly "13 Dollars". (3 points if 3 ways, 2 points if 2 ways, 1 point if 1 way, 0 point if no correct way)', '', '', '', '', '', '', '', '', '', '', ''],
        ['text', 'calc_combinations', '1. / 2. / 3.', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['select_one calc_score', 'calculation_score_select', 'SCORE ( /3)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 6. Abstraction (3 pts)
        ['begin group', 'g_abstraction', 'ABSTRACTION', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'abstraction_inst_practice', 'To what category these objects belong to? (e.g. orange - banana = fruit)', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'abstraction_1', 'train - boat', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'abstraction_2', 'north - south', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'abstraction_3', 'drum - flute', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 7. Delayed Recall (5 pts - free recall only; cues 0 pts)
        ['begin group', 'g_delayed_recall', 'DELAYED RECALL', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'delayed_inst', 'Points are awarded for recall with No cue (1 point for each item).', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'delayed_rose', 'ROSE - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_rose_cat', 'ROSE - Recall with category cue (type of flower)', '', '', '', "${delayed_rose} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_rose_mc', 'ROSE - Recall with multiple choice cue (rose / daisy / tulip)', '', '', '', "${delayed_rose} = 'no' and ${delayed_rose_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_chair', 'CHAIR - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_chair_cat', 'CHAIR - Recall with category cue (type of furniture)', '', '', '', "${delayed_chair} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_chair_mc', 'CHAIR - Recall with multiple choice cue (table / chair / bed)', '', '', '', "${delayed_chair} = 'no' and ${delayed_chair_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_hand', 'HAND - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_hand_cat', 'HAND - Recall with category cue (body part)', '', '', '', "${delayed_hand} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_hand_mc', 'HAND - Recall with multiple choice cue (foot / hand / knee)', '', '', '', "${delayed_hand} = 'no' and ${delayed_hand_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_blue', 'BLUE - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_blue_cat', 'BLUE - Recall with category cue (colour)', '', '', '', "${delayed_blue} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_blue_mc', 'BLUE - Recall with multiple choice cue (blue / brown / red)', '', '', '', "${delayed_blue} = 'no' and ${delayed_blue_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_spoon', 'SPOON - Recall with No cue', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_spoon_cat', 'SPOON - Recall with category cue (kitchen instrument)', '', '', '', "${delayed_spoon} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_spoon_mc', 'SPOON - Recall with multiple choice cue (fork / spoon / knife)', '', '', '', "${delayed_spoon} = 'no' and ${delayed_spoon_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 8. Visuoperception (3 pts)
        ['begin group', 'g_visuoperception', 'VISUOPERCEPTION', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'visuoperception_inst', 'Identify drawings. No more than 60 seconds. See complementary sheet. 3 points if N = 9-10; 2 points if N = 6-8; 1 point if N = 4-5; 0 point if N = 0-3.', '', '', '', '', '', '', '', '', '', '', 'visuoperception.png'],
        ['select_one yesno', 'vp_scissors', 'scissors', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_tshirt', 'T-shirt', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_banana', 'banana', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_lamp', 'lamp', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_candle', 'candle', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_watch', 'watch', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_cup', 'cup', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_leaf', 'leaf', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_key', 'key', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_spoon', 'spoon', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 9. Naming (4 pts)
        ['begin group', 'g_naming', 'NAMING', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'naming_inst', 'Identify animals. See complementary sheet.', '', '', '', '', '', '', '', '', '', '', 'naming.png'],
        ['select_one yesno', 'naming_zebra', 'zebra', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'naming_peacock', 'peacock', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'naming_tiger', 'tiger', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'naming_butterfly', 'butterfly', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 10. Attention (3 pts: 1 + 2)
        ['begin group', 'g_attention', 'ATTENTION', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'attention_white_inst', 'Name the numbers in circles. See complementary sheet. No point if 2 errors or more.', '', '', '', '', '', '', '', '', '', '', 'attention_white.png'],
        ['integer', 'attention_white_errors', 'ERROR N (White background)', '', '', '', '', '. >= 0 and . <= 50', 'Enter errors between 0 and 50.', 'yes', '', '', '', ''],
        ['note', 'attention_black_inst', 'Name the numbers in circles & squares. See complementary sheet. 2 points if 2 errors or less; 1 point if 3 errors; 0 point if 4 errors or more.', '', '', '', '', '', '', '', '', '', '', 'attention_black.png'],
        ['integer', 'attention_black_errors', 'ERROR N (Black background)', '', '', '', '', '. >= 0 and . <= 50', 'Enter errors between 0 and 50.', 'yes', '', '', '', ''],
        ['time', 'end_time', 'END TIME', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # Scoring calculations (exact PDF scoring rules, no invented interpretation)
        ['calculate', 'mocab_executive_total', '', '', '', '', '', '', '', '', 'if(${trail_making} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mocab_fluency_total', '', '', '', '', '', '', '', '', 'if(${fluency_count} != "" and ${fluency_count} >= 13, 2, if(${fluency_count} != "" and ${fluency_count} >= 8, 1, 0))', '', '', ''],
        ['calculate', 'mocab_orientation_total', '', '', '', '', '', '', '', '', 'if(${orientation_time} = "yes", 1, 0) + if(${orientation_day} = "yes", 1, 0) + if(${orientation_month} = "yes", 1, 0) + if(${orientation_year} = "yes", 1, 0) + if(${orientation_place} = "yes", 1, 0) + if(${orientation_city} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mocab_calculation_total', '', '', '', '', '', '', '', '', 'if(${calculation_score_select} != "", number(${calculation_score_select}), 0)', '', '', ''],
        ['calculate', 'mocab_abstraction_total', '', '', '', '', '', '', '', '', 'if(${abstraction_1} = "yes", 1, 0) + if(${abstraction_2} = "yes", 1, 0) + if(${abstraction_3} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mocab_memory_total', '', '', '', '', '', '', '', '', 'if(${delayed_rose} = "yes", 1, 0) + if(${delayed_chair} = "yes", 1, 0) + if(${delayed_hand} = "yes", 1, 0) + if(${delayed_blue} = "yes", 1, 0) + if(${delayed_spoon} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'visuoperception_count', '', '', '', '', '', '', '', '', 'if(${vp_scissors} = "yes", 1, 0) + if(${vp_tshirt} = "yes", 1, 0) + if(${vp_banana} = "yes", 1, 0) + if(${vp_lamp} = "yes", 1, 0) + if(${vp_candle} = "yes", 1, 0) + if(${vp_watch} = "yes", 1, 0) + if(${vp_cup} = "yes", 1, 0) + if(${vp_leaf} = "yes", 1, 0) + if(${vp_key} = "yes", 1, 0) + if(${vp_spoon} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mocab_visuoperception_total', '', '', '', '', '', '', '', '', 'if(${visuoperception_count} >= 9, 3, if(${visuoperception_count} >= 6, 2, if(${visuoperception_count} >= 4, 1, 0)))', '', '', ''],
        ['calculate', 'mocab_naming_total', '', '', '', '', '', '', '', '', 'if(${naming_zebra} = "yes", 1, 0) + if(${naming_peacock} = "yes", 1, 0) + if(${naming_tiger} = "yes", 1, 0) + if(${naming_butterfly} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mocab_attention_white_score', '', '', '', '', '', '', '', '', 'if(${attention_white_errors} != "" and ${attention_white_errors} <= 1, 1, 0)', '', '', ''],
        ['calculate', 'mocab_attention_black_score', '', '', '', '', '', '', '', '', 'if(${attention_black_errors} != "" and ${attention_black_errors} <= 2, 2, if(${attention_black_errors} = 3, 1, 0))', '', '', ''],
        ['calculate', 'mocab_attention_total', '', '', '', '', '', '', '', '', '${mocab_attention_white_score} + ${mocab_attention_black_score}', '', '', ''],
        ['calculate', 'mocab_raw_total', '', '', '', '', '', '', '', '', '${mocab_executive_total} + ${mocab_fluency_total} + ${mocab_orientation_total} + ${mocab_calculation_total} + ${mocab_abstraction_total} + ${mocab_memory_total} + ${mocab_visuoperception_total} + ${mocab_naming_total} + ${mocab_attention_total}', '', '', ''],
        ['calculate', 'mocab_education_adj', '', '', '', '', '', '', '', '', 'if(${education_years} != "" and ${education_years} < 4, 1, 0)', '', '', ''],
        ['calculate', 'mocab_literacy_adj', '', '', '', '', '', '', '', '', 'if(${is_illiterate} = "yes", 1, 0)', '', '', ''],
        ['calculate', 'mocab_total_score', '', '', '', '', '', '', '', '', 'if(${consent_reconfirmed} = "yes", min(${mocab_raw_total} + ${mocab_education_adj} + ${mocab_literacy_adj}, 30), "")', '', '', ''],

        # Total Time, Administered by & Date of exam (from PDF)
        ['begin group', 'g_admin', 'TOTAL SCORE & TIME', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['integer', 'total_time_min', 'TOTAL TIME (min)', '', '', '', consent_yes, '. >= 0 and . <= 180', 'Enter minutes between 0 and 180.', 'no', '', '', '', ''],
        ['integer', 'total_time_sec', 'TOTAL TIME (sec)', '', '', '', consent_yes, '. >= 0 and . <= 59', 'Enter seconds between 0 and 59.', 'no', '', '', '', ''],
        ['string', 'administered_by', 'Administered by', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['dateTime', 'assessment_datetime', 'Date of exam', '', '', '', '', '', '', 'yes', "if(. = '', now(), .)", '', '', ''],
        ['note', 'score_summary_note', 'TOTAL SCORE: ${mocab_total_score} / 30 (Add 1 point if education < 4 year AND add 1 point if illiterate)', '', '', '', consent_yes, '', '', '', '', '', '', ''],
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
        ['calc_score', '0', '0 point if no correct way', '0 పాయింట్లు'],
        ['calc_score', '1', '1 point if 1 way', '1 పాయింట్'],
        ['calc_score', '2', '2 points if 2 ways', '2 పాయింట్లు'],
        ['calc_score', '3', '3 points if 3 ways', '3 పాయింట్లు'],
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
    ws_settings.append(['MoCA-B (Basic)', 'moca_b', '2026-11-12 02-00', 'en'])

    os.makedirs(FORMS_APP_DIR, exist_ok=True)
    xlsx_path = os.path.join(FORMS_APP_DIR, 'moca_b.xlsx')
    wb.save(xlsx_path)
    print(f'Saved: {xlsx_path}')

    props_path = os.path.join(FORMS_APP_DIR, 'moca_b.properties.json')
    props = {
        'title': [
            {
                'locale': 'en',
                'content': 'Montreal Cognitive Assessment - Basic (MoCA-B)',
            }
        ],
        'icon': 'icon-healthcare-assessment',
        'context': {
            'person': True,
            'place': False,
            'expression': '!contact || summary.is_participant',
        },
    }
    with open(props_path, 'w', encoding='utf-8', newline='\n') as f:
        json.dump(props, f, indent=2)
        f.write('\n')
    print(f'Saved: {props_path}')


if __name__ == '__main__':
    build_moca_b_xlsx()
