"""create_moca_b_form.py - Generate the MoCA-B (Basic) XLSForm and properties.

Source of truth:
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
        'calculation', 'default', 'read_only', 'media::image::en', 'parameters',
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

        # Consent, Education & Literacy
        ['begin group', 'g_consent', 'Consent & Participant Education', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'consent_reconfirmed', 'Consent re-confirmed for the MoCA-B assessment?', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['integer', 'education_years', 'Years of formal education', '', 'Registration says: "${patient_education}". Enter completed years of schooling (0-30). +1 point is added if < 4 years (capped at 30).', '', consent_yes, '. >= 0 and . <= 30', 'Enter years between 0 and 30.', 'yes', '', '', '', ''],
        ['select_one yesno', 'is_illiterate', 'Is the participant considered illiterate?', '', 'Inability to read or write fluently in daily living. +1 point is added regardless of years of education (capped at 30).', '', consent_yes, '', '', 'yes', '', '', '', ''],
        ['time', 'start_time', 'Start time (record before administering Executive Function)', '', 'Optional: record start time (hh:mm).', '', consent_yes, '', '', 'no', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 1. Executive Function (1 pt)
        ['begin group', 'g_executive', 'Executive Function (Alternating Trail Making)', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'trail_making_inst', 'Ask the subject to draw a line alternating between a square with a number and a square with dots in increasing order: 1 -> (1 dot) -> 2 -> (2 dots) -> 3 -> (3 dots) -> 4 -> (4 dots) -> 5 -> (5 dots) -> 6 -> (6 dots). 1 point only if completed on first attempt without any errors or self-corrections.', '', '', '', '', '', '', '', '', '', '', 'trail_making.png'],
        ['select_one yesno', 'trail_making', 'Trail Making (1 -> 1 dot -> ... -> 6 -> 6 dots, zero errors/self-corrections)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 2. Immediate Recall (0 pts - 2 trials)
        ['begin group', 'g_memory', 'Immediate Recall (No point)', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'memory_inst_trial1', 'Read the 5 words at 1 word/second. Perform 2 trials even if the 1st trial is successful. No points are scored for Trials 1 and 2.', '', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'memory_words', 'Words: ROSE, CHAIR, HAND, BLUE, SPOON', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'mem_trial1', 'Trial 1', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mem_t1_rose', 'ROSE (Trial 1)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_chair', 'CHAIR (Trial 1)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_hand', 'HAND (Trial 1)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_blue', 'BLUE (Trial 1)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t1_spoon', 'SPOON (Trial 1)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['note', 'memory_inst_trial2', 'Read the same 5 words a second time. Inform the subject that they will be asked to recall these words again at the end of the test.', '', '', '', '', '', '', '', '', '', '', ''],
        ['begin group', 'mem_trial2', 'Trial 2', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'mem_t2_rose', 'ROSE (Trial 2)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_chair', 'CHAIR (Trial 2)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_hand', 'HAND (Trial 2)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_blue', 'BLUE (Trial 2)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'mem_t2_spoon', 'SPOON (Trial 2)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 3. Fluency (2 pts)
        ['begin group', 'g_fluency', 'Fluency', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'fluency_inst', 'Ask the subject to name as many FRUITS as possible in 1 minute (60 seconds). Record all words to ensure repeated words are not counted. Scoring: >=13 fruits = 2 pts; 8-12 fruits = 1 pt; <=7 fruits = 0 pts.', '', '', '', '', '', '', '', '', '', '', ''],
        ['text', 'fluency_words', 'Recorded fruit names (optional)', '', 'List named fruits to track repeats.', '', '', '', '', 'no', '', '', '', ''],
        ['integer', 'fluency_count', 'Total unique FRUITS named in 1 minute (N)', '', '0-7 = 0 pts, 8-12 = 1 pt, >=13 = 2 pts.', '', '', '. >= 0 and . <= 100', 'Enter a count between 0 and 100.', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 4. Orientation (6 pts)
        ['begin group', 'g_orientation', 'Orientation', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'orientation_inst', 'Without looking at a watch, ask approximate time (within +/- 2 hours accepted). Then ask day of the week, month, year, name of this place (hospital/clinic/office), and city.', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'orientation_time', 'Time (+/- 2 hr) correct?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_day', 'Day of week correct?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_month', 'Month correct?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_year', 'Year correct?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_place', 'Place (hospital/clinic/office) correct?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'orientation_city', 'City correct?', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 5. Calculation (3 pts)
        ['begin group', 'g_calculation', 'Calculation', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'calculation_inst', 'Provide ways to pay using 1 dollar coins/bills, 5 dollar and 10 dollar bills for an object that costs exactly "13 Dollars" without asking for change. Scoring: 3 or 4 correct ways = 3 pts; 2 ways = 2 pts; 1 way = 1 pt; 0 ways = 0 pts.', '', '', '', '', '', '', '', '', '', '', ''],
        ['text', 'calc_combinations', 'Recorded payment combinations (optional)', '', 'e.g. 10+1+1+1, 5+5+1+1+1, 5+8x1', '', '', '', '', 'no', '', '', '', ''],
        ['select_one calc_score', 'calculation_score_select', 'Number of correct combinations for 13 Dollars', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 6. Abstraction (3 pts)
        ['begin group', 'g_abstraction', 'Abstraction', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'abstraction_inst_practice', 'Practice (not scored): "To which category do an orange and a banana belong?" (Fruit). Administer the 3 scored pairs below (one prompt permitted per item if first response is concrete).', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'abstraction_1', 'train - boat (acceptable: means of transportation, travelling, vehicles; NOT iron/engines/fuel)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'abstraction_2', 'north - south (acceptable: cardinal directions, cardinal points, directions, hemispheres, regions)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'abstraction_3', 'drum - flute (acceptable: musical instruments; NOT wood/sound)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 7. Delayed Recall (5 pts)
        ['begin group', 'g_delayed_recall', 'Delayed Recall', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'delayed_inst', '1 point is allocated ONLY for words recalled with No cue (free recall). Category cue and multiple-choice cue responses are recorded for clinical info (0 points).', '', '', '', '', '', '', '', '', '', '', ''],
        ['select_one yesno', 'delayed_rose', 'ROSE - Recall with No cue (1 pt if Yes)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_rose_cat', 'ROSE - Category cue (type of flower)', '', '', '', "${delayed_rose} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_rose_mc', 'ROSE - Multiple choice cue (rose / daisy / tulip)', '', '', '', "${delayed_rose} = 'no' and ${delayed_rose_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_chair', 'CHAIR - Recall with No cue (1 pt if Yes)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_chair_cat', 'CHAIR - Category cue (type of furniture)', '', '', '', "${delayed_chair} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_chair_mc', 'CHAIR - Multiple choice cue (table / chair / bed)', '', '', '', "${delayed_chair} = 'no' and ${delayed_chair_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_hand', 'HAND - Recall with No cue (1 pt if Yes)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_hand_cat', 'HAND - Category cue (body part)', '', '', '', "${delayed_hand} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_hand_mc', 'HAND - Multiple choice cue (foot / hand / knee)', '', '', '', "${delayed_hand} = 'no' and ${delayed_hand_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_blue', 'BLUE - Recall with No cue (1 pt if Yes)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_blue_cat', 'BLUE - Category cue (colour)', '', '', '', "${delayed_blue} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_blue_mc', 'BLUE - Multiple choice cue (blue / brown / red)', '', '', '', "${delayed_blue} = 'no' and ${delayed_blue_cat} = 'no'", '', '', 'yes', '', '', '', ''],

        ['select_one yesno', 'delayed_spoon', 'SPOON - Recall with No cue (1 pt if Yes)', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_spoon_cat', 'SPOON - Category cue (kitchen instrument)', '', '', '', "${delayed_spoon} = 'no'", '', '', 'yes', '', '', '', ''],
        ['select_one delayed_cue', 'delayed_spoon_mc', 'SPOON - Multiple choice cue (fork / spoon / knife)', '', '', '', "${delayed_spoon} = 'no' and ${delayed_spoon_cat} = 'no'", '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 8. Visuoperception (3 pts)
        ['begin group', 'g_visuoperception', 'Visuoperception', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'visuoperception_inst', 'Identify as many superimposed drawings as possible within 60 seconds (do not rotate picture). Scoring: 9-10 objects = 3 pts; 6-8 = 2 pts; 4-5 = 1 pt; 0-3 = 0 pts.', '', '', '', '', '', '', '', '', '', '', 'visuoperception.png'],
        ['select_one yesno', 'vp_scissors', '1. scissors', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_tshirt', '2. T-shirt', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_banana', '3. banana', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_lamp', '4. lamp', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_candle', '5. candle', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_watch', '6. watch', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_cup', '7. cup', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_leaf', '8. leaf', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_key', '9. key', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['select_one yesno', 'vp_spoon', '10. spoon', '', '', 'horizontal-compact', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 9. Naming (4 pts)
        ['begin group', 'g_naming', 'Naming', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'naming_inst', 'Point to each animal and say: "Tell me the name of this animal." 1 point awarded per exact animal.', '', '', '', '', '', '', '', '', '', '', 'naming.png'],
        ['select_one naming_score', 'naming_zebra', 'Animal 1: Zebra (Horse/donkey NOT accepted)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one naming_score', 'naming_peacock', 'Animal 2: Peacock (Bird/chicken/other bird NOT accepted)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one naming_score', 'naming_tiger', 'Animal 3: Tiger (Cheetah/leopard/black tiger NOT accepted)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['select_one naming_score', 'naming_butterfly', 'Animal 4: Butterfly (Insect/other insect NOT accepted)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # 10. Attention (3 pts: 1 + 2)
        ['begin group', 'g_attention', 'Attention', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['note', 'attention_white_inst', 'White background: Ask subject to read out loud numbers in the CIRCLES only (do not read squares or triangles). Error = wrong shape, omission, or incorrect order. Scoring: <=1 error = 1 pt; >=2 errors = 0 pts.', '', '', '', '', '', '', '', '', '', '', 'attention_white.png'],
        ['integer', 'attention_white_errors', 'White background - Number of errors (<=1 error = 1 pt; >=2 = 0 pts)', '', '', '', '', '. >= 0 and . <= 50', 'Enter errors between 0 and 50.', 'yes', '', '', '', ''],
        ['note', 'attention_black_inst', 'Black background: Ask subject to read out loud numbers in the CIRCLES and SQUARES across both rows (do not read triangles). Scoring: <=2 errors = 2 pts; 3 errors = 1 pt; >=4 errors = 0 pts.', '', '', '', '', '', '', '', '', '', '', 'attention_black.png'],
        ['integer', 'attention_black_errors', 'Black background - Number of errors (<=2 = 2 pts; 3 = 1 pt; >=4 = 0 pts)', '', '', '', '', '. >= 0 and . <= 50', 'Enter errors between 0 and 50.', 'yes', '', '', '', ''],
        ['time', 'end_time', 'End time (after completing Attention)', '', 'Optional: record end time (hh:mm).', '', '', '', '', 'no', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # Scoring calculations
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
        ['calculate', 'mocab_interpretation', '', '', '', '', '', '', '', '', 'if(${consent_reconfirmed} != "yes", "not assessed (no consent)", if(${mocab_total_score} >= 26, "normal", "impaired"))', '', '', ''],

        # Admin, total time & score summary display
        ['begin group', 'g_admin', 'Administration & Timing', '', '', 'field-list', '', '', '', '', '', '', '', ''],
        ['integer', 'total_time_min', 'Total administration time - minutes (optional)', '', '', '', consent_yes, '. >= 0 and . <= 180', 'Enter minutes between 0 and 180.', 'no', '', '', '', ''],
        ['integer', 'total_time_sec', 'Total administration time - seconds (optional)', '', '', '', consent_yes, '. >= 0 and . <= 59', 'Enter seconds between 0 and 59.', 'no', '', '', '', ''],
        ['string', 'administered_by', 'Administered by (staff name)', '', '', '', '', '', '', 'yes', '', '', '', ''],
        ['dateTime', 'assessment_datetime', 'Date & Time of Assessment', '', '', '', '', '', '', 'yes', "if(. = '', now(), .)", '', '', ''],
        ['text', 'clinical_notes', 'Clinical notes (optional)', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['note', 'score_summary_note', '**MoCA-B Total Score:** ${mocab_total_score} / 30 (Raw: ${mocab_raw_total} + Education <4y: +${mocab_education_adj} + Illiterate: +${mocab_literacy_adj})', '', '', '', consent_yes, '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],

        # Sheet upload (repeat group, matching moca_assessment)
        ['begin group', 'g_moca_b_sheets', 'MoCA-B sheet / supporting documents', '', '', 'field-list', consent_yes, '', '', '', '', '', '', ''],
        ['begin repeat', 'moca_b_sheets', 'MoCA-B sheet upload', '', 'Upload the completed paper MoCA-B sheet (photo or PDF).', 'field-list', '', '', '', '', '', '', '', ''],
        ['image', 'moca_b_image', 'Upload MoCA-B sheet photo', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['file', 'moca_b_pdf', 'Upload MoCA-B sheet PDF', '', '', '', '', '', '', 'no', '', '', '', ''],
        ['end repeat', '', '', '', '', '', '', '', '', '', '', '', '', ''],
        ['end group', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ]

    for r in rows:
        params = 'max-pixels=1024' if len(r) > 1 and r[1] == 'moca_b_image' else ''
        ws.append(r + [params])

    # ============================================================
    # CHOICES SHEET
    # ============================================================
    ws_choices = wb.create_sheet('choices')
    ws_choices.append(['list_name', 'name', 'label::en', 'label::te'])

    choices = [
        ['yesno', 'yes', 'Yes', 'అవును'],
        ['yesno', 'no', 'No', 'కాదు'],
        ['naming_score', 'yes', 'Correct', 'సరిగ్గా'],
        ['naming_score', 'no', 'Incorrect', 'తప్పు'],
        ['naming_score', 'dk', "Don't know / Not assessed", 'తెలియదు'],
        ['calc_score', '0', '0 correct combinations (0 pts)', '0 సరైనవి (0)'],
        ['calc_score', '1', '1 correct combination (1 pt)', '1 సరైనది (1)'],
        ['calc_score', '2', '2 correct combinations (2 pts)', '2 సరైనవి (2)'],
        ['calc_score', '3', '3 or 4 correct combinations (3 pts)', '3 లేదా 4 సరైనవి (3)'],
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
    ws_settings.append(['MoCA-B (Basic)', 'moca_b', '2026-11-12 01-00', 'en'])

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
