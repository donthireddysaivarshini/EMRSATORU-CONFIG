"""
create_nims_scales.py - Generates the 6 new standalone assessment scales
from the NIMS survey definition for Satoru CHT EMR:
1. cog5  - CoG5 - Memory Screening Battery
2. mse   - Mental State Examination (MSE)
3. phq9  - PHQ-9 Depression Screening
4. ad8   - AD8 Dementia Screening Interview
5. easi  - Everyday Abilities Scale for India (EASI)
6. gds15 - Geriatric Depression Scale (GDS-15)
"""

import os
import openpyxl

HEADERS_SURVEY = [
    'type', 'name', 'label', 'hint', 'required', 'appearance',
    'relevant', 'calculation', 'constraint', 'constraint_message', 'media::image'
]
HEADERS_CHOICES = ['list_name', 'name', 'label']
HEADERS_SETTINGS = ['form_title', 'form_id', 'version', 'default_language']

BASE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'forms', 'app')

CHOICES_YES_NO = [
    {'list_name': 'yes_no', 'name': 'yes', 'label': 'Yes'},
    {'list_name': 'yes_no', 'name': 'no', 'label': 'No'},
    {'list_name': 'yes_no', 'name': '_not_completed___unable_to_und', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
    {'list_name': 'yes_no', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
    {'list_name': 'yes_no', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},
]

def make_inputs_and_details(include_education=False):
    rows = [
        {'type': 'begin group', 'name': 'inputs', 'label': 'Participant', 'appearance': 'field-list', 'relevant': "./source = 'user'"},
        {'type': 'hidden', 'name': 'source'},
        {'type': 'hidden', 'name': 'source_id'},
        {'type': 'begin group', 'name': 'user', 'label': 'User', 'appearance': 'field-list'},
        {'type': 'hidden', 'name': 'contact_id'},
        {'type': 'hidden', 'name': 'name'},
        {'type': 'end group'},
        {'type': 'begin group', 'name': 'contact', 'label': 'Select participant', 'appearance': 'field-list'},
        {'type': 'db:person', 'name': '_id', 'label': 'Select participant', 'appearance': 'select-contact type-participant'},
        {'type': 'hidden', 'name': 'patient_id'},
        {'type': 'hidden', 'name': 'name'},
        {'type': 'hidden', 'name': 'date_of_birth'},
        {'type': 'hidden', 'name': 'sex'},
        {'type': 'hidden', 'name': 'age'},
        {'type': 'hidden', 'name': 'phone'},
        {'type': 'end group'},
        {'type': 'end group'},
        {'type': 'calculate', 'name': 'patient_uuid', 'calculation': '../inputs/contact/_id'},
        {'type': 'calculate', 'name': 'patient_id', 'calculation': 'coalesce(../inputs/contact/patient_id, ../inputs/contact/_id)'},
        {'type': 'calculate', 'name': 'patient_name', 'calculation': '../inputs/contact/name'},
        {'type': 'calculate', 'name': 'screened_by', 'calculation': '../inputs/user/contact_id'},
        {'type': 'calculate', 'name': 'screened_by_name', 'calculation': '../inputs/user/name'},
        {'type': 'begin group', 'name': 'g_participant', 'label': 'Participant details & consent', 'appearance': 'field-list'},
        {'type': 'select_one yes_no', 'name': 'consent_confirmed', 'label': 'Did the participant (or guardian) provide informed consent for this assessment?', 'required': 'yes'},
        {'type': 'string', 'name': 'patient_name_display', 'label': 'Participant Name', 'required': 'yes', 'calculation': "if(. = '', ${patient_name}, .)"},
        {'type': 'date', 'name': 'assessment_date', 'label': 'Date of assessment', 'required': 'yes', 'calculation': "if(. = '', today(), .)", 'constraint': '. <= today()', 'constraint_message': 'Date of assessment cannot be in the future.'},
        {'type': 'date', 'name': 'dob', 'label': 'Date of birth', 'required': 'yes', 'constraint': '. <= today()', 'constraint_message': 'Date of birth cannot be in the future.'}
    ]
    if include_education:
        rows.extend([
            {'type': 'select_one km1en86', 'name': 'education', 'label': 'Education level', 'required': 'yes'},
            {'type': 'text', 'name': 'education_specify', 'label': 'Please specify education:', 'relevant': "${education} = 'not_studied' or ${education} = 'others'"},
            {'type': 'calculate', 'name': 'score_education', 'calculation': "if(${education}='not_studied' or ${education}='elementary_middle', 1, 0)"}
        ])
    rows.append({'type': 'end group'})
    return rows

def save_workbook(form_id, title, survey_rows, choices_rows, version='2026-11-05 17-00'):
    wb = openpyxl.Workbook()
    ws_survey = wb.active
    ws_survey.title = 'survey'
    ws_survey.append(HEADERS_SURVEY)
    for r in survey_rows:
        ws_survey.append([r.get(h) for h in HEADERS_SURVEY])

    ws_choices = wb.create_sheet('choices')
    ws_choices.append(HEADERS_CHOICES)
    for c in choices_rows:
        ws_choices.append([c.get(h) for h in HEADERS_CHOICES])

    ws_settings = wb.create_sheet('settings')
    ws_settings.append(HEADERS_SETTINGS)
    ws_settings.append([title, form_id, version, 'en'])

    path = os.path.join(BASE_DIR, f'{form_id}.xlsx')
    wb.save(path)
    print(f'Saved: {path}')

# ===========================================================================
# 1. COG-5
# ===========================================================================
def build_cog5():
    form_id = 'cog5'
    title = 'CoG5 - Memory Screening Battery'
    survey = make_inputs_and_details(include_education=True)
    survey.extend([
        {'type': 'begin group', 'name': 'grp_cog5_scale', 'label': 'CoG5 - Memory Screening Battery'},
        {'type': 'begin group', 'name': 'grp_picture_based_memory_impairment', 'label': 'Component-1 (Picture-Based Memory Impairment Screen)', 'appearance': 'field-list'},
        {'type': 'note', 'name': 'pmis_image', 'label': 'Here is a sheet with 4 pictures for you to learn and remember. Each picture belongs to a different category. Please say what you see in each picture out loud to me. To help you learn, I will tell you the category for each picture and ask you to name the picture that goes with each category cue. (For example: Which one is fruit?) Remember these pictures because later I will ask you to recall them.', 'media::image': 'cow.png'},
        {'type': 'select_one identify_images', 'name': 'able_identify', 'label': 'Q1. Able to say identify the images and cues.', 'required': 'yes'},
        {'type': 'end group'},

        {'type': 'begin group', 'name': 'grp_picture_based_screening_test', 'label': 'Component -2 (Pattern based screening test)', 'appearance': 'field-list'},
        {'type': 'note', 'name': 'triangle_thing', 'label': 'For this next part I will ask you to name which number goes with each symbol. Look up here at the top, each number has a symbol associated with it. Down here, we have the same symbols, but the numbers are missing. I want you to tell me which number goes with each symbol. For example: See this mark right here (point to first practice item), on the key the mark goes with 4, so you would say 4. What would you say for this mark? Now that you know how to do these, I want you to start right here and tell me the numbers that go with each of these shapes. Go in order (use finger to demonstrate). When you finish one row, keep going to the next row until I tell you to stop. Work as fast as you can without making mistakes. You can use your finger to keep track of where you are. Do you have any questions? (Answer questions). Ready? Go ahead. Time 90 Sec.', 'media::image': 'tri.png'},
        {'type': 'select_one pbst_options', 'name': 'pbst_result', 'label': 'Q2a. Number of Symbol Matched within 90 seconds', 'required': 'yes'},
        {'type': 'calculate', 'name': 'score_pbst', 'calculation': "if(${pbst_result}='more_25', 2, if(${pbst_result}='less_equal_25' or ${pbst_result}='unable', 1, 999))"},
        {'type': 'select_one pbst_upload_opts', 'name': 'pbst_upload_status', 'label': 'Q2b. Upload Score Sheet', 'required': 'yes', 'relevant': "${pbst_result} != 'na' and ${pbst_result} != 'nc_instruct' and ${pbst_result} != 'nc_dropout'"},
        {'type': 'image', 'name': 'pbst_upload_file', 'label': 'Upload File', 'required': 'yes', 'relevant': "${pbst_upload_status} = 'upload_now' or ${pbst_upload_status} = 'upload_later'"},
        {'type': 'text', 'name': 'pbst_upload_reason', 'label': 'Specify Reason for Not Applicable', 'required': 'yes', 'relevant': "${pbst_upload_status} = 'not_applicable'"},
        {'type': 'end group'},

        {'type': 'begin group', 'name': 'grp_mcr', 'label': 'Component - 3.1(Motoric Cognitive Risk)', 'appearance': 'field-list'},
        {'type': 'select_one memory_general', 'name': 'memory_general', 'label': 'Q3. In general, how is your memory for a person your age?', 'required': 'yes'},
        {'type': 'select_one thinking_problem', 'name': 'thinking_problem', 'label': 'Q4. Aside from memory, have you noticed other problems with your thinking skills?', 'required': 'yes'},
        {'type': 'text', 'name': 'thinking_problem_reason', 'label': 'Please specify reason:', 'required': 'yes', 'relevant': "${thinking_problem} = 'yes'"},
        {'type': 'select_one memory_problem', 'name': 'memory_problem', 'label': 'Q5. Do you feel you have memory problems?', 'required': 'yes'},
        {'type': 'select_one memory_worse', 'name': 'memory_worse', 'label': 'Q6. Do you feel that your memory is becoming worse?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'mcr_cog_flag', 'calculation': "if(selected(${memory_general},'fair') or selected(${memory_general},'poor') or selected(${thinking_problem},'yes') or selected(${memory_problem},'minor') or selected(${memory_problem},'major') or selected(${memory_worse},'yes_not_worried') or selected(${memory_worse},'yes_worried'),1,0)"},
        {'type': 'end group'},

        {'type': 'begin group', 'name': 'grp_timed_walk', 'label': 'Component -3.2 (Timed Walk Protocol)', 'appearance': 'field-list'},
        {'type': 'note', 'name': 'walk_note', 'label': '**Timed Walk Protocol**'},
        {'type': 'note', 'name': 'walk_intro', 'label': 'Next, I would like to test whether you can walk a very short distance comfortably. First, I want to make sure it is safe to carry out the test.'},
        {'type': 'select_one yes_no', 'name': 'leg_pain', 'label': 'Q7. Do you have any pain in your legs today?', 'required': 'yes'},
        {'type': 'select_one restriction_type', 'name': 'health_restriction', 'label': 'Q8. Do you have any problems from recent surgery, injury, or other health conditions?', 'required': 'yes'},
        {'type': 'select_one device_ability', 'name': 'assistive_device', 'label': 'Q9. If participant has cane, walker, or other device: Are you able to use it?', 'required': 'yes'},
        {'type': 'select_one walk_time_options', 'name': 'walk_time_result', 'label': 'Q10. Timed Walk (2.5m)', 'required': 'yes', 'relevant': "${health_restriction} = 'none' and ${assistive_device} != 'no'"},
        {'type': 'calculate', 'name': 'mcr_motor_flag', 'calculation': "if(${health_restriction}='surgery' or ${health_restriction}='injury' or ${health_restriction}='other' or ${health_restriction}='na' or ${health_restriction}='nc_instruct' or ${health_restriction}='nc_dropout' or ${assistive_device}='no' or ${assistive_device}='na' or ${assistive_device}='nc_instruct' or ${assistive_device}='nc_dropout' or ${walk_time_result}='slow' or ${walk_time_result}='na' or ${walk_time_result}='nc_instruct' or ${walk_time_result}='nc_dropout' or ${walk_time_result}='', 1, 0)"},
        {'type': 'end group'},

        {'type': 'begin group', 'name': 'grp_pmis_recall', 'label': 'Component - 1(Picture Based Memory Impairement Recall)', 'appearance': 'field-list'},
        {'type': 'select_one recall_combination', 'name': 'recall_combination', 'label': 'Q11. Earlier I showed you 4 pictures. Tell me as many of the 4 pictures as you can remember.', 'required': 'yes'},
        {'type': 'select_one free_recall', 'name': 'free_recall', 'label': 'Q12. How many pictures were identified FREE RECALL?', 'required': 'yes'},
        {'type': 'select_one cued_recall', 'name': 'cued_recall', 'label': 'Q13. How many pictures were identified with CUED RECALL? (Cued recall - use semantic cues)', 'required': 'yes', 'relevant': "${free_recall} != 'na' and ${free_recall} != 'nc_instruct' and ${free_recall} != 'nc_dropout'"},
        {'type': 'end group'},

        {'type': 'begin group', 'name': 'grp_symptom_checklist', 'label': 'Symptom Checklist (Optional)'},
        {'type': 'note', 'name': 'symptom_note', 'label': '**Q14. In the last year, are you suffering from any of the below-mentioned symptoms? Tick if yes**'},
        {'type': 'select_one yes_no', 'name': 'sym_sad', 'label': 'a) Feeling sad/irritable, feeling less energetic?', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'sym_anx', 'label': 'b) Feeling anxious, fearful or restless?', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'sym_sleep', 'label': 'c) Sleep disturbance?', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'sym_daytime', 'label': 'd) Excessive daytime sleepiness?', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'sym_snore', 'label': 'e) Snoring in sleep?', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'sym_voice', 'label': 'f) Hearing voices or seeing images when alone?', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'sym_mem', 'label': 'g) Problems with memory, way finding?', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'sym_speak', 'label': 'h) Difficulty in speaking, understanding?', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'sym_neuro', 'label': 'i) Any neurological problems (slowness, tremor)?', 'required': 'no'},
        {'type': 'calculate', 'name': 'score_q13_symptoms', 'calculation': "if(${sym_sad}='yes',1,0) + if(${sym_anx}='yes',1,0) + if(${sym_sleep}='yes',1,0) + if(${sym_daytime}='yes',1,0) + if(${sym_snore}='yes',1,0) + if(${sym_voice}='yes',1,0) + if(${sym_mem}='yes',1,0) + if(${sym_speak}='yes',1,0) + if(${sym_neuro}='yes',1,0)"},
        {'type': 'end group'},

        {'type': 'begin group', 'name': 'grp_risk_factors', 'label': 'Dementia Risk Checklist (Optional)', 'appearance': 'field-list'},
        {'type': 'note', 'name': 'risk_note', 'label': '**Q15. Dementia Risk & Comorbidity Screening Checklist.**'},
        {'type': 'select_one yes_no', 'name': 'risk_htn', 'label': 'a. Hypertension', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_dm', 'label': 'b. Diabetes', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_ldl', 'label': 'c. High LDL Cholesterol (dyslipidemia)', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_head', 'label': 'd. Serious head injury (unconscious or hospitalized)', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_urine', 'label': 'e. Urinary Incontinence (loss of bladder control)', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_alc', 'label': 'f. Alcohol use', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_phys', 'label': 'g. Physical Inactivity (lack of regular exercise)', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_obes', 'label': 'h. Overweight/Obesity', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_hear', 'label': 'i. Hearing impairment', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_vis', 'label': 'j. Visual impairment', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_soc', 'label': 'k. Social isolation', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_smoke', 'label': 'l. Current or past smoking/tobacco use', 'required': 'no'},
        {'type': 'select_one yes_no', 'name': 'risk_dep', 'label': 'm. From past 2 weeks feeling sad, low mood', 'required': 'no'},
        {'type': 'calculate', 'name': 'score_dementia_risk', 'calculation': "if(${risk_htn}='yes',1,0) + if(${risk_dm}='yes',1,0) + if(${risk_ldl}='yes',1,0) + if(${risk_head}='yes',1,0) + if(${risk_urine}='yes',1,0) + if(${risk_smoke}='yes',1,0) + if(${risk_alc}='yes',1,0) + if(${risk_phys}='yes',1,0) + if(${risk_obes}='yes',1,0) + if(${risk_hear}='yes',1,0) + if(${risk_vis}='yes',1,0) + if(${risk_dep}='yes',1,0) + if(${risk_soc}='yes',1,0) + ${score_education}"},
        {'type': 'end group'},

        # PMIS scoring
        {'type': 'calculate', 'name': 'calc_free_score', 'calculation': "if(${free_recall}='four',8,if(${free_recall}='three',6,if(${free_recall}='two',4,if(${free_recall}='one',2,if(${free_recall}='zero',0,999)))))"},
        {'type': 'calculate', 'name': 'calc_cued_score', 'calculation': "if(${cued_recall}='four',4,if(${cued_recall}='three',3,if(${cued_recall}='two',2,if(${cued_recall}='one',1,if(${cued_recall}='zero',0,0)))))"},
        {'type': 'calculate', 'name': 'score_pmis_total', 'calculation': "if(${calc_free_score} = 999, 999, ${calc_free_score} + ${calc_cued_score})"},
        {'type': 'calculate', 'name': 'pmis_flag', 'calculation': "if(${recall_combination}='other_combo' or ${able_identify}='no',1,0)"},
        {'type': 'calculate', 'name': 'PMIS_FINAL_RESULT', 'calculation': "if(${score_pmis_total}=999, 'DROPOUT/UNABLE TO UNDERSTAND', if(${score_pmis_total}=8 and ${pmis_flag}=0,'NORMAL','POSSIBLE IMPAIRMENT'))"},
        {'type': 'calculate', 'name': 'PBST_FINAL_RESULT', 'calculation': "if(${score_pbst}=2, 'NORMAL', if(${score_pbst}=1, 'POSSIBLE COGNITIVE SLOWING', 'DROPOUT/UNABLE TO UNDERSTAND'))"},
        {'type': 'calculate', 'name': 'MCR_FINAL_RESULT', 'calculation': "if(${mcr_cog_flag}=1 and ${mcr_motor_flag}=1,'POSITIVE: Cognitive Complaints & Motor Issue present',if(${mcr_cog_flag}=0 and ${mcr_motor_flag}=1,'NEGATIVE: Motor Issue present, No Cognitive Complaints',if(${mcr_cog_flag}=1 and ${mcr_motor_flag}=0,'NEGATIVE: Cognitive Complaints present, Normal Motor',if(${mcr_cog_flag}=0 and ${mcr_motor_flag}=0,'NEGATIVE: No Cognitive Complaints & Normal Motor','DROPOUT/UNABLE TO UNDERSTAND'))))"},
        {'type': 'calculate', 'name': 'FINAL_TRIAGE', 'calculation': "if(string(${pmis_flag})='1' or (number(${score_pmis_total})<8 and number(${score_pmis_total})!=999) or string(${score_pbst})='1' or (string(${mcr_cog_flag})='1' and string(${mcr_motor_flag})='1'),'FLAG FOR CLINICAL EVALUATION','ROUTINE MONITORING')"},
        {'type': 'calculate', 'name': 'calc_pmis_meaning', 'calculation': "if(${score_pmis_total}=999, 'The participant was unable to complete this test.', if(${score_pmis_total}=8 and ${pmis_flag}=0, 'The participant remembered the pictures well — memory appears normal.', 'The participant had difficulty remembering the pictures — this may suggest a memory concern.'))"},
        {'type': 'calculate', 'name': 'calc_pbst_meaning', 'calculation': "if(${score_pbst}=2, 'The participant matched symbols quickly — thinking speed appears normal.', if(${score_pbst}=1, 'The participant was a bit slow at matching symbols — this may suggest mild thinking slowness.', 'The participant was unable to complete this test.'))"},
        {'type': 'calculate', 'name': 'calc_mcr_meaning', 'calculation': "if(${mcr_cog_flag}=1 and ${mcr_motor_flag}=1, 'The participant reported memory concerns AND walks slowly — both are risk signs for dementia.', if(${mcr_cog_flag}=0 and ${mcr_motor_flag}=1, 'The participant walks slowly but does not report memory concerns — walking speed alone is a mild risk sign.', if(${mcr_cog_flag}=1 and ${mcr_motor_flag}=0, 'The participant reports memory concerns but walks at a normal speed — memory alone is a mild risk sign.', if(${mcr_cog_flag}=0 and ${mcr_motor_flag}=0, 'No memory concerns and normal walking speed — no risk signs detected.', 'The participant was unable to complete this test.'))))"},
        {'type': 'calculate', 'name': 'calc_risk_breakdown', 'calculation': "concat(if(${risk_htn}='yes', 'Hypertension, ', ''), if(${risk_dm}='yes', 'Diabetes, ', ''), if(${risk_ldl}='yes', 'High LDL Cholesterol, ', ''), if(${risk_head}='yes', 'Serious Head Injury, ', ''), if(${risk_urine}='yes', 'Urinary Incontinence, ', ''), if(${risk_alc}='yes', 'Alcohol Use, ', ''), if(${risk_phys}='yes', 'Physical Inactivity, ', ''), if(${risk_obes}='yes', 'Overweight / Obesity, ', ''), if(${risk_hear}='yes', 'Hearing Impairment, ', ''), if(${risk_vis}='yes', 'Visual Impairment, ', ''), if(${risk_soc}='yes', 'Social Isolation, ', ''), if(${risk_smoke}='yes', 'Smoking / Tobacco Use, ', ''), if(${risk_dep}='yes', 'Low Mood (past 2 weeks), ', ''), if(number(${score_education})>0, 'Low Education Level', ''))"},
        {'type': 'note', 'name': 'results_summary', 'label': '*** TEST COMPLETE: RESULTS ***'},
        {'type': 'note', 'name': 'display_pmis', 'label': '1. Picture Based Memory Impairment (PMIS) Result: ${PMIS_FINAL_RESULT} (Score: ${score_pmis_total} / 8)\n   👉 ${calc_pmis_meaning}'},
        {'type': 'note', 'name': 'display_pbst', 'label': '2. Pattern Based Screening (PBST) Result: ${PBST_FINAL_RESULT} (Score: ${score_pbst} / 2)\n   👉 ${calc_pbst_meaning}'},
        {'type': 'note', 'name': 'display_mcr', 'label': '3. Motoric Cognitive Risk (MCR) Result: ${MCR_FINAL_RESULT}\n   👉 ${calc_mcr_meaning}'},
        {'type': 'note', 'name': 'display_risk', 'label': '4. Dementia Risk Score: ${score_dementia_risk} / 14\n   Risk Factors Identified: ${calc_risk_breakdown}'},
        {'type': 'note', 'name': 'display_triage', 'label': 'OVERALL TRIAGE: ${FINAL_TRIAGE}'},
        {'type': 'calculate', 'name': 'calc_overall_conclusion', 'calculation': "if(string(${pmis_flag})='1' or (number(${score_pmis_total})<8 and number(${score_pmis_total})!=999) or string(${score_pbst})='1' or (string(${mcr_cog_flag})='1' and string(${mcr_motor_flag})='1'), 'SCREENED POSITIVE: Higher risk for dementia. Proceed to Stage-2.', 'SCREENED NEGATIVE: Lower risk for dementia.')"},
        {'type': 'note', 'name': 'display_overall', 'label': '**Overall conclusion :** ${calc_overall_conclusion}'},
        {'type': 'select_one cog5_status', 'name': 'Survey_Status_cog5', 'label': 'Survey Status', 'hint': 'Mark if the survey was completed or dropped out midway.', 'required': 'yes'},
        {'type': 'text', 'name': 'Reason_for_Dropout_cog5', 'label': 'Reason for Dropout / Refusal:', 'hint': 'Please explain why the survey was stopped.', 'required': 'yes', 'relevant': "${Survey_Status_cog5} = 'dropout___discontinued'"},
        {'type': 'end group'}
    ])

    choices = list(CHOICES_YES_NO)
    choices.extend([
        {'list_name': 'km1en86', 'name': 'not_studied', 'label': 'Not studied in school'},
        {'list_name': 'km1en86', 'name': 'elementary_middle', 'label': 'Elementary - middle school'},
        {'list_name': 'km1en86', 'name': 'high_school', 'label': 'Upto High School'},
        {'list_name': 'km1en86', 'name': 'completed', 'label': 'Completed High School'},
        {'list_name': 'km1en86', 'name': 'sslcpuc', 'label': 'S.S.L.C. /PUC (Intermediate)'},
        {'list_name': 'km1en86', 'name': 'vocational_technical', 'label': 'Vocational/ Technical'},
        {'list_name': 'km1en86', 'name': 'graduate_degree', 'label': 'Graduate - Degree'},
        {'list_name': 'km1en86', 'name': 'post_graduate', 'label': 'Post Graduate'},
        {'list_name': 'km1en86', 'name': 'phd', 'label': 'PhD'},
        {'list_name': 'km1en86', 'name': 'others', 'label': 'Others'},
        {'list_name': 'km1en86', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'km1en86', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'km1en86', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'identify_images', 'name': 'yes', 'label': 'Yes, in less than 5 attempts'},
        {'list_name': 'identify_images', 'name': 'no', 'label': 'No, the client was provided the correct name and cue of each item.'},
        {'list_name': 'identify_images', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'identify_images', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'identify_images', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'pbst_options', 'name': 'more_25', 'label': 'MORE THAN 25: Normal'},
        {'list_name': 'pbst_options', 'name': 'less_equal_25', 'label': 'LESS OR EQUAL TO 25: Possible cognitive impairment'},
        {'list_name': 'pbst_options', 'name': 'unable', 'label': 'Unable to understand practice'},
        {'list_name': 'pbst_options', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'pbst_options', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'pbst_options', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'pbst_upload_opts', 'name': 'upload_now', 'label': 'Upload now'},
        {'list_name': 'pbst_upload_opts', 'name': 'upload_later', 'label': 'Upload later'},
        {'list_name': 'pbst_upload_opts', 'name': 'not_applicable', 'label': 'Not applicable'},

        {'list_name': 'memory_general', 'name': 'good', 'label': 'Good'},
        {'list_name': 'memory_general', 'name': 'fair', 'label': 'Fair'},
        {'list_name': 'memory_general', 'name': 'poor', 'label': 'Poor'},
        {'list_name': 'memory_general', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'memory_general', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'memory_general', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'thinking_problem', 'name': 'no', 'label': 'No'},
        {'list_name': 'thinking_problem', 'name': 'yes', 'label': 'Yes'},
        {'list_name': 'thinking_problem', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'thinking_problem', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'thinking_problem', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'memory_problem', 'name': 'none', 'label': 'No problems'},
        {'list_name': 'memory_problem', 'name': 'minor', 'label': 'Some minor problems'},
        {'list_name': 'memory_problem', 'name': 'major', 'label': 'Major problems'},
        {'list_name': 'memory_problem', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'memory_problem', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'memory_problem', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'memory_worse', 'name': 'no', 'label': 'No problems'},
        {'list_name': 'memory_worse', 'name': 'yes_not_worried', 'label': 'Yes, but this does not worry me'},
        {'list_name': 'memory_worse', 'name': 'yes_worried', 'label': 'Yes, and this does worry me'},
        {'list_name': 'memory_worse', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'memory_worse', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'memory_worse', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'restriction_type', 'name': 'none', 'label': 'No apparent restriction (Continue)'},
        {'list_name': 'restriction_type', 'name': 'surgery', 'label': 'Yes, recent surgery (Skip Walk)'},
        {'list_name': 'restriction_type', 'name': 'injury', 'label': 'Yes, injury (Skip Walk)'},
        {'list_name': 'restriction_type', 'name': 'other', 'label': 'Yes, other health condition (Skip Walk)'},
        {'list_name': 'restriction_type', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'restriction_type', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'restriction_type', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'device_ability', 'name': 'yes', 'label': 'Yes, able to walk WITHOUT device'},
        {'list_name': 'device_ability', 'name': 'no', 'label': 'No, cannot walk without device'},
        {'list_name': 'device_ability', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'device_ability', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'device_ability', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'walk_time_options', 'name': 'normal', 'label': 'LESS THAN or EQUAL TO 5.5 seconds'},
        {'list_name': 'walk_time_options', 'name': 'slow', 'label': 'MORE THAN 5.5 seconds'},
        {'list_name': 'walk_time_options', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'walk_time_options', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'walk_time_options', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'recall_combination', 'name': 'four_free', 'label': '4 pictures Free recall'},
        {'list_name': 'recall_combination', 'name': 'three_free_one_cued', 'label': '3 pictures Free recall and 1 Cued recall'},
        {'list_name': 'recall_combination', 'name': 'other_combo', 'label': 'Any other combination'},
        {'list_name': 'recall_combination', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'recall_combination', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'recall_combination', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'free_recall', 'name': 'four', 'label': '4 pictures Free recall'},
        {'list_name': 'free_recall', 'name': 'three', 'label': '3 pictures Free recall'},
        {'list_name': 'free_recall', 'name': 'two', 'label': '2 pictures Free recall'},
        {'list_name': 'free_recall', 'name': 'one', 'label': '1 picture Free recall'},
        {'list_name': 'free_recall', 'name': 'zero', 'label': 'No pictures Free recall'},
        {'list_name': 'free_recall', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'free_recall', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'free_recall', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'cued_recall', 'name': 'four', 'label': '4 pictures recalled via Cue'},
        {'list_name': 'cued_recall', 'name': 'three', 'label': '3 pictures recalled via Cue'},
        {'list_name': 'cued_recall', 'name': 'two', 'label': '2 pictures recalled via Cue'},
        {'list_name': 'cued_recall', 'name': 'one', 'label': '1 picture recalled via Cue'},
        {'list_name': 'cued_recall', 'name': 'zero', 'label': 'No picture recalled via Cue'},
        {'list_name': 'cued_recall', 'name': 'na', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'cued_recall', 'name': 'nc_instruct', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'cued_recall', 'name': 'nc_dropout', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'cog5_status', 'name': 'completed', 'label': 'Completed'},
        {'list_name': 'cog5_status', 'name': 'dropout___discontinued', 'label': 'Dropout / Discontinued'}
    ])
    save_workbook(form_id, title, survey, choices)

# ===========================================================================
# 2. MSE
# ===========================================================================
def build_mse():
    form_id = 'mse'
    title = 'Mental State Examination (MSE)'
    survey = make_inputs_and_details()
    survey.extend([
        {'type': 'begin group', 'name': 'grp_mse_scale', 'label': 'Mental State Examination (MSE)'},
        {'type': 'select_one score_1', 'name': 'q1', 'label': '1. Is it morning or afternoon or evening?', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q2', 'label': '2. What day of the week is today?', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q3', 'label': '3. What date is it today?', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q4', 'label': '4. Which month is today?', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q5', 'label': '5. What season of the year is this?', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q6', 'label': '6. Under which post office does your village come?', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q7', 'label': '7. Which district does your village fall under?', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q8', 'label': '8. Which village are you from?', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q9', 'label': '9. Which block (if village has only blocks) OR Which number area is this?', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q10', 'label': '10. Which place is this?', 'required': 'yes'},
        {'type': 'select_one score_3', 'name': 'q11', 'label': '11. (I went to Delhi and brought three things – Mango, Chair and Coin) Can you tell me what are the three things I brought from Delhi?', 'required': 'yes'},
        {'type': 'select_one score_5', 'name': 'q12_a_b', 'label': '12.a. Now can you tell me the names of the days of the week starting from Sunday? b.Now can you tell me the names of the days backwards', 'required': 'yes'},
        {'type': 'select_one score_3', 'name': 'q13', 'label': '13. What are the names of the three things, which I told you that I have brought from Delhi?', 'required': 'yes'},
        {'type': 'note', 'name': 'q14_note', 'label': '14. (Show the subject the Mobile Phone/Wristwatch and pen) Can you tell me the name of these objects? (If Yes, items 15 & 16 apply.) (If No, item 15a applies.)'},
        {'type': 'select_one score_1', 'name': 'q15', 'label': '15a. Show him/her the Mobile Phone/Wristwatch and say – “What is this?” OR b.If necessary, Identification of watch by touching and ask, “What is this?”', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q16', 'label': '16. Show him/her the pen and say – “What is this?” OR If necessary, Identification of pen by touching and ask – “What is this?”', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q17', 'label': '17. Now I am going to say something, listen carefully and repeat it exactly as I say after I finish. Phrase : “NEITHER THIS NOR THAT”.', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q18', 'label': '18. Now look at my face and do exactly what I do. Now close your eyes.', 'required': 'yes'},
        {'type': 'select_one score_3', 'name': 'q19', 'label': '19. First take the paper in your right hand, then with your both hands, fold it into half once and then gives the paper back to me.', 'required': 'yes'},
        {'type': 'select_one score_1', 'name': 'q20', 'label': '20. Now say a line about your house? (something specifically about your houses)', 'hint': 'NOT INCLUDED IN MSE TOTAL if given -1, not given -0', 'required': 'no'},
        {'type': 'select_one score_3', 'name': 'q21', 'label': '21. Here is a drawing, you must copy this drawing exactly as shown in the space provided here.', 'hint': 'Score: Must draw two four-sided figures =1 One figure should be manually inside the other = 2 Orientation of the figure should be obviously appropriate =3', 'required': 'yes', 'media::image': 'drawing.png'},
        {'type': 'calculate', 'name': 'total_score', 'calculation': "if(number(${q1})<90, number(${q1}), 0) + if(number(${q2})<90, number(${q2}), 0) + if(number(${q3})<90, number(${q3}), 0) + if(number(${q4})<90, number(${q4}), 0) + if(number(${q5})<90, number(${q5}), 0) + if(number(${q6})<90, number(${q6}), 0) + if(number(${q7})<90, number(${q7}), 0) + if(number(${q8})<90, number(${q8}), 0) + if(number(${q9})<90, number(${q9}), 0) + if(number(${q10})<90, number(${q10}), 0) + if(number(${q11})<90, number(${q11}), 0) + if(number(${q12_a_b})<90, number(${q12_a_b}), 0) + if(number(${q13})<90, number(${q13}), 0) + if(number(${q15})<90, number(${q15}), 0) + if(number(${q16})<90, number(${q16}), 0) + if(number(${q17})<90, number(${q17}), 0) + if(number(${q18})<90, number(${q18}), 0) + if(number(${q19})<90, number(${q19}), 0) + if(number(${q21})<90, number(${q21}), 0)"},
        {'type': 'note', 'name': 'display_score_mse', 'label': '**Mental State Examination Score:** ${total_score} / 31\n\n**Interpretation:** Scores <18 (illiterates) or <27 (literates) indicate significant cognitive impairment and need detailed evaluation.'},
        {'type': 'select_one mse_status', 'name': 'Survey_Status_mse', 'label': 'Survey Status', 'hint': 'Mark if the survey was completed or dropped out midway.', 'required': 'yes'},
        {'type': 'text', 'name': 'Reason_for_Dropout_mse', 'label': 'Reason for Dropout / Refusal:', 'hint': 'Please explain why the survey was stopped.', 'required': 'yes', 'relevant': "${Survey_Status_mse} = 'dropout___discontinued'"},
        {'type': 'end group'}
    ])

    choices = list(CHOICES_YES_NO)
    choices.extend([
        {'list_name': 'score_1', 'name': '0', 'label': 'Wrong'},
        {'list_name': 'score_1', 'name': '1', 'label': 'Correct'},
        {'list_name': 'score_1', 'name': '99', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'score_1', 'name': '98', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'score_1', 'name': '97', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'score_3', 'name': '0', 'label': '0 Correct'},
        {'list_name': 'score_3', 'name': '1', 'label': '1 Correct'},
        {'list_name': 'score_3', 'name': '2', 'label': '2 Correct'},
        {'list_name': 'score_3', 'name': '3', 'label': '3 Correct'},
        {'list_name': 'score_3', 'name': '99', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'score_3', 'name': '98', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'score_3', 'name': '97', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'score_5', 'name': '0', 'label': '0 Correct'},
        {'list_name': 'score_5', 'name': '1', 'label': '1 Correct'},
        {'list_name': 'score_5', 'name': '2', 'label': '2 Correct'},
        {'list_name': 'score_5', 'name': '3', 'label': '3 Correct'},
        {'list_name': 'score_5', 'name': '4', 'label': '4 Correct'},
        {'list_name': 'score_5', 'name': '5', 'label': '5 Correct'},
        {'list_name': 'score_5', 'name': '99', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'score_5', 'name': '98', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'score_5', 'name': '97', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'mse_status', 'name': 'completed', 'label': 'Completed'},
        {'list_name': 'mse_status', 'name': 'dropout___discontinued', 'label': 'Dropout / Discontinued'}
    ])
    save_workbook(form_id, title, survey, choices)

# ===========================================================================
# 3. PHQ-9
# ===========================================================================
def build_phq9():
    form_id = 'phq9'
    title = 'PHQ-9 Depression Screening'
    survey = make_inputs_and_details()
    survey.extend([
        {'type': 'begin group', 'name': 'grp_phq9_scale', 'label': 'PHQ-9 Depression Screening'},
        {'type': 'select_one phq_freq', 'name': 'phq9_q1', 'label': '1. Little interest or pleasure in doing things', 'required': 'yes'},
        {'type': 'select_one phq_freq', 'name': 'phq9_q2', 'label': '2. Feeling down, depressed, or hopeless', 'required': 'yes'},
        {'type': 'select_one phq_freq', 'name': 'phq9_q3', 'label': '3. Trouble falling or staying asleep, or sleeping too much', 'required': 'yes'},
        {'type': 'select_one phq_freq', 'name': 'phq9_q4', 'label': '4. Feeling tired or having little energy', 'required': 'yes'},
        {'type': 'select_one phq_freq', 'name': 'phq9_q5', 'label': '5. Poor appetite or overeating', 'required': 'yes'},
        {'type': 'select_one phq_freq', 'name': 'phq9_q6', 'label': '6. Feeling bad about yourself — or that you are a failure or have let yourself or your family down', 'required': 'yes'},
        {'type': 'select_one phq_freq', 'name': 'phq9_q7', 'label': '7. Trouble concentrating on things, such as reading the newspaper or watching television', 'required': 'yes'},
        {'type': 'select_one phq_freq', 'name': 'phq9_q8', 'label': '8. Moving or speaking so slowly that other people could have noticed? Or the opposite — being so fidgety or restless that you have been moving around a lot more than usual', 'required': 'yes'},
        {'type': 'select_one phq_freq', 'name': 'phq9_q9', 'label': '9. Thoughts that you would be better off dead or of hurting yourself in some way', 'required': 'yes'},
        {'type': 'calculate', 'name': 'phq9_total', 'calculation': "if(number(${phq9_q1})<90, number(${phq9_q1}), 0) + if(number(${phq9_q2})<90, number(${phq9_q2}), 0) + if(number(${phq9_q3})<90, number(${phq9_q3}), 0) + if(number(${phq9_q4})<90, number(${phq9_q4}), 0) + if(number(${phq9_q5})<90, number(${phq9_q5}), 0) + if(number(${phq9_q6})<90, number(${phq9_q6}), 0) + if(number(${phq9_q7})<90, number(${phq9_q7}), 0) + if(number(${phq9_q8})<90, number(${phq9_q8}), 0) + if(number(${phq9_q9})<90, number(${phq9_q9}), 0)"},
        {'type': 'select_one phq_difficulty', 'name': 'phq9_difficulty', 'label': 'If you checked off any problems, how difficult have these problems made it for you to do your work, take care of things at home, or get along with other people?', 'required': 'no', 'relevant': "${phq9_total} > 0"},
        {'type': 'note', 'name': 'display_score_phq9', 'label': '**PHQ-9 (Depression) Score:** ${phq9_total} / 27\n\n**Interpretation:** Scores 5-9 mild depression, 10-14 moderate depression, 15-19 moderately severe depression, 20-27 severe depression.'},
        {'type': 'select_one phq9_status', 'name': 'Survey_Status_phq9', 'label': 'Survey Status', 'hint': 'Mark if the survey was completed or dropped out midway.', 'required': 'yes'},
        {'type': 'text', 'name': 'Reason_for_Dropout_phq9', 'label': 'Reason for Dropout / Refusal:', 'hint': 'Please explain why the survey was stopped.', 'required': 'yes', 'relevant': "${Survey_Status_phq9} = 'dropout___discontinued'"},
        {'type': 'end group'}
    ])

    choices = list(CHOICES_YES_NO)
    choices.extend([
        {'list_name': 'phq_freq', 'name': '0', 'label': 'Not at all'},
        {'list_name': 'phq_freq', 'name': '1', 'label': 'Several days'},
        {'list_name': 'phq_freq', 'name': '2', 'label': 'More than half the days'},
        {'list_name': 'phq_freq', 'name': '3', 'label': 'Nearly every day'},
        {'list_name': 'phq_freq', 'name': '99', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'phq_freq', 'name': '98', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'phq_freq', 'name': '97', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'phq_difficulty', 'name': '0', 'label': 'Not difficult at all'},
        {'list_name': 'phq_difficulty', 'name': '1', 'label': 'Somewhat difficult'},
        {'list_name': 'phq_difficulty', 'name': '2', 'label': 'Very difficult'},
        {'list_name': 'phq_difficulty', 'name': '3', 'label': 'Extremely difficult'},
        {'list_name': 'phq_difficulty', 'name': '99', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'phq_difficulty', 'name': '98', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'phq_difficulty', 'name': '97', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'phq9_status', 'name': 'completed', 'label': 'Completed'},
        {'list_name': 'phq9_status', 'name': 'dropout___discontinued', 'label': 'Dropout / Discontinued'}
    ])
    save_workbook(form_id, title, survey, choices)

# ===========================================================================
# 4. AD-8
# ===========================================================================
def build_ad8():
    form_id = 'ad8'
    title = 'AD8 Dementia Screening Interview'
    survey = make_inputs_and_details()
    survey.extend([
        {'type': 'begin group', 'name': 'grp_ad8_scale', 'label': 'AD8 Dementia Screening Interview'},
        {'type': 'note', 'name': '_Instructions_ad8', 'label': '**Instructions:** Remember, "Yes, a change" indicates that there has been a change in the last several years caused by cognitive (thinking and memory) problems.'},
        {'type': 'select_one ad8_options', 'name': 'ad8_q1', 'label': '1. Problems with judgment (bad financial decisions, problems with thinking)', 'required': 'yes'},
        {'type': 'select_one ad8_options', 'name': 'ad8_q2', 'label': '2. Less interest in hobbies/activities', 'required': 'yes'},
        {'type': 'select_one ad8_options', 'name': 'ad8_q3', 'label': '3. Repeats the same things over and over (questions, stories, statements)', 'required': 'yes'},
        {'type': 'select_one ad8_options', 'name': 'ad8_q4', 'label': '4. Trouble learning how to use a tool, appliance, or gadget', 'required': 'yes'},
        {'type': 'select_one ad8_options', 'name': 'ad8_q5', 'label': '5. Forgets correct month or year', 'required': 'yes'},
        {'type': 'select_one ad8_options', 'name': 'ad8_q6', 'label': '6. Trouble handling complicated financial affairs (taxes, bills, checkbook)', 'required': 'yes'},
        {'type': 'select_one ad8_options', 'name': 'ad8_q7', 'label': '7. Trouble remembering appointments', 'required': 'yes'},
        {'type': 'select_one ad8_options', 'name': 'ad8_q8', 'label': '8. Daily problems with thinking and/or memory', 'required': 'yes'},
        {'type': 'calculate', 'name': 'ad8_total', 'calculation': "if(number(${ad8_q1})<90, number(${ad8_q1}), 0) + if(number(${ad8_q2})<90, number(${ad8_q2}), 0) + if(number(${ad8_q3})<90, number(${ad8_q3}), 0) + if(number(${ad8_q4})<90, number(${ad8_q4}), 0) + if(number(${ad8_q5})<90, number(${ad8_q5}), 0) + if(number(${ad8_q6})<90, number(${ad8_q6}), 0) + if(number(${ad8_q7})<90, number(${ad8_q7}), 0) + if(number(${ad8_q8})<90, number(${ad8_q8}), 0)"},
        {'type': 'note', 'name': 'display_score_ad8', 'label': '**AD8 (Dementia) Score:** ${ad8_total} / 8\n\n**Interpretation:** Scores of 2 or above suggest that cognitive impairment is likely to be present and may require further assessment.'},
        {'type': 'select_one ad8_status', 'name': 'Survey_Status_ad8', 'label': 'Survey Status', 'hint': 'Mark if the survey was completed or dropped out midway.', 'required': 'yes'},
        {'type': 'text', 'name': 'Reason_for_Dropout_ad8', 'label': 'Reason for Dropout / Refusal:', 'hint': 'Please explain why the survey was stopped.', 'required': 'yes', 'relevant': "${Survey_Status_ad8} = 'dropout___discontinued'"},
        {'type': 'end group'}
    ])

    choices = list(CHOICES_YES_NO)
    choices.extend([
        {'list_name': 'ad8_options', 'name': '1', 'label': 'Yes, a change'},
        {'list_name': 'ad8_options', 'name': '0', 'label': 'No, no change'},
        {'list_name': 'ad8_options', 'name': '99', 'label': "Don't Know/ NA"},
        {'list_name': 'ad8_options', 'name': '98', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'ad8_options', 'name': '97', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'ad8_status', 'name': 'completed', 'label': 'Completed'},
        {'list_name': 'ad8_status', 'name': 'dropout___discontinued', 'label': 'Dropout / Discontinued'}
    ])
    save_workbook(form_id, title, survey, choices)

# ===========================================================================
# 5. EASI
# ===========================================================================
def build_easi():
    form_id = 'easi'
    title = 'Everyday Abilities Scale for India (EASI)'
    survey = make_inputs_and_details()
    survey.extend([
        {'type': 'begin group', 'name': 'grp_easi_scale', 'label': 'Everyday Abilities Scale for India (EASI)'},
        {'type': 'note', 'name': '_Instructions_easi', 'label': '**Instructions:** The question can be asked by the caregiver or if not available directly from participant.'},
        {'type': 'select_one easi_yesno', 'name': 'easi_raw1', 'label': '1. Forgets he/she has just eaten and asks for food again', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q1', 'calculation': "if(${easi_raw1}='' or number(${easi_raw1})>=97, 0, number(${easi_raw1}))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw2', 'label': '2. Urinates in an appropriate place', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q2', 'calculation': "if(${easi_raw2}='' or number(${easi_raw2})>=97, 0, if(number(${easi_raw2})=1, 0, 1))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw3', 'label': '3. Clothes get dirty from urine or stools', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q3', 'calculation': "if(${easi_raw3}='' or number(${easi_raw3})>=97, 0, number(${easi_raw3}))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw4', 'label': '4. Is his/her shirt buttoned properly?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q4', 'calculation': "if(${easi_raw4}='' or number(${easi_raw4})>=97, 0, if(number(${easi_raw4})=1, 0, 1))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw5', 'label': '5. Does he/she have their Dhoti/petticoat/saree tied properly?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q5', 'calculation': "if(${easi_raw5}='' or number(${easi_raw5})>=97, 0, if(number(${easi_raw5})=1, 0, 1))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw6', 'label': '6. He/she able to work as a member of a team i.e in a group activity, which requires different roles from people, will he/she be able to participate?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q6', 'calculation': "if(${easi_raw6}='' or number(${easi_raw6})>=97, 0, if(number(${easi_raw6})=1, 0, 1))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw7', 'label': '7. Does he/she express his/her opinion appropriately on important family matters, eg: marriage?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q7', 'calculation': "if(${easi_raw7}='' or number(${easi_raw7})>=97, 0, if(number(${easi_raw7})=1, 0, 1))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw8', 'label': '8. If he/she is assigned or himself/herself decides to undertake an important task, can he/she follow it through completion?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q8', 'calculation': "if(${easi_raw8}='' or number(${easi_raw8})>=97, 0, if(number(${easi_raw8})=1, 0, 1))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw9', 'label': '9. Is he/she able to remember important festivals such as Holi, Diwali?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q9', 'calculation': "if(${easi_raw9}='' or number(${easi_raw9})>=97, 0, if(number(${easi_raw9})=1, 0, 1))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw10', 'label': '10. If he/she is asked to deliver a message, does he/she remember to do so?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q10', 'calculation': "if(${easi_raw10}='' or number(${easi_raw10})>=97, 0, if(number(${easi_raw10})=1, 0, 1))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw11', 'label': '11. Does he/she appropriately discuss local/regional evens such as marriages, disasters, politics?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q11', 'calculation': "if(${easi_raw11}='' or number(${easi_raw11})>=97, 0, if(number(${easi_raw11})=1, 0, 1))"},

        {'type': 'select_one easi_yesno', 'name': 'easi_raw12', 'label': '12. Does he/she ever lose his/her way in the village/area?', 'required': 'yes'},
        {'type': 'calculate', 'name': 'easi_q12', 'calculation': "if(${easi_raw12}='' or number(${easi_raw12})>=97, 0, number(${easi_raw12}))"},

        {'type': 'calculate', 'name': 'easi_total', 'calculation': "coalesce(${easi_q1},0)+coalesce(${easi_q2},0)+coalesce(${easi_q3},0)+coalesce(${easi_q4},0)+coalesce(${easi_q5},0)+coalesce(${easi_q6},0)+coalesce(${easi_q7},0)+coalesce(${easi_q8},0)+coalesce(${easi_q9},0)+coalesce(${easi_q10},0)+coalesce(${easi_q11},0)+coalesce(${easi_q12},0)"},
        {'type': 'note', 'name': '_NOTE_easi', 'label': '**NOTE:** A higher total score reflects greater overall disability. If the subject could not generally perform the activity, the reason for non-performance was ascertained. Responses clearly indicative of physical or mental problems or both were so coded. Ambiguous responses were probed and only when it was obvious that no further clarification was forthcoming were they left as such.'},
        {'type': 'note', 'name': 'display_score_easi', 'label': '**EASI (Everyday Abilities) Score:** ${easi_total} / 12\n\n**Interpretation:** A higher total score reflects greater overall disability.'},
        {'type': 'select_one easi_status', 'name': 'Survey_Status_easi', 'label': 'Survey Status', 'hint': 'Mark if the survey was completed or dropped out midway.', 'required': 'yes'},
        {'type': 'text', 'name': 'Reason_for_Dropout_easi', 'label': 'Reason for Dropout / Refusal:', 'hint': 'Please explain why the survey was stopped.', 'required': 'yes', 'relevant': "${Survey_Status_easi} = 'dropout___discontinued'"},
        {'type': 'end group'}
    ])

    choices = list(CHOICES_YES_NO)
    choices.extend([
        {'list_name': 'easi_yesno', 'name': '1', 'label': 'Yes'},
        {'list_name': 'easi_yesno', 'name': '0', 'label': 'No'},
        {'list_name': 'easi_yesno', 'name': '99', 'label': 'Not completed - Unable to understand / answer (comprehension barrier)'},
        {'list_name': 'easi_yesno', 'name': '98', 'label': 'Not completed - Unable to follow instructions (could not complete tasks as instructed)'},
        {'list_name': 'easi_yesno', 'name': '97', 'label': 'Not completed - Dropout (left early / discontinued)'},

        {'list_name': 'easi_status', 'name': 'completed', 'label': 'Completed'},
        {'list_name': 'easi_status', 'name': 'dropout___discontinued', 'label': 'Dropout / Discontinued'}
    ])
    save_workbook(form_id, title, survey, choices)

# ===========================================================================
# 6. GDS-15
# ===========================================================================
def build_gds15():
    form_id = 'gds15'
    title = 'Geriatric Depression Scale (GDS-15)'
    survey = make_inputs_and_details()
    survey.extend([
        {'type': 'begin group', 'name': 'grp_gds15_scale', 'label': 'Geriatric Depression Scale (GDS-15)'},
        {'type': 'note', 'name': 'gds15_instructions', 'label': '**Instructions:** Circle/select the answer that best describes how the participant felt over the past week.'},
        {'type': 'select_one yes_no', 'name': 'gds_q1', 'label': '1. Are you basically satisfied with your life?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q2', 'label': '2. Have you dropped many of your activities and interests?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q3', 'label': '3. Do you feel that your life is empty?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q4', 'label': '4. Do you often get bored?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q5', 'label': '5. Are you in good spirits most of the time?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q6', 'label': '6. Are you afraid that something bad is going to happen to you?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q7', 'label': '7. Do you feel happy most of the time?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q8', 'label': '8. Do you often feel helpless?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q9', 'label': '9. Do you prefer to stay at home, rather than going out and doing things?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q10', 'label': '10. Do you feel that you have more problems with memory than most?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q11', 'label': '11. Do you think it is wonderful to be alive now?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q12', 'label': '12. Do you feel worthless the way you are now?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q13', 'label': '13. Do you feel full of energy?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q14', 'label': '14. Do you feel that your situation is hopeless?', 'required': 'yes'},
        {'type': 'select_one yes_no', 'name': 'gds_q15', 'label': '15. Do you think that most people are better off than you are?', 'required': 'yes'},

        {'type': 'calculate', 'name': 'score_gds_q1', 'calculation': "if(${gds_q1}='no',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q2', 'calculation': "if(${gds_q2}='yes',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q3', 'calculation': "if(${gds_q3}='yes',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q4', 'calculation': "if(${gds_q4}='yes',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q5', 'calculation': "if(${gds_q5}='no',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q6', 'calculation': "if(${gds_q6}='yes',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q7', 'calculation': "if(${gds_q7}='no',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q8', 'calculation': "if(${gds_q8}='yes',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q9', 'calculation': "if(${gds_q9}='yes',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q10', 'calculation': "if(${gds_q10}='yes',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q11', 'calculation': "if(${gds_q11}='no',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q12', 'calculation': "if(${gds_q12}='yes',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q13', 'calculation': "if(${gds_q13}='no',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q14', 'calculation': "if(${gds_q14}='yes',1,0)"},
        {'type': 'calculate', 'name': 'score_gds_q15', 'calculation': "if(${gds_q15}='yes',1,0)"},
        {'type': 'calculate', 'name': 'gds15_total', 'calculation': "${score_gds_q1} + ${score_gds_q2} + ${score_gds_q3} + ${score_gds_q4} + ${score_gds_q5} + ${score_gds_q6} + ${score_gds_q7} + ${score_gds_q8} + ${score_gds_q9} + ${score_gds_q10} + ${score_gds_q11} + ${score_gds_q12} + ${score_gds_q13} + ${score_gds_q14} + ${score_gds_q15}"},
        {'type': 'note', 'name': 'display_score_gds15', 'label': '**Geriatric Depression Scale (GDS-15) Score:** ${gds15_total} / 15\n\n**Interpretation:** A score of 5 or more suggests depression (0-4 normal, 5-8 mild, 9-11 moderate, 12-15 severe).'},
        {'type': 'select_one gds15_status', 'name': 'Survey_Status_gds15', 'label': 'Survey Status', 'hint': 'Mark if the survey was completed or dropped out midway.', 'required': 'yes'},
        {'type': 'text', 'name': 'Reason_for_Dropout_gds15', 'label': 'Reason for Dropout / Refusal:', 'hint': 'Please explain why the survey was stopped.', 'required': 'yes', 'relevant': "${Survey_Status_gds15} = 'dropout___discontinued'"},
        {'type': 'end group'}
    ])

    choices = list(CHOICES_YES_NO)
    choices.extend([
        {'list_name': 'gds15_status', 'name': 'completed', 'label': 'Completed'},
        {'list_name': 'gds15_status', 'name': 'dropout___discontinued', 'label': 'Dropout / Discontinued'}
    ])
    save_workbook(form_id, title, survey, choices)

if __name__ == '__main__':
    build_cog5()
    build_mse()
    build_phq9()
    build_ad8()
    build_easi()
    build_gds15()
    print('All 6 scales generated successfully.')

