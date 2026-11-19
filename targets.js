/**
 * satoru-config - targets
 *
 * This is the intern dashboard. Targets are computed on the device from the
 * documents that device has replicated, so an intern assigned to their own
 * screening_camp sees their own numbers and nobody else's. If two interns
 * share a camp, their numbers merge - see IMPLEMENTATION-GUIDE.md.
 *
 * `screened_by` is recorded on every report regardless, so a supervisor can
 * always break the numbers down by intern from an export.
 */

const {
  FORMS,
  CONTACT_TYPES,
  getField,
  num,
  isReportDoneByCurrentUser,
  isParticipantDoneByCurrentUser,
} = require('./nools-extras');

const consented = (contact, report) =>
  getField(report, 'g_consent.consent_confirmed') === 'yes';

module.exports = [

  // 1. Registered Participants (Intern / Admin Scoped)
  {
    id: 'participants-registered',
    type: 'count',
    icon: 'icon-person',
    goal: -1,
    translation_key: 'targets.registered.title',
    subtitle_translation_key: 'targets.this_month.subtitle',
    appliesTo: 'contacts',
    appliesToType: [CONTACT_TYPES.PARTICIPANT],
    appliesIf: (contact) => isParticipantDoneByCurrentUser(contact),
    idType: 'contact',
    date: 'reported',
  },

  // 2. Assessment Visits Completed (Form: assessment_visit / stage2_assessment)
  {
    id: 'stage2-completed',
    type: 'count',
    icon: 'icon-healthcare-assessment',
    goal: -1,
    translation_key: 'targets.stage2.title',
    subtitle_translation_key: 'targets.this_month.subtitle',
    appliesTo: 'reports',
    appliesToType: [FORMS.ASSESSMENT, FORMS.STAGE2],
    appliesIf: (contact, report) => isReportDoneByCurrentUser(report, contact),
    date: 'reported',
  },

  // 3. Memory Screenings Completed (Form: cog5 / stage1_screening)
  {
    id: 'stage1-completed',
    type: 'count',
    icon: 'icon-healthcare-assessment',
    goal: -1,
    translation_key: 'targets.stage1.title',
    subtitle_translation_key: 'targets.this_month.subtitle',
    appliesTo: 'reports',
    appliesToType: [FORMS.COG5, FORMS.STAGE1],
    appliesIf: (contact, report) =>
      isReportDoneByCurrentUser(report, contact) &&
      (report.form === FORMS.COG5 || consented(contact, report)),
    date: 'reported',
  },

  // 4. Clinical Scale Assessments Completed (MoCA, DDST-II, DST, VSMS)
  {
    id: 'clinical-scales-completed',
    type: 'count',
    icon: 'icon-healthcare-assessment',
    goal: -1,
    translation_key: 'targets.scales.title',
    subtitle_translation_key: 'targets.this_month.subtitle',
    appliesTo: 'reports',
    appliesToType: [
      FORMS.MOCA,
      FORMS.MOCA_INTEGRATED,
      FORMS.MOCA_B,
      FORMS.MOCA_TELUGU,
      FORMS.DDST,
      FORMS.DST,
      FORMS.VINELAND,
      FORMS.VSMS,
    ],
    appliesIf: (contact, report) => isReportDoneByCurrentUser(report, contact),
    date: 'reported',
  },

  // 5. Positive / Clinical Concern Identified (CoG5 Triage Flag / MoCA Below Normal / Clinical Concern)
  {
    id: 'cognitive-flag-count',
    type: 'count',
    icon: 'icon-healthcare-diagnosis',
    goal: -1,
    translation_key: 'targets.cognitive_flag.title',
    subtitle_translation_key: 'targets.this_month.subtitle',
    appliesTo: 'reports',
    appliesToType: [FORMS.COG5, FORMS.MOCA, FORMS.MOCA_INTEGRATED, FORMS.STAGE2],
    appliesIf: (contact, report) => {
      if (!isReportDoneByCurrentUser(report, contact)) {
        return false;
      }
      if (report.form === FORMS.COG5) {
        const triage = getField(report, 'grp_cog5_scale.FINAL_TRIAGE');
        return Boolean(triage && String(triage).includes('FLAG'));
      }
      if (report.form === FORMS.MOCA || report.form === FORMS.MOCA_INTEGRATED) {
        const interp = getField(report, 'interpretation') || getField(report, 'g_results.interpretation');
        return Boolean(interp && String(interp).toLowerCase().includes('below'));
      }
      return num(report, 'g_disposition.cognitive_flag') === 1;
    },
    date: 'reported',
  },
];
