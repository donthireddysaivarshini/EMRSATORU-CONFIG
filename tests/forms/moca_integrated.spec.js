/**
 * Unified MOCA form tests:
 * - Table 8.2 instrument selection (Telugu 7.1, MoCA-B, MoCA 8.1 English)
 * - MoCA 8.1 scoring + Memory Index Score (MIS /15)
 * - MoCA-B strict PDF scoring
 * - MoCA 7.1 Telugu scoring + Band A education-confounded flag
 * - Mandatory History Questions + No-consent behavior
 */
const { expect } = require('chai');
const { newHarness, repeat, fill } = require('./helpers');

const harness = newHarness();

const preAssessment = ({
  yearsSchooling = '15',
  canReadWrite = 'yes',
  testLanguage = 'english',
  luq = testLanguage === 'telugu' ? 'telugu' : (testLanguage === 'english' ? 'english' : 'neither'),
  informantPresent = 'yes',
} = {}) => [].concat(
  ['yes'], // consent_reconfirmed
  [
    yearsSchooling,
    canReadWrite,
    'English, Telugu',
    luq,
    'yes',
    testLanguage,
    informantPresent,
  ],
  informantPresent === 'yes' ? ['Ramesh', 'daily'] : [],
  ['yes', 'yes', 'yes', 'yes', 'yes', '2', 'none'], // sensory + testing conditions
);

const mandatoryHistoryAllNo = () => ['patient'].concat(repeat('no', 10), ['patient']);

const moca81Answers = ({
  pre = {},
  visuospatial = repeat('yes', 5),
  naming = repeat('yes', 3),
  attention = ['yes', 'yes', 'yes', '3'],
  sentences = ['yes', 'yes'],
  fluency = '14',
  abstraction = repeat('yes', 2),
  delayed = repeat('uncued', 5),
  orientation = repeat('yes', 6),
  history = mandatoryHistoryAllNo(),
} = {}) => [].concat(
  preAssessment({ testLanguage: 'english', yearsSchooling: '15', canReadWrite: 'yes', ...pre }),
  visuospatial,
  naming,
  repeat('yes', 10),
  attention,
  sentences,
  [fluency],
  abstraction,
  delayed,
  orientation,
  history,
  ['Tester'],
);

const mocaBAnswers = ({
  pre = {},
  trailMaking = 'yes',
  fluencyCount = '14',
  orientation = repeat('yes', 6),
  calculationScore = '3',
  abstraction = repeat('yes', 3),
  delayed = repeat(['yes'], 5),
  visuoperception = repeat('yes', 10),
  naming = repeat('yes', 4),
  whiteErrors = '0',
  blackErrors = '0',
} = {}) => [].concat(
  preAssessment({ testLanguage: 'english', yearsSchooling: '3', canReadWrite: 'no', ...pre }),
  ['09:00', trailMaking],
  repeat('yes', 10),
  ['apple, mango, banana', fluencyCount],
  orientation,
  ['10+1+1+1', calculationScore],
  abstraction,
  ...delayed,
  visuoperception,
  naming,
  [whiteErrors, blackErrors, '09:12', '12', '0'],
  mandatoryHistoryAllNo(),
  ['Tester'],
);

const mocaTeluguAnswers = ({
  pre = {},
  visuospatial = repeat('yes', 5),
  naming = repeat('yes', 3),
  attention = ['yes', 'yes', 'yes', '3'],
  sentences = ['yes', 'yes'],
  fluency = '12',
  abstraction = repeat('yes', 2),
  delayed = repeat(['yes'], 5),
  orientation = repeat('yes', 6),
} = {}) => [].concat(
  preAssessment({ testLanguage: 'telugu', yearsSchooling: '10', canReadWrite: 'yes', ...pre }),
  visuospatial,
  naming,
  repeat('yes', 10),
  attention,
  sentences,
  [fluency],
  abstraction,
  ...delayed,
  orientation,
  mandatoryHistoryAllNo(),
  ['Tester'],
);

describe('MOCA (Integrated) form & scoring', () => {
  before(async () => harness.start());
  after(async () => harness.stop());
  beforeEach(async () => harness.clear());

  it('English + Band B/C opens MoCA 8.1: 30/30, MIS 15/15, normal', async () => {
    const res = await fill(harness, 'moca', moca81Answers());
    expect(res.errors).to.deep.equal([]);
    expect(res.fields.g_edu_lang_informant.education_band).to.equal('C');
    expect(res.fields.selected_instrument).to.equal('moca_8_1');
    expect(res.fields.education_confounded).to.equal('no');
    expect(res.fields.moca_total_score).to.equal('30');
    expect(res.fields.moca_mis_score).to.equal('15');
    expect(res.fields.moca_interpretation).to.equal('Normal (>= 26/30)');
  });

  it('MoCA 8.1: MIS weights uncued (x3), category (x2), MC (x1), while Delayed Recall counts uncued only; <=12 yrs adds +1', async () => {
    const res = await fill(harness, 'moca', moca81Answers({
      pre: { yearsSchooling: '10', canReadWrite: 'yes' }, // Band B, <=12 yrs (+1)
      delayed: ['uncued', 'category', 'multiple_choice', 'not_recalled', 'uncued'],
    }));
    expect(res.errors).to.deep.equal([]);
    expect(res.fields.g_edu_lang_informant.education_band).to.equal('B');
    expect(res.fields.selected_instrument).to.equal('moca_8_1');
    expect(res.fields.m8_memory_total).to.equal('2'); // only 2 uncued
    expect(res.fields.moca_mis_score).to.equal('9'); // 3 + 2 + 1 + 0 + 3
    expect(res.fields.m8_education_adj).to.equal('1');
    expect(res.fields.moca_total_score).to.equal('28'); // 27 + 1
    expect(res.fields.moca_interpretation).to.equal('Normal (>= 26/30)');
  });

  it('English + unable to read/write simple sentence -> Band A -> opens MoCA-B', async () => {
    const res = await fill(harness, 'moca', mocaBAnswers({
      pre: { yearsSchooling: '6', canReadWrite: 'no' }, // >4 yrs but illiterate => Band A
    }));
    expect(res.errors).to.deep.equal([]);
    expect(res.fields.g_edu_lang_informant.education_band).to.equal('A');
    expect(res.fields.selected_instrument).to.equal('moca_b');
    expect(res.fields.moca_total_score).to.equal('30'); // capped at 30
    expect(res.fields.moca_interpretation || '').to.equal('');
  });

  it('Telugu + Band A -> opens Telugu MoCA 7.1 and flags education_confounded = yes', async () => {
    const res = await fill(harness, 'moca', mocaTeluguAnswers({
      pre: { yearsSchooling: '3', canReadWrite: 'no' },
    }));
    expect(res.errors).to.deep.equal([]);
    expect(res.fields.g_edu_lang_informant.education_band).to.equal('A');
    expect(res.fields.selected_instrument).to.equal('moca_telugu');
    expect(res.fields.education_confounded).to.equal('yes');
    expect(res.fields.moca_total_score).to.equal('30');
    expect(res.fields.moca_interpretation).to.equal('Normal (>= 26/30)');
  });

  it('Mandatory history conditional follow-ups (Q1 and Q10) work on MoCA 8.1', async () => {
    const history = ['informant', 'yes', '18 months ago', 'no', 'no', 'no', 'no', 'no', 'no', 'no', 'no', 'yes', 'NIMS 2025', 'informant'];
    const res = await fill(harness, 'moca', moca81Answers({
      orientation: repeat('no', 6), // 24 -> impaired
      history,
    }));
    expect(res.errors).to.deep.equal([]);
    expect(res.fields.moca_total_score).to.equal('24');
    expect(res.fields.moca_interpretation).to.equal('Below normal (< 26/30)');
    expect(res.fields.g_mandatory_history.mh_asked_to).to.equal('informant');
    expect(res.fields.g_mandatory_history.mh_answered_by).to.equal('informant');
    expect(res.fields.g_mandatory_history.mh_q1_duration).to.equal('18 months ago');
    expect(res.fields.g_mandatory_history.mh_q10_details).to.equal('NIMS 2025');
  });

  it('No consent: empty total score, interpretation not assessed (no consent)', async () => {
    const res = await fill(harness, 'moca', ['no', 'Tester']);
    expect(res.errors).to.deep.equal([]);
    expect(res.fields.moca_total_score || '').to.equal('');
    expect(res.fields.moca_interpretation).to.equal('Not assessed (no consent)');
  });
});

