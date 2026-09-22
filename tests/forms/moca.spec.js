/**
 * MoCA scoring, checked against the official MoCA instructions
 * (scales/moca/moca-instructions-english_2010.pdf).
 */
const { expect } = require('chai');
const { newHarness, repeat, fill } = require('./helpers');

const harness = newHarness();

// answers in the order the questions appear; override any section
const moca = ({
  educationYears = '15',
  visuospatial = repeat('yes', 5),
  naming = repeat('yes', 3),
  attention = ['yes', 'yes', 'yes', '3'],   // forward, backward, vigilance, serial 7s (0-3)
  sentences = ['yes', 'yes'],
  fluency = '14',
  abstraction = repeat('yes', 2),
  delayed = repeat(['yes', '', ''], 5),     // free recall, category cue, multiple choice
  orientation = repeat('yes', 6),
  assessmentTime,                           // blank = keep the automatic "now"
} = {}) => [].concat(
  ['yes', educationYears],
  visuospatial, naming,
  repeat('yes', 10),                        // learning trials: not scored
  attention, sentences, [fluency], abstraction,
  ...delayed, orientation,
  ['Tester'],
  assessmentTime ? [assessmentTime] : [],
);

describe('MoCA scoring', () => {
  before(async () => harness.start());
  after(async () => harness.stop());
  beforeEach(async () => harness.clear());

  const score = async (options) => {
    const { errors, fields } = await fill(harness, 'moca_assessment', moca(options));
    expect(errors, JSON.stringify(errors)).to.be.empty;
    return fields;
  };

  it('everything correct with 15 years of education scores 30, normal', async () => {
    const f = await score();
    expect(f.moca_visuospatial_total).to.equal('5');
    expect(f.moca_naming_total).to.equal('3');
    expect(f.moca_attention_total).to.equal('6');
    expect(f.moca_language_total).to.equal('3');
    expect(f.moca_abstraction_total).to.equal('2');
    expect(f.moca_memory_total).to.equal('5');
    expect(f.moca_orientation_total).to.equal('6');
    expect(f.moca_raw_total).to.equal('30');
    expect(f.moca_education_adj).to.equal('0');
    expect(f.moca_total_score).to.equal('30');
    expect(f.moca_interpretation).to.equal('normal');
  });

  it('fluency: 11 words earns the point, 10 does not', async () => {
    expect((await score({ fluency: '11' })).moca_language_total).to.equal('3');
    expect((await score({ fluency: '10' })).moca_language_total).to.equal('2');
  });

  it('keeps the typed fluency count (no formula may overwrite it)', async () => {
    const f = await score({ fluency: '17' });
    expect(f.g_language.fluency_count).to.equal('17');
  });

  it('education: 12 years or fewer adds 1 point, 13 does not', async () => {
    expect((await score({ educationYears: '12', orientation: ['no', 'yes', 'yes', 'yes', 'yes', 'yes'] })).moca_education_adj).to.equal('1');
    expect((await score({ educationYears: '0', orientation: ['no', 'yes', 'yes', 'yes', 'yes', 'yes'] })).moca_education_adj).to.equal('1');
    expect((await score({ educationYears: '13' })).moca_education_adj).to.equal('0');
  });

  it('cut-off: 25 + education point = 26 normal; 25 with 13+ years = impaired', async () => {
    const low = { fluency: '5', abstraction: ['no', 'yes'], orientation: ['no', 'no', 'yes', 'yes', 'yes', 'yes'], naming: ['yes', 'yes', 'no'] };
    let f = await score({ ...low, educationYears: '8' });
    expect(f.moca_raw_total).to.equal('25');
    expect(f.moca_total_score).to.equal('26');
    expect(f.moca_interpretation).to.equal('normal');
    f = await score({ ...low, educationYears: '16' });
    expect(f.moca_total_score).to.equal('25');
    expect(f.moca_interpretation).to.equal('impaired');
  });

  it('total is capped at 30', async () => {
    const f = await score({ educationYears: '5' });
    expect(f.moca_raw_total).to.equal('30');
    expect(f.moca_total_score).to.equal('30');
  });

  it('serial 7s score is added as 0-3 points', async () => {
    expect((await score({ attention: ['yes', 'yes', 'yes', '0'] })).moca_attention_total).to.equal('3');
    expect((await score({ attention: ['no', 'no', 'no', '2'] })).moca_attention_total).to.equal('2');
  });

  it('delayed recall: only words recalled without a cue count', async () => {
    const f = await score({ delayed: [['no', 'yes', ''], ['no', 'no', 'yes'], ['yes', '', ''], ['yes', '', ''], ['yes', '', '']] });
    expect(f.moca_memory_total).to.equal('3');
  });

  it('fluency count and years of education are required', async () => {
    let { errors } = await fill(harness, 'moca_assessment', moca({ fluency: '' }));
    expect(errors.map(e => e.question).join()).to.contain('Fluency count');
    ({ errors } = await fill(harness, 'moca_assessment', moca({ educationYears: '' })));
    expect(errors.map(e => e.question).join()).to.contain('Years of formal education');
  });

  it('assessment time defaults to now', async () => {
    const f = await score();
    
    const saved = new Date(f.g_admin.assessment_datetime);
    expect(Math.abs(Date.now() - saved.getTime())).to.be.below(10 * 60 * 1000);
  });

  it('assessment time can be set to when the paper test was given', async () => {
    const f = await score({ assessmentTime: '2026-09-01 10:30' });
    expect(f.g_admin.assessment_datetime).to.match(/^2026-09-01T10:30/);
  });

  it('no consent: no score and not "impaired"', async () => {
    const { errors, fields } = await fill(harness, 'moca_assessment', ['no', 'Tester']);
    expect(errors, JSON.stringify(errors)).to.be.empty;
    expect(fields.moca_total_score).to.equal('');
    expect(fields.moca_interpretation).to.equal('not assessed (no consent)');
  });
});
