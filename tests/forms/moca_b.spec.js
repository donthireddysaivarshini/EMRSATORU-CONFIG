const { expect } = require('chai');
const { newHarness, repeat, fill } = require('./helpers');

const harness = newHarness();

function buildAnswers(o = {}) {
  const delayedWord = (free, cat, mc) => {
    const out = [free ?? 'yes'];
    if ((free ?? 'yes') === 'no') {
      out.push(cat ?? 'no');
      if ((cat ?? 'no') === 'no') out.push(mc ?? 'no');
    }
    return out;
  };

  const vpYesCount = o.vpCount ?? 10;
  const vpAnswers = Array.from({ length: 10 }, (_, i) => (i < vpYesCount ? 'yes' : 'no'));

  const namingCount = o.namingCount ?? 4;
  const namingAnswers = Array.from({ length: 4 }, (_, i) => {
    if (o.namingAnswers) return o.namingAnswers[i];
    return i < namingCount ? 'yes' : 'no';
  });

  const absCount = o.abstractionCount ?? 3;
  const absAnswers = Array.from({ length: 3 }, (_, i) => (i < absCount ? 'yes' : 'no'));

  const orientCount = o.orientationCount ?? 6;
  const orientAnswers = Array.from({ length: 6 }, (_, i) => (i < orientCount ? 'yes' : 'no'));

  const delayed = o.delayed || {};

  return [
    'yes',                                    // consent_reconfirmed
    String(o.edu ?? 12),                      // education_years
    o.illiterate ?? 'no',                     // is_illiterate
    o.startTime ?? '',                        // start_time (optional)
    o.trail ?? 'yes',                         // trail_making
    ...repeat('yes', 5),                      // mem_trial1
    ...repeat('yes', 5),                      // mem_trial2
    o.fluencyWords ?? '',                     // fluency_words (optional)
    String(o.fluencyCount ?? 14),             // fluency_count
    ...orientAnswers,                         // orientation (6)
    o.calcCombo ?? '',                        // calc_combinations (optional)
    String(o.calcScore ?? 3),                 // calculation_score_select
    ...absAnswers,                            // abstraction (3)
    ...delayedWord(...(delayed.rose || [])),
    ...delayedWord(...(delayed.chair || [])),
    ...delayedWord(...(delayed.hand || [])),
    ...delayedWord(...(delayed.blue || [])),
    ...delayedWord(...(delayed.spoon || [])),
    ...vpAnswers,                             // visuoperception (10)
    ...namingAnswers,                         // naming (4)
    String(o.whiteErrors ?? 0),               // attention_white_errors
    String(o.blackErrors ?? 0),               // attention_black_errors
    o.endTime ?? '',                          // end_time (optional)
    o.totalMin ?? '',                         // total_time_min (optional)
    o.totalSec ?? '',                         // total_time_sec (optional)
    'Dr. Tester',                             // administered_by
    // assessment_datetime auto-defaults via calculate="if(. = '', now(), .)"
  ];
}

describe('MoCA-B scoring', function () {
  this.timeout(120000);
  before(async () => harness.start());
  after(async () => harness.stop());

  const run = async (opts) => {
    const { errors, fields } = await fill(harness, 'moca_b', buildAnswers(opts));
    expect(errors, JSON.stringify(errors)).to.be.empty;
    return fields;
  };

  it('perfect score: 30/30, normal', async () => {
    const f = await run();
    expect(f.mocab_executive_total).to.equal('1');
    expect(f.mocab_fluency_total).to.equal('2');
    expect(f.mocab_orientation_total).to.equal('6');
    expect(f.mocab_calculation_total).to.equal('3');
    expect(f.mocab_abstraction_total).to.equal('3');
    expect(f.mocab_memory_total).to.equal('5');
    expect(f.mocab_visuoperception_total).to.equal('3');
    expect(f.mocab_naming_total).to.equal('4');
    expect(f.mocab_attention_white_score).to.equal('1');
    expect(f.mocab_attention_black_score).to.equal('2');
    expect(f.mocab_attention_total).to.equal('3');
    expect(f.mocab_raw_total).to.equal('30');
    expect(f.mocab_education_adj).to.equal('0');
    expect(f.mocab_literacy_adj).to.equal('0');
    expect(f.mocab_total_score).to.equal('30');
    expect(f.mocab_interpretation).to.equal('normal');
  });

  it('fluency thresholds: 13->2, 12->1, 8->1, 7->0', async () => {
    expect((await run({ fluencyCount: 13 })).mocab_fluency_total).to.equal('2');
    expect((await run({ fluencyCount: 12 })).mocab_fluency_total).to.equal('1');
    expect((await run({ fluencyCount: 8 })).mocab_fluency_total).to.equal('1');
    expect((await run({ fluencyCount: 7 })).mocab_fluency_total).to.equal('0');
  });

  it('calculation (0-3) and abstraction (0-3)', async () => {
    const f = await run({ calcScore: 2, abstractionCount: 1 });
    expect(f.mocab_calculation_total).to.equal('2');
    expect(f.mocab_abstraction_total).to.equal('1');
    expect(f.mocab_raw_total).to.equal('27');
  });

  it('delayed recall: only free recall counts, cues score 0', async () => {
    const f = await run({
      delayed: {
        rose: ['no', 'yes'],         // category cue -> 0
        chair: ['no', 'no', 'yes'],  // mc cue -> 0
        hand: ['no', 'no', 'no'],    // not recalled -> 0
      },
    });
    expect(f.mocab_memory_total).to.equal('2');
    expect(f.mocab_total_score).to.equal('27');
  });

  it('visuoperception thresholds: 9->3, 8->2, 6->2, 5->1, 4->1, 3->0', async () => {
    expect((await run({ vpCount: 9 })).mocab_visuoperception_total).to.equal('3');
    expect((await run({ vpCount: 8 })).mocab_visuoperception_total).to.equal('2');
    expect((await run({ vpCount: 6 })).mocab_visuoperception_total).to.equal('2');
    expect((await run({ vpCount: 5 })).mocab_visuoperception_total).to.equal('1');
    expect((await run({ vpCount: 4 })).mocab_visuoperception_total).to.equal('1');
    expect((await run({ vpCount: 3 })).mocab_visuoperception_total).to.equal('0');
  });

  it('naming handles yes, no, and dk correctly', async () => {
    const f = await run({ namingAnswers: ['yes', 'yes', 'dk', 'no'] });
    expect(f.mocab_naming_total).to.equal('2');
  });

  it('attention error thresholds: white (<=1->1, >=2->0) and black (<=2->2, 3->1, >=4->0)', async () => {
    let f = await run({ whiteErrors: 1, blackErrors: 2 });
    expect(f.mocab_attention_white_score).to.equal('1');
    expect(f.mocab_attention_black_score).to.equal('2');
    expect(f.mocab_attention_total).to.equal('3');

    f = await run({ whiteErrors: 2, blackErrors: 3 });
    expect(f.mocab_attention_white_score).to.equal('0');
    expect(f.mocab_attention_black_score).to.equal('1');
    expect(f.mocab_attention_total).to.equal('1');

    f = await run({ whiteErrors: 5, blackErrors: 4 });
    expect(f.mocab_attention_total).to.equal('0');
  });

  it('adjustments: <4 yrs (+1), illiterate (+1), both (+2), capped at 30', async () => {
    // 25 raw + 1 (<4 yrs) = 26 -> normal
    let f = await run({ edu: 3, illiterate: 'no', namingCount: 2, abstractionCount: 0 });
    expect(f.mocab_raw_total).to.equal('25');
    expect(f.mocab_education_adj).to.equal('1');
    expect(f.mocab_literacy_adj).to.equal('0');
    expect(f.mocab_total_score).to.equal('26');
    expect(f.mocab_interpretation).to.equal('normal');

    // 25 raw + 1 (<4 yrs) + 1 (illiterate) = 27
    f = await run({ edu: 0, illiterate: 'yes', namingCount: 2, abstractionCount: 0 });
    expect(f.mocab_education_adj).to.equal('1');
    expect(f.mocab_literacy_adj).to.equal('1');
    expect(f.mocab_total_score).to.equal('27');

    // edu=4 -> no edu bonus
    f = await run({ edu: 4, illiterate: 'no', namingCount: 2, abstractionCount: 0 });
    expect(f.mocab_education_adj).to.equal('0');
    expect(f.mocab_total_score).to.equal('25');
    expect(f.mocab_interpretation).to.equal('impaired');

    // Capped at 30
    f = await run({ edu: 0, illiterate: 'yes' });
    expect(f.mocab_raw_total).to.equal('30');
    expect(f.mocab_total_score).to.equal('30');
  });

  it('no consent: empty total score, interpretation not assessed (no consent)', async () => {
    await harness.clear();
    const { errors, report } = await harness.fillForm('moca_b', ['no', 'Dr. Tester']);
    expect(errors, JSON.stringify(errors)).to.be.empty;
    expect(report.fields.mocab_total_score).to.equal('');
    expect(report.fields.mocab_interpretation).to.equal('not assessed (no consent)');
  });

  it('fluency_count is stored and scored without being overwritten', async () => {
    const f1 = await run({ fluencyCount: 11 });
    expect(f1.g_fluency.fluency_count).to.equal('11');
    expect(f1.mocab_fluency_total).to.equal('1');
  });
});

