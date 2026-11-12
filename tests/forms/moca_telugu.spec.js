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

  const delayed = o.delayed || {};
  const absCount = o.abstractionCount ?? 2;
  const absAnswers = Array.from({ length: 2 }, (_, i) => (i < absCount ? 'yes' : 'no'));
  const orientCount = o.orientationCount ?? 6;
  const orientAnswers = Array.from({ length: 6 }, (_, i) => (i < orientCount ? 'yes' : 'no'));

  return [
    'yes',                                // consent_reconfirmed
    String(o.edu ?? 15),                  // education_years
    ...repeat('yes', 5),                  // visuospatial (5)
    ...repeat('yes', 3),                  // naming (3)
    ...repeat('yes', 5),                  // mem_trial1
    ...repeat('yes', 5),                  // mem_trial2
    'yes',                                // digit_span_fwd
    'yes',                                // digit_span_bwd
    'yes',                                // vigilance
    String(o.serial7 ?? 3),               // serial_7s (0..3)
    'yes',                                // sentence_rep_1
    'yes',                                // sentence_rep_2
    String(o.fluencyCount ?? 12),         // fluency_count ('ప')
    ...absAnswers,                        // abstraction (2)
    ...delayedWord(...(delayed.kannu || [])),
    ...delayedWord(...(delayed.cheera || [])),
    ...delayedWord(...(delayed.gudi || [])),
    ...delayedWord(...(delayed.gulabi || [])),
    ...delayedWord(...(delayed.neelam || [])),
    ...orientAnswers,                     // orientation (6)
    'Dr. Tester',                         // administered_by
  ];
}

describe('MoCA Version 7.1 Telugu scoring', function () {
  this.timeout(120000);
  before(async () => harness.start());
  after(async () => harness.stop());

  const run = async (opts) => {
    const { errors, fields } = await fill(harness, 'moca_telugu', buildAnswers(opts));
    expect(errors, JSON.stringify(errors)).to.be.empty;
    return fields;
  };

  it('perfect score: 30/30, normal', async () => {
    const f = await run();
    expect(f.moca_te_visuospatial_total).to.equal('5');
    expect(f.moca_te_naming_total).to.equal('3');
    expect(f.moca_te_attention_total).to.equal('6');
    expect(f.moca_te_language_total).to.equal('3');
    expect(f.moca_te_abstraction_total).to.equal('2');
    expect(f.moca_te_memory_total).to.equal('5');
    expect(f.moca_te_orientation_total).to.equal('6');
    expect(f.moca_te_raw_total).to.equal('30');
    expect(f.moca_te_education_adj).to.equal('0');
    expect(f.moca_te_total_score).to.equal('30');
    expect(f.moca_te_interpretation).to.equal('normal');
  });

  it('fluency ("ప") threshold: 11->1, 10->0', async () => {
    expect((await run({ fluencyCount: 11 })).moca_te_language_total).to.equal('3');
    expect((await run({ fluencyCount: 10 })).moca_te_language_total).to.equal('2');
  });

  it('delayed recall: only uncued recall counts, cues score 0', async () => {
    const f = await run({
      delayed: {
        kannu: ['no', 'yes'],
        cheera: ['no', 'no', 'yes'],
        gudi: ['no', 'no', 'no'],
      },
    });
    expect(f.moca_te_memory_total).to.equal('2');
    expect(f.moca_te_total_score).to.equal('27');
  });

  it('education <= 12 yrs adds +1 (capped at 30) and >= 26 cutoff', async () => {
    let f = await run({ edu: 10, serial7: 0, abstractionCount: 0 });
    expect(f.moca_te_raw_total).to.equal('25');
    expect(f.moca_te_education_adj).to.equal('1');
    expect(f.moca_te_total_score).to.equal('26');
    expect(f.moca_te_interpretation).to.equal('normal');

    f = await run({ edu: 13, serial7: 0, abstractionCount: 0 });
    expect(f.moca_te_education_adj).to.equal('0');
    expect(f.moca_te_total_score).to.equal('25');
    expect(f.moca_te_interpretation).to.equal('impaired');

    f = await run({ edu: 8 });
    expect(f.moca_te_total_score).to.equal('30');
  });

  it('no consent: empty total score, interpretation not assessed (no consent)', async () => {
    await harness.clear();
    const { errors, report } = await harness.fillForm('moca_telugu', ['no', 'Dr. Tester']);
    expect(errors, JSON.stringify(errors)).to.be.empty;
    expect(report.fields.moca_te_total_score).to.equal('');
    expect(report.fields.moca_te_interpretation).to.equal('not assessed (no consent)');
  });
});

