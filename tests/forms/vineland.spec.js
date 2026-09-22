/**
 * VSMS: raw score = items answered yes; SQ = social age / chronological age x 100.
 * The raw score -> social age conversion is NOT verified (SCORING_AUDIT.md V1).
 */
const { expect } = require('chai');
const { newHarness, repeat, fill, valueOnOpen, today } = require('./helpers');

const harness = newHarness();

const vsms = ({ assessmentDate, dob, items, naReason }) =>
  ['yes', assessmentDate, dob, 'no'].concat(items, naReason ? [naReason] : []);

describe('VSMS scoring', () => {
  before(async () => harness.start());
  after(async () => harness.stop());

  const score = async (options) => {
    const { errors, fields } = await fill(harness, 'vineland', vsms(options));
    expect(errors, JSON.stringify(errors)).to.be.empty;
    return fields;
  };

  it('23.98 months sees the 0-2 year items; raw score counts yes answers', async () => {
    const items = repeat('yes', 34);
    items[3] = 'no';
    const f = await score({ assessmentDate: '2025-12-31', dob: '2024-01-01', items });
    expect(f.raw_score).to.equal('33');
    expect(f.total_no_items).to.equal('1');
    expect(f.score_status).to.equal('complete');
    expect([f.g_child.ca_years_display, f.g_child.ca_rem_months]).to.deep.equal(['1', '11']);
  });

  it('SQ = social age / chronological age x 100, rounded', async () => {
    const f = await score({ assessmentDate: '2025-12-31', dob: '2024-01-01', items: repeat('yes', 34) });
    const expected = Math.round(Number(f.social_age_months) / (730 / 30.4375) * 100);
    expect(f.social_quotient_rounded).to.equal(String(expected));
  });

  it('an item not assessed: incomplete', async () => {
    const items = repeat('yes', 34);
    items[0] = 'not_assessed';
    const f = await score({ assessmentDate: '2025-12-31', dob: '2024-01-01', items, naReason: 'not observed' });
    expect(f.score_status).to.equal('incomplete');
    expect(f.sq_interpretation).to.equal('Incomplete');
  });

  it.skip('raw score -> social age uses the VSMS manual table (waiting for NIMS)');

  it('newborn assessed on the day of birth: not interpreted as a delay', async () => {
    const f = await score({ assessmentDate: '2026-01-01', dob: '2026-01-01', items: repeat('yes', 17) });
    expect(f.social_quotient_rounded).to.equal('');
    expect(f.sq_interpretation).to.equal('Not interpretable (age under 1 day)');
  });

  it('date of birth after the assessment date is refused', async () => {
    const { errors } = await fill(harness, 'vineland', ['yes', '2026-01-01', '2026-02-01', 'no']);
    expect(errors.map(e => e.question).join()).to.contain('Child date of birth');
  });

  it('assessment date defaults to today', async () => {
    expect(await valueOnOpen(harness, 'vineland', '/data/g_child/assessment_date')).to.equal(today());
  });
});
