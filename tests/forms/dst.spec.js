/**
 * DST: each passed item earns (band width / items in band) months.
 * DA = sum; DQ = DA / chronological age x 100, only when nothing is "not assessed".
 */
const { expect } = require('chai');
const { newHarness, repeat, fill, valueOnOpen, today } = require('./helpers');

const harness = newHarness();

// [first item, last item] of the bands a child sees, youngest first
const dst = ({ assessmentDate, dob, items, naReason }) =>
  ['yes', assessmentDate, dob, 'no'].concat(items, naReason ? [naReason] : []);

describe('DST scoring', () => {
  before(async () => harness.start());
  after(async () => harness.stop());

  const score = async (options) => {
    const { errors, fields } = await fill(harness, 'dst', dst(options));
    expect(errors, JSON.stringify(errors)).to.be.empty;
    return fields;
  };

  it('23.98 months, all 29 items up to 24 months passed: DA 24, DQ 100', async () => {
    const f = await score({ assessmentDate: '2025-12-31', dob: '2024-01-01', items: repeat('pass', 29) });
    expect(f.developmental_age_months).to.equal('24');
    expect(f.score_status).to.equal('complete');
    expect(f.developmental_quotient_rounded).to.equal('100');
    expect([f.g_child.ca_years_display, f.g_child.ca_rem_months]).to.deep.equal(['1', '11']);
  });

  it('two failed 18-24 month items lose 1.5 months each: DA 21, DQ 88', async () => {
    const items = repeat('pass', 29);
    items[25] = 'fail';
    items[26] = 'fail';
    const f = await score({ assessmentDate: '2025-12-31', dob: '2024-01-01', items });
    expect(f.developmental_age_months).to.equal('21');
    expect(f.developmental_quotient_rounded).to.equal('88');
  });

  it('exactly 2 years shows "2y 0m" and sees the 2-3 year band', async () => {
    const f = await score({ assessmentDate: '2026-01-01', dob: '2024-01-01', items: repeat('pass', 35) });
    expect([f.g_child.ca_years_display, f.g_child.ca_rem_months]).to.deep.equal(['2', '0']);
    expect(f.developmental_age_months).to.equal('36');
    expect(f.developmental_quotient_rounded).to.equal('150');
  });

  it('an item not assessed: incomplete, no DQ', async () => {
    const items = repeat('pass', 29);
    items[0] = 'not_assessed';
    const f = await score({ assessmentDate: '2025-12-31', dob: '2024-01-01', items, naReason: 'child asleep' });
    expect(f.score_status).to.equal('incomplete');
    expect(f.developmental_quotient_rounded).to.equal('');
  });

  it('newborn assessed on the day of birth: no DQ, not "NaN"', async () => {
    const f = await score({ assessmentDate: '2026-01-01', dob: '2026-01-01', items: repeat('pass', 7) });
    expect(f.developmental_quotient_rounded).to.equal('');
  });

  it('assessment date defaults to today', async () => {
    expect(await valueOnOpen(harness, 'dst', '/data/g_child/assessment_date')).to.equal(today());
  });

  it('date of birth after the assessment date is refused', async () => {
    const { errors } = await fill(harness, 'dst', ['yes', '2026-01-01', '2026-02-01', 'no']);
    expect(errors.map(e => e.question).join()).to.contain('Child date of birth');
  });

  it('assessment date in the future is refused', async () => {
    const { errors } = await fill(harness, 'dst', ['yes', '2099-01-01', '2098-06-01', 'no']);
    expect(errors.map(e => e.question).join()).to.contain('Assessment date');
  });

  it('exactly 15 years on the assessment date is accepted, 15 years + 1 day refused', async () => {
    let r = await fill(harness, 'dst', ['yes', '2026-01-01', '2011-01-01', 'no'].concat(repeat('pass', 88)));
    expect(r.errors.map(e => e.question).join()).not.to.contain('Child date of birth');
    r = await fill(harness, 'dst', ['yes', '2026-01-02', '2011-01-01', 'no']);
    expect(r.errors.map(e => e.question).join()).to.contain('Child date of birth');
  });
});
