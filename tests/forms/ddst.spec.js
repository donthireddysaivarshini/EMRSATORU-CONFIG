/**
 * DDST-II, NIMS single-cut-off version (create_ddst_form.py): an item is shown
 * once the child is past its age, and a Fail/Refusal on a shown item is a delay.
 * Domain: 0 delays Normal, 1 Caution, 2+ Suspect.
 * Overall: Suspect if any domain is Suspect or 2+ domains are Caution.
 * These rules still need NIMS confirmation (SCORING_AUDIT.md X1, X2).
 */
const { expect } = require('chai');
const { newHarness, fill, ddstItems, valueOnOpen, today } = require('./helpers');

const harness = newHarness();
const TEST_DATE = '2026-01-07';
const DOB = '2024-01-01';          // 737 days = 24.2 months
const AGE_MONTHS = 737 / 30.4375;

const visible = ddstItems().filter(item => item.cutoff <= AGE_MONTHS);

const ddst = (fails = []) => [TEST_DATE, DOB, 'no']
  .concat(visible.map(item => fails.includes(item.name) ? 'F' : 'P'));

const firstIn = (domain, n = 0) => visible.filter(item => item.domain === domain)[n].name;

describe('DDST-II scoring', () => {
  before(async () => harness.start());
  after(async () => harness.stop());

  const score = async (fails) => {
    const { errors, fields } = await fill(harness, 'ddst', ddst(fails));
    expect(errors, JSON.stringify(errors)).to.be.empty;
    return fields;
  };

  it('shows only items the child is past', async () => {
    expect(visible.length).to.be.greaterThan(40);
    expect(ddstItems().length).to.equal(125);
  });

  it('all passed: every domain Normal, overall Normal', async () => {
    const f = await score();
    expect([f.d_ps_status, f.d_fm_status, f.d_lang_status, f.d_gm_status]).to.deep.equal(['Normal', 'Normal', 'Normal', 'Normal']);
    expect(f.overall).to.equal('Normal');
  });

  it('one delay: that domain is Caution, overall Normal (NIMS rule, to confirm)', async () => {
    const f = await score([firstIn('ps')]);
    expect(f.d_ps_delayed).to.equal('1');
    expect(f.d_ps_status).to.equal('Caution');
    expect(f.overall).to.equal('Normal');
  });

  it('two delays in one domain: Suspect', async () => {
    const f = await score([firstIn('lang'), firstIn('lang', 1)]);
    expect(f.d_lang_status).to.equal('Suspect');
    expect(f.overall).to.equal('Suspect');
  });

  it('one delay in two domains: Suspect', async () => {
    const f = await score([firstIn('fm'), firstIn('gm')]);
    expect(f.d_fm_status).to.equal('Caution');
    expect(f.d_gm_status).to.equal('Caution');
    expect(f.overall).to.equal('Suspect');
  });

  it('date of birth after the test date is refused (was a false "Normal")', async () => {
    const { errors } = await fill(harness, 'ddst', ['2026-01-01', '2026-02-01', 'no']);
    expect(errors.map(e => e.question).join()).to.contain('Child date of birth');
  });

  it('test date defaults to today', async () => {
    expect(await valueOnOpen(harness, 'ddst', '/data/g_meta/test_date')).to.equal(today());
  });
});
