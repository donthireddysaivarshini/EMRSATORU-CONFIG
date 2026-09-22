const path = require('path');
const Harness = require('cht-conf-test-harness');

const repo = path.join(__dirname, '..', '..');
// the forms panel points this at an uploaded form before publishing it
const formsDir = process.env.FORMS_DIR || path.join(repo, 'forms', 'app');

const newHarness = () => new Harness({
  directory: __dirname,
  appXFormFolderPath: formsDir,
  // form tests need no app settings (tasks/targets are not tested here)
  appSettingsPath: path.join(__dirname, 'app_settings.json'),
  headless: true,
  args: process.env.CHROMIUM_NO_SANDBOX ? ['--no-sandbox'] : [],
});

const repeat = (value, n) => Array(n).fill(value);

// fill a form (all questions are on one page) and return the saved fields
const fill = async (harness, form, answers) => {
  await harness.clear();
  const result = await harness.fillForm(form, answers);
  return { errors: result.errors, fields: result.report && result.report.fields };
};

// value of a question right after the form opens (before anything is filled in)
const valueOnOpen = async (harness, form, name) => {
  await harness.clear();
  await harness.loadForm(form);
  return harness.page.evaluate(n => document.querySelector(`[name="${n}"]`)?.value, name);
};

const today = () => new Date().toLocaleDateString('en-CA');

module.exports = { newHarness, repeat, fill, formsDir, valueOnOpen, today };

// DDST-II items in form order, with the age (months) from which each is shown
const ddstItems = () => {
  const fs = require('fs');
  const xml = fs.readFileSync(path.join(formsDir, 'ddst.xml'), 'utf8');
  const cutoffs = {};
  for (const m of xml.matchAll(/<bind nodeset="\/data\/g_[a-z]+\/([a-z]+_\d+)"[^>]*ca_months_dec {2}&gt;= ([0-9.]+)"/g)) {
    cutoffs[m[1]] = Number(m[2]);
  }
  const body = xml.slice(xml.indexOf('<h:body'));
  return [...body.matchAll(/<select1 ref="\/data\/g_[a-z]+\/([a-z]+_\d+)"/g)]
    .map(m => ({ name: m[1], domain: m[1].split('_')[0], cutoff: cutoffs[m[1]] }));
};

module.exports.ddstItems = ddstItems;
