/**
 * validate_exact_search.js - offline behavioural test for "disable partial
 * search" (cht-core webapp/src/ts/services/exact-search.service.ts).
 *
 * Transpiles the real service, runs it under Node against an in-memory fake
 * database, and asserts that with settings.search_exact_match on, a search
 * only returns a participant (or their reports) when the text is their full
 * name, full participant ID or full phone number.
 *
 * Run: node scripts/validate_exact_search.js   (exits non-zero on failure)
 */

const fs = require('fs');
const path = require('path');
const ts = require(path.join(__dirname, '..', '..', 'cht-core', 'node_modules', 'typescript'));

const CORE = path.join(__dirname, '..', '..', 'cht-core');
const SERVICE = path.join(CORE, 'webapp', 'src', 'ts', 'services', 'exact-search.service.ts');

let checks = 0;
const failures = [];
const check = (name, cond, extra) => {
  checks++;
  if (cond) {
    console.log(`  ok   ${name}`);
  } else {
    failures.push(name);
    console.log(`  FAIL ${name}${extra !== undefined ? '  ' + JSON.stringify(extra) : ''}`);
  }
};

// ---------------------------------------------------------------------------
// fake data: three participants who share a name prefix, plus their reports
// ---------------------------------------------------------------------------
const DOCS = [
  { _id: 'p-kamala', type: 'contact', contact_type: 'participant',
    name: 'Kamala Devi', cr_no: '79704', phone: '+91 90000 10001' },
  { _id: 'p-kamalakar', type: 'contact', contact_type: 'participant',
    name: 'Kamalakar Rao', cr_no: '79705', phone: '+91 90000 10002' },
  { _id: 'p-kavya', type: 'contact', contact_type: 'participant',
    name: 'Kavya', cr_no: '80001', phone: '+91 90000 10003' },

  { _id: 'r-kamala-dst', type: 'data_record', form: 'dst',
    fields: { patient_uuid: 'p-kamala', patient_id: '79704', patient_name: 'Kamala Devi' } },
  { _id: 'r-kamalakar-dst', type: 'data_record', form: 'dst',
    fields: { patient_uuid: 'p-kamalakar', patient_id: '79705', patient_name: 'Kamalakar Rao' } },
  // report whose subject is only linked by uuid - name must come from the contact
  { _id: 'r-kavya-moca', type: 'data_record', form: 'moca_assessment',
    fields: { patient_uuid: 'p-kavya' } },
];
const byId = Object.fromEntries(DOCS.map(d => [d._id, d]));
const dbService = {
  get: () => ({
    allDocs: async ({ keys }) => ({
      rows: keys.map(k => (byId[k] ? { id: k, doc: byId[k] } : { key: k, error: 'not_found' })),
    }),
  }),
};

// ---------------------------------------------------------------------------
function loadService() {
  const src = fs.readFileSync(SERVICE, 'utf8');
  const js = ts.transpileModule(src, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020,
      experimentalDecorators: true, emitDecoratorMetadata: false,
    },
  }).outputText;
  const module = { exports: {} };
  const requireStub = (name) => {
    if (name.startsWith('@angular/') || name.startsWith('@mm-services/')) {
      const noop = () => () => undefined;
      return new Proxy({}, { get: () => noop });
    }
    throw new Error(`unexpected import: ${name}`);
  };
  new Function('exports', 'require', 'module', js)(module.exports, requireStub, module);
  return new module.exports.ExactSearchService(dbService);
}

const ALL_CONTACTS = ['p-kamala', 'p-kamalakar', 'p-kavya'];
const ALL_REPORTS = ['r-kamala-dst', 'r-kamalakar-dst', 'r-kavya-moca'];
const ON = { search_exact_match: true };

(async () => {
  const svc = loadService();

  console.log('== when it applies ==');
  check('off when setting absent', !svc.isEnabled({}, 'Kamala Devi'));
  check('off when setting false', !svc.isEnabled({ search_exact_match: false }, 'Kamala Devi'));
  check('off with no search text (plain list view)', !svc.isEnabled(ON, ''));
  check('off for whitespace-only search', !svc.isEnabled(ON, '   '));
  check('off for field:value searches (already exact)', !svc.isEnabled(ON, 'patient_id:79704'));
  check('on for a normal search', svc.isEnabled(ON, 'Kamala Devi'));

  // the search library returns prefix matches; we narrow them
  const people = async q => (await svc.filter('contacts', ALL_CONTACTS, q)).sort();
  const reports = async q => (await svc.filter('reports', ALL_REPORTS, q)).sort();

  console.log('== People: partial text finds nothing ==');
  check('"Ka" -> nothing', (await people('Ka')).length === 0, await people('Ka'));
  check('"Kamala" -> nothing (not her full name)', (await people('Kamala')).length === 0, await people('Kamala'));
  check('"797" -> nothing (partial ID)', (await people('797')).length === 0, await people('797'));
  check('"10001" -> nothing (partial phone)', (await people('10001')).length === 0, await people('10001'));

  console.log('== People: complete values find exactly that person ==');
  check('full name', JSON.stringify(await people('Kamala Devi')) === '["p-kamala"]', await people('Kamala Devi'));
  check('full name, any case', JSON.stringify(await people('KAMALA devi')) === '["p-kamala"]');
  check('full name, extra spaces', JSON.stringify(await people('  Kamala   Devi ')) === '["p-kamala"]');
  check('full name does not also return Kamalakar', !(await people('Kamala Devi')).includes('p-kamalakar'));
  check('single-word full name', JSON.stringify(await people('Kavya')) === '["p-kavya"]', await people('Kavya'));
  check('full participant ID', JSON.stringify(await people('79704')) === '["p-kamala"]', await people('79704'));
  check('full phone as stored', JSON.stringify(await people('+91 90000 10001')) === '["p-kamala"]');
  check('full phone without country code', JSON.stringify(await people('9000010001')) === '["p-kamala"]', await people('9000010001'));
  check('full phone with dashes', JSON.stringify(await people('90000-10001')) === '["p-kamala"]');

  console.log('== Reports: same rule, matched on the participant ==');
  check('"Kamala" -> no reports', (await reports('Kamala')).length === 0, await reports('Kamala'));
  check('full name -> only her report', JSON.stringify(await reports('Kamala Devi')) === '["r-kamala-dst"]', await reports('Kamala Devi'));
  check('full ID -> only her report', JSON.stringify(await reports('79705')) === '["r-kamalakar-dst"]', await reports('79705'));
  check('subject known only by uuid, name from contact', JSON.stringify(await reports('Kavya')) === '["r-kavya-moca"]', await reports('Kavya'));
  check('subject known only by uuid, phone from contact', JSON.stringify(await reports('9000010003')) === '["r-kavya-moca"]', await reports('9000010003'));

  console.log('== robustness ==');
  check('empty id list', (await svc.filter('contacts', [], 'Kavya')).length === 0);
  check('ids that no longer exist are dropped', JSON.stringify(await svc.filter('contacts', ['gone', 'p-kavya'], 'Kavya')) === '["p-kavya"]');
  check('preserves the search library order', JSON.stringify(await svc.filter('contacts', ['p-kavya', 'p-kamala'], 'Kavya')) === '["p-kavya"]');

  console.log(`\n${checks} checks, ${failures.length} failures`);
  if (failures.length) {
    process.exit(1);
  }
  console.log('ALL EXACT-SEARCH CHECKS PASSED');
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
