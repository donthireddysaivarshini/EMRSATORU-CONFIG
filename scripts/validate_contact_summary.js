// validate_contact_summary.js - runs the COMPILED contact summary from
// app_settings.json the way CHT does, and checks the participant access gate.
//
//     node scripts/validate_contact_summary.js
const fs = require('fs');
const path = require('path');

const settings = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'app_settings.json'), 'utf8'));
// CHT: new Function('contact', 'reports', 'lineage', 'uhcStats', 'cht', 'targetDoc', script)
const summarise = new Function('contact', 'reports', 'lineage', 'uhcStats', 'cht', 'targetDoc', settings.contact_summary);

let failures = 0;
const check = (name, cond) => {
  console.log(`  ${cond ? 'ok  ' : 'FAIL'} ${name}`);
  if (!cond) {
    failures++;
  }
};

const cht = ({ perms = [], userContactId = 'intern-1', participantAccessAllowed } = {}) => ({
  v1: {
    hasPermissions: (p) => perms.includes(p),
    hasAnyPermission: () => false,
    analytics: { getTargetDocs: () => [] },
    userContactId,
    participantAccessAllowed,
  },
});
const participant = (creator) => ({
  _id: 'p1', type: 'contact', contact_type: 'participant', name: 'Ravi', cr_no: 'CR-1',
  created_by_person_uuid: creator, reported_date: Date.now(),
});
const restricted = (s) => !!(s && s.context && s.context.access_restricted);
const run = (contact, api) => summarise(contact, [], [], {}, api, undefined);

check('compiles and runs', typeof run(participant('intern-1'), cht()) === 'object');
check('supervisor sees any participant', !restricted(run(participant('other'), cht({ perms: ['can_view_participants'] }))));
check('older app: own participant visible', !restricted(run(participant('intern-1'), cht())));
check('older app: other intern\'s participant restricted', restricted(run(participant('other'), cht())));
check('app says allowed (e.g. unlocked by exact search) -> visible',
  !restricted(run(participant('other'), cht({ participantAccessAllowed: true }))));
check('app says not allowed (e.g. window expired) -> restricted',
  restricted(run(participant('intern-1'), cht({ participantAccessAllowed: false }))));
const full = run(participant('intern-1'), cht());
check('visible summary keeps its fields and context', full.fields.length > 1 && full.context.cr_no === 'CR-1'
  && 'assessment_count' in full.context);
check('camps are never restricted', !restricted(run({ _id: 'c', type: 'contact', contact_type: 'screening_camp' }, cht())));

console.log(failures ? `\n${failures} FAILED` : '\nALL CONTACT SUMMARY CHECKS PASSED');
process.exit(failures ? 1 : 0);
