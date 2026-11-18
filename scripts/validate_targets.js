// validate_targets.js - runs the COMPILED rules from app_settings.json
// the way CHT rules engine does, and verifies intern vs supervisor target emissions.
//
//     node scripts/validate_targets.js
const fs = require('fs');
const path = require('path');

const settings = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'app_settings.json'), 'utf8'));

// Try loading nootils from cht-core or fallback mock
let nootils;
try {
  nootils = require('../cht-core/node_modules/cht-nootils')({});
} catch {
  try {
    nootils = require('../../cht-core/node_modules/cht-nootils')({});
  } catch {
    nootils = {
      now: () => new Date(),
      addDate: (d, days) => new Date(new Date(d).getTime() + days * 86400000),
      isTimely: () => true,
    };
  }
}

const fn = new Function('c', 'Task', 'Target', 'Utils', 'user', 'cht', 'emit', settings.tasks.rules);

function runTargets(contactDoc, reports, userObj, chtObj) {
  const emissions = [];
  const emit = (type, obj) => { if (type === 'target') emissions.push(obj); };
  function Target(data) { Object.assign(this, data); }
  function Task(data) { Object.assign(this, data); }
  fn({ contact: contactDoc, reports: reports || [] }, Task, Target, nootils, userObj, chtObj, emit);
  return emissions;
}

function runTasks(contactDoc, reports, userObj, chtObj) {
  const tasks = [];
  const emit = (type, obj) => { if (type === 'task') tasks.push(obj); };
  function Target(data) { Object.assign(this, data); }
  function Task(data) { Object.assign(this, data); }
  fn({ contact: contactDoc, reports: reports || [] }, Task, Target, nootils, userObj, chtObj, emit);
  return tasks;
}

let failures = 0;
const check = (name, cond) => {
  console.log(`  ${cond ? 'ok  ' : 'FAIL'} ${name}`);
  if (!cond) {
    failures++;
  }
};

const intern1 = { _id: 'intern-1', name: 'Intern One' };
const intern2 = { _id: 'intern-2', name: 'Intern Two' };
const supervisor = { _id: 'sup-1', roles: ['supervisor'] };
const supervisorCht = { v1: { hasPermissions: (p) => p === 'can_browse_all_participants' || p === 'can_view_participants' } };

const pOwn = { _id: 'p1', type: 'contact', contact_type: 'participant', reported_date: Date.now(), created_by_person_uuid: 'intern-1' };
const pOther = { _id: 'p2', type: 'contact', contact_type: 'participant', reported_date: Date.now(), created_by_person_uuid: 'intern-2' };

const r1Own = { _id: 'r1', form: 'stage1_screening', reported_date: Date.now(), fields: { screened_by: 'intern-1', g_consent: { consent_confirmed: 'yes' } } };
const r1Other = { _id: 'r2', form: 'stage1_screening', reported_date: Date.now(), fields: { screened_by: 'intern-2', g_consent: { consent_confirmed: 'yes' } } };

// 1. Intern owns registration -> counted
const res1 = runTargets(pOwn, [], intern1, {});
check('intern sees their own registered participant', res1.some(t => t.type === 'participants-registered'));

// 2. Intern did not register and did not assess -> not counted
const res2 = runTargets(pOther, [], intern1, {});
check('intern does NOT see other intern participant without encounters', !res2.some(t => t.type === 'participants-registered'));

// 3. Intern assessed other intern\'s registered participant -> counted
const res3 = runTargets(pOther, [r1Own], intern1, {});
check('intern sees participant they assessed even if registered by other', res3.some(t => t.type === 'participants-registered'));
check('intern sees their own completed stage1', res3.some(t => t.type === 'stage1-completed'));

// 4. Other intern assessed -> not counted in intern 1\'s stage 1
const res4 = runTargets(pOther, [r1Other], intern1, {});
check('intern does NOT see other intern stage1', !res4.some(t => t.type === 'stage1-completed'));
check('intern does NOT see other intern participant with only other reports', !res4.some(t => t.type === 'participants-registered'));

// 5. Supervisor sees all participants
const res5 = runTargets(pOther, [], supervisor, supervisorCht);
check('supervisor sees any registered participant', res5.some(t => t.type === 'participants-registered'));

// 6. Supervisor sees all reports
const res6 = runTargets(pOther, [r1Other], supervisor, supervisorCht);
check('supervisor sees reports done by any intern', res6.some(t => t.type === 'stage1-completed'));

// 7. Tasks: Intern sees assessment task only for their own participant
const tOwn = runTasks(pOwn, [], intern1, {});
check('intern sees assessment visit task for their own participant',
      tOwn.some(t => t._id.includes('assessment-visit-due')));

// Verify modifyContent on the task
const taskDoc = tOwn.find(t => t._id.includes('assessment-visit-due'));
check('assessment visit task action passes source=task and contact',
      taskDoc && taskDoc.actions && taskDoc.actions[0] &&
      taskDoc.actions[0].content && taskDoc.actions[0].content.source === 'task' &&
      taskDoc.actions[0].content.source_id === 'p1');

// 8. Tasks: Intern does NOT see assessment task for other intern participant
const tOther = runTasks(pOther, [], intern1, {});
check('intern does NOT see assessment visit task for other intern participant',
      !tOther.some(t => t._id.includes('assessment-visit-due')));

console.log(failures ? `\n${failures} FAILED` : '\nALL TARGET CHECKS PASSED');
process.exit(failures ? 1 : 0);
