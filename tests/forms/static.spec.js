/**
 * Checks on the form files themselves (no browser).
 */
const fs = require('fs');
const path = require('path');
const { expect } = require('chai');

const formsDir = process.env.FORMS_DIR || path.join(__dirname, '..', '..', 'forms', 'app');
// FORM_IDS lets the forms panel check one uploaded form
const SCALES = process.env.FORM_IDS
  ? process.env.FORM_IDS.split(',')
  : fs.readdirSync(formsDir).filter(f => f.endsWith('.xml')).map(f => f.slice(0, -4));

describe('form files', () => {
  for (const form of SCALES) {
    const xml = fs.readFileSync(path.join(formsDir, `${form}.xml`), 'utf8');

    it(`${form}: no question the assessor fills in is overwritten by a formula`, () => {
      // A calculate on a visible question re-runs when a saved report is opened
      // for editing and silently replaces the answer (this lost MoCA fluency counts).
      // Allowed only as "fill while empty": if(. = '', <default>, .)
      const body = xml.slice(xml.indexOf('<h:body'));
      const questions = new Set([...body.matchAll(/<(?:input|select1|select|upload|range) [^>]*ref="([^"]+)"/g)].map(m => m[1]));
      const bad = [...xml.matchAll(/<bind nodeset="([^"]+)"([^>]*)\/>/g)]
        .filter(([, name, attrs]) => questions.has(name) && / calculate="/.test(attrs) && !/readonly="true\(\)"/.test(attrs))
        .filter(([, , attrs]) => !/ calculate="if\(\. = '', [^"]+, \.\)"/.test(attrs))
        .map(([, name]) => name);
      expect(bad).to.deep.equal([]);
    });

    it(`${form}: no first-load defaults (CHT never fires them)`, () => {
      // CHT always opens forms with a pre-filled instance, so odk-instance-first-load
      // never runs: a "default: today()" leaves the field blank. Use if(. = '', today(), .)
      expect(xml).not.to.contain('odk-instance-first-load');
    });

    it(`${form}: a saved true/false field is never used directly as a condition`, () => {
      // calculated comparisons are saved as the text "true"/"false"; in a later
      // formula any non-empty field counts as true (this made every DDST "Suspect")
      const decode = s => s.replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
      const formulas = [...xml.matchAll(/<bind nodeset="([^"]+)"([^>]*)\/>/g)]
        .flatMap(([, name, attrs]) => [...attrs.matchAll(/(?:calculate|relevant|constraint)="([^"]*)"/g)].map(m => [name, decode(m[1]).trim()]));
      const flags = formulas
        .filter(([, f]) => !/^(if|concat|floor|round|int|min|max|sum|count|coalesce|number|string|date|today|now|substr|selected-at|format-date)\(|^['"]|^\.\.\//.test(f))
        .filter(([, f]) => /(>|<|=| or | and )/.test(f))
        .map(([name]) => name);
      const misuse = formulas.filter(([, f]) => flags.some(flag => {
        const n = flag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        return new RegExp(`(^|if\\(|not\\(| or | and )\\s*${n}\\s*($|,|\\)| or | and )`).test(f);
      }));
      expect(misuse).to.deep.equal([]);
    });
  }
});
