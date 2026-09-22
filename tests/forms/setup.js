/**
 * cht-conf-test-harness 5.0.4 ships CHT 4.11's form engine and a 2021 Chromium.
 * Production runs CHT 5.3, so this swaps in the 5.3 engine kept in
 * engine/cht-core-5.3.0 (see README.md for how it was built) and lets the
 * harness use a modern Chromium from CHROMIUM_PATH. Safe to run repeatedly.
 */
const fs = require('fs');
const path = require('path');

const harnessDir = path.dirname(require.resolve('cht-conf-test-harness/package.json'));
const target = path.join(harnessDir, 'dist', 'cht-core-4-11');
const engine = path.join(__dirname, 'engine', 'cht-core-5.3.0');

const copy = (from, to) => {
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(from, to);
};

for (const f of ['main.js', 'polyfills.js', 'runtime.js', 'scripts.js', 'styles.css']) {
  copy(path.join(engine, f), path.join(target, f));
}
copy(path.join(engine, 'xsl', 'openrosa2html5form.xsl'), path.join(target, 'xsl', 'openrosa2html5form.xsl'));
for (const f of ['openrosa2html5form.xsl', 'openrosa2xmlmodel.xsl']) {
  copy(path.join(engine, 'enketo-transformer', 'xsl', f), path.join(target, 'enketo-transformer', 'xsl', f));
}

fs.writeFileSync(path.join(target, 'cht-form.js'), `// Satoru: load the CHT 5.3 form engine
(function () {
  var base = document.currentScript.src.replace(/cht-form\\.js.*$/, '');
  var css = document.createElement('link'); css.rel = 'stylesheet'; css.href = base + 'styles.css';
  document.head.appendChild(css);
  ['runtime.js', 'polyfills.js', 'scripts.js', 'main.js'].forEach(function (f) {
    var s = document.createElement('script'); s.src = base + f; s.async = false;
    document.head.appendChild(s);
  });
})();
`);

// use CHROMIUM_PATH directly: never download the bundled 2021 Chromium
const harnessJs = path.join(harnessDir, 'src', 'harness.js');
const original = [
  'const chromiumResolver = await PuppeteerChromiumResolver({ silent: !this.options.verbose });',
  '    this.options.executablePath = chromiumResolver.executablePath;',
].join('\n');
const patched = [
  "const chromiumResolver = { puppeteer: require('puppeteer-core'), executablePath: process.env.CHROMIUM_PATH };",
  '    this.options.executablePath = chromiumResolver.executablePath;',
].join('\n');
const src = fs.readFileSync(harnessJs, 'utf8');
if (!src.includes(patched)) {
  if (!src.includes(original)) {
    throw new Error('cht-conf-test-harness changed: cannot set the Chromium path. Check the pinned version.');
  }
  fs.writeFileSync(harnessJs, src.replace(original, patched));
}
if (!process.env.CHROMIUM_PATH) {
  throw new Error('Set CHROMIUM_PATH to a Chromium 120+ binary, e.g. CHROMIUM_PATH=/usr/bin/chromium');
}
console.log('form tests: CHT 5.3 engine installed, Chromium =', process.env.CHROMIUM_PATH);
