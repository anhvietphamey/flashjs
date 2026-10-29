// No npm needed: strips TypeScript types (Node 22.13+) and writes a static site to ./static
import { stripTypeScriptTypes } from 'node:module';
import fs from 'node:fs'; import path from 'node:path';
const XLSX_CDN = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs';
const out = 'static';
fs.rmSync(out, { recursive: true, force: true });
const walk = d => fs.readdirSync(d, { withFileTypes: true }).flatMap(e =>
  e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]);
for (const f of walk('src')) {
  const rel = path.relative('src', f), dest = path.join(out, rel.replace(/\.ts$/, '.js'));
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  if (f.endsWith('.ts')) {
    let js = stripTypeScriptTypes(fs.readFileSync(f, 'utf8'), { mode: 'strip' })
      .replace(/(from\s+'\.[^']*)\.ts'/g, "$1.js'")
      .replace(/from 'xlsx'/g, `from '${XLSX_CDN}'`);
    fs.writeFileSync(dest, js);
  } else fs.copyFileSync(f, dest);
}
let html = fs.readFileSync(path.join(out, 'taskpane', 'taskpane.html'), 'utf8')
  .replace('<script src="taskpane.js"></script>', '<script type="module" src="taskpane.js"></script>');
fs.writeFileSync(path.join(out, 'taskpane', 'taskpane.html'), html);
console.log('Done: ./static  (host it on any https site; taskpane at /taskpane/taskpane.html)');
