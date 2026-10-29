import { readWorkbook } from '../excel/reader.js';
import { LOCALES } from '../core/numbers.js';
import { linkSelectedTable, syncAll } from '../word/sync.js';
                                                   

const $ =                        (id        ) => document.getElementById(id)     ;
const log = (s        ) => { $('log').textContent += s + '\n'; };
let data = new Map                       ();

Office.onReady(() => {
  const saved = Office.context.document.settings.get('locale');
  if (saved) $                   ('locale').value = saved;

  $                   ('locale').onchange = e => {
    Office.context.document.settings.set('locale', (e.target                     ).value);
    Office.context.document.settings.saveAsync();
  };

  $                  ('file').onchange = async e => {
    const f = (e.target                    ).files?.[0]; if (!f) return;
    data = readWorkbook(await f.arrayBuffer());
    $                   ('names').innerHTML = [...data.keys()].map(n => `<option>${n}</option>`).join('');
    const none = data.size === 0;
    $                   ('names').disabled = $                   ('link').disabled = $                   ('sync').disabled = none;
    log(none ? 'No workbook-level named ranges in this file.' : `${data.size} named ranges found.`);
  };

  $('link').onclick = () => linkSelectedTable($                   ('names').value)
    .then(() => log('Table linked.')).catch(e => log(e.message));

  $('sync').onclick = () => syncAll(data, LOCALES[$                   ('locale').value], log)
    .catch(e => log(e.message));
});
