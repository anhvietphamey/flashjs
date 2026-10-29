import * as XLSX from 'https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs';
                                                   

/** Maps each workbook-scoped defined name to its grid of cells (cached values + number formats). */
export function readWorkbook(buf             )                             {
  const wb = XLSX.read(buf, { type: 'array', cellNF: true });
  const out = new Map                       ();
  for (const n of wb.Workbook?.Names ?? []) {
    if (n.Sheet != null || n.Name.startsWith('_xlnm')) continue;   // workbook-scoped only
    const m = /^'?(.+?)'?!(.+)$/.exec(n.Ref);
    const ws = m && wb.Sheets[m[1].replace(/''/g, "'")];
    if (!m || !ws) continue;
    const rg = XLSX.utils.decode_range(m[2].replace(/\$/g, ''));
    const grid                = [];
    for (let r = rg.s.r; r <= rg.e.r; r++) {
      const row              = [];
      for (let c = rg.s.c; c <= rg.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        row.push({ value: cell && (cell.t === 'n' || cell.t === 's') ? cell.v                    : null, fmt: cell?.z ?? 'General' });
      }
      grid.push(row);
    }
    out.set(n.Name, grid);
  }
  return out;
}
