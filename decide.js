import {                               parseText, formatNumber, styleFromWordText, styleFromExcelFmt } from './numbers.js';

                                                                         
/** write: new text for the Word cell (null = leave alone). flag: red highlight on/off. */
                                                                 

function normalize(x           )                         {
  if (typeof x.value === 'string') {
    const t = x.value.trim();
    return t === '' ? null : /^[-–—]$/.test(t) ? 0 : t;    // "-" in Excel means zero
  }
  return x.value;
}

export function decideCell(x           , wordText        , loc        , tbl            )           {
  const xv = normalize(x);
  const w = parseText(wordText, loc);
  if (xv === null) return { write: null, flag: w.kind !== 'empty' };   // Excel empty: keep Word text, flag it
  if (typeof xv === 'string') return { write: wordText.trim() === xv ? null : xv, flag: false };
  const style = w.kind === 'number' ? styleFromWordText(wordText, loc, tbl) : styleFromExcelFmt(x.fmt, tbl);
  const target = formatNumber(xv, style, loc);
  if (w.kind === 'number') {
    const t = parseText(target, loc);
    if (t.kind === 'number' && t.value === w.value) return { write: null, flag: false };  // already in sync
  }
  return { write: target, flag: false };
}
