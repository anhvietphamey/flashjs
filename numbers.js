// Pure number parsing/formatting. No Office or Excel dependencies.
                                                             
export const LOCALES                         = {
  en: { thousand: ',', decimal: '.' },
  vi: { thousand: '.', decimal: ',' },
};
                    
                     
                                                   
                     
                                                                                                                 
                                                                               

const esc = (s        ) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const DASH = /^[-–—]$/;

export function parseText(raw        , loc        )         {
  let s = raw.replace(/[\s\u00a0]+/g, ' ').trim();   // whitespace-only counts as empty
  if (!s) return { kind: 'empty' };
  if (DASH.test(s)) return { kind: 'number', value: 0, pct: false };
  let neg = false, pct = false;
  if (s.startsWith('(') && s.endsWith(')')) { neg = true; s = s.slice(1, -1).trim(); }
  if (s.endsWith('%')) { pct = true; s = s.slice(0, -1).trim(); }
  if (/^[-−–]/.test(s)) { neg = true; s = s.slice(1).trim(); }
  const re = new RegExp(`^(\\d{1,3}(${esc(loc.thousand)}\\d{3})+|\\d+)(${esc(loc.decimal)}\\d+)?$`);
  if (!re.test(s)) return { kind: 'text' };
  const n = Number(s.split(loc.thousand).join('').replace(loc.decimal, '.'));
  return { kind: 'number', value: neg ? -n : n, pct };
}

export function decimalsOf(raw        , loc        )         {
  const i = raw.lastIndexOf(loc.decimal);
  return i < 0 ? 0 : (/^\d+/.exec(raw.slice(i + 1))?.[0].length ?? 0);
}

/** Negative/zero conventions are inferred once per table from its existing text. */
export function scanTableStyle(texts          )             {
  const brackets = texts.some(t => /^\s*\(/.test(t));
  const minus = texts.some(t => /^\s*[-−]\s*\d/.test(t));
  return {
    negative: brackets || !minus ? 'brackets' : 'minus',   // financial statements default to brackets
    zero: texts.some(t => DASH.test(t.trim())) ? '-' : '0',
  };
}

export function styleFromWordText(text        , loc        , tbl            )           {
  return { decimals: decimalsOf(text, loc), percent: text.includes('%'), ...tbl };
}

/** Used when the Word cell is empty: take the style from the Excel cell's own number format. */
export function styleFromExcelFmt(z        , tbl            )           {
  const dec = /\.(0+)/.exec(z)?.[1].length ?? 0;
  const third = z.split(';')[2];
  return {
    decimals: dec,
    percent: z.includes('%'),
    negative: z.includes('(') ? 'brackets' : z.includes(';') ? 'minus' : tbl.negative,
    zero: third !== undefined ? (third.includes('-') ? '-' : '0') : tbl.zero,
  };
}

export function formatNumber(v        , st          , loc        )         {
  const x = st.percent ? v * 100 : v;
  const fixed = Math.abs(x).toFixed(st.decimals);
  const pct = st.percent ? '%' : '';
  if (Number(fixed) === 0) return st.zero === '-' ? '-' : fixed.replace('.', loc.decimal) + pct;
  const [i, f] = fixed.split('.');
  const body = i.replace(/\B(?=(\d{3})+(?!\d))/g, loc.thousand) + (f ? loc.decimal + f : '') + pct;
  return x < 0 ? (st.negative === 'brackets' ? `(${body})` : `-${body}`) : body;
}
