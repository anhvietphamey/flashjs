import {                 decideCell } from '../core/decide.js';
import {              scanTableStyle } from '../core/numbers.js';
import { isTextColumn,                  } from '../core/columns.js';

const RED = '#FF0000';

/** Link: wrap the table at the cursor in a content control tagged with the Excel defined name. */
export async function linkSelectedTable(name        ) {
  await Word.run(async ctx => {
    const t = ctx.document.getSelection().parentTableOrNullObject;
    t.load('isNullObject'); await ctx.sync();
    if (t.isNullObject) throw new Error('Place the cursor inside a table first.');
    const cc = t.getRange().insertContentControl('RichText');
    cc.tag = name; cc.title = name; cc.cannotDelete = true;
    await ctx.sync();
  });
}

/** Sync every tagged table. Changes are tracked. */
export async function syncAll(data                            , loc        ,
                              log                     , cfg              = {}) {
  await Word.run(async ctx => {
    const doc = ctx.document;
    doc.load('changeTrackingMode');
    const ccs = doc.contentControls; ccs.load('items/tag');
    await ctx.sync();
    const prevMode = doc.changeTrackingMode;
    doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    try {
      for (const cc of ccs.items) {
        if (!cc.tag) continue;
        const grid = data.get(cc.tag);
        if (!grid) { log(`${cc.tag}: no such name in the workbook`); continue; }
        const table = cc.tables.getFirstOrNullObject();
        table.load('values,isNullObject'); await ctx.sync();
        if (table.isNullObject) { log(`${cc.tag}: no table inside the control`); continue; }
        const vals             = table.values;
        if (vals.length !== grid.length || vals[0].length !== grid[0].length) {
          log(`${cc.tag}: size differs (Word ${vals.length}×${vals[0].length}, Excel ${grid.length}×${grid[0].length}); skipped`);
          continue;
        }
        const tbl = scanTableStyle(vals.flat());
        const jobs                                                                 = [];
        grid.forEach((row, r) => row.forEach((x, c) => {
          if (isTextColumn(cfg, c)) return;                // text columns: not synced yet
          const d = decideCell(x, vals[r][c], loc, tbl);
          const cell = table.getCell(r, c);
          cell.body.font.load('highlightColor');
          jobs.push({ cell, text: d.write, flag: d.flag });
        }));
        await ctx.sync();
        let written = 0, flagged = 0;
        for (const { cell, text, flag } of jobs) {
          if (text !== null) { cell.body.paragraphs.getFirst().insertText(text, 'Replace'); written++; }
          const cur = cell.body.font.highlightColor;
          if (flag && cur !== RED) cell.body.font.highlightColor = RED;
          if (!flag && cur) cell.body.font.highlightColor = null                     ;
          if (flag) flagged++;
        }
        await ctx.sync();
        log(`${cc.tag}: ${written} updated, ${flagged} flagged`);
      }
    } finally {
      doc.changeTrackingMode = prevMode;
      await ctx.sync();
    }
  });
}
