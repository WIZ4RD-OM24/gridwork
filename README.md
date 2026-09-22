# Gridwork

A spreadsheet in plain HTML and JavaScript with no dependencies and no build step. It includes its own formula engine, multiple sheets, dates, live charts, and Excel-style editing. Run as a Claude artifact, it adds live multi-person editing and a "Build with Claude" button that writes new sheets from a sentence.

**Try it:** https://wiz4rd-om24.github.io/gridwork/

## Run it

```bash
python -m http.server 8765
```

Then open http://localhost:8765. Any static file server works. The page loads `engine.js`, `xlsx.js` and `sample.js` from the same folder. Your workbook saves in the browser; use **Export** for a file you can keep.

```bash
node test.js      # 239 checks on the formula engine and .xlsx reader/writer
```

## What it does

**Formulas.** The tokenizer, parser (Excel operator precedence) and evaluator are written from scratch in [`engine.js`](engine.js), with about 70 functions:
- **Math & statistics:** SUM, AVERAGE, MEDIAN, STDEV, ROUND…
- **Conditional totals:** SUMIFS, COUNTIFS, AVERAGEIF, with criteria like `">=100"`, `"app*"` and `"<>done"`
- **Lookups:** XLOOKUP, VLOOKUP, INDEX/MATCH
- **Text:** CONCAT, SUBSTITUTE, TEXT…
- **Logic:** IF and IFERROR only evaluate the branch they need.

Circular references show `#CYCLE!`. Every error explains itself in the status bar.

**Workbooks.** Multiple sheets can reference each other (`Budget!C12`, `'My Sheet'!A1:B9`). Renaming a sheet rewrites every formula that points at it, as Excel does. Right-click a tab to duplicate or delete it.

**Dates.** Dates use Excel-compatible serial numbers, so they can be added and subtracted. You can type `2026-10-01`, `1 Oct 2026` or `Oct 1, 2026`. Short dates like `3/4/2026` follow your locale (day first in India, the UK and most of Europe). Date functions: TODAY, NOW, DATE, EDATE, EOMONTH, DATEDIF, NETWORKDAYS, WEEKDAY and TEXT (`"dddd dd mmm yyyy"`).

**Editing like Excel:**
- Function autocomplete, with a hint that highlights the argument you're typing.
- While writing a formula, click or drag across cells to insert references; each reference gets its own color.
- A fill handle that continues series (1, 2 → 3, 4; Week 1 → Week 2; dates step by day). Formulas shift relative references and keep `$`-anchored ones fixed.
- Copy and paste with Excel and Google Sheets, undo/redo, Ctrl+Arrow to jump to the edge of your data, resizable columns.
- Long text spills into empty cells next to it.

**Rows, columns and views.** Right-click a cell or use the **Data** menu:
- Insert or delete rows and columns. Formulas, charts and rules that point past the change shift with it; references to deleted cells become `#REF!`.
- Sort a table by any column, or filter a column to hide rows (click ▾ in the header).
- Freeze panes, so header rows and label columns stay put while you scroll.
- Conditional formatting: color a range when a formula is true, e.g. `=C2>B2` to turn over-budget rows red.

**Charts.** Select a table and press Chart, then pick a bar, stacked bar, line, area, donut or scatter chart. It redraws as the data changes and shows a tooltip on hover.

**Files.** Real Excel files: `.xlsx` import and export keep every sheet, formula, number format, column width and frozen header. CSV/TSV import (the delimiter is detected), CSV export per sheet, and a `.gridwork.json` file that keeps every sheet with its formulas, formatting and charts.

## Inside Claude

When `index.html` is published as a Claude artifact with the `db`, `room`, `user`, `sample` and `downloads` capabilities, it lights up features that aren't available locally:

- **Live collaboration.** The workbook lives in the artifact's shared database. Each sheet row is one document (`rows/<sheet>.<row>`), and edits are merged cell by cell, so two people typing in the same row don't overwrite each other. Writes to a document go out one at a time, and a large paste is sent at most six rows at a time. Other people's selections show as colored boxes with their names, their avatars appear in the footer, and a sheet tab shows a dot when someone is on it.
- **Build with Claude.** Describe a sheet and Claude writes it into a new tab with live formulas. The cells fill in as the answer arrives, and Claude can reference your existing sheets.

Without those capabilities (opened as a plain file, for example), the page falls back to saving in the browser.

**Outside Claude** (GitHub Pages or a local server), Build with Claude asks for your own [Anthropic API key](https://console.anthropic.com/settings/keys). The key is kept in your browser's storage and sent only to `api.anthropic.com`, and requests are billed to your API account. Anyone who can run scripts on the same origin can read it, so use a key with a spending limit.

## Files

| File | What it is |
|---|---|
| `index.html` | The app: grid, editing, sheets, charts, saving and sync |
| `engine.js` | Formula engine; works in the browser and in Node |
| `xlsx.js` | `.xlsx` reader and writer, using the browser's built-in zip compression |
| `sample.js` | The example workbook shown on first open |
| `test.js` | Engine checks; run with `node test.js` |

## License

[MIT](LICENSE)
