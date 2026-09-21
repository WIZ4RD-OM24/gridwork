# Gridwork

A spreadsheet in plain HTML and JavaScript with no dependencies and no build step. It includes its own formula engine, multiple sheets, dates, live charts, and Excel-style editing. Run as a Claude artifact, it adds live multi-person editing and a "Build with Claude" button that writes new sheets from a sentence.

## Run it

```bash
python -m http.server 8765
```

Then open http://localhost:8765. Any static file server works. The page loads `engine.js` and `sample.js` from the same folder. Your workbook saves in the browser; use **Export** for a file you can keep.

```bash
node test.js      # 192 checks on the formula engine
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

**Charts.** Select a table and press Chart for a bar or line chart. It redraws as the data changes and shows a tooltip on hover.

**Files.** CSV/TSV import (the delimiter is detected), CSV export per sheet, and a `.gridwork.json` file that keeps every sheet with its formulas, formatting and charts.

## Inside Claude

When `index.html` is published as a Claude artifact with the `db`, `room`, `user`, `sample` and `downloads` capabilities, it lights up features that aren't available locally:

- **Live collaboration.** The workbook lives in the artifact's shared database. Each sheet row is one document (`rows/<sheet>.<row>`), and edits are merged cell by cell, so two people typing in the same row don't overwrite each other. Writes to a document go out one at a time, and a large paste is sent at most six rows at a time. Other people's selections show as colored boxes with their names, their avatars appear in the footer, and a sheet tab shows a dot when someone is on it.
- **Build with Claude.** Describe a sheet and Claude writes it into a new tab with live formulas. The cells fill in as the answer arrives, and Claude can reference your existing sheets.

Without those capabilities (opened as a plain file, for example), the page falls back to saving in the browser.

## Files

| File | What it is |
|---|---|
| `index.html` | The app: grid, editing, sheets, charts, saving and sync |
| `engine.js` | Formula engine; works in the browser and in Node |
| `sample.js` | The example workbook shown on first open |
| `test.js` | Engine checks; run with `node test.js` |
