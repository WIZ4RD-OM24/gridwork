// The example workbook: shown on first open of a local copy, and used to seed the shared artifact.
(function (root) {
'use strict';
const budget = {
  A1: 'Household budget · September 2026',
  A2: 'Example data. Change any Actual and watch Left, Used, the totals, the chart and the Bills sheet follow.',
  A4: 'Category', B4: 'Budget', C4: 'Actual', D4: 'Left', E4: 'Used', F4: 'Status',
  A12: 'Total', B12: '=SUM(B5:B11)', C12: '=SUM(C5:C11)', D12: '=B12-C12', E12: '=C12/B12', F12: '=COUNTIF(F5:F11,"Over")&" over budget"',
  A14: 'Income', B14: '3400',
  A15: 'Savings rate', B15: '=C11/B14',
  A16: 'Left after spending', B16: '=B14-C12',
  A17: 'Biggest overspend', B17: '=IF(MIN(D5:D11)<0,INDEX(A5:A11,MATCH(MIN(D5:D11),D5:D11,0)),"None")',
};
[['Rent', 1450, 1450], ['Groceries', 520, 583.4], ['Utilities', 180, 162.75], ['Transport', 140, 96.2],
 ['Dining out', 160, 214.9], ['Subscriptions', 45, 61.97], ['Savings', 600, 600]].forEach(([name, plan, actual], i) => {
  const r = i + 5;
  Object.assign(budget, {
    ['A' + r]: name, ['B' + r]: String(plan), ['C' + r]: String(actual),
    ['D' + r]: `=B${r}-C${r}`, ['E' + r]: `=C${r}/B${r}`, ['F' + r]: `=IF(C${r}>B${r},"Over","OK")`,
  });
});

const bills = {
  A1: 'Bills · what’s due next',
  A2: 'Due dates count from today, so this sheet stays current. Rent pulls its amount from the Budget sheet.',
  A4: 'Bill', B4: 'Amount', C4: 'Due', D4: 'Days left', E4: 'Status',
  A11: 'Total', B11: '=SUM(B5:B9)',
  A12: 'Most urgent', B12: '=INDEX(A5:A9,MATCH(MIN(C5:C9),C5:C9,0))', C12: '=MIN(C5:C9)',
  A13: 'Share of spending', B13: '=B11/Budget!C12',
  A15: 'Lease ends', B15: '2027-03-31', C15: '=IFERROR(DATEDIF(TODAY(),B15,"M")&" months to go","Ended")',
  A16: 'Month ends', B16: '=EOMONTH(TODAY(),0)', C16: '=TEXT(B16,"dddd")',
};
[['Rent', '=Budget!C5', '=EOMONTH(TODAY(),0)+1'], ['Electricity', '84.3', '=TODAY()+6'], ['Internet', '49.99', '=TODAY()+13'],
 ['Phone', '29', '=TODAY()-2'], ['Insurance', '212', '=EDATE(TODAY(),2)']].forEach(([name, amount, due], i) => {
  const r = i + 5;
  Object.assign(bills, {
    ['A' + r]: name, ['B' + r]: amount, ['C' + r]: due, ['D' + r]: `=C${r}-TODAY()`,
    ['E' + r]: `=IF(D${r}<0,"Overdue",IF(D${r}<=7,"Due this week","Later"))`,
  });
});

const SAMPLE = {
  sheets: [
    {
      id: 'budget', name: 'Budget', cells: budget, widths: { A: 150, F: 140 }, chart: { range: 'A4:C11', type: 'bar' },
      freeze: { rows: 4, cols: 0 }, filter: { range: 'A4:F11', cols: [] },
      cf: [{ range: 'C5:C11', kind: 'formula', a: '=C5>B5', style: 'red' }, { range: 'F5:F11', kind: 'eq', a: 'OK', style: 'green' }],
      formats: [['A1', { b: 1 }], ['A2', { i: 1 }], ['A4:F4', { b: 1 }], ['A12:F12', { b: 1 }], ['B5:D12', { f: 'num' }],
        ['E5:E12', { f: 'pct' }], ['B14', { f: 'num' }], ['B15', { f: 'pct' }], ['B16', { f: 'num' }]],
    },
    {
      id: 'bills', name: 'Bills', cells: bills, widths: { A: 140, B: 120, C: 120, E: 130 }, chart: null,
      freeze: { rows: 4, cols: 0 }, filter: { range: 'A4:E9', cols: [] },
      cf: [{ range: 'E5:E9', kind: 'contains', a: 'Overdue', style: 'red' }, { range: 'E5:E9', kind: 'contains', a: 'this week', style: 'amber' },
        { range: 'D5:D9', kind: 'scale', style: 'redgreen' }],
      formats: [['A1', { b: 1 }], ['A2', { i: 1 }], ['A4:E4', { b: 1 }], ['A11:B11', { b: 1 }], ['B5:B11', { f: 'num' }],
        ['C5:C9', { f: 'date' }], ['C12', { f: 'date' }], ['B13', { f: 'pct' }], ['B15', { f: 'date' }], ['B16', { f: 'date' }]],
    },
  ],
};
if (typeof module === 'object' && module.exports) module.exports = SAMPLE; else root.GRIDWORK_SAMPLE = SAMPLE;
})(this);
