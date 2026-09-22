// Engine self-check: `node test.js` — exits non-zero on the first failure.
const assert = require('assert');
const E = require('./engine.js');

function sheet(cells) {
  const s = new E.Sheet();
  for (const [k, v] of Object.entries(cells)) s.set(k, v);
  s.recalc();
  return s;
}
const val = (cells, k = 'Z99') => { const s = sheet(cells); const v = s.value(k); return v instanceof E.Err ? v.code : v; };
const f = (formula, cells = {}) => val({ ...cells, Z99: formula });
let n = 0;
const eq = (a, b, msg) => { n++; assert.deepStrictEqual(a, b, msg); };

// addresses
eq(E.colName(0), 'A'); eq(E.colName(25), 'Z'); eq(E.colName(26), 'AA'); eq(E.colName(701), 'ZZ'); eq(E.colName(702), 'AAA');
eq(E.colIndex('AA'), 26); eq(E.parseKey('$B$3'), { r: 2, c: 1 }); eq(E.parseKey('A0'), null);

// literals
eq(E.literal('42'), 42); eq(E.literal('1,234.5'), 1234.5); eq(E.literal('50%'), 0.5); eq(E.literal('-.5e2'), -50);
eq(E.literal('true'), true); eq(E.literal("'123"), '123'); eq(E.literal('12 apples'), '12 apples'); eq(E.literal('0x1F'), '0x1F');

// precedence & operators (Excel rules: unary minus binds tighter than ^, ^ is left-assoc)
eq(f('=1+2*3'), 7); eq(f('=(1+2)*3'), 9); eq(f('=-2^2'), 4); eq(f('=2^3^2'), 64); eq(f('=2^-1'), 0.5);
eq(f('=50%*10'), 5); eq(f('=1+2&"x"'), '3x'); eq(f('=1+1=2'), true); eq(f('="a"<"B"'), true); eq(f('=0.1+0.2'), 0.1 + 0.2);
eq(f('=1<>1'), false); eq(f('=  3 *  ( 4 - 1 ) '), 9); eq(f('=1/0'), '#DIV/0!'); eq(f('="x"+1'), '#VALUE!');
eq(f('="3"+1'), 4); eq(f('=TRUE+1'), 2); eq(f('=A1+1'), 1); eq(f('=A1&"!"'), '!'); eq(f('=A1'), 0);

// parse errors & names
eq(f('=1+'), '#ERROR!'); eq(f('=SUM(1'), '#ERROR!'); eq(f('="abc'), '#ERROR!'); eq(f('=foo'), '#NAME?');
eq(f('=NOPE(1)'), '#NAME?'); eq(f('=ROUND()'), '#VALUE!'); eq(f('=LOG10(100)'), 2); eq(f('=sum(1,2)'), 3);

// references, ranges, errors propagate
const grid = { A1: '10', A2: '20', A3: '30', A4: 'text', A5: '', B1: '1', B2: '2', B3: '3' };
eq(f('=SUM(A1:A5)', grid), 60); eq(f('=SUM(A1:A3, 5, "5")', grid), 70); eq(f('=SUM(A4)', grid), 0);
eq(f('=AVERAGE(A1:A4)', grid), 20); eq(f('=COUNT(A1:A5)', grid), 3); eq(f('=COUNTA(A1:A5)', grid), 4);
eq(f('=COUNTBLANK(A1:A5)', grid), 1); eq(f('=MAX(A1:B3)', grid), 30); eq(f('=MIN(A1:B3)', grid), 1);
eq(f('=SUMPRODUCT(A1:A3,B1:B3)', grid), 140); eq(f('=MEDIAN(A1:A3, 100)', grid), 25);
eq(f('=A1:A3', grid), '#VALUE!'); eq(f('=SUM(A1:A3)/0', grid), '#DIV/0!');
eq(f('=SUM(A1:A2)', { A1: '=1/0', A2: '1' }), '#DIV/0!'); eq(f('=$A$1*2', grid), 20);

// cycles
eq(val({ A1: '=B1', B1: '=A1' }, 'A1'), '#CYCLE!'); eq(val({ A1: '=A1+1' }, 'A1'), '#CYCLE!');
eq(val({ A3: '=SUM(A1:A3)' }, 'A3'), '#CYCLE!');

// long fill-down chains don't blow the stack
const chain = { A1: '1' };
for (let r = 2; r <= 5000; r++) chain['A' + r] = `=A${r - 1}+1`;
eq(val(chain, 'A5000'), 5000);

// logic (IF is lazy)
eq(f('=IF(A1>5,"big","small")', grid), 'big'); eq(f('=IF(FALSE,1/0,7)'), 7); eq(f('=IF(1>2,1)'), false);
eq(f('=IFERROR(1/0,"n/a")'), 'n/a'); eq(f('=IFERROR(5,0)'), 5); eq(f('=AND(TRUE,1,A1)', grid), true);
eq(f('=OR(FALSE,0)'), false); eq(f('=NOT(0)'), true); eq(f('=ISBLANK(A5)', grid), true); eq(f('=ISERROR(1/0)'), true);
eq(f('=ISNUMBER(A1)', grid), true); eq(f('=ISTEXT(A4)', grid), true);

// math
eq(f('=ROUND(1.005,2)'), 1.01); eq(f('=ROUND(-2.5)'), -3); eq(f('=ROUND(1234,-2)'), 1200); eq(f('=ROUNDUP(1.001,2)'), 1.01);
eq(f('=ROUNDDOWN(-1.9)'), -1); eq(f('=MOD(-3,2)'), 1); eq(f('=INT(-1.5)'), -2); eq(f('=SQRT(-1)'), '#NUM!');
eq(f('=POWER(2,10)'), 1024); eq(f('=ABS(-3)'), 3); eq(f('=(-8)^(1/3)'), '#NUM!');
const rb = f('=RANDBETWEEN(1,3)'); assert(rb >= 1 && rb <= 3 && Number.isInteger(rb));

// text
eq(f('=CONCAT("a",A1:A2,TRUE)', grid), 'a1020TRUE'); eq(f('=LEN("hello")'), 5); eq(f('=UPPER("abc")'), 'ABC');
eq(f('=PROPER("hello wORLD-foo")'), 'Hello World-Foo'); eq(f('=TRIM("  a   b ")'), 'a b');
eq(f('=LEFT("abc")'), 'a'); eq(f('=RIGHT("abc",2)'), 'bc'); eq(f('=RIGHT("abc",0)'), ''); eq(f('=MID("abcdef",2,3)'), 'bcd');
eq(f('=FIND("c","abc")'), 3); eq(f('=FIND("z","abc")'), '#VALUE!'); eq(f('=SUBSTITUTE("a-b-c","-","+")'), 'a+b+c');
eq(f('=REPT("ab",3)'), 'ababab'); eq(f('=VALUE("1,500")'), 1500); eq(f('="say ""hi"""'), 'say "hi"');

// conditional aggregates
const sales = { A1: 'east', A2: 'west', A3: 'east', A4: 'north', B1: '100', B2: '250', B3: '50', B4: '' };
eq(f('=SUMIF(A1:A4,"east",B1:B4)', sales), 150); eq(f('=SUMIF(B1:B4,">60")', sales), 350);
eq(f('=COUNTIF(A1:A4,"e*")', sales), 2); eq(f('=COUNTIF(A1:A4,"<>east")', sales), 2); eq(f('=COUNTIF(B1:B4,"")', sales), 1);
eq(f('=COUNTIF(B1:B4,"<>")', sales), 3); eq(f('=COUNTIF(B1:B4,100)', sales), 1); eq(f('=AVERAGEIF(A1:A4,"east",B1:B4)', sales), 75);
eq(f('=SUMIFS(B1:B4,A1:A4,"east",B1:B4,">=60")', sales), 100); eq(f('=COUNTIFS(A1:A4,"?est")', sales), 1);
eq(f('=SUMIF(A1:A4,"east")', sales), 0);

// lookups
const t = { A1: 'apple', B1: '1.2', A2: 'banana', B2: '0.5', A3: 'cherry', B3: '3', D1: '0', E1: 'F', D2: '60', E2: 'C', D3: '80', E3: 'A' };
eq(f('=VLOOKUP("banana",A1:B3,2,FALSE)', t), 0.5); eq(f('=VLOOKUP("kiwi",A1:B3,2,FALSE)', t), '#N/A');
eq(f('=VLOOKUP(75,D1:E3,2)', t), 'C'); eq(f('=VLOOKUP(99,D1:E3,2,TRUE)', t), 'A'); eq(f('=VLOOKUP("apple",A1:B3,3,FALSE)', t), '#REF!');
eq(f('=XLOOKUP("cherry",A1:A3,B1:B3)', t), 3); eq(f('=XLOOKUP("kiwi",A1:A3,B1:B3,"none")', t), 'none');
eq(f('=MATCH("banana",A1:A3,0)', t), 2); eq(f('=MATCH(65,D1:D3)', t), 2); eq(f('=INDEX(A1:B3,3,2)', t), 3);
eq(f('=INDEX(A1:B3,9,1)', t), '#REF!'); eq(f('=INDEX(A1:A3,2)&"!"', t), 'banana!'); eq(f('=VLOOKUP("b*",A1:B3,2,FALSE)', t), 0.5);

// shifting formulas for copy/paste & fill
eq(E.shift('=A1+$B$1+C$1+$D1', 1, 1), '=B2+$B$1+D$1+$D2'); eq(E.shift('=SUM(A1:A3)*2', 2, 0), '=SUM(A3:A5)*2');
eq(E.shift('=A1&"A1"', 0, 1), '=B1&"A1"'); eq(E.shift('=A1', -1, 0), '=#REF!'); eq(E.shift('plain', 3, 3), 'plain');
eq(f('=#REF!+1'), '#REF!');
eq(E.tidy('=sum(a1:b2'), '=SUM(A1:B2)'); eq(E.tidy('=if(a1>0,"x(",round(b1'), '=IF(A1>0,"x(",ROUND(B1))'); eq(E.tidy('hello'), 'hello');

// reference spans for highlighting while editing
eq(E.refSpans('=A1+SUM(B2:C3)').map(s => [s.s, s.e, s.r0, s.c0, s.r1, s.c1]), [[1, 3, 0, 0, 0, 0], [8, 13, 1, 1, 2, 2]]);
eq(E.refSpans('=SUM(A1:'), [{ s: 5, e: 7, r0: 0, c0: 0, r1: 0, c1: 0, sheet: null }]);
eq(E.refSpans("='My Data'!B2+C3").map(x => x.sheet), ['My Data', null]);

// workbooks: cross-sheet references, quoting, renames, cycles across sheets
const book = new E.Workbook();
const budget = book.add('Budget', 'b'), data = book.add('My Data', 'd'), sum = book.add('Summary', 's');
budget.set('A1', '100'); budget.set('A2', '=A1*2');
data.set('B1', '1'); data.set('B2', '2'); data.set('B3', '3');
sum.set('A1', '=Budget!A2+1'); sum.set('A2', "=SUM('My Data'!B1:B3)"); sum.set('A3', '=budget!a1');
sum.set('A4', '=Nope!A1'); sum.set('A5', "=SUM('My Data'!B1:Budget!B3)");
budget.set('C1', '=Summary!C1'); sum.set('C1', '=Budget!C1');
book.recalc();
const bv = (s, k) => { const v = s.value(k); return v instanceof E.Err ? v.code : v; };
eq(bv(sum, 'A1'), 201); eq(bv(sum, 'A2'), 6); eq(bv(sum, 'A3'), 100); eq(bv(sum, 'A4'), '#REF!'); eq(bv(sum, 'A5'), '#ERROR!');
eq(bv(sum, 'C1'), '#CYCLE!'); eq(bv(budget, 'C1'), '#CYCLE!');
// A remote edit lands with no explicit recalc: dependents on other sheets must still read fresh.
budget.set('A1', '50'); eq(bv(sum, 'A1'), 101);
budget.set('A1', '100'); eq(bv(sum, 'A1'), 201);
eq(E.quoteSheet('Budget'), 'Budget'); eq(E.quoteSheet('My Data'), "'My Data'"); eq(E.quoteSheet("Bob's"), "'Bob''s'"); eq(E.quoteSheet('Q1'), "'Q1'");
eq(E.renameRefs("=Budget!A1+'budget'!B2+A1+Other!A1", 'Budget', 'Plan 2027'), "='Plan 2027'!A1+'Plan 2027'!B2+A1+Other!A1");
eq(E.shift("='My Data'!A1+$B$2", 1, 1), "='My Data'!B2+$B$2"); eq(E.tidy("=sum('My Data'!b1:b3"), "=SUM('My Data'!B1:B3)");
eq(new E.Sheet().book.sheets.length, 1);
const fresh = sheet({ A1: '1', B1: '=A1*10', C1: 'x' });
fresh.set('A1', '5'); eq(fresh.value('B1'), 50); fresh.set('C1', ''); eq(fresh.value('C1'), null);
const rn = book.add('Late', 'late'); rn.set('A1', '=Budget!A1+1'); eq(bv(rn, 'A1'), 101);

// dates: Excel-compatible serials, locale-aware parsing, date functions, TEXT
eq(E.literal('2024-01-01'), 45292); eq(E.literal('1900-03-01'), 61); eq(E.literal('Jan 1, 2024'), 45292); eq(E.literal('1 January 2024'), 45292);
eq(E.literal('2024-02-30'), '2024-02-30'); eq(E.literal('2024-01-01 18:00'), 45292.75); eq(E.literal('2024-01-01 6:00 pm'), 45292.75);
eq(E.literal('1/2/2024'), 45293); E.config.dateOrder = 'dmy'; eq(E.literal('1/2/2024'), 45323); E.config.dateOrder = 'mdy';
eq(E.literal('Sept 5 2024'), 45540); eq(E.literal('Septx 5 2024'), 'Septx 5 2024'); eq(E.literal('Sep 5, 2024'), 45540); eq(E.isoDate(45292), '2024-01-01'); eq(E.isoDate(45292.75), '2024-01-01 18:00');
eq(f('=DATE(2024,1,1)'), 45292); eq(f('=DATE(2024,14,1)'), E.literal('2025-02-01')); eq(f('=YEAR("2024-03-15")'), 2024);
eq(f('=MONTH(A1)', { A1: '2024-03-15' }), 3); eq(f('=DAY(A1)+0', { A1: '15 Mar 2024' }), 15);
eq(f('=WEEKDAY(DATE(2024,1,1))'), 2); eq(f('=WEEKDAY(DATE(2024,1,1),2)'), 1);
eq(f('=EDATE(DATE(2024,1,31),1)'), E.literal('2024-02-29')); eq(f('=EDATE(DATE(2024,3,31),-1)'), E.literal('2024-02-29'));
eq(f('=EOMONTH(DATE(2024,2,10),0)'), E.literal('2024-02-29')); eq(f('=EOMONTH(DATE(2024,2,10),-1)'), E.literal('2024-01-31'));
eq(f('=DAYS("2024-03-01","2024-02-01")'), 29); eq(f('=A2-A1', { A1: '2024-01-01', A2: '2024-12-31' }), 365);
eq(f('=DATEDIF("2020-05-10","2024-05-09","Y")'), 3); eq(f('=DATEDIF("2024-01-31","2024-03-01","M")'), 1); eq(f('=DATEDIF(2,1,"D")'), '#NUM!');
eq(f('=NETWORKDAYS("2024-01-01","2024-01-14")'), 10); eq(f('=TODAY()=INT(NOW())'), true);
eq(f('=TEXT(DATE(2024,1,1),"dddd dd mmm yyyy")'), 'Monday 01 Jan 2024'); eq(f('=TEXT(45292.75,"yyyy-mm-dd hh:mm")'), '2024-01-01 18:00');
eq(f('=TEXT(45292.75,"h:mm am/pm")'), '6:00 PM'); eq(f('=TEXT(1234.5,"#,##0.00")'), '1,234.50'); eq(f('=TEXT(0.256,"0.0%")'), '25.6%');
eq(f('=TEXT(5,"000")'), '005'); eq(f('=TEXT(1234.5,"$#,##0")'), '$1,235'); eq(f('=TEXT("abc","0.00")'), 'abc');
eq(f('=COUNTIF(A1:A3,">=2024-02-01")', { A1: '2024-01-15', A2: '2024-02-15', A3: '2024-03-15' }), 2);

// inserting / deleting rows and columns: references follow their cells
const ins = (raw, at, n = 1, axis = 'r', home = 'S') => E.adjustRefs(raw, { sheet: 'S', axis, at, n, del: false }, home);
const del = (raw, at, n = 1, axis = 'r', home = 'S') => E.adjustRefs(raw, { sheet: 'S', axis, at, n, del: true }, home);
eq(ins('=A5+A3', 4), '=A6+A3'); eq(ins('=SUM(A1:A10)', 4), '=SUM(A1:A11)'); eq(ins('=SUM(A5:A10)', 4), '=SUM(A6:A11)');
eq(ins('=SUM(A1:A4)', 4), '=SUM(A1:A4)'); eq(ins('=$A$5*$B5', 4, 2), '=$A$7*$B7'); eq(ins('=A10:A5', 4), '=A11:A6');
eq(ins('=Other!A5+A5', 4), '=Other!A5+A6'); eq(ins('=S!A5+A5', 4, 1, 'r', 'Other'), '=S!A6+A5'); eq(ins('="A5"&A5', 4), '="A5"&A6');
eq(ins('=B1+A1', 1, 1, 'c'), '=C1+A1'); eq(ins('=SUM(A1:C1)', 1, 1, 'c'), '=SUM(A1:D1)'); eq(ins('=A1', 0, 3, 'c'), '=D1');
eq(del('=A5', 4, 2), '=#REF!'); eq(del('=A7+A3', 4, 2), '=A5+A3'); eq(del('=SUM(A1:A10)', 4, 2), '=SUM(A1:A8)');
eq(del('=SUM(A5:A6)', 4, 2), '=SUM(#REF!)'); eq(del('=SUM(A6:A9)', 4, 2), '=SUM(A5:A7)'); eq(del('=SUM(A2:A5)', 4, 2), '=SUM(A2:A4)');
eq(del('=SUM(B1:D1)', 1, 1, 'c'), '=SUM(B1:C1)'); eq(del('=C1*2', 1, 1, 'c'), '=B1*2'); eq(del('=B1', 1, 1, 'c'), '=#REF!');
eq(E.adjustRefs("='My Data'!B2+B2", { sheet: 'My Data', axis: 'r', at: 1, n: 1, del: false }, 'X'), "='My Data'!B3+B2");
eq(ins('plain text', 0), 'plain text'); eq(ins('=SUM(', 0), '=SUM(');
const cfs = sheet({ A1: '5', B1: '3' });
eq(cfs.evalFormula('=A1>B1'), true); eq(cfs.evalFormula('A1*2'), 10); eq(cfs.evalFormula('=SUM(').code, '#ERROR!');

// every function has help text, and every help entry is a real function
for (const name of Object.keys(E.FUNCS).concat('IF', 'IFERROR')) assert(E.HELP[name], 'missing HELP for ' + name);
for (const name of Object.keys(E.HELP)) assert(E.FUNCS[name] || name === 'IF' || name === 'IFERROR', 'HELP for unknown ' + name);

// .xlsx: write the sample workbook, read it back, and read Excel's own conventions (async: uses the zip streams)
(async () => {
  const X = require('./xlsx.js'), S = require('./sample.js');
  eq(X.crc32(new TextEncoder().encode('hello')), 0x3610a686);
  const wb = new E.Workbook(), sheets = S.sheets.map(s => {
    const styles = {};
    for (const [rg, st] of s.formats) {
      const [a, b = a] = rg.split(':').map(E.parseKey);
      for (let r = a.r; r <= b.r; r++) for (let c = a.c; c <= b.c; c++) styles[E.key(r, c)] = { ...styles[E.key(r, c)], ...st };
    }
    const sh = wb.add(s.name, s.id);
    for (const [k, v] of Object.entries(s.cells)) sh.set(k, v);
    return { sh, out: { name: s.name, cells: s.cells, styles, widths: s.widths, freeze: { rows: 4, cols: 1 }, filter: { range: 'A4:F11' } } };
  });
  wb.recalc();
  for (const { sh, out } of sheets) out.values = Object.fromEntries([...sh.cells.keys()].map(k => [k, sh.value(k)]));
  const bytes = await X.toXlsx(sheets.map(x => x.out), { currency: '₹' });
  const back = await X.fromXlsx(bytes);
  eq(back.map(s => s.name), ['Budget', 'Bills']);
  sheets.forEach(({ out }, i) => {
    eq(back[i].cells, out.cells);
    eq(back[i].styles, out.styles);
    eq(back[i].widths, out.widths);
    eq(back[i].freeze, { rows: 4, cols: 1 });
    eq(back[i].filter, { range: 'A4:F11', cols: [] });
  });
  // Excel-style parts: shared strings, shared formulas, _xlfn. prefixes, booleans, text that looks like a number
  const XML = '<?xml version="1.0"?>', ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const hand = await X.zip([
    ['xl/workbook.xml', `${XML}<workbook ${ns}><sheets><sheet name="Data &amp; more" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `${XML}<Relationships><Relationship Id="rId1" Target="/xl/worksheets/sheet1.xml"/></Relationships>`],
    ['xl/sharedStrings.xml', `${XML}<sst><si><t>Price</t></si><si><r><t>Tot</t></r><r><t>al</t></r></si><si><t>007</t></si></sst>`],
    ['xl/worksheets/sheet1.xml', `${XML}<worksheet ${ns}><sheetData>` +
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>2</v></c><c r="C1" t="b"><v>1</v></c></row>' +
      '<row r="2"><c r="A2"><v>10</v></c><c r="B2"><f t="shared" ref="B2:B4" si="0">A2*2</f><v>20</v></c></row>' +
      '<row r="3"><c r="A3"><v>2.5</v></c><c r="B3"><f t="shared" si="0"/><v>5</v></c></row>' +
      '<row r="4"><c r="A4" t="s"><v>1</v></c><c r="B4"><f t="shared" si="0"/></c><c r="C4"><f>_xlfn.XLOOKUP("x",A1:A3,B1:B3,"none")</f></c></row>' +
      '</sheetData></worksheet>'],
  ]);
  const [h] = await X.fromXlsx(hand);
  eq(h.name, 'Data & more');
  eq(h.cells, { A1: 'Price', B1: "'007", C1: 'TRUE', A2: '10', B2: '=A2*2', A3: '2.5', B3: '=A3*2', A4: 'Total', B4: '=A4*2', C4: '=XLOOKUP("x",A1:A3,B1:B3,"none")' });
  let rejected = false;
  try { await X.fromXlsx(new TextEncoder().encode('not a zip file at all')); } catch { rejected = true; }
  eq(rejected, true);
  console.log(`ok — ${n} checks passed`);
})().catch(e => { console.error(e); process.exit(1); });
