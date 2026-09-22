// Excel .xlsx reader and writer, built on the platform's own zip streams (CompressionStream) — no library.
// Handles what Gridwork has: values, formulas (incl. shared ones), bold/italic, number formats, column widths,
// frozen panes and autofilter ranges. Plain script: sets window.Xlsx in the browser, module.exports in Node.
(function (root) {
'use strict';
const E = typeof module === 'object' && module.exports ? require('./engine.js') : root.Engine;

// ---------- zip ----------
const CRC = new Uint32Array(256).map((_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
function crc32(bytes) { let c = ~0; for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 255] ^ (c >>> 8); return ~c >>> 0; }
const pipe = async (bytes, stream) => new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(stream)).arrayBuffer());
function concat(parts) {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) { out.set(p, at); at += p.length; }
  return out;
}
async function zip(files) {
  const enc = new TextEncoder(), parts = [], central = [];
  let offset = 0;
  for (const [path, text] of files) {
    const name = enc.encode(path), raw = enc.encode(text), data = await pipe(raw, new CompressionStream('deflate-raw')), crc = crc32(raw);
    const local = new DataView(new ArrayBuffer(30)), dir = new DataView(new ArrayBuffer(46));
    // local header: signature, version, UTF-8 flag, deflate, time 0, date 1980-01-01, crc, sizes, name length
    [[0, 0x04034b50, 4], [4, 20, 2], [6, 0x800, 2], [8, 8, 2], [12, 0x21, 2], [14, crc, 4], [18, data.length, 4], [22, raw.length, 4], [26, name.length, 2]]
      .forEach(([at, v, size]) => size === 4 ? local.setUint32(at, v, true) : local.setUint16(at, v, true));
    [[0, 0x02014b50, 4], [4, 20, 2], [6, 20, 2], [8, 0x800, 2], [10, 8, 2], [14, 0x21, 2], [16, crc, 4], [20, data.length, 4], [24, raw.length, 4], [28, name.length, 2], [42, offset, 4]]
      .forEach(([at, v, size]) => size === 4 ? dir.setUint32(at, v, true) : dir.setUint16(at, v, true));
    parts.push(new Uint8Array(local.buffer), name, data);
    central.push(new Uint8Array(dir.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const size = central.reduce((n, p) => n + p.length, 0), end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.length, true);
  end.setUint16(10, files.length, true);
  end.setUint32(12, size, true);
  end.setUint32(16, offset, true);
  return concat([...parts, ...central, new Uint8Array(end.buffer)]);
}
// → { path: () => Promise<Uint8Array> }, inflating an entry only when it is read
function unzip(buf) {
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let e = u8.length - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error('This file isn’t an Excel workbook (.xlsx).');
  const files = {}, dec = new TextDecoder();
  let p = dv.getUint32(e + 16, true);
  for (let i = dv.getUint16(e + 10, true); i > 0; i--) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('This .xlsx file is damaged.');
    const method = dv.getUint16(p + 10, true), size = dv.getUint32(p + 20, true), nlen = dv.getUint16(p + 28, true), off = dv.getUint32(p + 42, true);
    const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
    const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true), data = u8.subarray(start, start + size);
    files[name] = () => method === 8 ? pipe(data, new DecompressionStream('deflate-raw')) : Promise.resolve(data);
    p += 46 + nlen + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
  }
  return files;
}

// ---------- xml ----------
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
const unesc = s => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (m, x) =>
  x[0] === '#' ? String.fromCodePoint(x[1] === 'x' || x[1] === 'X' ? parseInt(x.slice(2), 16) : +x.slice(1)) : { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[x.toLowerCase()]);
const attr = (tag, name) => { const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag); return m ? unesc(m[1]) : null; };
const texts = xml => [...xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(m => unesc(m[1])).join('');
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Formulas: Excel files spell newer functions _xlfn.XLOOKUP; Gridwork's AVG is Excel's AVERAGE.
const FUTURE = new Set(['XLOOKUP', 'CONCAT', 'DAYS']);
function renameFunctions(src, fn) {
  let toks;
  try { toks = E.tokenize(src); } catch { return src; }
  let out = src;
  for (let i = toks.length - 1; i >= 0; i--) {
    const t = toks[i], rep = t.k === 'id' ? fn(t.v.toUpperCase()) : null;
    if (rep) out = out.slice(0, t.s) + rep + out.slice(t.e);
  }
  return out;
}
const toExcel = f => renameFunctions(f, n => n === 'AVG' ? 'AVERAGE' : FUTURE.has(n) ? '_xlfn.' + n : null);
const fromExcel = f => renameFunctions(f, n => /^_XL(FN|WS)\./.test(n) ? n.replace(/^_XL(FN|WS)\./, '') : null);

// Number format ids → Gridwork formats ('num' 'int' 'pct' 'cur' 'date' 'dt')
function formatOf(id, custom) {
  if ((id >= 14 && id <= 17) || (id >= 27 && id <= 36) || (id >= 50 && id <= 58)) return 'date';
  if (id === 22) return 'dt';
  if (id === 9 || id === 10) return 'pct';
  if (id === 1 || id === 3) return 'int';
  if (id === 2 || id === 4) return 'num';
  if ((id >= 5 && id <= 8) || (id >= 37 && id <= 44)) return 'cur';
  const code = custom[id];
  if (!code) return '';
  const c = code.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '');
  if (/[yd]/i.test(c) || (/m/i.test(c) && !/[#0]/.test(c))) return /[hs]/i.test(c) ? 'dt' : 'date';
  if (c.includes('%')) return 'pct';
  if (/\[\$[^\]-]|[$€£¥₹]/.test(code)) return 'cur';
  if (/[0#]\.0/.test(c)) return 'num';
  return /[0#],[0#]/.test(c) ? 'int' : '';
}

// ---------- read ----------
async function fromXlsx(buf, limits = { rows: 1000, cols: 52 }) {
  const files = unzip(buf), dec = new TextDecoder();
  const text = async path => files[path] ? dec.decode(await files[path]()) : '';
  const wb = await text('xl/workbook.xml');
  if (!wb) throw new Error('This file isn’t an Excel workbook (.xlsx).');
  const rels = {};
  for (const m of (await text('xl/_rels/workbook.xml.rels')).matchAll(/<Relationship\b[^>]*>/g)) rels[attr(m[0], 'Id')] = attr(m[0], 'Target');
  const strings = [...(await text('xl/sharedStrings.xml')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => texts(m[1]));
  const st = await text('xl/styles.xml'), custom = {};
  for (const m of st.matchAll(/<numFmt\b[^>]*>/g)) custom[attr(m[0], 'numFmtId')] = attr(m[0], 'formatCode');
  const fonts = [...((/<fonts\b[\s\S]*?<\/fonts>/.exec(st) || [''])[0]).matchAll(/<font\b[^>]*?(?:\/>|>([\s\S]*?)<\/font>)/g)]
    .map(m => ({ b: /<b\b(?![^>]*val="(?:0|false)")/.test(m[1] || ''), i: /<i\b(?![^>]*val="(?:0|false)")/.test(m[1] || '') }));
  const xfs = [...((/<cellXfs\b[\s\S]*?<\/cellXfs>/.exec(st) || [''])[0]).matchAll(/<xf\b[^>]*>/g)].map(m => {
    const font = fonts[+attr(m[0], 'fontId') || 0] || {}, f = formatOf(+attr(m[0], 'numFmtId') || 0, custom), out = {};
    if (font.b) out.b = 1;
    if (font.i) out.i = 1;
    if (f) out.f = f;
    return Object.keys(out).length ? out : null;
  });
  const sheets = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    let target = rels[attr(m[0], 'r:id')];
    if (!target) continue;
    target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    const xml = await text(target);
    if (xml) sheets.push(readSheet(attr(m[0], 'name') || `Sheet${sheets.length + 1}`, xml, strings, xfs, limits));
  }
  return sheets;
}
function readSheet(name, xml, strings, xfs, limits) {
  const cells = {}, styles = {}, widths = {}, shared = {};
  for (const m of xml.matchAll(/<col\b[^>]*>/g)) {
    const min = +attr(m[0], 'min'), max = +attr(m[0], 'max'), w = +attr(m[0], 'width');
    if (w > 0 && attr(m[0], 'customWidth') !== '0') for (let c = min; c <= Math.min(max, limits.cols); c++) widths[E.colName(c - 1)] = Math.round(w * 7 + 5);
  }
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const tag = '<c' + m[1] + '>', p = E.parseKey(attr(tag, 'r') || '');
    if (!p || p.r >= limits.rows || p.c >= limits.cols) continue;
    const body = m[2] || '', t = attr(tag, 't'), st = xfs[+attr(tag, 's') || 0], k = E.key(p.r, p.c);
    const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(body), v = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1];
    let raw = '';
    if (fm) {
      const ftag = '<f' + fm[1] + '>', si = attr(ftag, 'si'), f = fm[2] != null ? unesc(fm[2]) : '';
      if (attr(ftag, 't') === 'shared' && si != null) {
        if (f) shared[si] = { f, r: p.r, c: p.c };
        const base = shared[si];
        if (base) raw = E.shift('=' + fromExcel(base.f), p.r - base.r, p.c - base.c); // shared formulas are written once, then shifted
      } else if (f) raw = '=' + fromExcel(f);
    }
    if (!raw) {
      const isText = t === 's' || t === 'inlineStr' || t === 'str';
      if (t === 's') raw = strings[+v] ?? '';
      else if (t === 'inlineStr') raw = texts(body);
      else if (t === 'str' || t === 'e') raw = unesc(v ?? '');
      else if (t === 'b') raw = v === '1' ? 'TRUE' : 'FALSE';
      else if (v != null && v !== '') raw = st?.f === 'date' || st?.f === 'dt' ? E.isoDate(+v) : String(+v);
      // Text that would otherwise read as a number, date or formula stays text.
      if (isText && raw && (raw[0] === '=' || raw[0] === "'" || typeof E.literal(raw) !== 'string')) raw = "'" + raw;
    }
    if (raw) cells[k] = raw;
    if (st) styles[k] = st;
  }
  const pane = /<pane\b[^>]*>/.exec(xml)?.[0], frozen = pane && /frozen/i.test(attr(pane, 'state') || '');
  const filter = /<autoFilter\b[^>]*>/.exec(xml)?.[0];
  return {
    name, cells, styles, widths, chart: null, cf: [],
    freeze: frozen ? { rows: Math.round(+attr(pane, 'ySplit') || 0), cols: Math.round(+attr(pane, 'xSplit') || 0) } : null,
    filter: filter && attr(filter, 'ref') ? { range: attr(filter, 'ref').replace(/\$/g, ''), cols: [] } : null,
  };
}

// ---------- write ----------
// sheets: [{ name, cells: {A1: raw}, values: {A1: computed}, styles: {A1: {b, i, f}}, widths: {A: px}, freeze, filter }]
async function toXlsx(sheets, { currency = '$' } = {}) {
  const FMTS = { num: 4, int: 3, pct: 164, cur: 165, date: 14, dt: 22 };
  const xfIndex = new Map([['0|0|', 0]]), xfs = [{ font: 0, fmt: 0 }];
  const xf = st => {
    if (!st) return 0;
    const font = (st.b ? 1 : 0) + (st.i ? 2 : 0), fmt = FMTS[st.f] || 0, id = `${st.b ? 1 : 0}|${st.i ? 1 : 0}|${st.f || ''}`;
    if (!xfIndex.has(id)) { xfIndex.set(id, xfs.length); xfs.push({ font, fmt }); }
    return xfIndex.get(id);
  };
  const cell = (k, raw, v, s) => {
    const S = s ? ` s="${s}"` : '';
    if (raw[0] === '=' && raw.length > 1) {
      const f = `<f>${esc(toExcel(raw.slice(1)))}</f>`;
      if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${k}"${S}>${f}<v>${v}</v></c>`;
      if (typeof v === 'boolean') return `<c r="${k}"${S} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
      if (v && typeof v === 'object' && v.code) return `<c r="${k}"${S} t="e">${f}<v>${/^#(REF!|DIV\/0!|N\/A|VALUE!|NAME\?|NUM!)$/.test(v.code) ? v.code : '#VALUE!'}</v></c>`;
      return `<c r="${k}"${S} t="str">${f}<v>${esc(v ?? '')}</v></c>`;
    }
    if (!raw) return `<c r="${k}"${S}/>`;
    const lit = E.literal(raw);
    if (typeof lit === 'number') return `<c r="${k}"${S}><v>${lit}</v></c>`;
    if (typeof lit === 'boolean') return `<c r="${k}"${S} t="b"><v>${lit ? 1 : 0}</v></c>`;
    return `<c r="${k}"${S} t="inlineStr"><is><t xml:space="preserve">${esc(lit)}</t></is></c>`;
  };
  const worksheets = sheets.map(s => {
    const rows = new Map();
    for (const k of new Set([...Object.keys(s.cells), ...Object.keys(s.styles || {})])) {
      const p = E.parseKey(k);
      if (!rows.has(p.r)) rows.set(p.r, []);
      rows.get(p.r).push([p.c, k]);
    }
    const data = [...rows.keys()].sort((a, b) => a - b).map(r => `<row r="${r + 1}">` +
      rows.get(r).sort((a, b) => a[0] - b[0]).map(([, k]) => cell(k, s.cells[k] || '', s.values?.[k], xf(s.styles?.[k]))).join('') + '</row>').join('');
    const cols = Object.entries(s.widths || {}).map(([L, px]) => [E.colIndex(L) + 1, px]).sort((a, b) => a[0] - b[0])
      .map(([c, px]) => `<col min="${c}" max="${c}" width="${Math.max(1, (px - 5) / 7).toFixed(2)}" customWidth="1"/>`).join('');
    const fz = s.freeze && (s.freeze.rows || s.freeze.cols) ? s.freeze : null;
    const pane = fz ? `<pane${fz.cols ? ` xSplit="${fz.cols}"` : ''}${fz.rows ? ` ySplit="${fz.rows}"` : ''} topLeftCell="${E.key(fz.rows || 0, fz.cols || 0)}" activePane="${fz.rows && fz.cols ? 'bottomRight' : fz.rows ? 'bottomLeft' : 'topRight'}" state="frozen"/>` : '';
    return XML + `<worksheet xmlns="${NS}" xmlns:r="${RNS}"><sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>` +
      (cols ? `<cols>${cols}</cols>` : '') + `<sheetData>${data}</sheetData>` + (s.filter?.range ? `<autoFilter ref="${esc(s.filter.range)}"/>` : '') + '</worksheet>';
  });
  const fonts = ['', '<b/>', '<i/>', '<b/><i/>'].map(x => `<font>${x}<sz val="11"/><name val="Calibri"/></font>`).join('');
  const styles = XML + `<styleSheet xmlns="${NS}">` +
    `<numFmts count="2"><numFmt numFmtId="164" formatCode="0.0%"/><numFmt numFmtId="165" formatCode="${esc(`"${currency}"#,##0.00`)}"/></numFmts>` +
    `<fonts count="4">${fonts}</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${xfs.length}">${xfs.map(x => `<xf numFmtId="${x.fmt}" fontId="${x.font}" fillId="0" borderId="0" xfId="0"${x.fmt ? ' applyNumberFormat="1"' : ''}${x.font ? ' applyFont="1"' : ''}/>`).join('')}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  const workbook = XML + `<workbook xmlns="${NS}" xmlns:r="${RNS}"><bookViews><workbookView/></bookViews><sheets>` +
    sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';
  const wbRels = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="${RNS}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
    `<Relationship Id="rId${sheets.length + 1}" Type="${RNS}/styles" Target="styles.xml"/></Relationships>`;
  const types = XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '</Types>';
  const rootRels = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Type="${RNS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  return zip([
    ['[Content_Types].xml', types], ['_rels/.rels', rootRels], ['xl/workbook.xml', workbook], ['xl/_rels/workbook.xml.rels', wbRels],
    ['xl/styles.xml', styles], ...worksheets.map((x, i) => [`xl/worksheets/sheet${i + 1}.xml`, x]),
  ]);
}

const api = { toXlsx, fromXlsx, zip, unzip, crc32 };
if (typeof module === 'object' && module.exports) module.exports = api; else root.Xlsx = api;
})(this);
