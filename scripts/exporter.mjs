// tabular-exporter —— 记录数组 → 带样式 Excel(xlsx) / 带 BOM 的 CSV
// 从 caijiqi-kaifa/electron/export/manager.ts 抽取并通用化。
// 零依赖：xlsx 用纯 JS 生成（ZIP STORE 法，内置 CRC32），CSV 自带 UTF-8 BOM。
// 用法：import { exportTable, exportToCsv, exportToExcel } from './scripts/exporter.mjs';

import { writeFileSync, readFileSync, unlinkSync } from 'node:fs';

// ───────────────── 通用工具 ─────────────────

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function csvEscape(v) {
  const s = String(v ?? '');
  if (s.includes(',') || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function toCellValue(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? '是' : '否';
  if (Array.isArray(v)) return v.join('\n');
  if (typeof v === 'object') return JSON.stringify(v);
  return v;
}

// ───────────────── CRC32 + ZIP(STORE) ─────────────────

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/**
 * 将多个 {name, data:Buffer} 打包成 STORE 法的 zip Buffer
 */
function zipStore(parts) {
  const chunks = [];
  const central = [];
  let offset = 0;

  for (const part of parts) {
    const nameBuf = Buffer.from(part.name, 'utf8');
    const data = part.data;
    const crc = crc32(data);

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8); // method 0 = store
    local.writeUInt16LE(0, 10); // time
    local.writeUInt16LE(0, 12); // date
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);

    chunks.push(local, nameBuf, data);

    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0, 8);
    cd.writeUInt16LE(0, 10);
    cd.writeUInt16LE(0, 12);
    cd.writeUInt16LE(0, 14);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(data.length, 20);
    cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(nameBuf.length, 28);
    cd.writeUInt16LE(0, 30);
    cd.writeUInt16LE(0, 32);
    cd.writeUInt16LE(0, 34);
    cd.writeUInt16LE(0, 36);
    cd.writeUInt32LE(0, 38);
    cd.writeUInt32LE(offset, 42);

    central.push(Buffer.concat([cd, nameBuf]));
    offset += local.length + nameBuf.length + data.length;
  }

  const centralBuf = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(parts.length, 8);
  eocd.writeUInt16LE(parts.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...chunks, centralBuf, eocd]);
}

// ───────────────── 样式 ─────────────────

const STYLE_HEADER = 1; // 加粗白字 + 蓝底
const STYLE_ZEBRA = 2; // 灰底斑马纹

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const LETTERS = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
function colLetter(idx) {
  let s = '';
  let n = idx + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = LETTERS[r] + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ───────────────── Excel 导出（纯 JS） ─────────────────

export function exportToExcel(rows, columns, opts = {}) {
  const sheetName = (opts.sheetName || 'Sheet1').slice(0, 31);
  const filePath = opts.filePath;

  const colsXml = columns
    .map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${c.width || 15}" customWidth="1"/>`)
    .join('');

  let body = '';
  // 表头
  columns.forEach((c, i) => {
    body += `<c r="${colLetter(i)}1" s="${STYLE_HEADER}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(c.header || c.key)}</t></is></c>`;
  });
  // 数据
  rows.forEach((row, ri) => {
    const r = ri + 2;
    const style = ri % 2 === 1 ? ` s="${STYLE_ZEBRA}"` : '';
    columns.forEach((c, i) => {
      const val = toCellValue(row[c.key]);
      if (typeof val === 'number' && !Number.isNaN(val)) {
        body += `<c r="${colLetter(i)}${r}"${style}><v>${val}</v></c>`;
      } else if (val === '' || val == null) {
        body += `<c r="${colLetter(i)}${r}"${style}/>`;
      } else {
        body += `<c r="${colLetter(i)}${r}"${style} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(val)}</t></is></c>`;
      }
    });
  });

  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>${colsXml}</cols>
  <sheetData>${body}</sheetData>
</worksheet>`;

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="${xmlEscape(sheetName)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const parts = [
    { name: '[Content_Types].xml', data: Buffer.from(contentTypes, 'utf8') },
    { name: '_rels/.rels', data: Buffer.from(rootRels, 'utf8') },
    { name: 'xl/workbook.xml', data: Buffer.from(workbookXml, 'utf8') },
    { name: 'xl/_rels/workbook.xml.rels', data: Buffer.from(workbookRels, 'utf8') },
    { name: 'xl/styles.xml', data: Buffer.from(STYLES_XML, 'utf8') },
    { name: 'xl/worksheets/sheet1.xml', data: Buffer.from(sheetXml, 'utf8') },
  ];

  const buf = zipStore(parts);
  if (filePath) writeFileSync(filePath, buf);
  return buf;
}

// ───────────────── CSV 导出 ─────────────────

export function exportToCsv(rows, columns, opts = {}) {
  const delimiter = opts.delimiter || ',';
  const headers = columns.map((c) => c.header || c.key);
  const lines = [headers.map(csvEscape).join(delimiter)];
  for (const row of rows) {
    lines.push(columns.map((c) => csvEscape(toCellValue(row[c.key]))).join(delimiter));
  }
  const content = '\uFEFF' + lines.join('\n'); // BOM 防中文乱码
  if (opts.filePath) writeFileSync(opts.filePath, content, 'utf8');
  return content;
}

// ───────────────── 统一入口 ─────────────────

export function exportTable(rows, columns, opts = {}) {
  const format = opts.format || 'xlsx';
  if (format === 'csv') return exportToCsv(rows, columns, opts);
  return exportToExcel(rows, columns, opts);
}

/**
 * 由 key 列表快速生成列配置（header 默认等于 key）
 */
export function columnsFromKeys(keys, widths = {}) {
  return keys.map((k) => ({ key: k, header: k, width: widths[k] || 15 }));
}

// ───────────────────────── 自测 ─────────────────────────
function _assert(name, cond) {
  if (!cond) { console.error('FAIL:', name); process.exitCode = 1; }
  else console.log('PASS:', name);
}

if (process.argv.includes('--selftest')) {
  const columns = [
    { key: 'title', header: '标题', width: 40 },
    { key: 'price', header: '价格', width: 12 },
    { key: 'inStock', header: '有货', width: 8 },
    { key: 'note', header: '备注', width: 30 },
  ];
  const rows = [
    { title: '无线鼠标, 黑色', price: 99.5, inStock: true, note: '热销' },
    { title: '机械键盘', price: 299, inStock: false, note: 'a,b\n多行' },
  ];

  // CSV
  const csv = exportToCsv(rows, columns, {});
  _assert('csv has BOM', csv.charCodeAt(0) === 0xfeff);
  _assert('csv escapes comma', csv.includes('"无线鼠标, 黑色"'));
  _assert('csv boolean 是', csv.includes('是') && csv.includes('否'));
  _assert('csv multiline quoted', csv.includes('"a,b\n多行"'));

  // Excel
  const xlsx = exportToExcel(rows, columns, {});
  _assert('xlsx is zip (PK)', xlsx[0] === 0x50 && xlsx[1] === 0x4b);
  _assert('xlsx has sheet part', xlsx.includes(Buffer.from('xl/worksheets/sheet1.xml')));
  _assert('xlsx has styles part', xlsx.includes(Buffer.from('xl/styles.xml')));
  _assert('xlsx contains header label', xlsx.includes(Buffer.from('标题')));
  _assert('xlsx contains data value', xlsx.includes(Buffer.from('无线鼠标')));
  _assert('xlsx contains number', xlsx.includes(Buffer.from('99.5')));

  // 落盘回读（zip 结构完整性）
  const outPath = new URL('_selftest.xlsx', import.meta.url).pathname.replace(/^\/+/, '');
  exportToExcel(rows, columns, { filePath: outPath });
  const back = readFileSync(outPath);
  _assert('xlsx written & readable', back.length === xlsx.length && back[0] === 0x50);
  try { unlinkSync(outPath); } catch (e) {}

  console.log('tabular-exporter self-test done.');
}
