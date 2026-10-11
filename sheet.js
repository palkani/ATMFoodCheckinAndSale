import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { appendToSheet, editWorkbook, VOLUNTEER_HEADERS } from './xl.js';
import { DATA_DIR } from './config.js';

export const WALKUP = 'Walk-up Sales';
export const VOLUNTEER_MEALS = 'Volunteer Meals';
const EXTRA_SHEETS = [WALKUP, VOLUNTEER_MEALS];

export const SHEET_PATH = process.env.SHEET_PATH || path.join(DATA_DIR, 'preorders.xlsx');
const CHECK_COLS = ['Checked In', 'Checked In At', 'Checked In By'];

// Writes are serialized so two check-ins never clobber each other's save.
let queue = Promise.resolve();
const locked = (fn) => (queue = queue.then(fn, fn));

const text = (v) => {
  if (v == null) return '';
  if (typeof v === 'object') return v.text ?? v.result ?? (v instanceof Date ? v.toISOString() : v.richText?.map((r) => r.text).join('') ?? '');
  return String(v);
};
const digits = (s) => s.replace(/\D/g, '');

async function open() {
  if (!fs.existsSync(SHEET_PATH)) return null;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SHEET_PATH);
  const ws = wb.worksheets.find((w) => !EXTRA_SHEETS.includes(w.name));
  if (!ws) return null;
  const headers = [];
  ws.getRow(1).eachCell({ includeEmpty: false }, (c, n) => headers.push({ n, label: text(c.value).trim() }));
  return { wb, ws, headers };
}

function findCol(headers, ...needles) {
  return headers.find((h) => needles.some((n) => h.label.toLowerCase().includes(n)));
}

function readRows({ ws, headers }) {
  const phoneH = findCol(headers, 'phone', 'mobile', 'cell', 'contact');
  const nameHs = headers.filter((h) => /name/i.test(h.label));
  const doneH = headers.find((h) => h.label === CHECK_COLS[0]);
  const atH = headers.find((h) => h.label === CHECK_COLS[1]);
  const shown = headers.filter((h) => !CHECK_COLS.includes(h.label));
  const emailH = findCol(headers, 'email');
  const slotH = findCol(headers, 'serving', 'schedule', 'time');
  const totalH = findCol(headers, 'total');
  // Quantity columns look like "$18.99 × 2"; payment-details columns are never sent to the browser.
  const isItem = (h, v) => /×\s*\d+\s*$/.test(v);
  const hidden = (h) => /paypal|credit card|card/i.test(h.label) || h === phoneH || h === emailH || nameHs.includes(h);
  const rows = [], itemLabels = new Set();
  ws.eachRow({ includeEmpty: false }, (row, rn) => {
    if (rn === 1) return;
    const cells = shown.map((h) => [h, text(row.getCell(h.n).value)]).filter(([, v]) => v !== '');
    if (!cells.length) return;
    cells.forEach(([h, v]) => isItem(h, v) && itemLabels.add(h.label));
    const items = cells.filter(([h, v]) => isItem(h, v)).map(([h, v]) => ({ label: h.label, qty: +v.match(/×\s*(\d+)\s*$/)[1] })).filter((i) => i.qty > 0);
    rows.push({
      row: rn,
      name: nameHs.map((h) => text(row.getCell(h.n).value)).join(' ').trim(),
      phone: phoneH ? text(row.getCell(phoneH.n).value) : '',
      email: emailH ? text(row.getCell(emailH.n).value) : '',
      slot: slotH ? text(row.getCell(slotH.n).value) : '',
      total: totalH ? text(row.getCell(totalH.n).value) : '',
      items,
      fields: cells.filter(([h, v]) => !isItem(h, v) && !hidden(h) && h !== slotH && h !== totalH && !/fee/i.test(h.label)).map(([h, v]) => [h.label, v]),
      checkedIn: doneH ? /^yes$/i.test(text(row.getCell(doneH.n).value)) : false,
      checkedInAt: atH ? text(row.getCell(atH.n).value) : '',
    });
  });
  return { rows, itemLabels: [...itemLabels], hasPhone: !!phoneH, hasName: nameHs.length > 0 };
}

export async function status() {
  const s = await open();
  if (!s) return { loaded: false };
  const { rows, hasPhone, hasName } = readRows(s);
  return { loaded: true, total: rows.length, checkedIn: rows.filter((r) => r.checkedIn).length, hasPhone, hasName, columns: s.headers.map((h) => h.label) };
}

export async function search(q) {
  const s = await open();
  if (!s) return [];
  const { rows } = readRows(s);
  q = q.trim().toLowerCase();
  const d = digits(q);
  const hit = !q ? rows : rows.filter((r) => r.name.toLowerCase().includes(q) || (d && digits(r.phone).includes(d)) || r.email.toLowerCase().includes(q) || r.fields.some(([k, v]) => /id/i.test(k) && v === q));
  return hit.sort((a, b) => a.checkedIn - b.checkedIn || a.name.localeCompare(b.name)).slice(0, 50);
}

export const save = (buf) => locked(async () => {
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(buf); } catch { throw Object.assign(new Error('Could not read that file. Upload an .xlsx file.'), { code: 400 }); }
  if (!wb.worksheets.length) throw Object.assign(new Error('Workbook has no sheets'), { code: 400 });
  fs.mkdirSync(path.dirname(SHEET_PATH), { recursive: true });
  // Walk-up sales and volunteer meals already recorded must survive replacing the preorder sheet.
  const old = new ExcelJS.Workbook();
  if (fs.existsSync(SHEET_PATH)) await old.xlsx.readFile(SHEET_PATH);
  let carried = false;
  for (const name of EXTRA_SHEETS) {
    const oldWs = old.getWorksheet(name);
    if (!oldWs || oldWs.rowCount < 2 || wb.getWorksheet(name)) continue;
    const headers = oldWs.getRow(1).values.slice(1);
    oldWs.eachRow((r, n) => { if (n > 1) appendToSheet(wb, name, r.values.slice(1), headers); });
    carried = true;
  }
  if (carried) await wb.xlsx.writeFile(SHEET_PATH); else fs.writeFileSync(SHEET_PATH, buf);
});

// Walk-up lunch sales go on their own tab of the preorder workbook.
export const appendVolunteerMeal = (values) => locked(() => editWorkbook(SHEET_PATH, (wb) => appendToSheet(wb, VOLUNTEER_MEALS, values, VOLUNTEER_HEADERS)));
export const appendWalkup = (values) => locked(() => editWorkbook(SHEET_PATH, (wb) => appendToSheet(wb, WALKUP, values)));

// Lunch counts per type, split by checked in / not yet, plus the number of orders (people).
export async function counts() {
  const s = await open();
  if (!s) return null;
  const { rows, itemLabels } = readRows(s);
  const items = new Map(itemLabels.map((label) => [label, { label, ordered: 0, checkedIn: 0 }]));
  for (const r of rows) for (const i of r.items) {
    const e = items.get(i.label) || { label: i.label, ordered: 0, checkedIn: 0 };
    e.ordered += i.qty; if (r.checkedIn) e.checkedIn += i.qty; items.set(i.label, e);
  }
  const done = rows.filter((r) => r.checkedIn).length;
  return { people: { total: rows.length, checkedIn: done, notCheckedIn: rows.length - done }, items: [...items.values()].sort((a, b) => a.label.localeCompare(b.label)) };
}

export async function summary() {
  const s = await open();
  return s ? readRows(s).rows : null;
}

export const setCheckIn = (rowNum, checked, by = '') => locked(async () => {
  const s = await open();
  if (!s) throw Object.assign(new Error('No sheet loaded'), { code: 400 });
  const { wb, ws } = s;
  const colFor = (label) => {
    let h = s.headers.find((x) => x.label === label);
    if (!h) { const n = Math.max(ws.columnCount, ...s.headers.map((x) => x.n)) + 1; ws.getRow(1).getCell(n).value = label; ws.getRow(1).getCell(n).font = { bold: true }; ws.getColumn(n).width = 20; h = { n, label }; s.headers.push(h); }
    return h.n;
  };
  const row = ws.getRow(rowNum);
  if (rowNum < 2 || !row.hasValues) throw Object.assign(new Error('Row not found'), { code: 404 });
  const doneN = colFor(CHECK_COLS[0]), atN = colFor(CHECK_COLS[1]), byN = colFor(CHECK_COLS[2]);
  if (checked && /^yes$/i.test(text(row.getCell(doneN).value))) throw Object.assign(new Error('Already checked in'), { code: 409 });
  row.getCell(doneN).value = checked ? 'Yes' : '';
  row.getCell(atN).value = checked ? new Date().toLocaleString() : '';
  row.getCell(byN).value = checked ? by : '';
  row.commit();
  const tmp = SHEET_PATH + '.tmp';
  await wb.xlsx.writeFile(tmp);
  fs.renameSync(tmp, SHEET_PATH);
});
