import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

export const SALE_HEADERS = ['Receipt #', 'Time', 'Items', 'Qty', 'Payment', 'Total ($)', 'Sold by'];

export function appendToSheet(wb, name, values) {
  let ws = wb.getWorksheet(name);
  if (!ws) {
    ws = wb.addWorksheet(name);
    ws.addRow(SALE_HEADERS).font = { bold: true };
    SALE_HEADERS.forEach((h, i) => (ws.getColumn(i + 1).width = i === 2 ? 44 : 16));
  }
  ws.addRow(values);
}

export const localStamp = (d = new Date()) => {
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
};

// Loads (or creates) the workbook at `file`, lets `fn` change it, then saves atomically.
export async function editWorkbook(file, fn) {
  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(file)) await wb.xlsx.readFile(file); else fs.mkdirSync(path.dirname(file), { recursive: true });
  fn(wb);
  const tmp = file + '.tmp';
  await wb.xlsx.writeFile(tmp);
  fs.renameSync(tmp, file);
}
