import path from 'node:path';
import { DATA_DIR } from './config.js';
import { appendToSheet, editWorkbook } from './xl.js';

export const SNACK_PATH = process.env.SNACK_SHEET_PATH || path.join(DATA_DIR, 'snack_sales.xlsx');
let queue = Promise.resolve();

// Excel 2 of 2: every snack sale is appended as a row, so the file is always current.
export const appendSnackSale = (values) => {
  const run = () => editWorkbook(SNACK_PATH, (wb) => appendToSheet(wb, 'Snack Sales', values));
  return (queue = queue.then(run, run));
};
