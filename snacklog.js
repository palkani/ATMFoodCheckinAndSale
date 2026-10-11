import path from 'node:path';
import { DATA_DIR } from './config.js';
import { appendToSheet, editWorkbook, VOLUNTEER_HEADERS } from './xl.js';

export const SNACK_PATH = process.env.SNACK_SHEET_PATH || path.join(DATA_DIR, 'snack_sales.xlsx');
let queue = Promise.resolve();

// Excel 2 of 2: every snack sale is appended as a row, so the file is always current.
const append = (name, values, headers) => {
  const run = () => editWorkbook(SNACK_PATH, (wb) => appendToSheet(wb, name, values, headers));
  return (queue = queue.then(run, run));
};
export const appendSnackSale = (values) => append('Snack Sales', values);
// Snacks handed to volunteers go on their own tab so they never mix with sales.
export const appendVolunteerSnack = (values) => append('Volunteer Snacks', values, VOLUNTEER_HEADERS);
