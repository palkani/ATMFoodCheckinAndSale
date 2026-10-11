import ExcelJS from 'exceljs';
import * as sheet from './sheet.js';
import { ROLES } from './auth.js';

const METHODS = ['cash', 'zelle', 'card'];
const norm = (s) => s.toLowerCase().replace(/[^a-z]/g, '');
const blank = () => ({ cash: 0, zelle: 0, card: 0, all: 0, count: 0 });

export async function buildReport(db, day) {
  const allRows = db.prepare("SELECT * FROM transactions WHERE date(created_at,'localtime')=? ORDER BY id").all(day);
  const rows = allRows.filter((r) => r.method !== 'volunteer'); // volunteer items are free, so they stay out of sales totals
  const byKind = { snack: blank(), lunch: blank() }, total = blank();
  const hours = new Map(), users = new Map();
  for (const r of rows) {
    for (const t of [byKind[r.kind], total]) { t[r.method] += r.total_cents; t.all += r.total_cents; t.count++; }
    const u = users.get(r.sold_by) || { user: ROLES[r.sold_by] || r.sold_by || '(unknown)', snack: 0, lunch: 0, count: 0, amount: 0 };
    u[r.kind] += r.total_cents; u.count++; u.amount += r.total_cents; users.set(r.sold_by, u);
    const h = new Date(r.created_at.replace(' ', 'T') + 'Z').getHours();
    const e = hours.get(h) || { hour: h, count: 0, amount: 0 }; e.count++; e.amount += r.total_cents; hours.set(h, e);
  }
  const items = db.prepare(`SELECT l.kind, l.name, SUM(l.qty) qty, SUM(l.qty * l.unit_cents) revenue,
      SUM(CASE WHEN t.method='cash' THEN l.qty ELSE 0 END) cash, SUM(CASE WHEN t.method='zelle' THEN l.qty ELSE 0 END) zelle, SUM(CASE WHEN t.method='card' THEN l.qty ELSE 0 END) card
    FROM sale_lines l JOIN transactions t ON t.id=l.tx_id WHERE date(t.created_at,'localtime')=? AND t.method != 'volunteer'
    GROUP BY l.kind, l.name ORDER BY l.kind DESC, qty DESC`).all(day);

  // Preorders come from the Excel sheet's current state (not date-bound).
  const pre = await sheet.summary();
  let preorders = { loaded: false };
  const lunches = new Map(); // normalized name -> { label, preorder, walkup }
  const bump = (label, k, n) => { const key = norm(label); const e = lunches.get(key) || { label, preorder: 0, walkup: 0, volunteer: 0 }; e[k] += n; lunches.set(key, e); };
  if (pre) {
    const itemMap = new Map(), slotMap = new Map();
    let prepaid = 0, prepaidPicked = 0;
    for (const p of pre) {
      const amt = parseFloat(p.total) || 0; prepaid += amt; if (p.checkedIn) prepaidPicked += amt;
      const sl = slotMap.get(p.slot || '(none)') || { slot: p.slot || '(none)', orders: 0, checkedIn: 0, lunches: 0, lunchesPicked: 0 };
      sl.orders++; if (p.checkedIn) sl.checkedIn++;
      for (const i of p.items) {
        const it = itemMap.get(i.label) || { label: i.label, ordered: 0, pickedUp: 0 };
        it.ordered += i.qty; if (p.checkedIn) it.pickedUp += i.qty; itemMap.set(i.label, it);
        sl.lunches += i.qty; if (p.checkedIn) sl.lunchesPicked += i.qty;
        bump(i.label, 'preorder', p.checkedIn ? i.qty : 0);
      }
      slotMap.set(sl.slot, sl);
    }
    preorders = { loaded: true, orders: pre.length, checkedIn: pre.filter((p) => p.checkedIn).length, items: [...itemMap.values()], slots: [...slotMap.values()].sort((a, b) => a.slot.localeCompare(b.slot)), prepaid, prepaidPicked };
  }
  for (const i of items.filter((x) => x.kind === 'lunch')) bump(i.name, 'walkup', i.qty);

  // Items given to volunteers: who got what, and what it would have sold for.
  const volLines = db.prepare(`SELECT t.id, t.volunteer, l.kind, l.name, l.qty, l.unit_cents FROM sale_lines l JOIN transactions t ON t.id=l.tx_id
    WHERE t.method='volunteer' AND date(t.created_at,'localtime')=? ORDER BY t.id`).all(day);
  const people = new Map(), volItems = new Map();
  for (const l of volLines) {
    const name = l.volunteer || '(no name)';
    const p = people.get(name) || { name, snacks: 0, meals: 0, qty: 0, value: 0, visits: new Set() };
    p[l.kind === 'lunch' ? 'meals' : 'snacks'] += l.qty; p.qty += l.qty; p.value += l.qty * l.unit_cents; p.visits.add(l.id); people.set(name, p);
    const k = l.kind + '|' + l.name, it = volItems.get(k) || { kind: l.kind, name: l.name, qty: 0, value: 0 };
    it.qty += l.qty; it.value += l.qty * l.unit_cents; volItems.set(k, it);
    if (l.kind === 'lunch') bump(l.name, 'volunteer', l.qty);
  }
  const volunteers = { people: [...people.values()].map((p) => ({ ...p, visits: p.visits.size })), items: [...volItems.values()], totalQty: volLines.reduce((n, l) => n + l.qty, 0), totalValue: volLines.reduce((n, l) => n + l.qty * l.unit_cents, 0) };
  return {
    date: day, total, byKind, byUser: [...users.values()], hourly: [...hours.values()].sort((a, b) => a.hour - b.hour), items, preorders,
    lunches: [...lunches.values()].map((e) => ({ ...e, total: e.preorder + e.walkup + e.volunteer })), volunteers, rows: allRows,
  };
}

export async function reportWorkbook(r) {
  const wb = new ExcelJS.Workbook();
  const d = (c) => c / 100;
  const add = (name, header, data, widths) => {
    const ws = wb.addWorksheet(name); ws.addRow(header).font = { bold: true };
    data.forEach((x) => ws.addRow(x)); header.forEach((_, i) => (ws.getColumn(i + 1).width = widths?.[i] || 18)); return ws;
  };
  add('Summary', ['Category', 'Cash ($)', 'Zelle ($)', 'Card ($)', 'Total ($)', 'Transactions'], [
    ['Snack sales', ...METHODS.map((m) => d(r.byKind.snack[m])), d(r.byKind.snack.all), r.byKind.snack.count],
    ['Walk-up lunch', ...METHODS.map((m) => d(r.byKind.lunch[m])), d(r.byKind.lunch.all), r.byKind.lunch.count],
    ['ALL', ...METHODS.map((m) => d(r.total[m])), d(r.total.all), r.total.count],
  ]);
  add('By Seller', ['Seller', 'Snack ($)', 'Walk-up lunch ($)', 'Total ($)', 'Transactions'], r.byUser.map((u) => [u.user, d(u.snack), d(u.lunch), d(u.amount), u.count]));
  add('Items Sold', ['Type', 'Item', 'Qty', 'Revenue ($)', 'Cash qty', 'Zelle qty', 'Card qty'], r.items.map((i) => [i.kind === 'lunch' ? 'Walk-up lunch' : 'Snack', i.name, i.qty, d(i.revenue), i.cash, i.zelle, i.card]), [16, 30]);
  add('By Hour', ['Hour', 'Transactions', 'Amount ($)'], r.hourly.map((h) => [`${h.hour}:00`, h.count, d(h.amount)]));
  add('Lunches Served', ['Lunch', 'Preorder picked up', 'Walk-up sold', 'Given to volunteers', 'Total served'], r.lunches.map((l) => [l.label, l.preorder, l.walkup, l.volunteer, l.total]), [36]);
  add('Volunteers', ['Volunteer', 'Snacks', 'Meals', 'Total items', 'Value ($)', 'Times'], r.volunteers.people.map((p) => [p.name, p.snacks, p.meals, p.qty, d(p.value), p.visits]), [26]);
  add('Volunteer Items', ['Type', 'Item', 'Qty', 'Value ($)'], r.volunteers.items.map((i) => [i.kind === 'lunch' ? 'Meal' : 'Snack', i.name, i.qty, d(i.value)]), [14, 30]);
  if (r.preorders.loaded) {
    add('Preorders by Item', ['Item', 'Ordered', 'Picked up', 'Remaining'], r.preorders.items.map((i) => [i.label, i.ordered, i.pickedUp, i.ordered - i.pickedUp]), [36]);
    add('Preorders by Slot', ['Serving time', 'Orders', 'Checked in', 'Lunches', 'Lunches picked up'], r.preorders.slots.map((s) => [s.slot, s.orders, s.checkedIn, s.lunches, s.lunchesPicked]), [26]);
  }
  add('Transactions', ['Receipt #', 'Time', 'Type', 'Details', 'Payment', 'Amount ($)', 'Sold by'], r.rows.map((t) => [t.id, new Date(t.created_at.replace(' ', 'T') + 'Z').toLocaleString(), t.kind === 'lunch' ? 'Walk-up lunch' : 'Snack', t.description, t.method, d(t.total_cents), ROLES[t.sold_by] || t.sold_by]), [10, 22, 16, 50]);
  return wb.xlsx.writeBuffer();
}
