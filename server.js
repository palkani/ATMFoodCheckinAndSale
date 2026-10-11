import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DatabaseSync } from 'node:sqlite';
import * as sheet from './sheet.js';
import { SNACK_PATH, appendSnackSale, appendVolunteerSnack } from './snacklog.js';
import { localStamp } from './xl.js';
import { buildReport, reportWorkbook } from './report.js';
import { initAuth, ROLES } from './auth.js';
import { DATA_DIR } from './config.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = process.env.PORT || 3000;
const METHODS = ['cash', 'zelle', 'card'];

fs.mkdirSync(DATA_DIR, { recursive: true });
const db = new DatabaseSync(path.join(DATA_DIR, 'app.db'));
db.exec(`
  CREATE TABLE IF NOT EXISTS snack_items (
    id INTEGER PRIMARY KEY, name TEXT NOT NULL, price_cents INTEGER NOT NULL, active INTEGER NOT NULL DEFAULT 1);
  CREATE TABLE IF NOT EXISTS transactions (
    id INTEGER PRIMARY KEY, kind TEXT NOT NULL, method TEXT NOT NULL, total_cents INTEGER NOT NULL,
    description TEXT NOT NULL, preorder_id INTEGER, created_at TEXT NOT NULL DEFAULT (datetime('now')));
  CREATE TABLE IF NOT EXISTS sale_lines (
    id INTEGER PRIMARY KEY, tx_id INTEGER NOT NULL, kind TEXT NOT NULL, name TEXT NOT NULL, qty INTEGER NOT NULL);
`);
if (!db.prepare('PRAGMA table_info(transactions)').all().some((c) => c.name === 'sold_by'))
  db.exec("ALTER TABLE transactions ADD COLUMN sold_by TEXT NOT NULL DEFAULT ''");
if (!db.prepare('PRAGMA table_info(snack_items)').all().some((c) => c.name === 'image'))
  db.exec('ALTER TABLE snack_items ADD COLUMN image TEXT');
const UPLOADS = path.join(DATA_DIR, 'images');
fs.mkdirSync(UPLOADS, { recursive: true });
const auth = initAuth(db);
if (!db.prepare('PRAGMA table_info(transactions)').all().some((c) => c.name === 'volunteer'))
  db.exec("ALTER TABLE transactions ADD COLUMN volunteer TEXT NOT NULL DEFAULT ''");
if (!db.prepare('PRAGMA table_info(sale_lines)').all().some((c) => c.name === 'unit_cents'))
  db.exec('ALTER TABLE sale_lines ADD COLUMN unit_cents INTEGER NOT NULL DEFAULT 0');
if (!db.prepare('PRAGMA table_info(snack_items)').all().some((c) => c.name === 'category'))
  db.exec("ALTER TABLE snack_items ADD COLUMN category TEXT NOT NULL DEFAULT 'snack'");
if (db.prepare("SELECT COUNT(*) c FROM snack_items WHERE category='lunch'").get().c === 0) {
  // Prices are set later from the screen; walk-up lunches are logged at $0.00 until then.
  for (const n of ['Veg Lunch', 'Non-Veg Lunch']) db.prepare("INSERT INTO snack_items (name, price_cents, category) VALUES (?, 0, 'lunch')").run(n);
}
if (db.prepare("SELECT COUNT(*) c FROM snack_items WHERE category='snack'").get().c === 0) {
  const ins = db.prepare('INSERT INTO snack_items (name, price_cents) VALUES (?, ?)');
  // Prices are set later by the Admin / Treasurer; items sell at $0.00 until then.
  for (const n of ['Tea', 'Water Bottle', 'Samosa', 'Pizza']) ins.run(n, 0);
}

// Built-in pictures for the starting menu; the Admin can replace any of them with a photo.
for (const [name, img] of [['Tea', 'tea'], ['Water Bottle', 'water-bottle'], ['Samosa', 'samosa'], ['Pizza', 'pizza'], ['Veg Lunch', 'veg-lunch'], ['Non-Veg Lunch', 'non-veg-lunch']])
  db.prepare('UPDATE snack_items SET image=? WHERE name=? AND image IS NULL').run(`/img/${img}.svg`, name);

const canSell = (role, kind) => role === 'admin' || (role === 'snack' && kind === 'snack') || (role === 'checkin' && kind === 'lunch');
const cents = (v) => Math.round(Number(v) * 100);
const fail = (code, msg) => Object.assign(new Error(msg), { code });
const logTx = (kind, method, total, description, soldBy, volunteer = '') =>
  db.prepare('INSERT INTO transactions (kind, method, total_cents, description, sold_by, volunteer) VALUES (?,?,?,?,?,?)')
    .run(kind, method, total, description, soldBy, volunteer);
const needMethod = (m) => { if (!METHODS.includes(m)) throw fail(400, 'Payment method must be cash, zelle or card'); };

const routes = {
  'GET /api/snacks': ({ query, user }) => {
    const category = query.get('category') || 'snack';
    if (!canSell(user.role, category)) throw fail(403, 'Your login cannot do this');
    return db.prepare('SELECT * FROM snack_items WHERE active=1 AND category=? ORDER BY name').all(category);
  },
  'POST /api/snacks/:id/image': ({ params, query, raw }) => {
    const ext = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/webp': 'webp' }[query.get('type')];
    if (!ext) throw fail(400, 'Use a PNG, JPEG or WebP picture');
    if (!raw.length || raw.length > 4_000_000) throw fail(400, 'Picture must be under 4 MB');
    const id = Number(params.id);
    if (!db.prepare('SELECT 1 FROM snack_items WHERE id=?').get(id)) throw fail(404, 'Item not found');
    const name = `${id}-${Date.now()}.${ext}`;
    fs.writeFileSync(path.join(UPLOADS, name), raw);
    db.prepare('UPDATE snack_items SET image=? WHERE id=?').run(`/uploads/${name}`, id);
    return { ok: true };
  },
  'POST /api/snacks': ({ body }) => {
    const price = cents(body.price);
    if (!body.name?.trim() || !(price >= 0)) throw fail(400, 'Name and price required');
    const category = body.category === 'lunch' ? 'lunch' : 'snack';
    db.prepare('INSERT INTO snack_items (name, price_cents, category) VALUES (?,?,?)').run(body.name.trim(), price, category);
    return { ok: true };
  },
  'PUT /api/snacks/:id': ({ params, body }) => {
    const price = cents(body.price);
    if (!(price >= 0)) throw fail(400, 'Valid price required');
    db.prepare('UPDATE snack_items SET price_cents=? WHERE id=?').run(price, params.id);
    return { ok: true };
  },
  'DELETE /api/snacks/:id': ({ params }) => { db.prepare('UPDATE snack_items SET active=0 WHERE id=?').run(params.id); return { ok: true }; },

  'POST /api/sales': async ({ body, user }) => {
    const isVolunteer = body.method === 'volunteer';
    const volunteer = isVolunteer ? String(body.volunteer || '').trim().slice(0, 60) : '';
    if (!isVolunteer) needMethod(body.method);
    if (!Array.isArray(body.items) || !body.items.length) throw fail(400, 'Cart is empty');
    let total = 0; const parts = []; const lines = []; const kinds = new Set();
    for (const { id, qty } of body.items) {
      const it = db.prepare('SELECT * FROM snack_items WHERE id=?').get(id);
      if (!it || !Number.isInteger(qty) || qty < 1) throw fail(400, 'Bad cart item');
      total += it.price_cents * qty; parts.push(`${qty}x ${it.name}`); lines.push([it.name, qty, it.price_cents]); kinds.add(it.category);
    }
    if (kinds.size !== 1) throw fail(400, 'Cannot mix snacks and lunch in one sale');
    const kind = [...kinds][0];
    if (!canSell(user.role, kind)) throw fail(403, 'Your login cannot sell this');
    // Volunteer items are free: nothing is charged, but the retail value stays on each line for reporting.
    const charged = isVolunteer ? 0 : total;
    const label = isVolunteer ? `Volunteer${volunteer ? ' ' + volunteer : ''}: ` : kind === 'lunch' ? 'Walk-up: ' : '';
    const tx = logTx(kind, body.method, charged, label + parts.join(', '), user.username, volunteer);
    for (const [name, qty, unit] of lines) db.prepare('INSERT INTO sale_lines (tx_id, kind, name, qty, unit_cents) VALUES (?,?,?,?,?)').run(tx.lastInsertRowid, kind, name, qty, unit);
    // The database is the record of truth; a failed Excel write (e.g. file open elsewhere) must not undo the sale.
    const qty = lines.reduce((n, l) => n + l[1], 0), id = Number(tx.lastInsertRowid), now = localStamp();
    let excel = true;
    try {
      if (isVolunteer) {
        const row = [id, now, volunteer, parts.join(', '), qty, total / 100, ROLES[user.role]];
        await (kind === 'lunch' ? sheet.appendVolunteerMeal(row) : appendVolunteerSnack(row));
      } else {
        const row = [id, now, parts.join(', '), qty, body.method, total / 100, ROLES[user.role]];
        await (kind === 'lunch' ? sheet.appendWalkup(row) : appendSnackSale(row));
      }
    }
    catch (e) { excel = false; console.error('Excel write failed:', e.message); }
    return { ok: true, total_cents: charged, volunteer: isVolunteer, excel };
  },

  'GET /api/sheet': async () => sheet.status(),
  'GET /api/sheet/counts': async () => (await sheet.counts()) || { people: { total: 0, checkedIn: 0, notCheckedIn: 0 }, items: [] },
  'GET /api/sheet/search': ({ query }) => sheet.search(query.get('q') || ''),
  'POST /api/sheet/upload': async ({ raw }) => { await sheet.save(raw); return sheet.status(); },
  'POST /api/sheet/checkin': async ({ body, user }) => { await sheet.setCheckIn(Number(body.row), true, ROLES[user.role]); return { ok: true }; },
  'POST /api/sheet/undo': async ({ body }) => { await sheet.setCheckIn(Number(body.row), false); return { ok: true }; },

  'GET /api/me': ({ user }) => ({ username: user.username, role: user.role, label: ROLES[user.role] }),
  'POST /api/users/password': ({ body }) => { auth.changePassword(body.username, body.password); return { ok: true }; },

  'GET /api/report': ({ query }) => buildReport(db, query.get('date') || new Date().toLocaleDateString('en-CA')),

  'GET /api/transactions': ({ query }) => {
    const day = query.get('date') || new Date().toLocaleDateString('en-CA');
    const rows = db.prepare("SELECT * FROM transactions WHERE date(created_at,'localtime')=? ORDER BY id DESC").all(day);
    const totals = { cash: 0, zelle: 0, card: 0, all: 0 };
    for (const r of rows) { totals[r.method] += r.total_cents; totals.all += r.total_cents; }
    const items = db.prepare(`SELECT l.kind, l.name, SUM(l.qty) qty FROM sale_lines l JOIN transactions t ON t.id=l.tx_id
      WHERE date(t.created_at,'localtime')=? GROUP BY l.kind, l.name ORDER BY l.kind DESC, l.name`).all(day);
    return { date: day, totals, rows, items };
  },
};

// Who may use what. Anything not listed is open to every logged-in user (menu reads, sales, pictures).
// Admin may do everything; 'snack' sells snacks only; 'checkin' handles preorder check-in and walk-up lunch only.
const HOME = { admin: '/snack', snack: '/snack', checkin: '/checkin' };
const PAGE_ROLES = { 'snack.html': ['admin', 'snack'], 'checkin.html': ['admin', 'checkin'], 'report.html': ['admin'], 'settings.html': ['admin'] };
const ADMIN_ONLY = new Set(['POST /api/snacks', 'POST /api/snacks/:id/image', 'PUT /api/snacks/:id', 'DELETE /api/snacks/:id', 'POST /api/sheet/upload', 'GET /api/report', 'GET /api/transactions', 'POST /api/users/password']);
const CHECKIN_API = new Set(['GET /api/sheet', 'GET /api/sheet/counts', 'GET /api/sheet/search', 'POST /api/sheet/checkin', 'POST /api/sheet/undo']);
const ADMIN_DOWNLOADS = new Set(['/sheet.xlsx', '/snack-sales.xlsx', '/report.xlsx']);
const PUBLIC_FILES = new Set(['login.html', 'style.css', 'common.js']);
const TRUST_PROXY = process.env.TRUST_PROXY === '1';
const clientIp = (req) => (TRUST_PROXY && req.headers['x-forwarded-for']?.split(',')[0].trim()) || req.socket.remoteAddress;
const isHttps = (req) => TRUST_PROXY && req.headers['x-forwarded-proto'] === 'https';
const cookie = (req, name) => (req.headers.cookie || '').split(/;\s*/).map((c) => c.split('=')).find(([k]) => k === name)?.[1];
const mime = { '.html': 'text/html', '.css': 'text/css', '.js': 'text/javascript', '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg', '.webp': 'image/webp' };
const matchRoute = (method, pathname) => {
  for (const key of Object.keys(routes)) {
    const [m, pat] = key.split(' ');
    if (m !== method) continue;
    const names = []; const re = new RegExp('^' + pat.replace(/:(\w+)/g, (_, n) => (names.push(n), '([^/]+)')) + '$');
    const hit = pathname.match(re);
    if (hit) return { key, fn: routes[key], params: Object.fromEntries(names.map((n, i) => [n, hit[i + 1]])) };
  }
};

http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://x');
  const send = (code, data, type = 'application/json') => { res.writeHead(code, { 'Content-Type': type }); res.end(type === 'application/json' ? JSON.stringify(data) : data); };
  try {
    if (url.pathname === '/healthz') return send(200, 'ok', 'text/plain');
    const token = cookie(req, 'sid');
    const user = auth.session(token);
    const page = { '/snack': 'snack.html', '/checkin': 'checkin.html', '/report': 'report.html', '/log': 'report.html', '/settings': 'settings.html', '/login': 'login.html' }[url.pathname];
    const file = page || url.pathname.slice(1);
    const redirect = (to) => { res.writeHead(302, { Location: to }); res.end(); };

    if (url.pathname === '/api/login' && req.method === 'POST') {
      const chunks = []; for await (const c of req) chunks.push(c);
      const b = JSON.parse(Buffer.concat(chunks) || '{}');
      const { token: t, maxAge } = auth.login(clientIp(req), b.username, b.password);
      res.writeHead(200, { 'Content-Type': 'application/json', 'Set-Cookie': `sid=${t}; HttpOnly; SameSite=Strict; Path=/; Max-Age=${maxAge}${isHttps(req) ? '; Secure' : ''}` });
      return res.end('{"ok":true}');
    }
    if (url.pathname === '/api/logout' && req.method === 'POST') {
      auth.logout(token);
      res.writeHead(200, { 'Content-Type': 'application/json', 'Set-Cookie': 'sid=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0' });
      return res.end('{"ok":true}');
    }
    if (url.pathname.startsWith('/uploads/') && user) {
      const f = path.join(UPLOADS, path.basename(url.pathname));
      if (!fs.existsSync(f)) return send(404, 'Not found', 'text/plain');
      res.writeHead(200, { 'Content-Type': mime[path.extname(f)] || 'application/octet-stream', 'Cache-Control': 'private, max-age=86400' });
      return res.end(fs.readFileSync(f));
    }
    const isApi = url.pathname.startsWith('/api/');
    if (!PUBLIC_FILES.has(file) || isApi) {
      if (!user) return isApi ? send(401, { error: 'Please log in' }) : redirect('/login');
      const key = isApi ? matchRoute(req.method, url.pathname)?.key : null;
      const allowed = ADMIN_DOWNLOADS.has(url.pathname) || ADMIN_ONLY.has(key) ? ['admin'] : CHECKIN_API.has(key) ? ['admin', 'checkin'] : PAGE_ROLES[file];
      if (allowed && !allowed.includes(user.role)) return isApi || ADMIN_DOWNLOADS.has(url.pathname) ? send(403, { error: 'Your login cannot do this' }) : redirect(HOME[user.role]);
    }
    if (url.pathname === '/' ) return redirect(user ? HOME[user.role] : '/login');
    if (url.pathname === '/login' && user) return redirect(HOME[user.role]);
    const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const downloads = { '/sheet.xlsx': [sheet.SHEET_PATH, 'lunch-preorders.xlsx'], '/snack-sales.xlsx': [SNACK_PATH, 'snack-sales.xlsx'] };
    if (downloads[url.pathname]) {
      const [file, name] = downloads[url.pathname];
      if (!fs.existsSync(file)) return send(404, 'Nothing recorded yet', 'text/plain');
      res.writeHead(200, { 'Content-Type': XLSX, 'Content-Disposition': `attachment; filename="${name}"` });
      return res.end(fs.readFileSync(file));
    }
    if (url.pathname === '/report.xlsx') {
      const day = url.searchParams.get('date') || new Date().toLocaleDateString('en-CA');
      res.writeHead(200, { 'Content-Type': XLSX, 'Content-Disposition': `attachment; filename="day-report-${day}.xlsx"` });
      return res.end(Buffer.from(await reportWorkbook(await buildReport(db, day))));
    }
    if (url.pathname.startsWith('/api/')) {
      const r = matchRoute(req.method, url.pathname);
      if (!r) return send(404, { error: 'Not found' });
      const chunks = []; for await (const c of req) chunks.push(c);
      const raw = Buffer.concat(chunks);
      const isJson = (req.headers['content-type'] || '').includes('json') && raw.length;
      const out = await r.fn({ user, params: r.params, query: url.searchParams, body: isJson ? JSON.parse(raw) : {}, raw });
      return send(200, out);
    }
    const full = path.join(__dirname, 'public', file);
    if (!full.startsWith(path.join(__dirname, 'public')) || !fs.existsSync(full) || fs.statSync(full).isDirectory()) return send(404, 'Not found', 'text/plain');
    send(200, fs.readFileSync(full), mime[path.extname(full)] || 'application/octet-stream');
  } catch (e) { send(Number.isInteger(e.code) ? e.code : 500, { error: e.message }); }
}).listen(PORT, () => console.log(`Running on http://localhost:${PORT}  (login at /login)  sheet: ${sheet.SHEET_PATH}`));
