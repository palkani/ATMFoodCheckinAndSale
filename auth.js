import crypto from 'node:crypto';

const SESSION_HOURS = 12;
export const ROLES = { admin: 'Admin / Treasurer', rep: 'Sales Rep' };
const sha = (s) => crypto.createHash('sha256').update(s).digest('hex');
const hashPw = (pw, salt) => crypto.scryptSync(pw, salt, 32).toString('hex');
const WORDS = ['mango', 'lotus', 'diya', 'jasmine', 'tamil', 'kolam', 'sweet', 'lamp', 'rangoli', 'banana'];
const randomPassword = () => `${WORDS[crypto.randomInt(WORDS.length)]}-${WORDS[crypto.randomInt(WORDS.length)]}-${crypto.randomInt(100, 1000)}`;

export function initAuth(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, role TEXT NOT NULL, salt TEXT NOT NULL, hash TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS sessions (token_hash TEXT PRIMARY KEY, username TEXT NOT NULL, expires_at INTEGER NOT NULL);
  `);
  const setPassword = (username, role, pw) => {
    const salt = crypto.randomBytes(16).toString('hex');
    db.prepare('INSERT INTO users (username, role, salt, hash) VALUES (?,?,?,?) ON CONFLICT(username) DO UPDATE SET salt=excluded.salt, hash=excluded.hash')
      .run(username, role, salt, hashPw(pw, salt));
  };
  // First start: create both accounts. Passwords come from the environment or are generated and printed once.
  for (const [username, env] of [['admin', 'ADMIN_PASSWORD'], ['rep', 'REP_PASSWORD']]) {
    if (db.prepare('SELECT 1 FROM users WHERE username=?').get(username)) continue;
    const pw = process.env[env] || randomPassword();
    setPassword(username, username, pw);
    console.log(`Created ${ROLES[username]} login — password: ${pw}   (change it under Settings)`);
  }

  const failures = new Map(); // ip -> { n, until }
  return {
    login(ip, username, password) {
      const f = failures.get(ip);
      if (f && f.until > Date.now()) throw Object.assign(new Error('Too many attempts. Wait a minute and try again.'), { code: 429 });
      const u = db.prepare('SELECT * FROM users WHERE username=?').get(String(username));
      const ok = u && crypto.timingSafeEqual(Buffer.from(hashPw(String(password), u.salt), 'hex'), Buffer.from(u.hash, 'hex'));
      if (!ok) {
        const expired = f && f.until && f.until <= Date.now();
        const n = (expired ? 0 : f?.n || 0) + 1;
        failures.set(ip, { n, until: n >= 5 ? Date.now() + 60_000 : 0 });
        throw Object.assign(new Error('Wrong password'), { code: 401 });
      }
      failures.delete(ip);
      const token = crypto.randomBytes(32).toString('hex');
      db.prepare('DELETE FROM sessions WHERE expires_at < ?').run(Date.now());
      db.prepare('INSERT INTO sessions (token_hash, username, expires_at) VALUES (?,?,?)').run(sha(token), u.username, Date.now() + SESSION_HOURS * 3600_000);
      return { token, maxAge: SESSION_HOURS * 3600 };
    },
    session(token) {
      if (!token) return null;
      return db.prepare('SELECT u.username, u.role FROM sessions s JOIN users u ON u.username=s.username WHERE s.token_hash=? AND s.expires_at>?').get(sha(token), Date.now()) || null;
    },
    logout(token) { if (token) db.prepare('DELETE FROM sessions WHERE token_hash=?').run(sha(token)); },
    changePassword(username, pw) {
      const u = db.prepare('SELECT role FROM users WHERE username=?').get(username);
      if (!u) throw Object.assign(new Error('Unknown user'), { code: 404 });
      if (String(pw).length < 6) throw Object.assign(new Error('Password must be at least 6 characters'), { code: 400 });
      setPassword(username, u.role, String(pw));
      db.prepare('DELETE FROM sessions WHERE username=?').run(username);
    },
  };
}
