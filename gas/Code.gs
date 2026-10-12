/**
 * Sales Ledger backend (Google Apps Script, bound to a Google Sheet).
 * No data and no passwords in this file. Data lives in the sheet; run setup() once (see Seed.gs).
 * Deploy as Web app: Execute as = Me, Who has access = Anyone. Every call still needs a login token.
 */
const SECS_ = ['sales', 'expense', 'cash', 'card', 'transfer', 'disc'];
const P_ = () => PropertiesService.getScriptProperties();
const SHEET_ = n => SpreadsheetApp.openById(P_().getProperty('SSID')).getSheetByName(n);
const HEAD_ = {
  entries: ['id', 'br', 'date', 'sec', 'cat', 'amt', 'ref', 'note', 'updated_at', 'updated_by'],
  cats: ['br', 'sec', 'label', 'counted', 'grp'],
  meta: ['br', 'json'],
  users: ['username', 'name', 'role', 'br', 'must_change', 'salt', 'hash']
};

function doGet() { return out_({ ok: true, service: 'sales-ledger' }); }

function doPost(e) {
  let r;
  try {
    const b = JSON.parse(e.postData.contents);
    r = route_(b);
  } catch (err) {
    r = { error: String(err && err.message || err) };
  }
  return out_(r);
}
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function route_(b) {
  if (b.a === 'login') return login_(b);
  const me = auth_(b.t);
  if (b.a === 'load') return load_(me);
  if (b.a === 'sync') return withLock_(() => sync_(me, b));
  if (b.a === 'chpw') return withLock_(() => chpw_(me, b));
  throw new Error('Unknown action');
}
function withLock_(fn) {
  const l = LockService.getScriptLock(); l.waitLock(25000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ---------- sheet helpers ---------- */
function read_(name) {
  const sh = SHEET_(name), n = sh.getLastRow(), h = HEAD_[name];
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, h.length).getValues().map(v => { const o = {}; h.forEach((k, i) => o[k] = v[i]); return o; });
}
const canSee_ = (me, br) => me.role === 'admin' || me.br === br;

/* ---------- auth ---------- */
const hex_ = bytes => bytes.map(x => ('0' + (x & 255).toString(16)).slice(-2)).join('');
function hashPw_(pw, salt) {
  let h = salt + pw;
  for (let i = 0; i < 200; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}
function sign_(s) { return hex_(Utilities.computeHmacSha256Signature(s, P_().getProperty('SECRET'))); }
function userRow_(username) {
  const sh = SHEET_('users'), n = sh.getLastRow();
  if (n < 2) return null;
  const vals = sh.getRange(2, 1, n - 1, 7).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).toLowerCase() === username) {
    const v = vals[i];
    return { row: i + 2, username: v[0], name: v[1], role: v[2], br: v[3] || null, must_change: v[4] === true || v[4] === 'TRUE', salt: v[5], hash: v[6] };
  }
  return null;
}
function token_(u) {
  const payload = u.username + '|' + (Date.now() + 30 * 864e5) + '|' + String(u.hash).slice(0, 10);
  return Utilities.base64EncodeWebSafe(payload) + '.' + sign_(payload);
}
function auth_(t) {
  if (!t || typeof t !== 'string' || t.indexOf('.') < 0) throw new Error('AUTH');
  const [p64, sig] = t.split('.');
  let payload;
  try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(p64)).getDataAsString(); } catch (e) { throw new Error('AUTH'); }
  if (sign_(payload) !== sig) throw new Error('AUTH');
  const [name, exp, hp] = payload.split('|');
  if (+exp < Date.now()) throw new Error('AUTH');
  const u = userRow_(name);
  if (!u || String(u.hash).slice(0, 10) !== hp) throw new Error('AUTH');
  return u;
}
function pub_(u) { return { id: u.username, username: u.username, name: u.name, role: u.role, br: u.br, must_change: u.must_change }; }
function login_(b) {
  const name = String(b.u || '').trim().toLowerCase(), cache = CacheService.getScriptCache(), k = 'fail:' + name;
  if (+(cache.get(k) || 0) >= 8) throw new Error('Too many attempts. Try again in 15 minutes.');
  const u = name ? userRow_(name) : null;
  if (!u || hashPw_(String(b.p || ''), u.salt) !== u.hash) {
    cache.put(k, String(+(cache.get(k) || 0) + 1), 900);
    Utilities.sleep(800);
    throw new Error('Wrong user ID or password.');
  }
  cache.remove(k);
  return { token: token_(u), me: pub_(u) };
}
function chpw_(me, b) {
  const np = String(b.np || '');
  if (np.length < 8) throw new Error('Use at least 8 characters.');
  if (np === '1234') throw new Error('Choose a different password.');
  const salt = Utilities.getUuid(), hash = hashPw_(np, salt);
  SHEET_('users').getRange(me.row, 5, 1, 3).setValues([[false, salt, hash]]);
  return { token: token_({ username: me.username, hash }), me: pub_({ ...me, must_change: false }) };
}

/* ---------- data ---------- */
function load_(me) {
  const users = read_('users').map(v => ({ username: v.username, name: v.name, role: v.role, br: v.br || null, must_change: v.must_change === true || v.must_change === 'TRUE' }));
  const profiles = (me.role === 'admin' ? users : users.filter(x => x.username === me.username)).map(x => ({ id: x.username, ...x }));
  return {
    me: pub_(me), profiles,
    meta: read_('meta').filter(r => canSee_(me, r.br)).map(r => ({ br: r.br, data: JSON.parse(r.json) })),
    cats: read_('cats').filter(r => canSee_(me, r.br)).map(r => ({ br: r.br, sec: r.sec, label: r.label, counted: r.counted === true || r.counted === 'TRUE', grp: String(r.grp || '') })),
    entries: read_('entries').filter(r => canSee_(me, r.br)).map(r => ({ id: r.id, br: r.br, date: String(r.date), sec: r.sec, cat: String(r.cat), amt: Number(r.amt), ref: String(r.ref), note: String(r.note) }))
  };
}
function sync_(me, b) {
  const up = b.up || [], upC = b.upC || [], del = b.del || [];
  if (up.length > 2000 || upC.length > 1000 || del.length > 2000) throw new Error('Too many changes at once');
  const who = me.username, now = new Date().toISOString();
  for (const e of up) {
    if (!canSee_(me, e.br)) throw new Error('Not allowed for branch ' + e.br);
    if (!SECS_.includes(e.sec) || !isFinite(e.amt) || !/^\d{4}-\d{2}-\d{2}$/.test(e.date) || !isFinite(e.id) || !String(e.cat || '').trim()) throw new Error('Invalid entry');
  }
  for (const c of upC) if (!canSee_(me, c.br) || !SECS_.includes(c.sec) || !String(c.label || '').trim()) throw new Error('Invalid category');

  // categories
  if (upC.length) {
    const sh = SHEET_('cats'), cur = read_('cats'), idx = {};
    cur.forEach((r, i) => idx[r.br + '|' + r.sec + '|' + r.label] = i + 2);
    const add = [];
    for (const c of upC) {
      const row = idx[c.br + '|' + c.sec + '|' + c.label];
      if (row) sh.getRange(row, 4, 1, 2).setValues([[!!c.counted, String(c.grp || '')]]); else add.push([c.br, c.sec, c.label, !!c.counted, String(c.grp || '')]);
    }
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 5).setValues(add);
    sh.getRange(1, 5).setValue('grp');
  }
  // entries
  const sh = SHEET_('entries'), n = sh.getLastRow();
  const col = n > 1 ? sh.getRange(2, 1, n - 1, 2).getValues() : [];   // id, br
  const rowOf = {}, brOf = {};
  col.forEach((v, i) => { rowOf[v[0]] = i + 2; brOf[v[0]] = v[1]; });
  const add = [];
  for (const e of up) {
    const vals = [e.id, e.br, e.date, e.sec, e.cat, Math.round(e.amt * 100) / 100, e.ref || '', e.note || '', now, who];
    if (rowOf[e.id]) {
      if (!canSee_(me, brOf[e.id])) throw new Error('Not allowed');
      sh.getRange(rowOf[e.id], 1, 1, 10).setValues([vals]);
    } else add.push(vals);
  }
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 10).setValues(add);
  const rows = del.filter(id => rowOf[id] && canSee_(me, brOf[id])).map(id => rowOf[id]).sort((a, b) => b - a);
  rows.forEach(r => sh.deleteRow(r));
  return { ok: true, up: up.length, del: rows.length };
}

/* ---------- one-time setup (run from the editor; needs Seed.gs) ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  P_().setProperty('SSID', ss.getId());
  if (!P_().getProperty('SECRET')) P_().setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  for (const name of Object.keys(HEAD_)) {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEAD_[name].length).setValues([HEAD_[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (first && ss.getSheets().length > 1) ss.deleteSheet(first);
  const E = SHEET_('entries'), C = SHEET_('cats'), M = SHEET_('meta');
  E.getRange('C:C').setNumberFormat('@'); E.getRange('E:E').setNumberFormat('@'); E.getRange('G:H').setNumberFormat('@'); E.getRange('A:A').setNumberFormat('0');
  C.getRange('A:C').setNumberFormat('@');
  if (E.getLastRow() < 2) {
    const m = Object.keys(SEED.meta).map(br => [br, JSON.stringify(SEED.meta[br])]);
    M.getRange(2, 1, m.length, 2).setValues(m);
    const cats = []; Object.keys(SEED.cats).forEach(br => SEED.cats[br].forEach(c => cats.push([br, c.sec, c.label, c.counted])));
    C.getRange(2, 1, cats.length, 4).setValues(cats);
    const now = new Date().toISOString();
    const ents = SEED.entries.map(e => [e.id, e.br, e.date, e.sec, e.cat, e.amt, e.ref || '', e.note || '', now, 'import']);
    E.getRange(2, 1, ents.length, 10).setValues(ents);
  }
  SHEET_('users').getRange('A:G').setNumberFormat('@');
  if (SHEET_('users').getLastRow() < 2) resetUsers();
  Logger.log('Setup done. Entries: ' + (E.getLastRow() - 1));
}
/** Re-creates all accounts from SEED_USERS (data is kept). Also the way to recover a forgotten password. */
function resetUsers() {
  const sh = SHEET_('users');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
  const rows = SEED_USERS.map(u => { const salt = Utilities.getUuid(); return [u.username, u.name, u.role, u.br || '', u.must_change, salt, hashPw_(u.pw, salt)]; });
  sh.getRange(2, 1, rows.length, 7).setValues(rows);
  sh.getRange(2, 5, rows.length, 1).setNumberFormat('General');
  sh.getRange(2, 5, rows.length, 1).setValues(rows.map(r => [r[4]]));
}
