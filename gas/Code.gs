/**
 * Sales Ledger backend (Google Apps Script, bound to a Google Sheet).
 * No data and no passwords in this file. Data lives in the sheet; run setup() once (see Seed.gs).
 * Deploy as Web app: Execute as = Me, Who has access = Anyone. Every call still needs a login token.
 */
const SECS_ = ['sales', 'expense', 'cash', 'card', 'transfer', 'disc'];
const P_ = () => PropertiesService.getScriptProperties();
const SHEET_ = n => {
  const ss = SpreadsheetApp.openById(P_().getProperty('SSID'));
  let sh = ss.getSheetByName(n);
  if (!sh && HEAD_[n]) { sh = ss.insertSheet(n); sh.getRange(1, 1, 1, HEAD_[n].length).setValues([HEAD_[n]]).setFontWeight('bold'); sh.setFrozenRows(1); if (n === 'wages' || n === 'deductions') sh.getRange('A:A').setNumberFormat('@'); if (n === 'deductions') sh.getRange('E:E').setNumberFormat('@'); }
  return sh;
};
const HEAD_ = {
  entries: ['id', 'br', 'date', 'sec', 'cat', 'amt', 'ref', 'note', 'updated_at', 'updated_by', 'file'],
  cats: ['br', 'sec', 'label', 'counted', 'grp'],
  meta: ['br', 'json'],
  users: ['username', 'name', 'role', 'br', 'must_change', 'salt', 'hash'],
  items: ['entry', 'br', 'date', 'supplier', 'item', 'qty', 'amt', 'cat'],
  wages: ['week', 'br', 'name', 'nick', 'mode', 'award', 'restricted', 'ordH', 'ordR', 'satH', 'satR', 'sunH', 'sunR', 'otH', 'otR', 'dayN', 'dayR', 'stp', 'stpReal', 'paid', 'note', 'realH', 'cashH', 'cashR', 'rent'],
  deductions: ['week', 'br', 'name', 'amount', 'apply_week', 'status', 'note'],
  checks: ['entry', 'br', 'amt', 'checked_by', 'checked_at', 'phys', 'digital', 'acc'],
  payees: ['br', 'supplier', 'account_name', 'bsb', 'acct', 'ref_mode', 'fixed_ref', 'bpay_biller', 'bpay_crn', 'note'],
  payers: ['entity', 'branches', 'bsb', 'acct']
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
  if (b.a === 'chk') return withLock_(() => chk_(me, b));
  if (b.a === 'inv') return withLock_(() => inv_(me, b));
  if (b.a === 'chpw') return withLock_(() => chpw_(me, b));
  if (b.a === 'inv_read') return invRead_(me, b);
  if (b.a === 'inv_img') return invImg_(me, b);
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
    meta: metaFor_(me),
    cats: read_('cats').filter(r => canSee_(me, r.br)).map(r => ({ br: r.br, sec: r.sec, label: r.label, counted: r.counted === true || r.counted === 'TRUE', grp: String(r.grp || '') })),
    entries: read_('entries').filter(r => canSee_(me, r.br)).map(r => ({ id: r.id, br: r.br, date: String(r.date), sec: r.sec, cat: String(r.cat), amt: Number(r.amt), ref: String(r.ref), note: String(r.note), file: String(r.file || '') })),
    items: itemsOf_(me),
    pay: me.role === 'admin' ? payInfo_() : null,
    checks: me.role === 'admin' ? read_('checks').map(r => ({ id: Number(r.entry), amt: r.amt === '' ? null : Number(r.amt), phys: r.phys === true || r.phys === 'TRUE', digital: r.digital === '' || r.digital === null ? null : Number(r.digital), acc: r.acc === true || r.acc === 'TRUE' })) : [],
    wages: wagesOf_(me), hiddenWage: hiddenWage_(me), deductions: me.role === 'admin' ? read_('deductions').map(d => ({ week: String(d.week), br: String(d.br), name: String(d.name), amount: Number(d.amount), apply_week: String(d.apply_week), status: String(d.status), note: String(d.note) })) : []
  };
}
/* Supplier bank details and the owner's paying accounts (admin only). The sheet is the only place they live; they are never in this file or the public repo. */
function payInfo_() {
  const s = v => String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim();
  return {
    payees: read_('payees').filter(r => s(r.supplier)).map(r => ({ br: s(r.br), supplier: s(r.supplier), name: s(r.account_name), bsb: s(r.bsb), acct: s(r.acct), mode: s(r.ref_mode), ref: s(r.fixed_ref), biller: s(r.bpay_biller), crn: s(r.bpay_crn), note: s(r.note) })),
    payers: read_('payers').filter(r => s(r.entity)).map(r => ({ entity: s(r.entity), branches: s(r.branches).split(',').map(x => x.trim()).filter(Boolean), bsb: s(r.bsb), acct: s(r.acct) }))
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
  for (const e of up) if (e.file && !/^[\w-]{10,}$/.test(String(e.file))) throw new Error('Invalid file');
  const itemsUp = b.items || [];
  for (const it of itemsUp) if (!canSee_(me, it.br) || !isFinite(it.entry) || !Array.isArray(it.lines) || it.lines.length > 200) throw new Error('Invalid items');
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
    const vals = [e.id, e.br, e.date, e.sec, e.cat, Math.round(e.amt * 100) / 100, e.ref || '', e.note || '', now, who, e.file || ''];
    if (rowOf[e.id]) {
      if (!canSee_(me, brOf[e.id])) throw new Error('Not allowed');
      sh.getRange(rowOf[e.id], 1, 1, 11).setValues([vals]);
    } else add.push(vals);
  }
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 11).setValues(add);
  if (up.some(e => e.file)) sh.getRange(1, 11).setValue('file');
  const rows = del.filter(id => rowOf[id] && canSee_(me, brOf[id])).map(id => rowOf[id]).sort((a, b) => b - a);
  rows.forEach(r => sh.deleteRow(r));
  if (itemsUp.length || del.length) saveItems_(me, itemsUp, del);
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

/** Adds (or resets) the admin account "danny". A random temporary password is written to the execution log;
 *  danny must change it at first login. Existing accounts and data are untouched. */
function addAdminDanny() {
  const sh = SHEET_('users');
  const pw = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  const salt = Utilities.getUuid();
  const row = ['danny', 'Danny', 'admin', '', true, salt, hashPw_(pw, salt)];
  const n = sh.getLastRow();
  let at = 0;
  if (n > 1) sh.getRange(2, 1, n - 1, 1).getValues().forEach((v, i) => { if (v[0] === 'danny') at = i + 2; });
  if (!at) at = n + 1;
  sh.getRange(at, 1, 1, 7).setNumberFormat('@').setValues([row]);
  sh.getRange(at, 5).setNumberFormat('General').setValue(true);
  Logger.log('danny created. Temporary password: ' + pw);
}

/* ---------- invoice photos + AI reading (Drive + Gemini) ---------- */
const ITEM_CATS_ = ['Grocery', 'Drink', 'Liquor', 'Equipment', 'Other'];

function itemsOf_(me) {
  const by = {}, order = [];
  read_('items').forEach(r => {
    if (!canSee_(me, r.br)) return;
    const k = String(r.entry);
    if (!by[k]) { by[k] = { entry: Number(r.entry), br: r.br, date: String(r.date), supplier: String(r.supplier), lines: [] }; order.push(k); }
    by[k].lines.push({ item: String(r.item), qty: r.qty === '' ? null : Number(r.qty), amt: Number(r.amt), cat: String(r.cat) });
  });
  return order.map(k => by[k]);
}
function saveItems_(me, itemsUp, delIds) {
  const sh = SHEET_('items'), n = sh.getLastRow();
  const drop = {};
  (delIds || []).forEach(id => drop[id] = 1);
  itemsUp.forEach(it => drop[it.entry] = 1);
  if (n > 1) {
    const col = sh.getRange(2, 1, n - 1, 2).getValues();   // entry, br
    const rows = [];
    col.forEach((v, i) => { if (drop[v[0]] && canSee_(me, v[1])) rows.push(i + 2); });
    rows.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  }
  const add = [];
  itemsUp.forEach(it => it.lines.forEach(l => add.push([it.entry, it.br, it.date, String(it.supplier || ''), String(l.item || '').slice(0, 200), l.qty === null || l.qty === undefined || l.qty === '' ? '' : Number(l.qty), Math.round(Number(l.amt || 0) * 100) / 100, ITEM_CATS_.includes(l.cat) ? l.cat : 'Other'])));
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 8).setValues(add);
}
function folder_(br) {
  const p = P_(), id = p.getProperty('INV_' + br);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  let root = null;
  const rid = p.getProperty('INV_ROOT');
  if (rid) { try { root = DriveApp.getFolderById(rid); } catch (e) {} }
  if (!root) { root = DriveApp.createFolder('Sales Ledger Invoices'); p.setProperty('INV_ROOT', root.getId()); }
  const f = root.createFolder(br);
  p.setProperty('INV_' + br, f.getId());
  return f;
}
const EXT_ = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'application/pdf': 'pdf' };
function invRead_(me, b) {
  const br = String(b.br || '');
  if (!canSee_(me, br) || !/^[A-Z]{2}$/.test(br)) throw new Error('Not allowed for branch ' + br);
  const mime = String(b.mime || ''), data = String(b.data || '');
  if (!EXT_[mime]) throw new Error('Use a JPG, PNG or PDF file.');
  if (data.length > 9e6) throw new Error('The file is too large. Take the photo again or use a smaller file.');
  const stamp = Utilities.formatDate(new Date(), 'Australia/Sydney', 'yyyyMMdd-HHmmss');
  const file = folder_(br).createFile(Utilities.newBlob(Utilities.base64Decode(data), mime, br + '_' + stamp + '.' + EXT_[mime]));
  file.setDescription(br);
  const out = { file: { id: file.getId(), name: file.getName() }, parsed: null, perr: '' };
  try { out.parsed = gemini_(mime, data); } catch (err) { out.perr = String(err && err.message || err); }
  if (out.parsed) {
    const p = out.parsed, clean = x => String(x || '').replace(/[^\w.-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 30);
    const nm = [br, p.date || stamp, clean(p.supplier), clean(p.invoice_no)].filter(Boolean).join('_') + '.' + EXT_[mime];
    try { file.setName(nm); out.file.name = nm; } catch (e) {}
  }
  return out;
}
function invImg_(me, b) {
  const id = String(b.id || '');
  if (!/^[\w-]{10,}$/.test(id)) throw new Error('Invalid file');
  const f = DriveApp.getFileById(id), br = String(f.getDescription() || '');
  if (!/^[A-Z]{2}$/.test(br) || !canSee_(me, br)) throw new Error('Not allowed');
  return { mime: f.getMimeType(), name: f.getName(), data: Utilities.base64Encode(f.getBlob().getBytes()) };
}
function models_(key) {
  const cache = CacheService.getScriptCache(), pref = P_().getProperty('GEMINI_MODEL');
  let list = cache.get('gem_models');
  list = list ? JSON.parse(list) : null;
  if (!list) {
    list = [];
    try {
      const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
      if (r.getResponseCode() === 200) {
        list = (JSON.parse(r.getContentText()).models || [])
          .filter(m => (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0)
          .map(m => String(m.name).replace(/^models\//, ''))
          .filter(n => /flash/.test(n) && !/image|tts|audio|live|embed|robotics|computer|native|dialog/.test(n))
          .sort((a, b) => (/preview|exp/.test(a) ? 1 : 0) - (/preview|exp/.test(b) ? 1 : 0) || (/lite/.test(a) ? 1 : 0) - (/lite/.test(b) ? 1 : 0) || (a < b ? 1 : -1))
          .slice(0, 5);
      }
    } catch (e) {}
    if (!list.length) list = ['gemini-2.5-flash', 'gemini-2.5-flash-lite'];
    cache.put('gem_models', JSON.stringify(list), 21600);
  }
  return pref ? [pref].concat(list.filter(x => x !== pref)) : list;
}
function gemini_(mime, b64) {
  const key = P_().getProperty('GEMINI_KEY');
  if (!key) throw new Error('Invoice reading is not set up yet.');
  const T = { type: 'STRING' }, N = { type: 'NUMBER', nullable: true };
  const body = {
    contents: [{ parts: [{ text:
      'You read supplier invoices and shop receipts for a sushi restaurant business in Australia. Copy the data exactly as printed; never guess. ' +
      'Dates are Australian DD/MM/YYYY; output the date as YYYY-MM-DD (empty string if not visible). Amounts are AUD. "total" is the final amount payable including GST. ' +
      '"items" has one entry per printed product line (name cleaned up but recognisable, quantity if shown, amount = line total as printed). ' +
      'Give every item a category: Grocery (food ingredients, fish, rice, vegetables, sauces, supermarket goods), Drink (non-alcoholic drinks), Liquor (alcohol), Equipment (utensils, packaging, cleaning supplies, appliances, repairs), or Other. ' +
      'Skip subtotal, GST, discount and payment lines (do not list them as items).' }, { inline_data: { mime_type: mime, data: b64 } }] }],
    generationConfig: {
      temperature: 0, responseMimeType: 'application/json',
      responseSchema: { type: 'OBJECT', properties: {
        supplier: T, date: T, invoice_no: T, total: N,
        items: { type: 'ARRAY', items: { type: 'OBJECT', properties: { name: T, qty: N, amount: N, category: { type: 'STRING', enum: ITEM_CATS_ } }, required: ['name', 'amount', 'category'] } }
      }, required: ['supplier', 'date', 'invoice_no', 'total', 'items'] }
    }
  };
  let last = '', limited = false;
  for (const m of models_(key)) {
    const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body), muteHttpExceptions: true });
    const code = r.getResponseCode(), txt = r.getContentText();
    if (code === 200) {
      let o;
      try { o = JSON.parse(JSON.parse(txt).candidates[0].content.parts[0].text); } catch (e) { last = 'unreadable answer'; continue; }
      P_().setProperty('GEMINI_MODEL', m);
      return norm_(o);
    }
    if (code === 429) limited = true;
    try { last = code + ' ' + JSON.parse(txt).error.message; } catch (e) { last = String(code); }
  }
  if (limited) throw new Error('The free reading limit is used up for now. Enter this invoice by hand, or try again later.');
  throw new Error('Could not read this invoice (' + String(last).slice(0, 140) + ').');
}
function norm_(o) {
  const num = x => (x === null || x === undefined || x === '' || !isFinite(x)) ? null : Math.round(Number(x) * 100) / 100;
  let date = String(o.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) date = '';
  return {
    supplier: String(o.supplier || '').trim().slice(0, 80), date: date, invoice_no: String(o.invoice_no || '').trim().slice(0, 40), total: num(o.total),
    items: (Array.isArray(o.items) ? o.items : []).slice(0, 150).map(i => ({ name: String(i.name || '').trim().slice(0, 120), qty: num(i.qty), amt: num(i.amount) || 0, cat: ITEM_CATS_.indexOf(i.category) >= 0 ? i.category : 'Other' })).filter(i => i.name)
  };
}

/** Run once from the editor to grant Drive + external-request permissions. */
function authorize() { DriveApp.getRootFolder(); UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=1', { headers: { 'x-goog-api-key': P_().getProperty('GEMINI_KEY') || '' }, muteHttpExceptions: true }); Logger.log('authorized'); }

/* ---------- wages: restricted staff are only ever sent to admins ---------- */
const isRestricted_ = r => r.restricted === true || r.restricted === 'TRUE';
function wagesOf_(me) {
  return read_('wages').filter(r => canSee_(me, r.br) && (me.role === 'admin' || !isRestricted_(r))).map(r => {
    const o = { week: String(r.week), br: String(r.br), name: String(r.name), nick: String(r.nick), mode: String(r.mode), award: String(r.award), restricted: isRestricted_(r), note: String(r.note) };
    ['ordH', 'ordR', 'satH', 'satR', 'sunH', 'sunR', 'otH', 'otR', 'dayN', 'dayR', 'stp', 'stpReal', 'paid', 'realH', 'cashH', 'cashR', 'rent'].forEach(k => o[k] = (r[k] === '' || r[k] === undefined || r[k] === null) ? null : Number(r[k]));
    return o;
  });
}
/** Managers only get the combined cost of their restricted staff (needed for wage % of sales), never names or amounts per person. */
function hiddenWage_(me) {
  if (me.role === 'admin') return [];
  const by = {};
  read_('wages').filter(r => r.br === me.br && isRestricted_(r)).forEach(r => { const k = r.br + '|' + r.week; by[k] = (by[k] || 0) + Number(r.stp || 0) + (Number(r.cashH || 0) && Number(r.cashR || 0) ? Number(r.cashH) * Number(r.cashR) : 0) + Number(r.rent || 0); });
  return Object.keys(by).map(k => ({ br: k.split('|')[0], week: k.split('|')[1], cost: Math.round(by[k] * 100) / 100 }));
}

/** Legacy Excel wage rows: managers must not see restricted staff (names or amounts). */
function metaFor_(me) {
  const rows = read_('meta').filter(r => canSee_(me, r.br)).map(r => ({ br: r.br, data: JSON.parse(r.json) }));
  if (me.role === 'admin') return rows;
  const bad = {};
  read_('wages').filter(isRestricted_).forEach(r => { bad[String(r.name).trim().toLowerCase()] = 1; });
  rows.forEach(r => { if (r.data && r.data.wages) r.data.wages = r.data.wages.filter(w => !bad[String(w.name).trim().toLowerCase()]); });
  return rows;
}

/** Admin-only invoice check record (one row per invoice in the checks sheet):
 *  amt = the sheet amount that was verified (tick), phys = physical invoice present, digital = amount on the supplier's digital invoice,
 *  acc = a difference between sheet and digital amount was accepted. Fields not sent are left as they are. */
function inv_(me, b) {
  if (me.role !== 'admin') throw new Error('Only admins can check invoices');
  const id = Number(b.id);
  if (!isFinite(id)) throw new Error('Invalid entry');
  const E = SHEET_('entries'), n = E.getLastRow(), ev = n > 1 ? E.getRange(2, 1, n - 1, 6).getValues() : [];
  const ent = ev.find(v => Number(v[0]) === id);
  if (!ent) throw new Error('Entry not found');
  const sh = SHEET_('checks'), m = sh.getLastRow(), rows = [];
  sh.getRange(1, 6, 1, 3).setValues([['phys', 'digital', 'acc']]);
  let cur = ['', '', '', '', '', '', '', ''];
  if (m > 1) sh.getRange(2, 1, m - 1, 8).getValues().forEach((v, i) => { if (Number(v[0]) === id) { rows.push(i + 2); if (rows.length === 1) cur = v; } });
  const entAmt = Number(ent[5]);
  let amt = cur[2], by = cur[3], at = cur[4], phys = cur[5] === true || cur[5] === 'TRUE', digital = cur[6], acc = cur[7] === true || cur[7] === 'TRUE';
  if ('phys' in b) phys = !!b.phys;
  if ('acc' in b) acc = !!b.acc;
  if ('digital' in b) {
    if (b.digital === null || b.digital === '') { digital = ''; acc = false; }
    else {
      const d = Math.round(Number(b.digital) * 100) / 100;
      if (!isFinite(d) || d < 0) throw new Error('Invalid amount');
      digital = d; if (Math.abs(d - entAmt) < 0.005) { b.tick = true; acc = false; } else { b.tick = false; }
    }
  }
  if ('tick' in b) { if (b.tick) { amt = entAmt; by = me.username; at = new Date().toISOString(); } else { amt = ''; } }
  rows.sort((x, y) => y - x).forEach(r => sh.deleteRow(r));
  if (amt !== '' || phys || digital !== '' || acc) sh.getRange(sh.getLastRow() + 1, 1, 1, 8).setValues([[id, ent[1], amt, by || me.username, at || new Date().toISOString(), phys, digital, acc]]);
  return { ok: true };
}
/** Older clients: invoice tick only. */
function chk_(me, b) { return inv_(me, { id: b.id, tick: !!b.on }); }
