/**
 * Last Seen: penghubung Google Sheet.
 * Tab yang dibuat otomatis:
 *   Skor      papan skor Kasus Harian
 *   Event     statistik permainan (mulai, tamat, gagal, ending, pilihan, harian)
 *   Ulasan    ulasan dan rating dari pemain
 *   Pemain    akun pemain (nama pengguna + PIN) untuk menyimpan progres
 *   Ringkasan angka utama dengan rumus, bisa dibaca langsung di Sheet
 *
 * Cara pakai:
 * 1. Tempel seluruh isi file ini di Apps Script.
 * 2. Jalankan setup() sekali.
 * 3. Project Settings > Script Properties: tambah ADMIN_KEY berisi kata sandi dasbor admin.
 * 4. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */

const TAB_SKOR = 'Skor', TAB_EVENT = 'Event', TAB_ULASAN = 'Ulasan', TAB_RINGKAS = 'Ringkasan', TAB_PEMAIN = 'Pemain';
const HEAD_PEMAIN = ['Nama pengguna', 'Salt', 'PIN hash', 'Token', 'Progres', 'Dibuat', 'Terakhir aktif', 'Gagal', 'Terkunci sampai'];
const P = { U: 0, SALT: 1, HASH: 2, TOK: 3, DATA: 4, MADE: 5, SEEN: 6, FAIL: 7, LOCK: 8 };
const HEAD_SKOR = ['Waktu', 'Kasus', 'Nama', 'Bintang', 'Detik', 'Salah', 'ID Pemain'];
const HEAD_EVENT = ['Waktu', 'Jenis', 'Bab', 'Mode', 'Bintang', 'Detail', 'ID Pemain', 'Bahasa'];
const HEAD_ULASAN = ['Waktu', 'Bab', 'Rating', 'Ulasan', 'ID Pemain', 'Bahasa', 'Status'];
const LEVELS = ['prolog', 'bab1', 'bab2', 'bab3', 'bab4', 'bab5'];

/* ---------- Setup ---------- */
function setup() {
  tab_(TAB_SKOR, HEAD_SKOR);
  tab_(TAB_EVENT, HEAD_EVENT);
  tab_(TAB_ULASAN, HEAD_ULASAN);
  tab_(TAB_PEMAIN, HEAD_PEMAIN);
  const r = tab_(TAB_RINGKAS, ['Metrik', 'Jumlah']);
  const rows = [
    ['Pemain unik', '=COUNTUNIQUE(Event!G2:G)'],
    ['Mulai Prolog', '=COUNTIFS(Event!B2:B,"start",Event!C2:C,"prolog")'],
    ['Tamat Bab 1', '=COUNTIFS(Event!B2:B,"finish",Event!C2:C,"bab1")'],
    ['Tamat Bab 5', '=COUNTIFS(Event!B2:B,"finish",Event!C2:C,"bab5")'],
    ['Ending Pulang', '=COUNTIFS(Event!B2:B,"ending",Event!F2:F,"home")'],
    ['Ending Dua Rumah', '=COUNTIFS(Event!B2:B,"ending",Event!F2:F,"two")'],
    ['Ending Terlambat', '=COUNTIFS(Event!B2:B,"fail",Event!F2:F,"late")'],
    ['Kasus Harian dimainkan', '=COUNTIF(Event!B2:B,"daily")'],
    ['Jumlah ulasan', '=COUNTA(Ulasan!C2:C)'],
    ['Rata-rata rating', '=IFERROR(ROUND(AVERAGE(Ulasan!C2:C),2),"-")'],
    ['Akun pemain', '=COUNTA(Pemain!A2:A)']
  ];
  r.getRange(2, 1, rows.length, 2).setValues(rows);
  r.autoResizeColumn(1);
}

/* ---------- Pintu masuk ---------- */
function doGet() {
  return json_({ ok: true, app: 'last-seen' });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false }); }
  switch (d.a) {
    case 'score': return json_(score_(d));
    case 'top': return json_(top_(parseInt(d.no, 10), String(d.id || '')));
    case 'ev': event_(d); return json_({ ok: true });
    case 'review': return json_(review_(d));
    case 'cfg': return json_(cfg_());
    case 'acc_new': return json_(accNew_(d));
    case 'acc_login': return json_(accLogin_(d));
    case 'acc_save': return json_(accSave_(d));
    case 'acc_load': return json_(accLoad_(d));
    case 'admin': return json_(admin_(d));
  }
  return json_({ ok: false });
}

/* ---------- Publik: skor, statistik, ulasan, pengumuman ---------- */
function score_(d) {
  const no = parseInt(d.no, 10), s = parseInt(d.s, 10), t = Math.round(Number(d.t)), w = parseInt(d.w, 10) || 0;
  const n = clean_(d.n, 12), id = clean_(d.id, 40);
  if (!n || !id || !(no > 0) || !(s >= 0 && s <= 3) || !(t >= 2 && t <= 3600)) return { ok: false };
  if (banned_().indexOf(n.toLowerCase()) >= 0) return { ok: false, err: 'name' };
  const owner = findUser_(n.toLowerCase());
  if (owner && !(String(d.u || '') === n.toLowerCase() && tokenOk_(owner.row, String(d.tok || '')))) return { ok: false, err: 'name' };
  return withLock_(() => {
    const sh = tab_(TAB_SKOR, HEAD_SKOR);
    const rows = sh.getDataRange().getValues();
    if (rows.some(r => r[1] == no && r[6] == id)) return { ok: true, dup: true };
    sh.appendRow([new Date(), no, n, s, t, w, id]);
    return { ok: true };
  });
}

function top_(no, id) {
  if (!(no > 0)) return { ok: false };
  const rows = tab_(TAB_SKOR, HEAD_SKOR).getDataRange().getValues().slice(1).filter(r => r[1] == no);
  rows.sort((a, b) => b[3] - a[3] || a[4] - b[4]);
  const rank = rows.findIndex(r => r[6] == id) + 1;
  return {
    ok: true, total: rows.length, rank: rank || null,
    top: rows.slice(0, 10).map(r => ({ n: String(r[2]), s: Number(r[3]), t: Number(r[4]), mine: r[6] == id }))
  };
}

function event_(d) {
  const allowed = ['start', 'finish', 'fail', 'ending', 'choice', 'daily'];
  const e = String(d.e || '');
  if (allowed.indexOf(e) < 0) return;
  tab_(TAB_EVENT, HEAD_EVENT).appendRow([
    new Date(), e, clean_(d.lv, 10), clean_(d.mode, 10),
    d.stars === undefined ? '' : Number(d.stars) || 0,
    clean_(d.x, 20), clean_(d.id, 40), clean_(d.lang, 2)
  ]);
}

function review_(d) {
  const rating = parseInt(d.r, 10), lv = clean_(d.lv, 10), id = clean_(d.id, 40), lang = clean_(d.lang, 2);
  if (!(rating >= 1 && rating <= 5) || !lv || !id) return { ok: false };
  const text = safeText_(d.x, 400);
  return withLock_(() => {
    const sh = tab_(TAB_ULASAN, HEAD_ULASAN);
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][4] == id && rows[i][1] == lv) {
        sh.getRange(i + 1, 1, 1, 7).setValues([[new Date(), lv, rating, text, id, lang, '']]);
        return { ok: true, updated: true };
      }
    }
    sh.appendRow([new Date(), lv, rating, text, id, lang, '']);
    return { ok: true };
  });
}

function cfg_() {
  const p = PropertiesService.getScriptProperties();
  const on = p.getProperty('ANN_ON') === '1';
  return { ok: true, ann: on ? (p.getProperty('ANN') || '') : '', annId: on ? (p.getProperty('ANN_ID') || '') : '' };
}

/* ---------- Admin ---------- */
function admin_(d) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!key || String(d.key || '') !== key) { Utilities.sleep(1200); return { ok: false, err: 'key' }; }
  switch (d.op) {
    case 'stats': return stats_();
    case 'reviews': return reviews_();
    case 'review_set': return reviewSet_(parseInt(d.row, 10), String(d.status || ''));
    case 'scores': return scores_(parseInt(d.no, 10));
    case 'score_del': return scoreDel_(parseInt(d.row, 10), String(d.pid || ''));
    case 'ban': return ban_(String(d.name || ''), !!d.off);
    case 'ann_get': { const c = PropertiesService.getScriptProperties(); return { ok: true, ann: c.getProperty('ANN') || '', on: c.getProperty('ANN_ON') === '1', banned: banned_() }; }
    case 'ann_set': return annSet_(String(d.ann || ''), !!d.on);
    case 'players': return players_();
    case 'player_pin': return playerPin_(parseInt(d.row, 10), String(d.u || ''));
    case 'player_del': return playerDel_(parseInt(d.row, 10), String(d.u || ''));
  }
  return { ok: false, err: 'op' };
}

function stats_() {
  const tz = Session.getScriptTimeZone();
  const day = dt => Utilities.formatDate(new Date(dt), tz, 'yyyy-MM-dd');
  const ev = tab_(TAB_EVENT, HEAD_EVENT).getDataRange().getValues().slice(1);
  const players = {}, langs = {}, modes = { story: 0, detective: 0 };
  const start = {}, finish = {}, fails = {}, stars = {};
  LEVELS.forEach(l => { start[l] = {}; finish[l] = {}; fails[l] = 0; stars[l] = [0, 0, 0, 0]; });
  const endings = { home: 0, two: 0, late: 0 }, choice = { leak: 0, quiet: 0 };
  const days = {};
  const today = day(new Date());
  for (let i = 13; i >= 0; i--) days[day(Date.now() - i * 864e5)] = { p: {}, ev: 0 };
  let daily = 0;
  ev.forEach(r => {
    const [t, e, lv, mode, st, x, id, lang] = r;
    players[id] = 1; if (lang) langs[id] = lang;
    const dk = day(t); if (days[dk]) { days[dk].p[id] = 1; days[dk].ev++; }
    if (e === 'start' && start[lv]) { start[lv][id] = 1; if (modes[mode] !== undefined) modes[mode]++; }
    if (e === 'finish' && finish[lv]) { finish[lv][id] = 1; const s = Number(st) || 0; if (s >= 1 && s <= 3) stars[lv][s]++; }
    if (e === 'fail' && fails[lv] !== undefined) { fails[lv]++; if (x === 'late') endings.late++; }
    if (e === 'ending' && endings[x] !== undefined) endings[x]++;
    if (e === 'choice' && choice[x] !== undefined) choice[x]++;
    if (e === 'daily') daily++;
  });
  const langCount = { id: 0, en: 0 };
  Object.keys(langs).forEach(k => { if (langCount[langs[k]] !== undefined) langCount[langs[k]]++; });
  const funnel = LEVELS.map(l => ({
    lv: l, start: Object.keys(start[l]).length, finish: Object.keys(finish[l]).length, fail: fails[l], stars: stars[l].slice(1)
  }));
  const sk = tab_(TAB_SKOR, HEAD_SKOR).getDataRange().getValues().slice(1);
  const cases = {};
  sk.forEach(r => {
    const no = r[1]; cases[no] = cases[no] || { no: no, n: 0, s: 0, t: 0, three: 0, last: r[0] };
    const c = cases[no]; c.n++; c.s += Number(r[3]); c.t += Number(r[4]); if (Number(r[3]) === 3) c.three++;
  });
  const caseList = Object.keys(cases).map(k => {
    const c = cases[k];
    return { no: c.no, n: c.n, avgS: +(c.s / c.n).toFixed(2), avgT: Math.round(c.t / c.n), pct3: Math.round(c.three / c.n * 100) };
  }).sort((a, b) => b.no - a.no);
  const rv = tab_(TAB_ULASAN, HEAD_ULASAN).getDataRange().getValues().slice(1);
  const dist = [0, 0, 0, 0, 0]; let sum = 0, fresh = 0;
  rv.forEach(r => { const g = Number(r[2]); if (g >= 1 && g <= 5) { dist[g - 1]++; sum += g; } if (!r[6]) fresh++; });
  return {
    ok: true, tz: tz, today: today,
    players: Object.keys(players).length,
    todayPlayers: days[today] ? Object.keys(days[today].p).length : 0,
    events: ev.length, daily: daily, modes: modes, lang: langCount,
    funnel: funnel, endings: endings, choice: choice,
    days: Object.keys(days).map(k => ({ d: k, p: Object.keys(days[k].p).length, ev: days[k].ev })),
    cases: caseList,
    accounts: Math.max(0, tab_(TAB_PEMAIN, HEAD_PEMAIN).getLastRow() - 1),
    reviews: { n: rv.length, avg: rv.length ? +(sum / rv.length).toFixed(2) : 0, dist: dist, fresh: fresh }
  };
}

function reviews_() {
  const rows = tab_(TAB_ULASAN, HEAD_ULASAN).getDataRange().getValues();
  const out = [];
  for (let i = rows.length - 1; i >= 1 && out.length < 300; i--) {
    const r = rows[i];
    out.push({ row: i + 1, t: new Date(r[0]).getTime(), lv: r[1], r: Number(r[2]), x: String(r[3]).replace(/^'/, ''), lang: r[5], st: r[6] });
  }
  return { ok: true, list: out };
}

function reviewSet_(row, status) {
  if (['', 'dibaca', 'disorot', 'disembunyikan'].indexOf(status) < 0) return { ok: false };
  const sh = tab_(TAB_ULASAN, HEAD_ULASAN);
  if (!(row >= 2 && row <= sh.getLastRow())) return { ok: false };
  sh.getRange(row, 7).setValue(status);
  return { ok: true };
}

function scores_(no) {
  const rows = tab_(TAB_SKOR, HEAD_SKOR).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (no > 0 && r[1] != no) continue;
    out.push({ row: i + 1, t: new Date(r[0]).getTime(), no: r[1], n: r[2], s: Number(r[3]), sec: Number(r[4]), w: Number(r[5]), pid: String(r[6]) });
  }
  out.sort((a, b) => b.s - a.s || a.sec - b.sec);
  return { ok: true, list: out.slice(0, 300) };
}

function scoreDel_(row, pid) {
  return withLock_(() => {
    const sh = tab_(TAB_SKOR, HEAD_SKOR);
    if (!(row >= 2 && row <= sh.getLastRow())) return { ok: false };
    if (String(sh.getRange(row, 7).getValue()) !== pid) return { ok: false, err: 'moved' };
    sh.deleteRow(row);
    return { ok: true };
  });
}

function banned_() {
  const v = PropertiesService.getScriptProperties().getProperty('BANNED');
  try { return v ? JSON.parse(v) : []; } catch (e) { return []; }
}

function ban_(name, off) {
  const n = clean_(name, 12).toLowerCase();
  if (!n) return { ok: false };
  let list = banned_().filter(x => x !== n);
  if (!off) list.push(n);
  PropertiesService.getScriptProperties().setProperty('BANNED', JSON.stringify(list.slice(-200)));
  return { ok: true, banned: list };
}

function annSet_(text, on) {
  const p = PropertiesService.getScriptProperties();
  const t = String(text).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, 200);
  if (t !== (p.getProperty('ANN') || '')) p.setProperty('ANN_ID', String(Date.now()));
  p.setProperty('ANN', t);
  p.setProperty('ANN_ON', on && t ? '1' : '0');
  return { ok: true };
}

/* ---------- Akun pemain ---------- */
function uname_(v) { const u = String(v || '').toLowerCase().trim(); return /^[a-z0-9_.]{3,15}$/.test(u) ? u : ''; }
function pinOk_(p) { return /^\d{6}$/.test(p) && !/^(\d)\1{5}$/.test(p) && ['123456', '654321', '012345', '123123'].indexOf(p) < 0; }
function hash_(salt, pin) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin)); }
function newTok_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function findUser_(u) {
  if (!u) return null;
  const rows = tab_(TAB_PEMAIN, HEAD_PEMAIN).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][P.U]) === u) return { row: i + 1, r: rows[i] };
  return null;
}
function toks_(r) { try { return JSON.parse(r[P.TOK] || '[]'); } catch (e) { return []; } }
function tokenOk_(row, tok) {
  if (!tok) return false;
  const r = tab_(TAB_PEMAIN, HEAD_PEMAIN).getRange(row, 1, 1, HEAD_PEMAIN.length).getValues()[0];
  return toks_(r).indexOf(tok) >= 0;
}
function data_(v) { const s = JSON.stringify(v || {}); return s.length <= 40000 ? s : null; }

function accNew_(d) {
  const u = uname_(d.u), pin = String(d.pin || '');
  if (!u) return { ok: false, err: 'uname' };
  if (!pinOk_(pin)) return { ok: false, err: 'weak' };
  if (banned_().indexOf(u) >= 0) return { ok: false, err: 'name' };
  const data = data_(d.data); if (data === null) return { ok: false, err: 'big' };
  return withLock_(() => {
    if (findUser_(u)) return { ok: false, err: 'taken' };
    const salt = Utilities.getUuid(), tok = newTok_(), now = new Date();
    tab_(TAB_PEMAIN, HEAD_PEMAIN).appendRow([u, salt, hash_(salt, pin), JSON.stringify([tok]), data, now, now, 0, '']);
    return { ok: true, u: u, tok: tok };
  });
}

function accLogin_(d) {
  const u = uname_(d.u), pin = String(d.pin || '');
  return withLock_(() => {
    const f = findUser_(u);
    if (!f) { Utilities.sleep(600); return { ok: false, err: 'cred' }; }
    const sh = tab_(TAB_PEMAIN, HEAD_PEMAIN), r = f.r, now = Date.now();
    const lockUntil = r[P.LOCK] ? new Date(r[P.LOCK]).getTime() : 0;
    if (lockUntil > now) return { ok: false, err: 'locked', wait: Math.ceil((lockUntil - now) / 60000) };
    if (hash_(r[P.SALT], pin) !== r[P.HASH]) {
      const fails = (Number(r[P.FAIL]) || 0) + 1;
      if (fails >= 5) { sh.getRange(f.row, P.FAIL + 1, 1, 2).setValues([[0, new Date(now + 15 * 60000)]]); return { ok: false, err: 'locked', wait: 15 }; }
      sh.getRange(f.row, P.FAIL + 1).setValue(fails);
      Utilities.sleep(600);
      return { ok: false, err: 'cred', left: 5 - fails };
    }
    const tok = newTok_(), list = toks_(r).concat([tok]).slice(-5);
    sh.getRange(f.row, P.TOK + 1).setValue(JSON.stringify(list));
    sh.getRange(f.row, P.SEEN + 1, 1, 3).setValues([[new Date(), 0, '']]);
    let data = {}; try { data = JSON.parse(r[P.DATA] || '{}'); } catch (e) {}
    return { ok: true, u: u, tok: tok, data: data };
  });
}

function accSave_(d) {
  const u = uname_(d.u), f = findUser_(u);
  if (!f || toks_(f.r).indexOf(String(d.tok || '')) < 0) return { ok: false, err: 'tok' };
  const data = data_(d.data); if (data === null) return { ok: false, err: 'big' };
  tab_(TAB_PEMAIN, HEAD_PEMAIN).getRange(f.row, P.DATA + 1, 1, 3).setValues([[data, f.r[P.MADE], new Date()]]);
  return { ok: true };
}

function accLoad_(d) {
  const u = uname_(d.u), f = findUser_(u);
  if (!f || toks_(f.r).indexOf(String(d.tok || '')) < 0) return { ok: false, err: 'tok' };
  let data = {}; try { data = JSON.parse(f.r[P.DATA] || '{}'); } catch (e) {}
  return { ok: true, data: data };
}

function players_() {
  const rows = tab_(TAB_PEMAIN, HEAD_PEMAIN).getDataRange().getValues();
  const out = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i]; let data = {}; try { data = JSON.parse(r[P.DATA] || '{}'); } catch (e) {}
    const best = data.best || {};
    out.push({
      row: i + 1, u: String(r[P.U]), made: new Date(r[P.MADE]).getTime(), seen: new Date(r[P.SEEN]).getTime(),
      done: LEVELS.filter(l => best[l]).length, stars: LEVELS.reduce((a, l) => a + (best[l] ? best[l].stars : 0), 0),
      streak: data.daily ? data.daily.streak || 0 : 0, endings: Object.keys(data.endings || {}),
      locked: r[P.LOCK] && new Date(r[P.LOCK]).getTime() > Date.now()
    });
  }
  return { ok: true, list: out };
}

function playerRow_(row, u) {
  const sh = tab_(TAB_PEMAIN, HEAD_PEMAIN);
  if (!(row >= 2 && row <= sh.getLastRow())) return null;
  return String(sh.getRange(row, 1).getValue()) === u ? sh : null;
}

function playerPin_(row, u) {
  return withLock_(() => {
    const sh = playerRow_(row, u); if (!sh) return { ok: false, err: 'moved' };
    let pin; do { pin = String(Math.floor(Math.random() * 1e6)).padStart(6, '0'); } while (!pinOk_(pin));
    const salt = Utilities.getUuid();
    sh.getRange(row, P.SALT + 1, 1, 3).setValues([[salt, hash_(salt, pin), '[]']]);
    sh.getRange(row, P.FAIL + 1, 1, 2).setValues([[0, '']]);
    return { ok: true, pin: pin };
  });
}

function playerDel_(row, u) {
  return withLock_(() => {
    const sh = playerRow_(row, u); if (!sh) return { ok: false, err: 'moved' };
    sh.deleteRow(row);
    return { ok: true };
  });
}

/* ---------- Bantuan ---------- */
// Nama dan label: hanya huruf, angka, spasi, titik, garis bawah.
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[^\p{L}\p{N} _.]/gu, '').trim().slice(0, max);
}
// Teks bebas (ulasan): tanda baca boleh, tapi tidak boleh terbaca sebagai rumus Sheet.
function safeText_(v, max) {
  let t = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function tab_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
