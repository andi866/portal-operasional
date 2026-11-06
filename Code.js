/**
 * PORTAL OPERASIONAL · ELS.ID Purwokerto (Terima Barang + Stok Harian)
 * Backend Google Apps Script (Web App) · v2
 *
 * Level akun: SUPERADMIN (semua + kelola akun), ADMIN (semua kecuali kelola akun), GUDANG (cek fisik, tanpa IO, rekap, indent).
 *
 * Sheet (dibuat otomatis oleh setupAwal):
 *   USERS · MASTER_BARANG (IMPORTRANGE) · PENERIMAAN · KIRIMAN · BARKUR · INPUT · TANPA_IO · SETTING
 */

const APP = {
  NAMA: 'Portal Operasional',
  CABANG: 'ELS.ID Purwokerto',
  TZ: 'Asia/Jakarta',
  FOLDER_NAMA: 'PORTAL OPERASIONAL',
  SESI_DETIK: 6 * 60 * 60,
  TOKEN_JAM: 20,
  CACHE_DETIK: 6 * 60 * 60,
  MAX_GAGAL: 5
};

const HEAD = {
  USERS: ['USERNAME', 'NAMA', 'ROLE', 'PIN', 'AKTIF', 'DIBUAT', 'AKSES_KOTA', 'MODUL', 'MENU'],
  MASTER_BARANG: ['KODE BARANG', 'NAMA BARANG'],
  PENERIMAAN: ['ID', 'DIBUAT', 'DIBUAT_OLEH', 'TGL_KIRIMAN', 'GUDANG_ASAL', 'FILE_KIRIMAN', 'FILE_BARKUR', 'FOLDER_URL', 'STATUS', 'SELESAI', 'SELESAI_OLEH', 'BARIS', 'QTY_DATA', 'QTY_FISIK', 'KODE_DICEK', 'JML_KODE', 'KODE_KURANG', 'KODE_TANPA_IO', 'ARSIP_ID', 'GUDANG_SELESAI', 'RINGKASAN_KOTA'],
  KIRIMAN: ['ID_TRM', 'LINE', 'NO_IO', 'TANGGAL', 'GUDANG', 'TUJUAN', 'KODE', 'NAMA', 'QTY'],
  BARKUR: ['ID_TRM', 'ID_PERMINTAAN', 'TGL_REQUEST', 'TGL_REQ_MS', 'GUDANG_TUJUAN', 'SALES', 'NO_SO', 'KODE', 'NAMA', 'QTY_DIMINTA', 'QTY_TERIMA_AWAL', 'MANUAL', 'OLEH', 'RAW'],
  INPUT: ['ID', 'ID_TRM', 'WAKTU', 'USERNAME', 'PETUGAS', 'KODE', 'NAMA_BARANG', 'QTY', 'KETERANGAN', 'ASAL_GUDANG', 'ADA_DI_IO', 'BATAL', 'BATAL_OLEH', 'BATAL_WAKTU'],
  TINDAK_LANJUT: ['ID', 'ID_TRM', 'KOTA', 'NO_IO', 'KODE', 'NAMA', 'JENIS', 'QTY', 'AKSI', 'STATUS', 'KETERANGAN', 'DIBUAT', 'OLEH', 'SELESAI', 'SELESAI_OLEH', 'CATATAN'],
  STOK_SESI: ['ID', 'TANGGAL', 'DIBUAT', 'OLEH', 'STATUS', 'UPLOAD_KE', 'UPLOAD_TERAKHIR', 'SELESAI', 'SELESAI_OLEH', 'JML_ITEM', 'QTY_SISTEM', 'QTY_FISIK', 'OK', 'CEK', 'BELUM', 'RIWAYAT_UPLOAD', 'FOLDER_URL', 'ARSIP_ID', 'CATATAN'],
  STOK_ITEM: ['ID_STK', 'KODE', 'NAMA', 'QTY_SISTEM', 'QTY_AWAL', 'KONSINYASI', 'KUNCI', 'KUNCI_SISTEM', 'KUNCI_FISIK', 'KUNCI_WAKTU', 'KUNCI_UPLOAD', 'DI_UPLOAD_TERAKHIR'],
  STOK_INPUT: ['ID', 'ID_STK', 'WAKTU', 'USERNAME', 'PETUGAS', 'KODE', 'NAMA', 'QTY', 'KETERANGAN', 'JENIS', 'BATAL', 'BATAL_OLEH', 'BATAL_WAKTU'],
  SETTING: ['KUNCI', 'NILAI']
};
const P = {};
Object.keys(HEAD).forEach(s => { P[s] = {}; HEAD[s].forEach((h, i) => P[s][h] = i); });
const ROLES = ['SUPERADMIN', 'ADMIN', 'GUDANG', 'SERVICE', 'SALES'];

/* ================= WEB APP ================= */

function doGet() {
  let out = null;
  const fid = getSetting_('INDEX_FILE_ID');
  if (fid) {
    let html = cacheGet_('index_html');
    if (!html) {
      try { html = DriveApp.getFileById(fid).getBlob().getDataAsString('UTF-8'); cachePut_('index_html', html, 600); } catch (e) { html = null; }
    }
    if (html) out = HtmlService.createHtmlOutput(html);
  }
  if (!out) out = HtmlService.createHtmlOutputFromFile('Index');
  return out
    .setTitle(APP.NAMA + ' · ' + APP.CABANG)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Portal Operasional')
    .addItem('Setup awal / perbaiki sheet', 'setupAwal')
    .addItem('Aktifkan arsip otomatis bulanan', 'pasangArsipOtomatis')
    .addItem('Arsipkan sekarang (data lama)', 'menuArsipSekarang')
    .addItem('Arsipkan satu tahun…', 'menuArsip')
    .addItem('Pakai file tampilan dari Drive…', 'menuIndexDrive')
    .addItem('Bersihkan cache', 'bersihkanCache')
    .addToUi();
}

/* tampilan (Index.html) dibaca dari file di Google Drive, supaya tidak perlu salin-tempel file besar */
function menuIndexDrive() {
  const ui = SpreadsheetApp.getUi();
  const cur = getSetting_('INDEX_FILE_ID');
  const r = ui.prompt('File tampilan dari Drive', 'Tempel link file Index.html di Google Drive.' + (cur ? '\nSaat ini: ' + cur + '\nKosongkan lalu OK untuk kembali memakai file Index di Apps Script.' : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const txt = str_(r.getResponseText());
  if (!txt) { setSetting_('INDEX_FILE_ID', ''); cacheDel_('index_html'); ui.alert('Kembali memakai file Index di Apps Script.'); return; }
  const m = txt.match(/[-\w]{25,}/);
  if (!m) { ui.alert('Link tidak dikenali. Salin link dari Drive: klik kanan file > Bagikan > Salin link.'); return; }
  try {
    const f = DriveApp.getFileById(m[0]);
    const html = f.getBlob().getDataAsString('UTF-8');
    const ok = html.indexOf('AKHIR FILE INDEX') >= 0;
    setSetting_('INDEX_FILE_ID', m[0]); cacheDel_('index_html');
    ui.alert('Tersambung ke file: ' + f.getName() + ' (' + Math.round(html.length / 1024) + ' KB)' + (ok ? '. Penanda akhir file ditemukan, file lengkap.' : '. PERINGATAN: penanda akhir file tidak ditemukan, file mungkin tidak lengkap.') + '\nBuka ulang link web app. Tidak perlu deploy ulang untuk update tampilan berikutnya.');
  } catch (e) { ui.alert('File tidak bisa dibuka: ' + e.message); }
}

/* ================= SETUP ================= */

function setupAwal() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]);
      if (sh.setFrozenRows) sh.setFrozenRows(1);
    } else if (sh.getLastColumn() < HEAD[name].length) {
      sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]);
    }
  });
  let tanpa = ss.getSheetByName('TANPA_IO');
  if (!tanpa) tanpa = ss.insertSheet('TANPA_IO');
  tanpa.getRange(1, 1).setFormula(
    "=QUERY(INPUT!A:N;\"select C, B, E, F, G, H, J, I, L, M where K = 'TIDAK' label C 'WAKTU', B 'NO PENERIMAAN', E 'PETUGAS', F 'KODE', G 'NAMA BARANG', H 'QTY', J 'DATANG DARI', I 'KETERANGAN', L 'BATAL', M 'BATAL OLEH'\";1)"
  );
  const users = sheet_('USERS');
  const all = rows_('USERS');
  if (!all.some(r => str_(r[0]))) {
    users.appendRow(['SUPERADMIN', 'SUPER ADMIN', 'SUPERADMIN', "'1234", 'YA', new Date()]);
  } else if (!all.some(r => up_(r[2]) === 'SUPERADMIN')) {
    const i = all.findIndex(r => up_(r[2]) === 'ADMIN');
    if (i >= 0) users.getRange(i + 2, 3).setValue('SUPERADMIN');
  }
  if (!getSetting_('FOLDER_ID')) {
    try { setSetting_('FOLDER_ID', DriveApp.createFolder(APP.FOLDER_NAMA).getId()); } catch (e) {}
  }
  if (!getSetting_('NAMA_CABANG')) setSetting_('NAMA_CABANG', APP.CABANG);
  const mb = sheet_('MASTER_BARANG');
  if (mb.getLastRow() < 2 && mb.getRange(1, 3).getValue() === '') {
    mb.getRange(1, 3).setValue('Ganti A1 dengan rumus IMPORTRANGE, contoh: =IMPORTRANGE("link-sheet-pusat";"MASTER!A1:B")');
  }
  bersihkanCache();
  return 'Setup selesai. Kalau ini instalasi baru, login: SUPERADMIN / PIN 1234, lalu segera ganti PIN di menu Akun.';
}
function bersihkanCache() {
  const c = CacheService.getScriptCache();
  c.remove('master_n'); c.remove('master_v'); c.remove('index_html_n');
  rows_('PENERIMAAN').map(r => str_(r[0])).filter(Boolean).slice(-300).forEach(id => c.remove('sess_' + id + '_n'));
}

/* ================= UTIL ================= */

function sheet_(name, ss) {
  const sh = (ss || SpreadsheetApp.getActive()).getSheetByName(name);
  if (!sh) throw new Error('Sheet ' + name + ' belum ada. Jalankan menu Terima Barang > Setup awal.');
  return sh;
}
function rows_(name, ss) {
  const sh = sheet_(name, ss), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEAD[name].length).getValues();
}
function txt_(v) { v = v == null ? '' : String(v); return v === '' ? '' : "'" + v; }
function ms_(v) { return v instanceof Date ? v.getTime() : (v === '' || v == null ? null : Number(v)); }
function str_(v) { return v == null ? '' : String(v).trim(); }
function up_(v) { return str_(v).toUpperCase(); }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Server sedang sibuk, coba lagi sebentar.');
  try { return fn(); } finally { lock.releaseLock(); }
}
function getSetting_(k) { const r = rows_('SETTING').find(x => x[0] === k); return r ? str_(r[1]) : ''; }
function setSetting_(k, v) {
  const sh = sheet_('SETTING'), all = rows_('SETTING');
  const i = all.findIndex(x => x[0] === k);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(v); else sh.appendRow([k, v]);
}
function hash_(pin, salt) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return 'h$' + salt + '$' + b.map(x => ((x + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function checkPin_(stored, pin) {
  stored = str_(stored);
  if (stored.indexOf('h$') === 0) return hash_(pin, stored.split('$')[1]) === stored;
  return stored === String(pin);
}
function newHash_(pin) { return hash_(String(pin), Utilities.getUuid().slice(0, 8)); }

/* cache besar dipotong-potong (batas CacheService 100 KB per kunci) */
function cachePut_(key, obj, ttl) {
  try {
    const s = JSON.stringify(obj), size = 90000, parts = {};
    const n = Math.ceil(s.length / size) || 1;
    if (n > 40) return;
    for (let i = 0; i < n; i++) parts[key + '_' + i] = s.slice(i * size, (i + 1) * size);
    parts[key + '_n'] = String(n);
    CacheService.getScriptCache().putAll(parts, ttl || APP.CACHE_DETIK);
  } catch (e) {}
}
function cacheGet_(key) {
  try {
    const c = CacheService.getScriptCache(), n = Number(c.get(key + '_n') || 0);
    if (!n) return null;
    const keys = []; for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const got = c.getAll(keys);
    let s = ''; for (let i = 0; i < n; i++) { if (got[keys[i]] == null) return null; s += got[keys[i]]; }
    return JSON.parse(s);
  } catch (e) { return null; }
}
function cacheDel_(key) { CacheService.getScriptCache().remove(key + '_n'); }

/* baca semua baris milik satu ID (bisa terpencar), dibaca per blok berurutan */
function rowsOf_(name, col, id, ss) {
  const sh = sheet_(name, ss), n = sh.getLastRow();
  if (n < 2) return [];
  const hits = sh.getRange(2, col + 1, n - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findAll();
  if (!hits.length) return [];
  const nums = hits.map(h => h.getRow()).sort((a, b) => a - b);
  const out = [], w = HEAD[name].length;
  const flush = (a, b) => sh.getRange(a, 1, b - a + 1, w).getValues().forEach((r, i) => { if (str_(r[col]) === id) out.push({ r: r, row: a + i }); });
  let start = nums[0], prev = nums[0];
  for (let i = 1; i < nums.length; i++) {
    if (nums[i] - prev > 30) { flush(start, prev); start = nums[i]; }
    prev = nums[i];
  }
  flush(start, prev);
  return out;
}
function deleteRowsOf_(name, rows) {
  const sh = sheet_(name);
  if (!rows.length) return;
  /* Google Sheets menolak kalau semua baris di bawah judul terhapus; sisakan satu baris kosong */
  try { if (sh.getMaxRows && rows.length >= sh.getMaxRows() - 1) sh.insertRowAfter(sh.getMaxRows()); } catch (e) {}
  const nums = rows.slice().sort((a, b) => b - a);
  let i = 0;
  while (i < nums.length) {
    const end = nums[i]; let start = end;
    while (i + 1 < nums.length && nums[i + 1] === start - 1) start = nums[++i];
    sh.deleteRows(start, end - start + 1); i++;
  }
}
function findRow_(name, col, value, ss) {
  const sh = sheet_(name, ss), n = sh.getLastRow();
  if (n < 2) return -1;
  const hit = sh.getRange(2, col + 1, n - 1, 1).createTextFinder(String(value)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : -1;
}
function appendRows_(name, rows) {
  if (!rows.length) return;
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/* ================= AUTH ================= */

function roleOf_(v) { v = up_(v); return ROLES.indexOf(v) >= 0 ? v : 'GUDANG'; }
function userPublic_(r) {
  const role = roleOf_(r[2]);
  let akses = role === 'SERVICE' ? aksesOf_(r[6]) : [];
  if (role === 'SERVICE' && !akses.length) akses = ['SERVICE'];
  const menu = menuOf_(role, r[7], r[8]);
  const modul = role === 'SUPERADMIN' ? ['TERIMA', 'STOK'] : ['TERIMA', 'STOK'].filter(m => menu.some(v => MENU_MOD[v] === m));
  return { username: up_(r[0]), nama: up_(r[1]) || up_(r[0]), role: role, aktif: up_(r[4]) !== 'TIDAK', akses: akses, modul: modul, menu: menu, menuCustom: !!str_(r[8]) };
}
/* ===== akses menu per akun ===== */
const MENU_MOD = { beranda: 'TERIMA', cek: 'TERIMA', tanpa: 'TERIMA', rekap: 'TERIMA', indent: 'TERIMA', tl: 'TERIMA', export: 'TERIMA', sberanda: 'STOK', shitung: 'STOK', sulang: 'STOK', srekap: 'STOK', slapor: 'STOK', sberita: 'STOK' };
const MENU_ROLE = {
  SUPERADMIN: Object.keys(MENU_MOD), ADMIN: Object.keys(MENU_MOD),
  GUDANG: ['beranda', 'cek', 'tanpa', 'rekap', 'indent', 'sberanda', 'shitung', 'sulang', 'srekap'],
  SERVICE: ['beranda', 'cek', 'tanpa', 'rekap', 'sberanda', 'shitung', 'sulang', 'srekap'],
  SALES: ['sberanda', 'shitung', 'sulang', 'srekap']
};
function menuOf_(role, modRaw, menuRaw) {
  const allowed = MENU_ROLE[role] || [];
  if (role === 'SUPERADMIN') return allowed.slice();
  const saved = str_(menuRaw).split(/[,;\s]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
  if (saved.length) { if (saved.indexOf('srekap') >= 0 && saved.indexOf('slapor') < 0) saved.push('slapor'); if (saved.indexOf('slapor') >= 0 && saved.indexOf('sberita') < 0) saved.push('sberita'); return allowed.filter(v => saved.indexOf(v) >= 0); }
  const mods = modulOf_(role, modRaw);
  return allowed.filter(v => mods.indexOf(MENU_MOD[v]) >= 0);
}
function apiMenuInfo(token) { auth_(token, 'SUPERADMIN'); return { role: MENU_ROLE, mod: MENU_MOD }; }
/* modul aplikasi per akun: TERIMA (terima barang) dan/atau STOK (stok harian) */
function modulOf_(role, v) {
  if (role === 'SUPERADMIN') return ['TERIMA', 'STOK'];
  const m = str_(v).split(/[,;\s]+/).map(up_).filter(x => x === 'TERIMA' || x === 'STOK');
  if (m.length) return [...new Set(m)];
  if (role === 'SALES') return ['STOK'];
  if (role === 'SERVICE') return ['TERIMA'];
  return ['TERIMA', 'STOK'];
}
function hasMod_(u, m) { return !!(u && u.modul && u.modul.indexOf(m) >= 0); }
function needMod_(u, m) { if (!hasMod_(u, m)) throw new Error('Akun ini tidak punya akses ke menu ' + (m === 'STOK' ? 'Stok Harian' : 'Terima Barang') + '.'); return u; }

/* ===== akses kota per akun (kosong = semua kota) ===== */
function aksesOf_(v) { return str_(v).split(/[,;]+/).map(s => up_(s)).filter(Boolean); }
function scoped_(u) { return !!(u && u.role === 'SERVICE' && u.akses && u.akses.length); }
function allowed_(u, g) { return !scoped_(u) || u.akses.some(k => up_(g).indexOf(k) >= 0); }
function metaFor_(u, m, nLines) {
  if (!scoped_(u)) return m;
  const x = JSON.parse(JSON.stringify(m));
  x.gudang = x.gudang.filter(g => allowed_(u, g));
  const d = {}; Object.keys(x.gudangDone || {}).forEach(g => { if (allowed_(u, g)) d[g] = x.gudangDone[g]; });
  x.gudangDone = d; x.folderUrl = '';
  const s = { baris: 0, data: 0, fisik: 0, dicek: 0, kode: 0, kurang: 0, tak: 0 };
  Object.keys(x.kota || {}).forEach(g => { if (!allowed_(u, g)) return; const k = x.kota[g]; Object.keys(s).forEach(f => s[f] += Number(k[f]) || 0); });
  if (nLines != null) s.baris = nLines;
  x.sum = s; delete x.kota;
  x.scoped = true;
  return x;
}
function indexFor_(u, all) {
  if (!scoped_(u)) return all;
  const vis = all.filter(m => m.gudang.some(g => allowed_(u, g)));
  fillKota_(pickPeriod_(vis, ''), 25); fillKota_(vis, 10);
  return vis.map(m => metaFor_(u, m));
}
/* ringkasan per kota untuk penerimaan lama yang belum punya (dihitung sekali lalu disimpan) */
function fillKota_(metas, limit) {
  let n = 0;
  metas.forEach(m => { if (m.kota || n >= limit) return; try { m.kota = computeKota_(m); writeKota_(m.id, m.kota); n++; } catch (e) {} });
}
function computeKota_(m) {
  const ss = srcSS_(m), lines = staticData_(m.id, ss).lines, scans = rowsOf_('INPUT', 1, m.id, ss).map(x => scanFromRow_(x.r)).filter(s => !s.undone);
  const out = {}, need = {}, keyG = {}, fis = {}, tak = {};
  const O = g => out[g] || (out[g] = { baris: 0, data: 0, fisik: 0, dicek: 0, kode: 0, kurang: 0, tak: 0 });
  lines.forEach(l => { const o = O(l.gudang); o.baris++; o.data += l.qty; const k = l.gudang + '|' + l.key; need[k] = (need[k] || 0) + l.qty; (keyG[l.key] = keyG[l.key] || []).indexOf(l.gudang) < 0 && keyG[l.key].push(l.gudang); });
  scans.forEach(s => {
    const gs = keyG[s.key];
    if (!gs) { if (s.asal) { tak[s.asal + '|' + s.key] = 1; O(s.asal).fisik += s.qty; } return; }
    const g = s.asal && gs.indexOf(s.asal) >= 0 ? s.asal : gs[0];
    fis[g + '|' + s.key] = (fis[g + '|' + s.key] || 0) + s.qty;
  });
  Object.keys(need).forEach(k => { const g = k.split('|')[0], o = O(g), f = fis[k] || 0; o.kode++; o.fisik += f; if (f > 0) o.dicek++; if (f < need[k]) o.kurang++; });
  Object.keys(tak).forEach(k => O(k.split('|')[0]).tak++);
  return out;
}
function writeKota_(id, kota) { const row = findRow_('PENERIMAAN', 0, id); if (row > 0) sheet_('PENERIMAAN').getRange(row, P.PENERIMAAN.RINGKASAN_KOTA + 1).setValue(JSON.stringify(kota)); }
function scanScope_(u, lines, scans) {
  if (!scoped_(u)) return scans;
  const keys = {}; lines.forEach(l => { if (allowed_(u, l.gudang)) keys[l.key] = 1; });
  return scans.filter(s => s.asal ? allowed_(u, s.asal) : keys[s.key]);
}
function sessionFor_(u, out) {
  if (!scoped_(u)) return out;
  const lines = out.lines.filter(l => allowed_(u, l.gudang));
  const keys = {}; lines.forEach(l => keys[l.key] = 1);
  if (!lines.length) throw new Error('Penerimaan ini tidak berisi IO untuk kota akun kamu.');
  return { meta: metaFor_(u, out.meta, lines.length), lines: lines, barkur: out.barkur.filter(b => keys[b.key]), scans: scanScope_(u, out.lines, out.scans), info: out.info };
}
/* level: kosong = semua akun · 'ADMIN' = admin & super admin · 'SUPERADMIN' = hanya super admin */
/* token bertanda tangan: v2.<user>.<iat>.<exp>.<sig> (5 bagian) — tidak hilang walaupun cache Google dibersihkan */
function tokSecret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('TOKEN_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('TOKEN_SECRET', s); }
  return s;
}
function tokSig_(payload) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, tokSecret_())).replace(/=+$/, ''); }
function makeToken_(username) {
  const iat = Date.now(), exp = iat + APP.TOKEN_JAM * 3600e3;
  const payload = 'v2.' + Utilities.base64EncodeWebSafe(username).replace(/=+$/, '') + '.' + iat + '.' + exp;
  return payload + '.' + tokSig_(payload);
}
function userByName_(username) {
  const cache = CacheService.getScriptCache(), k = 'usr_' + username, c = cache.get(k);
  if (c) { try { return JSON.parse(c); } catch (e) {} }
  const r = rows_('USERS').find(x => up_(x[0]) === username);
  if (!r) return null;
  const u = userPublic_(r);
  cache.put(k, JSON.stringify(u), 600);
  return u;
}
function auth_(token, level) {
  if (!token) throw new Error('SESI_HABIS');
  const cache = CacheService.getScriptCache();
  let u = null;
  if (String(token).indexOf('v2.') === 0) {
    const p = String(token).split('.');
    if (p.length !== 5) throw new Error('SESI_HABIS');
    const payload = p.slice(0, 4).join('.');
    if (tokSig_(payload) !== p[4]) throw new Error('SESI_HABIS');
    const iat = Number(p[2]), exp = Number(p[3]);
    if (!(exp > Date.now())) throw new Error('SESI_HABIS');
    let name = ''; try { name = Utilities.newBlob(Utilities.base64DecodeWebSafe(p[1] + '==='.slice((p[1].length + 3) % 4))).getDataAsString(); } catch (e) { throw new Error('SESI_HABIS'); }
    const bt = Number(PropertiesService.getScriptProperties().getProperty('blokts_' + up_(name)) || 0);
    if (bt && iat < bt) throw new Error('SESI_HABIS');
    u = userByName_(up_(name));
    if (!u || !u.aktif) throw new Error('SESI_HABIS');
  } else {
    const c = cache.get('tok_' + token);
    if (!c) throw new Error('SESI_HABIS');
    u = JSON.parse(c);
    cache.put('tok_' + token, c, APP.SESI_DETIK);
  }
  if (cache.get('blok_' + u.username)) throw new Error('SESI_HABIS');
  if (level === 'SUPERADMIN' && u.role !== 'SUPERADMIN') throw new Error('Menu ini khusus super admin.');
  if (level === 'ADMIN' && (u.role === 'GUDANG' || u.role === 'SERVICE' || u.role === 'SALES')) throw new Error('Menu ini khusus admin.');
  return u;
}
/* data stok ikut dikirim saat login (kalau akun membuka Stok) supaya tidak perlu bolak-balik ke server */
function stokBoot_(u, hint) {
  if (!hasMod_(u, 'STOK')) return null;
  hint = hint || {};
  if (!(hint.mod === 'stok' || !hasMod_(u, 'TERIMA'))) return null;
  try {
    stokEnsure_();
    const index = stokIndex_('');
    const aid = index.aktif ? index.aktif.id : (index.today ? index.today.id : '');
    if (!aid) return { index: index };
    if (str_(hint.snapId) === aid) { const dl = stokDelta_(aid, Number(hint.lastRow) || 0, str_(hint.trver)); dl.id = aid; return { index: index, delta: dl }; }
    const ses = stokSession_(aid); ses.id = aid;
    return { index: index, session: ses };
  } catch (e) { return null; }
}
function boot_(u, hint) {
  const all = hasMod_(u, 'TERIMA') ? indexFor_(u, allMeta_()) : [];
  let tlCount = 0; /* dihitung di belakang layar oleh aplikasi (apiTL) supaya login cepat */
  return { user: u, cabang: getSetting_('NAMA_CABANG') || APP.CABANG, index: pickPeriod_(all, ''), periods: periods_(all), years: years_(), masterV: masterVersion_(), tlCount: tlCount, stok: stokBoot_(u, hint) };
}
function apiLogin(username, pin, hint) {
  username = up_(username); pin = str_(pin);
  if (!username || !pin) throw new Error('Isi username dan PIN.');
  const cache = CacheService.getScriptCache(), fk = 'gagal_' + username;
  const gagal = Number(cache.get(fk) || 0);
  if (gagal >= APP.MAX_GAGAL) throw new Error('Terlalu banyak percobaan. Coba lagi 5 menit lagi.');
  const all = rows_('USERS');
  const i = all.findIndex(r => up_(r[0]) === username);
  if (i < 0 || !checkPin_(all[i][3], pin)) { cache.put(fk, String(gagal + 1), 300); throw new Error('Username atau PIN salah.'); }
  const u = userPublic_(all[i]);
  if (!u.aktif) throw new Error('Akun ini sedang nonaktif. Hubungi super admin.');
  if (str_(all[i][3]).indexOf('h$') !== 0) sheet_('USERS').getRange(i + 2, 4).setValue(newHash_(pin));
  cache.remove(fk); cache.remove('blok_' + username);
  const token = makeToken_(username);
  cache.put('usr_' + username, JSON.stringify(u), 600);
  const b = boot_(u, hint); b.token = token;
  return b;
}
function apiBoot(token, hint) { return boot_(auth_(token), hint); }
function apiLogout(token) { if (token) CacheService.getScriptCache().remove('tok_' + token); return true; }

/* ganti username (super admin): ikut memperbarui catatan input di sheet utama supaya riwayat & hak batal tetap nyambung */
function apiRenameUser(token, oldU, newU) {
  const me = auth_(token, 'SUPERADMIN');
  oldU = up_(oldU); newU = up_(newU).replace(/\s+/g, '');
  if (!newU) throw new Error('Isi username baru.');
  if (!/^[A-Z0-9._-]{2,30}$/.test(newU)) throw new Error('Username hanya boleh huruf, angka, titik, minus, atau garis bawah (2–30 karakter), tanpa spasi.');
  if (oldU === newU) return rows_('USERS').filter(r => str_(r[0])).map(userPublic_);
  return withLock_(() => {
    const sh = sheet_('USERS'), all = rows_('USERS');
    const i = all.findIndex(r => up_(r[0]) === oldU);
    if (i < 0) throw new Error('Akun ' + oldU + ' tidak ditemukan.');
    if (all.some(r => up_(r[0]) === newU)) throw new Error('Username ' + newU + ' sudah dipakai akun lain.');
    sh.getRange(i + 2, 1).setValue(txt_(newU));
    let n = 0;
    ['INPUT', 'STOK_INPUT'].forEach(name => {
      const s = SpreadsheetApp.getActive().getSheetByName(name); if (!s) return;
      const last = s.getLastRow(); if (last < 2) return;
      const rg = s.getRange(2, 4, last - 1, 1), v = rg.getValues();
      let changed = false;
      v.forEach(r => { if (up_(r[0]) === oldU) { r[0] = newU; changed = true; n++; } });
      if (changed) rg.setValues(v);
    });
    const cache = CacheService.getScriptCache();
    cache.removeAll(['usr_' + oldU, 'usr_' + newU, 'blok_' + newU]);
    cache.put('blok_' + oldU, '1', 120);
    PropertiesService.getScriptProperties().setProperty('blokts_' + oldU, String(Date.now()));
    const out = rows_('USERS').filter(r => str_(r[0])).map(userPublic_);
    out.renamed = n;
    return out;
  });
}
function apiChangePin(token, oldPin, newPin) {
  const me = auth_(token);
  if (!/^\d{4,6}$/.test(String(newPin))) throw new Error('PIN baru harus 4 sampai 6 angka.');
  return withLock_(() => {
    const all = rows_('USERS');
    const i = all.findIndex(r => up_(r[0]) === me.username);
    if (i < 0) throw new Error('Akun tidak ditemukan.');
    if (!checkPin_(all[i][3], String(oldPin))) throw new Error('PIN lama salah.');
    sheet_('USERS').getRange(i + 2, 4).setValue(newHash_(newPin));
    return true;
  });
}
function apiUsers(token) {
  auth_(token, 'SUPERADMIN');
  return rows_('USERS').filter(r => str_(r[0])).map(userPublic_);
}
function apiSaveUser(token, d) {
  const me = auth_(token, 'SUPERADMIN');
  return withLock_(() => {
    const sh = sheet_('USERS'), all = rows_('USERS');
    const username = up_(d.username).replace(/\s+/g, '');
    if (!username) throw new Error('Isi username.');
    const i = all.findIndex(r => up_(r[0]) === username);
    if (d.pin && !/^\d{4,6}$/.test(String(d.pin))) throw new Error('PIN harus 4 sampai 6 angka.');
    if (i < 0) {
      if (!d.pin) throw new Error('Isi PIN untuk akun baru.');
      sh.appendRow([username, up_(d.nama) || username, roleOf_(d.role), newHash_(d.pin), 'YA', new Date(), roleOf_(d.role) === 'SERVICE' ? txt_(aksesOf_(d.akses || 'SERVICE').join(', ')) : '', txt_(modulOf_(roleOf_(d.role), d.modul).join(',')), '']);
    } else {
      const row = i + 2;
      if (username === me.username && (d.aktif === false || (d.role && roleOf_(d.role) !== 'SUPERADMIN'))) throw new Error('Tidak bisa menonaktifkan atau menurunkan level akun sendiri.');
      if (d.nama != null) sh.getRange(row, 2).setValue(up_(d.nama));
      if (d.role) sh.getRange(row, 3).setValue(roleOf_(d.role));
      if (d.pin) sh.getRange(row, 4).setValue(newHash_(d.pin));
      if (d.aktif != null) sh.getRange(row, 5).setValue(d.aktif ? 'YA' : 'TIDAK');
      if (d.akses != null) sh.getRange(row, 7).setValue(txt_(aksesOf_(d.akses).join(', ')));
      if (d.modul != null) sh.getRange(row, 8).setValue(txt_(modulOf_(d.role ? roleOf_(d.role) : roleOf_(all[i][2]), d.modul).join(',')));
      if (d.role && d.modul == null) sh.getRange(row, 8).setValue('');
      if (d.menu != null) {
        const role = d.role ? roleOf_(d.role) : roleOf_(all[i][2]), allowed = MENU_ROLE[role] || [];
        const pick = (Array.isArray(d.menu) ? d.menu : String(d.menu).split(',')).map(s => String(s).trim().toLowerCase()).filter(v => allowed.indexOf(v) >= 0);
        if (!pick.length) throw new Error('Pilih minimal satu menu.');
        sh.getRange(row, 9).setValue(pick.join(','));
      }
      if (d.role && d.menu == null) sh.getRange(row, 9).setValue('');
      const cache = CacheService.getScriptCache();
      cache.remove('usr_' + username);
      if (d.aktif === false || d.role || d.akses != null || d.modul != null || d.menu != null) { cache.put('blok_' + username, '1', 120); PropertiesService.getScriptProperties().setProperty('blokts_' + username, String(Date.now())); }
      else if (d.aktif === true) cache.remove('blok_' + username);
    }
    return apiUsers(token);
  });
}

/* ================= MASTER ================= */

function masterVersion_() { return CacheService.getScriptCache().get('master_v') || ''; }
function apiMaster(token, force) {
  auth_(token);
  if (!force) { const c = cacheGet_('master'); if (c) return c; }
  const sh = sheet_('MASTER_BARANG'), n = sh.getLastRow();
  const seen = {}, rows = [];
  if (n >= 1) sh.getRange(1, 1, n, 2).getValues().forEach(r => {
    const k = str_(r[0]), nm = str_(r[1]);
    if (!k || !nm || k.charAt(0) === '#') return;
    const lk = k.toLowerCase(); if (lk === 'kode barang' || lk === 'kode sku' || lk === 'kode') return;
    const key = k.toUpperCase(); if (seen[key]) return;
    seen[key] = 1; rows.push([k, nm]);
  });
  const out = { v: String(Date.now()), rows: rows };
  cachePut_('master', out);
  CacheService.getScriptCache().put('master_v', out.v, APP.CACHE_DETIK);
  return out;
}

/* ================= PENERIMAAN ================= */

function metaFromRow_(r) {
  return {
    id: str_(r[0]), created: ms_(r[1]), createdBy: str_(r[2]),
    tanggal: str_(r[3]) ? str_(r[3]).split(',') : [], gudang: str_(r[4]) ? str_(r[4]).split('|') : [],
    files: { kiriman: str_(r[5]), barkur: str_(r[6]) }, folderUrl: str_(r[7]),
    status: str_(r[8]).toLowerCase() || 'berjalan', finished: ms_(r[9]), finishedBy: str_(r[10]),
    sum: { baris: Number(r[11]) || 0, data: Number(r[12]) || 0, fisik: Number(r[13]) || 0, dicek: Number(r[14]) || 0, kode: Number(r[15]) || 0, kurang: Number(r[16]) || 0, tak: Number(r[17]) || 0 },
    arsip: str_(r[18]), gudangDone: parseDone_(r[19]), kota: parseKota_(r[20])
  };
}
function parseKota_(v) { try { const o = str_(v) ? JSON.parse(str_(v)) : null; return o && typeof o === 'object' ? o : null; } catch (e) { return null; } }
function parseDone_(v) { try { const o = str_(v) ? JSON.parse(str_(v)) : {}; return o && typeof o === 'object' ? o : {}; } catch (e) { return {}; } }
function writeDone_(row, map) { sheet_('PENERIMAAN').getRange(row, P.PENERIMAAN.GUDANG_SELESAI + 1).setValue(Object.keys(map).length ? JSON.stringify(map) : ''); }
function gudangOfKey_(lines, key) { return [...new Set(lines.filter(l => l.key === key).map(l => l.gudang))]; }
function allGudang_(id, lines) {
  const set = {}; lines.forEach(l => { if (l.gudang) set[l.gudang] = 1; });
  rowsOf_('INPUT', 1, id).forEach(x => { if (up_(x.r[10]) === 'TIDAK' && up_(x.r[11]) !== 'YA' && str_(x.r[9])) set[str_(x.r[9])] = 1; });
  return Object.keys(set);
}
function lockedFor_(done, lines, key, asal) {
  if (asal) return !!done[asal];
  const gs = gudangOfKey_(lines, key);
  return gs.length > 0 && gs.every(g => done[g]);
}
function allMeta_() { return rows_('PENERIMAAN').filter(r => str_(r[0]) && up_(r[8]) !== 'DIHAPUS').map(metaFromRow_).reverse(); }
function ym_(ms) { return Utilities.formatDate(new Date(ms), APP.TZ, 'yyyy-MM'); }
/* period '' = semua yang berjalan + bulan ini · 'yyyy-MM' = satu bulan · 'yyyy' = satu tahun */
function pickPeriod_(all, period) {
  period = str_(period);
  if (!period) { const cur = ym_(Date.now()); return all.filter(m => m.status === 'berjalan' || ym_(m.created) === cur); }
  if (/^\d{4}$/.test(period)) return all.filter(m => ym_(m.created).slice(0, 4) === period);
  return all.filter(m => ym_(m.created) === period);
}
function periods_(all) {
  const map = {};
  all.forEach(m => {
    const k = ym_(m.created), p = map[k] || (map[k] = { ym: k, n: 0, data: 0, fisik: 0, kurang: 0, tak: 0, berjalan: 0, arsip: 0 });
    p.n++; p.data += m.sum.data; p.fisik += m.sum.fisik; p.kurang += m.sum.kurang; p.tak += m.sum.tak;
    if (m.status === 'berjalan') p.berjalan++; if (m.arsip) p.arsip++;
  });
  return Object.keys(map).sort().reverse().map(k => map[k]);
}
function index_() { return pickPeriod_(allMeta_(), ''); }
function apiIndex(token, period) { const u = needMod_(auth_(token), 'TERIMA'); const all = indexFor_(u, allMeta_()); return { index: pickPeriod_(all, period), periods: periods_(all) }; }

/* data lengkap satu bulan untuk rekap bulanan (dibaca dari sheet utama dan/atau arsip) */
function apiMonthData(token, ym) {
  auth_(token, 'ADMIN');
  const sess = pickPeriod_(allMeta_(), ym);
  const bySrc = {};
  sess.forEach(m => { const k = m.arsip || ''; (bySrc[k] = bySrc[k] || {})[m.id] = 1; });
  const lines = {}, scans = {};
  Object.keys(bySrc).forEach(src => {
    const ss = src ? SpreadsheetApp.openById(src) : null, ids = bySrc[src];
    rows_('KIRIMAN', ss).forEach(r => { const id = str_(r[0]); if (ids[id]) (lines[id] = lines[id] || []).push(lineFromRow_(r)); });
    rows_('INPUT', ss).forEach(r => { const id = str_(r[1]); if (ids[id]) (scans[id] = scans[id] || []).push(scanFromRow_(r)); });
  });
  return { sessions: sess, lines: lines, scans: scans };
}
function writeSummary_(row, s, u) {
  if (!s) return;
  const sh = sheet_('PENERIMAAN');
  if (u && scoped_(u)) {
    if (!s.kota) return;
    const cur = parseKota_(sh.getRange(row, P.PENERIMAAN.RINGKASAN_KOTA + 1).getValue()) || {};
    Object.keys(s.kota).forEach(g => { if (allowed_(u, g)) cur[g] = s.kota[g]; });
    sh.getRange(row, P.PENERIMAAN.RINGKASAN_KOTA + 1).setValue(JSON.stringify(cur));
    const t = { data: 0, fisik: 0, dicek: 0, kode: 0, kurang: 0, tak: 0 };
    Object.keys(cur).forEach(g => Object.keys(t).forEach(f => t[f] += Number(cur[g][f]) || 0));
    sh.getRange(row, P.PENERIMAAN.QTY_DATA + 1, 1, 6).setValues([[t.data, t.fisik, t.dicek, t.kode, t.kurang, t.tak]]);
    return;
  }
  sh.getRange(row, P.PENERIMAAN.QTY_DATA + 1, 1, 6).setValues([[s.data || 0, s.fisik || 0, s.dicek || 0, s.kode || 0, s.kurang || 0, s.tak || 0]]);
  if (s.kota) sh.getRange(row, P.PENERIMAAN.RINGKASAN_KOTA + 1).setValue(JSON.stringify(s.kota));
}
function sessRow_(id) {
  const row = findRow_('PENERIMAAN', 0, id);
  if (row < 0) throw new Error('Penerimaan ' + id + ' tidak ditemukan.');
  const r = sheet_('PENERIMAAN').getRange(row, 1, 1, HEAD.PENERIMAAN.length).getValues()[0];
  return { row: row, r: r, meta: metaFromRow_(r) };
}
function saveFile_(folder, f) {
  if (!f || !f.b64 || !folder) return;
  try {
    const mime = /\.xlsb$/i.test(f.name) ? 'application/vnd.ms-excel.sheet.binary.macroEnabled.12' : /\.xls$/i.test(f.name) ? 'application/vnd.ms-excel' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    folder.createFile(Utilities.newBlob(Utilities.base64Decode(f.b64), mime, f.name));
  } catch (e) {}
}
function subFolder_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function sessionFolder_(id, when) {
  const rootId = getSetting_('FOLDER_ID');
  if (!rootId) return null;
  try {
    const BLN = ['01-Januari', '02-Februari', '03-Maret', '04-April', '05-Mei', '06-Juni', '07-Juli', '08-Agustus', '09-September', '10-Oktober', '11-November', '12-Desember'];
    const root = DriveApp.getFolderById(rootId);
    return subFolder_(subFolder_(subFolder_(root, String(when.getFullYear())), BLN[when.getMonth()]), id);
  } catch (e) { return null; }
}
function kirimanRow_(id, l, line) { return [id, line, txt_(l.io), txt_(l.tanggal), txt_(l.gudang), txt_(l.tujuan), txt_(l.kode), txt_(l.nama), Number(l.qty) || 0]; }
function barkurRow_(id, b) {
  return [id, txt_(b.id), txt_(b.tglReqRaw), b.tglReq || '', txt_(b.gudangTujuan), txt_(b.sales), txt_(b.noSO), txt_(b.kode), txt_(b.nama),
    Number(b.qtyDiminta) || 0, Number(b.qtyTerimaAwal) || 0, b.manual ? 'YA' : '', txt_(b.by || ''), b.raw ? JSON.stringify(b.raw).slice(0, 45000) : ''];
}
function cleanLine_(l) { return { line: Number(l.line), io: str_(l.io), tanggal: str_(l.tanggal), gudang: str_(l.gudang), tujuan: str_(l.tujuan), kode: str_(l.kode), key: up_(l.kode), nama: str_(l.nama), qty: Number(l.qty) || 0 }; }
function cleanBarkur_(b) { return { id: str_(b.id), tglReqRaw: str_(b.tglReqRaw), tglReq: Number(b.tglReq) || null, gudangTujuan: str_(b.gudangTujuan), sales: str_(b.sales), noSO: str_(b.noSO), kode: str_(b.kode), key: up_(b.kode), nama: str_(b.nama), qtyDiminta: Number(b.qtyDiminta) || 0, qtyTerimaAwal: Number(b.qtyTerimaAwal) || 0, manual: !!b.manual, by: str_(b.by), raw: b.raw || null }; }
function lineFromRow_(r) { return cleanLine_({ line: r[1], io: r[2], tanggal: r[3], gudang: r[4], tujuan: r[5], kode: r[6], nama: r[7], qty: r[8] }); }
function refreshKirimanMeta_(row, lines, fileName, mode) {
  const sh = sheet_('PENERIMAAN');
  const tgl = [...new Set(lines.map(l => l.tanggal))].filter(Boolean).sort().join(',');
  const gud = [...new Set(lines.map(l => l.gudang))].filter(Boolean).join('|');
  sh.getRange(row, P.PENERIMAAN.TGL_KIRIMAN + 1, 1, 2).setValues([[txt_(tgl), txt_(gud)]]);
  const cur = str_(sh.getRange(row, P.PENERIMAAN.FILE_KIRIMAN + 1).getValue());
  sh.getRange(row, P.PENERIMAAN.FILE_KIRIMAN + 1).setValue(txt_(mode === 'tambah' && cur ? cur + ' + ' + fileName : fileName));
  sh.getRange(row, P.PENERIMAAN.BARIS + 1).setValue(lines.length);
}

function apiCreate(token, d) { try { return apiCreate_(token, d); } finally { trBump_(); } }
function apiCreate_(token, d) {
  const me = auth_(token, 'ADMIN');
  if (!d || !d.kiriman || !d.kiriman.lines || !d.kiriman.lines.length) throw new Error('Data kiriman kosong.');
  return withLock_(() => {
    const now = new Date();
    const ymd = Utilities.formatDate(now, APP.TZ, 'yyMMdd');
    const n = rows_('PENERIMAAN').filter(r => str_(r[0]).indexOf('TRM-' + ymd + '-') === 0).length + 1;
    const id = 'TRM-' + ymd + '-' + String(n).padStart(2, '0');
    const lines = d.kiriman.lines.map((l, i) => cleanLine_(Object.assign({}, l, { line: i + 1 })));
    appendRows_('KIRIMAN', lines.map(l => kirimanRow_(id, l, l.line)));
    const barkur = (d.barkur && d.barkur.rows ? d.barkur.rows : []).map(cleanBarkur_);
    if (barkur.length) appendRows_('BARKUR', barkur.map(b => barkurRow_(id, b)));
    const folder = sessionFolder_(id, now);
    saveFile_(folder, d.kiriman); if (d.barkur) saveFile_(folder, d.barkur);
    const tgl = [...new Set(lines.map(l => l.tanggal))].sort().join(',');
    const gud = [...new Set(lines.map(l => l.gudang))].join('|');
    const s = d.summary || {};
    sheet_('PENERIMAAN').appendRow([id, now, me.nama, txt_(tgl), txt_(gud), txt_(d.kiriman.name), txt_(d.barkur ? d.barkur.name : ''), folder ? folder.getUrl() : '', 'BERJALAN', '', '', lines.length, s.data || 0, 0, 0, s.kode || 0, s.kode || 0, 0, '', '', s.kota ? JSON.stringify(s.kota) : '']);
    cachePut_('sess_' + id, { lines: lines, barkur: barkur });
    return id;
  });
}

function scanFromRow_(r) {
  return { id: str_(r[0]), ts: ms_(r[2]), by: { id: up_(r[3]), nama: str_(r[4]) }, kode: str_(r[5]), key: up_(r[5]), nama: str_(r[6]), qty: Number(r[7]) || 0,
    ket: str_(r[8]), asal: str_(r[9]), undone: up_(r[11]) === 'YA', undoneBy: str_(r[12]) ? { nama: str_(r[12]) } : null, undoneAt: ms_(r[13]) };
}
function srcSS_(meta) { return meta.arsip ? SpreadsheetApp.openById(meta.arsip) : null; }
function staticData_(id, ss) {
  if (!ss) { const c = cacheGet_('sess_' + id); if (c) return c; }
  const lines = rowsOf_('KIRIMAN', 0, id, ss).map(x => lineFromRow_(x.r)).sort((a, b) => a.line - b.line);
  const barkur = rowsOf_('BARKUR', 0, id, ss).map(x => { const r = x.r; let raw = null; try { raw = str_(r[13]) ? JSON.parse(r[13]) : null; } catch (e) {}
    return cleanBarkur_({ id: r[1], tglReqRaw: r[2], tglReq: r[3], gudangTujuan: r[4], sales: r[5], noSO: r[6], kode: r[7], nama: r[8], qtyDiminta: r[9], qtyTerimaAwal: r[10], manual: up_(r[11]) === 'YA', by: r[12], raw: raw }); });
  const out = { lines: lines, barkur: barkur };
  if (!ss) cachePut_('sess_' + id, out);
  return out;
}
function apiSession(token, id) {
  const u = needMod_(auth_(token), 'TERIMA');
  const out = sessionFor_(u, sessionRaw_(id));
  out.tl = scoped_(u) ? [] : tlOf_(id);
  return out;
}
function sessionRaw_(id) {
  const s = sessRow_(id);
  if (s.meta.status === 'dihapus') throw new Error('Penerimaan ini sudah dihapus.');
  if (s.meta.arsip) {
    const c = cacheGet_('arsp_' + id);
    if (c) { c.meta = s.meta; return c; }
  }
  const ss = srcSS_(s.meta);
  const st = staticData_(id, ss);
  const out = { meta: s.meta, lines: st.lines, barkur: st.barkur, scans: rowsOf_('INPUT', 1, id, ss).map(x => scanFromRow_(x.r)) };
  if (s.meta.arsip) cachePut_('arsp_' + id, out);
  return out;
}
function apiScans(token, id) {
  const u = needMod_(auth_(token), 'TERIMA'); const s = sessRow_(id);
  const scans = rowsOf_('INPUT', 1, id, srcSS_(s.meta)).map(x => scanFromRow_(x.r));
  if (!scoped_(u)) return { scans: scans, meta: s.meta };
  const lines = staticData_(id, srcSS_(s.meta)).lines;
  return { scans: scanScope_(u, lines, scans), meta: metaFor_(u, s.meta, lines.filter(l => allowed_(u, l.gudang)).length) };
}

function openOnly_(s) {
  if (s.meta.arsip) throw new Error('Penerimaan ini sudah diarsipkan dan hanya bisa dilihat.');
  if (s.meta.status !== 'berjalan') throw new Error('Penerimaan ini sudah selesai. Minta admin membuka kembali.');
}
function trBump_() { try { CacheService.getScriptCache().put('trver', String(Date.now()), 21600); } catch (e) {} }
function apiSaveScan(token, id, sc, summary) { try { return apiSaveScan_(token, id, sc, summary); } finally { trBump_(); } }
function apiSaveScan_(token, id, sc, summary) {
  const me = needMod_(auth_(token), 'TERIMA');
  const q = Number(sc.qty), koreksi = !!sc.koreksi;
  if (koreksi) {
    if (!(me.role === 'ADMIN' || me.role === 'SUPERADMIN')) throw new Error('Koreksi hanya bisa dilakukan admin.');
    if (!q || Math.floor(q) !== q) throw new Error('Koreksi tidak mengubah apa pun.');
    sc.ket = 'Koreksi' + (str_(sc.ket) ? ' · ' + str_(sc.ket).replace(/^Koreksi\s*·?\s*/i, '') : '');
  } else if (!(q >= 1) || Math.floor(q) !== q) throw new Error('Qty harus angka bulat minimal 1.');
  return withLock_(() => {
    const s = sessRow_(id); openOnly_(s);
    if (scoped_(me)) {
      const own = staticData_(id).lines.filter(l => allowed_(me, l.gudang));
      if (sc.terdaftar && !own.some(l => l.key === up_(sc.kode))) throw new Error('Barang ini bukan bagian dari kota akun kamu.');
      if (sc.asal && !allowed_(me, sc.asal)) throw new Error('Kota asal ini di luar akses akun kamu.');
      if (!sc.asal) { const gs = [...new Set(own.filter(l => l.key === up_(sc.kode)).map(l => l.gudang))]; if (gs.length === 1) sc.asal = gs[0]; else throw new Error('Pilih kota asal barang.'); }
    }
    const done = s.meta.gudangDone;
    if (Object.keys(done).length) {
      if (!sc.terdaftar && !str_(sc.asal)) throw new Error('Pilih kota asal barang.');
      if (lockedFor_(done, staticData_(id).lines, up_(sc.kode), str_(sc.asal))) throw new Error('Kota ' + (str_(sc.asal) || 'asal barang ini') + ' sudah diselesaikan. Minta admin membuka kembali kota itu.');
    }
    const list = rowsOf_('INPUT', 1, id).map(x => scanFromRow_(x.r));
    const sid = str_(sc.id) || Utilities.getUuid();
    if (!list.some(x => x.id === sid)) {
      const now = new Date();
      sheet_('INPUT').appendRow([sid, id, now, me.username, me.nama, txt_(sc.kode), txt_(sc.nama), q, txt_(sc.ket), txt_(sc.asal), sc.terdaftar ? 'YA' : 'TIDAK', '', '', '']);
      list.push({ id: sid, ts: now.getTime(), by: { id: me.username, nama: me.nama }, kode: str_(sc.kode), key: up_(sc.kode), nama: str_(sc.nama), qty: q, ket: str_(sc.ket), asal: str_(sc.asal), undone: false, undoneBy: null, undoneAt: null });
    }
    writeSummary_(s.row, summary, me);
    return scoped_(me) ? scanScope_(me, staticData_(id).lines, list) : list;
  });
}
function apiUndo(token, id, scanId, undo, summary) { try { return apiUndo_(token, id, scanId, undo, summary); } finally { trBump_(); } }
function apiUndo_(token, id, scanId, undo, summary) {
  const me = auth_(token);
  return withLock_(() => {
    const s = sessRow_(id); openOnly_(s);
    const rows = rowsOf_('INPUT', 1, id);
    const hit = rows.find(x => str_(x.r[0]) === scanId);
    if (!hit) throw new Error('Catatan tidak ditemukan.');
    if ((me.role === 'GUDANG' || me.role === 'SERVICE' || me.role === 'SALES') && up_(hit.r[3]) !== me.username) throw new Error('Hanya bisa membatalkan catatan sendiri.');
    if (Object.keys(s.meta.gudangDone).length && lockedFor_(s.meta.gudangDone, staticData_(id).lines, up_(hit.r[5]), str_(hit.r[9]))) throw new Error('Kota asal catatan ini sudah diselesaikan.');
    if (scoped_(me) && str_(hit.r[9]) && !allowed_(me, str_(hit.r[9]))) throw new Error('Catatan ini di luar akses akun kamu.');
    const now = new Date();
    sheet_('INPUT').getRange(hit.row, 12, 1, 3).setValues([undo ? ['YA', me.nama, now] : ['', '', '']]);
    hit.r[11] = undo ? 'YA' : ''; hit.r[12] = undo ? me.nama : ''; hit.r[13] = undo ? now : '';
    writeSummary_(s.row, summary, me);
    const out = rows.map(x => scanFromRow_(x.r));
    return scoped_(me) ? scanScope_(me, staticData_(id).lines, out) : out;
  });
}
function apiFinish(token, id, finish, summary, tl) { try { return apiFinish_(token, id, finish, summary, tl); } finally { trBump_(); } }
function apiFinish_(token, id, finish, summary, tl) {
  const me = auth_(token, 'ADMIN');
  return withLock_(() => {
    const s = sessRow_(id);
    if (s.meta.arsip) throw new Error('Penerimaan yang sudah diarsipkan tidak bisa diubah.');
    sheet_('PENERIMAAN').getRange(s.row, P.PENERIMAAN.STATUS + 1, 1, 3).setValues([finish ? ['SELESAI', new Date(), me.nama] : ['BERJALAN', '', '']]);
    const done = finish ? s.meta.gudangDone : {};
    if (finish) {
      const at = Date.now(), open = allGudang_(id, staticData_(id).lines).filter(g => !done[g]);
      open.forEach(g => { done[g] = { at: at, by: me.nama }; });
      replaceTL_(id, open, tl || [], me);
    }
    writeDone_(s.row, done);
    writeSummary_(s.row, summary);
    const m = sessRow_(id).meta; m.tl = tlOf_(id);
    return m;
  });
}

/* selesaikan / buka kembali satu kota (gudang asal) */
function apiFinishGudang(token, id, gudang, finish, summary, tl) { try { return apiFinishGudang_(token, id, gudang, finish, summary, tl); } finally { trBump_(); } }
function apiFinishGudang_(token, id, gudang, finish, summary, tl) {
  const me = auth_(token, 'ADMIN');
  gudang = str_(gudang);
  return withLock_(() => {
    const s = sessRow_(id);
    if (s.meta.arsip) throw new Error('Penerimaan yang sudah diarsipkan tidak bisa diubah.');
    const all = allGudang_(id, staticData_(id).lines);
    if (all.indexOf(gudang) < 0) throw new Error('Kota ' + gudang + ' tidak ada di penerimaan ini.');
    const done = s.meta.gudangDone;
    const sh = sheet_('PENERIMAAN');
    if (finish) {
      if (s.meta.status !== 'berjalan') throw new Error('Penerimaan ini sudah selesai.');
      done[gudang] = { at: Date.now(), by: me.nama };
      replaceTL_(id, [gudang], tl || [], me);
      if (all.every(g => done[g])) sh.getRange(s.row, P.PENERIMAAN.STATUS + 1, 1, 3).setValues([['SELESAI', new Date(), me.nama]]);
    } else {
      delete done[gudang];
      if (s.meta.status === 'selesai') sh.getRange(s.row, P.PENERIMAAN.STATUS + 1, 1, 3).setValues([['BERJALAN', '', '']]);
    }
    writeDone_(s.row, done);
    writeSummary_(s.row, summary);
    const m = sessRow_(id).meta; m.tl = tlOf_(id);
    return m;
  });
}
function apiDelete(token, id) { try { return apiDelete_(token, id); } finally { trBump_(); } }
function apiDelete_(token, id) {
  auth_(token, 'ADMIN');
  return withLock_(() => { const s = sessRow_(id); sheet_('PENERIMAAN').getRange(s.row, P.PENERIMAAN.STATUS + 1).setValue('DIHAPUS'); return true; });
}

/* tambah IO susulan (mode 'tambah') atau ganti file IO (mode 'ganti') */
function apiEditKiriman(token, id, d, mode) { try { return apiEditKiriman_(token, id, d, mode); } finally { trBump_(); } }
function apiEditKiriman_(token, id, d, mode) {
  auth_(token, 'ADMIN');
  if (!d || !d.lines || !d.lines.length) throw new Error('File kiriman kosong.');
  if (mode !== 'tambah' && mode !== 'ganti') throw new Error('Pilihan tidak dikenal.');
  return withLock_(() => {
    const s = sessRow_(id);
    if (s.meta.arsip) throw new Error('Penerimaan ini sudah diarsipkan.');
    if (s.meta.status !== 'berjalan') throw new Error('Buka kembali penerimaan ini dulu sebelum mengubah IO.');
    const old = rowsOf_('KIRIMAN', 0, id);
    const info = { mode: mode, ditambah: 0, dilewati: [], ioBaru: [] };
    let lines;
    if (mode === 'tambah') {
      const oldLines = old.map(x => lineFromRow_(x.r));
      const ioLama = {}; oldLines.forEach(l => ioLama[l.io] = 1);
      const baru = d.lines.filter(l => !ioLama[str_(l.io)]);
      info.dilewati = [...new Set(d.lines.filter(l => ioLama[str_(l.io)]).map(l => str_(l.io)))];
      if (!baru.length) throw new Error('Semua No. IO di file ini sudah ada di ' + id + '. Pakai "Ganti file IO" kalau isinya berubah.');
      let max = oldLines.reduce((m, l) => Math.max(m, l.line), 0);
      const add = baru.map(l => cleanLine_(Object.assign({}, l, { line: ++max })));
      appendRows_('KIRIMAN', add.map(l => kirimanRow_(id, l, l.line)));
      lines = oldLines.concat(add);
      info.ditambah = add.length; info.ioBaru = [...new Set(add.map(l => l.io))];
      info.dibuka = [...new Set(add.map(l => l.gudang))].filter(g => s.meta.gudangDone[g]);
    } else {
      const sig = (ls, g) => ls.filter(l => l.gudang === g).map(l => l.io + '|' + l.key + '|' + l.qty).sort().join(';');
      const oldL = old.map(x => lineFromRow_(x.r));
      deleteRowsOf_('KIRIMAN', old.map(x => x.row));
      lines = d.lines.map((l, i) => cleanLine_(Object.assign({}, l, { line: i + 1 })));
      info.dibuka = Object.keys(s.meta.gudangDone).filter(g => sig(oldL, g) !== sig(lines, g));
      appendRows_('KIRIMAN', lines.map(l => kirimanRow_(id, l, l.line)));
      info.ditambah = lines.length; info.ioBaru = [...new Set(lines.map(l => l.io))];
    }
    refreshKirimanMeta_(s.row, lines, d.name, mode);
    if (info.dibuka.length) { const done = s.meta.gudangDone; info.dibuka.forEach(g => delete done[g]); writeDone_(s.row, done); }
    saveFile_(sessionFolder_(id, new Date(s.meta.created)), d);
    cacheDel_('sess_' + id);
    const out = apiSession(token, id);
    out.info = info;
    return out;
  });
}
/* hapus satu No. IO yang salah lampir (admin) */
function apiDeleteIO(token, id, io) { try { return apiDeleteIO_(token, id, io); } finally { trBump_(); } }
function apiDeleteIO_(token, id, io) {
  const me = auth_(token, 'ADMIN');
  io = str_(io);
  if (!io) throw new Error('No. IO kosong.');
  return withLock_(() => {
    const s = sessRow_(id);
    if (s.meta.arsip) throw new Error('Penerimaan ini sudah diarsipkan.');
    if (s.meta.status !== 'berjalan') throw new Error('Buka kembali penerimaan ini dulu sebelum menghapus IO.');
    const old = rowsOf_('KIRIMAN', 0, id).map(x => ({ row: x.row, l: lineFromRow_(x.r) }));
    const hapus = old.filter(x => str_(x.l.io) === io);
    if (!hapus.length) throw new Error('No. IO ' + io + ' tidak ditemukan di ' + id + '.');
    const sisa = old.filter(x => str_(x.l.io) !== io);
    if (!sisa.length) throw new Error('Ini satu-satunya No. IO di penerimaan ini. Pakai "Ganti file IO" atau hapus penerimaannya.');
    deleteRowsOf_('KIRIMAN', hapus.map(x => x.row));
    const lines = sisa.map(x => x.l);
    const sh = sheet_('PENERIMAAN');
    const tgl = [...new Set(lines.map(l => l.tanggal))].filter(Boolean).sort().join(',');
    const gud = [...new Set(lines.map(l => l.gudang))].filter(Boolean).join('|');
    sh.getRange(s.row, P.PENERIMAAN.TGL_KIRIMAN + 1, 1, 2).setValues([[txt_(tgl), txt_(gud)]]);
    sh.getRange(s.row, P.PENERIMAAN.BARIS + 1).setValue(lines.length);
    const kota = [...new Set(hapus.map(x => x.l.gudang))];
    const done = s.meta.gudangDone, dibuka = [], hilang = [];
    kota.forEach(g => { const masih = lines.some(l => l.gudang === g); if (!masih) hilang.push(g); if (done[g]) { delete done[g]; if (masih) dibuka.push(g); } });
    writeDone_(s.row, done);
    cacheDel_('sess_' + id);
    const out = apiSession(token, id);
    out.info = { io: io, baris: hapus.length, qty: hapus.reduce((a, x) => a + (Number(x.l.qty) || 0), 0), kota: kota, dibuka: dibuka, hilang: hilang, oleh: me.nama };
    return out;
  });
}
function apiReplaceBarkur(token, id, d) {
  auth_(token, 'ADMIN');
  return withLock_(() => {
    const s = sessRow_(id); openOnly_(s);
    deleteRowsOf_('BARKUR', rowsOf_('BARKUR', 0, id).filter(x => up_(x.r[11]) !== 'YA').map(x => x.row));
    appendRows_('BARKUR', (d.rows || []).map(b => barkurRow_(id, cleanBarkur_(b))));
    sheet_('PENERIMAAN').getRange(s.row, P.PENERIMAAN.FILE_BARKUR + 1).setValue(txt_(d.name));
    saveFile_(sessionFolder_(id, new Date(s.meta.created)), d);
    cacheDel_('sess_' + id);
    return apiSession(token, id);
  });
}
function apiAddIndent(token, id, b) {
  const me = auth_(token, 'ADMIN');
  return withLock_(() => {
    const s = sessRow_(id); openOnly_(s);
    b.manual = true; b.by = me.nama; b.id = 'MAN-' + Utilities.formatDate(new Date(), APP.TZ, 'yyMMddHHmmss');
    appendRows_('BARKUR', [barkurRow_(id, cleanBarkur_(b))]);
    cacheDel_('sess_' + id);
    return apiSession(token, id);
  });
}

/* edit / tambah satu request indent */
/* cari baris indent: ID Permintaan + kode barang (+ urutan kalau kembar), karena satu ID bisa berisi beberapa barang */
function barkurHit_(id, reqId, kode, nth) {
  const rows = rowsOf_('BARKUR', 0, id).filter(x => str_(x.r[1]) === str_(reqId));
  const byKode = str_(kode) ? rows.filter(x => up_(x.r[7]) === up_(kode)) : rows;
  return byKode[(Number(nth) || 1) - 1] || byKode[0] || null;
}
function barkurKeyed_(list, idOf, kodeOf) {
  const c = {}, out = {};
  list.forEach(x => { const base = str_(idOf(x)) + '|' + up_(kodeOf(x)); const n = c[base] = (c[base] || 0) + 1; out[n > 1 ? base + '#' + n : base] = x; });
  return out;
}
function apiSaveIndent(token, id, b) {
  const me = auth_(token, 'ADMIN');
  const q = Number(b.qtyDiminta), terima = Number(b.qtyTerimaAwal) || 0;
  if (!str_(b.kode)) throw new Error('Isi kode barang.');
  if (!(q >= 1) || Math.floor(q) !== q) throw new Error('Qty diminta minimal 1.');
  if (terima < 0 || Math.floor(terima) !== terima) throw new Error('Qty sudah diterima tidak valid.');
  return withLock_(() => {
    const s = sessRow_(id); openOnly_(s);
    const stamp = me.nama + ' · ' + Utilities.formatDate(new Date(), APP.TZ, 'dd/MM HH:mm');
    if (b.id) {
      const hit = barkurHit_(id, b.id, b.origKode, b.nth);
      if (!hit) throw new Error('Request ' + b.id + ' tidak ditemukan.');
      const r = hit.r; let raw = null; try { raw = str_(r[13]) ? JSON.parse(r[13]) : null; } catch (e) {}
      const d = cleanBarkur_({ id: r[1], tglReqRaw: b.tglReqRaw != null ? b.tglReqRaw : r[2], tglReq: r[3], gudangTujuan: b.gudangTujuan != null ? b.gudangTujuan : r[4],
        sales: b.sales, noSO: b.noSO, kode: b.kode, nama: b.nama, qtyDiminta: q, qtyTerimaAwal: terima, manual: up_(r[11]) === 'YA', by: 'edit ' + stamp, raw: raw });
      if (raw) Object.assign(raw, { 'Tanggal Request': d.tglReqRaw, 'Gudang Tujuan': d.gudangTujuan, 'Nama Sales': d.sales, 'No Invoice': d.noSO, 'Kode SKU': d.kode, 'Nama Barang': d.nama, 'Qty Diminta': d.qtyDiminta, 'Qty Terima': d.qtyTerimaAwal });
      d.raw = raw;
      sheet_('BARKUR').getRange(hit.row, 1, 1, HEAD.BARKUR.length).setValues([barkurRow_(id, d)]);
    } else {
      const d = cleanBarkur_(Object.assign({}, b, { id: 'MAN-' + Utilities.formatDate(new Date(), APP.TZ, 'yyMMddHHmmss'), manual: true, by: stamp, qtyDiminta: q, qtyTerimaAwal: terima }));
      appendRows_('BARKUR', [barkurRow_(id, d)]);
    }
    cacheDel_('sess_' + id);
    return apiSession(token, id);
  });
}
function apiDeleteIndent(token, id, bid, kode, nth) {
  auth_(token, 'ADMIN');
  return withLock_(() => {
    const s = sessRow_(id); openOnly_(s);
    const hit = barkurHit_(id, bid, kode, nth);
    if (!hit) throw new Error('Request tidak ditemukan.');
    sheet_('BARKUR').deleteRows(hit.row, 1);
    cacheDel_('sess_' + id);
    return apiSession(token, id);
  });
}
/* update dari file BARKUR terbaru: dicocokkan per ID Permintaan, indent manual tidak disentuh */
function apiMergeBarkur(token, id, d, hapusHilang) {
  const me = auth_(token, 'ADMIN');
  return withLock_(() => {
    const s = sessRow_(id); openOnly_(s);
    const sh = sheet_('BARKUR');
    /* dicocokkan per ID Permintaan + kode barang (satu ID bisa berisi beberapa barang) */
    const map = barkurKeyed_(rowsOf_('BARKUR', 0, id).filter(x => up_(x.r[11]) !== 'YA'), x => x.r[1], x => x.r[7]);
    const info = { baru: 0, berubah: 0, sama: 0, dihapus: 0, tetap: 0 }, seen = {}, adds = [];
    const fileRows = (d.rows || []).map(cleanBarkur_).filter(b => b.id);
    const fileMap = barkurKeyed_(fileRows, b => b.id, b => b.kode);
    Object.keys(fileMap).forEach(k => {
      const b = fileMap[k];
      seen[k] = 1;
      const x = map[k];
      if (!x) { adds.push(barkurRow_(id, b)); info.baru++; return; }
      const nr = barkurRow_(id, b); nr[12] = x.r[12];
      const changed = [2, 4, 5, 6, 7, 8, 9, 10].some(i => String(nr[i]).replace(/^'/, '') !== String(x.r[i]));
      if (changed) { sh.getRange(x.row, 1, 1, HEAD.BARKUR.length).setValues([nr]); info.berubah++; } else info.sama++;
    });
    const hilang = Object.keys(map).filter(k => !seen[k]);
    if (hapusHilang) { deleteRowsOf_('BARKUR', hilang.map(k => map[k].row)); info.dihapus = hilang.length; } else info.tetap = hilang.length;
    appendRows_('BARKUR', adds);
    sheet_('PENERIMAAN').getRange(s.row, P.PENERIMAAN.FILE_BARKUR + 1).setValue(txt_(d.name));
    saveFile_(sessionFolder_(id, new Date(s.meta.created)), d);
    cacheDel_('sess_' + id);
    const out = apiSession(token, id); out.info = info;
    return out;
  });
}

/* ================= TINDAK LANJUT SELISIH ================= */
const TL_AKSI = { RETUR_IO: 'Retur IO ke cabang', KIRIM_ULANG: 'Cabang kirim ulang', KEMBALI_CABANG: 'Kembalikan ke cabang', BUAT_IO: 'Minta cabang buat IO' };
function tlFromRow_(r, row) {
  return { id: str_(r[0]), trm: str_(r[1]), kota: str_(r[2]), io: str_(r[3]), kode: str_(r[4]), key: up_(r[4]), nama: str_(r[5]), jenis: up_(r[6]), qty: Number(r[7]) || 0,
    aksi: up_(r[8]), status: up_(r[9]) || 'MENUNGGU', ket: str_(r[10]), dibuat: ms_(r[11]), oleh: str_(r[12]), selesai: ms_(r[13]), selesaiOleh: str_(r[14]), catatan: str_(r[15]), row: row };
}
function tlSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('TINDAK_LANJUT');
  if (!sh) { sh = ss.insertSheet('TINDAK_LANJUT'); sh.getRange(1, 1, 1, HEAD.TINDAK_LANJUT.length).setValues([HEAD.TINDAK_LANJUT]); }
  return sh;
}
function tlRows_() { tlSheet_(); return rows_('TINDAK_LANJUT').map((r, i) => tlFromRow_(r, i + 2)).filter(t => t.id); }
function tlOf_(id) { tlSheet_(); return rowsOf_('TINDAK_LANJUT', 1, id).map(x => tlFromRow_(x.r, x.row)); }
function replaceTL_(id, kotas, items, me) {
  tlSheet_();
  const set = {}; kotas.forEach(k => set[k] = 1);
  deleteRowsOf_('TINDAK_LANJUT', rowsOf_('TINDAK_LANJUT', 1, id).filter(x => set[str_(x.r[2])]).map(x => x.row));
  const now = new Date(), stamp = Utilities.formatDate(now, APP.TZ, 'yyMMddHHmmss');
  const rows = (items || []).filter(t => set[str_(t.kota)] && TL_AKSI[up_(t.aksi)] && Number(t.qty) > 0).map((t, i) => {
    const aksi = up_(t.aksi), selesai = aksi === 'RETUR_IO';
    return ['TL-' + stamp + '-' + (i + 1), id, txt_(t.kota), txt_(t.io), txt_(t.kode), txt_(t.nama), up_(t.jenis), Number(t.qty), aksi, selesai ? 'SELESAI' : 'MENUNGGU', txt_(t.ket), now, me.nama, selesai ? now : '', selesai ? me.nama : '', selesai ? txt_('IO dikembalikan ke cabang') : ''];
  });
  appendRows_('TINDAK_LANJUT', rows);
}
function apiTL(token) {
  auth_(token, 'ADMIN');
  return tlOverview_();
}
function tlOverview_() {
  const list = tlRows_();
  const metas = allMeta_(), meta = {}; metas.forEach(m => meta[m.id] = m);
  const aktif = metas.filter(m => !m.arsip);
  fillKota_(aktif, 15);
  const hasTL = {}; list.forEach(t => hasTL[t.trm + '|' + t.kota] = 1);
  const perlu = [], jalan = [];
  aktif.forEach(m => {
    const k = m.kota || {};
    Object.keys(m.gudangDone || {}).forEach(g => { const x = k[g]; if (!hasTL[m.id + '|' + g] && x && (x.kurang > 0 || x.tak > 0)) perlu.push({ trm: m.id, kota: g, kurang: x.kurang, tak: x.tak, fisik: x.fisik, data: x.data, at: m.gudangDone[g].at, by: m.gudangDone[g].by }); });
    if (m.status === 'berjalan') m.gudang.forEach(g => { const x = k[g]; if ((m.gudangDone || {})[g] || !x || !(x.kurang > 0)) return; jalan.push({ trm: m.id, kota: g, kurang: x.kurang, kode: x.kode, dicek: x.dicek, fisik: x.fisik, data: x.data, created: m.created }); });
  });
  const wait = list.filter(t => t.status === 'MENUNGGU' && t.aksi === 'KIRIM_ULANG');
  if (wait.length) {
    const want = {}; wait.forEach(t => want[t.kota + '|' + t.key] = 1);
    const hits = {};
    rows_('KIRIMAN').forEach(r => { const k = str_(r[4]) + '|' + up_(r[6]); if (!want[k]) return; (hits[k] = hits[k] || []).push({ trm: str_(r[0]), io: str_(r[2]), qty: Number(r[8]) || 0 }); });
    wait.forEach(t => {
      const h = (hits[t.kota + '|' + t.key] || []).filter(x => x.trm !== t.trm && meta[x.trm] && meta[x.trm].created > t.dibuat);
      if (h.length) t.saran = h[0];
    });
  }
  list.forEach(t => { const m = meta[t.trm]; t.trmCreated = m ? m.created : null; delete t.row; });
  return { rows: list.sort((a, b) => (b.dibuat || 0) - (a.dibuat || 0)), perlu: perlu.sort((a, b) => (b.at || 0) - (a.at || 0)), jalan: jalan.sort((a, b) => (b.created || 0) - (a.created || 0)) };
}
/* isi / ubah tindak lanjut untuk kota yang sudah selesai (tanpa membuka ulang kota) */
function apiSaveTL(token, id, kota, items) {
  const me = auth_(token, 'ADMIN');
  return withLock_(() => { replaceTL_(id, [str_(kota)], items || [], me); return { tl: tlOf_(id) }; });
}
function apiTLUpdate(token, tlId, d) {
  const me = auth_(token, 'ADMIN');
  return withLock_(() => {
    const t = tlRows_().find(x => x.id === tlId);
    if (!t) throw new Error('Data tindak lanjut tidak ditemukan.');
    const sh = sheet_('TINDAK_LANJUT');
    if (d.aksi && TL_AKSI[up_(d.aksi)]) sh.getRange(t.row, 9).setValue(up_(d.aksi));
    if (d.status === 'SELESAI') sh.getRange(t.row, 10, 1, 1).setValue('SELESAI'), sh.getRange(t.row, 14, 1, 3).setValues([[new Date(), me.nama, txt_(d.catatan || '')]]);
    if (d.status === 'MENUNGGU') sh.getRange(t.row, 10, 1, 1).setValue('MENUNGGU'), sh.getRange(t.row, 14, 1, 3).setValues([['', '', '']]);
    return apiTL(token);
  });
}

/* ================= STOK HARIAN ================= */
/* kode barang yang disembunyikan dari stokan (awalan kode), diatur super admin */
function stokRules_() { try { const v = getSetting_('STOK_KECUALI'); if (v) return JSON.parse(v); } catch (e) {} return [{ p: 'OOW-', on: true }, { p: 'IW-', on: true }, { p: 'MOHAWK-', on: true }]; }
function stokExcl_(rules) { const act = (rules || stokRules_()).filter(r => r.on && str_(r.p)).map(r => up_(r.p)); return k => act.some(p => String(k).indexOf(p) === 0); }
function apiStokRules(token) { needMod_(auth_(token), 'STOK'); return stokRules_(); }
function apiSaveStokRules(token, rules) {
  auth_(token, 'SUPERADMIN');
  const clean = [], seen = {};
  (rules || []).forEach(r => { const p = up_(r.p).replace(/\s+/g, ''); if (!p || seen[p]) return; seen[p] = 1; clean.push({ p: p, on: !!r.on }); });
  setSetting_('STOK_KECUALI', JSON.stringify(clean));
  return clean;
}
function stokMeta_(r) {
  let riw = []; try { riw = str_(r[15]) ? JSON.parse(r[15]) : []; } catch (e) {}
  return { id: str_(r[0]), tanggal: str_(r[1]), created: ms_(r[2]), createdBy: str_(r[3]), status: str_(r[4]).toLowerCase() || 'berjalan', uploadKe: Number(r[5]) || 0, uploadAt: ms_(r[6]),
    finished: ms_(r[7]), finishedBy: str_(r[8]), sum: { item: Number(r[9]) || 0, sistem: Number(r[10]) || 0, fisik: Number(r[11]) || 0, ok: Number(r[12]) || 0, cek: Number(r[13]) || 0, belum: Number(r[14]) || 0 },
    riwayat: riw, folderUrl: str_(r[16]), arsip: str_(r[17]), catatan: stokNotes_(r[18]) };
}
function stokNotes_(v) { try { const o = str_(v) ? JSON.parse(str_(v)) : null; if (o && typeof o === 'object') return { umum: str_(o.umum), items: o.items || {}, ba: o.ba && typeof o.ba === 'object' ? o.ba : {} }; } catch (e) {} return { umum: '', items: {}, ba: {} }; }
/* isian berita acara stok opname (jam, cakupan, catatan, kronologi & tindak lanjut per kategori) */
function baClean_(v, depth) {
  if (depth > 3) return null;
  if (v == null) return '';
  if (typeof v === 'boolean' || typeof v === 'number') return v;
  if (typeof v === 'string') return v.slice(0, 1500);
  if (typeof v === 'object') { const o = {}; Object.keys(v).slice(0, 30).forEach(k => { o[String(k).slice(0, 30)] = baClean_(v[k], depth + 1); }); return o; }
  return '';
}
function apiStokBA(token, id, ba) {
  auth_(token, 'ADMIN');
  return withLock_(() => {
    const s = stokRow_(id), n = s.meta.catatan;
    n.ba = baClean_(ba || {}, 0);
    const txt = JSON.stringify(n);
    if (txt.length > 45000) throw new Error('Isian berita acara terlalu panjang.');
    sheet_('STOK_SESI').getRange(s.row, HEAD.STOK_SESI.indexOf('CATATAN') + 1).setValue(txt);
    return n.ba;
  });
}
/* catatan admin untuk stokan: key '' = catatan umum, selain itu catatan per kode barang */
function apiStokNote(token, id, key, text) {
  auth_(token, 'ADMIN');
  return withLock_(() => {
    const s = stokRow_(id), n = s.meta.catatan;
    const k = up_(key), t = str_(text).slice(0, 500);
    if (!k) n.umum = t; else if (t) n.items[k] = t; else delete n.items[k];
    sheet_('STOK_SESI').getRange(s.row, HEAD.STOK_SESI.indexOf('CATATAN') + 1).setValue(JSON.stringify(n));
    return n;
  });
}
/* selisih stokan sebelumnya yang sudah selesai, untuk tanda "kemarin" di stokan berikutnya */
function stokPrev_(meta) {
  const ck = 'stkprev_' + meta.id, cache = CacheService.getScriptCache(), hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const prev = rows_('STOK_SESI').filter(r => str_(r[0])).map(stokMeta_).filter(m => m.status === 'selesai' && m.id !== meta.id && (m.created || 0) < (meta.created || 0)).sort((a, b) => b.created - a.created)[0];
  if (!prev) return null;
  const ss = stokSrc_(prev), items = rowsOf_('STOK_ITEM', 0, prev.id, ss).map(x => stokItemFromRow_(x.r));
  const f = stokFisik_(rowsOf_('STOK_INPUT', 1, prev.id, ss).map(x => stokInputFromRow_(x.r))), sel = {};
  items.forEach(x => { if (x.kunci) return; const fis = f[x.key]; if (fis == null) { if (x.sistem > 0) sel[x.key] = 'B'; return; } if (fis !== x.sistem) sel[x.key] = fis - x.sistem; });
  Object.keys(f).forEach(k => { if (!items.some(x => x.key === k) && f[k]) sel[k] = f[k]; });
  const out = { id: prev.id, tanggal: prev.tanggal, sel: sel, catatan: prev.catatan.items };
  try { cache.put(ck, JSON.stringify(out), 21600); } catch (e) {}
  return out;
}
function stokRow_(id) {
  const row = findRow_('STOK_SESI', 0, id);
  if (row < 0) throw new Error('Stokan ' + id + ' tidak ditemukan.');
  const r = sheet_('STOK_SESI').getRange(row, 1, 1, HEAD.STOK_SESI.length).getValues()[0];
  return { row: row, meta: stokMeta_(r) };
}
function stokItemFromRow_(r) { return { kode: str_(r[1]), key: up_(r[1]), nama: str_(r[2]), sistem: Number(r[3]) || 0, awal: Number(r[4]) || 0, kons: Number(r[5]) || 0, kunci: up_(r[6]) === 'YA', kSistem: Number(r[7]) || 0, kFisik: Number(r[8]) || 0, kAt: ms_(r[9]), kUpload: Number(r[10]) || 0, diUpload: up_(r[11]) !== 'TIDAK' }; }
function stokItemRow_(id, x) { return [id, txt_(x.kode), txt_(x.nama), x.sistem, x.awal, x.kons || 0, x.kunci ? 'YA' : '', x.kunci ? x.kSistem : '', x.kunci ? x.kFisik : '', x.kunci && x.kAt ? new Date(x.kAt) : '', x.kunci ? x.kUpload : '', x.diUpload ? 'YA' : 'TIDAK']; }
function stokInputFromRow_(r) { return { id: str_(r[0]), ts: ms_(r[2]), by: { id: up_(r[3]), nama: str_(r[4]) }, kode: str_(r[5]), key: up_(r[5]), nama: str_(r[6]), qty: Number(r[7]) || 0, ket: str_(r[8]), jenis: up_(r[9]) || 'HITUNG', undone: up_(r[10]) === 'YA', undoneBy: str_(r[11]) ? { nama: str_(r[11]) } : null }; }
function stokSrc_(meta) { return meta.arsip ? SpreadsheetApp.openById(meta.arsip) : null; }
function stokFisik_(inputs) { const f = {}; inputs.forEach(s => { if (!s.undone) f[s.key] = (f[s.key] || 0) + s.qty; }); return f; }
function stokEnsure_() {
  const ss = SpreadsheetApp.getActive();
  ['STOK_SESI', 'STOK_ITEM', 'STOK_INPUT'].forEach(n => { if (!ss.getSheetByName(n)) { const sh = ss.insertSheet(n); sh.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]); } });
}
/* kunci barang yang sudah OK (fisik = sistem) supaya tidak berubah di upload berikutnya */
function stokLockOk_(items, fis, uploadKe) {
  const now = Date.now(); let n = 0;
  items.forEach(x => { if (x.kunci) return; const f = fis[x.key]; if (f == null) return; if (f === x.sistem) { x.kunci = true; x.kSistem = x.sistem; x.kFisik = f; x.kAt = now; x.kUpload = uploadKe; n++; } });
  return n;
}
function stokCacheItems_(id, items) { cachePut_('stkit_' + id, items); CacheService.getScriptCache().put('stklk_' + id, JSON.stringify(items.filter(x => x.kunci).map(x => x.key)), APP.CACHE_DETIK); }
function stokItems_(id, ss) {
  if (!ss) { const c = cacheGet_('stkit_' + id); if (c) return c; }
  const items = rowsOf_('STOK_ITEM', 0, id, ss).map(x => stokItemFromRow_(x.r));
  if (!ss) stokCacheItems_(id, items);
  return items;
}
function stokLockedKeys_(id) {
  const c = CacheService.getScriptCache().get('stklk_' + id);
  if (c) { try { return JSON.parse(c); } catch (e) {} }
  return stokItems_(id).filter(x => x.kunci).map(x => x.key);
}
function stokInputsSince_(id, sinceRow, ss) {
  const sh = sheet_('STOK_INPUT', ss), last = sh.getLastRow();
  if (!sinceRow || sinceRow < 1 || sinceRow > last) return { inputs: rowsOf_('STOK_INPUT', 1, id, ss).map(x => stokInputFromRow_(x.r)), lastRow: last, full: true };
  if (sinceRow === last) return { inputs: [], lastRow: last, full: false };
  const rows = sh.getRange(sinceRow + 1, 1, last - sinceRow, HEAD.STOK_INPUT.length).getValues().filter(r => str_(r[1]) === id).map(stokInputFromRow_);
  return { inputs: rows, lastRow: last, full: false };
}
function stokWriteItems_(id, items) {
  deleteRowsOf_('STOK_ITEM', rowsOf_('STOK_ITEM', 0, id).map(x => x.row));
  appendRows_('STOK_ITEM', items.map(x => stokItemRow_(id, x)));
  stokCacheItems_(id, items);
}
function stokCalc_(items, inputs) {
  const ex = stokExcl_();
  const f = stokFisik_(inputs), s = { item: 0, sistem: 0, fisik: 0, ok: 0, cek: 0, belum: 0 }, seen = {};
  items.forEach(x => {
    seen[x.key] = 1; if (ex(x.key)) return; const fis = f[x.key];
    if (!(x.sistem > 0) && fis == null && !x.kunci) return;
    s.item++; s.sistem += x.kunci ? x.kSistem : x.sistem; s.fisik += x.kunci ? x.kFisik : (fis || 0);
    if (x.kunci || (fis != null && fis === x.sistem)) s.ok++; else if (fis == null) s.belum++; else s.cek++;
  });
  Object.keys(f).forEach(k => { if (!seen[k] && !ex(k) && f[k] !== 0) { s.item++; s.fisik += f[k]; s.cek++; } });
  return s;
}
function stokSum_(row, s) {
  if (!s) return;
  sheet_('STOK_SESI').getRange(row, 10, 1, 6).setValues([[s.item || 0, s.sistem || 0, s.fisik || 0, s.ok || 0, s.cek || 0, s.belum || 0]]);
}
function apiStokIndex(token, period) { needMod_(auth_(token), 'STOK'); stokEnsure_(); return stokIndex_(period); }
function stokIndex_(period) {
  const all = rows_('STOK_SESI').filter(r => str_(r[0])).map(stokMeta_).reverse();
  const p = str_(period), cur = ym_(Date.now());
  const list = !p ? all.filter(m => m.status === 'berjalan' || ym_(m.created) === cur) : /^\d{4}$/.test(p) ? all.filter(m => ym_(m.created).slice(0, 4) === p) : all.filter(m => ym_(m.created) === p);
  const map = {};
  all.forEach(m => { const k = ym_(m.created), x = map[k] || (map[k] = { ym: k, n: 0, ok: 0, cek: 0, belum: 0, item: 0 }); x.n++; x.ok += m.sum.ok; x.cek += m.sum.cek; x.belum += m.sum.belum; x.item += m.sum.item; });
  const today = 'STK-' + Utilities.formatDate(new Date(), APP.TZ, 'yyMMdd');
  const aktif = all.find(m => m.status === 'berjalan' && !m.arsip) || null;
  return { index: list, periods: Object.keys(map).sort().reverse().map(k => map[k]), today: all.find(m => m.id === today) || null, aktif: aktif, hariIni: Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd'), kecuali: stokRules_(), petugasInput: stokPetugasInputOn_() };
}
/* cari barang di riwayat stokan (periode '' = berjalan & bulan ini, 'yyyy-MM', atau 'yyyy') */
/* id stokan yang sedang berjalan (paling baru), supaya HP petugas otomatis pindah ke stokan baru */
function stokAktifId_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('stkaktif');
  if (hit != null) return hit === '-' ? '' : hit;
  const a = rows_('STOK_SESI').filter(r => str_(r[0]) && up_(r[4]) === 'BERJALAN' && !str_(r[17])).map(stokMeta_).sort((x, y) => (y.created || 0) - (x.created || 0))[0];
  const id = a ? a.id : '';
  cache.put('stkaktif', id || '-', 30);
  return id;
}
/* hasil cek fisik Terima Barang (barang ada di IO) sejak stokan dimulai ikut dihitung sebagai fisik stokan */
function stokDariTerimaOn_() { return up_(getSetting_('STOK_DARI_TERIMA') || 'YA') !== 'TIDAK'; }
function stokTerimaInputs_(meta) {
  if (!meta || !stokDariTerimaOn_()) return [];
  const fin = meta.status === 'selesai';
  const ck = (fin ? 'stktrF_' : 'stktr_') + meta.id, cache = CacheService.getScriptCache(), ver = fin ? 'F' : (cache.get('trver') || '');
  const hit = cache.get(ck);
  if (hit) { try { const o = JSON.parse(hit); if (o.v === ver) return o.list; } catch (e) {} }
  const from = meta.created || 0, to = meta.status === 'selesai' && meta.finished ? meta.finished : Date.now() + 60000;
  const list = rows_('INPUT').filter(r => str_(r[0]) && up_(r[10]) !== 'TIDAK').map(r => { const s = scanFromRow_(r); s.trm = str_(r[1]); return s; })
    .filter(s => s.ts >= from && s.ts <= to && s.qty !== 0)
    .map(s => ({ id: 'T:' + s.id, ts: s.ts, by: s.by, kode: s.kode, key: s.key, nama: s.nama, qty: s.qty, ket: s.trm, jenis: 'TERIMA', undone: s.undone, undoneBy: s.undoneBy }));
  try { cache.put(ck, JSON.stringify({ v: ver, list: list }), fin ? 21600 : 600); } catch (e) {}
  return list;
}
/* IO kiriman cabang yang sedang diterima (berjalan / dibuat sejak stokan dimulai): qty di IO dan yang sudah dicek di Terima Barang */
function stokIncoming_(meta) {
  if (!meta || !stokDariTerimaOn_()) return {};
  const fin = meta.status === 'selesai';
  const ck = (fin ? 'stkincF_' : 'stkinc_') + meta.id, cache = CacheService.getScriptCache(), ver = fin ? 'F' : (cache.get('trver') || '');
  const hit = cache.get(ck);
  if (hit) { try { const o = JSON.parse(hit); if (o.v === ver) return o.m; } catch (e) {} }
  const to = meta.status === 'selesai' && meta.finished ? meta.finished : Date.now() + 60000;
  const ids = {}; allMeta_().forEach(m => { if (!m.arsip && ((m.status === 'berjalan' && (m.created || 0) <= to) || ((m.created || 0) >= (meta.created || 0) && (m.created || 0) <= to))) ids[m.id] = 1; });
  const out = {};
  if (Object.keys(ids).length) {
    rows_('KIRIMAN').forEach(r => { if (!ids[str_(r[0])]) return; const k = up_(r[6]); const o = out[k] || (out[k] = { io: 0, rec: 0 }); o.io += Number(r[8]) || 0; });
    rows_('INPUT').forEach(r => { if (!ids[str_(r[1])] || up_(r[11]) === 'YA' || up_(r[10]) === 'TIDAK') return; const k = up_(r[5]); if (!out[k]) return; out[k].rec += Number(r[7]) || 0; });
  }
  try { cache.put(ck, JSON.stringify({ v: ver, m: out }), fin ? 21600 : 600); } catch (e) {}
  return out;
}
function stokPetugasInputOn_() { return up_(getSetting_('STOK_PETUGAS_INPUT') || 'YA') !== 'TIDAK'; }
function apiStokSetting(token) { needMod_(auth_(token), 'STOK'); return { dariTerima: stokDariTerimaOn_(), petugasInput: stokPetugasInputOn_() }; }
function apiStokSetPetugasInput(token, on) { auth_(token, 'SUPERADMIN'); setSetting_('STOK_PETUGAS_INPUT', on ? 'YA' : 'TIDAK'); return stokPetugasInputOn_(); }
function apiStokSetDariTerima(token, on) { auth_(token, 'SUPERADMIN'); setSetting_('STOK_DARI_TERIMA', on ? 'YA' : 'TIDAK'); try { CacheService.getScriptCache().put('trver', String(Date.now()), 21600); } catch (e) {} return stokDariTerimaOn_(); }
function apiStokSearch(token, q, period) {
  needMod_(auth_(token), 'STOK'); stokEnsure_();
  const toks = str_(q).toLowerCase().split(/\s+/).filter(Boolean); if (!toks.length) return [];
  const all = rows_('STOK_SESI').filter(r => str_(r[0])).map(stokMeta_);
  const p = str_(period), cur = ym_(Date.now());
  const list = !p ? all.filter(m => m.status === 'berjalan' || ym_(m.created) === cur) : /^\d{4}$/.test(p) ? all.filter(m => ym_(m.created).slice(0, 4) === p) : all.filter(m => ym_(m.created) === p);
  const meta = {}; list.forEach(m => meta[m.id] = m);
  const srcs = {}; list.forEach(m => srcs[m.arsip || ''] = 1);
  const ex = stokExcl_(), hit = s => toks.every(t => s.toLowerCase().indexOf(t) >= 0), agg = {};
  Object.keys(srcs).forEach(src => {
    const ss = src ? SpreadsheetApp.openById(src) : null;
    rows_('STOK_ITEM', ss).forEach(r => { const id = str_(r[0]); if (!meta[id]) return; const x = stokItemFromRow_(r); if (ex(x.key) || !hit(x.kode + ' ' + x.nama)) return; agg[id + '|' + x.key] = { id: id, kode: x.kode, key: x.key, nama: x.nama, sistem: x.kunci ? x.kSistem : x.sistem, kunci: x.kunci, fisik: null, pinjam: 0, petugas: [] }; });
    rows_('STOK_INPUT', ss).forEach(r => { const id = str_(r[1]); if (!meta[id] || up_(r[10]) === 'YA') return; const k = id + '|' + up_(r[5]); const g = agg[k] || (hit(str_(r[5]) + ' ' + str_(r[6])) && !ex(up_(r[5])) ? (agg[k] = { id: id, kode: str_(r[5]), key: up_(r[5]), nama: str_(r[6]), sistem: 0, kunci: false, fisik: null, pinjam: 0, petugas: [] }) : null); if (!g) return;
      g.fisik = (g.fisik || 0) + (Number(r[7]) || 0); if (['PINJAM', 'KEEP', 'RUSAK', 'DISPLAY', 'LAINNYA'].indexOf(up_(r[9])) >= 0) g.pinjam += Number(r[7]) || 0; if (g.petugas.indexOf(str_(r[4])) < 0) g.petugas.push(str_(r[4])); });
  });
  return Object.keys(agg).map(k => { const g = agg[k]; const m = meta[g.id]; g.tanggal = m.tanggal; g.status = m.status; return g; })
    .sort((a, b) => (b.tanggal || '').localeCompare(a.tanggal || '')).slice(0, 300);
}
function apiStokSession(token, id) { needMod_(auth_(token), 'STOK'); stokEnsure_(); return stokSession_(id); }
function stokSession_(id) {
  const s = stokRow_(id), ss = stokSrc_(s.meta);
  const inp = stokInputsSince_(id, 0, ss);
  let prev = null; try { prev = stokPrev_(s.meta); } catch (e) {}
  return { meta: s.meta, items: stokItems_(id, ss), inputs: inp.inputs, lastRow: inp.lastRow, prev: prev, terima: ss ? [] : stokTerimaInputs_(s.meta), incoming: ss ? {} : stokIncoming_(s.meta), trver: s.meta.status === 'selesai' ? 'F' : (CacheService.getScriptCache().get('trver') || ''), dariTerima: stokDariTerimaOn_() };
}
/* sinkron ringan: hanya input baru sejak baris terakhir yang diketahui (sinceRow); 0 = ambil semua */
function apiStokInputs(token, id, sinceRow, trver) { needMod_(auth_(token), 'STOK'); return stokDelta_(id, sinceRow, trver); }
function stokDelta_(id, sinceRow, trver) {
  const s = stokRow_(id), ss = stokSrc_(s.meta), cache = CacheService.getScriptCache();
  const inp = stokInputsSince_(id, Number(sinceRow) || 0, ss);
  const tv = cache.get('trver') || '';
  const out = { meta: s.meta, inputs: inp.inputs, lastRow: inp.lastRow, full: inp.full, ver: cache.get('stkver_' + id) || '', aktif: stokAktifId_(), trver: tv };
  if (s.meta.status === 'selesai') { out.trver = 'F'; if (!ss && str_(trver) !== 'F') { out.terima = stokTerimaInputs_(s.meta); out.incoming = stokIncoming_(s.meta); } }
  else if (!ss && str_(trver) !== tv) { out.terima = stokTerimaInputs_(s.meta); out.incoming = stokIncoming_(s.meta); }
  return out;
}
/* upload stok sistem: stokan baru hari ini, atau update stok terbaru di hari yang sama */
function apiStokUpload(token, d) {
  const me = auth_(token, 'ADMIN'); needMod_(me, 'STOK'); stokEnsure_();
  if (!d || !d.items) throw new Error('Data stok kosong.');
  return withLock_(() => {
    const now = new Date();
    /* stokan yang masih berjalan (walaupun dari hari sebelumnya) dilanjutkan; stokan baru hanya kalau tidak ada yang berjalan */
    const open = rows_('STOK_SESI').filter(r => str_(r[0])).map(stokMeta_).filter(m => m.status === 'berjalan' && !m.arsip).sort((a, b) => (b.created || 0) - (a.created || 0))[0];
    const id = open ? open.id : 'STK-' + Utilities.formatDate(now, APP.TZ, 'yyMMdd');
    const incoming = {}; d.items.forEach(x => { const k = up_(x.kode); if (k) incoming[k] = { kode: str_(x.kode), key: k, nama: str_(x.nama), qty: Number(x.qty) || 0, kons: Number(x.kons) || 0 }; });
    let row = findRow_('STOK_SESI', 0, id), info = { baru: false, terkunci: 0, dikunciSekarang: 0, berubah: 0, tambah: 0, hilang: 0 };
    let items;
    if (row < 0) {
      items = Object.values(incoming).filter(x => x.qty > 0 || x.kons > 0).map(x => ({ kode: x.kode, key: x.key, nama: x.nama, sistem: x.qty, awal: x.qty, kons: x.kons, kunci: false, diUpload: true }));
      const folder = stokFolder_(id, now); if (folder) saveFile_(folder, d);
      sheet_('STOK_SESI').appendRow([id, txt_(Utilities.formatDate(now, APP.TZ, 'yyyy-MM-dd')), now, me.nama, 'BERJALAN', 1, now, '', '', 0, 0, 0, 0, 0, 0, JSON.stringify([{ ke: 1, at: now.getTime(), by: me.nama, file: str_(d.name), item: items.length, qty: items.reduce((a, x) => a + x.sistem, 0) }]), folder ? folder.getUrl() : '', '']);
      row = findRow_('STOK_SESI', 0, id);
      appendRows_('STOK_ITEM', items.map(x => stokItemRow_(id, x)));
      stokCacheItems_(id, items);
      info.baru = true;
    } else {
      const s = stokRow_(id);
      if (s.meta.status !== 'berjalan') throw new Error('Stokan hari ini sudah diselesaikan. Buka kembali dulu untuk upload ulang.');
      items = rowsOf_('STOK_ITEM', 0, id).map(x => stokItemFromRow_(x.r));
      const fis = stokFisik_(rowsOf_('STOK_INPUT', 1, id).map(x => stokInputFromRow_(x.r)).concat(stokTerimaInputs_(s.meta)));
      info.dikunciSekarang = stokLockOk_(items, fis, s.meta.uploadKe);
      const have = {};
      items.forEach(x => {
        have[x.key] = 1;
        const n = incoming[x.key];
        x.diUpload = !!n;
        if (n) x.kons = n.kons;
        if (x.kunci) { info.terkunci++; return; }
        const baru = n ? n.qty : 0;
        if (baru !== x.sistem) { info.berubah++; x.sistem = baru; }
        if (!n) info.hilang++;
      });
      Object.values(incoming).forEach(n => { if (have[n.key] || !(n.qty > 0 || n.kons > 0)) return; items.push({ kode: n.kode, key: n.key, nama: n.nama, sistem: n.qty, awal: n.qty, kons: n.kons, kunci: false, diUpload: true }); info.tambah++; });
      stokWriteItems_(id, items);
      const ke = s.meta.uploadKe + 1, riw = s.meta.riwayat.concat([{ ke: ke, at: now.getTime(), by: me.nama, file: str_(d.name), item: Object.keys(incoming).length, qty: Object.values(incoming).reduce((a, x) => a + x.qty, 0), berubah: info.berubah, dikunci: info.dikunciSekarang }]);
      sheet_('STOK_SESI').getRange(row, 6, 1, 2).setValues([[ke, now]]);
      sheet_('STOK_SESI').getRange(row, 16).setValue(JSON.stringify(riw));
      try { const f = stokFolder_(id, now); if (f) saveFile_(f, d); } catch (e) {}
    }
    CacheService.getScriptCache().remove('stkaktif');
    const out = apiStokSession(token, id); out.info = info;
    stokSum_(row, stokCalc_(out.items, out.inputs.concat(out.terima || [])));
    return out;
  });
}
function stokFolder_(id, when) {
  const rootId = getSetting_('FOLDER_ID'); if (!rootId) return null;
  try { const root = DriveApp.getFolderById(rootId); return subFolder_(subFolder_(subFolder_(subFolder_(root, 'STOK HARIAN'), String(when.getFullYear())), Utilities.formatDate(when, APP.TZ, 'MM')), id); } catch (e) { return null; }
}
function apiStokInput(token, id, sc, summary) {
  const me = needMod_(auth_(token), 'STOK');
  const admin = me.role === 'ADMIN' || me.role === 'SUPERADMIN';
  const jn = up_(sc.jenis), q = Number(sc.qty), koreksi = !!sc.koreksi, pinjam = !koreksi && ['PINJAM', 'KEEP', 'RUSAK', 'DISPLAY', 'LAINNYA'].indexOf(jn) >= 0;
  if (Math.floor(q) !== q) throw new Error('Qty harus angka bulat.');
  if (koreksi && !admin) throw new Error('Koreksi hanya bisa dilakukan admin.');
  if (!koreksi && q < 0) throw new Error('Qty tidak boleh minus.');
  if (pinjam && q < 1) throw new Error('Isi qty barang.');
  if (pinjam && (jn === 'PINJAM' || jn === 'KEEP' || jn === 'LAINNYA') && !str_(sc.ket)) throw new Error(jn === 'KEEP' ? 'Isi nama sales yang keep.' : jn === 'LAINNYA' ? 'Isi keterangan untuk pilihan Lainnya.' : 'Isi siapa yang meminjam.');
  if (koreksi && q === 0) throw new Error('Koreksi tidak mengubah apa pun.');
  const s = stokRow_(id);
  if (s.meta.arsip || s.meta.status !== 'berjalan') throw new Error('Stokan ini sudah selesai.');
  const key = up_(sc.kode), cache = CacheService.getScriptCache();
  const sid = str_(sc.id) || Utilities.getUuid();
  if (cache.get('stkin_' + sid)) return { ok: true, id: sid, dup: true };
  let unlocked = false;
  if (stokLockedKeys_(id).indexOf(key) >= 0) {
    if (!admin) throw new Error('Barang ini sudah OK dan terkunci. Hubungi admin kalau perlu koreksi.');
    if (!koreksi) throw new Error('Barang ini sudah terkunci OK. Pakai Koreksi kalau perlu diubah.');
    withLock_(() => {
      const hit = rowsOf_('STOK_ITEM', 0, id).find(x => up_(x.r[1]) === key);
      if (hit) sheet_('STOK_ITEM').getRange(hit.row, 7, 1, 5).setValues([['', '', '', '', '']]);
      cacheDel_('stkit_' + id); cache.remove('stklk_' + id);
    });
    unlocked = true;
  }
  /* tambah satu baris saja, tanpa antre kunci global, supaya banyak penghitung bisa simpan bersamaan */
  const now = new Date();
  sheet_('STOK_INPUT').appendRow([sid, id, now, me.username, me.nama, txt_(sc.kode), txt_(sc.nama), q, txt_(sc.ket), koreksi ? 'KOREKSI' : pinjam ? jn : 'HITUNG', '', '', '']);
  cache.put('stkin_' + sid, '1', APP.SESI_DETIK);
  try { stokSum_(s.row, summary); } catch (e) {}
  return { ok: true, id: sid, ts: now.getTime(), unlocked: unlocked };
}
function apiStokUndo(token, id, inputId, undo, summary) {
  const me = auth_(token, 'ADMIN');
  return withLock_(() => {
    const s = stokRow_(id);
    if (s.meta.arsip || s.meta.status !== 'berjalan') throw new Error('Stokan ini sudah selesai.');
    const rows = rowsOf_('STOK_INPUT', 1, id), hit = rows.find(x => str_(x.r[0]) === inputId);
    if (!hit) throw new Error('Input tidak ditemukan.');
    sheet_('STOK_INPUT').getRange(hit.row, 11, 1, 3).setValues([undo ? ['YA', me.nama, new Date()] : ['', '', '']]);
    CacheService.getScriptCache().put('stkver_' + id, String(Date.now()), APP.SESI_DETIK);
    stokSum_(s.row, summary);
    return { inputs: rowsOf_('STOK_INPUT', 1, id).map(x => stokInputFromRow_(x.r)) };
  });
}
function apiStokFinish(token, id, finish, summary, catatan) {
  const me = auth_(token, 'ADMIN');
  return withLock_(() => {
    const s = stokRow_(id);
    if (s.meta.arsip) throw new Error('Stokan yang sudah diarsipkan tidak bisa diubah.');
    if (finish) {
      const items = rowsOf_('STOK_ITEM', 0, id).map(x => stokItemFromRow_(x.r));
      const fis = stokFisik_(rowsOf_('STOK_INPUT', 1, id).map(x => stokInputFromRow_(x.r)).concat(stokTerimaInputs_(s.meta)));
      if (stokLockOk_(items, fis, s.meta.uploadKe)) stokWriteItems_(id, items); else stokCacheItems_(id, items);
    }
    if (finish && catatan != null) { const n = s.meta.catatan; n.umum = str_(catatan).slice(0, 500); sheet_('STOK_SESI').getRange(s.row, HEAD.STOK_SESI.indexOf('CATATAN') + 1).setValue(JSON.stringify(n)); }
    sheet_('STOK_SESI').getRange(s.row, 5).setValue(finish ? 'SELESAI' : 'BERJALAN');
    CacheService.getScriptCache().removeAll(['stkaktif', 'stktrF_' + id, 'stkincF_' + id]);
    sheet_('STOK_SESI').getRange(s.row, 8, 1, 2).setValues([finish ? [new Date(), me.nama] : ['', '']]);
    const out = apiStokSession(token, id);
    stokSum_(s.row, stokCalc_(out.items, out.inputs.concat(out.terima || [])));
    return out;
  });
}
function arsipStok_(pred) {
  const ps = sheet_('STOK_SESI'), prow = rows_('STOK_SESI');
  const byYear = {}, rowOf = {};
  prow.forEach((r, i) => { const m = stokMeta_(r); if (!m.id || m.arsip || m.status !== 'selesai' || !pred(m)) return; const y = new Date(m.created).getFullYear(); (byYear[y] = byYear[y] || {})[m.id] = 1; rowOf[m.id] = i + 2; });
  const years = Object.keys(byYear); if (!years.length) return 0;
  const all = { STOK_ITEM: rows_('STOK_ITEM'), STOK_INPUT: rows_('STOK_INPUT') }; let total = 0;
  years.forEach(y => {
    const ids = byYear[y];
    let aid = getSetting_('ARSIP_' + y), ass;
    if (aid) ass = SpreadsheetApp.openById(aid);
    else { ass = SpreadsheetApp.create('ARSIP PORTAL OPERASIONAL ' + y); aid = ass.getId(); setSetting_('ARSIP_' + y, aid); try { const root = DriveApp.getFolderById(getSetting_('FOLDER_ID')); DriveApp.getFileById(aid).moveTo(subFolder_(root, 'ARSIP')); } catch (e) {} }
    ['STOK_ITEM', 'STOK_INPUT'].forEach(name => {
      const col = name === 'STOK_INPUT' ? 1 : 0;
      let ash = ass.getSheetByName(name); if (!ash) { ash = ass.insertSheet(name); ash.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]); }
      const move = all[name].filter(r => ids[str_(r[col])]);
      if (move.length) ash.getRange(ash.getLastRow() + 1, 1, move.length, HEAD[name].length).setValues(move.map(r => r.map(protectText_)));
      all[name] = all[name].filter(r => !ids[str_(r[col])]);
    });
    Object.keys(ids).forEach(id => { ps.getRange(rowOf[id], HEAD.STOK_SESI.length).setValue(aid); total++; });
  });
  ['STOK_ITEM', 'STOK_INPUT'].forEach(name => { const sh = sheet_(name), n = sh.getLastRow(); if (n > 1) sh.getRange(2, 1, n - 1, HEAD[name].length).clearContent(); const keep = all[name]; if (keep.length) sh.getRange(2, 1, keep.length, HEAD[name].length).setValues(keep.map(r => r.map(protectText_))); });
  return total;
}

/* ================= CARI ARSIP ================= */

function years_() { return rows_('SETTING').filter(r => /^ARSIP_\d{4}$/.test(str_(r[0]))).map(r => str_(r[0]).slice(6)).sort().reverse(); }
function apiYears(token) { auth_(token); return years_(); }
function apiSearch(token, q, tahun) {
  const u = auth_(token);
  const toks = str_(q).toLowerCase().split(/\s+/).filter(Boolean);
  if (!toks.length) return [];
  const th = str_(tahun) || Utilities.formatDate(new Date(), APP.TZ, 'yyyy');
  const meta = {};
  rows_('PENERIMAAN').forEach(r => { if (up_(r[8]) === 'DIHAPUS') return; const m = metaFromRow_(r); if (ym_(m.created).slice(0, 4) === th) meta[m.id] = m; });
  const srcs = [null]; const aid = getSetting_('ARSIP_' + th); if (aid) srcs.push(SpreadsheetApp.openById(aid));
  const hit = s => toks.every(t => s.toLowerCase().indexOf(t) >= 0);
  const agg = {};
  const get = (id, key, kode, nama) => { const k = id + '|' + key; return agg[k] || (agg[k] = { id: id, key: key, kode: kode, nama: nama, data: 0, fisik: 0, ios: [], terdaftar: false, petugas: [] }); };
  srcs.forEach(ss => {
  rows_('KIRIMAN', ss).forEach(r => { const id = str_(r[0]); if (!meta[id] || !allowed_(u, str_(r[4]))) return; const kode = str_(r[6]), nama = str_(r[7]), io = str_(r[2]);
    if (!hit(kode + ' ' + nama + ' ' + io)) return; const g = get(id, up_(kode), kode, nama); g.data += Number(r[8]) || 0; g.terdaftar = true; if (g.ios.indexOf(io) < 0) g.ios.push(io); });
  });
  srcs.forEach(ss => {
  rows_('INPUT', ss).forEach(r => { const id = str_(r[1]); if (!meta[id] || up_(r[11]) === 'YA' || (scoped_(u) && !allowed_(u, str_(r[9])))) return; const kode = str_(r[5]), nama = str_(r[6]);
    const k = id + '|' + up_(kode); if (!agg[k] && !hit(kode + ' ' + nama)) return; const g = get(id, up_(kode), kode, nama); g.fisik += Number(r[7]) || 0; if (g.petugas.indexOf(str_(r[4])) < 0) g.petugas.push(str_(r[4])); });
  });
  return Object.keys(agg).map(k => { const g = agg[k]; g.meta = metaFor_(u, meta[g.id]); return g; })
    .sort((a, b) => (b.meta.created || 0) - (a.meta.created || 0)).slice(0, 300);
}

/* ================= ARSIP TAHUNAN ================= */

function menuArsip() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Arsipkan tahun', 'Tahun yang mau diarsipkan (contoh: 2026). Hanya penerimaan berstatus SELESAI yang dipindah.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  ui.alert(arsipkanTahun(Number(r.getResponseText())));
}
function arsipkanTahun(tahun) {
  if (!(tahun > 2000)) throw new Error('Tahun tidak valid.');
  return arsipkan_(m => new Date(m.created).getFullYear() === tahun);
}
/* otomatis tiap tanggal 1: arsipkan penerimaan SELESAI yang dibuat sebelum bulan lalu */
function arsipOtomatis() {
  const now = new Date(), batas = new Date(now.getFullYear(), now.getMonth() - 1, 1).getTime();
  const a = arsipkan_(m => m.created < batas);
  let b = 0; try { stokEnsure_(); b = withLock_(() => arsipStok_(m => m.created < batas)); } catch (e) {}
  return a + (b ? ' · ' + b + ' stokan harian diarsipkan.' : '');
}
function menuArsipSekarang() { SpreadsheetApp.getUi().alert(arsipOtomatis()); }
function pasangArsipOtomatis() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'arsipOtomatis').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('arsipOtomatis').timeBased().onMonthDay(1).atHour(2).create();
  SpreadsheetApp.getUi().alert('Arsip otomatis aktif. Setiap tanggal 1 jam 02.00, penerimaan berstatus SELESAI yang dibuat sebelum bulan lalu dipindah ke file arsip tahunannya. Data tetap bisa dibuka dari web.');
}
function arsipkan_(pred) {
  return withLock_(() => {
    const psh = sheet_('PENERIMAAN'), prow = rows_('PENERIMAAN');
    const byYear = {}, rowOf = {};
    prow.forEach((r, i) => { const m = metaFromRow_(r); if (!m.id || m.arsip || m.status !== 'selesai' || !pred(m)) return; const y = new Date(m.created).getFullYear(); (byYear[y] = byYear[y] || {})[m.id] = 1; rowOf[m.id] = i + 2; });
    const years = Object.keys(byYear);
    if (!years.length) return 'Tidak ada penerimaan yang perlu diarsipkan.';
    const all = {}; ['KIRIMAN', 'BARKUR', 'INPUT'].forEach(n => all[n] = rows_(n));
    let total = 0;
    years.forEach(y => {
      const ids = byYear[y];
      let aid = getSetting_('ARSIP_' + y), ass;
      if (aid) ass = SpreadsheetApp.openById(aid);
      else {
        ass = SpreadsheetApp.create('ARSIP PORTAL OPERASIONAL ' + y); aid = ass.getId();
        try { const root = DriveApp.getFolderById(getSetting_('FOLDER_ID')); DriveApp.getFileById(aid).moveTo(subFolder_(root, 'ARSIP')); } catch (e) {}
        setSetting_('ARSIP_' + y, aid);
      }
      ['KIRIMAN', 'BARKUR', 'INPUT'].forEach(name => {
        const col = name === 'INPUT' ? 1 : 0;
        let ash = ass.getSheetByName(name);
        if (!ash) { ash = ass.insertSheet(name); ash.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]); }
        const move = all[name].filter(r => ids[str_(r[col])]);
        if (move.length) ash.getRange(ash.getLastRow() + 1, 1, move.length, HEAD[name].length).setValues(move.map(r => r.map(protectText_)));
        all[name] = all[name].filter(r => !ids[str_(r[col])]);
      });
      const def = ass.getSheetByName('Sheet1') || ass.getSheetByName('Lembar1');
      if (def && ass.getSheets().length > 1) ass.deleteSheet(def);
      Object.keys(ids).forEach(id => { psh.getRange(rowOf[id], P.PENERIMAAN.ARSIP_ID + 1).setValue(aid); cacheDel_('sess_' + id); total++; });
    });
    ['KIRIMAN', 'BARKUR', 'INPUT'].forEach(name => {
      const sh = sheet_(name), n = sh.getLastRow();
      if (n > 1) sh.getRange(2, 1, n - 1, HEAD[name].length).clearContent();
      const keep = all[name];
      if (keep.length) sh.getRange(2, 1, keep.length, HEAD[name].length).setValues(keep.map(r => r.map(protectText_)));
      if (sh.getMaxRows && sh.getMaxRows() > keep.length + 500) { try { sh.deleteRows(keep.length + 2, sh.getMaxRows() - keep.length - 501); } catch (e) {} }
    });
    return total + ' penerimaan dipindah ke arsip tahunan (' + years.join(', ') + '). Sheet utama sekarang lebih ringan.';
  });
}
function protectText_(v) { return typeof v === 'string' && v !== '' && !/^h\$/.test(v) ? "'" + v : v; }

// AKHIR FILE CODE v10.1
