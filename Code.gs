/**
 * Profitability Sync — Google Sheets → Firestore (project profitability-gov3)
 *
 * Membaca sheet "Summary", "Breakdown Pareto", dan "Rekap SAP S3" di spreadsheet ini,
 * menyusun dataset dashboard, lalu menulisnya ke Firestore dokumen datasets/segment3-ytd2026.
 * Dashboard yang sedang terbuka langsung ikut ter-update (listener realtime).
 *
 * Alur otomatis:
 *   - Setiap perubahan di sheet  → trigger onChange menandai "ada perubahan".
 *   - Setiap 1 menit            → jika ada perubahan, data dibaca ulang dan dikirim
 *                                  (hanya bila isinya benar-benar berbeda, dicek via hash).
 * Penulisan memakai akun Firebase khusus ber-role "Sync", sehingga tunduk pada firestore.rules.
 */

const CFG = {
  PROJECT_ID: 'profitability-gov3',
  DATASET_ID: 'segment3-ytd2026',
  SHEET_SUMMARY: 'Summary',
  SHEET_PARETO: 'Breakdown Pareto',
  SHEET_LEDGER: 'Rekap SAP S3',
  SHEET_LOG: 'Log Sinkronisasi',
  DEBOUNCE_SEC: 20,           // tunggu sampai pengeditan berhenti sejenak sebelum sinkron
};
const MONTHS_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const PROP = PropertiesService.getScriptProperties();

/* ========================= MENU ========================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Profitability Sync')
    .addItem('Sinkronkan sekarang', 'syncNow')
    .addItem('Pratinjau data (tanpa kirim)', 'previewDataset')
    .addSeparator()
    .addItem('1. Atur koneksi Firebase', 'setupConnection')
    .addItem('2. Aktifkan sinkron otomatis', 'installTriggers')
    .addItem('Matikan sinkron otomatis', 'removeTriggers')
    .addItem('Cek status', 'showStatus')
    .addToUi();
}

/* ========================= SETUP ========================= */
function setupConnection() {
  const ui = SpreadsheetApp.getUi();
  const ask = (title, msg, key, secret) => {
    const cur = PROP.getProperty(key);
    const r = ui.prompt(title, msg + (cur ? `\n\nNilai saat ini: ${secret ? '•••• (tersimpan)' : cur}\nKosongkan untuk tetap memakai nilai ini.` : ''), ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) throw new Error('Dibatalkan.');
    const v = r.getResponseText().trim();
    if (v) PROP.setProperty(key, v);
    else if (!cur) throw new Error(`${title} wajib diisi.`);
  };
  try {
    ask('Web API Key', 'Firebase Console → Project settings → General → Web API Key.', 'FIREBASE_API_KEY', false);
    ask('Email akun Sync', 'Email akun Firebase yang dibuat khusus untuk sinkronisasi (role "Sync").', 'SYNC_EMAIL', false);
    ask('Password akun Sync', 'Password akun Sync. Disimpan di Script Properties, tidak terlihat di sheet.', 'SYNC_PASSWORD', true);
  } catch (e) { ui.alert(e.message); return; }
  CacheService.getScriptCache().remove('idToken');
  try {
    const t = getIdToken_();
    ui.alert(`Koneksi berhasil.\nLogin sebagai ${t.email}.\n\nLanjutkan ke menu "2. Aktifkan sinkron otomatis".`);
  } catch (e) { ui.alert('Koneksi gagal: ' + e.message); }
}

function installTriggers() {
  removeTriggers(true);
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onSheetChange_').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('onSheetChange_').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('syncIfDirty').timeBased().everyMinutes(1).create();
  PROP.setProperty('DIRTY', String(Date.now() - CFG.DEBOUNCE_SEC * 1000));
  SpreadsheetApp.getUi().alert('Sinkron otomatis aktif. Perubahan di sheet akan tampil di dashboard dalam ±1–2 menit.');
}

function removeTriggers(silent) {
  ScriptApp.getProjectTriggers()
    .filter(t => ['onSheetChange_', 'syncIfDirty'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (silent !== true) SpreadsheetApp.getUi().alert('Sinkron otomatis dimatikan.');
}

function showStatus() {
  const trig = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncIfDirty').length > 0;
  const last = PROP.getProperty('LAST_SYNC') || '–';
  const dirty = PROP.getProperty('DIRTY');
  SpreadsheetApp.getUi().alert(
    `Sinkron otomatis: ${trig ? 'AKTIF' : 'mati'}\n` +
    `Sinkron terakhir: ${last}\n` +
    `Perubahan menunggu: ${dirty ? 'ya' : 'tidak'}\n` +
    `Akun Sync: ${PROP.getProperty('SYNC_EMAIL') || '(belum diatur)'}`);
}

/* ========================= TRIGGERS ========================= */
function onSheetChange_(e) {
  // Perubahan pada sheet log tidak perlu memicu sinkron.
  try {
    if (e && e.range && e.range.getSheet().getName() === CFG.SHEET_LOG) return;
  } catch (err) { /* onChange tidak punya range */ }
  PROP.setProperty('DIRTY', String(Date.now()));
}

function syncIfDirty() {
  const dirty = Number(PROP.getProperty('DIRTY') || 0);
  if (!dirty) return;
  if (Date.now() - dirty < CFG.DEBOUNCE_SEC * 1000) return; // masih ada yang mengedit
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    PROP.deleteProperty('DIRTY');
    runSync_('otomatis');
  } catch (e) {
    PROP.setProperty('DIRTY', String(Date.now() - CFG.DEBOUNCE_SEC * 1000)); // coba lagi menit berikutnya
    logOnce_('GAGAL', e.message);
  } finally { lock.releaseLock(); }
}

function syncNow() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = runSync_('manual', true);
    ui.alert(`Terkirim: ${r.clients} pelanggan, ${r.months} bulan berdata.\nRevenue SAP ${fmt_(r.revSap)} · OPS ${fmt_(r.revOps)}.`);
  } catch (e) { logOnce_('GAGAL', e.message); ui.alert('Sinkron gagal: ' + e.message); }
}

function previewDataset() {
  const d = buildFromSpreadsheet_();
  const rs = d.clients.reduce((s, c) => s + c.rs, 0), ro = d.clients.reduce((s, c) => s + c.ro, 0);
  SpreadsheetApp.getUi().alert(
    `Pelanggan: ${d.clients.length}\nBulan berdata: ${d.dataMonths.map(i => MONTHS_ID[i]).join(', ')} ${d.year}\n` +
    `Revenue SAP ${fmt_(rs)} · OPS ${fmt_(ro)}\n` +
    `Label tertukar yang dikoreksi: ${d.swap.length ? d.swap.map(p => p.join(' ↔ ')).join('; ') : 'tidak ada'}\n` +
    `Catatan anomali: ${d.notes.length}\nUkuran: ${Math.round(JSON.stringify(d).length / 1024)} KB`);
}

/* ========================= SYNC CORE ========================= */
function runSync_(mode, force) {
  const data = buildFromSpreadsheet_();
  const json = JSON.stringify(data);
  if (json.length > 900000) throw new Error('Dataset melebihi batas 900 KB dokumen Firestore.');
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json));
  // Manual ("Sinkronkan sekarang") selalu mengirim; otomatis hanya bila isi berubah.
  if (!force && hash === PROP.getProperty('LAST_HASH')) return { skipped: true };

  const tok = getIdToken_();
  const base = `projects/${CFG.PROJECT_ID}/databases/(default)/documents`;
  const rs = data.clients.reduce((s, c) => s + c.rs, 0), ro = data.clients.reduce((s, c) => s + c.ro, 0);
  const body = {
    writes: [
      {
        update: {
          name: `${base}/datasets/${CFG.DATASET_ID}`,
          fields: {
            json: { stringValue: json },
            source: { stringValue: 'Google Sheets · ' + SpreadsheetApp.getActive().getName() },
            clients: { integerValue: String(data.clients.length) },
            hash: { stringValue: hash },
            updatedBy: { stringValue: tok.email },
          },
        },
        updateTransforms: [{ fieldPath: 'updatedAt', setToServerValue: 'REQUEST_TIME' }],
      },
      {
        update: {
          name: `${base}/audit/${Utilities.getUuid()}`,
          fields: {
            uid: { stringValue: tok.uid },
            email: { stringValue: tok.email },
            action: { stringValue: 'dataset_update' },
            detail: { stringValue: `Google Sheets (${mode}) · ${data.clients.length} pelanggan · data s.d. ${MONTHS_ID[data.lastMonth]} ${data.year}` },
            ua: { stringValue: 'Apps Script' },
          },
        },
        updateTransforms: [{ fieldPath: 'at', setToServerValue: 'REQUEST_TIME' }],
        currentDocument: { exists: false },
      },
    ],
  };
  const res = UrlFetchApp.fetch(`https://firestore.googleapis.com/v1/${base}:commit`, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + tok.idToken }, payload: JSON.stringify(body),
  });
  if (res.getResponseCode() !== 200) {
    const msg = safeErr_(res);
    if (/PERMISSION_DENIED/.test(msg)) throw new Error('Ditolak firestore.rules. Pastikan akun Sync punya dokumen users/{uid} dengan role "Sync" dan active: true. Detail: ' + msg);
    throw new Error('Firestore ' + res.getResponseCode() + ': ' + msg);
  }
  PROP.setProperty('LAST_HASH', hash);
  const stamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
  PROP.setProperty('LAST_SYNC', stamp);
  appendLog_([new Date(), 'OK', mode, data.clients.length, MONTHS_ID[data.dataMonths[0]] + '–' + MONTHS_ID[data.lastMonth], rs, ro, data.swap.length ? 'Label dikoreksi: ' + data.swap.map(p => p.join('↔')).join('; ') : '']);
  return { skipped: false, clients: data.clients.length, months: data.dataMonths.length, revSap: rs, revOps: ro };
}

function getIdToken_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('idToken');
  if (hit) return JSON.parse(hit);
  const key = PROP.getProperty('FIREBASE_API_KEY'), email = PROP.getProperty('SYNC_EMAIL'), pass = PROP.getProperty('SYNC_PASSWORD');
  if (!key || !email || !pass) throw new Error('Koneksi belum diatur. Jalankan menu "1. Atur koneksi Firebase".');
  const res = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=' + encodeURIComponent(key), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ email, password: pass, returnSecureToken: true }),
  });
  if (res.getResponseCode() !== 200) {
    const m = safeErr_(res);
    const map = { INVALID_LOGIN_CREDENTIALS: 'email atau password akun Sync salah', INVALID_PASSWORD: 'password akun Sync salah', EMAIL_NOT_FOUND: 'akun Sync tidak ditemukan', USER_DISABLED: 'akun Sync dinonaktifkan', 'API key not valid': 'Web API Key tidak valid' };
    const k = Object.keys(map).find(x => m.indexOf(x) >= 0);
    throw new Error('Login Firebase gagal: ' + (k ? map[k] : m));
  }
  const j = JSON.parse(res.getContentText());
  const tok = { idToken: j.idToken, uid: j.localId, email: j.email };
  cache.put('idToken', JSON.stringify(tok), Math.min(3000, Number(j.expiresIn || 3600) - 300));
  return tok;
}

function safeErr_(res) {
  try { const j = JSON.parse(res.getContentText()); return (j.error && (j.error.message || j.error.status)) || res.getContentText(); }
  catch (e) { return res.getContentText().slice(0, 300); }
}

/* ========================= LOG ========================= */
function appendLog_(row) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_LOG);
    sh.appendRow(['Waktu', 'Status', 'Mode', 'Pelanggan', 'Bulan', 'Revenue SAP', 'Revenue OPS', 'Catatan']);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm:ss');
    sh.getRange('F:G').setNumberFormat('#,##0');
  }
  sh.appendRow(row);
  if (sh.getLastRow() > 501) sh.deleteRows(2, sh.getLastRow() - 501);
}
function logOnce_(status, msg) {
  if (PROP.getProperty('LAST_ERR') === msg) return;
  PROP.setProperty('LAST_ERR', msg);
  try { appendLog_([new Date(), status, '', '', '', '', '', msg]); } catch (e) { /* abaikan */ }
}
function fmt_(v) { return 'Rp ' + Math.round(v).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }

/* ========================= BUILDER ========================= */
function buildFromSpreadsheet_() {
  const ss = SpreadsheetApp.getActive();
  const get = n => {
    const sh = ss.getSheetByName(n);
    if (!sh) throw new Error(`Sheet "${n}" tidak ditemukan.`);
    return sh.getDataRange().getValues();
  };
  PROP.deleteProperty('LAST_ERR');
  return buildDataset_({ summary: get(CFG.SHEET_SUMMARY), pareto: get(CFG.SHEET_PARETO), ledger: get(CFG.SHEET_LEDGER) });
}

/**
 * Dataset v2: semua angka dihitung dari "Rekap SAP S3" (ledger bulanan SAP & OPS, Jan–Des).
 * Summary & Breakdown Pareto dipakai sebagai pembanding. Fungsi ini identik dengan
 * buildProfitDataset() di dashboard, sehingga upload Excel dan sinkron Sheets menghasilkan data yang sama.
 */
function buildDataset_(S) { const d = buildProfitDataset(S); d.source = 'gsheet'; return d; }

/* ==== BUILDER DATASET v2 (dipakai bersama oleh dashboard & Apps Script) ====
 * Sumber angka utama: sheet "Rekap SAP S3" (ledger bulanan per IO, SAP & OPS).
 * Summary & Breakdown Pareto hanya dipakai sebagai pembanding (rekonsiliasi).
 * Output per pelanggan: array 12 bulan (Jan–Des) untuk revenue, direct cost per kategori,
 * dan depresiasi, masing-masing untuk SAP dan OPS. Dashboard menjumlahkan bulan sesuai filter periode.
 */
function buildProfitDataset(S) {
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  var CAT_KEYS = ['jasnaker', 'jarkom', 'gedung', 'kpk', 'lain', 'other'];
  var RAWCAT = { jasnaker: 'BEBAN JASNAKER', jarkom: 'BEBAN JARKOM', gedung: 'BEBAN MANDATORY GEDUNG', kpk: 'BEBAN KERJASAMA PIHAK KETIGA', lain: 'BEBAN LAIN - LAIN', other: 'BEBAN MARKETING' };
  var R = Math.round;
  var str = function (v) { return (v === null || v === undefined) ? '' : String(v).trim(); };
  var num = function (v) {
    if (v === null || v === undefined || v === '' || v === '-') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    if (v instanceof Date) return 0;
    var s = String(v).trim();
    if (s === '-' || s === '') return 0;
    s = s.replace(/^Rp\s*/i, '').replace(/\s/g, '');
    if (/^\(.*\)$/.test(s)) s = '-' + s.slice(1, -1);
    if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
    else if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
    var n = Number(s);
    return isFinite(n) ? n : 0;
  };
  var isTextNum = function (v) { return typeof v === 'string' && /\d/.test(v) && v.trim() !== '-'; };
  var z12 = function () { return [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; };
  var catKey = function (p) {
    var t = str(p).toUpperCase();
    if (t.indexOf('JASNAKER') >= 0) return 'jasnaker';
    if (t.indexOf('JARKOM') >= 0) return 'jarkom';
    if (t.indexOf('GEDUNG') >= 0) return 'gedung';
    if (t.indexOf('PIHAK KETIGA') >= 0) return 'kpk';
    if (t.indexOf('LAIN') >= 0) return 'lain';
    if (t.indexOf('MARKETING') >= 0) return 'other';
    return null;
  };

  /* ---------- 1. Ledger: cari baris header (tanggal bulan + "Grand Total") ---------- */
  var raw = S.ledger || [];
  var hr = -1;
  for (var r0 = 0; r0 < Math.min(raw.length, 40); r0++) {
    var row0 = raw[r0] || [];
    var hasDate = row0.some(function (h) { return h instanceof Date; });
    var hasGt = row0.some(function (h) { return str(h).toLowerCase() === 'grand total'; });
    if (hasDate && hasGt) { hr = r0; break; }
  }
  if (hr < 0) throw new Error('Sheet "Rekap SAP S3": baris header dengan kolom bulan (format tanggal) dan "Grand Total" tidak ditemukan.');
  var head = raw[hr];
  var col = function (name) { for (var i = 0; i < head.length; i++) if (str(head[i]).toLowerCase() === name) return i; return -1; };
  var cK = col('klien'), cL = col('layanan'), cS = col('source'), cP = col('pl'), cPar = col('pareto'), cGT = col('grand total');
  if (cK < 0 || cS < 0 || cP < 0) throw new Error('Sheet "Rekap SAP S3": kolom Klien / Source / PL tidak ditemukan di header.');
  var monthCols = [], year = null, skippedCols = 0;
  head.forEach(function (h, i) {
    if (!(h instanceof Date)) return;
    var d = new Date(h.getTime() + 12 * 3600e3); // hindari geser bulan karena zona waktu
    if (year === null) year = d.getFullYear();
    if (d.getFullYear() !== year) { skippedCols++; return; }
    monthCols.push([i, d.getMonth()]);
  });
  var dataMonths = monthCols.map(function (x) { return x[1]; }).sort(function (a, b) { return a - b; });
  var lastMonth = dataMonths[dataMonths.length - 1];
  var noteCol = cGT >= 0 ? cGT + 1 : -1;

  var C = {};                         // per pelanggan
  var blocks = {};                    // per (pelanggan|layanan|source): detail vs total
  var notes = [], negops = [], textRows = {}, totalMismatch = [];
  var get = function (n) {
    if (!C[n]) C[n] = { n: n, svcRev: {}, svcAny: {},
      rs: z12(), ro: z12(), dep: z12(), depo: z12(),
      c: { jasnaker: z12(), jarkom: z12(), gedung: z12(), kpk: z12(), lain: z12(), other: z12() },
      co: { jasnaker: z12(), jarkom: z12(), gedung: z12(), kpk: z12(), lain: z12(), other: z12() } };
    return C[n];
  };
  var blk = function (k) { return blocks[k] || (blocks[k] = { revT: null, revD: null, depT: null, depD: null, dirT: null, catSeen: false }); };
  var addArr = function (a, m, sign) { for (var i = 0; i < 12; i++) a[i] += sign * m[i]; };

  for (var r = hr + 1; r < raw.length; r++) {
    var row = raw[r] || [];
    var cl = str(row[cK]); if (!cl || cl.toLowerCase() === 'grand total' || cl.toLowerCase() === 'klien') continue;
    var lay = cL >= 0 ? str(row[cL]) : '', src = str(row[cS]).toUpperCase(), pl = str(row[cP]).toUpperCase(), par = cPar >= 0 ? str(row[cPar]) : '';
    if (src !== 'SAP' && src !== 'OPS') continue;
    var m = z12(), any = false;
    monthCols.forEach(function (x) { var v = num(row[x[0]]); m[x[1]] += v; if (v) any = true; if (isTextNum(row[x[0]])) textRows[cl] = true; });
    if (noteCol >= 0) { var nt = str(row[noteCol]); if (nt && !/^-?[\d.,]+$/.test(nt)) notes.push([cl, lay, src, par, nt]); }
    var c = get(cl); if (lay) c.svcAny[lay] = 1;
    var b = blk(cl + '|' + lay + '|' + src);
    var isTot = pl.indexOf('TOTAL') >= 0;
    if (pl.indexOf('PENDAPATAN') >= 0) {
      if (isTot) { b.revT = b.revT || z12(); addArr(b.revT, m, 1); } else { b.revD = b.revD || z12(); addArr(b.revD, m, 1); }
    } else if (pl.indexOf('DEPRESIASI') >= 0 || /DEPRESIASI/i.test(par)) {
      if (isTot) { b.depT = b.depT || z12(); addArr(b.depT, m, 1); } else { b.depD = b.depD || z12(); addArr(b.depD, m, 1); }
    } else if (isTot && pl.indexOf('DIRECT') >= 0) {
      b.dirT = b.dirT || z12(); addArr(b.dirT, m, 1);
    } else {
      var k = catKey(par);
      if (k && (pl === '' || pl.indexOf('DIRECT') >= 0)) { addArr(src === 'SAP' ? c.c[k] : c.co[k], m, 1); b.catSeen = true; b.catSum = b.catSum || z12(); addArr(b.catSum, m, 1); }
    }
  }
  /* revenue & depresiasi per blok: pakai baris Total, kalau tidak ada pakai baris detail */
  Object.keys(blocks).forEach(function (key) {
    var p = key.split('|'), cl = p[0], lay = p[1], src = p[2], b = blocks[key], c = C[cl];
    var rv = b.revT || b.revD;
    if (rv) {
      var tot = rv.reduce(function (s, v) { return s + v; }, 0);
      var sign;
      if (src === 'SAP') sign = -1;                       // SAP: pendapatan dicatat negatif (konvensi buku besar)
      else { sign = tot < 0 ? -1 : 1; if (tot < 0) negops.push([cl, lay, R(tot)]); } // OPS: seharusnya positif
      addArr(src === 'SAP' ? c.rs : c.ro, rv, sign);
      if (src === 'SAP') c.svcRev[lay] = (c.svcRev[lay] || 0) - tot;
    }
    var dp = b.depT || b.depD;
    if (dp) addArr(src === 'SAP' ? c.dep : c.depo, dp, 1);
    if (b.dirT && !b.catSeen) addArr(src === 'SAP' ? c.c.other : c.co.other, b.dirT, 1); // total tanpa rincian kategori
    if (b.dirT && b.catSeen) {
      var dT = b.dirT.reduce(function (s, v) { return s + v; }, 0), dC = b.catSum.reduce(function (s, v) { return s + v; }, 0);
      if (Math.abs(dT - dC) > 1000) totalMismatch.push([cl, lay, src, R(dT), R(dC)]);
    }
  });

  /* ---------- 2. Breakdown Pareto (pembanding) ---------- */
  var P = S.pareto || [], pc = {}, tot = {};
  var blocksP = [], titleRow = -1;
  for (var rp = 0; rp < Math.min(P.length, 10) && !blocksP.length; rp++) {
    (P[rp] || []).forEach(function (v, i) { if (typeof v === 'string' && /^beban /i.test(v.trim())) { var kk = catKey(v.replace(/lain lain/i, 'LAIN')); if (kk) blocksP.push([i, kk]); } });
    if (blocksP.length) titleRow = rp;
  }
  blocksP.forEach(function (bk) {
    for (var rr = titleRow + 2; rr < P.length; rr++) {
      var nn = str((P[rr] || [])[bk[0]]);
      if (nn.toLowerCase() === 'grand total') break;
      if (!nn || /^(row labels|\(blank\)|revenue|depresiasi)$/i.test(nn) || /^beban/i.test(nn)) continue;
      pc[nn] = pc[nn] || {}; pc[nn][bk[1]] = [R(num(P[rr][bk[0] + 2])), R(num(P[rr][bk[0] + 1]))];
    }
  });
  P.forEach(function (row) { var k = str((row || [])[0]); if (/^BEBAN /.test(k) || k === 'DEPRESIASI' || k === 'REVENUE') tot[k] = [R(num(row[2])), R(num(row[1]))]; });
  ['BEBAN JARKOM', 'BEBAN JASNAKER', 'BEBAN KERJASAMA PIHAK KETIGA', 'BEBAN LAIN - LAIN', 'BEBAN MANDATORY GEDUNG', 'BEBAN MARKETING', 'DEPRESIASI', 'REVENUE'].forEach(function (k) { if (!tot[k]) tot[k] = [0, 0]; });

  /* ---------- 3. Summary (pembanding) + deteksi label tertukar ---------- */
  var SM = S.summary || [], sumRows = [], start = -1;
  for (var rs = 0; rs < SM.length; rs++) if (str((SM[rs] || [])[0]).toLowerCase() === 'layanan') { start = rs + 2; break; }
  if (start > 0) for (var rs2 = start; rs2 < SM.length; rs2++) {
    var nm = str((SM[rs2] || [])[0]);
    if (nm.toLowerCase() === 'grand total') break;
    if (nm) sumRows.push({ lbl: nm, n: nm, rs: num(SM[rs2][1]), ro: num(SM[rs2][2]), ds: num(SM[rs2][4]), do: num(SM[rs2][5]) });
  }
  var sumArr = function (a) { return a.reduce(function (s, v) { return s + v; }, 0); };
  var ledRev = {}; Object.keys(C).forEach(function (n) { ledRev[n] = sumArr(C[n].rs); });
  var ledRevTo = function (n, last) { var c = C[n]; if (!c) return undefined; var s = 0; for (var i = 0; i <= last; i++) s += c.rs[i]; return s; };
  var parDir = {}; Object.keys(pc).forEach(function (n) { parDir[n] = Object.keys(pc[n]).reduce(function (s, k) { return s + pc[n][k][0]; }, 0); });
  var close = function (a, b2) { return a !== undefined && b2 !== undefined && Math.abs(a - b2) <= Math.max(2, Math.abs(b2) * 0.01); };
  var revMatch = function (row, name) { for (var mm = 0; mm < 12; mm++) if (close(ledRevTo(name, mm), row.rs)) return true; return false; };
  var score = function (row, name) { return (revMatch(row, name) ? 1 : 0) + (close(parDir[name], row.ds) ? 1 : 0); };
  var cand = [];
  for (var i1 = 0; i1 < sumRows.length; i1++) for (var j1 = i1 + 1; j1 < sumRows.length; j1++) {
    var A = sumRows[i1], B = sumRows[j1], ab = score(A, B.lbl), ba = score(B, A.lbl);
    var gain = ab + ba - score(A, A.lbl) - score(B, B.lbl);
    if (ab >= 1 && ba >= 1 && gain >= 2) cand.push([gain, i1, j1]);
  }
  cand.sort(function (x, y) { return y[0] - x[0]; });
  var used = {}, swap = [];
  cand.forEach(function (cc) { if (used[cc[1]] || used[cc[2]]) return; used[cc[1]] = used[cc[2]] = 1; var A = sumRows[cc[1]], B = sumRows[cc[2]]; A.n = B.lbl; B.n = A.lbl; swap.push([A.lbl, B.lbl]); });
  var sumBy = {}; sumRows.forEach(function (s) { sumBy[s.n] = s; });
  /* periode yang dicakup Summary (mis. Jan–Jul) per pelanggan */
  var sumPeriod = function (n) { var s = sumBy[n]; if (!s || !C[n]) return null; for (var mm = 0; mm < 12; mm++) if (close(ledRevTo(n, mm), s.rs) && Math.abs(s.rs) > 0) return mm; return -1; };

  /* ---------- 4. Susun pelanggan ---------- */
  var cls = function (s) {
    var t = s.toLowerCase();
    if (/^1500/.test(t) || t.indexOf('nomor akses') >= 0) return 'Nomor Akses / SIP Trunk';
    if (t.indexOf('waba') >= 0 || /^ota/.test(t)) return 'WhatsApp Business (WABA)';
    if (t.indexOf('omni') >= 0) return 'Omnichannel';
    if (t.indexOf('helpdesk') >= 0 || t.indexOf('programer') >= 0) return 'Helpdesk & IT Services';
    if (t.indexOf('crm') >= 0) return 'CRM Platform';
    return 'Contact Center Managed Service';
  };
  var order = sumRows.map(function (s) { return s.n; }).filter(function (n) { return C[n]; });
  Object.keys(C).sort().forEach(function (n) { if (order.indexOf(n) < 0) order.push(n); });
  var r12 = function (a) { return a.map(R); };
  var clients = order.map(function (n) {
    var c = C[n];
    var keys = Object.keys(c.svcAny).sort();
    var prim = '', best = -Infinity;
    Object.keys(c.svcRev).forEach(function (k) { if (c.svcRev[k] > best) { best = c.svcRev[k]; prim = k; } });
    if (!prim && keys.length) prim = keys[0];
    var m = { rs: r12(c.rs), ro: r12(c.ro), dep: r12(c.dep), depo: r12(c.depo), c: {}, co: {} };
    CAT_KEYS.forEach(function (k) { m.c[k] = r12(c.c[k]); m.co[k] = r12(c.co[k]); });
    var ds = z12(), dso = z12();
    CAT_KEYS.forEach(function (k) { for (var i = 0; i < 12; i++) { ds[i] += m.c[k][i]; dso[i] += m.co[k][i]; } });
    var cat = {}; ['jarkom', 'jasnaker', 'kpk', 'lain', 'gedung'].forEach(function (k) { cat[k] = [sumArr(m.c[k]), sumArr(m.co[k])]; });
    var s = sumBy[n];
    return {
      n: n, sl: prim ? cls(prim) : 'Lainnya', svc: keys, m: m,
      /* total seluruh bulan berdata (kompatibilitas modul lama) */
      rs: sumArr(m.rs), ro: sumArr(m.ro), ds: sumArr(ds), do: sumArr(dso), cat: cat, dep: [sumArr(m.dep), sumArr(m.depo)],
      mr: m.rs.slice(), md: ds.map(R),
      sum: s ? { lbl: s.lbl, rs: R(s.rs), ro: R(s.ro), ds: R(s.ds), do: R(s.do), upto: sumPeriod(n) } : null,
    };
  });
  /* pelanggan yang hanya ada di Summary */
  var onlySummary = sumRows.filter(function (s) { return !C[s.n]; }).map(function (s) { return s.n; });
  var mcat = {}; CAT_KEYS.forEach(function (k) { var o = {}; clients.forEach(function (c) { o[c.n] = c.m.c[k].slice(); }); mcat[RAWCAT[k]] = o; });
  var sapRev = 0, sapDir = 0; clients.forEach(function (c) { sapRev += c.rs; sapDir += c.ds; });

  return {
    v: 2, basis: 'ledger', year: year || new Date().getFullYear(),
    months: MON.slice(), dataMonths: dataMonths, lastMonth: lastMonth,
    clients: clients, tot: tot, notes: notes, negops: negops, mcat: mcat, swap: swap,
    raw: { sapRev: R(sapRev), sapDir: R(sapDir) },
    textRows: Object.keys(textRows), onlySummary: onlySummary, totalMismatch: totalMismatch, skippedCols: skippedCols,
    generatedAt: new Date().toISOString(),
  };
}

/* Untuk pengujian di luar Apps Script (Node.js). */
if (typeof module !== 'undefined') module.exports = { buildDataset_ };
