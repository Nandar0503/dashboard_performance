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
    ui.alert(`Terkirim: ${r.clients} pelanggan, ${r.months} bulan.\nRevenue SAP ${fmt_(r.revSap)} · OPS ${fmt_(r.revOps)}.`);
  } catch (e) { logOnce_('GAGAL', e.message); ui.alert('Sinkron gagal: ' + e.message); }
}

function previewDataset() {
  const d = buildFromSpreadsheet_();
  const rs = d.clients.reduce((s, c) => s + c.rs, 0), ro = d.clients.reduce((s, c) => s + c.ro, 0);
  SpreadsheetApp.getUi().alert(
    `Pelanggan: ${d.clients.length}\nBulan: ${d.months.join(', ')} ${d.year}\n` +
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
            detail: { stringValue: `Google Sheets (${mode}) · ${data.clients.length} pelanggan · ${data.months.length} bulan` },
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
  appendLog_([new Date(), 'OK', mode, data.clients.length, data.months.join('–'), rs, ro, data.swap.length ? 'Label dikoreksi: ' + data.swap.map(p => p.join('↔')).join('; ') : '']);
  return { skipped: false, clients: data.clients.length, months: data.months.length, revSap: rs, revOps: ro };
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
 * Fungsi murni: menerima nilai 2D dari tiga sheet, mengembalikan objek dataset dashboard.
 * (Diuji terpisah terhadap file Excel asli.)
 */
function buildDataset_(S) {
  const num = v => {
    if (v === null || v === undefined || v === '' || v === '-') return 0;
    if (typeof v === 'number') return v;
    if (v instanceof Date) return 0;
    let s = String(v).trim();
    if (s === '-' || s === '') return 0;
    if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
    else if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
    const n = Number(s);
    return isFinite(n) ? n : 0;
  };
  const str = v => (v === null || v === undefined) ? '' : String(v).trim();
  const R = Math.round;
  const textNum = (v) => typeof v === 'string' && /\d/.test(v) && v.trim() !== '-';

  /* ---- Ledger: Rekap SAP S3 ---- */
  const L = S.ledger;
  const head = L[0];
  const monthCols = [], monthLabels = [];
  let year = null, gtCol = -1;
  head.forEach((h, i) => {
    if (h instanceof Date) { monthCols.push(i); monthLabels.push(MONTHS_ID[h.getMonth()]); year = year || h.getFullYear(); }
    else if (str(h).toLowerCase() === 'grand total') gtCol = i;
  });
  if (!monthCols.length || gtCol < 0) throw new Error('Header bulan / "Grand Total" tidak ditemukan di baris pertama sheet ' + CFG.SHEET_LEDGER + '.');
  const NM = monthCols.length;
  const zeros = () => new Array(NM).fill(0);
  const svc = {}, mrev = {}, mdir = {}, mcat = {}, dep = {}, notes = [], negops = [], textRows = {};
  for (let r = 1; r < L.length; r++) {
    const row = L[r];
    const c = str(row[1]); if (!c) continue;
    const layanan = str(row[4]), src = str(row[5]), pl = str(row[6]), par = str(row[7]);
    const note = str(row[gtCol + 1]);
    if (note) notes.push([c, layanan, src, par, note]);
    const m = monthCols.map(i => num(row[i]));
    const gt = num(row[gtCol]);
    if (monthCols.some(i => textNum(row[i]))) textRows[c] = true;
    if (pl === '4.DEPRESIASI Total') { dep[c] = dep[c] || [0, 0]; dep[c][src === 'SAP' ? 0 : 1] += gt; }
    if (src === 'OPS' && pl === '1.PENDAPATAN Total' && gt < 0) negops.push([c, layanan, R(gt)]);
    if (src !== 'SAP') continue;
    if (pl === '1.PENDAPATAN Total') {
      svc[c] = svc[c] || {}; svc[c][layanan] = (svc[c][layanan] || 0) - gt;
      mrev[c] = mrev[c] || zeros(); m.forEach((v, i) => mrev[c][i] -= v);
    } else if (pl === '2.DIRECT Total') {
      mdir[c] = mdir[c] || zeros(); m.forEach((v, i) => mdir[c][i] += v);
    } else if (pl === '2.DIRECT' && par) {
      mcat[par] = mcat[par] || {}; mcat[par][c] = mcat[par][c] || zeros(); m.forEach((v, i) => mcat[par][c][i] += v);
    }
  }

  /* ---- Breakdown Pareto ---- */
  const P = S.pareto;
  const CATKEY = t => {
    t = t.toLowerCase();
    if (t.indexOf('jarkom') >= 0) return 'jarkom';
    if (t.indexOf('jasnaker') >= 0) return 'jasnaker';
    if (t.indexOf('pihak ketiga') >= 0) return 'kpk';
    if (t.indexOf('lain') >= 0) return 'lain';
    if (t.indexOf('gedung') >= 0) return 'gedung';
    return null;
  };
  let blocks = [], titleRow = -1;
  for (let r = 0; r < Math.min(P.length, 10) && !blocks.length; r++) {
    P[r].forEach((v, i) => { const k = typeof v === 'string' && /^beban /i.test(v.trim()) ? CATKEY(v) : null; if (k) blocks.push([i, k]); });
    if (blocks.length) titleRow = r;
  }
  if (!blocks.length) throw new Error('Judul blok "Beban ..." tidak ditemukan di sheet ' + CFG.SHEET_PARETO + '.');
  const pc = {};
  const skip = { 'grand total': 1, 'row labels': 1, '(blank)': 1, 'revenue': 1, 'depresiasi': 1 };
  blocks.forEach(([col, key]) => {
    for (let r = titleRow + 2; r < P.length; r++) {
      const n = str(P[r][col]);
      if (n.toLowerCase() === 'grand total') break;
      if (!n || skip[n.toLowerCase()] || /^beban/i.test(n)) continue;
      pc[n] = pc[n] || {};
      pc[n][key] = [R(num(P[r][col + 2])), R(num(P[r][col + 1]))]; // [SAP, OPS]
    }
  });
  const tot = {};
  P.forEach(row => {
    const k = str(row[0]);
    if (/^BEBAN /.test(k) || k === 'DEPRESIASI' || k === 'REVENUE') tot[k] = [R(num(row[2])), R(num(row[1]))];
  });
  ['BEBAN JARKOM', 'BEBAN JASNAKER', 'BEBAN KERJASAMA PIHAK KETIGA', 'BEBAN LAIN - LAIN', 'BEBAN MANDATORY GEDUNG', 'BEBAN MARKETING', 'DEPRESIASI', 'REVENUE']
    .forEach(k => { if (!tot[k]) tot[k] = [0, 0]; });

  /* ---- Summary ---- */
  const SM = S.summary;
  let start = -1;
  for (let r = 0; r < SM.length; r++) if (str(SM[r][0]).toLowerCase() === 'layanan') { start = r + 2; break; }
  if (start < 0) throw new Error('Tabel "Layanan" tidak ditemukan di sheet ' + CFG.SHEET_SUMMARY + '.');
  const rows = [];
  for (let r = start; r < SM.length; r++) {
    const n = str(SM[r][0]);
    if (n.toLowerCase() === 'grand total') break;
    if (!n) continue;
    rows.push({ lbl: n, n, rs: num(SM[r][1]), ro: num(SM[r][2]), ds: num(SM[r][4]), do: num(SM[r][5]) });
  }
  if (!rows.length) throw new Error('Tidak ada baris pelanggan di sheet ' + CFG.SHEET_SUMMARY + '.');

  /* ---- Deteksi label tertukar: bandingkan angka baris dengan ledger & Pareto ---- */
  const rawRev = {}; Object.keys(svc).forEach(c => rawRev[c] = Object.keys(svc[c]).reduce((s, k) => s + svc[c][k], 0));
  const parDir = {}; Object.keys(pc).forEach(c => parDir[c] = Object.keys(pc[c]).reduce((s, k) => s + pc[c][k][0], 0));
  const close = (a, b, absTol) => a !== undefined && b !== undefined && Math.abs(a - b) <= Math.max(absTol, Math.abs(b) * 0.01);
  const score = (row, name) => (close(rawRev[name], row.rs, 2) ? 1 : 0) + (close(parDir[name], row.ds, 2) ? 1 : 0);
  const cand = [];
  for (let i = 0; i < rows.length; i++) for (let j = i + 1; j < rows.length; j++) {
    const a = rows[i], b = rows[j];
    const ab = score(a, b.lbl), ba = score(b, a.lbl);
    const gain = ab + ba - score(a, a.lbl) - score(b, b.lbl);
    if (ab >= 1 && ba >= 1 && gain >= 2) cand.push([gain, i, j]);
  }
  cand.sort((x, y) => y[0] - x[0]);
  const used = {}, swap = [];
  cand.forEach(([g, i, j]) => {
    if (used[i] || used[j]) return;
    used[i] = used[j] = 1;
    const a = rows[i], b = rows[j];
    a.n = b.lbl; b.n = a.lbl;
    swap.push([a.lbl, b.lbl]);
  });

  /* ---- Gabungkan per pelanggan ---- */
  const cls = s => {
    const t = s.toLowerCase();
    if (/^1500/.test(t) || t.indexOf('nomor akses') >= 0) return 'Nomor Akses / SIP Trunk';
    if (t.indexOf('waba') >= 0 || /^ota/.test(t)) return 'WhatsApp Business (WABA)';
    if (t.indexOf('omni') >= 0) return 'Omnichannel';
    if (t.indexOf('helpdesk') >= 0 || t.indexOf('programer') >= 0) return 'Helpdesk & IT Services';
    if (t.indexOf('crm') >= 0) return 'CRM Platform';
    return 'Contact Center Managed Service';
  };
  const clients = rows.map(r => {
    const s = svc[r.n] || {};
    const keys = Object.keys(s).sort();
    let prim = '', best = -Infinity;
    keys.forEach(k => { if (s[k] > best) { best = s[k]; prim = k; } });
    return {
      n: r.n, lbl: r.lbl, rs: R(r.rs), ro: R(r.ro), ds: R(r.ds), do: R(r.do),
      sl: prim ? cls(prim) : 'Lainnya', svc: keys, cat: pc[r.n] || {},
      dep: (dep[r.n] || [0, 0]).map(R),
      mr: (mrev[r.n] || zeros()).map(R), md: (mdir[r.n] || zeros()).map(R),
    };
  });
  const mc = {};
  Object.keys(mcat).forEach(k => { mc[k] = {}; Object.keys(mcat[k]).forEach(c => mc[k][c] = mcat[k][c].map(R)); });
  const sumObj = o => Object.keys(o).reduce((s, c) => s + o[c].reduce((a, b) => a + b, 0), 0);

  return {
    clients, tot, notes, negops, mcat: mc, swap,
    raw: { sapRev: R(sumObj(mrev)), sapDir: R(sumObj(mdir)) },
    months: monthLabels, year: year || new Date().getFullYear(),
    textRows: Object.keys(textRows),
    source: 'gsheet', generatedAt: new Date().toISOString(),
  };
}

/* Untuk pengujian di luar Apps Script (Node.js). */
if (typeof module !== 'undefined') module.exports = { buildDataset_ };
