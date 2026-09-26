const APP = Object.freeze({
  NAME: 'Gold Planner',
  TZ: 'Asia/Jakarta',
  DB_KEY: 'GOLD_PLANNER_DB_ID',
  CACHE_KEY: 'GOLD_PLANNER_CACHE_V4',
  // Isi setelah folder PWA selesai di-hosting, contoh: https://gold-planner.netlify.app
  PWA_URL: '',
  API_VERSION: 1,
  API_PROXY_KEY: '1c909e9cb96b5008b9c3c08e3e5f7852167048fb6949b23d02cb651efaa86dde',
  API_MAX_BODY_CHARS: 750000,
  CACHE_SECONDS: 900,
  SESSION_SECONDS: 21600,
  VAULT_CHUNK_SIZE: 45000,
  VAULT_CHUNKS: 14,
  MAX_VAULT_CHARS: 600000,
  SOURCES: {
    ANTAM: 'https://www.logammulia.com/id/harga-emas-hari-ini',
    ANTAM_BUYBACK: 'https://www.logammulia.com/id/sell/gold',
    ANTAM_PROXY: 'https://logam-mulia-api.iamutaki.workers.dev/api/prices/logammulia',
    UBS: 'https://ubslifestyle.com/harga-buyback-hari-ini/',
    GALERI24: 'https://galeri24.co.id/harga-emas',
    WORLD_LIVE: 'https://api.gold-api.com/price/XAU',
    WORLD_FX: 'https://api.frankfurter.dev/v2/rate/USD/IDR',
    WORLD_HISTORY: 'https://stooq.com/q/d/l/?s=xauusd&i=d',
    WORLD_FALLBACK_LIVE: 'https://api.goldprice.dev/v1/prices?symbol=XAU-USD-SPOT',
    WORLD_FALLBACK_BARS: 'https://api.goldprice.dev/v1/bars?symbol=XAU-USD-SPOT&interval=1d&limit=40',
    WORLD_YAHOO: 'https://query1.finance.yahoo.com/v8/finance/chart/GC%3DF?range=1mo&interval=1d'
  }
});

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Endpoint untuk PWA V2. Dipanggil server Netlify, bukan langsung dari browser. */
function doPost(e) {
  try {
    const raw = e && e.postData ? String(e.postData.contents || '') : '';
    if (!raw || raw.length > APP.API_MAX_BODY_CHARS) throw new Error('Permintaan API kosong atau terlalu besar.');
    const request = JSON.parse(raw);
    if (String(request.proxyKey || '') !== APP.API_PROXY_KEY) throw new Error('Akses API ditolak.');
    if (Number(request.version) !== APP.API_VERSION) throw new Error('Versi API tidak didukung. Perbarui aplikasi.');
    const action = String(request.action || '');
    const args = Array.isArray(request.args) ? request.args : [];
    if (args.length > 3) throw new Error('Argumen API tidak valid.');

    let data;
    switch (action) {
      case 'getAuthSystemStatus': data = getAuthSystemStatus(); break;
      case 'getLoginChallenge': data = getLoginChallenge(args[0]); break;
      case 'loginAccount': data = loginAccount(args[0]); break;
      case 'resumeAccount': data = resumeAccount(args[0]); break;
      case 'registerAccount': data = registerAccount(args[0]); break;
      case 'saveEncryptedVault': data = saveEncryptedVault(args[0], args[1]); break;
      case 'logoutAccount': data = logoutAccount(args[0]); break;
      case 'getDashboardData': data = getDashboardData(Boolean(args[0])); break;
      default: throw new Error('Fungsi API tidak dikenali.');
    }
    return apiJson_({ ok: true, data: data });
  } catch (error) {
    return apiJson_({ ok: false, error: error && error.message ? String(error.message) : 'Terjadi kesalahan pada server.' });
  }
}

function apiJson_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Jalankan sekali dari editor Apps Script sebelum deploy. */
function setupApp() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(APP.DB_KEY);
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { id = ''; }
  }
  if (!id) {
    ss = SpreadsheetApp.create('Gold Planner - Database');
    props.setProperty(APP.DB_KEY, ss.getId());
  }

  ensureSheet_(ss, 'Purchases', [
    'ID', 'Tanggal', 'Merek', 'Gram', 'Jumlah', 'Harga Total', 'Tempat Beli', 'Catatan', 'Dibuat'
  ]);
  ensureSheet_(ss, 'PriceHistory', [
    'Timestamp', 'Merek', 'Gram', 'Harga Jual', 'Harga Buyback', 'Sumber', 'Status'
  ]);
  const settings = ensureSheet_(ss, 'Settings', ['Key', 'Value', 'Keterangan']);
  ensureSheet_(ss, 'AuthUsers', [
    'User ID', 'Auth Salt', 'Auth Verifier', 'Vault Salt', 'Dibuat', 'Login Terakhir', 'Status'
  ]);
  const vaultHeaders = ['User ID', 'IV', 'Revision', 'Diperbarui'];
  for (let v = 1; v <= APP.VAULT_CHUNKS; v++) vaultHeaders.push('Cipher ' + v);
  ensureSheet_(ss, 'EncryptedVaults', vaultHeaders);
  ensureSheet_(ss, 'DeviceSessions', ['Token Hash', 'User ID', 'Dibuat', 'Terakhir Digunakan']);
  ensureSettingsDefaults_(settings);
  settings.autoResizeColumns(1, 3);
  createRefreshTrigger_();
  refreshAllData();
  return { ok: true, spreadsheetUrl: ss.getUrl() };
}

function getDashboardData(forceRefresh) {
  const db = getDb_();
  ensureSettingsDefaults_(db.getSheetByName('Settings'));
  let snapshot = readCache_();
  const stale = !snapshot || !snapshot.fetchedAt || Date.now() - new Date(snapshot.fetchedAt).getTime() > APP.CACHE_SECONDS * 1000;
  if (forceRefresh || stale) {
    try { snapshot = refreshAllData(); } catch (e) {
      if (!snapshot) throw new Error('Belum ada data harga. Jalankan setupApp(), lalu coba lagi. Detail: ' + e.message);
      snapshot.refreshError = e.message;
    }
  }
  return {
    appName: APP.NAME,
    snapshot: snapshot,
    rules: getRules_(),
    sources: APP.SOURCES,
    pwaUrl: getPwaUrl_(),
    authRequired: true
  };
}

function getPwaUrl_() {
  const value = String(APP.PWA_URL || '').trim().replace(/\/$/, '');
  return /^https:\/\/[a-z0-9.-]+(?::\d+)?(?:\/[^\s]*)?$/i.test(value) ? value : '';
}

/**
 * Sistem akun zero-knowledge praktis:
 * - Browser mengirim userId yang sudah di-hash, verifier, dan ciphertext.
 * - Password serta isi pembelian tidak pernah dikirim ke server.
 * - Server hanya memeriksa verifier dan menyimpan brankas AES-GCM terenkripsi.
 */
function getLoginChallenge(userId) {
  validateUserId_(userId);
  const sheet = getAuthSheets_().users;
  const row = findRowById_(sheet, userId);
  if (!row) throw new Error('Akun tidak ditemukan. Periksa username atau buat akun baru.');
  const values = sheet.getRange(row, 1, 1, 7).getValues()[0];
  if (String(values[6] || 'active') !== 'active') throw new Error('Akun sedang dinonaktifkan.');
  return { authSalt: String(values[1]), vaultSalt: String(values[3]) };
}

function getAuthSystemStatus() {
  const auth = getAuthSheets_();
  return {
    ready: Boolean(auth.users && auth.vaults),
    version: 2,
    encryption: 'AES-GCM-256',
    passwordRecovery: false
  };
}

function registerAccount(payload) {
  validateRegistration_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const auth = getAuthSheets_();
    const users = auth.users;
    const vaults = auth.vaults;
    const existingUser = findRowById_(users, payload.userId);
    if (existingUser) {
      if (!findRowById_(vaults, payload.userId)) users.deleteRow(existingUser);
      else throw new Error('Email atau username sudah digunakan. Silakan masuk atau gunakan identitas lain.');
    }
    users.appendRow([
      payload.userId, payload.authSalt, payload.authVerifier, payload.vaultSalt,
      new Date(), new Date(), 'active'
    ]);
    try {
      writeVaultRow_(vaults, payload.userId, payload.iv, payload.ciphertext, 1, 0);
    } catch (e) {
      const incompleteRow = findRowById_(users, payload.userId);
      if (incompleteRow && !findRowById_(vaults, payload.userId)) users.deleteRow(incompleteRow);
      throw e;
    }
    return { sessionToken: createSession_(payload.userId), revision: 1 };
  } finally {
    lock.releaseLock();
  }
}

function loginAccount(payload) {
  if (!payload) throw new Error('Permintaan login tidak valid.');
  validateUserId_(payload.userId);
  validateBase64_(payload.authVerifier, 'Verifier login', 32, 256);
  checkLoginRate_(payload.userId);
  const auth = getAuthSheets_();
  const users = auth.users;
  const row = findRowById_(users, payload.userId);
  if (!row) throw new Error('Username atau password salah.');
  const values = users.getRange(row, 1, 1, 7).getValues()[0];
  if (String(values[6] || 'active') !== 'active') throw new Error('Akun sedang dinonaktifkan.');
  if (!constantTimeEqual_(String(values[2]), String(payload.authVerifier))) {
    recordLoginFailure_(payload.userId);
    throw new Error('Username atau password salah.');
  }
  CacheService.getScriptCache().remove(loginRateKey_(payload.userId));
  users.getRange(row, 6).setValue(new Date());
  const vault = readVault_(auth.vaults, payload.userId);
  delete vault.row;
  return {
    sessionToken: createSession_(payload.userId),
    vault: vault,
    vaultSalt: String(values[3])
  };
}

function resumeAccount(sessionToken) {
  const userId = requireSession_(sessionToken);
  const auth = getAuthSheets_();
  const vault = readVault_(auth.vaults, userId);
  delete vault.row;
  return { sessionToken: sessionToken, vault: vault };
}

function saveEncryptedVault(sessionToken, payload) {
  const userId = requireSession_(sessionToken);
  validateVaultPayload_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getAuthSheets_().vaults;
    const saved = readVault_(sheet, userId);
    const expected = Number(payload.revision || 0);
    if (Number(saved.revision || 0) !== expected) {
      throw new Error('Data berubah di perangkat lain. Login ulang sebelum menyimpan agar data tidak tertimpa.');
    }
    const nextRevision = expected + 1;
    writeVaultRow_(sheet, userId, payload.iv, payload.ciphertext, nextRevision, saved.row);
    return { ok: true, revision: nextRevision, updatedAt: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}

function logoutAccount(sessionToken) {
  if (sessionToken) CacheService.getScriptCache().remove(sessionKey_(sessionToken));
  if (sessionToken) {
    const sheet = getAuthSheets_().sessions;
    const row = findRowById_(sheet, sha256Hex_(String(sessionToken)));
    if (row) sheet.deleteRow(row);
  }
  return { ok: true };
}

function validateRegistration_(p) {
  if (!p) throw new Error('Data pendaftaran tidak valid.');
  validateUserId_(p.userId);
  validateBase64_(p.authSalt, 'Auth salt', 16, 128);
  validateBase64_(p.vaultSalt, 'Vault salt', 16, 128);
  validateBase64_(p.authVerifier, 'Auth verifier', 32, 256);
  validateVaultPayload_(p);
}

function validateVaultPayload_(p) {
  if (!p) throw new Error('Brankas terenkripsi tidak valid.');
  validateBase64_(p.iv, 'IV', 12, 64);
  validateBase64_(p.ciphertext, 'Ciphertext', 16, APP.MAX_VAULT_CHARS);
}

function validateUserId_(value) {
  if (!/^[a-f0-9]{64}$/.test(String(value || ''))) throw new Error('Identitas akun tidak valid.');
}

function validateBase64_(value, label, min, max) {
  const text = String(value || '');
  if (text.length < min || text.length > max || !/^[A-Za-z0-9+/]+={0,2}$/.test(text)) {
    throw new Error(label + ' tidak valid.');
  }
}

function findRowById_(sheet, id) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const match = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(id)).matchEntireCell(true).findNext();
  return match ? match.getRow() : 0;
}

function writeVaultRow_(sheet, userId, iv, ciphertext, revision, existingRow) {
  const chunks = [];
  const text = String(ciphertext);
  for (let i = 0; i < APP.VAULT_CHUNKS; i++) {
    chunks.push(text.slice(i * APP.VAULT_CHUNK_SIZE, (i + 1) * APP.VAULT_CHUNK_SIZE));
  }
  if (chunks.join('').length !== text.length) throw new Error('Brankas terlalu besar untuk disimpan.');
  const values = [userId, iv, revision, new Date()].concat(chunks);
  const row = existingRow || findRowById_(sheet, userId);
  if (row) sheet.getRange(row, 1, 1, values.length).setValues([values]);
  else sheet.appendRow(values);
}

function readVault_(sheet, userId) {
  const row = findRowById_(sheet, userId);
  if (!row) throw new Error('Brankas akun tidak ditemukan.');
  const values = sheet.getRange(row, 1, 1, 4 + APP.VAULT_CHUNKS).getValues()[0];
  return {
    iv: String(values[1] || ''), revision: Number(values[2] || 0),
    ciphertext: values.slice(4).join(''), row: row
  };
}

function createSession_(userId) {
  const token = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + '|' + Utilities.getUuid() + '|' + Date.now())
  ).replace(/=+$/g, '');
  CacheService.getScriptCache().put(sessionKey_(token), JSON.stringify({ userId: userId }), APP.SESSION_SECONDS);
  const sessions = getAuthSheets_().sessions;
  const tokenHash = sha256Hex_(token);
  sessions.appendRow([tokenHash, userId, new Date(), new Date()]);
  return token;
}

function requireSession_(token) {
  const raw = token ? CacheService.getScriptCache().get(sessionKey_(token)) : '';
  if (raw) {
    try { return JSON.parse(raw).userId; }
    catch (e) { throw new Error('Sesi tidak valid. Silakan login kembali.'); }
  }
  if (!token || String(token).length < 32 || String(token).length > 200) throw new Error('Sesi berakhir. Silakan login kembali.');
  const sessions = getAuthSheets_().sessions;
  const row = findRowById_(sessions, sha256Hex_(String(token)));
  if (!row) throw new Error('Sesi berakhir. Silakan login kembali.');
  const userId = String(sessions.getRange(row, 2).getValue() || '');
  validateUserId_(userId);
  CacheService.getScriptCache().put(sessionKey_(token), JSON.stringify({ userId: userId }), APP.SESSION_SECONDS);
  return userId;
}

function sessionKey_(token) {
  return 'GP_SESSION_' + sha256Hex_(String(token || ''));
}

function sha256Hex_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function (b) { const v = b < 0 ? b + 256 : b; return ('0' + v.toString(16)).slice(-2); }).join('');
}

function constantTimeEqual_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function loginRateKey_(userId) { return 'GP_LOGIN_FAIL_' + String(userId); }

function checkLoginRate_(userId) {
  const attempts = Number(CacheService.getScriptCache().get(loginRateKey_(userId)) || 0);
  if (attempts >= 8) throw new Error('Terlalu banyak percobaan login. Tunggu sekitar 15 menit lalu coba lagi.');
}

function recordLoginFailure_(userId) {
  const cache = CacheService.getScriptCache();
  const key = loginRateKey_(userId);
  cache.put(key, String(Number(cache.get(key) || 0) + 1), 900);
}

function getAuthSheets_() {
  const db = getDb_();
  const users = ensureSheet_(db, 'AuthUsers', [
    'User ID', 'Auth Salt', 'Auth Verifier', 'Vault Salt', 'Dibuat', 'Login Terakhir', 'Status'
  ]);
  const headers = ['User ID', 'IV', 'Revision', 'Diperbarui'];
  for (let i = 1; i <= APP.VAULT_CHUNKS; i++) headers.push('Cipher ' + i);
  const vaults = ensureSheet_(db, 'EncryptedVaults', headers);
  const sessions = ensureSheet_(db, 'DeviceSessions', ['Token Hash', 'User ID', 'Dibuat', 'Terakhir Digunakan']);
  return { users: users, vaults: vaults, sessions: sessions };
}

function saveSavingsTarget(form) {
  throw new Error('Fungsi lama dinonaktifkan. Target kini disimpan dalam brankas terenkripsi pengguna.');
}

function refreshAllData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const previous = readCache_() || { brands: [], world: null };
    const brands = ['ANTAM', 'UBS', 'GALERI24'].map(function (brand) {
      try { return fetchBrand_(brand); }
      catch (e) {
        const old = (previous.brands || []).find(function (x) { return x.brand === brand; });
        if (old) return Object.assign({}, old, { status: 'cached', error: e.message });
        return { brand: brand, rows: [], status: 'error', error: e.message, source: sourceForBrand_(brand) };
      }
    });
    let world;
    try { world = fetchWorldGold_(previous.world); }
    catch (e) {
      world = previous.world ? Object.assign({}, previous.world, { status: 'cached', error: e.message }) : { status: 'error', error: e.message, history: [] };
    }
    const snapshot = { fetchedAt: new Date().toISOString(), brands: brands, world: world };
    CacheService.getScriptCache().put(APP.CACHE_KEY, JSON.stringify(snapshot), 21600);
    PropertiesService.getScriptProperties().setProperty(APP.CACHE_KEY, JSON.stringify(snapshot));
    appendPriceHistory_(brands);
    return snapshot;
  } finally {
    lock.releaseLock();
  }
}

function addPurchase(form) {
  throw new Error('Fungsi lama dinonaktifkan. Pembelian kini disimpan dalam brankas terenkripsi pengguna.');
}

function deletePurchase(id) {
  throw new Error('Fungsi lama dinonaktifkan. Pembelian kini disimpan dalam brankas terenkripsi pengguna.');
}

function fetchBrand_(brand) {
  if (brand === 'ANTAM') {
    let rows = [], buybackPerGram = 0, status = 'official', recordedAt = '', errors = [];
    try {
      rows = parseAntamPriceTable_(fetchHtml_(APP.SOURCES.ANTAM));
    } catch (e) {
      errors.push('halaman harga: ' + e.message);
      try {
        const proxy = fetchJson_(APP.SOURCES.ANTAM_PROXY);
        rows = parseAntamProxy_(proxy);
        recordedAt = String((((proxy || {}).data || [])[0] || {}).recordedDate || '');
        status = 'official-proxy';
      } catch (proxyError) {
        errors.push('cache Logam Mulia: ' + proxyError.message);
        rows = antamOfficialSeedRows_();
        recordedAt = '2026-09-03';
        status = 'official-cache';
      }
    }
    try {
      buybackPerGram = extractBuybackPerGram_(fetchHtml_(APP.SOURCES.ANTAM_BUYBACK));
    } catch (e) { errors.push('halaman buyback: ' + e.message); }
    if (!(buybackPerGram > 0)) buybackPerGram = getLastOfficialAntamBuyback_();
    if (!(buybackPerGram > 0)) {
      buybackPerGram = 2492000;
      status = status === 'official' ? 'official-cache' : status;
    }
    rows.forEach(function (r) { r.buyback = Math.round(buybackPerGram * r.grams); });
    const antam = buildBrand_('ANTAM', rows, APP.SOURCES.ANTAM, status);
    antam.buybackSource = APP.SOURCES.ANTAM_BUYBACK;
    antam.sourceLabel = 'Logam Mulia';
    antam.recordedAt = recordedAt;
    antam.error = errors.join(' | ');
    return antam;
  }
  if (brand === 'UBS') {
    try {
      return buildBrand_('UBS', parseGenericTable_(fetchHtml_(APP.SOURCES.UBS)), APP.SOURCES.UBS, 'official');
    } catch (e) {
      return buildBrand_('UBS', parseGaleriSection_(fetchHtml_(APP.SOURCES.GALERI24), 'UBS'), APP.SOURCES.GALERI24, 'official-fallback');
    }
  }
  return buildBrand_('GALERI24', parseGaleriSection_(fetchHtml_(APP.SOURCES.GALERI24), 'GALERI 24'), APP.SOURCES.GALERI24, 'official');
}

function buildBrand_(brand, rows, source, status) {
  rows = dedupeRows_(rows).filter(function (r) { return r.grams > 0 && (r.sell > 0 || r.buyback > 0); });
  if (!rows.length) throw new Error('Format harga ' + brand + ' pada situs sumber tidak dikenali.');
  rows.sort(function (a, b) { return a.grams - b.grams; });
  return { brand: brand, rows: rows, source: source, status: status, updatedAt: new Date().toISOString() };
}

function parseAntamPriceTable_(html) {
  const tables = String(html).match(/<table\b[^>]*>[\s\S]*?<\/table>/gi) || [];
  let best = [];
  tables.forEach(function (table) {
    if (!/Harga\s*Dasar/i.test(textOnly_(table))) return;
    const rows = [];
    (table.match(/<tr\b[^>]*>[\s\S]*?<\/tr>/gi) || []).forEach(function (tr) {
      const cells = (tr.match(/<t[dh]\b[^>]*>[\s\S]*?<\/t[dh]>/gi) || []).map(textOnly_).filter(Boolean);
      if (cells.length < 2) return;
      const grams = parseWeight_(cells[0]);
      const prices = cells.slice(1).map(parsePriceNumber_).filter(function (n) { return n > 10000; });
      if (!grams || !prices.length) return;
      const baseSell = prices[0];
      const sell = prices[1] || Math.round(baseSell * 1.0025);
      rows.push({ grams: grams, baseSell: baseSell, sell: sell, buyback: 0, purchaseTaxRate: 0.0025 });
    });
    if (rows.length > best.length) best = rows;
  });
  if (best.length < 3) throw new Error('Tabel Emas Batangan pada Logam Mulia tidak dikenali.');
  return best;
}

function parseAntamProxy_(json) {
  const data = ((json || {}).data || []).filter(function (item) {
    return item.source === 'logammulia' && item.material === 'gold' && item.materialType === 'Emas Batangan';
  });
  const rows = data.map(function (item) {
    return { grams: Number(item.weight), baseSell: Number(item.sellPrice), sell: Number(item.sellPrice), buyback: 0, purchaseTaxRate: 0 };
  }).filter(function (row) { return row.grams > 0 && row.sell > 10000; });
  if (rows.length < 3) throw new Error('Data harga Emas Batangan tidak lengkap.');
  return rows;
}

function antamOfficialSeedRows_() {
  return [
    [0.5, 1372924], [1, 2645598], [2, 5231045], [3, 7821505], [5, 13002425], [10, 25949713],
    [25, 64748468], [50, 129417738], [100, 258757280], [250, 646627538], [500, 1293044550], [1000, 2586049000]
  ].map(function (row) { return { grams: row[0], sell: row[1], buyback: 0, purchaseTaxRate: 0.0025 }; });
}

function getLastOfficialAntamBuyback_() {
  try {
    const sheet = getDb_().getSheetByName('PriceHistory');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;
    const startRow = Math.max(2, lastRow - 999);
    const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 7).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      const grams = Number(rows[i][2]), buyback = Number(rows[i][4]), source = String(rows[i][5] || '');
      if (String(rows[i][1]) === 'ANTAM' && grams > 0 && buyback > 0 && /logammulia\.com/i.test(source)) return buyback / grams;
    }
  } catch (e) {}
  return 0;
}

function parseGenericTable_(html) {
  const rows = [];
  const matches = html.match(/<tr\b[^>]*>[\s\S]*?<\/tr>/gi) || [];
  matches.forEach(function (tr) {
    const cells = (tr.match(/<t[dh]\b[^>]*>[\s\S]*?<\/t[dh]>/gi) || []).map(textOnly_).filter(Boolean);
    if (cells.length < 2) return;
    const grams = parseWeight_(cells[0]);
    const money = cells.slice(1).map(parseRupiah_).filter(function (n) { return n > 10000; });
    if (!grams || !money.length) return;
    rows.push({ grams: grams, sell: money[0] || 0, buyback: money.length > 1 ? money[money.length - 1] : 0 });
  });
  if (rows.length) return rows;

  const lines = htmlToLines_(html);
  for (let i = 0; i < lines.length; i++) {
    const grams = parseWeight_(lines[i]);
    if (!grams) continue;
    const money = [];
    for (let j = i + 1; j < Math.min(i + 6, lines.length); j++) {
      const n = parseRupiah_(lines[j]);
      if (n > 10000) money.push(n);
      if (money.length === 2) break;
    }
    if (money.length) rows.push({ grams: grams, sell: money[0], buyback: money[1] || 0 });
  }
  return rows;
}

function parseGaleriSection_(html, vendor) {
  const lines = htmlToLines_(html);
  const target = ('HARGA ' + vendor).replace(/\s+/g, ' ').toUpperCase();
  let start = lines.findIndex(function (x) { return x.toUpperCase().replace(/\s+/g, ' ') === target; });
  if (start < 0) start = lines.findIndex(function (x) { return x.toUpperCase().indexOf(target) >= 0; });
  if (start < 0) throw new Error('Bagian harga ' + vendor + ' tidak ditemukan.');
  const rows = [];
  for (let i = start + 1; i < lines.length; i++) {
    if (i > start + 5 && /^HARGA\s+[A-Z]/i.test(lines[i]) && lines[i].toUpperCase().indexOf(target) < 0) break;
    const grams = parseWeight_(lines[i]);
    if (!grams) continue;
    const money = [];
    for (let j = i + 1; j < Math.min(i + 7, lines.length); j++) {
      const value = parseRupiah_(lines[j]);
      if (value >= 0 && /^RP/i.test(lines[j].replace(/\s/g, ''))) money.push(value);
      if (money.length === 2) break;
    }
    if (money.length === 2) rows.push({ grams: grams, sell: money[0], buyback: money[1] });
  }
  return rows;
}

function fetchWorldGold_(previousWorld) {
  const today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const start = Utilities.formatDate(new Date(Date.now() - 35 * 86400000), 'UTC', 'yyyy-MM-dd');
  const errors = [];
  let current = 0;
  let timestamp = new Date().toISOString();
  let liveSource = '';
  let bars = [];
  let historySource = '';
  let usdIdr = 0;

  try {
    const live = fetchJson_(APP.SOURCES.WORLD_LIVE);
    current = Number(live.price || 0);
    timestamp = normalizeApiDate_(live.updatedAt || live.updated_at || live.timestamp) || timestamp;
    liveSource = 'Gold-API XAU Spot';
  } catch (e) { errors.push('live utama: ' + e.message); }

  try {
    bars = fetchStooqHistory_(start, today);
    historySource = 'Stooq XAU/USD';
  } catch (e) { errors.push('riwayat utama: ' + e.message); }

  if (!current || bars.length < 2) {
    try {
      const fallback = fetchGoldpriceDev_(start, today);
      if (!current) {
        current = fallback.current;
        timestamp = fallback.timestamp || timestamp;
        liveSource = 'GoldPrice.dev XAU Spot';
      }
      if (bars.length < 2) {
        bars = fallback.bars;
        historySource = 'GoldPrice.dev XAU/USD';
      }
    } catch (e) { errors.push('cadangan XAU: ' + e.message); }
  }

  if (!current || bars.length < 2) {
    try {
      const yahoo = fetchYahooGold_();
      if (!current) {
        current = yahoo.current;
        timestamp = yahoo.timestamp || timestamp;
        liveSource = 'Yahoo Finance Gold Futures';
      }
      if (bars.length < 2) {
        bars = yahoo.bars;
        historySource = 'Yahoo Finance Gold Futures';
      }
    } catch (e) { errors.push('cadangan futures: ' + e.message); }
  }

  try {
    const fx = fetchJson_(APP.SOURCES.WORLD_FX);
    usdIdr = Number(fx.rate || 0);
  } catch (e) { errors.push('kurs USD/IDR: ' + e.message); }

  if (bars.length < 2 && previousWorld && previousWorld.history && previousWorld.history.length > 1) {
    bars = previousWorld.history;
    historySource = (previousWorld.historySource || previousWorld.source || 'snapshot terakhir') + ' (tersimpan)';
  }
  if (!current && previousWorld) {
    current = Number(previousWorld.usdOz || 0);
    timestamp = previousWorld.timestamp || timestamp;
    liveSource = (previousWorld.liveSource || previousWorld.source || 'snapshot terakhir') + ' (tersimpan)';
  }
  if (!usdIdr && previousWorld && previousWorld.usdOz > 0 && previousWorld.idrGram > 0) {
    usdIdr = Number(previousWorld.idrGram) * 31.1034768 / Number(previousWorld.usdOz);
  }
  if (!current && bars.length) current = Number(bars[bars.length - 1].close || 0);
  if (!current) throw new Error('Semua sumber harga emas dunia sedang tidak dapat diakses. ' + errors.join(' | '));

  bars = bars.filter(function (b) { return b.date && Number(b.close) > 0; })
    .sort(function (a, b) { return new Date(a.date) - new Date(b.date); }).slice(-30);
  const currentDate = String(timestamp).slice(0, 10);
  if (bars.length && String(bars[bars.length - 1].date).slice(0, 10) === currentDate) {
    bars[bars.length - 1].close = current;
  } else {
    bars.push({ date: currentDate, close: current });
  }
  const closes = bars.map(function (b) { return Number(b.close); });
  const analysis = analyzeSeries_(closes, current);
  return {
    status: liveSource.indexOf('(tersimpan)') >= 0 ? 'cached' : (errors.length ? 'partial' : 'live'),
    usdOz: current,
    usdIdr: usdIdr,
    idrGram: usdIdr ? current * usdIdr / 31.1034768 : Number(previousWorld && previousWorld.idrGram || 0),
    timestamp: timestamp,
    history: bars,
    analysis: analysis,
    source: liveSource,
    liveSource: liveSource,
    historySource: historySource,
    error: errors.length ? errors.join(' | ') : ''
  };
}

function fetchStooqHistory_(start, end) {
  const url = APP.SOURCES.WORLD_HISTORY + '&d1=' + start.replace(/-/g, '') + '&d2=' + end.replace(/-/g, '');
  const text = fetchText_(url);
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 3 || !/^date,/i.test(lines[0])) throw new Error('CSV kosong atau format berubah');
  const bars = lines.slice(1).map(function (line) {
    const c = line.split(',');
    return { date: c[0], open: Number(c[1]), high: Number(c[2]), low: Number(c[3]), close: Number(c[4]) };
  }).filter(function (b) { return b.date && b.close > 0; });
  if (bars.length < 2) throw new Error('riwayat kurang dari dua hari');
  return bars;
}

function fetchGoldpriceDev_(start, end) {
  const live = fetchJson_(APP.SOURCES.WORLD_FALLBACK_LIVE);
  const tick = (live.symbols || [])[0] || {};
  const url = APP.SOURCES.WORLD_FALLBACK_BARS + '&from=' + encodeURIComponent(start) + '&to=' + encodeURIComponent(end);
  const history = fetchJson_(url);
  const bars = (history.bars || history.series || []).map(function (b) {
    return { date: b.bar_start || b.date, open: Number(b.open), high: Number(b.high), low: Number(b.low), close: Number(b.close) };
  }).filter(function (b) { return b.date && b.close > 0; });
  return { current: Number(tick.price || 0), timestamp: tick.computed_at || '', bars: bars };
}

function fetchYahooGold_() {
  const json = fetchJson_(APP.SOURCES.WORLD_YAHOO);
  const result = (((json || {}).chart || {}).result || [])[0] || {};
  const timestamps = result.timestamp || [];
  const quote = ((((result.indicators || {}).quote || [])[0]) || {});
  const closes = quote.close || [];
  const bars = timestamps.map(function (t, i) {
    return { date: new Date(Number(t) * 1000).toISOString(), open: Number((quote.open || [])[i]), high: Number((quote.high || [])[i]), low: Number((quote.low || [])[i]), close: Number(closes[i]) };
  }).filter(function (b) { return b.close > 0; });
  const meta = result.meta || {};
  return { current: Number(meta.regularMarketPrice || (bars.length && bars[bars.length - 1].close) || 0), timestamp: meta.regularMarketTime ? new Date(meta.regularMarketTime * 1000).toISOString() : '', bars: bars };
}

function fetchJson_(url) { return JSON.parse(fetchText_(url)); }
function fetchText_(url) {
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GoldPlanner/2.0)', 'Accept': 'application/json,text/csv,text/plain,*/*' } });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('HTTP ' + code);
  return response.getContentText();
}
function normalizeApiDate_(value) {
  if (!value) return '';
  const n = Number(value);
  if (isFinite(n) && n > 0) return new Date(n < 1000000000000 ? n * 1000 : n).toISOString();
  const date = new Date(value);
  return isNaN(date.getTime()) ? '' : date.toISOString();
}

function analyzeSeries_(closes, current) {
  if (!closes.length) return { signal: 'Data belum cukup', score: 50, change1d: 0, change7d: 0, change30d: 0, ma7: current, ma30: current, high30: current, low30: current, rangePosition: 50, volatility: 0, trend: 'Belum tersedia' };
  const last7 = closes.slice(-7);
  const last30 = closes.slice(-30);
  const ma7 = avg_(last7);
  const ma30 = avg_(last30);
  const base7 = closes[Math.max(0, closes.length - 8)] || current;
  const base1 = closes[Math.max(0, closes.length - 2)] || current;
  const base30 = closes[0] || current;
  const change1d = base1 ? (current / base1 - 1) * 100 : 0;
  const change7d = base7 ? (current / base7 - 1) * 100 : 0;
  const change30d = base30 ? (current / base30 - 1) * 100 : 0;
  const high30 = Math.max.apply(null, last30);
  const low30 = Math.min.apply(null, last30);
  const rangePosition = high30 === low30 ? 50 : (current - low30) / (high30 - low30) * 100;
  const returns = [];
  for (let i = 1; i < last30.length; i++) returns.push((last30[i] / last30[i - 1] - 1) * 100);
  const volatility = stddev_(returns);
  let score = 50;
  if (current < ma7) score += 12; else score -= 8;
  if (ma7 < ma30) score += 12; else score -= 8;
  if (change7d < -2) score += 10;
  if (change7d > 3) score -= 12;
  score = Math.max(0, Math.min(100, Math.round(score)));
  let signal = 'Netral — pertimbangkan pembelian bertahap';
  if (score >= 70) signal = 'Harga melemah — menarik untuk dipantau';
  if (score <= 35) signal = 'Momentum kuat — hindari keputusan terburu-buru';
  let trend = 'Mendatar';
  if (ma7 > ma30 * 1.003) trend = 'Naik';
  if (ma7 < ma30 * 0.997) trend = 'Turun';
  return { signal: signal, score: score, change1d: change1d, change7d: change7d, change30d: change30d, ma7: ma7, ma30: ma30, high30: high30, low30: low30, rangePosition: Math.max(0, Math.min(100, rangePosition)), volatility: volatility, trend: trend };
}

function getRules_() {
  const settings = getSettings_();
  return [
    {
      brand: 'ANTAM', threshold: numSetting_(settings, 'ANTAM_TAX_THRESHOLD', 10000000),
      taxRate: numSetting_(settings, 'ANTAM_TAX_RATE', numSetting_(settings, 'ANTAM_TAX_NPWP', 0.015)),
      stampFee: numSetting_(settings, 'ANTAM_STAMP_FEE', 10000),
      taxLabel: 'PPh 22 buyback ANTAM', taxMode: 'above-threshold', source: APP.SOURCES.ANTAM_BUYBACK,
      notes: ['Harga buyback resmi dihitung per gram dan sama untuk semua pecahan serta tahun produksi.', 'PPh 22 sebesar 1,5% dipotong jika nilai transaksi di atas Rp10 juta; NIK wajib sesuai data identitas.', 'Materai Rp10.000 dikenakan pada transaksi buyback.', 'Pembayaran melalui transfer H+1 sampai H+3 hari kerja setelah transaksi di Butik Emas LM.']
    },
    {
      brand: 'UBS', threshold: numSetting_(settings, 'UBS_TAX_THRESHOLD', 0),
      taxRate: numSetting_(settings, 'UBS_TAX_RATE', 0), stampFee: numSetting_(settings, 'UBS_STAMP_FEE', 0),
      taxLabel: 'Pajak/potongan otomatis', taxMode: 'none', source: 'https://ubslifestyle.com/pedoman-buyback/',
      notes: ['Nilai logam mulia dihitung dari berat dikali harga buyback UBS saat ini.', 'Layanan UBS Lifestyle berlaku untuk produk yang dibeli melalui UBS Lifestyle dan memerlukan invoice asli.', 'Pajak atau potongan lain tidak dihitung otomatis; masukkan biaya aktual pada kolom potongan/biaya bila diberlakukan.']
    },
    {
      brand: 'GALERI24', threshold: numSetting_(settings, 'GALERI24_TAX_THRESHOLD', 10000000),
      taxRate: numSetting_(settings, 'GALERI24_TAX_RATE', numSetting_(settings, 'GALERI24_TAX_NPWP', 0.0025)),
      stampFee: numSetting_(settings, 'GALERI24_STAMP_FEE', 0),
      taxLabel: 'PPh 22 transaksi LJK bulion', taxMode: 'above-threshold', source: APP.SOURCES.GALERI24,
      notes: ['Harga final mengikuti sistem Galeri24 saat transaksi dan hasil pemeriksaan fisik.', 'Simulasi PPh 22 sebesar 0,25% di atas Rp10 juta digunakan untuk transaksi pembelian emas oleh LJK penyelenggara usaha bulion.', 'Kuitansi, kondisi produk, kanal transaksi, dan biaya outlet dapat memengaruhi nilai akhir.']
    }
  ];
}

function calculateSavingsGoal_(purchases, brands) {
  const settings = getSettings_();
  const targetGrams = Math.max(0, numSetting_(settings, 'SAVINGS_TARGET_GRAMS', 10));
  const targetDate = String(settings.SAVINGS_TARGET_DATE || '');
  const ownedGrams = purchases.filter(function (p) { return p.brand === 'ANTAM'; })
    .reduce(function (sum, p) { return sum + p.grams * p.qty; }, 0);
  const remainingGrams = Math.max(0, targetGrams - ownedGrams);
  const antam = brands.find(function (b) { return b.brand === 'ANTAM'; });
  const sellPerGram = estimateRate_(antam ? antam.rows : [], 1, 'sell');
  const estimatedRemainingCost = remainingGrams * sellPerGram;
  const progress = targetGrams > 0 ? Math.min(100, ownedGrams / targetGrams * 100) : 0;
  const monthsRemaining = targetDate ? monthsUntil_(targetDate) : 0;
  return {
    targetGrams: targetGrams, targetDate: targetDate, ownedGrams: ownedGrams,
    remainingGrams: remainingGrams, progress: progress, sellPerGram: sellPerGram,
    estimatedRemainingCost: estimatedRemainingCost, monthsRemaining: monthsRemaining,
    monthlyGrams: monthsRemaining > 0 ? remainingGrams / monthsRemaining : 0,
    monthlyAmount: monthsRemaining > 0 ? estimatedRemainingCost / monthsRemaining : 0,
    completed: targetGrams > 0 && ownedGrams >= targetGrams
  };
}

function calculatePortfolio_(purchases, brands) {
  let invested = 0, currentValue = 0, totalGrams = 0;
  const items = purchases.map(function (p) {
    const brand = brands.find(function (b) { return b.brand === p.brand; });
    const rate = estimateRate_(brand ? brand.rows : [], p.grams, 'buyback');
    const value = rate * p.grams * p.qty;
    invested += p.totalPrice;
    currentValue += value;
    totalGrams += p.grams * p.qty;
    return Object.assign({}, p, { currentValue: value, profit: value - p.totalPrice, rate: rate });
  });
  return { invested: invested, currentValue: currentValue, profit: currentValue - invested, totalGrams: totalGrams, items: items };
}

function getPurchases_() {
  const sheet = getDb_().getSheetByName('Purchases');
  const rows = sheet.getDataRange().getValues().slice(1);
  return rows.filter(function (r) { return r[0]; }).map(function (r) {
    return {
      id: String(r[0]), date: Utilities.formatDate(new Date(r[1]), APP.TZ, 'yyyy-MM-dd'),
      brand: String(r[2]), grams: Number(r[3]), qty: Number(r[4]), totalPrice: Number(r[5]),
      store: String(r[6] || ''), note: String(r[7] || '')
    };
  }).sort(function (a, b) { return b.date.localeCompare(a.date); });
}

function appendPriceHistory_(brands) {
  const sheet = getDb_().getSheetByName('PriceHistory');
  const now = new Date();
  const rows = [];
  brands.forEach(function (b) {
    (b.rows || []).forEach(function (r) { rows.push([now, b.brand, r.grams, r.sell, r.buyback, b.source, b.status]); });
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP.DB_KEY);
  if (!id) throw new Error('Database belum dibuat. Jalankan fungsi setupApp() satu kali dari editor Apps Script.');
  return SpreadsheetApp.openById(id);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#21163a').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureSettingsDefaults_(sheet) {
  const defaults = [
    ['ANTAM_TAX_THRESHOLD', 10000000, 'Ambang PPh 22 buyback ANTAM'],
    ['ANTAM_TAX_RATE', 0.015, 'PPh 22 buyback ANTAM sesuai halaman resmi'],
    ['ANTAM_STAMP_FEE', 10000, 'Biaya materai buyback ANTAM'],
    ['UBS_TAX_THRESHOLD', 0, 'Tidak menghitung pajak UBS secara otomatis'],
    ['UBS_TAX_RATE', 0, 'Isi bila kebijakan transaksi berubah'],
    ['UBS_STAMP_FEE', 0, 'Isi bila ada biaya tetap'],
    ['GALERI24_TAX_THRESHOLD', 10000000, 'Ambang transaksi LJK bulion'],
    ['GALERI24_TAX_RATE', 0.0025, 'PPh 22 pembelian emas oleh LJK bulion'],
    ['GALERI24_STAMP_FEE', 0, 'Isi bila ada biaya tetap'],
    ['SAVINGS_TARGET_GRAMS', 10, 'Target tabungan ANTAM dalam gram'],
    ['SAVINGS_TARGET_DATE', '', 'Tenggat target tabungan ANTAM']
  ];
  const existing = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) existing[String(r[0])] = true; });
  }
  const missing = defaults.filter(function (r) { return !existing[r[0]]; });
  if (missing.length) sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
}

function upsertSetting_(sheet, key, value, description) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === key) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[value, description]]);
      return;
    }
  }
  sheet.appendRow([key, value, description]);
}

function createRefreshTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'refreshAllData'; });
  if (!exists) ScriptApp.newTrigger('refreshAllData').timeBased().everyHours(1).create();
}

function getSettings_() {
  const rows = getDb_().getSheetByName('Settings').getDataRange().getValues().slice(1);
  const out = {};
  rows.forEach(function (r) { if (r[0]) out[String(r[0])] = r[1]; });
  return out;
}

function readCache_() {
  const raw = CacheService.getScriptCache().get(APP.CACHE_KEY) || PropertiesService.getScriptProperties().getProperty(APP.CACHE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function fetchHtml_(url) {
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GoldPlanner/1.0)', 'Accept-Language': 'id-ID,id;q=0.9' }
  });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error('Sumber merespons HTTP ' + response.getResponseCode());
  return response.getContentText();
}

function htmlToLines_(html) {
  return String(html).replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<br\s*\/?\s*>|<\/p>|<\/div>|<\/li>|<\/tr>|<\/td>|<\/th>|<\/h[1-6]>/gi, '\n')
    .replace(/<[^>]+>/g, ' ').split(/\n+/).map(decodeHtml_).map(function (s) { return s.replace(/\s+/g, ' ').trim(); }).filter(Boolean);
}

function textOnly_(html) { return decodeHtml_(String(html).replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim(); }
function decodeHtml_(s) { return String(s).replace(/&nbsp;|&#160;/gi, ' ').replace(/&amp;/gi, '&').replace(/&#39;|&apos;/gi, "'").replace(/&quot;/gi, '"').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>'); }
function parseWeight_(s) {
  const cleaned = String(s).replace(',', '.').trim();
  if (!/^(?:\d+(?:\.\d+)?)(?:\s*(?:gram|gr|g))?$/i.test(cleaned)) return 0;
  const n = Number((cleaned.match(/\d+(?:\.\d+)?/) || [0])[0]);
  return n > 0 && n <= 1000 ? n : 0;
}
function parseRupiah_(s) {
  if (!/rp/i.test(String(s))) return -1;
  const digits = String(s).replace(/[^0-9]/g, '');
  return digits ? Number(digits) : 0;
}
function parsePriceNumber_(s) {
  const text = String(s || '').trim();
  if (!/\d/.test(text)) return -1;
  const digits = text.replace(/[^0-9]/g, '');
  return digits ? Number(digits) : -1;
}
function extractBuybackPerGram_(html) {
  const lines = htmlToLines_(html);
  for (let i = 0; i < lines.length; i++) {
    if (!/buy\s*back|buyback|jual kembali/i.test(lines[i])) continue;
    for (let j = i; j < Math.min(i + 6, lines.length); j++) {
      const n = parseRupiah_(lines[j]);
      if (n > 100000) return n;
    }
  }
  return 0;
}
function dedupeRows_(rows) {
  const map = {};
  rows.forEach(function (r) {
    if (!map[r.grams] || (r.buyback > 0 && !map[r.grams].buyback)) map[r.grams] = r;
  });
  return Object.keys(map).map(function (k) { return map[k]; });
}
function estimateRate_(rows, grams, field) {
  if (!rows || !rows.length) return 0;
  const exact = rows.find(function (r) { return Math.abs(r.grams - grams) < 0.0001; });
  const chosen = exact || rows.reduce(function (best, r) { return Math.abs(r.grams - grams) < Math.abs(best.grams - grams) ? r : best; }, rows[0]);
  return Number(chosen[field] || 0) / chosen.grams;
}
function sourceForBrand_(brand) { return brand === 'ANTAM' ? APP.SOURCES.ANTAM : brand === 'UBS' ? APP.SOURCES.UBS : APP.SOURCES.GALERI24; }
function numSetting_(obj, key, fallback) { const n = Number(obj[key]); return isFinite(n) ? n : fallback; }
function monthsUntil_(dateText) {
  const end = new Date(String(dateText) + 'T23:59:59+07:00');
  if (isNaN(end.getTime()) || end.getTime() <= Date.now()) return 0;
  return Math.max(1, Math.ceil((end.getTime() - Date.now()) / (30.4375 * 86400000)));
}
function avg_(a) { return a.length ? a.reduce(function (x, y) { return x + y; }, 0) / a.length : 0; }
function stddev_(a) { if (!a.length) return 0; const m = avg_(a); return Math.sqrt(avg_(a.map(function (x) { return Math.pow(x - m, 2); }))); }
function sanitizeCell_(value) { const s = String(value || '').trim(); return /^[=+\-@]/.test(s) ? "'" + s : s; }
