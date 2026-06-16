// ============================================================
// KONFIGURASI KOLOM (0-based index, kolom A=0, B=1, C=2, ...)
// ============================================================
const CONFIG = {
  COL_CODE:          1,   // B: Code Barang MyBiz
  COL_NAME:         15,   // P: Nama Barang / Nama Departemen
  // Kolom STOK sisi kiri (C-J)
  COL_SUB_TOTAL:     9,   // J: Sub Total
  COL_RUSAK:        10,   // K: Rusak/SVC
  COL_SO:           11,   // L: SO
  COL_DP:           12,   // M: DP
  COL_KK:           13,   // N: KK
  COL_PJ:           14,   // O: PJ
  // Kolom sisi kanan (Q-AH)
  COL_ALL_STOK_FULL:16,   // Q: ALL STOK
  COL_OUT_START:    17,   // R: OUT YGY (pertama)
  COL_ALL_OUT:      24,   // Y: ALL OUT
  COL_OUT_QTY:      25,   // Z: Out Qty
  COL_ORDER_PLUS:   26,   // AA: ORDER (+) Qty
  COL_ORDER_MINUS:  27,   // AB: ORDER (-) Qty
  COL_IKIRA2:       28,   // AC: Ikira2 Habis
  COL_KET_KSG:      31,   // AF: KET KSG
  COL_PO:           32,   // AG: PO
  COL_KET_KEEP:     33,   // AH: KET KEEP
  TOTAL_COLS:       34,   // Baca tepat 34 kolom (A sampai AH)
  // Status bar ada di baris 1 (index 0), kolom C-M (index 2-12)
  ROW_STATUS:        0,   // Baris 1 (index 0) berisi status Data Stok OK dll
  COL_STATUS_START:  2,   // C: mulai status
  CABANG:    ['YGY','SMG','SLO','PWT','BBS','TGL','MDN'],
  DATA_START_ROW: 3       // Index 3 = baris ke-4 (0-based)
};

// ============================================================
// KONFIGURASI SHEET CEK VENDOR (0-based index)
// ============================================================
const VENDOR_CONFIG = {
  COL_UNIQ:         0,   // A: Kode unik per transaksi (NB-AC-A14-51M-31RN1, dst) — tidak dipakai matching
  COL_CODE:         1,   // B: Kode SKU base (NB-AC-A14-51M-31RN) — ini yang dicocokkan exact
  COL_DATE:         2,   // C: Tanggal
  COL_VENDOR:       3,   // D: Vendor
  COL_QTY:          4,   // E: SUM of Qty
  COL_HARGA_NORMAL: 5,   // F: Cek Harga Normal
  COL_HARGA_DISKON: 6,   // G: Cek Harga + Diskon
  COL_KET:          7,   // H: Keterangan
  TOTAL_COLS:       8,   // Baca 8 kolom (A-H)
  DATA_START_ROW:   2    // Index 2 = baris ke-3, data mulai sini
};

// ============================================================
// KONFIGURASI SHEET MASTER SKU (0-based index)
// Master SKU NB -> kolom: A=Kode, C=Type Laptop, J=Type Proc, L=Harga, M=Note
// ============================================================
const MASTER_CONFIG = {
  COL_CODE:         0,   // A: Kode Barang (dicocokkan ke Purchase kolom B)
  COL_TYPE_LAPTOP:  2,   // C: Type Laptop
  COL_RAM:          5,   // F: Ram
  COL_STORAGE:      6,   // G: Storage SSD
  COL_SIZE:         7,   // H: Size layar
  COL_TYPE_PROC:    9,   // J: Type Proc
  COL_PROCESSOR:   10,   // K: Processor (Type Proc 2)
  COL_HARGA:       11,   // L: Harga (untuk range)
  COL_NOTE:        12,   // M: Note / Catatan per SKU
  TOTAL_COLS:      13,   // Baca kolom A-M
  DATA_START_ROW:   1    // Index 1 = baris ke-2 (baris 1 = header)
};

// ============================================================
// KONFIGURASI SHEET MASTER SKU PC (struktur lebih simple dari NB)
// Master SKU PC -> kolom: A=Kode, I=Harga ELS
// ============================================================
const MASTER_PC_CONFIG = {
  COL_CODE:         0,   // A: Kode
  COL_HARGA:        8,   // I: Harga ELS
  TOTAL_COLS:       9,   // Baca kolom A-I
  DATA_START_ROW:   1    // Baris 1 = header, data mulai baris 2
};

// ============================================================
// SUMBER DATA EKSTERNAL — bypass IMPORTRANGE untuk semua field
// kecuali Processor (vlookup typeProc->processor di Master SKU NB).
// Spreadsheet sumber dibaca langsung via SpreadsheetApp.openById().
// Akun pemilik Apps Script harus punya akses ke ID ini.
// ============================================================
const EXTERNAL_SOURCES = {
  spreadsheetId: '12SUP1b7YfIEntEyj1T5YJwBS89uHtNk9lJTkFywkV_Q',
  NB: {
    productSheet: 'Laptop_Products',
    productCols: {
      code:       1,    // B: SKU
      typeLaptop: 3,    // D
      typeProc:   5,    // F
      ram:        6,    // G
      storage:    7,    // H
      size:       8,    // I
      harga:     10     // K
    },
    noteSheet: 'Struktur_Harga_NB',
    noteCols: {
      code: 1,          // B
      note: 18          // S
    }
  },
  PC: {
    productSheet: 'PC_Products',
    productCols: {
      code:  0,         // A
      harga: 8          // I
    },
    noteSheet: 'Struktur_Harga_PC',
    noteCols: {
      code: 1,          // B
      note: 9           // J
    }
  }
};

// ============================================================
// LOOKUP PROCESSOR — vlookup di Master SKU NB (active spreadsheet).
// Type Proc di kolom P (index 15) -> Processor di kolom Q (index 16).
// Hanya 2 kolom yang dibaca, jadi fast walaupun sheet active punya
// IMPORTRANGE di kolom lain.
// ============================================================
const PROCESSOR_LOOKUP = {
  sheetName:    'Master SKU NB',
  colTypeProc:  15,   // P (0-based)
  colProcessor: 16    // Q
};

// ============================================================
// WEB APP ENTRY
// 2 versi dalam 1 app:
//   ?mode=cache -> versi dengan in-memory cache (index_cache.html)
//   ?mode=fast  -> versi tanpa cache, lazy-load (index_fast.html)
//   default     -> fast
// ============================================================
function doGet(e) {
  var mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'fast';
  var file = (mode === 'cache') ? 'index_cache' : 'index_fast';
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle('Purchase Report Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// CACHE LAYER — bypass IMPORTRANGE re-evaluation
// CacheService: max 100KB/key & 9MB total → kita chunk jadi 90KB.
// Strategi:
//  - TTL pendek (90-600 detik) supaya data tidak terlalu stale
//  - Tombol "Update Data" pass forceFresh=true → bypass cache
//  - Cache disimpan jika respons sukses (skip kalau ada error)
// ============================================================
const CACHE_VER = 'v7';   // bump ini saat shape data berubah

function _cacheGet(key) {
  try {
    var cache = CacheService.getScriptCache();
    var meta = cache.get(key + '__meta');
    if (!meta) return null;
    var info = JSON.parse(meta);
    var keys = [];
    for (var i = 0; i < info.n; i++) keys.push(key + '__' + i);
    var got = cache.getAll(keys);
    var parts = [];
    for (var j = 0; j < info.n; j++) {
      var v = got[key + '__' + j];
      if (v == null) return null;   // ada chunk hilang → invalidate
      parts.push(v);
    }
    return JSON.parse(parts.join(''));
  } catch(e) {
    Logger.log('cacheGet ' + key + ' error: ' + e.toString());
    return null;
  }
}

function _cachePut(key, obj, ttlSec) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(obj);
    var size = 90000;
    var n = Math.ceil(json.length / size) || 1;
    var batch = {};
    for (var i = 0; i < n; i++) {
      batch[key + '__' + i] = json.substring(i * size, (i + 1) * size);
    }
    batch[key + '__meta'] = JSON.stringify({ n: n, t: Date.now() });
    cache.putAll(batch, ttlSec);
  } catch(e) {
    Logger.log('cachePut ' + key + ' error: ' + e.toString());
  }
}

function _cacheBust(key) {
  try {
    var cache = CacheService.getScriptCache();
    var meta = cache.get(key + '__meta');
    if (!meta) return;
    var info = JSON.parse(meta);
    var keys = [key + '__meta'];
    for (var i = 0; i < info.n; i++) keys.push(key + '__' + i);
    cache.removeAll(keys);
  } catch(e) {}
}

/* High-level wrapper: pakai cache atau fetch fresh */
function _cached(key, ttlSec, forceFresh, fetcher) {
  key = CACHE_VER + ':' + key;
  if (forceFresh) {
    _cacheBust(key);
  } else {
    var hit = _cacheGet(key);
    if (hit !== null) return hit;
  }
  var fresh = fetcher();
  if (fresh && fresh.success !== false) {
    _cachePut(key, fresh, ttlSec);
  }
  return fresh;
}

/* Endpoint: bust SEMUA cache (dipanggil dari tombol "Update Data") */
function bustAllCache() {
  try {
    var cache = CacheService.getScriptCache();
    // Hapus per key yang kita tahu (CacheService tidak support wildcard)
    var sheets = ['Purchase NB', 'Purchase PC'];
    sheets.forEach(function(s) {
      _cacheBust(CACHE_VER + ':rep_' + s);
    });
    _cacheBust(CACHE_VER + ':master_Master SKU NB');
    _cacheBust(CACHE_VER + ':master_Master SKU PC');
    _cacheBust(CACHE_VER + ':lastBuy');
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// ENDPOINT GABUNGAN — dipecah 2 level agar tabel CEPAT tampil:
// getBundleLight: report + status saja (sheet Purchase — ringan)
// getBundleHeavy: master + harga beli (sheet Master + Vendor — berat)
// getInitialData: getBundleLight + daftar sheet (buka app pertama kali)
// ============================================================
function getBundleLight(sheetName, forceFresh) {
  try {
    const report   = getReportData(sheetName, forceFresh);
    const statuses = (report && report.statuses) ? report.statuses : [];
    return {
      success:  report.success,
      report:   report,
      statuses: statuses,
      cached:   !!(report && report._fromCache)
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getBundleHeavy(sheetName, forceFresh) {
  try {
    // Auto-detect master sheet:
    //   "Purchase NB" → "Master SKU NB"
    //   "Purchase PC" → "Master SKU PC" (kalau ada; kalau tidak ditemukan,
    //                  getMasterData return success:false → master=null untuk PC)
    var masterName = null;
    var sn = String(sheetName || '');
    if (/purchase\s*nb/i.test(sn))      masterName = 'Master SKU NB';
    else if (/purchase\s*pc/i.test(sn)) masterName = 'Master SKU PC';
    var master  = masterName ? getMasterData(masterName, forceFresh) : null;
    const lastBuy = getLastPurchaseMap(forceFresh);
    return {
      success: true,
      master:  master,
      lastBuy: lastBuy
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getInitialData(sheetName, forceFresh) {
  try {
    const sn = getSheetNames();
    const sheets = (sn.success && sn.sheets && sn.sheets.length > 0)
      ? sn.sheets : ['Purchase NB','Purchase PC'];
    const active = sheetName || sheets[0];
    const bundle = getBundleLight(active, forceFresh);
    bundle.sheets = sheets;
    bundle.activeSheet = active;
    return bundle;
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// AMBIL NAMA SHEET (hanya yang mengandung "purchase")
// ============================================================
function getSheetNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const all = ss.getSheets().map(s => s.getName());
    // Exact match saja — "Purchase NB" dan "Purchase PC", bukan "Note Purchase ..."
    const filtered = all.filter(n => {
      const t = n.trim().toLowerCase();
      return t === 'purchase nb' || t === 'purchase pc';
    });
    return { success: true, sheets: filtered.length > 0 ? filtered : all.slice(0,2) };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// AMBIL DATA REPORT  (dengan cache layer — TTL 90 detik)
// ============================================================
function getReportData(sheetName, forceFresh) {
  return _cached('rep_' + sheetName, 90, forceFresh, function() {
    return _getReportDataRaw(sheetName);
  });
}

function _getReportDataRaw(sheetName) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const sh  = ss.getSheetByName(sheetName);
    if (!sh) return { success: false, error: 'Sheet "' + sheetName + '" tidak ditemukan.' };

    const lastRow = sh.getLastRow();
    if (lastRow < CONFIG.DATA_START_ROW + 1) {
      return { success: true, sheetName: sheetName, departments: [], statuses: [] };
    }

    // Ambil tepat CONFIG.TOTAL_COLS kolom — hindari baca kolom kosong ribuan
    const data = sh.getRange(1, 1, lastRow, CONFIG.TOTAL_COLS).getValues();

    // Status bar dari baris 1 (sudah ikut terbaca) — hemat 1 baca sheet
    const statuses = [];
    const statusRow = data[CONFIG.ROW_STATUS] || [];
    for (let si = CONFIG.COL_STATUS_START; si < statusRow.length; si++) {
      const sv = String(statusRow[si] || '').trim();
      if (sv) statuses.push(sv);
    }

    // PO header (mis: "PO : 21 Hari") — di baris 1, kolom Y (24) ke kanan.
    // Cell bisa terpisah: Y1="PO :", Z1=21, AA1="Hari"
    // Atau jadi satu cell: "PO : 21 Hari" (merged)
    // → scan kolom 24-30, ambil angka pertama yang ditemukan.
    let poHeader = 0;
    for (let pc = 24; pc <= Math.min(30, statusRow.length - 1); pc++) {
      const cell = statusRow[pc];
      if (typeof cell === 'number' && cell > 0) { poHeader = cell; break; }
      const m = String(cell || '').match(/(\d+(?:\.\d+)?)/);
      if (m) {
        const n = parseFloat(m[1]);
        if (!isNaN(n) && n > 0) { poHeader = n; break; }
      }
    }

    const departments = [];
    let currentDept = null;

    for (let r = CONFIG.DATA_START_ROW; r < data.length; r++) {
      const row   = data[r];
      const codeB = String(row[CONFIG.COL_CODE] || '').trim();
      const nameP = String(row[CONFIG.COL_NAME] || '').trim();

      if (!codeB && !nameP) continue; // baris kosong

      // -------------------------------------------------------
      // DETEKSI DEPARTEMEN:
      // Syarat PASTI departemen: kolom B == kolom P (identik)
      // -------------------------------------------------------
      const sameBP = codeB !== '' && nameP !== '' &&
                     codeB.toUpperCase() === nameP.toUpperCase();

      // Kode produk selalu mengandung '-' (misal NB-APP-MACBOOK-MHFA4ID/A)
      const looksLikeProductCode = codeB.includes('-') || codeB.includes('/');

      const isDept = sameBP ||
                     (!looksLikeProductCode && codeB !== '' && nameP === '') ||
                     (!codeB && nameP !== '');

      if (isDept) {
        currentDept = { name: nameP || codeB, items: [] };
        departments.push(currentDept);
      } else if (looksLikeProductCode && currentDept) {
        currentDept.items.push(parseItem(row, codeB, nameP));
      } else if (!currentDept && looksLikeProductCode) {
        // Produk sebelum departemen pertama — buat dept "Lainnya"
        currentDept = { name: 'Lainnya', items: [] };
        departments.push(currentDept);
        currentDept.items.push(parseItem(row, codeB, nameP));
      }
    }

    // Hapus departemen yang tidak punya item
    const clean = departments.filter(d => d.items && d.items.length > 0);

    return {
      success: true,
      sheetName: sheetName,
      departments: clean,
      statuses: statuses,
      poHeader: poHeader,
      generatedAt: new Date().toLocaleString('id-ID')
    };

  } catch(e) {
    // CATATAN: hindari '\n' di string log (rawan rusak saat copy-paste).
    Logger.log('getReportData error: ' + e.toString() + ' | stack: ' + (e.stack || ''));
    return { success: false, error: e.toString() };
  }
}

// Helper: parse satu baris produk
function parseItem(row, code, name) {
  const outCabang = {};
  CONFIG.CABANG.forEach(function(cab, i) {
    outCabang[cab] = toNum(row[CONFIG.COL_OUT_START + i]);
  });

  // Stok per cabang (kolom C-I = index 2-8, sesuai urutan CABANG)
  const stokCabang = {};
  CONFIG.CABANG.forEach(function(cab, i) {
    stokCabang[cab] = toNum(row[2 + i]);
  });

  return {
    code:       code,
    name:       name,
    subTotal:   toNum(row[CONFIG.COL_SUB_TOTAL]),
    rusak:      toNum(row[CONFIG.COL_RUSAK]),
    so:         toNum(row[CONFIG.COL_SO]),
    dp:         toNum(row[CONFIG.COL_DP]),
    kk:         toNum(row[CONFIG.COL_KK]),
    pj:         toNum(row[CONFIG.COL_PJ]),
    allStok:    toNum(row[CONFIG.COL_ALL_STOK_FULL]),
    stokCabang: stokCabang,
    outCabang:  outCabang,
    allOut:     toNum(row[CONFIG.COL_ALL_OUT]),
    outQty:     toNum(row[CONFIG.COL_OUT_QTY]),
    orderPlus:  toNum(row[CONFIG.COL_ORDER_PLUS]),
    orderMinus: toNum(row[CONFIG.COL_ORDER_MINUS]),
    ikira2:     row[CONFIG.COL_IKIRA2] != null ? row[CONFIG.COL_IKIRA2] : '',
    ketKsg:     String(row[CONFIG.COL_KET_KSG]  || ''),
    po:         String(row[CONFIG.COL_PO]        || ''),
    ketKeep:    toNum(row[CONFIG.COL_KET_KEEP])
  };
}

// ============================================================
// DATA MASTER SKU  (dengan cache layer — TTL 600 detik = 10 menit)
// Master SKU jarang berubah, cocok di-cache lama
// ============================================================
function getMasterData(masterSheetName, forceFresh) {
  return _cached('master_' + masterSheetName, 600, forceFresh, function() {
    return _getMasterDataRaw(masterSheetName);
  });
}

// ============================================================
// HELPER: Baca peta { CODE_UPPER -> value } dari sheet eksternal.
// kind: 'number' (toNum) atau 'text' (string trim).
// Return {} kalau sheet tidak ada / error (graceful degradation).
// ============================================================
function _readExternalKeyValueMap(spreadsheet, sheetName, colKey, colValue, kind) {
  try {
    const sh = spreadsheet.getSheetByName(sheetName);
    if (!sh) {
      Logger.log('External sheet "' + sheetName + '" tidak ditemukan');
      return {};
    }
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return {};
    const lastCol = Math.max(colKey, colValue) + 1;
    const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const map = {};
    for (let r = 1; r < data.length; r++) {  // skip baris 1 (header)
      const key = String(data[r][colKey] || '').trim();
      if (!key) continue;
      const upKey = key.toUpperCase();
      if (kind === 'number') {
        const n = toNum(data[r][colValue]);
        // ambil yg pertama atau yg lebih besar (kalau duplicate, pilih > 0)
        if (!(upKey in map) || (n > 0 && map[upKey] === 0)) map[upKey] = n;
      } else {
        const s = String(data[r][colValue] || '').trim();
        if (s && !map[upKey]) map[upKey] = s;
      }
    }
    return map;
  } catch(e) {
    Logger.log('readExternalKV error (' + sheetName + '): ' + e.toString());
    return {};
  }
}

// ============================================================
// HELPER: Baca SEMUA field NB dari sheet Laptop_Products eksternal
// (sekali baca untuk semua kolom yang dibutuhkan).
// Return: { map, sets } dimana sets={typeLaptop, ram, storage, size, typeProc}
// ============================================================
function _readNBProductsFromExternal(spreadsheet) {
  try {
    const cfg = EXTERNAL_SOURCES.NB.productCols;
    const sh = spreadsheet.getSheetByName(EXTERNAL_SOURCES.NB.productSheet);
    if (!sh) {
      Logger.log('External sheet Laptop_Products tidak ditemukan');
      return { map: {}, typeLaptops: [], typeProcs: [], rams: [], storages: [], sizes: [], priceMin: 0, priceMax: 0 };
    }
    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return { map: {}, typeLaptops: [], typeProcs: [], rams: [], storages: [], sizes: [], priceMin: 0, priceMax: 0 };
    }
    // Baca sampai kolom paling kanan yang dibutuhkan (max dari semua col index + 1)
    const maxCol = Math.max(cfg.code, cfg.typeLaptop, cfg.typeProc, cfg.ram, cfg.storage, cfg.size, cfg.harga) + 1;
    const data = sh.getRange(1, 1, lastRow, maxCol).getValues();

    const map = {};
    const laptopSet = {}, procSet = {}, ramSet = {}, storageSet = {}, sizeSet = {};
    let priceMin = Infinity, priceMax = -Infinity;

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const code = String(row[cfg.code] || '').trim();
      if (!code) continue;
      const upCode = code.toUpperCase();

      const typeLaptop = String(row[cfg.typeLaptop] || '').trim();
      const typeProc   = String(row[cfg.typeProc]   || '').trim();
      // RAM: ambil 5 karakter awal sesuai requirement sebelumnya
      const ram        = String(row[cfg.ram]        || '').substring(0, 5).trim();
      const storage    = String(row[cfg.storage]    || '').trim();
      // Size layar: ambil 3 karakter awal
      const size       = String(row[cfg.size]       || '').substring(0, 3).trim();
      const harga      = toNum(row[cfg.harga]);

      map[upCode] = {
        typeLaptop: typeLaptop,
        typeProc:   typeProc,
        processor:  '',          // diisi nanti via vlookup
        ram:        ram,
        storage:    storage,
        size:       size,
        harga:      harga,
        note:       ''           // diisi nanti dari Struktur_Harga_NB
      };

      if (typeLaptop) laptopSet[typeLaptop] = true;
      if (typeProc)   procSet[typeProc]     = true;
      if (ram)        ramSet[ram]           = true;
      if (storage)    storageSet[storage]   = true;
      if (size)       sizeSet[size]         = true;
      if (harga > 0) {
        if (harga < priceMin) priceMin = harga;
        if (harga > priceMax) priceMax = harga;
      }
    }

    if (priceMin === Infinity)  priceMin = 0;
    if (priceMax === -Infinity) priceMax = 0;

    return {
      map: map,
      typeLaptops: Object.keys(laptopSet).sort(),
      typeProcs:   Object.keys(procSet).sort(),
      rams:        Object.keys(ramSet).sort(),
      storages:    Object.keys(storageSet).sort(),
      sizes:       Object.keys(sizeSet).sort(function(a, b){
                     var na = parseFloat(a) || 0;
                     var nb = parseFloat(b) || 0;
                     if (na !== nb) return na - nb;
                     return a.localeCompare(b);
                   }),
      priceMin: priceMin,
      priceMax: priceMax
    };
  } catch(e) {
    Logger.log('_readNBProductsFromExternal error: ' + e.toString());
    return { map: {}, typeLaptops: [], typeProcs: [], rams: [], storages: [], sizes: [], priceMin: 0, priceMax: 0 };
  }
}

// ============================================================
// HELPER: Vlookup typeProc -> processor dari Master SKU NB (active).
// Hanya baca kolom P-Q. Return: { TYPE_PROC_UPPER -> processor }
// ============================================================
function _readProcessorLookupFromActive() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(PROCESSOR_LOOKUP.sheetName);
    if (!sh) return {};
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return {};

    // Baca hanya kolom P-Q (2 kolom) — getRange(startRow, startCol, numRows, numCols)
    const startCol = PROCESSOR_LOOKUP.colTypeProc + 1;     // 1-based -> 16 (kolom P)
    const numCols  = PROCESSOR_LOOKUP.colProcessor - PROCESSOR_LOOKUP.colTypeProc + 1;
    const data = sh.getRange(1, startCol, lastRow, numCols).getValues();

    const map = {};
    for (let r = 1; r < data.length; r++) {
      const tp   = String(data[r][0] || '').trim();
      const proc = String(data[r][1] || '').trim();
      if (tp && proc && !map[tp.toUpperCase()]) {
        map[tp.toUpperCase()] = proc;
      }
    }
    return map;
  } catch(e) {
    Logger.log('_readProcessorLookupFromActive error: ' + e.toString());
    return {};
  }
}

function _getMasterDataRaw(masterSheetName) {
  try {
    const isPC = /master\s*sku\s*pc/i.test(String(masterSheetName));
    const ext = SpreadsheetApp.openById(EXTERNAL_SOURCES.spreadsheetId);

    if (isPC) {
      // ───── PC: hanya harga + note dari sumber eksternal ─────
      const cfg = EXTERNAL_SOURCES.PC;
      const hargaMap = _readExternalKeyValueMap(ext, cfg.productSheet, cfg.productCols.code, cfg.productCols.harga, 'number');
      const noteMap  = _readExternalKeyValueMap(ext, cfg.noteSheet,    cfg.noteCols.code,    cfg.noteCols.note,    'text');

      const map = {};
      let priceMin = Infinity, priceMax = -Infinity;
      const allCodes = {};
      Object.keys(hargaMap).forEach(function(c){ allCodes[c] = true; });
      Object.keys(noteMap).forEach(function(c){ allCodes[c] = true; });

      Object.keys(allCodes).forEach(function(code){
        const harga = hargaMap[code] || 0;
        map[code] = {
          typeLaptop: '', typeProc: '', processor: '',
          ram: '', storage: '', size: '',
          harga: harga,
          note:  noteMap[code] || ''
        };
        if (harga > 0) {
          if (harga < priceMin) priceMin = harga;
          if (harga > priceMax) priceMax = harga;
        }
      });
      if (priceMin === Infinity)  priceMin = 0;
      if (priceMax === -Infinity) priceMax = 0;

      return {
        success: true,
        map: map,
        typeLaptops: [], typeProcs: [], processors: [],
        rams: [], storages: [], sizes: [],
        priceMin: priceMin, priceMax: priceMax
      };
    }

    // ───── NB: full read dari Laptop_Products + Note + Processor lookup ─────
    const products  = _readNBProductsFromExternal(ext);
    const noteMap   = _readExternalKeyValueMap(
      ext,
      EXTERNAL_SOURCES.NB.noteSheet,
      EXTERNAL_SOURCES.NB.noteCols.code,
      EXTERNAL_SOURCES.NB.noteCols.note,
      'text'
    );
    const procLookup = _readProcessorLookupFromActive();

    // Merge: tambahkan note + processor ke map products
    const processorSet = {};
    Object.keys(products.map).forEach(function(code){
      const entry = products.map[code];
      // Note
      if (noteMap[code]) entry.note = noteMap[code];
      // Processor (vlookup typeProc -> processor)
      if (entry.typeProc) {
        const proc = procLookup[entry.typeProc.toUpperCase()] || '';
        if (proc) {
          entry.processor = proc;
          processorSet[proc] = true;
        }
      }
    });

    // Tambahkan SKU yang ada di noteMap tapi tidak di products (rare edge case)
    Object.keys(noteMap).forEach(function(code){
      if (!products.map[code]) {
        products.map[code] = {
          typeLaptop: '', typeProc: '', processor: '',
          ram: '', storage: '', size: '',
          harga: 0,
          note: noteMap[code]
        };
      }
    });

    return {
      success: true,
      map: products.map,
      typeLaptops: products.typeLaptops,
      typeProcs:   products.typeProcs,
      processors:  Object.keys(processorSet).sort(),
      rams:        products.rams,
      storages:    products.storages,
      sizes:       products.sizes,
      priceMin:    products.priceMin,
      priceMax:    products.priceMax
    };

  } catch(e) {
    Logger.log('getMasterData error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// RIWAYAT PEMBELIAN DARI SHEET CEK VENDOR
// ============================================================
function getVendorHistory(itemCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Cari sheet bernama "Cek Vendor" (fleksibel)
    const vendorSh = ss.getSheets().find(function(s) {
      return s.getName().toLowerCase().replace(/\s/g,'').includes('cekvendor') ||
             s.getName().toLowerCase().replace(/\s/g,'').includes('vendor');
    });

    if (!vendorSh) return { success: false, error: 'Sheet "Cek Vendor" tidak ditemukan.' };

    const lastRow = vendorSh.getLastRow();
    if (lastRow <= VENDOR_CONFIG.DATA_START_ROW) return { success: true, history: [] };

    const startRow1 = VENDOR_CONFIG.DATA_START_ROW + 1; // 1-based untuk getRange
    const numRows   = lastRow - VENDOR_CONFIG.DATA_START_ROW;
    const data = vendorSh.getRange(startRow1, 1, numRows, VENDOR_CONFIG.TOTAL_COLS).getValues();

    const needle = itemCode.trim().toUpperCase();

    const filtered = [];
    for (var i = 0; i < data.length; i++) {
      var row    = data[i];
      var rowSku = String(row[VENDOR_CONFIG.COL_CODE] || '').trim().toUpperCase();
      if (!rowSku) continue;

      // Exact match SKU kolom B vs needle
      if (rowSku === needle) {
        var tglRaw = row[VENDOR_CONFIG.COL_DATE];
        var tanggal = '';
        if (tglRaw instanceof Date && !isNaN(tglRaw)) {
          tanggal = Utilities.formatDate(tglRaw, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        } else {
          tanggal = String(tglRaw || '');
        }

        var tglSort = (tglRaw instanceof Date && !isNaN(tglRaw)) ? tglRaw.getTime() : 0;

        filtered.push({
          tglSort:     tglSort,
          tanggal:     tanggal,
          vendor:      String(row[VENDOR_CONFIG.COL_VENDOR]       || ''),
          qty:         toNum(row[VENDOR_CONFIG.COL_QTY]),
          hargaNormal: fmtRupiah(row[VENDOR_CONFIG.COL_HARGA_NORMAL]),
          hargaDiskon: fmtRupiah(row[VENDOR_CONFIG.COL_HARGA_DISKON]),
          keterangan:  String(row[VENDOR_CONFIG.COL_KET]          || '')
        });
      }
    }

    // Sort descending by tanggal — terbaru di index 0, ambil 10 teratas
    filtered.sort(function(a, b) { return b.tglSort - a.tglSort; });
    var last10 = filtered.slice(0, 10);
    last10.forEach(function(h) { delete h.tglSort; });
    return { success: true, history: last10 };

  } catch(e) {
    Logger.log('getVendorHistory error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// PETA HARGA BELI TERAKHIR per SKU  (dengan cache — TTL 180 detik = 3 menit)
// Cek Vendor sheet bisa cukup besar, dan jarang berubah dalam menit-an
// ============================================================
function getLastPurchaseMap(forceFresh) {
  return _cached('lastBuy', 180, forceFresh, function() {
    return _getLastPurchaseMapRaw();
  });
}

function _getLastPurchaseMapRaw() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vendorSh = ss.getSheets().find(function(s) {
      return s.getName().toLowerCase().replace(/\s/g,'').includes('cekvendor') ||
             s.getName().toLowerCase().replace(/\s/g,'').includes('vendor');
    });
    if (!vendorSh) return { success: true, map: {} };

    const lastRow = vendorSh.getLastRow();
    if (lastRow <= VENDOR_CONFIG.DATA_START_ROW) return { success: true, map: {} };

    const startRow1 = VENDOR_CONFIG.DATA_START_ROW + 1;
    const numRows   = lastRow - VENDOR_CONFIG.DATA_START_ROW;
    const data = vendorSh.getRange(startRow1, 1, numRows, VENDOR_CONFIG.TOTAL_COLS).getValues();

    const map = {}; // SKU_UPPER -> { ts, harga, tanggal }

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var sku = String(row[VENDOR_CONFIG.COL_CODE] || '').trim().toUpperCase();
      if (!sku) continue;

      var tglRaw = row[VENDOR_CONFIG.COL_DATE];
      var ts = (tglRaw instanceof Date && !isNaN(tglRaw)) ? tglRaw.getTime() : 0;

      // Harga beli = "Harga Normal" (kolom F)
      var hn = parseFloat(String(row[VENDOR_CONFIG.COL_HARGA_NORMAL] || '').replace(/[^0-9.-]/g, ''));
      var harga = (!isNaN(hn) && hn > 0) ? hn : 0;

      var tanggal = (tglRaw instanceof Date && !isNaN(tglRaw))
        ? Utilities.formatDate(tglRaw, Session.getScriptTimeZone(), 'dd/MM/yyyy')
        : String(tglRaw || '');

      var prev = map[sku];
      /* Simpan hanya jika: belum ada prev, ATAU ts baru > ts lama (lebih baru)
         Jangan timpa entry valid (ts>0) dengan entry tanpa tanggal (ts=0) */
      if (!prev) {
        map[sku] = { ts: ts, harga: harga, tanggal: tanggal };
      } else if (ts > 0 && ts >= prev.ts) {
        map[sku] = { ts: ts, harga: harga, tanggal: tanggal };
      }
    }

    // Buang field bantu 'ts'
    Object.keys(map).forEach(function(k){ delete map[k].ts; });

    return { success: true, map: map };

  } catch(e) {
    Logger.log('getLastPurchaseMap error: ' + e.toString());
    return { success: false, error: e.toString(), map: {} };
  }
}

// ============================================================
// AMBIL STATUS BAR (baris 1: Data Stok OK, Sell Out OK, dll)
// ============================================================
function getStatusBar(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { success: false, error: 'Sheet tidak ditemukan' };
    const row = sh.getRange(1, 1, 1, 15).getValues()[0];
    const statuses = [];
    for (var i = 2; i < row.length; i++) {
      const v = String(row[i] || '').trim();
      if (v) statuses.push(v);
    }
    return { success: true, statuses: statuses };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// Refresh: ambil ulang report + status tanpa reload browser
function refreshReport(sheetName) {
  try {
    const report = getReportData(sheetName);
    const status = getStatusBar(sheetName);
    return {
      success: report.success,
      report:  report,
      status:  status,
      refreshedAt: new Date().toLocaleString('id-ID')
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// PUSH JUAL — SKU yang stok > 0 tapi pembelian terakhir < cutoff
// ============================================================
function getPushJualData(cutoffDate) {
  try {
    var cutoff = new Date(cutoffDate);
    if (isNaN(cutoff.getTime())) {
      return { success: false, error: 'Tanggal cutoff tidak valid.' };
    }

    // 1) Ambil peta pembelian terakhir per SKU
    var lbResult = getLastPurchaseMap();
    var lbMap = (lbResult && lbResult.success && lbResult.map) ? lbResult.map : {};

    // 2) Baca kedua sheet Purchase NB dan Purchase PC
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ['Purchase NB', 'Purchase PC'];
    var result = { success: true, nb: [], pc: [] };
    var now = new Date();

    for (var si = 0; si < sheets.length; si++) {
      var sheetName = sheets[si];
      var sh = ss.getSheetByName(sheetName);
      if (!sh) continue;

      var lastRow = sh.getLastRow();
      if (lastRow < CONFIG.DATA_START_ROW + 1) continue;

      var data = sh.getRange(1, 1, lastRow, CONFIG.TOTAL_COLS).getValues();
      var items = [];

      for (var r = CONFIG.DATA_START_ROW; r < data.length; r++) {
        var row = data[r];
        var code = String(row[CONFIG.COL_CODE] || '').trim();
        var name = String(row[CONFIG.COL_NAME] || '').trim();

        if (!code) continue;
        // Hanya produk (mengandung '-' atau '/')
        var isProduct = code.includes('-') || code.includes('/');
        if (!isProduct) continue;

        // Stok > 0
        var stok = toNum(row[CONFIG.COL_ALL_STOK_FULL]);
        if (stok <= 0) continue;

        // Cek riwayat beli
        var codeUpper = code.toUpperCase();
        var lb = lbMap[codeUpper] || null;

        var lastBuyDate = '';
        var usia = '';
        var keterangan = '';

        if (!lb || !lb.tanggal) {
          // Tidak punya riwayat beli
          keterangan = 'Copotan / Upgrade';
        } else {
          // Parse tanggal dd/MM/yyyy
          var parts = lb.tanggal.split('/');
          var buyDate = null;
          if (parts.length === 3) {
            buyDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
          }

          if (!buyDate || isNaN(buyDate.getTime())) {
            keterangan = 'Copotan / Upgrade';
          } else if (buyDate >= cutoff) {
            // Pembelian terakhir SETELAH cutoff → tidak masuk push jual
            continue;
          } else {
            lastBuyDate = lb.tanggal;
            // Hitung usia dalam hari
            var diffMs = now.getTime() - buyDate.getTime();
            var diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
            if (diffDays >= 60) {
              var diffMonths = Math.floor(diffDays / 30);
              usia = diffMonths + ' bulan';
            } else {
              usia = diffDays + ' hari';
            }
          }
        }

        items.push({
          code: code,
          name: name,
          stok: stok,
          lastBuyDate: lastBuyDate,
          usia: usia,
          keterangan: keterangan
        });
      }

      if (si === 0) result.nb = items;
      else result.pc = items;
    }

    return result;

  } catch (e) {
    Logger.log('getPushJualData error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// HELPERS
// ============================================================
function toNum(v) {
  var n = parseFloat(String(v || 0));
  return isNaN(n) ? 0 : n;
}

function fmtRupiah(val) {
  if (val === '' || val === null || val === undefined) return '-';
  var n = parseFloat(String(val).replace(/[^0-9.-]/g, ''));
  if (isNaN(n)) return String(val) || '-';
  if (n === 0) return 'Rp0';
  return 'Rp' + Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}
