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
// Master SKU NB -> kolom: A=Kode, C=Type Laptop, J=Type Proc, L=Harga
// ============================================================
const MASTER_CONFIG = {
  COL_CODE:         0,   // A: Kode Barang (dicocokkan ke Purchase kolom B)
  COL_TYPE_LAPTOP:  2,   // C: Type Laptop
  COL_TYPE_PROC:    9,   // J: Type Proc
  COL_HARGA:       11,   // L: Harga (untuk range)
  TOTAL_COLS:      12,   // Baca kolom A-L
  DATA_START_ROW:   1    // Index 1 = baris ke-2 (baris 1 = header)
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
// ENDPOINT GABUNGAN — dipecah 2 level agar tabel CEPAT tampil:
// getBundleLight: report + status saja (sheet Purchase — ringan)
// getBundleHeavy: master + harga beli (sheet Master + Vendor — berat)
// getInitialData: getBundleLight + daftar sheet (buka app pertama kali)
// ============================================================
function getBundleLight(sheetName) {
  try {
    const report   = getReportData(sheetName);
    const statuses = (report && report.statuses) ? report.statuses : [];
    return {
      success:  report.success,
      report:   report,
      statuses: statuses
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getBundleHeavy(sheetName) {
  try {
    // Master SKU hanya untuk Purchase NB
    const masterName = /purchase\s*nb/i.test(String(sheetName || '')) ? 'Master SKU NB' : null;
    const master  = masterName ? getMasterData(masterName) : null;
    const lastBuy = getLastPurchaseMap();
    return {
      success: true,
      master:  master,
      lastBuy: lastBuy
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getInitialData(sheetName) {
  try {
    const sn = getSheetNames();
    const sheets = (sn.success && sn.sheets && sn.sheets.length > 0)
      ? sn.sheets : ['Purchase NB','Purchase PC'];
    const active = sheetName || sheets[0];
    const bundle = getBundleLight(active);
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
// AMBIL DATA REPORT
// ============================================================
function getReportData(sheetName) {
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
// DATA MASTER SKU (untuk filter Type Laptop / Type Proc / Harga)
// ============================================================
function getMasterData(masterSheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(masterSheetName);
    if (!sh) return { success: false, error: 'Sheet "' + masterSheetName + '" tidak ditemukan.' };

    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return { success: true, map: {}, typeLaptops: [], typeProcs: [], priceMin: 0, priceMax: 0 };
    }

    const data = sh.getRange(1, 1, lastRow, MASTER_CONFIG.TOTAL_COLS).getValues();

    const map       = {};
    const laptopSet = {};
    const procSet   = {};
    let priceMin = Infinity, priceMax = -Infinity;

    for (let r = MASTER_CONFIG.DATA_START_ROW; r < data.length; r++) {
      const row  = data[r];
      const code = String(row[MASTER_CONFIG.COL_CODE] || '').trim();
      // Lewati baris kosong / baris departemen (tanpa '-' & '/')
      if (!code || (code.indexOf('-') === -1 && code.indexOf('/') === -1)) continue;

      const typeLaptop = String(row[MASTER_CONFIG.COL_TYPE_LAPTOP] || '').trim();
      const typeProc   = String(row[MASTER_CONFIG.COL_TYPE_PROC]   || '').trim();
      const harga      = toNum(row[MASTER_CONFIG.COL_HARGA]);

      map[code.toUpperCase()] = {
        typeLaptop: typeLaptop,
        typeProc:   typeProc,
        harga:      harga
      };

      if (typeLaptop) laptopSet[typeLaptop] = true;
      if (typeProc)   procSet[typeProc]     = true;
      if (harga > 0) {
        if (harga < priceMin) priceMin = harga;
        if (harga > priceMax) priceMax = harga;
      }
    }

    if (priceMin === Infinity)  priceMin = 0;
    if (priceMax === -Infinity) priceMax = 0;

    return {
      success:     true,
      map:         map,
      typeLaptops: Object.keys(laptopSet).sort(),
      typeProcs:   Object.keys(procSet).sort(),
      priceMin:    priceMin,
      priceMax:    priceMax
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
// PETA HARGA BELI TERAKHIR per SKU (dari sheet Cek Vendor)
// Dipakai untuk kolom "Harga Terakhir Beli" di tabel.
// ============================================================
function getLastPurchaseMap() {
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
      if (!prev || ts >= prev.ts) {
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
