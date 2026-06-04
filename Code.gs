/* SMART INVENTORY WMS
  Backend Logic untuk menghubungkan Spreadsheet dengan Web App
*/

// --- 1. CONFIGURATION & ROUTING ---

// ==========================================================
// GLOBAL: Ambil spreadsheet berdasarkan sheet_id penyewa
// ==========================================================
var LICENSE_API_URL = 'https://dataeasy.web.id/api/verify.php';

// GLOBAL: sheet_id aktif untuk request ini (reset setiap request)
var _ACTIVE_SHEET_ID = null;

function _getSheetId() {
  if (_ACTIVE_SHEET_ID && _ACTIVE_SHEET_ID.length > 10) return _ACTIVE_SHEET_ID;

  // Fallback ke ScriptProperties (backward compatibility single-tenant)
  var sp = PropertiesService.getScriptProperties();
  var fallback = sp.getProperty('bs_sheet_id');
  if (fallback && fallback.length > 10) return fallback;

  return null;
}

function _getSS() {
  var sheetId = _getSheetId();
  if (sheetId) {
    try { return SpreadsheetApp.openById(sheetId); } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _getSheet(name) {
  return _getSS().getSheetByName(name);
}

// Set sheet_id (dipanggil frontend setiap app load)
function setActiveSheetId(sheetId) {
  if (sheetId && sheetId.length > 10) {
    _ACTIVE_SHEET_ID = sheetId;
    PropertiesService.getScriptProperties().setProperty('bs_sheet_id', sheetId);
  }
}

// Resolve token → sheet_id (cek cache dulu, lalu API)
function _resolveToken(token) {
  if (!token) return null;
  var sp = PropertiesService.getScriptProperties();
  var cacheKey = 'sheet_' + token;

  // 1. Cek cache di ScriptProperties
  var cached = sp.getProperty(cacheKey);
  if (cached && cached.length > 10) {
    _ACTIVE_SHEET_ID = cached;
    sp.setProperty('bs_sheet_id', cached);
    return cached;
  }

  // 2. Panggil API untuk resolve
  try {
    var res = UrlFetchApp.fetch(LICENSE_API_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ action: 'get_sheet', code: token }),
      muteHttpExceptions: true
    });
    var data = JSON.parse(res.getContentText());
    if (data.ok && data.sheet_id) {
      _ACTIVE_SHEET_ID = data.sheet_id;
      sp.setProperty(cacheKey, data.sheet_id);
      sp.setProperty('bs_sheet_id', data.sheet_id);
      return data.sheet_id;
    }
  } catch(e) {}
  return null;
}

// WRAPPER: Frontend panggil ini untuk semua operasi. sheetId dikirim setiap kali.
function runWithSheet(sheetId, funcName, args) {
  if (sheetId && sheetId.length > 10) _ACTIVE_SHEET_ID = sheetId;
  var fn = this[funcName];
  if (typeof fn !== 'function') return { success: false, message: 'Fungsi tidak ditemukan: ' + funcName };
  return fn.apply(null, args || []);
}

function doGet(e) {
  // Ambil token dari URL parameter
  var token = (e && e.parameter && e.parameter.token) ? e.parameter.token : '';
  if (token) {
    _resolveToken(token);
  }

  var initialConfig = {
    page_title: 'SmartInv WMS',
    sidebar_name: 'SMART INVENTORY',
    sidebar_subtitle: 'BANTUSELLER V2.0'
  };

  try {
    var ss = _getSS();
    var sheet = ss.getSheetByName('Config');
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var data = sheet.getRange(2, 9, lastRow - 1, 2).getValues(); // Col I & J
        for (var i = 0; i < data.length; i++) {
          var key = String(data[i][0]).trim();
          var val = data[i][1];
          if (key && val && initialConfig.hasOwnProperty(key)) {
            initialConfig[key] = String(val);
          }
        }
      }
    }
  } catch(err) {
    // Gunakan default jika gagal
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.initialConfig = initialConfig;
  return template.evaluate()
      .setTitle(initialConfig.page_title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- CEK LISENSI BUATSIMPEL (dipanggil dari frontend) ---
function checkBuatSimpelLicense(code, email, sheetId) {
  if (!code) return { ok: false, msg: 'Kode kosong' };
  try {
    // Simpan sheet_id per kode lisensi di ScriptProperties
    if (sheetId && sheetId.length > 10) {
      _ACTIVE_SHEET_ID = sheetId;
      var sp = PropertiesService.getScriptProperties();
      sp.setProperty('sheet_' + code, sheetId);
      sp.setProperty('bs_sheet_id', sheetId);
    }
    var url = 'https://dataeasy.web.id/api/verify.php';
    var deviceId = sheetId ? 'gas_' + sheetId : 'gas_' + _getSS().getId();
    var payload = JSON.stringify({ action: 'activate', code: code, email: email || '', device_type: 'gudang', device_id: deviceId, sheet_id: sheetId || '' });
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: payload, muteHttpExceptions: true });
    return JSON.parse(res.getContentText());
  } catch(e) {
    return { ok: false, msg: 'Gagal cek lisensi: ' + e.toString() };
  }
}

function verifyBuatSimpelLicense(code) {
  if (!code) return { ok: false, msg: 'Kode kosong' };
  try {
    var url = 'https://dataeasy.web.id/api/verify.php';
    var deviceId = 'gas_' + _getSS().getId();
    var payload = JSON.stringify({ action: 'verify', code: code, device_id: deviceId });
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: payload, muteHttpExceptions: true });
    return JSON.parse(res.getContentText());
  } catch(e) {
    return { ok: false, msg: 'Offline - grace period' };
  }
}

// --- ENDPOINT UNTUK CHROME EXTENSION (ROAS MASTER PRO) ---
// Menerima data pesanan dari scraper Shopee via HTTP POST
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || '';
    var secret = body.secret || '';

    // Resolve sheet_id dari token jika ada
    var token = body.token || '';
    if (token) {
      _resolveToken(token);
    }

    // Validasi secret key
    var CONFIG_SECRET = getExtensionSecret_();
    if (secret !== CONFIG_SECRET) {
      return buildJsonResponse_({ success: false, message: 'Unauthorized: Invalid secret key' });
    }

    if (action === 'SYNC_ORDERS') {
      var source = String(body.source || 'shopee').toLowerCase(); // 'shopee' atau 'tiktok'
      var result = syncOrdersFromExtension_(body.orders || [], source);
      return buildJsonResponse_(result);
    }

    if (action === 'CHECK_CONNECTION') {
      return buildJsonResponse_({ success: true, message: 'Connected to Smart Inventory WMS' });
    }

    return buildJsonResponse_({ success: false, message: 'Unknown action: ' + action });

  } catch (err) {
    return buildJsonResponse_({ success: false, message: 'Server Error: ' + err.toString() });
  }
}

function buildJsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getExtensionSecret_() {
  // Cek cache dulu di PropertiesService
  try {
    var cached = PropertiesService.getScriptProperties().getProperty('ext_secret');
    if (cached) return cached;
  } catch(e) {}
  // Fallback: baca dari Config sheet
  try {
    var sheet = getSheet('Config');
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 'roasmaster2025';
    var data = sheet.getRange(2, 9, lastRow - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'extension_secret') {
        var secret = String(data[i][1]).trim();
        try { PropertiesService.getScriptProperties().setProperty('ext_secret', secret); } catch(e2) {}
        return secret;
      }
    }
  } catch(e) {}
  return 'roasmaster2025';
}

// Fungsi utama: Sinkronisasi pesanan dari Shopee ke Data_Pesanan
// Menggunakan logika yang sama dengan saveMassData() tapi khusus untuk extension
function syncOrdersFromExtension_(orders, source) {
  if (!orders || orders.length === 0) {
    return { success: false, message: 'Tidak ada data pesanan.' };
  }
  source = String(source || 'shopee').toLowerCase(); // 'shopee' atau 'tiktok'
  var channelName = source === 'tiktok' ? 'TikTok' : 'Shopee';

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    return { success: false, message: 'Server sibuk, coba lagi.' };
  }

  try {
    var ss = _getSS();
    var sheetPesanan = ss.getSheetByName('Data_Pesanan');
    var sheetProduk = ss.getSheetByName('Master_Produk');
    var sheetHistory = ss.getSheetByName('Riwayat_Stok');
    var sLogActivity = ss.getSheetByName('Data_Riwayat');
    var timestamp = new Date();

    // Cache Master Produk (SEKALI BACA untuk seluruh fungsi)
    var dataProduk = sheetProduk.getDataRange().getValues();

    // Map by SKU (case-insensitive) & array untuk fuzzy match - 1 loop saja
    var productBySku = {};
    var productBySkuLower = {};
    var productByName = [];

    for (var i = 1; i < dataProduk.length; i++) {
      var sku = String(dataProduk[i][3]).trim();
      var entry = {
        sku: sku,
        skuInduk: String(dataProduk[i][0]).trim(),
        namaProd: String(dataProduk[i][1]).trim(),
        varName: String(dataProduk[i][4]).trim(),
        hpp: Number(dataProduk[i][9]) || 0,
        sell: Number(dataProduk[i][11]) || 0,
        rowIndex: i // simpan index untuk batch update stok nanti
      };

      if (sku) {
        productBySku[sku] = entry;
        productBySkuLower[sku.toLowerCase()] = entry;
      }
      productByName.push(entry);
    }

    // Normalisasi teks
    function norm(raw) {
      return String(raw || '').toLowerCase().replace(/\s+/g, '').replace(/[-_\/]+/g, ',').trim();
    }

    // Matching produk berdasarkan source
    function findProduct(order) {
      var orderVar = norm(order.variasi);
      var orderNama = norm(order.namaProduk);

      // === MATCH BY NAMA + VARIASI ===
      if (orderNama) {
        for (var p2 = 0; p2 < productByName.length; p2++) {
          var prod2 = productByName[p2];
          var masterNama = norm(prod2.namaProd);
          if (!masterNama) continue;
          if (!(masterNama.includes(orderNama) || orderNama.includes(masterNama))) continue;
          if (orderVar) {
            var masterVar2 = norm(prod2.varName);
            if (masterVar2 === orderVar) return prod2;
            var oP = orderVar.split(',').filter(function(s){return s.length>0}).sort();
            var mP = masterVar2.split(',').filter(function(s){return s.length>0}).sort();
            if (oP.length > 0 && oP.length === mP.length && oP.join(',') === mP.join(',')) return prod2;
          } else {
            return prod2;
          }
        }
      }

      // === FALLBACK: SKU (hanya jika source bukan tiktok) ===
      if (source !== 'tiktok') {
        var orderSkuInduk = String(order.skuInduk || '').trim().toLowerCase();
        if (orderSkuInduk && orderVar) {
          for (var p = 0; p < productByName.length; p++) {
            var prod = productByName[p];
            if (prod.skuInduk.toLowerCase() !== orderSkuInduk) continue;
            var masterVar = norm(prod.varName);
            if (masterVar === orderVar) return prod;
            var oParts = orderVar.split(',').filter(function(s){return s.length>0}).sort();
            var mParts = masterVar.split(',').filter(function(s){return s.length>0}).sort();
            if (oParts.length > 0 && oParts.length === mParts.length && oParts.join(',') === mParts.join(',')) return prod;
          }
        }
        var skuVarLookup = String(order.skuVariasi || '').trim();
        if (skuVarLookup && productBySkuLower[skuVarLookup.toLowerCase()]) {
          return productBySkuLower[skuVarLookup.toLowerCase()];
        }
      }

      return null;
    }

    // Cache pesanan yang sudah ada (cek duplikat per ID + Variasi)
    var existingData = sheetPesanan.getDataRange().getValues();
    var dbKeys = new Set();
    for (var j = 1; j < existingData.length; j++) {
      dbKeys.add(String(existingData[j][0]).trim() + '|' + String(existingData[j][8]).trim().toLowerCase());
    }

    var newRows = [];
    var stockUpdates = [];
    var duplicateCount = 0;
    var successCount = 0;
    var matchedCount = 0;

    orders.forEach(function(order) {
      var orderId = String(order.idPesanan || '').trim();
      if (!orderId) return;

      var orderVar = String(order.variasi || '').trim().toLowerCase();
      var dupKey = orderId + '|' + orderVar;
      if (dbKeys.has(dupKey)) { duplicateCount++; return; }

      var matched = findProduct(order);
      var masterHpp = 0, masterSell = 0, matchedSku = '', matchedSkuInduk = '';

      if (matched) {
        masterHpp = matched.hpp;
        masterSell = matched.sell;
        matchedSku = matched.sku;
        matchedSkuInduk = matched.skuInduk;
        matchedCount++;
      }

      var qty = Number(order.qty) || 1;
      var row = new Array(28).fill('');
      row[0]  = orderId;
      row[1]  = order.tanggal || timestamp;
      row[2]  = channelName; // 'Shopee' atau 'TikTok'
      row[3]  = order.kurir || '';
      row[4]  = order.resi || '';
      row[5]  = order.skuInduk || matchedSkuInduk || '';
      row[6]  = order.namaProduk || '';
      row[7]  = order.skuVariasi || matchedSku || '';
      row[8]  = order.variasi || '';
      row[9]  = qty;
      row[10] = masterSell;
      row[11] = masterSell * qty;
      row[12] = masterHpp;
      row[13] = masterHpp * qty;
      row[14] = 'WAITING'; // TikTok "Selesai" tetap masuk WAITING → nanti sync tracking update
      row[15] = '';
      row[16] = '';
      row[18] = Number(order.diskon) || 0;
      row[19] = Number(order.totalHarga) || (masterSell * qty);
      row[21] = order.namaPembeli || '';
      row[22] = order.noTelp || '';
      row[23] = order.alamat || '';
      row[24] = Number(order.ongkir) || 0;
      row[25] = '';
      row[26] = order.kota || '';
      row[27] = order.provinsi || '';

      newRows.push(row);
      dbKeys.add(dupKey);
      successCount++;

    });

    // ===== SIMPAN PESANAN KE SHEET (BATCH) =====
    if (newRows.length > 0) {
      sheetPesanan.getRange(sheetPesanan.getLastRow() + 1, 1, newRows.length, 28).setValues(newRows);
    }

    // LOG ACTIVITY
    if (sLogActivity && successCount > 0) {
      sLogActivity.appendRow([
        timestamp, 'EXT-' + Date.now(), 'SYNC ' + channelName.toUpperCase(), '-',
        channelName, successCount, 'Chrome Extension',
        'Sync: ' + successCount + ' pesanan ' + channelName + ' (WAITING), Duplikat: ' + duplicateCount
      ]);
    }

    SpreadsheetApp.flush();

    var msg = 'Berhasil sync ' + successCount + ' pesanan (Status: WAITING).';
    if (duplicateCount > 0) msg += ' (' + duplicateCount + ' duplikat diskip)';

    return { success: true, message: msg, synced: successCount, matched: matchedCount, duplicates: duplicateCount };

  } catch (e) {
    return { success: false, message: 'Error sync: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Helper: Mendapatkan Sheet Aktif (Supaya script dinamis saat dicopy)
function getSheet(sheetName) {
  return _getSS().getSheetByName(sheetName);
}

// --- 2. DATA FETCHING (GET) ---

// Ambil Data Config untuk Dropdown
function getConfigData() {
  const sheet = getSheet('Config');
  const data = sheet.getDataRange().getValues();
  // Hapus header baris 1
  data.shift(); 
  return data;
}

// Ambil Dashboard Summary
function getDashboardSummary() {
  const sheetPesanan = getSheet('Data_Pesanan');
  const sheetProduk = getSheet('Master_Produk');
  
  // Ambil data (asumsi baris 1 adalah header)
  const dataPesanan = sheetPesanan.getDataRange().getValues();
  const dataProduk = sheetProduk.getDataRange().getValues();
  
  // 1. KPI COUNTERS
  let waiting = 0, packed = 0, shipped = 0, returned = 0;
  let totalAsset = 0;
  let lowStock = 0;

  // Struktur untuk Chart Trend (7 Hari Terakhir)
  const last7Days = {}; 
  const today = new Date();
  for(let d=6; d>=0; d--) {
    const date = new Date(today);
    date.setDate(today.getDate() - d);
    const key = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    last7Days[key] = 0;
  }

  // Loop Pesanan (Mulai baris 2/index 1)
  for (let i = 1; i < dataPesanan.length; i++) {
    const row = dataPesanan[i];
    const status = row[14]; // Kolom O: Status
    const tgl = row[1];     // Kolom B: Tanggal

    // Hitung Status
    if (status === 'WAITING') waiting++;
    else if (status === 'PACKED') packed++;
    else if (status === 'SHIPPED') shipped++;
    else if (status === 'RETURNED') returned++;

    // Hitung Trend (Berdasarkan Tanggal Pesanan)
    if (tgl instanceof Date) {
      const key = Utilities.formatDate(tgl, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (last7Days.hasOwnProperty(key)) {
        last7Days[key]++;
      }
    }
  }

  // 2. LOGIKA ASET, TOP PRODUK & KATEGORI
  const categoryMap = {};
  const assetList = [];

  for (let i = 1; i < dataProduk.length; i++) {
    const row = dataProduk[i];
    // Pastikan Nama Produk ada (validasi baris kosong)
    if(!row[1]) continue;

    const kategori = row[2] || 'Uncategorized';
    const stok = Number(row[7]) || 0;     // Kolom H
    const nilaiAset = Number(row[10]) || 0; // Kolom K

    // Hitung Total KPI
    totalAsset += nilaiAset;
    if (stok <= 5) lowStock++;

    // Agregasi Kategori (Untuk Pie Chart)
    if (!categoryMap[kategori]) categoryMap[kategori] = 0;
    categoryMap[kategori] += nilaiAset;

    // Kumpulkan data untuk Top 5 Aset
    assetList.push({
      name: row[1],
      category: kategori,
      stock: stok,
      asset: nilaiAset,
      img: row[14] || '' // Kolom O: Main Image
    });
  }

  // Sorting Data untuk Chart/Table
  // A. Top 5 Aset Terbesar
  const topAssets = assetList.sort((a, b) => b.asset - a.asset).slice(0, 5);

  // B. Data Trend (Convert Map to Array)
  const trendLabels = Object.keys(last7Days).map(k => {
    // Format label jadi "27 Dec" biar cantik
    const parts = k.split('-');
    const dateObj = new Date(parts[0], parts[1]-1, parts[2]);
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd MMM");
  });
  const trendValues = Object.values(last7Days);

  // C. Data Kategori (Top 5 Kategori by Value)
  const sortedCats = Object.entries(categoryMap).sort((a,b) => b[1] - a[1]).slice(0, 5);
  const catLabels = sortedCats.map(x => x[0]);
  const catValues = sortedCats.map(x => x[1]);

  return {
    orders: { waiting, packed, shipped, returned },
    inventory: { totalAsset, lowStock },
    charts: {
      trend: { labels: trendLabels, series: trendValues },
      category: { labels: catLabels, series: catValues }
    },
    topProducts: topAssets
  };
}

// Ambil Daftar Pesanan (Untuk Table Order & Gudang)
function getOrderList() {
  const sheet = getSheet('Data_Pesanan');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // Ambil data tanpa header, langsung reverse
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  return data.reverse();
}

// Ambil Master Produk (VERSI BERSIH & BENAR)
function getProductList() {
  const sheet = getSheet('Master_Produk');
  
  // Ambil Semua Data (Termasuk Kolom U yang ada di ujung kanan)
  // getDataRange otomatis mengambil sampai kolom terakhir yang ada datanya
  const data = sheet.getDataRange().getDisplayValues();
  
  // Hapus Header (Baris 1)
  data.shift(); 
  
  // Filter baris kosong
  return data.filter(row => row[1] && row[1].trim() !== "");
}

// --- 3. TRANSACTION LOGIC (POST) ---

// 1. FUNGSI INPUT ORDER MANUAL (DENGAN LOCK SYSTEM)
function addNewOrder(formObject) {
  const lock = LockService.getScriptLock();
  try {
      lock.waitLock(10000); // Tunggu antrian max 10 detik
  } catch (e) {
      return { success: false, message: "Sistem sibuk, coba sesaat lagi." };
  }

  try {
      const ss = _getSS();
      const sheet = ss.getSheetByName('Data_Pesanan');
      const sLogActivity = ss.getSheetByName('Data_Riwayat');
      const user = Session.getActiveUser().getEmail();
      const timestamp = new Date();
      
      let itemsToProcess = [];
      if (formObject.items && Array.isArray(formObject.items)) {
          itemsToProcess = formObject.items;
      } else {
          itemsToProcess.push(formObject);
      }

      let totalQty = 0;
      const newRows = [];

      itemsToProcess.forEach(item => {
          const qty = Number(item.qty) || 0;
          totalQty += qty;
          
          const hargaJualMaster = Number(item.hargaJual) || 0; 
          const diskon = Number(item.diskon) || 0;             
          const hargaBeli = Number(item.hargaBeli) || 0;       
          
          const totalJualStandard = qty * hargaJualMaster;
          const totalMarketplace = totalJualStandard - diskon;
          const totalModal = qty * hargaBeli;

          // [LOGIKA BARU] Text to Columns untuk Wilayah
          // Format dari frontend: "Kecamatan, Kota/Kab, Provinsi"
          let rawLoc = formObject.kecamatanString || ""; 
          let splitLoc = rawLoc.split(",").map(s => s.trim()); // Pisahkan koma & hapus spasi

          let valKecamatan = splitLoc[0] || ""; // Ambil kata pertama
          let valKota      = splitLoc[1] || ""; // Ambil kata kedua
          let valProvinsi  = splitLoc[2] || ""; // Ambil kata ketiga (jika ada)

          newRows.push([
            formObject.idPesanan,       // A
            timestamp,                  // B
            formObject.channel,         // C
            formObject.kurir,           // D
            formObject.resi || '',      // E
            
            item.skuInduk,              // F
            item.namaProduk,            // G
            item.skuVariasi,            // H
            item.variasi,               // I
       
            qty,                        // J
            hargaJualMaster,            // K
            totalJualStandard,          // L
            hargaBeli,          
            totalModal,                 // N
            'WAITING',                  // O
            '', '', '',                 // P-R (Reserved: Packed, Shipped, Returned Time)
            diskon,                     // S
            totalMarketplace,           // T
            
            '',                         // U (Reserved: Kondisi Retur / Stok Masuk) - Kita loncati agar aman
            formObject.custName,        // V - Nama Customer
            formObject.custPhone,       // W - No Telp
            formObject.custAddress,     // X - Alamat
            formObject.ongkir,           // Y - Ongkir

            valKecamatan,               // Z  (Kolom Baru: Kecamatan)
            valKota,                    // AA (Kolom Baru: Kota)
            valProvinsi                 // AB (Kolom Baru: Provinsi)
          ]);
      });

      // Simpan sekaligus (Batch) -> Lebih Cepat
      if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      }
      
      // Log Dashboard
      if (sLogActivity) {
          sLogActivity.appendRow([
              timestamp, formObject.idPesanan, 'NEW ORDER', '-', 
              formObject.channel, totalQty, user, 'Input Manual via Web'
          ]);
      }
      
      SpreadsheetApp.flush();
      return { success: true, message: "Pesanan berhasil disimpan!" };

  } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
  } finally {
      lock.releaseLock();
  }
}

// --- 4. SCANNER LOGIC (THE BRAIN) ---

// 1. FUNGSI SCANNER (EKSEKUSI STOK FISIK)
function processScan(barcode, mode, username, returCondition) {
  const ss = _getSS();
  const sheetPesanan = ss.getSheetByName('Data_Pesanan');
  const sheetProduk = ss.getSheetByName('Master_Produk');
  const sheetHistory = ss.getSheetByName('Riwayat_Stok');
  
  const dataPesanan = sheetPesanan.getDataRange().getValues();
  const timestamp = new Date();

  // Cari Pesanan by Resi
  let foundRowIndex = -1;
  let rowData = [];
  
  for (let i = 1; i < dataPesanan.length; i++) {
    if (String(dataPesanan[i][4]).trim().toLowerCase() === String(barcode).trim().toLowerCase()) { 
      foundRowIndex = i + 1; 
      rowData = dataPesanan[i];
      break;
    }
  }
  
  if (foundRowIndex === -1) return { success: false, message: "Resi tidak ditemukan!" };

  const skuVariasi = rowData[7]; 
  const qtyOrder = Number(rowData[9]); 
  const currentStatus = rowData[14]; 
  const namaProd = rowData[6] + " " + (rowData[8] || "");

  // --- 1. MODE IN (PACKING) ---
  if (mode === 'IN') {
    if (currentStatus !== 'WAITING' && !currentStatus.includes('PRE-ORDER')) {
       return { success: false, message: `Gagal: Status pesanan saat ini '${currentStatus}'.` };
    }
    // Update Status & Waktu Packing (Kolom P / 16)
    sheetPesanan.getRange(foundRowIndex, 15).setValue('PACKED');
    sheetPesanan.getRange(foundRowIndex, 16).setValue(timestamp);
    return { success: true, message: "Paket SIAP KIRIM (Packed)", detail: namaProd };
  }

  // --- 2. MODE OUT (KURIR) ---
  else if (mode === 'OUT') {
    if (currentStatus === 'SHIPPED') return { success: false, message: "Paket SUDAH jalan (Shipped)." };
    if (currentStatus !== 'PACKED') return { success: false, message: "Gagal: Paket belum dipacking!" };

    // Cari Produk
    const dataProduk = sheetProduk.getDataRange().getValues();
    let productRow = -1;
    let currentStock = 0;
    let currentOut = 0; 
    let productHPP = 0;

    for (let p = 1; p < dataProduk.length; p++) {
      if (String(dataProduk[p][3]) === String(skuVariasi)) { 
        productRow = p + 1;
        currentStock = Number(dataProduk[p][7]); 
        currentOut = Number(dataProduk[p][6]);   
        productHPP = Number(dataProduk[p][9]);   
        break;
      }
    }

    if (productRow === -1) return { success: false, message: "SKU Produk tidak ditemukan di Master!" };

    // Eksekusi Keluar
    const newStock = currentStock - qtyOrder;
    const newOut = currentOut + qtyOrder;
    
    // Update Master (BATCH: 2 cell sekaligus + 1 cell aset)
    sheetProduk.getRange(productRow, 7, 1, 2).setValues([[newOut, newStock]]);
    sheetProduk.getRange(productRow, 11).setValue(newStock * productHPP);

    // Update Pesanan (BATCH: status + timestamp)
    sheetPesanan.getRange(foundRowIndex, 15).setValue('SHIPPED'); 
    sheetPesanan.getRange(foundRowIndex, 17).setValue(timestamp); 

    // Log
    sheetHistory.appendRow([
      'OUT-' + Math.floor(Date.now()/1000), timestamp, 'OUT', 
      barcode, skuVariasi, namaProd, rowData[8], qtyOrder, 
      productHPP, qtyOrder * productHPP, newStock, 
      'Penjualan (Scan Out)', username, new Date()
    ]);

    return { success: true, message: `DIKIRIM. Stok -${qtyOrder}, Keluar +${qtyOrder}.` };
  }

  // --- 3. MODE RETUR (BARANG BALIK) ---
  else if (mode === 'RETUR') {
    if (currentStatus === 'RETURNED') return { success: false, message: "Paket sudah diretur sebelumnya." };

    const dataProduk = sheetProduk.getDataRange().getValues();
    let productRow = -1;
    let currentStock = 0;
    let currentMasuk = 0; // Ganti variable dari currentOut jadi currentMasuk
    let productHPP = 0;
    
    for (let p = 1; p < dataProduk.length; p++) {
      if (String(dataProduk[p][3]) === String(skuVariasi)) {
        productRow = p + 1;
        currentStock = Number(dataProduk[p][7]); // Col H (Sisa)
        // Ambil data Masuk dari Kolom U (Index 20)
        currentMasuk = (dataProduk[p].length > 20) ? Number(dataProduk[p][20]) : 0;
        productHPP = Number(dataProduk[p][9]); 
        break;
      }
    }

    let logNote = 'Retur: Kondisi ???';
    let newStock = currentStock;
    let newMasuk = currentMasuk;

    if (returCondition === 'GOOD') {
        newStock = currentStock + qtyOrder;
        newMasuk = currentMasuk + qtyOrder; 
        
        if (productRow !== -1) {
            // BATCH: Update Sisa(H) + Aset(K) bersamaan
            sheetProduk.getRange(productRow, 8).setValue(newStock);
            sheetProduk.getRange(productRow, 11).setValue(newStock * productHPP);
            sheetProduk.getRange(productRow, 21).setValue(newMasuk);
        }
        logNote = 'Retur (Bagus) - Restock';
    } else {
        logNote = 'Retur (Rusak/BS) - Stok Tetap';
    }

    sheetHistory.appendRow([
      'RET-' + Math.floor(Date.now()/1000), timestamp, 'RETUR',
      barcode, skuVariasi, namaProd, rowData[8], qtyOrder,
      0, 0, newStock, logNote, username, new Date()
    ]);

    // Update Status Pesanan
    sheetPesanan.getRange(foundRowIndex, 15).setValue('RETURNED'); // Col O
    sheetPesanan.getRange(foundRowIndex, 18).setValue(timestamp);  // Col R (Waktu Retur)
    
    // >>> BARU: SIMPAN KONDISI RETUR KE KOLOM U (Index 21) <<<
    sheetPesanan.getRange(foundRowIndex, 21).setValue(returCondition); 

    return { success: true, message: `Retur Diterima.` };
  }
  
  return { success: false, message: "Mode tidak valid" };
}

// --- API: HANDLE TRANSACTION (BACKEND YANG BENAR) ---

/* CODE.GS - UPDATE PENYIMPANAN 2 KOLOM URL */

function handleTransaction(payload) {
  const ss = _getSS();
  const sProd = ss.getSheetByName('Master_Produk');
  const sHist = ss.getSheetByName('Riwayat_Stok');
  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date();

  // Auto-ensure header "Admin Shopee" & "Admin Tiktok" di kolom V (22) & W (23)
  var headerV = sProd.getRange(1, 22).getValue();
  if (!headerV || String(headerV).trim() === '') {
    sProd.getRange(1, 22).setValue('Admin Shopee');
  }
  var headerW = sProd.getRange(1, 23).getValue();
  if (!headerW || String(headerW).trim() === '') {
    sProd.getRange(1, 23).setValue('Admin Tiktok');
  }
  
  if (payload.type === 'IN') {
      // --- KASUS 1: PRODUK BARU ---
      if (payload.isNew) {
          const parentSku = payload.parentSku;
          
          let mainImageUrl = "";
          if (payload.mainImage) {
              const fName = parentSku + "_MAIN_" + Date.now() + ".jpg";
              mainImageUrl = saveImageToDrive(payload.mainImage, fName);
          }

          payload.variations.forEach(v => {
             const skuFinal = v.sku;
             
             let variantImageUrl = ""; 
             if (v.image) {
                 const fName = skuFinal + "_" + Date.now() + ".jpg";
                 variantImageUrl = saveImageToDrive(v.image, fName);
             }
      
             const hpp = Number(v.price) || 0;
             const sell = Number(v.sellPrice) || 0;
             const asset = Number(v.qty) * hpp;
             const location = payload.location || '-';
             const marginVal = Number(v.adminPct) || 0;
             const marginTiktok = Number(v.adminTiktok) || 0;

             // appendRow - cara paling stabil
             sProd.appendRow([
                parentSku,          // A
                payload.name,       // B
                payload.category,   // C
                skuFinal,           // D
                v.name,             // E
                Number(v.qty),      // F
                0,                  // G
                Number(v.qty),      // H
                payload.unit,       // I
                hpp,                // J
                asset,              // K
                sell,               // L
                location,           // M
                mainImageUrl,       // N
                variantImageUrl,    // O
                v.barcode || '',    // P
                'StockIn',          // Q
                '',                 // R (bebas untuk fitur produksi)
                '',                 // S
                '',                 // T
                0,                  // U
                marginVal,          // V: Admin Shopee
                marginTiktok        // W: Admin Tiktok
             ]);

             sHist.appendRow(['TRX-'+Date.now(), timestamp, 'IN', payload.ref, skuFinal, payload.name, v.name, v.qty, hpp, asset, v.qty, 'New Product', user, timestamp]);
          });

          return { success: true, message: "Berhasil menambahkan " + payload.variations.length + " variasi produk baru!" };
      } 
      // --- KASUS 2: RESTOCK BARANG LAMA ---
      else {
          const dataProd = sProd.getDataRange().getValues();
          
          // Buat map SKU -> row index untuk lookup cepat & akurat
          // Normalisasi: trim + lowercase agar tidak gagal karena spasi/kapital
          var skuRowMap = {};
          for (var i = 1; i < dataProd.length; i++) {
            var skuKey = String(dataProd[i][3]).trim().toLowerCase();
            if (skuKey) skuRowMap[skuKey] = i;
          }
          
          var restockCount = 0;
          var failedSkus = [];
          var histRows = [];
          
          payload.items.forEach(function(item) {
               var inputSku = String(item.sku).trim();
               var lookupKey = inputSku.toLowerCase();
               var idx = skuRowMap[lookupKey];
               
               if (idx === undefined) {
                   // SKU tidak ditemukan - catat untuk feedback
                   failedSkus.push(inputSku);
                   return;
               }
               
               var rowNum = idx + 1;
               
               var oldStock = Number(dataProd[idx][7]) || 0;  // H: Sisa
               var oldHpp = Number(dataProd[idx][9]) || 0;    // J: HPP
               var oldMasuk = (dataProd[idx].length > 20) ? (Number(dataProd[idx][20]) || 0) : 0;
               
               var qtyMasuk = Number(item.qty) || 0;
               if (qtyMasuk <= 0) return; // Skip jika qty invalid
               
               var priceBeli = Number(item.price) || 0;
               if (priceBeli === 0) priceBeli = oldHpp;
               
               var newStock = oldStock + qtyMasuk;
               var newMasukTotal = oldMasuk + qtyMasuk;
               
               // Average HPP
               var newHpp = oldHpp;
               if (newStock > 0) {
                   newHpp = ((oldStock * oldHpp) + (qtyMasuk * priceBeli)) / newStock;
               }
               
               // UPDATE SHEET
               sProd.getRange(rowNum, 8).setValue(newStock);                    // H: Sisa
               sProd.getRange(rowNum, 10, 1, 2).setValues([[Math.round(newHpp), Math.round(newStock * newHpp)]]); // J+K
               sProd.getRange(rowNum, 21).setValue(newMasukTotal);             // U: Masuk
               
               // Update cache lokal agar item berikutnya pakai data terbaru
               dataProd[idx][7] = newStock;
               dataProd[idx][9] = newHpp;
               dataProd[idx][20] = newMasukTotal;
               
               histRows.push(['TRX-'+Date.now()+'-'+restockCount, timestamp, 'IN', payload.ref, inputSku, dataProd[idx][1], dataProd[idx][4], qtyMasuk, priceBeli, qtyMasuk*priceBeli, newStock, 'Restock', user, timestamp]);
               restockCount++;
          });
          
          // BATCH append riwayat
          if (histRows.length > 0) {
            sHist.getRange(sHist.getLastRow() + 1, 1, histRows.length, 14).setValues(histRows);
          }
          
          // Feedback yang jelas ke user
          if (restockCount === 0 && failedSkus.length > 0) {
            return { success: false, message: "Gagal: SKU tidak ditemukan di Master Produk (" + failedSkus.join(', ') + ")" };
          }
          var msg = "Berhasil restock " + restockCount + " produk!";
          if (failedSkus.length > 0) {
            msg += " (SKU tidak ditemukan: " + failedSkus.join(', ') + ")";
          }
          return { success: true, message: msg };
      }
  }
  return { success: false, message: "Tipe transaksi salah." };
}

// --- FUNGSI SAVE IMAGE (SOLUSI GAMBAR RUSAK) ---
function saveImageToDrive(actionData, fileName) {
  try {
    // [FIX] Deteksi apakah input berupa Object atau String
    let base64String = actionData;

    // Jika input adalah Object (misal: { data: "...", type: "..." })
    if (typeof actionData === 'object' && actionData !== null) {
        if (actionData.data) base64String = actionData.data;
        else if (actionData.base64) base64String = actionData.base64;
        else if (actionData.content) base64String = actionData.content;
    }

    // Validasi akhir: Harus String
    if (typeof base64String !== 'string') {
       throw new Error("Format data gambar tidak valid (Bukan String Base64)");
    }

    // Lanjut proses split
    var data = base64String.split(",");
    var cleanBase64 = data.length > 1 ? data[1] : data[0];
    
    var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), "image/jpeg", fileName);
    var file = DriveApp.createFile(blob);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    
    // Return URL Thumbnail
    return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000"; 
    
  } catch (e) {
    Logger.log("Error saveImage: " + e.toString());
    // Return string kosong atau error agar tidak merusak data sheet
    return ""; 
  }
}

// --- DELETE PRODUCT (SECURED) ---
function deleteProductBySku(sku, requestor) {
  
  // 1. CEK IZIN DULU (Hanya Admin & Owner yang boleh hapus)
  if (!validateUserAccess(requestor, 'Admin')) {
     return { success: false, message: "AKSES DITOLAK! Hanya Admin/Owner yang boleh menghapus data." };
  }

  // 2. JIKA LOLOS, LANJUT PROSES ASLI
  const sheet = _getSS().getSheetByName('Master_Produk');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]) === String(sku)) {
      sheet.deleteRow(i + 1);
      
      // Log siapa yang menghapus
      const sLog = _getSS().getSheetByName('Data_Riwayat');
      if(sLog) sLog.appendRow([new Date(), 'DEL-'+sku, 'DELETE PRODUCT', sku, '-', 0, requestor, 'Menghapus Produk Master']);

      return { success: true, message: "Produk berhasil dihapus." };
    }
  }
  
  return { success: false, message: "SKU tidak ditemukan." };
}

// --- FITUR EDIT PRODUK ---

// 1. Ambil Data Spesifik per SKU
function getProductBySku(sku, varName) {
  const sheet = getSheet('Master_Produk');
  const data = sheet.getDataRange().getValues();
  
  // Normalisasi input (bisa SKU Variasi, SKU Induk, atau Barcode)
  const searchKey = String(sku).trim().toUpperCase();
  // Normalisasi varName untuk pencocokan spesifik
  const searchVar = varName ? String(varName).trim().toLowerCase() : '';
  
  // Strategi: Cari di SKU Variasi (D), Barcode (P), dan SKU Induk (A)
  let fallbackResult = null;
  
  for (let i = 1; i < data.length; i++) {
    const rowSku = String(data[i][3]).trim().toUpperCase();      // Kolom D (SKU Variasi)
    const rowParentSku = String(data[i][0]).trim().toUpperCase(); // Kolom A (SKU Induk)
    const rowBarcode = String(data[i][15] || "").trim().toUpperCase(); // Kolom P (Barcode)
    const rowVar = String(data[i][4]).trim().toLowerCase();      // Kolom E (Variasi)
    
    // Match: SKU Variasi, Barcode, atau SKU Induk
    const skuMatch = (rowSku === searchKey || (rowBarcode !== "" && rowBarcode === searchKey) || rowParentSku === searchKey);
    
    if (skuMatch) {
      const row = data[i];
      const result = {
        found: true,
        rowIndex: i + 1,
        
        parentSku: row[0],
        name: row[1],
        category: row[2],
        sku: row[3],
        varName: row[4],
        
        // Stok & HPP
        qty: row[7],      // Stok Saat Ini (Kolom H)
        unit: row[8],     // Satuan (Kolom I)
        hpp: row[9],      // HPP (Kolom J)
        sellPrice: row[11], // Harga Jual (Kolom L)
        asset: row[10],   // Nilai Aset (Kolom K)
        
        location: row[12], // Lokasi (Kolom M)
        mainImage: row[13],
        varImage: row[14],
        barcode: row[15] || '', // Barcode (Kolom P)
        adminPct: row[21] || 0,  // Admin Shopee (Kolom V)
        adminTiktok: row[22] || 0  // Admin Tiktok (Kolom W)
      };
      
      // Jika varName dikirim, cocokkan juga variasi
      if (searchVar) {
        if (rowVar === searchVar) {
          return result; // EXACT MATCH: SKU + Variasi → langsung return
        }
        // Simpan sebagai fallback jika tidak ada exact match
        if (!fallbackResult) fallbackResult = result;
      } else {
        // Tidak ada varName → return match pertama (logika lama)
        return result;
      }
    }
  }
  
  // Jika ada varName tapi tidak ketemu exact match, return fallback
  if (fallbackResult) return fallbackResult;
  
  return { found: false };
}

// 1b. Ambil SEMUA variasi dari 1 produk (by Parent SKU)
function getProductAllVariants(parentSku) {
  const sheet = getSheet('Master_Produk');
  const data = sheet.getDataRange().getValues();
  const searchKey = String(parentSku).trim().toUpperCase();
  
  var variants = [];
  var productInfo = null;
  
  for (var i = 1; i < data.length; i++) {
    var rowParent = String(data[i][0]).trim().toUpperCase();
    if (rowParent !== searchKey) continue;
    
    var row = data[i];
    if (!productInfo) {
      productInfo = {
        parentSku: row[0],
        name: row[1],
        category: row[2],
        unit: row[8],
        location: row[12],
        mainImage: row[13]
      };
    }
    
    variants.push({
      sku: row[3],
      varName: row[4],
      qty: row[7],
      hpp: row[9],
      sellPrice: row[11],
      varImage: row[14],
      barcode: row[15] || '',
      adminPct: row[21] || 0,
      adminTiktok: row[22] || 0
    });
  }
  
  if (!productInfo || variants.length === 0) return { found: false };
  
  productInfo.found = true;
  productInfo.variants = variants;
  return productInfo;
}

// 2. Simpan Perubahan (Update Row)
function updateProductData(form) {
  const sheet = getSheet('Master_Produk');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for(let i=1; i<data.length; i++) {
      if(String(data[i][3]) === String(form.originalSku)) {
          rowIndex = i + 1;
          break;
      }
  }
  
  if(rowIndex === -1) return { success: false, message: "Produk tidak ditemukan." };

  // Update Data Standar
  sheet.getRange(rowIndex, 2).setValue(form.name);
  sheet.getRange(rowIndex, 3).setValue(form.category);
  sheet.getRange(rowIndex, 5).setValue(form.varName);
  sheet.getRange(rowIndex, 9).setValue(form.unit);
  sheet.getRange(rowIndex, 12).setValue(form.sellPrice);
  sheet.getRange(rowIndex, 13).setValue(form.location);
  
  // UPDATE BARCODE (Kolom 16 / P)
  if (form.barcode !== undefined) {
      sheet.getRange(rowIndex, 16).setValue(form.barcode);
  }

   // UPDATE ADMIN SHOPEE (Kolom 22 / V)
  if (form.adminPct !== undefined) {
      sheet.getRange(rowIndex, 22).setValue(Number(form.adminPct) || 0);
  }

  // UPDATE ADMIN TIKTOK (Kolom 23 / W)
  if (form.adminTiktok !== undefined) {
      sheet.getRange(rowIndex, 23).setValue(Number(form.adminTiktok) || 0);
  }

  // Recalculate Asset
  const currentStock = sheet.getRange(rowIndex, 8).getValue();
  const currentHpp = sheet.getRange(rowIndex, 10).getValue();
  sheet.getRange(rowIndex, 11).setValue(currentStock * currentHpp);

  // Update Foto
  if (form.newMainImage) {
      // Panggil fungsi saveImageToDrive (sekarang sudah aman menerima object/string)
      const url = saveImageToDrive(form.newMainImage, form.parentSku + "_MAIN_UPD_" + Date.now() + ".jpg");
      if(url && url !== "") sheet.getRange(rowIndex, 14).setValue(url);
  }
  
  if (form.newVarImage) {
      const url = saveImageToDrive(form.newVarImage, form.sku + "_VAR_UPD_" + Date.now() + ".jpg");
      // [FIX] Hapus baris duplikat yang lama
      if(url && url !== "") sheet.getRange(rowIndex, 15).setValue(url);
  }

  return { success: true, message: "Data Produk Berhasil Diupdate!" };
}

// FUNGSI UNTUK HAPUS ORDER
function deleteOrderById(orderId) {
  const sheet = getSheet('Data_Pesanan');
  const data = sheet.getDataRange().getValues();
  
  // Kita cari semua baris yang punya ID tersebut (karena 1 order bisa multi-row)
  // Kita hapus dari bawah ke atas agar index tidak bergeser saat loop
  let deletedCount = 0;
  
  for (let i = data.length - 1; i >= 1; i--) {
    // Kolom A (Index 0) adalah ID Pesanan
    if (String(data[i][0]) === String(orderId)) {
      sheet.deleteRow(i + 1); // +1 karena sheet index mulai dari 1
      deletedCount++;
    }
  }

  if (deletedCount > 0) {
    return { success: true, message: `Order ${orderId} berhasil dihapus.` };
  } else {
    return { success: false, message: "ID Pesanan tidak ditemukan." };
  }
}

// 2. FUNGSI UPLOAD MASSAL (FIX: HARGA MASTER & STATUS BERSIH)
function saveMassData(items) {
  const lock = LockService.getScriptLock();
  try {
      lock.waitLock(30000); 
  } catch (e) {
      return { success: false, message: "Server sibuk. Coba lagi." };
  }

  try {
      const ss = _getSS();
      const sheetPesanan = ss.getSheetByName('Data_Pesanan');
      const sheetProduk = ss.getSheetByName('Master_Produk');
      const sLogActivity = ss.getSheetByName('Data_Riwayat');
      const user = Session.getActiveUser().getEmail();
      const timestamp = new Date();
      
      // 1. CACHE DATA MASTER PRODUK (Untuk Ambil Harga & HPP)
      const dataProduk = sheetProduk.getDataRange().getValues();
      const productMap = {}; // Map: SKU -> {hpp, sellPrice}

      for(let i=1; i<dataProduk.length; i++) {
          const sku = String(dataProduk[i][3]).trim(); // SKU Variasi
          const hpp = Number(dataProduk[i][9]) || 0;   // Kolom J (HPP)
          const sell = Number(dataProduk[i][11]) || 0; // Kolom L (Harga Jual)
          
          productMap[sku] = { hpp: hpp, sell: sell };
      }

      // Cache Data Pesanan (Cek Duplikat)
      const existingData = sheetPesanan.getDataRange().getValues();
      const dbIds = new Set();
      for(let i=1; i<existingData.length; i++) {
          dbIds.add(String(existingData[i][0]).trim());
      }

      const newRows = [];
      let duplicateCount = 0;
      let successCount = 0;
      
      items.forEach(item => {
          const orderId = String(item.id).trim();
          if(dbIds.has(orderId)) { duplicateCount++; return; }

          // --- AMBIL DATA DARI MASTER PRODUK ---
          // Prioritas Cek: SKU Variasi -> SKU Induk
          const skuCheck = String(item.skuVariasi || item.skuInduk).trim();
          let masterHpp = 0;
          let masterSell = 0;

          if (productMap.hasOwnProperty(skuCheck)) {
              masterHpp = productMap[skuCheck].hpp;
              masterSell = productMap[skuCheck].sell;
          }

          // --- PERHITUNGAN KEUANGAN (SESUAI REQUEST) ---
          const qty = Number(item.qty) || 1; // Dari Kolom AH File
          
          // Kolom K: Harga Jual (@) -> Dari Master
          const unitSellPrice = masterSell; 
          
          // Kolom L: Total Penjualan -> Harga Master * Qty
          const totalSales = unitSellPrice * qty; 
          
          // Kolom M: Harga Beli (@) -> Dari Master
          const unitHpp = masterHpp; 
          
          // Kolom N: Total Modal -> HPP Master * Qty
          const totalModal = unitHpp * qty; 

          // Kolom T: Total Pembayaran -> Dari File Kolom AT
          const totalPaymentFile = Number(item.harga) || 0; 

          // --- STATUS LOGISTIK ---
          // Jika status berisi "WAITING" atau "PERLU DIKIRIM", paksa jadi "WAITING" bersih.
          let finalStatus = item.status;
          if (finalStatus && (finalStatus.includes('WAITING') || finalStatus.includes('PERLU'))) {
              finalStatus = 'WAITING'; // Hapus embel-embel (PRE-ORDER) / (UNKNOWN)
          }

          // --- SUSUN BARIS DATA (ARRAY 28 KOLOM) ---
          let row = new Array(28).fill('');

          row[0] = orderId;           // A: ID
          row[1] = item.tgl;          // B: Tanggal
          row[2] = item.channel;      // C: Channel
          row[3] = item.kurir;        // D: Kurir
          row[4] = item.resi;         // E: Resi

          row[5] = item.skuInduk;     // F: SKU Induk (File Kolom T)
          row[6] = item.nama;         // G: Nama Produk (File Kolom U)
          row[7] = item.skuVariasi;   // H: SKU Variasi (File Kolom V)
          row[8] = item.varian;       // I: Nama Variasi (File Kolom W)

          row[9] = qty;               // J: Jumlah (File Kolom AH)
          
          row[10] = unitSellPrice;    // K: Harga Jual @ (Master)
          row[11] = totalSales;       // L: Total Penjualan (Master * Qty)
          row[12] = unitHpp;          // M: Harga Beli @ (Master)
          row[13] = totalModal;       // N: Total Modal (Master * Qty)

          row[14] = finalStatus;      // O: Status (WAITING Bersih)

          // P-S Kosong
          row[18] = 0;                // S: Diskon (0 dulu atau sesuaikan)
          
          row[19] = totalPaymentFile; // T: Total Pembayaran (File Kolom AT)

          // --- DATA CUSTOMER (V-AB) ---
          row[21] = item.custName;    // V
          row[22] = item.custPhone;   // W
          row[23] = item.custAddress; // X
          row[24] = item.ongkir;      // Y
          row[25] = "";               // Z (Kecamatan Kosong)
          row[26] = item.kota;        // AA
          row[27] = item.provinsi;    // AB

          newRows.push(row);
          successCount++;
      });

      if (newRows.length > 0) {
          const lastRow = sheetPesanan.getLastRow();
          sheetPesanan.getRange(lastRow + 1, 1, newRows.length, 28).setValues(newRows);
      }

      if (sLogActivity && successCount > 0) {
          sLogActivity.appendRow([
              timestamp, 'BULK-' + Date.now(), 'MASS UPLOAD', '-', 
              'Import Excel', successCount, user, 
              `Berhasil Upload: ${successCount}. Duplikat: ${duplicateCount}`
          ]);
      }

      let msg = `Berhasil menyimpan ${successCount} data.`;
      if (duplicateCount > 0) msg += ` (${duplicateCount} duplikat).`;
      
      SpreadsheetApp.flush();
      return { success: successCount > 0, message: msg };

  } catch (e) {
      return { success: false, message: "Error Bulk Upload: " + e.toString() };
  } finally {
      lock.releaseLock();
  }
}

//TAMBAHAN FITUR BULK DELETE

function deleteBulkOrders(orderIds) {
  const sheet = getSheet('Data_Pesanan');
  const data = sheet.getDataRange().getValues();
  let deletedCount = 0;
  
  // Ubah array ID menjadi Set untuk pencarian cepat
  const idsToDelete = new Set(orderIds.map(String));

  // Loop dari bawah ke atas agar index tidak bergeser saat deleteRow
  for (let i = data.length - 1; i >= 1; i--) {
    const currentId = String(data[i][0]); // Kolom A = ID
    
    if (idsToDelete.has(currentId)) {
      sheet.deleteRow(i + 1); // +1 karena index sheet mulai dari 1
      deletedCount++;
    }
  }

  if (deletedCount > 0) {
    return { success: true, message: `Berhasil menghapus ${deletedCount} pesanan.` };
  } else {
    return { success: false, message: "Tidak ada data yang dihapus." };
  }
}

// [PERBAIKAN] LOGIKA OPNAME UNIVERSAL (PRODUK & BAHAN)
function processStockOpname(adjustments, opnameType) {
  // opnameType bisa 'PRODUCT' atau 'MATERIAL'
  if (!adjustments || adjustments.length === 0) return { success: false, message: "Tidak ada data." };
  
  const ss = _getSS();
  const sHist = ss.getSheetByName('Riwayat_Stok');
  const sLogActivity = ss.getSheetByName('Data_Riwayat');
  
  // 1. TENTUKAN TARGET SHEET & KOLOM BERDASARKAN TIPE
  let sheetMaster, colStok, colHpp, colAset, colDate;
  
  // Default ke PRODUCT jika tidak ada tipe
  const type = opnameType || 'PRODUCT'; 

  if (type === 'MATERIAL') {
      sheetMaster = ss.getSheetByName('Master_Bahan');
      // Mapping Kolom Master Bahan (Index mulai 0):
      // H(7)=Stok, I(8)=HPP, J(9)=Aset
      colStok = 7; 
      colHpp  = 8;
      colAset = 9;
      
      // [PERBAIKAN] Sesuai koreksi Anda: Kolom M = Index 12
      colDate = 12; 
  } else {
      sheetMaster = ss.getSheetByName('Master_Produk');
      // Mapping Kolom Master Produk:
      // H(7)=Stok, J(9)=HPP, K(10)=Aset, T(19)=Tgl Opname
      colStok = 7;
      colHpp  = 9;
      colAset = 10;
      colDate = 19; 
  }

  if (!sheetMaster) return { success: false, message: `Sheet Master untuk ${type} tidak ditemukan.` };

  const masterData = sheetMaster.getDataRange().getValues();
  const timestamp = new Date();
  const refCode = 'SO-' + (type === 'MATERIAL' ? 'MAT-' : 'PRD-') + Utilities.formatDate(timestamp, "Asia/Jakarta", "ddMMyy-HHmm");
  const user = 'Admin Opname'; 

  let countTotal = 0;
  const colSkuIndex = (type === 'MATERIAL') ? 0 : 3;
  const dateString = Utilities.formatDate(timestamp, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  const historyRows = [];
  const logRows = [];

  // Cek kolom material sekali saja di awal
  if (type === 'MATERIAL' && sheetMaster.getMaxColumns() < 13) {
    sheetMaster.insertColumnsAfter(sheetMaster.getMaxColumns(), 13 - sheetMaster.getMaxColumns());
  }

  adjustments.forEach(item => {
      const targetSku = String(item.sku).trim().toLowerCase();
      const diff = Number(item.diff);
      const newStock = Number(item.newStock);
      
      let productName = "";
      let productVar = "";
      let currentHpp = 0;

      for (let i = 1; i < masterData.length; i++) {
          if (String(masterData[i][colSkuIndex]).trim().toLowerCase() === targetSku) {
              
              productName = masterData[i][1]; 
              productVar  = (type === 'MATERIAL') ? '-' : masterData[i][4];
              currentHpp  = Number(masterData[i][colHpp]) || 0;

              // Update Stok & Aset
              sheetMaster.getRange(i + 1, colStok + 1).setValue(newStock);
              sheetMaster.getRange(i + 1, colAset + 1).setValue(newStock * currentHpp);
              
              try {
                 if (type === 'MATERIAL') {
                     sheetMaster.getRange(i + 1, 12, 1, 2).setValues([[item.note || '-', "'" + dateString]]);
                 } else {
                     sheetMaster.getRange(i + 1, 20).setValue("'" + dateString);
                 }
              } catch(e) {
                 Logger.log("Error Update Info: " + e);
              }
              break; 
          }
      }

      let logType = "OPNAME (OK)";
      let note = item.note || "Sesuai";
      if (diff > 0) { logType = "IN (ADJUST)"; note = item.note || "Selisih Lebih"; }
      else if (diff < 0) { logType = "OUT (ADJUST)"; note = item.note || "Selisih Kurang"; }

      const diffAbs = Math.abs(diff);

      historyRows.push([
          'ADJ-' + Math.floor(Math.random()*100000), timestamp, logType, refCode,
          item.sku, productName || 'Unknown', productVar || '-', diffAbs,
          currentHpp, diffAbs * currentHpp, newStock, note + ` (${type})`, user, timestamp
      ]);
      logRows.push([timestamp, refCode, logType, item.sku, productName, diffAbs, user, note + ` (${type})`]);
      countTotal++;
  });

  // BATCH APPEND riwayat + log
  if (historyRows.length > 0) {
    sHist.getRange(sHist.getLastRow() + 1, 1, historyRows.length, 14).setValues(historyRows);
  }
  if (logRows.length > 0 && sLogActivity) {
    sLogActivity.getRange(sLogActivity.getLastRow() + 1, 1, logRows.length, 8).setValues(logRows);
  }

  return { success: true, message: `Berhasil! ${countTotal} item ${type} telah divalidasi.` };
}

// [Code.gs] - MESIN PRODUKSI (AUTO HPP & STOK) - VERSI UPDATE OPERASIONAL

function ensureSheetWithHeader_(sheetName, headers) {
  const ss = _getSS();
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }

  if (headers && headers.length > 0 && sh.getLastRow() === 0) {
    sh.appendRow(headers);
  }

  return sh;
}

function safeJsonParse_(raw, fallback) {
  try {
    if (!raw) return fallback;
    return JSON.parse(raw);
  } catch (e) {
    return fallback;
  }
}

function runProductionToInventory_(form) {
  const ss = _getSS();
  const sProd = ss.getSheetByName('Master_Produk');
  const sMat = ss.getSheetByName('Master_Bahan');
  const sHist = ss.getSheetByName('Riwayat_Stok');
  const sLogProd = ensureSheetWithHeader_('Data_Produksi', ["ID Produksi", "Tanggal", "SKU", "Nama Produk", "Qty", "HPP Satuan", "Catatan", "User"]);
  const sLogActivity = ss.getSheetByName('Data_Riwayat');

  if (!sProd || !sMat || !sHist) {
    return { success: false, message: 'Sheet Master_Produk / Master_Bahan / Riwayat_Stok tidak lengkap.' };
  }

  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  const idProduksi = 'PRD-' + Math.floor(Date.now() / 1000);

  const dataProd = sProd.getDataRange().getValues();
  const dataMat = sMat.getDataRange().getValues();

  // 1. Cari produk output
  let rowOutput = -1;
  for (let i = 1; i < dataProd.length; i++) {
    if (String(dataProd[i][3]).trim() === String(form.outputSku).trim()) {
      rowOutput = i + 1;
      break;
    }
  }
  if (rowOutput === -1) return { success: false, message: 'Produk Output tidak ditemukan di Master.' };

  // 2. Validasi bahan
  let totalMaterialCost = 0;
  const matsToUpdate = [];
  const materials = Array.isArray(form.materials) ? form.materials : [];

  for (let m of materials) {
    const matSku = String(m.sku || '').trim();
    const matQtyNeed = Number(m.qty) || 0;
    let foundMatIndex = -1;

    for (let j = 1; j < dataMat.length; j++) {
      if (String(dataMat[j][0]).trim() === matSku) {
        foundMatIndex = j;
        break;
      }
    }

    if (foundMatIndex === -1) return { success: false, message: `Bahan '${m.sku}' tidak ditemukan.` };

    const currentStock = Number(dataMat[foundMatIndex][7]) || 0; // H: Sisa
    const currentHpp = Number(dataMat[foundMatIndex][8]) || 0;   // I: HPP

    if (currentStock < matQtyNeed) {
      return { success: false, message: `Stok Bahan Kurang: ${dataMat[foundMatIndex][1]} (Sisa: ${currentStock})` };
    }

    totalMaterialCost += (matQtyNeed * currentHpp);
    matsToUpdate.push({
      rowIndex: foundMatIndex + 1,
      qtyPakai: matQtyNeed,
      hpp: currentHpp,
      sku: dataMat[foundMatIndex][0],
      name: dataMat[foundMatIndex][1]
    });
  }

  // 3. Biaya operasional
  let totalOpCost = 0;
  let opLogText = "";
  const operationalCosts = Array.isArray(form.operationalCosts) ? form.operationalCosts : [];

  operationalCosts.forEach(op => {
    totalOpCost += Number(op.cost) || 0;
    if (op && op.name) opLogText += `${op.name}, `;
  });

  // 4. Potong stok bahan (gunakan data dari cache dataMat, bukan baca ulang)
  const prodHistRows = [];
  matsToUpdate.forEach(item => {
    const row = item.rowIndex;
    const idx = row - 1; // index array
    const valSisaLama = Number(dataMat[idx][7]) || 0;
    const valKeluarLama = Number(dataMat[idx][6]) || 0;

    const valKeluarBaru = valKeluarLama + item.qtyPakai;
    const valSisaBaru = valSisaLama - item.qtyPakai;

    // BATCH: G+H sekaligus, lalu J (aset)
    sMat.getRange(row, 7, 1, 2).setValues([[valKeluarBaru, valSisaBaru]]);
    sMat.getRange(row, 10).setValue(valSisaBaru * item.hpp);

    prodHistRows.push([
      idProduksi, timestamp, 'KELUAR (PRODUKSI)', '-', item.sku, item.name, '-',
      item.qtyPakai, item.hpp, item.qtyPakai * item.hpp, valSisaBaru,
      `Bahan utk: ${form.outputName}`, user, timestamp
    ]);
  });
  // BATCH APPEND riwayat
  if (prodHistRows.length > 0) {
    sHist.getRange(sHist.getLastRow() + 1, 1, prodHistRows.length, 14).setValues(prodHistRows);
  }

  // 5. Tambah stok barang jadi
  const totalProductionCost = totalMaterialCost + totalOpCost;

  const oldAwalOutput = Number(sProd.getRange(rowOutput, 6).getValue()) || 0;
  const oldMasukOutput = Number(sProd.getRange(rowOutput, 21).getValue()) || 0;
  const oldStockOutput = Number(sProd.getRange(rowOutput, 8).getValue()) || 0;
  const oldTotalAsset = Number(sProd.getRange(rowOutput, 11).getValue()) || 0;

  const qtyOutput = Number(form.outputQty) || 0;
  const newStockOutput = oldStockOutput + qtyOutput;
  const newTotalAsset = oldTotalAsset + totalProductionCost;
  const newHppOutput = newStockOutput > 0 ? Math.round(newTotalAsset / newStockOutput) : 0;

  if (oldAwalOutput === 0 && oldMasukOutput === 0) {
    sProd.getRange(rowOutput, 6).setValue(qtyOutput);
  } else {
    const newMasukOutput = oldMasukOutput + qtyOutput;
    sProd.getRange(rowOutput, 21).setValue(newMasukOutput);
  }

  sProd.getRange(rowOutput, 8).setValue(newStockOutput);
  sProd.getRange(rowOutput, 10).setValue(newHppOutput);
  sProd.getRange(rowOutput, 11).setValue(newTotalAsset);

  const note = opLogText ? `Produksi (Ops: ${opLogText})` : 'Produksi';

  sHist.appendRow([
    idProduksi, timestamp, 'IN (PRODUKSI)', '-', form.outputSku, form.outputName, '-',
    qtyOutput, newHppOutput, totalProductionCost, newStockOutput,
    note, user, timestamp
  ]);

  sLogProd.appendRow([
    idProduksi,
    timestamp,
    form.outputSku,
    form.outputName,
    qtyOutput,
    newHppOutput,
    note,
    user
  ]);

  if (sLogActivity) {
    sLogActivity.appendRow([
      timestamp,
      idProduksi,
      'PRODUKSI',
      form.outputSku,
      form.outputName,
      qtyOutput,
      user,
      note
    ]);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    message: `Produksi Berhasil. HPP Baru: Rp ${newHppOutput}`,
    hpp: newHppOutput,
    productionId: idProduksi
  };
}

// Backward compatibility: tetap bisa dipanggil langsung seperti flow lama
function handleProduction(form) {
  return runProductionToInventory_(form);
}

// SIMPAN DULU KE DATA PROSES JAHIT (BELUM POSTING STOK)
function queueProductionProcess(form) {
  const ss = _getSS();
  const sQueue = ensureSheetWithHeader_('Data_Proses_Jahit', [
    'ID Proses', 'Tanggal Input', 'SKU Output', 'Nama Output', 'Qty Output', 'Nama Penjahit',
    'JSON Bahan', 'JSON Operasional', 'Status', 'User Input', 'Tanggal Selesai',
    'Ringkasan', 'Catatan'
  ]);
  const sLogActivity = ss.getSheetByName('Data_Riwayat');

  const outputSku = String(form.outputSku || '').trim();
  const outputName = String(form.outputName || '').trim();
  const outputQty = Number(form.outputQty) || 0;
  const tailorName = String(form.tailorName || '').trim();
  const materials = Array.isArray(form.materials) ? form.materials : [];
  const operationalCosts = Array.isArray(form.operationalCosts) ? form.operationalCosts : [];

  if (!outputSku || !outputName || outputQty <= 0) {
    return { success: false, message: 'Data output produksi belum lengkap.' };
  }
  if (!tailorName) {
    return { success: false, message: 'Nama penjahit wajib diisi.' };
  }
  if (materials.length === 0) {
    return { success: false, message: 'Minimal harus ada 1 bahan baku.' };
  }

  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  const idProses = 'JHT-' + Math.floor(Date.now() / 1000) + '-' + Math.floor(Math.random() * 90 + 10);

  const summary = materials
    .slice(0, 5)
    .map(m => `${m.name || m.sku} (${m.qty})`)
    .join(', ');

  sQueue.appendRow([
    idProses,
    timestamp,
    outputSku,
    outputName,
    outputQty,
    tailorName,
    JSON.stringify(materials),
    JSON.stringify(operationalCosts),
    'PROSES',
    user,
    '',
    summary,
    'Menunggu tombol Selesai'
  ]);

  if (sLogActivity) {
    sLogActivity.appendRow([
      timestamp,
      idProses,
      'PROSES JAHIT',
      outputSku,
      outputName,
      outputQty,
      user,
      'Antrian produksi dibuat'
    ]);
  }

  return { success: true, message: `Data masuk ke Proses Jahit (${idProses}).`, id: idProses };
}

// LIST DATA PROSES JAHIT UNTUK FRONTEND
function getProductionProcessList() {
  const sheet = getSheet('Data_Proses_Jahit');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const limit = 100;
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;
  const colCount = Math.max(sheet.getLastColumn(), 13);
  const data = sheet.getRange(startRow, 1, numRows, colCount).getDisplayValues();

  return data.reverse().map(row => ({
    // Kompatibel dengan data lama (12 kolom) dan baru (13 kolom + penjahit)
    id: row[0],
    createdAt: row[1],
    outputSku: row[2],
    outputName: row[3],
    outputQty: row[4],
    tailorName: row.length >= 13 ? (row[5] || '-') : '-',
    status: row.length >= 13 ? (row[8] || 'PROSES') : (row[7] || 'PROSES'),
    createdBy: row.length >= 13 ? (row[9] || '-') : (row[8] || '-'),
    finishedAt: row.length >= 13 ? (row[10] || '') : (row[9] || ''),
    notes: row.length >= 13 ? (row[12] || '') : (row[11] || '')
  }));
}

// AKSI SELESAI: BARU POSTING KE MASTER PRODUK & STOK
function completeProductionProcess(processId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, message: 'Sistem sibuk. Coba lagi sesaat.' };
  }

  try {
    const ss = _getSS();
    const sheet = ss.getSheetByName('Data_Proses_Jahit');
    const sLogActivity = ss.getSheetByName('Data_Riwayat');

    if (!sheet) return { success: false, message: 'Sheet Data_Proses_Jahit tidak ditemukan.' };

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let rowData = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(processId).trim()) {
        rowIndex = i + 1;
        rowData = data[i];
        break;
      }
    }

    if (rowIndex === -1 || !rowData) {
      return { success: false, message: 'Data proses jahit tidak ditemukan.' };
    }

    const isNewFormat = rowData.length >= 13;
    const statusIndex = isNewFormat ? 8 : 7;
    const status = String(rowData[statusIndex] || '').toUpperCase();

    if (status === 'SELESAI') {
      return { success: false, message: 'Proses ini sudah selesai sebelumnya.' };
    }

    const tailorName = isNewFormat ? String(rowData[5] || '').trim() : '';
    const materialsRaw = isNewFormat ? rowData[6] : rowData[5];
    const opsRaw = isNewFormat ? rowData[7] : rowData[6];

    const form = {
      outputSku: rowData[2],
      outputName: rowData[3],
      outputQty: Number(rowData[4]) || 0,
      materials: safeJsonParse_(materialsRaw, []),
      operationalCosts: safeJsonParse_(opsRaw, [])
    };

    const result = runProductionToInventory_(form);
    if (!result.success) return result;

    const user = Session.getActiveUser().getEmail();
    const now = new Date();

    const statusCol = isNewFormat ? 9 : 8;
    const finishCol = isNewFormat ? 11 : 10;
    const noteCol = isNewFormat ? 13 : 12;

    sheet.getRange(rowIndex, statusCol).setValue('SELESAI');
    sheet.getRange(rowIndex, finishCol).setValue(now);
    sheet.getRange(rowIndex, noteCol).setValue(`Diposting ke stok oleh ${user} | HPP: Rp ${result.hpp}${tailorName ? ' | Penjahit: ' + tailorName : ''}`);

    if (sLogActivity) {
      sLogActivity.appendRow([
        now,
        processId,
        'SELESAI PROSES JAHIT',
        form.outputSku,
        form.outputName,
        form.outputQty,
        user,
        `Posting stok berhasil | Ref Produksi: ${result.productionId}${tailorName ? ' | Penjahit: ' + tailorName : ''}`
      ]);
    }

    SpreadsheetApp.flush();
    return { success: true, message: `Proses selesai. Data masuk ke Master Produk & Stok.` };
  } catch (err) {
    return { success: false, message: 'Error: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

// TAMBAHAN UNTUK MEMBACA RIWAYAT PRODUKSI
function getProductionLog() {
  const sheet = getSheet('Data_Produksi');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const limit = 50;
  let startRow = Math.max(2, lastRow - limit + 1);
  let numRows = lastRow - startRow + 1;

  const data = sheet.getRange(startRow, 1, numRows, 8).getDisplayValues();
  return data.reverse();
}

// --- MANAJEMEN RESEP (BOM) ---

// 1. SIMPAN RESEP BARU
function saveRecipe(form) {
  const ss = _getSS();
  const sheet = ss.getSheetByName('Master_Produk');
  
  // [TAMBAHAN] Definisi Log Activity
  const sLogActivity = ss.getSheetByName('Data_Riwayat');
  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  
  const data = sheet.getDataRange().getValues();
  const parentSku = String(form.parentSku).trim();
  
  let foundIndex = -1;
  let currentJson = {};
  
  // [PENTING] Deklarasikan variabel ini agar tidak error "ReferenceError"
  let productName = ""; 

  // 1. Cari Baris Produk
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === parentSku) { 
      foundIndex = i + 1;
      productName = data[i][1]; // Ambil nama produk dari Kolom B
      
      const rawJson = data[i][18];
      if(rawJson && rawJson !== "") {
        try { currentJson = JSON.parse(rawJson); } catch(e) {}
      }
      break;
    }
  }

  if (foundIndex === -1) return { success: false, message: "Produk tidak ditemukan di Master." };

  // 2. UPDATE STRUKTUR JSON
  currentJson.mats = form.ingredients.map(item => ({
      sku: item.sku,
      name: item.name,
      qty: Number(item.qty)
  }));
  currentJson.costs = form.costs.map(item => ({
      name: item.name,
      cost: Number(item.cost)
  }));
  const newJsonString = JSON.stringify(currentJson);
  sheet.getRange(foundIndex, 19).setValue(newJsonString); 

  // 3. UPDATE NOMINAL HPP
  if (form.newHpp && Number(form.newHpp) > 0) {
      sheet.getRange(foundIndex, 10).setValue(Number(form.newHpp));
  }

  // [LOG KE DASHBOARD]
  if (sLogActivity) {
      // Pastikan productName ada isinya, jika kosong pakai parentSku
      const logName = productName || parentSku || 'Unknown Product';
      
      sLogActivity.appendRow([
          timestamp,
          'RCP-UPD-' + Date.now(),
          'UPDATE RESEP',
          parentSku,
          logName, // Sekarang variabel ini sudah aman
          0,
          user,
          'Update Komposisi/Resep'
      ]);
  }

  return { success: true, message: "Resep & HPP berhasil diperbarui!" };
}

// 2. AMBIL RESEP (AUTO-FILL)
function getRecipeBySku(parentSku) {
  const sheet = _getSS().getSheetByName('Master_Resep');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const ingredients = [];
  
  // Filter baris yang SKU Induknya cocok
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(parentSku).trim()) {
      ingredients.push({
        sku: data[i][2], // SKU Bahan
        name: data[i][3], // Nama Bahan
        qty: data[i][4]   // Qty Standar
      });
    }
  }
  
  return ingredients;
}

// [Code.gs] - SIMPAN MULTI-VARIAN (V2)

function saveProductDesignV2(payload) {
  const ss = _getSS();
  const sheet = ss.getSheetByName('Master_Produk');
  const sLogActivity = ss.getSheetByName('Data_Riwayat');
  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  
  if (!sheet) return { success: false, message: "Sheet Master_Produk tidak ditemukan" };

  const header = payload.header;   // Data Induk (Nama Produk, Kategori, dll)
  const variants = payload.variants; // Array Variasi (Merah, XL, dll)

  // Siapkan JSON Resep
  const recipeObj = {
    mats: header.materials.map(m => ({ sku: m.sku, name: m.name, qty: Number(m.qty) })),
    costs: header.costs.map(c => ({ name: c.name, cost: Number(c.cost) }))
  };
  const jsonRecipe = JSON.stringify(recipeObj);

  // Loop setiap variasi
  variants.forEach(variant => {
      
      // 1. Tentukan Nama Variasi Murni (Untuk Kolom E)
      // Data dari frontend (variant.name) isinya hanya tag variasi, misal: "Merah - XL" atau "Single"
      let varNameOnly = variant.name; 
      if(!varNameOnly || varNameOnly === '') varNameOnly = 'Single';

      // 2. Tentukan Nama Lengkap Produk (Untuk Kolom B)
      // Logic: Jika Single -> "Kemeja Polos"
      //        Jika Varian -> "Kemeja Polos (Merah - XL)" atau "Kemeja Polos - Merah - XL"
      let fullName = header.name;
      if (varNameOnly !== 'Single') {
          fullName = `${header.name} (${varNameOnly})`; 
      }

      // URUTAN KOLOM MASTER_PRODUK (A - S):
      const rowData = [
          header.sku,                 // A - Parent SKU
          fullName,                   // B - [FIX] Nama Lengkap (Produk + Varian)
          header.category,            // C - Kategori
          variant.sku,                // D - SKU Variasi
          varNameOnly,                // E - Nama Variasi Saja
          
          0,                          // F - Stok Awal
          0,                          // G - Masuk
          0,                          // H - Keluar
          
          "Pcs",                      // I - Satuan (Default)
          Number(variant.hpp),        // J - HPP (Harga Pokok)
          
          0,                          // K - Total Aset (0 dulu)
          
          Number(variant.price),      // L - Harga Jual
          
          header.location || '-',     // M - Lokasi Rak
          
          payload.header.newMainImage || "", // N - Foto Utama
          payload.header.newVarImage || "",  // O - Foto Variasi
          variant.barcode || "",             // P - Barcode
          
          "JADI",                     // Q - Tipe Barang
          Number(variant.margin),     // R - Margin (%)
          
          jsonRecipe                  // S - Resep Lengkap (JSON)
      ];

      sheet.appendRow(rowData);
  });

  if (sLogActivity) {
      sLogActivity.appendRow([
          timestamp,
          'PRD-NEW-' + Date.now(), // ID
          'NEW PRODUCT',           // Tipe
          header.sku,              // SKU Induk
          header.name,             // Nama
          variants.length,         // Qty (Jumlah Varian)
          user,                    // User
          'Desain Produk Baru'     // Ket
      ]);
  }

  return { success: true, message: "Produk berhasil disimpan dengan nama yang benar!" };
}

// [Code.gs] - INPUT MASTER BAHAN BAKU (RAW MATERIAL)

function saveNewMaterial(form) {
  const ss = _getSS();
  const sheet = ss.getSheetByName('Master_Produk');
  const sLogActivity = ss.getSheetByName('Data_Riwayat');
  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  
  const sku = String(form.sku).trim().toUpperCase();
  const data = sheet.getDataRange().getValues();
  const isExist = data.some(row => String(row[3]).trim() === sku);
  
  if(isExist) return { success: false, message: `SKU ${sku} sudah terdaftar!` };
  
  sheet.appendRow([
    sku, form.name, form.category, sku, '-',
    0, 0, 0, 0, form.hpp, 0, 0, form.rack,
    'Raw Material', '', '', 'BAHAN', 0
  ]);

  if (sLogActivity) {
      sLogActivity.appendRow([
          timestamp, 'MAT-NEW-' + Date.now(), 'NEW ITEM',
          sku, form.name, 0, user, 'Input Bahan Baru (Manual)'
      ]);
  }
  
  return { success: true, message: "Bahan baku berhasil didaftarkan!" };
}

// [Code.gs] - UPDATE SAVE BULK (DUAL FUNGSI: NEW & RESTOCK)

function saveBulkMaterials(payload) {
  const ss = _getSS();
  const sMat = ss.getSheetByName('Master_Bahan');
  const sHist = ss.getSheetByName('Riwayat_Stok'); 
  const sLogActivity = ss.getSheetByName('Data_Riwayat');
  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  
  if (!sMat) return { success: false, message: "Sheet Master_Bahan tidak ditemukan." };

  const data = sMat.getDataRange().getValues();
  
  const skuMap = new Map();
  for(let i=1; i<data.length; i++) {
      const sku = String(data[i][0]).trim().toUpperCase();
      skuMap.set(sku, {
          rowIndex: i + 1,
          stok: Number(data[i][7]), 
          hpp: Number(data[i][8]),  
          masuk: Number(data[i][5]), 
          awal: Number(data[i][4])   
      });
  }
  
  const newRows = [];
  const historyRows = [];
  const logRows = [];
  let updatedCount = 0;
  let newCount = 0;

  payload.items.forEach(item => {
    const sku = String(item.sku).trim().toUpperCase();
    const qtyMasuk = Number(item.qty);
    const hargaBeli = Number(item.hpp);
    const note = item.invoice ? `Inv: ${item.invoice}` : 'Restock / New';

    if (skuMap.has(sku)) {
        const existing = skuMap.get(sku);
        const row = existing.rowIndex;
        const stokLama = existing.stok;
        const stokBaru = stokLama + qtyMasuk;
        const masukBaru = existing.masuk + qtyMasuk; 
        
        let newHpp = existing.hpp;
        if (stokBaru > 0) {
            newHpp = ((stokLama * existing.hpp) + (qtyMasuk * hargaBeli)) / stokBaru;
        }

        // BATCH: update 4 cell -> 2 panggilan (F+H bersamaan tidak bisa karena tidak bersebelahan)
        sMat.getRange(row, 6).setValue(masukBaru);
        sMat.getRange(row, 8, 1, 3).setValues([[stokBaru, newHpp, stokBaru * newHpp]]);
        
        historyRows.push([
            'MAT-IN-' + Math.floor(Date.now()/1000) + '-' + updatedCount, 
            timestamp, 'IN (BAHAN)', item.invoice || '-', sku, item.name, '-', 
            qtyMasuk, hargaBeli, qtyMasuk * hargaBeli, stokBaru, 
            'Restock Bahan', user, timestamp
        ]);
        logRows.push([timestamp, item.invoice || 'MAT-IN', 'IN (RESTOCK)', sku, item.name, qtyMasuk, user, 'Restock Bahan Baku']);
        updatedCount++;
    } else {
        const totalAset = qtyMasuk * hargaBeli;
        newRows.push([
            sku, item.name, item.category, item.unit,      
            qtyMasuk, 0, 0, qtyMasuk,       
            hargaBeli, totalAset, item.rack, note            
        ]);
        historyRows.push([
            'MAT-NEW-' + Math.floor(Date.now()/1000) + '-' + newCount, 
            timestamp, 'NEW ITEM', item.invoice || '-', sku, item.name, '-', 
            qtyMasuk, hargaBeli, totalAset, qtyMasuk, 
            'Stok Awal Baru', user, timestamp
        ]);
        logRows.push([timestamp, 'MAT-NEW', 'NEW ITEM', sku, item.name, qtyMasuk, user, 'Input Bahan Baru (Bulk)']);
        newCount++;
    }
  });

  // BATCH INSERT bahan baru
  if (newRows.length > 0) {
    sMat.getRange(sMat.getLastRow() + 1, 1, newRows.length, 12).setValues(newRows);
  }
  // BATCH INSERT riwayat stok
  if (historyRows.length > 0 && sHist) {
    sHist.getRange(sHist.getLastRow() + 1, 1, historyRows.length, 14).setValues(historyRows);
  }
  // BATCH INSERT log activity
  if (logRows.length > 0 && sLogActivity) {
    sLogActivity.getRange(sLogActivity.getLastRow() + 1, 1, logRows.length, 8).setValues(logRows);
  }

  return { success: true, message: `Selesai! ${newCount} Bahan Baru, ${updatedCount} Restock.` };
}

// 2. UPDATE GET LIST (Gunakan getDataRange agar aman dari error range)
function getMaterialList() {
  const sheet = _getSS().getSheetByName('Master_Bahan');
  if (!sheet) return [];
  
  // Ambil semua data yang ada (Aman meskipun kolom M belum ada)
  const data = sheet.getDataRange().getValues();
  
  // Hapus header baris pertama
  if (data.length > 0) data.shift();
  
  return data.filter(row => row[0] !== "" && row[1] !== "");
}

// [Code.gs] - FITUR EDIT & HAPUS BAHAN BAKU

// 1. HAPUS BAHAN
function deleteMaterialItem(sku) {
  const ss = _getSS();
  const sheet = ss.getSheetByName('Master_Bahan');
  
  const data = sheet.getDataRange().getValues();
  // Loop cari SKU (Kolom A / Index 0)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(sku).trim()) {
      sheet.deleteRow(i + 1); // Hapus Baris
      return { success: true, message: `Bahan ${sku} berhasil dihapus.` };
    }
  }
  return { success: false, message: `SKU ${sku} tidak ditemukan.` };
}

// 2. UPDATE BAHAN (EDIT)
function updateMaterialItem(form) {
  const ss = _getSS();
  const sheet = ss.getSheetByName('Master_Bahan');
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(form.sku).trim()) {
      // Update Kolom Tertentu Saja (Agar stok tidak rusak)
      // Kolom B(2): Nama, C(3): Kategori, D(4): Satuan, H(8): HPP, J(10): Rak
      
      // Update Nama
      sheet.getRange(i + 1, 2).setValue(form.name);
      // Update Kategori
      sheet.getRange(i + 1, 3).setValue(form.category);
      // Update HPP
      sheet.getRange(i + 1, 8).setValue(form.hpp);
      // Update Total Aset (Stok Sisa * HPP Baru)
      const sisa = sheet.getRange(i + 1, 7).getValue();
      sheet.getRange(i + 1, 9).setValue(sisa * form.hpp);
      // Update Rak
      sheet.getRange(i + 1, 10).setValue(form.rack);
      
      return { success: true, message: `Data bahan ${form.sku} berhasil diperbarui.` };
    }
  }
  return { success: false, message: `Gagal update. SKU ${form.sku} tidak ditemukan.` };
}

// [Code.gs] - HELPER UNTUK FORM DESAIN (AUTO SKU & KATEGORI)

// 1. AMBIL LIST KATEGORI UNIK DARI MASTER PRODUK
function getProductCategories() {
  const sheet = _getSS().getSheetByName('Master_Produk');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Ambil Kolom C (Kategori)
  const data = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
  
  // Filter Unik & Hapus Kosong
  const uniqueCats = [...new Set(data)].filter(c => c && c !== "");
  return uniqueCats.sort();
}

// 2. GENERATE SKU OTOMATIS BERURUTAN (SMART SKU)
function generateNextProductSku(categoryCode) {
  const sheet = _getSS().getSheetByName('Master_Produk');
  const lastRow = sheet.getLastRow();
  
  // 1. Normalisasi Prefix (Semua jadi Huruf Besar agar aman)
  // Jika categoryCode kosong, pakai "PRD".
  const rawCode = categoryCode ? categoryCode.trim() : 'PRD';
  const prefix = rawCode.toUpperCase() + '-'; 
  
  // Jika database kosong, mulai dari 1001
  if (lastRow < 2) return prefix + '1001'; 

  // 2. Ambil Semua SKU yang Ada (Kolom A)
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  
  let maxNum = 0;

  // 3. Loop cari angka tertinggi (Case Insensitive)
  data.forEach(sku => {
    if (sku) {
      const skuStr = String(sku).toUpperCase().trim(); // Paksa huruf besar
      
      if (skuStr.startsWith(prefix)) {
        // Hapus prefix, ambil angkanya saja
        // Contoh: "KMJ-1005" -> diambil "1005"
        const numPart = parseInt(skuStr.replace(prefix, ''));
        
        // Jika valid angka & lebih besar dari maxNum saat ini, update maxNum
        if (!isNaN(numPart) && numPart > maxNum) {
          maxNum = numPart;
        }
      }
    }
  });

  // Jika belum ada SKU dengan kode ini, start dari 1000
  if (maxNum === 0) maxNum = 1000;

  // 4. Return Urutan Selanjutnya (Max + 1)
  return prefix + (maxNum + 1);
}

// [Code.gs] - AMBIL KATEGORI DARI SHEET CONFIG

function getConfigCategories() {
  const ss = _getSS();
  const sheet = ss.getSheetByName('Config');
  
  // Jika sheet Config belum ada, return array kosong
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Ambil Range A2:B (Kolom A = Nama Kategori, Kolom B = Kode)
  // Kita ambil 2 kolom agar nanti bisa dipakai untuk generate SKU
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  
  // Filter baris yang Nama Kategorinya kosong
  return data.filter(row => row[0] && row[0] !== "");
}

// [Code.gs] - AMBIL DATA BAHAN UNTUK SEARCH DESAIN
function getMaterialOptions() {
  const sheet = _getSS().getSheetByName('Master_Bahan');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Ambil Range A2:I (Kolom C = Kategori ada di index 2)
  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  return data
    .filter(r => r[0] && r[1]) 
    .map(r => ({
      sku: r[0],
      name: r[1],
      category: r[2],   // <--- INI PENTING (Kolom C)
      unit: r[3],       
      hpp: Number(r[8]) 
    }));
}

// [Code.gs] - AMBIL RESEP DARI KOLOM 19 (DESAIN V2)

function getProductRecipe(sku) {
  const ss = _getSS();
  const sProd = ss.getSheetByName('Master_Produk');
  const sMat = ss.getSheetByName('Master_Bahan'); 
  
  const dataProd = sProd.getDataRange().getValues();
  
  // 1. Cari Produk berdasarkan SKU Variasi
  let recipeRaw = null;
  for (let i = 1; i < dataProd.length; i++) {
    if (String(dataProd[i][3]).trim() === String(sku).trim()) {
      recipeRaw = dataProd[i][18]; // Kolom S (Index 18) = JSON Resep
      break;
    }
  }

  if (!recipeRaw || recipeRaw === '') {
    return { success: false, message: 'Belum ada resep tersimpan.' };
  }

  try {
    const recipe = JSON.parse(recipeRaw);
    
    // 2. [PENTING] Ambil Harga & Satuan Terkini dari Master_Bahan
    const dataMat = sMat.getDataRange().getValues();
    const matMap = {}; // Map: SKU -> {hpp, unit}
    
    // Index Master Bahan: [0]SKU, [3]Satuan, [8]HPP
    for(let j=1; j<dataMat.length; j++) {
       const matSku = String(dataMat[j][0]).trim();
       const matUnit = dataMat[j][3]; 
       const matHpp = Number(dataMat[j][8]) || 0;
       
       matMap[matSku] = { unit: matUnit, hpp: matHpp };
    }

    // 3. Injeksi Data ke dalam Object Resep
    const ingredients = recipe.mats || recipe.materials || [];
    
    ingredients.forEach(item => {
        const matData = matMap[item.sku];
        if (matData) {
            item.unit = matData.unit; // Injeksi Satuan
            item.hpp = matData.hpp;   // Injeksi HPP
        } else {
            item.unit = '-';
            item.hpp = 0;
        }
    });
    
    // Kembalikan object resep yang sudah matang
    return { success: true, data: recipe };

  } catch (e) {
    return { success: false, message: 'Format resep rusak: ' + e.message };
  }
}

// ==========================================
// --- LOGIKA SCANNER GUDANG (EKSPEDISI) ---
// ==========================================

// ==========================================
// --- AUTO SCAN: 1x SCAN = SHIPPED + POTONG STOK ---
// Mode: 'AUTO' (default) atau 'RETUR'
// AUTO: WAITING/PACKED → langsung SHIPPED + stok terpotong
// RETUR: Status → RETURNED + restock jika kondisi GOOD
// ==========================================
function processAutoScan(barcode, mode, condition) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) {
    return { success: false, message: "Server Sedang Sibuk. Silakan scan ulang." };
  }

  try {
    if (!barcode) return { success: false, message: "Barcode kosong." };

    var ss = _getSS();
    var sheetOrder = ss.getSheetByName('Data_Pesanan');
    var sheetStok = ss.getSheetByName('Riwayat_Stok');
    var sLogActivity = ss.getSheetByName('Data_Riwayat');
    var user = Session.getActiveUser().getEmail();
    var timestamp = new Date();

    var sheetMaster = ss.getSheetByName('Master_Produk');
    if (!sheetMaster) sheetMaster = ss.getSheetByName('Master Produk & Stok');
    if (!sheetMaster) sheetMaster = ss.getSheetByName('Data_Produk');

    if (!sheetOrder) return { success: false, message: "Tab 'Data_Pesanan' tidak ditemukan." };
    if (!sheetMaster) return { success: false, message: "Tab Master Produk tidak ditemukan." };

    var data = sheetOrder.getDataRange().getValues();

    // Normalisasi status
    function normalizeStatus(raw) {
      var s = String(raw || "").replace(/\s+/g, ' ').trim().toUpperCase();
      if (s === 'SIAP KIRIM' || s === 'READY TO SHIP' || s === 'PACKING DONE') return 'PACKED';
      if (s === 'DIKIRIM' || s === 'SHIPPING' || s === 'SENT') return 'SHIPPED';
      if (s === 'MENUNGGU' || s === 'ANTRIAN' || s === 'TO PACK') return 'WAITING';
      return s;
    }

    function normalizeSkuValue(raw) {
      return String(raw || '').replace(/^'+/, '').replace(/\s+/g, '').trim();
    }

    function pickSkuCandidates(row) {
      var skuVar = normalizeSkuValue(row[7]);
      var skuInd = normalizeSkuValue(row[5]);
      var invalids = new Set(['', '-', 'N/A', 'NA', 'NULL', 'UNDEFINED']);
      var list = [];
      if (!invalids.has(String(skuVar).toUpperCase())) list.push(skuVar);
      if (!invalids.has(String(skuInd).toUpperCase()) && skuInd !== skuVar) list.push(skuInd);
      return list;
    }

    // Cari semua baris pesanan dengan resi ini
    var rowsToProcess = [];
    var currentStatus = "";
    var found = false;

    for (var i = 1; i < data.length; i++) {
      var rowResi = String(data[i][4]).trim().toLowerCase();
      var scanResi = String(barcode).trim().toLowerCase();
      if (rowResi === scanResi) {
        var skuCandidates = pickSkuCandidates(data[i]);
        rowsToProcess.push({
          rowIndex: i + 1,
          sku: skuCandidates[0] || '',
          skuCandidates: skuCandidates,
          qty: parseFloat(data[i][9]) || 0,
          nama: data[i][6],
          variasi: data[i][8]
        });
        if (!found) {
          currentStatus = normalizeStatus(data[i][14]);
          found = true;
        }
      }
    }

    if (!found) return { success: false, message: "Resi tidak ditemukan." };

    // Tentukan mode efektif (default AUTO)
    var effectiveMode = String(mode || 'AUTO').toUpperCase();

    var newStatus = "";
    var successTitle = "Berhasil";
    var successMsg = "";
    var logType = "";

    // --- MODE AUTO: 1x scan = langsung SHIPPED + potong stok ---
    if (effectiveMode === 'AUTO') {
      // Sudah SHIPPED? Info saja
      if (currentStatus === 'SHIPPED') {
        return { success: true, title: "Info", message: "Paket ini sudah dikirim sebelumnya." };
      }
      // Sudah RETURNED? Tolak
      if (currentStatus === 'RETURNED') {
        return { success: false, message: "Paket ini sudah diretur. Tidak bisa diproses ulang." };
      }
      // WAITING atau PACKED → langsung SHIPPED + potong stok
      if (currentStatus === 'WAITING' || currentStatus === 'PACKED' || currentStatus === '') {
        newStatus = "SHIPPED";
        successTitle = "Dikirim";
        processStockOut(sheetStok, sheetMaster, rowsToProcess, timestamp, barcode, user);
        successMsg = "Stok Terpotong & Status Terkirim.";
        logType = "AUTO SHIP";
      } else {
        return { success: false, message: "Status saat ini: " + currentStatus + ". Tidak bisa diproses." };
      }
    }

    // --- MODE RETUR: Kembalikan stok jika GOOD ---
    else if (effectiveMode === 'RETUR') {
      if (currentStatus === 'RETURNED') {
        return { success: true, title: "Info", message: "Paket ini sudah diretur sebelumnya." };
      }
      newStatus = "RETURNED";
      successTitle = "Retur Diterima";

      if (condition === 'GOOD') {
        processReturRestock(sheetStok, sheetMaster, rowsToProcess, timestamp, barcode, user);
        successMsg = "Stok Dikembalikan (Good).";
      } else {
        successMsg = "Barang Rusak (Bad). Stok tidak berubah.";
      }
      logType = "RETURN";
    }

    else {
      return { success: false, message: "Mode tidak dikenali: " + effectiveMode };
    }

    // UPDATE STATUS PESANAN
    var totalQty = 0;
    rowsToProcess.forEach(function(item) {
      totalQty += item.qty;
      var r = item.rowIndex;
      sheetOrder.getRange(r, 15).setValue(newStatus); // Col O

      if (effectiveMode === 'AUTO') {
        // Set waktu packed + shipped sekaligus (karena skip PACKED)
        sheetOrder.getRange(r, 16).setValue(timestamp); // Col P (Packed time)
        sheetOrder.getRange(r, 17).setValue(timestamp); // Col Q (Shipped time)
      }
      if (effectiveMode === 'RETUR') {
        sheetOrder.getRange(r, 18).setValue(timestamp); // Col R (Retur time)
        if (condition) sheetOrder.getRange(r, 21).setValue("Kondisi: " + condition);
      }
    });

    // LOG ACTIVITY
    if (sLogActivity) {
      var itemFirst = rowsToProcess[0];
      var detailName = itemFirst.nama + (rowsToProcess.length > 1 ? " (+Item Lain)" : "");
      sLogActivity.appendRow([timestamp, barcode, logType, itemFirst.sku, detailName, totalQty, user, successMsg]);
    }

    SpreadsheetApp.flush();
    return { success: true, title: successTitle, message: successMsg };

  } catch (err) {
    Logger.log("Error AutoScan: " + err);
    return { success: false, message: "Error Sistem: " + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function processExpeditionScan(barcode, mode, condition) {
  // [LOCK START] Mencegah tabrakan data saat scan cepat
  const lock = LockService.getScriptLock();
  try {
      // Tunggu antrian maksimal 15 detik. Jika server sibuk > 15dtk, batalkan.
      lock.waitLock(15000); 
  } catch (e) {
      return { success: false, message: "Server Sedang Sibuk. Silakan scan ulang." };
  }

  try {
      // --- LOGIKA UTAMA DIMULAI DISINI ---
      if (!barcode) return { success: false, message: "Barcode kosong." };
      
      const ss = _getSS();
      const sheetOrder = ss.getSheetByName('Data_Pesanan');
      const sheetStok = ss.getSheetByName('Riwayat_Stok');
      const sLogActivity = ss.getSheetByName('Data_Riwayat');
      const user = Session.getActiveUser().getEmail();
      const timestamp = new Date();

      // PENTING:
      // UI web app memakai tab Master_Produk untuk membaca stok.
      // Jadi prioritas harus Master_Produk dulu agar stok yang terlihat user ikut berubah.
      let sheetMaster = ss.getSheetByName('Master_Produk');
      if (!sheetMaster) sheetMaster = ss.getSheetByName('Master Produk & Stok');
      if (!sheetMaster) sheetMaster = ss.getSheetByName('Data_Produk');
      
      if (!sheetOrder) return { success: false, message: "Tab 'Data_Pesanan' tidak ditemukan." };
      if (!sheetMaster) return { success: false, message: "Tab Master Produk tidak ditemukan." };

      const data = sheetOrder.getDataRange().getValues();
      
      // Cari Data Pesanan
      let rowsToProcess = []; 
      let currentStatus = "";
      let found = false;

      // Normalisasi status dari sheet (antisipasi spasi/kapital/variasi teks)
      function normalizeStatus(raw) {
        const s = String(raw || "").replace(/\s+/g, ' ').trim().toUpperCase();
        if (s === 'SIAP KIRIM' || s === 'READY TO SHIP' || s === 'PACKING DONE') return 'PACKED';
        if (s === 'DIKIRIM' || s === 'SHIPPING' || s === 'SENT') return 'SHIPPED';
        if (s === 'MENUNGGU' || s === 'ANTRIAN' || s === 'TO PACK') return 'WAITING';
        return s;
      }
      
      function normalizeSkuValue(raw) {
        return String(raw || '')
          .replace(/^'+/, '') // buang apostrophe excel
          .replace(/\s+/g, '')
          .trim();
      }

      function pickSkuCandidates(row) {
        const skuVar = normalizeSkuValue(row[7]); // H: SKU Variasi
        const skuInd = normalizeSkuValue(row[5]); // F: SKU Induk
        const invalids = new Set(['', '-', 'N/A', 'NA', 'NULL', 'UNDEFINED']);

        const list = [];
        if (!invalids.has(String(skuVar).toUpperCase())) list.push(skuVar);
        if (!invalids.has(String(skuInd).toUpperCase()) && skuInd !== skuVar) list.push(skuInd);
        return list;
      }

      for (let i = 1; i < data.length; i++) {
        const rowResi = String(data[i][4]).trim().toLowerCase();
        const scanResi = String(barcode).trim().toLowerCase();
        if (rowResi === scanResi) {
          const skuCandidates = pickSkuCandidates(data[i]);
          rowsToProcess.push({
            rowIndex: i + 1,
            sku: skuCandidates[0] || '',
            skuCandidates: skuCandidates,
            qty: parseFloat(data[i][9]) || 0, 
            nama: data[i][6],
            variasi: data[i][8]
          });
          if (!found) {
            currentStatus = normalizeStatus(data[i][14]); 
            found = true;
          }
        }
      }

      if (!found) return { success: false, message: "Resi tidak ditemukan." };

      let newStatus = "";
      let successTitle = "Berhasil";
      let successMsg = "";
      let logType = ""; 

      // --- MODE IN (PACKING) ---
      if (mode === 'IN') {
        if (currentStatus === 'PACKED') return { success: true, title: "Info", message: "Sudah dipacking." };
        if (currentStatus !== 'WAITING' && currentStatus !== '') return { success: false, message: `Status: ${currentStatus}.` };
        
        newStatus = "PACKED";
        successTitle = "Siap Kirim";
        successMsg = `Packing selesai.`;
        logType = "PACKING";
      }
      
      // --- MODE OUT (KIRIM & POTONG STOK) ---
      else if (mode === 'OUT') {
        if (currentStatus === 'SHIPPED') return { success: true, title: "Info", message: "Sudah dikirim." };
        if (currentStatus !== 'PACKED') return { success: false, message: `OUT ditolak. Status saat ini: ${currentStatus || '-'} (wajib PACKED).` };
        
        newStatus = "SHIPPED";
        successTitle = "Dikirim";
        // Panggil helper update stok (User diteruskan)
        processStockOut(sheetStok, sheetMaster, rowsToProcess, timestamp, barcode, user);
        successMsg = `Stok Terpotong & Status Terkirim.`;
        logType = "SHIPPING";
      }

      // --- MODE RETUR (KEMBALIKAN STOK) ---
      else if (mode === 'RETUR') {
        newStatus = "RETURNED";
        successTitle = "Retur Diterima";
        
        if (condition === 'GOOD') {
           processReturRestock(sheetStok, sheetMaster, rowsToProcess, timestamp, barcode, user);
           successMsg = `Stok Dikembalikan (Good).`;
        } else {
           successMsg = `Barang Rusak (Bad).`;
        }
        logType = "RETURN";
      }

      // UPDATE STATUS
      let totalQty = 0;
      rowsToProcess.forEach(item => {
          totalQty += item.qty;
          const r = item.rowIndex;
          sheetOrder.getRange(r, 15).setValue(newStatus); 

          if (mode === 'IN') sheetOrder.getRange(r, 16).setValue(timestamp);
          if (mode === 'OUT') sheetOrder.getRange(r, 17).setValue(timestamp);
          if (mode === 'RETUR') {
              sheetOrder.getRange(r, 18).setValue(timestamp);
              if (condition) sheetOrder.getRange(r, 21).setValue("Kondisi: " + condition);
          }
      });

      // LOG ACTIVITY
      if (sLogActivity) {
          const itemFirst = rowsToProcess[0];
          const detailName = itemFirst.nama + (rowsToProcess.length > 1 ? " (+Item Lain)" : "");
          sLogActivity.appendRow([timestamp, barcode, logType, itemFirst.sku, detailName, totalQty, user, successMsg]);
      }
      
      SpreadsheetApp.flush(); // Paksa simpan perubahan sebelum Lock dilepas
      return { success: true, title: successTitle, message: successMsg };

  } catch (err) {
      Logger.log("Error Scanner: " + err);
      return { success: false, message: "Error Sistem: " + err.toString() };
  } finally {
      // [LOCK END] PENTING: Kunci harus dilepas apapun yang terjadi
      lock.releaseLock();
  }
}

// ============================================================
// --- FUNGSI UPDATE STOK YANG BENAR (KELUAR, SISA, ASET) ---
// ============================================================

function processStockOut(sheetStok, sheetMaster, items, time, resi, user) {
  const masterData = sheetMaster.getDataRange().getValues();

  function normalizeText(raw) {
    return String(raw || '').toLowerCase().replace(/\s+/g, '').replace(/[-_/]+/g, ',').trim();
  }

  function varMatch(itemVar, masterVar) {
    if (itemVar === masterVar) return true;
    var oP = itemVar.split(',').filter(function(s){return s.length>0}).sort();
    var mP = masterVar.split(',').filter(function(s){return s.length>0}).sort();
    return oP.length > 0 && oP.length === mP.length && oP.join(',') === mP.join(',');
  }

  const historyRows = [];

  items.forEach(item => {
      if (!item.qty) return;

      let matched = false;
      const itemSkuNorm = String(item.sku || '').trim().toLowerCase();
      const itemNamaNorm = normalizeText(item.nama);
      const itemVarNorm = normalizeText(item.variasi);

      // Matching 3 tier: SKU exact → Nama+Variasi → Variasi saja (fallback)
      var matchedRow = -1;

      // Tier 1: SKU exact match
      if (itemSkuNorm) {
        for (let m = 1; m < masterData.length; m++) {
          var masterSku = String(masterData[m][3] || '').trim().toLowerCase();
          if (masterSku && masterSku === itemSkuNorm) { matchedRow = m; break; }
        }
      }

      // Tier 2: Nama produk + Variasi
      if (matchedRow === -1 && itemNamaNorm) {
        for (let m = 1; m < masterData.length; m++) {
          var masterNama = normalizeText(masterData[m][1]);
          if (!masterNama) continue;
          if (!(masterNama.includes(itemNamaNorm) || itemNamaNorm.includes(masterNama))) continue;
          if (itemVarNorm) {
            var masterVar = normalizeText(masterData[m][4]);
            if (varMatch(itemVarNorm, masterVar)) { matchedRow = m; break; }
          } else {
            matchedRow = m; break;
          }
        }
      }

      // Tier 3: Variasi saja (fallback, hanya jika variasi cukup spesifik > 3 char)
      if (matchedRow === -1 && itemVarNorm && itemVarNorm.length > 3) {
        for (let m = 1; m < masterData.length; m++) {
          var masterVarF = normalizeText(masterData[m][4]);
          if (masterVarF && varMatch(itemVarNorm, masterVarF)) { matchedRow = m; break; }
        }
      }

      if (matchedRow > 0) {
        const oldKeluar = parseFloat(masterData[matchedRow][6]) || 0;
        const oldSisa   = parseFloat(masterData[matchedRow][7]) || 0;
        const hpp       = parseFloat(masterData[matchedRow][9]) || 0;
        const newKeluar = oldKeluar + item.qty;
        const newSisa   = oldSisa - item.qty;

        // Update sheet
        sheetMaster.getRange(matchedRow + 1, 7, 1, 2).setValues([[newKeluar, newSisa]]);
        sheetMaster.getRange(matchedRow + 1, 11).setValue(newSisa * hpp);

        // Update in-memory juga (agar item berikutnya baca data terbaru)
        masterData[matchedRow][6] = newKeluar;
        masterData[matchedRow][7] = newSisa;
        masterData[matchedRow][10] = newSisa * hpp;

        // History row SETELAH matching — dengan HPP dan sisa yang benar
        historyRows.push([
            'OUT-' + Math.floor(Math.random()*100000), time, 'OUT (SALES)', resi,
            item.sku || String(masterData[matchedRow][3] || ''), item.nama, item.variasi, item.qty,
            hpp, hpp * item.qty, newSisa,
            'Penjualan Scan Out', user || 'Admin Gudang', new Date()
        ]);
        matched = true;
      }

      if (!matched) {
        // Tetap tulis history (tanpa HPP) agar ada jejak
        historyRows.push([
            'OUT-' + Math.floor(Math.random()*100000), time, 'OUT (SALES)', resi,
            item.sku || '', item.nama, item.variasi, item.qty, 0, 0, 0,
            'Penjualan Scan Out (UNMATCHED)', user || 'Admin Gudang', new Date()
        ]);
        Logger.log('WARNING processStockOut: Produk tidak ketemu. sku=' + item.sku + ', nama=' + item.nama + ', variasi=' + item.variasi + ', resi=' + resi);
      }
  });

  // BATCH APPEND riwayat stok
  if (historyRows.length > 0 && sheetStok) {
    sheetStok.getRange(sheetStok.getLastRow() + 1, 1, historyRows.length, 14).setValues(historyRows);
  }
}

function processReturRestock(sheetStok, sheetMaster, items, time, resi, user) {
  const masterData = sheetMaster.getDataRange().getValues();

  function normalizeText(raw) {
    return String(raw || '').toLowerCase().replace(/\s+/g, '').replace(/[-_/]+/g, ',').trim();
  }

  function varMatch(itemVar, masterVar) {
    if (itemVar === masterVar) return true;
    var oP = itemVar.split(',').filter(function(s){return s.length>0}).sort();
    var mP = masterVar.split(',').filter(function(s){return s.length>0}).sort();
    return oP.length > 0 && oP.length === mP.length && oP.join(',') === mP.join(',');
  }

  const historyRows = [];

  items.forEach(item => {
      if (!item.qty) return;

      const itemSkuNorm = String(item.sku || '').trim().toLowerCase();
      const itemNamaNorm = normalizeText(item.nama);
      const itemVarNorm = normalizeText(item.variasi);

      var matchedRow = -1;

      // Tier 1: SKU exact
      if (itemSkuNorm) {
        for (let m = 1; m < masterData.length; m++) {
          var masterSku = String(masterData[m][3] || '').trim().toLowerCase();
          if (masterSku && masterSku === itemSkuNorm) { matchedRow = m; break; }
        }
      }

      // Tier 2: Nama + Variasi
      if (matchedRow === -1 && itemNamaNorm) {
        for (let m = 1; m < masterData.length; m++) {
          var masterNama = normalizeText(masterData[m][1]);
          if (!masterNama) continue;
          if (!(masterNama.includes(itemNamaNorm) || itemNamaNorm.includes(masterNama))) continue;
          if (itemVarNorm) {
            if (varMatch(itemVarNorm, normalizeText(masterData[m][4]))) { matchedRow = m; break; }
          } else { matchedRow = m; break; }
        }
      }

      // Tier 3: Variasi saja (fallback, > 3 char)
      if (matchedRow === -1 && itemVarNorm && itemVarNorm.length > 3) {
        for (let m = 1; m < masterData.length; m++) {
          if (varMatch(itemVarNorm, normalizeText(masterData[m][4]))) { matchedRow = m; break; }
        }
      }

      if (matchedRow > 0) {
        const oldSisa   = parseFloat(masterData[matchedRow][7]) || 0;
        const hpp       = parseFloat(masterData[matchedRow][9]) || 0;
        const oldMasuk  = (masterData[matchedRow].length > 20) ? parseFloat(masterData[matchedRow][20]) : 0;
        const newMasuk = oldMasuk + item.qty;
        const newSisa  = oldSisa + item.qty;

        sheetMaster.getRange(matchedRow + 1, 8).setValue(newSisa);
        sheetMaster.getRange(matchedRow + 1, 11).setValue(newSisa * hpp);
        sheetMaster.getRange(matchedRow + 1, 21).setValue(newMasuk);

        // Update in-memory
        masterData[matchedRow][7] = newSisa;
        masterData[matchedRow][10] = newSisa * hpp;
        if (masterData[matchedRow].length > 20) masterData[matchedRow][20] = newMasuk;

        historyRows.push([
            'RET-' + Math.floor(Math.random()*100000), time, 'IN (RETUR)', resi,
            item.sku || String(masterData[matchedRow][3] || ''), item.nama, item.variasi, item.qty,
            hpp, hpp * item.qty, newSisa,
            'Retur Good Condition', user || 'Admin Gudang', new Date()
        ]);
      } else {
        historyRows.push([
            'RET-' + Math.floor(Math.random()*100000), time, 'IN (RETUR)', resi,
            item.sku || '', item.nama, item.variasi, item.qty, 0, 0, 0,
            'Retur (UNMATCHED)', user || 'Admin Gudang', new Date()
        ]);
        Logger.log('WARNING processReturRestock: Produk tidak ketemu. sku=' + item.sku + ', nama=' + item.nama + ', variasi=' + item.variasi);
      }
  });

  // BATCH APPEND riwayat stok
  if (historyRows.length > 0 && sheetStok) {
    sheetStok.getRange(sheetStok.getLastRow() + 1, 1, historyRows.length, 14).setValues(historyRows);
  }
}

// ==========================================
// --- FITUR RIWAYAT MUTASI (STOCK CARD) ---
// ==========================================

// --- HELPER: KONVERSI TANGGAL GOOGLE SHEET ---
function parseSheetDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  
  // Jika berupa Angka Serial (contoh: 46036.9033)
  if (typeof value === 'number') {
      // Rumus konversi Serial Excel/Sheets ke JS Date
      // 25569 adalah offset hari antara 1 Jan 1970 dan 30 Des 1899
      return new Date(Math.round((value - 25569) * 86400 * 1000));
  }
  
  // Jika string, coba parsing standar
  return new Date(value);
}

// --- FUNGSI UTAMA ---
function getProductMutationHistory(sku) {
  const ss = _getSS();
  let history = [];

  // 1. NORMALISASI INPUT
  // Hapus spasi, lowercase, dan ganti dash aneh dengan dash standar
  const cleanStr = (s) => String(s || "").trim().toLowerCase().replace(/\u2013|\u2014/g, "-");
  
  let targetSku = "";
  if (typeof sku === 'object' && sku !== null) {
      targetSku = cleanStr(sku.sku || sku.toString());
  } else {
      targetSku = cleanStr(sku);
  }

  if (!targetSku || targetSku === "undefined") return [];

  // 2. STRATEGI PENCARIAN (Mapping Parent-Child)
  let searchSkus = new Set();
  searchSkus.add(targetSku);

  const sheetMaster = ss.getSheetByName('Master_Produk');
  if (sheetMaster) {
      const dataMaster = sheetMaster.getDataRange().getValues();
      for (let m = 1; m < dataMaster.length; m++) {
          const rowInduk = cleanStr(dataMaster[m][0]); // Col A
          const rowAnak  = cleanStr(dataMaster[m][3]); // Col D
          
          if (rowInduk === targetSku && rowAnak) searchSkus.add(rowAnak);
          if (rowAnak === targetSku && rowInduk) searchSkus.add(rowInduk);
      }
  }

  // 3. AMBIL DATA RIWAYAT
  const sheetStok = ss.getSheetByName('Riwayat_Stok');
  if (!sheetStok) return [];

  const data = sheetStok.getDataRange().getValues();
  // Loop mulai baris 2 (Index 1)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowSku = cleanStr(row[4]); // Kolom E (SKU)

    // LOGIKA PENCOCOKAN:
    // 1. Cek Exact Match di dalam Set (Induk/Anak)
    // 2. Cek Partial Match (StartsWith) untuk menangkap variasi jika mapping gagal
    //    Contoh: Target "PSM-001" akan menangkap "PSM-001-HITAM"
    const isMatch = searchSkus.has(rowSku) || rowSku.startsWith(targetSku + "-") || rowSku === targetSku;

    if (isMatch) {
        // Deteksi Tipe (IN/OUT)
        let rawType = String(row[2] || "").toUpperCase();
        let displayType = 'IN';
        if (rawType.match(/OUT|KELUAR|SALES|SHIPPED|SOLD/)) {
            displayType = 'OUT';
        }

        // Ambil Tanggal: Prioritas Timestamp (N/13) > Tanggal (B/1)
        // Gunakan helper parseSheetDate untuk menangani angka 46036...
        let dateVal = parseSheetDate(row[13]) || parseSheetDate(row[1]);
        
        // Format ISO String agar aman dikirim ke Frontend
        let dateIso = dateVal ? dateVal.toISOString() : null;

        history.push({
            date: dateIso,           // Kirim dalam format ISO String
            type: displayType,
            ref: row[3] || row[0],   // Ref / ID
            desc: row[11] || rawType,
            qty: Number(row[7]) || 0,
            user: String(row[12] || 'System')
        });
    }
  }

  // 4. SORTING (Terbaru di Atas)
  return history.sort((a, b) => {
      const tA = a.date ? new Date(a.date).getTime() : 0;
      const tB = b.date ? new Date(b.date).getTime() : 0;
      return tB - tA;
  });
}

// ==========================================
// --- DASHBOARD ANALYTICS (HYBRID V3.0) ---
// ==========================================

function getAdvancedDashboardData(startDateStr, endDateStr) {
  // 0. AMBIL CONFIG & SETUP
  const config = getAppConfig();
  const LIMIT_PROD = Number(config.limit_product) || 10;
  const LIMIT_MAT = Number(config.limit_material) || 20;

  const ss = _getSS();
  const sheetPesanan = ss.getSheetByName('Data_Pesanan');
  const sheetProduk = ss.getSheetByName('Master_Produk');
  const sheetBahan = ss.getSheetByName('Master_Bahan');
  const sheetRekap = ss.getSheetByName('Data_Rekap'); // [BARU] Baca Sheet Rekap

  // 1. Parsing Tanggal Filter
  let start, end;
  try {
      start = new Date(startDateStr); start.setHours(0,0,0,0);
      end = new Date(endDateStr); end.setHours(23,59,59,999);
  } catch(e) {
      const now = new Date();
      start = new Date(now.getFullYear(), now.getMonth(), 1);
      end = now;
  }

  // --- BAGIAN A: HITUNG SNAPSHOT ASET (REALTIME) ---
  // (Bagian ini tidak berubah, karena stok selalu real-time)
  let totalAssetProducts = 0;
  let lowStockCount = 0;
  const lowStockList = [];

  if (sheetProduk) {
      const dataProduk = sheetProduk.getDataRange().getValues();
      for (let i = 1; i < dataProduk.length; i++) {
          const row = dataProduk[i];
          if (!row[1] && !row[3]) continue;
          const stok = parseFloat(row[7]) || 0; 
          const aset = parseFloat(row[10]) || 0; 
          
          totalAssetProducts += aset;
          if(stok <= LIMIT_PROD) {
              lowStockCount++;
              if (lowStockList.length < 5) { 
                  lowStockList.push({ name: row[1], sku: row[3], variant: row[4]||'', qty: stok, unit: row[8]||'Pcs' });
              }
          }
      }
  }

  let totalAssetMaterials = 0;
  const lowMaterialList = [];
  if (sheetBahan) {
      const dataBahan = sheetBahan.getDataRange().getValues();
      for (let i = 1; i < dataBahan.length; i++) {
          const row = dataBahan[i];
          if(!row[0]) continue;
          
          const stokBahan = parseFloat(row[7]) || 0;
          const asetBahan = parseFloat(row[9]) || 0; 
          
          totalAssetMaterials += asetBahan;
          if (stokBahan <= LIMIT_MAT) {
              if (lowMaterialList.length < 5) {
                  lowMaterialList.push({ name: row[1], sku: row[0], qty: stokBahan, unit: row[3]||'Unit' });
              }
          }
      }
  }

  // --- BAGIAN B: HITUNG OMSET & PERFORMA (HYBRID) ---
  
  // Variabel Penampung Total
  let totalOmset = 0;
  let totalModal = 0;
  let totalProfit = 0;
  let totalQty = 0;
  let totalTrxCount = 0; // Gabungan transaksi live + arsip

  // Variabel Logistik (Hanya dari Data Live karena status di arsip pasti "Selesai")
  const setWaiting = new Set();
  const setPacked = new Set();
  const setShipped = new Set();
  const setReturned = new Set();
  let returnTotalValue = 0;
  let returnGoodItems = 0;
  let returnBadItems = 0;

  // Chart & Channel Data
  const chartData = {}; 
  const channelData = {};
  const productPerformance = {};

  // --- SUMBER 1: DATA ARSIP (REKAP) ---
  // Kita cek Data_Rekap dulu untuk mengisi angka masa lalu
  if (sheetRekap) {
    const dataRekap = sheetRekap.getDataRange().getValues();
    // Loop mulai baris 2 (Index 1)
    // Struktur Data_Rekap: [0]ID, [1]Bulan, [2]Tahun, [3]Omset, [4]Modal, [5]Profit, [6]Qty, [7]Trx
    for (let r = 1; r < dataRekap.length; r++) {
       const row = dataRekap[r];
       const idRekap = String(row[0]); // Contoh: "REKAP-2025-10"
       
       // Parsing Tanggal dari ID agar akurat
       if (idRekap.startsWith("REKAP-")) {
          const parts = idRekap.split("-"); // ["REKAP", "2025", "10"]
          const year = parseInt(parts[1]);
          const month = parseInt(parts[2]) - 1; // JS Month 0-11
          
          // Kita anggap tanggal 1 bulan tersebut untuk filter range
          const rekapDate = new Date(year, month, 1);
          
          // Cek apakah bulan rekap ini masuk dalam filter tanggal User?
          // Kita cek sederhana: apakah akhir bulan rekap >= start DAN awal bulan rekap <= end
          const endOfMonth = new Date(year, month + 1, 0);
          
          if (endOfMonth >= start && rekapDate <= end) {
             // MASUK RANGE! Tambahkan ke Total
             totalOmset += Number(row[3]) || 0;
             totalModal += Number(row[4]) || 0;
             totalProfit += Number(row[5]) || 0;
             totalQty += Number(row[6]) || 0;
             totalTrxCount += Number(row[7]) || 0;

             // Tambahkan ke Grafik (Sebagai 1 titik di tanggal 1 bulan tsb)
             const dateKey = Utilities.formatDate(rekapDate, Session.getScriptTimeZone(), "yyyy-MM-01");
             if (!chartData[dateKey]) chartData[dateKey] = { omset: 0, profit: 0 };
             chartData[dateKey].omset += (Number(row[3]) || 0);
             chartData[dateKey].profit += (Number(row[5]) || 0);
          }
       }
    }
  }

  // --- SUMBER 2: DATA LIVE (DATA_PESANAN) ---
  const uniqueOrdersLive = new Set(); // Untuk hitung count pesanan live

  if (sheetPesanan) {
      const dataPesanan = sheetPesanan.getDataRange().getValues();
      for (let i = 1; i < dataPesanan.length; i++) {
          const rawDate = dataPesanan[i][1];
          if(!rawDate) continue;

          const rowDate = new Date(rawDate);
          
          // Filter Tanggal
          if (rowDate >= start && rowDate <= end) {
              const orderId = String(dataPesanan[i][0]);
              const status = String(dataPesanan[i][14]).toUpperCase();
              
              const qtyRow = parseFloat(dataPesanan[i][9]) || 0;
              // Logika Omset: Prioritas Kolom T (Net Marketplace), lalu L (Gross)
              const omsetRow = parseFloat(dataPesanan[i][19]) || parseFloat(dataPesanan[i][11]) || 0;
              const modalRow = parseFloat(dataPesanan[i][13]) || 0;
              const profitRow = omsetRow - modalRow;
              const channel = String(dataPesanan[i][2]) || 'Manual';

              // 1. HITUNG OMSET (Hanya jika status valid)
              if (status !== 'CANCELED' && status !== 'RETURNED') {
                  uniqueOrdersLive.add(orderId);
                  
                  totalQty += qtyRow;
                  totalOmset += omsetRow;
                  totalModal += modalRow;
                  totalProfit += profitRow;

                  // Masukkan ke Chart Harian
                  const dateKey = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
                  if (!chartData[dateKey]) chartData[dateKey] = { omset: 0, profit: 0 };
                  chartData[dateKey].omset += omsetRow;
                  chartData[dateKey].profit += profitRow;

                  // Best Seller (Hanya dari data live yg detailnya ada)
                  let skuKey = String(dataPesanan[i][7] || "").trim();
                  if (!skuKey) skuKey = String(dataPesanan[i][5] || "").trim();

                  // 2. Ambil Nama Produk (Col G / Index 6) & Nama Variasi (Col I / Index 8)
                  const namaProd = String(dataPesanan[i][6] || "").trim();
                  const namaVar  = String(dataPesanan[i][8] || "").trim();

                  // 3. Proses Grouping
                  if (skuKey) {
                      if (!productPerformance[skuKey]) {
                          
                          // Susun Nama Tampilan: "Nama Produk (Nama Variasi)"
                          // Kita perlu menampilkan variasi agar Anda tahu mana yang Mocca, mana yang Hitam
                          let finalName = namaProd;
                          if (namaVar && namaVar !== '-' && namaVar !== '') {
                              finalName = `${namaProd} (${namaVar})`;
                          }

                          productPerformance[skuKey] = { 
                              name: finalName || skuKey, 
                              qty: 0 
                          };
                      }
                      
                      // 4. Akumulasi Jumlah (Qty dari variable qtyRow yang sudah diambil dari Kolom J)
                      productPerformance[skuKey].qty += qtyRow;
                  }
              }

              // 2. LOGISTIK (Hanya Data Live)
              if (status.includes('WAITING')) setWaiting.add(orderId);
              else if (status === 'PACKED') setPacked.add(orderId);
              else if (status === 'SHIPPED') setShipped.add(orderId);
              else if (status === 'RETURNED') {
                  setReturned.add(orderId);
                  returnTotalValue += omsetRow;
                  const condition = String(dataPesanan[i][20] || "").toUpperCase();
                  if (condition.includes('GOOD') || condition.includes('BAGUS') || condition.includes('OK')) {
                      returnGoodItems += qtyRow;
                  } else {
                      returnBadItems += qtyRow;
                  }
              }

              // Channel (Semua trafik dihitung)
              if (!channelData[channel]) channelData[channel] = 0;
              channelData[channel]++;
          }
      }
  }

  // Gabungkan Total Transaksi
  totalTrxCount += uniqueOrdersLive.size;

  // Sorting Chart Labels
  const sortedDates = Object.keys(chartData).sort();
  const chartLabels = sortedDates.map(d => {
      const parts = d.split('-'); 
      const dt = new Date(parts[0], parts[1]-1, parts[2]);
      // Jika tanggal 01 dan berasal dari Rekap, mungkin mau format "Okt 2025"
      // Tapi biar konsisten grafik garisnya, kita pakai format tanggal biasa
      return Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd MMM");
  });

  // Sorting Best Seller
  const sortedBestSellers = Object.values(productPerformance)
      .sort((a, b) => b.qty - a.qty) // Urutkan dari Qty Terbesar
      .slice(0, 5); // Ambil 5 Teratas

  return {
      financial: { 
          omset: totalOmset, 
          modal: totalModal, 
          profit: totalProfit, 
          count: totalTrxCount, 
          qty: totalQty 
      },
      logistics: { 
          waiting: setWaiting.size, 
          packed: setPacked.size, 
          shipped: setShipped.size, 
          returned: setReturned.size,
          returnValue: returnTotalValue,
          returnGood: returnGoodItems,
          returnBad: returnBadItems
      },
      inventory: { 
          assetTotal: totalAssetProducts + totalAssetMaterials,
          assetProducts: totalAssetProducts,
          assetMaterials: totalAssetMaterials,
          low: lowStockCount 
      },
      alerts: { products: lowStockList, materials: lowMaterialList },
      bestSellers: sortedBestSellers,
      charts: { 
          labels: chartLabels, 
          omset: sortedDates.map(d => chartData[d].omset), 
          profit: sortedDates.map(d => chartData[d].profit) 
      },
      channels: channelData
  };
}

// ==========================================
// --- FITUR PENGATURAN (CONFIG I & J) ---
// ==========================================

// 1. Ambil Data Konfigurasi (Mengembalikan Object)
function getAppConfig() {
  const ss = _getSS();
  const sheet = ss.getSheetByName('Config');
  
  // Default Config (Jika Sheet Kosong)
  const config = {
      app_name: 'Gudang Maju Jaya',
      owner_name: 'Admin',
      owner_role: 'Super Admin',
      owner_img: '',
      limit_product: 10,
      limit_material: 20,
      page_title: 'Smart Inventory',
      sidebar_name: 'SMART INVENTORY',
      sidebar_subtitle: 'BANTUSELLER V2.0'
  };

  if (!sheet) return config;

  // Ambil Data dari Kolom I (9) dan J (10)
  // Asumsi Baris 1 adalah Header, data mulai baris 2
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return config;

  const data = sheet.getRange(2, 9, lastRow - 1, 2).getValues(); // Ambil Col I & J
  
  data.forEach(row => {
      const key = String(row[0]).trim();
      const val = row[1];
      if (key) {
          config[key] = val;
      }
  });

  return config;
}

// 2. Simpan Konfigurasi (Handle Text & Image)
function saveAppConfig(form) {
  const ss = _getSS();
  let sheet = ss.getSheetByName('Config');
  
  // Jika sheet belum ada, buat baru (safety)
  if (!sheet) {
      sheet = ss.insertSheet('Config');
      sheet.getRange("I1").setValue("Key");
      sheet.getRange("J1").setValue("Value");
  }

  // Handle Upload Foto Profil Baru (Jika ada)
  let photoUrl = form.current_img; // Pakai URL lama defaultnya
  if (form.new_profile_img && form.new_profile_img.data) {
      const fileName = "PROFILE_" + Date.now() + ".jpg";
      photoUrl = saveImageToDrive(form.new_profile_img, fileName);
  }

  // Siapkan Data yang mau disimpan (Key -> Value)
  const updates = {
      app_name: form.app_name,
      owner_name: form.owner_name,
      owner_role: form.owner_role,
      owner_img: photoUrl,
      limit_product: Number(form.limit_product) || 10,
      limit_material: Number(form.limit_material) || 20,
      page_title: form.page_title || 'Smart Inventory',
      sidebar_name: form.sidebar_name || 'SMART INVENTORY',
      sidebar_subtitle: form.sidebar_subtitle || 'BANTUSELLER V2.0'
  };

  // Logika Update: Cek apakah Key sudah ada di Col I?
  // Jika ada update Col J. Jika tidak, Append.
  
  const lastRow = sheet.getLastRow();
  let existingData = [];
  if (lastRow >= 2) {
      existingData = sheet.getRange(2, 9, lastRow - 1, 1).getValues().flat(); // Array Key Saja
  }

  const keysToSave = Object.keys(updates);
  const newConfigRows = [];
  
  keysToSave.forEach(key => {
      const val = updates[key];
      const rowIndex = existingData.indexOf(key);

      if (rowIndex !== -1) {
          sheet.getRange(rowIndex + 2, 10).setValue(val);
      } else {
          newConfigRows.push([key, val]);
          existingData.push(key); 
      }
  });
  // BATCH INSERT config baru
  if (newConfigRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 9, newConfigRows.length, 2).setValues(newConfigRows);
  }

  return { success: true, message: "Pengaturan berhasil disimpan!", config: updates };
}

// --- 5. SETTINGS & CONFIGURATION (FIXED) ---

function saveConfig(data) {
  const ss = _getSS();
  let sheet = ss.getSheetByName('Config');
  
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.appendRow(['Group', 'Key', 'Value']); 
  }

  let photoUrl = data.current_profile_img;

  // 1. PROSES GAMBAR (Dengan Error Handling Ketat)
  if (data.profile_img_base64 && data.profile_img_base64.length > 20) {
    try {
      // Bersihkan header "data:image/jpeg;base64," jika ada
      let base64String = data.profile_img_base64;
      if (base64String.includes('base64,')) {
        base64String = base64String.split('base64,')[1];
      }

      // Decode
      const decodedBytes = Utilities.base64Decode(base64String);
      const blob = Utilities.newBlob(decodedBytes, "image/jpeg", "PROFILE_" + Date.now() + ".jpg");

      // Simpan ke Drive
      const folderName = "Product_Images"; 
      const folders = DriveApp.getFoldersByName(folderName);
      const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      // Generate URL
      photoUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();

    } catch (e) {
      Logger.log("GAGAL UPLOAD: " + e.toString());
      // Jangan stop script, biarkan url lama terpakai
    }
  }

  // 2. SIMPAN DATA LAIN
  const configData = [
    { key: 'app_name', val: data.app_name },
    { key: 'owner_name', val: data.owner_name },
    { key: 'owner_role', val: data.owner_role },
    { key: 'owner_img', val: photoUrl }, // Pastikan URL masuk sini
    { key: 'limit_product', val: data.limit_product },
    { key: 'limit_material', val: data.limit_material },
    { key: 'page_title', val: data.page_title || 'Smart Inventory' },
    { key: 'sidebar_name', val: data.sidebar_name || 'SMART INVENTORY' },
    { key: 'sidebar_subtitle', val: data.sidebar_subtitle || 'BANTUSELLER V2.0' }
  ];

  const sheetData = sheet.getDataRange().getValues();
  const newCfgRows = [];
  
  configData.forEach(item => {
    let found = false;
    for (let i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][8]) === item.key) {
        sheet.getRange(i + 1, 10).setValue(item.val);
        found = true;
        break;
      }
    }
    if (!found) {
      newCfgRows.push([item.key, item.val]);
    }
  });
  // BATCH INSERT config baru
  if (newCfgRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 9, newCfgRows.length, 2).setValues(newCfgRows);
  }

  return { success: true, newImg: photoUrl };
}

// --- FUNGSI KHUSUS ACTIVITY LOGS ---
function getActivityLogs() {
  try {
    const ss = _getSS();
    let sheet = ss.getSheetByName('Data_Riwayat');
    
    // 1. Jika sheet belum ada, buat baru & isi header
    if (!sheet) {
      sheet = ss.insertSheet('Data_Riwayat');
      sheet.appendRow(['Waktu', 'ID Transaksi', 'Tipe', 'SKU', 'Nama Barang', 'Qty', 'Oleh', 'Keterangan']);
      // Return status kosong dulu
      return JSON.stringify({ status: 'EMPTY', data: [] });
    }

    const data = sheet.getDataRange().getValues();
    
    // 2. Jika hanya ada header (baris <= 1)
    if (data.length <= 1) {
       return JSON.stringify({ status: 'EMPTY', data: [] });
    }

    data.shift(); // Hapus header

    // 3. SANITASI DATA (PENTING!)
    // Kita ubah semua tanggal menjadi Teks ISO dan pastikan tidak ada null
    const cleanData = data.map(row => {
      let timeVal = row[0];
      // Cek apakah kolom pertama adalah objek Date
      if (timeVal && typeof timeVal.toISOString === 'function') {
        timeVal = timeVal.toISOString(); 
      } else {
        timeVal = String(timeVal); // Paksa jadi string kalau bukan tanggal
      }

      return [
        timeVal,              // 0. Waktu (String Aman)
        String(row[1] || '-'), // 1. ID
        String(row[2] || '-'), // 2. Tipe
        String(row[3] || ''),  // 3. SKU
        String(row[4] || ''),  // 4. Nama
        row[5] || 0,           // 5. Qty
        String(row[6] || 'System'), // 6. User
        String(row[7] || '')   // 7. Keterangan
      ];
    });

    // 4. Return sebagai JSON String (Anti Macet)
    return JSON.stringify({
      status: 'SUCCESS',
      data: cleanData.reverse().slice(0, 100)
    });

  } catch (e) {
    Logger.log("ERROR LOGS: " + e.toString());
    // Kirim pesan error ke frontend
    return JSON.stringify({ status: 'ERROR', message: e.toString() });
  }
}

/* =========================================
   7. USER MANAGEMENT & AUTH SYSTEM (FINAL)
   ========================================= */

// --- LOGIN SYSTEM ---
function doLogin(username, password, sheetId) {
  if (sheetId && sheetId.length > 10) {
    _ACTIVE_SHEET_ID = sheetId;
    // Simpan sebagai "last active" agar fungsi berikutnya pakai sheet ini
    PropertiesService.getScriptProperties().setProperty('bs_sheet_id', sheetId);
  }
  const ss = _getSS();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { success: false, message: "DB Users Missing" };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // [PENTING] Index 0 = Username, Index 2 = Password (Kolom C)
    // Jangan pakai data[i][1] karena itu Email!
    if (String(data[i][0]) === String(username) && String(data[i][2]) === String(password)) {
       
       if (data[i][4] === 'Inactive') {
          return { success: false, message: "Akun dinonaktifkan." };
       }

       return { 
        success: true, 
        user: {
          username: data[i][0],
          role: data[i][3],
          fullname: data[i][0],
          access: data[i][4]
        }
      };
    }
  }
  return { success: false, message: "Username/Password Salah." };
}

/* =========================================
   BACKEND: USER MANAGEMENT (Users Sheet)
   ========================================= */

// 1. Ambil List Users
function getUsersList() {
  const ss = _getSS();
  let sheet = ss.getSheetByName('Users');
  
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    // Header Baru
    sheet.appendRow(['Username', 'Email', 'Password', 'Role', 'Access']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    return [];
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; 
  
  // Ambil range A2:E
  return sheet.getRange(2, 1, lastRow - 1, 5).getValues();
}

// 2. Tambah User Baru (Update simpan Email)
function addNewUser(form) {
  const ss = _getSS();
  let sheet = ss.getSheetByName('Users');
  if (!sheet) {
      sheet = ss.insertSheet('Users');
      sheet.appendRow(['Username', 'Email', 'Password', 'Role', 'Access']);
  }
  
  // Validasi Username
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(form.username).toLowerCase()) {
      return { success: false, message: "Username sudah digunakan!" };
    }
  }
  
  // Simpan Data (5 Kolom)
  sheet.appendRow([
    form.username, 
    form.email,
    form.password, 
    form.role, 
    form.access 
  ]);
  
  return { success: true, message: "User berhasil disimpan" };
}

/// --- DELETE USER (SECURED) ---
function deleteUserByUsername(targetUser, requestor) {
  // Hanya OWNER yang boleh menghapus user lain
  if (!validateUserAccess(requestor, 'Owner')) {
     return { success: false, message: "AKSES DITOLAK! Hanya Owner yang boleh mengelola User." };
  }

  const ss = _getSS();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(targetUser)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: "User tidak ditemukan." };
}

// 4. Update User Existing
function updateUser(form) {
  const ss = _getSS();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const oldUsername = form.old_username; // Username lama sebagai kunci pencarian
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(oldUsername)) {
       // Cek apakah password diisi? Jika kosong, pakai password lama (Kolom C / index 2)
       const currentPassword = data[i][2];
       const newPassword = (form.password && form.password.trim() !== "") ? form.password : currentPassword;
       
       // Update baris (5 Kolom: Username, Email, Password, Role, Access)
       sheet.getRange(i + 1, 1, 1, 5).setValues([[
         form.username,
         form.email,
         newPassword,
         form.role,
         form.access
       ]]);
       
       return { success: true, message: "Data user berhasil diperbarui." };
    }
  }
  return { success: false, message: "User asli tidak ditemukan." };
}

// ==========================================
// --- SYSTEM UTILITIES (LOGOUT FIX) ---
// ==========================================

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/* 5. BACKEND LOGIN */
function verifyUser(username, password, sheetId) {
  try {
    if (sheetId && sheetId.length > 10) {
      _ACTIVE_SHEET_ID = sheetId;
      PropertiesService.getScriptProperties().setProperty('bs_sheet_id', sheetId);
    }
    const ss = _getSS();
    // Pastikan nama sheet ini sesuai dengan sheet database user Anda
    const sheet = ss.getSheetByName('Users'); 
    
    if (!sheet) return { success: false, message: 'Database User tidak ditemukan.' };

    const data = sheet.getDataRange().getValues(); // Ambil semua data

    // Loop dari baris ke-2 (index 1) karena baris 1 biasanya header
    // Asumsi Kolom: A=Username, B=Email (opsional), C=Password, D=Role
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dbUser = String(row[0]).trim(); // Kolom A
      const dbPass = String(row[2]).trim(); // Kolom C (Password) - Sesuaikan Index!

      if (dbUser === String(username).trim() && dbPass === String(password).trim()) {
        return { 
          success: true, 
          role: row[3], // Kolom D (Role)
          access: row[4], // Kolom E (Access JSON)
          message: 'Login berhasil.' 
        };
      }
    }

    return { success: false, message: 'Username atau password salah.' };

  } catch (error) {
    return { success: false, message: 'Error server: ' + error.message };
  }
}

// ==========================================
// --- FITUR AUTO ARCHIVING (ROBOT PEMBERSIH) ---
// ==========================================

/* Fungsi Utama yang akan dijalankan oleh Trigger (Jam 2 Pagi) */
function runAutoArchiving() {
  const ss = _getSS();
  const sheetPesanan = ss.getSheetByName('Data_Pesanan');
  const sheetRekap = ss.getSheetByName('Data_Rekap');
  
  if (!sheetPesanan || !sheetRekap) {
    Logger.log("Sheet Data_Pesanan atau Data_Rekap tidak ditemukan.");
    return;
  }

  // 1. Ambil Konfigurasi
  const config = getAppConfig(); // Menggunakan fungsi yang sudah ada
  const LIMIT_DAYS = Number(config.archive_days) || 60; // Default 60 hari
  const FOLDER_NAME = config.archive_folder_name || 'DATABASE_ARSIP_TOKO';
  
  // Hitung Tanggal Batas (Cut-off)
  // Contoh: Hari ini 1 Januari. Batas 60 hari lalu = 1 November.
  // Data sebelum 1 November akan diarsip.
  const cutOffDate = new Date();
  cutOffDate.setDate(cutOffDate.getDate() - LIMIT_DAYS);
  cutOffDate.setHours(0, 0, 0, 0);

  // 2. Baca Data Pesanan
  const data = sheetPesanan.getDataRange().getValues();
  const headers = data[0]; // Simpan header untuk file arsip baru
  
  // Kita butuh Index Kolom Penting (Sesuaikan dengan struktur Sheet Anda)
  // Kolom A=0 (ID), B=1 (Tgl), O=14 (Status), K=10 (Harga Jual/Omset), M=12 (HPP/Modal), J=9 (Qty)
  // [PENTING] Pastikan index ini sesuai dengan kolom di Data_Pesanan Anda
  const IDX_ID = 0;
  const IDX_TGL = 1;
  const IDX_QTY = 9;
  const IDX_OMSET = 19; // Kolom T (Total Marketplace) - atau sesuaikan logika omset Anda
  const IDX_MODAL = 13; // Kolom N (Total Modal)
  const IDX_STATUS = 14;

  let rowsToArchive = [];
  let indexesToDelete = []; // Kita simpan nomor baris untuk dihapus nanti

  // Loop Data (Mulai dari baris 2)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tglPesanan = new Date(row[IDX_TGL]);
    const status = String(row[IDX_STATUS]).toUpperCase();

    // Kriteria Arsip:
    // 1. Status SUDAH SELESAI (Shipped, Returned, Canceled, Completed)
    // 2. Tanggal Pesanan < CutOffDate (Lebih tua dari batas)
    const isFinished = ['SHIPPED', 'RETURNED', 'CANCELED', 'COMPLETED', 'SELESAI'].includes(status);
    
    if (isFinished && tglPesanan < cutOffDate) {
      rowsToArchive.push(row);
      indexesToDelete.push(i + 1); // Simpan index baris (1-based untuk deleteRow)
    }
  }

  if (rowsToArchive.length === 0) {
    Logger.log("Tidak ada data lama yang perlu diarsip.");
    return "Nihil";
  }

  // 3. Kelompokkan Data per Bulan (Agar masuk ke file Arsip yang tepat)
  // Map: "2025-10" -> [Row A, Row B, ...]
  const groups = {};
  
  rowsToArchive.forEach(row => {
    const tgl = new Date(row[IDX_TGL]);
    // Buat Key Bulan: "2025-10"
    const key = Utilities.formatDate(tgl, Session.getScriptTimeZone(), "yyyy-MM");
    
    if (!groups[key]) groups[key] = [];
    groups[key].push(row);
  });

  // 4. Proses Pindah ke Arsip & Rekap (Per Bulan)
  const folderArsip = getOrCreateFolder(FOLDER_NAME);

  for (const keyBulan in groups) {
    const groupRows = groups[keyBulan];
    
    // a. Hitung Total Rekap untuk bulan ini
    let totalOmset = 0;
    let totalModal = 0;
    let totalQty = 0;
    
    groupRows.forEach(r => {
      // Gunakan parseFloat jaga-jaga kalau string
      totalOmset += Number(r[IDX_OMSET]) || Number(r[11]) || 0; // Fallback ke Col L jika T kosong
      totalModal += Number(r[IDX_MODAL]) || 0;
      totalQty += Number(r[IDX_QTY]) || 0;
    });
    
    const totalProfit = totalOmset - totalModal;
    
    // b. Simpan ke Sheet 'Data_Rekap'
    const parts = keyBulan.split('-'); // [2025, 10]
    const namaBulanStr = getNamaBulan(parts[1]); // Helper function
    
    sheetRekap.appendRow([
      `REKAP-${keyBulan}`, // ID Unik
      namaBulanStr,        // Bulan
      parts[0],            // Tahun
      totalOmset,
      totalModal,
      totalProfit,
      totalQty,
      groupRows.length,    // Total Transaksi
      new Date()           // Waktu Arsip
    ]);

    // c. Simpan ke File Spreadsheet Arsip (Google Drive)
    const fileName = `Arsip_Pesanan_${namaBulanStr}_${parts[0]}`; // Contoh: Arsip_Pesanan_Oktober_2025
    const ssArsip = getOrCreateSpreadsheet(folderArsip, fileName, headers);
    const sheetArsip = ssArsip.getSheets()[0];
    
    // Append Rows (Batch)
    // Mulai dari baris terakhir + 1
    const lastRowArsip = sheetArsip.getLastRow();
    sheetArsip.getRange(lastRowArsip + 1, 1, groupRows.length, groupRows[0].length).setValues(groupRows);
  }

  // 5. Hapus Data dari Sheet Utama (Delete from Bottom Up)
  // Kita harus hapus dari bawah ke atas agar index baris tidak bergeser saat dihapus
  indexesToDelete.sort((a, b) => b - a); // Urutkan besar ke kecil (Bawah ke Atas)
  
  indexesToDelete.forEach(rowIndex => {
    sheetPesanan.deleteRow(rowIndex);
  });

  Logger.log(`Sukses mengarsipkan ${rowsToArchive.length} baris data.`);
}

// --- HELPER FUNCTIONS (Otak Cadangan) ---

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

function getOrCreateSpreadsheet(folder, fileName, headers) {
  // Cek apakah file sudah ada di folder tersebut
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  } else {
    // Buat Spreadsheet Baru
    const ssNew = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(ssNew.getId());
    
    // Pindahkan file ke folder arsip (karena default create ada di Root Drive)
    file.moveTo(folder); 
    
    // Pasang Header
    ssNew.getSheets()[0].appendRow(headers);
    return ssNew;
  }
}

function getNamaBulan(monthIndex) {
  // monthIndex string "01" sampai "12"
  const idx = parseInt(monthIndex) - 1;
  const names = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return names[idx] || monthIndex;
}

// ==========================================
// --- FITUR AUTO ARSIP LOG (ROBOT PEMBERSIH) ---
// ==========================================

/**
 * Fungsi Utama: Memindahkan log lama ke file Spreadsheet terpisah
 * parameter 'testDays' opsional. Jika diisi, akan mengabaikan setting default 30 hari.
 */
function runLogArchiving(testDays) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Kunci script agar tidak bentrok
  } catch (e) {
    Logger.log('Script sibuk.');
    return;
  }

  try {
    const ss = _getSS();
    const sheet = ss.getSheetByName('Data_Riwayat');
    if (!sheet) return;

    // 1. TENTUKAN BATAS WAKTU
    // Jika testDays ada (misal 0), pakai itu. Jika tidak, default 30 hari.
    const LIMIT_DAYS = (testDays !== undefined) ? testDays : 30; 
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - LIMIT_DAYS);
    cutoffDate.setHours(0, 0, 0, 0); // Set ke awal hari

    Logger.log(`Menjalankan Arsip. Batas Tanggal: ${cutoffDate} (Lebih tua dari ${LIMIT_DAYS} hari)`);

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return; // Cuma header

    // 2. CARI BARIS YANG PERLU DIARSIP (Urut dari atas/terlama)
    // Asumsi: Data Log selalu urut waktu (Baris 2 adalah terlama)
    // Kita cari sampai baris berapa data itu < cutoffDate
    
    let lastRowToArchive = 0; 
    const rowsToMove = [];

    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][0]); // Kolom A = Waktu
      
      if (rowDate < cutoffDate) {
        rowsToMove.push(data[i]);
        lastRowToArchive = i;
      } else {
        // Karena data urut, begitu ketemu tanggal yg baru, stop.
        break; 
      }
    }

    if (rowsToMove.length === 0) {
      Logger.log("Tidak ada data lama yang perlu diarsip.");
      return "NIHIL";
    }

    // 3. PINDAHKAN KE FILE ARSIP EKSTERNAL
    const folderName = "DATABASE_ARSIP_LOGS";
    const year = cutoffDate.getFullYear();
    const fileName = `Arsip_Log_Activity_${year}`; // Contoh: Arsip_Log_Activity_2025
    
    const targetSS = getOrCreateArchiveLogFile(folderName, fileName, data[0]); // data[0] = Header
    const targetSheet = targetSS.getSheets()[0];

    // Simpan ke file arsip
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove);

    // 4. HAPUS DARI SHEET UTAMA (Batch Delete)
    // Hapus dari baris 2 sebanyak jumlah baris yg dipindah
    sheet.deleteRows(2, rowsToMove.length);

    const msg = `Sukses! ${rowsToMove.length} baris log dipindahkan ke file '${fileName}'.`;
    Logger.log(msg);
    return msg;

  } catch (e) {
    Logger.log("Error Archiving: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// Helper: Cari/Buat File Arsip di Folder Khusus
function getOrCreateArchiveLogFile(folderName, fileName, headers) {
  const folders = DriveApp.getFoldersByName(folderName);
  let folder;
  
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  } else {
    const newSS = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(newSS.getId());
    file.moveTo(folder); // Pindah ke folder arsip
    
    // Bikin Header di file baru
    newSS.getSheets()[0].appendRow(headers);
    return newSS;
  }
}

// ==========================================
// --- TOOLS UNTUK TESTING (JANGAN DIHAPUS DULU) ---
// ==========================================

// 1. GENERATE DATA PALSU JADUL (Untuk Testing)
function test_GenerateOldLogs() {
  const ss = _getSS();
  let sheet = ss.getSheetByName('Data_Riwayat');
  if(!sheet) sheet = ss.insertSheet('Data_Riwayat');

  const oldDate = new Date();
  oldDate.setDate(oldDate.getDate() - 60); // Mundur 60 Hari (2 Bulan lalu)
  
  // Buat 5 baris data "jadul"
  const dummyData = [];
  for(let i=1; i<=5; i++) {
    dummyData.push([
      oldDate, 
      `TEST-OLD-${i}`, 
      'TESTING', 
      'SKU-TEST', 
      'Barang Test Jadul', 
      1, 
      'tester@bot.com', 
      'Ini data palsu 60 hari lalu'
    ]);
  }
  
  // Masukkan ke baris paling ATAS (setelah header) agar urutan waktu valid
  // (Karena log biasanya appendRow di bawah, tapi data lama harusnya di atas)
  // Untuk simulasi, kita insertRowsAfter header (baris 1)
  sheet.insertRowsAfter(1, 5);
  sheet.getRange(2, 1, 5, 8).setValues(dummyData);
  
  SpreadsheetApp.getUi().alert(`Berhasil membuat 5 baris data tanggal ${oldDate.toLocaleDateString()} di baris atas.`);
}

// 2. TOMBOL RUN MANUAL (Hapus data > 1 hari)
// Ini mensimulasikan "Saya tidak mau nunggu 30 hari, arsipkan yg kemarin saja"
function test_RunArchiveNow() {
  const result = runLogArchiving(1); // Angka 1 artinya: Arsip data yang lebih tua dari 1 hari
  
  if(result === "NIHIL") {
    SpreadsheetApp.getUi().alert("Tidak ada data lama (lebih dari 1 hari) ditemukan.");
  } else {
    SpreadsheetApp.getUi().alert(result + "\n\nSilakan cek Google Drive folder 'DATABASE_ARSIP_LOGS'.");
  }
}

// ==========================================
// --- SETUP OTOMATIS (UNTUK USER BARU) ---
// ==========================================

// 1. MENU KHUSUS SAAT FILE DIBUKA
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ MENU ADMIN')
    .addItem('▶️ Aktifkan Robot Otomatis', 'setupAppTriggers')
    .addSeparator()
    .addItem('🗑️ Bersihkan Log Manual', 'test_RunArchiveNow')
    .addSeparator()
    .addItem('🏷️ Set Header Admin Shopee (Kolom V)', 'setupAdminShopeeHeader')
    .addToUi();
}

// Set header "Admin Shopee" & "Admin Tiktok" di kolom V (22) & W (23) Master_Produk
function setupAdminShopeeHeader() {
  var sheet = _getSS().getSheetByName('Master_Produk');
  if (!sheet) return;
  sheet.getRange(1, 22).setValue('Admin Shopee');
  sheet.getRange(1, 23).setValue('Admin Tiktok');
  SpreadsheetApp.getActiveSpreadsheet().toast('Header "Admin Shopee" (V) & "Admin Tiktok" (W) berhasil diset!', 'Berhasil', 3);
}

// 2. FUNGSI PEMBUAT TRIGGER (ONE CLICK SETUP)
function setupAppTriggers() {
  const ui = SpreadsheetApp.getUi();
  const userResponse = ui.alert(
    'Konfirmasi Setup',
    'Apakah Anda ingin mengaktifkan robot pembersih otomatis (Auto-Archive)?\nScript akan berjalan setiap hari jam 02:00 pagi.',
    ui.ButtonSet.YES_NO
  );

  if (userResponse == ui.Button.NO) return;

  // Hapus semua trigger lama dulu (biar tidak dobel-dobel kalau diklik berkali-kali)
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'runLogArchiving' || fn === 'runAutoArchiving' || fn === 'syncShopeeTracking') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // --- TRIGGER 1: BERSIHKAN LOG AKTIFITAS (Data_Riwayat) ---
  ScriptApp.newTrigger('runLogArchiving')
      .timeBased()
      .everyDays(1)   // Setiap 1 Hari
      .atHour(2)      // Jam 2 Pagi
      .create();

  // --- TRIGGER 2: ARSIP DATA PESANAN LAMA (Data_Pesanan) ---
  // (Pastikan fungsi 'runAutoArchiving' sudah ada di script Anda sebelumnya)
  // Jika belum ada, fungsi ini tidak perlu ditambahkan trigger-nya
  try {
     ScriptApp.newTrigger('runAutoArchiving')
      .timeBased()
      .everyDays(1)
      .atHour(2)
      .create();
  } catch(e) {
    // Abaikan jika fungsi archiving pesanan belum dibuat
  }

  // --- TRIGGER 3: SYNC TRACKING SHOPEE (Setiap 6 Jam) ---
  ScriptApp.newTrigger('syncShopeeTracking')
      .timeBased()
      .everyHours(6)
      .create();

  ui.alert('SUKSES! ✅\n\nRobot otomatis telah diaktifkan.\n- Arsip: setiap jam 2 pagi\n- Sync Tracking Shopee: setiap 6 jam');
}

// ==========================================
// --- FUNGSI UPDATE ORDER (DENGAN VALIDASI) ---
// ==========================================

function updateOrderData(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Mencegah bentrok edit
  } catch (e) {
    return { success: false, message: "Server sibuk. Coba lagi." };
  }

  try {
    const ss = _getSS();
    const sheetOrder = ss.getSheetByName('Data_Pesanan');
    const sheetProduk = ss.getSheetByName('Master_Produk');
    
    // 1. CARI PESANAN LAMA
    const dataOrder = sheetOrder.getDataRange().getValues();
    let rowIndex = -1;
    let oldData = null;

    // Loop cari ID Pesanan (Kolom A)
    for (let i = 1; i < dataOrder.length; i++) {
      if (String(dataOrder[i][0]) === String(form.idPesanan)) {
        rowIndex = i + 1;
        oldData = dataOrder[i];
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "ID Pesanan tidak ditemukan!" };

    // 2. VALIDASI STATUS (PENTING!)
    // Jika status sudah PACKED/SHIPPED, tolak edit Qty/SKU karena barang fisik sudah diproses.
    const currentStatus = String(oldData[14]).toUpperCase(); // Kolom O (Index 14)
    const isProcessed = (currentStatus === 'PACKED' || currentStatus === 'SHIPPED' || currentStatus === 'RETURNED');
    
    // Cek apakah User mengubah data sensitif (SKU atau Qty)
    const newQty = Number(form.qty);
    const oldQty = Number(oldData[9]);
    const isQtyChanged = (newQty !== oldQty);
    
    // Jika sudah diproses & user coba ganti Qty -> TOLAK
    if (isProcessed && isQtyChanged) {
       return { 
         success: false, 
         message: `GAGAL: Pesanan ini statusnya sudah '${currentStatus}'. Anda tidak boleh mengubah Qty karena barang sudah/sedang diproses. Silakan 'Batalkan Scan' dulu di menu Gudang.` 
       };
    }

    // 3. VALIDASI STOK (Hanya jika status masih WAITING)
    if (!isProcessed && isQtyChanged) {
        // Ambil Data Stok Terkini dari Master
        const targetSku = form.skuVariasi || form.skuInduk; // Prioritas Varian
        const prodData = getProductBySku(targetSku); // Helper yang sudah kita perbaiki
        
        if (!prodData.found) {
           return { success: false, message: `SKU '${targetSku}' tidak ditemukan di Master Produk.` };
        }

        const stokGudang = Number(prodData.qty);

        // LOGIKA PENGECEKAN:
        // Apakah stok gudang cukup untuk permintaan baru ini?
        // Ingat: Stok di Master adalah "Sisa Fisik di Rak".
        // Karena status pesanan ini masih 'WAITING', berarti dia BELUM mengambil barang fisik.
        // Jadi kita cukup cek: Apakah Stok Rak >= Qty Baru?
        
        if (stokGudang < newQty) {
           return { 
             success: false, 
             message: `STOK TIDAK CUKUP! Stok fisik hanya tersedia ${stokGudang} pcs, Anda meminta ${newQty} pcs.` 
           };
        }
    }

    // 4. HITUNG ULANG HARGA (Jaga-jaga harga berubah)
    const hargaJual = Number(form.hargaJual) || Number(oldData[10]);
    const totalJual = newQty * hargaJual;
    const diskon = Number(form.diskon) || 0;
    const totalNet = totalJual - diskon;

    // Ambil HPP Lama (Agar profit tidak lari)
    const hppSatuan = Number(oldData[12]); 
    const totalModal = newQty * hppSatuan;

    // 5. SIMPAN PERUBAHAN
    // Update Kolom Tertentu Saja:
    // [1]Tgl, [2]Channel, [3]Kurir, [4]Resi, [6]Nama, [8]Varian, [9]Qty, [11]Total, [13]Modal, [18]Diskon, [19]Net
    
    // BATCH: Update Metadata (C, D, E = kolom 3-5)
    sheetOrder.getRange(rowIndex, 3, 1, 3).setValues([[form.channel, form.kurir, form.resi]]);
    
    // Update Data Sensitif
    sheetOrder.getRange(rowIndex, 10).setValue(newQty);       // Col J (Qty)
    sheetOrder.getRange(rowIndex, 12).setValue(totalJual);    // Col L (Total Jual)
    sheetOrder.getRange(rowIndex, 14).setValue(totalModal);   // Col N (Total Modal)
    // BATCH: S+T (kolom 19-20)
    sheetOrder.getRange(rowIndex, 19, 1, 2).setValues([[diskon, totalNet]]);

    // Catat Log Aktifitas
    const sLogActivity = ss.getSheetByName('Data_Riwayat');
    if (sLogActivity) {
       const user = Session.getActiveUser().getEmail();
       sLogActivity.appendRow([
          new Date(), form.idPesanan, 'EDIT ORDER', targetSku, 
          form.namaProduk, newQty, user, 
          `Edit Data: Qty ${oldQty} -> ${newQty}`
       ]);
    }

    return { success: true, message: "Pesanan berhasil diperbarui!" };

  } catch (e) {
    return { success: false, message: "Error Update: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function test_EditOrder() {
   const result = updateOrderData({
      idPesanan: "ID-PESANAN-YANG-BARUSAN-DIBUAT", // Ganti ID asli
      qty: 100, // Sengaja bikin banyak biar gagal
      hargaJual: 50000,
      skuVariasi: "SKU-BARANGNYA"
   });
   Logger.log(result);
}

// --- BACKEND: AMBIL DETAIL ORDER LENGKAP (FIX GAMBAR & KECAMATAN) ---
function getFullOrderDetails(orderId) {
  try {
    const ss = _getSS();
    const sheet = ss.getSheetByName('Data_Pesanan');
    const sheetProd = ss.getSheetByName('Master_Produk'); // Perlu ini untuk ambil gambar
    
    if (!sheet) return { success: false, message: "Sheet 'Data_Pesanan' tidak ditemukan!" };

    // 1. BUAT MAP GAMBAR PRODUK (SKU -> URL)
    // Supaya saat edit, fotonya muncul
    const imgMap = {};
    if (sheetProd) {
       const dataProd = sheetProd.getDataRange().getValues();
       for (let p = 1; p < dataProd.length; p++) {
          const sku = String(dataProd[p][3]).trim(); // SKU Variasi
          const imgVar = dataProd[p][14]; // Foto Variasi
          const imgMain = dataProd[p][13]; // Foto Utama
          // Prioritas: Foto Varian -> Foto Utama
          if (sku) imgMap[sku] = imgVar || imgMain || "";
       }
    }

    const data = sheet.getDataRange().getValues();
    
    let orderHeader = null;
    let items = [];

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]).trim();
      const targetId = String(orderId).trim();

      if (rowId === targetId) {
        
        // Ambil Header (Sekali saja)
        if (!orderHeader) {
          let tglSafe = data[i][1];
          if (tglSafe instanceof Date) {
             const offset = tglSafe.getTimezoneOffset() * 60000;
             tglSafe = (new Date(tglSafe - offset)).toISOString().split('T')[0];
          } else {
             tglSafe = String(tglSafe);
          }

          // [PERBAIKAN KECAMATAN GABUNGAN]
          // Data tersimpan terpisah di Z(25), AA(26), AB(27)
          // Kita gabung lagi jadi satu string: "Kecamatan, Kota, Provinsi"
          let locFull = "";
          if (data[i][25]) {
             locFull = data[i][25]; // Kecamatan
             if (data[i][26]) locFull += `, ${data[i][26]}`; // Kota
             if (data[i][27]) locFull += `, ${data[i][27]}`; // Provinsi
          }

          orderHeader = {
            id: rowId,
            tgl: tglSafe, 
            channel: data[i][2],
            kurir: data[i][3],
            resi: data[i][4],
            status: data[i][14],
            
            // Data Customer
            custName: data[i][21] || '',
            custPhone: data[i][22] || '',
            custAddress: data[i][23] || '',
            ongkir: Number(data[i][24]) || 0,
            
            // Ini yang akan masuk ke field search kecamatan
            kecamatan: locFull 
          };
        }

        // Ambil Item
        const skuItem = String(data[i][7] || data[i][5]).trim(); // SKU Variasi atau Induk
        
        items.push({
          skuInduk: data[i][5],
          namaProduk: data[i][6],
          skuVariasi: data[i][7],
          variasi: data[i][8],
          qty: Number(data[i][9]),
          hargaJual: Number(data[i][10]),
          hpp: Number(data[i][12]), 
          diskon: Number(data[i][18]),
          
          // [PERBAIKAN] Ambil Gambar dari Map yang sudah kita buat
          img: imgMap[skuItem] || "" 
        });
      }
    }

    if (!orderHeader) {
      return { success: false, message: "ID Pesanan tidak ditemukan." };
    }

    return { success: true, header: orderHeader, items: items };

  } catch (err) {
    Logger.log("Error: " + err);
    return { success: false, message: "Error Backend: " + err.toString() };
  }
}

// --- BACKEND: PROSES EDIT FULL (DENGAN VALIDASI STOK KETAT) ---
function processFullOrderEdit(formObject) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: "Server sibuk. Coba lagi." };
  }

  try {
    const ss = _getSS();
    const sheet = ss.getSheetByName('Data_Pesanan');
    const orderId = String(formObject.idPesanan);
    
    // 1. VALIDASI STATUS PESANAN
    const data = sheet.getDataRange().getValues();
    let oldRowsIndices = [];
    let currentStatus = 'WAITING';
    
    // Cari semua baris milik ID ini
    for (let i = 1; i < data.length; i++) {
       if (String(data[i][0]) === orderId) {
          oldRowsIndices.push(i + 1); 
          currentStatus = String(data[i][14]).toUpperCase();
       }
    }

    if (oldRowsIndices.length === 0) return { success: false, message: "Order tidak ditemukan." };

    // Cek Status: Hanya boleh edit jika WAITING
    // Jika sudah dipacking/dikirim, stok fisik sudah bergerak, jadi edit dilarang.
    if (currentStatus === 'PACKED' || currentStatus === 'SHIPPED' || currentStatus === 'RETURNED') {
       return { success: false, message: `GAGAL: Pesanan ini statusnya sudah '${currentStatus}'. Stok sudah diproses. Silakan batalkan scan gudang dulu jika ingin revisi.` };
    }

    // 2. VALIDASI STOK (SAFETY CHECK)
    // Kita baca Master Produk dulu untuk tahu sisa stok saat ini
    const sheetProd = ss.getSheetByName('Master_Produk');
    const dataProd = sheetProd.getDataRange().getValues();
    const stockMap = new Map();

    // Mapping: SKU (Col D/Index 3) -> {Stok: Col H/Index 7, Nama: Col B/Index 1}
    for(let p=1; p<dataProd.length; p++) {
        const pSku = String(dataProd[p][3]).trim();
        const pStok = Number(dataProd[p][7]); // Stok Sisa Fisik
        const pNama = dataProd[p][1];
        stockMap.set(pSku, { stok: pStok, name: pNama });
    }

    const items = formObject.items;
    
    // Loop Cek Setiap Item yang mau disimpan
    for (let item of items) {
        // Prioritaskan SKU Variasi, kalau gak ada pakai SKU Induk
        const checkSku = String(item.skuVariasi || item.skuInduk).trim();
        const askQty = Number(item.qty);
        
        if (stockMap.has(checkSku)) {
            const prod = stockMap.get(checkSku);
            
            // LOGIKA: Karena status WAITING, berarti stok fisik di master BELUM terpotong oleh pesanan ini.
            // Jadi "Sisa Stok Master" adalah batas maksimal yang bisa diambil.
            if (prod.stok < askQty) {
                return { 
                    success: false, 
                    message: `STOK TIDAK CUKUP! Produk '${prod.name}' hanya sisa ${prod.stok}, Anda meminta ${askQty}.` 
                };
            }
        } else {
             // Jika SKU tidak ada di master
             return { success: false, message: `Error: SKU '${checkSku}' tidak ditemukan di Master Produk.` };
        }
    }

    // --- JIKA LOLOS VALIDASI, LANJUT SIMPAN ---

    // 3. HAPUS BARIS LAMA
    oldRowsIndices.sort((a, b) => b - a); // Hapus dari bawah ke atas
    oldRowsIndices.forEach(r => sheet.deleteRow(r));

    // 4. INSERT DATA BARU
    const timestamp = formObject.oldDate ? new Date(formObject.oldDate) : new Date(); // Pertahankan tanggal asli
    const user = Session.getActiveUser().getEmail();
    const newRows = [];

    items.forEach(item => {
        const qty = Number(item.qty) || 0;
        const hargaJual = Number(item.hargaJual) || 0;
        const diskon = Number(item.diskon) || 0;
        const hpp = Number(item.hargaBeli) || 0; // Pastikan ambil HPP yang dikirim frontend
        
        const totalJual = qty * hargaJual;
        const totalNet = totalJual - diskon;
        const totalModal = qty * hpp;

        newRows.push([
          orderId,                    // A (ID Tetap)
          timestamp,                  // B (Tgl Asli)
          formObject.channel,         // C
          formObject.kurir,           // D
          formObject.resi,            // E
          item.skuInduk,              // F
          item.namaProduk,            // G
          item.skuVariasi,            // H
          item.variasi,               // I
          qty,                        // J
          hargaJual,                  // K
          totalJual,                  // L
          hpp,                        // M
          totalModal,                 // N
          'WAITING',                  // O (Tetap Waiting)
          '', '', '',                 // P-R
          diskon,                     // S
          totalNet                    // T
        ]);
    });

    if (newRows.length > 0) {
       sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    // Log Activity
    const sLog = ss.getSheetByName('Data_Riwayat');
    if(sLog) sLog.appendRow([new Date(), orderId, 'EDIT ORDER', '-', 'Multi Item', items.length, user, 'Edit Pesanan (Re-Write)']);

    return { success: true, message: "Pesanan berhasil diperbarui!" };

  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// --- FITUR UNDO / PEMBATALAN TRANSAKSI ---
// ==========================================

function undoLastTransaction(barcode) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: "Server sibuk." };
  }

  try {
    const ss = _getSS();
    const sheetStok = ss.getSheetByName('Riwayat_Stok');
    const sheetOrder = ss.getSheetByName('Data_Pesanan');
    const sheetMaster = ss.getSheetByName('Master_Produk');

    // 1. CARI TRANSAKSI TERAKHIR DI RIWAYAT STOK
    // Kita cari dari bawah ke atas agar menemukan yang paling baru
    const dataStok = sheetStok.getDataRange().getValues();
    let logRowIndex = -1;
    let logData = null;

    for (let i = dataStok.length - 1; i >= 1; i--) {
      // Kolom D (Index 3) adalah Referensi/Resi
      // Kolom C (Index 2) adalah Tipe (IN/OUT)
      // Kita cari yang resinya cocok DAN belum dicancel
      if (String(dataStok[i][3]).trim() === String(barcode).trim() && 
          !String(dataStok[i][2]).includes('CANCEL')) {
         logRowIndex = i + 1;
         logData = dataStok[i];
         break;
      }
    }

    if (logRowIndex === -1) {
      return { success: false, message: "Tidak ditemukan riwayat transaksi aktif untuk resi ini." };
    }

    const type = logData[2]; // Tipe: OUT (SALES), PACKED, dll
    const sku = logData[4];  // SKU
    const qty = Number(logData[7]); // Qty yang dulu dikeluarkan

    // 2. PROSES UNDO BERDASARKAN TIPE
    
    // --- KASUS A: BATAL KIRIM (UNDO SCAN OUT) ---
    if (type.includes('OUT') || type.includes('SHIPPING')) {
       
       // A1. Kembalikan Stok Master
       const dataMaster = sheetMaster.getDataRange().getValues();
       for (let m = 1; m < dataMaster.length; m++) {
          if (String(dataMaster[m][3]) === String(sku)) {
             const row = m + 1;
             const newKeluar = Number(dataMaster[m][6]) - qty;
             const newSisa = Number(dataMaster[m][7]) + qty;
             const hpp = Number(dataMaster[m][9]);

             // BATCH: G+H sekaligus, lalu K
             sheetMaster.getRange(row, 7, 1, 2).setValues([[newKeluar, newSisa]]);
             sheetMaster.getRange(row, 11).setValue(newSisa * hpp);
             break;
          }
       }

       // A2. Kembalikan Status Pesanan (Jadi WAITING - karena AUTO mode skip PACKED)
       const dataOrder = sheetOrder.getDataRange().getValues();
       for (let o = 1; o < dataOrder.length; o++) {
          if (String(dataOrder[o][4]) === String(barcode)) { // Col E (Resi)
             sheetOrder.getRange(o + 1, 15).setValue('WAITING'); 
             // Hapus Tanggal Packing (Col P / 16) dan Kirim (Col Q / 17)
             sheetOrder.getRange(o + 1, 16).setValue(''); 
             sheetOrder.getRange(o + 1, 17).setValue(''); 
          }
       }

       // A3. Tandai Log sebagai Batal
       sheetStok.getRange(logRowIndex, 3).setValue('CANCELLED (OUT)');
       sheetStok.getRange(logRowIndex, 12).setValue('Transaksi Dibatalkan User');

       return { success: true, message: "✅ Scan DIBATALKAN. Stok dikembalikan (+"+qty+"). Status: WAITING." };
    }

    // --- KASUS B: BATAL PACKING (UNDO SCAN IN) ---
    else if (type === 'PACKING' || type === 'PACKED') {
       // Packing tidak mengurangi stok, jadi cuma ubah status pesanan
       const dataOrder = sheetOrder.getDataRange().getValues();
       for (let o = 1; o < dataOrder.length; o++) {
          if (String(dataOrder[o][4]) === String(barcode)) {
             sheetOrder.getRange(o + 1, 15).setValue('WAITING'); // Balik ke Waiting
             sheetOrder.getRange(o + 1, 16).setValue(''); // Hapus tgl packing
             break;
          }
       }
       
       // Update Log (Jika ada log packing)
       sheetStok.getRange(logRowIndex, 3).setValue('CANCELLED (PACKING)');
       
       return { success: true, message: "✅ Packing DIBATALKAN. Status pesanan kembali 'WAITING'." };
    }

    return { success: false, message: "Tipe transaksi tidak didukung untuk Undo otomatis." };

  } catch (e) {
    return { success: false, message: "Error Undo: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// --- SECURITY & PERMISSIONS HELPER ---
// ==========================================

/**
 * Fungsi Satpam: Mengecek apakah User punya wewenang.
 * @param {string} username - Username yang melakukan request
 * @param {string} minRole - Role minimal yang dibutuhkan ('Admin', 'Owner')
 * @return {boolean} - True jika diizinkan
 */
function validateUserAccess(username, minRole) {
  const ss = _getSS();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  // Normalisasi Role yang diminta (Huruf kecil & trim)
  const reqRole = String(minRole).trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    // Bandingkan Username (Kolom A) dengan input username
    if (String(data[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      
      const userRole = String(data[i][3]).trim().toLowerCase(); // Ambil Role User (Kolom D)

      // 1. Jika butuh OWNER, izinkan 'owner', 'super admin', atau 'pemilik'
      if (reqRole === 'owner') {
         return userRole === 'owner' || userRole === 'super admin' || userRole === 'pemilik';
      }
      
      // 2. Jika butuh ADMIN, izinkan 'admin', 'owner', dll
      if (reqRole === 'admin') {
         return userRole === 'admin' || userRole === 'owner' || userRole === 'super admin' || userRole === 'pemilik';
      }
      
      return true; // Untuk role Staff/Umum
    }
  }
  return false; // User tidak ditemukan
}

/* ===============================================================
   BACKEND: DATABASE WILAYAH INDONESIA (FIXED & ROBUST)
   Sumber: Edwardsamuel (Stable)
   =============================================================== */

/**
 * FUNGSI 1: GENERATOR DATA (Jalankan ini sekali saja!)
 * Fungsi ini akan men-download data resmi Indonesia dan menyimpannya ke Sheet.
 */
function setupDatabaseWilayah() {
  const ss = _getSS();
  let sheet = ss.getSheetByName('Master_Wilayah');
  
  // 1. Siapkan Sheet
  if (!sheet) {
    sheet = ss.insertSheet('Master_Wilayah');
  }
  sheet.clear();
  sheet.appendRow(['Kecamatan', 'Kota/Kabupaten', 'Provinsi']); // Header
  
  // Format Header
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  try {
    // URL Sumber Data (Edwardsamuel - Stabil)
    const BASE_URL = 'https://raw.githubusercontent.com/edwardsamuel/Wilayah-Administratif-Indonesia/master/csv';
    
    // Helper untuk Fetch agar aman
    const fetchData = (filename) => {
      const url = `${BASE_URL}/${filename}`;
      const options = { 'muteHttpExceptions': true };
      const resp = UrlFetchApp.fetch(url, options);
      if (resp.getResponseCode() !== 200) throw new Error(`Gagal download ${filename}`);
      return resp.getContentText();
    };

    // 2. Download & Mapping Provinsi
    // Format: [id, name]
    const csvProv = fetchData('provinces.csv');
    const mapProv = {};
    parseCsvCustom(csvProv).forEach(row => {
      if(row.length >= 2) mapProv[row[0]] = formatNama(row[1]);
    });

    // 3. Download & Mapping Kota/Kab
    // Format: [id, province_id, name]
    const csvReg = fetchData('regencies.csv');
    const mapReg = {};
    parseCsvCustom(csvReg).forEach(row => {
      if(row.length >= 3) {
        mapReg[row[0]] = {
          name: formatNama(row[2]),
          prov: mapProv[row[1]] || 'Unknown'
        };
      }
    });

    // 4. Download & Susun Kecamatan
    // Format: [id, regency_id, name]
    const csvDist = fetchData('districts.csv');
    const finalData = [];
    
    parseCsvCustom(csvDist).forEach(row => {
      if(row.length >= 3) {
        const regInfo = mapReg[row[1]];
        if (regInfo) {
          finalData.push([
            formatNama(row[2]), // Nama Kecamatan
            regInfo.name,       // Nama Kota
            regInfo.prov        // Nama Provinsi
          ]);
        }
      }
    });

    // 5. Simpan ke Sheet (Batch Write)
    if (finalData.length > 0) {
      // Tulis per 2000 baris agar tidak timeout jika data banyak
      const chunkSize = 2000;
      for (let i = 0; i < finalData.length; i += chunkSize) {
        const chunk = finalData.slice(i, i + chunkSize);
        sheet.getRange(sheet.getLastRow() + 1, 1, chunk.length, 3).setValues(chunk);
      }
      return `BERHASIL! ${finalData.length} data kecamatan telah tersimpan.`;
    } else {
      return "GAGAL: Data kosong setelah diproses.";
    }

  } catch (e) {
    return "ERROR CRITICAL: " + e.toString();
  }
}

/**
 * Helper: Parser CSV Manual yang lebih aman
 * Menangani koma, titik koma, dan new line
 */
function parseCsvCustom(text) {
  if (!text) return [];
  const lines = text.split('\n');
  const result = [];
  
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    if (!line) continue;
    
    // Deteksi pemisah (Koma atau Titik Koma)
    const separator = line.indexOf(';') > -1 ? ';' : ',';
    
    // Bersihkan tanda kutip jika ada
    const parts = line.split(separator).map(p => p.replace(/^"|"$/g, '').trim());
    result.push(parts);
  }
  return result;
}

/**
 * Helper: Format Nama (Title Case)
 * CONTOH: "DKI JAKARTA" -> "Dki Jakarta" (Biar rapi)
 */
function formatNama(str) {
  if(!str) return "";
  return str.toLowerCase().replace(/\b\w/g, s => s.toUpperCase());
}

/**
 * FUNGSI 2: PENCARIAN (Digunakan Frontend)
 */
function apiSearchKecamatan(query) {
  if (!query || query.length < 3) return [];

  const ss = _getSS();
  const sheet = ss.getSheetByName('Master_Wilayah');
  
  if (!sheet) return [{ id: null, label: "Jalankan 'setupDatabaseWilayah' dulu!" }];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [{ id: null, label: "Data wilayah kosong." }];

  // Ambil semua data (Caching agar cepat)
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const qLower = query.toLowerCase();
  const results = [];
  
  // Cari maksimal 15 hasil
  for (let i = 0; i < data.length; i++) {
    const kec = String(data[i][0]);
    const kota = String(data[i][1]);
    const prov = String(data[i][2]);
    
    if (kec.toLowerCase().includes(qLower) || kota.toLowerCase().includes(qLower)) {
      results.push({
        id: "0", 
        label: `${kec}, ${kota}, ${prov}`
      });
    }
    if (results.length >= 15) break;
  }

  return results.length > 0 ? results : [{ id: null, label: "Tidak ditemukan." }];
}

// Dummy Ongkir (Biar frontend gak error)
function apiCheckOngkir() { return 0; }

// Hitung Profit Bersih dari Data Pesanan + Master Produk
function getProfitData(startDate, endDate) {
  var ss = _getSS();
  var sheetPesanan = ss.getSheetByName('Data_Pesanan');
  var sheetProduk = ss.getSheetByName('Master_Produk');
  if (!sheetPesanan || !sheetProduk) return { penghasilan: 0, hpp: 0, profit: 0, orders: 0 };

  // Cache admin% + hpp dari Master_Produk
  // Key bisa SKU Variasi atau SKU Induk + Variasi
  var prodData = sheetProduk.getDataRange().getValues();
  var prodMap = {}; // key = SKU Variasi
  for (var p = 1; p < prodData.length; p++) {
    var sku = String(prodData[p][3]).trim();
    if (sku) {
      prodMap[sku.toLowerCase()] = {
        adminPct: Number(prodData[p][21]) || 0,  // Kolom V = Admin%
        hpp: Number(prodData[p][9]) || 0,         // Kolom J = HPP per unit
        jual: Number(prodData[p][11]) || 0        // Kolom L = Harga Jual
      };
    }
  }

  var data = sheetPesanan.getDataRange().getValues();
  var start = startDate ? new Date(startDate) : null;
  var end = endDate ? new Date(endDate) : null;
  if (end) { end.setHours(23, 59, 59); }

  var totalJual = 0;
  var totalAdmin = 0;
  var totalHpp = 0;
  var orderCount = 0;

  // Data perjalanan (SHIPPED) — belum masuk profit
  var onTheWayJual = 0;
  var onTheWayAdmin = 0;
  var onTheWayCount = 0;

  for (var i = 1; i < data.length; i++) {
    // FILTER 1: Hanya channel SHOPEE (Non Star, Star, Star+, Mall)
    var channel = String(data[i][2] || '').toLowerCase();
    if (channel.indexOf('shopee') === -1) continue;

    var status = String(data[i][14] || '').toUpperCase().trim();

    // FILTER Tanggal
    var tgl = new Date(data[i][1]);
    if (start && tgl < start) continue;
    if (end && tgl > end) continue;

    var qty = Number(data[i][9]) || 1;
    var skuVar = String(data[i][7]).trim().toLowerCase();

    var prod = prodMap[skuVar] || null;
    var hargaJual = prod ? prod.jual : (Number(data[i][11]) || 0) / qty;
    var adminPct = prod ? prod.adminPct : 0;
    var hppUnit = prod ? prod.hpp : (Number(data[i][13]) || 0) / qty;

    var jualTotal = hargaJual * qty;
    var adminRp = jualTotal * (adminPct / 100);
    var hppTotal = hppUnit * qty;

    // SHIPPED = dalam perjalanan (belum masuk profit). WAITING tidak dihitung.
    if (status !== 'SHIPPED' && status !== 'COMPLETED' && status !== 'LOST') continue;
    if (status === 'SHIPPED') {
      onTheWayJual += jualTotal;
      onTheWayAdmin += adminRp;
      onTheWayCount++;
      continue;
    }

    // COMPLETED dan LOST = masuk profit
    if (status !== 'COMPLETED' && status !== 'LOST') continue;

    totalJual += jualTotal;
    totalAdmin += adminRp;
    totalHpp += hppTotal;
    orderCount++;
  }

  var penghasilan = totalJual - totalAdmin;
  var profit = penghasilan - totalHpp;

  return {
    totalJual: Math.round(totalJual),
    totalAdmin: Math.round(totalAdmin),
    penghasilan: Math.round(penghasilan),
    totalHpp: Math.round(totalHpp),
    profit: Math.round(profit),
    orders: orderCount,
    onTheWay: {
      totalJual: Math.round(onTheWayJual),
      setelahAdmin: Math.round(onTheWayJual - onTheWayAdmin),
      count: onTheWayCount
    }
  };
}

// === PROFIT TIKTOK ===
function getTiktokProfitData(startDate, endDate) {
  var ss = _getSS();
  var sheetPesanan = ss.getSheetByName('Data_Pesanan');
  var sheetProduk = ss.getSheetByName('Master_Produk');
  if (!sheetPesanan || !sheetProduk) return { penghasilan: 0, totalHpp: 0, profit: 0, orders: 0 };

  // Cache adminTiktok% + hpp dari Master_Produk
  var prodData = sheetProduk.getDataRange().getValues();
  var prodMap = {};
  for (var p = 1; p < prodData.length; p++) {
    var sku = String(prodData[p][3]).trim();
    if (sku) {
      prodMap[sku.toLowerCase()] = {
        adminPct: Number(prodData[p][22]) || 0,  // Kolom W = Admin TikTok%
        hpp: Number(prodData[p][9]) || 0,
        jual: Number(prodData[p][11]) || 0
      };
    }
  }

  var data = sheetPesanan.getDataRange().getValues();
  var start = startDate ? new Date(startDate) : null;
  var end = endDate ? new Date(endDate) : null;
  if (end) { end.setHours(23, 59, 59); }

  var totalJual = 0, totalAdmin = 0, totalHpp = 0, orderCount = 0;
  var onTheWayJual = 0, onTheWayAdmin = 0, onTheWayCount = 0;

  for (var i = 1; i < data.length; i++) {
    var channel = String(data[i][2] || '').toLowerCase();
    if (channel.indexOf('tiktok') === -1) continue;

    var status = String(data[i][14] || '').toUpperCase().trim();
    var tgl = new Date(data[i][1]);
    if (start && tgl < start) continue;
    if (end && tgl > end) continue;

    var qty = Number(data[i][9]) || 1;
    var skuVar = String(data[i][7]).trim().toLowerCase();
    var prod = prodMap[skuVar] || null;
    var hargaJual = prod ? prod.jual : (Number(data[i][11]) || 0) / qty;
    var adminPct = prod ? prod.adminPct : 0;
    var hppUnit = prod ? prod.hpp : (Number(data[i][13]) || 0) / qty;

    var jualTotal = hargaJual * qty;
    var adminRp = jualTotal * (adminPct / 100);
    var hppTotal = hppUnit * qty;

    // Hanya SHIPPED masuk "dalam perjalanan". WAITING tidak dihitung.
    if (status !== 'SHIPPED' && status !== 'COMPLETED' && status !== 'LOST') continue;
    if (status === 'SHIPPED') {
      onTheWayJual += jualTotal;
      onTheWayAdmin += adminRp;
      onTheWayCount++;
      continue;
    }

    if (status !== 'COMPLETED' && status !== 'LOST') continue;

    totalJual += jualTotal;
    totalAdmin += adminRp;
    totalHpp += hppTotal;
    orderCount++;
  }

  var penghasilan = totalJual - totalAdmin;
  var profit = penghasilan - totalHpp;

  return {
    totalJual: Math.round(totalJual),
    totalAdmin: Math.round(totalAdmin),
    penghasilan: Math.round(penghasilan),
    totalHpp: Math.round(totalHpp),
    profit: Math.round(profit),
    orders: orderCount,
    onTheWay: {
      totalJual: Math.round(onTheWayJual),
      setelahAdmin: Math.round(onTheWayJual - onTheWayAdmin),
      count: onTheWayCount
    }
  };
}

// Edit Masal Produk (stok, hpp, jual)
function bulkUpdateProducts(items) {
  if (!items || items.length === 0) return { success: false, message: 'Tidak ada data' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch(e) { return { success: false, message: 'Server sibuk' }; }

  try {
    var ss = _getSS();
    var sheet = ss.getSheetByName('Master_Produk');
    var sHist = ss.getSheetByName('Riwayat_Stok');
    var data = sheet.getDataRange().getValues();
    var user = Session.getActiveUser().getEmail();
    var timestamp = new Date();
    var updated = 0;

    var histRows = [];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var sku = String(item.sku || '').trim();
      if (!sku) continue;

      for (var r = 1; r < data.length; r++) {
        if (String(data[r][3]).trim() !== sku) continue;

        var rowNum = r + 1;
        var oldSisa = Number(data[r][7]) || 0;
        var oldHpp = Number(data[r][9]) || 0;
        var oldJual = Number(data[r][11]) || 0;
        var changed = false;
        var changes = [];

        if (item.stok !== undefined && item.stok !== null && item.stok !== '') {
          var newStok = Number(item.stok);
          if (newStok !== oldSisa) {
            data[r][7] = newStok;
            changes.push('Stok: ' + oldSisa + ' -> ' + newStok);
            changed = true;
          }
        }

        if (item.hpp !== undefined && item.hpp !== null && item.hpp !== '') {
          var newHpp = Number(item.hpp);
          if (newHpp !== oldHpp) {
            data[r][9] = newHpp;
            changes.push('HPP: ' + oldHpp + ' -> ' + newHpp);
            changed = true;
          }
        }

        if (item.jual !== undefined && item.jual !== null && item.jual !== '') {
          var newJual = Number(item.jual);
          if (newJual !== oldJual) {
            data[r][11] = newJual;
            changes.push('Jual: ' + oldJual + ' -> ' + newJual);
            changed = true;
          }
        }

        if (changed) {
          var curStok = Number(data[r][7]) || 0;
          var curHpp = Number(data[r][9]) || 0;
          // BATCH: H, J, K, L sekaligus (kolom 8-12 = 5 kolom)
          sheet.getRange(rowNum, 8, 1, 5).setValues([[curStok, data[r][8], curHpp, Math.round(curStok * curHpp), data[r][11]]]);

          histRows.push(['EDIT-' + Math.floor(Date.now()/1000) + '-' + updated, timestamp, 'EDIT', 'Bulk Edit', sku, data[r][1], data[r][4], 0, curHpp, 0, curStok, changes.join(', '), user, timestamp]);
          updated++;
        }
        break;
      }
    }

    // BATCH APPEND riwayat
    if (histRows.length > 0 && sHist) {
      sHist.getRange(sHist.getLastRow() + 1, 1, histRows.length, 14).setValues(histRows);
    }
    SpreadsheetApp.flush();
    return { success: true, message: 'Berhasil update ' + updated + ' produk' };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================================
// BACKUP & RESET SYSTEM
// ==========================================================

// Daftar sheet yang bisa di-backup/reset
function getBackupSheetList() {
  var sheets = ['Data_Pesanan','Master_Produk','Riwayat_Stok','Data_Riwayat','Data_Rekap','Data_Produksi','Data_Proses_Jahit','Master_Bahan','Master_Resep'];
  var ss = _getSS();
  var result = [];
  for (var i = 0; i < sheets.length; i++) {
    var s = ss.getSheetByName(sheets[i]);
    if (s) {
      var rows = Math.max(0, s.getLastRow() - 1);
      result.push({ name: sheets[i], rows: rows });
    }
  }
  return result;
}

// Ambil data sheet untuk download (return array of {name, headers, data})
function getBackupData(sheetNames) {
  var ss = _getSS();
  var result = [];
  for (var i = 0; i < sheetNames.length; i++) {
    var s = ss.getSheetByName(sheetNames[i]);
    if (!s) continue;
    var all = s.getDataRange().getDisplayValues();
    if (all.length === 0) continue;
    result.push({ name: sheetNames[i], headers: all[0], data: all.slice(1) });
  }
  return result;
}

// Reset sheet (hapus semua data kecuali header)
function resetSheetData(sheetNames, confirmCode) {
  if (confirmCode !== 'RESET-CONFIRM') return { success: false, message: 'Kode konfirmasi salah' };
  var ss = _getSS();
  var sLog = ss.getSheetByName('Data_Riwayat');
  var user = Session.getActiveUser().getEmail();
  var timestamp = new Date();
  var count = 0;

  for (var i = 0; i < sheetNames.length; i++) {
    var s = ss.getSheetByName(sheetNames[i]);
    if (!s) continue;
    var lastRow = s.getLastRow();
    if (lastRow <= 1) continue;
    s.deleteRows(2, lastRow - 1);
    count++;
  }

  if (sLog) {
    sLog.appendRow([timestamp, 'RESET', 'RESET DATA', '-', sheetNames.join(', '), count, user, 'Reset ' + count + ' sheet via Backup Manager']);
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'Berhasil reset ' + count + ' sheet' };
}

// Backup ke Google Drive (buat spreadsheet baru)
function backupToDrive(sheetNames) {
  var ss = _getSS();
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd_HH-mm');
  var backupName = 'Backup_' + ss.getName() + '_' + timestamp;

  // Buat spreadsheet baru
  var newSS = SpreadsheetApp.create(backupName);

  // Hapus sheet default
  var defaultSheet = newSS.getSheets()[0];

  for (var i = 0; i < sheetNames.length; i++) {
    var src = ss.getSheetByName(sheetNames[i]);
    if (!src || src.getLastRow() === 0) continue;
    var copied = src.copyTo(newSS);
    copied.setName(sheetNames[i]);
  }

  // Hapus sheet default jika ada sheet lain
  if (newSS.getSheets().length > 1) {
    newSS.deleteSheet(defaultSheet);
  }

  // Log
  var sLog = ss.getSheetByName('Data_Riwayat');
  if (sLog) {
    sLog.appendRow([new Date(), 'BACKUP', 'BACKUP TO DRIVE', '-', sheetNames.join(', '), sheetNames.length, Session.getActiveUser().getEmail(), 'File: ' + backupName]);
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'Backup tersimpan di Google Drive', fileName: backupName, url: newSS.getUrl() };
}

// ===== AUTO BACKUP =====

// Setup trigger auto backup
function setupAutoBackup(interval) {
  // Hapus trigger lama
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runAutoBackup') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  var props = PropertiesService.getScriptProperties();

  if (!interval || interval === 'off') {
    props.setProperty('autoBackupInterval', 'off');
    props.deleteProperty('autoBackupMonths');
    return { success: true, message: 'Auto backup dinonaktifkan' };
  }

  props.setProperty('autoBackupInterval', interval);
  props.deleteProperty('autoBackupMonths');

  // Set waktu mulai agar countdown bisa jalan
  props.setProperty('lastAutoBackup', new Date().toISOString());

  if (interval === '1menit') {
    ScriptApp.newTrigger('runAutoBackup').timeBased().everyMinutes(1).create();
  } else if (interval === '1hari') {
    ScriptApp.newTrigger('runAutoBackup').timeBased().everyDays(1).atHour(2).create();
  } else if (interval === '1bulan') {
    ScriptApp.newTrigger('runAutoBackup').timeBased().onMonthDay(1).atHour(2).create();
  } else if (interval === '3bulan') {
    ScriptApp.newTrigger('runAutoBackup').timeBased().onMonthDay(1).atHour(2).create();
    props.setProperty('autoBackupMonths', '3');
  } else if (interval === '6bulan') {
    ScriptApp.newTrigger('runAutoBackup').timeBased().onMonthDay(1).atHour(2).create();
    props.setProperty('autoBackupMonths', '6');
  }

  // Log aktivasi
  try {
    var sLog = _getSS().getSheetByName('Data_Riwayat');
    if (sLog) sLog.appendRow([new Date(), 'AUTO-BACKUP', 'SETUP', '-', interval, '-', Session.getActiveUser().getEmail(), 'Auto backup diaktifkan: ' + interval]);
  } catch(e) {}

  return { success: true, message: 'Auto backup diaktifkan: ' + interval };
}

// Fungsi yang dipanggil trigger
function runAutoBackup() {
  var props = PropertiesService.getScriptProperties();
  var interval = props.getProperty('autoBackupInterval') || '';
  if (!interval || interval === 'off') return;

  var months = parseInt(props.getProperty('autoBackupMonths')) || 0;

  // Cek interval bulanan (3/6 bulan)
  if (months > 0) {
    var lastBackup = props.getProperty('lastAutoBackup');
    if (lastBackup) {
      var last = new Date(lastBackup);
      var now = new Date();
      var diffMonths = (now.getFullYear() - last.getFullYear()) * 12 + (now.getMonth() - last.getMonth());
      if (diffMonths < months) return;
    }
  }

  try {
    var sheets = ['Data_Pesanan', 'Master_Produk', 'Riwayat_Stok', 'Data_Riwayat', 'Master_Bahan'];
    var result = backupToDrive(sheets);
    props.setProperty('lastAutoBackup', new Date().toISOString());

    // Log sukses
    var sLog = _getSS().getSheetByName('Data_Riwayat');
    if (sLog) {
      sLog.appendRow([new Date(), 'AUTO-BACKUP', 'AUTO BACKUP DONE', '-', sheets.join(', '), sheets.length, 'Trigger', 'Otomatis (' + interval + '): ' + (result.fileName || 'OK')]);
    }
  } catch(e) {
    // Log error
    try {
      var sLog2 = _getSS().getSheetByName('Data_Riwayat');
      if (sLog2) sLog2.appendRow([new Date(), 'AUTO-BACKUP', 'AUTO BACKUP FAIL', '-', '-', 0, 'Trigger', 'Error: ' + e.toString()]);
    } catch(e2) {}
  }
}

// Cek status auto backup
function getAutoBackupStatus() {
  var props = PropertiesService.getScriptProperties();
  var interval = props.getProperty('autoBackupInterval') || 'off';
  var lastBackup = props.getProperty('lastAutoBackup') || null;

  // Hitung next backup
  var nextBackup = null;
  if (lastBackup && interval !== 'off') {
    var last = new Date(lastBackup);
    var intervals = {'1menit': 60000, '1hari': 86400000, '1bulan': 2592000000, '3bulan': 7776000000, '6bulan': 15552000000};
    var ms = intervals[interval] || 0;
    if (ms > 0) nextBackup = new Date(last.getTime() + ms).toISOString();
  }

  // Ambil log backup terakhir (5 terbaru) - cek kolom 1 dan 2
  var logs = [];
  try {
    var ss = _getSS();
    var sLog = ss.getSheetByName('Data_Riwayat');
    if (sLog && sLog.getLastRow() > 1) {
      var data = sLog.getDataRange().getDisplayValues();
      for (var i = data.length - 1; i >= 1; i--) {
        var col1 = String(data[i][1] || '').toUpperCase();
        var col2 = String(data[i][2] || '').toUpperCase();
        if (col1.includes('BACKUP') || col2.includes('BACKUP')) {
          logs.push({ date: String(data[i][0] || ''), type: col2 || col1, detail: String(data[i][7] || '') });
          if (logs.length >= 5) break;
        }
      }
    }
  } catch(e) {}

  return { interval: interval, lastBackup: lastBackup, nextBackup: nextBackup, logs: logs };
}

// ==========================================================
// LACAK RESI (via PHP Proxy - support SPX, J&T, JNE, dll)
// ==========================================================

function trackResi(awb) {
  if (!awb || awb === '-') return { success: false, message: 'Nomor resi kosong.' };
  var resi = String(awb).trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (resi.length < 8) return { success: false, message: 'Nomor resi terlalu pendek.' };

  // Ambil base URL dari Config, fallback ke default
  var baseUrl = '';
  try { baseUrl = getAppConfig().tracking_api_url || ''; } catch(e) {}
  if (!baseUrl) baseUrl = 'https://dataeasy.web.id/api/tracking.php';

  try {
    var url = baseUrl + '?awb=' + resi;
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    var body = JSON.parse(res.getContentText());
    return body;
  } catch(e) {
    return { success: false, message: 'Gagal menghubungi server tracking: ' + e.toString() };
  }
}

// ==========================================================
// FITUR HUTANG PIUTANG + MASTER KONTAK
// ==========================================================

// --- MASTER KONTAK (SUPPLIER & RESELLER) ---

function _ensureKontakSheet() {
  var ss = _getSS();
  var sh = ss.getSheetByName('Master_Kontak');
  if (!sh) {
    sh = ss.insertSheet('Master_Kontak');
    sh.appendRow(['Nama','Tipe','No HP','Alamat','Catatan','Created At']);
    sh.getRange(1,1,1,6).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getKontakList(tipe) {
  var sh = _ensureKontakSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2,1,last-1,6).getDisplayValues();
  if (tipe) {
    var t = String(tipe).toUpperCase();
    data = data.filter(function(r){ return String(r[1]).toUpperCase() === t; });
  }
  return data;
}

function addKontak(form) {
  var sh = _ensureKontakSheet();
  var nama = String(form.nama || '').trim();
  if (!nama) return { success: false, message: 'Nama wajib diisi.' };
  var tipe = String(form.tipe || 'SUPPLIER').toUpperCase();
  if (tipe !== 'SUPPLIER' && tipe !== 'RESELLER') tipe = 'SUPPLIER';

  // Cek duplikat nama+tipe
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nama.toLowerCase() && String(data[i][1]).toUpperCase() === tipe) {
      return { success: false, message: 'Kontak "' + nama + '" (' + tipe + ') sudah ada.' };
    }
  }

  sh.appendRow([nama, tipe, form.phone || '', form.alamat || '', form.catatan || '', new Date()]);
  return { success: true, message: 'Kontak berhasil ditambahkan.' };
}

function deleteKontak(nama, tipe) {
  var sh = _ensureKontakSheet();
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim().toLowerCase() === String(nama).trim().toLowerCase() &&
        String(data[i][1]).toUpperCase() === String(tipe).toUpperCase()) {
      sh.deleteRow(i + 1);
      return { success: true, message: 'Kontak dihapus.' };
    }
  }
  return { success: false, message: 'Kontak tidak ditemukan.' };
}

// --- DATA HUTANG PIUTANG ---

function _ensureHPSheet() {
  var ss = _getSS();
  var sh = ss.getSheetByName('Data_HutangPiutang');
  if (!sh) {
    sh = ss.insertSheet('Data_HutangPiutang');
    sh.appendRow(['ID','Tipe','Tanggal','Jatuh Tempo','Nama Pihak','SKU Barang','Nama Barang','Qty','Harga Satuan','Total','Status','Tanggal Lunas','Catatan','User Input','Created At','Sudah Dibayar']);
    sh.getRange(1,1,1,16).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getHutangPiutangList() {
  var sh = _ensureHPSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var cols = Math.max(sh.getLastColumn(), 16);
  var data = sh.getRange(2, 1, last - 1, cols).getDisplayValues();
  return data.reverse();
}

function addHutangPiutang(form) {
  var sh = _ensureHPSheet();
  var tipe = String(form.tipe || '').toUpperCase();
  if (tipe !== 'HUTANG' && tipe !== 'PIUTANG') return { success: false, message: 'Tipe tidak valid.' };

  var nama = String(form.namaPihak || '').trim();
  if (!nama) return { success: false, message: 'Nama pihak wajib diisi.' };

  var sku = String(form.sku || '').trim();
  var namaBarang = String(form.namaBarang || '').trim();
  if (!namaBarang) return { success: false, message: 'Nama barang wajib diisi.' };

  var qty = Number(form.qty) || 0;
  var harga = Number(form.harga) || 0;
  if (qty <= 0 || harga <= 0) return { success: false, message: 'Qty dan Harga harus lebih dari 0.' };

  var total = qty * harga;
  var tgl = form.tanggal || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  var jatuhTempo = form.jatuhTempo || '';
  var catatan = form.catatan || '';
  var user = Session.getActiveUser().getEmail() || 'Admin';

  var prefix = tipe === 'HUTANG' ? 'HT' : 'PT';
  var id = prefix + '-' + Math.floor(Date.now() / 1000);

  sh.appendRow([id, tipe, tgl, jatuhTempo, nama, sku, namaBarang, qty, harga, total, 'BELUM_LUNAS', '', catatan, user, new Date(), 0]);

  // Log activity
  var sLog = _getSS().getSheetByName('Data_Riwayat');
  if (sLog) {
    sLog.appendRow([new Date(), id, 'NEW ' + tipe, sku, namaBarang, total, user, tipe + ' ke ' + nama]);
  }

  return { success: true, message: tipe + ' berhasil dicatat.' };
}

function updateHPStatus(id, newStatus) {
  var sh = _ensureHPSheet();
  var data = sh.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      var status = String(newStatus).toUpperCase();
      sh.getRange(i + 1, 11).setValue(status); // Kolom K: Status

      if (status === 'LUNAS') {
        sh.getRange(i + 1, 12).setValue(Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')); // Kolom L: Tgl Lunas
      } else {
        sh.getRange(i + 1, 12).setValue('');
      }

      // Log
      var sLog = _getSS().getSheetByName('Data_Riwayat');
      if (sLog) {
        sLog.appendRow([new Date(), id, 'UPDATE ' + data[i][1], data[i][5], data[i][6], data[i][9], Session.getActiveUser().getEmail() || 'Admin', 'Status → ' + status]);
      }

      return { success: true, message: 'Status berhasil diubah ke ' + status + '.' };
    }
  }
  return { success: false, message: 'ID tidak ditemukan.' };
}

function bayarHutangPiutang(id, jumlahBayar) {
  var sh = _ensureHPSheet();
  var data = sh.getDataRange().getValues();
  var amount = Number(jumlahBayar) || 0;
  if (amount <= 0) return { success: false, message: 'Jumlah bayar harus lebih dari 0.' };

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      var totalTagihan = Number(data[i][9]) || 0;
      var sudahBayar = Number(data[i][15]) || 0; // Kolom P (index 15)
      var bayarBaru = sudahBayar + amount;
      var sisa = totalTagihan - bayarBaru;

      // Update kolom Sudah Dibayar (P = kolom 16)
      sh.getRange(i + 1, 16).setValue(bayarBaru);

      // Jika sudah lunas (sisa <= 0)
      if (sisa <= 0) {
        sh.getRange(i + 1, 11).setValue('LUNAS'); // Kolom K: Status
        sh.getRange(i + 1, 12).setValue(Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')); // Kolom L: Tgl Lunas
      }

      // Log
      var sLog = _getSS().getSheetByName('Data_Riwayat');
      var user = Session.getActiveUser().getEmail() || 'Admin';
      if (sLog) {
        var logMsg = sisa <= 0
          ? 'LUNAS (Bayar ' + amount + ', Total dibayar ' + bayarBaru + ')'
          : 'Bayar sebagian ' + amount + ' (Sisa ' + sisa + ')';
        sLog.appendRow([new Date(), id, 'BAYAR ' + data[i][1], data[i][5], data[i][6], amount, user, logMsg]);
      }

      var msg = sisa <= 0
        ? 'Pembayaran LUNAS! Total dibayar: Rp ' + bayarBaru.toLocaleString('id-ID')
        : 'Pembayaran Rp ' + amount.toLocaleString('id-ID') + ' tercatat. Sisa: Rp ' + sisa.toLocaleString('id-ID');

      return { success: true, message: msg };
    }
  }
  return { success: false, message: 'ID tidak ditemukan.' };
}

function deleteHutangPiutang(id) {
  var sh = _ensureHPSheet();
  var data = sh.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      // Log sebelum hapus
      var sLog = _getSS().getSheetByName('Data_Riwayat');
      if (sLog) {
        sLog.appendRow([new Date(), id, 'DELETE ' + data[i][1], data[i][5], data[i][6], data[i][9], Session.getActiveUser().getEmail() || 'Admin', 'Hapus record']);
      }
      sh.deleteRow(i + 1);
      return { success: true, message: 'Data berhasil dihapus.' };
    }
  }
  return { success: false, message: 'ID tidak ditemukan.' };
}

function getHPSummary() {
  var sh = _ensureHPSheet();
  var last = sh.getLastRow();
  if (last < 2) return { totalHutang: 0, totalPiutang: 0, selisih: 0, jatuhTempo: 0 };

  var data = sh.getRange(2,1,last-1,15).getValues();
  var totalHutang = 0, totalPiutang = 0, jatuhTempoCount = 0;
  var today = new Date();
  today.setHours(0,0,0,0);

  var nextWeek = new Date(today);
  nextWeek.setDate(nextWeek.getDate() + 7);

  for (var i = 0; i < data.length; i++) {
    var tipe = String(data[i][1]).toUpperCase();
    var total = Number(data[i][9]) || 0;
    var status = String(data[i][10]).toUpperCase();

    if (status === 'BELUM_LUNAS') {
      if (tipe === 'HUTANG') totalHutang += total;
      if (tipe === 'PIUTANG') totalPiutang += total;

      // Cek jatuh tempo dalam 7 hari ke depan atau sudah lewat
      var jt = data[i][3];
      if (jt) {
        var jtDate = new Date(jt);
        if (!isNaN(jtDate.getTime()) && jtDate <= nextWeek) {
          jatuhTempoCount++;
        }
      }
    }
  }

  return {
    totalHutang: totalHutang,
    totalPiutang: totalPiutang,
    selisih: totalPiutang - totalHutang,
    jatuhTempo: jatuhTempoCount
  };
}

// ==========================================================
// FITUR CEK ONGKIR (Scrape cek-ongkir.com)
// ==========================================================

function searchOngkirCity(term) {
  if (!term || term.length < 2) return [];
  try {
    var url = 'https://dataeasy.web.id/api/ongkir.php?action=search&term=' + encodeURIComponent(term);
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    var data = JSON.parse(res.getContentText());
    return (data || []).slice(0, 12);
  } catch(e) {
    return [];
  }
}

function cekOngkir(originId, destId, weightKg, couriers) {
  if (!originId || !destId) return { success: false, message: 'Pilih kota asal dan tujuan.' };
  var w = Number(weightKg) || 1;
  var c = couriers || 'jne:jnt:sicepat';

  try {
    var res = UrlFetchApp.fetch('https://dataeasy.web.id/api/ongkir.php', {
      method: 'post',
      payload: {
        action: 'cost',
        origin: String(originId),
        destination: String(destId),
        weight: String(w),
        courier: c
      },
      muteHttpExceptions: true,
      followRedirects: true
    });

    var body = JSON.parse(res.getContentText());
    return body;
  } catch(e) {
    return { success: false, message: 'Gagal cek ongkir: ' + e.toString() };
  }
}

// ==========================================================
// AUTO SYNC TRACKING SHOPEE
// ==========================================================

/**
 * Sinkronisasi status tracking pesanan Shopee yang masih SHIPPED.
 * Dipanggil oleh trigger (6 jam) atau manual dari frontend.
 * 
 * Alur:
 * - delivered → COMPLETED (masuk profit)
 * - lost → LOST (masuk profit, diasuransikan)
 * - returned → PENDING_RETURN (menunggu konfirmasi manual)
 * - lainnya → tetap SHIPPED
 */
function syncShopeeTracking(manualTrigger) {
  // Cek apakah penyewa ini mengaktifkan sync (SKIP jika manual trigger dari tombol)
  var syncCfg = _getSyncConfig();
  if (!manualTrigger && !syncCfg.isActive) return { success: true, message: 'Sync tidak aktif untuk penyewa ini.', results: {} };

  // Simpan timestamp sync ke sheet Config penyewa
  _setSyncConfig('sync_last', new Date().toISOString());

  var ss = _getSS();
  var sheetOrder = ss.getSheetByName('Data_Pesanan');
  var sLog = ss.getSheetByName('Data_Riwayat');
  if (!sheetOrder) return { success: false, message: 'Sheet Data_Pesanan tidak ditemukan.' };

  var data = sheetOrder.getDataRange().getValues();
  var timestamp = new Date();
  var results = { completed: 0, lost: 0, pendingReturn: 0, skipped: 0, errors: 0, total: 0, shipped: 0 };

  // Kumpulkan baris yang perlu dicek
  var toCheck = [];
  var debugInfo = { totalRows: data.length - 1, shopeeCount: 0, shippedCount: 0, withResi: 0, statusFound: {} };

  // === PRE-BUILD: Map orderId → items[] untuk stock lookup O(1) ===
  var orderItemsMap = {};

  for (var i = 1; i < data.length; i++) {
    var channel = String(data[i][2] || '').toLowerCase();
    var status = String(data[i][14] || '').toUpperCase().trim();
    var resi = String(data[i][4] || '').trim();
    var orderId = String(data[i][0] || '').trim();

    // Debug: catat semua status yang ditemukan
    if (!debugInfo.statusFound[status || '(kosong)']) debugInfo.statusFound[status || '(kosong)'] = 0;
    debugInfo.statusFound[status || '(kosong)']++;

    var isShopee = channel.indexOf('shopee') !== -1;
    var isTiktok = channel.indexOf('tiktok') !== -1;
    var isTracked = isShopee || isTiktok; // Shopee + TikTok di-track
    if (isTracked) debugInfo.shopeeCount++;
    if (isTracked && (status === 'SHIPPED' || status === 'WAITING')) debugInfo.shippedCount++;
    if (isTracked && (status === 'SHIPPED' || status === 'WAITING') && resi && resi !== '-') debugInfo.withResi++;

    // Cek pesanan Shopee/TikTok yang SHIPPED atau WAITING (punya resi)
    if (isTracked && (status === 'SHIPPED' || status === 'WAITING') && resi && resi !== '-') {
      toCheck.push({ rowIndex: i + 1, resi: resi, orderId: orderId, currentStatus: status });
    }

    // Pre-build order items map untuk stock cut (menghindari O(n) loop per order)
    if (orderId) {
      if (!orderItemsMap[orderId]) orderItemsMap[orderId] = [];
      var skuVar = String(data[i][7] || data[i][5] || '').trim();
      var itemNama = String(data[i][6] || '').trim();
      var itemVariasi = String(data[i][8] || '').trim();
      // Push selama ada nama produk ATAU sku (jangan skip item tanpa SKU)
      if (skuVar || itemNama) {
        orderItemsMap[orderId].push({
          rowIndex: i + 1,
          sku: skuVar,
          skuCandidates: [skuVar],
          qty: parseFloat(data[i][9]) || 0,
          nama: itemNama,
          variasi: itemVariasi
        });
      }
    }
  }

  results.total = toCheck.length;
  results.debug = debugInfo;

  if (toCheck.length === 0) {
    var debugMsg = 'Tidak ada pesanan untuk dicek. Detail: ' + 
      debugInfo.totalRows + ' total baris, ' + 
      debugInfo.shopeeCount + ' channel Shopee, ' + 
      debugInfo.shippedCount + ' status SHIPPED. ' +
      'Status ditemukan: ' + JSON.stringify(debugInfo.statusFound);
    return { success: true, message: debugMsg, results: results };
  }

  // === AUTO MULTI-BATCH: Tarik semua, proses per 100 ===
  var BATCH_SIZE = 100;
  var totalBatches = Math.ceil(toCheck.length / BATCH_SIZE);

  var baseUrl = '';
  try { baseUrl = getAppConfig().tracking_api_url || ''; } catch(e) {}
  if (!baseUrl) baseUrl = 'https://dataeasy.web.id/api/tracking.php';
  var batchUrl = baseUrl.replace(/tracking\.php$/, 'tracking-batch.php');

  var sheetMaster = ss.getSheetByName('Master_Produk');
  var sheetStok = ss.getSheetByName('Riwayat_Stok');
  var statusUpdates = [];
  var logRows = [];

  // === LOOP SEMUA BATCH ===
  for (var batchNum = 0; batchNum < totalBatches; batchNum++) {
    var batchStart = batchNum * BATCH_SIZE;
    var batch = toCheck.slice(batchStart, batchStart + BATCH_SIZE);

    // Kumpulkan AWB untuk batch ini
    var awbList = [];
    for (var b = 0; b < batch.length; b++) {
      awbList.push(String(batch[b].resi).trim().toUpperCase().replace(/[^A-Z0-9]/g, ''));
    }

    // Kirim ke PHP batch endpoint
    var allTrackResults = [];
    try {
      var response = UrlFetchApp.fetch(batchUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ awbs: awbList, secret: getExtensionSecret_() }),
        muteHttpExceptions: true,
        followRedirects: true
      });
      var respBody = JSON.parse(response.getContentText());
      if (respBody && respBody.success && respBody.results) {
        allTrackResults = respBody.results;
      } else {
        allTrackResults = _fallbackFetchAll(awbList, baseUrl);
      }
    } catch(e) {
      allTrackResults = _fallbackFetchAll(awbList, baseUrl);
    }

    // Proses hasil batch ini
    for (var j = 0; j < batch.length; j++) {
      var item = batch[j];
      var trackResult = allTrackResults[j] || null;

      try {
        if (!trackResult || !trackResult.success) { results.skipped++; continue; }

        var lastDesc = String(trackResult.status || '').toLowerCase();
        if (trackResult.history && trackResult.history.length > 0) {
          var firstHistory = String(trackResult.history[0].desc || '').toLowerCase();
          if (firstHistory) lastDesc = lastDesc + ' ' + firstHistory;
        }

        var newStatus = '';
        var logType = '';
        var needStockCut = false;

        var isDeliveredByCourier = lastDesc.indexOf('delivered by courier') !== -1 || lastDesc.indexOf('being delivered') !== -1 || lastDesc.indexOf('out for delivery') !== -1 || lastDesc.indexOf('sedang diantar') !== -1;
        if (!isDeliveredByCourier && (lastDesc.indexOf('delivered') !== -1 || lastDesc.indexOf('terkirim') !== -1 || lastDesc.indexOf('diterima') !== -1 || lastDesc.indexOf('sampai tujuan') !== -1 || lastDesc.indexOf('selesai') !== -1 || lastDesc.indexOf('sukses kirim') !== -1)) {
          newStatus = 'COMPLETED'; logType = 'AUTO COMPLETED'; results.completed++;
          if (item.currentStatus === 'WAITING') needStockCut = true;
        } else if (lastDesc.indexOf('lost') !== -1 || lastDesc.indexOf('hilang') !== -1) {
          newStatus = 'LOST'; logType = 'AUTO LOST'; results.lost++;
          if (item.currentStatus === 'WAITING') needStockCut = true;
        } else if (lastDesc.indexOf('returned') !== -1 || lastDesc.indexOf('dikembalikan') !== -1 || lastDesc.indexOf('retur') !== -1 || lastDesc.indexOf('undelivered') !== -1 || lastDesc.indexOf('failed delivery') !== -1) {
          newStatus = 'PENDING_RETURN'; logType = 'AUTO PENDING_RETURN'; results.pendingReturn++;
          if (item.currentStatus === 'WAITING') needStockCut = true;
        } else {
          var transitKeywords = ['pengiriman', 'transit', 'pickup', 'picked up', 'dikirim', 'on delivery', 'sorting', 'hub', 'warehouse', 'courier', 'delivering', 'received by sorting', 'manifested', 'departed', 'arrived'];
          var isTransit = item.currentStatus === 'WAITING' && trackResult.history && trackResult.history.length >= 1 && transitKeywords.some(function(k) { return lastDesc.indexOf(k) !== -1; });
          if (isTransit) {
            newStatus = 'SHIPPED'; logType = 'AUTO SHIPPED'; results.shipped++;
            needStockCut = true;
          } else { results.skipped++; continue; }
        }

        if (needStockCut && sheetMaster && sheetStok) {
          var orderItems = orderItemsMap[String(item.orderId).trim()] || [];
          if (orderItems.length > 0) {
            processStockOut(sheetStok, sheetMaster, orderItems, timestamp, item.resi, 'System (Auto Sync)');
          }
        }

        statusUpdates.push({ rowIndex: item.rowIndex, newStatus: newStatus, setTimestamp: (newStatus === 'SHIPPED' || newStatus === 'COMPLETED') });
        logRows.push([timestamp, item.orderId, logType, item.resi, '', 0, 'System (Auto Sync)', 'Tracking: ' + lastDesc]);

      } catch(e) { results.errors++; }
    }
  } // end batch loop

  // === BATCH WRITE KE SHEET (sekali untuk semua batch) ===
  for (var u = 0; u < statusUpdates.length; u++) {
    var upd = statusUpdates[u];
    sheetOrder.getRange(upd.rowIndex, 15).setValue(upd.newStatus);
    if (upd.setTimestamp) {
      sheetOrder.getRange(upd.rowIndex, 16).setValue(timestamp);
      sheetOrder.getRange(upd.rowIndex, 17).setValue(timestamp);
    }
  }

  if (sLog && logRows.length > 0) {
    var logLastRow = sLog.getLastRow();
    sLog.getRange(logLastRow + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
  }

  SpreadsheetApp.flush();

  var elapsed = Math.round((new Date() - timestamp) / 1000);
  var msg = 'Sync selesai (' + toCheck.length + ' resi, ' + totalBatches + ' batch, ' + elapsed + ' detik). ';
  if (results.shipped > 0) msg += results.shipped + ' DIKIRIM, ';
  if (results.completed > 0) msg += results.completed + ' COMPLETED, ';
  if (results.lost > 0) msg += results.lost + ' LOST, ';
  if (results.pendingReturn > 0) msg += results.pendingReturn + ' menunggu konfirmasi retur, ';
  if (results.skipped > 0) msg += results.skipped + ' belum ada update, ';
  if (results.errors > 0) msg += results.errors + ' error.';

  return { success: true, message: msg.replace(/, $/, '.'), results: results };
}

/**
 * Fallback: Jika batch endpoint PHP gagal, gunakan fetchAll() langsung ke tracking.php
 * Ini memastikan sync tetap jalan meskipun tracking-batch.php belum di-deploy
 */
function _fallbackFetchAll(awbList, baseUrl) {
  var CHUNK_SIZE = 20;
  var allResults = [];

  for (var c = 0; c < awbList.length; c += CHUNK_SIZE) {
    var chunk = awbList.slice(c, c + CHUNK_SIZE);
    var requests = [];
    for (var r = 0; r < chunk.length; r++) {
      requests.push({
        url: baseUrl + '?awb=' + chunk[r],
        muteHttpExceptions: true,
        followRedirects: true
      });
    }

    var responses = UrlFetchApp.fetchAll(requests);
    for (var p = 0; p < responses.length; p++) {
      try {
        allResults.push(JSON.parse(responses[p].getContentText()));
      } catch(e) {
        allResults.push({ success: false, message: 'Parse error' });
      }
    }

    if (c + CHUNK_SIZE < awbList.length) {
      Utilities.sleep(1000);
    }
  }

  return allResults;
}

/**
 * Konfirmasi retur dari PENDING_RETURN.
 * action: 'CONFIRM' → RETURNED + stok kembali
 * action: 'REJECT' → kembali ke SHIPPED
 */
function confirmPendingReturn(orderId, action) {
  var ss = _getSS();
  var sheetOrder = ss.getSheetByName('Data_Pesanan');
  var sheetMaster = ss.getSheetByName('Master_Produk');
  var sheetStok = ss.getSheetByName('Riwayat_Stok');
  var sLog = ss.getSheetByName('Data_Riwayat');
  var timestamp = new Date();
  var user = Session.getActiveUser().getEmail() || 'Admin';

  if (!sheetOrder) return { success: false, message: 'Sheet tidak ditemukan.' };

  var data = sheetOrder.getDataRange().getValues();
  var rowsToProcess = [];

  // Cari semua baris dengan orderId ini
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(orderId).trim()) {
      var status = String(data[i][14] || '').toUpperCase().trim();
      if (status !== 'PENDING_RETURN') continue;

      rowsToProcess.push({
        rowIndex: i + 1,
        sku: String(data[i][7] || data[i][5] || '').trim(),
        qty: parseFloat(data[i][9]) || 0,
        nama: data[i][6],
        variasi: data[i][8],
        resi: data[i][4]
      });
    }
  }

  if (rowsToProcess.length === 0) return { success: false, message: 'Pesanan PENDING_RETURN tidak ditemukan.' };

  var act = String(action).toUpperCase();

  if (act === 'CONFIRM') {
    // Kembalikan stok + set RETURNED
    if (sheetMaster && sheetStok) {
      processReturRestock(sheetStok, sheetMaster, rowsToProcess, timestamp, rowsToProcess[0].resi, user);
    }

    rowsToProcess.forEach(function(item) {
      sheetOrder.getRange(item.rowIndex, 15).setValue('RETURNED');
      sheetOrder.getRange(item.rowIndex, 18).setValue(timestamp);
      sheetOrder.getRange(item.rowIndex, 21).setValue('Kondisi: GOOD (Auto Tracking Return)');
    });

    if (sLog) sLog.appendRow([timestamp, orderId, 'CONFIRM RETURN', rowsToProcess[0].sku, rowsToProcess[0].nama, rowsToProcess[0].qty, user, 'Retur dikonfirmasi, stok dikembalikan']);

    SpreadsheetApp.flush();
    return { success: true, message: 'Retur dikonfirmasi. Stok dikembalikan.' };

  } else if (act === 'REJECT') {
    // Kembalikan ke SHIPPED
    rowsToProcess.forEach(function(item) {
      sheetOrder.getRange(item.rowIndex, 15).setValue('SHIPPED');
    });

    if (sLog) sLog.appendRow([timestamp, orderId, 'REJECT RETURN', rowsToProcess[0].sku, rowsToProcess[0].nama, 0, user, 'Pending return ditolak, kembali SHIPPED']);

    SpreadsheetApp.flush();
    return { success: true, message: 'Ditolak. Status kembali ke SHIPPED.' };
  }

  return { success: false, message: 'Action tidak valid.' };
}

/**
 * Ambil ringkasan status tracking untuk notifikasi
 */
function getTrackingSummary() {
  var sh = _getSS().getSheetByName('Data_Pesanan');
  if (!sh) return { lost: 0, pendingReturn: 0, shipped: 0 };

  var data = sh.getDataRange().getValues();
  var lost = 0, pendingReturn = 0, shipped = 0;

  for (var i = 1; i < data.length; i++) {
    var channel = String(data[i][2] || '').toLowerCase();
    if (channel.indexOf('shopee') === -1 && channel.indexOf('tiktok') === -1) continue;

    var status = String(data[i][14] || '').toUpperCase().trim();
    if (status === 'LOST') lost++;
    else if (status === 'PENDING_RETURN') pendingReturn++;
    else if (status === 'SHIPPED') shipped++;
  }

  return { lost: lost, pendingReturn: pendingReturn, shipped: shipped };
}

// --- SYNC CONFIG: Simpan per spreadsheet (per penyewa) di sheet Config kolom I/J ---

function _getSyncConfig() {
  var ss = _getSS();
  var sheet = ss.getSheetByName('Config');
  var result = { lastSync: '', isActive: false };
  if (!sheet) return result;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  var data = sheet.getRange(2, 9, lastRow - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    if (key === 'sync_last') result.lastSync = String(data[i][1] || '');
    if (key === 'sync_active') result.isActive = String(data[i][1]).toLowerCase() === 'true';
  }
  return result;
}

function _setSyncConfig(key, value) {
  var ss = _getSS();
  var sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange('I1').setValue('Key');
    sheet.getRange('J1').setValue('Value');
  }

  var lastRow = sheet.getLastRow();
  var found = false;
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) {
        sheet.getRange(i + 2, 10).setValue(value);
        found = true;
        break;
      }
    }
  }
  if (!found) {
    sheet.getRange(sheet.getLastRow() + 1, 9, 1, 2).setValues([[key, value]]);
  }
}

function getSyncTimestamp() {
  var cfg = _getSyncConfig();
  return { lastSync: cfg.lastSync, intervalHours: 6, isActive: cfg.isActive };
}

function toggleSyncTrackingTrigger(enable) {
  // Trigger tetap di ScriptProperties (global) karena GAS trigger tidak bisa per-sheet
  // Tapi config aktif/nonaktif disimpan per spreadsheet
  var triggers = ScriptApp.getProjectTriggers();
  var hasTrigger = false;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncShopeeTracking') {
      hasTrigger = true;
      break;
    }
  }

  if (enable) {
    // Buat trigger jika belum ada (shared, 1 trigger untuk semua penyewa)
    if (!hasTrigger) {
      ScriptApp.newTrigger('syncShopeeTracking')
        .timeBased()
        .everyHours(6)
        .create();
    }
    _setSyncConfig('sync_active', 'true');
    if (!_getSyncConfig().lastSync) {
      _setSyncConfig('sync_last', new Date().toISOString());
    }
    return { success: true, message: 'Sync Tracking Shopee diaktifkan (setiap 6 jam).' };
  } else {
    _setSyncConfig('sync_active', 'false');
    return { success: true, message: 'Sync Tracking Shopee dinonaktifkan.' };
  }
}

