// ==========================================================
// BACKEND GOOGLE APPS SCRIPT - BUBI PROJECT (HPP CALCULATOR)
// ==========================================================

function getDb() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Bubi project - HPP Calculator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper untuk load file HTML modular
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}

// --- AMBIL SEMUA DATA UNTUK INISIALISASI APLIKASI ---
function getInitialData() {
  const db = getDb();
  return {
    bahan: getSheetDataAsObject(db.getSheetByName('Master_Bahan')),
    mesin: getSheetDataAsObject(db.getSheetByName('Master_Mesin')),
    produk: getSheetDataAsObject(db.getSheetByName('Master_Produk'))
  };
}

function getSheetDataAsObject(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

// ==========================================
// HELPER: GENERATE AUTO INCREMENT ID (PRD-001, BHN-001, Dst)
// ==========================================
function generateAutoId(sheetName, prefix) {
  const ss = getDb();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return prefix + "-001";
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return prefix + "-001";
  }
  
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;
  
  ids.forEach(function(row) {
    const idStr = String(row[0]);
    if (idStr && idStr.indexOf(prefix + "-") !== -1) {
      const parts = idStr.split("-");
      const num = parseInt(parts[parts.length - 1], 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  });
  
  const nextNum = maxNum + 1;
  return prefix + "-" + String(nextNum).padStart(3, '0');
}

// Helper untuk konversi currency/string ke angka murni
function parseCurrencyToNumber(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  
  var str = String(val).trim();
  var cleaned = str.replace(/Rp|\s/g, '').replace(/\./g, '').replace(',', '.');
  var num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

// --- 1. SIMPAN MASTER BAHAN BAKU ---
function simpanBahanKeSheet(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getDb();
    var sheet = ss.getSheetByName('Master_Bahan');
    
    var idBahan = generateAutoId('Master_Bahan', 'BHN');
    var hargaBeli = parseFloat(payload.hargaBeli) || 0;
    var p = parseFloat(payload.p) || 0;
    var l = parseFloat(payload.l) || 0;
    var volume = parseFloat(payload.volume) || 0;
    var berat = parseFloat(payload.berat) || 0;
    var pcs = parseFloat(payload.pcs) || 0;

    var hargaSatuanDasar = 0;
    if (payload.tipe === 'Lembaran' && p > 0 && l > 0) {
      hargaSatuanDasar = hargaBeli / (p * l);
    } else if (payload.tipe === 'Panjang' && p > 0) {
      hargaSatuanDasar = hargaBeli / p;
    } else if (payload.tipe === 'Cairan' && volume > 0) {
      hargaSatuanDasar = hargaBeli / volume;
    } else if (payload.tipe === 'Satuan' && pcs > 0) {
      hargaSatuanDasar = hargaBeli / pcs;
    } else if (payload.tipe === 'Berat' && berat > 0) {
      hargaSatuanDasar = hargaBeli / berat;
    } else {
      hargaSatuanDasar = hargaBeli;
    }

    sheet.appendRow([
      idBahan,
      payload.nama,
      payload.tipe,
      payload.tipe,
      hargaBeli,
      p,
      l,
      volume,
      berat,
      pcs,
      Number(hargaSatuanDasar)
    ]);

    return { status: "success", message: "Bahan berhasil disimpan dengan ID: " + idBahan };
  } catch (e) {
    return { status: "error", message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- 2. SIMPAN MASTER MESIN ---
function simpanMesinKeSheet(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getDb();
    var sheet = ss.getSheetByName('Master_Mesin');
    
    var idMesin = generateAutoId('Master_Mesin', 'MSN');
    var biayaDepresiasiPerJam = payload.umurJam > 0 ? (payload.hargaBeli / payload.umurJam) : 0;
    var biayaListrikPerJam = (payload.watt / 1000) * payload.tarifKwh;
    var totalBiayaPerJam = biayaDepresiasiPerJam + biayaListrikPerJam;
    var biayaPerMenit = totalBiayaPerJam / 60;

    sheet.appendRow([
      idMesin,
      payload.nama,
      payload.hargaBeli,
      payload.umurJam,
      payload.watt,
      payload.tarifKwh,
      biayaPerMenit
    ]);

    return { status: "success", message: "Mesin berhasil disimpan dengan ID: " + idMesin };
  } catch (e) {
    return { status: "error", message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- 3. SIMPAN PRODUK LENGKAP BERSAMA BOM ---
function simpanProdukKeSheet(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = getDb();
    
    const sheetMaster = ss.getSheetByName('Master_Produk');
    const idProduk = generateAutoId('Master_Produk', 'PRD');
    const kodeProduk = payload.kodeProduk || idProduk;
    
    sheetMaster.appendRow([
      idProduk,
      kodeProduk,
      payload.namaProduk,
      payload.marginTarget || payload.marginTargetPersen || 0,
      parseCurrencyToNumber(payload.totalHPP),
      parseCurrencyToNumber(payload.hargaJual),
      new Date()
    ]);

    const sheetDetailBahan = ss.getSheetByName('Detail_BOM_Bahan');
    if (payload.detailBahan && payload.detailBahan.length > 0) {
      payload.detailBahan.forEach(function(item) {
        const idDetailBahan = generateAutoId('Detail_BOM_Bahan', 'BOMB');
        const subtotalBiaya = parseCurrencyToNumber(item.subtotal);

        sheetDetailBahan.appendRow([
          idDetailBahan,
          idProduk,
          item.idBahan,
          item.panjang || item.panjangPakai || 0,
          item.lebar || item.lebarPakai || 0,
          item.pemakaianMlGram || item.pemakaian || 0,
          item.qtyLayerPcs || 1,
          item.bahanTerbuang || item.wastePersen || 0,
          subtotalBiaya
        ]);
      });
    }

    const sheetDetailProses = ss.getSheetByName('Detail_BOM_Proses');
    if (payload.detailProses && payload.detailProses.length > 0) {
      payload.detailProses.forEach(function(item) {
        const idDetailProses = generateAutoId('Detail_BOM_Proses', 'BOMP');
        const subtotalBiaya = parseCurrencyToNumber(item.subtotal);

        sheetDetailProses.appendRow([
          idDetailProses,
          idProduk,
          item.idMesin || "-",
          item.namaProses,
          item.durasi || item.durasiMenit || 0,
          item.laborPerJam || item.tarifLaborJam || 0,
          subtotalBiaya
        ]);
      });
    }

    return { status: "success", message: "Data produk berhasil disimpan dengan ID: " + idProduk };

  } catch (e) {
    return { status: "error", message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- 4. AMBIL DATA MASTER PRODUK / KATALOG ---
function getMasterProdukData() {
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName('Master_Produk');
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];

    const result = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[0] || row[1]) {
        result.push({
          id: row[0],
          kodeProduk: row[1],
          namaProduk: row[2],
          marginPersen: row[3],
          totalHpp: row[4],
          hargaJual: row[5],
          tanggal: row[6] || ''
        });
      }
    }
    return result;
  } catch (e) {
    throw new Error("Gagal mengambil data katalog: " + e.message);
  }
}

// --- 5. HAPUS DATA DARI SHEET BASED ON ID ---
function hapusDataFromSheet(sheetName, idKey, idValue) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDb();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { status: "error", message: "Sheet tidak ditemukan" };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: "error", message: "Data kosong" };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idValue)) {
        sheet.deleteRow(i + 1);
        return { status: "success", message: "Data berhasil dihapus" };
      }
    }
    return { status: "error", message: "ID tidak ditemukan" };
  } catch (e) {
    return { status: "error", message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- 6. UPDATE MASTER BAHAN ---
function updateBahanKeSheet(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDb();
    const sheet = ss.getSheetByName('Master_Bahan');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.idBahan)) {
        const row = i + 1;
        var hargaBeli = parseFloat(payload.hargaBeli) || 0;
        var p = parseFloat(payload.p) || 0;
        var l = parseFloat(payload.l) || 0;
        var volume = parseFloat(payload.volume) || 0;
        var berat = parseFloat(payload.berat) || 0;
        var pcs = parseFloat(payload.pcs) || 0;

        var hargaSatuanDasar = 0;
        if (payload.tipe === 'Lembaran' && p > 0 && l > 0) {
          hargaSatuanDasar = hargaBeli / (p * l);
        } else if (payload.tipe === 'Panjang' && p > 0) {
          hargaSatuanDasar = hargaBeli / p;
        } else if (payload.tipe === 'Cairan' && volume > 0) {
          hargaSatuanDasar = hargaBeli / volume;
        } else if (payload.tipe === 'Satuan' && pcs > 0) {
          hargaSatuanDasar = hargaBeli / pcs;
        } else if (payload.tipe === 'Berat' && berat > 0) {
          hargaSatuanDasar = hargaBeli / berat;
        } else {
          hargaSatuanDasar = hargaBeli;
        }

        sheet.getRange(row, 2, 1, 10).setValues([[
          payload.nama, payload.tipe, payload.tipe, hargaBeli,
          p, l, volume, berat, pcs, Number(hargaSatuanDasar)
        ]]);

        return { status: "success", message: "Bahan berhasil diperbarui" };
      }
    }
    return { status: "error", message: "Data tidak ditemukan" };
  } catch (e) {
    return { status: "error", message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- 7. UPDATE MASTER MESIN ---
function updateMesinKeSheet(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDb();
    const sheet = ss.getSheetByName('Master_Mesin');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.idMesin)) {
        const row = i + 1;
        var biayaDepresiasiPerJam = payload.umurJam > 0 ? (payload.hargaBeli / payload.umurJam) : 0;
        var biayaListrikPerJam = (payload.watt / 1000) * payload.tarifKwh;
        var totalBiayaPerJam = biayaDepresiasiPerJam + biayaListrikPerJam;
        var biayaPerMenit = totalBiayaPerJam / 60;

        sheet.getRange(row, 2, 1, 6).setValues([[
          payload.nama, payload.hargaBeli, payload.umurJam,
          payload.watt, payload.tarifKwh, biayaPerMenit
        ]]);

        return { status: "success", message: "Mesin berhasil diperbarui" };
      }
    }
    return { status: "error", message: "Data tidak ditemukan" };
  } catch (e) {
    return { status: "error", message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MASTER MARKETPLACE
// ==========================================

function getMarketplacePrograms() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master_Marketplace");
    
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    // Ambil 6 kolom (A sampai F)
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); 
    
    const result = data.map((row, index) => {
      const statusRaw = String(row[5] || "").toUpperCase().trim();
      return {
        rowId: index + 2,
        idProgram: row[0],                                                           // Kolom A: ID
        channel: String(row[1] || "").trim(),                                        // Kolom B: Channel
        namaProgram: row[2],                                                         // Kolom C: Nama Program
        persenBiaya: parseFloat(row[3]) || 0,                                         // Kolom D: Potongan Persen
        fixedFee: parseFloat(row[4]) || 0,                                            // Kolom E: Fixed Fee
        statusAktif: (row[5] === true || statusRaw === "AKTIF" || statusRaw === "TRUE") // Kolom F: Status
      };
    });

    // URUTKAN BERDASARKAN CHANNEL (A-Z)
    result.sort((a, b) => a.channel.localeCompare(b.channel, 'id', { sensitivity: 'base' }));

    return result;
  } catch (err) {
    Logger.log("Error getMarketplacePrograms: " + err.message);
    return [];
  }
}

function simpanProgramMarketplace(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Master_Marketplace");
    
    if (!sheet) {
      sheet = ss.insertSheet("Master_Marketplace");
      sheet.appendRow(["ID", "Channel", "Nama_Program", "Potongan Persen (%)", "Fixed Fee (Rp)", "Status (Aktif / Non-Aktif)"]);
    }
    
    const statusStr = data.statusAktif ? 'AKTIF' : 'NON-AKTIF';
    
    if (data.rowId) {
      // UPDATE DATA EKSISTING (Mulai Kolom B / Index 2 sampai F / Status)
      // Kolom A (ID) tetap utuh, tidak tertimpa
      sheet.getRange(data.rowId, 2, 1, 5).setValues([[
        data.channel, 
        data.namaProgram, 
        data.persenBiaya, 
        data.fixedFee, 
        statusStr
      ]]);
    } else {
      // TAMBAH DATA BARU
      const lastRow = sheet.getLastRow();
      
      // BIKIN ID RAPI BERURUTAN (Contoh: MP-101, MP-102, dst.)
      // Jika lastRow = 1 (header), ID baru = MP-101. Jika lastRow = 4, ID baru = MP-104.
      const nextNumber = 100 + lastRow; 
      const newId = "MP-" + nextNumber;
      
      sheet.appendRow([
        newId, 
        data.channel, 
        data.namaProgram, 
        data.persenBiaya, 
        data.fixedFee, 
        statusStr
      ]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function hapusProgramMarketplace(rowId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master_Marketplace");
    if (!sheet) return { success: false, message: "Sheet tidak ditemukan" };

    // Hapus baris berdasarkan posisi fisik rowId di sheet
    sheet.deleteRow(rowId);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function toggleStatusMarketplace(rowId, isAktif) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master_Marketplace");
    if (!sheet) return { success: false, message: "Sheet tidak ditemukan" };

    const statusStr = isAktif ? 'AKTIF' : 'NON-AKTIF';
    
    // PERBAIKAN: Tembak Kolom 6 (Kolom F / Status), bukan Kolom 5!
    sheet.getRange(rowId, 6).setValue(statusStr); 
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// Katalog & Orderan (Lengkap dengan Edit & Hapus)
// ==========================================

// 1. Fungsi Mengambil Data Master Produk & Master Order
function getKatalogOrderData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // A. Ambil Data Master_Produk
  var sheetProduk = ss.getSheetByName('Master_Produk');
  var listMasterProduk = [];
  if (sheetProduk) {
    var dataProduk = sheetProduk.getDataRange().getValues();
    for (var i = 1; i < dataProduk.length; i++) {
      if (dataProduk[i][0]) { 
        listMasterProduk.push({
          idProduk: dataProduk[i][0],
          kodeProduk: dataProduk[i][1] || dataProduk[i][0],
          namaProduk: dataProduk[i][2],
          marginTarget: dataProduk[i][3],
          totalHPP: dataProduk[i][4],
          hargaJualRekomendasi: dataProduk[i][5] || dataProduk[i][4]
        });
      }
    }
  }

  // B. Ambil Data Master_Order
  var sheetOrder = ss.getSheetByName('Master_Order');
  var listMasterOrder = [];
  if (sheetOrder) {
    var dataOrder = sheetOrder.getDataRange().getValues();
    for (var j = 1; j < dataOrder.length; j++) {
      if (dataOrder[j][1]) { 
        listMasterOrder.push({
          tanggal: dataOrder[j][0],
          idOrder: dataOrder[j][1],
          kodeProduk: dataOrder[j][2],
          namaProduk: dataOrder[j][3],
          channel: dataOrder[j][4],
          qty: dataOrder[j][5],
          totalHPP: dataOrder[j][6],
          totalLabaBersih: dataOrder[j][7]
        });
      }
    }
  }

  return {
    masterProduk: listMasterProduk,
    masterOrder: listMasterOrder
  };
}

// 2. Fungsi Menyimpan & Mengedit Orderan di Sheet 'Master_Order'
function simpanOrderKeSheet(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master_Order');
  
  // Jika sheet belum ada, buat otomatis beserta header
  if (!sheet) {
    sheet = ss.insertSheet('Master_Order');
    sheet.appendRow([
      'Tanggal', 
      'ID_Order', 
      'Kode_Produk', 
      'Nama_Produk', 
      'Channel_Penjualan', 
      'Qty', 
      'Total_HPP', 
      'Total_Laba_Bersih'
    ]);
  }

  var dataOrder = sheet.getDataRange().getValues();
  var barisTarget = -1;

  // Cek apakah ini transaksi EDIT (cari baris berdasarkan ID Order)
  if (payload.isEdit) {
    for (var i = 1; i < dataOrder.length; i++) {
      if (dataOrder[i][1].toString() === payload.idOrder.toString()) { // Kolom ID_Order ada di indeks 1 (Kolom B)
        barisTarget = i + 1; // +1 karena index spreadsheet dimulai dari 1
        break;
      }
    }
  }

  var rowData = [
    payload.tanggal,
    payload.idOrder,
    payload.kodeProduk,
    payload.namaProduk,
    payload.channel,
    payload.qty,
    payload.totalHPP,
    payload.totalLabaBersih
  ];

  if (barisTarget > 0) {
    // JIKA EDIT: Timpa/update data pada baris lama
    sheet.getRange(barisTarget, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // JIKA BARU: Tambahkan baris baru paling bawah
    sheet.appendRow(rowData);
  }
  
  return { status: 'success' };
}

// 3. FUNGSI BARU: Menghapus Orderan Berdasarkan ID Order
function hapusOrderDariSheet(idOrder) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master_Order');
  
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };

  var dataOrder = sheet.getDataRange().getValues();
  var barisYangDihapus = -1;

  // Cari posisi baris berdasarkan ID Order
  for (var i = 1; i < dataOrder.length; i++) {
    if (dataOrder[i][1].toString() === idOrder.toString()) {
      barisYangDihapus = i + 1;
      break;
    }
  }

  if (barisYangDihapus > 0) {
    sheet.deleteRow(barisYangDihapus);
    return { status: 'success' };
  } else {
    throw new Error("ID Order tidak ditemukan di database!");
  }
}
