/**
 * ZettBOT Backend - Kasir Warteg Modern
 * Backend logika bisnis, CRUD, filtering anti-lag, pagination, Super Admin, & PIN Management.
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Kasir Warteg Modern - Zettbos POS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}

/**
 * PIN ADMIN MANAGEMENT VIA PROPERTIES SERVICE
 */
function getAdminPin() {
  var props = PropertiesService.getScriptProperties();
  var pin = props.getProperty('ADMIN_PIN');
  if (!pin) {
    pin = 'biskuitmeong';
    props.setProperty('ADMIN_PIN', pin);
  }
  return pin;
}

function verifyAdminPinBackend(inputPin) {
  var currentPin = getAdminPin();
  if (inputPin === currentPin) {
    return { success: true, message: 'Login Super Admin Berhasil!' };
  } else {
    return { success: false, message: 'PIN Super Admin Salah!' };
  }
}

function updateAdminPin(oldPin, newPin) {
  try {
    var currentPin = getAdminPin();
    if (String(oldPin).trim() !== String(currentPin).trim()) {
      return { success: false, message: 'PIN Lama Salah!' };
    }
    if (!newPin || String(newPin).trim().length < 4) {
      return { success: false, message: 'PIN Baru minimal 4 karakter!' };
    }
    
    PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', String(newPin).trim());
    return { success: true, message: 'PIN Super Admin Berhasil Diperbarui!' };
  } catch (err) {
    return { success: false, message: 'Gagal memperbarui PIN: ' + err.toString() };
  }
}

function generateTransactionId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Transaksi');
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd');
  var prefix = 'TRX-' + dateStr + '-';
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return prefix + '0001';
  }
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  var todayCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][0] || '');
    if (id.indexOf(prefix) === 0) {
      todayCount++;
    }
  }
  
  var nextSeq = todayCount + 1;
  var seqFormatted = ('0000' + nextSeq).slice(-4);
  return prefix + seqFormatted;
}

function getMenuItems(category) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Menu');
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return [];
    
    var items = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var item = {
        menuId: String(row[0] || ''),
        namaMenu: String(row[1] || ''),
        kategori: String(row[2] || ''),
        harga: Number(row[3] || 0),
        fotoUrl: String(row[4] || ''),
        status: String(row[5] || 'Tersedia')
      };
      
      if (!category || category === 'All' || item.kategori.toLowerCase() === category.toLowerCase()) {
        items.push(item);
      }
    }
    return items;
  } catch (err) {
    return [];
  }
}

function saveMenuItem(menuData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Menu');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Menu');
    }
    
    var data = sheet.getDataRange().getDisplayValues();
    var existingRowIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(menuData.menuId)) {
        existingRowIndex = i + 1;
        break;
      }
    }
    
    var rowValues = [
      String(menuData.menuId || ('MNU-' + Math.floor(100 + Math.random() * 900))),
      String(menuData.namaMenu || ''),
      String(menuData.kategori || 'Makanan'),
      Number(menuData.harga || 0),
      String(menuData.fotoUrl || 'https://images.unsplash.com/photo-1546069901-ba9599a7e63c?w=300'),
      String(menuData.status || 'Tersedia')
    ];
    
    if (existingRowIndex > 0) {
      sheet.getRange(existingRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: 'Menu berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function toggleMenuStatus(menuId, newStatus) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Menu');
    if (!sheet) return { success: false, message: 'Sheet Menu tidak ditemukan.' };
    
    var data = sheet.getDataRange().getDisplayValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(menuId)) {
        sheet.getRange(i + 1, 6).setValue(String(newStatus));
        SpreadsheetApp.flush();
        return { success: true, message: 'Status menu diperbarui!' };
      }
    }
    return { success: false, message: 'Menu tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveTransaction(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Transaksi');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Transaksi');
    }
    
    var trxId = generateTransactionId();
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    
    var newRow = [
      trxId,
      timestamp,
      Number(payload.totalBayar || 0),
      String(payload.metodePembayaran || 'Cash'),
      Number(payload.cashDibayar || 0),
      Number(payload.kembalian || 0),
      JSON.stringify(payload.detailPelanggan || []),
      String(payload.catatanPromo || '-')
    ];
    
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    
    return {
      success: true,
      transactionId: trxId,
      message: 'Transaksi berhasil disimpan!'
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getDashboardStats() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Transaksi');
    if (!sheet) return { omsetHariIni: 0, totalTransaksi: 0, makananTerlaris: '-', pctCash: 50, pctQris: 50 };
    
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { omsetHariIni: 0, totalTransaksi: 0, makananTerlaris: '-', pctCash: 50, pctQris: 50 };
    
    var todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');
    var omset = 0;
    var countToday = 0;
    var cashCount = 0;
    var qrisCount = 0;
    var itemCounts = {};
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowDate = String(row[1] || '').split(' ')[0];
      var total = Number(row[2] || 0);
      var method = String(row[3] || '');
      
      if (rowDate === todayStr) {
        omset += total;
        countToday++;
      }
      
      if (method.toUpperCase() === 'CASH') cashCount++;
      if (method.toUpperCase() === 'QRIS') qrisCount++;
      
      try {
        var details = JSON.parse(row[6] || '[]');
        details.forEach(function(cust) {
          if (cust.cart && Array.isArray(cust.cart)) {
            cust.cart.forEach(function(it) {
              itemCounts[it.namaMenu] = (itemCounts[it.namaMenu] || 0) + (Number(it.qty) || 1);
            });
          }
        });
      } catch (e) {}
    }
    
    var topItem = '-';
    var maxQty = 0;
    for (var name in itemCounts) {
      if (itemCounts[name] > maxQty) {
        maxQty = itemCounts[name];
        topItem = name + ' (' + maxQty + ')';
      }
    }
    
    var totalMethods = cashCount + qrisCount;
    var pctCash = totalMethods > 0 ? Math.round((cashCount / totalMethods) * 100) : 50;
    var pctQris = totalMethods > 0 ? Math.round((qrisCount / totalMethods) * 100) : 50;
    
    return {
      omsetHariIni: omset,
      totalTransaksi: countToday,
      makananTerlaris: topItem,
      pctCash: pctCash,
      pctQris: pctQris
    };
  } catch (err) {
    return { omsetHariIni: 0, totalTransaksi: 0, makananTerlaris: '-', pctCash: 50, pctQris: 50 };
  }
}

function getTransactionHistory(page, pageSize, searchQuery, dateFilter) {
  try {
    page = page || 1;
    pageSize = pageSize || 15;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Transaksi');
    if (!sheet) return { transactions: [], totalPages: 0, currentPage: 1, totalCount: 0 };
    
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { transactions: [], totalPages: 0, currentPage: 1, totalCount: 0 };
    
    var filtered = [];
    var search = (searchQuery || '').toLowerCase();
    var dateVal = (dateFilter || '').trim();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var trxId = String(row[0] || '');
      var dateStr = String(row[1] || '');
      var total = Number(row[2] || 0);
      var method = String(row[3] || '');
      var cash = Number(row[4] || 0);
      var change = Number(row[5] || 0);
      var details = String(row[6] || '[]');
      var promo = String(row[7] || '');

      var matchSearch = !search || trxId.toLowerCase().indexOf(search) >= 0 || method.toLowerCase().indexOf(search) >= 0 || details.toLowerCase().indexOf(search) >= 0;
      var matchDate = !dateVal || dateStr.indexOf(dateVal) >= 0;

      if (matchSearch && matchDate) {
        filtered.push({
          transactionId: trxId,
          tanggalJam: dateStr,
          totalBayar: total,
          metodePembayaran: method,
          cashDibayar: cash,
          kembalian: change,
          detailPelanggan: details,
          catatanPromo: promo
        });
      }
    }

    // Sort descending (terbaru di atas)
    filtered.reverse();

    var totalCount = filtered.length;
    var totalPages = Math.ceil(totalCount / pageSize) || 1;
    var startIndex = (page - 1) * pageSize;
    var paginated = filtered.slice(startIndex, startIndex + pageSize);

    return {
      transactions: paginated,
      totalPages: totalPages,
      currentPage: page,
      totalCount: totalCount
    };
  } catch (err) {
    return { transactions: [], totalPages: 0, currentPage: 1, totalCount: 0, error: err.toString() };
  }
}

function updateAdminPin(oldPin, newPin) {
  return updateAdminPinBackend(oldPin, newPin);
}

function backupDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuSheet = ss.getSheetByName('Menu');
    var trxSheet = ss.getSheetByName('Transaksi');
    
    var backupData = {
      timestamp: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss'),
      menu: menuSheet ? menuSheet.getDataRange().getDisplayValues() : [],
      transaksi: trxSheet ? trxSheet.getDataRange().getDisplayValues() : []
    };
    
    return {
      success: true,
      backupJson: JSON.stringify(backupData, null, 2),
      message: 'Backup database berhasil dibuat!'
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function restoreDatabase(jsonString) {
  try {
    var parsed = JSON.parse(jsonString);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (parsed.menu && Array.isArray(parsed.menu) && parsed.menu.length > 0) {
      var menuSheet = ss.getSheetByName('Menu');
      if (!menuSheet) menuSheet = ss.insertSheet('Menu');
      menuSheet.clear();
      menuSheet.getRange(1, 1, parsed.menu.length, parsed.menu[0].length).setValues(parsed.menu);
    }
    
    if (parsed.transaksi && Array.isArray(parsed.transaksi) && parsed.transaksi.length > 0) {
      var trxSheet = ss.getSheetByName('Transaksi');
      if (!trxSheet) trxSheet = ss.insertSheet('Transaksi');
      trxSheet.clear();
      trxSheet.getRange(1, 1, parsed.transaksi.length, parsed.transaksi[0].length).setValues(parsed.transaksi);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: 'Database berhasil dipulihkan (Restore Sukses)!' };
  } catch (err) {
    return { success: false, message: 'Gagal Restore: ' + err.toString() };
  }
}

function resetTransactionDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Transaksi');
    if (sheet) {
      sheet.clear();
      var trxHeaders = ['Transaction_ID', 'Tanggal_Jam', 'Total_Bayar', 'Metode_Pembayaran', 'Cash_Dibayar', 'Kembalian', 'Detail_Pelanggan_JSON', 'Catatan_Promo'];
      sheet.appendRow(trxHeaders);
      sheet.getRange(1, 1, 1, trxHeaders.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
      SpreadsheetApp.flush();
    }
    return { success: true, message: 'Riwayat transaksi berhasil dikosongkan!' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
