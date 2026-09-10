/**
 * ZettBOT Apps Script Assistant - Safe Database Setup
 * Memastikan struktur Spreadsheet siap tanpa menghapus data eksisting (Safe Migrate).
 */

function setupDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Setup Sheet Menu
    var menuSheet = ss.getSheetByName('Menu');
    var menuHeaders = ['Menu_ID', 'Nama_Menu', 'Kategori', 'Harga', 'Foto_URL', 'Status'];
    
    if (!menuSheet) {
      menuSheet = ss.insertSheet('Menu');
      menuSheet.appendRow(menuHeaders);
      menuSheet.getRange(1, 1, 1, menuHeaders.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
      
      // Seed Data Dummy Awal Lengkap (Makanan, Minuman, Tambahan, Cafe)
      var dummyMenu = [
        ['MNU-001', 'Nasi Rames Rendang', 'Makanan', 22000, 'https://images.unsplash.com/photo-1626777552726-4a6b54c97e46?w=300', 'Tersedia'],
        ['MNU-002', 'Ayam Goreng Lengkuas', 'Makanan', 18000, 'https://images.unsplash.com/photo-1626082927389-6cd097cdc6ec?w=300', 'Tersedia'],
        ['MNU-003', 'Tahu Tempe Bacem', 'Makanan', 5000, 'https://images.unsplash.com/photo-1546069901-ba9599a7e63c?w=300', 'Tersedia'],
        ['MNU-004', 'Telur Balado', 'Makanan', 7000, 'https://images.unsplash.com/photo-1512621776951-a57141f2eefd?w=300', 'Tersedia'],
        ['MNU-005', 'Es Teh Manis', 'Minuman', 5000, 'https://images.unsplash.com/photo-1556679343-c7306c1976bc?w=300', 'Tersedia'],
        ['MNU-006', 'Es Jeruk Peras', 'Minuman', 7000, 'https://images.unsplash.com/photo-1613478223719-2ab802602423?w=300', 'Tersedia'],
        ['MNU-007', 'Teh Hangat', 'Minuman', 4000, 'https://images.unsplash.com/photo-1576092768241-dec231879fc3?w=300', 'Tersedia'],
        ['MNU-008', 'Kerupuk', 'Tambahan', 2000, 'https://images.unsplash.com/photo-1599487488170-d11ec9c172f0?w=300', 'Tersedia'],
        ['MNU-009', 'Gorengan', 'Tambahan', 2000, 'https://images.unsplash.com/photo-1626082927389-6cd097cdc6ec?w=300', 'Tersedia'],
        ['MNU-010', 'Pisang', 'Tambahan', 3000, 'https://images.unsplash.com/photo-1571771894821-ce9b6c11b08e?w=300', 'Tersedia'],
        ['MNU-011', 'Kopi Kapal Api', 'Cafe', 5000, 'https://images.unsplash.com/photo-1514432324607-a09d9b4aefdd?w=300', 'Tersedia'],
        ['MNU-012', 'Indocafe', 'Cafe', 5000, 'https://images.unsplash.com/photo-1541167760496-1628856ab772?w=300', 'Tersedia'],
        ['MNU-013', 'Kopi ABC Susu', 'Cafe', 5000, 'https://images.unsplash.com/photo-1572442388796-11668ba67e53?w=300', 'Tersedia'],
        ['MNU-014', 'Nutrisari', 'Cafe', 5000, 'https://images.unsplash.com/photo-1613478223719-2ab802602423?w=300', 'Tersedia']
      ];
      menuSheet.getRange(2, 1, dummyMenu.length, menuHeaders.length).setValues(dummyMenu);
    } else {
      // Safe Migrate: Pastikan Header Sesuai
      var currentHeaders = menuSheet.getRange(1, 1, 1, menuHeaders.length).getDisplayValues()[0];
      if (currentHeaders.length === 0 || currentHeaders[0] === '') {
        menuSheet.getRange(1, 1, 1, menuHeaders.length).setValues([menuHeaders]);
      }
    }
    
    // 2. Setup Sheet Transaksi
    var trxSheet = ss.getSheetByName('Transaksi');
    var trxHeaders = ['Transaction_ID', 'Tanggal_Jam', 'Total_Bayar', 'Metode_Pembayaran', 'Cash_Dibayar', 'Kembalian', 'Detail_Pelanggan_JSON', 'Catatan_Promo'];
    
    if (!trxSheet) {
      trxSheet = ss.insertSheet('Transaksi');
      trxSheet.appendRow(trxHeaders);
      trxSheet.getRange(1, 1, 1, trxHeaders.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
    } else {
      var currentTrxHeaders = trxSheet.getRange(1, 1, 1, trxHeaders.length).getDisplayValues()[0];
      if (currentTrxHeaders.length === 0 || currentTrxHeaders[0] === '') {
        trxSheet.getRange(1, 1, 1, trxHeaders.length).setValues([trxTrxHeaders]);
      }
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: 'Database Warteg Modern berhasil disiapkan (Safe Migrate Selesai)!' };
  } catch (error) {
    return { success: false, message: 'Gagal setup database: ' + error.toString() };
  }
}
