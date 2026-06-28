/**
 * STREAMING_CHUNK: Menginisialisasi Entry Point dan Evaluasi Template HTML...
 * Backend Logic untuk WebApp Toko Online Cahya & Embun Florist
 * Terintegrasi dengan Google Sheets sebagai basis data yang aman.
 */

function doGet() {
  return HtmlService.createTemplateFromFile('frontend')
      .evaluate()
      .setTitle('Cahya & Embun Florist - Toko Bunga Bogor')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// SETUP DATABASE OTOMATIS
// Menyiapkan spreadsheet dengan semua tabel/sheet penting dan header default
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const tables = [
    { name: "Users", headers: ["UID", "Nama", "Username", "PasswordHash", "Role", "Kontak"] },
    { name: "Produk", headers: ["ID", "Nama Bunga", "Kategori", "Harga", "Image URL", "Deskripsi", "Stok", "Best Seller"] },
    { name: "Kategori", headers: ["ID", "Nama Kategori"] },
    { name: "Penjualan", headers: ["ID Transaksi", "Pelanggan", "Item Pesanan", "Total Transaksi", "Tanggal", "Status"] },
    { name: "Blogs", headers: ["ID", "Judul", "Ringkasan", "Konten", "Tanggal", "Penulis"] },
    { name: "ProfilToko", headers: ["Atribut", "Nilai"] }
  ];

  tables.forEach(table => {
    let sheet = ss.getSheetByName(table.name);
    if (!sheet) {
      sheet = ss.insertSheet(table.name);
      sheet.appendRow(table.headers);
      
      // Berikan baris styling visual pada header tabel
      sheet.getRange(1, 1, 1, table.headers.length)
           .setFontWeight("bold")
           .setBackground("#10b981")
           .setFontColor("#ffffff");
    }
  });

  // Isi data profil toko default jika masih kosong
  const profilSheet = ss.getSheetByName("ProfilToko");
  if (profilSheet.getLastRow() === 1) {
    profilSheet.appendRow(["Nama Toko", "Cahya & Embun Florist"]);
    profilSheet.appendRow(["Slogan", "Toko Bunga Bogor"]);
    profilSheet.appendRow(["Kontak", "085881654988"]);
    profilSheet.appendRow(["Alamat", "Jl. Cimanggu Barata, No. 82, RT 03 RW 04, Kelurahan Kedung Jaya, Kecamatan Tanah Sareal, Kota Bogor, 16164"]);
  }

  // Buat akun Default otomatis (Admin & Peserta/Pelanggan) jika kosong
  const userSheet = ss.getSheetByName("Users");
  if (userSheet.getLastRow() === 1) {
    // Admin setup (username: admin, password: edudigital)
    userSheet.appendRow(["U001", "Pengelola Toko", "admin", "edudigital", "Admin", "085881654988"]);
    // Peserta setup (username: peserta, password: edudigital)
    userSheet.appendRow(["U002", "Rian Wijaya", "peserta", "edudigital", "Peserta", "081234567890"]);
  }
}

/**
 * STREAMING_CHUNK: Merancang Validasi Login & Keamanan Pengguna...
 */
// VERIFIKASI LOGIN
// Membandingkan data input dari frontend dengan record aman di tabel Users
function verifyLogin(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    
    // Looping record mulai baris ke-2 (melewati header)
    for (let i = 1; i < data.length; i++) {
      const dbUser = data[i][2];
      const dbPass = data[i][3];
      const dbRole = data[i][4];
      const dbName = data[i][1];
      
      if (dbUser === username && dbPass === password) {
        return { success: true, role: dbRole, name: dbName };
      }
    }
    return { success: false, msg: "Kombinasi Username/Password salah." };
  } catch (e) {
    return { success: false, msg: "Kesalahan server: " + e.toString() };
  }
}

/**
 * STREAMING_CHUNK: Menyediakan API CRUD Produk dan Sinkronisasi Dinamis...
 */
// GET SINKRONISASI DATA UTAMA
function getDatabaseSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  return {
    categories: getSheetDataAsJson(ss.getSheetByName("Kategori")),
    products: getSheetDataAsJson(ss.getSheetByName("Produk")),
    blogs: getSheetDataAsJson(ss.getSheetByName("Blogs")),
    sales: getSheetDataAsJson(ss.getSheetByName("Penjualan"))
  };
}

// HELPER: Convert sheet data to clean JSON Array
function getSheetDataAsJson(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    headers.forEach((header, colIdx) => {
      obj[header.replace(/\s+/g, '_').toLowerCase()] = row[colIdx];
    });
    results.push(obj);
  }
  return results;
}

// SINKRONISASI PENJUALAN BARU
function addSale(saleData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Penjualan");
    sheet.appendRow([
      saleData.id,
      saleData.user,
      saleData.items,
      saleData.total,
      saleData.date,
      saleData.status
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, msg: e.toString() };
  }
}

// SINKRONISASI MANAJEMEN PRODUK (ADD / EDIT / DELETE)
function saveProductBackend(prod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Produk");
  const data = sheet.getDataRange().getValues();
  
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === prod.id) {
      foundRow = i + 1; // Baris riil di Excel (1-based index)
      break;
    }
  }

  const rowValues = [
    prod.id,
    prod.name,
    prod.category,
    prod.price,
    prod.img,
    prod.desc,
    prod.stock,
    prod.isBest ? "Ya" : "Tidak"
  ];

  if (foundRow !== -1) {
    // Aksi Update
    sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    // Aksi Insert
    sheet.appendRow(rowValues);
  }
  return { success: true };
}

function deleteProductBackend(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Produk");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, msg: "Produk tidak ditemukan." };
}