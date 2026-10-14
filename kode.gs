var SPREADSHEET_ID = '1KjMNJpSEXwip3mBN5iF2ebYbSyS42onLuBIKKYqPwcw';
var SHEET_NAME = 'Sheet1'; 
var EMAIL_DPRD = 'unangamaludin@gmail.com'; // ⚠ GANTI dengan email resmi DPRD Kabupaten Bogor

function doGet(e) {
  // 1. Pengaman jika fungsi dijalankan manual dari editor
  e = e || { parameter: {} };

  // 2. Jika URL mengandung parameter ?page=admin, tampilkan halaman admin
  if (e.parameter && e.parameter.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin.html')
        .setTitle('Dashboard Admin e-Tamu DPRD Kabupaten Bogor')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // 3. Default: Buka form pendaftaran
  return HtmlService.createHtmlOutputFromFile('Index.html')
      .setTitle('e-Tamu DPRD Kabupaten Bogor')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// FUNGSI STATISTIK (UNTUK HALAMAN INDEX)
// ==========================================
/**
 * Mengambil total orang (jumlah tamu) dan total agenda (jumlah baris) untuk HARI INI
 */
function getStatistikHariIni() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    var data = sheet.getDataRange().getDisplayValues();
    
    // Format tanggal hari ini dalam WIB (Asia/Jakarta) dengan format YYYY-MM-DD
    var todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    
    var totalAgenda = 0;
    var totalTamu = 0;
    
    // Looping dari baris ke-2 (mengabaikan header baris 1)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var jumlahTamuRombongan = parseInt(row[3]) || 1; // Kolom D (Jumlah Tamu, indeks 3)
      var tglKunjunganRaw = row[7];                   // Kolom H (Tanggal Kunjungan, indeks 7)
      
      if (tglKunjunganRaw) {
        var tglFormatted = "";
        
        // Pencocokan format tanggal YYYY-MM-DD atau DD/MM/YYYY
        if (tglKunjunganRaw.indexOf('-') !== -1) {
          tglFormatted = tglKunjunganRaw.split('T')[0]; // Ambil YYYY-MM-DD
        } else if (tglKunjunganRaw.indexOf('/') !== -1) {
          var parts = tglKunjunganRaw.split('/');
          if (parts.length === 3) {
            // Jika format DD/MM/YYYY -> diubah ke YYYY-MM-DD
            tglFormatted = parts[2] + '-' + ('0' + parts[1]).slice(-2) + '-' + ('0' + parts[0]).slice(-2);
          }
        }
        
        if (tglFormatted === todayStr) {
          totalAgenda += 1;
          totalTamu += jumlahTamuRombongan;
        }
      }
    }
    
    return {
      totalAgenda: totalAgenda,
      totalTamu: totalTamu
    };
  } catch (err) {
    Logger.log("Gagal mengambil statistik: " + err.toString());
    return { totalAgenda: 0, totalTamu: 0 };
  }
}

// ==========================================
// FUNGSI UNTUK ADMIN
// ==========================================
function getDataTamu() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  
  // Membaca data sebagai string agar rapi saat ditampilkan di HTML
  return sheet.getDataRange().getDisplayValues(); 
}

/**
 * Mengubah status permohonan (Setujui / Tolak) dan mengirim email otomatis ke Tamu
 */
function updateStatusTamu(rowIndex, statusBaru, alasanPenolakan) {
  alasanPenolakan = alasanPenolakan || "";
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    // Kolom 12 adalah kolom Status (L)
    sheet.getRange(rowIndex, 12).setValue(statusBaru);
    
    // Jika ditolak, simpan alasan penolakan di Kolom ke-13 (Kolom M)
    if (statusBaru.toString().toLowerCase().includes("tolak") && alasanPenolakan !== "") {
      try {
        if (sheet.getMaxColumns() < 13) {
          sheet.insertColumnAfter(sheet.getMaxColumns());
        }
        sheet.getRange(rowIndex, 13).setValue(alasanPenolakan);
      } catch (errKolom) {
        Logger.log("Abaikan jika gagal buat kolom 13: " + errKolom);
      }
    }

    // Ambil data baris tersebut untuk dikirim via email
    var rowData = sheet.getRange(rowIndex, 1, 1, 12).getDisplayValues()[0];
    
    var instansi  = rowData[1];
    var namaTamu  = rowData[2];
    var emailTamu = rowData[5]; // Kolom F (Email)
    var bertemu   = rowData[6];
    var tanggal   = rowData[7];
    var jam       = rowData[8];

    // Kirim email pemberitahuan ke Tamu jika email valid
    if (emailTamu && emailTamu.toString().includes('@')) {
      kirimEmailKeTamu(emailTamu, namaTamu, instansi, tanggal, jam, bertemu, statusBaru, alasanPenolakan);
    } else {
      Logger.log("Email tidak ditemukan atau salah format: " + emailTamu);
    }

    return { success: true, message: 'Status berhasil diubah menjadi "' + statusBaru + '"' };
  } catch (err) {
    return { success: false, message: 'Gagal mengubah status: ' + err.toString() };
  }
}

// ==========================================
// FUNGSI UNTUK FORM (INDEX)
// ==========================================
function simpanData(payload) {
  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(10000); // Mencegah bentrokan jika banyak yang submit bersamaan
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    // CEK APAKAH SHEET MASIH KOSONG / BELUM ADA HEADER
    if (sheet.getLastRow() === 0) {
      // Menyiapkan nama-nama kolom dalam bahasa Indonesia
      var headers = [
        "Timestamp",
        "Instansi",
        "Nama Pendaftar",
        "Jumlah Tamu",
        "No. HP / WA",
        "Email", 
        "Bertemu Dengan",
        "Tanggal Kunjungan",
        "Jam Kunjungan",
        "Maksud / Keperluan",
        "File Berkas",
        "Status",
        "Alasan Penolakan"
      ];
      sheet.appendRow(headers);
      
      // Membuat gaya (styling) header agar berwarna hijau gelap dan tulisan putih tebal
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold")                 
                 .setFontColor("white")                 
                 .setBackground("#054f39")              
                 .setHorizontalAlignment("center")      
                 .setVerticalAlignment("middle");       
                 
      // Membekukan baris pertama agar header selalu terlihat saat di-scroll
      sheet.setFrozenRows(1);

      // OTOMATIS MENYESUAIKAN LEBAR KOLOM AGAR RAPI
      sheet.autoResizeColumns(1, headers.length);

    }

    // PROSES UPLOAD FILE BERKAS
    var fileUrl = "-"; 
    if (payload.fileBase64 && payload.fileBase64 !== "") {
      try {
        var folderName = "e-Tamu Berkas DPRD";
        var folders = DriveApp.getFoldersByName(folderName);
        var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

        // Hapus prefix/header Base64 jika ada
        var cleanBase64 = payload.fileBase64.replace(/^data:.*;base64,/, "");

        var decodedBytes = Utilities.base64Decode(cleanBase64);
        var blob = Utilities.newBlob(
          decodedBytes, 
          payload.mimeType || 'application/octet-stream', 
          payload.fileName || 'berkas'
        );
        var file = folder.createFile(blob);
        
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
      } catch (errUpload) {
        fileUrl = "Error Upload: " + errUpload.toString();
      }
    }

    // SIMPAN DATA KE GOOGLE SHEETS
    var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    var statusAwal = "Menunggu";

    sheet.appendRow([
      timestamp,                  // Col A (Index 0)
      payload.instansi || '',     // Col B (Index 1)
      payload.nama || '',         // Col C (Index 2)
      payload.jumlahTamu || 1,    // Col D (Index 3)
      payload.noHp || '',         // Col E (Index 4)
      payload.email || '',        // Col F (Index 5)
      payload.bertemu || '',      // Col G (Index 6)
      payload.tanggal || '',      // Col H (Index 7)
      payload.jam || '',          // Col I (Index 8)
      payload.keperluan || '',    // Col J (Index 9)
      fileUrl,                    // Col K (Index 10)
      statusAwal                  // Col L (Index 11)
    ]);

    // KIRIM EMAIL NOTIFIKASI MASUK KE EMAIL DPRD
    kirimEmailKeDPRD({
      timestamp: timestamp,
      instansi: payload.instansi || '',
      nama: payload.nama || '',
      jumlahTamu: payload.jumlahTamu || 1,
      noHp: payload.noHp || '',
      email: payload.email || '',
      bertemu: payload.bertemu || '',
      tanggal: payload.tanggal || '',
      jam: payload.jam || '',
      keperluan: payload.keperluan || '',
      fileUrl: fileUrl
    });

    return { success: true, message: 'Data dan Berkas berhasil disimpan!' };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan: ' + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// FUNGSI PENGIRIMAN EMAIL OTOMATIS
// ==========================================

function kirimEmailKeDPRD(data) {
  try {
    var subject = '[e-Tamu DPRD] Permohonan Kunjungan Baru: ' + data.nama + ' (' + data.instansi + ')';
    var body = 
      'Yth. Sekretariat DPRD Kabupaten Bogor,\n\n' +
      'Telah masuk permohonan kunjungan baru via e-Tamu pada ' + data.timestamp + ':\n\n' +
      '• Nama Tamu       : ' + data.nama + '\n' +
      '• Instansi/Lembaga : ' + data.instansi + '\n' +
      '• Jumlah Tamu     : ' + data.jumlahTamu + ' orang\n' +
      '• No HP/WA        : ' + data.noHp + '\n' +
      '• Email Tamu      : ' + (data.email || '-') + '\n' +
      '• Ingin Bertemu   : ' + data.bertemu + '\n' +
      '• Rencana Jadwal  : ' + data.tanggal + ' (Jam ' + data.jam + ' WIB)\n' +
      '• Keperluan       : ' + data.keperluan + '\n' +
      '• Lampiran Berkas : ' + data.fileUrl + '\n\n' +
      'Silakan buka Dashboard Admin e-Tamu untuk menyetujui atau menolak permohonan ini.';

    GmailApp.sendEmail(EMAIL_DPRD, subject, body);
  } catch (err) {
    Logger.log("Gagal mengirim email ke DPRD: " + err.toString());
  }
}

function kirimEmailKeTamu(email, nama, instansi, tanggal, jam, bertemu, status, alasanPenolakan) {
  try {
    var subject = "";
    var body = "";
    
    var statusCek = (status || "").toString().toLowerCase();

    if (statusCek.includes("setuju")) {
      subject = '[Konfirmasi Tiket] Permohonan Kunjungan DPRD Kabupaten Bogor - DISETUJUI';
      body = 
        'Yth. Bpk/Ibu ' + nama + '\n' +
        '(' + instansi + ')\n\n' +
        'Permohonan kunjungan Anda ke DPRD Kabupaten Bogor telah DISETUJUI.\n\n' +
        'Rincian Jadwal Kunjungan:\n' +
        '• Tanggal   : ' + tanggal + '\n' +
        '• Waktu     : ' + jam + ' WIB\n' +
        '• Ingin Bertemu : ' + bertemu + '\n\n' +
        'Petunjuk Kedatangan:\n' +
        '1. Harap hadir 15 menit sebelum waktu kunjungan.\n' +
        '2. Tunjukkan email konfirmasi ini kepada petugas piket / Pos Satpam gedung DPRD.\n' +
        '3. Mematuhi tata tertib di lingkungan gedung DPRD Kabupaten Bogor.\n\n' +
        'Terima kasih.\n' +
        'Sekretariat DPRD Kabupaten Bogor';

    } else if (statusCek.includes("tolak")) {
      var teksAlasan = alasanPenolakan ? alasanPenolakan : "Mohon maaf, jadwal kunjungan penuh atau belum memenuhi syarat.";
      
      subject = '[Pemberitahuan] Status Permohonan Kunjungan DPRD Kabupaten Bogor';
      body = 
        'Yth. Bpk/Ibu ' + nama + '\n' +
        '(' + instansi + ')\n\n' +
        'Mohon maaf, permohonan kunjungan Anda ke DPRD Kabupaten Bogor saat ini BELUM DAPAT DISETUJUI / DITOLAK.\n\n' +
        'Alasan Penolakan:\n' +
        '"' + teksAlasan + '"\n\n' +
        'Anda dapat mengajukan permohonan kembali untuk jadwal/hari lain melalui sistem e-Tamu.\n\n' +
        'Terima kasih atas perhatian dan pengertian Anda.\n' +
        'Sekretariat DPRD Kabupaten Bogor';
    }

    if (subject !== "") {
      GmailApp.sendEmail(email, subject, body);
    } 
  } catch (err) {
    Logger.log("Gagal mengirim email ke tamu: " + err.toString());
  }
}