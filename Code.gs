// =========================================================================
// PENGATURAN ID (WAJIB DIISI SEBELUM DEPLOY)
// =========================================================================
const SPREADSHEET_ID = 'MASUKKAN_SPREADSHEET_ID_DI_SINI'; 
const FOLDER_ID = 'MASUKKAN_FOLDER_ID_DI_SINI'; 

// Konfigurasi Header Sheet (Disesuaikan dengan urutan kolom)
function setupSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Floor");
  if (!sheet) {
    sheet = ss.insertSheet("Floor");
    sheet.appendRow(["Tanggal", "Nama", "Role", "Shift", "Jam Masuk", "Jam Pulang", "Status", "Link Foto Masuk", "Link Foto Pulang"]);
    sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#f3f4f6");
    sheet.setFrozenRows(1);
  }
}

function doGet(e) {
  setupSheet();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Absensi Floor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Fungsi Upload Foto ke Drive
function uploadPhoto(base64Data, name, type) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    
    // Nama file: Nama_Masuk_1698765432...
    const fileName = `${name}_${type}_${new Date().getTime()}`;
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    
    const file = folder.createFile(blob);
    // Set permission agar foto bisa dilihat publik (oleh yang punya link)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Gagal Upload";
  }
}

// Fungsi Submit Absen
function submitAbsen(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Floor");
    if (!sheet) return { success: false, message: "Tab Floor tidak ditemukan" };

    const now = new Date();
    // Tanggal untuk ditampilkan (DD/MM/YYYY)
    const displayDateStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy");
    // Tanggal untuk pencarian sistem (YYYY-MM-DD)
    const searchDateStr = Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd");
    const timeStr = Utilities.formatDate(now, "GMT+7", "HH:mm");

    let photoUrl = "";
    if (data.photoData) {
      photoUrl = uploadPhoto(data.photoData, data.name, data.type);
    }

    if (data.type === 'masuk') {
      let status = "Hadir";
      const shiftHours = {
        "Morning": "08:00",
        "Middle": "11:00",
        "Evening": "12:00"
      };
      
      const batasMasuk = shiftHours[data.shift];
      if (batasMasuk && timeStr > batasMasuk) {
        status = "Telat";
      }

      // Format row: Tanggal, Nama, Role, Shift, Jam Masuk, Jam Pulang, Status, Foto Masuk, Foto Pulang
      sheet.appendRow([displayDateStr, data.name, data.role, data.shift, timeStr, "-", status, photoUrl, "-"]);
      return { success: true, message: `Berhasil Absen Masuk jam ${timeStr}. Status: ${status}` };
      
    } else if (data.type === 'pulang') {
      const dataValues = sheet.getDataRange().getValues();
      let foundRow = -1;
      
      // Cari dari bawah ke atas agar mendapat absen masuk terakhir di hari yang sama
      for (let i = dataValues.length - 1; i >= 1; i--) {
        const rowDate = dataValues[i][0];
        const rowName = String(dataValues[i][1]).trim();
        let sheetDateStr = "";
        
        if (rowDate instanceof Date) {
          sheetDateStr = Utilities.formatDate(rowDate, "GMT+7", "yyyy-MM-dd");
        } else {
          sheetDateStr = String(rowDate).trim();
          if(sheetDateStr.includes("/")) {
             const parts = sheetDateStr.split("/");
             if(parts.length === 3) {
                 sheetDateStr = `${parts[2]}-${parts[1]}-${parts[0]}`; 
             }
          }
        }

        // Cocokkan nama dan tanggal hari ini
        if (rowName === data.name.trim() && sheetDateStr === searchDateStr) {
          foundRow = i + 1; // Index array ke baris Sheet (+1)
          break;
        }
      }

      if (foundRow !== -1) {
        // Update Kolom F (Jam Pulang) dan Kolom I (Link Foto Pulang)
        sheet.getRange(foundRow, 6).setValue(timeStr); 
        sheet.getRange(foundRow, 9).setValue(photoUrl); 
        return { success: true, message: `Berhasil Absen Pulang jam ${timeStr}.` };
      } else {
        return { success: false, message: "Anda belum Absen Masuk hari ini!" };
      }
    }
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}
