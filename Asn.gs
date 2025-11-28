/**
 * ==========================================
 * MODUL SKP (SASARAN KINERJA PEGAWAI)
 * ==========================================
 */

/**
 * Mengambil Opsi Pegawai untuk Form SKP
 * Sumber: Spreadsheet SIABA_SKP_SOURCE sheet "Daftar ASN"
 * Kolom: A(NIP), B(Nama), C(Unit Kerja)
 */
function getSiabaSkpOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SKP_SOURCE);
    const sheet = ss.getSheetByName("Daftar ASN");
    
    if (!sheet || sheet.getLastRow() < 2) return [];

    // Ambil data A2:C
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();

    // Map ke object: A=NIP, B=Nama, C=Unit
    return data.map(row => ({
      nip: String(row[0]).trim(),  // Kolom A
      nama: String(row[1]).trim(), // Kolom B
      unit: String(row[2]).trim()  // Kolom C
    })).filter(item => item.unit && item.nama);

  } catch (e) {
    return handleError('getSiabaSkpOptions', e);
  }
}

/**
 * Memproses Form Unggah SKP
 * Simpan ke: Spreadsheet SIABA_SKP_DB
 * File ke: Folder SKP -> Folder Tahun
 */
function submitSiabaSkpForm(formData) {
  try {
    // 1. Validasi Folder Utama
    const folderId = FOLDER_CONFIG.SIABA_SKP_DOCS;
    if (!folderId || folderId.includes("GANTI")) {
        throw new Error("ID Folder SKP belum dikonfigurasi.");
    }
    const mainFolder = DriveApp.getFolderById(folderId);

    // 2. Manajemen Folder (Berdasarkan Tahun) - BARU
    // Fungsi getOrCreateFolder sudah ada di kode global Anda
    const targetFolder = getOrCreateFolder(mainFolder, String(formData.tahun));

    // 3. Simpan File
    const fileData = formData.fileData;
    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    
    // Nama File: SKP_Tahun_Nama
    const safeNama = formData.namaAsn.replace(/[^a-zA-Z0-9 ]/g, '');
    const newFileName = `SKP_${formData.tahun}_${safeNama}.pdf`;
    
    // Simpan di folder TAHUN, bukan main folder
    const file = targetFolder.createFile(blob).setName(newFileName);
    const fileUrl = file.getUrl();

    // 4. Simpan Data ke Spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SKP_DB);
    const sheetName = String(formData.tahun);
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["Timestamp", "Unit Kerja", "Nama ASN", "NIP", "Status", "Tahun", "File SKP"]);
    }
    
    const newRow = [
      new Date(),             
      formData.unitKerja,     
      formData.namaAsn,       
      "'" + formData.nip,     
      formData.status,        
      formData.tahun,         
      fileUrl                 
    ];

    sheet.appendRow(newRow);
    
    return "Data SKP berhasil diunggah.";

  } catch (e) {
    return handleError('submitSiabaSkpForm', e);
  }
}

/**
 * ==========================================
 * MODUL KELOLA DATA SKP
 * ==========================================
 */

/**
 * Mengambil Daftar Tahun (Nama Sheet) dari DB SKP
 */
function getSiabaSkpYearList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SKP_DB);
    const sheets = ss.getSheets();
    const years = [];
    
    // Ambil semua nama sheet yang berupa angka (Tahun)
    sheets.forEach(s => {
      const name = s.getName();
      if (/^\d{4}$/.test(name)) { // Regex cek 4 digit angka
        years.push(name);
      }
    });
    
    // Urutkan Descending (Terbaru di atas)
    return years.sort().reverse();
  } catch (e) {
    return handleError('getSiabaSkpYearList', e);
  }
}

/**
 * Mengambil Data SKP berdasarkan Tahun (Nama Sheet)
 */
function getSiabaSkpData(year, unitFilter) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SKP_DB);
    const sheet = ss.getSheetByName(String(year));
    
    if (!sheet) return { units: [], rows: [] }; // Sheet tahun tsb belum ada

    if (sheet.getLastRow() < 2) return { units: [], rows: [] };

    // Ambil Semua Data (A - J) -> 10 Kolom
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getDisplayValues();
    
    const rows = [];
    const unitsSet = new Set();

    // Mapping Index:
    // A(0)=Tgl, B(1)=Unit, C(2)=Nama, D(3)=NIP, E(4)=Status, 
    // F(5)=Tahun, G(6)=File, H(7)=Cek, I(8)=Ket, J(9)=Update
    
    data.forEach((row, index) => {
        const rowUnit = String(row[1] || "").trim(); // Kolom B
        if (rowUnit) unitsSet.add(rowUnit);

        if (unitFilter === "Semua" || rowUnit === unitFilter) {
            // Format Status (Kolom H / Index 7)
            let statusRaw = String(row[7] || "").trim().toUpperCase();
            let statusText = "Diproses";
            if (statusRaw === "V") statusText = "Diterima";
            else if (statusRaw === "X") statusText = "Ditolak";
            else if (statusRaw === "R") statusText = "Revisi";

            rows.push({
                _rowIndex: index + 2, // 1-based row index (header=1)
                _sheetName: String(year), // Penting untuk Edit/Hapus
                tgl: row[0],          // A
                nama: row[2],         // C
                nip: row[3],          // D
                statusPeg: row[4],    // E (Status Pegawai)
                tahun: row[5],        // F
                fileUrl: row[6],      // G
                statusCek: statusText,// H (Converted)
                ket: row[8],          // I
                update: row[9]        // J
            });
        }
    });

    // Sorting: Terbaru (Berdasarkan Tgl Aju / Col A)
    rows.sort((a, b) => {
        return new Date(b.tgl) - new Date(a.tgl);
    });

    return {
        units: Array.from(unitsSet).sort(),
        rows: rows
    };

  } catch (e) {
    return handleError('getSiabaSkpData', e);
  }
}

/**
 * Hapus Data SKP
 */
function deleteSiabaSkpData(sheetName, rowIndex, deleteCode) {
   try {
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SKP_DB);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) throw new Error("Sheet tidak ditemukan.");

    // Hapus File Fisik (Kolom G / Index 7)
    const fileUrl = sheet.getRange(rowIndex, 7).getValue(); 
    if (fileUrl && String(fileUrl).includes("drive.google.com")) {
        try {
            const fileId = String(fileUrl).match(/[-\w]{25,}/);
            if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch(e) {}
    }
    
    sheet.deleteRow(rowIndex);
    return "Data berhasil dihapus.";
  } catch (e) {
    return handleError('deleteSiabaSkpData', e);
  }
}

/**
 * Mengambil Data SKP untuk Edit
 * @param {String} sheetName - Nama sheet (Tahun asal)
 * @param {Number} rowIndex - Indeks baris (1-based)
 */
function getSiabaSkpEditData(sheetName, rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SKP_DB);
    const sheet = ss.getSheetByName(String(sheetName));
    
    if (!sheet) throw new Error("Sheet tahun " + sheetName + " tidak ditemukan.");

    // Ambil data baris (A-G) -> 7 Kolom (Sesuai struktur database SKP)
    // A=Timestamp, B=Unit, C=Nama, D=NIP, E=Status, F=Tahun, G=File
    const rowData = sheet.getRange(rowIndex, 1, 1, 7).getValues()[0];
    
    return {
        sheetName: sheetName, // Tahun Asal
        rowIndex: rowIndex,
        unitKerja: rowData[1],
        namaAsn: rowData[2],
        nip: String(rowData[3]).replace(/'/g, ''),
        status: rowData[4],
        tahun: rowData[5],
        fileUrl: rowData[6]
    };

  } catch (e) {
    return handleError('getSiabaSkpEditData', e);
  }
}

/**
 * Update Data SKP
 * Fitur: Pindah Sheet & Folder jika Tahun Berubah, Update Timestamp, Update Status Revisi
 */
function updateSiabaSkpData(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SKP_DB);
    const oldSheetName = String(formData.oldSheetName); // Tahun Lama
    const newSheetName = String(formData.tahun);        // Tahun Baru
    const rowIndex = parseInt(formData.rowIndex);
    
    const oldSheet = ss.getSheetByName(oldSheetName);
    if (!oldSheet) throw new Error("Sheet lama tidak ditemukan.");

    // 1. Siapkan Data File (Baru atau Lama)
    let fileUrl = formData.existingFileUrl;
    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.SIABA_SKP_DOCS);
    
    // Helper: Pindah/Simpan File ke Folder Tahun Baru
    const handleFileStorage = (targetYear, isNewUpload) => {
        const targetFolder = getOrCreateFolder(mainFolder, String(targetYear));
        
        // Skenario A: Upload File Baru
        if (isNewUpload && formData.fileData && formData.fileData.data) {
             // Hapus file lama jika ada
             if (fileUrl && String(fileUrl).includes("drive.google.com")) {
                 try {
                    const oldId = String(fileUrl).match(/[-\w]{25,}/)[0];
                    DriveApp.getFileById(oldId).setTrashed(true);
                 } catch(e){}
             }

             const decoded = Utilities.base64Decode(formData.fileData.data);
             const blob = Utilities.newBlob(decoded, formData.fileData.mimeType, formData.fileData.fileName);
             const safeNama = String(formData.namaAsn).replace(/[^a-zA-Z0-9 ]/g, '');
             const newName = `SKP_${targetYear}_${safeNama}.pdf`;
             
             const newFile = targetFolder.createFile(blob).setName(newName);
             return newFile.getUrl();
        }
        
        // Skenario B: File Tetap, Tapi Tahun Berubah (Pindah Folder)
        else if (oldSheetName !== newSheetName && fileUrl && String(fileUrl).includes("drive.google.com")) {
             try {
                const oldId = String(fileUrl).match(/[-\w]{25,}/)[0];
                const fileObj = DriveApp.getFileById(oldId);
                
                // Pindah folder (Add to new, Remove from old parents)
                fileObj.moveTo(targetFolder);
                
                // Rename file agar tahunnya sesuai
                const safeNama = String(formData.namaAsn).replace(/[^a-zA-Z0-9 ]/g, '');
                fileObj.setName(`SKP_${targetYear}_${safeNama}.pdf`);
                
                return fileObj.getUrl();
             } catch(e) {
                return fileUrl; // Fallback jika gagal pindah
             }
        }
        
        return fileUrl;
    };

    // Eksekusi Logic File
    const isNewUpload = (formData.fileData && formData.fileData.data);
    fileUrl = handleFileStorage(newSheetName, isNewUpload);

    // 2. Siapkan Data Baris Baru
    // A=Tgl, B=Unit, C=Nama, D=NIP, E=Status, F=Tahun, G=File, H=Cek, I=Ket, J=Update
    const oldData = oldSheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
    
    // PERBAIKAN: Status Cek (Kolom H) dipaksa menjadi "R" (Revisi)
    const statusCek = "R"; 
    
    const ket = oldData[8]; // Keterangan (I) tetap sama atau bisa dikosongkan jika perlu

    const dataRow = [
        oldData[0],           // A: Timestamp Awal (Tetap)
        formData.unitKerja,   // B
        formData.namaAsn,     // C
        "'" + formData.nip,   // D
        formData.status,      // E
        newSheetName,         // F
        fileUrl,              // G
        statusCek,            // H: "R" (REVISI)
        ket,                  // I
        new Date()            // J: UPDATE TIMESTAMP
    ];

    // 3. Simpan Data
    
    // KASUS 1: TAHUN BERUBAH (Pindah Sheet)
    if (oldSheetName !== newSheetName) {
        // Hapus dari sheet lama
        oldSheet.deleteRow(rowIndex);
        
        // Masukkan ke sheet baru
        let newSheet = ss.getSheetByName(newSheetName);
        if (!newSheet) {
            newSheet = ss.insertSheet(newSheetName);
            newSheet.appendRow(["Timestamp", "Unit Kerja", "Nama ASN", "NIP", "Status", "Tahun", "File SKP", "Cek", "Keterangan", "Update"]);
        }
        newSheet.appendRow(dataRow);
    } 
    
    // KASUS 2: TAHUN TETAP (Update di tempat)
    else {
        // Update Range A-J (10 Kolom)
        oldSheet.getRange(rowIndex, 1, 1, 10).setValues([dataRow]);
    }

    return "Data SKP berhasil diperbarui.";

  } catch (e) {
    return handleError('updateSiabaSkpData', e);
  }
}

/**
 * Mengambil Data Daftar Pengiriman SKP (Dinamis)
 * Sumber: Sheet "Daftar Kirim"
 * Kolom: B sampai Terakhir
 * Filter: Unit Kerja (Kolom A)
 */
function getSiabaSkpDaftarKirimData(unitFilter) {
  try {
    // Gunakan ID Spreadsheet SKP yang sudah ada
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SKP_DB);
    const sheet = ss.getSheetByName("Daftar Kirim");
    
    if (!sheet || sheet.getLastRow() < 2) {
        return { units: [], headers: [], rows: [] };
    }

    const range = sheet.getDataRange();
    const values = range.getDisplayValues();
    
    // Baris 1: Header
    // Ambil Header mulai dari Kolom B (Index 1) sampai akhir
    const allHeaders = values[0];
    const displayHeaders = allHeaders.slice(1); // Hapus Kolom A (Unit Kerja)

    // Cari Index Kolom "Nama ASN" untuk sorting (di dalam displayHeaders)
    // Jika header di spreadsheet adalah "Nama ASN", di array slice index-nya bergeser -1 dari aslinya
    // Kita cari text "Nama ASN" atau "Nama"
    let sortIndex = 0; // Default sort kolom pertama (Kolom B)
    for(let i=0; i<displayHeaders.length; i++) {
        if(displayHeaders[i].toLowerCase().includes("nama")) {
            sortIndex = i;
            break;
        }
    }

    const rows = [];
    const unitsSet = new Set();

    // Data mulai baris 2
    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const rowUnit = String(row[0] || "").trim(); // Kolom A
        
        if (rowUnit) unitsSet.add(rowUnit);

        if (unitFilter === "Semua" || rowUnit === unitFilter) {
            // Ambil data mulai Kolom B (Index 1) sampai akhir
            rows.push(row.slice(1));
        }
    }

    // Sorting Abjad berdasarkan Nama ASN (sortIndex)
    rows.sort((a, b) => String(a[sortIndex]).localeCompare(String(b[sortIndex])));

    return {
        units: Array.from(unitsSet).sort(),
        headers: displayHeaders,
        rows: rows
    };

  } catch (e) {
    return handleError('getSiabaSkpDaftarKirimData', e);
  }
}