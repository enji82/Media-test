// --- KONFIGURASI FOLDER ---
// ID Folder Utama (Penyimpanan Aktif)
function getSkArsipFolderIds() {
  try { return { 'MAIN_SK': FOLDER_CONFIG.MAIN_SK }; } catch(e) { return {}; }
}

// ID Folder Trash (Tempat Pembuangan)
const TRASH_FOLDER_ID = '1OB2Mxa_zvpYl7Vru9NEddYmBlU5SfYHL';
const ARSIP_ROOT_ID = '1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs';

// --- HELPER: DAPATKAN/BUAT FOLDER (AKTIF) ---
function getTargetFolder(tahun, semester) {
  const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
  return getOrCreateSubfolder(mainFolder, tahun, semester);
}

// --- HELPER: DAPATKAN/BUAT FOLDER (TRASH) ---
function getTrashTargetFolder(tahun, semester) {
  const trashRoot = DriveApp.getFolderById(TRASH_FOLDER_ID);
  return getOrCreateSubfolder(trashRoot, tahun, semester);
}

// Fungsi bantu umum untuk logika folder bertingkat
function getOrCreateSubfolder(rootFolder, level1Name, level2Name) {
  // 1. Level 1 (Tahun)
  let folder1;
  const iter1 = rootFolder.getFoldersByName(level1Name);
  if (iter1.hasNext()) { folder1 = iter1.next(); } 
  else { folder1 = rootFolder.createFolder(level1Name); }

  // 2. Level 2 (Semester)
  let folder2;
  const iter2 = folder1.getFoldersByName(level2Name);
  if (iter2.hasNext()) { folder2 = iter2.next(); } 
  else { folder2 = folder1.createFolder(level2Name); }

  return folder2;
}

// --- FUNGSI UTAMA: AMBIL DATA TABEL (SORTIR PRIORITAS UPDATE) ---
function getSKTableData(isKelola) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const ss = SpreadsheetApp.openById(config.id);
    const sheet = ss.getSheetByName(config.sheet);
    
    if (!sheet) throw new Error(`Sheet "${config.sheet}" tidak ditemukan.`);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; 

    const rows = data.slice(1);

    // INDEX KOLOM
    const I = {
        TS: 0, STATUS_SEK: 1, NAMA_SEK: 2, TA: 3, SMT: 4, NO_SK: 5, TGL_SK: 6, KRITERIA: 7, DOK: 8,
        STATUS: 9, KET: 10, UPDATE: 11, USER: 12, VERIF: 13
    };

    const safeDate = (v, fmt) => (v instanceof Date && !isNaN(v)) ? Utilities.formatDate(v, Session.getScriptTimeZone(), fmt) : (v || '-');

    // Update Helper di dalam getSKTableData juga
    const formatStatus = (val) => {
        let s = String(val).trim();
        const sl = s.toLowerCase();
        let cssClass = 'status-diproses'; 
        let displayText = s || 'Diproses';    

        if (sl === 'v' || sl === 'disetujui' || sl === 'ok') {
            cssClass = 'status-ok'; displayText = 'OK'; 
        } 
        else if (sl === 'x' || sl === 'ditolak') {
            cssClass = 'status-ditolak'; displayText = 'Ditolak';
        }
        else if (sl === 'revisi') {
            cssClass = 'status-revisi'; displayText = 'Revisi';
        }
        else if (sl === 'dihapus') {
            cssClass = 'status-dihapus'; displayText = 'Dihapus';
        }
        
        return `<span class="status-badge ${cssClass}">${displayText}</span>`;
    };

    const result = rows.map((r, i) => {
        let rowObj = {
            _rowIndex: i + 2,
            _source: 'SK',
            _statusRaw: r[I.STATUS], 
            
            'Nama Sekolah': r[I.NAMA_SEK] || '-',
            'Tahun Ajaran': r[I.TA] || '-',
            'Semester': r[I.SMT] || '-',
            'Nomor SK': r[I.NO_SK] || '-',
            'Tanggal SK': safeDate(r[I.TGL_SK], "dd/MM/yyyy"),
            'Kriteria SK': r[I.KRITERIA] || '-',
            'Status': formatStatus(r[I.STATUS]), 
            'Dokumen': r[I.DOK] || '',
            'Tanggal Unggah': safeDate(r[I.TS], "dd/MM/yyyy HH:mm:ss"),
            'User': r[I.USER] || '-'
        };

        if (isKelola) {
            rowObj['Keterangan'] = r[I.KET] || '-';
            rowObj['Verifikator'] = r[I.VERIF] || '-';
            rowObj['Update'] = safeDate(r[I.UPDATE], "dd/MM/yyyy HH:mm:ss"); 
        }
        
        // --- LOGIKA SORTIR CERDAS ---
        // Ambil waktu Unggah
        let timeUpload = (r[I.TS] instanceof Date) ? r[I.TS].getTime() : 0;
        
        // Ambil waktu Update (jika ada)
        let timeUpdate = (r[I.UPDATE] instanceof Date) ? r[I.UPDATE].getTime() : 0;
        
        // Bandingkan: Ambil mana yang paling besar (paling baru)
        // Jika belum pernah diupdate (timeUpdate 0), maka pakai timeUpload
        rowObj._sortKey = Math.max(timeUpload, timeUpdate);
        
        return rowObj;
    });

    // Sortir Descending (Paling Baru di Atas)
    result.sort((a, b) => b._sortKey - a._sortKey);
    
    return result;

  } catch (e) {
    return { error: e.message };
  }
}

// --- FUNGSI AMBIL DATA BARIS (EDIT) ---
function getSKDataByRow(rowIndex) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const range = sheet.getRange(rowIndex, 2, 1, 8); 
    const val = range.getValues()[0];

    let tglSK = '';
    if (val[5] instanceof Date) {
        tglSK = Utilities.formatDate(val[5], Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    return {
      statusSekolah: String(val[0]), 
      namaSekolah: String(val[1]),   
      tahunAjaran: String(val[2]),   
      semester: String(val[3]),      
      nomorSK: String(val[4]),       
      tanggalSK: tglSK,              
      kriteriaSK: String(val[6]),    
      dokumen: String(val[7])        
    };
  } catch (e) {
    return { error: "Gagal ambil data: " + e.message };
  }
}

// --- FUNGSI UPDATE DATA (EDIT) ---
function updateSKData(form) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const row = parseInt(form.rowIndex);
    
    // 1. Cek Status Kunci
    const statusVal = String(sheet.getRange(row, 10).getValue()).trim().toLowerCase();
    if (['ok', 'v', 'disetujui'].includes(statusVal)) {
        throw new Error("Data sudah disetujui (OK) dan tidak dapat diedit.");
    }

    // 2. Pindah File (Jika ada upload baru)
    const targetFolder = getTargetFolder(form.tahunAjaran, form.semester);
    if (form.fileData && form.fileData.data) {
       const blob = Utilities.newBlob(Utilities.base64Decode(form.fileData.data), form.fileData.mimeType, form.fileData.fileName);
       sheet.getRange(row, 9).setValue(targetFolder.createFile(blob).getUrl());
    } 
    // Logika pindah folder jika tahun/smt berubah (opsional, aman dihapus jika error)
    else {
       const oldUrl = sheet.getRange(row, 9).getValue();
       if(oldUrl && String(oldUrl).includes("drive")) {
           try { DriveApp.getFileById(oldUrl.match(/[-\w]{25,}/)[0]).moveTo(targetFolder); } catch(e){}
       }
    }

    // 3. LOGIKA BARU: KETERANGAN REVISI DINAMIS
    const editorName = form.editorName || "User"; // Nama dari Frontend
    
    // Ambil Keterangan Lama
    let oldKet = String(sheet.getRange(row, 11).getValue() || "-");
    if (oldKet === "-") oldKet = "";

    // Bersihkan tag revisi lama (Regex: hapus text dalam kurung yang diawali 'Direvisi oleh')
    // Agar tidak menumpuk: "Ket (Direvisi oleh A) (Direvisi oleh B)"
    let cleanKet = oldKet.replace(/\s*\(Direvisi oleh.*?\)/gi, "").trim();

    // Buat Tag Baru
    let newKet = cleanKet + " (Direvisi oleh " + editorName + ")";

    // 4. Update Status (Jika status 'Ditolak', kembalikan ke 'Revisi')
    let newStatus = statusVal;
    if (['ditolak', 'x'].includes(statusVal)) newStatus = 'Revisi';
    else if (['diproses', '', '-'].includes(statusVal)) newStatus = 'Diproses';

    // 5. SIMPAN KE SHEET
    // Data Utama (Kolom B s.d H)
    sheet.getRange(row, 2, 1, 7).setValues([[
        form.statusSekolah, form.namaSekolah, form.tahunAjaran, 
        form.semester, form.nomorSK, form.tanggalSK, form.kriteriaSK
    ]]);

    // Data Admin (Kolom J, K, L, M)
    sheet.getRange(row, 10).setValue(newStatus); 
    sheet.getRange(row, 11).setValue(newKet); // <-- Keterangan Baru
    sheet.getRange(row, 12).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"));
    sheet.getRange(row, 13).setValue(editorName); // User Update

    return "Data berhasil diperbarui.";
    
  } catch (e) {
    throw new Error(e.message);
  }
}

// --- FUNGSI HAPUS (PINDAH KE TRASH DENGAN NAMA & ROLE) ---
function deleteSKData(idx, code, actorName, actorRole) {
    try {
        const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
        if(String(code).trim() !== today) throw new Error("Kode hapus salah.");
        
        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const ss = SpreadsheetApp.openById(config.id);
        const sheet = ss.getSheetByName(config.sheet);
        
        const range = sheet.getRange(idx, 1, 1, sheet.getLastColumn());
        const values = range.getValues()[0];
        const tahun = values[3]; const smt = values[4]; const fileUrl = values[8];

        // --- UPDATE DATA UNTUK TRASH ---
        const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        
        values[9] = "Dihapus"; 
        
        // KETERANGAN: Diisi Nama Orang ("Dihapus oleh Budi")
        values[10] = "Dihapus oleh " + (actorName || "User"); 
        
        values[11] = timeStamp; 
        
        // USER: Diisi Role ("Administrator")
        values[12] = actorRole || "User"; 

        // Pindah ke Trash
        let trashSheet = ss.getSheetByName('Trash');
        if (!trashSheet) {
            trashSheet = ss.insertSheet('Trash');
            const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            trashSheet.appendRow(headers);
        }
        trashSheet.appendRow(values); 

        // Pindah File
        if (fileUrl && String(fileUrl).includes("drive")) {
            try { DriveApp.getFileById(fileUrl.match(/[-\w]{25,}/)[0]).moveTo(getTrashTargetFolder(tahun, smt)); } catch(e){}
        }

        sheet.deleteRow(idx);
        return "Data berhasil dihapus.";
    } catch(e) { throw new Error(e.message); }
}

// --- FUNGSI VERIFIKASI (ADMIN) ---
function updateSKVerificationStatus(idx, statusInput, ket, adminName) {
    try {
        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
        
        // Status Final
        let finalStatus = statusInput; // OK, Ditolak, Revisi, Diproses

        // Timestamp
        const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        
        // Update Kolom J (10) s.d N (14)
        // J=Status, K=Ket, L=Update, M=User(skip), N=Verifikator
        sheet.getRange(idx, 10).setValue(finalStatus);
        sheet.getRange(idx, 11).setValue(ket || "-");
        sheet.getRange(idx, 12).setValue(timeStamp);
        // Kolom 13 (User) jangan diubah karena itu user pengupload/pengedit
        sheet.getRange(idx, 14).setValue(adminName);
        
        return "Verifikasi berhasil disimpan.";
        
    } catch(e) {
        throw new Error(e.message);
    }
}

// --- FUNGSI UNGGAH BARU ---
function processManualForm(form) {
    try {
        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
        const targetFolder = getTargetFolder(form.tahunAjaran, form.semester);
        
        const blob = Utilities.newBlob(Utilities.base64Decode(form.fileData.data), form.fileData.mimeType, form.fileData.fileName);
        const file = targetFolder.createFile(blob);
        
        const row = [
            new Date(),
            form.statusSekolah, form.namaSekolah, form.tahunAjaran, form.semester,
            form.nomorSK, new Date(form.tanggalSK), form.kriteriaSK,
            file.getUrl(),
            'Diproses', '-', '', form.loggedInUser || 'User', '-'
        ];
        
        sheet.appendRow(row);
        return "Berhasil disimpan.";
    } catch(e) { throw new Error(e.message); }
}

// --- FUNGSI AMBIL DATA DROPDOWN (WAJIB ADA) ---
function getMasterSkOptions() {
  try {
    // 1. Coba ambil dari Sheet 'DropdownData' (Prioritas Utama)
    const config = SPREADSHEET_CONFIG.DROPDOWN_DATA; // Ambil dari Code.gs
    const ss = SpreadsheetApp.openById(config.id);
    const sheet = ss.getSheetByName(config.sheet);
    
    // Objek penampung hasil
    let options = {
        'Tahun Ajaran': [],
        'Semester': [],
        'Kriteria SK': [],
        'Nama Sekolah List': {} // Format: {'Negeri': ['SD A', 'SD B'], 'Swasta': ['SD C']}
    };

    if (sheet) {
        // Asumsi di Sheet DropdownData:
        // Kolom A: Tahun, B: Semester, C: Kriteria, D: Status Sek, E: Nama Sek
        const data = sheet.getDataRange().getValues();
        // Skip header (baris 1)
        for (let i = 1; i < data.length; i++) {
            if(data[i][0]) options['Tahun Ajaran'].push(data[i][0]);
            if(data[i][1]) options['Semester'].push(data[i][1]);
            if(data[i][2]) options['Kriteria SK'].push(data[i][2]);
            
            // Mapping Nama Sekolah berdasarkan Status (Kolom D & E)
            let status = data[i][3]; // Misal: Negeri/Swasta
            let nama = data[i][4];   // Misal: SDN 1 Secang
            if (status && nama) {
                if (!options['Nama Sekolah List'][status]) {
                    options['Nama Sekolah List'][status] = [];
                }
                options['Nama Sekolah List'][status].push(nama);
            }
        }
    } 
    
    // 2. FALLBACK (Jaga-jaga jika sheet DropdownData kosong/hilang)
    // Kita ambil data unik dari Sheet 'Respon SK' agar dropdown tidak kosong melompong
    else {
        const dbSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.id)
                                      .getSheetByName(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.sheet);
        if(dbSheet) {
            const data = dbSheet.getDataRange().getValues();
            // Ambil unik dari kolom yang sudah ada (Hanya darurat)
            let tahunSet = new Set(), smtSet = new Set(), namaSet = new Set();
            for(let i=1; i<data.length; i++) {
                if(data[i][3]) tahunSet.add(data[i][3]); // Kolom D (Tahun)
                if(data[i][4]) smtSet.add(data[i][4]);   // Kolom E (Semester)
                if(data[i][2]) namaSet.add(data[i][2]);  // Kolom C (Nama)
            }
            options['Tahun Ajaran'] = Array.from(tahunSet);
            options['Semester'] = Array.from(smtSet);
            // Struktur Sekolah Sederhana
            options['Nama Sekolah List'] = { 'Semua': Array.from(namaSet) };
        }
    }

    // Bersihkan duplikat & Sortir
    options['Tahun Ajaran'] = [...new Set(options['Tahun Ajaran'])].sort().reverse();
    options['Semester'] = [...new Set(options['Semester'])];
    options['Kriteria SK'] = [...new Set(options['Kriteria SK'])];

    return options;

  } catch (e) {
    // Kembalikan error agar frontend tahu
    return { error: e.message };
  }
}

// --- FUNGSI PROSES UPLOAD (WAJIB ADA JUGA) ---
function processManualForm(form) {
    try {
        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
        
        // 1. Upload File ke Drive
        // Fungsi getTargetFolder ada di script sebelumnya, pastikan ada!
        // Jika hilang, gunakan folder root sementara
        let folderId = FOLDER_CONFIG.MAIN_SK; 
        try { folderId = getTargetFolder(form.tahunAjaran, form.semester).getId(); } catch(e){}
        
        const folder = DriveApp.getFolderById(folderId);
        const blob = Utilities.newBlob(Utilities.base64Decode(form.fileData.data), form.fileData.mimeType, form.fileData.fileName);
        const file = folder.createFile(blob);
        
        // 2. Simpan ke Spreadsheet
        // Urutan: [Timestamp, StatusSek, NamaSek, Tahun, Smt, NoSK, TglSK, Kriteria, LinkFile, Status, Ket, Update, User, Verif]
        const row = [
            new Date(),
            form.statusSekolah, 
            form.namaSekolah, 
            form.tahunAjaran, 
            form.semester, 
            form.nomorSK, 
            new Date(form.tanggalSK), 
            form.kriteriaSK,
            file.getUrl(),
            'Diproses', '-', '', form.loggedInUser || 'User', '-'
        ];
        
        sheet.appendRow(row);
        return "Berhasil disimpan. Data sedang diproses.";
    } catch(e) { throw new Error(e.message); }
}

// --- FUNGSI AMBIL DATA STATUS (DAFTAR KIRIM - MATRIKS) ---
function getSKStatusData() {
  try {
    // ID dan Sheet
    const idSpreadsheet = '1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4'; 
    const ss = SpreadsheetApp.openById(idSpreadsheet);
    const sheet = ss.getSheetByName('Daftar Kirim SK');
    
    if (!sheet) throw new Error("Sheet 'Daftar Kirim SK' tidak ditemukan.");

    const data = sheet.getDataRange().getValues();
    if (data.length < 3) return { error: "Data kosong." };

    // Mapping Data Matriks
    const formattedData = data.map((row, rowIndex) => {
        if (rowIndex < 2) return row; // Header biarkan saja

        return row.map((cell, colIndex) => {
            if (colIndex === 0) return cell; // Kolom Nama Sekolah biarkan

            // Format Status -> Ke Class CSS Baru
            let s = String(cell).trim();
            const sl = s.toLowerCase();
            let css = 'status-diproses'; // Default
            let label = s;

            if (['v', 'ok', 'disetujui', 'sudah'].includes(sl)) {
                css = 'status-ok'; label = 'OK';
            }
            else if (['x', 'ditolak'].includes(sl)) {
                css = 'status-ditolak'; label = 'Ditolak';
            }
            else if (['revisi'].includes(sl)) {
                css = 'status-revisi'; label = 'Revisi';
            }
            else if (['belum', '-', ''].includes(sl)) {
                css = 'status-belum'; label = 'Belum';
            }
            
            return `<span class="status-badge ${css}">${label}</span>`;
        });
    });

    return formattedData;

  } catch (e) {
    return { error: e.message };
  }
}

function getSKRiwayatData() { return getSKTableData(false); }
function getSKKelolaData() { return getSKTableData(true); }

function getDriveContents(folderId) {
  try {
    const id = folderId || ARSIP_ROOT_ID;
    const folder = DriveApp.getFolderById(id);
    
    const folders = folder.getFolders();
    const files = folder.getFiles();
    
    let items = [];

    // 1. Ambil Folder
    while (folders.hasNext()) {
      const f = folders.next();
      items.push({
        id: f.getId(),
        name: f.getName(),
        type: 'folder',
        mime: 'application/vnd.google-apps.folder',
        updated: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy")
      });
    }

    // 2. Ambil File (Khusus PDF atau semua, kita ambil PDF utamanya)
    while (files.hasNext()) {
      const f = files.next();
      // Filter opsional: if (f.getMimeType() === MimeType.PDF)
      items.push({
        id: f.getId(),
        name: f.getName(),
        type: 'file',
        mime: f.getMimeType(),
        url: f.getUrl(),
        size: (f.getSize() / 1024).toFixed(0) + ' KB',
        updated: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy")
      });
    }

    // Sortir: Folder dulu, baru File (A-Z)
    items.sort((a, b) => {
        if (a.type === b.type) return a.name.localeCompare(b.name);
        return a.type === 'folder' ? -1 : 1;
    });

    return { 
        currentId: id, 
        currentName: folder.getName(), 
        items: items 
    };

  } catch (e) {
    return { error: "Gagal memuat arsip: " + e.message };
  }
}

// --- FUNGSI AMBIL DATA TRASH (FIX NULL ERROR) ---
function getSKTrashData() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const ss = SpreadsheetApp.openById(config.id);
    let sheet = ss.getSheetByName('Trash');
    
    // SKENARIO 1: Jika sheet Trash belum ada (Baru pertama kali atau habis dihapus)
    if (!sheet) {
        // Buat Sheet Baru
        sheet = ss.insertSheet('Trash');
        // Buat Header (Copy dari Sheet Utama agar sama persis)
        const mainSheet = ss.getSheetByName(config.sheet);
        const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
        sheet.appendRow(headers);
        
        // PENTING: Kembalikan Array Kosong agar Frontend tidak CRASH
        return []; 
    }

    const data = sheet.getDataRange().getValues();
    
    // SKENARIO 2: Sheet ada tapi isinya cuma header (Belum ada sampah)
    if (data.length < 2) {
        return []; // Kembalikan Array Kosong
    }

    const rows = data.slice(1); // Skip Header

    // Format Tanggal
    const safeDate = (v) => (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy") : (v || '-');
    const safeTime = (v) => (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") : (v || '-');

    // MAPPING DATA TRASH (SAMA PERSIS DENGAN TABEL UTAMA)
    // Karena kita memindahkan data 1:1, maka urutan kolomnya SAMA dengan Sheet Utama
    // A=Timestamp, B=StatusSek, C=Nama, D=Tahun, E=Smt, F=No, G=Tgl, H=Kriteria, I=Dok, J=Status, K=Ket, L=Update, M=User, N=Verif
    
    const result = rows.map((r) => {
        return {
            'Nama Sekolah': r[2] || '-', 
            'Tahun Ajaran': r[3] || '-', 
            'Semester': r[4] || '-',     
            'Nomor SK': r[5] || '-',     
            'Tanggal SK': safeDate(r[6]),
            'Kriteria SK': r[7] || '-',  
            'Dokumen': r[8] || '',       
            // Paksa Status jadi Merah
            'Status': `<span class="status-badge status-rejected">${r[9] || 'Dihapus'}</span>`,
            'Keterangan': r[10] || '-',
            'Update': safeTime(r[11]),   // Waktu Hapus
            'User': r[12] || '-',        // Siapa yang menghapus
            'Verifikator': r[13] || '-'
        };
    });

    // Urutkan dari yang paling baru dihapus (Paling bawah di sheet = Paling baru)
    return result.reverse();

  } catch (e) {
    // Kembalikan Object Error
    return { error: "Gagal memuat Trash: " + e.message };
  }
}

// ==========================================
// FUNGSI PEMANGGIL HALAMAN (WAJIB ADA)
// ==========================================
function getPageContent(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    return '<div style="color:red; padding:20px; text-align:center;">Error: Halaman "' + filename + '" tidak ditemukan.<br><small>' + e.message + '</small></div>';
  }
}

// --- FUNGSI WAJIB UNTUK LOAD HALAMAN ---
function srvLoadPage(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    return '<div style="color:red; text-align:center; padding:20px;">Error: Halaman <b>' + filename + '</b> tidak ditemukan.<br>Detail: ' + e.message + '</div>';
  }
}

// ==========================================
// FUNGSI PENGHUBUNG HALAMAN (WAJIB ADA)
// ==========================================
// Tanpa fungsi ini, layar akan blank saat ganti menu
function include(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    return '<div style="color:red; text-align:center; padding:20px;">' +
           '<h3>Gagal Memuat Halaman</h3>' +
           '<p>File <b>' + filename + '</b> tidak ditemukan.</p>' +
           '<small>Error: ' + e.message + '</small>' +
           '</div>';
  }
}