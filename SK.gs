/**
 * ===================================================================
 * MODUL SK PEMBAGIAN TUGAS (RENAME FILE & FOLDER OTOMATIS)
 * ===================================================================
 */

// --- 1. KONFIGURASI FOLDER ---
function getSkArsipFolderIds() {
  try { return { 'MAIN_SK': FOLDER_CONFIG.MAIN_SK }; } catch(e) { return {}; }
}

const TRASH_FOLDER_ID = '1OB2Mxa_zvpYl7Vru9NEddYmBlU5SfYHL';
const ARSIP_ROOT_ID = '1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs';

// Fungsi Helper: Cari/Buat Folder Bertingkat
function getTargetFolder(tahun, semester) {
  const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
  return getOrCreateSubfolder(mainFolder, tahun, semester);
}

// --- 2. FUNGSI DROPDOWN ---
function getMasterSkOptions() {
  try {
    let options = {
        'Tahun Ajaran': [], 'Semester': [], 'Kriteria SK': [], 'Nama Sekolah List': {} 
    };

    try {
        if (typeof SPREADSHEET_CONFIG !== 'undefined') {
            const config = SPREADSHEET_CONFIG.DROPDOWN_DATA;
            const ss = SpreadsheetApp.openById(config.id);
            const sheet = ss.getSheetByName(config.sheet);
            
            if (sheet) {
                const data = sheet.getDataRange().getValues();
                for (let i = 1; i < data.length; i++) {
                    if(data[i][0]) options['Tahun Ajaran'].push(String(data[i][0]));
                    if(data[i][1]) options['Semester'].push(String(data[i][1]));
                    
                    let status = String(data[i][2]).trim(); 
                    let nama = String(data[i][3]).trim();   
                    
                    if(data[i][4]) options['Kriteria SK'].push(String(data[i][4]));
                    
                    if (status && nama) {
                        if (!options['Nama Sekolah List'][status]) options['Nama Sekolah List'][status] = [];
                        options['Nama Sekolah List'][status].push(nama);
                    }
                }
            }
        }
    } catch(err) { console.log("Spreadsheet error: " + err.message); }

    if (options['Tahun Ajaran'].length === 0) options['Tahun Ajaran'] = ["2024/2025", "2025/2026"];
    if (options['Semester'].length === 0) options['Semester'] = ["Ganjil", "Genap"];
    if (options['Kriteria SK'].length === 0) options['Kriteria SK'] = ["PNS", "PPPK", "GTT"];
    if (Object.keys(options['Nama Sekolah List']).length === 0) {
        options['Nama Sekolah List'] = { 'Data Kosong': ['Cek Spreadsheet'] };
    }

    options['Tahun Ajaran'] = [...new Set(options['Tahun Ajaran'])].sort().reverse();
    options['Semester'] = [...new Set(options['Semester'])];
    options['Kriteria SK'] = [...new Set(options['Kriteria SK'])];
    
    return options;

  } catch (e) { return { error: "Error Backend: " + e.message }; }
}

// --- 3. FUNGSI UNGGAH (RENAME & SAVE FOLDER) ---
function processManualForm(form) {
    try {
        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
        
        // A. SIAPKAN FOLDER TUJUAN (Tahun -> Semester)
        // Jika folder belum ada, fungsi ini akan otomatis membuatnya
        let targetFolder;
        try { 
            targetFolder = getTargetFolder(form.tahunAjaran, form.semester); 
        } catch(e) {
            // Fallback ke folder utama jika gagal
            targetFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
        }
        
        // B. RENAME FILE OTOMATIS
        // 1. Bersihkan Tahun Ajaran (Ganti / jadi _)
        let tahunClean = String(form.tahunAjaran).replace(/\//g, '_');
        
        // 2. Ambil Ekstensi File Asli (misal .pdf)
        let originalName = form.fileData.fileName;
        let extension = "";
        if (originalName.indexOf('.') !== -1) {
            extension = originalName.substring(originalName.lastIndexOf('.'));
        }
        
        // 3. Susun Nama Baru: "Nama Sekolah - Tahun_Ajaran - Semester - Kriteria.pdf"
        let newFileName = `${form.namaSekolah} - ${tahunClean} - ${form.semester} - ${form.kriteriaSK}${extension}`;

        // C. BUAT FILE DI DRIVE
        const blob = Utilities.newBlob(Utilities.base64Decode(form.fileData.data), form.fileData.mimeType, newFileName);
        const file = targetFolder.createFile(blob);
        
        // D. SIMPAN DATA KE SHEET
        let tglSK = new Date(form.tanggalSK);
        tglSK.setHours(0, 0, 0, 0);

        const row = [
            new Date(), // Timestamp
            form.statusSekolah, 
            form.namaSekolah, 
            form.tahunAjaran, 
            form.semester,
            form.nomorSK, 
            tglSK, // Date Object
            form.kriteriaSK, 
            file.getUrl(), // Link File
            'Diproses', '-', '', form.loggedInUser || 'User', '-'
        ];
        
        sheet.appendRow(row);
        return "Berhasil disimpan.";
    } catch(e) { throw new Error(e.message); }
}

// --- 4. FUNGSI TABEL (READ) ---
function getSKTableData(isKelola) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if(!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; 
    
    const rows = data.slice(1);
    
    const safeDate = (v, f) => {
        if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), f);
        if (typeof v === 'string' && v.includes('/')) return v.replace("'", "").trim();
        return (v || '-');
    };
    
    const result = rows.map((r, i) => {
        let statusHtml = `<span class="status-badge status-diproses">${r[9]||'Diproses'}</span>`;
        let s = String(r[9]).toLowerCase();
        if(s==='ok'||s==='v'||s==='disetujui') statusHtml = `<span class="status-badge status-ok">OK</span>`;
        else if(s==='ditolak'||s==='x') statusHtml = `<span class="status-badge status-ditolak">Ditolak</span>`;
        else if(s==='revisi') statusHtml = `<span class="status-badge status-revisi">Revisi</span>`;

        let obj = {
            _rowIndex: i + 2,
            'Nama Sekolah': r[2] || '-', 
            'Tahun Ajaran': r[3] || '-', 
            'Semester': r[4] || '-',
            'Nomor SK': r[5] || '-', 
            'Tanggal SK': safeDate(r[6], "dd/MM/yyyy"), 
            'Kriteria SK': r[7] || '-', 
            'Status': statusHtml, 
            '_statusRaw': r[9] || 'Diproses',
            'Dokumen': r[8] || '',
            'Tanggal Unggah': safeDate(r[0], "dd/MM/yyyy HH:mm:ss"), 
            'User': r[12] || '-'
        };
        if(isKelola) {
            obj['Keterangan'] = r[10] || '-';
            obj['Verifikator'] = r[13] || '-';
            obj['Update'] = safeDate(r[11], "dd/MM/yyyy HH:mm:ss");
        }
        obj._sortKey = Math.max((r[0] instanceof Date?r[0].getTime():0), (r[11] instanceof Date?r[11].getTime():0));
        return obj;
    });
    return result.sort((a,b) => b._sortKey - a._sortKey);
  } catch(e) { return { error: e.message }; }
}

function getSKRiwayatData() { return getSKTableData(false); }
function getSKKelolaData() { return getSKTableData(true); }

// --- 5. FUNGSI EDIT & UPDATE ---
function getSKDataByRow(rowIndex) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const val = sheet.getRange(rowIndex, 2, 1, 8).getValues()[0];

    let rawTgl = val[5];
    let tglForInput = '';

    if (rawTgl instanceof Date) {
        tglForInput = Utilities.formatDate(rawTgl, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof rawTgl === 'string') {
        let clean = rawTgl.replace("'", "").trim();
        let p = clean.split("/");
        if (p.length === 3) tglForInput = `${p[2]}-${p[1]}-${p[0]}`;
    }

    return {
      statusSekolah: String(val[0]), namaSekolah: String(val[1]), tahunAjaran: String(val[2]),
      semester: String(val[3]), nomorSK: String(val[4]), 
      tanggalSK: tglForInput, 
      kriteriaSK: String(val[6]), dokumen: String(val[7])
    };
  } catch (e) { return { error: e.message }; }
}

function updateSKData(form) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const row = parseInt(form.rowIndex);
    
    // 1. Cek Status Kunci
    const statusVal = String(sheet.getRange(row, 10).getValue()).trim().toLowerCase();
    if (['ok', 'v', 'disetujui'].includes(statusVal)) {
        throw new Error("Data sudah disetujui, tidak bisa diedit.");
    }

    // 2. Siapkan Data Folder & Nama Baru
    const targetFolder = getTargetFolder(form.tahunAjaran, form.semester);
    
    let tahunClean = String(form.tahunAjaran).replace(/\//g, '_');
    // Format Nama Standar: Nama Sekolah - Tahun - Semester - Kriteria
    let baseName = `${form.namaSekolah} - ${tahunClean} - ${form.semester} - ${form.kriteriaSK}`;

    // 3. LOGIKA FILE (UPLOAD BARU vs PINDAHKAN LAMA)
    if (form.fileData && form.fileData.data) {
       // --- KASUS A: ADA UPLOAD FILE BARU ---
       
       // Deteksi ekstensi file baru
       let originalName = form.fileData.fileName;
       let extension = originalName.includes('.') ? originalName.substring(originalName.lastIndexOf('.')) : '.pdf';
       let newFileName = baseName + extension;

       // Upload ke folder target
       const blob = Utilities.newBlob(Utilities.base64Decode(form.fileData.data), form.fileData.mimeType, newFileName);
       const newFile = targetFolder.createFile(blob);
       
       // Update Link di Spreadsheet
       sheet.getRange(row, 9).setValue(newFile.getUrl());
       
       // (Opsional: File lama di spreadsheet dibiarkan menjadi sampah atau bisa dihapus manual)

    } else {
       // --- KASUS B: TIDAK ADA FILE BARU (HANYA EDIT DATA) ---
       // Kita harus memindahkan & rename file lama agar tetap rapi
       
       const currentUrl = String(sheet.getRange(row, 9).getValue());
       
       // Pastikan itu link Google Drive
       if (currentUrl.includes("drive.google.com") || currentUrl.includes("open?id=")) {
           try {
               // Ekstrak ID File Lama
               let fileId = "";
               let match = currentUrl.match(/[-\w]{25,}/);
               if (match) fileId = match[0];

               if (fileId) {
                   var oldFile = DriveApp.getFileById(fileId);
                   
                   // 1. PINDAHKAN FILE (MoveTo)
                   oldFile.moveTo(targetFolder);
                   
                   // 2. RENAME FILE (Agar sesuai data baru)
                   // Ambil ekstensi lama
                   let oldName = oldFile.getName();
                   let ext = oldName.includes('.') ? oldName.substring(oldName.lastIndexOf('.')) : '.pdf';
                   
                   oldFile.setName(baseName + ext);
               }
           } catch(err) {
               // Jika gagal memindahkan (misal file sudah dihapus manual), biarkan saja.
               console.log("Gagal memindahkan file: " + err.message);
           }
       }
    }

    // 4. UPDATE DATA SPREADSHEET
    let newStatus = statusVal; 
    if (['ditolak', 'x'].includes(statusVal)) newStatus = 'Revisi';
    else if (['', '-', 'belum'].includes(statusVal)) newStatus = 'Diproses';

    let tglSimpan = new Date(form.tanggalSK);
    tglSimpan.setHours(0, 0, 0, 0);

    // Update Kolom Data
    sheet.getRange(row, 2, 1, 7).setValues([[
        form.statusSekolah, form.namaSekolah, form.tahunAjaran, 
        form.semester, form.nomorSK, tglSimpan, form.kriteriaSK
    ]]);
    
    // Update Metadata
    sheet.getRange(row, 10).setValue(newStatus);
    
    let currentKet = String(sheet.getRange(row, 11).getValue());
    let revisiNote = `(Direvisi oleh ${form.editorName || "User"} pada ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy HH:mm")})`;
    
    if (currentKet.includes("(Direvisi oleh")) {
        currentKet = currentKet.replace(/\(Direvisi oleh.*?\)/, revisiNote);
    } else {
        currentKet = currentKet ? (currentKet + " " + revisiNote) : revisiNote;
    }
    
    sheet.getRange(row, 11).setValue(currentKet);
    sheet.getRange(row, 12).setValue(new Date());
    sheet.getRange(row, 13).setValue(form.editorName || "User");

    return "Data berhasil diperbarui.";
  } catch (e) { throw new Error(e.message); }
}

function updateSKVerificationStatus(data) {
    try {
        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
        const row = parseInt(data.rowIndex);
        
        // 1. Update Status (Kolom J / Index 10)
        sheet.getRange(row, 10).setValue(data.status);
        
        // 2. Update Keterangan (Kolom K / Index 11)
        // Kita timpa keterangan lama dengan yang baru dari admin
        sheet.getRange(row, 11).setValue(data.keterangan);
        
        // 3. Update Timestamp (Kolom L / Index 12)
        sheet.getRange(row, 12).setValue(new Date());
        
        // 4. Update Verifikator (Kolom N / Index 14)
        // Note: Kolom M (13) adalah User yg upload/edit, Kolom N (14) adalah Verifikator
        sheet.getRange(row, 14).setValue(data.verifikator);
        
        return "Status verifikasi berhasil diperbarui.";
        
    } catch (e) {
        throw new Error("Gagal verifikasi: " + e.message);
    }
}

function getSKStatusData() {
    try {
        const d = SpreadsheetApp.openById('1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4').getSheetByName('Daftar Kirim SK').getDataRange().getValues();
        if(d.length<3) return [];
        return d.map((r,i) => {
            if(i<2) return r;
            return r.map((c,j) => {
                if(j===0) return c;
                let s=String(c).toLowerCase().trim();
                let css='status-diproses'; let t=c;
                if(['v','ok'].includes(s)) { css='status-ok'; t='OK'; }
                else if(['x','ditolak'].includes(s)) { css='status-ditolak'; t='Ditolak'; }
                else if(['revisi'].includes(s)) { css='status-revisi'; t='Revisi'; }
                else if(['belum','','-'].includes(s)) { css='status-belum'; t='Belum'; }
                return `<span class="status-badge ${css}">${t}</span>`;
            });
        });
    } catch(e) { return {error:e.message}; }
}

// --- FUNGSI UTAMA HAPUS DATA ---
function deleteSKData(idx, code, actorName, actorRole) {
    try {
        // 1. Validasi Kode Keamanan (Format: YYYYMMDD hari ini)
        // Contoh: 20251211
        let todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
        
        if(String(code).trim() !== todayCode) {
            throw new Error("Kode hapus salah. Gunakan tanggal hari ini (Format: YYYYMMDD).");
        }

        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const ss = SpreadsheetApp.openById(config.id);
        const sheet = ss.getSheetByName(config.sheet);
        
        // Ambil seluruh data baris tersebut
        const vals = sheet.getRange(idx, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        // 2. PINDAHKAN FILE KE GOOGLE DRIVE TRASH (Sesuai Tahun/Semester)
        // Cek Kolom I (Index 8) apakah ada link file
        if(vals[8] && String(vals[8]).includes("drive")) { 
            try {
                // Ekstrak ID File
                let fileId = "";
                let match = vals[8].match(/[-\w]{25,}/);
                if (match) fileId = match[0];

                if (fileId) {
                    // Ambil Tahun (Index 3) dan Semester (Index 4) dari data
                    let tahun = vals[3];
                    let semester = vals[4];
                    
                    // Dapatkan/Buat folder tujuan di dalam Trash
                    let trashTarget = getTrashTargetFolder(tahun, semester);
                    
                    // Pindahkan file (MoveTo otomatis menghapus dari folder lama)
                    DriveApp.getFileById(fileId).moveTo(trashTarget); 
                }
            } catch(e) {
                console.log("Gagal memindahkan file ke Trash: " + e.message);
                // Kita biarkan error file ini (jangan throw), agar data di sheet tetap terhapus
            } 
        }
        
        // 3. PINDAHKAN DATA KE SHEET TRASH
        // Update metadata sebelum dipindah
        vals[9] = "Dihapus"; // Status
        vals[10] = "Dihapus oleh " + actorName; // Keterangan
        vals[11] = new Date(); // Tanggal Update
        vals[12] = actorRole; // User Executor
        
        let trashSheet = ss.getSheetByName('Trash');
        if(!trashSheet) { 
            trashSheet = ss.insertSheet('Trash'); 
            // Copy Header dari sheet utama jika sheet Trash baru dibuat
            let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            trashSheet.appendRow(headers);
        }
        
        trashSheet.appendRow(vals);
        
        // 4. HAPUS BARIS DARI SHEET UTAMA
        sheet.deleteRow(idx);
        
        return "Data berhasil dihapus dan dipindahkan ke Trash.";
        
    } catch(e) { 
        throw new Error(e.message); 
    }
}

// --- HELPER: CARI FOLDER DI DALAM TRASH ---
function getTrashTargetFolder(tahun, semester) {
  const trashRoot = DriveApp.getFolderById(TRASH_FOLDER_ID);
  // Menggunakan fungsi getOrCreateSubfolder yang sudah ada di SK.gs
  return getOrCreateSubfolder(trashRoot, tahun, semester);
}

// (Pastikan fungsi ini sudah ada di SK.gs Anda. Jika belum, copy juga yang ini)
function getOrCreateSubfolder(rootFolder, level1Name, level2Name) {
  // 1. Cek Folder Tahun (Level 1)
  let folder1;
  const iter1 = rootFolder.getFoldersByName(level1Name);
  if (iter1.hasNext()) { 
      folder1 = iter1.next(); 
  } else { 
      folder1 = rootFolder.createFolder(level1Name); 
  }

  // 2. Cek Folder Semester (Level 2)
  let folder2;
  const iter2 = folder1.getFoldersByName(level2Name);
  if (iter2.hasNext()) { 
      folder2 = iter2.next(); 
  } else { 
      folder2 = folder1.createFolder(level2Name); 
  }

  return folder2;
}

function getSKTrashData() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const ss = SpreadsheetApp.openById(config.id);
    const sheet = ss.getSheetByName('Trash');
    
    // Jika sheet Trash belum ada/kosong
    if(!sheet || sheet.getLastRow() < 2) return [];
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1); // Lewati header
    
    const safeDate = (v, f) => {
        if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), f);
        return (v || '-');
    };
    
    // Mapping Data Trash
    // Ingat struktur deleteSKData: 
    // Col 2:Nama, 3:Tahun, 4:Smt, 5:NoSK, 8:Link, 10:Ket(Siapa yg hapus), 11:Waktu Hapus
    const result = rows.map((r, i) => {
        return {
            _rowIndex: i + 2,
            'Nama Sekolah': r[2] || '-',
            'Tahun Ajaran': r[3] || '-',
            'Semester': r[4] || '-',
            'Nomor SK': r[5] || '-',
            'Dokumen': r[8] || '',
            'Info Hapus': r[10] || '-', // Berisi "Dihapus oleh..."
            'Waktu Hapus': safeDate(r[11], "dd/MM/yyyy HH:mm:ss"),
            // Buat kunci sortir berdasarkan waktu hapus (terbaru diatas)
            _sortKey: (r[11] instanceof Date ? r[11].getTime() : 0)
        };
    });
    
    // Sortir: Yang baru dihapus paling atas
    return result.sort((a,b) => b._sortKey - a._sortKey);
    
  } catch(e) { return { error: e.message }; }
}

function apiGetArsipFiles(folderId) {
  try {
    // 1. Tentukan Target Folder
    var targetId = folderId;
    
    // Jika null/kosong, gunakan Root
    if (!targetId) targetId = ARSIP_ROOT_ID;
    
    var folder = DriveApp.getFolderById(targetId);
    var folders = folder.getFolders();
    var files = folder.getFiles();
    
    var items = [];

    // 2. Ambil Subfolder
    while (folders.hasNext()) {
      var f = folders.next();
      items.push({
        id: f.getId(),
        name: f.getName(),
        type: 'folder',
        mime: 'application/vnd.google-apps.folder',
        updated: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy")
      });
    }

    // 3. Ambil File
    while (files.hasNext()) {
      var f = files.next();
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

    // 4. Sortir (Folder diatas)
    items.sort(function(a, b) {
        if (a.type === b.type) return a.name.localeCompare(b.name);
        return a.type === 'folder' ? -1 : 1;
    });

    return { 
        status: 'success',
        currentId: targetId,
        rootId: ARSIP_ROOT_ID,
        currentName: folder.getName(), 
        items: items 
    };

  } catch (e) {
    return { status: 'error', message: e.message };
  }
}