/**
 * ===================================================================
 * MODUL SK PEMBAGIAN TUGAS (RENAME FILE & FOLDER OTOMATIS)
 * ===================================================================
 */

function getSKArsipFiles(folderId) {
  try {
    // 1. Tentukan ID Folder Target
    var targetId = folderId;
    if (!targetId) {
        // Gunakan Folder Utama dari Config
        targetId = FOLDER_CONFIG.MAIN_SK; 
    }
    
    if(!targetId) throw new Error("ID Folder SK (MAIN_SK) belum disetting di code.gs.");

    const folder = DriveApp.getFolderById(targetId);
    
    let items = [];

    // 2. Ambil Sub-Folder
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
        let f = subFolders.next();
        items.push({
            id: f.getId(),
            name: f.getName(),
            type: 'folder',
            updated: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
            // Simpan timestamp untuk sorting
            _updatedTime: f.getLastUpdated().getTime() 
        });
    }

    // 3. Ambil File
    const fileIter = folder.getFiles();
    while (fileIter.hasNext()) {
        let f = fileIter.next();
        items.push({
            id: f.getId(),
            name: f.getName(),
            type: 'file',
            mime: f.getMimeType(),
            url: f.getUrl(),
            size: (f.getSize() / 1024).toFixed(0) + ' KB',
            updated: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
             _updatedTime: f.getLastUpdated().getTime()
        });
    }

    // 4. LOGIKA SORTING CUSTOM
    items.sort(function(a, b) {
        // A. Folder selalu di atas File
        if (a.type !== b.type) {
            return a.type === 'folder' ? -1 : 1;
        }

        // B. Sorting Nama
        var nameA = a.name.toLowerCase();
        var nameB = b.name.toLowerCase();

        // Cek apakah nama adalah Tahun (Angka 4 digit, misal 2023, 2024)
        var isYearA = /^\d{4}$/.test(nameA.split('/')[0].trim()); // Cek bagian depan jika ada slash
        var isYearB = /^\d{4}$/.test(nameB.split('/')[0].trim());

        if (isYearA && isYearB) {
             // URUTAN TAHUN: ASCENDING (Terlama -> Terbaru)
             // 2022, 2023, 2024...
             return nameA.localeCompare(nameB, undefined, {numeric: true});
        }

        // Cek Semester (Ganjil/Genap atau 1/2)
        // Kita ingin: Semester 1 (Ganjil) -> Semester 2 (Genap)
        // Logika: Ganjil < Genap secara alfabet? Tidak (Ganjil > Genap).
        // Jadi kita perlu custom check.
        if (nameA.includes('ganjil') && nameB.includes('genap')) return -1; // Ganjil dulu
        if (nameA.includes('genap') && nameB.includes('ganjil')) return 1;
        
        if (nameA.includes('1') && nameB.includes('2') && nameA.includes('semester')) return -1;
        if (nameA.includes('2') && nameB.includes('1') && nameA.includes('semester')) return 1;

        // Default: Alphabetical Ascending (A -> Z)
        return nameA.localeCompare(nameB, undefined, {numeric: true});
    });
    
    // Gabung: Folder duluan, baru File
    return { 
        currentFolder: folder.getName(),
        parentFolder: (targetId !== FOLDER_CONFIG.MAIN_SK) ? folder.getParents().next().getId() : null, 
        items: items 
    };

  } catch (e) {
    return { error: e.message };
  }
}

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

function processManualForm(form) {
    try {
        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
        
        // 1. Validasi File
        if (!form.fileData || !form.fileData.data) {
            throw new Error("File dokumen tidak ditemukan atau korup.");
        }

        // 2. Tentukan Folder Tujuan (Tahun/Semester)
        let targetFolder;
        try { 
            targetFolder = getTargetFolder(form.tahunAjaran, form.semester);
        } catch(e) {
            // Fallback ke Folder Utama jika gagal buat subfolder
            targetFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK); 
        }

        // 3. Simpan File ke Drive
        const blob = Utilities.newBlob(Utilities.base64Decode(form.fileData.data), form.fileData.mimeType, form.fileData.fileName);
        const file = targetFolder.createFile(blob);
        
        // Rename File agar rapi: "SK - [Nama Sekolah] - [Nomor SK]"
        const cleanName = (form.namaSekolah || "Sekolah").replace(/[^a-zA-Z0-9 ]/g, "");
        const cleanNomor = (form.nomorSK || "").replace(/[^a-zA-Z0-9 ]/g, "_");
        file.setName(`SK - ${cleanName} - ${cleanNomor}`);
        
        const fileUrl = file.getUrl();

        // 4. Siapkan Data Baris (Sesuai Urutan Kolom A-P)
        // A: Timestamp
        // B: Status Sekolah
        // C: Nama Sekolah
        // D: Tahun Ajaran
        // E: Semester
        // F: Nomor SK
        // G: Tanggal SK
        // H: Kriteria SK
        // I: Dokumen (URL)
        // J: Status (Default: 'Diproses')
        // K: Keterangan (Default: '-')
        // L: Pengunggah (Nama User)
        // M: Edit (Kosong)
        // N: Editor (Kosong)
        // O: Verifikasi (Kosong)
        // P: Verifikator (Kosong)

        const rowData = [
            new Date(),               // A: Timestamp
            form.statusSekolah,       // B: Status Sekolah
            form.namaSekolah,         // C: Nama Sekolah
            form.tahunAjaran,         // D: Tahun Ajaran
            form.semester,            // E: Semester
            "'" + form.nomorSK,       // F: Nomor SK (Pakai kutip agar tidak jadi rumus)
            form.tanggalSK,           // G: Tanggal SK
            form.kriteriaSK,          // H: Kriteria SK
            fileUrl,                  // I: Dokumen
            "Diproses",               // J: Status
            "-",                      // K: Keterangan
            form.loggedInUser || "User", // L: Pengunggah (Nama User yang Login)
            "",                       // M: Edit
            "",                       // N: Editor
            "",                       // O: Verifikasi
            ""                        // P: Verifikator
        ];

        // 5. Simpan ke Spreadsheet
        sheet.appendRow(rowData);

        return "Berhasil disimpan. Data Anda sedang diproses.";

    } catch (e) {
        throw new Error("Gagal menyimpan: " + e.message);
    }
}

function getSKTableData(isKelola) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if(!sheet) return { rows: [] }; 
    
    // Ambil semua data
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { rows: [] }; 
    
    const rows = data.slice(1); // Buang header
    
    // Helper untuk format tanggal tampilan
    const safeDate = (v) => {
        if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy");
        return (v || '-');
    };
    const safeTime = (v) => {
        if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        return (v || '-');
    };

    // Helper untuk parsing tanggal dari sheet ke Timestamp (miliseconds)
    // Penting: Handle jika cell kosong atau bukan objek Date
    const getTimestamp = (val) => {
        if (val instanceof Date) return val.getTime();
        // Jika string tanggal, coba parse (opsional, tergantung format di sheet)
        return 0; 
    };

    const result = rows.map((r, i) => {
        // Mapping Array Index (0-15) ke Object Keys
        // 0:Time (Unggah), 1:StatusSek, 2:NamaSek, 3:Thn, 4:Smt, 5:NoSK, 6:TglSK, 7:Krit, 
        // 8:Dok, 9:Status, 10:Ket, 11:Pengunggah, 
        // 12:EditTime, 13:Editor, 
        // 14:VerifTime, 15:Verifikator
        
        // --- LOGIKA STATUS BADGE ---
        let statusHtml = `<span class="status-badge status-diproses">${r[9]||'Diproses'}</span>`;
        let s = String(r[9]).toLowerCase();
        if(s.includes('ok')||s.includes('setuju')) statusHtml = `<span class="status-badge status-ok">OK</span>`;
        else if(s.includes('tolak')) statusHtml = `<span class="status-badge status-ditolak">Ditolak</span>`;
        else if(s.includes('revisi')) statusHtml = `<span class="status-badge status-revisi">Revisi</span>`;
        else if(s.includes('hapus')) statusHtml = `<span class="status-badge status-ditolak">Dihapus</span>`;

        // --- LOGIKA SORTING (AMBIL MAX DATE) ---
        // Kita bandingkan:
        // 1. Tanggal Unggah (Kolom A / Index 0)
        // 2. Tanggal Edit (Kolom M / Index 12)
        // 3. Tanggal Verifikasi (Kolom O / Index 14)
        
        let timeUnggah = getTimestamp(r[0]);
        let timeEdit   = getTimestamp(r[12]);
        let timeVerif  = getTimestamp(r[14]);
        
        // Cari waktu paling terakhir (terbesar) dari ketiga aktivitas
        let maxActivityTime = Math.max(timeUnggah, timeEdit, timeVerif);

        let obj = {
            _rowIndex: i + 2,
            
            // Kolom Utama Tampilan
            'Nama Sekolah': r[2] || '-',
            'Tahun Ajaran': r[3] || '-', 
            'Semester': r[4] || '-',
            'Nomor SK': r[5] || '-', 
            'Tanggal SK': safeDate(r[6]), 
            'Kriteria SK': r[7] || '-',
            'Dokumen': r[8] || '',
            'Status': statusHtml,
            '_statusRaw': r[9] || 'Diproses',
            'Keterangan': r[10] || '-',
            
            // Kolom History / Metadata
            'Tanggal Unggah': safeTime(r[0]),
            'Pengunggah': r[11] || '-',
            'Edit': safeTime(r[12]),
            'Editor': r[13] || '-',
            'Verifikasi': safeTime(r[14]),
            'Verifikator': r[15] || '-',
            
            // Simpan kunci sortir (Waktu aktivitas terakhir)
            _sortKey: maxActivityTime
        };
        
        return obj;
    });

    // Sort Descending (Terbaru di atas) berdasarkan _sortKey
    const sortedResult = result.sort((a, b) => b._sortKey - a._sortKey);
    return { rows: sortedResult };

  } catch(e) { 
      return { rows: [], error: e.message }; 
  }
}

function getSKRiwayatData() { return getSKTableData(false); }
function getSKKelolaData() {
  return getSKTableData(true);
}

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

function updateSKData(data) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    var row = parseInt(data.rowIndex);
    if (isNaN(row) || row < 2) throw new Error("ID Baris tidak valid.");

    // Update Kolom Data Utama
    sheet.getRange(row, 4).setValue(data.tahun);    // Col 4: Tahun
    sheet.getRange(row, 5).setValue(data.semester); // Col 5: Semester
    sheet.getRange(row, 6).setValue(data.nomor);    // Col 6: No SK
    sheet.getRange(row, 7).setValue(data.tanggal);  // Col 7: Tgl SK
    sheet.getRange(row, 8).setValue(data.kriteria); // Col 8: Kriteria

    // Update File (Jika ada)
    if (data.fileData) {
        const folderId = FOLDER_CONFIG.MAIN_SK; 
        const folder = DriveApp.getFolderById(folderId);
        const blob = Utilities.newBlob(Utilities.base64Decode(data.fileData.data), data.fileData.mimeType, data.fileData.fileName);
        const file = folder.createFile(blob);
        sheet.getRange(row, 9).setValue(file.getUrl()); // Col 9: Dokumen
    }

    // === LOGIKA HISTORY EDIT ===
    // Reset Status ke Diproses
    sheet.getRange(row, 10).setValue('Diproses'); 
    
    // Simpan Waktu Edit (Col 13 / M)
    sheet.getRange(row, 13).setValue(new Date());
    
    // Simpan Nama Editor (Col 14 / N)
    sheet.getRange(row, 14).setValue(data.editorName || 'User');

    return "Data berhasil diperbarui. Status kembali ke 'Diproses'.";

  } catch (e) {
    throw new Error("Gagal update: " + e.message);
  }
}

function updateSKVerificationStatus(data) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    var row = parseInt(data.rowIndex);
    if (isNaN(row) || row < 2) throw new Error("ID Baris tidak valid.");

    // Update Status (Col 10 / J)
    sheet.getRange(row, 10).setValue(data.status);
    
    // Update Keterangan (Col 11 / K)
    sheet.getRange(row, 11).setValue(data.keterangan || "-");
    
    // Update Waktu Verifikasi (Col 15 / O)
    sheet.getRange(row, 15).setValue(new Date());
    
    // Update Nama Verifikator (Col 16 / P)
    sheet.getRange(row, 16).setValue(data.verifikator || "Admin");

    return "Status verifikasi berhasil disimpan.";

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

function deleteSKData(rowIndex, code, deleterName) {
  try {
    // 1. Validasi Kode Keamanan (YYYYMMDD)
    let todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if(String(code).trim() !== todayCode) {
        throw new Error("Kode keamanan salah. Refresh halaman.");
    }

    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const ss = SpreadsheetApp.openById(config.id);
    const sheet = ss.getSheetByName(config.sheet);
    
    // 2. Siapkan Sheet Trash
    var trashSheet = ss.getSheetByName("Trash");
    if (!trashSheet) {
      trashSheet = ss.insertSheet("Trash");
      // Copy Header dari sheet utama (16 kolom)
      var headers = sheet.getRange(1, 1, 1, 16).getValues()[0];
      trashSheet.appendRow(headers);
    }

    var row = parseInt(rowIndex);
    
    // 3. Ambil Data Asli
    var lastCol = 16; // Kita kunci di 16 kolom sesuai struktur
    var range = sheet.getRange(row, 1, 1, lastCol);
    var values = range.getValues()[0];

    // 4. Modifikasi Data untuk Trash
    // Col 9 (Index 9) = Status
    values[9] = "Dihapus";
    
    // Col 10 (Index 10) = Keterangan
    // Format: "Dihapus oleh <Nama> pada <Waktu>"
    var waktuHapus = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    values[10] = "Dihapus oleh " + (deleterName || "Admin") + " pada " + waktuHapus;

    // 5. Pindahkan
    trashSheet.appendRow(values);
    sheet.deleteRow(row);

    return "Data berhasil dihapus dan.";

  } catch (e) {
    throw new Error("Gagal hapus: " + e.message);
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

function getSKDashboardStats() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if(!sheet) return { error: "Sheet tidak ditemukan" };
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { total: 0 }; 
    
    const rows = data.slice(1);
    
    let stats = {
        total: rows.length,
        today: 0,
        status: { ok: 0, revisi: 0, ditolak: 0, proses: 0 },
        monthlyTrend: Array(12).fill(0),
        semesters: {}, // Untuk menyimpan hitungan per semester
        recent: []
    };
    
    const now = new Date();
    const currentYear = now.getFullYear();
    const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
    
    rows.forEach(row => {
        // 1. PARSING TANGGAL (Kolom A / Index 0)
        // Kita paksa ubah jadi objek Date yang valid
        let dateObj = null;
        let rawDate = row[0];
        
        if (rawDate instanceof Date) {
            dateObj = rawDate;
        } else if (typeof rawDate === 'string') {
            // Coba parsing format string 'dd/MM/yyyy'
            let parts = rawDate.split(' ')[0].split('/');
            if (parts.length === 3) dateObj = new Date(parts[2], parts[1]-1, parts[0]);
        }

        // Hitung Tren & Hari Ini
        if (dateObj && !isNaN(dateObj.getTime())) {
            // Trend hanya tahun ini
            if (dateObj.getFullYear() === currentYear) {
                stats.monthlyTrend[dateObj.getMonth()]++;
            }
            // Hitung Hari Ini
            let rowDateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyyMMdd");
            if (rowDateStr === todayStr) stats.today++;
        }

        // 2. HITUNG STATUS (Kolom J / Index 9)
        let st = String(row[9] || '').toLowerCase();
        if (st.includes('ok') || st.includes('setuju') || st === 'v') stats.status.ok++;
        else if (st.includes('revisi')) stats.status.revisi++;
        else if (st.includes('tolak') || st === 'x') stats.status.ditolak++;
        else stats.status.proses++; // Default Diproses

        // 3. HITUNG SEMESTER (Kolom D=Tahun, E=Semester -> Index 3 & 4)
        let semKey = (row[3] || '?') + " (" + (row[4] || '?') + ")";
        if (!stats.semesters[semKey]) stats.semesters[semKey] = 0;
        stats.semesters[semKey]++;
    });

    // 4. SORTIR SEMESTER (Ambil 2 Terbanyak/Terbaru)
    // Kita ubah object ke array untuk disortir
    let semArray = [];
    for (let key in stats.semesters) {
        semArray.push({ label: key, count: stats.semesters[key] });
    }
    // Sortir berdasarkan Tahun (String descending) agar yang terbaru muncul
    stats.topSemesters = semArray.sort((a,b) => b.label.localeCompare(a.label)).slice(0, 2);

    // 5. RECENT ACTIVITY (5 Terakhir)
    // Ambil dari bawah (data baru biasanya di bawah)
    let recentRows = rows.length > 5 ? rows.slice(rows.length - 5) : rows;
    stats.recent = recentRows.reverse().map(r => ({
        sekolah: r[2],
        waktu: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "dd/MM/yy HH:mm") : String(r[0]).substring(0,16),
        status: r[9]
    }));
    
    return stats;
    
  } catch(e) { return { error: e.message }; }
}

function getSKStatusMatrixAll() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if(!sheet) return { years: [], rows: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { years: [], rows: [] };
    
    const rows = data.slice(1);
    
    // 1. Ambil Tahun Unik
    let years = [...new Set(rows.map(r => r[3]).filter(Boolean))];
    
    // [UBAH DI SINI] 
    // Gunakan .sort() saja agar Ascending (Tahun lama -> Tahun Baru di Kanan)
    years.sort(); 

    // 2. Mapping Data
    let schoolMap = {};

    rows.forEach(r => {
        let sekolah = r[2];
        let tahun = r[3];
        let semester = String(r[4] || "").toLowerCase(); 
        let statusRaw = r[9];

        if (!sekolah) return;
        if (!schoolMap[sekolah]) schoolMap[sekolah] = { nama: sekolah };

        let smtKey = 'Genap'; 
        if (semester.includes('ganjil') || semester.includes('sem 1') || semester.includes('semester 1') || semester.includes(' satu')) {
            smtKey = 'Ganjil';
        }
        
        let key = `${tahun}_${smtKey}`;
        schoolMap[sekolah][key] = statusRaw;
    });

    // 3. Konversi & Sort Sekolah
    let resultRows = Object.values(schoolMap);
    resultRows.sort((a, b) => a.nama.localeCompare(b.nama));

    return { 
        years: years, 
        rows: resultRows 
    };

  } catch (e) {
    return { error: e.message };
  }
}

// Helper untuk mengambil daftar Tahun Ajaran unik (Untuk dropdown filter)
function getListTahunAjaranSK() {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const data = sheet.getDataRange().getValues();
    let years = [];
    for(let i=1; i<data.length; i++) {
        if(data[i][3] && !years.includes(data[i][3])) years.push(data[i][3]);
    }
    return years.sort().reverse();
}