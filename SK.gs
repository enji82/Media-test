/**
 * ===================================================================
 * MODUL SK PEMBAGIAN TUGAS (FIXED & NO CONFLICT)
 * ===================================================================
 */

// --- 1. KONFIGURASI FOLDER ---
function getSkArsipFolderIds() {
  try { return { 'MAIN_SK': FOLDER_CONFIG.MAIN_SK }; } catch(e) { return {}; }
}

const TRASH_FOLDER_ID = '1OB2Mxa_zvpYl7Vru9NEddYmBlU5SfYHL';
const ARSIP_ROOT_ID = '1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs';

function getTargetFolder(tahun, semester) {
  const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
  return getOrCreateSubfolder(mainFolder, tahun, semester);
}

function getTrashTargetFolder(tahun, semester) {
  const trashRoot = DriveApp.getFolderById(TRASH_FOLDER_ID);
  return getOrCreateSubfolder(trashRoot, tahun, semester);
}

function getOrCreateSubfolder(rootFolder, level1Name, level2Name) {
  let folder1;
  const iter1 = rootFolder.getFoldersByName(level1Name);
  if (iter1.hasNext()) { folder1 = iter1.next(); } 
  else { folder1 = rootFolder.createFolder(level1Name); }

  let folder2;
  const iter2 = folder1.getFoldersByName(level2Name);
  if (iter2.hasNext()) { folder2 = iter2.next(); } 
  else { folder2 = folder1.createFolder(level2Name); }

  return folder2;
}

// --- 2. FUNGSI DROPDOWN (DENGAN DATA CADANGAN) ---
function getMasterSkOptions() {
  try {
    let options = {
        'Tahun Ajaran': [],
        'Semester': [],
        'Kriteria SK': [],
        'Nama Sekolah List': {}
    };

    // A. Coba Ambil dari Spreadsheet (Hanya jika Config tersedia)
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
                    if(data[i][2]) options['Kriteria SK'].push(String(data[i][2]));
                    
                    let status = String(data[i][3]).trim();
                    let nama = String(data[i][4]).trim();
                    if (status && nama) {
                        if (!options['Nama Sekolah List'][status]) options['Nama Sekolah List'][status] = [];
                        options['Nama Sekolah List'][status].push(nama);
                    }
                }
            }
        }
    } catch(err) { console.log("Spreadsheet error: " + err.message); }

    // B. DATA CADANGAN (HARDCODED) - Agar Dropdown PASTI MUNCUL
    if (options['Tahun Ajaran'].length === 0) options['Tahun Ajaran'] = ["2024/2025", "2025/2026", "2026/2027"];
    if (options['Semester'].length === 0) options['Semester'] = ["Ganjil", "Genap"];
    if (options['Kriteria SK'].length === 0) options['Kriteria SK'] = ["PNS", "PPPK", "GTT", "PTT"];
    if (Object.keys(options['Nama Sekolah List']).length === 0) {
        options['Nama Sekolah List'] = {
            'Negeri': ['SD Negeri 1 Secang', 'SD Negeri 2 Secang'],
            'Swasta': ['SD Muhammadiyah', 'SD IT']
        };
    }

    // C. Rapikan Data
    options['Tahun Ajaran'] = [...new Set(options['Tahun Ajaran'])].sort().reverse();
    options['Semester'] = [...new Set(options['Semester'])];
    options['Kriteria SK'] = [...new Set(options['Kriteria SK'])];
    
    return options;

  } catch (e) {
    return { error: "Error Backend: " + e.message };
  }
}

// --- 3. FUNGSI UNGGAH ---
function processManualForm(form) {
    try {
        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
        
        let folderId = FOLDER_CONFIG.MAIN_SK;
        try { folderId = getTargetFolder(form.tahunAjaran, form.semester).getId(); } catch(e){}
        
        const folder = DriveApp.getFolderById(folderId);
        const blob = Utilities.newBlob(Utilities.base64Decode(form.fileData.data), form.fileData.mimeType, form.fileData.fileName);
        const file = folder.createFile(blob);
        
        const row = [
            new Date(), form.statusSekolah, form.namaSekolah, form.tahunAjaran, form.semester,
            form.nomorSK, new Date(form.tanggalSK), form.kriteriaSK, file.getUrl(),
            'Diproses', '-', '', form.loggedInUser || 'User', '-'
        ];
        sheet.appendRow(row);
        return "Berhasil disimpan.";
    } catch(e) { throw new Error(e.message); }
}

// --- 4. FUNGSI TABEL (RIWAYAT & KELOLA) ---
function getSKTableData(isKelola) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if(!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; 
    
    const rows = data.slice(1);
    const safeDate = (v, f) => (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), f) : (v || '-');
    
    const result = rows.map((r, i) => {
        let statusHtml = `<span class="status-badge status-diproses">${r[9]||'Diproses'}</span>`;
        let s = String(r[9]).toLowerCase();
        if(s==='ok'||s==='v'||s==='disetujui') statusHtml = `<span class="status-badge status-ok">OK</span>`;
        else if(s==='ditolak'||s==='x') statusHtml = `<span class="status-badge status-ditolak">Ditolak</span>`;
        else if(s==='revisi') statusHtml = `<span class="status-badge status-revisi">Revisi</span>`;

        let obj = {
            _rowIndex: i + 2,
            'Nama Sekolah': r[2] || '-', 'Tahun Ajaran': r[3] || '-', 'Semester': r[4] || '-',
            'Nomor SK': r[5] || '-', 'Tanggal SK': safeDate(r[6], "dd/MM/yyyy"), 
            'Kriteria SK': r[7] || '-', 'Status': statusHtml, 'Dokumen': r[8] || '',
            'Tanggal Unggah': safeDate(r[0], "dd/MM/yyyy HH:mm:ss"), 'User': r[12] || '-'
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

function getSKDataByRow(rowIndex) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const val = sheet.getRange(rowIndex, 2, 1, 8).getValues()[0];
    return {
      statusSekolah: String(val[0]), namaSekolah: String(val[1]), tahunAjaran: String(val[2]),
      semester: String(val[3]), nomorSK: String(val[4]), 
      tanggalSK: (val[5] instanceof Date) ? Utilities.formatDate(val[5], Session.getScriptTimeZone(), "yyyy-MM-dd") : '',
      kriteriaSK: String(val[6]), dokumen: String(val[7])
    };
  } catch (e) { return { error: e.message }; }
}

function updateSKData(form) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const row = parseInt(form.rowIndex);
    const statusVal = String(sheet.getRange(row, 10).getValue()).trim().toLowerCase();
    
    if (['ok', 'v', 'disetujui'].includes(statusVal)) throw new Error("Data sudah disetujui, tidak bisa diedit.");

    const targetFolder = getTargetFolder(form.tahunAjaran, form.semester);
    if (form.fileData && form.fileData.data) {
       const blob = Utilities.newBlob(Utilities.base64Decode(form.fileData.data), form.fileData.mimeType, form.fileData.fileName);
       sheet.getRange(row, 9).setValue(targetFolder.createFile(blob).getUrl());
    } 

    let newStatus = ['ditolak','x'].includes(statusVal) ? 'Revisi' : (['diproses','','-'].includes(statusVal) ? 'Diproses' : statusVal);
    let oldKet = String(sheet.getRange(row, 11).getValue() || "-").replace(/\s*\(Direvisi oleh.*?\)/gi, "").trim();
    if(oldKet === "-") oldKet = "";
    
    sheet.getRange(row, 2, 1, 7).setValues([[form.statusSekolah, form.namaSekolah, form.tahunAjaran, form.semester, form.nomorSK, form.tanggalSK, form.kriteriaSK]]);
    sheet.getRange(row, 10).setValue(newStatus);
    sheet.getRange(row, 11).setValue(oldKet + " (Direvisi oleh " + (form.editorName||"User") + ")");
    sheet.getRange(row, 12).setValue(new Date());
    sheet.getRange(row, 13).setValue(form.editorName||"User");

    return "Data berhasil diperbarui.";
  } catch (e) { throw new Error(e.message); }
}

function deleteSKData(idx, code, actorName, actorRole) {
    try {
        if(String(code).trim() !== Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd")) throw new Error("Kode hapus salah.");
        const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
        const ss = SpreadsheetApp.openById(config.id);
        const sheet = ss.getSheetByName(config.sheet);
        
        const vals = sheet.getRange(idx, 1, 1, sheet.getLastColumn()).getValues()[0];
        vals[9]="Dihapus"; vals[10]="Dihapus oleh "+actorName; vals[11]=new Date(); vals[12]=actorRole;
        
        let trash = ss.getSheetByName('Trash');
        if(!trash) { trash=ss.insertSheet('Trash'); trash.appendRow(sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]); }
        trash.appendRow(vals);
        
        if(vals[8] && String(vals[8]).includes("drive")) { try{DriveApp.getFileById(vals[8].match(/[-\w]{25,}/)[0]).moveTo(getTrashTargetFolder(vals[3], vals[4]));}catch(e){} }
        
        sheet.deleteRow(idx);
        return "Data berhasil dihapus.";
    } catch(e) { throw new Error(e.message); }
}

function updateSKVerificationStatus(idx, st, ket, adm) {
    const s = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.id).getSheetByName(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.sheet);
    s.getRange(idx, 10).setValue(st); s.getRange(idx, 11).setValue(ket);
    s.getRange(idx, 12).setValue(new Date()); s.getRange(idx, 14).setValue(adm);
    return "Verifikasi disimpan.";
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

function getSKTrashData() {
    try {
        const s = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.id).getSheetByName('Trash');
        if(!s) return [];
        const d = s.getDataRange().getValues();
        if(d.length<2) return [];
        return d.slice(1).map(r => ({
            'Nama Sekolah': r[2], 'Tahun Ajaran': r[3], 'Status': `<span class="status-badge status-rejected">${r[9]}</span>`,
            'Keterangan': r[10], 'Update': Utilities.formatDate(r[11], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")
        })).reverse();
    } catch(e) { return {error:e.message}; }
}

function getDriveContents(fid) { /* Fungsi arsip */ return {}; }

// --- TIDAK ADA FUNGSI INCLUDE/SRVLOADPAGE DISINI ---
// (Karena sudah ada di Code.gs)