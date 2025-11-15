/**
 * ===================================================================
 * ======================= 1. KONFIGURASI PUSAT ======================
 * ===================================================================
 */
// 1. DAFTAR ID UNIK (BEST PRACTICE: SINGLE SOURCE OF TRUTH)
const SPREADSHEET_IDS = {
  SK_DATA: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4",
  PAUD_DATA: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs",
  SD_DATA: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s",
  
  // ID unik lainnya
  LAPBUL_GABUNGAN: "1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU",
  PTK_PAUD_DB: "1XetGkBymmN2NZQlXpzZ2MQyG0nhhZ0sXEPcNsLffhEU",
  PTK_SD_DB: "1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0",
  DATA_SEKOLAH: "1qeOYVfqFQdoTpysy55UIdKwAJv3VHo4df3g6u6m72Bs",   
  DROPDOWN_DATA: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA",
  FORM_OPTIONS_DB: "1prqqKQBYzkCNFmuzblNAZE41ag9rZTCiY2a0WvZCTvU",
  SIABA_REKAP: "1x3b-yzZbiqP2XfJNRC3XTbMmRTHLd8eEdUqAlKY3v9U",
  SIABA_TIDAK_PRESENSI: "1mjXz5l_cqBiiR3x9qJ7BU4yQ3f0ghERT9ph8CC608Zc",
};

const SPREADSHEET_CONFIG = {
  // --- Modul SK Pembagian Tugas ---
  SK_BAGI_TUGAS: { id: SPREADSHEET_IDS.SK_DATA, sheet: "SK Tabel Kirim" },
  SK_FORM_RESPONSES: { id: SPREADSHEET_IDS.SK_DATA, sheet: "Form Responses 1" },

  // --- Modul Laporan Bulanan & Data Murid ---
  LAPBUL_FORM_RESPONSES_PAUD: { id: SPREADSHEET_IDS.PAUD_DATA, sheet: "Form Responses 1" },
  LAPBUL_FORM_RESPONSES_SD: { id: SPREADSHEET_IDS.SD_DATA, sheet: "Input" },
  LAPBUL_GABUNGAN: { id: SPREADSHEET_IDS.LAPBUL_GABUNGAN },
  // Konfigurasi LAPBUL_RIWAYAT yang dirujuk dari ID pusat
  LAPBUL_RIWAYAT: {
      'PAUD': { 
          id: SPREADSHEET_IDS.PAUD_DATA,
          sheet: 'Form Responses 1' 
      },
      'SD': { 
          id: SPREADSHEET_IDS.SD_DATA,
          sheet: 'Input' 
      },
  },
  LAPBUL_STATUS: {
      'PAUD': { 
          id: SPREADSHEET_IDS.PAUD_DATA,
          sheet: 'Status' 
      },
      'SD': { 
          id: SPREADSHEET_IDS.SD_DATA,
          sheet: 'Status' 
      }
  },

  // --- Modul Data PTK ---
  PTK_PAUD_KEADAAN: { id: SPREADSHEET_IDS.PAUD_DATA, sheet: "Keadaan PTK PAUD" },
  PTK_PAUD_JUMLAH_BULANAN: { id: SPREADSHEET_IDS.PAUD_DATA, sheet: "Jumlah PTK Bulanan" },
  PTK_PAUD_DB: { id: SPREADSHEET_IDS.PTK_PAUD_DB, sheet: "PTK PAUD" },
  PTK_PAUD_TIDAK_AKTIF: { id: SPREADSHEET_IDS.PTK_PAUD_DB, sheet: "PTK PAUD Tidak Aktif" },
  PTK_SD_KEADAAN: { id: SPREADSHEET_IDS.SD_DATA, sheet: "Keadaan PTK SD" },
  PTK_SD_JUMLAH_BULANAN: { id: SPREADSHEET_IDS.SD_DATA, sheet: "PTK Bulanan SD"},
  PTK_SD_KEBUTUHAN: { id: SPREADSHEET_IDS.SD_DATA, sheet: "Kebutuhan Guru"},
  PTK_SD_DB: { id: SPREADSHEET_IDS.PTK_SD_DB },

  // --- Modul Data Murid ---
  MURID_PAUD_KELAS: { id: SPREADSHEET_IDS.PAUD_DATA, sheet: "Murid Kelas" },
  MURID_PAUD_JK: { id: SPREADSHEET_IDS.PAUD_DATA, sheet: "Murid JK" },
  MURID_PAUD_BULANAN: { id: SPREADSHEET_IDS.PAUD_DATA, sheet: "Murid Bulanan" },
  MURID_SD_KELAS: { id: SPREADSHEET_IDS.SD_DATA, sheet: "SD Tabel Kelas" },
  MURID_SD_ROMBEL: { id: SPREADSHEET_IDS.SD_DATA, sheet: "SD Tabel Rombel" },
  MURID_SD_JK: { id: SPREADSHEET_IDS.SD_DATA, sheet: "SD Tabel JK" },
  MURID_SD_AGAMA: { id: SPREADSHEET_IDS.SD_DATA, sheet: "SD Tabel Agama" },
  MURID_SD_BULANAN: { id: SPREADSHEET_IDS.SD_DATA, sheet: "SD Tabel Bulanan" },

  // --- Data Pendukung & Dropdown ---
  DATA_SEKOLAH: { id: SPREADSHEET_IDS.DATA_SEKOLAH },   
  DROPDOWN_DATA: { id: SPREADSHEET_IDS.DROPDOWN_DATA },
  FORM_OPTIONS_DB: { id: SPREADSHEET_IDS.FORM_OPTIONS_DB },

  // --- Data SIABA ---
  SIABA_REKAP: { id: SPREADSHEET_IDS.SIABA_REKAP, sheet: "Rekap Script" },
  SIABA_TIDAK_PRESENSI: { id: SPREADSHEET_IDS.SIABA_TIDAK_PRESENSI, sheet: "Rekap Script" },
};

const FOLDER_CONFIG = {
  MAIN_SK: "1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs",
  LAPBUL_KB: "18CxRT-eledBGRtHW1lFd2AZ8Bub6q5ra",
  LAPBUL_TK: "1WUNz_BSFmcwRVlrG67D2afm9oJ-bVI9H",
  LAPBUL_SD: "1I8DRQYpBbTt1mJwtD1WXVD6UK51TC8El",
};

const STATUS_COLUMNS_MAP = {
    'Tahun': 0, 'Bulan': 1, 'Jenjang': 2, 'Nama Sekolah': 3
};

const PAUD_STATUS_SHEET_ID = SPREADSHEET_IDS.PAUD_DATA;
const SD_STATUS_SHEET_ID = SPREADSHEET_IDS.SD_DATA;
const STATUS_SHEET_NAME = "Status"; // Nama sheet yang digunakan di kedua SS

const LAPBUL_STATUS_HEADERS = [
    "Nama Sekolah", "Jenjang", "Status", "Tahun", "Januari", 
    "Februari", "Maret", "April", "Mei", "Juni", 
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
];

const START_ROW = 2; 
const NUM_COLUMNS_TO_FETCH = 16; 
const STATUS_COLUMN_INDICES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];

const COLUMNS_MAP = {
    // PAUD (Sheet: "Form Responses 1")
    PAUD: {
        "Tanggal Unggah": 0,  // Kolom A
        "Bulan": 1,           // Kolom B
        "Tahun": 2,           // Kolom C
        "Status": 4,         
        "Jenjang": 6,         // Kolom G
        "Nama Sekolah": 7,    // Kolom H
        "Dokumen": 36,        // Kolom AK
        "Update": 37,         // Kolom AL
        "Jumlah Rombel": 5,   // Kolom F
    },
  
   // SD (Sheet: "Input")
    SD: {
        "Tanggal Unggah": 0,  // Kolom A
        "Bulan": 1,           // Kolom B
        "Tahun": 2,           // Kolom C
        "Status": 3,          // Kolom D
        "Nama Sekolah": 4,    // Kolom E
        "Dokumen": 7,         // Kolom H
        "Jenjang": 217,       // Kolom HJ (Indeks 217)
        "Update": 218,        // Kolom HK (Indeks 218)
        "Jumlah Rombel": 6,   // Kolom G (Asumsi)
    }
};

/**
 * ===================================================================
 * ===================== 2. FUNGSI INTI APLIKASI =====================
 * ===================================================================
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ===================================================================
 * ===================== 3. FUNGSI UTILITAS UMUM =====================
 * ===================================================================
 */

function handleError(functionName, error) {
  Logger.log(`Error di ${functionName}: ${error.message}\nStack: ${error.stack}`);
  return { error: error.message }; // Return object error ke client
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}

function getDataFromSheet(configKey) {
  const config = SPREADSHEET_CONFIG[configKey];
  if (!config) throw new Error(`Konfigurasi untuk '${configKey}' tidak ditemukan.`);
  const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
  if (!sheet) throw new Error(`Sheet '${config.sheet}' di spreadsheet '${config.id}' tidak ditemukan.`);
  return sheet.getDataRange().getDisplayValues();
}

function getCachedData(key, fetchFunction) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached != null) {
    return JSON.parse(cached);
  }
  const freshData = fetchFunction();
  cache.put(key, JSON.stringify(freshData), 21600); // Cache for 6 hours
  return freshData;
}


/**
 * ===================================================================
 * =================== MODUL GOOGLE DRIVE (ARSIP) ====================
 * ===================================================================
 */



function getFolders(folderId) {
  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const subFolders = parentFolder.getFolders();
    const folderList = [];
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      folderList.push({
        id: folder.getId(),
        name: folder.getName()
      });
    }
    folderList.sort((a, b) => b.name.localeCompare(a.name));
    return folderList;
  } catch (e) {
    return handleError("getFolders", e);
  }
}

function getFiles(folderId) {
  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const files = parentFolder.getFiles();
    const fileList = [];
    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        name: file.getName(),
        url: file.getUrl()
      });
    }
    fileList.sort((a, b) => a.name.localeCompare(b.name));
    return fileList;
  } catch (e) {
    return handleError("getFiles", e);
  }
}


/**
 * ===================================================================
 * ==================== MODUL DATA LAPORAN BULAN =====================
 * ===================================================================
 */



const PAUD_FORM_INDEX_MAP = [
    // 0-7: Data Lembaga (A-H)
    'Tanggal Unggah', 'Bulan', 'Tahun', 'NPSN', 'Status', 'Jumlah Rombel', 'Jenjang', 'Nama Sekolah', 
    
    // 8-21: Murid Usia (L/P)
    'murid_0_1_L', 'murid_0_1_P', 'murid_1_2_L', 'murid_1_2_P', 'murid_2_3_L', 'murid_2_3_P', 
    'murid_3_4_L', 'murid_3_4_P', 'murid_4_5_L', 'murid_4_5_P', 'murid_5_6_L', 'murid_5_6_P', 
    'murid_6_up_L', 'murid_6_up_P',
    
    // 22-25: Murid Rombel (L/P)
    'kelompok_A_L', 'kelompok_A_P', 'kelompok_B_L', 'kelompok_B_P', 

    // 26-29: PTK Kepsek/Guru
    'kepsek_ASN', 'kepsek_Non_ASN', 'guru_PNS', 'guru_PPPK', 'guru_GTY', 'guru_GTT',
    
    // 30-33: PTK Tendik
    'tendik_Penjaga', 'tendik_TAS', 'tendik_Pustakawan', 'tendik_Lainnya',
    
    'Dokumen', // Kolom Dokumen (Indeks 34)
    'Update' // Kolom Update (Indeks 35)
    // Catatan: Asumsi mapping ini sesuai dengan kolom Spreadsheet PAUD Anda.
];

const SD_FORM_INDEX_MAP = [
  // 0-7: Data Sekolah (A-H)
  'Tanggal Unggah', 'Bulan', 'Tahun', 'Status', 'Nama Sekolah', 'NPSN', 'Rombel', 'Dokumen',
  
  // 8-29: Kelas 1 (I-AC)
  'k1_jumlah_rombel', 'k1_rombel_tunggal_L', 'k1_rombel_tunggal_P',
  'k1_rombel_a_L', 'k1_rombel_a_P', 'k1_rombel_b_L', 'k1_rombel_b_P', 'k1_rombel_c_L', 'k1_rombel_c_P',
  'k1_agama_islam_L', 'k1_agama_islam_P', 'k1_agama_kristen_L', 'k1_agama_kristen_P', 'k1_agama_katolik_L', 'k1_agama_katolik_P', 
  'k1_agama_hindu_L', 'k1_agama_hindu_P', 'k1_agama_buddha_L', 'k1_agama_buddha_P', 'k1_agama_konghucu_L', 'k1_agama_konghucu_P',
  
  // 30-51: Kelas 2 (AD-AX)
  'k2_jumlah_rombel', 'k2_rombel_tunggal_L', 'k2_rombel_tunggal_P',
  'k2_rombel_a_L', 'k2_rombel_a_P', 'k2_rombel_b_L', 'k2_rombel_b_P', 'k2_rombel_c_L', 'k2_rombel_c_P',
  'k2_agama_islam_L', 'k2_agama_islam_P', 'k2_agama_kristen_L', 'k2_agama_kristen_P', 'k2_agama_katolik_L', 'k2_agama_katolik_P', 
  'k2_agama_hindu_L', 'k2_agama_hindu_P', 'k2_agama_buddha_L', 'k2_agama_buddha_P', 'k2_agama_konghucu_L', 'k2_agama_konghucu_P',

  // 52-73: Kelas 3 (AY-BS)
  'k3_jumlah_rombel', 'k3_rombel_tunggal_L', 'k3_rombel_tunggal_P',
  'k3_rombel_a_L', 'k3_rombel_a_P', 'k3_rombel_b_L', 'k3_rombel_b_P', 'k3_rombel_c_L', 'k3_rombel_c_P',
  'k3_agama_islam_L', 'k3_agama_islam_P', 'k3_agama_kristen_L', 'k3_agama_kristen_P', 'k3_agama_katolik_L', 'k3_agama_katolik_P', 
  'k3_agama_hindu_L', 'k3_agama_hindu_P', 'k3_agama_buddha_L', 'k3_agama_buddha_P', 'k3_agama_konghucu_L', 'k3_agama_konghucu_P',
  
  // 74-95: Kelas 4 (BT-CN)
  'k4_jumlah_rombel', 'k4_rombel_tunggal_L', 'k4_rombel_tunggal_P',
  'k4_rombel_a_L', 'k4_rombel_a_P', 'k4_rombel_b_L', 'k4_rombel_b_P', 'k4_rombel_c_L', 'k4_rombel_c_P',
  'k4_agama_islam_L', 'k4_agama_islam_P', 'k4_agama_kristen_L', 'k4_agama_kristen_P', 'k4_agama_katolik_L', 'k4_agama_katolik_P', 
  'k4_agama_hindu_L', 'k4_agama_hindu_P', 'k4_agama_buddha_L', 'k4_agama_buddha_P', 'k4_agama_konghucu_L', 'k4_agama_konghucu_P',

  // 96-117: Kelas 5 (CO-DI)
  'k5_jumlah_rombel', 'k5_rombel_tunggal_L', 'k5_rombel_tunggal_P',
  'k5_rombel_a_L', 'k5_rombel_a_P', 'k5_rombel_b_L', 'k5_rombel_b_P', 'k5_rombel_c_L', 'k5_rombel_c_P',
  'k5_agama_islam_L', 'k5_agama_islam_P', 'k5_agama_kristen_L', 'k5_agama_kristen_P', 'k5_agama_katolik_L', 'k5_agama_katolik_P', 
  'k5_agama_hindu_L', 'k5_agama_hindu_P', 'k5_agama_buddha_L', 'k5_agama_buddha_P', 'k5_agama_konghucu_L', 'k5_agama_konghucu_P',

  // 118-139: Kelas 6 (DJ-ED)
  'k6_jumlah_rombel', 'k6_rombel_tunggal_L', 'k6_rombel_tunggal_P',
  'k6_rombel_a_L', 'k6_rombel_a_P', 'k6_rombel_b_L', 'k6_rombel_b_P', 'k6_rombel_c_L', 'k6_rombel_c_P',
  'k6_agama_islam_L', 'k6_agama_islam_P', 'k6_agama_kristen_L', 'k6_agama_kristen_P', 'k6_agama_katolik_L', 'k6_agama_katolik_P', 
  'k6_agama_hindu_L', 'k6_agama_hindu_P', 'k6_agama_buddha_L', 'k6_agama_buddha_P', 'k6_agama_konghucu_L', 'k6_agama_konghucu_P',

  // --- INI ADALAH "PETA DATA" BARU (83 KOLOM PTK) ---
  // 140-142: Kepsek (3)
  'ptk_kepsek_pns', 'ptk_kepsek_pppk', 'ptk_kepsek_nonasn',
  
  // 143-147: Guru Kelas (5)
  'ptk_guru_kelas_pns', 'ptk_guru_kelas_pppk', 'ptk_guru_kelas_pppkpw', 'ptk_guru_kelas_gty', 'ptk_guru_kelas_gtt',
  // 148-152: Guru PAI (5)
  'ptk_guru_pai_pns', 'ptk_guru_pai_pppk', 'ptk_guru_pai_pppkpw', 'ptk_guru_pai_gty', 'ptk_guru_pai_gtt',
  // 153-157: Guru PJOK (5)
  'ptk_guru_pjok_pns', 'ptk_guru_pjok_pppk', 'ptk_guru_pjok_pppkpw', 'ptk_guru_pjok_gty', 'ptk_guru_pjok_gtt',
  // 158-162: Guru PA Kristen (5)
  'ptk_guru_kristen_pns', 'ptk_guru_kristen_pppk', 'ptk_guru_kristen_pppkpw', 'ptk_guru_kristen_gty', 'ptk_guru_kristen_gtt',
  // 163-167: Guru PA Katolik (5)
  'ptk_guru_katolik_pns', 'ptk_guru_katolik_pppk', 'ptk_guru_katolik_pppkpw', 'ptk_guru_katolik_gty', 'ptk_guru_katolik_gtt',
  // 168-172: Guru Bhs. Inggris (5)
  'ptk_guru_inggris_pns', 'ptk_guru_inggris_pppk', 'ptk_guru_inggris_pppkpw', 'ptk_guru_inggris_gty', 'ptk_guru_inggris_gtt',
  // 173-177: Guru Mapel Lain (5)
  'ptk_guru_lainnya_pns', 'ptk_guru_lainnya_pppk', 'ptk_guru_lainnya_pppkpw', 'ptk_guru_lainnya_gty', 'ptk_guru_lainnya_gtt',
  
  // 178-182: Tendik: Pengelola Umum (5)
  'ptk_tendik_pengelola_umum_pns', 'ptk_tendik_pengelola_umum_pppk', 'ptk_tendik_pengelola_umum_pppkpw', 'ptk_tendik_pengelola_umum_pty', 'ptk_tendik_pengelola_umum_ptt',
  // 183-187: Tendik: Operator (5)
  'ptk_tendik_operator_pns', 'ptk_tendik_operator_pppk', 'ptk_tendik_operator_pppkpw', 'ptk_tendik_operator_pty', 'ptk_tendik_operator_ptt',
  // 188-192: Tendik: Pengelola Layanan (5)
  'ptk_tendik_pengelola_layanan_pns', 'ptk_tendik_pengelola_layanan_pppk', 'ptk_tendik_pengelola_layanan_pppkpw', 'ptk_tendik_pengelola_layanan_pty', 'ptk_tendik_pengelola_layanan_ptt',
  // 193-197: Tendik: Penata (5)
  'ptk_tendik_penata_pns', 'ptk_tendik_penata_pppk', 'ptk_tendik_penata_pppkpw', 'ptk_tendik_penata_pty', 'ptk_tendik_penata_ptt',
  // 198-202: Tendik: Adm Perkantoran (5)
  'ptk_tendik_adm_pns', 'ptk_tendik_adm_pppk', 'ptk_tendik_adm_pppkpw', 'ptk_tendik_adm_pty', 'ptk_tendik_adm_ptt',
  // 203-207: Tendik: Penjaga (5)
  'ptk_tendik_penjaga_pns', 'ptk_tendik_penjaga_pppk', 'ptk_tendik_penjaga_pppkpw', 'ptk_tendik_penjaga_pty', 'ptk_tendik_penjaga_ptt',
  // 208-212: Tendik: TAS (5)
  'ptk_tendik_tas_pns', 'ptk_tendik_tas_pppk', 'ptk_tendik_tas_pppkpw', 'ptk_tendik_tas_pty', 'ptk_tendik_tas_ptt',
  // 213-217: Tendik: Pustakawan (5)
  'ptk_tendik_pustakawan_pns', 'ptk_tendik_pustakawan_pppk', 'ptk_tendik_pustakawan_pppkpw', 'ptk_tendik_pustakawan_pty', 'ptk_tendik_pustakawan_ptt',
  // 218-222: Tendik: Lainnya (5)
  'ptk_tendik_lainnya_pns', 'ptk_tendik_lainnya_pppk', 'ptk_tendik_lainnya_pppkpw', 'ptk_tendik_lainnya_pty', 'ptk_tendik_lainnya_ptt',

  // 223: Jenjang
  'Jenjang',
  // 224: Update
  'Update'
];



/**
 * ===================================================================
 * ======================== MODUL: DATA MURID ========================
 * ===================================================================
 */

function getMuridPaudKelasUsiaData() {
  try {
    return getDataFromSheet('MURID_PAUD_KELAS');
  } catch (e) {
    return handleError('getMuridPaudKelasUsiaData', e);
  }
}

function getMuridPaudJenisKelaminData() {
  try {
    return getDataFromSheet('MURID_PAUD_JK');
  } catch (e) {
    return handleError('getMuridPaudJenisKelaminData', e);
  }
}

function getMuridPaudJumlahBulananData() {
  try {
    return getDataFromSheet('MURID_PAUD_BULANAN');
  } catch (e) {
    return handleError('getMuridPaudJumlahBulananData', e);
  }
}

function getMuridSdKelasData() {
  try {
    return getDataFromSheet('MURID_SD_KELAS');
  } catch (e) {
    return handleError('getMuridSdKelasData', e);
  }
}

function getMuridSdRombelData() {
  try {
    return getDataFromSheet('MURID_SD_ROMBEL');
  } catch (e) {
    return handleError('getMuridSdRombelData', e);
  }
}

function getMuridSdJenisKelaminData() {
  try {
    return getDataFromSheet('MURID_SD_JK');
  } catch (e) {
    return handleError('getMuridSdJenisKelaminData', e);
  }
}

function getMuridSdAgamaData() {
  try {
    return getDataFromSheet('MURID_SD_AGAMA');
  } catch (e) {
    return handleError('getMuridSdAgamaData', e);
  }
}

function getMuridSdJumlahBulananData() {
  try {
    const config = SPREADSHEET_CONFIG.MURID_SD_BULANAN;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan");
    return sheet.getDataRange().getValues(); // Menggunakan getValues() untuk angka
  } catch (e) {
    return handleError('getMuridSdJumlahBulananData', e);
  }
}

/**
 * ===================================================================
 * ========================= MODUL: DATA PTK =========================
 * ===================================================================
 */
 



/**
 * ===================================================================
 * ========================== MODUL: DATA SK =========================
 * ===================================================================
 */





/**
 * [REFACTOR - FINAL V4] Mengambil data riwayat pengiriman SK.
 * Memperbaiki parsing tanggal untuk pengurutan yang benar.
 */


/**
 * ===================================================================
 * ======================== MODUL: DATA SIABA ========================
 * ===================================================================
 */

function getSiabaFilterOptions() {
  try {
    const ssDropdown = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
    const sheetUnitKerja = ssDropdown.getSheetByName("Unit Siaba");
    let unitKerjaOptions = [];
    if (sheetUnitKerja && sheetUnitKerja.getLastRow() > 1) {
      unitKerjaOptions = sheetUnitKerja.getRange(2, 1, sheetUnitKerja.getLastRow() - 1, 1)
                                      .getDisplayValues().flat().filter(Boolean).sort();
    }

    const ssSiaba = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SIABA_REKAP.id);
    const sheetSiaba = ssSiaba.getSheetByName(SPREADSHEET_CONFIG.SIABA_REKAP.sheet);
    if (!sheetSiaba || sheetSiaba.getLastRow() < 2) {
         throw new Error("Sheet Rekap SIABA tidak ditemukan atau kosong.");
    }

    const tahunBulanData = sheetSiaba.getRange(2, 1, sheetSiaba.getLastRow() - 1, 2).getDisplayValues();
    const uniqueTahun = [...new Set(tahunBulanData.map(row => row[0]))].filter(Boolean).sort().reverse();
    const uniqueBulan = [...new Set(tahunBulanData.map(row => row[1]))].filter(Boolean);
    
    const monthOrder = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    uniqueBulan.sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));

    return {
      'Tahun': uniqueTahun,
      'Bulan': uniqueBulan,
      'Unit Kerja': unitKerjaOptions
    };
  } catch (e) {
    return handleError('getSiabaFilterOptions', e);
  }
}

function getSiabaPresensiData(filters) {
  try {
    const { tahun, bulan } = filters;
    if (!tahun || !bulan) throw new Error("Filter Tahun dan Bulan wajib diisi.");

    const config = SPREADSHEET_CONFIG.SIABA_REKAP;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const filteredRows = dataRows.filter(row => String(row[0]) === String(tahun) && String(row[1]) === String(bulan));

    const startIndex = 2;
    const endIndex = 86;

    const displayHeaders = headers.slice(startIndex, endIndex + 1);
    const displayRows = filteredRows.map(row => row.slice(startIndex, endIndex + 1));
    
    const tpIndex = displayHeaders.indexOf('TP');
    const taIndex = displayHeaders.indexOf('TA');
    const plaIndex = displayHeaders.indexOf('PLA');
    const tapIndex = displayHeaders.indexOf('TAp');
    const tuIndex = displayHeaders.indexOf('TU');
    const namaIndex = displayHeaders.indexOf('Nama');

    if (tpIndex !== -1) {
      displayRows.sort((a, b) => {
        const compareDesc = (index) => {
            if (index === -1) return 0;
            const valB = parseInt(b[index], 10) || 0;
            const valA = parseInt(a[index], 10) || 0;
            return valB - valA;
        };
        
        let diff = compareDesc(tpIndex);
        if (diff !== 0) return diff;
        diff = compareDesc(taIndex);
        if (diff !== 0) return diff;
        diff = compareDesc(plaIndex);
        if (diff !== 0) return diff;
        diff = compareDesc(tapIndex);
        if (diff !== 0) return diff;
        diff = compareDesc(tuIndex);
        if (diff !== 0) return diff;

        if (namaIndex !== -1) {
            return (a[namaIndex] || "").localeCompare(b[namaIndex] || "");
        }
        return 0;
      });
    }

    return { headers: displayHeaders, rows: displayRows };

  } catch (e) {
    return handleError('getSiabaPresensiData', e);
  }
}

function getSiabaTidakPresensiFilterOptions() {
  try {
    const config = SPREADSHEET_CONFIG.SIABA_TIDAK_PRESENSI;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
         throw new Error("Sheet 'Rekap Script' untuk data Tidak Presensi tidak ditemukan atau kosong.");
    }

    const filterData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
    const uniqueTahun = [...new Set(filterData.map(row => row[0]))].filter(Boolean).sort().reverse();
    const uniqueBulan = [...new Set(filterData.map(row => row[1]))].filter(Boolean);
    const uniqueUnitKerja = [...new Set(filterData.map(row => row[2]))].filter(Boolean).sort();
    
    const monthOrder = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    uniqueBulan.sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));

    return {
      'Tahun': uniqueTahun,
      'Bulan': uniqueBulan,
      'Unit Kerja': uniqueUnitKerja
    };
  } catch (e) {
    return handleError('getSiabaTidakPresensiFilterOptions', e);
  }
}

function getSiabaTidakPresensiData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    if (!tahun || !bulan) throw new Error("Filter Tahun dan Bulan wajib diisi.");

    const config = SPREADSHEET_CONFIG.SIABA_TIDAK_PRESENSI;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const filteredRows = dataRows.filter(row => {
      const tahunMatch = String(row[0]) === String(tahun);
      const bulanMatch = String(row[1]) === String(bulan);
      const unitKerjaMatch = (unitKerja === "Semua") || (String(row[2]) === String(unitKerja));
      return tahunMatch && bulanMatch && unitKerjaMatch;
    });

    const startIndex = 3;
    const endIndex = 8;

    const displayHeaders = headers.slice(startIndex, endIndex + 1);
    const displayRows = filteredRows.map(row => row.slice(startIndex, endIndex + 1));
    
    const jumlahIndex = displayHeaders.indexOf('Jumlah');
    const namaIndex = displayHeaders.indexOf('Nama');

    displayRows.sort((a, b) => {
      const valB_jumlah = (jumlahIndex !== -1) ? (parseInt(b[jumlahIndex], 10) || 0) : 0;
      const valA_jumlah = (jumlahIndex !== -1) ? (parseInt(a[jumlahIndex], 10) || 0) : 0;
      if (valB_jumlah !== valA_jumlah) {
        return valB_jumlah - valA_jumlah;
      }
      if (namaIndex !== -1) {
          return (a[namaIndex] || "").localeCompare(b[namaIndex] || "");
      }
      return 0;
    });

    return { headers: displayHeaders, rows: displayRows };
  } catch (e) {
    return handleError('getSiabaTidakPresensiData', e);
  }
}