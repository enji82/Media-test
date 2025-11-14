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
 
function getPtkKeadaanPaudData() {
    try {
        const data = getDataFromSheet('PTK_PAUD_KEADAAN');
        if (!data || data.length < 2) {
            return { headers: [], rows: [], filterConfigs: [] };
        }
        
        const headers = data[0];
        const dataRows = data.slice(1);
        const jenjangIndex = headers.indexOf('Jenjang');

        const processedRows = dataRows.map(row => {
            const rowObject = {};
            headers.forEach((h, i) => rowObject[h] = row[i]);
            rowObject['_filterJenjang'] = row[jenjangIndex];
            return rowObject;
        });

        return { 
            headers: headers, 
            rows: processedRows,
            filterConfigs: [{ id: 'filterJenjang', dataColumn: '_filterJenjang' }]
        };
    } catch (e) {
        return handleError('getPtkKeadaanPaudData', e);
    }
}
 
function getPtkKeadaanSdData() {
  try {
    return getDataFromSheet('PTK_SD_KEADAAN');
  } catch (e) {
    return handleError('getKeadaanPtkSdData', e);
  }
}

function getPtkJumlahBulananPaudData() {
  try {
    const allData = getDataFromSheet('PTK_PAUD_JUMLAH_BULANAN');
    if (!allData || allData.length < 2) {
      return { headers: [], rows: [], filterConfigs: [] };
    }
    
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const jenjangIndex = headers.indexOf('Jenjang');
    const tahunIndex = headers.indexOf('Tahun');
    const bulanIndex = headers.indexOf('Bulan');
    
    const processedRows = dataRows.map(row => {
        const rowObject = {};
        headers.forEach((h, i) => rowObject[h] = row[i]);
        rowObject['_filterJenjang'] = row[jenjangIndex];
        rowObject['_filterTahun'] = row[tahunIndex];
        rowObject['_filterBulan'] = row[bulanIndex];
        return rowObject;
    });

    return {
        headers: headers,
        rows: processedRows,
        filterConfigs: [
            { id: 'filterTahun', dataColumn: '_filterTahun', sortReverse: true },
            { id: 'filterBulan', dataColumn: '_filterBulan', specialSort: 'bulan' },
            { id: 'filterJenjang', dataColumn: '_filterJenjang' },
            { id: 'filterNamaLembaga', dataColumn: 'Nama Lembaga', dependsOn: 'filterJenjang', dependencyColumn: '_filterJenjang' }
        ]
    };
  } catch (e) {
    return handleError('getPtkJumlahBulananPaudData', e);
  }
}

function getDaftarPtkPaudData() {
  try {
    const allData = getDataFromSheet('PTK_PAUD_DB');
    if (!allData || allData.length < 2) {
        return { headers: [], rows: [], filterConfigs: [] };
    }
    
    // Perbaikan 1: Membersihkan spasi dari semua header
    const headers = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    // Perbaikan 2: Menemukan indeks untuk SEMUA 6 filter
    const jenjangIndex = headers.indexOf('Jenjang');
    const lembagaIndex = headers.indexOf('Nama Lembaga');
    const statusIndex = headers.indexOf('Status');
    const pendidikanIndex = headers.indexOf('Pendidikan');
    const serdikIndex = headers.indexOf('Serdik');
    const dapodikIndex = headers.indexOf('Dapodik'); // <-- TARGET KITA

    // --- JEBAKAN LOG DIMULAI ---
    // Log ini akan mencatat apa yang dilihat server TEPAT SETELAH membersihkan spasi
    Logger.log("--- JEBAKAN LOG: getDaftarPtkPaudData ---");
    Logger.log("Header yang Ditemukan (setelah .trim()): " + JSON.stringify(headers));
    Logger.log("Indeks 'Dapodik' (Harusnya BUKAN -1): " + dapodikIndex);
    // --- AKHIR JEBAKAN LOG ---
    
    const processedRows = dataRows.map((row, i) => { // 'i' ditambahkan untuk index
        const rowObject = {};
        // Perbaikan 3: Loop ini menyalin SEMUA data
        headers.forEach((h, i) => {
          rowObject[h] = row[i];
        });
        
        // Perbaikan 4: Membuat data untuk SEMUA 6 filter
        rowObject['_filterJenjang'] = row[jenjangIndex];
        rowObject['_filterNamaLembaga'] = row[lembagaIndex];
        rowObject['_filterStatus'] = row[statusIndex];
        rowObject['_filterPendidikan'] = row[pendidikanIndex];
        rowObject['_filterSerdik'] = row[serdikIndex];
        rowObject['_filterDapodik'] = row[dapodikIndex]; // <-- TARGET KITA
        
        // --- JEBAKAN LOG 2: Catat data baris pertama ---
        if (i === 0) { // Hanya log baris data pertama
            Logger.log("Data Baris Pertama (Mentah): " + JSON.stringify(row));
            Logger.log("Data Baris Pertama (Diproses): " + JSON.stringify(rowObject));
            Logger.log("Nilai _filterDapodik (Harusnya 'Ya'/'Tidak'): " + rowObject['_filterDapodik']);
        }
        // --- AKHIR JEBAKAN LOG 2 ---

        return rowObject;
    });

    // (Urutkan data)
    processedRows.sort((a,b) => (a['Nama'] || "").localeCompare(b['Nama'] || ""));

    // (Kirim kembali header asli dari sheet)
    return {
        headers: headers,
        rows: processedRows,
        filterConfigs: [
            { id: 'filterJenjang', dataColumn: '_filterJenjang' },
            { id: 'filterNamaLembaga', dataColumn: '_filterNamaLembaga', dependsOn: 'filterJenjang', dependencyColumn: '_filterJenjang' },
            { id: 'filterStatus', dataColumn: '_filterStatus' },
            { id: 'filterPendidikan', dataColumn: '_filterPendidikan' },
            { id: 'filterSerdik', dataColumn: '_filterSerdik' },
            { id: 'filterDapodik', dataColumn: '_filterDapodik' }
        ]
    };

  } catch (e) {
    Logger.log("--- ERROR DI DALAM getDaftarPtkPaudData ---"); // Log error tambahan
    Logger.log(e);
    return handleError('getDaftarPtkPaudData', e);
  }
}

function getKelolaPtkPaudData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

    // 1. Ambil data mentah (untuk tanggal) dan data tampilan (untuk string)
    const allData = sheet.getDataRange().getValues();
    const allDisplayData = sheet.getDataRange().getDisplayValues();

    // 2. Bersihkan header (ini penting untuk 'Dapodik')
    const headers = allDisplayData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);
    const displayRows = allDisplayData.slice(1);

    const updateIndex = headers.indexOf('Update');
    const dateInputIndex = headers.indexOf('Tanggal Input');
    const parseDate = (value) => value instanceof Date && !isNaN(value) ? value.getTime() : 0;

    // 3. Gabungkan data dengan indeks baris aslinya
    const indexedData = dataRows.map((row, index) => ({
      originalRowIndex: index + 2,
      rowData: row,
      displayRow: displayRows[index]
    }));

    // 4. Urutkan berdasarkan Update (terbaru) lalu Tanggal Input (terbaru)
    indexedData.sort((a, b) => {
    // 1. Dapatkan tanggal mentah untuk baris A
    const updateA = parseDate(a.rowData[updateIndex]);
    const dateInputA = parseDate(a.rowData[dateInputIndex]);
    // 2. Tentukan tanggal terbaru untuk baris A (mana yang lebih besar)
    const newestDateA = Math.max(updateA, dateInputA);

    // 3. Dapatkan tanggal mentah untuk baris B
    const updateB = parseDate(b.rowData[updateIndex]);
    const dateInputB = parseDate(b.rowData[dateInputIndex]);
    // 4. Tentukan tanggal terbaru untuk baris B
    const newestDateB = Math.max(updateB, dateInputB);

    // 5. Urutkan berdasarkan tanggal terbaru (dari besar ke kecil)
    return newestDateB - newestDateA;
      });

    // 5. Susun ulang data menjadi objek yang rapi
    const finalData = indexedData.map(item => {
      const rowDataObject = {
        _rowIndex: item.originalRowIndex,
        _source: 'PAUD',
      };

      headers.forEach((header, i) => {
         // Pastikan Tanggal Input (Q) dan Update (R) diformat dengan benar
         if (header === 'Tanggal Input' || header === 'Update') {
            const rawDate = item.rowData[i];
            if (rawDate instanceof Date) {
               rowDataObject[header] = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
            } else {
               rowDataObject[header] = item.displayRow[i] || ""; // Fallback
            }
         } else {
            rowDataObject[header] = item.displayRow[i] || ""; // Gunakan display value untuk string
         }
      });
      return rowDataObject;
    });

    // 6. Kirim SEMUA header, biarkan javascript yang memilih
    return { headers: headers, rows: finalData }; 
  } catch (e) {
    return handleError('getKelolaPtkPaudData', e);
  }
}

function getPtkPaudDataByRow(rowIndex) {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const displayValues = sheet.getRange(rowIndex, 1, 1, headers.length).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      if (header === 'TMT' && values[i] instanceof Date) {
        rowData[header] = Utilities.formatDate(values[i], "UTC", "yyyy-MM-dd");
      } else {
        rowData[header] = displayValues[i];
      }
    });
    return rowData;
  } catch (e) {
    return handleError('getPtkPaudDataByRow', e);
  }
}

function updatePtkPaudData(formData) {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan.");

    const rowIndex = formData.rowIndex;
    if (!rowIndex) throw new Error("Row index tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const oldValues = range.getValues()[0];

    formData['Update'] = new Date();

    const newRowValues = headers.map((header, index) => {
      // Cek apakah data baru ada di formData
      if (formData.hasOwnProperty(header)) {

        // --- ▼▼▼ TAKTIK BARU UNTUK TMT ▼▼▼ ---
        if (header === 'TMT' && formData[header]) {
          // formData[header] adalah string "yyyy-mm-dd"
          return new Date(formData[header]);
        }
        // --- ▲▲▲ AKHIR TAKTIK ▲▲▲ ---

        return formData[header]; // Ambil data baru
      }
      return oldValues[index]; // Jika tidak ada, pertahankan data lama
    });

    range.setValues([newRowValues]); // 1. Simpan datanya

    // --- ▼▼▼ MISI FORMATTING (BARU) ▼▼▼ ---
    const tmtIndex = headers.indexOf('TMT');
    if (tmtIndex !== -1) {
      // 2. Terapkan format "dd-mm-yyyy" ke sel yang baru di-update
      sheet.getRange(rowIndex, tmtIndex + 1).setNumberFormat("dd-MM-yyyy");
    }
    // --- ▲▲▲ AKHIR MISI ▲▲▲ ---

    return "Data PTK berhasil diperbarui.";
  } catch (e) {
    throw new Error(`Gagal memperbarui data: ${e.message}`);
  }
}

function getNewPtkPaudOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.FORM_OPTIONS_DB.id);
    const sheet = ss.getSheetByName("Form PAUD");
    if (!sheet) throw new Error("Sheet 'Form PAUD' tidak ditemukan.");

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

    const jenjangOptions = [...new Set(data.map(row => row[0]).filter(Boolean))].sort();
    const statusOptions = [...new Set(data.map(row => row[2]).filter(Boolean))].sort();
    const jabatanOptions = [...new Set(data.map(row => row[3]).filter(Boolean))].sort();
    
    const lembagaMap = {};
    data.forEach(row => {
      const jenjang = row[0];
      const lembaga = row[1];
      if (jenjang && lembaga) {
        if (!lembagaMap[jenjang]) lembagaMap[jenjang] = [];
        if (!lembagaMap[jenjang].includes(lembaga)) lembagaMap[jenjang].push(lembaga);
      }
    });
    for (const jenjang in lembagaMap) {
      lembagaMap[jenjang].sort();
    }

    return {
      'Jenjang': jenjangOptions,
      'Nama Lembaga': lembagaMap,
      'Status': statusOptions,
      'Jabatan': jabatanOptions
    };
  } catch (e) {
    return handleError('getNewPtkPaudOptions', e);
  }
}

function addNewPtkPaud(formData) {
  try {
    const config = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

    const newRow = headers.map(header => {
      if (header === 'Tanggal Input') return new Date();

      // --- ▼▼▼ TAKTIK BARU UNTUK TMT ▼▼▼ ---
      if (header === 'TMT' && formData[header]) {
        // formData[header] adalah string "yyyy-mm-dd"
        // new Date(...) akan mengkonversinya ke Objek Date
        return new Date(formData[header]); 
      }
      // --- ▲▲▲ AKHIR TAKTIK ▲▲▲ ---

      return formData[header] || "";
    });

    sheet.appendRow(newRow); // 1. Simpan datanya

    // --- ▼▼▼ MISI FORMATTING (BARU) ▼▼▼ ---
    const lastRow = sheet.getLastRow();
    const tmtIndex = headers.indexOf('TMT');
    if (tmtIndex !== -1) {
      // 2. Terapkan format "dd-mm-yyyy" ke sel yang baru ditambahkan
      sheet.getRange(lastRow, tmtIndex + 1).setNumberFormat("dd-MM-yyyy");
    }
    // --- ▲▲▲ AKHIR MISI ▲▲▲ ---

    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

function deletePtkPaudData(rowIndex, deleteCode, alasan) {
  try {
    // 1. Validasi Kode Hapus
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");
    if (!alasan || String(alasan).trim() === "") throw new Error("Alasan tidak boleh kosong.");

    // 2. Buka kedua sheet (Sumber dan Target)
    const configSumber = SPREADSHEET_CONFIG.PTK_PAUD_DB;
    const configTarget = SPREADSHEET_CONFIG.PTK_PAUD_TIDAK_AKTIF;
    const ss = SpreadsheetApp.openById(configSumber.id); // Asumsi ID-nya sama
    
    const sheetSumber = ss.getSheetByName(configSumber.sheet);
    const sheetTarget = ss.getSheetByName(configTarget.sheet);
    if (!sheetSumber || !sheetTarget) throw new Error("Sheet 'PTK PAUD' atau 'PTK PAUD Tidak Aktif' tidak ditemukan.");

    // 3. Ambil Headers dari kedua sheet (bersihkan spasi)
    const headersSumber = sheetSumber.getRange(1, 1, 1, sheetSumber.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
    const headersTarget = sheetTarget.getRange(1, 1, 1, sheetTarget.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());

    // 4. Ambil data baris yang akan 'dihapus' (mentah, untuk menjaga format tanggal)
    const dataBarisSumber = sheetSumber.getRange(rowIndex, 1, 1, headersSumber.length).getValues()[0];

    // 5. Bangun baris baru untuk sheet "Tidak Aktif" (ini adalah 1D array, [data1, data2, ...])
    const barisBaruTarget = headersTarget.map(headerTarget => {
        if (headerTarget === 'Alasan') {
            return alasan; // Masukkan alasan
        }
        if (headerTarget === 'Update') {
            return new Date(); // Masukkan waktu pemindahan
        }
        
        // Cari data yang cocok dari sheet sumber
        const indexDiSumber = headersSumber.indexOf(headerTarget);
        if (indexDiSumber !== -1) {
            return dataBarisSumber[indexDiSumber]; // Salin data
        }
        return ""; // Kolom ada di target tapi tidak di sumber
    });

    // --- ▼▼▼ INI ADALAH PERBAIKAN UTAMA ▼▼▼ ---

    // 6. Buat baris kosong di Bawah Header (di Baris 2)
    sheetTarget.insertRowAfter(1); 

    // 7. Ambil range baris baru (Baris 2) dan isi datanya
    // (setValues MENGHARUSKAN 2D array, jadi kita bungkus [barisBaruTarget])
    sheetTarget.getRange(2, 1, 1, barisBaruTarget.length).setValues([barisBaruTarget]);

    // 8. (Opsional tapi Direkomendasikan) Format Ulang Tanggal di Baris 2
    const tmtIndex = headersTarget.indexOf('TMT');
    const tglInputIndex = headersTarget.indexOf('Tanggal Input');
    const updateIndex = headersTarget.indexOf('Update');
    
    if (tmtIndex !== -1) sheetTarget.getRange(2, tmtIndex + 1).setNumberFormat("dd-MM-yyyy");
    if (tglInputIndex !== -1) sheetTarget.getRange(2, tglInputIndex + 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    if (updateIndex !== -1) sheetTarget.getRange(2, updateIndex + 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");

    // --- ▲▲▲ AKHIR PERBAIKAN ▲▲▲ ---

    // 9. Hapus baris asli dari sheet "PTK PAUD"
    sheetSumber.deleteRow(rowIndex);
    
    return "Data PTK berhasil dinonaktifkan dan dipindah ke arsip.";
  } catch (e) {
    // Kirim pesan error yang spesifik ke client
    return handleError('deletePtkPaudData', e);
  }
}

function getPtkJumlahBulananSdData() {
  try {
    return getDataFromSheet('PTK_SD_JUMLAH_BULANAN');
  } catch (e) {
    return handleError('getPtkJumlahBulananSdData', e);
  }
}

function getDaftarPtkSdnData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_SD_DB;
    return SpreadsheetApp.openById(config.id).getSheetByName("PTK SDN").getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getDaftarPtkSdnData', e);
  }
}

function getDaftarPtkSdsData() {
  try {
    const config = SPREADSHEET_CONFIG.PTK_SD_DB;
    return SpreadsheetApp.openById(config.id).getSheetByName("PTK SDS").getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getDaftarPtkSdsData', e);
  }
}

function getKelolaPtkSdData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id);
    const sdnSheet = ss.getSheetByName("PTK SDN");
    const sdsSheet = ss.getSheetByName("PTK SDS");
    let combinedData = [];
    let allHeaders = []; 

    // Fungsi helper untuk mem-parsing tanggal
    const parseDate = (value) => value instanceof Date && !isNaN(value) ? value.getTime() : 0;

    // Fungsi untuk memproses satu sheet
    const processSheet = (sheet, sourceName) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      
      const range = sheet.getDataRange();
      const rawValues = range.getValues(); // Untuk tanggal
      const displayValues = range.getDisplayValues(); // Untuk string

      const headers = displayValues[0].map(h => String(h).trim());
      if (allHeaders.length === 0) allHeaders = headers; // Simpan header
      
      // ========================================================
      // === PERBAIKAN: Temukan semua indeks kolom yang kita butuhkan ===
      // ========================================================
      const updateIndex = headers.indexOf('Update');
      const inputIndex = headers.indexOf('Input');
      const namaIndex = headers.indexOf('Nama');
      const unitKerjaIndex = headers.indexOf('Unit Kerja');
      const statusIndex = headers.indexOf('Status');
      const jabatanIndex = headers.indexOf('Jabatan');
      
      // Ini adalah kunci perbaikannya:
      // Cari 'NIP' (untuk SDN) atau 'NIY' (untuk SDS)
      const nipIndex = headers.indexOf('NIP'); 
      const niyIndex = headers.indexOf('NIY');
      // ========================================================

      const dataRows = displayValues.slice(1);
      const rawRows = rawValues.slice(1);

      dataRows.forEach((row, index) => {
        // Jangan proses baris kosong
        if (!row[namaIndex]) return; 
        
        const rawRow = rawRows[index];
        const rowObject = {
          _rowIndex: index + 2,
          _source: sourceName,
        };

        // Ambil tanggal mentah untuk pengurutan
        const dateUpdate = parseDate(rawRow[updateIndex]);
        const dateInput = parseDate(rawRow[inputIndex]);
        rowObject._sortDate = Math.max(dateUpdate, dateInput); // Tanggal terbaru

        // ========================================================
        // === PERBAIKAN: Isi rowObject secara eksplisit ===
        // ========================================================
        rowObject['Nama'] = row[namaIndex] || "";
        rowObject['Unit Kerja'] = row[unitKerjaIndex] || "";
        rowObject['Status'] = row[statusIndex] || "";
        rowObject['Jabatan'] = row[jabatanIndex] || "";
        rowObject['Input'] = row[inputIndex] || "";
        rowObject['Update'] = row[updateIndex] || "";

        // Buat kolom 'NIP/NIY' gabungan
        if (nipIndex !== -1) { // Jika ini sheet SDN (ada kolom NIP)
            rowObject['NIP/NIY'] = row[nipIndex] || "";
        } else if (niyIndex !== -1) { // Jika ini sheet SDS (ada kolom NIY)
            rowObject['NIP/NIY'] = row[niyIndex] || "";
        } else {
            rowObject['NIP/NIY'] = ""; // Fallback
        }
        // ========================================================
        
        combinedData.push(rowObject);
      });
    };

    // Proses kedua sheet
    processSheet(sdnSheet, 'SDN');
    processSheet(sdsSheet, 'SDS');
    
    // Urutkan data gabungan berdasarkan tanggal terbaru
    combinedData.sort((a, b) => b._sortDate - a._sortDate);

    // Header yang diminta oleh Klien (javascript.html)
    const desiredHeaders = [
        "Nama", "Unit Kerja", "Status", "NIP/NIY", "Jabatan", "Aksi", "Input", "Update"
    ];

    return { headers: desiredHeaders, rows: combinedData };

  } catch (e) {
    return handleError('getKelolaPtkSdData', e);
  }
}

function getNewPtkSdOptions() {
  try {
    const ssOptions = SpreadsheetApp.openById(SPREADSHEET_CONFIG.FORM_OPTIONS_DB.id); // 1pr...
    const ssDropdown = SpreadsheetApp.openById(SPREADSHEET_IDS.DROPDOWN_DATA); // 1wi...

    // --- 1. Ambil Unit Kerja (dari ssDropdown / 1wi...) ---
    const sheetSDNS = ssDropdown.getSheetByName("Nama SDNS"); 
    const unitKerjaNegeri = [];
    const unitKerjaSwasta = [];

    if (sheetSDNS && sheetSDNS.getLastRow() > 1) {
        const data = sheetSDNS.getRange(2, 1, sheetSDNS.getLastRow() - 1, 2).getDisplayValues(); 
        data.forEach(row => {
            const status = row[0];
            const namaSekolah = row[1];
            if (status === "Negeri" && namaSekolah) {
                unitKerjaNegeri.push(namaSekolah);
            } else if (status === "Swasta" && namaSekolah) {
                unitKerjaSwasta.push(namaSekolah);
            }
        });
    }

    // --- 2. Helper untuk mengambil data PANGKAT dari ssOptions (1pr...) ---
    const getValuesFromOptionsDB = (sheetName, colLetter = 'A') => {
      const sheet = ssOptions.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`Peringatan: Sheet '${sheetName}' tidak ditemukan di FORM_OPTIONS_DB.`);
        return [];
      }
      return sheet.getRange(colLetter + '2:' + colLetter + sheet.getLastRow())
                  .getValues()
                  .flat()
                  .filter(value => String(value).trim() !== '');
    };
    
    return {
      'Unit Kerja Negeri': unitKerjaNegeri, // Dari 1wi...
      'Unit Kerja Swasta': unitKerjaSwasta, // Dari 1wi...
      
      // 'Agama', 'Pendidikan', 'Status' (diambil dari Klien)
      // 'Jabatan', 'Tugas Tambahan' (diambil dari Klien)
      
      // Data dinamis dari sheet 'Form SD' di 1pr...
      'Pangkat PNS': getValuesFromOptionsDB('Form SD', 'D'),
      'Pangkat PPPK': getValuesFromOptionsDB('Form SD', 'E'),
      'Pangkat PPPK PW': getValuesFromOptionsDB('Form SD', 'F'),
    };

  } catch (e) {
    return handleError('getNewPtkSdOptions', e);
  }
}

// GANTI FUNGSI 'addNewPtkSd' YANG LAMA DENGAN INI
function addNewPtkSd(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id); // 1HlyL...
    let sheet;
    
    // 1. Tentukan sheet target berdasarkan 'statusSekolah'
    if (formData.statusSekolah === 'Negeri') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (formData.statusSekolah === 'Swasta') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Status Sekolah (Negeri/Swasta) tidak valid.");
    }

    if (!sheet) throw new Error(`Sheet untuk status '${formData.statusSekolah}' tidak ditemukan.`);
    
    // 2. Ambil header dari sheet target
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    // 3. Tambahkan data meta
    formData['Input'] = new Date();
    
    // 4. Bangun baris baru berdasarkan header
    const newRow = headers.map(header => {
      let value = formData[header]; // Ambil data dari form

      // Aturan Khusus
      if (header === 'NUPTK' && value) {
        return "'" + value; // Simpan NUPTK sebagai Teks
      }
      if (header === 'TMT' && value) {
        return new Date(value); // Konversi string "yyyy-mm-dd" ke Objek Tanggal
      }
      
      // Jika data tidak ada di form (misal: 'Gol. Inpassing' di form Negeri)
      if (value === undefined) {
         return ""; // Isi dengan string kosong
      }

      return value;
    });
    
    // 5. Tambahkan baris ke sheet
    sheet.appendRow(newRow);
    
    // 6. Format ulang sel Tanggal
    const lastRow = sheet.getLastRow();
    const tmtIndex = headers.indexOf('TMT');
    if (tmtIndex !== -1) {
      sheet.getRange(lastRow, tmtIndex + 1).setNumberFormat("dd-MM-yyyy");
    }

    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

function getPtkSdDataByRow(rowIndex, source) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id);
    let sheet;
    if (source === 'SDN') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (source === 'SDS') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Sumber data tidak valid.");
    }
    if (!sheet) throw new Error(`Sheet '${source}' tidak ditemukan.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const displayValues = sheet.getRange(rowIndex, 1, 1, headers.length).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      if ((header === 'TMT' || header === 'TMT CPNS' || header === 'TMT PNS') && values[i] instanceof Date) {
        rowData[header] = Utilities.formatDate(values[i], "UTC", "yyyy-MM-dd");
      } else {
        rowData[header] = displayValues[i];
      }
    });
    return rowData;
  } catch (e) {
    return handleError('getPtkSdDataByRow', e);
  }
}

function updatePtkSdData(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id);
    let sheet;
    const source = formData.source;
    if (source === 'SDN') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (source === 'SDS') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Sumber data tidak valid.");
    }

    if (!sheet) throw new Error(`Sheet '${source}' tidak ditemukan.`);
    const rowIndex = formData.rowIndex;
    if (!rowIndex) throw new Error("Nomor baris (rowIndex) tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const oldValues = range.getValues()[0]; // Ambil nilai mentah (termasuk Date objects)

    formData['Update'] = new Date(); // Tambahkan stempel waktu update

    const newRowValues = headers.map((header, index) => {
      // Cek jika data baru ada di formData
      if (formData.hasOwnProperty(header)) {
        
        let value = formData[header];
        
        // --- PERBAIKAN LOGIKA PENYIMPANAN ---
        // 1. Ubah string tanggal "yyyy-mm-dd" kembali ke Objek Date
        if (header === 'TMT' && value) {
          return new Date(value); 
        }
        // 2. Tambahkan petik (') untuk NUPTK
        if (header === 'NUPTK' && value) {
          return "'" + value;
        }
        // 3. Jika input Pangkat/Gol untuk Non ASN (yang di-disabled)
        if (header === 'Pangkat, Gol./Ruang' && value === '— Tidak Perlu Diisi —') {
            return ""; // Simpan sebagai string kosong
        }
        
        return value; // Ambil data baru
      }
      
      // Jika tidak ada di formData (karena disabled, dll), pertahankan data lama
      return oldValues[index]; 
    });
    
    range.setValues([newRowValues]); // 1. Simpan datanya

    // 2. Format Ulang Tanggal
    const tmtIndex = headers.indexOf('TMT');
    if (tmtIndex !== -1) {
      sheet.getRange(rowIndex, tmtIndex + 1).setNumberFormat("dd-MM-yyyy");
    }

    return "Data PTK berhasil diperbarui.";
  } catch (e) {
    throw new Error(`Gagal memperbarui data: ${e.message}`);
  }
}

function getKebutuhanPtkSdnData() {
  try {
    return getDataFromSheet('PTK_SD_KEBUTUHAN');
  } catch (e) {
    return handleError('getKebutuhanPtkSdnData', e);
  }
}

function deletePtkSdData(rowIndex, source, deleteCode) {
  try {
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");

    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.PTK_SD_DB.id);
    let sheet;
    if (source === 'SDN') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (source === 'SDS') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Sumber data tidak valid: " + source);
    }

    if (!sheet) throw new Error("Sheet sumber '" + source + "' tidak ditemukan.");
    
    const maxRows = sheet.getLastRow();
    if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > maxRows) throw new Error("Nomor baris tidak valid.");

    sheet.deleteRow(rowIndex);
    return "Data PTK berhasil dihapus.";
  } catch (e) {
    throw new Error(`Gagal menghapus data: ${e.message}`);
  }
}


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