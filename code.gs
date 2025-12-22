/**
 * ===================================================================
 * KONFIGURASI TERPUSAT (PASTE ID ANDA DI SINI)
 * ===================================================================
 */

// 1. ID SPREADSHEET UMUM & SK/LAPBUL
const SPREADSHEET_CONFIG = {
  // Database Login User
  ID_UTAMA: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA", 
  SHEET_USERS: "Data User", 

  // Data Master (Sekolah, Tahun, dll)
  // ID ini saya ambil dari file Lapbul.gs Anda, sepertinya sudah benar
  DROPDOWN_DATA: { id: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA", sheet: "DropdownData" },

  // Modul SK
  SK_FORM_RESPONSES: { id: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4", sheet: "SK Pembagian Tugas" },
  SK_BAGI_TUGAS:     { id: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4", sheet: "SK Pembagian Tugas" }, // Biasanya sama dgn atas

  // Modul Laporan Bulanan (Respon Form)
  LAPBUL_FORM_RESPONSES_PAUD: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Input PAUD" },
  LAPBUL_FORM_RESPONSES_SD:   { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Input SD" },

  // Modul Lapbul (Tabel Riwayat & Status - Biasanya IDnya sama dengan Respon Form)
  LAPBUL_RIWAYAT: {
    PAUD: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Input PAUD" },
    SD:   { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Input SD" }
  },
  LAPBUL_STATUS: {
    PAUD: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Status PAUD" },
    SD:   { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Status SD" }
  },
  PTK_PAUD_KEADAAN: { 
    id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", 
    sheet: "Keadaan PTK PAUD" 
  },
  PTK_PAUD_JUMLAH_BULANAN: { 
    id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs",
    sheet: "Jumlah PTK Bulanan"
  },
  PTK_PAUD_DB: { 
    id: "1XetGkBymmN2NZQlXpzZ2MQyG0nhhZ0sXEPcNsLffhEU", 
    sheet: "PTK PAUD" 
  },
  // [BARU] Tambahan Config SD
  PTK_SD_DB: { 
    id: "1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0", 
    sheet: "PTK SD" 
  },
    PTK_PAUD_MUTASI: { 
    id: "1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0", 
    sheet: "PTK PAUD MUTASI" 
  }
};

// 2. ID SPREADSHEET KHUSUS SIABA (Wajib ada variabel bernama SPREADSHEET_IDS)
const SPREADSHEET_IDS = {
  SIABA_DB: "1sfbvyIZurU04gictep8hI-NnvicGs0wrDqANssVXt6o",           // Database Lupa Presensi & Pegawai
  SIABA_SALAH_DB: "1TZGrMiTuyvh2Xbo44RhJuWlQnOC5LzClsgIoNKtRFkY",        // Database Salah Presensi
  SIABA_DINAS_DB: "1I_2yUFGXnBJTCSW6oaT3D482YCs8TIRkKgQVBbvpa1M",        // Database Perjalanan Dinas
  SIABA_CUTI_DB: "1DhBjmLHFMuJqWM6yJHsm-1EKvHzG8U4zK2GuU-dIgn8",          // Database Cuti
  SIABA_REKAP_HELPER: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA"    // Helper Rekap Terlambat/Pulang Awal
};

// 3. ID FOLDER GOOGLE DRIVE (Tempat simpan PDF)
const FOLDER_CONFIG = {
  // Folder SK
  MAIN_SK: "1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs",

  // Folder Lapbul
  LAPBUL_KB: "18CxRT-eledBGRtHW1lFd2AZ8Bub6q5ra",
  LAPBUL_TK: "1WUNz_BSFmcwRVlrG67D2afm9oJ-bVI9H",
  LAPBUL_SD: "1I8DRQYpBbTt1mJwtD1WXVD6UK51TC8El",
  TRASH_SD: "1MpEgpCDrTX-SHjdNIa3aUpKUyYZpejrb",
  TRASH_PAUD: "1EUIOthRbotJQlSphxVZ-QAdewe17UCOU",

  // Folder Siaba
  SIABA_LUPA: "10kwGuGfwO5uFreEt7zBJZUaDx1fUSXo9",
  SIABA_DINAS: "1uPeOU7F_mgjZVyOLSsj-3LXGdq9rmmWl",
  SIABA_CUTI_DOCS: "1fAmqJXpmGIfEHoUeVm4LjnWvnwVwOfNM",
  SIABA_REKAP_ARCHIVE: "1MoGuseJNrOIMnkZNoqkKcK282jZpUkAm",
  SIABA_SKP_DOCS: "1DGYC8AtJFCpCZ0ou2ae9-5fc2-bWl20G",
  SIABA_PAK_DOCS: "1cvn-pOufs-OIbFQfqhmxc3fcmFuox4Sc",
};

/**
 * ===================================================================
 * FUNGSI UTAMA (ROUTER & LOGIN)
 * ===================================================================
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Media Informasi Korwil Disdikbud Secang')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  try {
    // Gunakan createTemplateFromFile agar lebih fleksibel
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    // JIKA ERROR (File tidak ketemu), KEMBALIKAN PESAN MERAH (BUKAN BLANK)
    return '<div style="color:red; padding:20px; text-align:center; border:1px solid red; margin:10px;">' +
           '<h3>Gagal Memuat Halaman</h3>' +
           '<p>File <b>"' + filename + '.html"</b> tidak ditemukan di Editor Script.</p>' +
           '<p>Cek nama file di sidebar kiri Anda.</p>' +
           '<small>Error detail: ' + e.message + '</small>' +
           '</div>';
  }
}

function getPage(page) {
  if (page === 'ptk_kelola_paud' || page === 'ptk_kelola_sd') {
      page = 'page_ptk_kelola'; 
  }
  // Panggil include yang sudah Anda punya, atau createTemplate langsung
  return HtmlService.createTemplateFromFile(page).evaluate().getContent();
}

function loginUser(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.ID_UTAMA);
  const sheet = ss.getSheetByName('Data User');
  
  if (!sheet) return { status: 'error', message: 'Sheet Data User tidak ditemukan.' };
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Cek Username (A) & Password (B)
    if (String(row[0]) === String(username) && String(row[1]) === String(password)) {
      
      return {
        status: 'success',
        user: {
          username: row[0], // Kolom A
          
          // --- KOREKSI MAPPING ---
          nama: row[2],     // Kolom C = Nama Lengkap
          role: row[3],     // Kolom D = Role (Admin/User)
          foto: row[4]      // Kolom E = Foto
          // -----------------------
        }
      };
    }
  }
  
  return { status: 'error', message: 'Username atau Password salah.' };
}

// --- Helper Functions ---

function handleError(functionName, error) {
  console.error(`Error di ${functionName}: ${error.message}`);
  return { error: error.message };
}

function getCachedData(key, fetchFunction, expirationInSeconds = 3600) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);
  const data = fetchFunction();
  if (data && !data.error) cache.put(key, JSON.stringify(data), expirationInSeconds);
  return data;
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

// Pengaman jika form submit secara native (mencegah blank)
function doPost(e) {
  return doGet(e);
}

function getDataFromSheet(configKey) {
  const conf = SPREADSHEET_CONFIG[configKey];
  if (!conf) throw new Error("Konfigurasi sheet tidak ditemukan: " + configKey);
  
  const ss = SpreadsheetApp.openById(conf.id);
  const sheet = ss.getSheetByName(conf.sheet);
  if (!sheet) throw new Error("Sheet tidak ditemukan: " + conf.sheet);
  
  // Ambil data sebagai String (Display Values) agar format angka/tanggal aman
  return sheet.getDataRange().getDisplayValues();
}