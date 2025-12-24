const LAPBUL_COLUMNS = {
    'PAUD': {
        // Kolom Utama
        'Tanggal Unggah': 0, // Kolom A
        'Bulan': 1,          // Kolom B
        'Tahun': 2,          // Kolom C
        'Status': 4,         // Kolom E
        'Rombel': 5,         // Kolom F
        'Jenjang': 6,        // Kolom G
        'Nama Sekolah': 7,   // Kolom H
        'Dokumen': 36,       // Kolom AK (Cek lagi di sheet Anda)
        
        // --- PERBAIKAN DI SINI (SESUAIKAN ANGKA-NYA) ---
        'User': 38,    // <--- Ganti dengan nomor kolom "User" yang benar
        'Update': 37   // <--- Ganti dengan nomor kolom "Tanggal Edit/Update"
    },
    'SD': {
        // Kolom Utama
        'Tanggal Unggah': 0, // Kolom A
        'Bulan': 1,          // Kolom B
        'Tahun': 2,          // Kolom C
        'Status': 3,         // Kolom D
        'Nama Sekolah': 4,   // Kolom E
        'Rombel': 6,         // Kolom G
        'Dokumen': 7,        // Kolom H
        'Jenjang': 217,      // Kolom HP (Cek lagi)
        
        // --- PERBAIKAN DI SINI (SESUAIKAN ANGKA-NYA) ---
        'User': 219,   // <--- Ganti dengan nomor kolom "User" yang benar
        'Update': 218,  // <--- Ganti dengan nomor kolom "Tanggal Edit/Update"
        'User Edit':220
    }
};

function getPaudSchoolLists() {
  const cacheKey = 'paud_school_lists_final_v4'; // Kunci cache baru
  return getCachedData(cacheKey, function() {
    
    // 1. Menggunakan kunci DROPDOWN_DATA (sesuai ID: 1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA)
    const config = SPREADSHEET_CONFIG.DROPDOWN_DATA; 
    
    if (!config || !config.id) {
        throw new Error("Konfigurasi ID Spreadsheet (DROPDOWN_DATA) tidak ditemukan.");
    }

    const ss = SpreadsheetApp.openById(config.id);
    if (!ss) {
        throw new Error("Gagal membuka Spreadsheet. Periksa ID atau izin akses.");
    }
    
    // 2. Gunakan nama Sheet yang benar: Form PAUD
    const sheet = ss.getSheetByName('Form PAUD');
    if (!sheet) {
        throw new Error("Sheet 'Form PAUD' tidak ditemukan di Spreadsheet Dropdown Data.");
    }
    
    if (sheet.getLastRow() < 2) {
        return { "KB": [], "TK": [] };
    }

    // 3. Ambil data Jenjang (Kolom A) dan Nama Lembaga (Kolom B)
    // Ambil data mulai dari baris 2, kolom 1 (A) sepanjang 2 kolom (A & B)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
    
    const lists = { "KB": [], "TK": [] };

    data.forEach(row => {
        const jenjang = String(row[0]).trim();
        const namaLembaga = String(row[1]).trim();
        
        if (jenjang && namaLembaga) {
            // Hanya tambahkan ke array yang sudah didefinisikan (KB atau TK)
            if (lists.hasOwnProperty(jenjang) && !lists[jenjang].includes(namaLembaga)) {
                lists[jenjang].push(namaLembaga);
            }
        }
    });
    
    // Urutkan Nama Lembaga
    lists["KB"].sort();
    lists["TK"].sort();

    return lists;
  });
}

function getSdSchoolLists() {
  try {
    // 1. Buka spreadsheet yang benar menggunakan DROPDOWN_DATA
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
    // 2. Akses sheet "Nama SDNS"
    const sheet = ss.getSheetByName("Nama SDNS");

    if (!sheet || sheet.getLastRow() < 2) {
      return { SDN: [], SDS: [] }; // Kembalikan array kosong jika sheet tidak ada/kosong
    }

    // 3. Ambil semua data dari baris kedua sampai akhir
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();

    const lists = {
      "SDN": [],
      "SDS": []
    };

    // 4. Loop melalui data dan pisahkan berdasarkan status di kolom A
    data.forEach(row => {
      const status = row[0]; // Kolom A (Negeri/Swasta)
      const namaSekolah = row[1]; // Kolom B (Nama SD)

      if (status === "Negeri" && namaSekolah) {
        lists.SDN.push(namaSekolah);
      } else if (status === "Swasta" && namaSekolah) {
        lists.SDS.push(namaSekolah);
      }
    });

    // Urutkan hasilnya
    lists.SDN.sort();
    lists.SDS.sort();

    return lists;

  } catch (e) {
    return handleError('getSdSchoolLists', e);
  }
}

function processLapbulFormPaud(formData) {
  try {
    const jenjang = formData.jenjang;
    let FOLDER_ID_LAPBUL;
    if (jenjang === 'KB') {
      FOLDER_ID_LAPBUL = FOLDER_CONFIG.LAPBUL_KB;
    } else if (jenjang === 'TK') {
      FOLDER_ID_LAPBUL = FOLDER_CONFIG.LAPBUL_TK;
    } else {
      throw new Error("Jenjang tidak valid: " + jenjang);
    }
    
    const mainFolder = DriveApp.getFolderById(FOLDER_ID_LAPBUL);
    const tahunFolder = getOrCreateFolder(mainFolder, formData.tahun);
    const bulanFolder = getOrCreateFolder(tahunFolder, formData.laporanBulan);

    const fileData = formData.fileData;
    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    const newFileName = `${formData.namaSekolah} - Lapbul ${formData.laporanBulan} ${formData.tahun}.pdf`;
    const newFile = bulanFolder.createFile(blob).setName(newFileName);
    const fileUrl = newFile.getUrl();

    const config = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);

    const newRow = [
      new Date(),
      formData.laporanBulan, formData.tahun, formData.npsn, formData.statusSekolah, formData.jumlahRombel, formData.jenjang, formData.namaSekolah,
      formData.murid_0_1_L, formData.murid_0_1_P, formData.murid_1_2_L, formData.murid_1_2_P,
      formData.murid_2_3_L, formData.murid_2_3_P, formData.murid_3_4_L, formData.murid_3_4_P,
      formData.murid_4_5_L, formData.murid_4_5_P, formData.murid_5_6_L, formData.murid_5_6_P,
      formData.murid_6_up_L, formData.murid_6_up_P, formData.kelompok_A_L, formData.kelompok_A_P,
      formData.kelompok_B_L, formData.kelompok_B_P, formData.kepsek_ASN, formData.kepsek_Non_ASN,
      formData.guru_PNS, formData.guru_PPPK, formData.guru_GTY, formData.guru_GTT,
      formData.tendik_Penjaga, formData.tendik_TAS, formData.tendik_Pustakawan, formData.tendik_Lainnya,
      fileUrl, formData.User
    ];
    sheet.appendRow(newRow);

    return "Sukses! Laporan Bulan PAUD berhasil dikirim.";
  } catch (e) {  
    return handleError('processLapbulFormPaud', e);
  }
}

// --- FUNGSI SIMPAN SD (AUTO MAPPING BERDASARKAN HEADER) ---
function processLapbulFormSd(formData) {
  try {
    const config = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error("Sheet Input SD tidak ditemukan.");

    // 1. SETUP FOLDER
    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.LAPBUL_SD);
    const tahunFolder = getOrCreateFolder(mainFolder, formData.Tahun);
    const bulanFolder = getOrCreateFolder(tahunFolder, formData.Bulan);
    
    // 2. PROSES FILE
    let fileUrl = "";
    if (formData.fileData && formData.fileData.data) {
        const decoded = Utilities.base64Decode(formData.fileData.data);
        const blob = Utilities.newBlob(decoded, formData.fileData.mimeType, formData.fileData.fileName);
        const newName = `${formData['Nama Sekolah']} - Lapbul SD ${formData.Bulan} ${formData.Tahun}.pdf`;
        const file = bulanFolder.createFile(blob).setName(newName);
        fileUrl = file.getUrl();
    }

    // 3. AUTO MAPPING (Input HTML -> Header Spreadsheet)
    // Ambil baris header (Baris 1)
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Siapkan baris kosong
    let newRow = new Array(headers.length).fill("");

    // Loop setiap kolom header untuk mencari pasangannya di formData
    headers.forEach((h, index) => {
        const header = String(h).trim();
        
        // A. Metadata Otomatis
        if (header === 'Tanggal Unggah') newRow[index] = new Date();
        else if (header === 'Dokumen') newRow[index] = fileUrl;
        else if (header === 'Update') newRow[index] = "";
        else if (header === 'Jenjang') newRow[index] = "SD";
        else if (header === 'User' || header === 'Pengunggah') { 
            newRow[index] = formData.User; 
        }
        
        // B. Data dari Form
        else if (formData.hasOwnProperty(header)) {
            let val = formData[header];
            // Jika angka kosong, simpan sebagai 0 agar bisa dihitung spreadsheet
            if (val === "" && !isNaN(Number(val))) newRow[index] = 0; 
            else newRow[index] = val;
        }
    });

    // 4. SIMPAN
    sheet.appendRow(newRow);
    return "Sukses! Laporan Bulan SD berhasil disimpan.";

  } catch (e) {
    throw new Error("Gagal menyimpan data SD: " + e.message);
  }
}

/**
 * FUNGSI PENGAMBIL DATA RIWAYAT (PAUD & SD) - VERSI ROBUST
 * Otomatis mencari kolom berdasarkan Header, meminimalisir error 'Data Kosong'.
 */
function getLapbulRiwayatData() {
  const result = [];
  
  // --- 1. AMBIL DATA PAUD ---
  try {
    const sheetPaud = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_RIWAYAT.PAUD.id)
                                    .getSheetByName(SPREADSHEET_CONFIG.LAPBUL_RIWAYAT.PAUD.sheet);
    
    // Ambil semua data sebagai String (DisplayValues) agar format tanggal/angka aman
    const dataPaud = sheetPaud.getDataRange().getDisplayValues();
    
    if (dataPaud.length > 1) {
      const headers = dataPaud[0].map(h => String(h).toLowerCase().trim());
      
      // Cari Index Kolom Penting (Flexible)
      // Kita cari posisi kolom berdasarkan kata kunci di header
      const idx = {
        Time: headers.findIndex(h => h.includes('tanggal') || h.includes('timestamp') || h.includes('waktu')),
        Bulan: headers.indexOf('bulan'),
        Tahun: headers.indexOf('tahun'),
        Nama: headers.findIndex(h => h.includes('nama') && (h.includes('sekolah') || h.includes('lembaga'))),
        Status: headers.indexOf('status'),
        Rombel: headers.findIndex(h => h.includes('rombel') || h.includes('kelompok')),
        File: headers.findIndex(h => h.includes('dokumen') || h.includes('file') || h.includes('link')),
        User: headers.findIndex(h => h === 'user' || h === 'pengunggah')
      };

      // Loop data (mulai baris ke-2)
      for (let i = 1; i < dataPaud.length; i++) {
        const row = dataPaud[i];
        
        // Ambil User: Jika kolom user ketemu pake itu, kalo nggak pake kolom terakhir
        let userVal = (idx.User > -1) ? row[idx.User] : row[row.length - 1];
        
        // Validasi User: Jangan sampai URL file dianggap user
        if (String(userVal).includes('http') || String(userVal).length > 50) userVal = '-';

        result.push({
          "Tanggal Unggah": (idx.Time > -1) ? row[idx.Time] : row[0],
          "Jenjang": "PAUD",
          "Nama Sekolah": (idx.Nama > -1) ? row[idx.Nama] : "Tanpa Nama",
          "Status": (idx.Status > -1) ? row[idx.Status] : "Terkirim",
          "Bulan": (idx.Bulan > -1) ? row[idx.Bulan] : "-",
          "Tahun": (idx.Tahun > -1) ? row[idx.Tahun] : "-",
          "Rombel": (idx.Rombel > -1) ? row[idx.Rombel] : "-",
          "Dokumen": (idx.File > -1) ? row[idx.File] : "",
          "User": userVal || '-',
          "_rowIndex": i + 1,
          "_source": "PAUD"
        });
      }
    }
  } catch (e) {
    console.error("Error Get Riwayat PAUD: " + e.message);
  }

  // --- 2. AMBIL DATA SD ---
  try {
    const sheetSd = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_RIWAYAT.SD.id)
                                  .getSheetByName(SPREADSHEET_CONFIG.LAPBUL_RIWAYAT.SD.sheet);
    const dataSd = sheetSd.getDataRange().getDisplayValues();
    
    if (dataSd.length > 1) {
      const headers = dataSd[0].map(h => String(h).toLowerCase().trim());
      
      // Cari Index Kolom (SD)
      const idx = {
        Time: headers.findIndex(h => h.includes('tanggal') || h.includes('timestamp')),
        Bulan: headers.indexOf('bulan'),
        Tahun: headers.indexOf('tahun'),
        Nama: headers.findIndex(h => h.includes('nama') && h.includes('sekolah')),
        // SD mungkin headernya "Jumlah Rombel" atau "Rombel"
        Rombel: headers.findIndex(h => h.includes('rombel')), 
        File: headers.findIndex(h => h.includes('dokumen') || h.includes('file')),
        Status: headers.indexOf('status'), // Jika ada
        User: headers.findIndex(h => h === 'user' || h === 'pengunggah')
      };

      for (let i = 1; i < dataSd.length; i++) {
        const row = dataSd[i];
        let userVal = (idx.User > -1) ? row[idx.User] : row[row.length - 1];
        if (String(userVal).includes('http')) userVal = '-';

        result.push({
          "Tanggal Unggah": (idx.Time > -1) ? row[idx.Time] : row[0],
          "Jenjang": "SD",
          "Nama Sekolah": (idx.Nama > -1) ? row[idx.Nama] : "Tanpa Nama",
          "Status": (idx.Status > -1) ? row[idx.Status] : "Terkirim",
          "Bulan": (idx.Bulan > -1) ? row[idx.Bulan] : "-",
          "Tahun": (idx.Tahun > -1) ? row[idx.Tahun] : "-",
          "Rombel": (idx.Rombel > -1) ? row[idx.Rombel] : "-",
          "Dokumen": (idx.File > -1) ? row[idx.File] : "",
          "User": userVal || '-',
          "_rowIndex": i + 1,
          "_source": "SD"
        });
      }
    }
  } catch (e) {
    console.error("Error Get Riwayat SD: " + e.message);
  }

  // --- 3. SORTING (TERBARU DI ATAS) ---
  // Kita parsing tanggal manual agar akurat
  result.sort((a, b) => {
    // Helper parse dd/MM/yyyy HH:mm:ss
    const parse = (str) => {
      if(!str) return 0;
      // Coba parsing standar
      let d = new Date(str);
      if(!isNaN(d)) return d.getTime();
      
      // Coba parsing manual format Indonesia (dd/mm/yyyy)
      let parts = String(str).split(/[/\s:]/); // Split by / space :
      if(parts.length >= 3) {
         // asumsi [dd, mm, yyyy, HH, MM, SS]
         return new Date(parts[2], parts[1]-1, parts[0], parts[3]||0, parts[4]||0, parts[5]||0).getTime();
      }
      return 0;
    };
    return parse(b["Tanggal Unggah"]) - parse(a["Tanggal Unggah"]);
  });

  return result;
}

function _getLimitedSheetData(sheetId, sheetName, startRow, numColumns) {
    try {
        const ss = SpreadsheetApp.openById(sheetId);
        if (!ss) return [];
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() < startRow) return [];
        
        // Ambil data dari baris tertentu hingga akhir, sebanyak numColumns (16 kolom = A sampai P)
        if (sheet.getLastRow() < startRow) return [];
        
        const dataRange = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, numColumns); 
        return dataRange.getDisplayValues();
    } catch (e) {
        Logger.log(`Error accessing data for ID ${sheetId}, sheet ${sheetName}: ${e.message}`);
        return [];
    }
}

const getSheetDataSecure = (sheet) => {
    // Jika hanya ada header atau sheet kosong, kembalikan null
    if (!sheet || sheet.getLastRow() < 2) return null;
    
    // Panggilan RPC (getValues dan getDisplayValues) hanya dilakukan di sini.
    const range = sheet.getDataRange();
    const displayValues = range.getDisplayValues(); 
    const rawValues = range.getValues(); // Mengambil Date Object mentah

    return {
        display: displayValues, 
        raw: rawValues,
        // Membersihkan header dari spasi ekstra (.trim()) untuk mapping yang akurat
        headers: displayValues[0].map(h => String(h).trim()) 
    };
};

const getCleanedDisplayValue = (displayRow, keyIndex) => {
    // Jika index tidak valid, kembalikan null
    if (keyIndex < 0) return null;
    // Mengambil nilai, mengubahnya ke String, dan menghapus spasi di awal/akhir (.trim())
    return String(displayRow[keyIndex] || '').trim();
};

/**
 * FUNGSI STATUS PENGIRIMAN (STANDARD TABLE VIEW)
 * Mengembalikan data JSON daftar sekolah beserta status per bulannya.
 */
function getLapbulStatusData(filter) {
  const tahun = (filter && filter.tahun) ? String(filter.tahun) : new Date().getFullYear().toString();
  const jenjangFilter = (filter && filter.jenjang) ? filter.jenjang : ""; 

  // Daftar Bulan untuk Iterasi
  const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  
  // Header Final yang akan dikirim ke Frontend
  const FINAL_HEADERS = ["Nama Sekolah", "Jenjang", "Status", ...months];

  const map = {}; // Object untuk menampung data unik per sekolah

  // Helper Processor
  const processSheet = (config, labelJenjang) => {
    try {
      const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
      const data = sheet.getDataRange().getDisplayValues();
      if (data.length < 2) return;

      const headers = data[0].map(h => String(h).toLowerCase().trim());
      
      // Cari Index Kolom
      const idx = {
        Nama: headers.findIndex(h => h.includes('nama') && (h.includes('sekolah') || h.includes('lembaga'))),
        Tahun: headers.indexOf('tahun'),
        Status: headers.findIndex(h => h === 'status'), // Kolom Status (OK/Revisi)
      };
      
      // Cari Index Kolom Bulan Secara Dinamis
      const monthIndices = {};
      months.forEach(m => {
          monthIndices[m] = headers.indexOf(m.toLowerCase());
      });

      if (idx.Nama === -1 || idx.Tahun === -1) return;

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (String(row[idx.Tahun]).trim() !== tahun) continue;

        const nama = String(row[idx.Nama]).trim();
        if(!nama) continue;

        if (!map[nama]) {
          map[nama] = { 
            "Nama Sekolah": nama,
            "Jenjang": labelJenjang, 
            "Status": "-", 
            // Inisialisasi bulan kosong
            ...months.reduce((acc, m) => ({ ...acc, [m]: "" }), {}) 
          };
        }

        // Update Status (Overwrite dengan data terbaru di baris bawah)
        if (idx.Status > -1 && row[idx.Status]) {
            map[nama]["Status"] = row[idx.Status];
        }

        // Cek Data Bulan
        months.forEach(m => {
            const colIdx = monthIndices[m];
            if (colIdx > -1) {
                const cellVal = row[colIdx];
                // Jika sel ada isinya (apapun itu: tanggal, v, checklist), anggap sudah setor
                if (cellVal && cellVal.trim() !== "" && cellVal.trim() !== "-") {
                    map[nama][m] = "✅"; // Standarkan jadi centang
                }
            }
        });
      }
    } catch (e) {
      console.error("Error Status " + labelJenjang + ": " + e.message);
    }
  };

  // Ambil Data Sesuai Filter
  if (jenjangFilter === "" || jenjangFilter === "PAUD") processSheet(SPREADSHEET_CONFIG.LAPBUL_STATUS.PAUD, "PAUD");
  if (jenjangFilter === "" || jenjangFilter === "SD")   processSheet(SPREADSHEET_CONFIG.LAPBUL_STATUS.SD, "SD");

  // Konversi Map ke Array Object
  const rows = Object.values(map).sort((a, b) => a["Nama Sekolah"].localeCompare(b["Nama Sekolah"]));

  // Return Format Standar SmartTable
  return { 
      headers: FINAL_HEADERS, 
      rows: rows 
  };
}

/**
 * FUNGSI STATUS PENGIRIMAN (INDEX POSITION FIX)
 * Mengambil data berdasarkan urutan kolom pasti yang diinfokan user:
 * [0] Nama, [1] Jenjang, [2] Status, [3] Tahun, [4-15] Bulan Jan-Des
 */
function getLapbulStatusData(filter) {
  // 1. Ambil Filter (Default Tahun Ini)
  const filterTahun = (filter && filter.tahun) ? String(filter.tahun) : new Date().getFullYear().toString();
  const filterJenjang = (filter && filter.jenjang) ? filter.jenjang : ""; 

  // Daftar Bulan untuk Key JSON
  const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  
  const map = {}; 

  // Helper Processor
  const processSheet = (spreadsheetId, sheetName, defaultJenjang) => {
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const sheet = ss.getSheetByName(sheetName);
      
      if (!sheet) {
          console.warn(`Sheet '${sheetName}' tidak ditemukan.`);
          return;
      }

      // Ambil Semua Data (Display Values agar format sesuai tampilan Excel)
      const data = sheet.getDataRange().getDisplayValues();
      if (data.length < 2) return;

      // --- SETUP INDEX KOLOM (HARDCODE SESUAI STRUKTUR ANDA) ---
      // Karena struktur sudah pasti, kita kunci index-nya
      const idx = {
        Nama: 0,    // Kolom A
        Jenjang: 1, // Kolom B
        Status: 2,  // Kolom C
        Tahun: 3,   // Kolom D
        BulanStart: 4 // Januari mulai dari Kolom E
      };

      // Loop Data (Mulai Baris 2)
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Cek apakah baris punya cukup kolom
        if (row.length < 4) continue;

        // 1. FILTER TAHUN
        // Pastikan kolom Tahun (Index 3) sama dengan filter
        const rowTahun = String(row[idx.Tahun]).trim();
        if (rowTahun !== filterTahun) continue;

        // 2. AMBIL NAMA SEKOLAH
        const nama = String(row[idx.Nama]).trim();
        if (!nama) continue;

        // 3. INISIALISASI DATA SEKOLAH
        if (!map[nama]) {
          map[nama] = { 
            "Nama Sekolah": nama,
            "Jenjang": (row[idx.Jenjang] || defaultJenjang),
            "Status": (row[idx.Status] || "-"),
            // Default bulan kosong dulu
            ...months.reduce((acc, m) => ({ ...acc, [m]: "" }), {}) 
          };
        }

        // 4. LOOP DATA BULAN (Index 4 s.d 15)
        // Kita ambil 12 kolom setelah kolom Tahun
        for (let m = 0; m < 12; m++) {
            const colIndex = idx.BulanStart + m;
            
            // Pastikan kolom ada di data row
            if (colIndex < row.length) {
                const cellVal = row[colIndex];
                const monthName = months[m]; // Januari, Februari...

                // LOGIKA CEK:
                // Jika sel TIDAK KOSONG dan TIDAK STRIP, anggap SUDAH (✅)
                if (cellVal && cellVal.trim().length > 0 && cellVal.trim() !== '-') {
                    map[nama][monthName] = "✅"; 
                }
            }
        }
      }
    } catch (e) {
      console.error("Error Processing " + sheetName + ": " + e.message);
    }
  };

  // --- EKSEKUSI (Hardcode ID & Nama Sheet agar Aman) ---
  // Pastikan ID di Code.gs benar. Nama Sheet "Status PAUD" & "Status SD" (Case Sensitive)
  
  if (filterJenjang === "" || filterJenjang === "PAUD") {
      processSheet(SPREADSHEET_CONFIG.LAPBUL_STATUS.PAUD.id, "Status PAUD", "PAUD");
  }
  if (filterJenjang === "" || filterJenjang === "SD") {
      processSheet(SPREADSHEET_CONFIG.LAPBUL_STATUS.SD.id, "Status SD", "SD");
  }

  // Sort & Return
  const rows = Object.values(map).sort((a, b) => a["Nama Sekolah"].localeCompare(b["Nama Sekolah"]));
  const finalHeaders = ["Nama Sekolah", "Jenjang", "Status", ...months];

  return { headers: finalHeaders, rows: rows };
}

const parseDateForSort = (dateStr) => {
    // KUNCI PERBAIKAN: Cek jika dateStr kosong (null, undefined, atau string kosong)
    if (!dateStr) return 0;
    
    if (dateStr instanceof Date && !isNaN(dateStr)) return dateStr.getTime();
    
    // Hanya jalankan .match jika dateStr adalah string
    if (typeof dateStr === 'string') {
        const parts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
        if (parts) {
            return new Date(parts[3], parts[2] - 1, parts[1], parts[4], parts[5], parts[6]).getTime();
        }
        const dateOnlyParts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        if (dateOnlyParts) {
            return new Date(dateOnlyParts[3], dateOnlyParts[2] - 1, dateOnlyParts[1]).getTime();
        }
    }
    return 0;
};

// GANTI FUNGSI HELPER INI (baris 86)
const formatDate = (cell) => {
    // KUNCI PERBAIKAN: Cek jika cell adalah Date object yang valid
    if (cell instanceof Date && !isNaN(cell)) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    }
    // Jika tidak valid (null, "", dll.), kembalikan string '-'
    return cell || '-'; 
};

// --- 1. UPDATE FUNGSI MAPPING (Support Kolom User & Sorting Cerdas) ---
const mapSheet = (sheet, source, timeZone) => {
    const sheetData = getSheetDataSecure(sheet);
    if (!sheetData) return [];

    const { display, raw, headers } = sheetData; // headers diambil dari Baris 1
    
    // Cari Index Kolom Secara Dinamis (Lebih Aman)
    const getIdx = (name) => headers.indexOf(name);
    
    const idxNama = getIdx("Nama Sekolah");
    const idxJenjang = getIdx("Jenjang");
    const idxStatus = getIdx("Status");
    const idxBulan = getIdx("Bulan");
    const idxTahun = getIdx("Tahun");
    const idxDokumen = getIdx("Dokumen");
    const idxUnggah = getIdx("Tanggal Unggah");
    const idxUpdate = getIdx("Update"); // Kolom Update
    
    // [PERUBAHAN 2] Cari Index Kolom User
    const idxUser = getIdx("User");           // Kolom Pengunggah
    const idxUserEdit = getIdx("User Edit");  // Kolom Penyunting Baru

    if (idxNama === -1 || idxUnggah === -1) return [];

    return display.slice(1).map((displayRow, index) => { 
        const rowIndex = index + 2; 
        const dateUnggahRaw = raw[rowIndex - 1][idxUnggah]; 
        const dateUpdateRaw = idxUpdate !== -1 ? raw[rowIndex - 1][idxUpdate] : null;

        if (!dateUnggahRaw) return null; 
        
        const dateUnggah = (dateUnggahRaw instanceof Date && !isNaN(dateUnggahRaw)) ? dateUnggahRaw.getTime() : 0;
        const dateUpdate = (dateUpdateRaw instanceof Date && !isNaN(dateUpdateRaw)) ? dateUpdateRaw.getTime() : 0;
        
        return {
            _rowIndex: rowIndex,
            _source: source,
            _sortDate: Math.max(dateUnggah, dateUpdate),
            
            "Nama Sekolah": getCleanedDisplayValue(displayRow, idxNama) || '-',
            "Jenjang": idxJenjang !== -1 ? (getCleanedDisplayValue(displayRow, idxJenjang) || (source==='SD'?'SD':'PAUD')) : (source==='SD'?'SD':'PAUD'),
            "Status": getCleanedDisplayValue(displayRow, idxStatus) || '-',
            "Bulan": getCleanedDisplayValue(displayRow, idxBulan) || '-',
            "Tahun": getCleanedDisplayValue(displayRow, idxTahun) || '-',
            "Dokumen": getCleanedDisplayValue(displayRow, idxDokumen) || '',
            
            "Tanggal Unggah": (dateUnggahRaw instanceof Date) ? Utilities.formatDate(dateUnggahRaw, timeZone, "dd/MM/yyyy HH:mm:ss") : '-',
            "Update": (dateUpdateRaw instanceof Date) ? Utilities.formatDate(dateUpdateRaw, timeZone, "dd/MM/yyyy HH:mm:ss") : '-',
            
            // [PERUBAHAN 3] Masukkan Data User
            "User": getCleanedDisplayValue(displayRow, idxUser) || '-',
            "User Edit": getCleanedDisplayValue(displayRow, idxUserEdit) || '-' 
        };
        
    }).filter(row => row !== null && row["Nama Sekolah"] !== '-'); 
};

// --- 2. UPDATE FUNGSI UTAMA (Ubah Urutan Header) ---
function getLapbulKelolaData() {
    try {
        const PAUD_CONFIG = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD;
        const SD_CONFIG = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD;
        
        const PAUD_SS = SpreadsheetApp.openById(PAUD_CONFIG.id);
        const SD_SS = SpreadsheetApp.openById(SD_CONFIG.id);
        
        const PAUD_SHEET = PAUD_SS.getSheetByName(PAUD_CONFIG.sheet);
        const SD_SHEET = SD_SS.getSheetByName(SD_CONFIG.sheet);
        
        const PAUD_TIME_ZONE = PAUD_SS.getSpreadsheetTimeZone();
        const SD_TIME_ZONE = SD_SS.getSpreadsheetTimeZone();
        
        // [PERUBAHAN 1] Tambahkan Header 'User' dan 'User Edit'
        const FINAL_HEADERS = [
            "Nama Sekolah", "Jenjang", "Status", "Bulan", "Tahun", 
            "Dokumen", "Aksi", "Tanggal Unggah", "Update", "User", "User Edit" 
        ];

        let combinedData = [];
        combinedData.push(...mapSheet(PAUD_SHEET, 'PAUD', PAUD_TIME_ZONE));
        combinedData.push(...mapSheet(SD_SHEET, 'SD', SD_TIME_ZONE));

        combinedData.sort((a, b) => b._sortDate - a._sortDate);

        return { headers: FINAL_HEADERS, rows: combinedData };

    } catch (e) {
        return handleError('getLapbulKelolaData', e);
    }
}

/**
 * FUNGSI AMBIL DATA UNTUK EDIT (DYNAMIC HEADER)
 * Membaca baris header (1) dan data baris (rowIndex) lalu menggabungkannya jadi Object.
 */
function getLapbulDataByRow(rowIndex, source) {
    try {
        const config = (source === 'SD') ? 
            SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD : 
            SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD;
            
        const ss = SpreadsheetApp.openById(config.id);
        const sheet = ss.getSheetByName(config.sheet);
        
        // Validasi Row Index
        const lastRow = sheet.getLastRow();
        if (rowIndex < 2 || rowIndex > lastRow) throw new Error("Baris data tidak valid.");

        // 1. Ambil Header (Baris 1) & Data (Baris Terpilih)
        const lastCol = sheet.getLastColumn();
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const dataRow = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0]; // Ambil data raw (biar tanggal bisa dibaca input date)

        // 2. Gabungkan jadi Object { "Nama Kolom": "Isi Data" }
        const result = {};
        headers.forEach((header, index) => {
            let val = dataRow[index];
            
            // Konversi Tanggal ke String YYYY-MM-DD (untuk input type="date")
            if (val instanceof Date) {
                val = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
            }
            // Konversi data kosong/null jadi string kosong
            if (val === null || val === undefined) val = "";
            
            result[header] = val;
        });

        // Kirim juga Link Dokumen asli (cari kolom yang isinya http)
        // (Opsional: sudah tercover di loop atas jika nama kolomnya 'Dokumen')
        
        return result;

    } catch (e) {
        return handleError('getLapbulDataByRow', e);
    }
}

/**
 * FUNGSI SIMPAN PERUBAHAN (UPDATE)
 * Otomatis mengisi kolom 'Update' dan 'User'
 */
function updateLapbulData(formObject) {
    try {
        const source = (formObject.source || 'PAUD').toUpperCase();
        const rowIndex = parseInt(formObject.rowIndex);
        if (!rowIndex) throw new Error("ID Baris tidak ditemukan.");

        const config = (source === 'SD') ? 
            SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD : 
            SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD;

        const ss = SpreadsheetApp.openById(config.id);
        const sheet = ss.getSheetByName(config.sheet);
        const lastCol = sheet.getLastColumn();
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

        // 1. Update Data Berdasarkan Header Input Form
        headers.forEach((header, index) => {
            if (formObject.hasOwnProperty(header)) {
                sheet.getRange(rowIndex, index + 1).setValue(formObject[header]);
            }
        });

        // 2. Simpan File Baru (Jika ada)
        if (formObject.fileData && formObject.fileData.data) {
            const blob = Utilities.newBlob(Utilities.base64Decode(formObject.fileData.data), formObject.fileData.mimeType, formObject.fileData.fileName);
            const folder = DriveApp.getFolderById((source==='SD') ? FOLDER_CONFIG.LAPBUL_SD : FOLDER_CONFIG.LAPBUL_KB);
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            
            const docIndex = headers.indexOf("Dokumen");
            if (docIndex > -1) sheet.getRange(rowIndex, docIndex + 1).setValue(file.getUrl());
        }

        // --- 3. FITUR OTOMATIS (Sesuai Request) ---
        
        // A. Isi Kolom 'Update' dengan Waktu Sekarang
        const updateIndex = headers.indexOf("Update");
        if (updateIndex > -1) {
            sheet.getRange(rowIndex, updateIndex + 1).setValue(new Date());
        }

        // B. Isi Kolom 'User' dengan Username Pengedit
        if (formObject.User) {
             const userIndex = headers.indexOf("User");
             if (userIndex > -1) {
                 sheet.getRange(rowIndex, userIndex + 1).setValue(formObject.User);
             }
        }

        return "Data berhasil diperbarui.";

    } catch (e) {
        return handleError('updateLapbulData', e);
    }
}

/**
 * FUNGSI HAPUS (SOFT DELETE)
 * Memindahkan data ke sheet "Trash" dan file ke folder Trash Drive
 */
function deleteLapbulData(rowIndex, source, deleteCode, deleterName, deleterRole) {
  try {
    // 1. Validasi Kode
    const today = new Date();
    const todayCode = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) {
      throw new Error("Kode Hapus salah. Gunakan format YYYYMMDD hari ini.");
    }

    // 2. Buka Sheet Sumber
    let configKey = (source === 'SD') ? 'LAPBUL_FORM_RESPONSES_SD' : 'LAPBUL_FORM_RESPONSES_PAUD';
    const config = SPREADSHEET_CONFIG[configKey];
    const ss = SpreadsheetApp.openById(config.id);
    const sheet = ss.getSheetByName(config.sheet);
    
    // 3. Ambil Data Baris yang akan dihapus
    const lastCol = sheet.getLastColumn();
    const range = sheet.getRange(rowIndex, 1, 1, lastCol);
    const rowValues = range.getValues()[0];
    
    // Identifikasi Kolom Penting (Berdasarkan Mapping di LAPBUL_COLUMNS)
    // PAUD & SD punya index yang mirip untuk Tahun(2) dan Bulan(1)
    const tahun = String(rowValues[2]); 
    const bulan = String(rowValues[1]);
    const fileUrl = String(rowValues[ (source==='SD' ? 7 : 36) ]); // SD col H(7), PAUD col AK(36)

    // 4. PINDAHKAN FILE KE TRASH DRIVE
    if (fileUrl && fileUrl.includes('http')) {
        const fileIdMatch = fileUrl.match(/[-\w]{25,}/);
        if (fileIdMatch) {
            try {
                const file = DriveApp.getFileById(fileIdMatch[0]);
                
                // Ambil Folder Trash Root
                const trashRootId = (source === 'SD') ? FOLDER_CONFIG.TRASH_SD : FOLDER_CONFIG.TRASH_PAUD;
                const trashRoot = DriveApp.getFolderById(trashRootId);
                
                // Buat/Cari Subfolder Tahun -> Bulan
                const yearFolder = getOrCreateFolder(trashRoot, tahun);
                const monthFolder = getOrCreateFolder(yearFolder, bulan);
                
                // Pindahkan File
                file.moveTo(monthFolder);
                
            } catch (err) {
                Logger.log("Gagal memindahkan file ke trash: " + err.message);
                // Lanjut saja meski file gagal dipindah (mungkin sudah hilang)
            }
        }
    }

    // 5. PINDAHKAN DATA KE SHEET "Trash"
    let trashSheet = ss.getSheetByName("Trash");
    if (!trashSheet) {
        // Buat sheet Trash jika belum ada
        trashSheet = ss.insertSheet("Trash");
        // Copy header
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        // Tambah header info hapus
        headers.push("Deleted By", "Role", "Deleted At");
        trashSheet.appendRow(headers);
    }

    // Tambahkan info penghapus ke data
    const trashRow = [...rowValues, deleterName, deleterRole, new Date()];
    trashSheet.appendRow(trashRow);

    // 6. HAPUS DARI SHEET UTAMA
    sheet.deleteRow(rowIndex);

    return "Data berhasil dipindahkan ke Sampah (Trash).";

  } catch (e) {
    return handleError("deleteLapbulData", e);
  }
}

function getLapbulInfo() {
  const cacheKey = 'lapbul_info_v2'; // Ubah kunci cache agar selalu memuat data baru
  return getCachedData(cacheKey, function() {
    try {
      // Menggunakan SPREADSHEET_CONFIG.DROPDOWN_DATA (ID: 1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA)
      const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
      const sheet = ss.getSheetByName('Informasi');
      
      if (!sheet || sheet.getLastRow() < 2) {
        return []; // Kembalikan array kosong jika sheet tidak valid atau hanya header
      }

      const lastRow = sheet.getLastRow();
      // Ambil data dari A2 sampai baris terakhir
      const range = sheet.getRange('A2:A' + lastRow);
      const values = range.getValues()
                          .flat()
                          .filter(item => String(item).trim() !== ''); // Filter baris kosong
      
      // Jika Anda ingin mengambil data dari kolom B, gunakan range 'B2:B' + lastRow. 
      // Saya asumsikan Anda ingin kolom A (Informasi Umum) di sini.

      return values;
    } catch (e) {
      Logger.log(`Error in getLapbulInfo fetch: ${e.message}`);
      return []; // Kembalikan array kosong untuk menghindari error di client
    }
  });
}

function getUnduhFormatInfo() {
  const cacheKey = 'unduh_format_info_v1';
  return getCachedData(cacheKey, function() {
    try {
      const ss = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
      const sheet = ss.getSheetByName('Informasi');
      if (!sheet || sheet.getLastRow() < 2) return [];
      const range = sheet.getRange('B2:B' + sheet.getLastRow());
      return range.getDisplayValues().flat().filter(item => String(item).trim() !== '');
    } catch (e) {
      return handleError('getUnduhFormatInfo', e);
    }
  });
}

function getLapbulArsipFolderIds() {
  try {
    return {
      'KB': FOLDER_CONFIG.LAPBUL_KB,
      'TK': FOLDER_CONFIG.LAPBUL_TK,
      'SD': FOLDER_CONFIG.LAPBUL_SD
    };
  } catch (e) {
    return handleError('getLapbulArsipFolderIds', e);
  }
}

/**
 * FUNGSI ARSIP LAPORAN BULAN
 * Merayapi Google Drive: Root (Jenjang) -> Tahun -> Bulan -> File
 */
function getLapbulArsipData(filter) {
  const fTahun = (filter && filter.tahun) ? String(filter.tahun) : new Date().getFullYear().toString();
  const fJenjang = (filter && filter.jenjang) ? filter.jenjang : ""; 
  const fBulan = (filter && filter.bulan) ? filter.bulan : "";

  // Mapping ID Folder Induk (Dari Code.gs / FOLDER_CONFIG)
  // Pastikan variabel global FOLDER_CONFIG terbaca
  const roots = {
    "KB": FOLDER_CONFIG.LAPBUL_KB,
    "TK": FOLDER_CONFIG.LAPBUL_TK,
    "SD": FOLDER_CONFIG.LAPBUL_SD
  };

  const rows = [];

  // Helper: Format Ukuran File
  const formatSize = (bytes) => {
    if (bytes === 0) return '0 B';
    const k = 1024, sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  };

  // Helper: Proses Satu Folder Jenjang
  const processJenjang = (label, folderId) => {
    try {
      const rootFolder = DriveApp.getFolderById(folderId);
      
      // 1. Cari Folder TAHUN (Exact Match)
      const yearFolders = rootFolder.getFoldersByName(fTahun);
      if (!yearFolders.hasNext()) return; // Tahun tidak ditemukan
      const yearFolder = yearFolders.next();

      // 2. Cari Folder BULAN
      // Jika filter bulan ada, cari spesifik. Jika tidak, ambil semua folder di dalam tahun.
      let monthFolders = [];
      if (fBulan) {
          const mIt = yearFolder.getFoldersByName(fBulan);
          if (mIt.hasNext()) monthFolders.push(mIt.next());
      } else {
          const mIt = yearFolder.getFolders();
          while (mIt.hasNext()) monthFolders.push(mIt.next());
      }

      // 3. Ambil File dari Setiap Folder Bulan
      monthFolders.forEach(mFolder => {
          const mName = mFolder.getName(); // Nama Bulan
          const files = mFolder.getFiles();
          
          while (files.hasNext()) {
              const file = files.next();
              // Simpan data file
              rows.push({
                  "Nama File": file.getName(),
                  "Jenjang": label,
                  "Tahun": fTahun,
                  "Bulan": mName,
                  "Ukuran": formatSize(file.getSize()),
                  "Link": file.getUrl(), // Untuk tombol Download/Lihat
                  "Waktu": Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
                  "Tipe": file.getMimeType()
              });
          }
      });
    } catch (e) {
      console.error(`Error Arsip ${label}: ${e.message}`);
    }
  };

  // Eksekusi Sesuai Filter Jenjang
  if (fJenjang === "" || fJenjang === "KB") processJenjang("KB", roots.KB);
  if (fJenjang === "" || fJenjang === "TK") processJenjang("TK", roots.TK);
  if (fJenjang === "" || fJenjang === "SD") processJenjang("SD", roots.SD);

  // Sort: Bulan (Mundur), lalu Nama File (A-Z)
  // Mapping Bulan ke Angka untuk sorting
  const monthMap = { "Januari":1, "Februari":2, "Maret":3, "April":4, "Mei":5, "Juni":6, "Juli":7, "Agustus":8, "September":9, "Oktober":10, "November":11, "Desember":12 };
  
  rows.sort((a, b) => {
      const mA = monthMap[a.Bulan] || 0;
      const mB = monthMap[b.Bulan] || 0;
      if (mB !== mA) return mB - mA; // Bulan terbaru di atas
      return a["Nama File"].localeCompare(b["Nama File"]);
  });

  return { 
      headers: ["Nama File", "Jenjang", "Tahun", "Bulan", "Ukuran", "Waktu", "Aksi"], 
      rows: rows 
  };
}

/**
 * FUNGSI FILE EXPLORER (ARSIP) - DENGAN SORTING BULAN
 */
function getLapbulArsipContent(folderId) {
  try {
    const items = [];
    const timeZone = Session.getScriptTimeZone();

    // --- 1. ROOT (BERANDA) ---
    if (!folderId) {
       if (typeof FOLDER_CONFIG === 'undefined') throw new Error("FOLDER_CONFIG Error");
       items.push({ id: FOLDER_CONFIG.LAPBUL_KB, name: "Arsip PAUD (KB)", type: "folder" });
       items.push({ id: FOLDER_CONFIG.LAPBUL_TK, name: "Arsip PAUD (TK)", type: "folder" });
       items.push({ id: FOLDER_CONFIG.LAPBUL_SD, name: "Arsip SD",        type: "folder" });
       
       return { currentName: "Beranda", parents: [], items: items };
    }

    // --- 2. ISI FOLDER ---
    const folder = DriveApp.getFolderById(folderId);
    
    // Ambil Sub-Folder
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
        const f = subFolders.next();
        items.push({
            id: f.getId(), name: f.getName(), type: "folder",
            updated: Utilities.formatDate(f.getLastUpdated(), timeZone, "dd/MM/yyyy HH:mm")
        });
    }

    // Ambil File
    const files = folder.getFiles();
    while (files.hasNext()) {
        const f = files.next();
        items.push({
            id: f.getId(), name: f.getName(), type: "file",
            mimeType: f.getMimeType(),
            size: (f.getSize() / 1024).toFixed(1) + " KB",
            url: f.getUrl(),
            updated: Utilities.formatDate(f.getLastUpdated(), timeZone, "dd/MM/yyyy HH:mm")
        });
    }

    // --- LOGIKA PENGURUTAN (SORTING) ---
    // Kamus Angka Bulan
    const monthMap = {
        "januari": 1, "februari": 2, "maret": 3, "april": 4, "mei": 5, "juni": 6,
        "juli": 7, "agustus": 8, "september": 9, "oktober": 10, "november": 11, "desember": 12
    };

    items.sort((a, b) => {
        // 1. Folder selalu di atas File
        if (a.type !== b.type) return a.type === 'folder' ? -1 : 1;
        
        const nameA = a.name.toLowerCase().trim();
        const nameB = b.name.toLowerCase().trim();

        // 2. Jika keduanya adalah BULAN, urutkan berdasarkan Kalender (1-12)
        if (monthMap[nameA] && monthMap[nameB]) {
            return monthMap[nameA] - monthMap[nameB];
        }

        // 3. Jika keduanya adalah TAHUN (Angka), urutkan Descending (Terbaru di atas)
        // Contoh: 2025 dulu, baru 2024
        if (!isNaN(nameA) && !isNaN(nameB) && nameA.length === 4) {
             return nameB.localeCompare(nameA); 
        }

        // 4. Sisanya urut Abjad biasa (A-Z)
        return a.name.localeCompare(b.name);
    });

    return { 
        currentId: folderId, 
        currentName: folder.getName(), 
        items: items 
    };

  } catch (e) {
    return { error: e.message };
  }
}

/**
 * FUNGSI AMBIL DATA TRASH (SAMA SEPERTI RIWAYAT TAPI DARI SHEET TRASH)
 */
function getLapbulTrashData() {
    try {
        const PAUD_CONFIG = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD;
        const SD_CONFIG = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD;
        
        let combinedData = [];

        // Helper untuk baca sheet Trash
        const fetchTrash = (config, source) => {
            const ss = SpreadsheetApp.openById(config.id);
            const sheet = ss.getSheetByName("Trash");
            if (!sheet || sheet.getLastRow() < 2) return [];

            const data = sheet.getDataRange().getDisplayValues();
            const headers = data[0];
            const rows = data.slice(1);

            // Cari index kolom penting berdasarkan nama header
            const getIdx = (name) => headers.indexOf(name);
            
            const idxNama = getIdx("Nama Sekolah");
            const idxBulan = getIdx("Bulan");
            const idxTahun = getIdx("Tahun");
            const idxDokumen = getIdx("Dokumen");
            const idxDelBy = getIdx("Deleted By");
            const idxDelAt = getIdx("Deleted At");

            return rows.map((row, i) => {
                // Filter baris kosong
                if (!row[idxNama]) return null;

                return {
                    "Nama Sekolah": row[idxNama] || '-',
                    "Jenjang": (source === 'SD') ? 'SD' : 'PAUD',
                    "Bulan": row[idxBulan] || '-',
                    "Tahun": row[idxTahun] || '-',
                    "Dokumen": row[idxDokumen] || '',
                    "Deleted By": row[idxDelBy] || '-',
                    "Deleted At": row[idxDelAt] || '-',
                    "_source": source
                };
            }).filter(r => r !== null);
        };

        // Ambil Data
        combinedData.push(...fetchTrash(PAUD_CONFIG, 'PAUD'));
        combinedData.push(...fetchTrash(SD_CONFIG, 'SD'));

        // Sort berdasarkan waktu hapus (terbaru di atas)
        // Kita parse string 'Deleted At' sederhana
        combinedData.sort((a, b) => new Date(b["Deleted At"]) - new Date(a["Deleted At"]));

        return combinedData;

    } catch (e) {
        return handleError('getLapbulTrashData', e);
    }
}

/**
 * STATISTIK DASHBOARD LAPBUL (FIX: GRAFIK & NAMA SEKOLAH)
 */
function getLapbulDashboardStats() {
    try {
        const PAUD_CONFIG = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD;
        const SD_CONFIG = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD;
        
        // Array Tren Bulanan (Jan-Des)
        let trendSD = new Array(12).fill(0);
        let trendPAUD = new Array(12).fill(0);
        
        let allRows = [];
        const now = new Date();
        const currentYear = now.getFullYear();
        const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();

        // Helper Ambil Data
        const fetch = (conf, jenjang) => {
            try {
                const ss = SpreadsheetApp.openById(conf.id);
                const sheet = ss.getSheetByName(conf.sheet);
                if (!sheet) return;

                const data = sheet.getDataRange().getValues();
                if (data.length < 2) return;

                const headers = data[0]; // Baris Header
                
                // 1. CARI INDEX KOLOM PENTING
                const idxNama = headers.indexOf("Nama Sekolah");
                // Cari "Tanggal Unggah" atau gunakan Timestamp (0) jika tidak ada
                let idxTgl = headers.indexOf("Tanggal Unggah");
                if (idxTgl === -1) idxTgl = 0; // Fallback ke Timestamp

                // Validasi: Jika kolom Nama Sekolah tidak ditemukan, skip
                if (idxNama === -1) return;

                data.slice(1).forEach(r => {
                    // Pastikan Nama Sekolah ada
                    if (!r[idxNama]) return;

                    const tgl = r[idxTgl];
                    const nama = r[idxNama];

                    // Masukkan ke List Gabungan
                    allRows.push({
                        tgl: tgl,
                        sekolah: String(nama),
                        jenjang: jenjang
                    });

                    // 2. ISI DATA TREN (Langsung di sini agar efisien)
                    if (tgl instanceof Date && tgl.getFullYear() === currentYear) {
                        const month = tgl.getMonth(); // 0 = Jan, 11 = Des
                        if (jenjang === 'SD') {
                            trendSD[month]++;
                        } else {
                            trendPAUD[month]++;
                        }
                    }
                });

            } catch (err) {
                Logger.log("Error fetch " + jenjang + ": " + err.message);
            }
        };

        fetch(PAUD_CONFIG, 'PAUD');
        fetch(SD_CONFIG, 'SD');

        // HITUNG TOTAL
        let total = allRows.length;
        let todayCount = 0;
        let countSD = 0;
        let countPAUD = 0;

        allRows.forEach(r => {
            if (r.jenjang === 'SD') countSD++; else countPAUD++;
            if (r.tgl instanceof Date && r.tgl.getTime() >= startOfDay) todayCount++;
        });

        // RECENT ACTIVITY (Sort Terbaru)
        allRows.sort((a, b) => {
            let dA = (a.tgl instanceof Date) ? a.tgl.getTime() : 0;
            let dB = (b.tgl instanceof Date) ? b.tgl.getTime() : 0;
            return dB - dA;
        });

        let recent = allRows.slice(0, 5).map(r => {
            let t = (r.tgl instanceof Date) ? r.tgl : new Date();
            let timeStr = Utilities.formatDate(t, Session.getScriptTimeZone(), "dd/MM HH:mm");
            return {
                sekolah: r.sekolah, // Sekarang pasti Nama Sekolah
                waktu: timeStr,
                jenjang: r.jenjang
            };
        });

        return {
            total: total,
            today: todayCount,
            countSD: countSD,
            countPAUD: countPAUD,
            trendSD: trendSD,     // Data Baru
            trendPAUD: trendPAUD, // Data Baru
            recent: recent
        };

    } catch (e) {
        return { error: e.message };
    }
}