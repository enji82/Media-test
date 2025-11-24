/**
 * ===================================================================
 * ======================== MODUL: DATA SIABA ========================
 * ===================================================================
 */

/**
 * [BARU] Helper internal untuk mencari ID Spreadsheet bulanan dari 'Lookup SIABA'.
 */
function _findSiabaSpreadsheetId(tahun, bulan) {
  try {
    const ssDropdown = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
    const lookupSheet = ssDropdown.getSheetByName("Lookup SIABA");
    
    if (!lookupSheet) {
      throw new Error("Sheet 'Lookup SIABA' tidak ditemukan di SPREADSHEET_DROPDOWN_DATA.");
    }
    
    // Ambil data lookup (Tahun, Bulan, ID)
    const lookupData = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 3).getDisplayValues();
    
    // Cari baris yang cocok
    for (const row of lookupData) {
      if (String(row[0]) === String(tahun) && String(row[1]) === String(bulan)) {
        return row[2]; // Kembalikan ID (kolom C)
      }
    }
    
    return null; // Tidak ditemukan
    
  } catch (e) {
    Logger.log(`Error di _findSiabaSpreadsheetId: ${e.message}`);
    return null;
  }
}


/**
 * [PERBAIKAN TOTAL] Mengambil data dari sheet bulanan yang spesifik, bukan dari master IMPORTRANGE.
 * Ini akan SANGAT CEPAT.
 */
function getSiabaPresensiData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    
    if (!tahun || !bulan) throw new Error("Filter Tahun dan Bulan wajib diisi.");
    if (!unitKerja) throw new Error("Filter Unit Kerja wajib diisi.");

    // 1. Cari ID Spreadsheet bulanan (Logika ini sudah benar)
    const targetSheetId = _findSiabaSpreadsheetId(tahun, bulan);
    
    if (!targetSheetId) {
      Logger.log(`Tidak ada Spreadsheet_ID ditemukan untuk ${bulan} ${tahun} di sheet 'Lookup SIABA'.`);
      return { headers: [], rows: [] };
    }

    // 2. Buka Spreadsheet dan Sheet "WebData" (Logika ini sudah benar)
    let ss;
    try {
      ss = SpreadsheetApp.openById(targetSheetId);
    } catch (e) {
      throw new Error(`Gagal membuka Spreadsheet ID: ${targetSheetId}. Pastikan ID di 'Lookup SIABA' benar.`);
    }

    const sheet = ss.getSheetByName("WebData"); 
    if (!sheet || sheet.getLastRow() < 2) {
       Logger.log(`Sheet 'WebData' tidak ditemukan atau kosong di Spreadsheet ID: ${targetSheetId}`);
       return { headers: [], rows: [] };
    }

    // 3. Ambil data mentah (Bersihkan header dari spasi)
    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0].map(h => String(h).trim()); // <-- .trim() penting
    const dataRows = allData.slice(1);

    // 4. Tentukan Indeks Kolom "Unit Kerja" (Kolom BO = Indeks 66)
    // (Berdasarkan konfirmasi Anda)
    const unitKerjaIndex = 66; 
    
    if (dataRows.length > 0 && dataRows[0].length <= unitKerjaIndex) {
      throw new Error(`Gagal memfilter: Sheet "WebData" hanya memiliki ${dataRows[0].length} kolom, tetapi kode mencoba membaca kolom BO (indeks 66).`);
    }

    // 5. Terapkan filter 'unitKerja' (Logika "Semua" tetap ada)
    const filteredRows = dataRows.filter(row => {
        const matchUnitKerja = (unitKerja === "Semua") || (String(row[unitKerjaIndex]) === String(unitKerja));
        return matchUnitKerja;
    });

    // --- PERBAIKAN UTAMA DIMULAI DI SINI ---

    // 6. Tentukan Header yang Anda Inginkan (SESUAI DAFTAR DARI ANDA)
    const desiredHeaders = [
        "NAMA ASN", "NIP", "TP", "HK", "H", "HA", "APL", 
        "TAp", "HU", "U", "TU", "CT", "CAP", "CS", "CM", "DD", "DL", 
        "TA", "Waktu TA", "PLA", "Waktu PLA", "LA", 
        "1D", "1P", "2D", "2P", "3D", "3P", "4D", "4P", "5D", "5P", 
        "6D", "6P", "7D", "7P", "8D", "8P", "9D", "9P", "10D", "10P", 
        "11D", "11P", "12D", "12P", "13D", "13P", "14D", "14P", "15D", "15P", 
        "16D", "16P", "17D", "17P", "18D", "18P", "19D", "19P", "20D", "20P", 
        "21D", "21P", "22D", "22P", "23D", "23P", "24D", "24P", "25D", "25P", 
        "26D", "26P", "27D", "27P", "28D", "28P", "29D", "29P", "30D", "30P", 
        "31D", "31P"
    ];

    // 7. Buat "Peta Indeks"
    const headerMap = {};
    headers.forEach((headerName, index) => {
        headerMap[headerName] = index;
    });
    
    // Periksa apakah "NAMA ASN" ada di WebData, jika tidak, coba "Nama"
    if (!headerMap.hasOwnProperty("NAMA ASN") && headerMap.hasOwnProperty("Nama")) {
      headerMap["NAMA ASN"] = headerMap["Nama"]; // Buat alias
    }

    // 8. Susun Ulang Data (Transformasi)
    const displayRows = filteredRows.map(row => {
        const newRow = [];
        for (const desiredHeader of desiredHeaders) {
            const indexDiWebData = headerMap[desiredHeader];
            
            if (indexDiWebData !== undefined) {
                newRow.push(row[indexDiWebData]);
            } else {
                newRow.push('-'); // Kolom tidak ditemukan di WebData
            }
        }
        return newRow;
    });
    
    // 9. Logika Sorting (PERBAIKAN: Menggunakan "NAMA ASN")
    const tpIndex = desiredHeaders.indexOf('TP');
    const taIndex = desiredHeaders.indexOf('TA');
    const plaIndex = desiredHeaders.indexOf('PLA');
    const tapIndex = desiredHeaders.indexOf('TAp');
    const tuIndex = desiredHeaders.indexOf('TU');
    const namaIndex = desiredHeaders.indexOf('NAMA ASN'); // <-- PERBAIKAN

    // --- AKHIR PERBAIKAN UTAMA ---

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

    // 10. Kembalikan data
    return { headers: desiredHeaders, rows: displayRows };
    
  } catch (e) {
    return handleError('getSiabaPresensiData', e);
  }
}


/**
 * [PERBAIKAN] Mengambil filter dari DROPDOWN_DATA, bukan sheet rekap yang besar.
 * Ini akan memuat dropdown filter secara instan.
 */
function getSiabaFilterOptions() {
  try {
    // 1. Buka spreadsheet DROPDOWN_DATA (CEPAT)
    const ssDropdown = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);

    // 2. Ambil Unit Kerja (Ini sudah benar dan cepat)
    const sheetUnitKerja = ssDropdown.getSheetByName("Unit Siaba");
    let unitKerjaOptions = [];
    if (sheetUnitKerja && sheetUnitKerja.getLastRow() > 1) {
      unitKerjaOptions = sheetUnitKerja.getRange(2, 1, sheetUnitKerja.getLastRow() - 1, 1)
                                      .getDisplayValues().flat().filter(Boolean).sort();
    }

    // 3. Ambil Tahun dan Bulan dari sheet "Filter Siaba" (BARU & CEPAT)
    const sheetFilterSiaba = ssDropdown.getSheetByName("Filter Siaba");
    if (!sheetFilterSiaba || sheetFilterSiaba.getLastRow() < 2) {
         throw new Error("Sheet 'Filter Siaba' tidak ditemukan atau kosong di SPREADSHEET_DROPDOWN_DATA.");
    }
    
    // Ambil data dari sheet kecil
    const filterData = sheetFilterSiaba.getRange(2, 1, sheetFilterSiaba.getLastRow() - 1, 2).getDisplayValues();
    
    const uniqueTahun = [...new Set(filterData.map(row => row[0]))].filter(Boolean).sort().reverse();
    const uniqueBulan = [...new Set(filterData.map(row => row[1]))].filter(Boolean);
    
    const monthOrder = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    uniqueBulan.sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));

    // 4. Kembalikan data filter
    return {
      'Tahun': uniqueTahun,
      'Bulan': uniqueBulan,
      'Unit Kerja': unitKerjaOptions
    };
    
  } catch (e) {
    return handleError('getSiabaFilterOptions', e);
  }
}

function getSiabaTidakPresensiData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    
    if (!tahun || !bulan) throw new Error("Filter Tahun dan Bulan wajib diisi.");
    if (!unitKerja) throw new Error("Filter Unit Kerja wajib diisi.");

    // 1. Cari ID Spreadsheet bulanan (Logika ini sudah benar)
    const targetSheetId = _findSiabaSpreadsheetId(tahun, bulan);
    
    if (!targetSheetId) {
      Logger.log(`Tidak ada Spreadsheet_ID ditemukan untuk ${bulan} ${tahun} di sheet 'Lookup SIABA'.`);
      return { headers: [], rows: [] };
    }

    // 2. Buka Spreadsheet dan Sheet "WebData" (Logika ini sudah benar)
    let ss;
    try {
      ss = SpreadsheetApp.openById(targetSheetId);
    } catch (e) {
      throw new Error(`Gagal membuka Spreadsheet ID: ${targetSheetId}. Pastikan ID di 'Lookup SIABA' benar.`);
    }

    const sheet = ss.getSheetByName("WebData"); 
    if (!sheet || sheet.getLastRow() < 2) {
       Logger.log(`Sheet 'WebData' tidak ditemukan atau kosong di Spreadsheet ID: ${targetSheetId}`);
       return { headers: [], rows: [] };
    }

    // 3. Ambil data mentah (Bersihkan header dari spasi)
    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0].map(h => String(h).trim()); 
    const dataRows = allData.slice(1);

    // 4. Cari Indeks Kolom "Unit Kerja" (Kolom BO = Indeks 66)
    const unitKerjaIndex = 66; 
    
    // 5. Cari Indeks Kolom "TP" (PENTING untuk filter)
    const tpIndex = headers.indexOf('TP');
    if (tpIndex === -1) {
        throw new Error(`Header "TP" tidak ditemukan di sheet "WebData".`);
    }

    // 6. Terapkan filter 'unitKerja' DAN 'TP > 0'
    const nipIndex = headers.indexOf('NIP');
    if (nipIndex === -1) {
        // Jika kolom NIP tidak ada di "WebData", kita tidak bisa memfilter dengan aman
        throw new Error(`Header "NIP" tidak ditemukan di sheet "WebData".`);
    }
    // --- AKHIR PERBAIKAN BARU ---

    // 7. Terapkan filter 'unitKerja' DAN 'TP > 0' DAN 'NIP valid'
    const filteredRows = dataRows.filter(row => {
        
        // --- INI PERBAIKAN BARU ---
        // Cek 1: Pastikan baris ini punya NIP yang valid
        const nipString = String(row[nipIndex] || '');
        
        // Buang SEMUA karakter non-digit (spasi, petik, huruf, dll.)
        const nipDigits = nipString.replace(/\D/g, ''); 
        
        // NIP valid jika HANYA berisi 18 digit
        const isNipValid = (nipDigits.length === 18);
        
        if (!isNipValid) {
             // Log ini akan membantu jika masih gagal, dan HANYA log jika BUKAN string kosong
             if (nipString.trim() !== "") {
                Logger.log(`Baris data dilewati (NIP tidak valid): ${nipString}`);
             }
             return false; // Lewati baris ini (termasuk baris total/kosong)
        }
        // --- AKHIR PERBAIKAN BARU ---
        
        const matchUnitKerja = (unitKerja === "Semua") || (String(row[unitKerjaIndex]) === String(unitKerja));
        const matchTP = (parseInt(row[tpIndex], 10) || 0) > 0; 
        
        return matchUnitKerja && matchTP;
    });

    // 7. Tentukan Header yang Anda Inginkan (SESUAI PERMINTAAN ANDA)
    const desiredHeaders = [
        "Nama ASN", 
        "NIP", 
        "Jumlah", 
        "Tidak Datang", 
        "Tidak Pulang", 
        "Tanggal Tidak Presensi"
    ];
    
    // 8. Tentukan Header SUMBER (Nama di "WebData")
    const sourceHeaders = [
        "NAMA ASN", 
        "NIP", 
        "TP", 
        "LAD", 
        "LAP", 
        "TGL LUPA"
    ];

    // 9. Buat "Peta Indeks"
    const headerMap = {};
    headers.forEach((headerName, index) => {
        headerMap[headerName] = index;
    });
    
    // Alias untuk "NAMA ASN" jika di WebData namanya "Nama"
    if (!headerMap.hasOwnProperty("NAMA ASN") && headerMap.hasOwnProperty("Nama")) {
      headerMap["NAMA ASN"] = headerMap["Nama"]; 
    }

    // 10. Susun Ulang Data (Transformasi) - DENGAN TRY...CATCH ANTI-CRASH
    const displayRows = []; // Buat array kosong baru

    // Dapatkan indeks NAMA ASN untuk logging error
    let namaIndex = headerMap["NAMA ASN"]; 
    if (namaIndex === undefined) namaIndex = headerMap["Nama"];

    filteredRows.forEach((row, rowIndex) => {
        const namaASN = (namaIndex !== undefined) ? row[namaIndex] : `Baris #${rowIndex}`;
        try {
            // Coba proses baris ini
            const newRow = [];
            for (const sourceHeader of sourceHeaders) { // sourceHeaders = ["NAMA ASN", "NIP", "TP", "LAD", "LAP", "TGL LUPA"]
                const indexDiWebData = headerMap[sourceHeader];
                
                if (indexDiWebData !== undefined) {
                    const cellValue = row[indexDiWebData];
                    
                    // Cek error eksplisit: Google Sheets bisa mengembalikan objek error
                    if (cellValue instanceof Error) {
                        newRow.push('#ERROR!'); // Ganti error dengan teks aman
                    } else {
                        newRow.push(cellValue); // Salin data sel
                    }
                } else {
                    newRow.push('-'); // Kolom tidak ditemukan di WebData
                }
            }
            displayRows.push(newRow); // Tambahkan baris yang berhasil diproses
        
        } catch (e) {
            // Jika ada error saat memproses baris ini,
            // catat di log server dan lewati baris ini.
            Logger.log(`Peringatan (getSiabaTidakPresensiData): Gagal memproses data untuk ASN '${namaASN}'. Error: ${e.message}.`);
        }
    });
    
    // 11. Logika Sorting (Berdasarkan "Jumlah" (TP) descending, lalu "Nama ASN")
    const sortJumlahIndex = desiredHeaders.indexOf('Jumlah'); // Indeks 2
    const sortNamaIndex = desiredHeaders.indexOf('Nama ASN');   // Indeks 0

    displayRows.sort((a, b) => {
      const valB_jumlah = (sortJumlahIndex !== -1) ? (parseInt(b[sortJumlahIndex], 10) || 0) : 0;
      const valA_jumlah = (sortJumlahIndex !== -1) ? (parseInt(a[sortJumlahIndex], 10) || 0) : 0;
      if (valB_jumlah !== valA_jumlah) {
        return valB_jumlah - valA_jumlah; // Urutkan Jumlah (descending)
      }
      if (sortNamaIndex !== -1) {
          return (a[sortNamaIndex] || "").localeCompare(b[sortNamaIndex] || ""); // Urutkan Nama (ascending)
      }
       return 0;
    });

    // 12. Kembalikan data
    return { headers: desiredHeaders, rows: displayRows };
    
  } catch (e) {
    return handleError('getSiabaTidakPresensiData', e);
  }
}

function getSiabaApelUpacaraData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    
    // 1. Cari ID Spreadsheet bulanan
    const targetSheetId = _findSiabaSpreadsheetId(tahun, bulan);
    if (!targetSheetId) {
      return { headers: [], rows: [] }; // Spreadsheet belum ada
    }

    // 2. Buka Spreadsheet dan Sheet "Draft Rekap Apel"
    let ss;
    try {
      ss = SpreadsheetApp.openById(targetSheetId);
    } catch (e) {
      throw new Error(`Gagal membuka Spreadsheet ID: ${targetSheetId}.`);
    }

    const sheet = ss.getSheetByName("Draft Rekap Apel"); 
    if (!sheet || sheet.getLastRow() < 2) {
       // Jika sheet tidak ada, kembalikan kosong
       return { headers: [], rows: [] };
    }

    // 3. Ambil data mentah
    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0].map(h => String(h).trim()); 
    const dataRows = allData.slice(1);

    // 4. Cari Indeks Kolom Penting
    // Cari "Unit Kerja" (bisa bernama "Unit Kerja", "Unit", atau di indeks tertentu)
    let unitKerjaIndex = headers.indexOf("Unit Kerja");
    if (unitKerjaIndex === -1) {
         // Fallback: Coba cari kolom yang isinya mirip unit kerja atau gunakan index 66 (default WebData) 
         // Namun untuk keamanan, kita coba cari "Nama Unit" atau sejenisnya
         unitKerjaIndex = headers.findIndex(h => h.toLowerCase().includes("unit kerja"));
    }
    
    // 5. Terapkan Filter Unit Kerja
    const filteredRows = dataRows.filter(row => {
        // Jika unitKerjaIndex tidak ketemu, kita loloskan semua (atau bisa throw error)
        // Asumsi: Jika Unit Kerja "Semua", lolos.
        if (unitKerja === "Semua") return true;
        
        // Jika kolom Unit Kerja ada, cek kecocokan
        if (unitKerjaIndex !== -1) {
             return String(row[unitKerjaIndex]) === String(unitKerja);
        }
        return true; // Default allow jika kolom tidak ditemukan (hati-hati)
    });

    // 6. Definisikan Header yang Diinginkan (Target)
    const dateHeaders = Array.from({length: 31}, (_, i) => String(i + 1));
    const desiredHeaders = [
        "Nama ASN", "NIP", 
        "HA", "A", "TA", 
        "HU", "U", "TU", 
        ...dateHeaders
    ];

    // 7. Mapping Header Sumber (Sheet) ke Target
    const headerMap = {};
    headers.forEach((headerName, index) => {
        headerMap[headerName] = index;
    });

    // === [BARU] PEMETAAN KOLOM KHUSUS ===
    // Format: "HEADER_DI_TABEL_WEB": "HEADER_DI_SPREADSHEET_SUMBER"
    const columnAliases = {
        "Nama ASN": "Nama",       // Alias Nama
        "HA": "HARI APEL",        // HA -> HARI APEL
        "A": "APL",               // A -> APL
        "TA": "TAPL",             // TA -> TAPL
        "HU": "HARI U",           // HU -> HARI U
        "U": "U",                 // U -> U (Sama)
        "TU": "TUPC"              // TU -> TUPC
    };

    // Terapkan alias: Jika header target tidak ketemu, cari pakai nama aliasnya
    for (const [targetKey, sourceKey] of Object.entries(columnAliases)) {
        if (headerMap[targetKey] === undefined && headerMap[sourceKey] !== undefined) {
            headerMap[targetKey] = headerMap[sourceKey];
        }
    }

    // 8. Transformasi Data (Anti-Crash)
    const displayRows = [];
    
    // Dapatkan indeks NAMA ASN untuk logging
    let namaIndex = headerMap["Nama ASN"];

    filteredRows.forEach((row, rowIndex) => {
        const namaASN = (namaIndex !== undefined) ? row[namaIndex] : `Baris #${rowIndex}`;
        try {
            const newRow = [];
            for (const targetHeader of desiredHeaders) {
                const sourceIndex = headerMap[targetHeader];
                
                if (sourceIndex !== undefined) {
                    const cellValue = row[sourceIndex];
                    if (cellValue instanceof Error) {
                        newRow.push('#ERROR!'); 
                    } else {
                        newRow.push(cellValue); 
                    }
                } else {
                    newRow.push('-'); // Kolom tidak ditemukan
                }
            }
            
            // Filter Tambahan: Hanya masukkan jika Nama ASN ada (Membersihkan baris kosong/total)
            // Cek kolom ke-0 (Nama ASN)
            if (newRow[0] && newRow[0] !== '-' && newRow[0].trim() !== '') {
                 displayRows.push(newRow);
            }
            
        } catch (e) {
            Logger.log(`Peringatan (getSiabaApelUpacaraData): Gagal baris '${namaASN}'. Error: ${e.message}.`);
        }
    });
    
    // Sorting: Nama ASN (Ascending)
    const sortNamaIndex = 0; // Index kolom Nama ASN di desiredHeaders
    displayRows.sort((a, b) => (a[sortNamaIndex] || "").localeCompare(b[sortNamaIndex] || ""));

    return { headers: desiredHeaders, rows: displayRows };

  } catch (e) {
    return handleError('getSiabaApelUpacaraData', e);
  }
}

function getSiabaLupaOptions() {
  // Cache selama 1 jam karena database pegawai jarang berubah
  const cacheKey = 'siaba_lupa_options_v1';
  return getCachedData(cacheKey, function() {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DB);
      const sheet = ss.getSheetByName("Database");
      if (!sheet) throw new Error("Sheet 'Database' tidak ditemukan.");

      // Ambil data dari baris 2 sampai akhir, kolom A(1) s/d C(3)
      // Kolom A: NIP, B: Nama ASN, C: Unit Kerja
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];
      
      const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      
      // Format data menjadi array of objects agar mudah diolah di client
      const formattedData = data.map(row => ({
        nip: String(row[0]).trim(),
        nama: String(row[1]).trim(),
        unit: String(row[2]).trim()
      })).filter(item => item.nama !== "" && item.unit !== "");

      return formattedData;

    } catch (e) {
      throw new Error(`Gagal mengambil data pegawai: ${e.message}`);
    }
  });
}

/**
 * Memproses formulir Lupa Presensi
 */
function processSiabaLupaForm(formData) {
  try {
    // 1. Validasi Folder Utama
    const folderId = FOLDER_CONFIG.SIABA_LUPA;
    if (!folderId || folderId === "GANTI_DENGAN_ID_FOLDER_DRIVE_UNTUK_SIMPAN_SURAT") {
       throw new Error("Folder penyimpanan belum dikonfigurasi (FOLDER_CONFIG.SIABA_LUPA).");
    }

    const mainFolder = DriveApp.getFolderById(folderId);

    // 2. Tentukan Folder Bulan (Berdasarkan Tanggal Lupa)
    // formData.Tanggal formatnya "YYYY-MM-DD"
    const tanggalLupa = new Date(formData.Tanggal);
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    
    // Nama Folder: "November 2025"
    const folderName = `${monthNames[tanggalLupa.getMonth()]} ${tanggalLupa.getFullYear()}`;
    
    // Cari atau Buat Folder Bulan tersebut (Menggunakan fungsi helper getOrCreateFolder di Code.gs)
    const targetFolder = getOrCreateFolder(mainFolder, folderName);

    // 3. Siapkan File
    const fileData = formData.fileData;
    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    
    // 4. Format Nama File: <Unit Kerja> <Nama ASN> <Tanggal Lupa Presensi>.pdf
    // Kita ubah format tanggal di nama file jadi DD-MM-YYYY agar lebih umum (opsional)
    const tglParts = formData.Tanggal.split('-'); // [YYYY, MM, DD]
    const tglFormatted = `${tglParts[2]}-${tglParts[1]}-${tglParts[0]}`; // DD-MM-YYYY
    
    const newFileName = `${formData['Unit Kerja']} ${formData['Nama ASN']} ${tglFormatted}.pdf`;
    
    // 5. Simpan File ke Folder Bulan
    const file = targetFolder.createFile(blob).setName(newFileName);
    const fileUrl = file.getUrl();

    // 6. Simpan Data ke Spreadsheet "Data Lupa"
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DB);
    let sheet = ss.getSheetByName("Data Lupa");
    
    if (!sheet) {
       sheet = ss.insertSheet("Data Lupa");
       sheet.appendRow(["Timestamp", "Unit Kerja", "Nama ASN", "NIP", "Tanggal Lupa", "Jenis Presensi", "Jam", "Komulatif", "Link Dokumen"]);
    }

    const jamLengkap = `${formData.Jam.toString().padStart(2, '0')}:${formData.Menit.toString().padStart(2, '0')}`;

    const newRow = [
      new Date(),
      formData['Unit Kerja'],
      formData['Nama ASN'],
      "'" + formData.NIP, 
      formData.Tanggal, // Di database tetap simpan format YYYY-MM-DD agar mudah disortir
      formData['Jenis Presensi'],
      jamLengkap,
      formData.Komulatif,
      fileUrl
    ];

    sheet.appendRow(newRow);
    
    return "Data Lupa Presensi berhasil disimpan!";

  } catch (e) {
    return handleError('processSiabaLupaForm', e);
  }
}

function getSiabaLupaOptions() {
  const cacheKey = 'siaba_lupa_options_v1';
  return getCachedData(cacheKey, function() {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DB);
      const sheet = ss.getSheetByName("Database");
      if (!sheet) throw new Error("Sheet 'Database' tidak ditemukan.");

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];
      
      // Ambil kolom A(NIP), B(Nama), C(Unit Kerja)
      const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      
      return data.map(row => ({
        nip: String(row[0]).trim(),
        nama: String(row[1]).trim(),
        unit: String(row[2]).trim()
      })).filter(item => item.nama !== "" && item.unit !== "");

    } catch (e) {
      throw new Error(`Gagal mengambil data pegawai: ${e.message}`);
    }
  });
}

/**
 * Memproses Upload Surat Pernyataan
 * Fitur: Auto-Folder Bulan & Rename File
 */
function processSiabaLupaForm(formData) {
  try {
    // 1. Validasi Folder Utama
    const folderId = FOLDER_CONFIG.SIABA_LUPA;
    if (!folderId || folderId === "GANTI_DENGAN_ID_FOLDER_DRIVE_UNTUK_SIMPAN_SURAT") {
       throw new Error("Folder penyimpanan belum dikonfigurasi (FOLDER_CONFIG.SIABA_LUPA).");
    }

    const mainFolder = DriveApp.getFolderById(folderId);

    // 2. Persiapan Variabel Tanggal (Dipakai untuk Folder, Nama File, dan Spreadsheet)
    const tanggalLupa = new Date(formData.Tanggal); // Objek Date dari input YYYY-MM-DD
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    
    // Pecah string tanggal "YYYY-MM-DD" menjadi bagian-bagian
    const tglParts = formData.Tanggal.split('-'); // [0]=YYYY, [1]=MM, [2]=DD
    const tglFormatted = `${tglParts[2]}-${tglParts[1]}-${tglParts[0]}`; // Format DD-MM-YYYY

    // 3. Tentukan Folder Bulan (Contoh: "November 2025")
    const folderName = `${monthNames[tanggalLupa.getMonth()]} ${tanggalLupa.getFullYear()}`;
    const targetFolder = getOrCreateFolder(mainFolder, folderName);

    // 4. Proses File (Rename & Upload)
    const fileData = formData.fileData;
    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    
    // Format Nama File: <Unit Kerja> <Nama ASN> <Tanggal>.pdf
    // Menggunakan tglFormatted (DD-MM-YYYY) agar seragam
    const newFileName = `${formData['Unit Kerja']} ${formData['Nama ASN']} ${tglFormatted}.pdf`;
    
    const file = targetFolder.createFile(blob).setName(newFileName);
    const fileUrl = file.getUrl();

    // 5. Simpan Data ke Spreadsheet "Data Lupa"
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DB);
    let sheet = ss.getSheetByName("Data Lupa");
    if (!sheet) {
       sheet = ss.insertSheet("Data Lupa");
       // Buat Header jika sheet baru dibuat
       sheet.appendRow(["Timestamp", "Unit Kerja", "Nama ASN", "NIP", "Tanggal Lupa", "Jenis Presensi", "Jam", "Komulatif", "Link Dokumen"]);
    }

    const jamLengkap = `${formData.Jam.toString().padStart(2, '0')}:${formData.Menit.toString().padStart(2, '0')}`;
    
    const newRow = [
      new Date(),                   // Kolom A: Waktu Unggah
      formData['Unit Kerja'],       // Kolom B
      formData['Nama ASN'],         // Kolom C
      "'" + formData.NIP,           // Kolom D (pakai petik agar jadi teks)
      tglFormatted,                 // Kolom E: Tanggal Lupa (Format DD-MM-YYYY)
      formData['Jenis Presensi'],   // Kolom F
      jamLengkap,                   // Kolom G
      formData.Komulatif,           // Kolom H
      fileUrl                       // Kolom I
    ];
    
    sheet.appendRow(newRow);
    
    return "Berhasil! Surat pernyataan telah disimpan.";

  } catch (e) {
    return handleError('processSiabaLupaForm', e);
  }
}

function getSiabaLupaCekData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    
    // 1. Validasi ID Spreadsheet
    if (!SPREADSHEET_IDS || !SPREADSHEET_IDS.SIABA_DB) {
        throw new Error("Konfigurasi ID Spreadsheet (SIABA_DB) tidak ditemukan.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DB);
    const sheet = ss.getSheetByName("Data Lupa");
    
    if (!sheet || sheet.getLastRow() < 2) {
       return { headers: [], rows: [] };
    }

    const range = sheet.getDataRange();
    // AMBIL DUA JENIS DATA:
    const rawValues = range.getValues();          // Untuk SORTING (Objek Date Asli)
    const displayValues = range.getDisplayValues(); // Untuk TAMPILAN (String Terformat)

    const headers = displayValues[0].map(h => String(h).trim());

    // Gabungkan data menjadi array objek agar Row & Display tetap sinkron saat disortir
    // Kita mulai dari indeks 1 (baris ke-2) untuk melewati header
    let combinedRows = [];
    for (let i = 1; i < rawValues.length; i++) {
        combinedRows.push({
            raw: rawValues[i],
            display: displayValues[i]
        });
    }

    // 3. Temukan Indeks Kolom
    const IDX = {
      TAHUN: headers.indexOf("Tahun"),
      BULAN: headers.indexOf("Bulan"),
      UNIT: headers.indexOf("Unit Kerja"),
      // Data Output
      NAMA: headers.indexOf("Nama ASN"),
      NIP: headers.indexOf("NIP"),
      PERIODE: headers.indexOf("Periode"),
      KOMULATIF: headers.indexOf("Komulatif"),
      TGL_LUPA: headers.indexOf("Tanggal Lupa"), 
      PRESENSI: headers.indexOf("Jenis Presensi"),
      JAM: headers.indexOf("Jam Lupa"), 
      FILE: headers.indexOf("File"),
      ACT: headers.indexOf("Act"),
      KET: headers.indexOf("Keterangan"),
      TIMESTAMP: headers.indexOf("Waktu Unggah") // Kolom A biasanya
    };
    
    // Fallback pencarian kolom jika nama tidak persis
    if (IDX.TIMESTAMP === -1) IDX.TIMESTAMP = 0; // Default kolom A
    if (IDX.KOMULATIF === -1) IDX.KOMULATIF = headers.findIndex(h => h.includes("Komulatif"));
    if (IDX.FILE === -1) IDX.FILE = headers.findIndex(h => h.toLowerCase().includes("file") || h.toLowerCase().includes("dokumen") || h.toLowerCase().includes("surat"));

    // 4. Proses Filter
    const filteredData = combinedRows.filter(item => {
      const row = item.display; // Gunakan display values untuk filter string
      
      const rowUnit = String(row[IDX.UNIT] || "").trim();
      const rowTahun = String(row[IDX.TAHUN] || "").trim();
      const rowBulan = String(row[IDX.BULAN] || "").trim();

      // Filter Unit Kerja
      const matchUnit = (unitKerja === "Semua" || rowUnit === unitKerja);
      if (!matchUnit) return false;

      // Jika kolom Tahun/Bulan tidak ada di sheet, loloskan saja (agar data tampil)
      if (IDX.TAHUN === -1 || IDX.BULAN === -1) return true;

      // Filter Tahun
      if (rowTahun !== String(tahun)) return false;

      // Filter Bulan (Nama Bulan)
      // Kita samakan dengan input filter yang berupa Nama Bulan ("Januari")
      if (rowBulan.toLowerCase() !== bulan.toLowerCase()) {
         // Cek juga jika di sheet isinya Angka (1-12)
         const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
         const bulanIndex = monthNames.indexOf(bulan) + 1;
         if (rowBulan !== String(bulanIndex)) {
             return false; 
         }
      }

      return true;
    });

    // 5. SORTING: Berdasarkan Tanggal Unggah (Timestamp) DESCENDING (Terbaru di Atas)
    // Menggunakan item.raw untuk mendapatkan Objek Date yang akurat
    filteredData.sort((a, b) => {
        const dateA = a.raw[IDX.TIMESTAMP];
        const dateB = b.raw[IDX.TIMESTAMP];
        
        // Pastikan nilai adalah Date valid, jika tidak anggap 0
        const timeA = (dateA instanceof Date) ? dateA.getTime() : 0;
        const timeB = (dateB instanceof Date) ? dateB.getTime() : 0;

        return timeB - timeA; // Besar ke Kecil (Terbaru ke Terlama)
    });

    // 6. Map ke Format Output
    const outputRows = filteredData.map(item => {
      const row = item.display; // Gunakan display values untuk output ke tabel
      return {
        "Nama ASN": row[IDX.NAMA] || "-",
        "NIP": row[IDX.NIP] || "-",
        "Periode": (IDX.PERIODE > -1) ? (row[IDX.PERIODE] || "-") : "-",
        "Komulatif": (IDX.KOMULATIF > -1) ? (row[IDX.KOMULATIF] || "-") : "-",
        "Tanggal": (IDX.TGL_LUPA > -1) ? (row[IDX.TGL_LUPA] || "-") : "-",
        "Presensi": (IDX.PRESENSI > -1) ? (row[IDX.PRESENSI] || "-") : "-",
        "Jam": (IDX.JAM > -1) ? (row[IDX.JAM] || "-") : "-",
        "Dokumen": (IDX.FILE > -1) ? (row[IDX.FILE] || "") : "",
        "Cek": (IDX.ACT > -1) ? (row[IDX.ACT] || "Belum") : "-",
        "Keterangan": (IDX.KET > -1) ? (row[IDX.KET] || "-") : "-",
        "Tanggal Unggah": (IDX.TIMESTAMP > -1) ? (row[IDX.TIMESTAMP] || "-") : "-"
      };
    });

    const desiredHeaders = [
      "Nama ASN", "NIP", "Periode", "Komulatif", "Tanggal", 
      "Presensi", "Jam", "Dokumen", "Cek", "Keterangan", "Tanggal Unggah"
    ];

    return { headers: desiredHeaders, rows: outputRows };

  } catch (e) {
    return handleError('getSiabaLupaCekData', e);
  }
}

function getSiabaLupaRekapOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DB);
    const sheet = ss.getSheetByName("Rekap Lupa Script");
    if (!sheet || sheet.getLastRow() < 2) return { tahun: [], unit: [] };

    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1); // Lewati header

    // Indeks: Unit Kerja (Col C -> 2), Tahun (Col Q -> 16)
    const idxUnit = 2;
    const idxTahun = 16;

    const years = new Set();
    const units = new Set();

    rows.forEach(row => {
      if (row[idxTahun]) years.add(String(row[idxTahun]));
      if (row[idxUnit]) units.add(String(row[idxUnit]));
    });

    return {
      tahun: Array.from(years).sort().reverse(),
      unit: Array.from(units).sort()
    };
  } catch (e) {
    return handleError('getSiabaLupaRekapOptions', e);
  }
}

/**
 * Mengambil Data Rekap Lupa Presensi sesuai Filter
 */
function getSiabaLupaRekapData(filters) {
  try {
    const { tahun, unitKerja } = filters;
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DB);
    const sheet = ss.getSheetByName("Rekap Lupa Script");
    
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues(); // Ambil nilai mentah (angka)
    const rows = data.slice(1); // Data tanpa header

    // Mapping Indeks Kolom (Sesuai Permintaan)
    // NIP(A=0), Nama(B=1), Unit(C=2), Jan(D=3)...Des(O=14), Jml(P=15), Thn(Q=16)
    const idx = {
      NIP: 0, NAMA: 1, UNIT: 2,
      JAN: 3, FEB: 4, MAR: 5, APR: 6, MEI: 7, JUN: 8,
      JUL: 9, AGT: 10, SEP: 11, OKT: 12, NOV: 13, DES: 14,
      JUMLAH: 15, TAHUN: 16
    };

    // Filter Data
    const filteredRows = rows.filter(row => {
      const rowTahun = String(row[idx.TAHUN]);
      const rowUnit = String(row[idx.UNIT]);

      const matchTahun = (rowTahun === String(tahun));
      const matchUnit = (unitKerja === "Semua" || rowUnit === unitKerja);

      return matchTahun && matchUnit;
    });

    // Urutkan berdasarkan Jumlah (Descending: Terbanyak di atas)
    // Jika jumlah sama, baru urutkan berdasarkan Nama
    filteredRows.sort((a, b) => {
        const diff = (parseFloat(b[idx.JUMLAH]) || 0) - (parseFloat(a[idx.JUMLAH]) || 0);
        if (diff !== 0) return diff;
        return String(a[idx.NAMA]).localeCompare(String(b[idx.NAMA]));
    });

    // Map ke format array sederhana untuk dikirim ke klien
    return filteredRows.map(row => [
      row[idx.NAMA], // 0
      row[idx.NIP],  // 1
      row[idx.JAN], row[idx.FEB], row[idx.MAR], row[idx.APR], row[idx.MEI], row[idx.JUN], // 2-7
      row[idx.JUL], row[idx.AGT], row[idx.SEP], row[idx.OKT], row[idx.NOV], row[idx.DES], // 8-13
      row[idx.JUMLAH] // 14
    ]);

  } catch (e) {
    return handleError('getSiabaLupaRekapData', e);
  }
}

function getSiabaSalahOptions() {
  try {
    // Gunakan ID langsung atau via config
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SALAH_DB);
    const sheet = ss.getSheetByName("Database");
    
    if (!sheet || sheet.getLastRow() < 2) return [];

    // Ambil data Kolom A(Unit), B(NIP), C(Nama)
    // getRange(baris, kolom, numBaris, numKolom) -> A2:C
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();

    // Map ke array objek yang bersih
    return data.map(row => ({
      unit: String(row[0]).trim(), // Kolom A
      nip: String(row[1]).trim(),  // Kolom B
      nama: String(row[2]).trim()  // Kolom C
    })).filter(item => item.unit && item.nama); // Filter baris kosong

  } catch (e) {
    return handleError('getSiabaSalahOptions', e);
  }
}

/**
 * Memproses Form Pengajuan Salah Presensi
 * Simpan ke: Spreadsheet SIABA_SALAH_DB sheet "Salah Presensi"
 */
function processSiabaSalahForm(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SALAH_DB);
    let sheet = ss.getSheetByName("Salah Presensi");
    
    // Buat sheet jika belum ada
    if (!sheet) {
      sheet = ss.insertSheet("Salah Presensi");
      sheet.appendRow([
        "Tanggal Pengajuan", "Unit Kerja", "Nama ASN", "NIP", 
        "Tanggal Salah", "Jam", "Tercatat", "Seharusnya"
      ]);
    }

    // Format Tanggal Pengajuan (Timestamp)
    const timestamp = new Date();

    // PERBAIKAN: Format Tanggal Salah (Input: YYYY-MM-DD -> Output: DD-MM-YYYY)
    // Menggunakan tanda strip (-) sesuai permintaan
    const tglParts = formData.tanggal.split('-'); // [YYYY, MM, DD]
    const tglFormatted = `${tglParts[2]}-${tglParts[1]}-${tglParts[0]}`;

    // Format Jam (HH:mm) - Memastikan 2 digit
    const jamLengkap = `${formData.jam.toString().padStart(2, '0')}:${formData.menit.toString().padStart(2, '0')}`;

    const newRow = [
      timestamp,              // A: Tanggal Pengajuan
      formData.unitKerja,     // B: Unit Kerja
      formData.namaAsn,       // C: Nama ASN
      "'" + formData.nip,     // D: NIP (pakai petik agar jadi teks)
      tglFormatted,           // E: Tanggal (DD-MM-YYYY)
      jamLengkap,             // F: Jam (HH:mm)
      formData.tercatat,      // G: Presensi Tercatat
      formData.benar          // H: Presensi yang Benar
    ];

    sheet.appendRow(newRow);
    
    return "Pengajuan berhasil dikirim.";

  } catch (e) {
    return handleError('processSiabaSalahForm', e);
  }
}

function getSiabaSalahFilterOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SALAH_DB);
    const sheet = ss.getSheetByName("Salah Presensi");
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { tahun: [], bulan: [], unit: [] };
    }

    // Ambil data dari baris 2 sampai akhir
    const data = sheet.getDataRange().getDisplayValues().slice(1);
    
    // Indeks Kolom (0-based): Unit(B=1), Tahun(L=11), Bulan(M=12)
    const idxUnit = 1;
    const idxTahun = 11;
    const idxBulan = 12;

    const years = new Set();
    const months = new Set();
    const units = new Set();

    data.forEach(row => {
      if (row[idxTahun]) years.add(String(row[idxTahun]).trim());
      if (row[idxBulan]) months.add(String(row[idxBulan]).trim());
      if (row[idxUnit]) units.add(String(row[idxUnit]).trim());
    });

    // Sortir Bulan (Januari - Desember)
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const sortedMonths = Array.from(months).sort((a, b) => {
       // Coba cek apakah bulan berupa angka atau nama
       const idxA = monthNames.indexOf(a);
       const idxB = monthNames.indexOf(b);
       if (idxA !== -1 && idxB !== -1) return idxA - idxB;
       return a.localeCompare(b);
    });

    return {
      tahun: Array.from(years).sort().reverse(),
      bulan: sortedMonths,
      unit: Array.from(units).sort()
    };

  } catch (e) {
    return handleError('getSiabaSalahFilterOptions', e);
  }
}

/**
 * Mengambil Data Tabel Cek Salah Presensi
 */
function getSiabaSalahCekData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_SALAH_DB);
    const sheet = ss.getSheetByName("Salah Presensi");
    
    if (!sheet || sheet.getLastRow() < 2) {
       return { headers: [], rows: [] };
    }

    const range = sheet.getDataRange();
    const rawValues = range.getValues(); // Untuk sorting tanggal
    const displayValues = range.getDisplayValues(); // Untuk tampilan
    
    const rows = [];
    // Gabungkan raw dan display, mulai baris 2 (index 1)
    for (let i = 1; i < rawValues.length; i++) {
        rows.push({
            raw: rawValues[i],
            display: displayValues[i]
        });
    }

    // Peta Indeks Kolom (0-based)
    // A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, M=12
    const IDX = {
      TIMESTAMP: 0, // A: Tanggal Pengajuan
      UNIT: 1,      // B: Unit Kerja
      NAMA: 2,      // C: Nama ASN
      NIP: 3,       // D: NIP
      TGL_SALAH: 4, // E: Tanggal Salah
      TERCATAT: 6,  // G: Tercatat
      REVISI: 7,    // H: Seharusnya
      CEK: 9,       // J: Cek/Act
      KET: 10,      // K: Keterangan
      TAHUN: 11,    // L: Tahun
      BULAN: 12     // M: Bulan
    };

    // Filter Data
    const filteredRows = rows.filter(item => {
      const row = item.display;
      const rowTahun = String(row[IDX.TAHUN] || "").trim();
      const rowBulan = String(row[IDX.BULAN] || "").trim();
      const rowUnit = String(row[IDX.UNIT] || "").trim();

      const matchTahun = (rowTahun === String(tahun));
      const matchBulan = (rowBulan === String(bulan));
      const matchUnit = (unitKerja === "Semua" || rowUnit === unitKerja);

      return matchTahun && matchBulan && matchUnit;
    });

    // Sorting: Tanggal Pengajuan (Timestamp) Terbaru di Atas
    filteredRows.sort((a, b) => {
        const dateA = a.raw[IDX.TIMESTAMP];
        const dateB = b.raw[IDX.TIMESTAMP];
        const timeA = (dateA instanceof Date) ? dateA.getTime() : 0;
        const timeB = (dateB instanceof Date) ? dateB.getTime() : 0;
        return timeB - timeA;
    });

    // Mapping Output
    const outputRows = filteredRows.map(item => {
      const row = item.display;
      return {
        "Nama ASN": row[IDX.NAMA] || "-",
        "NIP": row[IDX.NIP] || "-",
        "Tanggal": row[IDX.TGL_SALAH] || "-",
        "Tercatat": row[IDX.TERCATAT] || "-",
        "Revisi": row[IDX.REVISI] || "-",
        "Cek": row[IDX.CEK] || "Proses", // Default Proses jika kosong
        "Keterangan": row[IDX.KET] || "-",
        "Tanggal Pengajuan": row[IDX.TIMESTAMP] || "-"
      };
    });

    const headers = ["Nama ASN", "NIP", "Tanggal", "Tercatat", "Revisi", "Cek", "Keterangan", "Tanggal Pengajuan"];

    return { headers: headers, rows: outputRows };

  } catch (e) {
    return handleError('getSiabaSalahCekData', e);
  }
}

function getSiabaDinasOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DINAS_DB);
    const sheet = ss.getSheetByName("Database");
    if (!sheet || sheet.getLastRow() < 2) return [];

    // Ambil Kolom A(Unit), B(NIP), C(Nama)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();

    // Return data bersih
    return data.map(row => ({
      unit: String(row[0]).trim(),
      nip: String(row[1]).trim(),
      nama: String(row[2]).trim()
    })).filter(item => item.unit && item.nama);

  } catch (e) {
    return handleError('getSiabaDinasOptions', e);
  }
}

/**
 * Memproses Form Unggah SPT
 * Simpan ke: Spreadsheet SIABA_DINAS_DB sheet "SPPD"
 * Update: Auto ID (SPTXXXXX) di Kolom O, Update di Kolom P
 */
function processSiabaDinasForm(formData) {
  try {
    const folderId = FOLDER_CONFIG.SIABA_DINAS;
    const mainFolder = DriveApp.getFolderById(folderId);

    // 1. Generate ID Unik (SPT + 5 digit acak)
    const uniqueID = "SPT" + Math.floor(10000 + Math.random() * 90000);

    // 2. Simpan File SPT
    const fileData = formData.fileData;
    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    
    const safeNomorSPT = formData.nomorSPT.replace(/[^a-zA-Z0-9]/g, '_');
    const safeUnit = formData.unitKerja.replace(/[^a-zA-Z0-9 ]/g, '').trim();
    const newFileName = `SPT_${safeUnit}_${safeNomorSPT}.pdf`;
    
    const file = mainFolder.createFile(blob).setName(newFileName);
    const fileUrl = file.getUrl();

    // 3. Simpan Data
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DINAS_DB);
    let sheet = ss.getSheetByName("SPPD");
    
    if (!sheet) {
      sheet = ss.insertSheet("SPPD");
      sheet.appendRow([
          "Tgl Pengajuan", "Jenis PD", "No SPT", "Tgl SPT", "Tgl Mulai", "Tgl Selesai", 
          "Lokasi", "Kegiatan", "Unit Kerja", "Jml Peserta", "Nama ASN", "NIP", "File", "Jenis Dokumen", "ID", "Update"
      ]);
    }

    const formatDate = (dateStr) => {
        if (!dateStr) return "";
        const p = dateStr.split('-'); 
        return `${p[2]}-${p[1]}-${p[0]}`;
    };

    const tglPengajuan = new Date();
    const tglSPT = formatDate(formData.tanggalSPT);
    const tglMulai = formatDate(formData.tanggalMulai);
    const tglSelesai = formatDate(formData.tanggalSelesai);
    
    const pesertaCount = parseInt(formData.peserta) || 1;
    let listNama = [];
    let listNIP = [];

    for (let i = 1; i <= pesertaCount; i++) {
        const nama = formData[`nama_${i}`];
        const nip = formData[`nip_${i}`];
        if (nama && nip) {
            listNama.push(nama);
            listNIP.push(nip); 
        }
    }

    const combinedNama = listNama.join("<br>"); 
    const combinedNIP = listNIP.join("<br>");

    const newRow = [
        tglPengajuan,           // A
        formData.jenisPD,       // B
        formData.nomorSPT,      // C
        tglSPT,                 // D
        tglMulai,               // E
        tglSelesai,             // F
        formData.lokasi,        // G
        formData.kegiatan,      // H
        formData.unitKerja,     // I
        "'" + formData.peserta, // J
        combinedNama,           // K
        combinedNIP,            // L
        fileUrl,                // M
        formData.jenisDokumen,  // N (Final/Sementara)
        uniqueID,               // O (ID BARU)
        ""                      // P (Update kosong dulu)
    ];

    sheet.appendRow(newRow);

    return "Data berhasil disimpan. ID: " + uniqueID;

  } catch (e) {
    return handleError('processSiabaDinasForm', e);
  }
}

function getSiabaDinasFilterOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DINAS_DB);
    const sheet = ss.getSheetByName("SPPD");
    if (!sheet || sheet.getLastRow() < 2) return { tahun: [], bulan: [], unit: [] };

    // Ambil data Tgl Pengajuan (A=0) dan Unit Kerja (I=8)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    
    const years = new Set();
    const months = new Set();
    const units = new Set();
    const mNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    data.forEach(row => {
      const tgl = row[0]; // Tgl Pengajuan (Date Object)
      const unit = row[8]; // Unit Kerja

      if (tgl instanceof Date) {
        years.add(tgl.getFullYear());
        months.add(mNames[tgl.getMonth()]);
      }
      if (unit) units.add(String(unit).trim());
    });

    return {
      tahun: Array.from(years).sort().reverse(),
      bulan: Array.from(months), // Sorting bulan bisa ditangani di client
      unit: Array.from(units).sort()
    };
  } catch (e) {
    return handleError('getSiabaDinasFilterOptions', e);
  }
}

/**
 * Mengambil Data Tabel Cek Salah Presensi
 * PERBAIKAN: Menggunakan Raw Values untuk Filter Tanggal agar Akurat
 * & Penanganan Data Kosong (Robust)
 */
function getSiabaDinasCekData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DINAS_DB);
    const sheet = ss.getSheetByName("SPPD");
    
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

    const range = sheet.getDataRange();
    const rawValues = range.getValues(); // Untuk Filter & Sorting (Objek Asli)
    const displayValues = range.getDisplayValues(); // Untuk Tampilan Tabel (Teks)
    
    // Gabungkan data mulai dari baris ke-2 (indeks 1)
    const rows = [];
    for (let i = 1; i < rawValues.length; i++) {
        rows.push({
            raw: rawValues[i],
            display: displayValues[i],
            index: i + 1 // Simpan indeks baris asli (1-based)
        });
    }

    // Mapping Index (Sesuai Struktur Sheet SPPD)
    // A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, M=12, N=13, O=14, P=15, Q=16, R=17
    const IDX = {
      TGL_AJU: 0,    // A
      JENIS_PD: 1,   // B
      NO_SPT: 2,     // C
      TGL_SPT: 3,    // D
      TGL_MULAI: 4,  // E
      TGL_SELESAI: 5,// F
      LOKASI: 6,     // G
      KEGIATAN: 7,   // H
      UNIT: 8,       // I
      NAMA: 10,      // K
      NIP: 11,       // L
      DOKUMEN: 12,   // M
      JENIS_DOK: 13, // N
      ID: 14,        // O
      UPDATE: 15,    // P
      CEK: 16,       // Q
      KET: 17        // R
    };

    const mNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    // Filter Data
    const filtered = rows.filter(item => {
      const rawRow = item.raw;
      const dispRow = item.display;

      // Ambil Tanggal dari Data Mentah (Objek Date)
      const dateObj = rawRow[IDX.TGL_AJU];
      
      // Validasi Tanggal: Harus objek Date valid
      if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return false; 

      const rowYear = dateObj.getFullYear().toString();
      const rowMonth = mNames[dateObj.getMonth()];
      // Ambil unit kerja dengan aman (hindari undefined)
      const rowUnit = String(dispRow[IDX.UNIT] || "").trim();

      const matchTahun = (rowYear === String(tahun));
      const matchBulan = (rowMonth === String(bulan));
      const matchUnit = (unitKerja === "Semua" || rowUnit === unitKerja);

      return matchTahun && matchBulan && matchUnit;
    });

    // Sorting: Tanggal Pengajuan Terbaru (Descending)
    filtered.sort((a, b) => {
       const dateA = new Date(a.raw[IDX.TGL_AJU] || 0);
       const dateB = new Date(b.raw[IDX.TGL_AJU] || 0);
       return dateB - dateA; 
    });

    // Mapping Output (Gunakan Display Values)
    const result = filtered.map(item => {
        const r = item.display;
        
        // Helper aman untuk ambil nilai (return string kosong jika undefined)
        const getVal = (idx) => (r[idx] === undefined || r[idx] === null) ? "" : r[idx];

        return {
            "_rowIndex": item.index, 
            "Nama ASN": getVal(IDX.NAMA),
            "NIP": getVal(IDX.NIP),
            "Jenis PD": getVal(IDX.JENIS_PD),
            "Nomor SPT": getVal(IDX.NO_SPT),
            "Tanggal SPT": getVal(IDX.TGL_SPT),
            "Tanggal Mulai": getVal(IDX.TGL_MULAI),
            "Tanggal Selesai": getVal(IDX.TGL_SELESAI),
            "Lokasi Tujuan": getVal(IDX.LOKASI),
            "Kegiatan": getVal(IDX.KEGIATAN),
            "Dokumen": getVal(IDX.DOKUMEN),
            "Jenis Dokumen": getVal(IDX.JENIS_DOK),
            "Cek": getVal(IDX.CEK),
            "Keterangan": getVal(IDX.KET),
            "Tanggal Pengajuan": getVal(IDX.TGL_AJU),
            "Update": getVal(IDX.UPDATE),
            "ID": getVal(IDX.ID) // Penting untuk tombol Edit
        };
    });

    const headers = [
        "Nama ASN", "NIP", "Jenis PD", "Nomor SPT", "Tanggal SPT", "Tanggal Mulai", 
        "Tanggal Selesai", "Lokasi Tujuan", "Kegiatan", "Dokumen", "Jenis Dokumen", 
        "Cek", "Keterangan", "Aksi", "Tanggal Pengajuan", "Update"
    ];

    return { headers: headers, rows: result };

  } catch (e) {
    return handleError('getSiabaDinasCekData', e);
  }
}

function deleteSiabaDinasData(rowIndex, deleteCode) {
  try {
    // 1. Validasi Kode Hapus (Format: yyyyMMdd hari ini)
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DINAS_DB);
    const sheet = ss.getSheetByName("SPPD");
    
    // 2. Hapus File Fisik di Drive (Opsional, agar bersih)
    const fileUrl = sheet.getRange(rowIndex, 13).getValue(); // Kolom M (Dokumen)
    if (fileUrl && String(fileUrl).includes("drive.google.com")) {
        try {
            const fileId = String(fileUrl).match(/[-\w]{25,}/);
            if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch(e) { 
            Logger.log("Gagal hapus file: " + e.message); 
        }
    }

    // 3. Hapus Baris Data
    sheet.deleteRow(rowIndex);
    
    return "Data berhasil dihapus.";
  } catch (e) {
    return handleError('deleteSiabaDinasData', e);
  }
}

/**
 * Mengambil Data Edit Berdasarkan ID (Versi Robust/Aman)
 */
function getSiabaDinasEditData(idSpt) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DINAS_DB);
    const sheet = ss.getSheetByName("SPPD");
    
    // Pastikan kita mengambil area data yang cukup luas (sampai kolom P/16)
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 16); 
    
    if (lastRow < 2) throw new Error("Database kosong.");

    // Ambil semua data (mulai baris 2)
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = range.getValues(); // Raw values (Date object, dll)
    const displayValues = range.getDisplayValues(); // Display values (String terformat)
    
    let rowIndex = -1;
    let rowData = [];
    let displayRow = [];

    // Loop cari ID (Kolom O = Index 14)
    for (let i = 0; i < values.length; i++) {
        // Gunakan String() agar aman dari undefined
        const currentId = String(values[i][14] || "").trim();
        if (currentId === String(idSpt).trim()) {
            rowIndex = i + 2; // +2 karena mulai dari baris 2
            rowData = values[i];
            displayRow = displayValues[i];
            break;
        }
    }

    if (rowIndex === -1) throw new Error("Data tidak ditemukan (ID tidak cocok).");

    // Helper: Ambil nilai dengan aman (hindari undefined)
    const getVal = (arr, idx) => (arr && arr[idx] !== undefined) ? arr[idx] : "";

    // Helper: Format Tanggal untuk Input HTML (YYYY-MM-DD)
    const toInputDate = (dateObj) => {
        if (dateObj instanceof Date) {
            return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        // Fallback jika tersimpan sebagai string DD-MM-YYYY
        const s = String(dateObj);
        if (s.includes('-') && s.split('-')[0].length === 2) {
             const p = s.split('-');
             return `${p[2]}-${p[1]}-${p[0]}`;
        }
        return "";
    };

    return {
        rowIndex: rowIndex,
        id: idSpt,
        jenisPD: getVal(rowData, 1),
        nomorSPT: String(getVal(rowData, 2)), // Pastikan String
        tanggalSPT: toInputDate(getVal(rowData, 3)),
        tanggalMulai: toInputDate(getVal(rowData, 4)),
        tanggalSelesai: toInputDate(getVal(rowData, 5)),
        lokasi: getVal(rowData, 6),
        kegiatan: getVal(rowData, 7),
        unitKerja: getVal(rowData, 8),
        peserta: getVal(rowData, 9), 
        namaASN: String(getVal(displayRow, 10)), // Gunakan display row untuk teks
        nip: String(getVal(displayRow, 11)),     
        fileUrl: getVal(rowData, 12),
        jenisDokumen: getVal(rowData, 13)
    };

  } catch (e) {
    return handleError('getSiabaDinasEditData', e);
  }
}

/**
 * Update Full Data Berdasarkan ID (Versi Perbaikan)
 */
function updateSiabaDinasFullData(formData) {
  try {
    const idSpt = formData.idSpt;
    
    // Validasi ID
    if (!idSpt || String(idSpt).trim() === "") {
        throw new Error("Gagal menyimpan: ID SPT kosong atau tidak valid.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DINAS_DB);
    const sheet = ss.getSheetByName("SPPD");
    const folderId = FOLDER_CONFIG.SIABA_DINAS;

    // --- LOGIKA PENCARIAN BARIS (REVISI) ---
    // Ambil seluruh kolom ID (O)
    const idData = sheet.getRange("O:O").getDisplayValues().flat();
    
    // Cari index (tambah 1 karena array mulai dari 0, spreadsheet baris mulai dari 1)
    // Gunakan String comparison yang ketat
    let rowIndex = -1;
    for(let i=0; i<idData.length; i++) {
        if(String(idData[i]).trim() === String(idSpt).trim()) {
            rowIndex = i + 1;
            break;
        }
    }

    // Jika tidak ketemu, rowIndex tetap -1
    if (rowIndex < 1) {
        throw new Error("Data dengan ID '" + idSpt + "' tidak ditemukan di database. Mohon refresh halaman.");
    }
    // ----------------------------------------

    // 1. Handle File (Hanya jika ada file baru)
    let fileUrl = formData.existingFileUrl;
    if (formData.fileData && formData.fileData.data) {
        // Coba hapus file lama (opsional)
        try {
            const oldUrl = sheet.getRange(rowIndex, 13).getValue(); // Kolom M
            if (oldUrl && String(oldUrl).includes("drive.google.com")) {
                const fileId = String(oldUrl).match(/[-\w]{25,}/);
                if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
            }
        } catch(e) {
            // Abaikan error hapus file lama
        }

        // Upload baru
        const decodedData = Utilities.base64Decode(formData.fileData.data);
        const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, formData.fileData.fileName);
        
        const safeNomorSPT = String(formData.nomorSPT).replace(/[^a-zA-Z0-9]/g, '_');
        const safeUnit = String(formData.unitKerja).replace(/[^a-zA-Z0-9 ]/g, '').trim();
        const newFileName = `SPT_${safeUnit}_${safeNomorSPT}.pdf`;
        
        const file = DriveApp.getFolderById(folderId).createFile(blob).setName(newFileName);
        fileUrl = file.getUrl();
    }

    // 2. Format Data
    const formatDate = (dateStr) => {
        if (!dateStr) return "";
        const p = dateStr.split('-'); 
        return `${p[2]}-${p[1]}-${p[0]}`;
    };

    // 3. Proses Peserta
    const pesertaCount = parseInt(formData.peserta) || 1;
    let listNama = [];
    let listNIP = [];
    for (let i = 1; i <= pesertaCount; i++) {
        const nama = formData[`nama_${i}`];
        const nip = formData[`nip_${i}`];
        if (nama && nip) {
            listNama.push(nama);
            listNIP.push(nip); 
        }
    }
    const combinedNama = listNama.join("<br>"); 
    const combinedNIP = listNIP.join("<br>");

    // 4. Update Sheet (Range B:P -> Kolom 2 s/d 16)
    // Pastikan ID (Kolom O) tetap ada
    const dataRow = [
        formData.jenisPD,       // B (2)
        formData.nomorSPT,      // C (3)
        formatDate(formData.tanggalSPT), // D (4)
        formatDate(formData.tanggalMulai), // E (5)
        formatDate(formData.tanggalSelesai), // F (6)
        formData.lokasi,        // G (7)
        formData.kegiatan,      // H (8)
        formData.unitKerja,     // I (9)
        "'" + formData.peserta, // J (10)
        combinedNama,           // K (11)
        combinedNIP,            // L (12)
        fileUrl,                // M (13)
        formData.jenisDokumen,  // N (14)
        idSpt,                  // O (15) - Tulis ulang ID agar aman
        new Date()              // P (16) - Timestamp Update
    ];

    // Update 1 baris, 15 kolom (dari B sampai P)
    sheet.getRange(rowIndex, 2, 1, 15).setValues([dataRow]);

    return "Perubahan berhasil disimpan.";

  } catch (e) {
    return handleError('updateSiabaDinasFullData', e);
  }
}

/**
 * Update Full Data Berdasarkan ID (Versi TextFinder - Lebih Stabil)
 */
function updateSiabaDinasFullData(formData) {
  try {
    const idSpt = formData.idSpt;
    
    // Validasi ID
    if (!idSpt || String(idSpt).trim() === "") {
        throw new Error("Gagal menyimpan: ID SPT tidak ditemukan dalam formulir.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_DINAS_DB);
    const sheet = ss.getSheetByName("SPPD");
    const folderId = FOLDER_CONFIG.SIABA_DINAS;

    // --- PENCARIAN BARIS DENGAN TEXTFINDER ---
    // Mencari ID spesifik di Kolom O (agar tidak salah hapus baris lain)
    const columnO = sheet.getRange("O:O");
    const finder = columnO.createTextFinder(idSpt).matchEntireCell(true);
    const result = finder.findNext();

    if (!result) {
        throw new Error("Data dengan ID '" + idSpt + "' tidak ditemukan di database. Mohon refresh halaman.");
    }

    const rowIndex = result.getRow(); // Dapatkan nomor baris secara langsung

    // Validasi Baris (Jangan sampai mengedit header)
    if (rowIndex <= 1) {
        throw new Error("Kesalahan sistem: ID ditemukan di baris header.");
    }
    // ----------------------------------------

    // 1. Handle File (Hanya jika ada file baru)
    let fileUrl = formData.existingFileUrl;
    if (formData.fileData && formData.fileData.data) {
        // Coba hapus file lama (opsional)
        try {
            const oldUrl = sheet.getRange(rowIndex, 13).getValue(); // Kolom M
            if (oldUrl && String(oldUrl).includes("drive.google.com")) {
                const fileId = String(oldUrl).match(/[-\w]{25,}/);
                if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
            }
        } catch(e) {
            Logger.log("Gagal hapus file lama: " + e.message);
        }

        // Upload baru
        const decodedData = Utilities.base64Decode(formData.fileData.data);
        const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, formData.fileData.fileName);
        
        const safeNomorSPT = String(formData.nomorSPT).replace(/[^a-zA-Z0-9]/g, '_');
        const safeUnit = String(formData.unitKerja).replace(/[^a-zA-Z0-9 ]/g, '').trim();
        const newFileName = `SPT_${safeUnit}_${safeNomorSPT}.pdf`;
        
        const file = DriveApp.getFolderById(folderId).createFile(blob).setName(newFileName);
        fileUrl = file.getUrl();
    }

    // 2. Format Data
    const formatDate = (dateStr) => {
        if (!dateStr) return "";
        const p = dateStr.split('-'); 
        return `${p[2]}-${p[1]}-${p[0]}`;
    };

    // 3. Proses Peserta
    const pesertaCount = parseInt(formData.peserta) || 1;
    let listNama = [];
    let listNIP = [];
    for (let i = 1; i <= pesertaCount; i++) {
        const nama = formData[`nama_${i}`];
        const nip = formData[`nip_${i}`];
        if (nama && nip) {
            listNama.push(nama);
            listNIP.push(nip); 
        }
    }
    const combinedNama = listNama.join("<br>"); 
    const combinedNIP = listNIP.join("<br>");

    // 4. Update Sheet (Range B:P -> Kolom 2 s/d 16)
    // Menimpa data lama dengan data baru
    const dataRow = [
        formData.jenisPD,       // B
        formData.nomorSPT,      // C
        formatDate(formData.tanggalSPT), // D
        formatDate(formData.tanggalMulai), // E
        formatDate(formData.tanggalSelesai), // F
        formData.lokasi,        // G
        formData.kegiatan,      // H
        formData.unitKerja,     // I
        "'" + formData.peserta, // J
        combinedNama,           // K
        combinedNIP,            // L
        fileUrl,                // M
        formData.jenisDokumen,  // N
        idSpt,                  // O (ID Tetap)
        new Date()              // P (Waktu Update)
    ];

    // Eksekusi Update
    sheet.getRange(rowIndex, 2, 1, 15).setValues([dataRow]);

    return "Perubahan berhasil disimpan.";

  } catch (e) {
    return handleError('updateSiabaDinasFullData', e);
  }
}

/**
 * Mengambil Opsi Pegawai untuk Form Cuti
 * Sumber: Spreadsheet SIABA_CUTI_DB sheet "database"
 */
function getSiabaCutiOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    const sheet = ss.getSheetByName("database");
    if (!sheet || sheet.getLastRow() < 2) return [];

    // Ambil seluruh data karena kolomnya berjauhan (A sampai AL)
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1); // Lewati header

    // Indeks Kolom (0-based):
    // Nama=A(0), NIP=B(1), Alamat=G(6), Unit=AL(37)
    const idxUnit = 37; // AL
    const idxNama = 0;  // A
    const idxNip = 1;   // B
    const idxAlamat = 6;// G

    return rows.map(row => ({
      unit: String(row[idxUnit]).trim(),
      nama: String(row[idxNama]).trim(),
      nip: String(row[idxNip]).trim(),
      alamat: String(row[idxAlamat]).trim()
    })).filter(item => item.unit && item.nama); // Filter baris kosong

  } catch (e) {
    return handleError('getSiabaCutiOptions', e);
  }
}

/**
 * Memproses Form Pengajuan Cuti (UPDATE: Format Tanggal & Auto ID)
 * Simpan ke: Spreadsheet SIABA_CUTI_DB sheet "Form Ajuan"
 */
function submitSiabaCutiForm(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    let sheet = ss.getSheetByName("Form Ajuan");
    
    if (!sheet) {
      sheet = ss.insertSheet("Form Ajuan");
    }

    // --- PERBAIKAN FORMAT TANGGAL (DD mmmm YYYY) ---
    const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    
    const formatDate = (dateStr) => {
        if (!dateStr) return "";
        // Input dari HTML date picker selalu: YYYY-MM-DD
        const p = dateStr.split('-'); // [YYYY, MM, DD]
        const day = p[2];
        const monthIndex = parseInt(p[1], 10) - 1;
        const year = p[0];
        
        // Output: 23 November 2025
        return `${day} ${monthNames[monthIndex]} ${year}`;
    };
    // -----------------------------------------------

    // Generate ID Unik
    const idCuti = "CUTI" + new Date().getTime();
    const noHpText = "'" + String(formData.noHp).replace(/\D/g, '');

    const newRow = new Array(33).fill("");

    newRow[0] = new Date();               // A: Timestamp
    newRow[1] = formData.unitKerja;       // B
    newRow[2] = formData.namaAsn;         // C
    newRow[3] = "'" + formData.nip;       // D
    newRow[4] = formData.tugas;           // E
    newRow[5] = formData.status;          // F
    newRow[6] = formData.golongan;        // G
    newRow[7] = formData.jenisCuti;       // H
    newRow[8] = formatDate(formData.tglSurat);   // I (Format Baru)
    newRow[9] = formatDate(formData.tglMulai);   // J (Format Baru)
    newRow[10] = formatDate(formData.tglSelesai); // K (Format Baru)
    newRow[11] = formData.jmlHari;        // L
    newRow[12] = formData.alasan;         // M
    newRow[13] = formData.alamatCuti;     // N
    newRow[14] = noHpText;                // O
    newRow[15] = formData.usulan;         // P
    
    // Kolom AG (Index 32) = ID Cuti
    newRow[32] = idCuti;

    sheet.appendRow(newRow);
    
    return "Pengajuan Cuti berhasil dikirim.";

  } catch (e) {
    return handleError('submitSiabaCutiForm', e);
  }
}

/**
 * Mengambil Data Cuti untuk Edit berdasarkan ID (Kolom AG)
 */
function getSiabaCutiEditData(idCuti) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    const sheet = ss.getSheetByName("Form Ajuan");
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Data kosong.");

    // Ambil data kolom A sampai AG (1-33)
    const data = sheet.getRange(2, 1, lastRow - 1, 33).getValues();
    
    let rowData = null;
    let rowIndex = -1;

    // Loop cari ID di kolom AG (Index 32)
    for (let i = 0; i < data.length; i++) {
        if (String(data[i][32]) === String(idCuti)) {
            rowData = data[i];
            rowIndex = i + 2; // 1-based index + header
            break;
        }
    }

    if (!rowData) throw new Error("Data cuti tidak ditemukan.");

    // Helper Date YYYY-MM-DD
    const toInputDate = (val) => {
        if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (typeof val === 'string' && val.includes('-') && val.length===10) {
            // Asumsi format DD-MM-YYYY -> ubah ke YYYY-MM-DD
             const p = val.split('-');
             return `${p[2]}-${p[1]}-${p[0]}`;
        }
        return "";
    };

    return {
        id: idCuti,
        unitKerja: rowData[1],
        namaAsn: rowData[2],
        nip: String(rowData[3]).replace(/'/g, ''), // Hapus petik
        tugas: rowData[4],
        status: rowData[5],
        golongan: rowData[6],
        jenisCuti: rowData[7],
        tglSurat: toInputDate(rowData[8]),
        tglMulai: toInputDate(rowData[9]),
        tglSelesai: toInputDate(rowData[10]),
        jmlHari: rowData[11],
        alasan: rowData[12],
        alamatCuti: rowData[13],
        noHp: String(rowData[14]).replace(/'/g, ''),
        usulan: rowData[15]
    };

  } catch (e) {
    return handleError('getSiabaCutiEditData', e);
  }
}

/**
 * Update Data Cuti
 */
function updateSiabaCutiData(formData) {
  try {
    const idCuti = formData.idCuti;
    if (!idCuti) throw new Error("ID Cuti tidak ditemukan.");

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    const sheet = ss.getSheetByName("Form Ajuan");

    // Cari Baris dengan TextFinder (Lebih Cepat)
    const finder = sheet.getRange("AG:AG").createTextFinder(idCuti).matchEntireCell(true);
    const result = finder.findNext();
    
    if (!result) throw new Error("Data tidak ditemukan di database.");
    const rowIndex = result.getRow();

    // Helper Format Tanggal
    const formatDate = (dateStr) => {
        if (!dateStr) return "";
        const p = dateStr.split('-'); 
        return `${p[2]}-${p[1]}-${p[0]}`;
    };
    
    const noHpText = "'" + String(formData.noHp).replace(/\D/g, '');

    // Update Kolom E sampai P (Index 5 sampai 16 di sheet, kolom 5=E)
    // Data yang dikunci (Unit, Nama, NIP - Kolom B,C,D) tidak diupdate
    const dataToUpdate = [[
        formData.tugas,           // E
        formData.status,          // F
        formData.golongan,        // G
        formData.jenisCuti,       // H
        formatDate(formData.tglSurat),   // I
        formatDate(formData.tglMulai),   // J
        formatDate(formData.tglSelesai), // K
        formData.jmlHari,         // L
        formData.alasan,          // M
        formData.alamatCuti,      // N
        noHpText,                 // O
        formData.usulan           // P
    ]];

    // Tulis ke range E[Row]:P[Row]
    sheet.getRange(rowIndex, 5, 1, 12).setValues(dataToUpdate);
    
    // Opsional: Update Timestamp di Kolom A? (User tidak meminta, jadi biarkan timestamp asli)

    return "Data Cuti berhasil diperbarui.";

  } catch (e) {
    return handleError('updateSiabaCutiData', e);
  }
}

/**
 * Mengambil Opsi Filter untuk Halaman Unduh Formulir Cuti
 * Sumber: Sheet "Form Ajuan"
 * Filter: Tahun (Kolom X/Index 23), Unit Kerja (Kolom B/Index 1)
 * Data mulai baris 3 (Baris 2 berisi rumus)
 */
function getSiabaCutiUnduhFilterOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    const sheet = ss.getSheetByName("Form Ajuan");
    
    if (!sheet || sheet.getLastRow() < 3) return { tahun: [], unit: [] };

    const lastRow = sheet.getLastRow();
    // Ambil data sampai kolom X (Index 23) atau lebih aman secukupnya
    const data = sheet.getRange(3, 1, lastRow - 2, 25).getValues(); 
    
    const years = new Set();
    const units = new Set();

    data.forEach(row => {
      if (row[23]) years.add(String(row[23]).trim()); // Kolom X (Tahun)
      if (row[1]) units.add(String(row[1]).trim());   // Kolom B (Unit Kerja)
    });

    return {
      tahun: Array.from(years).sort().reverse(),
      unit: Array.from(units).sort()
    };
  } catch (e) {
    return handleError('getSiabaCutiUnduhFilterOptions', e);
  }
}

/**
 * Mengambil Data Tabel Unduh Formulir Cuti
 * Data mulai baris 3.
 */
function getSiabaCutiUnduhData(filters) {
  try {
    const { tahun, unitKerja } = filters;
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    const sheet = ss.getSheetByName("Form Ajuan");
    
    if (!sheet || sheet.getLastRow() < 3) return [];

    const range = sheet.getDataRange();
    const rawValues = range.getValues();        
    const displayValues = range.getDisplayValues(); 
    
    const rows = [];
    
    // Loop mulai dari i = 2 (Baris ke-3)
    for (let i = 2; i < rawValues.length; i++) {
        rows.push({
            raw: rawValues[i],
            display: displayValues[i],
            index: i + 1 // Index baris asli
        });
    }

    // Filter
    const filtered = rows.filter(item => {
        // Filter Tahun (X=23) & Unit (B=1)
        const rowTahun = String(item.display[23] || "").trim(); 
        const rowUnit = String(item.display[1] || "").trim(); 
        
        const matchTahun = (rowTahun === String(tahun));
        const matchUnit = (unitKerja === "Semua" || rowUnit === unitKerja);
        
        return matchTahun && matchUnit;
    });

    // Sorting: Tanggal Pengajuan Terbaru (Kolom A / Index 0)
    filtered.sort((a, b) => {
        const dateA = new Date(a.raw[0]);
        const dateB = new Date(b.raw[0]);
        return dateB - dateA; 
    });

    // Mapping Output
    const output = filtered.map(item => {
         const disp = item.display;
         
         // Status (Kolom R / Index 17)
         let statusRaw = String(disp[17] || "").trim().toUpperCase();
         let statusText = "Diproses";
         if (statusRaw === "V") statusText = "Disetujui";
         else if (statusRaw === "X") statusText = "Dibatalkan";
         
         // PERBAIKAN: Format File URL dari Kolom AT (Index 45)
         // A=0... Z=25, AA=26... AT=45
         let fileUrl = (disp.length > 45) ? disp[45] : "";

         return {
             _rowIndex: item.index,
             nama: disp[2],
             nip: disp[3],
             jenis: disp[16],
             tglMulai: disp[9],
             tglSelesai: disp[10],
             jmlHari: disp[11],
             usulan: disp[15],
             status: statusText,
             fileUrl: fileUrl,
             keterangan: disp[18], // Kolom S
             tglAjuan: disp[0] 
         };
    });

    return output;

  } catch (e) {
    return handleError('getSiabaCutiUnduhData', e);
  }
}

/**
 * Hapus Data Cuti
 */
function deleteSiabaCutiData(rowIndex, deleteCode) {
   try {
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");

    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    const sheet = ss.getSheetByName("Form Ajuan");
    
    // Hapus File jika ada (Kolom AH / Index 33)
    // Perhatikan: getRange pake 1-based index. Col AH = 34
    const fileUrl = sheet.getRange(rowIndex, 34).getValue(); 
    if (fileUrl && String(fileUrl).includes("drive.google.com")) {
        try {
            const fileId = String(fileUrl).match(/[-\w]{25,}/);
            if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch(e) {}
    }
    
    sheet.deleteRow(rowIndex);
    return "Data berhasil dihapus.";
  } catch (e) {
    return handleError('deleteSiabaCutiData', e);
  }
}

function getSiabaCutiData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    
    // 1. Cari ID Spreadsheet
    const targetSheetId = _findSiabaSpreadsheetId(tahun, bulan);
    if (!targetSheetId) return { headers: [], rows: [] };

    // 2. Buka Data
    const ss = SpreadsheetApp.openById(targetSheetId);
    const sheet = ss.getSheetByName("WebData");
    if (!sheet) return { headers: [], rows: [] };

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    // 3. Mapping Kolom
    const colMap = {};
    headers.forEach((h, i) => colMap[h] = i);
    
    // Alias Nama
    if (colMap["NAMA ASN"] === undefined && colMap["Nama"] !== undefined) colMap["NAMA ASN"] = colMap["Nama"];
    
    // Indeks Penting
    const idxUnit = 66; // Asumsi kolom Unit Kerja sama dgn presensi (index 66)
    const idxCT = colMap['CT'];
    const idxCS = colMap['CS'];
    const idxCAP = colMap['CAP'];
    const idxCM = colMap['CM'];

    // 4. Filter Data (Hanya yang punya nilai Cuti > 0)
    const filteredRows = dataRows.filter(row => {
        const matchUnit = (unitKerja === "Semua") || (String(row[idxUnit]) === String(unitKerja));
        if (!matchUnit) return false;
        
        // Cek apakah ada cuti apapun
        const valCT = parseInt(row[idxCT] || 0, 10);
        const valCS = parseInt(row[idxCS] || 0, 10);
        const valCAP = parseInt(row[idxCAP] || 0, 10);
        const valCM = parseInt(row[idxCM] || 0, 10);
        
        return (valCT + valCS + valCAP + valCM) > 0;
    });

    // 5. Susun Output
    const outputHeaders = ["Nama ASN", "NIP", "Cuti Tahunan", "Cuti Sakit", "Cuti Alasan Penting", "Cuti Melahirkan"];
    const outputRows = filteredRows.map(row => [
        row[colMap["NAMA ASN"]] || '-',
        row[colMap["NIP"]] || '-',
        row[idxCT] || '-',
        row[idxCS] || '-',
        row[idxCAP] || '-',
        row[idxCM] || '-'
    ]);
    
    // Sort by Nama
    outputRows.sort((a, b) => a[0].localeCompare(b[0]));

    return { headers: outputHeaders, rows: outputRows };

  } catch (e) {
    return handleError('getSiabaCutiData', e);
  }
}

/**
 * Mengambil Opsi Filter Unit Kerja untuk Halaman Sisa Cuti
 * Sumber: Sheet "Sisa CT" Kolom A
 */
function getSiabaCutiSisaFilterOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    const sheet = ss.getSheetByName("Sisa CT");
    if (!sheet || sheet.getLastRow() < 2) return [];

    // Ambil data Kolom A (Index 1, 1 kolom)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    
    const units = new Set();
    data.forEach(row => {
      if (row[0]) units.add(String(row[0]).trim());
    });

    return Array.from(units).sort();

  } catch (e) {
    return handleError('getSiabaCutiSisaFilterOptions', e);
  }
}

/**
 * Mengambil Data Tabel Sisa Cuti
 * Sumber: Sheet "Sisa CT"
 * Filter: Unit Kerja (Kolom A)
 * Output: Kolom B s/d L (11 Kolom Data)
 */
function getSiabaCutiSisaData(unitKerja) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    const sheet = ss.getSheetByName("Sisa CT");
    
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };

    // Ambil Header (Baris 1)
    // Kita ambil B sampai L (Kolom 2 sampai 12) = 11 Kolom
    const headersRaw = sheet.getRange(1, 2, 1, 11).getDisplayValues()[0];
    
    // Ambil Data (Baris 2 s/d Akhir)
    // Kita ambil A sampai L (Kolom 1 sampai 12)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getDisplayValues();

    const rows = [];
    
    data.forEach(row => {
        const rowUnit = String(row[0] || "").trim(); // Kolom A (Filter)
        
        if (unitKerja === "Semua" || rowUnit === unitKerja) {
            // Ambil data dari index 1 (Kolom B) sampai index 11 (Kolom L)
            // slice(1, 12) mengambil elemen ke-1, 2, ... 11.
            const rowData = row.slice(1, 12); 
            
            // Validasi: Pastikan rowData tidak kosong
            if (rowData.length > 0) {
                rows.push(rowData);
            }
        }
    });

    // Sort berdasarkan Nama (Kolom pertama di rowData = Kolom B asli)
    rows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));

    return { headers: headersRaw, rows: rows };

  } catch (e) {
    return handleError('getSiabaCutiSisaData', e);
  }
}

/**
 * Mengambil Daftar Tahun dari sheet Mapping "Data Cuti"
 * Kolom A: Tahun, Kolom B: Nama Sheet Target
 */
function getSiabaDataCutiYears() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    const sheet = ss.getSheetByName("Data Cuti");
    
    // Validasi sheet mapping
    if (!sheet || sheet.getLastRow() < 2) return [];

    // Ambil data kolom A (Tahun) mulai baris 2
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    
    // Ambil tahun yang unik dan tidak kosong
    const years = data.map(r => String(r[0]).trim()).filter(y => y !== "");
    
    // Return unique years, sorted descending (Terbaru di atas)
    return [...new Set(years)].sort().reverse();
  } catch (e) {
    return handleError('getSiabaDataCutiYears', e);
  }
}

/**
 * Mengambil Data Tabel beserta Header Dinamis
 */
function getSiabaDataCutiTable(year, unitFilter) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SIABA_CUTI_DB);
    
    // 1. Cek Mapping
    const mapSheet = ss.getSheetByName("Data Cuti");
    if (!mapSheet) throw new Error("Sheet mapping 'Data Cuti' tidak ditemukan.");
    
    const mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getDisplayValues();
    let targetSheetName = "";
    
    for(let i=0; i<mapData.length; i++) {
        if (String(mapData[i][0]).trim() === String(year).trim()) {
            targetSheetName = mapData[i][1];
            break;
        }
    }
    
    if (!targetSheetName) return { error: "Sheet untuk tahun " + year + " tidak ditemukan." };

    // 2. Buka Sheet Target
    const targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) return { error: "Sheet '" + targetSheetName + "' tidak ditemukan." };

    if (targetSheet.getLastRow() < 3) return { units: [], headers: [], rows: [] };

    // 3. Ambil Semua Data (Header + Isi)
    // Ambil Kolom A (Unit) sampai M (Index 12)
    const range = targetSheet.getRange(1, 1, targetSheet.getLastRow(), 13);
    const values = range.getDisplayValues();
    
    // Pisahkan Header (Baris 1 & 2) dan Data (Baris 3+)
    // Kita hanya butuh Kolom B-M untuk ditampilkan (slice(1, 13))
    const rawHeader1 = values[0].slice(1, 13);
    const rawHeader2 = values[1].slice(1, 13);
    
    // Data mulai index 2
    const dataRows = [];
    const unitsSet = new Set();

    for (let i = 2; i < values.length; i++) {
        const row = values[i];
        const rowUnit = String(row[0] || "").trim(); // Kolom A
        
        if (rowUnit) unitsSet.add(rowUnit);

        // Filter Unit
        if (unitFilter === "Semua" || rowUnit === unitFilter) {
            dataRows.push(row.slice(1, 13)); // Ambil B-M
        }
    }

    return {
        units: Array.from(unitsSet).sort(),
        headers: [rawHeader1, rawHeader2], // Kirim 2 baris header
        rows: dataRows
    };

  } catch (e) {
    return handleError('getSiabaDataCutiTable', e);
  }
}