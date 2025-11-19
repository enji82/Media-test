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
    const dateHeaders = Array.from({length: 31}, (_, i) => String(i + 1)); // "1", "2", ... "31"
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
    // Alias mapping jika nama di sheet berbeda
    if (!headerMap.hasOwnProperty("Nama ASN") && headerMap.hasOwnProperty("Nama")) headerMap["Nama ASN"] = headerMap["Nama"];

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