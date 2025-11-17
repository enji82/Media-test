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

    // 3. Ambil data mentah (Logika ini sudah benar)
    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0].map(h => String(h).trim()); // <-- Tambahkan .trim() untuk kebersihan
    const dataRows = allData.slice(1);

    // 4. Terapkan filter 'unitKerja' (Logika ini sudah benar)
    // Asumsi 'Unit Kerja' ada di data mentah di indeks 2
    const unitKerjaIndex = 2; 
    const filteredRows = dataRows.filter(row => {
        const matchUnitKerja = (unitKerja === "Semua") || (String(row[unitKerjaIndex]) === String(unitKerja));
        return matchUnitKerja;
    });

    // --- PERBAIKAN UTAMA DIMULAI DI SINI ---

    // 5. Tentukan Header yang Anda Inginkan (SESUAI URUTAN TABEL LAMA ANDA)
    // (Saya ambil dari page_siaba_daftar_presensi.html)
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

    // 6. Buat "Peta Indeks"
    // Ini memetakan nama header yang Anda inginkan ke posisi kolomnya di "WebData"
    const headerMap = {};
    headers.forEach((headerName, index) => {
        headerMap[headerName] = index;
    });

    // 7. Susun Ulang Data (Transformasi)
    const displayRows = filteredRows.map(row => {
        // Buat baris baru HANYA berisi data yang kita inginkan, SESUAI URUTAN
        const newRow = [];
        for (const desiredHeader of desiredHeaders) {
            const indexDiWebData = headerMap[desiredHeader];
            
            if (indexDiWebData !== undefined) {
                // Ambil data dari "WebData" menggunakan indeks yang benar
                newRow.push(row[indexDiWebData]);
            } else {
                // Jika kolom tidak ditemukan di "WebData", isi dengan strip
                newRow.push('-'); 
            }
        }
        return newRow;
    });
    
    // --- AKHIR PERBAIKAN UTAMA ---

    // 8. Logika Sorting (Sekarang menggunakan desiredHeaders)
    const tpIndex = desiredHeaders.indexOf('TP');
    const taIndex = desiredHeaders.indexOf('TA');
    const plaIndex = desiredHeaders.indexOf('PLA');
    const tapIndex = desiredHeaders.indexOf('TAp');
    const tuIndex = desiredHeaders.indexOf('TU');
    const namaIndex = desiredHeaders.indexOf('Nama'); // <-- Ini sekarang di indeks 0

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

    // 9. Kembalikan data (Sekarang menggunakan desiredHeaders)
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

// -----------------------------------------------------------------
// FUNGSI UNTUK "ASN TIDAK PRESENSI" (Biarkan seperti aslinya)
// -----------------------------------------------------------------

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