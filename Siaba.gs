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