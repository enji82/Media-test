function getSkArsipFolderIds() {
  try {
    return {
      'MAIN_SK': FOLDER_CONFIG.MAIN_SK
    };
  } catch (e) {
    return handleError('getSkArsipFolderIds', e);
  }
}

function getMasterSkOptions() {
  // Kunci cache unik
  const cacheKey = 'master_sk_options_v1';
  
  // Gunakan fungsi cache yang sudah ada
  return getCachedData(cacheKey, function() {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
      const getValuesFromSheet = (sheetName) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return [];
        return sheet.getRange('A2:A' + sheet.getLastRow()).getValues()
                    .flat()
                    .filter(value => String(value).trim() !== '');
      };

      return {
        'Nama SD': getValuesFromSheet('Nama SD').sort(),
        'Tahun Ajaran': getValuesFromSheet('Tahun Ajaran').sort().reverse(),
        'Semester': getValuesFromSheet('Semester').sort(),
        'Kriteria SK': getValuesFromSheet('Kriteria SK').sort()
      };
    } catch (e) {
      // Saat caching, kita lempar error agar tidak menyimpan cache yang rusak
      throw new Error(`Gagal mengambil SK Options: ${e.message}`);
    }
  });
}

function processManualForm(formData) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);  
    
    const tahunAjaranFolderName = formData.tahunAjaran.replace(/\//g, '-');
    const tahunAjaranFolder = getOrCreateFolder(mainFolder, tahunAjaranFolderName);

    const semesterFolderName = formData.semester;
    const targetFolder = getOrCreateFolder(tahunAjaranFolder, semesterFolderName);

    const newFilename = `${formData.namaSD} - ${tahunAjaranFolderName} - ${formData.semester} - ${formData.kriteriaSK}.pdf`;
    
    const decodedData = Utilities.base64Decode(formData.fileData.data);
    const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFilename);
    const newFile = targetFolder.createFile(blob);
    const fileUrl = newFile.getUrl();
    const newRow = [ new Date(), formData.namaSD, formData.tahunAjaran, formData.semester, formData.nomorSK, new Date(formData.tanggalSK), formData.kriteriaSK, fileUrl ];
    
    sheet.appendRow(newRow);
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 6).setNumberFormat("dd-MM-yyyy");

    return "Dokumen SK berhasil diunggah.";
  } catch (e) {
    return handleError('processManualForm', e);
  }
}

function getSKRiwayatData() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.id)
                                .getSheetByName(SPREADSHEET_CONFIG.SK_FORM_RESPONSES.sheet);
    const desiredHeaders = [
        "Nama SD",        // <-- Dipakai oleh filterNamaSekolah
        "Tahun Ajaran",   // <-- Dipakai oleh filterTahun
        "Semester",       // <-- Dipakai oleh filterSemester
        "Nomor SK", 
        "Tanggal SK", 
        "Kriteria SK", 
        "Dokumen", 
        "Tanggal Unggah"
    ];
    
    if (!sheet || sheet.getLastRow() < 2) {
      // Jika data kosong, pastikan headers tetap dikirim agar client-side bisa me-render filter.
      return { headers: desiredHeaders, rows: [] }; 
    }
    
    // Tetap gunakan getValues() untuk mendapatkan objek tanggal asli
    const allData = sheet.getDataRange().getValues(); 
    const originalHeaders = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    // Cari indeks kolom 'Tanggal Unggah' untuk sorting
    const timestampIndex = originalHeaders.indexOf('Tanggal Unggah');

    // Langsung urutkan berdasarkan objek tanggal mentah. Ini lebih cepat dan akurat.
    dataRows.sort((a, b) => {
        const dateA = a[timestampIndex] instanceof Date ? a[timestampIndex].getTime() : 0;
        const dateB = b[timestampIndex] instanceof Date ? b[timestampIndex].getTime() : 0;
        return dateB - dateA; // Mengurutkan dari terbaru (nilai terbesar) ke terlama
    });

    // Setelah diurutkan, baru kita format datanya untuk ditampilkan
    let structuredRows = dataRows.map(row => {
      const rowObject = {};
      originalHeaders.forEach((header, index) => {
        let cell = row[index];
        if (cell instanceof Date) {
          // Terapkan format yang berbeda berdasarkan nama header
          if (header === 'Tanggal SK') {
            rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
          } else { // Asumsikan sisanya adalah timestamp
            rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
          }
        } else {
          rowObject[header] = cell;
        }
      });
      return rowObject;
    });

    return {
      headers: desiredHeaders,
      rows: structuredRows
    };
  } catch (e) {
    return handleError('getSKRiwayatData', e);
  }
}

function getSKStatusData() {
  try {
    const config = SPREADSHEET_CONFIG.SK_BAGI_TUGAS;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);

    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }

    const dataRange = sheet.getDataRange();
    // Ambil nilai teks biasa
    const displayValues = dataRange.getDisplayValues();
    // Ambil nilai rich text untuk mendapatkan URL
    const richTextValues = dataRange.getRichTextValues();

    const headers = displayValues[0].map(h => String(h).trim());
    const dataRows = displayValues.slice(1);
    const richTextRows = richTextValues.slice(1);

    // Cari tahu di kolom mana 'Dokumen' berada
    const docIndex = headers.indexOf('Dokumen');

    const structuredRows = dataRows.map((row, rowIndex) => {
      const rowObject = {};
      headers.forEach((header, colIndex) => {
        // Jika ini adalah kolom "Dokumen" dan ada linknya, ambil URL-nya
        if (colIndex === docIndex) {
          const richText = richTextRows[rowIndex][colIndex];
          const linkUrl = richText.getLinkUrl();
          // Jika ada link, gunakan URL-nya. Jika tidak, gunakan teks biasa.
          rowObject[header] = linkUrl || row[colIndex]; 
        } else {
          rowObject[header] = row[colIndex];
        }
      });
      return rowObject;
    });

    return {
      headers: headers,
      rows: structuredRows
    };
  } catch (e) {
    return handleError('getSKStatusData', e);
  }
}

function getSKKelolaData() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
      return { headers: [], rows: [] };
    }

    const originalData = sheet.getDataRange().getValues();
    const originalHeaders = originalData[0].map(h => String(h).trim());
    const dataRows = originalData.slice(1);
    
    const parseDate = (value) => value instanceof Date && !isNaN(value) ? value : new Date(0);

    const indexedData = dataRows.map((row, index) => ({
      row: row,
      originalIndex: index + 2
    }));
    
    const updateIndex = originalHeaders.indexOf('Update');
    const timestampIndex = originalHeaders.indexOf('Tanggal Unggah');

    indexedData.sort((a, b) => {
      const dateB_update = parseDate(b.row[updateIndex]);
      const dateA_update = parseDate(a.row[updateIndex]);
      if (dateB_update.getTime() !== dateA_update.getTime()) {
        return dateB_update - dateA_update;
      }
      const dateB_timestamp = parseDate(b.row[timestampIndex]);
      const dateA_timestamp = parseDate(a.row[timestampIndex]);
      return dateB_timestamp - dateA_timestamp;
    });

    const structuredRows = indexedData.map(item => {
      const rowObject = {
        _rowIndex: item.originalIndex,
        _source: 'SK'
      };
      originalHeaders.forEach((header, i) => {
      let cell = item.row[i];
      // MODIFIKASI DIMULAI DI SINI
      if (header === 'Tanggal SK' && cell instanceof Date) {
      rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else if ((header === 'Tanggal Unggah' || header === 'Update') && cell instanceof Date) {
      rowObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      // MODIFIKASI SELESAI
      } else {
      rowObject[header] = cell;
    }
  });
  return rowObject;
    });
    
    const desiredHeaders = ["Nama SD", "Tahun Ajaran", "Semester", "Nomor SK", "Kriteria SK", "Dokumen", "Aksi", "Tanggal Unggah", "Update"];

    return {
      headers: desiredHeaders,
      rows: structuredRows
    };
  } catch (e) {
    return handleError("getSKKelolaData", e);
  }
}

function getSKDataByRow(rowIndex) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    // Ambil nilai mentah (RAW) untuk mendapatkan objek Date asli
    const rawValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Ambil nilai tampilan (DISPLAY) untuk konsistensi string/angka
    const displayValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const rowData = {};
    headers.forEach((header, i) => {
      // KUNCI PERBAIKAN: Format Tanggal SK ke YYYY-MM-DD
      if (header === 'Tanggal SK' && rawValues[i] instanceof Date) {
        // Format yang wajib untuk HTML input type="date"
        rowData[header] = Utilities.formatDate(rawValues[i], "UTC", "yyyy-MM-dd");
      } else {
        // Gunakan display value untuk field lain (Nomor SK, dll.)
        rowData[header] = displayValues[i];
      }
    });
    return rowData;
  } catch (e) {
    return handleError("getSKDataByRow", e);
  }
}

function updateSKData(formData) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const range = sheet.getRange(formData.rowIndex, 1, 1, headers.length);
    const existingRowValues = range.getDisplayValues()[0];
    const existingRowObject = {};
    headers.forEach((header, i) => { existingRowObject[header] = existingRowValues[i]; });

    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
    const tahunAjaranFolderName = existingRowObject['Tahun Ajaran'].replace(/\//g, '-');
    const tahunAjaranFolder = getOrCreateFolder(mainFolder, tahunAjaranFolderName);
    
    let fileUrl = existingRowObject['Dokumen'];
    const fileUrlIndex = headers.indexOf('Dokumen');

    const newSemesterFolderName = formData['Semester'];
    const newTargetFolder = getOrCreateFolder(tahunAjaranFolder, newSemesterFolderName);
    const newFilename = `${existingRowObject['Nama SD']} - ${tahunAjaranFolderName} - ${newSemesterFolderName} - ${formData['Kriteria SK']}.pdf`;

    if (formData.fileData && formData.fileData.data) {
      if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        try {
          const fileId = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
          if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch (e) {
          Logger.log(`Gagal menghapus file lama saat upload baru: ${e.message}`);
        }
      }
      
      const decodedData = Utilities.base64Decode(formData.fileData.data);
      const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFilename);
      const newFile = newTargetFolder.createFile(blob);
      fileUrl = newFile.getUrl();

    } else if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        const fileIdMatch = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
        if (fileIdMatch) {
            const fileId = fileIdMatch[0];
            const file = DriveApp.getFileById(fileId);
            const currentFileName = file.getName();
            const currentParentFolder = file.getParents().next();

            if (currentFileName !== newFilename || currentParentFolder.getName() !== newSemesterFolderName) {
                file.moveTo(newTargetFolder);
                file.setName(newFilename);
                fileUrl = file.getUrl();
            }
        }
    }
    
    formData['Dokumen'] = fileUrl;
    formData['Update'] = new Date();

    const newRowValuesForSheet = headers.map(header => {
      return formData.hasOwnProperty(header) ? formData[header] : existingRowObject[header];
    });

    sheet.getRange(formData.rowIndex, 1, 1, headers.length).setValues([newRowValuesForSheet]);
    
    const tanggalSKIndex = headers.indexOf('Tanggal SK');
    if (tanggalSKIndex !== -1) {
      sheet.getRange(formData.rowIndex, tanggalSKIndex + 1).setNumberFormat("dd-MM-yyyy");
    }
    
    return "Data berhasil diperbarui!";
  } catch (e) {
    return handleError('updateSKData', e);
  }
}

function deleteSKData(rowIndex, deleteCode) {
  try {
    const todayCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    if (String(deleteCode).trim() !== todayCode) throw new Error("Kode Hapus salah.");

    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const fileUrlIndex = headers.findIndex(h => h.trim().toLowerCase() === 'dokumen');
    
    if (fileUrlIndex !== -1) {
        const fileUrl = sheet.getRange(rowIndex, fileUrlIndex + 1).getValue();
        if (fileUrl && typeof fileUrl === 'string') {
            const fileId = fileUrl.match(/[-\w]{25,}/);
            if (fileId) {
                try {
                    DriveApp.getFileById(fileId[0]).setTrashed(true);
                } catch (err) {
                    Logger.log(`Gagal menghapus file dengan ID ${fileId[0]}: ${err.message}`);
                }
            }
        }
    }
    
    sheet.deleteRow(rowIndex);
    return "Data dan file terkait berhasil dihapus.";
  } catch (e) {
    return handleError("deleteSKData", e);
  }
}