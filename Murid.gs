function getMuridPaudKelasUsiaData() {
  const cacheKey = 'murid_paud_kelas_usia_v1';
  return getCachedData(cacheKey, () => {
    try {
      return getDataFromSheet('MURID_PAUD_KELAS');
    } catch (e) {
      throw new Error(`Gagal memuat data Murid PAUD Kelas Usia: ${e.message}`);
    }
  });
}

function getMuridPaudJenisKelaminData() {
  const cacheKey = 'murid_paud_jk_v1';
  return getCachedData(cacheKey, () => {
    try {
      return getDataFromSheet('MURID_PAUD_JK');
    } catch (e) {
      throw new Error(`Gagal memuat data Murid PAUD JK: ${e.message}`);
    }
  });
}

function getMuridPaudJumlahBulananData() {
  // Data bulanan mungkin lebih sering berubah, tapi kita tetap cache selama 1 jam (3600 detik)
  const cacheKey = 'murid_paud_bulanan_v1';
  return getCachedData(cacheKey, () => {
    try {
      return getDataFromSheet('MURID_PAUD_BULANAN');
    } catch (e) {
      throw new Error(`Gagal memuat data Murid PAUD Bulanan: ${e.message}`);
    }
  }, 3600); 
}

function getMuridSdKelasData() {
  const cacheKey = 'murid_sd_kelas_v1';
  return getCachedData(cacheKey, () => {
    try {
      return getDataFromSheet('MURID_SD_KELAS');
    } catch (e) {
      throw new Error(`Gagal memuat data Murid SD Kelas: ${e.message}`);
    }
  });
}

function getMuridSdRombelData() {
  const cacheKey = 'murid_sd_rombel_v1';
  return getCachedData(cacheKey, () => {
    try {
      return getDataFromSheet('MURID_SD_ROMBEL');
    } catch (e) {
      throw new Error(`Gagal memuat data Murid SD Rombel: ${e.message}`);
    }
  });
}

function getMuridSdJenisKelaminData() {
  const cacheKey = 'murid_sd_jk_v1';
  return getCachedData(cacheKey, () => {
    try {
      return getDataFromSheet('MURID_SD_JK');
    } catch (e) {
      throw new Error(`Gagal memuat data Murid SD JK: ${e.message}`);
    }
  });
}

function getMuridSdAgamaData() {
  const cacheKey = 'murid_sd_agama_v1';
  return getCachedData(cacheKey, () => {
    try {
      return getDataFromSheet('MURID_SD_AGAMA');
    } catch (e) {
      throw new Error(`Gagal memuat data Murid SD Agama: ${e.message}`);
    }
  });
}

function getMuridSdJumlahBulananData() {
  const cacheKey = 'murid_sd_bulanan_v1';
  return getCachedData(cacheKey, () => {
    try {
      const config = SPREADSHEET_CONFIG.MURID_SD_BULANAN;
      const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
      if (!sheet) throw new Error("Sheet tidak ditemukan");
      return sheet.getDataRange().getValues(); // Menggunakan getValues() untuk angka
    } catch (e) {
      throw new Error(`Gagal memuat data Murid SD Bulanan: ${e.message}`);
    }
  }, 3600); // Cache selama 1 jam
}