/**
 * ============================================================================
 * SiL TAHFIDZ - Sistem Laporan Perkembangan Tahfidz Siswa
 * MTsS Nurul Islam Indonesia
 * Backend Engine: Google Apps Script + Google Sheets + Google Drive
 * ============================================================================
 */

const APP_PROP_DB_ID = 'SIL_TAHFIDZ_SPREADSHEET_ID';
const APP_PROP_FOLDER_ID = 'SIL_TAHFIDZ_FOLDER_ID';

const SHEETS_SCHEMA = {
  GURU: {
    name: 'DATA GURU PENGUJI',
    headers: ['ID_GURU', 'USERNAME', 'PASSWORD', 'NAMA_LENGKAP', 'JABATAN', 'STATUS']
  },
  MADRASAH: {
    name: 'DATA MADRASAH',
    headers: ['ID_MADRASAH', 'KELAS', 'SEMESTER', 'NAMA_MADRASAH', 'NAMA_GURU_PENGUJI', 'NAMA_KEPALA_MADRASAH', 'TANGGAL_RAPOR', 'TAHUN_AJARAN', 'UPDATED_AT']
  },
  SISWA: {
    name: 'DATA SISWA',
    headers: ['ID_SISWA', 'KELAS', 'SEMESTER', 'NO_URUT', 'NAMA_SISWA', 'NIS_NISN', 'TTL', 'ALAMAT', 'NAMA_ORANGTUA', 'CREATED_AT']
  },
  SURAH: {
    name: 'DAFTAR SURAH',
    headers: ['ID_SURAH', 'KELAS', 'SEMESTER', 'NO_URUT', 'NAMA_SURAH', 'KETERANGAN', 'CREATED_AT']
  },
  PROGRES: {
    name: 'PROGRES HAFALAN',
    headers: ['ID_PROGRES', 'ID_SISWA', 'KELAS', 'SEMESTER', 'TARGET_SURAH', 'JUMLAH_TERCAPAI', 'TOTAL_SURAH', 'PERSENTASE', 'DETAIL_SURAH_JSON', 'CATATAN_GURU', 'UPDATED_AT']
  },
  RAPOR: {
    name: 'DATA RAPOR',
    headers: ['ID_RAPOR', 'ID_SISWA', 'NAMA_SISWA', 'KELAS', 'SEMESTER', 'TAHUN_AJARAN', 'PERSENTASE', 'FILE_DRIVE_ID', 'FILE_URL', 'GENERATED_AT']
  },
  KONFIGURASI: {
    name: 'KONFIGURASI',
    headers: ['KEY', 'VALUE', 'UPDATED_AT']
  }
};

/**
 * Entry Point Web App
 */
function doGet(e) {
  initDatabase();
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('SiL Tahfidz - MTsS Nurul Islam Indonesia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Inisialisasi Otomatis Folder Google Drive & Google Spreadsheet
 * Aman dijalankan berulang kali tanpa menimpa data yang sudah ada.
 */
function initDatabase() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty(APP_PROP_FOLDER_ID);
  let ssId = props.getProperty(APP_PROP_DB_ID);
  let folder, ss;

  // 1. Cek atau Buat Folder Utama di Google Drive (Private/Restricted)
  try {
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
    }
  } catch (e) {
    folder = null;
  }

  if (!folder) {
    const folders = DriveApp.getFoldersByName('SiL_Tahfidz_Database_MTsS_NII');
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('SiL_Tahfidz_Database_MTsS_NII');
      folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      folder.createFolder('Arsip_Rapor_PDF');
      folder.createFolder('Backup_Excel');
    }
    folderId = folder.getId();
    props.setProperty(APP_PROP_FOLDER_ID, folderId);
  }

  // 2. Cek atau Buat Spreadsheet Database Utama
  try {
    if (ssId) {
      ss = SpreadsheetApp.openById(ssId);
    }
  } catch (e) {
    ss = null;
  }

  if (!ss) {
    const files = folder.getFilesByName('DB_SiL_Tahfidz_MTsS_NII');
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create('DB_SiL_Tahfidz_MTsS_NII');
      const file = DriveApp.getFileById(ss.getId());
      file.moveTo(folder);
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    }
    ssId = ss.getId();
    props.setProperty(APP_PROP_DB_ID, ssId);
  }

  // 3. Pastikan semua Sheet & Header tersedia
  Object.keys(SHEETS_SCHEMA).forEach(key => {
    const schema = SHEETS_SCHEMA[key];
    let sheet = ss.getSheetByName(schema.name);
    if (!sheet) {
      sheet = ss.insertSheet(schema.name);
      sheet.appendRow(schema.headers);
      sheet.getRange(1, 1, 1, schema.headers.length)
        .setFontWeight('bold')
        .setBackground('#2D4739')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(schema.headers);
    }
  });

  // Hapus Sheet1 bawaan jika masih ada dan sheet lain sudah terbentuk
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // 4. Seed Data Awal (Guru Penguji Default & Konfigurasi) jika kosong
  const guruSheet = ss.getSheetByName(SHEETS_SCHEMA.GURU.name);
  if (guruSheet.getLastRow() <= 1) {
    guruSheet.appendRow(['GRU-001', 'guru', 'tahfidz123', 'Ustadz Penguji Tahfidz', 'Guru Penguji', 'AKTIF']);
  }

  const configSheet = ss.getSheetByName(SHEETS_SCHEMA.KONFIGURASI.name);
  if (configSheet.getLastRow() <= 1) {
    const now = new Date().toISOString();
    configSheet.appendRow(['APP_NAME', 'SiL Tahfidz', now]);
    configSheet.appendRow(['APP_SUBTITLE', 'Sistem Laporan Perkembangan Tahfidz Siswa MTsS Nurul Islam Indonesia', now]);
    configSheet.appendRow(['LANGUAGE', 'ID', now]);
    configSheet.appendRow(['SCHOOL_LOGO_BASE64', '', now]);
    configSheet.appendRow(['SPREADSHEET_ID', ssId, now]);
    configSheet.appendRow(['FOLDER_ID', folderId, now]);
  }

  return { ssId: ssId, folderId: folderId };
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty(APP_PROP_DB_ID);
  if (!ssId) {
    ssId = initDatabase().ssId;
  }
  return SpreadsheetApp.openById(ssId);
}

function generateId_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

/**
 * ============================================================================
 * KONFIGURASI & PENGATURAN APLIKASI
 * ============================================================================
 */
function getAppConfig() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEETS_SCHEMA.KONFIGURASI.name);
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      config[data[i][0]] = data[i][1];
    }
  }
  return config;
}

function saveAppConfig(updates) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.KONFIGURASI.name);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();

    Object.keys(updates).forEach(key => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(updates[key]);
          sheet.getRange(i + 1, 3).setValue(now);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([key, updates[key], now]);
      }
    });

    return { success: true, config: getAppConfig() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Verifikasi Login Opsional / Sesi Guru Penguji
 */
function verifyGuruLogin(username, password) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEETS_SCHEMA.GURU.name);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(username).trim() &&
        String(data[i][2]).trim() === String(password).trim() &&
        data[i][5] === 'AKTIF') {
      return {
        success: true,
        guru: {
          id: data[i][0],
          username: data[i][1],
          nama: data[i][3]
        }
      };
    }
  }
  return { success: false, message: 'Username atau Password tidak valid.' };
}

/**
 * ============================================================================
 * LOAD SEMUA DATA BERDASARKAN KELAS & SEMESTER
 * ============================================================================
 */
function getDashboardBundle(kelas, semester) {
  const ss = getSpreadsheet_();
  const kls = String(kelas).trim();
  const sem = String(semester).trim();

  // 1. Data Madrasah
  const madrasahSheet = ss.getSheetByName(SHEETS_SCHEMA.MADRASAH.name);
  const madrasahRows = madrasahSheet.getDataRange().getValues();
  let madrasah = {
    id: '',
    kelas: kls,
    semester: sem,
    namaMadrasah: 'MTsS Nurul Islam Indonesia',
    namaGuruPenguji: '',
    namaKepalaMadrasah: '',
    tanggalRapor: '',
    tahunAjaran: '2026/2027'
  };
  for (let i = 1; i < madrasahRows.length; i++) {
    if (String(madrasahRows[i][1]) === kls && String(madrasahRows[i][2]) === sem) {
      madrasah = {
        id: madrasahRows[i][0],
        kelas: madrasahRows[i][1],
        semester: madrasahRows[i][2],
        namaMadrasah: madrasahRows[i][3] || 'MTsS Nurul Islam Indonesia',
        namaGuruPenguji: madrasahRows[i][4] || '',
        namaKepalaMadrasah: madrasahRows[i][5] || '',
        tanggalRapor: madrasahRows[i][6] || '',
        tahunAjaran: madrasahRows[i][7] || '2026/2027'
      };
      break;
    }
  }

  // 2. Data Siswa
  const siswaSheet = ss.getSheetByName(SHEETS_SCHEMA.SISWA.name);
  const siswaRows = siswaSheet.getDataRange().getValues();
  const siswaList = [];
  for (let i = 1; i < siswaRows.length; i++) {
    if (String(siswaRows[i][1]) === kls && String(siswaRows[i][2]) === sem && siswaRows[i][4]) {
      siswaList.push({
        id: siswaRows[i][0],
        kelas: siswaRows[i][1],
        semester: siswaRows[i][2],
        no: siswaRows[i][3] || (siswaList.length + 1),
        nama: String(siswaRows[i][4]),
        nis: String(siswaRows[i][5] || ''),
        ttl: String(siswaRows[i][6] || ''),
        alamat: String(siswaRows[i][7] || ''),
        orangTua: String(siswaRows[i][8] || '')
      });
    }
  }

  // 3. Daftar Surah
  const surahSheet = ss.getSheetByName(SHEETS_SCHEMA.SURAH.name);
  const surahRows = surahSheet.getDataRange().getValues();
  const surahList = [];
  for (let i = 1; i < surahRows.length; i++) {
    if (String(surahRows[i][1]) === kls && String(surahRows[i][2]) === sem && surahRows[i][4]) {
      surahList.push({
        id: surahRows[i][0],
        kelas: surahRows[i][1],
        semester: surahRows[i][2],
        no: surahRows[i][3] || (surahList.length + 1),
        namaSurah: String(surahRows[i][4]),
        keterangan: String(surahRows[i][5] || '')
      });
    }
  }

  // 4. Progres Hafalan
  const progresSheet = ss.getSheetByName(SHEETS_SCHEMA.PROGRES.name);
  const progresRows = progresSheet.getDataRange().getValues();
  const progresMap = {};
  for (let i = 1; i < progresRows.length; i++) {
    if (String(progresRows[i][2]) === kls && String(progresRows[i][3]) === sem && progresRows[i][1]) {
      let detail = [];
      try {
        detail = JSON.parse(progresRows[i][8] || '[]');
      } catch (e) {
        detail = [];
      }
      progresMap[progresRows[i][1]] = {
        id: progresRows[i][0],
        idSiswa: progresRows[i][1],
        kelas: progresRows[i][2],
        semester: progresRows[i][3],
        targetSurah: progresRows[i][4] || '',
        jumlahTercapai: Number(progresRows[i][5] || 0),
        totalSurah: Number(progresRows[i][6] || surahList.length),
        persentase: Number(progresRows[i][7] || 0),
        detailSurah: detail,
        catatanGuru: progresRows[i][9] || ''
      };
    }
  }

  return {
    madrasah: madrasah,
    siswaList: siswaList,
    surahList: surahList,
    progresMap: progresMap
  };
}

/**
 * ============================================================================
 * CRUD DATA MADRASAH
 * ============================================================================
 */
function saveDataMadrasah(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.MADRASAH.name);
    const rows = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    const kls = String(payload.kelas).trim();
    const sem = String(payload.semester).trim();

    let rowIndex = -1;
    let id = payload.id || generateId_('MDR');

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === kls && String(rows[i][2]) === sem) {
        rowIndex = i + 1;
        id = rows[i][0] || id;
        break;
      }
    }

    const rowData = [
      id,
      kls,
      sem,
      payload.namaMadrasah || 'MTsS Nurul Islam Indonesia',
      payload.namaGuruPenguji || '',
      payload.namaKepalaMadrasah || '',
      payload.tanggalRapor || '',
      payload.tahunAjaran || '',
      now
    ];

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, message: 'Data Madrasah berhasil disimpan.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * CRUD DATA SISWA (MANUAL & BULK EXCEL)
 * ============================================================================
 */
function saveSiswa(payload) {
  if (!payload.nama || String(payload.nama).trim() === '') {
    throw new Error('Nama siswa wajib diisi.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.SISWA.name);
    const rows = sheet.getDataRange().getValues();
    const kls = String(payload.kelas).trim();
    const sem = String(payload.semester).trim();
    const now = new Date().toISOString();

    if (payload.id) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === payload.id) {
          sheet.getRange(i + 1, 4, 1, 6).setValues([[
            payload.no || rows[i][3],
            String(payload.nama).trim(),
            String(payload.nis || '').trim(),
            String(payload.ttl || '').trim(),
            String(payload.alamat || '').trim(),
            String(payload.orangTua || '').trim()
          ]]);
          return { success: true, message: 'Data siswa berhasil diperbarui.' };
        }
      }
    }

    // Hitung nomor urut kelas & semester
    let count = 0;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === kls && String(rows[i][2]) === sem) count++;
    }

    const newId = generateId_('SIS');
    sheet.appendRow([
      newId,
      kls,
      sem,
      count + 1,
      String(payload.nama).trim(),
      String(payload.nis || '').trim(),
      String(payload.ttl || '').trim(),
      String(payload.alamat || '').trim(),
      String(payload.orangTua || '').trim(),
      now
    ]);

    return { success: true, message: 'Data siswa berhasil ditambahkan.' };
  } finally {
    lock.releaseLock();
  }
}

function importSiswaBulk(kelas, semester, listSiswa) {
  if (!Array.isArray(listSiswa) || listSiswa.length === 0) {
    throw new Error('Tidak ada data siswa valid yang ditemukan pada file Excel.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.SISWA.name);
    const rows = sheet.getDataRange().getValues();
    const kls = String(kelas).trim();
    const sem = String(semester).trim();
    const now = new Date().toISOString();

    let count = 0;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === kls && String(rows[i][2]) === sem) count++;
    }

    const newRows = [];
    listSiswa.forEach((item, idx) => {
      if (item.nama && String(item.nama).trim() !== '') {
        count++;
        newRows.push([
          generateId_('SIS'),
          kls,
          sem,
          count,
          String(item.nama).trim(),
          String(item.nis || '').trim(),
          String(item.ttl || '').trim(),
          String(item.alamat || '').trim(),
          String(item.orangTua || '').trim(),
          now
        ]);
      }
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    return { success: true, count: newRows.length, message: newRows.length + ' data siswa berhasil diimpor ke database.' };
  } finally {
    lock.releaseLock();
  }
}

function deleteSiswa(idSiswa) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.SISWA.name);
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (rows[i][0] === idSiswa) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    // Hapus progres terkait jika ada
    const progSheet = ss.getSheetByName(SHEETS_SCHEMA.PROGRES.name);
    const progRows = progSheet.getDataRange().getValues();
    for (let i = progRows.length - 1; i >= 1; i--) {
      if (progRows[i][1] === idSiswa) {
        progSheet.deleteRow(i + 1);
      }
    }
    return { success: true, message: 'Data siswa berhasil dihapus.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * CRUD DAFTAR SURAH (MANUAL & BULK EXCEL)
 * ============================================================================
 */
function saveSurah(payload) {
  if (!payload.namaSurah || String(payload.namaSurah).trim() === '') {
    throw new Error('Nama Surah wajib diisi.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.SURAH.name);
    const rows = sheet.getDataRange().getValues();
    const kls = String(payload.kelas).trim();
    const sem = String(payload.semester).trim();
    const now = new Date().toISOString();

    if (payload.id) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === payload.id) {
          sheet.getRange(i + 1, 5, 1, 2).setValues([[
            String(payload.namaSurah).trim(),
            String(payload.keterangan || '').trim()
          ]]);
          return { success: true, message: 'Daftar surah berhasil diperbarui.' };
        }
      }
    }

    let count = 0;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === kls && String(rows[i][2]) === sem) count++;
    }

    sheet.appendRow([
      generateId_('SUR'),
      kls,
      sem,
      count + 1,
      String(payload.namaSurah).trim(),
      String(payload.keterangan || '').trim(),
      now
    ]);

    return { success: true, message: 'Surah berhasil ditambahkan.' };
  } finally {
    lock.releaseLock();
  }
}

function importSurahBulk(kelas, semester, listSurah) {
  if (!Array.isArray(listSurah) || listSurah.length === 0) {
    throw new Error('Tidak ada daftar surah valid yang ditemukan pada file Excel.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.SURAH.name);
    const rows = sheet.getDataRange().getValues();
    const kls = String(kelas).trim();
    const sem = String(semester).trim();
    const now = new Date().toISOString();

    let count = 0;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === kls && String(rows[i][2]) === sem) count++;
    }

    const newRows = [];
    listSurah.forEach(item => {
      if (item.namaSurah && String(item.namaSurah).trim() !== '') {
        count++;
        newRows.push([
          generateId_('SUR'),
          kls,
          sem,
          count,
          String(item.namaSurah).trim(),
          String(item.keterangan || '').trim(),
          now
        ]);
      }
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    return { success: true, count: newRows.length, message: newRows.length + ' surah berhasil diimpor.' };
  } finally {
    lock.releaseLock();
  }
}

function deleteSurah(idSurah) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.SURAH.name);
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (rows[i][0] === idSurah) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { success: true, message: 'Surah berhasil dihapus.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * CRUD PROGRES HAFALAN SISWA
 * ============================================================================
 */
function saveProgresHafalan(payload) {
  if (!payload.idSiswa) {
    throw new Error('Siswa belum dipilih.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.PROGRES.name);
    const rows = sheet.getDataRange().getValues();
    const kls = String(payload.kelas).trim();
    const sem = String(payload.semester).trim();
    const now = new Date().toISOString();

    const detail = Array.isArray(payload.detailSurah) ? payload.detailSurah : [];
    const totalSurah = Number(payload.totalSurah || detail.length || 0);
    const jumlahTercapai = detail.filter(d => d.status === 'Tercapai').length;
    const persentase = totalSurah > 0 ? Math.round((jumlahTercapai / totalSurah) * 1000) / 10 : 0;

    let rowIndex = -1;
    let id = payload.id || generateId_('PRG');

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][1] === payload.idSiswa && String(rows[i][2]) === kls && String(rows[i][3]) === sem) {
        rowIndex = i + 1;
        id = rows[i][0] || id;
        break;
      }
    }

    const rowData = [
      id,
      payload.idSiswa,
      kls,
      sem,
      String(payload.targetSurah || '').trim(),
      jumlahTercapai,
      totalSurah,
      persentase,
      JSON.stringify(detail),
      String(payload.catatanGuru || '').trim(),
      now
    ];

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return {
      success: true,
      message: 'Progres hafalan berhasil disimpan.',
      data: {
        id: id,
        idSiswa: payload.idSiswa,
        kelas: kls,
        semester: sem,
        targetSurah: rowData[4],
        jumlahTercapai: jumlahTercapai,
        totalSurah: totalSurah,
        persentase: persentase,
        detailSurah: detail,
        catatanGuru: rowData[9]
      }
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * ARSIP RAPOR PDF KE GOOGLE DRIVE & LOG SHEET DATA RAPOR
 * ============================================================================
 */
function archiveRaporPdfToDrive(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const folderId = props.getProperty(APP_PROP_FOLDER_ID) || initDatabase().folderId;
    const mainFolder = DriveApp.getFolderById(folderId);

    let raporFolder;
    const sub = mainFolder.getFoldersByName('Arsip_Rapor_PDF');
    raporFolder = sub.hasNext() ? sub.next() : mainFolder.createFolder('Arsip_Rapor_PDF');

    const base64Data = payload.base64Pdf.split(',')[1] || payload.base64Pdf;
    const decoded = Utilities.base64Decode(base64Data);
    const fileName = 'Rapor_Tahfidz_' + payload.namaSiswa.replace(/[^a-zA-Z0-9]/g, '_') + '_Kls' + payload.kelas + '_Sem' + payload.semester + '.pdf';
    const blob = Utilities.newBlob(decoded, 'application/pdf', fileName);

    // Hapus file rapor lama siswa yang sama agar folder tetap bersih
    const existing = raporFolder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    const file = raporFolder.createFile(blob);
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(SHEETS_SCHEMA.RAPOR.name);
    sheet.appendRow([
      generateId_('RPR'),
      payload.idSiswa,
      payload.namaSiswa,
      payload.kelas,
      payload.semester,
      payload.tahunAjaran || '',
      payload.persentase || 0,
      file.getId(),
      file.getUrl(),
      new Date().toISOString()
    ]);

    return { success: true, fileUrl: file.getUrl(), fileName: fileName };
  } finally {
    lock.releaseLock();
  }
}