const SPREADSHEET_ID = "1NTyd6qsyj95Q5og8ErAOWqZM_HroyzYxYBuJ1069jyc";
const SHEET_NAME = "user"; 
const DB_MP_SHEET_NAME = "Manpower";

/**
 * Normalisasi NRP untuk menghilangkan leading zeros saat pencocokan angka/string.
 * Contoh: "0113039" -> "113039"
 */
function normalizeNrp(val) {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  if (str.startsWith("'")) str = str.substring(1).trim();
  return str.replace(/^0+/, '');
}

/**
 * Format nilai NRP agar tersimpan sebagai Text murni dengan leading zero (misal: '01151181)
 */
function formatTextValue(val) {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  if (!str) return '';
  return str.startsWith("'") ? str : "'" + str;
}

// Header resmi pada sheet Manpower. Data organisasi dari sheet user hanya
// ditulis ke kolom ini; kolom data pribadi lain tetap tidak disentuh.
const MANPOWER_HEADER_ALIASES = {
  NRP: ['nrp', 'nik', 'id', 'user id'],
  NAMA: ['nama', 'name', 'karyawan', 'employee'],
  PERUSAHAAN: ['perusahaan', 'company', 'pt'],
  'KATEGORI AKUN': ['kategori akun', 'category kemitraan', 'category', 'kategori', 'kemitraan'],
  'LEVEL KARYAWAN': ['level karyawan', 'job rank', 'rank', 'level'],
  'GOLONGAN KARYAWAN': ['golongan karyawan', 'job group', 'jobgroup', 'group'],
  'STATUS KARYAWAN': ['status karyawan', 'status'],
  'SUB SECTION / SECTION': ['sub section / section', 'sub section', 'subsection']
};

function findHeaderByAliases(headers, aliases) {
  return headers.findIndex(header => aliases.includes(String(header).toLowerCase().trim()));
}

function mapUserValuesToManpowerHeaders(manpowerHeaders, values) {
  const mapped = {};
  Object.entries(MANPOWER_HEADER_ALIASES).forEach(([field, aliases]) => {
    const index = findHeaderByAliases(manpowerHeaders, aliases);
    if (index !== -1 && values[field] !== undefined) {
      mapped[manpowerHeaders[index]] = values[field];
    }
  });
  return mapped;
}

function getManpowerUserValues(headers, item) {
  return mapUserValuesToManpowerHeaders(headers, {
    NRP: formatTextValue(item.nrp),
    NAMA: String(item.nama || '').trim(),
    PERUSAHAAN: String(item.perusahaan || '').trim(),
    'KATEGORI AKUN': item.category || '',
    'LEVEL KARYAWAN': item.jobrank || '',
    'GOLONGAN KARYAWAN': item.jobgroup || '',
    'STATUS KARYAWAN': item.status || '',
    'SUB SECTION / SECTION': item.subsection || ''
  });
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Page')
    .setTitle('Manpower Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Mendapatkan data manpower dengan cepat pada saat page load awal (tanpa proses tulis binding berat).
 */
function getManpowerDataFast() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    let dbSheet = ss.getSheetByName(DB_MP_SHEET_NAME);
    
    if (!dbSheet) {
      dbSheet = ss.insertSheet(DB_MP_SHEET_NAME);
      dbSheet.appendRow(["NRP", "Nama", "Perusahaan"]);
    }

    const rows = sheet.getDataRange().getDisplayValues();
    if (!rows || rows.length <= 1) {
      return { data: [], dbFields: ["NRP", "Nama", "Perusahaan"] };
    }

    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const findCol = (keywords) => headers.findIndex(h => keywords.some(kw => h.includes(kw)));
    
    const nrpIdx = findCol(['nrp', 'nik', 'id', 'user id']);
    const namaIdx = findCol(['nama', 'name', 'karyawan', 'employee']);
    const posIdx = findCol(['position', 'jabatan', 'posisi']);
    const jgIdx = findCol(['job group', 'jobgroup', 'group']);
    const jrIdx = findCol(['job rank', 'rank', 'level']);
    const secIdx = headers.findIndex(h => h.includes('section') && !h.includes('sub'));
    const deptIdx = findCol(['departemen', 'department', 'divisi']);
    const subSecIdx = findCol(['sub section', 'subsection', 'sub']);
    const custIdx = findCol(['custodian', 'leader', 'atasan']);
    const statIdx = findCol(['status']);
    const compIdx = findCol(['perusahaan', 'company', 'pt']);
    const catIdx = findCol(['category kemitraan', 'category', 'kategori', 'kemitraan']);
    const roleIdx = findCol(['role', 'hak akses']);

    const actualNrpIdx = nrpIdx !== -1 ? nrpIdx : 0;
    const actualNamaIdx = namaIdx !== -1 ? namaIdx : 1;

    let dbMap = new Map();
    let dbFieldNames = ["NRP", "Nama", "Perusahaan"];
    const dbRows = dbSheet.getDataRange().getDisplayValues();
    
    if (dbRows && dbRows.length > 0) {
      const rawDbHeaders = dbRows[0].map(h => String(h).trim()).filter(Boolean);
      
      dbFieldNames = rawDbHeaders.filter(h => {
        const lower = h.toLowerCase();
        return lower !== 'superior';
      });
      
      const dbNrpIdx = rawDbHeaders.findIndex(h => h.toLowerCase() === 'nrp' || h.toLowerCase() === 'nik' || h.toLowerCase() === 'id');

      if (dbNrpIdx !== -1) {
        for (let j = 1; j < dbRows.length; j++) {
          let rNrp = String(dbRows[j][dbNrpIdx] || '').trim();
          if (rNrp) {
            let recordObj = {};
            rawDbHeaders.forEach((fieldName, colIdx) => {
              const lowerF = fieldName.toLowerCase();
              if (lowerF !== 'superior') {
                recordObj[fieldName] = dbRows[j][colIdx] !== undefined ? String(dbRows[j][colIdx]).trim() : '';
              }
            });
            dbMap.set(rNrp, recordObj);
            dbMap.set(normalizeNrp(rNrp), recordObj);
          }
        }
      }
    }

    if (dbFieldNames.length === 0) dbFieldNames = ["NRP", "Nama", "Perusahaan"];

    let data = [];
    for (let i = 1; i < rows.length; i++) {
      let r = rows[i];
      let currentNrp = String(r[actualNrpIdx] !== undefined ? r[actualNrpIdx] : '').trim();
      let currentNama = String(r[actualNamaIdx] !== undefined ? r[actualNamaIdx] : '').trim();

      if (!currentNrp && !currentNama) continue;

      let assignedRole = String(roleIdx !== -1 ? r[roleIdx] : 'user').toLowerCase().trim();
      if (!assignedRole) assignedRole = 'user';

      let dbRecord = dbMap.get(currentNrp) || dbMap.get(normalizeNrp(currentNrp)) || {};

      let sheetPerusahaan = String(compIdx !== -1 ? r[compIdx] : '').trim();
      if (sheetPerusahaan) {
        dbRecord['Perusahaan'] = sheetPerusahaan;
      } else if (dbRecord['Perusahaan']) {
        sheetPerusahaan = dbRecord['Perusahaan'];
      }

      data.push({
        rowIdx: i + 1,
        nrp: currentNrp,
        nama: currentNama,
        position: String(posIdx !== -1 ? r[posIdx] : '').trim(),
        jobgroup: String(jgIdx !== -1 ? r[jgIdx] : '').trim(),
        jobrank: String(jrIdx !== -1 ? r[jrIdx] : '').trim(),
        section: String(secIdx !== -1 ? r[secIdx] : '').trim(),
        departemen: String(deptIdx !== -1 ? r[deptIdx] : '').trim(),
        subsection: String(subSecIdx !== -1 ? r[subSecIdx] : '').trim(),
        custodian: String(custIdx !== -1 ? r[custIdx] : '').trim(),
        status: String(statIdx !== -1 ? r[statIdx] : 'ACTIVE').trim(),
        perusahaan: sheetPerusahaan,
        category: String(catIdx !== -1 ? r[catIdx] : '').trim(),
        role: assignedRole,
        dbRecord: dbRecord,
        dbFields: dbFieldNames
      });
    }

    return { data: data, dbFields: dbFieldNames };
  } catch (err) {
    return { data: [], dbFields: ["NRP", "Nama", "Perusahaan"], error: err.toString() };
  }
}

/**
 * Fungsi khusus untuk melakukan sinkronisasi binding dari sheet 'user' ke 'db_MP' 
 * hanya ketika tombol Synchronize ditekan secara manual.
 */
function executeManualSync() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    let dbSheet = ss.getSheetByName(DB_MP_SHEET_NAME);
    
    if (!dbSheet) {
      dbSheet = ss.insertSheet(DB_MP_SHEET_NAME);
      dbSheet.appendRow(["NRP", "Nama", "Perusahaan"]);
    }

    const rows = sheet.getDataRange().getDisplayValues();
    if (!rows || rows.length <= 1) {
      return getManpowerDataFast();
    }

    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    const findCol = (keywords) => headers.findIndex(h => keywords.some(kw => h.includes(kw)));
    
    const nrpIdx = findCol(['nrp', 'nik', 'id', 'user id']);
    const namaIdx = findCol(['nama', 'name', 'karyawan', 'employee']);
    const secIdx = headers.findIndex(h => h.includes('section') && !h.includes('sub'));
    const subSecIdx = findCol(['sub section', 'subsection', 'sub']);
    const posIdx = findCol(['position', 'jabatan', 'posisi']);
    const jgIdx = findCol(['job group', 'jobgroup', 'group']);
    const jrIdx = findCol(['job rank', 'rank', 'level']);
    const statIdx = findCol(['status']);
    const catIdx = findCol(['category kemitraan', 'category', 'kategori', 'kemitraan']);
    const compIdx = findCol(['perusahaan', 'company', 'pt']);

    const actualNrpIdx = nrpIdx !== -1 ? nrpIdx : 0;
    const actualNamaIdx = namaIdx !== -1 ? namaIdx : 1;

    const dbRows = dbSheet.getDataRange().getDisplayValues();
    let rawDbHeaders = dbRows.length > 0 ? dbRows[0].map(h => String(h).trim()).filter(Boolean) : ["NRP", "Nama", "Perusahaan"];
    
    let dbNrpColIdx = rawDbHeaders.findIndex(h => h.toLowerCase() === 'nrp' || h.toLowerCase() === 'nik' || h.toLowerCase() === 'id');
    if (dbNrpColIdx === -1) dbNrpColIdx = 0;

    let dbRowMap = new Map();
    for (let j = 1; j < dbRows.length; j++) {
      let rNrp = String(dbRows[j][dbNrpColIdx] || '').trim();
      if (rNrp) {
        dbRowMap.set(rNrp, j + 1);
        dbRowMap.set(normalizeNrp(rNrp), j + 1);
      }
    }

    for (let i = 1; i < rows.length; i++) {
      let r = rows[i];
      let currentNrp = String(r[actualNrpIdx] !== undefined ? r[actualNrpIdx] : '').trim();
      let currentNama = String(r[actualNamaIdx] !== undefined ? r[actualNamaIdx] : '').trim();
      if (!currentNrp && !currentNama) continue;

      let formattedNrp = formatTextValue(currentNrp);

      let existingDbRow = dbRowMap.get(currentNrp) || dbRowMap.get(normalizeNrp(currentNrp));

      let syncValues = mapUserValuesToManpowerHeaders(rawDbHeaders, {
        NRP: formattedNrp,
        NAMA: currentNama,
        'KATEGORI AKUN': catIdx !== -1 ? String(r[catIdx] || '').trim() : '',
        'LEVEL KARYAWAN': jrIdx !== -1 ? String(r[jrIdx] || '').trim() : '',
        'GOLONGAN KARYAWAN': jgIdx !== -1 ? String(r[jgIdx] || '').trim() : '',
        'STATUS KARYAWAN': statIdx !== -1 ? String(r[statIdx] || '').trim() : ''
      });

      if (existingDbRow) {
        rawDbHeaders.forEach((colName, colIdx) => {
          let matchKey = Object.keys(syncValues).find(k => k.toLowerCase() === colName.toLowerCase());
          if (matchKey !== undefined) {
            if (colIdx === dbNrpColIdx) {
              dbSheet.getRange(existingDbRow, colIdx + 1).setNumberFormat('@').setValue(formattedNrp);
            } else {
              dbSheet.getRange(existingDbRow, colIdx + 1).setValue(syncValues[matchKey]);
            }
          }
        });
      } else {
        let newDbRowArr = rawDbHeaders.map(colName => {
          let lowerC = colName.toLowerCase();
          if (lowerC === 'nrp' || lowerC === 'nik' || lowerC === 'id') return formattedNrp;
          let matchKey = Object.keys(syncValues).find(k => k.toLowerCase() === lowerC);
          return matchKey !== undefined ? syncValues[matchKey] : '';
        });
        dbSheet.appendRow(newDbRowArr);
        let lastRow = dbSheet.getLastRow();
        dbSheet.getRange(lastRow, dbNrpColIdx + 1).setNumberFormat('@').setValue(formattedNrp);
        dbRowMap.set(currentNrp, lastRow);
        dbRowMap.set(normalizeNrp(currentNrp), lastRow);
      }
    }

    return getManpowerDataFast();
  } catch (err) {
    return { data: [], dbFields: ["NRP", "Nama", "Perusahaan"], error: err.toString() };
  }
}

function createNewManpower(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const item = payload.item;
    const dbUpdates = payload.dbUpdates || {};

    const rows = sheet.getDataRange().getDisplayValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    
    const findCol = (keywords) => headers.findIndex(h => keywords.some(kw => h.includes(kw)));
    const nrpIdx = findCol(['nrp', 'nik', 'id']);
    
    const normItemNrp = normalizeNrp(item.nrp);
    for (let i = 1; i < rows.length; i++) {
      let currentSheetNrp = String(rows[i][nrpIdx !== -1 ? nrpIdx : 0] || '').trim();
      if (normalizeNrp(currentSheetNrp) === normItemNrp) {
        return { success: false, error: 'NRP ' + item.nrp + ' sudah terdaftar dalam sistem!' };
      }
    }

    let newRow = new Array(headers.length).fill('');
    const mapHeader = (keywords, val) => {
      let idx = findCol(keywords);
      if (idx !== -1) newRow[idx] = val;
    };

    const formattedNrp = formatTextValue(item.nrp);

    mapHeader(['nrp', 'nik', 'id'], formattedNrp);
    mapHeader(['nama', 'name', 'karyawan'], item.nama);
    mapHeader(['position', 'jabatan'], item.position);
    mapHeader(['job group', 'jobgroup', 'group'], item.jobgroup);
    mapHeader(['job rank', 'rank', 'level'], item.jobrank);
    mapHeader(['section'], item.section);
    mapHeader(['departemen', 'department', 'divisi'], item.departemen);
    mapHeader(['sub section', 'subsection', 'sub'], item.subsection);
    mapHeader(['custodian', 'leader', 'atasan'], item.custodian);
    mapHeader(['status'], item.status || 'ACTIVE');
    mapHeader(['perusahaan', 'company', 'pt'], item.perusahaan);
    mapHeader(['category kemitraan', 'category', 'kategori', 'kemitraan'], item.category);
    mapHeader(['role', 'hak akses'], item.role || 'user');

    sheet.appendRow(newRow);

    const lastUserRow = sheet.getLastRow();
    const actualNrpColIdx = nrpIdx !== -1 ? nrpIdx : 0;
    sheet.getRange(lastUserRow, actualNrpColIdx + 1).setNumberFormat('@').setValue(formattedNrp);

    let dbSheet = ss.getSheetByName(DB_MP_SHEET_NAME);
    if (!dbSheet) {
      dbSheet = ss.insertSheet(DB_MP_SHEET_NAME);
      dbSheet.appendRow(["NRP", "Nama", "Perusahaan"]);
    }

    let dbRows = dbSheet.getDataRange().getDisplayValues();
    let dbHeaders = dbRows.length > 0 ? dbRows[0].map(h => String(h).trim()) : ["NRP", "Nama", "Perusahaan"];

    delete dbUpdates['Superior'];
    delete dbUpdates['superior'];

    Object.keys(dbUpdates).forEach(k => {
      const lower = k.toLowerCase().trim();
      if (lower === 'section' || lower === 'jabatan' || lower === 'perusahaan') {
        delete dbUpdates[k];
      }
    });

    Object.keys(dbUpdates).forEach(field => {
      let trimmedField = String(field || '').trim();
      let foundCol = dbHeaders.findIndex(h => h.toLowerCase() === trimmedField.toLowerCase());
      if (foundCol === -1 && trimmedField) {
        dbSheet.getRange(1, dbHeaders.length + 1).setValue(trimmedField);
        dbHeaders.push(trimmedField);
      }
    });

    let finalRecordValues = { ...dbUpdates, ...getManpowerUserValues(dbHeaders, item) };
    let dbNrpColIdx = dbHeaders.findIndex(h => ['nrp', 'nik', 'id', 'user id'].includes(h.toLowerCase().trim()));
    if (dbNrpColIdx === -1) dbNrpColIdx = 0;

    let newDbRowArr = dbHeaders.map((colName, colIdx) => {
      if (colIdx === dbNrpColIdx) return formattedNrp;
      let trimmedCol = colName.trim();
      let matchKey = Object.keys(finalRecordValues).find(k => k.toLowerCase().trim() === trimmedCol.toLowerCase());
      return matchKey !== undefined ? finalRecordValues[matchKey] : '';
    });
    dbSheet.appendRow(newDbRowArr);

    const lastDbRow = dbSheet.getLastRow();
    dbSheet.getRange(lastDbRow, dbNrpColIdx + 1).setNumberFormat('@').setValue(formattedNrp);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function updateSingleManpower(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const item = payload.item;
    let dbUpdates = payload.dbUpdates || {};

    delete dbUpdates['Superior'];
    delete dbUpdates['superior'];

    Object.keys(dbUpdates).forEach(k => {
      const lower = k.toLowerCase().trim();
      if (lower === 'sub section / section' || lower === 'sub section' || lower === 'subsection') dbUpdates[k] = item.subsection;
      else if (lower === 'section' || lower === 'jabatan') delete dbUpdates[k];
    });

    let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const rows = sheet.getDataRange().getDisplayValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    
    const findCol = (keywords) => headers.findIndex(h => keywords.some(kw => h.includes(kw)));
    const nrpIdx = findCol(['nrp', 'nik', 'id']);
    const namaIdx = findCol(['nama', 'name', 'karyawan']);
    const posIdx = findCol(['position', 'jabatan']);
    const jgIdx = findCol(['job group', 'jobgroup', 'group']);
    const jrIdx = findCol(['job rank', 'rank', 'level']);
    const secIdx = headers.findIndex(h => h.includes('section') && !h.includes('sub'));
    const deptIdx = findCol(['departemen', 'department', 'divisi']);
    const subSecIdx = findCol(['sub section', 'subsection', 'sub']);
    const custIdx = findCol(['custodian', 'leader', 'atasan']);
    const statIdx = findCol(['status']);
    const compIdx = findCol(['perusahaan', 'company', 'pt']);
    const catIdx = findCol(['category kemitraan', 'category', 'kategori', 'kemitraan']);
    const roleIdx = findCol(['role', 'hak akses']);

    let targetRow = parseInt(item.rowIdx);
    let originalNrp = '';
    if (targetRow >= 2 && targetRow <= rows.length && nrpIdx !== -1) {
      originalNrp = String(rows[targetRow - 1][nrpIdx] || '').trim();
    }
    const normItemNrp = normalizeNrp(item.nrp);
    const normOriginalNrp = normalizeNrp(originalNrp);

    if (!targetRow || targetRow < 2) {
      for (let i = 1; i < rows.length; i++) {
        let currentSheetNrp = String(rows[i][nrpIdx !== -1 ? nrpIdx : 0] || '').trim();
        let normCur = normalizeNrp(currentSheetNrp);
        if ((normItemNrp && normCur === normItemNrp) || (normOriginalNrp && normCur === normOriginalNrp)) {
          targetRow = i + 1;
          break;
        }
      }
    }

    const formattedNrp = formatTextValue(item.nrp);

    if (targetRow >= 2 && targetRow <= rows.length) {
      if (nrpIdx !== -1) {
        sheet.getRange(targetRow, nrpIdx + 1).setNumberFormat('@').setValue(formattedNrp);
      }
      if (namaIdx !== -1) sheet.getRange(targetRow, namaIdx + 1).setValue(item.nama);
      if (posIdx !== -1) sheet.getRange(targetRow, posIdx + 1).setValue(item.position);
      if (jgIdx !== -1) sheet.getRange(targetRow, jgIdx + 1).setValue(item.jobgroup);
      if (jrIdx !== -1) sheet.getRange(targetRow, jrIdx + 1).setValue(item.jobrank);
      if (secIdx !== -1) sheet.getRange(targetRow, secIdx + 1).setValue(item.section);
      if (deptIdx !== -1) sheet.getRange(targetRow, deptIdx + 1).setValue(item.departemen);
      if (subSecIdx !== -1) sheet.getRange(targetRow, subSecIdx + 1).setValue(item.subsection);
      if (custIdx !== -1) sheet.getRange(targetRow, custIdx + 1).setValue(item.custodian);
      if (statIdx !== -1) sheet.getRange(targetRow, statIdx + 1).setValue(item.status);
      if (compIdx !== -1) sheet.getRange(targetRow, compIdx + 1).setValue(item.perusahaan);
      if (catIdx !== -1) sheet.getRange(targetRow, catIdx + 1).setValue(item.category);
      if (roleIdx !== -1) sheet.getRange(targetRow, roleIdx + 1).setValue(item.role);
    }

    let dbSheet = ss.getSheetByName(DB_MP_SHEET_NAME);
    if (!dbSheet) {
      dbSheet = ss.insertSheet(DB_MP_SHEET_NAME);
      dbSheet.appendRow(["NRP", "Nama", "Perusahaan"]);
    }

    // Gunakan getValues() agar NRP dengan leading zero tetap terbaca utuh
    let dbAllValues = dbSheet.getDataRange().getValues();
    let dbDisplayValues = dbSheet.getDataRange().getDisplayValues();
    let dbHeaders = dbAllValues.length > 0 ? dbAllValues[0].map(h => String(h).trim()) : ["NRP", "Nama", "Perusahaan"];
    
    Object.keys(dbUpdates).forEach(field => {
      let trimmedField = String(field || '').trim();
      let foundCol = dbHeaders.findIndex(h => h.toLowerCase().trim() === trimmedField.toLowerCase());
      if (foundCol === -1 && trimmedField) {
        dbSheet.getRange(1, dbHeaders.length + 1).setValue(trimmedField);
        dbHeaders.push(trimmedField);
      }
    });

    // Refresh setelah kemungkinan penambahan kolom baru
    dbAllValues = dbSheet.getDataRange().getValues();
    dbDisplayValues = dbSheet.getDataRange().getDisplayValues();
    dbHeaders = dbAllValues[0].map(h => String(h).trim());
    
    let dbNrpColIdx = dbHeaders.findIndex(h => {
      const lower = h.toLowerCase().trim();
      return lower === 'nrp' || lower === 'nik' || lower === 'id' || lower === 'user id';
    });
    if (dbNrpColIdx === -1) dbNrpColIdx = 0;

    let targetDbRow = -1;
    for (let j = 1; j < dbAllValues.length; j++) {
      // Coba raw value (getValues) dulu, lalu display value sebagai fallback
      let rawNrp = String(dbAllValues[j][dbNrpColIdx] || '').trim();
      let dispNrp = String(dbDisplayValues[j][dbNrpColIdx] || '').trim();
      let normRaw = normalizeNrp(rawNrp);
      let normDisp = normalizeNrp(dispNrp);
      if (
        (normItemNrp && (normRaw === normItemNrp || normDisp === normItemNrp)) ||
        (normOriginalNrp && (normRaw === normOriginalNrp || normDisp === normOriginalNrp))
      ) {
        targetDbRow = j + 1;
        break;
      }
    }

    let finalRecordValues = { ...dbUpdates, ...getManpowerUserValues(dbHeaders, item) };

    if (targetDbRow !== -1) {
      dbHeaders.forEach((colName, colIdx) => {
        let trimmedCol = colName.trim();
        let matchKey = Object.keys(finalRecordValues).find(k => k.toLowerCase().trim() === trimmedCol.toLowerCase());
        if (matchKey !== undefined) {
          if (colIdx === dbNrpColIdx) {
            dbSheet.getRange(targetDbRow, colIdx + 1).setNumberFormat('@').setValue(formattedNrp);
          } else {
            dbSheet.getRange(targetDbRow, colIdx + 1).setValue(finalRecordValues[matchKey]);
          }
        }
      });
    } else {
      let newRowArr = dbHeaders.map((colName, colIdx) => {
        if (colIdx === dbNrpColIdx) return formattedNrp;
        let trimmedCol = colName.trim();
        let matchKey = Object.keys(finalRecordValues).find(k => k.toLowerCase().trim() === trimmedCol.toLowerCase());
        return matchKey !== undefined ? finalRecordValues[matchKey] : '';
      });
      dbSheet.appendRow(newRowArr);
      const lastDbRow = dbSheet.getLastRow();
      dbSheet.getRange(lastDbRow, dbNrpColIdx + 1).setNumberFormat('@').setValue(formattedNrp);
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function updateRolesBatch(nrpList, newRole) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().trim());
    
    let nrpIdx = headers.findIndex(h => h.includes('nrp') || h.includes('nik') || h.includes('id'));
    let roleIdx = headers.findIndex(h => h.includes('role') || h.includes('hak akses'));

    if (nrpIdx === -1) nrpIdx = 0;
    if (roleIdx === -1) {
      roleIdx = headers.length;
      sheet.getRange(1, roleIdx + 1).setValue("Role");
    }

    const nrpNormSet = new Set(nrpList.map(n => normalizeNrp(n)));
    let count = 0;

    for (let i = 1; i < rows.length; i++) {
      let sheetNrp = String(rows[i][nrpIdx] || '').trim();
      if (nrpNormSet.has(normalizeNrp(sheetNrp))) {
        sheet.getRange(i + 1, roleIdx + 1).setValue(newRole);
        count++;
      }
    }

    return { success: true, count: count };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function deleteManpowerRows(nrpList) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let mainSheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const rows = mainSheet.getDataRange().getValues();
    const nrpNormSet = new Set(nrpList.map(n => normalizeNrp(n)));
    
    if (rows && rows.length > 1) {
      const headers = rows[0].map(h => String(h).toLowerCase().trim());
      let nrpIdx = headers.findIndex(h => h.includes('nrp') || h.includes('nik') || h.includes('id') || h.includes('user id'));
      if (nrpIdx === -1) nrpIdx = 0;

      for (let i = rows.length - 1; i >= 1; i--) {
        let sheetNrp = String(rows[i][nrpIdx] || '').trim();
        if (nrpNormSet.has(normalizeNrp(sheetNrp))) {
          mainSheet.deleteRow(i + 1);
        }
      }
    }

    let dbSheet = ss.getSheetByName(DB_MP_SHEET_NAME);
    if (dbSheet) {
      const dbRows = dbSheet.getDataRange().getValues();
      if (dbRows && dbRows.length > 1) {
        const dbHeaders = dbRows[0].map(h => String(h).toLowerCase().trim());
        let dbNrpIdx = dbHeaders.findIndex(h => h.toLowerCase() === 'nrp' || h.toLowerCase() === 'nik' || h.toLowerCase() === 'id');
        if (dbNrpIdx === -1) dbNrpIdx = 0;

        for (let j = dbRows.length - 1; j >= 1; j--) {
          let dbNrp = String(dbRows[j][dbNrpIdx] || '').trim();
          if (nrpNormSet.has(normalizeNrp(dbNrp))) {
            dbSheet.deleteRow(j + 1);
          }
        }
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
