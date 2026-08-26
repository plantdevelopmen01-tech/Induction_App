const SPREADSHEET_ID = "1NTyd6qsyj95Q5og8ErAOWqZM_HroyzYxYBuJ1069jyc";
const SHEET_NAME = "user"; 
const DB_MP_SHEET_NAME = "db_MP";

/**
 * Normalisasi NRP untuk menghilangkan leading zeros saat pencocokan angka/string.
 */
function normalizeNrp(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().replace(/^0+/, '');
}

/**
 * Format nilai NRP agar tersimpan sebagai Text murni dengan leading zero.
 */
function formatTextValue(val) {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  if (!str) return '';
  return str.startsWith("'") ? str : "'" + str;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Page')
    .setTitle('Manpower Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Mendapatkan data manpower dengan cepat pada saat page load awal.
 */
function getManpowerDataFast() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    let dbSheet = ss.getSheetByName(DB_MP_SHEET_NAME);
    
    if (!dbSheet) {
      dbSheet = ss.insertSheet(DB_MP_SHEET_NAME);
      dbSheet.appendRow(["NRP", "Nama", "Perusahaan", "Status"]);
    }

    const rows = sheet.getDataRange().getDisplayValues();
    if (!rows || rows.length <= 1) {
      return { data: [], dbFields: ["NRP", "Nama", "Perusahaan", "Status"] };
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
    const catIdx = findCol(['category', 'kategori']);
    const roleIdx = findCol(['role', 'hak akses']);

    const actualNrpIdx = nrpIdx !== -1 ? nrpIdx : 0;
    const actualNamaIdx = namaIdx !== -1 ? namaIdx : 1;

    let dbMap = new Map();
    let dbFieldNames = ["NRP", "Nama", "Perusahaan", "Status"];
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

    if (dbFieldNames.length === 0) dbFieldNames = ["NRP", "Nama", "Perusahaan", "Status"];

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

      let sheetStatus = String(statIdx !== -1 ? r[statIdx] : 'ACTIVE').trim();
      if (sheetStatus) {
        dbRecord['Status'] = sheetStatus;
      } else if (dbRecord['Status']) {
        sheetStatus = dbRecord['Status'];
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
        status: sheetStatus || 'ACTIVE',
        perusahaan: sheetPerusahaan,
        category: String(catIdx !== -1 ? r[catIdx] : '').trim(),
        role: assignedRole,
        dbRecord: dbRecord,
        dbFields: dbFieldNames
      });
    }

    return { data: data, dbFields: dbFieldNames };
  } catch (err) {
    return { data: [], dbFields: ["NRP", "Nama", "Perusahaan", "Status"], error: err.toString() };
  }
}

/**
 * Fungsi khusus untuk melakukan sinkronisasi binding dari sheet 'user' ke 'db_MP' 
 * termasuk header Status secara manual saat tombol Synchronize ditekan.
 */
function executeManualSync() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    let dbSheet = ss.getSheetByName(DB_MP_SHEET_NAME);
    
    if (!dbSheet) {
      dbSheet = ss.insertSheet(DB_MP_SHEET_NAME);
      dbSheet.appendRow(["NRP", "Nama", "Perusahaan", "Status"]);
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
    const jgIdx = findCol(['job group', 'jobgroup', 'group']);
    const compIdx = findCol(['perusahaan', 'company', 'pt']);
    const statIdx = findCol(['status']);

    const actualNrpIdx = nrpIdx !== -1 ? nrpIdx : 0;
    const actualNamaIdx = namaIdx !== -1 ? namaIdx : 1;

    const dbRows = dbSheet.getDataRange().getDisplayValues();
    let rawDbHeaders = dbRows.length > 0 ? dbRows[0].map(h => String(h).trim()).filter(Boolean) : ["NRP", "Nama", "Perusahaan", "Status"];
    
    let dbNrpColIdx = rawDbHeaders.findIndex(h => h.toLowerCase() === 'nrp' || h.toLowerCase() === 'nik' || h.toLowerCase() === 'id');
    if (dbNrpColIdx === -1) dbNrpColIdx = 0;

    const requiredBindFields = ['Section', 'Sub Section / Section', 'JABATAN', 'Perusahaan', 'Status'];
    requiredBindFields.forEach(field => {
      let found = rawDbHeaders.findIndex(h => h.toLowerCase() === field.toLowerCase());
      if (found === -1) {
        dbSheet.getRange(1, rawDbHeaders.length + 1).setValue(field);
        rawDbHeaders.push(field);
      }
    });

    const updatedDbRows = dbSheet.getDataRange().getDisplayValues();
    rawDbHeaders = updatedDbRows[0].map(h => String(h).trim());

    let dbRowMap = new Map();
    for (let j = 1; j < updatedDbRows.length; j++) {
      let rNrp = String(updatedDbRows[j][dbNrpColIdx] || '').trim();
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

      let sheetSec = secIdx !== -1 ? String(r[secIdx] || '').trim() : '';
      let sheetSubSec = subSecIdx !== -1 ? String(r[subSecIdx] || '').trim() : '';
      let sheetJg = jgIdx !== -1 ? String(r[jgIdx] || '').trim() : '';
      let sheetComp = compIdx !== -1 ? String(r[compIdx] || '').trim() : '';
      let sheetStatus = statIdx !== -1 ? String(r[statIdx] || '').trim() : 'ACTIVE';
      let formattedNrp = formatTextValue(currentNrp);

      let existingDbRow = dbRowMap.get(currentNrp) || dbRowMap.get(normalizeNrp(currentNrp));

      let syncValues = {
        'Section': sheetSec,
        'Sub Section / Section': sheetSubSec,
        'JABATAN': sheetJg,
        'Perusahaan': sheetComp,
        'Status': sheetStatus,
        'Nama': currentNama
      };

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
    return { data: [], dbFields: ["NRP", "Nama", "Perusahaan", "Status"], error: err.toString() };
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
    mapHeader(['category', 'kategori'], item.category);
    mapHeader(['role', 'hak akses'], item.role || 'user');

    sheet.appendRow(newRow);

    const lastUserRow = sheet.getLastRow();
    const actualNrpColIdx = nrpIdx !== -1 ? nrpIdx : 0;
    sheet.getRange(lastUserRow, actualNrpColIdx + 1).setNumberFormat('@').setValue(formattedNrp);

    let dbSheet = ss.getSheetByName(DB_MP_SHEET_NAME);
    if (!dbSheet) {
      dbSheet = ss.insertSheet(DB_MP_SHEET_NAME);
      dbSheet.appendRow(["NRP", "Nama", "Perusahaan", "Status"]);
    }

    let dbRows = dbSheet.getDataRange().getDisplayValues();
    let dbHeaders = dbRows.length > 0 ? dbRows[0].map(h => String(h).trim()) : ["NRP", "Nama", "Perusahaan", "Status"];

    dbUpdates['Perusahaan'] = item.perusahaan;
    dbUpdates['Status'] = item.status || 'ACTIVE';

    Object.keys(dbUpdates).forEach(field => {
      let foundCol = dbHeaders.findIndex(h => h.toLowerCase() === field.toLowerCase());
      if (foundCol === -1) {
        dbSheet.getRange(1, dbHeaders.length + 1).setValue(field);
        dbHeaders.push(field);
      }
    });

    let finalRecordValues = { 
      NRP: formattedNrp, 
      Nama: String(item.nama).trim(), 
      ...dbUpdates 
    };

    let newDbRowArr = dbHeaders.map(colName => {
      let matchKey = Object.keys(finalRecordValues).find(k => k.toLowerCase() === colName.toLowerCase());
      return matchKey !== undefined ? finalRecordValues[matchKey] : '';
    });
    dbSheet.appendRow(newDbRowArr);

    const lastDbRow = dbSheet.getLastRow();
    let dbNrpColIdx = dbHeaders.findIndex(h => h.toLowerCase() === 'nrp' || h.toLowerCase() === 'nik' || h.toLowerCase() === 'id');
    if (dbNrpColIdx === -1) dbNrpColIdx = 0;
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

    dbUpdates['Perusahaan'] = item.perusahaan;
    dbUpdates['Status'] = item.status || 'ACTIVE';

    Object.keys(dbUpdates).forEach(k => {
      const lower = k.toLowerCase().trim();
      if (lower === 'section') dbUpdates[k] = item.section;
      else if (lower === 'sub section / section' || lower === 'sub section' || lower === 'subsection') dbUpdates[k] = item.subsection;
      else if (lower === 'jabatan') dbUpdates[k] = item.jobgroup;
      else if (lower === 'perusahaan') dbUpdates[k] = item.perusahaan;
      else if (lower === 'status') dbUpdates[k] = item.status;
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
    const catIdx = findCol(['category', 'kategori']);
    const roleIdx = findCol(['role', 'hak akses']);

    let targetRow = parseInt(item.rowIdx);
    const normItemNrp = normalizeNrp(item.nrp);

    if (!targetRow || targetRow < 2) {
      for (let i = 1; i < rows.length; i++) {
        let currentSheetNrp = String(rows[i][nrpIdx !== -1 ? nrpIdx : 0] || '').trim();
        if (normalizeNrp(currentSheetNrp) === normItemNrp) {
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
      dbSheet.appendRow(["NRP", "Nama", "Perusahaan", "Status"]);
    }

    let dbRows = dbSheet.getDataRange().getDisplayValues();
    let dbHeaders = dbRows.length > 0 ? dbRows[0].map(h => String(h).trim()) : ["NRP", "Nama", "Perusahaan", "Status"];
    
    Object.keys(dbUpdates).forEach(field => {
      let foundCol = dbHeaders.findIndex(h => h.toLowerCase() === field.toLowerCase());
      if (foundCol === -1) {
        dbSheet.getRange(1, dbHeaders.length + 1).setValue(field);
        dbHeaders.push(field);
      }
    });

    dbRows = dbSheet.getDataRange().getDisplayValues();
    dbHeaders = dbRows[0].map(h => String(h).trim());
    
    let dbNrpColIdx = dbHeaders.findIndex(h => h.toLowerCase() === 'nrp' || h.toLowerCase() === 'nik' || h.toLowerCase() === 'id');
    if (dbNrpColIdx === -1) dbNrpColIdx = 0;

    let targetDbRow = -1;
    for (let j = 1; j < dbRows.length; j++) {
      let rNrp = String(dbRows[j][dbNrpColIdx] || '').trim();
      if (normalizeNrp(rNrp) === normItemNrp) {
        targetDbRow = j + 1;
        break;
      }
    }

    let finalRecordValues = { 
      NRP: formattedNrp, 
      Nama: String(item.nama).trim(), 
      ...dbUpdates 
    };

    if (targetDbRow !== -1) {
      dbHeaders.forEach((colName, colIdx) => {
        let matchKey = Object.keys(finalRecordValues).find(k => k.toLowerCase() === colName.toLowerCase());
        if (matchKey !== undefined) {
          if (colIdx === dbNrpColIdx) {
            dbSheet.getRange(targetDbRow, colIdx + 1).setNumberFormat('@').setValue(formattedNrp);
          } else {
            dbSheet.getRange(targetDbRow, colIdx + 1).setValue(finalRecordValues[matchKey]);
          }
        }
      });
    } else {
      let newRowArr = dbHeaders.map(colName => {
        let matchKey = Object.keys(finalRecordValues).find(k => k.toLowerCase() === colName.toLowerCase());
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
    
    for (let i = rows.length - 1; i >= 1; i--) {
      let sheetNrp = String(rows[i][0] || '').trim();
      if (nrpNormSet.has(normalizeNrp(sheetNrp))) {
        mainSheet.deleteRow(i + 1);
      }
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
