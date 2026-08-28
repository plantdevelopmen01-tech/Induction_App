const SPREADSHEET_ID = "1NTyd6qsyj95Q5og8ErAOWqZM_HroyzYxYBuJ1069jyc";
const SHEET_NAME = "user"; 
const DB_MP_SHEET_NAME = "Manpower"; // Diperbarui dari db_MP menjadi Manpower

/**
 * Normalisasi NRP: menghapus leading zero dan spasi tak terlihat
 * @param {string|number} val
 * @return {string}
 */
function normalizeNrp(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim().replace(/^0+/, '');
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Page')
    .evaluate()
    .setTitle('Manpower Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper untuk include HTML partial jika diperlukan
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Membaca data gabungan dari sheet 'user' dan sheet 'Manpower' secara cepat & presisi
 * @return {Object} { data: Array, dbFields: Array }
 */
function getManpowerDataFast() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(SHEET_NAME);
    const mpSheet = ss.getSheetByName(DB_MP_SHEET_NAME);

    if (!userSheet) {
      throw new Error(`Sheet utama '${SHEET_NAME}' tidak ditemukan.`);
    }

    // Read User Sheet
    const userValues = userSheet.getDataRange().getValues();
    if (userValues.length < 2) {
      return { data: [], dbFields: [] };
    }

    const userHeaders = userValues[0].map(h => String(h).trim());
    const nrpIdx = userHeaders.findIndex(h => h.toLowerCase() === 'nrp');
    const namaIdx = userHeaders.findIndex(h => h.toLowerCase() === 'nama');
    const posIdx = userHeaders.findIndex(h => h.toLowerCase() === 'position' || h.toLowerCase() === 'jabatan');
    const jgIdx = userHeaders.findIndex(h => h.toLowerCase() === 'job group' || h.toLowerCase() === 'jobgroup');
    const jrIdx = userHeaders.findIndex(h => h.toLowerCase() === 'job rank' || h.toLowerCase() === 'jobrank');
    const secIdx = userHeaders.findIndex(h => h.toLowerCase() === 'section');
    const subSecIdx = userHeaders.findIndex(h => h.toLowerCase() === 'sub section' || h.toLowerCase() === 'subsection');
    const custIdx = userHeaders.findIndex(h => h.toLowerCase() === 'custodian' || h.toLowerCase() === 'leader');
    const deptIdx = userHeaders.findIndex(h => h.toLowerCase() === 'departemen' || h.toLowerCase() === 'department');
    const statIdx = userHeaders.findIndex(h => h.toLowerCase() === 'status');
    const compIdx = userHeaders.findIndex(h => h.toLowerCase() === 'perusahaan' || h.toLowerCase() === 'company');
    const catIdx = userHeaders.findIndex(h => h.toLowerCase() === 'category' || h.toLowerCase() === 'kategori');
    const roleIdx = userHeaders.findIndex(h => h.toLowerCase() === 'role');

    // Read Manpower Sheet (Relasi Sekunder)
    let mpMap = new Map();
    let dbFields = [];

    if (mpSheet) {
      const mpValues = mpSheet.getDataRange().getValues();
      if (mpValues.length > 0) {
        dbFields = mpValues[0].map(f => String(f).trim());
        const mpNrpIdx = dbFields.findIndex(f => f.toLowerCase() === 'nrp');

        for (let i = 1; i < mpValues.length; i++) {
          const row = mpValues[i];
          const rawNrp = mpNrpIdx !== -1 ? row[mpNrpIdx] : '';
          const normKey = normalizeNrp(rawNrp);
          
          if (normKey) {
            let record = {};
            dbFields.forEach((field, fIdx) => {
              record[field] = row[fIdx] !== undefined ? row[fIdx] : '';
            });
            mpMap.set(normKey, { rowIndex: i + 1, record: record });
          }
        }
      }
    }

    // Assemble merged data array
    const manpowerData = [];
    for (let i = 1; i < userValues.length; i++) {
      const row = userValues[i];
      const rawNrp = nrpIdx !== -1 ? row[nrpIdx] : '';
      const normNrp = normalizeNrp(rawNrp);

      if (!normNrp && !row[namaIdx]) continue; // Skip empty rows

      const userStatus = statIdx !== -1 ? String(row[statIdx]).trim() : 'ACTIVE';
      const userSec = secIdx !== -1 ? String(row[secIdx]).trim() : '';
      const userSubSec = subSecIdx !== -1 ? String(row[subSecIdx]).trim() : '';
      const userJg = jgIdx !== -1 ? String(row[jgIdx]).trim() : '';

      // Fetch matching record from Manpower sheet map
      const mpData = mpMap.get(normNrp) || { rowIndex: null, record: {} };
      let dbRecord = { ...mpData.record };

      // BINDING OTOMATIS STATUS & RELASI KE SHEET MANPOWER
      dbRecord['Status'] = userStatus.toUpperCase();
      dbRecord['Section'] = userSec.toUpperCase();
      dbRecord['Sub Section / Section'] = userSubSec.toUpperCase();
      dbRecord['JABATAN'] = userJg.toUpperCase();

      manpowerData.push({
        rowIdx: i + 1,
        nrp: String(rawNrp).trim(),
        nama: namaIdx !== -1 ? String(row[namaIdx]).trim() : '',
        position: posIdx !== -1 ? String(row[posIdx]).trim() : '',
        jobgroup: userJg,
        jobrank: jrIdx !== -1 ? String(row[jrIdx]).trim() : '',
        section: userSec,
        subsection: userSubSec,
        custodian: custIdx !== -1 ? String(row[custIdx]).trim() : '',
        departemen: deptIdx !== -1 ? String(row[deptIdx]).trim() : '',
        status: userStatus,
        perusahaan: compIdx !== -1 ? String(row[compIdx]).trim() : '',
        category: catIdx !== -1 ? String(row[catIdx]).trim() : '',
        role: roleIdx !== -1 ? String(row[roleIdx]).trim().toLowerCase() : 'user',
        dbRecord: dbRecord,
        mpRowIdx: mpData.rowIndex
      });
    }

    return {
      data: manpowerData,
      dbFields: dbFields
    };

  } catch (err) {
    Logger.log('Error in getManpowerDataFast: ' + err.toString());
    throw new Error('Gagal mengambil data: ' + err.message);
  }
}

function executeManualSync() {
  return getManpowerDataFast();
}

/**
 * Menambahkan data manpower baru ke sheet 'user' dan sheet 'Manpower'
 */
function createNewManpower(payload) {
  try {
    const item = payload.item;
    const dbUpdates = payload.dbUpdates || {};

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(SHEET_NAME);
    const mpSheet = ss.getSheetByName(DB_MP_SHEET_NAME);

    if (!userSheet) throw new Error(`Sheet '${SHEET_NAME}' tidak ditemukan.`);

    const userHeaders = userSheet.getDataRange().getValues()[0].map(h => String(h).trim().toLowerCase());
    
    // Construct new user row array matching headers
    const newRow = userHeaders.map(header => {
      if (header === 'nrp') return item.nrp;
      if (header === 'nama') return item.nama;
      if (header === 'position' || header === 'jabatan') return item.position;
      if (header === 'job group' || header === 'jobgroup') return item.jobgroup;
      if (header === 'job rank' || header === 'jobrank') return item.jobrank;
      if (header === 'section') return item.section;
      if (header === 'sub section' || header === 'subsection') return item.subsection;
      if (header === 'custodian' || header === 'leader') return item.custodian;
      if (header === 'departemen' || header === 'department') return item.departemen;
      if (header === 'status') return item.status || 'ACTIVE';
      if (header === 'perusahaan' || header === 'company') return item.perusahaan;
      if (header === 'category' || header === 'kategori') return item.category;
      if (header === 'role') return item.role || 'user';
      return '';
    });

    userSheet.appendRow(newRow);

    // Update or Insert into Manpower Sheet
    if (mpSheet) {
      const mpValues = mpSheet.getDataRange().getValues();
      if (mpValues.length > 0) {
        const mpHeaders = mpValues[0].map(h => String(h).trim());
        const mpNrpIdx = mpHeaders.findIndex(h => h.toLowerCase() === 'nrp');

        const newMpRow = mpHeaders.map(header => {
          if (header.toLowerCase() === 'nrp') return item.nrp;
          if (header.toLowerCase() === 'nama') return item.nama;
          if (header.toLowerCase() === 'perusahaan') return item.perusahaan;
          if (header.toLowerCase() === 'status') return (item.status || 'ACTIVE').toUpperCase();
          
          if (dbUpdates[header] !== undefined) return dbUpdates[header];
          return '';
        });

        mpSheet.appendRow(newMpRow);
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Memperbarui data tunggal di sheet 'user' dan sheet 'Manpower'
 */
function updateSingleManpower(payload) {
  try {
    const item = payload.item;
    const dbUpdates = payload.dbUpdates || {};
    const normTargetNrp = normalizeNrp(item.nrp);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(SHEET_NAME);
    const mpSheet = ss.getSheetByName(DB_MP_SHEET_NAME);

    if (!userSheet) throw new Error(`Sheet '${SHEET_NAME}' tidak ditemukan.`);

    // 1. Update User Sheet
    const userValues = userSheet.getDataRange().getValues();
    const userHeaders = userValues[0].map(h => String(h).trim().toLowerCase());
    const nrpIdx = userHeaders.findIndex(h => h === 'nrp');

    let userTargetRow = -1;
    for (let i = 1; i < userValues.length; i++) {
      if (normalizeNrp(userValues[i][nrpIdx]) === normTargetNrp) {
        userTargetRow = i + 1;
        break;
      }
    }

    if (userTargetRow !== -1) {
      userHeaders.forEach((header, colIdx) => {
        if (header === 'nama') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.nama);
        else if (header === 'position' || header === 'jabatan') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.position);
        else if (header === 'job group' || header === 'jobgroup') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.jobgroup);
        else if (header === 'job rank' || header === 'jobrank') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.jobrank);
        else if (header === 'section') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.section);
        else if (header === 'sub section' || header === 'subsection') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.subsection);
        else if (header === 'custodian' || header === 'leader') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.custodian);
        else if (header === 'departemen' || header === 'department') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.departemen);
        else if (header === 'status') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.status);
        else if (header === 'perusahaan' || header === 'company') userSheet.getRange(userTargetRow, colIdx + 1).setValue(item.perusahaan);
      });
    }

    // 2. Update Manpower Sheet
    if (mpSheet) {
      const mpValues = mpSheet.getDataRange().getValues();
      if (mpValues.length > 0) {
        const mpHeaders = mpValues[0].map(h => String(h).trim());
        const mpNrpIdx = mpHeaders.findIndex(h => h.toLowerCase() === 'nrp');

        let mpTargetRow = -1;
        for (let i = 1; i < mpValues.length; i++) {
          if (normalizeNrp(mpValues[i][mpNrpIdx]) === normTargetNrp) {
            mpTargetRow = i + 1;
            break;
          }
        }

        // BINDING STATUS
        dbUpdates['Status'] = (item.status || 'ACTIVE').toUpperCase();
        dbUpdates['Section'] = (item.section || '').toUpperCase();
        dbUpdates['Sub Section / Section'] = (item.subsection || '').toUpperCase();
        dbUpdates['JABATAN'] = (item.jobgroup || '').toUpperCase();

        if (mpTargetRow !== -1) {
          mpHeaders.forEach((header, colIdx) => {
            if (dbUpdates[header] !== undefined) {
              mpSheet.getRange(mpTargetRow, colIdx + 1).setValue(dbUpdates[header]);
            }
          });
        } else {
          // Insert row if not exists in Manpower sheet
          const newMpRow = mpHeaders.map(header => {
            if (header.toLowerCase() === 'nrp') return item.nrp;
            if (header.toLowerCase() === 'nama') return item.nama;
            if (header.toLowerCase() === 'perusahaan') return item.perusahaan;
            return dbUpdates[header] !== undefined ? dbUpdates[header] : '';
          });
          mpSheet.appendRow(newMpRow);
        }
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Menghapus baris terpilih dari sheet 'user' dan sheet 'Manpower'
 */
function deleteManpowerRows(nrpList) {
  try {
    if (!nrpList || nrpList.length === 0) return { success: true, count: 0 };
    const normSet = new Set(nrpList.map(n => normalizeNrp(n)));

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(SHEET_NAME);
    const mpSheet = ss.getSheetByName(DB_MP_SHEET_NAME);

    let deletedCount = 0;

    // Delete from user sheet (bottom-up)
    if (userSheet) {
      const userValues = userSheet.getDataRange().getValues();
      const userHeaders = userValues[0].map(h => String(h).trim().toLowerCase());
      const nrpIdx = userHeaders.findIndex(h => h === 'nrp');

      for (let i = userValues.length - 1; i >= 1; i--) {
        if (normSet.has(normalizeNrp(userValues[i][nrpIdx]))) {
          userSheet.deleteRow(i + 1);
          deletedCount++;
        }
      }
    }

    // Delete from Manpower sheet (bottom-up)
    if (mpSheet) {
      const mpValues = mpSheet.getDataRange().getValues();
      if (mpValues.length > 0) {
        const mpHeaders = mpValues[0].map(h => String(h).trim());
        const mpNrpIdx = mpHeaders.findIndex(h => h.toLowerCase() === 'nrp');

        for (let i = mpValues.length - 1; i >= 1; i--) {
          if (normSet.has(normalizeNrp(mpValues[i][mpNrpIdx]))) {
            mpSheet.deleteRow(i + 1);
          }
        }
      }
    }

    return { success: true, count: deletedCount };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Memperbarui role pengguna secara masal di sheet 'user'
 */
function updateRolesBatch(nrpList, newRole) {
  try {
    if (!nrpList || nrpList.length === 0) return { success: true, count: 0 };
    const normSet = new Set(nrpList.map(n => normalizeNrp(n)));

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userSheet = ss.getSheetByName(SHEET_NAME);
    if (!userSheet) throw new Error(`Sheet '${SHEET_NAME}' tidak ditemukan.`);

    const userValues = userSheet.getDataRange().getValues();
    const userHeaders = userValues[0].map(h => String(h).trim().toLowerCase());
    const nrpIdx = userHeaders.findIndex(h => h === 'nrp');
    let roleIdx = userHeaders.findIndex(h => h === 'role');

    if (roleIdx === -1) {
      userSheet.getRange(1, userHeaders.length + 1).setValue('Role');
      roleIdx = userHeaders.length;
    }

    let updatedCount = 0;
    for (let i = 1; i < userValues.length; i++) {
      if (normSet.has(normalizeNrp(userValues[i][nrpIdx]))) {
        userSheet.getRange(i + 1, roleIdx + 1).setValue(newRole.toLowerCase());
        updatedCount++;
      }
    }

    return { success: true, count: updatedCount };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
