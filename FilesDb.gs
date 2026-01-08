const FILES_SHEET = 'FILES';

function files_ensure_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FILES_SHEET);
  if (!sh) sh = ss.insertSheet(FILES_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'Type', 'ProjectKey', 'MbomRev', 'BaseFormRev', 'AgileTabMDA', 'AgileTabCluster',
      'AgileRevCluster', 'ECO', 'Description', 'FileId', 'Url',
      'CreatedAt', 'CreatedBy', 'Status', 'Notes'
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function files_list_(type) {
  const rows = files_listAll_();
  if (!type) return rows;
  return rows.filter(r => r.Type === type);
}

function files_listAll_() {
  if (globalThis.__FILES_CACHE__ && Array.isArray(globalThis.__FILES_CACHE__)) {
    return globalThis.__FILES_CACHE__;
  }
  const sh = files_ensure_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx]);
    out.push(obj);
  }
  globalThis.__FILES_CACHE__ = out;
  return out;
}

function files_resetCache_() {
  globalThis.__FILES_CACHE__ = null;
}

function files_getByFileId_(fileId) {
  const id = String(fileId || '').trim();
  if (!id) return null;
  const rows = files_list_();
  return rows.find(r => String(r.FileId || '').trim() === id) || null;
}

function files_append_(rec) {
  const sh = files_ensure_();
  sh.appendRow([
    rec.type || '',
    rec.projectKey || '',
    rec.mbomRev || '',
    rec.baseFormRev || '',
    rec.agileTabMDA || '',
    rec.agileTabCluster || '',
    rec.agileRevCluster || '',
    rec.eco || '',
    rec.description || '',
    rec.fileId || '',
    rec.url || '',
    rec.createdAt ? new Date(rec.createdAt) : new Date(),
    rec.createdBy || '',
    rec.status || '',
    rec.notes || ''
  ]);
  files_resetCache_();
}

/**
 * Upsert by FileId. If exists, overwrite row with provided fields (keeps blanks if you pass blanks).
 */
function files_upsertByFileId_(rec) {
  const sh = files_ensure_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idxFileId = headers.indexOf('FileId');
  if (idxFileId < 0) throw new Error('FILES: missing FileId header');

  const id = String(rec.fileId || '').trim();
  if (!id) throw new Error('Upsert requires fileId');

  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idxFileId] || '').trim() === id) {
      rowIndex = i + 1;
      break;
    }
  }

  const row = [
    rec.type || '',
    rec.projectKey || '',
    rec.mbomRev || '',
    rec.baseFormRev || '',
    rec.agileTabMDA || '',
    rec.agileTabCluster || '',
    rec.agileRevCluster || '',
    rec.eco || '',
    rec.description || '',
    rec.fileId || '',
    rec.url || '',
    rec.createdAt ? new Date(rec.createdAt) : '',
    rec.createdBy || '',
    rec.status || '',
    rec.notes || ''
  ];

  if (rowIndex === -1) {
    sh.appendRow(row);
    files_resetCache_();
    return { ok: true, inserted: true };
  } else {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    files_resetCache_();
    return { ok: true, inserted: false };
  }
}

function files_getLatestBy_(type, predicateFn) {
  const rows = files_list_(type);
  const filtered = rows.filter(predicateFn);
  filtered.sort((a, b) => Number(b.MbomRev || 0) - Number(a.MbomRev || 0));
  return filtered[0] || null;
}

function files_nextRev_(type, projectKey) {
  const latest = files_getLatestBy_(type, r => (r.ProjectKey || '') === projectKey);
  const n = Number(latest?.MbomRev || 0);
  return n + 1;
}

function files_setStatus_(fileId, status) {
  const sh = files_ensure_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idxFileId = headers.indexOf('FileId');
  const idxStatus = headers.indexOf('Status');
  if (idxFileId < 0 || idxStatus < 0) throw new Error('FILES headers missing required columns');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idxFileId]).trim() === String(fileId).trim()) {
      sh.getRange(i + 1, idxStatus + 1).setValue(String(status || '').trim());
      files_resetCache_();
      return true;
    }
  }
  return false;
}
