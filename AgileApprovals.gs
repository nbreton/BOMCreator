const AGILE_APPROVALS_SHEET = 'AGILE_APPROVALS';

function agile_approval_ensure_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(AGILE_APPROVALS_SHEET);
  if (!sh) sh = ss.insertSheet(AGILE_APPROVALS_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['TabName', 'Status', 'UpdatedAt', 'UpdatedBy', 'Notes']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function agile_approval_getMap_() {
  const sh = agile_approval_ensure_();
  const values = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const tabName = String(values[i][0] || '').trim();
    if (!tabName) continue;
    map[tabName] = {
      status: String(values[i][1] || '').trim() || 'PENDING',
      updatedAt: values[i][2] || '',
      updatedBy: String(values[i][3] || '').trim(),
      notes: String(values[i][4] || '').trim()
    };
  }
  return map;
}

function agile_approval_status_(tabName) {
  const t = String(tabName || '').trim();
  if (!t) return 'PENDING';
  const map = agile_approval_getMap_();
  const entry = map[t];
  return entry ? String(entry.status || 'PENDING').toUpperCase() : 'PENDING';
}

function agile_approval_set_(tabName, status, notes) {
  auth_requireEditor_();

  const t = String(tabName || '').trim();
  if (!t) throw new Error('Missing tabName');

  const normalizedStatus = String(status || '').trim().toUpperCase();
  if (!['APPROVED', 'REJECTED'].includes(normalizedStatus)) {
    throw new Error('Status must be APPROVED or REJECTED');
  }

  const user = auth_getUser_();
  const sh = agile_approval_ensure_();
  const values = sh.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === t) {
      rowIndex = i + 1;
      break;
    }
  }

  const row = [t, normalizedStatus, new Date(), user.email || '', String(notes || '').trim()];
  if (rowIndex === -1) {
    sh.appendRow(row);
  } else {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  }

  return { ok: true, tabName: t, status: normalizedStatus };
}
