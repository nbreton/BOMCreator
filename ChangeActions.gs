const CHANGE_ACTIONS_DEFAULT_SPREADSHEET_ID = '1CUhpt4w7FjDj4PtGJm3sOHWVm_rfCYIZ1sLAS-XUzJ8';
const CHANGE_ACTIONS_DEFAULT_SHEETS = [
  'EoR CHANGE TRACKER',
  'MAIN ACTION TRACKER',
  'LESSON LEARNED TRACKER',
  'BOM REVIEW',
  'DWG REVIEW'
];
const CHANGE_ACTIONS_HEADER_ROW = 10;
const CHANGE_ACTIONS_MAX_ROWS = 2000;

function change_actions_config_() {
  const cfg = cfg_getAll_();
  const spreadsheetId = String(cfg.ECR_ACT_SOURCE_SPREADSHEET_ID || CHANGE_ACTIONS_DEFAULT_SPREADSHEET_ID || '').trim();
  const sheetsRaw = String(cfg.ECR_ACT_SOURCE_SHEETS || CHANGE_ACTIONS_DEFAULT_SHEETS.join(',')).trim();
  const sheetNames = sheetsRaw
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
  const uniqueSheetNames = Array.from(new Set(sheetNames));
  return { spreadsheetId, sheetNames: uniqueSheetNames };
}

function change_actions_normHeader_(s) {
  return String(s || '')
    .replace(/["']/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function change_actions_findHeader_(headerNorm, candidates) {
  for (const c of candidates) {
    const cn = change_actions_normHeader_(c);
    const idx = headerNorm.findIndex(h => h === cn || h.includes(cn));
    if (idx >= 0) return idx;
  }
  return -1;
}

function change_actions_list_() {
  const cfg = change_actions_config_();
  if (!cfg.spreadsheetId) {
    return { actions: [], sources: [], warning: 'Missing ECR_ACT_SOURCE_SPREADSHEET_ID configuration.' };
  }
  if (!cfg.sheetNames.length) {
    return { actions: [], sources: [], warning: 'Missing ECR_ACT_SOURCE_SHEETS configuration.' };
  }

  const actions = [];
  const warnings = [];

  const src = SpreadsheetApp.openById(cfg.spreadsheetId);
  cfg.sheetNames.forEach(sheetName => {
    const sh = src.getSheetByName(sheetName);
    if (!sh) {
      warnings.push(`Sheet not found: ${sheetName}`);
      return;
    }

    const lastCol = Math.max(1, sh.getLastColumn());
    const header = sh.getRange(CHANGE_ACTIONS_HEADER_ROW, 1, 1, lastCol).getValues()[0];
    const headerNorm = header.map(change_actions_normHeader_);

    const col = {
      actionNb: change_actions_findHeader_(headerNorm, ['action nb', 'action number']),
      projectName: change_actions_findHeader_(headerNorm, ['project name']),
      workPackage: change_actions_findHeader_(headerNorm, ['work package']),
      moduleType: change_actions_findHeader_(headerNorm, ['module type']),
      actionTitle: change_actions_findHeader_(headerNorm, ['action title']),
      referenceFile: change_actions_findHeader_(headerNorm, ['reference file', 'reference file ( ver id', 'reference file (ver id']),
      discipline: change_actions_findHeader_(headerNorm, ['discipline']),
      creationDate: change_actions_findHeader_(headerNorm, ['creation date']),
      priority: change_actions_findHeader_(headerNorm, ['priority']),
      raisedBy: change_actions_findHeader_(headerNorm, ['raised by']),
      assignedTo: change_actions_findHeader_(headerNorm, ['assigned to']),
      lifecycleStatus: change_actions_findHeader_(headerNorm, ['lifecycle status']),
      progress: change_actions_findHeader_(headerNorm, ['progress']),
      estimatedTime: change_actions_findHeader_(headerNorm, ['estimated time to complete']),
      comments: change_actions_findHeader_(headerNorm, ['comments']),
      xvtAction: change_actions_findHeader_(headerNorm, ['xvt action']),
      picture: change_actions_findHeader_(headerNorm, ['picture / image', 'picture/image', 'picture'])
    };

    if (col.actionNb < 0 || col.actionTitle < 0) {
      warnings.push(`Missing required headers in ${sheetName}. Expected "ACTION NB" and "ACTION Title" on row ${CHANGE_ACTIONS_HEADER_ROW}.`);
      return;
    }

    const lastRow = sh.getLastRow();
    const numRows = Math.max(0, lastRow - CHANGE_ACTIONS_HEADER_ROW);
    const limit = Math.min(numRows, CHANGE_ACTIONS_MAX_ROWS);
    if (limit <= 0) return;

    const data = sh.getRange(CHANGE_ACTIONS_HEADER_ROW + 1, 1, limit, lastCol).getValues();

    const sheetId = sh.getSheetId();
    data.forEach((row, idx) => {
      const actionNb = String(row[col.actionNb] || '').trim();
      const actionTitle = String(row[col.actionTitle] || '').trim();
      if (!actionNb && !actionTitle) return;

      actions.push({
        actionNb,
        projectName: String(row[col.projectName] || '').trim(),
        workPackage: String(row[col.workPackage] || '').trim(),
        moduleType: String(row[col.moduleType] || '').trim(),
        actionTitle,
        referenceFile: String(row[col.referenceFile] || '').trim(),
        discipline: String(row[col.discipline] || '').trim(),
        creationDate: row[col.creationDate] instanceof Date
          ? Utilities.formatDate(row[col.creationDate], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(row[col.creationDate] || '').trim(),
        priority: String(row[col.priority] || '').trim(),
        raisedBy: String(row[col.raisedBy] || '').trim(),
        assignedTo: String(row[col.assignedTo] || '').trim(),
        lifecycleStatus: String(row[col.lifecycleStatus] || '').trim(),
        progress: String(row[col.progress] || '').trim(),
        estimatedTime: String(row[col.estimatedTime] || '').trim(),
        comments: String(row[col.comments] || '').trim(),
        xvtAction: String(row[col.xvtAction] || '').trim(),
        picture: String(row[col.picture] || '').trim(),
        sourceSheet: sheetName,
        sourceRow: CHANGE_ACTIONS_HEADER_ROW + 1 + idx,
        sourceSheetId: sheetId,
        sourceSpreadsheetId: cfg.spreadsheetId
      });
    });
  });

  return {
    actions,
    sources: cfg.sheetNames,
    warning: warnings.length ? warnings.join(' | ') : ''
  };
}
