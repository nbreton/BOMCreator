//MbomOps.gs
function mbom_withCopyLock_(label, fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(120000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function mbom_createNewFormRevision_(params) {
  return mbom_withCopyLock_('CREATE_FORM', () => {
    const user = auth_requireEditor_();
    const cfg = cfg_getAll_();

    const baseFormId = params.baseFormFileId || cfg_get_('CURRENT_APPROVED_FORM_FILE_ID');
    const destFolderId = cfg_get_('FORMS_FOLDER_ID');
    const obsoleteFolderId = drive_getOrCreateSubfolderId_(destFolderId, 'Obsolete');

    const newFormRev = Number(params.newFormRev);
    if (!isFinite(newFormRev) || newFormRev <= 0) throw new Error('Invalid newFormRev');

    const changeRef = String(params.changeRef || params.ecrActRef || params.eco || '').trim(); // backward compatible

    const prefix = cfg_get_('NAME_PREFIX');
    const newName = `${prefix} - Form - Rev ${newFormRev}`;

    // Create new form (copy)
    const fileId = drive_copyFileWithRetry_(baseFormId, destFolderId, newName);
    const url = `https://docs.google.com/spreadsheets/d/${fileId}`;

    // Mark previous (base) form as obsolete by moving it
    mbom_obsoletePreviousForm_({ baseFormId, formsFolderId: destFolderId, formsObsoleteId: obsoleteFolderId });

    // Update Revision tab inside new file
    const ss = mbom_openSpreadsheetWithRetry_(fileId, { label: 'CREATE_FORM' });
    mbom_upsertRevisionRow_(ss, {
      revision: `Rev ${newFormRev}`,
      createdAt: new Date(),
      createdBy: user,
      changeRefLabel: 'ECR/ACT Ref',
      changeRefValue: changeRef,
      description: params.description || '',
      affectedItems: params.affectedItems || '',
      projectKey: 'GLOBAL',
      agileClusterRev: '',
      baseFormRev: params.baseFormRev || ''
    });

    files_append_({
      type: 'FORM',
      projectKey: 'GLOBAL',
      mbomRev: newFormRev,
      baseFormRev: params.baseFormRev || '',
      eco: changeRef, // stored in ECO column for backward compatibility
      description: params.description || '',
      fileId,
      url,
      fileName: newName,
      createdAt: new Date().toISOString(),
      createdBy: user,
      status: 'DRAFT',
      notes: params.notes || ''
    });

    // Email notification: new Form revision created
    try {
      notif_sendFormCreated_({
        mbomRev: newFormRev,
        fileId,
        url,
        createdBy: user,
        changeRef,
        description: params.description || ''
      });
    } catch (e) {
      log_warn_('Form created notification failed', { error: e.message });
    }

    log_info_('Created new Form revision', { fileId, newFormRev, baseFormId });
    return { ok: true, fileId, url, name: newName };
  });
}

function mbom_createReleasedForProject_(params) {
  return mbom_withCopyLock_('CREATE_RELEASED', () => {
    const user = auth_requireEditor_();
    const cfg = cfg_getAll_();

    const approvedFormId = params.approvedFormFileId || cfg_get_('CURRENT_APPROVED_FORM_FILE_ID');
    const releasedFolderId = cfg_get_('RELEASED_FOLDER_ID');
    const releasedObsoleteId = drive_getOrCreateSubfolderId_(releasedFolderId, 'Obsolete');

    const requireApprovedForm = cfg_bool_('REQUIRE_APPROVED_FORM', true);
    if (requireApprovedForm) {
      const formRec = files_getByFileId_(approvedFormId);
      if (!formRec || String(formRec.Type) !== 'FORM' || String(formRec.Status || '').toUpperCase() !== 'APPROVED') {
        throw new Error(`Controlled release blocked: selected Form file is not marked APPROVED in FILES.\nForm FileId: ${approvedFormId}`);
      }
    }

    const projectKey = String(params.projectKey || '').trim();
    if (!projectKey) throw new Error('Missing projectKey');

    const effective = projects_getEffective_(projectKey);
    const includeMda = (params.includeMda !== undefined) ? !!params.includeMda : !!effective.includeMda;

    const releaseRev = Number(params.releaseRev || files_nextRev_('RELEASED', projectKey));
    const prefix = cfg_get_('NAME_PREFIX');
    const newName = `${prefix} - ${projectKey} - Rev ${releaseRev}`;

    const fileId = drive_copyFileWithRetry_(approvedFormId, releasedFolderId, newName);
    const url = `https://docs.google.com/spreadsheets/d/${fileId}`;

    // Obsolete previous RELEASED after new copy is created (move to Obsolete folder)
    const obsoleteRes = mbom_obsoletePreviousReleased_({ projectKey, releasedFolderId, releasedObsoleteId });

    const ss = mbom_openSpreadsheetWithRetry_(fileId, { label: 'CREATE_RELEASED' });

    // Determine Agile tabs (latest by default)
    const site = projectKey.split('-')[0];
    const zoneNum = projectKey.split('-')[1];
    const clusterPart = `Zone ${zoneNum}`;

    const cluster = params.agileTabCluster
      ? { TabName: params.agileTabCluster, Rev: params.agileRevCluster || '' }
      : agile_getLatestTab_(site, clusterPart);

    if (!cluster || !cluster.TabName) throw new Error(`No Agile tab found for ${site} / ${clusterPart}`);

    let mda = null;
    if (includeMda) {
      mda = params.agileTabMDA
        ? { TabName: params.agileTabMDA, Rev: params.agileRevMDA || '' }
        : agile_getLatestTab_(site, 'MDA');

      if (!mda || !mda.TabName) throw new Error(`MDA required but no Agile tab found for ${site} / MDA`);
    }

    // Enforce Agile approvals if configured
    const requireApprovedAgile = cfg_bool_('REQUIRE_APPROVED_AGILE', true);
    if (requireApprovedAgile) {
      const stCl = agile_approval_status_(String(cluster.TabName));
      if (stCl !== 'APPROVED') throw new Error(`Controlled release blocked: Cluster Agile tab is not APPROVED: ${cluster.TabName} (${stCl}).`);

      if (includeMda) {
        const stMda = agile_approval_status_(String(mda.TabName));
        if (stMda !== 'APPROVED') throw new Error(`Controlled release blocked: MDA Agile tab is not APPROVED: ${mda.TabName} (${stMda}).`);
      }
    }

    // Apply configuration into the copied file (MDA optional)
    mbom_setAgileInputs_(ss, {
      downloadListId: cfg_get_('DOWNLOAD_LIST_SS_ID'),
      mdaTabName: includeMda ? String(mda.TabName) : '',
      clusterTabName: String(cluster.TabName)
    });

    const freeze = (params.freezeAgileInputs !== undefined)
      ? !!params.freezeAgileInputs
      : cfg_bool_('FREEZE_AGILE_INPUTS_DEFAULT', true);

    if (freeze) {
      mbom_freezeAgileInputs_(ss, cfg_get_('DOWNLOAD_LIST_SS_ID'),
        includeMda ? String(mda.TabName) : '',
        String(cluster.TabName)
      );
    }

    const clusterRow = agile_findTabByName_(String(cluster.TabName || ''));
    const mdaRow = includeMda ? agile_findTabByName_(String(mda.TabName || '')) : null;

    const clusterSupplier = String(effective.clusterBuswaySupplier || clusterRow?.BuswaySupplier || '');
    const mdaSupplier = includeMda ? String(effective.mdaBuswaySupplier || mdaRow?.BuswaySupplier || '') : '';

    const buswayClusterCode = String(params.buswayClusterCode || mbom_inferClusterCode_(clusterSupplier) || '');
    const buswayMdaCode = includeMda ? String(params.buswayMdaCode || mbom_inferMdaCode_(mdaSupplier) || '') : '';

    // Busway codes (read-only in UI, inferred from suppliers or overrides)
    mbom_setBuswayCodes_(ss, {
      mdaCode: includeMda ? buswayMdaCode : '',
      clusterCode: buswayClusterCode
    });

    const ecoValues = [];
    if (clusterRow && clusterRow.ECO) ecoValues.push(String(clusterRow.ECO || '').trim());
    if (includeMda && mdaRow && mdaRow.ECO) ecoValues.push(String(mdaRow.ECO || '').trim());
    const ecoFromAgile = Array.from(new Set(ecoValues.filter(Boolean))).join(' / ');
    const ecoFinal = String(params.eco || '').trim() || ecoFromAgile;

    // Revision tab update (RELEASED keeps ECO label)
    mbom_upsertRevisionRow_(ss, {
      revision: `Rev ${releaseRev}`,
      createdAt: new Date(),
      createdBy: user,
      changeRefLabel: 'ECO',
      changeRefValue: ecoFinal,
      description: params.description || '',
      affectedItems: params.affectedItems || '',
      projectKey,
      agileClusterRev: cluster.Rev || '',
      baseFormRev: params.baseFormRev || ''
    });

    mbom_updateReleasedRevisionSheet_(ss, {
      releaseRev,
      clusterRow,
      mdaRow,
      updateDate: new Date()
    });

    files_append_({
      type: 'RELEASED',
      projectKey,
      mbomRev: releaseRev,
      baseFormRev: params.baseFormRev || '',
      agileTabMDA: includeMda ? String(mda.TabName) : '',
      agileTabCluster: String(cluster.TabName),
      agileRevCluster: cluster.Rev || '',
      eco: ecoFinal,
      description: params.description || '',
      fileId,
      url,
      fileName: newName,
      createdAt: new Date().toISOString(),
      createdBy: user,
      status: 'RELEASED',
      notes: freeze ? 'Agile inputs copied as values' : 'Agile inputs copied'
    });

    // Email notification: RELEASED enqueue (digest grouped)
    try {
      notif_enqueueReleasedEvent_({
        projectKey,
        mbomRev: releaseRev,
        url,
        fileId,
        site,
        clusterTab: String(cluster.TabName),
        mdaTab: includeMda ? String(mda.TabName) : '',
        createdBy: user,
        createdAt: new Date().toISOString(),
        obsolete: obsoleteRes?.obsolete || null
      });
    } catch (e) {
      log_warn_('Released digest enqueue failed', { error: e.message });
    }

    log_info_('Created RELEASED mBOM', { projectKey, fileId, releaseRev, includeMda });
    return { ok: true, fileId, url, name: newName };
  });
}

function mbom_obsoletePreviousForm_(params) {
  const baseFormId = String(params.baseFormId || '').trim();
  if (!baseFormId) return { ok: true, skipped: 'missing baseFormId' };

  const fromFolderId = String(params.formsFolderId || '').trim();
  const toFolderId = String(params.formsObsoleteId || '').trim();
  if (!fromFolderId || !toFolderId) throw new Error('Missing Forms folder IDs for obsoleting');

  try {
    drive_moveFileToFolder_(baseFormId, fromFolderId, toFolderId);
    mbom_prefixObsoleteFileName_(baseFormId);
    files_setStatus_(baseFormId, 'OBSOLETE');
    log_info_('Obsoleted previous Form revision', { fileId: baseFormId });
    return { ok: true, moved: true };
  } catch (e) {
    log_warn_('Failed to obsolete previous Form revision', { fileId: baseFormId, error: e.message });
    return { ok: false, error: e.message };
  }
}

function mbom_obsoletePreviousReleased_(params) {
  const projectKey = String(params.projectKey || '').trim();
  if (!projectKey) return { ok: true, skipped: 'missing projectKey' };

  const fromFolderId = String(params.releasedFolderId || '').trim();
  const toFolderId = String(params.releasedObsoleteId || '').trim();
  if (!fromFolderId || !toFolderId) throw new Error('Missing Released folder IDs for obsoleting');

  const prev = files_getLatestBy_('RELEASED', r =>
    (r.ProjectKey || '') === projectKey &&
    String(r.Status || '').toUpperCase() !== 'OBSOLETE'
  );
  if (!prev || !prev.FileId) return { ok: true, skipped: 'no previous release' };

  const obsoleteInfo = {
    projectKey: String(prev.ProjectKey || projectKey),
    mbomRev: prev.MbomRev,
    fileId: prev.FileId,
    url: prev.Url || '',
    fileName: prev.FileName || ''
  };

  try {
    drive_moveFileToFolder_(prev.FileId, fromFolderId, toFolderId);
    const newName = mbom_prefixObsoleteFileName_(prev.FileId);
    if (newName) obsoleteInfo.fileName = newName;
    files_setStatus_(prev.FileId, 'OBSOLETE');
    log_info_('Obsoleted previous RELEASED mBOM', { projectKey, fileId: prev.FileId, mbomRev: prev.MbomRev });
    return { ok: true, moved: true, obsolete: obsoleteInfo };
  } catch (e) {
    log_warn_('Failed to obsolete previous RELEASED mBOM', { projectKey, fileId: prev.FileId, error: e.message });
    return { ok: false, error: e.message, obsolete: obsoleteInfo };
  }
}

function mbom_obsoleteFormFile_(params) {
  const fileId = String(params.fileId || '').trim();
  if (!fileId) throw new Error('Missing form fileId');

  const formsFolderId = cfg_get_('FORMS_FOLDER_ID');
  const obsoleteFolderId = drive_getOrCreateSubfolderId_(formsFolderId, 'Obsolete');

  try {
    drive_moveFileToFolder_(fileId, formsFolderId, obsoleteFolderId);
    mbom_prefixObsoleteFileName_(fileId);
    files_setStatus_(fileId, 'OBSOLETE');
    log_info_('Obsoleted Form file', { fileId });
    return { ok: true, moved: true };
  } catch (e) {
    log_warn_('Failed to obsolete Form file', { fileId, error: e.message });
    return { ok: false, error: e.message };
  }
}

function mbom_setAgileInputs_(ss, params) {
  const sourceIdRaw = String(params.downloadListId || '').trim();
  const sourceId = sourceIdRaw || '1q9Y2NgS4SAJGZ2OMtQafjTCCFWgQFldlmduQnW-qK7I';
  const copyRange = 'A1:U401';
  const destRange = 'A4:U404';

  const copyInputs = (sheetName, tabName, overrideSourceId) => {
    const tab = String(tabName || '').trim();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return false;
    sh.getRange('B1:B3').clearContent();
    if (!tab) {
      sh.getRange(destRange).clearContent();
      sh.getRange('A2').clearContent();
      return false;
    }
    const srcId = String(overrideSourceId || sourceId || '').trim();
    if (!srcId) return false;
    const src = SpreadsheetApp.openById(srcId);
    const srcSheet = src.getSheetByName(tab);
    if (!srcSheet) {
      log_warn_('Agile source tab missing', { sheet: sheetName, tab, sourceId: srcId });
      return false;
    }
    const values = srcSheet.getRange(copyRange).getValues();
    sh.getRange(destRange).setValues(values);
    const url = `https://docs.google.com/spreadsheets/d/${srcId}/edit#gid=${srcSheet.getSheetId()}`;
    sh.getRange('A2').setFormula(`=HYPERLINK("${url}","Open ${tab}")`);
    return true;
  };

  copyInputs('INPUT_BOM_AGILE_Cluster', params.clusterTabName);
  copyInputs('INPUT_BOM_AGILE_MDA', params.mdaTabName);
  copyInputs('INPUT_BOM_AGILE_WP5 KOPs', 'BOM Lvl0/Lvl1 WP5', '1q9Y2NgS4SAJGZ2OMtQafjTCCFWgQFldlmduQnW-qK7I');

  SpreadsheetApp.flush();
}

function mbom_updateReleasedRevisionSheet_(ss, params) {
  const sh = ss.getSheetByName('Revision');
  if (!sh) return;

  const releaseRev = params.releaseRev;
  if (releaseRev !== undefined && releaseRev !== null && releaseRev !== '') {
    sh.getRange('B1').setValue(releaseRev);
  }

  const headerInfo = mbom_findRevisionTableHeader_(sh);
  if (!headerInfo) {
    log_warn_('Revision table header not found', { sheet: 'Revision' });
    return;
  }

  const { headerRow, colMap } = headerInfo;
  const startRow = headerRow + 1;
  const today = Utilities.formatDate(params.updateDate || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const formatDate = (value) => (
    value instanceof Date
      ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(value || '').trim()
  );

  const entries = [
    { tabLabel: 'INPUT_BOM_AGILE_MDA', row: params.mdaRow || null },
    { tabLabel: 'INPUT_BOM_AGILE_Cluster', row: params.clusterRow || null }
  ];

  const maxCol = Math.max(1, ...Object.values(colMap).filter(v => v > 0));

  entries.forEach((entry, idx) => {
    const r = entry.row || {};
    const rowValues = new Array(maxCol).fill('');
    const setIf = (col, value) => { if (col > 0) rowValues[col - 1] = value; };

    setIf(colMap.site, String(r.Site || '').trim());
    setIf(colMap.part, String(r.Part || r.PartNorm || '').trim());
    setIf(colMap.tlaRef, String(r.TlaRef || '').trim());
    setIf(colMap.description, String(r.Description || '').trim());
    setIf(colMap.rev, r.Rev !== undefined ? r.Rev : '');
    setIf(colMap.downloadDate, formatDate(r.DownloadDate || ''));
    setIf(colMap.mbomDate, today);
    setIf(colMap.tabName, entry.tabLabel);

    sh.getRange(startRow + idx, 1, 1, maxCol).setValues([rowValues]);
  });

  SpreadsheetApp.flush();
}

function mbom_findRevisionTableHeader_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const lastRow = Math.max(1, sh.getLastRow());
  const scanRows = Math.min(20, lastRow);
  const values = sh.getRange(1, 1, scanRows, lastCol).getValues();
  const norm = (val) => String(val || '')
    .replace(/["']/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  const findIndex = (row, candidates) => {
    for (const c of candidates) {
      const cn = norm(c);
      const idx = row.findIndex(cell => cell === cn || cell.includes(cn));
      if (idx >= 0) return idx + 1;
    }
    return -1;
  };

  for (let i = 0; i < values.length; i++) {
    const row = values[i].map(norm);
    const hasSite = row.includes('site');
    const hasPart = row.some(cell => cell.startsWith('part'));
    const hasTab = row.some(cell => cell.includes('name of tab') || cell === 'tab');
    if (!hasSite || !hasPart || !hasTab) continue;

    const colMap = {
      site: findIndex(row, ['site']),
      part: findIndex(row, ['part']),
      tlaRef: findIndex(row, ['tla ref', 'tla']),
      description: findIndex(row, ['description', 'desc']),
      rev: findIndex(row, ['rev']),
      downloadDate: findIndex(row, ['date of downloading', 'date']),
      mbomDate: findIndex(row, ['date of mbom update', 'mbom update']),
      tabName: findIndex(row, ['name of tab', 'tab'])
    };

    return { headerRow: i + 1, colMap };
  }

  return null;
}

function mbom_openSpreadsheetWithRetry_(fileId, opts) {
  const id = String(fileId || '').trim();
  if (!id) throw new Error('Missing spreadsheet fileId');
  const label = String(opts?.label || '').trim();
  const attempts = Math.max(1, Number(opts?.attempts || 4));
  const baseSleepMs = Math.max(500, Number(opts?.baseSleepMs || 1500));
  let lastErr = null;
  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      lastErr = e;
      const sleepMs = Math.min(20000, baseSleepMs * Math.pow(2, attempt - 1));
      log_warn_('Spreadsheet open retry', {
        fileId: id,
        label,
        attempt,
        attempts,
        error: e.message
      });
      Utilities.sleep(sleepMs);
    }
  }
  throw lastErr || new Error('Failed to open spreadsheet');
}

function mbom_authorizeImportrange_(sheet, sourceId, tabName, rangeA1) {
  if (!sheet || !sourceId || !tabName || !rangeA1) return false;
  try {
    SpreadsheetApp.openById(sourceId);
    const tmp = sheet.getRange('Z1');
    const safeSource = String(sourceId || '').replace(/"/g, '""');
    const safeTab = String(tabName || '').replace(/"/g, '""');
    tmp.setFormula(`=IMPORTRANGE("${safeSource}","${safeTab}!${rangeA1}")`);
    SpreadsheetApp.flush();
    tmp.setValue('');
    return true;
  } catch (e) {
    log_warn_('IMPORTRANGE authorization failed', { sheet: sheet.getName(), error: e.message });
    return false;
  }
}

function mbom_freezeAgileInputs_(ss, downloadListId, mdaTabName, clusterTabName) {
  const rangeA1 = 'A4:U404';
  const freezeSheet = (sheetName, tabName) => {
    const tab = String(tabName || '').trim();
    if (!tab) return false;
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return false;
    const rng = sh.getRange(rangeA1);
    rng.copyTo(rng, { contentsOnly: true });
    return true;
  };

  freezeSheet('INPUT_BOM_AGILE_Cluster', clusterTabName);
  freezeSheet('INPUT_BOM_AGILE_MDA', mdaTabName);
  SpreadsheetApp.flush();
}


function mbom_setBuswayCodes_(ss, codes) {
  const clusterCode = String(codes.clusterCode || '').trim();
  const mdaCode = String(codes.mdaCode || '').trim();

  const setG2 = (sheetName, value) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return false;
    sh.getRange('G2').setValue(value);
    return true;
  };

  if (!setG2('Cluster', clusterCode)) {
    setG2('INPUT_BOM_AGILE_Cluster', clusterCode);
  }

  if (!setG2('MDA', mdaCode)) {
    setG2('INPUT_BOM_AGILE_MDA', mdaCode);
  }

  SpreadsheetApp.flush();
}

function mbom_inferMdaCode_(buswaySupplier) {
  const s = String(buswaySupplier || '').toUpperCase();
  if (s.includes('STARLINE')) return 'ST';
  if (s.includes('EI') || s.includes('E&I')) return 'EI';
  return '';
}

function mbom_inferClusterCode_(buswaySupplier) {
  const s = String(buswaySupplier || '').toUpperCase();
  if (s.includes('MARDIX')) return 'MA';
  if (s.includes('EAE')) return 'EA';
  if (s.includes('EI') || s.includes('E&I')) return 'EI';
  return '';
}

function mbom_prefixObsoleteFileName_(fileId) {
  const id = String(fileId || '').trim();
  if (!id) return '';
  const prefix = '[OBSELETE] ';
  try {
    const file = DriveApp.getFileById(id);
    const current = String(file.getName() || '').trim();
    if (current.startsWith(prefix)) {
      files_setFileName_(id, current);
      return current;
    }
    const newName = `${prefix}${current}`;
    file.setName(newName);
    files_setFileName_(id, newName);
    return newName;
  } catch (e) {
    log_warn_('Failed to rename obsolete file', { fileId: id, error: e.message });
    return '';
  }
}

/**
 * Revision upsert with flexible "Change Reference" column.
 * - For Forms: label "ECR/ACT Ref"
 * - For RELEASED: label "ECO"
 */
function mbom_upsertRevisionRow_(ss, meta) {
  const sh = ss.getSheetByName('Revision');
  if (!sh) return;

  // Ensure header exists
  const maxCols = Math.max(10, sh.getLastColumn());
  let header = sh.getRange(1, 1, 1, maxCols).getValues()[0].map(x => String(x || '').trim());

  const isHeaderEmpty = header.every(h => !h);
  if (isHeaderEmpty) {
    header = [
      'Revision',
      'Date',
      'Creator',
      (meta.changeRefLabel || 'ECO'),
      'Description',
      'Affected Items',
      'Project',
      'Agile Cluster Rev',
      'Base Form Rev',
      'Notes'
    ];
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }

  // Map columns by header name (robust)
  const idx = (names) => {
    const lower = header.map(h => String(h || '').toLowerCase());
    for (const n of names) {
      const k = String(n).toLowerCase();
      const i = lower.findIndex(h => h === k || h.includes(k));
      if (i >= 0) return i + 1;
    }
    return -1;
  };

  const cRevision = idx(['revision']);
  const cDate = idx(['date']);
  const cCreator = idx(['creator', 'created by']);
  const cChangeRef = idx([meta.changeRefLabel || '', 'ecr/act', 'eco', 'change ref']);
  const cDesc = idx(['description']);
  const cAffected = idx(['affected']);
  const cProject = idx(['project']);
  const cAgile = idx(['agile cluster']);
  const cBase = idx(['base form']);
  const cNotes = idx(['notes']);

  sh.insertRowBefore(2);

  const setIf = (col, val) => { if (col > 0) sh.getRange(2, col).setValue(val); };

  setIf(cRevision, meta.revision || '');
  setIf(cDate, meta.createdAt || new Date());
  setIf(cCreator, meta.createdBy || '');
  setIf(cChangeRef, meta.changeRefValue || '');
  setIf(cDesc, meta.description || '');
  setIf(cAffected, meta.affectedItems || '');
  setIf(cProject, meta.projectKey || '');
  setIf(cAgile, meta.agileClusterRev || '');
  setIf(cBase, meta.baseFormRev || '');
  setIf(cNotes, '');

  SpreadsheetApp.flush();
}
