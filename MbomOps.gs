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
    const ss = SpreadsheetApp.openById(fileId);
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

    // Obsolete previous RELEASED before creating new one (move to Obsolete folder)
    mbom_obsoletePreviousReleased_({ projectKey, releasedFolderId, releasedObsoleteId });

    const releaseRev = Number(params.releaseRev || files_nextRev_('RELEASED', projectKey));
    const prefix = cfg_get_('NAME_PREFIX');
    const newName = `${prefix} - ${projectKey} - Rev ${releaseRev}`;

    const fileId = drive_copyFileWithRetry_(approvedFormId, releasedFolderId, newName);
    const url = `https://docs.google.com/spreadsheets/d/${fileId}`;

    const ss = SpreadsheetApp.openById(fileId);

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

    // Busway codes (confirmed by user)
    mbom_setBuswayCodes_(ss, {
      mdaCode: includeMda ? String(params.buswayMdaCode || '') : '',
      clusterCode: String(params.buswayClusterCode || '')
    });

    // Revision tab update (RELEASED keeps ECO label)
    mbom_upsertRevisionRow_(ss, {
      revision: `Rev ${releaseRev}`,
      createdAt: new Date(),
      createdBy: user,
      changeRefLabel: 'ECO',
      changeRefValue: String(params.eco || '').trim(),
      description: params.description || '',
      affectedItems: params.affectedItems || '',
      projectKey,
      agileClusterRev: cluster.Rev || '',
      baseFormRev: params.baseFormRev || ''
    });

    files_append_({
      type: 'RELEASED',
      projectKey,
      mbomRev: releaseRev,
      baseFormRev: params.baseFormRev || '',
      agileTabMDA: includeMda ? String(mda.TabName) : '',
      agileTabCluster: String(cluster.TabName),
      agileRevCluster: cluster.Rev || '',
      eco: String(params.eco || '').trim(),
      description: params.description || '',
      fileId,
      url,
      createdAt: new Date().toISOString(),
      createdBy: user,
      status: 'RELEASED',
      notes: freeze ? 'Agile inputs frozen as values' : 'Agile inputs linked'
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
        createdAt: new Date().toISOString()
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

  try {
    drive_moveFileToFolder_(prev.FileId, fromFolderId, toFolderId);
    files_setStatus_(prev.FileId, 'OBSOLETE');
    log_info_('Obsoleted previous RELEASED mBOM', { projectKey, fileId: prev.FileId, mbomRev: prev.MbomRev });
    return { ok: true, moved: true };
  } catch (e) {
    log_warn_('Failed to obsolete previous RELEASED mBOM', { projectKey, fileId: prev.FileId, error: e.message });
    return { ok: false, error: e.message };
  }
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
