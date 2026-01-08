function files_refreshIndexFromDrive() {
  auth_requireEditor_();
  return files_refreshIndexFromDrive_();
}

function files_refreshIndexFromDrive_() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const cfg = cfg_getAll_();
    const formsFolderId = cfg_get_('FORMS_FOLDER_ID');
    const releasedFolderId = cfg_get_('RELEASED_FOLDER_ID');

    const obsoleteName = 'Obsolete';
    const formsObsoleteId = drive_getOrCreateSubfolderId_(formsFolderId, obsoleteName);
    const releasedObsoleteId = drive_getOrCreateSubfolderId_(releasedFolderId, obsoleteName);

    const activeForms = drive_listFilesInFolder_(formsFolderId);
    const obsoleteForms = drive_listFilesInFolder_(formsObsoleteId);
    const activeReleased = drive_listFilesInFolder_(releasedFolderId);
    const obsoleteReleased = drive_listFilesInFolder_(releasedObsoleteId);

    const prefix = (cfg.NAME_PREFIX || '').trim();

    let inserted = 0;
    let updated = 0;

    const processList = (list, isObsolete) => {
      list.forEach(f => {
        if (String(f.mimeType || '') !== MimeType.GOOGLE_SHEETS) return;
        const parsed = files_parseMbomName_(f.name, prefix);
        if (!parsed) return;

        const rec = {
          type: parsed.type,
          projectKey: parsed.projectKey,
          mbomRev: parsed.rev,
          baseFormRev: '',
          agileTabMDA: '',
          agileTabCluster: '',
          agileRevCluster: '',
          eco: '',
          description: '',
          fileId: f.id,
          url: f.url,
          createdAt: f.lastUpdated ? new Date(f.lastUpdated).toISOString() : '',
          createdBy: '',
          status: isObsolete ? 'OBSOLETE' : (parsed.type === 'RELEASED' ? 'RELEASED' : ''),
          notes: isObsolete ? 'Indexed from Obsolete folder' : 'Indexed from active folder'
        };

        const res = files_upsertByFileId_(rec);
        if (res.inserted) inserted++;
        else updated++;
      });
    };

    processList(activeForms, false);
    processList(obsoleteForms, true);
    processList(activeReleased, false);
    processList(obsoleteReleased, true);

    log_info_('Files index refreshed from Drive', { inserted, updated });
    return { ok: true, inserted, updated };
  } finally {
    lock.releaseLock();
  }
}

function files_parseMbomName_(name, prefix) {
  const n = String(name || '').trim();
  if (!n) return null;

  // Rev extraction
  const mRev = n.match(/-+\s*rev\s*(\d+)\s*$/i) || n.match(/\srev\s*(\d+)\s*$/i) || n.match(/rev\s*(\d+)\s*$/i);
  if (!mRev) return null;
  const rev = Number(mRev[1]);
  if (!isFinite(rev)) return null;

  const lower = n.toLowerCase();
  if (lower.includes(' - form - ') || lower.includes(' form - ')) {
    return { type: 'FORM', projectKey: 'GLOBAL', rev };
  }

  // RELEASED naming: "<prefix> - <projectKey> - Rev <n>"
  const idx = lower.lastIndexOf(' - rev');
  const left = idx > 0 ? n.substring(0, idx).trim() : n;

  // ProjectKey is after the last " - "
  const lastDash = left.lastIndexOf(' - ');
  const projectKey = lastDash >= 0 ? left.substring(lastDash + 3).trim() : '';

  if (!projectKey) return null;

  // If prefix provided, optionally ensure it matches (non-blocking)
  if (prefix && !n.startsWith(prefix)) {
    // still accept, to support legacy names
  }

  return { type: 'RELEASED', projectKey, rev };
}
