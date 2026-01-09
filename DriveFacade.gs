function drive_copyFileWithRetry_(srcFileId, destFolderId, newName, maxAttempts = 6) {
  const lock = LockService.getScriptLock();
  lock.waitLock(120000);
  try {
    const sizeBytes = drive_getFileSizeSafe_(srcFileId);
    const largeFile = sizeBytes >= 50 * 1024 * 1024;
    const attempts = Math.max(maxAttempts, largeFile ? 8 : maxAttempts);
    const openWaitMs = largeFile ? 4 * 60 * 1000 : 90 * 1000;
    const initialDelayMs = largeFile ? 6000 : 1000;
    let lastErr = null;
    for (let attempt = 1; attempt <= attempts; attempt++) {
      try {
        const fileId = drive_copyFile_(srcFileId, destFolderId, newName);
        Utilities.sleep(initialDelayMs);
        drive_openSpreadsheetWithRetry_(fileId, openWaitMs);
        drive_finalizeCopiedFile_(fileId, destFolderId, newName);
        return fileId;
      } catch (e) {
        lastErr = e;
        const sleepMs = Math.min(20000, 1200 * Math.pow(2, attempt));
        Utilities.sleep(sleepMs);
      }
    }
    throw lastErr || new Error('Drive copy failed');
  } finally {
    lock.releaseLock();
  }
}

function drive_finalizeCopiedFile_(fileId, destFolderId, newName) {
  const file = DriveApp.getFileById(fileId);
  if (newName) {
    file.setName(newName);
  }

  if (destFolderId) {
    let inDest = false;
    const parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === destFolderId) {
        inDest = true;
        break;
      }
    }
    if (!inDest) {
      const destFolder = DriveApp.getFolderById(destFolderId);
      destFolder.addFile(file);
    }
  }
}

function drive_copyFile_(srcFileId, destFolderId, newName) {
  const hasAdvancedDrive = (() => {
    try { return !!Drive && !!Drive.Files && typeof Drive.Files.copy === 'function'; }
    catch (e) { return false; }
  })();

  if (hasAdvancedDrive) {
    const resource = { title: newName, parents: [{ id: destFolderId }] };
    const copied = Drive.Files.copy(resource, srcFileId);
    return copied.id;
  }

  const src = DriveApp.getFileById(srcFileId);
  const folder = DriveApp.getFolderById(destFolderId);
  const copy = src.makeCopy(newName, folder);
  return copy.getId();
}

function drive_openSpreadsheetWithRetry_(fileId, maxWaitMs) {
  const start = Date.now();
  let lastErr = null;
  while (Date.now() - start < maxWaitMs) {
    try {
      SpreadsheetApp.openById(fileId);
      return true;
    } catch (e) {
      lastErr = e;
      Utilities.sleep(1500);
    }
  }
  throw lastErr || new Error('Timed out waiting for spreadsheet to become available');
}

function drive_getFileSizeSafe_(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    return Number(file.getSize() || 0);
  } catch (e) {
    return 0;
  }
}

function drive_getOrCreateSubfolderId_(parentFolderId, subfolderName) {
  const parent = DriveApp.getFolderById(parentFolderId);
  const it = parent.getFoldersByName(subfolderName);
  if (it.hasNext()) return it.next().getId();
  const created = parent.createFolder(subfolderName);
  return created.getId();
}

/**
 * "Moves" a file by adding it to destination folder and removing it from source folder.
 * Note: Drive supports multi-parent; this enforces a single-parent-like behavior for your structure.
 */
function drive_moveFileToFolder_(fileId, fromFolderId, toFolderId) {
  const hasAdvancedDrive = (() => {
    try { return !!Drive && !!Drive.Files && typeof Drive.Files.update === 'function'; }
    catch (e) { return false; }
  })();

  if (hasAdvancedDrive) {
    try {
      Drive.Files.update({}, fileId, null, {
        addParents: toFolderId,
        removeParents: fromFolderId,
        supportsAllDrives: true
      });
      return true;
    } catch (e) {
      log_warn_('Drive API move failed, falling back to DriveApp', { fileId, error: e.message });
    }
  }

  const file = DriveApp.getFileById(fileId);
  const fromFolder = DriveApp.getFolderById(fromFolderId);
  const toFolder = DriveApp.getFolderById(toFolderId);

  toFolder.addFile(file);

  // Remove only from the specified folder
  try { fromFolder.removeFile(file); } catch (e) {
    // If file was not in fromFolder or insufficient rights, ignore
  }
  return true;
}

function drive_listFilesInFolder_(folderId) {
  const cached = drive_listFilesFromIndex_(folderId);
  if (cached !== null) return cached;
  const folder = DriveApp.getFolderById(folderId);
  const it = folder.getFiles();
  const out = [];
  while (it.hasNext()) {
    const f = it.next();
    out.push({
      id: f.getId(),
      name: f.getName(),
      url: f.getUrl(),
      mimeType: f.getMimeType(),
      lastUpdated: f.getLastUpdated()
    });
  }
  return out;
}

const DRIVE_INDEX_META_KEY = 'DRIVE_INDEX_META';
const DRIVE_INDEX_RESUME_KEY = 'DRIVE_INDEX_RESUME';
const DRIVE_INDEX_CHUNK_PREFIX = 'DRIVE_INDEX_CHUNK_';
const DRIVE_INDEX_CACHE_KEY = 'DRIVE_INDEX_CACHE';
const DRIVE_INDEX_CACHE_TTL = 300;

function drive_hasAdvanced_() {
  try { return !!Drive && !!Drive.Files && typeof Drive.Files.list === 'function'; }
  catch (e) { return false; }
}

function drive_indexProps_() {
  return PropertiesService.getScriptProperties();
}

function drive_indexCache_() {
  return CacheService.getScriptCache();
}

function drive_getIndexMeta_() {
  const raw = drive_indexProps_().getProperty(DRIVE_INDEX_META_KEY);
  return raw ? JSON.parse(raw) : null;
}

function drive_setIndexMeta_(meta) {
  drive_indexProps_().setProperty(DRIVE_INDEX_META_KEY, JSON.stringify(meta || {}));
}

function drive_clearIndex_() {
  const props = drive_indexProps_();
  const meta = drive_getIndexMeta_();
  if (meta && meta.chunkCount) {
    for (let i = 0; i < meta.chunkCount; i++) {
      props.deleteProperty(`${DRIVE_INDEX_CHUNK_PREFIX}${i}`);
    }
  }
  props.deleteProperty(DRIVE_INDEX_META_KEY);
  props.deleteProperty(DRIVE_INDEX_RESUME_KEY);
  drive_indexCache_().remove(DRIVE_INDEX_CACHE_KEY);
}

function drive_storeIndexChunk_(chunkIndex, entries) {
  drive_indexProps_().setProperty(`${DRIVE_INDEX_CHUNK_PREFIX}${chunkIndex}`, JSON.stringify(entries || []));
}

function drive_getIndexResume_() {
  const raw = drive_indexProps_().getProperty(DRIVE_INDEX_RESUME_KEY);
  return raw ? JSON.parse(raw) : null;
}

function drive_setIndexResume_(resume) {
  drive_indexProps_().setProperty(DRIVE_INDEX_RESUME_KEY, JSON.stringify(resume || {}));
}

function drive_readIndexChunks_() {
  const meta = drive_getIndexMeta_();
  if (!meta || !meta.chunkCount) return [];
  const props = drive_indexProps_();
  const out = [];
  for (let i = 0; i < meta.chunkCount; i++) {
    const raw = props.getProperty(`${DRIVE_INDEX_CHUNK_PREFIX}${i}`);
    if (!raw) continue;
    try {
      const chunk = JSON.parse(raw);
      if (Array.isArray(chunk)) out.push(...chunk);
    } catch (e) {
      log_warn_('Drive index chunk parse failed', { chunk: i, error: e.message });
    }
  }
  return out;
}

function drive_getIndex_() {
  const cache = drive_indexCache_();
  const cached = cache.get(DRIVE_INDEX_CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const index = drive_readIndexChunks_();
  try {
    const raw = JSON.stringify(index || []);
    if (raw.length < 90000) {
      cache.put(DRIVE_INDEX_CACHE_KEY, raw, DRIVE_INDEX_CACHE_TTL);
    }
  } catch (e) {}
  return index;
}

function drive_listFilesFromIndex_(folderId) {
  const meta = drive_getIndexMeta_();
  if (!meta || !meta.completedAt) return null;
  const index = drive_getIndex_();
  const fid = String(folderId || '').trim();
  if (!fid) return null;
  return (index || []).filter(f => Array.isArray(f.parentIds) && f.parentIds.includes(fid));
}

function drive_getFileMetaCached_(fileId) {
  const id = String(fileId || '').trim();
  if (!id) return null;
  const cache = drive_indexCache_();
  const cached = cache.get(`DRIVE_META_${id}`);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const index = drive_getIndex_();
  const fromIndex = (index || []).find(f => f.id === id);
  if (fromIndex) {
    cache.put(`DRIVE_META_${id}`, JSON.stringify(fromIndex), DRIVE_INDEX_CACHE_TTL);
    return fromIndex;
  }
  try {
    const file = DriveApp.getFileById(id);
    const meta = {
      id,
      name: file.getName(),
      url: file.getUrl(),
      mimeType: file.getMimeType(),
      lastUpdated: file.getLastUpdated() ? file.getLastUpdated().toISOString() : '',
      parentIds: []
    };
    cache.put(`DRIVE_META_${id}`, JSON.stringify(meta), DRIVE_INDEX_CACHE_TTL);
    return meta;
  } catch (e) {
    return null;
  }
}

function drive_refreshIndexChunk_(opts) {
  opts = opts || {};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    return { ok: false, status: 'LOCKED', message: 'Drive index refresh already running.' };
  }
  try {
    const cfg = cfg_getAll_();
    const formsFolderId = cfg_get_('FORMS_FOLDER_ID');
    const releasedFolderId = cfg_get_('RELEASED_FOLDER_ID');
    const formsObsoleteId = drive_getOrCreateSubfolderId_(formsFolderId, 'Obsolete');
    const releasedObsoleteId = drive_getOrCreateSubfolderId_(releasedFolderId, 'Obsolete');
    const folderIds = [formsFolderId, releasedFolderId, formsObsoleteId, releasedObsoleteId].filter(Boolean);

    let resume = drive_getIndexResume_();
    if (!resume || opts.reset) {
      drive_clearIndex_();
      resume = {
        folderIds,
        folderIndex: 0,
        pageToken: '',
        chunkIndex: 0,
        count: 0,
        startedAt: new Date().toISOString()
      };
    }

    const pageSize = Math.max(50, Math.min(500, Number(opts.pageSize || 200)));
    const folderId = resume.folderIds[resume.folderIndex];
    if (!folderId) {
      drive_setIndexMeta_({ completedAt: new Date().toISOString(), chunkCount: resume.chunkIndex, count: resume.count, folderIds });
      drive_indexProps_().deleteProperty(DRIVE_INDEX_RESUME_KEY);
      return { ok: true, done: true, count: resume.count };
    }

    let entries = [];
    let nextPageToken = '';
    if (drive_hasAdvanced_()) {
      const res = Drive.Files.list({
        q: `'${folderId}' in parents and trashed = false`,
        maxResults: pageSize,
        pageToken: resume.pageToken || undefined,
        fields: 'items(id,title,mimeType,modifiedDate,alternateLink,parents(id)),nextPageToken'
      });
      const items = res.items || [];
      entries = items.map(item => ({
        id: item.id,
        name: item.title,
        url: item.alternateLink || '',
        mimeType: item.mimeType || '',
        lastUpdated: item.modifiedDate || '',
        parentIds: (item.parents || []).map(p => p.id)
      }));
      nextPageToken = res.nextPageToken || '';
    } else {
      const fallback = drive_listFilesInFolder_(folderId);
      entries = fallback.map(f => ({
        id: f.id,
        name: f.name,
        url: f.url,
        mimeType: f.mimeType,
        lastUpdated: f.lastUpdated ? new Date(f.lastUpdated).toISOString() : '',
        parentIds: [folderId]
      }));
      nextPageToken = '';
    }

    drive_storeIndexChunk_(resume.chunkIndex, entries);
    resume.chunkIndex += 1;
    resume.count += entries.length;

    if (nextPageToken) {
      resume.pageToken = nextPageToken;
    } else {
      resume.pageToken = '';
      resume.folderIndex += 1;
    }

    if (resume.folderIndex >= resume.folderIds.length && !resume.pageToken) {
      drive_setIndexMeta_({
        completedAt: new Date().toISOString(),
        chunkCount: resume.chunkIndex,
        count: resume.count,
        folderIds
      });
      drive_indexProps_().deleteProperty(DRIVE_INDEX_RESUME_KEY);
      return {
        ok: true,
        done: true,
        count: resume.count,
        benchmarkNote: 'Drive index refreshed. Calls reduced by caching metadata in PropertiesService + CacheService.'
      };
    }

    drive_setIndexResume_(resume);
    return {
      ok: true,
      done: false,
      count: resume.count,
      message: `Indexed ${entries.length} files from folder ${resume.folderIndex + 1}/${resume.folderIds.length}.`
    };
  } finally {
    lock.releaseLock();
  }
}
