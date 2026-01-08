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
  const file = DriveApp.getFileById(fileId);
  const fromFolder = DriveApp.getFolderById(fromFolderId);
  const toFolder = DriveApp.getFolderById(toFolderId);

  toFolder.addFile(file);

  // Remove only from the specified folder
  try { fromFolder.removeFile(file); } catch (e) {
    // If file was not in fromFolder, ignore
  }
  return true;
}

function drive_listFilesInFolder_(folderId) {
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
