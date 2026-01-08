function files_importReleasedFromText_(text) {
  auth_requireEditor_();

  const lines = String(text || '')
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(Boolean);

  let imported = 0;
  let skipped = 0;
  const errors = [];

  for (const line of lines) {
    try {
      const parts = line.split(',').map(s => s.trim()).filter(Boolean);

      let fileId = '';
      let projectKey = '';
      let rev = '';
      let status = 'RELEASED';

      if (parts.length >= 4) {
        fileId = extractFileId_(parts[0]);
        projectKey = parts[1];
        rev = parts[2];
        status = parts[3] || status;
      } else {
        fileId = extractFileId_(parts[0]);
      }

      if (!fileId) throw new Error(`Cannot extract file ID from: ${line}`);

      const file = DriveApp.getFileById(fileId);
      const name = file.getName();
      const url = file.getUrl();

      let parsed = files_parseMbomName_(name, (cfg_getAll_().NAME_PREFIX || '').trim());
      if (!parsed && parts.length < 4) {
        throw new Error(`Cannot parse name. Provide CSV: fileId,projectKey,rev,status. Name="${name}"`);
      }

      if (parsed) {
        projectKey = projectKey || parsed.projectKey;
        rev = rev || parsed.rev;
      }

      if (!projectKey || !rev) throw new Error(`Missing projectKey or rev for: ${line}`);

      const rec = {
        type: 'RELEASED',
        projectKey,
        mbomRev: Number(rev),
        baseFormRev: '',
        agileTabMDA: '',
        agileTabCluster: '',
        agileRevCluster: '',
        eco: '',
        description: 'Imported reference',
        fileId,
        url,
        createdAt: file.getLastUpdated() ? file.getLastUpdated().toISOString() : new Date().toISOString(),
        createdBy: '',
        status: String(status || 'RELEASED').toUpperCase(),
        notes: 'Imported via Import page'
      };

      const up = files_upsertByFileId_(rec);
      if (up.inserted) imported++;
      else imported++;

    } catch (e) {
      errors.push(String(e.message || e));
      skipped++;
    }
  }

  return { ok: true, imported, skipped, errors };
}

function extractFileId_(s) {
  const str = String(s || '').trim();
  if (!str) return '';
  // from URL
  const m = str.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m && m[1]) return m[1];
  // raw id
  if (/^[a-zA-Z0-9-_]{20,}$/.test(str)) return str;
  return '';
}
