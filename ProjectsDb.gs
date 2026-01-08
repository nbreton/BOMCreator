const PROJECTS_SHEET = 'PROJECTS';
const PROJECTS_HEADERS = [
  'ProjectKey',
  'ClusterGroup',
  'IncludeMDAOverride',
  'ClusterBuswaySupplier',
  'MdaBuswaySupplier',
  'Notes',
  'UpdatedAt',
  'UpdatedBy'
];

function projects_ensure_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(PROJECTS_SHEET);
  if (!sh) sh = ss.insertSheet(PROJECTS_SHEET);

  if (sh.getLastRow() === 0) {
    sh.appendRow(PROJECTS_HEADERS);
    sh.setFrozenRows(1);
  } else {
    projects_ensureHeaders_(sh);
  }
  return sh;
}

function projects_ensureHeaders_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const missing = PROJECTS_HEADERS.filter(h => !header.includes(h));
  if (!missing.length) return;
  sh.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
}

function projects_headerIndexMap_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const idx = {};
  header.forEach((h, i) => {
    if (h) idx[h] = i;
  });
  return idx;
}

function projects_getMap_() {
  const sh = projects_ensure_();
  const values = sh.getDataRange().getValues();
  const idx = projects_headerIndexMap_(sh);
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const pk = String(values[i][idx.ProjectKey] || '').trim();
    if (!pk) continue;
    const rawOverride = String(values[i][idx.IncludeMDAOverride] || '').trim();
    map[pk] = {
      clusterGroup: Number(values[i][idx.ClusterGroup] || ''),
      includeMdaOverride: projects_normalizeIncludeMdaOverride_(rawOverride),
      clusterBuswaySupplier: String(values[i][idx.ClusterBuswaySupplier] || '').trim(),
      mdaBuswaySupplier: String(values[i][idx.MdaBuswaySupplier] || '').trim(),
      notes: String(values[i][idx.Notes] || '').trim()
    };
  }
  return map;
}

function projects_normalizeIncludeMdaOverride_(value) {
  const v = String(value || '').trim().toUpperCase();
  if (!v) return '';
  if (['TRUE', 'YES', '1', 'Y'].includes(v)) return 'TRUE';
  if (['FALSE', 'NO', '0', 'N'].includes(v)) return 'FALSE';
  return v;
}

function projects_inferClusterGroup_(projectKey) {
  // Default: number after last hyphen (e.g., VBL2A-1 -> 1, LPP7A-2 -> 2)
  const pk = String(projectKey || '').trim();
  const m = pk.match(/-(\d+)\s*$/);
  if (m) return Number(m[1]);
  return 1; // safe default
}

function projects_shouldIncludeMda_(projectKey) {
  const map = projects_getMap_();
  const rec = map[String(projectKey || '').trim()];
  const inferred = projects_inferClusterGroup_(projectKey);
  const clusterGroup = (rec && isFinite(rec.clusterGroup) && rec.clusterGroup > 0) ? rec.clusterGroup : inferred;

  // Rule: ClusterGroup 1 => include MDA; otherwise no MDA
  let include = (clusterGroup === 1);

  if (rec && rec.includeMdaOverride) {
    const v = rec.includeMdaOverride.toUpperCase();
    if (['TRUE', 'YES', '1', 'Y'].includes(v)) include = true;
    if (['FALSE', 'NO', '0', 'N'].includes(v)) include = false;
  }
  return include;
}

function projects_getEffective_(projectKey) {
  const map = projects_getMap_();
  const pk = String(projectKey || '').trim();
  const rec = map[pk] || {};
  const inferred = projects_inferClusterGroup_(pk);
  const clusterGroup = (isFinite(rec.clusterGroup) && rec.clusterGroup > 0) ? rec.clusterGroup : inferred;
  const includeMda = projects_shouldIncludeMda_(pk);
  return {
    projectKey: pk,
    clusterGroup,
    includeMda,
    includeMdaOverride: rec.includeMdaOverride || '',
    clusterBuswaySupplier: rec.clusterBuswaySupplier || '',
    mdaBuswaySupplier: rec.mdaBuswaySupplier || '',
    notes: rec.notes || ''
  };
}

function projects_setClusterGroup_(projectKey, clusterGroup) {
  return projects_setSettings_(projectKey, { clusterGroup });
}

function projects_setSettings_(projectKey, settings) {
  auth_requireEditor_();

  const pk = String(projectKey || '').trim();
  if (!pk) throw new Error('Missing ProjectKey');
  const cfg = settings || {};
  const cg = (cfg.clusterGroup !== undefined && cfg.clusterGroup !== null) ? Number(cfg.clusterGroup) : null;
  if (cg !== null && (!isFinite(cg) || cg <= 0 || Math.floor(cg) !== cg)) {
    throw new Error('ClusterGroup must be a positive integer');
  }

  const includeMdaOverride = (cfg.includeMdaOverride === undefined || cfg.includeMdaOverride === null)
    ? ''
    : projects_normalizeIncludeMdaOverride_(cfg.includeMdaOverride);

  if (includeMdaOverride && !['TRUE', 'FALSE'].includes(includeMdaOverride)) {
    throw new Error('IncludeMDAOverride must be TRUE, FALSE, or blank');
  }

  const clusterBuswaySupplier = String(cfg.clusterBuswaySupplier || '').trim();
  const mdaBuswaySupplier = String(cfg.mdaBuswaySupplier || '').trim();

  const user = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '');
  const now = new Date();

  const sh = projects_ensure_();
  const values = sh.getDataRange().getValues();
  const idx = projects_headerIndexMap_(sh);

  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.ProjectKey] || '').trim() === pk) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    sh.appendRow([
      pk,
      cg || projects_inferClusterGroup_(pk),
      includeMdaOverride,
      clusterBuswaySupplier,
      mdaBuswaySupplier,
      '',
      now,
      user
    ]);
  } else {
    const row = values[rowIndex - 1];
    const newRow = row.slice();
    if (cg !== null) newRow[idx.ClusterGroup] = cg;
    if (cfg.includeMdaOverride !== undefined) newRow[idx.IncludeMDAOverride] = includeMdaOverride;
    if (cfg.clusterBuswaySupplier !== undefined) newRow[idx.ClusterBuswaySupplier] = clusterBuswaySupplier;
    if (cfg.mdaBuswaySupplier !== undefined) newRow[idx.MdaBuswaySupplier] = mdaBuswaySupplier;
    newRow[idx.UpdatedAt] = now;
    newRow[idx.UpdatedBy] = user;
    sh.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
  }

  return {
    ok: true,
    projectKey: pk,
    clusterGroup: cg,
    includeMdaOverride,
    clusterBuswaySupplier,
    mdaBuswaySupplier
  };
}

function projects_syncFromAgile_(projects) {
  const list = Array.isArray(projects) ? projects : [];
  if (!list.length) return { ok: true, added: 0 };

  const sh = projects_ensure_();
  const values = sh.getDataRange().getValues();
  const idx = projects_headerIndexMap_(sh);
  const existing = new Set();
  for (let i = 1; i < values.length; i++) {
    const pk = String(values[i][idx.ProjectKey] || '').trim();
    if (pk) existing.add(pk);
  }

  const user = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '');
  const now = new Date();
  const rows = [];

  list.forEach(p => {
    const pk = String(p.projectKey || '').trim();
    if (!pk || existing.has(pk)) return;
    const inferred = projects_inferClusterGroup_(pk);
    rows.push([pk, inferred, '', '', '', '', now, user]);
    existing.add(pk);
  });

  if (!rows.length) return { ok: true, added: 0 };
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, PROJECTS_HEADERS.length).setValues(rows);
  return { ok: true, added: rows.length };
}
