const PROJECTS_SHEET = 'PROJECTS';

function projects_ensure_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(PROJECTS_SHEET);
  if (!sh) sh = ss.insertSheet(PROJECTS_SHEET);

  if (sh.getLastRow() === 0) {
    sh.appendRow(['ProjectKey', 'ClusterGroup', 'IncludeMDAOverride', 'Notes', 'UpdatedAt', 'UpdatedBy']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function projects_getMap_() {
  const sh = projects_ensure_();
  const values = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) {
    const pk = String(values[i][0] || '').trim();
    if (!pk) continue;
    map[pk] = {
      clusterGroup: Number(values[i][1] || ''),
      includeMdaOverride: String(values[i][2] || '').trim(), // '', 'TRUE', 'FALSE'
      notes: String(values[i][3] || '').trim()
    };
  }
  return map;
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
  return { projectKey: pk, clusterGroup, includeMda, notes: rec.notes || '' };
}

function projects_setClusterGroup_(projectKey, clusterGroup) {
  auth_requireEditor_();

  const pk = String(projectKey || '').trim();
  const cg = Number(clusterGroup);
  if (!pk) throw new Error('Missing ProjectKey');
  if (!isFinite(cg) || cg <= 0 || Math.floor(cg) !== cg) throw new Error('ClusterGroup must be a positive integer');

  const user = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '');
  const now = new Date();

  const sh = projects_ensure_();
  const values = sh.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === pk) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    sh.appendRow([pk, cg, '', '', now, user]);
  } else {
    sh.getRange(rowIndex, 2, 1, 3).setValues([[cg, values[rowIndex - 1][2] || '', values[rowIndex - 1][3] || '']]);
    sh.getRange(rowIndex, 5, 1, 2).setValues([[now, user]]);
  }

  return { ok: true, projectKey: pk, clusterGroup: cg };
}

function projects_syncFromAgile_(projects) {
  const list = Array.isArray(projects) ? projects : [];
  if (!list.length) return { ok: true, added: 0 };

  const sh = projects_ensure_();
  const values = sh.getDataRange().getValues();
  const existing = new Set();
  for (let i = 1; i < values.length; i++) {
    const pk = String(values[i][0] || '').trim();
    if (pk) existing.add(pk);
  }

  const user = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '');
  const now = new Date();
  const rows = [];

  list.forEach(p => {
    const pk = String(p.projectKey || '').trim();
    if (!pk || existing.has(pk)) return;
    const inferred = projects_inferClusterGroup_(pk);
    rows.push([pk, inferred, '', '', now, user]);
    existing.add(pk);
  });

  if (!rows.length) return { ok: true, added: 0 };
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  return { ok: true, added: rows.length };
}
