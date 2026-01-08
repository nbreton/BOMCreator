function agile_refreshIndex() {
  auth_requireEditor_();
  return agile_refreshIndex_();
}

function agile_refreshIndex_() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const cfg = cfg_getAll_();
    const sourceId = cfg_get_('DOWNLOAD_LIST_SS_ID');
    const indexSheetName = cfg_get_('DOWNLOAD_LIST_INDEX_SHEET');
    const headerRow = Number(cfg.AGILE_HEADER_ROW || 3);
    const startRow = Number(cfg.AGILE_DATA_START_ROW || (headerRow + 1));

    const src = SpreadsheetApp.openById(sourceId);
    const sh = src.getSheetByName(indexSheetName);
    if (!sh) throw new Error(`Cannot find sheet "${indexSheetName}" in download list spreadsheet`);

    const lastCol = Math.max(1, sh.getLastColumn());
    const rawHeader = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    const headerNorm = rawHeader.map(agile_normHeader_);

    const col = {
      site: agile_findHeaderNorm_(headerNorm, ['site']),
      part: agile_findHeaderNorm_(headerNorm, ['part']),
      tla: agile_findHeaderNorm_(headerNorm, ['tla ref', 'tla']),
      desc: agile_findHeaderNorm_(headerNorm, ['description', 'desc']),
      rev: agile_findHeaderNorm_(headerNorm, ['rev']),
      date: agile_findHeaderNorm_(headerNorm, ['date of downloading', 'date']),
      tab: agile_findHeaderNorm_(headerNorm, ['name of tab', 'tab']),
      eco: agile_findHeaderNorm_(headerNorm, ['eco']),
      busway: agile_findHeaderNorm_(headerNorm, ['busway supplier', 'busway'])
    };

    const lastRow = sh.getLastRow();
    const numRows = Math.max(0, lastRow - startRow + 1);
    const data = numRows ? sh.getRange(startRow, 1, numRows, lastCol).getValues() : [];

    const records = [];
    for (const r of data) {
      const tab = String(r[col.tab] || '').trim();
      const site = String(r[col.site] || '').trim();
      const partRaw = String(r[col.part] || '').trim();
      if (!tab || !site || !partRaw) continue;

      const partNorm = agile_normalizePart_(partRaw);
      const tlaRef = String(r[col.tla] || '').trim();
      const description = String(r[col.desc] || '').trim();
      const buswaySupplier = String(r[col.busway] || '').trim();

      const rev = Number(String(r[col.rev] || '').trim());
      const eco = String(r[col.eco] || '').trim();

      const d = r[col.date];
      const dateStr = (d instanceof Date)
        ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(d || '').trim();

      const projectKey = agile_projectKey_(site, partNorm);
      const agileKey = `${site}||${partNorm}`;

      records.push({
        site, partRaw, partNorm, projectKey,
        tlaRef, description, buswaySupplier,
        rev: isFinite(rev) ? rev : '',
        dateStr, tab, eco, agileKey
      });
    }

    const latestRevByKey = {};
    for (const rec of records) {
      const v = Number(rec.rev || 0);
      const cur = latestRevByKey[rec.agileKey];
      if (cur === undefined || v > cur) latestRevByKey[rec.agileKey] = v;
    }

    const approvals = agile_approval_getMap_();

    const ss = SpreadsheetApp.getActive();
    let outSh = ss.getSheetByName('AGILE_INDEX');
    if (!outSh) outSh = ss.insertSheet('AGILE_INDEX');
    outSh.clear();

    const out = [[
      'Site', 'Part', 'PartNorm', 'ProjectKey',
      'TlaRef', 'Description', 'BuswaySupplier',
      'Rev', 'DownloadDate', 'TabName', 'ECO',
      'IsLatest', 'ApprovalStatus'
    ]];

    records.sort((a, b) =>
      (a.site || '').localeCompare(b.site || '') ||
      (a.partNorm || '').localeCompare(b.partNorm || '') ||
      Number(b.rev || 0) - Number(a.rev || 0)
    );

    for (const rec of records) {
      const isLatest = Number(rec.rev || 0) === Number(latestRevByKey[rec.agileKey] || 0);
      const approvalStatus = approvals[rec.tab]?.status || 'PENDING';
      out.push([
        rec.site,
        rec.partRaw,
        rec.partNorm,
        rec.projectKey,
        rec.tlaRef,
        rec.description,
        rec.buswaySupplier,
        rec.rev,
        rec.dateStr,
        rec.tab,
        rec.eco,
        isLatest,
        approvalStatus
      ]);
    }

    outSh.getRange(1, 1, out.length, out[0].length).setValues(out);
    outSh.setFrozenRows(1);

    // Record refresh timestamp (for diagnostics)
    PropertiesService.getScriptProperties().setProperty('AGILE_INDEX_LAST_REFRESH_AT', new Date().toISOString());

    // Optional notification hook (if Notifications.gs exists)
    try {
      if (typeof globalThis['notif_onAgileIndexRefreshed_'] === 'function') {
        const latest = agile_listLatest_();
        notif_onAgileIndexRefreshed_(latest);
      }
    } catch (e) {
      // never break refresh because of notifications
    }

    log_info_('Agile index refreshed', { count: records.length });
    return { ok: true, count: records.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Index state (non-refreshing)
 */
function agile_indexState_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('AGILE_INDEX');
  const last = PropertiesService.getScriptProperties().getProperty('AGILE_INDEX_LAST_REFRESH_AT') || '';
  if (!sh) return { exists: false, rows: 0, lastRefreshAt: last };
  const rows = Math.max(0, sh.getLastRow() - 1); // excluding header
  return { exists: true, rows, lastRefreshAt: last };
}

function agile_normHeader_(s) {
  return String(s || '')
    .replace(/["']/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function agile_findHeaderNorm_(headerNorm, candidates) {
  for (const c of candidates) {
    const cn = agile_normHeader_(c);
    const idx = headerNorm.findIndex(h => h === cn || h.includes(cn));
    if (idx >= 0) return idx;
  }
  throw new Error(`Cannot find header matching any of: ${candidates.join(', ')}`);
}

function agile_normalizePart_(part) {
  const p = String(part || '').trim();
  if (!p) return '';
  if (/^mda$/i.test(p)) return 'MDA';
  const m = p.match(/^zone\s*(\d+)$/i) || p.match(/^zone\s+(\d+)$/i);
  if (m) return `Zone ${m[1]}`;
  return p;
}

function agile_projectKey_(site, partNorm) {
  if (/^Zone\s+\d+$/i.test(partNorm)) {
    const n = partNorm.replace(/^Zone\s+/i, '').trim();
    return `${site}-${n}`;
  }
  if (/^MDA$/i.test(partNorm)) return `${site}-MDA`;
  return `${site}-${partNorm.replace(/\s+/g, '')}`;
}

function agile_isTrue_(v) {
  return v === true || String(v || '').trim().toUpperCase() === 'TRUE';
}

/**
 * Read AGILE_INDEX without auto-refresh. Returns [] if missing/empty.
 */
function agile_readIndex_(opts) {
  opts = opts || {};
  const useCache = opts.useCache !== false;
  if (useCache && Array.isArray(globalThis.__AGILE_INDEX_CACHE__)) {
    return globalThis.__AGILE_INDEX_CACHE__;
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('AGILE_INDEX');
  if (!sh || sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());

  const approvals = agile_approval_getMap_();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx]);
    const tab = String(obj.TabName || '').trim();
    obj.ApprovalStatus = approvals[tab]?.status || 'PENDING';
    out.push(obj);
  }

  if (useCache) globalThis.__AGILE_INDEX_CACHE__ = out;
  return out;
}

function agile_listLatest_() {
  return agile_listLatestFromRows_(agile_readIndex_());
}

function agile_listLatestFromRows_(rows) {
  const latest = (rows || []).filter(r => agile_isTrue_(r.IsLatest));
  latest.sort((a, b) =>
    String(a.Site || '').localeCompare(String(b.Site || '')) ||
    String(a.PartNorm || '').localeCompare(String(b.PartNorm || ''))
  );
  return latest.map(r => ({
    site: String(r.Site || ''),
    partNorm: String(r.PartNorm || ''),
    rev: r.Rev,
    tabName: String(r.TabName || ''),
    downloadDate: String(r.DownloadDate || ''),
    eco: String(r.ECO || ''),
    approvalStatus: String(r.ApprovalStatus || 'PENDING'),
    buswaySupplier: String(r.BuswaySupplier || ''),
    tlaRef: String(r.TlaRef || ''),
    description: String(r.Description || '')
  }));
}

function agile_getLatestTab_(site, part) {
  const partNorm = agile_normalizePart_(part);
  const rows = agile_readIndex_().filter(r =>
    String(r.Site || '') === String(site || '') &&
    String(r.PartNorm || '') === String(partNorm || '') &&
    agile_isTrue_(r.IsLatest)
  );
  rows.sort((a, b) => Number(b.Rev || 0) - Number(a.Rev || 0));
  return rows[0] || null;
}

function agile_listTabs_(site, part) {
  const partNorm = agile_normalizePart_(part);
  const rows = agile_readIndex_().filter(r =>
    String(r.Site || '') === String(site || '') &&
    String(r.PartNorm || '') === String(partNorm || '')
  );
  rows.sort((a, b) => Number(b.Rev || 0) - Number(a.Rev || 0));
  return rows.map(r => ({
    site: String(r.Site || ''),
    partNorm: String(r.PartNorm || ''),
    rev: r.Rev,
    tabName: String(r.TabName || ''),
    downloadDate: String(r.DownloadDate || ''),
    eco: String(r.ECO || ''),
    approvalStatus: String(r.ApprovalStatus || 'PENDING'),
    buswaySupplier: String(r.BuswaySupplier || ''),
    isLatest: agile_isTrue_(r.IsLatest),
    tlaRef: String(r.TlaRef || ''),
    description: String(r.Description || '')
  }));
}

function agile_findTabByName_(tabName) {
  const tab = String(tabName || '').trim();
  if (!tab) return null;
  const rows = agile_readIndex_().filter(r => String(r.TabName || '').trim() === tab);
  if (!rows.length) return null;
  rows.sort((a, b) => Number(b.Rev || 0) - Number(a.Rev || 0));
  return rows[0] || null;
}

/**
 * Restored and required by dashboard_build_()
 */
function agile_getProjects_() {
  return agile_getProjectsFromRows_(agile_readIndex_());
}

function agile_getProjectsFromRows_(rows) {
  rows = rows || [];
  const latestZoneRows = rows.filter(r =>
    agile_isTrue_(r.IsLatest) &&
    /^Zone\s+\d+$/i.test(String(r.PartNorm || ''))
  );

  const projects = {};
  for (const r of latestZoneRows) {
    const projectKey = String(r.ProjectKey || '').trim();
    if (!projectKey) continue;

    const site = String(r.Site || '').trim();
    const zone = String(r.PartNorm || '').replace(/^Zone\s+/i, '').trim();
    const mda = agile_getLatestTab_(site, 'MDA'); // may be null

    projects[projectKey] = {
      projectKey,
      site,
      zone,

      clusterTab: String(r.TabName || ''),
      clusterRev: r.Rev,
      clusterEco: String(r.ECO || ''),
      clusterDate: String(r.DownloadDate || ''),
      clusterApproval: String(r.ApprovalStatus || 'PENDING'),
      clusterBuswaySupplier: String(r.BuswaySupplier || ''),

      mdaTab: mda ? String(mda.TabName || '') : '',
      mdaRev: mda ? mda.Rev : '',
      mdaEco: mda ? String(mda.ECO || '') : '',
      mdaDate: mda ? String(mda.DownloadDate || '') : '',
      mdaApproval: mda ? String(mda.ApprovalStatus || 'PENDING') : 'PENDING',
      mdaBuswaySupplier: mda ? String(mda.BuswaySupplier || '') : ''
    };
  }

  return Object.values(projects).sort((a, b) => a.projectKey.localeCompare(b.projectKey));
}
