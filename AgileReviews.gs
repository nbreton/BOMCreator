const AGILE_REVIEWS_SHEET = 'AGILE_REVIEWS';
const AGILE_REVIEWS_HEADERS = [
  'TabName',
  'Site',
  'Part',
  'PartNorm',
  'ProjectKey',
  'Rev',
  'DownloadDate',
  'ProjectType',
  'ReviewStatus',
  'ReviewedAt',
  'ReviewedBy',
  'SummaryJson',
  'ExceptionsJson',
  'Notes'
];

function agile_review_ensure_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(AGILE_REVIEWS_SHEET);
  if (!sh) sh = ss.insertSheet(AGILE_REVIEWS_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(AGILE_REVIEWS_HEADERS);
    sh.setFrozenRows(1);
  } else {
    agile_review_ensureHeaders_(sh);
  }
  return sh;
}

function agile_review_ensureHeaders_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const missing = AGILE_REVIEWS_HEADERS.filter(h => !header.includes(h));
  if (!missing.length) return;
  sh.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
}

function agile_review_headerIndexMap_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const idx = {};
  header.forEach((h, i) => {
    if (h) idx[h] = i;
  });
  return idx;
}

function agile_review_projectType_(record) {
  const partNorm = String(record.partNorm || record.PartNorm || '').trim();
  if (partNorm.toUpperCase() === 'MDA') return 'MDA';
  const pk = String(record.projectKey || record.ProjectKey || '').trim();
  const effective = projects_getEffective_(pk);
  return effective.clusterGroup === 1 ? 'Cluster 1' : 'Cluster 2';
}

function agile_review_normHeader_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function agile_review_findHeaderIndex_(headerNorm, candidates) {
  for (const c of candidates) {
    const cn = agile_review_normHeader_(c);
    const idx = headerNorm.findIndex(h => h === cn || h.includes(cn));
    if (idx >= 0) return idx;
  }
  return -1;
}

function agile_review_parseQty_(value) {
  if (value === null || value === undefined) return 0;
  if (typeof value === 'string') {
    const cleaned = value.replace(/,/g, '').trim();
    if (!cleaned || cleaned === '-' || cleaned === '—') return 0;
    const n = Number(cleaned);
    return isFinite(n) ? n : 0;
  }
  const n = Number(value);
  return isFinite(n) ? n : 0;
}

function agile_review_loadShouldBe_() {
  if (globalThis.__AGILE_SHOULD_BE__) return globalThis.__AGILE_SHOULD_BE__;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('mBOM SHOULD BE /CLASSIFICATION');
  if (!sh) throw new Error('Missing sheet: mBOM SHOULD BE /CLASSIFICATION');

  const values = sh.getDataRange().getValues();
  if (!values.length) throw new Error('mBOM SHOULD BE /CLASSIFICATION is empty');

  const headerNorm = values[0].map(agile_review_normHeader_);
  const col = {
    wp: agile_review_findHeaderIndex_(headerNorm, ['work package', 'wp']),
    gpn: agile_review_findHeaderIndex_(headerNorm, ['gpn number', 'gpn']),
    classification: agile_review_findHeaderIndex_(headerNorm, ['classification']),
    qtyMda: agile_review_findHeaderIndex_(headerNorm, ['should be - mda qty', 'mda qty']),
    qtyCluster1: agile_review_findHeaderIndex_(headerNorm, ['should be - cluster 1 qty', 'cluster 1 qty']),
    qtyCluster2: agile_review_findHeaderIndex_(headerNorm, ['should be - cluster 2 qty', 'cluster 2 qty'])
  };

  const missing = Object.entries(col).filter(([, idx]) => idx < 0).map(([key]) => key);
  if (missing.length) throw new Error(`Missing required headers in should-be sheet: ${missing.join(', ')}`);

  const makeBucket = () => ({
    byGpn: {},
    byWpClass: {},
    totalsByWp: {},
    totalsByClass: {}
  });

  const data = {
    'MDA': makeBucket(),
    'Cluster 1': makeBucket(),
    'Cluster 2': makeBucket()
  };

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const gpn = String(row[col.gpn] || '').trim();
    const wp = String(row[col.wp] || '').trim();
    const classification = String(row[col.classification] || '').trim();
    if (!gpn && !wp && !classification) continue;

    const qtyByType = {
      'MDA': agile_review_parseQty_(row[col.qtyMda]),
      'Cluster 1': agile_review_parseQty_(row[col.qtyCluster1]),
      'Cluster 2': agile_review_parseQty_(row[col.qtyCluster2])
    };

    Object.entries(qtyByType).forEach(([type, qty]) => {
      if (!qty) return;
      const bucket = data[type];
      bucket.byGpn[gpn] = { gpn, wp, classification, qty };
      const wpKey = wp || 'UNSPECIFIED';
      const classKey = classification || 'UNCLASSIFIED';
      const wpClassKey = `${wpKey}||${classKey}`;
      bucket.byWpClass[wpClassKey] = (bucket.byWpClass[wpClassKey] || 0) + qty;
      bucket.totalsByWp[wpKey] = (bucket.totalsByWp[wpKey] || 0) + qty;
      bucket.totalsByClass[classKey] = (bucket.totalsByClass[classKey] || 0) + qty;
    });
  }

  globalThis.__AGILE_SHOULD_BE__ = data;
  return data;
}

function agile_review_loadAgileTab_(tabName, sourceId) {
  const tab = String(tabName || '').trim();
  if (!tab) {
    log_error_('Agile review missing tab name.', { tabName });
    throw new Error('Missing Agile tab name');
  }
  const srcId = String(sourceId || cfg_get_('DOWNLOAD_LIST_SS_ID') || '').trim();
  if (!srcId) {
    log_error_('Agile review missing DOWNLOAD_LIST_SS_ID.', { tab });
    throw new Error('Missing DOWNLOAD_LIST_SS_ID');
  }

  const ss = SpreadsheetApp.openById(srcId);
  const sh = ss.getSheetByName(tab);
  if (!sh) {
    log_error_('Agile review missing Agile tab.', { tab, sourceId: srcId });
    throw new Error(`Missing Agile tab: ${tab}`);
  }

  const lastRow = Math.max(1, sh.getLastRow());
  const lastCol = Math.max(1, sh.getLastColumn());
  const scanRows = Math.min(200, lastRow);
  const scanValues = sh.getRange(1, 1, scanRows, lastCol).getDisplayValues();

  let headerRowIndex = -1;
  let headerNorm = [];
  const scanDiagnostics = [];
  for (let i = 0; i < scanValues.length; i++) {
    const rowNorm = scanValues[i].map(agile_review_normHeader_);
    const hasGpn = agile_review_findHeaderIndex_(rowNorm, [
      'gpn number',
      'gpn',
      'item number',
      'number',
      'bom.item number',
      'bom.find number',
      'find number'
    ]) >= 0;
    const hasQty = agile_review_findHeaderIndex_(rowNorm, ['qty', 'quantity', 'total qty', 'bom.qty']) >= 0;
    if (i < 10) {
      scanDiagnostics.push({
        rowIndex: i + 1,
        hasGpn,
        hasQty,
        raw: scanValues[i],
        normalized: rowNorm
      });
    }
    if (hasGpn && hasQty) {
      headerRowIndex = i;
      headerNorm = rowNorm;
      break;
    }
  }

  if (headerRowIndex < 0) {
    log_warn_('Agile review header row not found.', {
      tab,
      sourceId: srcId,
      lastRow,
      lastCol,
      scanRows,
      sampleRows: scanDiagnostics
    });
    throw new Error(`Unable to locate header row in Agile tab: ${tab}`);
  }

  const col = {
    wp: agile_review_findHeaderIndex_(headerNorm, ['work package', 'wp']),
    gpn: agile_review_findHeaderIndex_(headerNorm, [
      'gpn number',
      'gpn',
      'item number',
      'number',
      'bom item number',
      'bom find number',
      'find number'
    ]),
    classification: agile_review_findHeaderIndex_(headerNorm, ['classification', 'family', 'commodity code', 'bom.bom category', 'bom.category']),
    qty: agile_review_findHeaderIndex_(headerNorm, ['qty', 'quantity', 'total qty', 'bom.qty']),
    itemType: agile_review_findHeaderIndex_(headerNorm, ['item type', 'bom.item type', 'part type', 'type']),
    description: agile_review_findHeaderIndex_(headerNorm, ['cad description', 'bom.item description', 'description'])
  };
  log_info_('Agile review header row detected.', {
    tab,
    sourceId: srcId,
    headerRow: headerRowIndex + 1,
    lastRow,
    lastCol,
    header: scanValues[headerRowIndex],
    columns: col
  });

  const startRow = headerRowIndex + 2;
  const dataRows = lastRow - startRow + 1;
  if (dataRows <= 0) return [];

  const values = sh.getRange(startRow, 1, dataRows, lastCol).getValues();
  const rows = [];
  let emptyStreak = 0;

  for (const row of values) {
    const gpn = String(row[col.gpn] || '').trim();
    const wp = col.wp >= 0 ? String(row[col.wp] || '').trim() : '';
    const classification = col.classification >= 0 ? String(row[col.classification] || '').trim() : '';
    const itemType = col.itemType >= 0 ? String(row[col.itemType] || '').trim() : '';
    const description = col.description >= 0 ? String(row[col.description] || '').trim() : '';
    if (!gpn && !wp && !classification && !itemType && !description) {
      emptyStreak += 1;
      if (emptyStreak >= 5) break;
      continue;
    }
    emptyStreak = 0;
    const qty = (col.qty >= 0) ? agile_review_parseQty_(row[col.qty]) : 1;
    rows.push({ gpn, wp, classification, qty, itemType, description });
  }

  return rows;
}

function agile_review_buildForTab_(record) {
  const tabName = String(record.tabName || record.TabName || '').trim();
  const site = String(record.site || record.Site || '').trim();
  const partNorm = String(record.partNorm || record.PartNorm || '').trim();
  const partRaw = String(record.partRaw || record.Part || '').trim();
  const projectKey = String(record.projectKey || record.ProjectKey || '').trim();
  const rev = record.rev !== undefined ? record.rev : record.Rev;
  const downloadDate = String(record.downloadDate || record.DownloadDate || '').trim();

  const summary = {
    tabName,
    site,
    partNorm,
    partRaw,
    projectKey,
    rev,
    downloadDate,
    projectType: agile_review_projectType_(record),
    totalLines: 0,
    totalQty: 0,
    uniqueGpn: 0,
    byWp: {},
    byWpClass: {},
    expectedByWpClass: {},
    notes: []
  };
  const exceptions = [];

  let shouldBe = null;
  try {
    shouldBe = agile_review_loadShouldBe_();
  } catch (e) {
    exceptions.push({
      type: 'CONFIG',
      message: e.message || String(e)
    });
    summary.notes.push('Should-be sheet unavailable.');
    return { summary, exceptions };
  }

  const expectedBucket = shouldBe[summary.projectType];
  if (!expectedBucket) {
    exceptions.push({
      type: 'CONFIG',
      message: `No should-be data for project type ${summary.projectType}`
    });
    return { summary, exceptions };
  }

  let agileRows = [];
  try {
    agileRows = agile_review_loadAgileTab_(tabName);
  } catch (e) {
    exceptions.push({
      type: 'SOURCE',
      message: e.message || String(e)
    });
    summary.notes.push('Agile tab could not be loaded.');
    return { summary, exceptions };
  }

  const actualByGpn = {};
  const actualByWpClass = {};
  const actualByWp = {};
  const actualByClass = {};

  const markException = (entry) => {
    exceptions.push({
      type: entry.type,
      workPackage: entry.workPackage || '',
      classification: entry.classification || '',
      gpn: entry.gpn || '',
      expectedQty: entry.expectedQty || 0,
      actualQty: entry.actualQty || 0,
      message: entry.message || ''
    });
  };

  agileRows.forEach(row => {
    const wp = row.wp || 'UNSPECIFIED';
    const classification = row.classification || 'UNCLASSIFIED';
    const key = `${wp}||${classification}`;

    const qty = row.qty || 0;
    summary.totalLines += 1;
    summary.totalQty += qty;

    if (row.gpn) {
      actualByGpn[row.gpn] = (actualByGpn[row.gpn] || 0) + qty;
    }

    actualByWpClass[key] = (actualByWpClass[key] || 0) + qty;
    actualByWp[wp] = (actualByWp[wp] || 0) + qty;
    actualByClass[classification] = (actualByClass[classification] || 0) + qty;

    const desc = `${row.itemType || ''} ${row.description || ''}`.toUpperCase();
    if (wp.toUpperCase() === 'WP5' && desc.includes('KOP-SPARES')) {
      markException({
        type: 'KOP-SPARES',
        workPackage: wp,
        classification,
        gpn: row.gpn,
        expectedQty: 0,
        actualQty: qty,
        message: 'KOP-SPARES item detected in WP5.'
      });
    }
  });

  summary.uniqueGpn = Object.keys(actualByGpn).length;
  summary.byWp = actualByWp;
  summary.byWpClass = actualByWpClass;
  summary.expectedByWpClass = expectedBucket.byWpClass || {};

  const strictWp = new Set(['WP1', 'WP4', 'WP5']);
  const totalsOnlyWp = new Set(['WP2']);
  const additionalWp = new Set(['WP3']);

  Object.values(expectedBucket.byGpn).forEach(expected => {
    const wp = expected.wp || 'UNSPECIFIED';
    if (totalsOnlyWp.has(wp.toUpperCase())) return;
    const actualQty = actualByGpn[expected.gpn] || 0;
    if (actualQty < expected.qty) {
      markException({
        type: 'MISSING_GPN',
        workPackage: wp,
        classification: expected.classification,
        gpn: expected.gpn,
        expectedQty: expected.qty,
        actualQty,
        message: 'Expected GPN missing or short.'
      });
    }
  });

  Object.keys(actualByGpn).forEach(gpn => {
    if (expectedBucket.byGpn[gpn]) return;
    const row = agileRows.find(r => r.gpn === gpn);
    const wp = (row && row.wp) ? String(row.wp).trim() : 'UNSPECIFIED';
    const classification = row ? row.classification : '';
    const upperWp = wp.toUpperCase();
    if (strictWp.has(upperWp)) {
      markException({
        type: 'EXTRA_GPN',
        workPackage: wp,
        classification,
        gpn,
        expectedQty: 0,
        actualQty: actualByGpn[gpn],
        message: 'Extra GPN detected in strict work package.'
      });
    } else if (additionalWp.has(upperWp)) {
      markException({
        type: 'ADDITIONAL_GPN',
        workPackage: wp,
        classification,
        gpn,
        expectedQty: 0,
        actualQty: actualByGpn[gpn],
        message: 'Additional GPN detected in WP3.'
      });
    }
  });

  Object.entries(expectedBucket.byWpClass).forEach(([wpClassKey, expectedQty]) => {
    const [wp, classification] = wpClassKey.split('||');
    const upperWp = String(wp || '').toUpperCase();
    const actualQty = actualByWpClass[wpClassKey] || 0;

    if (totalsOnlyWp.has(upperWp)) {
      if (actualQty !== expectedQty) {
        markException({
          type: 'TOTAL_MISMATCH',
          workPackage: wp,
          classification,
          expectedQty,
          actualQty,
          message: 'Total quantity mismatch for WP2 classification.'
        });
      }
    } else if (strictWp.has(upperWp) && actualQty > expectedQty) {
      markException({
        type: 'TOTAL_EXCESS',
        workPackage: wp,
        classification,
        expectedQty,
        actualQty,
        message: 'Total quantity exceeds expected for strict work package.'
      });
    }
  });

  return { summary, exceptions };
}

function agile_review_syncFromLatest_(opts) {
  const force = opts && opts.force === true;
  const latest = agile_listLatest_();
  if (!latest.length) return { ok: true, created: 0 };

  const sh = agile_review_ensure_();
  const values = sh.getDataRange().getValues();
  const idx = agile_review_headerIndexMap_(sh);
  const existing = new Set();
  for (let i = 1; i < values.length; i++) {
    const tab = String(values[i][idx.TabName] || '').trim();
    if (tab) existing.add(tab);
  }

  const rows = [];
  let created = 0;

  latest.forEach(record => {
    const tabName = String(record.tabName || '').trim();
    if (!tabName) return;
    if (!force && existing.has(tabName)) return;
    const { summary, exceptions } = agile_review_buildForTab_(record);
    rows.push([
      tabName,
      summary.site,
      summary.partNorm || summary.partRaw || '',
      summary.partNorm,
      summary.projectKey,
      summary.rev,
      summary.downloadDate,
      summary.projectType,
      'PENDING',
      '',
      '',
      JSON.stringify(summary),
      JSON.stringify(exceptions),
      ''
    ]);
    created += 1;
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, AGILE_REVIEWS_HEADERS.length).setValues(rows);
  }

  return { ok: true, created };
}

function agile_review_backfill_(opts) {
  auth_requireEditor_();
  const includeHistory = opts && opts.includeHistory === true;
  const rows = includeHistory ? agile_rowsToUi_(agile_readIndex_()) : agile_listLatest_();
  const sh = agile_review_ensure_();
  const values = sh.getDataRange().getValues();
  const idx = agile_review_headerIndexMap_(sh);

  const rowIndexByTab = {};
  for (let i = 1; i < values.length; i++) {
    const tab = String(values[i][idx.TabName] || '').trim();
    if (tab) rowIndexByTab[tab] = i + 1;
  }

  let updated = 0;
  let created = 0;

  rows.forEach(record => {
    const tabName = String(record.tabName || '').trim();
    if (!tabName) return;
    const { summary, exceptions } = agile_review_buildForTab_(record);
    const row = [
      tabName,
      summary.site,
      summary.partNorm || summary.partRaw || '',
      summary.partNorm,
      summary.projectKey,
      summary.rev,
      summary.downloadDate,
      summary.projectType,
      'PENDING',
      '',
      '',
      JSON.stringify(summary),
      JSON.stringify(exceptions),
      ''
    ];

    const existingRow = rowIndexByTab[tabName];
    if (existingRow) {
      sh.getRange(existingRow, 1, 1, row.length).setValues([row]);
      updated += 1;
    } else {
      sh.appendRow(row);
      created += 1;
    }
  });

  return { ok: true, created, updated };
}

function agile_review_list_(opts) {
  const statusFilter = String(opts?.status || '').trim().toUpperCase();
  const sh = agile_review_ensure_();
  const values = sh.getDataRange().getValues();
  const idx = agile_review_headerIndexMap_(sh);
  const out = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = String(row[idx.ReviewStatus] || '').trim().toUpperCase() || 'PENDING';
    if (statusFilter && status !== statusFilter) continue;
    const summary = (() => {
      try { return JSON.parse(String(row[idx.SummaryJson] || '{}')); } catch (e) { return {}; }
    })();
    const exceptions = (() => {
      try { return JSON.parse(String(row[idx.ExceptionsJson] || '[]')); } catch (e) { return []; }
    })();
    out.push({
      tabName: String(row[idx.TabName] || '').trim(),
      site: String(row[idx.Site] || '').trim(),
      part: String(row[idx.Part] || '').trim(),
      partNorm: String(row[idx.PartNorm] || '').trim(),
      projectKey: String(row[idx.ProjectKey] || '').trim(),
      rev: row[idx.Rev],
      downloadDate: String(row[idx.DownloadDate] || '').trim(),
      projectType: String(row[idx.ProjectType] || '').trim(),
      status,
      reviewedAt: row[idx.ReviewedAt],
      reviewedBy: String(row[idx.ReviewedBy] || '').trim(),
      summary,
      exceptions,
      notes: String(row[idx.Notes] || '').trim()
    });
  }

  out.sort((a, b) =>
    String(a.site || '').localeCompare(String(b.site || '')) ||
    String(a.partNorm || '').localeCompare(String(b.partNorm || '')) ||
    Number(b.rev || 0) - Number(a.rev || 0)
  );

  return out;
}

function agile_review_setStatus_(tabName, status, notes) {
  auth_requireEditor_();
  const tab = String(tabName || '').trim();
  if (!tab) throw new Error('Missing tabName');
  const normalizedStatus = String(status || '').trim().toUpperCase();
  if (!['APPROVED', 'REJECTED', 'PENDING'].includes(normalizedStatus)) {
    throw new Error('Status must be APPROVED, REJECTED, or PENDING');
  }

  const sh = agile_review_ensure_();
  const values = sh.getDataRange().getValues();
  const idx = agile_review_headerIndexMap_(sh);
  let rowIndex = -1;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx.TabName] || '').trim() === tab) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error(`No Agile review found for tab: ${tab}`);

  const user = auth_getUser_();
  const now = new Date();
  const row = values[rowIndex - 1];
  row[idx.ReviewStatus] = normalizedStatus;
  row[idx.ReviewedAt] = now;
  row[idx.ReviewedBy] = user.email || '';
  row[idx.Notes] = String(notes || '').trim();

  sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);

  return { ok: true, tabName: tab, status: normalizedStatus };
}

function agile_review_getPending_() {
  const rows = agile_review_list_({});
  return rows.filter(r => String(r.status || '').toUpperCase() !== 'APPROVED');
}
