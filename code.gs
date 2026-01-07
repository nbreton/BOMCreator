//Code.gs
function onOpen() {
 SpreadsheetApp.getUi()
   .createMenu('mBOM')
   .addItem('Open App (Window)', 'ui_openAppWindow')
   .addSeparator()
   .addItem('Refresh Agile Index', 'agile_refreshIndex')
   .addItem('Refresh Files Index (Drive)', 'files_refreshIndexFromDrive')
   .addSeparator()
   .addItem('Open Copy Monitoring (Jobs)', 'ui_openCopyMonitoring')
   .addToUi();
}


function ui_openAppWindow() {
 const html = HtmlService.createTemplateFromFile('Index')
   .evaluate()
   .setTitle('VERDON mBOM App')
   .setWidth(1320)
   .setHeight(880);


 SpreadsheetApp.getUi().showModelessDialog(html, 'VERDON mBOM App');
}


function ui_openCopyMonitoring() {
 // Opens the main app; monitoring is a dedicated tab in the UI.
 ui_openAppWindow();
}


function doGet(e) {
 const tpl = HtmlService.createTemplateFromFile('Index');
 tpl.__WEB_MODE__ = true;
 tpl.__QUERY__ = (e && e.parameter) ? e.parameter : {};
 return tpl.evaluate().setTitle('VERDON mBOM App');
}


function include(filename) {
 return HtmlService.createHtmlOutputFromFile(filename).getContent();
}




//Config.gs
const CFG_CACHE_SECONDS = 60;


function cfg_getAll_() {
 const cache = CacheService.getDocumentCache();
 const cached = cache.get('CFG_ALL');
 if (cached) return JSON.parse(cached);


 const ss = SpreadsheetApp.getActive();
 const sh = ss.getSheetByName('CONFIG');
 if (!sh) throw new Error('Missing sheet: CONFIG');


 const values = sh.getDataRange().getValues();
 const cfg = {};
 for (let i = 1; i < values.length; i++) {
   const k = String(values[i][0] || '').trim();
   if (!k) continue;
   cfg[k] = String(values[i][1] || '').trim();
 }


 cache.put('CFG_ALL', JSON.stringify(cfg), CFG_CACHE_SECONDS);
 return cfg;
}


function cfg_get_(key, required = true) {
 const cfg = cfg_getAll_();
 const v = cfg[key];
 if (required && (!v || !String(v).trim())) throw new Error(`Missing CONFIG value for key: ${key}`);
 return v;
}


function cfg_bool_(key, defaultVal = false) {
 const v = (cfg_getAll_()[key] || '').toString().trim().toUpperCase();
 if (!v) return defaultVal;
 return ['TRUE', 'YES', '1', 'Y', 'ON'].includes(v);
}


function cfg_list_() {
 const ss = SpreadsheetApp.getActive();
 const sh = ss.getSheetByName('CONFIG');
 if (!sh) throw new Error('Missing sheet: CONFIG');


 const values = sh.getDataRange().getValues();
 const out = [];
 for (let i = 1; i < values.length; i++) {
   const key = String(values[i][0] || '').trim();
   if (!key) continue;
   out.push({ key, value: String(values[i][1] || '').trim() });
 }
 out.sort((a, b) => a.key.localeCompare(b.key));
 return out;
}


function cfg_update_(updates) {
 const user = auth_getUser_(); // may be unknown, but update requires editor auth
 auth_requireEditor_();


 cfg_validateUpdates_(updates, user.email || '');


 const ss = SpreadsheetApp.getActive();
 const sh = ss.getSheetByName('CONFIG');
 if (!sh) throw new Error('Missing sheet: CONFIG');


 const values = sh.getDataRange().getValues();
 const rowByKey = {};
 for (let i = 1; i < values.length; i++) {
   const k = String(values[i][0] || '').trim();
   if (!k) continue;
   rowByKey[k] = i + 1;
 }


 const keys = Object.keys(updates || {});
 keys.forEach(key => {
   const v = updates[key];
   const val = (v === null || v === undefined) ? '' : String(v);
   if (rowByKey[key]) {
     sh.getRange(rowByKey[key], 2).setValue(val);
   } else {
     sh.appendRow([key, val]);
   }
 });


 CacheService.getDocumentCache().remove('CFG_ALL');
 return { ok: true, updated: keys.length };
}


function cfg_validateUpdates_(updates, userEmail) {
 updates = updates || {};


 if (Object.prototype.hasOwnProperty.call(updates, 'ALLOWED_EDITORS')) {
   const list = String(updates.ALLOWED_EDITORS || '')
     .split(',')
     .map(s => s.trim().toLowerCase())
     .filter(Boolean);


   // If they set a non-empty allowlist, ensure the current user remains included if known
   if (list.length && userEmail && !list.includes(userEmail.toLowerCase())) {
     throw new Error(`Refused: ALLOWED_EDITORS update would remove your access (${userEmail}).`);
   }
 }


 ['AGILE_HEADER_ROW', 'AGILE_DATA_START_ROW'].forEach(k => {
   if (Object.prototype.hasOwnProperty.call(updates, k) && String(updates[k]).trim() !== '') {
     const n = Number(updates[k]);
     if (!isFinite(n) || n <= 0 || Math.floor(n) !== n) {
       throw new Error(`Invalid ${k}. Must be a positive integer.`);
     }
   }
 });
}


/**
* Attempts to retrieve user identity in a robust way.
* In some Workspace contexts, email can be blank; we return isKnown=false in that case.
*/
function auth_getUser_() {
 let email = '';
 try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
 if (!email) {
   try { email = Session.getEffectiveUser().getEmail() || ''; } catch (e) {}
 }
 email = String(email || '').trim().toLowerCase();


 // Fallback: no email available
 if (!email) {
   return { isKnown: false, email: '', reason: 'Email not available in this context (Workspace policy / deployment mode).' };
 }
 return { isKnown: true, email, reason: '' };
}


/**
* Enforces editor allowlist for write actions.
* Rules:
* - If ALLOWED_EDITORS is empty/missing => allow writes (for initial setup).
* - If ALLOWED_EDITORS is non-empty and email is unknown => block with clear message.
* - If ALLOWED_EDITORS is non-empty and email not in list => block.
*/
function auth_requireEditor_() {
 const cfg = cfg_getAll_();
 const allowedRaw = (cfg['ALLOWED_EDITORS'] || '').trim();


 const user = auth_getUser_();


 // Note: if no allowlist is defined, do not block even if identity is unknown.
 if (!allowedRaw) return user.email || '(unknown user)';


 const allowed = allowedRaw
   .split(',')
   .map(s => s.trim().toLowerCase())
   .filter(Boolean);


 if (!user.isKnown) {
   throw new Error(
     'Access control is enabled (ALLOWED_EDITORS is set), but your email identity is not available in this environment.\n' +
     'Fix options:\n' +
     '1) Deploy as a Web App with appropriate access settings (recommended), or\n' +
     '2) Clear ALLOWED_EDITORS temporarily during setup, or\n' +
     '3) Ask your Workspace admin to allow user email visibility for Apps Script.'
   );
 }


 if (allowed.length && !allowed.includes(user.email)) {
   throw new Error(`Access denied for ${user.email}. Contact the mBOM admin.`);
 }
 return user.email;
}



//Log.gs
function log_info_(msg, data) { log_write_('INFO', msg, data); }
function log_warn_(msg, data) { log_write_('WARN', msg, data); }
function log_error_(msg, data) { log_write_('ERROR', msg, data); }


function log_write_(level, msg, data) {
 const ss = SpreadsheetApp.getActive();
 let sh = ss.getSheetByName('LOG');
 if (!sh) sh = ss.insertSheet('LOG');
 if (sh.getLastRow() === 0) {
   sh.appendRow(['Timestamp', 'Level', 'User', 'Message', 'Data(JSON)']);
 }
 const user = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '');
 const row = [new Date(), level, user, msg, data ? JSON.stringify(data) : ''];
 sh.appendRow(row);
}



//FilesDb.gs
const FILES_SHEET = 'FILES';


function files_ensure_() {
 const ss = SpreadsheetApp.getActive();
 let sh = ss.getSheetByName(FILES_SHEET);
 if (!sh) sh = ss.insertSheet(FILES_SHEET);
 if (sh.getLastRow() === 0) {
   sh.appendRow([
     'Type', 'ProjectKey', 'MbomRev', 'BaseFormRev', 'AgileTabMDA', 'AgileTabCluster',
     'AgileRevCluster', 'ECO', 'Description', 'FileId', 'Url',
     'CreatedAt', 'CreatedBy', 'Status', 'Notes'
   ]);
   sh.setFrozenRows(1);
 }
 return sh;
}


function files_list_(type) {
 const sh = files_ensure_();
 const values = sh.getDataRange().getValues();
 const headers = values[0];
 const out = [];
 for (let i = 1; i < values.length; i++) {
   const row = values[i];
   const obj = {};
   headers.forEach((h, idx) => obj[h] = row[idx]);
   if (type && obj.Type !== type) continue;
   out.push(obj);
 }
 return out;
}


function files_getByFileId_(fileId) {
 const id = String(fileId || '').trim();
 if (!id) return null;
 const rows = files_list_();
 return rows.find(r => String(r.FileId || '').trim() === id) || null;
}


function files_append_(rec) {
 const sh = files_ensure_();
 sh.appendRow([
   rec.type || '',
   rec.projectKey || '',
   rec.mbomRev || '',
   rec.baseFormRev || '',
   rec.agileTabMDA || '',
   rec.agileTabCluster || '',
   rec.agileRevCluster || '',
   rec.eco || '',
   rec.description || '',
   rec.fileId || '',
   rec.url || '',
   rec.createdAt ? new Date(rec.createdAt) : new Date(),
   rec.createdBy || '',
   rec.status || '',
   rec.notes || ''
 ]);
}


/**
* Upsert by FileId. If exists, overwrite row with provided fields (keeps blanks if you pass blanks).
*/
function files_upsertByFileId_(rec) {
 const sh = files_ensure_();
 const values = sh.getDataRange().getValues();
 const headers = values[0];
 const idxFileId = headers.indexOf('FileId');
 if (idxFileId < 0) throw new Error('FILES: missing FileId header');


 const id = String(rec.fileId || '').trim();
 if (!id) throw new Error('Upsert requires fileId');


 let rowIndex = -1;
 for (let i = 1; i < values.length; i++) {
   if (String(values[i][idxFileId] || '').trim() === id) {
     rowIndex = i + 1;
     break;
   }
 }


 const row = [
   rec.type || '',
   rec.projectKey || '',
   rec.mbomRev || '',
   rec.baseFormRev || '',
   rec.agileTabMDA || '',
   rec.agileTabCluster || '',
   rec.agileRevCluster || '',
   rec.eco || '',
   rec.description || '',
   rec.fileId || '',
   rec.url || '',
   rec.createdAt ? new Date(rec.createdAt) : '',
   rec.createdBy || '',
   rec.status || '',
   rec.notes || ''
 ];


 if (rowIndex === -1) {
   sh.appendRow(row);
   return { ok: true, inserted: true };
 } else {
   sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
   return { ok: true, inserted: false };
 }
}


function files_getLatestBy_(type, predicateFn) {
 const rows = files_list_(type);
 const filtered = rows.filter(predicateFn);
 filtered.sort((a, b) => Number(b.MbomRev || 0) - Number(a.MbomRev || 0));
 return filtered[0] || null;
}


function files_nextRev_(type, projectKey) {
 const latest = files_getLatestBy_(type, r => (r.ProjectKey || '') === projectKey);
 const n = Number(latest?.MbomRev || 0);
 return n + 1;
}


function files_setStatus_(fileId, status) {
 const sh = files_ensure_();
 const values = sh.getDataRange().getValues();
 const headers = values[0];
 const idxFileId = headers.indexOf('FileId');
 const idxStatus = headers.indexOf('Status');
 if (idxFileId < 0 || idxStatus < 0) throw new Error('FILES headers missing required columns');


 for (let i = 1; i < values.length; i++) {
   if (String(values[i][idxFileId]).trim() === String(fileId).trim()) {
     sh.getRange(i + 1, idxStatus + 1).setValue(String(status || '').trim());
     return true;
   }
 }
 return false;
}



//Agilelndex.gs
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
function agile_readIndex_() {
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
 return out;
}


function agile_listLatest_() {
 const rows = agile_readIndex_().filter(r => agile_isTrue_(r.IsLatest));
 rows.sort((a, b) =>
   String(a.Site || '').localeCompare(String(b.Site || '')) ||
   String(a.PartNorm || '').localeCompare(String(b.PartNorm || ''))
 );
 return rows.map(r => ({
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


/**
* Restored and required by dashboard_build_()
*/
function agile_getProjects_() {
 const rows = agile_readIndex_();
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



//MbomOps.gs
function mbom_createNewFormRevision_(params) {
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
}


function mbom_createReleasedForProject_(params) {
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
}


function mbom_obsoletePreviousReleased_(args) {
 const { projectKey, releasedFolderId, releasedObsoleteId } = args;


 const prev = files_getLatestBy_('RELEASED', r =>
   String(r.ProjectKey || '') === projectKey &&
   String(r.Status || '').toUpperCase() !== 'OBSOLETE'
 );
 if (!prev || !prev.FileId) return;


 try {
   drive_moveFileToFolder_(prev.FileId, releasedFolderId, releasedObsoleteId);
   files_setStatus_(prev.FileId, 'OBSOLETE');
   log_info_('Moved previous RELEASED to Obsolete', { projectKey, prevFileId: prev.FileId });
 } catch (e) {
   log_warn_('Failed to move previous RELEASED to Obsolete', { projectKey, prevFileId: prev.FileId, error: e.message });
 }
}


function mbom_obsoletePreviousForm_(args) {
 const { baseFormId, formsFolderId, formsObsoleteId } = args;
 if (!baseFormId) return;


 try {
   drive_moveFileToFolder_(baseFormId, formsFolderId, formsObsoleteId);
   files_setStatus_(baseFormId, 'OBSOLETE');
   log_info_('Moved previous Form to Obsolete', { baseFormId });
 } catch (e) {
   log_warn_('Failed to move previous Form to Obsolete', { baseFormId, error: e.message });
 }
}


function mbom_setAgileInputs_(ss, cfg) {
 const mdaSheet = ss.getSheetByName('INPUT_BOM_AGILE_MDA');
 const clSheet = ss.getSheetByName('INPUT_BOM_AGILE_Cluster');
 if (!mdaSheet || !clSheet) throw new Error('Missing input sheets: INPUT_BOM_AGILE_MDA / INPUT_BOM_AGILE_Cluster');


 mbom_setLabeledValue_(mdaSheet, 'FILE ID', cfg.downloadListId);
 mbom_setLabeledValue_(mdaSheet, 'SHEET NAME', cfg.mdaTabName || '');


 mbom_setLabeledValue_(clSheet, 'FILE ID', cfg.downloadListId);
 mbom_setLabeledValue_(clSheet, 'SHEET NAME', cfg.clusterTabName);


 SpreadsheetApp.flush();
}


function mbom_setLabeledValue_(sheet, label, value) {
 const rng = sheet.getRange(1, 1, Math.min(sheet.getMaxRows(), 200), 2).getValues();
 for (let i = 0; i < rng.length; i++) {
   const a = String(rng[i][0] || '').trim().toLowerCase();
   if (a === String(label).trim().toLowerCase()) {
     sheet.getRange(i + 1, 2).setValue(value);
     return true;
   }
 }
 if (label === 'FILE ID') sheet.getRange('B1').setValue(value);
 if (label === 'SHEET NAME') sheet.getRange('B2').setValue(value);
 return false;
}


function mbom_freezeAgileInputs_(ss, downloadListId, mdaTab, clusterTab) {
 const src = SpreadsheetApp.openById(downloadListId);


 const srcCl = src.getSheetByName(clusterTab);
 if (!srcCl) throw new Error(`Cannot find Agile tab: ${clusterTab}`);


 const dstCl = ss.getSheetByName('INPUT_BOM_AGILE_Cluster');
 const rangeA1 = 'A1:U400';
 dstCl.getRange(rangeA1).setValues(srcCl.getRange(rangeA1).getValues());


 if (mdaTab) {
   const srcMda = src.getSheetByName(mdaTab);
   if (!srcMda) throw new Error(`Cannot find Agile tab: ${mdaTab}`);
   const dstMda = ss.getSheetByName('INPUT_BOM_AGILE_MDA');
   dstMda.getRange(rangeA1).setValues(srcMda.getRange(rangeA1).getValues());
 }


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



//Jobs.gs
const JOB_PREFIX = 'JOB_';
const JOB_MAX_KEEP = 200;


function jobs_create_(type, params) {
 const user = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '');
 const id = `${Date.now()}_${Math.floor(Math.random() * 1e6)}`;
 const job = {
   id,
   type,
   status: 'QUEUED',
   createdAt: new Date().toISOString(),
   createdBy: user,
   startedAt: '',
   finishedAt: '',
   message: 'Queued',
   progressCurrent: 0,
   progressTotal: 0,
   cursor: 0,
   params: params || {},
   results: [],
   errors: []
 };
 jobs_put_(job);
 jobs_cleanupOld_();
 jobs_schedule_();
 return { ok: true, jobId: id };
}


function jobs_run_() {
 const lock = LockService.getScriptLock();
 if (!lock.tryLock(1000)) return;


 try {
   const job = jobs_getNext_();
   if (!job) {
     jobs_cleanupTriggers_();
     return;
   }


   if (job.status === 'QUEUED') {
     job.status = 'RUNNING';
     job.startedAt = job.startedAt || new Date().toISOString();
     job.message = 'Starting…';
     jobs_put_(job);
   }


   // Run one step/batch
   jobs_execute_(job);


   jobs_put_(job);


   // If still running, schedule next batch
   if (job.status === 'RUNNING') {
     jobs_schedule_();
   } else {
     jobs_cleanupTriggers_();
   }


 } catch (e) {
   log_error_('Job runner error', { message: e.message, stack: e.stack });
 } finally {
   lock.releaseLock();
 }
}


function jobs_execute_(job) {
 switch (job.type) {
   case 'CREATE_FORM':
     jobs_execCreateForm_(job);
     break;


   case 'CREATE_RELEASED_ONE':
     jobs_execCreateReleasedOne_(job);
     break;


   case 'CREATE_RELEASES_SELECTED':
     jobs_execCreateReleasedBatch_(job);
     break;


   case 'CREATE_RELEASES_ALL':
     jobs_prepareAllThenRun_(job);
     break;


   default:
     job.status = 'ERROR';
     job.finishedAt = new Date().toISOString();
     job.message = `Unknown job type: ${job.type}`;
     job.errors.push({ error: job.message });
 }
}


function jobs_execCreateForm_(job) {
 try {
   job.progressTotal = 1;
   job.progressCurrent = 0;
   job.message = 'Creating Form revision (copy)…';
   jobs_put_(job);


   const res = mbom_createNewFormRevision_(job.params || {});
   job.results.push(res);


   job.progressCurrent = 1;
   job.status = 'DONE';
   job.finishedAt = new Date().toISOString();
   job.message = 'Completed';


 } catch (e) {
   job.status = 'ERROR';
   job.finishedAt = new Date().toISOString();
   job.message = e.message;
   job.errors.push({ error: e.message });
   log_error_('CREATE_FORM failed', { jobId: job.id, error: e.message });
 }
}


function jobs_execCreateReleasedOne_(job) {
 try {
   job.progressTotal = 1;
   job.progressCurrent = 0;
   job.message = 'Creating RELEASED (copy)…';
   jobs_put_(job);


   const res = mbom_createReleasedForProject_(job.params || {});
   job.results.push(res);


   job.progressCurrent = 1;
   job.status = 'DONE';
   job.finishedAt = new Date().toISOString();
   job.message = 'Completed';


 } catch (e) {
   job.status = 'ERROR';
   job.finishedAt = new Date().toISOString();
   job.message = e.message;
   job.errors.push({ error: e.message });
   log_error_('CREATE_RELEASED_ONE failed', { jobId: job.id, error: e.message });
 }
}


function jobs_prepareAllThenRun_(job) {
 // Convert "ALL" into explicit project list once, then reuse batch executor.
 if (!job.params) job.params = {};
 if (!job.params.projectKeys || !Array.isArray(job.params.projectKeys)) {
   const rawProjects = agile_getProjects_(); // latest zone projects
   const keys = rawProjects.map(p => p.projectKey);


   // Optionally only include eligible approved (UI can request)
   if (job.params.onlyEligible === true) {
     const filtered = [];
     for (const pk of keys) {
       const eff = projects_getEffective_(pk);
       const includeMda = eff.includeMda;


       // Check approvals based on latest agile
       const site = pk.split('-')[0];
       const zone = pk.split('-')[1];
       const cl = agile_getLatestTab_(site, `Zone ${zone}`);
       if (!cl) continue;
       if (agile_approval_status_(cl.TabName) !== 'APPROVED') continue;


       if (includeMda) {
         const m = agile_getLatestTab_(site, 'MDA');
         if (!m) continue;
         if (agile_approval_status_(m.TabName) !== 'APPROVED') continue;
       }
       filtered.push(pk);
     }
     job.params.projectKeys = filtered;
   } else {
     job.params.projectKeys = keys;
   }
 }


 // Transform job to batch type and run
 job.type = 'CREATE_RELEASES_SELECTED';
 jobs_execCreateReleasedBatch_(job);
}


function jobs_execCreateReleasedBatch_(job) {
 const projectKeys = (job.params && Array.isArray(job.params.projectKeys)) ? job.params.projectKeys : [];
 const batchSize = Number((job.params && job.params.batchSize) || 3);


 if (!job.progressTotal) job.progressTotal = projectKeys.length;


 const end = Math.min(projectKeys.length, job.cursor + batchSize);
 job.message = `Creating RELEASED: ${job.cursor + 1}–${end} of ${projectKeys.length}`;
 jobs_put_(job);


 for (let i = job.cursor; i < end; i++) {
   const pk = projectKeys[i];
   try {
     const eff = projects_getEffective_(pk);
     const includeMda = (job.params.includeMdaOverride === true) ? true : (job.params.includeMdaOverride === false ? false : eff.includeMda);


     // Determine latest agile tabs unless provided per-project
     const site = pk.split('-')[0];
     const zone = pk.split('-')[1];


     const cl = agile_getLatestTab_(site, `Zone ${zone}`);
     if (!cl || !cl.TabName) throw new Error(`No latest Cluster Agile for ${pk}`);


     const m = includeMda ? agile_getLatestTab_(site, 'MDA') : null;
     if (includeMda && (!m || !m.TabName)) throw new Error(`MDA required but missing for ${pk}`);


     // Infer busway codes (user can override at job level)
     const clusterSupplier = String(cl.BuswaySupplier || '');
     const mdaSupplier = includeMda ? String(m.BuswaySupplier || '') : '';


     const buswayClusterCode = job.params.buswayClusterCode || jobs_inferClusterCode_(clusterSupplier);
     const buswayMdaCode = includeMda ? (job.params.buswayMdaCode || jobs_inferMdaCode_(mdaSupplier)) : '';


     const res = mbom_createReleasedForProject_({
       projectKey: pk,
       includeMda,
       agileTabCluster: cl.TabName,
       agileRevCluster: cl.Rev,
       agileTabMDA: includeMda ? m.TabName : '',
       eco: job.params.eco || '',
       description: job.params.description || 'Batch RELEASED creation',
       affectedItems: job.params.affectedItems || '',
       freezeAgileInputs: (job.params.freezeAgileInputs !== undefined) ? job.params.freezeAgileInputs : undefined,
       buswayClusterCode,
       buswayMdaCode
     });


     job.results.push({ projectKey: pk, url: res.url, fileId: res.fileId, name: res.name });
   } catch (e) {
     job.errors.push({ projectKey: pk, error: e.message });
     log_error_('Batch RELEASED failed', { jobId: job.id, projectKey: pk, error: e.message });
   }
 }


 job.cursor = end;
 job.progressCurrent = end;


 if (job.cursor >= projectKeys.length) {
   job.status = (job.errors.length > 0) ? 'DONE_WITH_ERRORS' : 'DONE';
   job.finishedAt = new Date().toISOString();
   job.message = 'Completed';
 } else {
   job.status = 'RUNNING';
 }
}


function jobs_inferMdaCode_(buswaySupplier) {
 const s = String(buswaySupplier || '').toUpperCase();
 if (s.includes('STARLINE')) return 'ST';
 if (s.includes('EI') || s.includes('E&I')) return 'EI';
 return '';
}


function jobs_inferClusterCode_(buswaySupplier) {
 const s = String(buswaySupplier || '').toUpperCase();
 if (s.includes('MARDIX')) return 'MA';
 if (s.includes('EAE')) return 'EA';
 if (s.includes('EI') || s.includes('E&I')) return 'EI';
 return '';
}


// -----------------------
// Storage + listing
// -----------------------
function jobs_props_() { return PropertiesService.getScriptProperties(); }
function jobs_key_(id) { return `${JOB_PREFIX}${id}`; }


function jobs_put_(job) {
 jobs_props_().setProperty(jobs_key_(job.id), JSON.stringify(job));
}


function jobs_get_(id) {
 const raw = jobs_props_().getProperty(jobs_key_(id));
 return raw ? JSON.parse(raw) : null;
}


function jobs_getNext_() {
 const props = jobs_props_().getProperties();
 const keys = Object.keys(props).filter(k => k.startsWith(JOB_PREFIX));
 const jobs = keys.map(k => JSON.parse(props[k]));


 // Oldest first, prioritize QUEUED then RUNNING
 jobs.sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt)));
 return jobs.find(j => j.status === 'QUEUED' || j.status === 'RUNNING') || null;
}


function jobs_list_(opts) {
 opts = opts || {};
 const limit = Number(opts.limit || 50);
 const activeOnly = opts.activeOnly === true;


 const props = jobs_props_().getProperties();
 const keys = Object.keys(props).filter(k => k.startsWith(JOB_PREFIX));
 const jobs = keys.map(k => JSON.parse(props[k]));


 let filtered = jobs;
 if (activeOnly) {
   filtered = jobs.filter(j => j.status === 'QUEUED' || j.status === 'RUNNING');
 }


 filtered.sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
 return filtered.slice(0, limit).map(jobs_publicView_);
}


function jobs_publicView_(job) {
 return {
   id: job.id,
   type: job.type,
   status: job.status,
   message: job.message || '',
   createdAt: job.createdAt,
   createdBy: job.createdBy,
   startedAt: job.startedAt,
   finishedAt: job.finishedAt,
   progressCurrent: job.progressCurrent || 0,
   progressTotal: job.progressTotal || 0,
   cursor: job.cursor || 0,
   resultsCount: (job.results || []).length,
   errorsCount: (job.errors || []).length,
   results: (job.results || []).slice(0, 10), // cap for UI
   errors: (job.errors || []).slice(0, 10)
 };
}


function jobs_summary_() {
 const props = jobs_props_().getProperties();
 const keys = Object.keys(props).filter(k => k.startsWith(JOB_PREFIX));
 const jobs = keys.map(k => JSON.parse(props[k]));


 const summary = { queued: 0, running: 0, done: 0, error: 0, doneWithErrors: 0 };
 let runningJob = null;


 for (const j of jobs) {
   if (j.status === 'QUEUED') summary.queued++;
   else if (j.status === 'RUNNING') { summary.running++; if (!runningJob) runningJob = j; }
   else if (j.status === 'DONE') summary.done++;
   else if (j.status === 'DONE_WITH_ERRORS') summary.doneWithErrors++;
   else if (j.status === 'ERROR') summary.error++;
 }


 return {
   summary,
   runningJob: runningJob ? jobs_publicView_(runningJob) : null
 };
}


function jobs_cleanupOld_() {
 const props = jobs_props_().getProperties();
 const keys = Object.keys(props).filter(k => k.startsWith(JOB_PREFIX));
 if (keys.length <= JOB_MAX_KEEP) return;


 const jobs = keys.map(k => ({ key: k, job: JSON.parse(props[k]) }));
 jobs.sort((a, b) => String(a.job.createdAt).localeCompare(String(b.job.createdAt)));


 const toDelete = jobs.slice(0, Math.max(0, jobs.length - JOB_MAX_KEEP));
 toDelete.forEach(x => jobs_props_().deleteProperty(x.key));
}


function jobs_getStatus_(jobId) {
 const job = jobs_get_(jobId);
 if (!job) return { ok: false, error: 'Job not found' };
 return { ok: true, job: jobs_publicView_(job) };
}


// -----------------------
// Triggers
// -----------------------
function jobs_schedule_() {
 jobs_cleanupTriggers_();
 ScriptApp.newTrigger('jobs_run_').timeBased().after(2000).create();
}


function jobs_cleanupTriggers_() {
 const triggers = ScriptApp.getProjectTriggers();
 triggers.forEach(t => {
   if (t.getHandlerFunction() === 'jobs_run_') ScriptApp.deleteTrigger(t);
 });
}



//Api.gs
function api_ping() {
 return { ok: true, ts: new Date().toISOString() };
}


function api_bootstrap() {
 const user = (() => {
   try {
     return (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '(unknown)');
   } catch (e) {
     return '(unknown)';
   }
 })();


 const cfg = (() => {
   try { return cfg_getAll_(); } catch (e) { return {}; }
 })();


 const webAppUrl = (() => {
   try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
 })();


 // Very small, always-fast payload
 return {
   ok: true,
   user,
   webAppUrl,
   config: {
     namePrefix: cfg.NAME_PREFIX || 'mBOM',
     freezeDefault: cfg_bool_('FREEZE_AGILE_INPUTS_DEFAULT', true),
     requireApprovedForm: cfg_bool_('REQUIRE_APPROVED_FORM', true),
     requireApprovedAgile: cfg_bool_('REQUIRE_APPROVED_AGILE', true),
     currentApprovedFormFileId: (cfg.CURRENT_APPROVED_FORM_FILE_ID || ''),
     companyLogoUrl: (cfg.COMPANY_LOGO_URL || ''),
     customerLogoUrl: (cfg.CUSTOMER_LOGO_URL || '')
   },
   indexState: (() => {
     try { return agile_indexState_(); } catch (e) { return { exists: false, rows: 0, lastRefreshAt: '' }; }
   })(),
   jobsSummary: (() => {
     try { return jobs_summary_(); } catch (e) { return { summary: { queued: 0, running: 0, done: 0, doneWithErrors: 0, error: 0 }, runningJob: null }; }
   })()
 };
}


function api_loadData(payload) {
 payload = payload || {};
 const diag = {
   startedAt: new Date().toISOString(),
   steps: [],
   limits: {},
   sizes: {},
   counts: {}
 };


 const step = (name, fn) => {
   const t0 = Date.now();
   try {
     const out = fn();
     diag.steps.push({ name, ms: Date.now() - t0, ok: true });
     return out;
   } catch (e) {
     diag.steps.push({ name, ms: Date.now() - t0, ok: false, error: e.message });
     throw e;
   }
 };


 try {
   // Limits (reduce payload risk)
   const limitForms = Number(payload.limitForms || 200);
   const limitReleases = Number(payload.limitReleases || 200);
   const limitAgileLatest = Number(payload.limitAgileLatest || 300);
   const jobsLimit = Number(payload.jobsLimit || 30);


   diag.limits = { limitForms, limitReleases, limitAgileLatest, jobsLimit };


   const dash = step('dashboard_build_', () => dashboard_build_());


   const configList = step('cfg_list_', () => cfg_list_());


   const formsAll = step('normalize_forms', () => dashboard_normalizeFilesForUi_(dash.forms || []));
   const releasesAll = step('normalize_releases', () => dashboard_normalizeFilesForUi_(dash.releases || []));


   const forms = formsAll.slice(0, limitForms);
   const releases = releasesAll.slice(0, limitReleases);


   const agileLatestAll = (dash.agileLatest || []);
   const agileLatest = agileLatestAll.slice(0, limitAgileLatest);


   const pendingForms = step('normalize_pendingForms', () => dashboard_normalizeFilesForUi_(dash.pendingForms || []));


   const jobsSummary = step('jobs_summary_', () => jobs_summary_());
   const jobsRecent = step('jobs_list_', () => {
     try {
       return jobs_list_({ limit: jobsLimit, activeOnly: false });
     } catch (e) {
       return [];
     }
   });


   // Notifications (optional)
   let notifStatus = { releasedQueueCount: 0, releasedLastSentAt: '', releasedNextSendAt: '' };
   let notifSettings = [];
   step('notif_optional', () => {
     try { if (typeof globalThis['notif_getStatus_'] === 'function') notifStatus = notif_getStatus_(); } catch (e) {}
     try { if (typeof globalThis['notif_listSettings_'] === 'function') notifSettings = notif_listSettings_(); } catch (e) {}
     return true;
   });


   diag.counts = {
     formsTotal: formsAll.length, formsSent: forms.length,
     releasesTotal: releasesAll.length, releasesSent: releases.length,
     agileLatestTotal: agileLatestAll.length, agileLatestSent: agileLatest.length,
     jobsSent: jobsRecent.length
   };


   // Approx size (best-effort)
   step('estimate_sizes', () => {
     const safeLen = (obj) => {
       try { return JSON.stringify(obj).length; } catch (e) { return -1; }
     };
     diag.sizes = {
       configList: safeLen(configList),
       dashboard: safeLen({ indexState: dash.indexState, projects: dash.projects, latestApprovedForm: dash.latestApprovedForm }),
       forms: safeLen(forms),
       releases: safeLen(releases),
       agileLatest: safeLen(agileLatest),
       jobsRecent: safeLen(jobsRecent),
       notifications: safeLen({ notifStatus, notifSettings })
     };
     return true;
   });


   return {
     ok: true,
     __diag: diag,


     configList,


     dashboard: {
       indexState: dash.indexState,
       projects: dash.projects || [],
       agileLatest,
       pendingAgile: (dash.pendingAgile || []).slice(0, limitAgileLatest),
       pendingForms,
       latestApprovedForm: dash.latestApprovedForm ? {
         mbomRev: dash.latestApprovedForm.MbomRev,
         fileId: String(dash.latestApprovedForm.FileId || ''),
         url: String(dash.latestApprovedForm.Url || ''),
         status: String(dash.latestApprovedForm.Status || '')
       } : null
     },


     forms,
     releases,


     jobs: {
       summary: jobsSummary.summary,
       runningJob: jobsSummary.runningJob,
       recent: jobsRecent
     },


     notifications: {
       status: notifStatus,
       settings: notifSettings
     }
   };


 } catch (e) {
   // Server-side log for Apps Script "Executions"
   try {
     log_error_('api_loadData failed', { error: e.message, stack: e.stack, diag });
   } catch (_) {}


   return {
     ok: false,
     error: e.message || String(e),
     stack: e.stack || '',
     __diag: diag
   };
 }
}




/* --- keep your existing endpoints below (unchanged) --- */


function api_refreshAgileIndex() { auth_requireEditor_(); return agile_refreshIndex_(); }
function api_refreshFilesIndex() { auth_requireEditor_(); return files_refreshIndexFromDrive_(); }


function api_listAgileTabs(payload) { payload = payload || {}; return { ok: true, rows: agile_listTabs_(payload.site, payload.part) }; }
function api_setAgileApproval(payload) { auth_requireEditor_(); payload = payload || {}; return agile_approval_set_(payload.tabName, payload.status, payload.notes || ''); }


function api_setProjectClusterGroup(payload) { auth_requireEditor_(); payload = payload || {}; return projects_setClusterGroup_(payload.projectKey, payload.clusterGroup); }


function api_scheduleFormRevision(payload) { auth_requireEditor_(); payload = payload || {}; payload.changeRef = payload.changeRef || payload.ecrActRef || ''; return jobs_create_('CREATE_FORM', payload); }
function api_scheduleReleasedForProject(payload) { auth_requireEditor_(); payload = payload || {}; return jobs_create_('CREATE_RELEASED_ONE', payload); }
function api_scheduleReleasedForSelected(payload) {
 auth_requireEditor_();
 payload = payload || {};
 const projectKeys = Array.isArray(payload.projectKeys) ? payload.projectKeys : [];
 return jobs_create_('CREATE_RELEASES_SELECTED', {
   projectKeys,
   batchSize: payload.batchSize || 3,
   freezeAgileInputs: payload.freezeAgileInputs,
   description: payload.description || '',
   eco: payload.eco || '',
   affectedItems: payload.affectedItems || '',
   onlyEligible: payload.onlyEligible === true
 });
}
function api_scheduleReleasedForAll(payload) {
 auth_requireEditor_();
 payload = payload || {};
 return jobs_create_('CREATE_RELEASES_ALL', {
   batchSize: payload.batchSize || 3,
   freezeAgileInputs: payload.freezeAgileInputs,
   description: payload.description || '',
   eco: payload.eco || '',
   affectedItems: payload.affectedItems || '',
   onlyEligible: payload.onlyEligible === true
 });
}


function api_listJobs(payload) { payload = payload || {}; return { ok: true, summary: jobs_summary_(), jobs: jobs_list_({ limit: payload.limit || 50, activeOnly: payload.activeOnly === true }) }; }
function api_jobStatus(jobId) { return jobs_getStatus_(jobId); }


function api_setApprovedFormFileId(fileId) {
 auth_requireEditor_();
 const ss = SpreadsheetApp.getActive();
 const sh = ss.getSheetByName('CONFIG');
 if (!sh) throw new Error('Missing sheet: CONFIG');


 const rng = sh.getDataRange().getValues();
 for (let i = 1; i < rng.length; i++) {
   if (String(rng[i][0]).trim() === 'CURRENT_APPROVED_FORM_FILE_ID') {
     sh.getRange(i + 1, 2).setValue(String(fileId || '').trim());
     CacheService.getDocumentCache().remove('CFG_ALL');
     return { ok: true };
   }
 }
 sh.appendRow(['CURRENT_APPROVED_FORM_FILE_ID', String(fileId || '').trim()]);
 CacheService.getDocumentCache().remove('CFG_ALL');
 return { ok: true };
}


function api_setFileStatus(payload) { auth_requireEditor_(); payload = payload || {}; const ok = files_setStatus_(payload.fileId, payload.status); return { ok }; }
function api_updateConfig(payload) { auth_requireEditor_(); payload = payload || {}; return cfg_update_(payload.updates || {}); }


function api_getProjectData(payload) { payload = payload || {}; return { ok: true, data: projectdata_get_(payload.projectKey) }; }
function api_importReleased(payload) { auth_requireEditor_(); payload = payload || {}; return files_importReleasedFromText_(payload.text || ''); }


function api_updateNotifSettings(payload) { auth_requireEditor_(); payload = payload || {}; return notif_updateSettings_(payload.updates || {}); }


//DriveFacade.gs
function drive_copyFileWithRetry_(srcFileId, destFolderId, newName, maxAttempts = 6) {
 const lock = LockService.getScriptLock();
 lock.waitLock(30000);
 try {
   let lastErr = null;
   for (let attempt = 1; attempt <= maxAttempts; attempt++) {
     try {
       const fileId = drive_copyFile_(srcFileId, destFolderId, newName);
       drive_openSpreadsheetWithRetry_(fileId, 90 * 1000);
       return fileId;
     } catch (e) {
       lastErr = e;
       const sleepMs = Math.min(15000, 1000 * Math.pow(2, attempt));
       Utilities.sleep(sleepMs);
     }
   }
   throw lastErr || new Error('Drive copy failed');
 } finally {
   lock.releaseLock();
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



//AgileApprovals.gs
const AGILE_APPROVALS_SHEET = 'AGILE_APPROVALS';


function agile_approvals_ensure_() {
 const ss = SpreadsheetApp.getActive();
 let sh = ss.getSheetByName(AGILE_APPROVALS_SHEET);
 if (!sh) sh = ss.insertSheet(AGILE_APPROVALS_SHEET);


 if (sh.getLastRow() === 0) {
   sh.appendRow(['TabName', 'ApprovalStatus', 'ApprovedBy', 'ApprovedAt', 'Notes']);
   sh.setFrozenRows(1);
 }
 return sh;
}


function agile_approval_normalizeStatus_(status) {
 const s = String(status || 'PENDING').trim().toUpperCase();
 if (['PENDING', 'APPROVED', 'REJECTED'].includes(s)) return s;
 throw new Error(`Invalid Agile approval status: ${status}`);
}


function agile_approval_getMap_() {
 const sh = agile_approvals_ensure_();
 const values = sh.getDataRange().getValues();
 const map = {};
 for (let i = 1; i < values.length; i++) {
   const tab = String(values[i][0] || '').trim();
   if (!tab) continue;
   map[tab] = {
     status: String(values[i][1] || 'PENDING').trim().toUpperCase() || 'PENDING',
     approvedBy: String(values[i][2] || '').trim(),
     approvedAt: values[i][3] || '',
     notes: String(values[i][4] || '').trim()
   };
 }
 return map;
}


function agile_approval_status_(tabName) {
 const t = String(tabName || '').trim();
 if (!t) return 'PENDING';
 const map = agile_approval_getMap_();
 return (map[t]?.status) ? map[t].status : 'PENDING';
}


function agile_approval_set_(tabName, status, notes) {
 const t = String(tabName || '').trim();
 if (!t) throw new Error('Missing TabName for Agile approval.');


 const st = agile_approval_normalizeStatus_(status);
 const user = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '');
 const now = new Date();


 const sh = agile_approvals_ensure_();
 const values = sh.getDataRange().getValues();


 let rowIndex = -1;
 for (let i = 1; i < values.length; i++) {
   if (String(values[i][0] || '').trim() === t) {
     rowIndex = i + 1;
     break;
   }
 }


 if (rowIndex === -1) {
   sh.appendRow([t, st, user, now, String(notes || '')]);
 } else {
   sh.getRange(rowIndex, 2, 1, 4).setValues([[st, user, now, String(notes || '')]]);
 }


 return { ok: true, tabName: t, status: st };
}



//Dashboard.gs
function dashboard_build_() {
 const indexState = agile_indexState_();


 const rawProjects = agile_getProjects_(); // will be [] if index empty
 const forms = files_list_('FORM');
 const releases = files_list_('RELEASED');
 const agileLatest = agile_listLatest_();


 const projects = rawProjects.map(p => {
   const effective = projects_getEffective_(p.projectKey);
   return {
     ...p,
     clusterGroup: effective.clusterGroup,
     includeMda: effective.includeMda
   };
 });


 const latestReleaseByProject = {};
 for (const r of releases) {
   const pk = String(r.ProjectKey || '').trim();
   if (!pk) continue;
   const cur = latestReleaseByProject[pk];
   const rev = Number(r.MbomRev || 0);
   if (!cur || rev > Number(cur.MbomRev || 0)) latestReleaseByProject[pk] = r;
 }


 const projectsView = projects.map(p => {
   const rel = latestReleaseByProject[p.projectKey] || null;
   return {
     ...p,
     latestReleased: rel ? {
       mbomRev: rel.MbomRev,
       status: String(rel.Status || ''),
       url: String(rel.Url || ''),
       fileId: String(rel.FileId || ''),
       agileTabCluster: String(rel.AgileTabCluster || ''),
       agileTabMDA: String(rel.AgileTabMDA || '')
     } : null
   };
 });


 const approvedForms = forms
   .filter(f => String(f.Status || '').toUpperCase() === 'APPROVED')
   .sort((a, b) => Number(b.MbomRev || 0) - Number(a.MbomRev || 0));
 const latestApprovedForm = approvedForms[0] || null;


 const pendingAgile = agileLatest.filter(a => String(a.approvalStatus || '').toUpperCase() !== 'APPROVED');
 const pendingForms = forms.filter(f => String(f.Status || '').toUpperCase() !== 'APPROVED');


 return {
   indexState,
   projects: projectsView,
   forms,
   releases,
   agileLatest,
   latestApprovedForm,
   pendingAgile,
   pendingForms
 };
}


function dashboard_normalizeFilesForUi_(rows) {
 return (rows || []).map(r => ({
   type: String(r.Type || ''),
   projectKey: String(r.ProjectKey || ''),
   mbomRev: r.MbomRev,
   status: String(r.Status || ''),
   fileId: String(r.FileId || ''),
   url: String(r.Url || ''),
   eco: String(r.ECO || ''),
   description: String(r.Description || ''),
   createdBy: String(r.CreatedBy || ''),
   createdAt: (r.CreatedAt instanceof Date) ? r.CreatedAt.toISOString() : String(r.CreatedAt || '')
 }));
}



//ProjectsDb.gs
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



//Fileslndex.gs
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



//Notifications.gs
const NOTIF_CONFIG_SHEET = 'NOTIF_CONFIG';
const NOTIF_LOG_SHEET = 'NOTIF_LOG';


const PROP_AGILE_SNAPSHOT = 'NOTIF_AGILE_LATEST_SNAPSHOT';
const PROP_RELEASED_QUEUE = 'NOTIF_RELEASED_QUEUE';
const PROP_RELEASED_LAST_SENT_AT = 'NOTIF_RELEASED_LAST_SENT_AT';
const PROP_RELEASED_NEXT_SEND_AT = 'NOTIF_RELEASED_NEXT_SEND_AT';


function notif_ensureSheets_() {
 const ss = SpreadsheetApp.getActive();


 let cfg = ss.getSheetByName(NOTIF_CONFIG_SHEET);
 if (!cfg) cfg = ss.insertSheet(NOTIF_CONFIG_SHEET);
 if (cfg.getLastRow() === 0) {
   cfg.appendRow(['Key', 'Value', 'Description']);
   cfg.setFrozenRows(1);
 }


 const defaults = [
   ['ENABLE_EMAILS_GLOBAL', 'TRUE', 'Master switch for all emails (TRUE/FALSE)'],
   ['ENABLE_AGILE_PUBLISH_EMAIL', 'TRUE', 'Email when new latest Agile BOM is detected'],
   ['ENABLE_FORM_CREATED_EMAIL', 'TRUE', 'Email when a new Form revision is created'],
   ['ENABLE_RELEASED_DIGEST_EMAIL', 'TRUE', 'Digest email for RELEASED creations (grouped)'],
   ['ENGINEERING_RECIPIENTS', '', 'Comma-separated emails for Engineering notifications'],
   ['OPS_RECIPIENTS', '', 'Comma-separated emails for Procurement/Production notifications'],
   ['RELEASED_DIGEST_HOURS', '4', 'Grouping window (hours) for RELEASED digest'],
   ['EMAIL_SENDER_NAME', 'VERDON mBOM App', 'Sender name for emails'],
   ['EMAIL_SUBJECT_PREFIX', '[VERDON mBOM]', 'Subject prefix']
 ];


 const values = cfg.getDataRange().getValues();
 const existing = new Set(values.slice(1).map(r => String(r[0] || '').trim()).filter(Boolean));
 defaults.forEach(d => {
   if (!existing.has(d[0])) cfg.appendRow(d);
 });


 let log = ss.getSheetByName(NOTIF_LOG_SHEET);
 if (!log) log = ss.insertSheet(NOTIF_LOG_SHEET);
 if (log.getLastRow() === 0) {
   log.appendRow(['Timestamp', 'Type', 'To', 'Subject', 'Details(JSON)']);
   log.setFrozenRows(1);
 }
}


function notif_listSettings_() {
 notif_ensureSheets_();
 const ss = SpreadsheetApp.getActive();
 const sh = ss.getSheetByName(NOTIF_CONFIG_SHEET);
 const values = sh.getDataRange().getValues();
 const out = [];
 for (let i = 1; i < values.length; i++) {
   const key = String(values[i][0] || '').trim();
   if (!key) continue;
   out.push({ key, value: String(values[i][1] || '').trim(), desc: String(values[i][2] || '').trim() });
 }
 out.sort((a, b) => a.key.localeCompare(b.key));
 return out;
}


function notif_getSettingsMap_() {
 const rows = notif_listSettings_();
 const map = {};
 rows.forEach(r => map[r.key] = r.value);
 return map;
}


function notif_updateSettings_(updates) {
 notif_ensureSheets_();
 const ss = SpreadsheetApp.getActive();
 const sh = ss.getSheetByName(NOTIF_CONFIG_SHEET);
 const values = sh.getDataRange().getValues();


 const rowByKey = {};
 for (let i = 1; i < values.length; i++) {
   const k = String(values[i][0] || '').trim();
   if (!k) continue;
   rowByKey[k] = i + 1;
 }


 const keys = Object.keys(updates || {});
 keys.forEach(k => {
   const v = String(updates[k] ?? '').trim();
   if (rowByKey[k]) {
     sh.getRange(rowByKey[k], 2).setValue(v);
   } else {
     sh.appendRow([k, v, '']);
   }
 });


 return { ok: true, updated: keys.length };
}


function notif_bool_(settings, key, defVal) {
 const v = String(settings[key] || '').trim().toUpperCase();
 if (!v) return defVal;
 return ['TRUE', 'YES', '1', 'Y', 'ON'].includes(v);
}


function notif_num_(settings, key, defVal) {
 const n = Number(String(settings[key] || '').trim());
 return isFinite(n) ? n : defVal;
}


function notif_log_(type, to, subject, details) {
 try {
   notif_ensureSheets_();
   const ss = SpreadsheetApp.getActive();
   const sh = ss.getSheetByName(NOTIF_LOG_SHEET);
   sh.appendRow([new Date(), type, to, subject, details ? JSON.stringify(details) : '']);
 } catch (e) {
   // do not fail main operation because of logging
 }
}


function notif_sendEmail_(to, subject, htmlBody) {
 const settings = notif_getSettingsMap_();
 const senderName = String(settings.EMAIL_SENDER_NAME || 'VERDON mBOM App').trim();


 if (!to || !String(to).trim()) return { ok: false, skipped: true, reason: 'No recipients' };


 MailApp.sendEmail({
   to: String(to),
   subject: String(subject),
   htmlBody: String(htmlBody),
   name: senderName
 });
 return { ok: true };
}


function notif_htmlTemplate_(title, subtitle, contentHtml) {
 const cssCard = 'max-width:760px;margin:0 auto;background:#ffffff;border:1px solid #dadce0;border-radius:12px;overflow:hidden;';
 const cssHdr = 'padding:16px 18px;background:#1a73e8;color:#ffffff;font-family:Arial,Helvetica,sans-serif;font-size:16px;font-weight:700;';
 const cssSub = 'padding:0 18px 10px 18px;background:#1a73e8;color:#e8f0fe;font-family:Arial,Helvetica,sans-serif;font-size:12px;';
 const cssBody = 'padding:18px;font-family:Arial,Helvetica,sans-serif;color:#202124;font-size:13px;line-height:1.45;';
 const cssFooter = 'padding:12px 18px;border-top:1px solid #dadce0;color:#5f6368;font-family:Arial,Helvetica,sans-serif;font-size:11px;';


 const footer = `This email was generated automatically by the VERDON mBOM App.`;


 return `
 <div style="background:#f8f9fa;padding:24px;">
   <div style="${cssCard}">
     <div style="${cssHdr}">${title}</div>
     <div style="${cssSub}">${subtitle || ''}</div>
     <div style="${cssBody}">
       ${contentHtml}
     </div>
     <div style="${cssFooter}">${footer}</div>
   </div>
 </div>`;
}


/**
* Agile publish notification:
* Sends when a new "latest" Agile tab is detected for any Site/PartNorm.
* Prevents initial flood: first snapshot creation does NOT send emails.
*/
function notif_onAgileIndexRefreshed_(latestList) {
 notif_ensureSheets_();
 const settings = notif_getSettingsMap_();


 if (!notif_bool_(settings, 'ENABLE_EMAILS_GLOBAL', true)) return;
 if (!notif_bool_(settings, 'ENABLE_AGILE_PUBLISH_EMAIL', true)) return;


 const to = String(settings.ENGINEERING_RECIPIENTS || '').trim();
 if (!to) return;


 const props = PropertiesService.getScriptProperties();
 const prevRaw = props.getProperty(PROP_AGILE_SNAPSHOT);
 const prev = prevRaw ? JSON.parse(prevRaw) : null;


 const cur = {};
 (latestList || []).forEach(x => {
   const key = `${x.site}||${x.partNorm}`;
   cur[key] = { tabName: x.tabName, rev: x.rev, date: x.downloadDate, buswaySupplier: x.buswaySupplier || '' };
 });


 // First run: store snapshot only (no email)
 if (!prev) {
   props.setProperty(PROP_AGILE_SNAPSHOT, JSON.stringify(cur));
   return;
 }


 const changes = [];
 Object.keys(cur).forEach(k => {
   const prevTab = prev[k]?.tabName || '';
   const curTab = cur[k]?.tabName || '';
   if (prevTab !== curTab && curTab) {
     const parts = k.split('||');
     changes.push({
       site: parts[0],
       partNorm: parts[1],
       tabName: cur[k].tabName,
       rev: cur[k].rev,
       date: cur[k].date,
       buswaySupplier: cur[k].buswaySupplier
     });
   }
 });


 props.setProperty(PROP_AGILE_SNAPSHOT, JSON.stringify(cur));


 if (!changes.length) return;


 const cfg = cfg_getAll_();
 const prefix = String(settings.EMAIL_SUBJECT_PREFIX || '[VERDON mBOM]').trim();
 const downloadId = (cfg.DOWNLOAD_LIST_SS_ID || '').trim();
 const downloadUrl = downloadId ? `https://docs.google.com/spreadsheets/d/${downloadId}` : '';
 const appUrl = ScriptApp.getService().getUrl() || '';


 const tableRows = changes.map(c => `
   <tr>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${c.site}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${c.partNorm}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${c.rev}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${c.date}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;font-family:monospace;">${c.tabName}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${c.buswaySupplier || ''}</td>
   </tr>`).join('');


 const content = `
   <p>New Agile BOM revision(s) were detected as <b>latest</b>. Please review and approve them in the mBOM App.</p>
   <table style="border-collapse:collapse;width:100%;border:1px solid #dadce0;">
     <thead>
       <tr style="background:#f8f9fa;">
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Site</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Part</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Rev</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Date</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Tab</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Busway Supplier</th>
       </tr>
     </thead>
     <tbody>${tableRows}</tbody>
   </table>
   <p style="margin-top:12px;">
     ${appUrl ? `<a href="${appUrl}" target="_blank">Open mBOM App</a>` : ''}
     ${downloadUrl ? ` • <a href="${downloadUrl}" target="_blank">Open “Downloading List of BOM lev 0”</a>` : ''}
   </p>`;


 const subject = `${prefix} Agile BOM published – review required (${changes.length})`;
 const html = notif_htmlTemplate_('Agile BOM published', 'New latest revision(s) detected', content);


 const sendRes = notif_sendEmail_(to, subject, html);
 if (sendRes.ok) notif_log_('AGILE_PUBLISHED', to, subject, { changes });
}


/**
* Form created notification (Engineering).
*/
function notif_sendFormCreated_(info) {
 notif_ensureSheets_();
 const settings = notif_getSettingsMap_();


 if (!notif_bool_(settings, 'ENABLE_EMAILS_GLOBAL', true)) return;
 if (!notif_bool_(settings, 'ENABLE_FORM_CREATED_EMAIL', true)) return;


 const to = String(settings.ENGINEERING_RECIPIENTS || '').trim();
 if (!to) return;


 const prefix = String(settings.EMAIL_SUBJECT_PREFIX || '[VERDON mBOM]').trim();


 const content = `
   <p>A new <b>mBOM Form revision</b> has been created.</p>
   <ul>
     <li><b>Revision:</b> Rev ${info.mbomRev}</li>
     <li><b>ECR/ACT Ref:</b> ${info.changeRef || ''}</li>
     <li><b>Created by:</b> ${info.createdBy || ''}</li>
   </ul>
   <p><a href="${info.url}" target="_blank">Open Form Spreadsheet</a></p>
   <p class="muted">Next steps: review and mark APPROVED in the app when validated.</p>`;


 const subject = `${prefix} New mBOM Form created – Rev ${info.mbomRev}`;
 const html = notif_htmlTemplate_('New mBOM Form Revision', `Rev ${info.mbomRev}`, content);


 const sendRes = notif_sendEmail_(to, subject, html);
 if (sendRes.ok) notif_log_('FORM_CREATED', to, subject, info);
}


/**
* Released digest enqueue (Ops digest grouped over RELEASED_DIGEST_HOURS)
* Behavior:
* - If last digest was sent more than digestHours ago => send immediately (single digest) and start cooldown.
* - Otherwise queue and send at the end of cooldown (lastSent + digestHours).
*/
function notif_enqueueReleasedEvent_(evt) {
 notif_ensureSheets_();
 const settings = notif_getSettingsMap_();


 if (!notif_bool_(settings, 'ENABLE_EMAILS_GLOBAL', true)) return;
 if (!notif_bool_(settings, 'ENABLE_RELEASED_DIGEST_EMAIL', true)) return;


 const to = String(settings.OPS_RECIPIENTS || '').trim();
 if (!to) return;


 const hours = Math.max(1, notif_num_(settings, 'RELEASED_DIGEST_HOURS', 4));
 const windowMs = hours * 3600 * 1000;


 const props = PropertiesService.getScriptProperties();
 const now = Date.now();


 // queue
 const queue = notif_getReleasedQueue_();
 queue.push(evt);
 props.setProperty(PROP_RELEASED_QUEUE, JSON.stringify(queue));


 const lastSent = Number(props.getProperty(PROP_RELEASED_LAST_SENT_AT) || 0);


 if (!lastSent || (now - lastSent) >= windowMs) {
   // Send immediately
   notif_sendReleasedDigestNow_();
   props.setProperty(PROP_RELEASED_LAST_SENT_AT, String(Date.now()));
   // schedule next flush at end of new window
   notif_scheduleReleasedDigestAt_(Date.now() + windowMs);
 } else {
   // Ensure flush is scheduled at end of current window
   notif_scheduleReleasedDigestAt_(lastSent + windowMs);
 }
}


function notif_getReleasedQueue_() {
 const props = PropertiesService.getScriptProperties();
 const raw = props.getProperty(PROP_RELEASED_QUEUE);
 if (!raw) return [];
 try { return JSON.parse(raw) || []; } catch (e) { return []; }
}


function notif_scheduleReleasedDigestAt_(targetMs) {
 const props = PropertiesService.getScriptProperties();
 const existing = Number(props.getProperty(PROP_RELEASED_NEXT_SEND_AT) || 0);
 const now = Date.now();


 // If a trigger is already scheduled for this target time (or later), do nothing.
 if (existing && existing === targetMs && existing > now) return;


 // Replace any existing triggers for this handler
 const triggers = ScriptApp.getProjectTriggers();
 triggers.forEach(t => {
   if (t.getHandlerFunction() === 'notif_releasedDigestRunner_') ScriptApp.deleteTrigger(t);
 });


 const delay = Math.max(0, targetMs - now);
 ScriptApp.newTrigger('notif_releasedDigestRunner_').timeBased().after(delay).create();
 props.setProperty(PROP_RELEASED_NEXT_SEND_AT, String(targetMs));
}


function notif_releasedDigestRunner_() {
 notif_sendReleasedDigestNow_();


 const settings = notif_getSettingsMap_();
 const hours = Math.max(1, notif_num_(settings, 'RELEASED_DIGEST_HOURS', 4));
 const windowMs = hours * 3600 * 1000;


 const props = PropertiesService.getScriptProperties();
 const queue = notif_getReleasedQueue_();


 // If queue is still non-empty (unlikely), schedule again
 if (queue.length) {
   props.setProperty(PROP_RELEASED_LAST_SENT_AT, String(Date.now()));
   notif_scheduleReleasedDigestAt_(Date.now() + windowMs);
 } else {
   // Cleanup scheduled time
   props.deleteProperty(PROP_RELEASED_NEXT_SEND_AT);
   // Cleanup triggers (no-op if already executed)
   const triggers = ScriptApp.getProjectTriggers();
   triggers.forEach(t => {
     if (t.getHandlerFunction() === 'notif_releasedDigestRunner_') ScriptApp.deleteTrigger(t);
   });
 }
}


function notif_sendReleasedDigestNow_() {
 notif_ensureSheets_();
 const settings = notif_getSettingsMap_();


 if (!notif_bool_(settings, 'ENABLE_EMAILS_GLOBAL', true)) return;
 if (!notif_bool_(settings, 'ENABLE_RELEASED_DIGEST_EMAIL', true)) return;


 const to = String(settings.OPS_RECIPIENTS || '').trim();
 if (!to) return;


 const props = PropertiesService.getScriptProperties();
 const queue = notif_getReleasedQueue_();
 if (!queue.length) return;


 // Clear queue first to avoid duplicate sends if error occurs later
 props.deleteProperty(PROP_RELEASED_QUEUE);


 const prefix = String(settings.EMAIL_SUBJECT_PREFIX || '[VERDON mBOM]').trim();


 const rows = queue.map(e => `
   <tr>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${e.site || ''}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${e.projectKey}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">Rev ${e.mbomRev}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">
       <a href="${e.url}" target="_blank">Open</a>
     </td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;font-family:monospace;">${e.clusterTab || ''}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;font-family:monospace;">${e.mdaTab || ''}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${e.createdBy || ''}</td>
     <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${e.createdAt ? String(e.createdAt).replace('T',' ').replace('Z','') : ''}</td>
   </tr>`).join('');


 const content = `
   <p>The following RELEASED mBOM(s) were created.</p>
   <table style="border-collapse:collapse;width:100%;border:1px solid #dadce0;">
     <thead>
       <tr style="background:#f8f9fa;">
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Site</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Project</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Release</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Link</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Cluster Tab</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">MDA Tab</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Created by</th>
         <th style="text-align:left;padding:8px;border-bottom:1px solid #dadce0;">Created at</th>
       </tr>
     </thead>
     <tbody>${rows}</tbody>
   </table>`;


 const subject = `${prefix} RELEASED mBOM digest (${queue.length})`;
 const html = notif_htmlTemplate_('RELEASED mBOM digest', `Count: ${queue.length}`, content);


 const sendRes = notif_sendEmail_(to, subject, html);
 if (sendRes.ok) notif_log_('RELEASED_DIGEST', to, subject, { count: queue.length, items: queue });
}


function notif_getStatus_() {
 notif_ensureSheets_();
 const props = PropertiesService.getScriptProperties();
 const queue = notif_getReleasedQueue_();
 const lastSent = props.getProperty(PROP_RELEASED_LAST_SENT_AT) || '';
 const nextSend = props.getProperty(PROP_RELEASED_NEXT_SEND_AT) || '';


 return {
   releasedQueueCount: queue.length,
   releasedLastSentAt: lastSent ? new Date(Number(lastSent)).toISOString() : '',
   releasedNextSendAt: nextSend ? new Date(Number(nextSend)).toISOString() : ''
 };
}



//ProjectData.gs
function projectdata_get_(projectKey) {
 const pk = String(projectKey || '').trim();
 if (!pk) throw new Error('Missing projectKey');


 const eff = projects_getEffective_(pk);
 const site = pk.split('-')[0];
 const zone = pk.split('-')[1];
 const includeMda = eff.includeMda;
 const clusterPart = `Zone ${zone}`;


 const clusterAgile = agile_listTabs_(site, clusterPart);
 const mdaAgile = includeMda ? agile_listTabs_(site, 'MDA') : [];


 const releasesAll = files_list_('RELEASED');
 const released = releasesAll
   .filter(r => String(r.ProjectKey || '') === pk)
   .sort((a, b) => Number(b.MbomRev || 0) - Number(a.MbomRev || 0))
   .map(r => ({
     mbomRev: r.MbomRev,
     status: String(r.Status || ''),
     url: String(r.Url || ''),
     fileId: String(r.FileId || ''),
     agileTabCluster: String(r.AgileTabCluster || ''),
     agileTabMDA: String(r.AgileTabMDA || ''),
     createdAt: (r.CreatedAt instanceof Date) ? r.CreatedAt.toISOString() : String(r.CreatedAt || '')
   }));


 return {
   projectKey: pk,
   site,
   zone,
   clusterGroup: eff.clusterGroup,
   includeMda,
   clusterAgile,
   mdaAgile,
   released
 };
}



//Import.gs
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



