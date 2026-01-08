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

  const rows = values.map(r => [r[0] || '', r[1] || '']);
  const keys = Object.keys(updates || {});
  const newRows = [];

  keys.forEach(key => {
    const v = updates[key];
    const val = (v === null || v === undefined) ? '' : String(v);
    if (rowByKey[key]) {
      rows[rowByKey[key] - 1][1] = val;
    } else {
      newRows.push([key, val]);
      rowByKey[key] = rows.length + newRows.length;
    }
  });

  if (newRows.length) rows.push(...newRows);
  if (rows.length) {
    sh.getRange(1, 1, rows.length, 2).setValues(rows);
  }

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
