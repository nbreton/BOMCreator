const LOG_SHEET_NAME_DEFAULT = 'LOGS';

function log_info_(msg, data) { log_write_('INFO', msg, data); }
function log_warn_(msg, data) { log_write_('WARN', msg, data); }
function log_error_(msg, data) { log_write_('ERROR', msg, data); }

function log_write_(level, msg, data) {
  const entry = {
    ts: new Date().toISOString(),
    level,
    user: log_getUser_(),
    message: String(msg || ''),
    data: data || null
  };

  try {
    Logger.log(JSON.stringify(entry));
  } catch (_) {
    Logger.log(`${entry.ts} [${entry.level}] ${entry.message}`);
  }

  if (!log_shouldWriteSheet_()) return;

  try {
    const ss = SpreadsheetApp.getActive();
    const sheetName = log_sheetName_();
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Timestamp', 'Level', 'User', 'Message', 'Data(JSON)']);
    }
    sh.appendRow([
      new Date(),
      entry.level,
      entry.user,
      entry.message,
      entry.data ? JSON.stringify(entry.data) : ''
    ]);
  } catch (_) {
    // Avoid throwing from logger
  }
}

function log_shouldWriteSheet_() {
  try {
    return cfg_bool_('LOGS_SHEET_ENABLED', false);
  } catch (e) {
    return false;
  }
}

function log_sheetName_() {
  try {
    const cfg = cfg_getAll_();
    return String(cfg.LOGS_SHEET_NAME || LOG_SHEET_NAME_DEFAULT).trim() || LOG_SHEET_NAME_DEFAULT;
  } catch (e) {
    return LOG_SHEET_NAME_DEFAULT;
  }
}

function log_getUser_() {
  let email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  if (!email) {
    try { email = Session.getEffectiveUser().getEmail() || ''; } catch (e) {}
  }
  return String(email || '').trim();
}
