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
