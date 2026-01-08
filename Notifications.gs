const NOTIF_CONFIG_SHEET = 'NOTIF_CONFIG';
const NOTIF_LOG_SHEET = 'NOTIF_LOG';

const PROP_AGILE_SNAPSHOT = 'NOTIF_AGILE_LATEST_SNAPSHOT';
const PROP_RELEASED_QUEUE = 'NOTIF_RELEASED_QUEUE';
const PROP_RELEASED_LAST_SENT_AT = 'NOTIF_RELEASED_LAST_SENT_AT';
const PROP_RELEASED_NEXT_SEND_AT = 'NOTIF_RELEASED_NEXT_SEND_AT';

function notif_escapeHtml_(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function notif_safeUrl_(value) {
  const url = String(value || '').trim();
  if (!url) return '';
  if (!/^https?:\/\//i.test(url)) return '';
  return url;
}

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

  const rows = values.map(r => [r[0] || '', r[1] || '', r[2] || '']);
  const keys = Object.keys(updates || {});
  const newRows = [];

  keys.forEach(k => {
    const v = String(updates[k] ?? '').trim();
    if (rowByKey[k]) {
      rows[rowByKey[k] - 1][1] = v;
    } else {
      newRows.push([k, v, '']);
      rowByKey[k] = rows.length + newRows.length;
    }
  });

  if (newRows.length) rows.push(...newRows);
  if (rows.length) {
    sh.getRange(1, 1, rows.length, 3).setValues(rows);
  }

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
  const safeTitle = notif_escapeHtml_(title);
  const safeSubtitle = subtitle ? notif_escapeHtml_(subtitle) : '';
  const cssCard = 'max-width:760px;margin:0 auto;background:#ffffff;border:1px solid #dadce0;border-radius:12px;overflow:hidden;';
  const cssHdr = 'padding:16px 18px;background:#1a73e8;color:#ffffff;font-family:Arial,Helvetica,sans-serif;font-size:16px;font-weight:700;';
  const cssSub = 'padding:0 18px 10px 18px;background:#1a73e8;color:#e8f0fe;font-family:Arial,Helvetica,sans-serif;font-size:12px;';
  const cssBody = 'padding:18px;font-family:Arial,Helvetica,sans-serif;color:#202124;font-size:13px;line-height:1.45;';
  const cssFooter = 'padding:12px 18px;border-top:1px solid #dadce0;color:#5f6368;font-family:Arial,Helvetica,sans-serif;font-size:11px;';

  const footer = `This email was generated automatically by the VERDON mBOM App.`;

  return `
  <div style="background:#f8f9fa;padding:24px;">
    <div style="${cssCard}">
      <div style="${cssHdr}">${safeTitle}</div>
      <div style="${cssSub}">${safeSubtitle}</div>
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
  const downloadUrl = notif_safeUrl_(downloadId ? `https://docs.google.com/spreadsheets/d/${downloadId}` : '');
  const appUrl = notif_safeUrl_(ScriptApp.getService().getUrl() || '');

  const tableRows = changes.map(c => `
    <tr>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${notif_escapeHtml_(c.site)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${notif_escapeHtml_(c.partNorm)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${notif_escapeHtml_(c.rev)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${notif_escapeHtml_(c.date)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;font-family:monospace;">${notif_escapeHtml_(c.tabName)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${notif_escapeHtml_(c.buswaySupplier || '')}</td>
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
      ${appUrl ? `<a href="${appUrl}" target="_blank" rel="noopener">Open mBOM App</a>` : ''}
      ${downloadUrl ? ` • <a href="${downloadUrl}" target="_blank" rel="noopener">Open “Downloading List of BOM lev 0”</a>` : ''}
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

  const safeUrl = notif_safeUrl_(info.url);
  const subjectRev = String(info.mbomRev ?? '');
  const safeRev = notif_escapeHtml_(subjectRev);
  const safeChangeRef = notif_escapeHtml_(info.changeRef || '');
  const safeCreatedBy = notif_escapeHtml_(info.createdBy || '');
  const content = `
    <p>A new <b>mBOM Form revision</b> has been created.</p>
    <ul>
      <li><b>Revision:</b> Rev ${safeRev}</li>
      <li><b>ECR/ACT Ref:</b> ${safeChangeRef}</li>
      <li><b>Created by:</b> ${safeCreatedBy}</li>
    </ul>
    ${safeUrl ? `<p><a href="${safeUrl}" target="_blank" rel="noopener">Open Form Spreadsheet</a></p>` : ''}
    <p class="muted">Next steps: review and mark APPROVED in the app when validated.</p>`;

  const subject = `${prefix} New mBOM Form created – Rev ${subjectRev}`;
  const html = notif_htmlTemplate_('New mBOM Form Revision', `Rev ${subjectRev}`, content);

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

  const rows = queue.map(e => {
    const safeUrl = notif_safeUrl_(e.url);
    return `
    <tr>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${notif_escapeHtml_(e.site || '')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${notif_escapeHtml_(e.projectKey)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">Rev ${notif_escapeHtml_(e.mbomRev)}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">
        ${safeUrl ? `<a href="${safeUrl}" target="_blank" rel="noopener">Open</a>` : ''}
      </td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;font-family:monospace;">${notif_escapeHtml_(e.clusterTab || '')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;font-family:monospace;">${notif_escapeHtml_(e.mdaTab || '')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${notif_escapeHtml_(e.createdBy || '')}</td>
      <td style="padding:6px 8px;border-bottom:1px solid #dadce0;">${notif_escapeHtml_(e.createdAt ? String(e.createdAt).replace('T',' ').replace('Z','') : '')}</td>
    </tr>`;
  }).join('');

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
