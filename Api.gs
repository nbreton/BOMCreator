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

    const sections = Array.isArray(payload.sections) ? payload.sections.map(String) : [];
    const loadAll = sections.length === 0;
    const wantProjects = loadAll || sections.includes('projects');
    const wantForms = loadAll || sections.includes('forms') || sections.includes('approvals');
    const wantReleases = loadAll || sections.includes('releases');
    const wantAgileLatest = loadAll || sections.includes('agileLatest') || sections.includes('approvals');
    const wantJobs = loadAll || sections.includes('jobs');
    const wantConfig = loadAll || sections.includes('config');
    const wantNotifications = loadAll || sections.includes('notifications');

    diag.sections = {
      loadAll,
      wantProjects,
      wantForms,
      wantReleases,
      wantAgileLatest,
      wantJobs,
      wantConfig,
      wantNotifications
    };

    const needsDashboard = wantProjects || wantForms || wantReleases || wantAgileLatest;
    const dash = needsDashboard
      ? step('dashboard_build_', () => dashboard_build_({
        includeProjects: wantProjects,
        includeFormsList: wantForms,
        includeReleasesList: wantReleases,
        includeAgileLatest: wantAgileLatest,
        includePending: loadAll,
        includeLatestApprovedForm: wantProjects
      }))
      : { indexState: agile_indexState_(), projects: [], forms: [], releases: [], agileLatest: [], pendingAgile: [], pendingForms: [], latestApprovedForm: null };

    const configList = wantConfig ? step('cfg_list_', () => cfg_list_()) : [];

    const formsAll = wantForms ? step('normalize_forms', () => dashboard_normalizeFilesForUi_(dash.forms || [])) : [];
    const releasesAll = wantReleases ? step('normalize_releases', () => dashboard_normalizeFilesForUi_(dash.releases || [])) : [];

    const forms = wantForms ? formsAll.slice(0, limitForms) : [];
    const releases = wantReleases ? releasesAll.slice(0, limitReleases) : [];

    const agileLatestAll = wantAgileLatest ? (dash.agileLatest || []) : [];
    const agileLatest = wantAgileLatest ? agileLatestAll.slice(0, limitAgileLatest) : [];

    const pendingForms = loadAll
      ? step('normalize_pendingForms', () => dashboard_normalizeFilesForUi_(dash.pendingForms || []))
      : [];

    const jobsSummary = wantJobs ? step('jobs_summary_', () => jobs_summary_()) : { summary: { queued: 0, running: 0, done: 0, doneWithErrors: 0, error: 0 }, runningJob: null };
    const jobsRecent = wantJobs ? step('jobs_list_', () => {
      try {
        return jobs_list_({ limit: jobsLimit, activeOnly: false });
      } catch (e) {
        return [];
      }
    }) : [];

    // Notifications (optional)
    let notifStatus = { releasedQueueCount: 0, releasedLastSentAt: '', releasedNextSendAt: '' };
    let notifSettings = [];
    if (wantNotifications) {
      step('notif_optional', () => {
        try { if (typeof globalThis['notif_getStatus_'] === 'function') notifStatus = notif_getStatus_(); } catch (e) {}
        try { if (typeof globalThis['notif_listSettings_'] === 'function') notifSettings = notif_listSettings_(); } catch (e) {}
        return true;
      });
    }

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

    const response = {
      ok: true,
      __diag: diag,

      configList,

      dashboard: {
        indexState: dash.indexState,
        projects: dash.projects || [],
        agileLatest,
        pendingAgile: loadAll ? (dash.pendingAgile || []).slice(0, limitAgileLatest) : [],
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
    return api_safeReturn_(response, 'api_loadData response serialization failed');

  } catch (e) {
    // Server-side log for Apps Script "Executions"
    try {
      log_error_('api_loadData failed', { error: e.message, stack: e.stack, diag });
    } catch (_) {}

    const errorResponse = {
      ok: false,
      error: e.message || String(e),
      stack: e.stack || '',
      __diag: diag
    };
    return api_safeReturn_(errorResponse, 'api_loadData error serialization failed');
  }
}

/**
 * Ensures server responses are JSON-serializable for google.script.run.
 */
function api_safeReturn_(payload, fallbackMessage) {
  try {
    return JSON.parse(JSON.stringify(payload));
  } catch (e) {
    return {
      ok: false,
      error: `${fallbackMessage || 'Response serialization failed'}: ${e.message || e}`,
      stack: e.stack || '',
      __diag: payload && payload.__diag ? payload.__diag : {}
    };
  }
}



/* --- keep your existing endpoints below (unchanged) --- */

function api_refreshAgileIndex() { auth_requireEditor_(); return agile_refreshIndex_(); }
function api_refreshFilesIndex() { auth_requireEditor_(); return files_refreshIndexFromDrive_(); }

function api_listAgileTabs(payload) {
 payload = payload || {};
 const response = { ok: true, rows: agile_listTabs_(payload.site, payload.part) };
 return api_safeReturn_(response, 'api_listAgileTabs response serialization failed');
}

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

function api_getProjectData(payload) {
 payload = payload || {};
 const response = { ok: true, data: projectdata_get_(payload.projectKey) };
 return api_safeReturn_(response, 'api_getProjectData response serialization failed');
}

function api_importReleased(payload) { auth_requireEditor_(); payload = payload || {}; return files_importReleasedFromText_(payload.text || ''); }

function api_updateNotifSettings(payload) { auth_requireEditor_(); payload = payload || {}; return notif_updateSettings_(payload.updates || {}); }


function api_listJobs(payload) { payload = payload || {}; return { ok: true, summary: jobs_summary_(), jobs: jobs_list_({ limit: payload.limit || 50, activeOnly: payload.activeOnly === true }) }; }
function api_jobStatus(jobId) { return jobs_getStatus_(jobId); }
function api_retryJob(payload) { auth_requireEditor_(); payload = payload || {}; return jobs_retry_(payload.jobId); }
function api_removeJob(payload) { auth_requireEditor_(); payload = payload || {}; return jobs_remove_(payload.jobId); }
function api_restartJobs(payload) { auth_requireEditor_(); payload = payload || {}; return jobs_restartRunner_(); }
