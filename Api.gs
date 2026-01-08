const API_VERSION = '2024.09.0';

class ApiError extends Error {
  constructor(code, message, details) {
    super(message);
    this.name = 'ApiError';
    this.code = code || 'INTERNAL';
    this.details = details || null;
  }
}

function api_handleRequest_(name, handler) {
  const startedMs = Date.now();
  const requestId = api_requestId_();
  const stackId = api_stackId_();
  try {
    const data = handler();
    return api_response_(data, startedMs, requestId, stackId);
  } catch (e) {
    const errorInfo = api_classifyError_(e);

    try {
      log_error_('API error', {
        name,
        requestId,
        code: errorInfo.code,
        message: errorInfo.message,
        details: errorInfo.details,
        stackId,
        stack: e && e.stack ? e.stack : ''
      });
      logEvent('api_error', {
        name,
        requestId,
        stackId,
        code: errorInfo.code,
        message: errorInfo.message
      });
    } catch (_) {}

    return api_errorResponse_(
      errorInfo.code,
      errorInfo.message,
      errorInfo.details,
      requestId,
      stackId,
      startedMs
    );
  }
}

function api_response_(data, startedMs, requestId, stackId) {
  return api_safeReturn_({
    ok: true,
    data: (data === undefined ? null : data),
    meta: {
      ms: Date.now() - startedMs,
      version: API_VERSION,
      requestId: requestId || '',
      stackId: stackId || ''
    }
  }, 'Response serialization failed');
}

function api_errorResponse_(code, message, details, requestId, stackId, startedMs) {
  return api_safeReturn_({
    ok: false,
    error: {
      code: code || 'INTERNAL',
      message: message || 'Request failed',
      details: details || null,
      requestId: requestId || '',
      stackId: stackId || ''
    },
    meta: {
      ms: Date.now() - startedMs,
      version: API_VERSION,
      requestId: requestId || '',
      stackId: stackId || ''
    }
  }, 'Error response serialization failed');
}

function api_safeReturn_(payload, fallbackMessage) {
  try {
    return JSON.parse(JSON.stringify(payload));
  } catch (e) {
    return {
      ok: false,
      error: {
        code: 'SERIALIZATION_ERROR',
        message: `${fallbackMessage || 'Response serialization failed'}: ${e.message || e}`,
        details: null,
        stackId: ''
      },
      meta: {
        ms: 0,
        version: API_VERSION
      }
    };
  }
}

function api_stackId_() {
  try {
    return Utilities.getUuid();
  } catch (e) {
    return String(Date.now());
  }
}

function api_requestId_() {
  try {
    return Utilities.getUuid();
  } catch (e) {
    return String(Date.now());
  }
}

function api_classifyError_(err) {
  if (err instanceof ApiError) {
    return {
      code: err.code || 'INTERNAL',
      message: err.message || 'Request failed',
      details: err.details || null
    };
  }

  const message = err && err.message ? err.message : String(err);
  const details = err && err.details ? err.details : null;
  const code = api_guessErrorCode_(message, err);

  return { code, message, details };
}

function api_guessErrorCode_(message, err) {
  if (err && err.code) return err.code;
  const msg = String(message || '').toLowerCase();
  if (msg.includes('access denied') || msg.includes('access control')) return 'FORBIDDEN';
  if (msg.includes('missing sheet')) return 'FAILED_PRECONDITION';
  if (msg.includes('not found')) return 'NOT_FOUND';
  if (msg.includes('invalid') || msg.includes('required')) return 'BAD_REQUEST';
  return 'INTERNAL';
}

function api_attachDetails_(err, details) {
  if (!details) return err;
  if (!err || typeof err !== 'object') return err;
  if (err.details && typeof err.details === 'object') {
    err.details = { ...err.details, ...details };
  } else {
    err.details = details;
  }
  return err;
}

function api_asObject_(value, label) {
  if (value === null || value === undefined) return {};
  if (typeof value !== 'object' || Array.isArray(value)) {
    throw new ApiError('BAD_REQUEST', `${label || 'payload'} must be an object.`);
  }
  return value;
}

function api_sanitizeText_(value, opts) {
  const str = String(value || '');
  const trimmed = opts && opts.trim === false ? str : str.trim();
  const withoutControls = trimmed.replace(/[\u0000-\u001F\u007F]/g, '');
  if (opts && opts.maxLen && withoutControls.length > opts.maxLen) {
    throw new ApiError('BAD_REQUEST', `${opts.label || 'Value'} exceeds ${opts.maxLen} characters.`);
  }
  return withoutControls;
}

function api_requireString_(value, label, opts) {
  const sanitized = api_sanitizeText_(value, { ...opts, label });
  if (!sanitized && !(opts && opts.allowEmpty)) {
    throw new ApiError('BAD_REQUEST', `${label} is required.`);
  }
  if (opts && opts.pattern && sanitized && !opts.pattern.test(sanitized)) {
    throw new ApiError('BAD_REQUEST', `${label} has invalid format.`);
  }
  return sanitized;
}

function api_optionalString_(value, label, opts) {
  const sanitized = api_sanitizeText_(value, { ...opts, label });
  if (!sanitized) return '';
  if (opts && opts.pattern && !opts.pattern.test(sanitized)) {
    throw new ApiError('BAD_REQUEST', `${label} has invalid format.`);
  }
  return sanitized;
}

function api_requireNumber_(value, label, opts) {
  const n = Number(value);
  if (!isFinite(n)) throw new ApiError('BAD_REQUEST', `${label} must be a number.`);
  if (opts && opts.integer && Math.floor(n) !== n) {
    throw new ApiError('BAD_REQUEST', `${label} must be an integer.`);
  }
  if (opts && typeof opts.min === 'number' && n < opts.min) {
    throw new ApiError('BAD_REQUEST', `${label} must be >= ${opts.min}.`);
  }
  if (opts && typeof opts.max === 'number' && n > opts.max) {
    throw new ApiError('BAD_REQUEST', `${label} must be <= ${opts.max}.`);
  }
  return n;
}

function api_optionalNumber_(value, label, opts) {
  if (value === null || value === undefined || value === '') {
    if (opts && Object.prototype.hasOwnProperty.call(opts, 'default')) return opts.default;
    return null;
  }
  return api_requireNumber_(value, label, opts);
}

function api_requireArray_(value, label, opts) {
  if (!Array.isArray(value)) throw new ApiError('BAD_REQUEST', `${label} must be an array.`);
  const maxLen = opts && opts.maxLen ? opts.maxLen : 500;
  if (value.length > maxLen) throw new ApiError('BAD_REQUEST', `${label} exceeds ${maxLen} items.`);
  return value;
}

function api_requireEnum_(value, label, allowed) {
  const sanitized = api_requireString_(value, label, { maxLen: 50 });
  const normalized = sanitized.toUpperCase();
  if (!allowed.includes(normalized)) {
    throw new ApiError('BAD_REQUEST', `${label} must be one of: ${allowed.join(', ')}.`);
  }
  return normalized;
}

function api_requireBoolean_(value, label) {
  if (typeof value === 'boolean') return value;
  if (value === 'true' || value === 'TRUE' || value === 1 || value === '1') return true;
  if (value === 'false' || value === 'FALSE' || value === 0 || value === '0' || value === '' || value === null || value === undefined) return false;
  throw new ApiError('BAD_REQUEST', `${label} must be boolean.`);
}

function api_ping() {
  return api_handleRequest_('api_ping', () => ({ ts: new Date().toISOString() }));
}

function api_bootstrap() {
  return api_handleRequest_('api_bootstrap', () => {
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

    const boolFromCfg = (key, defaultVal) => {
      const v = String(cfg[key] || '').trim().toUpperCase();
      if (!v) return defaultVal;
      return ['TRUE', 'YES', '1', 'Y', 'ON'].includes(v);
    };

    const webAppUrl = (() => {
      try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
    })();

    return {
      user,
      webAppUrl,
      config: {
        namePrefix: cfg.NAME_PREFIX || 'mBOM',
        freezeDefault: boolFromCfg('FREEZE_AGILE_INPUTS_DEFAULT', true),
        requireApprovedForm: boolFromCfg('REQUIRE_APPROVED_FORM', true),
        requireApprovedAgile: boolFromCfg('REQUIRE_APPROVED_AGILE', true),
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
  });
}

function api_loadData(payload) {
  return api_handleRequest_('api_loadData', () => {
    payload = api_asObject_(payload, 'payload');
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
      const limitForms = api_optionalNumber_(payload.limitForms, 'limitForms', { integer: true, min: 0, max: 1000, default: 200 });
      const limitReleases = api_optionalNumber_(payload.limitReleases, 'limitReleases', { integer: true, min: 0, max: 1000, default: 200 });
      const limitAgileLatest = api_optionalNumber_(payload.limitAgileLatest, 'limitAgileLatest', { integer: true, min: 0, max: 1000, default: 300 });
      const jobsLimit = api_optionalNumber_(payload.jobsLimit, 'jobsLimit', { integer: true, min: 0, max: 200, default: 30 });

      diag.limits = { limitForms, limitReleases, limitAgileLatest, jobsLimit };

      const allowedSections = ['projects', 'forms', 'releases', 'agileLatest', 'jobs', 'config', 'notifications', 'approvals'];
      const sections = Array.isArray(payload.sections) ? payload.sections.map(s => String(s || '').trim()).filter(Boolean) : [];
      const filteredSections = sections.filter(s => allowedSections.includes(s));
      const loadAll = filteredSections.length === 0;
      const wantProjects = loadAll || filteredSections.includes('projects');
      const wantForms = loadAll || filteredSections.includes('forms') || filteredSections.includes('approvals');
      const wantReleases = loadAll || filteredSections.includes('releases');
      const wantAgileLatest = loadAll || filteredSections.includes('agileLatest') || filteredSections.includes('approvals');
      const wantJobs = loadAll || filteredSections.includes('jobs');
      const wantConfig = loadAll || filteredSections.includes('config');
      const wantNotifications = loadAll || filteredSections.includes('notifications');

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

      const resume = api_asObject_(payload.resume || {}, 'resume');
      const formsOffset = api_optionalNumber_(resume.formsOffset, 'resume.formsOffset', { integer: true, min: 0, default: 0 });
      const releasesOffset = api_optionalNumber_(resume.releasesOffset, 'resume.releasesOffset', { integer: true, min: 0, default: 0 });
      const agileOffset = api_optionalNumber_(resume.agileOffset, 'resume.agileOffset', { integer: true, min: 0, default: 0 });
      const pendingAgileOffset = api_optionalNumber_(resume.pendingAgileOffset, 'resume.pendingAgileOffset', { integer: true, min: 0, default: 0 });
      const pendingFormsOffset = api_optionalNumber_(resume.pendingFormsOffset, 'resume.pendingFormsOffset', { integer: true, min: 0, default: 0 });

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

      const forms = wantForms ? formsAll.slice(formsOffset, formsOffset + limitForms) : [];
      const releases = wantReleases ? releasesAll.slice(releasesOffset, releasesOffset + limitReleases) : [];

      const agileLatestAll = wantAgileLatest ? (dash.agileLatest || []) : [];
      const agileLatest = wantAgileLatest ? agileLatestAll.slice(agileOffset, agileOffset + limitAgileLatest) : [];

      const pendingFormsAll = loadAll
        ? step('normalize_pendingForms', () => dashboard_normalizeFilesForUi_(dash.pendingForms || []))
        : [];

      const pendingForms = loadAll
        ? pendingFormsAll.slice(pendingFormsOffset, pendingFormsOffset + limitForms)
        : [];

      const pendingAgile = loadAll
        ? (dash.pendingAgile || []).slice(pendingAgileOffset, pendingAgileOffset + limitAgileLatest)
        : [];

      const jobsSummary = wantJobs ? step('jobs_summary_', () => jobs_summary_()) : { summary: { queued: 0, running: 0, done: 0, doneWithErrors: 0, error: 0 }, runningJob: null };
      const jobsRecent = wantJobs ? step('jobs_list_', () => {
        try {
          return jobs_list_({ limit: jobsLimit, activeOnly: false });
        } catch (e) {
          return [];
        }
      }) : [];

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
        formsTotal: formsAll.length,
        formsSent: forms.length,
        releasesTotal: releasesAll.length,
        releasesSent: releases.length,
        agileLatestTotal: agileLatestAll.length,
        agileLatestSent: agileLatest.length,
        pendingFormsTotal: pendingFormsAll.length,
        pendingFormsSent: pendingForms.length,
        pendingAgileTotal: (dash.pendingAgile || []).length,
        pendingAgileSent: pendingAgile.length,
        jobsSent: jobsRecent.length
      };

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

      const page = {
        next: {
          formsOffset: (formsOffset + forms.length < formsAll.length) ? formsOffset + forms.length : null,
          releasesOffset: (releasesOffset + releases.length < releasesAll.length) ? releasesOffset + releases.length : null,
          agileOffset: (agileOffset + agileLatest.length < agileLatestAll.length) ? agileOffset + agileLatest.length : null,
          pendingFormsOffset: (pendingFormsOffset + pendingForms.length < pendingFormsAll.length) ? pendingFormsOffset + pendingForms.length : null,
          pendingAgileOffset: (pendingAgileOffset + pendingAgile.length < (dash.pendingAgile || []).length) ? pendingAgileOffset + pendingAgile.length : null
        }
      };

      return {
        diag,
        page,
        configList,
        dashboard: {
          indexState: dash.indexState,
          projects: dash.projects || [],
          agileLatest,
          pendingAgile,
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
      api_attachDetails_(e, { diag });
      throw e;
    }
  });
}

function api_refreshAgileIndex() {
  return api_handleRequest_('api_refreshAgileIndex', () => {
    auth_requireEditor_();
    return agile_refreshIndex_();
  });
}

function api_refreshFilesIndex() {
  return api_handleRequest_('api_refreshFilesIndex', () => {
    auth_requireEditor_();
    return files_refreshIndexFromDrive_();
  });
}

function api_listAgileTabs(payload) {
  return api_handleRequest_('api_listAgileTabs', () => {
    payload = api_asObject_(payload, 'payload');
    const site = api_requireString_(payload.site, 'site', { maxLen: 100 });
    const part = api_requireString_(payload.part, 'part', { maxLen: 100 });
    return { rows: agile_listTabs_(site, part) };
  });
}

function api_setAgileApproval(payload) {
  return api_handleRequest_('api_setAgileApproval', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');
    const tabName = api_requireString_(payload.tabName, 'tabName', { maxLen: 200 });
    const status = api_requireEnum_(payload.status, 'status', ['APPROVED', 'REJECTED']);
    const notes = api_optionalString_(payload.notes, 'notes', { maxLen: 500 });
    return agile_approval_set_(tabName, status, notes || '');
  });
}

function api_setProjectClusterGroup(payload) {
  return api_handleRequest_('api_setProjectClusterGroup', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');
    const projectKey = api_requireString_(payload.projectKey, 'projectKey', { maxLen: 80, pattern: /^[A-Za-z0-9._-]+$/ });
    const clusterGroup = api_requireNumber_(payload.clusterGroup, 'clusterGroup', { integer: true, min: 1, max: 9 });
    return projects_setClusterGroup_(projectKey, clusterGroup);
  });
}

function api_scheduleFormRevision(payload) {
  return api_handleRequest_('api_scheduleFormRevision', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');

    const baseFormFileId = api_optionalString_(payload.baseFormFileId, 'baseFormFileId', { maxLen: 200 });
    const newFormRev = api_requireNumber_(payload.newFormRev, 'newFormRev', { integer: true, min: 1, max: 1000 });
    const changeRef = api_optionalString_(payload.changeRef || payload.ecrActRef, 'changeRef', { maxLen: 200 });
    const affectedItems = api_optionalString_(payload.affectedItems, 'affectedItems', { maxLen: 500 });
    const description = api_optionalString_(payload.description, 'description', { maxLen: 500 });

    return jobs_create_('CREATE_FORM', {
      baseFormFileId,
      newFormRev,
      changeRef,
      ecrActRef: changeRef,
      affectedItems,
      description
    });
  });
}

function api_scheduleReleasedForProject(payload) {
  return api_handleRequest_('api_scheduleReleasedForProject', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');

    const projectKey = api_requireString_(payload.projectKey, 'projectKey', { maxLen: 80, pattern: /^[A-Za-z0-9._-]+$/ });
    const includeMda = api_requireBoolean_(payload.includeMda, 'includeMda');
    const freezeAgileInputs = api_requireBoolean_(payload.freezeAgileInputs, 'freezeAgileInputs');

    const agileTabCluster = api_requireString_(payload.agileTabCluster, 'agileTabCluster', { maxLen: 200 });
    const agileTabMDA = includeMda
      ? api_requireString_(payload.agileTabMDA, 'agileTabMDA', { maxLen: 200 })
      : '';

    const buswayClusterCode = api_optionalString_(payload.buswayClusterCode, 'buswayClusterCode', { maxLen: 10, pattern: /^[A-Za-z0-9]*$/ });
    const buswayMdaCode = includeMda
      ? api_optionalString_(payload.buswayMdaCode, 'buswayMdaCode', { maxLen: 10, pattern: /^[A-Za-z0-9]*$/ })
      : '';

    const eco = api_optionalString_(payload.eco, 'eco', { maxLen: 120 });
    const affectedItems = api_optionalString_(payload.affectedItems, 'affectedItems', { maxLen: 500 });
    const description = api_optionalString_(payload.description, 'description', { maxLen: 500 });

    return jobs_create_('CREATE_RELEASED_ONE', {
      projectKey,
      includeMda,
      freezeAgileInputs,
      agileTabCluster,
      agileTabMDA,
      buswayClusterCode,
      buswayMdaCode,
      eco,
      affectedItems,
      description
    });
  });
}

function api_scheduleReleasedForSelected(payload) {
  return api_handleRequest_('api_scheduleReleasedForSelected', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');
    const projectKeys = api_requireArray_(payload.projectKeys, 'projectKeys', { maxLen: 200 })
      .map(key => api_requireString_(key, 'projectKey', { maxLen: 80, pattern: /^[A-Za-z0-9._-]+$/ }));

    const batchSize = api_optionalNumber_(payload.batchSize, 'batchSize', { integer: true, min: 1, max: 10, default: 3 });
    const freezeAgileInputs = api_requireBoolean_(payload.freezeAgileInputs, 'freezeAgileInputs');
    const onlyEligible = api_requireBoolean_(payload.onlyEligible, 'onlyEligible');
    const description = api_optionalString_(payload.description, 'description', { maxLen: 500 });
    const eco = api_optionalString_(payload.eco, 'eco', { maxLen: 120 });
    const affectedItems = api_optionalString_(payload.affectedItems, 'affectedItems', { maxLen: 500 });

    return jobs_create_('CREATE_RELEASES_SELECTED', {
      projectKeys,
      batchSize,
      freezeAgileInputs,
      description,
      eco,
      affectedItems,
      onlyEligible
    });
  });
}

function api_scheduleReleasedForAll(payload) {
  return api_handleRequest_('api_scheduleReleasedForAll', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');

    const batchSize = api_optionalNumber_(payload.batchSize, 'batchSize', { integer: true, min: 1, max: 10, default: 3 });
    const freezeAgileInputs = api_requireBoolean_(payload.freezeAgileInputs, 'freezeAgileInputs');
    const onlyEligible = api_requireBoolean_(payload.onlyEligible, 'onlyEligible');
    const description = api_optionalString_(payload.description, 'description', { maxLen: 500 });
    const eco = api_optionalString_(payload.eco, 'eco', { maxLen: 120 });
    const affectedItems = api_optionalString_(payload.affectedItems, 'affectedItems', { maxLen: 500 });

    return jobs_create_('CREATE_RELEASES_ALL', {
      batchSize,
      freezeAgileInputs,
      description,
      eco,
      affectedItems,
      onlyEligible
    });
  });
}

function api_setApprovedFormFileId(fileIdOrPayload) {
  return api_handleRequest_('api_setApprovedFormFileId', () => {
    auth_requireEditor_();
    let fileId = '';
    if (typeof fileIdOrPayload === 'object' && fileIdOrPayload !== null) {
      fileId = api_requireString_(fileIdOrPayload.fileId, 'fileId', { maxLen: 200 });
    } else {
      fileId = api_requireString_(fileIdOrPayload, 'fileId', { maxLen: 200 });
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('CONFIG');
    if (!sh) throw new Error('Missing sheet: CONFIG');

    const rng = sh.getDataRange().getValues();
    for (let i = 1; i < rng.length; i++) {
      if (String(rng[i][0]).trim() === 'CURRENT_APPROVED_FORM_FILE_ID') {
        sh.getRange(i + 1, 2).setValue(String(fileId || '').trim());
        CacheService.getDocumentCache().remove('CFG_ALL');
        return { updated: true };
      }
    }
    sh.appendRow(['CURRENT_APPROVED_FORM_FILE_ID', String(fileId || '').trim()]);
    CacheService.getDocumentCache().remove('CFG_ALL');
    return { updated: true };
  });
}

function api_setFileStatus(payload) {
  return api_handleRequest_('api_setFileStatus', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');
    const fileId = api_requireString_(payload.fileId, 'fileId', { maxLen: 200 });
    const status = api_requireEnum_(payload.status, 'status', ['DRAFT', 'APPROVED', 'OBSOLETE']);
    const ok = files_setStatus_(fileId, status);
    return { updated: ok };
  });
}

function api_updateConfig(payload) {
  return api_handleRequest_('api_updateConfig', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');
    const updates = api_asObject_(payload.updates, 'updates');
    const sanitized = {};
    Object.keys(updates).forEach(key => {
      const safeKey = api_requireString_(key, 'config key', { maxLen: 100, pattern: /^[A-Z0-9_]+$/ });
      const value = api_optionalString_(updates[key], `config ${safeKey}`, { maxLen: 500, trim: false });
      sanitized[safeKey] = value;
    });
    return cfg_update_(sanitized);
  });
}

function api_getProjectData(payload) {
  return api_handleRequest_('api_getProjectData', () => {
    payload = api_asObject_(payload, 'payload');
    const projectKey = api_requireString_(payload.projectKey, 'projectKey', { maxLen: 80, pattern: /^[A-Za-z0-9._-]+$/ });
    return projectdata_get_(projectKey);
  });
}

function api_importReleased(payload) {
  return api_handleRequest_('api_importReleased', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');
    const text = api_optionalString_(payload.text, 'text', { maxLen: 100000, trim: false });
    return files_importReleasedFromText_(text || '');
  });
}

function api_updateNotifSettings(payload) {
  return api_handleRequest_('api_updateNotifSettings', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');
    const updates = api_asObject_(payload.updates, 'updates');
    const sanitized = {};
    Object.keys(updates).forEach(key => {
      const safeKey = api_requireString_(key, 'notification key', { maxLen: 100, pattern: /^[A-Z0-9_]+$/ });
      const value = api_optionalString_(updates[key], `notification ${safeKey}`, { maxLen: 1000, trim: false });
      sanitized[safeKey] = value;
    });
    return notif_updateSettings_(sanitized);
  });
}

function api_listJobs(payload) {
  return api_handleRequest_('api_listJobs', () => {
    payload = api_asObject_(payload, 'payload');
    const limit = api_optionalNumber_(payload.limit, 'limit', { integer: true, min: 1, max: 200, default: 50 });
    const activeOnly = api_requireBoolean_(payload.activeOnly === undefined ? false : payload.activeOnly, 'activeOnly');
    return {
      summary: jobs_summary_(),
      jobs: jobs_list_({ limit, activeOnly })
    };
  });
}

function api_jobStatus(jobId) {
  return api_handleRequest_('api_jobStatus', () => {
    const id = api_requireString_(jobId, 'jobId', { maxLen: 80 });
    const res = jobs_getStatus_(id);
    if (!res || res.ok === false) {
      throw new ApiError('NOT_FOUND', 'Job not found');
    }
    return res.job;
  });
}

function api_retryJob(payload) {
  return api_handleRequest_('api_retryJob', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');
    const jobId = api_requireString_(payload.jobId, 'jobId', { maxLen: 80 });
    const res = jobs_retry_(jobId);
    if (res.ok === false) throw new ApiError('FAILED_PRECONDITION', res.error || 'Job retry failed');
    return res;
  });
}

function api_removeJob(payload) {
  return api_handleRequest_('api_removeJob', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload, 'payload');
    const jobId = api_requireString_(payload.jobId, 'jobId', { maxLen: 80 });
    const res = jobs_remove_(jobId);
    if (res.ok === false) throw new ApiError('FAILED_PRECONDITION', res.error || 'Job removal failed');
    return res;
  });
}

function api_restartJobs(payload) {
  return api_handleRequest_('api_restartJobs', () => {
    auth_requireEditor_();
    payload = api_asObject_(payload || {}, 'payload');
    return jobs_restartRunner_();
  });
}

function api_getArchitectureSpec() {
  return api_handleRequest_('api_getArchitectureSpec', () => ({
    spec: arch_getSpec_(),
    implementationSteps: arch_getImplementationSteps_()
  }));
}
