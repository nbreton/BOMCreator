const JOB_PREFIX = 'JOB_';
const JOB_MAX_KEEP = 200;
const JOB_STALE_MINUTES = 10;

/**
 * Create a new job and schedule execution.
 */
function jobs_create_(type, params) {
  const user = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '');
  const id = `${Date.now()}_${Math.floor(Math.random() * 1e6)}`;
  const job = {
    id,
    type,
    status: 'QUEUED',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    createdBy: user,
    startedAt: '',
    finishedAt: '',
    message: 'Queued',
    progressCurrent: 0,
    progressTotal: 0,
    cursor: 0,
    params: params || {},
    results: [],
    errors: [],
    retryCount: 0,
    cancelRequested: false
  };
  jobs_put_(job);
  jobs_cleanupOld_();
  jobs_schedule_();
  return { ok: true, jobId: id };
}

// -----------------------
// Storage helpers
// -----------------------
function jobs_props_() { return PropertiesService.getScriptProperties(); }

function jobs_key_(id) {
  return `${JOB_PREFIX}${id}`;
}

function jobs_put_(job) {
  job.updatedAt = new Date().toISOString();
  jobs_props_().setProperty(jobs_key_(job.id), JSON.stringify(job));
}

function jobs_get_(jobId) {
  const raw = jobs_props_().getProperty(jobs_key_(jobId));
  return raw ? JSON.parse(raw) : null;
}

function jobs_delete_(jobId) {
  jobs_props_().deleteProperty(jobs_key_(jobId));
  return { ok: true };
}

function jobs_remove_(jobId) {
  const job = jobs_get_(jobId);
  if (!job) return { ok: false, error: 'Job not found' };
  if (job.status === 'RUNNING') {
    return { ok: false, error: 'Cannot remove a running job. Retry later.' };
  }
  jobs_delete_(jobId);
  return { ok: true };
}

function jobs_cancel_(jobId) {
  const job = jobs_get_(jobId);
  if (!job) return { ok: false, error: 'Job not found' };
  if (['DONE', 'DONE_WITH_ERRORS', 'ERROR', 'CANCELLED'].includes(job.status)) {
    return { ok: false, error: `Cannot cancel job in status: ${job.status}` };
  }

  if (job.status === 'QUEUED') {
    job.status = 'CANCELLED';
    job.finishedAt = new Date().toISOString();
    job.message = 'Cancelled by user.';
  } else if (job.status === 'RUNNING') {
    job.cancelRequested = true;
    job.message = 'Cancel requested…';
  }
  jobs_put_(job);
  return { ok: true, jobId: job.id, status: job.status };
}

function jobs_restartRunner_() {
  jobs_cleanupTriggers_();
  jobs_schedule_();
  return { ok: true };
}

// -----------------------
// Runner + scheduler
// -----------------------
function jobs_run_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    jobs_schedule_();
    return;
  }

  try {
    jobs_markStale_();

    const job = jobs_getNext_();
    if (!job) {
      jobs_cleanupTriggers_();
      return;
    }

    if (job.cancelRequested) {
      job.status = 'CANCELLED';
      job.finishedAt = new Date().toISOString();
      job.message = 'Cancelled by user.';
      jobs_put_(job);
      jobs_cleanupTriggers_();
      return;
    }

    if (job.status === 'QUEUED') {
      job.status = 'RUNNING';
      job.startedAt = job.startedAt || new Date().toISOString();
      job.message = 'Starting…';
      jobs_put_(job);
    }

    // Execute one batch/step
    jobs_execute_(job);

    jobs_put_(job);

    // Continue if still running or queued work remains
    if (job.status === 'RUNNING') {
      jobs_schedule_();
    } else {
      const nextJob = jobs_getNext_();
      if (nextJob) {
        jobs_schedule_();
      } else {
        jobs_cleanupTriggers_();
      }
    }

  } catch (e) {
    log_error_('Job runner error', { message: e.message, stack: e.stack });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Executes a job. This is where your existing mBOM operations are invoked.
 * Ensure MbomOps.gs functions exist: mbom_createNewFormRevision_ and mbom_createReleasedForProject_.
 */
function jobs_execute_(job) {
  switch (job.type) {
    case 'CREATE_FORM':
      return jobs_execCreateForm_(job);

    case 'CREATE_RELEASED_ONE':
      return jobs_execCreateReleasedOne_(job);

    case 'CREATE_RELEASES_SELECTED':
      return jobs_execCreateReleasedBatch_(job);

    case 'CREATE_RELEASES_ALL':
      return jobs_prepareAllThenRun_(job);

    default:
      job.status = 'ERROR';
      job.finishedAt = new Date().toISOString();
      job.message = `Unknown job type: ${job.type}`;
      job.errors.push({ error: job.message });
      return;
  }
}

function jobs_execCreateForm_(job) {
  try {
    if (job.cancelRequested) {
      job.status = 'CANCELLED';
      job.finishedAt = new Date().toISOString();
      job.message = 'Cancelled by user.';
      return;
    }
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
    if (job.cancelRequested) {
      job.status = 'CANCELLED';
      job.finishedAt = new Date().toISOString();
      job.message = 'Cancelled by user.';
      return;
    }
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

/**
 * Convert ALL into explicit project list once, then run as SELECTED.
 */
function jobs_prepareAllThenRun_(job) {
  job.params = job.params || {};
  if (!Array.isArray(job.params.projectKeys)) {
    const allProjects = agile_getProjects_(); // requires AgileIndex.gs function
    const keys = allProjects.map(p => p.projectKey);

    if (job.params.onlyEligible === true) {
      const filtered = [];
      for (const pk of keys) {
        const eff = projects_getEffective_(pk);
        const includeMda = eff.includeMda;

        const site = pk.split('-')[0];
        const zone = pk.split('-')[1];

        const cl = agile_getLatestTab_(site, `Zone ${zone}`);
        if (!cl || !cl.TabName) continue;
        if (agile_approval_status_(cl.TabName) !== 'APPROVED') continue;

        if (includeMda) {
          const m = agile_getLatestTab_(site, 'MDA');
          if (!m || !m.TabName) continue;
          if (agile_approval_status_(m.TabName) !== 'APPROVED') continue;
        }
        filtered.push(pk);
      }
      job.params.projectKeys = filtered;
    } else {
      job.params.projectKeys = keys;
    }
  }

  job.type = 'CREATE_RELEASES_SELECTED';
  return jobs_execCreateReleasedBatch_(job);
}

function jobs_execCreateReleasedBatch_(job) {
  const projectKeys = (job.params && Array.isArray(job.params.projectKeys)) ? job.params.projectKeys : [];
  const batchSize = Number((job.params && job.params.batchSize) || 3);

  if (!job.progressTotal) job.progressTotal = projectKeys.length;

  // Edge case: empty list
  if (!projectKeys.length) {
    job.status = 'DONE';
    job.finishedAt = new Date().toISOString();
    job.message = 'No projects to process.';
    job.progressCurrent = 0;
    return;
  }

  const end = Math.min(projectKeys.length, job.cursor + batchSize);
  job.message = `Creating RELEASED: ${job.cursor + 1}–${end} of ${projectKeys.length}`;
  jobs_put_(job);

  for (let i = job.cursor; i < end; i++) {
    if (job.cancelRequested) {
      job.status = 'CANCELLED';
      job.finishedAt = new Date().toISOString();
      job.message = 'Cancelled by user.';
      return;
    }
    const pk = projectKeys[i];
    try {
      const eff = projects_getEffective_(pk);
      const includeMda = (job.params.includeMdaOverride === true)
        ? true
        : (job.params.includeMdaOverride === false ? false : eff.includeMda);

      const site = pk.split('-')[0];
      const zone = pk.split('-')[1];

      const cl = agile_getLatestTab_(site, `Zone ${zone}`);
      if (!cl || !cl.TabName) throw new Error(`No latest Cluster Agile for ${pk}`);

      const m = includeMda ? agile_getLatestTab_(site, 'MDA') : null;
      if (includeMda && (!m || !m.TabName)) throw new Error(`MDA required but missing for ${pk}`);

      // Infer busway codes (job-level overrides win)
      const clusterSupplier = eff.clusterBuswaySupplier || String(cl.BuswaySupplier || '');
      const mdaSupplier = includeMda ? (eff.mdaBuswaySupplier || String(m.BuswaySupplier || '')) : '';

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
// Listing, status, retry
// -----------------------
function jobs_publicView_(job) {
  return {
    id: job.id,
    type: job.type,
    status: job.status,
    message: job.message || '',
    createdAt: job.createdAt,
    updatedAt: job.updatedAt,
    createdBy: job.createdBy,
    startedAt: job.startedAt,
    finishedAt: job.finishedAt,
    progressCurrent: job.progressCurrent || 0,
    progressTotal: job.progressTotal || 0,
    cursor: job.cursor || 0,
    resultsCount: (job.results || []).length,
    errorsCount: (job.errors || []).length,
    results: (job.results || []).slice(0, 10),
    errors: (job.errors || []).slice(0, 10)
  };
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

function jobs_getStatus_(jobId) {
  const job = jobs_get_(jobId);
  if (!job) return { ok: false, error: 'Job not found' };
  return { ok: true, job: jobs_publicView_(job) };
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

function jobs_retry_(jobId) {
  const job = jobs_get_(jobId);
  if (!job) return { ok: false, error: 'Job not found' };
  if (!['ERROR', 'DONE_WITH_ERRORS'].includes(job.status)) {
    return { ok: false, error: `Cannot retry job in status: ${job.status}` };
  }

  if (job.type === 'CREATE_RELEASES_SELECTED' && Array.isArray(job.errors) && job.errors.length) {
    const failedKeys = job.errors.map(e => e.projectKey).filter(Boolean);
    if (failedKeys.length) {
      job.params = job.params || {};
      job.params.projectKeys = failedKeys;
    }
  }

  job.status = 'QUEUED';
  job.startedAt = '';
  job.finishedAt = '';
  job.message = `Retrying (attempt ${Number(job.retryCount || 0) + 1})`;
  job.progressCurrent = 0;
  job.progressTotal = 0;
  job.cursor = 0;
  job.results = [];
  job.errors = [];
  job.retryCount = Number(job.retryCount || 0) + 1;
  job.cancelRequested = false;
  jobs_put_(job);
  jobs_schedule_();
  return { ok: true, jobId: job.id };
}

// -----------------------
// Maintenance: stale jobs + cleanup
// -----------------------
function jobs_getNext_() {
  const props = jobs_props_().getProperties();
  const keys = Object.keys(props).filter(k => k.startsWith(JOB_PREFIX));
  const jobs = keys.map(k => JSON.parse(props[k]));

  jobs.sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt)));
  return jobs.find(j => j.status === 'QUEUED' || j.status === 'RUNNING') || null;
}

function jobs_markStale_() {
  const props = jobs_props_().getProperties();
  const keys = Object.keys(props).filter(k => k.startsWith(JOB_PREFIX));
  if (!keys.length) return;

  const now = Date.now();
  const staleMs = JOB_STALE_MINUTES * 60 * 1000;

  keys.forEach(k => {
    const job = JSON.parse(props[k]);
    if (job.status !== 'RUNNING') return;

    const ts = Date.parse(job.updatedAt || job.startedAt || job.createdAt || '');
    if (!isFinite(ts)) return;
    if (now - ts < staleMs) return;

    job.status = 'ERROR';
    job.finishedAt = new Date().toISOString();
    job.message = `Marked stale after ${JOB_STALE_MINUTES} minutes without updates.`;
    job.errors = job.errors || [];
    job.errors.push({ error: job.message });
    jobs_put_(job);
  });
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
