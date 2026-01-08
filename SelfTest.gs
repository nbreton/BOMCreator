function self_test() {
  return api_handleRequest_('self_test', () => {
    const checks = [];
    const ss = SpreadsheetApp.getActive();
    const requiredSheets = ['CONFIG', 'FILES', 'PROJECTS', 'AGILE_INDEX'];

    requiredSheets.forEach(name => {
      const exists = !!ss.getSheetByName(name);
      checks.push({
        name: `sheet:${name}`,
        ok: exists,
        detail: exists ? 'ok' : 'missing'
      });
    });

    let cfgOk = true;
    try { cfg_getAll_(); } catch (e) { cfgOk = false; }
    checks.push({
      name: 'config:read',
      ok: cfgOk,
      detail: cfgOk ? 'ok' : 'failed'
    });

    const requiredFns = ['api_bootstrap', 'api_loadData', 'api_listJobs', 'api_scheduleFormRevision', 'api_scheduleReleasedForProject'];
    requiredFns.forEach(fn => {
      const exists = typeof globalThis[fn] === 'function';
      checks.push({
        name: `fn:${fn}`,
        ok: exists,
        detail: exists ? 'ok' : 'missing'
      });
    });

    const overallOk = checks.every(c => c.ok);

    return {
      overallOk,
      checks
    };
  });
}
