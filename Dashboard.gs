function dashboard_build_(opts) {
  opts = opts || {};
  const includeProjects = opts.includeProjects !== false;
  const includeFormsList = opts.includeFormsList !== false;
  const includeReleasesList = opts.includeReleasesList !== false;
  const includeAgileLatest = opts.includeAgileLatest !== false;
  const includePending = opts.includePending !== false;
  const includeAgileAll = opts.includeAgileAll === true;
  const includeLatestApprovedForm = (opts.includeLatestApprovedForm !== undefined)
    ? opts.includeLatestApprovedForm
    : true;

  const indexState = agile_indexState_();

  const needAgileRows = includeProjects || includeAgileLatest || includePending || includeAgileAll;
  const agileRows = needAgileRows ? agile_readIndex_() : [];
  const rawProjects = includeProjects ? agile_getProjectsFromRows_(agileRows) : [];
  const agileLatest = (includeAgileLatest || includePending) ? agile_listLatestFromRows_(agileRows) : [];
  const agileAll = includeAgileAll ? dashboard_listAgileAllFromRows_(agileRows) : [];

  const needFiles = includeProjects || includeFormsList || includeReleasesList || includeLatestApprovedForm || includePending;
  const filesAll = needFiles ? files_listAll_() : [];
  const formsAll = (includeFormsList || includeLatestApprovedForm || includePending)
    ? filesAll.filter(f => f.Type === 'FORM')
    : [];
  const releasesAll = (includeReleasesList || includeProjects)
    ? filesAll.filter(f => f.Type === 'RELEASED')
    : [];

  if (includeProjects) {
    projects_syncFromAgile_(rawProjects);
  }

  const projects = includeProjects ? rawProjects.map(p => {
    const effective = projects_getEffective_(p.projectKey);
    return {
      ...p,
      clusterGroup: effective.clusterGroup,
      includeMda: effective.includeMda,
      includeMdaOverride: effective.includeMdaOverride || '',
      clusterBuswaySupplier: effective.clusterBuswaySupplier || p.clusterBuswaySupplier || '',
      mdaBuswaySupplier: effective.mdaBuswaySupplier || p.mdaBuswaySupplier || ''
    };
  }) : [];
  const agileTabUrls = includeProjects ? dashboard_buildAgileTabUrlMap_(
    projects.flatMap(p => [p.clusterTab, p.mdaTab])
  ) : {};

  const latestReleaseByProject = {};
  if (includeProjects) {
    for (const r of releasesAll) {
      const pk = String(r.ProjectKey || '').trim();
      if (!pk) continue;
      const cur = latestReleaseByProject[pk];
      const rev = Number(r.MbomRev || 0);
      if (!cur || rev > Number(cur.MbomRev || 0)) latestReleaseByProject[pk] = r;
    }
  }

  const projectsView = includeProjects ? projects.map(p => {
    const rel = latestReleaseByProject[p.projectKey] || null;
    return {
      ...p,
      clusterTabUrl: agileTabUrls[p.clusterTab] || '',
      mdaTabUrl: agileTabUrls[p.mdaTab] || '',
      hasNewAgileRevision: dashboard_hasNewAgileRevision_(p, rel),
      latestReleased: rel ? {
        mbomRev: rel.MbomRev,
        status: String(rel.Status || ''),
        url: String(rel.Url || ''),
        fileId: String(rel.FileId || ''),
        fileName: String(rel.FileName || ''),
        agileTabCluster: String(rel.AgileTabCluster || ''),
        agileTabMDA: String(rel.AgileTabMDA || '')
      } : null
    };
  }) : [];

  const approvedForms = includeLatestApprovedForm ? formsAll
    .filter(f => String(f.Status || '').toUpperCase() === 'APPROVED')
    .sort((a, b) => Number(b.MbomRev || 0) - Number(a.MbomRev || 0)) : [];
  const latestApprovedForm = approvedForms[0] || null;

  const pendingAgile = includePending
    ? agileLatest.filter(a => String(a.approvalStatus || '').toUpperCase() !== 'APPROVED')
    : [];
  const pendingForms = includePending
    ? formsAll.filter(f => String(f.Status || '').toUpperCase() !== 'APPROVED')
    : [];

  return {
    indexState,
    projects: projectsView,
    forms: includeFormsList ? formsAll : [],
    releases: includeReleasesList ? releasesAll : [],
    agileLatest,
    agileAll,
    latestApprovedForm,
    pendingAgile,
    pendingForms
  };
}

function dashboard_listAgileAllFromRows_(rows) {
  return (rows || []).map(r => ({
    site: String(r.Site || ''),
    partNorm: String(r.PartNorm || ''),
    rev: r.Rev,
    tabName: String(r.TabName || ''),
    downloadDate: String(r.DownloadDate || ''),
    eco: String(r.ECO || ''),
    approvalStatus: String(r.ApprovalStatus || 'PENDING'),
    buswaySupplier: String(r.BuswaySupplier || ''),
    isLatest: agile_isTrue_(r.IsLatest)
  }));
}

function dashboard_buildAgileTabUrlMap_(tabNames) {
  const names = Array.from(new Set((tabNames || []).map(n => String(n || '').trim()).filter(Boolean)));
  if (!names.length) return {};

  const sourceId = cfg_get_('DOWNLOAD_LIST_SS_ID');
  if (!sourceId) return {};

  const ss = SpreadsheetApp.openById(sourceId);
  const lookup = new Set(names);
  const out = {};

  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (!lookup.has(name)) return;
    out[name] = `https://docs.google.com/spreadsheets/d/${sourceId}/edit#gid=${sh.getSheetId()}`;
  });

  return out;
}

function dashboard_hasNewAgileRevision_(project, release) {
  if (!release) return false;

  const projectRevNum = Number(project?.clusterRev || 0);
  const releaseRevNum = Number(release?.AgileRevCluster || 0);

  if (isFinite(projectRevNum) && isFinite(releaseRevNum) && releaseRevNum > 0) {
    if (projectRevNum > releaseRevNum) return true;
  } else {
    const projectRev = String(project?.clusterRev || '').trim();
    const releaseRev = String(release?.AgileRevCluster || '').trim();
    if (projectRev && releaseRev && projectRev !== releaseRev) return true;
  }

  const latestClusterTab = String(project?.clusterTab || '').trim();
  const latestMdaTab = String(project?.mdaTab || '').trim();
  const releaseClusterTab = String(release?.AgileTabCluster || '').trim();
  const releaseMdaTab = String(release?.AgileTabMDA || '').trim();

  if (latestClusterTab && releaseClusterTab && latestClusterTab !== releaseClusterTab) return true;
  if (project?.includeMda && latestMdaTab && releaseMdaTab && latestMdaTab !== releaseMdaTab) return true;
  if (project?.includeMda && latestMdaTab && !releaseMdaTab) return true;

  return false;
}

function dashboard_normalizeFilesForUi_(rows) {
  return (rows || []).map(r => ({
    type: String(r.Type || ''),
    projectKey: String(r.ProjectKey || ''),
    mbomRev: r.MbomRev,
    status: String(r.Status || ''),
    fileId: String(r.FileId || ''),
    url: String(r.Url || ''),
    fileName: String(r.FileName || ''),
    eco: String(r.ECO || ''),
    description: String(r.Description || ''),
    createdBy: String(r.CreatedBy || ''),
    createdAt: (r.CreatedAt instanceof Date) ? r.CreatedAt.toISOString() : String(r.CreatedAt || '')
  }));
}
