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
