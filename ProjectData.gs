function projectdata_get_(projectKey) {
  const pk = String(projectKey || '').trim();
  if (!pk) throw new Error('Missing projectKey');

  const eff = projects_getEffective_(pk);
  const site = pk.split('-')[0];
  const zone = pk.split('-')[1];
  const includeMda = eff.includeMda;
  const clusterPart = `Zone ${zone}`;

  const clusterAgile = agile_listTabs_(site, clusterPart);
  const mdaAgile = includeMda ? agile_listTabs_(site, 'MDA') : [];

  const releasesAll = files_list_('RELEASED');
  const released = releasesAll
    .filter(r => String(r.ProjectKey || '') === pk)
    .sort((a, b) => Number(b.MbomRev || 0) - Number(a.MbomRev || 0))
    .map(r => ({
      mbomRev: r.MbomRev,
      status: String(r.Status || ''),
      url: String(r.Url || ''),
      fileId: String(r.FileId || ''),
      fileName: String(r.FileName || ''),
      agileTabCluster: String(r.AgileTabCluster || ''),
      agileTabMDA: String(r.AgileTabMDA || ''),
      createdAt: (r.CreatedAt instanceof Date) ? r.CreatedAt.toISOString() : String(r.CreatedAt || '')
    }));

  return {
    projectKey: pk,
    site,
    zone,
    clusterGroup: eff.clusterGroup,
    includeMda,
    clusterAgile,
    mdaAgile,
    released
  };
}

