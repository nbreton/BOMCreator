const ARCH_SPEC_VERSION = '2024.09.0';

function arch_getSpec_() {
  return {
    version: ARCH_SPEC_VERSION,
    layout: [
      { file: 'Config.gs', purpose: 'Config access, caching, access control.' },
      { file: 'Api.gs', purpose: 'Server API endpoints and request validation.' },
      { file: 'Architecture.gs', purpose: 'Architecture spec and implementation guidance.' },
      { file: 'AgileIndex.gs', purpose: 'Agile index refresh/read + project aggregation.' },
      { file: 'AgileApprovals.gs', purpose: 'Agile approval storage + helpers.' },
      { file: 'AgileReviews.gs', purpose: 'Agile BOM review generation + comparison helpers.' },
      { file: 'FilesDb.gs', purpose: 'FILES sheet read/write and revision helpers.' },
      { file: 'ProjectsDb.gs', purpose: 'PROJECTS sheet read/write and rules.' },
      { file: 'FilesIndex.gs', purpose: 'Drive indexing into FILES.' },
      { file: 'MbomOps.gs', purpose: 'Form and RELEASED creation workflows.' },
      { file: 'Jobs.gs', purpose: 'Background job orchestration.' },
      { file: 'Notifications.gs', purpose: 'Email configuration + senders.' },
      { file: 'Dashboard.gs', purpose: 'Dashboard aggregations for UI.' },
      { file: 'ProjectData.gs', purpose: 'Per-project data for UI.' },
      { file: 'Import.gs', purpose: 'Legacy RELEASED import.' },
      { file: 'DriveFacade.gs', purpose: 'Drive operations with retries.' },
      { file: 'Log.gs', purpose: 'Structured logging.' },
      { file: 'code.gs', purpose: 'Menu hooks + UI entry points.' },
      { file: 'Index.html', purpose: 'App shell + UI logic.' },
      { file: 'style.html', purpose: 'UI styling.' }
    ],
    dataModel: {
      sheets: [
        {
          name: 'CONFIG',
          key: 'Key (string)',
          columns: ['Key', 'Value'],
          validations: ['Key is required', 'Value may be blank', 'Admin-controlled keys are validated in Config.gs']
        },
        {
          name: 'AGILE_INDEX',
          key: 'Site + PartNorm + Rev',
          columns: [
            'Site', 'Part', 'PartNorm', 'ProjectKey',
            'TlaRef', 'Description', 'BuswaySupplier',
            'Rev', 'DownloadDate', 'TabName', 'ECO',
            'IsLatest', 'ApprovalStatus'
          ],
          validations: [
            'Generated from download list; refreshed by AgileIndex.gs',
            'IsLatest is boolean-like',
            'ApprovalStatus defaults to PENDING if not set'
          ]
        },
        {
          name: 'AGILE_APPROVALS',
          key: 'TabName',
          columns: ['TabName', 'Status', 'UpdatedAt', 'UpdatedBy', 'Notes'],
          validations: [
            'Status must be APPROVED or REJECTED',
            'UpdatedAt auto-filled on change'
          ]
        },
        {
          name: 'AGILE_REVIEWS',
          key: 'TabName',
          columns: [
            'TabName', 'Site', 'Part', 'PartNorm', 'ProjectKey', 'Rev', 'DownloadDate',
            'ProjectType', 'ReviewStatus', 'ReviewedAt', 'ReviewedBy', 'SummaryJson', 'ExceptionsJson', 'Notes'
          ],
          validations: [
            'ReviewStatus in {PENDING, APPROVED, REJECTED}',
            'SummaryJson/ExceptionsJson store review details'
          ]
        },
        {
          name: 'FILES',
          key: 'FileId',
          columns: [
            'Type', 'ProjectKey', 'MbomRev', 'BaseFormRev', 'AgileTabMDA', 'AgileTabCluster',
            'AgileRevCluster', 'ECO', 'Description', 'FileId', 'Url',
            'CreatedAt', 'CreatedBy', 'Status', 'Notes'
          ],
          validations: [
            'Type in {FORM, RELEASED}',
            'ProjectKey is GLOBAL for Form rows',
            'MbomRev is integer',
            'Status in {DRAFT, APPROVED, RELEASED, OBSOLETE}'
          ]
        },
        {
          name: 'PROJECTS',
          key: 'ProjectKey',
          columns: ['ProjectKey', 'ClusterGroup', 'IncludeMDAOverride', 'Notes', 'UpdatedAt', 'UpdatedBy'],
          validations: [
            'ClusterGroup is positive integer',
            'IncludeMDAOverride in {TRUE, FALSE, blank}'
          ]
        },
        {
          name: 'NOTIF_CONFIG',
          key: 'Key',
          columns: ['Key', 'Value', 'Description'],
          validations: ['Key required; values used by Notifications.gs']
        },
        {
          name: 'NOTIF_LOG',
          key: 'Timestamp',
          columns: ['Timestamp', 'Type', 'To', 'Subject', 'Details(JSON)'],
          validations: ['Append-only logging']
        },
        {
          name: 'LOGS',
          key: 'Timestamp',
          columns: ['Timestamp', 'Level', 'User', 'Message', 'Data(JSON)'],
          validations: ['Optional; controlled by CONFIG.LOGS_SHEET_ENABLED']
        }
      ]
    },
    apiSurface: [
      { name: 'api_bootstrap', request: null, response: 'user, webAppUrl, config, indexState, jobsSummary' },
      { name: 'api_loadData', request: '{sections[], limitForms, limitReleases, limitAgileLatest, jobsLimit}', response: 'projects/forms/releases/agileLatest/jobs/config/notifications/approvals' },
      { name: 'api_refreshAgileIndex', request: null, response: '{ok, count}' },
      { name: 'api_refreshFilesIndex', request: null, response: '{ok, inserted, updated}' },
      { name: 'api_listAgileTabs', request: '{site, part}', response: 'tab list with approvalStatus' },
      { name: 'api_setAgileApproval', request: '{tabName, status, notes}', response: '{ok, tabName, status}' },
      { name: 'api_setAgileReviewStatus', request: '{tabName, status, notes}', response: '{ok, tabName, status}' },
      { name: 'api_backfillAgileReviews', request: '{includeHistory}', response: '{ok, created, updated}' },
      { name: 'api_setProjectClusterGroup', request: '{projectKey, clusterGroup}', response: '{ok, projectKey, clusterGroup}' },
      { name: 'api_scheduleFormRevision', request: '{baseFormFileId?, newFormRev, changeRef, description, affectedItems}', response: '{ok, jobId}' },
      { name: 'api_scheduleReleasedForProject', request: '{projectKey, releaseRev?, eco?, description?, includeMda?, agileTabCluster?, agileTabMDA?, freezeAgileInputs?}', response: '{ok, jobId}' },
      { name: 'api_scheduleReleasedForSelected', request: '{projectKeys[], onlyEligible?}', response: '{ok, jobId}' },
      { name: 'api_scheduleReleasedForAll', request: '{onlyEligible?}', response: '{ok, jobId}' },
      { name: 'api_setApprovedFormFileId', request: '{fileId}', response: '{ok}' },
      { name: 'api_setFileStatus', request: '{fileId, status}', response: '{ok}' },
      { name: 'api_updateConfig', request: '{updates}', response: '{ok, updated}' },
      { name: 'api_getProjectData', request: '{projectKey}', response: 'project data + released history' },
      { name: 'api_importReleased', request: '{text}', response: '{ok, imported, skipped, errors[]}' },
      { name: 'api_updateNotifSettings', request: '{updates}', response: '{ok, updated}' },
      { name: 'api_listJobs', request: '{limit?}', response: '{jobs[]}' },
      { name: 'api_jobStatus', request: '{jobId}', response: '{job}' },
      { name: 'api_retryJob', request: '{jobId}', response: '{ok}' },
      { name: 'api_removeJob', request: '{jobId}', response: '{ok}' },
      { name: 'api_restartJobs', request: null, response: '{ok}' },
      { name: 'api_getArchitectureSpec', request: null, response: 'architecture spec + implementation steps' }
    ],
    uiIa: {
      screens: [
        { name: 'Dashboard', components: ['Jobs summary', 'Projects table', 'Forms table', 'Agile latest', 'Approvals queue'] },
        { name: 'Create mBOM Form', components: ['Form metadata', 'Schedule job'] },
        { name: 'Create RELEASED', components: ['Project selection', 'Agile tab selection', 'Validation', 'Schedule job'] },
        { name: 'Copy Monitoring', components: ['Job list', 'Progress details'] },
        { name: 'Project Data', components: ['Project detail', 'Released history'] },
        { name: 'Notifications', components: ['Settings grid', 'Send toggles'] },
        { name: 'Settings', components: ['CONFIG editor'] },
        { name: 'Import', components: ['Legacy RELEASED import'] },
        { name: 'README', components: ['In-app usage guide'] }
      ],
      navigation: {
        primaryTabs: ['dashboard', 'createForm', 'createReleased', 'monitoring', 'projectData', 'notifications', 'settings', 'import', 'readme'],
        secondaryTabs: ['projects', 'forms', 'agile', 'approvals']
      }
    },
    performancePlan: {
      batching: [
        'Use batch Jobs for Form/RELEASED creation to avoid timeouts.',
        'Use api_loadData sections to avoid loading entire data sets.'
      ],
      caching: [
        'CONFIG cached in document cache (Config.gs).',
        'FILES and AGILE_INDEX cached in memory per invocation.'
      ],
      indexing: [
        'AGILE_INDEX is a denormalized index refreshed from download list.',
        'FILES is indexed by FileId for upsert and status updates.',
        'PROJECTS syncs keys from AGILE_INDEX.'
      ],
      bottlenecks: [
        'Drive copy operations (Form/RELEASED) are the slowest path.',
        'Agile index refresh depends on the source download list size.',
        'Jobs scheduling uses script properties; keep JOB_MAX_KEEP low.'
      ]
    }
  };
}

function arch_getImplementationSteps_() {
  return [
    'Create/verify AGILE_APPROVALS sheet and wire Agile approval helpers.',
    'Expose architecture spec via API for UI/automation use.',
    'Keep data model sheets aligned with column order in FilesDb.gs/ProjectsDb.gs/Notifications.gs.',
    'Use api_loadData sectioning to limit payload size for dashboards.',
    'Monitor job queue size and adjust JOB_MAX_KEEP if needed.'
  ];
}
