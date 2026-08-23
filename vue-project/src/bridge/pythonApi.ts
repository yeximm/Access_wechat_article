type PywebviewStatus = {
  ok: boolean
  status: string
  webviewExists: boolean
  serverTime: string
  environment?: {
    appName?: string
    appVersion?: string
    systemLabel?: string
    pythonVersion?: string
    mitmproxyVersion?: string
    playwrightVersion?: string
    pywebviewVersion?: string
  }
}

export type RuntimeLogLevel = 'DEBUG' | 'INFO' | 'WARN' | 'ERROR'

export type ArticleRuntimeStage = {
  stage: string
  label: string
  status: string
  durationSeconds: number
  message: string
  createdAt: string
}

export type ArticleRuntimeRecord = {
  articleKey: string
  taskInfo: string
  status: string
  currentStage: string
  collectComments: boolean
  startedAt: string
  finishedAt: string
  durationSeconds?: number | null
  errorMessage: string
  stages: ArticleRuntimeStage[]
}

export type TaskRuntimeState = {
  currentAction?: string
  accountName?: string
  taskInfo?: string
  proxyStatusLabel?: string
  progressDone?: number
  progressTotalLabel?: string
  progressPercent?: number
  averageArticleSeconds?: number | null
  averageArticleDurationLabel?: string
  activeWorkerCount?: number
  totalWorkerCount?: number
  errorCount?: number
  latestError?: string
  articleRecords?: ArticleRuntimeRecord[]
}

export type TaskStatus = {
  ok: boolean
  status: string
  message?: string
  home?: HomeSnapshot
  runtimeState?: TaskRuntimeState
  auth?: AuthStatus
  hardware?: HardwareStatus
  proxy?: {
    host: string
    port: number
    enabled: boolean
    mitmEnabled?: boolean
    systemProxyEnabled?: boolean
    systemProxyActive?: boolean
    systemProxyReadable?: boolean
    systemProxyServer?: string
    systemProxyReadError?: string
    configuredProxyServer?: string
    mitmPort?: number
    mitmListenHost?: string
    mitmListenAddress?: string
    mitmPortOccupied?: boolean
    mitmPortAvailable?: boolean
    mitmPortOwner?: string
    mitmStartedAt?: number | null
  }
  config?: {
    autoSaveContent: boolean
    autoCleanTempFiles: boolean
    autoStartProxy: boolean
    enableSystemProxy?: boolean
    logLevel?: RuntimeLogLevel
    requestIntervalSeconds?: number
    startupDelaySeconds?: number
    verificationUrl?: string
    values?: Record<string, string>
  }
  workers?: string[]
  runOptions?: TaskRunOptions
  dbPath?: string
  appStartedAt?: string
  uptimeSeconds?: number
}

export type HealthCheckTarget = 'https' | 'ca' | 'proxy-port' | 'storage'

export type HealthCheckItem = {
  key: string
  label: string
  value?: string
  path?: string
  pathAbsolute?: string
  message?: string
  action?: string
  ok?: boolean
  exists?: boolean
  isDirectory?: boolean
  readable?: boolean
  writable?: boolean
}

export type HealthCheckResult = {
  ok: boolean
  target: HealthCheckTarget
  tone: 'success' | 'warning' | 'danger' | 'info'
  label: string
  message: string
  items: HealthCheckItem[]
  details: Record<string, unknown>
  checkedAt: string
}

export type StartupHealthCheckResult = {
  ok: boolean
  status: string
  cached: boolean
  results: Partial<Record<HealthCheckTarget, HealthCheckResult>>
  checkedAt: string
}

export type StartupSelfCheckItem = {
  key: string
  group: string
  label: string
  status: 'success' | 'warning' | 'failed'
  severity: 'info' | 'warning' | 'fatal'
  ok: boolean
  message: string
  action?: string
}

export type StartupSelfCheckState = {
  file_schema_version: number
  checked_version: string
  checked_data_schema_version: string
  checked_at: string
  status: string
  fatal_count: number
  warning_count: number
  duration_seconds: number
  items: StartupSelfCheckItem[]
}

export type StartupSelfCheckStatus = {
  ok: boolean
  needsSelfCheck: boolean
  currentVersion: string
  currentDataSchemaVersion: string
  state: StartupSelfCheckState | null
  statePath: string
}

export type StartupSelfCheckResult = {
  ok: boolean
  status: string
  fatalCount: number
  warningCount: number
  items: StartupSelfCheckItem[]
  checkedAt: string
  durationSeconds: number
  statePath: string
}

export type HardwareHistoryPoint = {
  timestamp: number | string
  cpuPercent?: number
  memoryPercent?: number
  memoryUsedBytes?: number
}

export type HardwareStatus = {
  cpuPercent: number
  memoryPercent: number
  memoryUsedBytes: number
  cpuLabel: string
  memoryLabel: string
  processCount?: number
  historySampleCount?: number
  sampleIntervalSeconds?: number
  history: HardwareHistoryPoint[]
  updatedAt?: string
}

export type HomeSnapshot = {
  status: string
  statusLabel: string
  accountName: string
  description: string
  originalCount: string
  friendFollowCount: string
  found: boolean
  message?: string
}

export type AuthStatus = {
  hasKeyUrl: boolean
  status: string
  statusLabel: string
  lastKeyUrlAt?: string
  lastKeyUrlSource?: string
  lastKeyUrlRedacted?: string
}

export type TaskRunOptions = {
  recordLimit: number
  dateFilterMode: 'all' | 'range' | 'before' | 'after'
  startDate?: string
  endDate?: string
  selections: {
    articleDetail: boolean
    offlineArchive?: boolean
    commentInfo: boolean
    skipCollectedRecords: boolean
    offlineArchiveMode?: 'standard' | 'beta'
  }
}

export type RuntimeConfigPayload = {
  autoSaveContent: boolean
  autoCleanTempFiles: boolean
  autoStartProxy: boolean
  enableSystemProxy: boolean
  logLevel: RuntimeLogLevel
  requestIntervalSeconds: number
  proxy: {
    host: string
    port: number
    startupDelaySeconds: number
    verificationUrl: string
  }
  values?: Record<string, string | number | boolean>
}

export type TaskLogItem = {
  level: string
  message: string
  source: string
  createdAt: string
  channel?: 'system' | 'main_flow' | 'article_task' | string
  phase?: string
  taskIndex?: number | string
  articleTaskId?: string
  articleTitle?: string
  errorId?: string
}

type TaskLogsResult = {
  ok: boolean
  items: TaskLogItem[]
}

type OpenRuntimeLogResult = {
  ok: boolean
  status: string
  logPath?: string
  logDir?: string
  message?: string
}

export type RuntimePathKey = 'projectDir' | 'outputDir' | 'storageDir' | 'logDir'

export type RuntimePathsResult = {
  ok: boolean
  status: string
  paths: Record<RuntimePathKey, string>
  message?: string
}

type OpenRuntimePathResult = {
  ok: boolean
  status: string
  key: RuntimePathKey
  path?: string
  message?: string
}

type SelectRuntimeDirectoryResult = {
  ok: boolean
  status: string
  configKey: string
  initialDir?: string
  path?: string
  selectedPath?: string
  message?: string
  taskStatus?: TaskStatus
}

type OpenArchiveArticleDirectoryResult = {
  ok: boolean
  status: string
  articleId: number
  path?: string
  message?: string
}

type SaveRuntimeConfigResult = {
  ok: boolean
  status: string
  message?: string
  configPath?: string
  taskStatus?: TaskStatus
}

type ResetRuntimeConfigResult = SaveRuntimeConfigResult & {
  backupPath?: string
}

export type CaCertificateStatus = {
  ok: boolean
  status: string
  installed: boolean
  label: string
  message?: string
  subject?: string
  storePath?: string
  thumbprint?: string
  currentCaPath?: string
  currentCaRelativePath?: string
  caFileExists?: boolean
  storeCertificateCount?: number
  projectCertificate?: MitmProjectCertificate | null
  projectCertificateInstalled?: boolean
  certificates?: MitmCertificateItem[]
}

export type MitmProjectCertificate = {
  path?: string
  thumbprint?: string
  subject?: string
  issuer?: string
  friendlyName?: string
  notBefore?: string
  notAfter?: string
}

export type MitmCertificateItem = {
  storePath: string
  thumbprint: string
  subject: string
  issuer?: string
  friendlyName?: string
  notBefore?: string
  notAfter?: string
  matchesProject?: boolean
}

type MitmCertificateListResult = {
  ok: boolean
  status: string
  count: number
  certificates: MitmCertificateItem[]
  message?: string
}

type MitmCertificateDeleteResult = {
  ok: boolean
  status: string
  deletedCount: number
  skippedCount: number
  deleted?: Array<{ thumbprint: string; storePath: string }>
  skipped?: Array<{ thumbprint: string; reason: string }>
  remainingCertificates?: MitmCertificateItem[]
  remainingCertificateCount?: number
  message?: string
}

type InstallCaCertificateResult = CaCertificateStatus

type ClearRuntimeCacheResult = {
  ok: boolean
  status: string
  removedCount: number
  removedFileCount: number
  removedDirectoryCount: number
  freedBytes: number
  skippedCount: number
  skipped?: Array<{ path: string; error: string }>
  tempDir?: string
  message?: string
}

type ProxyConnectionTestResult = {
  ok: boolean
  status: string
  message?: string
  url?: string
  proxy?: string
  statusCode?: number
  bytesRead?: number
}

export type WindowDiagnosticAction =
  | 'read-home'
  | 'activate-home'
  | 'first-article-click'
  | 'scroll-page'
  | 'bounce-scroll'
  | 'close-tab'

export type WindowDiagnosticResultItem = {
  label: string
  value: string
  cells?: Array<{ label: string, value: string }>
  kind?: 'summary' | 'operation' | 'discarded' | 'article'
  tone?: 'info' | 'success' | 'warning' | 'error'
  sequence?: number
}

export type WindowDiagnosticResult = {
  ok: boolean
  status: string
  action: WindowDiagnosticAction
  title?: string
  message?: string
  tone?: 'info' | 'success' | 'warning' | 'error'
  items?: WindowDiagnosticResultItem[]
}

export type WindowDiagnosticOptions = {
  scrollSteps?: number
}

export type WindowClickFlowDiagnosticOptions = {
  maxRecords: number
  dateFilterMode: 'all' | 'range' | 'before' | 'after'
  startDate?: string
  endDate?: string
}

export type ArticleDetailDiagnosticOptions = {
  skipCollectedRecords?: boolean
}

export type InitialContentStorageDiagnosticOptions = {
  skipCollectedRecords?: boolean
  storeArticleDetail?: boolean
}

export type ArticleDetailCommentsDiagnosticOptions = {
  skipCollectedRecords?: boolean
  storeArticleDetail?: boolean
  storeCommentInfo?: boolean
}

export type ArticleDetailOfflineCacheDiagnosticOptions = {
  skipCollectedRecords?: boolean
  storeArticleDetail?: boolean
  archiveOfflineContent?: boolean
  statefulOfflineCache?: boolean
}

export type ArticleDetailDiagnosticResult = {
  ok: boolean
  status: string
  jobId: string
  action:
    | 'single-article-detail'
    | 'initial-content-storage'
    | 'article-detail-comments'
    | 'article-detail-offline-cache'
    | 'window-click-flow'
  title?: string
  message?: string
  tone?: 'info' | 'success' | 'warning' | 'error'
  items?: WindowDiagnosticResultItem[]
  captureType?: 'html' | 'reference' | 'none' | string
  totalSeconds?: number
  htmlSource?: string
  archiveDir?: string
  articleId?: number
  accountId?: number
  historyId?: number
  attemptId?: string
  resourceManifest?: string[]
  commentCount?: number
  replyCount?: number
  commentPageCount?: number
  commentPath?: string
  commentAssetCount?: number
  commentAssetDir?: string
  offlineIndexPath?: string
  offlineResourceCount?: number
  offlineAssetsDir?: string
  offlineWarning?: string
  recognizedCount?: number
  skippedCount?: number
  stoppedByUser?: boolean
  traceDir?: string
  executionLogPath?: string
  resultPath?: string
}

export type ArchiveAccountItem = {
  id: number
  accountName: string
  createdTime: string
  updatedTime: string
  latestCollectTime: string
  articleCount: number
  savedCount: number
  failedCount: number
  sizeLabel: string
}

export type ArchiveArticleItem = {
  id: number
  accountId: number
  title: string
  publishedArticleTime: string
  articleLink: string
  recordType: string
  collectTime: string
  durationSeconds: number
  collectStatus: string
  statusLabel: string
  archiveDir: string
  archiveDirs: string[]
  sizeBytes: number
  sizeLabel: string
}

export type ArchiveSummary = {
  ok: boolean
  status: string
  accountCount: number
  articleCount: number
  dataType: string
  storageSizeBytes: number
  storageSizeLabel: string
  storageRoot: string
  dbPath: string
  message?: string
}

export type ArchiveCacheJobStatus = 'pending' | 'running' | 'done' | 'partial_failed' | 'failed' | 'missing'

export type ArchiveCacheResultItem = {
  articleId: number
  articleTitle: string
  ok: boolean
  status: string
  message: string
  archiveDir: string
  indexHtmlPath: string
  resourceCount: number
  warning: string
  elapsedSeconds: number
}

export type ArchiveCacheActiveProcess = {
  articleId: number
  articleTitle: string
  status: string
  step: string
  elapsedSeconds: number
}

export type ArchiveCacheJob = {
  ok: boolean
  jobId: string
  status: ArchiveCacheJobStatus
  total: number
  finished: number
  running: number
  skipped: number
  requestedTotal: number
  processed: number
  queued: number
  failed: number
  activeProcesses: ArchiveCacheActiveProcess[]
  concurrency: number
  results: ArchiveCacheResultItem[]
  message?: string
}

export type HistoryRecordItem = {
  id: number
  articleId: number
  accountId: number
  name: string
  account: string
  collectType: string
  collectTime: string
  recordTime: string
  startedTime: string
  finishedTime: string
  duration: string
  durationSeconds: number
  collectStatus: string
  status: string
  articleLink: string
  publishedArticleTime: string
  resourceTypes: string[]
  resourceTypeLabels: string[]
  errorStage: string
  errorStageLabel: string
  errorMessage: string
  outputDir: string
  recordSummary: HistoryRecordSummary
}

export type HistoryRecordSummaryItem = {
  key: string
  label: string
  value: string
}

export type HistoryRecordSummary = {
  kind: 'metrics' | 'status' | 'missing'
  items: HistoryRecordSummaryItem[]
  message: string
}

export type HistoryRecordsQuery = {
  page?: number
  pageSize?: number
  keyword?: string
  collectType?: string
  status?: string
  collectDate?: string
  collectStartDate?: string
  collectEndDate?: string
}

export type HistoryRecordsResult = {
  ok: boolean
  status: string
  page: number
  pageSize: number
  items: HistoryRecordItem[]
  total: number
  dbPath: string
  message?: string
}

export type HistorySuggestionsQuery = {
  keyword?: string
  limit?: number
}

export type HistorySuggestionsResult = {
  ok: boolean
  status: string
  items: string[]
  total: number
  dbPath?: string
  message?: string
}

export type HistoryTrendItem = {
  date: string
  label: string
  count: number
}

export type HistorySummary = {
  ok: boolean
  status: string
  totalRecords: number
  successfulRecords: number
  savedRecords?: number
  failedRecords: number
  successRate: number
  latestCollectDate: string
  collectedArticleCount: number
  averageDuration: string
  trend: HistoryTrendItem[]
  dbPath?: string
  message?: string
}

export type HistoryClearResult = {
  ok: boolean
  status: string
  deletedCount: number
  dbPath?: string
  message?: string
}

type ArchiveAccountsResult = {
  ok: boolean
  status: string
  items: ArchiveAccountItem[]
  total: number
  dbPath: string
  message?: string
}

type ArchiveAccountArticlesResult = {
  ok: boolean
  status: string
  accountId: number
  page: number
  pageSize: number
  items: ArchiveArticleItem[]
  total: number
  dbPath: string
  message?: string
}

type ArchiveDeleteResult = {
  ok: boolean
  status: string
  deletedArticleCount: number
  deletedAccountCount: number
  deletedArchiveDirCount: number
  missingArticleIds: number[]
  failures: Array<{ path: string; error: string }>
  message?: string
}

export type ArchiveExcelExportFile = {
  accountId: number
  accountName: string
  rowCount: number
  tempPath: string
  outputPath: string
  fileName: string
}

export type ArchiveExcelExportResult = {
  ok: boolean
  status: string
  exportedFileCount: number
  totalRowCount: number
  files: ArchiveExcelExportFile[]
  missingAccountIds: number[]
  message?: string
}

async function requestHttpApi<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(path, {
    ...init,
    headers: {
      'Content-Type': 'application/json',
      ...(init?.headers ?? {}),
    },
  })

  const payload = await response.json() as T
  if (!response.ok) {
    const message = typeof payload === 'object' && payload && 'message' in payload
      ? String((payload as { message?: unknown }).message)
      : `HTTP API 调用失败：${response.status}`
    throw new Error(message)
  }

  return payload
}

function postJson<T>(path: string, payload?: unknown): Promise<T> {
  return requestHttpApi<T>(path, {
    method: 'POST',
    body: JSON.stringify(payload ?? {}),
  })
}

function getJson<T>(path: string): Promise<T> {
  return requestHttpApi<T>(path, {
    method: 'GET',
  })
}

function parsePywebviewPayload<T>(payload: unknown): T {
  if (typeof payload === 'string') {
    return JSON.parse(payload) as T
  }
  return payload as T
}

// 统一封装 Vue 到 Python 的调用；业务接口统一走 FastAPI。
export async function getPythonStatus() {
  return requestHttpApi<PywebviewStatus>('/api/status')
}

export async function runStartupHealthChecks() {
  return postJson<StartupHealthCheckResult>('/api/health/startup')
}

export async function getStartupSelfCheckStatus() {
  return getJson<StartupSelfCheckStatus>('/api/startup-self-check/status')
}

export async function runStartupSelfCheck() {
  return postJson<StartupSelfCheckResult>('/api/startup-self-check/run')
}

export async function checkHealthTarget(target: HealthCheckTarget) {
  return postJson<HealthCheckResult>('/api/health/check', { target })
}

export async function startTask(options?: TaskRunOptions) {
  return postJson<TaskStatus>('/api/task/start', options)
}

export async function stopTask() {
  return postJson<TaskStatus>('/api/task/stop')
}

export async function startMitmProxy() {
  return postJson<TaskStatus>('/api/proxy/mitm/start')
}

export async function stopMitmProxy() {
  return postJson<TaskStatus>('/api/proxy/mitm/stop')
}

export async function enableSystemProxy() {
  return postJson<TaskStatus>('/api/proxy/system/enable')
}

export async function disableSystemProxy() {
  return postJson<TaskStatus>('/api/proxy/system/disable')
}

export async function getTaskStatus() {
  return requestHttpApi<TaskStatus>('/api/task/status')
}

export async function getHardwareStatus() {
  return requestHttpApi<HardwareStatus>('/api/runtime/hardware')
}

export async function getTaskLogs(limit = 100) {
  return requestHttpApi<TaskLogsResult>(`/api/task/logs?limit=${encodeURIComponent(limit)}`)
}

export async function openCurrentRuntimeLog() {
  return requestHttpApi<OpenRuntimeLogResult>('/api/log/current/open')
}

export async function listRuntimePaths() {
  return requestHttpApi<RuntimePathsResult>('/api/runtime/paths')
}

export async function openRuntimePath(key: RuntimePathKey) {
  return postJson<OpenRuntimePathResult>('/api/runtime/paths/open', { key })
}

export async function openArchiveArticleDirectory(articleId: number) {
  return postJson<OpenArchiveArticleDirectoryResult>('/api/archive/articles/open-directory', { articleId })
}

export async function saveRuntimeConfig(configPayload: RuntimeConfigPayload) {
  return postJson<SaveRuntimeConfigResult>('/api/config/save', configPayload)
}

export async function updateRuntimeConfig(configPayload: RuntimeConfigPayload) {
  return postJson<SaveRuntimeConfigResult>('/api/config/update', configPayload)
}

export async function resetRuntimeConfig() {
  return postJson<ResetRuntimeConfigResult>('/api/config/reset', {})
}

export async function selectRuntimeDirectory(configKey: string, currentPath: string) {
  return postJson<SelectRuntimeDirectoryResult>('/api/runtime/paths/select-directory', {
    configKey,
    currentPath,
  })
}

export async function checkCaCertificate() {
  return requestHttpApi<CaCertificateStatus>('/api/ca/status')
}

export async function installCaCertificate() {
  return postJson<InstallCaCertificateResult>('/api/ca/install')
}

export async function listMitmCaCertificates() {
  return requestHttpApi<MitmCertificateListResult>('/api/ca/mitm/list')
}

export async function deleteMitmCaCertificates(thumbprints: string[]) {
  return postJson<MitmCertificateDeleteResult>('/api/ca/mitm/delete', { thumbprints })
}

export async function clearRuntimeCache() {
  return postJson<ClearRuntimeCacheResult>('/api/cache/runtime/clear')
}

export async function testProxyConnection() {
  return postJson<ProxyConnectionTestResult>('/api/proxy/test')
}

export async function runWindowDiagnosticAction(
  action: WindowDiagnosticAction,
  options: WindowDiagnosticOptions = {},
) {
  return postJson<WindowDiagnosticResult>('/api/diagnostics/window', { action, ...options })
}

export async function startWindowClickFlowDiagnostic(options: WindowClickFlowDiagnosticOptions) {
  return postJson<ArticleDetailDiagnosticResult>('/api/diagnostics/window-click-flow', options)
}

export async function getWindowClickFlowDiagnostic(jobId: string) {
  return requestHttpApi<ArticleDetailDiagnosticResult>(
    '/api/diagnostics/window-click-flow/' + encodeURIComponent(jobId),
  )
}

export async function stopWindowClickFlowDiagnostic(jobId: string) {
  return postJson<ArticleDetailDiagnosticResult>(
    '/api/diagnostics/window-click-flow/' + encodeURIComponent(jobId) + '/stop',
    {},
  )
}

export async function startArticleDetailDiagnostic(
  options: ArticleDetailDiagnosticOptions = {},
) {
  return postJson<ArticleDetailDiagnosticResult>('/api/diagnostics/article-detail', options)
}

export async function getArticleDetailDiagnostic(jobId: string) {
  return requestHttpApi<ArticleDetailDiagnosticResult>(
    '/api/diagnostics/article-detail/' + encodeURIComponent(jobId),
  )
}

export async function startInitialContentStorageDiagnostic(
  options: InitialContentStorageDiagnosticOptions = {},
) {
  return postJson<ArticleDetailDiagnosticResult>('/api/diagnostics/initial-content-storage', options)
}

export async function getInitialContentStorageDiagnostic(jobId: string) {
  return requestHttpApi<ArticleDetailDiagnosticResult>(
    '/api/diagnostics/initial-content-storage/' + encodeURIComponent(jobId),
  )
}

export async function startArticleDetailCommentsDiagnostic(
  options: ArticleDetailCommentsDiagnosticOptions = {},
) {
  return postJson<ArticleDetailDiagnosticResult>('/api/diagnostics/article-detail-comments', options)
}

export async function getArticleDetailCommentsDiagnostic(jobId: string) {
  return requestHttpApi<ArticleDetailDiagnosticResult>(
    '/api/diagnostics/article-detail-comments/' + encodeURIComponent(jobId),
  )
}

export async function startArticleDetailOfflineCacheDiagnostic(
  options: ArticleDetailOfflineCacheDiagnosticOptions = {},
) {
  return postJson<ArticleDetailDiagnosticResult>('/api/diagnostics/article-detail-offline-cache', options)
}

export async function getArticleDetailOfflineCacheDiagnostic(jobId: string) {
  return requestHttpApi<ArticleDetailDiagnosticResult>(
    '/api/diagnostics/article-detail-offline-cache/' + encodeURIComponent(jobId),
  )
}

export async function listArchiveAccounts() {
  return requestHttpApi<ArchiveAccountsResult>('/api/archive/accounts')
}

export async function listArchiveSummary() {
  return requestHttpApi<ArchiveSummary>('/api/archive/summary')
}

export async function getHistoryRecords(query: HistoryRecordsQuery = {}) {
  const params = new URLSearchParams({
    page: String(query.page ?? 1),
    pageSize: String(query.pageSize ?? 15),
  })
  for (const [key, value] of Object.entries({
    keyword: query.keyword,
    collectType: query.collectType,
    status: query.status,
    collectDate: query.collectDate,
    collectStartDate: query.collectStartDate,
    collectEndDate: query.collectEndDate,
  })) {
    if (value) {
      params.set(key, String(value))
    }
  }
  return requestHttpApi<HistoryRecordsResult>(`/api/history/records?${params.toString()}`)
}

export async function getHistorySuggestions(query: HistorySuggestionsQuery = {}) {
  const params = new URLSearchParams()
  if (query.keyword) {
    params.set('keyword', query.keyword)
  }
  params.set('limit', String(query.limit ?? 20))
  return requestHttpApi<HistorySuggestionsResult>(`/api/history/suggestions?${params.toString()}`)
}

export async function getHistorySummary() {
  return requestHttpApi<HistorySummary>('/api/history/summary')
}

export async function clearHistoryRecords() {
  return requestHttpApi<HistoryClearResult>('/api/history/records', { method: 'DELETE' })
}

export async function listArchiveAccountArticles(accountId: number, page = 1, pageSize = 10) {
  const query = new URLSearchParams({
    page: String(page),
    pageSize: String(pageSize),
  })
  return requestHttpApi<ArchiveAccountArticlesResult>(
    `/api/archive/accounts/${encodeURIComponent(accountId)}/articles?${query.toString()}`,
  )
}

export async function deleteArchiveArticles(articleIds: number[]) {
  return requestHttpApi<ArchiveDeleteResult>('/api/archive/articles', {
    method: 'DELETE',
    body: JSON.stringify({ articleIds }),
  })
}

export async function deleteArchiveAccount(accountId: number) {
  return requestHttpApi<ArchiveDeleteResult>(
    `/api/archive/accounts/${encodeURIComponent(accountId)}`,
    { method: 'DELETE' },
  )
}

export async function deleteArchiveAll() {
  return requestHttpApi<ArchiveDeleteResult>('/api/archive', { method: 'DELETE' })
}

export async function exportArchiveAccountsToExcel(accountIds: number[]) {
  return postJson<ArchiveExcelExportResult>('/api/archive/export/accounts', { accountIds })
}

export async function cacheArchiveArticles(articleIds: number[]) {
  return postJson<ArchiveCacheJob>('/api/archive/cache/articles', { articleIds })
}

export async function cacheArchiveAccount(accountId: number) {
  return postJson<ArchiveCacheJob>(`/api/archive/accounts/${encodeURIComponent(accountId)}/cache`)
}

export async function getArchiveCacheJob(jobId: string) {
  return requestHttpApi<ArchiveCacheJob>(`/api/archive/cache/jobs/${encodeURIComponent(jobId)}`)
}
