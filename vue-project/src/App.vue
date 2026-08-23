<script setup lang="ts">
import AppIcon from './components/AppIcon.vue'
import { DownOutlined, QuestionCircleOutlined } from '@ant-design/icons-vue'
import { theme } from 'ant-design-vue'
import type { ThemeConfig } from 'ant-design-vue/es/config-provider/context'
import zhCN from 'ant-design-vue/es/locale/zh_CN'
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { DynamicScroller, DynamicScrollerItem, type DynamicScrollerExposed } from 'vue-virtual-scroller'
import 'vue-virtual-scroller/dist/vue-virtual-scroller.css'
import AppTopbar from './components/AppTopbar.vue'
import TopbarHealthDialog from './components/TopbarHealthDialog.vue'
import HardwareSparkline from './components/HardwareSparkline.vue'
import DataFilesPage from './pages/DataFilesPage.vue'
import HistoryPage from './pages/HistoryPage.vue'
import SettingsPage from './pages/SettingsPage.vue'
import {
  getPythonStatus,
  getHardwareStatus,
  getTaskLogs,
  getTaskStatus,
  startTask,
  stopTask,
  checkHealthTarget,
  getStartupSelfCheckStatus,
  listArchiveSummary,
  openRuntimePath,
  runStartupHealthChecks,
  runStartupSelfCheck,
  type ArchiveSummary,
  type HealthCheckResult,
  type StartupSelfCheckResult,
  type TaskLogItem,
  type TaskRunOptions,
  type TaskRuntimeState,
  type TaskStatus,
  type HardwareStatus,
} from './bridge/pythonApi'
import {
  getBrowserPreviewEnvironmentStatus,
  getEnvironmentErrorStatus,
  INITIAL_ENVIRONMENT_STATUS,
  resolveDesktopEnvironmentStatus,
  type ResolvedDesktopEnvironmentStatus,
} from './utils/desktopStatus'
import { formatLogMessageSegments, type LogMessageSegment } from './utils/logMessage'
import {
  canResumeLogAutoFollow,
  isLogScrollNearBottom,
  LOG_AUTO_FOLLOW_IDLE_MS,
  readLogScrollMetricsFromEvent,
} from './utils/logScroll'
import { buildArchiveSummaryStats } from './utils/archiveSummaryStats'
import { resolveTaskProgressSummary } from './utils/taskProgress'

type Tone = 'blue' | 'green' | 'red' | 'purple' | 'orange'
type TopbarHealthTone = 'success' | 'warning' | 'danger' | 'info'
type HealthCheckTarget = 'https' | 'ca' | 'proxy-port' | 'storage'
type TopbarHealthState = {
  value: string
  statusIcon: string
  tone: TopbarHealthTone
  message: string
  items: HealthCheckResult['items']
  checkedAt: string
  checking: boolean
}
type PageKey = 'home' | 'files' | 'history' | 'settings'
type NavItem = {
  label: string
  icon: string
  page: PageKey | null
  href?: string
}
type RuntimeLogLevel = 'INFO' | 'SUCCESS' | 'WARN' | 'ERROR'
type LogFilterLevel = 'ALL' | RuntimeLogLevel
type TaskDateFilterMode = 'all' | 'range' | 'before' | 'after'
type OfflineArchiveMode = 'standard' | 'beta'
type TaskDateRangeValue = [string, string] | null
type LogDisplayRow = {
  id: string
  time: string
  level: RuntimeLogLevel
  levelLabel: string
  levelClass: string
  message: string
  messageSegments: LogMessageSegment[]
  source: string
  channel: string
  phase: string
  taskIndex: number | string
  articleTaskId: string
  articleTitle: string
}
type RefreshablePageExpose = {
  refreshOnActivated?: () => void | Promise<void>
}

const isDark = ref(false)
const antThemeConfig = computed<ThemeConfig>(() => ({
  algorithm: isDark.value ? theme.darkAlgorithm : theme.defaultAlgorithm,
  token: {
    colorPrimary: '#357fd9',
    borderRadius: 8,
    controlHeight: 32,
    fontFamily: "'HarmonyOS Sans SC', 'Microsoft YaHei', sans-serif",
  },
}))
const activePage = ref<PageKey>('home')
const dataFilesPageRef = ref<RefreshablePageExpose | null>(null)
const historyPageRef = ref<RefreshablePageExpose | null>(null)
const settingsPageRef = ref<RefreshablePageExpose | null>(null)
const systemLogScrollerRef = ref<DynamicScrollerExposed<LogDisplayRow> | null>(null)
const articleTaskLogScrollerRef = ref<DynamicScrollerExposed<LogDisplayRow> | null>(null)
const shouldStickLogToBottom = ref(true)
const lastLogScrollUserInteractionAt = ref(0)
const githubUrl = 'https://github.com/yeximm/Access_wechat_article'
const quickStartUrl = 'https://github.com/yeximm/Access_wechat_article/blob/main/doc/quick_start.md'
const MAX_DESKTOP_STATUS_RETRIES = 12
const DESKTOP_STATUS_RETRY_DELAY_MS = 400
const LOG_POLL_LIMIT = 100
const HARDWARE_POLL_INTERVAL_MS = 500
const taskDateFilterMode = ref<TaskDateFilterMode>('all')
const unlimitedDateTaskCount = ref<number | null>(1)
const dateRangeTaskCount = ref<number | null>(0)
const latestDateTaskCount = ref<number | null>(0)
const earliestDateTaskCount = ref<number | null>(1)
const taskCountSnapshots: Record<TaskDateFilterMode, typeof unlimitedDateTaskCount> = {
  all: unlimitedDateTaskCount,
  range: dateRangeTaskCount,
  before: latestDateTaskCount,
  after: earliestDateTaskCount,
}
// 输入框只展示当前筛选模式的数量，切换模式时保留各自上一次输入值。
const pageCount = computed<number | null>({
  get: () => taskCountSnapshots[taskDateFilterMode.value].value,
  set: (value) => {
    taskCountSnapshots[taskDateFilterMode.value].value = value
  },
})
const taskDateFilterOpen = ref(false)
const taskStartDate = ref('')
const taskEndDate = ref('')
const taskDateRangeValue = computed<TaskDateRangeValue>({
  get: () => {
    if (!taskStartDate.value || !taskEndDate.value) {
      return null
    }
    return [taskStartDate.value, taskEndDate.value] as [string, string]
  },
  set: (value: TaskDateRangeValue) => {
    // Ant Design 范围选择器使用数组，页面内部继续保留独立的起止日期字符串。
    taskStartDate.value = value?.[0] ?? ''
    taskEndDate.value = value?.[1] ?? ''
  },
})
const taskDateFilterOptions: { label: string; value: TaskDateFilterMode }[] = [
  { label: '不限日期', value: 'all' },
  { label: '指定任务日期范围', value: 'range' },
  { label: '截止日期 (不早于)', value: 'before' },
  { label: '起始日期 (不晚于)', value: 'after' },
]
const taskDateFilterLabel = computed(() =>
  taskDateFilterOptions.find((option) => option.value === taskDateFilterMode.value)?.label ?? '不限日期',
)
const TASK_DATE_FILTER_HINTS: Record<TaskDateFilterMode, string> = {
  all: '从当前主页开始不限制文章发布时间',
  range: '采集起始日期至截止日期内发布的文章',
  before: '采集从当前主页开始到指定日期内发布的文章',
  after: '采集指定日期之前发布的文章',
}
const taskDateFilterHint = computed(() => TASK_DATE_FILTER_HINTS[taskDateFilterMode.value])
const desktopStatusLabel = ref('检测中')
const environmentStatus = ref({ ...INITIAL_ENVIRONMENT_STATUS })
const defaultHardwareStatus: HardwareStatus = {
  cpuPercent: 0,
  memoryPercent: 0,
  memoryUsedBytes: 0,
  cpuLabel: '0.0%',
  memoryLabel: '0.0%',
  processCount: 0,
  historySampleCount: 120,
  sampleIntervalSeconds: 0.5,
  history: [],
  updatedAt: '',
}
const defaultRuntimeState: TaskRuntimeState = {
  currentAction: '点击开始运行后，将从桌面主页窗口读取',
  accountName: '等待识别',
  taskInfo: '待获取',
  proxyStatusLabel: '空闲',
  progressDone: 0,
  progressTotalLabel: '全部',
  progressPercent: 0,
  averageArticleSeconds: null,
  averageArticleDurationLabel: '待统计',
  activeWorkerCount: 0,
  totalWorkerCount: 0,
  errorCount: 0,
  latestError: '',
  articleRecords: [],
}
const taskStatus = ref<TaskStatus>({
  ok: false,
  status: 'idle',
  proxy: { host: '127.0.0.1', port: 18000, enabled: false },
  hardware: defaultHardwareStatus,
  runtimeState: defaultRuntimeState,
  workers: [],
  home: {
    status: 'pending',
    statusLabel: '待准备',
    accountName: '等待识别，请先打开微信 PC 端公众号主页',
    description: '点击开始运行后，将从桌面主页窗口读取',
    originalCount: '待获取',
    friendFollowCount: '待获取',
    found: false,
  },
})
const taskLogs = ref<TaskLogItem[]>([])
const frontendRuntimeLogs = ref<TaskLogItem[]>([])
const archiveSummary = ref<ArchiveSummary | null>(null)
const activeLogLevel = ref<LogFilterLevel>('ALL')
const uptimeSeconds = ref(0)
const homeRuntimeState = ref<'idle' | 'detecting' | 'running'>('idle')
const statusValueTooltipKey = ref('')
const downloadSelections = ref({
  articleDetail: true,
  offlineArchive: false,
  commentInfo: true,
  skipCollectedRecords: true,
})
const offlineArchiveMode = ref<OfflineArchiveMode>('standard')
const offlineArchiveModeOpen = ref(false)
const mainTaskSelectionDefaultsApplied = ref(false)
const mainTaskSelectionDefaultsSignature = ref('')
let desktopStatusRetryTimer: number | undefined
let desktopStatusRetryCount = 0
let taskPollingTimer: number | undefined
let hardwarePollingTimer: number | undefined
let hardwareStatusErrorReported = false
let uptimeTimer: number | undefined
let logAutoFollowTimer: number | undefined
let startupHealthCheckRequested = false
let startupSelfCheckRequested = false
let pageRefreshSequence = 0

// 静态数据先用于还原可视化原型，后续接入 Python 后端后替换为真实运行状态。
const navGroups: { title: string; items: NavItem[] }[] = [
  {
    title: '数据采集',
    items: [
      { label: '主服务', icon: 'fa-solid fa-house', page: 'home' },
      { label: '快速开始', icon: 'fa-solid fa-book-open', page: null, href: quickStartUrl },
    ],
  },
  {
    title: '系统管理',
    items: [
      { label: '数据档案', icon: 'fa-regular fa-folder-open', page: 'files' },
      { label: '采集历史', icon: 'fa-regular fa-clock', page: 'history' },
      { label: '系统配置', icon: 'fa-solid fa-gear', page: 'settings' },
    ],
  },
]

const homeSnapshot = computed(() => taskStatus.value.home ?? {
  status: 'pending',
  statusLabel: '待准备',
  accountName: '等待识别，请先打开微信 PC 端公众号主页',
  description: '点击开始运行后，将从桌面主页窗口读取',
  originalCount: '待获取',
  friendFollowCount: '待获取',
  found: false,
})

const normalizedPageCount = computed(() => {
  const count = Number(pageCount.value)
  return Number.isFinite(count) ? Math.max(0, Math.floor(count)) : 1
})

const taskDateFieldLabel = computed(() => {
  const labels: Record<TaskDateFilterMode, string> = {
    all: '指定时间',
    range: '日期范围',
    before: '截止日期',
    after: '起始日期',
  }

  return labels[taskDateFilterMode.value]
})

function selectTaskDateFilterMode(mode: TaskDateFilterMode) {
  taskDateFilterMode.value = mode
  taskDateFilterOpen.value = false
}

const taskSettingsLocked = computed(() =>
  ['starting', 'running', 'stopping', 'cancelling'].includes(taskStatus.value.status),
)

const taskStatusLabel = computed(() => {
  if (homeRuntimeState.value === 'detecting') {
    return '检测主页窗口中'
  }

  if (taskStatus.value.status === 'idle') {
    return homeSnapshot.value.statusLabel || '待准备'
  }

  const labels: Record<string, string> = {
    idle: '待机',
    starting: '启动中',
    running: '采集中',
    stopping: '停止中',
    stopped: '已停止',
    completed: '已完成',
    success: '已完成',
    failed: '异常',
    cancelled: '已停止',
    error: '异常',
    'browser-preview': '浏览器预览',
  }

  return labels[taskStatus.value.status] ?? taskStatus.value.status
})

const taskStatusTone = computed<Tone>(() => {
  if (homeRuntimeState.value === 'detecting') {
    return 'blue'
  }

  if (['starting', 'running'].includes(taskStatus.value.status)) {
    return 'green'
  }

  if (['success', 'completed'].includes(taskStatus.value.status)) {
    return 'green'
  }

  if (homeSnapshot.value.status === 'ready') {
    return 'blue'
  }

  if (['not_found', 'failed', 'dependency_missing', 'content_unreadable'].includes(homeSnapshot.value.status)) {
    return 'red'
  }

  if (taskStatus.value.status === 'error' || taskStatus.value.status === 'failed') {
    return 'red'
  }

  return 'orange'
})

const HEALTH_CHECK_TARGETS: HealthCheckTarget[] = ['storage', 'ca', 'proxy-port', 'https']
const HEALTH_CHECK_META: Record<HealthCheckTarget, { label: string; icon: string }> = {
  https: { label: 'HTTPS 状态', icon: 'fa-solid fa-lock' },
  ca: { label: 'CA 证书', icon: 'fa-solid fa-shield-halved' },
  'proxy-port': { label: '代理端口', icon: 'fa-solid fa-network-wired' },
  storage: { label: '数据目录', icon: 'fa-solid fa-database' },
}

function buildTopbarCheckingState(message = '正在检测，请稍候...'): TopbarHealthState {
  return {
    value: '检测中',
    statusIcon: 'fa-solid fa-rotate',
    tone: 'warning',
    message,
    items: [],
    checkedAt: '',
    checking: true,
  }
}

const topbarHealthStates = ref<Record<HealthCheckTarget, TopbarHealthState>>({
  https: buildTopbarCheckingState('正在验证 HTTPS 代理链路...'),
  ca: buildTopbarCheckingState('正在比对项目证书与系统证书指纹...'),
  'proxy-port': buildTopbarCheckingState('正在检测代理端口占用状态...'),
  storage: buildTopbarCheckingState('正在检测配置目录读写权限...'),
})
const healthDialogVisible = ref(false)
const healthDialogTarget = ref<HealthCheckTarget>('https')
const healthCheckInProgress = ref<HealthCheckTarget | null>(null)
const startupHealthCheckRunning = ref(false)
const healthDialogMeta = computed(() => HEALTH_CHECK_META[healthDialogTarget.value])
const healthDialogState = computed(() => topbarHealthStates.value[healthDialogTarget.value])
const startupSelfCheckDialogVisible = ref(false)
const startupSelfCheckDialogState = ref<TopbarHealthState>(buildTopbarCheckingState('正在自检，请稍候...'))

function normalizeTopbarHealthValue(result: HealthCheckResult) {
  // 顶部空间有限，只显示短状态；完整原因仍保留在弹窗 message 和 items 中。
  if (result.ok) {
    return '正常'
  }

  return '异常'
}

function applyHealthCheckResult(result: HealthCheckResult) {
  topbarHealthStates.value[result.target] = {
    value: normalizeTopbarHealthValue(result),
    statusIcon: result.ok
      ? 'fa-solid fa-circle-check'
      : 'fa-solid fa-triangle-exclamation',
    tone: result.tone,
    message: result.message,
    items: result.items,
    checkedAt: result.checkedAt,
    checking: false,
  }
}

function applyHealthCheckFailure(target: HealthCheckTarget, message: string) {
  topbarHealthStates.value[target] = {
    value: '失败',
    statusIcon: 'fa-solid fa-triangle-exclamation',
    tone: 'danger',
    message,
    items: [],
    checkedAt: new Date().toISOString().slice(0, 19),
    checking: false,
  }
}

const topbarHealthItems = computed(() => [
  {
    key: 'https' as const,
    label: 'HTTPS 状态',
    value: topbarHealthStates.value.https.value,
    icon: 'fa-solid fa-lock',
    statusIcon: topbarHealthStates.value.https.statusIcon,
    tone: topbarHealthStates.value.https.tone,
    checking: topbarHealthStates.value.https.checking,
    disabled: startupHealthCheckRunning.value || healthCheckInProgress.value !== null,
  },
  {
    key: 'ca' as const,
    label: 'CA 证书',
    value: topbarHealthStates.value.ca.value,
    icon: 'fa-solid fa-shield-halved',
    statusIcon: topbarHealthStates.value.ca.statusIcon,
    tone: topbarHealthStates.value.ca.tone,
    checking: topbarHealthStates.value.ca.checking,
    disabled: startupHealthCheckRunning.value || healthCheckInProgress.value !== null,
  },
  {
    key: 'proxy-port' as const,
    label: '代理端口',
    value: topbarHealthStates.value['proxy-port'].value,
    icon: 'fa-solid fa-network-wired',
    statusIcon: topbarHealthStates.value['proxy-port'].statusIcon,
    tone: topbarHealthStates.value['proxy-port'].tone,
    checking: topbarHealthStates.value['proxy-port'].checking,
    disabled: startupHealthCheckRunning.value || healthCheckInProgress.value !== null,
  },
  {
    key: 'storage' as const,
    label: '数据目录',
    value: topbarHealthStates.value.storage.value,
    icon: 'fa-solid fa-database',
    statusIcon: topbarHealthStates.value.storage.statusIcon,
    tone: topbarHealthStates.value.storage.tone,
    checking: topbarHealthStates.value.storage.checking,
    disabled: startupHealthCheckRunning.value || healthCheckInProgress.value !== null,
  },
])

const hardwareStatus = computed(() => taskStatus.value.hardware ?? defaultHardwareStatus)
const hardwareCpuLabel = computed(() => hardwareStatus.value.cpuLabel || `${normalizePercent(hardwareStatus.value.cpuPercent).toFixed(1)}%`)
const hardwareMemoryLabel = computed(() => hardwareStatus.value.memoryLabel || `${normalizePercent(hardwareStatus.value.memoryPercent).toFixed(1)}%`)
const hardwareMemoryUsedLabel = computed(() => formatMemoryBytes(hardwareStatus.value.memoryUsedBytes))
const hardwareMemoryTooltip = computed(() => {
  const processCount = Number(hardwareStatus.value.processCount ?? 0)
  const processLabel = Number.isFinite(processCount) ? Math.max(0, Math.floor(processCount)) : 0
  return `当前软件占用 ${hardwareMemoryUsedLabel.value}，监控进程 ${processLabel} 个`
})
const hardwareHistory = computed(() => hardwareStatus.value.history ?? [])

function normalizePercent(value: unknown) {
  const percent = Number(value)
  if (!Number.isFinite(percent)) {
    return 0
  }
  return Math.min(100, Math.max(0, percent))
}

function formatMemoryBytes(value: unknown) {
  const bytes = Number(value)
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return '0 B'
  }

  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  let normalized = bytes
  let unitIndex = 0
  while (normalized >= 1024 && unitIndex < units.length - 1) {
    normalized /= 1024
    unitIndex += 1
  }

  const decimals = unitIndex === 0 ? 0 : normalized >= 100 ? 0 : 1
  return `${normalized.toFixed(decimals)} ${units[unitIndex]}`
}

const runtimeState = computed<TaskRuntimeState>(() => ({
  ...defaultRuntimeState,
  ...(taskStatus.value.runtimeState ?? {}),
}))

const taskProgressSummary = computed(() => resolveTaskProgressSummary({
  plannedCount: normalizedPageCount.value,
  logs: taskLogs.value,
}))
const taskProgressLabel = computed(() => (
  `${runtimeState.value.progressDone ?? taskProgressSummary.value.completedCount}/${runtimeState.value.progressTotalLabel || taskProgressSummary.value.totalLabel}`
))

const logTabs = [
  { level: 'ALL', label: 'ALL' },
  { level: 'INFO', label: 'INFO' },
  { level: 'SUCCESS', label: 'SUCCESS' },
  { level: 'WARN', label: 'WARN' },
  { level: 'ERROR', label: 'ERROR' },
] satisfies { level: LogFilterLevel; label: string }[]
const logSegmentOptions = logTabs.map((tab) => ({
  label: tab.label,
  value: tab.level,
}))

const logLevelLabels = {
  INFO: 'INFO',
  SUCCESS: 'SUCCESS',
  WARN: 'WARN',
  ERROR: 'ERROR',
} satisfies Record<RuntimeLogLevel, string>

const allLogRows = computed<LogDisplayRow[]>(() => {
  const mergedLogs = [...taskLogs.value, ...frontendRuntimeLogs.value]
    .sort((left, right) => String(left.createdAt || '').localeCompare(String(right.createdAt || '')))
  const rows = mergedLogs.map((item, index) => {
    const level = normalizeLogLevel(item.level)
    const channel = normalizeLogChannel(item.channel)

    return {
      id: `${item.createdAt || 'unknown'}-${item.source || 'runtime'}-${index}-${item.message}`,
      time: formatLogTime(item.createdAt),
      level,
      levelLabel: logLevelLabels[level],
      levelClass: level.toLowerCase(),
      message: item.message,
      messageSegments: formatLogMessageSegments(item.message),
      source: item.source || 'runtime',
      channel,
      phase: item.phase || '',
      taskIndex: item.taskIndex ?? '',
      articleTaskId: item.articleTaskId || '',
      articleTitle: item.articleTitle || '',
    }
  })

  return rows.filter((item) => activeLogLevel.value === 'ALL' || item.level === activeLogLevel.value)
})

const systemLogRows = computed<LogDisplayRow[]>(() =>
  allLogRows.value.filter((row) => !isArticleTaskLog(row)),
)
const articleTaskLogRows = computed<LogDisplayRow[]>(() =>
  allLogRows.value.filter((row) => isArticleTaskLog(row)),
)

const statusItems = computed(() => [
  {
    label: '当前状态',
    value: taskStatusLabel.value,
    icon: 'fa-solid fa-circle-play',
    tag: true,
    tone: taskStatusTone.value,
  },
  { label: '当前动作', value: runtimeState.value.currentAction || '待获取', icon: 'fa-solid fa-bolt' },
  { label: '公众号名称', value: runtimeState.value.accountName || '等待识别', icon: 'fa-brands fa-weixin' },
  { label: '任务信息', value: runtimeState.value.taskInfo || '待获取', icon: 'fa-regular fa-file-lines' },
  { label: '代理状态', value: runtimeState.value.proxyStatusLabel || '空闲', icon: 'fa-solid fa-network-wired' },
  { label: '采集进度', value: '', icon: 'fa-solid fa-gauge-high', progress: true },
  { label: '平均时长', value: runtimeState.value.averageArticleDurationLabel || '待统计', icon: 'fa-regular fa-clock' },
  {
    label: '活跃子进程',
    value: `${runtimeState.value.activeWorkerCount ?? 0}/${runtimeState.value.totalWorkerCount ?? 0}`,
    icon: 'fa-solid fa-list-check',
  },
  {
    label: '异常记录',
    value: String(runtimeState.value.errorCount ?? taskProgressSummary.value.failedCount),
    icon: 'fa-solid fa-triangle-exclamation',
    tone: 'red' as Tone,
  },
])

const stats = computed(() => buildArchiveSummaryStats(archiveSummary.value, formatDuration(uptimeSeconds.value)))

const mandatoryDownloadOption = { key: 'articleDetail', label: '文章详情', locked: true } as const

const offlineArchiveModeOptions = [
  { value: 'standard', label: '离线归档' },
  { value: 'beta', label: '离线归档 (beta)' },
] satisfies Array<{ value: OfflineArchiveMode; label: string }>

const selectedOfflineArchiveModeLabel = computed(() => (
  offlineArchiveModeOptions.find((option) => option.value === offlineArchiveMode.value)?.label ?? '离线归档'
))
const offlineArchiveSelected = computed(() => downloadSelections.value.offlineArchive)

const downloadOptions = [
  { key: 'commentInfo', label: '评论信息', locked: false },
  { key: 'skipCollectedRecords', label: '跳过已采集记录', locked: false },
] as const

const usageTips = [
  '点击开始运行前打开目标 PC 端',
  '手动打开目标公众号主页窗口',
  '开始运行后将自动读取主页信息',
  '出现跳转操作时无需手动干预',
]

const complianceTips = [
  '本工具仅用于学习研究与合规用途',
  '整理公开文章材料、字段内容记录',
  '不存储、分享或发布任何侵权内容',
  '请遵守相关法律、政策和平台条款',
]

const envItems = computed(() => [
  { name: 'Version', value: environmentStatus.value.appVersion, icon: 'fa-solid fa-cube' },
  { name: 'Python', value: environmentStatus.value.pythonVersion, icon: 'fa-brands fa-python' },
  { name: 'System', value: environmentStatus.value.systemLabel, icon: 'fa-brands fa-windows' },
  { name: 'Tauri', value: environmentStatus.value.desktopShellStatus, icon: 'fa-regular fa-window-maximize' },
  { name: 'MITMproxy', value: environmentStatus.value.mitmproxyVersion, icon: 'fa-solid fa-shield-halved' },
  { name: 'Playwright', value: environmentStatus.value.playwrightVersion, icon: 'fa-solid fa-window-restore' },
])

// 使用 Ant Design Vue 原生渐变配置，避免依赖内部 DOM 背景覆盖。
const progressStrokeColor = computed(() => ({
  from: isDark.value ? '#72a7ff' : '#2f80d9',
  to: isDark.value ? '#55d6bf' : '#37c2a3',
}))

const progressPercent = computed(() => {
  const runtimeProgress = Number(runtimeState.value.progressPercent)
  if (Number.isFinite(runtimeProgress)) {
    return Math.min(100, Math.max(0, Math.round(runtimeProgress)))
  }

  return taskProgressSummary.value.progressPercent
})

const progressLineStatus = computed(() => {
  return progressPercent.value > 0 && progressPercent.value < 100 ? 'active' : 'normal'
})

function parseConfigSwitchValue(value: unknown, fallback: boolean) {
  if (typeof value === 'boolean') {
    return value
  }

  if (typeof value !== 'string') {
    return fallback
  }

  const normalized = value.trim().toLowerCase()
  if (['开启', '开', 'true', '1', 'yes', 'on'].includes(normalized)) {
    return true
  }
  if (['关闭', '关', 'false', '0', 'no', 'off'].includes(normalized)) {
    return false
  }

  return fallback
}

function buildMainTaskSelectionDefaultsSignature(values: Record<string, string>) {
  // 只跟踪会影响主服务页默认勾选的配置项，避免普通状态轮询覆盖用户手动选择。
  return [
    values['single_article_task.comment_collection.enabled_by_default'] ?? '',
    values['single_article_task.offline_cache.enabled_by_default'] ?? '',
  ].join('|')
}

function applyMainTaskSelectionDefaults(values?: Record<string, string>) {
  // YAML 负责默认勾选；配置未变化时保留用户在主服务页的手动选择。
  if (taskSettingsLocked.value || !values) return

  const defaultsSignature = buildMainTaskSelectionDefaultsSignature(values)
  if (
    mainTaskSelectionDefaultsApplied.value
    && mainTaskSelectionDefaultsSignature.value === defaultsSignature
  ) return

  downloadSelections.value.commentInfo = parseConfigSwitchValue(
    values['single_article_task.comment_collection.enabled_by_default'],
    downloadSelections.value.commentInfo,
  )
  downloadSelections.value.offlineArchive = parseConfigSwitchValue(
    values['single_article_task.offline_cache.enabled_by_default'],
    downloadSelections.value.offlineArchive,
  )
  mainTaskSelectionDefaultsSignature.value = defaultsSignature
  mainTaskSelectionDefaultsApplied.value = true
}

function handleDownloadSelectionChange() {
  if (taskSettingsLocked.value) {
    return
  }

  mainTaskSelectionDefaultsApplied.value = true
}

function handleOfflineArchiveModeChange(value: OfflineArchiveMode) {
  if (taskSettingsLocked.value) {
    return
  }

  offlineArchiveMode.value = value
  offlineArchiveModeOpen.value = false
  mainTaskSelectionDefaultsApplied.value = true
}

function buildTaskRunOptions(): TaskRunOptions {
  const mode = taskDateFilterMode.value
  return {
    recordLimit: normalizedPageCount.value,
    dateFilterMode: mode,
    startDate: mode === 'range' || mode === 'after' ? (taskStartDate.value || undefined) : undefined,
    endDate: mode === 'range' || mode === 'before' ? (taskEndDate.value || undefined) : undefined,
    selections: {
      articleDetail: true,
      offlineArchive: downloadSelections.value.offlineArchive,
      commentInfo: downloadSelections.value.commentInfo,
      skipCollectedRecords: downloadSelections.value.skipCollectedRecords,
      offlineArchiveMode: offlineArchiveMode.value,
    },
  }
}

function validateTaskRunOptions(options: TaskRunOptions): string | null {
  if (options.dateFilterMode === 'range' && (!options.startDate || !options.endDate)) {
    return '请选择完整的日期范围。'
  }
  if (options.dateFilterMode === 'before' && !options.endDate) {
    return '请选择截止日期。'
  }
  if (options.dateFilterMode === 'after' && !options.startDate) {
    return '请选择起始日期。'
  }
  if (options.dateFilterMode === 'range' && options.startDate && options.endDate && options.startDate > options.endDate) {
    return '起始日期不能晚于截止日期。'
  }
  return null
}

async function handleStartTask() {
  if (taskSettingsLocked.value) {
    return
  }

  const options = buildTaskRunOptions()
  const validationMessage = validateTaskRunOptions(options)
  if (validationMessage) {
    appendFrontendRuntimeError(validationMessage)
    return
  }

  try {
    const status = await startTask(options)
    handleTaskStatusChanged(status)
    homeRuntimeState.value = 'running'
    startTaskPolling()
  } catch (error) {
    const message = `启动采集任务失败：${formatErrorMessage(error)}`
    appendFrontendRuntimeError(message)
    try {
      const status = await getTaskStatus()
      handleTaskStatusChanged(status)
    } catch (refreshError) {
      appendFrontendRuntimeError(`读取启动失败状态失败：${formatErrorMessage(refreshError)}`)
    }
  }
}

async function handleStopTask() {
  if (!['starting', 'running'].includes(taskStatus.value.status)) {
    return
  }

  try {
    const status = await stopTask()
    handleTaskStatusChanged(status)
    startTaskPolling()
  } catch (error) {
    appendFrontendRuntimeError(`停止采集任务失败：${formatErrorMessage(error)}`)
  }
}

function markHomeDetectionStarting() {
  homeRuntimeState.value = 'detecting'
}

function isStatusValueOverflowing(event: MouseEvent | FocusEvent) {
  const target = event.currentTarget as HTMLElement | null
  if (!target) {
    return false
  }

  return target.scrollWidth > target.clientWidth
}

function prepareStatusValueTooltip(event: MouseEvent | FocusEvent, key: string) {
  statusValueTooltipKey.value = isStatusValueOverflowing(event) ? key : ''
}

function hideStatusValueTooltip(key: string) {
  if (statusValueTooltipKey.value === key) {
    statusValueTooltipKey.value = ''
  }
}

function selectPage(page: PageKey | null) {
  if (page) {
    activePage.value = page
    void refreshActivePage(page)
  }
}

async function refreshActivePage(page: PageKey) {
  const refreshSequence = ++pageRefreshSequence
  await nextTick()
  if (refreshSequence !== pageRefreshSequence || activePage.value !== page) return
  await pageActivationRefreshers[page]()
}

function formatDuration(totalSeconds: number) {
  const safeSeconds = Math.max(0, Math.floor(totalSeconds))
  const hours = Math.floor(safeSeconds / 3600)
  const minutes = Math.floor((safeSeconds % 3600) / 60)
  const seconds = safeSeconds % 60

  return [hours, minutes, seconds].map((item) => String(item).padStart(2, '0')).join(':')
}

function normalizeLogLevel(level: string): RuntimeLogLevel {
  const normalized = String(level || 'INFO').toUpperCase()
  if (normalized === 'SUCCESS' || normalized === 'WARN' || normalized === 'ERROR') {
    return normalized
  }

  return 'INFO'
}

function normalizeLogChannel(channel?: string): string {
  const normalized = String(channel || 'system').trim().toLowerCase().replace(/-/g, '_')
  return normalized === 'article_task' || normalized === 'main_flow' ? normalized : 'system'
}

function isArticleTaskLog(row: LogDisplayRow): boolean {
  return row.channel === 'article_task'
}

function formatLogTime(createdAt: string) {
  const value = String(createdAt || '').trim()
  if (!value) {
    return '--:--:--'
  }

  const timeText = value.includes('T') ? value.split('T')[1] : value.split(' ')[1]
  return (timeText || value).slice(0, 8) || '--:--:--'
}

function formatErrorMessage(error: unknown) {
  if (error instanceof Error) {
    return error.message
  }

  return String(error || '未知异常')
}

async function refreshStartupHealthChecks() {
  if (startupHealthCheckRequested) {
    return
  }
  startupHealthCheckRequested = true
  startupHealthCheckRunning.value = true
  for (const target of HEALTH_CHECK_TARGETS) {
    topbarHealthStates.value[target] = buildTopbarCheckingState(
      `正在执行${HEALTH_CHECK_META[target].label}启动检测...`,
    )
  }

  try {
    const result = await runStartupHealthChecks()
    for (const target of HEALTH_CHECK_TARGETS) {
      const healthResult = result.results[target]
      if (healthResult) {
        applyHealthCheckResult(healthResult)
      } else {
        applyHealthCheckFailure(target, `${HEALTH_CHECK_META[target].label}未返回检测结果。`)
      }
    }
  } catch (error) {
    const message = `启动健康检测失败：${formatErrorMessage(error)}`
    for (const target of HEALTH_CHECK_TARGETS) {
      applyHealthCheckFailure(target, message)
    }
    appendFrontendRuntimeError(message)
  } finally {
    startupHealthCheckRunning.value = false
  }
}

function buildStartupSelfCheckItems(result: StartupSelfCheckResult): HealthCheckResult['items'] {
  return result.items.map((item) => ({
    key: item.key,
    label: item.label,
    value: item.message,
    message: item.action || '',
    ok: item.ok,
  }))
}

function applyStartupSelfCheckResult(result: StartupSelfCheckResult) {
  const hasWarning = result.warningCount > 0
  startupSelfCheckDialogState.value = {
    value: result.ok ? (hasWarning ? '有警告' : '通过') : '异常',
    statusIcon: result.ok
      ? 'fa-solid fa-circle-check'
      : 'fa-solid fa-triangle-exclamation',
    tone: result.ok ? (hasWarning ? 'warning' : 'success') : 'danger',
    message: result.ok
      ? '自检完成：' + result.fatalCount + ' 个致命问题，' + result.warningCount + ' 个警告。'
      : '自检发现 ' + result.fatalCount + ' 个致命问题，' + result.warningCount + ' 个警告。',
    items: buildStartupSelfCheckItems(result),
    checkedAt: result.checkedAt,
    checking: false,
  }
}

async function refreshStartupSelfCheck() {
  if (startupSelfCheckRequested) {
    return
  }
  startupSelfCheckRequested = true

  try {
    const status = await getStartupSelfCheckStatus()
    if (!status.needsSelfCheck) {
      return
    }

    startupSelfCheckDialogState.value = buildTopbarCheckingState('正在自检：检查程序运行环境、MITM、Playwright、SQLite 和存储配置...')
    startupSelfCheckDialogVisible.value = true
    const result = await runStartupSelfCheck()
    applyStartupSelfCheckResult(result)
    if (!result.ok) {
      appendFrontendRuntimeError('启动自检发现 ' + result.fatalCount + ' 个致命问题。')
    }
  } catch (error) {
    const message = '启动自检失败：' + formatErrorMessage(error)
    startupSelfCheckDialogVisible.value = true
    startupSelfCheckDialogState.value = {
      value: '失败',
      statusIcon: 'fa-solid fa-triangle-exclamation',
      tone: 'danger',
      message,
      items: [],
      checkedAt: new Date().toISOString().slice(0, 19),
      checking: false,
    }
    appendFrontendRuntimeError(message)
  }
}

async function handleTopbarHealthCheck(target: HealthCheckTarget) {
  if (startupHealthCheckRunning.value || healthCheckInProgress.value !== null) return
  healthDialogTarget.value = target
  healthDialogVisible.value = true
  const checkingState = buildTopbarCheckingState(`正在重新检测${HEALTH_CHECK_META[target].label}...`)
  topbarHealthStates.value[target] = checkingState
  healthCheckInProgress.value = target

  try {
    const result = await checkHealthTarget(target)
    applyHealthCheckResult(result)
  } catch (error) {
    const message = `${HEALTH_CHECK_META[target].label}检测失败：${formatErrorMessage(error)}`
    applyHealthCheckFailure(target, message)
    appendFrontendRuntimeError(message)
  } finally {
    healthCheckInProgress.value = null
  }
}

function appendFrontendRuntimeError(message: string) {
  const latest = frontendRuntimeLogs.value[frontendRuntimeLogs.value.length - 1]
  if (latest?.message === message) {
    return
  }

  frontendRuntimeLogs.value = [
    ...frontendRuntimeLogs.value,
    {
      level: 'ERROR',
      message,
      source: 'frontend',
      createdAt: new Date().toISOString().slice(0, 19),
      channel: 'system',
      phase: 'frontend',
    },
  ].slice(-LOG_POLL_LIMIT)
}

async function scrollLogTableToLatest() {
  await nextTick()
  if (
    (!systemLogRows.value.length && !articleTaskLogRows.value.length)
    || !shouldStickLogToBottom.value
  ) {
    return
  }

  systemLogScrollerRef.value?.scrollToBottom()
  articleTaskLogScrollerRef.value?.scrollToBottom()
}

function stopLogAutoFollowTimer() {
  window.clearTimeout(logAutoFollowTimer)
  logAutoFollowTimer = undefined
}

function scheduleLogAutoFollowResume() {
  stopLogAutoFollowTimer()
  if (shouldStickLogToBottom.value) {
    return
  }

  const lastInteractionAt = lastLogScrollUserInteractionAt.value || Date.now()
  const elapsed = Date.now() - lastInteractionAt
  const delay = Math.max(0, LOG_AUTO_FOLLOW_IDLE_MS - elapsed)
  logAutoFollowTimer = window.setTimeout(async () => {
    if (!canResumeLogAutoFollow(lastLogScrollUserInteractionAt.value, Date.now())) {
      scheduleLogAutoFollowResume()
      return
    }

    shouldStickLogToBottom.value = true
    await scrollLogTableToLatest()
  }, delay)
}

function markLogScrollUserInteraction() {
  lastLogScrollUserInteractionAt.value = Date.now()
  scheduleLogAutoFollowResume()
}

async function selectLogLevel(level: LogFilterLevel) {
  activeLogLevel.value = level
  shouldStickLogToBottom.value = true
  stopLogAutoFollowTimer()
  await scrollLogTableToLatest()
}

function handleLogSegmentChange(value: string | number) {
  const matched = logTabs.find((tab) => tab.level === value)
  if (!matched) {
    return
  }

  void selectLogLevel(matched.level)
}

function handleLogTableScroll(event: Event) {
  const metrics = readLogScrollMetricsFromEvent(event)
  if (!metrics) {
    return
  }

  const nearBottom = isLogScrollNearBottom(metrics)
  shouldStickLogToBottom.value = nearBottom
  if (nearBottom) {
    stopLogAutoFollowTimer()
    return
  }

  scheduleLogAutoFollowResume()
}

function stopDesktopStatusRetry() {
  window.clearTimeout(desktopStatusRetryTimer)
  desktopStatusRetryTimer = undefined
}

function applyDesktopEnvironmentStatus(resolved: ResolvedDesktopEnvironmentStatus) {
  desktopStatusLabel.value = resolved.desktopStatusLabel
  environmentStatus.value = resolved.environmentStatus
}

function scheduleDesktopStatusRetry() {
  stopDesktopStatusRetry()

  if (desktopStatusRetryCount >= MAX_DESKTOP_STATUS_RETRIES) {
    applyDesktopEnvironmentStatus(getBrowserPreviewEnvironmentStatus())
    return
  }

  desktopStatusRetryCount += 1
  desktopStatusRetryTimer = window.setTimeout(() => {
    refreshPythonStatus()
  }, DESKTOP_STATUS_RETRY_DELAY_MS)
}

async function refreshPythonStatus() {
  try {
    const status = await getPythonStatus()
    const resolved = resolveDesktopEnvironmentStatus(status)
    applyDesktopEnvironmentStatus(resolved)

    if (resolved.shouldRetry) {
      scheduleDesktopStatusRetry()
      return
    }

    stopDesktopStatusRetry()
  } catch {
    stopDesktopStatusRetry()
    applyDesktopEnvironmentStatus(getEnvironmentErrorStatus())
  }
}

function stopTaskPolling() {
  window.clearInterval(taskPollingTimer)
  taskPollingTimer = undefined
}

async function refreshTaskRuntime() {
  try {
    const [status, logs] = await Promise.all([getTaskStatus(), getTaskLogs(LOG_POLL_LIMIT)])
    taskStatus.value = {
      ...status,
      hardware: status.hardware ?? taskStatus.value.hardware ?? defaultHardwareStatus,
    }
    applyMainTaskSelectionDefaults(status.config?.values)
    uptimeSeconds.value = status.uptimeSeconds ?? uptimeSeconds.value
    taskLogs.value = Array.isArray(logs.items) ? logs.items : []
    homeRuntimeState.value = ['starting', 'running', 'stopping'].includes(status.status) ? 'running' : 'idle'
    syncTaskPolling(status)
  } catch (error) {
    appendFrontendRuntimeError(`读取运行状态失败：${formatErrorMessage(error)}`)
  }
}

async function refreshArchiveSummary() {
  try {
    archiveSummary.value = await listArchiveSummary()
  } catch (error) {
    archiveSummary.value = null
    appendFrontendRuntimeError(`读取数据统计失败：${formatErrorMessage(error)}`)
  }
}

function handleTaskStatusChanged(status: TaskStatus) {
  taskStatus.value = {
    ...status,
    hardware: status.hardware ?? taskStatus.value.hardware ?? defaultHardwareStatus,
  }
  applyMainTaskSelectionDefaults(status.config?.values)
  uptimeSeconds.value = status.uptimeSeconds ?? uptimeSeconds.value
  syncTaskPolling(status)
}

async function refreshHardwareStatus() {
  try {
    const hardware = await getHardwareStatus()
    taskStatus.value = {
      ...taskStatus.value,
      hardware,
    }
    hardwareStatusErrorReported = false
  } catch (error) {
    if (!hardwareStatusErrorReported) {
      hardwareStatusErrorReported = true
      appendFrontendRuntimeError(`读取硬件状态失败：${formatErrorMessage(error)}`)
    }
  }
}

async function refreshHomePageOnActivated() {
  await Promise.all([refreshTaskRuntime(), refreshArchiveSummary(), refreshHardwareStatus()])
}

const pageActivationRefreshers: Record<PageKey, () => void | Promise<void>> = {
  home: refreshHomePageOnActivated,
  files: () => dataFilesPageRef.value?.refreshOnActivated?.(),
  history: () => historyPageRef.value?.refreshOnActivated?.(),
  settings: () => settingsPageRef.value?.refreshOnActivated?.(),
}

function startHardwarePolling() {
  if (hardwarePollingTimer !== undefined) {
    return
  }
  hardwarePollingTimer = window.setInterval(() => {
    void refreshHardwareStatus()
  }, HARDWARE_POLL_INTERVAL_MS)
}

function stopHardwarePolling() {
  window.clearInterval(hardwarePollingTimer)
  hardwarePollingTimer = undefined
}

function startTaskPolling() {
  if (taskPollingTimer !== undefined) {
    return
  }

  taskPollingTimer = window.setInterval(() => {
    refreshTaskRuntime()
  }, 1500)
}

function shouldPollTaskRuntime(status: TaskStatus) {
  return ['starting', 'running', 'stopping'].includes(status.status)
}

// 诊断接口可能返回 managed-by-capture，这类结果不代表采集任务已结束。
function isTaskLifecycleStatus(status: TaskStatus) {
  return ['idle', 'starting', 'running', 'stopping', 'completed', 'cancelled', 'stopped', 'error', 'failed'].includes(status.status)
}

function syncTaskPolling(status: TaskStatus) {
  if (shouldPollTaskRuntime(status)) {
    startTaskPolling()
    return
  }

  if (!isTaskLifecycleStatus(status)) {
    return
  }

  stopTaskPolling()
}

function stopUptimeTimer() {
  window.clearInterval(uptimeTimer)
  uptimeTimer = undefined
}

function startUptimeTimer() {
  stopUptimeTimer()
  uptimeTimer = window.setInterval(() => {
    uptimeSeconds.value += 1
  }, 1000)
}

async function handleOpenLogFolder() {
  try {
    const result = await openRuntimePath('logDir')
    if (!result.ok) {
      appendFrontendRuntimeError(result.message || '打开日志目录失败。')
    }
  } catch (error) {
    appendFrontendRuntimeError(`打开日志目录失败：${formatErrorMessage(error)}`)
  }
}

onMounted(async () => {
  markHomeDetectionStarting()
  refreshPythonStatus()
  refreshArchiveSummary()
  startUptimeTimer()
  startHardwarePolling()
  await refreshTaskRuntime()
  await refreshHardwareStatus()
  await refreshStartupSelfCheck()
  await refreshStartupHealthChecks()
})
watch([systemLogRows, articleTaskLogRows], async () => {
  await scrollLogTableToLatest()
}, { flush: 'post' })
onBeforeUnmount(() => {
  stopDesktopStatusRetry()
  stopLogAutoFollowTimer()
  stopTaskPolling()
  stopHardwarePolling()
  stopUptimeTimer()
})
</script>

<template>
  <AConfigProvider :locale="zhCN" :theme="antThemeConfig">
    <main :class="['collector-app', { dark: isDark }]">
    <section class="workspace">
      <AppTopbar
        :is-dark="isDark"
        :github-url="githubUrl"
        :health-items="topbarHealthItems"
        @toggle-theme="isDark = !isDark"
        @health-check="handleTopbarHealthCheck"
      />

      <div class="dashboard">
        <aside class="sidebar panel">
          <nav
            v-for="group in navGroups"
            :key="group.title"
            class="nav-group"
            :aria-label="group.title"
          >
            <p class="nav-title">{{ group.title }} <span aria-hidden="true">⌁</span></p>
            <template
              v-for="item in group.items"
              :key="item.label"
            >
              <a
                v-if="item.href"
                class="nav-item"
                :href="item.href"
                target="_blank"
                rel="noopener noreferrer"
              >
                <AppIcon :icon="['nav-icon', item.icon]" />
                {{ item.label }}
              </a>
              <button
                v-else
                :class="['nav-item', { active: item.page === activePage }]"
                type="button"
                :disabled="!item.page"
                :aria-current="item.page === activePage ? 'page' : undefined"
                @click="selectPage(item.page)"
              >
                <AppIcon :icon="['nav-icon', item.icon]" />
                {{ item.label }}
              </button>
            </template>
          </nav>

          <img class="sidebar-book" src="/assets/watercolor-book-stack.png" alt="" />
        </aside>

        <template v-if="activePage === 'home'">
        <section class="task-column" aria-label="采集任务设置">
          <article class="task-card task-card-volume panel">
            <span class="step-badge">01</span>
            <img class="task-art task-art-list" src="/assets/watercolor-task-list.png" alt="" />
            <div class="task-volume-body">
              <h2>指定记录总量</h2>
              <div class="task-volume-row task-date-filter-row">
                <div class="task-control-label task-label-with-hint">
                  <span>日期筛选</span>
                  <ATooltip :title="taskDateFilterHint" placement="top">
                    <button
                      class="task-label-hint"
                      type="button"
                      aria-label="查看日期筛选说明"
                    >
                      <QuestionCircleOutlined />
                    </button>
                  </ATooltip>
                </div>
                <ADropdown
                  v-model:open="taskDateFilterOpen"
                  :trigger="['click']"
                  :disabled="taskSettingsLocked"
                  placement="bottomLeft"
                  overlay-class-name="task-date-filter-dropdown"
                >
                  <AButton
                    class="task-date-filter-trigger"
                    :disabled="taskSettingsLocked"
                    aria-label="日期筛选方式"
                  >
                    <span>{{ taskDateFilterLabel }}</span>
                    <DownOutlined
                      :class="['task-date-filter-chevron', { 'is-open': taskDateFilterOpen }]"
                    />
                  </AButton>
                  <template #overlay>
                    <AMenu :selected-keys="[taskDateFilterMode]">
                      <AMenuItem
                        v-for="option in taskDateFilterOptions"
                        :key="option.value"
                        @click="selectTaskDateFilterMode(option.value)"
                      >
                        {{ option.label }}
                      </AMenuItem>
                    </AMenu>
                  </template>
                </ADropdown>
              </div>

              <div class="task-volume-row task-count-row">
                <div class="task-control-label task-label-with-hint">
                  <label for="page-count">设定任务</label>
                  <ATooltip title="设为0时遍历日期范围内全部内容" placement="top">
                    <button
                      class="task-label-hint"
                      type="button"
                      aria-label="查看设定任务说明"
                    >
                      <QuestionCircleOutlined />
                    </button>
                  </ATooltip>
                </div>
                <div class="task-count-control">
                  <AInputNumber
                    id="page-count"
                    v-model:value="pageCount"
                    class="task-count-input"
                    :min="0"
                    :precision="0"
                    :controls="true"
                    :disabled="taskSettingsLocked"
                    aria-label="设定任务"
                  />
                </div>
              </div>

              <label class="task-volume-row task-date-row">
                <span class="task-control-label task-date-label download-option selected locked">
                  <ASpin class="task-date-label-spin" size="small" :spinning="true" />
                  {{ taskDateFieldLabel }}
                </span>
                <span class="task-date-control">
                  <ARangePicker
                    v-if="taskDateFilterMode === 'range'"
                    v-model:value="taskDateRangeValue"
                    class="task-date-picker task-date-range-picker"
                    :allow-clear="true"
                    :placeholder="['选择起始日期', '选择截止日期']"
                    format="YYYY-MM-DD"
                    value-format="YYYY-MM-DD"
                    popup-class-name="task-date-picker-panel"
                    :disabled="taskSettingsLocked"
                    aria-label="任务日期范围"
                  />
                  <ADatePicker
                    v-else-if="taskDateFilterMode === 'before'"
                    v-model:value="taskEndDate"
                    class="task-date-picker"
                    :allow-clear="true"
                    placeholder="选择截止日期"
                    format="YYYY-MM-DD  dddd"
                    value-format="YYYY-MM-DD"
                    popup-class-name="task-date-picker-panel"
                    :disabled="taskSettingsLocked"
                    aria-label="任务截止日期"
                  />
                  <ADatePicker
                    v-else-if="taskDateFilterMode === 'after'"
                    v-model:value="taskStartDate"
                    class="task-date-picker"
                    :allow-clear="true"
                    placeholder="选择起始日期"
                    format="YYYY-MM-DD  dddd"
                    value-format="YYYY-MM-DD"
                    popup-class-name="task-date-picker-panel"
                    :disabled="taskSettingsLocked"
                    aria-label="任务起始日期"
                  />
                  <ADatePicker
                    v-else-if="taskDateFilterMode === 'all'"
                    class="task-date-picker"
                    :value="null"
                    placeholder="不限日期，无需选择"
                    :disabled="true"
                    aria-label="不限日期"
                  />
                </span>
              </label>
            </div>
          </article>

          <article class="task-card task-card-content panel">
            <span class="step-badge">02</span>
            <div class="task-art-column">
              <img class="task-art task-art-heart" src="/assets/watercolor-task-heart.png" alt="" />
              <ACheckbox
                v-model:checked="downloadSelections[mandatoryDownloadOption.key]"
                :class="[
                  'download-option',
                  'article-detail-option',
                  {
                    selected: downloadSelections[mandatoryDownloadOption.key],
                    locked: mandatoryDownloadOption.locked || taskSettingsLocked,
                  },
                ]"
                :disabled="mandatoryDownloadOption.locked || taskSettingsLocked"
                title="文章详情为必选项，不能取消"
              >
                {{ mandatoryDownloadOption.label }}
              </ACheckbox>
            </div>
            <div class="task-body task-body-content">
              <h2>获取指定内容</h2>
              <div class="download-options" aria-label="选择获取内容">
                <div
                  :class="[
                    'download-option',
                    'download-option-combo',
                    {
                      selected: offlineArchiveSelected,
                      locked: taskSettingsLocked,
                    },
                  ]"
                >
                  <ACheckbox
                    v-model:checked="downloadSelections.offlineArchive"
                    class="offline-archive-checkbox"
                    :disabled="taskSettingsLocked"
                    :title="taskSettingsLocked ? '任务运行期间不能修改' : ''"
                    @change="handleDownloadSelectionChange"
                  >
                    {{ selectedOfflineArchiveModeLabel }}
                  </ACheckbox>
                  <ADropdown
                    v-model:open="offlineArchiveModeOpen"
                    :trigger="['click']"
                    :disabled="taskSettingsLocked"
                    placement="bottomRight"
                    overlay-class-name="offline-archive-mode-dropdown"
                  >
                    <AButton
                      class="offline-archive-mode-trigger"
                      :disabled="taskSettingsLocked"
                      aria-label="选择离线归档模式"
                    >
                      <DownOutlined
                        :class="['download-select-chevron', { 'is-open': offlineArchiveModeOpen }]"
                      />
                    </AButton>
                    <template #overlay>
                      <AMenu :selected-keys="[offlineArchiveMode]">
                        <AMenuItem
                          v-for="option in offlineArchiveModeOptions"
                          :key="option.value"
                          @click="handleOfflineArchiveModeChange(option.value)"
                        >
                          {{ option.label }}
                        </AMenuItem>
                      </AMenu>
                    </template>
                  </ADropdown>
                </div>
                <ACheckbox
                  v-for="option in downloadOptions"
                  :key="option.key"
                  v-model:checked="downloadSelections[option.key]"
                  :class="[
                    'download-option',
                    {
                      selected: downloadSelections[option.key],
                      locked: option.locked || taskSettingsLocked,
                    },
                  ]"
                  :disabled="option.locked || taskSettingsLocked"
                  :title="taskSettingsLocked ? '任务运行期间不能修改' : ''"
                  @change="handleDownloadSelectionChange"
                >
                  {{ option.label }}
                </ACheckbox>
              </div>
            </div>
          </article>

          <div class="control-panel panel" aria-label="任务控制">
            <AButton
              class="run-button"
              type="primary"
              html-type="button"
              size="large"
              :loading="taskStatus.status === 'starting'"
              :disabled="taskSettingsLocked"
              @click="handleStartTask"
            >
              <AppIcon icon="fa-solid fa-play" />
              <span class="button-label">开始运行</span>
            </AButton>
            <AButton
              class="stop-button"
              html-type="button"
              size="large"
              danger
              :loading="taskStatus.status === 'stopping'"
              :disabled="taskStatus.status !== 'running' && taskStatus.status !== 'starting'"
              @click="handleStopTask"
            >
              <AppIcon icon="fa-solid fa-stop" />
              <span class="button-label">停止</span>
            </AButton>
          </div>
        </section>

        <section class="status-column">
          <article class="status-card panel sprig-corner status-corner">
            <!-- 右上角插画使用裁剪后的独立图片，直接贴边显示，不再做遮挡式装饰。 -->
            <img class="corner-image status-corner-image" src="/assets/watercolor-flower-corner.png" alt="" />
            <div class="section-title">
              <svg
                class="title-icon status-title-logo"
                viewBox="0 0 1024 1024"
                aria-hidden="true"
                focusable="false"
              >
                <path
                  fill="currentColor"
                  d="M41.984 486.4l56.32 0q22.528 0 36.352-4.096t27.136-15.36l16.384-16.384q10.24-10.24 21.504-23.552t23.04-27.136 22.016-26.112q20.48-25.6 36.864-19.968t24.576 24.064l45.056 123.904q34.816-97.28 61.44-176.128 11.264-33.792 22.528-66.048t20.48-59.392 15.36-45.056 7.168-23.04q5.12-17.408 15.36-28.16t22.528-10.752q13.312 0 27.648 7.68t18.432 31.232q1.024 6.144 4.096 32.768t8.192 66.56 10.752 88.576 11.776 98.816q13.312 117.76 30.72 263.168 22.528-73.728 40.96-134.144 7.168-25.6 14.848-50.688t14.336-46.08 10.752-35.328 6.144-18.432q5.12-16.384 12.288-24.576t19.456-11.264q23.552-6.144 35.328 6.656t11.776 26.112q0 11.264-2.048 19.456-1.024 4.096-1.024 7.168l36.864 100.352q1.024-1.024 3.072-4.096 3.072-5.12 10.24-21.504 8.192-17.408 27.136-23.04t36.352-4.608q11.264 1.024 22.528 1.024l22.528 0 24.576 0 0 77.824q-6.144 1.024-12.288 1.024l-26.624 0q-14.336-1.024-26.624 7.68t-17.408 16.896q-4.096 8.192-14.336 31.744t-18.432 48.128q-4.096 12.288-13.312 16.384t-18.944 3.584-18.432-5.632-11.776-12.288-10.24-25.088-14.336-36.352q-9.216-21.504-18.432-46.08-26.624 84.992-48.128 156.672-9.216 30.72-18.432 60.416t-16.896 54.784-12.8 42.496-6.144 23.552q-5.12 24.576-19.456 38.912t-34.816 13.312q-21.504-2.048-30.72-18.432t-11.264-37.888q0-5.12-2.56-31.744t-6.656-67.584-8.704-91.136-9.728-102.4q-11.264-121.856-26.624-272.384-28.672 83.968-51.2 150.528-9.216 28.672-18.944 56.32t-17.408 50.176-12.8 37.888l-6.144 18.432q-5.12 13.312-13.824 24.576t-27.136 10.24q-18.432 0-27.648-12.8t-14.336-29.184q-3.072-8.192-11.264-31.744t-17.408-48.128q-11.264-28.672-23.552-63.488-7.168 9.216-15.36 18.432-7.168 8.192-15.36 17.92t-16.384 18.944q-16.384 19.456-30.72 26.624t-33.792 5.12q-10.24-1.024-27.136-0.512t-34.304 0.512q-19.456 1.024-43.008 1.024l0-79.872z"
                />
              </svg>
              <h2>程序运行状态</h2>
            </div>

            <div class="status-list">
              <div
                v-for="item in statusItems"
                :key="item.label"
                :class="['status-row', { 'with-progress': item.progress }]"
              >
                <AppIcon :icon="['row-icon', item.icon]" />
                <span class="status-label">{{ item.label }}：</span>
                <ATag
                  v-if="item.tag"
                  :class="['status-pill', item.tone]"
                  :color="item.tone"
                  :bordered="false"
                >
                  {{ item.value }}
                </ATag>
                <template v-else-if="item.progress">
                  <AProgress
                    class="progress-line"
                    aria-label="采集进度"
                    :percent="progressPercent"
                    :stroke-color="progressStrokeColor"
                    :status="progressLineStatus"
                    :show-info="false"
                    size="small"
                  />
                  <em>{{ taskProgressLabel }}</em>
                </template>
                <ATooltip
                  v-else
                  :title="item.value"
                  :open="statusValueTooltipKey === item.label"
                  placement="topLeft"
                  :destroy-tooltip-on-hide="true"
                >
                  <strong
                    :class="['status-value', 'status-value-ellipsis', 'status-description-value', item.tone]"
                    tabindex="0"
                    @mouseenter="prepareStatusValueTooltip($event, item.label)"
                    @mouseleave="hideStatusValueTooltip(item.label)"
                    @focus="prepareStatusValueTooltip($event, item.label)"
                    @blur="hideStatusValueTooltip(item.label)"
                  >
                    {{ item.value }}
                  </strong>
                </ATooltip>
              </div>
            </div>

            <div class="network-panel">
              <div class="hardware-labels" aria-label="硬件指标">
                <span class="hardware-label-row">
                  <AppIcon icon="hardware-label-icon hardware-cpu-icon fa-solid fa-microchip" />
                  <span class="hardware-label-box">
                    <span class="hardware-label-text">CPU 使用占比</span>
                  </span>
                </span>
                <span class="hardware-label-row">
                  <AppIcon icon="hardware-label-icon hardware-memory-icon fa-solid fa-memory" />
                  <span class="hardware-label-box">
                    <span class="hardware-label-text">物理内存占比</span>
                  </span>
                </span>
              </div>
              <div class="speed-content">
                <HardwareSparkline
                  class="mini-chart"
                  :points="hardwareHistory"
                />
                <div class="speed-values" aria-label="当前 CPU 和物理内存占比">
                  <strong class="hardware-cpu">{{ hardwareCpuLabel }}</strong>
                  <ATooltip :title="hardwareMemoryTooltip" placement="top">
                    <strong class="hardware-memory">{{ hardwareMemoryLabel }}</strong>
                  </ATooltip>
                </div>
              </div>
            </div>
          </article>
        </section>

        <aside class="right-column" aria-label="辅助信息">
          <section class="stats-card panel sprig-corner">
            <!-- 右上角插画使用裁剪后的独立图片，直接贴边显示。 -->
            <img class="corner-image stats-corner-image" src="/assets/watercolor-leaf-branch-a.png" alt="" />
            <div class="section-title">
              <span class="title-icon stats-title-frame" aria-hidden="true">
                <AppIcon icon="title-icon-glyph stats-title-icon fa-solid fa-chart-column" />
              </span>
              <h2>数据统计</h2>
            </div>
            <div class="stats-grid">
              <ACard
                v-for="item in stats"
                :key="item.label"
                :class="['stat-item', item.tone]"
                :bordered="false"
                hoverable
              >
                <AStatistic :title="item.label" :value="item.value" />
              </ACard>
            </div>
          </section>

          <section class="guide-card panel sprig-corner guide-corner">
            <!-- 右上角插画使用裁剪后的独立图片，避免伪元素负偏移造成遮挡。 -->
            <img class="corner-image guide-corner-image" src="/assets/watercolor-leaf-branch-b.png" alt="" />
            <div class="section-title">
              <span class="title-icon guide-title-frame" aria-hidden="true">
                <AppIcon icon="title-icon-glyph guide-title-icon fa-regular fa-bookmark" />
              </span>
              <h2>使用须知</h2>
            </div>
            <ACard class="notice info" :bordered="false" hoverable>
              <ACardMeta title="操作说明">
                <template #description>
                  <ul>
                    <li v-for="tip in usageTips" :key="tip">{{ tip }}</li>
                  </ul>
                </template>
              </ACardMeta>
            </ACard>
            <ACard class="notice warning" :bordered="false" hoverable>
              <ACardMeta title="合规提示">
                <template #description>
                  <ul>
                    <li v-for="tip in complianceTips" :key="tip">{{ tip }}</li>
                  </ul>
                </template>
              </ACardMeta>
            </ACard>
          </section>

          <section class="env-card panel">
            <div class="section-title title-with-pill">
              <span class="title-icon env-title-frame" aria-hidden="true">
                <span class="title-icon-glyph env-title-icon"></span>
              </span>
              <h2>运行环境</h2>
            </div>
            <div class="env-grid">
              <ACard v-for="item in envItems" :key="item.name" class="env-item" :bordered="false" hoverable>
                <div class="env-item-content">
                  <AppIcon :icon="['env-icon', item.icon]" />
                  <div class="env-item-text">
                    <strong>{{ item.name }}</strong>
                    <small>{{ item.value }}</small>
                  </div>
                </div>
              </ACard>
            </div>
          </section>
        </aside>

        <section class="log-card panel" aria-label="运行日志">
          <!-- 运行日志插画贴右下角显示，和日志内容保持清晰层级。 -->
          <img class="corner-image log-corner-image" src="/assets/watercolor-flower-branch.png" alt="" />
          <div class="log-header">
            <div class="section-title">
              <AppIcon icon="title-icon fa-regular fa-rectangle-list" />
              <h2>运行日志</h2>
            </div>
            <ASegmented
              class="log-tabs"
              aria-label="日志筛选"
              :value="activeLogLevel"
              :options="logSegmentOptions"
              @change="handleLogSegmentChange"
            />
            <div class="log-actions">
              <AButton class="log-folder-button" html-type="button" @click="handleOpenLogFolder">
                <AppIcon icon="fa-regular fa-folder-open" /> Open Log Folder
              </AButton>
            </div>
          </div>
          <div class="log-columns" aria-label="运行日志分栏">
            <section class="log-column" aria-label="软件活动和主流程日志">
              <div class="log-column-header">
                <h3>软件活动 / 主流程</h3>
                <span>{{ systemLogRows.length }} 条</span>
              </div>
              <DynamicScroller
                ref="systemLogScrollerRef"
                tabindex="0"
                class="log-table"
                :items="systemLogRows"
                key-field="id"
                :min-item-size="24"
                @scroll.passive="handleLogTableScroll"
                @pointerdown.passive="markLogScrollUserInteraction"
                @touchstart.passive="markLogScrollUserInteraction"
                @wheel.passive="markLogScrollUserInteraction"
                @keydown="markLogScrollUserInteraction"
              >
                <template #default="{ item: row, index, active }">
                  <DynamicScrollerItem
                    :item="row"
                    :active="active"
                    :data-index="index"
                    :size-dependencies="[row.message, row.levelLabel]"
                  >
                    <div :class="['log-row', `log-row-${row.levelClass}`]">
                      <time>[{{ row.time }}]</time>
                      <strong :class="['log-level', `log-level-${row.levelClass}`]">{{ row.levelLabel }}</strong>
                      <span class="log-message" :title="row.message">
                        <template v-for="(segment, segmentIndex) in row.messageSegments" :key="`${row.id}-${segmentIndex}`">
                          <span
                            v-if="segment.type === 'url'"
                            class="log-url"
                            :title="segment.fullText"
                          >{{ segment.text }}</span>
                          <span v-else>{{ segment.text }}</span>
                        </template>
                      </span>
                    </div>
                  </DynamicScrollerItem>
                </template>
                <template #empty>
                  <div class="log-empty">
                    暂无软件活动日志，程序开始运行后会显示主流程状态。
                  </div>
                </template>
              </DynamicScroller>
            </section>
            <section class="log-column" aria-label="单篇任务日志">
              <div class="log-column-header">
                <h3>单篇任务</h3>
                <span>{{ articleTaskLogRows.length }} 条</span>
              </div>
              <DynamicScroller
                ref="articleTaskLogScrollerRef"
                tabindex="0"
                class="log-table"
                :items="articleTaskLogRows"
                key-field="id"
                :min-item-size="24"
                @scroll.passive="handleLogTableScroll"
                @pointerdown.passive="markLogScrollUserInteraction"
                @touchstart.passive="markLogScrollUserInteraction"
                @wheel.passive="markLogScrollUserInteraction"
                @keydown="markLogScrollUserInteraction"
              >
                <template #default="{ item: row, index, active }">
                  <DynamicScrollerItem
                    :item="row"
                    :active="active"
                    :data-index="index"
                    :size-dependencies="[row.message, row.levelLabel]"
                  >
                    <div :class="['log-row', `log-row-${row.levelClass}`]">
                      <time>[{{ row.time }}]</time>
                      <strong :class="['log-level', `log-level-${row.levelClass}`]">{{ row.levelLabel }}</strong>
                      <span class="log-message" :title="row.message">
                        <template v-for="(segment, segmentIndex) in row.messageSegments" :key="`${row.id}-${segmentIndex}`">
                          <span
                            v-if="segment.type === 'url'"
                            class="log-url"
                            :title="segment.fullText"
                          >{{ segment.text }}</span>
                          <span v-else>{{ segment.text }}</span>
                        </template>
                      </span>
                    </div>
                  </DynamicScrollerItem>
                </template>
                <template #empty>
                  <div class="log-empty">
                    暂无单篇任务日志，主流程分发文章后会显示处理节点。
                  </div>
                </template>
              </DynamicScroller>
            </section>
          </div>
        </section>
        </template>

        <DataFilesPage
          v-else-if="activePage === 'files'"
          ref="dataFilesPageRef"
          class="management-area"
          :summary-stats="stats"
          @navigate="selectPage"
        />
        <HistoryPage v-else-if="activePage === 'history'" ref="historyPageRef" class="management-area" />
        <SettingsPage
          v-else
          ref="settingsPageRef"
          class="management-area"
          :environment-items="envItems"
          :task-status="taskStatus"
          @task-status-changed="handleTaskStatusChanged"
        />
      </div>
    </section>
    <TopbarHealthDialog
      :visible="healthDialogVisible"
      :title="`${healthDialogMeta.label}检测`"
      :icon="healthDialogMeta.icon"
      :tone="healthDialogState.tone"
      :status-label="healthDialogState.value"
      :message="healthDialogState.message"
      :items="healthDialogState.items"
      :checked-at="healthDialogState.checkedAt"
      :checking="healthDialogState.checking"
      @close="healthDialogVisible = false"
    />
    <TopbarHealthDialog
      :visible="startupSelfCheckDialogVisible"
      title="启动自检"
      icon="fa-solid fa-list-check"
      :tone="startupSelfCheckDialogState.tone"
      :status-label="startupSelfCheckDialogState.value"
      :message="startupSelfCheckDialogState.message"
      :items="startupSelfCheckDialogState.items"
      :checked-at="startupSelfCheckDialogState.checkedAt"
      :checking="startupSelfCheckDialogState.checking"
      @close="startupSelfCheckDialogVisible = false"
    />
    </main>
  </AConfigProvider>
</template>

<style scoped>
@font-face {
  font-family: 'HarmonyOS Sans SC';
  src: url('./assets/fonts/HarmonyOS_Sans_SC_Regular.ttf') format('truetype');
  font-display: swap;
  font-style: normal;
  font-weight: 400;
}

@font-face {
  font-family: 'HarmonyOS Sans SC';
  src: url('./assets/fonts/HarmonyOS_Sans_SC_Medium.ttf') format('truetype');
  font-display: swap;
  font-style: normal;
  font-weight: 500;
}

:global(*) {
  box-sizing: border-box;
}

:global(:root) {
  --design-width: 1600;
  --design-height: 900;
  --app-scale: min(
    calc(100vw / (var(--design-width) * 1px)),
    calc(100vh / (var(--design-height) * 1px))
  );
}

:global(body) {
  margin: 0;
  min-width: 320px;
  color: var(--ink);
  background: var(--page);
  font-family:
    'HarmonyOS Sans SC', 'Microsoft YaHei', 'PingFang SC', 'Segoe UI', system-ui, -apple-system,
    BlinkMacSystemFont, sans-serif;
  font-synthesis-weight: none;
}

:global(html),
:global(body),
:global(#app) {
  width: 100%;
  height: 100%;
  overflow: hidden;
}

button,
input {
  font: inherit;
}

button {
  cursor: pointer;
}

button:focus-visible,
input:focus-visible {
  outline: 3px solid rgba(48, 113, 204, 0.32);
  outline-offset: 3px;
}

.collector-app {
  --page: #eaf3fb;
  --paper: #fbfdff;
  --paper-soft: #f3f8fc;
  --frost-bg: rgba(252, 254, 255, 0.66);
  --frost-bg-strong: rgba(252, 254, 255, 0.78);
  --frost-bg-soft: rgba(241, 249, 252, 0.46);
  --frost-border: rgba(102, 145, 184, 0.32);
  --frost-highlight: rgba(255, 255, 255, 0.84);
  --frost-inner: rgba(89, 132, 174, 0.08);
  --frost-shadow: rgba(35, 69, 111, 0.12);
  --frost-blur: blur(18px) saturate(1.34);
  --paper-texture-opacity: 0.045;
  --paper-card-texture-opacity: 0.045;
  --paper-edge: rgba(102, 145, 184, 0.26);
  --paper-shadow-sm: 0 1px 2px rgba(28, 55, 82, 0.06);
  --paper-shadow-md: 0 4px 9px rgba(28, 55, 82, 0.06), 0 8px 14px rgba(28, 55, 82, 0.045);
  --paper-hover-shadow: 0 5px 10px rgba(28, 55, 82, 0.075), 0 10px 15px rgba(28, 55, 82, 0.05);
  --paper-fiber:
    radial-gradient(circle at 14% 18%, rgba(28, 55, 82, 0.2) 0 0.6px, transparent 0.9px),
    radial-gradient(circle at 82% 64%, rgba(28, 55, 82, 0.16) 0 0.5px, transparent 0.85px),
    radial-gradient(circle at 38% 78%, rgba(28, 55, 82, 0.1) 0 0.45px, transparent 0.8px),
    radial-gradient(circle at 66% 28%, rgba(255, 255, 255, 0.42) 0 0.55px, transparent 0.9px),
    linear-gradient(92deg, rgba(28, 55, 82, 0.14), transparent 34%, rgba(28, 55, 82, 0.1) 66%, transparent),
    linear-gradient(178deg, transparent 0%, rgba(28, 55, 82, 0.08) 48%, transparent 100%);
  --ink: #15386f;
  --ink-strong: #0c2d63;
  --ink-muted: #4d6c9f;
  --line: rgba(104, 141, 181, 0.3);
  --line-soft: rgba(104, 141, 181, 0.18);
  --blue: #2d75d6;
  --green: #1f8f69;
  --green-soft: #dff3e8;
  --red: #d9413f;
  --orange: #df7a35;
  --purple: #6651cc;
  --teal-wash: rgba(121, 182, 185, 0.28);
  display: flex;
  justify-content: center;
  align-items: flex-start;
  width: 100vw;
  height: 100vh;
  padding: 0;
  overflow: hidden;
  color: var(--ink);
  position: relative;
  isolation: isolate;
  background:
    radial-gradient(circle at 7% 6%, rgba(255, 255, 255, 0.96), transparent 24%),
    radial-gradient(circle at 94% 18%, rgba(210, 230, 246, 0.64), transparent 30%),
    linear-gradient(135deg, #f9fcff 0%, #eef6fd 54%, #f8fcff 100%);
}

.collector-app::before {
  content: '';
  position: fixed;
  inset: 0;
  z-index: 0;
  pointer-events: none;
  opacity: var(--paper-texture-opacity);
  background: var(--paper-fiber);
  mix-blend-mode: multiply;
}

.collector-app.dark {
  --page: #0f1726;
  --paper: #151f31;
  --paper-soft: #111a2b;
  --frost-bg: rgba(22, 32, 50, 0.9);
  --frost-bg-strong: rgba(25, 37, 58, 0.96);
  --frost-bg-soft: rgba(16, 25, 41, 0.86);
  --frost-border: rgba(128, 153, 188, 0.2);
  --frost-highlight: rgba(214, 226, 244, 0.055);
  --frost-inner: rgba(123, 149, 184, 0.055);
  --frost-shadow: rgba(0, 0, 0, 0.2);
  --frost-blur: blur(10px) saturate(1.04);
  --paper-texture-opacity: 0.024;
  --paper-card-texture-opacity: 0.022;
  --paper-edge: rgba(128, 153, 188, 0.18);
  --paper-shadow-sm: 0 1px 2px rgba(0, 0, 0, 0.16);
  --paper-shadow-md: 0 4px 9px rgba(0, 0, 0, 0.18), 0 8px 14px rgba(0, 0, 0, 0.16);
  --paper-hover-shadow: 0 5px 10px rgba(0, 0, 0, 0.22), 0 10px 15px rgba(0, 0, 0, 0.18);
  --paper-fiber:
    radial-gradient(circle at 14% 18%, rgba(214, 226, 244, 0.2) 0 0.6px, transparent 0.9px),
    radial-gradient(circle at 82% 64%, rgba(214, 226, 244, 0.14) 0 0.5px, transparent 0.85px),
    radial-gradient(circle at 38% 78%, rgba(214, 226, 244, 0.1) 0 0.45px, transparent 0.8px),
    linear-gradient(92deg, rgba(214, 226, 244, 0.12), transparent 34%, rgba(214, 226, 244, 0.08) 66%, transparent);
  --ink: #cbd8ea;
  --ink-strong: #e6eef8;
  --ink-muted: #8ea2bd;
  --line: rgba(128, 153, 188, 0.2);
  --line-soft: rgba(128, 153, 188, 0.1);
  --green-soft: rgba(64, 142, 111, 0.16);
  background:
    radial-gradient(circle at 8% 0%, rgba(53, 87, 135, 0.16), transparent 32%),
    radial-gradient(circle at 92% 12%, rgba(69, 89, 124, 0.12), transparent 34%),
    linear-gradient(135deg, #0d1421 0%, #111b2c 52%, #0e1726 100%);
}

.collector-app.dark::before {
  opacity: var(--paper-texture-opacity);
  background: var(--paper-fiber);
  mix-blend-mode: screen;
}

.workspace {
  position: relative;
  z-index: 1;
  width: 1600px;
  height: 900px;
  padding: 14px 16px;
  flex: 0 0 auto;
  transform: scale(var(--app-scale));
  transform-origin: top center;
  overflow: hidden;
}

.dark .sidebar-book,
.dark .corner-image {
  opacity: 0.3;
  mix-blend-mode: normal;
  filter: saturate(0.62) brightness(0.72) contrast(0.88);
}

.dark .sidebar-book {
  opacity: 0.24;
}

.dashboard {
  --section-gap: 14px;
  display: grid;
  grid-template-columns: 240px 430px 450px 400px;
  grid-template-rows: 462px 278px;
  grid-template-areas:
    'sidebar tasks status right'
    'sidebar logs logs right';
  gap: var(--section-gap) 16px;
  align-items: start;
  height: 754px;
}

.panel {
  position: relative;
  isolation: isolate;
  border: 1px solid var(--paper-edge);
  border-radius: 8px;
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm), var(--paper-shadow-md);
  overflow: hidden;
}

.panel::before {
  content: '';
  position: absolute;
  inset: 0;
  z-index: 0;
  pointer-events: none;
  background: var(--paper-fiber);
  opacity: var(--paper-card-texture-opacity);
}

.panel > * {
  position: relative;
  z-index: 1;
}

.dark .panel {
  border-color: var(--paper-edge);
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm), var(--paper-shadow-md);
}

.dark .panel::before {
  background: var(--paper-fiber);
  opacity: var(--paper-card-texture-opacity);
}

.sidebar {
  grid-area: sidebar;
  position: relative;
  height: 754px;
  padding: 30px 20px 18px;
}

.nav-group {
  position: relative;
  z-index: 2;
}

.nav-group + .nav-group {
  margin-top: 32px;
  padding-top: 28px;
  border-top: 1px dashed var(--line);
}

.nav-title {
  margin: 0 0 18px 12px;
  color: var(--ink-strong);
  font-size: 18px;
  font-weight: 500;
}

.nav-title span {
  color: var(--green);
}

.nav-item {
  display: flex;
  align-items: center;
  gap: 16px;
  width: 100%;
  min-height: 54px;
  margin-bottom: 12px;
  padding: 0 18px;
  border: 0;
  border-radius: 999px;
  color: var(--ink);
  background: transparent;
  font-size: 18px;
  font-weight: 500;
  text-align: left;
  text-decoration: none;
}

.nav-item:hover {
  background: rgba(74, 129, 183, 0.08);
}

.nav-item:disabled {
  cursor: not-allowed;
  opacity: 0.58;
}

.nav-item.active {
  color: #ffffff;
  background:
    radial-gradient(circle at 20% 10%, #ffffff4d, transparent 36%),
    linear-gradient(135deg, #5bbdc5eb, #2b8ea6eb);
  box-shadow: var(--paper-shadow-sm), inset 0 1px 0 rgba(255, 255, 255, 0.18);
}

.dark .nav-item.active {
  color: #eef6ff;
  background: #2b7f92;
}

.nav-icon {
  display: inline-grid;
  place-items: center;
  width: 32px;
  color: currentColor;
  font-size: 24px;
  line-height: 1;
}

.sidebar-book {
  position: absolute;
  left: 0;
  right: 0;
  bottom: -2px;
  z-index: 1;
  width: 100%;
  height: auto;
  object-fit: contain;
  object-position: center bottom;
  opacity: 0.72;
  pointer-events: none;
  mix-blend-mode: multiply;
  filter: saturate(0.92) contrast(0.96);
}

.task-column {
  grid-area: tasks;
  display: grid;
  gap: var(--section-gap);
}

.status-column {
  grid-area: status;
}

.right-column {
  grid-area: right;
  display: grid;
  grid-template-rows: 170px 344px 212px;
  gap: var(--section-gap);
  height: 754px;
}

.management-area {
  grid-column: 2 / 5;
  grid-row: 1 / 3;
  min-width: 0;
}

.task-card {
  display: grid;
  grid-template-columns: 118px minmax(0, 1fr);
  gap: 16px;
  height: 182px;
  padding: 18px 20px 16px;
}

.task-card-content {
  grid-template-columns: 118px minmax(0, 1fr);
  gap: 16px;
}

.task-volume-body {
  align-self: stretch;
  display: grid;
  grid-template-rows: auto repeat(3, 32px);
  align-content: center;
  gap: 7px;
  padding-top: 1px;
  min-width: 0;
}

.task-volume-body > h2 {
  margin-bottom: 3px;
}

.task-volume-row {
  display: grid;
  grid-template-columns: 88px minmax(0, 1fr);
  align-items: center;
  gap: 8px;
  min-width: 0;
}

.task-control-label {
  color: var(--ink-strong);
  font-size: 14px;
  font-weight: 500;
  line-height: 1.25;
  white-space: nowrap;
}

.task-label-with-hint {
  display: flex;
  align-items: center;
  gap: 6px;
  min-width: 0;
  padding-left: 8px;
}

.task-date-filter-trigger {
  position: relative;
  display: flex;
  align-items: center;
  justify-content: space-between;
  width: 100%;
  min-width: 0;
  height: 32px;
  padding: 0 22px 0 4px;
  border-color: var(--paper-edge);
  border-radius: 6px;
  color: var(--ink-strong);
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm);
  font-size: 14px;
  font-weight: 400;
  text-align: left;
  overflow: hidden;
}

.task-date-filter-trigger > span:first-child {
  display: block;
  width: 100%;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.task-date-filter-chevron {
  position: absolute;
  top: 50%;
  right: 8px;
  flex: 0 0 auto;
  color: var(--ink-muted);
  font-size: 8px;
  transform: translateY(-50%) rotate(0deg);
  transition: transform 160ms ease;
}

.task-date-filter-chevron.is-open {
  transform: translateY(-50%) rotate(180deg);
}

.task-count-control {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 7px;
  width: 100%;
  min-width: 0;
}

.task-count-input {
  flex: 1 1 auto;
  width: 100%;
  min-width: 0;
  max-width: none;
  height: 32px;
  border-color: var(--paper-edge);
  border-radius: 6px;
  color: var(--ink-strong);
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm);
  font-size: 14px;
  font-weight: 400;
}

.task-count-input:hover,
.task-count-input:focus-within {
  border-color: rgba(53, 127, 217, 0.56);
}

.task-count-input :deep(.ant-input-number-input-wrap) {
  height: 100%;
}

.task-count-input :deep(.ant-input-number-input) {
  height: 30px;
  padding: 0 36px 0 10px;
  color: var(--ink-strong);
  font-size: 14px;
  font-weight: 400;
  text-align: center;
}

.task-count-input :deep(.ant-input-number-handler-wrap) {
  width: 30px;
  border-inline-start-color: var(--line-soft);
  border-radius: 0 5px 5px 0;
  background: rgba(240, 246, 251, 0.92);
  opacity: 1;
}

.task-count-input :deep(.ant-input-number-handler) {
  color: var(--ink-muted);
}

.task-count-input :deep(.ant-input-number-handler:hover) {
  color: var(--blue);
}

.dark .task-count-input {
  border-color: var(--paper-edge);
  color: #e1eaf6;
  background: var(--frost-bg-strong);
}

.dark .task-count-input :deep(.ant-input-number-input) {
  color: #e1eaf6;
}

.dark .task-count-input :deep(.ant-input-number-handler-wrap) {
  border-inline-start-color: rgba(128, 153, 188, 0.22);
  background: rgba(28, 43, 65, 0.64);
}

.task-label-hint {
  display: grid;
  place-items: center;
  flex: 0 0 20px;
  width: 20px;
  height: 20px;
  padding: 0;
  border: 0;
  border-radius: 50%;
  color: #397fa8;
  background: transparent;
  box-shadow: none;
  font-size: 14px;
  cursor: help;
}

.task-label-hint:hover,
.task-label-hint:focus-visible {
  color: #286b94;
  background: transparent;
}

.task-label-hint:focus-visible {
  outline: 2px solid rgba(57, 127, 168, 0.24);
  outline-offset: 2px;
}

.task-date-row {
  position: relative;
  grid-template-columns: minmax(0, 1fr);
  width: 100%;
  transform: none;
}

.task-date-row > .task-control-label {
  position: absolute;
  top: 50%;
  right: calc(100% + 12px);
  transform: translateY(-50%);
}

.task-date-control {
  display: block;
  width: 100%;
  min-width: 0;
  height: 32px;
}

.task-date-picker {
  width: 100%;
  height: 32px;
}

.task-date-control :deep(.ant-picker) {
  width: 100%;
  height: 32px;
  border-color: var(--paper-edge);
  border-radius: 6px;
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm);
}

.task-date-control :deep(.ant-picker-input > input) {
  color: var(--ink-strong);
  font-size: 14px;
}

.task-date-control :deep(.ant-picker-input > input::placeholder) {
  color: var(--ink-muted);
  opacity: 1;
}

.task-date-control :deep(.ant-picker-range .ant-picker-input) {
  flex: 1 1 0;
  min-width: 0;
}

.task-date-control :deep(.ant-picker-range-separator) {
  display: grid;
  place-items: center;
  flex: 0 0 28px;
  padding: 0;
}

.dark .task-label-hint {
  color: #9bbbd7;
  background: transparent;
}

:global(.task-date-filter-dropdown.ant-dropdown .ant-dropdown-menu) {
  min-width: calc(180px * var(--app-scale));
  padding: calc(4px * var(--app-scale));
  border-radius: calc(6px * var(--app-scale));
}

:global(.task-date-filter-dropdown.ant-dropdown .ant-dropdown-menu .ant-dropdown-menu-item) {
  min-height: calc(32px * var(--app-scale));
  padding: calc(5px * var(--app-scale)) calc(12px * var(--app-scale));
  border-radius: calc(4px * var(--app-scale));
  font-size: calc(14px * var(--app-scale));
  font-weight: 400;
  line-height: 1.25;
}

:global(.task-date-picker-panel .ant-picker-panel-container) {
  /* 日历弹层挂载在 body，使用页面比例同步缩放其完整内部结构。 */
  zoom: var(--app-scale);
}

.step-badge {
  position: absolute;
  top: 16px;
  left: 22px;
  z-index: 1;
  display: grid;
  place-items: center;
  width: 44px;
  height: 44px;
  border-radius: 50%;
  color: #ffffff;
  background:
    radial-gradient(circle at 28% 20%, rgba(255, 255, 255, 0.5), transparent 35%),
    linear-gradient(135deg, #88bfd7, #508fb4);
  font-size: 21px;
  font-weight: 500;
}

.dark .step-badge {
  color: #e3edf8;
  background:
    radial-gradient(circle at 28% 20%, rgba(232, 242, 255, 0.18), transparent 35%),
    linear-gradient(135deg, #4479a0, #315d86);
  box-shadow:
    inset 0 1px 0 rgba(232, 242, 255, 0.16),
    0 0 10px rgba(61, 120, 177, 0.12);
}

.task-art {
  align-self: center;
  justify-self: center;
  width: 112px;
  height: 112px;
  margin-top: 24px;
  object-fit: contain;
  pointer-events: none;
}

.task-art-list {
  transform: translate(25px, -15px);
}

.dark .task-art {
  opacity: 0.66;
  filter: saturate(0.64) brightness(0.76) contrast(0.9);
}

.task-art-heart {
  width: 108px;
  height: 108px;
}

.task-art-column {
  align-self: stretch;
  justify-self: center;
  display: grid;
  grid-template-rows: auto auto;
  align-content: end;
  justify-items: center;
  gap: 0;
  width: 100%;
  min-width: 0;
}

.task-card-content .task-art {
  align-self: end;
  margin-top: 0;
}

.task-card-content .task-art-heart {
  width: 112px;
  height: 112px;
  transform: translate(-14px, 12px);
}

.task-body {
  align-self: center;
  min-width: 0;
}

.task-body h2,
.task-volume-body h2,
.section-title h2 {
  margin: 0;
  color: var(--ink-strong);
  font-size: 22px;
  line-height: 1.2;
  font-weight: 500;
}

.task-body p {
  margin: 12px 0 0;
  color: var(--ink);
  font-size: 15px;
  font-weight: 400;
}

.task-body-content {
  display: grid;
  align-content: start;
  gap: 10px;
  padding-top: 1px;
}

.field-row {
  display: grid;
  grid-template-columns: 1fr;
  gap: 10px;
  align-items: start;
  margin-top: 18px;
  color: var(--ink);
  font-weight: 400;
}

.field-row label {
  line-height: 1.2;
}

.download-options {
  justify-self: end;
  display: grid;
  gap: 7px;
  width: min(100%, 190px);
}

.download-option {
  display: flex;
  align-items: center;
  gap: 0;
  width: 100%;
  min-height: 32px;
  padding: 5px 10px;
  border: 1px solid rgba(104, 141, 181, 0.2);
  border-radius: 9px;
  color: var(--ink);
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm);
  text-align: left;
  font-size: 14px;
  font-weight: 400;
}

.download-option :deep(.ant-checkbox) {
  flex: 0 0 auto;
}

.download-option :deep(.ant-checkbox + span) {
  min-width: 0;
  padding-inline-start: 10px;
  padding-inline-end: 0;
}

.download-option-combo {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 28px;
  align-items: center;
  gap: 4px;
  padding-inline-end: 4px;
}

.offline-archive-checkbox {
  min-width: 0;
  color: inherit;
  font-size: inherit;
}

.offline-archive-checkbox :deep(span:last-child) {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.offline-archive-mode-trigger {
  display: grid;
  place-items: center;
  width: 28px;
  min-width: 28px;
  height: 28px;
  padding: 0;
  border: 0;
  border-radius: 6px;
  color: rgba(53, 127, 217, 0.86);
  background: transparent;
  box-shadow: none;
}

.offline-archive-mode-trigger:hover,
.offline-archive-mode-trigger:focus-visible {
  color: rgba(35, 101, 184, 0.96);
  background: rgba(217, 236, 255, 0.52);
}

.download-select-chevron {
  font-size: 11px;
  transition: transform 160ms ease;
}

.download-select-chevron.is-open {
  transform: rotate(180deg);
}

:global(.offline-archive-mode-dropdown.ant-dropdown .ant-dropdown-menu) {
  min-width: calc(168px * var(--app-scale));
  padding: calc(4px * var(--app-scale));
  border-radius: calc(6px * var(--app-scale));
}

:global(.offline-archive-mode-dropdown.ant-dropdown .ant-dropdown-menu .ant-dropdown-menu-item) {
  min-height: calc(32px * var(--app-scale));
  padding: calc(5px * var(--app-scale)) calc(12px * var(--app-scale));
  border-radius: calc(4px * var(--app-scale));
  font-size: calc(14px * var(--app-scale));
  font-weight: 400;
  line-height: 1.25;
}

.download-option:hover {
  background: rgba(233, 246, 247, 0.58);
}

.download-option.locked {
  border-color: rgba(111, 130, 155, 0.22);
  color: rgba(77, 108, 159, 0.74);
  background: rgba(230, 237, 245, 0.58);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.5);
  cursor: not-allowed;
}

.article-detail-option {
  width: 184px;
  min-height: 32px;
  padding-inline: 10px;
}

.task-date-label {
  justify-content: flex-start;
  gap: 8px;
  width: 120px;
  height: 32px;
  min-height: 32px;
  padding-inline: 10px;
  border-radius: 6px;
  cursor: default;
  pointer-events: none;
}

.task-date-label-spin {
  flex: 0 0 14px;
  width: 14px;
  height: 14px;
  color: rgba(77, 108, 159, 0.82);
  line-height: 1;
}

.dark .task-date-label-spin {
  color: rgba(153, 183, 219, 0.86);
}

.download-option.locked:hover {
  background: rgba(230, 237, 245, 0.58);
}

.download-option.selected {
  border-color: rgba(48, 132, 128, 0.32);
  color: var(--ink-strong);
  background: rgba(219, 241, 235, 0.64);
}

.download-option.selected.locked {
  border-color: rgba(111, 130, 155, 0.24);
  color: rgba(69, 91, 124, 0.78);
  background: rgba(230, 237, 245, 0.68);
}

.dark .download-option {
  border-color: rgba(117, 148, 187, 0.22);
  color: #b9c8dd;
  background: rgba(17, 27, 44, 0.66);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.04);
}

.dark .download-option:hover {
  background: rgba(24, 38, 60, 0.78);
}

.dark .download-option.locked,
.dark .download-option.locked:hover {
  border-color: rgba(117, 148, 187, 0.16);
  color: rgba(154, 171, 196, 0.62);
  background: rgba(24, 34, 50, 0.52);
}

.dark .download-option.selected {
  border-color: rgba(95, 151, 180, 0.36);
  color: #d7e7f2;
  background: rgba(35, 66, 84, 0.72);
}

.dark .download-option.selected.locked {
  border-color: rgba(117, 148, 187, 0.2);
  color: rgba(178, 190, 210, 0.68);
  background: rgba(31, 43, 59, 0.62);
}

.dark .offline-archive-mode-trigger {
  color: rgba(170, 204, 235, 0.88);
}

.dark .offline-archive-mode-trigger:hover,
.dark .offline-archive-mode-trigger:focus-visible {
  color: #d7e7f2;
  background: rgba(61, 96, 136, 0.36);
}

.control-panel {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 18px;
  height: 70px;
  padding: 10px 18px;
}

.run-button,
.stop-button {
  position: relative;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 14px;
  height: 50px;
  min-height: 50px;
  padding: 0 18px;
  border: 1px solid rgba(49, 92, 92, 0.18);
  border-radius: 14px;
  color: #ffffff;
  font-size: 20px;
  font-weight: 500;
  overflow: hidden;
  box-shadow: var(--paper-shadow-sm), var(--paper-shadow-md);
  text-shadow: none;
  transition:
    transform 160ms ease,
    filter 160ms ease,
    box-shadow 160ms ease,
    border-color 160ms ease;
}

.run-button::before,
.stop-button::before {
  display: none;
}

.run-button::after,
.stop-button::after {
  content: '';
  position: absolute;
  inset: 0;
  z-index: 0;
  pointer-events: none;
  background: var(--paper-fiber);
  mix-blend-mode: soft-light;
  opacity: 0.08;
}

.run-button i,
.stop-button i,
.button-label {
  position: relative;
  z-index: 1;
}

.run-button i,
.stop-button i {
  width: 18px;
  font-size: 17px;
  text-align: center;
}

.run-button {
  isolation: isolate;
}

.stop-button {
  isolation: isolate;
}

.run-button:hover,
.stop-button:hover {
  filter: brightness(1.03) saturate(1.02);
  transform: translateY(-1px);
  border-color: rgba(49, 92, 92, 0.24);
  box-shadow: var(--paper-shadow-sm), var(--paper-hover-shadow);
}

.run-button:disabled,
.stop-button:disabled {
  cursor: not-allowed;
  pointer-events: none;
  opacity: 0.46;
  filter: grayscale(0.72) saturate(0.55);
  transform: none;
  border-color: rgba(112, 130, 155, 0.2);
  box-shadow:
    inset 0 1px 0 rgba(255, 255, 255, 0.36),
    0 2px 5px rgba(35, 69, 111, 0.06);
  text-shadow: none;
}

.run-button:disabled:hover,
.stop-button:disabled:hover,
.run-button:disabled:active,
.stop-button:disabled:active {
  filter: grayscale(0.72) saturate(0.55);
  transform: none;
  border-color: rgba(112, 130, 155, 0.2);
  box-shadow:
    inset 0 1px 0 rgba(255, 255, 255, 0.36),
    0 2px 5px rgba(35, 69, 111, 0.06);
}

.run-button:active,
.stop-button:active {
  filter: brightness(0.98);
  transform: translateY(1px);
  box-shadow:
    inset 0 1px 0 rgba(255, 255, 255, 0.4),
    inset 0 2px 7px rgba(35, 69, 111, 0.11),
    0 4px 8px rgba(38, 70, 116, 0.08);
}

.run-button {
  background: #4aae9f;
}

.stop-button {
  background: #c9655f;
}

.dark .run-button,
.dark .stop-button {
  border-color: rgba(142, 174, 210, 0.28);
  box-shadow: var(--paper-shadow-sm), var(--paper-shadow-md);
  text-shadow: none;
}

.dark .run-button {
  background: #2f8d82;
}

.dark .stop-button {
  background: #864f53;
}

.dark .run-button::after,
.dark .stop-button::after {
  opacity: 0.08;
  mix-blend-mode: normal;
  background: var(--paper-fiber);
}

.dark .run-button:hover,
.dark .stop-button:hover {
  filter: brightness(1.03) saturate(1.02);
  border-color: rgba(153, 188, 225, 0.36);
  box-shadow:
    inset 0 1px 0 rgba(232, 242, 255, 0.12),
    0 0 0 1px rgba(67, 116, 184, 0.1),
    0 7px 12px rgba(0, 0, 0, 0.18);
}

.dark .run-button:disabled,
.dark .stop-button:disabled {
  opacity: 0.52;
  color: #8f9db1;
  border-color: rgba(117, 148, 187, 0.16);
  background: rgba(38, 46, 61, 0.78);
  box-shadow:
    inset 0 1px 0 rgba(232, 242, 255, 0.045),
    0 2px 5px rgba(0, 0, 0, 0.12);
}

.dark .run-button:disabled:hover,
.dark .stop-button:disabled:hover,
.dark .run-button:disabled:active,
.dark .stop-button:disabled:active {
  border-color: rgba(117, 148, 187, 0.16);
  box-shadow:
    inset 0 1px 0 rgba(232, 242, 255, 0.045),
    0 2px 5px rgba(0, 0, 0, 0.12);
}

.status-card {
  display: grid;
  grid-template-rows: auto minmax(0, 1fr) auto;
  height: 462px;
  padding: 20px 16px 18px;
}

.corner-image {
  position: absolute;
  opacity: 0.68;
  pointer-events: none;
  object-fit: contain;
  z-index: 1;
  mix-blend-mode: multiply;
  filter: saturate(0.9) contrast(0.95);
}

.status-corner-image,
.stats-corner-image,
.guide-corner-image {
  top: 0;
  right: 0;
  height: auto;
  object-position: top right;
}

.status-corner-image {
  width: 172px;
}

.stats-corner-image {
  width: 118px;
}

.guide-corner-image {
  width: 122px;
}

.section-title {
  display: flex;
  align-items: center;
  gap: 12px;
  min-width: 0;
}

.title-icon {
  display: inline-grid;
  place-items: center;
  width: 30px;
  height: 30px;
  flex: 0 0 auto;
  color: var(--blue);
  font-size: 22px;
  line-height: 1;
}

.stats-card .title-icon,
.guide-card .title-icon,
.env-card .title-icon {
  display: inline-grid;
  place-items: center;
  width: 34px;
  height: 34px;
  flex: 0 0 34px;
  line-height: 1;
}

.title-icon-glyph {
  display: block;
  color: var(--blue);
  line-height: 1;
}

.stats-title-icon {
  width: 30px;
  height: 30px;
  font-size: 30px;
}

.guide-title-icon {
  width: 31px;
  height: 31px;
  font-size: 31px;
  transform: translateY(1px);
}

.env-title-icon {
  width: 31px;
  height: 31px;
  background: var(--blue);
  -webkit-mask: url('/assets/computer-solid-full.svg') center / contain no-repeat;
  mask: url('/assets/computer-solid-full.svg') center / contain no-repeat;
}

.status-title-logo {
  display: block;
  width: 38px;
  height: 30px;
}

.status-list {
  display: grid;
  align-content: space-between;
  gap: 12px;
  min-height: 0;
  margin-top: 18px;
  padding-bottom: 18px;
  border-bottom: 1px dashed var(--line);
}

.status-row {
  display: grid;
  grid-template-columns: 18px 96px minmax(0, 1fr);
  gap: 6px;
  align-items: center;
  min-width: 0;
  color: var(--ink);
  font-size: 14px;
}

.status-row.with-progress {
  grid-template-columns: 18px 96px minmax(0, 1fr) max-content;
}

.row-icon {
  width: 18px;
  color: var(--blue);
  font-size: 14px;
  line-height: 1;
  text-align: center;
}

.status-label {
  color: var(--ink-muted);
  font-weight: 400;
  white-space: nowrap;
}

.status-row strong {
  color: var(--ink-strong);
  font-weight: 500;
}

.status-value {
  min-width: 0;
  width: 100%;
}

.status-value-ellipsis {
  display: block;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.status-value.status-description-value {
  cursor: help;
}

.status-description-value:focus-visible {
  outline: 3px solid rgba(45, 117, 214, 0.22);
  outline-offset: 2px;
  border-radius: 5px;
}

.status-pill {
  display: inline-flex;
  align-items: center;
  justify-self: start;
  min-height: 24px;
  margin-inline-end: 0;
  padding: 3px 8px;
  border: 0;
  border-radius: 6px;
  color: var(--green);
  background: var(--green-soft);
  font-size: 13px;
  font-weight: 500;
  white-space: nowrap;
}

.status-pill.orange {
  color: #b46b18;
  background: rgba(255, 236, 190, 0.72);
}

.status-pill.blue {
  color: #1d63bd;
  background: rgba(222, 237, 255, 0.78);
}

.status-pill.red {
  color: #b92f32;
  background: rgba(255, 226, 226, 0.78);
}

.status-pill.purple {
  color: #5f46c4;
  background: rgba(235, 230, 255, 0.78);
}

.progress-line {
  display: block;
  min-width: 0;
  line-height: 1;
}

.progress-line :deep(.ant-progress-outer) {
  display: block;
}

.progress-line :deep(.ant-progress-inner) {
  height: 8px;
  border-radius: 999px;
  background: rgba(53, 127, 217, 0.14);
}

.progress-line :deep(.ant-progress-bg) {
  height: 8px !important;
  border-radius: inherit;
}

.status-row em {
  color: var(--ink-muted);
  font-style: normal;
  font-weight: 500;
  white-space: nowrap;
}

.network-panel {
  display: grid;
  grid-template-columns: 118px minmax(0, 1fr);
  column-gap: 10px;
  align-items: center;
  min-height: 82px;
  padding-top: 18px;
}

.hardware-labels {
  display: grid;
  grid-template-rows: repeat(2, 1fr);
  align-items: center;
  height: 62px;
  font-size: 13px;
  font-weight: 400;
  line-height: 1.2;
  white-space: nowrap;
}

.hardware-label-row {
  display: grid;
  grid-template-columns: 16px 88px;
  align-items: center;
  gap: 7px;
  min-width: 0;
}

.hardware-label-icon {
  width: 16px;
  color: var(--blue);
  font-size: 14px;
  text-align: center;
}

.hardware-label-box {
  display: block;
  width: 88px;
}

.hardware-label-text {
  display: block;
  width: 100%;
  color: var(--ink-muted);
  font-weight: 400;
  text-align: justify;
  text-align-last: justify;
  text-justify: inter-character;
}

.network-panel strong {
  display: inline-flex;
  line-height: 1.2;
  font-size: 15px;
  white-space: nowrap;
}

.hardware-cpu {
  color: var(--green);
}

.hardware-memory {
  color: var(--blue);
}

.speed-content {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 54px;
  align-items: center;
  gap: 10px;
  min-width: 0;
}

.speed-values {
  display: grid;
  grid-template-rows: repeat(2, 1fr);
  gap: 0;
  align-items: center;
  justify-items: start;
  height: 62px;
  min-width: 0;
  white-space: nowrap;
}

.mini-chart {
  align-self: center;
  width: 100%;
  min-width: 0;
  height: 62px;
  transform: translateY(7px);
}

.stats-card,
.guide-card,
.env-card {
  padding: 18px 24px;
}

.stats-card {
  --blue: #2378d7;
  --green: #168a78;
  --purple: #6d56d7;
  height: 100%;
  padding-top: 16px;
  padding-bottom: 14px;
}

.stats-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 6px 24px;
  margin-top: 8px;
}

.stat-item {
  display: block;
  min-height: 40px;
  border-radius: 6px;
  background: transparent;
  box-shadow: none;
  cursor: default;
}

.stat-item:nth-child(odd) {
  border-right: 1px dashed var(--line);
}

.stat-item :deep(.ant-card-body) {
  display: grid;
  gap: 3px;
  min-height: 40px;
  padding: 0;
}

.stat-item :deep(.ant-statistic-title) {
  margin: 0;
  color: var(--ink-muted);
  font-size: 15px;
  font-weight: 400;
  line-height: 1.2;
}

.stat-item :deep(.ant-statistic-content) {
  color: var(--blue);
  font-size: 23px;
  line-height: 1.1;
  font-weight: 400;
}

.stat-item.green :deep(.ant-statistic-content) {
  color: var(--green);
}

.stat-item.purple :deep(.ant-statistic-content) {
  color: var(--purple);
}

.stat-item.orange :deep(.ant-statistic-content) {
  color: var(--orange);
}

.stat-item.red :deep(.ant-statistic-content) {
  color: var(--red);
}

.guide-card {
  display: grid;
  gap: 6px;
  height: 100%;
}

.notice {
  border: 1px solid rgba(104, 141, 181, 0.2);
  border-radius: 8px;
  box-shadow: var(--paper-shadow-sm);
  overflow: hidden;
  cursor: default;
  transition:
    transform 180ms ease,
    border-color 180ms ease,
    box-shadow 180ms ease,
    background 180ms ease;
}

.notice.ant-card-hoverable:hover {
  border-color: rgba(45, 111, 168, 0.28);
  box-shadow:
    0 5px 14px rgba(28, 55, 82, 0.12),
    0 12px 24px rgba(28, 55, 82, 0.1);
  transform: translateY(-2px);
}

.notice :deep(.ant-card-body) {
  padding: 12px 16px;
}

.notice :deep(.ant-card-meta-title) {
  margin: 0 0 6px;
  color: var(--blue);
  font-size: 16px;
  font-weight: 500;
  line-height: 1.25;
}

.notice :deep(.ant-card-meta-description) {
  color: inherit;
}

.notice.info {
  margin-top: 0;
  background: rgba(225, 240, 250, 0.58);
}

.notice.warning {
  background: rgba(255, 239, 219, 0.6);
}

.notice.warning :deep(.ant-card-meta-title) {
  color: var(--orange);
}

.notice ul {
  margin: 0;
  padding-left: 20px;
  color: var(--ink);
  line-height: 1.56;
  font-size: 14px;
  font-weight: 400;
}

.dark .notice {
  border-color: rgba(117, 148, 187, 0.18);
  box-shadow: var(--paper-shadow-sm);
}

.dark .notice.ant-card-hoverable:hover {
  border-color: rgba(130, 166, 210, 0.34);
  box-shadow:
    0 6px 16px rgba(0, 0, 0, 0.24),
    0 14px 26px rgba(0, 0, 0, 0.18);
}

.dark .notice.info {
  background: rgba(23, 45, 68, 0.68);
}

.dark .notice.warning {
  background: rgba(70, 57, 48, 0.62);
}

.dark .notice :deep(.ant-card-meta-title) {
  color: #76aef4;
}

.dark .notice.warning :deep(.ant-card-meta-title) {
  color: #d79a5e;
}

.dark .notice ul {
  color: #c3d0e2;
}

.title-with-pill {
  justify-content: flex-start;
}

.env-card {
  height: 100%;
  padding: 14px 18px;
}

.env-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 10px 8px;
  margin-top: 10px;
}

.env-item {
  display: block;
  min-height: 62px;
  min-width: 0;
  border: 1px solid rgba(104, 141, 181, 0.18);
  border-radius: 6px;
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm);
  overflow: hidden;
  cursor: default;
  transition:
    transform 180ms ease,
    border-color 180ms ease,
    box-shadow 180ms ease,
    background 180ms ease;
}

.env-item.ant-card-hoverable:hover {
  border-color: rgba(45, 111, 168, 0.28);
  box-shadow:
    0 5px 14px rgba(28, 55, 82, 0.12),
    0 12px 24px rgba(28, 55, 82, 0.1);
  transform: translateY(-2px);
}

.dark .env-item {
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm);
}

.dark .env-item.ant-card-hoverable:hover {
  border-color: rgba(130, 166, 210, 0.34);
  box-shadow:
    0 6px 16px rgba(0, 0, 0, 0.24),
    0 14px 26px rgba(0, 0, 0, 0.18);
}

.env-item :deep(.ant-card-body) {
  display: flex;
  align-items: center;
  justify-content: flex-start;
  min-height: 62px;
  padding: 8px 7px;
}

.env-item-content {
  display: grid;
  grid-template-columns: 20px minmax(0, 1fr);
  align-items: center;
  justify-items: start;
  gap: 6px;
  width: 100%;
  min-width: 0;
  text-align: left;
}

.env-item-text {
  display: grid;
  align-content: center;
  justify-items: start;
  gap: 2px;
  min-width: 0;
}

.env-icon {
  color: var(--blue);
  font-size: 17px;
  line-height: 1;
  text-align: center;
}

.env-item strong {
  min-width: 0;
  color: var(--ink-strong);
  font-size: 13px;
  font-weight: 500;
  white-space: nowrap;
}

.env-item small {
  min-width: 0;
  color: var(--ink-muted);
  font-size: 11.5px;
  font-weight: 500;
  white-space: nowrap;
}

.log-card {
  grid-area: logs;
  height: 278px;
  padding: 18px 22px;
}

.log-card::after {
  content: none;
}

.log-corner-image {
  right: 0;
  bottom: 0;
  z-index: 0;
  width: 268px;
  height: auto;
  opacity: 0.46;
  object-position: right bottom;
}

.dark .log-corner-image {
  opacity: 0.3;
}

.log-header {
  position: relative;
  z-index: 1;
  display: grid;
  grid-template-columns: auto 1fr auto;
  gap: 18px;
  align-items: center;
  padding-bottom: 8px;
  border-bottom: 1px solid var(--line);
}

.log-tabs,
.log-actions {
  display: flex;
  align-items: center;
  min-width: 0;
}

.log-tabs {
  justify-self: start;
  padding: 2px;
  border: 1px solid rgba(104, 141, 181, 0.2);
  border-radius: 8px;
  background: rgba(252, 254, 255, 0.62);
  box-shadow: var(--paper-shadow-sm);
}

.log-tabs :deep(.ant-segmented-group) {
  gap: 2px;
}

.log-tabs :deep(.ant-segmented-item) {
  min-height: 25px;
  border-radius: 6px;
  color: var(--ink);
  font-size: 13px;
  font-weight: 400;
  line-height: 25px;
  white-space: nowrap;
}

.log-tabs :deep(.ant-segmented-item-label) {
  min-height: 25px;
  padding: 0 9px;
  line-height: 25px;
}

.log-tabs :deep(.ant-segmented-item-selected) {
  color: var(--blue);
  background: rgba(225, 240, 250, 0.92);
  box-shadow: var(--paper-shadow-sm);
}

.log-tabs :deep(.ant-segmented-thumb) {
  background: rgba(225, 240, 250, 0.92);
  box-shadow: var(--paper-shadow-sm);
}

.log-actions {
  justify-content: flex-end;
  gap: 16px;
}

.log-folder-button {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  height: 30px;
  padding: 0 10px;
  border-color: rgba(104, 141, 181, 0.22);
  border-radius: 7px;
  color: var(--ink);
  background: rgba(252, 254, 255, 0.58);
  box-shadow: var(--paper-shadow-sm);
  font-size: 13px;
  font-weight: 400;
}

.log-folder-button:hover,
.log-folder-button:focus-visible {
  border-color: rgba(45, 111, 168, 0.28);
  color: var(--blue);
  background: rgba(225, 240, 250, 0.78);
}

.log-folder-button i {
  width: 14px;
  text-align: center;
}

.dark .log-tabs {
  border-color: rgba(128, 153, 188, 0.18);
  background: rgba(15, 24, 39, 0.48);
}

.dark .log-tabs :deep(.ant-segmented-item) {
  color: #c3d0e2;
}

.dark .log-tabs :deep(.ant-segmented-item-selected),
.dark .log-tabs :deep(.ant-segmented-thumb) {
  color: #8ab9f6;
  background: rgba(38, 58, 86, 0.86);
}

.dark .log-folder-button {
  border-color: rgba(128, 153, 188, 0.18);
  color: #c3d0e2;
  background: rgba(15, 24, 39, 0.48);
}

.dark .log-folder-button:hover,
.dark .log-folder-button:focus-visible {
  color: #8ab9f6;
  background: rgba(38, 58, 86, 0.78);
}

.log-columns {
  position: relative;
  z-index: 1;
  box-sizing: border-box;
  height: 200px;
  margin-top: 8px;
  display: grid;
  /* 日志左右两栏平分宽度，避免单篇任务栏视觉上压过主流程日志。 */
  grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
  gap: 10px;
}

.log-column {
  display: grid;
  grid-template-rows: 24px minmax(0, 1fr);
  min-width: 0;
  min-height: 0;
}

.log-column-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  min-width: 0;
  padding: 0 4px;
  color: var(--ink-muted);
  font-size: 12px;
  line-height: 1;
}

.log-column-header h3 {
  min-width: 0;
  margin: 0;
  color: var(--ink);
  font-size: 13px;
  font-weight: 500;
  line-height: 1;
  white-space: nowrap;
}

.log-column-header span {
  flex: none;
  color: var(--ink-muted);
  font-size: 12px;
  white-space: nowrap;
}

.log-table {
  position: relative;
  z-index: 1;
  box-sizing: border-box;
  height: 100%;
  margin-top: 0;
  padding: 8px 10px;
  border: 1px solid rgba(104, 141, 181, 0.2);
  border-radius: 8px;
  background: rgba(252, 254, 255, 0.56);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.62);
  overflow: auto;
  scrollbar-gutter: stable;
}

.dark .log-table {
  border-color: rgba(128, 153, 188, 0.18);
  background: rgba(15, 24, 39, 0.44);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.045);
}

.log-table :deep(.vue-recycle-scroller__item-view) {
  box-sizing: border-box;
  padding-right: 2px;
}

.log-row {
  position: relative;
  z-index: 1;
  display: grid;
  grid-template-columns: 68px 58px minmax(0, 1fr);
  gap: 6px;
  min-height: 24px;
  align-items: start;
  color: var(--ink);
  font-size: 14px;
  text-shadow: none;
  filter: none;
}

.log-row time,
.log-row strong {
  padding-top: 2px;
}

.log-row time {
  color: var(--ink-muted);
}

.log-level {
  color: var(--ink);
  font-weight: 400;
}

.log-level-info {
  color: var(--ink);
}

.log-message {
  min-width: 0;
  max-width: 100%;
  opacity: 1;
  filter: none;
  line-height: 1.45;
  text-shadow: none;
  white-space: normal;
  overflow: hidden;
  overflow-wrap: anywhere;
  word-break: break-word;
}

.log-url {
  display: block;
  max-width: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  vertical-align: bottom;
  white-space: nowrap;
}

.log-row-info .log-message {
  color: var(--ink);
}

.log-level-success,
.log-row-success .log-message {
  color: #16a15d;
}

.log-level-warn,
.log-row-warn .log-message {
  color: var(--orange);
}

.log-level-error,
.log-row-error .log-message {
  color: var(--red);
}

.log-empty {
  padding: 26px 0;
  color: var(--ink-muted);
  font-size: 14px;
  font-weight: 400;
  text-align: center;
}

.sprig-corner::after {
  content: none;
}

.blue {
  color: var(--blue) !important;
}

.green {
  color: var(--green) !important;
}

.red {
  color: var(--red) !important;
}

.purple {
  color: var(--purple) !important;
}

.orange {
  color: var(--orange) !important;
}

@supports not ((backdrop-filter: blur(1px)) or (-webkit-backdrop-filter: blur(1px))) {
  .panel,
  .task-count-input,
  .notice,
  .env-item {
    background: var(--paper);
  }
}

@media (prefers-reduced-transparency: reduce) {
  .panel,
  .task-count-input,
  .notice,
  .env-item {
    background: var(--paper);
    backdrop-filter: none;
    -webkit-backdrop-filter: none;
  }
}

@media (prefers-reduced-motion: reduce) {
  .task-date-filter-chevron {
    transition: none;
  }

  *,
  *::before,
  *::after {
    transition-duration: 0.01ms !important;
    animation-duration: 0.01ms !important;
    animation-iteration-count: 1 !important;
  }
}

</style>
