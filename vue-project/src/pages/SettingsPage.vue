<script setup lang="ts">
import AppIcon from '../components/AppIcon.vue'
import { computed, nextTick, reactive, ref, watch } from 'vue'
import { notification } from 'ant-design-vue'
import { resolveDiagnosticProgressMessage } from '../utils/diagnosticProgressMessage'
import {
  checkCaCertificate,
  clearRuntimeCache,
  deleteMitmCaCertificates,
  disableSystemProxy,
  enableSystemProxy,
  getTaskStatus,
  getArticleDetailDiagnostic,
  getArticleDetailCommentsDiagnostic,
  getArticleDetailOfflineCacheDiagnostic,
  getInitialContentStorageDiagnostic,
  getWindowClickFlowDiagnostic,
  installCaCertificate,
  listRuntimePaths,
  openRuntimePath,
  resetRuntimeConfig,
  saveRuntimeConfig,
  selectRuntimeDirectory,
  startMitmProxy,
  stopMitmProxy,
  testProxyConnection,
  updateRuntimeConfig,
  startArticleDetailDiagnostic,
  startArticleDetailCommentsDiagnostic,
  startArticleDetailOfflineCacheDiagnostic,
  startInitialContentStorageDiagnostic,
  startWindowClickFlowDiagnostic,
  stopWindowClickFlowDiagnostic,
  runStartupSelfCheck,
  runWindowDiagnosticAction,
} from '../bridge/pythonApi'
import type {
  CaCertificateStatus,
  MitmCertificateItem,
  MitmProjectCertificate,
  RuntimeConfigPayload,
  RuntimeLogLevel,
  RuntimePathKey,
  TaskStatus,
  WindowDiagnosticAction,
  WindowDiagnosticOptions,
  WindowDiagnosticResult,
  ArticleDetailDiagnosticResult,
  WindowClickFlowDiagnosticOptions,
  StartupSelfCheckResult,
} from '../bridge/pythonApi'

type EnvironmentItemTone = 'blue' | 'green' | 'orange' | 'red' | 'purple'

type EnvironmentItem = {
  name: string
  value: string
  icon: string
  tone?: EnvironmentItemTone
}

type DiagnosticResultTone = 'info' | 'success' | 'warning' | 'error'

type DiagnosticResultCell = {
  label: string
  value: string
}

type DiagnosticResultItem = DiagnosticResultCell & {
  cells?: DiagnosticResultCell[]
  kind?: 'summary' | 'operation' | 'discarded' | 'article'
  tone?: DiagnosticResultTone
  sequence?: number
}

type CaCertificateDialogMode = 'check' | 'install' | 'delete'
type CaCertificateDialogPhase = 'checking' | 'ready' | 'installing' | 'deleting' | 'done' | 'error'

const props = defineProps<{
  taskStatus?: TaskStatus | null
  environmentItems?: EnvironmentItem[]
}>()

const emit = defineEmits<{
  taskStatusChanged: [status: TaskStatus]
}>()

// 配置默认值集中放在组件内部，后续接后端保存接口时可以直接提交这个结构。
const CONFIG = {
  projectDir: '',
  storageDir: '',
  logDir: '',
  fileNameMode: '按发布日期和标题（YYYY-MM-DD HH-mm TITLE）',
  logLevel: 'INFO' as RuntimeLogLevel,
  requestIntervalSeconds: 2,
  proxyHost: '127.0.0.1',
  proxyPort: 18000,
  startupDelaySeconds: 0,
  trafficCheckUrl: 'https://mitm.it/',
}

const configForm = reactive({ ...CONFIG })

type NumericConfigKey =
  | 'requestIntervalSeconds'
  | 'proxyPort'
  | 'startupDelaySeconds'

const numericConfigLimits: Record<NumericConfigKey, { min: number; max: number; step: number }> = {
  requestIntervalSeconds: { min: 0, max: 3600, step: 1 },
  proxyPort: { min: 1, max: 65535, step: 1 },
  startupDelaySeconds: { min: 0, max: 300, step: 1 },
}

const fileNameModeOptions = [
  { label: CONFIG.fileNameMode, value: CONFIG.fileNameMode },
]

const logLevelOptions = [
  { label: 'DEBUG（调试信息）', value: 'DEBUG' },
  { label: 'INFO（一般信息）', value: 'INFO' },
  { label: 'WARN（警告信息）', value: 'WARN' },
  { label: 'ERROR（错误信息）', value: 'ERROR' },
]

const settings = reactive({
  autoStartProxy: false,
  enableSystemProxy: true,
  autoSaveContent: true,
  autoCleanTempFiles: true,
})

const caCertificateStatus = ref<CaCertificateStatus>({
  ok: false,
  status: 'unknown',
  installed: false,
  label: '未检测',
  message: '点击“检测状态”检查本机 CA 证书。',
})
const isCheckingCaCertificate = ref(false)
const isInstallingCaCertificate = ref(false)
const caCertificateDialogVisible = ref(false)
const caCertificateDialogMode = ref<CaCertificateDialogMode>('check')
const caCertificateDialogPhase = ref<CaCertificateDialogPhase>('checking')
const caCertificateDialogMessage = ref('')
const caCertificateDialogProjectCertificate = ref<MitmProjectCertificate | null>(null)
const caCertificateDialogSystemCertificates = ref<MitmCertificateItem[]>([])
const caCertificateDialogDeletedItems = ref<Array<{ thumbprint: string; storePath: string }>>([])
const caCertificateDialogSkippedItems = ref<Array<{ thumbprint: string; reason: string }>>([])

const caCertificateTone = computed(() => {
  if (isCheckingCaCertificate.value) {
    return 'checking'
  }

  if (caCertificateStatus.value.status === 'installed') {
    return 'success'
  }

  if (caCertificateStatus.value.status === 'missing') {
    return 'warning'
  }

  return 'error'
})

const caCertificateLabel = computed(() => (
  isCheckingCaCertificate.value ? '检测中' : caCertificateStatus.value.label
))

const configGuideItems = [
  '修改配置后请点击 保存配置 使设置生效。',
  '代理端口、代理模式变更后，建议重启 MITM。',
  '目录配置只影响新任务，旧归档不会自动移动。',
  'CA 证书未安装时，HTTPS 内容可能无法捕获。',
]

type SettingsCategoryKey =
  | 'basic'
  | 'main-flow'
  | 'single-article-task'
  | 'diagnostics'

type SettingsItemKey =
  | 'project-storage'
  | 'database-storage'
  | 'runtime-maintenance'
  | 'proxy-basic'
  | 'main-home-window'
  | 'main-home-scroll'
  | 'main-dispatch-control'
  | 'single-mitm-capture'
  | 'single-html-storage'
  | 'single-comment-collection'
  | 'single-offline-cache'
  | 'single-article-tab'
  | 'mitm-diagnostics'
  | 'window-diagnostics'
  | 'flow-diagnostics'

type SettingsItem = {
  key: SettingsItemKey
  label: string
  summary: string
  icon: string
}

type SettingsCategory = {
  key: SettingsCategoryKey
  label: string
  icon: string
  summary: string
  description: string
  items: SettingsItem[]
}

type SettingsDetailField = {
  label: string
  configKey: string
  value: string | SettingsDetailRangeValue
  description: string
  inputType?: 'text' | 'number' | 'number-stepper' | 'switch' | 'readonly-range'
  unit?: string
  min?: number
  max?: number
  step?: number
  browseLabel?: string
  browseAction?: (field: SettingsDetailField) => void | Promise<void>
  tone?: 'default' | 'readonly' | 'success' | 'warning'
}

type SettingsDetailRangeValue = {
  start: string
  end: string
  startLabel: string
  endLabel: string
}

type SettingsDetailControl = {
  kind:
    | 'log-level'
    | 'auto-clean'
    | 'proxy-host'
    | 'proxy-port'
    | 'verification-url'
    | 'startup-delay'
    | 'system-proxy'
    | 'mitm-proxy'
  label: string
  description: string
  configKey?: string
}

type SettingsDetailAction = {
  label: string
  buttonLabel?: string
  description: string
  detail?: string
  icon: string
  tone?: 'primary' | 'success' | 'orange' | 'danger' | 'ghost' | 'blue' | 'purple' | 'teal'
  disabled?: () => boolean
  showWindowClickFlowOptions?: boolean
  showScrollStepInput?: boolean
  showArticleDetailSkipCollectedOption?: boolean
  showInitialContentStorageOptions?: boolean
  showArticleDetailCommentsOptions?: boolean
  showArticleDetailOfflineCacheOptions?: boolean
  run: () => void | Promise<void>
}

type SettingsDetailContent = {
  note: string
  fields?: SettingsDetailField[]
  controls?: SettingsDetailControl[]
  actions?: SettingsDetailAction[]
}

type SettingsControlLayoutClass = 'compact-control' | 'wide-control'

const settingsNumberValues = reactive<Record<string, number>>({
  'basic_settings.runtime_maintenance.temp_retention_days': 7,
  'basic_settings.runtime_maintenance.log_retention_days': 30,
  'main_flow.home_window.activation_wait_seconds': 0.05,
  'main_flow.home_window.home_find_timeout_seconds': 3,
  'main_flow.home_scroll.scroll_initial_delay_seconds': 0.05,
  'main_flow.home_scroll.unchanged_before_bounce_seconds': 0.6,
  'main_flow.home_scroll.lazy_load_timeout_seconds': 3,
  'main_flow.home_scroll.bounce_up_steps': 2,
  'main_flow.home_scroll.bounce_pause_seconds': 0.2,
  'main_flow.home_scroll.bounce_down_steps': 6,
  'main_flow.home_scroll.bounce_attempts': 2,
  'main_flow.dispatch_control.single_task_interval_seconds': 0,
  'single_article_task.article_tab.article_title_stable_delay_seconds': 0.1,
  'single_article_task.article_tab.article_open_timeout_seconds': 12,
  'single_article_task.article_tab.article_close_confirm_timeout_seconds': 3,
  'single_article_task.mitm_capture.ready_timeout_seconds': 10,
  'single_article_task.mitm_capture.capture_timeout_seconds': 20,
  'single_article_task.mitm_capture.result_timeout_seconds': 11,
  'single_article_task.mitm_capture.listener_shutdown_timeout_seconds': 3,
  'single_article_task.html_storage.request_timeout_seconds': 10,
  'single_article_task.comment_collection.request_timeout_seconds': 10,
  'single_article_task.comment_collection.page_interval_seconds': 0.5,
  'single_article_task.comment_collection.top_level_max_pages': 50,
  'single_article_task.comment_collection.max_concurrent_processes': 3,
  'single_article_task.offline_cache.max_scroll_seconds': 30,
  'single_article_task.offline_cache.resource_timeout_seconds': 10,
  'single_article_task.offline_cache.max_concurrent_processes': 3,
})

const settingsToggleValues = reactive<Record<string, boolean>>({
  'basic_settings.proxy_settings.ssl_insecure': true,
  'single_article_task.mitm_capture.close_as_capture_deadline': true,
  'single_article_task.article_tab.restore_focus_after_close': true,
  'main_flow.home_scroll.bounce_enabled': true,
  'single_article_task.comment_collection.enabled_by_default': false,
  'single_article_task.offline_cache.enabled_by_default': false,
})

const runtimeConfigValues = reactive<Record<string, string>>({})
let runtimeConfigSyncTimer = 0

const readonlyRangeConfigKeys = new Set([
  'single_article_task.article_tab.article_title_poll_interval_seconds_range',
  'main_flow.home_scroll.date_seek_scroll_steps_range',
  'main_flow.home_scroll.scroll_probe_interval_seconds_range',
])

const wideControlKinds = new Set<SettingsDetailControl['kind']>([
  'verification-url',
])

function getSettingsControlLayoutClass(control: SettingsDetailControl): SettingsControlLayoutClass {
  return wideControlKinds.has(control.kind) ? 'wide-control' : 'compact-control'
}

function getSettingsFieldLayoutClass(field: SettingsDetailField): SettingsControlLayoutClass {
  const configKey = field.configKey.toLowerCase()
  const isLongPathValue = field.browseAction
    || configKey.includes('path')
    || configKey.includes('dir')
    || configKey.includes('url')

  // 路径、目录和 URL 需要保留宽输入区，短值和数字项才放到右侧紧凑区。
  return isLongPathValue ? 'wide-control' : 'compact-control'
}

function getSettingsConfigKeyPath(configKey?: string) {
  return configKey ? configKey.split('.').join('\u00A0-->\u00A0') : ''
}

function getSettingsFieldLabel(field: SettingsDetailField) {
  return field.unit ? `${field.label}（${field.unit}）` : field.label
}

const settingsCategories: SettingsCategory[] = [
  {
    key: 'diagnostics',
    label: '诊断工具',
    icon: 'fa-solid fa-list-check',
    summary: 'MITM 管理、窗口操作和基于主页窗口的流程测试入口',
    description: '诊断工具不是保存型配置，而是程序测试和排障入口。当前前端先固定入口样式，后端接口后续逐项接入。',
    items: [
      { key: 'mitm-diagnostics', label: 'MITM 管理', summary: 'CA 证书检测、安装、清除、HTTPS 校验、系统代理和 MITM 代理切换', icon: 'fa-solid fa-certificate' },
      { key: 'window-diagnostics', label: '窗口操作', summary: '微信主页激活、文章点击、滚动、回滚滚动和文章标签关闭', icon: 'fa-solid fa-window-restore' },
      { key: 'flow-diagnostics', label: '流程测试', summary: '基于主页窗口测试窗口点击、详情、内容存储和评论采集', icon: 'fa-solid fa-code-branch' },
    ],
  },
  {
    key: 'basic',
    label: '基础设置',
    icon: 'fa-solid fa-gear',
    summary: '项目存储、数据库、运行维护和代理基础参数',
    description: '基础设置对应 basic_settings 配置，集中展示项目目录、数据库配置、运行维护参数和代理基础信息。数据库存储当前只做只读展示，避免误改数据结构版本和数据库目录。',
    items: [
      { key: 'project-storage', label: '项目存储', summary: '文章归档目录、临时文件目录和日志目录', icon: 'fa-regular fa-folder-open' },
      { key: 'database-storage', label: '数据库存储', summary: '数据表结构版本和数据库目录', icon: 'fa-solid fa-database' },
      { key: 'runtime-maintenance', label: '运行维护', summary: '日志、启动清理和文件保留周期', icon: 'fa-solid fa-broom' },
      { key: 'proxy-basic', label: '代理基础设置', summary: '监听地址、监听端口、证书目录、CA 证书路径和代理验证地址', icon: 'fa-solid fa-plug-circle-check' },
    ],
  },
  {
    key: 'main-flow',
    label: '主流程',
    icon: 'fa-solid fa-code-branch',
    summary: '主页窗口定位、主页滚动和单篇任务分发节奏',
    description: '主流程对应 main_flow 配置，只负责围绕公众号主页窗口读取文章卡片、滚动定位和向单篇任务队列派发任务。',
    items: [
      { key: 'main-home-window', label: '主页窗口', summary: '主页窗口激活等待和主页定位超时', icon: 'fa-regular fa-hand-pointer' },
      { key: 'main-home-scroll', label: '主页滚动', summary: '日期定位步长、页面变化检测、懒加载和回弹滚动', icon: 'fa-solid fa-download' },
      { key: 'main-dispatch-control', label: '任务分发', summary: '主流程派发单篇任务之间的节奏控制', icon: 'fa-solid fa-network-wired' },
    ],
  },
  {
    key: 'single-article-task',
    label: '单篇任务',
    icon: 'fa-regular fa-file-lines',
    summary: '文章标签、MITM 捕获、HTML 存储、评论采集和离线缓存',
    description: '单篇任务对应 single_article_task 配置，控制每篇文章从详情获取到内容存储、评论采集和离线缓存的可调参数。',
    items: [
      { key: 'single-article-tab', label: '单篇标签操作', summary: '文章详情页标题检测、标题稳定等待、关闭确认和焦点恢复', icon: 'fa-solid fa-window-restore' },
      { key: 'single-mitm-capture', label: 'MITM 捕获', summary: '代理启动、系统代理接管、证书校验和捕获生命周期', icon: 'fa-solid fa-shield-halved' },
      { key: 'single-html-storage', label: '文章详情存储', summary: '使用 reference 参数重新请求文章 HTML 的超时时间', icon: 'fa-regular fa-file-code' },
      { key: 'single-comment-collection', label: '评论采集', summary: '默认开关、请求节奏、页数上限和子进程并发数', icon: 'fa-regular fa-clipboard' },
      { key: 'single-offline-cache', label: '离线缓存', summary: 'Playwright 打开文章后的滚动加载和离线资源下载超时', icon: 'fa-solid fa-file-shield' },
    ],
  },
]

const cacheCleaned = ref(false)
const isClearingCache = ref(false)
const isListingMitmCertificates = ref(false)
const isDeletingMitmCertificates = ref(false)
const mitmCertificateItems = caCertificateDialogSystemCertificates
const mitmCertificateMessage = caCertificateDialogMessage
const diagnosticResultDialogVisible = ref(false)
const diagnosticResultTitle = ref('')
const diagnosticResultMessage = ref('')
const diagnosticResultTone = ref<DiagnosticResultTone>('info')
const diagnosticResultItems = ref<DiagnosticResultItem[]>([])
const diagnosticResultListRef = ref<HTMLElement | null>(null)
const shouldAutoFollowDiagnosticResult = ref(true)
const isSyncingProxyState = ref(false)
const isHydratingConfig = ref(false)
const isApplyingDefaults = ref(false)
const resetDefaultsDialogVisible = ref(false)
const isTestingProxyConnection = ref(false)
const isWindowDiagnosticRunning = ref(false)
const windowDiagnosticScrollSteps = ref(3)
const isWindowClickFlowDiagnosticRunning = ref(false)
const isArticleDetailDiagnosticRunning = ref(false)
const isInitialContentStorageDiagnosticRunning = ref(false)
const isArticleDetailCommentsDiagnosticRunning = ref(false)
const isArticleDetailOfflineCacheDiagnosticRunning = ref(false)
const articleDetailSkipCollectedRecords = ref(false)
const initialContentStorageSkipCollectedRecords = ref(false)
const initialContentStorageStoreArticleDetail = ref(true)
const articleDetailCommentsSkipCollectedRecords = ref(false)
const articleDetailCommentsStoreArticleDetail = ref(true)
const articleDetailCommentsStoreCommentInfo = ref(true)
const articleDetailOfflineCacheSkipCollectedRecords = ref(false)
const articleDetailOfflineCacheStateful = ref(false)
const articleDetailOfflineCacheStoreArticleDetail = ref(true)
const articleDetailOfflineCacheArchiveContent = ref(true)
const activeWindowClickFlowJobId = ref('')
const windowClickFlowMaxRecords = ref(20)
const windowClickFlowDateFilterMode = ref<WindowClickFlowDiagnosticOptions['dateFilterMode']>('all')
const windowClickFlowStartDate = ref('')
const windowClickFlowEndDate = ref('')
const windowClickFlowDateRangeValue = computed<string[] | null>({
  get() {
    if (!windowClickFlowStartDate.value && !windowClickFlowEndDate.value) {
      return null
    }

    return [windowClickFlowStartDate.value, windowClickFlowEndDate.value]
  },
  set(value) {
    windowClickFlowStartDate.value = Array.isArray(value) ? String(value[0] ?? '') : ''
    windowClickFlowEndDate.value = Array.isArray(value) ? String(value[1] ?? '') : ''
  },
})
const windowClickFlowDateFilterOptions = [
  { label: '不限日期', value: 'all' },
  { label: '日期范围', value: 'range' },
  { label: '截止日期', value: 'before' },
  { label: '起始日期', value: 'after' },
]
const windowClickFlowUsesUnlimitedRecords = computed(() => (
  windowClickFlowDateFilterMode.value === 'range'
  || windowClickFlowDateFilterMode.value === 'before'
))
const isRunningStartupSelfCheck = ref(false)
const isSavingConfig = ref(false)
const runtimeStatus = ref<TaskStatus | null>(null)
let cacheCleanedTimer: number | undefined

const settingsEnvironmentItems = computed<EnvironmentItem[]>(() => props.environmentItems ?? [])
const currentTaskStatus = computed(() => props.taskStatus ?? runtimeStatus.value)
const isRuntimeCacheBusy = computed(() => (
  ['starting', 'running'].includes(String(currentTaskStatus.value?.status ?? '').toLowerCase())
  || isMitmProxyDiagnosticActive.value
  || isWindowDiagnosticRunning.value
  || isWindowClickFlowDiagnosticRunning.value
  || isArticleDetailDiagnosticRunning.value
  || isInitialContentStorageDiagnosticRunning.value
  || isArticleDetailCommentsDiagnosticRunning.value
  || isArticleDetailOfflineCacheDiagnosticRunning.value
))
const selectedSettingsCategoryKey = ref<SettingsCategoryKey>(settingsCategories[0]!.key)
const selectedSettingsItemKey = ref<SettingsItemKey | null>(null)
const selectedSettingsCategory = computed<SettingsCategory>(() => (
  settingsCategories.find((category) => category.key === selectedSettingsCategoryKey.value) ?? settingsCategories[0]!
))
const activeSettingsItems = computed(() => selectedSettingsCategory.value.items)
const selectedSettingsItem = computed<SettingsItem | null>(() => {
  if (!selectedSettingsItemKey.value) {
    return null
  }

  return activeSettingsItems.value.find((item) => item.key === selectedSettingsItemKey.value) ?? null
})
const selectedCategoryFieldCount = computed(() => activeSettingsItems.value.length)
const proxyActualHost = computed(() => String(configForm.proxyHost || CONFIG.proxyHost).trim() || CONFIG.proxyHost)
const proxyDisplayHost = computed(() => (isLoopbackHost(proxyActualHost.value) ? 'localhost' : proxyActualHost.value))
const proxyListenMeta = computed(() => {
  const actualHost = proxyActualHost.value
  if (isLoopbackHost(actualHost)) {
    return '本机回环地址，供系统代理和浏览器连接。'
  }

  return `当前配置绑定：${actualHost}`
})
const systemProxyCurrentServer = computed(() => String(currentTaskStatus.value?.proxy?.systemProxyServer ?? '').trim())
const systemProxyConfiguredServer = computed(() => (
  currentTaskStatus.value?.proxy?.configuredProxyServer
    ?? `${proxyActualHost.value}:${configForm.proxyPort}`
))
const isSystemProxyDiagnosticActive = computed(() => Boolean(currentTaskStatus.value?.proxy?.systemProxyActive))
const systemProxyDiagnosticLabel = computed(() => (isSystemProxyDiagnosticActive.value ? '关闭系统代理' : '开启系统代理'))
const systemProxyDiagnosticDescription = computed(() => (
  isSystemProxyDiagnosticActive.value
    ? '关闭当前 Windows 系统代理，避免系统流量继续指向本地代理。'
    : '开启 Windows 系统代理，并指向当前本地 MITM 监听地址。'
))
const systemProxyDiagnosticDetail = computed(() => (
  isSystemProxyDiagnosticActive.value
    ? `当前代理：${systemProxyCurrentServer.value || '未读取'}`
    : `目标代理：${systemProxyConfiguredServer.value}`
))
const systemProxyDiagnosticIcon = computed(() => (
  isSystemProxyDiagnosticActive.value ? 'fa-solid fa-stop' : 'fa-solid fa-circle-play'
))
const systemProxyDiagnosticTone = computed<'success' | 'orange'>(() => (
  isSystemProxyDiagnosticActive.value ? 'orange' : 'success'
))
const mitmProxyDiagnosticPort = computed(() => Number(
  currentTaskStatus.value?.proxy?.mitmPort
    ?? configForm.proxyPort
    ?? CONFIG.proxyPort,
))
const isMitmProxyDiagnosticActive = computed(() => Boolean(currentTaskStatus.value?.proxy?.mitmEnabled))
const isMitmProxyPortUnavailable = computed(() => (
  !isMitmProxyDiagnosticActive.value
    && currentTaskStatus.value?.proxy?.mitmPortOwner === 'external'
))
const mitmProxyDiagnosticLabel = computed(() => {
  if (isMitmProxyPortUnavailable.value) {
    return '代理端口不可用'
  }

  return isMitmProxyDiagnosticActive.value ? '关闭MITM代理' : '开启MITM代理'
})
const mitmProxyDiagnosticDescription = computed(() => (
  isMitmProxyPortUnavailable.value
    ? '当前代理监听端口已被占用，请先释放端口或修改代理设置中的监听端口。'
    : isMitmProxyDiagnosticActive.value
      ? '关闭本地 MITM 代理监听进程，停止继续接收系统代理转发流量。'
      : '启动本地 MITM 代理监听进程，供系统代理转发流量。'
))
const mitmProxyDiagnosticDetail = computed(() => {
  if (isMitmProxyPortUnavailable.value) {
    return `占用端口：${mitmProxyDiagnosticPort.value}`
  }

  return isMitmProxyDiagnosticActive.value
    ? `已监听端口：${mitmProxyDiagnosticPort.value}`
    : `待监听端口：${mitmProxyDiagnosticPort.value}`
})
const mitmProxyDiagnosticIcon = computed(() => (
  isMitmProxyPortUnavailable.value
    ? 'fa-solid fa-triangle-exclamation'
    : isMitmProxyDiagnosticActive.value ? 'fa-solid fa-stop' : 'fa-solid fa-circle-play'
))
const mitmProxyDiagnosticTone = computed<'teal' | 'orange' | 'danger'>(() => (
  isMitmProxyPortUnavailable.value ? 'danger' : isMitmProxyDiagnosticActive.value ? 'orange' : 'teal'
))

const selectedSettingsDetail = computed<SettingsDetailContent | null>(() => {
  switch (selectedSettingsItem.value?.key) {
    case 'project-storage':
      return {
        note: '目录值对应 data/custom.yaml 的 basic_settings.project_storage 配置。相对路径按项目根目录解析，目录输入框只允许选中复制，目录变更统一通过右侧浏览按钮完成。',
        fields: [
          { label: '文章归档目录', configKey: 'basic_settings.project_storage.article_storage_root', value: 'storages', description: '文章详情、评论、离线网页等资源的默认归档目录。', browseLabel: '浏览', browseAction: handlePendingBrowseAction },
          { label: '临时文件目录', configKey: 'basic_settings.project_storage.temp_dir', value: 'data/tmp', description: '运行中间文件、导出临时文件和探针结果的存放目录。', browseLabel: '浏览', browseAction: handlePendingBrowseAction },
          { label: '日志目录', configKey: 'basic_settings.project_storage.log_dir', value: 'data/logs', description: '程序运行日志写入目录，后续清理策略由运行维护控制。', browseLabel: '浏览', browseAction: handlePendingBrowseAction },
        ],
      }
    case 'database-storage':
      return {
        note: '数据库存储对应 basic_settings.database_settings，当前按只读配置展示，用于确认数据表结构版本和数据库目录。',
        fields: [
          { label: '数据表结构版本', configKey: 'basic_settings.database_settings.data_schema_version', value: 'v2.1', description: '用于精确匹配 data/sql/create_script/ 下的建表脚本。', tone: 'readonly' },
          { label: '数据库目录', configKey: 'basic_settings.database_settings.db_dir', value: 'data/sql', description: '当前版本 SQLite 数据库文件所在目录。', tone: 'readonly' },
        ],
      }
    case 'runtime-maintenance':
      return {
        note: '运行维护对应 basic_settings.runtime_maintenance，用于控制日志、启动清理和文件保留周期。',
        controls: [
          { kind: 'log-level', label: '日志级别', description: '控制程序输出日志的最低级别，默认 INFO。', configKey: 'basic_settings.runtime_maintenance.log_level' },
          { kind: 'auto-clean', label: '启动自动清理', description: '程序启动时是否自动清理上次遗留的临时文件。', configKey: 'basic_settings.runtime_maintenance.auto_clean_temp_files' },
        ],
        fields: [
          { label: '临时文件保留天数', configKey: 'basic_settings.runtime_maintenance.temp_retention_days', value: '7', description: '超过该天数的临时文件允许被清理。', inputType: 'number-stepper', unit: '天', min: 0, max: 365, step: 1 },
          { label: '日志文件保留天数', configKey: 'basic_settings.runtime_maintenance.log_retention_days', value: '30', description: '超过该天数的旧日志允许被清理。', inputType: 'number-stepper', unit: '天', min: 0, max: 365, step: 1 },
        ],
      }
    case 'proxy-basic':
      return {
        note: '代理基础设置对应 basic_settings.proxy_settings。监听地址和端口会影响系统代理指向，证书路径用于检测、安装和清理 MITM CA。',
        controls: [
          { kind: 'proxy-host', label: '监听地址', description: 'MITM 服务绑定的本机地址，当前默认 127.0.0.1。', configKey: 'basic_settings.proxy_settings.host' },
          { kind: 'proxy-port', label: '监听端口', description: 'MITM 服务监听端口，当前默认 18000。', configKey: 'basic_settings.proxy_settings.port' },
          { kind: 'verification-url', label: '代理验证地址', description: '用于代理证书验证、安装和手动诊断。', configKey: 'basic_settings.proxy_settings.verification_url' },
        ],
        fields: [
          { label: 'mitmproxy 配置目录', configKey: 'basic_settings.proxy_settings.confdir', value: '.mitmproxy', description: 'mitmproxy 配置和证书目录。', tone: 'readonly' },
          { label: 'CA 证书路径', configKey: 'basic_settings.proxy_settings.ca_cert_path', value: '.mitmproxy/mitmproxy-ca-cert.cer', description: '用于检测、安装或删除本项目生成的 CA 证书。', tone: 'readonly' },
        ],
      }
    case 'single-mitm-capture':
      return {
        note: 'MITM 捕获对应 single_article_task.mitm_capture，并引用 basic_settings.proxy_settings 中的代理启动和系统代理策略。',
        controls: [
          { kind: 'startup-delay', label: '代理启动额外等待（秒）', description: 'MITM 代理启动后的额外等待时间；通常为 0，由 READY 检测判断可用性。', configKey: 'basic_settings.proxy_settings.startup_delay_seconds' },
          { kind: 'system-proxy', label: '接管系统代理', description: '采集时是否允许程序接管 Windows 系统代理。', configKey: 'basic_settings.proxy_settings.enable_system_proxy' },
          { kind: 'mitm-proxy', label: 'MITM 运行状态', description: '由单篇文章采集 attempt 自动启动和停止。' },
        ],
        fields: [
          { label: '宽松证书校验', configKey: 'basic_settings.proxy_settings.ssl_insecure', value: '开启', description: 'MITM 连接上游 HTTPS 时是否允许忽略证书校验。', inputType: 'switch' },
          { label: 'READY 通知等待', configKey: 'single_article_task.mitm_capture.ready_timeout_seconds', value: '10.0', description: '父进程等待 MITM 子进程发出 READY 通知的最大时长。', inputType: 'number-stepper', unit: '秒', min: 0, max: 300, step: 0.5 },
          { label: '单次捕获总超时', configKey: 'single_article_task.mitm_capture.capture_timeout_seconds', value: '20.0', description: 'MITM READY 后等待文章请求被捕获的最长时间。', inputType: 'number-stepper', unit: '秒', min: 1, max: 3600, step: 1 },
          { label: '捕获结果等待', configKey: 'single_article_task.mitm_capture.result_timeout_seconds', value: '11.0', description: '发送停止捕获命令后，等待捕获结果返回的最长时间。', inputType: 'number-stepper', unit: '秒', min: 0, max: 600, step: 0.5 },
          { label: '监听器关闭超时', configKey: 'single_article_task.mitm_capture.listener_shutdown_timeout_seconds', value: '3.0', description: '停止 MITM listener 的最长等待时间。', inputType: 'number-stepper', unit: '秒', min: 0, max: 60, step: 0.5 },
          { label: '关闭标签作为捕获截止点', configKey: 'single_article_task.mitm_capture.close_as_capture_deadline', value: '开启', description: '关闭文章标签时立即结束本次捕获，不再接收新的捕获结果。', inputType: 'switch' },
        ],
      }
    case 'single-article-tab':
      return {
        note: '单篇标签操作用于确认文章详情页标题、稳定后关闭文章标签，并把焦点恢复到公众号主页。',
        fields: [
          { label: '标题稳定等待', configKey: 'single_article_task.article_tab.article_title_stable_delay_seconds', value: '0.1', description: '检测到目标文章标签标题后短暂等待；等待结束后立即关闭文章标签。', inputType: 'number-stepper', unit: '秒', min: 0, max: 5, step: 0.05 },
          { label: '详情页打开超时', configKey: 'single_article_task.article_tab.article_open_timeout_seconds', value: '12.0', description: '点击文章后等待详情页打开的最大时长；超过仍未识别目标标题则认为打开失败。', inputType: 'number-stepper', unit: '秒', min: 1, max: 120, step: 0.5 },
          { label: '标题轮询间隔', configKey: 'single_article_task.article_tab.article_title_poll_interval_seconds_range', value: { start: '0.05', end: '0.15', startLabel: '起始值', endLabel: '最大值' }, description: '点击文章后等待文章标签打开时使用；从起始间隔开始检测标题，未检测到时逐步放大。使用区间：0.05 秒 ~ 0.15 秒。', inputType: 'readonly-range', unit: '秒', tone: 'readonly' },
          { label: '关闭确认超时', configKey: 'single_article_task.article_tab.article_close_confirm_timeout_seconds', value: '3.0', description: '关闭文章标签后，最多等待窗口标题变化或标签消失的时长。', inputType: 'number-stepper', unit: '秒', min: 0, max: 60, step: 0.5 },
          { label: '关闭后恢复主页焦点', configKey: 'single_article_task.article_tab.restore_focus_after_close', value: '开启', description: '关闭文章标签后立即聚焦回公众号主页窗口。', inputType: 'switch' },
        ],
      }
    case 'main-home-window':
      return {
        note: '主页窗口设置用于控制主流程定位和短暂激活公众号主页窗口的等待时间。',
        fields: [
          { label: '激活主页窗口等待', configKey: 'main_flow.home_window.activation_wait_seconds', value: '0.05', description: '激活公众号主页窗口后等待窗口稳定，再继续读取 UIA 树或滚动。', inputType: 'number-stepper', unit: '秒', min: 0, max: 10, step: 0.05 },
          { label: '主页定位超时', configKey: 'main_flow.home_window.home_find_timeout_seconds', value: '3.0', description: '诊断工具查找微信主页窗口的最大等待时间；超过后直接提示先打开公众号主页。', inputType: 'number-stepper', unit: '秒', min: 0.5, max: 30, step: 0.5 },
        ],
      }
    case 'main-home-scroll':
      return {
        note: '主页滚动操作用于查找下一篇候选文章、判断页面变化、等待懒加载，并在必要时执行回弹滚动。',
        fields: [
          { label: '日期定位滚动步长', configKey: 'main_flow.home_scroll.date_seek_scroll_steps_range', value: { start: '3', end: '18', startLabel: '普通收录步长', endLabel: '日期定位最大步长' }, description: '日期定位阶段根据目标日期距离动态放大滚动步长；普通收录仍使用基础步长。使用区间：3 步 ~ 18 步。', inputType: 'readonly-range', unit: '步', tone: 'readonly' },
          { label: '滚动后读取等待', configKey: 'main_flow.home_scroll.scroll_initial_delay_seconds', value: '0.05', description: '每次滚动后先等待，再开始读取主页可见文章。', inputType: 'number-stepper', unit: '秒', min: 0, max: 5, step: 0.05 },
          { label: '变化检测轮询间隔', configKey: 'main_flow.home_scroll.scroll_probe_interval_seconds_range', value: { start: '0.1', end: '0.4', startLabel: '起始值', endLabel: '最大值' }, description: '滚动后用于判断 UIA 页面内容是否发生变化；从起始间隔开始检测，页面未变化时逐步放大。使用区间：0.1 秒 ~ 0.4 秒。', inputType: 'readonly-range', unit: '秒', tone: 'readonly' },
          { label: '无变化判定等待', configKey: 'main_flow.home_scroll.unchanged_before_bounce_seconds', value: '0.6', description: '滚动后页面持续无变化达到该时长，就认为本次滚动无效并准备回弹。', inputType: 'number-stepper', unit: '秒', min: 0, max: 10, step: 0.1 },
          { label: '懒加载最长等待', configKey: 'main_flow.home_scroll.lazy_load_timeout_seconds', value: '3.0', description: '检测到页面处于加载状态时，等待新内容出现的最长时间。', inputType: 'number-stepper', unit: '秒', min: 0, max: 60, step: 0.5 },
          { label: '启用回弹滚动', configKey: 'main_flow.home_scroll.bounce_enabled', value: '开启', description: '滚动到底且没有触发懒加载时，是否允许先上滚再下滚寻找新文章。', inputType: 'switch' },
          { label: '回弹向上步长', configKey: 'main_flow.home_scroll.bounce_up_steps', value: '2', description: '回弹滚动时先向上滚动的步数。', inputType: 'number-stepper', unit: '步', min: 0, max: 50, step: 1 },
          { label: '回弹等待时间', configKey: 'main_flow.home_scroll.bounce_pause_seconds', value: '0.2', description: '向上滚动后等待该时长，再执行向下滚动。', inputType: 'number-stepper', unit: '秒', min: 0, max: 10, step: 0.1 },
          { label: '回弹向下步长', configKey: 'main_flow.home_scroll.bounce_down_steps', value: '6', description: '回弹等待后再次向下滚动的步数。', inputType: 'number-stepper', unit: '步', min: 0, max: 80, step: 1 },
          { label: '回弹滚动次数', configKey: 'main_flow.home_scroll.bounce_attempts', value: '2', description: '每轮寻找候选文章时，最多尝试回弹滚动的次数。', inputType: 'number-stepper', unit: '次', min: 0, max: 20, step: 1 },
        ],
      }
    case 'main-dispatch-control':
      return {
        note: '任务分发用于控制主流程连续派发单篇任务时的额外等待，0 表示不主动等待。',
        fields: [
          { label: '单篇任务派发间隔', configKey: 'main_flow.dispatch_control.single_task_interval_seconds', value: '0', description: '主流程两次派发单篇任务之间的额外等待时间。', inputType: 'number-stepper', unit: '秒', min: 0, max: 3600, step: 1 },
        ],
      }
    case 'single-html-storage':
      return {
        note: '文章详情存储用于 MITM 捕获到 reference 参数后重新获取文章 HTML。',
        fields: [
          { label: '请求超时时间', configKey: 'single_article_task.html_storage.request_timeout_seconds', value: '10', description: '使用 reference 参数重新请求文章 HTML 时，单次请求允许等待的最长时间。', inputType: 'number-stepper', unit: '秒', min: 1, max: 300, step: 1 },
        ],
      }
    case 'single-comment-collection':
      return {
        note: '评论采集属于文章主采集后的可选资源获取任务，具体任务页仍可用前端开关覆盖默认值。',
        fields: [
          { label: '默认采集评论', configKey: 'single_article_task.comment_collection.enabled_by_default', value: '关闭', description: '是否默认采集评论；具体任务也可以由前端开关覆盖。', inputType: 'switch' },
          { label: '评论请求超时', configKey: 'single_article_task.comment_collection.request_timeout_seconds', value: '10', description: '每一次评论 HTTP 请求最多等待的时间。', inputType: 'number-stepper', unit: '秒', min: 1, max: 300, step: 1 },
          { label: '分页请求间隔', configKey: 'single_article_task.comment_collection.page_interval_seconds', value: '0.5', description: '评论分页请求之间的等待时间，避免连续请求过快。', inputType: 'number-stepper', unit: '秒', min: 0, max: 30, step: 0.1 },
          { label: '一级评论最大页数', configKey: 'single_article_task.comment_collection.top_level_max_pages', value: '50', description: '一级评论最多请求的页数。', inputType: 'number-stepper', unit: '页', min: 1, max: 1000, step: 1 },
          { label: '评论子进程最大并发数', configKey: 'single_article_task.comment_collection.max_concurrent_processes', value: '3', description: '同时运行的评论采集独立子进程数量；超过上限的文章会排队。', inputType: 'number-stepper', unit: '个', min: 1, max: 10, step: 1 },
        ],
      }
    case 'single-offline-cache':
      return {
        note: '离线缓存使用 Playwright 打开文章短链并下载资源。这里先按当前 custom.yaml 展示默认值。',
        fields: [
          { label: '默认离线归档', configKey: 'single_article_task.offline_cache.enabled_by_default', value: '关闭', description: '是否默认在主服务页勾选离线归档；具体任务仍可由前端开关覆盖。', inputType: 'switch' },
          { label: '最长滚动加载', configKey: 'single_article_task.offline_cache.max_scroll_seconds', value: '30', description: '离线缓存时最长滚动加载时间。', inputType: 'number-stepper', unit: '秒', min: 1, max: 600, step: 1 },
          { label: '资源下载超时', configKey: 'single_article_task.offline_cache.resource_timeout_seconds', value: '10', description: '单个离线资源下载超时。', inputType: 'number-stepper', unit: '秒', min: 1, max: 300, step: 1 },
          { label: '缓存子进程最大并发数', configKey: 'single_article_task.offline_cache.max_concurrent_processes', value: '3', description: '同时运行的 Playwright 离线缓存独立子进程数量；超过上限的文章会排队。', inputType: 'number-stepper', unit: '个', min: 1, max: 10, step: 1 },
        ],
      }
    case 'mitm-diagnostics':
      return {
        note: 'MITM 管理按钮用于 CA 证书、系统代理和 MITM 代理排障。已有前端能力会直接调用现有动作，未接入项后续再替换为真实后端接口。',
        actions: [
          { label: 'CA证书检测', description: '检测当前用户根证书库中的 MITM CA 状态', icon: 'fa-solid fa-magnifying-glass', tone: 'blue', disabled: () => isCheckingCaCertificate.value, run: handleCheckCaCertificate },
          { label: 'CA证书安装', description: '安装当前项目生成的 mitmproxy CA 证书', icon: 'fa-solid fa-certificate', tone: 'primary', disabled: () => isInstallingCaCertificate.value, run: openInstallCaCertificateDialog },
          { label: '清除CA证书', description: '检索并清除系统中的 mitmproxy 相关 CA 证书', icon: 'fa-solid fa-trash-can', tone: 'danger', disabled: () => isListingMitmCertificates.value || isDeletingMitmCertificates.value, run: handleOpenMitmCertificateDialog },
          { label: 'HTTPS 校验', description: '请求代理验证地址，检查代理连通性', icon: 'fa-solid fa-tower-broadcast', tone: 'purple', disabled: () => isTestingProxyConnection.value, run: handleTestProxyConnection },
          { label: systemProxyDiagnosticLabel.value, description: systemProxyDiagnosticDescription.value, detail: systemProxyDiagnosticDetail.value, icon: systemProxyDiagnosticIcon.value, tone: systemProxyDiagnosticTone.value, disabled: () => isSyncingProxyState.value, run: toggleSystemProxyDiagnosticAction },
          { label: mitmProxyDiagnosticLabel.value, description: mitmProxyDiagnosticDescription.value, detail: mitmProxyDiagnosticDetail.value, icon: mitmProxyDiagnosticIcon.value, tone: mitmProxyDiagnosticTone.value, disabled: () => isSyncingProxyState.value || isMitmProxyPortUnavailable.value, run: toggleMitmProxyDiagnosticAction },
        ],
      }
    case 'window-diagnostics':
      return {
        note: '窗口操作会直接调用后端窗口控制服务，只测试微信窗口识别、聚焦、点击、滚动和标签关闭，不启动 MITM，也不写入采集数据。',
        actions: [
          { label: '读取主页', description: '查找已打开的微信主页窗口，只读取并展示公众号名称。', icon: 'fa-regular fa-address-card', tone: 'blue', disabled: () => isWindowDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value, run: () => handleWindowDiagnosticAction('read-home') },
          { label: '激活主页', description: '立刻聚焦到微信主页窗口；未找到主页时提示先打开公众号主页。', icon: 'fa-solid fa-window-restore', tone: 'primary', disabled: () => isWindowDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value, run: () => handleWindowDiagnosticAction('activate-home') },
          { label: '首篇点击', description: '立刻聚焦主页窗口，找到首篇候选文章并点击打开；不等待标题确认，也不关闭文章标签。', icon: 'fa-solid fa-play', tone: 'success', disabled: () => isWindowDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value, run: () => handleWindowDiagnosticAction('first-article-click') },
          { label: '滚动页面', description: '立刻聚焦主页窗口，并按左侧临时步长执行一次向下滚动；该值不会写入 YAML。', icon: 'fa-solid fa-download', tone: 'purple', showScrollStepInput: true, disabled: () => isWindowDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value, run: () => handleWindowDiagnosticAction('scroll-page') },
          { label: '回弹滚动', description: '立刻聚焦主页窗口，并执行一次先上滚后下滚的回弹操作。', icon: 'fa-solid fa-rotate-right', tone: 'orange', disabled: () => isWindowDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value, run: () => handleWindowDiagnosticAction('bounce-scroll') },
          { label: '关闭标签', description: '立刻打开记录弹窗，查找第一个文章标签并通过 Ctrl+W 关闭。', icon: 'fa-solid fa-xmark', tone: 'danger', disabled: () => isWindowDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value, run: () => handleWindowDiagnosticAction('close-tab') },
        ],
      }
    case 'flow-diagnostics':
      return {
        note: '流程测试均基于当前微信主页窗口。已接入窗口测试、详情获取、初始内容存储、详情评论和离线缓存。',
        actions: [
          { label: '窗口点击流程', buttonLabel: '窗口测试', description: '首次立即激活公众号主页，按 UIA 日期组和文章卡片读取当前可视内容；滚动后重新读取并用日期加标题衔接，不点击文章、不启动 MITM。', icon: 'fa-solid fa-window-restore', tone: 'blue', showWindowClickFlowOptions: true, disabled: () => isWindowClickFlowDiagnosticRunning.value || isWindowDiagnosticRunning.value, run: handleWindowClickFlowDiagnosticAction },
          { label: '单篇文章详情流程', buttonLabel: '详情获取', description: '激活主页窗口并读取当前可视区第一篇文章卡片', icon: 'fa-regular fa-file-lines', tone: 'purple', showArticleDetailSkipCollectedOption: true, disabled: () => isArticleDetailDiagnosticRunning.value || isInitialContentStorageDiagnosticRunning.value || isArticleDetailCommentsDiagnosticRunning.value || isArticleDetailOfflineCacheDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value, run: handleArticleDetailDiagnosticAction },
          { label: '初始内容存储测试', buttonLabel: '初始内容存储', description: '复用单篇文章详情流程，解析 HTML 并存储初始文章内容', icon: 'fa-solid fa-box-archive', tone: 'success', showInitialContentStorageOptions: true, disabled: () => isInitialContentStorageDiagnosticRunning.value || isArticleDetailDiagnosticRunning.value || isArticleDetailCommentsDiagnosticRunning.value || isArticleDetailOfflineCacheDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value, run: handleInitialContentStorageDiagnosticAction },
          { label: '单篇评论存储测试', buttonLabel: '评论信息存储', description: '复用初始内容存储，随后启动独立评论子进程采集评论', icon: 'fa-regular fa-clipboard', tone: 'orange', showArticleDetailCommentsOptions: true, disabled: () => isArticleDetailCommentsDiagnosticRunning.value || isArticleDetailDiagnosticRunning.value || isInitialContentStorageDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value || isArticleDetailOfflineCacheDiagnosticRunning.value, run: handleArticleDetailCommentsDiagnosticAction },
          { label: '单篇离线缓存测试', buttonLabel: '离线缓存', description: '复用初始内容存储，随后启动 Playwright 子进程生成 index.html 和 assets', icon: 'fa-solid fa-download', tone: 'primary', showArticleDetailOfflineCacheOptions: true, disabled: () => isArticleDetailOfflineCacheDiagnosticRunning.value || isArticleDetailCommentsDiagnosticRunning.value || isArticleDetailDiagnosticRunning.value || isInitialContentStorageDiagnosticRunning.value || isWindowClickFlowDiagnosticRunning.value, run: handleArticleDetailOfflineCacheDiagnosticAction },
        ],
      }
    default:
      return null
  }
})

function isLoopbackHost(host: string) {
  return ['localhost', '127.0.0.1', '::1'].includes(host.toLowerCase())
}

function selectSettingsCategory(key: SettingsCategoryKey) {
  selectedSettingsCategoryKey.value = key
  selectedSettingsItemKey.value = null
}

function selectSettingsItem(key: SettingsItemKey) {
  selectedSettingsItemKey.value = key
}

function getNumericConfigValue(key: NumericConfigKey) {
  if (key === 'requestIntervalSeconds') {
    return configForm.requestIntervalSeconds
  }

  if (key === 'proxyPort') {
    return configForm.proxyPort
  }

  return configForm.startupDelaySeconds
}

function setNumericConfigValue(key: NumericConfigKey, value: number) {
  const limit = numericConfigLimits[key]
  const safeValue = Number.isFinite(value) ? value : limit.min
  const nextValue = Math.min(limit.max, Math.max(limit.min, Math.round(safeValue)))

  // 这里按字段单独赋值，避免 TypeScript 对联合 key 写入推断过窄。
  if (key === 'requestIntervalSeconds') {
    configForm.requestIntervalSeconds = nextValue
    return
  }

  if (key === 'proxyPort') {
    configForm.proxyPort = nextValue
    return
  }

  configForm.startupDelaySeconds = nextValue
}

function handleNumericConfigNumberChange(key: NumericConfigKey, value: number | string | null) {
  setNumericConfigValue(key, Number(value))
}

function getSettingsDisplayValue(field: SettingsDetailField) {
  if (typeof field.value !== 'string') {
    return field.value.start + ' ~ ' + field.value.end
  }

  return runtimeConfigValues[field.configKey] ?? field.value
}

function getSettingsRangeValue(field: SettingsDetailField): SettingsDetailRangeValue {
  if (typeof field.value !== 'string') {
    return field.value
  }

  return {
    start: field.value,
    end: field.value,
    startLabel: '起始值',
    endLabel: '结束值',
  }
}

function getRuntimeConfigDisplayValue(configKey: string) {
  return runtimeConfigValues[configKey] ?? ''
}

function setRuntimeConfigDisplayValue(configKey: string, value: string | number | boolean) {
  runtimeConfigValues[configKey] = String(value)
}

function parseRuntimeNumberValue(configKey: string, fallback: number) {
  const rawValue = runtimeConfigValues[configKey]
  if (rawValue === undefined || rawValue === '') {
    return fallback
  }

  const parsedValue = Number(rawValue)
  return Number.isFinite(parsedValue) ? parsedValue : fallback
}

function parseRuntimeSwitchValue(configKey: string, fallback: boolean) {
  const rawValue = runtimeConfigValues[configKey]
  if (rawValue === undefined) {
    return fallback
  }

  return ['开启', '开', 'true', '1', 'yes', 'on'].includes(rawValue.trim().toLowerCase())
}

function getSettingsNumberValue(field: SettingsDetailField) {
  const fallbackValue = settingsNumberValues[field.configKey] ?? Number(field.value) ?? 0
  return parseRuntimeNumberValue(field.configKey, fallbackValue)
}

function getSettingsNumberPrecision(field: SettingsDetailField) {
  const step = field.step ?? 1
  const raw = String(step)

  return raw.includes('.') ? raw.split('.')[1]?.length ?? 0 : 0
}

function setSettingsNumberValue(field: SettingsDetailField, value: number) {
  const min = field.min ?? 0
  const max = field.max ?? 999999
  const safeValue = Number.isFinite(value) ? value : min
  const clampedValue = Math.min(max, Math.max(min, safeValue))
  const factor = 10 ** getSettingsNumberPrecision(field)

  const nextValue = Math.round(clampedValue * factor) / factor
  settingsNumberValues[field.configKey] = nextValue
  setRuntimeConfigDisplayValue(field.configKey, nextValue)
  queueRuntimeConfigMemorySync()
}

function handleSettingsNumberFieldNumberChange(field: SettingsDetailField, value: number | string | null) {
  setSettingsNumberValue(field, Number(value))
}

function showConfigNotice(message: string, tone: 'info' | 'success' | 'warning' | 'error' = 'info') {
  notification[tone]({
    message,
    placement: 'bottomRight',
    duration: 2.4,
  })
}

function closeDiagnosticResultDialog() {
  if (isWindowClickFlowDiagnosticRunning.value) {
    return
  }
  diagnosticResultDialogVisible.value = false
}

function handleDiagnosticResultBackdropClick() {
  if (isWindowClickFlowDiagnosticRunning.value) {
    return
  }
  closeDiagnosticResultDialog()
}

function openDiagnosticResultDialog(options: {
  title: string
  message?: string
  tone?: DiagnosticResultTone
  items?: DiagnosticResultItem[]
  status?: string
  action?: string
}) {
  const isNewDialog = !diagnosticResultDialogVisible.value
  const items = options.items ?? []
  diagnosticResultTitle.value = options.title
  diagnosticResultMessage.value = resolveDiagnosticProgressMessage({
    status: options.status,
    action: options.action,
    tone: options.tone,
    message: options.message,
    items,
  })
  diagnosticResultTone.value = options.tone ?? 'info'
  diagnosticResultItems.value = items
  if (isNewDialog) {
    shouldAutoFollowDiagnosticResult.value = true
  }
  diagnosticResultDialogVisible.value = true
  scrollDiagnosticResultListToEnd()
}

function handleDiagnosticResultListScroll() {
  const list = diagnosticResultListRef.value
  if (!list) {
    return
  }
  const { scrollHeight, scrollTop, clientHeight } = list
  shouldAutoFollowDiagnosticResult.value = (
    scrollHeight - scrollTop - clientHeight <= 24
  )
}

function scrollDiagnosticResultListToEnd() {
  if (!shouldAutoFollowDiagnosticResult.value) {
    return
  }
  void nextTick(() => {
    const list = diagnosticResultListRef.value
    if (list && shouldAutoFollowDiagnosticResult.value) {
      list.scrollTop = list.scrollHeight
    }
  })
}

const caCertificateDialogTitle = computed(() => {
  if (caCertificateDialogMode.value === 'install') {
    return '安装 CA 证书'
  }

  if (caCertificateDialogMode.value === 'delete') {
    return '清除 CA 证书'
  }

  return 'CA 证书检测'
})

const caCertificateDialogIcon = computed(() => {
  if (caCertificateDialogMode.value === 'delete') {
    return 'fa-solid fa-trash-can'
  }

  if (caCertificateDialogMode.value === 'install') {
    return 'fa-solid fa-certificate'
  }

  return 'fa-solid fa-magnifying-glass'
})

const caCertificateDialogTone = computed(() => {
  if (caCertificateDialogPhase.value === 'error') {
    return 'danger'
  }

  if (caCertificateDialogMode.value === 'delete') {
    return 'danger'
  }

  if (caCertificateDialogMode.value === 'install') {
    return 'warning'
  }

  return caCertificateStatus.value.installed ? 'success' : 'info'
})

const isCaCertificateDialogBusy = computed(() => (
  caCertificateDialogPhase.value === 'checking'
    || caCertificateDialogPhase.value === 'installing'
    || caCertificateDialogPhase.value === 'deleting'
))

const caCertificateDialogCanConfirmInstall = computed(() => (
  caCertificateDialogMode.value === 'install'
    && caCertificateDialogPhase.value === 'ready'
    && Boolean(caCertificateStatus.value.caFileExists)
    && !caCertificateStatus.value.projectCertificateInstalled
))

const caCertificateDialogCanConfirmDelete = computed(() => (
  caCertificateDialogMode.value === 'delete'
    && caCertificateDialogPhase.value === 'ready'
    && mitmCertificateItems.value.length > 0
))

const caCertificateDialogCloseText = computed(() => (
  caCertificateDialogPhase.value === 'ready' ? '取消' : '知道了'
))

const caCertificateDialogProjectPath = computed(() => (
  caCertificateDialogProjectCertificate.value?.path
    || caCertificateStatus.value.currentCaRelativePath
    || caCertificateStatus.value.currentCaPath
    || getRuntimeConfigDisplayValue('basic_settings.proxy_settings.ca_cert_path')
    || '.mitmproxy/mitmproxy-ca-cert.cer'
))

const caCertificateDialogProjectThumbprint = computed(() => (
  caCertificateDialogProjectCertificate.value?.thumbprint
    || caCertificateStatus.value.thumbprint
    || ''
))

function openCaCertificateDialog(options: {
  mode: CaCertificateDialogMode
  phase: CaCertificateDialogPhase
  message: string
}) {
  caCertificateDialogMode.value = options.mode
  caCertificateDialogPhase.value = options.phase
  caCertificateDialogMessage.value = options.message
  caCertificateDialogDeletedItems.value = []
  caCertificateDialogSkippedItems.value = []
  caCertificateDialogProjectCertificate.value = caCertificateStatus.value.projectCertificate ?? null
  mitmCertificateItems.value = caCertificateStatus.value.certificates ?? []
  caCertificateDialogVisible.value = true
}

function applyCaCertificateStatusPayload(payload: CaCertificateStatus) {
  caCertificateStatus.value = payload
  caCertificateDialogProjectCertificate.value = payload.projectCertificate ?? null
  mitmCertificateItems.value = payload.certificates ?? []
}

function formatCaDialogValue(value: string | boolean | number | undefined | null) {
  if (typeof value === 'boolean') {
    return value ? '是' : '否'
  }

  if (value === undefined || value === null || value === '') {
    return '未知'
  }

  return String(value)
}

function getCertificateDisplayName(certificate: MitmCertificateItem) {
  return certificate.friendlyName || certificate.subject || 'mitmproxy'
}

function formatDiagnosticValue(value: unknown): string {
  if (typeof value === 'boolean') {
    return value ? '是' : '否'
  }
  if (value === null || value === undefined || value === '') {
    return '无'
  }
  if (Array.isArray(value)) {
    return value.length ? value.map((item) => formatDiagnosticValue(item)).join('，') : '无'
  }
  if (typeof value === 'object') {
    return JSON.stringify(value)
  }

  return String(value)
}

function buildDiagnosticItems(
  source: Record<string, unknown>,
  fields: Array<{ key: string, label: string }>,
): DiagnosticResultItem[] {
  return fields.map((field) => ({
    label: field.label,
    value: formatDiagnosticValue(source[field.key]),
  }))
}

function createDiagnosticResultRow(cells: DiagnosticResultCell[]): DiagnosticResultItem {
  const firstCell = cells[0] ?? { label: '', value: '' }

  return {
    ...firstCell,
    cells: cells.slice(1),
  }
}

function getDiagnosticResultPrimaryCell(item: DiagnosticResultItem): DiagnosticResultCell {
  return { label: item.label, value: item.value }
}

function getDiagnosticResultMetaCells(item: DiagnosticResultItem): DiagnosticResultCell[] {
  return item.cells ?? []
}

function getDiagnosticResultCells(item: DiagnosticResultItem): DiagnosticResultCell[] {
  return [getDiagnosticResultPrimaryCell(item), ...getDiagnosticResultMetaCells(item)]
}

function getDiagnosticResultItemClass(item: DiagnosticResultItem) {
  return [
    'diagnostic-result-item',
    item.kind ? `diagnostic-result-item--${item.kind}` : '',
    item.tone ? `diagnostic-result-item--${item.tone}` : '',
    { 'diagnostic-result-item--split': getDiagnosticResultMetaCells(item).length > 0 },
  ]
}

function createHttpsCheckDiagnosticItems(): DiagnosticResultItem[] {
  return [
    { label: 'MITM 代理', value: '等待开启' },
    { label: '系统代理', value: '等待开启' },
    { label: 'HTTPS 校验', value: '等待测试' },
    { label: '关闭系统代理', value: '等待恢复' },
    { label: '关闭 MITM 代理', value: '等待恢复' },
  ]
}

function updateDiagnosticItem(
  items: DiagnosticResultItem[],
  label: string,
  value: string,
): DiagnosticResultItem[] {
  return items.map((item) => {
    if (item.label !== label) {
      return item
    }

    if (!item.cells?.length) {
      return { ...item, value }
    }

    return {
      ...item,
      value,
    }
  })
}

function updateDiagnosticItemCells(
  items: DiagnosticResultItem[],
  label: string,
  cells: DiagnosticResultCell[],
): DiagnosticResultItem[] {
  return items.map((item) => (item.label === label ? createDiagnosticResultRow(cells) : item))
}

async function toggleSystemProxyDiagnosticAction() {
  if (isSyncingProxyState.value) {
    return
  }

  const enabled = !isSystemProxyDiagnosticActive.value
  isSyncingProxyState.value = true
  try {
    const result = enabled ? await enableSystemProxy() : await disableSystemProxy()
    emitTaskStatus(result)
    await syncProxySwitchState()
    if (!result.ok) {
      showConfigNotice(result.message ?? '系统代理修改失败。', 'error')
      return
    }

    showConfigNotice(result.message ?? (enabled ? '系统代理已开启。' : '系统代理已关闭。'), enabled ? 'success' : 'warning')
  } catch (error) {
    const message = error instanceof Error ? error.message : '系统代理修改失败。'
    showConfigNotice(message, 'error')
  } finally {
    isSyncingProxyState.value = false
  }
}

async function toggleMitmProxyDiagnosticAction() {
  if (isSyncingProxyState.value) {
    return
  }
  if (isMitmProxyPortUnavailable.value) {
    showConfigNotice('代理端口不可用，请先释放端口或修改监听端口。', 'error')
    return
  }

  const enabled = !isMitmProxyDiagnosticActive.value
  isSyncingProxyState.value = true
  try {
    const result = enabled ? await startMitmProxy() : await stopMitmProxy()
    emitTaskStatus(result)
    await syncProxySwitchState()
    if (!result.ok) {
      showConfigNotice(result.message ?? 'MITM 代理修改失败。', 'error')
      return
    }

    showConfigNotice(result.message ?? (enabled ? 'MITM 代理已开启。' : 'MITM 代理已关闭。'), enabled ? 'success' : 'warning')
  } catch (error) {
    const message = error instanceof Error ? error.message : 'MITM 代理修改失败。'
    showConfigNotice(message, 'error')
  } finally {
    isSyncingProxyState.value = false
  }
}

async function handlePendingBrowseAction(field: SettingsDetailField) {
  try {
    const displayValue = getSettingsDisplayValue(field)
    const result = await selectRuntimeDirectory(
      field.configKey,
      typeof displayValue === 'string' ? displayValue : '',
    )
    if (result.taskStatus) {
      emitTaskStatus(result.taskStatus)
    }
    if (!result.ok || !result.selectedPath) {
      showConfigNotice(result.message ?? '已取消目录选择。', result.status === 'cancelled' ? 'info' : 'error')
      return
    }

    setRuntimeConfigDisplayValue(field.configKey, result.selectedPath)
    showConfigNotice(result.message ?? field.label + '已同步到运行配置。', 'success')
  } catch (error) {
    showConfigNotice(error instanceof Error ? error.message : field.label + '目录选择失败。', 'error')
  }
}

function selectReadonlyInputText(event: Event) {
  const input = event.target as HTMLInputElement

  if (!input.readOnly) {
    return
  }

  // 目录和只读配置允许复制原值，但不允许在输入框内直接改写。
  input.select()
}

function handlePendingDiagnosticAction(label: string) {
  showConfigNotice(label + ' 的后端接口待接入。', 'info')
}

function isRealtimeDiagnosticRunning(status: string | undefined) {
  return status === 'running' || status === 'stop-requested'
}

async function handleWindowClickFlowDiagnosticAction() {
  if (isWindowClickFlowDiagnosticRunning.value) {
    return
  }

  const options = buildWindowClickFlowDiagnosticOptions()
  if (!options) {
    return
  }

  isWindowClickFlowDiagnosticRunning.value = true
  activeWindowClickFlowJobId.value = ''
  openDiagnosticResultDialog({
    title: '主页内容读取结果',
    message: '正在启动主页内容读取测试...',
    tone: 'info',
    items: [],
  })

  let pollTimer: number | undefined
  try {
    const started = await startWindowClickFlowDiagnostic(options)
    activeWindowClickFlowJobId.value = started.jobId
    applyWindowClickFlowDiagnosticResult(started)

    let latest: ArticleDetailDiagnosticResult = started
    while (isRealtimeDiagnosticRunning(latest.status)) {
      await new Promise<void>((resolve) => {
        pollTimer = window.setTimeout(resolve, 500)
      })
      latest = await getWindowClickFlowDiagnostic(started.jobId)
      applyWindowClickFlowDiagnosticResult(latest)
    }

    showConfigNotice(
      latest.message ?? '主页内容读取测试已结束。',
      latest.ok ? 'success' : 'warning',
    )
  } catch (error) {
    const message = error instanceof Error ? error.message : '主页内容读取测试失败。'
    openDiagnosticResultDialog({
      title: '主页内容读取结果',
      message,
      tone: 'error',
      items: [],
    })
    showConfigNotice(message, 'error')
  } finally {
    if (pollTimer !== undefined) {
      window.clearTimeout(pollTimer)
    }
    isWindowClickFlowDiagnosticRunning.value = false
    activeWindowClickFlowJobId.value = ''
  }
}

function buildWindowClickFlowDiagnosticOptions(): WindowClickFlowDiagnosticOptions | null {
  const maxRecords = windowClickFlowUsesUnlimitedRecords.value
    ? 0
    : Math.max(1, Math.min(20, Math.trunc(Number(windowClickFlowMaxRecords.value) || 1)))
  windowClickFlowMaxRecords.value = maxRecords
  const mode = windowClickFlowDateFilterMode.value
  const startDate = windowClickFlowStartDate.value
  const endDate = windowClickFlowEndDate.value
  if ((mode === 'range' || mode === 'after') && !startDate) {
    showConfigNotice('请填写起始日期。', 'warning')
    return null
  }
  if ((mode === 'range' || mode === 'before') && !endDate) {
    showConfigNotice('请填写截止日期。', 'warning')
    return null
  }
  if (mode === 'range' && startDate > endDate) {
    showConfigNotice('起始日期不能晚于截止日期。', 'warning')
    return null
  }
  return {
    maxRecords,
    dateFilterMode: mode,
    startDate: mode === 'range' || mode === 'after' ? startDate : undefined,
    endDate: mode === 'range' || mode === 'before' ? endDate : undefined,
  }
}

function handleWindowClickFlowMaxRecordsChange(value: number | string | null) {
  if (windowClickFlowUsesUnlimitedRecords.value) {
    windowClickFlowMaxRecords.value = 0
    return
  }

  const parsedValue = Number(value)
  const safeValue = Number.isFinite(parsedValue) ? parsedValue : windowClickFlowMaxRecords.value
  windowClickFlowMaxRecords.value = Math.max(1, Math.min(20, Math.round(safeValue)))
}

watch(windowClickFlowDateFilterMode, () => {
  // 范围和截止模式由日期边界结束；其他模式恢复默认的 20 篇上限。
  windowClickFlowMaxRecords.value = windowClickFlowUsesUnlimitedRecords.value ? 0 : 20
})

async function stopActiveWindowClickFlowDiagnostic() {
  const jobId = activeWindowClickFlowJobId.value
  if (!jobId) {
    return
  }
  try {
    const result = await stopWindowClickFlowDiagnostic(jobId)
    applyWindowClickFlowDiagnosticResult(result)
    showConfigNotice(result.message ?? '已请求停止主页内容读取。', 'warning')
  } catch (error) {
    const message = error instanceof Error ? error.message : '停止主页内容读取失败。'
    diagnosticResultMessage.value = message
    diagnosticResultTone.value = 'error'
    showConfigNotice(message, 'error')
  }
}

async function handleArticleDetailDiagnosticAction() {
  if (isArticleDetailDiagnosticRunning.value) {
    return
  }

  isArticleDetailDiagnosticRunning.value = true
  openDiagnosticResultDialog({
    title: '详情获取结果',
    message: '正在启动单篇文章详情获取...',
    tone: 'info',
    items: [
      { label: '流程', value: '单篇文章详情流程' },
      {
        label: '跳过已采集记录',
        value: articleDetailSkipCollectedRecords.value ? '开启' : '关闭',
      },
      { label: '状态', value: '启动中' },
    ],
  })

  let pollTimer: number | undefined
  try {
    const started = await startArticleDetailDiagnostic({
      skipCollectedRecords: articleDetailSkipCollectedRecords.value,
    })
    applyArticleDetailDiagnosticResult(started)
    const jobId = started.jobId
    let latest: ArticleDetailDiagnosticResult = started

    while (latest.status === 'running') {
      await delay(500)
      latest = await getArticleDetailDiagnostic(jobId)
      applyArticleDetailDiagnosticResult(latest)
    }

    showConfigNotice(
      latest.message ?? '单篇文章详情获取已完成。',
      latest.ok ? 'success' : 'warning',
    )
  } catch (error) {
    const message = error instanceof Error ? error.message : '单篇文章详情获取失败。'
    openDiagnosticResultDialog({
      title: '详情获取结果',
      message,
      tone: 'error',
      items: [
        { label: '流程', value: '单篇文章详情流程' },
        { label: '失败原因', value: message },
      ],
    })
    showConfigNotice(message, 'error')
  } finally {
    if (pollTimer !== undefined) {
      window.clearTimeout(pollTimer)
    }
    isArticleDetailDiagnosticRunning.value = false
  }
}

async function handleInitialContentStorageDiagnosticAction() {
  if (isInitialContentStorageDiagnosticRunning.value) {
    return
  }

  isInitialContentStorageDiagnosticRunning.value = true
  openDiagnosticResultDialog({
    title: '初始内容存储结果',
    message: '正在启动初始内容存储测试...',
    tone: 'info',
    items: [
      { label: '流程', value: '初始内容存储测试' },
      {
        label: '跳过已采集记录',
        value: initialContentStorageSkipCollectedRecords.value ? '开启' : '关闭',
      },
      {
        label: '存储文章详情',
        value: initialContentStorageStoreArticleDetail.value ? '开启（锁定）' : '关闭',
      },
      { label: '状态', value: '启动中' },
    ],
  })

  try {
    const started = await startInitialContentStorageDiagnostic({
      skipCollectedRecords: initialContentStorageSkipCollectedRecords.value,
      storeArticleDetail: initialContentStorageStoreArticleDetail.value,
    })
    applyArticleDetailDiagnosticResult(started)
    const jobId = started.jobId
    let latest: ArticleDetailDiagnosticResult = started

    while (latest.status === 'running') {
      await delay(500)
      latest = await getInitialContentStorageDiagnostic(jobId)
      applyArticleDetailDiagnosticResult(latest)
    }

    showConfigNotice(
      latest.message ?? '初始内容存储测试已完成。',
      latest.ok ? 'success' : 'warning',
    )
  } catch (error) {
    const message = error instanceof Error ? error.message : '初始内容存储测试失败。'
    openDiagnosticResultDialog({
      title: '初始内容存储结果',
      message,
      tone: 'error',
      items: [
        { label: '流程', value: '初始内容存储测试' },
        { label: '失败原因', value: message },
      ],
    })
    showConfigNotice(message, 'error')
  } finally {
    isInitialContentStorageDiagnosticRunning.value = false
  }
}

async function handleArticleDetailCommentsDiagnosticAction() {
  if (isArticleDetailCommentsDiagnosticRunning.value) {
    return
  }

  isArticleDetailCommentsDiagnosticRunning.value = true
  openDiagnosticResultDialog({
    title: '详情评论结果',
    message: '正在启动详情评论测试...',
    tone: 'info',
    items: [
      { label: '流程', value: '单篇文章详情评论' },
      { label: '状态', value: '启动中' },
    ],
  })

  try {
    const started = await startArticleDetailCommentsDiagnostic({
      skipCollectedRecords: articleDetailCommentsSkipCollectedRecords.value,
      storeArticleDetail: articleDetailCommentsStoreArticleDetail.value,
      storeCommentInfo: articleDetailCommentsStoreCommentInfo.value,
    })
    applyArticleDetailDiagnosticResult(started)
    const jobId = started.jobId
    let latest: ArticleDetailDiagnosticResult = started

    while (latest.status === 'running') {
      await delay(500)
      latest = await getArticleDetailCommentsDiagnostic(jobId)
      applyArticleDetailDiagnosticResult(latest)
    }

    showConfigNotice(
      latest.message ?? '详情评论测试已完成。',
      latest.ok ? 'success' : 'warning',
    )
  } catch (error) {
    const message = error instanceof Error ? error.message : '详情评论测试失败。'
    openDiagnosticResultDialog({
      title: '详情评论结果',
      message,
      tone: 'error',
      items: [
        { label: '流程', value: '单篇文章详情评论' },
        { label: '失败原因', value: message },
      ],
    })
    showConfigNotice(message, 'error')
  } finally {
    isArticleDetailCommentsDiagnosticRunning.value = false
  }
}

async function handleArticleDetailOfflineCacheDiagnosticAction() {
  if (isArticleDetailOfflineCacheDiagnosticRunning.value) {
    return
  }

  isArticleDetailOfflineCacheDiagnosticRunning.value = true
  openDiagnosticResultDialog({
    title: '单篇离线缓存结果',
    message: '正在启动单篇离线缓存测试...',
    tone: 'info',
    items: [
      { label: '流程', value: '单篇离线缓存测试' },
      {
        label: '跳过已采集记录',
        value: articleDetailOfflineCacheSkipCollectedRecords.value ? '开启' : '关闭',
      },
      {
        label: '带状态（bate）',
        value: articleDetailOfflineCacheStateful.value ? '开启' : '关闭',
      },
      { label: '存储文章详情', value: '开启（锁定）' },
      { label: '离线归档内容', value: '开启（锁定）' },
      { label: '状态', value: '启动中' },
    ],
  })

  try {
    const started = await startArticleDetailOfflineCacheDiagnostic({
      skipCollectedRecords: articleDetailOfflineCacheSkipCollectedRecords.value,
      statefulOfflineCache: articleDetailOfflineCacheStateful.value,
      storeArticleDetail: articleDetailOfflineCacheStoreArticleDetail.value,
      archiveOfflineContent: articleDetailOfflineCacheArchiveContent.value,
    })
    applyArticleDetailDiagnosticResult(started)
    const jobId = started.jobId
    let latest: ArticleDetailDiagnosticResult = started

    while (latest.status === 'running') {
      await delay(500)
      latest = await getArticleDetailOfflineCacheDiagnostic(jobId)
      applyArticleDetailDiagnosticResult(latest)
    }

    showConfigNotice(
      latest.message ?? '单篇离线缓存测试已完成。',
      latest.ok ? 'success' : 'warning',
    )
  } catch (error) {
    const message = error instanceof Error ? error.message : '单篇离线缓存测试失败。'
    openDiagnosticResultDialog({
      title: '单篇离线缓存结果',
      message,
      tone: 'error',
      items: [
        { label: '流程', value: '单篇离线缓存测试' },
        { label: '失败原因', value: message },
      ],
    })
    showConfigNotice(message, 'error')
  } finally {
    isArticleDetailOfflineCacheDiagnosticRunning.value = false
  }
}

function applyArticleDetailDiagnosticResult(result: ArticleDetailDiagnosticResult) {
  openDiagnosticResultDialog({
    title: result.title ?? '详情获取结果',
    message: result.message ?? '单篇文章详情获取正在执行...',
    tone: result.tone ?? (result.ok ? 'success' : 'info'),
    items: result.items ?? [],
    status: result.status,
    action: result.action,
  })
}

function applyWindowClickFlowDiagnosticResult(result: ArticleDetailDiagnosticResult) {
  // 窗口测试的操作和丢弃原因只写入诊断文件，弹窗保持为文章结果列表。
  const articleItems = (result.items ?? []).filter((item) => item.kind === 'article')
  openDiagnosticResultDialog({
    title: result.title ?? '主页内容读取结果',
    message: result.message ?? '主页内容读取正在执行...',
    tone: result.tone ?? (result.ok ? 'success' : 'info'),
    items: articleItems,
    status: result.status,
    action: result.action,
  })
}

function delay(milliseconds: number) {
  return new Promise<void>((resolve) => {
    window.setTimeout(resolve, milliseconds)
  })
}

function normalizeWindowDiagnosticScrollSteps(value = windowDiagnosticScrollSteps.value) {
  const parsedValue = Number(value)
  const safeValue = Number.isFinite(parsedValue) ? Math.trunc(parsedValue) : 1
  const normalizedValue = Math.min(200, Math.max(1, safeValue))
  windowDiagnosticScrollSteps.value = normalizedValue
  return normalizedValue
}

function handleWindowDiagnosticScrollStepsInput(event: Event) {
  const input = event.target as HTMLInputElement
  input.value = String(normalizeWindowDiagnosticScrollSteps(input.valueAsNumber))
}

function handleWindowDiagnosticScrollStepsNumberChange(value: number | string | null) {
  normalizeWindowDiagnosticScrollSteps(Number(value))
}

async function handleWindowDiagnosticAction(action: WindowDiagnosticAction) {
  if (isWindowDiagnosticRunning.value) {
    return
  }

  isWindowDiagnosticRunning.value = true
  openDiagnosticResultDialog({
    title: windowDiagnosticActionTitle(action),
    message: '正在执行窗口诊断...',
    tone: 'info',
    items: [
      { label: '动作', value: windowDiagnosticActionTitle(action) },
      { label: '状态', value: '执行中' },
    ],
  })

  try {
    const options: WindowDiagnosticOptions = action === 'scroll-page'
      ? { scrollSteps: normalizeWindowDiagnosticScrollSteps() }
      : {}
    const result = await runWindowDiagnosticAction(action, options)
    applyWindowDiagnosticResult(result)
    showConfigNotice(result.message ?? '窗口诊断已完成。', result.ok ? 'success' : 'warning')
  } catch (error) {
    const message = error instanceof Error ? error.message : '窗口诊断执行失败。'
    openDiagnosticResultDialog({
      title: windowDiagnosticActionTitle(action),
      message,
      tone: 'error',
      items: [
        { label: '动作', value: windowDiagnosticActionTitle(action) },
        { label: '失败原因', value: message },
      ],
    })
    showConfigNotice(message, 'error')
  } finally {
    isWindowDiagnosticRunning.value = false
  }
}

function applyWindowDiagnosticResult(result: WindowDiagnosticResult) {
  openDiagnosticResultDialog({
    title: result.title ?? windowDiagnosticActionTitle(result.action),
    message: result.message ?? '窗口诊断已返回结果。',
    tone: result.tone ?? (result.ok ? 'success' : 'error'),
    items: result.items ?? [],
    status: result.status,
    action: result.action,
  })
}

function windowDiagnosticActionTitle(action: WindowDiagnosticAction) {
  const titles: Record<WindowDiagnosticAction, string> = {
    'read-home': '读取主页结果',
    'activate-home': '激活主页结果',
    'first-article-click': '首篇点击结果',
    'scroll-page': '滚动页面结果',
    'bounce-scroll': '回弹滚动结果',
    'close-tab': '关闭标签结果',
  }
  return titles[action]
}

function emitTaskStatus(status: TaskStatus) {
  runtimeStatus.value = status
  emit('taskStatusChanged', status)
}

function applyProxyStatusSnapshot(status: TaskStatus | null | undefined) {
  if (!status) {
    return
  }

  // 这里只同步代理运行态；是否允许采集时接管系统代理属于独立配置。
  configForm.proxyHost = status.proxy?.host ?? configForm.proxyHost
  configForm.proxyPort = status.proxy?.port ?? configForm.proxyPort
  settings.autoStartProxy = Boolean(status.proxy?.mitmEnabled)
}

async function hydrateRuntimePaths() {
  const result = await listRuntimePaths()
  if (!result.ok) {
    showConfigNotice(result.message ?? '读取运行目录失败。', 'error')
    return
  }

  configForm.projectDir = result.paths.projectDir
  configForm.storageDir = result.paths.storageDir
  configForm.logDir = result.paths.logDir
}

async function handleOpenRuntimePath(key: RuntimePathKey) {
  try {
    const result = await openRuntimePath(key)
    showConfigNotice(result.message ?? (result.ok ? '已打开目录。' : '打开目录失败。'), result.ok ? 'success' : 'error')
  } catch (error) {
    showConfigNotice(error instanceof Error ? error.message : '打开目录失败。', 'error')
  }
}

function applyRuntimeConfigValues(values?: Record<string, string>) {
  if (!values) {
    return
  }

  for (const [key, value] of Object.entries(values)) {
    runtimeConfigValues[key] = String(value)
  }

  for (const key of Object.keys(settingsNumberValues)) {
    settingsNumberValues[key] = parseRuntimeNumberValue(key, settingsNumberValues[key] ?? 0)
  }
  for (const key of Object.keys(settingsToggleValues)) {
    settingsToggleValues[key] = parseRuntimeSwitchValue(key, settingsToggleValues[key] ?? false)
  }

  configForm.logLevel = (runtimeConfigValues['basic_settings.runtime_maintenance.log_level'] as RuntimeLogLevel | undefined) ?? configForm.logLevel
  configForm.requestIntervalSeconds = parseRuntimeNumberValue('main_flow.dispatch_control.single_task_interval_seconds', configForm.requestIntervalSeconds)
  configForm.proxyHost = runtimeConfigValues['basic_settings.proxy_settings.host'] ?? configForm.proxyHost
  configForm.proxyPort = parseRuntimeNumberValue('basic_settings.proxy_settings.port', configForm.proxyPort)
  configForm.startupDelaySeconds = parseRuntimeNumberValue('basic_settings.proxy_settings.startup_delay_seconds', configForm.startupDelaySeconds)
  configForm.trafficCheckUrl = runtimeConfigValues['basic_settings.proxy_settings.verification_url'] ?? configForm.trafficCheckUrl
  settings.autoCleanTempFiles = parseRuntimeSwitchValue('basic_settings.runtime_maintenance.auto_clean_temp_files', settings.autoCleanTempFiles)
  settings.enableSystemProxy = parseRuntimeSwitchValue('basic_settings.proxy_settings.enable_system_proxy', settings.enableSystemProxy)
}

function replaceRuntimeConfigValues(values?: Record<string, string>) {
  for (const key of Object.keys(runtimeConfigValues)) {
    delete runtimeConfigValues[key]
  }
  applyRuntimeConfigValues(values)
}

function formatSwitchValue(enabled: boolean) {
  return enabled ? '开启' : '关闭'
}

function buildRuntimeConfigValuesPayload() {
  const editableRuntimeConfigValues = Object.fromEntries(
    Object.entries(runtimeConfigValues).filter(([key]) => !readonlyRangeConfigKeys.has(key)),
  )

  return {
    ...editableRuntimeConfigValues,
    ...Object.fromEntries(
      Object.entries(settingsNumberValues).map(([key, value]) => [key, String(value)]),
    ),
    ...Object.fromEntries(
      Object.entries(settingsToggleValues).map(([key, value]) => [key, formatSwitchValue(value)]),
    ),
    'basic_settings.runtime_maintenance.log_level': configForm.logLevel,
    'basic_settings.runtime_maintenance.auto_clean_temp_files': formatSwitchValue(settings.autoCleanTempFiles),
    'basic_settings.proxy_settings.host': configForm.proxyHost,
    'basic_settings.proxy_settings.port': String(configForm.proxyPort),
    'basic_settings.proxy_settings.startup_delay_seconds': String(configForm.startupDelaySeconds),
    'basic_settings.proxy_settings.verification_url': configForm.trafficCheckUrl,
    'basic_settings.proxy_settings.enable_system_proxy': formatSwitchValue(settings.enableSystemProxy),
  }
}

async function syncRuntimeConfigToMemory() {
  if (isHydratingConfig.value || isApplyingDefaults.value) {
    return
  }

  const result = await updateRuntimeConfig(buildRuntimeConfigPayload())
  if (result.taskStatus) {
    emitTaskStatus(result.taskStatus)
  }
  if (!result.ok) {
    showConfigNotice(result.message ?? '同步运行配置失败。', 'error')
  }
}

function queueRuntimeConfigMemorySync() {
  if (isHydratingConfig.value || isApplyingDefaults.value) {
    return
  }

  window.clearTimeout(runtimeConfigSyncTimer)
  runtimeConfigSyncTimer = window.setTimeout(() => {
    void syncRuntimeConfigToMemory()
  }, 250)
}

async function syncProxySwitchState() {
  isSyncingProxyState.value = true
  try {
    const status = await getTaskStatus()
    emitTaskStatus(status)
    if (status.config) {
      isHydratingConfig.value = true
      applyRuntimeConfigValues(status.config.values)
      settings.autoSaveContent = Boolean(status.config.autoSaveContent)
      settings.autoCleanTempFiles = Boolean(status.config.autoCleanTempFiles)
      settings.enableSystemProxy = Boolean(status.config.enableSystemProxy ?? settings.enableSystemProxy)
      configForm.logLevel = status.config.logLevel ?? configForm.logLevel
      configForm.requestIntervalSeconds = Math.round(status.config.requestIntervalSeconds ?? configForm.requestIntervalSeconds)
      configForm.startupDelaySeconds = Math.round(status.config.startupDelaySeconds ?? configForm.startupDelaySeconds)
      configForm.trafficCheckUrl = status.config.verificationUrl ?? configForm.trafficCheckUrl
      await nextTick()
      isHydratingConfig.value = false
    }

    applyProxyStatusSnapshot(status)
  } finally {
    isHydratingConfig.value = false
    isSyncingProxyState.value = false
  }
}

function buildRuntimeConfigPayload(): RuntimeConfigPayload {
  return {
    autoSaveContent: settings.autoSaveContent,
    autoCleanTempFiles: settings.autoCleanTempFiles,
    autoStartProxy: settings.autoStartProxy,
    enableSystemProxy: settings.enableSystemProxy,
    logLevel: configForm.logLevel,
    requestIntervalSeconds: settingsNumberValues['main_flow.dispatch_control.single_task_interval_seconds'] ?? configForm.requestIntervalSeconds,
    proxy: {
      host: configForm.proxyHost,
      port: configForm.proxyPort,
      startupDelaySeconds: configForm.startupDelaySeconds,
      verificationUrl: configForm.trafficCheckUrl,
    },
    values: buildRuntimeConfigValuesPayload(),
  }
}

async function handleSaveConfig() {
  if (isSavingConfig.value) {
    return
  }

  isSavingConfig.value = true
  try {
    // 保存入口只负责持久化配置，真正写入 custom.yaml 的逻辑由 Python 后端统一处理。
    const result = await saveRuntimeConfig(buildRuntimeConfigPayload())
    if (result.ok) {
      if (result.taskStatus) {
        emitTaskStatus(result.taskStatus)
      }
      showConfigNotice(`已保存到 ${result.configPath ?? 'custom.yaml'}`, 'success')
      return
    }

    showConfigNotice(result.message ?? '保存失败，请查看运行日志。', 'error')
  } catch (error) {
    showConfigNotice(error instanceof Error ? error.message : '保存失败，请查看运行日志。', 'error')
  } finally {
    isSavingConfig.value = false
  }
}

async function handleClearCache() {
  if (isClearingCache.value) {
    return
  }

  isClearingCache.value = true
  try {
    const backendResult = await clearRuntimeCache()

    if ('caches' in window) {
      const cacheNames = await caches.keys()
      await Promise.all(cacheNames.map((cacheName) => caches.delete(cacheName)))
    }

    cacheCleaned.value = backendResult.ok
    showConfigNotice(
      backendResult.message ?? '已清理临时目录和浏览器 Cache Storage。',
      backendResult.ok ? 'success' : 'warning',
    )
    window.clearTimeout(cacheCleanedTimer)
    cacheCleanedTimer = window.setTimeout(() => {
      cacheCleaned.value = false
    }, 1600)
  } catch (error) {
    showConfigNotice(error instanceof Error ? error.message : '清理缓存失败。', 'error')
    cacheCleaned.value = false
  } finally {
    isClearingCache.value = false
  }
}

function startupSelfCheckTone(result: StartupSelfCheckResult): DiagnosticResultTone {
  if (!result.ok) {
    return 'error'
  }
  return result.warningCount > 0 ? 'warning' : 'success'
}

function startupSelfCheckStatusLabel(status: string) {
  const labels: Record<string, string> = {
    passed: '通过',
    passed_with_warnings: '有警告',
    failed: '异常',
    interrupted: '已中断',
  }
  return labels[status] ?? status
}

function buildStartupSelfCheckDiagnosticItems(result: StartupSelfCheckResult): DiagnosticResultItem[] {
  return result.items.map((item) => createDiagnosticResultRow([
    { label: '检查项', value: item.label },
    { label: '状态', value: startupSelfCheckStatusLabel(item.status) },
    { label: '结果', value: item.message || '无' },
    { label: '处理方式', value: item.action || '无需处理' },
  ]))
}

async function handleRunStartupSelfCheck() {
  if (isRunningStartupSelfCheck.value) {
    return
  }

  isRunningStartupSelfCheck.value = true
  openDiagnosticResultDialog({
    title: '启动自检',
    message: '正在自检：检查程序运行环境、MITM、Playwright、SQLite 和存储配置...',
    tone: 'info',
    items: [{ label: '启动自检', value: '正在执行' }],
  })

  try {
    const result = await runStartupSelfCheck()
    openDiagnosticResultDialog({
      title: '启动自检',
      message: result.ok
        ? '自检完成：' + result.fatalCount + ' 个致命问题，' + result.warningCount + ' 个警告。'
        : '自检发现 ' + result.fatalCount + ' 个致命问题，' + result.warningCount + ' 个警告。',
      tone: startupSelfCheckTone(result),
      status: result.status,
      items: buildStartupSelfCheckDiagnosticItems(result),
    })
    showConfigNotice(result.ok ? '启动自检完成。' : '启动自检发现异常，请查看弹窗。', result.ok ? 'success' : 'warning')
  } catch (error) {
    openDiagnosticResultDialog({
      title: '启动自检',
      message: error instanceof Error ? error.message : '启动自检失败。',
      tone: 'error',
      items: [{ label: '错误', value: error instanceof Error ? error.message : '启动自检失败。' }],
    })
    showConfigNotice('启动自检失败。', 'error')
  } finally {
    isRunningStartupSelfCheck.value = false
  }
}

function closeMitmCertificateDialog() {
  if (isCaCertificateDialogBusy.value) {
    return
  }

  caCertificateDialogVisible.value = false
}

async function handleOpenMitmCertificateDialog() {
  if (isListingMitmCertificates.value || isDeletingMitmCertificates.value || isCheckingCaCertificate.value) {
    return
  }

  openCaCertificateDialog({
    mode: 'delete',
    phase: 'checking',
    message: '正在检测项目 CA 证书和系统证书库...',
  })
  isListingMitmCertificates.value = true
  try {
    const result = await checkCaCertificate()
    applyCaCertificateStatusPayload(result)
    caCertificateDialogPhase.value = mitmCertificateItems.value.length ? 'ready' : 'done'
    mitmCertificateMessage.value = result.message ?? (mitmCertificateItems.value.length > 0
      ? `已检索到 ${mitmCertificateItems.value.length} 张系统 MITM 证书，确认后会删除这些列出的系统证书。`
      : '未检索到系统 MITM 相关证书，无需删除。')
    showConfigNotice(
      mitmCertificateMessage.value,
      result.ok ? 'info' : 'error',
    )
  } catch (error) {
    caCertificateDialogPhase.value = 'error'
    mitmCertificateMessage.value = error instanceof Error ? error.message : '检索 MITM 证书失败。'
    showConfigNotice(mitmCertificateMessage.value, 'error')
  } finally {
    isListingMitmCertificates.value = false
  }
}

async function handleConfirmDeleteMitmCertificates() {
  if (isDeletingMitmCertificates.value) {
    return
  }

  const thumbprints = mitmCertificateItems.value.map((item) => item.thumbprint).filter(Boolean)
  if (!thumbprints.length) {
    caCertificateDialogPhase.value = 'done'
    mitmCertificateMessage.value = '未检索到可删除的 MITM 证书。'
    showConfigNotice(mitmCertificateMessage.value, 'warning')
    return
  }

  isDeletingMitmCertificates.value = true
  caCertificateDialogPhase.value = 'deleting'
  mitmCertificateMessage.value = '正在删除弹窗中列出的系统 MITM 证书...'
  try {
    const result = await deleteMitmCaCertificates(thumbprints)
    caCertificateDialogPhase.value = result.ok ? 'done' : 'error'
    caCertificateDialogDeletedItems.value = result.deleted ?? []
    caCertificateDialogSkippedItems.value = result.skipped ?? []
    mitmCertificateItems.value = result.remainingCertificates ?? []
    mitmCertificateMessage.value = result.message ?? 'MITM 证书删除完成。'
    showConfigNotice(mitmCertificateMessage.value, result.ok ? 'success' : 'warning')
    try {
      const refreshStatus = await checkCaCertificate()
      applyCaCertificateStatusPayload(refreshStatus)
    } catch (refreshError) {
      showConfigNotice(
        refreshError instanceof Error ? `证书删除完成，但刷新状态失败：${refreshError.message}` : '证书删除完成，但刷新状态失败。',
        'warning',
      )
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : '删除 MITM 证书失败。'
    caCertificateDialogPhase.value = 'error'
    mitmCertificateMessage.value = message
    showConfigNotice(message, 'error')
  } finally {
    isDeletingMitmCertificates.value = false
  }
}

function handleConfigFieldNotice(label: string, value: string | number | boolean, tone: 'info' | 'success' | 'warning' | 'error' = 'info') {
  if (isSyncingProxyState.value || isHydratingConfig.value || isApplyingDefaults.value) {
    return
  }

  const normalizedValue = typeof value === 'boolean' ? (value ? '开启' : '关闭') : value
  showConfigNotice(`${label}已修改为 ${normalizedValue}`, tone)
  queueRuntimeConfigMemorySync()
}

function handleResetDefaults() {
  if (isApplyingDefaults.value || isSavingConfig.value) {
    return
  }

  if (isRuntimeCacheBusy.value) {
    showConfigNotice('当前有采集或诊断任务正在运行，暂不能恢复默认配置。', 'warning')
    return
  }

  resetDefaultsDialogVisible.value = true
}

function closeResetDefaultsDialog() {
  if (!isApplyingDefaults.value) {
    resetDefaultsDialogVisible.value = false
  }
}

async function confirmResetDefaults() {
  if (isApplyingDefaults.value) {
    return
  }

  if (isRuntimeCacheBusy.value) {
    resetDefaultsDialogVisible.value = false
    showConfigNotice('当前有采集或诊断任务正在运行，暂不能恢复默认配置。', 'warning')
    return
  }

  window.clearTimeout(runtimeConfigSyncTimer)
  isApplyingDefaults.value = true
  isHydratingConfig.value = true
  try {
    const result = await resetRuntimeConfig()
    if (!result.ok) {
      showConfigNotice(result.message ?? '恢复系统默认配置失败。', 'error')
      return
    }

    if (result.taskStatus) {
      emitTaskStatus(result.taskStatus)
      replaceRuntimeConfigValues(result.taskStatus.config?.values)
      settings.autoSaveContent = Boolean(result.taskStatus.config?.autoSaveContent)
      settings.autoCleanTempFiles = Boolean(result.taskStatus.config?.autoCleanTempFiles)
      settings.enableSystemProxy = Boolean(result.taskStatus.config?.enableSystemProxy ?? settings.enableSystemProxy)
      configForm.logLevel = result.taskStatus.config?.logLevel ?? configForm.logLevel
      configForm.requestIntervalSeconds = Number(result.taskStatus.config?.requestIntervalSeconds ?? configForm.requestIntervalSeconds)
      configForm.startupDelaySeconds = Number(result.taskStatus.config?.startupDelaySeconds ?? configForm.startupDelaySeconds)
      configForm.trafficCheckUrl = result.taskStatus.config?.verificationUrl ?? configForm.trafficCheckUrl
      applyProxyStatusSnapshot(result.taskStatus)
    }

    await nextTick()
    resetDefaultsDialogVisible.value = false
    const backupMessage = result.backupPath ? `，原配置已备份到 ${result.backupPath}` : ''
    showConfigNotice(`已恢复系统默认配置${backupMessage}`, 'success')
  } catch (error) {
    showConfigNotice(error instanceof Error ? error.message : '恢复系统默认配置失败。', 'error')
  } finally {
    isHydratingConfig.value = false
    isApplyingDefaults.value = false
  }
}

async function handleTestProxyConnection() {
  if (isTestingProxyConnection.value) {
    return
  }

  isTestingProxyConnection.value = true
  diagnosticResultItems.value = createHttpsCheckDiagnosticItems()
  openDiagnosticResultDialog({
    title: 'HTTPS 校验结果',
    message: '正在开启 MITM 代理...',
    tone: 'info',
    items: diagnosticResultItems.value,
  })

  let shouldDisableSystemProxy = false
  let shouldStopMitmProxy = false
  let finalMessage = 'HTTPS 校验流程已结束。'
  let finalTone: DiagnosticResultTone = 'info'
  try {
    diagnosticResultItems.value = updateDiagnosticItem(diagnosticResultItems.value, 'MITM 代理', '正在开启 MITM 代理...')
    const mitmStartResult = await startMitmProxy()
    emitTaskStatus(mitmStartResult)
    shouldStopMitmProxy = Boolean(mitmStartResult.ok)
    diagnosticResultItems.value = updateDiagnosticItem(
      diagnosticResultItems.value,
      'MITM 代理',
      mitmStartResult.ok ? '已开启' : '开启失败：' + (mitmStartResult.message ?? mitmStartResult.status),
    )
    if (!mitmStartResult.ok) {
      throw new Error(mitmStartResult.message ?? 'MITM 代理启动失败。')
    }

    diagnosticResultMessage.value = '正在开启系统代理...'
    diagnosticResultItems.value = updateDiagnosticItem(diagnosticResultItems.value, '系统代理', '正在开启系统代理...')
    const systemEnableResult = await enableSystemProxy()
    emitTaskStatus(systemEnableResult)
    shouldDisableSystemProxy = Boolean(systemEnableResult.ok)
    diagnosticResultItems.value = updateDiagnosticItem(
      diagnosticResultItems.value,
      '系统代理',
      systemEnableResult.ok ? '已开启' : '开启失败：' + (systemEnableResult.message ?? systemEnableResult.status),
    )
    if (!systemEnableResult.ok) {
      throw new Error(systemEnableResult.message ?? '系统代理开启失败。')
    }

    diagnosticResultMessage.value = '正在测试 HTTPS 是否已监听...'
    diagnosticResultItems.value = updateDiagnosticItem(diagnosticResultItems.value, 'HTTPS 校验', '正在测试 HTTPS 是否已监听...')
    const result = await testProxyConnection()
    diagnosticResultItems.value = updateDiagnosticItemCells(
      diagnosticResultItems.value,
      'MITM 代理',
      [
        { label: 'MITM 代理', value: '已开启' },
        { label: '代理地址', value: formatDiagnosticValue(result.proxy) },
      ],
    )
    diagnosticResultItems.value = updateDiagnosticItemCells(
      diagnosticResultItems.value,
      'HTTPS 校验',
      [
        { label: '验证地址', value: formatDiagnosticValue(result.url) },
        { label: 'HTTP 状态', value: formatDiagnosticValue(result.statusCode) },
      ],
    )
    finalMessage = result.message ?? (result.ok ? '代理连接测试通过。' : '代理连接测试失败。')
    finalTone = result.ok ? 'success' : 'error'
  } catch (error) {
    finalMessage = error instanceof Error ? error.message : '代理连接测试失败。'
    finalTone = 'error'
    if (!diagnosticResultDialogVisible.value) {
      openDiagnosticResultDialog({
        title: 'HTTPS 校验结果',
        message: finalMessage,
        tone: finalTone,
        items: diagnosticResultItems.value,
      })
    }
  } finally {
    diagnosticResultMessage.value = finalMessage + ' 正在恢复代理状态...'

    if (shouldDisableSystemProxy) {
      diagnosticResultItems.value = updateDiagnosticItem(diagnosticResultItems.value, '关闭系统代理', '正在关闭系统代理...')
      try {
        const systemDisableResult = await disableSystemProxy()
        emitTaskStatus(systemDisableResult)
        diagnosticResultItems.value = updateDiagnosticItem(
          diagnosticResultItems.value,
          '关闭系统代理',
          systemDisableResult.ok ? (systemDisableResult.message ?? '已关闭') : '关闭失败：' + (systemDisableResult.message ?? systemDisableResult.status),
        )
        if (!systemDisableResult.ok && finalTone === 'success') {
          finalTone = 'warning'
          finalMessage = finalMessage + ' 但系统代理关闭失败：' + (systemDisableResult.message ?? systemDisableResult.status)
        }
      } catch (cleanupError) {
        const cleanupMessage = cleanupError instanceof Error ? cleanupError.message : '系统代理关闭失败。'
        diagnosticResultItems.value = updateDiagnosticItem(diagnosticResultItems.value, '关闭系统代理', '关闭失败：' + cleanupMessage)
        if (finalTone === 'success') {
          finalTone = 'warning'
          finalMessage = finalMessage + ' 但系统代理关闭失败：' + cleanupMessage
        }
      }
    } else {
      diagnosticResultItems.value = updateDiagnosticItem(diagnosticResultItems.value, '关闭系统代理', '未开启，无需关闭')
    }

    if (shouldStopMitmProxy) {
      diagnosticResultItems.value = updateDiagnosticItem(diagnosticResultItems.value, '关闭 MITM 代理', '正在关闭 MITM 代理...')
      try {
        const mitmStopResult = await stopMitmProxy()
        emitTaskStatus(mitmStopResult)
        diagnosticResultItems.value = updateDiagnosticItem(
          diagnosticResultItems.value,
          '关闭 MITM 代理',
          mitmStopResult.ok ? '已关闭' : '关闭失败：' + (mitmStopResult.message ?? mitmStopResult.status),
        )
        if (!mitmStopResult.ok && finalTone === 'success') {
          finalTone = 'warning'
          finalMessage = finalMessage + ' 但 MITM 代理关闭失败：' + (mitmStopResult.message ?? mitmStopResult.status)
        }
      } catch (cleanupError) {
        const cleanupMessage = cleanupError instanceof Error ? cleanupError.message : 'MITM 代理关闭失败。'
        diagnosticResultItems.value = updateDiagnosticItem(diagnosticResultItems.value, '关闭 MITM 代理', '关闭失败：' + cleanupMessage)
        if (finalTone === 'success') {
          finalTone = 'warning'
          finalMessage = finalMessage + ' 但 MITM 代理关闭失败：' + cleanupMessage
        }
      }
    } else {
      diagnosticResultItems.value = updateDiagnosticItem(diagnosticResultItems.value, '关闭 MITM 代理', '未开启，无需关闭')
    }

    try {
      await syncProxySwitchState()
    } catch (syncError) {
      const syncMessage = syncError instanceof Error ? syncError.message : '代理状态刷新失败。'
      if (finalTone === 'success') {
        finalTone = 'warning'
      }
      finalMessage = finalMessage + ' ' + syncMessage
    }

    diagnosticResultTone.value = finalTone
    diagnosticResultMessage.value = finalMessage
    showConfigNotice(finalMessage, finalTone === 'success' ? 'success' : finalTone === 'warning' ? 'warning' : 'error')
    isTestingProxyConnection.value = false
  }
}

function buildUnknownCaStatus(message: string): CaCertificateStatus {
  return {
    ok: false,
    status: 'unknown',
    installed: false,
    label: '无法检测',
    message,
  }
}

async function handleCheckCaCertificate() {
  if (isCheckingCaCertificate.value) {
    return
  }

  openCaCertificateDialog({
    mode: 'check',
    phase: 'checking',
    message: '正在检测项目 CA 证书和系统证书库...',
  })
  isCheckingCaCertificate.value = true
  try {
    const result = await checkCaCertificate()
    applyCaCertificateStatusPayload(result)
    caCertificateDialogPhase.value = result.ok ? 'done' : 'error'
    caCertificateDialogMessage.value = result.message ?? result.label
    showConfigNotice(caCertificateDialogMessage.value, result.ok ? (result.installed ? 'success' : 'warning') : 'error')
  } catch (error) {
    caCertificateStatus.value = buildUnknownCaStatus(
      error instanceof Error ? error.message : '检测 CA 证书失败，请查看运行日志。',
    )
    caCertificateDialogPhase.value = 'error'
    caCertificateDialogMessage.value = caCertificateStatus.value.message ?? '检测 CA 证书失败。'
    showConfigNotice(caCertificateDialogMessage.value, 'error')
  } finally {
    isCheckingCaCertificate.value = false
  }
}

async function openInstallCaCertificateDialog() {
  if (isInstallingCaCertificate.value || isCheckingCaCertificate.value) {
    return
  }

  openCaCertificateDialog({
    mode: 'install',
    phase: 'checking',
    message: '正在读取当前项目 CA 证书信息...',
  })
  isCheckingCaCertificate.value = true
  try {
    const result = await checkCaCertificate()
    applyCaCertificateStatusPayload(result)
    if (!result.caFileExists) {
      caCertificateDialogPhase.value = 'error'
      caCertificateDialogMessage.value = result.message ?? '当前项目 CA 证书文件不存在，无法安装。'
    } else if (result.projectCertificateInstalled) {
      caCertificateDialogPhase.value = 'done'
      caCertificateDialogMessage.value = '项目 CA 证书已存在于当前用户根证书库，无需重复安装。'
    } else {
      caCertificateDialogPhase.value = 'ready'
      caCertificateDialogMessage.value = '请确认将当前项目 CA 证书安装到当前用户根证书库。'
    }
  } catch (error) {
    caCertificateStatus.value = buildUnknownCaStatus(
      error instanceof Error ? error.message : '读取 CA 证书信息失败。',
    )
    caCertificateDialogPhase.value = 'error'
    caCertificateDialogMessage.value = caCertificateStatus.value.message ?? '读取 CA 证书信息失败。'
    showConfigNotice(caCertificateDialogMessage.value, 'error')
  } finally {
    isCheckingCaCertificate.value = false
  }
}

function closeInstallCaCertificateDialog() {
  if (isCaCertificateDialogBusy.value) {
    return
  }

  caCertificateDialogVisible.value = false
}

async function confirmInstallCaCertificate() {
  if (isInstallingCaCertificate.value) {
    return
  }

  isInstallingCaCertificate.value = true
  caCertificateDialogPhase.value = 'installing'
  caCertificateDialogMessage.value = '正在安装项目 CA 证书到当前用户根证书库...'
  try {
    const result = await installCaCertificate()
    applyCaCertificateStatusPayload(result)
    caCertificateDialogPhase.value = result.ok ? 'done' : 'error'
    caCertificateDialogMessage.value = result.message ?? (result.ok ? 'CA 证书已安装。' : 'CA 证书安装失败。')
    openDiagnosticResultDialog({
      title: '安装 CA 证书结果',
      message: caCertificateDialogMessage.value,
      tone: result.ok ? 'success' : 'error',
      items: buildDiagnosticItems(result as unknown as Record<string, unknown>, [
        { key: 'status', label: '状态' },
        { key: 'currentCaPath', label: '证书文件' },
        { key: 'storePath', label: '安装位置' },
        { key: 'storeCertificateCount', label: '证书库数量' },
      ]),
    })
    showConfigNotice(caCertificateDialogMessage.value, result.ok ? 'success' : 'error')
  } catch (error) {
    caCertificateStatus.value = buildUnknownCaStatus(
      error instanceof Error ? error.message : '安装 CA 证书失败。',
    )
    caCertificateDialogPhase.value = 'error'
    caCertificateDialogMessage.value = caCertificateStatus.value.message ?? '安装 CA 证书失败。'
    openDiagnosticResultDialog({
      title: '安装 CA 证书结果',
      message: caCertificateDialogMessage.value,
      tone: 'error',
    })
    showConfigNotice(caCertificateDialogMessage.value, 'error')
  } finally {
    isInstallingCaCertificate.value = false
  }
}

watch(
  () => settings.enableSystemProxy,
  (enabled, previousValue) => {
    if (enabled !== previousValue) {
      handleConfigFieldNotice('采集时接管系统代理', enabled, enabled ? 'success' : 'warning')
    }
  },
)

watch(
  () => configForm.requestIntervalSeconds,
  (value, previousValue) => {
    if (value !== previousValue) {
      handleConfigFieldNotice('请求间隔时间', `${value} 秒`)
    }
  },
)

watch(
  () => configForm.proxyPort,
  (value, previousValue) => {
    if (value !== previousValue) {
      handleConfigFieldNotice('代理端口', value)
    }
  },
)

watch(
  () => configForm.startupDelaySeconds,
  (value, previousValue) => {
    if (value !== previousValue) {
      handleConfigFieldNotice('启动延迟', `${value} 秒`)
    }
  },
)

watch(
  () => configForm.trafficCheckUrl,
  (value, previousValue) => {
    if (value !== previousValue) {
      handleConfigFieldNotice('检测地址', value)
    }
  },
)

watch(
  () => configForm.logLevel,
  (value, previousValue) => {
    if (value !== previousValue) {
      handleConfigFieldNotice('日志等级', value)
    }
  },
)

watch(
  () => configForm.fileNameMode,
  (value, previousValue) => {
    if (value !== previousValue) {
      handleConfigFieldNotice('命名规则', value)
    }
  },
)

watch(
  () => settings.autoSaveContent,
  (enabled, previousValue) => {
    if (enabled !== previousValue) {
      handleConfigFieldNotice('内容自动保存', enabled, enabled ? 'success' : 'warning')
    }
  },
)

watch(
  () => settings.autoCleanTempFiles,
  (enabled, previousValue) => {
    if (enabled !== previousValue) {
      handleConfigFieldNotice('自动清理临时文件', enabled, enabled ? 'success' : 'warning')
    }
  },
)

watch(
  settingsToggleValues,
  (values) => {
    if (isSyncingProxyState.value || isHydratingConfig.value || isApplyingDefaults.value) {
      return
    }

    for (const [key, enabled] of Object.entries(values)) {
      runtimeConfigValues[key] = formatSwitchValue(enabled)
    }
    queueRuntimeConfigMemorySync()
  },
  { deep: true },
)

watch(
  currentTaskStatus,
  (status) => {
    if (isSyncingProxyState.value || isHydratingConfig.value) {
      return
    }

    applyProxyStatusSnapshot(status)
  },
  { immediate: true },
)

async function refreshOnActivated() {
  await Promise.all([hydrateRuntimePaths(), syncProxySwitchState()])
}

defineExpose({
  refreshOnActivated,
})
</script>

<template>
  <section class="management-page settings-page" aria-label="系统配置">
    <section class="settings-three-pane">
      <nav class="settings-primary-nav" aria-label="设置类别">
        <header class="settings-nav-header">
          <strong>设置类别</strong>
          <span>按 custom.yaml 分组</span>
        </header>
        <button
          v-for="category in settingsCategories"
          :key="category.key"
          :class="['settings-primary-item', { active: category.key === selectedSettingsCategory.key }]"
          type="button"
          @click="selectSettingsCategory(category.key)"
        >
          <AppIcon :icon="['settings-primary-icon', category.icon]" />
          <span>
            <strong>{{ category.label }}</strong>
          </span>
        </button>
      </nav>

      <nav class="settings-secondary-nav" aria-label="二级设置项">
        <header class="settings-nav-header">
          <strong>{{ selectedSettingsCategory.label }}</strong>
          <span>{{ selectedSettingsCategory.summary }}</span>
        </header>
        <button
          v-for="item in activeSettingsItems"
          :key="item.key"
          :class="['settings-secondary-item', { active: selectedSettingsItem?.key === item.key }]"
          type="button"
          @click="selectSettingsItem(item.key)"
        >
          <AppIcon :icon="['settings-secondary-icon', item.icon]" />
          <span>
            <strong>{{ item.label }}</strong>
          </span>
        </button>
      </nav>

      <section class="settings-detail-pane" aria-live="polite">
        <header :class="['settings-detail-header', { empty: !selectedSettingsItem }]">
          <div v-if="selectedSettingsItem" class="settings-detail-title-row">
            <span class="settings-detail-icon">
              <AppIcon :icon="selectedSettingsItem.icon" />
            </span>
            <h2>{{ selectedSettingsItem.label }}</h2>
          </div>
          <h2 v-else>请选择左侧具体设置项</h2>
          <small>{{ selectedSettingsItem?.summary ?? selectedSettingsCategory.description }}</small>
        </header>

        <template v-if="selectedSettingsItem">
          <section v-if="selectedSettingsDetail" class="settings-detail-section">
            <div class="settings-detail-content">
              <div v-if="selectedSettingsDetail.controls?.length" class="settings-config-list">
                <article
                  v-for="control in selectedSettingsDetail.controls"
                  :key="control.kind"
                  :class="['settings-config-row', getSettingsControlLayoutClass(control)]"
                >
                  <div class="settings-config-copy">
                    <strong>{{ control.label }}</strong>
                    <small>{{ control.description }}</small>
                  </div>
                  <code v-if="control.configKey" class="settings-config-keyline">{{ getSettingsConfigKeyPath(control.configKey) }}</code>

                  <div :class="['settings-config-control', getSettingsControlLayoutClass(control)]">
                    <ASelect
                      v-if="control.kind === 'log-level'"
                      v-model:value="configForm.logLevel"
                      class="settings-ant-control"
                      :options="logLevelOptions"
                      popup-class-name="settings-select-panel"
                    />
                    <AInput
                      v-else-if="control.kind === 'proxy-host'"
                      v-model:value="configForm.proxyHost"
                      class="settings-ant-control"
                      placeholder="127.0.0.1"
                    />
                    <AInput
                      v-else-if="control.kind === 'verification-url'"
                      v-model:value="configForm.trafficCheckUrl"
                      class="settings-ant-control"
                      placeholder="https://mitm.it/"
                    />
                    <AInputNumber
                      v-else-if="control.kind === 'proxy-port'"
                      class="settings-ant-number proxy-number-input"
                      :value="configForm.proxyPort"
                      :min="numericConfigLimits.proxyPort.min"
                      :max="numericConfigLimits.proxyPort.max"
                      :step="numericConfigLimits.proxyPort.step"
                      :precision="0"
                      :controls="true"
                      aria-label="代理端口"
                      @change="handleNumericConfigNumberChange('proxyPort', $event)"
                    />
                    <AInputNumber
                      v-else-if="control.kind === 'startup-delay'"
                      class="settings-ant-number proxy-number-input"
                      :value="configForm.startupDelaySeconds"
                      :min="numericConfigLimits.startupDelaySeconds.min"
                      :max="numericConfigLimits.startupDelaySeconds.max"
                      :step="numericConfigLimits.startupDelaySeconds.step"
                      :precision="0"
                      :controls="true"
                      aria-label="端口等待时间"
                      @change="handleNumericConfigNumberChange('startupDelaySeconds', $event)"
                    />
                    <div v-else-if="control.kind === 'auto-clean'" class="control-switch-line">
                      <ASwitch v-model:checked="settings.autoCleanTempFiles" class="settings-ant-switch" checked-children="开" un-checked-children="关" />
                      <span>{{ settings.autoCleanTempFiles ? '开启' : '关闭' }}</span>
                    </div>
                    <div v-else-if="control.kind === 'system-proxy'" class="control-switch-line">
                      <ASwitch v-model:checked="settings.enableSystemProxy" class="settings-ant-switch" checked-children="开" un-checked-children="关" />
                      <span>{{ settings.enableSystemProxy ? '允许接管' : '禁止接管' }}</span>
                    </div>
                    <div v-else-if="control.kind === 'mitm-proxy'" class="control-switch-line">
                      <ASwitch :checked="settings.autoStartProxy" class="settings-ant-switch" checked-children="开" un-checked-children="关" disabled />
                      <span>{{ settings.autoStartProxy ? '采集中运行' : '当前未运行' }}</span>
                    </div>
                  </div>
                </article>
              </div>

              <div v-if="selectedSettingsDetail.fields?.length" class="settings-config-list">
                <article
                  v-for="field in selectedSettingsDetail.fields"
                  :key="field.configKey"
                  :class="['settings-config-row', field.tone ?? 'default', getSettingsFieldLayoutClass(field)]"
                >
                  <div class="settings-config-copy">
                    <strong>{{ getSettingsFieldLabel(field) }}</strong>
                    <small>{{ field.description }}</small>
                  </div>
                  <code class="settings-config-keyline">{{ getSettingsConfigKeyPath(field.configKey) }}</code>
                  <div :class="['settings-config-control', getSettingsFieldLayoutClass(field)]">
                    <div v-if="field.inputType === 'switch'" class="control-switch-line">
                      <ASwitch v-model:checked="settingsToggleValues[field.configKey]" class="settings-ant-switch" checked-children="开" un-checked-children="关" />
                      <span>{{ settingsToggleValues[field.configKey] ? '开启' : '关闭' }}</span>
                    </div>
                    <div v-else-if="field.inputType === 'readonly-range'" class="settings-range-readonly">
                      <label>
                        <span>{{ getSettingsRangeValue(field).startLabel }}</span>
                        <input
                          class="settings-inline-input"
                          type="text"
                          :value="getSettingsRangeValue(field).start"
                          readonly
                          :aria-label="field.label + getSettingsRangeValue(field).startLabel"
                          @focus="selectReadonlyInputText"
                          @click="selectReadonlyInputText"
                        >
                      </label>
                      <span class="settings-range-separator">~</span>
                      <label>
                        <span>{{ getSettingsRangeValue(field).endLabel }}</span>
                        <input
                          class="settings-inline-input"
                          type="text"
                          :value="getSettingsRangeValue(field).end"
                          readonly
                          :aria-label="field.label + getSettingsRangeValue(field).endLabel"
                          @focus="selectReadonlyInputText"
                          @click="selectReadonlyInputText"
                        >
                      </label>
                      <span v-if="field.unit" class="settings-range-unit">{{ field.unit }}</span>
                    </div>
                    <AInputNumber
                      v-else-if="field.inputType === 'number-stepper'"
                      class="settings-ant-number settings-field-stepper"
                      :value="getSettingsNumberValue(field)"
                      :min="field.min ?? 0"
                      :max="field.max ?? 999999"
                      :step="field.step ?? 1"
                      :precision="getSettingsNumberPrecision(field)"
                      :controls="true"
                      :aria-label="field.label"
                      @change="handleSettingsNumberFieldNumberChange(field, $event)"
                    />
                    <input
                      v-else
                      class="settings-inline-input"
                      :type="field.inputType ?? 'text'"
                      :value="getSettingsDisplayValue(field)"
                      :readonly="field.tone === 'readonly' || Boolean(field.browseAction)"
                      :aria-label="field.label"
                      @focus="selectReadonlyInputText"
                      @click="selectReadonlyInputText"
                      @paste.prevent
                      @cut.prevent
                      @drop.prevent
                    >
                    <AButton
                      v-if="field.browseAction"
                      class="settings-ant-button settings-row-button ghost"
                      html-type="button"
                      @click="handlePendingBrowseAction(field)"
                    >
                      {{ field.browseLabel ?? '浏览' }}
                    </AButton>
                  </div>
                </article>
              </div>

              <div v-if="selectedSettingsDetail.actions?.length" class="diagnostic-action-grid settings-config-list">
                <article
                  v-for="action in selectedSettingsDetail.actions"
                  :key="action.label"
                  :class="[
                    'settings-config-row',
                    'compact-control',
                    'diagnostic-action-row',
                    action.tone ?? 'ghost',
                    { 'window-click-flow-row': action.showWindowClickFlowOptions },
                  ]"
                >
                  <div
                    v-if="action.showWindowClickFlowOptions"
                    class="window-click-flow-layout"
                  >
                    <div class="window-click-flow-title">
                      <strong>{{ action.label }}</strong>
                    </div>
                    <div class="window-click-flow-description-line">
                      <div class="window-click-flow-description">
                        <small>{{ action.description }}</small>
                        <span v-if="action.detail" class="diagnostic-action-detail">{{ action.detail }}</span>
                      </div>
                      <AButton
                        :class="['settings-ant-button', 'settings-row-button', 'diagnostic-action-button', action.tone ?? 'ghost']"
                        html-type="button"
                        :disabled="action.disabled?.() ?? false"
                        @click="action.run()"
                      >
                        <template #icon>
                          <AppIcon class="diagnostic-action-icon" :icon="action.icon" />
                        </template>
                        {{ action.buttonLabel ?? action.label }}
                      </AButton>
                    </div>
                    <div class="window-click-flow-fields">
                      <label class="window-click-flow-field window-click-flow-field--mode">
                        <span>日期筛选</span>
                        <ASelect
                          v-model:value="windowClickFlowDateFilterMode"
                          class="settings-ant-control window-click-flow-select"
                          :options="windowClickFlowDateFilterOptions"
                          popup-class-name="settings-select-panel"
                          :disabled="isWindowClickFlowDiagnosticRunning"
                        />
                      </label>
                      <label class="window-click-flow-field window-click-flow-field--count">
                        <span>任务数量</span>
                        <AInputNumber
                          class="settings-ant-number window-click-flow-stepper"
                          :value="windowClickFlowMaxRecords"
                          :min="windowClickFlowUsesUnlimitedRecords ? 0 : 1"
                          :max="20"
                          :precision="0"
                          :controls="true"
                          :disabled="isWindowClickFlowDiagnosticRunning || windowClickFlowUsesUnlimitedRecords"
                          aria-label="窗口测试任务数量"
                          @change="handleWindowClickFlowMaxRecordsChange"
                        />
                      </label>
                      <label
                        v-if="windowClickFlowDateFilterMode === 'range' || windowClickFlowDateFilterMode === 'after'"
                        :class="[
                          'window-click-flow-field',
                          'window-click-flow-field--date',
                          { 'window-click-flow-field--range': windowClickFlowDateFilterMode === 'range' },
                        ]"
                      >
                        <span v-if="windowClickFlowDateFilterMode === 'range'">日期范围</span>
                        <span v-else>起始日期</span>
                        <ARangePicker
                          v-if="windowClickFlowDateFilterMode === 'range'"
                          v-model:value="windowClickFlowDateRangeValue"
                          class="settings-ant-control window-click-flow-date-range-picker"
                          :allow-clear="true"
                          :placeholder="['选择起始日期', '选择截止日期']"
                          separator=" ~ "
                          format="YYYY-MM-DD"
                          value-format="YYYY-MM-DD"
                          popup-class-name="window-click-flow-date-range-picker-panel"
                          :disabled="isWindowClickFlowDiagnosticRunning"
                          aria-label="窗口测试起始日期和截止日期"
                        />
                        <ADatePicker
                          v-else
                          v-model:value="windowClickFlowStartDate"
                          class="settings-ant-control window-click-flow-date-picker"
                          :allow-clear="true"
                          placeholder="选择起始日期"
                          format="YYYY-MM-DD"
                          value-format="YYYY-MM-DD"
                          popup-class-name="window-click-flow-date-picker-panel"
                          :disabled="isWindowClickFlowDiagnosticRunning"
                          aria-label="窗口测试起始日期"
                        />
                      </label>
                      <label
                        v-if="windowClickFlowDateFilterMode === 'before'"
                        class="window-click-flow-field window-click-flow-field--date"
                      >
                        <span>截止日期</span>
                        <ADatePicker
                          v-model:value="windowClickFlowEndDate"
                          class="settings-ant-control window-click-flow-date-picker"
                          :allow-clear="true"
                          placeholder="选择截止日期"
                          format="YYYY-MM-DD"
                          value-format="YYYY-MM-DD"
                          popup-class-name="window-click-flow-date-picker-panel"
                          :disabled="isWindowClickFlowDiagnosticRunning"
                          aria-label="窗口测试截止日期"
                        />
                      </label>
                    </div>
                  </div>
                  <div v-else class="settings-config-copy">
                    <strong>{{ action.label }}</strong>
                    <small>{{ action.description }}</small>
                    <span v-if="action.detail" class="diagnostic-action-detail">{{ action.detail }}</span>
                  </div>
                  <div
                    v-if="!action.showWindowClickFlowOptions"
                    :class="[
                      'settings-config-control',
                      'action-control',
                      'compact-control',
                      {
                        'scroll-step-action-control': action.showScrollStepInput,
                        'article-detail-option-action-control': action.showArticleDetailSkipCollectedOption,
                        'initial-content-storage-option-action-control': action.showInitialContentStorageOptions,
                        'article-detail-comments-option-action-control': action.showArticleDetailCommentsOptions || action.showArticleDetailOfflineCacheOptions,
                        'article-detail-offline-cache-option-action-control': action.showArticleDetailOfflineCacheOptions,
                      },
                    ]"
                  >
                    <label v-if="action.showScrollStepInput" class="window-diagnostic-scroll-step">
                      <span>滚动步长</span>
                      <AInputNumber
                        class="settings-ant-number window-diagnostic-scroll-step-input"
                        :value="windowDiagnosticScrollSteps"
                        :min="1"
                        :max="200"
                        :step="1"
                        :precision="0"
                        :controls="true"
                        :disabled="action.disabled?.() ?? false"
                        aria-label="滚动页面步长"
                        @change="handleWindowDiagnosticScrollStepsNumberChange"
                      />
                    </label>
                    <ACheckbox
                      v-if="action.showArticleDetailSkipCollectedOption"
                      v-model:checked="articleDetailSkipCollectedRecords"
                      class="article-detail-skip-option"
                      :disabled="action.disabled?.() ?? false"
                      aria-label="详情获取跳过已采集记录"
                    >
                      跳过已采集记录
                    </ACheckbox>
                    <ACheckbox
                      v-if="action.showInitialContentStorageOptions"
                      v-model:checked="initialContentStorageSkipCollectedRecords"
                      class="article-detail-skip-option"
                      :disabled="action.disabled?.() ?? false"
                      aria-label="初始内容存储跳过已采集记录"
                    >
                      跳过已采集记录
                    </ACheckbox>
                    <ACheckbox
                      v-if="action.showInitialContentStorageOptions"
                      v-model:checked="initialContentStorageStoreArticleDetail"
                      class="article-detail-skip-option"
                      disabled
                      aria-label="初始内容存储文章详情"
                    >
                      存储文章详情
                    </ACheckbox>
                    <ACheckbox
                      v-if="action.showArticleDetailCommentsOptions"
                      v-model:checked="articleDetailCommentsSkipCollectedRecords"
                      class="article-detail-skip-option"
                      :disabled="action.disabled?.() ?? false"
                      aria-label="详情评论跳过已采集记录"
                    >
                      跳过已采集记录
                    </ACheckbox>
                    <div
                      v-if="action.showArticleDetailCommentsOptions"
                      class="article-detail-comments-option-stack"
                    >
                      <ACheckbox
                        v-model:checked="articleDetailCommentsStoreArticleDetail"
                        class="article-detail-skip-option"
                        disabled
                        aria-label="详情评论存储文章详情"
                      >
                        存储文章详情
                      </ACheckbox>
                      <ACheckbox
                        v-model:checked="articleDetailCommentsStoreCommentInfo"
                        class="article-detail-skip-option"
                        disabled
                        aria-label="详情评论存储评论信息"
                      >
                        存储评论信息
                      </ACheckbox>
                    </div>
                    <div
                      v-if="action.showArticleDetailOfflineCacheOptions"
                      class="article-detail-offline-cache-option-stack"
                    >
                      <ACheckbox
                        v-model:checked="articleDetailOfflineCacheSkipCollectedRecords"
                        class="article-detail-skip-option"
                        :disabled="action.disabled?.() ?? false"
                        aria-label="离线缓存跳过已采集记录"
                      >
                        跳过已采集记录
                      </ACheckbox>
                      <ACheckbox
                        v-model:checked="articleDetailOfflineCacheStateful"
                        class="article-detail-skip-option"
                        :disabled="action.disabled?.() ?? false"
                        aria-label="离线缓存带状态"
                      >
                        带状态（beta）
                      </ACheckbox>
                    </div>
                    <div
                      v-if="action.showArticleDetailOfflineCacheOptions"
                      class="article-detail-comments-option-stack"
                    >
                      <ACheckbox
                        v-model:checked="articleDetailOfflineCacheStoreArticleDetail"
                        class="article-detail-skip-option"
                        disabled
                        aria-label="离线缓存存储文章详情"
                      >
                        存储文章详情
                      </ACheckbox>
                      <ACheckbox
                        v-model:checked="articleDetailOfflineCacheArchiveContent"
                        class="article-detail-skip-option"
                        disabled
                        aria-label="离线缓存归档内容"
                      >
                        离线归档内容
                      </ACheckbox>
                    </div>
                    <AButton
                      :class="['settings-ant-button', 'settings-row-button', 'diagnostic-action-button', action.tone ?? 'ghost']"
                      html-type="button"
                      :disabled="action.disabled?.() ?? false"
                      @click="action.run()"
                    >
                      <template #icon>
                        <AppIcon class="diagnostic-action-icon" :icon="action.icon" />
                      </template>
                      {{ action.buttonLabel ?? action.label }}
                    </AButton>
                  </div>
                </article>
              </div>
            </div>
          </section>

        </template>

        <section v-else class="settings-category-overview">
          <h3>设置说明</h3>
          <p>{{ selectedSettingsCategory.description }}</p>
          <ul class="note-list compact">
            <li class="soft-note">
              <AppIcon icon="fa-solid fa-circle" />
              <span>当前分类包含 {{ selectedCategoryFieldCount }} 个具体设置项，请在第二列选择后编辑。</span>
            </li>
            <li v-for="item in configGuideItems" :key="item" class="soft-note">
              <AppIcon icon="fa-solid fa-circle" />
              <span>{{ item }}</span>
            </li>
          </ul>
        </section>
      </section>
    </section>

    <section class="settings-bottom-panels" aria-label="运行环境和快速操作">
      <section class="env-panel page-panel" aria-label="运行环境">
        <header class="settings-panel-header compact">
          <div class="config-action-heading">
            <span class="config-action-heading-icon" aria-hidden="true">
              <AppIcon icon="fa-solid fa-desktop" />
            </span>
            <span class="config-action-heading-copy">
              <strong>运行环境</strong>
              <span>当前依赖与运行状态</span>
            </span>
          </div>
        </header>
        <div class="env-check-grid">
          <div v-for="item in settingsEnvironmentItems" :key="item.name" :class="['env-check', `env-check--${item.tone ?? 'blue'}`]">
            <AppIcon :icon="['env-check-icon', item.icon]" />
            <span class="env-check-copy">
              <strong>{{ item.name }}</strong>
              <small :class="['env-check-value', item.tone ?? 'blue']">{{ item.value }}</small>
            </span>
          </div>
        </div>
      </section>

      <section class="config-actions page-panel" aria-label="配置操作">
        <header class="settings-panel-header compact">
          <div class="config-action-heading">
            <span class="config-action-heading-icon" aria-hidden="true">
              <AppIcon icon="fa-solid fa-file-shield" />
            </span>
            <span class="config-action-heading-copy">
              <strong>配置操作</strong>
              <span>保存、恢复、清理、代理检测和启动自检</span>
            </span>
          </div>
        </header>
        <div class="config-action-grid">
          <AButton class="settings-ant-button config-action-button success" html-type="button" :loading="isSavingConfig" :disabled="isApplyingDefaults" @click="handleSaveConfig">
            <template #icon>
              <AppIcon class="config-action-inline-icon" icon="fa-regular fa-floppy-disk" />
            </template>
            {{ isSavingConfig ? '保存中' : '保存配置' }}
          </AButton>
          <AButton class="settings-ant-button config-action-button primary" html-type="button" :loading="isApplyingDefaults" :disabled="isSavingConfig || isRuntimeCacheBusy" @click="handleResetDefaults">
            <template #icon>
              <AppIcon class="config-action-inline-icon" icon="fa-solid fa-rotate-right" />
            </template>
            {{ isApplyingDefaults ? '恢复中' : '恢复默认' }}
          </AButton>
          <AButton class="settings-ant-button config-action-button orange" html-type="button" :loading="isClearingCache" :disabled="isRuntimeCacheBusy" @click="handleClearCache">
            <template #icon>
              <AppIcon class="config-action-inline-icon" icon="fa-solid fa-broom" />
            </template>
            {{ isClearingCache ? '清理中' : cacheCleaned ? '已清理' : '清理缓存' }}
          </AButton>
          <AButton class="settings-ant-button config-action-button ghost" html-type="button" :loading="isRunningStartupSelfCheck" @click="handleRunStartupSelfCheck">
            <template #icon>
              <AppIcon class="config-action-inline-icon" icon="fa-solid fa-list-check" />
            </template>
            {{ isRunningStartupSelfCheck ? '自检中' : '重新自检' }}
          </AButton>
        </div>
      </section>
    </section>

    <teleport to="body">
      <transition name="mitm-cert-dialog-fade">
        <div
          v-if="diagnosticResultDialogVisible"
          class="mitm-cert-dialog-backdrop diagnostic-result-dialog-backdrop"
          role="presentation"
          @click.self="handleDiagnosticResultBackdropClick"
        >
          <section
            class="diagnostic-result-dialog"
            :class="'diagnostic-result-dialog--' + diagnosticResultTone"
            role="dialog"
            aria-modal="true"
            aria-labelledby="diagnostic-result-dialog-title"
          >
            <header class="diagnostic-result-dialog-header">
              <span class="diagnostic-result-dialog-icon" aria-hidden="true">
                <AppIcon icon="fa-solid fa-list-check" />
              </span>
              <div>
                <h3 id="diagnostic-result-dialog-title">{{ diagnosticResultTitle }}</h3>
                <p>{{ diagnosticResultMessage || '诊断动作已返回结果。' }}</p>
              </div>
            </header>

            <dl
              v-if="diagnosticResultItems.length"
              ref="diagnosticResultListRef"
              class="diagnostic-result-list"
              @scroll="handleDiagnosticResultListScroll"
            >
              <div
                v-for="(item, itemIndex) in diagnosticResultItems"
                :key="item.label + '-' + itemIndex"
                :class="getDiagnosticResultItemClass(item)"
              >
                <template v-if="getDiagnosticResultMetaCells(item).length">
                  <div
                    v-for="cell in getDiagnosticResultCells(item)"
                    :key="cell.label"
                    class="diagnostic-result-cell"
                  >
                    <dt>{{ cell.label }}</dt>
                    <dd>{{ cell.value }}</dd>
                  </div>
                </template>
                <div v-else class="diagnostic-result-item-main">
                  <dt>{{ getDiagnosticResultPrimaryCell(item).label }}</dt>
                  <dd>{{ getDiagnosticResultPrimaryCell(item).value }}</dd>
                </div>
              </div>
            </dl>

            <footer class="mitm-cert-dialog-actions">
              <AButton
                v-if="isWindowClickFlowDiagnosticRunning"
                class="settings-ant-button config-action-button orange"
                html-type="button"
                @click="stopActiveWindowClickFlowDiagnostic"
              >
                立即停止
              </AButton>
              <AButton
                class="settings-ant-button config-action-button primary"
                html-type="button"
                @click="closeDiagnosticResultDialog"
              >
                知道了
              </AButton>
            </footer>
          </section>
        </div>
      </transition>
    </teleport>

    <AModal
      v-model:open="resetDefaultsDialogVisible"
      class="reset-defaults-modal"
      title="恢复系统默认配置"
      centered
      ok-text="确认恢复"
      cancel-text="取消"
      :confirm-loading="isApplyingDefaults"
      :closable="!isApplyingDefaults"
      :mask-closable="!isApplyingDefaults"
      :cancel-button-props="{ disabled: isApplyingDefaults }"
      @ok="confirmResetDefaults"
      @cancel="closeResetDefaultsDialog"
    >
      <div class="reset-defaults-modal-content">
        <div class="diagnostic-result-dialog-header">
          <span class="diagnostic-result-dialog-icon reset-defaults-dialog-icon" aria-hidden="true">
            <AppIcon icon="fa-solid fa-rotate-right" />
          </span>
          <div>
            <p>确认后将使用 src/config/system.yaml 覆盖 data/custom.yaml。</p>
          </div>
        </div>

        <div class="reset-defaults-dialog-notice">
          <strong>当前自定义配置会先备份</strong>
          <span>备份文件保存为 data/custom.yaml.bak，恢复完成后页面将立即同步系统默认值。</span>
        </div>
      </div>
    </AModal>

    <teleport to="body">
      <transition name="mitm-cert-dialog-fade">
        <div
          v-if="caCertificateDialogVisible"
          class="mitm-cert-dialog-backdrop"
          role="presentation"
          @click.self="closeMitmCertificateDialog"
        >
          <section
            class="mitm-cert-dialog"
            :class="`mitm-cert-dialog--${caCertificateDialogTone}`"
            role="dialog"
            aria-modal="true"
            aria-labelledby="mitm-cert-dialog-title"
          >
            <header class="mitm-cert-dialog-header">
              <span class="mitm-cert-dialog-icon" aria-hidden="true">
                <AppIcon :icon="caCertificateDialogIcon" />
              </span>
              <div class="mitm-cert-dialog-copy">
                <h3 id="mitm-cert-dialog-title">{{ caCertificateDialogTitle }}</h3>
                <p>{{ caCertificateDialogMessage || '正在准备 CA 证书操作...' }}</p>
              </div>
            </header>

            <section class="mitm-cert-source-card mitm-cert-source-card--project mitm-cert-project">
              <div class="mitm-cert-source-heading">
                <span class="mitm-cert-source-icon" aria-hidden="true">
                  <AppIcon icon="fa-regular fa-folder-open" />
                </span>
                <div class="mitm-cert-source-title">
                  <strong>项目内部证书文件</strong>
                  <span>来自当前项目配置的 CA 证书路径，用于检测、安装和系统匹配。</span>
                </div>
                <span :class="['mitm-cert-badge', caCertificateStatus.projectCertificateInstalled ? 'success' : 'warning']">
                  项目内 · {{ caCertificateStatus.projectCertificateInstalled ? '已安装到系统' : '未匹配系统证书' }}
                </span>
              </div>
              <dl class="mitm-cert-meta">
                <div>
                  <dt>项目路径</dt>
                  <dd>{{ caCertificateDialogProjectPath }}</dd>
                </div>
                <div>
                  <dt>指纹</dt>
                  <dd>{{ formatCaDialogValue(caCertificateDialogProjectThumbprint) }}</dd>
                </div>
                <div>
                  <dt>有效期</dt>
                  <dd>
                    {{ caCertificateDialogProjectCertificate?.notBefore || '未知' }}
                    至
                    {{ caCertificateDialogProjectCertificate?.notAfter || '未知' }}
                  </dd>
                </div>
                <div>
                  <dt>颁发者</dt>
                  <dd>{{ caCertificateDialogProjectCertificate?.issuer || '未知' }}</dd>
                </div>
              </dl>
            </section>

            <section class="mitm-cert-source-card mitm-cert-source-card--system mitm-cert-list-section">
              <div class="mitm-cert-source-heading">
                <span class="mitm-cert-source-icon" aria-hidden="true">
                  <AppIcon icon="fa-solid fa-desktop" />
                </span>
                <div class="mitm-cert-source-title">
                  <strong>Windows 当前用户根证书库</strong>
                  <span>来自操作系统证书存储 Cert:\CurrentUser\Root，清除操作只处理弹窗列出的系统证书。</span>
                </div>
                <span class="mitm-cert-badge info">
                  系统证书库 · {{ isCaCertificateDialogBusy ? '检索中' : `${mitmCertificateItems.length} 张` }}
                </span>
              </div>
              <div v-if="isCaCertificateDialogBusy" class="mitm-cert-empty mitm-cert-empty--inline">
                {{ caCertificateDialogMessage || '正在处理...' }}
              </div>
              <div v-else-if="mitmCertificateItems.length" class="mitm-cert-list">
                <article
                  v-for="certificate in mitmCertificateItems"
                  :key="`${certificate.storePath}-${certificate.thumbprint}`"
                  class="mitm-cert-item"
                >
                  <div class="mitm-cert-item-main">
                    <strong>{{ getCertificateDisplayName(certificate) }}</strong>
                    <span>{{ certificate.storePath }}</span>
                  </div>
                  <dl class="mitm-cert-meta">
                    <div>
                      <dt>指纹</dt>
                      <dd>{{ certificate.thumbprint }}</dd>
                    </div>
                    <div>
                      <dt>匹配</dt>
                      <dd>{{ certificate.matchesProject ? '与项目证书一致' : 'mitmproxy 相关证书' }}</dd>
                    </div>
                    <div>
                      <dt>有效期</dt>
                      <dd>{{ certificate.notBefore || '未知' }} 至 {{ certificate.notAfter || '未知' }}</dd>
                    </div>
                    <div>
                      <dt>颁发者</dt>
                      <dd>{{ certificate.issuer || '未知' }}</dd>
                    </div>
                  </dl>
                </article>
              </div>
              <div v-else class="mitm-cert-empty mitm-cert-empty--inline">
                当前没有检索到系统 mitmproxy 相关证书。
              </div>
            </section>

            <section
              v-if="caCertificateDialogDeletedItems.length || caCertificateDialogSkippedItems.length"
              class="mitm-cert-delete-result"
            >
              <div class="mitm-cert-section-title">
                <strong>删除结果</strong>
                <span>成功 {{ caCertificateDialogDeletedItems.length }} 张 / 失败 {{ caCertificateDialogSkippedItems.length }} 张</span>
              </div>
              <dl class="mitm-cert-meta">
                <div v-for="item in caCertificateDialogDeletedItems" :key="`deleted-${item.thumbprint}`">
                  <dt>已删除</dt>
                  <dd>{{ item.thumbprint }}</dd>
                </div>
                <div v-for="item in caCertificateDialogSkippedItems" :key="`skipped-${item.thumbprint}`">
                  <dt>失败</dt>
                  <dd>{{ item.thumbprint }}：{{ item.reason }}</dd>
                </div>
              </dl>
            </section>

            <footer class="mitm-cert-dialog-actions">
              <AButton
                class="settings-ant-button config-action-button ghost"
                html-type="button"
                :disabled="isCaCertificateDialogBusy"
                @click="closeMitmCertificateDialog"
              >
                {{ caCertificateDialogCloseText }}
              </AButton>
              <AButton
                v-if="caCertificateDialogCanConfirmInstall"
                class="settings-ant-button config-action-button primary"
                html-type="button"
                :disabled="isInstallingCaCertificate"
                @click="confirmInstallCaCertificate"
              >
                <template #icon>
                  <AppIcon icon="fa-solid fa-certificate" />
                </template>
                确认安装
              </AButton>
              <AButton
                v-if="caCertificateDialogCanConfirmDelete"
                class="settings-ant-button config-action-button danger"
                html-type="button"
                :disabled="isDeletingMitmCertificates"
                @click="handleConfirmDeleteMitmCertificates"
              >
                <template #icon>
                  <AppIcon icon="fa-solid fa-trash-can" />
                </template>
                确认删除
              </AButton>
            </footer>
          </section>
        </div>
      </transition>
    </teleport>
  </section>
</template>

<style scoped>
.settings-page {
  --settings-brand: #0072EF;
  --settings-brand-hover: #0066D8;
  --settings-brand-active: #0057BD;
  --settings-brand-soft: #DCEBFF;
  --settings-brand-tint: #F1F7FF;
  --settings-brand-border: #B9D6FA;
  --settings-action: #4AAE9F;
  --settings-action-hover: #409D91;
  --settings-action-active: #368C82;
  --settings-action-soft: #E7F5F2;
  --settings-action-tint: #F4FBFA;
  --settings-action-ink: #2F6F66;
  --settings-action-border: rgba(74, 174, 159, 0.36);
  --settings-ink: #24364B;
  --settings-muted: #5D6F86;
  --settings-line: rgba(135, 161, 191, 0.24);
  --settings-surface: #FFFFFF;
  grid-template-columns: repeat(6, minmax(0, 1fr));
  grid-template-rows: minmax(0, 1fr) 216px;
  grid-template-areas:
    'center center center center center center'
    'bottom bottom bottom bottom bottom bottom';
}

.settings-three-pane {
  grid-area: center;
  display: grid;
  grid-template-columns: 230px 300px minmax(0, 1fr);
  gap: 16px;
  min-height: 0;
}

.settings-bottom-panels {
  grid-area: bottom;
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 16px;
  min-width: 0;
  min-height: 0;
}

.settings-primary-nav,
.settings-secondary-nav,
.settings-detail-pane {
  min-width: 0;
  min-height: 0;
  border: 1px solid var(--settings-line);
  border-radius: 8px;
  background: var(--settings-surface);
  box-shadow: var(--paper-shadow-sm), var(--paper-shadow-md);
}

.settings-primary-nav,
.settings-secondary-nav {
  display: grid;
  align-content: start;
  gap: 8px;
  padding: 14px;
  overflow: auto;
}

.settings-nav-header {
  display: grid;
  gap: 5px;
  padding: 2px 4px 8px;
  border-bottom: 1px solid rgba(104, 141, 181, 0.18);
}

.settings-nav-header strong {
  color: var(--settings-ink);
  font-size: 15px;
  font-weight: 600;
}

.settings-nav-header span {
  color: var(--settings-muted);
  font-size: 12px;
  font-weight: 500;
  line-height: 1.35;
}

.settings-primary-item,
.settings-secondary-item {
  --settings-menu-ink: #15386F;
  --settings-menu-active-ink: #15386F;
  --settings-menu-active-bg: rgba(21, 56, 111, 0.1);
  --settings-menu-hover-bg: rgba(21, 56, 111, 0.06);
  --settings-menu-active-border: rgba(21, 56, 111, 0.28);
  display: grid;
  grid-template-columns: 28px minmax(0, 1fr);
  align-items: center;
  gap: 8px;
  width: 100%;
  min-height: 42px;
  padding: 8px 9px;
  border: 1px solid transparent;
  border-radius: 8px;
  color: var(--settings-menu-ink);
  background: transparent;
  text-align: left;
  cursor: pointer;
}

.settings-secondary-item {
  min-height: 46px;
}

.settings-primary-item:hover,
.settings-secondary-item:hover {
  border-color: rgba(21, 56, 111, 0.18);
  background: var(--settings-menu-hover-bg);
}

.settings-primary-item.active,
.settings-secondary-item.active {
  border-color: var(--settings-menu-active-border);
  background: var(--settings-menu-active-bg);
  color: var(--settings-menu-active-ink);
}

.settings-primary-icon,
.settings-secondary-icon {
  display: grid;
  place-items: center;
  width: 24px;
  height: 24px;
  border-radius: 6px;
  color: var(--settings-menu-ink);
  background: rgba(21, 56, 111, 0.08);
  font-size: 12px;
}

.settings-primary-item.active .settings-primary-icon,
.settings-secondary-item.active .settings-secondary-icon {
  color: var(--settings-menu-active-ink);
  background: rgba(21, 56, 111, 0.12);
}

.settings-primary-item span,
.settings-secondary-item span {
  display: grid;
  min-width: 0;
}

.settings-primary-item strong,
.settings-secondary-item strong {
  overflow: hidden;
  color: currentColor;
  font-size: 14px;
  font-weight: 600;
  line-height: 1.16;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.settings-detail-pane {
  position: relative;
  display: grid;
  grid-template-rows: auto minmax(0, 1fr);
  gap: 14px;
  padding: 18px 22px;
  overflow: hidden;
}

.settings-detail-header {
  min-width: 0;
}

.settings-detail-title-row {
  display: flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
}

.settings-detail-icon {
  display: grid;
  flex: 0 0 auto;
  place-items: center;
  width: 32px;
  height: 32px;
  border: 1px solid rgba(21, 56, 111, 0.18);
  border-radius: 8px;
  color: #15386F;
  background: rgba(21, 56, 111, 0.1);
  font-size: 14px;
}

.settings-detail-header h2,
.settings-detail-header small {
  margin: 0;
}

.settings-detail-header h2 {
  min-width: 0;
  color: var(--settings-ink);
  font-size: 21px;
  font-weight: 700;
  line-height: 1.2;
}

.settings-detail-header small {
  display: block;
  margin-top: 6px;
  color: var(--settings-muted);
  font-size: 12.5px;
  font-weight: 500;
  line-height: 1.45;
}

.settings-detail-section,
.settings-category-overview {
  min-height: 0;
  overflow: auto;
}

.detail-form {
  display: grid;
  gap: 12px;
}

.settings-detail-content {
  display: grid;
  align-content: start;
  gap: 12px;
  min-width: 0;
}

.settings-config-list {
  display: grid;
  grid-template-columns: minmax(0, 1fr);
  gap: 10px;
  min-width: 0;
}

.settings-config-row {
  display: grid;
  grid-template-columns: minmax(0, 1fr);
  gap: 10px;
  min-width: 0;
  padding: 13px 14px;
  border: 1px solid rgba(135, 161, 191, 0.2);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.82);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.68);
}

.settings-config-row.compact-control {
  grid-template-columns: minmax(0, 1fr) max-content;
  align-items: start;
  column-gap: 16px;
  row-gap: 5px;
}

.settings-config-row.wide-control {
  grid-template-columns: minmax(0, 1fr);
  align-items: start;
}

.settings-config-row.readonly {
  border-color: rgba(0, 114, 239, 0.18);
  background: rgba(241, 247, 255, 0.72);
}

.settings-config-copy {
  display: grid;
  gap: 5px;
  min-width: 0;
}

.settings-config-row.compact-control .settings-config-copy {
  grid-column: 1;
  grid-row: 1;
}

.settings-config-row.compact-control .settings-config-keyline {
  grid-column: 1 / -1;
  grid-row: 2;
}

.settings-config-copy strong {
  color: var(--settings-ink);
  font-size: 15px;
  font-weight: 600;
  line-height: 1.25;
}

.settings-config-copy small {
  color: var(--settings-muted);
  font-size: 12.5px;
  font-weight: 500;
  line-height: 1.45;
}

.diagnostic-action-detail {
  display: block;
  width: fit-content;
  max-width: 100%;
  margin-top: 2px;
  color: var(--settings-brand);
  font-family: Consolas, 'SFMono-Regular', monospace;
  font-size: 12.5px;
  font-weight: 600;
  line-height: 1.35;
  overflow-wrap: anywhere;
}

.settings-config-keyline {
  display: block;
  width: fit-content;
  max-width: 100%;
  margin-top: 2px;
  color: var(--settings-brand);
  font-family: Consolas, 'SFMono-Regular', monospace;
  font-size: 12.5px;
  font-weight: 600;
  line-height: 1.35;
  overflow-wrap: break-word;
  word-break: normal;
  white-space: normal;
}

.settings-config-control {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto auto;
  align-items: center;
  gap: 8px;
  min-width: 0;
}

.settings-config-control.compact-control {
  grid-template-columns: max-content;
  align-self: start;
  grid-column: 2;
  grid-row: 1;
  justify-content: end;
  justify-self: end;
  padding-top: 5px;
  width: auto;
}

.settings-config-control.wide-control {
  grid-template-columns: minmax(0, 1fr) auto auto;
  justify-self: stretch;
  width: 100%;
}

.settings-config-control.action-control {
  grid-template-columns: max-content;
  justify-content: end;
}

.settings-inline-input {
  box-sizing: border-box;
  width: 100%;
  min-width: 0;
  height: 38px;
  padding: 0 12px;
  border: 1px solid var(--line);
  border-radius: 6px;
  color: var(--settings-ink);
  background: rgba(255, 255, 255, 0.86);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.68);
  font: inherit;
  font-size: 14px;
  font-weight: 600;
}

.settings-inline-input[readonly] {
  color: var(--settings-muted);
  background: rgba(241, 247, 255, 0.78);
  caret-color: transparent;
  cursor: text;
  user-select: text;
}

.settings-range-readonly {
  display: grid;
  grid-template-columns: minmax(86px, 1fr) auto minmax(86px, 1fr) auto;
  align-items: end;
  gap: 8px;
  width: min(360px, 100%);
}

.settings-range-readonly label {
  display: grid;
  gap: 4px;
  min-width: 0;
}

.settings-range-readonly label > span,
.settings-range-unit,
.settings-range-separator {
  color: var(--settings-muted);
  font-size: 12px;
  font-weight: 600;
  line-height: 1.2;
  white-space: nowrap;
}

.settings-range-readonly .settings-inline-input {
  width: 100%;
  text-align: center;
}

.settings-range-separator,
.settings-range-unit {
  align-self: center;
  padding-top: 18px;
}

.settings-config-control.compact-control .settings-inline-input {
  width: 156px;
  max-width: 100%;
}

.settings-row-button {
  min-width: 74px;
  height: 38px !important;
  padding-inline: 14px !important;
  white-space: nowrap;
}

@media (max-width: 760px) {
  .settings-config-row.compact-control {
    grid-template-columns: minmax(0, 1fr);
  }

  .settings-config-row.compact-control .settings-config-copy {
    grid-column: 1;
    grid-row: 1;
  }

  .settings-config-control.compact-control {
    grid-column: 1;
    grid-row: 2;
    justify-self: stretch;
    justify-content: start;
  }

  .settings-config-row.compact-control .settings-config-keyline {
    grid-column: 1;
    grid-row: 3;
  }
}

.config-field-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 10px;
  min-width: 0;
}

.config-field-card {
  display: grid;
  align-content: start;
  gap: 7px;
  min-width: 0;
  min-height: 108px;
  padding: 12px 13px;
  border: 1px solid rgba(135, 161, 191, 0.2);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.82);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.62);
}

.config-field-card.readonly {
  border-color: rgba(0, 114, 239, 0.18);
  background: rgba(241, 247, 255, 0.72);
}

.config-field-card.success {
  border-color: rgba(0, 114, 239, 0.18);
  background: rgba(241, 247, 255, 0.72);
}

.config-field-card.warning {
  border-color: rgba(217, 151, 39, 0.2);
  background: rgba(255, 248, 234, 0.62);
}

.config-field-key {
  overflow: hidden;
  color: var(--settings-brand);
  font-family: Consolas, 'SFMono-Regular', monospace;
  font-size: 11.5px;
  font-weight: 600;
  line-height: 1.2;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.config-field-card strong {
  min-width: 0;
  color: var(--settings-ink);
  font-size: 15px;
  font-weight: 600;
  line-height: 1.25;
  overflow-wrap: anywhere;
}

.config-field-card small {
  color: var(--settings-muted);
  font-size: 12px;
  font-weight: 500;
  line-height: 1.45;
}

.control-switch-line {
  display: inline-flex;
  align-items: center;
  justify-self: start;
  gap: 10px;
  min-height: 38px;
  color: var(--settings-ink);
  font-size: 13px;
  font-weight: 600;
}

.diagnostic-action-grid {
  display: grid;
  grid-template-columns: minmax(0, 1fr);
  gap: 10px;
  min-width: 0;
}

.diagnostic-action-grid.settings-config-list {
  gap: 12px;
}

.diagnostic-action-grid .settings-config-row {
  --diagnostic-row-accent: #357FD9;
  position: relative;
  overflow: hidden;
  padding: 14px 14px 14px 20px;
  border-color: rgba(104, 141, 181, 0.24);
  background:
    linear-gradient(180deg, rgba(255, 255, 255, 0.92), rgba(248, 251, 255, 0.86)),
    rgba(255, 255, 255, 0.86);
  box-shadow:
    0 4px 10px rgba(21, 56, 111, 0.06),
    inset 0 1px 0 rgba(255, 255, 255, 0.74);
  transition:
    border-color 160ms ease,
    box-shadow 160ms ease,
    background-color 160ms ease;
}

.diagnostic-action-grid .settings-config-row::before {
  content: '';
  position: absolute;
  top: 12px;
  bottom: 12px;
  left: 10px;
  width: 4px;
  border-radius: 999px;
  background: var(--diagnostic-row-accent);
  opacity: 0.9;
}

.diagnostic-action-grid .settings-config-row:hover {
  border-color: rgba(53, 127, 217, 0.32);
  box-shadow:
    0 6px 14px rgba(21, 56, 111, 0.09),
    inset 0 1px 0 rgba(255, 255, 255, 0.82);
}

.diagnostic-action-grid .settings-config-row.primary,
.diagnostic-action-grid .settings-config-row.blue {
  --diagnostic-row-accent: #357FD9;
}

.diagnostic-action-grid .settings-config-row.success {
  --diagnostic-row-accent: #35B889;
}

.diagnostic-action-grid .settings-config-row.orange {
  --diagnostic-row-accent: #F28B3C;
}

.diagnostic-action-grid .settings-config-row.teal {
  --diagnostic-row-accent: #159A91;
}

.diagnostic-action-grid .settings-config-row.purple {
  --diagnostic-row-accent: #8968CD;
}

.diagnostic-action-grid .settings-config-row.danger {
  --diagnostic-row-accent: #D74D4D;
}

.settings-config-control.compact-control.scroll-step-action-control {
  grid-template-columns: max-content max-content;
  gap: 12px;
}

.settings-config-control.compact-control.article-detail-option-action-control {
  grid-template-columns: max-content max-content;
  gap: 12px;
}

.settings-config-control.compact-control.initial-content-storage-option-action-control {
  grid-template-columns: max-content max-content max-content;
  gap: 12px;
}

.settings-config-control.compact-control.article-detail-comments-option-action-control {
  grid-template-columns: max-content max-content max-content;
  gap: 12px;
}

.settings-config-control.compact-control.article-detail-offline-cache-option-action-control {
  grid-template-columns: max-content max-content max-content;
  align-items: center;
  gap: 18px;
}

.article-detail-comments-option-stack {
  display: grid;
  gap: 4px;
}

.article-detail-offline-cache-option-stack {
  display: grid;
  gap: 4px;
}

.article-detail-skip-option {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  min-height: 34px;
  color: var(--settings-muted);
  font-size: 12px;
  font-weight: 600;
  white-space: nowrap;
  user-select: none;
}

.window-diagnostic-scroll-step {
  display: grid;
  grid-template-columns: auto 76px;
  align-items: center;
  gap: 8px;
  color: var(--settings-muted);
  font-size: 12px;
  font-weight: 600;
  white-space: nowrap;
}

.window-diagnostic-scroll-step-input {
  width: 76px;
}

.settings-config-row.window-click-flow-row {
  grid-template-columns: minmax(0, 1fr);
}

.window-click-flow-layout {
  display: grid;
  grid-template-rows: auto auto auto;
  gap: 8px;
  min-width: 0;
}

.window-click-flow-title {
  min-width: 0;
}

.window-click-flow-title strong {
  color: var(--settings-ink);
  font-size: 15px;
  font-weight: 600;
  line-height: 1.25;
}

.window-click-flow-description-line {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  min-width: 0;
  grid-row: 3;
}

.window-click-flow-description {
  display: grid;
  gap: 4px;
  min-width: 0;
}

.window-click-flow-description small {
  color: var(--settings-muted);
  font-size: 12.5px;
  font-weight: 500;
  line-height: 1.45;
}

.window-click-flow-fields {
  display: flex;
  flex-wrap: nowrap;
  align-items: end;
  gap: 6px;
  min-width: 0;
  overflow-x: auto;
  grid-row: 2;
}

.window-click-flow-field {
  display: grid;
  align-content: end;
  gap: 4px;
  min-width: 0;
}

.window-click-flow-field > span {
  color: var(--settings-muted);
  font-size: 11px;
  font-weight: 600;
  white-space: nowrap;
}

.window-click-flow-stepper {
  min-width: 116px;
  width: 116px;
  height: 38px;
  flex: 0 0 116px;
}

.window-click-flow-select {
  width: 100%;
  min-width: 0;
  flex: 0 0 116px;
}

.window-click-flow-field--mode {
  width: 116px;
  flex: 0 0 116px;
}

.window-click-flow-field--date {
  width: 142px;
  flex: 0 0 142px;
}

.window-click-flow-field--range {
  width: 248px;
  flex: 0 0 248px;
}

.window-click-flow-date-picker {
  width: 100%;
  height: 38px;
}

.window-click-flow-date-range-picker {
  width: 100%;
  height: 38px;
}

.window-click-flow-description-line > .settings-ant-button {
  flex: 0 0 168px;
}

@media (max-width: 1100px) {
  .window-click-flow-fields {
    justify-content: flex-start;
    overflow-x: auto;
  }

  .window-click-flow-description-line {
    align-items: flex-start;
  }
}

.diagnostic-action-grid :deep(.diagnostic-action-button) {
  --diagnostic-action-color: #2F6F66;
  --diagnostic-action-border: rgba(74, 174, 159, 0.3);
  --diagnostic-action-bg: #F4FBFA;
  --diagnostic-action-hover-color: #265F58;
  --diagnostic-action-hover-border: rgba(74, 174, 159, 0.5);
  --diagnostic-action-hover-bg: #E7F5F2;
  --diagnostic-action-icon-color: currentColor;
  --diagnostic-action-icon-bg: rgba(74, 174, 159, 0.14);

  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 168px;
  min-width: 168px;
  min-height: 52px;
  height: auto !important;
  padding: 7px 12px !important;
  color: var(--diagnostic-action-color);
  border-color: var(--diagnostic-action-border);
  border-radius: 12px;
  background: var(--diagnostic-action-bg);
  box-shadow: none;
  text-align: center;
  transition:
    color 0.16s ease,
    border-color 0.16s ease,
    background-color 0.16s ease;
}

.diagnostic-action-grid :deep(.diagnostic-action-button.ghost),
.diagnostic-action-grid :deep(.diagnostic-action-button) {
  color: var(--diagnostic-action-color);
  border-color: var(--diagnostic-action-border);
  background: var(--diagnostic-action-bg);
}

.diagnostic-action-grid :deep(.diagnostic-action-button:hover) {
  color: var(--diagnostic-action-hover-color);
  border-color: var(--diagnostic-action-hover-border);
  background: var(--diagnostic-action-hover-bg);
  box-shadow: none;
  transform: none;
}

.diagnostic-action-grid :deep(.diagnostic-action-button.success) {
  --diagnostic-action-color: #ffffff;
  --diagnostic-action-border: #35B889;
  --diagnostic-action-bg: #35B889;
  --diagnostic-action-hover-color: #ffffff;
  --diagnostic-action-hover-border: #2EA579;
  --diagnostic-action-hover-bg: #2EA579;
  --diagnostic-action-icon-color: #ffffff;
  --diagnostic-action-icon-bg: rgba(255, 255, 255, 0.22);
  color: #ffffff;
  border-color: #35B889;
  background: #35B889;
}

.diagnostic-action-grid :deep(.diagnostic-action-button.orange) {
  --diagnostic-action-color: #946012;
  --diagnostic-action-border: rgba(215, 154, 48, 0.3);
  --diagnostic-action-bg: #FFF8EA;
  --diagnostic-action-hover-color: #7A4E0E;
  --diagnostic-action-hover-border: rgba(215, 154, 48, 0.46);
  --diagnostic-action-hover-bg: #FFF1D6;
  --diagnostic-action-icon-color: #946012;
  --diagnostic-action-icon-bg: rgba(215, 154, 48, 0.16);
  color: #946012;
  border-color: rgba(215, 154, 48, 0.3);
  background: #FFF8EA;
}

.diagnostic-action-grid :deep(.diagnostic-action-button.primary) {
  --diagnostic-action-color: #ffffff;
  --diagnostic-action-border: #357FD9;
  --diagnostic-action-bg: #357FD9;
  --diagnostic-action-hover-color: #ffffff;
  --diagnostic-action-hover-border: #2267B8;
  --diagnostic-action-hover-bg: #2267B8;
  --diagnostic-action-icon-color: #ffffff;
  --diagnostic-action-icon-bg: rgba(255, 255, 255, 0.22);
  color: #ffffff;
  border-color: #357FD9;
  background: #357FD9;
}

.diagnostic-action-grid :deep(.diagnostic-action-button.blue) {
  --diagnostic-action-color: #235A93;
  --diagnostic-action-border: rgba(73, 139, 217, 0.32);
  --diagnostic-action-bg: #EEF6FF;
  --diagnostic-action-hover-color: #174B82;
  --diagnostic-action-hover-border: rgba(73, 139, 217, 0.48);
  --diagnostic-action-hover-bg: #DFEDFF;
  --diagnostic-action-icon-color: #235A93;
  --diagnostic-action-icon-bg: rgba(73, 139, 217, 0.16);
  color: #235A93;
  border-color: rgba(73, 139, 217, 0.32);
  background: #EEF6FF;
}

.diagnostic-action-grid :deep(.diagnostic-action-button.teal) {
  --diagnostic-action-color: #0F766E;
  --diagnostic-action-border: rgba(21, 154, 145, 0.3);
  --diagnostic-action-bg: #ECFDFB;
  --diagnostic-action-hover-color: #0B615B;
  --diagnostic-action-hover-border: rgba(21, 154, 145, 0.46);
  --diagnostic-action-hover-bg: #D7F7F3;
  --diagnostic-action-icon-color: #0F766E;
  --diagnostic-action-icon-bg: rgba(21, 154, 145, 0.16);
  color: #0F766E;
  border-color: rgba(21, 154, 145, 0.3);
  background: #ECFDFB;
}

.diagnostic-action-grid :deep(.diagnostic-action-button.purple) {
  --diagnostic-action-color: #6944A5;
  --diagnostic-action-border: rgba(137, 104, 205, 0.32);
  --diagnostic-action-bg: #F5F0FF;
  --diagnostic-action-hover-color: #553591;
  --diagnostic-action-hover-border: rgba(137, 104, 205, 0.48);
  --diagnostic-action-hover-bg: #ECE3FF;
  --diagnostic-action-icon-color: #6944A5;
  --diagnostic-action-icon-bg: rgba(137, 104, 205, 0.16);
  color: #6944A5;
  border-color: rgba(137, 104, 205, 0.32);
  background: #F5F0FF;
}

.diagnostic-action-grid :deep(.diagnostic-action-button.danger) {
  --diagnostic-action-color: #B4232E;
  --diagnostic-action-border: rgba(224, 76, 86, 0.28);
  --diagnostic-action-bg: #FFF1F1;
  --diagnostic-action-hover-color: #9F1D28;
  --diagnostic-action-hover-border: rgba(224, 76, 86, 0.44);
  --diagnostic-action-hover-bg: #FFE2E2;
  --diagnostic-action-icon-color: #B4232E;
  --diagnostic-action-icon-bg: rgba(224, 76, 86, 0.14);
  color: #B4232E;
  border-color: rgba(224, 76, 86, 0.28);
  background: #FFF1F1;
}

.diagnostic-action-grid :deep(.diagnostic-action-button .ant-btn-icon) {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 22px;
  margin-inline-end: 8px;
  line-height: 1;
}

.diagnostic-action-icon {
  display: grid;
  place-items: center;
  width: 22px;
  height: 22px;
  border-radius: 8px;
  color: var(--diagnostic-action-icon-color);
  background: var(--diagnostic-action-icon-bg);
  font-size: 12px;
}

.diagnostic-action-copy {
  display: grid;
  gap: 4px;
  min-width: 0;
}

.diagnostic-action-copy strong,
.diagnostic-action-copy small {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.diagnostic-action-copy strong {
  font-size: 14px;
  font-weight: 600;
}

.diagnostic-action-copy small {
  font-size: 11.5px;
  font-weight: 500;
  opacity: 0.82;
}

.two-column-detail {
  grid-template-columns: repeat(2, minmax(0, 1fr));
  align-items: start;
}

.detail-row {
  grid-template-columns: 150px minmax(0, 1fr);
  align-items: center;
  gap: 12px;
  min-height: 44px;
  padding: 10px 12px;
  border: 1px solid rgba(104, 141, 181, 0.18);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.44);
}

.two-column-detail .detail-row {
  grid-template-columns: minmax(0, 1fr);
  align-items: start;
}

.settings-detail-pane .settings-label small {
  display: block;
  margin-top: 4px;
}

.detail-note {
  margin: 0;
  padding: 10px 12px;
  border: 1px solid rgba(45, 117, 214, 0.14);
  border-radius: 8px;
  color: var(--ink-muted);
  background: rgba(222, 238, 250, 0.34);
  font-size: 12.5px;
  font-weight: 500;
  line-height: 1.55;
}

.settings-explain-card {
  display: grid;
  gap: 10px;
  padding: 16px 18px;
  border: 1px solid rgba(104, 141, 181, 0.2);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.52);
}

.settings-explain-card.warning {
  border-color: rgba(223, 122, 53, 0.25);
  background: rgba(250, 235, 218, 0.34);
}

.settings-explain-card h3,
.settings-category-overview h3,
.settings-explain-card p,
.settings-category-overview p {
  margin: 0;
}

.settings-explain-card h3,
.settings-category-overview h3 {
  color: var(--settings-ink);
  font-size: 17px;
  font-weight: 600;
}

.settings-explain-card p,
.settings-category-overview p {
  color: var(--settings-muted);
  font-size: 13px;
  font-weight: 500;
  line-height: 1.62;
}

.detail-bullet-list {
  display: grid;
  gap: 8px;
  margin: 0;
  padding-left: 18px;
  color: var(--settings-muted);
  font-size: 12.8px;
  font-weight: 500;
  line-height: 1.45;
}

.settings-category-overview {
  display: grid;
  align-content: start;
  gap: 12px;
}

.detail-wide-action {
  justify-self: start;
  min-width: 180px;
}

.config-note {
  grid-area: note;
}

.env-panel,
.config-actions,
.config-note {
  border-color: rgba(104, 141, 181, 0.24);
  background:
    linear-gradient(180deg, rgba(255, 255, 255, 0.74), rgba(242, 248, 252, 0.62));
  padding: 17px 22px;
}

.env-panel,
.config-actions {
  display: grid;
  align-content: start;
  min-width: 0;
  min-height: 0;
  overflow: hidden;
}

.env-panel {
  padding-inline: 16px;
}

.settings-form {
  gap: 12px;
  margin-top: 16px;
}

.settings-panel-header {
  display: grid;
  gap: 0;
  min-width: 0;
}

.settings-panel-header strong,
.settings-panel-header span {
  display: block;
  margin: 0;
}

.settings-panel-header strong {
  overflow: hidden;
  color: var(--settings-ink);
  font-size: 15px;
  font-weight: 600;
  line-height: 1.2;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.settings-panel-header span {
  margin-top: 4px;
  overflow: hidden;
  color: var(--settings-muted);
  font-size: 12px;
  font-weight: 500;
  line-height: 1.25;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.settings-panel-header.compact {
  grid-template-columns: minmax(0, 1fr) auto;
  align-items: center;
}

.settings-ant-control {
  width: 100%;
  min-width: 0;
}

.compact-number-input {
  width: 154px;
  justify-self: start;
}

.settings-page :deep(.settings-ant-control.ant-input),
.settings-page :deep(.settings-ant-control.ant-select),
.settings-page :deep(.settings-ant-control.ant-picker),
.settings-page :deep(.settings-ant-number.ant-input-number) {
  width: 100%;
  height: 38px;
  color: var(--settings-ink);
  font-size: 14px;
  font-weight: 600;
}

.settings-page :deep(.settings-ant-control.ant-input),
.settings-page :deep(.settings-ant-control.ant-select .ant-select-selector),
.settings-page :deep(.settings-ant-control.ant-picker),
.settings-page :deep(.settings-ant-number.ant-input-number) {
  border-color: var(--settings-line);
  border-radius: 8px;
  background: #ffffff;
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.6);
}

.settings-page :deep(.settings-ant-control.ant-input) {
  padding: 0 12px;
  line-height: 36px;
}

.settings-page :deep(.settings-ant-control.ant-input[readonly]) {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  cursor: text;
}

.settings-page :deep(.settings-ant-control.ant-select .ant-select-selector) {
  align-items: center;
  height: 38px !important;
  padding: 0 12px;
}

.settings-page :deep(.settings-ant-control.ant-select .ant-select-selection-item),
.settings-page :deep(.settings-ant-control.ant-select .ant-select-selection-placeholder),
.settings-page :deep(.settings-ant-control.ant-picker input),
.settings-page :deep(.settings-ant-control.ant-picker .ant-picker-input > input),
.settings-page :deep(.settings-ant-control.ant-picker .ant-picker-separator),
.settings-page :deep(.settings-ant-control.ant-picker .ant-picker-suffix),
.settings-page :deep(.settings-ant-control.ant-picker .ant-picker-clear) {
  color: var(--settings-ink);
  font-size: 14px;
  font-weight: 600;
}

.settings-page :deep(.settings-ant-control.ant-picker input::placeholder),
.settings-page :deep(.settings-ant-control.ant-select .ant-select-selection-placeholder) {
  color: rgba(21, 56, 111, 0.55);
}

.settings-page :deep(.window-click-flow-date-range-picker.ant-picker .ant-picker-input > input) {
  text-align: center;
}

:global(.settings-select-panel),
:global(.window-click-flow-date-picker-panel .ant-picker-panel-container),
:global(.window-click-flow-date-range-picker-panel .ant-picker-panel-container) {
  color: var(--ink);
  font-family: 'Microsoft YaHei', 'PingFang SC', 'Segoe UI', system-ui, sans-serif;
  font-size: 14px;
}

.settings-config-control.compact-control :deep(.settings-ant-control.ant-input),
.settings-config-control.compact-control :deep(.settings-ant-control.ant-select) {
  width: min(180px, 100%);
}

.proxy-number-input,
.settings-field-stepper,
.window-click-flow-stepper {
  width: 154px;
  justify-self: end;
}

.settings-page :deep(.settings-ant-number.ant-input-number) {
  min-width: 148px;
  max-width: 244px;
  border-color: var(--settings-line);
  border-radius: 8px;
  background: #ffffff;
  box-shadow:
    inset 0 1px 0 rgba(255, 255, 255, 0.84),
    0 2px 6px rgba(0, 114, 239, 0.06);
}

.settings-page :deep(.settings-ant-number .ant-input-number-input-wrap) {
  height: 100%;
}

.settings-page :deep(.settings-ant-number .ant-input-number-input) {
  height: 36px;
  color: var(--settings-ink);
  font-size: 14px;
  font-weight: 600;
  text-align: center;
}

.settings-page :deep(.settings-ant-number .ant-input-number-handler-wrap) {
  opacity: 1;
  border-start-end-radius: 8px;
  border-end-end-radius: 8px;
}

.settings-page :deep(.settings-ant-number .ant-input-number-handler) {
  color: var(--settings-muted);
}

.settings-page :deep(.settings-ant-number .ant-input-number-handler:hover) {
  color: var(--settings-action-ink);
}

.settings-ant-switch {
  flex-shrink: 0;
}

.settings-page :deep(.settings-ant-switch.ant-switch-checked) {
  background: var(--settings-action);
}

.settings-page :deep(.article-detail-skip-option.ant-checkbox-wrapper) {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  min-height: 30px;
  margin-inline-end: 0;
  color: var(--settings-ink);
  font-size: 13px;
  font-weight: 500;
}

.settings-page :deep(.article-detail-skip-option.ant-checkbox-wrapper .ant-checkbox + span) {
  padding-inline-start: 0;
  padding-inline-end: 0;
}

.settings-ant-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  min-width: 76px;
  height: 38px !important;
  padding: 0 12px !important;
  border: 1px solid var(--settings-action-border);
  border-radius: 8px;
  color: var(--settings-action-ink);
  background: var(--settings-action-tint);
  box-shadow: none;
  font-size: 14px;
  font-weight: 600;
  line-height: 1;
  transition:
    color 160ms ease,
    border-color 160ms ease,
    background-color 160ms ease;
}

.settings-page :deep(.settings-ant-button .ant-btn-icon),
.settings-page :deep(.settings-ant-button .ant-btn-icon > *) {
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.settings-ant-button.primary {
  color: #ffffff;
  border-color: var(--settings-action);
  background: var(--settings-action);
}

.settings-ant-button.success {
  color: #ffffff;
  border-color: var(--settings-action);
  background: var(--settings-action);
}

.settings-ant-button.orange {
  color: #946012;
  border-color: rgba(215, 154, 48, 0.36);
  background: #FFF8EA;
}

.settings-ant-button.danger {
  color: #ffffff;
  border-color: rgba(194, 59, 55, 0.72);
  background: #C9424D;
}

.settings-ant-button.ghost {
  color: var(--settings-action-ink);
  background: var(--settings-action-tint);
}

.settings-ant-button:not(:disabled):hover {
  color: var(--settings-action-ink);
  border-color: rgba(74, 174, 159, 0.5);
  background: var(--settings-action-soft);
  box-shadow: none;
}

.settings-ant-button.primary:not(:disabled):hover {
  color: #ffffff;
  border-color: var(--settings-action-hover);
  background: var(--settings-action-hover);
}

.settings-ant-button.success:not(:disabled):hover {
  color: #ffffff;
  border-color: var(--settings-action-hover);
  background: var(--settings-action-hover);
}

.settings-ant-button.orange:not(:disabled):hover {
  color: #7A4D0D;
  border-color: rgba(215, 154, 48, 0.48);
  background: #FFF3D5;
}

.settings-ant-button.danger:not(:disabled):hover {
  color: #ffffff;
  border-color: rgba(180, 35, 46, 0.82);
  background: #B4232E;
}

.proxy-listen-card {
  display: grid;
  grid-template-columns: 32px minmax(0, 1fr);
  align-items: center;
  gap: 10px;
  min-width: 0;
  height: 40px;
  padding: 4px 8px 4px 6px;
  border: 1px solid var(--settings-brand-border);
  border-radius: 8px;
  background: linear-gradient(180deg, #ffffff, var(--settings-brand-tint));
  box-shadow:
    inset 0 1px 0 rgba(255, 255, 255, 0.74),
    0 3px 8px rgba(0, 114, 239, 0.08);
}

.proxy-listen-icon {
  display: grid;
  place-items: center;
  width: 28px;
  height: 28px;
  border-radius: 8px;
  color: var(--settings-brand);
  background: rgba(0, 114, 239, 0.1);
  font-size: 13px;
}

.proxy-listen-copy {
  display: flex;
  align-items: baseline;
  gap: 10px;
  min-width: 0;
  overflow: hidden;
  white-space: nowrap;
}

.proxy-listen-host {
  flex: 0 1 auto;
  min-width: 0;
  max-width: 42%;
  overflow: hidden;
  color: var(--settings-ink);
  font-family: Consolas, 'SFMono-Regular', monospace;
  font-size: 19px;
  font-weight: 600;
  letter-spacing: 0;
  line-height: 1;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.proxy-listen-meta {
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
  color: var(--settings-muted);
  font-size: 12px;
  font-weight: 500;
  letter-spacing: 0;
  line-height: 1.15;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.proxy-row-note {
  align-self: center;
  min-width: 0;
  color: var(--settings-muted);
  font-size: 12.5px;
  line-height: 1.35;
  font-weight: 500;
}

.proxy-switch-line {
  min-height: 38px;
}

.proxy-link-control {
  color: var(--settings-brand);
  text-decoration: none;
}

.proxy-link-control:hover {
  text-decoration: underline;
}

.certificate-line {
  display: grid;
  grid-template-columns: 96px max-content;
  align-items: center;
  gap: 8px;
  justify-self: start;
  min-width: 0;
  height: 38px;
}

.certificate-actions {
  display: inline-grid;
  grid-template-columns: 128px 128px;
  gap: 8px;
  justify-self: end;
}

.certificate-action {
  width: 128px;
  min-width: 128px;
  height: 38px !important;
  padding: 0 !important;
  box-shadow: none;
  white-space: nowrap;
}

.certificate-action.secondary {
  color: var(--settings-action-ink);
  background: var(--settings-action-tint);
}

.certificate-status {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  justify-self: start;
  min-width: 96px;
  height: 38px;
  padding: 0 12px;
  border-radius: 6px;
  font-size: 14px;
  font-weight: 600;
  white-space: nowrap;
}

.certificate-status.success {
  color: #0f684f;
  background: rgba(223, 243, 232, 0.86);
}

.certificate-status.warning {
  color: #a35c12;
  background: rgba(250, 235, 218, 0.84);
}

.certificate-status.error {
  color: #ad3c38;
  background: rgba(253, 226, 224, 0.88);
}

.certificate-status.checking {
  color: var(--settings-brand);
  background: rgba(222, 236, 251, 0.86);
}

.browse-line {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 64px;
  gap: 0;
}

.settings-page :deep(.browse-input .vxe-input--wrapper) {
  border-radius: 6px 0 0 6px;
}

.browse-action {
  min-width: 64px;
  border-left: 0;
  border-radius: 0 6px 6px 0;
}

.env-check-grid {
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 8px;
  margin-top: 12px;
}

.env-check {
  grid-template-columns: 22px minmax(0, 1fr);
  align-items: center;
  gap: 5px;
  min-width: 0;
  min-height: 58px;
  padding: 10px 8px;
}

.env-check-icon {
  grid-row: auto;
  align-self: center;
  width: 20px;
  color: var(--blue);
  font-size: 18px;
  text-align: center;
}

.env-check--green .env-check-icon,
.env-check-value.green {
  color: #1f8f69;
}

.env-check--orange .env-check-icon,
.env-check-value.orange {
  color: #df7a35;
}

.env-check--red .env-check-icon,
.env-check-value.red {
  color: #d9413f;
}

.env-check--blue .env-check-icon,
.env-check-value.blue {
  color: var(--blue);
}

.env-check-copy {
  display: grid;
  gap: 3px;
  min-width: 0;
}

.env-check strong,
.env-check small {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.env-check strong {
  font-size: 14px;
  font-weight: 600;
}

.env-check small {
  font-size: 12px;
  font-weight: 500;
}

.config-action-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  grid-auto-rows: minmax(56px, 1fr);
  align-items: stretch;
  justify-items: stretch;
  gap: 12px;
  margin-top: 12px;
  min-height: 132px;
}

.config-action-heading {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
}

.config-action-heading-icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 30px;
  height: 30px;
  flex: 0 0 30px;
  color: #15386F;
  border: 1px solid rgba(21, 56, 111, 0.12);
  border-radius: 9px;
  background:
    linear-gradient(145deg, rgba(255, 255, 255, 0.86), rgba(232, 244, 255, 0.78)),
    #EEF7FF;
}

.settings-panel-header .config-action-heading-icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-top: 0;
  overflow: visible;
  color: #15386F;
  font-size: 18px;
  line-height: 1;
  text-overflow: clip;
  white-space: normal;
}

.settings-panel-header .config-action-heading-icon .app-icon {
  width: 18px;
  height: 18px;
  vertical-align: 0;
}

.config-action-heading-copy {
  display: inline-flex;
  flex-direction: column;
  min-width: 0;
}

.settings-panel-header .config-action-heading-copy {
  display: inline-flex;
  flex-direction: column;
  margin-top: 0;
  overflow: visible;
  color: inherit;
  font-size: inherit;
  font-weight: inherit;
  line-height: inherit;
  text-overflow: clip;
  white-space: normal;
}

.config-action-grid :deep(.config-action-button) {
  --config-action-color: #357FD9;
  --config-action-hover-color: #2267B8;
  --config-action-disabled-color: rgba(53, 127, 217, 0.42);
  --config-action-border: rgba(88, 153, 229, 0.24);
  --config-action-hover-border: rgba(53, 127, 217, 0.42);
  --config-action-bg: #F6FBFF;
  --config-action-hover-bg: #EDF6FF;
  --config-action-active-bg: #E1EFFD;
  --config-action-shadow-color: transparent;

  display: inline-flex;
  align-items: center;
  justify-content: center;
  box-sizing: border-box;
  width: 100%;
  min-width: 0;
  height: 100% !important;
  min-height: 56px;
  padding: 6px 14px;
  color: var(--config-action-color);
  border-color: var(--config-action-border);
  border-radius: 7px;
  background: var(--config-action-bg);
  font-size: 15px;
  font-weight: 600;
  text-align: center;
  box-shadow: none;
  transition:
    color 180ms ease,
    border-color 180ms ease,
    background-color 180ms ease;
}

.config-action-grid :deep(.config-action-button .ant-btn-icon+span),
.config-action-grid :deep(.config-action-button span+.ant-btn-icon) {
  margin-inline-start: 9px;
}

.config-action-grid :deep(.config-action-button .ant-btn-icon) {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex: 0 0 auto;
}

.config-action-inline-icon {
  flex: 0 0 auto;
  font-size: 18px;
  line-height: 1;
}

.config-action-button.success {
  --config-action-color: #ffffff;
  --config-action-hover-color: #ffffff;
  --config-action-border: #35B889;
  --config-action-hover-border: #2EA579;
  --config-action-bg: #35B889;
  --config-action-hover-bg: #2EA579;
  --config-action-active-bg: #258C68;
  --config-action-shadow-color: transparent;
}

.config-action-button.primary {
  --config-action-color: #ffffff;
  --config-action-hover-color: #ffffff;
  --config-action-border: #357FD9;
  --config-action-hover-border: #2267B8;
  --config-action-bg: #357FD9;
  --config-action-hover-bg: #2267B8;
  --config-action-active-bg: #1D579C;
  --config-action-shadow-color: transparent;
}

.config-action-button.orange {
  --config-action-color: #ffffff;
  --config-action-hover-color: #ffffff;
  --config-action-border: #F28B3C;
  --config-action-hover-border: #DA7328;
  --config-action-bg: #F28B3C;
  --config-action-hover-bg: #DA7328;
  --config-action-active-bg: #BE5F20;
  --config-action-shadow-color: transparent;
}

.config-action-button.danger {
  --config-action-color: #ffffff;
  --config-action-hover-color: #ffffff;
  --config-action-border: #D74D4D;
  --config-action-hover-border: #BF3D3D;
  --config-action-bg: #D74D4D;
  --config-action-hover-bg: #BF3D3D;
  --config-action-active-bg: #A93434;
  --config-action-shadow-color: transparent;
}

.config-action-button.ghost {
  --config-action-color: #357FD9;
  --config-action-hover-color: #2267B8;
  --config-action-disabled-color: rgba(53, 127, 217, 0.42);
  --config-action-border: rgba(88, 153, 229, 0.28);
  --config-action-hover-border: rgba(53, 127, 217, 0.46);
  --config-action-bg: #F6FBFF;
  --config-action-hover-bg: #EDF6FF;
  --config-action-active-bg: #E1EFFD;
  --config-action-shadow-color: transparent;
}

.config-action-grid :deep(.ant-btn.config-action-button:not(:disabled):hover) {
  transform: none;
  color: var(--config-action-hover-color);
  border-color: var(--config-action-hover-border);
  background: var(--config-action-hover-bg);
  box-shadow: none;
}

.config-action-grid :deep(.ant-btn.config-action-button:not(:disabled):active) {
  transform: translateY(0);
  color: var(--config-action-hover-color);
  border-color: var(--config-action-hover-border);
  background: var(--config-action-active-bg);
  box-shadow: none;
  filter: none;
}

.config-action-grid :deep(.ant-btn.config-action-button:focus) {
  color: var(--config-action-color);
  border-color: var(--config-action-border);
  background: var(--config-action-bg);
}

.config-action-grid :deep(.ant-btn.config-action-button:focus-visible) {
  outline: 3px solid rgba(53, 127, 217, 0.24);
  outline-offset: 2px;
}

.config-action-grid :deep(.ant-btn.config-action-button:disabled) {
  transform: none;
  color: var(--config-action-disabled-color);
  opacity: 0.58;
  border-color: var(--config-action-border);
  background: var(--config-action-bg);
  box-shadow: none;
  filter: saturate(0.58);
  cursor: not-allowed;
}

:global(.collector-app.dark .config-action-grid .config-action-button.ghost) {
  --config-action-color: #BFDDFB;
  --config-action-hover-color: #D7EAFF;
  --config-action-disabled-color: rgba(191, 221, 251, 0.46);
  --config-action-border: rgba(103, 163, 235, 0.3);
  --config-action-hover-border: rgba(103, 163, 235, 0.48);
  --config-action-bg: rgba(22, 46, 74, 0.58);
  --config-action-hover-bg: rgba(30, 60, 94, 0.72);
  --config-action-active-bg: rgba(20, 42, 68, 0.82);
  --config-action-shadow-color: transparent;
}

.restart-action,
.connection-action {
  height: 36px !important;
}

.test-button {
  width: 100%;
  margin-top: 0;
}

.config-note {
  min-height: 0;
}

.note-art {
  right: 0;
  bottom: 0;
  width: 156px;
  opacity: 0.28;
}

.note-list {
  display: grid;
  gap: 10px;
  margin: 18px 0 0;
  min-height: 118px;
  padding: 18px 18px 18px 16px;
  border: 1px solid rgba(104, 141, 181, 0.18);
  border-radius: 8px;
  background: rgba(222, 238, 250, 0.44);
  list-style: none;
}

.config-note .soft-note,
.note-list .soft-note {
  display: grid;
  grid-template-columns: 12px minmax(0, 1fr);
  gap: 10px;
  align-items: start;
  line-height: 1.35;
  color: var(--settings-muted);
  font-size: 12.8px;
  font-weight: 500;
}

.note-list .soft-note i {
  margin-top: 7px;
  color: var(--settings-brand);
  font-size: 5px;
}

.settings-label {
  display: grid;
  gap: 0;
  min-width: 0;
}

.form-row .settings-label,
.settings-label strong {
  color: var(--settings-ink);
  font-size: 14px;
  font-weight: 600;
}

.settings-label small {
  display: none;
  color: var(--settings-muted);
  font-size: 11.5px;
  line-height: 1.32;
  font-weight: 500;
}

.status-value {
  color: var(--ink-strong);
  font-size: 15px;
}

.status-value.success {
  color: var(--green);
}

@media (prefers-reduced-motion: reduce) {
  .config-action-grid :deep(.config-action-button) {
    transition: none;
    transform: none;
  }

  .config-action-grid :deep(.config-action-button:hover),
  .config-action-grid :deep(.config-action-button:active) {
    transform: none;
    filter: none;
  }

}

.mitm-cert-dialog-backdrop {
  position: fixed;
  inset: 0;
  z-index: 120;
  display: grid;
  place-items: center;
  padding: 24px;
  background: rgba(25, 43, 65, 0.42);
}

.mitm-cert-dialog {
  width: min(760px, calc(100vw - 48px));
  max-height: min(720px, calc(100vh - 48px));
  padding: 20px;
  border: 1px solid rgba(104, 141, 181, 0.32);
  border-radius: 10px;
  background: #FBFDFF;
  box-shadow: 0 18px 42px rgba(22, 45, 73, 0.28);
  isolation: isolate;
  overflow: hidden;
}

.mitm-cert-dialog-header {
  display: grid;
  grid-template-columns: 42px minmax(0, 1fr);
  gap: 12px;
  align-items: start;
}

.mitm-cert-dialog-copy {
  min-width: 0;
}

.mitm-cert-dialog-icon {
  display: grid;
  place-items: center;
  width: 42px;
  height: 42px;
  border-radius: 8px;
  color: #b53531;
  background: rgba(253, 226, 224, 0.88);
}

.mitm-cert-dialog h3 {
  margin: 0;
  color: var(--ink-strong);
  font-size: 18px;
  font-weight: 600;
}

.mitm-cert-dialog p {
  margin: 6px 0 0;
  color: var(--ink);
  font-size: 13px;
  font-weight: 500;
  line-height: 1.5;
  overflow-wrap: anywhere;
}

.mitm-cert-source-card {
  display: grid;
  gap: 14px;
  margin-top: 16px;
  padding: 14px;
  border: 1px solid rgba(104, 141, 181, 0.28);
  border-radius: 10px;
  background: rgba(246, 249, 252, 0.94);
}

.mitm-cert-source-card--project {
  border-color: rgba(49, 112, 173, 0.34);
  background: rgba(240, 247, 254, 0.94);
}

.mitm-cert-source-card--system {
  border-color: rgba(91, 105, 133, 0.34);
  background: rgba(247, 248, 252, 0.96);
}

.mitm-cert-source-heading {
  display: grid;
  grid-template-columns: 36px minmax(0, 1fr) max-content;
  gap: 10px;
  align-items: center;
  min-width: 0;
}

.mitm-cert-source-icon {
  display: grid;
  place-items: center;
  width: 36px;
  height: 36px;
  border-radius: 8px;
  color: #315f9d;
  background: rgba(222, 237, 252, 0.9);
}

.mitm-cert-source-card--system .mitm-cert-source-icon {
  color: #4d5871;
  background: rgba(231, 235, 243, 0.94);
}

.mitm-cert-source-title {
  display: grid;
  gap: 3px;
  min-width: 0;
}

.mitm-cert-source-title strong {
  color: var(--ink-strong);
  font-size: 14px;
  font-weight: 600;
}

.mitm-cert-source-title span {
  color: var(--ink-muted);
  font-size: 12px;
  font-weight: 500;
  line-height: 1.45;
}

.mitm-cert-badge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  justify-self: end;
  max-width: 100%;
  min-height: 24px;
  padding: 0 9px;
  border: 1px solid rgba(104, 141, 181, 0.24);
  border-radius: 999px;
  color: #42536f;
  background: rgba(255, 255, 255, 0.82);
  font-size: 12px;
  font-weight: 600;
  line-height: 1;
  white-space: nowrap;
}

.mitm-cert-badge.success {
  color: #116149;
  border-color: rgba(33, 134, 99, 0.28);
  background: rgba(224, 245, 237, 0.9);
}

.mitm-cert-badge.warning {
  color: #8a5310;
  border-color: rgba(205, 132, 39, 0.32);
  background: rgba(255, 241, 220, 0.92);
}

.mitm-cert-badge.info {
  color: #315f9d;
  border-color: rgba(49, 112, 173, 0.26);
  background: rgba(232, 241, 252, 0.92);
}

.mitm-cert-list {
  display: grid;
  gap: 10px;
  max-height: 430px;
  margin-top: 16px;
  padding-right: 4px;
  overflow: auto;
}

.mitm-cert-source-card .mitm-cert-list {
  margin-top: 0;
}

.mitm-cert-item {
  display: grid;
  gap: 10px;
  padding: 12px;
  border: 1px solid rgba(104, 141, 181, 0.2);
  border-radius: 8px;
  background: rgba(242, 247, 252, 0.72);
}

.mitm-cert-item-main {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  min-width: 0;
}

.mitm-cert-item-main strong {
  min-width: 0;
  color: var(--ink-strong);
  font-size: 14px;
  font-weight: 600;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.mitm-cert-item-main span {
  flex: none;
  color: var(--settings-brand);
  font-family: Consolas, 'SFMono-Regular', monospace;
  font-size: 12px;
  font-weight: 500;
}

.mitm-cert-meta {
  display: grid;
  gap: 8px;
  margin: 0;
}

.mitm-cert-meta div {
  display: grid;
  grid-template-columns: 54px minmax(0, 1fr);
  gap: 10px;
}

.mitm-cert-meta dt {
  color: var(--ink-muted);
  font-size: 12px;
  font-weight: 600;
}

.mitm-cert-meta dd {
  min-width: 0;
  margin: 0;
  color: var(--ink);
  font-size: 12px;
  font-weight: 500;
  overflow-wrap: anywhere;
}

.mitm-cert-empty {
  margin-top: 16px;
  padding: 18px;
  border: 1px dashed rgba(104, 141, 181, 0.34);
  border-radius: 8px;
  color: var(--ink-muted);
  background: rgba(242, 247, 252, 0.58);
  font-size: 13px;
  font-weight: 600;
  text-align: center;
}

.mitm-cert-empty--inline {
  margin-top: 0;
  padding: 14px;
}

.mitm-cert-dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 10px;
  margin-top: 18px;
}

.mitm-cert-dialog-actions .config-action-button {
  width: auto;
  min-width: 112px;
}

.mitm-cert-dialog-actions :deep(.ant-btn.config-action-button.ghost) {
  --config-action-color: #15386F;
  --config-action-border: rgba(53, 127, 217, 0.32);
  --config-action-bg: #EAF4FF;
  --config-action-hover-color: #0F2F61;
  --config-action-hover-border: rgba(53, 127, 217, 0.48);
  --config-action-hover-bg: #DDEEFF;
  color: #15386F;
  border-color: rgba(53, 127, 217, 0.32);
  background: #EAF4FF;
}

.mitm-cert-dialog-actions :deep(.ant-btn.config-action-button.ghost:not(:disabled):hover) {
  color: #0F2F61;
  border-color: rgba(53, 127, 217, 0.48);
  background: #DDEEFF;
}

.mitm-cert-dialog-actions :deep(.ant-btn.config-action-button.primary) {
  --config-action-color: #ffffff;
  --config-action-border: #357FD9;
  --config-action-bg: #357FD9;
  --config-action-hover-color: #ffffff;
  --config-action-hover-border: #2267B8;
  --config-action-hover-bg: #2267B8;
  --config-action-active-bg: #1D579C;
  color: #ffffff;
  border-color: #357FD9;
  background: #357FD9;
  opacity: 1;
}

.mitm-cert-dialog-actions :deep(.ant-btn.config-action-button.primary:not(:disabled):hover) {
  color: #ffffff;
  border-color: #2267B8;
  background: #2267B8;
}

.mitm-cert-dialog-actions :deep(.ant-btn.config-action-button.primary:not(:disabled):active) {
  color: #ffffff;
  border-color: #1D579C;
  background: #1D579C;
}

.diagnostic-result-dialog {
  display: flex;
  flex-direction: column;
  width: min(860px, calc(100vw - 48px));
  max-height: min(760px, calc(100vh - 40px));
  padding: 20px;
  border: 1px solid rgba(89, 130, 181, 0.34);
  border-radius: 10px;
  background: #F8FBFF;
  box-shadow: 0 6px 8px rgba(22, 45, 73, 0.18);
  overflow: hidden;
}

.diagnostic-result-dialog-header {
  flex: none;
  display: grid;
  grid-template-columns: 42px minmax(0, 1fr);
  gap: 12px;
  align-items: start;
}

.diagnostic-result-dialog-icon {
  display: grid;
  place-items: center;
  width: 42px;
  height: 42px;
  border: 1px solid rgba(21, 56, 111, 0.16);
  border-radius: 8px;
  color: #15386F;
  background: #E8F2FF;
}

.diagnostic-result-dialog--success .diagnostic-result-dialog-icon {
  color: #167255;
  border-color: rgba(22, 114, 85, 0.2);
  background: #E2F4ED;
}

.diagnostic-result-dialog--warning .diagnostic-result-dialog-icon {
  color: #8B5B13;
  border-color: rgba(139, 91, 19, 0.2);
  background: #FFF1D6;
}

.diagnostic-result-dialog--error .diagnostic-result-dialog-icon {
  color: #A82D39;
  border-color: rgba(168, 45, 57, 0.2);
  background: #FDE7EA;
}

.diagnostic-result-dialog h3 {
  margin: 0;
  color: #1F3148;
  font-size: 18px;
  font-weight: 600;
}

.diagnostic-result-dialog p {
  margin: 6px 0 0;
  color: #52657B;
  font-size: 13px;
  font-weight: 500;
  line-height: 1.5;
}

.reset-defaults-dialog {
  width: min(560px, calc(100vw - 48px));
}

.reset-defaults-dialog-icon {
  color: #9A5A12;
  border-color: rgba(194, 113, 28, 0.24);
  background: #FFF1DE;
}

.reset-defaults-dialog-notice {
  display: grid;
  gap: 6px;
  margin-top: 18px;
  padding: 14px 16px;
  border-radius: 8px;
  color: #5D472C;
  background: #FFF7EA;
}

.reset-defaults-dialog-notice strong {
  color: #7E4B12;
  font-size: 14px;
  font-weight: 600;
}

.reset-defaults-dialog-notice span {
  font-size: 13px;
  font-weight: 400;
  line-height: 1.55;
}

.diagnostic-result-list {
  display: grid;
  flex: 1;
  gap: 10px;
  min-height: 0;
  max-height: min(560px, calc(100vh - 220px));
  margin: 16px 0 0;
  padding-right: 6px;
  overflow-y: auto;
  overflow-x: hidden;
}

.diagnostic-result-item {
  display: grid;
  gap: 10px;
  min-width: 0;
  padding: 12px;
  border: 1px solid #C9D9EB;
  border-radius: 8px;
  background: #F1F6FC;
}

.diagnostic-result-item--operation {
  gap: 6px;
  padding: 9px 11px;
  border-color: #C6D8EA;
  background: #EDF4FB;
}

.diagnostic-result-item--discarded,
.diagnostic-result-item--warning.diagnostic-result-item--discarded {
  border-color: #E6C47D;
  background: #FFF7E8;
}

.diagnostic-result-item--discarded dt {
  color: #78551D;
}

.diagnostic-result-item--discarded dd {
  color: #5E431A;
}

.diagnostic-result-item--article {
  border-color: #B8D5C8;
  background: #F0F8F4;
}

.diagnostic-result-item--error {
  border-color: #E5B0B6;
  background: #FDEFF1;
}

.diagnostic-result-item.diagnostic-result-item--split {
  grid-template-columns: 116px minmax(0, 1fr) 92px minmax(0, 0.75fr);
  gap: 8px 12px;
  align-items: start;
}

.diagnostic-result-item-main {
  display: grid;
  grid-template-columns: minmax(128px, 176px) minmax(0, 1fr);
  gap: 10px 14px;
  align-items: start;
  min-width: 0;
}

.diagnostic-result-cells {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 8px;
  min-width: 0;
  padding-top: 10px;
  border-top: 1px solid rgba(104, 141, 181, 0.16);
}

.diagnostic-result-cell {
  display: grid;
  grid-template-columns: minmax(96px, 0.48fr) minmax(0, 1fr);
  gap: 8px;
  align-items: start;
  min-width: 0;
}

.diagnostic-result-item--split .diagnostic-result-cell {
  grid-column: span 2;
  grid-template-columns: minmax(92px, 0.45fr) minmax(0, 1fr);
}

.diagnostic-result-item--split .diagnostic-result-cell:nth-child(even) {
  grid-column: 3 / span 2;
}

.diagnostic-result-item dt {
  min-width: 0;
  color: #52657B;
  font-size: 12px;
  font-weight: 600;
  line-height: 1.45;
  overflow-wrap: anywhere;
}

.diagnostic-result-item dd {
  min-width: 0;
  margin: 0;
  color: #24364B;
  font-size: 12px;
  font-weight: 500;
  line-height: 1.45;
  overflow-wrap: anywhere;
  word-break: break-word;
  white-space: pre-wrap;
}

.diagnostic-result-dialog .mitm-cert-dialog-actions {
  flex: none;
  margin-top: 16px;
  padding-top: 14px;
  border-top: 1px solid #D7E3F0;
}

.diagnostic-result-dialog :deep(.ant-btn.config-action-button.primary) {
  --config-action-color: #ffffff;
  --config-action-border: #357FD9;
  --config-action-bg: #357FD9;
  --config-action-hover-color: #ffffff;
  --config-action-hover-border: #2267B8;
  --config-action-hover-bg: #2267B8;
  --config-action-active-bg: #1D579C;
  min-width: 128px;
  min-height: 40px;
  color: #ffffff;
  border-color: #357FD9;
  background: #357FD9;
  opacity: 1;
}

.diagnostic-result-dialog :deep(.ant-btn.config-action-button.primary:not(:disabled):hover) {
  color: #ffffff;
  border-color: #2267B8;
  background: #2267B8;
}

.diagnostic-result-dialog :deep(.ant-btn.config-action-button.primary:not(:disabled):active) {
  color: #ffffff;
  border-color: #1D579C;
  background: #1D579C;
}

@media (max-width: 760px) {
  .diagnostic-result-dialog {
    width: min(100%, calc(100vw - 32px));
    padding: 16px;
  }

  .diagnostic-result-list {
    max-height: calc(100vh - 210px);
  }

  .diagnostic-result-item-main,
  .diagnostic-result-cell {
    grid-template-columns: 1fr;
    gap: 4px;
  }
}

.mitm-cert-dialog-fade-enter-active,
.mitm-cert-dialog-fade-leave-active {
  transition: opacity 160ms ease;
}

.mitm-cert-dialog-fade-enter-active .mitm-cert-dialog,
.mitm-cert-dialog-fade-leave-active .mitm-cert-dialog {
  transition: transform 160ms ease;
}

.mitm-cert-dialog-fade-enter-from,
.mitm-cert-dialog-fade-leave-to {
  opacity: 0;
}

.mitm-cert-dialog-fade-enter-from .mitm-cert-dialog,
.mitm-cert-dialog-fade-leave-to .mitm-cert-dialog {
  transform: translateY(8px);
}

:global(.collector-app.dark) .settings-primary-nav,
:global(.collector-app.dark) .settings-secondary-nav,
:global(.collector-app.dark) .settings-detail-pane,
:global(.collector-app.dark) .env-panel,
:global(.collector-app.dark) .config-actions,
:global(.collector-app.dark) .config-note {
  border-color: rgba(98, 141, 196, 0.28);
  background:
    linear-gradient(180deg, rgba(23, 34, 54, 0.9), rgba(16, 25, 41, 0.82));
  box-shadow:
    inset 0 1px 0 rgba(214, 226, 244, 0.055),
    0 0 0 1px rgba(67, 116, 184, 0.1),
    0 5px 10px rgba(0, 0, 0, 0.14);
}

:global(.collector-app.dark) .settings-primary-item,
:global(.collector-app.dark) .settings-secondary-item {
  color: #b9c8dd;
}

:global(.collector-app.dark) .settings-primary-item:hover,
:global(.collector-app.dark) .settings-secondary-item:hover {
  background: rgba(36, 56, 84, 0.34);
}

:global(.collector-app.dark) .settings-primary-item.active,
:global(.collector-app.dark) .settings-secondary-item.active {
  color: #CFE2FF;
  background: rgba(89, 145, 220, 0.18);
}

:global(.collector-app.dark) .settings-primary-icon,
:global(.collector-app.dark) .settings-secondary-icon {
  color: #CFE2FF;
}

:global(.collector-app.dark) .settings-config-row,
:global(.collector-app.dark) .config-field-card,
:global(.collector-app.dark) .detail-row,
:global(.collector-app.dark) .settings-explain-card,
:global(.collector-app.dark) .settings-category-overview,
:global(.collector-app.dark) .detail-note {
  border-color: rgba(128, 153, 188, 0.14);
  background: rgba(15, 24, 39, 0.5);
}

:global(.collector-app.dark) .diagnostic-action-grid .settings-config-row {
  border-color: rgba(128, 153, 188, 0.2);
  background: rgba(15, 24, 39, 0.52);
  box-shadow:
    0 4px 10px rgba(0, 0, 0, 0.22),
    inset 0 1px 0 rgba(214, 226, 244, 0.045);
}

:global(.collector-app.dark) .diagnostic-action-grid .settings-config-row:hover {
  border-color: rgba(111, 154, 211, 0.34);
  box-shadow:
    0 6px 14px rgba(0, 0, 0, 0.28),
    inset 0 1px 0 rgba(214, 226, 244, 0.06);
}

:global(.collector-app.dark) .settings-config-row.readonly,
:global(.collector-app.dark) .config-field-card.readonly {
  background: rgba(20, 30, 46, 0.58);
}

:global(.collector-app.dark) .config-field-card.success {
  border-color: rgba(64, 142, 111, 0.24);
  background: rgba(64, 142, 111, 0.12);
}

:global(.collector-app.dark) .config-field-card.warning,
:global(.collector-app.dark) .settings-explain-card.warning {
  border-color: rgba(216, 180, 95, 0.24);
  background: rgba(101, 78, 43, 0.22);
}

:global(.collector-app.dark) .settings-config-copy strong,
:global(.collector-app.dark) .config-field-card strong,
:global(.collector-app.dark) .settings-explain-card h3,
:global(.collector-app.dark) .settings-category-overview h3,
:global(.collector-app.dark) .settings-panel-header strong,
:global(.collector-app.dark) .settings-label strong {
  color: #dce7f5;
}

:global(.collector-app.dark) .settings-config-copy small,
:global(.collector-app.dark) .config-field-card small,
:global(.collector-app.dark) .settings-explain-card p,
:global(.collector-app.dark) .settings-category-overview p,
:global(.collector-app.dark) .detail-bullet-list,
:global(.collector-app.dark) .detail-note,
:global(.collector-app.dark) .settings-panel-header span,
:global(.collector-app.dark) .settings-label small {
  color: #8ea2bd;
}

:global(.collector-app.dark) .config-action-heading-icon {
  color: #BFDDFB;
  border-color: rgba(103, 163, 235, 0.22);
  background:
    linear-gradient(145deg, rgba(33, 58, 91, 0.9), rgba(19, 36, 61, 0.82)),
    rgba(22, 46, 74, 0.72);
}

:global(.collector-app.dark) .settings-inline-input,
:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-input),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-select .ant-select-selector),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-picker),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-number.ant-input-number) {
  color: #cbd8ea;
  border-color: rgba(128, 153, 188, 0.2);
  background: rgba(15, 24, 39, 0.62);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.045);
}

:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-select .ant-select-selection-item),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-select .ant-select-selection-placeholder),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-picker input),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-picker .ant-picker-input > input),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-picker .ant-picker-separator),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-picker .ant-picker-suffix),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-control.ant-picker .ant-picker-clear),
:global(.collector-app.dark) .settings-page :deep(.settings-ant-number .ant-input-number-input),
:global(.collector-app.dark) .settings-inline-input[readonly] {
  color: #dce7f5;
}

:global(.collector-app.dark) .settings-page :deep(.settings-ant-number .ant-input-number-handler-wrap) {
  background: rgba(28, 43, 65, 0.46);
}

:global(.collector-app.dark) .settings-page :deep(.settings-ant-number .ant-input-number-handler) {
  color: #9db1cc;
  border-color: rgba(128, 153, 188, 0.2);
}

:global(.collector-app.dark) .settings-page :deep(.settings-ant-number .ant-input-number-handler:hover) {
  color: #dceaff;
}

:global(.collector-app.dark) .settings-row-button,
:global(.collector-app.dark) .settings-ant-button {
  color: #BFE9E3;
  border-color: rgba(112, 198, 186, 0.32);
  background: rgba(35, 70, 65, 0.62);
  box-shadow: none;
}

:global(.collector-app.dark) .settings-ant-button.primary {
  color: #f0f6ff;
  border-color: #2F8D82;
  background: #2F8D82;
}

:global(.collector-app.dark) .settings-ant-button.success {
  color: #f0f6ff;
  border-color: #2F8D82;
  background: #2F8D82;
}

:global(.collector-app.dark) .settings-ant-button.orange {
  color: #f7d99a;
  border-color: rgba(216, 180, 95, 0.26);
  background: rgba(69, 50, 31, 0.72);
}

:global(.collector-app.dark) .settings-ant-button.danger {
  color: #fff1f0;
  border-color: rgba(181, 83, 84, 0.34);
  background: rgba(105, 64, 70, 0.8);
}

:global(.collector-app.dark) .settings-ant-button.ghost {
  color: #BFE9E3;
  background: rgba(35, 70, 65, 0.62);
}

:global(.collector-app.dark) .settings-ant-button:not(:disabled):hover {
  color: #BFE9E3;
  border-color: rgba(112, 198, 186, 0.5);
  background: rgba(43, 86, 79, 0.72);
  box-shadow: none;
}

:global(.collector-app.dark) .settings-ant-button.primary:not(:disabled):hover,
:global(.collector-app.dark) .settings-ant-button.success:not(:disabled):hover {
  color: #f0f6ff;
  border-color: #3B9C90;
  background: #3B9C90;
}

:global(.collector-app.dark) .settings-ant-button.orange:not(:disabled):hover {
  color: #ffe4aa;
  border-color: rgba(216, 180, 95, 0.38);
  background: rgba(84, 61, 36, 0.82);
}

:global(.collector-app.dark) .settings-ant-button.danger:not(:disabled):hover {
  color: #fff1f0;
  border-color: rgba(181, 83, 84, 0.44);
  background: rgba(116, 70, 75, 0.86);
}

:global(.collector-app.dark) .proxy-listen-card {
  border-color: rgba(128, 153, 188, 0.18);
  background: rgba(15, 24, 39, 0.58);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.045);
}

:global(.collector-app.dark) .proxy-listen-icon {
  color: #8fbded;
  background: rgba(45, 117, 214, 0.16);
}

:global(.collector-app.dark) .proxy-listen-host {
  color: #dce7f5;
}

:global(.collector-app.dark) .diagnostic-action-grid :deep(.diagnostic-action-button) {
  --diagnostic-action-color: #BFE9E3;
  --diagnostic-action-border: rgba(112, 198, 186, 0.32);
  --diagnostic-action-bg: rgba(35, 70, 65, 0.62);
  --diagnostic-action-hover-color: #D8F5F0;
  --diagnostic-action-hover-border: rgba(112, 198, 186, 0.5);
  --diagnostic-action-hover-bg: rgba(43, 86, 79, 0.72);
  --diagnostic-action-icon-color: #BFE9E3;
  --diagnostic-action-icon-bg: rgba(112, 198, 186, 0.16);
  color: var(--diagnostic-action-color);
  border-color: var(--diagnostic-action-border);
  background: var(--diagnostic-action-bg);
  box-shadow: none;
}

:global(.collector-app.dark) .diagnostic-action-grid :deep(.diagnostic-action-button:hover) {
  color: var(--diagnostic-action-hover-color);
  border-color: var(--diagnostic-action-hover-border);
  background: var(--diagnostic-action-hover-bg);
}

:global(.collector-app.dark) .diagnostic-action-grid :deep(.diagnostic-action-button.success) {
  --diagnostic-action-color: #f0f6ff;
  --diagnostic-action-border: #2F8D82;
  --diagnostic-action-bg: #2F8D82;
  --diagnostic-action-hover-color: #f0f6ff;
  --diagnostic-action-hover-border: #3B9C90;
  --diagnostic-action-hover-bg: #3B9C90;
  --diagnostic-action-icon-color: #f0f6ff;
  --diagnostic-action-icon-bg: rgba(255, 255, 255, 0.18);
  color: #f0f6ff;
  border-color: #2F8D82;
  background: #2F8D82;
}

:global(.collector-app.dark) .diagnostic-action-grid :deep(.diagnostic-action-button.primary) {
  --diagnostic-action-color: #f0f6ff;
  --diagnostic-action-border: #2F6EBD;
  --diagnostic-action-bg: #2F6EBD;
  --diagnostic-action-hover-color: #ffffff;
  --diagnostic-action-hover-border: #3B7ED0;
  --diagnostic-action-hover-bg: #3B7ED0;
  --diagnostic-action-icon-color: #ffffff;
  --diagnostic-action-icon-bg: rgba(255, 255, 255, 0.18);
  color: #f0f6ff;
  border-color: #2F6EBD;
  background: #2F6EBD;
}

:global(.collector-app.dark) .diagnostic-action-grid :deep(.diagnostic-action-button.blue) {
  --diagnostic-action-color: #BFDDFB;
  --diagnostic-action-border: rgba(103, 163, 235, 0.32);
  --diagnostic-action-bg: rgba(29, 62, 99, 0.72);
  --diagnostic-action-hover-color: #D7EAFF;
  --diagnostic-action-hover-border: rgba(103, 163, 235, 0.5);
  --diagnostic-action-hover-bg: rgba(37, 75, 118, 0.82);
  --diagnostic-action-icon-color: #BFDDFB;
  --diagnostic-action-icon-bg: rgba(103, 163, 235, 0.16);
  color: #BFDDFB;
  border-color: rgba(103, 163, 235, 0.32);
  background: rgba(29, 62, 99, 0.72);
}

:global(.collector-app.dark) .diagnostic-action-grid :deep(.diagnostic-action-button.teal) {
  --diagnostic-action-color: #A7EEE5;
  --diagnostic-action-border: rgba(67, 190, 177, 0.34);
  --diagnostic-action-bg: rgba(26, 78, 73, 0.72);
  --diagnostic-action-hover-color: #C9FFF8;
  --diagnostic-action-hover-border: rgba(67, 190, 177, 0.5);
  --diagnostic-action-hover-bg: rgba(33, 94, 88, 0.84);
  --diagnostic-action-icon-color: #A7EEE5;
  --diagnostic-action-icon-bg: rgba(67, 190, 177, 0.18);
  color: #A7EEE5;
  border-color: rgba(67, 190, 177, 0.34);
  background: rgba(26, 78, 73, 0.72);
}

:global(.collector-app.dark) .diagnostic-action-grid :deep(.diagnostic-action-button.purple) {
  --diagnostic-action-color: #D7C7FF;
  --diagnostic-action-border: rgba(158, 129, 224, 0.32);
  --diagnostic-action-bg: rgba(57, 45, 87, 0.72);
  --diagnostic-action-hover-color: #E5DAFF;
  --diagnostic-action-hover-border: rgba(158, 129, 224, 0.5);
  --diagnostic-action-hover-bg: rgba(68, 54, 105, 0.82);
  --diagnostic-action-icon-color: #D7C7FF;
  --diagnostic-action-icon-bg: rgba(158, 129, 224, 0.16);
  color: #D7C7FF;
  border-color: rgba(158, 129, 224, 0.32);
  background: rgba(57, 45, 87, 0.72);
}

:global(.collector-app.dark) .diagnostic-action-grid :deep(.diagnostic-action-button.orange) {
  --diagnostic-action-color: #f7d99a;
  --diagnostic-action-border: rgba(216, 180, 95, 0.26);
  --diagnostic-action-bg: rgba(69, 50, 31, 0.72);
  --diagnostic-action-hover-color: #ffe4aa;
  --diagnostic-action-hover-border: rgba(216, 180, 95, 0.38);
  --diagnostic-action-hover-bg: rgba(84, 61, 36, 0.82);
  --diagnostic-action-icon-color: #f7d99a;
  --diagnostic-action-icon-bg: rgba(216, 180, 95, 0.16);
  color: #f7d99a;
  border-color: rgba(216, 180, 95, 0.26);
  background: rgba(69, 50, 31, 0.72);
}

:global(.collector-app.dark) .diagnostic-action-grid :deep(.diagnostic-action-button.danger) {
  --diagnostic-action-color: #fff1f0;
  --diagnostic-action-border: rgba(181, 83, 84, 0.34);
  --diagnostic-action-bg: rgba(105, 64, 70, 0.8);
  --diagnostic-action-hover-color: #fff1f0;
  --diagnostic-action-hover-border: rgba(181, 83, 84, 0.44);
  --diagnostic-action-hover-bg: rgba(116, 70, 75, 0.86);
  --diagnostic-action-icon-color: #fff1f0;
  --diagnostic-action-icon-bg: rgba(181, 83, 84, 0.18);
  color: #fff1f0;
  border-color: rgba(181, 83, 84, 0.34);
  background: rgba(105, 64, 70, 0.8);
}

:global(.collector-app.dark) .diagnostic-action-icon {
  color: var(--diagnostic-action-icon-color);
  background: var(--diagnostic-action-icon-bg);
}

:global(.collector-app.dark) .diagnostic-action-copy strong {
  color: #dce7f5;
}

:global(.collector-app.dark) .diagnostic-action-copy small {
  color: #8ea2bd;
}

:global(.collector-app.dark) .env-check {
  border-color: rgba(128, 153, 188, 0.16);
  background: rgba(15, 24, 39, 0.58);
}

:global(.collector-app.dark) .env-check strong {
  color: #dce7f5;
}

:global(.collector-app.dark) .env-check small {
  color: #8ea2bd;
}

:global(.collector-app.dark) .config-action-grid :deep(.config-action-button) {
  --config-action-color: #BFDDFB;
  --config-action-hover-color: #D7EAFF;
  --config-action-disabled-color: rgba(191, 221, 251, 0.46);
  --config-action-border: rgba(103, 163, 235, 0.28);
  --config-action-hover-border: rgba(103, 163, 235, 0.42);
  --config-action-bg: rgba(28, 52, 82, 0.68);
  --config-action-hover-bg: rgba(35, 64, 100, 0.78);
  --config-action-active-bg: rgba(24, 45, 72, 0.86);
  --config-action-shadow-color: transparent;
}

:global(.collector-app.dark) .config-action-button.primary {
  --config-action-border: #357FD9;
  --config-action-hover-border: #4B94EE;
  --config-action-bg: #357FD9;
  --config-action-hover-bg: #4B94EE;
}

:global(.collector-app.dark) .config-action-button.success {
  --config-action-border: #2EA579;
  --config-action-hover-border: #35B889;
  --config-action-bg: #2EA579;
  --config-action-hover-bg: #35B889;
}

:global(.collector-app.dark) .config-action-button.orange {
  --config-action-color: #ffffff;
  --config-action-hover-color: #ffffff;
  --config-action-border: #DA7328;
  --config-action-hover-border: #F28B3C;
  --config-action-bg: #C76624;
  --config-action-hover-bg: #DA7328;
}

:global(.collector-app.dark) .config-action-button.danger {
  --config-action-border: #BF3D3D;
  --config-action-bg: #B33A3A;
  --config-action-hover-bg: #CB4949;
}

:global(.collector-app.dark) .config-action-button.ghost {
  --config-action-color: #BFDDFB;
  --config-action-hover-color: #D7EAFF;
  --config-action-border: rgba(103, 163, 235, 0.3);
  --config-action-hover-border: rgba(103, 163, 235, 0.48);
  --config-action-bg: rgba(22, 46, 74, 0.58);
  --config-action-hover-bg: rgba(30, 60, 94, 0.72);
}

:global(.collector-app.dark) .mitm-cert-dialog-actions :deep(.ant-btn.config-action-button.primary) {
  --config-action-color: #ffffff;
  --config-action-border: #3D7FCC;
  --config-action-bg: #3D7FCC;
  --config-action-hover-color: #ffffff;
  --config-action-hover-border: #4B8DDB;
  --config-action-hover-bg: #4B8DDB;
  --config-action-active-bg: #2E68AE;
  color: #ffffff;
  border-color: #3D7FCC;
  background: #3D7FCC;
  opacity: 1;
}

:global(.collector-app.dark) .mitm-cert-dialog-actions :deep(.ant-btn.config-action-button.primary:not(:disabled):hover) {
  color: #ffffff;
  border-color: #4B8DDB;
  background: #4B8DDB;
}

:global(.collector-app.dark) .mitm-cert-dialog-actions :deep(.ant-btn.config-action-button.primary:not(:disabled):active) {
  color: #ffffff;
  border-color: #2E68AE;
  background: #2E68AE;
}

:global(.collector-app.dark) .mitm-cert-dialog-backdrop {
  background: rgba(7, 12, 22, 0.58);
}

:global(.collector-app.dark) .mitm-cert-dialog,
:global(.collector-app.dark) .diagnostic-result-dialog {
  border-color: rgba(111, 154, 211, 0.34);
  background: #16243A;
  box-shadow: 0 6px 8px rgba(0, 0, 0, 0.28);
}

:global(.collector-app.dark) .reset-defaults-dialog-icon {
  color: #F1C27D;
  border-color: rgba(218, 159, 77, 0.3);
  background: rgba(116, 75, 28, 0.5);
}

:global(.collector-app.dark) .reset-defaults-dialog-notice {
  color: #D9C7AD;
  background: rgba(88, 59, 25, 0.46);
}

:global(.collector-app.dark) .reset-defaults-dialog-notice strong {
  color: #F1C27D;
}

:global(.collector-app.dark) .mitm-cert-source-card,
:global(.collector-app.dark) .mitm-cert-item,
:global(.collector-app.dark) .mitm-cert-empty,
:global(.collector-app.dark) .diagnostic-result-item {
  border-color: rgba(126, 161, 205, 0.22);
  background: #1C2A40;
}

:global(.collector-app.dark) .diagnostic-result-cells {
  border-top-color: rgba(128, 153, 188, 0.14);
}

:global(.collector-app.dark) .diagnostic-result-item--operation {
  border-color: rgba(105, 157, 219, 0.28);
  background: #1A304B;
}

:global(.collector-app.dark) .diagnostic-result-item--discarded,
:global(.collector-app.dark) .diagnostic-result-item--warning.diagnostic-result-item--discarded {
  border-color: rgba(218, 172, 87, 0.36);
  background: #3A3020;
}

:global(.collector-app.dark) .diagnostic-result-item--discarded dt,
:global(.collector-app.dark) .diagnostic-result-item--discarded dd {
  color: #F1CF88;
}

:global(.collector-app.dark) .diagnostic-result-item--article {
  border-color: rgba(87, 174, 139, 0.32);
  background: #1D352F;
}

:global(.collector-app.dark) .diagnostic-result-item--error {
  border-color: rgba(207, 91, 102, 0.36);
  background: #3B252D;
}

:global(.collector-app.dark) .mitm-cert-source-icon,
:global(.collector-app.dark) .diagnostic-result-dialog-icon {
  color: #BFDDFB;
  border-color: rgba(103, 163, 235, 0.26);
  background: rgba(53, 127, 217, 0.2);
}

:global(.collector-app.dark) .diagnostic-result-dialog--success .diagnostic-result-dialog-icon {
  color: #9AD8BE;
  border-color: rgba(73, 174, 132, 0.28);
  background: rgba(47, 141, 130, 0.2);
}

:global(.collector-app.dark) .diagnostic-result-dialog--warning .diagnostic-result-dialog-icon {
  color: #F1CF88;
  border-color: rgba(216, 180, 95, 0.28);
  background: rgba(123, 88, 38, 0.26);
}

:global(.collector-app.dark) .diagnostic-result-dialog--error .diagnostic-result-dialog-icon {
  color: #FFB4BC;
  border-color: rgba(207, 91, 102, 0.3);
  background: rgba(116, 55, 64, 0.3);
}

:global(.collector-app.dark) .mitm-cert-source-card--system .mitm-cert-source-icon {
  color: #a9bfda;
  background: rgba(128, 153, 188, 0.14);
}

:global(.collector-app.dark) .mitm-cert-dialog h3,
:global(.collector-app.dark) .diagnostic-result-dialog h3,
:global(.collector-app.dark) .mitm-cert-source-title strong,
:global(.collector-app.dark) .mitm-cert-item-main strong {
  color: #E3ECF8;
}

:global(.collector-app.dark) .mitm-cert-dialog p,
:global(.collector-app.dark) .diagnostic-result-dialog p,
:global(.collector-app.dark) .mitm-cert-source-title span,
:global(.collector-app.dark) .mitm-cert-meta dt,
:global(.collector-app.dark) .diagnostic-result-item dt,
:global(.collector-app.dark) .mitm-cert-empty {
  color: #9AAEC7;
}

:global(.collector-app.dark) .mitm-cert-meta dd,
:global(.collector-app.dark) .diagnostic-result-item dd {
  color: #D7E2F0;
}

:global(.collector-app.dark) .diagnostic-result-dialog .mitm-cert-dialog-actions {
  border-top-color: #2E405A;
}

:global(.collector-app.dark) .diagnostic-result-dialog :deep(.ant-btn.config-action-button.primary) {
  --config-action-color: #ffffff;
  --config-action-border: #3D7FCC;
  --config-action-bg: #3D7FCC;
  --config-action-hover-color: #ffffff;
  --config-action-hover-border: #4B8DDB;
  --config-action-hover-bg: #4B8DDB;
  --config-action-active-bg: #2E68AE;
  color: #ffffff;
  border-color: #3D7FCC;
  background: #3D7FCC;
  opacity: 1;
}

:global(.collector-app.dark) .diagnostic-result-dialog :deep(.ant-btn.config-action-button.primary:not(:disabled):hover) {
  color: #ffffff;
  border-color: #4B8DDB;
  background: #4B8DDB;
}

:global(.collector-app.dark) .mitm-cert-badge {
  color: #cbd8ea;
  border-color: rgba(128, 153, 188, 0.18);
  background: rgba(17, 27, 44, 0.72);
}

:global(.collector-app.dark) .mitm-cert-badge.success {
  color: #83c7a8;
  border-color: rgba(64, 142, 111, 0.24);
  background: rgba(64, 142, 111, 0.14);
}

:global(.collector-app.dark) .mitm-cert-badge.warning {
  color: #d8b45f;
  border-color: rgba(216, 180, 95, 0.24);
  background: rgba(216, 180, 95, 0.14);
}

:global(.collector-app.dark) .mitm-cert-badge.info {
  color: #8fbded;
  border-color: rgba(45, 117, 214, 0.24);
  background: rgba(45, 117, 214, 0.14);
}

@media (max-width: 760px) {
  .mitm-cert-dialog {
    padding: 16px;
  }

  .mitm-cert-source-heading {
    grid-template-columns: 36px minmax(0, 1fr);
  }

  .mitm-cert-badge {
    grid-column: 1 / -1;
    justify-self: start;
    white-space: normal;
  }

  .mitm-cert-item-main {
    display: grid;
    justify-content: stretch;
  }

  .mitm-cert-item-main span {
    white-space: normal;
    overflow-wrap: anywhere;
  }
}

@media (prefers-reduced-motion: reduce) {
  .mitm-cert-dialog-fade-enter-active,
  .mitm-cert-dialog-fade-leave-active,
  .mitm-cert-dialog-fade-enter-active .mitm-cert-dialog,
  .mitm-cert-dialog-fade-leave-active .mitm-cert-dialog {
    transition: opacity 80ms ease;
  }

  .mitm-cert-dialog-fade-enter-from .mitm-cert-dialog,
  .mitm-cert-dialog-fade-leave-to .mitm-cert-dialog {
    transform: none;
  }
}

</style>
