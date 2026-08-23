<script setup lang="ts">
import AppIcon from '../components/AppIcon.vue'
import ArchiveDistributionChart from '../components/ArchiveDistributionChart.vue'
import { ReloadOutlined, SearchOutlined } from '@ant-design/icons-vue'
import { notification } from 'ant-design-vue'
import { computed, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import type { ArchiveAccountItem, ArchiveArticleItem, ArchiveCacheJob, ArchiveCacheResultItem } from '../bridge/pythonApi'
import {
  cacheArchiveAccount,
  cacheArchiveArticles,
  deleteArchiveAccount,
  deleteArchiveAll,
  deleteArchiveArticles,
  getArchiveCacheJob,
  exportArchiveAccountsToExcel,
  listArchiveAccountArticles,
  listArchiveAccounts,
  openArchiveArticleDirectory,
  openRuntimePath,
} from '../bridge/pythonApi'
import { buildArchiveDistribution, buildArchiveOverview } from '../utils/archiveDistribution'
import type { ArchiveSummaryStat } from '../utils/archiveSummaryStats'

type RecordRow = {
  id: number
  rowIndex: number
  title: string
  link: string
  archiveDir: string
  publishedAt: string
  size: string
}

type FileRow = {
  id: number
  account: string
  createdAt: string
  articleCount: number
  savedCount: number
  failedCount: number
}

type TableStateTone = 'default' | 'error'

type ArchiveTableStateRow = {
  rowKind: 'state'
  tableKey: string
  stateText: string
  stateTone: TableStateTone
}

type ArchiveTablePlaceholderRow = {
  rowKind: 'placeholder'
  tableKey: string
}

type AccountTableDataRow = FileRow & {
  rowKind: 'data'
  tableKey: string
  displayIndex: number
}

type RecordTableDataRow = RecordRow & {
  rowKind: 'data'
  tableKey: string
  displayIndex: number
}

type AccountTableRow = AccountTableDataRow | ArchiveTableStateRow | ArchiveTablePlaceholderRow
type RecordTableRow = RecordTableDataRow | ArchiveTableStateRow | ArchiveTablePlaceholderRow
type ArchiveTableRow = AccountTableRow | RecordTableRow

type ArchiveTableColumn = {
  title: string
  key: string
  dataIndex?: string
  width?: number | string
  align?: 'left' | 'center' | 'right'
  className?: string
  customCell?: (record: ArchiveTableRow) => { colSpan?: number }
}

type DeleteDialogMode = 'records' | 'account' | 'all'

type DeleteDialogState = {
  open: boolean
  mode: DeleteDialogMode
  title: string
  description: string
  summaryItems: Array<{ label: string; value: string }>
  detailTitle: string
  detailItems: string[]
  warning: string
  confirmText: string
  targetArticleIds: number[]
  targetAccount: FileRow | null
  errorMessage: string
}

type CacheNotificationTone = 'success' | 'warning' | 'error'

type CheckboxChangeEvent = {
  target: {
    checked: boolean
  }
}

type CollectDateRangeValue = [string, string] | null

type ArchiveAccountSelectOption = {
  label: string
  value: string
}

const props = defineProps<{
  summaryStats: ArchiveSummaryStat[]
}>()

const emit = defineEmits<{
  navigate: [page: 'home']
}>()

const selectedAccount = ref('')
const selectedCollectStartDate = ref('')
const selectedCollectEndDate = ref('')
const selectedPreviewFile = ref<FileRow | null>(null)
const selectedRecordIndexes = ref<number[]>([])
const fileListRows = ref<FileRow[]>([])
const recordRows = ref<RecordRow[]>([])
const recordTotal = ref(0)
const archiveAccountsLoading = ref(true)
const archiveAccountsError = ref('')
const archiveArticlesLoading = ref(false)
const archiveArticlesError = ref('')
const archiveDeleting = ref(false)
const archiveCaching = ref(false)
const archiveExporting = ref(false)
const activeCacheJobId = ref('')
const cacheJobSnapshot = ref<ArchiveCacheJob | null>(null)
const activeProcessIndex = ref(0)
const cacheProcessRotationPaused = ref(false)
const notifiedCacheArticleIds = ref<number[]>([])
const batchExportDialogOpen = ref(false)
const selectedExportAccountIds = ref<number[]>([])
const deleteDialog = ref<DeleteDialogState>({
  open: false,
  mode: 'records',
  title: '',
  description: '',
  summaryItems: [],
  detailTitle: '',
  detailItems: [],
  warning: '',
  confirmText: '确认删除',
  targetArticleIds: [],
  targetAccount: null,
  errorMessage: '',
})
const fileCurrentPage = ref(1)
const filePageSize = ref(10)
const recordCurrentPage = ref(1)
const recordPageSize = ref(10)
const archiveTablePageSize = 10
let archiveArticlesRequestId = 0
let cachePollingTimer: number | undefined
let cacheProcessRotationTimer: number | undefined

type PagerChangeParams = {
  currentPage: number
  pageSize: number
}

const accountOptions = computed(() => {
  return Array.from(new Set(fileListRows.value.map((file) => file.account)))
})

const accountSelectOptions = computed<ArchiveAccountSelectOption[]>(() => {
  return accountOptions.value.map((account) => ({
    label: account,
    value: account,
  }))
})

const selectedCollectDateRange = computed<CollectDateRangeValue>({
  get: (): CollectDateRangeValue => {
    if (!selectedCollectStartDate.value || !selectedCollectEndDate.value) {
      return null
    }
    return [selectedCollectStartDate.value, selectedCollectEndDate.value]
  },
  set: (value: CollectDateRangeValue) => {
    selectedCollectStartDate.value = value?.[0] ?? ''
    selectedCollectEndDate.value = value?.[1] ?? ''
  },
})

function filterArchiveAccountOption(input: string, option?: ArchiveAccountSelectOption) {
  const keyword = input.trim().toLowerCase()
  if (!keyword) {
    return true
  }
  return (option?.label ?? '').toLowerCase().includes(keyword)
}

const archiveMetricLabels = new Set(['文章数量', '存储占用'])

const archiveMetricIcons: Record<string, string> = {
  文章数量: 'fa-regular fa-file-lines',
  存储占用: 'fa-solid fa-database',
}

const archiveMetrics = computed(() => props.summaryStats.filter((item) => archiveMetricLabels.has(item.label)).map((item) => ({
  ...item,
  icon: archiveMetricIcons[item.label] || 'fa-solid fa-database',
})))

const cacheTaskStatusValue = computed(() => {
  const job = cacheJobSnapshot.value
  if (archiveCaching.value && !job) {
    return '正在准备'
  }
  if (!job) {
    return '当前空闲'
  }
  if (job.status === 'pending') {
    return '正在准备'
  }
  if (job.status === 'running') {
    return `${job.processed} / ${job.requestedTotal}`
  }
  if (job.status === 'done') {
    return `${job.requestedTotal} / ${job.requestedTotal}`
  }
  if (job.status === 'partial_failed') {
    return `${job.processed} / ${job.requestedTotal}`
  }
  return '任务异常'
})

const cacheTaskStatusSummary = computed(() => {
  const job = cacheJobSnapshot.value
  if (archiveCaching.value && !job) {
    return '正在创建缓存任务'
  }
  if (!job) {
    return '暂无缓存任务'
  }
  if (job.status === 'pending') {
    return `等待执行 ${job.requestedTotal} 篇`
  }
  if (job.status === 'running') {
    return `运行 ${job.running} · 排队 ${job.queued} · 跳过 ${job.skipped} · 失败 ${job.failed}`
  }
  if (job.status === 'done') {
    return `已完成 ${job.processed} 篇 · 跳过 ${job.skipped} 篇`
  }
  if (job.status === 'partial_failed') {
    return `已处理 ${job.processed} 篇 · 失败 ${job.failed} 篇`
  }
  return job.message || '缓存任务未正常完成'
})

const cacheTaskTone = computed(() => {
  const status = cacheJobSnapshot.value?.status
  if (status === 'done') {
    return 'green'
  }
  if (status === 'partial_failed') {
    return 'orange'
  }
  if (status === 'failed' || status === 'missing') {
    return 'red'
  }
  return archiveCaching.value ? 'blue' : 'purple'
})

const activeCacheProcesses = computed(() => cacheJobSnapshot.value?.activeProcesses || [])

const visibleActiveProcess = computed(() => {
  const processes = activeCacheProcesses.value
  if (processes.length === 0) {
    return null
  }
  return processes[activeProcessIndex.value % processes.length]
})

const cacheActiveProcessSummary = computed(() => {
  const process = visibleActiveProcess.value
  if (process) {
    return `${process.step} · ${process.elapsedSeconds.toFixed(1)} 秒`
  }
  if (archiveCaching.value && !cacheJobSnapshot.value) {
    return '正在等待缓存子进程启动'
  }
  if (cacheJobSnapshot.value?.status === 'running') {
    return cacheJobSnapshot.value.queued > 0 ? '正在等待下一个子进程' : '正在整理缓存结果'
  }
  return '暂无活跃子进程'
})

const filteredFileRows = computed(() => {
  return fileListRows.value.filter((file) => {
    const collectDate = file.createdAt.slice(0, 10)
    const matchAccount = !selectedAccount.value || file.account === selectedAccount.value
    const matchStartDate = !selectedCollectStartDate.value || collectDate >= selectedCollectStartDate.value
    const matchEndDate = !selectedCollectEndDate.value || collectDate <= selectedCollectEndDate.value

    return matchAccount && matchStartDate && matchEndDate
  })
})

const filePagerRows = computed(() => {
  return filteredFileRows.value
})

const filePagerTotal = computed(() => {
  return filePagerRows.value.length
})

const batchExportRows = computed(() => {
  return fileListRows.value
})

const archiveDistributionData = computed(() => buildArchiveDistribution(fileListRows.value))

const archiveOverview = computed(() => buildArchiveOverview(fileListRows.value))

const recordPagerTotal = computed(() => {
  return selectedPreviewFile.value ? recordTotal.value : 0
})

const visibleFileRows = computed(() => {
  const start = (fileCurrentPage.value - 1) * filePageSize.value
  return filePagerRows.value.slice(start, start + filePageSize.value)
})

const visibleRecordRows = computed(() => {
  return recordRows.value
})

const filePlaceholderRowCount = computed(() => {
  if (archiveAccountsLoading.value || archiveAccountsError.value || visibleFileRows.value.length === 0) {
    return 0
  }
  return Math.max(archiveTablePageSize - visibleFileRows.value.length, 0)
})

const recordPlaceholderRowCount = computed(() => {
  if (archiveArticlesLoading.value || archiveArticlesError.value || visibleRecordRows.value.length === 0) {
    return 0
  }
  return Math.max(archiveTablePageSize - visibleRecordRows.value.length, 0)
})

function makeStateCellAttrs(stateColumnKey: string, span: number) {
  return (columnKey: string) => (record: ArchiveTableRow) => {
    if (record.rowKind !== 'state') {
      return {}
    }

    return columnKey === stateColumnKey ? { colSpan: span } : { colSpan: 0 }
  }
}

const accountStateCellAttrs = makeStateCellAttrs('index', 5)
const recordStateCellAttrs = makeStateCellAttrs('selection', 5)

const accountTableColumns: ArchiveTableColumn[] = [
  { title: '序号', key: 'index', width: 46, align: 'center', customCell: accountStateCellAttrs('index') },
  { title: '公众号', key: 'account', dataIndex: 'account', align: 'center', className: 'account-name-col', customCell: accountStateCellAttrs('account') },
  { title: '采集时间', key: 'createdAt', dataIndex: 'createdAt', width: 104, align: 'center', className: 'account-time-col', customCell: accountStateCellAttrs('createdAt') },
  { title: '数量', key: 'articleCount', dataIndex: 'articleCount', width: 60, align: 'center', className: 'account-size-col', customCell: accountStateCellAttrs('articleCount') },
  { title: '操作', key: 'actions', width: 184, align: 'center', className: 'action-cell account-action-col', customCell: accountStateCellAttrs('actions') },
]

const recordTableColumns: ArchiveTableColumn[] = [
  { title: '', key: 'selection', width: 58, align: 'center', className: 'checkbox-cell record-action-col', customCell: recordStateCellAttrs('selection') },
  { title: '标题', key: 'title', dataIndex: 'title', align: 'left', className: 'record-title', customCell: recordStateCellAttrs('title') },
  { title: '文章发布时间', key: 'publishedAt', dataIndex: 'publishedAt', width: 142, align: 'center', className: 'record-time-col', customCell: recordStateCellAttrs('publishedAt') },
  { title: '大小', key: 'size', dataIndex: 'size', width: 76, align: 'center', className: 'record-size-col', customCell: recordStateCellAttrs('size') },
  { title: '操作', key: 'open', width: 86, align: 'center', className: 'record-open-col', customCell: recordStateCellAttrs('open') },
]

const batchExportTableColumns = [
  { title: '', key: 'selection', width: 54, align: 'center' as const },
  { title: '序号', key: 'index', width: 72, align: 'center' as const },
  { title: '公众号', key: 'account', dataIndex: 'account' },
  { title: '记录数', key: 'articleCount', dataIndex: 'articleCount', width: 116, align: 'center' as const },
]

// Ant Table 的默认空状态不会占满 10 行，这里把状态行和占位行并入数据源，保持分页表格高度稳定。
const accountTableRows = computed<AccountTableRow[]>(() => {
  if (archiveAccountsLoading.value) {
    return [{ rowKind: 'state', tableKey: 'account-loading', stateText: '正在读取数据库中的公众号列表...', stateTone: 'default' }]
  }
  if (archiveAccountsError.value) {
    return [{ rowKind: 'state', tableKey: 'account-error', stateText: archiveAccountsError.value, stateTone: 'error' }]
  }
  if (visibleFileRows.value.length === 0) {
    return [{ rowKind: 'state', tableKey: 'account-empty', stateText: '数据库中暂无公众号记录', stateTone: 'default' }]
  }

  const pageStart = (fileCurrentPage.value - 1) * filePageSize.value
  const rows: AccountTableRow[] = visibleFileRows.value.map((file, index) => ({
    ...file,
    rowKind: 'data',
    tableKey: `account-${file.id}`,
    displayIndex: pageStart + index + 1,
  }))

  for (let index = 0; index < filePlaceholderRowCount.value; index += 1) {
    rows.push({ rowKind: 'placeholder', tableKey: `account-placeholder-${index + 1}` })
  }

  return rows
})

const recordTableRows = computed<RecordTableRow[]>(() => {
  if (archiveArticlesLoading.value) {
    return [{ rowKind: 'state', tableKey: 'record-loading', stateText: '正在读取该公众号的记录详情...', stateTone: 'default' }]
  }
  if (archiveArticlesError.value) {
    return [{ rowKind: 'state', tableKey: 'record-error', stateText: archiveArticlesError.value, stateTone: 'error' }]
  }
  if (visibleRecordRows.value.length === 0) {
    return [{ rowKind: 'state', tableKey: 'record-empty', stateText: '该公众号暂无文章记录', stateTone: 'default' }]
  }

  const rows: RecordTableRow[] = visibleRecordRows.value.map((record, index) => ({
    ...record,
    rowKind: 'data',
    tableKey: `record-${record.id}`,
    displayIndex: index + 1,
  }))

  for (let index = 0; index < recordPlaceholderRowCount.value; index += 1) {
    rows.push({ rowKind: 'placeholder', tableKey: `record-placeholder-${index + 1}` })
  }

  return rows
})

function getArchiveTableRowKey(record: ArchiveTableRow) {
  return record.tableKey
}

function getBatchExportRowKey(record: FileRow) {
  return record.id
}

function getArchiveTableCustomRow(record: ArchiveTableRow) {
  if (record.rowKind === 'placeholder') {
    return { class: 'table-placeholder-row', 'aria-hidden': 'true' }
  }
  if (record.rowKind === 'state') {
    return { class: 'table-state-row' }
  }
  return {}
}

const allRecordIndexes = computed(() => {
  const pageStartIndex = (recordCurrentPage.value - 1) * recordPageSize.value
  return recordRows.value.map((_, index) => pageStartIndex + index)
})

const selectedRecordCount = computed(() => {
  return selectedRecordIndexes.value.length
})

const selectedRecordIds = computed(() => {
  const selectedIndexes = new Set(selectedRecordIndexes.value)
  return recordRows.value.filter((record) => selectedIndexes.has(record.rowIndex)).map((record) => record.id)
})

const isAllRecordsSelected = computed(() => {
  return allRecordIndexes.value.length > 0 && selectedRecordCount.value === allRecordIndexes.value.length
})

const isSomeRecordsSelected = computed(() => {
  return selectedRecordCount.value > 0 && !isAllRecordsSelected.value
})

const selectedExportCount = computed(() => {
  return selectedExportAccountIds.value.length
})

const selectedExportRecordCount = computed(() => {
  const selectedIds = new Set(selectedExportAccountIds.value)
  return batchExportRows.value.reduce((total, file) => {
    return selectedIds.has(file.id) ? total + file.articleCount : total
  }, 0)
})

const isAllExportAccountsSelected = computed(() => {
  return batchExportRows.value.length > 0 && selectedExportCount.value === batchExportRows.value.length
})

const isSomeExportAccountsSelected = computed(() => {
  return selectedExportCount.value > 0 && !isAllExportAccountsSelected.value
})

function formatArchiveAccountRow(item: ArchiveAccountItem): FileRow {
  const latestCollectTime = item.latestCollectTime || item.updatedTime || item.createdTime || '-'

  return {
    id: item.id,
    account: item.accountName || '未知公众号',
    createdAt: formatAccountCollectDate(latestCollectTime),
    articleCount: item.articleCount,
    savedCount: item.savedCount,
    failedCount: item.failedCount,
  }
}

function formatAccountCollectDate(value: string) {
  const matched = value.match(/^\d{4}-\d{2}-\d{2}/)
  return matched ? matched[0] : '-'
}

function formatArchiveArticleRow(item: ArchiveArticleItem, index: number): RecordRow {
  return {
    id: item.id,
    rowIndex: (recordCurrentPage.value - 1) * recordPageSize.value + index,
    title: item.title || '未命名文章',
    link: item.articleLink || '',
    archiveDir: item.archiveDir || '',
    publishedAt: item.publishedArticleTime || '-',
    size: item.sizeLabel || '0 B',
  }
}

async function loadArchiveAccounts() {
  archiveAccountsLoading.value = true
  archiveAccountsError.value = ''

  try {
    const result = await listArchiveAccounts()
    fileListRows.value = result.items.map(formatArchiveAccountRow)
  } catch (error) {
    archiveAccountsError.value = error instanceof Error ? error.message : '读取公众号列表失败'
    fileListRows.value = []
  } finally {
    archiveAccountsLoading.value = false
  }
}

function toggleAllRecords(event: CheckboxChangeEvent) {
  const checked = event.target.checked
  selectedRecordIndexes.value = checked ? [...allRecordIndexes.value] : []
}

function isRecordSelected(rowIndex: number) {
  return selectedRecordIndexes.value.includes(rowIndex)
}

function toggleRecordSelection(rowIndex: number, event: CheckboxChangeEvent) {
  const checked = event.target.checked
  const nextIndexes = new Set(selectedRecordIndexes.value)
  if (checked) {
    nextIndexes.add(rowIndex)
  } else {
    nextIndexes.delete(rowIndex)
  }
  selectedRecordIndexes.value = Array.from(nextIndexes).sort((left, right) => left - right)
}

// 记录详情按后端分页加载；每次打开或翻页时只让后端计算当前页文章目录大小。
async function loadArchiveArticles(file: FileRow, page = recordCurrentPage.value, pageSize = recordPageSize.value) {
  const requestId = ++archiveArticlesRequestId
  recordRows.value = []
  recordTotal.value = 0
  selectedRecordIndexes.value = []
  archiveArticlesLoading.value = true
  archiveArticlesError.value = ''

  try {
    const result = await listArchiveAccountArticles(file.id, page, pageSize)
    if (requestId !== archiveArticlesRequestId) {
      return
    }
    recordCurrentPage.value = result.page
    recordPageSize.value = result.pageSize
    recordTotal.value = result.total
    recordRows.value = result.items.map(formatArchiveArticleRow)
  } catch (error) {
    if (requestId !== archiveArticlesRequestId) {
      return
    }
    archiveArticlesError.value = error instanceof Error ? error.message : '读取记录详情失败'
  } finally {
    if (requestId === archiveArticlesRequestId) {
      archiveArticlesLoading.value = false
    }
  }
}

async function previewFile(file: FileRow) {
  selectedPreviewFile.value = file
  recordCurrentPage.value = 1
  recordPageSize.value = 10
  await loadArchiveArticles(file, 1, recordPageSize.value)
}

// 刷新列表后按公众号 ID 重新绑定当前项，保留用户正在查看的详情页位置。
async function handleRefreshArchiveData() {
  const selectedFileId = selectedPreviewFile.value?.id
  await loadArchiveAccounts()

  if (selectedFileId === undefined) {
    return
  }

  const refreshedFile = fileListRows.value.find((file) => file.id === selectedFileId)
  if (!refreshedFile) {
    selectedPreviewFile.value = null
    recordRows.value = []
    recordTotal.value = 0
    selectedRecordIndexes.value = []
    return
  }

  selectedPreviewFile.value = refreshedFile
  await loadArchiveArticles(refreshedFile, recordCurrentPage.value, recordPageSize.value)
}

function handleFilePageChange({ currentPage, pageSize }: PagerChangeParams) {
  fileCurrentPage.value = currentPage
  filePageSize.value = pageSize
}

function handleAntFilePageChange(page: number, pageSize: number) {
  handleFilePageChange({ currentPage: page, pageSize })
}

function handleRecordPageChange({ currentPage, pageSize }: PagerChangeParams) {
  recordCurrentPage.value = currentPage
  recordPageSize.value = pageSize
  if (selectedPreviewFile.value) {
    void loadArchiveArticles(selectedPreviewFile.value, currentPage, pageSize)
  }
}

function handleAntRecordPageChange(page: number, pageSize: number) {
  handleRecordPageChange({ currentPage: page, pageSize })
}

async function refreshArchiveViewAfterDelete() {
  await loadArchiveAccounts()
  if (!selectedPreviewFile.value) {
    return
  }

  const exists = fileListRows.value.some((file) => file.id === selectedPreviewFile.value?.id)
  if (!exists) {
    selectedPreviewFile.value = null
    recordRows.value = []
    recordTotal.value = 0
    selectedRecordIndexes.value = []
    return
  }

  const nextPage = Math.min(recordCurrentPage.value, Math.max(1, Math.ceil(recordPagerTotal.value / recordPageSize.value)))
  await loadArchiveArticles(selectedPreviewFile.value, nextPage, recordPageSize.value)
}

function formatDeleteFailureMessage(result: Awaited<ReturnType<typeof deleteArchiveArticles>>) {
  if (result.ok) {
    return ''
  }
  const firstFailure = result.failures[0]
  return firstFailure ? `部分本地目录删除失败：${firstFailure.path}` : '删除时存在未完成项，请查看后端日志'
}

function showCacheNotification(message: string, tone: CacheNotificationTone = 'success') {
  const options = {
    description: message,
    placement: 'bottomRight' as const,
    duration: 2.4,
  }

  if (tone === 'error') {
    notification.error({ message: '操作失败', ...options })
    return
  }
  if (tone === 'warning') {
    notification.warning({ message: '请注意', ...options })
    return
  }
  notification.success({ message: '操作成功', ...options })
}

async function handleOpenStorageDirectory() {
  try {
    const result = await openRuntimePath('storageDir')
    showCacheNotification(result.message || '已打开数据归档目录。', result.ok ? 'success' : 'error')
  } catch (error) {
    showCacheNotification(error instanceof Error ? error.message : '打开数据归档目录失败', 'error')
  }
}

function openArticleLink(record: RecordRow) {
  if (!record.link) {
    return
  }
  window.open(record.link, '_blank', 'noopener,noreferrer')
}

async function openRecordArchiveDirectory(record: RecordRow) {
  try {
    const result = await openArchiveArticleDirectory(record.id)
    showCacheNotification(result.message || '已打开文章归档目录。', result.ok ? 'success' : 'error')
  } catch (error) {
    showCacheNotification(error instanceof Error ? error.message : '打开文章归档目录失败', 'error')
  }
}

function clearCachePolling() {
  if (cachePollingTimer !== undefined) {
    window.clearInterval(cachePollingTimer)
    cachePollingTimer = undefined
  }
}

function clearCacheProcessRotation() {
  if (cacheProcessRotationTimer !== undefined) {
    window.clearInterval(cacheProcessRotationTimer)
    cacheProcessRotationTimer = undefined
  }
}

function startCacheProcessRotation() {
  clearCacheProcessRotation()
  cacheProcessRotationTimer = window.setInterval(() => {
    if (cacheProcessRotationPaused.value || activeCacheProcesses.value.length <= 1) {
      return
    }
    activeProcessIndex.value = (activeProcessIndex.value + 1) % activeCacheProcesses.value.length
  }, 2500)
}

function pauseCacheProcessRotation() {
  cacheProcessRotationPaused.value = true
}

function resumeCacheProcessRotation() {
  cacheProcessRotationPaused.value = false
}

async function refreshArchiveViewAfterCache() {
  await loadArchiveAccounts()
  if (selectedPreviewFile.value) {
    await loadArchiveArticles(selectedPreviewFile.value, recordCurrentPage.value, recordPageSize.value)
  }
}

function notifyCacheResults(results: ArchiveCacheResultItem[]) {
  const notified = new Set(notifiedCacheArticleIds.value)
  for (const item of results) {
    if (notified.has(item.articleId)) {
      continue
    }
    notified.add(item.articleId)
    if (item.ok) {
      showCacheNotification(`${item.articleTitle} 已缓存`, 'success')
    } else {
      showCacheNotification(`${item.articleTitle} 缓存失败`, 'error')
    }
  }
  notifiedCacheArticleIds.value = Array.from(notified)
}

async function handleCacheJobSnapshot(job: ArchiveCacheJob) {
  cacheJobSnapshot.value = job
  notifyCacheResults(job.results)
  const finished = ['done', 'partial_failed', 'failed', 'missing'].includes(job.status)
  if (!finished) {
    return
  }

  clearCachePolling()
  archiveCaching.value = false
  activeCacheJobId.value = ''
  if (job.status === 'done' && job.total === 0) {
    showCacheNotification(job.message || '没有可缓存的文章', 'warning')
  } else if (job.status === 'done' && job.skipped > 0) {
    showCacheNotification(job.message || `新增缓存 ${job.finished} 篇，跳过已有缓存 ${job.skipped} 篇。`, 'success')
  } else if (job.status === 'partial_failed') {
    showCacheNotification(job.message || '部分文章缓存失败', 'warning')
  } else if (job.status === 'failed' || job.status === 'missing') {
    showCacheNotification(job.message || '缓存任务失败', 'error')
  }
  await refreshArchiveViewAfterCache()
}

async function pollCacheJob(jobId: string) {
  try {
    const job = await getArchiveCacheJob(jobId)
    await handleCacheJobSnapshot(job)
  } catch (error) {
    clearCachePolling()
    archiveCaching.value = false
    activeCacheJobId.value = ''
    showCacheNotification(error instanceof Error ? error.message : '读取缓存任务状态失败', 'error')
  }
}

function startCachePolling(jobId: string) {
  clearCachePolling()
  activeCacheJobId.value = jobId
  cachePollingTimer = window.setInterval(() => {
    void pollCacheJob(jobId)
  }, 1000)
  void pollCacheJob(jobId)
}

async function startArchiveCacheJob(jobPromise: Promise<ArchiveCacheJob>) {
  if (archiveCaching.value) {
    showCacheNotification('已有缓存任务正在执行', 'warning')
    return
  }
  archiveCaching.value = true
  cacheJobSnapshot.value = null
  activeProcessIndex.value = 0
  notifiedCacheArticleIds.value = []
  try {
    const job = await jobPromise
    await handleCacheJobSnapshot(job)
    if (job.total === 0) {
      return
    }
    if (!['done', 'partial_failed', 'failed', 'missing'].includes(job.status)) {
      startCachePolling(job.jobId)
    }
  } catch (error) {
    archiveCaching.value = false
    showCacheNotification(error instanceof Error ? error.message : '创建缓存任务失败', 'error')
  }
}

async function cacheSelectedRecords() {
  const articleIds = selectedRecordIds.value
  if (articleIds.length === 0) {
    showCacheNotification('请先选择要缓存的文章', 'warning')
    return
  }
  await startArchiveCacheJob(cacheArchiveArticles(articleIds))
}

async function cacheAccountArticles(file: FileRow) {
  await startArchiveCacheJob(cacheArchiveAccount(file.id))
}

function openBatchExportDialog() {
  // 批量导出默认全选已有公众号，用户可以在弹窗内再缩小范围。
  selectedExportAccountIds.value = batchExportRows.value.map((file) => file.id)
  batchExportDialogOpen.value = true
}

function closeBatchExportDialog() {
  if (archiveExporting.value) {
    return
  }
  batchExportDialogOpen.value = false
}

function toggleAllExportAccounts(event: CheckboxChangeEvent) {
  const checked = event.target.checked
  selectedExportAccountIds.value = checked ? batchExportRows.value.map((file) => file.id) : []
}

function isExportAccountSelected(accountId: number) {
  return selectedExportAccountIds.value.includes(accountId)
}

function toggleExportAccount(accountId: number, event: CheckboxChangeEvent) {
  const checked = event.target.checked
  const nextIds = new Set(selectedExportAccountIds.value)
  if (checked) {
    nextIds.add(accountId)
  } else {
    nextIds.delete(accountId)
  }
  selectedExportAccountIds.value = Array.from(nextIds)
}

async function handleBatchExportExcel() {
  if (selectedExportAccountIds.value.length === 0) {
    showCacheNotification('请先选择要导出的公众号', 'warning')
    return
  }
  archiveExporting.value = true
  try {
    const result = await exportArchiveAccountsToExcel([...selectedExportAccountIds.value])
    if (!result.ok) {
      const message = result.message || '导出 Excel 失败'
      showCacheNotification(message, 'error')
      return
    }
    batchExportDialogOpen.value = false
    showCacheNotification(
      result.message || `已导出 ${result.exportedFileCount} 个 Excel 文件，共 ${result.totalRowCount} 条记录`,
      result.status === 'partial-failed' ? 'warning' : 'success',
    )
  } catch (error) {
    const message = error instanceof Error ? error.message : '导出 Excel 失败'
    showCacheNotification(message, message.includes('取消') ? 'warning' : 'error')
  } finally {
    archiveExporting.value = false
  }
}

function closeDeleteDialog() {
  if (archiveDeleting.value) {
    return
  }
  deleteDialog.value.open = false
  deleteDialog.value.errorMessage = ''
}

function openDeleteSelectedRecordsDialog() {
  if (!selectedPreviewFile.value || archiveDeleting.value) {
    return
  }
  const articleIds = selectedRecordIds.value
  if (articleIds.length === 0) {
    deleteDialog.value = {
      ...deleteDialog.value,
      open: true,
      mode: 'records',
      title: '请先选择记录',
      description: '需要先在记录详情表格中勾选一条或多条文章记录，才能执行删除。',
      summaryItems: [
        { label: '当前公众号', value: selectedPreviewFile.value.account },
        { label: '已选择', value: '0 条' },
      ],
      detailTitle: '',
      detailItems: [],
      warning: '未选择记录时不会删除任何数据库记录或本地归档目录。',
      confirmText: '我知道了',
      targetArticleIds: [],
      targetAccount: null,
      errorMessage: '',
    }
    return
  }

  const selectedIndexes = new Set(selectedRecordIndexes.value)
  const selectedRows = recordRows.value.filter((record) => selectedIndexes.has(record.rowIndex))
  deleteDialog.value = {
    open: true,
    mode: 'records',
    title: `删除选中的 ${articleIds.length} 条记录`,
    description: `将删除公众号「${selectedPreviewFile.value.account}」下选中的文章记录。`,
    summaryItems: [
      { label: '当前公众号', value: selectedPreviewFile.value.account },
      { label: '文章记录', value: `${articleIds.length} 条` },
      { label: '数据库影响', value: '删除 awa_public_articles 对应行' },
      { label: '本地存储', value: '删除文章目录及 _1 等重复目录' },
    ],
    detailTitle: '将删除的记录',
    detailItems: selectedRows.map((record) => `#${record.rowIndex + 1} ${record.title}`).slice(0, 5),
    warning: '删除后会同时清理本地归档目录；同一篇文章的重复目录会一并删除，无法从页面恢复。',
    confirmText: '确认删除记录',
    targetArticleIds: articleIds,
    targetAccount: null,
    errorMessage: '',
  }
}

async function executeDeleteSelectedRecords(articleIds: number[]) {
  archiveDeleting.value = true
  try {
    const result = await deleteArchiveArticles(articleIds)
    const failureMessage = formatDeleteFailureMessage(result)
    if (failureMessage) {
      deleteDialog.value.errorMessage = failureMessage
      return
    }
    deleteDialog.value.open = false
    await refreshArchiveViewAfterDelete()
  } catch (error) {
    deleteDialog.value.errorMessage = error instanceof Error ? error.message : '删除记录失败'
  } finally {
    archiveDeleting.value = false
  }
}

function openDeleteAccountDialog(file: FileRow) {
  if (archiveDeleting.value) {
    return
  }

  deleteDialog.value = {
    open: true,
    mode: 'account',
    title: `删除公众号「${file.account}」记录`,
    description: '将先删除该公众号下全部文章记录及本地归档，再删除公众号记录。',
    summaryItems: [
      { label: '公众号', value: file.account },
      { label: '文章记录', value: `${file.articleCount} 条` },
      { label: '成功记录', value: `${file.savedCount} 条` },
      { label: '失败记录', value: `${file.failedCount} 条` },
    ],
    detailTitle: '删除顺序',
    detailItems: [
      '逐篇删除 awa_public_articles 中的文章记录',
      '删除每篇文章本地目录及 _1 等重复目录',
      '最后删除 awa_public_accounts 中的公众号记录',
    ],
    warning: '这会删除该公众号下所有数据档案内容，包括本地存储文件；删除后无法从页面恢复。',
    confirmText: '确认删除公众号记录',
    targetArticleIds: [],
    targetAccount: file,
    errorMessage: '',
  }
}

async function executeDeleteAccount(file: FileRow) {
  archiveDeleting.value = true
  try {
    const result = await deleteArchiveAccount(file.id)
    const failureMessage = formatDeleteFailureMessage(result)
    if (failureMessage) {
      deleteDialog.value.errorMessage = failureMessage
      return
    }
    deleteDialog.value.open = false
    if (selectedPreviewFile.value?.id === file.id) {
      selectedPreviewFile.value = null
      recordRows.value = []
      recordTotal.value = 0
      selectedRecordIndexes.value = []
    }
    await loadArchiveAccounts()
  } catch (error) {
    deleteDialog.value.errorMessage = error instanceof Error ? error.message : '删除公众号失败'
  } finally {
    archiveDeleting.value = false
  }
}

function openDeleteAllArchivesDialog() {
  if (archiveDeleting.value) {
    return
  }

  const articleTotal = fileListRows.value.reduce((total, file) => total + file.articleCount, 0)
  deleteDialog.value = {
    open: true,
    mode: 'all',
    title: '删除全部数据档案',
    description: '将遍历公众号列表，删除全部文章记录、公众号记录和对应本地归档目录。',
    summaryItems: [
      { label: '公众号', value: `${fileListRows.value.length} 个` },
      { label: '文章记录', value: `${articleTotal} 条` },
      { label: '数据库影响', value: '清空 awa_public_articles / awa_public_accounts' },
      { label: '本地存储', value: '清理全部匹配归档目录' },
    ],
    detailTitle: '将删除的公众号',
    detailItems: fileListRows.value.map((file) => `${file.account}（${file.articleCount} 条）`).slice(0, 5),
    warning: '这是最高风险操作，会删除当前数据库中全部数据档案及匹配的本地存储目录，无法从页面恢复。',
    confirmText: '确认全部删除',
    targetArticleIds: [],
    targetAccount: null,
    errorMessage: '',
  }
}

async function executeDeleteAllArchives() {
  archiveDeleting.value = true
  try {
    const result = await deleteArchiveAll()
    const failureMessage = formatDeleteFailureMessage(result)
    if (failureMessage) {
      deleteDialog.value.errorMessage = failureMessage
      return
    }
    deleteDialog.value.open = false
    selectedPreviewFile.value = null
    selectedRecordIndexes.value = []
    recordRows.value = []
    recordTotal.value = 0
    await loadArchiveAccounts()
  } catch (error) {
    deleteDialog.value.errorMessage = error instanceof Error ? error.message : '全部删除失败'
  } finally {
    archiveDeleting.value = false
  }
}

async function confirmDeleteDialog() {
  if (archiveDeleting.value) {
    return
  }
  if (deleteDialog.value.targetArticleIds.length === 0 && deleteDialog.value.mode === 'records') {
    closeDeleteDialog()
    return
  }
  if (deleteDialog.value.mode === 'records') {
    await executeDeleteSelectedRecords(deleteDialog.value.targetArticleIds)
    return
  }
  if (deleteDialog.value.mode === 'account' && deleteDialog.value.targetAccount) {
    await executeDeleteAccount(deleteDialog.value.targetAccount)
    return
  }
  if (deleteDialog.value.mode === 'all') {
    await executeDeleteAllArchives()
  }
}

watch([selectedAccount, selectedCollectStartDate, selectedCollectEndDate], () => {
  fileCurrentPage.value = 1
  selectedPreviewFile.value = null
  selectedRecordIndexes.value = []
  recordRows.value = []
  recordTotal.value = 0
  archiveArticlesError.value = ''
  archiveArticlesLoading.value = false
})

watch(
  () => activeCacheProcesses.value.map((item) => item.articleId).join(','),
  () => {
    activeProcessIndex.value = 0
  },
)

async function refreshOnActivated() {
  await handleRefreshArchiveData()
}

defineExpose({
  refreshOnActivated,
})

onMounted(() => {
  startCacheProcessRotation()
})

onBeforeUnmount(() => {
  clearCachePolling()
  clearCacheProcessRotation()
})

</script>

<template>
  <section class="management-page files-page" aria-label="数据文件">
    <div class="metric-grid files-metrics config-summary-metrics">
      <article v-for="item in archiveMetrics" :key="item.label" class="metric-card page-panel">
        <span :class="['metric-icon', item.tone]">
          <AppIcon :icon="item.icon" />
        </span>
        <div class="metric-body">
          <span>{{ item.label }}</span>
          <strong :class="item.tone">{{ item.value }}</strong>
        </div>
      </article>

      <article class="cache-task-status-card metric-card page-panel" aria-live="polite">
        <span :class="['metric-icon', cacheTaskTone]">
          <AppIcon icon="fa-solid fa-download" />
        </span>
        <div class="cache-status-body">
          <div class="cache-status-head">
            <span>缓存任务</span>
            <strong :class="['cache-status-value', cacheTaskTone]">{{ cacheTaskStatusValue }}</strong>
          </div>
          <small :title="cacheTaskStatusSummary">{{ cacheTaskStatusSummary }}</small>
        </div>
      </article>

      <article
        class="cache-active-process-card metric-card page-panel"
        aria-live="polite"
        @mouseenter="pauseCacheProcessRotation"
        @mouseleave="resumeCacheProcessRotation"
      >
        <span class="metric-icon purple">
          <AppIcon icon="fa-solid fa-gears" />
        </span>
        <div class="cache-status-body">
          <div class="cache-status-head">
            <span>活跃子进程</span>
            <strong class="cache-status-value purple">
              {{ activeCacheProcesses.length }} / {{ cacheJobSnapshot?.concurrency || 0 }}
            </strong>
          </div>
          <Transition name="cache-process-slide" mode="out-in">
            <div
              :key="visibleActiveProcess?.articleId || 'cache-process-idle'"
              class="cache-process-content"
            >
              <span
                v-if="visibleActiveProcess"
                class="cache-process-title"
                :title="visibleActiveProcess.articleTitle"
              >
                {{ visibleActiveProcess.articleTitle }}
              </span>
              <small :title="cacheActiveProcessSummary">{{ cacheActiveProcessSummary }}</small>
            </div>
          </Transition>
        </div>
      </article>
    </div>

    <section class="file-list page-panel" aria-label="公众号列表">
      <div class="file-list-header">
        <h2 class="section-heading">
          <AppIcon icon="fa-regular fa-rectangle-list" />
          公众号列表
        </h2>

        <div class="filters-row file-filters">
          <AButton
            class="file-refresh-button"
            html-type="button"
            :loading="archiveAccountsLoading || archiveArticlesLoading"
            @click="handleRefreshArchiveData"
          >
            <ReloadOutlined />
            刷新
          </AButton>
          <ASelect
            v-model:value="selectedAccount"
            class="file-ant-control file-ant-select account-select-trigger"
            show-search
            allow-clear
            placeholder="选择公众号"
            :options="accountSelectOptions"
            :filter-option="filterArchiveAccountOption"
            :dropdown-match-select-width="false"
            popup-class-name="file-account-filter-dropdown"
            aria-label="选择公众号"
          >
            <template #suffixIcon>
              <SearchOutlined class="file-search-icon" />
            </template>
          </ASelect>
          <ARangePicker
            v-model:value="selectedCollectDateRange"
            class="file-date-picker file-date-range-picker"
            :allow-clear="true"
            format="YYYY-MM-DD"
            value-format="YYYY-MM-DD"
            separator=" ~ "
            popup-class-name="file-date-picker-panel"
            :placeholder="['采集起始日期', '采集结束日期']"
            aria-label="采集起始日期和采集结束日期"
          />
        </div>
      </div>

      <div class="table-wrap">
        <ATable
          class="archive-ant-table account-table"
          :columns="accountTableColumns"
          :data-source="accountTableRows"
          :pagination="false"
          :row-key="getArchiveTableRowKey"
          :custom-row="getArchiveTableCustomRow"
          :locale="{ emptyText: null }"
          size="small"
          table-layout="fixed"
        >
          <template #bodyCell="{ column, record }">
            <span
              v-if="record.rowKind === 'state' && column.key === 'index'"
              :class="['table-state-cell', { error: record.stateTone === 'error' }]"
            >
              {{ record.stateText }}
            </span>
            <template v-else-if="record.rowKind === 'data'">
              <template v-if="column.key === 'index'">
                {{ record.displayIndex }}
              </template>
              <span v-else-if="column.key === 'account'" class="account-name-cell" :title="record.account">
                {{ record.account }}
              </span>
              <template v-else-if="column.key === 'createdAt'">
                {{ record.createdAt }}
              </template>
              <template v-else-if="column.key === 'articleCount'">
                {{ record.articleCount }}
              </template>
              <span v-else-if="column.key === 'actions'" class="table-actions">
                <AButton class="text-link" type="link" html-type="button" @click="previewFile(record)">预览</AButton>
                <AButton
                  class="text-link"
                  type="link"
                  html-type="button"
                  :disabled="archiveCaching"
                  @click="cacheAccountArticles(record)"
                >
                  一键缓存
                </AButton>
                <AButton
                  class="text-link danger"
                  type="link"
                  html-type="button"
                  :disabled="archiveDeleting || archiveCaching"
                  @click="openDeleteAccountDialog(record)"
                >
                  删除
                </AButton>
              </span>
            </template>
          </template>
        </ATable>
      </div>

      <div class="pagination-bar">
        <span class="account-pagination-total">共 {{ filePagerTotal }} 条</span>
        <div class="pagination-controls">
          <APagination
            v-model:current="fileCurrentPage"
            v-model:page-size="filePageSize"
            class="file-ant-pager"
            size="small"
            :total="filePagerTotal"
            :page-size-options="[10]"
            :show-size-changer="false"
            :show-less-items="false"
            @change="handleAntFilePageChange"
          />
          <span class="fixed-page-size" aria-label="每页条数">{{ filePageSize }} 条/页</span>
        </div>
      </div>
    </section>

    <aside class="file-side">
      <section class="record-detail page-panel">
        <img class="panel-corner-art detail-art" src="/assets/watercolor-leaf-branch-a.png" alt="" />
        <div class="record-detail-header">
          <h2 class="section-heading">
            <AppIcon icon="fa-regular fa-file-lines" />
            记录详情
          </h2>

          <div v-if="selectedPreviewFile" class="record-actions" aria-label="记录详情操作">
            <div class="record-selection-summary" aria-live="polite">
              <span class="record-account-name">{{ selectedPreviewFile.account }}</span>
              <span>已选择 {{ selectedRecordCount }} 条</span>
            </div>
            <div class="record-action-buttons">
              <AButton
                class="action-button primary"
                html-type="button"
                :disabled="selectedRecordCount === 0 || archiveCaching || archiveArticlesLoading"
                @click="cacheSelectedRecords"
              >
                <AppIcon icon="fa-solid fa-download" />
                {{ archiveCaching ? '缓存中' : '缓存' }}
              </AButton>
              <AButton
                class="action-button danger"
                html-type="button"
                :disabled="selectedRecordCount === 0 || archiveDeleting || archiveCaching || archiveArticlesLoading"
                @click="openDeleteSelectedRecordsDialog"
              >
                <AppIcon icon="fa-regular fa-trash-can" />
                删除
              </AButton>
            </div>
          </div>
        </div>

        <template v-if="selectedPreviewFile">

          <div class="table-wrap record-table-wrap">
            <ATable
              class="archive-ant-table record-table"
              :columns="recordTableColumns"
              :data-source="recordTableRows"
              :pagination="false"
              :row-key="getArchiveTableRowKey"
              :custom-row="getArchiveTableCustomRow"
              :locale="{ emptyText: null }"
              size="small"
              table-layout="fixed"
            >
              <template #headerCell="{ column }">
                <ACheckbox
                  v-if="column.key === 'selection'"
                  class="record-checkbox"
                  :checked="isAllRecordsSelected"
                  :indeterminate="isSomeRecordsSelected"
                  :disabled="recordPagerTotal === 0 || archiveArticlesLoading || Boolean(archiveArticlesError)"
                  aria-label="全选记录详情列表"
                  @change="toggleAllRecords"
                />
                <template v-else>{{ column.title }}</template>
              </template>

              <template #bodyCell="{ column, record }">
                <span
                  v-if="record.rowKind === 'state' && column.key === 'selection'"
                  :class="['table-state-cell', { error: record.stateTone === 'error' }]"
                >
                  {{ record.stateText }}
                </span>
                <template v-else-if="record.rowKind === 'data'">
                  <ACheckbox
                    v-if="column.key === 'selection'"
                    class="record-checkbox"
                    :checked="isRecordSelected(record.rowIndex)"
                    :aria-label="`勾选第 ${record.displayIndex} 条记录`"
                    @change="toggleRecordSelection(record.rowIndex, $event)"
                  />
                  <AButton
                    v-else-if="column.key === 'title'"
                    class="record-title-link"
                    type="link"
                    html-type="button"
                    :disabled="!record.link"
                    :title="record.link ? `${record.title}\n${record.link}` : record.title"
                    @click="openArticleLink(record)"
                  >
                    {{ record.title }}
                  </AButton>
                  <template v-else-if="column.key === 'publishedAt'">
                    {{ record.publishedAt }}
                  </template>
                  <template v-else-if="column.key === 'size'">
                    {{ record.size }}
                  </template>
                  <span v-else-if="column.key === 'open'" class="record-open-cell">
                    <AButton
                      class="text-link"
                      type="link"
                      html-type="button"
                      @click="openRecordArchiveDirectory(record)"
                    >
                      打开目录
                    </AButton>
                  </span>
                </template>
              </template>
            </ATable>
          </div>

          <div class="record-pagination" aria-label="记录详情分页">
            <span class="record-pagination-total">共 {{ recordPagerTotal }} 条</span>
            <div class="pagination-controls">
              <APagination
                v-model:current="recordCurrentPage"
                v-model:page-size="recordPageSize"
                class="file-ant-pager record-ant-pager"
                size="small"
                :total="recordPagerTotal"
                :page-size-options="[10]"
                :show-size-changer="false"
                :show-less-items="false"
                @change="handleAntRecordPageChange"
              />
              <span class="fixed-page-size" aria-label="记录每页条数">{{ recordPageSize }} 条/页</span>
            </div>
          </div>
        </template>
        <div v-else class="record-empty-shell" aria-live="polite">
          <div v-if="archiveAccountsLoading" class="record-empty record-empty-loading">
            <ASkeleton
              active
              :title="{ width: '46%' }"
              :paragraph="{ rows: 6, width: ['68%', '92%', '84%', '76%', '64%', '54%'] }"
            />
          </div>

          <AResult
            v-else-if="archiveAccountsError"
            class="record-empty record-empty-result"
            status="warning"
            title="读取归档数据失败"
            :sub-title="archiveAccountsError"
          >
            <template #extra>
              <AButton class="record-empty-primary" html-type="button" @click="handleRefreshArchiveData">
                重新加载
              </AButton>
            </template>
          </AResult>

          <div v-else-if="archiveDistributionData.length === 0" class="record-empty record-empty-first-run">
            <span class="record-empty-icon">
              <AppIcon icon="fa-regular fa-folder-open" />
            </span>
            <h3>暂无归档数据</h3>
            <p>完成一次文章采集后，这里将展示公众号分布、文章数量和归档概览。</p>
            <AButton class="record-empty-primary" html-type="button" @click="emit('navigate', 'home')">
              前往主服务
            </AButton>
          </div>

          <div v-else class="record-empty record-overview-empty">
            <div class="archive-overview-guide" role="note">
              <span class="archive-overview-guide-icon">
                <AppIcon icon="fa-regular fa-hand-pointer" />
              </span>
              <div class="archive-overview-guide-copy">
                <strong>选择左侧公众号查看记录详情</strong>
                <span>点击左侧列表中的「预览」后，可查看文章标题、发布时间、归档大小和操作入口。</span>
              </div>
              <div class="archive-overview-guide-steps" aria-label="查看记录详情步骤">
                <span>选择公众号</span>
                <span>点击预览</span>
                <span>查看详情</span>
              </div>
            </div>
            <div class="archive-overview-chart-wrap">
              <ArchiveDistributionChart
                :data="archiveDistributionData"
                :total="archiveOverview.articleCount"
              />
            </div>
            <div class="archive-overview-copy">
              <span class="archive-overview-label">归档内容概览</span>
              <h3>已归档 {{ archiveOverview.articleCount }} 篇文章</h3>
              <dl class="archive-overview-stats">
                <div>
                  <dt>记录最多</dt>
                  <dd>{{ archiveOverview.topAccountName }} · {{ archiveOverview.topAccountArticleCount }} 篇</dd>
                </div>
                <div>
                  <dt>最近采集</dt>
                  <dd>{{ archiveOverview.latestCollectDate }}</dd>
                </div>
              </dl>
              <div class="archive-overview-legend" aria-label="公众号文章数量占比">
                <span v-for="item in archiveDistributionData" :key="item.name">
                  <i :style="{ backgroundColor: item.color }"></i>
                  {{ item.name }} {{ item.value }}
                </span>
              </div>
            </div>
          </div>
        </div>
      </section>
    </aside>

    <section class="bottom-quick-actions page-panel" aria-label="快速操作">
      <div class="quick-actions-head">
        <h2 class="section-heading">
          <AppIcon icon="fa-solid fa-bolt" />
          快速操作
        </h2>
      </div>

      <div class="bottom-quick-grid">
        <AButton class="action-button ghost" html-type="button" @click="handleOpenStorageDirectory">
          <AppIcon icon="fa-regular fa-folder-open" />
          打开目录
        </AButton>
        <AButton
          class="action-button purple"
          html-type="button"
          :disabled="fileListRows.length === 0"
          @click="openBatchExportDialog"
        >
          <AppIcon icon="fa-regular fa-file-excel" />
          批量导出
        </AButton>
        <AButton
          class="action-button danger"
          html-type="button"
          :disabled="archiveDeleting || archiveCaching || fileListRows.length === 0"
          @click="openDeleteAllArchivesDialog"
        >
          <AppIcon icon="fa-regular fa-trash-can" />
          全部删除
        </AButton>
      </div>
    </section>

    <AModal
      v-model:open="batchExportDialogOpen"
      class="batch-export-modal"
      :width="680"
      :closable="!archiveExporting"
      :keyboard="!archiveExporting"
      :mask-closable="!archiveExporting"
      centered
      @cancel="closeBatchExportDialog"
    >
      <template #title>
        <div class="batch-export-head">
          <span class="batch-export-icon" aria-hidden="true">
            <AppIcon icon="fa-regular fa-file-excel" />
          </span>
          <div>
            <h3 id="batch-export-dialog-title">已有公众号列表</h3>
            <p>选择需要导出的公众号，下一步将把对应记录写入 Excel 文件。</p>
          </div>
        </div>
      </template>

      <div class="batch-export-meta" aria-live="polite">
        已选择 {{ selectedExportCount }} 个公众号，共 {{ selectedExportRecordCount }} 条记录
      </div>

      <div class="batch-export-table-wrap">
        <ATable
          class="archive-ant-table batch-export-ant-table"
          :columns="batchExportTableColumns"
          :data-source="batchExportRows"
          :pagination="false"
          :row-key="getBatchExportRowKey"
          :locale="{ emptyText: '数据库中暂无公众号记录' }"
          size="small"
          table-layout="fixed"
        >
          <template #headerCell="{ column }">
            <ACheckbox
              v-if="column.key === 'selection'"
              class="record-checkbox"
              :checked="isAllExportAccountsSelected"
              :indeterminate="isSomeExportAccountsSelected"
              :disabled="archiveExporting"
              aria-label="选择全部公众号"
              @change="toggleAllExportAccounts"
            />
            <template v-else>{{ column.title }}</template>
          </template>

          <template #bodyCell="{ column, record, index }">
            <ACheckbox
              v-if="column.key === 'selection'"
              class="record-checkbox"
              :checked="isExportAccountSelected(record.id)"
              :disabled="archiveExporting"
              :aria-label="`选择公众号 ${record.account}`"
              @change="toggleExportAccount(record.id, $event)"
            />
            <template v-else-if="column.key === 'index'">{{ index + 1 }}</template>
            <span v-else-if="column.key === 'account'" class="export-account-cell" :title="record.account">
              {{ record.account }}
            </span>
            <template v-else-if="column.key === 'articleCount'">{{ record.articleCount }} 条</template>
          </template>
        </ATable>
      </div>

      <template #footer>
        <div class="batch-export-actions">
          <AButton class="action-button ghost" html-type="button" :disabled="archiveExporting" @click="closeBatchExportDialog">
            取消
          </AButton>
          <AButton
            class="action-button success"
            type="primary"
            html-type="button"
            :disabled="selectedExportCount === 0"
            :loading="archiveExporting"
            @click="handleBatchExportExcel"
          >
            <AppIcon v-if="!archiveExporting" icon="fa-regular fa-file-excel" />
            {{ archiveExporting ? '正在导出...' : '导出为excel' }}
          </AButton>
        </div>
      </template>
    </AModal>

    <AModal
      :open="deleteDialog.open"
      class="archive-delete-modal"
      :title="deleteDialog.title"
      :closable="!archiveDeleting"
      :keyboard="!archiveDeleting"
      :mask-closable="!archiveDeleting"
      :confirm-loading="archiveDeleting"
      :ok-text="deleteDialog.confirmText"
      :ok-button-props="{ danger: deleteDialog.confirmText !== '我知道了' }"
      cancel-text="取消"
      centered
      @cancel="closeDeleteDialog"
      @ok="confirmDeleteDialog"
    >
      <div class="archive-delete-content">
        <div class="archive-delete-intro">
          <span class="archive-delete-icon" aria-hidden="true">
            <AppIcon icon="fa-regular fa-trash-can" />
          </span>
          <p>{{ deleteDialog.description }}</p>
        </div>

        <dl v-if="deleteDialog.summaryItems.length" class="archive-delete-summary">
          <div v-for="item in deleteDialog.summaryItems" :key="item.label">
            <dt>{{ item.label }}</dt>
            <dd>{{ item.value }}</dd>
          </div>
        </dl>

        <div v-if="deleteDialog.detailItems.length" class="archive-delete-detail">
          <strong>{{ deleteDialog.detailTitle }}</strong>
          <ul>
            <li v-for="item in deleteDialog.detailItems" :key="item">{{ item }}</li>
          </ul>
        </div>

        <AAlert
          v-if="deleteDialog.warning"
          class="archive-delete-alert"
          type="warning"
          show-icon
          :message="deleteDialog.warning"
        />
        <AAlert
          v-if="deleteDialog.errorMessage"
          class="archive-delete-alert"
          type="error"
          show-icon
          :message="deleteDialog.errorMessage"
        />
      </div>
    </AModal>
  </section>
</template>

<style scoped>
.files-page {
  grid-template-columns: repeat(2, minmax(0, 1fr));
  grid-template-rows: 72px 548px 106px;
  grid-template-areas:
    'metrics metrics'
    'list side'
    'quick quick';
  --archive-table-header-height: 34px;
  --archive-table-row-height: 38.2px;
  --archive-table-body-height: calc(var(--archive-table-row-height) * 10);
}

.files-metrics {
  grid-area: metrics;
}

.files-metrics .metric-card {
  min-height: 72px;
}

.files-metrics > .metric-card {
  height: 72px;
  min-width: 0;
  overflow: hidden;
}

.cache-status-body {
  display: grid;
  align-content: center;
  gap: 3px;
  min-width: 0;
  overflow: hidden;
}

.cache-status-head {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 8px;
  min-width: 0;
}

.cache-status-head > span {
  overflow: hidden;
  color: var(--ink-strong);
  font-size: 14px;
  font-weight: 700;
  line-height: 1.15;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.cache-status-value {
  flex: 0 0 auto;
  color: var(--blue);
  font-size: 14px;
  font-weight: 700;
  line-height: 1.15;
  white-space: nowrap;
}

.cache-status-value.green {
  color: var(--green);
}

.cache-status-value.purple {
  color: var(--purple);
}

.cache-status-value.orange {
  color: var(--orange);
}

.cache-status-value.red {
  color: #d9413f;
}

.cache-status-body small,
.cache-process-title {
  display: block;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.cache-status-body small {
  color: var(--ink-muted);
  font-size: 11px;
  font-weight: 500;
  line-height: 1.2;
}

.cache-process-content {
  display: grid;
  gap: 1px;
  min-width: 0;
  overflow: hidden;
}

.cache-process-title {
  color: var(--ink);
  font-size: 12px;
  font-weight: 600;
  line-height: 1.2;
}

.cache-process-slide-enter-active,
.cache-process-slide-leave-active {
  transition: opacity 180ms ease, transform 180ms ease;
}

.cache-process-slide-enter-from {
  opacity: 0;
  transform: translateY(5px);
}

.cache-process-slide-leave-to {
  opacity: 0;
  transform: translateY(-5px);
}

.files-page .section-heading {
  font-weight: 500;
  letter-spacing: 0;
}

.file-list {
  grid-area: list;
  display: grid;
  grid-template-rows: auto minmax(0, 1fr) auto;
  row-gap: 4px;
  min-height: 0;
  padding: 18px 20px;
}

.file-list-header {
  display: grid;
  grid-template-columns: max-content minmax(0, 1fr);
  align-items: center;
  gap: 24px;
  min-height: 36px;
  min-width: 0;
}

.file-list-header .section-heading {
  margin: 0;
  white-space: nowrap;
}

.file-list,
.record-detail,
.bottom-quick-actions {
  box-shadow: var(--paper-shadow-sm), var(--paper-shadow-md);
}

.files-page :deep(.archive-ant-table .ant-table-thead > tr > th) {
  height: var(--archive-table-header-height);
  padding: 0 8px;
  border-bottom: 1px solid rgba(104, 141, 181, 0.16);
  color: var(--ink-strong);
  background: rgba(234, 244, 251, 0.48);
  font-weight: 500;
  text-align: center;
}

.files-page :deep(.archive-ant-table .ant-table-tbody) {
  height: var(--archive-table-body-height);
}

.files-page :deep(.archive-ant-table .ant-table-tbody > tr) {
  height: var(--archive-table-row-height);
}

.files-page :deep(.archive-ant-table .ant-table-tbody > tr > td) {
  box-sizing: border-box;
  height: var(--archive-table-row-height);
  padding: 0 8px;
  border-bottom: 0;
  box-shadow: inset 0 -1px 0 rgba(104, 141, 181, 0.09);
  color: var(--ink-strong);
  background: transparent;
  font-weight: 400;
}

.files-page :deep(.archive-ant-table .ant-table-tbody > tr:last-child > td) {
  border-bottom: 0;
  box-shadow: none;
}

.files-page :deep(.archive-ant-table .ant-table) {
  color: var(--ink-strong);
  background: transparent;
  font-size: 13px;
}

.files-page :deep(.archive-ant-table .ant-table-container),
.files-page :deep(.archive-ant-table .ant-table-content),
.files-page :deep(.archive-ant-table table) {
  background: transparent;
}

.files-page :deep(.archive-ant-table .ant-table-tbody > tr:hover > td) {
  background: rgba(45, 117, 214, 0.05);
}

.files-page :deep(.archive-ant-table .ant-btn.text-link) {
  height: 28px;
  min-height: 28px;
  line-height: 1;
}

.file-list .table-wrap {
  position: relative;
  z-index: 1;
  align-self: stretch;
  height: calc(var(--archive-table-header-height) + var(--archive-table-body-height));
  min-height: 0;
  margin-top: 8px;
  overflow: hidden;
}

.account-table {
  table-layout: fixed;
}

.account-table :deep(.ant-table table),
.record-table :deep(.ant-table table) {
  table-layout: fixed !important;
}

.account-index-col {
  width: 46px;
}

.account-name-col {
  width: auto;
}

.account-time-col {
  width: 104px;
}

.account-size-col {
  width: 60px;
}

.account-action-col {
  width: 184px;
}

.file-list .data-table th {
  text-align: center;
}

.file-list .data-table th.action-cell {
  text-align: center;
}

.file-list .data-table td:nth-child(1),
.file-list .data-table td:nth-child(2),
.file-list .data-table td:nth-child(3),
.file-list .data-table td:nth-child(4),
.file-list .data-table td.action-cell {
  text-align: center;
}

.account-table td:nth-child(2) {
  overflow: hidden;
  text-overflow: ellipsis;
}

.file-list :deep(.account-table .ant-table-thead > tr > th),
.file-list :deep(.account-table .ant-table-tbody > tr > td) {
  text-align: center;
}

.account-table :deep(.account-name-col) {
  overflow: hidden;
  text-overflow: ellipsis;
}

.account-name-cell {
  max-width: 100%;
  overflow: hidden;
  white-space: nowrap;
}

.account-table .table-actions {
  display: inline-flex;
  align-items: center;
  flex-wrap: nowrap;
  gap: 6px;
  justify-content: center;
}

.account-table .text-link {
  min-width: 50px;
  padding: 0 6px;
  cursor: pointer;
  font-size: 13px;
  line-height: 1.2;
}

.account-table .text-link:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}

.files-page .table-state-cell {
  display: grid;
  place-items: center;
  height: var(--archive-table-body-height);
  color: var(--ink-muted);
  font-size: 13px;
  font-weight: 400;
  text-align: center;
}

.files-page .table-placeholder-row {
  pointer-events: none;
}

.files-page :deep(.archive-ant-table .table-placeholder-row > td) {
  padding: 0;
  border-bottom: 0;
  box-shadow: none;
  background: transparent;
}

.files-page :deep(.archive-ant-table .table-state-row > td) {
  padding: 0;
  box-shadow: none;
}

.files-page .table-state-cell.error {
  color: #c93e3a;
}

.file-side {
  grid-area: side;
  display: grid;
  grid-template-rows: 1fr;
  min-width: 0;
}

.record-detail,
.bottom-quick-actions {
  padding: 18px 24px;
}

.record-detail {
  display: grid;
  grid-template-rows: auto minmax(0, 1fr) auto;
  row-gap: 4px;
  min-height: 0;
}

.bottom-quick-actions {
  padding: 12px 20px;
}

.detail-art {
  top: 0;
  right: 0;
  width: 116px;
}

.record-detail-header {
  display: grid;
  grid-template-columns: max-content minmax(0, 1fr);
  align-items: center;
  gap: 24px;
  min-height: 36px;
  min-width: 0;
}

.record-detail-header .section-heading {
  margin: 0;
  white-space: nowrap;
}

.record-actions {
  display: flex;
  align-items: center;
  justify-content: space-between;
  justify-self: end;
  gap: 10px;
  width: min(100%, 690px);
  min-width: 0;
  margin-top: 0;
}

.record-selection-summary {
  display: inline-flex;
  align-items: center;
  gap: 14px;
  flex: 1 1 auto;
  min-width: 0;
  min-height: 34px;
  padding: 0 12px;
  border: 1px solid var(--line-soft);
  border-radius: 8px;
  color: var(--ink);
  background: rgba(255, 255, 255, 0.36);
  box-shadow: var(--paper-shadow-sm);
  font-size: 13px;
  font-weight: 400;
  white-space: nowrap;
}

.record-account-name {
  color: var(--ink-strong);
  font-weight: 700;
}

.record-action-buttons {
  display: inline-flex;
  align-items: center;
  justify-content: flex-end;
  gap: 10px;
  margin-left: auto;
}

.record-actions .action-button {
  min-width: 82px;
}

.record-actions .action-button.primary {
  color: #ffffff;
  border-color: #2d70cc;
  background: #2d75d6;
  box-shadow: none;
}

.record-actions .action-button.primary:hover:not(:disabled) {
  color: #ffffff;
  border-color: #245fae;
  background: #245fae;
  box-shadow: none;
  transform: none;
}

.record-actions .action-button.danger {
  color: #ffffff;
  border-color: #c93e3a;
  background: #d9413f;
  box-shadow: none;
}

.record-actions .action-button.danger:hover:not(:disabled) {
  color: #ffffff;
  border-color: #b8322f;
  background: #b8322f;
  box-shadow: none;
  transform: none;
}

.record-empty-shell {
  min-width: 0;
  min-height: 0;
}

.record-empty {
  box-sizing: border-box;
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 462px;
  margin-top: 16px;
  padding: 24px;
  border: 1px dashed rgba(104, 141, 181, 0.26);
  border-radius: 10px;
  color: var(--ink-muted);
  background: rgba(255, 255, 255, 0.26);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.48);
  text-align: center;
}

.record-empty-loading {
  padding: 64px 54px;
  text-align: left;
}

.record-empty-loading :deep(.ant-skeleton) {
  width: min(100%, 420px);
}

.record-empty-result {
  display: block;
  padding-top: 92px;
}

.record-empty-result :deep(.ant-result-title) {
  color: var(--ink-strong);
  font-size: 17px;
  font-weight: 500;
}

.record-empty-result :deep(.ant-result-subtitle) {
  color: var(--ink-muted);
  font-size: 13px;
  font-weight: 400;
}

.record-empty-first-run {
  flex-direction: column;
  gap: 10px;
}

.record-empty-icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 58px;
  height: 58px;
  margin-bottom: 4px;
  border-radius: 12px;
  color: #2d75d6;
  background: rgba(45, 117, 214, 0.1);
  font-size: 27px;
}

.record-empty-first-run h3,
.archive-overview-copy h3 {
  margin: 0;
  color: var(--ink-strong);
  font-size: 18px;
  font-weight: 500;
  letter-spacing: 0;
}

.record-empty-first-run p {
  max-width: 330px;
  margin: 0;
  color: var(--ink-muted);
  font-size: 13px;
  font-weight: 400;
  line-height: 1.6;
}

.record-empty-primary {
  min-width: 104px;
  height: 32px;
  margin-top: 4px;
  border-color: #2d70cc;
  border-radius: 6px;
  color: #ffffff;
  background: #2d75d6;
  box-shadow: none;
  font-size: 13px;
  font-weight: 400;
}

.record-empty-primary:hover:not(:disabled) {
  border-color: #245fae;
  color: #ffffff;
  background: #245fae;
  box-shadow: none;
}

.record-overview-empty {
  display: grid;
  grid-template-columns: minmax(210px, 0.9fr) minmax(0, 1.1fr);
  align-items: center;
  gap: 22px 30px;
  padding: 26px 22px 30px;
  border: 0;
  background: transparent;
  box-shadow: none;
  text-align: left;
}

.archive-overview-chart-wrap {
  display: grid;
  gap: 14px;
  min-width: 0;
}

.archive-overview-guide {
  grid-column: 1 / -1;
  display: grid;
  grid-template-columns: 32px minmax(0, 1fr) auto;
  align-items: center;
  gap: 12px;
  min-width: 0;
  padding: 12px 14px;
  border: 1px solid rgba(45, 117, 214, 0.16);
  border-radius: 8px;
  color: #15386f;
  background: rgba(235, 246, 253, 0.72);
  white-space: normal;
}

.archive-overview-guide-icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  border-radius: 8px;
  color: #2d75d6;
  background: rgba(45, 117, 214, 0.1);
  font-size: 15px;
}

.archive-overview-guide-copy {
  display: grid;
  gap: 4px;
  min-width: 0;
}

.archive-overview-guide strong,
.archive-overview-guide-copy > span {
  overflow: visible;
  text-overflow: clip;
  white-space: normal;
}

.archive-overview-guide strong {
  font-size: 16px;
  font-weight: 500;
  line-height: 1.35;
}

.archive-overview-guide-copy > span {
  color: rgba(21, 56, 111, 0.74);
  font-size: 13px;
  font-weight: 400;
  line-height: 1.35;
}

.archive-overview-guide-steps {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  flex-wrap: wrap;
  gap: 7px;
  min-width: 0;
}

.archive-overview-guide-steps span {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 26px;
  padding: 0 10px;
  border: 1px solid rgba(45, 117, 214, 0.16);
  border-radius: 999px;
  color: #15386f;
  background: rgba(255, 255, 255, 0.62);
  font-size: 12px;
  font-weight: 400;
  line-height: 1;
  white-space: nowrap;
}

.archive-overview-copy {
  display: grid;
  align-content: center;
  gap: 10px;
  min-width: 0;
}

.archive-overview-label {
  color: #2d75d6;
  font-size: 13px;
  font-weight: 500;
}

.archive-overview-copy > p {
  margin: 0;
  color: var(--ink-muted);
  font-size: 13px;
  font-weight: 400;
  line-height: 1.7;
}

.archive-overview-stats {
  display: grid;
  gap: 7px;
  margin: 2px 0 0;
}

.archive-overview-stats > div {
  display: grid;
  grid-template-columns: 64px minmax(0, 1fr);
  gap: 8px;
  min-width: 0;
}

.archive-overview-stats dt,
.archive-overview-stats dd {
  margin: 0;
  font-size: 12px;
  font-weight: 400;
  line-height: 1.45;
}

.archive-overview-stats dt {
  color: var(--ink-muted);
}

.archive-overview-stats dd {
  overflow: hidden;
  color: var(--ink-strong);
  text-overflow: ellipsis;
  white-space: nowrap;
}

.archive-overview-legend {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 6px 12px;
  margin-top: 4px;
}

.archive-overview-legend span {
  display: flex;
  align-items: center;
  gap: 6px;
  min-width: 0;
  overflow: hidden;
  color: var(--ink);
  font-size: 11px;
  font-weight: 400;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.archive-overview-legend i {
  flex: 0 0 auto;
  width: 7px;
  height: 7px;
  border-radius: 50%;
}

.action-button.danger {
  color: #ffffff;
  border-color: rgba(217, 65, 63, 0.44);
  background:
    radial-gradient(circle at 20% 15%, rgba(255, 255, 255, 0.28), transparent 40%),
    linear-gradient(135deg, #e4635f, #c93e3a);
}

.record-table-wrap {
  align-self: stretch;
  height: calc(var(--archive-table-header-height) + var(--archive-table-body-height));
  min-height: 0;
  margin-top: 8px;
  overflow: hidden;
}

.record-table {
  table-layout: fixed;
  font-size: 13px;
}

.record-table :deep(.ant-table-thead > tr > th),
.record-table :deep(.ant-table-tbody > tr > td) {
  padding: 0 9px;
}

.record-table :deep(.ant-table-thead > tr > th) {
  z-index: 2;
  text-align: center;
}

.record-action-col {
  width: 58px;
}

.record-time-col {
  width: 142px;
}

.record-size-col {
  width: 76px;
}

.record-open-col {
  width: 86px;
}

.record-table .checkbox-cell,
.record-table .record-time,
.record-table .record-size,
.record-table .record-open-cell {
  text-align: center;
}

.record-table :deep(.checkbox-cell),
.record-table :deep(.record-time-col),
.record-table :deep(.record-size-col),
.record-table :deep(.record-open-col) {
  text-align: center;
}

.record-table :deep(.record-title) {
  min-width: 0;
}

.record-title {
  color: var(--ink-strong);
}

.record-title-link {
  display: block;
  width: 100%;
  height: auto;
  min-width: 0;
  padding: 0;
  border: 0;
  color: var(--blue);
  background: transparent;
  box-shadow: none;
  font: inherit;
  font-weight: 400;
  line-height: 1.2;
  text-align: left;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  cursor: pointer;
}

.record-title-link:hover:not(:disabled) {
  color: #1e5fae;
  text-decoration: none;
}

.record-title-link:disabled {
  color: var(--ink-strong);
  cursor: default;
}

.record-open-cell .text-link {
  padding: 0 1px;
  cursor: pointer;
  font-size: 13px;
  line-height: 1.2;
}

.account-table .text-link,
.record-open-cell .text-link {
  display: inline-grid;
  place-items: center;
  min-height: 26px;
  border-radius: 6px;
  background: rgba(45, 117, 214, 0.08);
  text-decoration: none;
  transition:
    background 150ms ease,
    color 150ms ease;
}

.account-table .table-actions .text-link {
  min-width: 50px;
  padding: 0 6px;
}

.file-list .pagination-bar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  width: 100%;
  box-sizing: border-box;
  min-height: 40px;
  padding-top: 4px;
  padding-right: 0;
  padding-left: 0;
  margin-top: 0;
}

.account-pagination-total {
  flex: 0 0 auto;
  margin: 0;
  padding: 0;
  padding-left: 8px;
}

.file-list .pagination-bar .pagination-controls {
  flex: 0 0 auto;
  width: auto;
  min-width: 0;
  margin-left: auto;
}

.file-list .pagination-bar .pagination-controls {
  display: inline-flex;
  align-items: center;
  justify-content: flex-end;
  gap: 8px;
  width: auto;
}

.account-table .text-link:hover:not(:disabled),
.record-open-cell .text-link:hover:not(:disabled) {
  background: rgba(45, 117, 214, 0.14);
}

.account-table .text-link.danger,
.record-open-cell .text-link.danger {
  color: var(--red);
  background: rgba(217, 65, 63, 0.08);
}

.account-table .text-link.danger:hover:not(:disabled),
.record-open-cell .text-link.danger:hover:not(:disabled) {
  background: rgba(217, 65, 63, 0.13);
}

.record-open-cell .text-link:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}

.record-checkbox {
  width: 16px;
  height: 16px;
  accent-color: #2d70cc;
  cursor: pointer;
  vertical-align: middle;
}

.record-checkbox:disabled {
  cursor: not-allowed;
  opacity: 0.48;
}

.record-pagination {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  width: 100%;
  min-height: 40px;
  margin-top: 8px;
  color: var(--ink);
  font-size: 13px;
  font-weight: 900;
}

.record-pagination-total {
  flex: 0 0 auto;
  margin: 0;
  padding: 0;
  padding-left: 8px;
  font-weight: 400;
}

.file-list :deep(.file-ant-pager.ant-pagination),
.record-detail :deep(.file-ant-pager.ant-pagination) {
  flex: 1 1 auto;
  min-width: 0;
  margin-left: auto;
  padding: 0;
  color: var(--ink);
  background: transparent;
  font-size: 13px;
  font-weight: 900;
}

.record-detail :deep(.record-ant-pager.ant-pagination) {
  flex: 0 0 auto;
  width: auto;
  margin-left: auto;
}

.pagination-controls {
  display: inline-flex;
  align-items: center;
  justify-content: flex-end;
  gap: 6px;
  width: auto;
  min-height: 32px;
}

.file-list :deep(.file-ant-pager .ant-pagination-prev),
.file-list :deep(.file-ant-pager .ant-pagination-next),
.file-list :deep(.file-ant-pager .ant-pagination-item),
.file-list :deep(.file-ant-pager .ant-pagination-jump-prev),
.file-list :deep(.file-ant-pager .ant-pagination-jump-next),
.record-detail :deep(.file-ant-pager .ant-pagination-prev),
.record-detail :deep(.file-ant-pager .ant-pagination-next),
.record-detail :deep(.file-ant-pager .ant-pagination-item),
.record-detail :deep(.file-ant-pager .ant-pagination-jump-prev),
.record-detail :deep(.file-ant-pager .ant-pagination-jump-next) {
  display: inline-grid;
  place-items: center;
  min-width: 32px;
  height: 32px;
  padding: 0 9px;
  border: 1px solid transparent;
  border-radius: 6px;
  color: var(--blue);
  background: rgba(255, 255, 255, 0.64);
  box-shadow: none;
  font: inherit;
  line-height: 1;
  cursor: pointer;
  transition:
    border-color 150ms ease,
    background 150ms ease,
    color 150ms ease;
}

.file-list :deep(.file-ant-pager .ant-pagination-prev),
.file-list :deep(.file-ant-pager .ant-pagination-next),
.record-detail :deep(.file-ant-pager .ant-pagination-prev),
.record-detail :deep(.file-ant-pager .ant-pagination-next) {
  width: 32px;
  padding: 0;
  color: rgba(77, 108, 159, 0.78);
}

.file-list :deep(.file-ant-pager .ant-pagination-jump-prev),
.file-list :deep(.file-ant-pager .ant-pagination-jump-next),
.record-detail :deep(.file-ant-pager .ant-pagination-jump-prev),
.record-detail :deep(.file-ant-pager .ant-pagination-jump-next) {
  min-width: 24px;
  padding: 0 6px;
  color: rgba(77, 108, 159, 0.72);
}

.file-list :deep(.file-ant-pager .ant-pagination-prev:hover:not(.ant-pagination-disabled)),
.file-list :deep(.file-ant-pager .ant-pagination-next:hover:not(.ant-pagination-disabled)),
.file-list :deep(.file-ant-pager .ant-pagination-item:hover:not(.ant-pagination-item-active)),
.file-list :deep(.file-ant-pager .ant-pagination-jump-prev:hover),
.file-list :deep(.file-ant-pager .ant-pagination-jump-next:hover),
.record-detail :deep(.file-ant-pager .ant-pagination-prev:hover:not(.ant-pagination-disabled)),
.record-detail :deep(.file-ant-pager .ant-pagination-next:hover:not(.ant-pagination-disabled)),
.record-detail :deep(.file-ant-pager .ant-pagination-item:hover:not(.ant-pagination-item-active)),
.record-detail :deep(.file-ant-pager .ant-pagination-jump-prev:hover),
.record-detail :deep(.file-ant-pager .ant-pagination-jump-next:hover) {
  border-color: rgba(45, 117, 214, 0.24);
  background: rgba(255, 255, 255, 0.86);
}

.file-list :deep(.file-ant-pager .ant-pagination-item-active),
.record-detail :deep(.file-ant-pager .ant-pagination-item-active) {
  color: #ffffff;
  border-color: rgba(45, 117, 214, 0.32);
  background: linear-gradient(135deg, #4d85dc, #2d70cc);
}

.file-list :deep(.file-ant-pager .ant-pagination-disabled),
.record-detail :deep(.file-ant-pager .ant-pagination-disabled) {
  cursor: not-allowed;
  color: rgba(77, 108, 159, 0.36);
  background: rgba(255, 255, 255, 0.46);
}

.file-list :deep(.file-ant-pager .ant-pagination-item-link),
.record-detail :deep(.file-ant-pager .ant-pagination-item-link) {
  line-height: 1;
}

.file-list :deep(.file-ant-pager .ant-pagination-item a),
.record-detail :deep(.file-ant-pager .ant-pagination-item a) {
  color: inherit;
}

.fixed-page-size {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 82px;
  height: 30px;
  border: 1px solid var(--line);
  border-radius: 7px;
  color: var(--ink);
  background: rgba(255, 255, 255, 0.46);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.62);
  font-size: 13px;
  font-weight: 400;
  white-space: nowrap;
}

.bottom-quick-actions {
  grid-area: quick;
}

.quick-actions-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 18px;
}

.bottom-quick-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 16px;
  margin-top: 10px;
}

.bottom-quick-grid .action-button {
  width: 100%;
  height: 40px;
  font-size: 15px;
  color: #15386f;
  border-color: rgba(96, 137, 178, 0.22);
  background: rgba(255, 255, 255, 0.5);
  box-shadow: var(--paper-shadow-sm);
  font-weight: 500;
}

.bottom-quick-grid .action-button:hover:not(:disabled) {
  color: #12376d;
  border-color: rgba(45, 117, 214, 0.28);
  background: rgba(234, 244, 251, 0.68);
  box-shadow: var(--paper-shadow-sm);
  transform: none;
}

.bottom-quick-grid .action-button:active:not(:disabled) {
  transform: none;
  background: rgba(222, 237, 248, 0.72);
}

.bottom-quick-grid .action-button.purple {
  color: #5c4eb5;
  border-color: rgba(111, 100, 214, 0.22);
  background: rgba(246, 244, 255, 0.78);
}

.bottom-quick-grid .action-button.danger {
  color: #b4232e;
  border-color: rgba(217, 65, 63, 0.24);
  background: rgba(255, 241, 241, 0.76);
}

.bottom-quick-grid .action-button:disabled {
  box-shadow: var(--paper-shadow-sm);
  transform: none;
}

.batch-export-modal :deep(.ant-modal-content),
.archive-delete-modal :deep(.ant-modal-content) {
  overflow: hidden;
  border: 1px solid rgba(104, 141, 181, 0.24);
  border-radius: 8px;
  box-shadow: 0 10px 18px rgba(23, 52, 86, 0.16);
}

.batch-export-modal :deep(.ant-modal-header) {
  margin: 0;
  padding: 18px 20px 16px;
  border-bottom: 1px solid rgba(104, 141, 181, 0.16);
}

.batch-export-modal :deep(.ant-modal-body) {
  padding: 0;
}

.batch-export-modal :deep(.ant-modal-footer) {
  margin: 0;
  padding: 14px 20px 18px;
  border-top: 1px solid rgba(104, 141, 181, 0.16);
}

.batch-export-head {
  display: grid;
  grid-template-columns: 42px minmax(0, 1fr);
  gap: 12px;
  align-items: center;
}

.batch-export-icon {
  display: grid;
  place-items: center;
  width: 38px;
  height: 38px;
  border: 1px solid rgba(31, 143, 105, 0.2);
  border-radius: 8px;
  color: #1f8f69;
  background: rgba(31, 143, 105, 0.12);
  font-size: 16px;
}

.batch-export-head h3 {
  margin: 0;
  font-size: 18px;
  font-weight: 500;
  line-height: 1.2;
}

.batch-export-head p {
  margin: 6px 0 0;
  font-size: 13px;
  font-weight: 400;
  line-height: 1.45;
  opacity: 0.72;
}

.batch-export-meta {
  margin: 14px 20px 0;
  font-size: 13px;
  font-weight: 400;
  line-height: 1.35;
  opacity: 0.76;
}

.batch-export-table-wrap {
  min-height: 220px;
  max-height: 380px;
  margin: 10px 20px 16px;
  overflow: auto;
}

.file-list .table-wrap,
.record-table-wrap,
.batch-export-table-wrap {
  border: 1px solid rgba(104, 141, 181, 0.18);
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.38);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.58);
}

.batch-export-table-wrap {
  background: transparent;
  box-shadow: none;
}

.batch-export-modal :deep(.batch-export-ant-table .ant-table) {
  background: transparent;
  font-size: 13px;
}

.batch-export-modal :deep(.batch-export-ant-table .ant-table-thead > tr > th),
.batch-export-modal :deep(.batch-export-ant-table .ant-table-tbody > tr > td) {
  height: 40px;
  padding: 0 12px;
}

.batch-export-modal :deep(.batch-export-ant-table .ant-table-thead > tr > th) {
  font-weight: 500;
  text-align: center;
}

.batch-export-modal :deep(.batch-export-ant-table .ant-table-cell) {
  text-align: center;
}

.batch-export-modal :deep(.batch-export-ant-table .ant-table-cell:nth-child(3)) {
  text-align: left;
}

.batch-export-modal :deep(.batch-export-ant-table .ant-empty) {
  margin-block: 52px;
  text-align: center;
  font-size: 13px;
  font-weight: 400;
}

.export-account-cell {
  display: block;
  overflow: hidden;
  font-weight: 400;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.batch-export-actions {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 10px;
}

.batch-export-actions .action-button {
  min-width: 104px;
}

.archive-delete-content {
  display: grid;
  gap: 14px;
}

.archive-delete-intro {
  display: grid;
  grid-template-columns: 38px minmax(0, 1fr);
  align-items: center;
  gap: 12px;
}

.archive-delete-icon {
  display: grid;
  place-items: center;
  width: 36px;
  height: 36px;
  border-radius: 8px;
  color: #c93e3a;
  background: rgba(217, 65, 63, 0.1);
}

.archive-delete-intro p {
  margin: 0;
  font-size: 14px;
  font-weight: 400;
  line-height: 1.55;
}

.archive-delete-summary {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 8px;
  margin: 0;
}

.archive-delete-summary > div {
  min-width: 0;
  padding: 10px 12px;
  border-radius: 7px;
  background: rgba(45, 117, 214, 0.06);
}

.archive-delete-summary dt,
.archive-delete-summary dd {
  overflow: hidden;
  margin: 0;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.archive-delete-summary dt {
  font-size: 12px;
  opacity: 0.68;
}

.archive-delete-summary dd {
  margin-top: 4px;
  font-size: 14px;
  font-weight: 500;
}

.archive-delete-detail strong {
  font-size: 13px;
  font-weight: 500;
}

.archive-delete-detail ul {
  max-height: 120px;
  margin: 8px 0 0;
  padding-left: 20px;
  overflow: auto;
  font-size: 13px;
  line-height: 1.6;
}

.archive-delete-alert :deep(.ant-alert-message) {
  font-weight: 400;
}

@media (prefers-reduced-motion: reduce) {
  .cache-process-slide-enter-active,
  .cache-process-slide-leave-active {
    transition: opacity 80ms ease;
  }

  .cache-process-slide-enter-from,
  .cache-process-slide-leave-to {
    transform: none;
  }
}

.data-table td i {
  width: 22px;
  color: var(--green);
}

.file-filters {
  display: grid;
  grid-template-columns: 72px minmax(0, 110px) minmax(0, 1fr);
  justify-self: end;
  position: relative;
  z-index: 8;
  width: min(100%, 520px);
  gap: 10px;
  margin-top: 0;
  padding: 0;
  border: 0;
  border-radius: 0;
  background: transparent;
  box-shadow: none;
}

.file-refresh-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 5px;
  width: 72px;
  height: 32px;
  padding: 0 10px;
  border-color: var(--paper-edge);
  border-radius: 6px;
  color: #15386f;
  background: var(--frost-bg-strong);
  box-shadow: none;
  font-size: 14px;
  font-weight: 400;
}

.file-refresh-button:hover:not(:disabled) {
  border-color: rgba(45, 117, 214, 0.34);
  color: #15386f;
  background: rgba(235, 246, 253, 0.92);
  box-shadow: none;
  transform: none;
}

.file-ant-control,
.file-date-picker {
  width: 100%;
  min-width: 0;
  height: 32px;
  border-color: var(--paper-edge);
  border-radius: 6px;
  color: var(--ink-strong);
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm);
  font-size: 14px;
  font-weight: 400;
}

.file-list :deep(.file-ant-select.ant-select) {
  width: 100%;
}

.file-list :deep(.file-ant-select.ant-select .ant-select-selector) {
  height: 32px;
  padding: 0 11px;
  border-color: var(--paper-edge);
  border-radius: 6px;
  color: var(--ink-strong);
  background: var(--frost-bg-strong);
  box-shadow: var(--paper-shadow-sm);
  font-size: 14px;
  font-weight: 400;
}

.file-list :deep(.file-ant-select.ant-select .ant-select-selection-item),
.file-list :deep(.file-ant-select.ant-select .ant-select-selection-placeholder),
.file-list :deep(.file-ant-select.ant-select .ant-select-selection-search-input) {
  color: var(--ink-strong);
  font-size: 14px;
  font-weight: 400;
  line-height: 30px;
}

.file-list :deep(.file-ant-select.ant-select .ant-select-selection-placeholder) {
  color: rgba(77, 108, 159, 0.72);
}

.file-search-icon {
  color: rgba(77, 108, 159, 0.78);
  font-size: 14px;
  pointer-events: none;
}

.file-list :deep(.file-ant-select.ant-select .ant-select-arrow),
.file-list :deep(.file-ant-select.ant-select .ant-select-clear) {
  color: rgba(77, 108, 159, 0.72);
}

.file-list :deep(.file-ant-select.ant-select .ant-select-clear) {
  background: transparent;
  font-size: 12px;
}

.file-list :deep(.file-ant-select.ant-select .ant-select-clear:hover) {
  color: rgba(21, 56, 111, 0.82);
}

.file-list :deep(.file-date-picker.ant-picker) {
  display: inline-flex;
  align-items: center;
  height: 32px;
  padding-inline: 12px;
}

.file-list :deep(.file-date-picker .ant-picker-input > input),
.file-list :deep(.file-date-picker .ant-picker-separator),
.file-list :deep(.file-date-picker .ant-picker-suffix),
.file-list :deep(.file-date-picker .ant-picker-clear) {
  color: var(--ink);
  font-size: 14px;
  font-weight: 400;
}

.file-list :deep(.file-date-picker .ant-picker-input > input) {
  text-align: center;
}

.file-list :deep(.file-date-picker .ant-picker-input > input::placeholder) {
  color: rgba(77, 108, 159, 0.72);
  font-weight: 400;
  text-align: center;
}

:global(.file-account-filter-dropdown.ant-select-dropdown) {
  width: max-content;
  min-width: calc(240px * var(--app-scale));
  max-width: calc(360px * var(--app-scale));
  padding: calc(4px * var(--app-scale));
  border: 1px solid var(--line);
  border-radius: calc(7px * var(--app-scale));
  color: var(--ink);
  background: #fbfdff;
  box-shadow: 0 12px 22px rgba(35, 69, 111, 0.14);
  font-size: calc(14px * var(--app-scale));
  font-weight: 400;
}

:global(.file-account-filter-dropdown.ant-select-dropdown .ant-select-item) {
  min-height: calc(32px * var(--app-scale));
  padding: calc(5px * var(--app-scale)) calc(12px * var(--app-scale));
  border-radius: calc(4px * var(--app-scale));
  color: var(--ink);
  font-size: calc(14px * var(--app-scale));
  font-weight: 400;
  line-height: 1.25;
}

:global(.file-account-filter-dropdown.ant-select-dropdown .ant-select-item-option-content) {
  overflow: visible;
  text-overflow: clip;
  white-space: nowrap;
}

:global(.file-account-filter-dropdown.ant-select-dropdown .ant-select-item-option-selected) {
  color: var(--ink-strong);
  background: rgba(45, 117, 214, 0.1);
  font-weight: 500;
}

:global(.file-date-picker-panel .ant-picker-panel-container) {
  zoom: var(--app-scale);
}

:global(.collector-app.dark) .record-selection-summary,
:global(.collector-app.dark) .record-pagination,
:global(.collector-app.dark) .fixed-page-size {
  color: #b9c8dd;
}

:global(.collector-app.dark) .file-filters {
  border: 0;
  background: transparent;
  box-shadow: none;
}

:global(.collector-app.dark) .file-list .table-wrap,
:global(.collector-app.dark) .record-table-wrap,
:global(.collector-app.dark) .batch-export-table-wrap {
  border-color: rgba(128, 153, 188, 0.18);
  background: rgba(15, 24, 39, 0.52);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.045);
}

:global(.collector-app.dark) .files-page :deep(.archive-ant-table .ant-table),
:global(.collector-app.dark) .files-page :deep(.archive-ant-table .ant-table-container),
:global(.collector-app.dark) .files-page :deep(.archive-ant-table .ant-table-content),
:global(.collector-app.dark) .files-page :deep(.archive-ant-table table) {
  color: #dce7f5;
  background: transparent;
}

:global(.collector-app.dark) .files-page :deep(.archive-ant-table .ant-table-thead > tr > th) {
  border-bottom-color: rgba(128, 153, 188, 0.18);
  color: #dce7f5;
  background: rgba(28, 43, 64, 0.62);
}

:global(.collector-app.dark) .files-page :deep(.archive-ant-table .ant-table-tbody > tr > td) {
  border-bottom-color: transparent;
  box-shadow: inset 0 -1px 0 rgba(128, 153, 188, 0.12);
  color: #dce7f5;
  background: transparent;
}

:global(.collector-app.dark) .files-page :deep(.archive-ant-table .ant-table-tbody > tr:last-child > td),
:global(.collector-app.dark) .files-page :deep(.archive-ant-table .table-placeholder-row > td),
:global(.collector-app.dark) .files-page :deep(.archive-ant-table .table-state-row > td) {
  box-shadow: none;
}

:global(.collector-app.dark) .files-page :deep(.archive-ant-table .ant-table-tbody > tr:hover > td) {
  background: rgba(83, 132, 190, 0.12);
}

:global(.collector-app.dark) .record-selection-summary {
  border-color: rgba(128, 153, 188, 0.18);
  background: rgba(15, 24, 39, 0.44);
}

:global(.collector-app.dark) .record-empty {
  border-color: rgba(128, 153, 188, 0.2);
  color: #8ea2bd;
  background: rgba(15, 24, 39, 0.36);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.045);
}

:global(.collector-app.dark) .record-overview-empty {
  border: 0;
  background: transparent;
  box-shadow: none;
}

:global(.collector-app.dark) .record-empty-loading :deep(.ant-skeleton-title),
:global(.collector-app.dark) .record-empty-loading :deep(.ant-skeleton-paragraph > li) {
  background: linear-gradient(90deg, rgba(128, 153, 188, 0.16), rgba(128, 153, 188, 0.28), rgba(128, 153, 188, 0.16));
}

:global(.collector-app.dark) .record-empty-result :deep(.ant-result-title),
:global(.collector-app.dark) .record-empty-first-run h3,
:global(.collector-app.dark) .archive-overview-copy h3,
:global(.collector-app.dark) .archive-overview-stats dd {
  color: #dce7f5;
}

:global(.collector-app.dark) .record-empty-result :deep(.ant-result-subtitle),
:global(.collector-app.dark) .record-empty-first-run p,
:global(.collector-app.dark) .archive-overview-guide-copy > span,
:global(.collector-app.dark) .archive-overview-stats dt,
:global(.collector-app.dark) .archive-overview-legend span {
  color: #9fb2cc;
}

:global(.collector-app.dark) .archive-overview-label {
  color: #8fbded;
}

:global(.collector-app.dark) .archive-overview-guide {
  border-color: rgba(111, 154, 211, 0.2);
  color: #dce7f5;
  background: rgba(31, 52, 80, 0.54);
}

:global(.collector-app.dark) .archive-overview-guide-icon {
  color: #8fbded;
  background: rgba(83, 132, 190, 0.16);
}

:global(.collector-app.dark) .archive-overview-guide-steps span {
  color: #cbd8ea;
  border-color: rgba(111, 154, 211, 0.18);
  background: rgba(15, 24, 39, 0.5);
}

:global(.collector-app.dark) .record-empty-icon {
  color: #8fbded;
  background: rgba(83, 132, 190, 0.16);
}

:global(.collector-app.dark) .record-empty-primary {
  color: #eef6ff;
  border-color: #3f7fc8;
  background: #2f6fb8;
}

:global(.collector-app.dark) .record-empty-primary:hover:not(:disabled) {
  color: #ffffff;
  border-color: #4d8ed8;
  background: #397dcc;
}

:global(.collector-app.dark) .record-title-link {
  color: #8fbded;
}

:global(.collector-app.dark) .record-title-link:hover:not(:disabled) {
  color: #b8d7f8;
}

:global(.collector-app.dark) .record-title-link:disabled {
  color: #dce7f5;
}

:global(.collector-app.dark) .action-button.danger {
  border-color: rgba(181, 83, 84, 0.34);
  background: linear-gradient(180deg, rgba(132, 82, 80, 0.82), rgba(105, 64, 70, 0.8));
}

:global(.collector-app.dark) .record-actions .action-button.primary {
  color: #eef6ff;
  border-color: #3f7fc8;
  background: #2f6fb8;
  box-shadow: none;
}

:global(.collector-app.dark) .record-actions .action-button.primary:hover:not(:disabled) {
  color: #ffffff;
  border-color: #4d8ed8;
  background: #397dcc;
  box-shadow: none;
  transform: none;
}

:global(.collector-app.dark) .record-actions .action-button.danger {
  color: #fff3f2;
  border-color: #b84d4a;
  background: #a9403d;
  box-shadow: none;
}

:global(.collector-app.dark) .record-actions .action-button.danger:hover:not(:disabled) {
  color: #ffffff;
  border-color: #cc5a56;
  background: #bd4a46;
  box-shadow: none;
  transform: none;
}

:global(.collector-app.dark) .file-refresh-button,
:global(.collector-app.dark) .file-list :deep(.file-ant-select.ant-select .ant-select-selector),
:global(.collector-app.dark) .file-list :deep(.file-date-picker.ant-picker) {
  border-color: rgba(128, 153, 188, 0.2);
  background: rgba(15, 24, 39, 0.62);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.045);
}

:global(.collector-app.dark) .file-refresh-button,
:global(.collector-app.dark) .file-search-icon,
:global(.collector-app.dark) .file-list :deep(.file-ant-select.ant-select .ant-select-selection-item),
:global(.collector-app.dark) .file-list :deep(.file-ant-select.ant-select .ant-select-selection-search-input),
:global(.collector-app.dark) .file-list :deep(.file-ant-select.ant-select .ant-select-arrow),
:global(.collector-app.dark) .file-list :deep(.file-ant-select.ant-select .ant-select-clear),
:global(.collector-app.dark) .file-list :deep(.file-date-picker .ant-picker-input > input),
:global(.collector-app.dark) .file-list :deep(.file-date-picker .ant-picker-separator),
:global(.collector-app.dark) .file-list :deep(.file-date-picker .ant-picker-suffix),
:global(.collector-app.dark) .file-list :deep(.file-date-picker .ant-picker-clear) {
  color: #cbd8ea;
}

:global(.collector-app.dark) .file-list :deep(.file-ant-select.ant-select .ant-select-selection-placeholder),
:global(.collector-app.dark) .file-list :deep(.file-date-picker .ant-picker-input > input::placeholder) {
  color: rgba(142, 162, 189, 0.72);
}

:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-prev),
:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-next),
:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-item),
:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-jump-prev),
:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-jump-next),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-prev),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-next),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-item),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-jump-prev),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-jump-next),
:global(.collector-app.dark) .fixed-page-size {
  color: #a9bfda;
  border-color: rgba(128, 153, 188, 0.16);
  background: rgba(17, 27, 44, 0.72);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.04);
}

:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-prev:hover:not(.ant-pagination-disabled)),
:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-next:hover:not(.ant-pagination-disabled)),
:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-item:hover:not(.ant-pagination-item-active)),
:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-jump-prev:hover),
:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-jump-next:hover),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-prev:hover:not(.ant-pagination-disabled)),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-next:hover:not(.ant-pagination-disabled)),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-item:hover:not(.ant-pagination-item-active)),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-jump-prev:hover),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-jump-next:hover) {
  color: #dceaff;
  border-color: rgba(111, 154, 211, 0.34);
  background: rgba(24, 38, 60, 0.86);
}

:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-item-active),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-item-active) {
  color: #f0f6ff;
  border-color: rgba(111, 154, 211, 0.36);
  background: linear-gradient(180deg, #376fb0, #285c9c);
}

:global(.collector-app.dark) .file-list :deep(.file-ant-pager .ant-pagination-disabled),
:global(.collector-app.dark) .record-detail :deep(.file-ant-pager .ant-pagination-disabled) {
  color: rgba(142, 162, 189, 0.42);
  background: rgba(17, 27, 44, 0.42);
}

</style>
