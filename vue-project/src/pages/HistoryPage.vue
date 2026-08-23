<script setup lang="ts">
import { DownOutlined, ReloadOutlined, SearchOutlined } from '@ant-design/icons-vue'
import AppIcon from '../components/AppIcon.vue'
import { notification } from 'ant-design-vue'
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import {
  clearHistoryRecords,
  getHistoryRecords,
  getHistorySuggestions,
  getHistorySummary,
  type HistoryRecordItem,
  type HistorySummary,
} from '../bridge/pythonApi'
import type { MetricCard } from '../data/mockData'
import {
  buildHistorySuggestionOptions,
  type HistorySuggestionQuery,
} from '../utils/historySuggestions'

type HistoryRecord = HistoryRecordItem
type HistoryDateRangeValue = [string, string] | null

type HistoryTableStateRow = {
  rowKind: 'state'
  tableKey: string
  stateText: string
  stateTone: 'default' | 'error'
}

type HistoryTablePlaceholderRow = {
  rowKind: 'placeholder'
  tableKey: string
}

type HistoryTableDataRow = HistoryRecord & {
  rowKind: 'data'
  tableKey: string
  displayIndex: number
}

type HistoryTableRow = HistoryTableDataRow | HistoryTableStateRow | HistoryTablePlaceholderRow

type HistoryTableColumn = {
  title: string
  key: string
  dataIndex?: string
  width?: number | string
  align?: 'left' | 'center' | 'right'
  className?: string
  customCell?: (record: HistoryTableRow) => { colSpan?: number }
}

const HISTORY_VISIBLE_ROWS = 15
const HISTORY_TABLE_HEADER_HEIGHT = 32
const DEFAULT_HISTORY_TABLE_BODY_HEIGHT = 480
const HISTORY_QUERY_DEBOUNCE_MS = 250
const HISTORY_TABLE_COLUMN_COUNT = 7
const HISTORY_FILTER_TYPE_DETAIL_KEY = 'history-filter:type:detail'
const HISTORY_FILTER_TYPE_COMMENTS_KEY = 'history-filter:type:comments'
const HISTORY_FILTER_STATUS_SUCCESS_KEY = 'history-filter:status:success'
const HISTORY_FILTER_STATUS_FAILED_KEY = 'history-filter:status:failed'
const HISTORY_FILTER_COLLECT_TYPE_KEYS = ['history-filter:type:detail', 'history-filter:type:comments']
const HISTORY_FILTER_STATUS_KEYS = ['history-filter:status:success', 'history-filter:status:failed']

// Ant Design Vue Select 的空字符串会被当成已选值，未输入时保持 undefined 才会显示 placeholder。
const keyword = ref<string | undefined>(undefined)
const selectedCollectType = ref('')
const selectedStatus = ref('')
const selectedCollectStartDate = ref('')
const selectedCollectEndDate = ref('')
const historyCurrentPage = ref(1)
const historyPageSize = ref(HISTORY_VISIBLE_ROWS)
const selectedHistoryId = ref<number | null>(null)
const historyRecords = ref<HistoryRecord[]>([])
const historyTotal = ref(0)
const historyLoading = ref(false)
const historyError = ref('')
const historySummary = ref<HistorySummary | null>(null)
const summaryLoading = ref(false)
const historyClearing = ref(false)
const clearHistoryDialogOpen = ref(false)
const clearHistoryError = ref('')
const historySuggestions = ref<string[]>([])
const historyTableWrapRef = ref<HTMLElement | null>(null)
const historyTableBodyHeight = ref(DEFAULT_HISTORY_TABLE_BODY_HEIGHT)
let historyTableResizeObserver: ResizeObserver | null = null
let historyRecordsDebounceTimer: ReturnType<typeof setTimeout> | null = null
let historySuggestionsDebounceTimer: ReturnType<typeof setTimeout> | null = null
let historyRecordsRequestId = 0
let historySuggestionsRequestId = 0

// 表格固定 15 行展示，行高由分页上方的剩余空间反推；向下取值避免 15 行累计后产生内部滚动。
const historyRowHeight = computed(() => Math.floor((historyTableBodyHeight.value / HISTORY_VISIBLE_ROWS) * 100) / 100)
const historyGridStyle = computed<Record<string, string>>(() => ({
  '--history-row-height': `${historyRowHeight.value}px`,
  '--history-body-height': `${historyTableBodyHeight.value}px`,
}))

const historyFilterTreeData = [
  {
    title: '采集类型',
    value: 'history-filter:type',
    selectable: false,
    disableCheckbox: true,
    children: [
      { title: '文章详情', value: HISTORY_FILTER_TYPE_DETAIL_KEY },
      { title: '评论信息', value: HISTORY_FILTER_TYPE_COMMENTS_KEY },
    ],
  },
  {
    title: '任务状态',
    value: 'history-filter:status',
    selectable: false,
    disableCheckbox: true,
    children: [
      { title: '成功', value: HISTORY_FILTER_STATUS_SUCCESS_KEY },
      { title: '失败', value: HISTORY_FILTER_STATUS_FAILED_KEY },
    ],
  },
]

const collectTypeValueByFilterKey: Record<string, string> = {
  [HISTORY_FILTER_TYPE_DETAIL_KEY]: '文章详情',
  [HISTORY_FILTER_TYPE_COMMENTS_KEY]: '评论信息',
}

const collectTypeFilterKeyByValue: Record<string, string> = {
  文章详情: HISTORY_FILTER_TYPE_DETAIL_KEY,
  评论信息: HISTORY_FILTER_TYPE_COMMENTS_KEY,
}

const statusValueByFilterKey: Record<string, string> = {
  [HISTORY_FILTER_STATUS_SUCCESS_KEY]: 'success',
  [HISTORY_FILTER_STATUS_FAILED_KEY]: 'failed',
}

const statusFilterKeyByValue: Record<string, string> = {
  success: HISTORY_FILTER_STATUS_SUCCESS_KEY,
  failed: HISTORY_FILTER_STATUS_FAILED_KEY,
}

const statusLabelByValue: Record<string, string> = {
  success: '成功',
  failed: '失败',
}

const historyMetrics = computed<MetricCard[]>(() => {
  const summary = historySummary.value
  return [
    {
      label: '历史任务数',
      value: String(summary?.totalRecords ?? 0),
      icon: 'fa-regular fa-clipboard',
      tone: 'blue',
      hint: '来自 awa_fetch_history',
    },
    {
      label: '成功率',
      value: `${summary?.successRate ?? 0}%`,
      icon: 'fa-solid fa-check',
      tone: 'green',
      hint: 'success / 全部任务',
    },
    {
      label: '最近采集日期',
      value: summary?.latestCollectDate || '-',
      icon: 'fa-regular fa-calendar-days',
      tone: 'blue',
      hint: '按 started_time 统计',
    },
    {
      label: '累计采集文章',
      value: String(summary?.collectedArticleCount ?? 0),
      icon: 'fa-regular fa-file-lines',
      tone: 'orange',
      hint: '成功文章按 ID 去重',
    },
  ]
})

const chartBars = computed(() => historySummary.value?.trend.map((item) => item.count) ?? [])
const chartLabels = computed(() => historySummary.value?.trend.map((item) => item.label) ?? [])
const chartMaxValue = computed(() => Math.max(1, ...chartBars.value))
const canClearHistoryRecords = computed(() => {
  return !historyClearing.value && (historySummary.value?.totalRecords ?? historyTotal.value) > 0
})

// 根据采集结果文本统一映射状态色，避免表格和详情里的状态样式各写一套判断。
function getStatusTone(status: string) {
  if (status === '成功') {
    return 'success'
  }

  if (status === '失败') {
    return 'danger'
  }

  if (status === '部分完成') {
    return 'warning'
  }

  return 'neutral'
}

// 采集记录只展示开始采集时间，保留标准 HH:mm:ss，便于和真实采集日志对照。
const historyNameOptions = computed(() => {
  return buildHistorySuggestionOptions(historySuggestions.value)
})

const selectedHistoryFilterKeys = computed<string[]>({
  get: () => {
    const selectedKeys: string[] = []
    const collectTypeKey = collectTypeFilterKeyByValue[selectedCollectType.value]
    const statusKey = statusFilterKeyByValue[selectedStatus.value]

    if (collectTypeKey) {
      selectedKeys.push(collectTypeKey)
    }
    if (statusKey) {
      selectedKeys.push(statusKey)
    }

    return selectedKeys
  },
  set: (value: string[]) => applyHistoryFilterKeys(value),
})

const historyFilterSelectionLabel = computed(() => {
  const selectedLabels = [
    selectedCollectType.value,
    statusLabelByValue[selectedStatus.value],
  ].filter(Boolean)

  return selectedLabels.length > 0 ? `筛选：${selectedLabels.join(' · ')}` : '筛选：全部'
})

function normalizeHistoryKeyword(value: string | undefined) {
  const normalizedValue = value?.trim()
  return normalizedValue ? normalizedValue : undefined
}

function getLastHistoryFilterKey(value: string[], allowedKeys: string[]) {
  const matchedKeys = value.filter((key) => allowedKeys.includes(key))
  return matchedKeys.at(-1) ?? ''
}

function applyHistoryFilterKeys(value: string[]) {
  const checkedKeys = Array.isArray(value) ? value : []
  const collectTypeKey = getLastHistoryFilterKey(checkedKeys, HISTORY_FILTER_COLLECT_TYPE_KEYS)
  const statusKey = getLastHistoryFilterKey(checkedKeys, HISTORY_FILTER_STATUS_KEYS)

  selectedCollectType.value = collectTypeValueByFilterKey[collectTypeKey] ?? ''
  selectedStatus.value = statusValueByFilterKey[statusKey] ?? ''
}

function renderHistoryFilterTagPlaceholder() {
  return historyFilterSelectionLabel.value
}

function truncateAccountName(account: string) {
  const chars = Array.from(account || '')
  return chars.length > 10 ? `${chars.slice(0, 9).join('')}...` : account
}

const selectedCollectDateRange = computed<HistoryDateRangeValue>({
  get: (): HistoryDateRangeValue => {
    if (!selectedCollectStartDate.value || !selectedCollectEndDate.value) {
      return null
    }
    return [selectedCollectStartDate.value, selectedCollectEndDate.value]
  },
  set: (value: HistoryDateRangeValue) => {
    selectedCollectStartDate.value = value?.[0] ?? ''
    selectedCollectEndDate.value = value?.[1] ?? ''
  },
})

function makeHistoryStateCellAttrs(stateColumnKey: string) {
  return (columnKey: string) => (record: HistoryTableRow) => {
    if (record.rowKind !== 'state') {
      return {}
    }

    return columnKey === stateColumnKey ? { colSpan: HISTORY_TABLE_COLUMN_COUNT } : { colSpan: 0 }
  }
}

const historyStateCellAttrs = makeHistoryStateCellAttrs('seq')

const historyTableColumns: HistoryTableColumn[] = [
  { title: '序号', key: 'seq', width: 48, align: 'center', customCell: historyStateCellAttrs('seq') },
  {
    title: '记录名称',
    key: 'name',
    dataIndex: 'name',
    width: 200,
    align: 'left',
    customCell: historyStateCellAttrs('name'),
  },
  {
    title: '公众号',
    key: 'account',
    dataIndex: 'account',
    width: 128,
    align: 'center',
    customCell: historyStateCellAttrs('account'),
  },
  {
    title: '采集类型',
    key: 'collectType',
    dataIndex: 'collectType',
    width: 82,
    align: 'center',
    customCell: historyStateCellAttrs('collectType'),
  },
  {
    title: '记录时间',
    key: 'collectTime',
    dataIndex: 'collectTime',
    width: 168,
    align: 'center',
    customCell: historyStateCellAttrs('collectTime'),
  },
  {
    title: '状态',
    key: 'status',
    dataIndex: 'status',
    width: 86,
    align: 'center',
    className: 'history-status-column',
    customCell: historyStateCellAttrs('status'),
  },
  { title: '操作', key: 'action', width: 60, align: 'center', customCell: historyStateCellAttrs('action') },
]

const historyPlaceholderRowCount = computed(() => {
  if (historyLoading.value || historyError.value || historyRecords.value.length === 0) {
    return 0
  }
  return Math.max(HISTORY_VISIBLE_ROWS - historyRecords.value.length, 0)
})

const historyTableRows = computed<HistoryTableRow[]>(() => {
  if (historyLoading.value) {
    return [{ rowKind: 'state', tableKey: 'history-loading', stateText: '正在读取采集记录...', stateTone: 'default' }]
  }
  if (historyError.value) {
    return [{ rowKind: 'state', tableKey: 'history-error', stateText: historyError.value, stateTone: 'error' }]
  }
  if (historyRecords.value.length === 0) {
    return [{ rowKind: 'state', tableKey: 'history-empty', stateText: '暂无采集记录', stateTone: 'default' }]
  }

  const pageStart = (historyCurrentPage.value - 1) * historyPageSize.value
  const rows: HistoryTableRow[] = historyRecords.value.map((record, index) => ({
    ...record,
    rowKind: 'data',
    tableKey: `history-${record.id}`,
    displayIndex: pageStart + index + 1,
  }))

  for (let index = 0; index < historyPlaceholderRowCount.value; index += 1) {
    rows.push({ rowKind: 'placeholder', tableKey: `history-placeholder-${index + 1}` })
  }

  return rows
})

function getHistoryTableRowKey(record: HistoryTableRow) {
  return record.tableKey
}

function getHistoryTableCustomRow(record: HistoryTableRow) {
  if (record.rowKind === 'placeholder') {
    return { class: 'history-placeholder-row', 'aria-hidden': 'true' }
  }
  if (record.rowKind === 'state') {
    return { class: 'history-state-row' }
  }
  return {}
}

const selectedHistoryRecord = computed(() => {
  if (!selectedHistoryId.value) {
    return null
  }

  return historyRecords.value.find((row) => row.id === selectedHistoryId.value) ?? null
})

const recordDetail = computed(() => {
  const record = selectedHistoryRecord.value

  if (!record) {
    return null
  }

  return {
    name: record.name,
    account: record.account,
    collectType: record.collectType,
    startedTime: record.startedTime,
    finishedTime: record.finishedTime,
    publishedArticleTime: record.publishedArticleTime,
    articleLink: record.articleLink,
    duration: record.duration,
    status: record.status,
    collectStatus: record.collectStatus,
    resourceTypeLabels: record.resourceTypeLabels,
    outputDir: record.outputDir,
    errorStageLabel: record.errorStageLabel,
    errorMessage: record.errorMessage,
  }
})

async function loadHistorySummary() {
  summaryLoading.value = true
  try {
    historySummary.value = await getHistorySummary()
  } catch (error) {
    // 汇总失败不阻断列表和候选加载，页面用空指标明确降级。
    historySummary.value = null
  } finally {
    summaryLoading.value = false
  }
}

async function loadHistorySuggestions(query: Partial<HistorySuggestionQuery> = {}) {
  const requestId = ++historySuggestionsRequestId
  try {
    const result = await getHistorySuggestions({
      keyword: query.keyword ?? keyword.value ?? '',
      limit: query.limit ?? 30,
    })
    if (requestId !== historySuggestionsRequestId) {
      return
    }
    historySuggestions.value = result.items
  } catch {
    if (requestId !== historySuggestionsRequestId) {
      return
    }
    // 候选只影响输入提示，失败时保留空列表，不阻断历史记录查询。
    historySuggestions.value = []
  }
}

function scheduleHistorySuggestionsLoad(query: Partial<HistorySuggestionQuery> = {}) {
  if (historySuggestionsDebounceTimer) {
    clearTimeout(historySuggestionsDebounceTimer)
  }
  historySuggestionsDebounceTimer = setTimeout(() => {
    historySuggestionsDebounceTimer = null
    void loadHistorySuggestions(query)
  }, HISTORY_QUERY_DEBOUNCE_MS)
}

function updateHistoryTableMetrics() {
  const tableWrap = historyTableWrapRef.value

  if (!tableWrap) {
    return
  }

  const nextBodyHeight = tableWrap.clientHeight - HISTORY_TABLE_HEADER_HEIGHT

  if (nextBodyHeight <= 0) {
    return
  }

  historyTableBodyHeight.value = Number(nextBodyHeight.toFixed(2))
}

async function loadHistoryRecords(page = historyCurrentPage.value, pageSize = historyPageSize.value) {
  const requestId = ++historyRecordsRequestId
  historyLoading.value = true
  historyError.value = ''
  try {
    const result = await getHistoryRecords({
      page,
      pageSize,
      keyword: keyword.value ?? '',
      collectType: selectedCollectType.value,
      status: selectedStatus.value,
      collectStartDate: selectedCollectStartDate.value,
      collectEndDate: selectedCollectEndDate.value,
    })
    if (requestId !== historyRecordsRequestId) {
      return
    }
    historyCurrentPage.value = result.page
    historyPageSize.value = result.pageSize
    historyTotal.value = result.total
    historyRecords.value = result.items
    if (selectedHistoryId.value && !result.items.some((item) => item.id === selectedHistoryId.value)) {
      selectedHistoryId.value = null
    }
  } catch (error) {
    if (requestId !== historyRecordsRequestId) {
      return
    }
    historyError.value = error instanceof Error ? error.message : '读取采集记录失败'
    historyRecords.value = []
    historyTotal.value = 0
  } finally {
    if (requestId === historyRecordsRequestId) {
      historyLoading.value = false
    }
  }
}

function scheduleHistoryRecordsLoad() {
  if (historyRecordsDebounceTimer) {
    clearTimeout(historyRecordsDebounceTimer)
  }
  historyRecordsDebounceTimer = setTimeout(() => {
    historyRecordsDebounceTimer = null
    void loadHistoryRecords(1, historyPageSize.value)
  }, HISTORY_QUERY_DEBOUNCE_MS)
}

function clearHistoryQueryTimers() {
  if (historyRecordsDebounceTimer) {
    clearTimeout(historyRecordsDebounceTimer)
    historyRecordsDebounceTimer = null
  }
  if (historySuggestionsDebounceTimer) {
    clearTimeout(historySuggestionsDebounceTimer)
    historySuggestionsDebounceTimer = null
  }
}

async function resetHistoryFilters() {
  clearHistoryQueryTimers()
  keyword.value = undefined
  selectedCollectType.value = ''
  selectedStatus.value = ''
  selectedCollectStartDate.value = ''
  selectedCollectEndDate.value = ''
  historyCurrentPage.value = 1
  await Promise.all([loadHistorySummary(), loadHistorySuggestions(), loadHistoryRecords(1, historyPageSize.value)])
}

function openClearHistoryDialog() {
  if (!canClearHistoryRecords.value) {
    return
  }
  clearHistoryError.value = ''
  clearHistoryDialogOpen.value = true
}

function closeClearHistoryDialog() {
  if (historyClearing.value) {
    return
  }
  clearHistoryDialogOpen.value = false
}

async function confirmClearHistory() {
  if (historyClearing.value) {
    return
  }

  historyClearing.value = true
  clearHistoryError.value = ''
  try {
    const result = await clearHistoryRecords()
    selectedHistoryId.value = null
    historyCurrentPage.value = 1
    clearHistoryDialogOpen.value = false
    notification.success({
      message: '采集历史已清空',
      description: result.message ?? `已清空 ${result.deletedCount} 条采集历史记录。`,
      placement: 'bottomRight',
    })
    await Promise.all([loadHistorySummary(), loadHistorySuggestions(), loadHistoryRecords(1, historyPageSize.value)])
  } catch (error) {
    const message = error instanceof Error ? error.message : '清空采集历史失败'
    clearHistoryError.value = message
    notification.error({
      message: '清空失败',
      description: message,
      placement: 'bottomRight',
    })
  } finally {
    historyClearing.value = false
  }
}

function selectHistoryRecord(row: HistoryRecord) {
  selectedHistoryId.value = row.id
}

function handleHistoryKeywordSearch(value: string) {
  keyword.value = normalizeHistoryKeyword(value)
  scheduleHistorySuggestionsLoad({ keyword: keyword.value ?? '' })
}

function handleHistoryKeywordChange(value: string | undefined) {
  keyword.value = normalizeHistoryKeyword(value)
}

async function handleAntHistoryPageChange(page: number, pageSize: number) {
  historyCurrentPage.value = page
  historyPageSize.value = pageSize
  await loadHistoryRecords(page, pageSize)
}

watch(keyword, () => {
  historyCurrentPage.value = 1
  scheduleHistoryRecordsLoad()
  scheduleHistorySuggestionsLoad()
})

watch([selectedCollectType, selectedStatus, selectedCollectStartDate, selectedCollectEndDate], async () => {
  historyCurrentPage.value = 1
  await loadHistoryRecords(1, historyPageSize.value)
})

async function refreshOnActivated() {
  clearHistoryQueryTimers()
  await Promise.all([
    loadHistorySummary(),
    loadHistorySuggestions(),
    loadHistoryRecords(historyCurrentPage.value, historyPageSize.value),
  ])
  await nextTick()
  updateHistoryTableMetrics()
}

defineExpose({
  refreshOnActivated,
})

onMounted(async () => {
  await nextTick()
  updateHistoryTableMetrics()

  if (typeof ResizeObserver !== 'undefined' && historyTableWrapRef.value) {
    historyTableResizeObserver = new ResizeObserver(updateHistoryTableMetrics)
    historyTableResizeObserver.observe(historyTableWrapRef.value)
  }
})

onBeforeUnmount(() => {
  clearHistoryQueryTimers()
  historyRecordsRequestId += 1
  historySuggestionsRequestId += 1
  historyTableResizeObserver?.disconnect()
})
</script>

<template>
  <section class="management-page history-page" aria-label="采集历史">
    <div class="metric-grid history-metrics config-summary-metrics">
      <article v-for="item in historyMetrics" :key="item.label" class="metric-card page-panel">
        <span :class="['metric-icon', item.tone]">
          <AppIcon :icon="item.icon" />
        </span>
        <div class="metric-body">
          <span>{{ item.label }}</span>
          <strong :class="item.tone">{{ item.value }}</strong>
        </div>
      </article>
    </div>

    <section class="history-list page-panel" aria-label="采集记录">
      <div class="history-list-header">
        <h2 class="section-heading">
          <AppIcon icon="fa-regular fa-rectangle-list" />
          采集记录
        </h2>

        <div class="filters-row history-filters">
          <AButton class="history-refresh-button" html-type="button" :loading="historyLoading" @click="resetHistoryFilters">
            <ReloadOutlined />
            刷新
          </AButton>
          <ASelect
            v-model:value="keyword"
            class="history-ant-control history-ant-select history-keyword"
            show-search
            allow-clear
            placeholder="搜索公众号或文章标题"
            :options="historyNameOptions"
            :filter-option="false"
            :dropdown-match-select-width="false"
            popup-class-name="history-keyword-panel"
            aria-label="搜索公众号或文章标题"
            @search="handleHistoryKeywordSearch"
            @change="handleHistoryKeywordChange"
          >
            <template #suffixIcon>
              <SearchOutlined class="history-search-icon" />
            </template>
          </ASelect>
          <ATreeSelect
            v-model:value="selectedHistoryFilterKeys"
            class="history-ant-control history-ant-select history-filter-tree"
            :tree-data="historyFilterTreeData"
            tree-checkable
            tree-default-expand-all
            :show-search="false"
            allow-clear
            :max-tag-count="0"
            :max-tag-placeholder="renderHistoryFilterTagPlaceholder"
            placeholder="筛选：全部"
            popup-class-name="history-filter-select-panel"
            aria-label="采集筛选"
          >
            <template #suffixIcon>
              <DownOutlined class="history-select-chevron" />
            </template>
          </ATreeSelect>
          <ARangePicker
            v-model:value="selectedCollectDateRange"
            class="history-date-picker history-date-range-picker"
            :allow-clear="true"
            format="YYYY-MM-DD"
            value-format="YYYY-MM-DD"
            separator=" ~ "
            popup-class-name="history-date-range-picker-panel"
            :placeholder="['采集起始日期', '采集结束日期']"
            aria-label="采集起始日期和采集结束日期"
          />
        </div>
      </div>

      <div ref="historyTableWrapRef" class="history-table-wrap">
        <ATable
          class="history-ant-table"
          :style="historyGridStyle"
          :columns="historyTableColumns"
          :data-source="historyTableRows"
          :pagination="false"
          :row-key="getHistoryTableRowKey"
          :custom-row="getHistoryTableCustomRow"
          :locale="{ emptyText: null }"
          size="small"
          table-layout="fixed"
        >
          <template #bodyCell="{ column, record }">
            <span
              v-if="record.rowKind === 'state' && column.key === 'seq'"
              :class="['history-table-state', { error: record.stateTone === 'error' }]"
            >
              {{ record.stateText }}
            </span>
            <template v-else-if="record.rowKind === 'data'">
              <template v-if="column.key === 'seq'">
                {{ record.displayIndex }}
              </template>
              <span v-else-if="column.key === 'name'" class="history-title-cell" :title="record.name">
                {{ record.name }}
              </span>
              <span v-else-if="column.key === 'account'" class="history-account-cell" :title="record.account">
                {{ truncateAccountName(record.account) }}
              </span>
              <template v-else-if="column.key === 'collectType'">
                {{ record.collectType }}
              </template>
              <template v-else-if="column.key === 'collectTime'">
                {{ record.collectTime }}
              </template>
              <ATag
                v-else-if="column.key === 'status'"
                class="history-status-tag"
                :class="'status-' + getStatusTone(record.status)"
                :bordered="false"
              >
                {{ record.status }}
              </ATag>
              <AButton
                v-else-if="column.key === 'action'"
                class="history-view-link"
                type="link"
                html-type="button"
                @click="selectHistoryRecord(record)"
              >
                查看
              </AButton>
            </template>
          </template>
        </ATable>
      </div>

      <div class="history-ant-pagination" aria-label="采集历史分页">
        <span class="history-total">共 {{ historyTotal }} 条</span>
        <div class="history-pagination-controls">
          <APagination
            v-model:current="historyCurrentPage"
            v-model:page-size="historyPageSize"
            class="history-ant-pager"
            size="small"
            :total="historyTotal"
            :page-size-options="[HISTORY_VISIBLE_ROWS]"
            :show-size-changer="false"
            :show-less-items="false"
            @change="handleAntHistoryPageChange"
          />
          <span class="history-page-size">{{ historyPageSize }} 条/页</span>
        </div>
      </div>
    </section>

    <aside class="history-side">
      <section class="task-detail page-panel">
        <img class="panel-corner-art detail-art" src="/assets/watercolor-leaf-branch-b.png" alt="" />
        <h2 class="section-heading">
          <AppIcon icon="fa-regular fa-file-lines" />
          记录详情
        </h2>
        <div v-if="recordDetail" class="detail-list">
          <div class="detail-row">
            <span>公众号名</span>
            <strong>{{ recordDetail.account }}</strong>
          </div>
          <div class="detail-row">
            <span>记录名称</span>
            <strong>{{ recordDetail.name }}</strong>
          </div>
          <div class="detail-row">
            <span>文章发布时间</span>
            <strong>{{ recordDetail.publishedArticleTime || '-' }}</strong>
          </div>
          <div class="detail-row">
            <span>开始时间</span>
            <strong>{{ recordDetail.startedTime || '-' }}</strong>
          </div>
          <div class="detail-row">
            <span>结束时间</span>
            <strong>{{ recordDetail.finishedTime || '-' }}</strong>
          </div>
          <div class="detail-row">
            <span>采集类型</span>
            <strong>{{ recordDetail.collectType }}</strong>
          </div>
          <div class="detail-row">
            <span>执行时长</span>
            <strong>{{ recordDetail.duration }}</strong>
          </div>
          <div class="detail-row">
            <span>文章链接</span>
            <strong class="detail-value-wrap" :title="recordDetail.articleLink">
              {{ recordDetail.articleLink || '-' }}
            </strong>
          </div>
          <div class="detail-row">
            <span>运行结果</span>
            <strong>
              <ATag
                class="history-status-tag"
                :class="'status-' + getStatusTone(recordDetail.status)"
                :bordered="false"
              >
                {{ recordDetail.status }}
              </ATag>
            </strong>
          </div>
          <template v-if="recordDetail.collectStatus === 'failed'">
            <div class="detail-row">
              <span>失败阶段</span>
              <strong>{{ recordDetail.errorStageLabel || '-' }}</strong>
            </div>
            <div class="detail-row">
              <span>失败原因</span>
              <strong class="detail-value-wrap">{{ recordDetail.errorMessage || '未记录具体原因' }}</strong>
            </div>
          </template>
          <template v-else>
            <div class="detail-row">
              <span>资源类型</span>
              <strong class="detail-value-wrap">
                {{ recordDetail.resourceTypeLabels.join('、') || '-' }}
              </strong>
            </div>
            <div class="detail-row">
              <span>输出目录</span>
              <strong class="detail-value-wrap" :title="recordDetail.outputDir">
                {{ recordDetail.outputDir || '-' }}
              </strong>
            </div>
          </template>
        </div>
        <div v-else class="detail-empty">
          <AppIcon icon="fa-regular fa-hand-pointer" />
          <strong>点击左侧“查看”</strong>
          <span>选择一条采集记录后，这里会显示文章标题、公众号、采集类型、采集时间、执行时长和记录内容。</span>
        </div>
      </section>

      <section class="history-stats page-panel">
        <div class="history-stats-header">
          <h2 class="section-heading">
            <AppIcon icon="fa-solid fa-chart-column" />
            历史统计
          </h2>
          <AButton
            class="history-clear-button"
            danger
            html-type="button"
            :disabled="!canClearHistoryRecords"
            :loading="historyClearing"
            @click="openClearHistoryDialog"
          >
            <AppIcon v-if="!historyClearing" icon="fa-regular fa-trash-can" />
            {{ historyClearing ? '清空中' : '清空记录' }}
          </AButton>
        </div>
        <p class="chart-caption">近日采集趋势</p>
        <div class="trend-chart" aria-label="近日采集趋势">
          <div v-for="(bar, index) in chartBars" :key="chartLabels[index]" class="trend-item">
            <ATooltip
              :title="chartLabels[index] + '：' + bar + ' 条'"
              placement="top"
            >
              <span
                class="trend-bar"
                :style="{ height: Math.round((bar / chartMaxValue) * 92) + 'px' }"
              ></span>
            </ATooltip>
            <small>{{ chartLabels[index] }}</small>
          </div>
        </div>
        <div class="summary-grid">
          <div>
            <AppIcon icon="fa-solid fa-check" />
            <strong>{{ historySummary?.successfulRecords ?? 0 }}</strong>
            <span>成功记录</span>
          </div>
          <div>
            <AppIcon icon="fa-solid fa-xmark" />
            <strong>{{ historySummary?.failedRecords ?? 0 }}</strong>
            <span>失败记录</span>
          </div>
          <div>
            <AppIcon icon="fa-regular fa-clock" />
            <strong>{{ summaryLoading ? '读取中' : (historySummary?.averageDuration ?? '00:00.00') }}</strong>
            <span>平均耗时</span>
          </div>
        </div>
      </section>
    </aside>

    <AModal
      v-model:open="clearHistoryDialogOpen"
      class="history-clear-modal"
      title="清空采集历史"
      :closable="!historyClearing"
      :keyboard="!historyClearing"
      :mask-closable="!historyClearing"
      :confirm-loading="historyClearing"
      :ok-button-props="{ danger: true }"
      ok-text="确认清空"
      cancel-text="取消"
      centered
      @cancel="closeClearHistoryDialog"
      @ok="confirmClearHistory"
    >
      <div class="history-clear-content">
        <div class="history-clear-intro">
          <span class="history-clear-icon" aria-hidden="true">
            <AppIcon icon="fa-regular fa-trash-can" />
          </span>
          <p>将清空全部采集历史流水记录，历史统计和采集记录列表会同步刷新。</p>
        </div>
        <dl class="history-clear-summary">
          <div>
            <dt>影响数据表</dt>
            <dd>awa_fetch_history</dd>
          </div>
          <div>
            <dt>当前记录数</dt>
            <dd>{{ historySummary?.totalRecords ?? historyTotal }} 条</dd>
          </div>
        </dl>
        <AAlert
          class="history-clear-alert"
          type="warning"
          show-icon
          message="不会删除文章归档和公众号数据。"
        />
        <AAlert
          v-if="clearHistoryError"
          class="history-clear-alert"
          type="error"
          show-icon
          :message="clearHistoryError"
        />
      </div>
    </AModal>
  </section>
</template>

<style scoped>
.history-page {
  height: 100%;
  grid-template-columns: minmax(0, 1fr) 420px;
  grid-template-rows: 72px minmax(0, 1fr);
  grid-template-areas:
    'metrics metrics'
    'list side';
}

.history-metrics {
  grid-area: metrics;
}

.history-list {
  grid-area: list;
  display: flex;
  flex-direction: column;
  min-height: 0;
  padding: 16px 20px 8px;
}

.history-list-header {
  display: grid;
  grid-template-columns: max-content minmax(0, 1fr);
  align-items: center;
  gap: 24px;
  min-height: 48px;
  min-width: 0;
}

.history-list-header .section-heading {
  margin: 0;
  white-space: nowrap;
}

.history-side {
  grid-area: side;
  display: grid;
  /* 右侧总高度跟主内容行一致，第二块吃掉剩余空间，避免底部超出左侧列表。 */
  grid-template-rows: 348px minmax(0, 1fr);
  gap: 14px;
  min-height: 0;
  min-width: 0;
}

.task-detail,
.history-stats,
.history-log {
  padding: 18px 24px;
}

.history-stats-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}

.history-clear-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 5px;
  min-width: 94px;
  height: 32px;
  padding: 0 11px;
  border-color: rgba(217, 65, 63, 0.24);
  border-radius: 6px;
  color: #b84242;
  background: rgba(255, 255, 255, 0.56);
  box-shadow: var(--paper-shadow-sm);
  font-size: 13px;
  font-weight: 500;
}

.history-clear-button:hover,
.history-clear-button:focus-visible {
  border-color: rgba(217, 65, 63, 0.42);
  color: #a83636;
  background: rgba(255, 241, 240, 0.92);
}

.history-clear-button:disabled {
  color: rgba(77, 108, 159, 0.42);
  border-color: rgba(104, 141, 181, 0.14);
  background: rgba(255, 255, 255, 0.4);
}

.history-clear-content {
  display: grid;
  gap: 14px;
}

.history-clear-intro {
  display: grid;
  grid-template-columns: 38px minmax(0, 1fr);
  align-items: center;
  gap: 12px;
}

.history-clear-icon {
  display: grid;
  place-items: center;
  width: 38px;
  height: 38px;
  border-radius: 10px;
  color: #d9413f;
  background: rgba(255, 241, 240, 0.92);
}

.history-clear-intro p {
  margin: 0;
  color: var(--ink);
  font-size: 14px;
  line-height: 1.7;
}

.history-clear-summary {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
  margin: 0;
}

.history-clear-summary div {
  padding: 10px 12px;
  border: 1px solid var(--line-soft);
  border-radius: 7px;
  background: rgba(234, 244, 251, 0.38);
}

.history-clear-summary dt {
  margin-bottom: 4px;
  color: var(--ink-muted);
  font-size: 12px;
}

.history-clear-summary dd {
  margin: 0;
  color: var(--ink-strong);
  font-size: 14px;
  font-weight: 500;
}

.history-clear-alert {
  border-radius: 7px;
}

.history-filters {
  grid-template-columns: 76px minmax(154px, 176px) minmax(132px, 152px) minmax(250px, 1fr);
  justify-self: end;
  position: relative;
  z-index: 8;
  width: min(100%, 696px);
  gap: 10px;
  margin-top: 0;
  padding: 0;
  border: 0;
  border-radius: 0;
  background: transparent;
  box-shadow: none;
}

.history-ant-control,
.history-date-picker,
.history-refresh-button {
  width: 100%;
  min-width: 0;
  height: 38px;
  font-size: 14px;
  font-weight: 400;
}

.history-keyword,
.history-filter-tree,
.history-date-range-picker {
  min-width: 0;
}

.history-list :deep(.history-ant-select.ant-select) {
  width: 100%;
}

.history-list :deep(.history-ant-select.ant-select .ant-select-selector) {
  height: 38px;
  padding: 0 11px;
  border-color: var(--line);
  border-radius: 6px;
  color: var(--ink);
  background: rgba(255, 255, 255, 0.48);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.6);
  font-size: 14px;
  font-weight: 400;
}

.history-list :deep(.history-ant-select.ant-select .ant-select-selection-item),
.history-list :deep(.history-ant-select.ant-select .ant-select-selection-placeholder),
.history-list :deep(.history-ant-select.ant-select .ant-select-selection-search-input) {
  color: var(--ink);
  font-size: 14px;
  font-weight: 400;
  line-height: 36px;
}

.history-list :deep(.history-ant-select.ant-select .ant-select-selection-placeholder) {
  color: rgba(77, 108, 159, 0.72);
}

.history-search-icon {
  color: rgba(77, 108, 159, 0.78);
  font-size: 14px;
}

.history-select-chevron {
  color: rgba(77, 108, 159, 0.84);
  font-size: 10px;
  pointer-events: none;
  transform: rotate(0deg);
  transition: transform 160ms ease;
}

.history-select-chevron.is-open {
  transform: rotate(180deg);
}

.history-list :deep(.history-filter-tree.ant-select-open .history-select-chevron) {
  transform: rotate(180deg);
}

.history-list :deep(.history-ant-select.ant-select .ant-select-arrow),
.history-list :deep(.history-date-picker.ant-picker .ant-picker-suffix),
.history-list :deep(.history-date-picker.ant-picker .ant-picker-clear) {
  color: rgba(77, 108, 159, 0.72);
}

.history-list :deep(.history-date-picker.ant-picker) {
  width: 100%;
  height: 38px;
  padding: 0 11px;
  border-color: var(--line);
  border-radius: 6px;
  color: var(--ink);
  background: rgba(255, 255, 255, 0.48);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.6);
  font-size: 14px;
  font-weight: 400;
}

.history-list :deep(.history-date-picker.ant-picker input),
.history-list :deep(.history-date-picker.ant-picker .ant-picker-input > input) {
  color: var(--ink);
  font-size: 14px;
  font-weight: 400;
}

.history-list :deep(.history-date-picker.ant-picker input::placeholder) {
  color: rgba(77, 108, 159, 0.72);
}

.history-refresh-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 0 10px;
  border-color: rgba(45, 117, 214, 0.24);
  border-radius: 6px;
  color: var(--blue);
  background: rgba(255, 255, 255, 0.54);
  box-shadow: var(--paper-shadow-sm);
  font-size: 14px;
  font-weight: 500;
}

.history-refresh-button:hover,
.history-refresh-button:focus-visible {
  border-color: rgba(45, 117, 214, 0.42);
  color: var(--blue);
  background: rgba(238, 247, 255, 0.9);
}

:global(.history-date-range-picker-panel.ant-picker-dropdown) {
  color: var(--ink);
  font-size: 14px;
  font-weight: 400;
}

:global(.history-keyword-panel.ant-select-dropdown) {
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

:global(.history-filter-select-panel.ant-select-dropdown) {
  min-width: calc(152px * var(--app-scale));
  max-width: calc(176px * var(--app-scale));
  padding: calc(4px * var(--app-scale));
  border: 1px solid var(--line);
  border-radius: calc(7px * var(--app-scale));
  color: var(--ink);
  background: #fbfdff;
  box-shadow: 0 12px 22px rgba(35, 69, 111, 0.14);
  font-size: calc(14px * var(--app-scale));
  font-weight: 400;
}

:global(.history-keyword-panel .ant-select-item),
:global(.history-filter-select-panel .ant-select-item) {
  min-height: calc(32px * var(--app-scale));
  padding: calc(5px * var(--app-scale)) calc(12px * var(--app-scale));
  border-radius: calc(4px * var(--app-scale));
  color: var(--ink);
  font-size: calc(14px * var(--app-scale));
  font-weight: 400;
  line-height: 1.25;
}

:global(.history-keyword-panel .ant-select-item-option-content) {
  overflow: visible;
  text-overflow: clip;
  white-space: nowrap;
}

:global(.history-keyword-panel .ant-select-item-option-selected),
:global(.history-filter-select-panel .ant-select-item-option-selected) {
  color: var(--ink-strong);
  background: rgba(45, 117, 214, 0.1);
  font-weight: 500;
}

:global(.history-filter-select-panel .ant-select-tree) {
  color: var(--ink);
  background: transparent;
  font-size: calc(14px * var(--app-scale));
  line-height: 1.25;
}

:global(.history-filter-select-panel .ant-select-tree-treenode) {
  align-items: center;
  min-height: calc(26px * var(--app-scale));
  padding: calc(1px * var(--app-scale)) 0;
}

:global(.history-filter-select-panel .ant-select-tree-node-content-wrapper) {
  min-height: calc(26px * var(--app-scale));
  padding-inline: calc(4px * var(--app-scale));
  border-radius: calc(4px * var(--app-scale));
  color: var(--ink);
  line-height: calc(26px * var(--app-scale));
}

:global(.history-filter-select-panel .ant-select-tree-node-content-wrapper:hover) {
  background: rgba(45, 117, 214, 0.08);
}

:global(.history-filter-select-panel .ant-select-tree-switcher) {
  width: calc(18px * var(--app-scale));
  line-height: calc(26px * var(--app-scale));
}

:global(.history-filter-select-panel .ant-select-tree-indent-unit) {
  width: calc(14px * var(--app-scale));
}

:global(.history-filter-select-panel .ant-select-tree-checkbox) {
  margin-inline-end: calc(6px * var(--app-scale));
}

:global(.history-filter-select-panel .ant-select-tree-checkbox-inner) {
  width: calc(15px * var(--app-scale));
  height: calc(15px * var(--app-scale));
}

:global(.history-filter-select-panel .ant-select-tree-checkbox-checked .ant-select-tree-checkbox-inner) {
  border-color: #2d75d6;
  background: #2d75d6;
}

:global(.history-date-range-picker-panel .ant-picker-panel-container) {
  zoom: var(--app-scale);
}

.history-table-wrap {
  position: relative;
  --history-header-height: 32px;
  --history-body-height: 480px;
  flex: 1 1 0;
  height: auto;
  min-height: 0;
  margin-top: 6px;
  border: 1px solid rgba(104, 141, 181, 0.18);
  border-radius: 7px;
  background: rgba(255, 255, 255, 0.38);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.58);
  overflow: hidden;
}

.history-list :deep(.history-ant-table) {
  height: calc(var(--history-header-height) + var(--history-body-height));
  color: var(--ink);
  background: transparent;
  font-size: 14px;
  font-weight: 400;
}

.history-list :deep(.history-ant-table .ant-table) {
  width: 100%;
  color: var(--ink);
  background: transparent;
  font-size: 14px;
  font-weight: 400;
}

.history-list :deep(.history-ant-table .ant-table-container),
.history-list :deep(.history-ant-table .ant-table-content),
.history-list :deep(.history-ant-table table) {
  width: 100%;
  background: transparent;
}

.history-list :deep(.history-ant-table .ant-table-thead > tr > th) {
  height: var(--history-header-height);
  padding: 0 8px;
  border-color: var(--line-soft);
  color: var(--ink-strong);
  background: rgba(234, 244, 251, 0.54);
  font-size: 14px;
  font-weight: 500;
  line-height: var(--history-header-height);
}

.history-list :deep(.history-ant-table .ant-table-tbody > tr) {
  height: var(--history-row-height, 32px);
}

.history-list :deep(.history-ant-table .ant-table-tbody > tr > td) {
  height: var(--history-row-height, 32px);
  padding: 0 8px;
  border-color: var(--line-soft);
  color: var(--ink);
  background: transparent;
  font-size: 14px;
  font-weight: 400;
  line-height: 1.2;
}

.history-list :deep(.history-ant-table .ant-table-tbody > tr:hover > td) {
  background: rgba(74, 129, 183, 0.08);
}

.history-list :deep(.history-ant-table .history-placeholder-row > td) {
  color: transparent;
  background: transparent;
  pointer-events: none;
}

.history-list :deep(.history-ant-table .history-state-row > td) {
  height: var(--history-body-height);
  padding: 0 8px;
  border-color: transparent;
  background: transparent;
}

.history-list :deep(.history-ant-table .ant-table-tbody > tr > td.ant-table-cell-ellipsis) {
  overflow: hidden;
}

.history-list :deep(.history-ant-table .history-status-tag) {
  min-width: 70px;
  margin: 0;
  border: 1px solid transparent;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 500;
  line-height: 22px;
}

.history-list :deep(.history-ant-table .history-status-tag.status-success) {
  border-color: rgba(31, 143, 105, 0.16);
  color: var(--green);
  background: rgba(31, 143, 105, 0.14);
}

.history-list :deep(.history-ant-table .history-status-tag.status-warning) {
  border-color: rgba(223, 122, 53, 0.18);
  color: var(--orange);
  background: rgba(223, 122, 53, 0.15);
}

.history-list :deep(.history-ant-table .history-status-tag.status-danger) {
  border-color: rgba(217, 65, 63, 0.16);
  color: var(--red);
  background: rgba(217, 65, 63, 0.13);
}

.history-list :deep(.history-ant-table .history-status-tag.status-neutral) {
  border-color: rgba(104, 141, 181, 0.18);
  color: var(--ink-muted);
  background: rgba(104, 141, 181, 0.12);
}

.history-title-cell {
  display: block;
  width: 100%;
  overflow: hidden;
  color: var(--ink-strong);
  text-align: left;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.history-account-cell {
  display: block;
  max-width: 10em;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.history-table-state {
  display: grid;
  place-items: center;
  min-height: 120px;
  color: var(--ink);
  font-size: 14px;
  font-weight: 500;
}

.history-table-state.error {
  color: var(--red);
}
.history-title-cell {
  display: block;
  width: 100%;
  overflow: hidden;
  color: var(--ink-strong);
  text-align: left;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.history-account-cell {
  display: block;
  max-width: 10em;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.history-table-state {
  display: grid;
  place-items: center;
  min-height: 120px;
  color: var(--ink);
  font-size: 14px;
  font-weight: 500;
}

.history-table-state.error {
  color: var(--red);
}

.record-content-row {
  align-items: start;
}

.task-detail .detail-list {
  gap: 8px;
  margin-top: 14px;
}

.task-detail .detail-row {
  grid-template-columns: 82px minmax(0, 1fr);
  gap: 8px;
  font-size: 13px;
}

.task-detail .detail-row strong {
  min-width: 0;
  font-weight: 500;
}

.detail-value-wrap {
  overflow-wrap: anywhere;
  white-space: normal;
}

.detail-row .record-content {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 6px;
  overflow: visible;
  white-space: normal;
}

.record-content span {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 22px;
  padding: 0 6px;
  border: 1px solid rgba(104, 141, 181, 0.2);
  border-radius: 6px;
  color: var(--ink);
  background: rgba(255, 255, 255, 0.36);
  font-size: 11px;
  font-weight: 400;
}

.history-view-link {
  padding: 0 1px;
  color: var(--blue);
  font-size: 13px;
  font-weight: 500;
  line-height: 1.2;
}

.history-view-link {
  padding: 0 1px;
  color: var(--blue);
  font-size: 13px;
  font-weight: 500;
  line-height: 1.2;
}

.history-view-link:hover,
.history-view-link:focus-visible {
  color: #245eaf;
}

.history-ant-pagination {
  display: flex;
  flex: 0 0 auto;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  margin-top: 6px;
  min-height: 31px;
  color: var(--ink);
  font-size: 13px;
  font-weight: 400;
}

.history-total {
  flex: 0 0 auto;
  color: var(--ink);
  white-space: nowrap;
}

.history-pagination-controls {
  display: inline-flex;
  align-items: center;
  justify-content: flex-end;
  gap: 10px;
  min-width: 0;
}

.history-list :deep(.history-ant-pager.ant-pagination) {
  margin: 0;
  color: var(--ink);
  font-size: 13px;
  font-weight: 400;
}

.history-list :deep(.history-ant-pager .ant-pagination-item),
.history-list :deep(.history-ant-pager .ant-pagination-prev),
.history-list :deep(.history-ant-pager .ant-pagination-next) {
  min-width: 28px;
  height: 28px;
  line-height: 26px;
}

.history-list :deep(.history-ant-pager .ant-pagination-item),
.history-list :deep(.history-ant-pager .ant-pagination-prev .ant-pagination-item-link),
.history-list :deep(.history-ant-pager .ant-pagination-next .ant-pagination-item-link) {
  border-color: transparent;
  border-radius: 6px;
  color: var(--blue);
  background: rgba(255, 255, 255, 0.64);
}

.history-list :deep(.history-ant-pager .ant-pagination-item:hover),
.history-list :deep(.history-ant-pager .ant-pagination-prev:hover .ant-pagination-item-link),
.history-list :deep(.history-ant-pager .ant-pagination-next:hover .ant-pagination-item-link) {
  border-color: rgba(45, 117, 214, 0.24);
  background: rgba(238, 247, 255, 0.9);
}

.history-list :deep(.history-ant-pager .ant-pagination-item-active) {
  border-color: rgba(45, 117, 214, 0.32);
  background: linear-gradient(135deg, #4d85dc, #2d70cc);
}

.history-list :deep(.history-ant-pager .ant-pagination-item-active a) {
  color: #ffffff;
}

.history-list :deep(.history-ant-pager .ant-pagination-disabled .ant-pagination-item-link) {
  color: rgba(77, 108, 159, 0.36);
  background: rgba(255, 255, 255, 0.46);
}

.history-page-size {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 82px;
  height: 30px;
  padding: 0 12px;
  border: 1px solid var(--line);
  border-radius: 7px;
  color: var(--ink);
  background: rgba(255, 255, 255, 0.46);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.62);
  font-size: 13px;
  font-weight: 400;
  white-space: nowrap;
}
.detail-art {
  top: 0;
  right: 0;
  width: 118px;
}

.detail-empty {
  display: grid;
  justify-items: start;
  gap: 10px;
  margin-top: 28px;
  padding: 18px;
  border: 1px dashed rgba(104, 141, 181, 0.34);
  border-radius: 8px;
  color: var(--ink);
  background: rgba(255, 255, 255, 0.28);
}

.detail-empty i {
  display: grid;
  place-items: center;
  width: 34px;
  height: 34px;
  border-radius: 50%;
  color: #ffffff;
  background: linear-gradient(135deg, #74aefa, #2d75d6);
  font-size: 15px;
}

.detail-empty strong {
  color: var(--ink-strong);
  font-size: 15px;
  font-weight: 500;
}

.detail-empty span {
  color: var(--ink);
  line-height: 1.62;
  font-size: 13px;
  font-weight: 400;
}

.chart-caption {
  margin: 12px 0 6px;
  color: var(--ink);
  font-size: 14px;
  font-weight: 500;
}

.trend-chart {
  position: relative;
  display: grid;
  grid-template-columns: repeat(7, minmax(0, 1fr));
  align-items: end;
  height: 112px;
  padding: 8px 10px 0;
}

.trend-item {
  position: relative;
  z-index: 1;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: flex-end;
  gap: 7px;
  height: 100%;
}

.trend-bar {
  display: block;
  width: 24px;
  border-radius: 5px 5px 0 0;
  background:
    linear-gradient(180deg, rgba(255, 255, 255, 0.34), transparent 34%),
    linear-gradient(180deg, #79aee8, #2d75d6);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.48);
  transition:
    filter 150ms ease,
    transform 150ms ease;
}

.trend-bar:hover {
  filter: saturate(1.08) brightness(1.04);
  transform: translateY(-2px);
}

.trend-item small {
  line-height: 1;
  color: var(--ink);
  font-size: 12px;
  font-weight: 500;
}

.summary-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 8px;
  margin-top: 16px;
}

.summary-grid div {
  display: grid;
  justify-items: center;
  gap: 3px;
  min-height: 60px;
  padding: 8px 6px;
  border: 1px solid var(--line-soft);
  border-radius: 7px;
  background: rgba(255, 255, 255, 0.34);
  box-shadow: var(--paper-shadow-sm);
}

.summary-grid i {
  color: var(--green);
  font-size: 19px;
}

.summary-grid div:nth-child(2) i,
.summary-grid div:nth-child(2) strong {
  color: var(--red);
}

.summary-grid div:nth-child(3) i,
.summary-grid div:nth-child(3) strong {
  color: var(--blue);
}

.summary-grid strong {
  color: var(--green);
  font-size: 19px;
  font-weight: 500;
}

.summary-grid span {
  color: var(--ink);
  font-size: 12px;
  font-weight: 400;
}

.history-log {
  grid-area: log;
}

.log-art {
  right: 0;
  bottom: 0;
  width: 250px;
}

.log-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 18px;
}

.log-buttons {
  display: inline-flex;
  gap: 10px;
}

.history-log-table {
  display: grid;
  gap: 8px;
  margin-top: 14px;
}

.history-log-row {
  display: grid;
  grid-template-columns: 86px 70px minmax(0, 1fr);
  gap: 16px;
  color: var(--ink);
  font-size: 14px;
  font-weight: 700;
}

.history-log-row time,
.history-log-row strong {
  font-family: Consolas, 'SFMono-Regular', monospace;
}

.history-log-row time {
  color: var(--ink-muted);
}

.history-log-row strong {
  color: var(--green);
}

.history-log-row strong.warn {
  color: var(--orange);
}

.history-log-row strong.success {
  color: #16a15d;
}

.page-size {
  width: 96px;
  height: 32px;
}

:global(.collector-app.dark) .history-list :deep(.history-ant-control),
:global(.collector-app.dark) .history-list :deep(.history-date-picker.ant-picker),
:global(.collector-app.dark) .history-refresh-button {
  border-color: rgba(128, 153, 188, 0.2);
  color: #cbd8ea;
  background: rgba(15, 24, 39, 0.62);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.045);
}

:global(.collector-app.dark) .history-list :deep(.history-ant-select.ant-select .ant-select-selector) {
  border-color: rgba(128, 153, 188, 0.2);
  color: #cbd8ea;
  background: rgba(15, 24, 39, 0.56);
}

:global(.collector-app.dark) .history-list :deep(.history-ant-select.ant-select .ant-select-selection-item),
:global(.collector-app.dark) .history-list :deep(.history-ant-select.ant-select .ant-select-selection-placeholder),
:global(.collector-app.dark) .history-list :deep(.history-date-picker.ant-picker input),
:global(.collector-app.dark) .history-list :deep(.history-date-picker.ant-picker .ant-picker-input > input) {
  color: #cbd8ea;
}

:global(.collector-app.dark) .history-list :deep(.history-ant-select.ant-select .ant-select-selection-placeholder),
:global(.collector-app.dark) .history-list :deep(.history-date-picker.ant-picker input::placeholder) {
  color: rgba(142, 162, 189, 0.72);
}

:global(.collector-app.dark) .history-search-icon,
:global(.collector-app.dark) .history-select-chevron {
  color: #9fc3ef;
}

:global(.collector-app.dark) .history-refresh-button {
  border-color: rgba(128, 153, 188, 0.2);
  color: #8fbded;
}

:global(.collector-app.dark) .history-refresh-button:hover,
:global(.collector-app.dark) .history-refresh-button:focus-visible {
  border-color: rgba(111, 154, 211, 0.34);
  color: #dceaff;
  background: rgba(24, 38, 60, 0.86);
}

:global(.collector-app.dark) .history-table-wrap {
  border-color: rgba(128, 153, 188, 0.14);
  background: rgba(15, 24, 39, 0.52);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.045);
}

:global(.collector-app.dark) .history-list :deep(.history-ant-table .ant-table),
:global(.collector-app.dark) .history-list :deep(.history-ant-table .ant-table-container),
:global(.collector-app.dark) .history-list :deep(.history-ant-table .ant-table-content),
:global(.collector-app.dark) .history-list :deep(.history-ant-table table) {
  color: #c5d3e6;
  background: transparent;
}

:global(.collector-app.dark) .history-list :deep(.history-ant-table .ant-table-thead > tr > th) {
  color: #dce7f5;
  border-color: rgba(128, 153, 188, 0.12);
  background: rgba(24, 37, 58, 0.86);
}

:global(.collector-app.dark) .history-list :deep(.history-ant-table .ant-table-tbody > tr > td) {
  color: #c5d3e6;
  border-color: rgba(128, 153, 188, 0.1);
  background: transparent;
}

:global(.collector-app.dark) .history-list :deep(.history-ant-table .ant-table-tbody > tr:hover > td) {
  background: rgba(36, 56, 84, 0.32);
}

:global(.collector-app.dark) .history-title-cell,
:global(.collector-app.dark) .history-account-cell {
  color: #dce7f5;
}

:global(.collector-app.dark) .history-table-state,
:global(.collector-app.dark) .history-total,
:global(.collector-app.dark) .history-page-size {
  color: #8ea2bd;
}

:global(.collector-app.dark) .history-list :deep(.history-view-link) {
  color: #8fbded;
}

:global(.collector-app.dark) .history-list :deep(.history-ant-pager .ant-pagination-item),
:global(.collector-app.dark) .history-list :deep(.history-ant-pager .ant-pagination-prev .ant-pagination-item-link),
:global(.collector-app.dark) .history-list :deep(.history-ant-pager .ant-pagination-next .ant-pagination-item-link),
:global(.collector-app.dark) .history-page-size {
  color: #a9bfda;
  border-color: rgba(128, 153, 188, 0.16);
  background: rgba(17, 27, 44, 0.72);
  box-shadow: inset 0 1px 0 rgba(214, 226, 244, 0.04);
}

:global(.collector-app.dark) .history-list :deep(.history-ant-pager .ant-pagination-item:hover),
:global(.collector-app.dark) .history-list :deep(.history-ant-pager .ant-pagination-prev:hover .ant-pagination-item-link),
:global(.collector-app.dark) .history-list :deep(.history-ant-pager .ant-pagination-next:hover .ant-pagination-item-link) {
  color: #dceaff;
  border-color: rgba(111, 154, 211, 0.34);
  background: rgba(24, 38, 60, 0.86);
}

:global(.collector-app.dark) .history-list :deep(.history-ant-pager .ant-pagination-item-active) {
  border-color: rgba(111, 154, 211, 0.36);
  background: #2f6fb5;
}

:global(.collector-app.dark) .history-list :deep(.history-ant-pager .ant-pagination-item-active a) {
  color: #f0f6ff;
}

:global(.collector-app.dark) .history-list :deep(.history-ant-pager .ant-pagination-disabled .ant-pagination-item-link) {
  color: rgba(142, 162, 189, 0.42);
  background: rgba(17, 27, 44, 0.42);
}
:global(.collector-app.dark) .task-detail .detail-row,
:global(.collector-app.dark) .trend-item,
:global(.collector-app.dark) .history-log-row {
  border-color: rgba(128, 153, 188, 0.14);
  background: rgba(15, 24, 39, 0.5);
}

:global(.collector-app.dark) .detail-empty {
  color: #8ea2bd;
  background: rgba(15, 24, 39, 0.36);
}

:global(.collector-app.dark) .detail-empty strong,
:global(.collector-app.dark) .task-detail .detail-row strong,
:global(.collector-app.dark) .detail-row .record-content {
  color: #dce7f5;
}

:global(.collector-app.dark) .trend-bar {
  background:
    linear-gradient(180deg, rgba(214, 226, 244, 0.14), transparent 34%),
    linear-gradient(180deg, #4d8ed2, #2f6fb5);
}

:global(.collector-app.dark) .chart-caption,
:global(.collector-app.dark) .trend-item small,
:global(.collector-app.dark) .history-log-row time {
  color: #8ea2bd;
}
</style>
