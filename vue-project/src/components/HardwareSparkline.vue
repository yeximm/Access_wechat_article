<script setup lang="ts">
import { Chart } from '@antv/g2'
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import type { HardwareHistoryPoint } from '../bridge/pythonApi'

type HardwareSeries = 'CPU' | '物理内存'

type HardwareChartDatum = {
  elapsedSeconds: number
  time: string
  type: HardwareSeries
  value: number
  displayValue: number
  shape: 'smooth'
}

type HardwareLane = {
  base: number
  amplitude: number
}

type HardwareValueDomain = {
  min: number
  max: number
}

const CHART_HEIGHT = 62
const MIN_VISIBLE_RANGE_PERCENT = 1

// 两条固定泳道避免 CPU / 内存百分比接近时完全重叠，图表只表达趋势。
const CPU_LANE: HardwareLane = {
  base: 46,
  amplitude: 12,
}

const MEMORY_LANE: HardwareLane = {
  base: 16,
  amplitude: 12,
}

const props = defineProps<{
  points: HardwareHistoryPoint[]
}>()

const chartContainerRef = ref<HTMLDivElement | null>(null)
let chart: InstanceType<typeof Chart> | null = null

const normalizedPoints = computed<HardwareHistoryPoint[]>(() => {
  const points = props.points?.slice(-40) ?? []
  if (points.length >= 2) {
    return points
  }

  if (points.length === 1) {
    return [buildEmptyPoint(0), points[0]!]
  }

  return Array.from({ length: 12 }, (_, index) => buildEmptyPoint(index))
})

const chartData = computed<HardwareChartDatum[]>(() => {
  const points = normalizedPoints.value
  const firstTimestamp = resolvePointTimestamp(points[0], 0)
  const cpuDomain = resolveAdaptiveDomain(points.map((point) => point.cpuPercent))
  const memoryDomain = resolveAdaptiveDomain(points.map((point) => point.memoryPercent))
  return points.flatMap((point, index) => {
    const cpu = normalizePercent(point.cpuPercent)
    const memory = normalizePercent(point.memoryPercent)
    const time = resolvePointTime(point, index)
    const elapsedSeconds = resolveElapsedSeconds(point, index, firstTimestamp)
    return [
      {
        elapsedSeconds,
        time,
        type: 'CPU' as const,
        value: cpu,
        displayValue: resolveLaneValue(cpu, CPU_LANE, cpuDomain),
        shape: 'smooth' as const,
      },
      {
        elapsedSeconds,
        time,
        type: '物理内存' as const,
        value: memory,
        displayValue: resolveLaneValue(memory, MEMORY_LANE, memoryDomain),
        shape: 'smooth' as const,
      },
    ]
  })
})

function buildEmptyPoint(index: number): HardwareHistoryPoint {
  return {
    timestamp: index,
    cpuPercent: 0,
    memoryPercent: 0,
  }
}

function resolvePointTimestamp(point: HardwareHistoryPoint | undefined, index: number) {
  if (!point) {
    return index * 1000
  }
  if (typeof point.timestamp === 'string') {
    const date = new Date(point.timestamp)
    if (!Number.isNaN(date.getTime())) {
      return date.getTime()
    }
  }
  const numericTimestamp = Number(point.timestamp)
  if (Number.isFinite(numericTimestamp)) {
    return numericTimestamp
  }
  return index * 1000
}

function resolveElapsedSeconds(point: HardwareHistoryPoint, index: number, firstTimestamp: number) {
  const timestamp = resolvePointTimestamp(point, index)
  const elapsed = (timestamp - firstTimestamp) / 1000
  if (!Number.isFinite(elapsed) || elapsed < 0) {
    return index
  }
  return elapsed
}

function resolvePointTime(point: HardwareHistoryPoint, index: number) {
  if (typeof point.timestamp === 'string') {
    const date = new Date(point.timestamp)
    if (!Number.isNaN(date.getTime())) {
      return date.toLocaleTimeString('zh-CN', { hour12: false })
    }
    return point.timestamp
  }
  return String(index)
}

function normalizePercent(value: unknown) {
  const percent = Number(value)
  if (!Number.isFinite(percent)) {
    return 0
  }
  return Math.min(100, Math.max(0, percent))
}

function resolveAdaptiveDomain(values: unknown[]): HardwareValueDomain {
  const normalizedValues = values.map(normalizePercent)
  if (normalizedValues.length === 0) {
    return {
      min: 0,
      max: MIN_VISIBLE_RANGE_PERCENT,
    }
  }

  const min = Math.min(...normalizedValues)
  const max = Math.max(...normalizedValues)
  const actualRange = max - min
  const visibleRange = Math.max(actualRange, MIN_VISIBLE_RANGE_PERCENT)
  const center = (min + max) / 2

  // 小占比场景下，真实数值仍显示在右侧；折线只把近期波动放大到可视范围内。
  let domainMin = center - visibleRange / 2
  let domainMax = center + visibleRange / 2

  if (domainMin < 0) {
    domainMax = Math.min(100, domainMax - domainMin)
    domainMin = 0
  }

  if (domainMax > 100) {
    domainMin = Math.max(0, domainMin - (domainMax - 100))
    domainMax = 100
  }

  return {
    min: domainMin,
    max: domainMax,
  }
}

function resolveLaneValue(percent: number, lane: HardwareLane, domain: HardwareValueDomain) {
  const value = normalizePercent(percent)
  const range = Math.max(domain.max - domain.min, MIN_VISIBLE_RANGE_PERCENT)
  const ratio = Math.min(1, Math.max(0, (value - domain.min) / range))
  return lane.base + ratio * lane.amplitude
}

function resolveChartColors() {
  const styleSource = chartContainerRef.value ?? document.documentElement
  const styles = window.getComputedStyle(styleSource)
  return {
    cpu: styles.getPropertyValue('--green').trim() || '#1f8f69',
    memory: styles.getPropertyValue('--blue').trim() || '#2d75d6',
  }
}

function buildChartOptions() {
  const colors = resolveChartColors()
  return {
    type: 'line',
    autoFit: true,
    data: chartData.value,
    padding: 0,
    margin: 0,
    inset: 0,
    encode: {
      x: 'elapsedSeconds',
      y: 'displayValue',
      color: 'type',
      shape: 'smooth',
    },
    scale: {
      x: {
        nice: false,
        padding: 0,
      },
      y: {
        domainMin: 0,
        domainMax: CHART_HEIGHT,
        nice: false,
      },
      color: {
        domain: ['CPU', '物理内存'],
        range: [colors.cpu, colors.memory],
      },
    },
    axis: false,
    legend: false,
    tooltip: false,
    style: {
      lineWidth: 2.6,
      lineCap: 'round',
      lineJoin: 'round',
    },
    animate: false,
  }
}

async function renderChart() {
  await nextTick()
  if (!chartContainerRef.value) {
    return
  }

  destroyChart()
  chart = new Chart({
    container: chartContainerRef.value,
    autoFit: true,
    height: CHART_HEIGHT,
  })
  chart.options(buildChartOptions())
  await chart.render()
}

function destroyChart() {
  if (!chart) {
    return
  }
  chart.destroy()
  chart = null
}

watch(
  chartData,
  (data) => {
    // 运行状态轮询只更新数据，不重建图表，避免折线闪烁。
    void chart?.changeData(data)
  },
  { deep: true },
)

onMounted(() => {
  void renderChart()
})

onBeforeUnmount(() => {
  destroyChart()
})
</script>

<template>
  <div
    ref="chartContainerRef"
    class="hardware-sparkline"
    role="img"
    aria-label="当前 CPU 和物理内存占比趋势"
  ></div>
</template>

<style scoped>
.hardware-sparkline {
  display: block;
  width: 100%;
  min-width: 0;
  height: 62px;
  overflow: hidden;
}
</style>
