import type { TaskStatus } from '../bridge/pythonApi'
import type { EnvironmentStatus } from './desktopStatus'

export type SettingsMetricTone = 'blue' | 'green' | 'red' | 'purple' | 'orange'

export type SettingsRuntimeMetric = {
  label: string
  value: string
  icon: string
  tone: SettingsMetricTone
  hint?: string
}

const WINDOW_DETECTION_ICON = 'fa-solid fa-desktop'

export function buildSettingsRuntimeMetrics(
  status?: Partial<TaskStatus> | null,
  environmentStatus?: Partial<EnvironmentStatus> | null,
): SettingsRuntimeMetric[] {
  return [
    {
      label: '应用版本',
      value: normalizeAppVersion(environmentStatus?.appVersion),
      icon: 'fa-solid fa-code-branch',
      tone: 'blue',
      hint: '读取自程序运行环境配置',
    },
    resolveProgramMetric(status),
    resolveWindowMetric(status),
    resolveAuthMetric(status),
  ]
}

function normalizeAppVersion(value: unknown): string {
  const text = String(value ?? '').trim()
  if (!text) {
    return '检测中'
  }
  return text
}

function resolveProgramMetric(status?: Partial<TaskStatus> | null): SettingsRuntimeMetric {
  const runtimeStatus = String(status?.status || 'idle')
  const map: Record<string, Pick<SettingsRuntimeMetric, 'value' | 'tone' | 'hint'>> = {
    idle: { value: '待机', tone: 'orange', hint: '当前没有采集任务运行' },
    starting: { value: '启动中', tone: 'blue', hint: '正在准备采集环境' },
    running: { value: '采集中', tone: 'green', hint: '主服务正在执行采集任务' },
    stopped: { value: '已停止', tone: 'orange', hint: '当前没有采集任务运行' },
    error: { value: '异常', tone: 'red', hint: status?.message || '请查看运行日志' },
    'browser-preview': { value: '浏览器预览', tone: 'blue', hint: '当前处于前端预览环境' },
  }
  const resolved = map[runtimeStatus] ?? { value: runtimeStatus || '未知', tone: 'orange' as const, hint: '运行状态未识别' }

  return {
    label: '程序状态',
    value: resolved.value,
    icon: 'fa-solid fa-heart-pulse',
    tone: resolved.tone,
    hint: resolved.hint,
  }
}

function resolveWindowMetric(status?: Partial<TaskStatus> | null): SettingsRuntimeMetric {
  const home = status?.home
  const homeStatus = String(home?.status || '')
  const statusLabel = String(home?.statusLabel || '').trim()
  const accountName = String(home?.accountName || '').trim()
  const value = statusLabel || (home?.found ? '已检测到主页窗口' : '未检测到主页窗口')

  if (homeStatus === 'ready' || home?.found) {
    return {
      label: '窗口检测',
      value,
      icon: WINDOW_DETECTION_ICON,
      tone: 'green',
      hint: accountName || '已读取主页信息',
    }
  }

  if (homeStatus === 'display_cached') {
    return {
      label: '窗口检测',
      value,
      icon: WINDOW_DETECTION_ICON,
      tone: 'orange',
      hint: accountName || '沿用上次识别信息',
    }
  }

  if (['not_found', 'failed', 'dependency_missing', 'content_unreadable', 'display_unavailable'].includes(homeStatus)) {
    return {
      label: '窗口检测',
      value,
      icon: WINDOW_DETECTION_ICON,
      tone: 'red',
      hint: home?.message || '请打开微信 PC 公众号主页',
    }
  }

  return {
    label: '窗口检测',
    value: value || '检测中',
    icon: WINDOW_DETECTION_ICON,
    tone: 'blue',
    hint: '等待主页窗口检测结果',
  }
}

function resolveAuthMetric(status?: Partial<TaskStatus> | null): SettingsRuntimeMetric {
  const auth = status?.auth
  const proxy = status?.proxy
  const mitmRunning = Boolean(proxy?.mitmEnabled)

  if (auth?.hasKeyUrl) {
    return {
      label: '鉴权状态',
      value: auth.statusLabel || '已获取鉴权',
      icon: 'fa-solid fa-certificate',
      tone: 'green',
      hint: '已捕获带 key 的文章 URL',
    }
  }

  if (mitmRunning || auth?.status === 'waiting') {
    return {
      label: '鉴权状态',
      value: auth?.statusLabel || '等待鉴权',
      icon: 'fa-solid fa-certificate',
      tone: 'orange',
      hint: '打开文章后等待 MITM 捕获 key URL',
    }
  }

  return {
    label: '鉴权状态',
    value: auth?.statusLabel || '未启动代理',
    icon: 'fa-solid fa-certificate',
    tone: 'orange',
    hint: '启动 MITM 后打开文章获取 key',
  }
}
