type DesktopEnvironment = {
  appVersion?: string
  systemLabel?: string
  pythonVersion?: string
  mitmproxyVersion?: string
  playwrightVersion?: string
  desktopShell?: string
  desktopShellStatus?: string
}

type DesktopStatusPayload = {
  ok: boolean
  status: string
  environment?: DesktopEnvironment
}

export type EnvironmentStatus = {
  systemLabel: string
  appVersion: string
  pythonVersion: string
  mitmproxyVersion: string
  playwrightVersion: string
  desktopShellStatus: string
}

export type ResolvedDesktopEnvironmentStatus = {
  shouldRetry: boolean
  desktopStatusLabel: string
  environmentStatus: EnvironmentStatus
}

export const INITIAL_ENVIRONMENT_STATUS: EnvironmentStatus = {
  systemLabel: '检测中',
  appVersion: '检测中',
  pythonVersion: '检测中',
  mitmproxyVersion: '检测中',
  playwrightVersion: '检测中',
  desktopShellStatus: '检测中',
}

function formatDesktopShellStatus(
  source: DesktopEnvironment | undefined,
  fallbackLabel: string,
): string {
  if (!source) {
    return fallbackLabel
  }

  const shell = String(source.desktopShell || 'tauri').trim().toLowerCase()
  const shellLabel = shell === 'tauri' ? 'Tauri' : shell || 'Tauri'
  const statusLabel = source.desktopShellStatus || fallbackLabel
  return `${shellLabel} ${statusLabel}`.trim()
}

function buildEnvironmentStatus(
  source: DesktopEnvironment | undefined,
  fallbackLabel: string,
): EnvironmentStatus {
  return {
    systemLabel: source?.systemLabel || '未知系统',
    appVersion: source?.appVersion || '未设置',
    pythonVersion: source?.pythonVersion || '未知版本',
    mitmproxyVersion: source?.mitmproxyVersion || '未知版本',
    playwrightVersion: source?.playwrightVersion || '未知版本',
    desktopShellStatus: formatDesktopShellStatus(source, fallbackLabel),
  }
}

export function resolveDesktopEnvironmentStatus(
  status: DesktopStatusPayload,
): ResolvedDesktopEnvironmentStatus {
  if (status.status === 'browser-preview') {
    return {
      shouldRetry: true,
      desktopStatusLabel: '检测中',
      environmentStatus: { ...INITIAL_ENVIRONMENT_STATUS },
    }
  }

  const desktopStatusLabel = status.ok ? '后端已连接' : '浏览器预览'
  return {
    shouldRetry: false,
    desktopStatusLabel,
    environmentStatus: buildEnvironmentStatus(status.environment, desktopStatusLabel),
  }
}

export function getBrowserPreviewEnvironmentStatus(): ResolvedDesktopEnvironmentStatus {
  const desktopStatusLabel = '浏览器预览'
  return {
    shouldRetry: false,
    desktopStatusLabel,
    environmentStatus: buildEnvironmentStatus(undefined, desktopStatusLabel),
  }
}

export function getEnvironmentErrorStatus(): ResolvedDesktopEnvironmentStatus {
  return {
    shouldRetry: false,
    desktopStatusLabel: '连接失败',
    environmentStatus: {
      systemLabel: '读取失败',
      appVersion: '读取失败',
      pythonVersion: '读取失败',
      mitmproxyVersion: '读取失败',
      playwrightVersion: '读取失败',
      desktopShellStatus: '读取失败',
    },
  }
}
