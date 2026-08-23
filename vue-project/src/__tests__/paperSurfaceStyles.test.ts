import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import test from 'node:test'
import { fileURLToPath } from 'node:url'

const currentDir = dirname(fileURLToPath(import.meta.url))
const appVue = readFileSync(resolve(currentDir, '../App.vue'), 'utf8')
const appTopbarVue = readFileSync(resolve(currentDir, '../components/AppTopbar.vue'), 'utf8')
const pagesCss = readFileSync(resolve(currentDir, '../styles/pages.css'), 'utf8')
const settingsPageVue = readFileSync(resolve(currentDir, '../pages/SettingsPage.vue'), 'utf8')

function extractRule(source: string, selector: string) {
  const normalizedSource = source.replace(/\r\n/g, '\n')
  const escaped = selector.replace(/[.*+?^\${}()|[\]\\]/g, '\\$&')
  const match = normalizedSource.match(new RegExp(escaped + '\\s*\\{[\\s\\S]*?\\n\\}'))

  assert.ok(match, selector + ' rule should exist')

  return match[0]
}

test('全局外壳提供统一轻纸质感变量，供首页和管理页复用', () => {
  assert.match(appVue, /--paper-texture-opacity:\s*0\.045;/)
  assert.match(appVue, /--paper-card-texture-opacity:\s*0\.045;/)
  assert.match(appVue, /--paper-edge:\s*rgba\(102, 145, 184, 0\.26\);/)
  assert.match(appVue, /--paper-shadow-sm:\s*0 1px 2px rgba\(28, 55, 82, 0\.06\);/)
  assert.match(appVue, /--paper-shadow-md:\s*0 4px 9px rgba\(28, 55, 82, 0\.06\),\s*0 8px 14px rgba\(28, 55, 82, 0\.045\);/)
  assert.match(appVue, /--paper-hover-shadow:\s*0 5px 10px rgba\(28, 55, 82, 0\.075\),\s*0 10px 15px rgba\(28, 55, 82, 0\.05\);/)
  assert.match(appVue, /--paper-fiber:/)
  assert.match(appVue, /--paper-texture-opacity:\s*0\.024;/)
  assert.match(appVue, /--paper-card-texture-opacity:\s*0\.022;/)
})

test('program status panel removes proxy row and stretches remaining content', () => {
  const statusCardRule = extractRule(appVue, '.status-card')
  const statusListRule = extractRule(appVue, '.status-list')
  const statusRowRule = extractRule(appVue, '.status-row')
  const statusProgressRowRule = extractRule(appVue, '.status-row.with-progress')
  const statusProgressValueRule = extractRule(appVue, '.status-row em')
  const networkPanelRule = extractRule(appVue, '.network-panel')
  const speedContentRule = extractRule(appVue, '.speed-content')
  const speedValuesRule = extractRule(appVue, '.speed-values')

  assert.doesNotMatch(appVue, /当前代理状态/)
  assert.doesNotMatch(appVue, /proxyDisplayStatus|resolveProxyDisplayStatus/)
  assert.match(statusCardRule, /display:\s*grid;/)
  assert.match(statusCardRule, /grid-template-rows:\s*auto minmax\(0, 1fr\) auto;/)
  assert.match(statusCardRule, /padding:\s*20px 16px 18px;/)
  assert.match(statusListRule, /align-content:\s*space-between;/)
  assert.match(statusListRule, /min-height:\s*0;/)
  assert.match(statusRowRule, /grid-template-columns:\s*18px 96px minmax\(0, 1fr\);/)
  assert.match(statusRowRule, /gap:\s*6px;/)
  assert.match(statusProgressRowRule, /grid-template-columns:\s*18px 96px minmax\(0, 1fr\) max-content;/)
  assert.match(statusProgressValueRule, /white-space:\s*nowrap;/)
  assert.match(networkPanelRule, /min-height:\s*82px;/)
  assert.match(networkPanelRule, /padding-top:\s*18px;/)
  assert.match(networkPanelRule, /grid-template-columns:\s*18px 96px minmax\(0, 1fr\);/)
  assert.match(speedContentRule, /grid-template-columns:\s*minmax\(0, 1fr\) 82px;/)
  assert.match(speedValuesRule, /justify-items:\s*start;/)
})

test('首页纸质风格降低过重字重，保留清晰层级', () => {
  const navTitleRule = extractRule(appVue, '.nav-title')
  const navItemRule = extractRule(appVue, '.nav-item')
  const headingRule = extractRule(appVue, '.task-body h2,\n.section-title h2')
  const taskBodyRule = extractRule(appVue, '.task-body p')
  const fieldRowRule = extractRule(appVue, '.field-row')
  const runButtonRule = extractRule(appVue, '.run-button,\n.stop-button')
  const statusStrongRule = extractRule(appVue, '.status-row strong')
  const statLabelRule = extractRule(appVue, '.stat-item span')
  const statValueRule = extractRule(appVue, '.stat-item strong')
  const noticeListRule = extractRule(appVue, '.notice ul')
  const logTabsRule = extractRule(appVue, '.log-tabs button,\n.log-actions button')
  const logLevelRule = extractRule(appVue, '.log-level')

  assert.match(navTitleRule, /font-weight:\s*500;/)
  assert.match(navItemRule, /font-weight:\s*500;/)
  assert.match(headingRule, /font-weight:\s*500;/)
  assert.match(taskBodyRule, /font-weight:\s*400;/)
  assert.match(fieldRowRule, /font-weight:\s*400;/)
  assert.match(runButtonRule, /font-weight:\s*500;/)
  assert.match(statusStrongRule, /font-weight:\s*500;/)
  assert.match(statLabelRule, /font-weight:\s*400;/)
  assert.match(statValueRule, /font-weight:\s*400;/)
  assert.match(noticeListRule, /font-weight:\s*400;/)
  assert.match(logTabsRule, /font-weight:\s*400;/)
  assert.match(logLevelRule, /font-weight:\s*400;/)
})

test('home page uses HarmonyOS Sans SC while keeping legacy brand title style', () => {
  const bodyRule = extractRule(appVue, ':global(body)')
  const brandTitleRule = extractRule(appTopbarVue, '.brand h1')
  const brandSubtitleRule = extractRule(appTopbarVue, '.brand p')
  const healthLabelRule = extractRule(appTopbarVue, '.health-label')
  const healthValueRule = extractRule(appTopbarVue, '.health-value')
  const githubRule = extractRule(appTopbarVue, '.github-pill')

  assert.match(appVue, /@font-face\s*\{[\s\S]*font-family:\s*'HarmonyOS Sans SC';[\s\S]*HarmonyOS_Sans_SC_Regular\.ttf[\s\S]*font-weight:\s*400;/)
  assert.match(appVue, /@font-face\s*\{[\s\S]*font-family:\s*'HarmonyOS Sans SC';[\s\S]*HarmonyOS_Sans_SC_Medium\.ttf[\s\S]*font-weight:\s*500;/)
  assert.match(bodyRule, /font-family:[\s\S]*'HarmonyOS Sans SC'[\s\S]*sans-serif;/)
  assert.match(bodyRule, /font-synthesis-weight:\s*none;/)
  assert.match(brandTitleRule, /font-family:\s*Georgia, 'Times New Roman', 'Noto Serif SC', serif;/)

  for (const source of [appVue, appTopbarVue]) {
    assert.doesNotMatch(source, /font-weight:\s*[6-9]00;/)
  }

  for (const rule of [brandSubtitleRule, healthValueRule]) {
    assert.match(rule, /font-weight:\s*400;/)
  }

  for (const rule of [brandTitleRule, healthLabelRule, githubRule]) {
    assert.match(rule, /font-weight:\s*500;/)
  }
})

test('首页和管理页通用面板改为哑光纸质 surface，不再使用玻璃模糊', () => {
  const panelRule = extractRule(appVue, '.panel')
  const panelBeforeRule = extractRule(appVue, '.panel::before')
  const pagePanelRule = extractRule(pagesCss, '.page-panel')
  const pagePanelBeforeRule = extractRule(pagesCss, '.page-panel::before')

  for (const rule of [panelRule, pagePanelRule]) {
    assert.match(rule, /background:\s*var\(--frost-bg-strong\);/)
    assert.match(rule, /box-shadow:\s*var\(--paper-shadow-sm\),\s*var\(--paper-shadow-md\);/)
    assert.doesNotMatch(rule, /backdrop-filter/)
    assert.doesNotMatch(rule, /linear-gradient\(135deg, var\(--frost-highlight\)/)
  }

  for (const rule of [panelBeforeRule, pagePanelBeforeRule]) {
    assert.match(rule, /background:\s*var\(--paper-fiber\);/)
    assert.match(rule, /opacity:\s*var\(--paper-card-texture-opacity\);/)
  }
})

test('应用左侧主导航和系统配置三段式菜单都恢复绿色选中态', () => {
  const activeNavRule = extractRule(appVue, '.nav-item.active')
  const settingsMenuRule = extractRule(settingsPageVue, '.settings-primary-item,\n.settings-secondary-item')
  const settingsActiveRule = extractRule(settingsPageVue, '.settings-primary-item.active,\n.settings-secondary-item.active')

  assert.match(activeNavRule, /radial-gradient\(circle at 20% 10%, #ffffff4d, transparent 36%\)/)
  assert.match(activeNavRule, /linear-gradient\(135deg, #5bbdc5eb, #2b8ea6eb\)/)
  assert.doesNotMatch(activeNavRule, /#4eae9ceb|#2e8480eb|#2f8d82/)

  assert.match(settingsMenuRule, /--settings-menu-active-bg:\s*#E7F5F2;/)
  assert.match(settingsMenuRule, /--settings-menu-hover-bg:\s*#F4FBFA;/)
  assert.match(settingsActiveRule, /color:\s*var\(--settings-menu-active-ink\);/)
  assert.doesNotMatch(settingsActiveRule, /0, 114, 239|#DCEBFF/)
})

test('顶部状态块和工具按钮使用纸质资料标签 surface，避免强玻璃效果', () => {
  const healthItemRule = extractRule(appTopbarVue, '.health-item')
  const githubRule = extractRule(appTopbarVue, '.github-pill')
  const themeSwitchRule = extractRule(appTopbarVue, '.theme-switch')
  const topbarSurfaceRule = extractRule(
    appTopbarVue,
    '.health-item::before,\n.github-pill::before,\n.theme-switch::before',
  )
  const topbarSurfaceContentRule = extractRule(
    appTopbarVue,
    '.health-item > *,\n.github-pill > *,\n.theme-switch > *',
  )

  for (const rule of [healthItemRule, githubRule, themeSwitchRule]) {
    assert.match(rule, /background:\s*var\(--frost-bg-strong\);/)
    assert.match(rule, /box-shadow:\s*var\(--paper-shadow-sm\);/)
    assert.doesNotMatch(rule, /backdrop-filter|linear-gradient\(135deg/)
  }

  assert.match(topbarSurfaceRule, /background:\s*var\(--paper-fiber\);/)
  assert.match(topbarSurfaceRule, /opacity:\s*calc\(var\(--paper-card-texture-opacity\) \* 1\.25\);/)
  assert.match(topbarSurfaceRule, /mix-blend-mode:\s*multiply;/)
  assert.match(topbarSurfaceContentRule, /z-index:\s*1;/)
})

test('首页局部控件使用同一套轻纸质 surface，不再保留局部玻璃模糊', () => {
  const taskCountInputRule = extractRule(appVue, '.task-count-input')
  const downloadOptionRule = extractRule(appVue, '.download-option')
  const noticeRule = extractRule(appVue, '.notice')
  const envItemRule = extractRule(appVue, '.env-item')

  for (const rule of [taskCountInputRule, downloadOptionRule, noticeRule, envItemRule]) {
    assert.match(rule, /box-shadow:\s*var\(--paper-shadow-sm\);/)
    assert.doesNotMatch(rule, /backdrop-filter|linear-gradient\(135deg/)
  }
})

test('使用须知的操作说明和合规提示使用 Ant Design Card Meta 展示', () => {
  const rightColumnRule = extractRule(appVue, '.right-column')
  const guideCardRule = extractRule(appVue, '.guide-card')
  const envGridRule = extractRule(appVue, '.env-grid')
  const envItemRule = extractRule(appVue, '.env-item')
  const envHoverRule = extractRule(appVue, '.env-item.ant-card-hoverable:hover')
  const envCardBodyRule = extractRule(appVue, '.env-item :deep(.ant-card-body)')
  const envItemContentRule = extractRule(appVue, '.env-item-content')
  const noticeInfoRule = extractRule(appVue, '.notice.info')
  const noticeHoverRule = extractRule(appVue, '.notice.ant-card-hoverable:hover')
  const noticeMetaTitleRule = extractRule(appVue, '.notice :deep(.ant-card-meta-title)')
  const noticeMetaDescriptionRule = extractRule(appVue, '.notice :deep(.ant-card-meta-description)')

  assert.match(appVue, /<ACard\s+class="notice info"[^>]*hoverable[\s\S]*?<ACardMeta title="操作说明">[\s\S]*?<template #description>[\s\S]*?usageTips/)
  assert.match(appVue, /<ACard\s+class="notice warning"[^>]*hoverable[\s\S]*?<ACardMeta title="合规提示">[\s\S]*?<template #description>[\s\S]*?complianceTips/)
  assert.match(appVue, /<ACard\s+v-for="item in envItems"[\s\S]*?class="env-item"[\s\S]*?:bordered="false"[\s\S]*hoverable[\s\S]*?<div class="env-item-content">/)
  assert.doesNotMatch(appVue, /<div class="notice info">[\s\S]*?<h3>操作说明<\/h3>/)
  assert.doesNotMatch(appVue, /<div class="notice warning">[\s\S]*?<h3>合规提示<\/h3>/)
  assert.match(rightColumnRule, /grid-template-rows:\s*170px 344px 212px;/)
  assert.match(guideCardRule, /gap:\s*6px;/)
  assert.match(noticeInfoRule, /margin-top:\s*0;/)
  assert.match(envGridRule, /gap:\s*10px 8px;/)
  assert.match(envGridRule, /margin-top:\s*10px;/)
  assert.match(envItemRule, /min-height:\s*62px;/)
  assert.match(envItemRule, /transition:[\s\S]*?transform 180ms ease/)
  assert.match(envHoverRule, /transform:\s*translateY\(-2px\);/)
  assert.match(envHoverRule, /box-shadow:[\s\S]*?0 5px 14px rgba\(28, 55, 82, 0\.12\)/)
  assert.match(envCardBodyRule, /align-items:\s*center;/)
  assert.match(envCardBodyRule, /justify-content:\s*flex-start;/)
  assert.match(envItemContentRule, /justify-items:\s*start;/)
  assert.match(envItemContentRule, /text-align:\s*left;/)
  assert.match(noticeHoverRule, /transform:\s*translateY\(-2px\);/)
  assert.match(noticeHoverRule, /box-shadow:[\s\S]*?0 5px 14px rgba\(28, 55, 82, 0\.12\)/)
  assert.match(noticeMetaTitleRule, /font-weight:\s*500;/)
  assert.match(noticeMetaDescriptionRule, /color:\s*inherit;/)
})

test('home page action status stats and log controls use Ant Design Vue components', () => {
  assert.match(appVue, /<AButton\s+[\s\S]*?class="run-button"[\s\S]*?html-type="button"[\s\S]*?<AppIcon icon="fa-solid fa-play"/)
  assert.match(appVue, /<AButton\s+[\s\S]*?class="stop-button"[\s\S]*?html-type="button"[\s\S]*?<AppIcon icon="fa-solid fa-stop"/)
  assert.match(appVue, /<ATag\s+[\s\S]*?:class="\['status-pill', item\.tone\]"/)
  assert.match(appVue, /<AProgress\s+[\s\S]*?class="progress-line"/)
  assert.match(appVue, /<AProgress\s+[\s\S]*?:stroke-color="progressStrokeColor"/)
  assert.match(appVue, /<AProgress\s+[\s\S]*?:status="progressLineStatus"/)
  assert.match(appVue, /<ACard\s+v-for="item in stats"[\s\S]*?:class="\['stat-item', item\.tone\]"[\s\S]*?<AStatistic/)
  assert.match(appVue, /<ASegmented\s+[\s\S]*?class="log-tabs"/)
  assert.match(appVue, /<AButton\s+class="log-folder-button"[\s\S]*?Open Log Folder/)
  assert.doesNotMatch(appVue, /<div class="log-tabs"[\s\S]*?<button/)
  assert.doesNotMatch(appVue, /<button[^>]*class="run-button"/)
  assert.doesNotMatch(appVue, /<button[^>]*class="stop-button"/)
})

test('右侧信息卡片标题图标使用统一视觉尺寸', () => {
  const rightTitleIconRule = extractRule(
    appVue,
    '.stats-card .title-icon,\n.guide-card .title-icon,\n.env-card .title-icon',
  )
  const titleGlyphRule = extractRule(appVue, '.title-icon-glyph')
  const statsTitleIconRule = extractRule(appVue, '.stats-title-icon')
  const guideTitleIconRule = extractRule(appVue, '.guide-title-icon')
  const envTitleIconRule = extractRule(appVue, '.env-title-icon')

  assert.match(appVue, /title-icon stats-title-frame/)
  assert.match(appVue, /title-icon guide-title-frame/)
  assert.match(appVue, /title-icon env-title-frame/)
  assert.match(rightTitleIconRule, /width:\s*34px;/)
  assert.match(rightTitleIconRule, /height:\s*34px;/)
  assert.match(rightTitleIconRule, /place-items:\s*center;/)
  assert.doesNotMatch(rightTitleIconRule, /font-size:\s*30px;/)
  assert.match(titleGlyphRule, /display:\s*block;/)
  assert.match(statsTitleIconRule, /width:\s*30px;/)
  assert.match(statsTitleIconRule, /height:\s*30px;/)
  assert.match(guideTitleIconRule, /width:\s*31px;/)
  assert.match(guideTitleIconRule, /height:\s*31px;/)
  assert.match(envTitleIconRule, /width:\s*31px;/)
  assert.match(envTitleIconRule, /height:\s*31px;/)
  assert.match(envTitleIconRule, /center \/ contain no-repeat/)
})

test('运行日志使用独立阅读底色，纸质纹理和插画不进入文字层', () => {
  const logHeaderRule = extractRule(appVue, '.log-header')
  const logColumnsRule = extractRule(appVue, '.log-columns')
  const logTableRule = extractRule(appVue, '.log-table')
  const logRowRule = extractRule(appVue, '.log-row')
  const logMessageRule = extractRule(appVue, '.log-message')
  const logCornerRule = extractRule(appVue, '.log-corner-image')

  assert.match(logHeaderRule, /padding-bottom:\s*8px;/)
  assert.match(logColumnsRule, /height:\s*200px;/)
  assert.match(logColumnsRule, /margin-top:\s*8px;/)
  assert.match(logColumnsRule, /grid-template-columns:\s*minmax\(0, 1fr\) minmax\(0, 1fr\);/)
  assert.match(logTableRule, /height:\s*100%;/)
  assert.match(logTableRule, /margin-top:\s*0;/)
  assert.match(logTableRule, /background:\s*rgba\(252, 254, 255, 0\.56\);/)
  assert.match(logTableRule, /border:\s*1px solid rgba\(104, 141, 181, 0\.2\);/)
  assert.match(logTableRule, /box-shadow:\s*inset 0 1px 0 rgba\(255, 255, 255, 0\.62\);/)
  assert.match(logRowRule, /grid-template-columns:\s*68px 58px minmax\(0, 1fr\);/)
  assert.match(logRowRule, /gap:\s*6px;/)
  assert.match(logRowRule, /text-shadow:\s*none;/)
  assert.match(logRowRule, /filter:\s*none;/)
  assert.match(logMessageRule, /opacity:\s*1;/)
  assert.match(logMessageRule, /filter:\s*none;/)
  assert.match(logMessageRule, /overflow-wrap:\s*anywhere;/)
  assert.match(logMessageRule, /word-break:\s*break-word;/)
  assert.match(logCornerRule, /z-index:\s*0;/)
  assert.match(logCornerRule, /opacity:\s*0\.46;/)
  assert.match(appVue, /\.dark \.log-table\s*\{[\s\S]*?background:\s*rgba\(15, 24, 39, 0\.44\);/)
})
