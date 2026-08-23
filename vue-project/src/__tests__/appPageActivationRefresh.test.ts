import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import test from 'node:test'

const appVue = readFileSync(new URL('../App.vue', import.meta.url), 'utf8')
const dataFilesPageVue = readFileSync(new URL('../pages/DataFilesPage.vue', import.meta.url), 'utf8')
const historyPageVue = readFileSync(new URL('../pages/HistoryPage.vue', import.meta.url), 'utf8')
const settingsPageVue = readFileSync(new URL('../pages/SettingsPage.vue', import.meta.url), 'utf8')

function requireBlock(source: string, pattern: RegExp, label: string) {
  const match = source.match(pattern)
  assert.ok(match, `${label} should exist`)
  return match[0]
}

test('App menu selection refreshes the activated page after switching', () => {
  const selectPageBlock = requireBlock(appVue, /function selectPage\(page: PageKey \| null\)[\s\S]*?\n\}/, 'selectPage')
  const refreshBlock = requireBlock(appVue, /async function refreshActivePage\(page: PageKey\)[\s\S]*?\n\}/, 'refreshActivePage')

  assert.match(appVue, /const dataFilesPageRef = ref<RefreshablePageExpose \| null>\(null\)/)
  assert.match(appVue, /const historyPageRef = ref<RefreshablePageExpose \| null>\(null\)/)
  assert.match(appVue, /const settingsPageRef = ref<RefreshablePageExpose \| null>\(null\)/)
  assert.match(selectPageBlock, /activePage\.value = page/)
  assert.match(selectPageBlock, /void refreshActivePage\(page\)/)
  assert.match(refreshBlock, /await nextTick\(\)/)
  assert.match(appVue, /const pageActivationRefreshers: Record<PageKey, \(\) => void \| Promise<void>> = \{[\s\S]*?home: refreshHomePageOnActivated/)
  assert.match(appVue, /files: \(\) => dataFilesPageRef\.value\?\.refreshOnActivated\?\.\(\)/)
  assert.match(appVue, /history: \(\) => historyPageRef\.value\?\.refreshOnActivated\?\.\(\)/)
  assert.match(appVue, /settings: \(\) => settingsPageRef\.value\?\.refreshOnActivated\?\.\(\)/)
  assert.match(appVue, /async function refreshHomePageOnActivated\(\)[\s\S]*?refreshTaskRuntime\(\)[\s\S]*?refreshArchiveSummary\(\)[\s\S]*?refreshHardwareStatus\(\)/)
})

test('main service reapplies default task switches only when config signature changes', () => {
  const defaultsBlock = requireBlock(appVue, /function applyMainTaskSelectionDefaults\(values\?: Record<string, string>\)[\s\S]*?\n\}/, 'applyMainTaskSelectionDefaults')
  const manualChangeBlock = requireBlock(appVue, /function handleDownloadSelectionChange\(\)[\s\S]*?\n\}/, 'handleDownloadSelectionChange')

  assert.match(appVue, /const mainTaskSelectionDefaultsSignature = ref\(''\)/)
  assert.match(appVue, /function buildMainTaskSelectionDefaultsSignature\(values: Record<string, string>\)/)
  assert.match(defaultsBlock, /const defaultsSignature = buildMainTaskSelectionDefaultsSignature\(values\)/)
  assert.match(defaultsBlock, /mainTaskSelectionDefaultsApplied\.value[\s\S]*?mainTaskSelectionDefaultsSignature\.value === defaultsSignature/)
  assert.match(defaultsBlock, /mainTaskSelectionDefaultsSignature\.value = defaultsSignature/)
  assert.match(manualChangeBlock, /mainTaskSelectionDefaultsApplied\.value = true/)
  assert.doesNotMatch(defaultsBlock, /if \(mainTaskSelectionDefaultsApplied\.value \|\| taskSettingsLocked\.value \|\| !values\)/)
})

test('management pages expose refreshOnActivated for menu switching', () => {
  assert.match(dataFilesPageVue, /async function refreshOnActivated\(\)[\s\S]*?handleRefreshArchiveData\(\)/)
  assert.match(dataFilesPageVue, /defineExpose\(\{[\s\S]*?refreshOnActivated/)

  assert.match(historyPageVue, /async function refreshOnActivated\(\)[\s\S]*?loadHistorySummary\(\)[\s\S]*?loadHistorySuggestions\(\)[\s\S]*?loadHistoryRecords/)
  assert.match(historyPageVue, /defineExpose\(\{[\s\S]*?refreshOnActivated/)

  assert.match(settingsPageVue, /async function refreshOnActivated\(\)[\s\S]*?hydrateRuntimePaths\(\)[\s\S]*?syncProxySwitchState\(\)/)
  assert.match(settingsPageVue, /defineExpose\(\{[\s\S]*?refreshOnActivated/)
})
