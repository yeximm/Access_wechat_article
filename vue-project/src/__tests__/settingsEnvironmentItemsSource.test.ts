import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import test from 'node:test'
import { fileURLToPath } from 'node:url'

const currentDir = dirname(fileURLToPath(import.meta.url))
const appVue = readFileSync(resolve(currentDir, '../App.vue'), 'utf8')
const settingsPageVue = readFileSync(resolve(currentDir, '../pages/SettingsPage.vue'), 'utf8')
const settingsEnvironmentItemsPath = resolve(currentDir, '../utils/settingsEnvironmentItems.ts')

test('系统配置页运行环境直接使用 App.vue 的 envItems，不再单独维护设置页列表', () => {
  assert.match(appVue, /const envItems = computed\(\(\) => \[/)
  assert.match(appVue, /name:\s*'Version'/)
  assert.match(appVue, /name:\s*'Python'/)
  assert.match(appVue, /name:\s*'System'/)
  assert.match(appVue, /name:\s*'Tauri'/)
  assert.match(appVue, /name:\s*'MITMproxy'/)
  assert.match(appVue, /name:\s*'Playwright'/)
  assert.match(appVue, /<SettingsPage[\s\S]*:environment-items="envItems"/)
  assert.match(settingsPageVue, /environmentItems\?: EnvironmentItem\[\]/)
  assert.match(settingsPageVue, /props\.environmentItems \?\? \[\]/)
  assert.doesNotMatch(settingsPageVue, /buildSettingsEnvironmentItems/)
  assert.equal(existsSync(settingsEnvironmentItemsPath), false)
})
