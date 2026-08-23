import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import test from 'node:test'

const appVue = readFileSync(new URL('../App.vue', import.meta.url), 'utf-8')

test('任务启动和运行期间锁定数量及内容设置', () => {
  assert.match(appVue, /const taskSettingsLocked = computed\([\s\S]*starting[\s\S]*running/)
  assert.match(appVue, /id="page-count"[\s\S]*:disabled="taskSettingsLocked"/)
  assert.match(appVue, /<ACheckbox[\s\S]*v-for="option in downloadOptions"[\s\S]*:disabled="option\.locked \|\| taskSettingsLocked"/)
  assert.match(appVue, /function handleDownloadSelectionChange\(\)[\s\S]*if \(taskSettingsLocked\.value\)[\s\S]*return/)
  assert.doesNotMatch(appVue, /function toggleDownloadOption\(/)
})

test('主服务启动失败时保留可见错误并恢复可编辑状态', () => {
  assert.match(appVue, /async function handleStartTask\(/)
  assert.match(appVue, /启动采集任务失败/)
  assert.match(appVue, /await startTask\(/)
  assert.match(appVue, /const status = await getTaskStatus\(\)[\s\S]*handleTaskStatusChanged\(status\)/)
})
