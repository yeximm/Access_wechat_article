import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import test from 'node:test'

const appVue = readFileSync(new URL('../App.vue', import.meta.url), 'utf-8')

test('主服务开始和停止按钮绑定主流程启动停止逻辑', () => {
  assert.match(appVue, /<span class="button-label">开始运行<\/span>/)
  assert.match(appVue, /<span class="button-label">停止<\/span>/)
  assert.match(appVue, /@click="handleStartTask"/)
  assert.match(appVue, /@click="handleStopTask"/)
  assert.match(appVue, /async function handleStartTask\(/)
  assert.match(appVue, /async function handleStopTask\(/)
  assert.match(appVue, /await startTask\(/)
  assert.match(appVue, /await stopTask\(/)
})

test('主服务启动请求包含日期筛选和单篇后置选项', () => {
  assert.match(appVue, /function buildTaskRunOptions\(\): TaskRunOptions \{[\s\S]*dateFilterMode/)
  assert.match(appVue, /offlineArchiveMode/)
  assert.match(appVue, /skipCollectedRecords/)
  assert.match(appVue, /startDate/)
  assert.match(appVue, /endDate/)
})
