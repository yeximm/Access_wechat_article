import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import test from 'node:test'
import { fileURLToPath } from 'node:url'

const currentDir = dirname(fileURLToPath(import.meta.url))
const appVue = readFileSync(resolve(currentDir, '../App.vue'), 'utf8')
const pythonApi = readFileSync(resolve(currentDir, '../bridge/pythonApi.ts'), 'utf8')

test('运行状态声明单篇文章阶段记录', () => {
  assert.match(pythonApi, /export type ArticleRuntimeRecord = \{/)
  assert.match(pythonApi, /articleRecords\?: ArticleRuntimeRecord\[\]/)
})

test('主服务日志只显示后端摘要而不逐条展开文章阶段', () => {
  assert.match(appVue, /const LOG_POLL_LIMIT = 100/)
  assert.doesNotMatch(appVue, /function buildArticleRuntimeLogItems\(/)
  assert.doesNotMatch(appVue, /articleRuntimeLogs/)
  assert.match(appVue, /const systemLogRows = computed<LogDisplayRow\[\]>\(/)
  assert.match(appVue, /const articleTaskLogRows = computed<LogDisplayRow\[\]>\(/)
  assert.match(appVue, /isArticleTaskLog\(row\)/)
})

test('main service logs are split into software activity and article task columns', () => {
  assert.match(appVue, /<div class="log-columns"/)
  assert.match(appVue, /<h3>软件活动 \/ 主流程<\/h3>/)
  assert.match(appVue, /<h3>单篇任务<\/h3>/)
  assert.match(appVue, /:items="systemLogRows"/)
  assert.match(appVue, /:items="articleTaskLogRows"/)
})
