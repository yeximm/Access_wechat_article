import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import test from 'node:test'
import { fileURLToPath } from 'node:url'

const currentDir = dirname(fileURLToPath(import.meta.url))
const appVue = readFileSync(resolve(currentDir, '../App.vue'), 'utf8')

test('download options are compact and right aligned', () => {
  const match = appVue.match(/\.download-options\s*\{(?<body>[\s\S]*?)\}/)

  assert.ok(match?.groups?.body, 'missing .download-options style block')
  assert.match(match.groups.body, /justify-self:\s*end;/)
  assert.match(match.groups.body, /width:\s*min\(100%,\s*190px\);/)
})

test('指定记录总量插画向右移动 5px并保留纵向位置', () => {
  const listArtRule = appVue.match(/\.task-art-list\s*\{(?<body>[\s\S]*?)\}/)
  const heartArtRule = appVue.match(/\.task-card-content \.task-art-heart\s*\{(?<body>[\s\S]*?)\}/)

  assert.ok(listArtRule?.groups?.body, 'missing .task-art-list style block')
  assert.ok(heartArtRule?.groups?.body, 'missing .task-card-content .task-art-heart style block')
  assert.match(listArtRule.groups.body, /transform:\s*translate\(25px,\s*-15px\);/)
  assert.match(heartArtRule.groups.body, /transform:\s*translate\(-14px,\s*12px\);/)
})

test('获取指定内容把文章详情移到插画下方并在右侧加入离线归档', () => {
  assert.match(appVue, /mainTaskSelectionDefaultsApplied/)
  assert.match(appVue, /single_article_task\.comment_collection\.enabled_by_default/)
  assert.match(appVue, /single_article_task\.offline_cache\.enabled_by_default/)
  assert.doesNotMatch(appVue, /data_acquisition\.(comment_collection|offline_cache)\.enabled_by_default/)
  assert.match(appVue, /downloadSelections\.value\.offlineArchive\s*=\s*parseConfigSwitchValue/)
  assert.match(appVue, /downloadSelections\.value\.commentInfo\s*=\s*parseConfigSwitchValue/)
  assert.match(appVue, /const mandatoryDownloadOption = \{ key: 'articleDetail', label: '文章详情', locked: true \}/)
  assert.match(appVue, /type OfflineArchiveMode = 'standard' \| 'beta'/)
  assert.match(appVue, /const offlineArchiveMode = ref<OfflineArchiveMode>\('standard'\)/)
  assert.match(appVue, /const offlineArchiveModeOpen = ref\(false\)/)
  assert.match(appVue, /const offlineArchiveModeOptions = \[[\s\S]*\{ value: 'standard', label: '离线归档' \}[\s\S]*\{ value: 'beta', label: '离线归档 \(beta\)' \}/)
  assert.match(appVue, /const downloadOptions = \[\s*\{ key: 'commentInfo', label: '评论信息', locked: false \}/)
  assert.match(appVue, /const downloadSelections = ref\(\{[\s\S]*offlineArchive: false/)

  const contentCardMatch = appVue.match(/<article class="task-card task-card-content panel">[\s\S]*?<\/article>/)
  assert.ok(contentCardMatch, '获取指定内容卡片应使用独立布局类')
  const contentCard = contentCardMatch[0]

  assert.match(contentCard, /<div class="task-art-column">[\s\S]*<img class="task-art task-art-heart"/)
  assert.match(contentCard, /<ACheckbox[\s\S]*?v-model:checked="downloadSelections\[mandatoryDownloadOption\.key\]"/)
  assert.match(contentCard, /class="\[\s*'download-option',\s*'article-detail-option'/)
  assert.match(contentCard, /mandatoryDownloadOption\.label/)
  assert.match(contentCard, /<div[\s\S]*class="\[[\s\S]*'download-option-combo'[\s\S]*offlineArchiveSelected/)
  assert.match(contentCard, /<ACheckbox[\s\S]*?v-model:checked="downloadSelections\.offlineArchive"[\s\S]*?\{\{ selectedOfflineArchiveModeLabel \}\}[\s\S]*?<\/ACheckbox>/)
  assert.match(contentCard, /<ADropdown[\s\S]*?v-model:open="offlineArchiveModeOpen"[\s\S]*?:trigger="\['click'\]"[\s\S]*?overlay-class-name="offline-archive-mode-dropdown"/)
  assert.match(contentCard, /<AButton[\s\S]*?class="offline-archive-mode-trigger"/)
  assert.match(contentCard, /<DownOutlined\s+:class="\['download-select-chevron', \{ 'is-open': offlineArchiveModeOpen \}\]"\s+\/>/)
  assert.match(contentCard, /<AMenu :selected-keys="\[offlineArchiveMode\]">[\s\S]*?<AMenuItem[\s\S]*?v-for="option in offlineArchiveModeOptions"/)
  assert.doesNotMatch(contentCard, /<ASelect[\s\S]*?class="offline-archive-mode-select"/)
  assert.match(contentCard, /<ACheckbox[\s\S]*?v-for="option in downloadOptions"/)
  assert.match(contentCard, /v-model:checked="downloadSelections\[option\.key\]"/)
  assert.match(contentCard, /v-for="option in downloadOptions"[\s\S]*\{\{ option\.label \}\}[\s\S]*<\/ACheckbox>/)
  assert.doesNotMatch(contentCard, /v-for="option in downloadOptions"[\s\S]*离线归档[\s\S]*<\/ACheckbox>/)
  assert.doesNotMatch(contentCard, /v-for="option in downloadOptions"[\s\S]*文章详情/)
  assert.doesNotMatch(contentCard, /role="checkbox"|class="option-box"/)
  assert.doesNotMatch(contentCard, /<button[\s\S]*?download-option/)
})

test('离线归档保留勾选框并通过右侧下拉选择模式', () => {
  assert.match(appVue, /function handleOfflineArchiveModeChange\(value: OfflineArchiveMode\)/)
  assert.match(appVue, /offlineArchiveMode\.value = value/)
  assert.match(appVue, /offlineArchiveModeOpen\.value = false/)
  assert.match(appVue, /const selectedOfflineArchiveModeLabel = computed\(\(\) =>/)
  assert.match(appVue, /\.download-option-combo\s*\{[\s\S]*grid-template-columns:\s*minmax\(0, 1fr\) 28px;/)
  assert.match(appVue, /\.offline-archive-mode-trigger\s*\{[\s\S]*width:\s*28px;/)
  assert.match(appVue, /\.offline-archive-mode-trigger\s*\{[\s\S]*place-items:\s*center;/)
  assert.match(appVue, /\.download-select-chevron\.is-open\s*\{[\s\S]*transform:\s*rotate\(180deg\);/)
  const dropdownRule = appVue.match(
    /:global\(\.offline-archive-mode-dropdown\.ant-dropdown \.ant-dropdown-menu\)\s*\{(?<body>[\s\S]*?)\}/,
  )
  const dropdownItemRule = appVue.match(
    /:global\(\.offline-archive-mode-dropdown\.ant-dropdown \.ant-dropdown-menu \.ant-dropdown-menu-item\)\s*\{(?<body>[\s\S]*?)\}/,
  )

  assert.ok(dropdownRule?.groups?.body)
  assert.match(dropdownRule.groups.body, /min-width:\s*calc\(168px \* var\(--app-scale\)\);/)
  assert.match(dropdownRule.groups.body, /padding:\s*calc\(4px \* var\(--app-scale\)\);/)
  assert.doesNotMatch(dropdownRule.groups.body, /zoom:/)
  assert.ok(dropdownItemRule?.groups?.body)
  assert.match(dropdownItemRule.groups.body, /min-height:\s*calc\(32px \* var\(--app-scale\)\);/)
  assert.match(dropdownItemRule.groups.body, /font-size:\s*calc\(14px \* var\(--app-scale\)\);/)
})

test('获取指定内容插画下方按钮使用左侧独立布局区域', () => {
  const contentCardRule = appVue.match(/\.task-card-content\s*\{(?<body>[\s\S]*?)\}/)
  const artColumnRule = appVue.match(/\.task-art-column\s*\{(?<body>[\s\S]*?)\}/)
  const heartArtRule = appVue.match(/\.task-card-content \.task-art-heart\s*\{(?<body>[\s\S]*?)\}/)
  const articleDetailRule = appVue.match(/\.article-detail-option\s*\{(?<body>[\s\S]*?)\}/)

  assert.ok(contentCardRule?.groups?.body, 'missing .task-card-content style block')
  assert.ok(artColumnRule?.groups?.body, 'missing .task-art-column style block')
  assert.ok(heartArtRule?.groups?.body, 'missing .task-card-content .task-art-heart style block')
  assert.ok(articleDetailRule?.groups?.body, 'missing .article-detail-option style block')
  assert.match(contentCardRule.groups.body, /grid-template-columns:\s*118px minmax\(0, 1fr\);/)
  assert.match(contentCardRule.groups.body, /gap:\s*16px;/)
  assert.match(artColumnRule.groups.body, /grid-template-rows:\s*auto auto;/)
  assert.match(artColumnRule.groups.body, /align-content:\s*end;/)
  assert.match(artColumnRule.groups.body, /gap:\s*0;/)
  assert.match(heartArtRule.groups.body, /width:\s*112px;/)
  assert.match(heartArtRule.groups.body, /height:\s*112px;/)
  assert.match(heartArtRule.groups.body, /transform:\s*translate\(-14px,\s*12px\);/)
  assert.match(articleDetailRule.groups.body, /width:\s*184px;/)
  assert.match(articleDetailRule.groups.body, /min-height:\s*32px;/)
})
