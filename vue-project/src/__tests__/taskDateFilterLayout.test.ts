import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import test from 'node:test'
import { fileURLToPath } from 'node:url'

const appVue = readFileSync(fileURLToPath(new URL('../App.vue', import.meta.url)), 'utf8')

test('日期筛选按钮使用 Ant Design 官方向下指示图标', () => {
  assert.match(appVue, /import \{[^}]*DownOutlined[^}]*\} from '@ant-design\/icons-vue'/)
  assert.match(appVue, /const taskDateFilterOpen = ref\(false\)/)
  assert.match(appVue, /v-model:open="taskDateFilterOpen"/)
  assert.doesNotMatch(appVue, /:get-popup-container="getTaskPopupContainer"/)
  assert.doesNotMatch(appVue, /function getTaskPopupContainer/)
  assert.match(
    appVue,
    /<DownOutlined\s+:class="\['task-date-filter-chevron', \{ 'is-open': taskDateFilterOpen \}\]"\s+\/>/,
  )
  assert.doesNotMatch(
    appVue,
    /<AppIcon class="task-date-filter-chevron" icon="fa-solid fa-chevron-down" \/>/,
  )
  assert.match(
    appVue,
    /\.task-date-filter-chevron\.is-open\s*\{[\s\S]*?transform:\s*translateY\(-50%\) rotate\(180deg\);/,
  )
  assert.match(
    appVue,
    /@media \(prefers-reduced-motion:\s*reduce\)\s*\{[\s\S]*?\.task-date-filter-chevron\s*\{[\s\S]*?transition:\s*none;/,
  )
})

test('指定记录总量卡片保留标题并以下拉菜单提供四种日期筛选方式', () => {
  assert.match(appVue, /type TaskDateFilterMode = 'all' \| 'range' \| 'before' \| 'after'/)
  assert.match(appVue, /const taskDateFilterMode = ref<TaskDateFilterMode>\('all'\)/)
  assert.match(appVue, /\{ label: '不限日期', value: 'all' \}/)
  assert.match(appVue, /\{ label: '指定任务日期范围', value: 'range' \}/)
  assert.match(appVue, /\{ label: '截止日期 \(不早于\)', value: 'before' \}/)
  assert.match(appVue, /\{ label: '起始日期 \(不晚于\)', value: 'after' \}/)

  const card = appVue.match(
    /<article class="task-card task-card-volume panel">[\s\S]*?<\/article>/,
  )
  assert.ok(card, '缺少指定记录总量的独立卡片布局')
  assert.match(card[0], /<h2>指定记录总量<\/h2>/)
  assert.match(card[0], /<ADropdown/)
  assert.match(card[0], /:trigger="\['click'\]"/)
  assert.match(card[0], /class="task-date-filter-trigger"/)
  assert.match(card[0], /\{\{ taskDateFilterLabel \}\}/)
  assert.match(card[0], /<AMenu :selected-keys="\[taskDateFilterMode\]">/)
  assert.match(card[0], /<AMenuItem[\s\S]*?v-for="option in taskDateFilterOptions"/)
  assert.match(card[0], /@click="selectTaskDateFilterMode\(option\.value\)"/)
  assert.match(
    appVue,
    /function selectTaskDateFilterMode\(mode: TaskDateFilterMode\)\s*\{[\s\S]*?taskDateFilterMode\.value = mode[\s\S]*?taskDateFilterOpen\.value = false[\s\S]*?\}/,
  )
  assert.doesNotMatch(card[0], /<VxeSelect/)
  assert.doesNotMatch(appVue, /taskDatePopupConfig/)
  assert.doesNotMatch(card[0], /role="radiogroup"/)
  assert.doesNotMatch(card[0], /type="radio"/)
  assert.doesNotMatch(card[0], /默认为 1，设为 0 时遍历全部内容/)

  const dropdownMenuItemRule = appVue.match(
    /:global\(\.task-date-filter-dropdown\.ant-dropdown \.ant-dropdown-menu \.ant-dropdown-menu-item\)\s*\{(?<body>[\s\S]*?)\}/,
  )
  assert.ok(dropdownMenuItemRule?.groups?.body)
  assert.match(dropdownMenuItemRule.groups.body, /font-size:\s*calc\(14px \* var\(--app-scale\)\);/)
  assert.match(dropdownMenuItemRule.groups.body, /font-weight:\s*400;/)
  assert.match(dropdownMenuItemRule.groups.body, /line-height:\s*1\.25;/)
  assert.match(appVue, /:global\(:root\)\s*\{[\s\S]*?--app-scale:/)

  const triggerRule = appVue.match(/\.task-date-filter-trigger\s*\{(?<body>[\s\S]*?)\}/)
  const triggerTextRule = appVue.match(
    /\.task-date-filter-trigger > span:first-child\s*\{(?<body>[\s\S]*?)\}/,
  )
  assert.ok(triggerRule?.groups?.body)
  assert.match(triggerRule.groups.body, /position:\s*relative;/)
  assert.match(triggerRule.groups.body, /padding:\s*0 22px 0 4px;/)
  assert.match(triggerRule.groups.body, /overflow:\s*hidden;/)
  assert.ok(triggerTextRule?.groups?.body)
  assert.match(triggerTextRule.groups.body, /display:\s*block;/)
  assert.match(triggerTextRule.groups.body, /width:\s*100%;/)
  assert.match(triggerTextRule.groups.body, /min-width:\s*0;/)
  assert.match(triggerTextRule.groups.body, /overflow:\s*hidden;/)
  assert.match(triggerTextRule.groups.body, /text-overflow:\s*ellipsis;/)
  assert.match(triggerTextRule.groups.body, /white-space:\s*nowrap;/)

  const chevronRule = appVue.match(/\.task-date-filter-chevron\s*\{(?<body>[\s\S]*?)\}/)
  assert.ok(chevronRule?.groups?.body)
  assert.match(chevronRule.groups.body, /position:\s*absolute;/)
  assert.match(chevronRule.groups.body, /top:\s*50%;/)
  assert.match(chevronRule.groups.body, /right:\s*8px;/)
  assert.match(chevronRule.groups.body, /transform:\s*translateY\(-50%\) rotate\(0deg\);/)
})

test('四种日期筛选分别保存任务数量快照并使用规定初始值', () => {
  assert.match(appVue, /const unlimitedDateTaskCount = ref<number \| null>\(1\)/)
  assert.match(appVue, /const dateRangeTaskCount = ref<number \| null>\(0\)/)
  assert.match(appVue, /const latestDateTaskCount = ref<number \| null>\(0\)/)
  assert.match(appVue, /const earliestDateTaskCount = ref<number \| null>\(1\)/)
  assert.match(
    appVue,
    /const taskCountSnapshots: Record<TaskDateFilterMode, typeof unlimitedDateTaskCount> = \{[\s\S]*?all: unlimitedDateTaskCount,[\s\S]*?range: dateRangeTaskCount,[\s\S]*?before: latestDateTaskCount,[\s\S]*?after: earliestDateTaskCount,[\s\S]*?\}/,
  )
  assert.match(
    appVue,
    /const pageCount = computed<number \| null>\(\{[\s\S]*?get: \(\) => taskCountSnapshots\[taskDateFilterMode\.value\]\.value,[\s\S]*?taskCountSnapshots\[taskDateFilterMode\.value\]\.value = value[\s\S]*?\}\)/,
  )
})

test('日期筛选提示随当前选项显示对应说明', () => {
  assert.match(appVue, /const TASK_DATE_FILTER_HINTS: Record<TaskDateFilterMode, string> = \{/)
  assert.match(appVue, /all:\s*'从当前主页开始不限制文章发布时间'/)
  assert.match(appVue, /range:\s*'采集起始日期至截止日期内发布的文章'/)
  assert.match(appVue, /before:\s*'采集从当前主页开始到指定日期内发布的文章'/)
  assert.match(appVue, /after:\s*'采集指定日期之前发布的文章'/)
  assert.match(
    appVue,
    /const taskDateFilterHint = computed\(\(\) => TASK_DATE_FILTER_HINTS\[taskDateFilterMode\.value\]\)/,
  )
  assert.match(
    appVue,
    /<ATooltip :title="taskDateFilterHint" placement="top">[\s\S]*?aria-label="查看日期筛选说明"/,
  )
  assert.doesNotMatch(appVue, /title="选择本次任务的日期筛选方式"/)
})

test('日期控件使用 Ant Design DatePicker 系列并按筛选模式切换', () => {
  assert.match(appVue, /const taskStartDate = ref\(''\)/)
  assert.match(appVue, /const taskEndDate = ref\(''\)/)
  assert.match(appVue, /const taskDateRangeValue = computed<TaskDateRangeValue>/)
  assert.match(appVue, /class="[^"]*task-date-row[^"]*"/)
  assert.match(appVue, /<ARangePicker[\s\S]*?v-model:value="taskDateRangeValue"/)
  assert.match(appVue, /<ADatePicker[\s\S]*?v-model:value="taskEndDate"/)
  assert.match(appVue, /<ADatePicker[\s\S]*?v-model:value="taskStartDate"/)
  assert.match(appVue, /value-format="YYYY-MM-DD"/)
  assert.match(appVue, /taskDateFilterMode === 'all'[\s\S]*?:disabled="true"/)
  assert.match(appVue, /all:\s*'指定时间'/)
  assert.doesNotMatch(appVue, /<VxeDate(?:Range)?Picker/)
  assert.match(
    appVue,
    /\.task-volume-body\s*\{[\s\S]*?grid-template-rows:\s*auto repeat\(3, 32px\);/,
  )
})

test('task volume rows match the content option height and spacing', () => {
  const volumeBodyRule = appVue.match(/\.task-volume-body\s*\{(?<body>[\s\S]*?)\}/)
  const volumeHeadingRule = appVue.match(/\.task-volume-body > h2\s*\{(?<body>[\s\S]*?)\}/)
  const filterTriggerRule = appVue.match(/\.task-date-filter-trigger\s*\{(?<body>[\s\S]*?)\}/)
  const countControlRule = appVue.match(/\.task-count-control\s*\{(?<body>[\s\S]*?)\}/)
  const countInputRule = appVue.match(/\.task-count-input\s*\{(?<body>[\s\S]*?)\}/)
  const dateControlRule = appVue.match(/\.task-date-control\s*\{(?<body>[\s\S]*?)\}/)
  const datePickerRule = appVue.match(/\.task-date-picker\s*\{(?<body>[\s\S]*?)\}/)
  const downloadOptionsRule = appVue.match(/\.download-options\s*\{(?<body>[\s\S]*?)\}/)
  const downloadOptionRule = appVue.match(/\.download-option\s*\{(?<body>[\s\S]*?)\}/)

  assert.ok(volumeBodyRule?.groups?.body)
  assert.match(volumeBodyRule.groups.body, /grid-template-rows:\s*auto repeat\(3, 32px\);/)
  assert.match(volumeBodyRule.groups.body, /align-content:\s*center;/)
  assert.match(volumeBodyRule.groups.body, /gap:\s*7px;/)
  assert.match(volumeBodyRule.groups.body, /padding-top:\s*1px;/)
  assert.ok(volumeHeadingRule?.groups?.body)
  assert.match(volumeHeadingRule.groups.body, /margin-bottom:\s*3px;/)

  for (const rule of [filterTriggerRule, countInputRule, dateControlRule, datePickerRule]) {
    assert.ok(rule?.groups?.body)
    assert.match(rule.groups.body, /height:\s*32px;/)
  }

  assert.ok(countControlRule?.groups?.body)
  assert.match(countControlRule.groups.body, /width:\s*100%;/)
  assert.ok(countInputRule?.groups?.body)
  assert.match(countInputRule.groups.body, /flex:\s*1 1 auto;/)
  assert.match(countInputRule.groups.body, /max-width:\s*none;/)

  assert.ok(downloadOptionsRule?.groups?.body)
  assert.match(downloadOptionsRule.groups.body, /gap:\s*7px;/)
  assert.ok(downloadOptionRule?.groups?.body)
  assert.match(downloadOptionRule.groups.body, /min-height:\s*32px;/)
})

test('日期选择弹层跟随页面比例整体缩放', () => {
  const popupClassMatches = appVue.match(/popup-class-name="task-date-picker-panel"/g) ?? []
  assert.equal(popupClassMatches.length, 3)

  const popupScaleRule = appVue.match(
    /:global\(\.task-date-picker-panel \.ant-picker-panel-container\)\s*\{(?<body>[\s\S]*?)\}/,
  )
  assert.ok(popupScaleRule?.groups?.body)
  assert.match(popupScaleRule.groups.body, /zoom:\s*var\(--app-scale\);/)
})

test('任务数量使用 Ant Design InputNumber 并保留零值说明浮层', () => {
  const card = appVue.match(
    /<article class="task-card task-card-volume panel">[\s\S]*?<\/article>/,
  )
  assert.ok(card)
  const countRow = card[0].match(
    /<div class="task-volume-row task-count-row">[\s\S]*?(?=\n\s*<label class="task-volume-row task-date-row">)/,
  )
  assert.ok(countRow)
  assert.match(
    countRow[0],
    /<div class="task-control-label task-label-with-hint">\s*<label for="page-count">设定任务<\/label>[\s\S]*?<ATooltip title="设为0时遍历日期范围内全部内容" placement="top">/,
  )
  assert.match(card[0], /<AInputNumber/)
  assert.match(card[0], /v-model:value="pageCount"/)
  assert.match(card[0], /class="task-count-input"/)
  assert.match(card[0], /:min="0"/)
  assert.match(card[0], /:precision="0"/)
  assert.match(card[0], /:controls="true"/)
  assert.match(card[0], /:disabled="taskSettingsLocked"/)
  assert.match(card[0], /<ATooltip/)
  assert.match(card[0], /title="设为0时遍历日期范围内全部内容"/)
  assert.match(card[0], /placement="top"/)
  assert.match(card[0], /class="task-label-hint"/)
  assert.match(card[0], /aria-label="设定任务"/)
  assert.match(card[0], /aria-label="查看设定任务说明"/)
  assert.match(countRow[0], /<div class="task-count-control">\s*<AInputNumber/)
  assert.doesNotMatch(countRow[0], /<div class="task-count-control">[\s\S]*?<ATooltip/)
  assert.doesNotMatch(card[0], />任务数量<\/label>|aria-label="任务数量"/)
  assert.doesNotMatch(card[0], /showDescriptionTooltip|moveDescriptionTooltip|hideDescriptionTooltip/)
  assert.doesNotMatch(card[0], /class="number-stepper"/)
  assert.doesNotMatch(card[0], /<input[\s\S]*?type="number"/)
  assert.doesNotMatch(appVue, /startPageCountHold|handlePageStep|pageHoldDelayTimer|pageHoldIntervalTimer/)
})

test('task labels use Ant Design question icons without circular backgrounds', () => {
  const card = appVue.match(
    /<article class="task-card task-card-volume panel">[\s\S]*?<\/article>/,
  )

  assert.ok(card)
  assert.match(
    appVue,
    /import \{[^}]*QuestionCircleOutlined[^}]*\} from '@ant-design\/icons-vue'/,
  )
  assert.match(card[0], /<QuestionCircleOutlined\s*\/>/)
  assert.doesNotMatch(card[0], /fa-regular fa-circle-question/)
  assert.equal(card[0].match(/class="task-label-hint"/g)?.length, 2)
  assert.match(
    card[0],
    /<span>日期筛选<\/span>[\s\S]*?<ATooltip :title="taskDateFilterHint" placement="top">/,
  )

  const labelWithHintRule = appVue.match(/\.task-label-with-hint\s*\{(?<body>[\s\S]*?)\}/)
  const hintRule = appVue.match(/\.task-label-hint\s*\{(?<body>[\s\S]*?)\}/)
  const hintInteractionRule = appVue.match(
    /\.task-label-hint:hover,\s*\.task-label-hint:focus-visible\s*\{(?<body>[\s\S]*?)\}/,
  )
  const darkHintRule = appVue.match(/\.dark \.task-label-hint\s*\{(?<body>[\s\S]*?)\}/)

  assert.ok(labelWithHintRule?.groups?.body)
  assert.match(labelWithHintRule.groups.body, /padding-left:\s*8px;/)
  assert.match(labelWithHintRule.groups.body, /gap:\s*6px;/)
  assert.ok(hintRule?.groups?.body)
  assert.match(hintRule.groups.body, /width:\s*20px;/)
  assert.match(hintRule.groups.body, /height:\s*20px;/)
  assert.match(hintRule.groups.body, /background:\s*transparent;/)
  assert.match(hintRule.groups.body, /box-shadow:\s*none;/)
  assert.ok(hintInteractionRule?.groups?.body)
  assert.match(hintInteractionRule.groups.body, /background:\s*transparent;/)
  assert.ok(darkHintRule?.groups?.body)
  assert.match(darkHintRule.groups.body, /background:\s*transparent;/)

  const taskRowRule = appVue.match(/\.task-volume-row\s*\{(?<body>[\s\S]*?)\}/)
  assert.ok(taskRowRule?.groups?.body)
  assert.match(taskRowRule.groups.body, /grid-template-columns:\s*88px minmax\(0, 1fr\);/)
})

test('日期筛选写入主流程后端参数', () => {
  const builder = appVue.match(/function buildTaskRunOptions\(\): TaskRunOptions \{[\s\S]*?\n\}/)
  assert.ok(builder)
  assert.match(builder[0], /dateFilterMode/)
  assert.match(builder[0], /startDate/)
  assert.match(builder[0], /endDate/)
})

test('新增设置标签与获取内容选项使用相同字号', () => {
  const taskLabelRule = appVue.match(/\.task-control-label\s*\{(?<body>[\s\S]*?)\}/)
  const downloadOptionRule = appVue.match(/\.download-option\s*\{(?<body>[\s\S]*?)\}/)
  const dateInputRule = appVue.match(
    /\.task-date-control :deep\(\.ant-picker-input > input\)\s*\{(?<body>[\s\S]*?)\}/,
  )

  assert.ok(taskLabelRule?.groups?.body)
  assert.ok(downloadOptionRule?.groups?.body)
  assert.ok(dateInputRule?.groups?.body)
  assert.match(taskLabelRule.groups.body, /font-size:\s*14px;/)
  assert.match(downloadOptionRule.groups.body, /font-size:\s*14px;/)
  assert.match(dateInputRule.groups.body, /font-size:\s*14px;/)
})

test('日期控件底行与获取内容多选框保持一致的底部留白', () => {
  const dateRowRule = appVue.match(/\.task-date-row\s*\{(?<body>[\s\S]*?)\}/)

  assert.ok(dateRowRule?.groups?.body)
  assert.match(dateRowRule.groups.body, /position:\s*relative;/)
  assert.match(dateRowRule.groups.body, /grid-template-columns:\s*minmax\(0,\s*1fr\);/)
  assert.match(dateRowRule.groups.body, /width:\s*100%;/)
  assert.match(dateRowRule.groups.body, /transform:\s*none;/)
})

test('日期操作行恢复简短标签并复用文章详情选项样式', () => {
  assert.match(appVue, /all: '指定时间'/)
  assert.match(appVue, /range: '日期范围'/)
  assert.match(appVue, /before: '截止日期'/)
  assert.match(appVue, /after: '起始日期'/)
  assert.match(
    appVue,
    /class="task-control-label task-date-label download-option selected locked"/,
  )

  const dateLabelRule = appVue.match(
    /\.task-date-label\s*\{(?<body>[\s\S]*?)\}/,
  )
  assert.ok(dateLabelRule?.groups?.body)
  assert.match(dateLabelRule.groups.body, /width:\s*120px;/)
  assert.match(dateLabelRule.groups.body, /height:\s*32px;/)
  assert.match(dateLabelRule.groups.body, /min-height:\s*32px;/)
  assert.match(dateLabelRule.groups.body, /cursor:\s*default;/)
  assert.match(dateLabelRule.groups.body, /pointer-events:\s*none;/)

  const dateLabelPositionRule = appVue.match(
    /\.task-date-row > \.task-control-label\s*\{(?<body>[\s\S]*?)\}/,
  )
  assert.ok(dateLabelPositionRule?.groups?.body)
  assert.match(dateLabelPositionRule.groups.body, /right:\s*calc\(100% \+ 12px\);/)
})

test('日期范围的两个输入区等宽且使用官方箭头居中分隔', () => {
  const rangeInputRule = appVue.match(
    /\.task-date-control :deep\(\.ant-picker-range \.ant-picker-input\)\s*\{(?<body>[\s\S]*?)\}/,
  )
  const rangeSeparatorRule = appVue.match(
    /\.task-date-control :deep\(\.ant-picker-range-separator\)\s*\{(?<body>[\s\S]*?)\}/,
  )

  assert.ok(rangeInputRule?.groups?.body)
  assert.match(rangeInputRule.groups.body, /flex:\s*1 1 0;/)
  assert.match(rangeInputRule.groups.body, /min-width:\s*0;/)
  assert.ok(rangeSeparatorRule?.groups?.body)
  assert.match(rangeSeparatorRule.groups.body, /display:\s*grid;/)
  assert.match(rangeSeparatorRule.groups.body, /place-items:\s*center;/)
  assert.match(rangeSeparatorRule.groups.body, /flex:\s*0 0 28px;/)
  assert.match(rangeSeparatorRule.groups.body, /padding:\s*0;/)
  assert.doesNotMatch(appVue, /#separator/)
  assert.doesNotMatch(appVue, /task-date-separator/)
})

test('单日期选择后显示中文星期但仍保存标准日期字符串', () => {
  assert.match(
    appVue,
    /taskDateFilterMode === 'before'[\s\S]*?format="YYYY-MM-DD {2}dddd"[\s\S]*?value-format="YYYY-MM-DD"/,
  )
  assert.match(
    appVue,
    /taskDateFilterMode === 'after'[\s\S]*?format="YYYY-MM-DD {2}dddd"[\s\S]*?value-format="YYYY-MM-DD"/,
  )
  assert.match(
    appVue,
    /<ARangePicker[\s\S]*?format="YYYY-MM-DD"[\s\S]*?value-format="YYYY-MM-DD"/,
  )
})
