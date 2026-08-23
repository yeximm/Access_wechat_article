# 快速入手

本文面向第一次阅读、使用或准备二次开发 Access WeChat Article 的使用者。

它的目标不是替代完整安装文档和功能文档，而是帮助你快速建立程序主流程的架构认知：

- 程序启动后如何运行；
- 主流程由哪些大模块组成；
- 每个大模块内部又负责哪些子能力。

如需查看完整安装步骤和页面功能说明，请先阅读：

- [安装说明](./install.md)
- [功能说明](./features.md)

## 1. 快速启动

项目使用 `uv` 管理 Python 环境。进入项目根目录后执行：

```bash
uv sync
```

首次使用 Playwright 离线缓存功能前，需要把 Chromium 浏览器安装到项目目录：

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH=".playwright-browsers"
uv run playwright install chromium
```

启动本地 API 后端：

```bash
uv run python dev_server.py
```

`main.py` 保留为旧命令兼容入口，内部同样启动 `dev_server.py` 的 FastAPI 服务。后端启动后，可以在浏览器访问本地网页端：

```text
http://127.0.0.1:8766/
```

开发调试后端接口时仍使用同一入口：

```bash
uv run python dev_server.py
```

> 注意：当前 Python 入口只启动本地 FastAPI 服务，不再打开旧桌面窗口；桌面 exe 封装将由 Tauri 承接。

## 2. 主流程总览

主流程指“主服务页面点击开始运行后，程序从当前打开的公众号/服务号主页逐篇采集文章”的流程。

从架构上看，它可以分为七个大模块：

1. **任务编排模块**：接收开始/停止指令，组织配置、进程、日志和 worker 生命周期。
2. **窗口控制模块**：识别主页窗口、识别文章候选、滚动主页、点击文章、关闭详情窗口。
3. **代理监听与捕获模块**：通过本地 MITM 监听微信内置浏览器请求，捕获当前文章的 HTML 和请求信息。
4. **文章详情解析模块**：从文章 HTML 中提取标题、公众号、发布时间、短链接和互动指标。
5. **评论采集模块**：在用户勾选评论信息时，继续获取评论、回复、评论图片和表情信息。
6. **数据清洗与存储模块**：清洗字段、生成本地目录、写入 JSON/HTML 归档和 SQLite 索引。
7. **状态日志与异常保护模块**：贯穿全流程，负责前端状态、运行日志、失败记录、停止和资源收尾。

这些模块不是简单地互相替代，而是围绕“单篇文章采集循环”协同工作：窗口模块打开一篇文章，代理模块捕获这篇文章的请求，详情和评论模块解析数据，存储模块落地结果，然后窗口模块关闭详情页并继续下一篇。

## 3. 主流程程序流程图

![9a2375dc-bee5-4f6e-8e9f-327d046cb854](./quick_start/9a2375dc-bee5-4f6e-8e9f-327d046cb854.png)

## 4. 任务编排模块

任务编排模块是主流程的入口和调度层，负责把前端的开始、停止、状态查询和日志查询请求转换成可控的后台任务。

它不直接解析文章内容，也不直接操作微信窗口，而是集中处理运行选项、主页窗口快照、worker 启动、进程生命周期、任务状态和运行日志。

主要代码在 `src/core/task_manager.py`、`src/core/process_manager.py`、`src/services/task_service.py`、`src/workers/article_capture.py` 和 `src/core/events.py`。

![任务编排流程图](./quick_start/9441cadc-2945-4083-8617-dc06120b7d43.png)

前端、Tauri 桌面壳和 FastAPI 都只是入口，最终会汇总到 `TaskManager.start_task`；真正决定能不能启动采集的是运行选项规范化、主页窗口检测和 `ProcessManager` 启动 worker 这几步。

图中的失败/取消路径单独下沉，是因为启动失败不应该混入文章采集循环。比如主页窗口不可用、已有 worker 正在运行、启动过程被取消，都应该尽早返回状态和日志，而不是让后面的窗口点击、MITM 捕获或存储模块继续执行。

## 5. 窗口控制模块

窗口控制模块负责所有和微信 PC 客户端 UI 直接交互的动作：找到真实可见的公众号/服务号/订阅号主页，读取账号和文章候选，滚动主页，点击目标文章，并在单篇采集完成后关闭微信内置浏览器详情窗口。

它的核心代码集中在 `src/modules/window/wechat_home_window_finder.py`、`src/workers/wechat_home.py`、`src/modules/window/home_article_cursor.py`、`src/workers/home_article_clicker.py`、`src/modules/window/article_window_flow.py`、`src/modules/window/detail_window_manager.py` 和 `src/modules/window/home_window_focus_guard.py`。

![主页窗口识别流程图](./quick_start/cd423299-39c1-443d-b553-75ddec977702.png)

主页窗口识别需要排除微信聊天主窗口、隐藏壳窗口、微信内置浏览器详情页和本程序窗口。

文章候选优先选择当前屏幕可见、有坐标、符合文章区域特征的节点，不把 UI 树里所有文字都当成标题。

详情窗口关闭逻辑只能关闭文章详情窗口，不关主页窗口。

## 6. 代理监听与捕获模块

代理监听与捕获模块负责拿到“用户本次点击文章后产生的真实文章请求”。它通过本地 MITM 代理监听微信内置浏览器访问 `mp.weixin.qq.com` 的请求，过滤无关资源和诊断噪声，等待当前文章的主 HTML，并生成后续解析和存储所需的 capture report；主要代码在 `src/workers/mitm_worker.py`、`src/modules/proxy/mitm_capture_waiter.py`、`src/modules/proxy/mitm_controller.py`、`src/modules/proxy/proxy_manager.py`、`src/modules/proxy/system_proxy.py` 和 `src/modules/proxy/certificate.py`。

![代理监听与捕获流程图](./quick_start/adaa2696-8cd6-49e4-a828-0e72279804e4.png)

MITM 模块不负责选择文章，也不主动刷新文章页来补救失败；它只监听、过滤、匹配并上报请求证据。

图里的 capture report 是后续模块的交接物，里面包含主 HTML、URL、请求摘要、脱敏后的关键参数和诊断信息。

日志中涉及 `key`、`pass_ticket`、`appmsg_token`、Cookie 等字段时脱敏，避免把短时间有效的敏感参数直接暴露到日志中。

## 7. 文章详情解析模块

文章详情解析模块负责把 MITM 捕获到的文章 HTML 转换成结构化字段，是从“网页内容”进入“可存储数据”的第一层。

它主要解析公众号名、文章标题、发布时间、短链接、IP 属地，以及阅读量、点赞量、转发量、推荐量、评论量等可用指标。

主要代码在 `src/modules/detail/article_detail.py`、`src/modules/detail/account_identity.py` 和 `src/modules/storage/article_archive_store.py`。

![文章详情解析流程图](./quick_start/4640ce7a-0c0a-4922-adef-44ab11721634.png)

窗口识别拿到的账号名、标题或时间只能作为兜底，最终落库前更可信的是文章详情 HTML 中解析出的字段。

短链接是存储模块判断一篇文章是否成功保存的重要字段。如果详情中无法解析到有效的 `https://mp.weixin.qq.com/s/...` 短链接，后续不应该强行写入 `saved` 记录，而是进入失败路径，避免污染 SQLite 的去重键。

## 8. 评论采集模块

评论采集模块是可选链路，只有用户在主服务页面勾选“评论信息”时才会进入。

它依赖文章详情 HTML 中的评论接口参数，用原请求上下文读取一级评论、评论回复、评论图片和表情资源，并把结果写入当前文章归档目录。

主要代码在 `src/modules/detail/comment_detail.py`、`src/workers/comment_worker.py` 和相关测试 `tests/test_comment_fetcher.py`。

![评论采集流程图](./quick_start/17f0ded7-4bbf-4940-a3ba-3ba44f098a06.png)

评论采集和文章详情保存不是同一个成功条件。

文章详情已经解析成功时，即使评论接口失败，也不应该直接否定文章详情本身的保存结果，而是记录评论失败原因并继续走文章详情存储链路。

图里评论分页、回复读取和资源保存是相对容易受接口状态影响的部分。评论接口依赖临时参数、网络状态和微信返回内容，失败原因需要写入日志或归档信息，方便后续判断是参数过期、接口无数据，还是请求本身异常。

## 9. 数据清洗与存储模块

数据清洗与存储模块负责把前面模块得到的结果落地，核心原则是“大内容进本地归档目录，轻量索引进 SQLite”。

它会清洗公众号名、标题和发布时间，分配 Windows 可用的文章目录，写入 `article_detail.json`、主 HTML 证据和可选评论文件，再构造 SQLite 记录并区分 `saved` / `failed` 状态。

主要代码在 `src/modules/storage/article_archive_store.py`、`src/modules/storage/path_builder.py`、`src/modules/storage/sqlite_store.py`、`src/modules/storage/public_article_store.py` 和 `src/modules/storage/awa_public_schema.sql`。

![数据清洗与存储流程图](./quick_start/028ef0ca-db2b-427d-9075-86e1f1174adb.png)

存储图要分开看两条线：本地归档和 SQLite 索引。本地归档保存结构化详情、原始 HTML 证据和评论相关文件；SQLite 只保存账号、标题、发布时间、短链接、采集类型、采集时间、耗时和采集状态等轻量字段。

点击/捕获失败、归档异常、解析不到短链接、SQLite 入库异常都可能汇入失败记录逻辑。

失败记录不写 `failed://` 这类占位短链，避免污染后续按 `account_id + article_link` 进行成功文章去重的业务键。

## 10. 状态日志与异常保护模块

状态日志与异常保护模块贯穿整个主流程，负责让任务“看得见、停得住、查得到”。

worker、MITM、详情解析、评论采集和存储过程都会通过事件队列上报日志或状态，`TaskManager` 负责统一归一化、等级过滤、内存裁剪、文件落盘和前端查询。

主要代码在 `src/core/task_manager.py`、`src/core/events.py`、`src/core/progress_logger.py`、`src/core/file_logger.py`、`src/core/log_levels.py`、`src/core/process_manager.py` 和 `src/workers/mitm_worker.py`。

![日志链路流程图](./quick_start/fce77fe3-de0e-4093-829f-6c08d203581e.png)

`ProgressLogger` 和 `put_event` 负责把 worker 侧的步骤、进度、成功、警告和错误放入 `event_queue`；`TaskManager._drain_worker_events` 再把事件分成普通日志、采集状态、鉴权状态和流量事件分别处理。

日志有两个主要出口：一份保存在内存 `_logs` 中供 `/api/task/logs` 快速读取，另一份通过 `SessionFileLogger.write` 写入 `data/logs/yyyy-mm-dd/*.log`。

任务结束或异常时，系统会裁剪内存日志、更新任务状态，并尽量完成详情窗口清理、worker 停止和代理状态恢复。

## 11. 主流程之外的辅助模块

除了主流程，项目还有一些围绕数据维护和使用体验的辅助模块。这些模块不一定参与“点击主页文章并采集”的主循环，但会复用 SQLite、本地归档和前端 API。

### 11.1 数据档案模块

数据档案页用于查看公众号列表、文章记录、缓存状态、删除记录、打开归档目录和批量导出 Excel。

相关代码：

- `src/modules/storage/archive_delete_service.py`
- `src/modules/storage/archive_excel_export_service.py`
- `src/modules/storage/archive_storage_info.py`

### 11.2 Playwright 离线缓存模块

离线缓存模块用于根据 SQLite 中已保存的短链接，使用 Playwright 打开文章并保存离线 `index.html` 和正文实际引用资源。

它和主流程采集不同，不依赖 MITM 的临时 key 参数。

相关代码：

- `src/modules/html_archive/`
- `src/workers/article_html_archive_worker.py`

### 11.3 系统配置和环境维护模块

系统配置页负责运行目录、日志等级、代理开关、证书安装、缓存清理等能力。

相关代码：

- `dev_server.py`
- `src/config/`
- `src/modules/proxy/`
- `src/modules/system/`

## 12. 开发定位建议

如果你要修改主流程，建议按问题类型优先定位：

- 启动无响应、停止不了：先看 `src/core/task_manager.py` 和 `src/core/process_manager.py`。
- 主页识别错误：先看 `src/modules/window/wechat_home_window_finder.py` 和 `src/workers/wechat_home.py`。
- 点击了错误文章：先看 `src/modules/window/home_article_cursor.py` 和 `src/workers/home_article_clicker.py`。
- 误关主页窗口：先看 `src/modules/window/detail_window_manager.py` 和 `src/modules/window/article_window_flow.py`。
- 捕获不到文章请求：先看 `src/modules/proxy/mitm_capture_waiter.py` 和 `src/workers/mitm_worker.py`。
- 文章详情字段缺失：先看 `src/modules/detail/article_detail.py`。
- 评论慢或评论缺失：先看 `src/modules/detail/comment_detail.py`。
- SQLite 记录异常：先看 `src/modules/storage/sqlite_store.py` 和 `src/modules/storage/article_archive_store.py`。
- 数据档案页异常：先看 `src/modules/storage/` 下的归档服务和 `vue-project/src/pages/DataFilesPage.vue`。

## 13. 推荐验证命令

窗口识别和主流程候选相关：

```bash
uv run python -m unittest tests.test_wechat_window_activation tests.test_home_article_cursor
```

归档、缓存、导出、清理相关：

```bash
uv run python -m unittest tests.test_archive_cache_service tests.test_archive_delete_service tests.test_archive_excel_export_service tests.test_runtime_cleanup tests.test_sqlite_store
```

入口文件语法检查：

```bash
uv run python -m py_compile main.py dev_server.py
```

如果涉及真实微信窗口、系统代理、MITM 或浏览器测试，测试结束后需要确认：

- `127.0.0.1:8766` 没有遗留自己启动的服务。
- `127.0.0.1:18000` / `localhost:18000` 没有遗留 MITM 子进程。
- 系统代理已恢复到测试前状态。
- 没有把测试产物写到项目根目录、`your_disk_path\tmp` 或未受控目录。
