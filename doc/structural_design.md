# 各任务逻辑设计

这一部分主要用于理解任务运行顺序，不强调具体类名和代码细节。

整体原则：

- 单篇文章数据采集是项目核心，其他任务都围绕它提供前置、后置或辅助能力。
- 主流程只做业务编排和消息转发，不直接解析 HTML、不直接写归档、不直接操作 MITM 内部细节。
- 各级进程通过标准消息触发任务和传递结果；普通变量只属于当前进程，跨进程数据统一通过通信通道传输。
- MITM 独立进程只负责代理生命周期、监听、捕获、内存暂存和结果通知，不负责文章解析、评论采集、离线缓存和 Excel 导出。
- 窗口任务只负责识别、点击、关闭和恢复焦点，不负责网络捕获和数据保存。
- HTML 解析与保存任务负责接收 MITM 捕获结果；如果是 HTML 就直接解析，如果是 reference 就先请求 HTML，再统一解析保存。
- 当前版本数据库初始化任务是程序启动前置任务。程序只定位并执行 `data/sql/create_script/` 中已经存在的对应版本建表 SQL，不生成或改写建表 SQL。
- 文章采集、历史记录和数据档案只接收启动阶段已经确认可用的数据库路径，不在业务任务中重复初始化数据库。
- 系统代理和 MITM 端口是全局资源，同一时间只允许一个文章采集任务接管；失败、取消和异常退出都必须先恢复系统代理，再停止 MITM。
- 所有任务都要有明确输入、截止条件、输出结果和失败原因。

## 数据库初始化任务

```text
程序启动
-> ConfigService 先读取 src/config/system.yaml，再用 data/custom.yaml 覆盖同名字段
-> 校验并生成只读 AppConfig，保存到当前进程内存
task 启动
-> 接收内存中的 AppConfig，不再次读取 YAML
   -> 读取 software.data_schema_version：当前数据表版本，例如 v2.1
   -> 读取 storage.db_dir：当前数据库目录，例如 data/sql
   -> 读取 storage.db_file_name：当前数据库文件名，例如 awa-v2.1.sqlite3
-> 组合当前数据库路径：data/sql/awa-v2.1.sqlite3
-> 根据 data_schema_version 精确定位现有建表脚本
   -> v2.1 对应 data/sql/create_script/create_awa_v2_1.sql
   -> 不扫描“最新 SQL”，不由程序生成或修改建表 SQL
   -> 如果对应版本脚本不存在：停止初始化，提示配置版本与脚本不匹配
-> 检查当前版本数据库是否存在
   -> 如果存在：认为当前版本数据库已经初始化
      -> 不重复执行建表脚本
      -> 验证数据库可以打开，并确认三张必要表存在
      -> 直接返回当前数据库路径
   -> 如果不存在：进入数据库创建流程
-> 创建当前版本数据库
   -> 先创建临时数据库文件：data/sql/awa-v2.1.sqlite3.tmp
   -> 对临时数据库执行已存在的 create_awa_v2_1.sql
   -> 程序只执行脚本，不拼接或生成 CREATE TABLE SQL
   -> 执行成功后重命名为 data/sql/awa-v2.1.sqlite3
   -> 执行失败时删除本次临时文件，不生成不完整的正式数据库
   -> 返回当前数据库路径
task 结束
```

数据库初始化任务只在程序启动阶段执行。当前 `v2.1` 使用已经存在的 `data/sql/create_script/create_awa_v2_1.sql`，该脚本定义 `awa_public_accounts`、`awa_public_articles`、`awa_fetch_history` 三张表及相关索引。程序不得在运行时生成另一份建表 SQL。配置文件也只在应用启动或用户明确执行配置重载时读取；文章循环、窗口模块、代理模块和存储模块只接收内存配置或明确参数。

后续所有业务任务只通过 `TaskContext.db_path` 接收已经可用的当前版本数据库路径。业务任务不关心数据库文件是本次启动创建的，还是之前已经存在的，也不再调用数据库初始化服务。

## 文章采集整体流程

```text
task 启动
-> TaskManager 创建任务并获取文章采集独占锁
   -> 如果已有采集任务接管系统代理：拒绝重复启动并返回占用任务信息
   -> 获取成功后生成本次任务唯一 proxy_lease_id，供所有 MITM 尝试校验同一份代理接管权
-> 接收 TaskContext：task_id、proxy_lease_id、db_path、storage_root、temp_dir、started_at、取消令牌
-> 读取采集配置：目标成功数量、最大尝试次数、是否采集评论、请求间隔、单篇失败额外重试次数
-> 执行运行前预检
   -> 检查 TaskContext.db_path 指向的 SQLite 是否可连接
   -> 检查 storages/ 是否可写
   -> 检查 MITM 端口是否被无关进程占用
   -> 检查 CA 证书状态，必要时提示安装
-> 查找微信主页窗口：确认是公众号/服务号/订阅号主页，不是聊天窗口或文章详情页
-> 读取主页信息：公众号名称、简介、原创、朋友关注等信息
-> 进入文章采集循环
   -> 识别当前可见文章卡片
   -> 选择下一篇需要采集的文章
   -> 在全局最大尝试次数允许范围内，为目标文章执行首次尝试和必要的重试
      -> 每进入一次真实采集尝试就生成新的 attempt_id，并把总尝试次数加 1
      -> 每次尝试只调用一次单篇文章采集任务；首次尝试和重试都计入总尝试次数
      -> 每次尝试恰好追加一条 awa_fetch_history；成功记录随文章事务写入，MITM 未就绪、窗口检测失败、无捕获结果、解析或保存失败由主流程补写失败记录
      -> 成功则结束当前目标的重试；失败时仅在单篇额外重试次数和全局最大尝试次数都未耗尽时重试
   -> 接收目标文章最终结果：成功、失败、跳过
   -> 更新成功数量、跳过数量和失败记录
   -> 每次循环检查取消令牌
-> 判断是否继续
   -> 已达到目标成功数量：结束循环
   -> 已达到最大尝试次数：结束循环，避免持续失败导致无限循环
   -> 当前可见文章已经处理完：滚动主页，继续识别下一批文章卡片
   -> 主页无法继续滚动或连续无候选文章：结束循环
-> 执行收尾清理：关闭残留文章详情页、清理 data/tmp/<task_id>/、确认系统代理未被残留接管
-> 释放文章采集独占锁
-> 输出整体采集结果：成功数量、失败数量、跳过数量、失败原因列表
task 结束
```

“采集数量”统一表示成功保存的文章数量。`最大尝试次数` 是整个任务的真实采集尝试硬上限，首次尝试和重试都占用一次；过滤卡片或主动跳过不计入。`单篇失败重试次数` 表示同一目标在首次失败后最多还能额外尝试多少次，但剩余全局尝试次数不足时必须立即停止。尝试次数在进入单篇任务、准备创建本次 MITM 子进程时递增，因此 MITM 启动失败、未收到 READY、窗口检测失败和未捕获到内容也都算一次。每次重试前必须完成上一次详情窗口、系统代理和 MITM 子进程的清理，并重新刷新文章卡片坐标。

`awa_fetch_history` 按每次真实采集尝试恰好追加一条记录，而不是只记录最终成功。文章保存成功时，由 HTML 解析与保存服务在文章 UPSERT 的同一事务中写成功历史；其他失败由 `ArticleCaptureService` 在收到单次尝试结果后写失败历史。尚未解析出文章或公众号 ID 的早期失败允许 `article_id`、`account_id` 为空，并使用目标公众号名称、目标标题、失败阶段和错误原因保留必要上下文；如果 SQLite 本身不可写，则至少写入任务事件和日志，不能为了记录历史而掩盖原始失败。

文章采集整体流程也必须放在 `try/finally` 中。即使预检失败、用户取消或主流程异常，仍要执行当前任务清理并释放文章采集独占锁。

## 主页文章选择任务

```text
task 启动
-> 接收主页窗口和采集配置
-> 读取当前可见区域的文章卡片
   -> 获取文章标题
   -> 获取文章卡片区域
   -> 获取可点击坐标
   -> 记录卡片在当前页面中的顺序
-> 过滤不可采集卡片
   -> 标题为空则跳过
   -> 坐标无效则跳过
   -> 本轮已经点击过的 ArticleTarget.fingerprint 则跳过，避免滚动后重复点击
   -> 本轮近期失败且未到重试条件的 fingerprint 则跳过
   -> 视频号、贴图、非文章区域等干扰内容则跳过
-> 选择下一篇目标文章
   -> 优先选择当前可见区域中顺序最靠前的有效文章
   -> 如果没有有效文章，返回需要滚动主页
-> 输出目标文章：文章标题、点击坐标、所属公众号、主页窗口标识
task 结束
```

主页阶段只能基于 `ArticleTarget.fingerprint` 做本轮去重，不能仅凭标题判断数据库中是否已经采集。永久去重必须等 HTML 解析出文章短链后，按 SQLite 唯一键 `(account_id, article_link)` 判断。

## 单篇文章采集任务

```text
task 启动（只代表一次采集尝试，不在内部重试）
-> 接收采集配置、TaskContext、ArticleTarget、attempt_id 和取消令牌
-> 刷新主页可见文章卡片，重新确认目标标题和点击坐标仍有效
-> 为本次 attempt_id 新建 MITM 捕获子进程和独立通信通道，并发送 START_CAPTURE
   -> 不复用上一次尝试的子进程、通信通道或 CaptureBuffer
   -> 子进程独立执行 MITM 监听、系统代理接管、捕获变量维护和代理恢复
-> 等待 MITM 子进程发送 READY
   -> READY 表示 MITM 已监听、系统代理已指向 MITM、后续请求可以被捕获
   -> 未收到 READY：取消本次点击，执行子进程清理并返回失败
-> 收到 READY 后点击文章
-> 检测本次点击后新出现的微信文章详情窗口
   -> 必须确认是微信内置浏览器文章详情，不是主页或旧的残留窗口
   -> 必须检测到文章标签标题，并与目标标题一致或经过标准化后可确认一致
   -> 检测到目标标题说明文章已经打开；此时 HTML/reference 应已在本次 CaptureBuffer 中，不为 MITM 额外等待
   -> 标题出现后等待 0.5s，只用于确认窗口状态稳定
   -> 在窗口检测超时前仍未出现目标标题：停止本次捕获并返回窗口检测失败
-> 只关闭本次文章详情窗口，不关闭主页窗口
-> 向 MITM 子进程发送 STOP_CAPTURE
-> MITM 子进程冻结当前捕获变量、完成代理关闭后返回 RESULT
   -> status=success、capture_type=html：直接使用捕获的 response HTML
   -> status=success、capture_type=reference：后续使用 reference 请求 HTML
   -> status=failed、capture_type=none：标题已出现但变量中没有 HTML/reference，直接判定捕获失败
-> 主流程接收 MitmCaptureResult，并转交 HTML 解析与保存任务
-> 如果文章保存成功且开启评论采集：进入评论采集任务
   -> 评论失败只记录子任务失败和 warning，不改变文章主采集成功状态
-> 在 finally 中执行单次尝试收尾
   -> 关闭本次详情窗口
   -> 子进程仍存活时发送 CANCEL，并等待其按“系统代理优先恢复、MITM 随后停止”的顺序退出
   -> 超时仍未退出时，由主流程根据 PROXY_SNAPSHOT 执行任务级代理残留恢复，再 terminate/kill 并 join 子进程
   -> 本次 task 结束前必须确认该 MITM 子进程已退出，后续尝试不得复用旧进程
-> 输出本次尝试结果：成功、失败、跳过、文章目录、SQLite 记录 ID、warning、失败阶段和失败原因
task 结束
```

是否对同一目标重试由上层 `ArticleCaptureService` 决定，单篇任务只执行一次。这样总尝试次数只有一个计数入口，避免“单篇内部重试”绕过全局最大尝试次数。

## MITM捕获进程任务

```text
task 启动（每次采集尝试新建一个进程）
-> 接收捕获配置、task_id、attempt_id、proxy_lease_id：URL 匹配规则、reference 匹配规则、监听超时、代理配置
-> 接收本次尝试专用的双向通信通道
   -> 主流程发送：START_CAPTURE、STOP_CAPTURE、CANCEL
   -> 子进程发送：READY、RESULT、FAILED
   -> 所有消息携带 task_id、attempt_id；忽略与本进程不匹配的过期消息
-> 子进程启动后接收本次 START_CAPTURE，立即进入代理和捕获会话，不作为常驻进程等待下一篇文章
-> 执行代理开启任务
   -> 代理开启任务在修改系统代理前发送 PROXY_SNAPSHOT；主流程保存本次快照，只用于异常退出恢复
   -> 先启动 MITM 监听
   -> 再开启系统代理
   -> 开启成功后记录 listen_started_at，并通知主流程：READY
-> 监听 READY 之后即将产生的网络请求
   -> 只处理微信文章相关请求
   -> 忽略 listen_started_at 之前已经存在的旧请求
   -> 忽略与目标文章明显无关的请求
-> 捕获 reference
   -> 如果 reference 先出现：写入本进程 CaptureBuffer
   -> 保留 reference，并在收到 STOP_CAPTURE 前继续监听 HTML，不主动结束捕获
-> 捕获 response HTML
   -> 如果 HTML 有效：写入本进程 CaptureBuffer
   -> 即使 reference 已先到，也将最终捕获结果升级为 HTML
   -> 不主动关闭文章窗口，不直接解析或写本地文件
-> 接收主流程 STOP_CAPTURE
   -> STOP_CAPTURE 表示主流程已经检测到并关闭本次目标文章窗口
   -> 立即冻结 CaptureBuffer，不再接收本篇文章的新结果
-> 生成标准 MitmCaptureResult
   -> HTML 存在：status=success、capture_type=html
   -> HTML 不存在但 reference 存在：status=success、capture_type=reference
   -> 两者都不存在：status=failed、capture_type=none，并记录失败阶段和原因
-> 执行代理关闭任务
   -> 先恢复/关闭系统代理
   -> 再关闭 MITM 监听
-> 通过通信通道向主流程发送最终 RESULT
   -> HTML 结果包含 HTML、request 摘要、HTML 来源 mitm_response
   -> reference 结果包含 reference、必要 headers、URL、临时参数摘要
   -> RESULT 携带 task_id、attempt_id；普通进程变量不能直接交给其他进程，主流程接收序列化结果后再转发给解析或保存子任务
-> 发送 RESULT 后退出当前子进程；主流程负责 join，本进程不等待也不处理下一次采集
-> 如果收到 CANCEL 或发生异常
   -> 仍然先恢复系统代理，再停止 MITM
   -> 发送 FAILED，并确保子进程退出
task 结束
```

## 代理开启任务

```text
task 启动
-> 接收 MITM 捕获进程的启动请求
-> 校验主流程传入的 proxy_lease_id：确认它仍持有 TaskManager 的全局代理接管权，不在子进程内重复获取第二把锁
-> 读取当前系统代理状态：保存用户原本代理快照
-> 在修改系统代理前把代理快照发送给主流程，保证子进程异常退出时仍可恢复
-> 检查 MITM 端口：确认没有被无关进程占用
-> 在当前 MITM 子进程中启动监听实例
-> 等待 MITM 就绪：确认 127.0.0.1:18000 可连接
-> 开启系统代理：将系统代理指向 MITM 地址
-> 校验系统代理状态：确认当前代理确实指向 MITM
-> 发布代理开启结果
   -> 成功：返回 READY
   -> 失败：回滚系统代理和 MITM，返回失败原因
task 结束
```

系统代理是全局资源。上述顺序可以避免系统代理指向尚未启动或已经停止的 MITM，从而降低其他应用断网风险；但使用证书锁定、明确拒绝 MITM 证书的第三方应用，在采集期间仍可能不兼容，不能承诺所有应用完全不受影响。

常规文章采集不在每次代理开启或关闭时发起 HTTPS 连通性请求。HTTPS 检测只用于程序启动预检、系统配置页手动检测或代理异常排查，避免每篇文章重复联网校验和延长代理占用时间。

## 代理关闭任务

```text
task 启动
-> 接收 MITM 捕获进程的关闭请求
-> 读取当前代理状态：确认系统代理是否仍指向本次 MITM
-> 恢复/关闭系统代理
   -> 如果启动前有代理快照：恢复到启动前状态
   -> 如果启动前没有代理：关闭系统代理
   -> 如果当前代理已经不是本次 MITM：只记录状态，不覆盖用户新设置
-> 停止 MITM 监听：关闭当前子进程内的 MITM 监听实例
-> 检查端口释放：确认 18000 端口不再被当前 MITM 占用
-> 检查系统代理：确认系统代理没有继续指向 127.0.0.1:18000
-> 释放系统代理接管权
-> 发布代理关闭结果：成功、失败原因、残留状态
task 结束
```

## html解析与保存

```text
task 启动
-> 接收 MITM 标准捕获结果：capture_type=html 或 capture_type=reference
-> 接收目标文章信息：目标标题、所属公众号、点击坐标、采集开始时间
-> 接收请求证据：HTML 来源、request 摘要、reference 摘要
-> 准备文章 HTML
   -> 如果 MITM 结果是 HTML：直接作为原始文章页面，HTML 来源标记为 mitm_response
   -> 如果 MITM 结果是 reference：先构造文章 HTML 请求
      -> 使用 reference 中的文章 URL 和关键参数
      -> 复用必要 request headers
      -> 设置请求超时时间
      -> 显式绕过本地 MITM 代理，避免受到已经关闭的系统代理影响
   -> reference 请求成功：得到文章 HTML，HTML 来源标记为 reference_request
   -> reference 请求失败：返回失败阶段、失败原因、HTTP 状态或异常信息，结束本篇文章保存
-> 敏感参数处理
   -> key、pass_ticket、appmsg_token 等只作为本地临时证据保存
   -> 日志、前端提示、错误报告必须脱敏
-> 解析基础信息
   -> 解析公众号名称
   -> 解析文章标题
   -> 解析发布时间
   -> 解析文章短链
   -> 解析 IP 属地
-> 解析统计指标
   -> 解析听众量
   -> 解析阅读量
   -> 解析点赞量
   -> 解析转发量
   -> 解析推荐量
   -> 解析评论量
   -> 无法获取的指标保留为空值，不默认写成 0
-> 执行覆盖前完整校验
   -> HTML 必须是有效微信文章页面
   -> 公众号名称、文章标题、发布时间和文章短链必须有效
   -> 解析标题与目标标题必须一致，或经过统一标准化后可确认一致
   -> 标题明显不一致、正文无效或短链无效：记录风险并停止，不覆盖任何已有成功归档
-> 查询 SQLite 永久去重记录
   -> 先按唯一 account_name 查询或创建公众号，再按 (account_id, article_link) 查询已有文章
   -> 已存在且解析数据有效：更新同一条文章记录并覆盖本次成功产生的资源
   -> 不存在：创建新的文章索引记录
-> 生成确定性归档目录
   -> 已有文章且 archive_dir 非空：直接复用原 archive_dir，不因标题、公众号名或发布时间展示变化而迁移目录
   -> 新文章或旧记录缺少 archive_dir：按以下规则首次生成目录
      -> 使用公众号名称作为一级目录
      -> 公众号唯一身份只使用规范化后的 account_name，不读取、不比较 account_biz
      -> 使用“规范化 account_name + 规范化文章短链”计算 SHA-256，截取前 12 位作为稳定 article_key
      -> 文章目录固定为“发布时间 + 清洗后的文章标题 + __ + article_key”
      -> 相同文章每次得到同一目录，不再追加 _1、_2、_3 等递增后缀
      -> 将相对 storages/ 的路径保存到 awa_public_articles.archive_dir
-> 在 data/tmp/<task_id>/article_stage/<article_key>/ 准备本次文件
   -> 原始 HTML：origin/main.html
   -> 请求摘要：origin/request.json
   -> 文章详情：article_detail.json
-> 生成 article_detail.json
   -> 写入公众号名称、文章标题、发布时间、短链、IP 属地
   -> 写入听众、阅读、点赞、转发、推荐、评论等指标
   -> 写入 HTML 来源：mitm_response 或 reference_request
   -> 写入采集耗时
   -> 写入采集时间
   -> 写入记录状态
-> 覆盖本次成功产生的资源
   -> 主采集成功只覆盖 origin/main.html、origin/request.json、article_detail.json
   -> 本次没有执行评论或离线缓存时，保留已有 comments/final.json、comments/assets/、index.html 和 assets/
   -> 替换前把将被覆盖的旧文件移动到同文件系统的本次备份目录；新文章记录本次新建目录
   -> 暂存文件校验或替换失败：立即恢复旧文件，不留下半成品
-> 写入 SQLite
   -> 保存或更新公众号记录
   -> 文章记录按 (account_id, article_link) 执行 UPSERT，重复采集不新增文章 ID
   -> 更新 article_title、published_article_time、archive_dir、last_collected_time、updated_time
   -> 根据本地实际存在的资源刷新 resource_types_json；本地文件是资源是否存在的真实依据，SQLite 字段只做快速索引
   -> 文章索引更新和本次成功 awa_fetch_history 在同一个 SQLite 事务中提交
   -> SQLite 提交成功：删除本次旧文件备份，完成覆盖
   -> SQLite 提交失败：回滚事务并执行文件补偿回滚；已有文章恢复旧文件，新文章删除本次新建资源和空目录
   -> 补偿完成后返回数据库保存失败，由 ArticleCaptureService 统一尝试追加本次失败历史；SQLite 仍不可写时只记录任务事件和日志
   -> 不保留“文件已更新但数据库仍是旧状态”的结果
-> 输出保存结果
   -> 成功：返回文章目录、article_detail.json 路径、SQLite 记录 ID
   -> 失败：返回失败阶段、失败原因、上一次有效归档路径（如果存在）；本次失败产生的新文件已回滚
task 结束
```

## 评论采集任务

```text
task 启动
-> 接收文章保存结果：文章目录、article_detail.json、origin/request.json、文章 HTML
-> 判断是否具备评论采集条件
   -> 先从 origin/main.html 解析 HTML 评论数
   -> 没有解析到评论数或评论数为 0：判定无评论，不构建评论请求参数，不发起评论接口请求，写入空 comments/final.json 覆盖旧评论结果
   -> HTML 评论数大于 0：继续构建评论请求参数
   -> 评论参数不足：记录评论采集跳过原因，不影响文章采集成功状态
   -> 没有开启评论采集：直接跳过
-> 提取评论请求参数
   -> 从文章 URL、HTML、request 摘要中提取评论接口需要的参数
   -> 只使用本篇文章保存的本地证据，不重新启动 MITM
-> 请求第一页评论：获取一级评论列表和分页信息
-> 判断是否还有下一页
   -> 有下一页：继续请求
   -> 没有下一页：停止分页
-> 遍历评论回复：对有回复的评论继续请求回复接口
-> 合并评论数据：精选评论、普通评论、回复列表
-> 评论去重：按评论 ID 或内容特征去重
-> 下载评论资源：头像、图片、表情包动图等需要本地保存的资源
-> 在 data/tmp/<task_id>/comment_stage/ 准备 comments/final.json 和 comments/assets/ 相关资源
-> 校验评论结果有效后，覆盖本篇文章已有评论文件
   -> 替换前备份本次会覆盖的旧评论文件
   -> 评论采集失败时保留上一次成功的评论文件
   -> 不修改正文证据和离线缓存资源
-> 在同一 SQLite 事务中刷新 resource_types_json，并追加一条 comment_fetch 成功历史
   -> SQLite 提交成功：删除旧评论文件备份
   -> SQLite 提交失败：恢复旧评论文件，再尽力用独立短事务追加一条 comment_fetch 失败历史
   -> 请求或文件阶段失败：不覆盖旧文件，只追加一条 comment_fetch 失败历史
-> 输出评论采集结果：成功、跳过、失败原因、评论数量、资源数量
task 结束
```

## 离线缓存任务

```text
task 启动
-> 读取缓存目标：从 SQLite 查询需要缓存的文章数据
-> 为每篇文章创建缓存任务：文章标题、短链、公众号、归档目录
-> 判断是否需要缓存
   -> 已存在 index.html 且不强制刷新：跳过
   -> 缺少短链或归档目录：记录失败原因
-> 使用 Playwright 打开文章
   -> 只使用 SQLite 保存的文章短链
   -> 不使用 MITM 捕获的 key、pass_ticket 等敏感参数
   -> 不开启系统代理，不依赖 MITM
-> 等待正文加载：确认页面主体内容出现
-> 自适应滚动页面：短文章快速结束，长文章尽量加载完整
-> 收集页面资源：图片、样式、正文引用资源
-> 过滤无关资源：不保存无意义 JS、字体、表情等资源
-> 保存资源到 assets/
-> 重写 HTML/CSS 链接：将远程资源地址替换成本地资源路径
-> 在 data/tmp/<task_id>/offline_stage/ 准备 index.html 和 assets/
-> 校验离线页面有效后，覆盖文章归档目录中的 index.html 和 assets/
   -> 替换前备份本次会覆盖的旧离线页面和资源
   -> 缓存失败时保留上一次成功的离线页面
   -> 不修改正文证据和评论资源
-> 在同一 SQLite 事务中刷新 resource_types_json，并追加一条 offline_cache 成功历史
   -> SQLite 提交成功：删除旧离线资源备份
   -> SQLite 提交失败：恢复旧离线资源，再尽力用独立短事务追加一条 offline_cache 失败历史
   -> 页面获取或文件阶段失败：不覆盖旧文件，只追加一条 offline_cache 失败历史
-> 记录缓存结果：成功、跳过、失败原因、资源数量、耗时
task 结束
```

## Excel 导出任务

```text
task 启动
-> 接收导出目标：一个或多个公众号 ID、用户选择的导出目录
-> 校验导出目录：确认目录存在且可写，不写项目根目录
-> 查询 SQLite：读取公众号下所有文章记录
-> 根据 awa_public_articles.archive_dir 定位每篇文章的 article_detail.json
-> 补齐统计字段：听众、阅读、点赞、转发、推荐、评论等数据
-> 处理缺失详情
   -> 缺少 article_detail.json：保留 SQLite 基础字段并写明状态
   -> 统计字段为空：保持空值，不默认写 0
-> 生成 Excel：每个公众号单独生成一个文件
-> 写入临时目录：先写入 data/tmp/archive_excel_export_*
-> 复制到用户选择目录：复制成功后返回最终路径
-> 清理临时文件：只清理本次导出产生的临时文件
-> 返回导出结果：文件路径、文章数量、失败原因
task 结束
```

## 归档删除任务

```text
task 启动
-> 接收删除目标：文章 ID、公众号 ID 或全部归档
-> 查询 SQLite：找到对应文章记录和公众号记录
-> 读取 awa_public_articles.archive_dir：定位 storages 下的文章目录
-> 执行删除前校验
   -> 只允许删除 storages/ 下的归档目录
   -> 路径异常或越界时直接拒绝删除
-> 将待删除归档移动到 data/tmp/<task_id>/delete_trash/，暂不永久删除
-> 在数据库事务中删除文章记录
   -> 删除公众号时，必须先确认该公众号已没有文章记录，符合外键 RESTRICT 约束
-> 数据库事务成功：永久清理本次 delete_trash
-> 数据库事务失败：将归档目录恢复到原位置，并记录失败原因
-> 汇总删除结果：成功数量、失败数量、跳过数量
-> 返回删除报告：每条失败原因都要保留
task 结束
```

## 运行时清理任务

```text
task 启动
-> 定位任务性质：这是异常兜底和整体收尾任务，不替代单篇 MITM 捕获进程的正常代理关闭任务
-> 接收 task_id：只清理 data/tmp/<task_id>/，不清理其他任务目录
-> 检查当前任务：确认没有仍在运行的子进程
-> 检查 MITM 进程：只关闭当前 task_id 启动或明确归属当前任务的 MITM 进程
-> 检查 18000 端口：确认 MITM 端口是否释放
-> 检查系统代理
   -> 如果仍指向 127.0.0.1:18000：恢复或关闭
   -> 如果已经被用户改成其他代理：只记录状态，不覆盖用户设置
-> 检查残留文章详情窗口：只关闭确认属于微信内置浏览器文章详情的窗口
-> 释放当前任务持有的文章采集独占锁
-> 记录清理结果：清理了哪些文件、关闭了哪些进程、恢复了哪些代理状态
task 结束
```

`8766` 是正在运行的 FastAPI 服务端口，不属于普通业务任务清理范围。停止后端端口只允许由桌面程序整体退出流程或测试收尾流程处理。

# src 整体结构设计

目录保证每层职责清楚、数据流标准、单篇文章采集链路能跑通。

保留 7 个一级目录：

```text
src/
  app/        # 程序入口层：API、桌面桥接、静态页面挂载
  webview/    # 前端构建后的静态页面产物，必须保留
  services/   # 业务流程编排层：把 modules 串成完整任务
  modules/    # 单一能力工具层：代理、窗口、请求、归档、系统、进程
  domain/     # 标准数据结构、状态、错误码、结果对象
  storage/    # SQLite 初始化与仓储读写
  config/     # system.yaml/custom.yaml 分层读取、保存、恢复和校验
```

暂时不单独创建 `workers/`。第一阶段的后台任务可以先由 `services/task/task_manager.py` 承接；等后续任务并发、取消、队列、进度推送变复杂后，再把后台 runner 独立出来。

核心依赖方向：

```text
app -> services
webview -> app API / Tauri desktop shell
services -> modules / storage / config / domain
modules -> domain / config
storage -> domain
domain -> 不依赖其他业务目录
```

不要反向调用：

- `modules` 不调用 `services`。
- `storage` 不调用 `services` 和 `modules`。
- `app` 不直接操作 MITM、微信窗口、SQLite 细节。
- `modules` 和 `storage` 不直接读取 `custom.yaml`；由 `services` 传入 `AppConfig`、`TaskContext` 或明确参数。
- `webview` 不放 Python 业务代码，只放前端构建产物。

## 标准数据流

所有任务尽量使用同一套数据流，方便前端展示状态，也方便后续排查问题。

```text
ApiRequest
-> TaskManager 创建任务、独占锁和取消令牌
-> Service 接收 TaskCommand + TaskContext
   -> 调用 Module，接收 ToolResult
   -> 按业务阶段调用 Repository 持久化
   -> 运行过程中持续发布 TaskEvent 给 TaskManager
-> Service 返回 TaskResult
-> TaskManager 保存最终状态和事件
-> ApiResponse 返回任务摘要或查询结果
```

第一阶段只需要定义少量通用对象：

- `TaskCommand`：任务输入，例如目标成功数量、最大尝试次数、是否采集评论、请求间隔、失败重试次数。
- `TaskContext`：任务运行上下文，例如 `task_id`、`proxy_lease_id`、`db_path`、`storage_root`、`temp_dir`、`started_at`、取消令牌。
- `ToolResult`：`modules` 工具返回结果，包含 `status`、`data`、`error_code`、`message`、`duration_seconds`。
- `TaskEvent`：任务过程事件，包含 `task_id`、`stage`、`status`、`message`、`payload`、`created_at`。
- `TaskResult`：`services` 对外结果，包含 `status`、`data`、`error_code`、`message`、`warnings`；完整事件由 `TaskManager` 管理，不重复塞入最终结果。
- `MitmCaptureResult`：MITM 子进程返回的可序列化结果，包含 `task_id`、`attempt_id`、`status`、`capture_type`、HTML/reference 数据、请求摘要、错误阶段和错误信息。
- `ProcessMessage`：进程通信消息，必须包含 `task_id`、`attempt_id`、消息类型和可序列化 payload；START_CAPTURE 的 payload 还包含 `proxy_lease_id`，防止旧进程消息串入新尝试或越权接管代理。
- `ResourceManifest`：根据文章目录中的真实文件生成的资源清单；SQLite 的 `resource_types_json` 只是快速索引，不是资源存在性的唯一依据。

第一阶段建议统一这些标志：

- `task_status`：`pending`、`running`、`success`、`failed`、`skipped`、`cancelled`。
- `capture_type`：`html`、`reference`、`none`；失败属于 `status`，不再把 `failed` 当成捕获类型。
- `resource_type`：`article_detail`、`origin_html`、`origin_request`、`comment_detail`、`comment_assets`、`offline_html`、`offline_assets`。
- `process_message_type`：`start_capture`、`proxy_snapshot`、`ready`、`stop_capture`、`result`、`cancel`、`failed`。
- `stage`：`db_init`、`preflight`、`home_scan`、`single_capture`、`mitm_capture`、`html_save`、`comment_collect`、`offline_cache`、`export`、`cleanup`。

# app

`app` 是入口层，只负责接收请求、做参数转换、调用 `services`。这一层不写采集细节，不直接操作 MITM、微信窗口、SQLite、Playwright。

## api

- `server.py`
  - 创建 FastAPI 应用。
  - 注册 API 路由。
  - 挂载 `src/webview/` 静态页面。

- `capture_routes.py`
  - 提供文章采集启动、停止、状态查询接口。
  - 启动、停止和状态查询统一调用 `services/task/task_manager.py`。
  - `TaskManager` 在后台调用 `services/capture/article_capture_service.py`，路由不绕过任务管理器直接启动长任务。

- `archive_routes.py`
  - 提供文章列表查询、离线缓存、Excel 导出、归档删除接口。
  - 调用 `services/archive/`。

- `system_routes.py`
  - 提供系统状态、代理检测、CA 证书检测、端口检查接口。
  - 调用 `services/runtime/system_service.py`，不直接调用系统模块。

- `config_routes.py`
  - 提供配置读取、保存、重置接口。
  - 调用 `services/config/config_service.py`，不直接读写 YAML。

## desktop

- Tauri 桌面壳
  - 负责加载 Vue 构建产物或开发服务器页面。
  - 只承接桌面窗口、安装包和必要的系统级桌面能力。
  - 业务能力优先走 FastAPI `/api/...`，避免 API 重复。

# webview

`webview` 必须保留，它放的是前端构建后的静态页面文件，用于 FastAPI 静态加载，也可作为 Tauri 打包时的前端产物来源。

这一层不是前端源码目录，也不是后端业务目录。前端源码仍放在项目根目录的 `vue-project/src/`。修改页面时先改 `vue-project/src/`，再构建输出到 `src/webview/`。

- `index.html`
  - 前端构建后的入口页面。
  - 由 FastAPI 返回给浏览器预览，也可由 Tauri 打包后加载。

- `assets/`
  - 前端构建后的 JS、CSS、图片等静态资源。
  - 由 Vite 构建生成，不在这里手写业务逻辑。

- `favicon` / `icons`
  - 程序窗口图标或页面图标。
  - 如果由前端构建工具生成，就跟随构建产物一起更新。

前端源码建议结构：

```text
vue-project/src/
  pages/
  components/
  api/
  stores/
  router/
  types/
  utils/
  bridge/
```

# services

`services` 是业务流程编排层。它负责把多个 `modules` 工具按任务生命周期串起来，但不把具体工具实现写在 service 内部。

## capture

- `article_capture_service.py`
  - 编排文章采集整体流程。
  - 接收已经包含 `db_path` 的 `TaskContext`，不负责数据库初始化。
  - 负责运行前预检、主页扫描、文章循环、成功/尝试计数、单篇额外重试、滚动、取消检查和收尾清理。
  - 全局尝试计数只在这里维护；首次尝试和重试都调用一次 `single_article_capture_service.py`，并共同受最大尝试次数限制。
  - 保证每次真实尝试恰好一条 `awa_fetch_history`：成功历史由保存事务写入；其他未记录历史的结果由本服务补写失败历史，包括 MITM、窗口、捕获和解析阶段的早期失败。

- `home_article_service.py`
  - 编排主页文章选择。
  - 调用窗口识别、主页信息读取、文章卡片识别、滚动工具。
  - 输出下一篇要采集的 `ArticleTarget`。

- `single_article_capture_service.py`
  - 编排单篇文章的一次采集尝试，是项目核心服务；内部不执行重试。
  - 流程：刷新目标 -> 新建 MITM 子进程 -> 等待 READY -> 点击文章 -> 检测目标标题 -> 关闭文章标签 -> 发送 STOP_CAPTURE -> 接收 RESULT -> 解析保存。
  - 检测到目标文章标题后不额外等待 MITM；STOP_CAPTURE 到达时 HTML 优先，否则使用已暂存的 reference，两者都没有则失败。
  - 在 `finally` 中关闭文章窗口并确认子进程已按正常顺序恢复系统代理；只有子进程异常或超时时才由主流程按快照兜底恢复，再强制终止并 join 残留子进程。
  - 不直接解析 HTML，不直接写 SQLite。

- `mitm_process_control_service.py`
  - 每次采集尝试都新建 MITM 独立子进程和通信通道，不使用进程池或复用旧进程。
  - 编排子进程启动、消息通信、取消和回收，不实现代理底层操作。
  - 负责接收并暂存 PROXY_SNAPSHOT、等待 READY、发送 STOP_CAPTURE/CANCEL、接收 RESULT/FAILED，并把可序列化结果交还主流程。
  - 所有消息校验 `task_id`、`attempt_id`；任务结束时确认子进程已退出，超时才强制终止。
  - MITM 子进程入口放在 `modules/processes/mitm_capture_process.py`。

- `html_parse_save_service.py`
  - 编排 HTML 解析与保存。
  - `capture_type=html` 时直接解析保存。
  - `capture_type=reference` 时先请求 HTML，再统一解析保存。
  - 解析和完整校验通过后，按 `(account_id, article_link)` 决定新增或覆盖。
  - 使用确定性目录和暂存目录，只替换本次成功生成的资源。
  - 文件替换和 SQLite 提交之间使用补偿回滚：数据库失败时恢复旧文件或清理本次新建目录，不依赖后续修复任务消除不一致。
  - 输出文章目录、SQLite 记录 ID、资源清单。

- `comment_collect_service.py`
  - 编排评论采集。
  - 从已保存的文章详情或请求证据中提取评论参数。
  - 保存 `comments/final.json` 和 `comments/assets/`，并更新文章资源清单。
  - 评论文件覆盖与 SQLite 更新使用同一套补偿回滚，保证失败时恢复上一次成功评论资源。

## archive

- `archive_service.py`
  - 提供文章档案查询、资源状态刷新、归档删除等本地数据任务。
  - 不负责联网请求。

- `offline_cache_service.py`
  - 编排离线网页缓存。
  - 调用 Playwright 打开短链、滚动加载、保存 `index.html` 和 `assets/`。
  - 离线文件覆盖与 SQLite 更新使用同一套补偿回滚，保证失败时恢复上一次成功离线资源。

- `archive_export_service.py`
  - 编排 Excel 导出。
  - 从 SQLite 查询文章索引，根据 `archive_dir` 读取本地 `article_detail.json` 补齐统计字段。

## runtime

- `database_init_service.py`
  - 只在程序启动阶段检查当前版本数据库文件是否存在。
  - 根据 `software.data_schema_version` 精确读取 `data/sql/create_script/` 中已经存在的对应脚本，例如 `v2.1` 读取 `create_awa_v2_1.sql`。
  - 程序只执行现有脚本，不生成、拼接或改写建表 SQL。
  - 数据库已存在时检查可以打开且三张必要表存在，然后返回数据库路径。

- `preflight_service.py`
  - 编排运行前预检。
  - 检查数据库、归档目录、MITM 端口、CA 证书、系统依赖；需要时可做一次 HTTPS 连通性检测，但不在每篇采集的代理启停阶段重复执行。

- `system_service.py`
  - 提供系统状态、端口、代理和 CA 证书相关业务入口。
  - 手动代理操作也要遵守文章采集独占锁，避免覆盖采集任务的代理快照。

- `cleanup_service.py`
  - 编排运行时清理。
  - 只清理当前 `task_id` 的临时目录、子进程、系统代理残留和文章详情窗口。
  - 不停止 FastAPI 的 8766 端口；后端退出由桌面程序整体生命周期管理。

## config

- `config_service.py`
  - 作为 API 与 `config/` 之间的业务入口。
  - 编排配置读取、校验、保存和重置，路由不直接操作 YAML。

## task

- `task_manager.py`
  - 管理任务创建、后台执行、取消令牌、状态查询、事件记录和最终结果。
  - 文章采集启动时获取全局采集独占锁，任务结束或异常时释放。
  - 路由只通过 TaskManager 启动文章采集，TaskManager 再调用 ArticleCaptureService。
  - 第一阶段用它替代独立 `workers/` 目录。

# modules

`modules` 是单一能力工具层。每个文件只做一类动作，外部通过类或函数实例化调用。它不编排完整业务流程，不直接写 SQLite，不调用 `services`。

如果某个文件后续变大，再继续拆细；第一阶段不需要为了“看起来很模块化”拆出太多空文件。

## proxy

- `proxy_lifecycle.py`
  - 只封装 MITM 子进程内部的代理与监听启停顺序，不负责创建业务任务或管理主流程。
  - 开启：先启动 MITM，再开启系统代理。
  - 关闭：先恢复/关闭系统代理，再关闭 MITM。
  - 保存启动前代理快照，并在失败、取消和异常退出时执行同样的恢复顺序。

- `ca_certificate.py`
  - 检查、安装、删除 MITM CA 证书。

- `https_probe.py`
  - 仅供程序启动预检、系统配置页手动检测和代理异常排查，不嵌入每次代理开启/关闭流程。
  - 检测 HTTPS 代理连通性。
  - 返回是否可用、失败原因、耗时。

- `mitm_capture.py`
  - 匹配目标文章请求。
  - 提取 response HTML 或 reference。
  - 维护当前 MITM 子进程内的 `CaptureBuffer`：reference 先到时暂存，HTML 到达时升级为 HTML。
  - 收到 STOP_CAPTURE 后冻结变量；HTML/reference 都不存在时返回 `status=failed、capture_type=none`。

## processes

`processes` 放在 `modules` 下，只放独立进程入口和底层进程通信工具，不写业务编排。

- `mitm_capture_process.py`
  - 单篇文章单次尝试的 MITM 独立进程入口；每次新建，完成后退出，不复用。
  - 在子进程中执行代理开启、MITM 监听、捕获、代理关闭。
  - 接收主流程的 START_CAPTURE、STOP_CAPTURE、CANCEL，并校验 `task_id`、`attempt_id`。
  - 向主流程发送 READY、RESULT、FAILED；最终 RESULT 只在系统代理恢复、MITM 停止后发送。
  - 不直接把普通变量共享给其他子进程，统一返回可序列化 `MitmCaptureResult`。

- `process_channel.py`
  - 封装主进程和子进程通信。
  - 提供标准消息封装、发送、等待、序列化校验和超时处理。
  - 主流程作为消息中转者，把 MITM 结果转交解析或保存子任务。

- `process_launcher.py`
  - 启动、等待、终止子进程。
  - 返回 PID、退出码、运行耗时。
  - 正常路径先等待子进程自行退出；取消或超时路径使用已接收的 `ProxySnapshot` 恢复仍指向本次 MITM 的系统代理，再 terminate/kill 并 join，保证 task 结束时没有本次残留进程。

## window

- `wechat_window.py`
  - 区分微信聊天主窗口、公众号主页、微信内置浏览器文章详情窗口。
  - 返回标准窗口信息。

- `home_reader.py`
  - 读取主页公众号名称、简介、原创、朋友关注等信息。

- `article_card_detector.py`
  - 识别主页文章卡片。
  - 输出标题、区域、点击坐标、页面顺序。
  - 过滤视频号、贴图、非文章区域。

- `tab_operator.py`
  - 检测文章详情 tab 是否打开。
  - 只关闭文章详情页，不关闭主页窗口。

- `mouse_operator.py`
  - 执行指定坐标点击。
  - 执行一次鼠标滚轮滚动。
  - 必要时恢复焦点。

## request

- `article_html_requester.py`
  - 使用 reference 参数重新请求文章 HTML。
  - 控制 headers、超时、重试、请求间隔。
  - 不解析、不保存。

- `article_parser.py`
  - 从 HTML 解析公众号名、标题、发布时间、短链、IP 属地、阅读点赞等指标。
  - 无法获取的指标返回空值，不默认写成 `0`。

- `comment_client.py`
  - 请求评论接口。
  - 处理分页、回复评论、评论解析。

## archive

- `archive_path_builder.py`
  - 根据公众号名称、文章短链、发布时间和标题生成确定性文章归档目录。
  - 只使用规范化 `account_name` 和短链生成稳定 `article_key`，不依赖 `account_biz`；相同文章始终映射到同一 `archive_dir`。
  - 处理 Windows 非法文件名和过长路径，不使用 `_1`、`_2` 等递增目录后缀。

- `article_file_store.py`
  - 先在 `data/tmp/<task_id>/` 准备并校验文件，再替换文章目录中本次成功生成的资源。
  - 保存 `article_detail.json`、原始 HTML、请求摘要、评论文件，并保留本次未执行任务的已有资源。
  - 替换旧文件前建立本次备份，提供提交后删除备份和数据库失败后恢复文件的补偿操作。
  - 不写 SQLite。

- `resource_manifest_builder.py`
  - 根据本地真实文件生成 `ResourceManifest`。
  - 标记文章详情、原始 HTML、请求摘要、评论、离线网页、离线资源是否存在。
  - 生成结果用于刷新 SQLite 的 `resource_types_json`；数据库字段只是快速索引。

- `offline_archiver.py`
  - 使用 Playwright 打开文章、滚动加载、保存资源、重写 HTML/CSS 链接。

- `excel_writer.py`
  - 接收标准导出数据并生成 Excel。
  - 不直接查询 SQLite。

## system

- `paths.py`
  - 获取项目根目录、配置路径、数据库目录、归档目录、日志目录、临时目录。
  - 接收 `AppConfig` 或 `TaskContext` 后解析路径，不自行读取 `custom.yaml`。
  - 为每个任务生成 `data/tmp/<task_id>/`，避免任务之间互相清理临时文件。

- `file_ops.py`
  - 目录列表、创建目录、删除文件、删除文件夹。
  - 删除时必须限制在明确允许的路径范围内。

- `process_ports.py`
  - 检查进程、端口占用、端口绑定关系。
  - 只关闭明确属于当前任务或当前项目的进程。

- `time_utils.py`
  - 格式化 SQLite 时间、文章发布时间、目录名时间、当前时间。

# domain

`domain` 用来统一数据对象、状态、错误码和结果结构。它不执行任务，也不访问系统资源。

- `models.py`
  - 定义核心数据对象：`TaskCommand`、`TaskContext`、`ArticleTarget`、`ArticleDetail`、`MitmCaptureResult`、`ResourceManifest`、`ProcessMessage`。

- `enums.py`
  - 定义 `TaskStatus`、`TaskStage`、`CaptureType`、`ResourceType`、`ProcessMessageType`、`ErrorCode`。

- `results.py`
  - 定义 `ToolResult`、`ServiceResult`、`TaskResult`。

- `events.py`
  - 定义 `TaskEvent`。
  - 统一前端进度展示和日志记录。

- `errors.py`
  - 定义业务异常和错误码映射。
  - 避免前端靠中文错误文案判断状态。

# storage

`storage` 只负责 SQLite 初始化和仓储读写，不参与采集流程判断，不解析 HTML，不处理窗口和代理。

## sqlite

- `connection.py`
  - 管理 SQLite 连接、事务、提交、回滚、关闭；每次新建连接后都执行 `PRAGMA foreign_keys = ON`，不能只依赖建表脚本所在连接。

- `schema_loader.py`
  - 根据 `data_schema_version` 从 `data/sql/create_script/` 精确定位已经存在的建表 SQL。
  - 例如 `v2.1` 只允许读取 `create_awa_v2_1.sql`，不扫描“最新脚本”。
  - 返回脚本路径、版本号、脚本内容；不生成或修改 SQL。

- `database_initializer.py`
  - 根据 `database_init_service.py` 传入的数据库路径和现有脚本内容执行建表。
  - 先创建临时数据库，脚本执行成功后重命名为正式数据库。
  - 只负责执行 SQL 和数据库文件操作，不决定脚本版本。

## repositories

- `account_repository.py`
  - 保存和查询公众号记录。
  - 只按唯一 `account_name` 保存、查询和去重公众号，不使用 `account_biz` 参与身份判断。

- `article_repository.py`
  - 保存、查询和更新 `awa_public_articles`。
  - 按唯一键 `(account_id, article_link)` 执行 UPSERT，重复采集更新同一文章 ID。
  - 保存确定性 `archive_dir`，并根据本地资源清单更新 `resource_types_json`。

- `fetch_history_repository.py`
  - 保存联网资源获取历史。
  - 只记录文章详情、评论详情、离线网页等网络获取任务。
  - `status` 只使用 `success` 和 `failed`。
  - 只作为追加式历史记录，不代表文章资源当前是否存在。
  - 支持 `article_id`、`account_id` 为空的早期失败记录；SQLite 不可写时由 TaskManager 事件和日志兜底。

# config

`config` 负责读取、保存、系统默认值和配置对象转换。业务层不要直接读取 YAML 文件。应用入口通过 `services/config/config_service.py` 在启动阶段以 `src/config/system.yaml` 为基础，再用 `data/custom.yaml` 中存在的字段覆盖，后续任务共享同一个冻结的 `AppConfig`；只有明确调用 `reload()` 或恢复系统默认配置时，才允许重新读取并原子替换内存配置。

- `app_config.py`
  - 定义配置对象。
  - 覆盖软件信息、存储信息、代理信息、请求配置、评论配置、离线缓存配置、运行时配置。

- `config_loader.py`
  - 读取 `src/config/system.yaml` 和 `data/custom.yaml`。
  - 以系统 YAML 为基础，深度合并用户 YAML 覆盖值。
  - 返回 `AppConfig`。

- `services/config/config_service.py`
  - 作为应用程序读取配置的统一业务入口。
  - 初始化时加载并保存内存配置，普通 `current` 访问不读取磁盘。
  - 显式重载成功后才替换内存对象；校验失败时继续保留原配置。
  - 恢复默认时先校验 `system.yaml`，再备份并覆盖 `custom.yaml`，成功后同步内存配置。

- `app/runtime_context.py`
  - 应用入口统一创建 `ConfigService`。
  - 将同一个内存 `AppConfig` 传给窗口工厂、单篇捕获设置和后续业务服务。

- `config_writer.py`
  - 保存或恢复用户可修改配置。
  - 恢复默认时先生成 `custom.yaml.bak`，再通过同目录临时文件原子替换 `custom.yaml`。

- `config_validator.py`
  - 校验版本号、路径、端口、请求间隔、重试次数等关键字段。

- `system.yaml`
  - 作为唯一系统默认配置来源。
  - 当 `custom.yaml` 缺少字段时提供基础值，也是“恢复默认”操作的覆盖来源。
