# 安装说明

本文只说明如何从零开始获取项目、安装依赖并启动程序。具体功能说明请看 [`features.md`](./features.md)。

## 1. 环境要求

建议环境：

- Windows 10 / Windows 11
- Python 3.13 或更高版本
- uv，用于管理 Python 环境和依赖
- Git，可选；使用 ZIP 包时不需要 Git

本项目会在项目目录内使用这些运行目录：

- `data/`：用户配置、SQLite 数据库、日志和临时文件
- `.mitmproxy/`：mitmproxy 证书和配置
- `.playwright-browsers/`：Playwright 浏览器
- `storages/`：文章归档数据

## 2. 获取项目

### 方式一：使用 Git clone

进入你希望保存项目的目录，然后执行：

```bash
git clone https://github.com/yeximm/Access_wechat_article.git
cd Access_wechat_article
```

如果你使用自己的 fork 或本地仓库地址，把仓库地址替换成实际地址即可。

### 方式二：使用 ZIP 包

1. 打开项目 GitHub 页面。
2. 点击 `Code`。
3. 点击 `Download ZIP`。
4. 解压到本地目录，例如：

```text
your_local_path\Access_wechat_article
```

建议不要解压到系统保护目录，也尽量避免路径过长。

## 3. 安装 Python 依赖

进入项目根目录：

```bash
cd your_local_path\Access_wechat_article
```

使用 uv 同步依赖：

```bash
uv sync
```

项目依赖以这两个文件为准：

- `pyproject.toml`
- `uv.lock`

本项目不再维护 `requirements.txt`。如果你在旧文档或旧版本中看到 `pip install -r requirements.txt`，请改用 `uv sync`。

不要把依赖安装到系统 Python、用户目录或全局环境。

如果提示找不到 `uv`，请先安装 uv。安装完成后重新打开终端，再回到项目根目录执行 `uv sync`。

## 4. 安装 Playwright 浏览器

Playwright Chromium 浏览器需要安装到项目目录 `.playwright-browsers/`。

PowerShell 中执行：

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH=".playwright-browsers"
uv run playwright install chromium
```

cmd 中执行：

```bat
set PLAYWRIGHT_BROWSERS_PATH=.playwright-browsers
uv run playwright install chromium
```

安装后项目根目录下应出现：

```text
.playwright-browsers/
```

## 5. 启动本地 API 后端

在项目根目录执行：

```bash
uv run python dev_server.py
```

程序会启动本地 FastAPI 服务。`main.py` 保留为旧命令兼容入口，也会转调同一个后端服务；桌面窗口后续由 Tauri 封装承接。


## 6. 安装验证

完成安装后，执行：

```bash
uv run python dev_server.py
```

后端启动后，可以在浏览器访问网页端：

```text
http://127.0.0.1:8766/
```

如果根路径没有自动打开页面，可以访问：

```text
http://127.0.0.1:8766/index.html
```

如果能打开桌面窗口，说明基础运行正常。


## 7. 更新项目

如果使用 Git clone：

```bash
git pull
uv sync
```

如果依赖或 Playwright 版本变化，再执行：

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH=".playwright-browsers"
uv run playwright install chromium
```

如果使用 ZIP 包更新，建议先备份：

```text
data/custom.yaml
data/awa_public.sqlite3
storages/
.mitmproxy/
```

然后再替换项目代码。

## 8. 安装相关常见问题

### uv 命令不存在

说明本机还没有安装 uv，或安装后终端没有刷新环境变量。安装 uv 后重新打开终端，再执行：

```bash
uv --version
```

### Playwright 提示找不到浏览器

重新在项目根目录执行：

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH=".playwright-browsers"
uv run playwright install chromium
```

并确认项目根目录存在 `.playwright-browsers/`。

### 网页端无法访问

网页端依赖 `dev_server.py` 启动的本地服务。请先确认本地 API 后端正在运行，然后访问：

```text
http://127.0.0.1:8766/
```

如果端口被占用，关闭旧的程序进程后重新执行：

```bash
uv run python dev_server.py
```

