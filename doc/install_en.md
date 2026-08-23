# Installation Guide

This document explains how to obtain the project, install dependencies, install the Playwright browser, and start the application from scratch. For feature details, see [Feature Guide](./features_en.md). For the project overview, see [README_en.md](./README_en.md).

## 1. Requirements

Recommended environment:

- Windows 10 / Windows 11
- Python 3.13 or later
- uv, used for Python environment and dependency management
- Git, optional if you use the ZIP package

The project uses these runtime directories inside the project root:

- `data/`: user configuration, SQLite database, logs, and temporary files
- `.mitmproxy/`: mitmproxy certificates and configuration
- `.playwright-browsers/`: Playwright browsers
- `storages/`: article archive data

## 2. Get The Project

### Option A: Git Clone

Go to the directory where you want to store the project, then run:

```bash
git clone https://github.com/yeximm/Access_wechat_article.git
cd Access_wechat_article
```

If you use your own fork or local repository URL, replace the repository URL accordingly.

### Option B: ZIP Package

1. Open the project GitHub page.
2. Click `Code`.
3. Click `Download ZIP`.
4. Extract it to a local directory, for example:

```text
your_local_path\Access_wechat_article
```

Avoid extracting the project into a system-protected directory. It is also recommended to avoid overly long paths.

## 3. Install Python Dependencies

Enter the project root:

```bash
cd your_local_path\Access_wechat_article
```

Sync dependencies with uv:

```bash
uv sync
```

Project dependencies are defined by:

- `pyproject.toml`
- `uv.lock`

This project no longer maintains `requirements.txt`. If old documentation or older versions mention `pip install -r requirements.txt`, use `uv sync` instead.

Do not install dependencies into system Python, user directories, or global environments.

If `uv` is not found, install uv first, reopen your terminal, return to the project root, and run `uv sync` again.

## 4. Install Playwright Browser

Playwright Chromium should be installed into the project directory `.playwright-browsers/`.

Run in PowerShell:

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH=".playwright-browsers"
uv run playwright install chromium
```

Run in cmd:

```bat
set PLAYWRIGHT_BROWSERS_PATH=.playwright-browsers
uv run playwright install chromium
```

After installation, the project root should contain:

```text
.playwright-browsers/
```

## 5. Start The Desktop Application

Run from the project root:

```bash
uv run python dev_server.py
```

The program starts the local FastAPI service. `main.py` remains as a compatibility entrypoint and delegates to the same backend service; the desktop window will be handled by the Tauri packaging path.

## 6. Verify Installation

After installation, run:

```bash
uv run python dev_server.py
```

After the backend starts, you can also open the local web interface in your browser:

```text
http://127.0.0.1:8766/
```

If the root path does not open automatically, try:

```text
http://127.0.0.1:8766/index.html
```

If the desktop window opens, the basic runtime is ready.

## 7. Update The Project

If you use Git clone:

```bash
git pull
uv sync
```

If dependencies or the Playwright version change, run again:

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH=".playwright-browsers"
uv run playwright install chromium
```

If you update from a ZIP package, back up these files and directories first:

```text
data/custom.yaml
data/awa_public.sqlite3
storages/
.mitmproxy/
```

Then replace the project code.

## 8. Installation FAQ

### uv Command Not Found

This usually means uv is not installed, or the terminal has not refreshed environment variables. Install uv, reopen the terminal, and run:

```bash
uv --version
```

### Playwright Cannot Find Browser

Run again from the project root:

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH=".playwright-browsers"
uv run playwright install chromium
```

Make sure `.playwright-browsers/` exists in the project root.

### Local Web Interface Cannot Be Opened

The local web interface depends on the service started by `dev_server.py`. Make sure the local API backend is running, then visit:

```text
http://127.0.0.1:8766/
```

If the port is occupied, close the old process and run again:

```bash
uv run python dev_server.py
```
