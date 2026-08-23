# Contribution And Issue Guide

Thank you for helping improve Access WeChat Article. This project involves WeChat PC window recognition, local MITM, system proxy, SQLite, local archives, and Playwright offline caching.

To help maintainers locate problems faster, please submit issues that are structured, redacted, and reproducible.

This document focuses on how to submit issues. Pull request guidance is kept to the basics.

## Before Submitting An Issue

Before opening an issue, please do three things:

1. Search existing issues to check whether the same problem has already been reported.
2. Try to reproduce the problem with the latest code or latest release.
3. Prepare the smallest useful reproduction information. Do not upload full databases, full logs, certificates, Cookie values, or links containing sensitive parameters.

If the problem only appears for a specific public account, article, or time range, please mention that in the issue. A public account name, article short link, operation steps, and actual result are usually much more helpful than saying “it does not work.”

Issue URL: https://github.com/yeximm/Access_wechat_article/issues

## Issue Title Suggestions

Use the format “page/function + symptom” when possible.

Good examples:

- `Main service: home window bounces repeatedly after selecting a start date`
- `Diagnostics: window test detects duplicate article cards`
- `Offline cache: images are missing after stateful beta archive`
- `Comment collection: comment count differs from HTML comment count`
- `System settings: default switches do not refresh after changing menus`

Avoid vague titles:

- `Problem`
- `Error`
- `Cannot run`
- `Please help`

## Issue Template

Please use the following format when possible.

```markdown
## Problem Description

Briefly describe the problem you encountered.

## Related Public Account Or Article

- Public account name:
- Article link:
- Article title:
- Publication time:

> If the link contains sensitive parameters such as key, pass_ticket, appmsg_token, uin, or Cookie, redact them first.

## Steps To Reproduce

1. Which page did you open:
2. Which button did you click:
3. Which settings did you select:
4. Which step did the program reach:

## Actual Result

Describe what actually happened.

## Expected Result

Describe what you expected the program to do.

## Runtime Environment

- Operating system:
- Python version:
- WeChat PC version:
- Whether system proxy was enabled:
- Whether the mitmproxy CA certificate was installed:
- Startup command, such as uv run python dev_server.py / uv run python main.py:

## Logs Or Screenshots

Paste key log snippets or upload redacted screenshots.

## What You Have Tried

Describe what you have already tried, such as restarting the backend, reopening the WeChat home page, reinstalling the certificate, or clearing cache.
```

## Extra Information By Problem Type

Different issues need different information. Add the relevant details below depending on the problem type.

### Home Window Recognition / Scrolling / Article Card Issues

Please provide as much as possible:

- Public account name.
- Current date filter mode, such as no date limit, date range, end date, or start date.
- Task count.
- The approximate date currently visible on the home page.
- Whether bounce scrolling, no scrolling, duplicate articles, or missing articles occurred.
- Key lines from diagnostics or main service logs.
- Redacted screenshots, preferably including the article list, date groups, and result dialog.

### Article Detail Capture Issues

Please provide as much as possible:

- Public account name.
- Article title.
- Article short link, preferably in the form `https://mp.weixin.qq.com/s/...`.
- Whether “skip collected records” was selected.
- Whether system proxy and MITM were enabled.
- Final status shown in the detail capture dialog.
- Whether HTML was captured, or only a reference was captured.

### Initial Content Storage / Comment Collection Issues

Please provide as much as possible:

- Article directory path, such as `storages/account_name/article_directory`.
- Whether `article_detail.json` exists.
- Whether `origin/main.html` exists.
- Comment count, reply count, page count, and stop reason from the comment result.
- Key error logs.

Do not upload full HTML, full comment JSON, or full databases directly. Paste only redacted key fields when needed.

### Offline Cache / Image / Video Issues

Please provide as much as possible:

- Original article short link.
- Local `index.html` path.
- Whether you used standard offline archive or “stateful beta.”
- Which resource type is abnormal: image, GIF, normal video, WeChat Channels card, or “Read Original.”
- Screenshot of the local page opened in a browser.
- Key console errors or log snippets containing messages such as “resource skipped” or “download failed.”

If the issue involves video resources, please state whether it is a normal `<video>` / iframe video or a WeChat Channels `mp-common-videosnap` card.

### Installation / Startup / Certificate / Proxy Issues

Please provide as much as possible:

- The command you ran.
- The key part of the full error stack.
- Whether the system proxy was occupied.
- Whether the mitmproxy CA certificate was installed successfully.
- Whether `https://mitm.it/` or HTTPS status check passed.

Do not upload the `.mitmproxy/` directory, certificate files, private keys, or system certificate screenshots.

## Sensitive Information Redaction

Do not include real values of the following in issues, screenshots, logs, or pull requests:

- Cookie
- Set-Cookie
- `key`
- `pass_ticket`
- `appmsg_token`
- `uin`
- `wxtoken`
- Temporary query parameters from complete WeChat article URLs
- SQLite databases
- Certificates or private keys
- Complete article archive directories
- Complete comment files

Recommended redaction:

- For long tokens, keep only a short prefix and suffix, such as `abc...xyz`.
- For Cookie, keep only field names and remove all real values.
- Prefer WeChat article short links. If a long link must be posted, remove parameters such as `key`, `pass_ticket`, and `appmsg_token`.
- Paste only dozens of log lines that explain the problem. Do not upload full log files.

## Good Issue Example

````markdown
## Problem Description

After selecting “start date 2026-08-14” in the main service, the home window bounces twice, then the log says “date group snapshot did not change.”

## Related Public Account Or Article

- Public account name: People’s Daily
- Article link: Not opened yet; the problem occurs on the home article list
- Target date: 2026-08-14

## Steps To Reproduce

1. Open the WeChat PC public account home page.
2. Select “start date” on the main service page.
3. Choose 2026-08-14.
4. Set task count to 20.
5. Click “Start.”

## Actual Result

The home window bounces twice, then the log says the date group snapshot did not change. The search speed is noticeably slower than the window test button.

## Expected Result

The main flow should reuse the window test button’s scrolling and date-positioning logic, quickly approach the target date, and start collecting article cards near that date.

## Runtime Environment

- Operating system: Windows 11
- Startup command: uv run python dev_server.py
- WeChat PC version: To be added
- Whether system proxy was enabled: Automatically enabled by the program
- CA certificate: Installed

## Logs Or Screenshots

Key log:

```text
[14:20:03] WARN Date group snapshot did not change; no new date found after bounce
```
````

## Pull Request Basics

If you plan to submit a PR, please try to:

- Keep one PR focused on one clear problem.
- Explain what functionality changed and which pages or modules are affected.
- Add necessary tests, or explain why local verification is not possible.
- Do not commit personal data, databases, certificates, logs, cached pages, or exported files.
- Update related documentation when changing user-visible behavior.

This project involves local proxy, certificates, and WeChat window control. If your change touches these capabilities, please explain how the system proxy is restored after abnormal exit, how wrong-window closing is avoided, and how sensitive parameters are protected.

## Files That Should Not Be Committed

The following content may contain personal data, certificates, runtime logs, article materials, or short-lived request parameters, and should not be committed to Git:

- `.mitmproxy/`
- `.playwright-browsers/`
- `data/*.sqlite3`
- `data/sql/*.sqlite3`
- `data/logs/`
- `data/tmp/`
- `data/runtime/`
- `storages/`
- `tests/artifacts/`
- `tests/output/`
- Exported `.xlsx`, `.csv`, `.db`, `.sqlite`, or `.sqlite3` files

If test samples are needed, use minimal redacted fake data.
