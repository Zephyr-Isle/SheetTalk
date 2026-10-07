<div align="center">
  <img src="assets/icon.png" width="88" alt="SheetTalk">

# SheetTalk

**AI copilot for Excel & WPS — chat in natural language, act on the spreadsheet you have open**

简体中文 | [English](README.en.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/Python-3.10%2B-green)
![Excel · WPS](https://img.shields.io/badge/Excel%20·%20WPS-107c41)

<img src="assets/screenshot.png" width="380" alt="SheetTalk main window">

</div>

---

A **native Windows desktop app** (PySide6 / QML, Win11 Fluent style) that drives both
**Microsoft Excel** and **WPS Spreadsheets** through COM automation, bringing
Microsoft 365 Copilot-style AI to your spreadsheets: analyze data, write formulas,
create charts, clean data, and format ranges.

Two forms in one app: a **standalone desktop window** and an **in-app sidebar**
(Excel add-in / WPS JS add-in) sharing the same conversation — use both at once.

## Features

- **Conversational editing** — "sort by sales descending and highlight the top 10", done
- **Data analysis** — reads table structure, summarizes key findings and anomalies
- **Formula expert** — generates and fills formulas in batch (written in English, displayed localized)
- **Charts** — picks the right chart type and inserts it
- **Formatting** — bold, colors, number formats, alignment, column width
- **Full Markdown** — headings, tables, code blocks, streamed token by token
- **Plan mode** — complex tasks show an execution plan first
- **Selection as context** — the range you select in Excel is attached automatically
- **Two forms** — desktop window + in-app sidebar, one shared session
- **Dark / light theme** — follows Windows automatically, synced across both UIs
- **Tray background** — closing the window keeps the sidebar service alive; auto-starts with Excel/WPS
- **Excel & WPS** — auto-detects Microsoft Excel → WPS Spreadsheets → legacy WPS
- **Any model** — any OpenAI-compatible endpoint: DeepSeek / GLM / Kimi / Qwen / OpenAI / local Ollama

## Quick start

### Option 1: one-click install (recommended)

1. Double-click **一键安装.bat** (installs deps → registers ribbon buttons → sets up auto-start → creates a desktop shortcut)
2. Open Excel/WPS, click **表答 → 侧边栏** on the ribbon (or just use the desktop window)
3. Click the gear icon, paste an API key (once), then chat in natural language

### Option 2: from source

```bash
cd excel-ai-agent
pip install -r requirements.txt
python main.py
```

### Option 3: portable executable

`build_exe.bat` builds with [Nuitka](https://nuitka.net); the result is a single file
`dist\ExcelAI.exe` (~34 MB, onefile-compressed, Qt/Python runtime bundled). Copy it
anywhere and run, no Python required (still needs an API key on first run). The first
launch unpacks the runtime to `%LOCALAPPDATA%\SheetTalk\<version>` (a few seconds);
later launches reuse it instantly.

## Excel sidebar (Office add-in)

1. Click **一键安装 / 修复** in the app settings (registers the trusted catalog)
2. **Fully quit and restart Excel** (Excel reads trusted catalogs at startup only)
3. Ribbon: **Insert → Get Add-ins → "Shared Folder" → SheetTalk → Add**
   (modern Office no longer honors auto-install; this manual step is needed once)
4. A **表答** group appears at the far right of the **Home** tab → click **侧边栏**

## WPS sidebar (WPS JS add-in)

1. Click **一键安装 / 修复** (registers Office + WPS in one go)
2. **Fully quit and restart WPS Spreadsheets**
3. A **表答** tab appears on the ribbon → click **侧边栏**

WPS 12.1.0.16910+ disables JS add-ins by default: open `office6\cfgs\oem.ini` in the
WPS install folder, add `JsApiPlugin=true` under `[support]`, save and restart WPS
(the installer detects and hints this automatically).

## How it works

```
┌──────────────────────┐  Qt signals (thread-safe) ┌──────────────────┐  COM bridge  ┌──────────────────┐
│  Native window (QML) │ ◄───────────────────────  │ Python controller│ ◄──────────► │  Excel / WPS      │
│  chat UI + Markdown  │                           │ Agent + tool calls│              │  (your workbook)  │
└──────────────────────┘                           │  ↕ LLM streaming │              └──────────────────┘
                                                   │                  ▲
┌──────────────────────┐                           │ shared session    │
│ In-app sidebar       │ ◄── 127.0.0.1:8765 ───────┘ (built-in HTTP)  │
│ (add-in task pane)   │        SSE + REST
└──────────────────────┘
```

The AI drives the workbook through **74 spreadsheet tools** (read/write/formulas/charts/
formatting/sorting/find-replace/row-column insert-delete/freeze panes/merge/protection/
pivot tables/conditional formats/hyperlinks/visibility/sheets…) via function calling; every step is streamed into the chat.
All COM calls run serially on a dedicated thread, fully isolated from the UI. The local
service listens on `127.0.0.1:8765` only — nothing leaves your machine.

## Requirements

- Windows 10/11
- Python 3.10+ (verified on 3.14; not needed for the packaged build)
- Microsoft Excel (any recent version) **or** WPS Spreadsheets
- Any OpenAI-compatible API key (DeepSeek has free quota; Ollama works fully offline)

## Security

- API keys stay in local `config.json` (`%APPDATA%\ExcelAI` for packaged builds) and are
  sent only to the endpoint you configure
- The built-in HTTP service binds to `127.0.0.1:8765` only — never exposed to the LAN
- All spreadsheet operations run locally via COM; no third-party servers involved

## Project layout

```
excel-ai-agent/
├── main.py               # entry point (--smoke / --watch / --watch-guard)
├── app/
│   ├── qt_app.py         # Qt bootstrap (Fluent style, unified icon, tray, QML)
│   ├── qt_controller.py  # controller: signals ↔ agent/COM threads, status polling
│   ├── chat.py           # shared chat session (stream throttle, events, cancel)
│   ├── web_server.py     # built-in HTTP service (sidebar page + REST/SSE, 127.0.0.1)
│   ├── excel_bridge.py   # COM bridge (Excel/WPS dual compat, serialized)
│   ├── agent.py          # agent loop (context injection + tool calls + events)
│   ├── tools.py          # 74 spreadsheet tools: registry + tool-RAG delivery + execution
│   ├── llm.py            # OpenAI-compatible client (SSE + function calling)
│   ├── watcher.py        # follow-launch monitor (auto start/stop with Excel/WPS)
│   ├── installer.py      # one-click install/uninstall (add-ins + autostart + shortcut)
│   ├── runtime.py        # runtime detection (source / PyInstaller / Nuitka)
│   └── config.py         # config.json read/write
├── qml/Main.qml          # Win11 Fluent UI (native Markdown rendering)
├── addin/                # in-app sidebar
│   ├── manifest.xml      # Office add-in manifest (standard VersionOverrides)
│   ├── addin.html        # sidebar page (streaming chat + Markdown + dark theme)
│   ├── sideload.py/.bat  # registration: Office trusted catalog + WPS JS add-in
│   └── wps/              # WPS JS add-in sources (ribbon.xml + main.js)
├── assets/               # unified icon (icon.ico / icon.png) + generator make_icon.py
├── selftest.py           # offline self-test (no Excel / API key needed)
├── live_demo_test.py     # live smoke test (needs Excel open)
├── install.py            # one-click install/uninstall CLI entry
├── requirements.txt
├── build_exe.bat         # Nuitka build script
└── LICENSE               # MIT license
```

## Development

```bash
python selftest.py          # offline self-check: no Excel, no API key needed
python live_demo_test.py    # end-to-end: requires Excel/WPS open (creates a demo workbook)
python main.py --smoke      # UI smoke test: exits after ~6s and writes _gui_shot.png
build_exe.bat               # Nuitka onefile build → dist\ExcelAI.exe (~34 MB)
```

Conventions and architecture details for contributors: see [AGENTS.md](AGENTS.md).

## License

Released under the [MIT License](LICENSE) — free to use, modify and redistribute, including commercially, with attribution.

Copyright (c) 2026 Zephyr-Isle
