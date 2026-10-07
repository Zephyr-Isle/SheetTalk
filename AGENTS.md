# AGENTS.md

> 给 AI 编码代理(Cline / Cursor / Copilot 等)的项目说明。改动代码前请先读完本文件。

## 项目概览

**表答 SheetTalk**(原「Excel AI 助手」)—— 一个*原生 Windows 桌面应用*(PySide6/Qt + QML Material Design),
通过 COM 自动化同时驱动 **Microsoft Excel** 与 **WPS 表格**,提供类 Microsoft 365 Copilot 的
对话式 AI 能力(读写表格、公式、图表、格式、排序、数据分析)。

- 语言/运行时:**Python 3.10+**(已在 3.14 验证),仅支持 **Windows**(依赖 COM / pywin32)。
- UI:QML(`qml/Main.qml`)+ Material 风格;Python 侧通过 `QObject` 属性/信号/槽暴露给 QML。
- AI:任意 **OpenAI 兼容** `/chat/completions` 接口(SSE 流式 + function calling),
  内置 DeepSeek / 智谱 / Kimi / 通义 / OpenAI / Ollama 预设。
- 形态:① 独立桌面窗口;② Excel 内侧边栏(Office Web 加载项,由内置 HTTP 服务提供页面),
  两者共享同一份对话历史与 COM 桥接。

## 环境与命令

```bash
# 安装依赖(无 pyproject / setup.py,只有 requirements.txt)
pip install -r requirements.txt        # PySide6>=6.6, requests>=2.31, pywin32>=306

# 运行
python main.py                         # 启动桌面应用
python main.py --watch                 # 跟随启动:检测到 Excel/WPS 才拉起主程序(单次)
python main.py --watch-guard           # 跟随启动(常驻):表格在则主程序在,表格全关则收掉
python main.py --smoke                 # 启动 ~6 秒后自动退出并截图 _gui_shot.png(冒烟测试)
python addin/sideload.py               # 注册加载项(Office+WPS),并自动装上「跟随启动」
python addin/sideload.py remove        # 反向:两侧加载项 + 跟随启动一并移除
python addin/sideload.py watch/unwatch # 只单独登记/取消「跟随启动」(pythonw 无窗口常驻监控)
# 跟随启动监控无控制台,日志只在 %APPDATA%\ExcelAI\logs\app.log

# 测试(项目没有 pytest,就这两个独立脚本)
python selftest.py                     # 离线自检:不需 Excel、不需 API Key,打印 SELFTEST OK
python live_demo_test.py               # 真机端到端:必须先打开 Excel/WPS,会新建演示工作簿,打印 LIVE TEST OK

# 打包
build_exe.bat                          # Nuitka 打包 → dist\main.dist\ExcelAI.exe(约 420MB)
```

**验证改动的最低成本路径**:改了 Agent/工具/事件流 → 跑 `python selftest.py`;
改了 COM 桥接 → 跑 `python live_demo_test.py`(需真机 Excel);
改了 UI → 跑 `python main.py --smoke` 看 `_gui_shot.png` 截图;改了加载项页面 → 启动应用后访问
`http://localhost:8765/addin.html`。

## 目录结构

```
excel-ai-agent/
├── main.py                # 入口:argparse(--smoke / --watch / --watch-guard)、logging、拉起 Qt 应用
├── app/
│   ├── qt_app.py          # Qt 引导:Fluent 风格、统一图标、托盘后台、QML 加载、启动加载项 HTTP 服务
│   ├── qt_controller.py   # Controller(QObject):暴露给 QML 的属性/槽;信号跨线程回 UI
│   ├── chat.py            # ChatSession:框架无关会话(历史、流式节流 ~120ms、事件翻译)
│   ├── agent.py           # run_agent:系统提示词 + LLM 循环 + 工具调用(事件流)
│   ├── tools.py           # TOOLS(42 个 function-calling schema)+ Executor(执行与描述)
│   ├── llm.py             # OpenAI 兼容客户端(SSE 流式累积、非流式回退、LLMError)
│   ├── excel_bridge.py    # ExcelBridge + ComWorker:COM 桥接,探测 Excel→WPS(Ket)→WPS(ET)
│   ├── web_server.py      # 内置 HTTP 服务(仅 127.0.0.1:8765):加载项页面 + REST/SSE、samples 归一化
│   ├── watcher.py         # 跟随启动监控(guard 常驻 / once 单次)、TCP 探测、退出标记
│   ├── installer.py       # 一键安装/卸载(加载项 + 开机自启 + 桌面快捷方式),供 GUI 与 install.py 共用
│   ├── runtime.py         # 运行形态探测(开发态 / PyInstaller / Nuitka 的统一 is_frozen / exe_dir)
│   ├── icon_design.py     # 统一图标设计光栅化(ico/png/HTTP 接口共用,别另画图标)
│   ├── prune_dist.py      # Nuitka 产物修剪(build_exe.bat 打包后自动调用)
│   └── config.py          # 配置读写(PRESETS / DEFAULTS / 加锁读写 config.json)
├── qml/Main.qml           # 全部界面(单文件 QML:对话列表、设置对话框、Markdown 渲染)
├── addin/                 # 侧边栏:manifest.xml(Office)、addin.html(两侧共用)、
│                         #   wps/(WPS JS 加载项:ribbon.xml + main.js + js/main.js)、
│                         #   sideload.py/.bat(两侧注册)、marked/purify
├── assets/                # 统一图标 icon.ico / icon.png + 生成器 make_icon.py + README 截图
├── selftest.py            # 离线自检(fake Bridge + fake LLM,断言事件序列 + 加固断言)
├── live_demo_test.py      # 真机冒烟(写入→公式→格式→排序→图表→回读;--clean 才关残留工作簿)
├── install.py             # 一键安装/卸载命令行入口(供 一键安装.bat / 一键卸载.bat 调用)
├── requirements.txt
└── build_exe.bat          # Nuitka 打包(含 qml/、addin/、assets/ 资源,产物 dist\main.dist)
```

## 架构与线程模型

```
QML(Main.qml)  ◄── Qt 信号(QueuedConnection)──  qt_controller.Controller
      │                                              │   ▲
      │ ctrl.sendMessage()/属性绑定                  │   │ emit 回调(event dict)
      ▼                                              ▼   │
                                  chat.ChatSession ──► agent.run_agent
                                                          │  ↕ requests(SSE)
                                             tools.Executor ──► excel_bridge.ExcelBridge
                                                                       │
                                                          ComWorker(专用线程,CoInitialize)
                                                                       ▼
                                                          Excel / WPS(COM)
```

**硬性约束(改动时必须遵守)**:

1. **所有 COM 调用必须经 `ExcelBridge` 的 `ComWorker`**(`bridge.submit(fn)`),它把调用串行固定在
   一个已 `CoInitialize` 的线程上。**不要**在 Qt 主线程、Agent 线程或 HTTP 线程直接碰 `win32com`。
2. **Agent 循环跑在独立 Python 线程**,通过 `emit(event_dict)` 推事件;事件跨到 UI 必须走
   **Qt 信号**(`ctrl.chatEvent.emit(...)`),绝不能从工作线程直接操作 QML。
3. **`ChatSession.send` 有非阻塞锁**:忙时发 `toast` 事件并返回 `False`,别再叠加并发保护。
4. `shutdown()` 会置 `_quitting`,退出阶段丢弃事件,避免关闭时 crash——新增的信号发射也要检查它。

## 事件协议(改 UI / 改 Agent 时保持一致)

`agent.run_agent` 的 emit 事件 `type`:
`status / plan / text_start / delta / note_flush / step / step_result / error / cancelled / done`

经 `ChatSession` 翻译后的 UI 事件 `kind`:
`user_msg / plan / text_start / delta / note_flush / step / step_result / status / error / cancelled / done / toast / need_settings / clear`

- `delta`:流式增量文本,由 ChatSession 按 ~120ms 节流后才下发。
- `step` / `step_result`:工具调用卡片(`sid` 关联,`summary` / `brief` / `ok`)。
  `step` 还带 `ref`(来源区域,如 `Sheet1!A1:C20`),由 `Executor.describe_ref` 生成,
  用于步骤卡片右侧的引用溯源标签。
- `plan`:计划模式下模型先给出的执行计划,UI 渲染成蓝色计划卡片。
- `cancelled`:用户点「停止」后 agent 收尾,UI 显示"已停止"。
- `note_flush`:把"工具调用前的叙述"降级为小字备注,正文留给最终回答。
- 新增事件类型时:`agent.py` 产出 → `chat.py` 翻译 → `qml/Main.qml` 的 `onChatEvent` switch 处理
  → `addin/addin.html` 的对应处理函数,四处都要改。

### 中止(Stop)的实现约定

`run_agent(..., cancel=None)` 收一个 `threading.Event`。检查点只有两处:
**每轮循环开头**、**每个工具调用之前**。绝不能放在工具执行中间 —— 那样工作簿会停在
半改状态。上一次已流出的正文要用 `acc.clear()`(而不是 `acc = []`)落进 history,
否则闭包内 `acc` 会被当成局部变量而抛 `UnboundLocalError`。

UI 侧:桌面端 `ctrl.cancelRun()` → `session.cancel()`;
加载项 `POST /api/cancel`。两端的发送按钮在 busy 时都变成灰色「停止」。


## 代码风格与约定

- **注释、docstring、UI 文案、日志、工具 description 全部用简体中文**;标识符用英文。
- 文件顶部必须有模块 docstring,说明职责与对外接口(参考现有 `app/*.py`)。
- 依赖只用 `requirements.txt` 里的:`PySide6`、`requests`、`pywin32`。**不要引入新的第三方库**
  (QML 侧同理:只用 QtQuick / QtQuick.Controls / QtQuick.Layouts 自带能力)。
- 类型标注非强制(桥接层用了 `from __future__ import annotations`),风格保持与邻近代码一致。
- 错误处理范式:
  - 面向用户的可读错误 → 抛 `excel_bridge.BridgeError`,由 `Executor.execute` 捕获转成 `{"error": ...}`。
  - 工具**永远返回 dict**,不能抛异常(`selftest.py` 专门断言了这一点)。
  - LLM 层错误 → 抛 `llm.LLMError`,由 `agent.run_agent` 转成 `error` 事件。
- 单文件 QML:`qml/Main.qml` 已 30KB+,新增界面优先加在现有结构里(属性、组件、`Connections` 事件接线),
  颜色用顶部已定义的语义属性(`surface / outline / subText / okColor / grad1...`),不要散落硬编码色值。

### 如何新增一个表格工具

1. `app/tools.py` 的 `TOOLS` 列表加一条 OpenAI function schema(name/description/parameters,中文描述)。
2. `app/tools.py` 的 `Executor` 加 `_t_<name>(self, ...)` 方法,转调 `self.bridge.*`;可选在
   `Executor.describe` 里加一句中文摘要、在 `brief_result` 里加一种结果摘要。
3. `app/excel_bridge.py` 加对应桥接方法,内部必须经 `self.worker.submit(...)`,
   失败抛 `BridgeError(中文原因)`。
4. 若引入新的操作规则,同步更新 `app/agent.py` 的 `SYSTEM_TEMPLATE`(`{context}` 占位符是
   `str.replace` 替换的,别写成 f-string)。
5. 跑 `python selftest.py` 验证。

### 如何改 LLM / 提示词

- `app/llm.py`:流式优先,HTTP 400 且错误文本含 stream/sse 时自动回退非流式;
  SSE 行必须按 UTF-8 解码(否则中文乱码,见源码注释)。
- `app/agent.py`:`HISTORY_TURNS = 16`(历史轮数)、单条内容截断 4000 字、工具结果截断 4000 字、
  `max_steps` 默认 30、下限 5、上限 500 —— 调整这些常量时同步改 `agent.py` 底部的超限提示文案
  与 `config.max_steps` 的校验。

## 配置与密钥

- 配置文件:开发态在项目根 `config.json`;打包态在 `%APPDATA%\ExcelAI\config.json`。
- 字段以 `config.DEFAULTS` 为准:`provider / base_url / model / api_key / theme_mode / samples / max_steps / plan_mode`;
  `save()` 只写入 `DEFAULTS` 里存在的键,**新增配置项必须先加进 `DEFAULTS`**。
- `theme_mode`:`auto`(默认,跟随系统亮暗,Python 侧 3s 轮询 `AppsUseLightTheme`)/ `light` / `dark`;
  旧键 `theme` 保留但不再使用。`samples`:输入框上方示例按钮列表,`None`=内置默认,`[]`=用户清空。
- 主题令牌集中在 `qml/Main.qml` 顶部(`accent` = Excel 绿 `#107c41`),新增界面颜色一律引用
  `win.*` 语义属性;未显式设色的 Label 会在深色下隐形(有 `palette` 兜底,但仍建议显式写 color)。
- ⚠️ **当前仓库根目录的 `config.json` 里含有真实 API Key。不要把它提交到远程仓库或分享出去;
  建议改为 `config.example.json` 并把 `config.json` 加入 `.gitignore`,同时作废该 Key。**
- 无任何对外监听:HTTP 服务只绑 `127.0.0.1:8765`(端口必须与 `addin/manifest.xml` 一致)。

## 加载项(Excel / WPS 侧边栏)

两种宿主、两套注册机制,但**共用同一个 `addin.html` 页面与同一个 8765 服务**:

| 宿主 | 机制 | 注册函数 | 落地目录 |
|---|---|---|---|
| Office(Excel) | Office Web Add-in 受信任目录 | `register_office()` | `%APPDATA%\ExcelAI\addin-catalog` |
| WPS 表格 | WPS JS 加载项(jsaddons) | `register_wps()` | `%APPDATA%\kingsoft\wps\jsaddons\ExcelAI_1.0` |

- `python addin/sideload.py [office|wps|remove …]` 默认两侧都注册**并自动登记「跟随启动」**;`remove`(不带参数)会连跟随启动一起撤销;`sideload.bat` 只是转发。
- `addin/manifest.xml` 的 `SourceLocation` 指向 `http://localhost:8765/addin.html`,
  改端口必须三处同步:manifest、`app/web_server.py`、README。
- `addin.html` 用本地 `marked.min.js` + `purify.min.js` 渲染 Markdown;改页面后注意
  DOMPurify 白名单,避免破坏渲染。
- **受信任目录有新旧两套注册表布局**:`Wef\TrustedCatalogs\{GUID}`(子键,含 `Id`/`Url`/`Flags=1`,
  `Url` **必须是 UNC 路径**)+ 旧的 `Wef\TrustedCatalog`(值名=本地路径,DWORD 1)。
  新版 Office(约 2024 起)只认前者;`register_office()` 两套都写,`remove_office()` 两套都删。
- **`Wef\AutoInstallAddins` 自动安装在新版已失效**:该机制是老版本 Office 的,
  新版(16.0.20430+,2026-10 实测)注册后只把加载项记进 `Wef\AddinLifecycle`
  (Id 出现即「清单已通过校验、目录已枚举」的证据),**并不会真正安装** ——
  必须在 Excel 里「插入 → 获取加载项 → 共享文件夹 → 添加」手动点一次,之后按用户
  常驻。键还是照写(兼容老版本),`remove_office()` 一并撤销。
- **改 manifest 必须递增 `<Version>` 且清 `Wef\Cache`**:Office 按 Id+版本缓存清单,
  不清缓存就一直用旧清单,新加的功能区按钮不会出现。`_clear_wef_cache()` 遍历
  `Wef\Cache\{Excel,Word,PowerPoint}` 删子键。注意缓存挂在 **`Wef`** 下,
  不是 `Wef\TrustedCatalog` 下 —— 写错会静默无效(代码里已用 `WEF_BASE`)。
- **manifest 命名空间分两套**:`Icon`/`Action`/`TaskpaneId`/`SourceLocation`/`Image`/`Url`
  属于 `bt:`(officeappbasictypes);`Label`/`Supertip`/`Hosts`/`DefaultSettings`/`Commands`
  属于默认命名空间。漏 `bt:` 前缀 Office 会**直接拒绝加载整个清单**,症状是加载项根本不出来。
- **功能区按钮必须写在 `<VersionOverrides>` 里**(taskpaneappversionoverrides → Hosts →
  `Host xsi:type="Workbook"` → DesktopFormFactor → ExtensionPoint),控件要
  `xsi:type="Button"`,文案走 Resources 的 `bt:ShortStrings`/`bt:LongStrings`。
  顶层裸写 `<Commands>`/`<ExtensionPoint>` 不符合 schema —— 2026-10 按旧结构怎么注册
  按钮都不出现,重写成标准结构才修好(Version 升到 1.3.0.0)。
- `Url` 由 `_to_unc()` 从本地路径推导(`C:\x` → `\\localhost\C$\x`)。**改目录位置要同步这里**,
  否则 Office 静默忽略。若 C$ 形式不可读(部分机器的管理共享对普通权限进程关闭),
  `register_office()` 会自动创建专用共享 `ExcelAI$`(弹一次 UAC,永久生效)并改用
  `\\localhost\ExcelAI$`;`remove_office()` 会顺带删共享。
- 排查时注意:PowerShell 的 `Get-ItemProperty` 对含 `{}` 的路径解析不可靠,
  用 `reg query "HKCU\...\TrustedCatalogs" /s` 才准。
- WPS 侧文件:`addin/wps/ribbon.xml`(功能区按钮)、`addin/wps/js/main.js`
  (`toggleExcelAiPane`,真正实现)、`addin/wps/main.js`(垫片,仅当 WPS 生成的 index.html
  引入的是根目录 main.js 时才兜底)。**不要在 `addin/wps/` 里手写 `index.html`**,WPS 会自己生成。
- WPS 侧任务窗格地址是 `addin.html?host=wps`;该参数让页面跳过 Office JS 加载。
  改 `addin.html` 顶部脚本时保留这个分支,否则 WPS 里会产生无谓的外网请求。
- WPS 12.1.0.16910+ 默认关闭 JS 加载项:`_check_oem_ini()` 只检测并提示改
  `office6\cfgs\oem.ini` 的 `JsApiPlugin=true`,**不要**自动写 Program Files 下的文件。
- `publish.xml` 可能损坏或缺根标签:`_read_publish()` 会备份成 `.bad` 并重建;
  增删条目只操作 `name == "ExcelAI"` 的那一项,别动别人的加载项。

## 常见坑

- **单实例**:`app/qt_app.py` 启动时探测 `127.0.0.1:8765/api/status`,带 `addin=true` 说明
  已有实例(含 `--watch` 拉起的)就直接退出。避免两个窗口、两份 COM 桥接互相抢占。
- **`GetActiveObject` 只能拿到「一个」Excel 实例**:用户多开 Excel 时,它返回的是
  注册在 ROT 里的那个,不一定是用户正在看的那份 —— 表现是「状态栏显示 A 文件,
  实际改的是 B 文件」。所以 `_wb()` 的取值顺序是:用户显式指定(`use_workbook`)
  → `ActiveWindow.Parent`(用户眼前这个窗口)→ `ActiveWorkbook` → 最近打开的。
  `_snapshot` 额外返回 `open_workbooks` 与 `app_visible`,供界面提示改的是哪个文件。
- **`app.Visible=False` 的残留实例不能操作**:这种实例会**静默丢弃**结构性修改
  (删工作表、RemoveDuplicates 都返回成功但数据没变),非常难查。调试时不要在
  用户正在使用的 Excel 里 `new_workbook()`;确实要建,先确认实例可见。
  另外发现行为异常时先看 `app.Visible`,再用任务管理器确认是不是多开了 Excel 进程。
- **跟随启动**(`app/watcher.py`):只枚举进程名,不碰 COM。用
  `win32ts.WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 1, 0)`——参数固定为
  `(0, 1, 0)`,其他组合会报「参数错误」;pywin32 **没有** 封装 `ToolHelp32Snapshot`。
  拉起 GUI 必须带 `DETACHED_PROCESS`,否则子进程继承控制台句柄会把调用方终端卡住。
- **`app_running()` 是 TCP 探测,禁止换成 HTTP**:`/api/status` 要过 COM 队列且
  urllib 的 opener 会把首次请求时的代理配置缓存到进程退出 —— 两个坑都会让 guard
  误判「主程序不在」并反复拉起重复实例(症状:app.log 以秒级/分钟级周期刷
  「检测到…已启动 / 已拉起主程序 / 检测到主程序已在运行,本次启动退出」,而
  curl / 新进程探测都正常)。TCP 握手由内核完成,对代理与 COM 忙碌免疫。
  严格的 HTTP 校验单独放在 `app_responding()`(带 addin 标识),只用于单实例判定。
- **关窗 = 托盘后台运行**:QML `onClosing` 拦截后只 `hide()`,进程不退出(8765
  侧边栏服务保持可用);真正退出只走托盘菜单「退出」,退出前写
  `user_quit.flag`(见 `watcher.mark_user_quit`),guard 见到标记就**不再拉起**,
  直到主程序被再次启动(`launch_app()`/`qt_app` 启动时清标记)。别把这个标记
  机制删了 —— 否则用户退出后窗口又会被 guard 弹回来。
- **窗口落点保护**:QML 里预显示阶段设置的坐标在 DPI 缩放下会被二次换算,
  窗口可能整体落到屏幕外(标题栏不可达)。`Main.qml` 的 `clampIntoScreen()`
  在每次显示后把窗口钳回当前屏幕可见区,别删。
- 单实例 + `launch_app()` 构成幂等保护:重复注册开机自启不会拉出多个主程序。
- **打包用 Nuitka**(`--enable-plugin=pyside6`,见 build_exe.bat):产物
  `dist\main.dist\`;pyside6 插件会因 QtQuickControls2 自动带上 qml 系列插件与模块。
  Nuitka **不设 sys.frozen**,frozen 判断一律走 `app/runtime.py` 的 `is_frozen()`
  (开发态/PyInstaller/Nuitka 三态统一);数据文件 qml/addin/assets 就在 exe 旁边,
  由 `runtime.exe_dir()` 查找 —— 别按 PyInstaller 的 `_MEIPASS` 习惯写路径。
- 冻结态的配置/痕迹目录是 `%APPDATA%\ExcelAI`,开发态是项目根 —— 两态的
  config.json、user_quit.flag 互不相通,开发/打包混用时注意。
- **samples 的形状是三端契约**:config 里存 `{icon, text, prompt}` **对象数组**
  (桌面端设置区按对象渲染,直接用 modelData.icon/text);侧边栏设置按
  「提示词一行一条」编辑;`web_server._normalize_samples` 在保存入口统一归一化
  (字符串/畸形输入也转成对象,空数组转成 null=恢复内置默认)。三端任何一处
  修改 samples 的读写,都必须保持这个形状,否则另一端会渲染出 [object Object]。
- **HTTP 请求体上限 2MB**(`_read_json`):超出直接回 413。对话消息远用不到,
  防的是异常/恶意请求把内存打爆;新增路由时保留这个防线。

- `pythoncom`/`win32com` 只能在 `ComWorker` 线程调用,否则偶发 `CoInitialize has not been called`。
- Excel 弹出模态对话框会阻塞 COM:`ComWorker.submit` 超时(90s)会抛
  `BridgeError("表格程序没有响应…")`,不要把它当 bug 修。

### pywin32 / Excel COM 的几个坑(全部实测踩过,别再犯)

- **`ComWorker` 必须泵消息**:`CoInitialize` 后只阻塞在队列上会让 Excel 属 STA 的重入
  调用失败(表现为 `Validation.Add` 报「发生意外」,而普通读写正常,极易漏测)。
  现在空闲时 `pythoncom.PumpWaitingMessages()`。
- **`Range.Address` / `Range.End` 是「属性」不是方法**:`rng.Address(True, True)` 会抛
  `'str' object is not callable`;要「工作表名!地址」只能自己拼 `"%s!%s" % (ws.Name, rng.Address)`。
- **`wb.PivotTables().Count` 会抛「找不到成员」**:动态派发下这个 collection 成员访问不可用。
  透视表命名改用时间戳,不要去数现有透视表。
- **`Range.Sort`/`Range.Copy` 的部分具名可选参数会被动态派发丢弃**(实测 Header/Key2/
  Order2/Copy 的 After 全部无效):Sort 排完必须**校验表头位置**,不对就还原原始块并回退
  Python 侧排序(`_py_sort_rows`:中文按拼音、多键稳定排序;回退会把公式变为值,结果带
  note);copy_sheet 不用 Copy,改「新建表 + Cells.Copy」。凡新增用到带可选参数的 COM
  方法,先在隔离实例里实测参数是否真的生效,别假设具名参数可用。
- **透视值字段必须用 `AddDataField(field, 标签, 函数)` 一次完成**;先设
  `Orientation=4` 再设 `Function` 会报「不能设置 Function 属性」。
  另外 `pt.PivotFields("1")`(数字索引)是非法调用,只能用字段名。
- **透视表落点要留 ≥3 列空间**:只找「第一个空列」会让 Excel 报目标区域无效,
  并弹模态对话框**卡死后续所有 COM 调用**。现在按「源区域与 UsedRange 的右边界 +1」落点。
- **区域上有 AutoFilter / ListObject 会影响后续操作**:建透视表与去重前都要先
  `ws.AutoFilterMode = False` 并把压在区域上的 ListObject `Unlist()`。
- **`Validation` 的操作顺序**:`Delete()` 之后**不能**立刻设 `IgnoreBlank`/`InCellDropdown`
  (抛「发生意外」),必须先 `Add` 出规则再设属性。
- **`RemoveDuplicates` 会「返回成功但一行没删」**:且 `Columns` 必须是列序号 **tuple**,
  传逗号字符串会抛「发生意外」。现在先试原生、检测行数没变则回退到 Python 侧
  (`_dup_row_indexes` 读数据比对后从后往前删行)。
- **强杀 EXCEL.EXE 会污染 win32com 缓存**:之后 `GetActiveObject` 报
  `(-2147221021, 操作无法使用)`,删掉 `%LOCALAPPDATA%\Temp\gen_py` 即可恢复。
  另外 Excel 无打开工作簿时 `GetActiveObject` 也可能失败,先 `Workbooks.Add()`。
- 读取单次上限约 500 格、写入单次上限 10 万格(工具 description 与系统提示词都依赖这个约定)。
- `write_range` 中以 `=` 开头的字符串会按**公式**写入;纯文本别带等号(系统提示词第 4 条)。
- 公式一律英文函数名 + 英文逗号,由 Excel/WPS 自动本地化显示。
- 修改工具 schema 后,旧会话历史里可能残留不匹配的 tool_calls,必要时提醒用户"清空对话"。
- `__pycache__/` 目录存在(含已删除的 `server.py` 残留),生成物不要入库。

## 安全/隐私红线

- API Key 只存本机、只发往用户自填的 `base_url`;不要新增任何外发遥测或第三方域名。
- 对表格的操作全部走本地 COM;不要引入"上传表格数据到云端"的路径。
- 新增 HTTP 路由必须维持仅监听 127.0.0.1;`addin.html` 是本地可信页面,注入的 HTML 仍需净化。
