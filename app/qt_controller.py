"""Qt 控制器:把 Excel 桥接与对话会话暴露给 QML。

- 所有 COM 调用走 bridge 自带的 COM 工作线程
- 每条消息的 Agent 循环跑在独立 Python 线程,事件经 Qt 信号(队列连接)送到 QML
- 流式增量由 ChatSession 统一节流(~120ms)
"""
import threading

from PySide6.QtCore import Property, QObject, QTimer, Signal, Slot

from app import config, installer, llm
from app.chat import ChatSession
from app.excel_bridge import ExcelBridge

EMPTY_STATUS = {"connected": False, "host_label": None, "workbook": None}


def _read_system_dark():
    """读 Windows 系统外观:AppsUseLightTheme=0 表示系统当前是深色。读不到时按亮色。"""
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return int(val) == 0
    except Exception:
        return False


class Controller(QObject):
    chatEvent = Signal("QVariantMap")
    testFinished = Signal(bool, str)
    statusChanged = Signal()
    busyChanged = Signal()
    themeChanged = Signal()
    installBusyChanged = Signal()
    # 窗口被关成后台(托盘)后发出,qt_app 用来弹一次托盘气泡提示
    hiddenToTray = Signal()
    # HTTP 线程请求把主窗口调回前台(/api/show;跨线程,走队列连接)
    showWindow = Signal()

    def __init__(self):
        super().__init__()
        self.bridge = ExcelBridge()
        self.session = ChatSession(self.bridge)
        self._busy = False
        self._status = dict(EMPTY_STATUS)
        self._statusRefreshing = False
        self._quitting = False
        self._install_busy = False
        # 主题可能被侧边栏(设置页)修改:轮询对比,变化即发 themeChanged 让桌面跟随
        self._last_theme_mode = config.load().get("theme_mode", "auto")

        self._timer = QTimer(self)
        self._timer.timeout.connect(self.refreshStatus)
        self._timer.start(4000)
        QTimer.singleShot(200, self.refreshStatus)

        # 跟随系统亮暗:轮询注册表(3s),变化时发 themeChanged,QML 的 win.dark 会重算
        self._system_dark = _read_system_dark()
        self._theme_timer = QTimer(self)
        self._theme_timer.timeout.connect(self._poll_system_theme)
        self._theme_timer.start(3000)

        import os
        if os.environ.get("EXCELAI_SMOKE_CHAT"):
            QTimer.singleShot(1500, self._smoke_chat)

    # ---------- 暴露给 QML 的属性 ----------

    def _get_busy(self):
        return self._busy
    busy = Property(bool, _get_busy, notify=busyChanged)

    # ---------- 一键安装忙碌态(防止重复点击;工作线程里改值,用信号跨回 UI) ----------

    def _set_install_busy(self, on):
        if self._install_busy != on:
            self._install_busy = on
            if not self._quitting:
                self.installBusyChanged.emit()

    def _get_install_busy(self):
        return self._install_busy
    installBusy = Property(bool, _get_install_busy, notify=installBusyChanged)

    def _get_statusLine(self):
        s = self._status
        if s.get("connected") and s.get("workbook"):
            return f'{s["host_label"]} · {s["workbook"]}'
        if s.get("connected"):
            return f'{s.get("host_label")} · 无工作簿'
        return "未连接表格"
    statusLine = Property(str, _get_statusLine, notify=statusChanged)

    def _get_connected(self):
        return bool(self._status.get("connected") and self._status.get("workbook"))
    connected = Property(bool, _get_connected, notify=statusChanged)

    def _get_warn(self):
        return bool(self._status.get("connected")) and not self._status.get("workbook")
    warnState = Property(bool, _get_warn, notify=statusChanged)

    def _get_dark(self):
        return config.load().get("theme_mode") == "dark"
    darkTheme = Property(bool, _get_dark, notify=themeChanged)

    def _poll_system_theme(self):
        dark = _read_system_dark()
        if dark != self._system_dark:
            self._system_dark = dark
            if not self._quitting:
                self.themeChanged.emit()

    def _get_system_dark(self):
        """Windows 系统当前是否深色(theme_mode=auto 时界面跟随它)。"""
        return self._system_dark
    systemDark = Property(bool, _get_system_dark, notify=themeChanged)

    def _get_theme_mode(self):
        """界面主题模式:auto(跟随系统,默认)/ light / dark。"""
        m = config.load().get("theme_mode", "auto")
        return m if m in ("auto", "light", "dark") else "auto"
    themeMode = Property(str, _get_theme_mode, notify=themeChanged)

    def _get_selection_ref(self):
        """当前选区引用,如「Sheet1!$A$1:$C$10」;无选区返回空串。"""
        s = self._status
        sel = s.get("selection")
        if not (s.get("connected") and sel):
            return ""
        return "%s!%s" % (s.get("active_sheet") or "?", sel)
    selectionRef = Property(str, _get_selection_ref, notify=statusChanged)

    def _get_selection_dims(self):
        """把 $A$1:$C$10 这类选区解析成「10 行 × 3 列」;解析不出返回空串。"""
        import re
        sel = (self._status or {}).get("selection") or ""
        try:
            parts = sel.split(":")
            if len(parts) != 2:
                return ""
            rc = []
            for part in parts:
                m = re.match(r"\$?([A-Z]+)\$?(\d+)$", part.strip())
                if not m:
                    return ""
                col = 0
                for ch in m.group(1):
                    col = col * 26 + (ord(ch) - 64)
                rc.append((int(m.group(2)), col))
            rows = abs(rc[1][0] - rc[0][0]) + 1
            cols = abs(rc[1][1] - rc[0][1]) + 1
            return "%d 行 × %d 列" % (rows, cols)
        except Exception:
            return ""
    selectionDims = Property(str, _get_selection_dims, notify=statusChanged)

    # ---------- 状态轮询 / 连接 ----------

    @Slot()
    def refreshStatus(self):
        if self._statusRefreshing:
            return
        self._statusRefreshing = True

        def work():
            try:
                s = self.bridge.status()
            except Exception:
                s = dict(EMPTY_STATUS)
            self._status = s
            # 主题可能被侧边栏(设置页)修改:变化时让桌面端立即跟随
            theme = config.load().get("theme_mode", "auto")
            theme_changed = theme != self._last_theme_mode
            self._last_theme_mode = theme
            self._statusRefreshing = False
            if not self._quitting:
                self.statusChanged.emit()
                if theme_changed:
                    self.themeChanged.emit()
        threading.Thread(target=work, daemon=True).start()

    @Slot()
    def reconnect(self):
        self.refreshStatus()

    @Slot(str)
    def launch(self, host):
        def work():
            try:
                self.bridge.launch(host)
            except Exception as e:
                self._emit_ui({"kind": "toast", "text": str(e)})
            self.refreshStatus()
        threading.Thread(target=work, daemon=True).start()

    # ---------- 对话 ----------

    def _emit_ui(self, ev):
        if not self._quitting:
            self.chatEvent.emit(ev)

    @Slot(str, bool)
    def sendMessage(self, text, with_selection=True):
        text = (text or "").strip()
        if not text:
            return
        if self._busy:
            self._emit_ui({"kind": "toast", "text": "正在处理上一条请求,请稍候…"})
            return
        self._busy = True
        self.busyChanged.emit()

        def work():
            try:
                self.session.send(text, self._emit_ui, use_selection=bool(with_selection))
            finally:
                self._busy = False
                if not self._quitting:
                    self.busyChanged.emit()
        threading.Thread(target=work, daemon=True).start()

    @Slot()
    def cancelRun(self):
        """中止当前这一轮(对齐 Copilot 输入框右下角的 Stop)。"""
        if not self._busy:
            return
        self.session.cancel()
        self._emit_ui({"kind": "status", "text": "正在停止…"})

    @Slot()
    def clearChat(self):
        if self._busy:
            # 忙碌中清空会让运行中的回合把回复写进新 history,产生孤立上下文;
            # 与侧边栏同口径:先停止再清空。
            self._emit_ui({"kind": "toast", "text": "正在处理请求,请先停止再清空"})
            return
        self.session.clear()
        self._emit_ui({"kind": "clear"})

    # ---------- 安装 / 自启动(设置页「一键安装」区块) ----------

    def _get_watch_enabled(self):
        """是否已登记开机自启(设置页开关的回显)。"""
        try:
            return installer.watch_enabled()
        except Exception:
            return False
    watchEnabled = Property(bool, _get_watch_enabled, notify=statusChanged)

    @Slot()
    def installAddins(self):
        """一键安装:Excel/WPS 工具栏按钮 + 开机自启 + 桌面快捷方式。"""
        if self._install_busy:
            self._emit_ui({"kind": "toast", "text": "正在安装,请稍候…"})
            return
        self._set_install_busy(True)
        self._emit_ui({"kind": "toast", "text": "正在安装工具栏按钮与自启动…"})

        def work():
            ok, msg = installer.install_all()
            self._set_install_busy(False)
            if not self._quitting:
                self._emit_ui({"kind": "toast", "text": msg})
                self.statusChanged.emit()
        threading.Thread(target=work, daemon=True).start()

    @Slot()
    def removeAddins(self):
        if self._install_busy:
            self._emit_ui({"kind": "toast", "text": "\u6b63\u5728\u5904\u7406,\u8bf7\u7a0d\u5019\u2026"})
            return
        self._set_install_busy(True)
        self._emit_ui({"kind": "toast", "text": "\u6b63\u5728\u79fb\u9664\u5de5\u5177\u680f\u6309\u94ae\u4e0e\u81ea\u542f\u52a8\u2026"})
        def work():
            ok, msg = installer.remove_all()
            self._set_install_busy(False)
            if not self._quitting:
                self._emit_ui({"kind": "toast", "text": msg})
                self.statusChanged.emit()
        threading.Thread(target=work, daemon=True).start()


    @Slot(bool)
    def setWatchEnabled(self, on):
        """开关开机自启;状态变化通过 statusChanged 让开关回显真实值。"""
        def work():
            ok, msg = installer.set_watch(bool(on))
            if not self._quitting:
                self._emit_ui({"kind": "toast", "text": msg})
                self.statusChanged.emit()
        threading.Thread(target=work, daemon=True).start()

    # ---------- 设置 ----------

    @Slot(result="QVariantMap")
    def loadSettings(self):
        return dict(config.load())

    @Slot(result="QVariantMap")
    def presets(self):
        return dict(config.PRESETS)

    @Slot(str, str, str, str, int, bool)
    def saveSettings(self, provider, base_url, model, api_key, max_steps, plan_mode=False):
        # 与 agent.py 的钳制口径保持一致,并兜住 QML 传来的非数字
        try:
            steps = int(max_steps)
        except (TypeError, ValueError):
            steps = 30
        steps = max(5, min(steps, 500))
        config.save({"provider": provider, "base_url": base_url.strip(),
                     "model": model.strip(), "api_key": api_key.strip(),
                     "max_steps": steps, "plan_mode": bool(plan_mode)})
        self.statusChanged.emit()

    @Slot(bool)
    def setPlanMode(self, on):
        config.save({"plan_mode": bool(on)})

    @Slot(str)
    def setThemeMode(self, mode):
        """界面主题:auto=跟随系统亮暗 / light / dark,保存后 themeChanged 立即生效。"""
        if mode in ("auto", "light", "dark"):
            config.save({"theme_mode": mode})
            self.themeChanged.emit()

    @Slot("QVariantList")
    def saveSamples(self, items):
        """持久化示例按钮(增删即时保存);空数组表示用户清空。"""
        config.save({"samples": [dict(i) for i in (items or [])]})

    @Slot(str, str, str)
    def testLlm(self, base_url, model, api_key):
        s = config.load()
        base_url = base_url.strip() or s.get("base_url", "")
        model = model.strip() or s.get("model", "")
        api_key = api_key.strip() if api_key.strip() else s.get("api_key", "")

        def work():
            try:
                llm.chat_completions(base_url, api_key, model,
                                     [{"role": "user", "content": "回复:ok"}],
                                     max_tokens=8, timeout=30)
                self.testFinished.emit(True, "✓ 连接成功")
            except llm.LLMError as e:
                self.testFinished.emit(False, "✗ " + str(e)[:300])
        threading.Thread(target=work, daemon=True).start()

    def shutdown(self):
        self._quitting = True

    # ---------- 托盘后台运行 ----------

    @Slot()
    def windowHiddenToTray(self):
        """QML 关闭按钮回调:窗口转后台而不是退出进程(侧边栏服务保持可用)。"""
        self.hiddenToTray.emit()

    @Slot(result="QVariantMap")
    def selectionPreview(self):
        """选区芯片点击:读取当前框选的地址与内容预览(短超时,不长时间卡 UI)。"""
        try:
            return dict(self.bridge.get_selection(timeout=5) or {})
        except Exception:
            return {}

    @Slot(str, bool, bool)
    def confirmRun(self, sid, allowed, remember=False):
        """确认条回调:回传危险工具二次确认结果(只是置事件,线程安全)。

        remember=True 时该工具在本会话内不再重复确认。
        """
        self.session.confirm(sid, allowed, remember)

    @Slot()
    def requestShowWindow(self):
        """/api/show 由 HTTP 线程调用;信号队列连接,窗口操作在主线程执行。"""
        self.showWindow.emit()

    # ---------- UI 冒烟测试钩子(EXCELAI_SMOKE_CHAT=1) ----------

    def _smoke_chat(self):
        emit = self._emit_ui
        # 用户消息带选区附件(验证附件行布局);text_start 后让事件循环继续跑,
        # --smoke 的定时截图(4s)恰好落在停顿期内,捕获 AI 卡片的加载动画。
        # 注意:不能用 time.sleep —— 它在主线程里会冻住事件循环,截图定时器也被推迟。
        self._busy = True
        self.busyChanged.emit()
        emit({"kind": "user_msg", "text": "分析当前表格,告诉我销售额最高的三个月",
              "with_selection": True})
        emit({"kind": "text_start"})
        # 危险操作确认条(验证 Codex 风格下拉 + 红色确认按钮的 UI)
        emit({"kind": "confirm", "sid": "smoke-confirm", "tool": "delete_sheet",
              "summary": "删除工作表:Sheet1(含 12 行数据)"})
        from PySide6.QtCore import QTimer
        QTimer.singleShot(3400, lambda: self._smoke_chat_rest(emit))

    def _smoke_chat_rest(self, emit):
        import time as _t
        emit({"kind": "status", "text": "已连接 Microsoft Excel,正在分析…"})
        for part in ("我先", "看一下", "数据…"):
            emit({"kind": "delta", "text": part})
            _t.sleep(0.08)
        emit({"kind": "note_flush"})
        emit({"kind": "step", "sid": "s1", "summary": "读取 活动表!A1:D20"})
        emit({"kind": "step_result", "sid": "s1", "ok": True, "brief": "✓ 返回 20 行数据"})
        emit({"kind": "text_start"})
        md = ("## 结论\n\n销售额最高的三个月如下:\n\n"
              "| 月份 | 销售额 | 环比 |\n|---|---|---|\n"
              "| 2025-12 | 242,000 | +3.0% |\n| 2025-11 | 235,000 | +7.8% |\n| 2025-10 | 218,000 | -9.5% |\n\n"
              "其中 **2025-12** 最高;可用下面公式复现:\n\n```excel\n=SUMIFS(B:B,A:A,\"2025-12\")\n```\n\n"
              "需要我把这三个月标成绿色吗?")
        for i in range(0, len(md), 36):
            emit({"kind": "delta", "text": md[i:i + 36]})
            _t.sleep(0.03)
        self._busy = False
        if not self._quitting:
            self.busyChanged.emit()
        emit({"kind": "done"})
