"""Qt 应用引导:创建 QApplication,注册控制器,加载 QML,装配托盘后台运行。"""
import http.client
import logging
import os
import sys

# 诊断探针:打包态下 PySide6 绑定缺依赖时,抛出的往往是裸 ImportError(细节被
# shiboken/libpyside 吞掉),这里逐个导入把真凶与完整异常链打到 stderr,
# 便于排查"打包后启动即崩"。正常环境无副作用:这些模块随后本来就会被导入。
import traceback as _tb

# 诊断探针:打包态下 PySide6 绑定缺依赖时,抛出的往往是裸 ImportError(细节被
# shiboken/libpyside 吞掉),这里逐个导入把真凶与完整异常链打到 stderr,
# 便于排查"打包后启动即崩"。正常环境无副作用:这些模块随后本来就会被导入。
import ctypes as _ct


def _probe_loadlib(tag):
    _p = os.path.join(os.path.dirname(__import__("PySide6").__file__), "QtQuick.pyd")
    try:
        _ct.WinDLL(_p)
        sys.stderr.write("[probe:%s] LoadLibrary QtQuick.pyd OK\n" % tag)
        return True
    except OSError as _e:
        sys.stderr.write("[probe:%s] LoadLibrary QtQuick.pyd FAIL: %r\n" % (tag, _e))
        return False


_probe_loadlib("fresh")   # 任何 PySide6 绑定导入之前:纯净进程状态

for _mod in ("QtCore", "QtGui", "QtNetwork", "QtOpenGL", "QtQml", "QtQuick",
             "QtQuickControls2", "QtWidgets"):
    try:
        __import__("PySide6." + _mod)
        if _mod in ("QtQml", "QtOpenGL", "QtNetwork"):
            _probe_loadlib("after-" + _mod)
    except Exception as _e:
        sys.stderr.write("[probe] PySide6.%s FAIL repr=%r args=%r winerror=%r cause=%r\n"
                         % (_mod, _e, _e.args, getattr(_e, "winerror", None),
                            repr(_e.__cause__)[:200]))

from PySide6.QtCore import QTimer, QUrl, Qt
from PySide6.QtGui import (QBrush, QColor, QFont, QFontDatabase, QIcon,
                         QLinearGradient, QPainter, QPen, QPixmap)
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuickControls2 import QQuickStyle
from PySide6.QtWidgets import QApplication, QMenu, QSystemTrayIcon


def qml_dir():
    """QML 目录查找顺序:PyInstaller _MEIPASS → Nuitka/独立目录(exe 旁边)→ 项目根。"""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass and os.path.isdir(os.path.join(meipass, "qml")):
        return os.path.join(meipass, "qml")
    from app import runtime
    bundled = os.path.join(runtime.exe_dir(), "qml")
    if os.path.isdir(bundled):
        return bundled
    return os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "qml")


def _make_icon():
    # 统一图标:优先加载 assets/icon.ico(多尺寸,与桌面快捷方式/exe 同源,
    # 生成器见 assets/make_icon.py),标题栏/托盘可拿到点对点的 16/32px;
    # 找不到再回退到下面的程序绘制(同一设计)。
    bases = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        bases.append(meipass)
    from app import runtime
    bases.append(runtime.exe_dir())
    bases.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    for base in bases:
        ico = os.path.join(base, "assets", "icon.ico")
        if os.path.isfile(ico):
            ic = QIcon(ico)
            if not ic.isNull():
                return ic
    pm = QPixmap(64, 64)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    g = QLinearGradient(0, 0, 64, 64)
    g.setColorAt(0, QColor("#107c41"))   # Excel 绿(与 Main.qml 的 accent 一致)
    g.setColorAt(1, QColor("#0b5c30"))
    p.setBrush(QBrush(g))
    p.setPen(Qt.NoPen)
    # 与 assets/make_icon.py 同一设计:圆角方块几乎占满画布(四周只留 2/64)
    p.drawRoundedRect(2, 2, 60, 60, 15, 15)
    p.setPen(QPen(QColor(255, 255, 255, 235), 3.2, Qt.SolidLine, Qt.RoundCap))
    for y in (19, 32, 45):
        p.drawLine(14, y, 50, y)
    for x in (25, 39):
        p.drawLine(x, 14, x, 50)
    p.end()
    return QIcon(pm)


def _pick_font_family(candidates):
    """从候选字体名里挑第一个系统里真实存在的。

    Win11 首选 Segoe UI Variable,中文界面再回退微软雅黑 UI。
    注意 QML 的 font.family 只接受单个字体名,字体栈要在 Python 侧用
    QFont.setFamilies() 设置,否则整串会被当成一个不存在的字体名而回退到衬线体。
    """
    available = set(QFontDatabase.families())
    picked = [c for c in candidates if c in available]
    return picked or ["Microsoft YaHei UI", "Segoe UI", "sans-serif"]


def run_qt_app(smoke_timeout=0):
    # 界面已改为自绘 Win11 Fluent 风格(qml/Main.qml 内统一令牌 + component),
    # 这里用 Basic 风格作为底层默认,避免 Basic 之外的平台样式(如 Material/Fusion)
    # 覆盖我们显式定义的 background/contentItem 造成观感不一致。
    QQuickStyle.setStyle("Basic")
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    # 用 QApplication(而非 QGuiApplication)是因为托盘 QSystemTrayIcon 属于 QtWidgets;
    # 对 QML 应用无副作用。
    app = QApplication(sys.argv)
    app.setApplicationName("表答 SheetTalk")
    app.setApplicationDisplayName("表答 SheetTalk")
    _app_icon = _make_icon()
    app.setWindowIcon(_app_icon)
    # 关闭最后一个窗口 ≠ 退出(关窗转托盘后台),没有 QGuiApplication.quitOnLastWindowClosed
    # 的默认语义干扰,这里显式关掉,退出只走托盘菜单。
    app.setQuitOnLastWindowClosed(False)

    # 全局字体:Win11 观感的关键(与 qml/Main.qml 的 win.fontStack 保持同一候选顺序)
    _f = QFont()
    _f.setFamilies(_pick_font_family(["Segoe UI Variable Text", "Microsoft YaHei UI", "Segoe UI"]))
    _f.setPixelSize(13)
    app.setFont(_f)

    from app.qt_controller import Controller
    ctrl = Controller()
    app.aboutToQuit.connect(ctrl.shutdown)

    # 启动期自检:关键不变量(确认流/工具注册/校验/检索)失败时在 UI 红色横幅明示,
    # 不带“运行时才炸”的风险;结果同时进 app.log 与 /api/status(侧边栏同步展示)
    _startup_issues = []
    try:
        from app.agent import startup_selfcheck
        from app import runtime
        _startup_issues = list(startup_selfcheck())
        runtime.STARTUP_ISSUES = _startup_issues
        if _startup_issues:
            logging.getLogger("excelai").error(
                "启动自检发现 %d 项异常:%s", len(_startup_issues), "; ".join(_startup_issues))
    except Exception:
        logging.getLogger("excelai").exception("启动自检执行失败")

    # Excel 内侧边栏(Office 加载项)服务,固定 8765 端口(与 addin/manifest.xml 一致)
    # 单实例保护:端口已被本程序占用说明已有一个实例在跑(可能是 --watch 拉起的),
    # 此时通知那个实例把窗口调到前台,再退出,避免开出两个窗口、两份 COM 桥接互相抢占。
    try:
        from app import web_server
        from app.watcher import PORT, app_responding, clear_user_quit
        # 冒烟测试(--smoke)只渲染截图几秒、不发消息不改表,允许在主程序
        # 已运行时另开一个临时窗口;正式启动仍严格单实例。
        # 用 app_responding(严格 HTTP 校验)而不是 TCP 探测:8765 被别的
        # 程序占用时宁可自己起一个(起不来会留日志),也别静默退出。
        if app_responding() and not smoke_timeout:
            logging.getLogger("excelai").info("检测到主程序已在运行,本次启动退出")
            # 主程序多半缩在托盘后台:让老实例把窗口调到前台,避免「双击了没反应」
            try:
                c = http.client.HTTPConnection("127.0.0.1", PORT, timeout=3)
                c.request("GET", "/api/show")
                c.getresponse().read()
                c.close()
            except OSError:
                pass
            return 0
        # 本次要真正运行了:撤销「主动退出」标记,guard 恢复跟随
        clear_user_quit()
        if web_server.create_server(ctrl.bridge, ctrl.session,
                                    on_show=ctrl.requestShowWindow):
            logging.getLogger("excelai").info("加载项服务: http://localhost:8765/addin.html")
    except Exception:
        logging.getLogger("excelai").exception("加载项服务启动失败")

    engine = QQmlApplicationEngine()
    engine.rootContext().setContextProperty("ctrl", ctrl)
    engine.rootContext().setContextProperty("startupIssues", _startup_issues)
    engine.load(QUrl.fromLocalFile(os.path.join(qml_dir(), "Main.qml")))
    if not engine.rootObjects():
        print("QML 加载失败,请查看上方警告", file=sys.stderr)
        return 1

    # ---------------- 托盘:关窗后台运行的唯一入口 ----------------
    win = engine.rootObjects()[0]
    # 任务栏/标题栏图标:给窗口本体也显式设一次。QQuickWindow 不一定会继承
    # QApplication 的窗口图标(打包态曾出现任务栏按钮空白),多设一次成本为零。
    try:
        win.setIcon(_app_icon)
    except Exception:
        logging.getLogger("excelai").debug("设置窗口图标失败", exc_info=True)

    # ---------------- 原生顶边闭环校正(修复「标题栏在屏幕外」) ----------------
    # 实测(Windows11 + Qt6,96% 与 200% 缩放均复现):QML 侧 y 对应的原生窗口
    # 顶边 ≈ y − 30(标题栏高度),只按 QML y 钳制标题栏照样出屏。这里读原生
    # GetWindowRect(Qt 进程是 PMv2 感知,拿到物理坐标)做闭环:顶边不足就抬高
    # QML y,反复几次直到落在屏幕内。启动最初 ~2s 内完成,不影响用户拖动。
    def _clamp_native_top(attempt=0):
        try:
            import ctypes
            from ctypes import wintypes
            rc = wintypes.RECT()
            if not ctypes.windll.user32.GetWindowRect(
                    wintypes.HWND(int(win.winId())), ctypes.byref(rc)):
                raise OSError("GetWindowRect 失败")
            dpr = float(win.devicePixelRatio()) or 1.0
            scr = win.screen()
            target = (scr.availableGeometry().top() + 4) * dpr if scr else 4 * dpr
            delta = target - rc.top
            _log.debug("原生顶边校正 第%d次:top=%s target=%.0f delta=%.0f",
                       attempt + 1, rc.top, target, delta)
            if abs(delta) <= 3:
                return True
            win.setY(win.y() + delta / dpr)
        except Exception:
            _log.debug("原生顶边校正失败", exc_info=True)
            return True     # 修不了就放弃,别无限重试
        if attempt >= 6:
            return True
        QTimer.singleShot(350, lambda: _clamp_native_top(attempt + 1))
        return False

    QTimer.singleShot(350, _clamp_native_top)


    def _show_window():
        # 注意:PySide6 的 QQuickWindow 没有 isMinimized()(调用即抛 AttributeError,
        # 会导致托盘点击永远无法恢复窗口),最小化状态要用 windowStates() 判断。
        states = win.windowStates()
        if states & Qt.WindowMinimized:
            win.setWindowStates((states & ~Qt.WindowMinimized) | Qt.WindowActive)
        else:
            win.show()
        win.raise_()
        win.requestActivate()

    def _quit_from_tray():
        # 留下「用户主动退出」标记:guard 监控看到后不再把窗口拉回来
        from app.watcher import mark_user_quit
        mark_user_quit()
        app.quit()

    def _on_tray_activated(reason):
        # 左键单击/双击都恢复窗口;右键(Context)交给菜单。任何异常都进日志,
        # 便于排查「点了托盘没反应」。
        _log.info("托盘激活:reason=%s", reason)
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            try:
                _show_window()
            except Exception:
                _log.exception("托盘恢复窗口失败")

    _log = logging.getLogger("excelai")
    tray = QSystemTrayIcon(_make_icon())
    tray.setToolTip("表答 SheetTalk")
    menu = QMenu()
    act_show = menu.addAction("显示主窗口")
    act_show.triggered.connect(_show_window)
    menu.addSeparator()
    act_quit = menu.addAction("退出")
    act_quit.triggered.connect(_quit_from_tray)
    tray.setContextMenu(menu)
    tray.activated.connect(_on_tray_activated)
    # 点托盘气泡也恢复窗口
    tray.messageClicked.connect(_show_window)
    ctrl.showWindow.connect(_show_window)
    tray.show()

    first_hide = [False]

    def _on_hidden_to_tray():
        if not first_hide[0]:
            first_hide[0] = True
            tray.showMessage("表答 SheetTalk",
                             "已转入后台运行,Excel 侧边栏服务保持可用。\n"
                             "双击托盘图标重新打开;右键托盘图标可退出。",
                             QSystemTrayIcon.Information, 4000)

    ctrl.hiddenToTray.connect(_on_hidden_to_tray)

    if smoke_timeout:
        # 冒烟测试:渲染几秒后把窗口自截图存盘再退出(不截取用户桌面)
        def _shot_and_quit():
            try:
                img = win.grabWindow()
                out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                   "_gui_shot.png")
                img.save(out)
            except Exception as e:
                print("自截图失败:", e, file=sys.stderr)
            app.quit()
        if "--smoke-chat" in sys.argv:
            # README 截图:启动后自动发一条只读演示消息,流式回复直接渲染进窗口,
            # 截图窗口期(main.py 里放宽到 40s)足够覆盖 LLM 回复
            QTimer.singleShot(1800, lambda: ctrl.sendMessage(
                "分析当前工作表的数据,总结关键发现", False))
        QTimer.singleShot(max(1500, smoke_timeout - 2000), _shot_and_quit)
        QTimer.singleShot(smoke_timeout + 2500, app.quit)
    return app.exec()
