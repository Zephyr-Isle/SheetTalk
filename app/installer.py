"""一键安装 / 卸载 / 自启动:桌面设置界面与 install.py 共用的安装逻辑。

职责:
- 注册 Office(Excel)与 WPS 的侧边栏 + 功能区「工具栏按钮」(复用 addin/sideload.py)
- 登记 / 撤销「跟随启动」开机自启(HKCU\\...\\Run)
- 创建 / 删除桌面快捷方式(免命令行,小白双击即用)

对外接口:
  install_all()          -> (ok: bool, summary: str)   装齐:两侧加载项 + 跟随启动 + 桌面快捷方式
  remove_all()           -> (ok: bool, summary: str)   全部移除
  watch_enabled()        -> bool                       是否已登记开机自启
  set_watch(on: bool)    -> (ok: bool, summary: str)   单独开关开机自启

所有注册动作都是幂等的(sideload 内部按固定 GUID/值名覆盖写)。
"""
import contextlib
import importlib.util
import io
import logging
import os
import sys

from app import runtime
from app.web_server import addin_dir

log = logging.getLogger("excelai")

SHORTCUT_NAME = "表答 SheetTalk.lnk"


# ---------- 复用 addin/sideload.py(按文件路径加载,开发态与打包态都可用) ----------

_sideload_cache = None


def _sideload():
    global _sideload_cache
    if _sideload_cache is not None:
        return _sideload_cache
    path = os.path.join(addin_dir(), "sideload.py")
    if not os.path.isfile(path):
        raise FileNotFoundError("找不到注册脚本:%s" % path)
    spec = importlib.util.spec_from_file_location("excelai_sideload", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    _sideload_cache = mod
    return mod


def _quiet(fn, *args):
    """执行 sideload 里的函数并把它的 print 输出收进日志(GUI 下没有控制台)。"""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        fn(*args)
    out = buf.getvalue().strip()
    if out:
        log.info("%s 输出:\n%s", getattr(fn, "__name__", "sideload"), out)
    return out


# ---------- 桌面快捷方式 ----------

def _project_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _icon_ico_path():
    """返回快捷方式图标:统一用 assets/icon.ico(与窗口/托盘/exe 同一设计)。

    - 打包态:图标已由 PyInstaller --icon 内嵌进 exe,直接指向 exe
    - 开发态:优先仓库 assets/icon.ico(多尺寸 16→256)
    - 找不到时回退到 %APPDATA% 现场生成单尺寸 ico(兜底,设计略有差异)
    """
    if runtime.is_frozen():
        return sys.executable
    repo_ico = os.path.join(_project_root(), "assets", "icon.ico")
    if os.path.isfile(repo_ico):
        return repo_ico
    try:
        ico_dir = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "ExcelAI")
        os.makedirs(ico_dir, exist_ok=True)
        ico_path = os.path.join(ico_dir, "appicon.ico")
        if os.path.isfile(ico_path) and os.path.getsize(ico_path) > 0:
            return ico_path
        from app.web_server import _png_icon
        png = _png_icon(32)
        # ICO 容器:ICONDIR + 1 条 ICONDIRENTRY + PNG 数据(Vista 起支持 PNG 压缩图标)
        import struct
        header = struct.pack("<HHH", 0, 1, 1)
        entry = struct.pack("<BBBBHHII", 32, 32, 0, 0, 1, 32, len(png), 6 + 16)
        with open(ico_path, "wb") as f:
            f.write(header + entry + png)
        return ico_path
    except Exception as e:
        log.debug("生成快捷方式图标失败:%s", e)
        return None


def _shortcut_target():
    """返回 (可执行文件, 参数, 工作目录);打包态直接指向 exe。"""
    if runtime.is_frozen():
        return sys.executable, "", os.path.dirname(sys.executable)
    # 开发态:优先 pythonw.exe(无黑窗),参数是 main.py
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    exe = pyw if os.path.isfile(pyw) else sys.executable
    root = _project_root()
    return exe, '"%s"' % os.path.join(root, "main.py"), root


def ensure_shortcut():
    """创建/覆盖桌面快捷方式。返回 (ok, 说明)。"""
    try:
        import win32com.client  # pywin32,requirements 内
    except Exception as e:
        return False, "创建桌面快捷方式失败(缺少 pywin32):%s" % e
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        # SpecialFolders("Desktop") 自动跟随 OneDrive 桌面重定向
        desktop = shell.SpecialFolders("Desktop")
        lnk = shell.CreateShortcut(os.path.join(desktop, SHORTCUT_NAME))
        target, args, workdir = _shortcut_target()
        lnk.TargetPath = target
        if args:
            lnk.Arguments = args
        lnk.WorkingDirectory = workdir
        lnk.Description = "表答 SheetTalk - Excel AI 表格助手"
        icon = _icon_ico_path()
        if icon:
            lnk.IconLocation = "%s,0" % icon
        lnk.Save()
        log.info("桌面快捷方式已就绪:%s", os.path.join(desktop, SHORTCUT_NAME))
        return True, "桌面快捷方式已创建"
    except Exception as e:
        log.exception("创建桌面快捷方式失败")
        return False, "创建桌面快捷方式失败:%s" % e


def remove_shortcut():
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        path = os.path.join(shell.SpecialFolders("Desktop"), SHORTCUT_NAME)
        if os.path.isfile(path):
            os.remove(path)
            return True, "桌面快捷方式已删除"
        return True, "桌面快捷方式不存在,无需删除"
    except Exception as e:
        return False, "删除桌面快捷方式失败:%s" % e


# ---------- 装 / 卸 / 自启动 ----------

def install_all():
    """一键安装:Office 按钮 + WPS 按钮 + 开机自启 + 桌面快捷方式。"""
    problems = []
    try:
        m = _sideload()
    except Exception as e:
        return False, "找不到注册脚本:%s" % e
    for fn_name, label in (("register_office", "Excel 侧"), ("register_wps", "WPS 侧"),
                           ("install_watch", "开机自启")):
        try:
            _quiet(getattr(m, fn_name))
        except Exception as e:
            problems.append("%s注册失败:%s" % (label, e))
    try:
        ok, msg = ensure_shortcut()
        if not ok:
            problems.append(msg)
    except Exception as e:
        problems.append("桌面快捷方式失败:%s" % e)

    if problems:
        return False, "部分步骤未完成:" + ";".join(problems)
    return True, ("安装完成:Excel/WPS 工具栏按钮 + 开机自启 + 桌面快捷方式。"
                  "重启 Excel/WPS 后即可在功能区看到「表答」。")


def remove_all():
    """一键卸载:两侧加载项 + 跟随启动 + 桌面快捷方式。"""
    problems = []
    try:
        m = _sideload()
    except Exception as e:
        return False, "找不到注册脚本:%s" % e
    for fn_name, label in (("remove_office", "Excel 侧"), ("remove_wps", "WPS 侧"),
                           ("remove_watch", "开机自启")):
        try:
            _quiet(getattr(m, fn_name))
        except Exception as e:
            problems.append("%s移除失败:%s" % (label, e))
    try:
        ok, msg = remove_shortcut()
        if not ok:
            problems.append(msg)
    except Exception as e:
        problems.append("桌面快捷方式删除失败:%s" % e)

    if problems:
        return False, "部分步骤未完成:" + ";".join(problems)
    return True, "已卸载:功能区按钮、开机自启与桌面快捷方式均已移除。"


def watch_enabled():
    """当前是否登记了开机自启(HKCU\\...\\Run 里的跟随启动项)。"""
    try:
        m = _sideload()
        if m.winreg is None:
            return False
        with m.winreg.OpenKey(m.winreg.HKEY_CURRENT_USER, m.RUN_KEY, 0,
                              m.winreg.KEY_READ) as k:
            m.winreg.QueryValueEx(k, m.RUN_VALUE)
        return True
    except OSError:
        return False
    except Exception as e:
        log.debug("读取自启状态失败:%s", e)
        return False


def set_watch(on):
    """单独开关开机自启。"""
    try:
        m = _sideload()
    except Exception as e:
        return False, "找不到注册脚本:%s" % e
    try:
        _quiet(m.install_watch if on else m.remove_watch)
    except Exception as e:
        return False, ("开启" if on else "关闭") + "自启动失败:%s" % e
    if on:
        return True, "已开启:开机后自动待命,打开 Excel/WPS 即就绪"
    return True, "已关闭开机自启(本次运行不受影响)"
