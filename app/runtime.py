"""运行形态探测:开发态 / PyInstaller / Nuitka 统一判断。

对外接口:
  is_frozen()  -> bool  是否以打包后的独立程序运行
  exe_dir()    -> str   数据所在目录(打包态=解包目录;开发态=项目根)
  launch_exe() -> str   打包态=用户启动的那个 exe 的绝对路径;开发态=None

两种打包器的约定不同:
  - PyInstaller:sys.frozen=True,数据在 sys._MEIPASS 下
  - Nuitka:无 sys.frozen(独立模块内有 __compiled__ 标记),数据文件就在 exe 旁边
所有 frozen 判断一律走这里,别再直接 getattr(sys, "frozen", False)。
"""
import os
import sys

# Nuitka 编译后本模块的全局命名空间里会出现 __compiled__;开发态没有
_IS_NUITKA = "__compiled__" in globals()

# 启动期自检结果(qt_app 启动时写入):非空 = 有关键不变量被破坏,
# web_server 会把它放进 /api/status,UI 以红色横幅明示
STARTUP_ISSUES = []


def is_frozen():
    return bool(getattr(sys, "frozen", False)) or _IS_NUITKA


def exe_dir():
    """数据文件所在目录:打包态返回解包目录,开发态返回项目根目录。

    打包态用 __file__ 而不是 sys.executable:qml/addin/assets 被解包到
    运行目录,只有编译模块的 __file__ 指向数据真正所在的位置,而
    sys.executable 指向的是解包目录里一个并不存在的 python.exe(见 launch_exe)。
    """
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def launch_exe():
    r"""返回「用户实际启动的那个可执行文件」的绝对路径;开发态返回 None。

    只有打包态才需要它(自启项 / 桌面快捷方式 / guard 拉起主程序)。
    Nuitka onefile 下 sys.executable 与 __file__ 都指向解包缓存目录里的副本:
    4.2.2 实测 sys.executable = %LOCALAPPDATA%\SheetTalk\<版本>\python.exe,
    而该目录里只有 ExcelAI.exe —— 这个 python.exe 根本不存在,照它登记开机
    自启 / Popen 只会静默失败(用户看到的是「跟随启动没生效」)。
    可靠的是 sys.argv[0]:Nuitka 手册把 onefile 的 argv[0] 定义为
    「the original executable path」,实测也正是用户双击的那个 exe。
    新版 Nuitka 另下发 NUITKA_ONEFILE_BINARY(4.2.2 为空),有则优先。
    """
    if not is_frozen():
        return None
    binary = os.environ.get("NUITKA_ONEFILE_BINARY") or ""
    if binary.lower().endswith(".exe"):
        return os.path.abspath(binary)
    raw = sys.argv[0] or ""
    if raw.lower().endswith(".exe"):
        return os.path.abspath(raw)
    # 异常启动方式(拿不到 exe 形态的 argv[0])才退回 sys.executable
    return os.path.abspath(sys.executable)
