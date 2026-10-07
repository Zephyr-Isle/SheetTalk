"""运行形态探测:开发态 / PyInstaller / Nuitka 统一判断。

对外接口:
  is_frozen() -> bool   是否以打包后的独立程序运行
  exe_dir()   -> str    程序所在目录(打包态=exe 旁边;开发态=项目根)

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
    """程序所在目录:打包态返回 exe 所在目录,开发态返回项目根目录。"""
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
