"""修剪 Nuitka 打包产物中本程序用不到的 Qt 模块(build_exe.bat 打包后自动调用)。

Nuitka 的 pyside6 插件会把整个 QML 数据目录和全套 Qt DLL 带进产物,其中包括完全
用不到的大模块(单个 QtWebEngine 就 195MB)。按下面的黑名单修剪后体积约减半。

修剪范围仅限 dist 目录内;Main.qml 只 import QtQuick/Controls/Layouts/Window,
被删模块均无运行时引用。用法: python -m app.prune_dist dist/main.dist
"""
import os
import shutil
import sys

# DLL 按前缀删(qt6<前缀>*.dll);保留:core/gui/widgets/qml*/quick*/svg/network/
# opengl*/shadertools 等本程序实际用到的模块
DELETE_DLL_PREFIXES = (
    "qt63d", "qt6charts", "qt6concurrent", "qt6datavisualization", "qt6graphs",
    "qt6labs", "qt6location", "qt6multimedia", "qt6pdf", "qt6positioning",
    "qt6quick3d", "qt6quicktest", "qt6quicktimeline", "qt6quickvectorimage",
    "qt6remoteobjects", "qt6scxml", "qt6sensors", "qt6spatialaudio", "qt6sql",
    "qt6statemachine", "qt6test", "qt6texttospeech", "qt6virtualkeyboard",
    "qt6webchannel", "qt6webengine", "qt6websockets", "qt6webview",
)

# PySide6/qml 下整目录删除(Main.qml 只 import QtQuick/Controls/Layouts/Window)
DELETE_QML_DIRS = (
    "Qt3D", "Qt5Compat", "QtCharts", "QtDataVisualization", "QtGraphs",
    "QtLocation", "QtMultimedia", "QtNetwork", "QtPositioning", "QtRemoteObjects",
    "QtScxml", "QtSensors", "QtTest", "QtTextToSpeech", "QtWebChannel",
    "QtWebEngine", "QtWebSockets", "QtWebView", "QtQuick3D",
)


def prune(dist_dir):
    dist_dir = os.path.abspath(dist_dir)
    if not os.path.isdir(dist_dir):
        print("产物目录不存在:", dist_dir)
        return 1
    saved = 0
    root = dist_dir
    for name in os.listdir(root):
        low = name.lower()
        if low.endswith(".dll") and low.startswith(DELETE_DLL_PREFIXES):
            p = os.path.join(root, name)
            saved += os.path.getsize(p)
            os.remove(p)
            print("  删 DLL:", name)
    # addin 数据目录里的 __pycache__(installer 导入 sideload.py 时生成)不进包
    for dirpath, dirnames, _ in os.walk(os.path.join(root, "addin")):
        for d in list(dirnames):
            if d == "__pycache__":
                p = os.path.join(dirpath, d)
                for dp, _, files in os.walk(p):
                    for fn in files:
                        saved += os.path.getsize(os.path.join(dp, fn))
                shutil.rmtree(p, ignore_errors=True)
                print("  删 addin 缓存:", p)
    qml_root = os.path.join(root, "PySide6", "qml")
    if os.path.isdir(qml_root):
        for name in os.listdir(qml_root):
            if name in DELETE_QML_DIRS:
                p = os.path.join(qml_root, name)
                for dirpath, _, files in os.walk(p):
                    for fn in files:
                        saved += os.path.getsize(os.path.join(dirpath, fn))
                shutil.rmtree(p, ignore_errors=True)
                print("  删 QML 目录:", name)
    print("修剪完成,共释放 %.1f MB" % (saved / 1048576))
    return 0


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join("dist", "main.dist")
    sys.exit(prune(target))
