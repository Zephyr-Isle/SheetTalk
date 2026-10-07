"""修剪 Nuitka 打包产物中本程序用不到的 Qt 模块(build_exe.bat 打包后自动调用)。

Nuitka 的 pyside6 插件会把整个 QML 数据目录和全套 Qt DLL 带进产物。按下面的
黑名单修剪后,产物只保留本程序实际用到的部分。用法: python -m app.prune_dist dist/main.dist

删减依据(改清单前先核对):
  - Main.qml 只 import QtQuick/Controls/Layouts/Window,QtQuick.Controls 用
    QQuickStyle.setStyle("Basic") 显式钉死 Basic 样式 → 其余样式全可删;
  - qt6quick.dll 硬依赖 qt6network/qt6opengl/qt6qmlmeta/qt6qmlmodels.dll
    (objdump -p 实测),这四个 DLL 必须保留;
  - opengl32sw.dll 是动态加载的软件 OpenGL 兜底,Win10/11 走 D3D11/12 RHI,
    删除不影响加载,仅在无 GPU 的极端设备上失去兜底。
被删模块均无运行时引用;删错的表现是 exe 启动即崩,用 --smoke 实测把关。
"""
import os
import shutil
import sys

# DLL 按前缀删(qt6<前缀>*.dll)
DELETE_DLL_PREFIXES = (
    "qt63d", "qt6charts", "qt6concurrent", "qt6datavisualization", "qt6graphs",
    "qt6labs", "qt6location", "qt6multimedia", "qt6pdf", "qt6positioning",
    "qt6quick3d", "qt6quickdialogs", "qt6quickparticles", "qt6quickshapes",
    "qt6quicktest", "qt6quicktimeline", "qt6quickvectorimage",
    "qt6remoteobjects", "qt6scxml", "qt6sensors", "qt6shadertools", "qt6spatialaudio",
    "qt6sql", "qt6statemachine", "qt6svg", "qt6test", "qt6texttospeech",
    "qt6virtualkeyboard", "qt6webchannel", "qt6webengine", "qt6websockets", "qt6webview",
    "qt6quickeffects",
    # 仅被动态加载、且无任何 import 方的 QML 插件。注意 qt6qmlworkerscript 删不得:
    # qt6qmlmeta.dll(→qt6quick.dll 的硬依赖)硬链接它,缺了 QtQuick 绑定导入即崩
    # (实测:报裸 ImportError,极难排查)。qt6qmlnetwork/qmllocalstorage/qmlxmllistmodel
    # 无任何链接方,可删。
    "qt6qmllocalstorage", "qt6qmlnetwork", "qt6qmlxmllistmodel",
    # 软件 OpenGL 兜底(20MB),动态加载,D3D11 RHI 默认路径用不到
    "opengl32sw",
)

# QtQuick.Controls 只用 Basic 样式(qt_app 显式 setStyle("Basic")),其余样式 DLL 全删。
# basicstyleimpl 是 Basic 的实现库,qmldir 静态引用,缺了 QML 加载失败(实测),必须保留
KEEP_QC2_DLLS = {"qt6quickcontrols2.dll", "qt6quickcontrols2impl.dll",
                 "qt6quickcontrols2basic.dll", "qt6quickcontrols2basicstyleimpl.dll"}

# PySide6/qml 下整目录删除
DELETE_QML_DIRS = (
    "Qt",  # Qt/labs 系列_dll 已按前缀删,目录里只剩 qml 壳
    "Qt3D", "Qt5Compat", "QtCharts", "QtDataVisualization", "QtGraphs",
    "QtLocation", "QtMultimedia", "QtNetwork", "QtPositioning", "QtRemoteObjects",
    "QtScxml", "QtSensors", "QtTest", "QtTextToSpeech", "QtWebChannel",
    "QtWebEngine", "QtWebSockets", "QtWebView", "QtQuick3D",
)

# qml 根之下的子目录删除(相对 PySide6/qml):Quick 的边角模块与其余控件样式
DELETE_QML_SUBDIRS = (
    "QtQuick/Dialogs", "QtQuick/Effects", "QtQuick/LocalStorage",
    "QtQuick/NativeStyle", "QtQuick/Particles", "QtQuick/Pdf",
    "QtQuick/Scene2D", "QtQuick/Scene3D", "QtQuick/Shapes",
    "QtQuick/Timeline", "QtQuick/VectorImage", "QtQuick/VirtualKeyboard",
    "QtQuick/tooling",
    "QtQuick/Controls/FluentWinUI3", "QtQuick/Controls/Fusion",
    "QtQuick/Controls/Imagine", "QtQuick/Controls/iOS", "QtQuick/Controls/macOS",
    "QtQuick/Controls/Material", "QtQuick/Controls/NativeStyle",
    "QtQuick/Controls/Universal", "QtQuick/Controls/Windows",
    "QtQuick/Controls/designer",
)

# qml 根下仅服务设计器/静态检查的工具文件(运行时不需要)
DELETE_QML_FILES = ("QtQuick/plugins.qmltypes",)

# 插件目录:图标引擎只用 ico/png(内置),TLS 走 python 的 ssl(不走 QtNetwork),
# styles 是 QtWidgets 的原生样式(托盘菜单会退回 Fusion 观感,可接受)
DELETE_PLUGIN_DIRS = ("iconengines", "tls", "styles")
# imageformats / platforms 只保留白名单内的插件
IMAGEFORMATS_KEEP = {"qico.dll"}
PLATFORMS_KEEP = {"qwindows.dll"}

# 产物根目录里可直接删的第三方包(cryptography 10MB,源码无一处 import;
# https 证书链由 requests 的 certifi 提供,不受影响)
DELETE_ROOT_DIRS = ("cryptography",)


def _rmtree_size(path):
    total = 0
    for dirpath, _, files in os.walk(path):
        for fn in files:
            total += os.path.getsize(os.path.join(dirpath, fn))
    shutil.rmtree(path, ignore_errors=True)
    return total


def prune(dist_dir):
    dist_dir = os.path.abspath(dist_dir)
    if not os.path.isdir(dist_dir):
        print("产物目录不存在:", dist_dir)
        return 1
    saved = 0
    root = dist_dir
    for name in os.listdir(root):
        low = name.lower()
        if not low.endswith(".dll"):
            continue
        if low.startswith(DELETE_DLL_PREFIXES) or (
            low.startswith("qt6quickcontrols2") and low not in KEEP_QC2_DLLS
        ):
            p = os.path.join(root, name)
            saved += os.path.getsize(p)
            os.remove(p)
            print("  删 DLL:", name)
    # 第三方包整目录
    for name in DELETE_ROOT_DIRS:
        p = os.path.join(root, name)
        if os.path.isdir(p):
            saved += _rmtree_size(p)
            print("  删 包:", name)
    # addin 数据目录里的 __pycache__(installer 导入 sideload.py 时生成)不进包
    for dirpath, dirnames, _ in os.walk(os.path.join(root, "addin")):
        for d in list(dirnames):
            if d == "__pycache__":
                p = os.path.join(dirpath, d)
                saved += _rmtree_size(p)
                print("  删 addin 缓存:", p)
    qml_root = os.path.join(root, "PySide6", "qml")
    if os.path.isdir(qml_root):
        for name in os.listdir(qml_root):
            if name in DELETE_QML_DIRS:
                p = os.path.join(qml_root, name)
                saved += _rmtree_size(p)
                print("  删 QML 目录:", name)
        for rel in DELETE_QML_SUBDIRS:
            p = os.path.join(qml_root, *rel.split("/"))
            if os.path.isdir(p):
                saved += _rmtree_size(p)
                print("  删 QML 子目录:", rel)
        for rel in DELETE_QML_FILES:
            p = os.path.join(qml_root, *rel.split("/"))
            if os.path.isfile(p):
                saved += os.path.getsize(p)
                os.remove(p)
                print("  删 QML 工具文件:", rel)
    # qt-plugins:整目录删 + imageformats/platforms 按白名单保留
    plugins_root = os.path.join(root, "PySide6", "qt-plugins")
    if os.path.isdir(plugins_root):
        for name in DELETE_PLUGIN_DIRS:
            p = os.path.join(plugins_root, name)
            if os.path.isdir(p):
                saved += _rmtree_size(p)
                print("  删 插件目录:", name)
        for sub, keep in (("imageformats", IMAGEFORMATS_KEEP), ("platforms", PLATFORMS_KEEP)):
            sub_dir = os.path.join(plugins_root, sub)
            if not os.path.isdir(sub_dir):
                continue
            for name in os.listdir(sub_dir):
                if name.lower() not in keep:
                    p = os.path.join(sub_dir, name)
                    saved += os.path.getsize(p) if os.path.isfile(p) else _rmtree_size(p)
                    os.remove(p) if os.path.isfile(p) else shutil.rmtree(p, ignore_errors=True)
                    print("  删 插件:", sub + "/" + name)
    print("修剪完成,共释放 %.1f MB" % (saved / 1048576))
    return 0


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join("dist", "main.dist")
    sys.exit(prune(target))
