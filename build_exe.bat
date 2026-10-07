@echo off
rem 表答 SheetTalk - Nuitka 打包脚本(产物: dist\main.dist\ExcelAI.exe)
rem
rem 两个坑,改这个文件前必读:
rem 1. 必须保存为 CRLF 换行(LF 下 ^ 续行会被 cmd 拆碎);
rem 2. cmd 在 chcp 65001 下解析含中文的长行会错位,所以 nuitka 命令
rem    写成单行 ASCII,中文只留在 rem 和末尾 chcp 之后的 echo 里,
rem    元数据 product-name 也因此用英文。
setlocal
cd /d "%~dp0"

python -m pip install nuitka ordered-set || goto :fail

rem pyside6 插件自动带上 QML 所需的 Qt 模块与插件;qml/addin/assets 以数据目录随包分发(运行时按 exe 旁查找,见 app/runtime.py)。首次构建需 C 编译器(MSVC,或让 Nuitka 自动下载 MinGW),耗时 10-30 分钟属正常。
rem 注意:Nuitka 4.2+ 的参数解析只认「--选项=值」,空格分隔会报 requires an argument,所有带值参数必须写等号形式。
rem nofollow:仅 cryptography(10MB)/win32ui 无源码引用可排除。PySide6 的 QtNetwork/QtOpenGL
rem 绑定虽然源码没 import,但 shiboken 按链接关系强制要求(实测:QtQml 需要 QtNetwork,QtQuick 需要 QtOpenGL),缺了启动即 ImportError,必须随包。
rem nofollow:cryptography(10MB)/win32ui 无源码引用可排除。PySide6 的 QtNetwork/QtOpenGL
rem 绑定虽然源码没 import,但 shiboken 按链接关系强制要求,缺了启动即 ImportError,必须随包。
rem noinclude-dlls:onefile 会把整个 Qt 载荷压进 exe,必须按 prune_dist 同款黑名单
rem 在打包期排除用不到的 Qt DLL(清单依据见 app/prune_dist.py 头注释)。
python -m nuitka --mode=onefile --enable-plugin=pyside6 --windows-console-mode=disable --windows-icon-from-ico=assets/icon.ico --include-data-dir=qml=qml --include-data-dir=addin=addin --include-data-dir=assets=assets --onefile-tempdir-spec="{CACHE_DIR}/{PRODUCT}/{VERSION}" --noinclude-dlls=opengl32sw* --noinclude-dlls=qt63d* --noinclude-dlls=qt6charts* --noinclude-dlls=qt6concurrent* --noinclude-dlls=qt6datavis* --noinclude-dlls=qt6graphs* --noinclude-dlls=qt6labs* --noinclude-dlls=qt6location* --noinclude-dlls=qt6multimedia* --noinclude-dlls=qt6pdf* --noinclude-dlls=qt6positioning* --noinclude-dlls=qt6quick3d* --noinclude-dlls=qt6quickdialogs* --noinclude-dlls=qt6quickparticles* --noinclude-dlls=qt6quickshapes* --noinclude-dlls=qt6quicktest* --noinclude-dlls=qt6quicktimeline* --noinclude-dlls=qt6quickvectorimage* --noinclude-dlls=qt6remoteobjects* --noinclude-dlls=qt6scxml* --noinclude-dlls=qt6sensors* --noinclude-dlls=qt6shadertools* --noinclude-dlls=qt6spatialaudio* --noinclude-dlls=qt6sql* --noinclude-dlls=qt6statemachine* --noinclude-dlls=qt6svg* --noinclude-dlls=qt6test* --noinclude-dlls=qt6texttospeech* --noinclude-dlls=qt6virtualkeyboard* --noinclude-dlls=qt6webchannel* --noinclude-dlls=qt6webengine* --noinclude-dlls=qt6websockets* --noinclude-dlls=qt6webview* --noinclude-dlls=qt6quickeffects* --noinclude-dlls=qt6qmllocalstorage* --noinclude-dlls=qt6qmlnetwork* --noinclude-dlls=qt6qmlxmllistmodel* --noinclude-dlls=qt6quickcontrols2fluentwinui3* --noinclude-dlls=qt6quickcontrols2fusion* --noinclude-dlls=qt6quickcontrols2imagine* --noinclude-dlls=qt6quickcontrols2material* --noinclude-dlls=qt6quickcontrols2universal* --noinclude-dlls=qt6quickcontrols2windowsstyle* --noinclude-dlls=qt6quickcontrols2macos* --noinclude-dlls=qt6quickcontrols2ios* --noinclude-dlls=qt6quickcontrols2native* --noinclude-qt-plugins=styles,tls,iconengines --nofollow-import-to=cryptography --nofollow-import-to=win32ui --python-flag=no_asserts --python-flag=no_docstrings --output-dir=dist --output-filename=ExcelAI.exe --company-name=ExcelAI --product-name="SheetTalk" --file-version=1.4.0 --product-version=1.4.0 --assume-yes-for-downloads --lto=no main.py || goto :fail

rem 修剪产物里用不到的 Qt 大模块(WebEngine/3D/Charts 等,约省 200MB)
python -m app.prune_dist dist\main.dist || goto :fail

if not exist "dist\ExcelAI.exe" (
    echo 打包失败: 找不到 dist\ExcelAI.exe
    goto :fail
)

chcp 65001 >nul
echo.
echo ==============================================
echo 打包完成: dist\main.dist\ExcelAI.exe
echo 分发请复制整个 dist\main.dist 文件夹
echo ==============================================
pause
exit /b 0

:fail
chcp 65001 >nul
echo.
echo 打包失败,请检查上方错误信息
pause
exit /b 1
