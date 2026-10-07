@echo off
chcp 65001 >nul
rem 表答 SheetTalk - Nuitka 打包脚本(产物: dist\main.dist\ExcelAI.exe)
setlocal
cd /d "%~dp0"

python -m pip install nuitka ordered-set || goto :fail

rem pyside6 插件自动带上 QML 所需的 Qt 模块与插件;qml/addin/assets 以数据目录随包
rem 分发(运行时按 exe 旁查找,见 app/runtime.py)。首次构建需 C 编译器(MSVC,
rem 或让 Nuitka 自动下载 MinGW),耗时 10-30 分钟属正常。
python -m nuitka --standalone --enable-plugin=pyside6 ^
  --windows-console-mode=disable ^
  --windows-icon-from-ico=assets\icon.ico ^
  --include-data-dir=qml=qml ^
  --include-data-dir=addin=addin ^
  --include-data-dir=assets=assets ^
  --output-dir dist --output-filename ExcelAI.exe ^
  --company-name ExcelAI --product-name "表答 SheetTalk" ^
  --file-version 1.3.0 --product-version 1.3.0 ^
  --assume-yes-for-downloads --lto=no ^
  main.py || goto :fail

rem 修剪产物里用不到的 Qt 大模块(WebEngine/3D/Charts 等,约省 200MB)
python -m app.prune_dist dist\main.dist || goto :fail

echo.
echo ==============================================
echo 打包完成: dist\main.dist\ExcelAI.exe
echo 分发请复制整个 dist\main.dist 文件夹
echo ==============================================
pause
exit /b 0

:fail
echo.
echo 打包失败,请检查上方错误信息
pause
exit /b 1
