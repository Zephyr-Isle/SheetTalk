@echo off
rem 表答 SheetTalk - 一键安装(小白专用:双击即可,无需任何命令行知识)
rem 编码:UTF-8,配合 chcp 65001 显示中文
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo.
echo ===========================================
echo    表答 SheetTalk 一键安装
echo ===========================================
echo.

rem ---- 找 Python:依次尝试 python / py,都没有则提示安装 ----
where python >nul 2>nul
if not errorlevel 1 goto :run
where py >nul 2>nul
if not errorlevel 1 goto :runpy

echo [X] 没有检测到 Python。
echo.
echo     请先安装 Python 3.10 或更高版本:
echo     https://www.python.org/downloads/
echo.
echo     安装时务必勾选 "Add python.exe to PATH"
echo     装好后重新双击本文件即可。
echo.
pause
exit /b 1

:run
echo [+] 使用 Python:python
echo.
python install.py
set "RC=%ERRORLEVEL%"
goto :done

:runpy
echo [+] 使用 Python:py -3
echo.
py -3 install.py
set "RC=%ERRORLEVEL%"

:done
echo.
if "%RC%"=="0" (
  echo ===========================================
  echo   安装结束,本窗口可以关闭了
  echo ===========================================
) else (
  echo [X] 安装过程出错(错误码 %RC%),请截图上面的信息。
)
echo.
pause
exit /b %RC%
