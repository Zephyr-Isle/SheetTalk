@echo off
rem 表答 SheetTalk - 一键卸载(小白专用)
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo.
echo ===========================================
echo    表答 SheetTalk 一键卸载
echo ===========================================
echo.

where python >nul 2>nul
if not errorlevel 1 goto :run
where py >nul 2>nul
if not errorlevel 1 goto :runpy

echo [X] 没有检测到 Python,无法执行卸载。
echo.
pause
exit /b 1

:run
python install.py --uninstall
set "RC=%ERRORLEVEL%"
goto :done

:runpy
py -3 install.py --uninstall
set "RC=%ERRORLEVEL%"

:done
echo.
if "%RC%"=="0" (
  echo ===========================================
  echo   卸载结束,本窗口可以关闭了
  echo ===========================================
) else (
  echo [X] 卸载过程出错(错误码 %RC%),请截图上面的信息。
)
echo.
pause
exit /b %RC%
