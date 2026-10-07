@echo off
rem 注册「表答 SheetTalk」侧边栏:Office(Excel)+ WPS 表格,注册后需重启对应软件
rem 用法: sideload.bat [all|office|wps|remove|remove office|remove wps]
python "%~dp0sideload.py" %*
pause
