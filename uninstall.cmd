@echo off
setlocal
cd /d "%~dp0"
set "NODE=node"
if exist "%~dp0runtime\node.exe" set "NODE=%~dp0runtime\node.exe"
echo.
echo 正在卸载 Excel AI Agent...
echo.
"%NODE%" scripts\install-persistent.js remove
if errorlevel 1 (
  echo.
  echo 卸载过程出错，请查看上方报错。
  pause
  exit /b 1
)
echo.
echo 已卸载。请完全退出 Excel。
pause
