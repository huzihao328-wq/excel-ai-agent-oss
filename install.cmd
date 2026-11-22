@echo off
setlocal
cd /d "%~dp0"
set "NODE=node"
if exist "%~dp0runtime\node.exe" set "NODE=%~dp0runtime\node.exe"
echo.
echo 正在安装 Excel AI Agent（不访问 npm）...
echo 升级请直接装，不要先卸载，记忆和技能会保留。
echo.
"%NODE%" scripts\install-persistent.js
if errorlevel 1 (
  echo.
  echo 安装失败。请查看上方报错，或打开 INSTALL.AGENT.md 的故障排除。
  pause
  exit /b 1
)
echo.
echo 安装完成。请完全退出 Excel 再重新打开，功能区应出现 AI Agent。
pause
