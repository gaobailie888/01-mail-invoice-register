@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动“邮箱发票自动登记”本地向导...
python start.py
if errorlevel 1 (
  echo.
  echo 启动失败。请确认已安装 Python 3.10 或更高版本。
  pause
)
