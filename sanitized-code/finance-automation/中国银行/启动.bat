@echo off
REM 静默启动中国银行单步支付脚本：仅弹出 Chromium 浏览器，不显示控制台窗口
cd /d "%~dp0"
start "" pythonw open_boc.py
