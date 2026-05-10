@echo off
REM 静默启动工商银行单笔支付脚本：仅弹出 Chromium 浏览器，不显示控制台窗口
cd /d "%~dp0"
start "" pythonw open_icbc.py
