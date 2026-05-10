招行制单脚本换机包

目标环境
- 分辨率：1920 x 1080
- 缩放：100%
- 系统：Windows
- U-BANK：建议管理员运行，且先手工确认能正常登录和单笔转账经办

包内主要文件
- run_zhidan.bat：运行制单脚本
- install_requirements.bat：安装 Python 依赖
- requirements.txt：Python 依赖清单
- 招行\skills\制单单账号单笔转账skill\制单.py：制单主脚本
- 招行\ubank_common.py：登录、点击、按键等公共函数
- 招行\.env.example：登录密码模板，复制为 .env 后填写
- M3直供合同付款数据获取\bank_form.json：当前制单读取的数据文件
- M3直供合同付款数据获取\step*.py：生成/提取 bank_form.json 的相关脚本

新电脑第一次使用
1. 安装 Python，并确认命令行里 python 可用。
2. 双击 install_requirements.bat 安装依赖。
3. 复制 招行\.env.example 为 招行\.env，填写 LOGIN_PWD 和 CERT_PWD。
4. 如果 U-BANK 快捷方式不是桌面上的“招行U-BANK.lnk”，在 招行\.env 里填写 CMB_UBANK_SHORTCUT_PATH。
5. 确认 Windows 显示设置为 1920 x 1080、缩放 100%。
6. 双击 run_zhidan.bat 运行。

重要说明
- 这个包没有复制原电脑的真实 招行\.env，避免把密码散落到桌面包里。
- 制单脚本当前保留“直接键盘逐字输入”，不做 Ctrl+A/Backspace 清空，也不粘贴用途。
- 坐标兜底已按前台窗口比例计算，并在脚本启动时启用 DPI awareness，适配 1920x1080/100%。
- Firmbank.exe 本身之前出现过不稳定崩溃；如果新电脑也崩，优先重装 U-BANK/安全控件/UKey 驱动，并加白名单、管理员运行。
