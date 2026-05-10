# 财务付款 Agent 演示包

这是 EasyClaw 黑客松参赛作品「财务付款 Agent」的公开安全演示包。

## 作品简介

财务付款 Agent 是一个运行在 Windows 本机的财务自动化 Agent，用于把合同付款流程从人工查单、复制字段、打开网银、逐项填表，收敛为一条可监控、可追踪、可恢复的自动化流水线。

核心闭环：

1. 监控 OA 合同付款列表新增单据。
2. 打开付款详情并抽取收款户名、收方账号、开户行、金额、用途、单据编号等字段。
3. 生成统一的 `bank_form.json` 结构。
4. 根据付款单位选择对应 U 盾端口。
5. 驱动招行 U-BANK、工行、中行、兴业、农行等银行入口完成制单填表。
6. 保留人工确认环节，Agent 不直接替代资金最终确认。

## 技术栈

- Python
- Playwright
- pywinauto
- pyautogui
- Windows USB Hub 控制
- JSON 状态文件与本地浏览器 profile

## 安全说明

本仓库包含公开演示页和 `sanitized-code/finance-automation` 脱敏代码副本。公开包不包含真实业务代码中的敏感内容：

- 不包含 `.env` 或登录凭据
- 不包含银行账号、付款数据、客户名称
- 不包含真实 OA 导出的 JSON、状态文件或浏览器登录态
- 不包含银行页面截图
- 不包含二进制驱动、缓存和日志

`sanitized-code/finance-automation` 保留了 OA 数据获取、合同付款监控、银行制单入口、SKILL.md 和说明文档，真实域名、付款单位、本机路径和长账号数字均已替换为示例值。

## 演示链接

https://lrl0817.github.io/easyclaw-finance-agent-demo/

## 代码仓库

https://github.com/LRL0817/easyclaw-finance-agent-demo
