# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目目标

监控示例 OA（oa.example.local）合同付款列表新增数据，自动抽取字段后驱动各家网银（招行 U-BANK / 工行 / 中行 / 兴业 / 农行）完成单笔转账填单。Windows-only，全部脚本以 Python 实现，靠 Playwright（网页银行 + OA）+ pywinauto / pyautogui（桌面客户端）混合驱动。

## 顶层目录与角色

| 目录 | 角色 |
|---|---|
| `M3直供合同付款数据获取/` | OA 抓取流水线（Playwright），step1_open → step2_navigate → step3_finance_contract → step4_extract，输出 `latest.json` + `bank_form.json` + `data/{flow_id}.json` |
| `监控合同付款直供/` | 长跑监控 `monitor.py`，定时刷新报表 iframe → 比对 seen_flows → 抽数据 → 调招行制单 |
| `招行/` | U-BANK 桌面客户端制单（pywinauto），共享 `ubank_common.py`；多个 skill 子目录对应单/双账号 × 单笔/批量 × 制单/日记账 |
| `工商银行/`, `中国银行/`, `农业银行/` | 各家企业网银 Playwright 自动化（headful），桌面 `.lnk` 启动或直接 goto |
| `兴业银行/` | 桌面客户端 pyautogui 自动化，依赖屏幕坐标（1920×1080） |

每个银行目录下有自己的 `.env`（凭据）、`screenshots/`（调试落盘）、`SKILL.md`（流程说明）。Playwright 银行的登录态持久化在 `.browser_profile/`。

## 端到端闭环

```
监控合同付款直供/monitor.py  (每 30s 刷新合同付款列表 iframe)
  ↓ 比对 state.json::seen_flows，按出现顺序逐条处理新流水号（oldest-first）
  ↓ 蜂鸣 + 在原列表点击新行 → 弹出『查询穿透显示』
step4_extract.extract_fields(detail) → save_extracted()
  ↓ 写 latest.json + bank_form.json + data/{flow_id}.json + bank_form_{flow_id}.json
  ↓ 必填字段缺失直接 raise，不会送到制单
subprocess.run(招行/skills/制单单账号单笔转账skill/制单.py)  (阻塞，超时 5 分钟)
  ↓ rc==0 才把 flow_id 标记为 seen；非 0 不标记 → 下轮重试
返回监控继续轮询
```

监控脚本带守护：`run_session` 抛异常后由 `run` 等 30s 重启；浏览器会话默认每 6h 主动重启（`--restart-hours 0` 关闭）；连续 5 次刷新失败或读到空列表会触发 `recover_list` 重新登录导航。

## bank_form.json 字段契约

`监控/save_extracted` 与 `step4_extract` 都按这个 schema 输出，所有银行制单脚本都从这里读取（招行制单的实际路径是 `M3直供合同付款数据获取/bank_form.json`）：

```
付款单位名称   ← OA 付款单位名称        # 用于决定切哪个 U 盾（见下）
收方账号       ← OA 银行账户
收方户名       ← OA 收款单位名称
开户银行       ← OA 开户银行（按 BANK_HEADS 前缀拆出总行名）
支行名称       ← OA 开户银行（保留完整支行名）
金额           ← OA 申请金额（去千分位）
用途           ← 固定 "货款"（normalize_purpose；不再使用 OA 申请说明）
业务参考号     ← OA 单据编号（招行『其他信息』里，红星规则下不填）
```

`BANK_HEADS` 在 `step4_extract.py` 维护，长前缀必须排在前面（避免「中国邮政储蓄银行」被「邮储银行」抢匹配）。

## 招行制单退出码（监控用来判断是否标记 seen）

定义在 `监控合同付款直供/monitor.py::ZHIDAN_EXIT_REASONS`：

```
0  成功
-1 制单脚本缺失 / 超时 / 调用异常
2  USB Hub 切换失败 或 付款单位未配置
3  招行登录失败
4  未找到招行主界面窗口
5  未能进入目标制单页面
6  付款账号选择失败 或 批量导入失败
7  双账号批量导入失败
8  招行客户端崩溃（Firmbank.exe 应用程序错误弹窗被检测到）
9  招行经办按钮定位或点击失败
```

新增 / 修改制单脚本时务必同步这张表。

## 多公司 U 盾切换

招行制单脚本通过 USB Hub 在物理上隔离多枚 U 盾。匹配关系硬编码在 `招行/skills/制单单账号单笔转账skill/制单.py::_USB_HUB_PORT_BY_PAYER`：

```
"示例付款单位A" → 端口 10
"示例付款单位B" → 端口 11
```

控制脚本路径：`$CODEX_HOME/skills/usb-hub/scripts/hub_ctrl.py`（默认 `~/.codex/skills/usb-hub/scripts/hub_ctrl.py`），不在本仓库内。`bank_form.json::付款单位名称` 缺失或不在表里 → 制单 rc=2 → 监控不标记 seen 但也不会重试到死，需人工查清新付款单位归属哪个端口再补到表里。

## 常用命令

工作目录：`%USERPROFILE%\Desktop\财务\`。所有命令都用 PowerShell + 中文路径，路径带空格的 skill 子目录要加双引号。

```powershell
# 启动监控（默认 30s 一轮）
python 监控合同付款直供\monitor.py

# 自定义间隔
python 监控合同付款直供\monitor.py --interval 60

# 只跑一轮
python 监控合同付款直供\monitor.py --once

# 抽取后跳过招行制单（开发调试，且不标记 seen）
python 监控合同付款直供\monitor.py --skip-zhidan

# 重置基线（首次运行会把当前列表整批当 seen，避免回放历史）
Remove-Item 监控合同付款直供\state.json

# 单步抓数据（不经过监控）
python M3直供合同付款数据获取\step1_open.py        # 仅登录
python M3直供合同付款数据获取\step4_extract.py     # 抓首条 → 写 bank_form.json

# 招行制单（读 bank_form.json）
python -X utf8 "招行\skills\制单单账号单笔转账skill\制单.py"

# 其他银行入口
python 工商银行\open_icbc.py
python 中国银行\open_boc.py
python "兴业银行\skills\单笔转账\open_bank.py"
python "农业银行\skills\单笔转账\open_browser.py"
```

无构建 / 测试 / lint。所有脚本单文件直接 `python <file>` 即可，依赖装在系统 Python（`pip install playwright pywinauto pyautogui pygetwindow pywin32 python-dotenv pillow` + `playwright install chromium`）。

## 跨脚本约定

- **不配 PYTHONPATH**：跨目录 import 一律用 `sys.path.insert(0, ...)` 自处理。监控通过这个机制 import M3 目录下的 `step1_open` 与 `step4_extract`；招行 skills 子目录的 `制单.py` 也用同样手法把 `招行/` 加进 path 来 `import ubank_common`。改动相对路径前先看脚本顶部的 `_ZHIDAN_ROOT` / `M3_DIR` 等常量。
- **持久化目录互斥**：`.browser_profile/` 同一时间只能被一个 Chromium 进程独占。运行 step2/step3/step4 前必须先关掉 step1 的窗口，否则报 `TargetClosedError`。监控因为单进程串行所以无此问题。
- **必填校验在写文件前**：`save_extracted` 在 schema 缺失关键字段时 raise，让监控捕获后跳过本条但不标 seen，避免把残缺数据塞给制单。
- **seen_flows 是有序 dict**：`monitor.py` 用 dict 保插入序，磁盘 list 是「老 → 新」，截断保留最近 `SEEN_FLOWS_KEEP=2000` 条。新增逻辑要保证插入位置在尾部。
- **流水号正则**：`YLHT[A-Z0-9-]*-\d{4}-[A-Z0-9-]+-付-\d{4}-\d+`，集中在 `monitor.py::FLOW_PATTERN` 与 `step4_extract.py::FLOW_PATTERN`，改一边记得改另一边。
- **凭据**：每个银行目录下的 `.env` 都不进版本控制，OA 凭据在 `M3直供合同付款数据获取\.env`（`OA_USER` / `OA_PASS`）。

## 已知缺口

- 招行制单只填红星必填字段（用户明确规则），「业务参考号」`bank_form.json` 里有但脚本不会展开「其他信息」折叠区填它。
- 多公司多 U 盾仅覆盖示例付款单位A / 示例付款单位B两家；新公司需先补 `_USB_HUB_PORT_BY_PAYER` 与对应物理端口接线。
- 招行制单当前以 `taskkill /F /T /IM Firmbank.exe` 关 U-BANK 收尾，不走 Alt+F4 / 点 X / 页面导航；成功路径会先点击「经办」按钮，若定位/点击失败则 rc=9 且不会标记 seen。
- Firmbank.exe 12.0.0.6 在「单笔转账经办」页面存在偶发崩溃；制单脚本启动 WinEventHook 守护线程（窗口创建/显示/标题变化）并保留 20ms 轮询兜底，发现应用程序错误弹窗后立刻强杀 Firmbank 并以 rc=8 退出。Python 层仍不能承诺绝对不可见；若 Windows 已先绘制一帧，只有 DLL / 进程内拦截才更接近硬保证。
