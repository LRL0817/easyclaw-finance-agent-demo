---
name: 监控
description: 监控示例 OA 合同付款列表是否出现新数据。该列表必须刷新才能拿到最新内容，本 skill 通过定时刷新内嵌报表 iframe 并比对首行流水号实现自动监测。当用户要求监控合同付款是否有新数据时使用。
---

# 监控

定时刷新示例 OA 报表分析 → 财务报表 → 合同付款 列表，比对首行流水号
（`YLHTSP-YYYY-XXXX-付-YYYY-XXXXX`），发现新数据即提示。

## 前置条件

- 已经跑过一次 `M3直供合同付款数据获取\step1_open.py`，
  `.browser_profile/` 中保留了登录态。本 skill 直接复用该目录，免登录。
- `pip install playwright && playwright install chromium`。

## 文件清单

| 文件 | 说明 |
|---|---|
| `monitor.py` | 监控脚本，自带 CLI 参数 `--interval`（轮询秒数，默认 30）、`--once`（只检查一次） |
| `state.json` | 上次看到的首行流水号（用于跨会话保留基线） |

## 工作原理

1. 复用 `M3直供合同付款数据获取/.browser_profile/` 启动持久化 Chromium。
2. 直接 goto 报表分析 URL → 点击『财务报表』→ 点击『合同付款』，定位包含
   `YLHTSP-...-付-...` 的内嵌 iframe。
3. 抓取该 iframe 内首行的『流水号』文本作为基线。
4. 每 `interval` 秒刷新一次（`frame.goto(frame.url)`，比 `page.reload` 快、不重建外层路径），重新读取首行流水号。
5. 如果新流水号 != 上次记录，提示『发现新数据』（终端打印 + `winsound.Beep` 蜂鸣 + 写入 `state.json`）。

## 用法

```bash
python "监控合同付款直供\monitor.py"               # 每 30s 轮询
python "监控合同付款直供\monitor.py" --interval 60 # 自定义间隔
python "监控合同付款直供\monitor.py" --once        # 只查一次
```

按 Ctrl+C 或关闭浏览器窗口即结束。
