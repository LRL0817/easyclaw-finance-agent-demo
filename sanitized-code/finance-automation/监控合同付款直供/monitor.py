"""监控 - 示例 OA 合同付款列表新数据轮询

完整执行 登录 → 集团空间 → 翻页 → 报表中心 → 报表分析 → 财务报表 → 合同付款，
然后定时刷新报表 iframe，比对首行流水号，发现新流水号即提示。
"""
import argparse
import builtins
import json
import logging
import os
import re
import subprocess
import sys
import time
import traceback
from datetime import datetime
from logging.handlers import RotatingFileHandler

try:
    import winsound
except ImportError:
    winsound = None

from playwright.sync_api import sync_playwright

# 复用 M3 脚本的 step1_open 模块（含 try_login / 持久化目录 / .env 加载）
M3_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "M3直供合同付款数据获取")
)
sys.path.insert(0, M3_DIR)
from step1_open import OA_URL, USER_DATA_DIR, try_login  # noqa: E402
from step4_extract import BANK_HEADS, extract_fields, normalize_purpose  # noqa: E402


REQUIRED_BANK_FORM_FIELDS = ("付款单位名称", "收方账号", "收方户名", "开户银行", "金额")


def save_extracted(data: dict, flow_id: str | None = None) -> None:
    """写出抽取结果。除了覆盖 latest.json / bank_form.json，
    若提供 flow_id 还会另存到 M3_DIR/data/{flow_id}.json 与 bank_form_{flow_id}.json。

    必填字段缺失时直接 raise — process_new_flows 会捕获，不会标记 seen，
    避免把残缺数据塞给招行制单脚本。
    """
    data_dir = os.path.join(M3_DIR, "data")
    os.makedirs(data_dir, exist_ok=True)

    latest = os.path.join(M3_DIR, "latest.json")
    with open(latest, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[{now()}] 已写出: {latest}")

    full_bank = data.get("开户银行", "") or ""
    bank_head = full_bank
    for h in BANK_HEADS:
        if full_bank.startswith(h):
            bank_head = h
            break
    bank_form = {
        "付款单位名称": data.get("付款单位名称", ""),
        "收方账号": data.get("银行账户", ""),
        "收方户名": data.get("收款单位名称", ""),
        "开户银行": bank_head,
        "支行名称": full_bank,
        "金额": (data.get("申请金额", "") or "").replace(",", ""),
        "用途": normalize_purpose(),
        "业务参考号": data.get("单据编号", ""),
    }

    missing = [k for k in REQUIRED_BANK_FORM_FIELDS if not (bank_form.get(k) or "").strip()]
    if missing:
        raise RuntimeError(
            f"bank_form 必填字段缺失 {missing}；原始抽取: {data}"
        )

    bank_path = os.path.join(M3_DIR, "bank_form.json")
    with open(bank_path, "w", encoding="utf-8") as f:
        json.dump(bank_form, f, ensure_ascii=False, indent=2)
    print(f"[{now()}] 已写出: {bank_path}")

    if flow_id:
        safe_id = flow_id.replace("/", "_")
        per_raw = os.path.join(data_dir, f"{safe_id}.json")
        per_form = os.path.join(data_dir, f"bank_form_{safe_id}.json")
        with open(per_raw, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        with open(per_form, "w", encoding="utf-8") as f:
            json.dump(bank_form, f, ensure_ascii=False, indent=2)
        print(f"[{now()}] 已写出: {per_raw}")

STATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "state.json")
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
LOG_PATH = os.path.join(LOG_DIR, "monitor.log")
FLOW_PATTERN = r"YLHT[A-Z0-9-]*-\d{4}-[A-Z0-9-]+-付-\d{4}-\d+"
FLOW_RE = re.compile(FLOW_PATTERN)
DEFAULT_BROWSER_RESTART_SECONDS = 6 * 60 * 60
DEFAULT_FAILURE_THRESHOLD = 5
SUPERVISOR_RESTART_DELAY_SECONDS = 30
ZHIDAN_TIMEOUT_SECONDS = 5 * 60
SEEN_FLOWS_KEEP = 2000  # 保留最近多少条已处理流水号
ZHIDAN_EXIT_REASONS = {
    0: "成功",
    -1: "制单脚本缺失、超时或调用异常",
    2: "USB Hub 切换失败或付款单位未配置",
    3: "招行登录失败",
    4: "未找到招行主界面窗口",
    5: "未能进入目标制单页面",
    6: "付款账号选择失败或批量导入失败",
    7: "双账号批量导入失败",
    8: "招行客户端 Firmbank.exe 崩溃",
    9: "招行经办按钮定位或点击失败",
}
_PRINT_PATCHED = False


def describe_zhidan_rc(rc: int) -> str:
    return ZHIDAN_EXIT_REASONS.get(rc, "未知制单错误")


def _serialize_seen(seen: dict) -> list:
    """seen 用 dict 保存（dict 在 Python 3.7+ 保插入序），新元素插入在尾部，
    所以 [-N:] = 最新 N 条。这样换浏览器会话恢复后顺序仍然稳定，
    截断时也不会随机丢历史。"""
    return list(seen.keys())[-SEEN_FLOWS_KEEP:]


# ---------------- 提示 / 状态 ----------------

def setup_logging() -> None:
    """把现有 print 同步写入滚动日志，尽量少侵入原脚本。"""
    global _PRINT_PATCHED
    os.makedirs(LOG_DIR, exist_ok=True)
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    logger = logging.getLogger("monitor")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    if not logger.handlers:
        handler = RotatingFileHandler(
            LOG_PATH,
            maxBytes=10 * 1024 * 1024,
            backupCount=10,
            encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logger.addHandler(handler)

    if _PRINT_PATCHED:
        return

    original_print = builtins.print

    def tee_print(*args, **kwargs):
        try:
            original_print(*args, **kwargs)
        except UnicodeEncodeError:
            safe_args = [
                str(arg).encode("utf-8", errors="replace").decode("utf-8", errors="replace")
                for arg in args
            ]
            original_print(*safe_args, **kwargs)
        file = kwargs.get("file")
        if file not in (None, sys.stdout, sys.stderr):
            return
        sep = kwargs.get("sep", " ")
        end = kwargs.get("end", "\n")
        msg = sep.join(str(arg) for arg in args)
        if end and end != "\n":
            msg += end
        logger.info(msg.rstrip("\n"))

    builtins.print = tee_print
    _PRINT_PATCHED = True
    print(f"[{now()}] 日志文件: {LOG_PATH}")

def beep():
    if winsound is not None:
        try:
            for _ in range(3):
                winsound.Beep(1200, 250)
                time.sleep(0.1)
        except Exception:
            pass


def load_state() -> dict:
    if os.path.exists(STATE_PATH):
        try:
            with open(STATE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_state(state: dict) -> None:
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def mark_state(
    state: dict,
    *,
    status: str = "running",
    error: str | None = None,
    consecutive_failures: int | None = None,
) -> None:
    timestamp = datetime.now().isoformat(timespec="seconds")
    state["last_heartbeat"] = timestamp
    state["status"] = status
    if status in {"running", "baseline_ready"} and error == "":
        state["last_success_at"] = timestamp
    if error is not None:
        state["last_error"] = error
    if consecutive_failures is not None:
        state["consecutive_failures"] = consecutive_failures
    save_state(state)


def now() -> str:
    return datetime.now().strftime("%H:%M:%S")


# ---------------- 招行制单 ----------------

ZHIDAN_ROOT = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "招行")
)
ZHIDAN_SCRIPT = os.path.join(ZHIDAN_ROOT, "skills", "制单单账号单笔转账skill", "制单.py")

def run_zhidan(skip: bool = False) -> int:
    """阻塞调用招行制单脚本。返回退出码。"""
    if skip:
        print(f"[{now()}] 测试模式：跳过招行制单")
        return 0
    if not os.path.exists(ZHIDAN_SCRIPT):
        print(f"[{now()}] ⚠ 未找到制单脚本: {ZHIDAN_SCRIPT}")
        return -1
    print(f"[{now()}] 启动招行制单: {ZHIDAN_SCRIPT}")
    try:
        proc = subprocess.run(
            [sys.executable, "-X", "utf8", ZHIDAN_SCRIPT],
            cwd=ZHIDAN_ROOT,
            timeout=ZHIDAN_TIMEOUT_SECONDS,
        )
        print(f"[{now()}] 制单退出码: {proc.returncode} ({describe_zhidan_rc(proc.returncode)})")
        return proc.returncode
    except subprocess.TimeoutExpired:
        print(f"[{now()}] 制单超时 {ZHIDAN_TIMEOUT_SECONDS}s，已终止本次制单")
        return -1
    except Exception as e:
        print(f"[{now()}] 制单调用异常: {e}")
        print(traceback.format_exc())
        return -1




# ---------------- step2 导航动作 ----------------

def click_group_space(page) -> None:
    page.locator('text=集团空间').first.click()
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(800)


def click_menu_next_arrow(page) -> None:
    for sel in [
        '[class*="next"]:visible',
        '[class*="right-arrow"]:visible',
        '[class*="arrow-right"]:visible',
        'i[class*="next"]:visible',
        'button[aria-label*="next" i]',
    ]:
        loc = page.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible():
                loc.click()
                page.wait_for_timeout(600)
                return
        except Exception:
            continue
    raise RuntimeError("未找到菜单右侧『下一页』箭头")


def click_report_center(page) -> None:
    page.locator('text=报表中心').first.click()
    page.wait_for_load_state("domcontentloaded")


def click_report_analysis(context, page):
    """点击『报表分析』。该项可能在新页签打开，返回 (报表分析所在 page)。"""
    item = page.locator('text=报表分析').first
    item.wait_for(state="visible", timeout=5000)
    try:
        with context.expect_page(timeout=4000) as new_info:
            item.click()
        new_page = new_info.value
        new_page.wait_for_load_state("domcontentloaded")
        return new_page
    except Exception:
        # 同页签打开
        page.wait_for_load_state("domcontentloaded")
        return page


# ---------------- 财务报表 → 合同付款 ----------------

def click_finance_report(page) -> None:
    page.wait_for_selector('text=报表中心', timeout=10000)
    candidates = page.locator('xpath=//*[normalize-space(text())="财务报表"]')
    clicked = False
    for i in range(candidates.count()):
        el = candidates.nth(i)
        try:
            if el.is_visible():
                el.scroll_into_view_if_needed()
                el.click()
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        page.locator('text=财务报表').first.click(force=True)
    page.wait_for_function(
        "() => /财务报表\\s*\\(/.test(document.body.innerText)", timeout=8000
    )
    page.wait_for_timeout(400)


def click_contract_payment(page) -> None:
    page.locator('text=合同付款').first.click()
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(1500)


# ---------------- iframe 找首行流水号 ----------------

def find_table_frame(page, deadline_ms: int = 12000):
    waited, step = 0, 500
    while waited < deadline_ms:
        for fr in page.frames:
            try:
                if fr.locator(f"text=/{FLOW_PATTERN}/").first.count() > 0:
                    return fr
            except Exception:
                continue
        page.wait_for_timeout(step)
        waited += step
    raise RuntimeError("未找到合同付款列表 iframe")


def read_first_flow(frame) -> str | None:
    try:
        txt = frame.evaluate("() => document.body.innerText")
    except Exception:
        return None
    m = FLOW_RE.search(txt or "")
    return m.group(0) if m else None


def read_all_flows(frame) -> list[str]:
    """按页面顺序返回所有流水号（去重保序，列表通常顶部最新）。"""
    try:
        txt = frame.evaluate("() => document.body.innerText")
    except Exception:
        return []
    seen, out = set(), []
    for m in FLOW_RE.findall(txt or ""):
        if m not in seen:
            seen.add(m)
            out.append(m)
    return out


def click_row_by_flow(context, frame, flow_id: str):
    """根据流水号文本点击对应行，等待新页面打开并返回详情 page。"""
    cell = frame.locator(f'text="{flow_id}"').first
    cell.wait_for(state="visible", timeout=8000)
    cell.scroll_into_view_if_needed()
    row_handle = cell.evaluate_handle("el => el.closest('tr') || el")
    try:
        with context.expect_page(timeout=15000) as new_info:
            try:
                row_handle.as_element().click()
            except Exception:
                cell.click()
        detail = new_info.value
    except Exception:
        with context.expect_page(timeout=15000) as new_info:
            row_handle.as_element().dblclick()
        detail = new_info.value
    detail.wait_for_load_state("domcontentloaded")
    # 等表格渲染：扫所有 frame 找『申请人』
    waited, step = 0, 500
    while waited < 25000:
        for fr in [detail] + list(detail.frames):
            try:
                if fr.locator('text=申请人').first.count() > 0:
                    detail.wait_for_timeout(800)
                    return detail
            except Exception:
                continue
        detail.wait_for_timeout(step)
        waited += step
    return detail


def refresh_frame(frame) -> None:
    url = frame.url
    if not url:
        return
    try:
        frame.goto(url, wait_until="domcontentloaded", timeout=20000)
    except Exception:
        pass


# ---------------- 主流程 ----------------

def navigate_to_list(context):
    """走完整 step1+step2+step3 路径，返回报表分析所在 page。"""
    page = context.pages[0] if context.pages else context.new_page()
    page.goto(OA_URL, wait_until="domcontentloaded")
    if try_login(page):
        print(f"[{now()}] 已自动登录。")
        page.wait_for_timeout(1500)
    else:
        print(f"[{now()}] 已是登录态。")

    click_group_space(page)
    print(f"[{now()}] 进入『集团空间』。")
    click_menu_next_arrow(page)
    print(f"[{now()}] 翻页。")
    click_report_center(page)
    print(f"[{now()}] 点击『报表中心』。")
    report_page = click_report_analysis(context, page)
    print(f"[{now()}] 进入『报表分析』: {report_page.url}")

    click_finance_report(report_page)
    print(f"[{now()}] 点击『财务报表』。")
    click_contract_payment(report_page)
    print(f"[{now()}] 进入『合同付款』。")
    return report_page


def process_new_flows(
    context,
    page,
    frame,
    new_flows: list[str],
    seen: set,
    state: dict,
    *,
    skip_zhidan: bool = False,
    mark_seen: bool = True,
) -> list[str]:
    """依次为每个新流水号点击行 → 抽字段 → 写文件 → 跑制单。每条独立落盘。
    返回成功处理的列表。"""
    done = []
    # 列表中越靠下越早出现，按出现时间顺序处理：oldest-first → 倒序处理
    for flow_id in reversed(new_flows):
        print(f"[{now()}] 处理新流水: {flow_id}")
        detail = None
        try:
            try:
                frame.locator(f'text="{flow_id}"').first.wait_for(timeout=3000)
            except Exception:
                frame = find_table_frame(page)
            detail = click_row_by_flow(context, frame, flow_id)
            print(f"[{now()}] 详情页: {detail.url}")
            data = extract_fields(detail)
            print(json.dumps(data, ensure_ascii=False, indent=2))
            save_extracted(data, flow_id=flow_id)
            done.append(flow_id)
            # 关详情页避免和制单 GUI 抢焦点
            try:
                detail.close()
                detail = None
            except Exception:
                pass
            # 每条单独跑制单；只有制单退出码为 0 才标记成功
            rc = run_zhidan(skip=skip_zhidan)
            if rc == 0 and mark_seen:
                seen[flow_id] = None  # dict 保序：新流水号插在尾部
                state["seen_flows"] = _serialize_seen(seen)
                state["updated_at"] = datetime.now().isoformat(timespec="seconds")
                save_state(state)
                print(f"[{now()}] ✅ {flow_id} 处理完成")
            elif rc == 0:
                print(f"[{now()}] ✅ {flow_id} 抽取完成；测试模式未标记 seen")
            else:
                # 制单失败：不写 seen，下次轮询会重试。把这条从 done 里移掉。
                if done and done[-1] == flow_id:
                    done.pop()
                print(f"[{now()}] ❌ {flow_id} 制单失败 (rc={rc}, {describe_zhidan_rc(rc)})，未标记，下轮重试")
        except Exception as e:
            print(f"[{now()}] 抽取 {flow_id} 失败: {e}")
            print(traceback.format_exc())
            mark_state(state, status="flow_failed", error=f"{flow_id}: {e}")
        finally:
            if detail is not None:
                try:
                    detail.close()
                except Exception:
                    pass
    print(f"[{now()}] 🔄 本批 {len(done)}/{len(new_flows)} 条处理完成，恢复监控")
    return done


def recover_list(context, state: dict, reason: str):
    """重新走 OA 导航，常用于登录态过期、iframe 丢失、页面空白后的恢复。"""
    print(f"[{now()}] 尝试恢复列表页: {reason}")
    mark_state(state, status="recovering", error=reason)
    page = navigate_to_list(context)
    frame = find_table_frame(page)
    print(f"[{now()}] 列表页恢复完成")
    return page, frame


def run_session(
    interval: int,
    once: bool,
    restart_seconds: int | None,
    failure_threshold: int,
    skip_zhidan: bool,
) -> None:
    """运行一个浏览器会话。返回后由外层守护决定是否重开浏览器。"""
    state = load_state()
    # seen 用 dict 保插入序：磁盘上的 list 顺序代表「老 → 新」，
    # 直接 dict.fromkeys 后续可以 [-N:] 截掉最老的。
    seen = dict.fromkeys(state.get("seen_flows", []) or [])
    if "last_flow" in state and state["last_flow"]:
        seen.setdefault(state["last_flow"], None)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            USER_DATA_DIR,
            headless=False,
            viewport={"width": 1366, "height": 800},
        )
        session_started = time.monotonic()
        consecutive_failures = 0
        try:
            page = navigate_to_list(context)
            frame = find_table_frame(page)

            flows = read_all_flows(frame)
            if not flows:
                raise RuntimeError("启动后未读取到流水号，可能仍在登录页或报表未加载")
            print(f"[{now()}] 当前列表共 {len(flows)} 条流水号，首条: {flows[0] if flows else None}")
            mark_state(state, status="running", error="", consecutive_failures=0)

            if not seen:
                # 首次运行：把当前所有流水号当作基线，避免回溯历史。
                # flows 是 OA 列表顶部最新 → 反转一次让最新插入到 dict 尾部，
                # 与增量分支语义一致（[-N:] 一律是「最新 N 条」）。
                seen.update(dict.fromkeys(reversed(flows)))
                state["seen_flows"] = _serialize_seen(seen)
                state["updated_at"] = datetime.now().isoformat(timespec="seconds")
                mark_state(state, status="baseline_ready", error="", consecutive_failures=0)
                print(f"[{now()}] 首次运行，建立基线 {len(seen)} 条")
            else:
                new_flows = [f for f in flows if f not in seen]
                if new_flows:
                    print(f"[{now()}] 🔔 启动时发现 {len(new_flows)} 条新数据: {new_flows}")
                    beep()
                    process_new_flows(
                        context,
                        page,
                        frame,
                        new_flows,
                        seen,
                        state,
                        skip_zhidan=skip_zhidan,
                        mark_seen=not skip_zhidan,
                    )
                    try:
                        frame = find_table_frame(page)
                    except Exception as e:
                        page, frame = recover_list(context, state, f"启动处理后找不到 iframe: {e}")

            if once:
                return

            print(f"[{now()}] 进入轮询，间隔 {interval}s。Ctrl+C 退出。")
            while True:
                if restart_seconds and time.monotonic() - session_started >= restart_seconds:
                    print(f"[{now()}] 已运行 {restart_seconds}s，按计划重启浏览器会话")
                    mark_state(state, status="scheduled_restart", error="", consecutive_failures=0)
                    return

                time.sleep(interval)
                try:
                    refresh_frame(frame)
                    page.wait_for_timeout(800)
                    try:
                        frame.locator(f"text=/{FLOW_PATTERN}/").first.wait_for(
                            timeout=8000
                        )
                    except Exception:
                        frame = find_table_frame(page)
                    flows = read_all_flows(frame)
                except Exception as e:
                    consecutive_failures += 1
                    print(f"[{now()}] 刷新失败 ({consecutive_failures}/{failure_threshold}): {e}")
                    print(traceback.format_exc())
                    mark_state(
                        state,
                        status="poll_failed",
                        error=str(e),
                        consecutive_failures=consecutive_failures,
                    )
                    if consecutive_failures >= failure_threshold:
                        try:
                            page, frame = recover_list(
                                context,
                                state,
                                f"连续刷新失败 {consecutive_failures} 次",
                            )
                            consecutive_failures = 0
                            mark_state(state, status="running", error="", consecutive_failures=0)
                        except Exception as recover_error:
                            print(f"[{now()}] 浏览器会话恢复失败: {recover_error}")
                            print(traceback.format_exc())
                            mark_state(
                                state,
                                status="session_failed",
                                error=str(recover_error),
                                consecutive_failures=consecutive_failures,
                            )
                            raise
                    continue

                if not flows:
                    consecutive_failures += 1
                    print(f"[{now()}] 列表为空 ({consecutive_failures}/{failure_threshold})")
                    mark_state(
                        state,
                        status="empty_list",
                        error="列表为空",
                        consecutive_failures=consecutive_failures,
                    )
                    if consecutive_failures >= failure_threshold:
                        try:
                            page, frame = recover_list(
                                context,
                                state,
                                f"连续读取空列表 {consecutive_failures} 次",
                            )
                            consecutive_failures = 0
                            mark_state(state, status="running", error="", consecutive_failures=0)
                        except Exception as recover_error:
                            print(f"[{now()}] 空列表恢复失败: {recover_error}")
                            print(traceback.format_exc())
                            mark_state(
                                state,
                                status="session_failed",
                                error=str(recover_error),
                                consecutive_failures=consecutive_failures,
                            )
                            raise
                    continue

                consecutive_failures = 0
                mark_state(state, status="running", error="", consecutive_failures=0)
                new_flows = [f for f in flows if f not in seen]
                if new_flows:
                    print(f"[{now()}] 🔔 发现 {len(new_flows)} 条新数据: {new_flows}")
                    beep()
                    process_new_flows(
                        context,
                        page,
                        frame,
                        new_flows,
                        seen,
                        state,
                        skip_zhidan=skip_zhidan,
                        mark_seen=not skip_zhidan,
                    )
                    # 处理完后回到列表页 frame；浏览器若挂了则重新导航
                    try:
                        frame = find_table_frame(page)
                    except Exception as e:
                        try:
                            page, frame = recover_list(context, state, f"处理后找不到 iframe: {e}")
                        except Exception as recover_error:
                            print(f"[{now()}] 浏览器恢复失败: {recover_error}")
                            print(traceback.format_exc())
                            mark_state(state, status="session_failed", error=str(recover_error))
                            raise
                    print(f"[{now()}] ▶ 监控已恢复，继续轮询")
                else:
                    print(f"[{now()}] 无变化 (顶: {flows[0]})")
        finally:
            try:
                context.close()
            except Exception:
                pass


def run(
    interval: int,
    once: bool,
    restart_seconds: int | None = DEFAULT_BROWSER_RESTART_SECONDS,
    failure_threshold: int = DEFAULT_FAILURE_THRESHOLD,
    skip_zhidan: bool = False,
) -> None:
    if once:
        run_session(
            interval,
            once=True,
            restart_seconds=restart_seconds,
            failure_threshold=failure_threshold,
            skip_zhidan=skip_zhidan,
        )
        return

    while True:
        try:
            run_session(
                interval,
                once=False,
                restart_seconds=restart_seconds,
                failure_threshold=failure_threshold,
                skip_zhidan=skip_zhidan,
            )
        except KeyboardInterrupt:
            print(f"\n[{now()}] 收到 Ctrl+C，退出。")
            break
        except Exception as e:
            print(f"[{now()}] 监控会话异常退出: {e}")
            print(traceback.format_exc())

        print(f"[{now()}] {SUPERVISOR_RESTART_DELAY_SECONDS}s 后重启监控会话")
        time.sleep(SUPERVISOR_RESTART_DELAY_SECONDS)


def main():
    ap = argparse.ArgumentParser(description="监控示例 OA 合同付款列表新数据")
    ap.add_argument("--interval", type=int, default=30, help="轮询秒数 (默认 30)")
    ap.add_argument("--once", action="store_true", help="只检查一次后退出")
    ap.add_argument("--restart-hours", type=float, default=6.0, help="浏览器会话定时重启小时数，<=0 表示不重启")
    ap.add_argument("--failure-threshold", type=int, default=DEFAULT_FAILURE_THRESHOLD, help="连续失败多少次后重登/恢复")
    ap.add_argument("--skip-zhidan", action="store_true", help="测试用：抽取数据后跳过招行制单，且不标记 seen")
    args = ap.parse_args()
    setup_logging()
    restart_seconds = None if args.restart_hours <= 0 else int(args.restart_hours * 60 * 60)
    run(
        args.interval,
        args.once,
        restart_seconds=restart_seconds,
        failure_threshold=args.failure_threshold,
        skip_zhidan=args.skip_zhidan,
    )


if __name__ == "__main__":
    main()
