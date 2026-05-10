"""获取数据 - 步骤4: 打开合同付款列表 → 点击首行 → 在弹出新页中抽取字段

抽出的字段（用于后续招商银行单笔转账填单）：
  申请人 / 申请部门 / 申请日期 / 申请金额
  付款方式
  合同名称 / 合同编号
  收款单位名称 / 付款单位名称
  开户银行 / 银行账户
  合同金额 / 累计申请金额
  申请说明 / 类别 / 单据编号
"""
import json
import os
import re
from playwright.sync_api import sync_playwright

from step1_open import USER_DATA_DIR, OA_URL, try_login
from step3_finance_contract import REPORT_URL

FLOW_PATTERN = r"YLHT[A-Z0-9-]*-\d{4}-[A-Z0-9-]+-付-\d{4}-\d+"

WANT_LABELS = [
    "申请人", "申请部门", "申请日期", "申请金额",
    "付款方式",
    "合同名称", "合同编号",
    "收款单位名称", "付款单位名称",
    "开户银行", "银行账户",
    "合同金额", "累计申请金额",
    "申请说明", "类别",
]


# 招行制单需要把「开户银行」拆成「总行名 + 完整支行名」。
# 注意顺序：长前缀放前面，短前缀（如「邮储银行」）放后面，
# 否则「中国邮政储蓄银行xxx」会被「邮储银行」先命中。
BANK_HEADS = [
    "中国工商银行", "中国建设银行", "中国农业银行",
    "中国民生银行", "中国邮政储蓄银行", "中国银行",
    "招商银行", "交通银行", "中信银行", "浦发银行",
    "兴业银行", "光大银行", "华夏银行", "平安银行",
    "广发银行", "邮储银行", "宁波银行", "北京银行",
    "上海银行", "浙商银行", "渤海银行",
]


def normalize_purpose() -> str:
    # 用户规则：用途始终填「货款」，不读 OA 申请说明。
    return "货款"


def ensure_logged_in_and_open_list(context):
    """打开合同付款列表页，必要时登录。返回 page。"""
    page = context.pages[0] if context.pages else context.new_page()
    page.goto(REPORT_URL, wait_until="domcontentloaded")
    if "main.do" in page.url or page.locator('input[type="password"]').count() > 0:
        page.goto(OA_URL, wait_until="domcontentloaded")
        try_login(page)
        page.wait_for_timeout(1500)
        page.goto(REPORT_URL, wait_until="domcontentloaded")
    # 进入合同付款列表
    page.locator('text=财务报表').first.click()
    page.wait_for_timeout(500)
    page.locator('text=合同付款').first.click()
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(1500)
    return page


def find_frame_with_rows(page):
    """遍历所有 frame，找到第一个包含合同付款流水号文本的 frame。"""
    deadline_ms = 12000
    waited = 0
    step = 500
    while waited < deadline_ms:
        for fr in page.frames:
            try:
                loc = fr.locator(f"text=/{FLOW_PATTERN}/").first
                if loc.count() > 0:
                    return fr, loc
            except Exception:
                continue
        page.wait_for_timeout(step)
        waited += step
    raise RuntimeError("未在任何 frame 中找到合同付款列表行")


def click_first_row_open_detail(context, page):
    """点击列表第一条数据行，等待新页面打开并返回。"""
    fr, first_cell = find_frame_with_rows(page)
    first_cell.scroll_into_view_if_needed()
    # 首选：点击该单元格所在的整行 <tr>，若不存在再回落到单元格本身
    row_handle = first_cell.evaluate_handle("el => el.closest('tr') || el")
    try:
        with context.expect_page(timeout=15000) as new_info:
            try:
                row_handle.as_element().click()
            except Exception:
                first_cell.click()
        detail = new_info.value
    except Exception:
        # 兜底：双击行
        with context.expect_page(timeout=15000) as new_info:
            row_handle.as_element().dblclick()
        detail = new_info.value
    detail.wait_for_load_state("domcontentloaded")
    # 等表格渲染完：扫描所有 frame 直到出现『申请人』文字
    deadline = 25000
    waited = 0
    step = 500
    while waited < deadline:
        for fr in [detail] + list(detail.frames):
            try:
                if fr.locator('text=申请人').first.count() > 0:
                    detail.wait_for_timeout(800)
                    return detail
            except Exception:
                continue
        detail.wait_for_timeout(step)
        waited += step
    print("⚠ 等待详情页『申请人』超时，仍尝试抽取")
    return detail


def extract_fields(detail_page) -> dict:
    """在详情页里按标签抓取字段值。"""
    result: dict = {}

    # 单据编号: 标题旁，遍历所有 frame 找带 -付- 的编号
    targets_for_bill = [detail_page] + list(detail_page.frames)
    for tgt in targets_for_bill:
        try:
            bill_loc = tgt.locator(f"text=/{FLOW_PATTERN}/").first
            if bill_loc.count() > 0:
                result["单据编号"] = bill_loc.inner_text(timeout=2000).strip()
                break
        except Exception:
            continue

    # 用 JS 在 DOM 中扫描 table 单元格：找到包含目标标签文字的格，取其后续兄弟格的文本
    js = """
    (labels) => {
      const norm = (s) => (s || "").replace(/\\s+/g, "").trim();
      const out = {};
      const tds = Array.from(document.querySelectorAll('td, th, div, span'));
      for (const lab of labels) {
        for (let i = 0; i < tds.length; i++) {
          if (norm(tds[i].innerText) === lab) {
            // 优先：同 tr 下下一个 td
            const tr = tds[i].closest('tr');
            if (tr) {
              const cells = Array.from(tr.children);
              const idx = cells.indexOf(tds[i].closest('td,th'));
              if (idx >= 0 && idx + 1 < cells.length) {
                const v = norm(cells[idx + 1].innerText);
                if (v && v !== lab) { out[lab] = v; break; }
              }
            }
          }
        }
      }
      return out;
    }
    """
    # 在主页面 + 所有子 frame 中扫描
    targets = [detail_page] + [fr for fr in detail_page.frames if fr != detail_page.main_frame]
    for tgt in targets:
        try:
            scanned = tgt.evaluate(js, WANT_LABELS)
            for k, v in scanned.items():
                if v and k not in result:
                    result[k] = v
        except Exception:
            continue

    return result


def run():
    os.makedirs(USER_DATA_DIR, exist_ok=True)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            USER_DATA_DIR,
            headless=False,
            viewport={"width": 1366, "height": 800},
        )
        page = ensure_logged_in_and_open_list(context)
        print("已进入合同付款列表。")

        detail = click_first_row_open_detail(context, page)
        print("已打开首行详情页:", detail.url)

        data = extract_fields(detail)
        print("\n=== 抽取结果 ===")
        print(json.dumps(data, ensure_ascii=False, indent=2))

        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "latest.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"\n已保存到: {out_path}")

        # 拆分 OA 的开户银行（如『中国工商银行成都高新综合保税区支行』）
        # 为招行制单脚本所需的：开户银行（仅总行名）+ 支行名称（完整）
        full_bank = data.get("开户银行", "") or ""
        bank_head = full_bank
        for h in BANK_HEADS:
            if full_bank.startswith(h):
                bank_head = h
                break

        # 对照招行制单 fill_transfer_form 所需字段
        bank_form = {
            "付款单位名称": data.get("付款单位名称", ""),
            "收方账号": data.get("银行账户", ""),
            "收方户名": data.get("收款单位名称", ""),
            "开户银行": bank_head,                          # 总行名（用于下拉搜索）
            "支行名称": full_bank,                          # 完整支行名（含『...支行』）
            "金额": (data.get("申请金额", "") or "").replace(",", ""),
            "用途": normalize_purpose(),
            "业务参考号": data.get("单据编号", ""),
        }
        bank_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bank_form.json")
        with open(bank_path, "w", encoding="utf-8") as f:
            json.dump(bank_form, f, ensure_ascii=False, indent=2)
        print(f"招行表单字段已保存到: {bank_path}")
        print("\n=== 招行表单映射 ===")
        print(json.dumps(bank_form, ensure_ascii=False, indent=2))

        print("\n抽取完成，关闭浏览器。")
        try:
            context.close()
        except Exception:
            pass


if __name__ == "__main__":
    run()
