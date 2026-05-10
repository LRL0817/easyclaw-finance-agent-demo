"""获取数据 - 步骤3: 进入报表分析页 → 财务报表 → 合同付款

报表分析在新窗口打开。这里直接用已知的报表中心 URL 打开，避免依赖 step2 的导航链。
持久化用户目录里的 cookie 已携带登录态。
"""
import os
from playwright.sync_api import sync_playwright

from step1_open import USER_DATA_DIR, OA_URL, try_login

REPORT_URL = (
    "https://oa.example.local/seeyon/vreport/vReport.do"
    "?method=vReportView&portalId=000000000000000000&_resourceCode=F08_report_view"
)


def click_finance_report(page) -> None:
    """点击左侧『财务报表』。"""
    page.wait_for_selector('text=报表中心', timeout=10000)
    candidates = page.locator('xpath=//*[normalize-space(text())="财务报表"]')
    n = candidates.count()
    clicked = False
    for i in range(n):
        el = candidates.nth(i)
        try:
            if not el.is_visible():
                continue
            el.scroll_into_view_if_needed()
            el.click()
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        page.locator('text=财务报表').first.click(force=True)
    page.wait_for_function(
        "() => /财务报表\\s*\\(/.test(document.body.innerText)",
        timeout=8000,
    )
    page.wait_for_timeout(500)


def click_contract_payment(page) -> None:
    """点击『合同付款』报表卡片。"""
    item = page.locator('text=合同付款').first
    item.wait_for(state="visible", timeout=8000)
    item.click()
    page.wait_for_load_state("domcontentloaded")


def run():
    os.makedirs(USER_DATA_DIR, exist_ok=True)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            USER_DATA_DIR,
            headless=False,
            viewport={"width": 1366, "height": 800},
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(REPORT_URL, wait_until="domcontentloaded")

        # 若被重定向到登录页则登录后重新进入
        if "main.do" in page.url or page.locator('input[type="password"]').count() > 0:
            page.goto(OA_URL, wait_until="domcontentloaded")
            try_login(page)
            page.wait_for_timeout(1500)
            page.goto(REPORT_URL, wait_until="domcontentloaded")

        print("已打开报表分析:", page.url)

        click_finance_report(page)
        print("已点击『财务报表』。")

        click_contract_payment(page)
        print("已点击『合同付款』，当前页:", page.url)

        print("关闭浏览器窗口即结束。")
        page.wait_for_event("close", timeout=0)


if __name__ == "__main__":
    run()
