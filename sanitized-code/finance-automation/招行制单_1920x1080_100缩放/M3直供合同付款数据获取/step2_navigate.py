"""获取数据 - 步骤2: 进入「集团空间」→ 翻页 → 进入「报表中心」

复用 step1_open.py 的持久化用户目录与登录函数。
"""
import os
from playwright.sync_api import sync_playwright

from step1_open import OA_URL, USER_DATA_DIR, try_login


def click_group_space(page) -> None:
    """点击右上角『集团空间』标签。"""
    page.locator('text=集团空间').first.click()
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(800)


def click_menu_next_arrow(page) -> None:
    """点击菜单条右侧的下一页箭头一次。"""
    # 优先：根据 SVG / 类名常见命名
    candidates = [
        '[class*="next"]:visible',
        '[class*="right-arrow"]:visible',
        '[class*="arrow-right"]:visible',
        'i[class*="next"]:visible',
        'button[aria-label*="next" i]',
    ]
    for sel in candidates:
        loc = page.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible():
                loc.click()
                page.wait_for_timeout(600)
                return
        except Exception:
            continue
    raise RuntimeError("未找到菜单右侧『下一页』箭头按钮")


def click_report_center(page) -> None:
    """点击『报表中心』。"""
    page.locator('text=报表中心').first.click()
    page.wait_for_load_state("domcontentloaded")


def click_report_analysis(page) -> None:
    """点击下拉菜单里的『报表分析』。"""
    item = page.locator('text=报表分析').first
    item.wait_for(state="visible", timeout=5000)
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
        page.goto(OA_URL, wait_until="domcontentloaded")

        if try_login(page):
            print("已自动登录。")
            page.wait_for_timeout(1500)
        else:
            print("已是登录态。")

        click_group_space(page)
        print("已切换到『集团空间』。")

        click_menu_next_arrow(page)
        print("已点击菜单右侧箭头。")

        click_report_center(page)
        print("已点击『报表中心』。")

        click_report_analysis(page)
        print("已进入『报表分析』，当前页:", page.url)

        print("关闭浏览器窗口即结束。")
        page.wait_for_event("close", timeout=0)


if __name__ == "__main__":
    run()
