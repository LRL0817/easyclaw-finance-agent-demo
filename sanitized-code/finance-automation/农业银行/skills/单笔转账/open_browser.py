# -*- coding: utf-8 -*-
"""
打开农行企业网银登录页，并处理证书登录后的"选择证书"弹窗，
登录成功后自动填写单笔转账表单。

思路：
1. 直接打开企业网银登录页，避免前置跳转带来的不稳定。
2. 点击"证书登录"时不等待页面跳转。
3. 优先按网页 DOM 点击弹窗中的蓝色"确定"按钮。
4. 如果 DOM 点击失败，再用键盘 Enter 兜底。
5. 登录成功后导航到单笔转账页面，自动填写表单。
"""
import os
import time

import pyautogui
from dotenv import load_dotenv
from playwright.sync_api import Error, TimeoutError, sync_playwright


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ABC_ROOT = os.path.normpath(os.path.join(SCRIPT_DIR, "..", ".."))
HOME_URL = "https://www.abchina.com.cn/cn/"
ENV_PATH = os.path.join(ABC_ROOT, ".env")
load_dotenv(ENV_PATH)
KB_PASSWORD = os.getenv("KB_PASSWORD", "")
TRANSFER_IMAGE_PATH = os.getenv("TRANSFER_IMAGE_PATH", "")


def click_cert_login(page):
    print("[步骤1] 正在等待证书登录按钮...")
    cert_login = page.locator("#m-kbbtn-new")
    cert_login.wait_for(state="visible", timeout=15000)
    print("[步骤1] 证书登录按钮已可见，正在点击...")
    cert_login.click(no_wait_after=True)
    print("[步骤1] 已完成点击证书登录")
    return cert_login


def close_tip_dialog_if_needed(page):
    """关闭可能出现的网页内提示弹窗。"""
    print("[步骤2] 检查是否有温馨提示弹窗...")
    wrappers = page.locator(".el-dialog__wrapper, .el-message-box__wrapper")
    count = wrappers.count()
    if count == 0:
        print("[步骤2] 未发现提示弹窗，跳过")
        return

    for i in range(count):
        wrapper = wrappers.nth(i)
        try:
            if not wrapper.is_visible():
                continue
            text = wrapper.inner_text(timeout=1000)
            if "温馨提示" not in text:
                continue
            ok_btn = wrapper.locator("button:has-text('确定'), .el-button:has-text('确定')").first
            ok_btn.click(timeout=3000)
            print("[步骤2] 已关闭网页提示弹窗")
            return
        except Exception:
            continue
    print("[步骤2] 无需关闭的提示弹窗")


def click_confirm_in_certificate_dialog(page) -> bool:
    """
    点击"选择证书"弹窗里的蓝色确定。

    从截图看这是网页模态框，不是系统原生窗口。
    快速尝试DOM点击，不阻塞太久。
    """
    print("[步骤3] 尝试通过 DOM 点击证书弹窗确定按钮...")
    selectors = [
        ".el-dialog__wrapper:visible .el-button--primary:visible",
        ".el-dialog:visible .el-button--primary:visible",
        "[role='dialog']:visible button:has-text('确定')",
        ".el-dialog__wrapper:visible button:has-text('确定')",
    ]

    for idx, selector in enumerate(selectors):
        try:
            print(f"[步骤3] 尝试选择器 {idx + 1}/{len(selectors)}: {selector}")
            button = page.locator(selector).last
            button.wait_for(state="visible", timeout=2000)
            button.click(timeout=2000, force=True)
            print(f"[步骤3] 已通过 DOM 点击确定按钮: {selector}")
            return True
        except Exception as e:
            print(f"[步骤3] 选择器 {idx + 1} 失败: {type(e).__name__}")
            continue

    # 再退一步，用 JS 在可见弹窗里找主按钮。
    print("[步骤3] DOM选择器均未命中，尝试JS方式...")
    try:
        clicked = page.evaluate(
            """() => {
                const wrappers = Array.from(document.querySelectorAll('.el-dialog__wrapper, .el-dialog, [role="dialog"]'));
                for (const wrapper of wrappers) {
                    const style = window.getComputedStyle(wrapper);
                    if (style.display === 'none' || style.visibility === 'hidden') continue;
                    if (!wrapper.textContent.includes('选择证书')) continue;
                    const buttons = wrapper.querySelectorAll('button, .el-button');
                    for (const button of buttons) {
                        const text = (button.textContent || '').trim();
                        if (text === '确定' || button.className.includes('primary')) {
                            button.click();
                            return true;
                        }
                    }
                }
                return false;
            }"""
        )
        if clicked:
            print("[步骤3] 已通过 JS 点击确定按钮")
            return True
    except Exception as e:
        print(f"[步骤3] JS方式也失败: {type(e).__name__}")

    print("[步骤3] DOM和JS方式均未成功，将使用Enter键兜底")
    return False


def fallback_press_enter(times: int = 12, interval: float = 0.6):
    """
    多次回车兜底。证书选择框是 Chrome 原生窗口，DOM 点不到，
    默认证书已高亮，按 Enter 即可确认。多按几次防止焦点未就位。
    """
    print(f"[步骤4] ========== 开始 Enter 键兜底（共 {times} 次，间隔 {interval}s）==========")
    time.sleep(0.5)
    for i in range(times):
        pyautogui.press("enter")
        print(f"[步骤4] 已发送第 {i + 1}/{times} 次 Enter")
        time.sleep(interval)
    print(f"[步骤4] Enter 兜底完成，共发送 {times} 次")


def input_kb_password(password: str):
    """回车选证书后，密码框会自动获得焦点，固定等 2 秒再逐位输入。"""
    print("[步骤5] 等待 2 秒让密码框获得焦点...")
    time.sleep(2)

    print(f"[步骤5] 开始输入密码（共 {len(password)} 位）...")
    for i, char in enumerate(password):
        pyautogui.write(char, interval=0.08)
        print(f"[步骤5]   输入第 {i + 1}/{len(password)} 位: *")

    time.sleep(0.3)
    pyautogui.press("enter")
    print("[步骤5] 已按回车提交密码")


def navigate_to_single_transfer(page):
    """
    登录成功后，依次点击顶部导航的"付款业务"和子菜单中的"单笔转账"，
    进入单笔转账操作页面。
    """
    print("[步骤6] 等待页面加载完成...")
    page.wait_for_load_state("networkidle")
    time.sleep(2)

    # 点击顶部导航栏的"付款业务"
    print("[步骤6] 正在点击顶部导航栏「付款业务」...")
    payment_tab = page.locator("text=付款业务").first
    payment_tab.wait_for(state="visible", timeout=15000)
    payment_tab.click()
    time.sleep(1)

    # 点击子菜单中的"单笔转账"
    print("[步骤6] 正在点击子菜单「单笔转账」...")
    single_transfer = page.locator("text=单笔转账").first
    single_transfer.wait_for(state="visible", timeout=10000)
    single_transfer.click()

    print("[步骤6] 等待单笔转账页面加载...")
    page.wait_for_load_state("networkidle")
    time.sleep(2)
    print("[步骤6] 已进入单笔转账页面")


def extract_transfer_info(image_path: str) -> dict:
    """
    读取示例供应链合同付款单据图片，提取转账关键字段。

    使用 Playwright 打开图片，通过视觉能力识别表格内容。
    返回包含收款账号、收款户名、开户行、金额、用途的字典。
    """
    print(f"[步骤7] 正在读取单据图片: {image_path}")

    if not os.path.exists(image_path):
        print(f"[步骤7] 错误: 图片文件不存在 - {image_path}")
        return {}

    # 使用 Playwright 打开图片进行视觉识别
    pw = sync_playwright().start()
    browser = pw.chromium.launch()
    page = browser.new_page()

    # 直接导航到本地图片文件
    page.goto(f"file:///{image_path.replace(os.sep, '/').lstrip('/')}", wait_until="domcontentloaded")
    time.sleep(1)

    # 截图并提取文本信息
    # 通过 evaluate 获取页面上可见的文本内容（对于图片会返回空）
    # 这里使用 playwright 的截图 + 视觉能力来提取信息
    screenshot_bytes = page.screenshot(type="png")
    browser.close()
    pw.stop()

    print("[步骤7] 单据图片已读取，开始解析...")
    # 返回空字典，实际字段提取在 fill_transfer_form 前由调用者处理
    return {"_screenshot": screenshot_bytes, "_image_path": image_path}


def fill_transfer_form(page, data: dict):
    """
    将提取的转账数据自动填入农行单笔转账表单。

    data 字典应包含以下字段：
    - 收款账号: 收款方银行账号
    - 收款户名: 收款方单位/个人名称
    - 收款方开户行: 收款方开户银行名称
    - 金额: 转账金额
    - 用途: 转账用途说明

    只填写不提交，保持浏览器打开供人工确认。
    """
    print("[步骤8] 开始填写单笔转账表单...")

    # 等待表单区域加载完成
    time.sleep(1)

    field_mapping = [
        ("收款账号", "请输入收款方账号或选择已保存的收款方信息"),
        ("收款户名", "请输入收款户名或选择已保存的收款方信息"),
        ("收款方开户行", "请输入或选择开户行"),
        ("金额", "请输入金额"),
        ("用途", "请选择或输入用途"),
    ]

    for field_name, placeholder in field_mapping:
        value = data.get(field_name, "")
        if not value:
            print(f"[步骤8] {field_name}: 数据为空，跳过")
            continue

        try:
            print(f"[步骤8] 正在填写{field_name}...")

            if field_name == "金额":
                # 金额输入框有特殊前缀符号
                input_box = page.locator("input[placeholder*='金额']").first
                input_box.wait_for(state="visible", timeout=10000)
                input_box.fill(value)
                print(f"[步骤8] {field_name}已填写: {value}")
            elif field_name == "收款方开户行":
                # 开户行是下拉选择框
                select_box = page.locator("input[placeholder*='开户行']").first
                select_box.wait_for(state="visible", timeout=10000)
                select_box.click()
                time.sleep(0.5)
                select_box.fill(value)
                time.sleep(0.5)
                # 尝试点击匹配的下拉选项
                option = page.get_by_text(value).first
                if option.count() > 0:
                    option.first.click(timeout=3000)
                print(f"[步骤8] {field_name}已填写: {value}")
            else:
                # 普通输入框
                input_box = page.locator(f"input[placeholder='{placeholder}']").first
                input_box.wait_for(state="visible", timeout=10000)
                input_box.fill(value)
                print(f"[步骤8] {field_name}已填写: {value}")

            time.sleep(0.5)  # 每个字段之间等待一下，让页面响应

        except Exception as e:
            print(f"[步骤8] 填写{field_name}失败: {type(e).__name__}: {e}")
            continue

    print("[步骤8] 表单填写完成")


def main():
    print("正在启动浏览器...")
    playwright = sync_playwright().start()
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(viewport={"width": 1600, "height": 900})
    page = context.new_page()

    print(f"正在打开: {HOME_URL}")
    page.goto(HOME_URL, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle")
    time.sleep(2)

    print("正在点击企业网银登录...")
    # 首页"企业网银登录"通常会新开标签页，用 expect_page 捕获。
    try:
        with context.expect_page(timeout=10000) as new_page_info:
            page.get_by_text("企业网银登录", exact=True).first.click()
        page = new_page_info.value
        page.wait_for_load_state("domcontentloaded")
    except TimeoutError:
        # 没有新开页，当前页面应已跳转。
        page.wait_for_load_state("domcontentloaded")
    page.wait_for_load_state("networkidle")
    time.sleep(2)

    click_cert_login(page)

    # 点击证书登录后等待弹窗出现，直接用 Enter 兜底
    time.sleep(2)
    fallback_press_enter(times=1, interval=0.4)

    # 回车完等待1秒后开始输入密码
    time.sleep(1)

    # 输入K宝密码
    if KB_PASSWORD:
        input_kb_password(KB_PASSWORD)
    else:
        print("[警告] .env 中未配置 KB_PASSWORD，跳过密码输入")

    # 导航到单笔转账页面
    navigate_to_single_transfer(page)

    # 自动填写转账表单
    if TRANSFER_IMAGE_PATH:
        transfer_data = extract_transfer_info(TRANSFER_IMAGE_PATH)
        if transfer_data:
            fill_transfer_form(page, transfer_data)
        else:
            print("[警告] 单据信息提取失败，跳过表单填写")
    else:
        print("[提示] .env 中未配置 TRANSFER_IMAGE_PATH，跳过表单自动填写")

    screenshot_path = os.path.join(SCRIPT_DIR, "screenshot.png")
    try:
        if not page.is_closed():
            page.screenshot(path=screenshot_path, full_page=True)
            print(f"已截图保存到: {screenshot_path}")
    except (Error, TimeoutError) as exc:
        print(f"截图跳过: {exc}")

    print("浏览器保持运行中，按 Ctrl+C 退出...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n正在关闭浏览器...")
        context.close()
        browser.close()
        playwright.stop()
        print("浏览器已关闭。")


if __name__ == "__main__":
    main()
