"""获取数据 - 步骤1: 打开示例 OA 主页并自动登录 (使用 Playwright 自带 Chromium)

- 使用持久化用户目录 (user_data_dir)，登录后 cookie / 会话将保留，下次直接进入。
- 若仍处在登录页，则自动填入账号密码并点击登录。
"""
import os
from playwright.sync_api import sync_playwright

OA_URL = os.getenv("OA_URL", "https://oa.example.local/seeyon/main.do")
USERNAME = os.getenv("OA_USER", "")
PASSWORD = os.getenv("OA_PASS", "")

# 简易 .env 加载（同目录），仅当对应环境变量未设置时使用
_ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.exists(_ENV_PATH):
    with open(_ENV_PATH, "r", encoding="utf-8") as _f:
        for _line in _f:
            _line = _line.strip()
            if not _line or _line.startswith("#") or "=" not in _line:
                continue
            _k, _v = _line.split("=", 1)
            _k, _v = _k.strip(), _v.strip().strip('"').strip("'")
            os.environ.setdefault(_k, _v)
    OA_URL = os.environ.get("OA_URL", OA_URL)
    USERNAME = os.environ.get("OA_USER", USERNAME)
    PASSWORD = os.environ.get("OA_PASS", PASSWORD)

# 浏览器用户数据目录，保存 cookie / 登录态
USER_DATA_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    ".browser_profile",
)


def try_login(page) -> bool:
    """若当前页面是登录页，则填表并提交。返回是否执行了登录。"""
    try:
        password_input = page.locator('input[type="password"]').first
        password_input.wait_for(state="visible", timeout=4000)
    except Exception:
        return False

    if not USERNAME or not PASSWORD:
        raise RuntimeError(
            "未配置 OA 账号密码：请在脚本同目录创建 .env，填入 OA_USER / OA_PASS，"
            "或设置同名环境变量。可参考 .env.example。"
        )

    # 用户名: 取页面里第一个 type 不是 password / hidden 的可见 input
    username_input = page.locator(
        'input:not([type="password"]):not([type="hidden"])'
    ).first
    username_input.fill(USERNAME)
    password_input.fill(PASSWORD)

    # 登录按钮: 优先匹配文字"登 录"/"登录"，再回退到 button
    for sel in [
        'text=/^\\s*登\\s*录\\s*$/',
        'button:has-text("登录")',
        'button:has-text("登 录")',
        'input[type="submit"]',
    ]:
        btn = page.locator(sel).first
        if btn.count() > 0:
            btn.click()
            break
    else:
        password_input.press("Enter")

    page.wait_for_load_state("domcontentloaded")
    return True


def open_oa_main():
    os.makedirs(USER_DATA_DIR, exist_ok=True)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            USER_DATA_DIR,
            headless=False,
            viewport={"width": 1366, "height": 800},
            args=["--remote-debugging-port=9222"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(OA_URL, wait_until="domcontentloaded")

        if try_login(page):
            print("已自动填写账号密码并提交登录。")
        else:
            print("已是登录态或未检测到登录表单。")

        print("当前页面:", page.url)
        print("关闭浏览器窗口即结束（关闭后下次运行将保留登录状态）。")
        page.wait_for_event("close", timeout=0)


if __name__ == "__main__":
    open_oa_main()
