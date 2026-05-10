"""
兴业银行企业网银 - 登录公共模块

提供登录前的通用流程（打开应用 -> 网盾密码 -> 填写登录页），
调用方在 login() 返回后自行决定何时点击 "登录" 按钮。

使用示例：
    from cib_login import login, click_login_button

    ctx = login()           # 打开应用、填完账号/密码/勾选同意，停在点登录前
    click_login_button(ctx) # 真正点击登录
"""
import os
import sys
import time
import pyautogui
import pygetwindow as gw
import win32gui


_CIB_ROOT = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(_CIB_ROOT, ".env")
SHORTCUT_PATH = os.environ.get(
    "CIB_SHORTCUT_PATH",
    os.path.join(os.path.expanduser("~"), "Desktop", "兴业银行企业网银.lnk"),
)
SCREENSHOT_DIR = os.path.join(_CIB_ROOT, "screenshots")

LOGIN_COORDS = {
    "dropdown":  (1331, 395),
    "password":  (1331, 453),
    "checkbox":  (1132, 489),
    "login_btn": (1331, 577),
}


def screenshot(name):
    if not os.path.exists(SCREENSHOT_DIR):
        os.makedirs(SCREENSHOT_DIR)
    path = os.path.join(SCREENSHOT_DIR, f"{name}.png")
    pyautogui.screenshot(path)
    print(f"  [截图] {path}")


def load_config(env_path=ENV_PATH):
    config = {}
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith("#") or not line or "=" not in line:
                continue
            key, value = line.split("=", 1)
            config[key.strip()] = value.strip()
    return config


def find_window():
    for w in gw.getAllWindows():
        if w.title and "兴业" in str(w.title):
            return w
    return None


def find_ukey_dialog():
    """查找 '验证网盾密码' 对话框（含子窗口）"""
    for w in gw.getAllWindows():
        if w.title and "验证网盾密码" in str(w.title):
            return w

    found = []

    def _enum_child(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if title and "验证网盾密码" in title and win32gui.IsWindowVisible(hwnd):
            found.append(hwnd)

    def _enum_top(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if title and "验证网盾密码" in title and win32gui.IsWindowVisible(hwnd):
            found.append(hwnd)
        win32gui.EnumChildWindows(hwnd, _enum_child, None)

    win32gui.EnumWindows(_enum_top, None)

    if not found:
        return None

    hwnd = found[0]
    l, t, r, b = win32gui.GetWindowRect(hwnd)

    class _W:
        pass
    w = _W()
    w.title = win32gui.GetWindowText(hwnd)
    w.left, w.top = l, t
    w.width, w.height = r - l, b - t
    w._hwnd = hwnd
    w.activate = lambda: win32gui.SetForegroundWindow(hwnd)
    return w


def _type_with_capslock(s, interval=0.1):
    caps_on = False
    for ch in s:
        need_caps = ch.isalpha() and ch.isupper()
        if need_caps != caps_on:
            pyautogui.press("capslock")
            time.sleep(0.1)
            caps_on = need_caps
        pyautogui.write(ch.lower() if ch.isalpha() else ch, interval=0)
        time.sleep(interval)
    if caps_on:
        pyautogui.press("capslock")
        time.sleep(0.1)


def handle_ukey_dialog(ukey_pwd, wait_seconds=10, tag=""):
    """检测并处理网盾密码对话框。若未出现返回 False。"""
    print(f"  检测网盾密码对话框（最多等待{wait_seconds}秒）...")
    ukey_win = None
    for i in range(wait_seconds):
        ukey_win = find_ukey_dialog()
        if ukey_win:
            break
        time.sleep(1)

    if not ukey_win:
        print("  未出现网盾密码对话框，跳过")
        return False

    print(f"  检测到对话框: '{ukey_win.title}' @ ({ukey_win.left}, {ukey_win.top})")
    time.sleep(0.8)
    try:
        ukey_win.activate()
    except Exception:
        pass
    time.sleep(0.3)

    pyautogui.write(ukey_pwd, interval=0.08)
    print(f"  已输入网盾密码 (长度:{len(ukey_pwd)})")
    time.sleep(0.4)
    screenshot(f"网盾_输密码{tag}")

    pyautogui.press("enter")
    time.sleep(1.5)
    screenshot(f"网盾_确认{tag}")
    return True


def open_bank_window():
    """[1] 打开网银应用，返回主窗口对象"""
    print("\n[1] 打开网银应用...")
    win = find_window()
    if not win:
        print(f"  未找到窗口，打开快捷方式: {SHORTCUT_PATH}")
        os.startfile(SHORTCUT_PATH)
        for i in range(30):
            time.sleep(1)
            win = find_window()
            if win:
                break
            print(f"  等待窗口... ({i+1}s)")
        if not win:
            print("  错误: 打开快捷方式后仍未找到窗口")
            sys.exit(1)
        time.sleep(3)

    print(f"  窗口: '{win.title}' @ ({win.left}, {win.top}) {win.width}x{win.height}")
    win.activate()
    if not win.visible:
        win.show()
    time.sleep(1)
    screenshot("01_开始前")
    return win


def handle_startup_ukey(ukey_pwd, wait_seconds=5):
    """[2] 启动阶段的网盾密码对话框（使用 capslock 安全输入法）"""
    print("\n[2] 检测网盾密码对话框...")
    time.sleep(2)

    ukey_win = None
    for i in range(wait_seconds):
        ukey_win = find_ukey_dialog()
        if ukey_win:
            break
        time.sleep(1)

    if ukey_win:
        print(f"  检测到网盾密码对话框: '{ukey_win.title}'，开始输入密码...")
        _type_with_capslock(ukey_pwd)
        print(f"  已输入网盾密码 (长度:{len(ukey_pwd)})")
        time.sleep(0.4)
        screenshot("网盾_输密码_启动")
        pyautogui.press("enter")
        time.sleep(1.5)
        screenshot("网盾_确认_启动")
        return True
    else:
        print("  未检测到网盾密码对话框，跳过")
        return False


def fill_login_form(cert_pwd, coords=None):
    """[3.1-3.4] 登录页：选账号 -> 输密码 -> 勾同意（不点登录）"""
    coords = coords or LOGIN_COORDS
    print("\n[3] 登录页流程（不点登录）...")

    print("  [3.1] 点击登录名下拉框...")
    pyautogui.click(*coords["dropdown"])
    time.sleep(1.8)
    screenshot("03_1_下拉框")

    print("  [3.2] 选择 AJVKLE...")
    opt_x = coords["dropdown"][0]
    opt_y = coords["dropdown"][1] + 49
    pyautogui.click(opt_x, opt_y)
    time.sleep(1.0)
    screenshot("03_2_选账号")

    print("  [3.3] 点击密码框并输入密码...")
    pyautogui.click(*coords["password"])
    time.sleep(0.5)
    pyautogui.write(cert_pwd, interval=0.05)
    print(f"    已输入登录密码 (长度:{len(cert_pwd)})")
    time.sleep(0.5)
    screenshot("03_3_输密码")

    print("  [3.4] 勾选同意复选框...")
    pyautogui.click(*coords["checkbox"])
    time.sleep(0.8)
    screenshot("03_4_勾同意")


def click_login_button(ctx=None, coords=None):
    """[3.5] 点击登录按钮"""
    coords = coords or (ctx or {}).get("coords") or LOGIN_COORDS
    print("  [3.5] 点击登录按钮...")
    pyautogui.click(*coords["login_btn"])
    time.sleep(2)
    screenshot("03_5_点登录")


def login(env_path=ENV_PATH):
    """完整执行 '点登录之前' 的所有步骤。

    返回 dict: {win, ukey_pwd, cert_pwd, coords}，调用方可用它 click_login_button。
    """
    config = load_config(env_path)
    ukey_pwd = config.get("LOGIN_PWD", "")
    cert_pwd = config.get("CERT_PWD", "")

    print("=" * 50)
    print("兴业银行网银 - 登录前公共流程")
    print("=" * 50)

    win = open_bank_window()
    handle_startup_ukey(ukey_pwd)
    fill_login_form(cert_pwd)

    print("\n[✓] 已停在点登录按钮之前。")
    return {
        "win": win,
        "ukey_pwd": ukey_pwd,
        "cert_pwd": cert_pwd,
        "coords": LOGIN_COORDS,
    }


if __name__ == "__main__":
    login()
