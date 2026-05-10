"""
中国银行 - 打开企业网银登录页

使用 Playwright 启动有头 Chromium 浏览器，打开中国银行官网，
点击「企业网上银行登录」，跳转到企业网银登录页后脚本退出。
"""

import ctypes
import json
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

BOC_URL = "https://netc1.igtb.boc.cn/#/login-page?redirect=%2Findex"
BOC_EXTENSION_ID = "nhhdpdhiemjpkaikglglhabjafffdjfo"
BOC_EXTENSION_NAME = "BOC Certificate Application Extension"

ENV_FILE = Path(__file__).resolve().parent / ".env"
LOG_FILE = Path(__file__).resolve().parent / "boc.log"
PROXY_ENV_KEYS = (
    "HTTP_PROXY",
    "HTTPS_PROXY",
    "ALL_PROXY",
    "NO_PROXY",
    "http_proxy",
    "https_proxy",
    "all_proxy",
    "no_proxy",
)



# Native Chrome certificate chooser is outside the page DOM, so Playwright keyboard
# events may not reach it. Use the same Enter strategy as the ICBC script:
# Interception HID keyboard first, SendInput scancode fallback.
_ITC_KB_SLOT = -1
try:
    import interception as _itc  # type: ignore
    from interception.inputs import _g_context as _itc_ctx  # type: ignore

    for _i in range(10):
        try:
            _hwid = _itc_ctx.devices[_i].get_HWID()
        except Exception:
            _hwid = None
        if _hwid and _itc_ctx.is_keyboard(_i) and "HID\\VID_" in _hwid:
            _ITC_KB_SLOT = _i
            break
    if _ITC_KB_SLOT >= 0:
        _itc.set_devices(keyboard=_ITC_KB_SLOT)
        _ITC_OK = True
    else:
        _ITC_OK = False
except Exception:
    _itc = None
    _ITC_OK = False

_USER32 = ctypes.WinDLL("user32", use_last_error=True)
_VK_RETURN = 0x0D
_INPUT_KEYBOARD = 1
_KEYEVENTF_KEYUP = 0x0002
_KEYEVENTF_SCANCODE = 0x0008
_KEYEVENTF_EXTENDEDKEY = 0x0001
_MAPVK_VK_TO_VSC = 0


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class _HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", ctypes.c_ulong),
        ("wParamL", ctypes.c_ushort),
        ("wParamH", ctypes.c_ushort),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [
        ("ki", _KEYBDINPUT),
        ("mi", _MOUSEINPUT),
        ("hi", _HARDWAREINPUT),
    ]


class _INPUT(ctypes.Structure):
    _fields_ = [
        ("type", ctypes.c_ulong),
        ("u", _INPUT_UNION),
    ]


def _send_scancode(vk: int) -> None:
    scan = _USER32.MapVirtualKeyW(vk, _MAPVK_VK_TO_VSC)
    extended = vk in (0x0D,) and False
    down = _INPUT(type=_INPUT_KEYBOARD)
    down.u.ki = _KEYBDINPUT(
        wVk=0,
        wScan=scan,
        dwFlags=_KEYEVENTF_SCANCODE | (_KEYEVENTF_EXTENDEDKEY if extended else 0),
        time=0,
        dwExtraInfo=None,
    )
    up = _INPUT(type=_INPUT_KEYBOARD)
    up.u.ki = _KEYBDINPUT(
        wVk=0,
        wScan=scan,
        dwFlags=_KEYEVENTF_SCANCODE | _KEYEVENTF_KEYUP
        | (_KEYEVENTF_EXTENDEDKEY if extended else 0),
        time=0,
        dwExtraInfo=None,
    )
    _USER32.SendInput(1, ctypes.byref(down), ctypes.sizeof(_INPUT))
    time.sleep(0.05)
    _USER32.SendInput(1, ctypes.byref(up), ctypes.sizeof(_INPUT))
    time.sleep(0.08)


def _send_enter() -> None:
    if _ITC_OK:
        _itc.press("enter")
        time.sleep(0.08)
        return
    _send_scancode(_VK_RETURN)


def _send_ascii(text: str) -> None:
    """Send ASCII letters/digits with the same keyboard backend as the ICBC script."""
    if _ITC_OK:
        _itc.write(text.lower(), interval=0.05)
        return
    for ch in text:
        if ch.isalpha():
            vk = ord(ch.upper())
        elif ch.isdigit():
            vk = ord(ch)
        else:
            continue
        _send_scancode(vk)
        time.sleep(0.05)


def _disable_proxy_for_playwright() -> None:
    for key in PROXY_ENV_KEYS:
        os.environ.pop(key, None)


def _load_env_file(path: Path) -> None:
    if not path.is_file():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key = key.strip()
        val = val.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = val


_load_env_file(ENV_FILE)
BOC_USHIELD_PASSWORD = os.environ.get("BOC_USHIELD_PASSWORD", "").strip()


def _setup_stdout_utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _setup_file_logging() -> None:
    try:
        log = LOG_FILE.open("a", encoding="utf-8", buffering=1)
        sys.stdout = log
        sys.stderr = log
        print("\n========== BOC open ==========", flush=True)
    except Exception:
        pass


def _wait_full_load(page, timeout_ms: int = 60000) -> None:
    try:
        page.wait_for_load_state("domcontentloaded", timeout=timeout_ms)
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=timeout_ms)
    except Exception:
        pass


def _click_certificate_login(page, timeout_ms: int = 15000) -> bool:
    """Click the BOC login tab labeled "数字证书登录"."""
    for frame in page.frames:
        try:
            tab = frame.get_by_text("数字证书登录", exact=True).first
            tab.wait_for(state="visible", timeout=timeout_ms)
            tab.click(no_wait_after=True, timeout=timeout_ms)
            return True
        except Exception:
            pass

        try:
            if frame.evaluate(
                r"""
            () => {
                const normalize = (text) => (text || '').replace(/\s+/g, '').trim();
                const visible = (el) => {
                    const style = window.getComputedStyle(el);
                    const box = el.getBoundingClientRect();
                    return style.visibility !== 'hidden'
                        && style.display !== 'none'
                        && box.width > 0
                        && box.height > 0;
                };
                const candidates = Array.from(document.querySelectorAll('button,a,div,span,li'))
                    .filter((el) => visible(el) && normalize(el.innerText || el.textContent) === '数字证书登录')
                    .sort((a, b) => b.getBoundingClientRect().width - a.getBoundingClientRect().width);
                if (!candidates.length) return false;
                candidates[0].click();
                return true;
            }
            """
            ):
                return True
        except Exception:
            pass
    return False


def _click_password_input(page, timeout_ms: int = 8000) -> bool:
    """Find the 用户密码 input and click it to focus."""
    selectors = (
        'input[type="password"]',
        'input[placeholder*="密码"]',
        'input[placeholder*="登录密码"]',
    )
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=timeout_ms)
            loc.scroll_into_view_if_needed(timeout=2000)
            loc.click(timeout=timeout_ms)
            print(f"[已点击] 用户密码输入框 ({sel})", flush=True)
            return True
        except Exception:
            continue

    # Fallback: click just under the "用户密码" label
    try:
        label = page.get_by_text("用户密码", exact=True).first
        box = label.bounding_box()
        if box:
            cx = box["x"] + box["width"] / 2
            cy = box["y"] + box["height"] + 28
            page.mouse.click(cx, cy)
            print(f"[已点击] 用户密码框 (标签下方 {cx:.0f},{cy:.0f})", flush=True)
            return True
    except Exception:
        pass
    return False


def _click_login_button(page, timeout_ms: int = 6000) -> bool:
    """Click the red 登录 button on the password-login panel."""
    # role-based exact match avoids matching the "密码登录" tab
    try:
        page.get_by_role("button", name="登录", exact=True).first.click(timeout=timeout_ms)
        print("[已点击] 登录按钮 (role=button)", flush=True)
        return True
    except Exception:
        pass

    selectors = (
        'button:text-is("登录")',
        '[class*="login-btn"]',
        '[class*="loginBtn"]',
        'div[role="button"]:has-text("登录")',
    )
    for sel in selectors:
        try:
            btn = page.locator(sel).first
            btn.wait_for(state="visible", timeout=timeout_ms)
            btn.scroll_into_view_if_needed(timeout=2000)
            btn.click(timeout=timeout_ms)
            print(f"[已点击] 登录按钮 ({sel})", flush=True)
            return True
        except Exception:
            continue

    # Last resort: hit Enter while password input still has focus
    try:
        page.keyboard.press("Enter")
        print("[已按回车] 触发登录", flush=True)
        return True
    except Exception:
        return False


def _fill_page_password_and_login(page, password: str, timeout_ms: int = 30000) -> bool:
    """Click password box → type password → click 登录."""
    try:
        page.wait_for_url("**netc2.igtb.boc.cn**", timeout=timeout_ms)
    except Exception:
        pass

    # Wait for the 用户密码 label so the form is fully rendered
    deadline = time.time() + timeout_ms / 1000
    while time.time() < deadline:
        try:
            if page.get_by_text("用户密码", exact=True).first.is_visible(timeout=500):
                break
        except Exception:
            pass
        time.sleep(0.5)

    try:
        page.bring_to_front()
    except Exception:
        pass

    # Step 1: click the password input to focus it
    if not _click_password_input(page, timeout_ms=8000):
        print("[警告] 未找到用户密码输入框", flush=True)
        return False
    time.sleep(0.4)

    # Step 2: clear, then type the password with real key events
    try:
        page.keyboard.press("Control+A")
        time.sleep(0.1)
        page.keyboard.press("Delete")
        time.sleep(0.1)
    except Exception:
        pass

    typed = False
    try:
        page.keyboard.type(password, delay=80)
        typed = True
    except Exception:
        pass
    if not typed:
        try:
            _send_ascii(password)
            typed = True
        except Exception:
            pass
    if not typed:
        print("[警告] 输入用户密码失败", flush=True)
        return False
    print("[已输入] 用户密码", flush=True)
    time.sleep(0.6)

    # Step 3: click the red 登录 button
    if not _click_login_button(page, timeout_ms=6000):
        print("[警告] 未能点击登录按钮", flush=True)
        return False
    return True


def _is_page_password_login_ready(page) -> bool:
    """Return true when the netc2 web password-login page is visible."""
    if "netc2.igtb.boc.cn" in page.url:
        return True
    for frame in page.frames:
        try:
            if frame.get_by_text("用户密码", exact=True).first.is_visible(timeout=300):
                return True
        except Exception:
            pass
    return False


def _maybe_submit_ushield_password(page, password: str, timeout_s: float = 4.0) -> bool:
    """Only enter U-shield PIN if the web login page has not appeared yet."""
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        if _is_page_password_login_ready(page):
            return False
        time.sleep(0.5)
    _send_ascii(password)
    _send_enter()
    time.sleep(10)
    return True


def _click_payment_transfer(page, timeout_ms: int = 60000) -> bool:
    """登录完成后：点击顶部「付款服务」→「转账汇款」。"""
    # 等待主页菜单渲染出来（付款服务 文本可见即可）
    deadline = time.time() + timeout_ms / 1000
    payment_loc = None
    payment_frame = None
    while time.time() < deadline:
        for frame in page.frames:
            try:
                loc = frame.get_by_text("付款服务", exact=True).first
                if loc.is_visible(timeout=300):
                    payment_loc = loc
                    payment_frame = frame
                    break
            except Exception:
                continue
        if payment_loc is not None:
            break
        time.sleep(0.5)

    if payment_loc is None:
        print("[警告] 未找到 付款服务 菜单", flush=True)
        return False

    try:
        payment_loc.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass
    try:
        payment_loc.hover(timeout=3000)
    except Exception:
        pass
    time.sleep(0.3)
    try:
        payment_loc.click(timeout=5000)
        print("[已点击] 付款服务", flush=True)
    except Exception:
        print("[警告] 点击 付款服务 失败", flush=True)
        return False
    time.sleep(0.8)

    # 子菜单中点击「转账汇款」
    end = time.time() + 15
    while time.time() < end:
        frames = [payment_frame] if payment_frame else []
        frames += [f for f in page.frames if f is not payment_frame]
        for frame in frames:
            if frame is None:
                continue
            try:
                sub = frame.get_by_text("转账汇款", exact=True).first
                if sub.is_visible(timeout=300):
                    try:
                        sub.scroll_into_view_if_needed(timeout=1500)
                    except Exception:
                        pass
                    try:
                        sub.click(timeout=5000)
                        print("[已点击] 转账汇款", flush=True)
                        return True
                    except Exception:
                        pass
            except Exception:
                continue
        # 子菜单可能 hover 才出现，再悬停一次
        try:
            payment_loc.hover(timeout=1500)
        except Exception:
            pass
        time.sleep(0.5)

    print("[警告] 未找到 转账汇款 入口", flush=True)
    return False


# 测试用假数据（仅用于填写必填项，不会自动提交）
TEST_PAYEE_ACCOUNT = "000000000000000000"
TEST_PAYEE_NAME = "四川福融泰合供应链管理有限公司"
TEST_PAYEE_BANK_NAME = "中国工商银行成都高新综合保税区支行"
TEST_PAYEE_BANK_CODE = "000000000000000000"
TEST_AMOUNT = "100.00"


def _find_input_by_label(frame, label_text: str):
    """在同一行（或紧邻）查找带星号的必填输入框。"""
    try:
        return frame.evaluate_handle(
            r"""
            (label) => {
                const norm = (s) => (s || '').replace(/\s+/g, '');
                const all = Array.from(document.querySelectorAll('label,div,span,p'));
                const target = all.find((el) => {
                    const t = norm(el.innerText || el.textContent);
                    return t === label || t === '*' + label || t.endsWith(label);
                });
                if (!target) return null;
                let node = target;
                for (let i = 0; i < 6 && node; i++) {
                    const inp = node.querySelector && node.querySelector('input:not([type=hidden]),textarea');
                    if (inp) return inp;
                    node = node.parentElement;
                }
                return null;
            }
            """,
            label_text,
        )
    except Exception:
        return None


def _click_radio_by_text(frame, text: str) -> bool:
    try:
        return bool(frame.evaluate(
            r"""
            (txt) => {
                const norm = (s) => (s || '').replace(/\s+/g, '');
                const labels = Array.from(document.querySelectorAll('label,span,div'));
                for (const el of labels) {
                    if (norm(el.innerText || el.textContent) === txt) {
                        const radio = el.closest('label')?.querySelector('input[type=radio]')
                            || el.parentElement?.querySelector('input[type=radio]')
                            || el.previousElementSibling;
                        const target = radio || el;
                        target.click();
                        return true;
                    }
                }
                return false;
            }
            """,
            text,
        ))
    except Exception:
        return False


def _fill_input_handle(handle, value: str) -> bool:
    if not handle:
        return False
    try:
        handle.evaluate("(el) => { el.focus(); el.value = ''; }")
        handle.type(value, delay=30)
        handle.evaluate(
            "(el) => { el.dispatchEvent(new Event('input',{bubbles:true})); el.dispatchEvent(new Event('change',{bubbles:true})); el.blur(); }"
        )
        return True
    except Exception:
        return False


def _fill_transfer_form(page, timeout_ms: int = 60000) -> bool:
    """填写转账汇款页面的所有红星必填项（测试用假数据）。"""
    deadline = time.time() + timeout_ms / 1000
    target_frame = None
    while time.time() < deadline:
        for frame in page.frames:
            try:
                if frame.get_by_text("收款人账号", exact=False).first.is_visible(timeout=300):
                    target_frame = frame
                    break
            except Exception:
                continue
        if target_frame is not None:
            break
        time.sleep(0.5)

    if target_frame is None:
        print("[警告] 未找到 转账汇款 表单", flush=True)
        return False

    time.sleep(1.0)

    # 收款人账号
    h = _find_input_by_label(target_frame, "收款人账号")
    if _fill_input_handle(h, TEST_PAYEE_ACCOUNT):
        print(f"[已填] 收款人账号 = {TEST_PAYEE_ACCOUNT}", flush=True)
    time.sleep(0.4)

    # 收款人户名
    h = _find_input_by_label(target_frame, "收款人户名")
    if _fill_input_handle(h, TEST_PAYEE_NAME):
        print(f"[已填] 收款人户名 = {TEST_PAYEE_NAME}", flush=True)
    time.sleep(0.4)

    # 收款行类型 = 他行
    if _click_radio_by_text(target_frame, "他行"):
        print("[已选] 收款行类型 = 他行", flush=True)
    time.sleep(0.4)

    # 收款人开户行名称
    h = _find_input_by_label(target_frame, "收款人开户行名称")
    if _fill_input_handle(h, TEST_PAYEE_BANK_NAME):
        print(f"[已填] 收款人开户行名称 = {TEST_PAYEE_BANK_NAME}", flush=True)
    time.sleep(0.6)

    # 收款人开户行行号
    h = _find_input_by_label(target_frame, "收款人开户行行号")
    if _fill_input_handle(h, TEST_PAYEE_BANK_CODE):
        print(f"[已填] 收款人开户行行号 = {TEST_PAYEE_BANK_CODE}", flush=True)
    time.sleep(0.4)

    # 收款人类型 = 单位
    if _click_radio_by_text(target_frame, "单位"):
        print("[已选] 收款人类型 = 单位", flush=True)
    time.sleep(0.4)

    # 金额
    h = _find_input_by_label(target_frame, "金额")
    if _fill_input_handle(h, TEST_AMOUNT):
        print(f"[已填] 金额 = {TEST_AMOUNT}", flush=True)
    time.sleep(0.3)

    print("[完成] 转账表单必填项已填写（未提交）", flush=True)
    return True


def _prepare_chrome_extensions() -> str:
    """Create a temp profile and copy only the BOC certificate extension."""
    chrome_ext_dir = (
        Path(os.environ["LOCALAPPDATA"])
        / "Google" / "Chrome" / "User Data" / "Default" / "Extensions"
    )
    if not chrome_ext_dir.is_dir():
        raise SystemExit(f"[终止] 未找到 Chrome 扩展目录：{chrome_ext_dir}")

    source_ext_dir = chrome_ext_dir / BOC_EXTENSION_ID
    if not source_ext_dir.is_dir():
        raise SystemExit(
            f"[终止] 未找到中行证书扩展 {BOC_EXTENSION_ID}。请先在 Chrome 中安装 {BOC_EXTENSION_NAME}。"
        )

    versions = [
        ver for ver in source_ext_dir.iterdir()
        if ver.is_dir() and (ver / "manifest.json").is_file()
    ]
    if not versions:
        raise SystemExit(f"[终止] 中行证书扩展目录缺少 manifest.json：{source_ext_dir}")

    selected = max(versions, key=lambda ver: ver.stat().st_mtime)
    tmp = Path(tempfile.mkdtemp(prefix="boc_playwright_"))
    try:
        dst = tmp / "Default" / "Extensions" / BOC_EXTENSION_ID / selected.name
        shutil.copytree(selected, dst)
        manifest = selected / "manifest.json"
        m = json.loads(manifest.read_text(encoding="utf-8"))
        print(f"[扩展] {m.get('name', BOC_EXTENSION_NAME)} ({selected.name})", flush=True)
    except Exception:
        shutil.rmtree(tmp, ignore_errors=True)
        raise
    _write_chrome_preferences(tmp)
    return str(tmp)


def _write_chrome_preferences(profile_dir: Path) -> None:
    """Persist Chrome prefs in the temporary Playwright profile."""
    default_dir = profile_dir / "Default"
    default_dir.mkdir(parents=True, exist_ok=True)
    prefs_path = default_dir / "Preferences"
    prefs = {
        "browser": {
            "enable_spellchecking": False,
        },
        "intl": {
            "accept_languages": "zh-CN,zh",
        },
        "translate": {
            "enabled": False,
            "blocked_languages": ["en"],
            "site_blacklist": [
                "netc1.igtb.boc.cn",
                "netc2.igtb.boc.cn",
            ],
        },
        "translate_blocked_languages": ["en"],
        "translate_site_blacklist": [
            "netc1.igtb.boc.cn",
            "netc2.igtb.boc.cn",
        ],
    }
    prefs_path.write_text(
        json.dumps(prefs, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> None:
    _setup_stdout_utf8()
    if not sys.stdout or not sys.stdout.isatty():
        _setup_file_logging()

    if _ITC_OK:
        print(f"[输入后端] Interception 驱动（键盘 slot={_ITC_KB_SLOT}）", flush=True)
    else:
        print("[输入后端] SendInput 回退（证书弹窗可能被拦截）", flush=True)

    _disable_proxy_for_playwright()
    profile_dir = _prepare_chrome_extensions()

    # 收集所有扩展版本目录路径
    ext_base = Path(profile_dir) / "Default" / "Extensions"
    ext_dirs = []
    for ext_id in ext_base.iterdir():
        if not ext_id.is_dir():
            continue
        for ver in ext_id.iterdir():
            if (ver / "manifest.json").is_file():
                ext_dirs.append(str(ver))
                break

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            viewport={"width": 1366, "height": 860},
            locale="zh-CN",
            ignore_default_args=["--disable-extensions"],
            args=[
                "--no-proxy-server",
                "--proxy-server=direct://",
                "--proxy-bypass-list=*",
                "--disable-blink-features=AutomationControlled",
                "--disable-features=Translate,TranslateUI",
                "--lang=zh-CN",
                f"--load-extension={','.join(ext_dirs)}",
            ],
        )
        page = context.new_page()
        page.goto(BOC_URL, wait_until="domcontentloaded")
        _wait_full_load(page, timeout_ms=60000)
        print(f"[已打开] {BOC_URL}", flush=True)
        if _click_certificate_login(page):
            print("[已点击] 数字证书登录", flush=True)
            time.sleep(1)
            _send_enter()
            print("[已按回车] 确认证书", flush=True)
            if BOC_USHIELD_PASSWORD:
                if _maybe_submit_ushield_password(page, BOC_USHIELD_PASSWORD):
                    print("[已输入] U盾密码并提交", flush=True)
                else:
                    print("[跳过] 未出现 U盾密码输入，已进入网页登录页", flush=True)
                if _fill_page_password_and_login(page, BOC_USHIELD_PASSWORD):
                    print("[已输入] 登录页用户密码并点击登录", flush=True)
                    _wait_full_load(page, timeout_ms=60000)
                    if _click_payment_transfer(page, timeout_ms=60000):
                        print("[已进入] 转账汇款", flush=True)
                        _wait_full_load(page, timeout_ms=60000)
                        _fill_transfer_form(page, timeout_ms=60000)
                    else:
                        print("[警告] 未能进入 转账汇款", flush=True)
                else:
                    print("[警告] 未能自动填写登录页用户密码", flush=True)
            else:
                print("[跳过] 未配置 BOC_USHIELD_PASSWORD", flush=True)
        else:
            print("[警告] 未找到或未能点击：数字证书登录", flush=True)
        print(f"[完成] 企业网银登录页：{page.url}", flush=True)
        print("浏览器保持打开，关闭窗口或按 Ctrl+C 退出。", flush=True)

        try:
            while context.browser.is_connected():
                time.sleep(1)
        except KeyboardInterrupt:
            pass
        finally:
            try:
                context.close()
            except Exception:
                pass
            shutil.rmtree(profile_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
