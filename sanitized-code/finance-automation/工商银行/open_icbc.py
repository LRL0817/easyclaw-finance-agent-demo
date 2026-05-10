"""
工商银行单笔支付 - 入口脚本

使用 Playwright 启动一个有头 Chromium 浏览器，打开工商银行官网
（https://www.icbc.com.cn/），自动跳转到企业网银登录页，点击 U盾登录，
并在原生「选择证书」与 U盾 PIN 弹窗中自动按回车 / 输入密码。

浏览器窗口被用户关闭或脚本收到 Ctrl+C 时退出。

依赖：
    pip install playwright
    python -m playwright install chromium
"""

import ctypes
import os
import sys
import time
from dataclasses import dataclass
from pathlib import Path


from playwright.sync_api import BrowserContext, Frame, Locator, Page, sync_playwright

ICBC_URL = "https://www.icbc.com.cn/"
CORP_LOGIN_TEXT = "企业网上银行登录"
USHIELD_LOGIN_TEXT = "U盾登录"
PAYMENT_MENU_TEXT = "付款业务"
SINGLE_PAYMENT_TEXT = "逐笔支付"

ENV_FILE = Path(__file__).resolve().parent / ".env"
LOG_FILE = Path(__file__).resolve().parent / "icbc.log"
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


def _disable_proxy_for_playwright() -> None:
    """Keep this script's Playwright browser on a direct connection."""
    for key in PROXY_ENV_KEYS:
        os.environ.pop(key, None)


def _load_env_file(path: Path) -> None:
    """极简 .env 解析：KEY=VALUE，每行一条；忽略空行/#注释；已存在的环境变量不覆盖。"""
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

# U盾 PIN：从环境变量 / .env 读取。请把密码放在与脚本同级的 .env 中：
#     ICBC_USHIELD_PASSWORD=你的U盾PIN
USHIELD_PASSWORD = os.environ.get("ICBC_USHIELD_PASSWORD", "")


@dataclass(frozen=True)
class PaymentData:
    """单笔付款页面需要自动填入的数据；必须由 .env 提供，避免误用历史默认值。"""

    payee_name: str
    payee_account: str
    payee_bank: str
    amount: str
    purpose: str
    remark: str

    @classmethod
    def from_env(cls) -> "PaymentData":
        required = {
            "payee_name": "ICBC_PAYEE_NAME",
            "payee_account": "ICBC_PAYEE_ACCOUNT",
            "payee_bank": "ICBC_PAYEE_BANK",
            "amount": "ICBC_PAYMENT_AMOUNT",
            "purpose": "ICBC_PAYMENT_PURPOSE",
        }
        values: dict[str, str] = {}
        missing: list[str] = []
        for field, key in required.items():
            val = os.environ.get(key, "").strip()
            if not val:
                missing.append(key)
            values[field] = val
        if missing:
            raise SystemExit(
                "[终止] 以下付款字段未在 .env 中配置，拒绝使用历史默认值："
                + ", ".join(missing)
            )
        values["remark"] = os.environ.get("ICBC_PAYMENT_REMARK", "").strip()
        return cls(**values)


PAYMENT_DATA = PaymentData.from_env()


# ----------- 键盘自动化（用于 Chromium 原生弹窗 / U盾 PIN 弹窗） -----------
# 关键：U盾 PIN 弹窗是工行的安全输入控件，会通过低级钩子检测 LLKHF_INJECTED，
# SendInput 注入会被丢弃。优先走 Interception 内核驱动（HID 层模拟，无注入标志）；
# 驱动未安装/未加载时回退到 SendInput 扫描码（适用于系统弹窗如证书选择）。
#
# 注意：interception 的 `auto_capture_devices` 会从 slot 0 开始挑第一个 is_keyboard
# 的设备，HP 笔记本上 slot 0 通常是 `ACPI\\VEN_HPQ&DEV_8001`（电源/Fn 热键 ACPI
# 设备），不是真正的键盘。往那个 slot 发按键，应用层完全收不到。所以这里跳过
# ACPI / 无 HWID 的 slot，专门挑 HWID 含 `HID\\VID_` 的真键盘。
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
except Exception as _e:  # 驱动未装/未加载/包未装
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
    """用 SendInput + 扫描码发送一个按键（按下+抬起）。绕过部分安全输入过滤。"""
    scan = _USER32.MapVirtualKeyW(vk, _MAPVK_VK_TO_VSC)
    extended = vk in (0x0D,) and False  # Enter 不是扩展键

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
    """逐字发送 ASCII 字母/数字。优先走 Interception（绕过 U盾安全输入），
    驱动不可用时回退到 SendInput 扫描码。"""
    if _ITC_OK:
        # Interception 的 write 仅小写；对工行 PIN 通常都是数字/小写字母，足够。
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


# ----------- 工具 -----------
def _setup_stdout_utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _setup_file_logging() -> None:
    """pythonw 启动时也保留一份日志，方便排查页面定位问题。"""
    try:
        log = LOG_FILE.open("a", encoding="utf-8", buffering=1)
        sys.stdout = log
        sys.stderr = log
        print("\n========== ICBC run ==========", flush=True)
    except Exception:
        pass


def _wait_full_load(page: Page, timeout_ms: int = 60000) -> None:
    """先等 DOM，再尽量等 networkidle；网慢时 networkidle 可能超时，忽略即可。"""
    try:
        page.wait_for_load_state("domcontentloaded", timeout=timeout_ms)
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=timeout_ms)
    except Exception:
        pass


def _xpath_literal(text: str) -> str:
    """把任意字符串转成 XPath 字符串字面量。"""
    if "'" not in text:
        return f"'{text}'"
    if '"' not in text:
        return f'"{text}"'
    parts = text.split("'")
    return "concat(" + ', "\'", '.join(f"'{part}'" for part in parts) + ")"


def _all_frames(page: Page) -> list[Frame]:
    return [page.main_frame, *[frame for frame in page.frames if frame != page.main_frame]]


def _label_variants(label: str) -> list[str]:
    aliases = {
        "收款账号": ["收款账号", "收款账户"],
        "收款账户": ["收款账号", "收款账户"],
        "收款银行/行别": ["收款银行/行别", "收款银行行别"],
    }
    return aliases.get(label, [label])


def _try_visible(locator: Locator, timeout_ms: int = 1000) -> bool:
    try:
        locator.wait_for(state="visible", timeout=timeout_ms)
        return True
    except Exception:
        return False


def _fill_locator(locator: Locator, value: str) -> bool:
    """尽量用真实输入事件填值；失败时回退到 DOM 赋值并派发 input/change。"""
    try:
        locator.click(timeout=2000)
        locator.fill(value, timeout=3000)
        return True
    except Exception:
        pass

    try:
        locator.evaluate(
            """(el, value) => {
                el.focus();
                el.value = value;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }""",
            value,
        )
        return True
    except Exception:
        return False


def _fill_by_label_geometry(page: Page, label: str, value: str) -> bool:
    """按页面可见标签的位置，填同一行右侧最近的输入框。适配工行自定义表单。"""
    script = """({ label, value }) => {
        const normalize = (text) => (text || '')
            .replace(/[\\s:*：]/g, '')
            .trim();
        const wanted = normalize(label);
        const visible = (el) => {
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none'
                && style.visibility !== 'hidden'
                && rect.width > 0
                && rect.height > 0;
        };
        const textNodes = Array.from(document.querySelectorAll('label, span, div, td, th, p'))
            .filter((el) => {
                if (!visible(el) || normalize(el.innerText || el.textContent) !== wanted) {
                    return false;
                }
                const rect = el.getBoundingClientRect();
                return rect.width < 260 && rect.height < 90;
            });
        const controls = Array.from(document.querySelectorAll('input, textarea, [contenteditable="true"]'))
            .filter((el) => visible(el) && !el.disabled && el.type !== 'hidden');

        for (const textEl of textNodes) {
            textEl.scrollIntoView({ block: 'center', inline: 'nearest' });
            const labelRect = textEl.getBoundingClientRect();
            const labelY = labelRect.top + labelRect.height / 2;
            const candidates = controls
                .map((control) => {
                    const rect = control.getBoundingClientRect();
                    const controlY = rect.top + rect.height / 2;
                    return {
                        control,
                        rect,
                        rowGap: Math.abs(controlY - labelY),
                        xGap: rect.left - labelRect.right,
                    };
                })
                .filter((item) => item.xGap > -20 && item.rowGap < 48)
                .sort((a, b) => (a.rowGap - b.rowGap) || (a.xGap - b.xGap));
            const match = candidates[0]?.control;
            if (!match) {
                continue;
            }

            match.focus();
            if (match.isContentEditable) {
                match.textContent = value;
            } else {
                const proto = match instanceof HTMLTextAreaElement
                    ? HTMLTextAreaElement.prototype
                    : HTMLInputElement.prototype;
                const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
                if (setter) {
                    setter.call(match, value);
                } else {
                    match.value = value;
                }
            }
            for (const name of ['input', 'change', 'blur']) {
                match.dispatchEvent(new Event(name, { bubbles: true }));
            }
            return true;
        }
        return false;
    }"""
    for variant in _label_variants(label):
        for frame in _all_frames(page):
            try:
                if frame.evaluate(script, {"label": variant, "value": value}):
                    print(f"[已填入] {label}：{value}", flush=True)
                    return True
            except Exception:
                pass
    return False


def _focus_by_label_geometry(page: Page, label: str) -> bool:
    """把焦点放到指定标签同行右侧的输入框上。"""
    script = """(label) => {
        const normalize = (text) => (text || '')
            .replace(/[\\s:*：]/g, '')
            .trim();
        const wanted = normalize(label);
        const visible = (el) => {
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none'
                && style.visibility !== 'hidden'
                && rect.width > 0
                && rect.height > 0;
        };
        const textNodes = Array.from(document.querySelectorAll('label, span, div, td, th, p'))
            .filter((el) => {
                if (!visible(el) || normalize(el.innerText || el.textContent) !== wanted) {
                    return false;
                }
                const rect = el.getBoundingClientRect();
                return rect.width < 260 && rect.height < 90;
            });
        const controls = Array.from(document.querySelectorAll('input, textarea, [contenteditable="true"]'))
            .filter((el) => visible(el) && !el.disabled && el.type !== 'hidden');

        for (const textEl of textNodes) {
            textEl.scrollIntoView({ block: 'center', inline: 'nearest' });
            const labelRect = textEl.getBoundingClientRect();
            const labelY = labelRect.top + labelRect.height / 2;
            const candidates = controls
                .map((control) => {
                    const rect = control.getBoundingClientRect();
                    const controlY = rect.top + rect.height / 2;
                    return {
                        control,
                        rowGap: Math.abs(controlY - labelY),
                        xGap: rect.left - labelRect.right,
                    };
                })
                .filter((item) => item.xGap > -20 && item.rowGap < 48)
                .sort((a, b) => (a.rowGap - b.rowGap) || (a.xGap - b.xGap));
            const match = candidates[0]?.control;
            if (!match) {
                continue;
            }
            match.focus();
            match.click();
            return true;
        }
        return false;
    }"""
    for variant in _label_variants(label):
        for frame in _all_frames(page):
            try:
                if frame.evaluate(script, variant):
                    return True
            except Exception:
                pass
    return False


def _fill_by_label_keyboard(page: Page, label: str, value: str) -> bool:
    """对会吞 DOM 赋值的控件，用页面键盘输入一遍。"""
    if not _focus_by_label_geometry(page, label):
        return False
    try:
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        page.keyboard.insert_text(value)
        time.sleep(0.6)
        print(f"[已键入] {label}：{value}", flush=True)
        return True
    except Exception:
        return False


def _blur_active_control(page: Page) -> None:
    """收起当前下拉/弹层。不要点固定坐标，会误中其它输入框（之前 click(360,360)
    会落在表单区域、把 收款账号/收款单位 焦点搅乱）。只发 Esc 即可。"""
    try:
        page.keyboard.press("Escape")
        time.sleep(0.2)
    except Exception:
        pass


def _click_first_dropdown_item(page: Page) -> bool:
    """点击当前聚焦输入框正下方下拉里的第一个可选项。

    关键：必须以 document.activeElement 为锚点过滤候选——否则 div/span
    选择器会扫到左侧导航栏的菜单项（之前就误点了「指令查询」）。把候选限制
    在聚焦输入框正下方一定范围内，左侧导航就被自然排除了。"""
    script = """() => {
        const focused = document.activeElement;
        const fRect = (focused && focused.getBoundingClientRect)
            ? focused.getBoundingClientRect() : null;
        const visible = (el) => {
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none'
                && style.visibility !== 'hidden'
                && rect.width > 0
                && rect.height > 0;
        };
        const ignored = new Set([
            '热门银行', 'ABCDE', 'FGHJ', 'KMN', 'PQRST', 'WXYZ', '手工录入',
            '搜索银行名称', '可下拉查询，模糊匹配',
            '指令查询', '账户查询', '我的工作台'
        ]);
        // 候选必须在弹层容器内：用 role / 高 z-index 浮层 / 常见下拉类名判断；
        // 否则会掉到下一行表单标签（如「*汇款金额： 调整限额」）上。
        const POPUP_ROLES = new Set(['dialog', 'menu', 'listbox', 'tooltip', 'tablist', 'option', 'menuitem']);
        const POPUP_CLASS_RE = /(dropdown|select|popup|popover|menu|listbox|combobox|cascader|picker|panel|overlay|modal)/i;
        const isInPopup = (el) => {
            let node = el;
            for (let depth = 0; node && node !== document.body && depth < 12; depth++) {
                const role = node.getAttribute && node.getAttribute('role');
                if (role && POPUP_ROLES.has(role)) return true;
                const cls = (node.className && node.className.baseVal) || node.className || '';
                if (typeof cls === 'string' && POPUP_CLASS_RE.test(cls)) return true;
                const style = getComputedStyle(node);
                const z = parseInt(style.zIndex, 10);
                if ((style.position === 'fixed' || style.position === 'absolute') && !isNaN(z) && z >= 100) {
                    return true;
                }
                node = node.parentElement;
            }
            return false;
        };
        const looksLikeFormLabel = (text) => {
            // 工行表单标签都带 * 前缀或末尾 :/： 冒号，且常含「调整限额」「收款」「汇款」等。
            return text.startsWith('*')
                || /[：:]\\s*$/.test(text)
                || /(调整限额|收款单位|收款账号|收款银行|汇款金额|汇款用途|备注|预约执行)/.test(text);
        };
        const candidates = Array.from(document.querySelectorAll(
            '[role="option"], [role="menuitem"], li, .el-select-dropdown__item, .ant-select-item-option, .ivu-select-item, div, span'
        ))
            .filter((el) => {
                if (!visible(el)) return false;
                const rect = el.getBoundingClientRect();
                const text = (el.innerText || el.textContent || '').trim();
                if (!text || text.length < 2 || text.length > 80) return false;
                if (ignored.has(text)) return false;
                if (text.includes('搜索银行名称')) return false;
                if (looksLikeFormLabel(text)) return false;
                if (rect.width < 30 || rect.height < 18) return false;
                if (fRect) {
                    if (rect.top < fRect.bottom - 5) return false;
                    if (rect.top > fRect.bottom + 600) return false;
                    const cx = rect.left + rect.width / 2;
                    if (cx < fRect.left - 100 || cx > fRect.right + 320) return false;
                } else {
                    if (rect.top < 120) return false;
                }
                if (!isInPopup(el)) return false;
                return true;
            })
            .sort((a, b) => {
                const ar = a.getBoundingClientRect();
                const br = b.getBoundingClientRect();
                return (ar.top - br.top) || (ar.left - br.left);
            });
        const item = candidates[0];
        if (!item) {
            return null;
        }
        item.click();
        return (item.innerText || item.textContent || '').trim();
    }"""
    for frame in _all_frames(page):
        try:
            text = frame.evaluate(script)
            if text:
                print(f"[已点击第一个下拉项] {text}", flush=True)
                return True
        except Exception:
            pass
    return False


def _input_candidates(frame: Frame, label: str) -> list[Locator]:
    text = _xpath_literal(label)
    return [
        frame.get_by_label(label).first,
        frame.get_by_placeholder(label).first,
        frame.locator(
            "xpath=//*[self::label or self::span or self::div or self::td]"
            f"[contains(normalize-space(.), {text})]"
            "/following::*[self::input or self::textarea][1]"
        ).first,
        frame.locator(
            "xpath=//*[self::label or self::span or self::div or self::td]"
            f"[contains(normalize-space(.), {text})]"
            "/following::*[@contenteditable='true'][1]"
        ).first,
    ]


def _fill_by_label(page: Page, label: str, value: str) -> bool:
    if _fill_by_label_geometry(page, label, value):
        return True

    strict_labels = {
        "收款单位",
        "收款账号",
        "收款银行/行别",
        "汇款金额",
        "汇款用途",
        "备注",
    }
    if label in strict_labels:
        print(f"[警告] 未能精确定位「{label}」", flush=True)
        return False

    for frame in _all_frames(page):
        for candidate in _input_candidates(frame, label):
            if _try_visible(candidate) and _fill_locator(candidate, value):
                print(f"[已填入] {label}：{value}", flush=True)
                return True
    print(f"[警告] 未能填入「{label}」", flush=True)
    return False


def _click_matching_option(page: Page, text: str, timeout_ms: int = 1500) -> bool:
    for frame in _all_frames(page):
        option = frame.get_by_text(text, exact=True).last
        if _try_visible(option, timeout_ms=timeout_ms):
            try:
                option.click(timeout=2000)
                return True
            except Exception:
                pass
    return False


def _bank_search_term(value: str) -> str:
    """工行银行 picker 只要银行总行名，不要支行全称。"""
    index = value.find("银行")
    if index >= 0:
        return value[: index + len("银行")]
    return value.strip()


def _fill_bank_search_box(page: Page, value: str) -> bool:
    """只在银行选择弹窗里的「搜索银行名称」输入框填写银行总行名。

    页面右上角红色导航栏里有一个全站搜索框，placeholder 也含「搜索」二字，
    所以 input[placeholder*='搜索'] 这种宽泛选择器会误命中那里。这里改成
    先按弹窗内特征文字（热门银行 / ABCDE / FGHJ / KMN / PQRST / WXYZ / 手工
    录入）锁定弹窗容器，再仅在该容器内部查找输入框，并校验输入框的 bounding
    box 必须落在弹窗容器内。找不到就返回 False，让上层 fallback 接管。"""

    script = r"""(value) => {
        const visible = (el) => {
            if (!el || !el.getBoundingClientRect) return false;
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none'
                && style.visibility !== 'hidden'
                && style.opacity !== '0'
                && rect.width > 0
                && rect.height > 0;
        };
        const POPUP_MARKERS = ['热门银行', 'ABCDE', 'FGHJ', 'KMN', 'PQRST', 'WXYZ', '手工录入'];
        const findMarkerEls = (marker) => {
            const out = [];
            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null);
            while (walker.nextNode()) {
                const node = walker.currentNode;
                const text = (node.nodeValue || '').replace(/\s+/g, '');
                if (!text || !text.includes(marker)) continue;
                const el = node.parentElement;
                if (el && visible(el)) out.push(el);
            }
            return out;
        };
        // 锚点：「热门银行」是弹窗里最稳定的标识。
        const anchors = findMarkerEls('热门银行');
        if (!anchors.length) return { ok: false, reason: 'no-anchor' };

        // 沿 anchor 向上找一个同时包含尽可能多 marker 的祖先作为弹窗容器。
        const containsText = (root, marker) => {
            const t = (root.innerText || root.textContent || '').replace(/\s+/g, '');
            return t.includes(marker);
        };
        let popup = null;
        let bestScore = -1;
        for (const anchor of anchors) {
            let node = anchor;
            for (let depth = 0; node && node !== document.body && depth < 18; depth++) {
                if (visible(node)) {
                    let score = 0;
                    for (const m of POPUP_MARKERS) {
                        if (containsText(node, m)) score++;
                    }
                    const r = node.getBoundingClientRect();
                    // 弹窗一般中等大小、不会撑满整个文档，也不会贴在最顶部（顶部 60px 是导航栏区）。
                    const reasonable = r.width >= 280 && r.height >= 200
                        && r.width <= window.innerWidth + 4
                        && r.top >= 0;
                    if (score >= 3 && reasonable && score > bestScore) {
                        bestScore = score;
                        popup = node;
                    }
                }
                node = node.parentElement;
            }
        }
        if (!popup) return { ok: false, reason: 'no-popup' };

        const popupRect = popup.getBoundingClientRect();
        const inside = (rect) => {
            return rect.left >= popupRect.left - 1
                && rect.top >= popupRect.top - 1
                && rect.right <= popupRect.right + 1
                && rect.bottom <= popupRect.bottom + 1;
        };

        // 只在弹窗容器内部找候选输入框。
        const all = Array.from(popup.querySelectorAll('input, textarea, [contenteditable="true"]'))
            .filter((el) => visible(el) && !el.disabled && el.type !== 'hidden' && inside(el.getBoundingClientRect()));
        if (!all.length) return { ok: false, reason: 'no-input-in-popup' };

        const score = (el) => {
            const ph = (el.placeholder || '') + ' ' + (el.getAttribute('aria-label') || '');
            if (ph.includes('搜索银行名称')) return 0;
            if (ph.includes('银行名称')) return 1;
            if (ph.includes('搜索')) return 2;
            return 3;
        };
        all.sort((a, b) => {
            const sa = score(a), sb = score(b);
            if (sa !== sb) return sa - sb;
            return a.getBoundingClientRect().top - b.getBoundingClientRect().top;
        });

        const match = all[0];
        // 二次校验：bounding box 必须真的落在弹窗内。
        const mr = match.getBoundingClientRect();
        if (!inside(mr)) return { ok: false, reason: 'input-outside-popup' };

        match.focus();
        try { match.click(); } catch (e) {}
        if (match.isContentEditable) {
            match.textContent = '';
            match.textContent = value;
        } else {
            const proto = match instanceof HTMLTextAreaElement
                ? HTMLTextAreaElement.prototype
                : HTMLInputElement.prototype;
            const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
            if (setter) {
                setter.call(match, '');
                setter.call(match, value);
            } else {
                match.value = value;
            }
        }
        for (const name of ['input', 'change']) {
            match.dispatchEvent(new Event(name, { bubbles: true }));
        }
        return {
            ok: true,
            placeholder: match.placeholder || match.getAttribute('aria-label') || '',
            popupRect: { left: popupRect.left, top: popupRect.top, right: popupRect.right, bottom: popupRect.bottom },
            inputRect: { left: mr.left, top: mr.top, right: mr.right, bottom: mr.bottom },
        };
    }"""

    for frame in _all_frames(page):
        try:
            result = frame.evaluate(script, value)
        except Exception:
            continue
        if isinstance(result, dict) and result.get("ok"):
            ph = result.get("placeholder", "")
            print(f"[弹窗内填写] 搜索银行名称（placeholder={ph!r}）", flush=True)
            return True
    print("[警告] 未在银行弹窗内找到「搜索银行名称」输入框，跳过弹窗内填写", flush=True)
    return False


def _click_bank_option(page: Page, text: str) -> bool:
    """只点击银行 picker 里的精确银行候选。"""
    script = """(text) => {
        const visible = (el) => {
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none'
                && style.visibility !== 'hidden'
                && rect.width > 0
                && rect.height > 0;
        };
        const POPUP_ROLES = new Set(['dialog', 'menu', 'listbox', 'option', 'menuitem']);
        const POPUP_CLASS_RE = /(dropdown|select|popup|popover|menu|listbox|combobox|cascader|picker|panel|overlay|modal)/i;
        const isInPopup = (el) => {
            let node = el;
            for (let depth = 0; node && node !== document.body && depth < 12; depth++) {
                const role = node.getAttribute && node.getAttribute('role');
                if (role && POPUP_ROLES.has(role)) return true;
                const cls = (node.className && node.className.baseVal) || node.className || '';
                if (typeof cls === 'string' && POPUP_CLASS_RE.test(cls)) return true;
                const style = getComputedStyle(node);
                const z = parseInt(style.zIndex, 10);
                if ((style.position === 'fixed' || style.position === 'absolute') && !isNaN(z) && z >= 100) {
                    return true;
                }
                node = node.parentElement;
            }
            return false;
        };
        const candidates = Array.from(document.querySelectorAll('[role="option"], [role="menuitem"], li, div, span, a'))
            .filter((el) => {
                if (!visible(el)) return false;
                const value = (el.innerText || el.textContent || '').trim();
                if (value !== text) return false;
                if (!isInPopup(el)) return false;
                const rect = el.getBoundingClientRect();
                return rect.width >= 30 && rect.height >= 18;
            })
            .sort((a, b) => {
                const ar = a.getBoundingClientRect();
                const br = b.getBoundingClientRect();
                return (br.top - ar.top) || (br.left - ar.left);
            });
        const item = candidates[0];
        if (!item) return false;
        item.click();
        return true;
    }"""
    for frame in _all_frames(page):
        try:
            if frame.evaluate(script, text):
                return True
        except Exception:
            pass

    return _click_matching_option(page, text, timeout_ms=2000)


def _fill_autocomplete(page: Page, label: str, value: str) -> bool:
    if label in ("收款单位", "收款账号", "收款账户"):
        filled = _fill_by_label_keyboard(page, label, value) or _fill_by_label(page, label, value)
    else:
        filled = _fill_by_label(page, label, value)
    if not filled:
        return False

    # 工行页面的收款人/行别常是模糊匹配框：填入后回车或点下拉项确认。
    time.sleep(0.8)
    if _click_matching_option(page, value):
        print(f"[已选择] {label} 下拉项", flush=True)
        _blur_active_control(page)
        return True
    try:
        page.keyboard.press("Enter")
        time.sleep(0.3)
        page.keyboard.press("Tab")
    except Exception:
        pass
    _blur_active_control(page)
    return True


def _wait_bank_popup(page: Page, timeout_ms: int = 4000) -> bool:
    """等银行弹窗渲染完成；用「热门银行」tab 文字作为弹窗就绪信号。"""
    deadline = time.monotonic() + timeout_ms / 1000.0
    while time.monotonic() < deadline:
        for frame in _all_frames(page):
            try:
                tab = frame.get_by_text("热门银行", exact=True).first
                if _try_visible(tab, timeout_ms=300):
                    return True
            except Exception:
                pass
    return False


def _fill_bank_and_click_first(page: Page, value: str) -> bool:
    search_term = _bank_search_term(value)
    if not _focus_by_label_geometry(page, "收款银行/行别"):
        print("[警告] 未能打开「收款银行/行别」下拉", flush=True)
        return False

    if not _wait_bank_popup(page, timeout_ms=4000):
        print("[警告] 未在 4s 内看到银行弹窗的「热门银行」tab", flush=True)
    time.sleep(0.6)
    if _fill_bank_search_box(page, search_term):
        print(f"[已键入] 银行搜索名称：{search_term}", flush=True)
        time.sleep(1.2)
        if _click_bank_option(page, search_term):
            print(f"[已选择] 收款银行/行别：{search_term}", flush=True)
            _blur_active_control(page)
            return True

    print(f"[警告] 未能通过银行弹窗选择「{search_term}」，尝试通用下拉确认", flush=True)
    if _fill_by_label_keyboard(page, "收款银行/行别", search_term) and _click_first_dropdown_item(page):
        _blur_active_control(page)
        return True
    try:
        page.keyboard.press("ArrowDown")
        time.sleep(0.2)
        page.keyboard.press("Enter")
        print("[已键盘确认] 收款银行/行别第一个选项", flush=True)
    except Exception:
        pass
    _blur_active_control(page)
    return True


def _select_dropdown_value(page: Page, label: str, value: str) -> bool:
    """打开下拉控件并选择指定文本项。"""
    if not _focus_by_label_geometry(page, label):
        if not _fill_by_label(page, label, value):
            return False
    try:
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        page.keyboard.insert_text(value)
        time.sleep(0.5)
    except Exception:
        pass

    try:
        page.keyboard.press("ArrowDown")
        time.sleep(0.2)
        page.keyboard.press("Enter")
        time.sleep(0.3)
        _blur_active_control(page)
        print(f"[已键盘选择] {label}：{value}", flush=True)
        return True
    except Exception:
        pass

    if _click_matching_option(page, value, timeout_ms=2000):
        print(f"[已选择] {label}：{value}", flush=True)
        _blur_active_control(page)
        return True

    try:
        page.keyboard.press("Enter")
        time.sleep(0.2)
    except Exception:
        pass
    _blur_active_control(page)
    print(f"[已填入] {label}：{value}", flush=True)
    return True


def _click_text(page: Page, text: str, timeout_ms: int = 15000) -> bool:
    for frame in _all_frames(page):
        target = frame.get_by_text(text, exact=True).first
        if not _try_visible(target, timeout_ms=1000):
            continue
        try:
            target.click(timeout=timeout_ms)
            print(f"[已点击] {text}", flush=True)
            return True
        except Exception:
            try:
                target.click(force=True, timeout=timeout_ms)
                print(f"[已点击] {text}（force）", flush=True)
                return True
            except Exception:
                pass
    return False


# ----------- 业务步骤 -----------
def _click_corp_login(context: BrowserContext, page: Page) -> Page:
    """点击「企业网上银行登录」。若打开新标签页则切到新页并返回该页。"""
    try:
        target = page.get_by_text(CORP_LOGIN_TEXT, exact=True).first
        target.wait_for(state="visible", timeout=20000)
    except Exception as e:
        print(f"[警告] 未找到「{CORP_LOGIN_TEXT}」入口：{e}", flush=True)
        return page

    try:
        with context.expect_page(timeout=10000) as new_page_info:
            target.click()
        new_page = new_page_info.value
        new_page.bring_to_front()
        print(f"[已点击] {CORP_LOGIN_TEXT} → 新标签页：{new_page.url}", flush=True)
        return new_page
    except Exception:
        print(f"[已点击] {CORP_LOGIN_TEXT} → 当前页跳转：{page.url}", flush=True)
        return page


def _click_ushield_login(context: BrowserContext, page: Page) -> Page:
    """在企业网银登录页点击「U盾登录」，并自动应付证书选择和 PIN 输入弹窗。"""
    print("[等待] 第二页加载完成（domcontentloaded + networkidle）…", flush=True)
    _wait_full_load(page, timeout_ms=60000)
    print(f"[就绪] 第二页：{page.url}", flush=True)

    try:
        target = page.get_by_text(USHIELD_LOGIN_TEXT, exact=True).first
        target.wait_for(state="visible", timeout=30000)
    except Exception as e:
        print(f"[警告] 未找到「{USHIELD_LOGIN_TEXT}」按钮：{e}", flush=True)
        return page

    try:
        target.click(no_wait_after=True)
        print(f"[已点击] {USHIELD_LOGIN_TEXT}", flush=True)
    except Exception as e:
        print(f"[警告] 点击「{USHIELD_LOGIN_TEXT}」失败：{e}", flush=True)
        return page

    # 1) 等证书弹窗出现，按回车确认
    time.sleep(1)
    _send_enter()
    print("[已按回车] 确认证书", flush=True)

    # 2) 等 PIN 弹窗出现，直接输入密码 + 回车
    if not USHIELD_PASSWORD:
        print(
            "[跳过] 未配置 U盾 PIN：请在脚本同级 .env 中设置 ICBC_USHIELD_PASSWORD",
            flush=True,
        )
        return page
    time.sleep(10)
    _send_ascii(USHIELD_PASSWORD)
    _send_enter()
    print("[已输入] U盾密码并提交（第 1 次）", flush=True)

    # 3) 工行 U盾会再弹一次 PIN 输入框，重复一次：等弹窗 → 输密码 + 回车
    time.sleep(2)
    _send_ascii(USHIELD_PASSWORD)
    _send_enter()
    print("[已输入] U盾密码并提交（第 2 次）", flush=True)

    return page


def _click_payment_menu(page: Page) -> None:
    """登录完成、落在企业网银首页后，点击顶部菜单「付款业务」。"""
    print("[等待] 登录后首页加载完成（domcontentloaded + networkidle）…", flush=True)
    _wait_full_load(page, timeout_ms=60000)
    print(f"[就绪] 登录后首页：{page.url}", flush=True)

    try:
        target = page.get_by_text(PAYMENT_MENU_TEXT, exact=True).first
        target.wait_for(state="visible", timeout=30000)
    except Exception as e:
        print(f"[警告] 未找到「{PAYMENT_MENU_TEXT}」菜单：{e}", flush=True)
        return

    try:
        target.click()
        print(f"[已点击] {PAYMENT_MENU_TEXT}", flush=True)
    except Exception:
        # 广告浮层可能拦截点击，跳过 actionability 检查再来一次
        try:
            target.click(force=True)
            print(f"[已点击] {PAYMENT_MENU_TEXT}（force）", flush=True)
        except Exception as e:
            print(f"[警告] 点击「{PAYMENT_MENU_TEXT}」失败：{e}", flush=True)


def _open_single_payment_page(page: Page) -> None:
    """进入「付款业务 - 逐笔支付」页面；若已经在该页，则这一步自然跳过。"""
    print("[准备] 打开逐笔支付页面", flush=True)
    _wait_full_load(page, timeout_ms=30000)
    if _click_text(page, SINGLE_PAYMENT_TEXT, timeout_ms=15000):
        _wait_full_load(page, timeout_ms=30000)
        return
    print(f"[警告] 未找到「{SINGLE_PAYMENT_TEXT}」，请确认当前已进入付款业务页面", flush=True)


def _fill_text_only(page: Page, label: str, value: str) -> bool:
    """填入纯文本字段：键入后只发 Esc 收起可能弹出的浮层，不按 Enter，
    也不点任何下拉候选——避免触发 ICBC 表单的联动校验/清空逻辑。
    汇款用途、备注、汇款金额、收款账号都走这条路径。"""
    filled = _fill_by_label_keyboard(page, label, value) or _fill_by_label(page, label, value)
    if not filled:
        return False
    try:
        page.keyboard.press("Escape")
        time.sleep(0.3)
    except Exception:
        pass
    return True


def _fill_payee_name_first(page: Page, value: str) -> bool:
    """收款单位必须最先填。

    工行的「收款单位」是带地址簿模糊匹配的 autocomplete 输入框，有两个坑：
    1) 按 Esc 会被解释成「取消输入」，把已键入但还没从下拉选中的文本清空。
    2) blur 时若没匹配到地址簿条目，部分模式下也会被清空。

    所以这里改成：键入 → 等下拉浮出 → 优先点完全匹配的下拉项（提交为
    地址簿选择）→ 没匹配项就保留手工录入文本，不发 Esc / Tab，让后续
    填「收款账号」的点击自然触发 blur。"""
    filled = _fill_by_label_keyboard(page, "收款单位", value) or _fill_by_label(
        page, "收款单位", value
    )
    if not filled:
        return False
    time.sleep(1.0)
    if _click_matching_option(page, value, timeout_ms=2000):
        print(f"[已选择] 收款单位地址簿匹配项：{value}", flush=True)
    return True


def _ensure_payee_name_filled(page: Page, value: str) -> None:
    """收款单位 autocomplete 在 blur 时若没匹配地址簿条目会把已键入的文本清空。
    填完整张表后这里再核对一次：还原成空就用 DOM 原生 setter 直接写回，并仅
    派发 input/change 事件，避免再次触发 blur 清空逻辑。"""
    script = """({ label, value }) => {
        const normalize = (text) => (text || '').replace(/[\\s:*：]/g, '').trim();
        const wanted = normalize(label);
        const visible = (el) => {
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none'
                && style.visibility !== 'hidden'
                && rect.width > 0
                && rect.height > 0;
        };
        const textNodes = Array.from(document.querySelectorAll('label, span, div, td, th, p'))
            .filter((el) => {
                if (!visible(el) || normalize(el.innerText || el.textContent) !== wanted) return false;
                const r = el.getBoundingClientRect();
                return r.width < 260 && r.height < 90;
            });
        const controls = Array.from(document.querySelectorAll('input, textarea, [contenteditable="true"]'))
            .filter((el) => visible(el) && !el.disabled && el.type !== 'hidden');
        for (const t of textNodes) {
            const lr = t.getBoundingClientRect();
            const ly = lr.top + lr.height / 2;
            const cands = controls
                .map((c) => {
                    const r = c.getBoundingClientRect();
                    return {
                        c,
                        dy: Math.abs(r.top + r.height / 2 - ly),
                        dx: r.left - lr.right,
                    };
                })
                .filter((x) => x.dx > -20 && x.dy < 48)
                .sort((a, b) => (a.dy - b.dy) || (a.dx - b.dx));
            const match = cands[0]?.c;
            if (!match) continue;
            const current = match.isContentEditable ? (match.textContent || '') : (match.value || '');
            if (current === value) return 'ok';
            match.focus();
            if (match.isContentEditable) {
                match.textContent = value;
            } else {
                const proto = match instanceof HTMLTextAreaElement
                    ? HTMLTextAreaElement.prototype
                    : HTMLInputElement.prototype;
                const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
                if (setter) setter.call(match, value); else match.value = value;
            }
            for (const name of ['input', 'change']) {
                match.dispatchEvent(new Event(name, { bubbles: true }));
            }
            return 'restored';
        }
        return 'not-found';
    }"""
    for variant in _label_variants("收款单位"):
        for frame in _all_frames(page):
            try:
                result = frame.evaluate(script, {"label": variant, "value": value})
            except Exception:
                continue
            if result == "ok":
                return
            if result == "restored":
                print(f"[已还原] 收款单位：{value}", flush=True)
                return
    print("[警告] 收尾核对时未能定位到「收款单位」输入框", flush=True)


def _fill_payment_form(page: Page, data: PaymentData) -> None:
    """自动填单笔付款字段，不提交。

    顺序很关键：必须先填「收款单位」。ICBC 的收款单位字段会触发地址簿匹配，
    没匹配上时会把「收款账号 / 收款银行」联动清空。先把收款单位填好并主动
    Tab 触发清空逻辑（此时其它字段还没值，无影响），再补其它字段。"""
    print("[准备] 自动填入单笔付款数据", flush=True)
    _wait_full_load(page, timeout_ms=30000)

    _fill_payee_name_first(page, data.payee_name)
    _fill_text_only(page, "收款账号", data.payee_account)
    _fill_bank_and_click_first(page, data.payee_bank)
    _fill_text_only(page, "汇款金额", data.amount)
    _fill_text_only(page, "汇款用途", data.purpose)
    if data.remark:
        _fill_text_only(page, "备注", data.remark)

    # 走到这里，收款单位往往已被 autocomplete blur 清空，最后再用 DOM 写一遍。
    _ensure_payee_name_filled(page, data.payee_name)

    print("[完成] 已填入付款表单，请人工核对后再继续后续支付步骤。", flush=True)


def main() -> None:
    _setup_stdout_utf8()
    if not sys.stdout or not sys.stdout.isatty():
        _setup_file_logging()
    if _ITC_OK:
        backend = f"Interception 驱动（键盘 slot={_ITC_KB_SLOT}）"
    else:
        backend = "SendInput 回退（U盾 PIN 可能被拦截）"
    print(f"[输入后端] {backend}", flush=True)
    _disable_proxy_for_playwright()
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=[
                "--no-proxy-server",
                "--proxy-server=direct://",
                "--proxy-bypass-list=*",
            ],
        )
        context = browser.new_context(
            viewport={"width": 1366, "height": 860},
            locale="zh-CN",
        )
        page = context.new_page()
        page.goto(ICBC_URL, wait_until="domcontentloaded")

        print(f"[已打开] {ICBC_URL}", flush=True)

        corp_page = _click_corp_login(context, page)
        _click_ushield_login(context, corp_page)
        _click_payment_menu(corp_page)
        _open_single_payment_page(corp_page)
        _fill_payment_form(corp_page, PAYMENT_DATA)

        print("请在浏览器中核对自动填入的付款信息，并手动完成后续支付操作。", flush=True)
        print("关闭浏览器窗口（或在终端按 Ctrl+C）即可退出脚本。", flush=True)

        try:
            while browser.is_connected():
                time.sleep(1)
        except KeyboardInterrupt:
            pass
        finally:
            try:
                context.close()
            except Exception:
                pass
            try:
                browser.close()
            except Exception:
                pass


if __name__ == "__main__":
    main()
