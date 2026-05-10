# -*- coding: utf-8 -*-
"""
招行U-BANK - 制单：自动登录并进入单笔转账经办页面，填写转账表单
"""
import time
import subprocess
import sys
import threading
import win32gui
import win32api
import win32con
import win32clipboard
import os
import ctypes
from ctypes import wintypes
from datetime import datetime

# 自动把 招行/ 根目录加入 import 路径，方便直接 import ubank_common，无需配置 PYTHONPATH
_ZHIDAN_ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
if _ZHIDAN_ROOT not in sys.path:
    sys.path.insert(0, _ZHIDAN_ROOT)

from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from ubank_common import (
    click_at, press_keys, open_ubank, login_ubank,
    wait_for_main_window, click_control_by_name,
)

# 截图保存目录
_FINANCE_ROOT = os.path.dirname(_ZHIDAN_ROOT)
_SCREENSHOT_DIR = os.path.join(_ZHIDAN_ROOT, "screenshots")
os.makedirs(_SCREENSHOT_DIR, exist_ok=True)

_FORM_PATH = os.path.join(_FINANCE_ROOT, "M3直供合同付款数据获取", "bank_form.json")
_CODEX_HOME = os.environ.get("CODEX_HOME") or os.path.join(os.path.expanduser("~"), ".codex")
_USB_HUB_CTRL = os.path.join(_CODEX_HOME, "skills", "usb-hub", "scripts", "hub_ctrl.py")
_USB_HUB_PORT_BY_PAYER = {
    "示例付款单位A": 10,
    "示例付款单位B": 11,
}
UBANK_CRASH_EXIT_CODE = 8
SUBMIT_CLICK_EXIT_CODE = 9
CRASH_DIAG_MODE = os.environ.get("UBANK_CRASH_DIAG") == "1"
_EVENT_SYSTEM_DIALOGSTART = 0x0010
_EVENT_OBJECT_CREATE = 0x8000
_EVENT_OBJECT_SHOW = 0x8002
_EVENT_OBJECT_NAMECHANGE = 0x800C
_WINEVENT_OUTOFCONTEXT = 0x0000
_WINEVENT_SKIPOWNPROCESS = 0x0002
_WINEVENT_FLAGS = _WINEVENT_OUTOFCONTEXT | _WINEVENT_SKIPOWNPROCESS
_PM_REMOVE = 0x0001
TARGET_SCREEN_WIDTH = 1920
TARGET_SCREEN_HEIGHT = 1080
TARGET_SCALE = "100%"
_WinEventProcType = ctypes.WINFUNCTYPE(
    None,
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.HWND,
    wintypes.LONG,
    wintypes.LONG,
    wintypes.DWORD,
    wintypes.DWORD,
)


def _enable_dpi_awareness():
    """Use real screen pixels so UIA rectangles and fallback clicks match 100% DPI machines."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware.
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


_enable_dpi_awareness()


class UBankCrashDetected(RuntimeError):
    """Raised when the CMB U-BANK native client shows an application error."""


def _is_ubank_crash_title(title):
    title = title or ""
    is_app_error = "应用程序错误" in title or "Application Error" in title
    is_ubank = (
        "Firmbank.exe" in title
        or "招商银行企业银行" in title
        or "企业银行" in title
        or "U-BANK" in title
    )
    return is_app_error and is_ubank


def _find_ubank_crash_dialog():
    """Return the U-BANK crash dialog hwnd/title if Firmbank.exe has crashed."""
    matches = []

    def _enum_handler(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd) or ""
        except Exception:
            return

        if _is_ubank_crash_title(title):
            matches.append((hwnd, title))

    win32gui.EnumWindows(_enum_handler, None)
    return matches[0] if matches else (None, "")


def _dismiss_crash_dialog(hwnd):
    """Dismiss the specific crash dialog so the next retry is not blocked."""
    if not hwnd:
        return

    buttons = []

    def _enum_child(child_hwnd, _):
        try:
            text = (win32gui.GetWindowText(child_hwnd) or "").strip()
            if text in {"确定", "OK"}:
                buttons.append(child_hwnd)
        except Exception:
            pass

    try:
        win32gui.EnumChildWindows(hwnd, _enum_child, None)
        if buttons:
            left, top, right, bottom = win32gui.GetWindowRect(buttons[0])
            click_at((left + right) // 2, (top + bottom) // 2)
        else:
            win32gui.SetForegroundWindow(hwnd)
            press_keys((0x0D, 0), (0x0D, 2))
        time.sleep(0.3)
        if win32gui.IsWindow(hwnd):
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            time.sleep(0.2)
    except Exception as e:
        print(f"关闭招行崩溃弹窗失败: {e}")


def _clear_stale_ubank_crash_dialogs():
    """启动前清理上一次运行遗留的应用错误弹窗。"""
    for _ in range(3):
        hwnd, title = _find_ubank_crash_dialog()
        if not hwnd:
            return
        print(f"清理启动前残留的招行崩溃弹窗: {title}")
        _dismiss_crash_dialog(hwnd)
        time.sleep(0.5)


_GLOBAL_WATCHER = None  # 由 main() 启动的全程崩溃守护线程，模块级唯一实例


def _check_ubank_crash(stage="", dismiss=True):
    """检查 U-BANK 是否崩溃。

    全局守护线程 _GLOBAL_WATCHER 已在 20ms 轮询窗口；这里主要读它的 fired 标记。
    极少数场景（守护线程未启动 / 刚启动还没扫到）下，再做一次本地扫描兜底。
    """
    # 1) 守护线程已经检测并灭过弹窗 → 主流程直接 raise 让脚本以 rc=8 终止
    if _GLOBAL_WATCHER is not None and _GLOBAL_WATCHER.fired:
        title = _GLOBAL_WATCHER.fired_title
        where = f"（{stage}）" if stage else ""
        print(f"守护线程已捕获并强杀招行崩溃{where}: {title}")
        raise UBankCrashDetected(title)

    # 2) 兜底：自己再扫一次（守护线程未启动时这是唯一防线）
    hwnd, title = _find_ubank_crash_dialog()
    if not hwnd:
        return

    where = f"（{stage}）" if stage else ""
    print(f"检测到招行客户端崩溃{where}: {title}")
    if dismiss and not CRASH_DIAG_MODE:
        _kill_all_firmbank()
    elif CRASH_DIAG_MODE:
        print("诊断模式已开启：保留崩溃现场，等待 Windows 写入 LocalDump")
    raise UBankCrashDetected(title)


def _sleep_and_check(seconds, stage=""):
    end_ts = time.time() + seconds
    while time.time() < end_ts:
        _check_ubank_crash(stage)
        time.sleep(min(0.5, max(0, end_ts - time.time())))
    _check_ubank_crash(stage)


def _trace_fill_step(stage):
    ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    print(f"[填表TRACE {ts}] {stage}", flush=True)


def _log_display_profile():
    width = win32api.GetSystemMetrics(0)
    height = win32api.GetSystemMetrics(1)
    print(
        f"当前屏幕: {width}x{height}; 目标适配: "
        f"{TARGET_SCREEN_WIDTH}x{TARGET_SCREEN_HEIGHT}, 缩放 {TARGET_SCALE}"
    )
    if width != TARGET_SCREEN_WIDTH or height != TARGET_SCREEN_HEIGHT:
        print("提示: 坐标兜底按前台窗口比例计算，非目标分辨率也可尝试，但建议新电脑设为 1920x1080/100%")


def _is_debug_input_label(target_label):
    return "金额" in target_label or "用途" in target_label


def _load_bank_form():
    """读取转账表单 JSON；制单和 U 盾选择共用同一份数据。"""
    import json as _json

    try:
        with open(_FORM_PATH, "r", encoding="utf-8") as _f:
            form = _json.load(_f)
        print(f"已加载 bank_form.json: {_FORM_PATH}")
        return form, True
    except Exception as _e:
        print(f"未能加载 {_FORM_PATH}: {_e}，使用内置示例值")
        return {}, False


def _resolve_usb_hub_port(form):
    payer_name = (form.get("付款单位名称") or form.get("付款单位") or "").strip()
    if not payer_name:
        print("bank_form.json 缺少付款单位名称，无法选择 U 盾端口")
        return None

    for payer_keyword, port in _USB_HUB_PORT_BY_PAYER.items():
        if payer_keyword in payer_name:
            print(f"付款单位[{payer_name}]匹配[{payer_keyword}]，准备切换 USB Hub {port} 口")
            return port

    print(f"付款单位[{payer_name}]未配置 USB Hub 端口，终止制单以避免用错 U 盾")
    return None


def _switch_usb_hub_for_form(form, form_loaded):
    """U-BANK 启动前按付款单位独占打开对应 U 盾端口。"""
    if not form_loaded:
        print("未加载到 bank_form.json，跳过 USB Hub 切换，仅用于本地示例调试")
        return True

    port = _resolve_usb_hub_port(form)
    if port is None:
        return False

    if not os.path.exists(_USB_HUB_CTRL):
        print(f"未找到 USB Hub 控制脚本: {_USB_HUB_CTRL}")
        return False

    try:
        proc = subprocess.run(
            [sys.executable, _USB_HUB_CTRL, "only", str(port), "--settle", "2"],
            timeout=45,
        )
        print(f"USB Hub {port} 口切换退出码: {proc.returncode}")
        return proc.returncode == 0
    except subprocess.TimeoutExpired:
        print(f"USB Hub {port} 口切换超时")
        return False
    except Exception as e:
        print(f"USB Hub {port} 口切换异常: {e}")
        return False


def _screenshot(label):
    """截取前台窗口截图保存到本地"""
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return None
        ts = datetime.now().strftime("%H%M%S")
        filename = f"{ts}_{label}.png"
        filepath = os.path.join(_SCREENSHOT_DIR, filename)
        # 使用powershell截图（兼容性最好）
        cmd = f'''
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bmp = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
$bmp.Save("{filepath.replace("\\", "/")}")
$g.Dispose()
$bmp.Dispose()
Write-Output "OK"
'''
        result = subprocess.run(
            ["powershell", "-command", cmd],
            capture_output=True, text=True, encoding="utf-8", timeout=10
        )
        print(f"已截图: {filename}")
        return filepath
    except Exception as e:
        print(f"截图失败: {e}")
        return None


def _click_dropdown_option(control, target_text):
    """递归在下拉列表中查找并点击目标选项"""
    try:
        name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
        if name and name.strip() == target_text.strip() and control.is_visible():
            rect = control.rectangle()
            if rect.width() > 5 and rect.height() > 5:
                click_at(rect.mid_point().x, rect.mid_point().y)
                print(f"已选择下拉项: {target_text}")
                return True
    except Exception:
        pass
    try:
        for child in control.children():
            if _click_dropdown_option(child, target_text):
                return True
    except Exception:
        pass
    return False


def find_and_input_text(main_win, target_label, text_value, input_mode="type"):
    """根据标签名模糊查找相邻的文本输入框并输入内容"""
    debug_input = _is_debug_input_label(target_label)
    use_clipboard = input_mode == "paste"

    def _trace(step):
        if debug_input:
            _trace_fill_step(f"{target_label}: {step}")

    def _check(step):
        if debug_input:
            _check_ubank_crash(f"fill-debug: {target_label} {step}")

    # 先用模糊匹配查找标签控件（标签名中包含目标关键字即可）
    label_control = None

    def _find_label(control):
        nonlocal label_control
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            # 模糊匹配：目标关键字在标签名中，或标签名在目标关键字中
            if name and name.strip() and (
                target_label.strip() in name or
                name.strip() in target_label.strip() or
                target_label.replace(" ", "")[:3] in name.replace(" ", "")
            ):
                label_control = control
                return True
        except Exception:
            pass
        try:
            for child in control.children():
                if _find_label(child):
                    return True
        except Exception:
            pass
        return False

    print(f"正在查找标签: {target_label}")
    _trace("开始查找标签")
    _find_label(main_win)
    _check("标签查找后")

    # 如果模糊匹配没找到，尝试遍历所有可见文本控件打印调试信息
    if not label_control:
        all_labels = []

        def _collect_texts(control):
            try:
                name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
                ctrl_type = control.element_info.control_type if hasattr(control.element_info, "control_type") else ""
                cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
                if name and name.strip() and ("Text" in ctrl_type or "Static" in cls):
                    all_labels.append(f"[{ctrl_type}|{cls}] {name.strip()}")
            except Exception:
                pass
            try:
                for child in control.children():
                    _collect_texts(child)
            except Exception:
                pass

        _collect_texts(main_win)
        print(f"  未精确找到: {target_label}，当前页面可见文本控件:")
        for lbl in all_labels[:50]:
            print(f"  {lbl}")

    if not label_control:
        print(f"未找到标签: {target_label}")
        return False

    try:
        label_rect = label_control.rectangle()
        _trace(
            f"命中标签 rect=({label_rect.left},{label_rect.top},"
            f"{label_rect.right},{label_rect.bottom})"
        )
    except Exception:
        _trace("命中标签但读取 rect 失败")

    # 找到标签后，在其父容器中找可编辑的输入框
    def _find_nearby_edit(label_ctrl):
        """从标签控件的父容器中找到可编辑的输入框"""
        candidates = []
        try:
            parent = label_ctrl.parent()
            if parent is None:
                return None
            # 获取标签的位置信息
            label_rect = label_ctrl.rectangle()
            label_y = label_rect.top

            for sibling in parent.children():
                if not sibling.is_visible():
                    continue
                cls_name = sibling.element_info.class_name if hasattr(sibling.element_info, "class_name") else ""
                ctrl_type = sibling.element_info.control_type if hasattr(sibling.element_info, "control_type") else ""
                sname = sibling.element_info.name if hasattr(sibling.element_info, "name") and sibling.element_info.name else ""

                # 匹配编辑类控件
                is_edit = any(kw in cls_name for kw in ["Edit", "edit"]) or any(kw in str(ctrl_type) for kw in ["Edit", "Document", "ComboBox"])
                if is_edit:
                    srect = sibling.rectangle()
                    # 优先选择与标签在同一行或附近（Y坐标接近）且是空输入框的
                    is_empty = not sname or len(sname.strip()) < 2 or "输入" in sname or "请输" in sname or "选择" in sname
                    y_diff = abs(srect.top - label_y)
                    candidates.append((sibling, y_diff, is_empty))
        except Exception:
            pass

        if not candidates:
            return None
        # 排序：优先选Y距离近的、空框的
        candidates.sort(key=lambda x: (x[1], not x[2]))
        return candidates[0][0]

    edit_ctrl = _find_nearby_edit(label_control)
    _check("输入框查找后")
    if edit_ctrl:
        try:
            pre_focus_rect = None
            try:
                pre_focus_rect = edit_ctrl.rectangle()
                _trace(
                    f"命中输入框 rect=({pre_focus_rect.left},{pre_focus_rect.top},"
                    f"{pre_focus_rect.right},{pre_focus_rect.bottom})"
                )
            except Exception:
                _trace("命中输入框但读取 rect 失败")

            _trace("set_focus 前")
            edit_ctrl.set_focus()
            time.sleep(0.3)  # 加长等待确保聚焦完成
            _trace("set_focus 后")
            _check("set_focus 后")
            # 判断控件类型，选择合适的输入方式
            ctrl_type = edit_ctrl.element_info.control_type if hasattr(edit_ctrl.element_info, "control_type") else ""
            cls_name = edit_ctrl.element_info.class_name if hasattr(edit_ctrl.element_info, "class_name") else ""
            _trace(f"控件类型 type={ctrl_type}, class={cls_name}")

            # 下拉框/组合框
            is_combo = "Combo" in ctrl_type or "Combo" in cls_name or "ComboBox" in str(ctrl_type)

            if is_combo:
                rect = edit_ctrl.rectangle()
                rect_bad = (
                    rect.width() <= 5
                    or rect.height() <= 5
                    or (rect.mid_point().x == 0 and rect.mid_point().y == 0)
                )
                if rect_bad:
                    # 招行 ComboBox set_focus 后 UIA rect 偶发塌陷为 (0,0,0,0)，
                    # 回退到 set_focus 前缓存的 rect 再点。
                    if (
                        pre_focus_rect is not None
                        and pre_focus_rect.width() > 5
                        and pre_focus_rect.height() > 5
                        and not (pre_focus_rect.mid_point().x == 0 and pre_focus_rect.mid_point().y == 0)
                    ):
                        _trace(
                            f"Combo rect 塌陷，回退 set_focus 前 rect=({pre_focus_rect.left},"
                            f"{pre_focus_rect.top},{pre_focus_rect.right},{pre_focus_rect.bottom})"
                        )
                        rect = pre_focus_rect
                    else:
                        _trace(
                            f"Combo rect 不可信，放弃 UIA 输入 rect=({rect.left},{rect.top},"
                            f"{rect.right},{rect.bottom})"
                        )
                        return False
                _trace(f"Combo 点击前 center=({rect.mid_point().x},{rect.mid_point().y})")
                click_at(rect.mid_point().x, rect.mid_point().y)
                time.sleep(0.5)
                _trace("Combo 点击后")
                _check("Combo 点击后")

            if text_value:
                if use_clipboard:
                    _trace("跳过清空字段，剪贴板粘贴前")
                    if not _paste_text_via_clipboard(text_value):
                        return False
                    _trace("剪贴板粘贴后")
                    _check("剪贴板粘贴后")
                else:
                    _trace("跳过清空字段，直接输入")
                    _trace("慢速键入前")
                    _type_text_slow(text_value, char_delay=0.09 if "金额" not in target_label else 0.12)
                    time.sleep(0.4)
                    _trace("慢速键入后")
                    _check("慢速键入后")
            else:
                _trace("空值跳过清空字段")

            action = "粘贴" if use_clipboard else "输入"
            print(f"已{action} [{target_label}]: {text_value}")
            _trace("输入函数即将返回 True")
            _check("输入函数返回前")
            return True
        except UBankCrashDetected:
            raise
        except Exception as e:
            # 打印详细错误和控件信息用于调试
            try:
                c_type = edit_ctrl.element_info.control_type if hasattr(edit_ctrl.element_info, "control_type") else ""
                c_cls = edit_ctrl.element_info.class_name if hasattr(edit_ctrl.element_info, "class_name") else ""
                print(f"  控件详情: type={c_type}, class={c_cls}")
            except Exception:
                pass
            print(f"输入失败 [{target_label}]: {e}")
    else:
        print(f"未找到 {target_label} 对应的输入框")

    return False


def select_radio_by_name(main_win, group_name, option_name):
    """根据选项名称点击对应的单选按钮"""

    def _find_and_click_radio(control):
        """递归查找包含目标文字的单选按钮并点击"""
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            if name and option_name in name:
                ctrl_type = control.element_info.control_type if hasattr(control.element_info, "control_type") else ""
                rect = control.rectangle()

                if control.is_visible() and rect.width() > 5 and rect.height() > 5:
                    # RadioButton 或 RadioButton类型的控件
                    if "Radio" in ctrl_type or "RadioButton" in str(ctrl_type):
                        click_at(rect.mid_point().x, rect.mid_point().y)
                        print(f"已选中: {option_name}")
                        return True
                    # 也可能是普通文本控件，直接点击
                    else:
                        click_at(rect.mid_point().x, rect.mid_point().y)
                        print(f"已选中(文本点击): {option_name}")
                        return True
        except Exception:
            pass

        try:
            for child in control.children():
                result = _find_and_click_radio(child)
                if result:
                    return True
        except Exception:
            pass
        return False

    print(f"正在选择: {group_name} -> {option_name}")
    _find_and_click_radio(main_win)


def _uia_collect_confirm_candidates(control, out_list, host_rect):
    """递归收集名称含「确认/确定」的可点击控件候选（优先小面积按钮类）"""
    try:
        if not control.is_visible():
            return
        ei = control.element_info
        name = ei.name if hasattr(ei, "name") and ei.name else ""
        nm = name.replace(" ", "") if name else ""
        ctype = str(ei.control_type) if hasattr(ei, "control_type") and ei.control_type else ""
        cls = ei.class_name if hasattr(ei, "class_name") and ei.class_name else ""
        rect = control.rectangle()
        w, h = rect.width(), rect.height()
        if w >= 6 and h >= 6 and ("确认" in nm or "确定" in nm):
            area = w * h
            is_std_btn = ("Button" in ctype) or ("SplitButton" in ctype) or ("Button" in cls)
            # 面积过大且非标准按钮时，更像容器文案，只继续扫子节点
            if area > 20000 and not is_std_btn:
                for ch in control.children():
                    _uia_collect_confirm_candidates(ch, out_list)
                return
            pri_name = 0 if "确认" in nm else 1
            pri_type = 0 if is_std_btn else (1 if ("Hyperlink" in ctype or "Custom" in ctype) else 2)
            mx, my = rect.mid_point().x, rect.mid_point().y
            # 保护：不点击窗口右上角附近（避免误点关闭“X”）
            hr_left, hr_top, hr_right, hr_bottom = host_rect
            close_zone_x = hr_right - 90
            close_zone_y = hr_top + 50
            if mx >= close_zone_x and my <= close_zone_y:
                return
            out_list.append((pri_name, pri_type, area, mx, my, name.strip()[:48]))
    except Exception:
        pass
    try:
        for ch in control.children():
            _uia_collect_confirm_candidates(ch, out_list, host_rect)
    except Exception:
        pass


def _rect_close(a, b, tol=30):
    """判断两个矩形是否近似相同（用于匹配前台窗口）"""
    return (
        abs(a[0] - b[0]) <= tol and
        abs(a[1] - b[1]) <= tol and
        abs(a[2] - b[2]) <= tol and
        abs(a[3] - b[3]) <= tol
    )


def _try_uia_confirm_click(desktop):
    """仅在前台窗口的 UIA 树中查找并点击「确认/确定」"""
    fg_hwnd = win32gui.GetForegroundWindow()
    if not fg_hwnd:
        return False
    fg_rect = win32gui.GetWindowRect(fg_hwnd)
    candidates = []
    for w in desktop.windows():
        try:
            wr = w.rectangle()
            host_rect = (wr.left, wr.top, wr.right, wr.bottom)
            # 只在前台窗口中查找，避免命中主页面里的“确认”按钮
            if not _rect_close(host_rect, fg_rect):
                continue
            _uia_collect_confirm_candidates(w, candidates, host_rect)
        except Exception:
            pass
    if not candidates:
        return False
    candidates.sort(key=lambda t: (t[0], t[1], t[2]))
    _, _, _, x, y, label = candidates[0]
    print(f"UIA 定位到确认类控件: {label!r} -> ({x}, {y})")
    click_at(x, y)
    time.sleep(0.35)
    return True


def _click_foreground_client(cx_ratio, cy_ratio):
    """在前台窗口客户区按比例点击，用于抢焦点（比例 0~1，相对窗口宽高）"""
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return False
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w = right - left
    h = bottom - top
    if w < 80 or h < 60:
        return False
    x = int(left + w * cx_ratio)
    y = int(top + h * cy_ratio)
    click_at(x, y)
    time.sleep(0.2)
    return True


def _wheel_foreground_client(cx_ratio, cy_ratio, clicks):
    """在前台窗口客户区按比例滚轮滚动；不点击任何控件。clicks<0 向下。"""
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return False
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w = right - left
    h = bottom - top
    if w < 80 or h < 60:
        return False
    x = int(left + w * cx_ratio)
    y = int(top + h * cy_ratio)
    win32api.SetCursorPos((x, y))
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, int(clicks * 120), 0)
    time.sleep(0.2)
    return True


def _type_ascii_text_slow(text_value, char_delay=0.12):
    """逐字符键入 ASCII 字段，避免触发网页控件的 paste/onchange 快路径。"""
    key_map = {
        ".": 0xBE,
        ",": 0xBC,
        "-": 0xBD,
        " ": 0x20,
    }
    for ch in text_value or "":
        if "0" <= ch <= "9":
            vk = ord(ch)
        elif ch in key_map:
            vk = key_map[ch]
        else:
            raise ValueError(f"慢速键入只支持 ASCII 数字金额字符，遇到: {ch!r}")
        press_keys((vk, 0), (vk, 2))
        time.sleep(char_delay)


def _type_text_slow(text_value, char_delay=0.09):
    """逐字符键入文本；中文通过 VK_PACKET 发送，不走剪贴板。"""
    text = text_value or ""
    if text and all(("0" <= ch <= "9") or ch in ".,- " for ch in text):
        _type_ascii_text_slow(text, char_delay=char_delay)
        return

    for ch in text:
        send_keys(ch, pause=0, with_spaces=True, vk_packet=True)
        time.sleep(char_delay)


def _read_clipboard_text():
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                return win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT), True
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass
    return "", False


def _write_clipboard_text(text_value):
    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text_value or "")
        finally:
            win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        print(f"写入剪贴板失败: {e}")
        return False


def _paste_text_via_clipboard(text_value):
    """设置剪贴板后发送 Ctrl+V；不做 Ctrl+A/Backspace 清空。"""
    old_text, has_old_text = _read_clipboard_text()
    if not _write_clipboard_text(text_value):
        return False
    try:
        time.sleep(0.12)
        press_keys((0x11, 0), (0x56, 0), (0x56, 2), (0x11, 2))  # Ctrl+V
        time.sleep(0.45)
        return True
    finally:
        if has_old_text:
            _write_clipboard_text(old_text)


def _click_ratio_and_paste(cx_ratio, cy_ratio, text_value, field_name, input_mode="type"):
    """U-BANK 页面 UIA 标签不稳定时，按窗口比例点击输入框并输入/粘贴。"""
    debug_input = _is_debug_input_label(field_name)
    use_clipboard = input_mode == "paste"
    if debug_input:
        _trace_fill_step(f"{field_name}: 坐标兜底点击前 ratio=({cx_ratio},{cy_ratio})")
    if not _click_foreground_client(cx_ratio, cy_ratio):
        print(f"{field_name} 坐标兜底失败：未找到前台窗口")
        return False
    if debug_input:
        action = "粘贴" if use_clipboard else "键入"
        _trace_fill_step(f"{field_name}: 坐标兜底点击后，{action}前")
        _check_ubank_crash(f"fill-debug: {field_name} 坐标兜底点击后")
    # U-BANK 的 Chrome/原生桥接对 Ctrl+A + Backspace 很敏感，曾在金额框清空后崩溃。
    # 坐标兜底也只做聚焦后输入/粘贴，不清空。
    time.sleep(0.15)
    if use_clipboard:
        if not _paste_text_via_clipboard(text_value):
            return False
    else:
        _type_text_slow(text_value, char_delay=0.09 if "金额" not in field_name else 0.12)
    time.sleep(0.4)
    if debug_input:
        _trace_fill_step(f"{field_name}: 坐标兜底输入后")
        _check_ubank_crash(f"fill-debug: {field_name} 坐标兜底输入后")
    print(f"已通过坐标兜底输入 [{field_name}]: {text_value or ''}")
    return True


def _dismiss_transfer_overlays():
    """关闭小招提醒/引导遮罩，避免表单变灰导致控件不可定位。"""
    print("清理页面遮罩/提醒...")
    for _ in range(3):
        press_keys((0x1B, 0), (0x1B, 2))  # Esc
        time.sleep(0.25)
    for x_ratio, y_ratio in [(0.285, 0.700), (0.960, 0.515), (0.960, 0.675)]:
        _click_foreground_client(x_ratio, y_ratio)
        time.sleep(0.25)


def _try_spatial_confirm_click():
    """自绘/Web 弹窗常见布局：主按钮在底部偏右，对前台窗体该位置点击"""
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w = right - left
    h = bottom - top
    if w < 200 or h < 100:
        return
    x = right - int(max(72, w * 0.20))
    y = bottom - int(max(32, h * 0.14))
    print(f"几何点击确认区域（前台窗右下角附近）: ({x}, {y})，窗口约 {w}x{h}")
    click_at(x, y)
    time.sleep(0.3)


def _find_x_close_button(main_win):
    """在主窗口 UIA 树里找右上角的关闭(X)按钮，返回控件或 None。

    招行 U-BANK 12.0.0.6 标题栏是自绘的，但 UIA 一般还是能枚举到 Button 类控件。
    匹配条件：按钮类、宽高合理（小图标按钮）、位置在主窗口右上角带内。
    """
    try:
        win_rect = main_win.rectangle()
    except Exception as e:
        print(f"获取主窗口矩形失败: {e}")
        return None

    win_right = win_rect.right
    win_top = win_rect.top
    candidates = []

    def _walk(control):
        try:
            if not control.is_visible():
                return
        except Exception:
            return
        try:
            ei = control.element_info
            name = (ei.name or "") if hasattr(ei, "name") else ""
            ctype = str(ei.control_type) if hasattr(ei, "control_type") and ei.control_type else ""
            cls = (ei.class_name or "") if hasattr(ei, "class_name") else ""
            rect = control.rectangle()
            w, h = rect.width(), rect.height()

            # 关闭按钮特征：尺寸像图标按钮（10~60px 宽高），且贴在主窗右上角
            in_top_right = (
                rect.right >= win_right - 80
                and rect.top <= win_top + 60
                and rect.right <= win_right + 5
            )
            is_button_like = "Button" in ctype or "Button" in cls or "Close" in cls
            is_close_named = any(kw in name for kw in ["关闭", "Close", "close", "×", "X"])
            small_icon = 8 <= w <= 80 and 8 <= h <= 60

            if in_top_right and small_icon and (is_button_like or is_close_named):
                # 评分：名字直接命中"关闭/Close"优先；越靠右上角越优先
                pri_name = 0 if is_close_named else 1
                pri_pos = (win_right - rect.right) + (rect.top - win_top)
                candidates.append((pri_name, pri_pos, rect, name, ctype, cls))
        except Exception:
            pass
        try:
            for child in control.children():
                _walk(child)
        except Exception:
            pass

    _walk(main_win)
    if not candidates:
        return None
    candidates.sort(key=lambda t: (t[0], t[1]))
    pri_name, pri_pos, rect, name, ctype, cls = candidates[0]
    print(f"UIA 命中 X 按钮候选: name={name!r}, type={ctype}, class={cls}, "
          f"rect=({rect.left},{rect.top},{rect.right},{rect.bottom})")
    return rect


def _kill_all_firmbank():
    """强杀所有 Firmbank.exe 进程，连同子进程一起。返回是否成功。

    注意：不解码 stdout/stderr —— Windows taskkill 输出是 GBK，但脚本运行在
    `python -X utf8` 下默认按 UTF-8 解码会触发 UnicodeDecodeError。这里只关心
    退出码，bytes 直接丢弃。
    """
    try:
        proc = subprocess.run(
            ["taskkill", "/F", "/T", "/IM", "Firmbank.exe"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=10,
        )
        return proc.returncode == 0
    except Exception as e:
        print(f"taskkill Firmbank 异常: {e}")
        return False


def _is_firmbank_alive():
    """通过 EnumWindows 判断 Firmbank 主窗口是否还活着（比 tasklist 快）。"""
    alive = [False]

    def _enum(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd) or ""
        except Exception:
            return
        if "招商银行企业银行" in title or "Firmbank" in title:
            alive[0] = True

    try:
        win32gui.EnumWindows(_enum, None)
    except Exception:
        pass
    return alive[0]


class _CrashWatcher:
    """后台守护：发现 Firmbank 崩溃弹窗立刻 taskkill /F /T，
    尽量压缩弹窗可见时间。

    Why: U-BANK 12.0.0.6 在转账经办页关闭路径有 use-after-free，无法治本，
    只能在弹窗刚出现的时候立刻强杀整个进程树，让弹窗连同进程一起消失。
    优先用 SetWinEventHook 监听窗口创建/显示/标题变化，再用 20ms 轮询兜底。
    """

    def __init__(self, poll_interval=0.02):
        self._stop = threading.Event()
        self._fired = threading.Event()
        self._interval = poll_interval
        self._thread = None
        self._fired_title = ""
        self._fire_lock = threading.Lock()
        self._user32 = None
        self._hooks = []
        self._win_event_proc = None

    def start(self):
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _loop(self):
        self._install_win_event_hooks()
        try:
            while not self._stop.is_set() and not self.fired:
                self._pump_win_messages()
                if self.fired:
                    return

                hwnd, title = _find_ubank_crash_dialog()
                if hwnd:
                    self._fire(title)
                    return

                self._stop.wait(0.005 if self._hooks else self._interval)
        finally:
            self._uninstall_win_event_hooks()

    def _install_win_event_hooks(self):
        try:
            self._user32 = ctypes.WinDLL("user32", use_last_error=True)
            self._win_event_proc = _WinEventProcType(self._on_win_event)
            self._user32.SetWinEventHook.argtypes = [
                wintypes.DWORD,
                wintypes.DWORD,
                wintypes.HANDLE,
                _WinEventProcType,
                wintypes.DWORD,
                wintypes.DWORD,
                wintypes.DWORD,
            ]
            self._user32.SetWinEventHook.restype = wintypes.HANDLE
            self._user32.UnhookWinEvent.argtypes = [wintypes.HANDLE]
            self._user32.UnhookWinEvent.restype = wintypes.BOOL
            self._user32.PeekMessageW.argtypes = [
                ctypes.POINTER(wintypes.MSG),
                wintypes.HWND,
                wintypes.UINT,
                wintypes.UINT,
                wintypes.UINT,
            ]
            self._user32.PeekMessageW.restype = wintypes.BOOL
            self._user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
            self._user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]

            for event_id in (
                _EVENT_SYSTEM_DIALOGSTART,
                _EVENT_OBJECT_CREATE,
                _EVENT_OBJECT_SHOW,
                _EVENT_OBJECT_NAMECHANGE,
            ):
                hook = self._user32.SetWinEventHook(
                    event_id,
                    event_id,
                    None,
                    self._win_event_proc,
                    0,
                    0,
                    _WINEVENT_FLAGS,
                )
                if hook:
                    self._hooks.append(hook)

            if self._hooks:
                print("崩溃守护已安装 WinEventHook（窗口创建/显示/标题变化）")
            else:
                err = ctypes.get_last_error()
                print(f"WinEventHook 安装失败，退回 20ms 轮询: {err}")
        except Exception as e:
            self._hooks = []
            self._win_event_proc = None
            print(f"WinEventHook 安装异常，退回 20ms 轮询: {e}")

    def _uninstall_win_event_hooks(self):
        if not self._user32:
            return
        for hook in self._hooks:
            try:
                self._user32.UnhookWinEvent(hook)
            except Exception:
                pass
        self._hooks = []

    def _pump_win_messages(self):
        if not self._user32:
            return
        try:
            msg = wintypes.MSG()
            while self._user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, _PM_REMOVE):
                self._user32.TranslateMessage(ctypes.byref(msg))
                self._user32.DispatchMessageW(ctypes.byref(msg))
        except Exception:
            pass

    def _on_win_event(self, _hook, _event, hwnd, _id_object, _id_child, _thread, _time_ms):
        if self._stop.is_set() or self.fired or not hwnd:
            return
        self._handle_possible_dialog(hwnd)

    def _handle_possible_dialog(self, hwnd):
        try:
            title = win32gui.GetWindowText(hwnd) or ""
        except Exception:
            return False
        if not _is_ubank_crash_title(title):
            return False
        self._fire(title)
        return True

    def _fire(self, title):
        with self._fire_lock:
            if self._fired.is_set():
                return False
            self._fired_title = title
            self._fired.set()
        if CRASH_DIAG_MODE:
            print("诊断模式已开启：崩溃守护不强杀 Firmbank，等待 Windows 写入 LocalDump")
        else:
            _kill_all_firmbank()
        return True

    def stop(self, timeout=1.0):
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=timeout)

    @property
    def fired(self):
        return self._fired.is_set()

    @property
    def fired_title(self):
        return self._fired_title


def _close_with_confirm(main_win):
    """关闭 U-BANK：直接 taskkill /F /T 整个进程树。

    Why: Firmbank 12.0.0.6 在转账经办页接到 WM_CLOSE 必触发内部 COM 对象的
    use-after-free，弹出错误对话框给用户看。任何"优雅关闭"路径（Alt+F4、点 X、
    导航离开页面）都会触发同一条 cleanup 链，无法避免。守护线程式高频轮询
    强杀的速度也跟不上 Windows 的弹窗渲染速度，用户依然能看到弹窗。

    所以彻底改思路：根本不发 WM_CLOSE，直接对 Firmbank 进程树 taskkill /F /T。
    Windows 直接 TerminateProcess，不走 Firmbank 自己的 cleanup → bug 不被触发
    → 弹窗根本不出现。

    CLAUDE.md 已明确：制单成功路径已经点击「经办」按钮；关闭只是收尾清理。
    所以"优雅关闭"对监控流程而言并非必要功能，taskkill 是合规的关闭手段。
    """
    _ = main_win  # 不再使用，保留参数以维持调用方签名兼容
    print("开始关闭程序...")

    if not _is_firmbank_alive():
        print("Firmbank 已不在运行，无需关闭")
        return

    # 直接强杀整个进程树
    print("发送 taskkill /F /T /IM Firmbank.exe")
    _kill_all_firmbank()

    # 等进程树清干净（Windows 杀子进程有秒级延迟）
    end_ts = time.time() + 6
    while time.time() < end_ts:
        if not _is_firmbank_alive():
            break
        time.sleep(0.2)

    # 仍有残留 → 再发一次（极少数子进程可能逗留较久）
    if _is_firmbank_alive():
        time.sleep(0.5)
        _kill_all_firmbank()
        time.sleep(1)

    if _is_firmbank_alive():
        print("Firmbank 仍未关闭，请人工检查")
    else:
        print("Firmbank 已关闭")


def _click_confirm_button(desktop, button_text, step_desc):
    """
    退出确认弹窗的多策略顺序（嵌入 Web 时 UIA 往往不完整）：
    1) 轮询 UIA 树，找名称含「确认/确定」的按钮类控件并点击
    2) 前台窗体中心点一下抢焦点，再点右下角常见主按钮区，再发 Enter（默认按钮）
    3) 多按几次 Tab 后 Enter，覆盖焦点顺序与控件数量变化的情况
    """
    _ = button_text  # 保留参数与旧调用一致，逻辑已不依赖单一文案匹配 Text
    _sleep_and_check(1.0, f"{step_desc}确认前")

    # 持续监听一小段时间，避免“弹窗刚出现时还未完成 UIA 树注册”
    end_ts = time.time() + 12
    while time.time() < end_ts:
        _check_ubank_crash(f"{step_desc}确认监听")
        if _try_uia_confirm_click(desktop):
            print(f"{step_desc} 已通过 UIA 点击确认")
            _check_ubank_crash(f"{step_desc}确认后")
            return

        # UIA 未命中时，按“前台小弹窗几何主按钮”点击（已实测更稳定）
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            # 仅对中小弹窗启用，避免误点主界面
            if 320 <= width <= 1200 and 160 <= height <= 700:
                # 关闭确认弹窗中，蓝色“确认”稳定在底部偏中右
                x = int(left + width * 0.56)
                y = int(top + height * 0.84)
                print(f"{step_desc} UIA 未命中，几何点击确认: ({x}, {y}), size=({width}x{height})")
                click_at(x, y)
                _sleep_and_check(0.2, f"{step_desc}几何点击后")
                press_keys((0x0D, 0), (0x0D, 2))
                _check_ubank_crash(f"{step_desc}回车后")
                return

        _sleep_and_check(0.25, f"{step_desc}确认轮询")

    # 最终兜底：仅发送键盘确认，不做额外危险坐标点击
    print(f"{step_desc} 监听超时，改用 Enter + Tab 兜底")
    press_keys((0x0D, 0), (0x0D, 2))
    _sleep_and_check(0.3, f"{step_desc}兜底回车后")
    for _ in range(14):
        press_keys((0x09, 0), (0x09, 2))
        _sleep_and_check(0.12, f"{step_desc}兜底Tab")
    press_keys((0x0D, 0), (0x0D, 2))
    _check_ubank_crash(f"{step_desc}兜底确认后")
    print(f"{step_desc} 兜底确认序列已发送")


def _click_bank_dropdown_option(main_win):
    """输入开户银行名称后，等待下拉列表出现，点击银行选项确认"""

    def _find_bank_items(control, out_list):
        """递归查找开户银行下拉列表中的选项项"""
        try:
            if not control.is_visible():
                return
        except Exception:
            return
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            rect = control.rectangle()
            # 匹配单行高度的银行名称选项（排除容器聚合文本）
            name_len = len(name.strip()) if name else 0
            is_single_item = 2 <= name_len <= 20
            is_single_row = 15 <= rect.height() <= 55 and rect.width() > 60
            if name and is_single_item and is_single_row and (
                "招商" in name or "中国银行" in name or "工商" in name or
                "建设" in name or "农业" in name or "交通" in name or
                "中信" in name or "浦发" in name or "兴业" in name
            ):
                out_list.append((rect, name))
        except Exception:
            pass
        try:
            for child in control.children():
                _find_bank_items(child, out_list)
        except Exception:
            pass

    print("等待开户银行下拉列表...")
    end_time = time.time() + 6
    while time.time() < end_time:
        items = []
        _find_bank_items(main_win, items)
        if items:
            # 按Y坐标排序，取第一个（最上面的选项）
            items.sort(key=lambda x: x[0].top)
            first_rect, first_name = items[0]
            mx = first_rect.mid_point().x
            my = first_rect.mid_point().y
            print(f"找到银行选项: {first_name}，位置({mx}, {my})")
            click_at(mx, my)
            print(f"已选择: {first_name}")
            return True
        time.sleep(0.3)

    print("未检测到开户银行下拉列表选项")
    return False


def _double_click_first_branch_option(main_win):
    """输入支行名称后，等待下拉列表出现，双击第一个选项（定位到'行'字附近）"""

    def _find_branch_dropdown_items(control, out_list):
        """递归查找支行下拉列表中的单独选项项"""
        try:
            if not control.is_visible():
                return
        except Exception:
            return
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
            rect = control.rectangle()
            # 匹配单行高度的"支行"文本选项（排除容器聚合文本）
            name_len = len(name.strip()) if name else 0
            # 单个支行名称通常 10~40 字符；容器会包含大量文字（>100字符）
            is_single_item = 5 <= name_len <= 60
            is_single_row = 15 <= rect.height() <= 55 and rect.width() > 80
            if name and "支行" in name and is_single_item and is_single_row:
                out_list.append((rect, name, ctrl_type))
        except Exception:
            pass
        try:
            for child in control.children():
                _find_branch_dropdown_items(child, out_list)
        except Exception:
            pass

    print("等待支行下拉列表...")
    end_time = time.time() + 8
    while time.time() < end_time:
        items = []
        _find_branch_dropdown_items(main_win, items)
        if items:
            # 按Y坐标排序，取第一个（最上面的选项）
            items.sort(key=lambda x: x[0].top)
            first_rect, first_name, first_type = items[0]
            mx = first_rect.mid_point().x + 60  # 偏右，靠近"行"字位置
            my = first_rect.mid_point().y
            print(f"找到支行选项列表，第一项[{first_type}]: {first_name}，位置({mx}, {my})")

            # 双击：第一次点击
            click_at(mx, my)
            time.sleep(0.15)
            # 第二次点击
            click_at(mx, my)
            print(f"已双击选择: {first_name}")
            return True
        time.sleep(0.3)

    print("未检测到支行下拉列表选项")
    return False


def _input_bank_name(main_win, bank_name):
    """专门处理开户银行输入：复用通用查找 -> 逐字键入 -> 点击下拉选项"""
    print(f"开始填写开户银行: {bank_name}")

    # 截图：填写前
    _screenshot("01_开户银行填前")

    # 1. 用通用函数查找并聚焦输入框
    find_and_input_text(main_win, "开户银行", "")
    time.sleep(0.5)

    # 2. 逐字键入银行名称，不走清空/剪贴板路径。
    _type_text_slow(bank_name, char_delay=0.1)
    print(f"已输入开户银行: {bank_name}")
    time.sleep(1.5)

    # 截图：填入后
    _screenshot("02_开户银行填后")

    # 4. 等待下拉并点击银行选项（而不是直接回车）
    _click_bank_option(main_win, bank_name)
    time.sleep(0.5)

    # 5. Tab 切走焦点
    press_keys((0x09, 0), (0x09, 2))
    time.sleep(0.3)

    print(f"开户银行填写完成")
    return True


def _click_bank_option(main_win, bank_name):
    """在开户银行下拉列表中找到并点击银行选项"""

    def _find_bank_items(control, out_list):
        try:
            if not control.is_visible():
                return
        except Exception:
            return
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            rect = control.rectangle()
            # 银行名较短(2~25字)，排除输入框本身
            is_input = any(kw in ctrl_type for kw in ["Edit", "Document"])
            name_len = len(name.strip()) if name else 0
            is_single_row = 15 <= rect.height() <= 55 and rect.width() > 60
            if not is_input and name and 2 <= name_len <= 25 and is_single_row and bank_name[:3] in name:
                out_list.append((rect, name))
        except Exception:
            pass
        try:
            for child in control.children():
                _find_bank_items(child, out_list)
        except Exception:
            pass

    end_time = time.time() + 6
    while time.time() < end_time:
        items = []
        _find_bank_items(main_win, items)
        if items:
            items.sort(key=lambda x: x[0].top)
            r, n = items[0]
            click_at(r.mid_point().x, r.mid_point().y)
            print(f"已选择开户银行下拉项: {n}")
            return True
        time.sleep(0.3)

    print("未找到开户银行下拉项，尝试回车")
    press_keys((0x0D, 0), (0x0D, 2))


def _check_skip_branch(main_win):
    """检测支行输入框是否显示'无需填写'提示或被禁用，若是则跳过"""

    def _find_branch_input(control):
        """查找支行名称/联行号输入框"""
        nonlocal branch_edit
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
            is_edit = any(kw in cls for kw in ["Edit", "edit"]) or any(kw in ctrl_type for kw in ["Edit", "Document"])
            if is_edit and ("联行号" in name or "支行名称" in name) and control.is_visible():
                rect = control.rectangle()
                if rect.width() > 50 and rect.height() > 5:
                    branch_edit = control
                    return True
        except Exception:
            pass
        try:
            for child in control.children():
                if _find_branch_input(child):
                    return True
        except Exception:
            pass
        return False

    def _scan_nearby_hints(control, rect_hint):
        """扫描支行输入框附近区域的所有可见文本，查找'无需填写'等提示"""
        try:
            if not control.is_visible():
                return
        except Exception:
            return
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
            # 只看非编辑类控件（Text、Static、Hyperlink等）
            is_edit = any(kw in cls for kw in ["Edit", "edit"]) or any(kw in ctrl_type for kw in ["Edit", "Document", "ComboBox"])
            if not is_edit and name and name.strip():
                try:
                    r = control.rectangle()
                    # 检查是否在支行框下方（Y坐标接近）且在合理范围内
                    y_diff = abs(r.top - rect_hint[1])
                    if 0 <= y_diff <= 40 and abs(r.left - rect_hint[0]) < 300:
                        if "无需填写" in name or "则无需" in name:
                            print(f"扫描到支行提示文本: {name}")
                            raise StopIteration()
                except StopIteration:
                    raise
                except Exception:
                    pass
        except StopIteration:
            raise
        except Exception:
            pass
        try:
            for child in control.children():
                _scan_nearby_hints(child, rect_hint)
        except StopIteration:
            raise
        except Exception:
            pass

    branch_edit = None
    _find_branch_input(main_win)

    # 方案1：找不到支行输入框 -> 跳过
    if not branch_edit:
        print("未找到支行输入框，跳过")
        return True

    try:
        br_rect = branch_edit.rectangle()

        # 方案2：检测输入框是否被禁用（招商银行时支行框变灰）
        try:
            enabled = branch_edit.is_enabled() if hasattr(branch_edit, "is_enabled") else True
            if not enabled:
                print("支行输入框已禁用，跳过")
                return True
        except Exception:
            pass

        # 方案3：获取输入框值/名称中的提示词
        val = ""
        try:
            val = branch_edit.get_value() if hasattr(branch_edit, "get_value") else ""
        except Exception:
            pass
        name = branch_edit.element_info.name if hasattr(branch_edit.element_info, "name") else ""
        combined = (val + " " + name).lower()

        for h in ["无需填写", "则无需"]:
            if h in combined:
                print(f"支行框内容含'{h}'，跳过")
                return True

        # 方案4：扫描支行框附近的可见文本（找placeholder提示）
        hint_rect = (br_rect.left, br_rect.bottom + 5)
        try:
            _scan_nearby_hints(main_win, hint_rect)
        except StopIteration:
            print("检测到支行区域'无需填写'提示，跳过")
            return True

    except Exception as e:
        print(f"检测支行状态异常: {e}")

    print("支行需要填写")
    return False

    return False


def _input_branch_name(main_win, label_text, text_value):
    """精确查找支行输入框并填入：只匹配包含'联行号'的标签，避免误匹配其他字段"""

    def _find_exact_label(control):
        """精确查找支行标签（必须含'联行号'关键字）"""
        nonlocal label_ctrl
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
            # 排除输入框本身，只找文本标签
            is_edit = any(kw in cls for kw in ["Edit", "edit"]) or any(kw in ctrl_type for kw in ["Edit", "Document"])
            # 支宽匹配：含"支行名称"或"联行号"即可
            if not is_edit and ("联行号" in name or "支行名称" in name) and control.is_visible():
                rect = control.rectangle()
                if rect.width() > 3 and rect.height() > 3:
                    label_ctrl = control
                    return True
        except Exception:
            pass
        try:
            for child in control.children():
                if _find_exact_label(child):
                    return True
        except Exception:
            pass
        return False

    def _find_nearby_edit(label):
        """从标签旁找编辑框"""
        candidates = []
        try:
            parent = label.parent()
            if parent is None:
                return None
            ly = label.rectangle().top
            for sib in parent.children():
                if not sib.is_visible():
                    continue
                stype = str(sib.element_info.control_type) if hasattr(sib.element_info, "control_type") else ""
                scls = sib.element_info.class_name if hasattr(sib.element_info, "class_name") else ""
                is_edit = any(kw in scls for kw in ["Edit", "edit"]) or any(kw in stype for kw in ["Edit", "ComboBox", "Combo"])
                if is_edit:
                    sr = sib.rectangle()
                    yd = abs(sr.top - ly)
                    candidates.append((sib, yd))
        except Exception:
            pass
        if not candidates:
            return None
        candidates.sort(key=lambda x: x[1])
        return candidates[0][0]

    # 查找标签
    label_ctrl = None
    _find_exact_label(main_win)
    if not label_ctrl:
        print(f"未找到支行标签[{label_text}]")
        return False

    # 找输入框
    edit_ctrl = _find_nearby_edit(label_ctrl)
    if not edit_ctrl:
        print("未找到支行输入框")
        return False

    # 点击聚焦 -> 清空 -> 逐字键入
    edit_ctrl.set_focus()
    time.sleep(0.4)
    _type_text_slow(text_value, char_delay=0.09)
    time.sleep(0.5)
    print(f"已输入 [{label_text}]: {text_value}")
    return True


def _try_select_branch_dropdown(main_win):
    """检测支行下拉列表，只选择真正的选项（严格排除Edit/输入框控件）"""

    def _find_valid_branch_items(control, out_list):
        """递归查找支行下拉选项（排除输入框自身）"""
        try:
            if not control.is_visible():
                return
        except Exception:
            return
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
            rect = control.rectangle()
            # 只排除真正的输入框（Edit、Document）
            is_input_ctrl = any(kw in ctrl_type for kw in ["Edit", "Document"])
            # 只排除输入框本身（名称含"联行号"且是输入类控件）
            is_input_name = "联行号" in name and is_input_ctrl
            if is_input_name:
                return  # 只跳过支行输入框自身

            # 匹配条件：含"支行"、单行高度、合理宽度、非空名称
            name_len = len(name.strip()) if name else 0
            is_single_item = 5 <= name_len <= 60
            is_single_row = 15 <= rect.height() <= 55 and rect.width() > 80
            if name and "支行" in name and is_single_item and is_single_row:
                out_list.append((rect, name))
        except Exception:
            pass
        try:
            for child in control.children():
                _find_valid_branch_items(child, out_list)
        except Exception:
            pass

    print("检测支行下拉列表...")
    end_time = time.time() + 6
    while time.time() < end_time:
        items = []
        _find_valid_branch_items(main_win, items)
        if items:
            items.sort(key=lambda x: x[0].top)
            first_rect, first_name = items[0]
            mx = first_rect.mid_point().x + 60  # 偏右靠近"行"字
            my = first_rect.mid_point().y
            print(f"找到有效支行选项: {first_name}，位置({mx}, {my})")
            click_at(mx, my)
            time.sleep(0.15)
            click_at(mx, my)  # 双击
            print(f"已双击选择: {first_name}")
            return True
        time.sleep(0.3)

    print("未找到有效支行下拉选项")
    return False


def _collect_submit_button_candidates(main_win):
    """遍历 UIA 树，定位真正的「经办」按钮候选；只读控件，不点击。

    页面顶部标题、页签和面包屑里也会出现「单笔转账经办」。这些不是提交按钮，
    所以这里仅把可见且文案规整后等于「经办」的控件列为可点击候选。
    """
    candidates = []
    related = []

    def _walk(control):
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
            visible = control.is_visible()
            rect = control.rectangle()
        except Exception:
            name = ""
            ctrl_type = ""
            cls = ""
            visible = False
            rect = None

        if name and "经办" in name and rect is not None:
            w = rect.width()
            h = rect.height()
            if visible and w > 5 and h > 5:
                stripped = name.strip()
                normalized = "".join(stripped.split())
                is_buttonish = "Button" in ctrl_type or "Button" in cls or "button" in cls.lower()
                item = {
                    "name": stripped,
                    "normalized": normalized,
                    "type": ctrl_type,
                    "class": cls,
                    "rect": rect,
                    "center": (rect.mid_point().x, rect.mid_point().y),
                    "control": control,
                    "buttonish": is_buttonish,
                }
                if normalized == "经办":
                    priority = 0 if is_buttonish else 1
                    # 同名候选里优先选更像按钮的，其次选更靠下的，避开顶部导航区。
                    candidates.append((priority, -rect.top, rect.left, item))
                else:
                    related.append(item)

        try:
            for child in control.children():
                _walk(child)
        except Exception:
            pass

    _walk(main_win)
    candidates.sort(key=lambda t: (t[0], t[1], t[2]))
    return [item for _, _, _, item in candidates], related


def _print_submit_probe(stage, main_win, prefer_confirmation=False):
    candidates, related = _collect_submit_button_candidates(main_win)
    if candidates:
        print(f"{stage}: 找到 {len(candidates)} 个「经办」按钮候选")
        for idx, item in enumerate(candidates[:5], start=1):
            rect = item["rect"]
            cx, cy = item["center"]
            print(
                f"  候选#{idx}: name={item['name']!r}, type={item['type']}, "
                f"class={item['class']}, rect=({rect.left},{rect.top},{rect.right},{rect.bottom}), "
                f"center=({cx},{cy})"
            )
        selected = candidates[0]
        if prefer_confirmation and len(candidates) > 1:
            bottom_top = candidates[0]["rect"].top
            for item in candidates[1:]:
                if item["rect"].top < bottom_top - 10:
                    selected = item
                    break
        rect = selected["rect"]
        cx, cy = selected["center"]
        print(
            f"{stage}: 选中「经办」候选 name={selected['name']!r}, "
            f"rect=({rect.left},{rect.top},{rect.right},{rect.bottom}), center=({cx},{cy})"
        )
        return selected

    print(f"{stage}: 未找到可见的「经办」按钮候选")
    for item in related[:5]:
        rect = item["rect"]
        cx, cy = item["center"]
        print(
            f"  排除的相关控件: name={item['name']!r}, type={item['type']}, "
            f"class={item['class']}, rect=({rect.left},{rect.top},{rect.right},{rect.bottom}), "
            f"center=({cx},{cy})"
        )
    return None


def _has_submit_validation_error(main_win):
    """点击经办后检查页面是否仍因必填/格式问题停在表单。"""
    needles = ("请填写正确的经办信息", "请填写正确", "经办信息")
    hits = []

    def _walk(control):
        try:
            if not control.is_visible():
                return
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            if any(needle in name for needle in needles):
                hits.append(name.strip())
                return
        except Exception:
            pass

        try:
            for child in control.children():
                _walk(child)
        except Exception:
            pass

    _walk(main_win)
    if hits:
        print(f"经办后检测到页面校验提示: {hits[0]}")
        return True
    return False


def _wait_for_submit_validation_error(main_win, timeout=2.5):
    end_ts = time.time() + timeout
    while time.time() < end_ts:
        _check_ubank_crash("经办后校验提示检测")
        if _has_submit_validation_error(main_win):
            _screenshot("10_经办校验失败")
            return True
        time.sleep(0.3)
    return False


def _click_submit_candidate(item, stage_desc):
    rect = item["rect"]
    cx, cy = item["center"]
    print(
        f"{stage_desc}: 准备点击「经办」按钮 name={item['name']!r}, "
        f"type={item['type']}, class={item['class']}, "
        f"rect=({rect.left},{rect.top},{rect.right},{rect.bottom}), center=({cx},{cy})"
    )

    control = item.get("control")
    if control is not None and item.get("buttonish"):
        try:
            control.invoke()
            print(f"{stage_desc}: 已通过 UIA invoke 点击「经办」按钮")
            time.sleep(0.8)
            return True
        except Exception as e:
            print(f"{stage_desc}: UIA invoke 点击失败，改用坐标点击: {e}")

    click_at(cx, cy)
    print(f"{stage_desc}: 已通过坐标点击「经办」按钮")
    time.sleep(0.8)
    return True


def _click_submit_button(main_win, prefer_confirmation=False):
    """填表后滚动到底部并点击真正的「经办」按钮。"""
    print("开始定位并点击「经办」按钮...")
    _check_ubank_crash("点击经办前")
    _screenshot("07_经办点击_滚动前")
    best = _print_submit_probe("滚动前", main_win, prefer_confirmation=prefer_confirmation)
    if best:
        if _click_submit_candidate(best, "滚动前"):
            _screenshot("09_经办点击后")
            if _wait_for_submit_validation_error(main_win):
                return False
            return True

    try:
        main_win.set_focus()
    except Exception:
        pass

    # 先用 End/PageDown，再用滚轮，把页面底部按钮区带入可见区域。
    for _ in range(2):
        press_keys((0x23, 0), (0x23, 2))  # End
        time.sleep(0.25)
        _check_ubank_crash("点击经办 End 滚动")
    for _ in range(2):
        press_keys((0x22, 0), (0x22, 2))  # PageDown
        time.sleep(0.25)
        _check_ubank_crash("点击经办 PageDown 滚动")
    for _ in range(6):
        _wheel_foreground_client(0.70, 0.72, -5)
        _check_ubank_crash("点击经办 滚轮滚动")

    _screenshot("08_经办点击_滚动后")
    best = _print_submit_probe("滚动后", main_win, prefer_confirmation=prefer_confirmation)
    if best:
        rect = best["rect"]
        cx, cy = best["center"]
        print(f"经办按钮最终定位: rect=({rect.left},{rect.top},{rect.right},{rect.bottom}), center=({cx},{cy})")
        if _click_submit_candidate(best, "滚动后"):
            _screenshot("09_经办点击后")
            if _wait_for_submit_validation_error(main_win):
                return False
            return True
    else:
        print("经办按钮最终定位失败：滚动到底部后仍未在 UIA 树中发现可见候选，已保存截图供排查")
    return False


def fill_transfer_form(main_win, form=None, form_loaded=False):
    """
    填写单笔转账经办表单
    优先从当前财务目录下 M3直供合同付款数据获取\\bank_form.json 读取字段；
    若不存在则回退到内置示例值。
    """
    _form = form or {}

    if form_loaded:
        payee_acct = _form.get("收方账号", "")
        payee_name = _form.get("收方户名", "")
        bank_head = _form.get("开户银行", "")
        branch_full = _form.get("支行名称", "")
        amount = _form.get("金额", "")
        purpose = _form.get("用途", "")
    else:
        payee_acct = _form.get("收方账号") or "000000000000000000"
        payee_name = _form.get("收方户名") or "倪瑞"
        bank_head = _form.get("开户银行") or "中国银行"
        branch_full = _form.get("支行名称") or "中国银行内蒙古锡林浩特市团结支行"
        amount = _form.get("金额") or "1748.96"
        purpose = _form.get("用途") or "渠道锡林浩特结算1月分佣"

    _check_ubank_crash("fill: 进入 fill_transfer_form")
    _dismiss_transfer_overlays()
    _check_ubank_crash("fill: dismiss_overlays 后")

    # 收方账号
    if not find_and_input_text(main_win, "收方账号", payee_acct):
        _click_ratio_and_paste(0.545, 0.557, payee_acct, "收方账号")
    _check_ubank_crash("fill: 收方账号 后")

    # 收方户名
    if not find_and_input_text(main_win, "收方户名", payee_name):
        _click_ratio_and_paste(0.545, 0.616, payee_name, "收方户名")
    _check_ubank_crash("fill: 收方户名 后")

    # 开户银行：专用函数处理（搜索下拉框）
    _input_bank_name(main_win, bank_head)
    _check_ubank_crash("fill: 开户银行 后")

    # 支行名称/联行号
    branch_name = branch_full
    _screenshot("03_支行检测")
    print(f"开始检测支行是否需要填写...")

    # 检测支行输入框当前内容，判断是否需要填写
    skip_branch = _check_skip_branch(main_win)
    _check_ubank_crash("fill: 支行 skip 检测后")
    if skip_branch:
        print("检测到支行无需填写，已跳过支行步骤")
        _screenshot("05_支行跳过")
    else:
        print(f"开始填写支行名称: {branch_name}")
        branch_ok = _input_branch_name(main_win, "支行名称/联行号", branch_name)
        _check_ubank_crash("fill: 支行 输入后")
        if branch_ok:
            time.sleep(1.5)
            _check_ubank_crash("fill: 支行 输入后等待")
            _screenshot("04_支行填后下拉前")
            branch_clicked = _try_select_branch_dropdown(main_win)
            _check_ubank_crash("fill: 支行 双击下拉后")
            if not branch_clicked:
                print("支行下拉列表未检测到有效选项，已跳过支行选择")
            _screenshot("05_支行选择后")

    _trace_fill_step("支行阶段结束，稳定等待 3s 开始")
    _sleep_and_check(3, "fill: 支行阶段结束稳定等待")
    _trace_fill_step("支行阶段结束，稳定等待 3s 完成")

    # ---- 转账信息 ----

    # 金额
    _trace_fill_step("金额 输入开始")
    if not find_and_input_text(main_win, "金额(￥)", amount):
        _click_ratio_and_paste(0.545, 0.843, amount, "金额")
    _trace_fill_step("金额 输入函数返回后")
    _check_ubank_crash("fill: 金额 后")
    _sleep_and_check(1, "fill: 金额 后 1s 稳定等待")
    _trace_fill_step("金额 后 1s 稳定等待完成")

    # 用途
    _trace_fill_step("用途 输入开始")
    if not find_and_input_text(main_win, "用途", purpose):
        _click_ratio_and_paste(0.525, 0.897, purpose, "用途")
    _trace_fill_step("用途 输入函数返回后")
    _check_ubank_crash("fill: 用途 后")
    _sleep_and_check(1, "fill: 用途 后 1s 稳定等待")
    _trace_fill_step("用途 后 1s 稳定等待完成")

    _screenshot("06_全部填完")
    _check_ubank_crash("fill: 全部填完 截图后")


def _main_impl():
    """主流程：打开 -> 登录 -> 点击转账支付 -> 点击单笔转账经办 -> 填写表单 -> 到达制单页

    任何中途失败都以非零退出码结束，让 monitor.run_zhidan 把它当失败处理、
    不要标记 seen，下一轮自动重试。
    """

    # 1. 读取表单并切换 U 盾
    form, form_loaded = _load_bank_form()
    _log_display_profile()
    # [临时] 调试模式：跳过 USB Hub 切换，直接用当前已插着的 U 盾测试制单流程
    print("[调试] 已跳过 USB Hub 切换，直接进入制单流程")
    # if not _switch_usb_hub_for_form(form, form_loaded):
    #     print("USB Hub 切换失败或付款单位未配置，终止执行")
    #     sys.exit(2)

    # 2. 打开应用
    _check_ubank_crash("启动前")
    if open_ubank() is False:
        sys.exit(2)
    _sleep_and_check(1, "启动后")

    # 3. 登录
    success = login_ubank()
    _check_ubank_crash("登录后")
    if not success:
        print("登录失败，终止执行")
        sys.exit(3)

    # 4. 等待主界面加载
    desktop = Desktop(backend="uia")
    print("等待主界面加载...")
    main_win = wait_for_main_window(desktop)
    if not main_win:
        print("未找到主界面窗口，终止执行")
        sys.exit(4)

    _sleep_and_check(10, "等待主界面页面加载")  # 等待页面完全加载
    print("页面加载完成...")

    # 5. 点击顶部导航栏的"转账支付"
    _check_ubank_crash("点击转账支付前")
    click_control_by_name(main_win, "转账支付")
    _sleep_and_check(2, "等待转账支付菜单")  # 等待转账支付菜单展开/跳转

    # 6. 点击左侧"单笔转账经办"
    _check_ubank_crash("点击单笔转账经办前")
    click_control_by_name(main_win, "单笔转账经办")
    _sleep_and_check(3, "等待单笔转账经办页面")  # 等待单笔转账经办页面加载

    # 7. 验证是否成功到达单笔转账经办页面
    print("验证目标页面...")
    success_flag = False
    try:
        for _ in range(10):
            _check_ubank_crash("验证目标页面")
            windows = desktop.windows()
            for w in windows:
                title = w.window_text()
                if "单笔转账经办" in title or "单笔转账" in title:
                    success_flag = True
                    print(f"已到达目标页面: {title}")
                    break
            if success_flag:
                break
            _sleep_and_check(1, "等待目标页面")
    except Exception as e:
        if isinstance(e, UBankCrashDetected):
            raise
        print(f"验证过程异常: {e}")

    if not success_flag:
        print("未能确认进入单笔转账经办页面")
        _sleep_and_check(1, "目标页面失败后准备退出")
        _close_with_confirm(main_win)
        sys.exit(5)

    print("已成功进入单笔转账经办页面")

    # 8. 填写转账表单
    _sleep_and_check(2, "等待表单控件加载")  # 等待表单控件完全加载
    print("开始填写表单...")
    _check_ubank_crash("填写表单前")
    fill_transfer_form(main_win, form, form_loaded=form_loaded)
    _check_ubank_crash("填写表单后")

    if not _click_submit_button(main_win):
        print("未能点击第一次「经办」按钮，终止执行，避免将未提交流水标记为成功")
        _close_with_confirm(main_win)
        sys.exit(SUBMIT_CLICK_EXIT_CODE)
    _check_ubank_crash("第一次经办按钮点击后")
    _sleep_and_check(3, "第一次经办点击后等待二次经办页面")
    _screenshot("10_第一次经办等待后")

    print("开始使用同一套定位方法点击第二次「经办」按钮...")
    if not _click_submit_button(main_win, prefer_confirmation=True):
        print("未能点击第二次「经办」按钮，终止执行，避免将未完整经办流水标记为成功")
        _close_with_confirm(main_win)
        sys.exit(SUBMIT_CLICK_EXIT_CODE)
    _check_ubank_crash("第二次经办按钮点击后")
    _sleep_and_check(3, "第二次经办点击后等待结果")
    _screenshot("11_第二次经办点击后等待结果")

    print("表单填写完成，准备退出 U-BANK...")

    # 9. 退出：直接 taskkill，不触发 Firmbank 的关闭路径
    _sleep_and_check(1, "退出前")
    _close_with_confirm(main_win)

    print("========== 制单流程完成 ==========")
    print("表单填写完毕，U-BANK 关闭策略由 _close_with_confirm 上方日志说明")
    print("==================================")


def main():
    """启动全局崩溃守护线程后运行制单主流程。"""
    global _GLOBAL_WATCHER

    _clear_stale_ubank_crash_dialogs()

    watcher = _CrashWatcher(poll_interval=0.02)
    _GLOBAL_WATCHER = watcher
    watcher.start()
    print("已启动 U-BANK 崩溃守护线程（WinEventHook + 20ms 轮询兜底）")
    if CRASH_DIAG_MODE:
        print("已开启 U-BANK 崩溃诊断模式：发现崩溃时不强杀，LocalDump 目录为 C:\\tmp\\FirmbankDumps")

    try:
        try:
            _main_impl()
        except UBankCrashDetected:
            raise
        except Exception:
            if watcher.fired:
                raise UBankCrashDetected(watcher.fired_title) from None
            raise

        if watcher.fired:
            raise UBankCrashDetected(watcher.fired_title)
    finally:
        watcher.stop()
        if _GLOBAL_WATCHER is watcher:
            _GLOBAL_WATCHER = None


if __name__ == "__main__":
    try:
        main()
    except UBankCrashDetected as e:
        print(f"招行客户端崩溃，制单失败: {e}")
        sys.exit(UBANK_CRASH_EXIT_CODE)
