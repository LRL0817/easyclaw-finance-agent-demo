# -*- coding: utf-8 -*-
"""
招行U-BANK 公共模块 - 登录、截图、记录保存等共用功能
"""
import time
import os
import json
import re
import win32api
import win32con
import win32gui
from datetime import datetime
from dotenv import load_dotenv
from pywinauto import Desktop

# 加载.env配置文件（密码存储在此）
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(SCRIPT_DIR, ".env"))
SCREENSHOT_DIR = os.path.join(SCRIPT_DIR, "screenshots")


def ensure_dir(path):
    """确保目录存在"""
    if not os.path.exists(path):
        os.makedirs(path)


def click_at(x, y):
    """在指定坐标处点击鼠标左键"""
    win32api.SetCursorPos((x, y))
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def press_keys(*keys_with_action):
    """
    模拟组合按键，如 press_keys((0x11, 0), (0x56, 0), (0x56, 2), (0x11, 2))
    每个元素为 (vk_code, action)，action: 0=down, 2=up
    """
    for vk, action in keys_with_action:
        win32api.keybd_event(vk, 0, action, 0)
        time.sleep(0.05)


def capture_window_screenshot(win, save_path):
    """对窗口进行截图保存"""
    try:
        from PIL import ImageGrab
        rect = win.rectangle()
        screenshot = ImageGrab.grab(bbox=(
            int(rect.left), int(rect.top),
            int(rect.right), int(rect.bottom)
        ))
        screenshot.save(save_path)
        print(f"已截图保存: {save_path}")
        return True
    except ImportError:
        print("PIL未安装，使用备用截图方式...")
        return _capture_backup(save_path)
    except Exception as e:
        print(f"截图失败: {e}")
        return False


def _capture_backup(save_path):
    """备用截图方式：使用win32gui截取整个屏幕"""
    try:
        import win32gui
        import win32ui
        from ctypes import windll
        from PIL import Image

        hwnd = win32gui.GetDesktopWindow()
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top

        hdesktop = win32gui.GetDesktopWindow()
        hwndDC = win32gui.GetWindowDC(hdesktop)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)

        result = windll.user32.PrintWindow(hdesktop, saveDC.GetSafeHdc(), 0)

        try:
            if result == 1:
                bmpinfo = saveBitMap.GetInfo()
                bmpstr = saveBitMap.GetBitmapBits(True)
                img = Image.frombuffer(
                    'RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1
                )
                img.save(save_path)
                print(f"已截图保存(备用方式): {save_path}")
                return True
            else:
                print(f"PrintWindow返回值异常: {result}")
                return False
        finally:
            mfcDC.DeleteDC()
            saveDC.DeleteDC()
            win32gui.ReleaseDC(hdesktop, hwndDC)
            win32gui.DeleteObject(saveBitMap.GetHandle())

    except Exception as e2:
        print(f"备用截图也失败: {e2}")
        return False


def save_record(data, screenshot_path):
    """将数据记录保存为JSON文件"""
    record = {
        "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "数据": data,
        "截图路径": os.path.basename(screenshot_path) if screenshot_path else None,
    }
    # 保存到记录文件（追加模式）
    records_file = os.path.join(SCREENSHOT_DIR, "records.json")

    records = []
    if os.path.exists(records_file):
        with open(records_file, "r", encoding="utf-8") as f:
            try:
                records = json.load(f)
            except json.JSONDecodeError:
                records = []

    records.append(record)

    with open(records_file, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    print("数据记录已保存")
    return record


def open_ubank():
    """打开招行U-BANK应用"""
    lnk_path = os.environ.get("CMB_UBANK_SHORTCUT_PATH")
    if not lnk_path:
        candidates = [
            os.path.join(os.environ.get("PUBLIC", r"%USERPROFILE%"), "Desktop", "招行U-BANK.lnk"),
            os.path.join(os.path.expanduser("~"), "Desktop", "招行U-BANK.lnk"),
        ]
        lnk_path = next((path for path in candidates if os.path.exists(path)), candidates[0])
    os.startfile(lnk_path)
    print("已启动招行U-BANK")


def login_ubank():
    """自动输入密码并登录 - 密码从.env文件读取，若已登录则跳过"""
    desktop = Desktop(backend="uia")

    # 先检查是否已经登录（直接出现主窗口而非登录窗口）
    for _ in range(10):
        for w in desktop.windows():
            t = w.window_text()
            if ("招商银行" in t or "企业银行" in t) and "联机登录" not in t and t.strip():
                print(f"检测到已登录状态: {t}")
                return True
        time.sleep(1)

    # 未登录则走正常登录流程
    login_pwd = os.getenv("LOGIN_PWD", "")
    cert_pwd = os.getenv("CERT_PWD", "")
    if not login_pwd or not cert_pwd:
        print("错误: 未在.env中找到密码配置，请检查.env文件")
        return None

    for _ in range(30):
        win = desktop.window(title="联机登录")
        if win.exists(timeout=1):
            break
        time.sleep(1)
    else:
        print("未找到招行U-BANK登录窗口")
        return None

    print(f"已连接到窗口: {win.window_text()}")
    time.sleep(1)

    # 输入登录密码（通过遍历子控件找到ATL:Edit类型的密码框）
    _fill_edit_fields(win, login_pwd, cert_pwd)

    time.sleep(0.5)

    # 点击登录按钮（通过坐标点击，避免COM错误）
    try:
        login_btn = win.child_window(auto_id="2020", class_name="Button")
        if login_btn.exists():
            rect = login_btn.rectangle()
            click_at(rect.mid_point().x, rect.mid_point().y)
            print("已点击登录按钮")
            return True
        else:
            print("未找到登录按钮")
    except Exception as e:
        print(f"点击登录按钮失败: {e}")

    return False


def _fill_edit_fields(win, login_pwd, cert_pwd):
    """遍历登录窗口的ATL:Edit控件，依次输入登录密码和证书密码"""
    edit_count = 0
    for child in win.children():
        try:
            cls = child.element_info.class_name if hasattr(child.element_info, "class_name") else ""
            if "ATL" in cls and child.is_visible():
                for sub in child.children():
                    sub_cls = sub.element_info.class_name if hasattr(sub.element_info, "class_name") else ""
                    if "Edit" in sub_cls:
                        edit_count += 1
                        if edit_count == 1:
                            sub.set_focus()
                            sub.type_keys(login_pwd)
                            print("已输入登录密码")
                        elif edit_count == 2:
                            sub.set_focus()
                            sub.type_keys(cert_pwd)
                            print("已输入证书密码")
                            break
        except Exception:
            continue
    if edit_count == 0:
        print("未找到登录密码输入框")
    elif edit_count < 2:
        print("未找到证书密码输入框")


def wait_for_main_window(desktop):
    """等待主界面窗口加载完成"""
    possible_titles = ["V12", "U-BANK", "招商银行", "企业银行"]
    main_win = None

    for _ in range(20):
        windows = desktop.windows()
        for w in windows:
            title = w.window_text()
            for pt in possible_titles:
                if pt in title and "联机登录" not in title:
                    main_win = w
                    break
            if main_win:
                break
        if main_win:
            break
        time.sleep(1)

    if not main_win:
        for w in desktop.windows():
            title = w.window_text()
            if title.strip() and title not in ["任务栏", "", "Program Manager"]:
                if hasattr(w.element_info, "control_type") and w.element_info.control_type == "Window":
                    main_win = w
                    break

    if main_win:
        print(f"已连接主界面窗口: {main_win.window_text()}")
    else:
        print("未找到主界面窗口")
    return main_win


def click_control_by_name(main_win, target_name):
    """递归查找指定名称的控件并点击"""

    def _find_and_click(control):
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            if name and name.strip() == target_name:
                print(f"找到控件: {name}")
                rect = control.rectangle()
                click_at(rect.mid_point().x, rect.mid_point().y)
                print(f"已点击: {target_name}")
                return True
        except Exception:
            pass

        try:
            for child in control.children():
                result = _find_and_click(child)
                if result:
                    return True
        except Exception:
            pass
        return False

    print(f"正在查找: {target_name}...")
    clicked = _find_and_click(main_win)
    if not clicked:
        print(f"未找到控件: {target_name}")
    return clicked


def click_control_by_names(main_win, target_names):
    """按顺序尝试点击多个候选控件名称，命中一个即返回"""
    for name in target_names:
        if click_control_by_name(main_win, name):
            return True
    print(f"候选控件均未命中: {target_names}")
    return False


def _rect_close(a, b, tol=30):
    """判断两个矩形是否近似相同（用于匹配前台窗口）"""
    return (
        abs(a[0] - b[0]) <= tol and
        abs(a[1] - b[1]) <= tol and
        abs(a[2] - b[2]) <= tol and
        abs(a[3] - b[3]) <= tol
    )


def _uia_collect_confirm_candidates(control, out_list, host_rect):
    """递归收集名称含“确认/确定”的可点击控件候选"""
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
            if area > 20000 and not is_std_btn:
                for ch in control.children():
                    _uia_collect_confirm_candidates(ch, out_list, host_rect)
                return
            pri_name = 0 if "确认" in nm else 1
            pri_type = 0 if is_std_btn else (1 if ("Hyperlink" in ctype or "Custom" in ctype) else 2)
            mx, my = rect.mid_point().x, rect.mid_point().y
            hr_left, hr_top, hr_right, _ = host_rect
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


def _try_uia_confirm_click(desktop):
    """仅在前台窗口 UIA 树中查找并点击确认按钮"""
    fg_hwnd = win32gui.GetForegroundWindow()
    if not fg_hwnd:
        return False
    fg_rect = win32gui.GetWindowRect(fg_hwnd)
    candidates = []
    for w in desktop.windows():
        try:
            wr = w.rectangle()
            host_rect = (wr.left, wr.top, wr.right, wr.bottom)
            if not _rect_close(host_rect, fg_rect):
                continue
            _uia_collect_confirm_candidates(w, candidates, host_rect)
        except Exception:
            pass
    if not candidates:
        return False
    candidates.sort(key=lambda t: (t[0], t[1], t[2]))
    _, _, _, x, y, label = candidates[0]
    print(f"UIA 定位到确认控件: {label!r} -> ({x}, {y})")
    click_at(x, y)
    time.sleep(0.35)
    return True


def _click_confirm_button(desktop, step_desc):
    """点击单次确认弹窗（优先 UIA，次选几何点击）"""
    time.sleep(1.0)
    end_ts = time.time() + 12
    while time.time() < end_ts:
        if _try_uia_confirm_click(desktop):
            print(f"{step_desc} 已通过 UIA 点击确认")
            return
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            if 320 <= width <= 1200 and 160 <= height <= 700:
                x = int(left + width * 0.56)
                y = int(top + height * 0.84)
                print(f"{step_desc} UIA 未命中，几何点击确认: ({x}, {y})")
                click_at(x, y)
                time.sleep(0.2)
                press_keys((0x0D, 0), (0x0D, 2))
                return
        time.sleep(0.25)
    print(f"{step_desc} 监听超时，发送 Enter + Tab 兜底")
    press_keys((0x0D, 0), (0x0D, 2))
    time.sleep(0.3)
    for _ in range(14):
        press_keys((0x09, 0), (0x09, 2))
        time.sleep(0.12)
    press_keys((0x0D, 0), (0x0D, 2))


def close_with_confirm(main_win):
    """关闭 U-BANK 并连续处理两次确认弹窗"""
    print("开始执行退出流程...")
    try:
        main_win.set_focus()
        time.sleep(0.3)
    except Exception:
        pass

    press_keys(
        (0x12, 0), (0x73, 0),  # Alt down, F4 down
        (0x73, 2), (0x12, 2),  # F4 up, Alt up
    )
    print("已发送 Alt+F4")
    time.sleep(1.5)

    desktop = Desktop(backend="uia")
    _click_confirm_button(desktop, "第1次")
    time.sleep(2)
    _click_confirm_button(desktop, "第2次")
    print("退出流程完成")
