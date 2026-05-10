# -*- coding: utf-8 -*-
"""
招行U-BANK - 制单：自动登录并进入单笔转账经办页面，填写转账表单
"""
import time
import subprocess
import win32gui
import os
import sys
from datetime import datetime

_ZHIDAN_ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
if _ZHIDAN_ROOT not in sys.path:
    sys.path.insert(0, _ZHIDAN_ROOT)

from pywinauto import Desktop
from ubank_common import (
    click_at, press_keys, open_ubank, login_ubank,
    wait_for_main_window, click_control_by_name,
)


_PS_SET_CLIPBOARD = (
    "[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false); "
    "$txt = [Console]::In.ReadToEnd(); "
    "Set-Clipboard -Value $txt"
)


def _set_clipboard_text(text_value):
    """通过 stdin 传 UTF-8 给 PowerShell；失败 raise 避免粘贴旧剪贴板。"""
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-Command", _PS_SET_CLIPBOARD],
        input=text_value or "", text=True, encoding="utf-8",
        capture_output=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"Set-Clipboard 失败 rc={proc.returncode}: {proc.stderr.strip()}")


# 截图保存目录
_SCREENSHOT_DIR = os.path.join(_ZHIDAN_ROOT, "screenshots")
os.makedirs(_SCREENSHOT_DIR, exist_ok=True)

# ============================================================
# 账号选择配置（只需修改此参数即可切换付款账号）
# 可选值: "001" 或 "002"
# ============================================================
SELECT_ACCOUNT = "001"


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


def find_and_input_text(main_win, target_label, text_value):
    """根据标签名模糊查找相邻的文本输入框并输入内容"""

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
    _find_label(main_win)

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
    if edit_ctrl:
        try:
            edit_ctrl.set_focus()
            time.sleep(0.3)  # 加长等待确保聚焦完成
            # 判断控件类型，选择合适的输入方式
            ctrl_type = edit_ctrl.element_info.control_type if hasattr(edit_ctrl.element_info, "control_type") else ""
            cls_name = edit_ctrl.element_info.class_name if hasattr(edit_ctrl.element_info, "class_name") else ""

            # 下拉框/组合框
            is_combo = "Combo" in ctrl_type or "Combo" in cls_name or "ComboBox" in str(ctrl_type)

            if is_combo:
                rect = edit_ctrl.rectangle()
                click_at(rect.mid_point().x, rect.mid_point().y)
                time.sleep(0.5)

            # 统一用剪贴板+Ctrl+V粘贴（解决中文首字符丢失问题）
            if text_value:
                # 先全选已有内容（Ctrl+A），再粘贴覆盖
                press_keys(
                    (0x11, 0), (0x41, 0),  # Ctrl down, A down
                    (0x41, 2), (0x11, 2),  # A up, Ctrl up
                )
                time.sleep(0.15)
                _set_clipboard_text(text_value)
                time.sleep(0.2)
                # Ctrl+V 粘贴
                press_keys(
                    (0x11, 0), (0x56, 0),  # Ctrl down, V down
                    (0x56, 2), (0x11, 2),  # V up, Ctrl up
                )
            else:
                # 普通文本框：直接键盘输入
                edit_ctrl.type_keys(text_value)

            print(f"已输入 [{target_label}]: {text_value}")
            return True
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


def _close_with_confirm(main_win):
    """点击右上角X关闭窗口，并处理后续的确认弹窗"""
    print("开始关闭程序...")

    # 1. 聚焦主窗口并发送 Alt+F4 触发关闭（U-BANK 自绘标题栏，坐标点 X 不可靠）
    try:
        main_win.set_focus()
        time.sleep(0.3)
        # Alt down, F4 down, F4 up, Alt up
        press_keys(
            (0x12, 0), (0x73, 0),  # ALT down, F4 down
            (0x73, 2), (0x12, 2),  # F4 up, ALT up
        )
        print("已发送 Alt+F4")
    except Exception as e:
        print(f"发送 Alt+F4 失败: {e}")
        return

    time.sleep(1.5)

    # 2. 点击第一个确认弹窗
    desktop = Desktop(backend="uia")
    _click_confirm_button(desktop, "确认", "第1次")

    time.sleep(2)

    # 3. 点击第二个确认弹窗
    _click_confirm_button(desktop, "确认", "第2次")

    print("程序已退出")


def _click_confirm_button(desktop, button_text, step_desc):
    """
    退出确认弹窗的多策略顺序（嵌入 Web 时 UIA 往往不完整）：
    1) 轮询 UIA 树，找名称含「确认/确定」的按钮类控件并点击
    2) 前台窗体中心点一下抢焦点，再点右下角常见主按钮区，再发 Enter（默认按钮）
    3) 多按几次 Tab 后 Enter，覆盖焦点顺序与控件数量变化的情况
    """
    _ = button_text  # 保留参数与旧调用一致，逻辑已不依赖单一文案匹配 Text
    time.sleep(1.0)

    # 持续监听一小段时间，避免“弹窗刚出现时还未完成 UIA 树注册”
    end_ts = time.time() + 12
    while time.time() < end_ts:
        if _try_uia_confirm_click(desktop):
            print(f"{step_desc} 已通过 UIA 点击确认")
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
                time.sleep(0.2)
                press_keys((0x0D, 0), (0x0D, 2))
                return

        time.sleep(0.25)

    # 最终兜底：仅发送键盘确认，不做额外危险坐标点击
    print(f"{step_desc} 监听超时，改用 Enter + Tab 兜底")
    press_keys((0x0D, 0), (0x0D, 2))
    time.sleep(0.3)
    for _ in range(14):
        press_keys((0x09, 0), (0x09, 2))
        time.sleep(0.12)
    press_keys((0x0D, 0), (0x0D, 2))
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
    """专门处理开户银行输入：复用通用查找 -> 清空 -> 粘贴 -> 点击下拉选项"""
    print(f"开始填写开户银行: {bank_name}")

    # 截图：填写前
    _screenshot("01_开户银行填前")

    # 1. 用通用函数查找并聚焦输入框
    find_and_input_text(main_win, "开户银行", "")
    time.sleep(0.5)

    # 2. 全选 + Backspace 清空
    press_keys((0x11, 0), (0x41, 0), (0x41, 2), (0x11, 2))  # Ctrl+A
    time.sleep(0.15)
    for _ in range(30):
        press_keys((0x08, 0), (0x08, 2))  # Backspace
    time.sleep(0.2)

    # 3. 剪贴板粘贴银行名称
    _set_clipboard_text(bank_name)
    time.sleep(0.2)
    press_keys((0x11, 0), (0x56, 0), (0x56, 2), (0x11, 2))
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

    # 点击聚焦 -> 清空 -> 粘贴
    edit_ctrl.set_focus()
    time.sleep(0.4)
    press_keys((0x11, 0), (0x41, 0), (0x41, 2), (0x11, 2))  # Ctrl+A
    time.sleep(0.15)
    for _ in range(30):
        press_keys((0x08, 0), (0x08, 2))  # Backspace
    time.sleep(0.15)

    _set_clipboard_text(text_value)
    time.sleep(0.2)
    press_keys((0x11, 0), (0x56, 0), (0x56, 2), (0x11, 2))
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


def _select_pay_account(main_win, account_suffix):
    """
    在进入单笔转账经办页面后，点击付款账号下拉框，选择指定账号（001或002）
    下拉选项格式类似: "1109 6079 5610 001, 北京巡鲜电子商务有限公司,..."
    只需匹配账号后缀(如"001")即可定位目标选项

    定位策略：直接查找包含账号数字模式的ComboBox/编辑控件
    （因为"付方账号"标签在Web UI中不暴露为Text控件）
    """

    def _find_pay_account_combo(control):
        """查找付款账号组合框（名称含'付方账号'的ComboBox控件）"""
        nonlocal pay_combo_ctrl
        try:
            if not control.is_visible():
                return False
        except Exception:
            return False
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
            rect = control.rectangle()
            # 匹配组合框类控件
            is_combo = any(kw in cls for kw in ["ComboBox", "Combo"]) or \
                       any(kw in ctrl_type for kw in ["ComboBox", "Combo"])
            # 控件名中包含"付方账号"
            if is_combo and name and "付方账号" in name and rect.width() > 50 and rect.height() > 5:
                pay_combo_ctrl = control
                return True
        except Exception:
            pass
        try:
            for child in control.children():
                if _find_pay_account_combo(child):
                    return True
        except Exception:
            pass
        return False

    def _find_pay_account_dropdown_items(control, out_list):
        """递归查找付款账号下拉列表中的选项"""
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
            # 排除输入框/编辑控件本身
            is_input = any(kw in ctrl_type for kw in ["Edit", "Document"]) or any(kw in cls for kw in ["Edit", "edit"])
            if is_input:
                return
            # 匹配包含目标账号后缀的选项（单行、合理尺寸）
            name_len = len(name.strip()) if name else 0
            is_single_row = 15 <= rect.height() <= 55 and rect.width() > 100
            if name and account_suffix in name and is_single_row and 5 <= name_len <= 80:
                out_list.append((rect, name))
        except Exception:
            pass
        try:
            for child in control.children():
                _find_pay_account_dropdown_items(child, out_list)
        except Exception:
            pass

    print(f"准备选择付款账号: ...{account_suffix}")

    # 第一步：通过账号数字模式找到付款账号组合框
    pay_combo_ctrl = None
    _find_pay_account_combo(main_win)

    if not pay_combo_ctrl:
        print("未找到付款账号输入框，跳过账号选择")
        return False

    # 第二步：多种策略展开下拉并选择目标账号
    try:
        # 先通过附近Text控件判断当前已选中的账号值
        def _find_displayed_account_value(control):
            """在付款账号ComboBox附近查找显示的实际账号文本"""
            nonlocal displayed_value
            try:
                name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
                ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
                cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
                # 匹配包含账号数字模式的Text控件（排除ComboBox自身）
                is_text = ctrl_type == "Text"
                if is_text and name and len(name) > 20:
                    import re
                    if re.search(r'\d{4}\s+\d{4}\s+\d{4}', name):
                        displayed_value = name
                        return True
            except Exception:
                pass
            try:
                for child in control.children():
                    if _find_displayed_account_value(child):
                        return True
            except Exception:
                pass
            return False

        displayed_value = ""
        _find_displayed_account_value(main_win)
        print(f"付款账号当前显示值: {displayed_value[:70]}..." if displayed_value else "未找到显示值")

        # 如果当前值已包含目标账号后缀，说明已是目标账号，无需切换
        if account_suffix in str(displayed_value):
            print(f"付款账号已是目标 ...{account_suffix}，无需切换")
            return True

        rect = pay_combo_ctrl.rectangle()
        cx = rect.mid_point().x
        cy = rect.mid_point().y

        # 尝试pywinauto原生的select方法（对标准ComboBox有效）
        try:
            all_texts = pay_combo_ctrl.texts() if hasattr(pay_combo_ctrl, "texts") else []
            if all_texts:
                print(f"ComboBox可用选项({len(all_texts)}个): {[t[:40] for t in all_texts[:5]]}")
                for txt in all_texts:
                    if account_suffix in str(txt):
                        pay_combo_ctrl.select(txt)
                        print(f"已通过select()选择: {txt[:50]}...")
                        time.sleep(0.5)
                        return True
        except Exception as e:
            print(f"  select/texts方法不可用: {e}")

        # 策略：聚焦ComboBox -> 点击输入区（左侧）-> 清空 -> 输入目标筛选文本
        # （Web嵌入的Ant Design/Element Select控件：输入文本可触发过滤+自动选择）
        rect = pay_combo_ctrl.rectangle()

        # 点击ComboBox左侧输入区域（不是箭头）
        click_at(rect.left + 100, rect.mid_point().y)
        time.sleep(0.5)
        print("已点击付款账号输入区域")

        # 全选并清空
        press_keys((0x11, 0), (0x41, 0), (0x41, 2), (0x11, 2))  # Ctrl+A
        time.sleep(0.15)

        # 输入目标账号后缀来搜索匹配
        _set_clipboard_text(account_suffix)
        time.sleep(0.2)
        press_keys((0x11, 0), (0x56, 0), (0x56, 2), (0x11, 2))  # Ctrl+V
        print(f"已输入筛选: {account_suffix}")
        time.sleep(2.0)  # 给足够时间让Web响应

        # 检查值是否变化
        new_displayed = ""
        _find_displayed_account_value(main_win)
        new_displayed = displayed_value
        if account_suffix in str(new_displayed):
            print(f"输入筛选后已匹配 ...{account_suffix}: {new_displayed[:50]}...")
            time.sleep(0.5)
            return True
        print(f"筛选后显示: {new_displayed[:60] if new_displayed else '(未检测到)'}...")

        # 尝试Enter确认选择
        press_keys((0x0D, 0), (0x0D, 2))  # Enter
        time.sleep(1.5)

        _find_displayed_account_value(main_win)
        new_displayed = displayed_value
        if account_suffix in str(new_displayed):
            print(f"Enter后已匹配 ...{account_suffix}")
            time.sleep(0.3)
            return True

    except Exception as e:
        print(f"操作付款账号失败: {e}")

    print(f"未能切换付款账号至 ...{account_suffix}（Web嵌入控件的限制）")
    return False


def fill_transfer_form(main_win):
    """
    填写单笔转账经办表单
    根据合同付款单据截图中的字段逐一输入：
    - 收方账号、收方户名、开户银行、支行名称联行号
    - 金额、用途
    """
    # ---- 合同付款单据：YLHTFK-2026-00154，申请人王志飞，日期2026-03-26 ----

    # 收方账号：来自合同付款 - 银行账户
    find_and_input_text(main_win, "收方账号", "000000000000000000")

    # 收方户名：来自合同付款 - 收款单位名称
    find_and_input_text(main_win, "收方户名", "倪瑞")

    # 开户银行：专用函数处理（搜索下拉框，需清空后输入并选择选项）
    _input_bank_name(main_win, "中国银行")

    # 支行名称/联行号：先检测是否需要填写（有些银行无需填支行）
    branch_name = "中国银行内蒙古锡林浩特市团结支行"
    _screenshot("03_支行检测")
    print(f"开始检测支行是否需要填写...")

    # 检测支行输入框当前内容，判断是否需要填写
    skip_branch = _check_skip_branch(main_win)
    if skip_branch:
        print("检测到支行无需填写，已跳过支行步骤")
        _screenshot("05_支行跳过")
    else:
        print(f"开始填写支行名称: {branch_name}")
        branch_ok = _input_branch_name(main_win, "支行名称/联行号", branch_name)
        if branch_ok:
            time.sleep(1.5)
            _screenshot("04_支行填后下拉前")
            branch_clicked = _try_select_branch_dropdown(main_win)
            if not branch_clicked:
                print("支行下拉列表未检测到有效选项，已跳过支行选择")
            _screenshot("05_支行选择后")

    # ---- 转账信息 ----

    # 金额：来自合同付款 - 申请金额
    find_and_input_text(main_win, "金额(￥)", "1748.96")

    # 用途：下拉框选择或输入（来自合同付款 - 申请说明前半句）
    find_and_input_text(main_win, "用途", "渠道锡林浩特结算1月分佣")

    _screenshot("06_全部填完")


def main():
    """主流程：打开 -> 登录 -> 点击转账支付 -> 点击单笔转账经办 -> 填写表单 -> 到达制单页"""

    # 1. 打开应用
    open_ubank()

    # 2. 登录
    success = login_ubank()
    if not success:
        print("登录失败，终止执行")
        sys.exit(3)

    # 3. 等待主界面加载
    desktop = Desktop(backend="uia")
    print("等待主界面加载...")
    main_win = wait_for_main_window(desktop)
    if not main_win:
        print("未找到主界面窗口，终止执行")
        sys.exit(4)

    time.sleep(10)  # 等待页面完全加载
    print("页面加载完成...")

    # 4. 点击顶部导航栏的"转账支付"
    click_control_by_name(main_win, "转账支付")
    time.sleep(2)  # 等待转账支付菜单展开/跳转

    # 5. 点击左侧"单笔转账经办"
    click_control_by_name(main_win, "单笔转账经办")
    time.sleep(3)  # 等待单笔转账经办页面加载

    # 6. 验证是否成功到达单笔转账经办页面
    print("验证目标页面...")
    success_flag = False
    try:
        for _ in range(10):
            windows = desktop.windows()
            for w in windows:
                title = w.window_text()
                if "单笔转账经办" in title or "单笔转账" in title:
                    success_flag = True
                    print(f"已到达目标页面: {title}")
                    break
            if success_flag:
                break
            time.sleep(1)
    except Exception as e:
        print(f"验证过程异常: {e}")

    if success_flag:
        print("已成功进入单笔转账经办页面")

        # 7. 选择付款账号（通过顶部 SELECT_ACCOUNT 参数控制选001或002）
        time.sleep(2)  # 等待付款账号控件完全加载
        if not _select_pay_account(main_win, SELECT_ACCOUNT):
            print("付款账号选择失败，终止执行")
            sys.exit(6)

        # 8. 填写转账表单
        time.sleep(1)  # 等待表单控件完全加载
        print("开始填写表单...")
        fill_transfer_form(main_win)

        print("========== 制单流程完成 ==========")
        print("已进入单笔转账经办页面并完成表单填写")
        print("==================================")

        # 8. 退出：点击右上角X，然后点击两次确定
        time.sleep(1)
        _close_with_confirm(main_win)
    else:
        print("未能确认进入单笔转账经办页面")
        time.sleep(1)
        _close_with_confirm(main_win)
        sys.exit(5)


if __name__ == "__main__":
    main()
