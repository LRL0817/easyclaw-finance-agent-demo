# -*- coding: utf-8 -*-
"""
招行U-BANK - 批量制单：自动登录并进入批量转账经办页面
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
# 导入文件配置（只需修改此参数即可切换Excel文件）
# 填写Excel文件的完整路径，例如: r"%USERPROFILE%\data.xlsx"
# ============================================================
IMPORT_FILE_PATH = os.environ.get(
    "CMB_IMPORT_FILE_PATH",
    os.path.join(_ZHIDAN_ROOT, "招行批量付款-克拉波.xlsx"),
)



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
        if w >= 6 and h >= 6 and ("确认" in nm or "确定" in nm or "继续导入" in nm or "导入正确" in nm):
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
            # 保护：不点击窗口右上角附近（避免误点关闭"X"）
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
    """在前台窗口的UIA树中查找并点击「确认/确定/继续导入」等按钮"""
    fg_hwnd = win32gui.GetForegroundWindow()
    if not fg_hwnd:
        return False
    fg_rect = win32gui.GetWindowRect(fg_hwnd)
    candidates = []
    for w in desktop.windows():
        try:
            wr = w.rectangle()
            host_rect = (wr.left, wr.top, wr.right, wr.bottom)
            # 只在前台窗口中查找，避免命中主页面里的"确认"按钮
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


def _close_with_confirm(main_win):
    """关闭窗口，并处理后续的确认弹窗"""
    print("开始关闭程序...")

    try:
        main_win.set_focus()
        time.sleep(0.3)
    except Exception:
        pass

    # 1. 先尝试 Alt+F4
    press_keys(
        (0x12, 0), (0x73, 0),
        (0x73, 2), (0x12, 2),
    )
    print("已发送 Alt+F4")
    time.sleep(2)

    # 2. 检查窗口是否还在（有些页面不响应Alt+F4），若还在则直接点右上角X
    hwnd = win32gui.GetForegroundWindow()
    if hwnd:
        try:
            title = win32gui.GetWindowText(hwnd)
            # 窗口标题还包含U-BANK/招商银行说明没关掉
            if "U-BANK" in title or "招商银行" in title or "企业银行" in title:
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                # 点击右上角X按钮区域（窗口最右上角）
                x = right - 15
                y = top + 10
                print(f"Alt+F4未生效，直接点击关闭X({x}, {y})")
                click_at(x, y)
                time.sleep(1.5)
        except Exception as e:
            print(f"检测窗口状态异常: {e}")

    # 3. 点击第一个确认弹窗
    desktop = Desktop(backend="uia")
    _click_confirm_button(desktop, "确认", "第1次")

    time.sleep(2)

    # 4. 点击第二个确认弹窗
    _click_confirm_button(desktop, "确认", "第2次")

    print("程序已退出")


def _click_confirm_button(desktop, button_text, step_desc):
    """
    退出确认弹窗的多策略：
    1) 轮询 UIA 树找确认按钮
    2) 前台小弹窗几何点击（蓝色确认按钮在底部偏右）
    3) Enter + Tab 兜底
    """
    _ = button_text
    time.sleep(1.0)

    # 持续监听，弹窗可能延迟出现
    end_ts = time.time() + 10
    while time.time() < end_ts:
        if _try_uia_confirm_click(desktop):
            print(f"{step_desc} 已通过 UIA 点击确认")
            return

        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top
            # 中小弹窗尺寸范围
            if 320 <= width <= 1200 and 160 <= height <= 700:
                x = int(left + width * 0.56)
                y = int(top + height * 0.84)
                print(f"{step_desc} 几何点击确认: ({x}, {y}), size=({width}x{height})")
                click_at(x, y)
                time.sleep(0.5)
                press_keys((0x0D, 0), (0x0D, 2))
                return

        time.sleep(0.25)

    print(f"{step_desc} 监听超时，改用 Enter + Tab 兜底")
    press_keys((0x0D, 0), (0x0D, 2))
    time.sleep(0.3)
    for _ in range(14):
        press_keys((0x09, 0), (0x09, 2))
        time.sleep(0.12)
    press_keys((0x0D, 0), (0x0D, 2))
    print(f"{step_desc} 兜底确认序列已发送")


def _import_excel_file(file_path):
    """
    点击导入后，在Windows文件选择对话框中填入文件路径并确认，
    然后在弹出的提示框中点击"继续导入"
    流程：文件对话框 → 填路径 → 点打开 → 提示弹窗 → 点继续导入
    """
    print(f"开始导入文件: {file_path}")

    # 第一步：等待文件选择对话框出现
    print("等待文件选择对话框...")
    file_dialog_found = False
    for _ in range(20):  # 等待最多10秒
        try:
            # 检查前台窗口
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                title = win32gui.GetWindowText(hwnd)
                if "打开" in title:
                    file_dialog_found = True
                    print(f"检测到文件对话框: {title}")
                    break

            # 遍历所有顶层窗口查找对话框（#32770类名）
            def find_dialog_callback(hwnd, _):
                nonlocal file_dialog_found
                try:
                    cls_name = win32gui.GetClassName(hwnd)
                    title = win32gui.GetWindowText(hwnd)
                    rect = win32gui.GetWindowRect(hwnd)
                    w = rect[2] - rect[0]
                    h = rect[3] - rect[1]
                    if "打开" in title and w > 300 and h > 200:
                        file_dialog_found = True
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.3)
                        print(f"检测到文件对话框(枚举): {title}")
                except Exception:
                    pass

            win32gui.EnumWindows(find_dialog_callback, None)
            if file_dialog_found:
                break
        except Exception as e:
            print(f"检测对话框异常: {e}")
        time.sleep(0.5)

    if not file_dialog_found:
        print("未检测到文件选择对话框")
        return False

    time.sleep(1)

    # 第二步：在文件名输入框中填入完整文件路径，然后点"打开"
    # 先点击文件名输入框区域获取焦点，再粘贴路径
    try:
        _set_clipboard_text(file_path)
        time.sleep(0.2)

        # Ctrl+V 粘贴路径到文件名输入框
        press_keys(
            (0x11, 0), (0x56, 0),  # Ctrl down, V down
            (0x56, 2), (0x11, 2),  # V up, Ctrl up
        )
        print(f"已输入文件路径: {file_path}")
        time.sleep(1)

        # 点回车或点击"打开"按钮确认选择
        press_keys((0x0D, 0), (0x0D, 2))  # Enter
        print("已按回车确认文件选择")
        time.sleep(3)  # 给更长时间让弹窗渲染

    except Exception as e:
        print(f"填写文件路径失败: {e}")
        return False

    # 第三步：等待提示弹窗，用UIA轮询查找并点击"继续导入"
    print("等待'继续导入'弹窗按钮...")
    desktop2 = Desktop(backend="uia")
    end_ts = time.time() + 15
    clicked_continue = False
    while time.time() < end_ts:
        if _try_uia_confirm_click(desktop2):
            print("已点击'继续导入'")
            clicked_continue = True
            break
        time.sleep(0.4)

    if not clicked_continue:
        print("未找到'继续导入'按钮，尝试几何点击...")
        # 兜底：在屏幕中央偏右下的位置点击（Web弹窗通常居中）
        hwnd = win32gui.GetForegroundWindow()
        if hwnd:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            x = int(left + (right - left) * 0.58)
            y = int(top + (bottom - top) * 0.55)
            click_at(x, y)
            print(f"兜底几何点击({x}, {y})")

    time.sleep(3)

    # 第四步：等待错误信息弹窗，点击"导入正确列表"
    print("等待'导入正确列表'按钮...")
    end_ts2 = time.time() + 15
    clicked_import_ok = False
    while time.time() < end_ts2:
        if _try_uia_confirm_click(desktop2):
            print("已点击'导入正确列表'")
            clicked_import_ok = True
            break
        time.sleep(0.4)

    if not clicked_import_ok:
        print("未找到'导入正确列表'按钮")

    time.sleep(3)

    # 第五步：等待数据加载完成后，点击"经办"按钮
    print("等待数据加载完成，准备点击经办...")
    time.sleep(3)

    main_win2 = wait_for_main_window(desktop2)
    if main_win2:
        try:
            click_control_by_name(main_win2, "经办")
            print("已点击经办按钮")
            time.sleep(2)
        except Exception as e:
            print(f"点击经办失败: {e}")

    print("========== 批量制单导入+经办完成 ==========")
    return True


def main():
    """主流程：打开 -> 登录 -> 点击转账支付 -> 点击批量转账经办 -> 到达批量制单页"""

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

    # 5. 点击左侧"批量转账经办"
    click_control_by_name(main_win, "批量转账经办")
    time.sleep(3)  # 等待批量转账经办页面加载

    # 6. 验证是否成功到达批量转账经办页面
    print("验证目标页面...")
    success_flag = False
    try:
        for _ in range(10):
            windows = desktop.windows()
            for w in windows:
                title = w.window_text()
                if "批量转账经办" in title or "批量转账" in title:
                    success_flag = True
                    print(f"已到达目标页面: {title}")
                    break
            if success_flag:
                break
            time.sleep(1)
    except Exception as e:
        print(f"验证过程异常: {e}")

    if success_flag:
        print("已成功进入批量转账经办页面")

        # 7. 点击"导入"按钮
        time.sleep(2)  # 等待页面控件完全加载
        click_control_by_name(main_win, "导入")
        print("已点击导入按钮")
        time.sleep(2)

        # 8. 在文件对话框中填入Excel路径，点打开，再点继续导入
        if not _import_excel_file(IMPORT_FILE_PATH):
            print("批量导入失败，终止执行")
            sys.exit(6)

        print("========== 批量制单流程完成 ==========")
        print("已进入批量转账经办页面并点击导入")
        print("====================================")

        # 9. 退出功能
        time.sleep(1)
        _close_with_confirm(main_win)
    else:
        print("未能确认进入批量转账经办页面")
        time.sleep(1)
        _close_with_confirm(main_win)
        sys.exit(5)


if __name__ == "__main__":
    main()
