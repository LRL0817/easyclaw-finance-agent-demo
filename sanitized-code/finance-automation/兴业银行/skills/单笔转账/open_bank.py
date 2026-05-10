"""
兴业银行企业网银 - 自动登录

流程：
  [1] 打开快捷方式（若网银窗口未打开）
  [2] 检测并填写"验证网盾密码"对话框（可能不出现，使用 .env 第一项 LOGIN_PWD）
  [3] 登录页：选账号 -> 输密码(.env 第二项 CERT_PWD) -> 勾同意 -> 点登录
  [4] 登录后若再次出现"验证网盾密码"对话框，同样用 LOGIN_PWD 填写
"""
import os
import sys
import time
import subprocess
import pyautogui
import pygetwindow as gw
import win32api
import win32gui
from pywinauto import Desktop


def _find_nearby_edit(label_ctrl):
    """从标签的父容器里找距离最近的 Edit/ComboBox 控件"""
    try:
        parent = label_ctrl.parent()
        if parent is None:
            return None
        ly = label_ctrl.rectangle().top
        candidates = []
        for sib in parent.children():
            try:
                if not sib.is_visible():
                    continue
                stype = str(sib.element_info.control_type) if hasattr(sib.element_info, "control_type") else ""
                scls = sib.element_info.class_name if hasattr(sib.element_info, "class_name") else ""
                if any(kw in scls for kw in ["Edit", "edit"]) or any(kw in stype for kw in ["Edit", "Document", "ComboBox", "Combo"]):
                    sr = sib.rectangle()
                    candidates.append((sib, abs(sr.top - ly)))
            except Exception:
                pass
        if not candidates:
            return None
        candidates.sort(key=lambda x: x[1])
        return candidates[0][0]
    except Exception:
        return None


def _find_label(main_win, label_text):
    """递归查找标签控件（排除 Edit/Document），名称等于或刚好包含 label_text"""
    result = {"ctrl": None}

    def _walk(control):
        if result["ctrl"]:
            return True
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
            is_edit = any(kw in cls for kw in ["Edit", "edit"]) or any(kw in ctrl_type for kw in ["Edit", "Document"])
            if not is_edit and name and control.is_visible():
                s = name.strip().lstrip("*").strip()
                if s == label_text or (label_text in s and len(s) <= len(label_text) + 4):
                    result["ctrl"] = control
                    return True
        except Exception:
            pass
        try:
            for child in control.children():
                if _walk(child):
                    return True
        except Exception:
            pass
        return False

    _walk(main_win)
    return result["ctrl"]


_PS_SET_CLIPBOARD = (
    # Windows PowerShell 5.1 默认按系统 ANSI 解 stdin，必须改成 UTF-8 才能正确传中文。
    "[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false); "
    "$txt = [Console]::In.ReadToEnd(); "
    "Set-Clipboard -Value $txt"
)


def _set_clipboard(value):
    """通过 stdin 传 UTF-8 给 PowerShell；失败 raise 避免粘贴旧剪贴板。"""
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-Command", _PS_SET_CLIPBOARD],
        input=value or "", text=True, encoding="utf-8",
        capture_output=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"Set-Clipboard 失败 rc={proc.returncode}: {proc.stderr.strip()}")


def fill_field_by_label(main_win, label_text, value, use_clipboard=False):
    """按标签找邻近输入框，点击聚焦，清空，输入"""
    label = _find_label(main_win, label_text)
    if not label:
        print(f"    未找到标签 [{label_text}]")
        return False
    edit = _find_nearby_edit(label)
    if not edit:
        print(f"    [{label_text}] 找不到输入框")
        return False
    try:
        r = edit.rectangle()
        pyautogui.click(r.mid_point().x, r.mid_point().y)
        time.sleep(0.4)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        pyautogui.press("delete")
        time.sleep(0.15)
        if use_clipboard:
            _set_clipboard(value)
            time.sleep(0.2)
            pyautogui.hotkey("ctrl", "v")
        else:
            pyautogui.write(value, interval=0.05)
        print(f"    已填 [{label_text}]: {value}")
        time.sleep(0.4)
        return True
    except Exception as e:
        print(f"    填 [{label_text}] 失败: {e}")
        return False


def _panel_has_bank(main_win, bank_name):
    """检测银行 logo 面板是否已弹出（通过能否找到银行名判定）"""
    found = {"ok": False}

    def _walk(control):
        if found["ok"]:
            return
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            if name and name.strip() == bank_name and control.is_visible():
                found["ok"] = True
                return
        except Exception:
            pass
        try:
            for child in control.children():
                _walk(child)
        except Exception:
            pass

    _walk(main_win)
    return found["ok"]


def dump_bank_debug_controls(main_win, stage, keywords=None, limit=80):
    """打印收款行相关的可见 UIA 控件，辅助定位下拉/查询/填入问题。"""
    if main_win is None:
        print(f"    [收款行调试:{stage}] main_win=None")
        return
    if keywords is None:
        keywords = [
            "收款行", "开户", "银行", "招商", "福州", "仓山",
            "查询", "搜索", "关键", "行号", "填入", "确定", "选择",
        ]
    rows = []

    def _walk(control, depth=0):
        try:
            if not control.is_visible():
                return
            info = control.element_info
            name = info.name if hasattr(info, "name") and info.name else ""
            ctrl_type = str(info.control_type) if hasattr(info, "control_type") else ""
            cls = info.class_name if hasattr(info, "class_name") else ""
            if name and any(k in name for k in keywords):
                r = control.rectangle()
                rows.append((depth, name.strip(), ctrl_type, cls, r.left, r.top, r.right, r.bottom))
        except Exception:
            pass
        try:
            for child in control.children():
                _walk(child, depth + 1)
        except Exception:
            pass

    _walk(main_win)
    print(f"    [收款行调试:{stage}] 匹配控件 {len(rows)} 个")
    for depth, name, ctrl_type, cls, left, top, right, bottom in rows[:limit]:
        print(f"      depth={depth} name={name!r} type={ctrl_type} class={cls} rect=({left},{top},{right},{bottom})")
    if len(rows) > limit:
        print(f"      ... 还有 {len(rows) - limit} 个控件未打印")


def is_bank_panel_open(main_win):
    """判断收款行搜索面板是否仍然展开。"""
    found = {"ok": False}

    def _walk(control):
        if found["ok"]:
            return
        try:
            if not control.is_visible():
                return
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            if name and ("查询结果" in name or "请输入关键词或完整行号查询" in name):
                found["ok"] = True
                return
        except Exception:
            pass
        try:
            for child in control.children():
                _walk(child)
        except Exception:
            pass

    _walk(main_win)
    return found["ok"]


def click_field_dropdown(main_win, label_text, verify_bank=None, max_retry=3):
    label = _find_label(main_win, label_text)
    if not label:
        print(f"    未找到标签 [{label_text}]")
        return False
    # 以标签位置为基准：标签右侧 200 像素是下拉框中心区域
    try:
        lr = label.rectangle()
    except Exception:
        return False
    cx = lr.right + 200
    cy = lr.mid_point().y
    print(f"    标签 [{label_text}] rect=({lr.left},{lr.top},{lr.right},{lr.bottom}) -> 点 ({cx}, {cy})")
    screenshot("bank_debug_01_下拉前")
    dump_bank_debug_controls(main_win, "下拉前")
    for attempt in range(max_retry):
        try:
            pyautogui.click(cx, cy)
            print(f"    点击下拉 尝试{attempt + 1}")
            time.sleep(1.2)
            screenshot(f"bank_debug_02_下拉后_尝试{attempt + 1}")
            dump_bank_debug_controls(main_win, f"下拉后_尝试{attempt + 1}")
            if not verify_bank or _panel_has_bank(main_win, verify_bank):
                return True
            print(f"    面板未出现 [{verify_bank}]，重试")
        except Exception as e:
            print(f"    点击异常: {e}")
    return False


def click_match_fill_button(main_win, timeout=5):
    """找 '智能匹配结果' 行旁边的 '填入' 按钮并点击（策略1：自动匹配收款行）"""
    end = time.time() + timeout
    attempt = 0
    while time.time() < end:
        attempt += 1
        target = {"c": None}

        def _walk(control):
            if target["c"]:
                return
            try:
                name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
                if name and name.strip() == "填入" and control.is_visible():
                    r = control.rectangle()
                    if r.width() > 3 and r.height() > 3:
                        target["c"] = control
                        return
            except Exception:
                pass
            try:
                for child in control.children():
                    _walk(child)
            except Exception:
                pass

        _walk(main_win)
        if target["c"]:
            r = target["c"].rectangle()
            print(f"    找到 [填入] 按钮 rect=({r.left},{r.top},{r.right},{r.bottom})")
            screenshot("bank_debug_填入按钮_点击前")
            pyautogui.click(r.mid_point().x, r.mid_point().y)
            print(f"    已点 [填入] @ ({r.mid_point().x}, {r.mid_point().y})")
            time.sleep(0.6)
            screenshot("bank_debug_填入按钮_点击后")
            return True
        if attempt in (1, 4, 8):
            print(f"    第{attempt}次未找到 [填入]，打印可见控件快照")
            dump_bank_debug_controls(main_win, f"找填入_第{attempt}次")
        time.sleep(0.4)
    print("    未发现 [智能匹配 -> 填入]")
    return False


def _find_option_in_tree(main_win, option_text):
    """在下拉展开后的 UIA 树中查找目标选项控件"""
    found = {"c": None}

    def _walk(control):
        if found["c"]:
            return
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            if name and name.strip() == option_text and control.is_visible():
                r = control.rectangle()
                if r.height() > 5 and r.width() > 10:
                    found["c"] = control
                    return
        except Exception:
            pass
        try:
            for child in control.children():
                _walk(child)
        except Exception:
            pass

    _walk(main_win)
    return found["c"]


def _find_nearby_combo(label_ctrl):
    """从标签的父容器里找距离最近的下拉/组合框控件（Electron 应用可能注册为 Button 或 Edit）"""
    try:
        parent = label_ctrl.parent()
        if parent is None:
            return None
        ly = label_ctrl.rectangle().top
        candidates = []
        for sib in parent.children():
            try:
                if not sib.is_visible():
                    continue
                scls = sib.element_info.class_name if hasattr(sib.element_info, "class_name") else ""
                stype = str(sib.element_info.control_type) if hasattr(sib.element_info, "control_type") else ""
                # 匹配多种可能的下拉控件类型（含 Electron 自定义实现）
                is_combo = (
                    any(kw in stype for kw in ["ComboBox", "Combo", "List", "Select"]) or
                    any(kw in stype for kw in ["Button"]) and ("arrow" in scls.lower() or "drop" in scls.lower()) or
                    any(kw in stype for kw in ["Edit", "Document"])
                )
                if is_combo:
                    sr = sib.rectangle()
                    candidates.append((sib, abs(sr.top - ly)))
            except Exception:
                pass
        if not candidates:
            return None
        candidates.sort(key=lambda x: x[1])
        return candidates[0][0]
    except Exception:
        return None


def select_combobox_option(main_win, label_text, option_text):
    """打开下拉，在 UIA 树里找目标项点击（找不到则滚动列表重试）"""
    label = _find_label(main_win, label_text)
    if not label:
        print(f"    未找到标签 [{label_text}]")
        return False

    # 通过标签找旁边的 ComboBox 控件
    combo = _find_nearby_combo(label)

    if combo:
        r = combo.rectangle()
        cx, cy = r.mid_point().x, r.mid_point().y
        print(f"    找到控件 @ ({cx}, {cy}) type={combo.element_info.control_type}")

        # 先用 pywinauto 让控件获得焦点
        try:
            combo.set_focus()
            time.sleep(0.5)
            print(f"    已用 pywinauto set_focus")
        except Exception as e:
            print(f"    set_focus 失败: {e}，回退鼠标点击聚焦")
            pyautogui.click(cx, cy)
            time.sleep(0.5)

        # 用 F4 或 Alt+Down 打开下拉（标准 Windows 下拉快捷键）
        for hotkey in [("f4",), ("alt", "down")]:
            try:
                if len(hotkey) == 1:
                    pyautogui.press(hotkey[0])
                else:
                    pyautogui.hotkey(*hotkey)
                time.sleep(1.5)
                print(f"    已按 {hotkey} 打开下拉 [{label_text}]")
                break
            except Exception:
                continue

        # 检查是否已展开
        opt_ctrl = _find_option_in_tree(main_win, option_text)
        if not opt_ctrl:
            # 最后尝试 click_input
            try:
                combo.click_input()
                time.sleep(1.5)
                print(f"    回退 click_input 打开下拉")
            except Exception:
                pass
    else:
        lr = label.rectangle()
        cx = lr.right + 200
        cy = lr.mid_point().y
        pyautogui.click(cx, cy)
        time.sleep(1.5)
        print(f"    未找到 ComboBox 控件，使用偏移坐标 ({cx}, {cy})")

    # 先在当前可见区域查找
    opt_ctrl = _find_option_in_tree(main_win, option_text)
    if opt_ctrl:
        r = opt_ctrl.rectangle()
        pyautogui.click(r.mid_point().x, r.mid_point().y)
        print(f"    已选 [{label_text}]: {option_text}")
        return True

    # 当前视野没找到，用键盘向下翻页查找（最多按 15 次 Down）
    print(f"    当前视野未找到 [{option_text}]，开始键盘翻页查找...")
    for key_i in range(15):
        pyautogui.press("down")
        time.sleep(0.3)
        opt_ctrl = _find_option_in_tree(main_win, option_text)
        if opt_ctrl:
            r = opt_ctrl.rectangle()
            pyautogui.click(r.mid_point().x, r.mid_point().y)
            print(f"    已选 [{label_text}]: {option_text} (按键{key_i + 1}次后)")
            return True

    # 键盘也找不到，最后尝试鼠标滚轮滚动
    print(f"    键盘未找到，尝试滚轮滚动...")
    scroll_x = cx
    scroll_y = cy + 80
    for scroll_i in range(5):
        pyautogui.moveTo(scroll_x, scroll_y)
        pyautogui.scroll(-300)
        time.sleep(0.6)
        opt_ctrl = _find_option_in_tree(main_win, option_text)
        if opt_ctrl:
            r = opt_ctrl.rectangle()
            pyautogui.click(r.mid_point().x, r.mid_point().y)
            print(f"    已选 [{label_text}]: {option_text} (滚动{scroll_i + 1}次后)")
            return True

    print(f"    下拉中未找到 [{option_text}]")
    return False


def pick_branch_option(main_win, keyword, target_name):
    """在收款行弹出面板的搜索框里粘贴关键字，然后点结果里匹配的支行"""
    print(f"    收款行搜索 keyword={keyword!r}, target={target_name!r}")
    screenshot("bank_debug_03_进入支行搜索函数")
    dump_bank_debug_controls(main_win, "进入支行搜索函数")
    # 找搜索框
    search = None

    def _find_search(control):
        nonlocal search
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            ctrl_type = str(control.element_info.control_type) if hasattr(control.element_info, "control_type") else ""
            cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
            is_edit = any(kw in cls for kw in ["Edit", "edit"]) or any(kw in ctrl_type for kw in ["Edit", "Document"])
            if is_edit and control.is_visible() and ("关键" in name or "行号" in name or "查询" in name):
                search = control
                return True
        except Exception:
            pass
        try:
            for child in control.children():
                if _find_search(child):
                    return True
        except Exception:
            pass
        return False

    _find_search(main_win)
    if not search:
        print("    未找到收款行搜索框")
        screenshot("bank_debug_04_未找到搜索框")
        dump_bank_debug_controls(main_win, "未找到搜索框")
        return False
    try:
        r = search.rectangle()
        print(f"    找到搜索框 name={search.element_info.name!r} type={search.element_info.control_type} rect=({r.left},{r.top},{r.right},{r.bottom})")
        pyautogui.click(r.mid_point().x, r.mid_point().y)
        time.sleep(0.3)
        screenshot("bank_debug_05_点击搜索框")
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        pyautogui.press("delete")
        time.sleep(0.1)
        _set_clipboard(keyword)
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.8)
        screenshot("bank_debug_06_输入搜索关键字")
        dump_bank_debug_controls(main_win, "输入搜索关键字后")
        # 点击 '查询' 按钮
        query_btn = {"c": None}

        def _walk_q(control):
            if query_btn["c"]:
                return
            try:
                name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
                if name and name.strip() == "查询" and control.is_visible():
                    r = control.rectangle()
                    if r.width() > 3 and r.height() > 3:
                        query_btn["c"] = control
                        return
            except Exception:
                pass
            try:
                for child in control.children():
                    _walk_q(child)
            except Exception:
                pass

        _walk_q(main_win)
        if query_btn["c"]:
            r = query_btn["c"].rectangle()
            print(f"    找到 [查询] rect=({r.left},{r.top},{r.right},{r.bottom})")
            screenshot("bank_debug_07_查询按钮_点击前")
            pyautogui.click(r.mid_point().x, r.mid_point().y)
            print(f"    已点 [查询]")
        else:
            print("    未找到 [查询] 按钮，回车查询")
            pyautogui.press("enter")
        time.sleep(1.8)
        screenshot("bank_debug_08_查询后")
        dump_bank_debug_controls(main_win, "查询后")
    except Exception as e:
        print(f"    搜索框输入失败: {e}")
        screenshot("bank_debug_搜索框输入失败")
        return False

    # 查询完成后，固定点击查询结果的第一个。
    time.sleep(0.5)
    screenshot("bank_debug_09_点击第一个结果前")
    pyautogui.click(930, 480)
    print("    已点击 [固定位置第一个结果] (930, 480)")
    time.sleep(1.0)
    screenshot("bank_debug_10_点击第一个结果后")
    dump_bank_debug_controls(main_win, "点击第一个结果后")
    if is_bank_panel_open(main_win):
        print("    收款行面板仍展开，固定点未选中结果")
        return False
    print("    固定结果点已完成收款行选择，无需点击 [填入]")
    return True


def click_control_by_name(main_win, target_name, timeout=8):
    """在 UIA 树中按名称模糊查找控件并点击（参考招行 ubank_common.click_control_by_name）"""

    def _find_and_click(control):
        try:
            name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
            if name and name.strip() == target_name.strip():
                rect = control.rectangle()
                if control.is_visible() and rect.width() > 3 and rect.height() > 3:
                    cx, cy = rect.mid_point().x, rect.mid_point().y
                    pyautogui.click(cx, cy)
                    print(f"  已点击 [{target_name}] @ ({cx}, {cy})")
                    return True
        except Exception:
            pass
        try:
            for child in control.children():
                if _find_and_click(child):
                    return True
        except Exception:
            pass
        return False

    end = time.time() + timeout
    while time.time() < end:
        if _find_and_click(main_win):
            return True
        time.sleep(0.3)
    print(f"  未找到控件: {target_name}")
    return False


def click_next_step(main_win, timeout=5):
    """点击单笔转账表单底部的“下一步”按钮。"""
    if main_win is not None:
        try:
            main_win.set_focus()
        except Exception:
            pass
        if click_control_by_name(main_win, "下一步", timeout=timeout):
            return True

    # 兜底坐标：基于 1920x1080 下单笔转账页面底部蓝色“下一步”按钮。
    pyautogui.click(1037, 811)
    print("  已点击 [下一步] @ (1037, 811)")
    return True


def click_submit(main_win, timeout=5):
    """点击信息确认页底部的“提交”按钮。"""
    if main_win is not None:
        try:
            main_win.set_focus()
        except Exception:
            pass
        if click_control_by_name(main_win, "提交", timeout=timeout):
            return True

    # 兜底坐标：基于 1920x1080 信息确认页底部蓝色“提交”按钮。
    pyautogui.click(1139, 889)
    print("  已点击 [提交] @ (1139, 889)")
    return True


def scroll_to_bank_section():
    """把表单滚动位置归一化到收款行区域，避免相对滚动导致固定点漂移。"""
    pyautogui.moveTo(960, 500)
    pyautogui.scroll(2000)
    time.sleep(0.8)
    screenshot("06_5_回到顶部")
    pyautogui.scroll(-500)
    time.sleep(0.8)
    screenshot("06_5_下翻后")


_CIB_ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
ENV_PATH = os.path.join(_CIB_ROOT, ".env")
SHORTCUT_PATH = os.environ.get(
    "CIB_SHORTCUT_PATH",
    os.path.join(os.path.expanduser("~"), "Desktop", "兴业银行企业网银.lnk"),
)
SCREENSHOT_DIR = os.path.join(_CIB_ROOT, "screenshots")


def screenshot(name):
    if not os.path.exists(SCREENSHOT_DIR):
        os.makedirs(SCREENSHOT_DIR)
    path = os.path.join(SCREENSHOT_DIR, f"{name}.png")
    pyautogui.screenshot(path)
    print(f"  [截图] {path}")


def load_config():
    config = {}
    with open(ENV_PATH, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith("#") or not line or "=" not in line:
                continue
            key, value = line.split("=", 1)
            config[key.strip()] = value.strip()
    return config


def find_window():
    for w in gw.getAllWindows():
        if not w.title:
            continue
        t = str(w.title)
        # 排除网盾助手等后台小窗口（坐标异常或尺寸过小说明不是主窗口）
        if "兴业管家网盾助手" in t or "网盾助手" in t:
            continue
        # 过滤掉隐藏/最小化到托盘的窗口（left=-32000 是 Windows 隐藏窗口特征）
        if w.left < -1000 or w.width < 200 or w.height < 100:
            continue
        if "兴业" in t:
            return w
    return None


def find_ukey_dialog():
    """查找 '验证网盾密码' 对话框窗口（含子窗口）。返回带 left/top/width/height/title/activate 的对象。"""
    # 先找顶层窗口
    for w in gw.getAllWindows():
        if w.title and "验证网盾密码" in str(w.title):
            return w

    # 回退：用 win32 枚举所有顶层和子窗口
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


def locate_ukey_password_box():
    """Locate the UKey password input by looking at the current screenshot."""
    try:
        img = pyautogui.screenshot().convert("RGB")
        width, height = img.size
        px = img.load()

        def is_title_blue(rgb):
            r, g, b = rgb
            return r < 80 and 95 <= g <= 180 and 130 <= b <= 230

        candidates = []
        for y in range(80, height - 120):
            start = None
            for x in range(0, width):
                if is_title_blue(px[x, y]):
                    if start is None:
                        start = x
                elif start is not None:
                    if x - start >= 300:
                        candidates.append((y, start, x))
                    start = None
            if start is not None and width - start >= 300:
                candidates.append((y, start, width))

        if not candidates:
            return None

        y, left, right = sorted(candidates, key=lambda item: item[2] - item[1], reverse=True)[0]
        same_bar = [item for item in candidates if abs(item[1] - left) < 8 and abs(item[2] - right) < 8]
        top = min(item[0] for item in same_bar)
        bottom = max(item[0] for item in same_bar)

        runs = []
        search_left = left + 160
        search_right = right - 15
        search_top = bottom + 15
        search_bottom = min(bottom + 95, height)
        for yy in range(search_top, search_bottom):
            start = None
            for xx in range(search_left, search_right):
                r, g, b = px[xx, yy]
                is_white = r >= 245 and g >= 245 and b >= 245
                if is_white:
                    if start is None:
                        start = xx
                elif start is not None:
                    if xx - start >= 150:
                        runs.append((yy, start, xx))
                    start = None
            if start is not None and search_right - start >= 150:
                runs.append((yy, start, search_right))

        if not runs:
            return None

        yy, input_left, input_right = sorted(runs, key=lambda item: item[2] - item[1], reverse=True)[0]
        same_input = [item for item in runs if abs(item[1] - input_left) < 8 and abs(item[2] - input_right) < 8]
        input_top = min(item[0] for item in same_input)
        input_bottom = max(item[0] for item in same_input)
        return ((input_left + input_right) // 2, (input_top + input_bottom) // 2)
    except Exception:
        return None


def focus_ukey_password_field(ukey_win):
    """Bring the UKey dialog forward and click the password edit box."""
    try:
        ukey_win.activate()
    except Exception:
        pass
    time.sleep(0.3)
    try:
        located = locate_ukey_password_box()
        if located:
            x, y = located
            print(f"  聚焦网盾密码框 image=({x},{y})")
            pyautogui.click(x, y)
            time.sleep(0.15)
            pyautogui.click(x, y)
            time.sleep(0.2)
            return

        raw_x = ukey_win.left + int(ukey_win.width * 0.68)
        raw_y = ukey_win.top + int(ukey_win.height * 0.22)
        screen_w, screen_h = pyautogui.size()
        metric_w = win32api.GetSystemMetrics(0)
        metric_h = win32api.GetSystemMetrics(1)
        scale_x = screen_w / metric_w if metric_w else 1
        scale_y = screen_h / metric_h if metric_h else 1
        x = int(raw_x * scale_x)
        y = int(raw_y * scale_y)
        print(f"  聚焦网盾密码框 raw=({raw_x},{raw_y}) click=({x},{y}) scale=({scale_x:.3f},{scale_y:.3f})")
        pyautogui.click(x, y)
        time.sleep(0.15)
        pyautogui.click(x, y)
        time.sleep(0.2)
    except Exception:
        pass


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

    print(f"  检测到对话框: '{ukey_win.title}' @ ({ukey_win.left}, {ukey_win.top}) {ukey_win.width}x{ukey_win.height}")
    time.sleep(0.8)
    focus_ukey_password_field(ukey_win)

    # 直接输入密码（对话框打开时输入框已聚焦），不点击输入框
    pyautogui.write(ukey_pwd, interval=0.08)
    print(f"  已输入网盾密码 (长度:{len(ukey_pwd)})")
    time.sleep(0.4)
    screenshot(f"网盾_输密码{tag}")

    # 按回车确认
    pyautogui.press("enter")
    time.sleep(1.5)
    screenshot(f"网盾_确认{tag}")
    if find_ukey_dialog():
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("backspace")
        print("  错误: 网盾密码框仍未关闭，停止后续流程")
        sys.exit(1)
    return True


def main():
    config = load_config()
    ukey_pwd = config.get("LOGIN_PWD", "")   # .env 第一项：网盾密码
    cert_pwd = config.get("CERT_PWD", "")    # .env 第二项：登录页密码

    print("=" * 50)
    print("兴业银行网银自动登录")
    print("=" * 50)

    # ========== [1] 打开应用 ==========
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
    try:
        win.activate()
    except Exception as e:
        print(f"  窗口激活异常，继续执行: {e}")
    if not win.visible:
        win.show()
    time.sleep(1)
    screenshot("01_开始前")

    # ========== [2] 网盾密码（启动阶段，可能不出现） ==========
    print("\n[2] 检测网盾密码对话框...")
    time.sleep(2)

    def type_with_capslock(s, interval=0.1):
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

    ukey_win = None
    for i in range(5):
        ukey_win = find_ukey_dialog()
        if ukey_win:
            break
        time.sleep(1)

    if ukey_win:
        print(f"  检测到网盾密码对话框: '{ukey_win.title}'，开始输入密码...")
        focus_ukey_password_field(ukey_win)
        type_with_capslock(ukey_pwd)
        print(f"  已输入网盾密码 (长度:{len(ukey_pwd)})")
        time.sleep(0.4)
        screenshot("网盾_输密码_启动")
        pyautogui.press("enter")
        time.sleep(1.5)
        screenshot("网盾_确认_启动")
        if find_ukey_dialog():
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")
            print("  错误: 网盾密码框仍未关闭，停止后续流程")
            sys.exit(1)
    else:
        print("  未检测到网盾密码对话框，跳过")

    # ========== [3] 登录页流程 ==========
    print("\n[3] 登录页流程...")
    # 坐标基于 1920x1080 完整截图测量（登录卡片位于右侧）
    coords = {
        "dropdown":  (1331, 395),
        "password":  (1331, 453),
        "checkbox":  (1132, 489),
        "login_btn": (1331, 577),
    }

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

    print("  [3.5] 点击登录按钮...")
    pyautogui.click(*coords["login_btn"])
    time.sleep(2)
    screenshot("03_5_点登录")

    # ========== [4] 登录后若再次出现网盾密码对话框 ==========
    print("\n[4] 处理登录后可能出现的网盾密码对话框...")
    handle_ukey_dialog(ukey_pwd, wait_seconds=10, tag="_登录后")

    # ========== [5] 进入单笔转账（UIA 按控件名识别） ==========
    print("\n[5] 导航到 转账付款 -> 单笔转账 ...")
    time.sleep(2)
    desktop = Desktop(backend="uia")
    main_win = None
    for w in desktop.windows():
        try:
            if w.window_text() and "兴业" in w.window_text():
                main_win = w
                break
        except Exception:
            pass
    if main_win is None:
        print("  UIA 未找到兴业主窗口，回退坐标点击")
        pyautogui.click(368, 167)
        time.sleep(1.2)
        screenshot("05_1_转账付款")
        pyautogui.click(162, 289)
    else:
        print("  [5.1] 点击 '转账付款' ...")
        click_control_by_name(main_win, "转账付款")
        time.sleep(1.2)
        screenshot("05_1_转账付款")

        print("  [5.2] 点击 '单笔转账' (坐标) ...")
        pyautogui.click(162, 289)
    time.sleep(2)
    screenshot("05_2_单笔转账")

    # ========== [6] 填写转账表单 ==========
    print("\n[6] 填写转账表单...")
    time.sleep(6)

    desktop_uia = Desktop(backend="uia")
    main_win = None
    for w in desktop_uia.windows():
        try:
            if w.window_text() and "兴业" in w.window_text():
                main_win = w
                break
        except Exception:
            pass
    if main_win is None:
        print("  未找到兴业主窗口，跳过填表")
    else:
        AMOUNT = "0.01"
        ACCT_NO = "000000000000000000"
        ACCT_NAME = "陆零壹（福州）网络信息有限公司"
        BANK = "招商银行"
        BRANCH_KEYWORD = "福州仓山"
        BRANCH_FULL = "招商银行股份有限公司福州仓山支行"
        PURPOSE_OPTION = "其他合法用途"
        PURPOSE_TEXT = "陆零壹结算2026年1月-2026年2月分佣，金额总计0.01元，请审批。"

        print("  [6.1] 金额")
        fill_field_by_label(main_win, "金额", AMOUNT, use_clipboard=False)
        screenshot("06_1_金额")

        print("  [6.2] 选 单位账户")
        click_control_by_name(main_win, "单位账户", timeout=3)
        time.sleep(0.5)

        print("  [6.3] 收款账号")
        fill_field_by_label(main_win, "收款账号", ACCT_NO, use_clipboard=False)

        print("  [6.4] 收款户名")
        fill_field_by_label(main_win, "收款户名", ACCT_NAME, use_clipboard=True)
        screenshot("06_4_户名")

        print("  [6.5] 收款行")
        scroll_to_bank_section()
        # 策略1：智能匹配结果的 '填入' 按钮
        if click_match_fill_button(main_win, timeout=3):
            print("    策略1 成功：使用智能匹配")
            time.sleep(1.0)
        else:
            print("    策略2：面板搜索")
            click_field_dropdown(main_win, "收款行", verify_bank=BANK, max_retry=3)
            time.sleep(0.8)
            screenshot("06_5a_面板")
            # 只输入开户银行前4个字
            if not pick_branch_option(main_win, BANK, BRANCH_FULL):
                print("  错误: 收款行未选中，停止后续下一步/提交")
                screenshot("06_5_收款行未选中_停止")
                sys.exit(1)
        screenshot("06_5b_选完支行")
        # 刷新 UIA 树快照
        time.sleep(1)
        for w in desktop_uia.windows():
            try:
                if w.window_text() and "兴业" in w.window_text():
                    main_win = w
                    break
            except Exception:
                pass

        print("  [6.6] 用途 (直接填写)")
        fill_field_by_label(main_win, "用途", PURPOSE_TEXT, use_clipboard=True)
        screenshot("06_6_用途")

        screenshot("06_全部填完")

        print("  [6.7] 点击下一步")
        click_next_step(main_win)
        time.sleep(2)
        screenshot("06_7_下一步")

        print("  [6.8] 点击提交")
        click_submit(main_win)
        time.sleep(2)
        screenshot("06_8_提交")

    # 退出前全量截图，便于排查表单填写结果
    time.sleep(1)
    screenshot("06_退出前")

    # ========== [7] 关闭网银 ==========
    # print("\n[7] 关闭网银 (点击右上角 X)...")
    # time.sleep(1)
    # try:
    #     if main_win is not None:
    #         main_win.set_focus()
    # except Exception:
    #     pass
    # time.sleep(0.3)
    # # 右上角关闭按钮 (基于 1920x1080)
    # pyautogui.click(1790, 46)
    # time.sleep(1.5)
    # # 若弹出确认框，回车确认
    # pyautogui.press("enter")
    # time.sleep(1.2)
    # screenshot("07_退出")
    print("\n[7] 已按要求跳过关闭网银，窗口保持打开")
    screenshot("07_保留窗口")

    print("\n" + "=" * 50)
    print("执行完毕!")
    print("=" * 50)


if __name__ == "__main__":
    main()
