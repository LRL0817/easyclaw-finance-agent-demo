# -*- coding: utf-8 -*-
"""
诊断脚本：扫描弹窗内所有控件，找出真正的"确认"按钮
"""
import time
from pywinauto import Desktop

desktop = Desktop(backend="uia")
print("等待5秒，请确保确认弹窗已弹出...")
time.sleep(5)

target_controls = []

def _scan_control(control, depth=0):
    """递归扫描所有控件，收集包含'确认'的"""
    try:
        name = control.element_info.name if hasattr(control.element_info, "name") and control.element_info.name else ""
        ctrl_type = control.element_info.control_type if hasattr(control.element_info, "control_type") else ""
        cls = control.element_info.class_name if hasattr(control.element_info, "class_name") else ""
        auto_id = control.element_info.auto_id if hasattr(control.element_info, "auto_id") else ""

        rect = control.rectangle()
        visible = control.is_visible()

        # 收集所有包含"确认"或"取消"的控件（用于对比）
        if "确认" in name or "取消" in name or "确 认" in name:
            target_controls.append({
                "depth": depth,
                "name": name,
                "type": ctrl_type,
                "class": cls,
                "auto_id": auto_id,
                "visible": visible,
                "width": rect.width(),
                "height": rect.height(),
                "x": rect.mid_point().x,
                "y": rect.mid_point().y,
            })
    except Exception as e:
        pass

    try:
        for child in control.children():
            _scan_control(child, depth + 1)
    except Exception:
        pass


print("开始扫描所有窗口...")
for w in desktop.windows():
    title = w.window_text()
    print(f"\n=== 窗口: {title} ===")
    _scan_control(w)

print(f"\n\n========== 共找到 {len(target_controls)} 个目标控件 ==========\n")
for i, c in enumerate(target_controls):
    print(f"[{i}] depth={c['depth']}")
    print(f"    name     = '{c['name']}'")
    print(f"    type     = {c['type']}")
    print(f"    class    = {c['class']}")
    print(f"    auto_id  = '{c['auto_id']}'")
    print(f"    visible  = {c['visible']}")
    print(f"    size     = {c['width']} x {c['height']}")
    print(f"    position = ({c['x']}, {c['y']})")
    print()
