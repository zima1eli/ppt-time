"""Opt-in named Windows desktop commands and low-frequency meeting-room telemetry."""
import ctypes
from ctypes import wintypes
import os
import platform
import subprocess
import sys
import time
from pathlib import Path
from PySide6.QtCore import QObject, QTimer
from PySide6.QtWidgets import QCheckBox, QDialog, QLabel, QVBoxLayout


ACTIONS = {"desk_show": (0x5B, 0x44), "desk_tasks": (0x5B, 0x09),
           "desk_switch": (0x12, 0x09), "desk_project": (0x5B, 0x50),
           "desk_close": (0x12, 0x73), "desk_lock": (), "desk_files": ()}


class _Memory(ctypes.Structure):
    _fields_ = [("length", wintypes.DWORD), ("load", wintypes.DWORD)] + [
        (name, ctypes.c_ulonglong) for name in ("total", "avail", "page", "free_page", "virt", "free_virt", "ext")]


class _Power(ctypes.Structure):
    _fields_ = [(name, ctypes.c_ubyte) for name in ("ac", "flag", "percent", "state")] + [
        ("remaining", wintypes.DWORD), ("capacity", wintypes.DWORD)]


class _Adapter(ctypes.Structure):
    _fields_ = [("size", wintypes.DWORD), ("name", ctypes.c_wchar * 32),
                ("description", ctypes.c_wchar * 128), ("flags", wintypes.DWORD),
                ("device_id", ctypes.c_wchar * 128), ("key", ctypes.c_wchar * 128)]


def _cpu_times():
    idle, kernel, user = (wintypes.FILETIME() for _ in range(3))
    if not ctypes.windll.kernel32.GetSystemTimes(ctypes.byref(idle), ctypes.byref(kernel), ctypes.byref(user)):
        return None
    def value(ft):
        return (ft.dwHighDateTime << 32) + ft.dwLowDateTime
    return value(idle), value(kernel) + value(user)


def _adapters():
    names = []
    for index in range(32):
        device = _Adapter()
        device.size = ctypes.sizeof(device)
        if not ctypes.windll.user32.EnumDisplayDevicesW(None, index, ctypes.byref(device), 0):
            break
        if device.flags & 1 and device.description and device.description not in names:
            names.append(device.description)
    return "、".join(names)


class OperatorTools(QObject):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.enabled = False
        self.dialog = None
        self.cache = {}
        self.sampled = -1e10
        self.times = None
        self.graphics = None

    def metrics(self):
        now = time.monotonic()
        if now - self.sampled < 2:
            return self.cache
        self.sampled = now
        data = {"os": platform.system()}
        if sys.platform == "win32":
            if self.graphics is None:
                self.graphics = _adapters()
            data["graphics"] = self.graphics
            current = _cpu_times()
            if self.times and current and current[1] > self.times[1]:
                percent = round(100 * (1 - (current[0] - self.times[0]) / (current[1] - self.times[1])))
                data["cpu_percent"] = min(100, max(0, percent))
            self.times = current
            memory = _Memory()
            memory.length = ctypes.sizeof(memory)
            if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory)):
                data["memory_percent"] = memory.load
            power = _Power()
            if ctypes.windll.kernel32.GetSystemPowerStatus(ctypes.byref(power)):
                data["power"] = {0:"电池", 1:"外接电源"}.get(power.ac, "未知")
                if power.percent <= 100 and power.flag != 128:
                    data["battery_percent"] = power.percent
        self.cache = data
        return data

    def execute(self, action):
        if not self.enabled or sys.platform != "win32" or action not in ACTIONS:
            return False
        user32 = ctypes.windll.user32
        if action == "desk_lock":
            return bool(user32.LockWorkStation())
        if action == "desk_files":
            subprocess.Popen([str(Path(os.environ["SystemRoot"]) / "explorer.exe")])
            return True
        # The operator can trigger only these named shortcuts for the current session.
        for key in ACTIONS[action]:
            user32.keybd_event(key, 0, 0, 0)
        for key in reversed(ACTIONS[action]):
            user32.keybd_event(key, 0, 2, 0)
        return True

    def open_settings(self):
        if self.dialog is None:
            self.dialog = QDialog()
            self.dialog.setWindowTitle("ppt time · 讲台状态")
            self.dialog.setMinimumWidth(360)
            layout = QVBoxLayout(self.dialog)
            self.details = QLabel()
            self.details.setWordWrap(True)
            layout.addWidget(self.details)
            toggle = QCheckBox("本次运行允许遥控端操作 Windows 桌面")
            toggle.setEnabled(sys.platform == "win32")
            toggle.toggled.connect(lambda checked: setattr(self, "enabled", checked))
            layout.addWidget(toggle)
            hint = QLabel("控制链接持有者可显示桌面、打开资源管理器、切换或关闭前台窗口、打开投影设置和锁屏。")
            hint.setWordWrap(True)
            layout.addWidget(hint)
            self.refresh = QTimer(self.dialog)
            self.refresh.setInterval(2200)
            self.refresh.timeout.connect(self._details)
            self.dialog.finished.connect(lambda _: self.refresh.stop())
        self._details()
        self.refresh.start()
        self.dialog.show()
        self.dialog.raise_()
        self.dialog.activateWindow()

    def _details(self):
        values = self.metrics()
        fields = [("系统", "os"), ("CPU", "cpu_percent"), ("内存", "memory_percent"),
                  ("图形设备", "graphics"), ("供电", "power"), ("电量", "battery_percent")]
        self.details.setText("\n".join(f"{title}　{values[key]}{'%' if key.endswith('_percent') else ''}"
                                      for title, key in fields if key in values and values[key] != ""))

    def close(self):
        self.enabled = False
        if self.dialog:
            self.dialog.close()
            self.dialog.deleteLater()
            self.dialog = None
