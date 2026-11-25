"""Settings and Windows content sampling for the audience timer strip."""
from dataclasses import asdict, dataclass, fields
from pathlib import Path
import json
import sys

from PySide6.QtCore import QIODevice, QSaveFile, Qt
from PySide6.QtGui import QColor, QImage
from PySide6.QtWidgets import (QCheckBox, QColorDialog, QComboBox, QDialog, QDialogButtonBox,
                               QFontComboBox, QFormLayout, QHBoxLayout, QLabel, QPushButton,
                               QSpinBox, QTabWidget, QVBoxLayout, QWidget)


ANCHORS = {"top-center": "顶部居中", "top-left": "左上", "top-right": "右上",
           "bottom-center": "底部居中", "bottom-left": "左下", "bottom-right": "右下"}
SIZE_MODES = {"auto": "自动（匹配初始浮窗）", "75": "75%", "100": "100%（原大小）",
              "125": "125%", "150": "150%", "160": "160%", "175": "175%",
              "200": "200%", "250": "250%", "300": "300%"}
FONT_SIZE_RANGE = (10, 80)


@dataclass(frozen=True)
class OverlayOptions:
    anchor: str = "top-center"
    x_offset: int = 0
    edge_gap: int = -1
    radius: int = -1
    time_size: int = 20
    page_size: int = 16
    font: str = ""
    foreground: str = ""
    background: str = ""
    smart_fade: bool = False
    fade_percent: int = 18
    lcd: bool = False
    bold: bool = False
    secondary_offsets: bool = False
    secondary_x: int = 0
    secondary_gap: int = -1
    exclude_capture: bool = False
    flash_last_minute: bool = False
    flash_expiry: bool = False
    size_mode: str = "auto"

    def offsets(self, secondary):
        if secondary and self.secondary_offsets:
            return self.secondary_x, self.secondary_gap
        return self.x_offset, self.edge_gap

    @classmethod
    def load(cls, path):
        try:
            raw = json.loads(Path(path).read_text(encoding="utf-8"))
        except (OSError, ValueError, UnicodeError):
            return cls()
        if not isinstance(raw, dict):
            return cls()
        if "size_mode" not in raw:
            # Earlier versions rendered custom projection font sizes at their native scale.
            for key in ("time_size", "page_size"):
                size = raw.get(key)
                if (type(size) is int and FONT_SIZE_RANGE[0] <= size <= FONT_SIZE_RANGE[1]
                        and size != getattr(cls, key)):
                    raw["size_mode"] = "100"
                    break
        return cls.from_mapping(raw)

    @classmethod
    def from_mapping(cls, raw, strict=False):
        known = {f.name for f in fields(cls)}
        if strict and (not isinstance(raw, dict) or raw.keys() - known):
            raise ValueError("未知的放映条外观设置。")
        values = {}
        ranges = {"x_offset": (-4096, 4096), "edge_gap": (-1, 2048), "radius": (-1, 96),
                  "secondary_x": (-4096, 4096), "secondary_gap": (-1, 2048),
                  "time_size": FONT_SIZE_RANGE, "page_size": FONT_SIZE_RANGE, "fade_percent": (0, 95)}
        flags = {"smart_fade", "lcd", "bold", "secondary_offsets", "exclude_capture", "flash_last_minute", "flash_expiry"}
        for key, item in raw.items():
            if key in ranges and type(item) is int:
                low, high = ranges[key]
                if strict and not low <= item <= high:
                    raise ValueError("放映条数值超出范围。")
                values[key] = max(low, min(high, item))
            elif key in flags and isinstance(item, bool):
                values[key] = item
            elif key == "anchor" and isinstance(item, str) and item in ANCHORS:
                values[key] = item
            elif key == "size_mode" and isinstance(item, str) and item in SIZE_MODES:
                values[key] = item
            elif key == "font" and isinstance(item, str) and (not strict or len(item) <= 128):
                values[key] = item[:128]
            elif key in ("foreground", "background") and isinstance(item, str) and (
                    not item or (len(item) <= 32 and QColor.isValidColorName(item))):
                values[key] = QColor(item).name() if item else ""
            elif strict:
                raise ValueError("放映条设置值无效。")
        return cls(**values)

    def save(self, path):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        file = QSaveFile(str(path))
        data = json.dumps(asdict(self), ensure_ascii=False, indent=2).encode("utf-8")
        if not file.open(QIODevice.OpenModeFlag.WriteOnly) or file.write(data) != len(data) or not file.commit():
            raise OSError(file.errorString())


class OverlayDialog(QDialog):
    def __init__(self, options, parent=None):
        super().__init__(parent)
        self.setWindowTitle("放映条外观 · ppt time")
        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs)
        def section(name):
            page = QWidget()
            tabs.addTab(page, name)
            return QFormLayout(page)
        form = section("大小与位置")
        self.size_mode = QComboBox()
        for key, label in SIZE_MODES.items():
            self.size_mode.addItem(label, key)
        self.size_mode.setCurrentIndex(max(0, self.size_mode.findData(options.size_mode)))
        form.addRow("放映条大小", self.size_mode)
        self.anchor = QComboBox()
        for key, label in ANCHORS.items():
            self.anchor.addItem(label, key)
        self.anchor.setCurrentIndex(self.anchor.findData(options.anchor))
        form.addRow("位置", self.anchor)

        def spin(label, low, high, value, suffix=" px"):
            widget = QSpinBox()
            widget.setRange(low, high)
            widget.setValue(value)
            widget.setSuffix(suffix)
            form.addRow(label, widget)
            return widget

        self.x_offset = spin("水平偏移（向右为正）", -4096, 4096, options.x_offset)
        self.edge_gap = spin("距上下边缘", -1, 2048, options.edge_gap)
        self.edge_gap.setSpecialValueText("随样式")
        self.secondary_offsets = QCheckBox("扩展屏使用独立位置偏移")
        self.secondary_offsets.setChecked(options.secondary_offsets)
        form.addRow(self.secondary_offsets)
        self.secondary_x = spin("扩展屏水平偏移", -4096, 4096, options.secondary_x)
        self.secondary_gap = spin("扩展屏距上下边缘", -1, 2048, options.secondary_gap)
        self.secondary_gap.setSpecialValueText("随样式")
        for field in (self.secondary_x, self.secondary_gap):
            field.setEnabled(options.secondary_offsets)
            self.secondary_offsets.toggled.connect(field.setEnabled)
        self.radius = spin("圆角半径", -1, 96, options.radius)
        self.radius.setSpecialValueText("随样式")
        form = section("字体与颜色")
        self.time_size = spin("时间字号", *FONT_SIZE_RANGE, options.time_size)
        self.page_size = spin("页码字号", *FONT_SIZE_RANGE, options.page_size)
        self.font = QFontComboBox()
        self.font.insertItem(0, "随系统")
        self.font.setCurrentIndex(max(0, self.font.findText(options.font)) if options.font else 0)
        form.addRow("字体", self.font)
        self.lcd = QCheckBox("七段液晶数字")
        self.lcd.setChecked(options.lcd)
        self.lcd.toggled.connect(self.font.setDisabled)
        self.font.setDisabled(options.lcd)
        self.bold = QCheckBox("加粗数字")
        self.bold.setChecked(options.bold)
        form.addRow("数字样式", self.lcd)
        form.addRow("", self.bold)
        self.colors = {"foreground": options.foreground, "background": options.background}
        for key, label in (("foreground", "文字颜色"), ("background", "背景颜色")):
            row = QWidget()
            buttons = QHBoxLayout(row)
            buttons.setContentsMargins(0, 0, 0, 0)
            choose, reset = QPushButton(self.colors[key] or "随基础设置"), QPushButton("重置")
            def pick(_checked, key=key, choose=choose):
                selected = QColorDialog.getColor(QColor(self.colors[key] or ("white" if key == "foreground" else "#141414")), self)
                if selected.isValid():
                    self.colors[key] = selected.name()
                    choose.setText(selected.name())
            def clear(_checked, key=key, choose=choose):
                self.colors[key] = ""
                choose.setText("随基础设置")
            choose.clicked.connect(pick)
            reset.clicked.connect(clear)
            buttons.addWidget(choose)
            buttons.addWidget(reset)
            form.addRow(label, row)
        form = section("提醒与录屏")
        self.flash_last_minute = QCheckBox("最后 60 秒让倒计时数字缓慢闪动")
        self.flash_last_minute.setChecked(options.flash_last_minute)
        self.flash_expiry = QCheckBox("到时后闪动 4 秒")
        self.flash_expiry.setChecked(options.flash_expiry)
        form.addRow("倒计时提醒", self.flash_last_minute)
        form.addRow("", self.flash_expiry)
        self.exclude_capture = QCheckBox("在兼容的屏幕录制中隐藏放映条")
        self.exclude_capture.setChecked(options.exclude_capture)
        self.exclude_capture.setEnabled(capture_supported())
        self.exclude_capture.setToolTip("适用于 Windows 10 2004 及以上版本支持窗口排除的录屏工具。排除录屏时，系统对放映条统一应用背景透明度。")
        form.addRow("录屏", self.exclude_capture)
        self.smart = QCheckBox("遮住幻灯片文字或图像细节时淡出")
        self.smart.setChecked(options.smart_fade)
        self.smart.setEnabled(sys.platform == "win32")
        form.addRow("智能闪避", self.smart)
        self.fade = spin("淡出后保留的不透明度", 0, 95, options.fade_percent, "%")
        self.fade.setEnabled(self.smart.isChecked() and self.smart.isEnabled())
        self.smart.toggled.connect(self.fade.setEnabled)
        note = QLabel("Windows 可读取放映窗口画面时启用智能闪避。")
        note.setWordWrap(True)
        form.addRow(note)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def options(self):
        return OverlayOptions(size_mode=self.size_mode.currentData(),
                              anchor=self.anchor.currentData(), x_offset=self.x_offset.value(),
                              edge_gap=self.edge_gap.value(), radius=self.radius.value(),
                              time_size=self.time_size.value(), page_size=self.page_size.value(),
                              font=self.font.currentText() if self.font.currentIndex() > 0 else "",
                              foreground=self.colors["foreground"], background=self.colors["background"],
                              smart_fade=self.smart.isChecked(), fade_percent=self.fade.value(),
                              lcd=self.lcd.isChecked(), bold=self.bold.isChecked(),
                              secondary_offsets=self.secondary_offsets.isChecked(),
                              secondary_x=self.secondary_x.value(), secondary_gap=self.secondary_gap.value(),
                              exclude_capture=self.exclude_capture.isChecked(),
                              flash_last_minute=self.flash_last_minute.isChecked(), flash_expiry=self.flash_expiry.isChecked())


def capture_supported():
    return sys.platform == "win32" and sys.getwindowsversion().build >= 19041


def set_capture_affinity(hwnd, excluded):
    if not capture_supported():
        return False
    import ctypes
    from ctypes import wintypes
    set_affinity = ctypes.windll.user32.SetWindowDisplayAffinity
    set_affinity.argtypes = (wintypes.HWND, wintypes.DWORD)
    set_affinity.restype = wintypes.BOOL
    return bool(set_affinity(hwnd, 0x11 if excluded else 0))


def content_score(image):
    """Fraction of neighbouring samples separated by a visible colour edge."""
    if image.isNull() or image.width() < 3 or image.height() < 3:
        return 0.
    small = image.scaled(min(128, image.width()), min(48, image.height()),
                         Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation)
    small = small.convertToFormat(QImage.Format.Format_RGB32)
    data = small.constBits()
    stride = small.bytesPerLine()
    edges = 0
    for y in range(1, small.height()):
        for x in range(1, small.width()):
            here = stride * y + x * 4
            for other in (here - 4, here - stride):
                if max(abs(data[here+c] - data[other+c]) for c in range(3)) >= 36:
                    edges += 1
    return edges / (2 * (small.width() - 1) * (small.height() - 1))


def sample_underlay(source_hwnd, overlay_hwnd):
    """Read just the overlapping rectangle from the slide window's own drawing.

    The desktop is never composited into the sample. An unsupported renderer,
    minimized source or non-overlapping mirror simply leaves the overlay visible.
    Called on a daemon worker: applications can take time to answer PrintWindow.
    """
    if sys.platform != "win32" or not source_hwnd or source_hwnd == overlay_hwnd:
        return None
    import ctypes
    from ctypes import wintypes
    import win32con
    import win32gui
    import win32ui
    if not win32gui.IsWindow(source_hwnd) or not win32gui.IsWindowVisible(source_hwnd) or win32gui.IsIconic(source_hwnd):
        return None
    left, top, right, bottom = win32gui.GetWindowRect(overlay_hwnd)
    sx, sy, sr, sb = win32gui.GetWindowRect(source_hwnd)
    if left < sx or top < sy or right > sr or bottom > sb:
        return None
    width, height = right-left, bottom-top
    # Some renderers reset the DC origin when printing, so use the source
    # window's natural coordinates, then read back only the exact footprint.
    paint_width, paint_height = right-sx, bottom-sy
    if width < 3 or height < 3 or width * height > 2_000_000 or paint_width * paint_height > 34_000_000:
        return None
    paint = ctypes.windll.user32.PrintWindow
    paint.argtypes = (wintypes.HWND, wintypes.HDC, wintypes.UINT)
    paint.restype = wintypes.BOOL
    window_dc = win32gui.GetWindowDC(source_hwnd)
    paint_dc = memory_dc = paint_bitmap = bitmap = paint_previous = previous = None
    try:
        paint_dc = win32gui.CreateCompatibleDC(window_dc)
        paint_bitmap = win32gui.CreateCompatibleBitmap(window_dc, paint_width, paint_height)
        paint_previous = win32gui.SelectObject(paint_dc, paint_bitmap)
        if not paint(source_hwnd, paint_dc, 2):  # PW_RENDERFULLCONTENT on Windows 8.1+
            return None
        memory_dc = win32gui.CreateCompatibleDC(window_dc)
        bitmap = win32gui.CreateCompatibleBitmap(window_dc, width, height)
        previous = win32gui.SelectObject(memory_dc, bitmap)
        win32gui.BitBlt(memory_dc, 0, 0, width, height, paint_dc, left-sx, top-sy, win32con.SRCCOPY)
        # GetBitmapBits returns device-format BGRA for a modern 32-bit display.
        handle = win32ui.CreateBitmapFromHandle(bitmap)
        info = handle.GetInfo()
        if info['bmBitsPixel'] != 32:
            return None
        data = handle.GetBitmapBits(True)
        image = QImage(data, width, height, info['bmWidthBytes'], QImage.Format.Format_RGB32)
        return content_score(image)
    finally:
        if previous:
            win32gui.SelectObject(memory_dc, previous)
        if bitmap:
            win32gui.DeleteObject(bitmap)
        if memory_dc:
            win32gui.DeleteDC(memory_dc)
        if paint_previous:
            win32gui.SelectObject(paint_dc, paint_previous)
        if paint_bitmap:
            win32gui.DeleteObject(paint_bitmap)
        if paint_dc:
            win32gui.DeleteDC(paint_dc)
        win32gui.ReleaseDC(source_hwnd, window_dc)
