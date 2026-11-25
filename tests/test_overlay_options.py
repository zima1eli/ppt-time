import json
import os
from pathlib import Path
import sys
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtCore import QRect, Qt
from PySide6.QtGui import QColor, QImage, QPainter
from PySide6.QtWidgets import QApplication, QSystemTrayIcon, QWidget
from types import SimpleNamespace

from overlay_options import SIZE_MODES, OverlayDialog, OverlayOptions, content_score, sample_underlay
from presentation_display import PresentationDisplay
from presentation_overlay import PresentationOverlay
from slideshow import SlideSnapshot


class OverlayOptionsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def test_position_font_radius_and_inherited_colors(self):
        bar = PresentationOverlay()
        try:
            bg = QColor("#112233")
            bg.setAlpha(180)
            options = OverlayOptions(anchor="bottom-right", x_offset=-17, edge_gap=9, radius=4,
                                     time_size=46, background="#660055")
            bar.configure("notch", 100, bg, QColor("white"), .8, options)
            bar._set_screen_rect(QRect(-1920, 0, 1920, 1080))
            self.assertEqual(bar.x() + bar.width(), -17)
            self.assertEqual(bar.y() + bar.height(), 1071)
            self.assertEqual(bar._background.name(), "#660055")
            self.assertEqual(bar._background.alpha(), 180)
            self.assertGreater(bar.height(), 46)
            self.assertTrue(bar._outline().contains(bar._outline().boundingRect().center()))
            bar.configure("notch", 100, bg, QColor("white"), .8)
            self.assertEqual(bar._background, bg)
            self.assertEqual(bar.y(), 0)
        finally:
            bar.close()

    def test_settings_roundtrip_and_recovery(self):
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as temp:
            path = Path(temp.name)
        try:
            options = OverlayOptions(smart_fade=True, anchor="top-right", background="#303e3a")
            options.save(path)
            self.assertEqual(options, OverlayOptions.load(path))
            path.write_text(json.dumps({"radius": 500, "time_size": -5, "font": None,
                                        "foreground": "bad shade", "anchor": []}))
            recovered = OverlayOptions.load(path)
            self.assertEqual((recovered.radius, recovered.time_size, recovered.anchor), (96, 10, "top-center"))
            path.write_text("not json")
            self.assertEqual(OverlayOptions.load(path), OverlayOptions())
        finally:
            path.unlink(missing_ok=True)

    def test_dialog_defaults_and_roundtrip(self):
        dlg = OverlayDialog(OverlayOptions())
        self.assertEqual(dlg.options(), OverlayOptions())
        dlg.anchor.setCurrentIndex(dlg.anchor.findData("bottom-left"))
        dlg.time_size.setValue(27)
        self.assertEqual((dlg.options().anchor, dlg.options().time_size), ("bottom-left", 27))
        dlg.close()

    def test_projection_size_migrates_only_valid_old_custom_fonts(self):
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as temp:
            path = Path(temp.name)
        self.addCleanup(path.unlink, missing_ok=True)
        for raw, expected in (
                ({}, "auto"), ({"time_size": 20, "page_size": 16}, "auto"),
                ({"time_size": 24}, "100"), ({"page_size": 25}, "100"),
                ({"time_size": -3, "page_size": "24"}, "auto"),
                ({"time_size": 32, "size_mode": "auto"}, "auto"),
                ({"time_size": 32, "size_mode": "175"}, "175"),
                ({"size_mode": "future-option"}, "auto")):
            with self.subTest(raw=raw):
                path.write_text(json.dumps(raw), encoding="utf-8")
                self.assertEqual(OverlayOptions.load(path).size_mode, expected)
        self.assertEqual(OverlayOptions.from_mapping({"time_size": 24}, strict=True).size_mode, "auto")
        for bad in (125, True, None, "120", "125%", {}):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                OverlayOptions.from_mapping({"size_mode": bad}, strict=True)

    def test_projection_size_native_web_choices_and_roundtrip(self):
        from remote_appearance import APPEARANCE_HTML
        choices = APPEARANCE_HTML.split('<select name="size_mode">', 1)[1].split("</select>", 1)[0]
        dlg = OverlayDialog(OverlayOptions(size_mode="160", page_size=23))
        self.addCleanup(dlg.close)
        self.assertEqual(dlg.options().size_mode, "160")
        for index, (mode, label) in enumerate(SIZE_MODES.items()):
            self.assertEqual(dlg.size_mode.itemData(index), mode)
            self.assertEqual(dlg.size_mode.itemText(index), label)
            self.assertIn(f'<option value="{mode}">{label}</option>', choices)
            dlg.size_mode.setCurrentIndex(index)
            options = dlg.options()
            self.assertEqual((options.size_mode, options.page_size), (mode, 23))
            self.assertEqual(OverlayOptions.from_mapping({"size_mode": mode}, strict=True).size_mode, mode)
    def test_lcd_bold_and_secondary_offsets(self):
        options = OverlayOptions(lcd=True, bold=True, secondary_offsets=True, x_offset=11, edge_gap=6,
                                 secondary_x=-200, secondary_gap=44, flash_last_minute=True, flash_expiry=True)
        dialog = OverlayDialog(options)
        self.assertEqual(options, dialog.options())
        dialog.close()
        self.assertEqual((11,6), options.offsets(False))
        self.assertEqual((-200,44), options.offsets(True))
        bar = PresentationOverlay()
        bar.set_content("+01:40:17", "9/12")
        bar.configure("notch", 100, QColor("black"), QColor("white"), 1, options)
        original = bar.grab().toImage()
        self.assertGreater(bar._time_width, 80)
        bar.configure("notch", 100, QColor("black"), QColor("white"), 1, OverlayOptions())
        self.assertNotEqual(original, bar.grab().toImage())
        bar.close()

    def test_remote_atomic_changes_and_warning_flashes(self):
        cfg = SimpleNamespace(presentation_style="pill", presentation_screen="primary",
                              presentation_pages=True, presentation_only=False, color="#ffffff",
                              font_transparency=0, global_transparency=0, ui_scale=100, bg_qcolor=lambda: QColor("black"))
        controls = QWidget()
        d = PresentationDisplay(cfg, controls)
        self.addCleanup(d.close)
        self.addCleanup(controls.close)
        d.set_snapshot(SlideSnapshot(True, 1, 3))
        self.assertTrue(d.remote_appearance({"lcd": True, "flash_last_minute": True, "flash_expiry": True}))
        self.assertFalse(d.remote_appearance({"radius": 300, "lcd": False}))
        self.assertFalse(d.remote_appearance({"private": 2}))
        self.assertTrue(d.appearance.lcd)
        self.assertTrue(d.remote_appearance({"size_mode": "250"}))
        self.assertFalse(d.remote_appearance({"size_mode": 250, "lcd": False}))
        self.assertTrue(d.appearance.lcd)
        self.assertEqual(d.appearance.size_mode, "250")
        d.set_countdown_state(58, True, True)
        self.assertTrue(d._flash_timer.isActive())
        d._step_flash()
        self.assertTrue(next(iter(d.overlays.values()))._flash_dim)
        d.set_timer_running(False)
        self.assertFalse(d._flash_timer.isActive())
        d.set_countdown_state(0, True, False)
        d.flash_expired()
        self.assertTrue(d._flash_timer.isActive())
        with patch("presentation_display.time.monotonic", return_value=d._flash_until + 1):
            d._step_flash()
        self.assertFalse(d._flash_timer.isActive())
        d.set_countdown_state(50, False, True)
        self.assertFalse(d._flash_timer.isActive())

    @unittest.skipUnless(sys.platform == "win32", "Windows recording exclusion")
    def test_native_recording_affinity_can_be_restored(self):
        from overlay_options import capture_supported
        import ctypes
        from ctypes import wintypes
        if self.qt.platformName() != "windows" or not capture_supported():
            self.skipTest("Modern Windows desktop compositor required")
        cfg = SimpleNamespace(presentation_style="pill", presentation_screen="primary",
                              presentation_pages=True, presentation_only=False, color="#ffffff",
                              font_transparency=0, global_transparency=5, ui_scale=100, bg_qcolor=lambda: QColor("#da123456"))
        controls = QWidget()
        d = PresentationDisplay(cfg, controls)
        self.addCleanup(d.close)
        self.addCleanup(controls.close)
        d.start_preview("pill","primary",True)
        affinity = ctypes.windll.user32.GetWindowDisplayAffinity
        affinity.argtypes = wintypes.HWND, ctypes.POINTER(wintypes.DWORD)
        for excluded in (True,False):
            self.assertTrue(d.remote_appearance({"exclude_capture":excluded}))
            self.qt.processEvents()
            w, = d.overlays.values()
            self.assertEqual(w.capture_excluded, excluded)
            result = wintypes.DWORD()
            self.assertTrue(affinity(int(w.winId()), ctypes.byref(result)))
            self.assertEqual(result.value, 0x11 if excluded else 0)

    def test_content_score_uniform_background_and_text(self):
        image = QImage(220, 44, QImage.Format.Format_RGB32)
        image.fill(QColor("#f2f2f2"))
        self.assertEqual(content_score(image), 0)
        painter = QPainter(image)
        painter.setPen(QColor("black"))
        painter.drawText(image.rect(), Qt.AlignmentFlag.AlignCenter, "Slide 04   Workshop agenda")
        painter.end()
        self.assertGreater(content_score(image), .024)

    @unittest.skipUnless(sys.platform == "win32", "Windows window sampling")
    def test_native_sampler_ignores_its_own_opaque_digits(self):
        if self.qt.platformName() != "windows":
            self.skipTest("Run with QT_QPA_PLATFORM=windows for the native compositor")

        class Slide(QWidget):
            def paintEvent(self, event):
                painter = QPainter(self)
                painter.fillRect(self.rect(), QColor("white"))
                for x in range(70, 340, 8):
                    painter.fillRect(QRect(x, 54, 4, 60), QColor("black"))

        source = Slide(None, Qt.WindowType.FramelessWindowHint)
        bar = PresentationOverlay()
        self.addCleanup(source.close)
        self.addCleanup(bar.close)
        source.setGeometry(100, 100, 500, 210)
        source.show()
        bar.configure("pill", 100, QColor("black"), QColor("white"), 1.)
        bar.set_content("22:02", "16/53")
        bar.show()
        for y, busy in ((164, True), (234, False)):
            bar.move(215, y)
            self.qt.processEvents()
            hwnd, cover = int(source.winId()), int(bar.winId())
            output = []
            worker = threading.Thread(target=lambda: output.append(sample_underlay(hwnd, cover)), daemon=True)
            worker.start()
            deadline = time.monotonic() + 4
            while worker.is_alive() and time.monotonic() < deadline:
                self.qt.processEvents()
                time.sleep(.01)
            self.assertTrue(output)
            self.assertIsNotNone(output[0])
            self.assertEqual(output[0] >= .024, busy)

    def test_pages_only_waits_for_real_pages_and_overrides_classic(self):
        class Config:
            presentation_style = "classic"
            presentation_screen = "primary"
            presentation_pages = False
            presentation_only = False
            color = "#ffffff"
            font_transparency = 0
            global_transparency = 0
            ui_scale = 100
            def bg_qcolor(self):
                return QColor("black")
        controls = QWidget()
        display = PresentationDisplay(Config(), controls)
        with patch.object(QSystemTrayIcon, "isSystemTrayAvailable", return_value=True):
            display.set_pages_only(True)
            display.set_snapshot(SlideSnapshot(True))
            overlay, = display.overlays.values()
            self.assertFalse(overlay.isVisible())
            display.set_snapshot(SlideSnapshot(True, 2, 5))
            self.assertTrue(overlay.isVisible())
            self.assertEqual((overlay._time, overlay._page), ("", "2/5"))
            display.set_pages_only(False)
            self.assertFalse(display.overlays)
            self.assertTrue(controls.isVisible())
        display.close()
        controls.close()
