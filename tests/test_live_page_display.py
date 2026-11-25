import os
import unittest
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtCore import QRect
from PySide6.QtGui import QColor, QFontDatabase
from PySide6.QtWidgets import QApplication, QWidget

from overlay_options import OverlayOptions
from presentation_display import PresentationDisplay
from presentation_overlay import PresentationOverlay
from slideshow import SlideSnapshot


class ClassicBar(QWidget):
    def set_page_text(self, page):
        self.page = page


class LivePageDisplayTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def setUp(self):
        self.cfg = SimpleNamespace(
            presentation_style="classic", presentation_screen="primary", presentation_pages=True,
            presentation_only=False, color="#ffffff", font_transparency=0, global_transparency=0,
            ui_scale=100, bg_qcolor=lambda: QColor("black"),
        )
        self.controls = ClassicBar()
        self.display = PresentationDisplay(self.cfg, self.controls)
        self.addCleanup(self.controls.close)
        self.addCleanup(self.display.close)

    def test_classic_shares_live_pages_and_clears_unknown_disabled_and_exit(self):
        d = self.display
        d.set_manual(True)
        d.set_snapshot(SlideSnapshot(True, 1, 4, document="one.pptx"))
        self.assertEqual(self.controls.page, "1/4")
        d.set_snapshot(SlideSnapshot(True, 3, 4, document="one.pptx"))
        d.update_content("09:53", "#dd2222")
        self.assertEqual(self.controls.page, "3/4")
        self.cfg.presentation_pages = False
        d.refresh()
        self.assertEqual(self.controls.page, "")
        self.cfg.presentation_pages = True
        d.refresh()
        self.assertEqual(self.controls.page, "3/4")
        d.set_snapshot(SlideSnapshot(True, document="two.pptx"))
        self.assertEqual(self.controls.page, "")
        d.set_snapshot(SlideSnapshot(True, 1, 8, document="two.pptx"))
        self.assertEqual(self.controls.page, "1/8")
        d.set_snapshot(SlideSnapshot(True, 1, 8, document="two.pptx", ended=True))
        self.assertEqual(self.controls.page, "")
        d.set_snapshot(SlideSnapshot(False, 1, 8))
        self.assertIsNone(d.snapshot)
        self.assertEqual(self.controls.page, "")
        self.assertFalse(d.overlays)

    def test_preview_started_during_live_show_ends_with_that_show(self):
        d = self.display
        self.cfg.presentation_style = "shoulder"
        d.update_content("04:27", "")
        d.set_snapshot(SlideSnapshot(True, 7, 12))
        d.start_preview("pill", "primary", True)
        overlay, = d.overlays.values()
        self.assertEqual((overlay._time, overlay._page), ("04:27", "7/12"))
        d.set_snapshot(None)
        self.assertIsNone(d.preview)
        self.assertFalse(d.overlays)
        self.assertFalse(d._preview_timer.isActive())

    def test_projection_sizes_follow_desktop_and_each_screen_independently(self):
        bar = PresentationOverlay()
        self.addCleanup(bar.close)
        bar.set_content("05:00", "2/18")
        bar.configure("shoulder", 142, QColor("black"), QColor("white"), 1, desktop_time_size=27)
        for width, height, factor in ((1920, 1080, 1), (1280, 720, .875), (3840, 2160, 2)):
            bar._set_screen_rect(QRect(0, 0, width, height))
            self.assertAlmostEqual(bar._time_font.pixelSize() * bar._paint_scale, 27 * 1.42 * factor)
        bar.configure("shoulder", 142, QColor("black"), QColor("white"), 1,
                      OverlayOptions(size_mode="100"), desktop_time_size=27)
        fixed = bar.size()
        self.assertAlmostEqual(bar._paint_scale, 1.42)
        bar._set_screen_rect(QRect(0, 0, 1920, 1080))
        self.assertEqual(bar.size(), fixed)

    def test_large_digits_fit_projection_and_render_live_page_changes(self):
        bar = PresentationOverlay()
        self.addCleanup(bar.close)
        bar._set_screen_rect(QRect(0, 0, 640, 480))
        for style in ("notch", "pill", "shoulder"):
            with self.subTest(style=style):
                bar.configure(style, 300, QColor("black"), QColor("white"), 1,
                              OverlayOptions(time_size=70, page_size=64, lcd=not QFontDatabase.families(),
                                             size_mode="100"))
                bar.set_content("01:59:39", "103/240")
                self.assertEqual(bar.width(), 640)

                def painted_page():
                    pixmap = bar.grab()
                    scale = pixmap.devicePixelRatio() * bar._paint_scale
                    x = (bar._logical_width - bar._padding - bar._page_width) * scale
                    return pixmap.toImage().copy(round(x), 0, round(bar._page_width * scale), pixmap.height())

                first = painted_page()
                self.assertTrue(any(first.pixelColor(x, y).lightness() > 150
                                    for x in range(first.width()) for y in range(first.height())))
                bar.set_content("01:59:39", "108/240")
                self.assertNotEqual(first, painted_page())
                bar._set_screen_rect(QRect(0, 0, 3840, 2160))
                self.assertEqual(bar._paint_scale, 3)
                bar._set_screen_rect(QRect(0, 0, 640, 480))
