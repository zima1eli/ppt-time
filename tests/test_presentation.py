import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication
import main
from slideshow import SlideSnapshot


class PresentationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def setUp(self):
        for mock in (patch.object(main.Config, "load"), patch.object(main.Config, "save"),
                     patch.object(main, "HAS_KEYBOARD", False), patch.object(main.SlideMonitor, "start")):
            mock.start()
            self.addCleanup(mock.stop)
        self.app = main.App()
        self.app.config.global_sound = False
        self.addCleanup(self.app._exit)

    def sample(self, active, current=None, total=None):
        snapshot = SlideSnapshot(active, current, total)
        self.app._check_ppt_status(snapshot)
        self.app._check_ppt_status(snapshot)

    def test_show_following_manual_pause_and_unknown_page(self):
        self.sample(True, 3, 18)
        self.assertFalse(self.app.controller.paused)
        overlay, = self.app.display.overlays.values()
        self.assertEqual(overlay._page, "3/18")
        self.sample(True, 15, 18)
        self.assertEqual(overlay._page, "15/18")
        self.app.float_bar.btn_toggle.click()
        self.sample(True)
        self.assertTrue(self.app.controller.paused)
        self.assertEqual(overlay._page, "")
        self.sample(False)
        self.assertFalse(self.app.display.overlays)

    def test_same_office_window_can_supply_total_for_a_direct_partial_sample(self):
        player = self.app.player
        direct = SlideSnapshot(True, current=3, document="A.pptx", hwnd=101)
        player.owned_snapshot = direct
        self.app._external_snapshot = SlideSnapshot(True, 3, 20, document="A.pptx", hwnd=101)
        self.assertEqual(self.app._effective_snapshot().page_text, "3/20")
        self.app._external_snapshot = SlideSnapshot(True, 2, 20, document="A.pptx", hwnd=101)
        self.assertEqual(self.app._effective_snapshot(), direct)
        self.app._external_snapshot = SlideSnapshot(True, 3, 20, document="B.pptx", hwnd=102)
        self.assertEqual(self.app._effective_snapshot(), direct)

    def test_tray_status_tracks_real_pages_manual_mode_and_saved_settings(self):
        status = self.app.tray_status
        self.assertIn("等待", status.text())
        self.app._preview()
        self.assertIn("等待", status.text())
        self.sample(True)
        self.assertEqual(status.text(), "放映中 · 未读取到页码")
        self.sample(True, 2, 4)
        self.assertEqual(status.text(), "放映中 · 页码 2/4")
        self.assertIn(status.text(), self.app.tray.toolTip())

        dialog = main.SettingsWindow(self.app.config)
        self.addCleanup(dialog.deleteLater)
        dialog.switch_presentation_pages.setChecked(False)
        self.app._save_settings(dialog)
        self.assertIn("页码显示已关闭", status.text())
        dialog.switch_presentation_pages.setChecked(True)
        self.app._save_settings(dialog)
        self.assertIn("2/4", status.text())
        self.sample(False)
        manual = self.app.float_bar.manual_presentation_action
        manual.setChecked(True)
        self.assertEqual(status.text(), "手动计时 · 未读取到页码")
        manual.setChecked(False)
        self.assertIn("等待", status.text())

        dialog.cmb_presentation_style.setCurrentIndex(dialog.cmb_presentation_style.findData("classic"))
        dialog.switch_ppt_auto_start.setChecked(False)
        dialog.switch_suspend_hotkeys.setChecked(False)
        self.app._save_settings(dialog)
        self.assertTrue(dialog.presentation_pages_row.isEnabled())
        self.assertTrue(self.app._ppt_watch_enabled())
        dialog.switch_presentation_pages.setChecked(False)
        self.app._save_settings(dialog)
        self.assertEqual(status.text(), "放映检测已关闭")
        self.app.show_timing.set_options({"pages_only": True})
        self.assertTrue(self.app._ppt_watch_enabled())
        self.assertIn("等待", status.text())

    def test_preview_cannot_leak_example_numbers_into_show(self):
        self.app._preview()
        preview, = self.app.display.overlays.values()
        self.assertEqual(preview._page, "16/53")
        self.assertTrue(self.app.controller.paused)
        self.sample(True)
        overlay, = self.app.display.overlays.values()
        self.assertEqual(overlay._page, "")
        self.assertEqual(overlay._time, "05:00")
        self.sample(False)
        self.assertTrue(self.app.controller.paused)

    def test_silent_warning_after_delay_and_fractional_pause(self):
        ctrl = self.app.controller
        with patch.object(main.time, "monotonic", return_value=100):
            ctrl.toggle_pause()
        with patch.object(main.time, "monotonic", return_value=391.25):
            ctrl._tick()
            self.assertEqual(self.app.display.time_text, "00:09")
            self.assertEqual(self.app.display.alert_color, "#ff4444")
        with patch.object(main.time, "monotonic", return_value=391.75):
            ctrl.toggle_pause()
        self.assertEqual(ctrl._remaining_float, 8.25)

    def test_manual_timer_keeps_running_after_a_show_and_removes_its_pages(self):
        action = self.app.float_bar.manual_presentation_action
        action.setChecked(True)
        self.assertFalse(self.app.controller.paused)
        self.sample(True, 2, 3)
        overlay, = self.app.display.overlays.values()
        self.assertEqual(overlay._page, "2/3")
        self.sample(False)
        self.assertEqual(overlay._page, "")
        self.assertFalse(self.app.controller.paused)
        action.setChecked(False)
        self.assertTrue(self.app.controller.paused)
        self.assertFalse(self.app.display.overlays)

    def test_style_save_preserves_running_time(self):
        self.sample(True, 4, 10)
        old_deadline = self.app.controller._target_time
        dialog = main.SettingsWindow(self.app.config)
        self.addCleanup(dialog.deleteLater)
        dialog.cmb_presentation_style.setCurrentIndex(dialog.cmb_presentation_style.findData("pill"))
        self.app._save_settings(dialog)
        self.assertFalse(self.app.controller.paused)
        self.assertEqual(old_deadline, self.app.controller._target_time)
        overlay, = self.app.display.overlays.values()
        self.assertEqual(overlay._style, "pill")
