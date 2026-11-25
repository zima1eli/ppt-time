import os
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from show_timing import ShowTiming, TimedStages, TimingDialog, timing_options


def snapshot(page, document="C:/会议/晨会.pptx"):
    return SimpleNamespace(active=True, current=page, document=document,
                           page_text=f"{page}/12" if page else "")


class ShowTimingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QApplication.instance() or QApplication([])

    def setUp(self):
        self.now = 100.
        timer = patch("show_timing.time.monotonic", side_effect=lambda: self.now)
        timer.start()
        self.addCleanup(timer.stop)
        self.sound = Mock()
        self.config = SimpleNamespace(
            stages=[SimpleNamespace(label="报告", seconds=30, count_up=False)],
            global_sound=True, countdown_10s_sound=False, countdown_beep_seconds=10,
            alerts=[], alert_sound_path="", end_sound_path="", auto_advance=False,
            ppt_auto_start=True, presentation_screen="primary")
        self.clock = TimedStages(self.config, self.sound)
        self.clock.stop()
        self.addCleanup(self.clock.stop)
        self.display = SimpleNamespace(config=self.config, manual=False, set_pages_only=Mock(),
                                       _target_screens=Mock(return_value=[self.qt.primaryScreen()]))
        self.save = Mock()
        self.timing = ShowTiming(self.clock, self.display, {}, self.save)
        self.addCleanup(self.timing.close)
        self.ticks = []
        self.clock.tick.connect(lambda label, seconds: self.ticks.append(seconds))
        self.expired = Mock()
        self.clock.expired.connect(self.expired)

    def test_overtime_pause_and_resume_keep_fractions_and_report_only_once(self):
        self.timing.set_options({"overtime": True})
        self.clock.set_duration(3, start=True)
        self.now += 8.75
        self.clock._tick()
        self.assertEqual(self.ticks[-2:], [0, -5])
        self.expired.assert_called_once()
        self.sound.assert_called_once()
        self.clock.toggle_pause()
        self.assertTrue(self.clock.can_resume)
        self.now += 100
        self.clock.toggle_pause()
        self.now += .5
        self.clock._tick()
        self.assertEqual(self.ticks[-1], -6)
        self.clock.set_duration(20, start=False)
        self.assertEqual(self.ticks[-1], 20)
        self.assertTrue(self.clock.paused)
        self.assertFalse(self.clock._overtime)

    def test_auto_transition_precedes_end_layer_and_overtime(self):
        self.config.auto_advance = True
        self.config.stages.append(SimpleNamespace(label="答辩", seconds=15, count_up=True))
        self.timing.set_options({"overtime": True, "end_mode": "black"})
        self.timing.update_snapshot(snapshot(1))
        self.clock.set_duration(2, True)
        self.now += 2
        self.clock._tick()
        self.assertEqual(self.clock._stage_idx, 1)
        self.assertEqual(self.ticks[-1], 0)
        self.assertFalse(self.timing._windows)
        self.expired.assert_not_called()

    def test_full_path_filename_and_generic_rules_and_reentry(self):
        self.timing.set_options({"rules": [
            {"document": "", "page": "*", "seconds": 120},
            {"document": "", "page": 2, "seconds": 4},
            {"document": "晨会.pptx", "page": 2, "seconds": 8},
            {"document": "c:/会议/晨会.pptx", "page": 2, "seconds": 14}]})
        self.timing.update_snapshot(snapshot(1))
        self.assertEqual(self.ticks[-1], 120)
        self.clock.toggle_pause()
        self.timing.update_snapshot(snapshot(2))
        self.assertEqual(self.ticks[-1], 14)
        deadline = self.clock._target_time
        self.now += 7
        self.timing.update_snapshot(snapshot(None, ""))
        self.timing.update_snapshot(snapshot(2))
        self.assertEqual(self.clock._target_time, deadline)
        self.timing.update_snapshot(snapshot(3))
        self.assertEqual(self.clock._target_time, deadline)
        self.timing.update_snapshot(snapshot(2))
        self.assertEqual(self.clock._target_time, self.now + 14)
        self.timing.update_snapshot(snapshot(2, "d:/其他/晨会.pptx"))
        self.assertEqual(self.ticks[-1], 8)
        self.timing.update_snapshot(snapshot(2, "d:/其他/报告.pptx"))
        self.assertEqual(self.ticks[-1], 4)
        self.timing.update_snapshot(None)
        self.timing.update_snapshot(snapshot(1))
        self.assertEqual(self.ticks[-1], 120)

    def test_pages_only_pauses_both_clock_and_sound(self):
        self.timing.update_snapshot(snapshot(2))
        self.clock.set_duration(3, True)
        self.now += 1.25
        self.timing.set_options({"pages_only": True, "end_mode": "black"})
        self.assertTrue(self.clock.paused)
        self.assertEqual(self.clock._remaining_float, 1.75)
        self.assertFalse(self.clock.can_resume)
        self.clock.toggle_pause()
        self.now += 60
        self.clock._tick()
        self.sound.assert_not_called()
        self.display.set_pages_only.assert_called_with(True)
        self.timing.set_options({})
        self.assertFalse(self.clock.paused)
        self.assertAlmostEqual(self.clock._target_time - self.now, 1.75)

    def test_operator_pause_and_current_page_budget_survive_mode_or_clock_reset(self):
        self.timing.set_options({"rules": [{"page": 2, "seconds": 9}]})
        self.timing.update_snapshot(snapshot(2))
        self.clock.toggle_pause()
        self.now += 2
        self.clock.toggle_pause()
        self.timing.set_options(dict(self.timing.options, pages_only=True))
        self.timing.set_options(dict(self.timing.options, pages_only=False))
        self.assertTrue(self.clock.paused)
        self.assertEqual(self.clock._remaining_float, 7)
        self.clock.restart_stage()
        self.assertEqual(self.ticks[-1], 9)
        self.clock.stop()
        self.assertEqual(self.ticks[-1], 30)

    def test_black_layer_and_locked_notice_clear_on_show_exit(self):
        self.timing.update_snapshot(snapshot(3))
        for mode, allow, should_dismiss in (("black", False, True), ("notice", True, True),
                                            ("notice", False, False)):
            with self.subTest(mode=mode, allow=allow):
                self.timing.set_options({"end_mode": mode, "allow_close": allow,
                                         "end_screen": "all", "message": "演讲结束"})
                self.clock.set_duration(1, True)
                self.now += 1
                self.clock._tick()
                window, = self.timing._windows.values()
                self.assertEqual(window.label.text(), "演讲结束")
                self.display._target_screens.assert_called_with("all")
                if mode == "black":
                    self.assertEqual(window.geometry(), self.qt.primaryScreen().geometry())
                QTest.keyClick(window, Qt.Key.Key_Escape)
                self.assertEqual(not self.timing._windows, should_dismiss)
                self.timing.update_snapshot(None)
                self.assertFalse(self.timing._windows)
                self.timing.update_snapshot(snapshot(3))

    def test_options_round_trip_and_dialog_current_page(self):
        options = timing_options({"rules": [{"document": "", "page": "*", "seconds": "01:10"}],
                                  "end_mode": "notice", "message": "<测试>"})
        dialog = TimingDialog(options, "D:/deck.pptx", 7)
        self.addCleanup(dialog.deleteLater)
        dialog._add_current()
        dialog.accept()
        self.assertEqual(dialog.options["rules"][0]["seconds"], 70)
        self.assertEqual(dialog.options["rules"][1], {"document": "D:/deck.pptx", "page": 7, "seconds": 60})
        self.assertEqual(timing_options(dialog.options), dialog.options)
        self.assertEqual(timing_options({"rules": [{"seconds": "NaN"}]})["rules"], [])
