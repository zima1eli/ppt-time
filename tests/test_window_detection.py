import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import main


class AudienceWindowTests(unittest.TestCase):
    def detect(self, windows, foreground):
        # Each row is (handle, class, process, caption, fullscreen).
        lookup = {r[0]: r for r in windows}
        user32 = Mock()
        user32.IsWindowVisible.return_value = True
        user32.IsIconic.return_value = False
        user32.GetForegroundWindow.return_value = foreground

        def class_name(hwnd, dest, _):
            dest.value = lookup[hwnd][1]
            return len(dest.value)

        def enumerate_windows(visit, context):
            for hwnd in lookup:
                if not visit(hwnd, context):
                    break
            return True

        user32.GetClassNameW.side_effect = class_name
        user32.EnumWindows.side_effect = enumerate_windows
        with patch.object(main.platform, "system", return_value="Windows"), \
             patch.object(main.ctypes, "windll", SimpleNamespace(user32=user32), create=True), \
             patch.object(main, "WNDENUMPROC_TYPE", lambda cb: cb), \
             patch.object(main, "get_window_process_name", lambda h: lookup[h][2]), \
             patch.object(main, "_get_window_title", lambda h: lookup[h][3]), \
             patch.object(main, "is_window_fullscreen", lambda h: lookup[h][4]), \
             patch("slideshow._wps_fullscreen_surface", lambda h: lookup[h][4]):
            return main.find_slideshow_window()

    def test_secondary_audience_wins_when_presenter_or_overlay_has_focus(self):
        windows = [
            (12, "Qt6QWindowIcon", "ppt time.exe", "ppt time", True),
            (21, "Qt5QWindowIcon", "wpp.exe", "WPS 演讲者视图", True),
            (22, "Qt5QWindowIcon", "wpp.exe", "Meeting.dps", True),
        ]
        for hwnd in (12, 21, 22):
            with self.subTest(hwnd=hwnd):
                self.assertEqual(self.detect(windows, hwnd), 22)

    def test_real_audience_class_outranks_generic_office_pane(self):
        windows = [
            (10, "PaneClassDC", "powerpnt.exe", "PowerPoint", True),
            (11, "screenClass", "powerpnt.exe", "PowerPoint Slide Show - [Presenter View.pptx]", True),
        ]
        self.assertEqual(self.detect(windows, 10), 11)

    def test_an_editing_pane_and_presenter_alone_are_not_an_audience_window(self):
        rows = [
            (10, "Mso_Subwindow_Class", "powerpnt.exe", "幻灯片放映.pptx - PowerPoint", False),
            (11, "screenClass", "powerpnt.exe", "PowerPoint Presenter View - [Deck]", True),
        ]
        self.assertEqual(self.detect(rows, 11), 0)
