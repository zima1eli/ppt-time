import sys
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

if sys.platform == "win32":
    import win32con
    from PySide6.QtCore import QRectF
    from PySide6.QtGui import QPainterPath
    import projection_context as context


@unittest.skipUnless(sys.platform == "win32", "Windows projection menu")
class ProjectionContextTests(unittest.TestCase):
    def test_right_click_repeats_and_drag_out_or_modified_click_passes(self):
        opened = Mock()
        click = context._StripClick(opened)
        position = SimpleNamespace(x=25, y=5)
        hit = (123, 25, 5)
        with patch.object(context, "_hit_strip", return_value=hit) as find, \
                patch.object(context.win32api, "GetAsyncKeyState", return_value=0) as keys:
            def mouse(message):
                return click.handle(message, (), position)
            for _ in range(3):
                self.assertTrue(mouse(win32con.WM_RBUTTONDOWN))
                self.assertTrue(mouse(win32con.WM_RBUTTONUP))
            self.assertEqual(opened.call_count, 3)
            self.assertTrue(mouse(win32con.WM_RBUTTONDOWN))
            find.return_value = None
            self.assertTrue(mouse(win32con.WM_RBUTTONUP))
            self.assertFalse(mouse(win32con.WM_RBUTTONDOWN))
            find.return_value = hit
            self.assertFalse(mouse(win32con.WM_RBUTTONUP))
            for modifier in (win32con.VK_CONTROL, win32con.VK_SHIFT):
                keys.side_effect = lambda key: 0x8000 if key == modifier else 0
                self.assertFalse(mouse(win32con.WM_RBUTTONDOWN))
                self.assertFalse(mouse(win32con.WM_RBUTTONUP))
            self.assertEqual(opened.call_count, 3)

    def test_hit_shape_native_coordinates_and_window_above(self):
        path = QPainterPath()
        path.addEllipse(QRectF(0, 0, 1, 1))
        targets = ((2, path),)
        with patch.object(context.win32gui, "IsWindowVisible", return_value=True), \
                patch.object(context.win32gui, "GetWindowRect", return_value=(-500, 80, -300, 180)), \
                patch.object(context.win32gui, "WindowFromPoint", return_value=12), \
                patch.object(context.win32gui, "GetAncestor", return_value=10), \
                patch.object(context.win32gui, "GetWindow", return_value=2) as above:
            self.assertEqual(context._hit_strip(targets, -400, 130), (2, 100, 50))
            self.assertIsNone(context._hit_strip(targets, -498, 81))
            self.assertIsNone(context._hit_strip(targets, -300, 130))
            above.return_value = 0
            self.assertIsNone(context._hit_strip(targets, -400, 130))
