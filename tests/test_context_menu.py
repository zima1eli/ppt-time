import ctypes
import os
import sys
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtCore import QObject, QPoint, Qt, QTimer
from PySide6.QtGui import QContextMenuEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QDialog, QSystemTrayIcon

from main import App, FloatBar


class MenuOwner(QObject):
    """Install the production tray menu without launching Office or registering hotkeys."""

    def __init__(self):
        super().__init__()
        self.float_bar = FloatBar()
        self.file_dialogs = 0
        self.conference = SimpleNamespace(open_panel=Mock())
        self.lan = self.cloud = self.operator = SimpleNamespace(open_settings=Mock())
        self.show_timing = SimpleNamespace(edit=Mock(), dismiss=Mock())
        self.display = SimpleNamespace(edit_appearance=Mock())
        self.powerpoint = self.wps_pdf = self.wps_presentation = SimpleNamespace(
            install=Mock(), status=Mock(), uninstall=Mock())
        self._update_tray_status = Mock()
        self._open_settings = Mock()
        self._exit = Mock()
        self._copy_page_diagnostics = Mock()
        self._tray_activated = Mock()
        self._addin_action = Mock()
        with patch.object(QSystemTrayIcon, "show"):
            App._create_tray(self)

    def _pick_file(self):
        dialog = QDialog()
        QTimer.singleShot(10, dialog.reject)
        dialog.exec()
        self.file_dialogs += 1


class ContextMenuTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.owner = MenuOwner()
        self.bar = self.owner.float_bar
        self.menu = self.owner.tray.contextMenu()
        self.bar.show()
        QTest.qWait(30)
        self.addCleanup(self.owner.deleteLater)
        self.addCleanup(self.bar.close)
        self.addCleanup(self.menu.close)

    def open_bar_menu(self):
        point = QPoint(35, 24)
        if sys.platform == "win32" and self.app.platformName() == "windows":
            from ctypes import wintypes
            post = ctypes.windll.user32.PostMessageW
            post.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
            hwnd = int(self.bar.winId())
            packed = (point.y() << 16) | point.x()
            for message, buttons in ((0x0204, 2), (0x0205, 0)):
                self.assertTrue(post(hwnd, message, buttons, packed))
        else:
            event = QContextMenuEvent(QContextMenuEvent.Reason.Mouse, point, self.bar.mapToGlobal(point))
            QApplication.sendEvent(self.bar, event)
        QTest.qWait(50)
        self.assertIs(self.app.activePopupWidget(), self.menu)

    def select(self, menu, label):
        action = next(action for action in menu.actions() if action.text() == label)
        menu.setActiveAction(action)
        QTest.keyClick(menu, Qt.Key.Key_Return)
        QTest.qWait(30)
        self.assertFalse(self.menu.isVisible())

    def test_reopens_after_cancel_and_after_modal_action(self):
        self.open_bar_menu()
        QTest.keyClick(self.menu, Qt.Key.Key_Escape)
        for _ in range(2):
            self.open_bar_menu()
            self.select(self.menu, "打开放映文件…")
        self.assertEqual(self.owner.file_dialogs, 2)
        self.open_bar_menu()
        self.select(self.menu, "会议控制台…")
        self.owner.conference.open_panel.assert_called_once()

    def test_tray_submenu_and_hidden_bar_popup_share_live_actions(self):
        self.open_bar_menu()
        self.select(self.menu, "打开放映文件…")
        self.bar.hide()
        self.bar.show_context_menu_at(QPoint(100, 100))
        QTest.qWait(30)
        self.assertIs(self.app.activePopupWidget(), self.menu)
        action = next(action for action in self.menu.actions() if action.text() == "PowerPoint 页码联动")
        self.menu.setActiveAction(action)
        QTest.keyClick(self.menu, Qt.Key.Key_Right)
        QTest.qWait(40)
        self.assertTrue(action.menu().isVisible())
        self.select(action.menu(), "查看联动状态")
        self.owner._addin_action.assert_called_once_with("PowerPoint 页码联动", self.owner.powerpoint.status)
        self.bar.show()
        self.open_bar_menu()
        self.select(self.menu, "计时器设置…")
        self.owner._open_settings.assert_called_once()


if __name__ == "__main__":
    unittest.main()
