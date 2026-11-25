import os
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QColor, QPainter, QPdfWriter
from PySide6.QtWidgets import QApplication

from meeting_player import MeetingPlayer, _start_current_windows
from slideshow import _mac_snapshot, _page_labels, _window_snapshot


class MeetingPlayerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_pdf_display_page_order_and_end(self):
        player = MeetingPlayer()
        observed = []
        errors = []
        player.snapshot_changed.connect(observed.append)
        player.error.connect(errors.append)
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as opened:
            path = Path(opened.name)
        self.addCleanup(path.unlink, missing_ok=True)
        writer = QPdfWriter(str(path))
        painter = QPainter(writer)
        for index, color in enumerate(("red", "green", "blue")):
            if index:
                writer.newPage()
            painter.fillRect(0, 0, writer.width(), writer.height(), QColor(color))
        painter.end()
        del writer
        try:
            self.assertTrue(player.open_file(path), errors)
            self.app.processEvents()
            self.assertEqual(observed[-1].document, str(path.resolve()))
            self.assertEqual(observed[-1].page_text, "1/3")
            first = player.window.grab().toImage()
            self.assertGreater(first.pixelColor(first.rect().center()).red(), 220)
            self.assertTrue(player.command("next"))
            self.assertEqual(observed[-1].page_text, "2/3")
            second = player.window.grab().toImage()
            self.assertGreater(second.pixelColor(second.rect().center()).green(), 90)
            self.assertTrue(player.command("last"))
            self.assertEqual(observed[-1].page_text, "3/3")
            self.assertTrue(player.command("first"))
            self.assertEqual(observed[-1].page_text, "1/3")
            self.assertTrue(player.command("next"))
            self.assertTrue(player.command("prev"))
            self.assertEqual(observed[-1].page_text, "1/3")
            self.assertFalse(player.open_file(path.with_suffix(".missing.pdf")))
            self.assertEqual(player.owned_snapshot.page_text, "1/3")
            for expected in ("2/3", "3/3", ""):
                self.assertTrue(player.command("next"))
                self.assertEqual(observed[-1].page_text, expected)
            self.assertTrue(observed[-1].ended)
            self.assertFalse(observed[-1].active)
            self.assertIsNone(player.owned_snapshot)
            self.assertFalse(player.command("next"))
        finally:
            player.close()
            # Deferred deletion releases QPdfDocument's Windows file handle.
            from PySide6.QtCore import QCoreApplication, QEvent
            QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)

    def test_office_snapshot_names_document_and_end_state(self):
        doc = SimpleNamespace(FullName="talk.pptx", Slides=SimpleNamespace(Count=8))
        view = SimpleNamespace(State=1, Slide=SimpleNamespace(SlideIndex=6))
        window = SimpleNamespace(Presentation=doc, View=view)
        snapshot = _window_snapshot(window, ("Monitor", (0, 0)), 765)
        self.assertEqual((snapshot.page_text, snapshot.document, snapshot.hwnd), ("6/8", "talk.pptx", 765))
        view.State = 5
        snapshot = _window_snapshot(window, ("Monitor", (0, 0)), 765)
        self.assertTrue(snapshot.ended)
        self.assertEqual(snapshot.page_text, "")

    def test_mac_parser_retains_document_with_separator(self):
        snapshot = _mac_snapshot("show|3/8|-120,40|/talks/June|Q&A.pptx")
        self.assertEqual(snapshot.document, "/talks/June|Q&A.pptx")
        self.assertEqual(snapshot.page_text, "3/8")
        self.assertEqual(snapshot.screen_point, (-120, 40))
        self.assertTrue(_mac_snapshot("show||").active)
        self.assertFalse(_mac_snapshot("idle").active)

    def test_status_labels_require_a_page_context_and_valid_range(self):
        self.assertEqual(_page_labels("幻灯片 8／24"), {(8, 24)})
        self.assertEqual(_page_labels("第2张，共9张"), {(2, 9)})
        self.assertEqual(_page_labels("Slide 4 of 13"), {(4, 13)})
        self.assertEqual(_page_labels("8/21"), set())
        self.assertEqual(_page_labels("8/21", True), {(8, 21)})
        self.assertEqual(_page_labels("8/3", True), set())
        self.assertEqual(_page_labels("Page 2/4  Page 3/4"), {(2, 4), (3, 4)})

    @unittest.skipUnless(os.name == "nt", "Windows slideshow integration")
    def test_bounded_wps_fallback_and_ambiguous_status(self):
        from slideshow import read_windows_pages
        with patch("slideshow._process_name", return_value="wpp.exe"), \
             patch("slideshow.windows_show_window", return_value=(None, ("", None))), \
             patch("win32gui.IsWindowVisible", return_value=True), \
             patch("win32gui.GetClassName", return_value="KSlideShow"), \
             patch("win32gui.GetWindowText", return_value="WPS - [slide 9 of 12.pptx]"), \
             patch("slideshow._screen_info", return_value=("Projector", (4,2))), \
             patch("slideshow._pid", return_value=7), \
             patch("slideshow._wps_accessible_labels", return_value={(3, 12)}) as access:
            result = read_windows_pages(100)
            self.assertEqual((result.page_text, result.hwnd, result.screen_name), ("3/12", 100, "Projector"))
            access.assert_called_once_with(100, 7)
            access.return_value = {(3, 12), (4, 12)}
            self.assertEqual(read_windows_pages(100).page_text, "")

    @unittest.skipUnless(os.name == "nt", "Windows slideshow integration")
    def test_start_current_uses_live_document_or_rejects_unrelated_editor(self):
        doc = Mock()
        doc.Slides.Count = 5
        app = SimpleNamespace(ActivePresentation=doc, HWND=10, SlideShowWindows=SimpleNamespace(Count=0))
        with patch("win32gui.GetForegroundWindow", return_value=12), \
             patch("meeting_player._pid", return_value=8), \
             patch("meeting_player._running_office_apps", return_value=[app]):
            _start_current_windows()
            doc.SlideShowSettings.Run.assert_called_once_with()
            app.SlideShowWindows.Count = 1
            _start_current_windows()
            doc.SlideShowSettings.Run.assert_called_once_with()
        with patch("win32gui.GetForegroundWindow", return_value=51), \
             patch("meeting_player._pid", return_value=8), \
             patch("meeting_player._running_office_apps", return_value=[]), \
             patch("win32gui.IsWindowVisible", return_value=True), \
             patch("meeting_player._process_name", return_value="notepad.exe"), \
             patch("win32gui.EnumWindows", side_effect=lambda fn, arg: fn(51, arg)), \
             patch("win32gui.PostMessage") as send:
            with self.assertRaises(ValueError):
                _start_current_windows()
            send.assert_not_called()
