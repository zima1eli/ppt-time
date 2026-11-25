import os
from pathlib import Path
import tempfile
import threading
import time
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication

from meeting_player import MeetingPlayer, _open_windows_deck, _read_opened_show
from slideshow import SlideSnapshot


@unittest.skipUnless(os.name == "nt", "Windows Office automation")
class OpenedOfficeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def wait_until(self, condition):
        limit = time.monotonic() + 3
        while not condition() and time.monotonic() < limit:
            self.app.processEvents()
            time.sleep(.015)
        self.assertTrue(condition())

    def test_selected_show_emits_live_pages_then_clears_without_rot(self):
        with tempfile.NamedTemporaryFile(suffix=".pptx") as fixture:
            path = Path(fixture.name).resolve()
            doc = SimpleNamespace(FullName=str(path), Slides=SimpleNamespace(Count=8), Close=Mock())
            view = SimpleNamespace(State=1, Slide=SimpleNamespace(SlideIndex=1))
            view.Next = Mock(side_effect=lambda: setattr(view.Slide, "SlideIndex", 6))
            view.Exit = Mock(side_effect=lambda: setattr(view, "State", 5))
            show = SimpleNamespace(Presentation=doc, View=view, HWND=123)
            office = SimpleNamespace(SlideShowWindows=SimpleNamespace(Count=1))
            opened = (office, doc, show, False)
            reader_threads = []

            def read(owned, document):
                reader_threads.append(threading.get_ident())
                return _read_opened_show(owned, document)

            with patch("meeting_player._open_windows_deck", return_value=opened), \
                 patch("meeting_player._read_opened_show", side_effect=read), \
                 patch("meeting_player._screen_info", return_value=("Projector", (10, 10))), \
                 patch("win32gui.IsWindowVisible", return_value=True), \
                 patch("win32gui.IsIconic", return_value=False), \
                 patch("win32gui.GetForegroundWindow", return_value=123), \
                 patch("meeting_player._windows_command", side_effect=AssertionError("Unexpected ROT lookup")):
                player = MeetingPlayer()
                observed, errors = [], []
                player.snapshot_changed.connect(observed.append)
                player.error.connect(errors.append)
                try:
                    self.assertTrue(player.open_file(path))
                    self.wait_until(lambda: player.owned_snapshot is not None)
                    self.assertEqual(player.owned_snapshot.page_text, "1/8")
                    self.assertTrue(player.command("next"))
                    self.wait_until(lambda: player.owned_snapshot and player.owned_snapshot.page_text == "6/8")
                    view.Slide.SlideIndex = 8  # physical keyboard or click inside PowerPoint
                    self.wait_until(lambda: player.owned_snapshot and player.owned_snapshot.page_text == "8/8")
                    view.State = 5
                    self.wait_until(lambda: player.owned_snapshot is None)
                    self.assertTrue(observed[-1].ended)
                    self.assertEqual(observed[-1].document, str(path))
                    view.State = 1
                    self.wait_until(lambda: player.owned_snapshot is not None)
                    self.assertTrue(player.command("stop"))
                    self.wait_until(lambda: player.owned_snapshot is None)
                    self.assertEqual(view.Exit.call_count, 1)
                    self.assertFalse(doc.Close.called)
                    self.assertEqual(errors, [])
                    self.assertEqual(len(set(reader_threads)), 1)
                    self.assertNotEqual(reader_threads[0], threading.get_ident())
                finally:
                    player.close()
                    if player._thread:
                        player._thread.join(1)

    def test_superseded_and_unreadable_results_never_mask_external_pages(self):
        player = MeetingPlayer()
        try:
            old = SlideSnapshot(True, 2, 5, document="formerly.pptx")
            player._generation = 3
            player._accept_office_sample(2, old)
            self.assertIsNone(player.owned_snapshot)
            player._accept_office_sample(3, old)
            player._accept_office_sample(3, SlideSnapshot(True, document="formerly.pptx"))
            self.assertIsNone(player.owned_snapshot)
        finally:
            player.close()

    def test_reuses_an_already_open_and_running_document(self):
        path = Path("already-open.pptx").resolve()
        slide_show = SimpleNamespace(View=SimpleNamespace(State=1))
        doc = SimpleNamespace(FullName=str(path), SlideShowWindow=slide_show, SlideShowSettings=Mock())
        presentations = Mock()
        presentations.Count = 1
        presentations.Item.return_value = doc
        app = SimpleNamespace(Presentations=presentations)
        with patch("win32com.client.dynamic.Dispatch", return_value=app):
            self.assertEqual(_open_windows_deck(path), (app, doc, slide_show, False))
        presentations.Open.assert_not_called()
        doc.SlideShowSettings.Run.assert_not_called()
