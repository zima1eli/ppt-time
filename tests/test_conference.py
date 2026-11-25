import json
import os
import shutil
import tempfile
import unittest
import uuid
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QColor, QImage
from PySide6.QtWidgets import QApplication, QPushButton

from conference import Board, ConferenceDesk, Session
from slideshow import SlideSnapshot


class ConferenceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.root = Path(tempfile.gettempdir()) / f"ppt-time-board-{uuid.uuid4().hex}"
        self.root.mkdir()
        self.addCleanup(shutil.rmtree, self.root)
        self.state = self.root / "board.json"
        self.desk = ConferenceDesk(self.state)
        self.addCleanup(self.desk.close)

    def picture(self, name):
        path = self.root / name
        img = QImage(32, 18, QImage.Format.Format_RGB32)
        img.fill(QColor("#609798"))
        self.assertTrue(img.save(str(path)))
        return str(path)

    def test_scenes_real_window_snapshot_and_return(self):
        a, b = self.picture("a.png"), self.picture("b.png")
        d = self.desk
        self.assertTrue(d.update_board(home=a, images=[a, b], auto_home=True, safe_end=True))
        w, = d._windows.values()
        self.assertTrue(w.isVisible())
        self.assertFalse(w.grab().isNull())
        self.assertEqual(w.geometry(), QApplication.primaryScreen().geometry())
        d.set_snapshot(SlideSnapshot(True, 1, 3))
        self.assertEqual(d.scene, "none")
        d.set_snapshot(SlideSnapshot(True, 3, 3))
        self.assertEqual(d.scene, "none")
        d.set_snapshot(SlideSnapshot(True, ended=True))
        self.assertEqual(d.scene, "home")
        d.set_snapshot(SlideSnapshot(True, 3, 3))
        self.assertEqual(d.scene, "none")
        self.assertTrue(d.command("previous"))
        self.assertEqual(d.index, 1)
        d.set_snapshot(SlideSnapshot(True, ended=True))
        d.set_snapshot(None)
        self.assertEqual(d.scene, "interstitial")
        self.assertTrue(d.command("next"))
        self.assertEqual(d.index, 0)
        d.set_snapshot(SlideSnapshot(True, 1, 2))
        d.set_snapshot(None)
        self.assertEqual(d.scene, "home")
        self.assertTrue(d.command("hide"))
        self.assertFalse(d._windows)

    def test_local_file_board_and_agenda_signals_persist(self):
        path = self.root / "opening.pptx"
        path.write_bytes(b"test")
        d = self.desk
        self.assertTrue(d.add_files([path, path, self.root / "missing"]))
        one = Session("Opening", "09:30", "Alice", str(path), 90)
        two = Session("Closing", seconds=300)
        self.assertTrue(d.update_board(agenda=[one, two]))
        texts, durations = [], []
        d.open_file_requested.connect(texts.append)
        d.duration_requested.connect(durations.append)
        self.assertTrue(d.command("agenda", 0))
        self.assertEqual(texts, [str(path)])
        self.assertEqual(durations, [90])
        self.assertTrue(d.update_board(agenda=[two, one]))
        self.assertEqual(d.active_session, 1)
        self.assertEqual(Board.load(self.state), d.board)
        self.assertNotIn(str(self.root), json.dumps(d.status()))
        path.unlink()
        self.assertFalse(d.command("agenda", 1))
        self.assertEqual(texts, [str(path)])

    def test_remote_agenda_rebinds_only_local_matching_files(self):
        d = self.desk
        row = {"title": "Talk", "start": "13:00", "speaker": "Alex", "seconds": 300,
               "file": "new.pptx"}
        self.assertTrue(d.import_agenda([row]))
        self.assertEqual(d.status()["agenda"], [{**row, "uid": d.board.agenda[0].uid}])
        self.assertFalse(d.command("agenda", 0))
        for unsafe in ("../../outside.pptx", "C:outside.pptx"):
            self.assertFalse(d.import_agenda([{**row, "file": unsafe}]))
        path = self.root / "new.pptx"
        path.touch()
        d.add_files([path])
        opened = []
        d.open_file_requested.connect(opened.append)
        self.assertTrue(d.command("agenda", 0))
        self.assertEqual(opened, [str(path)])
        second = self.root / "another" / path.name
        second.parent.mkdir()
        second.touch()
        d.add_files([second])
        self.assertFalse(d.command("agenda", 0))

    def test_invalid_config_missing_image_and_write_failure(self):
        self.state.write_text('{"images":42,"agenda":[null],"home":99}', encoding="utf-8")
        self.assertEqual(Board.load(self.state), Board())
        d = self.desk
        self.assertTrue(d.update_board(images=[str(self.root/"gone.png")]))
        self.assertFalse(d.command("interstitial", 0))
        self.assertFalse(d.command("interstitial", -1))
        self.assertEqual(d.scene, "none")
        self.assertTrue(d.command("home"))
        self.assertFalse(d.command("file", True))
        original = self.state.read_bytes()
        with patch.object(Path, "mkdir", side_effect=OSError("read only")):
            self.assertFalse(d.update_board(title="temporary"))
        self.assertEqual(self.state.read_bytes(), original)
        self.assertEqual(d.board.title, "")

    def test_panel_and_reordered_playlist(self):
        d = self.desk
        a, b = self.picture("one.png"), self.picture("two.png")
        d.open_panel()
        self.assertEqual(d.panel.tabs.count(), 4)
        self.assertTrue(d.update_board(images=[a,b]))
        self.assertTrue(d.command("interstitial", 0))
        d.panel.images.setCurrentRow(0)
        d.panel._move("images", 1)
        self.assertEqual(d.board.images, [b,a])
        self.assertEqual(d.index, 1)
        d.panel._remove("images")
        self.assertEqual(d.scene, "none")

    def test_project_agenda_paginate_and_follow_active_session(self):
        d = self.desk
        rows = [Session(f"Talk {i} " + "long " * 30, "10:20", "Speaker") for i in range(15)]
        self.assertTrue(d.update_board(title="Track B", agenda=rows))
        d.open_panel()
        project = next(b for b in d.panel.findChildren(QPushButton) if b.text() == "投放议程看板")
        project.click()
        canvas, = d._windows.values()
        self.assertTrue(canvas.isVisible())
        self.assertEqual(d.status()["board_pages"], 3)
        self.assertTrue(d.command("board_next"))
        self.assertEqual(canvas._page, 1)
        self.assertTrue(d.command("board_previous"))
        self.assertTrue(d.command("agenda", 14))
        self.assertEqual((canvas._active_row, canvas._page), (14, 2))
        canvas.resize(720, 420)
        image = canvas.grab().toImage()
        colors = {image.pixelColor(x, y).name() for x in range(30, 700, 20) for y in range(10, 411, 5)}
        self.assertIn("#21495a", colors)
        self.assertTrue(d.command("board_next"))
        self.assertEqual(canvas._page, 2)
        self.assertTrue(d.update_board(agenda=[]))
        self.assertEqual((canvas._page, canvas._active_row), (0, -1))
        self.assertTrue(d.command("home"))
        self.assertIsNone(canvas._agenda)
        self.assertTrue(d.command("hide"))
        self.assertFalse(d._windows)
