import csv
import os
from pathlib import Path
import shutil
import tempfile
import unittest
import uuid
from zipfile import ZipFile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtWidgets import QApplication

from agenda_manager import AgendaManager, filter_sessions, read_schedule
from conference import Board, ConferenceDesk, Session
from slideshow import SlideSnapshot


class AgendaTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.root = Path(tempfile.gettempdir()) / ("agenda-" + uuid.uuid4().hex)
        self.root.mkdir()
        self.addCleanup(shutil.rmtree, self.root)
        self.desk = ConferenceDesk(self.root / "state.json")
        self.addCleanup(self.desk.close)

    def test_csv_accepts_chinese_and_validates_before_returning(self):
        csv_path = self.root / "meeting.csv"
        with csv_path.open("w", newline="", encoding="gb18030") as output:
            writer = csv.writer(output)
            writer.writerow(["议题", "会议", "会场", "日期", "开始时间", "讲者", "单位", "时长（分钟）", "状态"])
            writer.writerow(["新进展", "年会", "A", "2026/11/24", "9:05", "李明", "四川", "1.5", "待播放"])
        session, = read_schedule(csv_path)
        self.assertEqual((session.title, session.date, session.start, session.seconds, session.status),
                         ("新进展", "2026-11-24", "09:05", 90, "待开始"))
        self.assertEqual(filter_sessions([session], meeting="年会", room="A", date="2026-11-24",
                                         status="待开始", search="四川"), [(0, session)])
        with csv_path.open("a", encoding="gb18030") as output:
            output.write("bad,,,2026-02-30,,,,,")
        with self.assertRaisesRegex(ValueError, "第 3 行"):
            read_schedule(csv_path)

    def test_xlsx_shared_inline_sparse_and_excel_dates(self):
        path = self.root / "source.xlsx"
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        with ZipFile(path, "w") as x:
            x.writestr("xl/workbook.xml", f'<workbook xmlns="{ns}" xmlns:r="{rel}"><workbookPr date1904="1"/><sheets><sheet r:id="rId7"/></sheets></workbook>')
            x.writestr("xl/_rels/workbook.xml.rels", f'<Relationships><Relationship Id="rId7" Type="{rel}/worksheet" Target="worksheets/sheet5.xml"/></Relationships>')
            x.writestr("xl/sharedStrings.xml", f'<sst xmlns="{ns}"><si><r><t>主</t></r><r><t>题</t></r></si></sst>')
            x.writestr("xl/worksheets/sheet5.xml", f'''<worksheet xmlns="{ns}"><sheetData>
                <row><c r="A1" t="s"><v>0</v></c><c r="C1" t="inlineStr"><is><t>日期</t></is></c><c r="D1" t="inlineStr"><is><t>开始时间</t></is></c></row>
                <row><c r="A2" t="inlineStr"><is><t>Opening</t></is></c><c r="C2"><v>1</v></c><c r="D2"><v>0.375</v></c></row></sheetData></worksheet>''')
        session, = read_schedule(path)
        self.assertEqual((session.title, session.date, session.start), ("Opening", "1904-01-02", "09:00"))

    def test_metadata_roundtrip_and_playback(self):
        row = {"title": "讲座", "speaker": "甲", "start": "10:00", "seconds": 30, "file": "",
               "meeting": "年会", "room": "B", "date": "2026-11-24", "organization": "大学", "status": "待开始"}
        self.assertTrue(self.desk.import_agenda([row]))
        self.assertEqual(self.desk.status()["agenda"], [{**row, "uid": self.desk.board.agenda[0].uid}])
        self.assertEqual(Board.load(self.desk.path).meetings, ["年会"])
        self.assertTrue(self.desk.command("agenda", 0))
        self.assertEqual(self.desk.status()["agenda"][0]["status"], "进行中")
        self.desk.set_snapshot(SlideSnapshot(True, 1, 2))
        self.desk.set_snapshot(None)
        self.assertEqual(Board.load(self.desk.path).agenda[0].status, "已完成")
        self.assertFalse(self.desk.import_agenda([{**row, "date": "2026-02-31"}]))
        self.assertEqual(self.desk.board.agenda[0].status, "已完成")

    def test_manager_views_filter_and_selection_original_index(self):
        self.desk.update_board(agenda=[Session("Later", "12:00", room="A"),
                                       Session("Earlier", "08:00", room="B")])
        manager = AgendaManager(self.desk)
        self.addCleanup(manager.deleteLater)
        self.assertEqual(manager.table.rowCount(), 2)
        manager.filters["room"].setCurrentIndex(manager.filters["room"].findData("B"))
        self.assertEqual(manager.table.rowCount(), 1)
        manager.table.selectRow(0)
        self.assertEqual(manager.selected(), [1])
        manager.filters["room"].setCurrentIndex(0)
        manager.mode.setCurrentIndex(2)
        manager.cards.setCurrentRow(0)
        self.assertEqual(manager.selected(), [1])
        manager.mode.setCurrentIndex(1)
        self.assertEqual(manager.cards.count(), 2)

    def test_room_sync_preserves_bound_file_and_live_row_for_existing_uid(self):
        file = self.root / "talk.pptx"
        file.touch()
        entry = Session("Talk", document=str(file), room="A")
        self.desk.update_board(agenda=[entry])
        self.desk.command("agenda", 0)
        edited = {**self.desk.status()["agenda"][0], "speaker": "Jane"}
        other = Session("New").public()
        self.assertTrue(self.desk.import_agenda([other, edited]))
        self.assertEqual(self.desk.active_session, 1)
        self.assertEqual(self.desk.board.agenda[1].document, str(file))
        self.assertFalse(self.desk.import_agenda([other, other]))
        self.assertFalse(self.desk.import_agenda([{**other, "uid": 42}]))
        edited["file"] = "different.pptx"
        self.assertTrue(self.desk.import_agenda([edited]))
        self.assertEqual(self.desk.board.agenda[0].document, "")
