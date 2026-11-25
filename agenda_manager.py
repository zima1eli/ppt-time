"""Operator-facing multi-room schedule management and CSV/XLSX intake."""

import csv
from dataclasses import replace
from datetime import datetime, timedelta
import io
from pathlib import Path
import posixpath
import re
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView, QComboBox, QDialog, QFileDialog,
    QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QListView, QListWidget,
    QListWidgetItem, QMessageBox, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QStackedWidget,
)

from conference import SESSION_STATES, Session, SessionEditor, _button


HEADERS = {
    "title": ("议题", "主题", "报告题目", "演讲题目", "日程名称", "议程名称", "title", "subject"),
    "meeting": ("会议名称", "会议", "meeting"),
    "room": ("会场", "会场名称", "分会场", "room"),
    "date": ("会议日期", "日期", "date"),
    "start": ("开始时间", "时间", "开始", "start"),
    "speaker": ("讲者", "演讲者", "报告人", "姓名", "speaker"),
    "organization": ("单位", "讲者单位", "机构", "organization", "affiliation"),
    "status": ("状态", "日程状态", "播放状态", "status"),
    "file_name": ("文件名", "附件", "文件", "file"),
    "seconds": ("时长秒", "时长(秒)", "时长（秒）", "秒", "seconds"),
    "minutes": ("时长", "分钟", "时长(分钟)", "时长（分钟）", "minutes"),
}
HEADER_MAP = {alias.casefold(): field for field, aliases in HEADERS.items() for alias in aliases}
XML = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
STATUS_ALIASES = {"未开始": "待开始", "待播放": "待开始", "播放中": "进行中", "已结束": "已完成",
                  "pending": "待开始", "active": "进行中", "done": "已完成"}


def _xlsx_rows(path):
    with ZipFile(path) as book:
        def xml(name):
            info = book.getinfo(name)
            if info.file_size > 32 * 1024 * 1024:
                raise ValueError("表格内容过大，请只保留需要导入的日程。")
            return ElementTree.fromstring(book.read(info))
        workbook = xml("xl/workbook.xml")
        props = workbook.find(XML + "workbookPr")
        uses_1904 = props is not None and props.get("date1904") in ("1", "true")
        links = {el.get("Id"): el.get("Target") for el in xml("xl/_rels/workbook.xml.rels")
                 if el.get("Type", "").endswith("/worksheet") and el.get("TargetMode") != "External"}
        sheet = next((s for s in workbook.iter(XML + "sheet")
                      if s.get("state", "visible") == "visible" and s.get(REL) in links), None)
        if sheet is None:
            raise ValueError("工作簿没有可见的工作表。")
        part = links[sheet.get(REL)]
        sheet_path = part.lstrip("/") if part.startswith("/") else posixpath.normpath(posixpath.join("xl", part))
        shared = []
        if "xl/sharedStrings.xml" in book.namelist():
            shared = ["".join(t.text or "" for t in item.iter(XML + "t"))
                      for item in xml("xl/sharedStrings.xml").findall(XML + "si")]
        rows = []
        for row in xml(sheet_path).iter(XML + "row"):
            values = []
            for cell in row.findall(XML + "c"):
                ref = re.match("[A-Z]+", cell.get("r", ""))
                col = len(values)
                if ref:
                    col = 0
                    for ch in ref[0]:
                        col = col * 26 + ord(ch) - 64
                    col -= 1
                if not 0 <= col < 128:
                    continue
                val = cell.findtext(XML + "v", "")
                if cell.get("t") == "s":
                    number = int(val)
                    if not 0 <= number < len(shared):
                        raise ValueError("Excel 共享文本索引无效。")
                    val = shared[number]
                elif cell.get("t") == "inlineStr":
                    val = "".join(t.text or "" for t in cell.iter(XML + "t"))
                if col >= len(values):
                    values.extend([""] * (col - len(values) + 1))
                values[col] = val
            if any(values):
                rows.append(values)
                if len(rows) > 2001:
                    raise ValueError("单次最多导入 2000 条议程。")
    return rows, uses_1904


def _csv_rows(path):
    if path.stat().st_size > 12 * 1024 * 1024:
        raise ValueError("文件过大，请只保留会议日程。")
    data = path.read_bytes()
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        content = data.decode("utf-16")
    else:
        try:
            content = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            content = data.decode("gb18030")
    try:
        dialect = csv.Sniffer().sniff(content[:8192], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    return [row for row in csv.reader(io.StringIO(content, newline=""), dialect) if any(row)]


def _date(raw, uses_1904=False):
    if not raw:
        return ""
    if re.fullmatch(r"\d{1,5}(?:\.\d+)?", raw):
        epoch = datetime(1904, 1, 1) if uses_1904 else datetime(1899, 12, 30)
        return (epoch + timedelta(days=int(float(raw)))).date().isoformat()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y%m%d", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(raw, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError("无效日期 " + raw)


def _time(raw):
    if not raw:
        return ""
    if re.fullmatch(r"0(?:\.\d+)?", raw):
        minutes = round(float(raw) * 1440) % 1440
        return f"{minutes // 60:02d}:{minutes % 60:02d}"
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(raw, fmt).strftime("%H:%M")
        except ValueError:
            pass
    raise ValueError("无效开始时间 " + raw)


def read_schedule(path):
    path = Path(path)
    use_1904 = False
    if path.suffix.lower() == ".xlsx":
        try:
            rows, use_1904 = _xlsx_rows(path)
        except (KeyError, IndexError, BadZipFile, ElementTree.ParseError) as error:
            raise ValueError("无法读取此 XLSX 工作簿。") from error
    elif path.suffix.lower() == ".csv":
        rows = _csv_rows(path)
    else:
        raise ValueError("请选择 CSV 或 XLSX 文件。")
    if not rows or len(rows) > 2001:
        raise ValueError("文件应包含表头及不超过 2000 条议程。")
    headers = [HEADER_MAP.get(str(col).strip().casefold().replace(" ", "")) for col in rows[0]]
    if "title" not in headers:
        raise ValueError("首行需要“议题”或“主题”列。可另加会议、会场、日期、开始时间、讲者、单位、时长（分钟）、状态、文件名。")
    result = []
    for number, row in enumerate(rows[1:], 2):
        values = {key: val.strip() for key, val in zip(headers, row) if key and val.strip()}
        if not values:
            continue
        try:
            if not values.get("title"):
                raise ValueError("缺少议题")
            values["start"] = _time(values.get("start", ""))
            values["date"] = _date(values.get("date", ""), use_1904)
            state = values.get("status", "")
            state = STATUS_ALIASES.get(state.casefold(), state)
            if state and state not in SESSION_STATES:
                raise ValueError("状态应为待开始、进行中或已完成")
            values["status"] = state
            minute_value = values.pop("minutes", "0")
            seconds = round(float(values["seconds"])) if "seconds" in values else round(float(minute_value)*60)
            if not 0 <= seconds <= 86400:
                raise ValueError("时长超出 0 至 24 小时")
            values["seconds"] = seconds
            if "file_name" in values:
                filename = values["file_name"].replace("\\", "/").rsplit("/", 1)[-1]
                if not filename or filename in (".", "..") or ":" in filename:
                    raise ValueError("附件名无效")
                values["file_name"] = filename
            result.append(Session(**values))
        except (TypeError, ValueError, OverflowError) as exc:
            raise ValueError(f"第 {number} 行：{exc}") from exc
    if not result:
        raise ValueError("没有可导入的议程。")
    return result


def filter_sessions(agenda, meeting="", room="", date="", status="", search=""):
    needle = search.strip().casefold()
    return [(i, entry) for i, entry in enumerate(agenda)
            if (not meeting or meeting == entry.meeting)
            and (not room or room == entry.room) and (not date or date == entry.date)
            and (not status or status == (entry.status or SESSION_STATES[0]))
            and (not needle or any(needle in val.casefold() for val in
                                   (entry.title, entry.speaker, entry.organization, entry.meeting, entry.room)))]


class AgendaManager(QDialog):
    def __init__(self, desk, parent=None):
        super().__init__(parent)
        self.desk = desk
        self.setWindowTitle("ppt time · 日程管理")
        self.resize(960, 570)
        outer = QVBoxLayout(self)
        top = QHBoxLayout()
        self.filters = {}
        for key, name in (("meeting", "会议"), ("room", "会场"), ("date", "日期"), ("status", "状态")):
            box = QComboBox()
            box.setAccessibleName(name)
            box.currentIndexChanged.connect(self.draw)
            self.filters[key] = box
            top.addWidget(box)
        self.search = QLineEdit()
        self.search.setPlaceholderText("搜索议题、讲者或单位")
        self.search.textChanged.connect(self.draw)
        top.addWidget(self.search, 1)
        self.mode = QComboBox()
        self.mode.addItems(["表格", "卡片", "时间线"])
        self.mode.currentIndexChanged.connect(self.draw)
        top.addWidget(self.mode)
        outer.addLayout(top)
        self.summary = QLabel()
        outer.addWidget(self.summary)
        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(["日期", "开始", "议题", "讲者", "单位", "会议", "会场", "状态"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.cellDoubleClicked.connect(lambda *_: self.edit(True))
        self.cards = QListWidget()
        self.cards.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.cards.setResizeMode(QListView.ResizeMode.Adjust)
        self.cards.setMovement(QListView.Movement.Static)
        self.cards.itemDoubleClicked.connect(lambda *_: self.edit(True))
        self.content = QStackedWidget()
        self.content.addWidget(self.table)
        self.content.addWidget(self.cards)
        outer.addWidget(self.content, 1)
        buttons = QHBoxLayout()
        for title, handler in (("添加", lambda: self.edit(False)), ("编辑", lambda: self.edit(True)),
                               ("删除选中", self.remove), ("批量修改…", self.batch),
                               ("管理会议…", self.groups), ("导入 CSV / XLSX…", self.import_rows)):
            buttons.addWidget(_button(title, handler))
        buttons.addStretch()
        outer.addLayout(buttons)
        self.desk.changed.connect(self.reload)
        self.reload()

    def reload(self):
        rows = self.desk.board.agenda
        values = {key: sorted({getattr(row, key) for row in rows if getattr(row, key)})
                  for key in ("meeting", "room", "date")}
        values["meeting"] = sorted(set(self.desk.board.meetings) | set(values["meeting"]))
        values["status"] = SESSION_STATES
        for key, box in self.filters.items():
            selected = box.currentData()
            box.blockSignals(True)
            box.clear()
            box.addItem("全部" + box.accessibleName(), "")
            for val in values[key]:
                box.addItem(val, val)
            box.setCurrentIndex(max(0, box.findData(selected)))
            box.blockSignals(False)
        self.draw()

    def draw(self, *_):
        rows = filter_sessions(self.desk.board.agenda, search=self.search.text(),
                               **{k: box.currentData() or "" for k, box in self.filters.items()})
        counts = {status: sum((r.status or SESSION_STATES[0]) == status for _, r in rows)
                  for status in SESSION_STATES}
        self.summary.setText("  ·  ".join([f"{len(rows)} 项议程", *
                             (f"{s} {n}" for s, n in counts.items()),
                             f"已关联文件 {sum(bool(r.document or r.file_name) for _,r in rows)}"]))
        mode = self.mode.currentIndex()
        self.content.setCurrentIndex(0 if mode == 0 else 1)
        if mode == 2:
            rows.sort(key=lambda pair: (pair[1].date, pair[1].start, pair[1].room, pair[0]))
        if mode == 0:
            self.table.setRowCount(len(rows))
            for i, (original, r) in enumerate(rows):
                for j, field in enumerate((r.date, r.start, r.title, r.speaker, r.organization, r.meeting,
                                           r.room, r.status or SESSION_STATES[0])):
                    item = QTableWidgetItem(field)
                    item.setData(Qt.ItemDataRole.UserRole, original)
                    self.table.setItem(i, j, item)
        else:
            self.cards.clear()
            self.cards.setViewMode(QListView.ViewMode.IconMode if mode == 1 else QListView.ViewMode.ListMode)
            self.cards.setWordWrap(mode == 1)
            self.cards.setSpacing(8 if mode == 1 else 3)
            for original, r in rows:
                time = " ".join(x for x in (r.date, r.start) if x)
                subject = " · ".join(x for x in (r.title, r.speaker, r.organization) if x)
                group = " · ".join(x for x in (r.meeting, r.room, r.status or SESSION_STATES[0]) if x)
                label = "\n".join(x for x in (time, subject, group) if x) if mode == 1 else "    ".join(x for x in (time, subject, group) if x)
                item = QListWidgetItem(label)
                item.setData(Qt.ItemDataRole.UserRole, original)
                self.cards.addItem(item)

    def selected(self):
        if self.mode.currentIndex() == 0:
            return sorted({item.data(Qt.ItemDataRole.UserRole) for item in self.table.selectedItems()})
        return sorted(item.data(Qt.ItemDataRole.UserRole) for item in self.cards.selectedItems())

    def save(self, rows, meetings=None):
        names = self.desk.board.meetings if meetings is None else meetings
        names = list(dict.fromkeys([*names, *(row.meeting for row in rows if row.meeting)]))
        files = list(dict.fromkeys([*self.desk.board.files, *(row.document for row in rows if row.document)]))
        if not self.desk.update_board(agenda=rows, meetings=names, files=files):
            QMessageBox.warning(self, self.windowTitle(), self.desk.last_error)

    def edit(self, existing):
        selected = self.selected()
        if existing and len(selected) != 1:
            QMessageBox.information(self, self.windowTitle(), "请选择一条议程编辑。")
            return
        rows = list(self.desk.board.agenda)
        row = rows[selected[0]] if existing else Session("", meeting=self.filters["meeting"].currentData() or "")
        dialog = SessionEditor(row, self, self.desk.board.meetings)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            rows = list(self.desk.board.agenda)
            if existing:
                index = next((i for i, current in enumerate(rows) if current.uid == row.uid), None)
                if index is None or rows[index] != row:
                    QMessageBox.warning(self, self.windowTitle(), "此议程已被其他操作修改，请重新选择。")
                    dialog.deleteLater()
                    return
                rows[index] = dialog.session()
            else:
                rows.append(dialog.session())
            self.save(rows)
        dialog.deleteLater()

    def remove(self):
        selected = {self.desk.board.agenda[i].uid for i in self.selected()}
        if selected and QMessageBox.question(self, self.windowTitle(), f"删除选中的 {len(selected)} 条议程？") == QMessageBox.StandardButton.Yes:
            self.save([r for r in self.desk.board.agenda if r.uid not in selected])

    def batch(self):
        selected = {self.desk.board.agenda[i].uid for i in self.selected()}
        if not selected:
            QMessageBox.information(self, self.windowTitle(), "请先选中需要修改的议程。")
            return
        fields = {"会议": "meeting", "日期": "date", "会场": "room", "状态": "status"}
        label, ok = QInputDialog.getItem(self, "批量修改", "修改字段", list(fields), 0, False)
        if not ok:
            return
        key = fields[label]
        if key == "status":
            value, ok = QInputDialog.getItem(self, "批量修改", label, SESSION_STATES, 0, False)
        else:
            value, ok = QInputDialog.getText(self, "批量修改", label + ("（YYYY-MM-DD）" if key == "date" else ""))
        if not ok:
            return
        try:
            value = _date(value.strip()) if key == "date" else value.strip()
        except ValueError as exc:
            QMessageBox.warning(self, self.windowTitle(), str(exc))
            return
        self.save([replace(row, **{key: value}) if row.uid in selected else row
                   for row in self.desk.board.agenda])

    def groups(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("会议分组")
        layout = QVBoxLayout(dialog)
        names = QListWidget()
        layout.addWidget(names)
        def refresh():
            names.clear()
            names.addItems(sorted(set(self.desk.board.meetings) | {s.meeting for s in self.desk.board.agenda if s.meeting}))
        def change(mode):
            current = names.currentItem()
            if mode != "add" and current is None:
                return
            old = current.text() if mode != "add" else ""
            value = ""
            if mode != "delete":
                value, ok = QInputDialog.getText(dialog, "会议名称", "名称", text=old)
                value = value.strip()
                if not ok or not value or value == old:
                    return
                if names.findItems(value, Qt.MatchFlag.MatchExactly):
                    QMessageBox.warning(dialog, "会议名称", "此会议名称已经存在。")
                    return
            elif QMessageBox.question(dialog,"移除会议","清除会议分组并保留该会议的议程？") != QMessageBox.StandardButton.Yes:
                return
            meetings = [names.item(i).text() for i in range(names.count()) if names.item(i).text() != old]
            if value:
                meetings.append(value)
            rows = [replace(s, meeting=value) if old and s.meeting == old else s for s in self.desk.board.agenda]
            self.save(rows, meetings)
            refresh()
        refresh()
        row = QHBoxLayout()
        for title, mode in (("添加","add"),("重命名","rename"),("移除","delete")):
            row.addWidget(_button(title, lambda m=mode: change(m)))
        layout.addLayout(row)
        dialog.exec()
        dialog.deleteLater()

    def import_rows(self):
        file, _ = QFileDialog.getOpenFileName(self, "导入会议日程", "", "会议日程 (*.csv *.xlsx)")
        if not file:
            return
        try:
            entries = read_schedule(file)
        except (OSError, ValueError, csv.Error) as exc:
            QMessageBox.warning(self, self.windowTitle(), str(exc))
            return
        dlg = QMessageBox(self)
        dlg.setWindowTitle("导入日程")
        dlg.setText(f"读取到 {len(entries)} 条议程。首条为“{entries[0].title}”。")
        append = dlg.addButton("追加到现有议程", QMessageBox.ButtonRole.AcceptRole)
        overwrite = dlg.addButton("替换现有议程", QMessageBox.ButtonRole.DestructiveRole)
        dlg.addButton(QMessageBox.StandardButton.Cancel)
        dlg.setDefaultButton(append)
        dlg.exec()
        if dlg.clickedButton() is append:
            self.save(self.desk.board.agenda + entries)
        elif dlg.clickedButton() is overwrite:
            self.save(entries)
        dlg.deleteLater()
