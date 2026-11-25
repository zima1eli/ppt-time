"""Local conference schedule, presentation library and audience scenes."""

import json
import re
from dataclasses import asdict, dataclass, field, replace
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from PySide6.QtCore import QIODevice, QObject, QPoint, QRect, QSaveFile, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QFont, QImageReader, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog,
    QAbstractItemView, QFormLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QListWidget, QMessageBox, QPushButton, QSpinBox, QTabWidget, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)


IMAGE_FILTER = "图片 (*.png *.jpg *.jpeg *.webp *.bmp *.gif *.tif *.tiff)"
MEDIA_FILTER = ("会议文件 (*.ppt *.pptx *.pps *.ppsx *.dps *.pdf *.mp4 *.mov *.mkv *.avi "
                "*.m4v *.wmv *.webm);;所有文件 (*)")


@dataclass
class Session:
    title: str
    start: str = ""
    speaker: str = ""
    document: str = ""
    seconds: int = 0
    file_name: str = ""
    meeting: str = ""
    room: str = ""
    date: str = ""
    organization: str = ""
    status: str = ""
    uid: str = field(default_factory=lambda: uuid4().hex, compare=False)

    def public(self):
        result = {"uid": self.uid, "title": self.title, "speaker": self.speaker, "start": self.start,
                  "seconds": self.seconds, "file": Path(self.document).name or self.file_name}
        result.update({key: getattr(self, key) for key in SESSION_FIELDS if getattr(self, key)})
        return result


SESSION_FIELDS = ("meeting", "room", "date", "organization", "status")
SESSION_STATES = ("待开始", "进行中", "已完成")


@dataclass
class Board:
    title: str = ""
    home: str = ""
    clock: bool = True
    auto_home: bool = False
    safe_end: bool = False
    screen: str = "follow"
    images: list[str] = field(default_factory=list)
    files: list[str] = field(default_factory=list)
    agenda: list[Session] = field(default_factory=list)
    meetings: list[str] = field(default_factory=list)

    @classmethod
    def load(cls, path):
        board = cls()
        try:
            values = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, ValueError):
            return board
        if not isinstance(values, dict):
            return board
        for name in ("title", "home", "screen", "clock", "auto_home", "safe_end"):
            value = values.get(name)
            if type(value) is type(getattr(board, name)):
                setattr(board, name, value)
        for name in ("images", "files", "meetings"):
            if isinstance(values.get(name), list):
                setattr(board, name, list(dict.fromkeys(
                    item for item in values[name] if isinstance(item, str) and item
                )))
        if isinstance(values.get("agenda"), list):
            seen = set()
            for row in values["agenda"]:
                if not isinstance(row, dict) or not isinstance(row.get("title"), str):
                    continue
                entry = Session(row["title"])
                for key in ("start", "speaker", "document", "file_name", *SESSION_FIELDS):
                    if isinstance(row.get(key), str):
                        setattr(entry, key, row[key])
                key = row.get("uid")
                if isinstance(key, str) and re.fullmatch(r"[0-9a-f]{32}", key) and key not in seen:
                    entry.uid = key
                seen.add(entry.uid)
                if type(row.get("seconds")) is int:
                    entry.seconds = max(0, min(row["seconds"], 24 * 3600))
                board.agenda.append(entry)
        return board


def _photo(path):
    if not path:
        return QPixmap()
    reader = QImageReader(str(path))
    reader.setAutoTransform(True)
    # Normalize large stills before caching them for projector repainting.
    size = reader.size()
    if size.isValid() and max(size.width(), size.height()) > 8192:
        reader.setScaledSize(size.scaled(8192, 8192, Qt.AspectRatioMode.KeepAspectRatio))
    return QPixmap.fromImage(reader.read())


class SceneCanvas(QWidget):
    dismiss = Signal()
    turn_board_page = Signal(int)
    AGENDA_ROWS = 6

    def __init__(self, audience=False):
        flags = Qt.WindowType.Widget
        if audience:
            flags = (Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint |
                     Qt.WindowType.WindowStaysOnTopHint)
        super().__init__(None, flags)
        self.setWindowTitle("ppt time · 大会画面")
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, audience)
        self._picture = QPixmap()
        self._path = None
        self._title = ""
        self._clock = False
        self._agenda = None
        self._active_row = -1
        self._page = 0
        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self.update)

    def set_scene(self, path, picture, title="", clock=False):
        self._path, self._picture, self._title, self._clock = path, picture, title, clock
        self._agenda = None
        self._sync_clock()
        self.update()

    def set_board(self, board, active_row, page):
        self._agenda, self._title = board.agenda, board.title
        self._active_row, self._page = active_row, page
        self._clock = False
        self._sync_clock()
        self.update()

    def _sync_clock(self):
        if self._clock and self.isVisible():
            if not self._timer.isActive():
                self._timer.start()
        else:
            self._timer.stop()

    def showEvent(self, event):
        super().showEvent(event)
        self._sync_clock()

    def hideEvent(self, event):
        self._timer.stop()
        super().hideEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.dismiss.emit()
            event.accept()
        elif self._agenda is not None and event.key() in (
                Qt.Key.Key_Right, Qt.Key.Key_PageDown, Qt.Key.Key_Left, Qt.Key.Key_PageUp):
            self.turn_board_page.emit(1 if event.key() in (Qt.Key.Key_Right, Qt.Key.Key_PageDown) else -1)
            event.accept()
        else:
            super().keyPressEvent(event)

    def mouseDoubleClickEvent(self, event):
        self.dismiss.emit()
        event.accept()

    def contextMenuEvent(self, event):
        self.dismiss.emit()
        event.accept()

    def paintEvent(self, _):
        p = QPainter(self)
        if self._agenda is not None:
            self._paint_agenda(p)
            p.end()
            return
        p.fillRect(self.rect(), QColor("#12202c") if self._picture.isNull() else Qt.GlobalColor.black)
        p.setRenderHints(QPainter.RenderHint.Antialiasing | QPainter.RenderHint.SmoothPixmapTransform)
        if not self._picture.isNull():
            size = self._picture.size().scaled(self.size(), Qt.AspectRatioMode.KeepAspectRatio)
            x, y = (self.width() - size.width()) // 2, (self.height() - size.height()) // 2
            p.drawPixmap(QRect(x, y, size.width(), size.height()), self._picture)
        elif self._title:
            p.setPen(Qt.GlobalColor.white)
            font = QFont(self.font())
            font.setPixelSize(max(16, int(self.height() * .075)))
            font.setBold(True)
            p.setFont(font)
            margin = max(16, self.width() // 16)
            p.drawText(self.rect().adjusted(margin, margin, -margin, -margin),
                       Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap, self._title)
        if self._clock:
            now = datetime.now()
            line = now.strftime("%Y-%m-%d   %H:%M:%S")
            font = QFont(self.font())
            font.setPixelSize(max(12, round(self.height() * .031)))
            p.setFont(font)
            metrics = p.fontMetrics()
            pad = max(8, metrics.height() // 2)
            width, height = metrics.horizontalAdvance(line) + 2 * pad, metrics.height() + pad
            box = QRect(self.width() - width - pad, self.height() - height - pad, width, height)
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(10, 15, 22, 180))
            p.drawRoundedRect(box, pad / 2, pad / 2)
            p.setPen(Qt.GlobalColor.white)
            p.drawText(box, Qt.AlignmentFlag.AlignCenter, line)
        p.end()

    def _paint_agenda(self, p):
        w, h = self.width(), self.height()
        p.fillRect(self.rect(), QColor("#12202c"))
        margin = max(7, min(w // 16, h // 14))
        width = max(1, w - margin * 2)
        short = min(w, h)

        def line(rect, text, size, color, bold=False, align=Qt.AlignmentFlag.AlignLeft):
            font = QFont(self.font())
            font.setPixelSize(max(7, size))
            font.setBold(bold)
            p.setFont(font)
            p.setPen(QColor(color))
            fit = p.fontMetrics().elidedText(text, Qt.TextElideMode.ElideRight, max(0, rect.width()))
            p.drawText(rect, Qt.AlignmentFlag.AlignVCenter | align, fit)

        header_h = max(22, round(h * .125))
        title = self._title + " · 会议议程" if self._title else "会议议程"
        line(QRect(margin, margin // 2, width, header_h), title, int(short * .054), "#ffffff", True)
        label_y = margin // 2 + header_h
        label_h, footer_h = max(14, h // 16), max(14, h // 14)
        rows_y = label_y + label_h
        row_h = max(1, (h - margin // 2 - footer_h - rows_y) // self.AGENDA_ROWS)
        gap = max(3, width // 60)
        speaker_w = width // 5
        time_w = width // 6
        title_w = width - time_w - speaker_w - 4*gap
        columns = ((margin+gap, time_w-gap), (margin+time_w+gap, title_w+gap),
                   (margin+width-speaker_w, speaker_w-gap))
        for (x, length), text in zip(columns, ("时间", "议题", "讲者")):
            line(QRect(x, label_y, length, label_h), text, int(short * .025), "#a5c6d1")

        start = self._page * self.AGENDA_ROWS
        for row_no, entry in enumerate(self._agenda[start:start+self.AGENDA_ROWS], start=start):
            y = rows_y + (row_no-start)*row_h
            row = QRect(margin, y, width, row_h)
            active = row_no == self._active_row
            if active:
                p.fillRect(row.adjusted(0, 1, 0, -1), QColor("#21495a"))
                p.fillRect(QRect(margin, y+1, max(3, w//230), max(1, row_h-2)), QColor("#72e0cd"))
            else:
                p.fillRect(QRect(margin, y+row_h-1, width, 1), QColor("#2c414d"))
            for col, ((x, length), text) in enumerate(zip(columns, (entry.start, entry.title, entry.speaker))):
                line(QRect(x, y, length, row_h), text, int(short * (.032 if col == 1 else .027)),
                     "#ffffff" if col == 1 or active else "#c3d6df", active or col == 1)
        if not self._agenda:
            line(QRect(margin+gap, rows_y, width-2*gap, row_h), "暂无议程", int(short*.032), "#c3d6df")
        pages = max(1, (len(self._agenda) + self.AGENDA_ROWS - 1) // self.AGENDA_ROWS)
        if pages > 1:
            line(QRect(margin, h-footer_h-margin//2, width, footer_h), f"{self._page+1} / {pages}",
                 int(short*.026), "#a5c6d1", align=Qt.AlignmentFlag.AlignRight)


class ConferenceDesk(QObject):
    changed = Signal()
    open_file_requested = Signal(str)
    duration_requested = Signal(int)

    def __init__(self, state_path, parent=None):
        super().__init__(parent)
        self.path = Path(state_path)
        self.board = Board.load(self.path)
        self.scene = "none"
        self.index = -1
        self.active_session = -1
        self.board_page = 0
        self.last_error = ""
        self._show = False
        self._terminal = False
        self._terminal_home = False
        self._follow = ""
        self._windows = {}
        self._cached_path = None
        self._picture = QPixmap()
        self.panel = None
        app = QApplication.instance()
        app.screenAdded.connect(self._screens_changed)
        app.screenRemoved.connect(self._screens_changed)
        if self.board.auto_home:
            self.command("home")

    def _fail(self, message):
        self.last_error = message
        return False

    def update_board(self, **updates):
        new = replace(self.board, **updates)
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            return self._fail(str(exc))
        output = QSaveFile(str(self.path))
        raw = json.dumps(asdict(new), ensure_ascii=False, indent=2).encode("utf-8")
        if (not output.open(QIODevice.OpenModeFlag.WriteOnly) or
                output.write(raw) != len(raw) or not output.commit()):
            return self._fail("无法保存会务设置：" + output.errorString())
        previous = self.board
        self.board = new
        if previous.agenda != new.agenda and self.active_session >= 0:
            active = previous.agenda[self.active_session]
            self.active_session = next((i for i, s in enumerate(new.agenda) if s.uid == active.uid), -1)
        if self.scene == "interstitial" and previous.images != new.images:
            selected = previous.images[self.index]
            if selected in new.images:
                self.index = new.images.index(selected)
            else:
                self.scene, self.index = "none", -1
        if new.auto_home and not previous.auto_home and not self._show and self.scene == "none":
            self.scene = "home"
        self.last_error = ""
        self.board_page = min(self.board_page, self._last_board_page())
        self._display()
        self.changed.emit()
        return True

    def add_files(self, paths):
        current = list(self.board.files)
        for path in paths:
            name = str(Path(path).absolute())
            if name not in current and Path(name).is_file():
                current.append(name)
        return self.update_board(files=current) if current != self.board.files else True

    def status(self):
        return {
            "scene": self.scene, "title": self.board.title, "index": self.index,
            "active_session": self.active_session,
            "board_page": self.board_page + 1, "board_pages": self._last_board_page() + 1,
            "agenda": [s.public() for s in self.board.agenda],
            "meetings": self.board.meetings.copy(),
            "interstitials": [Path(path).name for path in self.board.images],
            "files": [Path(path).name for path in self.board.files],
        }

    def import_agenda(self, payload):
        """Replace the agenda with public meeting data; file access stays local."""
        if not isinstance(payload, list) or len(payload) > 2000:
            return self._fail("议程数据格式不正确。")
        entries = []
        local = {entry.uid: entry for entry in self.board.agenda}
        legacy = {}
        for entry in self.board.agenda:
            identity = (entry.title, entry.start, entry.meeting, entry.room, entry.date,
                        Path(entry.document).name or entry.file_name)
            legacy[identity] = entry if identity not in legacy else None
        seen = set()
        for item in payload:
            if not isinstance(item, dict):
                return self._fail("议程条目格式不正确。")
            fields = [item.get(key, "") for key in ("title", "start", "speaker", "file")]
            seconds = item.get("seconds", 0)
            if (any(not isinstance(field, str) or len(field) > 1024 for field in fields)
                    or not fields[0].strip() or type(seconds) is not int
                    or not 0 <= seconds <= 24 * 3600):
                return self._fail("议程的标题或时长无效。")
            title, start, speaker, filename = fields
            if filename and (filename in (".", "..") or any(c in filename for c in "/\\:\0")):
                return self._fail("同步附件必须使用文件名。")
            extras = {key: item.get(key, "") for key in SESSION_FIELDS}
            if any(not isinstance(field, str) or len(field) > 1024 for field in extras.values()):
                return self._fail("会议信息格式不正确。")
            if extras["status"] and extras["status"] not in SESSION_STATES:
                return self._fail("议程状态不正确。")
            if extras["date"]:
                try:
                    extras["date"] = datetime.strptime(extras["date"], "%Y-%m-%d").date().isoformat()
                except ValueError:
                    return self._fail("会议日期不正确。")
            entry = Session(title, start, speaker, seconds=seconds, file_name=filename, **extras)
            uid = item.get("uid")
            if uid is not None:
                if not isinstance(uid, str) or not re.fullmatch(r"[0-9a-f]{32}", uid) or uid in seen:
                    return self._fail("议程编号无效或重复。")
                entry.uid = uid
            else:
                old_match = legacy.get((title, start, entry.meeting, entry.room, entry.date, filename))
                if old_match and old_match.uid not in seen:
                    entry.uid = old_match.uid
            seen.add(entry.uid)
            old = local.get(entry.uid)
            if old and old.document and Path(old.document).name == filename:
                entry.document = old.document
                entry.file_name = old.file_name
            entries.append(entry)
        if self.board.agenda == entries and list(local) == [entry.uid for entry in entries]:
            return True
        meetings = list(dict.fromkeys([*self.board.meetings, *(s.meeting for s in entries if s.meeting)]))
        return self.update_board(agenda=entries, meetings=meetings)

    def set_snapshot(self, snapshot):
        show = bool(snapshot is not None and snapshot.active)
        terminal = show and bool(getattr(snapshot, "ended", False))
        previous_screen = self._follow
        if show:
            screen = self._observed_screen(snapshot)
            if screen:
                self._follow = screen.name()
        entered, exited = show and not self._show, self._show and not show
        crossed_end = terminal and not self._terminal
        left_end = self._terminal and show and not terminal
        self._show, self._terminal = show, terminal
        if crossed_end and self.board.safe_end and self.scene == "none":
            self.command("home")
            self._terminal_home = True
        elif entered or (left_end and self._terminal_home):
            self.command("hide")
        elif exited and self.board.auto_home and self.scene == "none":
            self.command("home")
        elif previous_screen != self._follow and self.board.screen == "follow":
            self._display()
        if not terminal:
            self._terminal_home = False
        if exited and 0 <= self.active_session < len(self.board.agenda):
            rows = self.board.agenda.copy()
            entry = rows[self.active_session]
            if entry.status == "进行中":
                rows[self.active_session] = replace(entry, status="已完成")
                self.update_board(agenda=rows)

    @staticmethod
    def _observed_screen(snapshot):
        name = (getattr(snapshot, "screen_name", "") or "").casefold().rsplit("\\", 1)[-1]
        for screen in QApplication.screens():
            if name and screen.name().casefold().rsplit("\\", 1)[-1] == name:
                return screen
        point = getattr(snapshot, "screen_point", None)
        return QApplication.screenAt(QPoint(*point)) if point is not None else None

    def _targets(self):
        screens = QApplication.screens()
        if self.board.screen == "all":
            return screens
        primary = QApplication.primaryScreen()
        selected = primary if self.board.screen == "primary" else None
        name = self._follow if self.board.screen == "follow" else self.board.screen
        selected = next((s for s in screens if s.name() == name), selected)
        if not selected:
            selected = next((s for s in screens if s != primary), primary)
        return [selected] if selected else []

    def media_screen(self):
        """A PDF or video has one audience window even when clocks appear on both screens."""
        targets = self._targets()
        primary = QApplication.primaryScreen()
        return next((screen for screen in targets if screen != primary), targets[0] if targets else primary)

    def _screens_changed(self, *_):
        self._display()
        self.changed.emit()

    def _scene_values(self, scene, index):
        if scene == "home":
            return self.board.home, self.board.title, self.board.clock
        if scene == "board":
            return "", "", False
        return self.board.images[index], "", False

    def _last_board_page(self):
        return max(0, (len(self.board.agenda) - 1) // SceneCanvas.AGENDA_ROWS)

    def _display(self):
        targets = self._targets() if self.scene != "none" else []
        for screen in list(self._windows):
            if screen not in targets:
                window = self._windows.pop(screen)
                window.hide()
                window.deleteLater()
        if not targets:
            return
        path, title, clock = self._scene_values(self.scene, self.index)
        if self._cached_path != path:
            self._cached_path, self._picture = path, _photo(path)
        for screen in targets:
            window = self._windows.get(screen)
            if window is None:
                window = self._windows[screen] = SceneCanvas(audience=True)
                window.dismiss.connect(lambda: self.command("hide"))
                window.turn_board_page.connect(lambda delta: self.command("board_next" if delta > 0 else "board_previous"))
                window.setScreen(screen)
                screen.geometryChanged.connect(window.setGeometry)
            window.setGeometry(screen.geometry())
            if self.scene == "board":
                window.set_board(self.board, self.active_session, self.board_page)
            else:
                window.set_scene(path, self._picture, title, clock)
            if not window.isVisible():
                window.show()

    def command(self, action, value=None):
        self.last_error = ""
        board = self.board
        if action in ("file", "agenda"):
            rows = board.files if action == "file" else board.agenda
            if type(value) is not int or not 0 <= value < len(rows):
                return self._fail("该条目不存在。")
            session = rows[value] if action == "agenda" else None
            name = session.document if session else rows[value]
            if session and not name and session.file_name:
                matches = [p for p in board.files if Path(p).name.casefold() == session.file_name.casefold()]
                if len(matches) != 1:
                    return self._fail("请在议程中选择本会场的附件：" + session.file_name)
                name = matches[0]
            if name and not Path(name).is_file():
                return self._fail("会议文件已移动或删除。")
            if session:
                sessions = board.agenda.copy()
                if 0 <= self.active_session < len(sessions) and self.active_session != value:
                    prior = sessions[self.active_session]
                    if prior.status == "进行中":
                        sessions[self.active_session] = replace(prior, status="已完成")
                sessions[value] = replace(session, status="进行中")
                if sessions != board.agenda and not self.update_board(agenda=sessions):
                    return False
                self.active_session = value
                if session.seconds:
                    self.duration_requested.emit(session.seconds)
            if name:
                self.command("hide")
                self.open_file_requested.emit(name)
            elif session and self.scene == "board":
                self.board_page = value // SceneCanvas.AGENDA_ROWS
                self._display()
            self.changed.emit()
            return True
        if action in ("next", "previous", "interstitial"):
            if not board.images:
                return self._fail("请先添加串场图片。")
            if action == "next":
                value = (self.index + 1) % len(board.images)
            elif action == "previous":
                value = (self.index - 1) % len(board.images) if self.index >= 0 else len(board.images) - 1
            if type(value) is not int or not 0 <= value < len(board.images):
                return self._fail("该图片不存在。")
            path = board.images[value]
            image = _photo(path)
            if image.isNull():
                return self._fail("串场图片无法读取：" + Path(path).name)
            self._cached_path, self._picture = path, image
            self.scene, self.index = "interstitial", value
        elif action == "home":
            self.scene = "home"
        elif action == "board":
            self.scene = "board"
            self.board_page = max(0, self.active_session) // SceneCanvas.AGENDA_ROWS
        elif action in ("board_previous", "board_next"):
            if self.scene != "board":
                return self._fail("请先投放议程看板。")
            self.board_page = max(0, min(self._last_board_page(), self.board_page + (1 if action == "board_next" else -1)))
        elif action == "hide":
            if self.scene == "none":
                return True
            self.scene = "none"
        else:
            return self._fail("未知的会务操作。")
        self._terminal_home = False
        self._display()
        self.changed.emit()
        return True

    def open_panel(self):
        if self.panel is None:
            self.panel = BoardPanel(self)
        self.panel.show()
        self.panel.raise_()
        self.panel.activateWindow()

    def close(self):
        app = QApplication.instance()
        app.screenAdded.disconnect(self._screens_changed)
        app.screenRemoved.disconnect(self._screens_changed)
        self.command("hide")
        if self.panel is not None:
            self.panel.close()
            self.panel.deleteLater()
            self.panel = None


def _button(text, callback):
    button = QPushButton(text)
    button.clicked.connect(lambda _checked=False: callback())
    return button


class SessionEditor(QDialog):
    def __init__(self, session, parent, meetings=()):
        super().__init__(parent)
        self.original = session
        self.setWindowTitle("议程信息")
        form = QFormLayout(self)
        self.title_field = QLineEdit(session.title)
        self.start_field = QLineEdit(session.start)
        self.start_field.setPlaceholderText("例如 09:30")
        self.speaker = QLineEdit(session.speaker)
        self.meeting = QComboBox()
        self.meeting.setEditable(True)
        self.meeting.addItems(meetings)
        self.meeting.setCurrentText(session.meeting)
        self.room = QLineEdit(session.room)
        self.date = QLineEdit(session.date)
        self.date.setPlaceholderText("YYYY-MM-DD")
        self.organization = QLineEdit(session.organization)
        self.state = QComboBox()
        self.state.addItems(SESSION_STATES)
        self.state.setCurrentText(session.status or SESSION_STATES[0])
        self.duration = QSpinBox()
        self.duration.setRange(0, 24 * 3600)
        self.duration.setSpecialValueText("沿用当前计时")
        self.duration.setSuffix(" 秒")
        self.duration.setValue(session.seconds)
        self.file = QLineEdit(session.document)
        self.previous_filename = session.file_name
        if self.previous_filename:
            self.file.setPlaceholderText("同步附件：" + self.previous_filename)
        file_row = QHBoxLayout()
        file_row.addWidget(self.file)
        file_row.addWidget(_button("浏览…", self.choose_file))
        form.addRow("议题", self.title_field)
        form.addRow("会议", self.meeting)
        form.addRow("日期", self.date)
        form.addRow("会场", self.room)
        form.addRow("开始时间", self.start_field)
        form.addRow("讲者", self.speaker)
        form.addRow("单位", self.organization)
        form.addRow("状态", self.state)
        form.addRow("发言时长", self.duration)
        form.addRow("会议文件", file_row)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                                   QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.validate)
        buttons.rejected.connect(self.reject)
        form.addRow(buttons)

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择会议文件", self.file.text(), MEDIA_FILTER)
        if path:
            self.file.setText(path)

    def validate(self):
        if not self.title_field.text().strip():
            self.title_field.setFocus()
            return
        start = self.start_field.text().strip()
        date = self.date.text().strip()
        if date:
            try:
                self.date.setText(datetime.strptime(date, "%Y-%m-%d").date().isoformat())
            except ValueError:
                QMessageBox.information(self, self.windowTitle(), "请输入有效日期，例如 2026-11-24。")
                return
        if start:
            try:
                self.start_field.setText(datetime.strptime(start, "%H:%M").strftime("%H:%M"))
            except ValueError:
                QMessageBox.information(self, self.windowTitle(), "请用 24 小时制输入时间，例如 09:30。")
                return
        if self.file.text().strip() and not Path(self.file.text().strip()).is_file():
            QMessageBox.information(self, self.windowTitle(), "请选择已有文件，或清空文件路径。")
            return
        self.accept()

    def session(self):
        document = self.file.text().strip()
        return replace(self.original, title=self.title_field.text().strip(), start=self.start_field.text().strip(),
                       speaker=self.speaker.text().strip(), document=document, seconds=self.duration.value(),
                       file_name="" if document else self.previous_filename, meeting=self.meeting.currentText().strip(),
                       date=self.date.text().strip(), room=self.room.text().strip(),
                       organization=self.organization.text().strip(), status=self.state.currentText())


class BoardPanel(QDialog):
    def __init__(self, desk):
        super().__init__()
        self.desk = desk
        self.setWindowTitle("ppt time · 会务中控")
        self.resize(750, 500)
        outer = QVBoxLayout(self)
        self.tabs = QTabWidget()
        outer.addWidget(self.tabs)
        self._home_page()
        self._images_page()
        self._agenda_page()
        self._files_page()
        self.status_line = QLabel()
        footer = QHBoxLayout()
        footer.addWidget(self.status_line, 1)
        footer.addWidget(_button("收起大屏画面", lambda: self._command("hide")))
        outer.addLayout(footer)
        self.desk.changed.connect(self.reload)
        self.reload()

    def _check(self, result):
        if not result:
            QMessageBox.warning(self, self.windowTitle(), self.desk.last_error)
        return result

    def _command(self, name, index=None):
        self._check(self.desk.command(name, index))

    def _save(self, **changes):
        return self._check(self.desk.update_board(**changes))

    def _buttons(self, layout, specs):
        row = QHBoxLayout()
        for text, callback in specs:
            row.addWidget(_button(text, callback))
        row.addStretch()
        layout.addLayout(row)

    def _home_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        form = QFormLayout()
        self.event_title = QLineEdit()
        self.event_title.setPlaceholderText("未使用背景图片时在大屏居中显示")
        self.event_title.editingFinished.connect(self._save_title)
        form.addRow("会议名称", self.event_title)
        bg = QHBoxLayout()
        self.image_label = QLabel()
        bg.addWidget(self.image_label, 1)
        bg.addWidget(_button("选择图片…", self._home_image))
        bg.addWidget(_button("清除", lambda: self._save(home="")))
        form.addRow("大会主画面", bg)
        self.screen = QComboBox()
        self.screen.activated.connect(lambda i: self._save(screen=self.screen.itemData(i)))
        form.addRow("输出屏幕", self.screen)
        self.checks = {}
        for key, label in (("clock", "在主画面显示日期和时间"),
                           ("auto_home", "放映结束后自动显示主画面"),
                           ("safe_end", "收到末页结束信号时提前切回主画面")):
            box = QCheckBox(label)
            box.clicked.connect(lambda checked=False, k=key: self._save(**{k: checked}))
            self.checks[key] = box
            form.addRow("", box)
        layout.addLayout(form)
        self._buttons(layout, [("预览", lambda: self._preview("home")),
                               ("投放主画面", lambda: self._command("home"))])
        layout.addStretch()
        self.tabs.addTab(page, "大会主画面")

    def _save_title(self):
        title = self.event_title.text().strip()
        if title != self.desk.board.title:
            self._save(title=title)

    def _home_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择大会主画面", self.desk.board.home, IMAGE_FILTER)
        if path:
            if _photo(path).isNull():
                QMessageBox.warning(self, self.windowTitle(), "无法读取所选图片。")
            else:
                self._save(home=path)

    def _images_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        self.images = QListWidget()
        self.images.itemDoubleClicked.connect(lambda _: self._command("interstitial", self.images.currentRow()))
        layout.addWidget(self.images)
        self._buttons(layout, [("添加图片…", self._add_images),
                               ("移除", lambda: self._remove("images")),
                               ("上移", lambda: self._move("images", -1)),
                               ("下移", lambda: self._move("images", 1)),
                               ("预览", lambda: self._preview("interstitial"))])
        self._buttons(layout, [("投放选中图片", lambda: self._command("interstitial", self.images.currentRow())),
                               ("上一张", lambda: self._command("previous")),
                               ("下一张", lambda: self._command("next"))])
        self.tabs.addTab(page, "串场图片")

    def _add_images(self):
        items, _ = QFileDialog.getOpenFileNames(self, "添加串场图片", "", IMAGE_FILTER)
        if items:
            self._save(images=list(dict.fromkeys([*self.desk.board.images, *items])))

    def _agenda_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        self.agenda = QTableWidget(0, 5)
        self.agenda.setHorizontalHeaderLabels(["开始", "议题", "讲者", "时长", "文件"])
        self.agenda.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.agenda.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.agenda.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.agenda.verticalHeader().setVisible(False)
        self.agenda.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.agenda.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.agenda.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.agenda.cellDoubleClicked.connect(lambda _r, _c: self._edit_agenda(True))
        layout.addWidget(self.agenda)
        self._buttons(layout, [("添加议程…", lambda: self._edit_agenda(False)),
                               ("编辑…", lambda: self._edit_agenda(True)),
                               ("移除", lambda: self._remove("agenda")),
                               ("上移", lambda: self._move("agenda", -1)),
                               ("下移", lambda: self._move("agenda", 1)),
                               ("开始此议程", lambda: self._command("agenda", self.agenda.currentRow()))])
        self._buttons(layout, [("日程管理与导入…", self._manage_agenda),
                               ("投放议程看板", lambda: self._command("board")),
                               ("看板上一页", lambda: self._command("board_previous")),
                               ("看板下一页", lambda: self._command("board_next"))])
        self.tabs.addTab(page, "会议议程")

    def _manage_agenda(self):
        from agenda_manager import AgendaManager
        dialog = AgendaManager(self.desk, self)
        dialog.exec()
        dialog.deleteLater()

    def _edit_agenda(self, edit):
        rows = list(self.desk.board.agenda)
        index = self.agenda.currentRow()
        if edit and index < 0:
            return
        old = rows[index] if edit else Session("")
        dialog = SessionEditor(old, self, self.desk.board.meetings)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            entry = dialog.session()
            rows = list(self.desk.board.agenda)
            if edit:
                index = next((i for i, row in enumerate(rows) if row.uid == old.uid), None)
                if index is None or rows[index] != old:
                    QMessageBox.warning(self, self.windowTitle(), "此议程已被更新，请重新选择。")
                    dialog.deleteLater()
                    return
                rows[index] = entry
            else:
                rows.append(entry)
                index = len(rows) - 1
            files = list(self.desk.board.files)
            if entry.document and entry.document not in files:
                files.append(entry.document)
            meetings = list(dict.fromkeys([*self.desk.board.meetings, *([entry.meeting] if entry.meeting else [])]))
            if self._save(agenda=rows, files=files, meetings=meetings):
                self.agenda.selectRow(index)
        dialog.deleteLater()

    def _files_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        self.files = QListWidget()
        self.files.itemDoubleClicked.connect(lambda _: self._command("file", self.files.currentRow()))
        layout.addWidget(self.files)
        self._buttons(layout, [("添加会议文件…", self._add_files),
                               ("移除", lambda: self._remove("files")),
                               ("上移", lambda: self._move("files", -1)),
                               ("下移", lambda: self._move("files", 1)),
                               ("打开文件", lambda: self._command("file", self.files.currentRow()))])
        self.tabs.addTab(page, "文件看板")

    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "添加会议文件", "", MEDIA_FILTER)
        if paths:
            self._check(self.desk.add_files(paths))

    def _remove(self, key):
        index = getattr(self, key).currentRow()
        if index >= 0:
            rows = list(getattr(self.desk.board, key))
            rows.pop(index)
            self._save(**{key: rows})

    def _move(self, key, offset):
        widget = getattr(self, key)
        index = widget.currentRow()
        rows = list(getattr(self.desk.board, key))
        target = index + offset
        if index >= 0 and 0 <= target < len(rows):
            rows[index], rows[target] = rows[target], rows[index]
            if self._save(**{key: rows}):
                if key == "agenda":
                    widget.selectRow(target)
                else:
                    widget.setCurrentRow(target)

    def _preview(self, scene):
        index = self.images.currentRow() if scene == "interstitial" else -1
        if scene == "interstitial" and index < 0:
            return
        path, title, clock = self.desk._scene_values(scene, index)
        picture = _photo(path)
        if path and picture.isNull():
            QMessageBox.warning(self, self.windowTitle(), "无法读取图片：" + Path(path).name)
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("会场画面预览")
        layout = QVBoxLayout(dialog)
        canvas = SceneCanvas()
        canvas.set_scene(path, picture, title, clock)
        canvas.dismiss.connect(dialog.reject)
        layout.addWidget(canvas)
        dialog.resize(720, 420)
        dialog.exec()
        dialog.deleteLater()

    def reload(self):
        board = self.desk.board
        if not self.event_title.hasFocus():
            self.event_title.setText(board.title)
        self.image_label.setText(Path(board.home).name if board.home else "纯色背景")
        self.image_label.setToolTip(board.home)
        self.screen.clear()
        for key, label in (("follow", "跟随放映屏（候场优先副屏）"), ("primary", "主屏幕"),
                           ("all", "全部屏幕")):
            self.screen.addItem(label, key)
        for screen in QApplication.screens():
            self.screen.addItem(screen.name() or "本机显示器", screen.name())
        index = self.screen.findData(board.screen)
        if index < 0:
            self.screen.addItem(board.screen + "（未连接）", board.screen)
            index = self.screen.count() - 1
        self.screen.setCurrentIndex(index)
        for key, box in self.checks.items():
            box.setChecked(getattr(board, key))
        for key in ("images", "files"):
            widget = getattr(self, key)
            previous_row = widget.currentRow()
            widget.clear()
            for path in getattr(board, key):
                widget.addItem(Path(path).name)
                widget.item(widget.count() - 1).setToolTip(path)
            if widget.count():
                widget.setCurrentRow(min(max(0, previous_row), widget.count() - 1))
        previous = self.agenda.currentRow()
        self.agenda.setRowCount(len(board.agenda))
        for i, row in enumerate(board.agenda):
            minutes, seconds = divmod(row.seconds, 60)
            for j, field in enumerate((row.start, row.title, row.speaker,
                                       f"{minutes}:{seconds:02d}" if row.seconds else "—",
                                       Path(row.document).name or row.file_name)):
                self.agenda.setItem(i, j, QTableWidgetItem(field))
        if board.agenda:
            self.agenda.selectRow(min(max(0, previous), len(board.agenda)-1))
        status = {"home": "正在投放大会主画面", "none": "会务画面已收起",
                  "board": f"正在投放会议议程 · {self.desk.board_page+1}/{self.desk._last_board_page()+1}",
                  "interstitial": "正在投放串场图片"}.get(self.desk.scene, "")
        if self.desk.active_session >= 0:
            status += "  ·  " + board.agenda[self.desk.active_session].title
        self.status_line.setText(status)
