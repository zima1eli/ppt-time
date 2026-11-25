"""Presentation clocks, per-slide time budgets and audience timeout messages."""

from __future__ import annotations

import math
import ntpath
import time

from PySide6.QtCore import QObject, Qt, QTimer, Signal
from PySide6.QtGui import QFontMetrics
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox,
    QFileDialog, QFormLayout, QHBoxLayout, QHeaderView, QLabel,
    QMessageBox, QPlainTextEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)


MAX_SECONDS = 99 * 3600 + 59 * 60 + 59
DEFAULT_MESSAGE = "时间到\nTime’s Up"


def _duration(text):
    """Accept seconds, MM:SS or HH:MM:SS."""
    parts = str(text).strip().split(":")
    if not 1 <= len(parts) <= 3 or any(not p.isascii() or not p.isdecimal() or len(p) > 6 for p in parts):
        raise ValueError("时长可以填写秒数、分:秒或时:分:秒。")
    numbers = [int(part) for part in parts]
    if any(value >= 60 for value in numbers[1:]):
        raise ValueError("时长的进位应为 60。")
    seconds = 0
    for value in numbers:
        seconds = 60 * seconds + value
    if not 1 <= seconds <= MAX_SECONDS:
        raise ValueError("时长须在 1 秒到 99 小时 59 分 59 秒之间。")
    return seconds


def _clock(seconds):
    minutes, seconds = divmod(seconds, 60)
    if minutes < 60:
        return f"{minutes:02d}:{seconds:02d}"
    hours, minutes = divmod(minutes, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def _source_key(value):
    return ntpath.normpath(value).casefold() if value else ""


def _rule(document, page, seconds):
    document = str(document or "").strip()
    if document == "*":
        document = ""
    page = str(page).strip()
    if page in ("*", "整份", "整份文稿"):
        page = "0"
    if not page.isascii() or not page.isdecimal() or len(page) > 6:
        raise ValueError("页码请填写正整数；整份文稿可填 *。")
    return {"document": document, "page": int(page), "seconds": _duration(seconds)}


def timing_options(data=None):
    data = data if isinstance(data, dict) else {}
    rules = []
    seen = set()
    rows = data.get("rules")
    for row in rows if isinstance(rows, list) else []:
        if not isinstance(row, dict):
            continue
        try:
            entry = _rule(row.get("document", ""), row.get("page", 0), row.get("seconds", ""))
        except ValueError:
            continue
        key = _source_key(entry["document"]), entry["page"]
        if key not in seen:
            seen.add(key)
            rules.append(entry)
    mode = data.get("end_mode")
    target = data.get("end_screen")
    message = data.get("message")
    return {
        "pages_only": data.get("pages_only") is True,
        "overtime": data.get("overtime") is True,
        "end_mode": mode if mode in ("notice", "black") else "off",
        "end_screen": target if target in ("all", "primary", "slideshow") else "display",
        "allow_close": data.get("allow_close") is not False,
        "message": message.strip()[:500] if isinstance(message, str) and message.strip() else DEFAULT_MESSAGE,
        "rules": rules,
    }


class TimedStages(QObject):
    """Monotonic stage clock; negative tick values denote elapsed overtime."""

    tick = Signal(str, int)
    stage_changed = Signal(int, str)
    alert_triggered = Signal(str)
    loop_restarted = Signal()
    state_changed = Signal(bool)
    expired = Signal()

    def __init__(self, config, play_sound):
        super().__init__()
        self.config = config
        self._play_sound = play_sound
        self._timer = QTimer(self)
        self._timer.setInterval(200)
        self._timer.timeout.connect(self._tick)
        self.running = False
        self.paused = True
        self.timing_enabled = True
        self.overtime_enabled = False
        self._overtime = False
        self._stage_idx = 0
        self._duration = 0
        self._override_seconds = None
        self._count_up = False
        self._remaining_float = 0.0
        self._target_time = 0.0
        self._last_displayed = -1
        self._zero_triggered = False

    @property
    def can_resume(self):
        return self.timing_enabled and (self._remaining_float > 0 or self._overtime)

    @property
    def at_end(self):
        return self._zero_triggered and self.paused and not self._overtime

    def set_timing_enabled(self, enabled):
        if not enabled and not self.paused:
            self._remaining_float = self._target_time - time.monotonic()
            if not self._overtime:
                self._remaining_float = max(0.001, self._remaining_float)
            self._stop_and_wait()
        self.timing_enabled = enabled

    def set_overtime(self, enabled):
        self.overtime_enabled = enabled
        if self._overtime and not enabled:
            self._overtime = False
            self._remaining_float = 0.0
            self._stop_and_wait()
            self.tick.emit(self.config.stages[self._stage_idx].label, 0)

    def toggle_pause(self):
        if not self.can_resume:
            return
        if self.paused:
            self.running = True
            self.paused = False
            self._target_time = time.monotonic() + self._remaining_float
            self._timer.start()
            self.state_changed.emit(True)
        else:
            self._tick()
            if not self.paused:
                self._stop_and_wait()

    def _stop_and_wait(self):
        self._timer.stop()
        self.paused = True
        self.state_changed.emit(False)

    def _load_stage(self, seconds=None):
        if not self.config.stages:
            return
        stage = self.config.stages[self._stage_idx]
        self._override_seconds = seconds
        self._count_up = stage.count_up if seconds is None else False
        self._duration = stage.seconds if seconds is None else seconds
        self._remaining_float = float(self._duration)
        self._last_displayed = self._duration
        self._zero_triggered = self._overtime = False
        self.stage_changed.emit(self._stage_idx, stage.label)
        self.tick.emit(stage.label, 0 if self._count_up else self._duration)

    def set_duration(self, seconds, start=None):
        """Set a temporary countdown. Preserve its running state when start is None."""
        duration = _duration(seconds)
        resume = not self.paused if start is None else bool(start)
        if not self.config.stages:
            return
        self._stop_and_wait()
        self._load_stage(duration)
        if resume and self.timing_enabled:
            self.toggle_pause()

    def restart_stage(self):
        self._load_stage(self._override_seconds)
        self._stop_and_wait()

    def prev_stage(self):
        if self.config.stages:
            self._stage_idx = (self._stage_idx - 1) % len(self.config.stages)
            self._load_stage()
            self._stop_and_wait()

    def next_stage(self):
        if self.config.stages:
            self._advance_idx()
            self._load_stage()
            self._stop_and_wait()

    def stop(self):
        self._stage_idx = 0
        self._load_stage()
        self._stop_and_wait()
        self.running = False

    def _advance_idx(self):
        self._stage_idx += 1
        if self._stage_idx >= len(self.config.stages):
            self._stage_idx = 0
            self.loop_restarted.emit()

    def _tick(self):
        if self.paused or not self.config.stages:
            return
        now = time.monotonic()
        self._remaining_float = self._target_time - now
        balance = math.ceil(self._remaining_float)
        remaining = max(0, balance)
        stage = self.config.stages[self._stage_idx]
        if self._overtime:
            if balance != self._last_displayed:
                self._last_displayed = balance
                self.tick.emit(stage.label, balance)
            return
        if remaining != self._last_displayed:
            previous = self._last_displayed
            self._last_displayed = remaining
            self.tick.emit(stage.label, self._duration - remaining if self._count_up else remaining)
            self._warn(remaining, previous)
        if self._remaining_float > 0 or self._zero_triggered:
            return
        self._zero_triggered = True
        if self.config.global_sound:
            self._play_sound(self.config.end_sound_path, 2000)
        if self.config.auto_advance:
            self._advance_idx()
            self._load_stage()
            self._target_time = now + self._remaining_float
        else:
            if self.overtime_enabled:
                self._overtime = True
                if balance < 0:
                    self._last_displayed = balance
                    self.tick.emit(stage.label, balance)
            else:
                self._remaining_float = 0.0
                self._stop_and_wait()
            self.expired.emit()

    def _warn(self, remaining, previous):
        if remaining <= 0:
            return
        if self.config.global_sound and self.config.countdown_10s_sound:
            try:
                beep = max(1, int(self.config.countdown_beep_seconds))
            except (ValueError, TypeError):
                beep = 10
            if remaining <= beep:
                self._play_sound(self.config.alert_sound_path, 200)
        crossed = [a for a in self.config.alerts if remaining <= a.seconds < previous]
        if crossed:
            alert = min(crossed, key=lambda a: a.seconds)
            self.alert_triggered.emit(alert.color)
            if self.config.global_sound and alert.play_sound:
                self._play_sound(self.config.alert_sound_path, 200)


class _TimeoutWindow(QWidget):
    dismissed = Signal()

    def __init__(self, options):
        super().__init__(None, Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint |
                         Qt.WindowType.WindowStaysOnTopHint)
        self.black = options["end_mode"] == "black"
        self.allow_close = self.black or options["allow_close"]
        self.setWindowTitle("ppt time · 时间到")
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, False)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setObjectName("endWindow")
        self.setStyleSheet(
            "QWidget#endWindow {background: #000; color: white;}"
            if self.black else
            "QWidget#endWindow {background: #222832; color: white; border: 1px solid #657183;}"
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        self.label = QLabel(options["message"])
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setTextFormat(Qt.TextFormat.PlainText)
        self.label.setWordWrap(True)
        self.label.setStyleSheet("color: white; background: transparent;")
        layout.addWidget(self.label, 1)
        if self.allow_close:
            self.hint = QLabel("按 Esc、空格或回车返回")
            self.hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.hint.setStyleSheet("color: #c1c5cd; background: transparent; font-size: 14px;")
            layout.addWidget(self.hint)

    def place(self, screen):
        self.setScreen(screen)
        box = screen.geometry()
        if self.black:
            self.setGeometry(box)
        else:
            self.resize(min(box.width(), max(280, int(box.width() * .54))),
                        min(box.height(), max(140, int(box.height() * .32))))
            self.move(box.center() - self.rect().center())
        font = self.label.font()
        font.setBold(True)
        label_width = max(10, self.width() - 48)
        label_height = max(10, self.height() - 70)
        # Fit long custom messages on conference projectors and small secondary screens.
        flags = int(Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap | Qt.TextFlag.TextWrapAnywhere)
        for size in range(min(104 if self.black else 52, max(12, label_height // 3)), 9, -2):
            font.setPixelSize(size)
            measured = QFontMetrics(font).boundingRect(0, 0, label_width, 10000, flags, self.label.text())
            if measured.height() <= label_height and measured.width() <= label_width:
                break
        self.label.setFont(font)

    def keyPressEvent(self, event):
        keys = (Qt.Key.Key_Escape, Qt.Key.Key_Space, Qt.Key.Key_Enter, Qt.Key.Key_Return)
        if event.key() in keys:
            event.accept()
            if self.allow_close:
                self.dismissed.emit()
            return
        super().keyPressEvent(event)

    def closeEvent(self, event):
        event.ignore()
        if self.allow_close:
            self.dismissed.emit()


class TimingDialog(QDialog):
    def __init__(self, options, document="", page=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ppt time · 放映计时选项")
        self.setMinimumWidth(560)
        self.resize(640, 575)
        self.document = document
        self.page = page
        self.options = options
        layout = QVBoxLayout(self)
        self.pages = QCheckBox("仅显示页码，暂停计时和声音提醒")
        self.pages.setChecked(options["pages_only"])
        self.overtime = QCheckBox("时间用尽后继续显示超时秒数（+00:01）")
        self.overtime.setChecked(options["overtime"])
        layout.addWidget(self.pages)
        layout.addWidget(self.overtime)
        form = QFormLayout()
        self.end_mode = QComboBox()
        for label, value in (("关闭", "off"), ("文字提示", "notice"), ("全屏黑屏", "black")):
            self.end_mode.addItem(label, value)
        self.end_screen = QComboBox()
        for label, value in (("跟随计时条", "display"), ("放映屏", "slideshow"),
                             ("主屏幕", "primary"), ("所有屏幕", "all")):
            self.end_screen.addItem(label, value)
        self.message = QPlainTextEdit()
        self.message.setPlainText(options["message"])
        self.message.setMaximumHeight(66)
        self.allow_close = QCheckBox("允许用 Esc、空格或回车关闭文字提示")
        self.allow_close.setChecked(options["allow_close"])
        self.end_mode.setCurrentIndex(self.end_mode.findData(options["end_mode"]))
        self.end_screen.setCurrentIndex(self.end_screen.findData(options["end_screen"]))
        form.addRow("时间到", self.end_mode)
        form.addRow("提示屏幕", self.end_screen)
        form.addRow("提示文字", self.message)
        form.addRow(self.allow_close)
        layout.addLayout(form)
        label = QLabel("文稿与单页时长")
        font = label.font()
        font.setBold(True)
        label.setFont(font)
        layout.addWidget(label)
        explanation = QLabel("文稿留空可用于所有文件。页码填 * 表示整份文稿的开始时长；"
                             "指定页码后，每次进入该页都会重新开始该时长，未设规则的页面延续当前计时。")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["文稿路径或文件名", "页码", "时长（分:秒）"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        for row in options["rules"]:
            self._add(row["document"], row["page"], row["seconds"])
        layout.addWidget(self.table, 1)
        actions = QHBoxLayout()
        for text, action in (("添加规则", lambda: self._add("", 1, 60)),
                             ("选文稿", self._pick_document), ("当前放映页", self._add_current),
                             ("移除所选", self._remove_selected)):
            button = QPushButton(text)
            button.clicked.connect(action)
            button.setEnabled(text != "当前放映页" or bool(page))
            actions.addWidget(button)
        actions.addStretch()
        layout.addLayout(actions)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.pages.toggled.connect(self._enable_controls)
        self.end_mode.currentIndexChanged.connect(self._enable_controls)
        self._enable_controls()

    def _enable_controls(self, *_):
        timing = not self.pages.isChecked()
        mode = self.end_mode.currentData()
        self.overtime.setEnabled(timing)
        self.end_mode.setEnabled(timing)
        self.end_screen.setEnabled(timing and mode != "off")
        self.message.setEnabled(timing and mode != "off")
        self.allow_close.setEnabled(timing and mode == "notice")

    def _add(self, document, page, seconds):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for column, text in enumerate((document, str(page) if page else "*", _clock(seconds))):
            self.table.setItem(row, column, QTableWidgetItem(text))
        self.table.setCurrentCell(row, 0)

    def _pick_document(self):
        documents, _ = QFileDialog.getOpenFileNames(
            self, "选择文稿", "", "演示文稿 (*.ppt *.pptx *.pptm *.pps *.ppsx *.dps *.dpt *.pdf);;所有文件 (*)")
        for document in documents:
            self._add(document, 0, 300)

    def _add_current(self):
        self._add(self.document, self.page, 60)

    def _remove_selected(self):
        for row in sorted({cell.row() for cell in self.table.selectedIndexes()}, reverse=True):
            self.table.removeRow(row)

    def accept(self):
        rules = []
        seen = set()
        for row in range(self.table.rowCount()):
            cells = [self.table.item(row, i) for i in range(3)]
            try:
                entry = _rule(*(cell.text() if cell else "" for cell in cells))
                key = _source_key(entry["document"]), entry["page"]
                if key in seen:
                    raise ValueError("这份文稿的该页已设置过时长。")
                seen.add(key)
                rules.append(entry)
            except ValueError as error:
                self.table.setCurrentCell(row, 1)
                QMessageBox.warning(self, "请检查计时规则", f"第 {row + 1} 行：{error}")
                return
        if len(self.message.toPlainText().strip()) > 500:
            QMessageBox.warning(self, "请检查提示文字", "提示文字最多可填写 500 个字符。")
            return
        self.options = timing_options({
            "pages_only": self.pages.isChecked(), "overtime": self.overtime.isChecked(),
            "end_mode": self.end_mode.currentData(), "end_screen": self.end_screen.currentData(),
            "message": self.message.toPlainText(), "allow_close": self.allow_close.isChecked(),
            "rules": rules,
        })
        super().accept()


class ShowTiming(QObject):
    """Bind presentation identity and the visible screens to the shared clock."""

    def __init__(self, controller, display, options, save_callback, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.display = display
        self.save_callback = save_callback
        self.options = timing_options(options)
        self.snapshot = None
        self._entered = False
        self._document = ""
        self._page = None
        self._applied_rule = None
        self._deck_rule = None
        self._windows = {}
        self._active_notice = False
        self._dialog = None
        self._resume_after_page_mode = False
        self._entered_during_page_mode = False
        self._apply_options()
        controller.expired.connect(self._timeout)
        controller.stage_changed.connect(self.dismiss)
        app = QApplication.instance()
        app.screenAdded.connect(self._place_windows)
        app.screenRemoved.connect(self._place_windows)

    @property
    def document(self):
        if self.snapshot is None:
            return ""
        for field in ("document", "document_path", "source_path"):
            value = getattr(self.snapshot, field, "")
            if isinstance(value, str) and value:
                return value
        return ""

    def _apply_options(self):
        clock = self.controller
        pages = self.options["pages_only"]
        if pages and clock.timing_enabled:
            self._resume_after_page_mode = not clock.paused
        old_enabled = clock.timing_enabled
        clock.set_timing_enabled(not pages)
        clock.set_overtime(self.options["overtime"])
        self.display.set_pages_only(pages)
        if pages:
            self.dismiss()
        elif not old_enabled:
            if self._resume_after_page_mode or (
                self._entered_during_page_mode and self.controller.config.ppt_auto_start
            ):
                if clock.paused:
                    clock.toggle_pause()
            self._resume_after_page_mode = False
            self._entered_during_page_mode = False

    def edit(self):
        if self._dialog is not None:
            self._dialog.raise_()
            self._dialog.activateWindow()
            return
        self._dialog = TimingDialog(self.options, self.document, self._page)
        try:
            if self._dialog.exec() == QDialog.DialogCode.Accepted:
                self.set_options(self._dialog.options)
        finally:
            self._dialog.deleteLater()
            self._dialog = None

    def set_options(self, options):
        updated = timing_options(options)
        if updated == self.options:
            return
        before = self.options
        self.options = updated
        self.save_callback(updated)
        self._apply_options()
        if self._active_notice and any(before[key] != updated[key] for key in
                                       ("end_mode", "message", "end_screen", "allow_close")):
            self.dismiss()
            self._timeout()
        if self.snapshot and before["rules"] != updated["rules"]:
            self.update_snapshot(self.snapshot, settings_changed=True)

    def _match(self, document, page):
        if page is None:
            return None
        normalized = _source_key(document)
        basename = ntpath.basename(normalized)
        matches = [r for r in self.options["rules"] if r["page"] == page]
        for key in (normalized, basename, ""):
            for entry in matches:
                if key == _source_key(entry["document"]):
                    return entry
        return None

    def update_snapshot(self, snapshot, settings_changed=False):
        """Called after a stable show edge or an updated slide sample."""
        self.snapshot = snapshot if snapshot and snapshot.active else None
        if self.snapshot is None:
            self._entered = False
            self._document = ""
            self._page = None
            self._applied_rule = None
            self._deck_rule = None
            self._entered_during_page_mode = False
            if not self.display.manual:
                self._resume_after_page_mode = False
                self.dismiss()
            return
        document = _source_key(self.document)
        page = self.snapshot.current if self.snapshot.page_text else None
        known_change = bool(document and self._document and document != self._document)
        new_deck = not self._entered or known_change
        learned_document = bool(document and not self._document)
        new_page = page is not None and self._page != page
        if self._active_notice:
            self._place_windows()
        if not (new_deck or learned_document or new_page or settings_changed):
            return
        if known_change:
            self._page = None
        if new_deck and self.options["pages_only"]:
            self._entered_during_page_mode = True
        self._entered = True
        self._document = document or self._document
        if page is not None:
            self._page = page
        rule = self._match(document, page)
        if new_deck or learned_document or settings_changed:
            deck_rule = self._match(document, 0)
            if rule is None and (new_deck or deck_rule != self._deck_rule):
                rule = deck_rule
            self._deck_rule = deck_rule
        if rule is None:
            self._applied_rule = None
            return
        if not (new_deck or new_page) and rule == self._applied_rule:
            return
        self._applied_rule = rule
        start = not self.controller.paused or (
            self.controller.at_end and (self.controller.config.ppt_auto_start or self.display.manual))
        self.controller.set_duration(rule["seconds"], start=start)

    def set_duration(self, seconds, start=None):
        self.controller.set_duration(seconds, start)

    def _timeout(self):
        if self.options["end_mode"] == "off" or self.options["pages_only"]:
            return
        if not self.display.manual and not self._entered:
            return
        self._active_notice = True
        self._place_windows()

    def _place_windows(self, *_):
        if not self._active_notice:
            return
        selection = self.options["end_screen"]
        if selection == "display":
            selection = self.display.config.presentation_screen
        targets = self.display._target_screens(selection)
        for screen in list(self._windows):
            if screen not in targets:
                self._remove_window(screen)
        added = False
        for screen in targets:
            if screen not in self._windows:
                window = _TimeoutWindow(self.options)
                window.dismissed.connect(self.dismiss)
                screen.geometryChanged.connect(self._place_windows)
                self._windows[screen] = window
                window.place(screen)
                window.show()
                added = True
            else:
                self._windows[screen].place(screen)
        if added and targets:
            window = self._windows[targets[0]]
            window.raise_()
            window.activateWindow()
            window.setFocus()

    def dismiss(self, *_):
        self._active_notice = False
        for screen in list(self._windows):
            self._remove_window(screen)

    def _remove_window(self, screen):
        window = self._windows.pop(screen)
        try:
            screen.geometryChanged.disconnect(self._place_windows)
        except (TypeError, RuntimeError):
            pass
        window.hide()
        window.deleteLater()

    def close(self):
        self.dismiss()
        if self._dialog is not None:
            self._dialog.reject()
