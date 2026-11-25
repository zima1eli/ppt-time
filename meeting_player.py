"""Audience presentation windows and document-scoped meeting controls."""

from __future__ import annotations

from dataclasses import replace
import os
from pathlib import Path
import queue
import re
import shutil
import subprocess
import sys
import threading
import time

from PySide6.QtCore import QObject, QRect, Qt, QTimer, QUrl, Signal, Slot
from PySide6.QtGui import QGuiApplication, QPainter, QShortcut, QKeySequence
from PySide6.QtWidgets import QWidget, QVBoxLayout

from slideshow import (SlideSnapshot, _confirmed_fullscreen, _office_hwnd, _pid, _process_name,
                       _running_office_apps, _screen_info, _window_snapshot, is_adobe_show, windows_show_window)


OFFICE = {".ppt", ".pptx", ".pps", ".ppsx", ".dps"}
VIDEOS = {".mp4", ".m4v", ".mov", ".mkv", ".webm", ".avi", ".wmv", ".mpg", ".mpeg"}
SUPPORTED_EXTENSIONS = OFFICE | VIDEOS | {".pdf"}


class _AudienceWindow(QWidget):
    changed = Signal(object)
    closed = Signal(bool)
    failed = Signal(str)

    def __init__(self, path):
        super().__init__(None, Qt.WindowType.Window | Qt.WindowType.FramelessWindowHint)
        self.path = path
        self.ended = False
        self.setWindowTitle(f"ppt time · {path.name}")
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.setStyleSheet("background-color: black")
        self.shortcuts = []
        for key, action in (("Right", "next"), ("Down", "next"), ("PgDown", "next"),
                            ("Left", "previous"), ("Up", "previous"), ("PgUp", "previous"),
                            ("Esc", "stop"), ("Space", "space"), ("Return", "next")):
            shortcut = QShortcut(QKeySequence(key), self)
            shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
            shortcut.activated.connect(lambda command=action: self.control(command))
            self.shortcuts.append(shortcut)

    def present(self, screen):
        target = screen or QGuiApplication.primaryScreen()
        self.setScreen(target)
        self.setGeometry(target.geometry())
        self.showFullScreen()
        self.raise_()
        self.activateWindow()
        self.notify()

    def snapshot(self, current=None, total=None):
        screen = self.screen()
        center = screen.geometry().center()
        return SlideSnapshot(True, current, total, screen.name(), (center.x(), center.y()),
                             str(self.path), self.ended, int(self.winId()) if sys.platform == "win32" else 0)

    def notify(self):
        self.changed.emit(self.snapshot())

    def finish(self):
        self.ended = True
        self.close()

    def closeEvent(self, event):
        self.closed.emit(self.ended)
        super().closeEvent(event)


class _PdfWindow(_AudienceWindow):
    def __init__(self, path):
        from PySide6.QtPdf import QPdfDocument
        super().__init__(path)
        self.document = QPdfDocument(self)
        error = self.document.load(str(path))
        if error != QPdfDocument.Error.None_ or self.document.pageCount() < 1:
            self.document.close()
            self.deleteLater()
            if error == QPdfDocument.Error.IncorrectPassword:
                raise ValueError("PDF 需要密码，请选择可直接打开的 PDF。")
            raise ValueError("无法读取这个 PDF。")
        self.index = 0
        self.wheel_delta = 0
        self.cached_key = None
        self.image = None
        for key, index in (("Home", 0), ("End", self.document.pageCount() - 1)):
            shortcut = QShortcut(QKeySequence(key), self)
            shortcut.activated.connect(lambda page=index: self._go_to(page))
            self.shortcuts.append(shortcut)

    def snapshot(self):
        return super().snapshot(self.index + 1, self.document.pageCount())

    def control(self, action):
        if action == "stop":
            self.close()
        elif action == "first":
            self._go_to(0)
        elif action == "last":
            self._go_to(self.document.pageCount() - 1)
        elif action in {"next", "space", "previous"}:
            new_index = max(0, self.index + (-1 if action == "previous" else 1))
            if new_index >= self.document.pageCount():
                self.finish()
            else:
                self._go_to(new_index)

    def _go_to(self, index):
        if index != self.index:
            self.index = index
            self.update()
            self.notify()

    def wheelEvent(self, event):
        self.wheel_delta += event.angleDelta().y()
        if abs(self.wheel_delta) >= 120:
            self.control("previous" if self.wheel_delta > 0 else "next")
            self.wheel_delta = 0
        event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() in {Qt.MouseButton.LeftButton, Qt.MouseButton.RightButton}:
            self.control("next" if event.button() == Qt.MouseButton.LeftButton else "previous")

    def closeEvent(self, event):
        self.document.close()
        super().closeEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), Qt.GlobalColor.black)
        if self.document.pageCount() == 0:
            return
        size = self.document.pagePointSize(self.index).toSize()
        if size.isEmpty():
            return
        size.scale(self.size(), Qt.AspectRatioMode.KeepAspectRatio)
        pixel_size = size * self.devicePixelRatioF()
        key = (self.index, pixel_size.width(), pixel_size.height())
        if key != self.cached_key:
            self.image = self.document.render(self.index, pixel_size)
            self.cached_key = key
        if self.image is not None and not self.image.isNull():
            target = QRect(0, 0, size.width(), size.height())
            target.moveCenter(self.rect().center())
            painter.drawImage(target, self.image)


class _VideoWindow(_AudienceWindow):
    def __init__(self, path):
        from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
        from PySide6.QtMultimediaWidgets import QVideoWidget
        super().__init__(path)
        video = QVideoWidget(self)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(video)
        self.media = QMediaPlayer(self)
        self.audio = QAudioOutput(self)
        self.media.setAudioOutput(self.audio)
        self.media.setVideoOutput(video)
        self.media.errorOccurred.connect(self._playback_error)
        self.media.mediaStatusChanged.connect(self._media_status)

    def present(self, screen):
        super().present(screen)
        self.media.setSource(QUrl.fromLocalFile(str(self.path)))
        self.media.play()

    def _media_status(self, status):
        if status == self.media.MediaStatus.EndOfMedia:
            self.finish()

    def _playback_error(self, error, detail):
        self.failed.emit(f"视频播放失败：{detail}")
        self.close()

    def control(self, action):
        if action == "stop":
            self.close()
        elif action in {"first", "last"} and self.media.isSeekable():
            self.media.setPosition(0 if action == "first" else max(0, self.media.duration() - 1))
        elif action == "space":
            if self.media.isPlaying():
                self.media.pause()
            else:
                self.media.play()
        elif action in {"next", "previous"} and self.media.isSeekable():
            self.media.setPosition(max(0, min(self.media.duration(), self.media.position()
                                              + (10000 if action == "next" else -10000))))

    def closeEvent(self, event):
        self.media.stop()
        super().closeEvent(event)


def _open_windows_deck(path):
    import winreg
    from win32com.client import dynamic
    prefer_wps = path.suffix.lower() == ".dps"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
                            + "\\" + path.suffix + r"\UserChoice") as key:
            preference = winreg.QueryValueEx(key, "ProgId")[0].lower()
        prefer_wps = prefer_wps or "kingsoft" in preference or "wpp" in preference or "wps" in preference
    except OSError:
        pass
    apps = ["KWPP.Application", "WPP.Application", "PowerPoint.Application"]
    if not prefer_wps:
        apps.insert(0, apps.pop())
    for progid in apps:
        try:
            app = dynamic.Dispatch(progid)
        except Exception:
            continue
        # An already-open document belongs to the operator. Only close a
        # presentation we opened ourselves, including if starting the show fails.
        presentations = app.Presentations
        doc = None
        for i in range(1, int(presentations.Count) + 1):
            existing = presentations.Item(i)
            if _same_office_file(existing.FullName, path):
                doc = existing
                break
        owned = doc is None
        if owned:
            # Office automation defaults can differ from an interactive open.
            # Disable document macros while reading the selected meeting file.
            security = None
            try:
                security = app.AutomationSecurity
                app.AutomationSecurity = 3
            except Exception:
                if progid == "PowerPoint.Application":
                    raise ValueError("无法设置文稿的安全打开模式。")
            try:
                doc = presentations.Open(str(path), -1, 0, -1)
            finally:
                if security is not None:
                    app.AutomationSecurity = security
        try:
            show = None
            if not owned:
                try:
                    current = doc.SlideShowWindow
                    if int(current.View.State) in (1, 2, 3, 4):
                        show = current
                except Exception:
                    pass  # A document in editing mode has no SlideShowWindow.
            if show is None:
                show = doc.SlideShowSettings.Run()
            return app, doc, show, owned
        except Exception:
            if owned:
                doc.Close()
            raise
    raise ValueError("未找到可启动放映的 PowerPoint 或 WPS 演示。")


def _same_office_file(left, right):
    if os.path.normcase(str(left)) == os.path.normcase(str(right)):
        return True
    try:
        return os.path.samefile(left, right)  # Office may expand a junction or short path.
    except (OSError, TypeError, ValueError):
        return False


def _read_opened_show(owned, path):
    """Sample only the exact show returned by Run, on its original COM thread."""
    import win32gui
    if not owned or owned[2] is None:
        return SlideSnapshot()
    app, _, show, _ = owned
    hwnd = _office_hwnd(show)
    if hwnd and (not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd)):
        return SlideSnapshot()
    try:
        show_count = int(app.SlideShowWindows.Count)
    except Exception:
        show_count = None  # Some WPS versions omit the application's collection.
    if show_count == 0 or (show_count and show_count > 1 and not show.Active):
        return SlideSnapshot()
    if hwnd:
        foreground = win32gui.GetForegroundWindow()
        if foreground and foreground != hwnd and _pid(foreground) != _pid(hwnd):
            # A separate foreground slideshow takes precedence over this still
            # open document. Settings, the clock and presenter controls do not.
            foreground_class = win32gui.GetClassName(foreground).lower()
            if (foreground_class in {"screenclass", "kslideshow", "wps_slideshow", "wpp_slideshow_class"}
                    or _confirmed_fullscreen(foreground) or is_adobe_show(foreground)):
                return SlideSnapshot()
    snapshot = _window_snapshot(show, _screen_info(hwnd), hwnd)
    if snapshot.document:
        expected = str(path) if os.path.isabs(snapshot.document) else Path(path).name
        if not _same_office_file(snapshot.document, expected):
            return SlideSnapshot()
    if snapshot.ended:
        return SlideSnapshot(document=str(path), ended=True)
    if not snapshot.page_text:
        return SlideSnapshot()
    if not os.path.isabs(snapshot.document):
        snapshot = replace(snapshot, document=str(path))
    return snapshot


def _windows_command(action, snapshot):
    import win32api
    import win32con
    import win32gui
    hwnd = snapshot.hwnd
    if not hwnd or not win32gui.IsWindow(hwnd):
        raise ValueError("当前没有可控制的放映窗口。")
    name = _process_name(hwnd)
    if name not in {"powerpnt.exe", "pptview.exe", "wpp.exe", "wps.exe", "wpspdf.exe", "acrord32.exe", "acrobat.exe"}:
        raise ValueError("当前窗口已不再属于演示程序。")
    if name in {"acrord32.exe", "acrobat.exe"}:
        if not is_adobe_show(hwnd):
            raise ValueError("PDF 已退出放映。")
    elif name == "wpspdf.exe" or (name == "wps.exe" and snapshot.document.lower().endswith(".pdf")):
        from wps_pdf import _wps_pdf_window
        if _wps_pdf_window(snapshot.document, hwnd) != hwnd:
            raise ValueError("WPS PDF 已离开放映画面。")
    else:
        window, _ = windows_show_window(hwnd)
        if window is not None:
            if snapshot.document:
                current = str(window.Presentation.FullName)
                if os.path.normcase(current) != os.path.normcase(snapshot.document):
                    raise ValueError("放映文件已切换。")
            getattr(window.View, {"next":"Next", "previous":"Previous", "stop":"Exit",
                                  "first":"First", "last":"Last"}[action])()
            return
        # Versions of WPS without live COM still accept keyboard messages sent
        # specifically to their slideshow window, including while a desk has focus.
        class_name = win32gui.GetClassName(hwnd).lower()
        bounds = win32api.GetMonitorInfo(win32api.MonitorFromWindow(hwnd))["Monitor"]
        is_full = all(abs(a-b) <= 4 for a, b in zip(win32gui.GetWindowRect(hwnd), bounds))
        has_caption = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE) & win32con.WS_CAPTION
        if class_name not in {"screenclass", "kslideshow", "wps_slideshow", "wpp_slideshow_class"} and (not is_full or has_caption):
            raise ValueError("演示程序已离开放映画面。")
    key = {"next":win32con.VK_NEXT, "previous":win32con.VK_PRIOR, "stop":win32con.VK_ESCAPE,
           "first":win32con.VK_HOME, "last":win32con.VK_END}[action]
    scan = win32api.MapVirtualKey(key, 0)
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, key, 1 | (scan << 16))
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, key, 1 | (scan << 16) | (3 << 30))


def _start_current_windows():
    import win32api
    import win32con
    import win32gui
    foreground = win32gui.GetForegroundWindow()
    foreground_pid = _pid(foreground)
    candidates = []
    for app in _running_office_apps():
        try:
            doc = app.ActivePresentation
            if doc is None or int(doc.Slides.Count) <= 0:
                continue
            candidates.append((app, doc))
        except Exception:
            continue
    preferred = [(app, doc) for app, doc in candidates
                 if foreground_pid and _pid(_office_hwnd(app)) == foreground_pid]
    choices = preferred or candidates
    if len(choices) == 1:
        app, doc = choices[0]
        try:
            if int(app.SlideShowWindows.Count) > 0:
                return
        except Exception:
            pass
        doc.SlideShowSettings.Run()
        return
    if choices:
        raise ValueError("多个演示程序已打开文稿，请先切换到要放映的编辑窗口。")

    def is_editor(hwnd):
        if not hwnd or not win32gui.IsWindowVisible(hwnd):
            return False
        try:
            name = _process_name(hwnd)
            if name not in {"powerpnt.exe", "wpp.exe", "wps.exe"}:
                return False
            title = win32gui.GetWindowText(hwnd)
            cls = win32gui.GetClassName(hwnd).lower()
            if cls == "#32770" or "slideshow" in cls or cls == "screenclass":
                return False
            if name == "powerpnt.exe":
                return cls == "pptframeclass" and bool(title.strip()) and title.strip().lower() != "powerpoint"
            # WPS may share an executable with Writer and Spreadsheets. For
            # keyboard-only editions require an explicit presentation filename.
            return bool(re.search(r"\.(?:pptx?|ppsx?|dps)(?:$|[\s\]\)】）\-—])", title, re.I))
        except Exception:
            return False

    if is_editor(foreground):
        target = foreground
    else:
        matches = []
        def visit(hwnd, _):
            if is_editor(hwnd):
                matches.append(hwnd)
            return True
        win32gui.EnumWindows(visit, None)
        if len(matches) != 1:
            raise ValueError("请打开要放映的 PowerPoint 或 WPS 演示文稿并切换到编辑窗口。")
        target = matches[0]
    if not is_editor(target):
        raise ValueError("演示编辑窗口已关闭。")
    key = win32con.VK_F5
    scan = win32api.MapVirtualKey(key, 0)
    win32gui.PostMessage(target, win32con.WM_KEYDOWN, key, 1 | (scan << 16))
    win32gui.PostMessage(target, win32con.WM_KEYUP, key, 1 | (scan << 16) | (3 << 30))


class MeetingPlayer(QObject):
    snapshot_changed = Signal(object)
    error = Signal(str)
    _sample_ready = Signal(int, object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.window = None
        self.owned_snapshot = None
        self._queue = queue.Queue(32)
        self._thread = None
        self._closing = threading.Event()
        self._generation = 0
        self._sample_ready.connect(self._accept_office_sample, Qt.ConnectionType.QueuedConnection)
        self._sample_expiry = QTimer(self)
        self._sample_expiry.setSingleShot(True)
        self._sample_expiry.setInterval(1800)
        self._sample_expiry.timeout.connect(self._clear_office_snapshot)

    def _enqueue(self, item, generation=None):
        try:
            self._queue.put_nowait((*item, self._generation if generation is None else generation))
        except queue.Full:
            self.error.emit("演示程序正在响应，请稍后再试。")
            return False
        if self._thread is None:
            self._thread = threading.Thread(target=self._worker, name="Meeting document control", daemon=True)
            self._thread.start()
        return True

    def _worker(self):
        if sys.platform == "win32":
            import pythoncom
            pythoncom.CoInitialize()
        owned = None
        watch_path = None
        watch_generation = -1
        next_sample = 0.0
        try:
            while not self._closing.is_set():
                timeout = max(0.0, next_sample - time.monotonic()) if watch_path is not None else None
                try:
                    operation, data, generation = self._queue.get(timeout=timeout)
                except queue.Empty:
                    operation, data, generation = "sample", None, watch_generation
                if self._closing.is_set():
                    break
                try:
                    if owned and watch_generation < self._generation:
                        self._release(owned)
                        owned = watch_path = None
                    if generation < self._generation:
                        continue
                    if operation == "sample":
                        pass
                    elif operation in {"release", "open"}:
                        watch_path = None
                        if owned:
                            self._release(owned)
                            owned = None
                        if operation == "release":
                            continue
                        if sys.platform == "win32":
                            owned = _open_windows_deck(data)
                            watch_path, watch_generation = data, generation
                            next_sample = 0.0
                        elif sys.platform == "darwin":
                            self._mac_command("open", str(data))
                        else:
                            executable = shutil.which("libreoffice") or shutil.which("soffice")
                            if not executable:
                                raise ValueError("请安装 LibreOffice 以放映演示文稿。")
                            subprocess.Popen([executable, "--show", str(data)])
                    elif operation == "start_current":
                        _start_current_windows()
                    elif sys.platform == "win32":
                        shown = _office_hwnd(owned[2]) if owned else 0
                        same_document = (data.document and watch_path is not None and
                                         _same_office_file(data.document, watch_path))
                        same_show = (owned is not None and owned[2] is not None and generation == watch_generation
                                     and ((shown and shown == data.hwnd and (not data.document or same_document))
                                          or (not data.hwnd and same_document)))
                        if same_show:
                            getattr(owned[2].View, {"next":"Next", "previous":"Previous", "stop":"Exit",
                                                   "first":"First", "last":"Last"}[operation])()
                        else:
                            _windows_command(operation, data)
                        if same_show and operation == "stop":
                            self._sample_ready.emit(watch_generation, SlideSnapshot())
                            self._release(owned, exit_show=False)
                            owned = watch_path = None
                        next_sample = 0.0
                    elif sys.platform == "darwin":
                        self._mac_command(operation, data.document)
                    else:
                        raise ValueError("此平台的外部文稿请使用演示程序翻页。")
                except Exception as exc:
                    if not self._closing.is_set():
                        self.error.emit(f"演示操作失败：{exc}")
                if watch_path is not None and time.monotonic() >= next_sample:
                    started = time.monotonic()
                    try:
                        snapshot = _read_opened_show(owned, watch_path)
                    except Exception:
                        snapshot = SlideSnapshot()
                    if time.monotonic() - started > 1.8:
                        snapshot = SlideSnapshot()
                    self._sample_ready.emit(watch_generation, snapshot)
                    next_sample = time.monotonic() + .4
        finally:
            if owned:
                self._release(owned)
            if sys.platform == "win32":
                pythoncom.CoUninitialize()

    @staticmethod
    def _release(owned, exit_show=True):
        _, doc, show, opened_here = owned
        if opened_here:
            if exit_show:
                try:
                    show.View.Exit()
                except Exception:
                    pass
            try:
                doc.Close()
            except Exception:
                pass

    @staticmethod
    def _mac_command(operation, path):
        if operation == "open":
            script = '''
on run argv
    tell application id "com.microsoft.Powerpoint"
        open (POSIX file (item 1 of argv))
        run slide show slide show settings of active presentation
    end tell
end run'''
        else:
            # Operate only the active show, never the current application's UI.
            action = {"next":"go to next slide", "previous":"go to previous slide", "stop":"exit slide show",
                      "first":"go to first slide", "last":"go to last slide"}[operation]
            script = '''
on run argv
    tell application id "com.microsoft.Powerpoint"
        if (count of slide show windows) is not 1 then error "未找到唯一的放映窗口"
        set w to slide show window 1
        if (item 1 of argv) is not "" then
            if (full name of presentation of w as text) is not (item 1 of argv) then error "放映文件已切换"
        end if
        ''' + action + ''' slide show view of w
    end tell
end run'''
        subprocess.run(["/usr/bin/osascript", "-e", script, path], check=True,
                       capture_output=True, text=True, timeout=25)

    def open_file(self, path, screen=None):
        try:
            path = Path(path).expanduser().resolve(strict=True)
            suffix = path.suffix.lower()
            if not path.is_file() or suffix not in SUPPORTED_EXTENSIONS:
                raise ValueError("请选择 PPT、PDF 或视频文件。")
            if suffix in OFFICE:
                generation = self._generation + 1
                if not self._enqueue(("open", path), generation):
                    return False
                self._generation = generation
                self._clear_office_snapshot()
                self._close_window()
                return True
            window = _PdfWindow(path) if suffix == ".pdf" else _VideoWindow(path)
            self._generation += 1
            self._clear_office_snapshot()
            self._close_window()
            if self._thread is not None:
                self._enqueue(("release", None))
            self.window = window
            window.changed.connect(self._on_snapshot)
            window.closed.connect(self._on_closed)
            window.failed.connect(self.error.emit)
            window.present(screen)
            return True
        except (ValueError, OSError, ImportError, TypeError, RuntimeError) as exc:
            self.error.emit(str(exc))
            return False

    def _clear_office_snapshot(self):
        if self.window is None and self.owned_snapshot is not None:
            previous = self.owned_snapshot
            self.owned_snapshot = None
            self._sample_expiry.stop()
            self.snapshot_changed.emit(SlideSnapshot(document=previous.document))

    @Slot(int, object)
    def _accept_office_sample(self, generation, snapshot):
        if generation != self._generation or self.window is not None or self._closing.is_set():
            return
        if snapshot.active and snapshot.page_text:
            self._sample_expiry.start()
            if snapshot != self.owned_snapshot:
                self._on_snapshot(snapshot)
        elif self.owned_snapshot is not None:
            self._sample_expiry.stop()
            self.owned_snapshot = None
            self.snapshot_changed.emit(snapshot)

    def _on_snapshot(self, snapshot):
        self.owned_snapshot = snapshot
        self.snapshot_changed.emit(snapshot)

    def _on_closed(self, ended):
        was = self.owned_snapshot
        self.owned_snapshot = self.window = None
        self.snapshot_changed.emit(SlideSnapshot(document=was.document if was else "", ended=ended))

    def _close_window(self):
        if self.window is not None:
            self.window.close()

    def command(self, action, snapshot=None):
        if action == "prev":
            action = "previous"
        if action not in {"previous", "next", "stop", "first", "last"}:
            return False
        if self.window is not None:
            self.window.control(action)
            return True
        if snapshot is None:
            snapshot = self.owned_snapshot
        if snapshot is None or not snapshot.active:
            self.error.emit("当前没有正在放映的文稿。")
            return False
        return self._enqueue((action, snapshot))

    def start_current(self):
        if self.window is not None:
            return True
        if sys.platform != "win32":
            self.error.emit("请在演示程序内启动放映，或从会务台选择文稿。")
            return False
        return self._enqueue(("start_current", None))

    def close(self):
        self._clear_office_snapshot()
        self._close_window()
        self._closing.set()
        try:
            self._queue.put_nowait(("close", None, self._generation))
        except queue.Full:
            pass
