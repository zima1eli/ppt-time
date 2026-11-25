"""Place the audience-facing timer on the selected screens."""
from pathlib import Path
from dataclasses import asdict
from queue import Empty, SimpleQueue
import sys
import threading
import time

from PySide6.QtCore import QObject, QPoint, QTimer, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox, QSystemTrayIcon

from overlay_options import OverlayDialog, OverlayOptions, capture_supported, sample_underlay
from presentation_overlay import PresentationOverlay


class PresentationDisplay(QObject):
    _strip_context_requested = Signal(object, int, int)

    def __init__(self, config, float_bar, parent=None):
        super().__init__(parent)
        self.config = config
        self.float_bar = float_bar
        self.snapshot = None
        self.manual = False
        self.time_text = ""
        self.alert_color = ""
        self.overlays = {}
        self.preview = None
        self.pages_only = False
        self.appearance = OverlayOptions()
        self.appearance_path = None
        self._sample_timer = QTimer(self)
        self._sample_timer.setInterval(650)
        self._sample_timer.timeout.connect(self._sample_background)
        self._sample_worker = None
        self._samples = SimpleQueue()
        self._last_sample = 0.
        self._countdown = None
        self._clock_running = False
        self._flash_until = 0.
        self._flash_dim = False
        self._flash_timer = QTimer(self)
        self._flash_timer.setInterval(500)
        self._flash_timer.timeout.connect(self._step_flash)
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.timeout.connect(self.end_preview)
        app = QApplication.instance()
        self._projection_context = None
        if sys.platform == "win32" and app.platformName() == "windows" and hasattr(float_bar, "show_context_menu_at"):
            from projection_context import ProjectionContext
            self._projection_context = ProjectionContext(self._strip_context_requested.emit)
            self._strip_context_requested.connect(self._show_strip_menu, Qt.ConnectionType.QueuedConnection)
        app.screenAdded.connect(self.refresh)
        app.screenRemoved.connect(self.refresh)

    def set_appearance_path(self, path):
        self.appearance_path = Path(path)
        self.appearance = OverlayOptions.load(self.appearance_path)
        self.refresh()

    def edit_appearance(self, parent=None):
        dialog = OverlayDialog(self.appearance, parent)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            options = dialog.options()
            if self.appearance_path is not None:
                try:
                    options.save(self.appearance_path)
                except OSError as error:
                    QMessageBox.warning(parent, "ppt time", f"无法保存放映条外观设置。\n{error}")
                    return
            self.appearance = options
            self.refresh()

    def remote_appearance(self, patch):
        if not isinstance(patch, dict) or not patch:
            return False
        if patch.get("exclude_capture") is True and not capture_supported():
            return False
        try:
            updated = OverlayOptions.from_mapping({**asdict(self.appearance), **patch}, strict=True)
            if self.appearance_path is not None:
                updated.save(self.appearance_path)
        except (ValueError, TypeError, OSError):
            return False
        self.appearance = updated
        self.refresh()
        return True

    def set_countdown_state(self, seconds, is_countdown, running=True):
        if self._countdown is not None and self._countdown <= 0 < seconds:
            self._flash_until = 0.
        self._clock_running = running
        self._countdown = seconds if is_countdown else None
        self._sync_flash()

    def set_timer_running(self, running):
        self._clock_running = running
        self._sync_flash()

    def flash_expired(self):
        if self.appearance.flash_expiry and not self.pages_only and self.overlays:
            self._flash_until = time.monotonic() + 4
            self._sync_flash()

    def _sync_flash(self):
        last_minute = (self.appearance.flash_last_minute and self._clock_running and
                       self._countdown is not None and 0 < self._countdown <= 60)
        end_flash = self.appearance.flash_expiry and time.monotonic() < self._flash_until
        if self.overlays and not self.pages_only and (self.snapshot or self.manual) and (last_minute or end_flash):
            if not self._flash_timer.isActive():
                self._flash_timer.start()
        else:
            self._flash_timer.stop()
            self._flash_dim = False
            for overlay in self.overlays.values():
                overlay.set_flash_dim(False)

    def _step_flash(self):
        self._sync_flash()
        if self._flash_timer.isActive():
            self._flash_dim = not self._flash_dim
            for overlay in self.overlays.values():
                overlay.set_flash_dim(self._flash_dim)

    def set_pages_only(self, enabled):
        if self.pages_only != bool(enabled):
            self.pages_only = bool(enabled)
            self.refresh()

    def set_snapshot(self, snapshot):
        if snapshot is not None and not snapshot.active:
            snapshot = None
        previous = self.snapshot
        if previous == snapshot:
            return
        self.snapshot = snapshot
        if previous is not None and snapshot is not None:
            if (previous.screen_name, previous.screen_point) == (snapshot.screen_name, snapshot.screen_point):
                self._paint_overlays()
                self._sync_sampling()
                return
        if previous is None or snapshot is None:
            # A preview belongs to the desktop or show in which it was started.
            self.preview = None
            self._preview_timer.stop()
        self.refresh()

    def start_preview(self, style, screen, pages):
        self.preview = (style if style != "classic" else "shoulder", screen, pages)
        self._preview_timer.start(8000)
        self.refresh()

    def set_manual(self, enabled):
        self.manual = enabled
        self.preview = None
        self._preview_timer.stop()
        self.refresh()

    def end_preview(self):
        if self.preview is not None:
            self.preview = None
            self._preview_timer.stop()
            self.refresh()

    def _target_screens(self, selection):
        screens = QApplication.screens()
        if selection == "all":
            return screens
        target = None
        if selection == "slideshow" and self.snapshot is not None:
            name = self.snapshot.screen_name.casefold().rsplit("\\", 1)[-1]
            if name:
                target = next((s for s in screens
                               if s.name().casefold().rsplit("\\", 1)[-1] == name), None)
            if target is None and self.snapshot.screen_point is not None:
                target = QApplication.screenAt(QPoint(*self.snapshot.screen_point))
        target = target or QApplication.primaryScreen()
        return [target] if target else []

    def refresh(self, *_):
        c = self.config
        style, selection, _ = self.preview or (
            c.presentation_style, c.presentation_screen, c.presentation_pages)
        if self.pages_only and style == "classic":
            style = "notch"
        showing = self.snapshot is not None or self.manual
        show_overlay = (showing or self.preview is not None) and style != "classic"
        # Keep a reachable control surface on desktops without a notification area.
        show_controls = ((showing and style == "classic") or
                         (not showing and not c.presentation_only) or
                         not QSystemTrayIcon.isSystemTrayAvailable())
        self.float_bar.setVisible(show_controls)
        targets = self._target_screens(selection) if show_overlay else []
        capture_surface = (self.appearance.exclude_capture and capture_supported() and
                           QApplication.platformName() == "windows")
        for screen in list(self.overlays):
            if screen not in targets or self.overlays[screen].capture_surface != capture_surface:
                overlay = self.overlays.pop(screen)
                overlay.close()
                overlay.deleteLater()
        for screen in targets:
            if screen not in self.overlays:
                self.overlays[screen] = PresentationOverlay(capture_surface)
            overlay = self.overlays[screen]
            foreground = QColor(c.color)
            foreground.setAlphaF(max(0.0, min(1.0, 1 - c.font_transparency / 100)))
            overlay.configure(style, c.ui_scale, c.bg_qcolor(), foreground,
                              max(0.3, min(1.0, 1 - c.global_transparency / 100)), self.appearance,
                              desktop_time_size=getattr(c, "font_size", 32))
            overlay.place_on(screen)
        self._paint_overlays()
        self._sync_sampling()
        self._sync_flash()

    def update_content(self, time_text, alert_color):
        self.time_text, self.alert_color = time_text, alert_color
        self._paint_overlays()

    def _paint_overlays(self):
        pages = self.preview[2] if self.preview else self.config.presentation_pages
        page = ""
        if (pages or self.pages_only) and self.snapshot is not None and not self.snapshot.ended:
            page = self.snapshot.page_text
        classic = self.config.presentation_style == "classic" and not self.preview and not self.pages_only
        if hasattr(self.float_bar, "set_page_text"):
            self.float_bar.set_page_text(page if classic else "")

        time_text, alert = self.time_text, self.alert_color
        if self.preview is not None and self.snapshot is None and not self.manual:
            time_text, alert = "22:02", ""
            page = "16/53" if pages or self.pages_only else ""
        for overlay in self.overlays.values():
            overlay.set_content("" if self.pages_only else time_text, page, alert)
            overlay.setVisible(not self.pages_only or bool(page))
            overlay.set_flash_dim(self._flash_dim)
        if self._projection_context is not None:
            self._projection_context.watch((int(w.winId()), w.context_outline())
                                           for w in self.overlays.values() if w.isVisible())

    def _show_strip_menu(self, hwnd, x, y):
        if QApplication.activeModalWidget() is None:
            for overlay in self.overlays.values():
                if overlay.isVisible() and int(overlay.winId()) == hwnd:
                    dpr = overlay.devicePixelRatioF()
                    point = overlay.mapToGlobal(QPoint(round(x / dpr), round(y / dpr)))
                    self.float_bar.show_context_menu_at(point)
                    break

    def _sync_sampling(self):
        enabled = (sys.platform == "win32" and self.appearance.smart_fade and self.overlays
                   and getattr(self.snapshot, "hwnd", 0))
        if enabled and not self._sample_timer.isActive():
            self._sample_timer.start()
        elif not enabled:
            self._sample_timer.stop()
            for overlay in self.overlays.values():
                overlay.set_faded(False)

    def _sample_background(self):
        source = getattr(self.snapshot, "hwnd", 0)
        if not source:
            return
        # Let slow/off-screen sources time out without blocking the presenter or
        # accumulating bitmap jobs when the operator changes slides quickly.
        frames = {int(w.winId()): (w, w.geometry().getRect()) for w in self.overlays.values() if w.isVisible()}
        frame_key = (source, getattr(self.snapshot, "current", None),
                     tuple((hwnd, rect) for hwnd, (_, rect) in frames.items()))
        try:
            key, started, scores = self._samples.get_nowait()
            if key == frame_key and time.monotonic() - started < 2.5:
                self._last_sample = started
                for hwnd, score in scores.items():
                    overlay = frames[hwnd][0]
                    # Different entry/exit thresholds prevent flicker on gradual transitions.
                    threshold = .014 if overlay._faded else .024
                    overlay.set_faded(score is not None and score >= threshold)
        except Empty:
            pass
        if time.monotonic() - self._last_sample > 2.5:
            for overlay, _ in frames.values():
                overlay.set_faded(False)
        if self._sample_worker is not None and self._sample_worker.is_alive():
            return
        if not frames:
            return
        def collect(queue, key, handles):
            started = time.monotonic()
            scores = {}
            for hwnd in handles:
                try:
                    scores[hwnd] = sample_underlay(source, hwnd)
                except Exception:
                    scores[hwnd] = None  # A window can disappear between two native calls.
            queue.put((key, started, scores))
        self._sample_worker = threading.Thread(target=collect, args=(self._samples, frame_key, tuple(frames)),
                                               name="Slide detail sampler", daemon=True)
        self._sample_worker.start()

    def close(self):
        if self._projection_context is not None:
            self._projection_context.watch(())
        self._flash_timer.stop()
        self._sample_timer.stop()
        self._preview_timer.stop()
        for overlay in self.overlays.values():
            overlay.close()
        self.overlays.clear()
