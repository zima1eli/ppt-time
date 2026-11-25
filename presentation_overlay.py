"""A passive audience display for a running slide show."""

import math

from PySide6.QtCore import QPointF, QPropertyAnimation, QRect, QRectF, Qt
from PySide6.QtGui import (QBitmap, QColor, QFont, QFontMetricsF, QGuiApplication, QImage,
                           QPainter, QPainterPath, QPen, QRegion, QScreen, QTransform)
from PySide6.QtWidgets import QWidget

from overlay_options import OverlayOptions, capture_supported, set_capture_affinity


# a, b, c, d, e, f, g denote the top segment, clockwise perimeter, and middle.
_DIGITS = ("abcdef", "bc", "abdeg", "abcdg", "bcfg", "acdfg", "acdefg", "abc", "abcdefg", "abcdfg")
_LINES = dict(a=(.12,.08,.48,.08), b=(.54,.15,.54,.43), c=(.54,.57,.54,.85),
              d=(.12,.92,.48,.92), e=(.06,.57,.06,.85), f=(.06,.15,.06,.43), g=(.12,.5,.48,.5))


def _lcd_width(text, pixels):
    return sum(.29 if c == ":" else .66 for c in text) * pixels


def _paint_lcd(p, rect, text, pixels, color, bold):
    p.save()
    p.translate(rect.center().x() - _lcd_width(text, pixels)/2, rect.center().y()-pixels/2)
    p.scale(pixels, pixels)
    width = .115 if bold else .075
    p.setBrush(color)
    for char in text:
        p.setPen(QPen(color, width, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        if char.isascii() and char.isdecimal():
            for segment in _DIGITS[int(char)]:
                x1,y1,x2,y2 = _LINES[segment]
                p.drawLine(QPointF(x1,y1), QPointF(x2,y2))
        elif char == ":":
            p.drawPoint(QPointF(.14,.32))
            p.drawPoint(QPointF(.14,.70))
        elif char == "/":
            p.drawLine(QPointF(.12,.91), QPointF(.49,.08))
        elif char in "+-":
            p.drawLine(QPointF(.09,.5), QPointF(.52,.5))
            if char == "+":
                p.drawLine(QPointF(.305,.29), QPointF(.305,.71))
        p.translate(.29 if char == ":" else .66, 0)
    p.restore()


class PresentationOverlay(QWidget):
    HEIGHT = 36
    ICON = 15
    GAP = 6
    DIVIDER_SPACE = 21

    def __init__(self, capture_surface=False):
        flags = (Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint
                 | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.WindowDoesNotAcceptFocus
                 | Qt.WindowType.WindowTransparentForInput)
        super().__init__(None, flags)
        self.setWindowTitle("ppt time 放映浮层")
        # Windows 10 rejects display affinity on UpdateLayeredWindow surfaces.
        # A shaped, DWM-composited top-level window accepts affinity and keeps
        # the same native click-through / non-activation flags.
        self.capture_surface = capture_surface
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, not capture_surface)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_MacAlwaysShowToolWindow)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._target_screen = None
        self._screen_rect = QRect()
        self._style = "notch"
        self._scale = 1.0
        self._desktop_time_size = 32
        self._background = QColor("#141414")
        self._foreground = QColor("#ffffff")
        self._alert = QColor()
        self._time = "00:00"
        self._page = ""
        self.options = OverlayOptions()
        self._faded = False
        self._flash_dim = False
        self._affinity_key = None
        self.capture_excluded = False
        self._base_opacity = 1.
        self._fade_animation = QPropertyAnimation(self, b"windowOpacity", self)
        self._fade_animation.setDuration(180)
        self._time_font = self._font(20)
        self._page_font = self._font(16)
        self._measure()

    @staticmethod
    def _font(pixels, family="", bold=False):
        font = QFont()
        font.setFamilies(([family] if family else []) + ["Segoe UI", "Helvetica Neue", "Arial"])
        font.setWeight(QFont.Weight.Bold if bold else QFont.Weight.DemiBold)
        font.setPixelSize(pixels)
        return font

    def configure(self, style: str, scale: int, background: QColor, foreground: QColor,
                  opacity: float, options=None, desktop_time_size=32):
        self._style = style if style in ("notch", "pill", "shoulder") else "notch"
        self._scale = max(0.25, min(4.0, scale / 100.0))
        self._desktop_time_size = desktop_time_size
        self.options = options or OverlayOptions()
        self._background = QColor(self.options.background or background)
        self._background.setAlphaF(background.alphaF())
        self._foreground = QColor(self.options.foreground or foreground)
        self._foreground.setAlphaF(foreground.alphaF())
        self._base_opacity = max(0.0, min(1.0, opacity))
        self._fade_animation.stop()
        if not self.options.smart_fade:
            self._faded = False
        self.setWindowOpacity(self._target_opacity())
        self._time_font = self._font(self.options.time_size, self.options.font, self.options.bold)
        self._page_font = self._font(self.options.page_size, self.options.font, self.options.bold)
        if self.isVisible():
            self._apply_capture()
        self._measure()
        self.update()

    def _apply_capture(self):
        if not capture_supported() or QGuiApplication.platformName() != "windows":
            return
        key = int(self.winId()), self.options.exclude_capture
        if self._affinity_key != key:
            self._affinity_key = key
            self.capture_excluded = set_capture_affinity(*key) and key[1]

    def showEvent(self, event):
        super().showEvent(event)
        self._apply_capture()

    def set_flash_dim(self, dim):
        if self._flash_dim != dim:
            self._flash_dim = dim
            self.update()

    def _target_opacity(self):
        alpha = self._base_opacity
        if self.capture_surface:
            alpha *= self._background.alphaF() if self._background.alpha() else 1.
        return alpha * (self.options.fade_percent / 100. if self._faded else 1.)

    def set_faded(self, faded):
        if self._faded == faded:
            return
        self._faded = faded
        self._fade_animation.stop()
        self._fade_animation.setStartValue(self.windowOpacity())
        self._fade_animation.setEndValue(self._target_opacity())
        self._fade_animation.start()

    def set_content(self, time_text: str, page_text: str = "", alert_color: str = ""):
        alert = QColor(alert_color) if alert_color else QColor()
        if (time_text, page_text, alert) == (self._time, self._page, self._alert):
            return
        new_text = (time_text, page_text) != (self._time, self._page)
        self._time, self._page, self._alert = time_text, page_text, alert
        if new_text:
            self._measure()
        self.update()

    def place_on(self, screen: QScreen):
        if screen is None:
            return
        if self._target_screen is not screen:
            if self._target_screen is not None:
                self._target_screen.geometryChanged.disconnect(self._set_screen_rect)
                self._target_screen.destroyed.disconnect(self._screen_removed)
            self._target_screen = screen
            screen.geometryChanged.connect(self._set_screen_rect)
            screen.destroyed.connect(self._screen_removed)
            self.setScreen(screen)
        self._set_screen_rect(screen.geometry())

    def _screen_removed(self):
        self._target_screen = None
        self._screen_rect = QRect()

    def _set_screen_rect(self, rect):
        if rect != self._screen_rect:
            self._screen_rect = rect
            self._fit_to_screen()
        else:
            self._reposition()

    def _reposition(self):
        if self._screen_rect.isEmpty():
            return
        free_x = max(0, self._screen_rect.width() - self.width())
        free_y = max(0, self._screen_rect.height() - self.height())
        side, alignment = self.options.anchor.split("-")
        x = {"left": 0, "center": free_x // 2, "right": free_x}[alignment]
        secondary = self._target_screen is not None and self._target_screen != QGuiApplication.primaryScreen()
        x_offset, gap = self.options.offsets(secondary)
        x = self._screen_rect.x() + max(0, min(free_x, x + x_offset))
        if gap < 0:
            gap = round(8 * self._paint_scale) if self._style == "pill" else 0
        y = self._screen_rect.y() + (min(free_y, gap) if side == "top" else max(0, free_y - gap))
        if x != self.x() or y != self.y():
            self.move(x, y)

    def _text_width(self, text, font):
        if self.options.lcd and all(c in "0123456789:/+-" for c in text):
            return math.ceil(_lcd_width(text, font.pixelSize()))
        metrics = QFontMetricsF(font)
        digit_width = max(metrics.horizontalAdvance(str(n)) for n in range(10))
        return math.ceil(max(metrics.horizontalAdvance(text),
                             sum(digit_width if c.isdecimal() else metrics.horizontalAdvance(c) for c in text)))

    def _draw_digits(self, painter, x, width, text, font, color):
        rect = QRectF(x, 0, width, self._logical_height)
        if self.options.lcd and all(c in "0123456789:/+-" for c in text):
            _paint_lcd(painter, rect, text, font.pixelSize(), color, self.options.bold)
        else:
            painter.setPen(color)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, text)

    def _measure(self):
        self._logical_height = max(self.HEIGHT, math.ceil(max(QFontMetricsF(self._time_font).height(),
                                                           QFontMetricsF(self._page_font).height()) + 6))
        self._padding = 21 if self._style == "shoulder" else 14
        width = self._padding * 2
        self._time_width = self._text_width(self._time, self._time_font) if self._time else 0
        if self._time:
            width += self.ICON + self.GAP + self._time_width
        if self._page:
            # Reserve the numerator's eventual full width while turning through a deck.
            num, slash, total = self._page.partition("/")
            sample = num.rjust(len(total), "0") + slash + total if num.isdecimal() and total.isdecimal() else self._page
            self._page_width = self._text_width(sample, self._page_font)
            width += self.ICON + self.GAP + self._page_width
        if self._time and self._page:
            width += self.DIVIDER_SPACE
        self._logical_width = width
        self._fit_to_screen()

    def _fit_to_screen(self):
        if self.options.size_mode == "auto":
            # QScreen geometry is already in logical pixels. Matching the two
            # clock sizes here lets Qt handle Windows DPI exactly once.
            short_edge = min(self._screen_rect.width(), self._screen_rect.height())
            screen_factor = max(0.875, min(2.0, short_edge / 1080.0)) if short_edge > 0 else 1.0
            size_factor = max(1., self._desktop_time_size / OverlayOptions.time_size) * screen_factor
        else:
            size_factor = int(self.options.size_mode) / 100.0
        self._paint_scale = self._scale * size_factor
        if not self._screen_rect.isEmpty():
            self._paint_scale = min(self._paint_scale, self._screen_rect.width() / self._logical_width,
                                    self._screen_rect.height() / self._logical_height)
        new_width = math.ceil(self._logical_width * self._paint_scale)
        new_height = math.ceil(self._logical_height * self._paint_scale)
        if not self._screen_rect.isEmpty():
            new_width = min(new_width, self._screen_rect.width())
            new_height = min(new_height, self._screen_rect.height())
        if (new_width, new_height) != (self.width(), self.height()):
            self.setFixedSize(new_width, new_height)
        self._reposition()
        if self.capture_surface:
            self._update_capture_shape()

    def _update_capture_shape(self):
        if self._background.alpha():
            shape = QTransform.fromScale(self._paint_scale, self._paint_scale).map(self._outline())
            region = QRegion(shape.toFillPolygon().toPolygon())
        else:
            # A fully transparent background must leave the digits visible.
            ink = QImage(self.size(), QImage.Format.Format_ARGB32_Premultiplied)
            ink.fill(Qt.GlobalColor.transparent)
            painter = QPainter(ink)
            self._paint(painter)
            painter.end()
            region = QRegion(QBitmap.fromImage(ink.createAlphaMask()))
        self.setMask(region if not region.isEmpty() else QRegion(-2, -2, 1, 1))

    def _outline(self):
        w, h = self._logical_width, self._logical_height
        path = QPainterPath()
        if self._style == "pill":
            r = h / 2 if self.options.radius < 0 else min(self.options.radius, h / 2, w / 2)
            path.addRoundedRect(QRectF(0, 0, w, h), r, r)
            return path
        shoulder = 9 if self._style == "shoulder" else 0
        radius = self.options.radius if self.options.radius >= 0 else (15 if shoulder else 9)
        radius = min(radius, h - shoulder, max(0, w / 2 - shoulder))
        path.moveTo(0, 0)
        path.lineTo(w, 0)
        if shoulder:
            path.quadTo(w - shoulder, 0, w - shoulder, shoulder)
        path.lineTo(w - shoulder, h - radius)
        path.quadTo(w - shoulder, h, w - shoulder - radius, h)
        path.lineTo(shoulder + radius, h)
        path.quadTo(shoulder, h, shoulder, h - radius)
        if shoulder:
            path.lineTo(shoulder, shoulder)
            path.quadTo(shoulder, 0, 0, 0)
        path.closeSubpath()
        if self.options.anchor.startswith("bottom"):
            return QTransform(1, 0, 0, -1, 0, h).map(path)
        return path

    def context_outline(self):
        return QTransform.fromScale(1 / self._logical_width, 1 / self._logical_height).map(self._outline())

    def _glyph(self, painter, x, color, clock=False):
        painter.save()
        painter.translate(x, (self._logical_height - self.ICON) / 2)
        painter.setPen(QPen(color, 1.45, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap,
                            Qt.PenJoinStyle.RoundJoin))
        if clock:
            painter.drawEllipse(QRectF(1, 1, 13, 13))
            path = QPainterPath(QPointF(7.5, 3.8))
            path.lineTo(7.5, 7.6)
            path.lineTo(10.4, 9.1)
        else:
            path = QPainterPath(QPointF(3, 1))
            path.lineTo(9, 1)
            path.lineTo(12.5, 4.6)
            path.lineTo(12.5, 14)
            path.lineTo(3, 14)
            path.closeSubpath()
            path.moveTo(9, 1.3)
            path.lineTo(9, 5)
            path.lineTo(12, 5)
            path.moveTo(5.8, 9.3)
            path.lineTo(9.6, 9.3)
        painter.drawPath(path)
        painter.restore()

    def paintEvent(self, event):
        painter = QPainter(self)
        self._paint(painter)
        painter.end()

    def _paint(self, painter):
        painter.setRenderHints(QPainter.RenderHint.Antialiasing | QPainter.RenderHint.TextAntialiasing)
        painter.scale(self._paint_scale, self._paint_scale)
        background = QColor(self._background)
        if self.capture_surface and background.alpha():
            background.setAlpha(255)
        painter.fillPath(self._outline(), background)
        x = self._padding
        if self._time:
            time_color = QColor(self._alert if self._alert.isValid() else self._foreground)
            if self._flash_dim:
                time_color.setAlphaF(time_color.alphaF() * .36)
            self._glyph(painter, x, time_color, clock=True)
            x += self.ICON + self.GAP
            self._draw_digits(painter, x, self._time_width, self._time, self._time_font, time_color)
            x += self._time_width
        if self._time and self._page:
            divider = QColor(self._foreground)
            divider.setAlphaF(divider.alphaF() * 0.38)
            painter.setPen(QPen(divider, 1))
            centre = x + self.DIVIDER_SPACE / 2
            painter.drawLine(QPointF(centre, 10), QPointF(centre, self._logical_height - 10))
            x += self.DIVIDER_SPACE
        if self._page:
            self._glyph(painter, x, self._foreground)
            x += self.ICON + self.GAP
            self._draw_digits(painter, x, self._page_width, self._page, self._page_font, self._foreground)
