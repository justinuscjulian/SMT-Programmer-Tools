from PySide6.QtCore import Property, QEasingCurve, QEvent, QPropertyAnimation, QRectF, Qt
from PySide6.QtGui import QLinearGradient, QPainter, QPen
from PySide6.QtWidgets import QStyle, QStyleOptionToolButton, QStylePainter, QToolButton

from ui.color_utils import mix_color, theme_color, with_alpha


class LiquidNavButton(QToolButton):
    def __init__(self, theme_manager, parent=None):
        super().__init__(parent)
        self._theme = theme_manager.theme
        self._hover_progress = 0.0
        self._active_progress = 0.0
        self.setAttribute(Qt.WA_Hover, True)
        self.setMouseTracking(True)

        self._hover_animation = self._build_animation(b"hoverProgress", 180)
        self._active_animation = self._build_animation(b"activeProgress", 240)
        theme_manager.changed.connect(self.set_theme)

    def _build_animation(self, prop, duration):
        animation = QPropertyAnimation(self, prop, self)
        animation.setDuration(duration)
        animation.setEasingCurve(QEasingCurve.OutCubic)
        return animation

    def set_theme(self, theme):
        self._theme = theme
        self.update()

    def set_active(self, active):
        self._animate(self._active_animation, 1.0 if active else 0.0)

    def event(self, event):
        if event.type() == QEvent.HoverEnter:
            self._animate(self._hover_animation, 1.0)
        elif event.type() == QEvent.HoverLeave:
            self._animate(self._hover_animation, 0.0)
        return super().event(event)

    def _animate(self, animation, target):
        animation.stop()
        animation.setStartValue(animation.propertyName() and self._property_value(animation.propertyName()))
        animation.setEndValue(target)
        animation.start()

    def _property_value(self, prop):
        if bytes(prop) == b"hoverProgress":
            return self._hover_progress
        return self._active_progress

    def paintEvent(self, event):
        del event
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        self._paint_liquid_background(painter)
        painter.end()

        style_painter = QStylePainter(self)
        option = QStyleOptionToolButton()
        self.initStyleOption(option)
        style_painter.drawControl(QStyle.CE_ToolButtonLabel, option)

    def _paint_liquid_background(self, painter):
        theme = self._theme
        rect = QRectF(self.rect()).adjusted(1.0, 1.0, -1.0, -1.0)
        radius = min(rect.height() / 2.2, 18)
        hover = self._hover_progress
        active = self._active_progress
        glow = max(hover * 0.72, active)

        glass = theme_color(theme.get("glass_strong"), "#ffffff")
        surface = theme_color(theme.get("surface_high"), "#ffffff")
        border = theme_color(theme.get("border_soft"), "#ffffff")
        accent = theme_color(theme.get("accent"), "#0a84ff")

        top_color = mix_color(surface, accent, 0.08 + active * 0.16)
        bottom_color = mix_color(glass, accent, hover * 0.12 + active * 0.2)
        top_color = with_alpha(top_color, 64 + hover * 46 + active * 72)
        bottom_color = with_alpha(bottom_color, 42 + hover * 58 + active * 82)

        fill = QLinearGradient(rect.topLeft(), rect.bottomRight())
        fill.setColorAt(0.0, top_color)
        fill.setColorAt(0.52, with_alpha(glass, 32 + glow * 82))
        fill.setColorAt(1.0, bottom_color)
        painter.setPen(Qt.NoPen)
        painter.setBrush(fill)
        painter.drawRoundedRect(rect, radius, radius)

        highlight = QLinearGradient(rect.topLeft(), rect.bottomLeft())
        highlight.setColorAt(0.0, with_alpha(theme_color("#ffffff"), 54 + hover * 42))
        highlight.setColorAt(0.45, with_alpha(theme_color("#ffffff"), 12 + active * 18))
        highlight.setColorAt(1.0, with_alpha(theme_color("#ffffff"), 0))
        painter.setBrush(highlight)
        painter.drawRoundedRect(rect.adjusted(1.2, 1.2, -1.2, -rect.height() * 0.42), radius, radius)

        pen_color = mix_color(border, accent, hover * 0.35 + active * 0.5)
        painter.setBrush(Qt.NoBrush)
        painter.setPen(QPen(with_alpha(pen_color, 74 + glow * 82), 1.15))
        painter.drawRoundedRect(rect, radius, radius)

    def get_hover_progress(self):
        return self._hover_progress

    def set_hover_progress(self, value):
        self._hover_progress = float(value)
        self.update()

    def get_active_progress(self):
        return self._active_progress

    def set_active_progress(self, value):
        self._active_progress = float(value)
        self.update()

    hoverProgress = Property(float, get_hover_progress, set_hover_progress)
    activeProgress = Property(float, get_active_progress, set_active_progress)
