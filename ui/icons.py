from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPainterPath, QPen, QPixmap


def make_icon(name, color="#2563eb", size=24):
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor(color), 2.1)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)

    rect = QRectF(4, 4, size - 8, size - 8)
    if name == "menu":
        for y in (7, 12, 17):
            painter.drawLine(QPointF(5, y), QPointF(19, y))
    elif name == "compare":
        painter.drawRoundedRect(QRectF(4, 5, 7, 14), 2.5, 2.5)
        painter.drawRoundedRect(QRectF(13, 5, 7, 14), 2.5, 2.5)
        painter.drawLine(QPointF(8, 9), QPointF(16, 9))
        painter.drawLine(QPointF(16, 15), QPointF(8, 15))
    elif name == "machine":
        painter.drawRoundedRect(QRectF(6, 6, 12, 12), 3, 3)
        for x in (8, 12, 16):
            painter.drawLine(QPointF(x, 3), QPointF(x, 6))
            painter.drawLine(QPointF(x, 18), QPointF(x, 21))
        for y in (8, 12, 16):
            painter.drawLine(QPointF(3, y), QPointF(6, y))
            painter.drawLine(QPointF(18, y), QPointF(21, y))
        painter.drawEllipse(QPointF(12, 12), 2.5, 2.5)
    elif name == "history":
        painter.drawEllipse(rect)
        painter.drawLine(QPointF(12, 12), QPointF(12, 7))
        painter.drawLine(QPointF(12, 12), QPointF(16, 14))
    elif name == "tools":
        path = QPainterPath()
        path.moveTo(6, 18)
        path.lineTo(14.5, 9.5)
        path.moveTo(14.2, 5.4)
        path.cubicTo(16, 4, 18.6, 4.1, 20, 5.6)
        path.lineTo(16.6, 9)
        path.lineTo(15, 7.4)
        path.lineTo(11, 11.4)
        painter.drawPath(path)
        painter.drawEllipse(QPointF(6, 18), 2.1, 2.1)
    elif name == "worksheet":
        painter.drawRoundedRect(QRectF(5, 4, 14, 16), 2.5, 2.5)
        painter.drawLine(QPointF(8, 8), QPointF(16, 8))
        painter.drawLine(QPointF(8, 12), QPointF(16, 12))
        painter.drawLine(QPointF(8, 16), QPointF(13, 16))
    elif name == "all_in_one":
        painter.drawRoundedRect(QRectF(4, 4, 7, 7), 2, 2)
        painter.drawRoundedRect(QRectF(13, 4, 7, 7), 2, 2)
        painter.drawRoundedRect(QRectF(4, 13, 7, 7), 2, 2)
        painter.drawRoundedRect(QRectF(13, 13, 7, 7), 2, 2)
        painter.drawLine(QPointF(11, 7.5), QPointF(13, 7.5))
        painter.drawLine(QPointF(7.5, 11), QPointF(7.5, 13))
        painter.drawLine(QPointF(11, 16.5), QPointF(13, 16.5))
    elif name == "arrow_left":
        painter.drawLine(QPointF(15.5, 6), QPointF(9.5, 12))
        painter.drawLine(QPointF(9.5, 12), QPointF(15.5, 18))
        painter.drawLine(QPointF(10, 12), QPointF(20, 12))
    elif name == "arrow_right":
        painter.drawLine(QPointF(8.5, 6), QPointF(14.5, 12))
        painter.drawLine(QPointF(14.5, 12), QPointF(8.5, 18))
        painter.drawLine(QPointF(4, 12), QPointF(14, 12))
    elif name == "sun":
        painter.drawEllipse(QPointF(12, 12), 4, 4)
        for a, b in [
            ((12, 3), (12, 5)),
            ((12, 19), (12, 21)),
            ((3, 12), (5, 12)),
            ((19, 12), (21, 12)),
            ((5.6, 5.6), (7.1, 7.1)),
            ((16.9, 16.9), (18.4, 18.4)),
            ((18.4, 5.6), (16.9, 7.1)),
            ((7.1, 16.9), (5.6, 18.4)),
        ]:
            painter.drawLine(QPointF(*a), QPointF(*b))
    elif name == "moon":
        path = QPainterPath()
        path.moveTo(18, 16.4)
        path.cubicTo(13, 18.4, 7.8, 15.1, 7.8, 9.8)
        path.cubicTo(7.8, 7, 9.2, 4.8, 11.3, 3.5)
        path.cubicTo(9.8, 8.4, 13.4, 13.2, 18.6, 12.7)
        path.cubicTo(18.6, 14.1, 18.4, 15.3, 18, 16.4)
        painter.drawPath(path)
    else:
        painter.drawEllipse(rect)

    painter.end()
    return QIcon(pixmap)
