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
    elif name == "plan":
        painter.drawRoundedRect(QRectF(5, 3.5, 14, 17), 2.5, 2.5)
        painter.drawLine(QPointF(8, 8), QPointF(16, 8))
        painter.drawLine(QPointF(8, 12), QPointF(16, 12))
        painter.drawLine(QPointF(8, 16), QPointF(12, 16))
        painter.drawLine(QPointF(14, 16), QPointF(17.5, 19.5))
    elif name == "all_in_one":
        painter.drawRoundedRect(QRectF(4, 4, 7, 7), 2, 2)
        painter.drawRoundedRect(QRectF(13, 4, 7, 7), 2, 2)
        painter.drawRoundedRect(QRectF(4, 13, 7, 7), 2, 2)
        painter.drawRoundedRect(QRectF(13, 13, 7, 7), 2, 2)
        painter.drawLine(QPointF(11, 7.5), QPointF(13, 7.5))
        painter.drawLine(QPointF(7.5, 11), QPointF(7.5, 13))
        painter.drawLine(QPointF(11, 16.5), QPointF(13, 16.5))
    elif name == "insert_point":
        painter.drawEllipse(QPointF(12, 12), 7, 7)
        painter.drawEllipse(QPointF(12, 12), 2.4, 2.4)
        painter.drawLine(QPointF(12, 3), QPointF(12, 7))
        painter.drawLine(QPointF(12, 17), QPointF(12, 21))
        painter.drawLine(QPointF(3, 12), QPointF(7, 12))
        painter.drawLine(QPointF(17, 12), QPointF(21, 12))
    elif name == "feeder_mapping":
        painter.drawRoundedRect(QRectF(4, 5, 16, 14), 2.5, 2.5)
        for y in (8.5, 12, 15.5):
            painter.drawLine(QPointF(7, y), QPointF(13.5, y))
            painter.drawLine(QPointF(16.5, y), QPointF(17.5, y))
        painter.drawLine(QPointF(6, 20.5), QPointF(18, 20.5))
    elif name == "feeder_compare":
        painter.drawRoundedRect(QRectF(4, 5, 6, 14), 2.2, 2.2)
        painter.drawRoundedRect(QRectF(14, 5, 6, 14), 2.2, 2.2)
        painter.drawLine(QPointF(7, 9), QPointF(17, 9))
        painter.drawLine(QPointF(17, 15), QPointF(7, 15))
        painter.drawLine(QPointF(15.3, 7.3), QPointF(17.2, 9))
        painter.drawLine(QPointF(15.3, 10.7), QPointF(17.2, 9))
        painter.drawLine(QPointF(8.7, 13.3), QPointF(6.8, 15))
        painter.drawLine(QPointF(8.7, 16.7), QPointF(6.8, 15))
    elif name == "feeder_reuse":
        painter.drawRoundedRect(QRectF(4, 5, 7, 14), 2.2, 2.2)
        painter.drawRoundedRect(QRectF(13, 5, 7, 14), 2.2, 2.2)
        painter.drawLine(QPointF(7.5, 9), QPointF(16.5, 9))
        painter.drawLine(QPointF(16.5, 15), QPointF(7.5, 15))
        painter.drawLine(QPointF(14.8, 7.4), QPointF(16.8, 9))
        painter.drawLine(QPointF(14.8, 10.6), QPointF(16.8, 9))
        painter.drawLine(QPointF(9.2, 13.4), QPointF(7.2, 15))
        painter.drawLine(QPointF(9.2, 16.6), QPointF(7.2, 15))
    elif name == "model_group":
        painter.drawRoundedRect(QRectF(4, 5, 6, 6), 2, 2)
        painter.drawRoundedRect(QRectF(14, 5, 6, 6), 2, 2)
        painter.drawRoundedRect(QRectF(9, 14, 6, 6), 2, 2)
        painter.drawLine(QPointF(10, 8), QPointF(14, 8))
        painter.drawLine(QPointF(8.5, 11), QPointF(10.5, 14))
        painter.drawLine(QPointF(15.5, 11), QPointF(13.5, 14))
        painter.drawEllipse(QPointF(12, 12), 1.4, 1.4)
    elif name == "used_part_component":
        painter.drawRoundedRect(QRectF(4, 4, 16, 16), 2.5, 2.5)
        painter.drawLine(QPointF(4, 9), QPointF(20, 9))
        painter.drawLine(QPointF(4, 14), QPointF(20, 14))
        painter.drawLine(QPointF(9, 4), QPointF(9, 20))
        painter.drawLine(QPointF(14, 4), QPointF(14, 20))
        painter.drawEllipse(QPointF(16.8, 16.8), 1.3, 1.3)
    elif name == "component_usage":
        painter.drawRoundedRect(QRectF(4, 4, 12, 14), 2.5, 2.5)
        painter.drawLine(QPointF(7, 8), QPointF(13, 8))
        painter.drawLine(QPointF(7, 11.5), QPointF(13, 11.5))
        painter.drawLine(QPointF(7, 15), QPointF(11, 15))
        painter.drawEllipse(QPointF(15.5, 15.5), 3.6, 3.6)
        painter.drawLine(QPointF(18, 18), QPointF(21, 21))
    elif name == "help":
        painter.drawEllipse(QPointF(12, 12), 8, 8)
        path = QPainterPath()
        path.moveTo(9, 9.2)
        path.cubicTo(9.4, 7.4, 11, 6.5, 12.7, 6.7)
        path.cubicTo(14.5, 6.9, 15.6, 8.1, 15.4, 9.7)
        path.cubicTo(15.2, 11.1, 13.9, 11.8, 12.9, 12.5)
        path.cubicTo(12.2, 13, 12, 13.4, 12, 14.2)
        painter.drawPath(path)
        painter.drawPoint(QPointF(12, 17.2))
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
