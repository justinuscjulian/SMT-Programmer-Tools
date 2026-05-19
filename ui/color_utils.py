import re

from PySide6.QtGui import QColor


RGBA_RE = re.compile(r"rgba\((\d+),\s*(\d+),\s*(\d+),\s*(\d+)\)", re.IGNORECASE)


def theme_color(value, fallback="#ffffff"):
    if not value:
        return QColor(fallback)

    match = RGBA_RE.fullmatch(str(value).strip())
    if match:
        red, green, blue, alpha = (int(part) for part in match.groups())
        return QColor(red, green, blue, alpha)

    color = QColor(str(value))
    return color if color.isValid() else QColor(fallback)


def with_alpha(color, alpha):
    result = QColor(color)
    result.setAlpha(max(0, min(255, int(alpha))))
    return result


def mix_color(left, right, amount):
    amount = max(0.0, min(1.0, float(amount)))
    return QColor(
        round(left.red() + (right.red() - left.red()) * amount),
        round(left.green() + (right.green() - left.green()) * amount),
        round(left.blue() + (right.blue() - left.blue()) * amount),
        round(left.alpha() + (right.alpha() - left.alpha()) * amount),
    )
