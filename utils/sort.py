import re


def natural_sort_key(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r"(\d+)", str(text))]

