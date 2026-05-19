from pathlib import Path


ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")


def read_lines_with_fallback(path):
    path = Path(path)
    last_error = None
    for encoding in ENCODINGS:
        try:
            with path.open("r", encoding=encoding) as handle:
                return handle.readlines(), encoding
        except UnicodeDecodeError as exc:
            last_error = exc
    if last_error:
        raise last_error
    raise UnicodeDecodeError("unknown", b"", 0, 1, "Unable to read file")

