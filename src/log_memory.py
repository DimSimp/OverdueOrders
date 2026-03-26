from __future__ import annotations

import logging
from collections import deque

_handler: _MemoryLogHandler | None = None


class _MemoryLogHandler(logging.Handler):
    def __init__(self, maxlen: int = 5000):
        super().__init__()
        self._lines: deque[str] = deque(maxlen=maxlen)

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self._lines.append(self.format(record))
        except Exception:
            self.handleError(record)

    def get_lines(self) -> list[str]:
        return list(self._lines)


def install() -> None:
    global _handler
    if _handler is not None:
        return
    _handler = _MemoryLogHandler()
    _handler.setFormatter(logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%H:%M:%S",
    ))
    logging.getLogger().addHandler(_handler)


def get_lines() -> list[str]:
    if _handler is None:
        return []
    return _handler.get_lines()
