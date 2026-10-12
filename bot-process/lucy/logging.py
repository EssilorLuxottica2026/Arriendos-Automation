"""Logging helpers for Bot Lucy progress reporting."""
import sys
from typing import Callable, Optional


_callback: Optional[Callable[[str], None]] = None


def set_log_callback(callback: Optional[Callable[[str], None]]) -> None:
    global _callback
    _callback = callback


def log(message: str) -> None:
    if sys.stdout is not None:
        print(message)
    if _callback is not None:
        _callback(message)
