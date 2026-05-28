"""Structured logging for harness operations."""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

_logger: logging.Logger | None = None


def get_logger(name: str = "harness_pptx", log_dir: str | None = None) -> logging.Logger:
    global _logger
    if _logger is not None:
        return _logger

    _logger = logging.getLogger(name)
    _logger.setLevel(logging.DEBUG)
    _logger.handlers.clear()

    # Console handler
    ch = logging.StreamHandler(sys.stderr)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter(
        "[%(asctime)s] %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    ))
    _logger.addHandler(ch)

    # File handler (if log_dir provided)
    if log_dir:
        path = Path(log_dir)
        path.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(
            path / f"build_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        )
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter(
            "[%(asctime)s] %(levelname)s %(name)s: %(message)s"
        ))
        _logger.addHandler(fh)

    return _logger
