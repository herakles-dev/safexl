"""Shared helper for tests that need a third-party workbook.

The workbooks are not stored in the repository; ``fetch_corpus.py`` downloads
them from their pinned upstream sources. A test that needs one skips when it
is absent. Not collected by pytest (no ``test_`` name).
"""
from __future__ import annotations

from pathlib import Path

import pytest

SKIP_REASON = "run research/xlsx-bench/fetch_corpus.py"


def require(path):
    """Return ``path`` unchanged, or skip the calling test if the file is absent."""
    if not Path(path).is_file():
        pytest.skip(SKIP_REASON)
    return path
