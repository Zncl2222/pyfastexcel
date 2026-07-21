from __future__ import annotations

import pytest

from pyfastexcel.manager import StyleManager


@pytest.fixture(autouse=True)
def isolate_process_style_defaults():
    """Production saves retain defaults; tests isolate that process-level state."""
    StyleManager.reset_style_configs()
    yield
    StyleManager.reset_style_configs()
