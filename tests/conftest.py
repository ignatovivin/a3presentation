from __future__ import annotations

import os
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEST_TEMPLATES_DIR = PROJECT_ROOT / "tests" / "fixtures" / "templates"

os.environ.setdefault("TEMPLATES_DIR", str(TEST_TEMPLATES_DIR))
