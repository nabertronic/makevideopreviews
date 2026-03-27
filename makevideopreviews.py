#!/usr/bin/env python3
from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parent
SRC_ROOT = PROJECT_ROOT / "src"

if SRC_ROOT.exists():
    sys.path.insert(0, str(SRC_ROOT))

from makevideopreviews.app import run


if __name__ == "__main__":
    raise SystemExit(run())
