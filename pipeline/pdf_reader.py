from pathlib import Path
import sys


BASE_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from lease_accounting.pipeline.pdf_reader import *  # noqa: F401,F403

