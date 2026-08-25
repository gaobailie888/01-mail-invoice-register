#!/usr/bin/env python3
"""Install missing dependencies, then start the local setup wizard."""
from __future__ import annotations
import importlib.util
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
missing = [name for name in ("pdfplumber", "openpyxl") if importlib.util.find_spec(name) is None]
if missing:
    print("首次运行：正在安装必要依赖…")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(ROOT / "requirements.txt")])
subprocess.check_call([sys.executable, str(ROOT / "scripts" / "setup_wizard.py")])
