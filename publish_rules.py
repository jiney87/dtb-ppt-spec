#!/usr/bin/env python3
"""Validate the complete release before pushing main."""
from pathlib import Path
import subprocess
import sys

root = Path(__file__).resolve().parent
subprocess.run([sys.executable, str(root / "scripts/validate_release.py")], check=True)
subprocess.run(["git", "push", "origin", "main"], cwd=root, check=True)
