#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CSMP Assistant v8.0 — Entry point."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from gui.app import RUSApp

if __name__ == "__main__":
    RUSApp().mainloop()
