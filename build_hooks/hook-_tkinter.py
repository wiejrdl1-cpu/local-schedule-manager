"""Collect Tcl/Tk libraries without running PyInstaller's Tcl probe."""

from pathlib import Path
import sys


_tcl_root = Path(sys.base_prefix) / "tcl"
datas = [
    (str(_tcl_root / "tcl8.6"), "_tcl_data"),
    (str(_tcl_root / "tk8.6"), "_tk_data"),
]
