# Excel Diff Application
# This file makes the directory a Python package

from .ui.main_window import run_gui
from .core.differ import diff_excel

__version__ = "1.0.0"
__author__ = "Excel Diff Team"

__all__ = ["run_gui", "diff_excel"]