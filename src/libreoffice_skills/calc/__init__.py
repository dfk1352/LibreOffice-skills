"""Calc skill package."""

from libreoffice_skills.calc.cells import get_cell, set_cell
from libreoffice_skills.calc.core import create_spreadsheet
from libreoffice_skills.calc.snapshot import snapshot_area

__all__ = ["create_spreadsheet", "get_cell", "set_cell", "snapshot_area"]
