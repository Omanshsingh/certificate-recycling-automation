"""Separate installed program resources from the operator's unchanged data."""
from pathlib import Path

PROGRAM_ROOT = Path(__file__).resolve().parent
DATA_ROOT = PROGRAM_ROOT.parent if PROGRAM_ROOT.name == 'App' else PROGRAM_ROOT


def resource(root, name):
    root = Path(root)
    direct = root / name
    return direct if direct.exists() else root / 'App' / name
