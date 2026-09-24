"""Compatibility entry point for daily tasks installed before folder cleanup."""
from pathlib import Path
import runpy
import sys

if __name__ == '__main__':
    program = Path(__file__).resolve().parent / 'App'
    sys.path.insert(0, str(program))
    runpy.run_path(str(program / 'scheduled_run.py'), run_name='__main__')
