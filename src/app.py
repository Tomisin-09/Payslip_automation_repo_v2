# src/app.py
from __future__ import annotations

import os
import sys

def _resource_path(relative_path: str) -> str:
    """
    When packaged with PyInstaller, files are unpacked to sys._MEIPASS.
    In normal runs, use project root.
    """
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)

def main() -> None:
    # Optionally set cwd to a stable location (e.g., the folder where the exe lives)
    # This helps non-technical users who double click the exe.
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd()
    os.chdir(exe_dir)

    # If you load config via relative path, ensure it points correctly.
    # Example: config/settings.yml next to the exe if you ship it that way.
    # Otherwise keep your existing path logic.
    from main import main as run_main  # assumes main.py has a main() function
    run_main()

if __name__ == "__main__":
    main()
