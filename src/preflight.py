# src/preflight.py
from __future__ import annotations

import importlib
import platform
import sys
from typing import Iterable, List, Tuple


def is_windows() -> bool:
    return platform.system().lower() == "windows"


def check_python(min_major: int = 3, min_minor: int = 9) -> None:
    v = sys.version_info
    if (v.major, v.minor) < (min_major, min_minor):
        raise SystemExit(
            "\n".join(
                [
                    "Unsupported Python version.",
                    f"Found: {v.major}.{v.minor}.{v.micro}",
                    f"Required: {min_major}.{min_minor}+",
                ]
            )
        )


def check_required_modules(required: Iterable[Tuple[str, str]]) -> None:
    """
    required: iterable of (import_name, pip_name)
    Example: ("yaml", "PyYAML")
    """
    missing: List[Tuple[str, str]] = []
    for import_name, pip_name in required:
        try:
            importlib.import_module(import_name)
        except Exception:
            missing.append((import_name, pip_name))

    if missing:
        lines = [
            "Missing Python dependencies.",
            "",
            "Install requirements and try again:",
        ]
        if is_windows():
            lines.append("  py -m pip install -r requirements.txt")
        else:
            lines.append("  python3 -m pip install -r requirements.txt")

        lines.append("")
        lines.append("Missing modules:")
        for import_name, pip_name in missing:
            lines.append(f"  - {import_name} (pip package: {pip_name})")

        raise SystemExit("\n".join(lines))


def resolve_capabilities(
    pdf_enabled: bool,
    email_enabled: bool,
    *,
    allow_dev_fallback: bool = True,
) -> Tuple[bool, bool]:
    """
    On non-Windows, optionally downgrade PDF/email to False so macOS/Linux can run
    the XLSX-only workflow without failing.

    Returns: (pdf_enabled_resolved, email_enabled_resolved)
    """
    if is_windows():
        return pdf_enabled, email_enabled

    if not allow_dev_fallback:
        return pdf_enabled, email_enabled

    if pdf_enabled or email_enabled:
        print("Running on a non-Windows OS.")
        print("PDF export and email sending are Windows-only and will be skipped.")
        print("XLSX generation will still run.")
        print("")
    return False, False


def check_windows_com_deps(pdf_enabled: bool, email_enabled: bool) -> None:
    """
    If Windows and PDF/email enabled, ensure pywin32 is installed.
    """
    if not is_windows():
        return
    if not (pdf_enabled or email_enabled):
        return

    try:
        importlib.import_module("win32com.client")
    except Exception:
        raise SystemExit(
            "\n".join(
                [
                    "Missing Windows COM dependency: pywin32",
                    "",
                    "Install it and try again:",
                    "  py -m pip install pywin32",
                ]
            )
        )


def print_runtime_banner(pdf_enabled: bool, email_enabled: bool) -> None:
    print("Preflight OK")
    print(f"Python: {sys.version.split()[0]}")
    print(f"Executable: {sys.executable}")
    print(f"OS: {platform.system()} {platform.release()}")
    print(f"PDF enabled (effective): {pdf_enabled}")
    print(f"Email enabled (effective): {email_enabled}")
    print("")
