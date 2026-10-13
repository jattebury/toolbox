"""Run every check suite against ../merge_pdfs.py.

    python tests/run_all.py

The window checks open and close real windows, so run this on a desktop session
and leave the mouse alone for a minute. The secured-PDF suite runs in "present"
or "missing" mode depending on whether the cryptography package is installed.
"""
import importlib.util
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPT = HERE.parent / "merge_pdfs.py"
LAUNCHER = HERE.parent / "Combine PDFs.pyw"
crypto = "present" if importlib.util.find_spec("cryptography") else "missing"

with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
    work = Path(tmp)
    suites = [
        ("command line", "cli_checks.py", [SCRIPT, work / "cli"]),
        ("window", "gui_checks.py", [SCRIPT, work / "gui"]),
        ("review fixes", "review_fix_checks.py", [SCRIPT, work / "review"]),
        ("earlier packets", "earlier_packet_checks.py", [SCRIPT, work / "packet"]),
        (f"secured PDFs (cryptography {crypto})", "secured_pdf_checks.py",
         ["check", SCRIPT, HERE / "fixtures" / "secured", work / "secured", crypto]),
        ("launcher", "launcher_checks.py", [LAUNCHER, work / "launcher"]),
    ]
    failed = []
    for name, suite, args in suites:
        print(f"\n=== {name}", flush=True)
        if subprocess.run([sys.executable, str(HERE / suite), *map(str, args)]).returncode:
            failed.append(name)

print("\n" + ("All suites passed." if not failed else f"FAILED: {', '.join(failed)}"))
sys.exit(1 if failed else 0)
