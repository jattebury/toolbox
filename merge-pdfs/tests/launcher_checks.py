"""Exercise Combine PDFs.pyw with fake neighbours and no real dialogs. Usage: launcher_checks.py LAUNCHER WORKDIR"""
import json
import runpy
import shutil
import sys
import tkinter.messagebox as mb
from pathlib import Path

launcher, work = Path(sys.argv[1]), Path(sys.argv[2])
shutil.rmtree(work, ignore_errors=True)
shown, fails = [], []
mb.showerror = lambda title, msg, **kw: shown.append((title, msg))


def check(name, ok, detail=""):
    print(("PASS  " if ok else "FAIL  ") + name + ("" if ok else f"   -> {detail}"))
    if not ok:
        fails.append(name)


def case(name, fake_script):
    """Copy the launcher into its own folder, optionally beside a fake merge_pdfs.py, and run it."""
    d = work / name
    d.mkdir(parents=True)
    copy = d / launcher.name
    shutil.copy(launcher, copy)
    if fake_script is not None:
        (d / "merge_pdfs.py").write_text(fake_script, encoding="utf-8")
    shown.clear()
    sys.argv = [str(copy), "dropped on the icon.pdf"]
    try:
        runpy.run_path(str(copy), run_name="__main__")
    except SystemExit as e:
        return e
    return None


case("missing", None)
check("merge_pdfs.py missing -> says to keep the files together",
      len(shown) == 1 and "Can't find merge_pdfs.py" in shown[0][1] and str(work / "missing") in shown[0][1], shown)

case("nopkg", "import pypdf_not_installed_xyz.errors\n")
check("package missing -> gives the install command",
      len(shown) == 1 and "python -m pip install pypdf_not_installed_xyz" in shown[0][1], shown)

case("boom", "raise RuntimeError('boom')\n")
check("other startup error -> shown, not swallowed", len(shown) == 1 and "RuntimeError: boom" in shown[0][1], shown)

fake = ("import json, pathlib, sys\n"
        "pathlib.Path(__file__).with_name('argv.json').write_text(json.dumps(sys.argv))\n"
        "sys.exit(0)\n")
result = case("argv", fake)
argv = json.loads((work / "argv" / "argv.json").read_text())
check("always window mode, even with a file dropped on the icon", argv == [str(work / "argv" / "merge_pdfs.py")], argv)
check("normal exit isn't reported as an error", isinstance(result, SystemExit) and result.code == 0 and not shown,
      (result, shown))

print(f"\n{'ALL LAUNCHER TESTS PASSED' if not fails else f'{len(fails)} LAUNCHER FAILURE(S): {fails}'}")
sys.exit(1 if fails else 0)
