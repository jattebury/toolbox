"""Double-click to open the Combine PDFs window.

Runs merge_pdfs.py from this same folder, with no console window. Keep the two
files together. If the window can't start, a message box says why.
"""
import runpy
import sys
from pathlib import Path

script = Path(__file__).with_name("merge_pdfs.py")
sys.argv = [str(script)]  # no arguments: always the window, even if files are dropped on this icon
try:
    runpy.run_path(str(script), run_name="__main__")
except Exception as e:
    # There's no console here, so a startup error would otherwise vanish without a trace.
    import tkinter as tk
    from tkinter import messagebox

    if not script.exists():
        detail = f"Can't find merge_pdfs.py.\n\nKeep it in the same folder as this launcher:\n    {script.parent}"
    elif isinstance(e, ModuleNotFoundError) and e.name:
        package = e.name.split(".")[0]
        detail = f"The {package} package isn't installed for this Python.\n\nTo fix it, run:\n    python -m pip install {package}"
    else:
        detail = f"{type(e).__name__}: {e}"
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Combine PDFs", f"{detail}\n\n(Python: {sys.executable})", parent=root)
    root.destroy()
