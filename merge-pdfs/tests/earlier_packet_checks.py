"""Earlier combined packets are recognised by their bookmarks and left out.

Usage: earlier_packet_checks.py SCRIPT WORKDIR
"""
import importlib.util
import shutil
import subprocess
import sys
import tkinter as tk
from pathlib import Path

from pypdf import PdfReader, PdfWriter

script, work = Path(sys.argv[1]), Path(sys.argv[2])
spec = importlib.util.spec_from_file_location("merge_pdfs", script)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
if work.exists():
    shutil.rmtree(work)
fails = []


def check(name, ok, detail=""):
    print(("PASS  " if ok else "FAIL  ") + name + ("" if ok else f"   -> {detail}"))
    if not ok:
        fails.append(name)


def make(path, width, pages=1):
    w = PdfWriter()
    for _ in range(pages):
        w.add_blank_page(width=width, height=200)
    path.parent.mkdir(parents=True, exist_ok=True)
    w.write(path)


d = work / "Board Packet"
for name, w_, n in [("1. Agenda.pdf", 101, 1), ("2. Minutes.pdf", 102, 3), ("3. Treasurer.pdf", 103, 2)]:
    make(d / name, w_, n)
# An earlier packet saved inside the folder under a name the filename rule wouldn't catch.
m.merge(m.list_pdfs(d), work / "tmp.pdf")
shutil.move(work / "tmp.pdf", d / "September 2026 Packet.pdf")
# A PDF whose bookmarks name only one other file is an ordinary input.
w = PdfWriter()
w.add_blank_page(width=104, height=200)
w.add_outline_item("1. Agenda", 0)
w.write(d / "4. Notes.pdf")

pdfs = m.list_pdfs(d)
check("packet recognised by its bookmarks", m.looks_combined(d / "September 2026 Packet.pdf", pdfs))
check("one matching bookmark isn't enough", not m.looks_combined(d / "4. Notes.pdf", pdfs))
check("ordinary inputs aren't flagged", not any(m.looks_combined(p, pdfs) for p in pdfs if p.name[0].isdigit()))

p = subprocess.run([sys.executable, str(script), str(d), "-o", str(work / "cli.pdf")], capture_output=True, text=True)
got = [int(pg.mediabox.width) for pg in PdfReader(work / "cli.pdf").pages]
check("CLI leaves the old packet out and says so", p.returncode == 0 and got == [101, 102, 102, 102, 103, 103, 104]
      and "left out September 2026 Packet.pdf: looks like an earlier combined file" in p.stderr, (got, p.stderr))

m.enable_dpi_awareness()
root = tk.Tk()
app = m.MergeApp(root)
app.load_folder(d)
root.update()
names = [app.paths[i].name for i in app.tree.get_children()]
check("window leaves the old packet out", names == ["1. Agenda.pdf", "2. Minutes.pdf", "3. Treasurer.pdf", "4. Notes.pdf"], names)
check("window status says so", "Left out September 2026 Packet.pdf" in app.status["text"], app.status["text"])
root.destroy()

print(f"\n{'ALL EARLIER-PACKET CHECKS PASSED' if not fails else f'{len(fails)} EARLIER-PACKET FAILURE(S): {fails}'}")
sys.exit(1 if fails else 0)
