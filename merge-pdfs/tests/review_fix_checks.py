"""Regression tests for the review findings. Usage: review_fix_checks.py SCRIPT WORKDIR"""
import hashlib
import importlib.util
import io
import os
import shutil
import subprocess
import sys
import tkinter as tk
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, DecodedStreamObject, IndirectObject, NameObject, NullObject

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


def build(path, width, n=1, marker=b"P"):
    """PDF whose pages have real content streams; page width encodes which file it is."""
    w = PdfWriter()
    for i in range(1, n + 1):
        page = w.add_blank_page(width, 200)
        s = DecodedStreamObject()
        s.set_data(b"% " + marker + (b" page %d\n0 0 m 10 10 l S\n" % i))
        page[NameObject("/Contents")] = w._add_object(s)
    buf = io.BytesIO()
    w.write(buf)
    data = buf.getvalue()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return data


def widths(pdf):
    return [int(p.mediabox.width) for p in PdfReader(pdf).pages]


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


# ---------------------------------------------------------------- damaged files
def damaged_case(kind):
    d = work / "damaged" / kind
    build(d / "1. Agenda.pdf", 101, 1)
    build(d / "3. Treasurer.pdf", 103, 2)
    mid = d / "2. Check Detail.pdf"
    if kind == "A-bad-stream":
        data = build(mid, 102, 4, b"CD")
        at = data.index(b"% CD page 3")
        end = data.index(b"endstream", at)
        mid.write_bytes(data[:end] + b"endstreaX" + data[end + 9:])
    elif kind in ("B-null-annot", "C-dangling-annot"):
        w = PdfWriter()
        for _ in range(2):
            w.add_blank_page(102, 200)
        bad = NullObject() if kind == "B-null-annot" else IndirectObject(9999, 0, w)
        w.pages[0][NameObject("/Annots")] = ArrayObject([bad])
        with open(mid, "wb") as f:
            w.write(f)
    elif kind == "D-truncated":
        data = build(mid, 102, 6, b"CD")
        mid.write_bytes(data[:data.index(b"% CD page 4")])
    return d


for kind in ("A-bad-stream", "B-null-annot", "C-dangling-annot", "D-truncated"):
    d = damaged_case(kind)
    out = d.parent / f"{kind}.pdf"
    merged, skipped = m.merge(m.list_documents(d), out)
    got = widths(out)
    check(f"damaged {kind}: skipped file contributes no pages", got == [101, 103, 103]
          and [p.name for p, _ in skipped] == ["2. Check Detail.pdf"]
          and sum(n for _, n in merged) == len(got), (got, skipped))
    check(f"damaged {kind}: bookmarks match what's inside",
          [o.title for o in PdfReader(out).outline] == ["1. Agenda", "3. Treasurer"])

# ---------------------------------------------------------------- CLI -o naming an input
cli = work / "cli" / "Board Packet"
for name, w_ in [("1. Agenda.pdf", 101), ("2. Minutes.pdf", 102), ("3. Treasurer.pdf", 103)]:
    build(cli / name, w_)
before = {p.name: sha(p) for p in cli.iterdir()}


def run(*args, cwd=None):
    p = subprocess.run([sys.executable, str(script), *map(str, args)], capture_output=True, text=True, cwd=cwd)
    return p.returncode, p.stdout, p.stderr


for label, args, cwd in [
    ("-o names an input", (cli, "-o", cli / "1. Agenda.pdf"), None),
    ("-o names an input, different case", (cli, "-o", cli / "2. MINUTES.PDF"), None),
    ("-o names an input, relative from inside the folder", (".", "-o", "3. Treasurer.pdf"), cli),
]:
    code, out, err = run(*args, cwd=cwd)
    check(f"CLI refuses: {label}", code == 1 and "is one of the files being combined" in err, (code, out, err))
check("CLI: all inputs unchanged", {p.name: sha(p) for p in cli.iterdir()} == before,
      sorted(p.name for p in cli.iterdir()))
for attempt in (1, 2):
    code, out, err = run(cli)
    check(f"CLI default re-run still works (run {attempt})", code == 0 and widths(cli / "merged.pdf") == [101, 102, 103], (code, err))

# ---------------------------------------------------------------- window
m.enable_dpi_awareness()
root = tk.Tk()
app = m.MergeApp(root)
t = app.tree
opened, dialogs = [], []
m.open_file = lambda p: opened.append(Path(p).name)
m.messagebox.showerror = lambda title, msg, **kw: dialogs.append(("error", title, msg))
m.messagebox.askyesno = lambda title, msg, **kw: dialogs.append(("yesno", title, msg)) or False
clock = [10_000]


def ev(seq, gap=1000, **kw):
    """Generate an event with its own timestamp, so separate clicks aren't read as double-clicks."""
    clock[0] += gap
    t.event_generate(seq, time=clock[0], **kw)


def settle():
    for _ in range(3):
        root.update()


def names():
    return [app.paths[i].name for i in t.get_children()]


def visible(iid):
    return bool(t.bbox(iid))


long_dir = work / "long"
for i in range(1, 21):
    build(long_dir / f"{i:02d}. Item {i:02d}.pdf", 100 + i)
app.load_folder(long_dir)
settle()
original = names()

# Find a window height that leaves the last visible row cut off.
half = None
for height in range(480, 380, -3):
    root.geometry(f"{root.winfo_width()}x{height}")
    settle()
    t.yview_moveto(0)
    settle()
    for r in t.get_children():
        bb = t.bbox(r)
        if bb and bb[1] + bb[3] > t.winfo_height():
            half = r
            break
    if half:
        break
check("found a half-hidden bottom row to click", half is not None)
if half:
    bb = t.bbox(half)
    x, y = bb[0] + 40, bb[1] + (t.winfo_height() - bb[1]) // 2
    yview = t.yview()
    ev("<ButtonPress-1>", x=x, y=y); settle()
    ev("<B1-Motion>", gap=20, x=x + 1, y=y); settle()
    ev("<ButtonRelease-1>", gap=20, x=x + 1, y=y); settle()
    check("click+jiggle on half-hidden row doesn't reorder", names() == original, [n[:2] for n in names()])
    check("click on half-hidden row doesn't scroll", t.yview() == yview, (t.yview(), yview))
    check("click selects that row", t.selection() == (half,), t.selection())
    clock[0] += 5000
    ev("<ButtonPress-1>", x=x, y=y); ev("<ButtonRelease-1>", gap=40, x=x, y=y)
    ev("<ButtonPress-1>", gap=40, x=x, y=y); ev("<ButtonRelease-1>", gap=40, x=x, y=y); settle()
    check("double-click on half-hidden row opens that file", opened[-1:] == [app.paths[half].name], opened)

# Drag far past the bottom edge: the dragged row must stay in view.
root.geometry(f"{root.winfo_width()}x480")
settle()
t.yview_moveto(0)
settle()
row = t.get_children()[2]
bb = t.bbox(row)
x, y = bb[0] + 40, bb[1] + bb[3] // 2
ev("<ButtonPress-1>", x=x, y=y); settle()
for yy in range(y, t.winfo_height() + 60, 6):
    ev("<B1-Motion>", gap=20, x=x, y=yy); settle()
for _ in range(4):
    ev("<B1-Motion>", gap=20, x=x, y=t.winfo_height() + 60); settle()
ev("<ButtonRelease-1>", gap=20, x=x, y=t.winfo_height() + 60); settle()
check("drag past bottom edge keeps the row in view", visible(row), (t.index(row), t.yview()))
check("drag past bottom edge moved it down", t.index(row) > 2, t.index(row))
check("nothing lost while dragging", sorted(names()) == sorted(original))

# Drag up onto the column headings of a scrolled list: row stays in view.
t.yview_moveto(1)
settle()
row = [r for r in t.get_children() if visible(r)][-2]
bb = t.bbox(row)
x, y = bb[0] + 40, bb[1] + bb[3] // 2
ev("<ButtonPress-1>", x=x, y=y); settle()
for yy in range(y, 4, -6):
    ev("<B1-Motion>", gap=20, x=x, y=yy); settle()
for _ in range(3):
    ev("<B1-Motion>", gap=20, x=x, y=5); settle()
ev("<ButtonRelease-1>", gap=20, x=x, y=5); settle()
check("drag onto headings keeps the row in view", visible(row), (t.index(row), t.yview()))

# Double-clicking a heading or column divider in a scrolled list opens nothing.
t.yview_moveto(0.5)
settle()
n_open = len(opened)
for hx in (200, t.bbox(t.get_children()[0], "n")[2] if visible(t.get_children()[0]) else 40):
    clock[0] += 5000
    ev("<ButtonPress-1>", x=hx, y=5); ev("<ButtonRelease-1>", gap=40, x=hx, y=5)
    ev("<ButtonPress-1>", gap=40, x=hx, y=5); ev("<ButtonRelease-1>", gap=40, x=hx, y=5); settle()
check("double-click on heading/divider opens nothing", len(opened) == n_open, opened[n_open:])

# Reload after the folder was renamed in Explorer: the list and its order survive.
rn = work / "rename" / "Board Packet"
for name, w_ in [("1. A.pdf", 101), ("2. B.pdf", 102), ("3. C.pdf", 103)]:
    build(rn / name, w_)
app.load_folder(rn)
settle()
t.selection_set(t.get_children()[2])
app.move_to_top()
arranged = names()
status = app.status["text"]
rn.rename(rn.with_name("Board Packet FINAL"))
app.reload()
settle()
check("reload of a renamed folder shows an error", dialogs[-1:] and dialogs[-1][1] == "Can't open folder", dialogs[-1:])
check("...and keeps the arranged list", names() == arranged and app.status["text"] == status
      and str(app.combine_button["state"]) == "normal", (names(), app.status["text"]))
m.filedialog.askdirectory = lambda **kw: str(work / "does not exist")
app.choose_folder()
settle()
check("choosing a missing folder keeps the list", names() == arranged and dialogs[-1][1] == "Can't open folder")

# Earlier combined files are left out of the list, and the status line says so.
lo = work / "leftout" / "Board Packet"
for name, w_ in [("1. A.pdf", 101), ("2. B.pdf", 102), ("merged.pdf", 150), ("Board Packet.pdf", 160)]:
    build(lo / name, w_)
app.load_folder(lo)
settle()
check("earlier combined files left out", names() == ["1. A.pdf", "2. B.pdf"], names())
check("status names what was left out", "Left out Board Packet.pdf, merged.pdf" in app.status["text"], app.status["text"])
m.filedialog.askopenfilenames = lambda **kw: (str(lo / "merged.pdf"),)
app.add_files()
settle()
check("Add PDFs can still include one", names()[-1] == "merged.pdf"
      and "Left out Board Packet.pdf" in app.status["text"] and "merged.pdf:" not in app.status["text"], app.status["text"])

# Selecting a file that will be skipped says why.
why = work / "why"
build(why / "1. Fine.pdf", 101)
w = PdfWriter()
w.add_blank_page(102, 200)
w.encrypt("secret", algorithm="RC4-128")
w.write(why / "2. Locked.pdf")
shutil.copy(work / "damaged" / "A-bad-stream" / "2. Check Detail.pdf", why / "3. Damaged.pdf")
app.load_folder(why)
settle()
rows = t.get_children()
check("status counts files that will be skipped", "2 can't be included" in app.status["text"], app.status["text"])
t.selection_set(rows[1]); settle()
check("selecting a locked file says it needs a password", "2. Locked.pdf\" will be skipped: needs a password to open" in app.status["text"], app.status["text"])
t.selection_set(rows[2]); settle()
check("selecting a damaged file says it's damaged", "damaged or not a valid PDF" in app.status["text"], app.status["text"])
t.selection_set(rows[0]); settle()
check("selecting a good file shows the totals again", app.status["text"].startswith("3 files"), app.status["text"])
m.filedialog.asksaveasfilename = lambda **kw: str(why.parent / "why.pdf")
app.combine()
check("Done dialog gives each skip reason", dialogs[-1][0] == "yesno" and "2. Locked.pdf: needs a password to open" in dialogs[-1][2]
      and "3. Damaged.pdf: damaged or not a valid PDF" in dialogs[-1][2], dialogs[-1:])
check("output has only the good file", widths(why.parent / "why.pdf") == [101])

root.destroy()
print(f"\n{'ALL REVIEW-FIX TESTS PASSED' if not fails else f'{len(fails)} REVIEW-FIX FAILURE(S): {fails}'}")
sys.exit(1 if fails else 0)
