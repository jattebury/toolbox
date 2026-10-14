"""Drives MergeApp the way a person would. Usage: gui_checks.py SCRIPT WORKDIR [SCREENSHOT.png]"""
import ctypes
import hashlib
import importlib.util
import shutil
import sys
import time
import tkinter as tk
from ctypes import wintypes
from pathlib import Path

from pypdf import PdfReader, PdfWriter

script, work = Path(sys.argv[1]), Path(sys.argv[2])
shot = Path(sys.argv[3]) if len(sys.argv) > 3 else None
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


def make(path, width, pages=1, user_pw=None):
    w = PdfWriter()
    for _ in range(pages):
        w.add_blank_page(width=width, height=200)
    if user_pw:
        w.encrypt(user_pw, algorithm="RC4-128")
    path.parent.mkdir(parents=True, exist_ok=True)
    w.write(path)


src = work / "2026" / "September" / "Board Packet"
FILES = {"1. Board Agenda.pdf": (101, 1), "2. Meeting Minutes August.pdf": (102, 5),
         "3. Balance Sheet.pdf": (103, 2), "4. Check Detail.pdf": (104, 2), "10. Education Report.pdf": (110, 1)}
for name, (w, p) in FILES.items():
    make(src / name, w, p)
make(src / "5. Locked Report.pdf", 105, user_pw="secret")
(src / "notes.txt").write_text("not a pdf")
extra = work / "Downloads" / "Treasurer Letter.pdf"
make(extra, 150)
A, M, B, C, L, E = ("1. Board Agenda.pdf", "2. Meeting Minutes August.pdf", "3. Balance Sheet.pdf",
                    "4. Check Detail.pdf", "5. Locked Report.pdf", "10. Education Report.pdf")

m.enable_dpi_awareness()
root = tk.Tk()
app = m.MergeApp(root)
t = app.tree
root.update()
opened = []
m.open_file = lambda path: opened.append(Path(path).name)  # never launch a real viewer
clock = [10_000]


def names():
    return [app.paths[i].name for i in t.get_children()]


def iid(name):
    return next(i for i, p in app.paths.items() if p.name == name)


def select(*ns):
    t.selection_set([iid(n) for n in ns])
    root.update()


def settle():
    for _ in range(3):
        root.update()
        time.sleep(0.05)


check("empty state", "Choose a folder" in app.status["text"] and str(app.combine_button["state"]) == "disabled",
      (app.status["text"], app.combine_button["state"]))

app.load_folder(src)
settle()
check("natural order, locked file listed", names() == [A, M, B, C, L, E], names())
check("numbering", [str(t.set(i, "n")) for i in t.get_children()] == ["1", "2", "3", "4", "5", "6"], [t.set(i, "n") for i in t.get_children()])
check("page counts", [t.set(i, "pages") for i in t.get_children()] == ["1", "5", "2", "2", "can't open", "1"],
      [t.set(i, "pages") for i in t.get_children()])
check("locked row tagged", "bad" in t.item(iid(L), "tags"))
check("status", app.status["text"].startswith("6 files") and "11 pages" in app.status["text"]
      and "1 can't be included" in app.status["text"], app.status["text"])
check("combine enabled", str(app.combine_button["state"]) == "normal")
check("long folder path scrolled to show its end", app.folder_entry.xview()[1] == 1.0, app.folder_entry.xview())

if shot:
    from PIL import ImageGrab
    select(B)
    root.attributes("-topmost", True)
    root.lift()
    settle()
    time.sleep(0.4)
    settle()
    rect = wintypes.RECT()
    ctypes.windll.dwmapi.DwmGetWindowAttribute(int(root.wm_frame(), 16), 9, ctypes.byref(rect), ctypes.sizeof(rect))
    ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom), all_screens=True).save(shot)
    root.attributes("-topmost", False)
    print(f"screenshot -> {shot} ({rect.right - rect.left}x{rect.bottom - rect.top})")

# Buttons
select(A); app.move(1)
check("move down", names() == [M, A, B, C, L, E], names())
select(A, C); app.move(-1)
check("move up, two separate rows", names() == [A, M, C, B, L, E], names())
select(A, M); app.move(-1)
check("move up blocked at top keeps block", names() == [A, M, C, B, L, E], names())
select(E); app.move(1)
check("move down blocked at bottom", names() == [A, M, C, B, L, E], names())
select(A, L); app.move_to_bottom()
check("move to bottom keeps relative order", names() == [M, C, B, E, A, L], names())
select(E); app.move_to_top()
check("move to top", names() == [E, M, C, B, A, L], names())
check("renumbered", [str(t.set(i, "n")) for i in t.get_children()] == ["1", "2", "3", "4", "5", "6"], [t.set(i, "n") for i in t.get_children()])
check("selection kept after move", t.selection() == (iid(E),), t.selection())

# Keyboard
select(M); t.focus_force(); t.focus(iid(M)); settle()
t.event_generate("<Alt-Down>"); settle()
check("Alt+Down", names() == [E, C, M, B, A, L], names())
t.event_generate("<Alt-Up>"); settle()
check("Alt+Up", names() == [E, M, C, B, A, L], names())

# Drag and drop, moving the mouse in small steps like a real drag
app.reload(); settle()


def drag(from_row, to_row, modifier=""):
    rows = t.get_children()
    x, y0, _, h = t.bbox(rows[from_row], "file")
    _, y1, _, _ = t.bbox(rows[to_row], "file")
    x, y0, y1 = x + 20, y0 + h // 2, y1 + h // 2
    clock[0] += 5000  # well apart from the last click, so it is not read as a double-click
    t.event_generate(f"<{modifier}ButtonPress-1>", x=x, y=y0, time=clock[0]); settle()
    step = 3 if y1 > y0 else -3
    for y in list(range(y0, y1, step)) + [y1]:
        clock[0] += 10
        t.event_generate(f"<{modifier}B1-Motion>", x=x, y=y, time=clock[0])
        root.update()
    clock[0] += 10
    t.event_generate(f"<{modifier}ButtonRelease-1>", x=x, y=y1, time=clock[0]); settle()


drag(0, 2)
check("drag down lands on the drop row", names() == [M, B, A, C, L, E], names())
check("dragged row stays selected", t.selection() == (iid(A),), t.selection())
drag(5, 0)
check("drag up to the top", names() == [E, M, B, A, C, L], names())
drag(1, 5)
check("drag down to the bottom", names() == [E, B, A, C, L, M], names())
before = names()
drag(0, 3, modifier="Shift-")
check("shift-drag selects instead of moving", names() == before, names())
check("no file opened by clicks or drags", opened == [], opened)
check("cursor reset after drag", str(t["cursor"]) == "", t["cursor"])

# Combine (the locked file is still listed and must be reported as skipped)
asked, dialogs = {}, []
m.filedialog.asksaveasfilename = lambda **kw: asked.update(kw) or str(src.parent / "Board Packet.pdf")
m.messagebox.askyesno = lambda title, msg, **kw: dialogs.append(("yesno", title, msg)) or False
m.messagebox.showerror = lambda title, msg, **kw: dialogs.append(("error", title, msg))
order = names()
app.combine()
out = src.parent / "Board Packet.pdf"
check("save dialog suggests September/Board Packet.pdf",
      Path(asked.get("initialdir", "")) == src.parent and asked.get("initialfile") == "Board Packet.pdf", asked)
r = PdfReader(out)
expected = [FILES[n][0] for n in order if n != L for _ in range(FILES[n][1])]
check("output follows list order", [int(p.mediabox.width) for p in r.pages] == expected,
      ([int(p.mediabox.width) for p in r.pages], expected))
check("bookmarks follow list order", [o.title for o in r.outline] == [Path(n).stem for n in order if n != L])
check("done dialog reports skip", bool(dialogs) and dialogs[-1][0] == "yesno" and L in dialogs[-1][2]
      and "11 pages" in dialogs[-1][2], dialogs[-1:])
check("no partial file left", not list(src.parent.glob("~*.part")))

# Refuses to overwrite one of its own inputs
victim = src / B
before_hash = hashlib.sha256(victim.read_bytes()).hexdigest()
m.filedialog.asksaveasfilename = lambda **kw: str(victim)
app.combine()
check("won't overwrite an input", dialogs[-1][:2] == ("error", "Choose a different name")
      and hashlib.sha256(victim.read_bytes()).hexdigest() == before_hash, dialogs[-1:])

# Output open in a viewer (held with no sharing)
k32 = ctypes.windll.kernel32
k32.CreateFileW.restype = wintypes.HANDLE
m.filedialog.asksaveasfilename = lambda **kw: str(out)
before_hash = hashlib.sha256(out.read_bytes()).hexdigest()
h = k32.CreateFileW(str(out), 0x80000000, 0, None, 3, 0x80, None)
try:
    app.combine()
finally:
    k32.CloseHandle(h)
check("locked output -> friendly error", dialogs[-1][:2] == ("error", "Couldn't save") and "close it" in dialogs[-1][2],
      dialogs[-1:])
check("locked output untouched, no partial", hashlib.sha256(out.read_bytes()).hexdigest() == before_hash
      and not list(src.parent.glob("~*.part")))

# Cancelled save dialog does nothing
n_dialogs = len(dialogs)
m.filedialog.asksaveasfilename = lambda **kw: ""
app.combine()
check("cancel save = no-op", len(dialogs) == n_dialogs)

# Add from elsewhere (with a duplicate), then remove
m.filedialog.askopenfilenames = lambda **kw: (str(extra), str(src / M))
app.add_files(); settle()
check("add appends new file once, skips duplicate", names().count(M) == 1 and names()[-1] == extra.name, names())
check("added file shows where it's from", "(from " in t.set(iid(extra.name), "file"), t.set(iid(extra.name), "file"))
select(L, A)
app.remove(); settle()
check("remove takes rows off the list, files untouched", L not in names() and A not in names()
      and (src / L).exists() and (src / A).exists(), names())
check("selection moves to a neighbour", len(t.selection()) == 1, t.selection())
check("status drops the skip warning", "can't be included" not in app.status["text"], app.status["text"])
t.focus_force(); settle()
n_before = len(names())
t.event_generate("<Delete>"); settle()
check("Delete key removes", len(names()) == n_before - 1, names())

# Reload restores the folder in natural order
app.reload(); settle()
check("reload restores folder order", names() == [A, M, B, C, L, E], names())

# Folders with nothing usable
only_bad = work / "only_bad"
make(only_bad / "x.pdf", 100, user_pw="pw")
app.load_folder(only_bad); settle()
check("all-unreadable folder disables Combine", str(app.combine_button["state"]) == "disabled")
empty = work / "empty"
empty.mkdir()
app.load_folder(empty); settle()
check("empty folder status", app.status["text"] == "No files listed." and str(app.combine_button["state"]) == "disabled",
      app.status["text"])

# Unexpected errors surface as a dialog instead of vanishing
root.after(0, lambda: 1 / 0)
settle()
check("callback errors show a dialog", dialogs[-1][:2] == ("error", "Something went wrong")
      and "ZeroDivisionError" in dialogs[-1][2], dialogs[-1:])

root.destroy()

# No arguments opens the window; arguments use the command line
calls = []
m.run_gui = lambda: calls.append("gui")
m.run_cli = lambda argv: calls.append(("cli", argv)) or 0
sys.argv = ["merge_pdfs.py"]
m.main()
sys.argv = ["merge_pdfs.py", "some folder", "-o", "x.pdf"]
m.main()
check("dispatch: no args -> window, args -> CLI", calls == ["gui", ("cli", ["some folder", "-o", "x.pdf"])], calls)

print(f"\n{'ALL GUI TESTS PASSED' if not fails else f'{len(fails)} GUI FAILURE(S): {fails}'}")
sys.exit(1 if fails else 0)
