"""Converting other file types: the plumbing, tested with a fake converter (no Word needed).

Usage: conversion_checks.py SCRIPT WORKDIR
"""
import contextlib
import importlib.util
import io
import os
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
work.mkdir(parents=True)
fails = []


def check(name, ok, detail=""):
    print(("PASS  " if ok else "FAIL  ") + name + ("" if ok else f"   -> {detail}"))
    if not ok:
        fails.append(name)


class Fake(m.Converter):
    """A ".fake" file holds "<pages> <width>". One starting with "bad" is refused, "crash" raises an
    unexpected error, and "stuck" fails in a way not worth retrying."""

    label = "fake documents"
    extensions = (".fake",)

    def __init__(self):
        self.sessions = 0
        self.calls = []
        self.absolute = []
        self.down = None
        self.fail_setup = False

    def unavailable(self):
        return self.down

    @contextlib.contextmanager
    def session(self):
        self.sessions += 1
        if self.fail_setup:
            raise RuntimeError("couldn't start the fake converter")

        def convert(source, pdf):
            self.calls.append(source.name)
            self.absolute.append(source.is_absolute())
            text = source.read_text()
            if text.startswith("bad"):
                raise m.PdfProblem("the fake converter refused it")
            if text.startswith("crash"):
                raise ValueError("kaboom")
            if text.startswith("stuck"):
                raise m.PdfProblem("the fake converter gave up", retry=False)
            pages, width = map(int, text.split())
            w = PdfWriter()
            for _ in range(pages):
                w.add_blank_page(width=width, height=200)
            w.write(pdf)
        yield convert


fake = Fake()
m.CONVERTERS[:] = [fake]


def make_pdf(path, width, pages=1):
    w = PdfWriter()
    for _ in range(pages):
        w.add_blank_page(width=width, height=200)
    path.parent.mkdir(parents=True, exist_ok=True)
    w.write(path)


def write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


def widths(pdf):
    return [int(p.mediabox.width) for p in PdfReader(pdf).pages]


def set_mtime(path, seconds_from_now):
    t = path.stat().st_mtime + seconds_from_now
    os.utime(path, (t, t))


def run_cli(args):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = m.run_cli([str(a) for a in args])
    return code, out.getvalue(), err.getvalue()


# ---------------------------------------------------------------- the registry
check("supported extensions come from the converters", m.supported_extensions() == {".pdf", ".fake"}, m.supported_extensions())
check("messages name the kinds", m.kinds_text() == "PDFs or fake documents" and m.kinds_text("and") == "PDFs and fake documents")
check("Add files dialog offers every kind", m.file_types()[0][1] == "*.pdf *.fake" and ("fake documents", "*.fake") in m.file_types(),
      m.file_types())

listing = work / "listing"
make_pdf(listing / "a.pdf", 101)
write(listing / "b.fake", "1 102")
write(listing / "D.FAKE", "1 104")
write(listing / "~$b.fake", "lock")
write(listing / "c.txt", "not combinable")
check("folder listing: PDFs and convertible files, no lock files or others",
      [p.name for p in m.list_documents(listing)] == ["a.pdf", "b.fake", "D.FAKE"], m.list_documents(listing))

# ---------------------------------------------------------------- Conversions
c = m.Conversions()
seen = []
files = [listing / "a.pdf", listing / "b.fake", listing / "D.FAKE"]
c.convert(files, lambda i, total, f: seen.append((i, total, f.name)))
check("one session for the batch, PDFs left alone", fake.sessions == 1 and fake.calls == ["b.fake", "D.FAKE"], (fake.sessions, fake.calls))
check("progress reported before each conversion", seen == [(0, 2, "b.fake"), (1, 2, "D.FAKE")], seen)
check("a PDF is its own PDF", c.pdf_for(listing / "a.pdf") == listing / "a.pdf")
copy_b = c.pdf_for(listing / "b.fake")
check("converted copy lives in the temporary folder", copy_b.parent == c.folder and widths(copy_b) == [102], copy_b)
c.convert(files)
check("unchanged files aren't converted again", fake.calls == ["b.fake", "D.FAKE"] and fake.sessions == 1, fake.calls)

write(listing / "b.fake", "2 112")
c.convert(files)
check("a changed file is converted again (and only it)", fake.calls == ["b.fake", "D.FAKE", "b.fake"], fake.calls)
check("...its old copy is deleted", not copy_b.exists())
check("...and the new copy has the new content", widths(c.pdf_for(listing / "b.fake")) == [112, 112])
stat = (listing / "D.FAKE").stat()
os.utime(listing / "D.FAKE", ns=(stat.st_atime_ns, stat.st_mtime_ns + 2_000_000_000))
c.convert(files)
check("a re-saved file (same size, new time) is converted again", fake.calls[-1] == "D.FAKE", fake.calls)

write(listing / "bad.fake", "bad")
calls = len(fake.calls)
c.convert([listing / "bad.fake"])
try:
    c.pdf_for(listing / "bad.fake")
    check("a failed conversion raises PdfProblem with the reason", False)
except m.PdfProblem as e:
    check("a failed conversion raises PdfProblem with the reason", str(e) == "the fake converter refused it", e)
check("asking again for a failed file doesn't retry it", len(fake.calls) == calls + 1, fake.calls)
c.convert([listing / "bad.fake"])
check("convert() does retry earlier failures", len(fake.calls) == calls + 2, fake.calls)

fake.down = "install the fake converter first"
write(listing / "e.fake", "1 105")
sessions = fake.sessions
c.convert([listing / "e.fake"])
try:
    c.pdf_for(listing / "e.fake")
    ok = False
except m.PdfProblem as e:
    ok = str(e) == "install the fake converter first"
check("an unavailable converter: no session, and its reason is the file's reason", ok and fake.sessions == sessions)
fake.down = None
c.convert([listing / "e.fake"])
check("...and once it's available the file converts", widths(c.pdf_for(listing / "e.fake")) == [105])

gone = listing / "gone.fake"
write(gone, "1 106")
c.convert([gone])
gone.unlink()
try:
    c.pdf_for(gone)
    check("a file deleted after converting can't be used", False)
except m.PdfProblem as e:
    check("a file deleted after converting can't be used", "couldn't read" in str(e), e)

write(listing / "x.weird", "?")
try:
    c.pdf_for(listing / "x.weird")
    check("a file no converter handles is reported", False)
except m.PdfProblem as e:
    check("a file no converter handles is reported", "can't be converted" in str(e), e)

probe = ("import importlib.util, sys; from pathlib import Path\n"
         "spec = importlib.util.spec_from_file_location('m', sys.argv[1]); m = importlib.util.module_from_spec(spec)\n"
         "spec.loader.exec_module(m)\n"
         "class F(m.Converter):\n"
         "    extensions = ('.fake',)\n"
         "    def session(self):\n"
         "        import contextlib\n"
         "        @contextlib.contextmanager\n"
         "        def s():\n"
         "            yield lambda src, pdf: Path(pdf).write_bytes(b'%PDF')\n"
         "        return s()\n"
         "m.CONVERTERS[:] = [F()]\n"
         "c = m.Conversions(); c.convert([Path(sys.argv[2])]); print(c.folder)\n")
out = subprocess.run([sys.executable, "-c", probe, str(script), str(listing / "D.FAKE")], capture_output=True, text=True)
folder = Path(out.stdout.strip()) if out.stdout.strip() else None
check("the temporary folder is deleted when the program exits", out.returncode == 0 and folder and not folder.exists(),
      (out.returncode, out.stdout, out.stderr[-300:]))

# Converters get absolute paths (Word runs in its own process, with its own current folder).
odd = work / "odd"
make_pdf(odd / "1. Fine.pdf", 101)
write(odd / "2. Crash.fake", "crash")
write(odd / "3. Ok.fake", "1 103")
here = os.getcwd()
os.chdir(odd)
try:
    fake.absolute.clear()
    code, _, _ = run_cli([".", "-o", work / "odd-dot.pdf"])
finally:
    os.chdir(here)
check("converters get absolute paths, even for the command line's '.'", code == 0 and fake.absolute and all(fake.absolute), fake.absolute)

c2 = m.Conversions()
c2.convert(m.list_documents(odd))
check("an unexpected converter error becomes the file's reason",
      c2.results[(odd / "2. Crash.fake").resolve()][1] == "couldn't convert it (kaboom)", c2.results)
check("...and the next file still converts", widths(c2.pdf_for(odd / "3. Ok.fake")) == [103])
fake.fail_setup = True
c3 = m.Conversions()
c3.convert(m.list_documents(odd))
try:
    c3.pdf_for(odd / "3. Ok.fake")
    ok = None
except m.PdfProblem as e:
    ok = str(e)
check("a converter that can't start: every file gets the reason", ok == "couldn't convert fake documents (couldn't start the fake converter)", ok)
fake.fail_setup = False

write(odd / "4. Stuck.fake", "stuck")
c4 = m.Conversions()
c4.convert([odd / "4. Stuck.fake"])
calls = len(fake.calls)
c4.convert([odd / "4. Stuck.fake"])
check("a failure not worth retrying isn't retried while the file is unchanged", len(fake.calls) == calls, fake.calls[calls:])
c4.forget_failures()
c4.convert([odd / "4. Stuck.fake"])
check("...until failures are forgotten (Reload)", len(fake.calls) == calls + 1, fake.calls[calls:])
write(odd / "4. Stuck.fake", "stuck again")
c4.convert([odd / "4. Stuck.fake"])
check("...or the file changes", len(fake.calls) == calls + 2, fake.calls[calls:])

# The same document saved in two forms: only the one saved last is used.
tw = work / "twins" / "Board Packet"
make_pdf(tw / "1. Agenda.pdf", 101)
write(tw / "2. Minutes.fake", "1 120")
make_pdf(tw / "2. Minutes.pdf", 121)
set_mtime(tw / "2. Minutes.fake", -60)            # the PDF was saved after the document
make_pdf(tw / "3. Report.pdf", 131)
write(tw / "3. Report.fake", "1 130")
set_mtime(tw / "3. Report.pdf", -60)              # the document was edited after its PDF was made
write(tw / "4. Same.fake", "1 140")
make_pdf(tw / "4. Same.pdf", 141)
same = (tw / "4. Same.pdf").stat().st_mtime_ns
os.utime(tw / "4. Same.fake", ns=(same, same))   # saved at the same moment
write(tw / "5. Close.fake", "1 150")
make_pdf(tw / "5. Close.pdf", 151)
close = (tw / "5. Close.pdf").stat().st_mtime_ns
os.utime(tw / "5. Close.fake", ns=(close + 1_000_000_000, close + 1_000_000_000))  # a second apart: copied together
twins = {p.name: why for p, why in m.same_document_twins(m.list_documents(tw)).items()}
check("a document and its PDF copy: the one saved last is kept (the PDF if saved together)", twins == {
    "2. Minutes.fake": "same document as 2. Minutes.pdf, which was saved later",
    "3. Report.pdf": "same document as 3. Report.fake, which was saved later",
    "4. Same.fake": "same document as 4. Same.pdf",
    "5. Close.fake": "same document as 5. Close.pdf"}, twins)
code, out, err = run_cli([tw, "-o", work / "twins" / "cli.pdf"])
check("CLI leaves the older form out and says so", code == 0 and widths(work / "twins" / "cli.pdf") == [101, 121, 130, 141, 151]
      and "left out 2. Minutes.fake: same document as 2. Minutes.pdf, which was saved later" in err, (code, err))

# Documents that look open in Word (Office's "~$" lock file is beside them).
ow = work / "open"
write(ow / "2. Minutes.fake", "1 150")
write(ow / "~$ Minutes.fake", "lock")    # long names: the first two characters replaced
write(ow / "Notes.fake", "1 151")
write(ow / "~$Notes.fake", "lock")       # short names: "~$" in front
write(ow / "Other.fake", "1 152")
check("documents open in Word are spotted by their lock files",
      [p.name for p in m.open_in_word(m.list_documents(ow))] == ["2. Minutes.fake", "Notes.fake"])
code, out, err = run_cli([ow, "-o", work / "open-cli.pdf"])
check("CLI notes documents open in Word", code == 0 and "note: 2. Minutes.fake is open in Word" in err and "note: Notes.fake is open in Word" in err, err)

# ---------------------------------------------------------------- merge()
mix = work / "mix" / "Board Packet"
make_pdf(mix / "1. Agenda.pdf", 101)
write(mix / "2. Minutes.fake", "2 102")
write(mix / "3. Bad.fake", "bad")
make_pdf(mix / "4. Report.pdf", 104)
write(mix / "~$Minutes.fake", "lock")
docs = m.list_documents(mix)
merged, skipped = m.merge(docs, work / "mix" / "out.pdf")
check("merge converts as it goes, in list order", widths(work / "mix" / "out.pdf") == [101, 102, 102, 104],
      widths(work / "mix" / "out.pdf"))
check("bookmarks use the original file names", [o.title for o in PdfReader(work / "mix" / "out.pdf").outline]
      == ["1. Agenda", "2. Minutes", "4. Report"])
check("merge reports page counts and skips", [(p.name, n) for p, n in merged] == [("1. Agenda.pdf", 1), ("2. Minutes.fake", 2), ("4. Report.pdf", 1)]
      and [(p.name, why) for p, why in skipped] == [("3. Bad.fake", "the fake converter refused it")], (merged, skipped))
calls = len(fake.calls)
try:
    m.merge(docs, mix / "2. Minutes.fake")
    check("won't save over an input, and checks before converting anything", False)
except ValueError:
    check("won't save over an input, and checks before converting anything", len(fake.calls) == calls)

code, out, err = run_cli([mix, "-o", work / "mix" / "cli.pdf"])
check("CLI: converts, says so, and reports", code == 0 and "converting 2. Minutes.fake" in out
      and "+ 2. Minutes.fake (2 pages)" in out and "skipped 3. Bad.fake: the fake converter refused it" in err, (code, out, err))
check("CLI: output in natural order", widths(work / "mix" / "cli.pdf") == [101, 102, 102, 104])
(work / "empty").mkdir()
code, out, err = run_cli([work / "empty"])
check("CLI: empty folder names what it looked for", code == 1 and "No PDFs or fake documents found" in err, err)

# ---------------------------------------------------------------- the window
m.enable_dpi_awareness()
root = tk.Tk()
app = m.MergeApp(root)
t = app.tree
dialogs, answers = [], []
m.open_file = lambda p: None
m.messagebox.showerror = lambda title, msg, **kw: dialogs.append(("error", title, msg))
m.messagebox.askyesno = lambda title, msg, **kw: dialogs.append(("yesno", title, msg)) or (answers.pop(0) if answers else False)
progress_text = []
real_progress = app._show_progress


def spy(i, total, path):
    real_progress(i, total, path)
    progress_text.append(app.status["text"])


app._show_progress = spy
root.update()
check("window: empty status names the kinds", app.status["text"] == "Choose a folder to list its PDFs and fake documents.", app.status["text"])

app.load_folder(mix)
root.update()
rows = {app.paths[i].name: i for i in t.get_children()}
check("window lists PDFs and converted files, not lock files", list(rows) == ["1. Agenda.pdf", "2. Minutes.fake", "3. Bad.fake", "4. Report.pdf"],
      list(rows))
check("window shows converted page counts", str(t.set(rows["2. Minutes.fake"], "pages")) == "2" and t.set(rows["3. Bad.fake"], "pages") == "can't open",
      [t.set(i, "pages") for i in rows.values()])
check("window shows conversion progress", progress_text[:2] == ["Converting 2. Minutes.fake (1 of 2)…", "Converting 3. Bad.fake (2 of 2)…"],
      progress_text)
t.selection_set(rows["3. Bad.fake"])
root.update()
check("selecting a file that failed to convert says why", app.status["text"] == '"3. Bad.fake" will be skipped: the fake converter refused it',
      app.status["text"])

offered = {}
m.filedialog.askopenfilenames = lambda **kw: offered.update(kw) or (str(work / "elsewhere" / "Extra.fake"),)
write(work / "elsewhere" / "Extra.fake", "1 150")
app.add_files()
root.update()
check("Add files dialog offers the converter's files", ("fake documents", "*.fake") in offered.get("filetypes", []), offered)
check("an added file is converted and listed", app.paths[t.get_children()[-1]].name == "Extra.fake"
      and str(t.set(t.get_children()[-1], "pages")) == "1")

m.filedialog.asksaveasfilename = lambda **kw: str(work / "mix" / "window.pdf")
calls = len(fake.calls)
progress_text.clear()
app.combine()
check("combine reuses copies made when listing", len(fake.calls) == calls + 1 and fake.calls[-1] == "3. Bad.fake", fake.calls[calls:])
check("combine output follows the list", widths(work / "mix" / "window.pdf") == [101, 102, 102, 104, 150], widths(work / "mix" / "window.pdf"))
check("status isn't left saying Converting", "Converting" not in app.status["text"], app.status["text"])

write(mix / "2. Minutes.fake", "3 122")  # edited and saved after the folder was listed
write(mix / "3. Bad.fake", "1 133")      # and the problem file fixed
app.combine()
check("combine converts files edited since they were listed", widths(work / "mix" / "window.pdf") == [101, 122, 122, 122, 133, 104, 150],
      widths(work / "mix" / "window.pdf"))
check("...and says so while it works", any("Converting 2. Minutes.fake" in s for s in progress_text), progress_text)
check("...and the list shows the new page counts", str(t.set(rows["2. Minutes.fake"], "pages")) == "3"
      and str(t.set(rows["3. Bad.fake"], "pages")) == "1" and "bad" not in t.item(rows["3. Bad.fake"], "tags"),
      [t.set(i, "pages") for i in rows.values()])
check("...and the new total", app.status["text"].startswith("5 files  ·  7 pages"), app.status["text"])

fake.fail_setup = True
app.load_folder(odd)
root.update()
check("a converter that can't start doesn't stop the listing", [app.paths[i].name for i in t.get_children()][:1] == ["1. Fine.pdf"]
      and all(t.set(i, "pages") == "can't open" for i in t.get_children() if app.paths[i].suffix == ".fake"),
      [(app.paths[i].name, t.set(i, "pages")) for i in t.get_children()])
fake.fail_setup = False

app.load_folder(tw)
root.update()
check("window leaves the older form of a document out", [app.paths[i].name for i in t.get_children()]
      == ["1. Agenda.pdf", "2. Minutes.pdf", "3. Report.fake", "4. Same.pdf", "5. Close.pdf"])
check("...and says so", "Left out 2. Minutes.fake: same document as 2. Minutes.pdf, which was saved later" in app.status["text"],
      app.status["text"])

app.load_folder(ow)
root.update()
check("window notes documents open in Word", "Open in Word: 2. Minutes.fake, Notes.fake" in app.status["text"], app.status["text"])
m.filedialog.asksaveasfilename = lambda **kw: str(work / "open-window.pdf")
answers[:] = [False]
app.combine()
check("Combine asks first when documents are open in Word; No goes back", dialogs[-1][1] == "Open in Word"
      and not (work / "open-window.pdf").exists(), dialogs[-1:])
answers[:] = [True, False]
app.combine()
check("...Yes combines anyway", (work / "open-window.pdf").exists() and widths(work / "open-window.pdf") == [150, 151, 152])

calls = len(fake.calls)
app.load_folder(odd)
app.load_folder(odd)
first = len(fake.calls)
app.reload()
check("Reload tries a failure not worth retrying again", any(n == "4. Stuck.fake" for n in fake.calls[first:]), fake.calls[calls:])
root.destroy()

# ---------------------------------------------------------------- Word's availability checks (Word itself not used)
word = m.WordConverter()
saved = {k: sys.modules.get(k) for k in ("win32com", "win32com.client")}
sys.modules["win32com"] = sys.modules["win32com.client"] = None  # as if pywin32 weren't installed
try:
    msg = word.unavailable()
finally:
    for k, v in saved.items():
        if v is None:
            sys.modules.pop(k, None)
        else:
            sys.modules[k] = v
check("without pywin32: says how to install it", msg == "to include Word documents, run:  python -m pip install pywin32", msg)
import winreg  # noqa: E402

real_open = winreg.OpenKey
winreg.OpenKey = lambda *a, **k: (_ for _ in ()).throw(OSError("not registered"))
try:
    msg = word.unavailable()
finally:
    winreg.OpenKey = real_open
check("without Word: says Word is needed", msg == "converting Word documents needs Microsoft Word, which isn't installed on this computer", msg)

print(f"\n{'ALL CONVERSION CHECKS PASSED' if not fails else f'{len(fails)} CONVERSION FAILURE(S): {fails}'}")
sys.exit(1 if fails else 0)
