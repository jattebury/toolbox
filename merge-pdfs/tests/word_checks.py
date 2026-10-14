"""Word documents converted by the real Microsoft Word. Skips (and passes) if Word or pywin32 isn't available.

Usage: word_checks.py SCRIPT WORKDIR

Everything happens in WORKDIR, in hidden copies of Word started for the purpose. At the end it
checks that Word processes that were already running are still there, that none of the copies
it started are left behind, and that Word's crash-tracking lists don't mention its test files.
"""
import csv
import importlib.util
import shutil
import subprocess
import sys
import time
import tkinter as tk
import winreg
from pathlib import Path

from pypdf import PdfReader, PdfWriter

script, work = Path(sys.argv[1]), Path(sys.argv[2])
spec = importlib.util.spec_from_file_location("merge_pdfs", script)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
why = m.WordConverter().unavailable()
if why:
    print(f"SKIPPED: {why}")
    sys.exit(0)
import win32com.client  # noqa: E402

if work.exists():
    shutil.rmtree(work)
fails = []


def check(name, ok, detail=""):
    print(("PASS  " if ok else "FAIL  ") + name + ("" if ok else f"   -> {detail}"))
    if not ok:
        fails.append(name)


def word_pids():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW).stdout
    return {int(r[1]) for r in csv.reader(out.splitlines()) if len(r) > 1 and r[1].isdigit()}


def resiliency_mentions(text):
    """Entries in Word's crash-tracking lists (document recovery, "caused a serious error") that mention `text`."""
    found = []

    def walk(path):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
        except OSError:
            return
        i = 0
        while True:
            try:
                name, value, _ = winreg.EnumValue(key, i)
            except OSError:
                break
            data = value.decode("utf-16-le", "ignore") if isinstance(value, bytes) else str(value)
            if text.lower() in data.lower():
                found.append(f"{path}\\{name}")
            i += 1
        j = 0
        while True:
            try:
                sub = winreg.EnumKey(key, j)
            except OSError:
                break
            walk(f"{path}\\{sub}")
            j += 1

    walk(r"Software\Microsoft\Office\16.0\Word\Resiliency")
    return found


def pdf_text(pdf):
    return " ".join(p.extract_text() or "" for p in PdfReader(pdf).pages)


def private_word():
    w = win32com.client.DispatchEx("Word.Application")
    if w.Visible or w.Documents.Count:
        print("SKIPPED: couldn't get a private copy of Word to make test documents")
        sys.exit(0)
    w.DisplayAlerts = 0
    return w


before = word_pids()
src = work / "Board Packet"
src.mkdir(parents=True)

# ---------------------------------------------------------------- test documents, made by a Word of our own
maker = private_word()
links_before = maker.Options.UpdateLinksAtOpen


def make_doc(path, pages=1, fmt=16, password="", heading=None, comment=None, tracked=None, text=None):
    d = maker.Documents.Add()
    if heading:
        first = d.Paragraphs(1).Range
        first.Text = heading
        first.Style = -2  # wdStyleHeading1
        d.Content.InsertParagraphAfter()
    if text:
        d.Content.InsertAfter(text)
    for i in range(pages if not text else 0):
        end = d.Content.End - 1
        d.Range(end, end).InsertAfter(f"{path.stem} page {i + 1}")
        if i < pages - 1:
            end = d.Content.End - 1
            d.Range(end, end).InsertBreak(7)  # wdPageBreak
    if comment:
        d.Comments.Add(d.Paragraphs(1).Range, comment)
    if tracked:
        d.TrackRevisions = True
        end = d.Content.End - 1
        d.Range(end, end).InsertAfter(tracked)
        d.TrackRevisions = False
    if password:
        d.Password = password  # (SaveAs2's Password argument is silently ignored)
    d.SaveAs2(FileName=str(path), FileFormat=fmt)
    d.Close(SaveChanges=0)


w = PdfWriter()
w.add_blank_page(612, 792)
w.write(src / "1. Agenda.pdf")
make_doc(src / "2. Minutes.docx", pages=3, heading="Call to Order")
make_doc(src / "3. Old Report.doc", pages=2, fmt=0)  # wdFormatDocument: Word 97-2003
make_doc(src / "4. Locked.docx", password="pw")
(src / "5. Broken.docx").write_bytes(b"PK\x03\x04 not really a Word document" * 40)
make_doc(src / "6. Reviewed.docx", comment="COMMENTTEXT", tracked=" TRACKEDINSERT")
(src / "~$2. Minutes.docx").write_bytes(b"\x0bSomeone" + b"\x00" * 150)  # Word's lock file

# A report with an automatic link to another document, saved while that said SEPTEMBER.
other = work / "linked"
other.mkdir()
make_doc(other / "Figures.docx", text="SEPTEMBER FIGURES")
d = maker.Documents.Add()
d.Content.Text = "Report: "
end = d.Content.End - 1
link_target = str(other / "Figures.docx").replace("\\", "\\\\")
d.Fields.Add(d.Range(end, end), 56, f'Word.Document.12 "{link_target}" "" \\a \\t', False)  # wdFieldLink
d.SaveAs2(FileName=str(other / "Report.docx"), FileFormat=16)
d.Close(SaveChanges=0)
d = maker.Documents.Open(FileName=str(other / "Figures.docx"))  # next month, the linked file changes
d.Content.Text = "OCTOBER FIGURES"
d.Save()
d.Close(SaveChanges=0)

# A long document, to make sure watching for question boxes doesn't interfere with a long export.
d = maker.Documents.Add()
for i in range(120):
    end = d.Content.End - 1
    d.Range(end, end).InsertAfter(f"Long report page {i + 1}")
    if i < 119:
        end = d.Content.End - 1
        d.Range(end, end).InsertBreak(7)
d.SaveAs2(FileName=str(other / "Long.docx"), FileFormat=16)
d.Close(SaveChanges=0)
maker.Quit(SaveChanges=0)
del maker, d

check("the password test file really is encrypted", (src / "4. Locked.docx").read_bytes()[:4] == bytes.fromhex("D0CF11E0"))
files = m.list_documents(src)
check("folder listing includes .doc/.docx, not Word's lock file",
      [p.name for p in files] == ["1. Agenda.pdf", "2. Minutes.docx", "3. Old Report.doc", "4. Locked.docx",
                                  "5. Broken.docx", "6. Reviewed.docx"], [p.name for p in files])

# ---------------------------------------------------------------- converting
c = m.Conversions()
seen = []
started = time.perf_counter()
c.convert(files, lambda i, total, f: seen.append(f.name))
print(f"   converted {len(seen)} Word files in {time.perf_counter() - started:.1f}s (one Word start)")


def outcome(path, conversions=c):
    try:
        return len(PdfReader(conversions.pdf_for(path)).pages)
    except m.PdfProblem as e:
        return str(e)


check(".docx converts, with every page", outcome(src / "2. Minutes.docx") == 3, outcome(src / "2. Minutes.docx"))
check(".doc (Word 97-2003) converts", outcome(src / "3. Old Report.doc") == 2, outcome(src / "3. Old Report.doc"))
check("password-protected: skipped, not prompted for", outcome(src / "4. Locked.docx") == "needs a password to open",
      outcome(src / "4. Locked.docx"))
check("corrupt file: skipped with Word's reason", str(outcome(src / "5. Broken.docx")).startswith("Word couldn't convert it (")
      and "corrupt" in str(outcome(src / "5. Broken.docx")).lower(), outcome(src / "5. Broken.docx"))
reviewed = pdf_text(c.pdf_for(src / "6. Reviewed.docx"))
check("comments aren't printed in the PDF", "COMMENTTEXT" not in reviewed, reviewed)
check("tracked changes show as the document currently reads", "TRACKEDINSERT" in reviewed, reviewed)

c2 = m.Conversions()
c2.convert([other / "Report.docx", other / "Long.docx"])
linked = pdf_text(c2.pdf_for(other / "Report.docx"))
check("linked content appears as last saved, not fetched afresh", "SEPTEMBER" in linked and "OCTOBER" not in linked, linked)
check("a long document converts completely", outcome(other / "Long.docx", c2) == 120, outcome(other / "Long.docx", c2))

# ---------------------------------------------------------------- merging
out = work / "packet.pdf"
merged, skipped = m.merge(files, out, c)
reader = PdfReader(out)
top = [o.title for o in reader.outline if not isinstance(o, list)]
check("packet has every page that converted", len(reader.pages) == 1 + 3 + 2 + 1, len(reader.pages))
check("bookmarks use the Word file names", top == ["1. Agenda", "2. Minutes", "3. Old Report", "6. Reviewed"], top)
nested = [o for o in reader.outline if isinstance(o, list)]
check("a Word heading becomes a bookmark inside its document's", any(getattr(x, "title", "") == "Call to Order" for x in sum(nested, [])),
      [[getattr(x, "title", x) for x in n] for n in nested])
check("skipped files are reported with reasons", [p.name for p, _ in skipped] == ["4. Locked.docx", "5. Broken.docx"], skipped)

p = subprocess.run([sys.executable, str(script), str(src), "-o", str(work / "cli.pdf")], capture_output=True, text=True)
check("command line converts Word files", p.returncode == 0 and "converting 2. Minutes.docx" in p.stdout
      and len(PdfReader(work / "cli.pdf").pages) == 7, (p.returncode, p.stdout, p.stderr))
p = subprocess.run([sys.executable, str(script.resolve()), ".", "-o", str(work / "cli-dot.pdf")], cwd=src,
                   capture_output=True, text=True)
check("command line run from inside the folder ('.') converts Word files", p.returncode == 0
      and len(PdfReader(work / "cli-dot.pdf").pages) == 7 and "couldn't find" not in p.stderr, (p.returncode, p.stderr))

# ---------------------------------------------------------------- the window
m.enable_dpi_awareness()
root = tk.Tk()
app = m.MergeApp(root)
answers, asked = [], []
m.messagebox.askyesno = lambda title, msg, **k: asked.append(title) or (answers.pop(0) if answers else False)
m.filedialog.asksaveasfilename = lambda **kw: str(work / "window.pdf")
app.load_folder(src)
root.update()
rows = {app.paths[i].name: i for i in app.tree.get_children()}
pages = {n: str(app.tree.set(i, "pages")) for n, i in rows.items()}
check("window shows Word files with page counts", pages["2. Minutes.docx"] == "3" and pages["3. Old Report.doc"] == "2", pages)
check("window notes a document with a Word lock file as open in Word", "Open in Word: 2. Minutes.docx" in app.status["text"],
      app.status["text"])
app.tree.selection_set(rows["4. Locked.docx"])
root.update()
check("window says why a Word file will be skipped", "will be skipped: needs a password to open" in app.status["text"], app.status["text"])
answers[:] = [True, False]  # "Combine anyway?" yes; "Open it now?" no
app.combine()
check("Combine asks before going ahead with a document open in Word", asked[:1] == ["Open in Word"], asked)
check("window combines Word files", len(PdfReader(work / "window.pdf").pages) == 7)
root.destroy()

# ---------------------------------------------------------------- questions, stuck copies, and what Word is left with
# A question box nobody can see is cancelled straight away (here: the password prompt the dummy
# password normally prevents), instead of waiting for the timeout.
word = m.HiddenWord(timeout=120)
saved_options = dict(m.WORD_OPEN_OPTIONS)
del m.WORD_OPEN_OPTIONS["PasswordDocument"]
started = time.perf_counter()
try:
    word.convert(src / "4. Locked.docx", work / "asked.pdf")
    result = "converted?!"
except m.PdfProblem as e:
    result = str(e)
took = time.perf_counter() - started
m.WORD_OPEN_OPTIONS.update(saved_options)
check("a question box is cancelled at once, with a reason", result.startswith("Word wanted to ask something") and took < 15,
      (result, f"{took:.1f}s"))
check("our copy of Word never runs macros", word.app.AutomationSecurity == 3)
check("...and doesn't update links", word.app.Options.UpdateLinksAtOpen is False)
recent_before = [r.Path + "\\" + r.Name for r in word.app.RecentFiles]
doc = word.app.Documents.Open(FileName=str(src / "3. Old Report.doc"), **m.WORD_OPEN_OPTIONS)
check("documents are opened read-only", doc.ReadOnly is True, doc.ReadOnly)
doc.Close(SaveChanges=0)
recent_after = [r.Path + "\\" + r.Name for r in word.app.RecentFiles]
check("documents don't go on Word's Recent list", recent_after == recent_before
      and not any("Old Report" in r for r in recent_after), recent_after[:3])
word.close()

# A Word that doesn't finish in time is shut down and replaced. (Simulated: a stand-in for Word's
# document list that never answers, so the real Word is shut down while it's idle, not mid-document.)
word = m.HiddenWord(timeout=1)
word._start()


class NeverAnswers:
    def Open(self, **options):
        time.sleep(4)
        raise RuntimeError("no answer")


class StandIn:
    def __init__(self, app):
        self._app = app
        self.Documents = NeverAnswers()

    def __getattr__(self, name):
        return getattr(self._app, name)


first_pid = word.pid
word.app = StandIn(word.app)
try:
    word.convert(src / "2. Minutes.docx", work / "stuck.pdf")
    result = "converted?!"
except m.PdfProblem as e:
    result = e
check("a Word that takes too long is shut down", "stopped responding" in str(result) and getattr(result, "retry", True) is False, result)
check("...its process is gone", first_pid not in word_pids(), first_pid)
word.timeout = 120
try:
    word.convert(src / "2. Minutes.docx", work / "after.pdf")
    ok = len(PdfReader(work / "after.pdf").pages) == 3
except m.PdfProblem as e:
    ok = e
check("...and the next file gets a fresh copy of Word", ok is True and word.pid != first_pid, ok)
word.close()

# ---------------------------------------------------------------- what's left behind
deadline = time.time() + 45
while time.time() < deadline and word_pids() - before:
    time.sleep(1)
after = word_pids()
check("Word processes that were already running are still running", before <= after, (sorted(before), sorted(after)))
check("no hidden copies of Word left behind", not (after - before), sorted(after - before))
checker = private_word()
links_after = checker.Options.UpdateLinksAtOpen
checker.Quit(SaveChanges=0)
del checker
check("Word's own 'update links' setting is as it was", links_after == links_before, (links_before, links_after))
check("Word's crash-tracking lists don't mention the test files", not resiliency_mentions(str(work)), resiliency_mentions(str(work)))

print(f"\n{'ALL WORD CHECKS PASSED' if not fails else f'{len(fails)} WORD FAILURE(S): {fails}'}")
sys.exit(1 if fails else 0)
