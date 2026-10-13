"""Command-line regression tests for merge_pdfs.py. Usage: cli_checks.py SCRIPT WORKDIR"""
import ctypes
import hashlib
import shutil
import subprocess
import sys
from ctypes import wintypes
from pathlib import Path

from pypdf import PdfReader, PdfWriter

script, work = sys.argv[1], Path(sys.argv[2])
if work.exists():
    shutil.rmtree(work)
fails = []


def check(name, ok, detail=""):
    print(("PASS  " if ok else "FAIL  ") + name + ("" if ok else f"   -> {detail}"))
    if not ok:
        fails.append(name)


def make(path, width, pages=1, user_pw=None, owner_pw=None):
    w = PdfWriter()
    for _ in range(pages):
        w.add_blank_page(width=width, height=200)
    if user_pw is not None:
        w.encrypt(user_pw, owner_password=owner_pw, algorithm="RC4-128")
    path.parent.mkdir(parents=True, exist_ok=True)
    w.write(path)


def fixtures(d, stale=False):
    for name, width, pages in [("file1.pdf", 101, 1), ("file2.pdf", 102, 2), ("file10.pdf", 110, 3), ("Upper.PDF", 150, 1)]:
        make(d / name, width, pages)
    if stale:
        make(d / "merged.pdf", 999, 5)
    (d / "broken.pdf").write_bytes(b"not really a pdf")
    (d / "notes.txt").write_text("ignore me")


def run(*args):
    p = subprocess.run([sys.executable, script, *map(str, args)], capture_output=True, text=True)
    return p.returncode, p.stdout, p.stderr


def layout(pdf):
    r = PdfReader(pdf)
    return [int(p.mediabox.width) for p in r.pages], [o.title for o in r.outline]


def sha(p):
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


EXPECT = ([101, 102, 102, 110, 110, 110, 150], ["file1", "file2", "file10", "Upper"])

a = work / "a"
fixtures(a, stale=True)
for attempt in (1, 2):
    code, out, err = run(a)
    check(f"default output run {attempt}: exit 0 + order + bookmarks", code == 0 and layout(a / "merged.pdf") == EXPECT, (code, out, err))
check("broken file reported on stderr", "skipped broken.pdf" in err, err)
check("singular 'page' wording", "(1 page)" in out, out)

b = work / "b" / "Board Packet"
fixtures(b)
code, out, err = run(b, "-o", b.parent)
check("-o existing folder", code == 0 and layout(b.parent / "merged.pdf") == EXPECT, (code, err))
code, out, err = run(b, "-o", b.parent / "Packet Sept.pdf")
check("-o file path", code == 0 and layout(b.parent / "Packet Sept.pdf") == EXPECT, (code, err))
code, out, err = run(b, "-o", b.parent / "October")
check("-o new folder, no slash", code == 0 and (b.parent / "October").is_dir() and layout(b.parent / "October" / "merged.pdf") == EXPECT, (code, err))
code, out, err = run(b, "-o", str(b.parent / "New Folder") + '"')
check("-o with stray PowerShell quote", code == 0 and layout(b.parent / "New Folder" / "merged.pdf") == EXPECT, (code, err))
code, out, err = run(str(b) + '"')
check("FOLDER with stray PowerShell quote", code == 0 and layout(b / "merged.pdf") == EXPECT, (code, err))

(work / "empty").mkdir()
code, out, err = run(work / "empty")
check("empty folder -> exit 1", code == 1 and "No PDFs found" in err, (code, err))
code, out, err = run(work / "nope")
check("missing folder -> exit 1", code == 1 and "Not a folder" in err, (code, err))

enc = work / "enc"
make(enc / "a_plain.pdf", 100)
make(enc / "b_locked.pdf", 100, user_pw="secret")
make(enc / "c_ownerpw.pdf", 100, user_pw="", owner_pw="owner")
code, out, err = run(enc)
check("encrypted: user-pw skipped, owner-pw merged", code == 0 and len(layout(enc / "merged.pdf")[0]) == 2 and "skipped b_locked.pdf" in err, (code, out, err))

bad = work / "bad"
bad.mkdir()
(bad / "broken.pdf").write_bytes(b"junk")
code, out, err = run(bad)
check("only unreadable PDFs -> exit 1, nothing written", code == 1 and "No readable PDFs" in err and not (bad / "merged.pdf").exists(), (code, err))

# Hold the output open with no sharing, the way a PDF viewer can.
k32 = ctypes.windll.kernel32
k32.CreateFileW.restype = wintypes.HANDLE
before = sha(a / "merged.pdf")
h = k32.CreateFileW(str(a / "merged.pdf"), 0x80000000, 0, None, 3, 0x80, None)
try:
    code, out, err = run(a)
finally:
    k32.CloseHandle(h)
check("locked output -> exit 1 + friendly message", code == 1 and "is it open in another program" in err, (code, err))
check("locked output left intact, no partial file", sha(a / "merged.pdf") == before and not list(a.glob("~*.part")), list(a.iterdir()))

code, out, err = run("-h")
check("-h mentions the window", code == 0 and "no arguments" in out, out)

print(f"\n{'ALL CLI TESTS PASSED' if not fails else f'{len(fails)} CLI FAILURE(S): {fails}'}")
sys.exit(1 if fails else 0)
