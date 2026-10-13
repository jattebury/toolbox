"""Secured (AES-encrypted) PDFs that open without a password.

    secured_pdf_checks.py make DIR                                    (needs cryptography; rebuilds the fixtures)
    secured_pdf_checks.py check SCRIPT FIXTURES WORKDIR missing|present

"missing" checks the behaviour when the cryptography package isn't installed (those files are
skipped with an install hint); "present" checks that they merge normally once it is.
"""
import importlib.util
import shutil
import subprocess
import sys
import tkinter as tk
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, NameObject

FILES = [("1. Agenda.pdf", 101, None), ("2. Bank Statement.pdf", 102, "AES-256"),
         ("3. Audit Letter.pdf", 103, "AES-128"), ("4. Minutes.pdf", 104, "RC4-128")]

if sys.argv[1] == "make":
    d = Path(sys.argv[2])
    shutil.rmtree(d, ignore_errors=True)
    d.mkdir(parents=True)
    for name, width, alg in FILES:
        w = PdfWriter()
        for i in range(2):
            page = w.add_blank_page(width, 200)
            s = DecodedStreamObject()
            s.set_data(b"BT /F1 12 Tf 10 10 Td (page) Tj ET\n")
            page[NameObject("/Contents")] = w._add_object(s)
        if alg:
            w.encrypt(user_password="", owner_password="owner", algorithm=alg)
        w.write(d / name)
    print("made", [(n, a) for n, _, a in FILES])
    sys.exit(0)

script, d, work, expect = Path(sys.argv[2]), Path(sys.argv[3]), Path(sys.argv[4]), sys.argv[5]
spec = importlib.util.spec_from_file_location("merge_pdfs", script)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
work.mkdir(parents=True, exist_ok=True)
fails = []


def check(name, ok, detail=""):
    print(("PASS  " if ok else "FAIL  ") + f"[{expect}] " + name + ("" if ok else f"   -> {detail}"))
    if not ok:
        fails.append(name)


AES = {"2. Bank Statement.pdf", "3. Audit Letter.pdf"}
out = work / f"merged_{expect}.pdf"
merged, skipped = m.merge(m.list_pdfs(d), out)
got = [int(p.mediabox.width) for p in PdfReader(out).pages]
if expect == "missing":
    check("AES files skipped with the install hint", {p.name for p, _ in skipped} == AES
          and all(why == m.AES_HINT for _, why in skipped), skipped)
    check("everything else merged", got == [101, 101, 104, 104], got)
else:
    check("nothing skipped", skipped == [], skipped)
    check("AES files merged in order", got == [101, 101, 102, 102, 103, 103, 104, 104], got)

p = subprocess.run([sys.executable, str(script), str(d), "-o", str(work / f"cli_{expect}.pdf")],
                   capture_output=True, text=True)
if expect == "missing":
    check("CLI tells how to fix it", p.returncode == 0 and p.stderr.count("python -m pip install cryptography") == 2, p.stderr)
else:
    check("CLI merges all 4", p.returncode == 0 and "from 4 of 4 files" in p.stdout, (p.stdout, p.stderr))

root = tk.Tk()
app = m.MergeApp(root)
app.load_folder(d)
root.update()
rows = {app.paths[i].name: i for i in app.tree.get_children()}
pages = {n: str(app.tree.set(i, "pages")) for n, i in rows.items()}
if expect == "missing":
    check("AES rows flagged when the folder loads (not only at Combine)",
          pages["2. Bank Statement.pdf"] == "can't open" and pages["3. Audit Letter.pdf"] == "can't open", pages)
    app.tree.selection_set(rows["3. Audit Letter.pdf"])
    root.update()
    check("selecting one shows the fix", "python -m pip install cryptography" in app.status["text"], app.status["text"])
else:
    check("AES rows show page counts", pages["2. Bank Statement.pdf"] == "2" and pages["3. Audit Letter.pdf"] == "2", pages)
root.destroy()

print(f"\n{'ALL SECURED-PDF CHECKS PASSED' if not fails else f'{len(fails)} SECURED-PDF FAILURE(S): {fails}'} ({expect})")
sys.exit(1 if fails else 0)
