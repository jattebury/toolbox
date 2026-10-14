# merge-pdfs

Combine a folder of PDFs and Word documents into one PDF, in the order you choose.
Made for assembling a monthly board packet: pick the folder, drag the reports into
order, click **Combine PDFs…**.

## Setup

Python 3.10 or newer with Tkinter (included in the python.org installer for Windows), then:

```
python -m pip install -r requirements.txt
```

- `pywin32` (plus Microsoft Word) is only needed for Word documents.
- `cryptography` is only needed for "secured" PDFs — files that open without a
  password but have editing locked with AES (common for bank statements).

Without them, those files are skipped with a note saying how to fix it.

## Using the window

Double-click **`Combine PDFs.pyw`** (keep it in the same folder as
`merge_pdfs.py`), or run:

```
python merge_pdfs.py
```

1. **Choose folder…** lists the folder's PDFs and Word documents in filename order,
   with page counts. Word documents are converted to PDF as they're listed.
2. Put them in order: drag rows, use the **Move** buttons, or press Alt+↑/↓.
   **Add files…** brings in files from other folders; **Remove from list** takes a
   row off the list (the file itself is untouched). Double-click a row to open it.
3. **Combine PDFs…** asks where to save. It suggests the folder above the one you
   chose, named after it — e.g. `September\Board Packet.pdf`.

Files that can't be included show in red; click one to see why.

## Command line

```
python merge_pdfs.py FOLDER [-o OUTPUT]
```

Merges every PDF and Word document in `FOLDER` in natural filename order
(`file2.pdf` before `file10.pdf`). `OUTPUT` can be a `.pdf` path or a folder; it
defaults to `FOLDER/merged.pdf`.

## What it takes care of

- **Bookmarks** — the combined PDF has one bookmark per source file (headings in Word
  documents become bookmarks inside theirs).
- **Word documents** (`.doc`, `.docx`) are converted by a hidden copy of Microsoft Word
  of their own, so documents you have open in Word aren't touched. They're converted
  as last saved, linked content (a pasted-link table, say) included — save your edits
  first. If a document looks open in Word, the window and command line say so, and
  Combine asks before going ahead. A document saved after it was listed is converted
  again when you combine. Comments aren't printed. Macros never run. If Word stops to
  ask a question nobody can see, the question is cancelled and the file skipped with
  a note; if Word stops responding, it's shut down after two minutes.
- **A document saved in two forms is used once** — `Minutes.docx` beside its "Save as
  PDF" copy `Minutes.pdf`: the one saved last is used (the PDF, if they were saved
  together), and the other is noted as left out.
- **Problem files are skipped, with the reason** — damaged files, files that need a
  password, and files a missing package would be needed for. A file is either merged
  completely or not at all.
- **Earlier packets are left out** — `merged.pdf`, `<folder name>.pdf`, or any PDF
  whose bookmarks name the other files in the folder, so an old packet saved in the
  folder isn't combined in a second time. The window and command line both say what
  was left out; **Add files…** can still include it.
- **Your files are safe** — it refuses to save over one of the files being combined,
  and writes to a temporary file first, so a failed save never leaves a half-written
  PDF in place of a good one. If the output is open in a PDF viewer, it says so.

## Adding another file type

Conversion is pluggable. To handle another kind of file (images, Excel, PowerPoint…),
subclass `Converter` in `merge_pdfs.py` — give it a `label`, its `extensions`, an
`unavailable()` check, and a `session()` that converts a batch — and add it to
`CONVERTERS`. The folder listing, the **Add files…** dialog, the command line and the
merge pick it up from there. `WordConverter` is the worked example;
`tests/conversion_checks.py` shows a minimal one.

## Tests

```
python tests/run_all.py
```

Runs eight suites (over 200 checks): command line, window (driven with simulated
clicks, drags and keys), the fixes from a code review, earlier-packet detection,
secured PDFs, the launcher, the conversion plumbing (with a fake converter), and
Word documents converted by the real Word (skipped if Word or pywin32 isn't
available). The window checks open and close real windows, so run them on a desktop
session. `tests/fixtures/secured` holds small made-up secured PDFs; rebuild them with
`python tests/secured_pdf_checks.py make tests/fixtures/secured` (needs `cryptography`).
