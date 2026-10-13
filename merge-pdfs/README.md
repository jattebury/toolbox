# merge-pdfs

Combine a folder of PDFs into one file, in the order you choose. Made for
assembling a monthly board packet: pick the folder, drag the reports into
order, click **Combine PDFs…**.

## Setup

Python 3.10 or newer with Tkinter (included in the python.org installer for Windows), then:

```
python -m pip install -r requirements.txt
```

`cryptography` is only needed for "secured" PDFs — files that open without a
password but have editing locked with AES (common for bank statements). Without
it, those files are skipped with a note saying how to fix it.

## Using the window

Double-click **`Combine PDFs.pyw`** (keep it in the same folder as
`merge_pdfs.py`), or run:

```
python merge_pdfs.py
```

1. **Choose folder…** lists the folder's PDFs in filename order, with page counts.
2. Put them in order: drag rows, use the **Move** buttons, or press Alt+↑/↓.
   **Add PDFs…** brings in files from other folders; **Remove from list** takes a
   row off the list (the file itself is untouched). Double-click a row to preview it.
3. **Combine PDFs…** asks where to save. It suggests the folder above the one you
   chose, named after it — e.g. `September\Board Packet.pdf`.

Files that can't be included show in red; click one to see why.

## Command line

```
python merge_pdfs.py FOLDER [-o OUTPUT]
```

Merges every PDF in `FOLDER` in natural filename order (`file2.pdf` before
`file10.pdf`). `OUTPUT` can be a `.pdf` path or a folder; it defaults to
`FOLDER/merged.pdf`.

## What it takes care of

- **Bookmarks** — the combined PDF has one bookmark per source file.
- **Problem files are skipped, with the reason** — damaged files, files that need a
  password, and secured files when `cryptography` isn't installed. A file is either
  merged completely or not at all.
- **Earlier packets are left out** — `merged.pdf`, `<folder name>.pdf`, or any PDF
  whose bookmarks name the other files in the folder, so an old packet saved in the
  folder isn't combined in a second time. The window and command line both say what
  was left out; **Add PDFs…** can still include it.
- **Your files are safe** — it refuses to save over one of the files being combined,
  and writes to a temporary file first, so a failed save never leaves a half-written
  PDF in place of a good one. If the output is open in a PDF viewer, it says so.

## Tests

```
python tests/run_all.py
```

Runs over 110 checks across six suites: command line, window (driven with simulated
clicks, drags and keys), the fixes from a code review, earlier-packet detection,
secured PDFs, and the launcher. The window checks open and close real windows, so
run them on a desktop session. `tests/fixtures/secured` holds small made-up secured
PDFs; rebuild them with `python tests/secured_pdf_checks.py make tests/fixtures/secured`
(needs `cryptography`).
