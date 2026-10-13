"""Combine PDFs into one file, with a small window for choosing and ordering them.

Window -- run with no arguments:
    python merge_pdfs.py
    (or double-click "Combine PDFs.pyw" in the same folder -- no console window)

    Choose a folder, drag its PDFs into the order you want (or use the
    buttons), then click "Combine PDFs..." and pick where to save.

Command line -- merge a whole folder in natural filename order
(file2.pdf before file10.pdf):
    python merge_pdfs.py FOLDER [-o OUTPUT]

    FOLDER  folder containing the PDFs ("." for the current directory)
    OUTPUT  a .pdf file path, or a folder to write merged.pdf into (created if needed);
            defaults to FOLDER/merged.pdf (never re-merged into itself). It can't be
            one of the PDFs being combined.

Either way, each source file gets a bookmark in the output; PDFs that are
damaged or need a password are skipped with the reason; and an earlier
combined packet found in the folder (merged.pdf, "<folder name>.pdf", or any
PDF whose bookmarks name the other files) is left out, with a note saying so.

Requires: pip install pypdf cryptography
          (cryptography is only needed for "secured" AES-encrypted PDFs)
"""

import argparse
import logging
import os
import re
import sys
import tkinter as tk
import traceback
import webbrowser
from contextlib import contextmanager
from pathlib import Path
from tkinter import filedialog, font, messagebox, ttk

from pypdf import PdfReader, PdfWriter
from pypdf.errors import DependencyError

# pypdf logs low-level parse warnings for broken files; our "skipped" messages cover them.
logging.getLogger("pypdf").setLevel(logging.ERROR)

# Why a "secured" PDF (opens without a password, editing locked with AES) can't be read.
AES_HINT = "secured (AES-encrypted) PDF; to include it, run:  python -m pip install cryptography"


# ---------------------------------------------------------------------------
# Merging


class PdfProblem(Exception):
    """A PDF can't be included; the message says why, in plain words."""


def natural_key(path: Path) -> list:
    """Sort key that orders embedded numbers numerically, case-insensitively."""
    return [int(tok) if tok.isdigit() else tok.lower() for tok in re.split(r"(\d+)", path.name)]


def plural(n: int, word: str) -> str:
    return f"{n} {word}{'' if n == 1 else 's'}"


def list_pdfs(folder: Path, exclude: Path | None = None) -> list[Path]:
    """PDFs directly inside `folder` in natural filename order, leaving out `exclude`."""
    skip = exclude.resolve() if exclude else None
    return sorted(
        (p for p in folder.iterdir()
         if p.is_file() and p.suffix.lower() == ".pdf" and p.resolve() != skip),
        key=natural_key,
    )


def open_pdf(path: Path) -> PdfReader:
    """Open a PDF and check it can be copied, raising PdfProblem if not."""
    try:
        reader = PdfReader(path)
        locked = reader.is_encrypted and not reader.decrypt("")  # "" opens files with only an owner password
        if not locked:
            # Trial copy into a throwaway writer: PdfWriter.append adds pages one at a time,
            # so a file that fails partway must fail here, not leave some of its pages in the output.
            PdfWriter().append(reader)
    except DependencyError:
        raise PdfProblem(AES_HINT) from None
    except OSError as e:
        raise PdfProblem(f"couldn't read the file ({e.strerror or e})") from None
    except Exception as e:
        raise PdfProblem(f"damaged or not a valid PDF ({str(e) or type(e).__name__})") from None
    if locked:
        raise PdfProblem("needs a password to open")
    return reader


def looks_combined(pdf: Path, folder_pdfs: list[Path]) -> bool:
    """True if `pdf` looks like an earlier output of this tool: its bookmarks name two or more of the other PDFs."""
    others = {p.stem for p in folder_pdfs if p != pdf}
    try:
        titles = {item.title for item in PdfReader(pdf).outline if not isinstance(item, list)}
    except Exception:
        return False
    return len(others & titles) >= 2


def merge(pdfs: list[Path], output: Path) -> tuple[list[tuple[Path, int]], list[tuple[Path, str]]]:
    """Merge `pdfs`, in order, into `output`, with a bookmark per file.

    Returns (merged, skipped): (path, page count) for each file merged, and
    (path, reason) for each file that couldn't be included. Writes nothing if
    every file was skipped. Raises ValueError if `output` is one of `pdfs`,
    and OSError if it can't be written -- in which case any existing file at
    `output` is left untouched.
    """
    if output.resolve() in {p.resolve() for p in pdfs}:
        raise ValueError(f'"{output.name}" is one of the files being combined. Choose a different name.')

    writer = PdfWriter()
    merged, skipped = [], []
    for pdf in pdfs:
        try:
            reader = open_pdf(pdf)
            writer.append(reader, outline_item=pdf.stem)
        except Exception as e:
            skipped.append((pdf, str(e) or type(e).__name__))
        else:
            merged.append((pdf, len(reader.pages)))

    if merged:
        output.parent.mkdir(parents=True, exist_ok=True)
        # Write beside the target, then swap it in, so a failed save never
        # leaves a half-written file where a good one used to be.
        partial = output.with_name(f"~{output.name}.part")
        try:
            with open(partial, "wb") as f:
                writer.write(f)
            os.replace(partial, output)
        finally:
            partial.unlink(missing_ok=True)
    return merged, skipped


# ---------------------------------------------------------------------------
# Command line


def run_cli(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Combine all PDFs in a folder into one file. Run with no arguments to open a window instead.")
    parser.add_argument("folder", nargs="?", default=".", help="folder containing PDFs (default: current directory)")
    parser.add_argument("-o", "--output", help="output .pdf file, or folder to put merged.pdf in (default: FOLDER/merged.pdf)")
    args = parser.parse_args(argv)

    # Windows PowerShell 5.1 passes a quoted path ending in a backslash, like 'C:\My Folder\',
    # as C:\My Folder" -- drop that stray quote (it can't appear in a Windows filename).
    folder = Path(args.folder.rstrip('"'))
    if not folder.is_dir():
        print(f"Not a folder: {folder}", file=sys.stderr)
        return 1

    if args.output:
        output = Path(args.output.rstrip('"'))
        # Anything not named *.pdf is a folder to put merged.pdf in.
        if output.is_dir() or output.suffix.lower() != ".pdf":
            output = output / "merged.pdf"
    else:
        output = folder / "merged.pdf"

    # Leave out the default merged.pdf so re-running doesn't fold the last result back in.
    # Any other PDF already in FOLDER stays an input, so merge() refuses to overwrite it.
    pdfs = list_pdfs(folder, exclude=output if output.name.lower() == "merged.pdf" else None)
    # Leave out earlier combined packets too, so an old packet saved in FOLDER isn't combined in again.
    earlier = [p for p in pdfs if looks_combined(p, pdfs)]
    for p in earlier:
        print(f"  left out {p.name}: looks like an earlier combined file", file=sys.stderr)
    pdfs = [p for p in pdfs if p not in earlier]
    if not pdfs:
        print(f"No PDFs found in {folder}", file=sys.stderr)
        return 1

    try:
        merged, skipped = merge(pdfs, output)
    except PermissionError:
        # On Windows this usually means the output PDF is open in a viewer.
        print(f"Can't write {output}: permission denied (is it open in another program?)", file=sys.stderr)
        return 1
    except ValueError as e:
        print(e, file=sys.stderr)
        return 1

    for pdf, pages in merged:
        print(f"  + {pdf.name} ({plural(pages, 'page')})")
    for pdf, reason in skipped:
        print(f"  skipped {pdf.name}: {reason}", file=sys.stderr)
    if not merged:
        print("No readable PDFs to merge.", file=sys.stderr)
        return 1

    total = sum(pages for _, pages in merged)
    print(f"Wrote {output}: {plural(total, 'page')} from {len(merged)} of {plural(len(pdfs), 'file')}")
    return 0


# ---------------------------------------------------------------------------
# Window


def open_file(path: Path) -> None:
    """Open a file in its default app (for a PDF, the usual viewer)."""
    if hasattr(os, "startfile"):
        os.startfile(path)
    else:
        webbrowser.open(path.as_uri())


def enable_dpi_awareness() -> None:
    """Have Windows draw the window crisply on high-DPI screens instead of stretching it."""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


class MergeApp:
    """Window for choosing a folder, putting its PDFs in order, and combining them."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.folder: Path | None = None
        self.paths: dict[str, Path] = {}        # list row id -> file
        self.pages: dict[str, int | None] = {}  # list row id -> page count, None if it can't be included
        self.reasons: dict[str, str] = {}       # list row id -> why it can't be included
        self.left_out: list[Path] = []          # earlier combined files found in the folder
        self._next_id = 0
        self._dragging = ""

        root.title("Combine PDFs")
        root.report_callback_exception = self._show_error
        scale = root.winfo_fpixels("1i") / 96  # 1.0 at 100% Windows display scaling

        def px(n: float) -> int:
            return round(n * scale)

        root.geometry(f"{px(820)}x{px(480)}")
        root.minsize(px(560), px(320))
        line = font.nametofont("TkDefaultFont").metrics("linespace")
        ttk.Style(root).configure("Treeview", rowheight=line + px(8))

        frame = ttk.Frame(root, padding=px(12))
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        # Folder picker
        top = ttk.Frame(frame)
        top.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, px(8)))
        top.columnconfigure(1, weight=1)
        ttk.Label(top, text="Folder:").grid(row=0, column=0, padx=(0, px(6)))
        self.folder_text = tk.StringVar()
        self.folder_entry = ttk.Entry(top, textvariable=self.folder_text, state="readonly")
        self.folder_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(top, text="Choose folder…", command=self.choose_folder).grid(row=0, column=2, padx=(px(6), 0))

        # File list
        box = ttk.Frame(frame)
        box.grid(row=1, column=0, sticky="nsew")
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        tree = self.tree = ttk.Treeview(box, columns=("n", "file", "pages"), show="headings", selectmode="extended")
        tree.heading("n", text="#")
        tree.heading("file", text="File", anchor="w")
        tree.heading("pages", text="Pages")
        tree.column("n", width=px(40), stretch=False, anchor="e")
        tree.column("file", width=px(520), anchor="w")
        tree.column("pages", width=px(90), stretch=False, anchor="center")
        tree.tag_configure("bad", foreground="#b3261e")
        scroll = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        tree.bind("<ButtonPress-1>", self._drag_start, add="+")
        tree.bind("<B1-Motion>", self._drag_motion, add="+")
        tree.bind("<ButtonRelease-1>", self._drag_end, add="+")
        tree.bind("<Double-1>", self._open_row)
        tree.bind("<<TreeviewSelect>>", lambda e: self._update_status())
        tree.bind("<Alt-Up>", lambda e: self.move(-1) or "break")
        tree.bind("<Alt-Down>", lambda e: self.move(1) or "break")
        tree.bind("<Delete>", lambda e: self.remove())

        # Buttons beside the list
        side = ttk.Frame(frame)
        side.grid(row=1, column=1, sticky="n", padx=(px(10), 0))
        buttons = [
            ("Move to top", self.move_to_top),
            ("Move up", lambda: self.move(-1)),
            ("Move down", lambda: self.move(1)),
            ("Move to bottom", self.move_to_bottom),
            None,
            ("Add PDFs…", self.add_files),
            ("Remove from list", self.remove),
            ("Reload folder", self.reload),
        ]
        for row, spec in enumerate(buttons):
            if spec is None:
                ttk.Separator(side).grid(row=row, column=0, sticky="ew", pady=px(8))
            else:
                ttk.Button(side, text=spec[0], command=spec[1]).grid(row=row, column=0, sticky="ew", pady=px(2))

        # Hint, status, and the main action
        ttk.Label(frame, foreground="gray",
                  text="Drag rows to reorder  ·  Alt+↑/↓ moves the selection  ·  Double-click a row to open it").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(px(6), 0))
        bottom = ttk.Frame(frame)
        bottom.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(px(8), 0))
        bottom.columnconfigure(0, weight=1)
        self.status = ttk.Label(bottom, wraplength=px(640))
        self.status.grid(row=0, column=0, sticky="w")
        self.combine_button = ttk.Button(bottom, text="Combine PDFs…", command=self.combine)
        self.combine_button.grid(row=0, column=1, sticky="e")
        self._refresh()

    # -- Choosing files

    def choose_folder(self) -> None:
        chosen = filedialog.askdirectory(parent=self.root, title="Choose the folder with your PDFs",
                                         initialdir=str(self.folder or Path.cwd()), mustexist=True)
        if chosen:
            self.load_folder(Path(chosen))

    def load_folder(self, folder: Path) -> None:
        """List `folder`'s PDFs in natural filename order, replacing the current list."""
        # Read the folder first, so a folder that's gone or unreadable leaves the current list alone.
        try:
            pdfs = list_pdfs(folder)
        except OSError as e:
            messagebox.showerror("Can't open folder", f"Couldn't read {folder}:\n\n{e.strerror or e}", parent=self.root)
            return
        self.folder = folder
        self.folder_text.set(str(folder))
        self.folder_entry.xview_moveto(1)  # a long path shows its end, where the folder name is
        self.tree.delete(*self.tree.get_children())
        self.paths.clear()
        self.pages.clear()
        self.reasons.clear()
        # Leave out earlier combined packets so they aren't combined in a second time: merged.pdf
        # from the command line, "<folder name>.pdf", or any PDF whose bookmarks name the other files.
        own = {"merged.pdf", f"{folder.name}.pdf".lower()}
        with self._busy():
            self.left_out = [p for p in pdfs if p.name.lower() in own or looks_combined(p, pdfs)]
            self._add([p for p in pdfs if p not in self.left_out])

    def reload(self) -> None:
        if self.folder:
            self.load_folder(self.folder)

    def add_files(self) -> None:
        chosen = filedialog.askopenfilenames(parent=self.root, title="Add PDFs",
                                             initialdir=str(self.folder or Path.cwd()),
                                             filetypes=[("PDF files", "*.pdf")])
        if chosen:
            with self._busy():
                self._add([Path(p) for p in chosen])

    def _add(self, files: list[Path]) -> None:
        """Append files to the end of the list, skipping any already on it."""
        listed = {p.resolve() for p in self.paths.values()}
        for path in files:
            if path.resolve() in listed:
                continue
            listed.add(path.resolve())
            iid = f"row{self._next_id}"
            self._next_id += 1
            try:
                pages = len(open_pdf(path).pages)
            except PdfProblem as e:
                pages = None
                self.reasons[iid] = str(e)
            self.paths[iid] = path
            self.pages[iid] = pages
            name = path.name if path.parent == self.folder else f"{path.name}   (from {path.parent})"
            self.tree.insert("", "end", iid=iid, values=("", name, "can't open" if pages is None else pages),
                             tags=("bad",) if pages is None else ())
        self._refresh()

    # -- Reordering

    def move(self, step: int) -> None:
        """Shift the selected rows one place up (-1) or down (+1), keeping them in order."""
        order = list(self.tree.get_children())
        chosen = set(self.tree.selection())
        if step > 0:
            order.reverse()
        # Walk from the leading edge, swapping each selected row past an unselected neighbour.
        for i in range(1, len(order)):
            if order[i] in chosen and order[i - 1] not in chosen:
                order[i - 1], order[i] = order[i], order[i - 1]
        if step > 0:
            order.reverse()
        self._reorder(order)

    def move_to_top(self) -> None:
        order = self.tree.get_children()
        chosen = set(self.tree.selection())
        self._reorder([i for i in order if i in chosen] + [i for i in order if i not in chosen])

    def move_to_bottom(self) -> None:
        order = self.tree.get_children()
        chosen = set(self.tree.selection())
        self._reorder([i for i in order if i not in chosen] + [i for i in order if i in chosen])

    def remove(self) -> None:
        """Take the selected rows off the list. The files themselves are not touched."""
        chosen = set(self.tree.selection())
        if not chosen:
            return
        order = self.tree.get_children()
        last = max(order.index(i) for i in chosen)
        rest = [i for i in order if i not in chosen]
        # Select the row after the removed ones (or the last remaining row) so Delete can repeat.
        neighbour = order[last + 1] if last + 1 < len(order) else (rest[-1] if rest else "")
        self.tree.delete(*chosen)
        for iid in chosen:
            del self.paths[iid], self.pages[iid]
            self.reasons.pop(iid, None)
        if neighbour:
            self.tree.selection_set(neighbour)
            self.tree.focus(neighbour)
        self._refresh()

    def _reorder(self, order: list[str]) -> None:
        for index, iid in enumerate(order):
            self.tree.move(iid, "", index)
        self._refresh()
        if self.tree.selection():
            self.tree.see(self.tree.selection()[0])

    def _drag_start(self, event: tk.Event) -> str | None:
        # A plain click on a row can start a drag; Shift/Ctrl-clicks just change the selection.
        on_row = self.tree.identify_region(event.x, event.y) == "cell"
        self._dragging = self.tree.identify_row(event.y) if on_row and not event.state & 0x0005 else ""
        if self._dragging:
            # Select the row here instead of in ttk's own handler, which scrolls a half-hidden
            # row into view and so slides the next row under the pointer (a tiny jiggle would
            # then swap two files, and a double-click would open the wrong one).
            self.tree.focus_set()
            self.tree.focus(self._dragging)
            self.tree.selection_set(self._dragging)
            return "break"
        return None

    def _drag_motion(self, event: tk.Event) -> None:
        if not self._dragging:
            return
        # Outside the rows (headings, or past either edge), Tk's identify_row names rows
        # that are scrolled out of view, so scroll and use the edge row instead.
        shown = [r for r in self.tree.get_children() if self.tree.bbox(r)]
        top = self.tree.bbox(shown[0])[1] if shown else 0  # rows start below the headings
        bottom = self.tree.winfo_height() - 1
        if event.y < top:
            self.tree.yview_scroll(-1, "units")
        elif event.y > bottom:
            self.tree.yview_scroll(1, "units")
        target = self.tree.identify_row(min(max(event.y, top + 1), bottom))  # y == top maps to the row above
        if target and target != self._dragging:
            self.tree.move(self._dragging, "", self.tree.index(target))
            self.tree.configure(cursor="sb_v_double_arrow")
            self._refresh()

    def _drag_end(self, event: tk.Event) -> None:
        self._dragging = ""
        self.tree.configure(cursor="")

    # -- Combining

    def combine(self) -> None:
        pdfs = [self.paths[iid] for iid in self.tree.get_children()]
        if not pdfs:
            return
        # Suggest saving beside the folder, named after it: September\Board Packet.pdf
        name = self.folder.name if self.folder and self.folder.name else "merged"
        chosen = filedialog.asksaveasfilename(
            parent=self.root, title="Save combined PDF as",
            initialdir=str(self.folder.parent if self.folder else Path.cwd()),
            initialfile=f"{name}.pdf", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not chosen:
            return
        output = Path(chosen)
        try:
            with self._busy():
                merged, skipped = merge(pdfs, output)
        except ValueError as e:
            messagebox.showerror("Choose a different name", str(e), parent=self.root)
            return
        except PermissionError:
            messagebox.showerror("Couldn't save", f'Windows wouldn\'t let this save "{output.name}".\n\n'
                                 "If it's open in a PDF viewer, close it and try again.", parent=self.root)
            return
        except OSError as e:
            messagebox.showerror("Couldn't save", str(e), parent=self.root)
            return

        if not merged:
            messagebox.showerror("Nothing to combine", "None of the files could be included.", parent=self.root)
            return
        total = sum(pages for _, pages in merged)
        message = f'Saved "{output.name}": {plural(total, "page")} from {plural(len(merged), "file")}.'
        if skipped:
            message += "\n\nSkipped:\n" + "\n".join(f"  • {p.name}: {why}" for p, why in skipped)
        if messagebox.askyesno("Done", f"{message}\n\nOpen it now?", parent=self.root):
            open_file(output)

    # -- Helpers

    def _refresh(self) -> None:
        """Renumber the rows, then update the status line and the Combine button."""
        for n, iid in enumerate(self.tree.get_children(), 1):
            self.tree.set(iid, "n", n)
        self._update_status()

    def _update_status(self) -> None:
        rows = self.tree.get_children()
        bad = [iid for iid in rows if self.pages[iid] is None]
        selected = self.tree.selection()
        if len(selected) == 1 and selected[0] in bad:
            # Say why the selected file can't be included (and how to fix it, where possible).
            text = f'"{self.paths[selected[0]].name}" will be skipped: {self.reasons[selected[0]]}'
        elif rows:
            text = f"{plural(len(rows), 'file')}  ·  {plural(sum(self.pages[i] or 0 for i in rows), 'page')}"
            if bad:
                text += f"  ·  {len(bad)} can't be included and will be skipped (click {'it' if len(bad) == 1 else 'one'} to see why)"
        else:
            text = "No PDFs listed." if self.folder else "Choose a folder to list its PDFs."
        left_out = [p.name for p in self.left_out if p not in self.paths.values()]
        if left_out:
            text += f"  ·  Left out {', '.join(left_out)}: looks like an earlier combined file (Add PDFs… to include it)"
        self.status.configure(text=text)
        self.combine_button.configure(state="normal" if len(rows) > len(bad) else "disabled")

    @contextmanager
    def _busy(self):
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            yield
        finally:
            self.root.configure(cursor="")

    def _open_row(self, event: tk.Event) -> None:
        # Only rows open a file; ignore double-clicks on the headings, column dividers, or empty space.
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        if iid:
            open_file(self.paths[iid])

    def _show_error(self, exc_type, exc, tb) -> None:
        traceback.print_exception(exc_type, exc, tb)
        messagebox.showerror("Something went wrong", f"{exc_type.__name__}: {exc}", parent=self.root)


def run_gui() -> None:
    enable_dpi_awareness()
    root = tk.Tk()
    MergeApp(root)
    root.mainloop()


def main() -> int:
    if len(sys.argv) == 1:
        run_gui()
        return 0
    return run_cli(sys.argv[1:])


if __name__ == "__main__":
    sys.exit(main())
