"""Combine PDFs and Word documents into one PDF, with a small window for choosing and ordering them.

Window -- run with no arguments:
    python merge_pdfs.py
    (or double-click "Combine PDFs.pyw" in the same folder -- no console window)

    Choose a folder, drag its files into the order you want (or use the
    buttons), then click "Combine PDFs..." and pick where to save.

Command line -- merge a whole folder in natural filename order
(file2.pdf before file10.pdf):
    python merge_pdfs.py FOLDER [-o OUTPUT]

    FOLDER  folder containing the files ("." for the current directory)
    OUTPUT  a .pdf file path, or a folder to write merged.pdf into (created if needed);
            defaults to FOLDER/merged.pdf (never re-merged into itself). It can't be
            one of the files being combined.

Word documents (.doc, .docx) are converted to PDF first, as last saved, by a hidden
copy of Microsoft Word of their own -- documents you have open in Word aren't touched,
but changes you haven't saved there aren't included (you're warned).

Either way, each source file gets a bookmark in the output; files that are damaged
or need a password are skipped with the reason; an earlier combined packet found in
the folder (merged.pdf, "<folder name>.pdf", or any PDF whose bookmarks name the other
files) is left out; and of a document saved in two forms (Minutes.docx and its
Minutes.pdf), only the one saved last is used. Anything left out is noted.

To handle another kind of file, add a Converter -- see "Converting other file
types to PDF" below.

Requires: pip install pypdf cryptography pywin32
          (cryptography is only needed for "secured" AES-encrypted PDFs, and
          pywin32 only for Word documents)
"""

import argparse
import atexit
import functools
import logging
import os
import re
import shutil
import sys
import tempfile
import threading
import time
import tkinter as tk
import traceback
import uuid
import webbrowser
from collections.abc import Callable
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
    """A file can't be included; the message says why, in plain words. `retry` is False when
    trying again before the file changes would only waste time (Word stopped responding, say)."""

    def __init__(self, message: str, retry: bool = True):
        super().__init__(message)
        self.retry = retry


def natural_key(path: Path) -> list:
    """Sort key that orders embedded numbers numerically, case-insensitively."""
    return [int(tok) if tok.isdigit() else tok.lower() for tok in re.split(r"(\d+)", path.name)]


def plural(n: int, word: str) -> str:
    return f"{n} {word}{'' if n == 1 else 's'}"


def list_documents(folder: Path, exclude: Path | None = None) -> list[Path]:
    """Files directly inside `folder` that can be combined (PDFs, plus anything a Converter
    handles), in natural filename order, leaving out `exclude` and Word's "~$" lock files."""
    kinds = supported_extensions()
    skip = exclude.resolve() if exclude else None
    return sorted(
        (p for p in folder.iterdir()
         if p.is_file() and p.suffix.lower() in kinds and not p.name.startswith("~$") and p.resolve() != skip),
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


def looks_combined(path: Path, folder_files: list[Path]) -> bool:
    """True if `path` looks like an earlier output of this tool: a PDF whose bookmarks name two or more of the other files."""
    if path.suffix.lower() != ".pdf":
        return False
    others = {p.stem for p in folder_files if p != path}
    try:
        titles = {item.title for item in PdfReader(path).outline if not isinstance(item, list)}
    except Exception:
        return False
    return len(others & titles) >= 2


def same_document_twins(files: list[Path]) -> dict[Path, str]:
    """Files to leave out because another of `files` is the same document in another form --
    "Minutes.docx" beside "Minutes.pdf", its "Save as PDF" copy, say. Of each such set the one
    saved last is kept; files saved within a couple of seconds of each other (copied or
    downloaded together, say) count as saved together, and then the PDF is kept.
    Returns {file to leave out: why}."""
    together = 2_000_000_000  # nanoseconds

    def saved(p: Path) -> int:
        try:
            return p.stat().st_mtime_ns
        except OSError:
            return 0

    sets: dict[tuple[Path, str], list[Path]] = {}
    for p in files:
        sets.setdefault((p.parent, p.stem.lower()), []).append(p)
    twins = {}
    for group in sets.values():
        if len(group) > 1:
            newest = max(saved(p) for p in group)
            latest = [p for p in group if newest - saved(p) <= together]
            keep = next((p for p in latest if p.suffix.lower() == ".pdf"), max(latest, key=saved))
            for p in group:
                if p != keep:
                    later = ", which was saved later" if saved(keep) - saved(p) > together else ""
                    twins[p] = f"same document as {keep.name}{later}"
    return twins


def open_in_word(files: list[Path]) -> list[Path]:
    """Which of `files` look open in Word (or Excel, PowerPoint...) right now. Office keeps a "~$"
    lock file beside a document while it's open: the document's name with its first two
    characters replaced, or with "~$" in front of a short name. Changes not yet saved there
    won't be in the PDF."""
    locks: dict[Path, set[str]] = {}
    found = []
    for f in files:
        if f.suffix.lower() == ".pdf":
            continue
        if f.parent not in locks:
            try:
                locks[f.parent] = {p.name[2:].lower() for p in f.parent.iterdir() if p.name.startswith("~$")}
            except OSError:
                locks[f.parent] = set()
        name = f.name.lower()
        if name in locks[f.parent] or name[2:] in locks[f.parent]:
            found.append(f)
    return found


def merge(files: list[Path], output: Path, conversions: "Conversions | None" = None,
          progress: "Callable[[int, int, Path], None] | None" = None,
          ) -> tuple[list[tuple[Path, int]], list[tuple[Path, str]]]:
    """Merge `files`, in order, into the PDF `output`, with a bookmark per file.

    Files that aren't PDFs are converted first; pass `conversions` to reuse copies already
    made, and `progress` to hear about each conversion (see Conversions.convert).
    Returns (merged, skipped): (path, page count) for each file merged, and (path, reason)
    for each file that couldn't be included. Writes nothing if every file was skipped.
    Raises ValueError if `output` is one of `files`, and OSError if it can't be written --
    in which case any existing file at `output` is left untouched.
    """
    if output.resolve() in {p.resolve() for p in files}:
        raise ValueError(f'"{output.name}" is one of the files being combined. Choose a different name.')
    if conversions is None:
        conversions = Conversions()
    conversions.convert(files, progress)

    writer = PdfWriter()
    merged, skipped = [], []
    for path in files:
        try:
            reader = open_pdf(conversions.pdf_for(path))
            writer.append(reader, outline_item=path.stem)
        except Exception as e:
            skipped.append((path, str(e) or type(e).__name__))
        else:
            merged.append((path, len(reader.pages)))

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
# Converting other file types to PDF
#
# Each Converter turns one kind of file into PDFs. To handle another kind, subclass
# Converter and add an instance to CONVERTERS: the folder listing, the "Add files"
# dialog, the command line and merge() all pick it up from there.


class Converter:
    """Turns one kind of file (Word documents, say) into PDFs."""

    label = ""                        # what the files are called, plural: "Word documents"
    extensions: tuple[str, ...] = ()  # lower-case, with the dot: (".doc", ".docx")

    def unavailable(self) -> str | None:
        """Why this converter can't work on this computer, and how to fix that -- or None if it can."""
        return None

    def session(self):
        """A context manager for converting a batch of files. It yields convert(source, pdf),
        which writes the PDF version of `source` to `pdf` (both absolute paths) or raises
        PdfProblem saying why not. Expensive setup (starting Word, say) belongs here, so a
        batch pays for it once. Any other error, here or in convert(), becomes the reason
        the files concerned are skipped."""
        raise NotImplementedError


class WordConverter(Converter):
    """Converts Word documents with Microsoft Word, in a hidden copy of Word of its own."""

    label = "Word documents"
    extensions = (".doc", ".docx")
    timeout = 120  # seconds one document may take before a stuck Word is shut down

    def unavailable(self) -> str | None:
        if sys.platform != "win32":
            return "converting Word documents needs Microsoft Word on Windows"
        import winreg
        try:
            winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application").Close()
        except OSError:
            return "converting Word documents needs Microsoft Word, which isn't installed on this computer"
        try:
            import win32com.client  # noqa: F401  (part of pywin32)
        except ImportError:
            return "to include Word documents, run:  python -m pip install pywin32"
        return None

    @contextmanager
    def session(self):
        word = HiddenWord(self.timeout)
        try:
            yield word.convert
        finally:
            word.close()


# How the hidden Word opens each document.
WORD_OPEN_OPTIONS = dict(
    ReadOnly=True, AddToRecentFiles=False, Visible=False,
    ConfirmConversions=False, NoEncodingDialog=True,  # no "convert this file?" or "which encoding?" questions
    PasswordDocument="#",  # a wrong password makes a protected file fail rather than ask for one
)


class HiddenWord:
    """A private, invisible copy of Word: started on first use, started afresh if it gets stuck,
    and shut down by close(). It never uses a Word that's already running, so documents open
    in your own Word aren't touched."""

    def __init__(self, timeout: float):
        self.timeout = timeout
        self.app = None
        self.process = None        # handle to our copy's process, to watch its windows and shut it down
        self.pid = 0
        self.update_links = None   # Word's "update links at open" setting as it was, to put back
        self.start_failure = None  # why Word couldn't start; not tried again for the rest of this batch

    def _start(self) -> None:
        if self.start_failure:
            raise PdfProblem(self.start_failure)
        import pywintypes
        import win32com.client
        app = None
        try:
            app = win32com.client.DispatchEx("Word.Application")  # a new copy of Word, not the one on screen
            if patiently(lambda: app.Visible or app.Documents.Count > 0):
                app = None  # somebody's Word, not a fresh copy: don't change it, open files in it, or quit it
                raise PdfProblem("Word didn't start a separate copy of itself; close Word and try again")
            caption = f"merge-pdfs {uuid.uuid4().hex}"
            patiently(lambda: setattr(app, "Caption", caption))  # so we can find its window, and so its process
            self.process, self.pid = process_of_window("OpusApp", caption)
            patiently(lambda: setattr(app, "DisplayAlerts", 0))       # wdAlertsNone
            patiently(lambda: setattr(app, "AutomationSecurity", 3))  # msoAutomationSecurityForceDisable: no macros
            self.update_links = patiently(lambda: app.Options.UpdateLinksAtOpen)
            # Linked content (a pasted-link table, say) as last saved, not fetched afresh -- possibly
            # from unsaved edits open in another program. Word keeps this setting, so close() puts it back.
            patiently(lambda: setattr(app.Options, "UpdateLinksAtOpen", False))
        except (PdfProblem, pywintypes.com_error) as e:
            self.app = app
            self.close()  # don't leave a half-started copy running, invisible, until you log off
            self.start_failure = str(e) if isinstance(e, PdfProblem) else f"Word couldn't start ({word_error_text(e)})"
            raise PdfProblem(self.start_failure) from None
        self.app = app

    def convert(self, source: Path, pdf: Path) -> None:
        if self.app is None:
            self._start()
        watch = WordWatch(self.process, self.pid, self.timeout)
        doc = None
        try:
            doc = self.app.Documents.Open(FileName=str(source), **WORD_OPEN_OPTIONS)
            doc.ExportAsFixedFormat(OutputFileName=str(pdf), ExportFormat=17,  # wdExportFormatPDF
                                    OpenAfterExport=False, OptimizeFor=0,       # wdExportOptimizeForPrint
                                    CreateBookmarks=1,                          # wdExportCreateHeadingBookmarks
                                    DocStructureTags=True)
        except Exception as e:
            watch.stop()
            if watch.stuck:
                raise PdfProblem(f"Word stopped responding while converting it "
                                 f"(gave up after {self.timeout:g} seconds)", retry=False) from None
            if watch.questions:
                raise PdfProblem("Word wanted to ask something before converting it. Open it in Word once, "
                                 "answer the question and close it, then try again") from None
            if not self._alive():
                raise PdfProblem("Word closed unexpectedly while converting it") from None
            raise PdfProblem(word_problem(e)) from None
        finally:
            watch.stop()
            if watch.stuck or not self._alive():
                self._forget()  # it's gone; the next file gets a fresh copy of Word
            elif doc is not None:
                try:
                    doc.Close(SaveChanges=0)  # wdDoNotSaveChanges
                except Exception:
                    pass

    def _alive(self) -> bool:
        try:
            self.app.Documents.Count
            return True
        except Exception:
            return False

    def _forget(self) -> None:
        """Let go of a copy of Word that died or was shut down. (A copy shut down this way saves
        nothing, so Word's settings are as they were.)"""
        self.app = None
        if self.process:
            terminate_process(self.process)  # in case it's only half gone
            close_process_handle(self.process)
            self.process, self.pid = None, 0

    def close(self) -> None:
        app, self.app = self.app, None
        if app is not None:
            try:
                if self.update_links is not None:
                    app.Options.UpdateLinksAtOpen = self.update_links  # Word saves its settings as it quits
                app.Quit(SaveChanges=0)
            except Exception:
                pass
            del app
        if self.process:
            # Word can take a few seconds to exit after Quit; make sure our copy really goes,
            # without holding anyone up.
            threading.Thread(target=reap_process, args=(self.process,), daemon=True).start()
            self.process, self.pid = None, 0


class WordWatch:
    """Watches our hidden Word while it works on one document. A question box it puts up can't be
    seen, let alone answered, so it's closed (the same as clicking Cancel), which makes the
    document fail at once. If Word still hasn't finished after `timeout` seconds, it's shut down.
    (Shutting Word down while a question is showing would make Word distrust that file next time;
    closing the question first avoids that.) Without a process to watch, it does nothing."""

    QUESTION_CLASSES = ("#32770", "NUIDialog")  # standard and Office message boxes; plus Word's "bosa_sdm..." dialogs

    def __init__(self, process, pid: int, timeout: float):
        self.stuck = False
        self.questions = 0
        self._done = threading.Event()
        self._thread = None
        if process:
            self._thread = threading.Thread(target=self._watch, args=(process, pid, time.monotonic() + timeout),
                                            daemon=True)
            self._thread.start()

    def _watch(self, process, pid: int, deadline: float) -> None:
        while not self._done.wait(0.25):
            for hwnd, window_class in visible_windows_of(pid):
                if window_class in self.QUESTION_CLASSES or window_class.startswith("bosa_sdm"):
                    self.questions += 1
                    close_window(hwnd)
            if time.monotonic() >= deadline:
                self.stuck = True
                terminate_process(process)
                return

    def stop(self) -> None:
        self._done.set()
        if self._thread:
            self._thread.join(timeout=5)


def patiently(call, seconds: float = 10):
    """call(), retried while Word says it's busy -- as a copy that has only just started can be."""
    import pywintypes
    deadline = time.monotonic() + seconds
    while True:
        try:
            return call()
        except pywintypes.com_error as e:
            if e.hresult not in (-2147418111, -2147417846) or time.monotonic() > deadline:  # RPC_E_CALL_REJECTED,
                raise                                                                        # RPC_E_SERVERCALL_RETRYLATER
            time.sleep(0.25)


def word_error_text(error: Exception) -> str:
    """Word's own description of an automation error, if it gave one."""
    info = getattr(error, "excepinfo", None)
    text = str(info[2]).strip() if info and len(info) > 2 and info[2] else ""
    return text or getattr(error, "strerror", None) or str(error)


def word_problem(error: Exception) -> str:
    """Plain words for why Word couldn't convert a file."""
    text = word_error_text(error)
    if "password" in text.lower():
        return "needs a password to open"
    return f"Word couldn't convert it ({text})"


@functools.cache
def _win32():
    """user32 and kernel32, with the signatures used below (handles are 64-bit)."""
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.WinDLL("user32")
    kernel32 = ctypes.WinDLL("kernel32")
    user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
    user32.FindWindowW.restype = wintypes.HWND
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.EnumWindows.argtypes = [ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM), wintypes.LPARAM]
    user32.IsWindowVisible.argtypes = [wintypes.HWND]
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel32.WaitForSingleObject.restype = wintypes.DWORD
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    return user32, kernel32


def process_of_window(window_class: str, title: str):
    """(handle, process ID) of the process that owns the top-level window with this class and
    title, or (None, 0). Holding the handle keeps that process ID from being reused, so shutting
    the process down later is safe."""
    import ctypes
    from ctypes import wintypes
    user32, kernel32 = _win32()
    hwnd = user32.FindWindowW(window_class, title)
    pid = wintypes.DWORD()
    if not hwnd or not user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid)):
        return None, 0
    handle = kernel32.OpenProcess(0x0001 | 0x00100000, False, pid.value)  # PROCESS_TERMINATE | SYNCHRONIZE
    return (handle, pid.value) if handle else (None, 0)


def visible_windows_of(pid: int) -> list[tuple[int, str]]:
    """(window, class name) for each visible top-level window of process `pid`."""
    import ctypes
    from ctypes import wintypes
    user32 = _win32()[0]
    found = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def each(hwnd, _):
        owner = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
        if owner.value == pid and user32.IsWindowVisible(hwnd):
            name = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, name, 256)
            found.append((hwnd, name.value))
        return True

    user32.EnumWindows(each, 0)
    return found


def close_window(hwnd) -> None:
    _win32()[0].PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE: for a dialog, the same as Cancel


def terminate_process(handle) -> None:
    _win32()[1].TerminateProcess(handle, 1)


def close_process_handle(handle) -> None:
    _win32()[1].CloseHandle(handle)


def reap_process(handle, seconds: float = 30) -> None:
    """Give a process `seconds` to exit, shut it down if it hasn't, and let go of its handle."""
    kernel32 = _win32()[1]
    if kernel32.WaitForSingleObject(handle, int(seconds * 1000)) == 0x102:  # WAIT_TIMEOUT
        kernel32.TerminateProcess(handle, 1)
    kernel32.CloseHandle(handle)


CONVERTERS: list[Converter] = [WordConverter()]


def converter_for(path: Path) -> Converter | None:
    extension = path.suffix.lower()
    return next((c for c in CONVERTERS if extension in c.extensions), None)


def supported_extensions() -> set[str]:
    return {".pdf", *(e for c in CONVERTERS for e in c.extensions)}


def kinds_text(conjunction: str = "or") -> str:
    """What can be combined, for messages: "PDFs or Word documents"."""
    names = ["PDFs", *(c.label for c in CONVERTERS)]
    return names[0] if len(names) == 1 else f"{', '.join(names[:-1])} {conjunction} {names[-1]}"


class Conversions:
    """PDF copies of files that aren't PDFs, made on demand in a temporary folder that's
    deleted when the program exits. A copy is remade when its file changes."""

    def __init__(self):
        self.folder: Path | None = None
        # file -> (its size and modified time when converted, its PDF copy or why there isn't one)
        self.results: dict[Path, tuple[tuple[int, int] | None, Path | str]] = {}
        self.gave_up: set[Path] = set()  # files whose failure isn't worth retrying until they change

    @staticmethod
    def _stamp(path: Path) -> tuple[int, int]:
        info = path.stat()
        return info.st_mtime_ns, info.st_size

    def _current(self, path: Path) -> Path | str | None:
        """The recorded result for `path` if the file hasn't changed since, else None."""
        entry = self.results.get(path.resolve())
        try:
            return entry[1] if entry and entry[0] == self._stamp(path) else None
        except OSError:
            return None

    def _needs_converting(self, path: Path) -> bool:
        result = self._current(path)
        if isinstance(result, Path):
            return False
        return not (result is not None and path.resolve() in self.gave_up)

    def forget_failures(self) -> None:
        """Try every failed file again next time, even those not worth retrying unchanged."""
        self.gave_up.clear()

    def convert(self, files: list[Path], progress: Callable[[int, int, Path], None] | None = None) -> None:
        """Make PDF copies of whichever `files` need one: those that aren't PDFs and are new,
        changed since their copy, or failed last time (unless that failure isn't worth
        retrying). Files of one kind share one converter session (one start of Word, say).
        `progress(i, total, file)` is called before each file is converted."""
        todo = [f for f in files if f.suffix.lower() != ".pdf" and self._needs_converting(f)]
        if not todo:
            return
        if self.folder is None:
            self.folder = Path(tempfile.mkdtemp(prefix="merge_pdfs_"))
            atexit.register(shutil.rmtree, self.folder, ignore_errors=True)
        groups: dict[Converter | None, list[Path]] = {}
        for f in todo:
            groups.setdefault(converter_for(f), []).append(f)
        done = 0
        for converter, group in groups.items():
            why = converter.unavailable() if converter else "this kind of file can't be converted to PDF"
            if why:
                for f in group:
                    self._record(f, self._stamp_or_none(f), why)
                done += len(group)
                continue
            pending = list(group)
            try:
                with converter.session() as convert_one:
                    while pending:
                        f = pending.pop(0)
                        if progress:
                            progress(done, len(todo), f)
                        done += 1
                        stamp = self._stamp_or_none(f)  # before converting, so a save mid-way means "changed"
                        if stamp is None:
                            self._record(f, None, "couldn't read the file")
                            continue
                        pdf = self.folder / f"{uuid.uuid4().hex}.pdf"
                        try:
                            # Absolute paths: a converter may run in another process (Word does),
                            # with a different current folder.
                            convert_one(f.absolute(), pdf)
                        except Exception as e:
                            pdf.unlink(missing_ok=True)
                            problem = e if isinstance(e, PdfProblem) else PdfProblem(
                                f"couldn't convert it ({str(e) or type(e).__name__})")
                            self._record(f, stamp, str(problem), retry=problem.retry)
                        else:
                            self._record(f, stamp, pdf)
            except Exception as e:  # the converter couldn't get going (or finish up) at all
                why = str(e) if isinstance(e, PdfProblem) else f"couldn't convert {converter.label} ({str(e) or type(e).__name__})"
                for f in pending:
                    self._record(f, self._stamp_or_none(f), why)
                done += len(pending)

    def _stamp_or_none(self, path: Path) -> tuple[int, int] | None:
        try:
            return self._stamp(path)
        except OSError:
            return None

    def _record(self, path: Path, stamp, result: Path | str, retry: bool = True) -> None:
        key = path.resolve()
        old = self.results.get(key)
        if old and isinstance(old[1], Path) and old[1] != result:
            old[1].unlink(missing_ok=True)  # the copy it replaces
        self.results[key] = (stamp, result)
        if isinstance(result, str) and not retry:
            self.gave_up.add(key)
        else:
            self.gave_up.discard(key)

    def pdf_for(self, path: Path) -> Path:
        """The PDF to merge for `path`: the file itself if it's a PDF, otherwise its current PDF
        copy (made now if there isn't one). Raises PdfProblem if there can't be one."""
        if path.suffix.lower() == ".pdf":
            return path
        result = self._current(path)
        if result is None:
            self.convert([path])
            result = self.results[path.resolve()][1]
        if isinstance(result, Path):
            return result
        raise PdfProblem(result)


def file_types() -> list[tuple[str, str]]:
    """Choices for the "Add files" dialog's file-type list."""
    kinds = [("PDF files", "*.pdf")] + [(c.label, " ".join(f"*{e}" for e in c.extensions)) for c in CONVERTERS]
    return [(f"{kinds_text()}", " ".join(pattern for _, pattern in kinds))] + kinds


# ---------------------------------------------------------------------------
# Command line


def run_cli(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description=f"Combine the {kinds_text('and')} in a folder into one PDF. "
                    "Run with no arguments to open a window instead.")
    parser.add_argument("folder", nargs="?", default=".", help="folder containing the files (default: current directory)")
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
    # Any other file already in FOLDER stays an input, so merge() refuses to overwrite it.
    files = list_documents(folder, exclude=output if output.name.lower() == "merged.pdf" else None)
    # Leave out earlier combined packets too, so an old packet saved in FOLDER isn't combined in
    # again, and the older form of a document saved in two (Minutes.docx and its Minutes.pdf).
    left_out = {p: "looks like an earlier combined file" for p in files if looks_combined(p, files)}
    left_out |= same_document_twins([p for p in files if p not in left_out])
    for p, why in left_out.items():
        print(f"  left out {p.name}: {why}", file=sys.stderr)
    files = [p for p in files if p not in left_out]
    if not files:
        print(f"No {kinds_text()} found in {folder}", file=sys.stderr)
        return 1
    for p in open_in_word(files):
        print(f"  note: {p.name} is open in Word; changes not saved there won't be included", file=sys.stderr)

    try:
        merged, skipped = merge(files, output, progress=lambda i, total, f: print(f"  converting {f.name} ...", flush=True))
    except PermissionError:
        # On Windows this usually means the output PDF is open in a viewer.
        print(f"Can't write {output}: permission denied (is it open in another program?)", file=sys.stderr)
        return 1
    except ValueError as e:
        print(e, file=sys.stderr)
        return 1

    for path, pages in merged:
        print(f"  + {path.name} ({plural(pages, 'page')})")
    for path, reason in skipped:
        print(f"  skipped {path.name}: {reason}", file=sys.stderr)
    if not merged:
        print("No readable files to merge.", file=sys.stderr)
        return 1

    total = sum(pages for _, pages in merged)
    print(f"Wrote {output}: {plural(total, 'page')} from {len(merged)} of {plural(len(files), 'file')}")
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
    """Window for choosing a folder, putting its files in order, and combining them."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.folder: Path | None = None
        self.paths: dict[str, Path] = {}        # list row id -> file
        self.pages: dict[str, int | None] = {}  # list row id -> page count, None if it can't be included
        self.reasons: dict[str, str] = {}       # list row id -> why it can't be included
        self.left_out: dict[Path, str] = {}     # files in the folder not listed, and why
        self.open_elsewhere: list[Path] = []    # listed documents that look open in Word
        self.conversions = Conversions()        # PDF copies of the Word documents (and so on) listed
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
            ("Add files…", self.add_files),
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
        chosen = filedialog.askdirectory(parent=self.root, title="Choose the folder with your files",
                                         initialdir=str(self.folder or Path.cwd()), mustexist=True)
        if chosen:
            self.load_folder(Path(chosen))

    def load_folder(self, folder: Path) -> None:
        """List `folder`'s files in natural filename order, replacing the current list."""
        # Read the folder first, so a folder that's gone or unreadable leaves the current list alone.
        try:
            files = list_documents(folder)
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
        # Leave out earlier combined packets so they aren't combined in a second time (merged.pdf
        # from the command line, "<folder name>.pdf", or any PDF whose bookmarks name the other
        # files), and the older form of a document saved in two (Minutes.docx and Minutes.pdf).
        own = {"merged.pdf", f"{folder.name}.pdf".lower()}
        with self._busy():
            self.left_out = {p: "looks like an earlier combined file" for p in files
                             if p.name.lower() in own or looks_combined(p, files)}
            self.left_out |= same_document_twins([p for p in files if p not in self.left_out])
            self._add([p for p in files if p not in self.left_out])

    def reload(self) -> None:
        if self.folder:
            self.conversions.forget_failures()  # an explicit "try again"
            self.load_folder(self.folder)

    def add_files(self) -> None:
        chosen = filedialog.askopenfilenames(parent=self.root, title="Add files",
                                             initialdir=str(self.folder or Path.cwd()), filetypes=file_types())
        if chosen:
            with self._busy():
                self._add([Path(p) for p in chosen])

    def _add(self, files: list[Path]) -> None:
        """Append files to the end of the list, skipping any already on it. Files that aren't
        PDFs are converted now, so their page counts (or problems) show straight away."""
        listed = {p.resolve() for p in self.paths.values()}
        new = []
        for path in files:
            if path.resolve() not in listed:
                listed.add(path.resolve())
                new.append(path)
        self.conversions.convert(new, self._show_progress)
        for path in new:
            iid = f"row{self._next_id}"
            self._next_id += 1
            try:
                pages = len(open_pdf(self.conversions.pdf_for(path)).pages)
            except PdfProblem as e:
                pages = None
                self.reasons[iid] = str(e)
            self.paths[iid] = path
            self.pages[iid] = pages
            name = path.name if path.parent == self.folder else f"{path.name}   (from {path.parent})"
            self.tree.insert("", "end", iid=iid, values=("", name, "can't open" if pages is None else pages),
                             tags=("bad",) if pages is None else ())
        self.open_elsewhere = open_in_word(list(self.paths.values()))
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
        self.open_elsewhere = [p for p in self.open_elsewhere if p in self.paths.values()]
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
        files = [self.paths[iid] for iid in self.tree.get_children()]
        if not files:
            return
        # Word documents still open in Word: unsaved changes there won't make it into the PDF.
        self.open_elsewhere = open_in_word(files)
        if self.open_elsewhere:
            names = "\n".join(f"  • {p.name}" for p in self.open_elsewhere)
            if not messagebox.askyesno(
                    "Open in Word", f"These are open in Word:\n{names}\n\nChanges you haven't saved there won't be "
                    "in the PDF. Save them in Word first if you need to.\n\nCombine now?", parent=self.root):
                self._update_status()
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
                # Word documents edited since they were listed are converted again here.
                merged, skipped = merge(files, output, self.conversions, self._show_progress)
            self._show_results(merged, skipped)
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
        finally:
            self._update_status()  # replace any "Converting..." message

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

    def _show_results(self, merged: list[tuple[Path, int]], skipped: list[tuple[Path, str]]) -> None:
        """Update the rows from what Combine found: a file edited since it was listed may now
        have a different page count, or be included or skipped where it wasn't before."""
        pages, reasons = dict(merged), dict(skipped)
        for iid, path in self.paths.items():
            if path in pages:
                self.pages[iid] = pages[path]
                self.reasons.pop(iid, None)
            elif path in reasons:
                self.pages[iid] = None
                self.reasons[iid] = reasons[path]
            else:
                continue
            bad = self.pages[iid] is None
            self.tree.set(iid, "pages", "can't open" if bad else self.pages[iid])
            self.tree.item(iid, tags=("bad",) if bad else ())

    def _show_progress(self, i: int, total: int, path: Path) -> None:
        self.status.configure(text=f"Converting {path.name} ({i + 1} of {total})…")
        self.root.update_idletasks()

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
            text = "No files listed." if self.folder else f"Choose a folder to list its {kinds_text('and')}."
        if self.open_elsewhere:
            names = ", ".join(p.name for p in self.open_elsewhere)
            text += f"  ·  Open in Word: {names} (changes not saved there won't be included)"
        listed = set(self.paths.values())
        notes: dict[str, list[str]] = {}
        for p, why in self.left_out.items():
            if p not in listed:
                notes.setdefault(why, []).append(p.name)
        if notes:
            text += "  ·  " + "; ".join(f"Left out {', '.join(names)}: {why}" for why, names in notes.items())
            text += " (Add files… to include it)"
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
