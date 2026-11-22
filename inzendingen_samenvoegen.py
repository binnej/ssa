# -*- coding: utf-8 -*-
"""
Inzendingen samenvoegen tot PDF / Merge student submissions into one PDF
========================================================================
Reads a zip file with student submissions and creates ONE pdf with:
  - the newest (or all) submissions of each student (Word via Microsoft Word, or PDF),
  - a small header on top (name, class, number of submissions, page number, custom field),
    with a border on the first page of each submission,
  - optional: every submission starts on an odd page (recto-verso),
  - optional: an empty page (with header) for students with an empty folder,
  - bookmarks per class and per student.

Everything happens locally. Nothing is sent anywhere.
Temporary files are deleted afterwards.
"""
import io
import os
import re
import shutil
import sys
import tempfile
import threading
import queue
import zipfile
from collections import Counter

from pypdf import PdfReader, PdfWriter, Transformation
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# --------------------------------------------------------------------------
# Texts / translations (add a language here if you want)
# --------------------------------------------------------------------------
STR = {
    "nl": {
        "app_title": "Inzendingen samenvoegen tot PDF",
        # header
        "h_name": "Naam", "h_class": "Klas", "h_count": "Aantal inzendingen",
        "h_multiple": "Meerdere documenten ingediend",          # vet, zoals gevraagd
        "h_failed": "Document kon niet verwerkt worden",
        "h_nodoc": "Geen bruikbaar document gevonden",
        "h_file": "Bestand %d/%d",
        # bookmarks
        "bm_multi": " (%d bestanden)", "bm_file": "Bestand %d/%d: %s",
        # log / status
        "no_students": "Geen studentenmappen gevonden in het zipbestand.",
        "loose": "Los bestand buiten een studentenmap genegeerd: %s",
        "multi": "%s: %d bestanden ingediend.",
        "ignored": "%s: genegeerd (geen %s): %s",
        "ok": "[%d/%d] %s: %s (%s)",
        "fail": "%s: %s kon niet verwerkt worden (%s)",
        "nousable": "%s: geen bruikbaar document (%d bestand(en), geen %s).",
        "emptyfolder": "%s: lege map.",
        "pw": "beveiligd met een wachtwoord", "nopages": "document bevat geen pagina's",
        "word_only_win": "Word-modus werkt enkel op Windows met Microsoft Word.",
        "word_missing": "Microsoft Word kon niet gestart worden. Is Word geïnstalleerd op deze computer?",
        "word_starting": "Word wordt gestart... even geduld, dit kan tot een minuut duren.",
        "word_first": "Eerste document wordt omgezet... Word is traag bij de start, even geduld.",
        "working": "Bezig: %d van %d",
        "saved_as": "Opgeslagen als: %s",
        "summary": "Klaar: %d studenten, %d pagina's. Lege mappen: %d, meerdere bestanden: %d, mislukt: %d.",
        "cancelled": "Geannuleerd.", "error": "Fout: %s",
        "perm": "Kan het PDF-bestand niet opslaan. Is het nog geopend in een ander programma?",
        # GUI
        "language": "Taal:",
        "step1": "1. Kies het zipbestand met de inzendingen", "browse": "Bladeren...",
        "step2": "2. Opties",
        "type": "Type inzendingen:", "word": "Word (doc, docx, odt)", "pdf": "PDF",
        "structure": "Mapstructuur:",
        "st_auto": "Automatisch herkennen", "st_yes": "Met klassen (klas/student)",
        "st_no": "Zonder klassen (student)",
        "rv": "Recto-verso: elke inzending begint op een oneven pagina",
        "empty": "Lege pagina met koptekst voor studenten met een lege map",
        "allfiles": "Alle inzendingen van een student opnemen (niet enkel de nieuwste)",
        "noheader": "Geen koptekst (hoofding) toevoegen",
        "hdrall": "Koptekst op elke pagina (anders enkel op de eerste pagina)",
        "pagenum": "Paginanummer in koptekst:",
        "pn_sub": "per inzending (1 / 4)", "pn_all": "in de hele PDF", "pn_none": "geen",
        "custom": "Extra veld in koptekst:",
        "custom_hint": "Bv. Score: ____ / 20  (wordt afgekapt als het niet past)",
        "step3": "3. PDF maken", "make": "PDF maken", "cancel": "Annuleren", "open": "PDF openen",
        "privacy": "Alle verwerking gebeurt lokaal op deze computer; er wordt niets verstuurd.",
        "pick_zip": "Kies het zipbestand", "zip_files": "Zipbestand", "all_files": "Alle bestanden",
        "need_zip": "Kies eerst een zipbestand.", "save_as": "Opslaan als",
        "cant_read": "Kan het zipbestand niet lezen: %s", "none_found": "Geen studentenmappen gevonden.",
        "found_cls": "Gevonden: %d klassen, %d studenten",
        "found_nocls": "Gevonden: %d studenten (zonder klassen)",
        "found_tail": " | lege mappen: %d | meerdere bestanden: %d",
    },
    "en": {
        "app_title": "Merge student submissions into one PDF",
        "h_name": "Name", "h_class": "Class", "h_count": "Submissions",
        "h_multiple": "Multiple documents submitted",
        "h_failed": "Document could not be processed",
        "h_nodoc": "No usable document found",
        "h_file": "File %d/%d",
        "bm_multi": " (%d files)", "bm_file": "File %d/%d: %s",
        "no_students": "No student folders found in the zip file.",
        "loose": "Loose file outside a student folder ignored: %s",
        "multi": "%s: %d files submitted.",
        "ignored": "%s: ignored (not %s): %s",
        "ok": "[%d/%d] %s: %s (%s)",
        "fail": "%s: %s could not be processed (%s)",
        "nousable": "%s: no usable document (%d file(s), none of type %s).",
        "emptyfolder": "%s: empty folder.",
        "pw": "password protected", "nopages": "document has no pages",
        "word_only_win": "Word mode only works on Windows with Microsoft Word.",
        "word_missing": "Microsoft Word could not be started. Is Word installed on this computer?",
        "word_starting": "Starting Word... please wait, this can take up to a minute.",
        "word_first": "Converting the first document... Word is slow at start-up, please wait.",
        "working": "Working: %d of %d",
        "saved_as": "Saved as: %s",
        "summary": "Done: %d students, %d pages. Empty folders: %d, multiple files: %d, failed: %d.",
        "cancelled": "Cancelled.", "error": "Error: %s",
        "perm": "Cannot save the PDF file. Is it still open in another program?",
        "language": "Language:",
        "step1": "1. Choose the zip file with the submissions", "browse": "Browse...",
        "step2": "2. Options",
        "type": "Type of submissions:", "word": "Word (doc, docx, odt)", "pdf": "PDF",
        "structure": "Folder structure:",
        "st_auto": "Detect automatically", "st_yes": "With classes (class/student)",
        "st_no": "Without classes (student)",
        "rv": "Duplex: every submission starts on an odd page",
        "empty": "Empty page with header for students with an empty folder",
        "allfiles": "Include all submissions of a student (not only the newest)",
        "noheader": "No header at all",
        "hdrall": "Header on every page (otherwise only on the first page)",
        "pagenum": "Page number in header:",
        "pn_sub": "per submission (1 / 4)", "pn_all": "in the whole PDF", "pn_none": "none",
        "custom": "Extra field in header:",
        "custom_hint": "E.g. Score: ____ / 20  (chopped off if it does not fit)",
        "step3": "3. Create the PDF", "make": "Create PDF", "cancel": "Cancel", "open": "Open PDF",
        "privacy": "All processing happens locally on this computer; nothing is sent anywhere.",
        "pick_zip": "Choose the zip file", "zip_files": "Zip file", "all_files": "All files",
        "need_zip": "Please choose a zip file first.", "save_as": "Save as",
        "cant_read": "Cannot read the zip file: %s", "none_found": "No student folders found.",
        "found_cls": "Found: %d classes, %d students",
        "found_nocls": "Found: %d students (no classes)",
        "found_tail": " | empty folders: %d | multiple files: %d",
    },
}
_lang = "en"


def set_lang(lang):
    global _lang
    _lang = lang if lang in STR else "en"


def tr(key, *args):
    s = STR[_lang].get(key) or STR["en"][key]
    return s % args if args else s


def detect_language():
    """Dutch if Windows (UI language or region) or the system locale is Dutch; else English."""
    try:
        if sys.platform == "win32":
            import ctypes
            k = ctypes.windll.kernel32
            if (k.GetUserDefaultUILanguage() & 0x3FF) == 0x13 or \
               (k.GetUserDefaultLCID() & 0x3FF) == 0x13:      # 0x13 = LANG_DUTCH
                return "nl"
            return "en"
    except Exception:
        pass
    try:
        import locale
        loc = locale.getlocale()[0] or os.environ.get("LANG", "") or ""
        if loc.lower().startswith("nl"):
            return "nl"
    except Exception:
        pass
    return "en"


# --------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------
MM = 72.0 / 25.4
A4 = (595.2756, 841.8898)
WORD_EXT = {".doc", ".docx", ".docm", ".odt", ".rtf"}
PDF_EXT = {".pdf"}
JUNK_FILES = {".ds_store", "thumbs.db", "desktop.ini"}

HEADER_FONT_SIZE = 9.0      # pt (smaller for very long names)
HEADER_TOP = 5 * MM         # distance from top edge of page to top of header box
HEADER_LEFT = 8 * MM
HEADER_SEP = "   |   "


class Cancelled(Exception):
    pass


# --------------------------------------------------------------------------
# Reading the zip and detecting the structure
# --------------------------------------------------------------------------
class FileRef:
    __slots__ = ("info", "path", "ext")

    def __init__(self, info, path):
        self.info = info
        self.path = path
        self.ext = os.path.splitext(path)[1].lower()


class Student:
    def __init__(self, cls, name):
        self.cls = cls
        self.name = name
        self.files = []


class ScanResult:
    def __init__(self):
        self.students = []
        self.loose_files = []
        self.depth = 0
        self.has_classes = False

    @property
    def n_classes(self):
        return len({s.cls for s in self.students if s.cls})


def fix_name(info):
    """Repair names with accents from zips without the UTF-8 flag."""
    name = info.filename
    if not (info.flag_bits & 0x800):
        try:
            name = name.encode("cp437").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return name.replace("\\", "/")


def clean_student_name(name):
    # Moodle exports: "Jan Peeters_123456_assignsubmission_file_"
    name = re.sub(r"_\d+_assignsubmission_\w*$", "", name)
    return name.strip()


def is_junk(parts):
    last = parts[-1]
    return (
        "__MACOSX" in parts
        or last.lower() in JUNK_FILES
        or last.startswith("~$")
        or last.startswith("._")
    )


def scan_zip(zf, classes_option="auto"):
    """classes_option: 'auto' | 'yes' | 'no'"""
    items = []
    for info in zf.infolist():
        name = fix_name(info)
        is_dir = name.endswith("/")
        parts = [p for p in name.split("/") if p]
        if not parts or is_junk(parts):
            continue
        items.append((parts, is_dir, info))

    # Which folders are "student folders"? Folders that directly contain files,
    # or empty leaf folders. The most common depth wins.
    file_dirs, all_dirs, parent_dirs = set(), set(), set()
    for parts, is_dir, _ in items:
        if is_dir:
            all_dirs.add(tuple(parts))
        elif len(parts) > 1:
            file_dirs.add(tuple(parts[:-1]))
    for d in list(all_dirs) + list(file_dirs):
        for k in range(1, len(d)):
            parent_dirs.add(d[:k])
    leaf_dirs = {d for d in all_dirs if d not in parent_dirs}
    candidates = file_dirs | leaf_dirs

    res = ScanResult()
    if not candidates:
        return res
    depth = Counter(len(d) for d in candidates).most_common()
    best = max(c for _, c in depth)
    res.depth = min(d for d, c in depth if c == best)
    D = res.depth

    if classes_option == "auto":
        res.has_classes = D >= 2
    else:
        res.has_classes = (classes_option == "yes") and D >= 2

    by_key = {}   # order = order of first appearance in the zip
    for parts, is_dir, info in items:
        if len(parts) <= D and (is_dir or len(parts) - 1 < D):
            if not is_dir:
                res.loose_files.append("/".join(parts))
            if not (is_dir and len(parts) == D):
                continue
        key = tuple(parts[:D])
        st = by_key.get(key)
        if st is None:
            cls = key[D - 2] if res.has_classes else ""
            st = Student(cls, clean_student_name(key[D - 1]))
            by_key[key] = st
            res.students.append(st)
        if not is_dir:
            st.files.append(FileRef(info, "/".join(parts)))
    return res


def pick_files(student, exts, all_files):
    """Newest suitable file, or (all_files) all suitable files, oldest first."""
    cands = [(f.info.date_time, idx, f) for idx, f in enumerate(student.files) if f.ext in exts]
    cands.sort(key=lambda t: (t[0], t[1]))
    files = [t[2] for t in cands]
    if not files:
        return []
    return files if all_files else [files[-1]]


# --------------------------------------------------------------------------
# Header (overlay at the top of the page)
# --------------------------------------------------------------------------
_fonts_registered = None


def _register_fonts():
    """Returns (regular, bold) of a Unicode font, or None."""
    global _fonts_registered
    if _fonts_registered is not None:
        return _fonts_registered or None
    win = os.environ.get("WINDIR", r"C:\Windows")
    pairs = [
        (os.path.join(win, "Fonts", "arial.ttf"), os.path.join(win, "Fonts", "arialbd.ttf")),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
         "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
    ]
    for reg, bold in pairs:
        if os.path.exists(reg) and os.path.exists(bold):
            try:
                pdfmetrics.registerFont(TTFont("HdrReg", reg))
                pdfmetrics.registerFont(TTFont("HdrBold", bold))
                _fonts_registered = ("HdrReg", "HdrBold")
                return _fonts_registered
            except Exception:
                pass
    _fonts_registered = ()
    return None


def pick_fonts(segments, custom=""):
    """Helvetica (small file) unless there are characters it does not know."""
    text = "".join(t for t, _ in segments) + custom
    try:
        text.encode("cp1252")
        return ("Helvetica", "Helvetica-Bold")
    except UnicodeEncodeError:
        return _register_fonts() or ("Helvetica", "Helvetica-Bold")


def header_segments(student, n_files, file_no=None, file_total=1, extra=None):
    segs = [("%s: %s" % (tr("h_name"), student.name), False)]
    if student.cls:
        segs.append((HEADER_SEP + "%s: %s" % (tr("h_class"), student.cls), False))
    segs.append((HEADER_SEP + "%s: " % tr("h_count"), False))
    segs.append((str(n_files), n_files > 1))
    if n_files > 1:
        segs.append(("   " + tr("h_multiple"), True))
    if file_no is not None:
        segs.append((HEADER_SEP + tr("h_file", file_no, file_total), False))
    if extra:
        segs.append((HEADER_SEP + extra, True))
    return segs


def make_overlay(w, h, segments, custom, fonts, border):
    reg, bold = fonts
    if reg == "Helvetica":  # Helvetica only knows cp1252
        segments = [(t.encode("cp1252", "replace").decode("cp1252"), b) for t, b in segments]
        custom = custom.encode("cp1252", "replace").decode("cp1252")

    def width(segs, sz):
        return sum(pdfmetrics.stringWidth(t, bold if b else reg, sz) for t, b in segs)

    size = HEADER_FONT_SIZE
    maxw = w - 2 * HEADER_LEFT
    while width(segments, size) > maxw and size > 5:
        size -= 0.25

    # custom field: appended at the end and chopped off if it does not fit
    segs = list(segments)
    if custom:
        remaining = maxw - width(segments, size)
        txt = HEADER_SEP + custom
        while txt and pdfmetrics.stringWidth(txt, reg, size) > remaining:
            txt = txt[:-1]
        if len(txt) > len(HEADER_SEP.rstrip()) + 1:
            segs.append((txt, False))

    tw = width(segs, size)
    pad = 1.5 * MM
    band_h = size + 2 * pad
    top = h - HEADER_TOP

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    c.setFillColorRGB(1, 1, 1)      # white background, only around the text
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.9)
    c.rect(HEADER_LEFT - pad, top - band_h, tw + 2 * pad, band_h,
           stroke=1 if border else 0, fill=1)
    c.setFillColorRGB(0, 0, 0)
    x, y = HEADER_LEFT, top - pad - size * 0.8
    for text, b in segs:
        font = bold if b else reg
        c.setFont(font, size)
        c.drawString(x, y, text)
        x += pdfmetrics.stringWidth(text, font, size)
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def stamp_page(page, segments, custom, fonts, border):
    try:
        if getattr(page, "rotation", 0) % 360:
            page.transfer_rotation_to_content()
    except Exception:
        pass
    box = page.cropbox
    left, bottom = float(box.left), float(box.bottom)
    w, h = float(box.width), float(box.height)
    ov = make_overlay(w, h, segments, custom, fonts, border)
    page.merge_transformed_page(ov, Transformation().translate(left, bottom))


def stamp_all(writer, jobs, opts):
    """jobs: (page_index, base_segments, is_first_page, local_no, local_total)"""
    total = len(writer.pages)
    custom = (opts.get("custom_text") or "").strip()
    mode = opts.get("page_numbers", "sub")
    for idx, segs, is_first, li, lt in jobs:
        s = list(segs)
        if mode == "sub":
            s.append((HEADER_SEP + "%d / %d" % (li, lt), False))
        elif mode == "all":
            s.append((HEADER_SEP + "%d / %d" % (idx + 1, total), False))
        stamp_page(writer.pages[idx], s, custom, pick_fonts(s, custom), border=is_first)


# --------------------------------------------------------------------------
# Word -> PDF through Microsoft Word (installed locally)
# --------------------------------------------------------------------------
class WordConverter:
    def __init__(self):
        if sys.platform != "win32":
            raise RuntimeError(tr("word_only_win"))
        try:
            import win32com.client
            self.word = win32com.client.DispatchEx("Word.Application")  # own instance
        except Exception:
            raise RuntimeError(tr("word_missing"))
        self.word.Visible = False
        self.word.DisplayAlerts = 0
        try:
            self.word.AutomationSecurity = 3     # never run macros
        except Exception:
            pass

    def convert(self, src, dst):
        doc = None
        try:
            # Positional: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
            # PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument.
            # A dummy password makes protected documents fail immediately instead
            # of showing an (invisible) password prompt.
            doc = self.word.Documents.Open(src, False, True, False, "x", "x", False, "x")
            try:
                doc.ExportAsFixedFormat(dst, 17)     # 17 = wdExportFormatPDF
            except Exception:
                doc.SaveAs2(dst, 17)
        finally:
            if doc is not None:
                try:
                    doc.Close(0)                      # 0 = do not save
                except Exception:
                    pass
        return dst

    def close(self):
        try:
            self.word.Quit(0)
        except Exception:
            pass


# --------------------------------------------------------------------------
# Main logic
# --------------------------------------------------------------------------
def _extract(zf, fref, tmpdir, idx):
    dst = os.path.join(tmpdir, "%d%s" % (idx, fref.ext))
    with zf.open(fref.info) as src, open(dst, "wb") as out:
        shutil.copyfileobj(src, out)
    return dst


def _silent_remove(*paths):
    for p in paths:
        try:
            os.remove(p)
        except OSError:
            pass


def build_pdf(zip_path, out_path, opts, log, progress, cancelled,
              status=None, converter_factory=None):
    """
    opts: mode ('word'|'pdf'), classes ('auto'|'yes'|'no'), recto_verso, empty_pages,
          all_files, no_header, header_all_pages (bool), page_numbers ('sub'|'all'|'none'),
          custom_text (str), lang ('nl'|'en')
    log(level, text): level in 'info' | 'warn' | 'error'
    progress(current, total)
    status(text_or_None, busy): None = back to the normal "working x of y" text
    """
    if opts.get("lang"):
        set_lang(opts["lang"])
    status = status or (lambda text, busy=False: None)
    word_mode = opts["mode"] == "word"
    exts = WORD_EXT if word_mode else PDF_EXT
    stats = {"students": 0, "pages": 0, "multiple": 0, "empty": 0, "failed": 0}

    with zipfile.ZipFile(zip_path) as zf:
        scan = scan_zip(zf, opts.get("classes", "auto"))
        students = scan.students
        if not students:
            raise RuntimeError(tr("no_students"))
        for lf in scan.loose_files:
            log("warn", tr("loose", lf))

        tmpdir = tempfile.mkdtemp(prefix="inzendingen_")
        converter = None
        try:
            warm = False
            if word_mode:
                status(tr("word_starting"), True)
                converter = (converter_factory or WordConverter)()
                status(tr("word_first"), True)
                warm = True

            writer = PdfWriter()
            class_items = {}
            jobs = []
            total = len(students)

            def pad_to_odd():
                if opts["recto_verso"] and len(writer.pages) % 2 == 1:
                    writer.add_blank_page(A4[0], A4[1])

            for i, st in enumerate(students, 1):
                if cancelled():
                    raise Cancelled()
                progress(i - 1, total)
                if warm:                      # progress() resets the text; show it again
                    status(tr("word_first"), True)
                n = len(st.files)
                label = st.name + (" (%s)" % st.cls if st.cls else "")
                files = pick_files(st, exts, opts.get("all_files", False))
                m = len(files)

                if n > 1:
                    stats["multiple"] += 1
                    log("warn", tr("multi", label, n))
                unsuitable = [os.path.basename(f.path) for f in st.files if f.ext not in exts]
                if unsuitable and files:
                    log("warn", tr("ignored", label, "/".join(sorted(exts)), ", ".join(unsuitable)))
                if not files:
                    if n > 0:
                        stats["failed"] += 1
                        log("error", tr("nousable", label, n, "/".join(sorted(exts))))
                    else:
                        stats["empty"] += 1
                        log("warn", tr("emptyfolder", label))

                student_item = None
                added = False
                for k, f in enumerate(files or [None], 1):
                    pages_src, failure = None, None
                    if f is None:
                        if n == 0 and not opts["empty_pages"]:
                            break
                        failure = tr("h_nodoc") if n > 0 else None
                    else:
                        src = pdf = None
                        try:
                            src = _extract(zf, f, tmpdir, i * 1000 + k)
                            pdf = converter.convert(src, os.path.join(tmpdir, "%d_%d.pdf" % (i, k))) \
                                if converter else src
                            reader = PdfReader(pdf)
                            if reader.is_encrypted and not reader.decrypt(""):
                                raise ValueError(tr("pw"))
                            pages_src = list(reader.pages)
                            if not pages_src:
                                raise ValueError(tr("nopages"))
                            d = f.info.date_time
                            log("info", tr("ok", i, total, label, os.path.basename(f.path),
                                           "%04d-%02d-%02d %02d:%02d" % tuple(d[:5])))
                        except Exception as e:
                            pages_src = None
                            failure = tr("h_failed")
                            stats["failed"] += 1
                            log("error", tr("fail", label, os.path.basename(f.path), e))
                        finally:
                            _silent_remove(*(p for p in (src, pdf) if p))
                        if warm:
                            warm = False
                            status(None, False)

                    segs = header_segments(st, n, k if m > 1 else None, m, failure)
                    pad_to_odd()
                    first = len(writer.pages)
                    if pages_src:
                        for j, p in enumerate(pages_src):
                            writer.add_page(p)
                            if j == 0 or opts["header_all_pages"]:
                                jobs.append((len(writer.pages) - 1, segs, j == 0, j + 1, len(pages_src)))
                    else:
                        writer.add_blank_page(A4[0], A4[1])
                        jobs.append((len(writer.pages) - 1, segs, True, 1, 1))

                    if student_item is None:
                        parent = None
                        if st.cls:
                            if st.cls not in class_items:
                                class_items[st.cls] = writer.add_outline_item(st.cls, first)
                            parent = class_items[st.cls]
                        title = st.name + (tr("bm_multi", n) if n > 1 else "")
                        student_item = writer.add_outline_item(title, first, parent)
                    if m > 1:
                        writer.add_outline_item(tr("bm_file", k, m, os.path.basename(f.path)),
                                                first, student_item)
                    added = True
                if added:
                    stats["students"] += 1

            progress(total, total)
            if not opts.get("no_header"):
                stamp_all(writer, jobs, opts)
            try:
                writer.page_mode = "/UseOutlines"
            except Exception:
                pass
            with open(out_path, "wb") as fh:
                writer.write(fh)
            stats["pages"] = len(writer.pages)
        finally:
            if converter:
                converter.close()
            shutil.rmtree(tmpdir, ignore_errors=True)
    return stats


# --------------------------------------------------------------------------
# Window (Tkinter: included in Python and in the .exe)
# --------------------------------------------------------------------------
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    set_lang(detect_language())
    root = tk.Tk()
    root.geometry("800x760")
    root.minsize(700, 660)

    q = queue.Queue()
    state = {"cancel": False, "out": None, "progress": (0, 1)}
    texts = []          # (widget, key) for re-translating

    def reg(widget, key):
        texts.append((widget, key))
        widget.config(text=tr(key))
        return widget

    lang_var = tk.StringVar(value="Nederlands" if detect_language() == "nl" else "English")
    zip_var = tk.StringVar()
    mode_var = tk.StringVar(value="word")
    rv_var = tk.BooleanVar(value=True)
    empty_var = tk.BooleanVar(value=True)
    allfiles_var = tk.BooleanVar(value=False)
    nohdr_var = tk.BooleanVar(value=False)
    hdrall_var = tk.BooleanVar(value=True)
    custom_var = tk.StringVar(value="")
    STRUCT_KEYS = ["auto", "yes", "no"]
    PN_KEYS = ["sub", "all", "none"]

    main = ttk.Frame(root, padding=12)
    main.pack(fill="both", expand=True)
    main.columnconfigure(1, weight=1)
    main.rowconfigure(10, weight=1)

    reg(ttk.Label(main, font=("Segoe UI", 10, "bold")), "step1").grid(
        row=0, column=0, columnspan=2, sticky="w")
    lf = ttk.Frame(main)
    lf.grid(row=0, column=2, sticky="e")
    reg(ttk.Label(lf), "language").pack(side="left")
    lang_cb = ttk.Combobox(lf, textvariable=lang_var, values=["Nederlands", "English"],
                           state="readonly", width=11)
    lang_cb.pack(side="left", padx=4)

    ttk.Entry(main, textvariable=zip_var).grid(row=1, column=0, columnspan=2,
                                               sticky="ew", pady=4, padx=(0, 6))
    reg(ttk.Button(main, command=lambda: choose_zip()), "browse").grid(row=1, column=2)
    info_lbl = ttk.Label(main, text="", foreground="#444")
    info_lbl.grid(row=2, column=0, columnspan=3, sticky="w")

    reg(ttk.Label(main, font=("Segoe UI", 10, "bold")), "step2").grid(
        row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
    opt = ttk.Frame(main)
    opt.grid(row=4, column=0, columnspan=3, sticky="ew")
    opt.columnconfigure(2, weight=1)

    reg(ttk.Label(opt), "type").grid(row=0, column=0, sticky="w")
    reg(ttk.Radiobutton(opt, variable=mode_var, value="word"), "word").grid(
        row=0, column=1, sticky="w", padx=8)
    reg(ttk.Radiobutton(opt, variable=mode_var, value="pdf"), "pdf").grid(
        row=0, column=2, sticky="w")
    reg(ttk.Label(opt), "structure").grid(row=1, column=0, sticky="w", pady=3)
    struct_cb = ttk.Combobox(opt, state="readonly", width=32)
    struct_cb.grid(row=1, column=1, columnspan=2, sticky="w", padx=8)
    struct_cb.bind("<<ComboboxSelected>>", lambda e: refresh_info())

    reg(ttk.Checkbutton(opt, variable=rv_var), "rv").grid(row=2, column=0, columnspan=3, sticky="w")
    reg(ttk.Checkbutton(opt, variable=empty_var), "empty").grid(row=3, column=0, columnspan=3, sticky="w")
    reg(ttk.Checkbutton(opt, variable=allfiles_var), "allfiles").grid(row=4, column=0, columnspan=3, sticky="w")
    nohdr_chk = reg(ttk.Checkbutton(opt, variable=nohdr_var, command=lambda: update_states()), "noheader")
    nohdr_chk.grid(row=5, column=0, columnspan=3, sticky="w")
    hdrall_chk = reg(ttk.Checkbutton(opt, variable=hdrall_var), "hdrall")
    hdrall_chk.grid(row=6, column=0, columnspan=3, sticky="w")
    pn_lbl = reg(ttk.Label(opt), "pagenum")
    pn_lbl.grid(row=7, column=0, sticky="w", pady=3)
    pn_cb = ttk.Combobox(opt, state="readonly", width=26)
    pn_cb.grid(row=7, column=1, columnspan=2, sticky="w", padx=8)
    cu_lbl = reg(ttk.Label(opt), "custom")
    cu_lbl.grid(row=8, column=0, sticky="w")
    cu_entry = ttk.Entry(opt, textvariable=custom_var, width=40)
    cu_entry.grid(row=8, column=1, columnspan=2, sticky="w", padx=8)
    cu_hint = reg(ttk.Label(opt, foreground="#666"), "custom_hint")
    cu_hint.grid(row=9, column=1, columnspan=2, sticky="w", padx=8)

    reg(ttk.Label(main, font=("Segoe UI", 10, "bold")), "step3").grid(
        row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))
    btns = ttk.Frame(main)
    btns.grid(row=6, column=0, columnspan=3, sticky="w", pady=4)
    go_btn = reg(ttk.Button(btns, command=lambda: start()), "make")
    go_btn.pack(side="left")
    cancel_btn = reg(ttk.Button(btns, state="disabled",
                                command=lambda: state.update(cancel=True)), "cancel")
    cancel_btn.pack(side="left", padx=6)
    open_btn = reg(ttk.Button(btns, state="disabled",
                              command=lambda: state["out"] and os.startfile(state["out"])), "open")
    open_btn.pack(side="left")

    bar = ttk.Progressbar(main, mode="determinate")
    bar.grid(row=7, column=0, columnspan=3, sticky="ew", pady=4)
    status = ttk.Label(main, text="", wraplength=740)
    status.grid(row=8, column=0, columnspan=3, sticky="w")

    logbox = tk.Text(main, height=8, wrap="word", state="disabled", font=("Consolas", 9))
    logbox.grid(row=10, column=0, columnspan=2, sticky="nsew", pady=(4, 0))
    sb = ttk.Scrollbar(main, command=logbox.yview)
    sb.grid(row=10, column=2, sticky="ns", pady=(4, 0))
    logbox.configure(yscrollcommand=sb.set)
    logbox.tag_configure("warn", foreground="#b36b00")
    logbox.tag_configure("error", foreground="#c00000")
    reg(ttk.Label(main, foreground="#666"), "privacy").grid(
        row=11, column=0, columnspan=3, sticky="w", pady=(6, 0))

    def update_states():
        st = "disabled" if nohdr_var.get() else "normal"
        ro = "disabled" if nohdr_var.get() else "readonly"
        hdrall_chk.config(state=st)
        pn_cb.config(state=ro)
        cu_entry.config(state=st)

    def retranslate():
        set_lang("nl" if lang_var.get() == "Nederlands" else "en")
        root.title(tr("app_title"))
        for w, k in texts:
            w.config(text=tr(k))
        s_idx = max(struct_cb.current(), 0)
        p_idx = max(pn_cb.current(), 0)
        struct_cb.config(values=[tr("st_auto"), tr("st_yes"), tr("st_no")])
        pn_cb.config(values=[tr("pn_sub"), tr("pn_all"), tr("pn_none")])
        struct_cb.current(s_idx)
        pn_cb.current(p_idx)
        refresh_info()

    lang_cb.bind("<<ComboboxSelected>>", lambda e: retranslate())

    def write_log(level, text):
        logbox.configure(state="normal")
        logbox.insert("end", text + "\n", level if level in ("warn", "error") else ())
        logbox.see("end")
        logbox.configure(state="disabled")

    def refresh_info():
        path = zip_var.get().strip()
        if not path or not os.path.isfile(path):
            info_lbl.config(text="")
            return
        try:
            with zipfile.ZipFile(path) as zf:
                sc = scan_zip(zf, STRUCT_KEYS[max(struct_cb.current(), 0)])
        except Exception as e:
            info_lbl.config(text=tr("cant_read", e))
            return
        if not sc.students:
            info_lbl.config(text=tr("none_found"))
            return
        empty = sum(1 for s in sc.students if not s.files)
        multi = sum(1 for s in sc.students if len(s.files) > 1)
        txt = (tr("found_cls", sc.n_classes, len(sc.students)) if sc.has_classes
               else tr("found_nocls", len(sc.students)))
        info_lbl.config(text=txt + tr("found_tail", empty, multi))

    def choose_zip():
        p = filedialog.askopenfilename(
            title=tr("pick_zip"), filetypes=[(tr("zip_files"), "*.zip"), (tr("all_files"), "*.*")])
        if p:
            zip_var.set(p)
            refresh_info()

    def start():
        zp = zip_var.get().strip()
        if not zp or not os.path.isfile(zp):
            messagebox.showwarning(tr("app_title"), tr("need_zip"))
            return
        base = os.path.splitext(os.path.basename(zp))[0]
        out = filedialog.asksaveasfilename(
            title=tr("save_as"), defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialdir=os.path.dirname(zp), initialfile=base + "_samengevoegd.pdf")
        if not out:
            return
        opts = {"mode": mode_var.get(), "classes": STRUCT_KEYS[max(struct_cb.current(), 0)],
                "recto_verso": rv_var.get(), "empty_pages": empty_var.get(),
                "all_files": allfiles_var.get(), "no_header": nohdr_var.get(),
                "header_all_pages": hdrall_var.get(),
                "page_numbers": PN_KEYS[max(pn_cb.current(), 0)],
                "custom_text": custom_var.get(),
                "lang": "nl" if lang_var.get() == "Nederlands" else "en"}
        state.update(cancel=False, out=None, progress=(0, 1))
        go_btn.config(state="disabled")
        cancel_btn.config(state="normal")
        open_btn.config(state="disabled")
        lang_cb.config(state="disabled")
        logbox.configure(state="normal")
        logbox.delete("1.0", "end")
        logbox.configure(state="disabled")
        bar["value"] = 0
        status.config(text="")

        def worker():
            com = False
            try:
                if opts["mode"] == "word" and sys.platform == "win32":
                    import pythoncom
                    pythoncom.CoInitialize()
                    com = True
                stats = build_pdf(
                    zp, out, opts,
                    log=lambda lv, t: q.put(("log", lv, t)),
                    progress=lambda c, t: q.put(("progress", c, t)),
                    cancelled=lambda: state["cancel"],
                    status=lambda text, busy=False: q.put(("status", text, busy)))
                q.put(("done", True, (out, stats)))
            except Cancelled:
                q.put(("done", False, tr("cancelled")))
            except PermissionError:
                q.put(("done", False, tr("perm")))
            except Exception as e:
                q.put(("done", False, tr("error", e)))
            finally:
                if com:
                    import pythoncom
                    pythoncom.CoUninitialize()

        threading.Thread(target=worker, daemon=True).start()

    def show_progress():
        c, t = state["progress"]
        bar.stop()
        bar.config(mode="determinate", maximum=max(t, 1), value=c)
        status.config(text=tr("working", c, t))

    def poll():
        try:
            while True:
                msg = q.get_nowait()
                if msg[0] == "log":
                    write_log(msg[1], msg[2])
                elif msg[0] == "progress":
                    state["progress"] = (msg[1], msg[2])
                    show_progress()
                elif msg[0] == "status":
                    if msg[2]:                      # busy (e.g. starting Word)
                        bar.config(mode="indeterminate")
                        bar.start(15)
                        status.config(text=msg[1] or "")
                    else:
                        show_progress()
                elif msg[0] == "done":
                    bar.stop()
                    bar.config(mode="determinate")
                    go_btn.config(state="normal")
                    cancel_btn.config(state="disabled")
                    lang_cb.config(state="readonly")
                    if msg[1]:
                        out, st = msg[2]
                        state["out"] = out
                        open_btn.config(state="normal")
                        bar["value"] = bar["maximum"]
                        summary = tr("summary", st["students"], st["pages"], st["empty"],
                                     st["multiple"], st["failed"])
                        status.config(text=summary)
                        write_log("info", "\n" + summary + "\n" + tr("saved_as", out))
                        messagebox.showinfo(tr("app_title"), summary)
                    else:
                        status.config(text=msg[2])
                        write_log("error", msg[2])
                        messagebox.showerror(tr("app_title"), msg[2])
        except queue.Empty:
            pass
        root.after(100, poll)

    retranslate()
    update_states()
    poll()
    root.mainloop()


if __name__ == "__main__":
    run_gui()
