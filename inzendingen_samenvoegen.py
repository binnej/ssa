# -*- coding: utf-8 -*-
"""
Inzendingen samenvoegen tot PDF
===============================
Leest een zipbestand met studenteninzendingen en maakt EEN pdf met:
  - de nieuwste inzending van elke student (Word: via Microsoft Word, of PDF),
  - bovenaan een klein koptekstje (naam, klas, aantal inzendingen),
  - optioneel: elke inzending start op een oneven pagina (recto-verso),
  - optioneel: een lege pagina met koptekst voor studenten met een lege map,
  - bladwijzers per klas en per student.

Alles gebeurt lokaal op de computer. Er wordt niets verstuurd.
De tijdelijke bestanden worden na afloop verwijderd.
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
# Teksten (hier aanpassen indien nodig)
# --------------------------------------------------------------------------
APP_TITLE = "Inzendingen samenvoegen tot PDF"
TXT_NAME = "Naam"
TXT_CLASS = "Klas"
TXT_COUNT = "Aantal inzendingen"
TXT_MULTIPLE = "Meerdere documenten ingediend"      # vet, zoals gevraagd
TXT_FAILED = "Document kon niet verwerkt worden"
TXT_NO_DOC = "Geen bruikbaar document gevonden"

# --------------------------------------------------------------------------
# Instellingen
# --------------------------------------------------------------------------
MM = 72.0 / 25.4
A4 = (595.2756, 841.8898)
WORD_EXT = {".doc", ".docx", ".docm", ".odt", ".rtf"}
PDF_EXT = {".pdf"}
JUNK_FILES = {".ds_store", "thumbs.db", "desktop.ini"}

HEADER_FONT_SIZE = 9.0      # pt (wordt kleiner bij zeer lange namen)
HEADER_TOP = 5 * MM         # afstand bovenrand pagina tot bovenkant kader
HEADER_LEFT = 8 * MM


class Cancelled(Exception):
    pass


# --------------------------------------------------------------------------
# Zip inlezen en structuur herkennen
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
    """Herstel namen met accenten uit zips zonder UTF-8 vlag."""
    name = info.filename
    if not (info.flag_bits & 0x800):
        try:
            name = name.encode("cp437").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return name.replace("\\", "/")


def clean_student_name(name):
    # Moodle-exports: "Jan Peeters_123456_assignsubmission_file_"
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

    # Welke mappen zijn "studentmappen"? Mappen die rechtstreeks bestanden
    # bevatten, of lege eindmappen. De meest voorkomende diepte wint.
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

    by_key = {}   # volgorde = volgorde van eerste voorkomen in de zip
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


def pick_latest(student, exts):
    best, best_key = None, None
    for idx, f in enumerate(student.files):
        if f.ext in exts:
            k = (f.info.date_time, idx)
            if best_key is None or k > best_key:
                best, best_key = f, k
    return best


# --------------------------------------------------------------------------
# Koptekst (overlay bovenaan de pagina)
# --------------------------------------------------------------------------
_fonts_registered = None


def _register_fonts():
    """Geeft (regular, bold) van een Unicode-font terug, of None."""
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


def pick_fonts(segments):
    """Helvetica (klein bestand) tenzij er tekens zijn die het niet kent."""
    text = "".join(t for t, _ in segments)
    try:
        text.encode("cp1252")
        return ("Helvetica", "Helvetica-Bold")
    except UnicodeEncodeError:
        return _register_fonts() or ("Helvetica", "Helvetica-Bold")


def header_segments(student, n_files, extra=None):
    segs = [("%s: %s" % (TXT_NAME, student.name), False)]
    if student.cls:
        segs.append(("   |   %s: %s" % (TXT_CLASS, student.cls), False))
    segs.append(("   |   %s: " % TXT_COUNT, False))
    segs.append((str(n_files), n_files > 1))
    if n_files > 1:
        segs.append(("   " + TXT_MULTIPLE, True))
    if extra:
        segs.append(("   |   " + extra, True))
    return segs


def make_overlay(w, h, segments, fonts):
    reg, bold = fonts
    safe = fonts[0] == "Helvetica"
    if safe:  # Helvetica kent enkel cp1252
        segments = [(t.encode("cp1252", "replace").decode("cp1252"), b) for t, b in segments]

    def total(sz):
        return sum(pdfmetrics.stringWidth(t, bold if b else reg, sz) for t, b in segments)

    size = HEADER_FONT_SIZE
    maxw = w - 2 * HEADER_LEFT
    while total(size) > maxw and size > 5:
        size -= 0.25
    tw = total(size)
    pad = 1.5 * MM
    band_h = size + 2 * pad
    top = h - HEADER_TOP

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    # witte achtergrond, enkel rond de tekst, zodat die leesbaar blijft
    c.setFillColorRGB(1, 1, 1)
    c.setFillAlpha(1)
    c.rect(HEADER_LEFT - pad, top - band_h, tw + 2 * pad, band_h, stroke=0, fill=1)
    c.setFillAlpha(1)
    c.setFillColorRGB(0, 0, 0)
    x, y = HEADER_LEFT, top - pad - size * 0.8
    for text, b in segments:
        font = bold if b else reg
        c.setFont(font, size)
        c.drawString(x, y, text)
        x += pdfmetrics.stringWidth(text, font, size)
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def stamp_page(page, segments, fonts, cache):
    try:
        if getattr(page, "rotation", 0) % 360:
            page.transfer_rotation_to_content()
    except Exception:
        pass
    box = page.cropbox
    left, bottom = float(box.left), float(box.bottom)
    w, h = float(box.width), float(box.height)
    key = (round(w, 1), round(h, 1))
    ov = cache.get(key)
    if ov is None:
        ov = make_overlay(w, h, segments, fonts)
        cache[key] = ov
    page.merge_transformed_page(ov, Transformation().translate(left, bottom))


# --------------------------------------------------------------------------
# Word -> PDF via Microsoft Word (lokaal geinstalleerd)
# --------------------------------------------------------------------------
class WordConverter:
    def __init__(self):
        if sys.platform != "win32":
            raise RuntimeError("Word-modus werkt enkel op Windows met Microsoft Word.")
        try:
            import win32com.client
            self.word = win32com.client.DispatchEx("Word.Application")  # eigen instantie
        except Exception:
            raise RuntimeError("Microsoft Word kon niet gestart worden. "
                               "Is Word geinstalleerd op deze computer?")
        self.word.Visible = False
        self.word.DisplayAlerts = 0
        try:
            self.word.AutomationSecurity = 3     # macro's nooit uitvoeren
        except Exception:
            pass

    def convert(self, src, dst):
        doc = None
        try:
            # Positioneel: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
            # PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument.
            # Een dummy-wachtwoord zorgt dat beveiligde documenten meteen mislukken
            # in plaats van een (onzichtbare) wachtwoordvraag te tonen.
            doc = self.word.Documents.Open(src, False, True, False, "x", "x", False, "x")
            try:
                doc.ExportAsFixedFormat(dst, 17)     # 17 = wdExportFormatPDF
            except Exception:
                doc.SaveAs2(dst, 17)
        finally:
            if doc is not None:
                try:
                    doc.Close(0)                      # 0 = niet opslaan
                except Exception:
                    pass
        return dst

    def close(self):
        try:
            self.word.Quit(0)
        except Exception:
            pass


# --------------------------------------------------------------------------
# Hoofdlogica
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


def build_pdf(zip_path, out_path, opts, log, progress, cancelled, converter_factory=None):
    """
    opts: mode ('word'|'pdf'), classes ('auto'|'yes'|'no'), recto_verso (bool),
          empty_pages (bool), header_all_pages (bool)
    log(level, text): level in 'info' | 'warn' | 'error'
    """
    word_mode = opts["mode"] == "word"
    exts = WORD_EXT if word_mode else PDF_EXT
    stats = {"students": 0, "pages": 0, "multiple": 0, "empty": 0, "failed": 0}

    with zipfile.ZipFile(zip_path) as zf:
        scan = scan_zip(zf, opts.get("classes", "auto"))
        students = scan.students
        if not students:
            raise RuntimeError("Geen studentenmappen gevonden in het zipbestand.")
        for lf in scan.loose_files:
            log("warn", "Los bestand buiten een studentenmap genegeerd: %s" % lf)

        tmpdir = tempfile.mkdtemp(prefix="inzendingen_")
        converter = None
        try:
            if word_mode:
                converter = (converter_factory or WordConverter)()
            writer = PdfWriter()
            class_items = {}
            total = len(students)

            def pad_to_odd():
                if opts["recto_verso"] and len(writer.pages) % 2 == 1:
                    writer.add_blank_page(A4[0], A4[1])

            def add_bookmarks(st, first_page):
                parent = None
                if st.cls:
                    if st.cls not in class_items:
                        class_items[st.cls] = writer.add_outline_item(st.cls, first_page)
                    parent = class_items[st.cls]
                title = st.name
                if len(st.files) > 1:
                    title += " (%d bestanden)" % len(st.files)
                writer.add_outline_item(title, first_page, parent)

            for i, st in enumerate(students, 1):
                if cancelled():
                    raise Cancelled()
                progress(i - 1, total)
                n = len(st.files)
                label = st.name + (" (%s)" % st.cls if st.cls else "")
                chosen = pick_latest(st, exts)
                src_pages, failure = None, None

                if n > 1:
                    stats["multiple"] += 1
                    log("warn", "%s: %d bestanden ingediend." % (label, n))
                if chosen is not None:
                    src = pdf = None
                    try:
                        src = _extract(zf, chosen, tmpdir, i)
                        pdf = converter.convert(src, os.path.join(tmpdir, "%d.pdf" % i)) \
                            if converter else src
                        reader = PdfReader(pdf)
                        if reader.is_encrypted and not reader.decrypt(""):
                            raise ValueError("beveiligd met een wachtwoord")
                        src_pages = list(reader.pages)
                        if not src_pages:
                            raise ValueError("document bevat geen pagina's")
                        d = chosen.info.date_time
                        log("info", "[%d/%d] %s: %s (%04d-%02d-%02d %02d:%02d)" % (
                            i, total, label, os.path.basename(chosen.path), *d[:5]))
                    except Exception as e:
                        src_pages = None
                        failure = TXT_FAILED
                        stats["failed"] += 1
                        log("error", "%s: %s kon niet verwerkt worden (%s)" % (
                            label, os.path.basename(chosen.path), e))
                    finally:
                        _silent_remove(*(p for p in (src, pdf) if p))
                elif n > 0:
                    failure = TXT_NO_DOC
                    stats["failed"] += 1
                    log("error", "%s: geen bruikbaar document (%d bestand(en), "
                                 "geen %s)." % (label, n, "/".join(sorted(exts))))
                else:
                    stats["empty"] += 1
                    log("warn", "%s: lege map." % label)

                segs = header_segments(st, n, failure)
                fonts = pick_fonts(segs)
                cache = {}

                if src_pages:
                    pad_to_odd()
                    first = len(writer.pages)
                    for k, p in enumerate(src_pages):
                        newp = writer.add_page(p)
                        if k == 0 or opts["header_all_pages"]:
                            stamp_page(newp, segs, fonts, cache)
                    add_bookmarks(st, first)
                elif failure or opts["empty_pages"]:
                    pad_to_odd()
                    first = len(writer.pages)
                    newp = writer.add_blank_page(A4[0], A4[1])
                    stamp_page(newp, segs, fonts, cache)
                    add_bookmarks(st, first)
                else:
                    continue
                stats["students"] += 1

            progress(total, total)
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
# Venster (Tkinter, zit standaard in Python en in het .exe-bestand)
# --------------------------------------------------------------------------
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("760x600")
    root.minsize(660, 520)

    q = queue.Queue()
    state = {"cancel": False, "running": False, "out": None}

    zip_var = tk.StringVar()
    mode_var = tk.StringVar(value="word")
    rv_var = tk.BooleanVar(value=True)
    empty_var = tk.BooleanVar(value=True)
    all_var = tk.BooleanVar(value=True)
    struct_var = tk.StringVar(value="Automatisch herkennen")
    STRUCT = {"Automatisch herkennen": "auto", "Met klassen (klas/student)": "yes",
              "Zonder klassen (student)": "no"}

    main = ttk.Frame(root, padding=12)
    main.pack(fill="both", expand=True)
    main.columnconfigure(1, weight=1)
    main.rowconfigure(9, weight=1)

    ttk.Label(main, text="1. Kies het zipbestand met de inzendingen",
              font=("Segoe UI", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
    ttk.Entry(main, textvariable=zip_var).grid(row=1, column=0, columnspan=2,
                                               sticky="ew", pady=4, padx=(0, 6))
    ttk.Button(main, text="Bladeren...", command=lambda: choose_zip()).grid(row=1, column=2)
    info_lbl = ttk.Label(main, text="", foreground="#444")
    info_lbl.grid(row=2, column=0, columnspan=3, sticky="w")

    ttk.Label(main, text="2. Opties", font=("Segoe UI", 10, "bold")).grid(
        row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))
    opt = ttk.Frame(main)
    opt.grid(row=4, column=0, columnspan=3, sticky="w")
    ttk.Label(opt, text="Type inzendingen:").grid(row=0, column=0, sticky="w")
    ttk.Radiobutton(opt, text="Word (doc, docx, odt)", variable=mode_var,
                    value="word").grid(row=0, column=1, sticky="w", padx=8)
    ttk.Radiobutton(opt, text="PDF", variable=mode_var,
                    value="pdf").grid(row=0, column=2, sticky="w")
    ttk.Label(opt, text="Mapstructuur:").grid(row=1, column=0, sticky="w", pady=3)
    cb = ttk.Combobox(opt, textvariable=struct_var, values=list(STRUCT), state="readonly", width=30)
    cb.grid(row=1, column=1, columnspan=2, sticky="w", padx=8)
    cb.bind("<<ComboboxSelected>>", lambda e: refresh_info())
    ttk.Checkbutton(opt, text="Recto-verso: elke inzending begint op een oneven pagina",
                    variable=rv_var).grid(row=2, column=0, columnspan=3, sticky="w")
    ttk.Checkbutton(opt, text="Lege pagina met koptekst voor studenten met een lege map",
                    variable=empty_var).grid(row=3, column=0, columnspan=3, sticky="w")
    ttk.Checkbutton(opt, text="Koptekst op elke pagina (anders enkel op de eerste pagina)",
                    variable=all_var).grid(row=4, column=0, columnspan=3, sticky="w")

    ttk.Label(main, text="3. PDF maken", font=("Segoe UI", 10, "bold")).grid(
        row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))
    btns = ttk.Frame(main)
    btns.grid(row=6, column=0, columnspan=3, sticky="w", pady=4)
    go_btn = ttk.Button(btns, text="PDF maken", command=lambda: start())
    go_btn.pack(side="left")
    cancel_btn = ttk.Button(btns, text="Annuleren", state="disabled",
                            command=lambda: state.update(cancel=True))
    cancel_btn.pack(side="left", padx=6)
    open_btn = ttk.Button(btns, text="PDF openen", state="disabled",
                          command=lambda: state["out"] and os.startfile(state["out"]))
    open_btn.pack(side="left")

    bar = ttk.Progressbar(main, mode="determinate")
    bar.grid(row=7, column=0, columnspan=3, sticky="ew", pady=4)
    status = ttk.Label(main, text="")
    status.grid(row=8, column=0, columnspan=3, sticky="w")

    logbox = tk.Text(main, height=10, wrap="word", state="disabled", font=("Consolas", 9))
    logbox.grid(row=9, column=0, columnspan=2, sticky="nsew", pady=(4, 0))
    sb = ttk.Scrollbar(main, command=logbox.yview)
    sb.grid(row=9, column=2, sticky="ns", pady=(4, 0))
    logbox.configure(yscrollcommand=sb.set)
    logbox.tag_configure("warn", foreground="#b36b00")
    logbox.tag_configure("error", foreground="#c00000")
    ttk.Label(main, text="Alle verwerking gebeurt lokaal op deze computer; er wordt niets verstuurd.",
              foreground="#666").grid(row=10, column=0, columnspan=3, sticky="w", pady=(6, 0))

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
                sc = scan_zip(zf, STRUCT[struct_var.get()])
        except Exception as e:
            info_lbl.config(text="Kan het zipbestand niet lezen: %s" % e)
            return
        if not sc.students:
            info_lbl.config(text="Geen studentenmappen gevonden.")
            return
        empty = sum(1 for s in sc.students if not s.files)
        multi = sum(1 for s in sc.students if len(s.files) > 1)
        txt = ("Gevonden: %d klassen, %d studenten" % (sc.n_classes, len(sc.students))
               if sc.has_classes else "Gevonden: %d studenten (zonder klassen)" % len(sc.students))
        txt += " | lege mappen: %d | meerdere bestanden: %d" % (empty, multi)
        info_lbl.config(text=txt)

    def choose_zip():
        p = filedialog.askopenfilename(title="Kies het zipbestand",
                                       filetypes=[("Zipbestand", "*.zip"), ("Alle bestanden", "*.*")])
        if p:
            zip_var.set(p)
            refresh_info()

    def start():
        zp = zip_var.get().strip()
        if not zp or not os.path.isfile(zp):
            messagebox.showwarning(APP_TITLE, "Kies eerst een zipbestand.")
            return
        base = os.path.splitext(os.path.basename(zp))[0]
        out = filedialog.asksaveasfilename(
            title="Opslaan als", defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialdir=os.path.dirname(zp), initialfile=base + "_samengevoegd.pdf")
        if not out:
            return
        opts = {"mode": mode_var.get(), "classes": STRUCT[struct_var.get()],
                "recto_verso": rv_var.get(), "empty_pages": empty_var.get(),
                "header_all_pages": all_var.get()}
        state.update(cancel=False, running=True, out=None)
        go_btn.config(state="disabled")
        cancel_btn.config(state="normal")
        open_btn.config(state="disabled")
        logbox.configure(state="normal")
        logbox.delete("1.0", "end")
        logbox.configure(state="disabled")
        bar["value"] = 0
        status.config(text="Bezig...")

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
                    cancelled=lambda: state["cancel"])
                q.put(("done", True, (out, stats)))
            except Cancelled:
                q.put(("done", False, "Geannuleerd."))
            except PermissionError:
                q.put(("done", False, "Kan het PDF-bestand niet opslaan. "
                                      "Is het nog geopend in een ander programma?"))
            except Exception as e:
                q.put(("done", False, "Fout: %s" % e))
            finally:
                if com:
                    import pythoncom
                    pythoncom.CoUninitialize()

        threading.Thread(target=worker, daemon=True).start()

    def poll():
        try:
            while True:
                msg = q.get_nowait()
                if msg[0] == "log":
                    write_log(msg[1], msg[2])
                elif msg[0] == "progress":
                    bar["maximum"] = max(msg[2], 1)
                    bar["value"] = msg[1]
                    status.config(text="Bezig: %d van %d" % (msg[1], msg[2]))
                elif msg[0] == "done":
                    state["running"] = False
                    go_btn.config(state="normal")
                    cancel_btn.config(state="disabled")
                    if msg[1]:
                        out, st = msg[2]
                        state["out"] = out
                        open_btn.config(state="normal")
                        bar["value"] = bar["maximum"]
                        summary = ("Klaar: %d studenten, %d pagina's. Lege mappen: %d, "
                                   "meerdere bestanden: %d, mislukt: %d." % (
                                       st["students"], st["pages"], st["empty"],
                                       st["multiple"], st["failed"]))
                        status.config(text=summary)
                        write_log("info", "\n" + summary + "\nOpgeslagen als: " + out)
                        messagebox.showinfo(APP_TITLE, summary)
                    else:
                        status.config(text=msg[2])
                        write_log("error", msg[2])
                        messagebox.showerror(APP_TITLE, msg[2])
        except queue.Empty:
            pass
        root.after(100, poll)

    poll()
    root.mainloop()


if __name__ == "__main__":
    run_gui()
