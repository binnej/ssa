# Inzendingen samenvoegen tot PDF

> ## ⚠️ This is a vibe-coded project
> This program was written entirely by an AI assistant (Claude) in a single chat session, directed by a non-programmer's requirements and **not reviewed line by line by a human**. The complete conversation is included in the [appendix](#appendix-the-complete-chat-session).
>
> **What was tested:** zip scanning, class/no-class detection, latest-file selection, header stamping, recto-verso padding, blank header pages and bookmarks, using synthetic zips in PDF mode on Linux.
> **What was *not* tested:** the Microsoft Word conversion (needs Windows + Word), the Tkinter window, and the PyInstaller `.exe` build.
> **Please try it on a small real zip before relying on it.** Student work is involved, so check the output page by page the first time. Use at your own risk.

Merge a zip of student submissions into **one PDF**, with a small header on each submission (student name, class, number of files submitted). Everything runs locally; nothing is uploaded anywhere (GDPR-friendly).

The UI and header texts are in **Dutch** (Flemish school context). Strings are at the top of the Python file if you want to translate them.

## Features

- Reads a zip in one of two layouts (auto-detected, can be overridden in the window):
  1. `Class/Student/files…` (class shown in the header and bookmarks)
  2. `Student/files…` (no class)
- Takes the **most recent file** per student, by the timestamp stored in the zip.
- **Word mode:** `.doc`, `.docx`, `.docm`, `.odt`, `.rtf` are converted to PDF by your locally installed **Microsoft Word**, so the layout is what Word shows.
- **PDF mode:** student PDFs are used as they are (no Word needed).
- **Header** (9 pt, one line at the top edge of the page, drawn over the document with a small white box; the page is not shrunk):
  - `Naam: …  |  Klas: …  |  Aantal inzendingen: N`
  - If N > 1, the count and **Meerdere documenten ingediend** are printed in bold.
  - Option to put it on every page (default) or only the first page of each submission.
- **Recto-verso option:** every submission starts on an odd page; a white page is inserted when needed.
- **Empty folder option:** a page with only the header (count 0) for students with an empty folder. Students with unusable files (corrupt, password-protected, wrong type) always get a header page with a bold note.
- **PDF bookmarks** per class and per student, in zip order.
- Log with warnings (empty folders, multiple files, failures). Moodle name suffixes (`_123456_assignsubmission_file_`) are removed from student names.
- Temporary files are deleted after each student.

## Requirements

- Windows.
- Microsoft Word (only for Word mode; PDF mode does not need it).
- To **run from source**: Python 3.9+ with `pypdf`, `reportlab`, `pywin32`.
- To **build the `.exe`**: additionally `pyinstaller` (done by `build_exe.bat`).

## Usage

### For teachers (with the `.exe`)
1. Double-click `InzendingenSamenvoegen.exe`.
2. Choose the zip, choose Word or PDF mode, tick the options, click **PDF maken**.
3. Choose where to save the PDF.

Windows SmartScreen may show "unknown publisher" on first run (the exe is unsigned): choose *More info → Run anyway*, or have IT approve it.

### Building the `.exe` (once, by one person)
```bat
build_exe.bat
```
This installs the dependencies and produces `dist\InzendingenSamenvoegen.exe` with PyInstaller (`--onefile --windowed`).

### Running from source
```bat
pip install pypdf reportlab pywin32
python inzendingen_samenvoegen.py
```

## Files

| File | Purpose |
|---|---|
| `inzendingen_samenvoegen.py` | The whole program (logic + Tkinter window) |
| `build_exe.bat` | Builds the standalone `.exe` with PyInstaller |
| `LEESMIJ.txt` | Short Dutch instructions for teachers |

## How it works

1. `scan_zip` finds the student folders: the most common depth of folders that directly contain files (or are empty leaf folders). Depth 1 = no classes, depth ≥ 2 = classes.
2. For each student, the newest file with a suitable extension is extracted to a temp folder.
3. Word mode: Word is started invisibly (own instance, macros disabled, read-only open, dummy password so protected files fail instead of prompting) and exports the file to PDF.
4. `pypdf` merges the pages; `reportlab` draws the header overlay, which is merged onto the page top. Helvetica is used for Western characters; if a name has other characters (e.g. Cyrillic), Arial (or DejaVu/Liberation as fallback) is embedded.
5. Bookmarks and the final PDF are written; temp files are removed.

## Known limitations

- Word conversion is untested (see the notice at the top).
- Documents are not rescaled: a Letter-size document stays Letter-size (blank/padding pages are A4).
- The header can cover document text at the very top of the page (by design).
- Only the newest file per student is used; there is no manual override.
- Empty folders only show up if the zip contains explicit folder entries; some zip tools drop empty folders.
- A Word document that hangs or shows an unexpected dialog can stall the run (there is no timeout).
- Memory use grows with the size of all submissions (pypdf keeps the output in memory); fine for ~50–200 students with normal documents.
- Arabic, CJK and other scripts not covered by Arial are not guaranteed to render in the header.

## License

No license has been chosen yet. Add one before publishing if you want others to reuse the code.

---

# Appendix: the complete chat session

Below is the full conversation in which this project was specified and built, in order. Tool calls (memory notes, file creation, test runs) are not reproduced, only the messages exchanged. Only the user's messages and Claude's replies are included.

## Turn 1: User

> Claude, I have a zip file containing student submissions. The folderstructure in the zip file is:
> Scenario 1:
>
> * Each class has its own folder. The name of the folder is the name of the class.
> * In that folder are other folders, each folder corresponds to the name of one student.
>
> Scenario 2:
>
> * Zip file just contains folders with student's names, no class info.
>
> Each student submits one word file. If the student submits multiple files, the most recent is taken.
>
> Make a website that generates a PDF document. The PDF document contains:
>
> * Each student's submission.
> * A header to the top of the page is added with following info:
>    * Student's name
>    * Class in case this is scenario 1
>    * The amount of submissions for that student, if it is more than one, put it in bold. Add the text in bold: "Meerdere documenten ingediend".
>    * Make the header small so that its interference with the original document is minimal, but the header is still easy to read and the document can be printed.
> * If recto-verso is enabled, A new submission starts at an odd page. A white page is added if needed.
>
> Add an option to add an empty page with header in case a student's folder is empty.
> No data should leave the local machine as these documents are private (GDPR).
> Analyse this prompt first, give me suggestions for improvements on the website if you think of something. The website is preferably one html file. Is everything clear? Need more info?

## Turn 1: Claude

> The spec is mostly clear. The one real problem is converting Word to PDF.
>
> **The core technical issue**
>
> A single offline HTML file can't convert .docx to PDF faithfully. Browser-only options (docx-preview, mammoth) re-render the document as HTML, which loses layout fidelity: tables, text boxes, fonts, headers/footers, equations and page breaks can all shift. For work you may grade or print, that matters.
>
> Options:
> 1. **Pure browser (one HTML file):** libraries are inlined so it works offline and nothing leaves the machine. Lower fidelity, and page numbering and page breaks are approximate.
> 2. **Local helper plus HTML UI:** a small Python script (or the LibreOffice command line) converts each docx to PDF on your PC, then the page merges and adds headers. Much better fidelity, but it is not a single HTML file.
> 3. **Hybrid:** one HTML file with a "fast preview" mode (option 1), plus an optional script for high-fidelity conversion.
>
> I'd recommend option 2 or 3 if layout accuracy matters.
>
> **Questions**
>
> 1. Which fidelity do you need: "readable" or "looks exactly like in Word"?
> 2. How do I determine "most recent"? Zip timestamps are often wrong, since LMS exports frequently reset them. Alternatives are the filename (e.g. a date or version number) or the Word document's own "last modified" metadata. Which LMS do the files come from?
> 3. Is it only .docx, or also .doc, .odt and PDF?
> 4. Does "number of submissions" count all files in the folder, or only Word files?
> 5. Page size: always A4?
>
> **Suggested improvements**
>
> - **Auto-detect the scenario** (with or without class folders) and show a preview of what was found: classes, students, file counts.
> - **Header placement:** by default, overlay it in the top margin (about 8 pt text, one line). If the document's top margin is too small, optionally shrink the page content slightly instead.
> - **Sorting:** alphabetical by class, then by student, with a choice of first name or last name.
> - **PDF bookmarks** per class and student for easy navigation.
> - **Output split:** one PDF for everything, or one PDF per class.
> - **Warnings list:** empty folders, unreadable or corrupt files, non-Word files, students with multiple submissions, with the chosen file shown.
> - **Summary page** (optional) at the start: list of students, submission counts and page numbers.
> - **Manual override:** pick a different file for a student when several were submitted.
> - **Page numbering:** optionally continuous numbering or a per-student "page x of y" in the header.
> - **Privacy:** no network calls at all (all libraries inlined, no CDNs, no fonts or analytics), and an optional note in the UI that files are processed in memory only.
> - **Performance:** large classes can use a lot of memory, so process students one at a time and show a progress bar.
>
> Let me know your answers to the questions above, especially 1 and 2, and I'll build it.

## Turn 2: User

> Can you make a word plugin instead? That may be easier?
> 1) Fidelity: looks exactly as in word, if possible. Word documents are pretty standard so no weirdd corner cases expected.
> 2) Use the timestamp of the file, I checked, it is reliable.
> 3) I expect also doc and odt mostly. Can you add two modes to the program: word submissions (odt, doc and docx) and pdf? If this is too large a burden for a word plugin, skip pdf.
> 4) all files in the folder. It is a sort of warning that something is wrong with the upload.
> 5) Yes. A4.
> Suggested improvements:
> autodetect: sure!
> Header placement: do not shrink the page. Just put it on top of the document. If it overlays on document text, so be it.
> Sorting: order of folders in the zip file.
> PDF bookmarks: would be nice, also a word equivalent, but only nice to have.
> Output split: not needed
> warning list: optional, but probably too difficult for a word plugin.
> Summary page: no
> Manual override: not needed
> numbering: no
> privacy: correct, if the website downloads some online library, that is fine, just do not send submissions to any server.
> Performance: Max is 200 students. More realistic: 50 - 100.

## Turn 2: Claude

> A Word add-in is not easier, and it can't do all of this.
>
> **Why an add-in is a poor fit**
> - Add-ins can only insert `.docx`, not `.doc` or `.odt`, so two of your three formats are out. PDF mode is out too.
> - Headers and odd-page breaks would have to be built per section inside Word, which is fiddly.
> - An add-in has to be served from a local web server and sideloaded, which is more setup for you than a script.
>
> **What I'd recommend instead: a small local Windows program (Python plus your installed Word)**
> - Word itself converts each `.doc`, `.docx` and `.odt` to PDF, so the layout is exactly what Word shows.
> - Python then stamps the header on top of page 1 of each submission (overlay, no shrinking), adds blank pages for odd-page starts, adds an empty header page for empty folders, and adds PDF bookmarks per student.
> - PDF mode becomes trivial, since PDFs need no conversion.
> - It is built with a simple window: pick the zip, choose Word or PDF mode, and set the recto-verso and empty-folder options.
> - It auto-detects the scenario and sorts by zip order. The newest file is picked by timestamp, and the count is all files in the folder.
> - Everything runs locally and nothing is uploaded. 200 students is fine.
> - The warnings list (empty folders, multiple files) is easy here, so I'd include it.
>
> **Requirements:** Windows with Word installed, and Python with `pywin32` and `pypdf`/`reportlab`. I can also give you a one-click `.bat` launcher, or a packaged `.exe` if you prefer.
>
> **Alternative:** if you'd rather skip Python, a Word macro (VBA) can do the conversion and merging, but the header stamping and bookmarks are clumsier there.
>
> Shall I build the Python program? Just let me know whether Word is installed on the machine where you'll run it.

## Turn 3: User

> I want this program to be easy usable. I need something that teachers can download, double click and it should work. Not sure if python is installed on their machines...

## Turn 3: Claude

*(Claude read the PDF skill, wrote `inzendingen_samenvoegen.py`, tested it with synthetic zips in PDF mode, inspected the rendered headers, and created `build_exe.bat` and `LEESMIJ.txt`. A short status message was sent while it was checking the header rendering.)*

> Quick status: the core logic works in my tests (scan, latest-file pick, odd-page padding, bookmarks). I'm now checking the header rendering visually before packaging.

*(Final message of that turn:)*

> I've built the program, but only half of it is tested. Teachers get a double-clickable `.exe` that needs no Python. I can't build the `.exe` here, so someone has to build it once on a Windows PC.
>
> I tested the zip scanning and PDF stamping with fake zips in PDF mode. I could not test the Word conversion or the window, because there's no Word or Tkinter in my environment. Please try it on one real zip before handing it out.
>
> **What it does**
> - **Layout detection:** it works out from the zip whether there are class folders, and a dropdown lets you override that.
> - **Which file counts:** it uses the newest file by timestamp. It also strips Moodle suffixes like `_12345_assignsubmission_file_` from names.
> - **Word mode:** Word itself converts `.doc`, `.docx` and `.odt` to PDF, so the layout matches what Word shows.
> - **PDF mode:** it uses the PDFs directly and doesn't need Word.
> - **Header:** a small 9 pt line at the top edge, drawn over the page with a white box only around the text, so the page is not shrunk. It shows name, class and number of files. When there is more than one file, the count and "Meerdere documenten ingediend" are bold.
> - **Recto-verso:** each submission starts on an odd page, with a white page added when needed. Empty folders get a header-only page if you tick the option.
> - **Bookmarks:** one per class and one per student. Students come out in zip order.
> - **Log:** the window shows a progress bar and a log of empty folders and multiple files.
> - **Privacy:** nothing is sent anywhere, and temporary files are deleted after each student.
>
> **To make the `.exe`:** on one Windows PC with Python, double-click `build_exe.bat`. It produces `dist\InzendingenSamenvoegen.exe`, which teachers can then copy and double-click. They need Microsoft Word for Word mode, and nothing extra for PDF mode.
>
> **Things I decided that you may want to change**
> - The header appears on every page of a submission. There is a checkbox to limit it to the first page.
> - A student whose files aren't usable (a `.zip`, a corrupt file, a password-protected document) gets a header page with a bold note, so they don't silently disappear.
> - The window and header text are in Dutch.
> - Windows SmartScreen may show "unknown publisher" the first time someone runs the `.exe`. They choose "More info" and then "Run anyway", or your IT team can approve it.
> - I assumed the zip contains explicit empty folder entries. Some zip tools drop empty folders, in which case those students won't appear at all.
>
> Files are in `/mnt/user-data/outputs/inzendingen-samenvoegen/`:
> - inzendingen_samenvoegen.py
> - build_exe.bat
> - LEESMIJ.txt

## Turn 4: User

> Please make a readme.md for github. Include this entire chat session as an appendix. Make clear that this is a vibe-coded project at the beginning of the readme.md

## Turn 4: Claude

> This README.md, with the vibe-coded notice at the top and this chat session as the appendix.
