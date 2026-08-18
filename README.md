# Book to LaTeX & PDF

[![Tests](https://github.com/kkmmee94/book-to-latex/actions/workflows/tests.yml/badge.svg)](https://github.com/kkmmee94/book-to-latex/actions/workflows/tests.yml)
[![Release builds](https://github.com/kkmmee94/book-to-latex/actions/workflows/release.yml/badge.svg)](https://github.com/kkmmee94/book-to-latex/actions/workflows/release.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A local-first desktop app, browser app, and command-line tool that turns documents into LaTeX and a compiled PDF. It is designed so a nontechnical user can choose a file, choose how closely the output should match it, and receive the result without configuring OCR, AI endpoints, or LaTeX commands.

## What the user receives

Every normal conversion keeps the output folder simple:

```text
MyDocument.tex                        LaTeX source
MyDocument.pdf                        compiled, ready-to-read PDF
MyDocument_conversion_report.txt      plain-language result or error details
MyDocument_assets/                    only when photos/logos/source visuals are required
```

Temporary page renders are deleted automatically. Required visual assets remain because deleting them would make the LaTeX impossible to compile. A compiler log is kept only when PDF creation fails. Advanced per-page review files are opt-in.

## Three simple choices

| Choice | Use it for | Result |
|---|---|---|
| **Reconstruct and polish** | prose, reports, books and fully editable documents | exact wording in a professionally redesigned LaTeX document |
| **Enhance a scan or lecture** | scanned lectures, notes, mathematics and older documents | clearer editable reconstruction that keeps the source's recognizable character |
| **Keep original pages unchanged** | documents that must remain visually unchanged | original PDF pages are placed directly into LaTeX with shapes, colors, headers, footers and images unchanged |

Every visible option in the desktop app has a hover explanation. Technical controls live under **Advanced settings**.

## Page and visual decisions are independent

- **Finished page size:** same physical shape as the source, A4, or US Letter.
- **Use of pages:** compact continuous flow or one source page per output page. Both remain available for editable modes.
- **Original pages unchanged:** always one source page per output page because changing page boundaries cannot preserve the page unchanged. It may use the source page size, A4, or US Letter.
- **Natural photographs:** keep photographs of people, animals, places and objects, or replace them with objective written descriptions.
- **Semantic visuals:** graphs, charts, tables, equations, flowcharts and technical diagrams are reconstructed as editable LaTeX only when their values and geometry can be recovered without guessing. If exact reconstruction is unsafe, the source visual is retained so it never disappears.

Temporary analysis renders are deleted. Retained photos, logos, screenshots, and non-reconstructable source visuals are project dependencies and are listed in the conversion report.

## Language support

The first language release includes:

- **English** — pdfLaTeX compilation and English OCR;
- **Arabic** — right-to-left XeLaTeX output, Unicode Arabic, Polyglossia, Amiri, and bundled official Arabic Tesseract OCR data.

The language registry is deliberately extensible. The AI is instructed to preserve the selected source language—not translate or transliterate it. See [docs/LANGUAGES.md](docs/LANGUAGES.md).

## Supported inputs

- PDF, DOCX, ODT, RTF, HTML, EML, and EPUB;
- PPTX and ODP presentations;
- XLSX, XLSM, XLS, ODS, CSV, and TSV;
- PNG, JPG, TIFF, BMP, GIF, WebP, HEIC, and HEIF;
- TXT, Markdown, LaTeX, JSON, XML, YAML, source code, logs, and other safely readable text;
- legacy DOC, DOT, PPT, PPS, XLS, and WPS through LibreOffice.

An arbitrary binary file is not meaningful language-model input. The app accepts every safely readable format it can extract or OCR and gives a clear error for encrypted, damaged, executable, media, or unsupported proprietary data.

## Install

Release builds provide:

- a Windows installer and portable GUI executable;
- a macOS application disk image;
- a Linux `.deb` package and portable GUI binary;
- Python wheel and source distribution for the CLI.

For the Python CLI on any platform:

```bash
pipx install book-to-latex
book-to-latex --help
```

From a source checkout:

```bash
python -m venv .venv
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
book-to-latex-gui
```

Linux users may also need their distribution's `python3-tk` package.

## Optional system tools

The app detects these tools at startup and uses them automatically:

| Capability | Windows | macOS | Linux |
|---|---|---|---|
| Compile PDF | [MiKTeX](https://miktex.org/download) | MiKTeX or TeX Live | TeX Live |
| OCR | [Tesseract](https://tesseract-ocr.github.io/tessdoc/Installation.html) | `brew install tesseract` | distribution Tesseract package |
| Local AI | [Ollama](https://ollama.com/download) | Ollama | Ollama |
| Legacy Office | [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/) | LibreOffice | LibreOffice |
| Enhanced PDF inspection | [Poppler](https://poppler.freedesktop.org/) | `brew install poppler` | distribution Poppler package |

Arabic OCR data ships with the app; the Tesseract executable is still required. Arabic PDF compilation requires XeLaTeX, which is included with normal MiKTeX and TeX Live installations.

The language selector defaults to automatic detection. When Arabic or Chinese OCR data is missing, the app downloads the official `tessdata_fast` language file into the user's application-data folder and reuses it. Users do not need to install Tesseract language packs manually.

The desktop and browser interfaces include **Check for updates**. It checks the public GitHub release feed and opens the official installer/download page only when a newer version exists.

## Local AI

At startup, the app lists installed Ollama models and automatically chooses a suitable one:

- `book-latex-qwen3-local-uncensored:8b` for clean editable conversion;
- `book-latex-qwen35-vision:9b` for visual mathematics and enhanced scans/lectures;
- original-pages mode does not use AI.

`UNCENSORED` is a visible user label. The provided conversion workflow identifies the source weights as the official `Qwen/Qwen3-8B` model; the label is not an independent claim about the weights.

Set up the standard models:

```powershell
.\setup_local_model.ps1
```

Import an existing Hugging Face Qwen3 folder exactly:

```powershell
.\setup_exact_local_model.ps1 -Weights "C:\path\to\Qwen3-8B-hf"
```

The second script converts the safetensors to GGUF, quantizes to Q4_K_M, creates the visible Ollama model, tests it, and removes the temporary F16 file.

## CLI examples

Clean local conversion and compilation:

```bash
book-to-latex --input manuscript.docx --output manuscript.tex --no-llm --compile-pdf --no-review
```

Arabic scanned PDF:

```bash
book-to-latex --input arabic-scan.pdf --output arabic-scan.tex --ocr --ocr-lang ara --document-language ara --no-llm --compile-pdf --no-review
```

Private Ollama conversion:

```bash
book-to-latex --input lecture.pdf --output lecture.tex --provider ollama --model book-latex-qwen3-local-uncensored:8b --compile-pdf --no-review
```

Compact editable A4 output with photograph descriptions:

```bash
book-to-latex --input lecture.pdf --output lecture.tex --provider ollama --model book-latex-qwen35-vision:9b --vision --redraw-graphs --preserve-layout --preserve-color --page-size a4 --page-flow compact --photo-handling describe --compile-pdf --no-review
```

Exact page-for-page copy on the original page size:

```bash
book-to-latex --input original.pdf --output exact-copy.tex --image-only --page-size source --page-flow source_pages --compile-pdf --no-review
```

## Reliability behavior

- AI is forbidden from inventing image filenames.
- A visual-inventory pass distinguishes semantic graphs/tables/diagrams from natural photographs, logos, screenshots and artwork.
- Invented or random semantic redraws are rejected; the original source visual is retained when exact data cannot be recovered.
- Before compilation, unavailable `\includegraphics` references are replaced with a visible notice and listed in the report.
- Common generated table, package, and TikZ mistakes are repaired before compiling.
- The compiler runs twice for references, then removes `.aux`, `.toc`, `.out`, `.log`, and other temporary build files.
- A `_compile.log` survives only after a failed compilation.
- Numeric fidelity and text similarity are checked; detailed per-page review is optional.
- Exact PDF mode references the original PDF pages directly, so it creates no page-image folder.

## Development

```bash
python -m pip install -e ".[dev]"
python -m ruff check .
python -m unittest discover -s tests -v
python -m build
```

The CI test matrix covers Windows, macOS, Linux, Python 3.10, and Python 3.12. Release automation builds the GUI and CLI artifacts for all three operating systems.

## Project map

```text
book_to_latex.py              conversion engine and CLI
book_to_latex_gui.py          primary Tk desktop interface
book_to_latex_streamlit.py    optional browser interface
assets/tessdata/              bundled OCR language data
docs/                         user, language, architecture, and release docs
models/                       Ollama Modelfile templates
packaging/                    Windows, macOS, and Linux installer metadata
tests/                        regression suite
.github/workflows/            cross-platform CI and release builds
```

Read [USER_GUIDE.md](USER_GUIDE.md) for the nontechnical walkthrough and [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for contributor details.

## License

MIT. The bundled `ara.traineddata` comes from the official Tesseract `tessdata_fast` repository and retains its upstream licensing terms; see [assets/tessdata/README.md](assets/tessdata/README.md).
