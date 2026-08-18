# Book to LaTeX & PDF — Complete User Guide

This guide assumes no programming or LaTeX experience.

## Before your first conversion

For ordinary documents, ebooks, spreadsheets, presentations, images, PDFs, and readable text/code files, open `book-reader.exe` and start immediately.

The app checks automatically whether a PDF page has readable text and uses OCR only when necessary.

## A safe first conversion

1. Open the application.
2. Select **Choose file…**.
3. Select the file you want to convert. The chooser includes an **All files** option.
4. The app suggests a filename ending in `_latex.tex`. Change it only if desired.
5. Choose **Clean and editable**, **Stay close to the original layout**, or **Exact visual copy**.
6. Choose original colours or black and white.
7. Select **Create LaTeX and PDF**.
8. The app automatically chooses the AI model, OCR behavior, review checks and compiler.
9. Wait until the message begins with **Finished**.
10. Select **Open finished PDF** to read the result and **Open report** to see what happened.

Move the pointer over any label, box, or checkbox and pause briefly. A yellow explanation appears.

## What files can I choose?

The app accepts any selected file and has dedicated readers for:

- Word DOCX, OpenDocument ODT, RTF, HTML, and saved EML email;
- EPUB ebooks;
- PowerPoint PPTX and OpenDocument ODP presentations;
- Excel XLSX/XLSM/XLS, ODS, CSV, and TSV data;
- PDF documents;
- PNG, JPG/JPEG, TIFF, BMP, GIF, WebP, and HEIC/HEIF images;
- ordinary text, Markdown, JSON, XML, YAML, TOML, configuration, log, SQL, LaTeX, and common source-code files.

Unknown extensions are accepted when the content can be safely decoded as text. Legacy DOC, DOT, PPT, PPS, XLS, and WPS files are read through LibreOffice when it is installed.

Before the optional LLM is called, the app extracts readable text, slide content, table rows, ebook sections, or OCR text. Raw executable, encrypted, damaged, proprietary, audio, and video bytes are not sent to the LLM as if they were text; the app gives a clear error when no safe extractor exists.

## Understanding the three output styles

### Clean and editable

Use this for almost every manuscript. The generated words are actual LaTeX text that can be edited, searched, copied, and restyled.

This mode tries to preserve content, not exact typography or page coordinates. PDF text extraction does not reliably expose every font, column, margin, or floating picture.

### Stay close to the original layout

This automatically uses the installed vision model. It reads the rendered page image together with layout-preserved extracted text and reconstructs mathematics, tables, columns and structure as editable LaTeX. It aims to stay close, but exact coordinates and fonts are not guaranteed.

### Exact visual copy

For PDFs, this places each original PDF page directly into the LaTeX output. It does not create a folder containing hundreds of temporary page pictures. Use it when appearance is more important than editing.

The result looks closest to the source, including diagrams, signatures, formulas, colors, and layout. Because each page is a picture, its words cannot be edited as LaTeX text.

## Technical controls are optional

The main window intentionally hides model names, endpoints, OCR switches, token limits, retries and matching percentages. Select **Advanced settings** only if you need a page range, poetry line-break preservation, another AI service, a custom style guide, or graph redrawing.

## Understanding conversion methods

### No AI — reliable basic conversion

The app extracts content using the reader for that format, then safely handles LaTeX special characters such as `&`, `%`, `_`, `#`, braces, and common mathematical symbols.

Advantages:

- no internet connection;
- book text never leaves the computer;
- predictable output;
- no API cost.

It does not infer headings, tables, complex formulas, or semantic formatting.

### Ollama — private AI on this computer

Ollama runs AI models locally. At startup, the app automatically lists every installed Ollama model and selects the recommended Book-to-LaTeX model. Select another model from the dropdown or use **Refresh local models** after installing one—no model name or endpoint needs to be typed.

Run `setup_local_model.ps1` to install the two configured models:

- `book-latex-qwen3-local-uncensored:8b` for clean editable conversion from the exact local Qwen3 safetensors;
- `book-latex-qwen3:8b` as a fallback text model;
- `book-latex-qwen35-vision:9b` for mathematical and visual PDF/image pages.

The result depends on the selected model and available computer memory. Always inspect the review report.

### Vision-assisted pages

Select **Let a vision AI inspect each PDF/image page** with a model whose name contains `vision`, `VL`, or `qwen3.5`. The app sends a rendered page picture alongside layout-preserved text. This lets the model interpret stacked fractions, summation limits, superscripts, matrices, cases, diagrams, and columns that plain text extraction damages.

The text-only local Qwen3 models cannot inspect images. **Stay close to the original layout** automatically uses the installed Qwen3.5 vision model.

### Project style guide

An optional TXT/Markdown style guide is applied to every page. Use it to define custom macros, enumeration conventions, graph colours, cross-reference commands, source-typo policy, or house style. A conservative built-in guide is always present, and `STYLE_GUIDE_EXAMPLE.md` demonstrates a linked problem/solution handbook.

### Experimental graph redrawing

With vision enabled, the graph option asks for editable TikZ/pgfplots code. The prompt forbids guessed coordinates or statistics, but graph code must still be checked against the source.

### OpenAI-compatible service

Enter:

- the exact model name;
- the service's chat-completions address;
- its API key.

The app sends one page of extracted text per request. The API key stays in memory and is not placed in the LaTeX project or review report. The service itself may retain requests according to its own policies.

If a request fails, the app performs safe local conversion for that page and marks it **Needs review**.

## Scanned PDF options

### Read scanned pages with OCR

OCR means optical character recognition. It tries to recognize text inside a page picture.

Image inputs use OCR automatically in editable-text modes. For PDFs, enable **Read scanned PDF pages with OCR** when page text cannot be selected.

OCR is never perfectly accurate. Names, punctuation, page numbers, columns, and unusual fonts deserve special attention in the review report.

### Re-read every PDF page with OCR

Normally, the app uses the PDF's embedded text and runs OCR only when a selected page contains no embedded text. Re-reading every page ignores that shortcut.

Use this only when the embedded text is visibly wrong. It makes conversion slower.

### Document language

The main window currently offers:

| Language | Code |
|---|---|
| English | `eng` |
| Arabic | `ara` |

Choose the language already used in the source. The app preserves it; it does not translate. Arabic automatically enables right-to-left LaTeX, Unicode Arabic, bundled Arabic OCR data, and XeLaTeX compilation.

## Quality options

### Exact text match

The app compares original visible text with visible text reconstructed from the LaTeX. Line wrapping and repeated whitespace are ignored; capitalization and punctuation are not.

Use this recommended setting when text fidelity matters.

### Allow a small text difference

This is useful when an AI model applies harmless formatting that the exact comparison cannot fully reconstruct. A 5% allowed difference means a page passes at 95% similarity or higher.

The percentage is a guide, not proof that meaning is correct. Review highlighted pages.

### Easy conversion report

Every conversion creates one small `_conversion_report.txt` file. It states whether the PDF compiled, which compiler was used, whether the app repaired anything, and where to find a technical error log if compilation failed.

Temporary page images are deleted automatically. The compiler log is deleted after a successful build.

### Optional detailed review

Enable **Create a detailed per-page review** under Advanced settings when you need `review_report.html`. It opens in a normal browser and shows:

- the source and LaTeX side by side;
- the text-match percentage;
- whether safe fallback was used;
- warnings;
- exact per-page number matching, including missing and extra values;
- a button to show only pages needing review.

The report is local and does not upload anything.

The review folder also contains:

- `source_extracted.txt` — all layout-preserved extracted content;
- `verification.json` and `verification_summary.md` — whole-document word counts, token differences, text similarity, numeric-token comparison, and digit-sequence checking that tolerates damaged formula spacing without tolerating changed digits;
- `source_pdf_analysis.json` and `.md` — PDF producer, creator, date, page size, image/raster coverage, text volume, and vector-object counts.

### Ready-to-read PDF

The app asks pdfLaTeX to compile English projects and XeLaTeX to compile Arabic projects. It validates image references and common generated structures before compiling twice. If compilation still fails, the `.tex` file, short report, and `_compile.log` remain available.

Install [MiKTeX](https://miktex.org/download) on Windows or macOS. On Linux, install TeX Live from the distribution package manager. Restart the app afterward.

## Advanced options

| Option | Meaning |
|---|---|
| First page | First source page to convert; starts at 1 |
| Last page | Final page to convert; 0 means all remaining pages |
| Lines per content unit | Splits long text, document, ebook, and spreadsheet content into manageable conversion units |
| Scan quality | OCR and page-picture resolution; 220 DPI is the balanced default |
| Keep original line breaks | Useful for poetry and forms, usually off for prose |
| Keep colors in page pictures | Creates color instead of grayscale source-page images |
| Save separate files for every page | Adds a `_pages` directory for detailed editing |
| Create a detailed per-page review | Adds HTML, CSV, JSON, and page-by-page verification; off by default |
| LaTeX content only | Omits the document wrapper; intended for experienced LaTeX users |
| Maximum AI output | Increase only if a long AI page is cut off |
| AI creativity | Keep at 0 to reduce rewriting |
| AI timeout | Maximum wait for one page request |
| AI retry count | Extra connection attempts before safe fallback |
| Project style guide | Shared custom LaTeX conventions applied to every unit |

## Cancellation and existing files

Select **Cancel** to stop after the current page or network request. The finished `.tex` file is written atomically near the end, so cancelling does not replace it with a partial document.

The app asks before replacing an existing `.tex` file. Generated page and review folders are refreshed using known generated filenames; unrelated files are not recursively deleted.

## Troubleshooting

### Open PDF is disabled

Read the status message and open the conversion report. The usual causes are:

- MiKTeX or TeX Live is not installed;
- the AI produced a rare LaTeX error the automatic repair pass could not resolve;
- **Create LaTeX content only** is enabled.

The `.tex` file and technical `_compile.log` still exist. Select **Show details** to open a separate resizable window with a working scrollbar.

### OCR is not ready

Install both parts when running from source:

1. [Tesseract OCR](https://tesseract-ocr.github.io/tessdoc/Installation.html).
2. The Python packages with `python -m pip install -r requirements.txt` when running from source.

Restart the app. On Windows, the app checks `C:\Program Files\Tesseract-OCR\tesseract.exe` automatically.

### Ollama connection fails

Confirm that:

- Ollama is open/running;
- the model is installed;
- the model name exactly matches Ollama;
- the address is `http://127.0.0.1:11434/api/chat` unless Ollama was customized.

The affected page will use safe local conversion and appear in the review report.

### Online AI reports unauthorized

Check the API key, model name, and connection address. Some providers use a different OpenAI-compatible address.

### The PDF has duplicated pages

You selected **Editable text + pictures of each PDF page**. That mode intentionally includes both the source-page picture and editable text. Select **Editable text** for text only.

### Exact appearance cannot be selected

It works with PDF and image input. Extracted Word, ebook, spreadsheet, presentation, and ordinary text content has no single original rendered page for the app to reproduce exactly.

### A page has no text

If it is a scanned page, enable OCR. If it contains only an illustration, use a page-picture output style.

## Privacy checklist

- No AI: source text stays on this computer.
- Ollama: text is sent to the local Ollama address you selected.
- OpenAI-compatible: page text is sent to the entered service address.
- API keys are not written to output or review files.
- Review reports contain the book text; protect their folder like the original manuscript.
