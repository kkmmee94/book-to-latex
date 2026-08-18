# Language support

Language support is a first-class conversion setting rather than an OCR-only detail.

## Available now

| Language | App code | OCR | LaTeX engine | Direction |
|---|---:|---|---|---|
| English | `eng` | system Tesseract English data | pdfLaTeX | left-to-right |
| Arabic | `ara` | bundled `tessdata_fast/ara.traineddata` | XeLaTeX + Polyglossia + Amiri | right-to-left |

The selected language is sent to every conversion stage:

1. OCR receives its Tesseract language code.
2. The AI prompt is told to preserve the original language and not translate or transliterate it.
3. Arabic output preserves Unicode Arabic and right-to-left reading order.
4. The LaTeX preamble selects the appropriate engine, language package, and font.
5. The conversion report records the language and compiler.

## Adding another language

1. Register its code and label in `DOCUMENT_LANGUAGES` in `book_to_latex.py`.
2. Add a language-specific preamble branch only when normal pdfLaTeX is insufficient.
3. If OCR should work without a system language pack, place the official `.traineddata` under `assets/tessdata/` and include it in the PyInstaller/wheel manifests.
4. Add an OCR, prompt, compilation, and temporary-file regression test.
5. Update this document and the user guide.

The GUI and browser interface read the same registry, so a registered language appears consistently in both.
