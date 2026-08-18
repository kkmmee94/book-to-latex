# Troubleshooting

## No compiled PDF

Open `<name>_conversion_report.txt`. If compilation failed, it points to `<name>_compile.log`. Successful builds remove that technical log automatically.

The app repairs unavailable image references, underspecified tables, missing common package declarations, and frequent TikZ compatibility mistakes before compilation. The LaTeX source remains available even when an unusual error survives.

## Arabic text is disconnected or reversed

Confirm **Arabic** was selected before conversion and that XeLaTeX is installed. The generated preamble should contain `polyglossia`, `\setdefaultlanguage{arabic}`, and the Amiri font.

## Arabic scan returns no text

The app bundles Arabic recognition data, but it still requires the Tesseract executable. Install Tesseract and restart the app.

## Show details does not fit

**Show details** opens a separate resizable window. Use its scrollbar or mouse wheel; closing it does not stop the conversion.

## Output folder contains page images

Normal Clean, Close, and PDF Exact conversions remove temporary page renders. Persistent `_images`, `_pages`, or `_review` folders are created only by advanced output choices; rerun with those choices off to clean the known generated folders.
