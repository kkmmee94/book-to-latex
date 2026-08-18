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

Normal Clean, Close, and PDF Exact conversions remove temporary page renders. A `_assets` folder is different: it contains photographs, logos, screenshots, or source visuals that the LaTeX actually references and therefore must not be deleted. The conversion report lists every required asset. Persistent `_pages` or `_review` folders are created only by advanced output choices.

## Close layout is not identical

Close layout creates editable LaTeX and therefore reconstructs the document. Choose **Exact visual copy** when colors, shapes, headers, footers, photographs, fonts and positions must remain unchanged. Exact copy is necessarily one source page per output page.

## A graph was kept as an image

The visual inventory found that its exact data or geometry could not be recovered without guessing. The app retains the original graph so it remains accurate and visible. Graphs with fully readable values are reconstructed semantically; generated random or approximate data is rejected.
