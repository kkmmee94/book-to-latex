# Development and release guide

## Local checks

```bash
python -m pip install -e ".[dev]"
python -m ruff check .
python -m unittest discover -s tests -v
python -m build
```

PDF and OCR tests skip only when their external executable is unavailable. A release workstation should have pdfLaTeX, XeLaTeX, and Tesseract so those paths run.

## Architecture

- `book_to_latex.py` owns extraction, OCR, AI requests, verification, LaTeX generation, repair, compilation, reporting, cleanup, and the CLI.
- `book_to_latex_gui.py` is the nontechnical Tk workflow.
- `book_to_latex_streamlit.py` is an optional browser interface.
- Optional executables are discovered at runtime. They are not imported as Python libraries or assumed to be on one operating system.
- Page images are temporary unless an advanced output explicitly needs persistent assets.
- Generated-directory cleanup validates that the exact target is beside the selected output.

## Release artifacts

`.github/workflows/release.yml` builds:

- Windows portable `.exe` and installer;
- macOS `.app` inside a `.dmg`;
- Linux portable binary and `.deb`;
- Python wheel and source distribution for the CLI.

Tag a release with a semantic version such as `v1.1.0`. CI tests run before packaging. GitHub release publication requires the normal repository `contents: write` workflow permission.

## Dependency policy

Python dependencies are bounded by major version in `pyproject.toml`. External programs remain optional and are documented rather than silently downloaded by the app. The only bundled OCR model is the official Arabic `tessdata_fast` file.

## Pull requests

Keep user-facing language plain. New visible controls require a hover/help explanation. Add tests for compile failures, cleanup behavior, and language routing when those paths change.
