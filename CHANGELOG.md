# Changelog

## 1.2.0 — 2026-08-18

- Separated visual fidelity, physical page size, and page usage into clear user choices.
- Added source-size, A4, and US Letter output.
- Added compact continuous flow while retaining one-source-page-per-output-page mode.
- Added Keep photographs and Replace photographs with descriptions policies.
- Added visual inventory and enforcement for semantic graph, chart, table, equation, infographic, and technical-diagram reconstruction.
- Retains the source visual when exact semantic reconstruction would require guessing, so images never silently disappear.
- Preserves required assets while continuing to delete temporary page renders and unused extracted images.
- Prevents reconstructed source pages from spilling onto unintended extra pages.

## 1.1.0 — 2026-08-18

- Added Arabic OCR, right-to-left Unicode LaTeX, Polyglossia/Amiri, and XeLaTeX compilation.
- Added a shared language registry used by desktop, browser, core, CLI, and reports.
- Prevented AI-invented image filenames and added pre-compilation repair for unavailable graphics, malformed tables, missing package declarations, and common TikZ mistakes.
- Changed default output to LaTeX, compiled PDF, and one concise conversion report.
- Made page renders temporary and removed successful compiler logs.
- Changed exact PDF mode to reference original PDF pages without a generated PNG folder.
- Moved technical details to a separate resizable, scrollable window.
- Added cross-platform packaging and release metadata for GUI and CLI.
