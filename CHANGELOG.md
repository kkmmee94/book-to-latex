# Changelog

## 1.3.2 — 2026-08-18

- Fixed HTTPS certificate verification in packaged macOS builds by shipping a current CA bundle.
- Replaced free-text OpenAI model and endpoint fields with guided model choices, an API-key link, and a connection test.
- Added automatic discovery for Ollama, LM Studio, Jan, llama.cpp, GPT4All, Hugging Face, GGUF, Safetensors, and MLX model locations.
- Added one-click installation of a recommended Ollama text-and-vision model on Windows, macOS, and Linux.
- Replaced the Windows-only local vision-model error with cross-platform setup guidance and OpenAI fallback.
- Repaired additional AI-generated TikZ faults, including invalid triangle/marker shapes, grid coordinates, placeholder plots, colors, paired ranges, and text-mode markers.

## 1.3.1 — 2026-08-18

- Restored a consistently English application interface.
- Arabic selection now affects only the generated document's language and RTL layout; it no longer mixes Arabic into English controls.

## 1.3.0 — 2026-08-18

- Renamed the main workflows to Reconstruct and polish, Enhance a scan or lecture, and Keep original pages unchanged.
- Added a Check for updates button backed by the public GitHub release feed.
- Added automatic Arabic/Chinese script detection and on-demand official Tesseract language-data installation.
- Added Chinese Unicode XeLaTeX/CTeX compilation.
- Added further automatic repairs for boxed floats, TikZ coordinates/colors/legends, split math, Unicode math symbols, and generated probability plots.
- Repaired and compiled the reported 148-page `module01_close_layout` conversion.

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
