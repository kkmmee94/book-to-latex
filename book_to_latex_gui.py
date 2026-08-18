#!/usr/bin/env python3
"""Simple desktop interface for converting documents to LaTeX and PDF."""

from __future__ import annotations

import os
import platform
import queue
import subprocess
import threading
import tkinter as tk
import traceback
import webbrowser
from collections.abc import Iterable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from book_to_latex import (
    APP_VERSION,
    DEFAULT_OLLAMA_ENDPOINT,
    DEFAULT_OPENAI_COMPAT_ENDPOINT,
    IMAGE_EXTENSIONS,
    MATCH_MODE_PERCENT,
    PAGE_FLOW_COMPACT,
    PAGE_FLOW_SOURCE,
    PAGE_SIZE_A4,
    PAGE_SIZE_LETTER,
    PAGE_SIZE_SOURCE,
    PHOTO_DESCRIBE,
    PHOTO_KEEP,
    SUPPORTED_INPUT_EXTENSIONS,
    ConversionCancelled,
    check_for_updates,
    convert_book_to_latex,
    discover_local_ai,
    install_recommended_local_model,
    model_supports_vision,
    openai_connection_info,
    runtime_capabilities,
)

LOOK_CLEAN = "Reconstruct and polish (fully editable)"
LOOK_CLOSE = "Enhance a scan or lecture"
LOOK_EXACT = "Keep original pages unchanged"

COLOUR_KEEP = "Keep the original colours"
COLOUR_MONO = "Black and white"

PAGE_SIZE_LABELS = {
    "Same size as the original": PAGE_SIZE_SOURCE,
    "A4 pages": PAGE_SIZE_A4,
    "US Letter pages": PAGE_SIZE_LETTER,
}
PAGE_FLOW_LABELS = {
    "Compact — use fewer pages": PAGE_FLOW_COMPACT,
    "Keep every source page separate": PAGE_FLOW_SOURCE,
}
PHOTO_LABELS = {
    "Keep real photographs": PHOTO_KEEP,
    "Replace photographs with descriptions": PHOTO_DESCRIBE,
}
LANGUAGE_LABELS = {
    "Detect automatically": "auto",
    "English": "eng",
    "Arabic": "ara",
    "Chinese (Simplified)": "chi_sim",
    "Chinese (Traditional)": "chi_tra",
}

AI_AUTO = "Automatic (recommended)"
AI_OLLAMA = "Local AI — private on this computer"
AI_OPENAI = "OpenAI — online"
AI_NONE = "Basic conversion without AI"

OPENAI_MODEL_LABELS = {
    "Balanced quality and cost — GPT-5.6 Terra (recommended)": "gpt-5.6-terra",
    "Best quality — GPT-5.6 Sol": "gpt-5.6-sol",
    "Lower cost — GPT-5.6 Luna": "gpt-5.6-luna",
}

class ToolTip:
    def __init__(self, widget: tk.Widget, text: str, delay_ms: int = 450) -> None:
        self.widget = widget
        self.text = text
        self.delay_ms = delay_ms
        self.after_id: str | None = None
        self.window: tk.Toplevel | None = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event: tk.Event | None = None) -> None:
        self._cancel()
        self.after_id = self.widget.after(self.delay_ms, self._show)

    def _cancel(self) -> None:
        if self.after_id is not None:
            self.widget.after_cancel(self.after_id)
            self.after_id = None

    def _show(self) -> None:
        self.after_id = None
        if self.window is not None or not self.text:
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        self.window.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self.window,
            text=self.text,
            justify="left",
            wraplength=390,
            background="#fff8d8",
            foreground="#172033",
            relief="solid",
            borderwidth=1,
            padx=10,
            pady=8,
            font=("Segoe UI", 9),
        ).pack()

    def _hide(self, _event: tk.Event | None = None) -> None:
        self._cancel()
        if self.window is not None:
            self.window.destroy()
            self.window = None


def add_tooltip(widgets: tk.Widget | Iterable[tk.Widget], text: str) -> None:
    if isinstance(widgets, tk.Widget):
        widgets = [widgets]
    for widget in widgets:
        ToolTip(widget, text)


def open_in_system(path: Path) -> None:
    if os.name == "nt":
        os.startfile(str(path))
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class BookToLatexGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Book to LaTeX & PDF")
        self.root.geometry("960x850")
        self.root.minsize(820, 720)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.running = False
        self.last_result: dict[str, object] | None = None
        self.details_window: tk.Toplevel | None = None
        self.details_text: tk.Text | None = None
        self.log_messages: list[str] = []
        self.capabilities = runtime_capabilities()
        self.local_models: list[dict[str, object]] = []
        self.local_model_by_label: dict[str, dict[str, object]] = {}

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.look_var = tk.StringVar(value=LOOK_CLEAN)
        self.colour_var = tk.StringVar(value=COLOUR_KEEP)
        self.page_size_var = tk.StringVar(value="A4 pages")
        self.page_flow_var = tk.StringVar(value="Compact — use fewer pages")
        self.photo_var = tk.StringVar(value="Keep real photographs")

        self.language_by_label = LANGUAGE_LABELS.copy()
        default_language = "Detect automatically" if "Detect automatically" in self.language_by_label else next(
            iter(self.language_by_label)
        )
        self.language_var = tk.StringVar(value=default_language)

        # Expert settings are intentionally hidden from the main workflow.
        self.ai_choice_var = tk.StringVar(value=AI_AUTO)
        self.local_model_var = tk.StringVar()
        self.openai_model_by_label = OPENAI_MODEL_LABELS.copy()
        self.openai_model_var = tk.StringVar(value=next(iter(self.openai_model_by_label)))
        self.api_key_var = tk.StringVar(value=os.environ.get("OPENAI_API_KEY", ""))
        self.ai_connection_status_var = tk.StringVar(value="Choose Automatic unless you want a specific service.")
        self.start_page_var = tk.StringVar(value="1")
        self.end_page_var = tk.StringVar(value="0")
        self.lines_per_unit_var = tk.StringVar(value="70")
        self.ocr_force_var = tk.BooleanVar(value=False)
        self.ocr_dpi_var = tk.StringVar(value="220")
        self.max_tokens_var = tk.StringVar(value="4096")
        self.temperature_var = tk.StringVar(value="0.0")
        self.timeout_var = tk.StringVar(value="180")
        self.retries_var = tk.StringVar(value="2")
        self.allowed_difference_var = tk.StringVar(value="10.0")
        self.keep_page_files_var = tk.BooleanVar(value=False)
        self.detailed_review_var = tk.BooleanVar(value=False)
        self.keep_line_breaks_var = tk.BooleanVar(value=False)
        self.redraw_graphs_var = tk.BooleanVar(value=True)
        self.style_guide_var = tk.StringVar()

        self._configure_styles()
        self._build_ui()
        self._update_look_explanation()
        self.root.after(100, self._drain_queue)
        self.root.after(250, self._refresh_local_models)

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("vista" if os.name == "nt" else "clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 22, "bold"), foreground="#243b6b")
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#526079")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 11, "bold"), foreground="#243b6b")
        style.configure("Choice.TCheckbutton", font=("Segoe UI", 10, "bold"))
        style.configure("Choice.TRadiobutton", font=("Segoe UI", 10, "bold"))
        style.configure("Hint.TLabel", foreground="#667085")
        style.configure("Info.TLabel", foreground="#243b6b", font=("Segoe UI", 10))
        style.configure("Action.TButton", font=("Segoe UI", 11, "bold"), padding=(18, 7))

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=(20, 8))
        main.pack(fill="both", expand=True)

        heading = ttk.Frame(main)
        heading.pack(fill="x", pady=(0, 7))
        ttk.Label(heading, text="Book → LaTeX & PDF", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            heading,
            text="Choose how the result should look. The app handles OCR, AI, checking and PDF compilation for you.",
            style="Subtitle.TLabel",
        ).pack(side="left", anchor="w", pady=(3, 0))
        ttk.Button(heading, text="How it works", command=self._show_help).pack(side="right", padx=(6, 0))
        ttk.Button(heading, text="More settings", command=self._open_advanced).pack(side="right")
        self.update_button = ttk.Button(
            heading,
            text=f"Check for updates (v{APP_VERSION})",
            command=self._check_for_updates,
        )
        self.update_button.pack(side="right", padx=(0, 6))
        add_tooltip(
            self.update_button,
            "Checks the public GitHub release page and offers the correct installer when a newer version exists.",
        )

        files = ttk.LabelFrame(main, text="1. Choose your file", style="Section.TLabelframe")
        files.pack(fill="x", pady=6)
        files.columnconfigure(1, weight=1)
        input_label = ttk.Label(files, text="Original file")
        input_label.grid(row=0, column=0, sticky="w", padx=10, pady=6)
        input_entry = ttk.Entry(files, textvariable=self.input_var)
        input_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=6)
        input_button = ttk.Button(files, text="Choose file…", command=self._pick_input)
        input_button.grid(row=0, column=2, padx=10, pady=6)
        add_tooltip(
            [input_label, input_entry, input_button],
            "Choose the PDF, document, ebook, spreadsheet, presentation, image or text file you want to turn into LaTeX.",
        )
        output_label = ttk.Label(files, text="Save finished project as")
        output_label.grid(row=1, column=0, sticky="w", padx=10, pady=6)
        output_entry = ttk.Entry(files, textvariable=self.output_var)
        output_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=6)
        output_button = ttk.Button(files, text="Choose location…", command=self._pick_output)
        output_button.grid(row=1, column=2, padx=10, pady=6)
        add_tooltip(
            [output_label, output_entry, output_button],
            "The editable .tex file, compiled PDF and review report will be saved together at this location.",
        )

        appearance = ttk.LabelFrame(
            main, text="2. How should the result look?", style="Section.TLabelframe"
        )
        appearance.pack(fill="x", pady=6)
        choices = [
            (
                LOOK_CLEAN,
                "Rebuilds the entire document as polished editable LaTeX. Every word stays exact, while page design, tables and spacing are improved instead of copied.",
            ),
            (
                LOOK_CLOSE,
                "Best for scanned lectures, notes and older documents. Makes them clearer and editable while keeping the character and structure of the source.",
            ),
            (
                LOOK_EXACT,
                "Keeps every original page visually unchanged. Searchable text stays searchable when it existed in the source, but the LaTeX text is not reconstructed.",
            ),
        ]
        for row, (value, description) in enumerate(choices):
            radio = ttk.Radiobutton(
                appearance,
                text=value,
                value=value,
                variable=self.look_var,
                style="Choice.TRadiobutton",
                command=self._update_look_explanation,
            )
            radio.grid(row=row * 2, column=0, sticky="w", padx=12, pady=(5, 0))
            ttk.Label(
                appearance,
                text=description,
                style="Hint.TLabel",
                wraplength=790,
            ).grid(row=row * 2 + 1, column=0, sticky="w", padx=(36, 12), pady=(0, 3))

        preferences = ttk.LabelFrame(
            main, text="3. Page, language and pictures", style="Section.TLabelframe"
        )
        preferences.pack(fill="x", pady=6)
        preferences.columnconfigure(1, weight=1)
        preferences.columnconfigure(3, weight=1)

        ttk.Label(preferences, text="Document language").grid(
            row=0, column=0, sticky="w", padx=(12, 6), pady=7
        )
        language = ttk.Combobox(
            preferences,
            textvariable=self.language_var,
            values=list(self.language_by_label),
            state="readonly",
            width=24,
        )
        language.grid(row=0, column=1, sticky="ew", padx=(0, 14), pady=7)
        add_tooltip(
            language,
            "Automatic detection is recommended. If Arabic or Chinese OCR data is missing, the app downloads the official Tesseract language file and remembers it. Choosing Arabic affects the generated document's RTL layout, not the app interface language.",
        )

        ttk.Label(preferences, text="Colour").grid(
            row=0, column=2, sticky="w", padx=(8, 5), pady=7
        )
        self.colour_keep_button = ttk.Radiobutton(
            preferences,
            text=COLOUR_KEEP,
            value=COLOUR_KEEP,
            variable=self.colour_var,
        )
        self.colour_keep_button.grid(row=0, column=3, sticky="w", padx=5, pady=7)
        self.colour_mono_button = ttk.Radiobutton(
            preferences,
            text=COLOUR_MONO,
            value=COLOUR_MONO,
            variable=self.colour_var,
        )
        self.colour_mono_button.grid(row=0, column=4, sticky="w", padx=(5, 12), pady=7)
        add_tooltip(
            [self.colour_keep_button, self.colour_mono_button],
            "Enhanced mode reproduces visible colors in headings, boxes, charts and diagrams. Original-pages mode always keeps the source colors.",
        )

        ttk.Label(preferences, text="Finished page size").grid(
            row=1, column=0, sticky="w", padx=(12, 6), pady=7
        )
        self.page_size_combo = ttk.Combobox(
            preferences,
            textvariable=self.page_size_var,
            values=list(PAGE_SIZE_LABELS),
            state="readonly",
            width=28,
        )
        self.page_size_combo.grid(row=1, column=1, sticky="ew", padx=(0, 14), pady=7)
        add_tooltip(
            self.page_size_combo,
            "Same size preserves the source PDF's physical page shape. A4 or US Letter places the result on that paper size. This is separate from how faithfully content is reconstructed.",
        )

        ttk.Label(preferences, text="Use of pages").grid(
            row=1, column=2, sticky="w", padx=(8, 5), pady=7
        )
        self.page_flow_combo = ttk.Combobox(
            preferences,
            textvariable=self.page_flow_var,
            values=list(PAGE_FLOW_LABELS),
            state="readonly",
            width=34,
        )
        self.page_flow_combo.grid(
            row=1, column=3, columnspan=2, sticky="ew", padx=(5, 12), pady=7
        )
        add_tooltip(
            self.page_flow_combo,
            "Compact lets editable content flow naturally and use fewer pages. Keep every source page separate preserves each source page boundary. Original-pages mode always keeps one source page per output page.",
        )

        ttk.Label(preferences, text="Real photographs").grid(
            row=2, column=0, sticky="w", padx=(12, 6), pady=7
        )
        self.photo_combo = ttk.Combobox(
            preferences,
            textvariable=self.photo_var,
            values=list(PHOTO_LABELS),
            state="readonly",
            width=34,
        )
        self.photo_combo.grid(
            row=2, column=1, columnspan=2, sticky="ew", padx=(0, 14), pady=7
        )
        add_tooltip(
            self.photo_combo,
            "Keep real photographs as required project assets, or replace photographs of people, animals, places and objects with concise written descriptions. Graphs, tables, equations and technical diagrams are recreated as editable LaTeX either way.",
        )
        ttk.Label(
            preferences,
            text="Graphs, tables, equations and technical diagrams are recreated as editable LaTeX automatically.",
            style="Hint.TLabel",
            wraplength=420,
        ).grid(row=2, column=3, columnspan=2, sticky="w", padx=(5, 12), pady=7)

        self.explanation_label = ttk.Label(
            main,
            text="",
            style="Info.TLabel",
            wraplength=820,
            padding=(12, 5),
        )
        self.explanation_label.pack(fill="x", pady=(5, 2))
        self.readiness_label = ttk.Label(
            main,
            text="Checking local tools…",
            style="Hint.TLabel",
        )
        self.readiness_label.pack(anchor="w", padx=12)
        ttk.Button(
            main,
            text="Choose or set up AI…",
            command=lambda: self._open_advanced("ai"),
        ).pack(anchor="w", padx=12, pady=(4, 2))

        ttk.Label(
            main,
            text="You will receive: a LaTeX file, compiled PDF, and conversion report. Only required source/photo assets are kept; temporary page renders are deleted.",
            style="Hint.TLabel",
        ).pack(anchor="w", padx=12, pady=(2, 8))

        self.progress = ttk.Progressbar(main, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(4, 5))
        self.status_label = ttk.Label(main, text="Ready — choose a file")
        self.status_label.pack(fill="x")

        actions = ttk.Frame(main)
        actions.pack(fill="x", pady=(8, 0))
        self.start_button = ttk.Button(
            actions,
            text="Create LaTeX and PDF",
            style="Action.TButton",
            command=self._start,
        )
        self.start_button.pack(side="left")
        self.cancel_button = ttk.Button(
            actions, text="Cancel", command=self._cancel, state="disabled"
        )
        self.cancel_button.pack(side="left", padx=(8, 0))
        self.details_button = ttk.Button(
            actions, text="Show details", command=self._toggle_details
        )
        self.details_button.pack(side="left", padx=(8, 0))
        self.open_folder_button = ttk.Button(
            actions, text="Open output folder", command=self._open_output_folder, state="disabled"
        )
        self.open_folder_button.pack(side="right", padx=(6, 0))
        self.open_review_button = ttk.Button(
            actions, text="Open report", command=self._open_review, state="disabled"
        )
        self.open_review_button.pack(side="right", padx=(6, 0))
        self.open_pdf_button = ttk.Button(
            actions, text="Open finished PDF", command=self._open_pdf, state="disabled"
        )
        self.open_pdf_button.pack(side="right")

    def _update_look_explanation(self) -> None:
        messages = {
            LOOK_CLEAN: "Reconstructs every word into editable LaTeX and applies a clean professional design. Tables are redesigned and chosen graphs remain visible even when their appearance changes.",
            LOOK_CLOSE: "Enhances scans or lecture notes while retaining their recognizable structure. Vision AI restores mathematics, tables, visuals, colors and meaningful page elements.",
            LOOK_EXACT: "Keeps original pages visually unchanged. It always uses one source page per output page; choose Same size, A4 or US Letter.",
        }
        look = self.look_var.get()
        self.explanation_label.configure(text=messages[look])
        exact = look == LOOK_EXACT
        if exact:
            self.page_flow_var.set("Keep every source page separate")
            self.page_size_var.set("Same size as the original")
            self.colour_var.set(COLOUR_KEEP)
        self.page_flow_combo.configure(state="disabled" if exact else "readonly")
        self.photo_combo.configure(state="disabled" if exact else "readonly")
        self.colour_keep_button.configure(state="disabled" if exact else "normal")
        self.colour_mono_button.configure(state="disabled" if exact else "normal")
        self._update_suggested_output_name()

    def _update_suggested_output_name(self) -> None:
        input_text = self.input_var.get().strip()
        if not input_text:
            return
        source = Path(input_text)
        current_text = self.output_var.get().strip()
        suggested_endings = ("_latex.tex", "_close_layout.tex", "_exact_copy.tex")
        if current_text and not any(Path(current_text).name == f"{source.stem}{ending}" for ending in suggested_endings):
            return
        suffix = {
            LOOK_CLEAN: "_latex.tex",
            LOOK_CLOSE: "_close_layout.tex",
            LOOK_EXACT: "_exact_copy.tex",
        }[self.look_var.get()]
        self.output_var.set(str(source.with_name(f"{source.stem}{suffix}")))

    def _pick_input(self) -> None:
        supported_pattern = " ".join(
            f"*{extension}" for extension in sorted(SUPPORTED_INPUT_EXTENSIONS)
        )
        selected = filedialog.askopenfilename(
            title="Choose the file to convert",
            filetypes=[
                ("Supported documents and images", supported_pattern),
                ("PDF documents", "*.pdf"),
                ("Documents and ebooks", "*.docx *.odt *.rtf *.epub *.html *.eml"),
                ("Images", "*.png *.jpg *.jpeg *.tif *.tiff *.webp *.heic *.heif"),
                ("All files", "*.*"),
            ],
        )
        if not selected:
            return
        source = Path(selected)
        self.input_var.set(selected)
        if not self.output_var.get().strip():
            self._update_suggested_output_name()
        visual = source.suffix.lower() == ".pdf" or source.suffix.lower() in IMAGE_EXTENSIONS
        if self.look_var.get() == LOOK_EXACT and not visual:
            self.look_var.set(LOOK_CLOSE)
            self._update_look_explanation()

    def _pick_output(self) -> None:
        initial = Path(self.output_var.get()) if self.output_var.get().strip() else None
        selected = filedialog.asksaveasfilename(
            title="Save the finished LaTeX project",
            defaultextension=".tex",
            initialdir=str(initial.parent) if initial else None,
            initialfile=initial.name if initial else "converted_document.tex",
            filetypes=[("LaTeX project", "*.tex")],
        )
        if selected:
            self.output_var.set(selected)

    def _show_help(self) -> None:
        messagebox.showinfo(
            "How it works",
            "1. Choose the original file.\n\n"
            "2. Choose Reconstruct and polish, Enhance a scan or lecture, or Keep original pages unchanged.\n\n"
            "3. Choose the page size and whether editable content should be compact or keep every source page separate.\n\n"
            "4. Choose whether real photographs should be kept or replaced with descriptions. Graphs, tables and technical diagrams are recreated as LaTeX automatically.\n\n"
            "5. Select Create LaTeX and PDF.\n\n"
            "Original-pages mode is page-for-page and never compact. Reconstructed and enhanced modes can be compact. The app automatically handles models, OCR, compilation and checking.",
        )

    def _open_advanced(self, initial_tab: str = "") -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("More settings")
        dialog.geometry("680x560")
        dialog.transient(self.root)
        notebook = ttk.Notebook(dialog)
        notebook.pack(fill="both", expand=True, padx=12, pady=12)

        pages_tab = ttk.Frame(notebook, padding=14)
        ai_tab = ttk.Frame(notebook, padding=14)
        style_tab = ttk.Frame(notebook, padding=14)
        notebook.add(pages_tab, text="Pages and checking")
        notebook.add(ai_tab, text="AI connection")
        notebook.add(style_tab, text="Style and graphs")

        self._advanced_field(pages_tab, 0, "First page", self.start_page_var, "1 starts at the beginning.")
        self._advanced_field(pages_tab, 1, "Last page", self.end_page_var, "0 means the final page.")
        self._advanced_field(
            pages_tab,
            2,
            "Lines per content unit",
            self.lines_per_unit_var,
            "Splits long documents into manageable AI requests.",
        )
        self._advanced_field(pages_tab, 3, "OCR quality (DPI)", self.ocr_dpi_var, "220 is balanced.")
        self._advanced_field(
            pages_tab,
            4,
            "Allowed text difference (%)",
            self.allowed_difference_var,
            "Numbers are always checked separately.",
        )
        ttk.Checkbutton(
            pages_tab,
            text="Force OCR on every selected PDF page",
            variable=self.ocr_force_var,
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Checkbutton(
            pages_tab,
            text="Keep original line breaks (useful for poetry)",
            variable=self.keep_line_breaks_var,
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Checkbutton(
            pages_tab,
            text="Save separate source and LaTeX files for every page",
            variable=self.keep_page_files_var,
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=8)
        detailed_review = ttk.Checkbutton(
            pages_tab,
            text="Create a detailed per-page review (uses more disk space)",
            variable=self.detailed_review_var,
        )
        detailed_review.grid(row=8, column=0, columnspan=2, sticky="w", pady=8)
        add_tooltip(
            detailed_review,
            "Normally the app creates one short report. Enable this only when you want page-by-page source text, verification data, and an HTML review.",
        )

        ai_tab.columnconfigure(1, weight=1)
        ttk.Label(ai_tab, text="AI service").grid(row=0, column=0, sticky="w", pady=7)
        ai_choice = ttk.Combobox(
            ai_tab,
            textvariable=self.ai_choice_var,
            values=[AI_AUTO, AI_OLLAMA, AI_OPENAI, AI_NONE],
            state="readonly",
        )
        ai_choice.grid(row=0, column=1, sticky="ew", pady=7)
        ai_choice.bind("<<ComboboxSelected>>", lambda _event: self._update_advanced_ai_state())
        add_tooltip(
            ai_choice,
            "Automatic uses a ready local model first. Choose OpenAI when you prefer online AI, or Basic conversion when you do not want AI.",
        )

        ttk.Label(ai_tab, text="Local model").grid(row=1, column=0, sticky="w", pady=7)
        self.local_model_combo = ttk.Combobox(
            ai_tab,
            textvariable=self.local_model_var,
            values=list(self.local_model_by_label),
            state="readonly",
        )
        self.local_model_combo.grid(row=1, column=1, sticky="ew", pady=7)
        local_actions = ttk.Frame(ai_tab)
        local_actions.grid(row=2, column=1, sticky="w", pady=(0, 10))
        ttk.Button(local_actions, text="Search this computer again", command=self._refresh_local_models).pack(
            side="left"
        )
        ttk.Button(
            local_actions,
            text="Install recommended local AI…",
            command=self._install_local_ai,
        ).pack(side="left", padx=(8, 0))

        ttk.Label(ai_tab, text="OpenAI model").grid(row=3, column=0, sticky="w", pady=7)
        self.openai_model_combo = ttk.Combobox(
            ai_tab,
            textvariable=self.openai_model_var,
            values=list(self.openai_model_by_label),
            state="readonly",
        )
        self.openai_model_combo.grid(row=3, column=1, sticky="ew", pady=7)
        add_tooltip(
            self.openai_model_combo,
            "The recommended balanced model is already selected. Every listed model can read text and page images.",
        )

        ttk.Label(ai_tab, text="OpenAI API key").grid(row=4, column=0, sticky="w", pady=7)
        self.api_key_entry = ttk.Entry(ai_tab, textvariable=self.api_key_var, show="•")
        self.api_key_entry.grid(row=4, column=1, sticky="ew", pady=7)
        add_tooltip(
            self.api_key_entry,
            "Paste the private key from your OpenAI API account. The app uses it only for your conversion requests.",
        )
        online_actions = ttk.Frame(ai_tab)
        online_actions.grid(row=5, column=1, sticky="w", pady=(0, 8))
        ttk.Button(
            online_actions,
            text="Get an OpenAI API key",
            command=lambda: webbrowser.open("https://platform.openai.com/api-keys"),
        ).pack(side="left")
        ttk.Button(
            online_actions,
            text="Test connection and choose model",
            command=self._test_openai_connection,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(
            ai_tab,
            textvariable=self.ai_connection_status_var,
            style="Hint.TLabel",
            wraplength=500,
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(2, 10))

        self._advanced_field(ai_tab, 7, "Maximum AI output", self.max_tokens_var, "4096 suits dense pages.")
        self._advanced_field(ai_tab, 8, "AI creativity", self.temperature_var, "0 preserves source fidelity.")
        self._advanced_field(ai_tab, 9, "Timeout (seconds)", self.timeout_var, "Maximum wait per page.")
        self._advanced_field(ai_tab, 10, "Retry count", self.retries_var, "Extra attempts after a connection problem.")
        ttk.Label(
            ai_tab,
            text="You never need to type a model name or connection address. Automatic uses a detected local model; OpenAI uses the choice above.",
            style="Hint.TLabel",
            wraplength=610,
        ).grid(row=11, column=0, columnspan=2, sticky="w", pady=12)

        style_tab.columnconfigure(1, weight=1)
        ttk.Label(style_tab, text="Project style guide").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Entry(style_tab, textvariable=self.style_guide_var).grid(
            row=0, column=1, sticky="ew", pady=8
        )
        ttk.Button(style_tab, text="Choose…", command=self._pick_style_guide).grid(
            row=0, column=2, padx=(8, 0), pady=8
        )
        ttk.Checkbutton(
            style_tab,
            text="Recreate graphs, tables and technical diagrams as editable LaTeX",
            variable=self.redraw_graphs_var,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=10)
        ttk.Label(
            style_tab,
            text="These controls are optional. The built-in style already preserves words and numbers, uses standard mathematics environments, and produces a compiled document.",
            style="Hint.TLabel",
            wraplength=610,
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=12)

        ttk.Button(dialog, text="Done", command=dialog.destroy).pack(pady=(0, 12))
        self._update_advanced_ai_state()
        if initial_tab == "ai":
            notebook.select(ai_tab)

    def _advanced_field(
        self,
        parent: ttk.Frame,
        row: int,
        label_text: str,
        variable: tk.StringVar,
        help_text: str,
        *,
        secret: bool = False,
    ) -> None:
        parent.columnconfigure(1, weight=1)
        label = ttk.Label(parent, text=label_text)
        label.grid(row=row, column=0, sticky="w", padx=(0, 12), pady=7)
        entry = ttk.Entry(parent, textvariable=variable, show="•" if secret else "")
        entry.grid(row=row, column=1, sticky="ew", pady=7)
        add_tooltip([label, entry], help_text)

    def _update_advanced_ai_state(self) -> None:
        if not hasattr(self, "local_model_combo") or not self.local_model_combo.winfo_exists():
            return
        choice = self.ai_choice_var.get()
        if choice == AI_OLLAMA:
            self.local_model_combo.configure(state="readonly" if self.local_model_by_label else "disabled")
            self.openai_model_combo.configure(state="disabled")
            self.api_key_entry.configure(state="disabled")
        elif choice == AI_OPENAI:
            self.local_model_combo.configure(state="disabled")
            self.openai_model_combo.configure(state="readonly")
            self.api_key_entry.configure(state="normal")
        else:
            self.local_model_combo.configure(state="disabled")
            self.openai_model_combo.configure(state="disabled")
            self.api_key_entry.configure(state="disabled")

    def _pick_style_guide(self) -> None:
        selected = filedialog.askopenfilename(
            title="Choose a style guide",
            filetypes=[("Style guide", "*.md *.txt"), ("All files", "*.*")],
        )
        if selected:
            self.style_guide_var.set(selected)

    def _refresh_local_models(self) -> None:
        self.readiness_label.configure(text="Checking local AI models…")

        def load() -> None:
            self.events.put(("local_models", discover_local_ai(timeout=2.0, auto_start_ollama=True)))

        threading.Thread(target=load, daemon=True).start()

    def _test_openai_connection(self) -> None:
        key = self.api_key_var.get().strip()
        if not key:
            messagebox.showinfo(
                "OpenAI needs an API key",
                "Select Get an OpenAI API key, copy your private key, and paste it here. The model is already selected for you.",
            )
            return
        self.ai_connection_status_var.set("Testing OpenAI and loading the models available to your account…")

        def test() -> None:
            self.events.put(("openai_connection", openai_connection_info(key, timeout=15)))

        threading.Thread(target=test, daemon=True).start()

    def _install_local_ai(self) -> None:
        if not messagebox.askyesno(
            "Install recommended local AI?",
            "The app will ask Ollama to download Qwen 3.5 9B, which can read both text and page images. The download is several gigabytes and may take a while. Continue?",
        ):
            return
        self.ai_connection_status_var.set("Installing the recommended local AI model…")
        self.readiness_label.configure(text="Installing local AI — this can take a while…")

        def install() -> None:
            self.events.put(("local_install", install_recommended_local_model()))

        threading.Thread(target=install, daemon=True).start()

    def _check_for_updates(self) -> None:
        self.update_button.configure(state="disabled", text="Checking for updates…")

        def check() -> None:
            self.events.put(("update_check", check_for_updates(timeout=10)))

        threading.Thread(target=check, daemon=True).start()

    def _handle_update_check(self, info: dict[str, object]) -> None:
        self.update_button.configure(
            state="normal",
            text=f"Check for updates (v{APP_VERSION})",
        )
        if not info.get("success"):
            if messagebox.askyesno(
                "Could not check for updates",
                "The automatic check could not reach GitHub.\n\n"
                f"{info.get('error') or 'Unknown connection error'}\n\n"
                "Open the official release page in your browser instead?",
            ):
                webbrowser.open(str(info.get("release_url")))
            return
        latest = str(info.get("latest_version") or "")
        if not info.get("update_available"):
            messagebox.showinfo(
                "You are up to date",
                f"Book to LaTeX v{APP_VERSION} is the newest public version.",
            )
            return
        if messagebox.askyesno(
            "Update available",
            f"Book to LaTeX v{latest} is available.\n\nOpen the official download page now?",
        ):
            webbrowser.open(str(info.get("release_url")))

    def _handle_local_models(self, info: dict[str, object]) -> None:
        self.local_models = [dict(model) for model in (info.get("models") or [])]
        self.local_model_by_label = {}
        for model in self.local_models:
            name = str(model.get("name") or "")
            source = str(model.get("source") or "Local AI")
            capability = "text + images" if model.get("vision") else "text"
            label = f"{name} — {source}, {capability}"
            self.local_model_by_label[label] = model
        if hasattr(self, "local_model_combo") and self.local_model_combo.winfo_exists():
            self.local_model_combo.configure(values=list(self.local_model_by_label))
        selected = self._find_text_model()
        if selected:
            label = next(
                (label for label, item in self.local_model_by_label.items() if item is selected),
                "",
            )
            if label and self.local_model_var.get() not in self.local_model_by_label:
                self.local_model_var.set(label)
        model_files = list(info.get("model_files") or [])
        if info.get("available") and self.local_models:
            has_vision = bool(self._find_vision_model())
            vision_note = "vision AI ready" if has_vision else "vision AI not installed"
            uncensored_note = (
                "UNCENSORED local model ready · "
                if any(str(model.get("name")) == "book-latex-qwen3-local-uncensored:8b" for model in self.local_models)
                else ""
            )
            self.readiness_label.configure(
                text=f"Ready: {uncensored_note}{len(self.local_models)} local model(s) detected; {vision_note}."
            )
            self.ai_connection_status_var.set(
                f"Found {len(self.local_models)} ready local model(s). Automatic will choose the right one."
            )
        elif model_files:
            self.readiness_label.configure(
                text=f"Found {len(model_files)} local model file(s), but their local AI app/server is not running."
            )
            self.ai_connection_status_var.set(
                "Local model files were found. Start Ollama, LM Studio, Jan, GPT4All, or llama.cpp, then select Search this computer again."
            )
        else:
            self.readiness_label.configure(
                text="Ready without local AI. OCR and PDF compilation are available."
            )
            self.ai_connection_status_var.set(
                "No ready local AI was found. You can use OpenAI or install the recommended local model."
            )
        self._update_advanced_ai_state()

    def _handle_openai_connection(self, info: dict[str, object]) -> None:
        if not info.get("available"):
            self.ai_connection_status_var.set(f"OpenAI connection failed: {info.get('error')}")
            messagebox.showerror("OpenAI could not connect", str(info.get("error") or "Unknown error"))
            return
        models = [str(model) for model in (info.get("models") or [])]
        for model in models:
            if model not in self.openai_model_by_label.values():
                self.openai_model_by_label[f"Available to your account — {model}"] = model
        if models:
            selected_model = models[0]
            selected_label = next(
                label for label, model in self.openai_model_by_label.items() if model == selected_model
            )
            self.openai_model_var.set(selected_label)
        if hasattr(self, "openai_model_combo") and self.openai_model_combo.winfo_exists():
            self.openai_model_combo.configure(values=list(self.openai_model_by_label))
        self.ai_connection_status_var.set(
            "OpenAI is connected. A page-image-capable model has been selected automatically."
        )

    def _handle_local_install(self, info: dict[str, object]) -> None:
        if info.get("success"):
            self.ai_connection_status_var.set("The recommended local AI is installed and ready.")
            self.ai_choice_var.set(AI_AUTO)
            self._refresh_local_models()
            messagebox.showinfo("Local AI ready", "The recommended text-and-image model is installed. Automatic can now use it.")
            return
        self.ai_connection_status_var.set(f"Local AI installation could not finish: {info.get('error')}")
        if info.get("needs_ollama") and messagebox.askyesno(
            "Install Ollama first",
            "Ollama is required to run the recommended model privately. Open the Ollama download page now?",
        ):
            webbrowser.open("https://ollama.com/download")
            return
        messagebox.showerror("Local AI installation could not finish", str(info.get("error") or "Unknown error"))

    def _find_text_model(self) -> dict[str, object] | None:
        preferred = [
            "book-latex-qwen3-local-uncensored:8b",
            "book-latex-qwen3:8b",
            "qwen3:8b",
        ]
        for name in preferred:
            found = next((model for model in self.local_models if str(model.get("name")) == name), None)
            if found:
                return found
        return next(
            (model for model in self.local_models if not bool(model.get("vision"))),
            self.local_models[0] if self.local_models else None,
        )

    def _find_vision_model(self) -> dict[str, object] | None:
        preferred = ["book-latex-qwen35-vision:9b", "qwen3.5:9b"]
        for name in preferred:
            found = next((model for model in self.local_models if str(model.get("name")) == name), None)
            if found:
                return found
        return next((model for model in self.local_models if bool(model.get("vision"))), None)

    @staticmethod
    def _looks_visual(model_name: str) -> bool:
        return model_supports_vision(model_name)

    def _read_numbers(self) -> dict[str, Any] | None:
        try:
            return {
                "start_page": int(self.start_page_var.get()),
                "end_page": int(self.end_page_var.get()),
                "lines_per_page": int(self.lines_per_unit_var.get()),
                "ocr_dpi": int(self.ocr_dpi_var.get()),
                "max_tokens": int(self.max_tokens_var.get()),
                "temperature": float(self.temperature_var.get()),
                "timeout": float(self.timeout_var.get()),
                "retries": int(self.retries_var.get()),
                "match_error_percent": float(self.allowed_difference_var.get()),
            }
        except ValueError:
            messagebox.showerror(
                "Check advanced settings",
                "One of the numeric More settings fields contains invalid text.",
            )
            return None

    def _choose_ai(self, look: str) -> tuple[str, str, str, bool, bool, str] | None:
        choice = self.ai_choice_var.get()
        if look == LOOK_EXACT:
            return "openai", "", DEFAULT_OPENAI_COMPAT_ENDPOINT, True, False, ""
        if choice == AI_NONE:
            return "openai", "", DEFAULT_OPENAI_COMPAT_ENDPOINT, True, False, ""
        if choice == AI_OPENAI:
            key = self.api_key_var.get().strip() or os.environ.get("OPENAI_API_KEY", "").strip()
            if not key:
                messagebox.showerror(
                    "OpenAI needs one API key",
                    "The OpenAI model is already selected. Select Choose or set up AI, get an OpenAI API key, and paste it into the single key box.",
                )
                self._open_advanced("ai")
                return None
            model = self.openai_model_by_label.get(
                self.openai_model_var.get(),
                "gpt-5.6-terra",
            )
            return "openai", model, DEFAULT_OPENAI_COMPAT_ENDPOINT, False, True, key
        if choice == AI_OLLAMA:
            selected = self.local_model_by_label.get(self.local_model_var.get()) or self._find_text_model()
            if not selected:
                messagebox.showerror(
                    "No ready local AI found",
                    "Select Choose or set up AI, then search again, install the recommended local AI, or choose OpenAI.",
                )
                self._open_advanced("ai")
                return None
            return (
                str(selected.get("provider") or "ollama"),
                str(selected.get("name") or ""),
                str(selected.get("endpoint") or DEFAULT_OLLAMA_ENDPOINT),
                False,
                bool(selected.get("vision")),
                "",
            )

        if look == LOOK_CLOSE:
            selected = self._find_vision_model()
            if selected:
                return (
                    str(selected.get("provider") or "ollama"),
                    str(selected.get("name") or ""),
                    str(selected.get("endpoint") or DEFAULT_OLLAMA_ENDPOINT),
                    False,
                    True,
                    "",
                )
            key = self.api_key_var.get().strip() or os.environ.get("OPENAI_API_KEY", "").strip()
            if key:
                model = self.openai_model_by_label.get(self.openai_model_var.get(), "gpt-5.6-terra")
                return "openai", model, DEFAULT_OPENAI_COMPAT_ENDPOINT, False, True, key
            messagebox.showerror(
                "One AI choice is needed for enhanced pages",
                "No page-image model is ready yet. Select Choose or set up AI, then either paste an OpenAI API key or install the recommended private local AI. You will not need to type a model name.",
            )
            self._open_advanced("ai")
            return None
        selected = self._find_text_model()
        if selected:
            return (
                str(selected.get("provider") or "ollama"),
                str(selected.get("name") or ""),
                str(selected.get("endpoint") or DEFAULT_OLLAMA_ENDPOINT),
                False,
                bool(selected.get("vision")),
                "",
            )
        key = self.api_key_var.get().strip() or os.environ.get("OPENAI_API_KEY", "").strip()
        if key:
            model = self.openai_model_by_label.get(self.openai_model_var.get(), "gpt-5.6-terra")
            return "openai", model, DEFAULT_OPENAI_COMPAT_ENDPOINT, False, True, key
        return "openai", "", DEFAULT_OPENAI_COMPAT_ENDPOINT, True, False, ""

    def _start(self) -> None:
        if self.running:
            return
        input_text = self.input_var.get().strip()
        output_text = self.output_var.get().strip()
        if not input_text or not Path(input_text).is_file():
            messagebox.showerror("Choose a file", "Choose the original file first.")
            return
        if not output_text:
            messagebox.showerror("Choose a destination", "Choose where the finished project should be saved.")
            return
        output_path = Path(output_text)
        if output_path.suffix.lower() != ".tex":
            output_path = output_path.with_suffix(".tex")
            self.output_var.set(str(output_path))

        extension = Path(input_text).suffix.lower()
        visual_input = extension == ".pdf" or extension in IMAGE_EXTENSIONS
        look = self.look_var.get()
        if look == LOOK_EXACT and not visual_input:
            messagebox.showerror(
                "Visual source required",
                "Keeping original pages unchanged is available for PDFs and images. Choose another result type for this file.",
            )
            return
        numbers = self._read_numbers()
        if numbers is None:
            return
        selected_ai = self._choose_ai(look)
        if selected_ai is None:
            return
        provider, model, endpoint, no_llm, model_is_visual, api_key = selected_ai
        if look == LOOK_CLOSE and not model_is_visual and not no_llm:
            messagebox.showerror(
                "Vision model required",
                "The selected model cannot inspect page images. Select Automatic or a vision model.",
            )
            return
        if output_path.exists() and not messagebox.askyesno(
            "Replace existing project?",
            f"{output_path.name} already exists. Replace it?",
        ):
            return

        style_guide = ""
        if self.style_guide_var.get().strip():
            try:
                style_guide = Path(self.style_guide_var.get()).read_text(
                    encoding="utf-8", errors="replace"
                )
            except OSError as exc:
                messagebox.showerror("Cannot read style guide", str(exc))
                return

        exact = look == LOOK_EXACT
        close = look == LOOK_CLOSE
        colour = self.colour_var.get() == COLOUR_KEEP
        use_ocr = extension == ".pdf" and not exact
        ocr_language = self.language_by_label.get(self.language_var.get(), "eng")
        page_size = PAGE_SIZE_LABELS.get(self.page_size_var.get(), PAGE_SIZE_A4)
        page_flow = PAGE_FLOW_LABELS.get(self.page_flow_var.get(), PAGE_FLOW_COMPACT)
        photo_handling = PHOTO_LABELS.get(self.photo_var.get(), PHOTO_KEEP)
        options: dict[str, Any] = {
            "input_path": Path(input_text),
            "output_path": output_path,
            "provider": provider,
            "model": model,
            "endpoint": endpoint,
            "api_key": api_key,
            "no_llm": no_llm,
            "strict_mode": True,
            "match_mode": MATCH_MODE_PERCENT,
            "no_review": not self.detailed_review_var.get(),
            "use_ocr": use_ocr,
            "ocr_force": self.ocr_force_var.get(),
            "ocr_lang": ocr_language,
            "document_language": ocr_language,
            "page_size": page_size,
            "page_flow": page_flow,
            "photo_handling": photo_handling,
            "preserve_graphs": False,
            "preserve_layout": close or exact or self.keep_line_breaks_var.get(),
            "preserve_color": True if exact else colour and close,
            "image_only": exact,
            "vision_mode": close and not no_llm,
            "redraw_graphs": close and self.redraw_graphs_var.get(),
            "style_guide": style_guide,
            "compile_pdf": True,
            "keep_page_files": self.keep_page_files_var.get(),
            "no_wrapper": False,
            "backoff": 1.5,
            **numbers,
        }
        self.cancel_event.clear()
        self.last_result = None
        self._clear_log()
        self._set_running(True)
        self.status_label.configure(text="Starting…")
        self.worker = threading.Thread(target=self._convert, args=(options,), daemon=True)
        self.worker.start()

    def _convert(self, options: dict[str, Any]) -> None:
        try:
            result = convert_book_to_latex(
                **options,
                progress_callback=lambda current, total, message: self.events.put(
                    ("progress", (current, total, message))
                ),
                log_callback=lambda message: self.events.put(("log", message)),
                cancel_callback=self.cancel_event.is_set,
            )
            self.events.put(("result", result))
        except ConversionCancelled:
            self.events.put(("cancelled", None))
        except Exception as exc:  # noqa: BLE001
            self.events.put(("log", traceback.format_exc()))
            self.events.put(("error", str(exc)))
        finally:
            self.events.put(("finished", None))

    def _cancel(self) -> None:
        if self.running:
            self.cancel_event.set()
            self.cancel_button.configure(state="disabled")
            self.status_label.configure(text="Cancelling after the current page…")

    def _set_running(self, running: bool) -> None:
        self.running = running
        self.start_button.configure(state="disabled" if running else "normal")
        self.cancel_button.configure(state="normal" if running else "disabled")
        if running:
            self.progress.configure(value=0, maximum=100)
            self.open_folder_button.configure(state="disabled")
            self.open_pdf_button.configure(state="disabled")
            self.open_review_button.configure(state="disabled")

    def _toggle_details(self) -> None:
        if self.details_window is not None and self.details_window.winfo_exists():
            self.details_window.lift()
            self.details_window.focus_force()
            return
        window = tk.Toplevel(self.root)
        window.title("Conversion details")
        window.geometry("820x500")
        window.minsize(560, 300)
        frame = ttk.Frame(window, padding=10)
        frame.pack(fill="both", expand=True)
        text_widget = tk.Text(frame, wrap="word", state="normal", font=("Consolas", 9))
        scroll = ttk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scroll.set)
        text_widget.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        text_widget.insert("1.0", "\n".join(self.log_messages))
        text_widget.see("end")
        text_widget.configure(state="disabled")
        self.details_window = window
        self.details_text = text_widget
        self.details_button.configure(text="Details open")
        window.protocol("WM_DELETE_WINDOW", self._close_details)

    def _close_details(self) -> None:
        if self.details_window is not None:
            self.details_window.destroy()
        self.details_window = None
        self.details_text = None
        self.details_button.configure(text="Show details")

    def _clear_log(self) -> None:
        self.log_messages.clear()
        if self.details_text is not None and self.details_text.winfo_exists():
            self.details_text.configure(state="normal")
            self.details_text.delete("1.0", "end")
            self.details_text.configure(state="disabled")

    def _append_log(self, message: str) -> None:
        self.log_messages.append(message)
        if self.details_text is not None and self.details_text.winfo_exists():
            self.details_text.configure(state="normal")
            self.details_text.insert("end", f"{message}\n")
            self.details_text.see("end")
            self.details_text.configure(state="disabled")

    def _handle_result(self, result: dict[str, object]) -> None:
        self.last_result = result
        uncertain = int(result.get("uncertain_count", 0))
        pdf_path = result.get("pdf_path")
        compilation = result.get("compilation") or {}
        if pdf_path:
            if result.get("exact_visual_mode"):
                message = f"Finished — exact page copy created ({result['converted_pages']} source page(s))"
            elif result.get("page_flow") == PAGE_FLOW_COMPACT:
                message = f"Finished — compact editable LaTeX and PDF created ({result['converted_pages']} source unit(s))"
            else:
                message = f"Finished — editable LaTeX and PDF created with source page boundaries ({result['converted_pages']} unit(s))"
        else:
            message = "LaTeX was created, but the PDF needs attention"
        if uncertain:
            message += f"; {uncertain} item(s) highlighted in the review"
        self.status_label.configure(text=message)
        self.progress.configure(value=100, maximum=100)
        self.open_folder_button.configure(state="normal")
        self.open_pdf_button.configure(state="normal" if pdf_path else "disabled")
        self.open_review_button.configure(state="normal" if result.get("report_path") else "disabled")
        summary = message + "."
        if not pdf_path and compilation.get("message"):
            summary += f"\n\nPDF note: {compilation['message']}"
        if uncertain:
            summary += "\n\nOpen the review report to see what needs checking."
        messagebox.showinfo("Your files are ready", summary)

    def _drain_queue(self) -> None:
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "log":
                    self._append_log(str(value))
                elif kind == "progress":
                    current, total, message = value
                    self.progress.configure(
                        maximum=max(total, 1), value=min(current, max(total, 1))
                    )
                    self.status_label.configure(text=message)
                elif kind == "local_models":
                    self._handle_local_models(value)
                elif kind == "openai_connection":
                    self._handle_openai_connection(value)
                elif kind == "local_install":
                    self._handle_local_install(value)
                elif kind == "update_check":
                    self._handle_update_check(value)
                elif kind == "result":
                    self._handle_result(value)
                elif kind == "cancelled":
                    self.status_label.configure(text="Cancelled — the previous finished files were not replaced")
                elif kind == "error":
                    self.status_label.configure(text="Could not finish — select Show details for more information")
                    messagebox.showerror(
                        "Conversion could not finish",
                        f"{value}\n\nSelect Show details for technical information.",
                    )
                elif kind == "finished":
                    self._set_running(False)
        except queue.Empty:
            pass
        finally:
            if self.root.winfo_exists():
                self.root.after(100, self._drain_queue)

    def _open_output_folder(self) -> None:
        if self.last_result:
            open_in_system(Path(str(self.last_result["output_path"])).parent)

    def _open_pdf(self) -> None:
        if self.last_result and self.last_result.get("pdf_path"):
            open_in_system(Path(str(self.last_result["pdf_path"])))

    def _open_review(self) -> None:
        if not self.last_result or not self.last_result.get("report_path"):
            return
        report = Path(str(self.last_result["report_path"]))
        if report.is_file():
            open_in_system(report)

    def _on_close(self) -> None:
        if self.running and not messagebox.askyesno(
            "Conversion is running", "Stop the conversion and close the app?"
        ):
            return
        self.cancel_event.set()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    BookToLatexGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
