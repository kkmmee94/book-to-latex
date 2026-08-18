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
from collections.abc import Iterable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from book_to_latex import (
    DEFAULT_OLLAMA_ENDPOINT,
    DEFAULT_OPENAI_COMPAT_ENDPOINT,
    DOCUMENT_LANGUAGES,
    IMAGE_EXTENSIONS,
    MATCH_MODE_PERCENT,
    SUPPORTED_INPUT_EXTENSIONS,
    ConversionCancelled,
    convert_book_to_latex,
    ollama_connection_info,
    runtime_capabilities,
)

LOOK_CLEAN = "Clean and editable (recommended)"
LOOK_CLOSE = "Stay close to the original layout"
LOOK_EXACT = "Exact visual copy"

COLOUR_KEEP = "Keep the original colours"
COLOUR_MONO = "Black and white"

AI_AUTO = "Automatic (recommended)"
AI_OLLAMA = "Choose an installed Ollama model"
AI_OPENAI = "Use an OpenAI-compatible service"
AI_NONE = "Do not use AI"

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
        self.root.geometry("880x720")
        self.root.minsize(760, 650)
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
        self.local_model_names: list[str] = []

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.look_var = tk.StringVar(value=LOOK_CLEAN)
        self.colour_var = tk.StringVar(value=COLOUR_KEEP)

        self.language_by_label = {label: code for code, label in DOCUMENT_LANGUAGES.items()}
        default_language = "English" if "English" in self.language_by_label else next(
            iter(self.language_by_label)
        )
        self.language_var = tk.StringVar(value=default_language)

        # Expert settings are intentionally hidden from the main workflow.
        self.ai_choice_var = tk.StringVar(value=AI_AUTO)
        self.advanced_model_var = tk.StringVar()
        self.endpoint_var = tk.StringVar(value=DEFAULT_OLLAMA_ENDPOINT)
        self.api_key_var = tk.StringVar()
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
        self.redraw_graphs_var = tk.BooleanVar(value=False)
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
        ttk.Button(heading, text="Advanced settings", command=self._open_advanced).pack(side="right")

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
                "Best for novels, poetry, essays and documents you want to edit. Uses clean readable formatting rather than copying every page position.",
            ),
            (
                LOOK_CLOSE,
                "Best for mathematics, tables and structured pages. A vision model studies each page and recreates the layout as editable LaTeX.",
            ),
            (
                LOOK_EXACT,
                "Looks identical because each original page is placed directly into LaTeX. The page text is not editable, and temporary page pictures are not kept.",
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
            main, text="3. Language and colour", style="Section.TLabelframe"
        )
        preferences.pack(fill="x", pady=6)
        ttk.Label(preferences, text="Colour").grid(row=0, column=0, sticky="w", padx=12, pady=10)
        colour_keep = ttk.Radiobutton(
            preferences,
            text=COLOUR_KEEP,
            value=COLOUR_KEEP,
            variable=self.colour_var,
        )
        colour_keep.grid(row=0, column=1, sticky="w", padx=8, pady=10)
        colour_mono = ttk.Radiobutton(
            preferences,
            text=COLOUR_MONO,
            value=COLOUR_MONO,
            variable=self.colour_var,
        )
        colour_mono.grid(row=0, column=2, sticky="w", padx=8, pady=10)
        add_tooltip(
            [colour_keep, colour_mono],
            "This controls preserved page pictures and visual reconstruction. Ordinary LaTeX text remains normal black text.",
        )
        ttk.Label(preferences, text="Document language").grid(
            row=0, column=3, sticky="e", padx=(24, 5), pady=10
        )
        language = ttk.Combobox(
            preferences,
            textvariable=self.language_var,
            values=list(self.language_by_label),
            state="readonly",
            width=20,
        )
        language.grid(row=0, column=4, sticky="w", padx=(5, 12), pady=10)
        add_tooltip(
            language,
            "Choose the language already used in the document. Arabic enables right-to-left LaTeX, Arabic OCR, and XeLaTeX automatically. The app does not translate unless you ask it to.",
        )

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

        ttk.Label(
            main,
            text="You will receive: an editable LaTeX file, a compiled PDF, and a quality-check report.",
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
            LOOK_CLEAN: "The app will use the fast local text model when helpful, automatically OCR unreadable pages, create editable LaTeX, compile it, and verify the words and numbers.",
            LOOK_CLOSE: "The app will automatically choose the installed vision model, inspect both the page image and extracted text, reconstruct mathematics and structure, compile the result, and flag anything uncertain.",
            LOOK_EXACT: "The app will preserve each PDF/image page exactly inside a LaTeX document, apply your colour choice, compile it, and skip unnecessary AI/OCR work.",
        }
        self.explanation_label.configure(text=messages[self.look_var.get()])

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
            self.output_var.set(str(source.with_name(f"{source.stem}_latex.tex")))
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
            "2. Choose Clean and editable, Stay close to the original, or Exact visual copy.\n\n"
            "3. Choose colour or black and white.\n\n"
            "4. Select Create LaTeX and PDF.\n\n"
            "The app automatically chooses the correct local AI model, checks whether OCR is needed, analyses PDF quality, compiles the LaTeX, and creates a review report. You do not need to understand those technical steps.",
        )

    def _open_advanced(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Advanced settings")
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
        ttk.Label(ai_tab, text="AI choice").grid(row=0, column=0, sticky="w", pady=7)
        ai_choice = ttk.Combobox(
            ai_tab,
            textvariable=self.ai_choice_var,
            values=[AI_AUTO, AI_OLLAMA, AI_OPENAI, AI_NONE],
            state="readonly",
        )
        ai_choice.grid(row=0, column=1, sticky="ew", pady=7)
        ai_choice.bind("<<ComboboxSelected>>", lambda _event: self._update_advanced_ai_state())
        ttk.Label(ai_tab, text="Model").grid(row=1, column=0, sticky="w", pady=7)
        self.advanced_model_combo = ttk.Combobox(
            ai_tab,
            textvariable=self.advanced_model_var,
            values=self.local_model_names,
        )
        self.advanced_model_combo.grid(row=1, column=1, sticky="ew", pady=7)
        ttk.Button(ai_tab, text="Refresh installed models", command=self._refresh_local_models).grid(
            row=2, column=1, sticky="w", pady=(0, 10)
        )
        self._advanced_field(ai_tab, 3, "Connection address", self.endpoint_var, "Automatic uses Ollama's local address.")
        self._advanced_field(ai_tab, 4, "API key", self.api_key_var, "Only needed by online services.", secret=True)
        self._advanced_field(ai_tab, 5, "Maximum AI output", self.max_tokens_var, "4096 suits dense pages.")
        self._advanced_field(ai_tab, 6, "AI creativity", self.temperature_var, "0 preserves source fidelity.")
        self._advanced_field(ai_tab, 7, "Timeout (seconds)", self.timeout_var, "Maximum wait per page.")
        self._advanced_field(ai_tab, 8, "Retry count", self.retries_var, "Extra attempts after a connection problem.")
        ttk.Label(
            ai_tab,
            text="Automatic uses the text model for clean documents and the vision model for close-layout PDFs/images.",
            style="Hint.TLabel",
            wraplength=610,
        ).grid(row=9, column=0, columnspan=2, sticky="w", pady=12)

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
            text="Ask the vision model to redraw graphs with TikZ/pgfplots",
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
        if not hasattr(self, "advanced_model_combo"):
            return
        choice = self.ai_choice_var.get()
        if choice == AI_OLLAMA:
            self.advanced_model_combo.configure(state="readonly" if self.local_model_names else "normal")
            self.endpoint_var.set(DEFAULT_OLLAMA_ENDPOINT)
        elif choice == AI_OPENAI:
            self.advanced_model_combo.configure(state="normal")
            if self.endpoint_var.get() == DEFAULT_OLLAMA_ENDPOINT:
                self.endpoint_var.set(DEFAULT_OPENAI_COMPAT_ENDPOINT)
        else:
            self.advanced_model_combo.configure(state="disabled")

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
            self.events.put(("ollama_models", ollama_connection_info(timeout=4.0)))

        threading.Thread(target=load, daemon=True).start()

    def _handle_ollama_models(self, info: dict[str, object]) -> None:
        models = info.get("models") or []
        self.local_model_names = [str(model["name"]) for model in models]
        if hasattr(self, "advanced_model_combo"):
            self.advanced_model_combo.configure(values=self.local_model_names)
        text_model = self._find_text_model()
        if text_model and not self.advanced_model_var.get():
            self.advanced_model_var.set(text_model)
        if info.get("available") and self.local_model_names:
            has_vision = bool(self._find_vision_model())
            vision_note = "vision AI ready" if has_vision else "vision AI not installed"
            uncensored_note = (
                "UNCENSORED local model ready · "
                if "book-latex-qwen3-local-uncensored:8b" in self.local_model_names
                else ""
            )
            self.readiness_label.configure(
                text=f"Ready: {uncensored_note}OCR, PDF analysis and PDF compilation available ({vision_note})."
            )
        elif info.get("available"):
            self.readiness_label.configure(
                text="Ready without AI. OCR and PDF compilation are available."
            )
        else:
            self.readiness_label.configure(
                text="Ready without local AI. OCR and PDF compilation are available."
            )

    def _find_text_model(self) -> str | None:
        preferred = [
            "book-latex-qwen3-local-uncensored:8b",
            "book-latex-qwen3:8b",
            "qwen3:8b",
        ]
        for name in preferred:
            if name in self.local_model_names:
                return name
        return next((name for name in self.local_model_names if not self._looks_visual(name)), None)

    def _find_vision_model(self) -> str | None:
        preferred = ["book-latex-qwen35-vision:9b", "qwen3.5:9b"]
        for name in preferred:
            if name in self.local_model_names:
                return name
        return next((name for name in self.local_model_names if self._looks_visual(name)), None)

    @staticmethod
    def _looks_visual(model_name: str) -> bool:
        lowered = model_name.lower()
        return any(
            token in lowered
            for token in ("vision", "-vl", "qwen3.5", "qwen35", "llava", "gpt-4o", "gpt-4.1")
        )

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
                "One of the numeric Advanced settings contains invalid text.",
            )
            return None

    def _choose_ai(self, look: str) -> tuple[str, str, str, bool, bool] | None:
        choice = self.ai_choice_var.get()
        if look == LOOK_EXACT:
            return "openai", "", DEFAULT_OPENAI_COMPAT_ENDPOINT, True, False
        if choice == AI_NONE:
            return "openai", "", DEFAULT_OPENAI_COMPAT_ENDPOINT, True, False
        if choice == AI_OPENAI:
            model = self.advanced_model_var.get().strip()
            if not model:
                messagebox.showerror("Missing AI model", "Enter the online model in Advanced settings.")
                return None
            return "openai", model, self.endpoint_var.get().strip(), False, False
        if choice == AI_OLLAMA:
            model = self.advanced_model_var.get().strip()
            if not model:
                messagebox.showerror("Missing local model", "Select an Ollama model in Advanced settings.")
                return None
            return "ollama", model, DEFAULT_OLLAMA_ENDPOINT, False, self._looks_visual(model)

        if look == LOOK_CLOSE:
            model = self._find_vision_model()
            if not model:
                messagebox.showerror(
                    "Vision model required",
                    "Close-layout conversion needs the local vision model. Run setup_local_model.bat, then reopen the app.",
                )
                return None
            return "ollama", model, DEFAULT_OLLAMA_ENDPOINT, False, True
        model = self._find_text_model()
        if model:
            return "ollama", model, DEFAULT_OLLAMA_ENDPOINT, False, False
        return "openai", "", DEFAULT_OPENAI_COMPAT_ENDPOINT, True, False

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
                "Exact visual copy is available for PDFs and images. Choose another appearance for this file.",
            )
            return
        numbers = self._read_numbers()
        if numbers is None:
            return
        selected_ai = self._choose_ai(look)
        if selected_ai is None:
            return
        provider, model, endpoint, no_llm, model_is_visual = selected_ai
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
        options: dict[str, Any] = {
            "input_path": Path(input_text),
            "output_path": output_path,
            "provider": provider,
            "model": model,
            "endpoint": endpoint,
            "api_key": self.api_key_var.get().strip(),
            "no_llm": no_llm,
            "strict_mode": True,
            "match_mode": MATCH_MODE_PERCENT,
            "no_review": not self.detailed_review_var.get(),
            "use_ocr": use_ocr,
            "ocr_force": self.ocr_force_var.get(),
            "ocr_lang": ocr_language,
            "document_language": ocr_language,
            "preserve_graphs": False,
            "preserve_layout": close or exact or self.keep_line_breaks_var.get(),
            "preserve_color": colour and (close or exact),
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
            message = f"Finished — LaTeX and PDF created ({result['converted_pages']} page/unit(s))"
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
                elif kind == "ollama_models":
                    self._handle_ollama_models(value)
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
