#!/usr/bin/env python3
"""Simple browser interface for Book to LaTeX & PDF."""

from __future__ import annotations

import csv
import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

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
    check_for_updates,
    convert_book_to_latex,
    ollama_connection_info,
    runtime_capabilities,
)

LOOK_CLEAN = "Reconstruct and polish (fully editable)"
LOOK_CLOSE = "Enhance a scan or lecture"
LOOK_EXACT = "Keep original pages unchanged"
COLOUR_KEEP = "Keep the original colours"
COLOUR_MONO = "Black and white"
AI_AUTO = "Automatic (recommended)"
AI_OLLAMA = "Choose an installed Ollama model"
AI_OPENAI = "Use an OpenAI-compatible service"
AI_NONE = "Do not use AI"

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

def _read_csv_rows(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8", newline="") as stream:
        return list(csv.DictReader(stream))


def _add_to_zip(zf: zipfile.ZipFile, path: Path) -> None:
    if path.is_file():
        zf.write(path, path.name)
    elif path.is_dir():
        for child in sorted(path.rglob("*")):
            if child.is_file():
                zf.write(child, f"{path.name}/{child.relative_to(path).as_posix()}")


def _zip_finished_project(result: dict[str, object]) -> bytes:
    buffer = io.BytesIO()
    seen: set[Path] = set()
    compilation = result.get("compilation") or {}
    review = result.get("review") or {}
    candidates = [
        result.get("output_path"),
        result.get("pdf_path"),
        result.get("report_path"),
        result.get("input_path") if result.get("exact_visual_mode") else None,
        result.get("images_dir"),
        result.get("assets_dir"),
        result.get("page_files_dir"),
        compilation.get("log_path"),
        review.get("review_dir"),
    ]
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for raw_path in candidates:
            if not raw_path:
                continue
            path = Path(str(raw_path))
            if path in seen or not path.exists():
                continue
            seen.add(path)
            _add_to_zip(archive, path)
    return buffer.getvalue()


@st.cache_data(ttl=5, show_spinner=False)
def _local_ai_info() -> dict[str, object]:
    return ollama_connection_info(timeout=4.0)


def _looks_visual(model_name: str) -> bool:
    lowered = model_name.lower()
    return any(
        token in lowered
        for token in ("vision", "-vl", "qwen3.5", "qwen35", "llava", "gpt-4o", "gpt-4.1")
    )


def _find_text_model(models: list[str]) -> str | None:
    for preferred in (
        "book-latex-qwen3-local-uncensored:8b",
        "book-latex-qwen3:8b",
        "qwen3:8b",
    ):
        if preferred in models:
            return preferred
    return next((model for model in models if not _looks_visual(model)), None)


def _find_vision_model(models: list[str]) -> str | None:
    for preferred in ("book-latex-qwen35-vision:9b", "qwen3.5:9b"):
        if preferred in models:
            return preferred
    return next((model for model in models if _looks_visual(model)), None)


st.set_page_config(page_title="Book to LaTeX & PDF", page_icon="📘", layout="wide")
st.title("📘 Book → LaTeX & PDF")
st.caption("Choose how the finished result should look. OCR, AI, checking and PDF compilation happen automatically.")
if st.button(f"Check for updates (v{APP_VERSION})"):
    update_info = check_for_updates(timeout=10)
    if not update_info.get("success"):
        st.error(f"Could not check GitHub: {update_info.get('error')}")
    elif update_info.get("update_available"):
        st.warning(f"Version {update_info['latest_version']} is available.")
        st.link_button("Open official download page", str(update_info["release_url"]))
    else:
        st.success(f"Version {APP_VERSION} is up to date.")

with st.expander("How it works"):
    st.markdown(
        """
1. Choose the original file.
2. Choose **Reconstruct and polish**, **Enhance a scan or lecture**, or **Keep original pages unchanged**.
3. Choose page size and compact or source-page flow.
4. Choose whether real photographs stay visible or become written descriptions.
5. Select **Create LaTeX and PDF**.

The result includes the LaTeX file, compiled PDF and a quality-check report. Technical controls are hidden in Advanced settings.
"""
    )

uploaded = st.file_uploader(
    "1. Choose the original file",
    help="PDF, document, ebook, spreadsheet, presentation, image or readable text/code file.",
)
if not uploaded:
    st.info("Choose a file to begin.")
    st.stop()

safe_upload_name = Path(uploaded.name).name
input_extension = Path(safe_upload_name).suffix.lower()
input_is_visual = input_extension == ".pdf" or input_extension in IMAGE_EXTENSIONS
capabilities = runtime_capabilities()
ai_info = _local_ai_info()
local_models = [str(model["name"]) for model in (ai_info.get("models") or [])]
text_model = _find_text_model(local_models)
vision_model = _find_vision_model(local_models)

if "book-latex-qwen3-local-uncensored:8b" in local_models:
    st.success("UNCENSORED local Qwen3 model ready · Vision AI, OCR and PDF compilation ready")
elif local_models:
    st.success("Local AI, OCR and PDF compilation are ready")
else:
    st.warning("No local AI model was detected. Clean basic conversion remains available.")

st.subheader("2. How should the result look?")
look = st.radio(
    "Appearance",
    [LOOK_CLEAN, LOOK_CLOSE, LOOK_EXACT],
    label_visibility="collapsed",
)
look_descriptions = {
    LOOK_CLEAN: "Rebuilds every word as polished editable LaTeX. The app improves page design, tables and spacing rather than copying them.",
    LOOK_CLOSE: "Enhances scans and lecture notes while keeping their recognizable structure, mathematics, visuals and meaningful page elements.",
    LOOK_EXACT: "Keeps each original page visually unchanged. Existing searchable text remains searchable, but the LaTeX text is not reconstructed.",
}
st.info(look_descriptions[look])

if look == LOOK_EXACT and not input_is_visual:
    st.error("Keeping original pages unchanged is available for PDFs and images. Choose another result type.")

st.subheader("3. Page, language and pictures")
left, right = st.columns(2)
with left:
    colour_choice = st.radio(
        "Colour",
        [COLOUR_KEEP] if look == LOOK_EXACT else [COLOUR_KEEP, COLOUR_MONO],
        horizontal=True,
        help="Enhanced mode reproduces visual colors. Original-pages mode always keeps source colors.",
    )
    page_size_label = st.selectbox(
        "Finished page size",
        list(PAGE_SIZE_LABELS),
        index=0 if look == LOOK_EXACT else 1,
        help="Same size preserves the source PDF shape. A4 or US Letter places the result on that paper size.",
    )
with right:
    language_labels = LANGUAGE_LABELS
    language_label = st.selectbox(
        "Document language",
        list(language_labels),
        help="Automatic detection is recommended. Missing official Arabic or Chinese Tesseract data is downloaded once and reused. The app preserves the source language and does not translate it.",
    )
    if look == LOOK_EXACT:
        page_flow_label = "Keep every source page separate"
        st.selectbox("Use of pages", [page_flow_label], disabled=True)
    else:
        page_flow_label = st.selectbox(
            "Use of pages",
            list(PAGE_FLOW_LABELS),
            help="Compact uses fewer pages. The other choice keeps every source page boundary.",
        )

photo_label = st.selectbox(
    "Real photographs",
    list(PHOTO_LABELS),
    disabled=look == LOOK_EXACT,
    help="Keep photographs as required project assets or replace them with descriptions. Graphs, tables, equations and technical diagrams are recreated as editable LaTeX either way.",
)
st.caption("Graphs, tables, equations and technical diagrams are recreated as editable LaTeX automatically in enhanced mode.")

default_output_dir = str(Path.cwd())
default_suffix = {
    LOOK_CLEAN: "_latex.tex",
    LOOK_CLOSE: "_close_layout.tex",
    LOOK_EXACT: "_exact_copy.tex",
}[look]
default_output_name = f"{Path(safe_upload_name).stem}{default_suffix}"

with st.expander("Advanced settings — optional"):
    st.caption("Most people should leave these settings unchanged.")
    output_dir = st.text_input(
        "Save folder on this computer",
        value=default_output_dir,
        help="This path is on the computer running Streamlit.",
    )
    output_name = st.text_input("LaTeX filename", value=default_output_name)
    ai_choice = st.selectbox(
        "AI choice",
        [AI_AUTO, AI_OLLAMA, AI_OPENAI, AI_NONE],
        help="Automatic picks the text or vision model based on the selected appearance.",
    )
    selected_model = ""
    endpoint = DEFAULT_OLLAMA_ENDPOINT
    api_key = ""
    if ai_choice == AI_OLLAMA:
        selected_model = st.selectbox("Installed Ollama model", local_models) if local_models else st.text_input("Ollama model")
    elif ai_choice == AI_OPENAI:
        selected_model = st.text_input("Online model name")
        endpoint = st.text_input("Connection address", value=DEFAULT_OPENAI_COMPAT_ENDPOINT)
        api_key = st.text_input("API key", type="password")
    if st.button("Refresh installed local models"):
        _local_ai_info.clear()
        st.rerun()

    start_page = st.number_input("First page", min_value=1, value=1)
    end_page = st.number_input("Last page (0 = all)", min_value=0, value=0)
    lines_per_page = st.number_input("Lines per content unit", min_value=1, value=70)
    ocr_dpi = st.slider("OCR/page-image quality", 150, 450, 220, 10)
    ocr_force = st.checkbox("Force OCR on every PDF page", value=False)
    keep_line_breaks = st.checkbox("Keep original line breaks (useful for poetry)", value=False)
    keep_page_files = st.checkbox("Save separate files for every page", value=False)
    detailed_review = st.checkbox(
        "Create a detailed per-page review",
        value=False,
        help="Normally the app creates one short report and removes temporary page images.",
    )
    allowed_difference = st.slider("Allowed text difference (%)", 0.0, 50.0, 10.0, 0.5)
    max_tokens = st.number_input("Maximum AI output per page", 200, 16000, 4096, 100)
    temperature = st.slider("AI creativity", 0.0, 2.0, 0.0, 0.05)
    timeout = st.number_input("AI timeout (seconds)", 10, 1200, 180)
    retries = st.number_input("AI retry count", 0, 8, 2)
    redraw_graphs = st.checkbox(
        "Recreate graphs, tables and technical diagrams as editable LaTeX",
        value=True,
    )
    style_guide_upload = st.file_uploader("Project style guide", type=["md", "txt"])

if "output_dir" not in locals():
    output_dir = default_output_dir
    output_name = default_output_name
    ai_choice = AI_AUTO
    selected_model = ""
    endpoint = DEFAULT_OLLAMA_ENDPOINT
    api_key = ""
    start_page = 1
    end_page = 0
    lines_per_page = 70
    ocr_dpi = 220
    ocr_force = False
    keep_line_breaks = False
    keep_page_files = False
    detailed_review = False
    allowed_difference = 10.0
    max_tokens = 4096
    temperature = 0.0
    timeout = 180
    retries = 2
    redraw_graphs = True
    style_guide_upload = None

start = st.button("Create LaTeX and PDF", type="primary", use_container_width=True)
if not start:
    st.stop()
if look == LOOK_EXACT and not input_is_visual:
    st.stop()

if look == LOOK_EXACT or ai_choice == AI_NONE:
    provider, model, endpoint, no_llm, vision_mode = "openai", "", DEFAULT_OPENAI_COMPAT_ENDPOINT, True, False
elif ai_choice == AI_OPENAI:
    provider, model, no_llm = "openai", selected_model.strip(), False
    vision_mode = look == LOOK_CLOSE and _looks_visual(model)
elif ai_choice == AI_OLLAMA:
    provider, model, endpoint, no_llm = "ollama", selected_model.strip(), DEFAULT_OLLAMA_ENDPOINT, False
    vision_mode = look == LOOK_CLOSE and _looks_visual(model)
elif look == LOOK_CLOSE:
    if not vision_model:
        st.error("Enhancing a scan or lecture needs the local vision model. Run setup_local_model.bat first.")
        st.stop()
    provider, model, endpoint, no_llm, vision_mode = "ollama", vision_model, DEFAULT_OLLAMA_ENDPOINT, False, True
else:
    provider, model, endpoint = "ollama", text_model or "", DEFAULT_OLLAMA_ENDPOINT
    no_llm, vision_mode = not bool(text_model), False

if not no_llm and not model:
    st.error("The selected AI choice needs a model name.")
    st.stop()
if look == LOOK_CLOSE and not vision_mode and not no_llm:
    st.error("Enhancing a scan or lecture requires a vision-capable model.")
    st.stop()

clean_output_name = Path(output_name.strip() or default_output_name).name
if not clean_output_name.lower().endswith(".tex"):
    clean_output_name += ".tex"
output_directory = Path(output_dir).expanduser().resolve()
output_directory.mkdir(parents=True, exist_ok=True)
output_path = output_directory / clean_output_name
style_guide = (
    style_guide_upload.getvalue().decode("utf-8", errors="replace")
    if style_guide_upload is not None
    else ""
)

status = st.empty()
progress = st.progress(0.0)
log_box = st.empty()
log_lines: list[str] = []


def on_log(message: str) -> None:
    log_lines.append(str(message))


def on_progress(current: int, total: int, message: str) -> None:
    progress.progress(0.0 if total <= 0 else min(current / total, 1.0))
    status.info(message)


with tempfile.TemporaryDirectory() as temporary:
    exact = look == LOOK_EXACT
    input_path = (
        output_directory / f"{output_path.stem}_required_source{input_extension}"
        if exact
        else Path(temporary) / safe_upload_name
    )
    input_path.write_bytes(uploaded.getvalue())
    try:
        close = look == LOOK_CLOSE
        result = convert_book_to_latex(
            input_path=input_path,
            output_path=output_path,
            provider=provider,
            model=model,
            endpoint=endpoint,
            api_key=api_key,
            no_llm=no_llm,
            strict_mode=True,
            start_page=int(start_page),
            end_page=int(end_page),
            lines_per_page=int(lines_per_page),
            max_tokens=int(max_tokens),
            temperature=float(temperature),
            timeout=float(timeout),
            retries=int(retries),
            backoff=1.5,
            keep_page_files=keep_page_files,
            no_wrapper=False,
            match_mode=MATCH_MODE_PERCENT,
            match_error_percent=float(allowed_difference),
            no_review=not detailed_review,
            use_ocr=input_extension == ".pdf" and not exact,
            ocr_force=ocr_force,
            ocr_dpi=int(ocr_dpi),
            ocr_lang=language_labels[language_label],
            document_language=language_labels[language_label],
            page_size=PAGE_SIZE_LABELS[page_size_label],
            page_flow=PAGE_FLOW_LABELS[page_flow_label],
            photo_handling=PHOTO_LABELS[photo_label],
            preserve_graphs=False,
            preserve_layout=close or exact or keep_line_breaks,
            preserve_color=True if exact else (colour_choice == COLOUR_KEEP) and close,
            image_only=exact,
            vision_mode=vision_mode,
            redraw_graphs=close and redraw_graphs,
            style_guide=style_guide,
            compile_pdf=True,
            progress_callback=on_progress,
            log_callback=on_log,
        )
    except Exception as exc:  # noqa: BLE001
        status.error("Conversion could not finish")
        st.error(str(exc))
        with st.expander("Technical details"):
            st.code("\n".join(log_lines[-400:]), language="text")
        st.stop()

progress.progress(1.0)
status.success("Your LaTeX project and PDF are ready")
left, right = st.columns(2)
with left:
    st.success(f"LaTeX: `{result['output_path']}`")
    if result.get("pdf_path"):
        st.success(f"PDF: `{result['pdf_path']}`")
    elif result.get("compilation"):
        st.warning(str(result["compilation"].get("message")))
    if result.get("uncertain_count"):
        st.warning(f"{result['uncertain_count']} item(s) are highlighted in the review report.")
with right:
    st.download_button(
        "Download complete project",
        data=_zip_finished_project(result),
        file_name=f"{output_path.stem}_complete.zip",
        mime="application/zip",
        use_container_width=True,
    )
    if result.get("pdf_path"):
        pdf_path = Path(str(result["pdf_path"]))
        st.download_button(
            "Download finished PDF",
            data=pdf_path.read_bytes(),
            file_name=pdf_path.name,
            mime="application/pdf",
            use_container_width=True,
        )

review = result.get("review")
short_report = Path(str(result["report_path"]))
st.download_button(
    "Download conversion report",
    data=short_report.read_bytes(),
    file_name=short_report.name,
    mime="text/plain",
)
if review:
    report_path = Path(str(review["report"]))
    st.download_button(
        "Download quality-check report",
        data=report_path.read_bytes(),
        file_name=report_path.name,
        mime="text/html",
    )
    uncertain_rows = _read_csv_rows(Path(str(review["uncertain_csv"])))
    if uncertain_rows:
        with st.expander("Items that need checking"):
            st.dataframe(uncertain_rows, use_container_width=True, hide_index=True)

with st.expander("Technical details"):
    st.code("\n".join(log_lines[-400:]), language="text")
