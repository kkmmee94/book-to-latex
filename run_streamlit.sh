#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
streamlit run book_to_latex_streamlit.py
