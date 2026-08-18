@echo off
setlocal

cd /d "%~dp0"
streamlit run book_to_latex_streamlit.py
