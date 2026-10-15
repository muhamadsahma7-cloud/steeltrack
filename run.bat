@echo off
cd /d "%~dp0streamlit_app"
python -m streamlit run app.py
pause
