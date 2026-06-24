@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist __pycache__ rmdir /s /q __pycache__
streamlit run app.py
pause
