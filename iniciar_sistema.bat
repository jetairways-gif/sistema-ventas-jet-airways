@echo off
title Sistema de Ventas por Cursos
cd /d "%~dp0"
python -m streamlit run app.py
pause
