@echo off
chcp 65001 >nul
title Renomeador de PDFs v5.0 (Interface Grafica Local)
cd /d "%~dp0"
python gui.py
