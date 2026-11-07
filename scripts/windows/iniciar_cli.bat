@echo off
chcp 65001 >nul
title Renomeador de PDFs v6.0 (Modo CLI)
cd /d "%~dp0..\.."
python main.py --cli

