@echo off
chcp 65001 >nul
title Renomeador de PDFs v4.0
cd /d "%~dp0"
python renomear.py
