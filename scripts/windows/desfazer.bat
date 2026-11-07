@echo off
chcp 65001 >nul
title Desfazer Renomeacoes
cd /d "%~dp0..\.."
python main.py --undo

