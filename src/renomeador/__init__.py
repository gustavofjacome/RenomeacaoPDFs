# -*- coding: utf-8 -*-
"""
Pacote Renomeador de PDFs
-------------------------
Módulos:
- core: motor de leitura de planilhas, 3 prioridades, OCR local e operações PDF
- gui: backend do servidor web local (127.0.0.1) e rotas da API
- desfazer: utilitário de reversão de logs
"""

from renomeador import core
from renomeador import gui
from renomeador import desfazer

__all__ = ["core", "gui", "desfazer"]
