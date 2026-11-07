# -*- coding: utf-8 -*-
"""
Ponto de entrada do Renomeador de PDFs
--------------------------------------
Permite executar:
    python main.py          (Inicia a Interface Gráfica Web Local v6.5)
    python main.py --cli    (Inicia no Modo Terminal / CLI)
    python main.py --undo   (Inicia a ferramenta de reversão)
"""

import sys
from pathlib import Path

# Adiciona a pasta src/ ao sys.path para importação dos módulos
raiz = Path(__file__).resolve().parent
src_dir = raiz / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from renomeador import gui, core, desfazer


def main():
    if "--cli" in sys.argv:
        sys.argv.remove("--cli")
        import argparse
        parser = argparse.ArgumentParser(description="Renomeador de PDFs - CLI")
        parser.add_argument("--planilha", default=core.NOME_PLANILHA_PADRAO)
        parser.add_argument("--pasta_pdfs", default=core.PASTA_PDFS_PADRAO)
        parser.add_argument("--auto-iniciar", action="store_true")
        parser.add_argument("--nao-abrir-pdf", action="store_true")
        parser.add_argument("--sem-pausa", action="store_true")
        core.main(parser.parse_args())
    elif "--undo" in sys.argv:
        sys.argv.remove("--undo")
        import argparse
        parser = argparse.ArgumentParser(description="Desfazer Renomeações")
        parser.add_argument("--pasta_pdfs", default=core.PASTA_PDFS_PADRAO)
        parser.add_argument("--dry-run", action="store_true")
        desfazer.main(parser.parse_args())
    else:
        gui.main()


if __name__ == "__main__":
    main()
