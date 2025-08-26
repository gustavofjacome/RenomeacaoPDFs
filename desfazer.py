# -*- coding: utf-8 -*-
import os
import re
import argparse
from pathlib import Path

ARQUIVO_LOG = "renomeacao.log"
PASTA_PDFS_PADRAO = "pdfs"

def parse_log_line(line):
    """Extrai nomes de uma linha de log."""
    match = re.search(r"SUCESSO: '(.*?)' -> '(.*?)'", line)
    return match.groups() if match else (None, None)

def main(args):
    pdfs_dir = Path(args.pasta_pdfs)
    log_file = Path(ARQUIVO_LOG)
    dry_run = args.dry_run

    os.system('cls' if os.name == 'nt' else 'clear')
    print(f"[INÍCIO] Reversão - {'SIMULAÇÃO' if dry_run else 'REAL'}")
    print("=" * 50)

    if not log_file.exists():
        print("[ERRO] Log não encontrado."); return
    if not pdfs_dir.exists():
        print("[ERRO] Pasta PDF não encontrada."); return

    renomeacoes = []
    for linha in log_file.read_text(encoding='utf-8').splitlines():
        a, b = parse_log_line(linha)
        if a and b: renomeacoes.append((a, b))

    if not renomeacoes:
        print("[INFO] Nenhuma renomeação encontrada."); return

    if not dry_run:
        print(f"[AVISO] {len(renomeacoes)} arquivos serão revertidos.")
        if input("Confirmar (s/n)? ").lower() != 's': return

    revertidos = erros = 0
    for antigo, novo in reversed(renomeacoes):
        caminho_novo = pdfs_dir / novo
        caminho_antigo = pdfs_dir / antigo
        print(f"Revertendo {novo} -> {antigo}")
        if not caminho_novo.exists():
            print("  [IGNORADO] Arquivo não existe."); continue
        if caminho_antigo.exists():
            print("  [ERRO] Nome original já existe."); erros += 1; continue
        try:
            if not dry_run: caminho_novo.rename(caminho_antigo)
            print("  [OK]" if not dry_run else "  [SIMULADO]")
            revertidos += 1
        except Exception as e:
            print(f"  [ERRO] {e}"); erros += 1

    print("="*50)
    print(f"Revertidos: {revertidos}\nErros: {erros}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--pasta_pdfs', default=PASTA_PDFS_PADRAO)
    parser.add_argument('--dry-run', action='store_true')
    main(parser.parse_args())
    input("Pressione Enter para sair...")
