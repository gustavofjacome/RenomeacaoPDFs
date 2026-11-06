# -*- coding: utf-8 -*-
"""
Utilitário de Reversão em Lote (desfazer.py)
--------------------------------------------
Permite reverter as renomeações registradas em `renomeacao.log`.
Ignora automaticamente ações que já foram desfeitas pelo botão [Z] durante a sessão.
"""

import argparse
import os
import re
from datetime import datetime
from pathlib import Path
from typing import List, Tuple

ARQUIVO_LOG = "renomeacao.log"
PASTA_PDFS_PADRAO = "pdfs"


def extrair_renomeacoes_ativas(log_file: Path) -> List[Tuple[str, str]]:
    """
    Lê o arquivo de log e retorna apenas as renomeações que ainda estão ativas
    (descontando as que já foram revertidas com [Z]).
    """
    ativas: List[Tuple[str, str]] = []
    for linha in log_file.read_text(encoding="utf-8", errors="ignore").splitlines():
        m_suc = re.search(r"SUCESSO:\s*'(.*?)'\s*->\s*'(.*?)'", linha)
        if m_suc:
            ativas.append((m_suc.group(1), m_suc.group(2)))
            continue

        m_rev = re.search(r"REVERTIDO:\s*'(.*?)'\s*->\s*'(.*?)'", linha)
        if m_rev:
            atual, original = m_rev.group(1), m_rev.group(2)
            # Remover a última ocorrência correspondente
            for i in range(len(ativas) - 1, -1, -1):
                if ativas[i] == (original, atual):
                    ativas.pop(i)
                    break

    return ativas


def registrar_reversao_no_log(log_file: Path, atual: str, original: str):
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"[{agora}] REVERTIDO: '{atual}' -> '{original}'\n")


def main(args):
    script_dir = Path(__file__).resolve().parent
    pdfs_dir = Path(args.pasta_pdfs)
    if not pdfs_dir.is_absolute() and not pdfs_dir.exists():
        pdfs_dir = script_dir / args.pasta_pdfs

    log_file = script_dir / ARQUIVO_LOG
    dry_run = args.dry_run

    os.system("cls" if os.name == "nt" else "clear")
    print("=" * 60)
    print(f" ↩️  REVERSÃO DE RENOMEAÇÕES ({'SIMULAÇÃO' if dry_run else 'MODO REAL'})")
    print("=" * 60)

    if not log_file.exists():
        print("[ERRO] Arquivo 'renomeacao.log' não encontrado.")
        return
    if not pdfs_dir.exists():
        print(f"[ERRO] Pasta '{pdfs_dir}' não encontrada.")
        return

    renomeacoes = extrair_renomeacoes_ativas(log_file)
    if not renomeacoes:
        print("[INFO] Nenhuma renomeação ativa encontrada para reverter.")
        return

    if args.ultimos and args.ultimos > 0:
        renomeacoes_alvo = list(reversed(renomeacoes[-args.ultimos :]))
    else:
        renomeacoes_alvo = list(reversed(renomeacoes))

    print(f"Encontradas {len(renomeacoes)} renomeações ativas no histórico.")
    print(f"Serão revertidas: {len(renomeacoes_alvo)} arquivo(s).\n")

    for orig, novo in renomeacoes_alvo[:10]:
        print(f"  • {novo}  ->  {orig}")
    if len(renomeacoes_alvo) > 10:
        print(f"  ... e mais {len(renomeacoes_alvo) - 10} arquivo(s).")

    if not dry_run and not args.sim:
        resp = input("\nConfirmar reversão? (s/n): ").strip().lower()
        if resp != "s":
            print("Operação cancelada.")
            return

    revertidos = 0
    erros = 0
    print()
    for antigo, novo in renomeacoes_alvo:
        caminho_novo = pdfs_dir / novo
        caminho_antigo = pdfs_dir / antigo
        print(f"Revertendo: {novo} -> {antigo} ... ", end="")
        if not caminho_novo.exists():
            print("[IGNORADO - arquivo atual não encontrado]")
            continue
        if caminho_antigo.exists():
            print("[ERRO - arquivo original já existe]")
            erros += 1
            continue
        try:
            if not dry_run:
                caminho_novo.rename(caminho_antigo)
                registrar_reversao_no_log(log_file, novo, antigo)
            print("[OK]" if not dry_run else "[SIMULADO]")
            revertidos += 1
        except Exception as e:
            print(f"[ERRO: {e}]")
            erros += 1

    print("=" * 60)
    print(f"✅ Revertidos com sucesso: {revertidos}  |  ❌ Erros: {erros}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Reverte renomeações feitas pelo renomear.py")
    parser.add_argument("--pasta_pdfs", default=PASTA_PDFS_PADRAO, help="Pasta onde estão os PDFs")
    parser.add_argument("--dry-run", action="store_true", help="Apenas simula sem renomear nada")
    parser.add_argument("--ultimos", type=int, default=0, help="Reverter apenas os últimos N arquivos")
    parser.add_argument("--sim", action="store_true", help="Confirma automaticamente sem perguntar")
    main(parser.parse_args())
    if not any(arg in ("--sim", "--dry-run") for arg in os.sys.argv):
        try:
            input("\nPressione Enter para sair...")
        except (EOFError, KeyboardInterrupt):
            pass
