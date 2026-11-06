# -*- coding: utf-8 -*-
"""
Renomeador Semiautomático de PDFs - Interface Gráfica Web Local (v5.0 GUI)
--------------------------------------------------------------------------
- Roda 100% local em http://127.0.0.1 (nenhum dado sai do seu computador).
- Interface limpa, minimalista e rápida com ícones padrão Font Awesome.
- Visualizador de PDF embutido na própria página ao lado dos comandos e da fila.
- Suporte completo a atalhos de 1 tecla (S/Enter, E, B, D, N, P, Z).
"""

import argparse
import json
import re
import shutil
import socket
import threading
import time
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import renomear as core

PORTA_PADRAO = 8765
HOST_LOCAL = "127.0.0.1"


def limpar_texto_status(texto: str) -> str:
    """Remove códigos ANSI e emojis de mensagens vindas do módulo CLI."""
    sem_ansi = re.sub(r"\033\[[0-9;]*m", "", str(texto))
    # Remover emojis comuns mantendo acentuação latina
    sem_emoji = re.sub(r"[\U00010000-\U0010ffff]|[\u2600-\u27bf]|\ufe0f", "", sem_ansi)
    return " ".join(sem_emoji.split())


class EstadoGUI:
    def __init__(self, planilha_path: Path, pdfs_dir: Path, script_dir: Path):
        self.lock = threading.Lock()
        self.script_dir = script_dir
        self.planilha_path = planilha_path
        self.pdfs_dir = pdfs_dir
        self.log_path = script_dir / core.ARQUIVO_LOG
        self.backup_divisao_dir = pdfs_dir / core.PASTA_BACKUP_DIVISAO

        self.itens_planilha: List[core.ItemPlanilha] = []
        self.fila_pdfs: List[Path] = []
        self.pdfs_ja_renomeados: List[Path] = []
        self.historico: List[core.AcaoHistorico] = []

        self.idx_pdf: int = 0
        self.idx_nome: Optional[int] = None

        self.total_renomeados: int = 0
        self.total_editados: int = 0
        self.total_divididos: int = 0
        self.total_pulados: int = 0
        self.ultimo_evento: str = "Pronto"
        self.erro_carregamento: str = ""

        self._cache_paginas: Dict[str, Optional[int]] = {}
        self.recarregar_tudo()

    def recarregar_tudo(self):
        self.pdfs_dir.mkdir(parents=True, exist_ok=True)
        self.erro_carregamento = ""

        itens, msg = core.carregar_planilha(self.planilha_path)
        if not itens:
            self.itens_planilha = []
            self.erro_carregamento = msg
        else:
            self.itens_planilha = itens

        fila_pendentes, ja_ren, _ = core.sincronizar_pdfs_e_planilha(
            self.pdfs_dir, self.itens_planilha, self.log_path
        )
        self.fila_pdfs = fila_pendentes
        self.pdfs_ja_renomeados = ja_ren
        self.idx_pdf = 0
        self.idx_nome = core.proximo_indice_nome_pendente(self.itens_planilha, 0)
        self._cache_paginas.clear()

    def reescanear_pasta_pdfs(self):
        self.pdfs_dir.mkdir(parents=True, exist_ok=True)
        fila_pendentes, ja_ren, _ = core.sincronizar_pdfs_e_planilha(
            self.pdfs_dir, self.itens_planilha, self.log_path
        )
        self.fila_pdfs = fila_pendentes
        self.pdfs_ja_renomeados = ja_ren
        if self.idx_pdf >= len(self.fila_pdfs):
            self.idx_pdf = max(0, len(self.fila_pdfs) - 1) if self.fila_pdfs else 0
        if (
            self.idx_nome is None
            or self.idx_nome >= len(self.itens_planilha)
            or self.itens_planilha[self.idx_nome].concluido
        ):
            self.idx_nome = core.proximo_indice_nome_pendente(self.itens_planilha, 0)
        self._cache_paginas.clear()

    def pdf_atual(self) -> Optional[Path]:
        if 0 <= self.idx_pdf < len(self.fila_pdfs):
            return self.fila_pdfs[self.idx_pdf]
        return None

    def pdf_proximo(self) -> Optional[Path]:
        if 0 <= self.idx_pdf + 1 < len(self.fila_pdfs):
            return self.fila_pdfs[self.idx_pdf + 1]
        return None

    def item_atual(self) -> Optional[core.ItemPlanilha]:
        if (
            self.idx_nome is not None
            and 0 <= self.idx_nome < len(self.itens_planilha)
            and not self.itens_planilha[self.idx_nome].concluido
        ):
            return self.itens_planilha[self.idx_nome]
        self.idx_nome = core.proximo_indice_nome_pendente(self.itens_planilha, 0)
        if self.idx_nome is not None:
            return self.itens_planilha[self.idx_nome]
        return None

    def obter_paginas_pdf(self, pdf_path: Optional[Path]) -> Optional[int]:
        if pdf_path is None or not pdf_path.exists():
            return None
        chave = f"{pdf_path.name}:{pdf_path.stat().st_mtime_ns}"
        if chave not in self._cache_paginas:
            self._cache_paginas[chave] = core.contar_paginas_pdf(pdf_path)
        return self._cache_paginas[chave]

    def salvar_relatorio(self):
        restantes_nomes = sum(1 for it in self.itens_planilha if not it.concluido)
        try:
            with open(self.script_dir / core.ARQUIVO_RELATORIO, "w", encoding="utf-8") as f:
                f.write(f"RELATÓRIO DE RENOMEAÇÃO (GUI v5.0) - {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 60 + "\n")
                f.write(f"Renomeados nesta sessão : {self.total_renomeados}\n")
                f.write(f"Editados manualmente    : {self.total_editados}\n")
                f.write(f"PDFs divididos (2+ p.)  : {self.total_divididos}\n")
                f.write(f"PDFs pulados            : {self.total_pulados}\n")
                f.write(f"Nomes restantes planilha: {restantes_nomes}/{len(self.itens_planilha)}\n")
        except Exception:
            pass

    def serializar_estado(self) -> dict:
        pdf_cur = self.pdf_atual()
        pdf_next = self.pdf_proximo()
        item_cur = self.item_atual()
        paginas_cur = self.obter_paginas_pdf(pdf_cur)

        total_planilha = len(self.itens_planilha)
        concluidos_planilha = sum(1 for it in self.itens_planilha if it.concluido)

        return {
            "tem_pdf": pdf_cur is not None,
            "pdf_atual": {
                "nome": pdf_cur.name,
                "url": f"/api/pdf?name={urllib.parse.quote(pdf_cur.name)}&t={int(pdf_cur.stat().st_mtime)}",
                "paginas": paginas_cur,
                "posicao": self.idx_pdf + 1,
                "total_fila": len(self.fila_pdfs),
                "restantes": max(0, len(self.fila_pdfs) - self.idx_pdf),
            }
            if pdf_cur and pdf_cur.exists()
            else None,
            "pdf_proximo_url": (
                f"/api/pdf?name={urllib.parse.quote(pdf_next.name)}&t={int(pdf_next.stat().st_mtime)}"
                if pdf_next and pdf_next.exists()
                else None
            ),
            "item_atual": {
                "idx": self.idx_nome,
                "indice_original": item_cur.indice_original,
                "nome_completo": item_cur.nome_completo,
                "matricula": item_cur.matricula,
                "nome_formatado": item_cur.nome_formatado,
            }
            if item_cur
            else None,
            "progresso": {
                "total_planilha": total_planilha,
                "concluidos_planilha": concluidos_planilha,
                "pendentes_planilha": total_planilha - concluidos_planilha,
                "percentual": round((100.0 * concluidos_planilha / total_planilha), 1)
                if total_planilha > 0
                else 0.0,
                "renomeados_sessao": self.total_renomeados,
                "editados_sessao": self.total_editados,
                "divididos_sessao": self.total_divididos,
                "pulados_sessao": self.total_pulados,
                "pode_desfazer": len(self.historico) > 0,
                "qtd_historico": len(self.historico),
            },
            "ultimo_evento": self.ultimo_evento,
            "erro_carregamento": self.erro_carregamento,
            "planilha_nome": self.planilha_path.name,
            "pasta_pdfs": str(self.pdfs_dir),
            "fila_nomes": [
                {
                    "idx": idx,
                    "indice_original": it.indice_original,
                    "nome_completo": it.nome_completo,
                    "matricula": it.matricula,
                    "nome_formatado": it.nome_formatado,
                    "concluido": it.concluido,
                    "atual": idx == self.idx_nome and not it.concluido,
                }
                for idx, it in enumerate(self.itens_planilha)
            ],
        }


# ==============================================================================
# AÇÕES DE NEGÓCIO
# ==============================================================================

def acao_confirmar_atual(estado: EstadoGUI, idx_escolhido: Optional[int] = None) -> Tuple[bool, str]:
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None or not pdf_cur.exists():
        return False, "Nenhum arquivo PDF pendente selecionado."

    alvo_idx = idx_escolhido if idx_escolhido is not None else estado.idx_nome
    if alvo_idx is None or not (0 <= alvo_idx < len(estado.itens_planilha)):
        return False, "Não há nome selecionado na fila. Use Editar (E) para informar um nome."

    item = estado.itens_planilha[alvo_idx]
    destino, erro = core.executar_renomeacao_arquivo(
        pdf_cur, item.nome_formatado, estado.pdfs_dir, estado.log_path
    )
    if not destino:
        return False, f"Erro ao renomear: {erro}"

    item.concluido = True
    estado.historico.append(
        core.AcaoHistorico(
            tipo="renomear",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=estado.idx_nome if estado.idx_nome is not None else 0,
            pdf_origem=pdf_cur,
            pdf_destino=destino,
            idx_nome_concluido=alvo_idx,
            descricao=f"{pdf_cur.name} -> {destino.name}",
        )
    )
    estado.total_renomeados += 1
    estado.idx_pdf += 1
    if alvo_idx == estado.idx_nome:
        estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, alvo_idx + 1)
    else:
        if estado.idx_nome is None or estado.itens_planilha[estado.idx_nome].concluido:
            estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, 0)

    estado.ultimo_evento = f"Renomeado: {pdf_cur.name} -> {destino.name}"
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_editar_manual(
    estado: EstadoGUI,
    nome_digitado: str,
    matricula_digitada: str,
    avancar_planilha: bool,
) -> Tuple[bool, str]:
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None or not pdf_cur.exists():
        return False, "Nenhum arquivo PDF pendente selecionado."

    _, _, nome_formatado = core.formatar_nome_e_matricula(nome_digitado, matricula_digitada)
    if not nome_formatado:
        return False, "Informe um nome válido."

    destino, erro = core.executar_renomeacao_arquivo(
        pdf_cur, nome_formatado, estado.pdfs_dir, estado.log_path
    )
    if not destino:
        return False, f"Erro ao renomear: {erro}"

    idx_concluido = None
    if avancar_planilha and estado.idx_nome is not None and 0 <= estado.idx_nome < len(estado.itens_planilha):
        estado.itens_planilha[estado.idx_nome].concluido = True
        idx_concluido = estado.idx_nome

    estado.historico.append(
        core.AcaoHistorico(
            tipo="renomear",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=estado.idx_nome if estado.idx_nome is not None else 0,
            pdf_origem=pdf_cur,
            pdf_destino=destino,
            idx_nome_concluido=idx_concluido,
            descricao=f"Editado: {pdf_cur.name} -> {destino.name}",
        )
    )
    estado.total_renomeados += 1
    estado.total_editados += 1
    estado.idx_pdf += 1
    if avancar_planilha and estado.idx_nome is not None:
        estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, estado.idx_nome + 1)

    modo_txt = "fila avançou" if avancar_planilha else "fila mantida"
    estado.ultimo_evento = f"Renomeado ( manual ): {pdf_cur.name} -> {destino.name} ({modo_txt})"
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_pular_nome(estado: EstadoGUI) -> Tuple[bool, str]:
    item_cur = estado.item_atual()
    if item_cur is None or estado.idx_nome is None:
        return False, "Não há mais nomes pendentes na fila."

    nome_antigo = item_cur.nome_formatado
    prox = core.proximo_indice_nome_pendente(estado.itens_planilha, estado.idx_nome + 1)
    if prox == estado.idx_nome:
        return False, "Este é o único nome restante na fila."

    estado.historico.append(
        core.AcaoHistorico(
            tipo="pular_nome",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=estado.idx_nome,
            descricao=f"Pulou nome {nome_antigo}",
        )
    )
    estado.idx_nome = prox
    estado.ultimo_evento = f"Nome pulado: {nome_antigo}"
    return True, estado.ultimo_evento


def acao_pular_pdf(estado: EstadoGUI, avancar_nome_junto: bool = False) -> Tuple[bool, str]:
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None:
        return False, "Não há PDF atual para pular."

    estado.historico.append(
        core.AcaoHistorico(
            tipo="pular_pdf",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=estado.idx_nome if estado.idx_nome is not None else 0,
            descricao=f"Pulou PDF {pdf_cur.name}",
        )
    )
    estado.total_pulados += 1
    estado.idx_pdf += 1
    if avancar_nome_junto and estado.idx_nome is not None:
        estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, estado.idx_nome + 1)

    estado.ultimo_evento = f"PDF pulado: {pdf_cur.name}"
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_dividir_pdf(
    estado: EstadoGUI,
    modo: str,
    pagina_corte: int = 1,
    ja_renomear_parte1: bool = True,
) -> Tuple[bool, str]:
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None or not pdf_cur.exists():
        return False, "Nenhum arquivo PDF pendente selecionado."

    num_paginas = estado.obter_paginas_pdf(pdf_cur)
    estado.backup_divisao_dir.mkdir(parents=True, exist_ok=True)
    backup_path = core.obter_destino_unico(estado.backup_divisao_dir, pdf_cur.stem)

    novos_pdfs: List[Path] = []

    if modo == "duplicar" or (num_paginas == 1 and modo != "todas"):
        parte1 = core.obter_destino_unico(pdf_cur.parent, f"{pdf_cur.stem}_parte1")
        parte2 = core.obter_destino_unico(pdf_cur.parent, f"{pdf_cur.stem}_parte2")
        try:
            shutil.copy2(pdf_cur, backup_path)
            shutil.copy2(pdf_cur, parte1)
            shutil.move(str(pdf_cur), str(parte2))
            novos_pdfs = [parte1, parte2]
        except Exception as e:
            return False, f"Erro ao duplicar PDF: {e}"
    else:
        total_p = num_paginas if (num_paginas and num_paginas >= 2) else 2
        if modo == "todas":
            intervalos = [(p, p) for p in range(1, total_p + 1)]
        else:
            corte = max(1, min(int(pagina_corte), total_p - 1))
            intervalos = [(1, corte), (corte + 1, total_p)]

        novos_pdfs, status = core.dividir_pdf_em_intervalos(pdf_cur, intervalos)
        if not novos_pdfs:
            if "pip install pypdf" in status and core._instalar_pypdf_usuario():
                novos_pdfs, status = core.dividir_pdf_em_intervalos(pdf_cur, intervalos)
            if not novos_pdfs:
                return False, status

        try:
            shutil.move(str(pdf_cur), str(backup_path))
        except Exception:
            pass

    estado.fila_pdfs[estado.idx_pdf : estado.idx_pdf + 1] = novos_pdfs
    estado.historico.append(
        core.AcaoHistorico(
            tipo="dividir",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=estado.idx_nome if estado.idx_nome is not None else 0,
            pdf_origem=pdf_cur,
            arquivos_criados=list(novos_pdfs),
            backup_origem=backup_path,
            descricao=f"Dividido {pdf_cur.name} em {len(novos_pdfs)} partes",
        )
    )
    estado.total_divididos += 1

    if ja_renomear_parte1 and estado.item_atual() is not None:
        ok, _ = acao_confirmar_atual(estado)
        if ok:
            estado.ultimo_evento = (
                f"Dividido em {len(novos_pdfs)} partes e Parte 1 renomeada. Exibindo Parte 2."
            )
            return True, estado.ultimo_evento

    estado.ultimo_evento = f"PDF dividido em {len(novos_pdfs)} partes ({', '.join(p.name for p in novos_pdfs)})"
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_desfazer(estado: EstadoGUI) -> Tuple[bool, str]:
    novo_idx_pdf, novo_idx_nome, msg = core.desfazer_ultima_acao(
        estado.historico, estado.fila_pdfs, estado.itens_planilha, estado.log_path
    )
    msg_limpa = limpar_texto_status(msg)
    if novo_idx_pdf is None:
        return False, msg_limpa

    estado.idx_pdf = novo_idx_pdf
    estado.idx_nome = novo_idx_nome
    if estado.total_renomeados > 0 and "voltou a ser" in msg_limpa:
        estado.total_renomeados -= 1
    estado.ultimo_evento = msg_limpa
    estado.salvar_relatorio()
    return True, msg_limpa


def acao_colar_planilha_texto(estado: EstadoGUI, texto_colado: str) -> Tuple[bool, str]:
    linhas_brutas = [l.strip() for l in texto_colado.splitlines() if l.strip()]
    if not linhas_brutas:
        return False, "O texto colado está vazio."

    csv_destino = estado.script_dir / "nomes_colados.csv"
    csv_destino.write_text(texto_colado, encoding="utf-8")

    itens, msg = core.carregar_planilha(csv_destino)
    if not itens:
        return False, f"Não foi possível extrair nomes: {msg}"

    estado.planilha_path = csv_destino
    estado.recarregar_tudo()
    estado.ultimo_evento = f"{len(itens)} nomes carregados da área de transferência."
    return True, estado.ultimo_evento


# ==============================================================================
# INTERFACE WEB MINIMALISTA COM ÍCONES FONT AWESOME (SVG VETORIAL OFFLINE-SAFE)
# ==============================================================================

PAGINA_HTML = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Renomeador de PDFs</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">
<style>
  :root {
    --bg-app: #09090b;
    --bg-surface: #121215;
    --bg-elevated: #18181b;
    --bg-hover: #27272a;
    --border: #27272a;
    --border-strong: #3f3f46;
    --text-primary: #f4f4f5;
    --text-secondary: #a1a1aa;
    --text-muted: #71717a;
    --accent: #10b981;
    --accent-hover: #059669;
    --warning: #f59e0b;
    --danger: #ef4444;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI", Roboto, sans-serif;
    background: var(--bg-app);
    color: var(--text-primary);
    height: 100vh;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    user-select: none;
    -webkit-font-smoothing: antialiased;
  }

  /* Ícones SVG inline padrão Font Awesome (garante funcionamento mesmo sem internet no trabalho) */
  .fa-svg {
    width: 1em;
    height: 1em;
    fill: currentColor;
    display: inline-block;
    vertical-align: -0.125em;
    flex-shrink: 0;
  }

  /* HEADER MINIMALISTA */
  header {
    height: 46px;
    background: var(--bg-surface);
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 16px;
    flex-shrink: 0;
  }

  .brand {
    display: flex;
    align-items: center;
    gap: 9px;
    font-size: 14px;
    font-weight: 600;
    color: var(--text-primary);
    letter-spacing: -0.01em;
  }
  .brand .fa-svg { color: var(--text-secondary); }

  .header-center {
    display: flex;
    align-items: center;
    gap: 14px;
    font-size: 12.5px;
    color: var(--text-secondary);
  }
  .progress-track {
    width: 140px;
    height: 6px;
    background: var(--bg-elevated);
    border-radius: 3px;
    overflow: hidden;
    border: 1px solid var(--border);
  }
  .progress-fill {
    height: 100%;
    background: var(--accent);
    width: 0%;
    transition: width 0.2s ease;
  }

  .header-actions {
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .btn-ghost {
    background: transparent;
    color: var(--text-secondary);
    border: 1px solid var(--border);
    padding: 5px 10px;
    border-radius: 5px;
    font-size: 12px;
    font-weight: 500;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    transition: all 0.12s;
  }
  .btn-ghost:hover:not(:disabled) {
    background: var(--bg-hover);
    color: var(--text-primary);
    border-color: var(--border-strong);
  }
  .btn-ghost:disabled {
    opacity: 0.35;
    cursor: default;
  }

  /* WORKSPACE PRINCIPAL */
  .workspace {
    flex: 1;
    display: flex;
    overflow: hidden;
  }

  /* ESQUERDA: VISUALIZADOR DE PDF */
  .pdf-section {
    flex: 1;
    display: flex;
    flex-direction: column;
    background: #050507;
    border-right: 1px solid var(--border);
    min-width: 0;
  }

  .pdf-bar {
    height: 38px;
    background: var(--bg-surface);
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 14px;
    font-size: 12.5px;
  }
  .pdf-meta {
    display: flex;
    align-items: center;
    gap: 8px;
    color: var(--text-secondary);
  }
  .pdf-name {
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    color: var(--text-primary);
    font-weight: 600;
  }
  .tag-pages {
    font-size: 11px;
    padding: 2px 7px;
    border-radius: 4px;
    background: var(--bg-elevated);
    border: 1px solid var(--border);
    color: var(--text-secondary);
  }
  .tag-pages.warn {
    background: rgba(245, 158, 11, 0.12);
    border-color: rgba(245, 158, 11, 0.4);
    color: #fbbf24;
    cursor: pointer;
    font-weight: 600;
  }

  .pdf-viewport {
    flex: 1;
    position: relative;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  #pdf-frame {
    width: 100%;
    height: 100%;
    border: none;
    background: #121215;
  }

  .empty-box {
    text-align: center;
    max-width: 360px;
    padding: 32px;
    color: var(--text-secondary);
  }
  .empty-box h3 {
    color: var(--text-primary);
    font-size: 16px;
    margin: 12px 0 6px;
    font-weight: 600;
  }
  .empty-box p {
    font-size: 13px;
    line-height: 1.5;
    margin-bottom: 16px;
  }

  /* DIREITA: PAINEL DE CONTROLE LIMPO */
  .sidebar {
    width: 390px;
    display: flex;
    flex-direction: column;
    background: var(--bg-surface);
    flex-shrink: 0;
  }

  /* BLOCO DE INFORMAÇÃO DO ALUNO ATUAL */
  .current-card {
    padding: 16px;
    border-bottom: 1px solid var(--border);
  }
  .section-label {
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--text-muted);
    display: flex;
    justify-content: space-between;
    margin-bottom: 8px;
  }
  .person-name {
    font-size: 18px;
    font-weight: 600;
    color: var(--text-primary);
    line-height: 1.3;
    margin-bottom: 4px;
    user-select: text;
  }
  .person-id {
    font-size: 13px;
    color: var(--text-secondary);
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    margin-bottom: 12px;
    user-select: text;
  }
  .target-file-box {
    background: var(--bg-app);
    border: 1px solid var(--border-strong);
    border-radius: 6px;
    padding: 8px 10px;
    display: flex;
    align-items: center;
    gap: 8px;
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    font-size: 13px;
    color: var(--accent);
    word-break: break-all;
    user-select: text;
  }

  /* GRADE DE COMANDOS */
  .controls-block {
    padding: 14px 16px;
    border-bottom: 1px solid var(--border);
    display: flex;
    flex-direction: column;
    gap: 8px;
  }

  .btn-primary {
    width: 100%;
    height: 40px;
    border: none;
    border-radius: 6px;
    background: var(--accent);
    color: #052e16;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 14px;
    transition: background 0.12s;
  }
  .btn-primary:hover { background: var(--accent-hover); color: #ffffff; }

  .btn-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px;
  }

  .btn-secondary {
    height: 34px;
    padding: 0 10px;
    border-radius: 5px;
    border: 1px solid var(--border);
    background: var(--bg-elevated);
    color: var(--text-primary);
    font-size: 12.5px;
    font-weight: 500;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: space-between;
    transition: background 0.12s, border-color 0.12s;
  }
  .btn-secondary:hover {
    background: var(--bg-hover);
    border-color: var(--border-strong);
  }
  .btn-label {
    display: inline-flex;
    align-items: center;
    gap: 7px;
  }
  .btn-label .fa-svg { color: var(--text-secondary); }

  .shortcut {
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    font-size: 11px;
    color: var(--text-muted);
    background: rgba(255, 255, 255, 0.05);
    padding: 1px 5px;
    border-radius: 3px;
  }
  .btn-primary .shortcut {
    color: #052e16;
    background: rgba(0, 0, 0, 0.12);
  }
  .btn-primary:hover .shortcut { color: #ffffff; }

  /* PAINÉIS CONTEXTUAIS (EDITAR / DIVIDIR / PLANILHA) */
  .panel-drawer {
    display: none;
    padding: 14px 16px;
    background: var(--bg-elevated);
    border-bottom: 1px solid var(--border-strong);
  }
  .panel-drawer.open { display: block; }
  .drawer-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: 12px;
    font-weight: 600;
    color: var(--text-primary);
    margin-bottom: 10px;
  }
  .input-clean {
    width: 100%;
    padding: 8px 10px;
    background: var(--bg-app);
    border: 1px solid var(--border-strong);
    border-radius: 5px;
    color: var(--text-primary);
    font-size: 13px;
    margin-bottom: 8px;
  }
  .input-clean:focus { outline: none; border-color: var(--accent); }

  .radio-group {
    display: flex;
    flex-direction: column;
    gap: 6px;
    font-size: 12px;
    color: var(--text-secondary);
    margin: 6px 0 12px;
  }
  .radio-group label {
    display: flex;
    align-items: center;
    gap: 7px;
    cursor: pointer;
  }

  /* LISTA DA FILA */
  .queue-section {
    flex: 1;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    padding: 12px 16px;
  }
  .search-wrap {
    position: relative;
    margin-bottom: 10px;
  }
  .search-wrap .fa-svg {
    position: absolute;
    left: 10px;
    top: 50%;
    transform: translateY(-50%);
    color: var(--text-muted);
    font-size: 12px;
  }
  .search-input {
    width: 100%;
    padding: 7px 10px 7px 30px;
    background: var(--bg-app);
    border: 1px solid var(--border);
    border-radius: 5px;
    color: var(--text-primary);
    font-size: 12.5px;
  }
  .search-input:focus { outline: none; border-color: var(--border-strong); }

  .queue-list {
    flex: 1;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 4px;
  }
  .queue-row {
    padding: 7px 9px;
    border-radius: 5px;
    border: 1px solid transparent;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    font-size: 12.5px;
    transition: background 0.1s;
  }
  .queue-row:hover {
    background: var(--bg-elevated);
  }
  .queue-row.current {
    background: rgba(16, 185, 129, 0.08);
    border-color: rgba(16, 185, 129, 0.3);
  }
  .queue-row.done {
    opacity: 0.4;
  }
  .row-main {
    min-width: 0;
    flex: 1;
  }
  .row-title {
    font-weight: 500;
    color: var(--text-primary);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .row-sub {
    font-size: 11px;
    color: var(--text-muted);
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .row-actions {
    display: none;
    gap: 4px;
    flex-shrink: 0;
  }
  .queue-row:hover .row-actions { display: flex; }

  .btn-row {
    background: var(--bg-hover);
    color: var(--text-primary);
    border: 1px solid var(--border-strong);
    padding: 3px 7px;
    border-radius: 4px;
    font-size: 11px;
    cursor: pointer;
  }
  .btn-row:hover { background: var(--accent); color: #052e16; border-color: var(--accent); }

  /* RODAPÉ DISCRETO */
  footer {
    height: 28px;
    background: var(--bg-surface);
    border-top: 1px solid var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 16px;
    font-size: 11.5px;
    color: var(--text-muted);
    flex-shrink: 0;
  }
</style>
</head>
<body>

<!-- DEFINIÇÃO DE ÍCONES FONT AWESOME (SVG VETORIAL EMBUTIDO) -->
<svg style="display:none;" xmlns="http://www.w3.org/2000/svg">
  <symbol id="fa-file-pdf" viewBox="0 0 512 512"><path d="M0 64C0 28.7 28.7 0 64 0L224 0l0 128c0 17.7 14.3 32 32 32l128 0 0 144-208 0c-35.3 0-64 28.7-64 64l0 144-48 0c-35.3 0-64-28.7-64-64L0 64zm384 64l-128 0L256 0 384 128zM176 352l32 0c30.9 0 56 25.1 56 56s-25.1 56-56 56l-16 0 0 32c0 8.8-7.2 16-16 16s-16-7.2-16-16l0-48 0-80c0-8.8 7.2-16 16-16zm32 80c13.3 0 24-10.7 24-24s-10.7-24-24-24l-16 0 0 48 16 0zm96-80l32 0c26.5 0 48 21.5 48 48l0 64c0 26.5-21.5 48-48 48l-32 0c-8.8 0-16-7.2-16-16l0-128c0-8.8 7.2-16 16-16zm32 128c8.8 0 16-7.2 16-16l0-64c0-8.8-7.2-16-16-16l-16 0 0 96 16 0zm80-112c0-8.8 7.2-16 16-16l48 0c8.8 0 16 7.2 16 16s-7.2 16-16 16l-32 0 0 32 32 0c8.8 0 16 7.2 16 16s-7.2 16-16 16l-32 0 0 48c0 8.8-7.2 16-16 16s-16-7.2-16-16l0-64 0-64z"/></symbol>
  <symbol id="fa-check" viewBox="0 0 448 512"><path d="M438.6 105.4c12.5 12.5 12.5 32.8 0 45.3l-256 256c-12.5 12.5-32.8 12.5-45.3 0l-128-128c-12.5-12.5-12.5-32.8 0-45.3s32.8-12.5 45.3 0L160 338.7 393.4 105.4c12.5-12.5 32.8-12.5 45.3 0z"/></symbol>
  <symbol id="fa-pen" viewBox="0 0 512 512"><path d="M362.7 19.3L314.3 67.7 444.3 197.7l48.4-48.4c25-25 25-65.5 0-90.5L453.3 19.3c-25-25-65.5-25-90.5 0zm-71 71L58.6 323.5c-10.4 10.4-18 23.3-22.2 37.4L1 481.2C-1.5 489.7 .8 498.8 7 505s15.3 8.5 23.7 6.1l120.3-35.4c14.1-4.2 27-11.8 37.4-22.2L421.7 220.3 291.7 90.3z"/></symbol>
  <symbol id="fa-scissors" viewBox="0 0 512 512"><path d="M256 192l-39.5-39.5c4.9-12.6 7.5-26.2 7.5-40.5C224 50.1 173.9 0 112 0S0 50.1 0 112s50.1 112 112 112c14.3 0 27.9-2.7 40.5-7.5L192 256l-39.5 39.5c-12.6-4.9-26.2-7.5-40.5-7.5C50.1 288 0 338.1 0 400s50.1 112 112 112s112-50.1 112-112c0-14.3-2.7-27.9-7.5-40.5L499.3 76.7c6.2-6.2 6.2-16.4 0-22.6L457.9 12.7c-6.2-6.2-16.4-6.2-22.6 0L256 192zm-144-32a48 48 0 1 1 0-96 48 48 0 1 1 0 96zm0 288a48 48 0 1 1 0-96 48 48 0 1 1 0 96zm192-128l133.3 133.3c6.2 6.2 16.4 6.2 22.6 0l41.4-41.4c6.2-6.2 6.2-16.4 0-22.6L368 256l-64 64z"/></symbol>
  <symbol id="fa-forward-step" viewBox="0 0 320 512"><path d="M52.5 440.6c-9.5 7.9-22.8 9.7-34.1 4.4S0 428.4 0 416L0 96C0 83.6 7.2 72.3 18.4 67s24.5-3.6 34.1 4.4l192 160L256 241l0-145c0-17.7 14.3-32 32-32s32 14.3 32 32l0 320c0 17.7-14.3 32-32 32s-32-14.3-32-32l0-145-11.5 9.6-192 160z"/></symbol>
  <symbol id="fa-rotate-left" viewBox="0 0 512 512"><path d="M125.7 160l50.3 0c17.7 0 32 14.3 32 32s-14.3 32-32 32L48 224c-17.7 0-32-14.3-32-32L16 64c0-17.7 14.3-32 32-32s32 14.3 32 32l0 51.2L97.6 97.6c87.5-87.5 229.3-87.5 316.8 0s87.5 229.3 0 316.8s-229.3 87.5-316.8 0c-12.5-12.5-12.5-32.8 0-45.3s32.8-12.5 45.3 0c62.5 62.5 163.8 62.5 226.3 0s62.5-163.8 0-226.3s-163.8-62.5-226.3 0L125.7 160z"/></symbol>
  <symbol id="fa-magnifying-glass" viewBox="0 0 512 512"><path d="M416 208c0 45.9-14.9 88.3-40 122.7L502.6 457.4c12.5 12.5 12.5 32.8 0 45.3s-32.8 12.5-45.3 0L330.7 376c-34.4 25.2-76.8 40-122.7 40C93.1 416 0 322.9 0 208S93.1 0 208 0S416 93.1 416 208zM208 352a144 144 0 1 0 0-288 144 144 0 1 0 0 288z"/></symbol>
  <symbol id="fa-arrows-rotate" viewBox="0 0 512 512"><path d="M105.1 202.6c7.7-21.8 20.2-42.3 37.8-59.8c62.5-62.5 163.8-62.5 226.3 0L386.3 160 352 160c-17.7 0-32 14.3-32 32s14.3 32 32 32l111.5 0c0 0 0 0 0 0l.4 0c17.7 0 32-14.3 32-32l0-112c0-17.7-14.3-32-32-32s-32 14.3-32 32l0 35.2L414.4 97.6c-87.5-87.5-229.3-87.5-316.8 0C73.2 122 55.6 150.7 44.8 181.4c-5.9 16.7 2.9 34.9 19.5 40.8s34.9-2.9 40.8-19.5zM39 289.3c-5 1.5-9.8 4.2-13.7 8.2c-4 4-6.7 8.8-8.1 14c-.3 1.2-.6 2.5-.8 3.8c-.3 1.7-.4 3.4-.4 5.1L16 432c0 17.7 14.3 32 32 32s32-14.3 32-32l0-35.1 17.6 17.5c0 0 0 0 0 0c87.5 87.4 229.3 87.4 316.7 0c24.4-24.4 42.1-53.1 52.9-83.8c5.9-16.7-2.9-34.9-19.5-40.8s-34.9 2.9-40.8 19.5c-7.7 21.8-20.2 42.3-37.8 59.8c-62.5 62.5-163.8 62.5-226.3 0l-.1-.1L125.6 352l34.4 0c17.7 0 32-14.3 32-32s-14.3-32-32-32L48.4 288c-1.6 0-3.2 .1-4.8 .3s-3.1 .5-4.6 1z"/></symbol>
  <symbol id="fa-table" viewBox="0 0 512 512"><path d="M64 256l0-96 160 0 0 96L64 256zm0 64l160 0 0 96L64 416l0-96zm224 96l0-96 160 0 0 96-160 0zM448 256l-160 0 0-96 160 0 0 96zM64 32C28.7 32 0 60.7 0 96L0 416c0 35.3 28.7 64 64 64l384 0c35.3 0 64-28.7 64-64l0-320c0-35.3-28.7-64-64-64L64 32z"/></symbol>
  <symbol id="fa-xmark" viewBox="0 0 384 512"><path d="M342.6 150.6c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0L192 210.7 86.6 105.4c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3L146.7 256 41.4 361.4c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0L192 301.3 297.4 406.6c12.5 12.5 32.8 12.5 45.3 0s12.5-32.8 0-45.3L237.3 256 342.6 150.6z"/></symbol>
</svg>

<header>
  <div class="brand">
    <svg class="fa-svg"><use href="#fa-file-pdf"></use></svg>
    <span>Renomeador de PDFs</span>
  </div>

  <div class="header-center">
    <span id="lbl-progresso">0 / 0</span>
    <div class="progress-track">
      <div class="progress-fill" id="barra-progresso"></div>
    </div>
    <span>Sessão: <strong id="lbl-renomeados" style="color:var(--text-primary);">0</strong></span>
  </div>

  <div class="header-actions">
    <button class="btn-ghost" id="btn-undo" onclick="desfazerUltimo()" title="Desfazer última ação (Tecla Z)">
      <svg class="fa-svg"><use href="#fa-rotate-left"></use></svg>
      <span>Desfazer</span>
      <span class="shortcut">Z</span>
    </button>
    <button class="btn-ghost" onclick="alternarGaveta('gaveta-planilha')" title="Colar dados do Excel ou recarregar planilha">
      <svg class="fa-svg"><use href="#fa-table"></use></svg>
      <span>Planilha</span>
    </button>
    <button class="btn-ghost" onclick="reescanearPasta()" title="Atualizar arquivos da pasta pdfs/">
      <svg class="fa-svg"><use href="#fa-arrows-rotate"></use></svg>
      <span>Atualizar</span>
    </button>
  </div>
</header>

<div class="workspace">
  <!-- PAINEL ESQUERDO: VISUALIZADOR DE PDF -->
  <section class="pdf-section">
    <div class="pdf-bar">
      <div class="pdf-meta">
        <svg class="fa-svg"><use href="#fa-file-pdf"></use></svg>
        <span class="pdf-name" id="lbl-pdf-nome">—</span>
        <span id="lbl-pdf-pos">(0 / 0)</span>
      </div>
      <div>
        <span class="tag-pages" id="badge-paginas" onclick="alternarGaveta('gaveta-dividir')">—</span>
      </div>
    </div>

    <div class="pdf-viewport">
      <iframe id="pdf-frame" title="Visualizador de PDF"></iframe>
      <div class="empty-box" id="empty-state" style="display:none;">
        <svg class="fa-svg" style="font-size:28px; color:var(--accent);"><use href="#fa-check"></use></svg>
        <h3>Nenhum PDF pendente</h3>
        <p id="empty-msg">Todos os arquivos da pasta <code>pdfs/</code> foram processados.</p>
        <button class="btn-ghost" onclick="reescanearPasta()">
          <svg class="fa-svg"><use href="#fa-arrows-rotate"></use></svg>
          <span>Verificar novamente</span>
        </button>
      </div>
    </div>
  </section>

  <!-- PAINEL DIREITO: CONTROLES E FILA -->
  <aside class="sidebar">
    <!-- DADOS DO ITEM ATUAL -->
    <div class="current-card">
      <div class="section-label">
        <span>Nome na fila</span>
        <span id="lbl-item-idx">#—</span>
      </div>
      <div class="person-name" id="lbl-aluno-nome">—</div>
      <div class="person-id" id="lbl-aluno-mat">—</div>
      <div class="target-file-box">
        <svg class="fa-svg" style="color:var(--accent);"><use href="#fa-file-pdf"></use></svg>
        <span id="lbl-aluno-formatado">—.pdf</span>
      </div>
    </div>

    <!-- AÇÕES PRINCIPAIS -->
    <div class="controls-block">
      <button class="btn-primary" id="btn-confirmar" onclick="confirmarAtual()">
        <span class="btn-label">
          <svg class="fa-svg"><use href="#fa-check"></use></svg>
          <span>Confirmar e Renomear</span>
        </span>
        <span class="shortcut">S / Enter</span>
      </button>

      <div class="btn-grid">
        <button class="btn-secondary" onclick="alternarGaveta('gaveta-editar')">
          <span class="btn-label">
            <svg class="fa-svg"><use href="#fa-pen"></use></svg>
            <span>Editar</span>
          </span>
          <span class="shortcut">E</span>
        </button>
        <button class="btn-secondary" onclick="alternarGaveta('gaveta-dividir')">
          <span class="btn-label">
            <svg class="fa-svg"><use href="#fa-scissors"></use></svg>
            <span>Dividir PDF</span>
          </span>
          <span class="shortcut">D</span>
        </button>
        <button class="btn-secondary" onclick="pularNome()" title="Avança apenas o nome da planilha e mantém o PDF atual">
          <span class="btn-label">
            <svg class="fa-svg"><use href="#fa-forward-step"></use></svg>
            <span>Pular Nome</span>
          </span>
          <span class="shortcut">N</span>
        </button>
        <button class="btn-secondary" onclick="pularPdf()" title="Pula este arquivo PDF e mantém o nome atual para o próximo">
          <span class="btn-label">
            <svg class="fa-svg"><use href="#fa-forward-step"></use></svg>
            <span>Pular PDF</span>
          </span>
          <span class="shortcut">P</span>
        </button>
      </div>
    </div>

    <!-- GAVETA: EDITAR / PESSOA EXTRA -->
    <div class="panel-drawer" id="gaveta-editar">
      <div class="drawer-header">
        <span>Editar nome / Pessoa fora da planilha</span>
        <button class="btn-ghost" style="border:none; padding:2px 5px;" onclick="fecharGavetas()">
          <svg class="fa-svg"><use href="#fa-xmark"></use></svg>
        </button>
      </div>
      <input type="text" id="inp-edit-nome" class="input-clean"
             placeholder="Nome Completo e Matrícula"
             oninput="atualizarPreviewEdicao()"
             onkeydown="if(event.key==='Enter'){salvarEdicaoManual();} else if(event.key==='Escape'){fecharGavetas();}">
      <div style="font-size:11.5px; color:var(--accent); font-family:monospace; margin-bottom:6px;" id="preview-edit-formatado">
        Arquivo: —
      </div>
      <div class="radio-group">
        <label>
          <input type="radio" name="modo-edit" value="extra" checked>
          <span>Pessoa fora da planilha (manter fila atual)</span>
        </label>
        <label>
          <input type="radio" name="modo-edit" value="substituir">
          <span>Substituir o nome atual da fila</span>
        </label>
      </div>
      <div class="btn-grid">
        <button class="btn-primary" style="height:32px; font-size:12.5px; justify-content:center;" onclick="salvarEdicaoManual()">Salvar</button>
        <button class="btn-secondary" style="height:32px; justify-content:center;" onclick="fecharGavetas()">Cancelar</button>
      </div>
    </div>

    <!-- GAVETA: DIVIDIR PDF -->
    <div class="panel-drawer" id="gaveta-dividir">
      <div class="drawer-header">
        <span id="lbl-info-divisao">Dividir arquivo PDF</span>
        <button class="btn-ghost" style="border:none; padding:2px 5px;" onclick="fecharGavetas()">
          <svg class="fa-svg"><use href="#fa-xmark"></use></svg>
        </button>
      </div>
      <div style="display:flex; flex-direction:column; gap:6px;">
        <button class="btn-primary" style="height:34px; font-size:12.5px; justify-content:center;"
                onclick="executarDivisao('corte', 1, true)">
          Separar Pág. 1 (renomear) + Restante
        </button>
        <div style="display:flex; gap:6px; align-items:center;">
          <span style="font-size:12px; color:var(--text-secondary);">1ª parte até pág.:</span>
          <input type="number" id="inp-corte-pag" value="1" min="1" max="99" class="input-clean" style="width:60px; margin:0; padding:5px;">
          <button class="btn-secondary" style="flex:1; justify-content:center;" onclick="executarDivisaoCustom()">Cortar página</button>
        </div>
        <div class="btn-grid">
          <button class="btn-secondary" style="justify-content:center;" onclick="executarDivisao('todas', 1, false)">1 PDF por pág.</button>
          <button class="btn-secondary" style="justify-content:center;" onclick="executarDivisao('duplicar', 1, true)">Duplicar arquivo</button>
        </div>
      </div>
    </div>

    <!-- GAVETA: COLAR PLANILHA DO EXCEL -->
    <div class="panel-drawer" id="gaveta-planilha">
      <div class="drawer-header">
        <span>Colar dados da planilha (Nome + Matrícula)</span>
        <button class="btn-ghost" style="border:none; padding:2px 5px;" onclick="fecharGavetas()">
          <svg class="fa-svg"><use href="#fa-xmark"></use></svg>
        </button>
      </div>
      <textarea id="txt-colar-excel" class="input-clean" style="height:90px; font-family:monospace; font-size:11.5px; resize:none;"
                placeholder="Cole aqui as colunas copiadas do Excel..."></textarea>
      <div class="btn-grid">
        <button class="btn-primary" style="height:32px; font-size:12.5px; justify-content:center;" onclick="enviarPlanilhaColada()">Carregar texto</button>
        <button class="btn-secondary" style="height:32px; justify-content:center;" onclick="recarregarPlanilhaDisco()">Ler nomes.xlsx</button>
      </div>
    </div>

    <!-- FILA DE NOMES E BUSCA -->
    <div class="queue-section">
      <div class="search-wrap">
        <svg class="fa-svg"><use href="#fa-magnifying-glass"></use></svg>
        <input type="text" id="inp-busca" class="search-input"
               placeholder="Buscar nome ou matrícula (B)..."
               oninput="renderizarListaNomes()"
               onkeydown="if(event.key==='Escape'){this.value=''; renderizarListaNomes(); this.blur();}">
      </div>
      <div class="queue-list" id="lista-nomes"></div>
    </div>
  </aside>
</div>

<footer>
  <span id="status-toast">Pronto</span>
  <span>Teclas: <strong>S</strong> Confirmar · <strong>E</strong> Editar · <strong>D</strong> Dividir · <strong>B</strong> Buscar · <strong>N</strong> Pular Nome · <strong>P</strong> Pular PDF · <strong>Z</strong> Desfazer</span>
</footer>

<link rel="prefetch" id="pdf-prefetch" href="">

<script>
let ESTADO = null;
let urlPdfAtualNoIframe = "";
const PARTICULAS = new Set(["de", "da", "do", "dos", "das", "e", "del", "d", "d'"]);

function formatarPreviewLocal(texto) {
  let s = (texto || "").trim();
  if (!s) return "—";
  if (s.toLowerCase().endsWith(".pdf")) s = s.slice(0, -4).trim();
  if (s.includes("_") && !s.includes(" ")) return s + ".pdf";

  let mat = "";
  const m = s.match(/^(.+?)[\s\-_;,]+(\d{4,})$/);
  if (m) {
    s = m[1].trim();
    mat = m[2].trim();
  }
  const palavras = s.replace(/[_\-]+/g, " ").split(/\s+/).filter(Boolean);
  const filtradas = palavras.filter(p => !PARTICULAS.has(p.toLowerCase()));
  const escolhidas = (filtradas.length ? filtradas : palavras).slice(0, 2).map(p => {
    return p.charAt(0).toUpperCase() + p.slice(1).toLowerCase();
  });
  const base = mat ? [...escolhidas, mat].join("_") : escolhidas.join("_");
  return (base || "—") + ".pdf";
}

function atualizarPreviewEdicao() {
  const val = document.getElementById("inp-edit-nome").value;
  document.getElementById("preview-edit-formatado").textContent = "Arquivo: " + formatarPreviewLocal(val);
}

async function apiPost(rota, dados = {}) {
  try {
    const resp = await fetch(rota, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(dados)
    });
    const json = await resp.json();
    if (json.estado) {
      aplicarEstado(json.estado);
    }
    if (!json.ok && json.mensagem) {
      mostrarStatus(json.mensagem, true);
    }
    return json;
  } catch (err) {
    mostrarStatus("Erro de comunicação local", true);
  }
}

async function carregarEstadoInicial() {
  const resp = await fetch("/api/state");
  const json = await resp.json();
  aplicarEstado(json);
}

function mostrarStatus(msg, erro = false) {
  const el = document.getElementById("status-toast");
  el.textContent = msg;
  el.style.color = erro ? "#ef4444" : "#a1a1aa";
}

function aplicarEstado(st) {
  ESTADO = st;
  const p = st.progresso;
  document.getElementById("lbl-progresso").textContent =
    `${p.concluidos_planilha} / ${p.total_planilha} (${p.percentual}%)`;
  document.getElementById("barra-progresso").style.width = `${p.percentual}%`;
  document.getElementById("lbl-renomeados").textContent = p.renomeados_sessao;
  document.getElementById("btn-undo").disabled = !p.pode_desfazer;

  mostrarStatus(st.ultimo_evento || "Pronto", Boolean(st.erro_carregamento));

  const iframe = document.getElementById("pdf-frame");
  const empty = document.getElementById("empty-state");

  if (st.tem_pdf && st.pdf_atual) {
    iframe.style.display = "block";
    empty.style.display = "none";
    document.getElementById("lbl-pdf-nome").textContent = st.pdf_atual.nome;
    document.getElementById("lbl-pdf-pos").textContent =
      `(${st.pdf_atual.posicao} / ${st.pdf_atual.total_fila})`;

    const badge = document.getElementById("badge-paginas");
    const pags = st.pdf_atual.paginas;
    if (pags && pags >= 2) {
      badge.textContent = `${pags} páginas — Dividir (D)`;
      badge.className = "tag-pages warn";
    } else if (pags === 1) {
      badge.textContent = "1 página";
      badge.className = "tag-pages";
    } else {
      badge.textContent = "PDF";
      badge.className = "tag-pages";
    }

    const novaUrl = st.pdf_atual.url + "#toolbar=1&navpanes=0&view=FitH";
    if (urlPdfAtualNoIframe !== novaUrl) {
      urlPdfAtualNoIframe = novaUrl;
      iframe.src = novaUrl;
    }

    document.getElementById("pdf-prefetch").href = st.pdf_proximo_url || "";
  } else {
    iframe.style.display = "none";
    empty.style.display = "block";
    urlPdfAtualNoIframe = "";
    iframe.src = "about:blank";
    document.getElementById("lbl-pdf-nome").textContent = "—";
    document.getElementById("lbl-pdf-pos").textContent = "";
    document.getElementById("badge-paginas").textContent = "—";
    document.getElementById("badge-paginas").className = "tag-pages";
    if (st.erro_carregamento) {
      document.getElementById("empty-msg").textContent = st.erro_carregamento;
    }
  }

  if (st.item_atual) {
    document.getElementById("lbl-item-idx").textContent =
      `#${st.item_atual.indice_original} de ${p.total_planilha}`;
    document.getElementById("lbl-aluno-nome").textContent = st.item_atual.nome_completo;
    document.getElementById("lbl-aluno-mat").textContent =
      st.item_atual.matricula || "Sem matrícula separada";
    document.getElementById("lbl-aluno-formatado").textContent =
      `${st.item_atual.nome_formatado}.pdf`;
  } else {
    document.getElementById("lbl-item-idx").textContent = "Concluído";
    document.getElementById("lbl-aluno-nome").textContent = "Fila finalizada";
    document.getElementById("lbl-aluno-mat").textContent = "Use Editar (E) para informar outro nome";
    document.getElementById("lbl-aluno-formatado").textContent = "—";
  }

  renderizarListaNomes();
}

function renderizarListaNomes() {
  if (!ESTADO) return;
  const container = document.getElementById("lista-nomes");
  const filtro = (document.getElementById("inp-busca").value || "").trim().toLowerCase();

  const itens = ESTADO.fila_nomes.filter(it => {
    if (!filtro) return true;
    const alvo = `${it.nome_completo} ${it.matricula} ${it.nome_formatado}`.toLowerCase();
    return alvo.includes(filtro);
  });

  container.innerHTML = itens.map(it => {
    const cls = it.atual ? "queue-row current" : (it.concluido ? "queue-row done" : "queue-row");
    return `
      <div class="${cls}">
        <div class="row-main">
          <div class="row-title">${it.indice_original}. ${it.nome_completo}</div>
          <div class="row-sub">${it.nome_formatado}.pdf</div>
        </div>
        <div class="row-actions">
          <button class="btn-row" onclick="confirmarComIndice(${it.idx})" title="Renomear PDF atual com este nome">
            Usar
          </button>
          ${!it.concluido && !it.atual ? `
            <button class="btn-row" onclick="apontarFila(${it.idx})" title="Posicionar fila neste item">
              Focar
            </button>
          ` : ""}
        </div>
      </div>
    `;
  }).join("");
}

function fecharGavetas() {
  for (const id of ["gaveta-editar", "gaveta-dividir", "gaveta-planilha"]) {
    document.getElementById(id).classList.remove("open");
  }
}

function alternarGaveta(id) {
  const el = document.getElementById(id);
  const estavaAberta = el.classList.contains("open");
  fecharGavetas();
  if (!estavaAberta) {
    el.classList.add("open");
    if (id === "gaveta-editar") {
      const inp = document.getElementById("inp-edit-nome");
      inp.value = "";
      atualizarPreviewEdicao();
      setTimeout(() => inp.focus(), 20);
    } else if (id === "gaveta-dividir" && ESTADO && ESTADO.pdf_atual) {
      const pags = ESTADO.pdf_atual.paginas;
      document.getElementById("lbl-info-divisao").textContent =
        pags ? `Dividir PDF (${pags} páginas)` : "Dividir PDF";
    } else if (id === "gaveta-planilha") {
      setTimeout(() => document.getElementById("txt-colar-excel").focus(), 20);
    }
  }
}

async function confirmarAtual() {
  fecharGavetas();
  await apiPost("/api/confirm", {});
}

async function confirmarComIndice(idx) {
  fecharGavetas();
  await apiPost("/api/confirm", { idx_escolhido: idx });
}

async function apontarFila(idx) {
  await apiPost("/api/select_name", { idx_nome: idx });
}

async function salvarEdicaoManual() {
  const nome = document.getElementById("inp-edit-nome").value.trim();
  if (!nome) return;
  const modo = document.querySelector('input[name="modo-edit"]:checked').value;
  fecharGavetas();
  await apiPost("/api/edit", {
    nome: nome,
    matricula: "",
    avancar_planilha: modo === "substituir"
  });
}

async function executarDivisao(modo, paginaCorte, jaRenomearParte1) {
  fecharGavetas();
  await apiPost("/api/split", {
    modo: modo,
    pagina_corte: paginaCorte,
    ja_renomear_parte1: jaRenomearParte1
  });
}

async function executarDivisaoCustom() {
  const corte = parseInt(document.getElementById("inp-corte-pag").value || "1", 10);
  await executarDivisao("corte", corte, true);
}

async function pularNome() {
  fecharGavetas();
  await apiPost("/api/skip_name", {});
}

async function pularPdf() {
  fecharGavetas();
  await apiPost("/api/skip_pdf", { avancar_nome_junto: false });
}

async function desfazerUltimo() {
  fecharGavetas();
  await apiPost("/api/undo", {});
}

async function reescanearPasta() {
  await apiPost("/api/rescan", {});
}

async function recarregarPlanilhaDisco() {
  fecharGavetas();
  await apiPost("/api/reload_sheet", {});
}

async function enviarPlanilhaColada() {
  const texto = document.getElementById("txt-colar-excel").value;
  const res = await apiPost("/api/paste_sheet", { texto: texto });
  if (res && res.ok) {
    fecharGavetas();
  }
}

// Mantém os atalhos de 1 tecla ativos mesmo após clicar dentro do PDF
setInterval(() => {
  if (document.activeElement && document.activeElement.id === "pdf-frame") {
    document.activeElement.blur();
    window.focus();
  }
}, 250);

window.addEventListener("keydown", (e) => {
  const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : "";
  if (tag === "input" || tag === "textarea" || tag === "select") {
    return;
  }

  if (e.key === "Escape") {
    fecharGavetas();
    return;
  }

  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
    e.preventDefault();
    desfazerUltimo();
    return;
  }

  if (e.ctrlKey || e.altKey || e.metaKey) return;

  const k = e.key.toLowerCase();
  if (k === "s" || e.key === "Enter") {
    e.preventDefault();
    confirmarAtual();
  } else if (k === "e") {
    e.preventDefault();
    alternarGaveta("gaveta-editar");
  } else if (k === "d") {
    e.preventDefault();
    alternarGaveta("gaveta-dividir");
  } else if (k === "b" || k === "/") {
    e.preventDefault();
    const inp = document.getElementById("inp-busca");
    inp.focus();
    inp.select();
  } else if (k === "n") {
    e.preventDefault();
    pularNome();
  } else if (k === "p") {
    e.preventDefault();
    pularPdf();
  } else if (k === "z") {
    e.preventDefault();
    desfazerUltimo();
  }
});

window.addEventListener("DOMContentLoaded", carregarEstadoInicial);
</script>
</body>
</html>
"""


class ManipuladorHTTPGUI(BaseHTTPRequestHandler):
    estado: EstadoGUI = None  # type: ignore

    def log_message(self, format, *args):
        pass

    def _enviar_json(self, dados: dict, status: int = 200):
        corpo = json.dumps(dados, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(corpo)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        rota = parsed.path

        if rota == "/" or rota == "/index.html":
            corpo = PAGINA_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(corpo)))
            self.end_headers()
            self.wfile.write(corpo)
            return

        if rota == "/api/state":
            with self.estado.lock:
                self._enviar_json(self.estado.serializar_estado())
            return

        if rota == "/api/pdf":
            qs = urllib.parse.parse_qs(parsed.query)
            nome_arq = qs.get("name", [""])[0]
            nome_seguro = Path(nome_arq).name
            caminho_pdf = (self.estado.pdfs_dir / nome_seguro).resolve()
            if (
                not nome_seguro
                or not caminho_pdf.exists()
                or not caminho_pdf.is_file()
                or caminho_pdf.parent != self.estado.pdfs_dir.resolve()
            ):
                self.send_error(404, "Arquivo PDF nao encontrado")
                return

            dados_pdf = caminho_pdf.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'inline; filename="{nome_seguro}"')
            self.send_header("Content-Length", str(len(dados_pdf)))
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()
            self.wfile.write(dados_pdf)
            return

        self.send_error(404, "Nao encontrado")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        rota = parsed.path
        tamanho = int(self.headers.get("Content-Length", "0"))
        corpo_raw = self.rfile.read(tamanho) if tamanho > 0 else b"{}"
        try:
            payload = json.loads(corpo_raw.decode("utf-8"))
        except Exception:
            payload = {}

        with self.estado.lock:
            ok = True
            msg = ""

            if rota == "/api/confirm":
                idx_esc = payload.get("idx_escolhido")
                ok, msg = acao_confirmar_atual(
                    self.estado, int(idx_esc) if idx_esc is not None else None
                )

            elif rota == "/api/edit":
                ok, msg = acao_editar_manual(
                    self.estado,
                    str(payload.get("nome", "")),
                    str(payload.get("matricula", "")),
                    bool(payload.get("avancar_planilha", False)),
                )

            elif rota == "/api/split":
                ok, msg = acao_dividir_pdf(
                    self.estado,
                    modo=str(payload.get("modo", "corte")),
                    pagina_corte=int(payload.get("pagina_corte", 1)),
                    ja_renomear_parte1=bool(payload.get("ja_renomear_parte1", True)),
                )

            elif rota == "/api/skip_name":
                ok, msg = acao_pular_nome(self.estado)

            elif rota == "/api/skip_pdf":
                ok, msg = acao_pular_pdf(
                    self.estado, bool(payload.get("avancar_nome_junto", False))
                )

            elif rota == "/api/undo":
                ok, msg = acao_desfazer(self.estado)

            elif rota == "/api/rescan":
                self.estado.reescanear_pasta_pdfs()
                msg = f"Pasta atualizada: {len(self.estado.fila_pdfs)} PDF(s) pendente(s)."

            elif rota == "/api/reload_sheet":
                self.estado.planilha_path = self.estado.script_dir / core.NOME_PLANILHA_PADRAO
                self.estado.recarregar_tudo()
                msg = f"Planilha {self.estado.planilha_path.name} recarregada."

            elif rota == "/api/paste_sheet":
                ok, msg = acao_colar_planilha_texto(
                    self.estado, str(payload.get("texto", ""))
                )

            elif rota == "/api/select_name":
                idx_n = payload.get("idx_nome")
                if idx_n is not None and 0 <= int(idx_n) < len(self.estado.itens_planilha):
                    self.estado.idx_nome = int(idx_n)
                    self.estado.ultimo_evento = (
                        f"Fila posicionada em {self.estado.itens_planilha[self.estado.idx_nome].nome_formatado}"
                    )
                    msg = self.estado.ultimo_evento

            else:
                self._enviar_json({"ok": False, "mensagem": "Rota inválida"}, status=404)
                return

            self._enviar_json(
                {
                    "ok": ok,
                    "mensagem": msg,
                    "estado": self.estado.serializar_estado(),
                }
            )


def escolher_porta_livre(porta_inicial: int) -> int:
    for p in range(porta_inicial, porta_inicial + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind((HOST_LOCAL, p))
                return p
            except OSError:
                continue
    return porta_inicial


def main():
    core.ativar_cores_windows()
    parser = argparse.ArgumentParser(
        description="Renomeador de PDFs v5.0 - Interface Gráfica Web Local (127.0.0.1)"
    )
    parser.add_argument(
        "--planilha",
        default=core.NOME_PLANILHA_PADRAO,
        help=f"Caminho da planilha Excel ou CSV (padrão: {core.NOME_PLANILHA_PADRAO})",
    )
    parser.add_argument(
        "--pasta_pdfs",
        default=core.PASTA_PDFS_PADRAO,
        help=f"Pasta com os arquivos PDF (padrão: {core.PASTA_PDFS_PADRAO})",
    )
    parser.add_argument(
        "--porta",
        type=int,
        default=PORTA_PADRAO,
        help=f"Porta local (padrão: {PORTA_PADRAO})",
    )
    parser.add_argument(
        "--nao-abrir-navegador",
        action="store_true",
        help="Não abre o navegador automaticamente ao iniciar",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    planilha_path = Path(args.planilha)
    if not planilha_path.is_absolute() and not planilha_path.exists():
        planilha_path = script_dir / args.planilha

    pdfs_dir = Path(args.pasta_pdfs)
    if not pdfs_dir.is_absolute() and not pdfs_dir.exists():
        pdfs_dir = script_dir / args.pasta_pdfs

    estado = EstadoGUI(planilha_path=planilha_path, pdfs_dir=pdfs_dir, script_dir=script_dir)
    ManipuladorHTTPGUI.estado = estado

    porta = escolher_porta_livre(args.porta)
    url = f"http://{HOST_LOCAL}:{porta}"

    servidor = ThreadingHTTPServer((HOST_LOCAL, porta), ManipuladorHTTPGUI)

    print("=" * 68)
    print(f"  RENOMEADOR DE PDFs v5.0 (Interface Local)")
    print("=" * 68)
    print(f"  Endereco local : {url}")
    print(f"  Planilha ativa : {planilha_path.name} ({len(estado.itens_planilha)} nomes)")
    print(f"  Pasta de PDFs  : {pdfs_dir} ({len(estado.fila_pdfs)} pendentes)")
    print("-" * 68)
    print("  Pressione Ctrl+C para encerrar.")
    print("=" * 68)

    if not args.nao_abrir_navegador:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()

    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nEncerrando servidor local...")
    finally:
        servidor.server_close()


if __name__ == "__main__":
    main()
