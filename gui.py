# -*- coding: utf-8 -*-
"""
Renomeador Semiautomático de PDFs - Interface Gráfica Web Local (v5.0 GUI)
--------------------------------------------------------------------------
- Roda 100% local em http://127.0.0.1 (nenhum dado sai do seu computador).
- Não exige permissões de administrador nem frameworks pesados (usa a biblioteca
  padrão do Python `http.server` + motor inteligente do `renomear.py`).
- Visualizador de PDF embutido na própria página ao lado dos comandos e da fila.
- Suporte completo a atalhos de teclado (S/Enter, E, B, D, N, P, Z) e cliques rápidos.
- Permite colar dados direto do Excel (Ctrl+C / Ctrl+V) ou enviar planilha .xlsx/.csv.
"""

import argparse
import io
import json
import mimetypes
import os
import shutil
import socket
import sys
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


# ==============================================================================
# ESTADO DA SESSÃO DA GUI
# ==============================================================================

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
        self.ultimo_evento: str = "Sessão iniciada. Pronto para renomear!"
        self.erro_carregamento: str = ""

        # Cache leve de número de páginas por arquivo
        self._cache_paginas: Dict[str, Optional[int]] = {}

        self.recarregar_tudo()

    def recarregar_tudo(self):
        """Recarrega planilha e sincroniza com os PDFs da pasta."""
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
        """Reescaneia a pasta de PDFs preservando o andamento atual."""
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
            "fila_pdfs_nomes": [
                {
                    "idx": idx,
                    "nome": p.name,
                    "atual": idx == self.idx_pdf,
                }
                for idx, p in enumerate(self.fila_pdfs[self.idx_pdf : self.idx_pdf + 30], start=self.idx_pdf)
            ],
        }


# ==============================================================================
# AÇÕES DE NEGÓCIO DA GUI
# ==============================================================================

def acao_confirmar_atual(estado: EstadoGUI, idx_escolhido: Optional[int] = None) -> Tuple[bool, str]:
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None or not pdf_cur.exists():
        return False, "Nenhum arquivo PDF pendente selecionado."

    alvo_idx = idx_escolhido if idx_escolhido is not None else estado.idx_nome
    if alvo_idx is None or not (0 <= alvo_idx < len(estado.itens_planilha)):
        return False, "Não há nome selecionado na fila. Use [E] Editar para digitar um nome manual."

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
        # Foi escolhido fora de ordem via busca/clique na lista: manter ponteiro atual se ainda pendente
        if estado.idx_nome is None or estado.itens_planilha[estado.idx_nome].concluido:
            estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, 0)

    estado.ultimo_evento = f"✅ '{pdf_cur.name}' renomeado para '{destino.name}'"
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
        return False, "Digite um nome válido para o arquivo."

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

    modo_txt = "fila avançou" if avancar_planilha else "pessoa extra - fila mantida"
    estado.ultimo_evento = f"✏️ '{pdf_cur.name}' -> '{destino.name}' ({modo_txt})"
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_pular_nome(estado: EstadoGUI) -> Tuple[bool, str]:
    item_cur = estado.item_atual()
    if item_cur is None or estado.idx_nome is None:
        return False, "Não há mais nomes pendentes na fila para pular."

    nome_antigo = item_cur.nome_formatado
    prox = core.proximo_indice_nome_pendente(estado.itens_planilha, estado.idx_nome + 1)
    if prox == estado.idx_nome:
        return False, "Este é o único nome restante na fila."

    estado.historico.append(
        core.AcaoHistorico(
            tipo="pular_nome",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=estado.idx_nome,
            descricao=f"Pulou nome '{nome_antigo}'",
        )
    )
    estado.idx_nome = prox
    estado.ultimo_evento = f"⏭️ Nome '{nome_antigo}' pulado (mantido o mesmo PDF)"
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
            descricao=f"Pulou PDF '{pdf_cur.name}'",
        )
    )
    estado.total_pulados += 1
    estado.idx_pdf += 1
    if avancar_nome_junto and estado.idx_nome is not None:
        estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, estado.idx_nome + 1)

    estado.ultimo_evento = f"⏭️ PDF '{pdf_cur.name}' pulado"
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_dividir_pdf(
    estado: EstadoGUI,
    modo: str,
    pagina_corte: int = 1,
    ja_renomear_parte1: bool = True,
) -> Tuple[bool, str]:
    """
    Divide ou duplica o PDF atual:
    - modo == 'duplicar': duplica o PDF inteiro em _parte1 e _parte2
    - modo == 'todas': separa cada página em um PDF individual
    - modo == 'corte': separa Páginas 1..pagina_corte (Parte 1) e (pagina_corte+1)..N (Parte 2)
    """
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
            # Tentar instalar pypdf automaticamente se falhou por falta de biblioteca
            if "pip install pypdf" in status and core._instalar_pypdf_usuario():
                novos_pdfs, status = core.dividir_pdf_em_intervalos(pdf_cur, intervalos)
            if not novos_pdfs:
                return False, status

        try:
            shutil.move(str(pdf_cur), str(backup_path))
        except Exception:
            pass

    # Substituir na fila de PDFs
    estado.fila_pdfs[estado.idx_pdf : estado.idx_pdf + 1] = novos_pdfs
    estado.historico.append(
        core.AcaoHistorico(
            tipo="dividir",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=estado.idx_nome if estado.idx_nome is not None else 0,
            pdf_origem=pdf_cur,
            arquivos_criados=list(novos_pdfs),
            backup_origem=backup_path,
            descricao=f"Dividido '{pdf_cur.name}' em {len(novos_pdfs)} partes",
        )
    )
    estado.total_divididos += 1

    if ja_renomear_parte1 and estado.item_atual() is not None:
        ok, msg_ren = acao_confirmar_atual(estado)
        if ok:
            estado.ultimo_evento = (
                f"✂️ PDF dividido em {len(novos_pdfs)} partes e Parte 1 já renomeada! "
                f"Exibindo Parte 2 agora."
            )
            return True, estado.ultimo_evento

    estado.ultimo_evento = f"✂️ '{pdf_cur.name}' dividido em {len(novos_pdfs)} partes ({', '.join(p.name for p in novos_pdfs)})"
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_desfazer(estado: EstadoGUI) -> Tuple[bool, str]:
    novo_idx_pdf, novo_idx_nome, msg = core.desfazer_ultima_acao(
        estado.historico, estado.fila_pdfs, estado.itens_planilha, estado.log_path
    )
    # Limpar códigos ANSI da mensagem do core para exibir bonito na web
    msg_limpa = core.re.sub(r"\033\[[0-9;]*m", "", msg)
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
    """
    Recebe texto copiado direto do Excel/Planilhas (Ctrl+C -> Ctrl+V),
    formata todos os nomes e salva em `nomes.csv` (ou atualiza a sessão).
    """
    linhas_brutas = [l.strip() for l in texto_colado.splitlines() if l.strip()]
    if not linhas_brutas:
        return False, "O texto colado está vazio."

    # Salvar em um arquivo CSV/TSV temporário ou atualizar planilha ativa
    csv_destino = estado.script_dir / "nomes_colados.csv"
    csv_destino.write_text(texto_colado, encoding="utf-8")

    itens, msg = core.carregar_planilha(csv_destino)
    if not itens:
        return False, f"Não foi possível extrair nomes do texto colado: {msg}"

    estado.planilha_path = csv_destino
    estado.recarregar_tudo()
    estado.ultimo_evento = f"📋 {len(itens)} nomes carregados via colagem rápida!"
    return True, estado.ultimo_evento


# ==============================================================================
# INTERFACE WEB SINGLE-PAGE (HTML5 + CSS3 + JS PURO - ZERO DEPENDÊNCIAS EXTERNAS)
# ==============================================================================

PAGINA_HTML = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Renomeador de PDFs v5.0 — Localhost GUI</title>
<style>
  :root {
    --bg-main: #0f172a;
    --bg-panel: #1e293b;
    --bg-card: #263349;
    --bg-input: #0b1120;
    --border: #334155;
    --text: #f8fafc;
    --text-muted: #94a3b8;
    --green: #10b981;
    --green-hover: #059669;
    --amber: #f59e0b;
    --amber-hover: #d97706;
    --purple: #8b5cf6;
    --purple-hover: #7c3aed;
    --blue: #3b82f6;
    --blue-hover: #2563eb;
    --red: #ef4444;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Inter, sans-serif;
    background: var(--bg-main);
    color: var(--text);
    height: 100vh;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    user-select: none;
  }

  /* BARRA SUPERIOR */
  header {
    height: 54px;
    background: var(--bg-panel);
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 18px;
    flex-shrink: 0;
  }
  .brand {
    display: flex;
    align-items: center;
    gap: 10px;
    font-weight: 700;
    font-size: 16px;
  }
  .badge-local {
    background: rgba(16, 185, 129, 0.15);
    color: #34d399;
    border: 1px solid rgba(16, 185, 129, 0.35);
    font-size: 11px;
    padding: 3px 8px;
    border-radius: 999px;
    font-weight: 600;
  }
  .top-stats {
    display: flex;
    align-items: center;
    gap: 16px;
    font-size: 13px;
  }
  .progress-wrap {
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .progress-bar-bg {
    width: 160px;
    height: 10px;
    background: var(--bg-input);
    border-radius: 999px;
    overflow: hidden;
    border: 1px solid var(--border);
  }
  .progress-bar-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--blue), var(--green));
    width: 0%;
    transition: width 0.2s ease;
  }
  .top-actions {
    display: flex;
    gap: 8px;
  }
  .btn-top {
    background: var(--bg-card);
    color: var(--text);
    border: 1px solid var(--border);
    padding: 6px 12px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.15s;
  }
  .btn-top:hover { background: #334155; }

  /* LAYOUT DIVIDIDO PRINCIPAL */
  .workspace {
    flex: 1;
    display: flex;
    overflow: hidden;
  }

  /* PAINEL ESQUERDO: VISUALIZADOR DE PDF EMBUTIDO */
  .pdf-pane {
    flex: 1.55;
    display: flex;
    flex-direction: column;
    border-right: 1px solid var(--border);
    background: #090d16;
    min-width: 0;
  }
  .pdf-toolbar {
    min-height: 44px;
    background: var(--bg-card);
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 6px 14px;
    gap: 10px;
    flex-wrap: wrap;
  }
  .pdf-file-info {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 14px;
    font-weight: 600;
  }
  .pdf-filename {
    color: #fbbf24;
    font-family: monospace;
    font-size: 15px;
    background: rgba(245, 158, 11, 0.12);
    padding: 3px 8px;
    border-radius: 5px;
    border: 1px solid rgba(245, 158, 11, 0.3);
  }
  .page-badge {
    font-size: 12px;
    padding: 3px 9px;
    border-radius: 6px;
    font-weight: 700;
    background: #334155;
    color: #e2e8f0;
  }
  .page-badge.multi-page {
    background: rgba(245, 158, 11, 0.22);
    color: #fcd34d;
    border: 1px solid #f59e0b;
    cursor: pointer;
    animation: pulse-border 1.6s infinite;
  }
  @keyframes pulse-border {
    0%, 100% { box-shadow: 0 0 0 0 rgba(245, 158, 11, 0.4); }
    50% { box-shadow: 0 0 0 5px rgba(245, 158, 11, 0); }
  }
  .pdf-container {
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
    background: #1e293b;
  }
  .empty-state {
    text-align: center;
    padding: 40px;
    max-width: 480px;
  }
  .empty-state h2 { margin-bottom: 10px; color: #34d399; }
  .empty-state p { color: var(--text-muted); line-height: 1.5; margin-bottom: 18px; }

  /* PAINEL DIREITO: CONTROLES, DADOS E FILA */
  .control-pane {
    width: 470px;
    max-width: 46vw;
    display: flex;
    flex-direction: column;
    background: var(--bg-panel);
    flex-shrink: 0;
  }

  /* CARTÃO DO ALUNO ATUAL */
  .target-card {
    padding: 16px;
    background: var(--bg-card);
    border-bottom: 1px solid var(--border);
  }
  .target-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--text-muted);
    margin-bottom: 8px;
    font-weight: 700;
  }
  .student-fullname {
    font-size: 20px;
    font-weight: 800;
    color: #ffffff;
    line-height: 1.25;
    margin-bottom: 6px;
    user-select: text;
  }
  .student-mat {
    font-size: 14px;
    color: #93c5fd;
    font-family: monospace;
    margin-bottom: 10px;
    user-select: text;
  }
  .formatted-preview-box {
    background: var(--bg-input);
    border: 1px solid #10b981;
    border-radius: 8px;
    padding: 9px 12px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
  }
  .formatted-filename {
    font-family: monospace;
    font-size: 14.5px;
    font-weight: 700;
    color: #34d399;
    word-break: break-all;
    user-select: text;
  }

  /* BOTÕES DE AÇÃO RÁPIDA */
  .actions-section {
    padding: 14px 16px;
    border-bottom: 1px solid var(--border);
    display: flex;
    flex-direction: column;
    gap: 9px;
  }
  .btn-main-confirm {
    width: 100%;
    padding: 14px;
    border: none;
    border-radius: 8px;
    background: var(--green);
    color: #052e16;
    font-size: 16px;
    font-weight: 800;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    transition: transform 0.08s, background 0.15s;
    box-shadow: 0 4px 12px rgba(16, 185, 129, 0.25);
  }
  .btn-main-confirm:hover { background: #34d399; }
  .btn-main-confirm:active { transform: scale(0.99); }

  .action-grid-2 {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 8px;
  }
  .action-grid-3 {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 8px;
  }
  .btn-action {
    padding: 10px 8px;
    border-radius: 7px;
    border: 1px solid var(--border);
    background: var(--bg-card);
    color: var(--text);
    font-size: 13px;
    font-weight: 700;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    transition: background 0.15s, border-color 0.15s;
  }
  .btn-action:hover { background: #334155; border-color: #64748b; }
  .btn-action.edit { border-color: rgba(245, 158, 11, 0.5); color: #fcd34d; }
  .btn-action.split { border-color: rgba(139, 92, 246, 0.5); color: #c4b5fd; }
  .btn-action.undo { border-color: rgba(59, 130, 246, 0.5); color: #93c5fd; }
  .kbd {
    background: rgba(15, 23, 42, 0.65);
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 4px;
    padding: 1px 5px;
    font-family: monospace;
    font-size: 11px;
  }

  /* PAINEL EXPANSÍVEL (EDITAR OU DIVIDIR) */
  .drawer {
    display: none;
    background: #172033;
    border-bottom: 2px solid var(--amber);
    padding: 14px 16px;
  }
  .drawer.open { display: block; }
  .drawer h4 {
    font-size: 13px;
    margin-bottom: 8px;
    color: #fcd34d;
    display: flex;
    justify-content: space-between;
  }
  .input-field {
    width: 100%;
    padding: 9px 11px;
    background: var(--bg-input);
    border: 1px solid var(--border);
    border-radius: 6px;
    color: var(--text);
    font-size: 14px;
    margin-bottom: 8px;
  }
  .input-field:focus { outline: none; border-color: var(--blue); }
  .radio-row {
    display: flex;
    flex-direction: column;
    gap: 6px;
    font-size: 12.5px;
    margin: 8px 0 10px;
    color: #cbd5e1;
  }
  .radio-row label {
    display: flex;
    align-items: center;
    gap: 7px;
    cursor: pointer;
  }

  /* ABAS DA FILA E PLANILHA */
  .tabs-bar {
    display: flex;
    border-bottom: 1px solid var(--border);
    background: var(--bg-main);
  }
  .tab-btn {
    flex: 1;
    padding: 10px 6px;
    background: transparent;
    border: none;
    border-bottom: 2px solid transparent;
    color: var(--text-muted);
    font-size: 12.5px;
    font-weight: 700;
    cursor: pointer;
  }
  .tab-btn.active {
    color: var(--text);
    border-bottom-color: var(--blue);
    background: var(--bg-panel);
  }

  .tab-content {
    flex: 1;
    display: none;
    flex-direction: column;
    overflow: hidden;
    padding: 12px 14px;
  }
  .tab-content.active { display: flex; }

  .search-bar {
    margin-bottom: 8px;
  }
  .queue-list {
    flex: 1;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 6px;
    padding-right: 4px;
  }
  .queue-item {
    padding: 8px 10px;
    border-radius: 7px;
    background: var(--bg-card);
    border: 1px solid var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    font-size: 12.5px;
  }
  .queue-item.current {
    border-color: var(--green);
    background: rgba(16, 185, 129, 0.12);
  }
  .queue-item.done {
    opacity: 0.48;
  }
  .queue-item-info {
    min-width: 0;
    flex: 1;
  }
  .queue-item-title {
    font-weight: 700;
    color: #f1f5f9;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .queue-item-sub {
    font-size: 11px;
    color: var(--text-muted);
    font-family: monospace;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .queue-item-btns {
    display: flex;
    gap: 5px;
    flex-shrink: 0;
  }
  .btn-mini {
    padding: 4px 8px;
    border-radius: 5px;
    border: 1px solid var(--border);
    background: var(--bg-main);
    color: var(--text);
    font-size: 11px;
    font-weight: 600;
    cursor: pointer;
  }
  .btn-mini:hover { background: var(--blue); border-color: var(--blue); }

  /* RODAPÉ DE STATUS */
  footer {
    height: 32px;
    background: var(--bg-input);
    border-top: 1px solid var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 16px;
    font-size: 12px;
    color: var(--text-muted);
    flex-shrink: 0;
  }
  #status-toast {
    color: #34d399;
    font-weight: 600;
  }
</style>
</head>
<body>

<header>
  <div class="brand">
    <span>📄 Renomeador de PDFs v5.0</span>
    <span class="badge-local">🔒 100% Localhost (Dados Privados)</span>
  </div>

  <div class="top-stats">
    <div class="progress-wrap">
      <span id="lbl-progresso">0/0 (0%)</span>
      <div class="progress-bar-bg">
        <div class="progress-bar-fill" id="barra-progresso"></div>
      </div>
    </div>
    <span>✅ Sessão: <strong id="lbl-renomeados">0</strong></span>
    <span>✂️ Divididos: <strong id="lbl-divididos">0</strong></span>
  </div>

  <div class="top-actions">
    <button class="btn-top" onclick="abrirAba('aba-planilha')">📋 Colar do Excel / Planilha</button>
    <button class="btn-top" onclick="reescanearPasta()">🔄 Atualizar Pasta PDFs</button>
  </div>
</header>

<div class="workspace">
  <!-- ESQUERDA: PDF EMBUTIDO -->
  <section class="pdf-pane">
    <div class="pdf-toolbar">
      <div class="pdf-file-info">
        <span>Arquivo:</span>
        <span class="pdf-filename" id="lbl-pdf-nome">—</span>
        <span id="lbl-pdf-pos" style="color: var(--text-muted); font-size: 12px;">(0/0)</span>
        <span class="page-badge" id="badge-paginas" onclick="abrirGavetaDividir()">—</span>
      </div>
      <div style="display:flex; gap:8px;">
        <button class="btn-top" onclick="recarregarIframePdf()" title="Recarregar visualização do PDF">🔄 Recarregar PDF</button>
        <button class="btn-top" onclick="focarAtalhos()" title="Clique aqui se os atalhos de 1 tecla pararem após clicar dentro do PDF">⌨️ Ativar Teclas Rápidas</button>
      </div>
    </div>

    <div class="pdf-container">
      <iframe id="pdf-frame" title="Visualizador de PDF"></iframe>
      <div class="empty-state" id="empty-state" style="display:none;">
        <h2>🎉 Nenhum PDF Pendente!</h2>
        <p id="empty-msg">Todos os arquivos PDF da pasta <code>pdfs/</code> já foram renomeados ou a pasta está vazia.</p>
        <button class="btn-main-confirm" onclick="reescanearPasta()">🔄 Verificar Pasta Novamente</button>
      </div>
    </div>
  </section>

  <!-- DIREITA: COMANDOS E FILA -->
  <aside class="control-pane">
    <!-- CARTÃO DO ALUNO SUGERIDO -->
    <div class="target-card">
      <div class="target-header">
        <span>👤 Nome Esperado na Fila</span>
        <span id="lbl-item-idx">Item #—</span>
      </div>
      <div class="student-fullname" id="lbl-aluno-nome">—</div>
      <div class="student-mat" id="lbl-aluno-mat">Matrícula: —</div>
      <div class="formatted-preview-box">
        <span class="formatted-filename" id="lbl-aluno-formatado">—.pdf</span>
      </div>
    </div>

    <!-- BOTÕES DE AÇÃO -->
    <div class="actions-section">
      <button class="btn-main-confirm" id="btn-confirmar" onclick="confirmarAtual()">
        <span>✅ SIM, ESTÁ CORRETO — RENOMEAR</span>
        <span class="kbd">S / Enter</span>
      </button>

      <div class="action-grid-2">
        <button class="btn-action edit" onclick="abrirGavetaEditar()">
          <span>✏️ Editar / Pessoa Extra</span>
          <span class="kbd">E</span>
        </button>
        <button class="btn-action split" onclick="abrirGavetaDividir()">
          <span>✂️ Dividir PDF (2+ p.)</span>
          <span class="kbd">D</span>
        </button>
      </div>

      <div class="action-grid-3">
        <button class="btn-action" onclick="pularNome()" title="A pessoa da planilha faltou (sem PDF). Avança só o nome e mantém este PDF.">
          <span>⏭️ Pular Nome</span>
          <span class="kbd">N</span>
        </button>
        <button class="btn-action" onclick="pularPdf()" title="Deixa este arquivo PDF para depois, mantendo o nome atual para o próximo PDF.">
          <span>⏭️ Pular PDF</span>
          <span class="kbd">P</span>
        </button>
        <button class="btn-action undo" onclick="desfazerUltimo()" id="btn-undo" title="Desfaz imediatamente a última renomeação ou divisão">
          <span>↩️ Desfazer</span>
          <span class="kbd">Z</span>
        </button>
      </div>
    </div>

    <!-- GAVETA 1: EDITAR / PESSOA FORA DA PLANILHA -->
    <div class="drawer" id="gaveta-editar">
      <h4>
        <span>✏️ Digitar Nome Manual / Pessoa Fora da Planilha</span>
        <span style="cursor:pointer; color:var(--text-muted);" onclick="fecharGavetas()">✕ [Esc]</span>
      </h4>
      <input type="text" id="inp-edit-nome" class="input-field"
             placeholder="Ex: Maria da Silva 20201011100005 (formata sozinho!)"
             oninput="atualizarPreviewEdicao()"
             onkeydown="if(event.key==='Enter'){salvarEdicaoManual();} else if(event.key==='Escape'){fecharGavetas();}">
      <div style="font-size:12px; color:#34d399; font-family:monospace; margin-bottom:6px;" id="preview-edit-formatado">
        Prévia: —
      </div>
      <div class="radio-row">
        <label>
          <input type="radio" name="modo-edit" value="extra" checked>
          <span><strong>Pessoa EXTRA</strong> (não está na planilha — manter fila atual para o próximo PDF)</span>
        </label>
        <label>
          <input type="radio" name="modo-edit" value="substituir">
          <span><strong>Substituir</strong> o nome atual da fila (avançar a fila da planilha)</span>
        </label>
      </div>
      <div class="action-grid-2">
        <button class="btn-main-confirm" style="padding:9px; font-size:13px;" onclick="salvarEdicaoManual()">✅ Renomear Agora [Enter]</button>
        <button class="btn-action" onclick="fecharGavetas()">Cancelar [Esc]</button>
      </div>
    </div>

    <!-- GAVETA 2: DIVIDIR PDF COM 2+ PESSOAS -->
    <div class="drawer" id="gaveta-dividir" style="border-bottom-color: var(--purple);">
      <h4 style="color:#c4b5fd;">
        <span>✂️ Dividir PDF com 2 ou mais Pessoas</span>
        <span style="cursor:pointer; color:var(--text-muted);" onclick="fecharGavetas()">✕ [Esc]</span>
      </h4>
      <p style="font-size:12px; color:var(--text-muted); margin-bottom:10px;" id="lbl-info-divisao">
        Separe as páginas deste PDF na hora sem sair do navegador:
      </p>
      <div style="display:flex; flex-direction:column; gap:7px;">
        <button class="btn-main-confirm" style="padding:10px; font-size:13px; background:var(--purple); color:#fff;"
                onclick="executarDivisao('corte', 1, true)">
          ⚡ Separar Pág. 1 (renomear p/ atual) + Restante (próximo da fila)
        </button>
        <div style="display:flex; gap:6px; align-items:center;">
          <span style="font-size:12px;">1ª pessoa vai até a pág.:</span>
          <input type="number" id="inp-corte-pag" value="1" min="1" max="99" class="input-field" style="width:65px; margin:0; padding:6px;">
          <button class="btn-action" style="flex:1; padding:7px;" onclick="executarDivisaoCustom()">✂️ Cortar nesta página</button>
        </div>
        <div class="action-grid-2">
          <button class="btn-action" onclick="executarDivisao('todas', 1, false)">📄 1 PDF por página</button>
          <button class="btn-action" onclick="executarDivisao('duplicar', 1, true)">📑 Duplicar PDF (mesma pág.)</button>
        </div>
      </div>
    </div>

    <!-- ABAS INFERIORES: BUSCA NA PLANILHA / COLAR EXCEL / FILA DE PDFS -->
    <div class="tabs-bar">
      <button class="tab-btn active" id="btn-aba-nomes" onclick="abrirAba('aba-nomes')">🔎 Fila / Busca [B]</button>
      <button class="tab-btn" id="btn-aba-planilha" onclick="abrirAba('aba-planilha')">📋 Colar do Excel</button>
      <button class="tab-btn" id="btn-aba-pdfs" onclick="abrirAba('aba-pdfs')">📂 Arquivos PDF</button>
    </div>

    <!-- CONTEÚDO ABA 1: FILA DE NOMES + BUSCA FORA DE ORDEM -->
    <div class="tab-content active" id="aba-nomes">
      <div class="search-bar">
        <input type="text" id="inp-busca" class="input-field" style="margin-bottom:0;"
               placeholder="🔎 Buscar nome ou matrícula fora de ordem (Tecla B)..."
               oninput="renderizarListaNomes()"
               onkeydown="if(event.key==='Escape'){this.value=''; renderizarListaNomes(); this.blur();}">
      </div>
      <div class="queue-list" id="lista-nomes"></div>
    </div>

    <!-- CONTEÚDO ABA 2: COLAR DIRETO DO EXCEL -->
    <div class="tab-content" id="aba-planilha">
      <p style="font-size:12px; color:var(--text-muted); margin-bottom:8px; line-height:1.4;">
        Copie as <strong>2 colunas (Nome Completo e Matrícula)</strong> direto do Excel (<code>Ctrl+C</code>) e cole abaixo (<code>Ctrl+V</code>). O sistema já remove <code>de/da/do/dos</code> e formata tudo sozinho!
      </p>
      <textarea id="txt-colar-excel" class="input-field" style="flex:1; font-family:monospace; font-size:12px; resize:none;"
                placeholder="Gustavo Ferreira Jácome&#9;20201011100005&#10;Maria da Silva Santos&#9;20252014040028"></textarea>
      <div class="action-grid-2">
        <button class="btn-main-confirm" style="padding:10px; font-size:13px;" onclick="enviarPlanilhaColada()">📥 Carregar Nomes Colados</button>
        <button class="btn-action" onclick="recarregarPlanilhaDisco()">🔄 Recarregar nomes.xlsx</button>
      </div>
    </div>

    <!-- CONTEÚDO ABA 3: FILA DE ARQUIVOS PDF -->
    <div class="tab-content" id="aba-pdfs">
      <div class="queue-list" id="lista-pdfs"></div>
    </div>
  </aside>
</div>

<footer>
  <span id="status-toast">Iniciando...</span>
  <span>Atalhos: <strong>[S/Enter]</strong> Confirmar • <strong>[E]</strong> Editar • <strong>[D]</strong> Dividir • <strong>[B]</strong> Buscar • <strong>[N]</strong> Pular Nome • <strong>[P]</strong> Pular PDF • <strong>[Z]</strong> Desfazer</span>
</footer>

<link rel="prefetch" id="pdf-prefetch" href="">

<script>
let ESTADO = null;
let urlPdfAtualNoIframe = "";
const PARTICULAS = new Set(["de", "da", "do", "dos", "das", "e", "del", "d", "d'"]);

// Formatação instantânea no navegador (espelha a regra do Python para preview ao vivo)
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
  document.getElementById("preview-edit-formatado").textContent = "Como vai ficar: " + formatarPreviewLocal(val);
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
      mostrarStatus("⚠️ " + json.mensagem, true);
    }
    return json;
  } catch (err) {
    mostrarStatus("❌ Erro de comunicação com o servidor local: " + err, true);
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
  el.style.color = erro ? "#f87171" : "#34d399";
}

function aplicarEstado(st) {
  ESTADO = st;
  const p = st.progresso;
  document.getElementById("lbl-progresso").textContent =
    `${p.concluidos_planilha}/${p.total_planilha} (${p.percentual}%)`;
  document.getElementById("barra-progresso").style.width = `${p.percentual}%`;
  document.getElementById("lbl-renomeados").textContent = p.renomeados_sessao;
  document.getElementById("lbl-divididos").textContent = p.divididos_sessao;
  document.getElementById("btn-undo").disabled = !p.pode_desfazer;

  mostrarStatus(st.ultimo_evento || "", Boolean(st.erro_carregamento));

  const iframe = document.getElementById("pdf-frame");
  const empty = document.getElementById("empty-state");

  if (st.tem_pdf && st.pdf_atual) {
    iframe.style.display = "block";
    empty.style.display = "none";
    document.getElementById("lbl-pdf-nome").textContent = st.pdf_atual.nome;
    document.getElementById("lbl-pdf-pos").textContent =
      `(${st.pdf_atual.posicao} de ${st.pdf_atual.total_fila})`;

    const badge = document.getElementById("badge-paginas");
    const pags = st.pdf_atual.paginas;
    if (pags && pags >= 2) {
      badge.textContent = `⚠️ ${pags} páginas (2+ pessoas? Clique ou [D] p/ dividir)`;
      badge.className = "page-badge multi-page";
    } else if (pags === 1) {
      badge.textContent = "1 página";
      badge.className = "page-badge";
    } else {
      badge.textContent = "PDF";
      badge.className = "page-badge";
    }

    const novaUrl = st.pdf_atual.url + "#toolbar=1&navpanes=0&view=FitH";
    if (urlPdfAtualNoIframe !== novaUrl) {
      urlPdfAtualNoIframe = novaUrl;
      iframe.src = novaUrl;
    }

    // Pré-carregar o próximo PDF em background para transição instantânea
    const prefetch = document.getElementById("pdf-prefetch");
    prefetch.href = st.pdf_proximo_url || "";
  } else {
    iframe.style.display = "none";
    empty.style.display = "block";
    urlPdfAtualNoIframe = "";
    iframe.src = "about:blank";
    document.getElementById("lbl-pdf-nome").textContent = "Concluído";
    document.getElementById("lbl-pdf-pos").textContent = "";
    document.getElementById("badge-paginas").textContent = "—";
    document.getElementById("badge-paginas").className = "page-badge";
    if (st.erro_carregamento) {
      document.getElementById("empty-msg").textContent = st.erro_carregamento;
    }
  }

  // Atualizar cartão do aluno esperado
  if (st.item_atual) {
    document.getElementById("lbl-item-idx").textContent =
      `Item #${st.item_atual.indice_original} de ${p.total_planilha}`;
    document.getElementById("lbl-aluno-nome").textContent = st.item_atual.nome_completo;
    document.getElementById("lbl-aluno-mat").textContent =
      st.item_atual.matricula ? `Matrícula: ${st.item_atual.matricula}` : "Sem matrícula separada";
    document.getElementById("lbl-aluno-formatado").textContent =
      `${st.item_atual.nome_formatado}.pdf`;
  } else {
    document.getElementById("lbl-item-idx").textContent = "Fila concluída";
    document.getElementById("lbl-aluno-nome").textContent = "Todos os nomes da planilha foram usados!";
    document.getElementById("lbl-aluno-mat").textContent = "Use [E] para digitar um nome extra ou recarregue a planilha.";
    document.getElementById("lbl-aluno-formatado").textContent = "—";
  }

  renderizarListaNomes();
  renderizarListaPdfs();
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
    const cls = it.atual ? "queue-item current" : (it.concluido ? "queue-item done" : "queue-item");
    return `
      <div class="${cls}">
        <div class="queue-item-info">
          <div class="queue-item-title">#${it.indice_original} ${it.nome_completo}</div>
          <div class="queue-item-sub">${it.nome_formatado}.pdf</div>
        </div>
        <div class="queue-item-btns">
          <button class="btn-mini" onclick="confirmarComIndice(${it.idx})" title="Renomear o PDF aberto agora para esta pessoa">
            🎯 Usar neste PDF
          </button>
          ${!it.concluido && !it.atual ? `
            <button class="btn-mini" onclick="apontarFila(${it.idx})" title="Mover o ponteiro da fila para começar desta pessoa">
              👁️ Fila aqui
            </button>
          ` : ""}
        </div>
      </div>
    `;
  }).join("");
}

function renderizarListaPdfs() {
  if (!ESTADO) return;
  const container = document.getElementById("lista-pdfs");
  container.innerHTML = (ESTADO.fila_pdfs_nomes || []).map(p => `
    <div class="queue-item ${p.atual ? 'current' : ''}">
      <div class="queue-item-info">
        <div class="queue-item-title">📄 ${p.nome}</div>
      </div>
      <div class="queue-item-btns">
        ${!p.atual ? `<button class="btn-mini" onclick="apontarPdf(${p.idx})">Abrir este</button>` : `<span>Atual</span>`}
      </div>
    </div>
  `).join("");
}

function fecharGavetas() {
  document.getElementById("gaveta-editar").classList.remove("open");
  document.getElementById("gaveta-dividir").classList.remove("open");
}

function abrirGavetaEditar() {
  fecharGavetas();
  const gaveta = document.getElementById("gaveta-editar");
  gaveta.classList.add("open");
  const inp = document.getElementById("inp-edit-nome");
  inp.value = "";
  atualizarPreviewEdicao();
  setTimeout(() => inp.focus(), 30);
}

function abrirGavetaDividir() {
  fecharGavetas();
  const gaveta = document.getElementById("gaveta-dividir");
  gaveta.classList.add("open");
  if (ESTADO && ESTADO.pdf_atual && ESTADO.pdf_atual.paginas) {
    document.getElementById("lbl-info-divisao").textContent =
      `O arquivo '${ESTADO.pdf_atual.nome}' possui ${ESTADO.pdf_atual.paginas} página(s). Como deseja dividir?`;
  }
}

function abrirAba(idAba) {
  for (const id of ["aba-nomes", "aba-planilha", "aba-pdfs"]) {
    document.getElementById(id).classList.toggle("active", id === idAba);
    document.getElementById("btn-" + id).classList.toggle("active", id === idAba);
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

async function apontarPdf(idx) {
  await apiPost("/api/select_pdf", { idx_pdf: idx });
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
  await apiPost("/api/reload_sheet", {});
}

async function enviarPlanilhaColada() {
  const texto = document.getElementById("txt-colar-excel").value;
  const res = await apiPost("/api/paste_sheet", { texto: texto });
  if (res && res.ok) {
    abrirAba("aba-nomes");
  }
}

function recarregarIframePdf() {
  const iframe = document.getElementById("pdf-frame");
  if (iframe && urlPdfAtualNoIframe) {
    iframe.src = urlPdfAtualNoIframe;
  }
}

function focarAtalhos() {
  window.focus();
  if (document.activeElement && document.activeElement.id === "pdf-frame") {
    document.activeElement.blur();
  }
  mostrarStatus("⌨️ Atalhos de 1 tecla prontos! (S/Enter, E, D, B, N, P, Z)");
}

// Mantém os atalhos de 1 tecla sempre ativos mesmo se clicar no PDF para rolar a página
setInterval(() => {
  if (document.activeElement && document.activeElement.id === "pdf-frame") {
    document.activeElement.blur();
    window.focus();
  }
}, 250);

// Atalhos globais de 1 tecla (quando não estiver digitando em um campo de texto)
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
    abrirGavetaEditar();
  } else if (k === "d") {
    e.preventDefault();
    abrirGavetaDividir();
  } else if (k === "b" || k === "/") {
    e.preventDefault();
    abrirAba("aba-nomes");
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


# ==============================================================================
# SERVIDOR HTTP LOCAL (127.0.0.1 APENAS)
# ==============================================================================

class ManipuladorHTTPGUI(BaseHTTPRequestHandler):
    estado: EstadoGUI = None  # type: ignore

    def log_message(self, format, *args):
        # Silenciar logs repetitivos no terminal para manter limpo
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
            # Proteção contra path traversal: usar apenas o nome do arquivo dentro de pdfs_dir
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
                msg = f"🔄 Pasta atualizada: {len(self.estado.fila_pdfs)} PDF(s) pendente(s)."

            elif rota == "/api/reload_sheet":
                self.estado.planilha_path = self.estado.script_dir / core.NOME_PLANILHA_PADRAO
                self.estado.recarregar_tudo()
                msg = f"🔄 Planilha '{self.estado.planilha_path.name}' recarregada!"

            elif rota == "/api/paste_sheet":
                ok, msg = acao_colar_planilha_texto(
                    self.estado, str(payload.get("texto", ""))
                )

            elif rota == "/api/select_name":
                idx_n = payload.get("idx_nome")
                if idx_n is not None and 0 <= int(idx_n) < len(self.estado.itens_planilha):
                    self.estado.idx_nome = int(idx_n)
                    self.estado.ultimo_evento = (
                        f"👁️ Fila apontada para '{self.estado.itens_planilha[self.estado.idx_nome].nome_formatado}'"
                    )
                    msg = self.estado.ultimo_evento

            elif rota == "/api/select_pdf":
                idx_p = payload.get("idx_pdf")
                if idx_p is not None and 0 <= int(idx_p) < len(self.estado.fila_pdfs):
                    self.estado.idx_pdf = int(idx_p)
                    self.estado.ultimo_evento = (
                        f"📂 Exibindo PDF '{self.estado.fila_pdfs[self.estado.idx_pdf].name}'"
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

    print("=" * 72)
    print(f"{core.Cor.BOLD}{core.Cor.VERDE}  🌐 RENOMEADOR DE PDFs v5.0 — INTERFACE GRÁFICA LOCAL (GUI){core.Cor.RESET}")
    print("=" * 72)
    print(f"  🔒 Servidor 100% local (offline/privado): {core.Cor.BOLD}{core.Cor.CIANO}{url}{core.Cor.RESET}")
    print(f"  📊 Planilha ativa : {planilha_path.name} ({len(estado.itens_planilha)} nomes)")
    print(f"  📂 Pasta de PDFs  : {pdfs_dir} ({len(estado.fila_pdfs)} PDFs pendentes)")
    print("─" * 72)
    print("  💡 O navegador será aberto automaticamente. Para encerrar, pressione Ctrl+C.")
    print("=" * 72)

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
