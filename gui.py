# -*- coding: utf-8 -*-
"""
Renomeador Semiautomático de PDFs - Backend Web Local (v6.5)
------------------------------------------------------------
Servidor HTTP local (127.0.0.1) limpo e desacoplado:
- Os arquivos de interface (HTML, CSS e JS) ficam organizados na pasta `web/`.
- Este módulo gerencia apenas o estado da sessão, as rotas de API (`/api/...`)
  e a comunicação com o motor de PDFs e OCR local (`renomear.py`).
"""

import argparse
import base64
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
    sem_emoji = re.sub(r"[\U00010000-\U0010ffff]|[\u2600-\u27bf]|\ufe0f", "", sem_ansi)
    return " ".join(sem_emoji.split())


class EstadoGUI:
    def __init__(self, planilha_path: Path, pdfs_dir: Path, script_dir: Path):
        self.lock = threading.Lock()
        self.script_dir = script_dir
        self.web_dir = script_dir / "web"
        self.planilha_path = planilha_path
        self.pdfs_dir = pdfs_dir
        self.log_path = script_dir / core.ARQUIVO_LOG
        self.memoria_ia_path = script_dir / core.ARQUIVO_MEMORIA_IA
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
        self.total_juntados: int = 0
        self.total_pulados: int = 0
        self.inicio_sessao: float = time.time()
        self.timestamps_renomeacao: List[float] = []
        self.ultimo_evento: str = "Sistema pronto"
        self.erro_carregamento: str = ""

        self._cache_paginas: Dict[str, Optional[int]] = {}
        self._cache_brancas: Dict[str, List[int]] = {}
        self._cache_ia: Dict[str, dict] = {}
        self._cache_log_mtime: int = -1
        self._cache_log_dados: dict = {}
        # Mapa de diagnostico OCR por indice de item da planilha: idx_item -> dict
        self.diagnostico_fila_ocr: Dict[int, dict] = {}

        self.recarregar_tudo()

    def _limpar_caches_pdf(self):
        self._cache_paginas.clear()
        self._cache_brancas.clear()
        self._cache_ia.clear()

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
        self.diagnostico_fila_ocr.clear()
        self._limpar_caches_pdf()

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
        self._limpar_caches_pdf()

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

    def ultimo_pdf_renomeado(self) -> Optional[Path]:
        for acao in reversed(self.historico):
            if acao.tipo == "renomear" and acao.pdf_destino and acao.pdf_destino.exists():
                return acao.pdf_destino
        return None

    def obter_paginas_pdf(self, pdf_path: Optional[Path]) -> Optional[int]:
        if pdf_path is None or not pdf_path.exists():
            return None
        chave = f"{pdf_path.name}:{pdf_path.stat().st_mtime_ns}"
        if chave not in self._cache_paginas:
            self._cache_paginas[chave] = core.contar_paginas_pdf(pdf_path)
        return self._cache_paginas[chave]

    def obter_paginas_brancas(self, pdf_path: Optional[Path]) -> List[int]:
        if pdf_path is None or not pdf_path.exists():
            return []
        chave = f"{pdf_path.name}:{pdf_path.stat().st_mtime_ns}"
        if chave not in self._cache_brancas:
            self._cache_brancas[chave] = core.detectar_paginas_em_branco(pdf_path)
        return self._cache_brancas[chave]

    def obter_verificacao_ia(self, pdf_path: Optional[Path], usar_ocr: bool = True) -> Optional[dict]:
        if pdf_path is None or not pdf_path.exists() or not self.itens_planilha:
            return None
        mtime_mem = self.memoria_ia_path.stat().st_mtime_ns if self.memoria_ia_path.exists() else 0
        chave = f"{pdf_path.name}:{pdf_path.stat().st_mtime_ns}:{self.idx_nome}:{int(usar_ocr)}:{mtime_mem}"
        if chave not in self._cache_ia:
            res = core.verificar_pdf_contra_planilha(
                pdf_path,
                self.itens_planilha,
                self.idx_nome,
                usar_ocr=usar_ocr,
                caminho_memoria_ia=self.memoria_ia_path,
            )
            self._cache_ia[chave] = res
            if self.idx_nome is not None:
                self.diagnostico_fila_ocr[self.idx_nome] = {
                    "pdf_nome": pdf_path.name,
                    "nivel_cor": res.get("nivel_cor", "vermelho"),
                    "status": res.get("status", "duvida"),
                    "confianca": res.get("confianca", 0.0),
                    "motivo": res.get("motivo", ""),
                    "diagnostico_causa": res.get("diagnostico_causa", ""),
                }
        return self._cache_ia[chave]

    def analisar_arquivo_log(self) -> dict:
        """Lê `renomeacao.log` (com cache por mtime) para alimentar o painel de estatísticas."""
        if not self.log_path.exists():
            return {
                "total_sucesso": 0,
                "total_revertidos": 0,
                "total_erros": 0,
                "renomeados_hoje": 0,
                "por_hora": [],
                "por_dia": [],
                "ultimas_operacoes": [],
            }
        try:
            mtime = self.log_path.stat().st_mtime_ns
            if mtime == self._cache_log_mtime and self._cache_log_dados:
                return self._cache_log_dados

            hoje_str = time.strftime("%Y-%m-%d")
            sucessos = 0
            revertidos = 0
            erros = 0
            hoje_qtd = 0
            contagem_horas: Dict[str, int] = {}
            contagem_dias: Dict[str, int] = {}
            ultimas: List[dict] = []

            linhas = self.log_path.read_text(encoding="utf-8", errors="ignore").splitlines()
            for linha in linhas:
                m_ts = re.match(r"^\[(\d{4}-\d{2}-\d{2})\s+(\d{2}):(\d{2}):(\d{2})\]\s+(.*)$", linha)
                if not m_ts:
                    continue
                data_iso, hh, mm, ss, resto = m_ts.groups()
                horario_curto = f"{hh}:{mm}:{ss}"
                partes_data = data_iso.split("-")
                data_br = f"{partes_data[2]}/{partes_data[1]}" if len(partes_data) == 3 else data_iso

                if resto.startswith("SUCESSO:"):
                    sucessos += 1
                    if data_iso == hoje_str:
                        hoje_qtd += 1
                    contagem_horas[f"{hh}h"] = contagem_horas.get(f"{hh}h", 0) + 1
                    contagem_dias[data_br] = contagem_dias.get(data_br, 0) + 1
                    m_arq = re.search(r"'([^']+)'\s*->\s*'([^']+)'", resto)
                    if m_arq:
                        ultimas.append({
                            "data": data_br,
                            "horario": horario_curto,
                            "tipo": "SUCESSO",
                            "origem": m_arq.group(1),
                            "destino": m_arq.group(2),
                        })
                elif resto.startswith("REVERTIDO:"):
                    revertidos += 1
                    m_arq = re.search(r"'([^']+)'\s*->\s*'([^']+)'", resto)
                    if m_arq:
                        ultimas.append({
                            "data": data_br,
                            "horario": horario_curto,
                            "tipo": "DESFEITO",
                            "origem": m_arq.group(1),
                            "destino": m_arq.group(2),
                        })
                elif resto.startswith("ERRO:"):
                    erros += 1

            por_hora_lista = [
                {"rotulo": k, "valor": v}
                for k, v in sorted(contagem_horas.items(), key=lambda x: x[0])
            ]
            ultimos_dias_items = list(contagem_dias.items())[-7:]
            por_dia_lista = [{"rotulo": k, "valor": v} for k, v in ultimos_dias_items]

            res = {
                "total_sucesso": sucessos,
                "total_revertidos": revertidos,
                "total_erros": erros,
                "renomeados_hoje": hoje_qtd,
                "por_hora": por_hora_lista,
                "por_dia": por_dia_lista,
                "ultimas_operacoes": list(reversed(ultimas[-30:])),
            }
            self._cache_log_mtime = mtime
            self._cache_log_dados = res
            return res
        except Exception:
            return {
                "total_sucesso": 0,
                "total_revertidos": 0,
                "total_erros": 0,
                "renomeados_hoje": 0,
                "por_hora": [],
                "por_dia": [],
                "ultimas_operacoes": [],
            }

    def obter_estatisticas_completas(self) -> dict:
        total_plan = len(self.itens_planilha)
        concluidos = sum(1 for it in self.itens_planilha if it.concluido)
        pendentes = total_plan - concluidos
        pdfs_restantes = max(0, len(self.fila_pdfs) - self.idx_pdf)

        # Distribuição por Prioridade (1: Matrícula, 2: CPF, 3: Nome Completo)
        p1_tot = sum(1 for it in self.itens_planilha if it.matricula)
        p1_ok = sum(1 for it in self.itens_planilha if it.matricula and it.concluido)
        p2_tot = sum(1 for it in self.itens_planilha if not it.matricula and it.cpf)
        p2_ok = sum(1 for it in self.itens_planilha if not it.matricula and it.cpf and it.concluido)
        p3_tot = sum(1 for it in self.itens_planilha if not it.matricula and not it.cpf)
        p3_ok = sum(1 for it in self.itens_planilha if not it.matricula and not it.cpf and it.concluido)

        # Semáforo OCR atual
        ocr_verde = 0
        ocr_amarelo = 0
        ocr_vermelho = 0
        for idx, it in enumerate(self.itens_planilha):
            if it.concluido:
                ocr_verde += 1
            elif idx in self.diagnostico_fila_ocr:
                cor = self.diagnostico_fila_ocr[idx].get("nivel_cor", "vermelho")
                if cor == "verde":
                    ocr_verde += 1
                elif cor == "amarelo":
                    ocr_amarelo += 1
                else:
                    ocr_vermelho += 1
        ocr_nao_lido = max(0, total_plan - (ocr_verde + ocr_amarelo + ocr_vermelho))

        # Velocidade e ETA da sessão
        decorrido_seg = max(1.0, time.time() - self.inicio_sessao)
        if len(self.timestamps_renomeacao) >= 2:
            janela_seg = max(1.0, self.timestamps_renomeacao[-1] - self.inicio_sessao)
            ritmo_min = round((self.total_renomeados / janela_seg) * 60.0, 1)
            seg_por_pdf = round(janela_seg / max(1, self.total_renomeados), 1)
        elif self.total_renomeados == 1:
            ritmo_min = round((1.0 / max(5.0, min(decorrido_seg, 60.0))) * 60.0, 1)
            seg_por_pdf = round(min(decorrido_seg, 60.0), 1)
        else:
            ritmo_min = 0.0
            seg_por_pdf = 0.0

        eta_min = round(pdfs_restantes / ritmo_min, 1) if (ritmo_min > 0 and pdfs_restantes > 0) else 0.0

        return {
            "resumo": {
                "total_planilha": total_plan,
                "concluidos_planilha": concluidos,
                "pendentes_planilha": pendentes,
                "pdfs_restantes_fila": pdfs_restantes,
                "alunos_sem_pdf_estimado": max(0, pendentes - pdfs_restantes),
                "pdfs_extras_estimado": max(0, pdfs_restantes - pendentes),
                "tempo_sessao_min": round(decorrido_seg / 60.0, 1),
                "ritmo_por_minuto": ritmo_min,
                "segundos_por_pdf": seg_por_pdf,
                "eta_minutos": eta_min,
            },
            "prioridades": {
                "p1": {"rotulo": "Prioridade 1 (Matrícula)", "total": p1_tot, "concluidos": p1_ok},
                "p2": {"rotulo": "Prioridade 2 (CPF)", "total": p2_tot, "concluidos": p2_ok},
                "p3": {"rotulo": "Prioridade 3 (Nome Completo)", "total": p3_tot, "concluidos": p3_ok},
            },
            "semaforo_ocr": {
                "verde": ocr_verde,
                "amarelo": ocr_amarelo,
                "vermelho": ocr_vermelho,
                "nao_analisado": ocr_nao_lido,
            },
            "operacoes_sessao": {
                "renomeados": self.total_renomeados,
                "editados_manual": self.total_editados,
                "divididos": self.total_divididos,
                "juntados": self.total_juntados,
                "pulados": self.total_pulados,
            },
            "historico_log": self.analisar_arquivo_log(),
        }

    def salvar_relatorio(self):
        restantes_nomes = sum(1 for it in self.itens_planilha if not it.concluido)
        try:
            with open(self.script_dir / core.ARQUIVO_RELATORIO, "w", encoding="utf-8") as f:
                f.write(f"RELATÓRIO DE RENOMEAÇÃO (GUI v6.5) - {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 60 + "\n")
                f.write(f"Renomeados nesta sessão : {self.total_renomeados}\n")
                f.write(f"Editados manualmente    : {self.total_editados}\n")
                f.write(f"PDFs divididos (2+ p.)  : {self.total_divididos}\n")
                f.write(f"PDFs juntados           : {self.total_juntados}\n")
                f.write(f"PDFs pulados            : {self.total_pulados}\n")
                f.write(f"Nomes restantes planilha: {restantes_nomes}/{len(self.itens_planilha)}\n")
        except Exception:
            pass

    def serializar_estado(self) -> dict:
        pdf_cur = self.pdf_atual()
        pdf_next = self.pdf_proximo()
        item_cur = self.item_atual()
        paginas_cur = self.obter_paginas_pdf(pdf_cur)
        brancas_cur = self.obter_paginas_brancas(pdf_cur) if (paginas_cur and paginas_cur >= 2) else []
        ia_cur = self.obter_verificacao_ia(pdf_cur, usar_ocr=True) if pdf_cur else None
        ult_ren = self.ultimo_pdf_renomeado()

        arquivo_existente_aluno = None
        if item_cur:
            cand = self.pdfs_dir / f"{item_cur.nome_formatado}.pdf"
            if cand.exists():
                arquivo_existente_aluno = cand.name

        total_planilha = len(self.itens_planilha)
        concluidos_planilha = sum(1 for it in self.itens_planilha if it.concluido)
        mem_ia = core.carregar_memoria_ia(self.memoria_ia_path)
        variacoes_ocr = mem_ia.get("variacoes_ocr", {})
        padroes_alunos = mem_ia.get("padroes_alunos", {})
        termos_ignorados = mem_ia.get("termos_ignorados", [])

        # Mapear qual PDF da fila corresponde teoricamente a cada próximo item pendente
        mapa_pdf_previsto: Dict[int, str] = {}
        cursor_pdf = self.idx_pdf
        if self.idx_nome is not None:
            for i in range(self.idx_nome, len(self.itens_planilha)):
                if not self.itens_planilha[i].concluido:
                    if cursor_pdf < len(self.fila_pdfs):
                        mapa_pdf_previsto[i] = self.fila_pdfs[cursor_pdf].name
                        cursor_pdf += 1

        return {
            "tem_pdf": pdf_cur is not None,
            "pdf_atual": {
                "nome": pdf_cur.name,
                "url": f"/api/pdf?name={urllib.parse.quote(pdf_cur.name)}&t={pdf_cur.stat().st_mtime_ns}",
                "paginas": paginas_cur,
                "paginas_brancas": brancas_cur,
                "posicao": self.idx_pdf + 1,
                "total_fila": len(self.fila_pdfs),
                "restantes": max(0, len(self.fila_pdfs) - self.idx_pdf),
                "tem_proximo": pdf_next is not None,
                "proximo_nome": pdf_next.name if pdf_next else None,
            }
            if pdf_cur and pdf_cur.exists()
            else None,
            "pdf_proximo_url": (
                f"/api/pdf?name={urllib.parse.quote(pdf_next.name)}&t={pdf_next.stat().st_mtime_ns}"
                if pdf_next and pdf_next.exists()
                else None
            ),
            "ultimo_renomeado": ult_ren.name if ult_ren else None,
            "item_atual": {
                "idx": self.idx_nome,
                "indice_original": item_cur.indice_original,
                "nome_completo": item_cur.nome_completo,
                "matricula": item_cur.matricula,
                "cpf": item_cur.cpf,
                "nome_formatado": item_cur.nome_formatado,
                "prioridade": 1 if item_cur.matricula else (2 if item_cur.cpf else 3),
                "arquivo_existente": arquivo_existente_aluno,
            }
            if item_cur
            else None,
            "verificacao_ia": ia_cur,
            "memoria_ia": {
                "total_confirmacoes": int(mem_ia.get("total_confirmacoes", 0)),
                "total_regras": len(variacoes_ocr) + len(padroes_alunos) + len(termos_ignorados),
                "total_variacoes_ocr": len(variacoes_ocr),
                "total_padroes_alunos": len(padroes_alunos),
                "variacoes_ocr": variacoes_ocr,
                "termos_ignorados": termos_ignorados,
                "ultimos_aprendizados": mem_ia.get("ultimos_aprendizados", [])[:12],
            },
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
                "juntados_sessao": self.total_juntados,
                "pulados_sessao": self.total_pulados,
                "pode_desfazer": len(self.historico) > 0,
                "qtd_historico": len(self.historico),
            },
            "estatisticas": self.obter_estatisticas_completas(),
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
                    "cpf": it.cpf,
                    "prioridade": 1 if it.matricula else (2 if it.cpf else 3),
                    "nome_formatado": it.nome_formatado,
                    "concluido": it.concluido,
                    "arquivo_gerado": it.arquivo_gerado,
                    "pdf_previsto": mapa_pdf_previsto.get(idx, ""),
                    "atual": idx == self.idx_nome and not it.concluido,
                    "ocr": self.diagnostico_fila_ocr.get(idx),
                }
                for idx, it in enumerate(self.itens_planilha)
            ],
        }


# ==============================================================================
# AÇÕES DE NEGÓCIO
# ==============================================================================

def acao_confirmar_atual(
    estado: EstadoGUI,
    idx_escolhido: Optional[int] = None,
    juntar_se_existir: bool = False,
    ensinar_ia: bool = False,
) -> Tuple[bool, str]:
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None or not pdf_cur.exists():
        return False, "Nenhum arquivo PDF pendente selecionado."

    alvo_idx = idx_escolhido if idx_escolhido is not None else estado.idx_nome
    if alvo_idx is None or not (0 <= alvo_idx < len(estado.itens_planilha)):
        return False, "Não há nome selecionado na fila. Use Editar (E) para informar um nome."

    item = estado.itens_planilha[alvo_idx]
    caminho_existente = estado.pdfs_dir / f"{item.nome_formatado}.pdf"

    if juntar_se_existir and caminho_existente.exists():
        return acao_anexar_a_destino(estado, caminho_existente, marcar_idx_nome=alvo_idx)

    # Treinar a IA local continuamente antes de mover/renomear o PDF
    qtd_aprendidos, msg_ia = core.registrar_aprendizado_ia(
        pdf_cur,
        item,
        caminho_memoria=estado.memoria_ia_path,
        forcar_padrao_documento=ensinar_ia,
    )

    destino, erro = core.executar_renomeacao_arquivo(
        pdf_cur, item.nome_formatado, estado.pdfs_dir, estado.log_path
    )
    if not destino:
        return False, f"Erro ao renomear: {erro}"

    item.concluido = True
    item.arquivo_gerado = destino.name
    estado.diagnostico_fila_ocr[alvo_idx] = {
        "pdf_nome": pdf_cur.name,
        "nivel_cor": "verde",
        "status": "concluido",
        "confianca": 100.0,
        "motivo": f"Renomeado ({pdf_cur.name} -> {destino.name})",
    }
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
    estado.timestamps_renomeacao.append(time.time())
    estado.idx_pdf += 1
    if alvo_idx == estado.idx_nome:
        estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, alvo_idx + 1)
    else:
        if estado.idx_nome is None or estado.itens_planilha[estado.idx_nome].concluido:
            estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, 0)

    sufixo_ia = f" · [{msg_ia}]" if (ensinar_ia or qtd_aprendidos > 0) else ""
    estado.ultimo_evento = f"Renomeado: {pdf_cur.name} -> {destino.name}{sufixo_ia}"
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_editar_manual(
    estado: EstadoGUI,
    nome_digitado: str,
    matricula_digitada: str,
    cpf_digitado: str,
    avancar_planilha: bool,
) -> Tuple[bool, str]:
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None or not pdf_cur.exists():
        return False, "Nenhum arquivo PDF pendente selecionado."

    _, _, nome_formatado = core.formatar_nome_e_matricula(
        nome_digitado, matricula_digitada, cpf_digitado
    )
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
        estado.itens_planilha[estado.idx_nome].arquivo_gerado = destino.name
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
    estado.ultimo_evento = f"Renomeado (manual): {pdf_cur.name} -> {destino.name} ({modo_txt})"
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_juntar_pdfs(estado: EstadoGUI, quantidade: int = 2) -> Tuple[bool, str]:
    qtd = max(2, int(quantidade))
    alvos = estado.fila_pdfs[estado.idx_pdf : estado.idx_pdf + qtd]
    if len(alvos) < 2:
        return False, "Não há outros PDFs suficientes na fila após o atual para juntar."

    estado.backup_divisao_dir.mkdir(parents=True, exist_ok=True)
    backups: List[Path] = []
    try:
        for p in alvos:
            bkp = core.obter_destino_unico(estado.backup_divisao_dir, p.stem)
            shutil.copy2(p, bkp)
            backups.append(bkp)
    except Exception as e:
        return False, f"Erro ao criar backup antes de juntar: {e}"

    pdf_principal = alvos[0]
    ok, msg = core.juntar_arquivos_pdf(alvos, pdf_principal)
    if not ok:
        return False, msg

    for extra in alvos[1:]:
        try:
            extra.unlink(missing_ok=True)
        except Exception:
            pass

    del estado.fila_pdfs[estado.idx_pdf + 1 : estado.idx_pdf + len(alvos)]
    estado.historico.append(
        core.AcaoHistorico(
            tipo="juntar",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=estado.idx_nome if estado.idx_nome is not None else 0,
            arquivos_criados=[pdf_principal],
            arquivos_originais=list(alvos),
            backups_originais=backups,
            descricao=f"Juntou {len(alvos)} PDFs ({', '.join(p.name for p in alvos)})",
        )
    )
    estado.total_juntados += 1
    estado._limpar_caches_pdf()
    estado.ultimo_evento = f"Juntou {len(alvos)} PDFs em '{pdf_principal.name}'."
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_anexar_a_destino(
    estado: EstadoGUI,
    pdf_destino_existente: Path,
    marcar_idx_nome: Optional[int] = None,
) -> Tuple[bool, str]:
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None or not pdf_cur.exists():
        return False, "Nenhum PDF atual selecionado."
    if not pdf_destino_existente.exists():
        return False, "O arquivo PDF de destino não foi encontrado."

    estado.backup_divisao_dir.mkdir(parents=True, exist_ok=True)
    bkp_dest = core.obter_destino_unico(estado.backup_divisao_dir, pdf_destino_existente.stem)
    bkp_cur = core.obter_destino_unico(estado.backup_divisao_dir, pdf_cur.stem)
    try:
        shutil.copy2(pdf_destino_existente, bkp_dest)
        shutil.copy2(pdf_cur, bkp_cur)
    except Exception as e:
        return False, f"Erro ao criar backup para anexar: {e}"

    ok, msg = core.juntar_arquivos_pdf([pdf_destino_existente, pdf_cur], pdf_destino_existente)
    if not ok:
        return False, msg

    try:
        pdf_cur.unlink(missing_ok=True)
    except Exception:
        pass

    idx_ant_nome = estado.idx_nome if estado.idx_nome is not None else 0
    if marcar_idx_nome is not None and 0 <= marcar_idx_nome < len(estado.itens_planilha):
        estado.itens_planilha[marcar_idx_nome].concluido = True
        estado.itens_planilha[marcar_idx_nome].arquivo_gerado = pdf_destino_existente.name
        if marcar_idx_nome == estado.idx_nome:
            estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, marcar_idx_nome + 1)

    estado.historico.append(
        core.AcaoHistorico(
            tipo="anexar_anterior",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=idx_ant_nome,
            arquivos_originais=[pdf_destino_existente, pdf_cur],
            backups_originais=[bkp_dest, bkp_cur],
            descricao=f"Anexou {pdf_cur.name} ao final de {pdf_destino_existente.name}",
        )
    )
    estado.total_juntados += 1
    estado.idx_pdf += 1
    estado._limpar_caches_pdf()
    estado.ultimo_evento = f"Anexado '{pdf_cur.name}' ao final de '{pdf_destino_existente.name}'."
    estado.salvar_relatorio()
    return True, estado.ultimo_evento


def acao_anexar_ao_anterior(estado: EstadoGUI) -> Tuple[bool, str]:
    ult = estado.ultimo_pdf_renomeado()
    if ult is None:
        return False, "Ainda não há nenhum PDF renomeado nesta sessão para anexar."
    return acao_anexar_a_destino(estado, ult)


def _executar_edicao_pdf_com_backup(
    estado: EstadoGUI,
    funcao_edicao,
) -> Tuple[bool, str]:
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None or not pdf_cur.exists():
        return False, "Nenhum arquivo PDF pendente selecionado."

    estado.backup_divisao_dir.mkdir(parents=True, exist_ok=True)
    bkp = core.obter_destino_unico(estado.backup_divisao_dir, pdf_cur.stem)
    try:
        shutil.copy2(pdf_cur, bkp)
    except Exception as e:
        return False, f"Erro ao criar backup: {e}"

    ok, msg = funcao_edicao(pdf_cur)
    if not ok:
        bkp.unlink(missing_ok=True)
        return False, msg

    estado.historico.append(
        core.AcaoHistorico(
            tipo="editar_pdf",
            idx_pdf_anterior=estado.idx_pdf,
            idx_nome_anterior=estado.idx_nome if estado.idx_nome is not None else 0,
            pdf_origem=pdf_cur,
            backup_origem=bkp,
            descricao=msg,
        )
    )
    estado._limpar_caches_pdf()
    estado.ultimo_evento = msg
    return True, msg


def acao_diagnosticar_fila_ocr(estado: EstadoGUI, limite: int = 30) -> Tuple[bool, str]:
    """
    Analisa os próximos PDFs pendentes da fila com IA/OCR local e preenche o semáforo
    (verde / amarelo / vermelho) ao lado de cada item na lista da fila!
    """
    if not estado.fila_pdfs or estado.idx_pdf >= len(estado.fila_pdfs):
        return False, "Não há PDFs pendentes na fila para diagnosticar."

    pendentes_indices = [
        i for i in range(estado.idx_nome or 0, len(estado.itens_planilha))
        if not estado.itens_planilha[i].concluido
    ]
    pdfs_alvos = estado.fila_pdfs[estado.idx_pdf : estado.idx_pdf + max(1, limite)]

    verdes = 0
    amarelos = 0
    vermelhos = 0

    for offset, pdf_cand in enumerate(pdfs_alvos):
        if not pdf_cand.exists():
            continue
        idx_esperado = pendentes_indices[offset] if offset < len(pendentes_indices) else None
        res = core.verificar_pdf_contra_planilha(
            pdf_cand,
            estado.itens_planilha,
            idx_esperado,
            usar_ocr=True,
            caminho_memoria_ia=estado.memoria_ia_path,
        )
        cor = res.get("nivel_cor", "vermelho")
        if cor == "verde":
            verdes += 1
        elif cor == "amarelo":
            amarelos += 1
        else:
            vermelhos += 1

        if idx_esperado is not None:
            estado.diagnostico_fila_ocr[idx_esperado] = {
                "pdf_nome": pdf_cand.name,
                "nivel_cor": cor,
                "status": res.get("status", "duvida"),
                "confianca": res.get("confianca", 0.0),
                "motivo": res.get("motivo", ""),
                "diagnostico_causa": res.get("diagnostico_causa", ""),
                "idx_sugerido": res.get("idx_sugerido"),
                "nome_sugerido": res.get("nome_completo", ""),
            }

    estado.ultimo_evento = (
        f"Diagnóstico OCR concluído em {len(pdfs_alvos)} PDF(s): "
        f"{verdes} OK (verde) · {amarelos} Revisão (amarelo) · {vermelhos} Não lido (vermelho)."
    )
    return True, estado.ultimo_evento


def acao_auto_renomear_confirmados(estado: EstadoGUI) -> Tuple[bool, str]:
    """
    Renomeia automaticamente os PDFs da ponta da fila que tiverem status 'confirmado_atual' (verde),
    parando assim que encontrar um amarelo (TCC em grupo / fora de ordem / dúvida familiar) ou vermelho.
    """
    if not estado.fila_pdfs or estado.idx_pdf >= len(estado.fila_pdfs):
        return False, "Não há PDFs pendentes na fila."

    renomeados_auto = 0
    while estado.idx_pdf < len(estado.fila_pdfs):
        pdf_cand = estado.pdf_atual()
        if pdf_cand is None or not pdf_cand.exists():
            break

        res = core.verificar_pdf_contra_planilha(
            pdf_cand,
            estado.itens_planilha,
            estado.idx_nome,
            usar_ocr=True,
            caminho_memoria_ia=estado.memoria_ia_path,
        )
        # Só auto-renomeia no escuro se for VERDE (confirmou o aluno atual e não é TCC em grupo nem familiar)
        if (
            res.get("nivel_cor") == "verde"
            and res.get("status") == "confirmado_atual"
            and res.get("idx_sugerido") is not None
            and float(res.get("confianca", 0.0)) >= 82.0
        ):
            ok, _ = acao_confirmar_atual(estado, idx_escolhido=int(res["idx_sugerido"]))
            if ok:
                renomeados_auto += 1
                continue
        break

    # Também rodar o diagnóstico nos próximos PDFs para colorir a lista
    acao_diagnosticar_fila_ocr(estado, limite=20)

    if renomeados_auto == 0:
        estado.ultimo_evento = (
            "O PDF atual requer revisão manual (amarelo/vermelho). Diagnóstico da fila atualizado."
        )
        return False, estado.ultimo_evento

    estado.ultimo_evento = (
        f"Auto-renomeados {renomeados_auto} PDF(s) 100% confirmados! Parado no próximo item de revisão."
    )
    return True, estado.ultimo_evento


def acao_duplicar_para_grupo_tcc(estado: EstadoGUI, indices_grupo: List[int]) -> Tuple[bool, str]:
    """
    Quando um PDF é um TCC em grupo com N alunos da planilha, cria cópias renomeadas
    para todos os N alunos do grupo de uma só vez e avança para o próximo PDF!
    """
    pdf_cur = estado.pdf_atual()
    if pdf_cur is None or not pdf_cur.exists():
        return False, "Nenhum arquivo PDF pendente selecionado."

    indices_validos = [
        int(i) for i in indices_grupo if 0 <= int(i) < len(estado.itens_planilha)
    ]
    if not indices_validos:
        return False, "Nenhum aluno válido selecionado para o grupo."

    estado.backup_divisao_dir.mkdir(parents=True, exist_ok=True)
    bkp = core.obter_destino_unico(estado.backup_divisao_dir, pdf_cur.stem)
    shutil.copy2(pdf_cur, bkp)

    criados: List[Path] = []
    try:
        for idx_al in indices_validos:
            item = estado.itens_planilha[idx_al]
            dest = core.obter_destino_unico(estado.pdfs_dir, item.nome_formatado)
            shutil.copy2(pdf_cur, dest)
            core.registrar_log(estado.log_path, f"SUCESSO: '{pdf_cur.name}' -> '{dest.name}' (TCC/Grupo)")
            item.concluido = True
            item.arquivo_gerado = dest.name
            criados.append(dest)
            estado.total_renomeados += 1

        pdf_cur.unlink(missing_ok=True)
    except Exception as e:
        return False, f"Erro ao gerar cópias do grupo: {e}"

    estado.idx_pdf += 1
    estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, 0)
    estado._limpar_caches_pdf()
    estado.ultimo_evento = (
        f"TCC/Grupo salvo para {len(criados)} alunos: {', '.join(p.name for p in criados)}"
    )
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
    estado._limpar_caches_pdf()

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
    estado._limpar_caches_pdf()
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
    estado.ultimo_evento = f"{len(itens)} nomes carregados com sucesso."
    return True, estado.ultimo_evento


def acao_upload_planilha(estado: EstadoGUI, nome_arquivo: str, conteudo_b64: str) -> Tuple[bool, str]:
    sufixo = Path(nome_arquivo).suffix.lower()
    if sufixo not in (".xlsx", ".csv", ".tsv", ".txt"):
        return False, "Formato não suportado. Envie um arquivo .xlsx, .csv ou .tsv."
    try:
        dados = base64.b64decode(conteudo_b64)
        destino = estado.script_dir / f"planilha_importada{sufixo}"
        destino.write_bytes(dados)
        itens, msg = core.carregar_planilha(destino)
        if not itens:
            return False, f"Erro ao ler arquivo enviado: {msg}"
        estado.planilha_path = destino
        estado.recarregar_tudo()
        estado.ultimo_evento = f"Planilha '{nome_arquivo}' importada ({len(itens)} alunos)."
        return True, estado.ultimo_evento
    except Exception as e:
        return False, f"Falha ao importar planilha: {e}"


def acao_adicionar_pessoa_fila(
    estado: EstadoGUI,
    nome: str,
    matricula: str,
    cpf: str,
) -> Tuple[bool, str]:
    mat_limpa = core.limpar_matricula(matricula)
    cpf_limpo = core.limpar_cpf(cpf)
    nome_completo, _, nome_formatado = core.formatar_nome_e_matricula(nome, mat_limpa, cpf_limpo)
    if not nome_formatado:
        return False, "Informe um nome válido para adicionar à lista."

    novo_idx_orig = (max((it.indice_original for it in estado.itens_planilha), default=0) + 1)
    estado.itens_planilha.append(
        core.ItemPlanilha(
            indice_original=novo_idx_orig,
            nome_completo=nome_completo,
            matricula=mat_limpa,
            cpf=cpf_limpo,
            nome_formatado=nome_formatado,
        )
    )
    if estado.idx_nome is None:
        estado.idx_nome = len(estado.itens_planilha) - 1
    estado.ultimo_evento = f"Aluno '{nome_completo}' ({nome_formatado}.pdf) adicionado à fila."
    return True, estado.ultimo_evento


def acao_remover_pessoa_fila(estado: EstadoGUI, idx: int) -> Tuple[bool, str]:
    if not (0 <= idx < len(estado.itens_planilha)):
        return False, "Índice de aluno inválido."
    removido = estado.itens_planilha.pop(idx)
    estado.diagnostico_fila_ocr.clear()
    estado.idx_nome = core.proximo_indice_nome_pendente(estado.itens_planilha, 0)
    estado.ultimo_evento = f"Aluno '{removido.nome_completo}' removido da fila."
    return True, estado.ultimo_evento


# ==============================================================================
# SERVIDOR HTTP LOCAL (SERVE `web/` + `/api/...`)
# ==============================================================================

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

    def _servir_arquivo_web(self, nome_arquivo: str, content_type: str):
        caminho = (self.estado.web_dir / nome_arquivo).resolve()
        if not caminho.exists() or caminho.parent != self.estado.web_dir.resolve():
            self.send_error(404, f"Arquivo web nao encontrado: {nome_arquivo}")
            return
        dados = caminho.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(dados)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(dados)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        rota = parsed.path

        if rota in ("/", "/index.html"):
            self._servir_arquivo_web("index.html", "text/html; charset=utf-8")
            return
        if rota == "/styles.css":
            self._servir_arquivo_web("styles.css", "text/css; charset=utf-8")
            return
        if rota == "/app.js":
            self._servir_arquivo_web("app.js", "application/javascript; charset=utf-8")
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

        if rota == "/api/page_image":
            qs = urllib.parse.parse_qs(parsed.query)
            nome_arq = qs.get("name", [""])[0]
            pagina = int(qs.get("page", ["1"])[0] or "1")
            nome_seguro = Path(nome_arq).name
            caminho_pdf = (self.estado.pdfs_dir / nome_seguro).resolve()
            if (
                not nome_seguro
                or not caminho_pdf.exists()
                or not caminho_pdf.is_file()
                or caminho_pdf.parent != self.estado.pdfs_dir.resolve()
            ):
                self.send_error(404, "PDF nao encontrado")
                return

            png_bytes = core.renderizar_pagina_png(caminho_pdf, pagina_1based=pagina, dpi=115)
            if not png_bytes:
                self.send_error(500, "Nao foi possivel renderizar a pagina")
                return

            self.send_response(200)
            self.send_header("Content-Type", "image/png")
            self.send_header("Content-Length", str(len(png_bytes)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(png_bytes)
            return

        if rota == "/api/export_report":
            with self.estado.lock:
                arq_gerado, _ = core.exportar_planilha_conferencia(
                    self.estado.itens_planilha,
                    self.estado.pdfs_dir,
                    self.estado.script_dir / core.ARQUIVO_CONFERENCIA_XLSX,
                    self.estado.script_dir / core.ARQUIVO_CONFERENCIA_CSV,
                )
                dados = arq_gerado.read_bytes()
                mime = (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    if arq_gerado.suffix.lower() == ".xlsx"
                    else "text/csv; charset=utf-8"
                )
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Disposition", f'attachment; filename="{arq_gerado.name}"')
            self.send_header("Content-Length", str(len(dados)))
            self.end_headers()
            self.wfile.write(dados)
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
                juntar_ex = bool(payload.get("juntar_se_existir", False))
                ensinar_ia = bool(payload.get("ensinar_ia", False))
                ok, msg = acao_confirmar_atual(
                    self.estado,
                    int(idx_esc) if idx_esc is not None else None,
                    juntar_se_existir=juntar_ex,
                    ensinar_ia=ensinar_ia,
                )

            elif rota == "/api/train_ai":
                modo = str(payload.get("modo", "ensinar_atual"))
                if modo == "ignorar_termo":
                    termo = str(payload.get("termo", ""))
                    ok, msg = core.adicionar_termo_ignorado_ia(termo, self.estado.memoria_ia_path)
                    self.estado._limpar_caches_pdf()
                    self.estado.ultimo_evento = msg
                elif modo == "remover_ignorado":
                    termo = str(payload.get("termo", ""))
                    ok, msg = core.remover_termo_ignorado_ia(termo, self.estado.memoria_ia_path)
                    self.estado._limpar_caches_pdf()
                    self.estado.ultimo_evento = msg
                else:
                    # ensinar_atual: aprende o padrao do PDF atual para o aluno atual e ja confirma se solicitado
                    confirmar_apos = bool(payload.get("confirmar", True))
                    idx_alvo = payload.get("idx_nome")
                    idx_final = int(idx_alvo) if idx_alvo is not None else self.estado.idx_nome
                    if confirmar_apos:
                        ok, msg = acao_confirmar_atual(
                            self.estado,
                            idx_escolhido=idx_final,
                            ensinar_ia=True,
                        )
                    else:
                        pdf_c = self.estado.pdf_atual()
                        if pdf_c and idx_final is not None and 0 <= idx_final < len(self.estado.itens_planilha):
                            _, msg = core.registrar_aprendizado_ia(
                                pdf_c,
                                self.estado.itens_planilha[idx_final],
                                caminho_memoria=self.estado.memoria_ia_path,
                                forcar_padrao_documento=True,
                            )
                            self.estado._limpar_caches_pdf()
                            self.estado.ultimo_evento = msg
                            ok = True
                        else:
                            ok, msg = False, "Nenhum PDF ou aluno selecionado para ensinar a IA."

            elif rota == "/api/edit":
                ok, msg = acao_editar_manual(
                    self.estado,
                    str(payload.get("nome", "")),
                    str(payload.get("matricula", "")),
                    str(payload.get("cpf", "")),
                    bool(payload.get("avancar_planilha", False)),
                )

            elif rota == "/api/split":
                ok, msg = acao_dividir_pdf(
                    self.estado,
                    modo=str(payload.get("modo", "corte")),
                    pagina_corte=int(payload.get("pagina_corte", 1)),
                    ja_renomear_parte1=bool(payload.get("ja_renomear_parte1", True)),
                )

            elif rota == "/api/merge":
                ok, msg = acao_juntar_pdfs(
                    self.estado,
                    quantidade=int(payload.get("quantidade", 2)),
                )

            elif rota == "/api/append_previous":
                ok, msg = acao_anexar_ao_anterior(self.estado)

            elif rota == "/api/rotate":
                ang = int(payload.get("angulo", 90))
                pag = int(payload.get("pagina", 0))
                ok, msg = _executar_edicao_pdf_com_backup(
                    self.estado,
                    lambda p: core.rotacionar_paginas_pdf(p, ang, pagina_alvo=pag if pag > 0 else None),
                )

            elif rota == "/api/auto_orient":
                ok, msg = _executar_edicao_pdf_com_backup(
                    self.estado,
                    lambda p: core.auto_orientar_pdf(p),
                )

            elif rota == "/api/auto_crop_borders":
                pag = int(payload.get("pagina", 0))
                ok, msg = _executar_edicao_pdf_com_backup(
                    self.estado,
                    lambda p: core.auto_recortar_bordas_pretas(p, pagina_alvo=pag if pag > 0 else None),
                )

            elif rota == "/api/crop_page":
                pag = int(payload.get("pagina", 0))
                x0 = float(payload.get("x0", 0.0))
                y0 = float(payload.get("y0", 0.0))
                x1 = float(payload.get("x1", 1.0))
                y1 = float(payload.get("y1", 1.0))
                ok, msg = _executar_edicao_pdf_com_backup(
                    self.estado,
                    lambda p: core.recortar_pagina_pdf(
                        p, x0, y0, x1, y1, pagina_alvo=pag if pag > 0 else None
                    ),
                )

            elif rota == "/api/delete_pages":
                pags = [int(x) for x in payload.get("paginas", [])]
                ok, msg = _executar_edicao_pdf_com_backup(
                    self.estado,
                    lambda p: core.excluir_paginas_pdf(p, pags),
                )

            elif rota == "/api/reorder_pages":
                inv = bool(payload.get("inverter", True))
                ordem = payload.get("ordem")
                ok, msg = _executar_edicao_pdf_com_backup(
                    self.estado,
                    lambda p: core.reordenar_paginas_pdf(p, nova_ordem_1based=ordem, inverter=inv),
                )

            elif rota == "/api/scan_ocr_queue":
                limite = int(payload.get("limite", 25))
                ok, msg = acao_diagnosticar_fila_ocr(self.estado, limite=limite)

            elif rota == "/api/auto_batch_ocr":
                ok, msg = acao_auto_renomear_confirmados(self.estado)

            elif rota == "/api/duplicate_group":
                indices = payload.get("indices", [])
                ok, msg = acao_duplicar_para_grupo_tcc(self.estado, indices)

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

            elif rota == "/api/upload_sheet":
                ok, msg = acao_upload_planilha(
                    self.estado,
                    str(payload.get("filename", "planilha.xlsx")),
                    str(payload.get("content_b64", "")),
                )

            elif rota == "/api/add_person":
                ok, msg = acao_adicionar_pessoa_fila(
                    self.estado,
                    str(payload.get("nome", "")),
                    str(payload.get("matricula", "")),
                    str(payload.get("cpf", "")),
                )

            elif rota == "/api/delete_person":
                ok, msg = acao_remover_pessoa_fila(
                    self.estado, int(payload.get("idx", -1))
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
        description="Renomeador de PDFs v6.5 - Interface Gráfica Web Local (127.0.0.1)"
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
    print("  RENOMEADOR DE PDFs v6.5 (Arquitetura Modular web/ + PyMuPDF)")
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
