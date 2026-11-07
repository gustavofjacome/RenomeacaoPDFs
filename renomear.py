# -*- coding: utf-8 -*-
"""
Renomeador Semiautomático de PDFs (CLI + Core v6.0)
---------------------------------------------------
Melhorias v6.0:
- Usa PyMuPDF moderno (`import pymupdf`) em vez do antigo `fitz` depreciado.
- Regra de 3 Prioridades de Renomeação:
  1) Primeiro_Segundo_Matricula (quando possui matrícula válida)
  2) Primeiro_Segundo_CPF (quando matrícula é vazia / "não consta", mas possui CPF)
  3) nome_completo_separado_por_underline (em minúsculas quando não tem matrícula nem CPF)
- Juntar PDFs (Merge) e Anexar ao PDF Anterior já renomeado.
- Girar páginas (90°, -90°, 180°) individualmente ou no PDF inteiro + Auto-Orientação.
- Recorte individual de página (manual na GUI) + Remoção automática de bordas pretas de scanner.
- Detecção e exclusão de páginas em branco ou páginas específicas + Inverter ordem de páginas.
- Leitura inteligente local (Texto embutido + OCR local leve no cabeçalho) para verificar
  Nome/Matrícula/CPF na planilha e auto-renomear os confirmados (100% offline/local).
- Exportação de Planilha de Conferência Final (.xlsx / .csv) com Concluídos, Faltantes e Extras.
"""

import argparse
import csv
import difflib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
import warnings
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

warnings.filterwarnings("ignore", category=DeprecationWarning)

# --- CONSTANTES ---
NOME_PLANILHA_PADRAO = "nomes.xlsx"
PASTA_PDFS_PADRAO = "pdfs"
PASTA_BACKUP_DIVISAO = ".backup_divididos"
ARQUIVO_LOG = "renomeacao.log"
ARQUIVO_RELATORIO = "relatorio_final.txt"
ARQUIVO_CONFERENCIA_XLSX = "relatorio_conferencia.xlsx"
ARQUIVO_CONFERENCIA_CSV = "relatorio_conferencia.csv"
ARQUIVO_MEMORIA_IA = "memoria_ia.json"

PARTICULAS_IGNORADAS = {"de", "da", "do", "dos", "das", "e", "del", "d", "d'"}

VALORES_NAO_CONSTA = {
    "",
    "nan",
    "none",
    "null",
    "nil",
    "-",
    "--",
    "---",
    "s/m",
    "sm",
    "s/n",
    "sn",
    "n/c",
    "nc",
    "n/a",
    "na",
    "nao consta",
    "não consta",
    "nao possui",
    "não possui",
    "sem matricula",
    "sem matrícula",
    "sem cpf",
    "inexistente",
    "0",
    "00",
    "000",
}

CABECALHOS_CONHECIDOS = {
    "nome",
    "nomes",
    "nome completo",
    "nome_completo",
    "nome do pdf",
    "nome_do_pdf",
    "aluno",
    "estudante",
    "discente",
    "candidato",
    "servidor",
    "pessoa",
    "matricula",
    "matrícula",
    "siape",
    "registro",
    "ra",
    "cpf",
    "documento",
    "id",
}


def _obter_pymupdf():
    """
    Importa o PyMuPDF moderno (`import pymupdf`) sem disparar o aviso de depreciação do antigo `fitz`.
    """
    try:
        import pymupdf  # type: ignore
        return pymupdf
    except ImportError:
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                import fitz as pymupdf  # type: ignore
                return pymupdf
        except ImportError:
            return None


# --- CORES ANSI (com suporte a Windows) ---
def ativar_cores_windows():
    if os.name == "nt":
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.GetStdHandle(-11)
            mode = ctypes.c_ulong()
            if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                kernel32.SetConsoleMode(handle, mode.value | 0x0004)
        except Exception:
            pass


class Cor:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    VERDE = "\033[92m"
    AMARELO = "\033[93m"
    VERMELHO = "\033[91m"
    CIANO = "\033[96m"
    AZUL = "\033[94m"
    MAGENTA = "\033[95m"


@dataclass
class ItemPlanilha:
    indice_original: int  # 1-based na planilha
    nome_completo: str
    matricula: str
    nome_formatado: str
    cpf: str = ""
    concluido: bool = False
    arquivo_gerado: str = ""


@dataclass
class AcaoHistorico:
    tipo: str  # 'renomear', 'pular_nome', 'pular_pdf', 'dividir', 'juntar', 'editar_pdf', 'anexar_anterior'
    idx_pdf_anterior: int
    idx_nome_anterior: int
    pdf_origem: Optional[Path] = None
    pdf_destino: Optional[Path] = None
    idx_nome_concluido: Optional[int] = None
    arquivos_criados: List[Path] = field(default_factory=list)
    arquivos_originais: List[Path] = field(default_factory=list)
    backups_originais: List[Path] = field(default_factory=list)
    backup_origem: Optional[Path] = None
    descricao: str = ""


# ==============================================================================
# 1. FORMATAÇÃO DE NOMES, MATRÍCULAS E CPF (3 NÍVEIS DE PRIORIDADE)
# ==============================================================================

def remover_acentos(texto: str) -> str:
    nfkd = unicodedata.normalize("NFKD", str(texto or ""))
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def normalizar_busca(texto: str) -> str:
    """Remove acentos e coloca em minúsculas para busca rápida."""
    return remover_acentos(texto).lower()


def limpar_nome_arquivo(nome: str) -> str:
    """Remove caracteres inválidos para arquivos no Windows/Linux."""
    caracteres_invalidos = r'[/\\:*?"<>|]'
    nome_limpo = re.sub(caracteres_invalidos, "-", str(nome))
    nome_limpo = " ".join(nome_limpo.split())
    nome_limpo = nome_limpo.strip(". ")
    return nome_limpo[:200] if len(nome_limpo) > 200 else nome_limpo


def _eh_valor_nao_consta(valor: Optional[str]) -> bool:
    if valor is None:
        return True
    limpo = normalizar_busca(str(valor)).strip(" ._-/\\")
    if limpo in VALORES_NAO_CONSTA:
        return True
    if any(termo in limpo for termo in ("nao consta", "sem matricula", "sem cpf", "nao informado", "desconhecid")):
        return True
    return False


def limpar_cpf(valor: Optional[str]) -> str:
    """
    Limpa e valida se o valor representa um CPF (retorna apenas os dígitos, ex: 11 dígitos).
    Retorna '' se for vazio ou 'não consta'.
    """
    if _eh_valor_nao_consta(valor):
        return ""
    texto = str(valor).strip()
    if re.match(r"^\d+\.0+$", texto):
        texto = texto.split(".")[0]
    elif re.match(r"^\d+\.\d+e\+\d+$", texto.lower()):
        try:
            texto = str(int(float(texto)))
        except Exception:
            pass
    digitos = re.sub(r"\D", "", texto)
    if not digitos:
        return ""
    # Se veio com 9, 10 ou 11 dígitos (quando o Excel comeu zero à esquerda do CPF)
    if 9 <= len(digitos) <= 11:
        return digitos.zfill(11)
    return digitos


def limpar_matricula(valor: Optional[str]) -> str:
    """Limpa matrícula vinda do Excel (remove .0 de floats, ignora 'não consta', etc.)."""
    if _eh_valor_nao_consta(valor):
        return ""
    texto = str(valor).strip()
    if re.match(r"^\d+\.0+$", texto):
        texto = texto.split(".")[0]
    elif re.match(r"^\d+\.\d+e\+\d+$", texto.lower()):
        try:
            texto = str(int(float(texto)))
        except Exception:
            pass
    # Se parece um CPF formatado (000.000.000-00), limpar pontuação
    if re.match(r"^\d{3}\.\d{3}\.\d{3}-\d{2}$", texto):
        return re.sub(r"\D", "", texto)
    limpo = limpar_nome_arquivo(texto).replace(" ", "")
    # Se sobrou apenas letras sem nenhum dígito (ex: "NC", " isento"), não é matrícula
    if not any(ch.isdigit() for ch in limpo):
        return ""
    return limpo


def capitalizar_palavra_nome(palavra: str) -> str:
    """Formata uma palavra do nome (ex: 'GUSTAVO' -> 'Gustavo', 'Jácome' -> 'Jácome')."""
    if not palavra:
        return ""
    if palavra.isupper() or palavra.islower():
        return palavra.capitalize()
    return palavra[0].upper() + palavra[1:]


def extrair_primeiro_segundo_nome(nome_completo: str) -> Tuple[str, ...]:
    """
    Extrai o primeiro e o segundo nome da pessoa ignorando preposições
    (de, da, do, dos, das, e).
    """
    texto = re.sub(r"[_\-]+", " ", str(nome_completo).strip())
    palavras_brutas = [p.strip(",.;:") for p in texto.split() if p.strip(",.;:")]
    if not palavras_brutas:
        return ()

    palavras_filtradas = [
        p for p in palavras_brutas if p.lower() not in PARTICULAS_IGNORADAS
    ]
    escolhidas = palavras_filtradas if palavras_filtradas else palavras_brutas
    primeiras_duas = escolhidas[:2]
    return tuple(capitalizar_palavra_nome(p) for p in primeiras_duas)


def formatar_nome_completo_underline(nome_completo: str) -> str:
    """
    Formata o nome completo separado por '_' em minúsculas sem acentos
    para a Prioridade 3 (quando não há matrícula nem CPF).
    Exemplo: 'Gustavo Ferreira Jácome de Lira' -> 'gustavo_ferreira_jacome_de_lira'
    """
    texto = remover_acentos(str(nome_completo or "").strip())
    texto = re.sub(r"[_\-]+", " ", texto)
    palavras = []
    for p in texto.split():
        p_limpa = re.sub(r"[^a-zA-Z0-9]", "", p).lower()
        if p_limpa:
            palavras.append(p_limpa)
    return limpar_nome_arquivo("_".join(palavras))


def formatar_nome_e_matricula(
    nome_ou_linha: str,
    matricula: Optional[str] = None,
    cpf: Optional[str] = None,
    forcar_reformatacao: bool = False,
) -> Tuple[str, str, str]:
    """
    Aplica a regra de 3 prioridades de renomeação:
    - Prioridade 1 (Tem Matrícula): Primeiro_Segundo_Matricula
    - Prioridade 2 (Não tem Matrícula / 'Não consta', mas tem CPF): Primeiro_Segundo_CPF
    - Prioridade 3 (Sem Matrícula e sem CPF): nome_completo_separado_por_underline (minúsculo)

    Retorna: (nome_completo_legivel, identificador_usado, nome_formatado_final)
    """
    nome_str = str(nome_ou_linha or "").strip()
    if nome_str.lower().endswith(".pdf"):
        nome_str = nome_str[:-4].strip()

    mat_str = limpar_matricula(matricula)
    cpf_str = limpar_cpf(cpf)

    # Caso 1: Veio apenas 1 coluna já com underscores
    if not mat_str and not cpf_str and "_" in nome_str and " " not in nome_str:
        partes = [p for p in nome_str.split("_") if p]
        if len(partes) >= 2 and re.match(r"^\d{3,}$", partes[-1]):
            id_detectado = partes[-1]
            partes_nome = partes[:-1]
            if forcar_reformatacao or len(partes_nome) > 2 or any(
                p.lower() in PARTICULAS_IGNORADAS for p in partes_nome
            ):
                nomes_limpos = extrair_primeiro_segundo_nome(" ".join(partes_nome))
                formatado = "_".join(list(nomes_limpos) + [id_detectado])
            else:
                formatado = "_".join(partes)
            return " ".join(partes_nome), id_detectado, limpar_nome_arquivo(formatado)
        # Se já é nome_completo_com_underline sem números (Prioridade 3)
        nome_legivel = " ".join(capitalizar_palavra_nome(p) for p in partes)
        return nome_legivel, "", formatar_nome_completo_underline(nome_legivel)

    # Caso 2: Usuário digitou Nome + CPF ou Matrícula na mesma linha (ex: no modo Editar)
    if not mat_str and not cpf_str:
        # Verificar se termina com CPF formatado (000.000.000-00) ou número
        match_cpf_mask = re.match(r"^(.+?)[\s\-_;,]+(\d{3}\.?\d{3}\.?\d{3}[\-\.]?\d{2})$", nome_str)
        if match_cpf_mask:
            nome_str = match_cpf_mask.group(1).strip()
            cpf_str = limpar_cpf(match_cpf_mask.group(2))
        else:
            match_final = re.match(r"^(.+?)[\s\-_;,]+(\d{4,})$", nome_str)
            if match_final:
                nome_str = match_final.group(1).strip()
                num_final = match_final.group(2).strip()
                if len(num_final) == 11:
                    cpf_str = num_final
                else:
                    mat_str = num_final

    # Remover sufixos "não consta" caso tenham sido digitados junto ao nome
    nome_str = re.sub(
        r"[\s\-_;,]+(n[aã]o\s+consta|sem\s+matr[ií]cula|s/m|n/c)$",
        "",
        nome_str,
        flags=re.IGNORECASE,
    ).strip()

    partes_nome = extrair_primeiro_segundo_nome(nome_str)
    if not partes_nome:
        ident = mat_str or cpf_str
        return nome_str, ident, limpar_nome_arquivo(nome_str or ident)

    # PRIORIDADE 1: Primeiro e Segundo Nome + Matrícula
    if mat_str:
        formatado = "_".join(list(partes_nome) + [mat_str])
        return nome_str, mat_str, limpar_nome_arquivo(formatado)

    # PRIORIDADE 2: Primeiro e Segundo Nome + CPF
    if cpf_str:
        formatado = "_".join(list(partes_nome) + [cpf_str])
        return nome_str, cpf_str, limpar_nome_arquivo(formatado)

    # PRIORIDADE 3: Nome completo separado por '_' em minúsculas
    formatado_completo = formatar_nome_completo_underline(nome_str)
    return nome_str, "", formatado_completo


# ==============================================================================
# 2. LEITURA DE PLANILHA (.XLSX NATIVO + OPENPYXL + .CSV/.TSV) COM SUPORTE A CPF
# ==============================================================================

def _ler_xlsx_nativo(caminho: Path) -> List[List[str]]:
    """Lê a primeira aba de um arquivo .xlsx usando apenas bibliotecas padrão do Python."""
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(caminho, "r") as z:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            root_ss = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root_ss.findall("s:si", ns):
                texto = "".join((t.text or "") for t in si.findall(".//s:t", ns))
                shared_strings.append(texto)

        worksheets = sorted(
            [n for n in z.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        )
        if not worksheets:
            return []

        root_ws = ET.fromstring(z.read(worksheets[0]))
        linhas: List[List[str]] = []

        def col_para_idx(ref_celula: str) -> int:
            letras = "".join(ch for ch in ref_celula if ch.isalpha()).upper()
            idx = 0
            for ch in letras:
                idx = idx * 26 + (ord(ch) - ord("A") + 1)
            return max(idx - 1, 0)

        for row_el in root_ws.findall(".//s:row", ns):
            valores_linha: dict = {}
            max_col = -1
            for c_el in row_el.findall("s:c", ns):
                ref = c_el.attrib.get("r", "")
                col_idx = col_para_idx(ref) if ref else (max_col + 1)
                tipo = c_el.attrib.get("t", "")
                v_el = c_el.find("s:v", ns)
                is_el = c_el.find(".//s:t", ns)

                valor = ""
                if tipo == "s" and v_el is not None and v_el.text is not None:
                    idx_ss = int(v_el.text)
                    if 0 <= idx_ss < len(shared_strings):
                        valor = shared_strings[idx_ss]
                elif tipo == "inlineStr" and is_el is not None and is_el.text is not None:
                    valor = is_el.text
                elif v_el is not None and v_el.text is not None:
                    valor = v_el.text

                valores_linha[col_idx] = str(valor).strip()
                if col_idx > max_col:
                    max_col = col_idx

            if max_col >= 0:
                linha_lista = [valores_linha.get(c, "") for c in range(max_col + 1)]
                if any(celula != "" for celula in linha_lista):
                    linhas.append(linha_lista)

        return linhas


def _ler_linhas_brutas_planilha(caminho: Path) -> List[List[str]]:
    """Lê linhas de .xlsx, .csv, .tsv ou .txt."""
    sufixo = caminho.suffix.lower()

    if sufixo in (".csv", ".tsv", ".txt"):
        conteudo = ""
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                conteudo = caminho.read_text(encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        linhas_texto = [l for l in conteudo.splitlines() if l.strip()]
        if not linhas_texto:
            return []
        delimitador = "\t" if sufixo == ".tsv" else ","
        try:
            dialeto = csv.Sniffer().sniff(linhas_texto[0], delimiters=",;\t|")
            delimitador = dialeto.delimiter
        except Exception:
            if ";" in linhas_texto[0] and "," not in linhas_texto[0]:
                delimitador = ";"
            elif "\t" in linhas_texto[0]:
                delimitador = "\t"
        leitor = csv.reader(linhas_texto, delimiter=delimitador)
        return [[c.strip() for c in row] for row in leitor if any(c.strip() for c in row)]

    try:
        import openpyxl  # type: ignore
        wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
        ws = wb.active
        linhas = []
        for row in ws.iter_rows(values_only=True):
            linha_str = ["" if val is None else str(val).strip() for val in row]
            if any(c != "" for c in linha_str):
                linhas.append(linha_str)
        wb.close()
        return linhas
    except ImportError:
        return _ler_xlsx_nativo(caminho)


def _eh_linha_cabecalho(linha: List[str]) -> bool:
    celulas_nao_vazias = [c.strip().lower() for c in linha if c.strip()]
    if not celulas_nao_vazias:
        return True
    for c in celulas_nao_vazias:
        if c in CABECALHOS_CONHECIDOS:
            return True
    if len(celulas_nao_vazias) >= 2:
        col2 = celulas_nao_vazias[1]
        if not any(ch.isdigit() for ch in col2) and any(
            pal in col2 for pal in ("matr", "insc", "id", "cod", "cód", "num", "núm", "ra", "cpf")
        ):
            return True
    return False


def _detectar_colunas_nome_matricula_cpf(
    linhas: List[List[str]],
) -> Tuple[int, Optional[int], Optional[int], int]:
    """
    Retorna (idx_coluna_nome, idx_coluna_matricula, idx_coluna_cpf, linha_inicial_dados).
    """
    if not linhas:
        return 0, None, None, 0

    primeira = linhas[0]
    tem_cabecalho = _eh_linha_cabecalho(primeira)
    inicio = 1 if tem_cabecalho else 0

    col_nome: Optional[int] = None
    col_mat: Optional[int] = None
    col_cpf: Optional[int] = None

    if tem_cabecalho:
        for idx, texto in enumerate(primeira):
            t = texto.strip().lower()
            if col_nome is None and any(
                k in t for k in ("nome", "aluno", "discente", "estudante", "pessoa", "candidato")
            ):
                col_nome = idx
            elif col_cpf is None and ("cpf" in t or "doc" in t):
                col_cpf = idx
            elif col_mat is None and any(
                k in t for k in ("matr", "ra", "siape", "registro", "insc", "id")
            ):
                col_mat = idx

        if col_nome is not None and (col_mat is not None or col_cpf is not None):
            return col_nome, col_mat, col_cpf, inicio

    amostras = linhas[inicio : inicio + 15]
    if not amostras:
        return 0, None, None, inicio

    colunas_ativas = set()
    for l in amostras:
        for idx, val in enumerate(l):
            if val.strip():
                colunas_ativas.add(idx)

    cols_ordenadas = sorted(colunas_ativas)
    if len(cols_ordenadas) >= 3:
        c0, c1, c2 = cols_ordenadas[0], cols_ordenadas[1], cols_ordenadas[2]
        return c0, c1, c2, inicio
    elif len(cols_ordenadas) == 2:
        c0, c1 = cols_ordenadas[0], cols_ordenadas[1]
        c0_eh_num = sum(
            1 for l in amostras if len(l) > c0 and re.match(r"^[\d\.\-]+$", l[c0].strip())
        )
        c1_eh_num = sum(
            1 for l in amostras if len(l) > c1 and re.match(r"^[\d\.\-]+$", l[c1].strip())
        )
        if c0_eh_num > c1_eh_num:
            return c1, c0, None, inicio
        return c0, c1, None, inicio

    col_unica = cols_ordenadas[0] if cols_ordenadas else 0
    return col_unica, None, None, inicio


def carregar_planilha(planilha_path: Path) -> Tuple[List[ItemPlanilha], str]:
    """
    Carrega a planilha (.xlsx, .csv, .tsv) e retorna a lista de ItemPlanilha já formatados
    com suporte às 3 prioridades (Matrícula -> CPF -> nome_completo_underline).
    """
    if not planilha_path.exists():
        return [], f"Planilha '{planilha_path}' não encontrada."

    try:
        linhas = _ler_linhas_brutas_planilha(planilha_path)
    except Exception as e:
        return [], f"Erro ao ler planilha '{planilha_path}': {e}"

    if not linhas:
        return [], f"A planilha '{planilha_path}' está vazia."

    col_nome, col_mat, col_cpf, linha_inicio = _detectar_colunas_nome_matricula_cpf(linhas)
    itens: List[ItemPlanilha] = []

    for idx_linha, linha in enumerate(linhas[linha_inicio:], start=1):
        val_nome = linha[col_nome].strip() if len(linha) > col_nome else ""
        val_mat = (
            linha[col_mat].strip()
            if (col_mat is not None and len(linha) > col_mat)
            else ""
        )
        val_cpf = (
            linha[col_cpf].strip()
            if (col_cpf is not None and len(linha) > col_cpf)
            else ""
        )

        if not val_nome and not val_mat and not val_cpf:
            continue
        if val_nome.lower() in ("nan", "none"):
            continue

        mat_limpa = limpar_matricula(val_mat)
        cpf_limpo = limpar_cpf(val_cpf)

        # Se não havia coluna de CPF separada e o campo matrícula contém um CPF (com máscara), separar
        if not cpf_limpo and re.match(r"^\d{3}\.\d{3}\.\d{3}-\d{2}$", val_mat):
            cpf_limpo = limpar_cpf(val_mat)
            mat_limpa = ""

        nome_completo, ident_usado, nome_formatado = formatar_nome_e_matricula(
            val_nome, mat_limpa, cpf_limpo
        )
        if not nome_formatado:
            continue

        itens.append(
            ItemPlanilha(
                indice_original=idx_linha,
                nome_completo=nome_completo,
                matricula=mat_limpa,
                cpf=cpf_limpo,
                nome_formatado=nome_formatado,
            )
        )

    if not itens:
        return [], "Nenhum nome válido foi encontrado na planilha (verifique se ela não está vazia)."

    return itens, "OK"


# ==============================================================================
# 3. DETECÇÃO INTELIGENTE DE PDFs PENDENTES VS JÁ RENOMEADOS
# ==============================================================================

def _chave_ordenacao_natural(caminho: Path) -> Tuple:
    nome = caminho.stem
    partes = re.split(r"(\d+)", nome)
    chave = []
    for p in partes:
        if p.isdigit():
            chave.append((0, int(p)))
        else:
            chave.append((1, p.lower()))
    return tuple(chave)


def _carregar_destinos_do_log(log_path: Path) -> Set[str]:
    destinos: Set[str] = set()
    if not log_path.exists():
        return destinos
    try:
        for linha in log_path.read_text(encoding="utf-8", errors="ignore").splitlines():
            m = re.search(r"SUCESSO:\s*'(.*?)'\s*->\s*'(.*?)'", linha)
            if m:
                destinos.add(Path(m.group(2)).stem.lower())
            m_rev = re.search(r"REVERTIDO:\s*'(.*?)'\s*->\s*'(.*?)'", linha)
            if m_rev:
                destinos.discard(Path(m_rev.group(1)).stem.lower())
    except Exception:
        pass
    return destinos


def _eh_pdf_ja_renomeado(
    pdf_path: Path,
    stems_planilha_lower: Set[str],
    destinos_log_lower: Set[str],
) -> bool:
    stem = pdf_path.stem
    stem_lower = stem.lower()

    if stem.startswith(".") or re.search(r"_(parte|unido)\d*$", stem_lower):
        return False

    base_sem_contador = re.sub(r"_\d+$", "", stem_lower)
    if stem_lower in stems_planilha_lower or base_sem_contador in stems_planilha_lower:
        return True

    if stem_lower in destinos_log_lower:
        return True

    if re.match(r"^[A-Za-zÀ-ÿ]+_[A-Za-zÀ-ÿ]+_\d{4,}(_\d+)?$", stem):
        return True

    return False


def sincronizar_pdfs_e_planilha(
    pdfs_dir: Path,
    itens_planilha: List[ItemPlanilha],
    log_path: Path,
) -> Tuple[List[Path], List[Path], int]:
    if not pdfs_dir.exists():
        return [], [], 0

    todos_pdfs = [
        f
        for f in pdfs_dir.iterdir()
        if f.is_file() and f.suffix.lower() == ".pdf"
    ]

    stems_planilha_lower = {item.nome_formatado.lower() for item in itens_planilha}
    destinos_log_lower = _carregar_destinos_do_log(log_path)

    pdfs_pendentes: List[Path] = []
    pdfs_ja_renomeados: List[Path] = []
    stems_renomeados_existentes: Counter = Counter()
    mapa_arquivo_por_stem: Dict[str, str] = {}

    for pdf in todos_pdfs:
        if _eh_pdf_ja_renomeado(pdf, stems_planilha_lower, destinos_log_lower):
            pdfs_ja_renomeados.append(pdf)
            stem_low = pdf.stem.lower()
            stems_renomeados_existentes[stem_low] += 1
            mapa_arquivo_por_stem.setdefault(stem_low, pdf.name)
            base_low = re.sub(r"_\d+$", "", stem_low)
            if base_low != stem_low and base_low in stems_planilha_lower:
                stems_renomeados_existentes[base_low] += 1
                mapa_arquivo_por_stem.setdefault(base_low, pdf.name)
        else:
            pdfs_pendentes.append(pdf)

    pdfs_pendentes.sort(key=_chave_ordenacao_natural)
    pdfs_ja_renomeados.sort(key=_chave_ordenacao_natural)

    nomes_concluidos = 0
    usados_por_stem: Counter = Counter()
    for item in itens_planilha:
        chave = item.nome_formatado.lower()
        if usados_por_stem[chave] < stems_renomeados_existentes[chave]:
            item.concluido = True
            item.arquivo_gerado = mapa_arquivo_por_stem.get(chave, f"{item.nome_formatado}.pdf")
            usados_por_stem[chave] += 1
            nomes_concluidos += 1
        else:
            item.concluido = False

    return pdfs_pendentes, pdfs_ja_renomeados, nomes_concluidos


# ==============================================================================
# 4. OPERAÇÕES DE PDF (PyMuPDF / pypdf): DIVIDIR, JUNTAR, ANEXAR, GIRAR, CORTAR,
#    EXCLUIR PÁGINA, REORDENAR, AUTO-LIMPAR BORDAS PRETAS E DETECTAR BRANCAS
# ==============================================================================

def contar_paginas_pdf(caminho_pdf: Path) -> Optional[int]:
    """Retorna o número de páginas de um PDF usando PyMuPDF ou pypdf."""
    pymupdf = _obter_pymupdf()
    if pymupdf is not None:
        try:
            with pymupdf.open(str(caminho_pdf)) as doc:
                return len(doc)
        except Exception:
            pass

    for mod_name in ("pypdf", "PyPDF2"):
        try:
            mod = __import__(mod_name)
            reader = mod.PdfReader(str(caminho_pdf))
            return len(reader.pages)
        except ImportError:
            pass
        except Exception:
            break

    if shutil.which("qpdf"):
        try:
            res = subprocess.run(
                ["qpdf", "--show-npages", str(caminho_pdf)],
                capture_output=True,
                text=True,
                check=True,
            )
            return int(res.stdout.strip())
        except Exception:
            pass

    try:
        dados = caminho_pdf.read_bytes()
        paginas = re.findall(rb"/Type\s*/Page\b(?!s)", dados)
        if paginas:
            return len(paginas)
    except Exception:
        pass

    return None


def _instalar_pypdf_usuario() -> bool:
    print(f"\n{Cor.CIANO}📦 Instalando biblioteca 'pymupdf' / 'pypdf' (--user)...{Cor.RESET}")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pymupdf", "pypdf", "--user"])
        return True
    except Exception as e:
        print(f"{Cor.VERMELHO}Não foi possível instalar automaticamente: {e}{Cor.RESET}")
        return False


def dividir_pdf_em_intervalos(
    caminho_pdf: Path,
    intervalos: List[Tuple[int, int]],
) -> Tuple[List[Path], str]:
    """
    Divide `caminho_pdf` em múltiplos arquivos conforme `intervalos` (1-based, inclusivo).
    """
    pasta = caminho_pdf.parent
    stem = caminho_pdf.stem
    arquivos_saida: List[Path] = []
    for idx in range(1, len(intervalos) + 1):
        candidato = pasta / f"{stem}_parte{idx}.pdf"
        contador = 1
        while candidato.exists():
            candidato = pasta / f"{stem}_parte{idx}_{contador}.pdf"
            contador += 1
        arquivos_saida.append(candidato)

    # 1. PyMuPDF (preferencial: rápido e sem warnings)
    pymupdf = _obter_pymupdf()
    if pymupdf is not None:
        try:
            with pymupdf.open(str(caminho_pdf)) as doc:
                for (ini, fim), saida in zip(intervalos, arquivos_saida):
                    novo_doc = pymupdf.open()
                    novo_doc.insert_pdf(doc, from_page=ini - 1, to_page=fim - 1)
                    novo_doc.save(str(saida))
                    novo_doc.close()
            return arquivos_saida, "OK"
        except Exception as e:
            return [], f"Erro ao dividir com PyMuPDF: {e}"

    # 2. pypdf / PyPDF2
    for mod_name in ("pypdf", "PyPDF2"):
        try:
            mod = __import__(mod_name)
            reader = mod.PdfReader(str(caminho_pdf))
            for (ini, fim), saida in zip(intervalos, arquivos_saida):
                writer = mod.PdfWriter()
                for p in range(ini - 1, fim):
                    writer.add_page(reader.pages[p])
                with open(saida, "wb") as f_out:
                    writer.write(f_out)
            return arquivos_saida, "OK"
        except ImportError:
            pass
        except Exception as e:
            return [], f"Erro ao dividir com {mod_name}: {e}"

    return [], "Biblioteca de PDF não encontrada (instale com: pip install pymupdf pypdf)."


def juntar_arquivos_pdf(
    caminhos_entrada: List[Path],
    caminho_saida: Path,
) -> Tuple[bool, str]:
    """
    Une 2 ou mais arquivos PDF em `caminho_saida`.
    """
    if len(caminhos_entrada) < 2:
        return False, "Selecione pelo menos 2 arquivos PDF para juntar."

    pymupdf = _obter_pymupdf()
    if pymupdf is not None:
        try:
            doc_final = pymupdf.open()
            for p in caminhos_entrada:
                with pymupdf.open(str(p)) as doc_parte:
                    doc_final.insert_pdf(doc_parte)
            temp_out = caminho_saida.parent / f".tmp_merge_{caminho_saida.name}"
            doc_final.save(str(temp_out))
            doc_final.close()
            temp_out.replace(caminho_saida)
            return True, "OK"
        except Exception as e:
            return False, f"Erro ao juntar PDFs com PyMuPDF: {e}"

    for mod_name in ("pypdf", "PyPDF2"):
        try:
            mod = __import__(mod_name)
            writer = mod.PdfWriter()
            for p in caminhos_entrada:
                reader = mod.PdfReader(str(p))
                for page in reader.pages:
                    writer.add_page(page)
            temp_out = caminho_saida.parent / f".tmp_merge_{caminho_saida.name}"
            with open(temp_out, "wb") as f_out:
                writer.write(f_out)
            temp_out.replace(caminho_saida)
            return True, "OK"
        except ImportError:
            pass
        except Exception as e:
            return False, f"Erro ao juntar PDFs com {mod_name}: {e}"

    return False, "Instale 'pymupdf' ou 'pypdf' para juntar PDFs."


def _salvar_pdf_modificado_seguro(doc, caminho_pdf: Path):
    """Salva um documento PyMuPDF sobrescrevendo o original com segurança."""
    temp_path = caminho_pdf.parent / f".tmp_edit_{caminho_pdf.name}"
    doc.save(str(temp_path), deflate=True)
    doc.close()
    temp_path.replace(caminho_pdf)


def rotacionar_paginas_pdf(
    caminho_pdf: Path,
    angulo: int,
    pagina_alvo: Optional[int] = None,
) -> Tuple[bool, str]:
    """
    Gira páginas de `caminho_pdf` em `angulo` graus (90, -90, 180, 270).
    Se `pagina_alvo` for None ou 0, gira todas as páginas; senão gira apenas a página `pagina_alvo` (1-based).
    """
    pymupdf = _obter_pymupdf()
    if pymupdf is not None:
        try:
            doc = pymupdf.open(str(caminho_pdf))
            total = len(doc)
            if pagina_alvo and 1 <= pagina_alvo <= total:
                indices = [pagina_alvo - 1]
            else:
                indices = list(range(total))
            for idx in indices:
                page = doc[idx]
                page.set_rotation((page.rotation + angulo) % 360)
            _salvar_pdf_modificado_seguro(doc, caminho_pdf)
            alvo_txt = f"página {pagina_alvo}" if (pagina_alvo and 1 <= pagina_alvo <= total) else f"todas as {total} página(s)"
            return True, f"Rotação de {angulo}° aplicada em {alvo_txt}."
        except Exception as e:
            return False, f"Erro ao girar PDF: {e}"

    for mod_name in ("pypdf", "PyPDF2"):
        try:
            mod = __import__(mod_name)
            reader = mod.PdfReader(str(caminho_pdf))
            writer = mod.PdfWriter()
            total = len(reader.pages)
            for idx, page in enumerate(reader.pages):
                if not pagina_alvo or pagina_alvo == idx + 1:
                    page.rotate(angulo)
                writer.add_page(page)
            temp_path = caminho_pdf.parent / f".tmp_rot_{caminho_pdf.name}"
            with open(temp_path, "wb") as f_out:
                writer.write(f_out)
            temp_path.replace(caminho_pdf)
            return True, f"Rotação de {angulo}° aplicada."
        except ImportError:
            pass
        except Exception as e:
            return False, f"Erro ao girar PDF: {e}"

    return False, "Biblioteca PyMuPDF/pypdf não encontrada."


def excluir_paginas_pdf(
    caminho_pdf: Path,
    paginas_1based: List[int],
) -> Tuple[bool, str]:
    """
    Remove uma ou mais páginas (1-based) do arquivo PDF.
    Não permite remover todas as páginas do documento.
    """
    pymupdf = _obter_pymupdf()
    if pymupdf is None:
        return False, "Instale 'pymupdf' para excluir páginas."
    try:
        doc = pymupdf.open(str(caminho_pdf))
        total = len(doc)
        indices_remover = sorted({p - 1 for p in paginas_1based if 1 <= p <= total}, reverse=True)
        if not indices_remover:
            doc.close()
            return False, "Nenhuma página válida selecionada para excluir."
        if len(indices_remover) >= total:
            doc.close()
            return False, "O PDF possui apenas 1 página (não é possível deixar o PDF com 0 páginas)."
        for idx in indices_remover:
            doc.delete_page(idx)
        _salvar_pdf_modificado_seguro(doc, caminho_pdf)
        return True, f"Página(s) {', '.join(str(i + 1) for i in reversed(indices_remover))} removida(s) com sucesso."
    except Exception as e:
        return False, f"Erro ao excluir página: {e}"


def reordenar_paginas_pdf(
    caminho_pdf: Path,
    nova_ordem_1based: Optional[List[int]] = None,
    inverter: bool = False,
) -> Tuple[bool, str]:
    """
    Inverte a ordem de todas as páginas ou aplica `nova_ordem_1based` (ex: [2, 1, 3]).
    """
    pymupdf = _obter_pymupdf()
    if pymupdf is None:
        return False, "Instale 'pymupdf' para reordenar páginas."
    try:
        doc = pymupdf.open(str(caminho_pdf))
        total = len(doc)
        if total < 2:
            doc.close()
            return False, "O PDF possui apenas 1 página."

        if inverter or not nova_ordem_1based:
            ordem_0based = list(reversed(range(total)))
            desc = "Ordem das páginas invertida"
        else:
            ordem_0based = [p - 1 for p in nova_ordem_1based if 1 <= p <= total]
            if not ordem_0based:
                doc.close()
                return False, "Ordem de páginas inválida."
            desc = f"Páginas reordenadas para {', '.join(str(i + 1) for i in ordem_0based)}"

        doc.select(ordem_0based)
        _salvar_pdf_modificado_seguro(doc, caminho_pdf)
        return True, f"{desc} ({len(ordem_0based)} páginas)."
    except Exception as e:
        return False, f"Erro ao reordenar páginas: {e}"


def recortar_pagina_pdf(
    caminho_pdf: Path,
    x0_rel: float,
    y0_rel: float,
    x1_rel: float,
    y1_rel: float,
    pagina_alvo: Optional[int] = None,
) -> Tuple[bool, str]:
    """
    Aplica um recorte (CropBox) baseado em coordenadas relativas (0.0 a 1.0)
    vistas pelo usuário na imagem renderizada da página.
    Suporta páginas com ou sem rotação prévia!
    """
    pymupdf = _obter_pymupdf()
    if pymupdf is None:
        return False, "Instale 'pymupdf' para recortar páginas."

    x0_r = max(0.0, min(1.0, min(x0_rel, x1_rel)))
    x1_r = max(0.0, min(1.0, max(x0_rel, x1_rel)))
    y0_r = max(0.0, min(1.0, min(y0_rel, y1_rel)))
    y1_r = max(0.0, min(1.0, max(y0_rel, y1_rel)))

    if (x1_r - x0_r) < 0.03 or (y1_r - y0_r) < 0.03:
        return False, "Área de recorte muito pequena. Selecione uma região maior."

    try:
        doc = pymupdf.open(str(caminho_pdf))
        total = len(doc)
        indices = [pagina_alvo - 1] if (pagina_alvo and 1 <= pagina_alvo <= total) else list(range(total))

        for idx in indices:
            page = doc[idx]
            vis_rect = page.rect
            vx0 = vis_rect.x0 + x0_r * vis_rect.width
            vy0 = vis_rect.y0 + y0_r * vis_rect.height
            vx1 = vis_rect.x0 + x1_r * vis_rect.width
            vy1 = vis_rect.y0 + y1_r * vis_rect.height
            novo_vis = pymupdf.Rect(vx0, vy0, vx1, vy1)
            # Converter do espaço visual da página para o espaço da CropBox
            novo_crop = novo_vis * page.derotation_matrix
            novo_crop.normalize()
            page.set_cropbox(novo_crop & page.mediabox)

        _salvar_pdf_modificado_seguro(doc, caminho_pdf)
        alvo_txt = f"página {pagina_alvo}" if (pagina_alvo and 1 <= pagina_alvo <= total) else f"todas as {total} página(s)"
        return True, f"Recorte aplicado em {alvo_txt}."
    except Exception as e:
        return False, f"Erro ao recortar PDF: {e}"


def auto_recortar_bordas_pretas(
    caminho_pdf: Path,
    pagina_alvo: Optional[int] = None,
) -> Tuple[bool, str]:
    """
    Detecta e remove automaticamente as "bordinhas pretas" / sombras escuras de scanner
    nas 4 extremidades da folha usando visão computacional leve (PyMuPDF + numpy).
    """
    pymupdf = _obter_pymupdf()
    if pymupdf is None:
        return False, "Instale 'pymupdf' para remover bordas pretas."

    try:
        import numpy as np  # type: ignore
    except ImportError:
        return False, "Biblioteca 'numpy' necessária para auto-recorte."

    try:
        doc = pymupdf.open(str(caminho_pdf))
        total = len(doc)
        indices = [pagina_alvo - 1] if (pagina_alvo and 1 <= pagina_alvo <= total) else list(range(total))
        paginas_ajustadas = 0

        for idx in indices:
            page = doc[idx]
            pix = page.get_pixmap(dpi=100, colorspace=pymupdf.csGRAY, alpha=False)
            h, w = pix.height, pix.width
            if h < 40 or w < 40:
                continue

            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape((h, w))

            # Uma linha/coluna de borda de scanner costuma ser escura (< 165) nas extremidades
            # ou ter blocos pretos (< 85) nos cantos.
            # Varremos de fora para dentro até encontrar o início consistente do papel claro (> 185).
            max_scan_y = int(h * 0.22)
            max_scan_x = int(w * 0.22)

            top = 0
            for y in range(max_scan_y):
                linha = arr[y, :]
                frac_escura = float(np.mean(linha < 150))
                frac_muito_preta = float(np.mean(linha < 70))
                if frac_escura > 0.12 or frac_muito_preta > 0.04:
                    top = y + 1
                elif top > 0 and float(np.mean(linha > 200)) > 0.88:
                    break

            bottom = h
            for y in range(h - 1, h - 1 - max_scan_y, -1):
                linha = arr[y, :]
                frac_escura = float(np.mean(linha < 150))
                frac_muito_preta = float(np.mean(linha < 70))
                if frac_escura > 0.12 or frac_muito_preta > 0.04:
                    bottom = y
                elif bottom < h and float(np.mean(linha > 200)) > 0.88:
                    break

            left = 0
            for x in range(max_scan_x):
                col = arr[top:bottom, x] if bottom > top else arr[:, x]
                frac_escura = float(np.mean(col < 150))
                frac_muito_preta = float(np.mean(col < 70))
                if frac_escura > 0.12 or frac_muito_preta > 0.04:
                    left = x + 1
                elif left > 0 and float(np.mean(col > 200)) > 0.88:
                    break

            right = w
            for x in range(w - 1, w - 1 - max_scan_x, -1):
                col = arr[top:bottom, x] if bottom > top else arr[:, x]
                frac_escura = float(np.mean(col < 150))
                frac_muito_preta = float(np.mean(col < 70))
                if frac_escura > 0.12 or frac_muito_preta > 0.04:
                    right = x
                elif right < w and float(np.mean(col > 200)) > 0.88:
                    break

            # Adicionar margem interna de segurança de 2 pixels se detectou borda preta
            if top > 0:
                top = min(top + 2, max_scan_y)
            if bottom < h:
                bottom = max(bottom - 2, h - max_scan_y)
            if left > 0:
                left = min(left + 2, max_scan_x)
            if right < w:
                right = max(right - 2, w - max_scan_x)

            if top > 1 or left > 1 or bottom < (h - 1) or right < (w - 1):
                vis_rect = page.rect
                vx0 = vis_rect.x0 + (left / float(w)) * vis_rect.width
                vy0 = vis_rect.y0 + (top / float(h)) * vis_rect.height
                vx1 = vis_rect.x0 + (right / float(w)) * vis_rect.width
                vy1 = vis_rect.y0 + (bottom / float(h)) * vis_rect.height
                novo_vis = pymupdf.Rect(vx0, vy0, vx1, vy1)
                novo_crop = novo_vis * page.derotation_matrix
                novo_crop.normalize()
                page.set_cropbox(novo_crop & page.mediabox)
                paginas_ajustadas += 1

        if paginas_ajustadas == 0:
            doc.close()
            return False, "Nenhuma borda preta significativa foi detectada nas margens."

        _salvar_pdf_modificado_seguro(doc, caminho_pdf)
        return True, f"Bordas pretas removidas automaticamente em {paginas_ajustadas} página(s)!"
    except Exception as e:
        return False, f"Erro ao limpar bordas pretas: {e}"


def auto_orientar_pdf(caminho_pdf: Path) -> Tuple[bool, str]:
    """
    Detecta se as páginas estão com o texto na vertical (90°/270°) ou de cabeça para baixo (180°)
    usando análise de perfil de projeção de linhas de texto + Tesseract OSD se disponível.
    """
    pymupdf = _obter_pymupdf()
    if pymupdf is None:
        return False, "Instale 'pymupdf' para auto-orientar páginas."

    try:
        import numpy as np  # type: ignore
    except ImportError:
        np = None

    try:
        doc = pymupdf.open(str(caminho_pdf))
        ajustadas = 0

        for page in doc:
            angulo_correcao = 0

            # 1. Tentar Tesseract OSD se disponível no sistema para detectar 90/180/270 com precisão
            if shutil.which("tesseract"):
                try:
                    pix = page.get_pixmap(dpi=100, colorspace=pymupdf.csGRAY, alpha=False)
                    png_bytes = pix.tobytes("png")
                    res = subprocess.run(
                        ["tesseract", "stdin", "stdout", "--psm", "0"],
                        input=png_bytes,
                        capture_output=True,
                        timeout=2.5,
                    )
                    osd_txt = (res.stdout.decode("utf-8", errors="ignore") + res.stderr.decode("utf-8", errors="ignore"))
                    m_rot = re.search(r"Rotate:\s*(\d+)", osd_txt)
                    m_conf = re.search(r"Orientation confidence:\s*([\d\.]+)", osd_txt)
                    if m_rot:
                        rot = int(m_rot.group(1))
                        conf = float(m_conf.group(1)) if m_conf else 2.0
                        if rot in (90, 180, 270) and conf >= 1.2:
                            angulo_correcao = rot
                except Exception:
                    pass

            # 2. Fallback rápido por Projeção Horizontal vs Vertical (detecta páginas deitadas em 0.02s)
            if angulo_correcao == 0 and np is not None:
                pix = page.get_pixmap(dpi=72, colorspace=pymupdf.csGRAY, alpha=False)
                h, w = pix.height, pix.width
                if h > 50 and w > 50:
                    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape((h, w))
                    # Recortar centro (ignorar bordas)
                    miolo = arr[int(h * 0.1) : int(h * 0.9), int(w * 0.1) : int(w * 0.9)]
                    tinta = (miolo < 170).astype(np.float32)
                    # Variância das médias das linhas vs colunas: linhas horizontais de texto geram alta variância horizontal
                    var_horiz = float(np.var(np.mean(tinta, axis=1)))
                    var_vert = float(np.var(np.mean(tinta, axis=0)))
                    if var_vert > var_horiz * 1.65 and var_vert > 0.0005:
                        angulo_correcao = 90

            if angulo_correcao != 0:
                page.set_rotation((page.rotation + angulo_correcao) % 360)
                ajustadas += 1

        if ajustadas == 0:
            doc.close()
            return False, "As páginas já parecem estar na orientação correta."

        _salvar_pdf_modificado_seguro(doc, caminho_pdf)
        return True, f"Orientação corrigida automaticamente em {ajustadas} página(s)!"
    except Exception as e:
        return False, f"Erro ao auto-orientar PDF: {e}"


def detectar_paginas_em_branco(caminho_pdf: Path) -> List[int]:
    """
    Retorna a lista de páginas (1-based) que parecem estar em branco (>99.2% claras sem texto).
    Muito útil quando o scanner escaneia frente e verso de folha simples.
    """
    pymupdf = _obter_pymupdf()
    if pymupdf is None:
        return []
    try:
        import numpy as np  # type: ignore
    except ImportError:
        return []

    em_branco: List[int] = []
    try:
        with pymupdf.open(str(caminho_pdf)) as doc:
            if len(doc) < 2:
                return []
            for idx, page in enumerate(doc):
                if len(page.get_text("text").strip()) > 10:
                    continue
                pix = page.get_pixmap(dpi=36, colorspace=pymupdf.csGRAY, alpha=False)
                h, w = pix.height, pix.width
                if h < 20 or w < 20:
                    continue
                arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape((h, w))
                # Ignorar 8% das bordas (caso haja bordinha preta do scanner na página em branco)
                miolo = arr[int(h * 0.08) : int(h * 0.92), int(w * 0.08) : int(w * 0.92)]
                frac_escura = float(np.mean(miolo < 185))
                if frac_escura < 0.004:  # Menos de 0.4% de pixels escuros no miolo
                    em_branco.append(idx + 1)
    except Exception:
        pass
    return em_branco


def renderizar_pagina_png(caminho_pdf: Path, pagina_1based: int = 1, dpi: int = 110) -> Optional[bytes]:
    """Renderiza uma página do PDF como imagem PNG para o editor de recorte visual da GUI."""
    pymupdf = _obter_pymupdf()
    if pymupdf is None:
        return None
    try:
        with pymupdf.open(str(caminho_pdf)) as doc:
            if len(doc) == 0:
                return None
            idx = max(0, min(len(doc) - 1, pagina_1based - 1))
            pix = doc[idx].get_pixmap(dpi=dpi, alpha=False)
            return pix.tobytes("png")
    except Exception:
        return None


# ==============================================================================
# 5. "IA LOCAL" EVOLUTIVA: LEITURA DO PDF, MEMÓRIA DE TREINAMENTO CONTÍNUO
#    (memoria_ia.json) E DIAGNÓSTICO INTELIGENTE DE CAUSA DE ERROS
# ==============================================================================

def _caminho_memoria_ia_padrao(caminho_pdf: Optional[Path] = None) -> Path:
    if caminho_pdf is not None:
        pasta_projeto = caminho_pdf.resolve().parent.parent
        if pasta_projeto.exists():
            return pasta_projeto / ARQUIVO_MEMORIA_IA
    return Path(ARQUIVO_MEMORIA_IA).resolve()


def carregar_memoria_ia(caminho_memoria: Optional[Path] = None) -> dict:
    """
    Carrega a base de conhecimento local da IA (`memoria_ia.json`).
    Funciona 100% offline e leve:
    - `variacoes_ocr`: mapeia palavras lidas com ruído pelo OCR para a grafia correta (ex: 'gustawo' -> 'gustavo')
    - `padroes_alunos`: mapeia tokens/assinaturas específicas aprendidas de documentos difíceis para o aluno correto
    - `termos_ignorados`: nomes de coordenadores, diretores, secretários ou cabeçalhos que a IA deve ignorar
    - `total_confirmacoes`: contador de quantos documentos já treinaram a IA
    - `ultimos_aprendizados`: histórico recente do que a IA aprendeu
    """
    alvo = caminho_memoria or _caminho_memoria_ia_padrao()
    estrutura_padrao = {
        "total_confirmacoes": 0,
        "variacoes_ocr": {},
        "padroes_alunos": {},
        "termos_ignorados": [],
        "ultimos_aprendizados": [],
    }
    if not alvo.exists():
        return estrutura_padrao
    try:
        dados = json.loads(alvo.read_text(encoding="utf-8"))
        if not isinstance(dados, dict):
            return estrutura_padrao
        for k, v in estrutura_padrao.items():
            if k not in dados or not isinstance(dados[k], type(v)):
                dados[k] = v
        return dados
    except Exception:
        return estrutura_padrao


def salvar_memoria_ia(dados: dict, caminho_memoria: Optional[Path] = None) -> bool:
    alvo = caminho_memoria or _caminho_memoria_ia_padrao()
    try:
        alvo.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
        return True
    except Exception:
        return False


def adicionar_termo_ignorado_ia(termo: str, caminho_memoria: Optional[Path] = None) -> Tuple[bool, str]:
    """Ensina a IA a ignorar um nome recorrente (ex: diretor, coordenador, secretaria, familiar)."""
    termo_limpo = normalizar_busca(termo).strip()
    if len(termo_limpo) < 3:
        return False, "Digite um nome ou termo com pelo menos 3 letras para ignorar."
    mem = carregar_memoria_ia(caminho_memoria)
    ignorados: List[str] = mem.get("termos_ignorados", [])
    if termo_limpo not in ignorados:
        ignorados.append(termo_limpo)
        mem["termos_ignorados"] = ignorados
        historico = mem.get("ultimos_aprendizados", [])
        historico.insert(0, f"Termo ignorado adicionado: '{termo.strip()}'")
        mem["ultimos_aprendizados"] = historico[:25]
        salvar_memoria_ia(mem, caminho_memoria)
        return True, f"A IA agora ignorará '{termo.strip()}' durante a leitura dos PDFs."
    return False, f"O termo '{termo.strip()}' já está na lista de ignorados da IA."


def remover_termo_ignorado_ia(termo: str, caminho_memoria: Optional[Path] = None) -> Tuple[bool, str]:
    termo_limpo = normalizar_busca(termo).strip()
    mem = carregar_memoria_ia(caminho_memoria)
    ignorados: List[str] = mem.get("termos_ignorados", [])
    novos = [t for t in ignorados if normalizar_busca(t).strip() != termo_limpo]
    if len(novos) == len(ignorados):
        return False, "Termo não encontrado na memória da IA."
    mem["termos_ignorados"] = novos
    salvar_memoria_ia(mem, caminho_memoria)
    return True, f"Termo '{termo}' removido da lista de ignorados."


def registrar_aprendizado_ia(
    caminho_pdf: Path,
    item_confirmado: ItemPlanilha,
    caminho_memoria: Optional[Path] = None,
    forcar_padrao_documento: bool = False,
) -> Tuple[int, str]:
    """
    Treina a IA local com base na confirmação/correção do usuário:
    1) Aprende variações de OCR (ex: quando o scanner/Tesseract troca letras do nome do aluno).
    2) Se `forcar_padrao_documento=True` (ou quando o nome estava difícil de ler), memoriza
       palavras-chave/identificadores daquele documento associando ao aluno confirmado.
    Retorna (qtd_novos_padroes_aprendidos, resumo_aprendizado).
    """
    alvo_mem = caminho_memoria or _caminho_memoria_ia_padrao(caminho_pdf)
    mem = carregar_memoria_ia(alvo_mem)
    mem["total_confirmacoes"] = int(mem.get("total_confirmacoes", 0)) + 1

    texto_bruto, _ = extrair_texto_pdf_local(caminho_pdf, usar_ocr_se_imagem=True, pagina_1based=1)
    if not texto_bruto:
        salvar_memoria_ia(mem, alvo_mem)
        return 0, "Confirmação contabilizada (PDF sem camada de texto para extrair padrões)."

    texto_norm = _remover_linhas_filiacao(normalizar_busca(texto_bruto))
    palavras_pdf = set(re.findall(r"[a-z0-9]{3,}", texto_norm))
    nome_aluno_norm = normalizar_busca(item_confirmado.nome_completo)
    palavras_aluno = [
        p for p in nome_aluno_norm.split() if len(p) >= 4 and p not in PARTICULAS_IGNORADAS
    ]

    variacoes: Dict[str, str] = mem.get("variacoes_ocr", {})
    padroes_alunos: Dict[str, List[str]] = mem.get("padroes_alunos", {})
    historico: List[str] = mem.get("ultimos_aprendizados", [])
    novos_aprendizados = 0
    detalhes: List[str] = []

    # 1. Aprender correções de grafia/ruído de OCR para as palavras do nome do aluno
    for pal_real in palavras_aluno:
        if pal_real in palavras_pdf:
            continue
        melhor_token = ""
        melhor_sim = 0.0
        for tok in palavras_pdf:
            if not tok.isalpha() or abs(len(tok) - len(pal_real)) > 2:
                continue
            sim = difflib.SequenceMatcher(None, pal_real, tok).ratio()
            if sim > melhor_sim:
                melhor_sim = sim
                melhor_token = tok
        if 0.66 <= melhor_sim < 0.99 and melhor_token:
            if variacoes.get(melhor_token) != pal_real:
                variacoes[melhor_token] = pal_real
                novos_aprendizados += 1
                msg = f"OCR '{melhor_token}' -> '{pal_real}' ({item_confirmado.nome_formatado})"
                detalhes.append(msg)
                historico.insert(0, msg)

    # 2. Se o usuário clicou em "Ensinar IA" (ou o documento tinha baixa leitura direta),
    #    memorizar tokens distintivos do cabeçalho vinculados a este aluno
    acertos_diretos = sum(1 for p in palavras_aluno if p in palavras_pdf)
    if forcar_padrao_documento or (palavras_aluno and acertos_diretos < max(1, len(palavras_aluno) // 2)):
        stopwords_doc = {
            "universidade", "instituto", "federal", "ministerio", "educacao", "secretaria",
            "historico", "escolar", "declaracao", "certidao", "certificado", "diploma",
            "aluno", "discente", "nome", "matricula", "cpf", "data", "nascimento", "curso",
            "campus", "periodo", "semestre", "assinatura", "diretor", "coordenador", "pagina",
            "brasil", "estado", "republica", "documento", "identidade", "orgao", "expedidor",
        } | PARTICULAS_IGNORADAS
        tokens_distintivos = []
        for tok in re.findall(r"[a-z0-9]{4,}", texto_norm):
            if tok not in stopwords_doc and tok not in tokens_distintivos:
                tokens_distintivos.append(tok)
            if len(tokens_distintivos) >= 8:
                break

        if tokens_distintivos:
            chave_aluno = nome_aluno_norm
            atuais = set(padroes_alunos.get(chave_aluno, []))
            antes = len(atuais)
            atuais.update(tokens_distintivos[:6])
            padroes_alunos[chave_aluno] = sorted(atuais)[:18]
            if len(atuais) > antes:
                novos_aprendizados += 1
                msg = f"Padrão de documento memorizado p/ '{item_confirmado.nome_completo}'"
                detalhes.append(msg)
                historico.insert(0, msg)

    mem["variacoes_ocr"] = variacoes
    mem["padroes_alunos"] = padroes_alunos
    mem["ultimos_aprendizados"] = historico[:25]
    salvar_memoria_ia(mem, alvo_mem)

    if detalhes:
        return novos_aprendizados, "IA aprendeu: " + "; ".join(detalhes)
    return 0, f"IA registrou a confirmação de '{item_confirmado.nome_formatado}' na base de experiência."


def extrair_texto_pdf_local(
    caminho_pdf: Path,
    usar_ocr_se_imagem: bool = True,
    pagina_1based: int = 1,
) -> Tuple[str, str]:
    """
    Extrai texto de uma página do PDF (`pagina_1based`, padrão=1) de forma 100% local e rápida:
    1) Tenta ler camada de texto embutida (0.005s).
    2) Se estiver vazio (PDF escaneado como imagem pura) e `usar_ocr_se_imagem=True`,
        recorta os 45% superiores da página (onde ficam cabeçalho/nome/matrícula/CPF)
        e roda Tesseract OCR local (se disponível) em ~0.3s!
    Retorna (texto_extraido, metodo_usado).
    """
    pymupdf = _obter_pymupdf()
    if pymupdf is None or not caminho_pdf.exists():
        return "", "indisponivel"

    try:
        with pymupdf.open(str(caminho_pdf)) as doc:
            if len(doc) == 0:
                return "", "vazio"
            idx_pag = max(0, min(len(doc) - 1, pagina_1based - 1))
            page = doc[idx_pag]
            texto_nativo = page.get_text("text").strip()
            if len(texto_nativo) >= 15:
                return texto_nativo, "texto_embutido"

            if not usar_ocr_se_imagem or not shutil.which("tesseract"):
                return texto_nativo, "sem_ocr"

            # Truque para PC fraco: renderizar primeiro os 45% do topo da página em 140 DPI
            rect = page.rect
            clip_topo = pymupdf.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * 0.45)
            pix = page.get_pixmap(dpi=140, colorspace=pymupdf.csGRAY, clip=clip_topo, alpha=False)
            png_bytes = pix.tobytes("png")

            idiomas = "por+eng"
            res = subprocess.run(
                ["tesseract", "stdin", "stdout", "-l", idiomas, "--psm", "6"],
                input=png_bytes,
                capture_output=True,
                timeout=3.0,
            )
            if res.returncode != 0:
                res = subprocess.run(
                    ["tesseract", "stdin", "stdout", "--psm", "6"],
                    input=png_bytes,
                    capture_output=True,
                    timeout=3.0,
                )
            texto_ocr = res.stdout.decode("utf-8", errors="ignore").strip()

            # Se o topo tinha pouquíssimo texto, tenta a página inteira em 115 DPI
            if len(texto_ocr) < 20 and pagina_1based == 1:
                pix_full = page.get_pixmap(dpi=115, colorspace=pymupdf.csGRAY, alpha=False)
                res_full = subprocess.run(
                    ["tesseract", "stdin", "stdout", "--psm", "6"],
                    input=pix_full.tobytes("png"),
                    capture_output=True,
                    timeout=4.0,
                )
                texto_ocr = res_full.stdout.decode("utf-8", errors="ignore").strip()

            return texto_ocr, "ocr_local"
    except Exception:
        return "", "erro"


def _detectar_suspeita_pagina_virada(caminho_pdf: Path) -> bool:
    """
    Verifica rapidamente (~0.015s) se a 1ª página parece estar escaneada deitada (90°/270°),
    causando falha no OCR tradicional.
    """
    pymupdf = _obter_pymupdf()
    if pymupdf is None:
        return False
    try:
        import numpy as np  # type: ignore
        with pymupdf.open(str(caminho_pdf)) as doc:
            if len(doc) == 0:
                return False
            page = doc[0]
            pix = page.get_pixmap(dpi=60, colorspace=pymupdf.csGRAY, alpha=False)
            h, w = pix.height, pix.width
            if h < 40 or w < 40:
                return False
            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape((h, w))
            miolo = arr[int(h * 0.12) : int(h * 0.88), int(w * 0.12) : int(w * 0.88)]
            tinta = (miolo < 170).astype(np.float32)
            var_horiz = float(np.var(np.mean(tinta, axis=1)))
            var_vert = float(np.var(np.mean(tinta, axis=0)))
            return bool(var_vert > var_horiz * 1.55 and var_vert > 0.0004)
    except Exception:
        return False


def _remover_linhas_filiacao(texto_norm: str, termos_ignorados_ia: Optional[List[str]] = None) -> str:
    """
    Filtra linhas ou trechos que claramente se referem a familiares (pai, mãe, filiação,
    genitor, responsável, cônjuge) dentro de um dossiê de aluno, além de termos/nomes
    institucionais que o usuário ensinou a IA a ignorar.
    """
    padroes_familia = (
        r"\b(pai|mae|filiacao|genitor|genitora|responsavel legal|nome do pai|nome da mae|conjuge)\s*[:\-]"
    )
    linhas_filtradas = []
    for linha in texto_norm.splitlines():
        if re.search(padroes_familia, linha):
            continue
        linha_limpa = linha
        if termos_ignorados_ia:
            for termo in termos_ignorados_ia:
                if termo and len(termo) >= 3 and termo in linha_limpa:
                    linha_limpa = linha_limpa.replace(termo, " ")
        linhas_filtradas.append(linha_limpa)
    return "\n".join(linhas_filtradas)


def _avaliar_texto_contra_itens(
    texto_bruto: str,
    itens_planilha: List[ItemPlanilha],
    idx_nome_atual: Optional[int],
    mem_ia: dict,
) -> Tuple[List[dict], Optional[dict], bool, str]:
    """
    Avalia um texto extraído contra todos os itens da planilha usando também a Memória da IA.
    Retorna (candidatos_fortes, melhor_parcial, suspeita_familiar_atual, trecho_limpo).
    """
    termos_ignorados = mem_ia.get("termos_ignorados", [])
    variacoes_ocr: Dict[str, str] = mem_ia.get("variacoes_ocr", {})
    padroes_alunos: Dict[str, List[str]] = mem_ia.get("padroes_alunos", {})

    texto_norm_completo = normalizar_busca(texto_bruto)
    texto_norm = _remover_linhas_filiacao(texto_norm_completo, termos_ignorados)
    texto_sem_pontuacao = re.sub(r"[\.\-]", "", texto_norm)

    palavras_brutas = set(re.findall(r"[a-z0-9]{3,}", texto_norm))
    # Aplicar correções aprendidas pela IA (ex: 'gustawo' -> 'gustavo')
    palavras_pdf = set(palavras_brutas)
    for tok in list(palavras_brutas):
        if tok in variacoes_ocr:
            corrigida = variacoes_ocr[tok]
            palavras_pdf.add(corrigida)
            texto_norm += f" {corrigida}"

    linhas_norm = [l.strip() for l in texto_norm.splitlines() if len(l.strip()) >= 4]
    trecho_limpo = " | ".join(linhas_norm[:3])[:140]

    candidatos_fortes: List[dict] = []
    melhor_parcial: Optional[dict] = None
    suspeita_familiar_atual = False

    for idx, item in enumerate(itens_planilha):
        score_item = 0.0
        motivo_item = ""
        usou_memoria = False

        # 1. Conferir Matrícula do aluno
        if item.matricula and len(item.matricula) >= 5 and item.matricula in texto_norm.replace(" ", ""):
            score_item = 0.99
            motivo_item = f"Matrícula {item.matricula} confirmada"

        # 2. Conferir CPF
        if score_item < 0.95 and item.cpf and len(item.cpf) >= 9:
            if item.cpf in texto_sem_pontuacao:
                score_item = 0.96
                motivo_item = f"CPF {item.cpf} confirmado"

        # 3. Conferir Nome Completo do Aluno + Memória da IA
        nome_aluno_norm = normalizar_busca(item.nome_completo)
        palavras_aluno = [
            p for p in nome_aluno_norm.split() if len(p) >= 3 and p not in PARTICULAS_IGNORADAS
        ]
        if palavras_aluno:
            primeiro_nome = palavras_aluno[0]
            tem_primeiro_nome = (
                primeiro_nome in palavras_pdf
                or any(
                    difflib.SequenceMatcher(None, primeiro_nome, w).ratio() >= 0.86
                    for w in palavras_pdf
                    if abs(len(w) - len(primeiro_nome)) <= 1
                )
            )

            if nome_aluno_norm in texto_norm and len(nome_aluno_norm) >= 8:
                if score_item < 0.95:
                    score_item = 0.95
                    motivo_item = "Nome completo do aluno encontrado"
            else:
                acertos_palavras = 0.0
                for pal in palavras_aluno:
                    if pal in palavras_pdf:
                        acertos_palavras += 1.0
                    else:
                        melhor_pal = max(
                            (
                                difflib.SequenceMatcher(None, pal, w).ratio()
                                for w in palavras_pdf
                                if abs(len(w) - len(pal)) <= 2
                            ),
                            default=0.0,
                        )
                        if melhor_pal >= 0.84:
                            acertos_palavras += melhor_pal

                proporcao_nome = acertos_palavras / float(len(palavras_aluno))

                # Verificar se a Memória da IA conhece padrões específicos deste aluno
                tokens_mem_aluno = padroes_alunos.get(nome_aluno_norm, [])
                bonus_memoria = 0.0
                if tokens_mem_aluno:
                    matches_mem = sum(1 for t in tokens_mem_aluno if t in palavras_pdf)
                    if matches_mem >= 3:
                        bonus_memoria = min(0.32, matches_mem * 0.08)
                        usou_memoria = True

                if not tem_primeiro_nome and bonus_memoria < 0.20:
                    if idx == idx_nome_atual and proporcao_nome >= 0.5:
                        suspeita_familiar_atual = True
                    score_nome = proporcao_nome * 0.45
                else:
                    segunda_ok = (
                        len(palavras_aluno) < 2
                        or palavras_aluno[1] in palavras_pdf
                        or any(
                            difflib.SequenceMatcher(None, palavras_aluno[1], w).ratio() >= 0.84
                            for w in palavras_pdf
                        )
                    )
                    score_nome = min(0.94, proporcao_nome * (0.93 if segunda_ok else 0.72) + bonus_memoria)

                if score_nome > score_item:
                    score_item = score_nome
                    sufixo_mem = " + Memória IA" if usou_memoria else ""
                    motivo_item = f"Nome compatível ({int(round(score_nome * 100))}%{sufixo_mem})"

        info_cand = {
            "idx": idx,
            "indice_original": item.indice_original,
            "nome_completo": item.nome_completo,
            "nome_formatado": item.nome_formatado,
            "concluido": item.concluido,
            "score": round(score_item * 100, 1),
            "motivo": motivo_item,
            "usou_memoria": usou_memoria,
        }

        if score_item >= 0.78:
            candidatos_fortes.append(info_cand)
        elif score_item >= 0.42:
            if melhor_parcial is None or info_cand["score"] > melhor_parcial["score"]:
                melhor_parcial = info_cand

    candidatos_fortes.sort(key=lambda x: (x["score"], x["idx"] == idx_nome_atual), reverse=True)
    return candidatos_fortes, melhor_parcial, suspeita_familiar_atual, trecho_limpo


def verificar_pdf_contra_planilha(
    caminho_pdf: Path,
    itens_planilha: List[ItemPlanilha],
    idx_nome_atual: Optional[int],
    usar_ocr: bool = True,
    caminho_memoria_ia: Optional[Path] = None,
) -> dict:
    """
    Analisa o PDF atual contra a planilha usando OCR Local + Memória Evolutiva da IA (`memoria_ia.json`),
    e gera um **Diagnóstico Inteligente de Causa de Erro** (`diagnostico_causa`, `explicacao_causa`)
    junto com **Sugestões de Ação em 1 Clique** (`sugestoes_acao`).
    """
    mem_ia = carregar_memoria_ia(caminho_memoria_ia or _caminho_memoria_ia_padrao(caminho_pdf))
    num_paginas = contar_paginas_pdf(caminho_pdf) or 1
    texto_bruto, metodo = extrair_texto_pdf_local(caminho_pdf, usar_ocr_se_imagem=usar_ocr, pagina_1based=1)

    nome_esperado_txt = (
        itens_planilha[idx_nome_atual].nome_completo
        if (idx_nome_atual is not None and 0 <= idx_nome_atual < len(itens_planilha))
        else "aluno atual"
    )
    nome_fmt_esperado = (
        itens_planilha[idx_nome_atual].nome_formatado
        if (idx_nome_atual is not None and 0 <= idx_nome_atual < len(itens_planilha))
        else ""
    )

    # Caso 0: PDF sem texto legível na Página 1
    if not texto_bruto or len(texto_bruto.strip()) < 4:
        suspeita_virada = _detectar_suspeita_pagina_virada(caminho_pdf)
        sugestoes: List[dict] = []
        if suspeita_virada:
            diag_titulo = "Página possivelmente escaneada deitada ou invertida"
            diag_exp = (
                "O OCR não conseguiu ler as palavras porque as linhas de texto parecem estar na vertical (90°/270°). "
                "Clique em Auto-Orientar para corrigir a posição da folha e reanalisar."
            )
            sugestoes.append({
                "tipo": "auto_orientar",
                "rotulo": "Auto-Orientar Página e Reanalisar",
                "icone": "fa-compass",
                "estilo": "warning",
            })
        else:
            diag_titulo = "Documento manuscrito, imagem muito clara ou verso sem cabeçalho"
            diag_exp = (
                f"Não foi encontrado texto impresso legível no topo. Pode ser um documento manuscrito de '{nome_esperado_txt}', "
                "uma página em branco ou o verso solto do documento do aluno anterior."
            )
        sugestoes.append({
            "tipo": "anexar_anterior",
            "rotulo": "É verso/continuação? Anexar ao Aluno Anterior",
            "icone": "fa-file-import",
            "estilo": "info",
        })
        if idx_nome_atual is not None:
            sugestoes.append({
                "tipo": "confirmar_atual",
                "rotulo": f"Confirmar manualmente p/ {nome_fmt_esperado}",
                "icone": "fa-check",
                "estilo": "success",
            })
        return {
            "status": "sem_texto",
            "nivel_cor": "vermelho",
            "metodo": metodo,
            "confianca": 0.0,
            "idx_sugerido": None,
            "bate_com_atual": False,
            "candidatos_grupo": [],
            "motivo": "Sem texto legível ou manuscrito — requer leitura manual",
            "trecho": "",
            "diagnostico_causa": diag_titulo,
            "explicacao_causa": diag_exp,
            "sugestoes_acao": sugestoes,
        }

    candidatos_fortes, melhor_parcial, suspeita_familiar_atual, trecho_limpo = _avaliar_texto_contra_itens(
        texto_bruto, itens_planilha, idx_nome_atual, mem_ia
    )

    # Verificar se há um 2º aluno diferente escaneado a partir da Página 2 (quando o PDF tem 2+ páginas)
    outro_aluno_pag2: Optional[dict] = None
    if num_paginas >= 2:
        texto_p2, _ = extrair_texto_pdf_local(caminho_pdf, usar_ocr_se_imagem=False, pagina_1based=2)
        if texto_p2 and len(texto_p2.strip()) >= 15:
            cands_p2, _, _, _ = _avaliar_texto_contra_itens(texto_p2, itens_planilha, idx_nome_atual, mem_ia)
            if cands_p2:
                id_p1 = candidatos_fortes[0]["idx"] if candidatos_fortes else idx_nome_atual
                for cp2 in cands_p2:
                    if cp2["idx"] != id_p1 and cp2["score"] >= 82.0:
                        outro_aluno_pag2 = cp2
                        break

    # CASO ESPECIAL: 2 alunos diferentes em páginas separadas do mesmo PDF (esqueceu de separar no scanner!)
    if outro_aluno_pag2 is not None and len(candidatos_fortes) == 1:
        aluno_p1 = candidatos_fortes[0]
        return {
            "status": "dois_alunos_paginas_diferentes",
            "nivel_cor": "amarelo",
            "metodo": metodo,
            "confianca": aluno_p1["score"],
            "idx_sugerido": aluno_p1["idx"],
            "indice_original": aluno_p1["indice_original"],
            "nome_completo": aluno_p1["nome_completo"],
            "nome_formatado": aluno_p1["nome_formatado"],
            "bate_com_atual": aluno_p1["idx"] == idx_nome_atual,
            "candidatos_grupo": [aluno_p1, outro_aluno_pag2],
            "motivo": (
                f"Alerta: Pág. 1 é de {aluno_p1['nome_formatado']}, mas Pág. 2 parece ser de {outro_aluno_pag2['nome_formatado']}!"
            ),
            "trecho": trecho_limpo,
            "diagnostico_causa": "Dois alunos diferentes escaneados no mesmo arquivo PDF",
            "explicacao_causa": (
                f"A IA identificou '{aluno_p1['nome_completo']}' na Página 1, porém encontrou "
                f"'{outro_aluno_pag2['nome_completo']}' na Página 2. Provavelmente duas fichas foram puxadas juntas no scanner."
            ),
            "sugestoes_acao": [
                {
                    "tipo": "dividir_pag1",
                    "rotulo": f"Separar Página 1 ({aluno_p1['nome_formatado']}) do restante",
                    "icone": "fa-scissors",
                    "estilo": "warning",
                },
                {
                    "tipo": "confirmar_atual",
                    "rotulo": f"Manter junto e renomear p/ {aluno_p1['nome_formatado']}",
                    "icone": "fa-check",
                    "estilo": "primary",
                },
            ],
        }

    # CASO 1: TCC EM GRUPO (2 ou mais alunos da planilha na mesma página/capa)
    if len(candidatos_fortes) >= 2 and candidatos_fortes[1]["score"] >= 80.0:
        nomes_grupo = ", ".join(c["nome_completo"].split()[0] for c in candidatos_fortes[:4])
        bate_algum_atual = any(c["idx"] == idx_nome_atual for c in candidatos_fortes)
        escolhido = next((c for c in candidatos_fortes if c["idx"] == idx_nome_atual), candidatos_fortes[0])
        return {
            "status": "tcc_grupo",
            "nivel_cor": "amarelo",
            "metodo": metodo,
            "confianca": escolhido["score"],
            "idx_sugerido": escolhido["idx"],
            "indice_original": escolhido["indice_original"],
            "nome_completo": escolhido["nome_completo"],
            "nome_formatado": escolhido["nome_formatado"],
            "bate_com_atual": bate_algum_atual,
            "candidatos_grupo": candidatos_fortes[:5],
            "motivo": f"TCC/Grupo detectado: {len(candidatos_fortes)} alunos da planilha neste PDF ({nomes_grupo})",
            "trecho": trecho_limpo,
            "diagnostico_causa": f"Trabalho/TCC em Grupo ({len(candidatos_fortes)} integrantes da planilha na mesma capa)",
            "explicacao_causa": (
                f"Foram encontrados os alunos {', '.join(c['nome_completo'] for c in candidatos_fortes[:4])} na mesma página. "
                "Se cada aluno precisa ter seu arquivo individual na pasta, duplique o PDF antes de renomear."
            ),
            "sugestoes_acao": [
                {
                    "tipo": "duplicar_pdf",
                    "rotulo": "Duplicar PDF inteiro (Criar cópia p/ outro integrante)",
                    "icone": "fa-copy",
                    "estilo": "warning",
                },
                {
                    "tipo": "usar_sugerido_e_confirmar",
                    "idx_nome": escolhido["idx"],
                    "rotulo": f"Renomear este para {escolhido['nome_formatado']}",
                    "icone": "fa-check",
                    "estilo": "success",
                },
            ],
        }

    # CASO 2: 1 ÚNICO ALUNO ENCONTRADO COM FORTE CONFIANÇA
    if len(candidatos_fortes) == 1 or (len(candidatos_fortes) >= 1 and candidatos_fortes[0]["score"] >= 78.0):
        melhor = candidatos_fortes[0]
        bate_atual = melhor["idx"] == idx_nome_atual

        if bate_atual:
            return {
                "status": "confirmado_atual",
                "nivel_cor": "verde",
                "metodo": metodo,
                "confianca": melhor["score"],
                "idx_sugerido": melhor["idx"],
                "indice_original": melhor["indice_original"],
                "nome_completo": melhor["nome_completo"],
                "nome_formatado": melhor["nome_formatado"],
                "bate_com_atual": True,
                "candidatos_grupo": candidatos_fortes[:3],
                "motivo": melhor["motivo"],
                "trecho": trecho_limpo,
                "diagnostico_causa": "Alinhamento perfeito com a fila da planilha",
                "explicacao_causa": (
                    f"Os dados lidos no documento conferem com o aluno esperado na fila ({melhor['nome_completo']})."
                ),
                "sugestoes_acao": [
                    {
                        "tipo": "confirmar_atual",
                        "rotulo": f"Confirmar {melhor['nome_formatado']} [Enter]",
                        "icone": "fa-check",
                        "estilo": "success",
                    }
                ],
            }

        # Se NÃO bate com o aluno atual (`encontrado_outro`), diagnosticar o que aconteceu na fila!
        idx_enc = melhor["idx"]
        ja_concluido = bool(melhor.get("concluido"))

        if ja_concluido or (idx_nome_atual is not None and idx_enc < idx_nome_atual):
            diag_titulo = f"Continuação ou 2ª via de aluno já processado ({melhor['nome_formatado']})"
            diag_exp = (
                f"O documento pertence a '{melhor['nome_completo']}' (#{melhor['indice_original']}), que está atrás na fila "
                "ou já foi renomeado. Provavelmente esta folha é uma continuação/verso solto do documento dele."
            )
            sugestoes_outro = [
                {
                    "tipo": "anexar_anterior",
                    "rotulo": "Anexar ao último PDF renomeado [A]",
                    "icone": "fa-file-import",
                    "estilo": "warning",
                },
                {
                    "tipo": "usar_sugerido_e_confirmar",
                    "idx_nome": idx_enc,
                    "rotulo": f"Salvar como 2º arquivo ({melhor['nome_formatado']}_1.pdf)",
                    "icone": "fa-check-double",
                    "estilo": "primary",
                },
            ]
        else:
            salto = (idx_enc - idx_nome_atual) if idx_nome_atual is not None else 1
            nomes_pulados = []
            if idx_nome_atual is not None:
                for i_p in range(idx_nome_atual, min(idx_enc, idx_nome_atual + 3)):
                    if not itens_planilha[i_p].concluido:
                        nomes_pulados.append(itens_planilha[i_p].nome_completo.split()[0])
            txt_pulados = ", ".join(nomes_pulados) if nomes_pulados else nome_esperado_txt
            diag_titulo = f"Desalinhamento de fila: {salto} aluno(s) pulado(s) / sem documento"
            diag_exp = (
                f"A fila esperava '{nome_esperado_txt}', mas o PDF pertence a '{melhor['nome_completo']}' "
                f"({salto} posição(ões) à frente na planilha). Provavelmente {txt_pulados} faltou ou não teve PDF escaneado."
            )
            sugestoes_outro = [
                {
                    "tipo": "usar_sugerido_e_confirmar",
                    "idx_nome": idx_enc,
                    "rotulo": f"Pular faltante(s) e renomear p/ {melhor['nome_formatado']}",
                    "icone": "fa-bolt",
                    "estilo": "warning",
                },
                {
                    "tipo": "selecionar_nome",
                    "idx_nome": idx_enc,
                    "rotulo": f"Apenas posicionar fila em {melhor['nome_formatado']}",
                    "icone": "fa-crosshairs",
                    "estilo": "info",
                },
            ]

        return {
            "status": "encontrado_outro",
            "nivel_cor": "amarelo",
            "metodo": metodo,
            "confianca": melhor["score"],
            "idx_sugerido": idx_enc,
            "indice_original": melhor["indice_original"],
            "nome_completo": melhor["nome_completo"],
            "nome_formatado": melhor["nome_formatado"],
            "bate_com_atual": False,
            "candidatos_grupo": candidatos_fortes[:3],
            "motivo": melhor["motivo"],
            "trecho": trecho_limpo,
            "diagnostico_causa": diag_titulo,
            "explicacao_causa": diag_exp,
            "sugestoes_acao": sugestoes_outro,
        }

    # CASO 3: SUSPEITA DE DOCUMENTO DE FAMILIAR (Sobrenome bateu, mas 1º nome é diferente, ex: Pai/Mãe no dossiê)
    if suspeita_familiar_atual and idx_nome_atual is not None:
        return {
            "status": "duvida_familiar",
            "nivel_cor": "amarelo",
            "metodo": metodo,
            "confianca": 58.0,
            "idx_sugerido": idx_nome_atual,
            "bate_com_atual": False,
            "candidatos_grupo": [],
            "motivo": "Apenas sobrenome coincide (possível doc. de familiar/pai/mãe no dossiê) — confira",
            "trecho": trecho_limpo,
            "diagnostico_causa": "Possível documento de familiar (Pai/Mãe/Responsável) dentro do dossiê",
            "explicacao_causa": (
                f"O sobrenome de '{nome_esperado_txt}' aparece no documento, mas o primeiro nome é diferente. "
                "Pode ser o RG/CNH do pai ou mãe compondo o dossiê do aluno, ou um irmão homônimo."
            ),
            "sugestoes_acao": [
                {
                    "tipo": "ensinar_ia_atual",
                    "rotulo": f"Pertence ao dossiê de {nome_fmt_esperado} (Confirmar e Ensinar IA)",
                    "icone": "fa-graduation-cap",
                    "estilo": "success",
                },
                {
                    "tipo": "anexar_anterior",
                    "rotulo": "Anexar ao PDF do aluno anterior [A]",
                    "icone": "fa-file-import",
                    "estilo": "info",
                },
            ],
        }

    # CASO 4: DÚVIDA / NENHUM ALUNO COM ALTA CONFIANÇA (Analisar candidato parcial, rotação ou pessoa extra)
    sugestoes_duvida: List[dict] = []
    if melhor_parcial is not None:
        diag_titulo = (
            f"Leitura parcial ({melhor_parcial['score']}% parecido com {melhor_parcial['nome_completo']})"
        )
        diag_exp = (
            f"O OCR encontrou semelhança parcial com '{melhor_parcial['nome_completo']}', mas abaixo do limite de segurança "
            "(carimbo sobre o nome, fonte borrada ou abreviação). Se você confirmar e ensinar a IA, ela aprenderá essa variação!"
        )
        if melhor_parcial["idx"] == idx_nome_atual:
            sugestoes_duvida.append({
                "tipo": "ensinar_ia_atual",
                "rotulo": f"Confirmar {nome_fmt_esperado} e Ensinar IA",
                "icone": "fa-graduation-cap",
                "estilo": "success",
            })
        else:
            sugestoes_duvida.append({
                "tipo": "usar_sugerido_e_confirmar",
                "idx_nome": melhor_parcial["idx"],
                "rotulo": f"Renomear p/ {melhor_parcial['nome_formatado']} e Ensinar IA",
                "icone": "fa-graduation-cap",
                "estilo": "warning",
            })
    else:
        diag_titulo = "Nome do documento não consta na planilha ou é página intermediária sem cabeçalho"
        diag_exp = (
            "O texto extraído não coincide com os alunos da planilha. Isso costuma ocorrer em: "
            "(1) Aluno extra fora da planilha, (2) Página 2/verso solto do aluno anterior, ou (3) Nome de coordenador/diretor em destaque."
        )
        if idx_nome_atual is not None:
            sugestoes_duvida.append({
                "tipo": "ensinar_ia_atual",
                "rotulo": f"Forçar {nome_fmt_esperado} e Ensinar Padrão à IA",
                "icone": "fa-graduation-cap",
                "estilo": "success",
            })

    if _detectar_suspeita_pagina_virada(caminho_pdf):
        sugestoes_duvida.insert(0, {
            "tipo": "auto_orientar",
            "rotulo": "Página parece virada: Auto-Orientar e Reanalisar",
            "icone": "fa-compass",
            "estilo": "warning",
        })

    sugestoes_duvida.append({
        "tipo": "anexar_anterior",
        "rotulo": "Anexar ao PDF do aluno anterior [A]",
        "icone": "fa-file-import",
        "estilo": "info",
    })
    sugestoes_duvida.append({
        "tipo": "abrir_modal_manual",
        "rotulo": "Cadastrar Pessoa Extra (Fora da Planilha) [E]",
        "icone": "fa-user-plus",
        "estilo": "primary",
    })

    return {
        "status": "duvida",
        "nivel_cor": "vermelho",
        "metodo": metodo,
        "confianca": melhor_parcial["score"] if melhor_parcial else 0.0,
        "idx_sugerido": melhor_parcial["idx"] if melhor_parcial else None,
        "bate_com_atual": False,
        "candidatos_grupo": [melhor_parcial] if melhor_parcial else [],
        "motivo": (
            f"Semelhança parcial ({melhor_parcial['score']}%) com {melhor_parcial['nome_formatado']} — confira"
            if melhor_parcial
            else "Nenhum aluno da planilha identificado neste PDF — veja o diagnóstico da IA"
        ),
        "trecho": trecho_limpo,
        "diagnostico_causa": diag_titulo,
        "explicacao_causa": diag_exp,
        "sugestoes_acao": sugestoes_duvida,
    }


# ==============================================================================
# 6. EXPORTAÇÃO DE PLANILHA / RELATÓRIO DE CONFERÊNCIA FINAL (.XLSX E .CSV)
# ==============================================================================

def exportar_planilha_conferencia(
    itens_planilha: List[ItemPlanilha],
    pdfs_dir: Path,
    caminho_xlsx: Path,
    caminho_csv: Path,
) -> Tuple[Path, str]:
    """
    Exporta a planilha completa de conferência com:
    - Todos os alunos da planilha (Status: Concluído vs Faltante / Sem PDF)
    - Alunos Extras (arquivos PDF renomeados na pasta que não estavam na planilha)
    Retorna (caminho_arquivo_gerado, mensagem).
    """
    stems_planilha = {it.nome_formatado.lower() for it in itens_planilha}
    extras: List[str] = []
    if pdfs_dir.exists():
        for pdf in sorted(pdfs_dir.iterdir(), key=_chave_ordenacao_natural):
            if pdf.is_file() and pdf.suffix.lower() == ".pdf" and not pdf.stem.startswith("."):
                stem_low = pdf.stem.lower()
                base_low = re.sub(r"_\d+$", "", stem_low)
                if stem_low not in stems_planilha and base_low not in stems_planilha:
                    if not re.match(r"^\d+$", pdf.stem):
                        extras.append(pdf.name)

    cabecalho = [
        "Ordem Planilha",
        "Nome Completo",
        "Matrícula",
        "CPF",
        "Nome Padrão do Arquivo",
        "Status",
        "Arquivo Salvo na Pasta",
    ]
    linhas_dados = []
    for it in itens_planilha:
        status = "CONCLUÍDO" if it.concluido else "FALTANTE (SEM PDF)"
        arq = it.arquivo_gerado or (f"{it.nome_formatado}.pdf" if it.concluido else "")
        linhas_dados.append([
            str(it.indice_original),
            it.nome_completo,
            it.matricula,
            it.cpf,
            f"{it.nome_formatado}.pdf",
            status,
            arq,
        ])

    for idx_ex, nome_arq in enumerate(extras, start=1):
        linhas_dados.append([
            f"EXTRA-{idx_ex}",
            Path(nome_arq).stem.replace("_", " "),
            "",
            "",
            nome_arq,
            "EXTRA (FORA DA PLANILHA)",
            nome_arq,
        ])

    # Sempre salva CSV (compatível com qualquer Excel)
    try:
        with open(caminho_csv, "w", encoding="utf-8-sig", newline="") as f_csv:
            writer = csv.writer(f_csv, delimiter=";")
            writer.writerow(cabecalho)
            writer.writerows(linhas_dados)
    except Exception:
        pass

    # Tentar salvar .xlsx formatado com openpyxl
    try:
        import openpyxl  # type: ignore
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Conferencia_PDFs"
        ws.append(cabecalho)
        for linha in linhas_dados:
            ws.append(linha)
        wb.save(str(caminho_xlsx))
        wb.close()
        return caminho_xlsx, f"Planilha de conferência salva em '{caminho_xlsx.name}'."
    except Exception:
        return caminho_csv, f"Planilha de conferência salva em '{caminho_csv.name}'."


# ==============================================================================
# 7. INTERAÇÃO COM O SISTEMA E UTILITÁRIOS DE LOG / DESFAZER
# ==============================================================================

def limpar_tela():
    os.system("cls" if os.name == "nt" else "clear")


def trazer_terminal_para_frente():
    if os.name != "nt":
        return False
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        console_window = kernel32.GetConsoleWindow()
        if console_window:
            user32.ShowWindow(console_window, 9)
            user32.SetForegroundWindow(console_window)
            user32.SetFocus(console_window)
            return True
    except Exception:
        pass
    return False


def abrir_pdf(caminho_pdf: Path) -> bool:
    try:
        if os.name == "nt":
            os.startfile(str(caminho_pdf))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(
                ["open", str(caminho_pdf)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        else:
            subprocess.Popen(
                ["xdg-open", str(caminho_pdf)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        return True
    except Exception as e:
        print(f"{Cor.VERMELHO}Erro ao abrir PDF '{caminho_pdf.name}': {e}{Cor.RESET}")
        return False


def ler_tecla_unica(opcoes_validas: Set[str]) -> str:
    if os.name == "nt":
        import msvcrt
        while True:
            ch = msvcrt.getwch()
            if ch == "\x03":
                raise KeyboardInterrupt
            if ch in ("\x00", "\xe0"):
                msvcrt.getwch()
                continue
            tecla = ch.lower()
            if tecla in opcoes_validas:
                print(tecla.upper())
                return tecla
    else:
        import termios
        import tty
        fd = sys.stdin.fileno()
        if not os.isatty(fd):
            linha = sys.stdin.readline()
            if not linha:
                return "q"
            tecla = linha.strip().lower()[:1]
            print(tecla.upper())
            return tecla if tecla in opcoes_validas else "q"

        config_antiga = termios.tcgetattr(fd)
        try:
            tty.setraw(fd)
            while True:
                ch = sys.stdin.read(1)
                if ch == "\x03":
                    raise KeyboardInterrupt
                tecla = ch.lower()
                if tecla in opcoes_validas:
                    sys.stdout.write(tecla.upper() + "\r\n")
                    sys.stdout.flush()
                    return tecla
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, config_antiga)


def registrar_log(log_path: Path, mensagem: str):
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{agora}] {mensagem}\n")


def gerar_barra_progresso(atual: int, total: int, largura: int = 36) -> str:
    if total <= 0:
        return ""
    proporcao = min(max(atual / float(total), 0.0), 1.0)
    preenchido = int(largura * proporcao)
    barra = "█" * preenchido + "░" * (largura - preenchido)
    percentual = 100.0 * proporcao
    return f"[{Cor.CIANO}{barra}{Cor.RESET}] {percentual:5.1f}% ({atual}/{total})"


def proximo_indice_nome_pendente(itens: List[ItemPlanilha], a_partir_de: int = 0) -> Optional[int]:
    for idx in range(max(0, a_partir_de), len(itens)):
        if not itens[idx].concluido:
            return idx
    for idx in range(0, min(max(0, a_partir_de), len(itens))):
        if not itens[idx].concluido:
            return idx
    return None


def obter_destino_unico(pasta: Path, nome_base: str) -> Path:
    nome_limpo = limpar_nome_arquivo(nome_base)
    if nome_limpo.lower().endswith(".pdf"):
        nome_limpo = nome_limpo[:-4]
    destino = pasta / f"{nome_limpo}.pdf"
    contador = 1
    while destino.exists():
        destino = pasta / f"{nome_limpo}_{contador}.pdf"
        contador += 1
    return destino


def executar_renomeacao_arquivo(
    arquivo_pdf: Path,
    nome_final: str,
    pdfs_dir: Path,
    log_path: Path,
) -> Tuple[Optional[Path], Optional[str]]:
    destino = obter_destino_unico(pdfs_dir, nome_final)
    try:
        arquivo_pdf.rename(destino)
        registrar_log(log_path, f"SUCESSO: '{arquivo_pdf.name}' -> '{destino.name}'")
        return destino, None
    except Exception as e:
        registrar_log(log_path, f"ERRO: '{arquivo_pdf.name}' -> '{nome_final}': {e}")
        return None, str(e)


def desfazer_ultima_acao(
    historico: List[AcaoHistorico],
    fila_pdfs: List[Path],
    itens_planilha: List[ItemPlanilha],
    log_path: Path,
) -> Tuple[Optional[int], Optional[int], str]:
    """Desfaz a última ação registrada no histórico da sessão (renomear, dividir, juntar, anexar, editar_pdf, pular)."""
    if not historico:
        return None, None, f"{Cor.AMARELO}⚠️ Nada para desfazer nesta sessão.{Cor.RESET}"

    acao = historico.pop()

    if acao.tipo == "renomear":
        if acao.pdf_destino and acao.pdf_origem and acao.pdf_destino.exists():
            try:
                acao.pdf_destino.rename(acao.pdf_origem)
                registrar_log(
                    log_path,
                    f"REVERTIDO: '{acao.pdf_destino.name}' -> '{acao.pdf_origem.name}'",
                )
            except Exception as e:
                return None, None, f"{Cor.VERMELHO}❌ Erro ao reverter arquivo: {e}{Cor.RESET}"
        if acao.idx_nome_concluido is not None and 0 <= acao.idx_nome_concluido < len(itens_planilha):
            itens_planilha[acao.idx_nome_concluido].concluido = False
            itens_planilha[acao.idx_nome_concluido].arquivo_gerado = ""
        return (
            acao.idx_pdf_anterior,
            acao.idx_nome_anterior,
            f"{Cor.AMARELO}↩️ Desfeito: '{acao.pdf_destino.name if acao.pdf_destino else ''}' voltou a ser '{acao.pdf_origem.name if acao.pdf_origem else ''}'{Cor.RESET}",
        )

    elif acao.tipo == "dividir":
        for parte in acao.arquivos_criados:
            if parte in fila_pdfs:
                fila_pdfs.remove(parte)
            if parte.exists():
                try:
                    parte.unlink()
                except Exception:
                    pass
        if acao.backup_origem and acao.backup_origem.exists() and acao.pdf_origem:
            try:
                shutil.move(str(acao.backup_origem), str(acao.pdf_origem))
                fila_pdfs.insert(acao.idx_pdf_anterior, acao.pdf_origem)
            except Exception as e:
                return None, None, f"{Cor.VERMELHO}Erro ao restaurar PDF dividido: {e}{Cor.RESET}"
        return (
            acao.idx_pdf_anterior,
            acao.idx_nome_anterior,
            f"{Cor.AMARELO}↩️ Divisão desfeita: '{acao.pdf_origem.name if acao.pdf_origem else ''}' restaurado.{Cor.RESET}",
        )

    elif acao.tipo == "juntar":
        # Remover o PDF unido se não for um dos originais que será sobrescrito, e restaurar todos os originais
        for criado in acao.arquivos_criados:
            if criado in fila_pdfs:
                fila_pdfs.remove(criado)
            if criado.exists():
                try:
                    criado.unlink()
                except Exception:
                    pass
        restaurados: List[Path] = []
        for orig, bkp in zip(acao.arquivos_originais, acao.backups_originais):
            if bkp.exists():
                try:
                    shutil.move(str(bkp), str(orig))
                    restaurados.append(orig)
                except Exception as e:
                    return None, None, f"{Cor.VERMELHO}Erro ao restaurar PDF juntado: {e}{Cor.RESET}"
        fila_pdfs[acao.idx_pdf_anterior : acao.idx_pdf_anterior] = restaurados
        return (
            acao.idx_pdf_anterior,
            acao.idx_nome_anterior,
            f"{Cor.AMARELO}↩️ Junção desfeita: {len(restaurados)} PDFs separados novamente.{Cor.RESET}",
        )

    elif acao.tipo == "anexar_anterior":
        # Restaurar o PDF destino (anterior) e o PDF origem (atual) a partir de seus backups
        for orig, bkp in zip(acao.arquivos_originais, acao.backups_originais):
            if bkp.exists():
                try:
                    shutil.move(str(bkp), str(orig))
                except Exception as e:
                    return None, None, f"{Cor.VERMELHO}Erro ao desfazer anexação: {e}{Cor.RESET}"
        return (
            acao.idx_pdf_anterior,
            acao.idx_nome_anterior,
            f"{Cor.AMARELO}↩️ Anexação ao PDF anterior desfeita ({acao.descricao}).{Cor.RESET}",
        )

    elif acao.tipo == "editar_pdf":
        if acao.backup_origem and acao.backup_origem.exists() and acao.pdf_origem:
            try:
                shutil.move(str(acao.backup_origem), str(acao.pdf_origem))
            except Exception as e:
                return None, None, f"{Cor.VERMELHO}Erro ao desfazer edição de página: {e}{Cor.RESET}"
        return (
            acao.idx_pdf_anterior,
            acao.idx_nome_anterior,
            f"{Cor.AMARELO}↩️ Edição de página desfeita ({acao.descricao}).{Cor.RESET}",
        )

    else:
        return (
            acao.idx_pdf_anterior,
            acao.idx_nome_anterior,
            f"{Cor.AMARELO}↩️ Passo anterior restaurado ({acao.descricao}).{Cor.RESET}",
        )


# ==============================================================================
# 8. MENUS INTERATIVOS DO CLI (v6.0)
# ==============================================================================

def menu_editar_nome(item_sugerido: Optional[ItemPlanilha]) -> Optional[Tuple[str, bool]]:
    print("\n" + "─" * 70)
    print(f"{Cor.BOLD}{Cor.AMARELO}✏️  MODO EDIÇÃO / PESSOA FORA DA PLANILHA (3 PRIORIDADES){Cor.RESET}")
    print(
        f"{Cor.DIM}• Com Matrícula -> Primeiro_Segundo_Matricula\n"
        f"• Sem Matrícula + Com CPF -> Primeiro_Segundo_CPF\n"
        f"• Sem Matrícula e Sem CPF -> nome_completo_separado_por_underline{Cor.RESET}"
    )
    entrada = input("\nDigite o Nome Completo (ou Enter para cancelar): ").strip()
    if not entrada:
        return None

    tem_digitos = any(ch.isdigit() for ch in entrada)
    matricula_extra = ""
    cpf_extra = ""
    if not tem_digitos and "_" not in entrada:
        matricula_extra = input("Digite a Matrícula (ou Enter se não constar): ").strip()
        if _eh_valor_nao_consta(matricula_extra):
            cpf_extra = input("Digite o CPF (ou Enter se não constar p/ usar nome_completo): ").strip()

    _, _, nome_formatado = formatar_nome_e_matricula(entrada, matricula_extra, cpf_extra)
    if not nome_formatado:
        return None

    print(f"\n➡️  Arquivo ficará: {Cor.VERDE}{Cor.BOLD}{nome_formatado}.pdf{Cor.RESET}")

    if item_sugerido is not None:
        print(
            f"\nO que fazer com o nome que estava na fila ({Cor.CIANO}{item_sugerido.nome_formatado}{Cor.RESET})?"
        )
        print(f"  {Cor.BOLD}[1]{Cor.RESET} Pessoa EXTRA que não estava na planilha (MANTER '{item_sugerido.nome_formatado}' para o próximo PDF) {Cor.VERDE}[Padrão - Enter]{Cor.RESET}")
        print(f"  {Cor.BOLD}[2]{Cor.RESET} SUBSTITUIR '{item_sugerido.nome_formatado}' (avançar a fila da planilha também)")
        print(f"  {Cor.BOLD}[C]{Cor.RESET} Cancelar")
        escolha = input("Escolha [1/2/C]: ").strip().lower()
        if escolha == "c":
            return None
        avancar_planilha = escolha == "2"
    else:
        avancar_planilha = False

    return nome_formatado, avancar_planilha


def menu_buscar_na_planilha(itens: List[ItemPlanilha]) -> Optional[int]:
    print("\n" + "─" * 70)
    print(f"{Cor.BOLD}{Cor.CIANO}🔎 BUSCAR PESSOA NA PLANILHA (NOME, MATRÍCULA OU CPF){Cor.RESET}")
    termo = input("Digite parte do nome, matrícula ou CPF (ou Enter para cancelar): ").strip()
    if not termo:
        return None

    termo_norm = normalizar_busca(termo)
    encontrados: List[int] = []
    for idx, item in enumerate(itens):
        alvo = normalizar_busca(f"{item.nome_completo} {item.matricula} {item.cpf} {item.nome_formatado}")
        if termo_norm in alvo:
            encontrados.append(idx)

    if not encontrados:
        print(f"{Cor.VERMELHO}Nenhum nome encontrado com '{termo}'.{Cor.RESET}")
        input("Pressione Enter para voltar...")
        return None

    print("\nResultados encontrados:")
    for pos, idx in enumerate(encontrados[:15], start=1):
        it = itens[idx]
        status = f"{Cor.DIM}(já renomeado){Cor.RESET}" if it.concluido else f"{Cor.VERDE}(pendente){Cor.RESET}"
        print(
            f"  {Cor.BOLD}[{pos}]{Cor.RESET} {it.nome_formatado} "
            f"{Cor.DIM}({it.nome_completo}){Cor.RESET} {status}"
        )

    escolha = input("\nDigite o número desejado (ou Enter para cancelar): ").strip()
    if not escolha.isdigit():
        return None
    num = int(escolha)
    if 1 <= num <= min(len(encontrados), 15):
        return encontrados[num - 1]
    return None


def menu_dividir_pdf(
    arquivo_pdf: Path,
    backup_dir: Path,
) -> Optional[Tuple[List[Path], Path]]:
    print("\n" + "─" * 70)
    print(f"{Cor.BOLD}{Cor.MAGENTA}✂️  DIVIDIR PDF COM 2+ PESSOAS ({arquivo_pdf.name}){Cor.RESET}")
    num_paginas = contar_paginas_pdf(arquivo_pdf)

    if num_paginas is not None:
        print(f"📄 Páginas detectadas neste PDF: {Cor.BOLD}{num_paginas}{Cor.RESET}")
    else:
        print(f"{Cor.AMARELO}Não foi possível detectar o número de páginas antecipadamente.{Cor.RESET}")

    intervalos: List[Tuple[int, int]] = []
    duplicar_inteiro = False

    if num_paginas == 1:
        print("\n⚠️  Este arquivo possui apenas 1 página.")
        print(f"  {Cor.BOLD}[D]{Cor.RESET} Duplicar o PDF inteiro (criar cópia para poder renomear para as 2 pessoas)")
        print(f"  {Cor.BOLD}[C]{Cor.RESET} Cancelar")
        esc = input("Escolha [D/C]: ").strip().lower()
        if esc != "d":
            return None
        duplicar_inteiro = True

    elif num_paginas == 2:
        print(f"\n  {Cor.BOLD}[Enter ou 1]{Cor.RESET} Separar Página 1 (1ª pessoa) e Página 2 (2ª pessoa) {Cor.VERDE}[Padrão]{Cor.RESET}")
        print(f"  {Cor.BOLD}[D]{Cor.RESET}          Duplicar o PDF inteiro para 2 pessoas sem cortar páginas")
        print(f"  {Cor.BOLD}[C]{Cor.RESET}          Cancelar")
        esc = input("Escolha [Enter/1/D/C]: ").strip().lower()
        if esc == "c":
            return None
        elif esc == "d":
            duplicar_inteiro = True
        else:
            intervalos = [(1, 1), (2, 2)]

    else:
        total_txt = str(num_paginas) if num_paginas else "N"
        print(f"\nComo deseja separar as {total_txt} páginas?")
        print(f"  {Cor.BOLD}[Enter ou 1]{Cor.RESET} Página 1 para a 1ª pessoa | Páginas 2..{total_txt} para a próxima")
        if num_paginas and num_paginas > 2:
            print(f"  {Cor.BOLD}[2..{num_paginas-1}]{Cor.RESET}       Digite até qual página vai a 1ª pessoa (ex: '2' -> Págs 1-2 e Págs 3-{num_paginas})")
            print(f"  {Cor.BOLD}[T]{Cor.RESET}          Separar TODAS as {num_paginas} páginas em PDFs individuais (1 página por PDF)")
        print(f"  {Cor.BOLD}[D]{Cor.RESET}          Duplicar o PDF inteiro para 2 pessoas sem cortar páginas")
        print(f"  {Cor.BOLD}[C]{Cor.RESET}          Cancelar")
        esc = input("Sua escolha: ").strip().lower()
        if esc == "c":
            return None
        elif esc == "d":
            duplicar_inteiro = True
        elif esc == "t" and num_paginas:
            intervalos = [(p, p) for p in range(1, num_paginas + 1)]
        elif esc == "" or esc == "1":
            fim = num_paginas if num_paginas else 2
            intervalos = [(1, 1), (2, fim)]
        elif esc.isdigit() and num_paginas and 1 <= int(esc) < num_paginas:
            corte = int(esc)
            intervalos = [(1, corte), (corte + 1, num_paginas)]
        else:
            print(f"{Cor.VERMELHO}Opção inválida.{Cor.RESET}")
            input("Enter para voltar...")
            return None

    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_path = obter_destino_unico(backup_dir, arquivo_pdf.stem)

    if duplicar_inteiro:
        parte1 = obter_destino_unico(arquivo_pdf.parent, f"{arquivo_pdf.stem}_parte1")
        parte2 = obter_destino_unico(arquivo_pdf.parent, f"{arquivo_pdf.stem}_parte2")
        try:
            shutil.copy2(arquivo_pdf, backup_path)
            shutil.copy2(arquivo_pdf, parte1)
            shutil.move(str(arquivo_pdf), str(parte2))
            return [parte1, parte2], backup_path
        except Exception as e:
            print(f"{Cor.VERMELHO}Erro ao duplicar PDF: {e}{Cor.RESET}")
            input("Enter para continuar...")
            return None

    novos_pdfs, status = dividir_pdf_em_intervalos(arquivo_pdf, intervalos)
    if not novos_pdfs:
        print(f"{Cor.VERMELHO}❌ {status}{Cor.RESET}")
        input("Enter para continuar...")
        return None

    try:
        shutil.move(str(arquivo_pdf), str(backup_path))
    except Exception as e:
        print(f"{Cor.AMARELO}Aviso: não foi possível mover original para backup: {e}{Cor.RESET}")

    return novos_pdfs, backup_path


def exibir_painel_principal(
    pdf_atual: Path,
    pos_pdf: int,
    total_pdfs_sessao: int,
    item_atual: Optional[ItemPlanilha],
    itens_planilha: List[ItemPlanilha],
    idx_nome: Optional[int],
    total_renomeados_sessao: int,
    ultimo_evento: str,
):
    limpar_tela()
    total_planilha = len(itens_planilha)
    concluidos_planilha = sum(1 for it in itens_planilha if it.concluido)

    print(f"{Cor.BOLD}{Cor.AZUL}╔══════════════════════════════════════════════════════════════════════╗{Cor.RESET}")
    print(f"{Cor.BOLD}{Cor.AZUL}║         📄 RENOMEADOR SEMIAUTOMÁTICO DE PDFs - CLI v6.0              ║{Cor.RESET}")
    print(f"{Cor.BOLD}{Cor.AZUL}╚══════════════════════════════════════════════════════════════════════╝{Cor.RESET}")

    print(f" Progresso Planilha: {gerar_barra_progresso(concluidos_planilha, total_planilha)}")
    print(f" PDFs nesta sessão : {pos_pdf}/{total_pdfs_sessao}  |  Renomeados agora: {Cor.VERDE}{total_renomeados_sessao}{Cor.RESET}")

    if ultimo_evento:
        print(f" Última ação       : {ultimo_evento}")

    print("─" * 72)
    print(f" 📂 {Cor.BOLD}ARQUIVO PDF ATUAL :{Cor.RESET} {Cor.AMARELO}{Cor.BOLD}{pdf_atual.name}{Cor.RESET}")

    if item_atual is not None:
        print(f" 👤 {Cor.BOLD}NOME NA PLANILHA  :{Cor.RESET} {item_atual.nome_completo}")
        if item_atual.matricula:
            print(f" 🔢 {Cor.BOLD}MATRÍCULA (Prio 1):{Cor.RESET} {item_atual.matricula}")
        elif item_atual.cpf:
            print(f" 🪪 {Cor.BOLD}CPF (Prioridade 2):{Cor.RESET} {item_atual.cpf}")
        else:
            print(f" 📝 {Cor.BOLD}MODO (Prioridade 3):{Cor.RESET} Nome completo com '_'")
        print(
            f" 🎯 {Cor.BOLD}COMO VAI FICAR    :{Cor.RESET} "
            f"{Cor.VERDE}{Cor.BOLD}{item_atual.nome_formatado}.pdf{Cor.RESET} "
            f"{Cor.DIM}(Item #{item_atual.indice_original}/{total_planilha}){Cor.RESET}"
        )

        proximos: List[str] = []
        if idx_nome is not None:
            for p_idx in range(idx_nome + 1, len(itens_planilha)):
                if not itens_planilha[p_idx].concluido:
                    proximos.append(itens_planilha[p_idx].nome_formatado)
                if len(proximos) >= 2:
                    break
        if proximos:
            print(f" ⏭️  {Cor.DIM}Próximos da fila  : {'  ->  '.join(proximos)}{Cor.RESET}")
    else:
        print(
            f" ⚠️  {Cor.AMARELO}Todos os nomes da planilha já foram concluídos! "
            f"Use [E] para digitar um nome manual ou [B] para buscar.{Cor.RESET}"
        )

    print("─" * 72)
    print(
        f" {Cor.BOLD}{Cor.VERDE}[S]{Cor.RESET} Sim (Renomear)      "
        f"{Cor.BOLD}{Cor.AMARELO}[E]{Cor.RESET} Editar / Pessoa Extra   "
        f"{Cor.BOLD}{Cor.CIANO}[B]{Cor.RESET} Buscar na Planilha"
    )
    print(
        f" {Cor.BOLD}{Cor.MAGENTA}[D]{Cor.RESET} Dividir PDF         "
        f"{Cor.BOLD}{Cor.MAGENTA}[J]{Cor.RESET} Juntar c/ Próximo PDF   "
        f"{Cor.BOLD}{Cor.CIANO}[G]{Cor.RESET} Girar / Limpar Bordas"
    )
    print(
        f" {Cor.BOLD}[N]{Cor.RESET} Pular Nome da Fila  "
        f"{Cor.BOLD}[P]{Cor.RESET} Pular Arquivo PDF       "
        f"{Cor.BOLD}[Z]{Cor.RESET} Desfazer Último"
    )
    print(
        f" {Cor.BOLD}[R]{Cor.RESET} Reabrir PDF         "
        f"{Cor.BOLD}{Cor.VERMELHO}[Q]{Cor.RESET} Sair e Gerar Planilha de Conferência"
    )
    print("─" * 72)
    print(" Escolha uma opção (sem precisar de Enter): ", end="", flush=True)


def main(args):
    ativar_cores_windows()
    script_dir = Path(__file__).resolve().parent
    planilha_path = Path(args.planilha)
    if not planilha_path.is_absolute() and not planilha_path.exists():
        planilha_path = script_dir / args.planilha

    pdfs_dir = Path(args.pasta_pdfs)
    if not pdfs_dir.is_absolute() and not pdfs_dir.exists():
        pdfs_dir = script_dir / args.pasta_pdfs

    log_path = script_dir / ARQUIVO_LOG
    backup_divisao_dir = pdfs_dir / PASTA_BACKUP_DIVISAO

    limpar_tela()
    print("=" * 72)
    print(f"{Cor.BOLD}     📄 RENOMEADOR SEMIAUTOMÁTICO DE PDFs - CLI v6.0{Cor.RESET}")
    print("=" * 72)

    itens_planilha, msg = carregar_planilha(planilha_path)
    if not itens_planilha:
        print(f"\n{Cor.VERMELHO}[ERRO] {msg}{Cor.RESET}")
        if not args.sem_pausa:
            input("\nPressione Enter para sair...")
        return

    if not pdfs_dir.exists():
        print(f"\n{Cor.VERMELHO}[ERRO] Pasta de PDFs '{pdfs_dir}' não encontrada.{Cor.RESET}")
        if not args.sem_pausa:
            input("\nPressione Enter para sair...")
        return

    fila_pdfs, pdfs_ja_renomeados, qtd_nomes_concluidos = sincronizar_pdfs_e_planilha(
        pdfs_dir, itens_planilha, log_path
    )

    print(f"\n📊 {Cor.BOLD}Resumo carregado:{Cor.RESET}")
    print(f"   • Total de nomes na planilha : {Cor.BOLD}{len(itens_planilha)}{Cor.RESET}")
    print(f"   • Nomes já concluídos na pasta: {Cor.VERDE}{qtd_nomes_concluidos}{Cor.RESET}")
    print(f"   • Nomes pendentes na fila     : {Cor.AMARELO}{len(itens_planilha) - qtd_nomes_concluidos}{Cor.RESET}")
    print(f"   • PDFs pendentes para renomear: {Cor.CIANO}{len(fila_pdfs)}{Cor.RESET}")

    if not fila_pdfs:
        print(f"\n{Cor.VERDE}🎉 Nenhum PDF pendente encontrado em '{pdfs_dir}'! Tudo já está renomeado.{Cor.RESET}")
        if not args.sem_pausa:
            input("\nPressione Enter para sair...")
        return

    if not args.auto_iniciar:
        resp_ini = input("\nPressione [Enter] para iniciar (ou 'q' para sair): ").strip().lower()
        if resp_ini == "q":
            return

    historico: List[AcaoHistorico] = []
    idx_pdf = 0
    idx_nome = proximo_indice_nome_pendente(itens_planilha, 0)

    total_renomeados = 0
    total_editados = 0
    total_pulados = 0
    total_divididos = 0
    arquivos_pulados: List[dict] = []
    erros: List[dict] = []
    ultimo_evento = ""
    ultimo_pdf_aberto: Optional[Path] = None

    try:
        while idx_pdf < len(fila_pdfs):
            arquivo_pdf = fila_pdfs[idx_pdf]
            item_atual = (
                itens_planilha[idx_nome]
                if (idx_nome is not None and 0 <= idx_nome < len(itens_planilha) and not itens_planilha[idx_nome].concluido)
                else None
            )
            if item_atual is None:
                idx_nome = proximo_indice_nome_pendente(itens_planilha, 0)
                item_atual = itens_planilha[idx_nome] if idx_nome is not None else None

            exibir_painel_principal(
                pdf_atual=arquivo_pdf,
                pos_pdf=idx_pdf + 1,
                total_pdfs_sessao=len(fila_pdfs),
                item_atual=item_atual,
                itens_planilha=itens_planilha,
                idx_nome=idx_nome,
                total_renomeados_sessao=total_renomeados,
                ultimo_evento=ultimo_evento,
            )

            if ultimo_pdf_aberto != arquivo_pdf and not args.nao_abrir_pdf:
                if abrir_pdf(arquivo_pdf):
                    ultimo_pdf_aberto = arquivo_pdf
                    trazer_terminal_para_frente()
                else:
                    erros.append({"arquivo": arquivo_pdf.name, "erro": "Falha ao abrir visualizador"})

            tecla = ler_tecla_unica({"s", "e", "b", "d", "j", "g", "n", "p", "z", "r", "q"})

            if tecla == "r":
                abrir_pdf(arquivo_pdf)
                trazer_terminal_para_frente()
                ultimo_evento = f"🔄 PDF '{arquivo_pdf.name}' reaberto."
                continue

            elif tecla == "z":
                novo_idx_pdf, novo_idx_nome, msg_undo = desfazer_ultima_acao(
                    historico, fila_pdfs, itens_planilha, log_path
                )
                ultimo_evento = msg_undo
                if novo_idx_pdf is not None:
                    idx_pdf = novo_idx_pdf
                    idx_nome = novo_idx_nome
                    ultimo_pdf_aberto = None
                continue

            elif tecla == "j":
                if idx_pdf + 1 >= len(fila_pdfs):
                    ultimo_evento = f"{Cor.AMARELO}⚠️ Não há próximo PDF na fila para juntar.{Cor.RESET}"
                    continue
                prox_pdf = fila_pdfs[idx_pdf + 1]
                backup_divisao_dir.mkdir(parents=True, exist_ok=True)
                bkp1 = obter_destino_unico(backup_divisao_dir, arquivo_pdf.stem)
                bkp2 = obter_destino_unico(backup_divisao_dir, prox_pdf.stem)
                shutil.copy2(arquivo_pdf, bkp1)
                shutil.copy2(prox_pdf, bkp2)
                ok_j, msg_j = juntar_arquivos_pdf([arquivo_pdf, prox_pdf], arquivo_pdf)
                if ok_j:
                    prox_pdf.unlink(missing_ok=True)
                    fila_pdfs.pop(idx_pdf + 1)
                    historico.append(
                        AcaoHistorico(
                            tipo="juntar",
                            idx_pdf_anterior=idx_pdf,
                            idx_nome_anterior=idx_nome if idx_nome is not None else 0,
                            arquivos_criados=[arquivo_pdf],
                            arquivos_originais=[arquivo_pdf, prox_pdf],
                            backups_originais=[bkp1, bkp2],
                            descricao=f"Juntou {arquivo_pdf.name} + {prox_pdf.name}",
                        )
                    )
                    ultimo_pdf_aberto = None
                    ultimo_evento = f"{Cor.VERDE}📎 Juntou '{arquivo_pdf.name}' + '{prox_pdf.name}'!{Cor.RESET}"
                else:
                    ultimo_evento = f"{Cor.VERMELHO}❌ {msg_j}{Cor.RESET}"
                continue

            elif tecla == "g":
                print("\nOpções de Página: [1] Girar 90° Dir | [2] Girar 90° Esq | [3] Girar 180° | [4] Auto-Limpar Bordas Pretas | [5] Auto-Orientar")
                esc_g = input("Escolha [1-5 ou Enter p/ cancelar]: ").strip()
                if esc_g in ("1", "2", "3", "4", "5"):
                    backup_divisao_dir.mkdir(parents=True, exist_ok=True)
                    bkp = obter_destino_unico(backup_divisao_dir, arquivo_pdf.stem)
                    shutil.copy2(arquivo_pdf, bkp)
                    if esc_g == "1":
                        ok_e, msg_e = rotacionar_paginas_pdf(arquivo_pdf, 90)
                    elif esc_g == "2":
                        ok_e, msg_e = rotacionar_paginas_pdf(arquivo_pdf, -90)
                    elif esc_g == "3":
                        ok_e, msg_e = rotacionar_paginas_pdf(arquivo_pdf, 180)
                    elif esc_g == "4":
                        ok_e, msg_e = auto_recortar_bordas_pretas(arquivo_pdf)
                    else:
                        ok_e, msg_e = auto_orientar_pdf(arquivo_pdf)
                    if ok_e:
                        historico.append(
                            AcaoHistorico(
                                tipo="editar_pdf",
                                idx_pdf_anterior=idx_pdf,
                                idx_nome_anterior=idx_nome if idx_nome is not None else 0,
                                pdf_origem=arquivo_pdf,
                                backup_origem=bkp,
                                descricao=msg_e,
                            )
                        )
                        ultimo_pdf_aberto = None
                        ultimo_evento = f"{Cor.VERDE}✨ {msg_e}{Cor.RESET}"
                    else:
                        bkp.unlink(missing_ok=True)
                        ultimo_evento = f"{Cor.AMARELO}⚠️ {msg_e}{Cor.RESET}"
                continue

            elif tecla == "s":
                if item_atual is None or idx_nome is None:
                    ultimo_evento = f"{Cor.AMARELO}⚠️ Não há nome na fila! Use [E] para digitar um nome.{Cor.RESET}"
                    continue

                destino, erro = executar_renomeacao_arquivo(
                    arquivo_pdf, item_atual.nome_formatado, pdfs_dir, log_path
                )
                if destino:
                    item_atual.concluido = True
                    item_atual.arquivo_gerado = destino.name
                    historico.append(
                        AcaoHistorico(
                            tipo="renomear",
                            idx_pdf_anterior=idx_pdf,
                            idx_nome_anterior=idx_nome,
                            pdf_origem=arquivo_pdf,
                            pdf_destino=destino,
                            idx_nome_concluido=idx_nome,
                            descricao=f"{arquivo_pdf.name} -> {destino.name}",
                        )
                    )
                    total_renomeados += 1
                    ultimo_evento = f"{Cor.VERDE}✅ '{arquivo_pdf.name}' -> '{destino.name}'{Cor.RESET}"
                    idx_pdf += 1
                    idx_nome = proximo_indice_nome_pendente(itens_planilha, idx_nome + 1)
                else:
                    erros.append({"arquivo": arquivo_pdf.name, "erro": erro or "Erro desconhecido"})
                    ultimo_evento = f"{Cor.VERMELHO}❌ Erro ao renomear: {erro}{Cor.RESET}"

            elif tecla == "e":
                resultado_edicao = menu_editar_nome(item_atual)
                if resultado_edicao is None:
                    ultimo_evento = "✏️ Edição cancelada."
                    continue

                nome_customizado, avancar_planilha = resultado_edicao
                destino, erro = executar_renomeacao_arquivo(
                    arquivo_pdf, nome_customizado, pdfs_dir, log_path
                )
                if destino:
                    idx_concluido = None
                    if avancar_planilha and item_atual is not None and idx_nome is not None:
                        item_atual.concluido = True
                        item_atual.arquivo_gerado = destino.name
                        idx_concluido = idx_nome

                    historico.append(
                        AcaoHistorico(
                            tipo="renomear",
                            idx_pdf_anterior=idx_pdf,
                            idx_nome_anterior=idx_nome if idx_nome is not None else 0,
                            pdf_origem=arquivo_pdf,
                            pdf_destino=destino,
                            idx_nome_concluido=idx_concluido,
                            descricao=f"Editado: {arquivo_pdf.name} -> {destino.name}",
                        )
                    )
                    total_renomeados += 1
                    total_editados += 1
                    idx_pdf += 1
                    if avancar_planilha and idx_nome is not None:
                        idx_nome = proximo_indice_nome_pendente(itens_planilha, idx_nome + 1)
                    ultimo_evento = f"{Cor.VERDE}✏️ Editado: '{arquivo_pdf.name}' -> '{destino.name}'{Cor.RESET}"
                else:
                    erros.append({"arquivo": arquivo_pdf.name, "erro": erro or "Erro"})
                    ultimo_evento = f"{Cor.VERMELHO}❌ Erro ao renomear editado: {erro}{Cor.RESET}"

            elif tecla == "b":
                idx_escolhido = menu_buscar_na_planilha(itens_planilha)
                if idx_escolhido is None:
                    ultimo_evento = "🔎 Busca cancelada."
                    continue

                item_escolhido = itens_planilha[idx_escolhido]
                destino, erro = executar_renomeacao_arquivo(
                    arquivo_pdf, item_escolhido.nome_formatado, pdfs_dir, log_path
                )
                if destino:
                    item_escolhido.concluido = True
                    item_escolhido.arquivo_gerado = destino.name
                    historico.append(
                        AcaoHistorico(
                            tipo="renomear",
                            idx_pdf_anterior=idx_pdf,
                            idx_nome_anterior=idx_nome if idx_nome is not None else 0,
                            pdf_origem=arquivo_pdf,
                            pdf_destino=destino,
                            idx_nome_concluido=idx_escolhido,
                            descricao=f"Busca: {arquivo_pdf.name} -> {destino.name}",
                        )
                    )
                    total_renomeados += 1
                    idx_pdf += 1
                    if idx_escolhido == idx_nome:
                        idx_nome = proximo_indice_nome_pendente(itens_planilha, idx_nome + 1)
                    ultimo_evento = f"{Cor.VERDE}🔎 '{arquivo_pdf.name}' -> '{destino.name}' (via busca){Cor.RESET}"
                else:
                    erros.append({"arquivo": arquivo_pdf.name, "erro": erro or "Erro"})
                    ultimo_evento = f"{Cor.VERMELHO}❌ Erro ao renomear: {erro}{Cor.RESET}"

            elif tecla == "d":
                res_div = menu_dividir_pdf(arquivo_pdf, backup_divisao_dir)
                if res_div is None:
                    ultimo_evento = "✂️ Divisão cancelada."
                    continue

                novos_pdfs, backup_path = res_div
                fila_pdfs[idx_pdf : idx_pdf + 1] = novos_pdfs
                historico.append(
                    AcaoHistorico(
                        tipo="dividir",
                        idx_pdf_anterior=idx_pdf,
                        idx_nome_anterior=idx_nome if idx_nome is not None else 0,
                        pdf_origem=arquivo_pdf,
                        arquivos_criados=list(novos_pdfs),
                        backup_origem=backup_path,
                        descricao=f"Dividido {arquivo_pdf.name} em {len(novos_pdfs)} partes",
                    )
                )
                total_divididos += 1
                ultimo_pdf_aberto = None
                ultimo_evento = f"{Cor.MAGENTA}✂️ '{arquivo_pdf.name}' dividido em {len(novos_pdfs)} partes!{Cor.RESET}"

            elif tecla == "n":
                if idx_nome is None or item_atual is None:
                    ultimo_evento = f"{Cor.AMARELO}⚠️ Não há mais nomes na fila para pular.{Cor.RESET}"
                    continue
                nome_pulado = item_atual.nome_formatado
                historico.append(
                    AcaoHistorico(
                        tipo="pular_nome",
                        idx_pdf_anterior=idx_pdf,
                        idx_nome_anterior=idx_nome,
                        descricao=f"Pulou nome '{nome_pulado}'",
                    )
                )
                prox = proximo_indice_nome_pendente(itens_planilha, idx_nome + 1)
                if prox == idx_nome:
                    ultimo_evento = f"{Cor.AMARELO}⚠️ Este é o único nome restante na fila.{Cor.RESET}"
                else:
                    idx_nome = prox
                    ultimo_evento = f"{Cor.AMARELO}⏭️ Nome '{nome_pulado}' pulado.{Cor.RESET}"

            elif tecla == "p":
                historico.append(
                    AcaoHistorico(
                        tipo="pular_pdf",
                        idx_pdf_anterior=idx_pdf,
                        idx_nome_anterior=idx_nome if idx_nome is not None else 0,
                        descricao=f"Pulou PDF '{arquivo_pdf.name}'",
                    )
                )
                arquivos_pulados.append(
                    {
                        "arquivo": arquivo_pdf.name,
                        "sugestao": item_atual.nome_formatado if item_atual else "",
                    }
                )
                total_pulados += 1
                idx_pdf += 1
                ultimo_evento = f"{Cor.AMARELO}⏭️ Arquivo '{arquivo_pdf.name}' pulado.{Cor.RESET}"

            elif tecla == "q":
                print(f"\n{Cor.AMARELO}Encerrando sessão a pedido do usuário...{Cor.RESET}")
                break

    except KeyboardInterrupt:
        print(f"\n{Cor.AMARELO}Processo interrompido pelo usuário.{Cor.RESET}")

    finally:
        exportar_planilha_conferencia(
            itens_planilha,
            pdfs_dir,
            script_dir / ARQUIVO_CONFERENCIA_XLSX,
            script_dir / ARQUIVO_CONFERENCIA_CSV,
        )
        print("\n" + "=" * 72)
        print(
            f"📊 {Cor.BOLD}RESUMO DA SESSÃO:{Cor.RESET} "
            f"{Cor.VERDE}{total_renomeados} renomeados{Cor.RESET} | "
            f"{Cor.CIANO}{total_editados} editados{Cor.RESET} | "
            f"{Cor.MAGENTA}{total_divididos} divididos{Cor.RESET} | "
            f"{Cor.AMARELO}{total_pulados} pulados{Cor.RESET}"
        )
        print(f"📥 Planilha de conferência exportada automaticamente: {ARQUIVO_CONFERENCIA_XLSX}")
        print("=" * 72)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Renomeador Semiautomático de PDFs (CLI v6.0)"
    )
    parser.add_argument(
        "--planilha",
        default=NOME_PLANILHA_PADRAO,
        help=f"Caminho da planilha Excel ou CSV (padrão: {NOME_PLANILHA_PADRAO})",
    )
    parser.add_argument(
        "--pasta_pdfs",
        default=PASTA_PDFS_PADRAO,
        help=f"Pasta com os arquivos PDF (padrão: {PASTA_PDFS_PADRAO})",
    )
    parser.add_argument(
        "--auto-iniciar",
        action="store_true",
        help="Inicia direto sem pedir confirmação inicial",
    )
    parser.add_argument(
        "--nao-abrir-pdf",
        action="store_true",
        help="Não abre o visualizador de PDF automaticamente (útil para testes)",
    )
    parser.add_argument(
        "--sem-pausa",
        action="store_true",
        help="Não aguarda Enter ao encerrar",
    )
    main(parser.parse_args())
