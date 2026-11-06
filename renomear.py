# -*- coding: utf-8 -*-
"""
Renomeador Semiautomático de PDFs (CLI v4.0)
--------------------------------------------
Melhorias principais:
- Lê planilhas (.xlsx, .csv, .tsv) com 2 colunas (Nome Completo + Matrícula) ou 1 coluna já formatada.
- Formata automaticamente: PrimeiroNome_SegundoNome_Matricula (ignorando de, da, do, dos, das, e).
- Funciona mesmo sem pandas/openpyxl instalados (leitor nativo de .xlsx e .csv incluso).
- Filas independentes de PDFs e Nomes (nunca desalinha quando falta alguém ou tem PDF extra).
- Retomada inteligente: detecta automaticamente quais PDFs e nomes já foram renomeados na pasta,
  sem precisar apagar linhas da planilha manualmente ao parar e continuar.
- Botão [E]ditar corrigido: aceita Nome Completo + Matrícula (formata sozinho) e permite indicar
  se é uma pessoa EXTRA (fora da planilha) ou se substitui o nome atual da fila.
- Botão [B]uscar: pesquisa qualquer pessoa da planilha fora de ordem pelo nome ou matrícula.
- Botão [D]ividir PDF: divide PDFs com 2 ou mais pessoas direto pelo terminal (ou duplica PDF de 1 pág).
- Botão [N] Pular Nome (quando a pessoa da planilha não tem PDF) e [P] Pular PDF.
- Botão [Z] Desfazer imediato na própria tela para reverter o último passo sem sair do programa.
"""

import argparse
import csv
import os
import re
import shutil
import subprocess
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Set, Tuple

# --- CONSTANTES ---
NOME_PLANILHA_PADRAO = "nomes.xlsx"
PASTA_PDFS_PADRAO = "pdfs"
PASTA_BACKUP_DIVISAO = ".backup_divididos"
ARQUIVO_LOG = "renomeacao.log"
ARQUIVO_RELATORIO = "relatorio_final.txt"

PARTICULAS_IGNORADAS = {"de", "da", "do", "dos", "das", "e", "del", "d", "d'"}

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
    "id",
}


# --- CORES ANSI (com suporte a Windows) ---
def ativar_cores_windows():
    if os.name == "nt":
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            # ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
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
    concluido: bool = False


@dataclass
class AcaoHistorico:
    tipo: str  # 'renomear', 'pular_nome', 'pular_pdf', 'dividir'
    idx_pdf_anterior: int
    idx_nome_anterior: int
    pdf_origem: Optional[Path] = None
    pdf_destino: Optional[Path] = None
    idx_nome_concluido: Optional[int] = None
    arquivos_criados: List[Path] = field(default_factory=list)
    backup_origem: Optional[Path] = None
    descricao: str = ""


# ==============================================================================
# 1. FORMATAÇÃO DE NOMES E MATRÍCULAS
# ==============================================================================

def limpar_nome_arquivo(nome: str) -> str:
    """Remove caracteres inválidos para arquivos no Windows/Linux."""
    caracteres_invalidos = r'[/\\:*?"<>|]'
    nome_limpo = re.sub(caracteres_invalidos, "-", str(nome))
    nome_limpo = " ".join(nome_limpo.split())
    nome_limpo = nome_limpo.strip(". ")
    return nome_limpo[:200] if len(nome_limpo) > 200 else nome_limpo


def limpar_matricula(valor: str) -> str:
    """Limpa matrícula vinda do Excel (remove .0 de floats, espaços extras, etc.)."""
    if valor is None:
        return ""
    texto = str(valor).strip()
    if not texto or texto.lower() in ("nan", "none"):
        return ""
    # Caso o Excel tenha salvo número inteiro como float (ex: 20201011100005.0)
    if re.match(r"^\d+\.0+$", texto):
        texto = texto.split(".")[0]
    # Caso notação científica do Excel (raro, mas possível se float)
    elif re.match(r"^\d+\.\d+e\+\d+$", texto.lower()):
        try:
            texto = str(int(float(texto)))
        except Exception:
            pass
    return limpar_nome_arquivo(texto).replace(" ", "")


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
    Exemplo: 'Gustavo da Silva Jacome' -> ('Gustavo', 'Silva')
    Exemplo: 'Gustavo Jacome' -> ('Gustavo', 'Jacome')
    """
    # Substituir underscores por espaço caso alguém passe algo semi-formatado
    texto = re.sub(r"[_\-]+", " ", str(nome_completo).strip())
    palavras_brutas = [p.strip(",.;:") for p in texto.split() if p.strip(",.;:")]
    if not palavras_brutas:
        return ()

    palavras_filtradas = [
        p for p in palavras_brutas if p.lower() not in PARTICULAS_IGNORADAS
    ]
    # Se por acaso todas eram partículas (improvável), usa as brutas
    escolhidas = palavras_filtradas if palavras_filtradas else palavras_brutas
    primeiras_duas = escolhidas[:2]
    return tuple(capitalizar_palavra_nome(p) for p in primeiras_duas)


def formatar_nome_e_matricula(
    nome_ou_linha: str,
    matricula: Optional[str] = None,
    forcar_reformatacao: bool = False,
) -> Tuple[str, str, str]:
    """
    Recebe Nome Completo + Matrícula (ou uma string única) e retorna:
    (nome_completo_legivel, matricula_limpa, nome_formatado_final)

    Exemplos:
    - ('Gustavo da Silva Jacome', '20201011100005')
      -> ('Gustavo da Silva Jacome', '20201011100005', 'Gustavo_Silva_20201011100005')
    - ('Gustavo Jacome', '20201011100005')
      -> ('Gustavo Jacome', '20201011100005', 'Gustavo_Jacome_20201011100005')
    - ('Gustavo_Jacome_20201011100005', '')
      -> ('Gustavo Jacome', '20201011100005', 'Gustavo_Jacome_20201011100005')
    - ('Maria das Dores Pereira 20201011100009', '')
      -> ('Maria das Dores Pereira', '20201011100009', 'Maria_Dores_20201011100009')
    """
    nome_str = str(nome_ou_linha or "").strip()
    if nome_str.lower().endswith(".pdf"):
        nome_str = nome_str[:-4].strip()

    mat_str = limpar_matricula(matricula) if matricula else ""

    # Caso 1: Veio apenas 1 coluna e já está com underscores (ex: Gustavo_Jacome_20201011100005)
    if not mat_str and "_" in nome_str and " " not in nome_str:
        partes = [p for p in nome_str.split("_") if p]
        if len(partes) >= 2 and re.match(r"^\d+$", partes[-1]):
            mat_detectada = partes[-1]
            partes_nome = partes[:-1]
            if forcar_reformatacao or len(partes_nome) > 2 or any(
                p.lower() in PARTICULAS_IGNORADAS for p in partes_nome
            ):
                nomes_limpos = extrair_primeiro_segundo_nome(" ".join(partes_nome))
                formatado = "_".join(list(nomes_limpos) + [mat_detectada])
            else:
                formatado = "_".join(partes)
            return " ".join(partes_nome), mat_detectada, limpar_nome_arquivo(formatado)
        return nome_str.replace("_", " "), "", limpar_nome_arquivo(nome_str)

    # Caso 2: Matrícula não veio separada, mas está no final da mesma string
    # Ex: "Gustavo da Silva Jacome 20201011100005" ou "Gustavo Jacome - 20201011100005"
    if not mat_str:
        match_final = re.match(r"^(.+?)[\s\-_;,]+(\d{4,})$", nome_str)
        if match_final:
            nome_str = match_final.group(1).strip()
            mat_str = match_final.group(2).strip()

    partes_nome = extrair_primeiro_segundo_nome(nome_str)
    if not partes_nome:
        return nome_str, mat_str, limpar_nome_arquivo(nome_str or mat_str)

    if mat_str:
        formatado = "_".join(list(partes_nome) + [mat_str])
    else:
        formatado = "_".join(partes_nome)

    return nome_str, mat_str, limpar_nome_arquivo(formatado)


# ==============================================================================
# 2. LEITURA DE PLANILHA (.XLSX NATIVO + OPENPYXL/PANDAS + .CSV/.TSV)
# ==============================================================================

def _ler_xlsx_nativo(caminho: Path) -> List[List[str]]:
    """
    Lê a primeira aba de um arquivo .xlsx usando apenas bibliotecas padrão do Python
    (zipfile + xml.etree.ElementTree). Não requer pandas nem openpyxl!
    """
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(caminho, "r") as z:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            root_ss = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root_ss.findall("s:si", ns):
                texto = "".join((t.text or "") for t in si.findall(".//s:t", ns))
                shared_strings.append(texto)

        # Encontrar a primeira worksheet
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

    # Para .xlsx: tentar openpyxl se instalado, senão leitor nativo
    try:
        import openpyxl  # type: ignore
        wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
        ws = wb.active
        linhas = []
        for row in ws.iter_rows(values_only=True):
            linha_str = [
                "" if val is None else str(val).strip()
                for val in row
            ]
            if any(c != "" for c in linha_str):
                linhas.append(linha_str)
        wb.close()
        return linhas
    except ImportError:
        return _ler_xlsx_nativo(caminho)


def _eh_linha_cabecalho(linha: List[str]) -> bool:
    """Detecta se a primeira linha da planilha é um cabeçalho (ex: 'Nome', 'Matrícula')."""
    celulas_nao_vazias = [c.strip().lower() for c in linha if c.strip()]
    if not celulas_nao_vazias:
        return True
    for c in celulas_nao_vazias:
        if c in CABECALHOS_CONHECIDOS:
            return True
    # Se tiver 2 colunas e a 2ª coluna for texto puro sem dígitos (ex: "Matrícula"), é cabeçalho
    if len(celulas_nao_vazias) >= 2:
        col2 = celulas_nao_vazias[1]
        if not any(ch.isdigit() for ch in col2) and any(
            pal in col2 for pal in ("matr", "insc", "id", "cod", "cód", "num", "núm", "ra")
        ):
            return True
    return False


def _detectar_colunas_nome_matricula(linhas: List[List[str]]) -> Tuple[int, Optional[int], int]:
    """
    Retorna (idx_coluna_nome, idx_coluna_matricula_ou_None, linha_inicial_dados).
    Suporta planilhas com:
    - Coluna A = Nome Completo, Coluna B = Matrícula
    - Coluna A = Matrícula, Coluna B = Nome Completo
    - Coluna A = Nome já formatado (ex: Gustavo_Jacome_20201011100005)
    - Colunas extras (detecta pelo cabeçalho se houver)
    """
    if not linhas:
        return 0, None, 0

    primeira = linhas[0]
    tem_cabecalho = _eh_linha_cabecalho(primeira)
    inicio = 1 if tem_cabecalho else 0

    if tem_cabecalho:
        col_nome = None
        col_mat = None
        for idx, texto in enumerate(primeira):
            t = texto.strip().lower()
            if col_nome is None and any(k in t for k in ("nome", "aluno", "discente", "estudante", "pessoa", "candidato")):
                col_nome = idx
            elif col_mat is None and any(k in t for k in ("matr", "ra", "siape", "registro", "insc", "id", "cpf")):
                col_mat = idx
        if col_nome is not None and col_mat is not None:
            return col_nome, col_mat, inicio

    # Inspecionar primeiras linhas de dados para descobrir as colunas
    amostras = linhas[inicio : inicio + 10]
    if not amostras:
        return 0, None, inicio

    # Verificar quantas colunas não vazias existem em média
    colunas_ativas = set()
    for l in amostras:
        for idx, val in enumerate(l):
            if val.strip():
                colunas_ativas.add(idx)

    cols_ordenadas = sorted(colunas_ativas)
    if len(cols_ordenadas) >= 2:
        c0, c1 = cols_ordenadas[0], cols_ordenadas[1]
        # Verificar se c0 é matrícula (só números) e c1 é nome (letras)
        c0_eh_num = sum(1 for l in amostras if len(l) > c0 and re.match(r"^\d+(\.0+)?$", l[c0].strip()))
        c1_eh_num = sum(1 for l in amostras if len(l) > c1 and re.match(r"^\d+(\.0+)?$", l[c1].strip()))
        if c0_eh_num > c1_eh_num:
            return c1, c0, inicio
        return c0, c1, inicio

    col_unica = cols_ordenadas[0] if cols_ordenadas else 0
    return col_unica, None, inicio


def carregar_planilha(planilha_path: Path) -> Tuple[List[ItemPlanilha], str]:
    """
    Carrega a planilha (.xlsx, .csv, .tsv) e retorna a lista de ItemPlanilha já formatados.
    """
    if not planilha_path.exists():
        return [], f"Planilha '{planilha_path}' não encontrada."

    try:
        linhas = _ler_linhas_brutas_planilha(planilha_path)
    except Exception as e:
        return [], f"Erro ao ler planilha '{planilha_path}': {e}"

    if not linhas:
        return [], f"A planilha '{planilha_path}' está vazia."

    col_nome, col_mat, linha_inicio = _detectar_colunas_nome_matricula(linhas)
    itens: List[ItemPlanilha] = []

    for idx_linha, linha in enumerate(linhas[linha_inicio:], start=1):
        val_nome = linha[col_nome].strip() if len(linha) > col_nome else ""
        val_mat = (
            linha[col_mat].strip()
            if (col_mat is not None and len(linha) > col_mat)
            else ""
        )

        if not val_nome and not val_mat:
            continue
        if val_nome.lower() in ("nan", "none"):
            continue

        nome_completo, matricula, nome_formatado = formatar_nome_e_matricula(val_nome, val_mat)
        if not nome_formatado:
            continue

        itens.append(
            ItemPlanilha(
                indice_original=idx_linha,
                nome_completo=nome_completo,
                matricula=matricula,
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
    """
    Ordena arquivos numericamente de forma natural:
    .001.pdf, .002.pdf, .002_parte1.pdf, .002_parte2.pdf, .010.pdf
    """
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
    """Lê o renomeacao.log para saber quais arquivos já foram resultado de renomeação."""
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
    """
    Determina se um arquivo PDF na pasta `pdfs/` já é um PDF renomeado (concluído)
    ou se é um arquivo bruto de digitalização (ex: .001.pdf, 001.pdf, scan001.pdf).
    """
    stem = pdf_path.stem
    stem_lower = stem.lower()

    # Se começa com ponto (ex: .001.pdf) ou termina com _parte1 / _parte2 (gerado na divisão), é pendente!
    if stem.startswith(".") or re.search(r"_parte\d+$", stem_lower):
        return False

    # Se bate com um nome da planilha (ou sufixo _1, _2 de duplicata)
    base_sem_contador = re.sub(r"_\d+$", "", stem_lower)
    if stem_lower in stems_planilha_lower or base_sem_contador in stems_planilha_lower:
        return True

    # Se consta no log como arquivo renomeado
    if stem_lower in destinos_log_lower:
        return True

    # Se segue claramente o padrão Nome_Sobrenome_Matricula (ex: Maria_Silva_20201011100005)
    if re.match(r"^[A-Za-zÀ-ÿ]+_[A-Za-zÀ-ÿ]+_\d{4,}(_\d+)?$", stem):
        return True

    return False


def sincronizar_pdfs_e_planilha(
    pdfs_dir: Path,
    itens_planilha: List[ItemPlanilha],
    log_path: Path,
) -> Tuple[List[Path], List[Path], int]:
    """
    Separa os PDFs da pasta em (pdfs_pendentes, pdfs_ja_renomeados)
    e marca automaticamente como `concluido=True` os itens da planilha cujo PDF já existe!
    Retorna (pdfs_pendentes, pdfs_ja_renomeados, qtd_nomes_ja_concluidos).
    """
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

    for pdf in todos_pdfs:
        if _eh_pdf_ja_renomeado(pdf, stems_planilha_lower, destinos_log_lower):
            pdfs_ja_renomeados.append(pdf)
            stem_low = pdf.stem.lower()
            stems_renomeados_existentes[stem_low] += 1
            # Também contabilizar caso tenha sufixo de colisão (ex: Nome_Sobrenome_Mat_1)
            base_low = re.sub(r"_\d+$", "", stem_low)
            if base_low != stem_low and base_low in stems_planilha_lower:
                stems_renomeados_existentes[base_low] += 1
        else:
            pdfs_pendentes.append(pdf)

    pdfs_pendentes.sort(key=_chave_ordenacao_natural)
    pdfs_ja_renomeados.sort(key=_chave_ordenacao_natural)

    # Marcar na planilha os nomes que já possuem arquivo renomeado na pasta
    nomes_concluidos = 0
    usados_por_stem: Counter = Counter()
    for item in itens_planilha:
        chave = item.nome_formatado.lower()
        if usados_por_stem[chave] < stems_renomeados_existentes[chave]:
            item.concluido = True
            usados_por_stem[chave] += 1
            nomes_concluidos += 1
        else:
            item.concluido = False

    return pdfs_pendentes, pdfs_ja_renomeados, nomes_concluidos


# ==============================================================================
# 4. DIVISÃO DE PDFs (MULTI-BACKEND: pypdf / PyMuPDF / PyPDF2 / qpdf)
# ==============================================================================

def contar_paginas_pdf(caminho_pdf: Path) -> Optional[int]:
    """Retorna o número de páginas de um PDF usando a ferramenta disponível."""
    # 1. Tentar pypdf / PyPDF2
    for mod_name in ("pypdf", "PyPDF2"):
        try:
            mod = __import__(mod_name)
            reader = mod.PdfReader(str(caminho_pdf))
            return len(reader.pages)
        except ImportError:
            pass
        except Exception:
            break

    # 2. Tentar PyMuPDF (fitz)
    try:
        import fitz  # type: ignore
        with fitz.open(str(caminho_pdf)) as doc:
            return len(doc)
    except ImportError:
        pass
    except Exception:
        pass

    # 3. Tentar qpdf (CLI)
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

    # 4. Fallback leve: contar objetos /Type /Page no binário do PDF
    try:
        dados = caminho_pdf.read_bytes()
        paginas = re.findall(rb"/Type\s*/Page\b(?!s)", dados)
        if paginas:
            return len(paginas)
    except Exception:
        pass

    return None


def _instalar_pypdf_usuario() -> bool:
    """Tenta instalar pypdf no modo --user (sem precisar de admin)."""
    print(f"\n{Cor.CIANO}📦 Instalando biblioteca leve 'pypdf' (--user, sem admin)...{Cor.RESET}")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf", "--user"])
        return True
    except Exception as e:
        print(f"{Cor.VERMELHO}Não foi possível instalar pypdf automaticamente: {e}{Cor.RESET}")
        return False


def dividir_pdf_em_intervalos(
    caminho_pdf: Path,
    intervalos: List[Tuple[int, int]],
) -> Tuple[List[Path], str]:
    """
    Divide `caminho_pdf` em múltiplos arquivos conforme `intervalos` (1-based, inclusivo).
    Exemplo: intervalos = [(1, 1), (2, 2)] cria `<stem>_parte1.pdf` e `<stem>_parte2.pdf`.
    Retorna (lista_de_arquivos_criados, mensagem_status).
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

    # 1. Tentar pypdf ou PyPDF2
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

    # 2. Tentar PyMuPDF (fitz)
    try:
        import fitz  # type: ignore
        with fitz.open(str(caminho_pdf)) as doc:
            for (ini, fim), saida in zip(intervalos, arquivos_saida):
                novo_doc = fitz.open()
                novo_doc.insert_pdf(doc, from_page=ini - 1, to_page=fim - 1)
                novo_doc.save(str(saida))
                novo_doc.close()
        return arquivos_saida, "OK"
    except ImportError:
        pass
    except Exception as e:
        return [], f"Erro ao dividir com PyMuPDF: {e}"

    # 3. Tentar qpdf (disponível em Linux/macOS ou Windows se instalado)
    if shutil.which("qpdf"):
        try:
            for (ini, fim), saida in zip(intervalos, arquivos_saida):
                subprocess.run(
                    [
                        "qpdf",
                        str(caminho_pdf),
                        "--pages",
                        ".",
                        f"{ini}-{fim}",
                        "--",
                        str(saida),
                    ],
                    check=True,
                    capture_output=True,
                )
            return arquivos_saida, "OK"
        except Exception as e:
            return [], f"Erro ao dividir com qpdf: {e}"

    # 4. Se nenhuma biblioteca de divisão estiver instalada, oferecer instalar pypdf (--user)
    print(
        f"\n{Cor.AMARELO}⚠️ Para dividir páginas de PDF, é necessário o pacote 'pypdf' ou 'PyMuPDF'.{Cor.RESET}"
    )
    resp = input("Deseja instalar 'pypdf' agora automaticamente (sem precisar de admin)? (s/n): ").strip().lower()
    if resp == "s" and _instalar_pypdf_usuario():
        return dividir_pdf_em_intervalos(caminho_pdf, intervalos)

    return [], "Biblioteca de divisão de PDF não encontrada (instale com: pip install pypdf --user)."


# ==============================================================================
# 5. INTERAÇÃO COM O SISTEMA (ABRIR PDF, TECLADO, FOCO NO TERMINAL)
# ==============================================================================

def limpar_tela():
    os.system("cls" if os.name == "nt" else "clear")


def trazer_terminal_para_frente():
    """Traz a janela do terminal para frente no Windows após abrir o PDF."""
    if os.name != "nt":
        return False
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        console_window = kernel32.GetConsoleWindow()
        if console_window:
            user32.ShowWindow(console_window, 9)  # SW_RESTORE
            user32.SetForegroundWindow(console_window)
            user32.SetFocus(console_window)
            return True
    except Exception:
        pass
    return False


def abrir_pdf(caminho_pdf: Path) -> bool:
    """Abre o arquivo PDF no visualizador padrão do sistema."""
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
    """
    Lê uma única tecla no terminal sem precisar apertar Enter
    (compatível com Windows via msvcrt e Linux/macOS via termios).
    """
    if os.name == "nt":
        import msvcrt
        while True:
            ch = msvcrt.getwch()
            if ch == "\x03":  # Ctrl+C
                raise KeyboardInterrupt
            if ch in ("\x00", "\xe0"):  # Teclas especiais (setas, F1, etc.)
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
            # Modo não-interativo (ex: testes automatizados via pipe)
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
                if ch == "\x03":  # Ctrl+C
                    raise KeyboardInterrupt
                tecla = ch.lower()
                if tecla in opcoes_validas:
                    sys.stdout.write(tecla.upper() + "\r\n")
                    sys.stdout.flush()
                    return tecla
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, config_antiga)


# ==============================================================================
# 6. UTILITÁRIOS DE INTERFACE E LOG
# ==============================================================================

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
    """Encontra o próximo índice de nome ainda não concluído na planilha."""
    for idx in range(max(0, a_partir_de), len(itens)):
        if not itens[idx].concluido:
            return idx
    for idx in range(0, min(max(0, a_partir_de), len(itens))):
        if not itens[idx].concluido:
            return idx
    return None


def obter_destino_unico(pasta: Path, nome_base: str) -> Path:
    """Garante que o arquivo de destino não sobrescreva outro existente."""
    nome_limpo = limpar_nome_arquivo(nome_base)
    if nome_limpo.lower().endswith(".pdf"):
        nome_limpo = nome_limpo[:-4]
    destino = pasta / f"{nome_limpo}.pdf"
    contador = 1
    while destino.exists():
        destino = pasta / f"{nome_limpo}_{contador}.pdf"
        contador += 1
    return destino


def normalizar_busca(texto: str) -> str:
    """Remove acentos e coloca em minúsculas para busca rápida."""
    nfkd = unicodedata.normalize("NFKD", str(texto))
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sem_acento.lower()


# ==============================================================================
# 7. AÇÕES AVANÇADAS DO MENU (EDITAR, BUSCAR, DIVIDIR, DESFAZER)
# ==============================================================================

def menu_editar_nome(item_sugerido: Optional[ItemPlanilha]) -> Optional[Tuple[str, bool]]:
    """
    Menu para [E]ditar / Digitar outro nome manualmente.
    Retorna (nome_formatado_final, avancar_fila_planilha) ou None se cancelado.
    """
    print("\n" + "─" * 70)
    print(f"{Cor.BOLD}{Cor.AMARELO}✏️  MODO EDIÇÃO / PESSOA FORA DA PLANILHA{Cor.RESET}")
    print(
        f"{Cor.DIM}Dica: Você pode digitar 'Nome Completo Matrícula' (ex: Maria da Silva 20201011100005)\n"
        f"que o sistema já remove de/da/do e formata com '_' automaticamente!{Cor.RESET}"
    )
    entrada = input("\nDigite o Nome (ou Enter para cancelar): ").strip()
    if not entrada:
        return None

    # Se a pessoa digitou nome sem matrícula e sem underscore, perguntar matrícula opcional
    tem_digitos = any(ch.isdigit() for ch in entrada)
    matricula_extra = ""
    if not tem_digitos and "_" not in entrada:
        matricula_extra = input("Digite a Matrícula (ou Enter se já digitou / não tiver): ").strip()

    _, _, nome_formatado = formatar_nome_e_matricula(entrada, matricula_extra)
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
    """
    Permite buscar qualquer pessoa na planilha pelo nome ou matrícula (quando o PDF veio fora de ordem).
    Retorna o índice na lista `itens` ou None se cancelado.
    """
    print("\n" + "─" * 70)
    print(f"{Cor.BOLD}{Cor.CIANO}🔎 BUSCAR PESSOA NA PLANILHA (FORA DE ORDEM){Cor.RESET}")
    termo = input("Digite parte do nome ou matrícula (ou Enter para cancelar): ").strip()
    if not termo:
        return None

    termo_norm = normalizar_busca(termo)
    encontrados: List[int] = []
    for idx, item in enumerate(itens):
        alvo = normalizar_busca(f"{item.nome_completo} {item.matricula} {item.nome_formatado}")
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
) -> Optional[ Tuple[List[Path], Path] ]:
    """
    Menu interativo para dividir um PDF que veio com 2 (ou mais) pessoas digitalizadas juntas.
    Retorna (lista_de_novos_pdfs, caminho_backup_original) ou None se cancelado.
    """
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
        parte1 = arquivo_pdf.parent / f"{arquivo_pdf.stem}_parte1.pdf"
        parte2 = arquivo_pdf.parent / f"{arquivo_pdf.stem}_parte2.pdf"
        parte1 = obter_destino_unico(arquivo_pdf.parent, parte1.stem)
        parte2 = obter_destino_unico(arquivo_pdf.parent, parte2.stem)
        try:
            shutil.copy2(arquivo_pdf, backup_path)
            shutil.copy2(arquivo_pdf, parte1)
            shutil.move(str(arquivo_pdf), str(parte2))
            return [parte1, parte2], backup_path
        except Exception as e:
            print(f"{Cor.VERMELHO}Erro ao duplicar PDF: {e}{Cor.RESET}")
            input("Enter para continuar...")
            return None

    # Dividir por páginas
    novos_pdfs, status = dividir_pdf_em_intervalos(arquivo_pdf, intervalos)
    if not novos_pdfs:
        print(f"{Cor.VERMELHO}❌ {status}{Cor.RESET}")
        input("Enter para continuar...")
        return None

    try:
        shutil.move(str(arquivo_pdf), str(backup_path))
    except Exception as e:
        print(f"{Cor.AMARELO}Aviso: não foi possível mover original para backup: {e}{Cor.RESET}")

    print(
        f"{Cor.VERDE}✅ PDF dividido em {len(novos_pdfs)} partes: "
        f"{', '.join(p.name for p in novos_pdfs)}{Cor.RESET}"
    )
    return novos_pdfs, backup_path


# ==============================================================================
# 8. FLUXO PRINCIPAL INTERATIVO (CLI)
# ==============================================================================

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
    print(f"{Cor.BOLD}{Cor.AZUL}║         📄 RENOMEADOR SEMIAUTOMÁTICO DE PDFs - CLI v4.0              ║{Cor.RESET}")
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
            print(f" 🔢 {Cor.BOLD}MATRÍCULA         :{Cor.RESET} {item_atual.matricula}")
        print(
            f" 🎯 {Cor.BOLD}COMO VAI FICAR    :{Cor.RESET} "
            f"{Cor.VERDE}{Cor.BOLD}{item_atual.nome_formatado}.pdf{Cor.RESET} "
            f"{Cor.DIM}(Item #{item_atual.indice_original}/{total_planilha}){Cor.RESET}"
        )

        # Mostrar prévia dos próximos 2 nomes pendentes da fila para ajudar a identificar desalinhamentos
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
        f" {Cor.BOLD}{Cor.MAGENTA}[D]{Cor.RESET} Dividir PDF (2+ p.) "
        f"{Cor.BOLD}[N]{Cor.RESET} Pular Nome da Fila      "
        f"{Cor.BOLD}[P]{Cor.RESET} Pular Arquivo PDF"
    )
    print(
        f" {Cor.BOLD}[Z]{Cor.RESET} Desfazer Último     "
        f"{Cor.BOLD}[R]{Cor.RESET} Reabrir PDF             "
        f"{Cor.BOLD}{Cor.VERMELHO}[Q]{Cor.RESET} Sair"
    )
    print("─" * 72)
    print(" Escolha uma opção (sem precisar de Enter): ", end="", flush=True)


def executar_renomeacao_arquivo(
    arquivo_pdf: Path,
    nome_final: str,
    pdfs_dir: Path,
    log_path: Path,
) -> Tuple[Optional[Path], Optional[str]]:
    """Renomeia `arquivo_pdf` para `nome_final.pdf` na pasta `pdfs_dir`."""
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
    """Desfaz a última ação registrada no histórico da sessão."""
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
                return (
                    None,
                    None,
                    f"{Cor.VERMELHO}❌ Erro ao reverter arquivo: {e}{Cor.RESET}",
                )
        if acao.idx_nome_concluido is not None and 0 <= acao.idx_nome_concluido < len(itens_planilha):
            itens_planilha[acao.idx_nome_concluido].concluido = False
        return (
            acao.idx_pdf_anterior,
            acao.idx_nome_anterior,
            f"{Cor.AMARELO}↩️ Desfeito: '{acao.pdf_destino.name if acao.pdf_destino else ''}' voltou a ser '{acao.pdf_origem.name if acao.pdf_origem else ''}'{Cor.RESET}",
        )

    elif acao.tipo == "dividir":
        # Remover partes criadas e restaurar backup original
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

    else:
        # pular_nome ou pular_pdf
        return (
            acao.idx_pdf_anterior,
            acao.idx_nome_anterior,
            f"{Cor.AMARELO}↩️ Passo anterior restaurado ({acao.descricao}).{Cor.RESET}",
        )


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
    print(f"{Cor.BOLD}     📄 RENOMEADOR SEMIAUTOMÁTICO DE PDFs - CLI v4.0{Cor.RESET}")
    print("=" * 72)

    itens_planilha, msg = carregar_planilha(planilha_path)
    if not itens_planilha:
        print(f"\n{Cor.VERMELHO}[ERRO] {msg}{Cor.RESET}")
        if not args.sem_pausa:
            input("\nPressione Enter para sair...")
        return

    if not pdfs_dir.exists():
        print(f"\n{Cor.VERMELHO}[ERRO] Pasta de PDFs '{pdfs_dir}' não encontrada.{Cor.RESET}")
        print(f"Crie a pasta '{pdfs_dir.name}' e coloque os arquivos PDF nela.")
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
    print(f"   • PDFs pendentes para renomear: {Cor.CIANO}{len(fila_pdfs)}{Cor.RESET} (e {len(pdfs_ja_renomeados)} já renomeados ignorados)")

    # Verificar nomes duplicados entre os pendentes
    nomes_pendentes_str = [it.nome_formatado for it in itens_planilha if not it.concluido]
    duplicados = [n for n, c in Counter(nomes_pendentes_str).items() if c > 1]
    if duplicados:
        print(
            f"\n{Cor.AMARELO}⚠️ Aviso: {len(duplicados)} nome(s) pendente(s) aparecem mais de uma vez na planilha "
            f"(ex: {', '.join(duplicados[:3])}). O sistema adicionará _1, _2 se necessário.{Cor.RESET}"
        )

    if not fila_pdfs:
        print(f"\n{Cor.VERDE}🎉 Nenhum PDF pendente encontrado em '{pdfs_dir}'! Tudo já está renomeado.{Cor.RESET}")
        if not args.sem_pausa:
            input("\nPressione Enter para sair...")
        return

    if not args.auto_iniciar:
        print(f"\n{Cor.DIM}Exemplo do 1º nome formatado: {itens_planilha[0].nome_formatado}.pdf{Cor.RESET}")
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

            # Só abre o PDF automaticamente se mudou de arquivo (evita reabrir repetidamente ao pular apenas o nome)
            if ultimo_pdf_aberto != arquivo_pdf and not args.nao_abrir_pdf:
                if abrir_pdf(arquivo_pdf):
                    ultimo_pdf_aberto = arquivo_pdf
                    trazer_terminal_para_frente()
                else:
                    erros.append({"arquivo": arquivo_pdf.name, "erro": "Falha ao abrir visualizador"})

            tecla = ler_tecla_unica({"s", "e", "b", "d", "n", "p", "z", "r", "q"})

            # ------------------------------------------------------------------
            # [R] REABRIR PDF ATUAL
            # ------------------------------------------------------------------
            if tecla == "r":
                abrir_pdf(arquivo_pdf)
                trazer_terminal_para_frente()
                ultimo_evento = f"🔄 PDF '{arquivo_pdf.name}' reaberto."
                continue

            # ------------------------------------------------------------------
            # [Z] DESFAZER ÚLTIMA AÇÃO
            # ------------------------------------------------------------------
            elif tecla == "z":
                novo_idx_pdf, novo_idx_nome, msg_undo = desfazer_ultima_acao(
                    historico, fila_pdfs, itens_planilha, log_path
                )
                ultimo_evento = msg_undo
                if novo_idx_pdf is not None:
                    idx_pdf = novo_idx_pdf
                    idx_nome = novo_idx_nome
                    if total_renomeados > 0 and "voltou a ser" in msg_undo:
                        total_renomeados -= 1
                continue

            # ------------------------------------------------------------------
            # [S] SIM - CONFIRMAR NOME ATUAL DA FILA
            # ------------------------------------------------------------------
            elif tecla == "s":
                if item_atual is None or idx_nome is None:
                    ultimo_evento = f"{Cor.AMARELO}⚠️ Não há nome na fila! Use [E] para digitar um nome.{Cor.RESET}"
                    continue

                destino, erro = executar_renomeacao_arquivo(
                    arquivo_pdf, item_atual.nome_formatado, pdfs_dir, log_path
                )
                if destino:
                    item_atual.concluido = True
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
                    input(f"\n{Cor.VERMELHO}Erro ao renomear: {erro}. Pressione Enter...{Cor.RESET}")

            # ------------------------------------------------------------------
            # [E] EDITAR / DIGITAR PESSOA FORA DA PLANILHA
            # ------------------------------------------------------------------
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
                    ultimo_evento = (
                        f"{Cor.VERDE}✏️ Editado: '{arquivo_pdf.name}' -> '{destino.name}' "
                        f"({'fila avançou' if avancar_planilha else 'fila mantida p/ próximo PDF'}){Cor.RESET}"
                    )
                else:
                    erros.append({"arquivo": arquivo_pdf.name, "erro": erro or "Erro"})
                    ultimo_evento = f"{Cor.VERMELHO}❌ Erro ao renomear editado: {erro}{Cor.RESET}"

            # ------------------------------------------------------------------
            # [B] BUSCAR PESSOA NA PLANILHA (FORA DE ORDEM)
            # ------------------------------------------------------------------
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
                    # Se escolheu justamente o atual, avança; se escolheu outro fora de ordem, mantém o atual!
                    if idx_escolhido == idx_nome:
                        idx_nome = proximo_indice_nome_pendente(itens_planilha, idx_nome + 1)
                    ultimo_evento = f"{Cor.VERDE}🔎 '{arquivo_pdf.name}' -> '{destino.name}' (via busca){Cor.RESET}"
                else:
                    erros.append({"arquivo": arquivo_pdf.name, "erro": erro or "Erro"})
                    ultimo_evento = f"{Cor.VERMELHO}❌ Erro ao renomear: {erro}{Cor.RESET}"

            # ------------------------------------------------------------------
            # [D] DIVIDIR PDF (2 OU MAIS PESSOAS NO MESMO ARQUIVO)
            # ------------------------------------------------------------------
            elif tecla == "d":
                res_div = menu_dividir_pdf(arquivo_pdf, backup_divisao_dir)
                if res_div is None:
                    ultimo_evento = "✂️ Divisão cancelada."
                    continue

                novos_pdfs, backup_path = res_div
                # Substituir o arquivo atual na fila pelas novas partes
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
                ultimo_pdf_aberto = None  # Para abrir ou atualizar na sequência

                # Oferecer já renomear a Parte 1 para o nome atual da fila!
                if item_atual is not None and idx_nome is not None and novos_pdfs:
                    parte1 = novos_pdfs[0]
                    print(
                        f"\nDeseja já renomear a 1ª parte ({Cor.AMARELO}{parte1.name}{Cor.RESET}) "
                        f"para {Cor.VERDE}{Cor.BOLD}{item_atual.nome_formatado}.pdf{Cor.RESET}?"
                    )
                    conf = input("[Enter/S] Sim, renomear agora | [N] Não, voltar ao painel: ").strip().lower()
                    if conf in ("", "s", "sim", "y"):
                        destino, erro = executar_renomeacao_arquivo(
                            parte1, item_atual.nome_formatado, pdfs_dir, log_path
                        )
                        if destino:
                            item_atual.concluido = True
                            historico.append(
                                AcaoHistorico(
                                    tipo="renomear",
                                    idx_pdf_anterior=idx_pdf,
                                    idx_nome_anterior=idx_nome,
                                    pdf_origem=parte1,
                                    pdf_destino=destino,
                                    idx_nome_concluido=idx_nome,
                                    descricao=f"{parte1.name} -> {destino.name}",
                                )
                            )
                            total_renomeados += 1
                            idx_pdf += 1
                            idx_nome = proximo_indice_nome_pendente(itens_planilha, idx_nome + 1)
                            ultimo_evento = (
                                f"{Cor.VERDE}✂️ Dividido e Parte 1 renomeada para '{destino.name}'! "
                                f"Abrindo Parte 2...{Cor.RESET}"
                            )
                            continue

                ultimo_evento = (
                    f"{Cor.MAGENTA}✂️ '{arquivo_pdf.name}' dividido em {len(novos_pdfs)} partes!{Cor.RESET}"
                )

            # ------------------------------------------------------------------
            # [N] PULAR NOME DA PLANILHA (PESSOA DA PLANILHA NÃO TEM PDF)
            # ------------------------------------------------------------------
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
                    ultimo_evento = f"{Cor.AMARELO}⏭️ Nome '{nome_pulado}' pulado (mantido mesmo PDF '{arquivo_pdf.name}').{Cor.RESET}"

            # ------------------------------------------------------------------
            # [P] PULAR ARQUIVO PDF ATUAL
            # ------------------------------------------------------------------
            elif tecla == "p":
                print(f"\nAo pular o arquivo '{Cor.AMARELO}{arquivo_pdf.name}{Cor.RESET}':")
                print(f"  {Cor.BOLD}[1]{Cor.RESET} Pular APENAS o PDF (manter o nome '{item_atual.nome_formatado if item_atual else ''}' para o próximo PDF) {Cor.VERDE}[Padrão - Enter]{Cor.RESET}")
                print(f"  {Cor.BOLD}[2]{Cor.RESET} Pular PDF e TAMBÉM pular o nome atual da planilha")
                print(f"  {Cor.BOLD}[C]{Cor.RESET} Cancelar")
                esc_p = input("Escolha [1/2/C]: ").strip().lower()
                if esc_p == "c":
                    ultimo_evento = "Pulo cancelado."
                    continue

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
                if esc_p == "2" and idx_nome is not None:
                    idx_nome = proximo_indice_nome_pendente(itens_planilha, idx_nome + 1)
                ultimo_evento = f"{Cor.AMARELO}⏭️ Arquivo '{arquivo_pdf.name}' pulado.{Cor.RESET}"

            # ------------------------------------------------------------------
            # [Q] SAIR
            # ------------------------------------------------------------------
            elif tecla == "q":
                print(f"\n{Cor.AMARELO}Encerrando sessão a pedido do usuário...{Cor.RESET}")
                break

    except KeyboardInterrupt:
        print(f"\n{Cor.AMARELO}Processo interrompido pelo usuário.{Cor.RESET}")

    finally:
        print("\n" + "=" * 72)
        print(
            f"📊 {Cor.BOLD}RESUMO DA SESSÃO:{Cor.RESET} "
            f"{Cor.VERDE}{total_renomeados} renomeados{Cor.RESET} | "
            f"{Cor.CIANO}{total_editados} editados{Cor.RESET} | "
            f"{Cor.MAGENTA}{total_divididos} divididos{Cor.RESET} | "
            f"{Cor.AMARELO}{total_pulados} pulados{Cor.RESET} | "
            f"{Cor.VERMELHO}{len(erros)} erros{Cor.RESET}"
        )
        restantes_pdfs = max(0, len(fila_pdfs) - idx_pdf)
        restantes_nomes = sum(1 for it in itens_planilha if not it.concluido)
        print(
            f"📌 {Cor.BOLD}Status Geral:{Cor.RESET} Restam {restantes_pdfs} PDF(s) e {restantes_nomes} nome(s) pendentes."
        )
        if restantes_pdfs > 0:
            print(
                f"💡 {Cor.CIANO}Dica: Ao rodar o script novamente, ele continuará automaticamente "
                f"dos {restantes_pdfs} PDFs restantes sem precisar mexer na planilha!{Cor.RESET}"
            )
        print("=" * 72)

        try:
            with open(script_dir / ARQUIVO_RELATORIO, "w", encoding="utf-8") as f:
                f.write(f"RELATÓRIO DE RENOMEAÇÃO - {datetime.now():%Y-%m-%d %H:%M:%S}\n")
                f.write("=" * 60 + "\n")
                f.write(f"Renomeados nesta sessão : {total_renomeados}\n")
                f.write(f"Editados manualmente    : {total_editados}\n")
                f.write(f"PDFs divididos (2+ p.)  : {total_divididos}\n")
                f.write(f"PDFs pulados            : {total_pulados}\n")
                f.write(f"Erros                   : {len(erros)}\n")
                f.write(f"Nomes restantes planilha: {restantes_nomes}/{len(itens_planilha)}\n")

                if arquivos_pulados:
                    f.write("\nARQUIVOS PULADOS:\n")
                    for item in arquivos_pulados:
                        f.write(f"  - {item['arquivo']} (sugestão na hora: {item['sugestao']})\n")

                if erros:
                    f.write("\nERROS:\n")
                    for erro in erros:
                        f.write(f"  - {erro['arquivo']}: {erro['erro']}\n")
        except Exception as e:
            print(f"Aviso: não foi possível salvar relatório: {e}")

        if not args.sem_pausa:
            try:
                input("\nPressione Enter para sair...")
            except (EOFError, KeyboardInterrupt):
                pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Renomeador Semiautomático de PDFs (CLI v4.0)"
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

    try:
        main(parser.parse_args())
    except Exception as e:
        print(f"\nErro crítico inesperado: {e}")
        try:
            input("Pressione ENTER para sair...")
        except (EOFError, KeyboardInterrupt):
            pass
