# -*- coding: utf-8 -*-
import os
import subprocess
import pandas as pd
import re
import time
import sys
import argparse
import msvcrt
from pathlib import Path
from datetime import datetime
from collections import Counter

# --- CONSTANTES ---
NOME_PLANILHA_PADRAO = "nomes.xlsx"
PASTA_PDFS_PADRAO = "pdfs"
ARQUIVO_PROGRESSO = ".progresso.txt"
ARQUIVO_LOG = "renomeacao.log"
ARQUIVO_RELATORIO = "relatorio_final.txt"

def limpar_nome_arquivo(nome):
    """Limpa e valida um nome de arquivo para Windows (remove caracteres inválidos)."""
    caracteres_invalidos = r'[/\\:*?"<>|]'
    nome_limpo = re.sub(caracteres_invalidos, '-', nome)
    nome_limpo = ' '.join(nome_limpo.split())
    nome_limpo = nome_limpo.strip('. ')
    return nome_limpo[:200] if len(nome_limpo) > 200 else nome_limpo

def trazer_terminal_para_frente():
    """Tenta trazer o terminal para frente (apenas Windows)."""
    try:
        powershell_cmd = (
            "Add-Type -AssemblyName Microsoft.VisualBasic; "
            "[Microsoft.VisualBasic.Interaction]::AppActivate((Get-Process -Id $PID).MainWindowHandle)"
        )
        subprocess.run(["powershell", "-WindowStyle", "Hidden", "-Command", powershell_cmd],
                       timeout=2, check=False)
    except Exception:
        pass

def validar_planilha(nomes):
    """Verifica duplicatas na planilha e retorna avisos."""
    avisos = []
    duplicados = [n for n, c in Counter(nomes).items() if c > 1]
    if duplicados:
        aviso = f"⚠️ {len(duplicados)} nome(s) duplicado(s) encontrados."
        aviso += f" Exemplo: {', '.join(duplicados[:3])}"
        if len(duplicados) > 3: aviso += "..."
        avisos.append(aviso)
    return avisos

def carregar_dados(planilha_path, pdfs_dir):
    """Carrega planilha e PDFs da pasta."""
    try:
        df = pd.read_excel(planilha_path, engine='openpyxl')
        nomes = df.iloc[:, 0].dropna().astype(str).tolist()
        if not nomes: return None, None, "Planilha vazia."
    except Exception as e:
        return None, None, f"Erro ao carregar planilha: {e}"
    if not pdfs_dir.exists():
        return None, None, f"Pasta '{pdfs_dir}' não encontrada."
    pdf_files = sorted(
        list(pdfs_dir.glob("*.pdf")),
        key=lambda f: int(re.findall(r'\d+', f.stem)[0]) if re.findall(r'\d+', f.stem) else 0
    )
    if not pdf_files: return None, None, "Nenhum PDF encontrado."
    return nomes, pdf_files, "OK"

def abrir_pdf(caminho_pdf):
    """Abre PDF com visualizador padrão."""
    try:
        os.startfile(str(caminho_pdf))
        return True
    except Exception:
        return False

def obter_resposta_rapida():
    """Captura tecla única (sem Enter)."""
    print("\n[s]im | [p]ular | [e]ditar | [r]eabrir | [n/q] sair")
    return msvcrt.getch().decode('utf-8').lower()

def imprimir_barra_progresso(iteracao, total, prefixo='', sufixo='', comprimento=50):
    if total == 0: return
    percentual = f"{100 * (iteracao / float(total)):.1f}"
    preenchido = int(comprimento * iteracao // total)
    barra = '█' * preenchido + '-' * (comprimento - preenchido)
    sys.stdout.write(f'\r{prefixo} |{barra}| {percentual}% {sufixo}')
    sys.stdout.flush()

def registrar_log(mensagem):
    agora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(ARQUIVO_LOG, 'a', encoding='utf-8') as f:
        f.write(f"[{agora}] {mensagem}\n")

def salvar_progresso(index_atual, script_dir):
    with open(script_dir / ARQUIVO_PROGRESSO, 'w') as f:
        f.write(str(index_atual))

def carregar_progresso(script_dir):
    progresso_file = script_dir / ARQUIVO_PROGRESSO
    if progresso_file.exists():
        try:
            return int(progresso_file.read_text().strip())
        except Exception:
            return 0
    return 0

def main(args):
    script_dir = Path(__file__).parent
    planilha_path = Path(args.planilha)
    pdfs_dir = Path(args.pasta_pdfs)

    os.system('cls' if os.name == 'nt' else 'clear')
    print("[INÍCIO] Renomeador Manual Assistido v3.1")
    print("="*60)

    nomes, arquivos_pdf, msg = carregar_dados(planilha_path, pdfs_dir)
    if not nomes:
        print(f"[ERRO] {msg}"); input("Enter para sair..."); return

    print(f"[OK] {len(nomes)} nomes e {len(arquivos_pdf)} PDFs carregados.")

    avisos = validar_planilha(nomes)
    if avisos:
        print("\n".join(avisos))
        if input("Continuar mesmo assim? (s/n): ").lower() != 's':
            return

    inicio = 0
    progresso_salvo = carregar_progresso(script_dir)
    if 0 < progresso_salvo < min(len(arquivos_pdf), len(nomes)):
        if input(f"Retomar do item {progresso_salvo+1}? (s/n): ").lower() == 's':
            inicio = progresso_salvo

    nomes_a_processar = nomes[inicio:]
    arquivos_a_processar = arquivos_pdf[inicio:]

    total_renomeados = total_pulados = total_editados = 0
    arquivos_pulados, erros = [], []

    imprimir_barra_progresso(0, len(arquivos_a_processar), prefixo='Progresso:', sufixo='')

    try:
        for i, (arquivo_pdf, nome_esperado) in enumerate(zip(arquivos_a_processar, nomes_a_processar), start=inicio):
            nome_final = nome_esperado
            editado = False
            while True:
                os.system('cls' if os.name == 'nt' else 'clear')
                print(f"Item {i+1}/{len(nomes)}")
                print(f"Arquivo: {arquivo_pdf.name}\nNome sugerido: {nome_esperado}")

                if not abrir_pdf(arquivo_pdf):
                    erros.append({'arquivo': arquivo_pdf.name, 'erro': 'Falha ao abrir'}); break

                trazer_terminal_para_frente()
                resp = obter_resposta_rapida()

                if resp == 'e':
                    novo_nome = input("Novo nome (sem .pdf): ").strip()
                    if novo_nome:
                        nome_final = limpar_nome_arquivo(novo_nome)
                        editado = True
                        resp = 's'
                    else: continue
                if resp == 'r': continue
                elif resp == 's':
                    destino = pdfs_dir / f"{nome_final}.pdf"
                    contador = 1
                    while destino.exists():
                        destino = pdfs_dir / f"{nome_final}_{contador}.pdf"
                        contador += 1
                    try:
                        arquivo_pdf.rename(destino)
                        registrar_log(f"SUCESSO: '{arquivo_pdf.name}' -> '{destino.name}'")
                        total_renomeados += 1
                        if editado: total_editados += 1
                        salvar_progresso(i+1, script_dir)
                    except Exception as e:
                        erros.append({'arquivo': arquivo_pdf.name, 'erro': str(e)})
                        registrar_log(f"ERRO: {e}")
                elif resp == 'p':
                    arquivos_pulados.append({'arquivo': arquivo_pdf.name, 'sugestao': nome_esperado})
                    total_pulados += 1
                elif resp in ['n', 'q']: raise KeyboardInterrupt
                else: continue
                break
            imprimir_barra_progresso(i - inicio + 1, len(arquivos_a_processar), prefixo='Progresso:', sufixo='')
    except KeyboardInterrupt:
        print("\n[PAUSADO pelo usuário]")

    with open(script_dir / ARQUIVO_RELATORIO, 'w', encoding='utf-8') as f:
        f.write(f"RELATÓRIO {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        f.write(f"Renomeados: {total_renomeados} (editados: {total_editados})\n")
        f.write(f"Pulados: {total_pulados}\n")
        f.write(f"Erros: {len(erros)}\n")

    print("\n[FIM] Relatório salvo.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--planilha', default=NOME_PLANILHA_PADRAO)
    parser.add_argument('--pasta_pdfs', default=PASTA_PDFS_PADRAO)
    main(parser.parse_args())
    input("Pressione Enter para sair...")
