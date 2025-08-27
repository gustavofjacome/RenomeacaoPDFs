# -*- coding: utf-8 -*-
import os
import subprocess
import pandas as pd
import re
import time
import sys
import argparse
from pathlib import Path
from datetime import datetime
from collections import Counter
import ctypes
import msvcrt

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
    """Traz o terminal para frente usando API do Windows."""
    try:
        # Pega o handle da janela do console atual
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        
        # Pega o handle da janela do console
        console_window = kernel32.GetConsoleWindow()
        if console_window:
            # Traz a janela para frente
            user32.SetForegroundWindow(console_window)
            # Foca na janela
            user32.SetFocus(console_window)
            # Garante que esteja visível
            user32.ShowWindow(console_window, 9)  # SW_RESTORE
            return True
    except Exception as e:
        print(f"Aviso: Não foi possível trazer terminal para frente: {e}")
    return False

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
    """Abre PDF com visualizador padrão - SEM DELAY."""
    try:
        os.startfile(str(caminho_pdf))
        return True
    except Exception as e:
        print(f"Erro ao abrir PDF: {e}")
        return False

def obter_resposta():
    """Captura tecla única sem Enter - MODO RÁPIDO."""
    print("\n[S]im | [P]ular | [E]ditar | [R]eabrir | [N/Q]uit")
    print("Escolha: ", end='', flush=True)
    
    while True:
        try:
            tecla = msvcrt.getch().decode('utf-8').lower()
            if tecla in ['s', 'p', 'e', 'r', 'n', 'q']:
                print(tecla.upper())  # Mostra a tecla pressionada
                return tecla
        except (KeyboardInterrupt, UnicodeDecodeError):
            return 'q'

def imprimir_barra_progresso(iteracao, total, prefixo='', sufixo='', comprimento=50):
    if total == 0: return
    percentual = f"{100 * (iteracao / float(total)):.1f}"
    preenchido = int(comprimento * iteracao // total)
    barra = '█' * preenchido + '-' * (comprimento - preenchido)
    print(f'{prefixo} |{barra}| {percentual}% {sufixo}')

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

def aguardar_so_se_erro(mensagem="ERRO - Enter para continuar..."):
    """Só pausa em caso de erro."""
    try:
        input(mensagem)
    except (KeyboardInterrupt, EOFError):
        pass

def main(args):
    script_dir = Path(__file__).parent
    planilha_path = Path(args.planilha)
    pdfs_dir = Path(args.pasta_pdfs)

    try:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("="*70)
        print("     RENOMEADOR MANUAL ASSISTIDO v3.2 - VERSÃO CORRIGIDA")
        print("="*70)
        print()

        nomes, arquivos_pdf, msg = carregar_dados(planilha_path, pdfs_dir)
        if not nomes:
            print(f"[ERRO] {msg}")
            aguardar_so_se_erro()
            return

        print(f"✅ {len(nomes)} nomes | {len(arquivos_pdf)} PDFs")

        avisos = validar_planilha(nomes)
        if avisos:
            for aviso in avisos:
                print(aviso)
            resposta = input("Continuar? (s/n): ").lower()
            if resposta != 's':
                return

        # Verificar progresso salvo
        inicio = 0
        progresso_salvo = carregar_progresso(script_dir)
        if 0 < progresso_salvo < min(len(arquivos_pdf), len(nomes)):
            resposta = input(f"Retomar do item {progresso_salvo+1}? (s/n): ").lower()
            if resposta == 's':
                inicio = progresso_salvo

        nomes_a_processar = nomes[inicio:]
        arquivos_a_processar = arquivos_pdf[inicio:]

        total_renomeados = total_pulados = total_editados = 0
        arquivos_pulados, erros = [], []

        print("INICIANDO...")
        
        # Loop principal - MODO RÁPIDO
        for i, (arquivo_pdf, nome_esperado) in enumerate(zip(arquivos_a_processar, nomes_a_processar), start=inicio):
            nome_final = nome_esperado
            editado = False
            
            while True:
                os.system('cls')
                
                print(f"ITEM {i+1}/{len(nomes)} | {arquivo_pdf.name}")
                print(f"Nome: {nome_esperado}")
                imprimir_barra_progresso(i - inicio + 1, len(arquivos_a_processar), 
                                        'Progresso:', f'({i+1}/{len(nomes)})')

                # Abrir PDF
                if not abrir_pdf(arquivo_pdf):
                    print("❌ Falha ao abrir PDF!")
                    erros.append({'arquivo': arquivo_pdf.name, 'erro': 'Falha ao abrir'})
                    aguardar_so_se_erro()
                    break

                # Tentar trazer terminal para frente
                trazer_terminal_para_frente()

                # Obter resposta RÁPIDA
                resp = obter_resposta()

                if resp == 'e':
                    novo_nome = input("Novo nome: ").strip()
                    if novo_nome:
                        nome_final = limpar_nome_arquivo(novo_nome)
                        editado = True
                        resp = 's'  # Auto-confirma
                    else:
                        continue

                elif resp == 'r':
                    continue

                elif resp == 's':
                    # Renomear RÁPIDO
                    destino = pdfs_dir / f"{nome_final}.pdf"
                    contador = 1
                    while destino.exists():
                        destino = pdfs_dir / f"{nome_final}_{contador}.pdf"
                        contador += 1
                    
                    try:
                        arquivo_pdf.rename(destino)
                        print(f"✅ -> {destino.name}")
                        registrar_log(f"SUCESSO: '{arquivo_pdf.name}' -> '{destino.name}'")
                        total_renomeados += 1
                        if editado: 
                            total_editados += 1
                        salvar_progresso(i+1, script_dir)
                        
                    except Exception as e:
                        print(f"❌ ERRO: {e}")
                        erros.append({'arquivo': arquivo_pdf.name, 'erro': str(e)})
                        registrar_log(f"ERRO: {e}")
                        aguardar_so_se_erro()

                elif resp == 'p':
                    print("⏭️ PULADO")
                    arquivos_pulados.append({'arquivo': arquivo_pdf.name, 'sugestao': nome_esperado})
                    total_pulados += 1

                elif resp in ['n', 'q']:
                    raise KeyboardInterrupt

                break

    except KeyboardInterrupt:
        print("\nPROCESSO INTERROMPIDO")

    except Exception as e:
        print(f"\nERRO: {e}")
        registrar_log(f"ERRO CRÍTICO: {e}")

    finally:
        # Relatório final RÁPIDO
        print(f"\n📊 RESUMO: {total_renomeados} renomeados | {total_pulados} pulados | {len(erros)} erros")
        
        with open(script_dir / ARQUIVO_RELATORIO, 'w', encoding='utf-8') as f:
            f.write(f"RELATÓRIO - {datetime.now():%Y-%m-%d %H:%M:%S}\n")
            f.write(f"Renomeados: {total_renomeados} | Editados: {total_editados}\n")
            f.write(f"Pulados: {total_pulados} | Erros: {len(erros)}\n")
            
            if arquivos_pulados:
                f.write("\nPULADOS:\n")
                for item in arquivos_pulados:
                    f.write(f"- {item['arquivo']}\n")
            
            if erros:
                f.write("\nERROS:\n")
                for erro in erros:
                    f.write(f"- {erro['arquivo']}: {erro['erro']}\n")

        # Limpar progresso se terminou
        if inicio + len(arquivos_a_processar) >= len(nomes):
            progresso_file = script_dir / ARQUIVO_PROGRESSO
            if progresso_file.exists():
                progresso_file.unlink()

        print("FIM - Relatório salvo")
        input("Enter para sair...")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Renomeador Manual Assistido de PDFs")
    parser.add_argument('--planilha', default=NOME_PLANILHA_PADRAO,
                       help=f"Caminho da planilha Excel (padrão: {NOME_PLANILHA_PADRAO})")
    parser.add_argument('--pasta_pdfs', default=PASTA_PDFS_PADRAO,
                       help=f"Pasta com os PDFs (padrão: {PASTA_PDFS_PADRAO})")
    
    try:
        main(parser.parse_args())
    except Exception as e:
        print(f"Erro crítico: {e}")
        input("Pressione ENTER para sair...")