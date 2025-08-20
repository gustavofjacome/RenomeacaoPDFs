import os
import subprocess
import pandas as pd
import re
import time
from pathlib import Path
import sys


def limpar_nome_arquivo(nome):
    caracteres_invalidos = r'[/\\:*?"<>|]'
    nome_limpo = re.sub(caracteres_invalidos, '-', nome)
    nome_limpo = ' '.join(nome_limpo.split())
    nome_limpo = nome_limpo.rstrip('.')
    if len(nome_limpo) > 200:
        nome_limpo = nome_limpo[:200].rstrip()
    return nome_limpo


def trazer_terminal_para_frente():
    try:
        powershell_cmd = '''
        Add-Type -TypeDefinition '
        using System;
        using System.Runtime.InteropServices;
        public class Win32 {
            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
            [DllImport("user32.dll")]
            public static extern bool BringWindowToTop(IntPtr hWnd);
            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            [DllImport("kernel32.dll")]
            public static extern IntPtr GetConsoleWindow();
            [DllImport("user32.dll")]
            public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        }';
        $consoleWindow = [Win32]::GetConsoleWindow();
        if ($consoleWindow -ne [IntPtr]::Zero) {
            [Win32]::ShowWindow($consoleWindow, 9);
            [Win32]::BringWindowToTop($consoleWindow);
            [Win32]::SetForegroundWindow($consoleWindow);
            Start-Sleep -Milliseconds 200;
            [Win32]::SetForegroundWindow($consoleWindow);
        }
        '''
        
        subprocess.run([
            "powershell", "-WindowStyle", "Hidden", "-Command", powershell_cmd
        ], capture_output=True, timeout=5)
        
    except Exception:
        try:
            subprocess.run([
                "powershell", "-WindowStyle", "Hidden", "-Command",
                "$console = [console]::WindowHeight; " +
                "Add-Type -AssemblyName Microsoft.VisualBasic; " +
                "[Microsoft.VisualBasic.Interaction]::AppActivate((Get-Process -Name cmd,powershell,WindowsTerminal -ErrorAction SilentlyContinue | Select-Object -First 1).Id); " +
                "Start-Sleep -Milliseconds 300"
            ], capture_output=True, timeout=5)
        except:
            pass


def verificar_estrutura_pastas():
    script_dir = Path(__file__).parent
    
    planilha_path = script_dir / "nomes.xlsx"
    if not planilha_path.exists():
        return False, f"Planilha 'nomes.xlsx' não encontrada em: {planilha_path}"
    
    pdfs_dir = script_dir / "pdfs"
    if not pdfs_dir.exists():
        return False, f"Pasta 'pdfs' não encontrada em: {pdfs_dir}"
    
    pdf_files = list(pdfs_dir.glob("*.pdf"))
    if not pdf_files:
        return False, f"Nenhum arquivo PDF encontrado na pasta: {pdfs_dir}"
    
    return True, "Estrutura verificada com sucesso"


def carregar_nomes_planilha():
    try:
        script_dir = Path(__file__).parent
        planilha_path = script_dir / "nomes.xlsx"
        
        df = pd.read_excel(planilha_path, engine='openpyxl')
        primeira_coluna = df.iloc[:, 0]
        nomes = primeira_coluna.dropna().astype(str).tolist()
        
        if not nomes:
            return False, "Nenhum nome encontrado na primeira coluna da planilha"
        
        return True, nomes
        
    except Exception as e:
        return False, f"Erro ao carregar planilha: {str(e)}"


def obter_arquivos_pdf():
    script_dir = Path(__file__).parent
    pdfs_dir = script_dir / "pdfs"
    
    pdf_files = list(pdfs_dir.glob("*.pdf"))
    
    def extrair_numero(arquivo):
        nome = arquivo.stem
        numeros = re.findall(r'\d+', nome)
        return int(numeros[0]) if numeros else 0
    
    pdf_files.sort(key=extrair_numero)
    
    return pdf_files


def abrir_pdf(caminho_pdf):
    try:
        os.startfile(str(caminho_pdf))
        return True
    except Exception:
        try:
            subprocess.run(['start', str(caminho_pdf)], shell=True, check=True)
            return True
        except Exception:
            return False


def obter_resposta_usuario(nome_esperado, arquivo_atual):
    print(f"\n{'='*60}")
    print(f"Arquivo atual: {arquivo_atual.name}")
    print(f"Nome esperado: {nome_esperado}")
    print(f"{'='*60}")
    
    while True:
        print("\nOpcoes:")
        print("  's' = Sim, renomear este arquivo")
        print("  'n' = Nao, parar processo completamente")
        print("  'p' = Pular este arquivo e continuar")
        print("  'q' = Sair do programa")
        
        resposta = input("\nEste PDF corresponde ao nome esperado? (s/n/p/q): ").lower().strip()
        
        if resposta in ['s', 'sim', 'y', 'yes']:
            return 's'
        elif resposta in ['n', 'não', 'nao', 'no']:
            return 'n'
        elif resposta in ['p', 'pular', 'skip']:
            return 'p'
        elif resposta in ['q', 'quit', 'sair']:
            return 'q'
        else:
            print("\nResposta invalida! Use: s, n, p ou q")


def renomear_arquivo(arquivo_original, novo_nome):
    try:
        nome_limpo = limpar_nome_arquivo(novo_nome)
        novo_caminho = arquivo_original.parent / f"{nome_limpo}.pdf"
        
        contador = 1
        caminho_final = novo_caminho
        while caminho_final.exists():
            nome_com_contador = f"{nome_limpo}_{contador}"
            caminho_final = arquivo_original.parent / f"{nome_com_contador}.pdf"
            contador += 1
        
        arquivo_original.rename(caminho_final)
        
        return True, f"Arquivo renomeado para: {caminho_final.name}"
        
    except Exception as e:
        return False, f"Erro ao renomear arquivo: {str(e)}"


def main():
    print("[INICIO] Renomeador Automatico de PDFs - Windows 10")
    print("=" * 60)
    
    print("[CHECK] Verificando estrutura de pastas...")
    sucesso, mensagem = verificar_estrutura_pastas()
    if not sucesso:
        print(f"[ERRO] {mensagem}")
        input("Pressione Enter para sair...")
        return
    print(f"[OK] {mensagem}")
    
    print("\n[CHECK] Carregando nomes da planilha...")
    sucesso, resultado = carregar_nomes_planilha()
    if not sucesso:
        print(f"[ERRO] {resultado}")
        input("Pressione Enter para sair...")
        return
    nomes = resultado
    print(f"[OK] {len(nomes)} nomes carregados da planilha")
    
    print("\n[CHECK] Buscando arquivos PDF...")
    arquivos_pdf = obter_arquivos_pdf()
    print(f"[OK] {len(arquivos_pdf)} arquivos PDF encontrados")
    
    if len(nomes) != len(arquivos_pdf):
        print(f"\n[AVISO] {len(nomes)} nomes na planilha vs {len(arquivos_pdf)} PDFs encontrados")
        continuar = input("Deseja continuar mesmo assim? (s/n): ").lower().strip()
        if continuar not in ['s', 'sim', 'y', 'yes']:
            print("[CANCELADO] Operacao cancelada pelo usuario.")
            return
    
    print(f"\n[INICIO] Processo de renomeacao iniciado...")
    print("-" * 60)
    
    total_renomeados = 0
    total_pulados = 0
    arquivos_pulados = []
    
    for i, (arquivo_pdf, nome_esperado) in enumerate(zip(arquivos_pdf, nomes), 1):
        print(f"\n[{i:03d}/{min(len(arquivos_pdf), len(nomes)):03d}] PROCESSANDO")
        print(f">> Arquivo: {arquivo_pdf.name}")
        print(f">> Nome esperado: {nome_esperado}")
        
        print("[ACAO] Abrindo PDF...")
        if not abrir_pdf(arquivo_pdf):
            print("[ERRO] Nao foi possivel abrir o PDF automaticamente.")
            print(f"       Abra manualmente: {arquivo_pdf}")
        
        print("[AGUARDE] PDF abrindo...")
        time.sleep(0.8)
        
        print("[ACAO] Trazendo terminal para frente...")
        trazer_terminal_para_frente()
        time.sleep(0.3)
        
        trazer_terminal_para_frente()
        
        input("\n[PAUSA] Pressione Enter quando o PDF estiver aberto...")
        
        resposta = obter_resposta_usuario(nome_esperado, arquivo_pdf)
        
        if resposta == 's':
            sucesso, mensagem = renomear_arquivo(arquivo_pdf, nome_esperado)
            if sucesso:
                print(f"[SUCESSO] {mensagem}")
                total_renomeados += 1
            else:
                print(f"[ERRO] {mensagem}")
                break
                
        elif resposta == 'p':
            print(f"[PULADO] Arquivo {arquivo_pdf.name} foi pulado")
            
            info_pulado = {
                'arquivo': arquivo_pdf.name,
                'nome_esperado': nome_esperado,
                'posicao': i,
                'posicao_planilha': i
            }
            arquivos_pulados.append(info_pulado)
            
            total_pulados += 1
            continue
            
        elif resposta == 'q':
            print(f"\n[SAIDA] Programa encerrado pelo usuario")
            break
            
        else:
            print(f"\n[PARADO] Processo interrompido no arquivo: {arquivo_pdf.name}")
            break
    
    print(f"\n{'='*60}")
    print("[FINAL] RELATORIO DE PROCESSAMENTO")
    print(f"{'='*60}")
    print(f"Arquivos renomeados: {total_renomeados}")
    print(f"Arquivos pulados:    {total_pulados}")
    print(f"Total processado:    {total_renomeados + total_pulados}")
    
    if arquivos_pulados:
        print(f"\n{'-'*60}")
        print("[DETALHES] ARQUIVOS PULADOS - Para adicionar na planilha:")
        print(f"{'-'*60}")
        
        for pulado in arquivos_pulados:
            print(f"Nome: {pulado['nome_esperado']}")
            print(f"  |-> Arquivo original: {pulado['arquivo']}")
            print(f"  |-> Posicao na planilha: linha {pulado['posicao_planilha']}")
            print(f"  |-> (Inserir este nome na linha {pulado['posicao_planilha']} da planilha)")
            print()
        
        print(f"{'-'*60}")
    
    if total_renomeados > 0:
        print(f"[SUCESSO] Processo concluido com {total_renomeados} arquivos renomeados!")
    else:
        print(f"[INFO] Nenhum arquivo foi renomeado.")
    
    input("\n[FIM] Pressione Enter para finalizar...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n[INTERROMPIDO] Processo cancelado pelo usuario (Ctrl+C)")
    except Exception as e:
        print(f"\n\n[ERRO FATAL] Erro inesperado: {str(e)}")
        print("Verifique se todos os arquivos estao no local correto.")
        input("Pressione Enter para finalizar...")