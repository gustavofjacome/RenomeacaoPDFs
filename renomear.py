import os
import subprocess
import pandas as pd
import re
from pathlib import Path
import sys


def limpar_nome_arquivo(nome):
    """
    Remove ou substitui caracteres inválidos para nomes de arquivos no Windows.
    
    Args:
        nome (str): Nome original que pode conter caracteres inválidos
        
    Returns:
        str: Nome limpo, seguro para usar como nome de arquivo
    """
    # Define os caracteres inválidos no Windows:
    caracteres_invalidos = r'[/\\:*?"<>|]'
    
    # Substitui caracteres inválidos por hífen
    nome_limpo = re.sub(caracteres_invalidos, '-', nome)
    
    # Remove espaços extras e caracteres de controle
    nome_limpo = ' '.join(nome_limpo.split())
    
    # Remove pontos no final (não permitidos no Windows)
    nome_limpo = nome_limpo.rstrip('.')
    
    # Limita o tamanho do nome
    if len(nome_limpo) > 200:
        nome_limpo = nome_limpo[:200].rstrip()
    
    return nome_limpo


def verificar_estrutura_pastas():
    """
    Verifica se a estrutura de pastas e arquivos necessários existe.
    
    Returns:
        tuple: (sucesso: bool, mensagem: str)
    """
    script_dir = Path(__file__).parent
    
    # Verifica se a planilha existe
    planilha_path = script_dir / "nomes.xlsx"
    if not planilha_path.exists():
        return False, f"Planilha 'nomes.xlsx' não encontrada em: {planilha_path}"
    
    # Verifica se a pasta pdfs existe
    pdfs_dir = script_dir / "pdfs"
    if not pdfs_dir.exists():
        return False, f"Pasta 'pdfs' não encontrada em: {pdfs_dir}"
    
    # Verifica se há arquivos PDF na pasta
    pdf_files = list(pdfs_dir.glob("*.pdf"))
    if not pdf_files:
        return False, f"Nenhum arquivo PDF encontrado na pasta: {pdfs_dir}"
    
    return True, "Estrutura verificada com sucesso"


def carregar_nomes_planilha():
    """
    Carrega a lista de nomes da primeira coluna da planilha Excel.
    
    Returns:
        tuple: (sucesso: bool, dados: list ou mensagem_erro: str)
    """
    try:
        script_dir = Path(__file__).parent
        planilha_path = script_dir / "nomes.xlsx"
        
        # Lê a planilha Excel
        df = pd.read_excel(planilha_path, engine='openpyxl')
        
        # Pega a primeira coluna
        primeira_coluna = df.iloc[:, 0]
        
        # Remove valores vazios/nulos e converte para lista
        nomes = primeira_coluna.dropna().astype(str).tolist()
        
        if not nomes:
            return False, "Nenhum nome encontrado na primeira coluna da planilha"
        
        return True, nomes
        
    except Exception as e:
        return False, f"Erro ao carregar planilha: {str(e)}"


def obter_arquivos_pdf():
    """
    Obtém lista ordenada de arquivos PDF na pasta 'pdfs'.
    Ordena numericamente baseado no padrão .001.pdf, .002.pdf, etc.
    
    Returns:
        list: Lista de caminhos dos arquivos PDF ordenados
    """
    script_dir = Path(__file__).parent
    pdfs_dir = script_dir / "pdfs"
    
    # Busca todos os arquivos PDF
    pdf_files = list(pdfs_dir.glob("*.pdf"))
    
    def extrair_numero(arquivo):
        """Extrai número do nome do arquivo para ordenação"""
        nome = arquivo.stem
        numeros = re.findall(r'\d+', nome)
        return int(numeros[0]) if numeros else 0
    
    # Ordena os arquivos numericamente
    pdf_files.sort(key=extrair_numero)
    
    return pdf_files


def abrir_pdf(caminho_pdf):
    """
    Abre o arquivo PDF no visualizador padrão do sistema.
    
    Args:
        caminho_pdf (Path): Caminho para o arquivo PDF
        
    Returns:
        bool: True se conseguiu abrir, False caso contrário
    """
    try:
        # Tenta abrir com o comando 'start' do Windows
        subprocess.run(['start', str(caminho_pdf)], shell=True, check=True)
        return True
    except subprocess.CalledProcessError:
        try:
            # Alternativa: usar os.startfile (específico do Windows)
            os.startfile(str(caminho_pdf))
            return True
        except Exception:
            return False


def obter_resposta_usuario(nome_esperado, arquivo_atual):
    """
    Solicita confirmação do usuário se o PDF corresponde ao nome esperado.
    
    Args:
        nome_esperado (str): Nome esperado da planilha
        arquivo_atual (Path): Caminho do arquivo PDF atual
        
    Returns:
        str: 's' para sim, 'n' para não
    """
    print(f"\n{'='*60}")
    print(f"Arquivo atual: {arquivo_atual.name}")
    print(f"Nome esperado: {nome_esperado}")
    print(f"{'='*60}")
    
    while True:
        resposta = input("Este PDF corresponde ao nome esperado? (s/n): ").lower().strip()
        if resposta in ['s', 'sim', 'y', 'yes']:
            return 's'
        elif resposta in ['n', 'não', 'nao', 'no']:
            return 'n'
        else:
            print("Por favor, responda 's' para sim ou 'n' para não.")


def renomear_arquivo(arquivo_original, novo_nome):
    """
    Renomeia o arquivo PDF com o novo nome.
    
    Args:
        arquivo_original (Path): Caminho do arquivo original
        novo_nome (str): Novo nome (sem extensão)
        
    Returns:
        tuple: (sucesso: bool, mensagem: str)
    """
    try:
        # Limpa o nome e adiciona extensão .pdf
        nome_limpo = limpar_nome_arquivo(novo_nome)
        novo_caminho = arquivo_original.parent / f"{nome_limpo}.pdf"
        
        # Verifica se já existe um arquivo com esse nome
        contador = 1
        caminho_final = novo_caminho
        while caminho_final.exists():
            nome_com_contador = f"{nome_limpo}_{contador}"
            caminho_final = arquivo_original.parent / f"{nome_com_contador}.pdf"
            contador += 1
        
        # Renomeia o arquivo
        arquivo_original.rename(caminho_final)
        
        return True, f"Arquivo renomeado para: {caminho_final.name}"
        
    except Exception as e:
        return False, f"Erro ao renomear arquivo: {str(e)}"


def main():
    """
    Função principal do programa.
    Coordena todo o fluxo de renomeação dos arquivos.
    """
    print("🔄 Iniciando Renomeador Automático de PDFs...")
    print("=" * 60)
    
    # Verifica estrutura de pastas
    print("📁 Verificando estrutura de pastas...")
    sucesso, mensagem = verificar_estrutura_pastas()
    if not sucesso:
        print(f"❌ Erro: {mensagem}")
        input("Pressione Enter para sair...")
        return
    print(f"✅ {mensagem}")
    
    # Carrega nomes da planilha
    print("\n📊 Carregando nomes da planilha...")
    sucesso, resultado = carregar_nomes_planilha()
    if not sucesso:
        print(f"❌ Erro: {resultado}")
        input("Pressione Enter para sair...")
        return
    nomes = resultado
    print(f"✅ {len(nomes)} nomes carregados da planilha")
    
    # Obtém arquivos PDF
    print("\n📄 Buscando arquivos PDF...")
    arquivos_pdf = obter_arquivos_pdf()
    print(f"✅ {len(arquivos_pdf)} arquivos PDF encontrados")
    
    # Verifica correspondência entre nomes e arquivos
    if len(nomes) != len(arquivos_pdf):
        print(f"\n⚠️  Atenção: {len(nomes)} nomes na planilha vs {len(arquivos_pdf)} PDFs encontrados")
        continuar = input("Deseja continuar mesmo assim? (s/n): ").lower().strip()
        if continuar not in ['s', 'sim', 'y', 'yes']:
            print("Operação cancelada.")
            return
    
    # Processo principal de renomeação
    print(f"\n🚀 Iniciando processo de renomeação...")
    print("📝 Para cada PDF, confirme se corresponde ao nome esperado:")
    print("   's' = Sim, renomear")
    print("   'n' = Não, parar processo")
    
    total_renomeados = 0
    
    # Loop principal - processa cada PDF
    for i, (arquivo_pdf, nome_esperado) in enumerate(zip(arquivos_pdf, nomes), 1):
        print(f"\n📄 Processando {i}/{min(len(arquivos_pdf), len(nomes))}")
        
        # Abre o PDF para visualização
        print(f"🔍 Abrindo PDF: {arquivo_pdf.name}")
        if not abrir_pdf(arquivo_pdf):
            print("⚠️  Não foi possível abrir o PDF automaticamente.")
            print(f"   Abra manualmente: {arquivo_pdf}")
        
        # Aguarda o usuário visualizar o PDF
        input("\n⏳ Pressione Enter quando o PDF estiver aberto e você puder visualizá-lo...")
        
        # Solicita confirmação do usuário
        resposta = obter_resposta_usuario(nome_esperado, arquivo_pdf)
        
        if resposta == 's':
            # Renomeia o arquivo
            sucesso, mensagem = renomear_arquivo(arquivo_pdf, nome_esperado)
            if sucesso:
                print(f"✅ {mensagem}")
                total_renomeados += 1
            else:
                print(f"❌ {mensagem}")
                break
        else:
            # Para o processo
            print(f"\n⏹️  Processo interrompido pelo usuário no arquivo: {arquivo_pdf.name}")
            print(f"📊 Total de arquivos renomeados: {total_renomeados}")
            break
    
    # Exibe resultado final
    if total_renomeados > 0:
        print(f"\n🎉 Processo concluído!")
        print(f"📊 Total de arquivos renomeados: {total_renomeados}")
    else:
        print(f"\n🔄 Nenhum arquivo foi renomeado.")
    
    input("\n⏳ Pressione Enter para finalizar...")


if __name__ == "__main__":
    """
    Ponto de entrada do programa.
    Trata interrupções e erros inesperados.
    """
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n⏹️  Processo interrompido pelo usuário (Ctrl+C)")
    except Exception as e:
        print(f"\n\n❌ Erro inesperado: {str(e)}")
        print("Por favor, verifique se todos os arquivos estão no local correto.")
        input("Pressione Enter para finalizar...")