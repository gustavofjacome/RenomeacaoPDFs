import os
import subprocess
import pandas as pd
import re
from pathlib import Path
import sys
import fitz  # PyMuPDF
from fuzzywuzzy import fuzz
import unicodedata
from datetime import datetime


class CatalogacaoValidator:
    def __init__(self):
        self.relatorio = {
            'processados': 0,
            'renomeados': 0,
            'nao_catalogados': [],
            'catalogados_duplicados': [],
            'baixa_confianca': [],
            'erros_ocr': []
        }
    
    def gerar_relatorio_final(self):
        """Gera relatório final com sugestões de correção."""
        print(f"\n" + "="*80)
        print(f"📊 RELATÓRIO FINAL DE VALIDAÇÃO DE CATALOGAÇÃO")
        print(f"="*80)
        
        print(f"\n✅ ESTATÍSTICAS GERAIS:")
        print(f"   • Arquivos processados: {self.relatorio['processados']}")
        print(f"   • Arquivos renomeados: {self.relatorio['renomeados']}")
        print(f"   • Documentos não catalogados: {len(self.relatorio['nao_catalogados'])}")
        print(f"   • Possíveis duplicações: {len(self.relatorio['catalogados_duplicados'])}")
        print(f"   • Baixa confiança OCR: {len(self.relatorio['baixa_confianca'])}")
        print(f"   • Erros de OCR: {len(self.relatorio['erros_ocr'])}")
        
        if self.relatorio['nao_catalogados']:
            print(f"\n⚠️ DOCUMENTOS NÃO CATALOGADOS (precisam ser adicionados na planilha):")
            for item in self.relatorio['nao_catalogados']:
                print(f"   • Posição {item['posicao']:3d}: '{item['nome_encontrado']}'")
                print(f"     Arquivo: {item['arquivo']}")
                print(f"     Adicionar na linha {item['posicao']} da planilha de catalogação")
                print()
        
        if self.relatorio['catalogados_duplicados']:
            print(f"\n🔄 POSSÍVEIS DUPLICAÇÕES NA CATALOGAÇÃO:")
            for item in self.relatorio['catalogados_duplicados']:
                print(f"   • Nome: '{item['nome_catalogado']}'")
                print(f"     Posições na planilha: {item['posicoes']}")
                print(f"     Recomendação: Manter apenas uma entrada, remover as outras")
                print()
        
        if self.relatorio['baixa_confianca']:
            print(f"\n🔍 NECESSITAM VERIFICAÇÃO MANUAL (baixa confiança OCR):")
            for item in self.relatorio['baixa_confianca']:
                print(f"   • Posição {item['posicao']:3d}: {item['arquivo']}")
                print(f"     Catalogado: '{item['nome_catalogado']}'")
                print(f"     OCR encontrou: '{item['nome_ocr']}'")
                print(f"     Similaridade: {item['confianca']}%")
                print()
        
        # Salvar relatório em arquivo
        self.salvar_relatorio_arquivo()
    
    def salvar_relatorio_arquivo(self):
        """Salva o relatório detalhado em arquivo texto."""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"relatorio_catalogacao_{timestamp}.txt"
            
            with open(nome_arquivo, 'w', encoding='utf-8') as f:
                f.write(f"RELATÓRIO DE VALIDAÇÃO DE CATALOGAÇÃO\n")
                f.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                f.write(f"{'='*80}\n\n")
                
                f.write(f"ESTATÍSTICAS GERAIS:\n")
                f.write(f"Arquivos processados: {self.relatorio['processados']}\n")
                f.write(f"Arquivos renomeados: {self.relatorio['renomeados']}\n")
                f.write(f"Documentos não catalogados: {len(self.relatorio['nao_catalogados'])}\n")
                f.write(f"Possíveis duplicações: {len(self.relatorio['catalogados_duplicados'])}\n\n")
                
                if self.relatorio['nao_catalogados']:
                    f.write(f"DOCUMENTOS NÃO CATALOGADOS:\n")
                    f.write(f"(Adicionar estes nomes na planilha de catalogação)\n\n")
                    for item in self.relatorio['nao_catalogados']:
                        f.write(f"Posição: {item['posicao']}\n")
                        f.write(f"Nome encontrado: {item['nome_encontrado']}\n")
                        f.write(f"Arquivo: {item['arquivo']}\n")
                        f.write(f"Inserir na linha {item['posicao']} da planilha\n\n")
                
                if self.relatorio['catalogados_duplicados']:
                    f.write(f"DUPLICAÇÕES NA CATALOGAÇÃO:\n")
                    for item in self.relatorio['catalogados_duplicados']:
                        f.write(f"Nome: {item['nome_catalogado']}\n")
                        f.write(f"Posições: {item['posicoes']}\n")
                        f.write(f"Ação: Manter apenas uma entrada\n\n")
            
            print(f"📄 Relatório salvo em: {nome_arquivo}")
            
        except Exception as e:
            print(f"⚠️ Erro ao salvar relatório: {e}")


def limpar_nome_arquivo(nome):
    """Remove caracteres inválidos para nomes de arquivos."""
    caracteres_invalidos = r'[/\\:*?"<>|]'
    nome_limpo = re.sub(caracteres_invalidos, '-', nome)
    nome_limpo = ' '.join(nome_limpo.split())
    nome_limpo = nome_limpo.rstrip('.')
    
    if len(nome_limpo) > 200:
        nome_limpo = nome_limpo[:200].rstrip()
    
    return nome_limpo


def normalizar_texto(texto):
    """Normaliza texto removendo acentos e convertendo para minúsculas."""
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(char for char in texto if unicodedata.category(char) != 'Mn')
    return texto.lower().strip()


def extrair_nome_do_padrao_planilha(nome_padrao):
    """Extrai nome do padrão: PrimeiroNome_SegundoNome_Matricula"""
    partes = nome_padrao.split('_')
    
    if len(partes) < 2:
        return nome_padrao
    
    primeiro_nome = partes[0]
    segundo_nome = partes[1]
    
    artigos = ['da', 'de', 'do', 'das', 'dos', 'o', 'd']
    
    def formatar_palavra(palavra):
        if palavra.lower() in artigos:
            return palavra.lower()
        return palavra.capitalize()
    
    primeiro_formatado = formatar_palavra(primeiro_nome)
    segundo_formatado = formatar_palavra(segundo_nome)
    
    return f"{primeiro_formatado} {segundo_formatado}"


def extrair_texto_pdf_ocr(caminho_pdf):
    """Extrai texto do PDF usando PyMuPDF."""
    try:
        doc = fitz.open(str(caminho_pdf))
        texto_completo = ""
        
        # Extrai texto de todas as páginas (foca nas primeiras para performance)
        max_paginas = min(3, len(doc))  # Máximo 3 páginas
        for pagina_num in range(max_paginas):
            pagina = doc.load_page(pagina_num)
            texto_pagina = pagina.get_text()
            texto_completo += texto_pagina + "\n"
        
        doc.close()
        
        if not texto_completo.strip():
            return False, "Nenhum texto encontrado no PDF"
        
        return True, texto_completo.strip()
        
    except Exception as e:
        return False, f"Erro ao extrair texto do PDF: {str(e)}"


def extrair_possiveis_nomes(texto):
    """Extrai possíveis nomes de pessoa do texto."""
    # Padrões mais específicos para documentos brasileiros
    padroes = [
        r'nome[:\s]*([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+(?:\s+[a-záàâãéêíóôõúç]*\s*[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+)+)',
        r'aluno[:\s]*([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+(?:\s+[a-záàâãéêíóôõúç]*\s*[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+)+)',
        r'estudante[:\s]*([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+(?:\s+[a-záàâãéêíóôõúç]*\s*[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+)+)',
        r'discente[:\s]*([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+(?:\s+[a-záàâãéêíóôõúç]*\s*[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+)+)',
        # Nome completo em linhas separadas (comum em documentos)
        r'^([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+(?:\s+[a-záàâãéêíóôõúç]*\s*[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ][a-záàâãéêíóôõúç]+)+)$',
    ]
    
    nomes_encontrados = []
    
    for linha in texto.split('\n'):
        linha = linha.strip()
        
        # Filtros para evitar falsos positivos
        if len(linha) < 5 or len(linha) > 80:
            continue
        if any(palavra in linha.lower() for palavra in ['página', 'data', 'protocolo', 'processo']):
            continue
            
        for padrao in padroes:
            matches = re.findall(padrao, linha, re.MULTILINE)
            for match in matches:
                if isinstance(match, tuple):
                    match = match[0]
                
                # Valida se é um nome válido
                palavras = match.split()
                if len(palavras) >= 2 and all(len(p) > 1 for p in palavras):
                    # Remove nomes muito comuns que não são pessoas
                    palavras_invalidas = ['ministerio', 'secretaria', 'departamento', 'coordenacao']
                    if not any(inv in match.lower() for inv in palavras_invalidas):
                        nomes_encontrados.append(match.strip())
    
    # Remove duplicatas mantendo ordem
    nomes_unicos = []
    for nome in nomes_encontrados:
        if nome not in nomes_unicos:
            nomes_unicos.append(nome)
    
    return nomes_unicos


def calcular_similaridade(nome1, nome2):
    """Calcula similaridade entre dois nomes."""
    nome1_norm = normalizar_texto(nome1)
    nome2_norm = normalizar_texto(nome2)
    
    # Várias métricas de similaridade
    ratio = fuzz.ratio(nome1_norm, nome2_norm)
    partial_ratio = fuzz.partial_ratio(nome1_norm, nome2_norm)
    token_sort = fuzz.token_sort_ratio(nome1_norm, nome2_norm)
    token_set = fuzz.token_set_ratio(nome1_norm, nome2_norm)
    
    return max(ratio, partial_ratio, token_sort, token_set)


def encontrar_melhor_correspondencia(nomes_pdf, nome_esperado, limiar_minimo=65):
    """Encontra melhor correspondência entre nomes extraídos e esperado."""
    if not nomes_pdf:
        return False, "", 0
    
    melhor_pontuacao = 0
    melhor_nome = ""
    
    for nome_pdf in nomes_pdf:
        pontuacao = calcular_similaridade(nome_pdf, nome_esperado)
        
        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_nome = nome_pdf
    
    encontrou = melhor_pontuacao >= limiar_minimo
    return encontrou, melhor_nome, melhor_pontuacao


def detectar_duplicacoes_catalogacao(nomes_catalogados):
    """Detecta possíveis duplicações na catalogação."""
    duplicacoes = []
    nomes_processados = set()
    
    for i, nome in enumerate(nomes_catalogados):
        nome_normalizado = normalizar_texto(extrair_nome_do_padrao_planilha(nome))
        
        if nome_normalizado in nomes_processados:
            # Encontra todas as posições deste nome
            posicoes = []
            for j, outro_nome in enumerate(nomes_catalogados):
                outro_normalizado = normalizar_texto(extrair_nome_do_padrao_planilha(outro_nome))
                if outro_normalizado == nome_normalizado:
                    posicoes.append(j + 1)  # +1 para linha da planilha
            
            if len(posicoes) > 1:
                duplicacoes.append({
                    'nome_catalogado': extrair_nome_do_padrao_planilha(nome),
                    'posicoes': posicoes
                })
        
        nomes_processados.add(nome_normalizado)
    
    return duplicacoes


def processar_pdf_com_validacao(caminho_pdf, posicao, nome_esperado=None, validator=None):
    """Processa PDF com validação de catalogação."""
    # Extrai texto do PDF
    sucesso, resultado = extrair_texto_pdf_ocr(caminho_pdf)
    if not sucesso:
        if validator:
            validator.relatorio['erros_ocr'].append({
                'arquivo': caminho_pdf.name,
                'posicao': posicao,
                'erro': resultado
            })
        return 'erro_ocr', 0, f"Erro OCR: {resultado}", None
    
    texto_pdf = resultado
    nomes_encontrados = extrair_possiveis_nomes(texto_pdf)
    
    if not nomes_encontrados:
        if validator:
            validator.relatorio['erros_ocr'].append({
                'arquivo': caminho_pdf.name,
                'posicao': posicao,
                'erro': "Nenhum nome encontrado"
            })
        return 'erro_ocr', 0, "Nenhum nome encontrado no OCR", None
    
    # Se tem nome esperado da catalogação
    if nome_esperado:
        nome_esperado_completo = extrair_nome_do_padrao_planilha(nome_esperado)
        encontrou, melhor_nome, pontuacao = encontrar_melhor_correspondencia(
            nomes_encontrados, nome_esperado_completo, limiar_minimo=65
        )
        
        if encontrou and pontuacao >= 75:  # Alta confiança
            return 'renomear', pontuacao, f"Correspondência encontrada: '{melhor_nome}' (Similaridade: {pontuacao}%)", nome_esperado
        elif encontrou:  # Média confiança
            if validator:
                validator.relatorio['baixa_confianca'].append({
                    'arquivo': caminho_pdf.name,
                    'posicao': posicao,
                    'nome_catalogado': nome_esperado_completo,
                    'nome_ocr': melhor_nome,
                    'confianca': pontuacao
                })
            return 'baixa_confianca', pontuacao, f"Baixa confiança: '{melhor_nome}' vs '{nome_esperado_completo}' ({pontuacao}%)", None
        else:
            return 'nao_corresponde', 0, f"Não corresponde ao catalogado: '{nome_esperado_completo}'", None
    
    # Se não tem nome catalogado (documento não catalogado)
    melhor_nome_encontrado = max(nomes_encontrados, key=len) if nomes_encontrados else "Nome não identificado"
    
    if validator:
        validator.relatorio['nao_catalogados'].append({
            'arquivo': caminho_pdf.name,
            'posicao': posicao,
            'nome_encontrado': melhor_nome_encontrado
        })
    
    return 'nao_catalogado', 0, f"Documento não catalogado. Nome encontrado: '{melhor_nome_encontrado}'", None


def verificar_estrutura_pastas():
    """Verifica estrutura de pastas."""
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
    """Carrega nomes da planilha."""
    try:
        script_dir = Path(__file__).parent
        planilha_path = script_dir / "nomes.xlsx"
        
        df = pd.read_excel(planilha_path, engine='openpyxl')
        primeira_coluna = df.iloc[:, 0]
        nomes = primeira_coluna.dropna().astype(str).tolist()
        
        return True, nomes
        
    except Exception as e:
        return False, f"Erro ao carregar planilha: {str(e)}"


def obter_arquivos_pdf():
    """Obtém lista ordenada de PDFs."""
    script_dir = Path(__file__).parent
    pdfs_dir = script_dir / "pdfs"
    
    pdf_files = list(pdfs_dir.glob("*.pdf"))
    
    def extrair_numero(arquivo):
        nome = arquivo.stem
        numeros = re.findall(r'\d+', nome)
        return int(numeros[0]) if numeros else 0
    
    pdf_files.sort(key=extrair_numero)
    return pdf_files


def renomear_arquivo(arquivo_original, novo_nome):
    """Renomeia arquivo PDF."""
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
        return True, f"Renomeado para: {caminho_final.name}"
        
    except Exception as e:
        return False, f"Erro ao renomear: {str(e)}"


def main():
    """Função principal."""
    print("🔍 Validador de Catalogação e Renomeador de PDFs")
    print("=" * 70)
    
    # Verifica dependências
    try:
        import fitz
        from fuzzywuzzy import fuzz
    except ImportError:
        print("❌ Instale as dependências:")
        print("   pip install PyMuPDF fuzzywuzzy python-Levenshtein")
        input("Pressione Enter para sair...")
        return
    
    # Verifica estrutura
    print("📁 Verificando estrutura...")
    sucesso, mensagem = verificar_estrutura_pastas()
    if not sucesso:
        print(f"❌ {mensagem}")
        input("Pressione Enter para sair...")
        return
    print(f"✅ {mensagem}")
    
    # Carrega catalogação
    print("\n📊 Carregando catalogação...")
    sucesso, nomes_catalogados = carregar_nomes_planilha()
    if not sucesso:
        print(f"❌ {nomes_catalogados}")
        input("Pressione Enter para sair...")
        return
    print(f"✅ {len(nomes_catalogados)} nomes catalogados")
    
    # Obtém PDFs
    print("\n📄 Carregando PDFs...")
    arquivos_pdf = obter_arquivos_pdf()
    print(f"✅ {len(arquivos_pdf)} arquivos PDF encontrados")
    
    # Detecta duplicações na catalogação
    print(f"\n🔄 Verificando duplicações na catalogação...")
    duplicacoes = detectar_duplicacoes_catalogacao(nomes_catalogados)
    if duplicacoes:
        print(f"⚠️ Encontradas {len(duplicacoes)} possíveis duplicações")
    else:
        print(f"✅ Nenhuma duplicação detectada")
    
    # Inicializa validador
    validator = CatalogacaoValidator()
    if duplicacoes:
        validator.relatorio['catalogados_duplicados'] = duplicacoes
    
    print(f"\n⚙️ Configurações:")
    print(f"   • Limiar de confiança: 75% (alta) / 65% (baixa)")
    print(f"   • Modo: Validação automática com relatório")
    
    confirmar = input(f"\n▶️ Iniciar processamento? (s/n): ").lower().strip()
    if confirmar not in ['s', 'sim', 'y', 'yes']:
        return
    
    # Processamento principal
    print(f"\n🚀 Processando arquivos...")
    
    for i, arquivo_pdf in enumerate(arquivos_pdf):
        posicao = i + 1
        validator.relatorio['processados'] += 1
        
        print(f"\n📄 [{posicao:3d}] {arquivo_pdf.name}")
        
        # Verifica se há nome catalogado para esta posição
        nome_catalogado = None
        if posicao <= len(nomes_catalogados):
            nome_catalogado = nomes_catalogados[posicao - 1]
        
        # Processa o PDF
        resultado, confianca, detalhes, nome_para_renomear = processar_pdf_com_validacao(
            arquivo_pdf, posicao, nome_catalogado, validator
        )
        
        print(f"     {detalhes}")
        
        # Ações baseadas no resultado
        if resultado == 'renomear' and nome_para_renomear:
            sucesso, msg = renomear_arquivo(arquivo_pdf, nome_para_renomear)
            if sucesso:
                print(f"     ✅ {msg}")
                validator.relatorio['renomeados'] += 1
            else:
                print(f"     ❌ {msg}")
        elif resultado == 'nao_catalogado':
            print(f"     ⚠️ Necessita catalogação na posição {posicao}")
        elif resultado == 'baixa_confianca':
            print(f"     🔍 Necessita verificação manual")
        elif resultado == 'erro_ocr':
            print(f"     ❌ Problema no OCR")
    
    # Gera relatório final
    validator.gerar_relatorio_final()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n⏹️ Processo interrompido")
    except Exception as e:
        print(f"\n\n❌ Erro: {str(e)}")
        input("Pressione Enter para finalizar...")