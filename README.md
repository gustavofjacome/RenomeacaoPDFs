# 📄 Renomeador Automático de PDFs

## 📋 Descrição

Sistema para automatizar a renomeação de arquivos PDF digitalizados usando uma planilha Excel como referência.

O sistema abre cada PDF para visualização, permite conferir se corresponde ao nome esperado, e renomeia automaticamente após confirmação.


## 🗂️ Estrutura de Arquivos

```
projeto/
├── renomear.py          # Script principal
├── nomes.xlsx           # Planilha com os nomes (primeira coluna)
├── pdfs/               # Pasta com os PDFs para renomear
│   ├── .001.pdf
│   ├── .002.pdf
│   ├── .003.pdf
│   └── ...
└── README.md           # Este arquivo
```

## 🚀 Instalação

### 1. Clone o Repositório

```bash
git clone https://github.com/gustavofjacome/RenomeacaoPDFs
cd nome RenomeacaoPDFs

```

### 2. Instale as Dependências

```bash
pip install pandas openpyxl --user 
```

**OBS:** O parâmetro `--user` instala apenas para o usuário atual, dispensando privilégios de administrador.

### 3. Verificar Python

Certifique-se de ter Python instalado:

```bash
python --version
```

## 📊 Preparação dos Dados

### Planilha Excel (nomes.xlsx)

1. Crie uma planilha Excel chamada `nomes.xlsx`
2. Na **primeira coluna (A)**, coloque os nomes desejados, um por linha:

```
Coluna A
--------
João_Silva_20251213400007
Taubaté_da_Silva_20251213400008
Gus_Carvalho_20251213400009
William_Silva_20251213400010
...
```

3. Salve o arquivo na pasta principal do projeto

### Arquivos PDF

1. Coloque os PDFs na pasta `pdfs/`
2. Os arquivos devem seguir padrão numérico: `.001.pdf`, `.002.pdf`, `.003.pdf`, etc.
3. A ordem dos PDFs deve corresponder à ordem dos nomes na planilha

## 🎯 Como Usar

### 1. Execute o Script

```bash
python renomear.py  #certifique-se de estar na pasta do RenomeacaoPDFs para usar este comando
```

### 2. Siga os Comandos do Script

O programa irá:

1. **Verificar** se todos os arquivos necessários existem
2. **Contar** quantos nomes e PDFs foram encontrados
3. **Para cada PDF**:
   - Abrir o arquivo no visualizador padrão
   - Mostrar qual nome é esperado
   - Aguardar sua confirmação:
     - Digite **'s'** se o PDF corresponde ao nome → arquivo será renomeado
     - Digite **'n'** se não corresponde → processo será interrompido

### 3. Exemplo de Uso

```
📄 Processando 1/25
🔍 Abrindo PDF: .001.pdf

⏳ Pressione Enter quando o PDF estiver aberto e você puder visualizá-lo...

============================================================
Arquivo atual: .001.pdf
Nome esperado: Taubaté_da_Silva_20251213400008
============================================================
Este PDF corresponde ao nome esperado? (s/n): s

✅ Arquivo renomeado para: Taubaté_da_Silva_20251213400008.pdf
```

## ⚙️ Características Técnicas

### Tratamento de Nomes
- Remove caracteres inválidos para Windows: `/ \ : * ? " < > |`
- Limita tamanho dos nomes para evitar problemas do sistema
- Adiciona numeração automática em caso de nomes duplicados

### Compatibilidade
- **Windows**: Totalmente compatível
- **Python**: Versão 3.6+

### Dependências
- `pandas`: Leitura de planilhas Excel
- `openpyxl`: Engine para arquivos .xlsx
- `pathlib`: Manipulação de caminhos (incluída no Python)
- `subprocess`: Execução de comandos do sistema (incluída no Python)

## 🛠️ Solução de Problemas

### "pandas não encontrado"
```bash
pip install pandas openpyxl --user --upgrade
```

### "python não é reconhecido"
1. Instale o Python do site oficial: https://python.org
2. Marque "Add to PATH" durante a instalação
3. Reinicie o terminal

### PDF não abre automaticamente
- O script tentará abrir no visualizador padrão
- Se falhar, abra manualmente o arquivo mostrado na tela
- No Windows, o padrão geralmente é o Microsoft Edge

### Planilha não é reconhecida
- Verifique se o arquivo se chama exatamente `nomes.xlsx`
- Certifique-se de que está na mesma pasta do script
- Confirme que os nomes estão na primeira coluna (A)

### Arquivos não são encontrados
- Verifique se a pasta `pdfs/` existe
- Confirme que os PDFs têm extensão `.pdf`
- Certifique-se de que há arquivos na pasta

## 🔒 Segurança

- **Backup recomendado**: Sempre faça backup dos PDFs originais antes de executar
- **Processo reversível**: Mantenha uma cópia dos arquivos originais
- **Interrupção segura**: Use Ctrl+C ou responda 'n' para parar a qualquer momento

**⭐ Se este projeto foi útil, considere dar uma estrela no repositório!**