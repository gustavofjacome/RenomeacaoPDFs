# 📄 Renomeador Semiautomático de PDFs (CLI v4.0)

Sistema em linha de comando (CLI) para renomear PDFs digitalizados em massa com conferência rápida por 1 tecla, formatação automática de nomes (`Primeiro_Segundo_Matricula`), divisão de PDFs com 2 pessoas no mesmo arquivo e retomada inteligente.

---

## 🏢 Como usar no PC do Trabalho (Sem dor de cabeça e sem Admin)

### 1. Como pegar esta versão no PC do trabalho
Se você já tem a pasta do repositório lá no trabalho, abra o terminal na pasta e rode:
```bash
git fetch origin
git checkout v4-cli
git pull origin v4-cli
```
*(Ou, se preferir nem usar Git no trabalho, basta abrir a branch `v4-cli` no GitHub, clicar no botão verde **Code -> Download ZIP** e extrair a pasta!)*

### 2. Como abrir o programa (Zero esforço)
- **No Windows:** Basta dar **2 cliques no arquivo `iniciar.bat`**! (Ou rodar `python renomear.py` no terminal).
- **Não precisa instalar nada antes:** O script já lê `.xlsx` e `.csv` nativamente com o próprio Python puro (sem precisar instalar `pandas` nem `openpyxl`).
- Se você precisar dividir um PDF de 2 páginas (`[D]`) pela primeira vez no Windows, o próprio programa pergunta na tela se você quer que ele instale o `pypdf` automaticamente no modo usuário (sem precisar de senha de administrador).

---

## ✨ O que mudou na v4.0?

1. **Adeus fórmula no Excel**: Você não precisa mais criar fórmulas para juntar `Primeiro_Segundo_Matricula` nem remover `de`, `da`, `do`, `dos`, `das`, `e` na mão. Basta colocar as **2 colunas (`Nome Completo` e `Matrícula`)** direto em `nomes.xlsx` (ou continuar usando 1 coluna já formatada — ambos funcionam!).
2. **Filas Independentes (Nunca mais desalinha)**:
   - **Pessoa fora da planilha (`[E]` Editar)**: Você digita o nome + matrícula (o script formata sozinho!) e escolhe manter a fila atual para o próximo PDF.
   - **Pessoa da planilha veio fora de ordem (`[B]` Buscar)**: Pesquise parte do nome ou matrícula direto no terminal, selecione a pessoa e renomeie sem bagunçar a fila.
   - **Pessoa da planilha faltou / sem PDF (`[N]` Pular Nome)**: Pula apenas o nome da planilha e continua no mesmo PDF aberto.
3. **2 Pessoas no mesmo PDF (`[D]` Dividir PDF)**:
   - Divide o PDF (ex: Página 1 para a 1ª pessoa e Página 2 para a 2ª pessoa) **direto pelo terminal**, já renomeia a Parte 1 e coloca a Parte 2 como o próximo arquivo da fila!
4. **Retomada Automática (Sem apagar linhas da planilha)**:
   - Se você fechar o programa no meio e abrir de novo depois, ele reconhece sozinho quais arquivos já foram renomeados na pasta `pdfs/` e pula automaticamente os nomes correspondentes da planilha.
5. **Desfazer Imediato (`[Z]`)**:
   - Apertou `[S]` por engano? Aperte `[Z]` na mesma hora para desfazer a última ação e voltar ao arquivo anterior.

---

## 📊 Preparação da Planilha (`nomes.xlsx` ou `.csv`)

Você pode usar `nomes.xlsx` de **duas formas** (o sistema detecta automaticamente):

### Opção A (Recomendada): 2 Colunas (Nome Completo + Matrícula)
Basta colar o nome completo na Coluna A e a matrícula na Coluna B:

| Nome Completo | Matrícula |
| :--- | :--- |
| Gustavo de Oliveira Jácome | 20201011100005 |
| MARIA DAS DORES DA SILVA | 20251213400008 |
| João Pedro dos Santos | 20251213400009 |

O script removerá automaticamente `de`, `da`, `do`, `dos`, `das`, `e` e gerará:
- `Gustavo_Oliveira_20201011100005.pdf`
- `Maria_Dores_20251213400008.pdf`
- `João_Pedro_20251213400009.pdf`

### Opção B: 1 Coluna já formatada (Modo clássico)
| NOME DO PDF |
| :--- |
| Gustavo_Jacome_20201011100005 |
| João_Silva_20251213400007 |

---

## 🎯 Atalhos do Terminal (1 Tecla, sem precisar de Enter)

Coloque os PDFs digitalizados na pasta `pdfs/` (ex: `.001.pdf`, `.002.pdf`, etc.) e abra o `iniciar.bat` (ou `python renomear.py`):

| Tecla | Função | Quando usar? |
| :---: | :--- | :--- |
| **`S`** | **Sim (Confirmar)** | O PDF aberto corresponde ao nome exibido na tela. |
| **`E`** | **Editar / Pessoa Extra** | O PDF é de alguém **que não está na planilha** (ou você quer digitar outro nome). Permite manter o nome atual da fila para o próximo PDF! |
| **`B`** | **Buscar na Planilha** | O PDF é de alguém que **está na planilha, mas fora de ordem**. |
| **`D`** | **Dividir PDF (2+ pessoas)** | Duas ou mais pessoas foram digitalizadas juntas no mesmo PDF. Separa as páginas na hora! |
| **`N`** | **Pular Nome da Fila** | A pessoa sugerida na planilha não tem PDF (faltou), e o PDF aberto já é da próxima pessoa. |
| **`P`** | **Pular Arquivo PDF** | Deixar este arquivo PDF para verificar depois. |
| **`Z`** | **Desfazer Último** | Reverte imediatamente a última ação/renomeação feita na sessão. |
| **`R`** | **Reabrir PDF** | Abre novamente o PDF atual no visualizador padrão. |
| **`Q`** | **Sair** | Salva o relatório e encerra (ao abrir de novo, continua de onde parou!). |

---

## ↩️ Reverter em Lote (`desfazer.bat` ou `desfazer.py`)

Caso queira reverter várias renomeações pelo histórico do arquivo `renomeacao.log`, dê 2 cliques em `desfazer.bat` ou rode:

```bash
# Simular o que seria revertido (sem alterar arquivos)
python desfazer.py --dry-run

# Reverter apenas as últimas 5 renomeações
python desfazer.py --ultimos 5

# Reverter todas as renomeações ativas do log
python desfazer.py
```
