# 🌐 Renomeador Semiautomático de PDFs — v5.0 (Interface Gráfica Local + CLI)

Sistema completo para renomear PDFs digitalizados em massa de forma **ultrarrápida**, rodando **100% em `localhost` (`127.0.0.1`)** para máxima privacidade de dados sensíveis e **sem exigir permissão de administrador**.

---

## 🚀 Como Iniciar a Interface Gráfica (Web GUI v5.0)

### No Windows (PC do Trabalho)
Basta dar **2 cliques no arquivo `iniciar_gui.bat`**!

### No Terminal (Windows ou Linux)
```bash
python gui.py
```

O sistema iniciará um servidor local privado (`http://127.0.0.1:8765`) e **abrirá automaticamente o navegador** com a tela dividida:
- **Lado Esquerdo (Visualizador de PDF Embutido):** O PDF abre **dentro da própria página** (você não precisa mais dividir janela de terminal com janela de navegador, nem fechar abas!). Além disso, o sistema já faz *pré-carregamento* do próximo PDF em segundo plano para a troca ser instantânea.
- **Lado Direito (Painel de Controle e Fila):** Mostra em destaque o **Nome Completo**, a **Matrícula**, **como o arquivo vai ficar** (`Primeiro_Segundo_Matricula.pdf`), todos os botões de ação e a fila interativa da planilha.

---

## 🔒 Segurança e Privacidade (Dados Sensíveis)

- O servidor roda estritamente no endereço de loopback **`127.0.0.1`** (acessível apenas pela própria máquina).
- **Zero internet / Zero nuvem:** Nenhum arquivo PDF, nome ou matrícula é enviado para a internet. Funciona inclusive com o cabo de rede desconectado.
- **Zero dependências externas obrigatórias:** Usa o servidor nativo do Python (`http.server`) e leitor nativo de `.xlsx`/`.csv`, não precisando instalar Flask, Node.js ou pacotes pesados no computador do trabalho.

---

## ⚡ Superpoderes da Versão GUI (`v5-gui`)

1. **PDF Embutido na Página + Alerta Automático de 2+ Páginas:**
   - Quando um PDF possui **2 ou mais páginas**, um alerta piscante amarelo aparece automaticamente no topo do PDF (`⚠️ 2 páginas — 2+ pessoas? Clique ou [D] p/ dividir`), evitando que duas pessoas no mesmo arquivo passem despercebidas.
2. **Atalhos de 1 Tecla no Navegador (Modo Speed-Run):**
   - Você pode operar tudo pelo mouse **ou** apenas apertando **1 tecla** (sem precisar clicar em nada):
     - **`S`** ou **`Enter`**: **Confirmar e Renomear** (passa para o próximo PDF na hora!).
     - **`E`**: Abre a gaveta de **Editar / Pessoa Fora da Planilha** (você digita `Nome Completo Matrícula` e vê a prévia formatada em tempo real!).
     - **`D`**: Abre a gaveta de **Dividir PDF (2+ pessoas)** — com 1 clique você separa a Página 1 (já renomeando para o aluno atual) e joga a Página 2 para a tela!
     - **`B`** ou **`/`**: Foca na **Busca Instantânea** da planilha para achar qualquer aluno fora de ordem.
     - **`N`**: **Pular Nome da Fila** (quando o aluno da planilha faltou / não tem PDF).
     - **`P`**: **Pular Arquivo PDF** (deixa o PDF atual para depois, mantendo o nome da fila).
     - **`Z`** (ou `Ctrl+Z`): **Desfazer Última Ação** na mesma hora!
3. **Colar Direto do Excel (`Ctrl+C` -> `Ctrl+V`):**
   - Na aba **"📋 Colar do Excel"**, você pode simplesmente copiar as 2 colunas (`Nome Completo` e `Matrícula`) lá no Excel e colar direto na caixa de texto do site! O sistema já remove `de`, `da`, `do`, `dos`, `das`, `e` e monta a fila na hora, sem nem precisar salvar planilha se não quiser.
4. **Fila Interativa com Clique Rápido:**
   - Na aba **"🔎 Fila / Busca"**, você vê todos os nomes da planilha (quem já foi concluído, quem é o atual e quem está pendente).
   - Se um PDF veio fora de ordem, basta clicar em **`🎯 Usar neste PDF`** ao lado do nome da pessoa na lista!
5. **Retomada Inteligente:**
   - Se fechar o navegador ou parar no meio, basta abrir de novo depois: ele detecta sozinho na pasta `pdfs/` quais arquivos e nomes já foram renomeados e continua exatamente dos que faltam.

---

## 📊 Como Preparar os Dados

1. **Planilha (`nomes.xlsx` ou colar na tela):**
   - **Opção recomendada (2 colunas):** Coluna A com `Nome Completo` e Coluna B com `Matrícula`. O sistema formata automaticamente para `Primeiro_Segundo_Matricula.pdf` (ignorando `de`, `da`, `do`, `dos`, `das`, `e`).
   - **Opção clássica (1 coluna):** Coluna A já no formato `Gustavo_Jacome_20201011100005`.
2. **Pasta `pdfs/`:**
   - Coloque os PDFs digitalizados na pasta `pdfs/` (ex: `.001.pdf`, `.002.pdf`, etc.).
   - Se adicionar novos PDFs enquanto o site já estiver aberto, basta clicar no botão **"🔄 Atualizar Pasta PDFs"** no topo da página!

---

## 💻 Modo Terminal (CLI v4.0) também disponível

Se em algum momento quiser rodar apenas no modo texto no terminal, ele continua disponível na mesma pasta:
```bash
python renomear.py
```
E para reverter várias renomeações em lote pelo histórico:
```bash
python desfazer.py
```
