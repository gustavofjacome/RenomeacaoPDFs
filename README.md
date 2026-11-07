# 🌐 Renomeador Semiautomático de PDFs — v6.0 (Web GUI Local + PyMuPDF + OCR Leve)

Sistema completo para renomear, dividir, juntar, girar, recortar e auditar PDFs digitalizados em massa de forma **ultrarrápida**, rodando **100% em `localhost` (`127.0.0.1`)** para máxima privacidade de dados sensíveis e **sem precisar de chave de API ou internet**.

---

## 🚀 Como Iniciar a Interface Gráfica (Web GUI v6.0)

### No Windows (PC do Trabalho)
Basta dar **2 cliques no arquivo `iniciar_gui.bat`**!

### No Terminal (Windows ou Linux)
```bash
source .venv/bin/activate
python gui.py
```

---

## 🎯 Regra Automática das 3 Prioridades de Renomeação

O sistema lê planilhas com 1, 2 ou 3 colunas (`Nome Completo`, `Matrícula`, `CPF`) e aplica automaticamente a hierarquia:

1. **Prioridade 1 (Possui Matrícula válida):**
   Usa `Primeiro_Segundo_Matricula.pdf` (ignorando preposições `de`, `da`, `do`, `dos`, `das`, `e`).
   *Exemplo:* `Gustavo_Ferreira_20201011100005.pdf`
2. **Prioridade 2 (Matrícula vazia / `"Não consta"`, mas possui CPF):**
   Usa `Primeiro_Segundo_CPF.pdf`.
   *Exemplo:* `Gustavo_Ferreira_12345678900.pdf`
3. **Prioridade 3 (Sem Matrícula e sem CPF):**
   Usa o **nome completo separado por `_` em letras minúsculas** (sem acentos).
   *Exemplo:* `gustavo_ferreira_jacome_de_lira.pdf`

---

## ⚡ Superpoderes da Versão v6.0

1. **Juntar PDFs (`[J]`) e Anexar ao Aluno Anterior (`[A]`):**
   - **`[J]` Juntar PDFs:** Une o PDF atual com o próximo PDF da fila (ou $N$ PDFs seguidos) quando um documento de um mesmo aluno foi escaneado separado.
   - **`[A]` + Aluno Anterior:** Acabou de renomear um aluno e percebeu que a folha seguinte ainda era dele? Aperte **`A`** e ela é colada diretamente no final do PDF que você acabou de salvar!
   - **Alerta de Aluno Repetido:** Se o aluno atual já tiver um PDF salvo na pasta, aparece um botão **`📎 Juntar a ele`** para unir as páginas em vez de duplicar.
2. **Girar Páginas (`[R]`) e Auto-Orientação (`🧭 Auto-Orientar`):**
   - Gire páginas em `90°`, `-90°` ou `180°` com 1 clique/tecla, ou clique em **Auto-Orientar** para o sistema colocar páginas "deitadas" na orientação de leitura automaticamente.
3. **Editor Visual de Página (`[C]`) + Remoção de "Bordinhas Pretas" de Scanner:**
   - **🪄 Tirar Bordas Pretas:** Analisa as 4 margens da folha escaneada e remove aquelas faixas pretas de canto de scanner com 1 clique.
   - **✂️ Recorte Manual com o Mouse (`[C]`):** Desenhe um retângulo com o mouse direto sobre a página para recortar exatamente a área desejada (na página atual ou em todas).
   - **⚪ Alerta de Página em Branco + Exclusão/Inversão:** Detecta páginas em branco sozinho (verso vazio de scanner) e permite excluir qualquer página ou **inverter a ordem das páginas** (`🔄 Inverter Ordem`).
4. **Verificação Automática, Diagnóstico de Causa de Erro e Treinamento Contínuo da IA (`memoria_ia.json`):**
   - Extrai o texto do cabeçalho da 1ª página (e também inspeciona a 2ª página quando há múltiplas páginas) para conferir **Matrícula, CPF e Nome** com toda a planilha.
   - **Diagnóstico Inteligente ("O que pode ter acontecido?"):** Quando há divergência ou dúvida (amarelo/vermelho), a IA explica a provável causa do problema e oferece botões de **1 clique** para resolver na hora:
     - *Desalinhamento de fila (aluno faltante)* -> botão para pular o faltante e renomear para o aluno encontrado.
     - *Dois alunos escaneados juntos no mesmo PDF (Pág. 1 ≠ Pág. 2)* -> botão para dividir na Página 1.
     - *Continuação / verso solto do aluno anterior* -> botão para anexar ao PDF anterior (`[A]`).
     - *Documento de familiar no dossiê (Pai/Mãe)* -> ignora linhas de `Filiação:`/`Pai:`/`Mãe:` e alerta se apenas o sobrenome bater.
     - *TCC / Trabalho em Grupo* -> detecta 2+ alunos da planilha na mesma capa e oferece duplicação automática para todos os integrantes.
     - *Página deitada/invertida* -> sugere Auto-Orientar e reanalisar.
   - **Treinamento Contínuo Local (`memoria_ia.json`):**
     - A cada confirmação (ou ao clicar em **Ensinar IA**), o sistema aprende variações de grafia/ruído de OCR (ex: `ferrelra -> ferreira`) e padrões de documentos daquele aluno.
     - Você também pode clicar em **Ignorar Nome** para ensinar a IA a ignorar nomes recorrentes de coordenadores, diretores ou secretários.
5. **Painel de Estatísticas & Métricas de Produtividade (100% Offline):**
   - Nova aba dedicada com indicadores em tempo real:
     - **KPIs:** Concluídos x Pendentes, Velocidade (PDFs/min e seg/arquivo), Tempo de Sessão e Estimativa de Conclusão (ETA).
     - **Regras de Nomenclatura:** Gráficos de progresso segmentados pelas 3 prioridades (Matrícula, CPF e Nome Completo).
     - **Eficiência OCR:** Semáforo consolidado de leituras (Verde OK, Amarelo Revisão, Vermelho Dúvida e Não Analisados).
     - **Volume Diário & Histórico:** Gráfico de colunas de volume de renomeação por dia e tabela detalhada das últimas 30 operações registradas no log.
5. **Exportação de Planilha de Conferência Final (`📥 Baixar Conferência`):**
   - Baixa uma planilha `.xlsx` / `.csv` listando todos os alunos **Concluídos**, **Faltantes (Sem PDF)** e **Extras (Fora da planilha)**.
6. **Desfazer Universal (`[Z]` / `Ctrl+Z`):**
   - Reverte renomeações, divisões, junções, anexações, rotações e recortes na mesma hora!
