let ESTADO = null;
let urlPdfAtualNoIframe = "";
let paginaCropAtual = 1;
let arrastandoCrop = false;
let cropInicio = { x: 0, y: 0 };
let cropSelecao = null;

const PARTICULAS = new Set(["de", "da", "do", "dos", "das", "e", "del", "d", "d'"]);
const NAO_CONSTA = new Set([
  "", "nan", "none", "-", "--", "s/m", "sm", "n/c", "nc",
  "nao consta", "não consta", "sem matricula", "sem matrícula"
]);

function removerAcentosJS(s) {
  return (s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function ehNaoConstaJS(v) {
  const limpo = removerAcentosJS(v || "").toLowerCase().trim();
  return NAO_CONSTA.has(limpo) || limpo.includes("nao consta") || limpo.includes("sem matricula");
}

function formatarPreview3Prioridades(nomeBruto, matBruta, cpfBruto) {
  let s = (nomeBruto || "").trim();
  if (!s) return "—";
  if (s.toLowerCase().endsWith(".pdf")) s = s.slice(0, -4).trim();

  let mat = ehNaoConstaJS(matBruta) ? "" : (matBruta || "").trim().replace(/\s+/g, "");
  let cpf = ehNaoConstaJS(cpfBruto) ? "" : (cpfBruto || "").replace(/\D/g, "");

  if (!mat && !cpf) {
    const mCpf = s.match(/^(.+?)[\s\-_;,]+(\d{3}\.?\d{3}\.?\d{3}[\-\.]?\d{2})$/);
    if (mCpf) {
      s = mCpf[1].trim();
      cpf = mCpf[2].replace(/\D/g, "");
    } else {
      const mNum = s.match(/^(.+?)[\s\-_;,]+(\d{4,})$/);
      if (mNum) {
        s = mNum[1].trim();
        const num = mNum[2].trim();
        if (num.length === 11) cpf = num;
        else mat = num;
      }
    }
  }

  s = s.replace(/[\s\-_;,]+(n[aã]o\s+consta|sem\s+matr[ií]cula|s\/m|n\/c)$/i, "").trim();
  const palavras = s.replace(/[_\-]+/g, " ").split(/\s+/).filter(Boolean);
  if (!palavras.length) return "—";

  const filtradas = palavras.filter(p => !PARTICULAS.has(p.toLowerCase()));
  const duasPrimeiras = (filtradas.length ? filtradas : palavras).slice(0, 2).map(p => {
    return p.charAt(0).toUpperCase() + p.slice(1).toLowerCase();
  });

  if (mat && /\d/.test(mat)) {
    return [...duasPrimeiras, mat].join("_") + ".pdf (Prioridade 1: Matrícula)";
  }
  if (cpf) {
    return [...duasPrimeiras, cpf].join("_") + ".pdf (Prioridade 2: CPF)";
  }
  const completoMin = palavras
    .map(p => removerAcentosJS(p).replace(/[^a-zA-Z0-9]/g, "").toLowerCase())
    .filter(Boolean)
    .join("_");
  return (completoMin || "—") + ".pdf (Prioridade 3: Nome Completo)";
}

function atualizarPreviewEdicao() {
  const n = document.getElementById("inp-edit-nome").value;
  const m = document.getElementById("inp-edit-mat").value;
  const c = document.getElementById("inp-edit-cpf").value;
  document.getElementById("preview-edit-formatado").textContent =
    "Arquivo: " + formatarPreview3Prioridades(n, m, c);
}

/* ========================================================================= */
/* NAVEGAÇÃO DE ABAS E MODAL DE ATALHOS                                      */
/* ========================================================================= */
function alternarAba(nomeAba) {
  document.getElementById("aba-operacao").classList.toggle("active", nomeAba === "operacao");
  document.getElementById("aba-dados").classList.toggle("active", nomeAba === "dados");
  document.getElementById("aba-estatisticas").classList.toggle("active", nomeAba === "estatisticas");
  document.getElementById("tab-btn-operacao").classList.toggle("active", nomeAba === "operacao");
  document.getElementById("tab-btn-dados").classList.toggle("active", nomeAba === "dados");
  document.getElementById("tab-btn-estatisticas").classList.toggle("active", nomeAba === "estatisticas");

  if (nomeAba === "estatisticas" && ESTADO && ESTADO.estatisticas) {
    renderizarPainelEstatisticas(ESTADO.estatisticas);
  }
}

function abrirModalAtalhos() {
  document.getElementById("modal-atalhos").classList.add("open");
}

function fecharModalAtalhos() {
  document.getElementById("modal-atalhos").classList.remove("open");
}

function alternarModalAtalhos() {
  document.getElementById("modal-atalhos").classList.toggle("open");
}

/* ========================================================================= */
/* COMUNICAÇÃO COM O SERVIDOR LOCAL                                          */
/* ========================================================================= */
async function apiPost(rota, dados = {}) {
  try {
    const resp = await fetch(rota, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(dados)
    });
    const json = await resp.json();
    if (json.estado) {
      aplicarEstado(json.estado);
    }
    if (!json.ok && json.mensagem) {
      mostrarStatus(json.mensagem, true);
    }
    return json;
  } catch (err) {
    mostrarStatus("Erro de comunicação local", true);
  }
}

async function carregarEstadoInicial() {
  const resp = await fetch("/api/state");
  const json = await resp.json();
  aplicarEstado(json);
}

function mostrarStatus(msg, erro = false) {
  const el = document.getElementById("status-toast");
  el.textContent = msg;
  el.style.color = erro ? "#ef4444" : "#a1a1aa";
}

function aplicarEstado(st) {
  ESTADO = st;
  const p = st.progresso;
  document.getElementById("lbl-progresso").textContent =
    `${p.concluidos_planilha} / ${p.total_planilha} (${p.percentual}%)`;
  document.getElementById("barra-progresso").style.width = `${p.percentual}%`;
  document.getElementById("lbl-renomeados").textContent = p.renomeados_sessao;
  if (st.memoria_ia && document.getElementById("lbl-ia-regras")) {
    document.getElementById("lbl-ia-regras").textContent =
      `${st.memoria_ia.total_regras} regras (${st.memoria_ia.total_confirmacoes} exp.)`;
  }
  document.getElementById("lbl-tab-total").textContent = p.total_planilha;
  document.getElementById("lbl-planilha-ativa").textContent = `Planilha ativa: ${st.planilha_nome}`;
  document.getElementById("btn-undo").disabled = !p.pode_desfazer;
  document.getElementById("btn-anexar-ant").disabled = !st.ultimo_renomeado;
  document.getElementById("btn-anexar-ant").title = st.ultimo_renomeado
    ? `Anexar este PDF ao final de '${st.ultimo_renomeado}' (Tecla A)`
    : "Renomeie pelo menos 1 PDF antes para poder anexar ao anterior";

  mostrarStatus(st.ultimo_evento || "Pronto", Boolean(st.erro_carregamento));

  const iframe = document.getElementById("pdf-frame");
  const empty = document.getElementById("empty-state");
  const badgeBrancas = document.getElementById("badge-brancas");

  if (st.tem_pdf && st.pdf_atual) {
    iframe.style.display = "block";
    empty.style.display = "none";
    document.getElementById("lbl-pdf-nome").textContent = st.pdf_atual.nome;
    document.getElementById("lbl-preview-origem").textContent = st.pdf_atual.nome;
    document.getElementById("lbl-pdf-pos").textContent =
      `(${st.pdf_atual.posicao} de ${st.pdf_atual.total_fila})`;

    const badge = document.getElementById("badge-paginas");
    const pags = st.pdf_atual.paginas;
    if (pags && pags >= 2) {
      badge.innerHTML = `<i class="fa-solid fa-layer-group"></i> ${pags} páginas — Dividir (D)`;
      badge.className = "tag-badge warn";
    } else if (pags === 1) {
      badge.textContent = "1 página";
      badge.className = "tag-badge";
    } else {
      badge.textContent = "PDF";
      badge.className = "tag-badge";
    }

    const brancas = st.pdf_atual.paginas_brancas || [];
    if (brancas.length > 0) {
      badgeBrancas.style.display = "inline-block";
      badgeBrancas.innerHTML = `<i class="fa-solid fa-trash-can"></i> Pág. ${brancas.join(", ")} em branco (Remover)`;
    } else {
      badgeBrancas.style.display = "none";
    }

    const novaUrl = st.pdf_atual.url + "#toolbar=1&navpanes=0&view=FitH";
    if (urlPdfAtualNoIframe !== novaUrl) {
      urlPdfAtualNoIframe = novaUrl;
      iframe.src = novaUrl;
      if (document.getElementById("crop-overlay").classList.contains("open")) {
        paginaCropAtual = 1;
        carregarImagemCrop();
      }
    }

    document.getElementById("pdf-prefetch").href = st.pdf_proximo_url || "";
  } else {
    iframe.style.display = "none";
    empty.style.display = "block";
    badgeBrancas.style.display = "none";
    fecharModoRecorte();
    urlPdfAtualNoIframe = "";
    iframe.src = "about:blank";
    document.getElementById("lbl-pdf-nome").textContent = "—";
    document.getElementById("lbl-preview-origem").textContent = "—";
    document.getElementById("lbl-pdf-pos").textContent = "";
    document.getElementById("badge-paginas").textContent = "—";
    document.getElementById("badge-paginas").className = "tag-badge";
    if (st.erro_carregamento) {
      document.getElementById("empty-msg").textContent = st.erro_carregamento;
    }
  }

  const bannerDup = document.getElementById("banner-duplicado");
  if (st.item_atual) {
    document.getElementById("lbl-item-idx").textContent =
      `#${st.item_atual.indice_original}/${p.total_planilha}`;
    document.getElementById("lbl-aluno-nome").textContent = st.item_atual.nome_completo;

    const prio = st.item_atual.prioridade;
    const lblPrio = document.getElementById("lbl-prioridade");
    if (prio === 1) {
      lblPrio.textContent = "Prioridade 1: Matrícula";
      document.getElementById("lbl-aluno-mat").textContent = `Matrícula: ${st.item_atual.matricula}`;
    } else if (prio === 2) {
      lblPrio.textContent = "Prioridade 2: CPF";
      document.getElementById("lbl-aluno-mat").textContent = `Sem matrícula · CPF: ${st.item_atual.cpf}`;
    } else {
      lblPrio.textContent = "Prioridade 3: Nome Completo";
      document.getElementById("lbl-aluno-mat").textContent = "Sem matrícula e sem CPF (usa nome_completo)";
    }

    document.getElementById("lbl-aluno-formatado").textContent =
      `${st.item_atual.nome_formatado}.pdf`;

    if (st.item_atual.arquivo_existente) {
      bannerDup.style.display = "flex";
      document.getElementById("txt-duplicado").textContent =
        `Já existe '${st.item_atual.arquivo_existente}' na pasta.`;
    } else {
      bannerDup.style.display = "none";
    }
  } else {
    document.getElementById("lbl-item-idx").textContent = "Concluído";
    document.getElementById("lbl-prioridade").textContent = "—";
    document.getElementById("lbl-aluno-nome").textContent = "Fila finalizada";
    document.getElementById("lbl-aluno-mat").textContent = "Use Editar (E) para informar outro nome";
    document.getElementById("lbl-aluno-formatado").textContent = "—";
    bannerDup.style.display = "none";
  }

  // Banner de Diagnóstico OCR + Análise Inteligente de Causa de Erro + Botões de 1 Clique
  const cardAtual = document.getElementById("card-aluno-atual");
  const bannerIA = document.getElementById("banner-ia");
  const icoIA = document.getElementById("ico-ia");
  const txtIA = document.getElementById("txt-ia");
  const boxDiagCausa = document.getElementById("box-diag-causa");
  const lblDiagTitulo = document.getElementById("lbl-diag-titulo");
  const lblDiagExp = document.getElementById("lbl-diag-explicacao");
  const boxSugestoes = document.getElementById("box-sugestoes-ia");
  const ia = st.verificacao_ia;

  cardAtual.className = "current-card";
  if (boxSugestoes) boxSugestoes.innerHTML = "";
  if (boxDiagCausa) boxDiagCausa.style.display = "none";

  if (ia && st.tem_pdf) {
    const cor = ia.nivel_cor || "vermelho";
    cardAtual.classList.add(`borda-${cor}`);
    bannerIA.style.display = "flex";
    bannerIA.className = `status-banner ai-diagnostic-panel ${cor}`;

    if (ia.status === "confirmado_atual") {
      icoIA.className = "fa-solid fa-circle-check";
      txtIA.textContent = `OCR Confirmado: ${ia.motivo}`;
    } else if (ia.status === "tcc_grupo") {
      icoIA.className = "fa-solid fa-users";
      txtIA.textContent = `Revisão (Amarelo): ${ia.motivo}`;
    } else if (ia.status === "dois_alunos_paginas_diferentes") {
      icoIA.className = "fa-solid fa-scissors";
      txtIA.textContent = ia.motivo;
    } else if (ia.status === "encontrado_outro") {
      icoIA.className = "fa-solid fa-triangle-exclamation";
      txtIA.textContent = `Fora de ordem: PDF é de #${ia.indice_original} ${ia.nome_completo} (${ia.confianca}%)`;
    } else if (ia.status === "duvida_familiar") {
      icoIA.className = "fa-solid fa-user-shield";
      txtIA.textContent = `Revisão Familiar: ${ia.motivo}`;
    } else {
      icoIA.className = "fa-solid fa-circle-exclamation";
      txtIA.textContent = `Dúvida / Não lido: ${ia.motivo}`;
    }

    // Mostrar caixa de diagnóstico de causa quando houver alerta/erro (amarelo ou vermelho)
    if (cor !== "verde" && ia.diagnostico_causa && boxDiagCausa) {
      boxDiagCausa.style.display = "block";
      lblDiagTitulo.textContent = `O que pode ter acontecido: ${ia.diagnostico_causa}`;
      lblDiagExp.textContent = ia.explicacao_causa || "";
    }

    // Renderizar botões de sugestões inteligentes de 1 clique
    if (boxSugestoes) {
      const botoes = [];
      const sugs = ia.sugestoes_acao || [];
      sugs.forEach((s, idxSug) => {
        if (cor === "verde" && s.tipo === "confirmar_atual") return;
        const icone = s.icone || "fa-wand-magic-sparkles";
        const estilo = s.estilo || "";
        botoes.push(`
          <button class="btn-smart-action ${estilo}" onclick="executarSugestaoInteligente(${idxSug})">
            <i class="fa-solid ${icone}"></i> ${s.rotulo}
          </button>
        `);
      });

      if (ia.status === "tcc_grupo") {
        botoes.push(`
          <button class="btn-smart-action warning" onclick="duplicarParaGrupoTCC()">
            <i class="fa-solid fa-users"></i> Salvar cópias p/ todos do grupo
          </button>
        `);
      }
      boxSugestoes.innerHTML = botoes.join("");
    }
  } else {
    bannerIA.style.display = "none";
  }

  renderizarListaNomes();
  renderizarTabelaDados();
  renderizarPainelMemoriaIA(st.memoria_ia);
  if (st.estatisticas) {
    renderizarPainelEstatisticas(st.estatisticas);
  }
}

/* ========================================================================= */
/* RENDERIZAÇÃO DA LISTA DE PROCESSOS COM SEMÁFORO OCR E BOTÕES CLAROS       */
/* ========================================================================= */
function obterBadgeOCR(it) {
  if (it.concluido) {
    return `<span class="ocr-pill verde"><i class="fa-solid fa-check"></i> Concluído</span>`;
  }
  if (!it.ocr) {
    return "";
  }
  const cor = it.ocr.nivel_cor || "vermelho";
  if (cor === "verde") {
    return `<span class="ocr-pill verde" title="${it.ocr.motivo || ''}"><i class="fa-solid fa-circle-check"></i> OCR OK</span>`;
  }
  if (cor === "amarelo") {
    return `<span class="ocr-pill amarelo" title="${it.ocr.motivo || ''}"><i class="fa-solid fa-triangle-exclamation"></i> Revisar</span>`;
  }
  return `<span class="ocr-pill vermelho" title="${it.ocr.motivo || ''}"><i class="fa-solid fa-circle-xmark"></i> Dúvida / Não lido</span>`;
}

function renderizarListaNomes() {
  if (!ESTADO) return;
  const container = document.getElementById("lista-nomes");
  const filtro = removerAcentosJS(document.getElementById("inp-busca").value || "").trim().toLowerCase();

  const itens = ESTADO.fila_nomes.filter(it => {
    if (!filtro) return true;
    const alvo = removerAcentosJS(
      `${it.nome_completo} ${it.matricula} ${it.cpf} ${it.nome_formatado} ${it.pdf_previsto}`
    ).toLowerCase();
    return alvo.includes(filtro);
  });

  container.innerHTML = itens.map(it => {
    let classes = ["queue-row"];
    if (it.atual) classes.push("current");
    if (it.concluido) classes.push("done");
    if (!it.concluido && it.ocr && it.ocr.nivel_cor) {
      classes.push(`ocr-${it.ocr.nivel_cor}`);
    }

    const origemTxt = it.concluido
      ? (it.ocr && it.ocr.pdf_nome ? it.ocr.pdf_nome : "Salvo")
      : (it.pdf_previsto ? it.pdf_previsto : "Sem PDF na fila");

    const notaOcr = (!it.concluido && it.ocr && it.ocr.motivo)
      ? `<div class="row-ocr-note"><i class="fa-solid fa-microchip"></i> ${it.ocr.motivo}</div>`
      : "";

    return `
      <div class="${classes.join(" ")}">
        <div class="row-top">
          <div class="row-title">${it.indice_original}. ${it.nome_completo}</div>
          ${obterBadgeOCR(it)}
        </div>
        <div class="row-files">
          <span class="file-flow">
            ${origemTxt} <i class="fa-solid fa-arrow-right" style="font-size:9px; margin:0 3px;"></i>
            <strong>${it.nome_formatado}.pdf</strong>
          </span>
        </div>
        ${notaOcr}
        <div class="row-actions">
          <button class="btn-row primary" onclick="confirmarComIndice(${it.idx})" title="Aplica este aluno imediatamente ao PDF que está aberto na tela">
            <i class="fa-solid fa-check"></i> Renomear c/ este
          </button>
          ${!it.concluido && !it.atual ? `
            <button class="btn-row" onclick="apontarFila(${it.idx})" title="Posiciona a fila neste aluno sem renomear ainda">
              <i class="fa-solid fa-crosshairs"></i> Selecionar na fila
            </button>
          ` : ""}
        </div>
      </div>
    `;
  }).join("");
}

function renderizarTabelaDados() {
  if (!ESTADO) return;
  const tbody = document.getElementById("tbody-dados");
  if (!tbody) return;

  tbody.innerHTML = ESTADO.fila_nomes.map(it => {
    const docId = it.matricula
      ? `Mat: ${it.matricula}`
      : (it.cpf ? `CPF: ${it.cpf}` : "Não consta");
    const prioTxt = it.prioridade === 1
      ? "1 · Matrícula"
      : (it.prioridade === 2 ? "2 · CPF" : "3 · Nome Completo");
    const arqOrigem = it.concluido
      ? (it.arquivo_gerado || `${it.nome_formatado}.pdf`)
      : `${it.pdf_previsto || "—"} → ${it.nome_formatado}.pdf`;

    return `
      <tr>
        <td>${it.indice_original}</td>
        <td><strong>${it.nome_completo}</strong></td>
        <td><code>${docId}</code></td>
        <td>${prioTxt}</td>
        <td><code>${arqOrigem}</code></td>
        <td>${obterBadgeOCR(it) || '<span style="color:var(--text-muted);">Pendente</span>'}</td>
        <td>
          <button class="btn-row" onclick="apontarFila(${it.idx}); alternarAba('operacao');">
            <i class="fa-solid fa-arrow-up-right-from-square"></i> Ir p/ Aluno
          </button>
          <button class="btn-row" onclick="removerPessoaFila(${it.idx})" title="Remover este aluno da fila">
            <i class="fa-solid fa-trash-can"></i>
          </button>
        </td>
      </tr>
    `;
  }).join("");
}

/* ========================================================================= */
/* AÇÕES DE INTERFACE                                                        */
/* ========================================================================= */
function fecharGavetas() {
  for (const id of ["gaveta-editar", "gaveta-dividir", "gaveta-juntar"]) {
    const el = document.getElementById(id);
    if (el) el.classList.remove("open");
  }
}

function alternarGaveta(id) {
  const el = document.getElementById(id);
  const estavaAberta = el.classList.contains("open");
  fecharGavetas();
  if (!estavaAberta) {
    el.classList.add("open");
    if (id === "gaveta-editar") {
      document.getElementById("inp-edit-nome").value = "";
      document.getElementById("inp-edit-mat").value = "";
      document.getElementById("inp-edit-cpf").value = "";
      atualizarPreviewEdicao();
      setTimeout(() => document.getElementById("inp-edit-nome").focus(), 20);
    } else if (id === "gaveta-dividir" && ESTADO && ESTADO.pdf_atual) {
      const pags = ESTADO.pdf_atual.paginas;
      document.getElementById("lbl-info-divisao").innerHTML =
        `<i class="fa-solid fa-scissors"></i> Dividir PDF (${pags || 1} páginas)`;
    } else if (id === "gaveta-juntar" && ESTADO && ESTADO.pdf_atual) {
      const prox = ESTADO.pdf_atual.proximo_nome;
      document.getElementById("lbl-info-juntar").textContent = prox
        ? `Próximo da fila: ${prox} (Total restantes: ${ESTADO.pdf_atual.restantes})`
        : "Não há outro PDF após o atual na fila.";
    }
  }
}

async function confirmarAtual() {
  fecharGavetas();
  await apiPost("/api/confirm", {});
}

async function confirmarJuntandoAoExistente() {
  fecharGavetas();
  await apiPost("/api/confirm", { juntar_se_existir: true });
}

async function confirmarComIndice(idx) {
  fecharGavetas();
  await apiPost("/api/confirm", { idx_escolhido: idx });
}

async function usarSugestaoIA() {
  if (!ESTADO || !ESTADO.verificacao_ia || ESTADO.verificacao_ia.idx_sugerido == null) return;
  await confirmarComIndice(ESTADO.verificacao_ia.idx_sugerido);
}

async function executarSugestaoInteligente(idxSugestao) {
  if (!ESTADO || !ESTADO.verificacao_ia || !ESTADO.verificacao_ia.sugestoes_acao) return;
  const sug = ESTADO.verificacao_ia.sugestoes_acao[idxSugestao];
  if (!sug) return;

  if (sug.tipo === "confirmar_atual") {
    await confirmarAtual();
  } else if (sug.tipo === "ensinar_ia_atual") {
    await apiPost("/api/train_ai", { modo: "ensinar_atual", confirmar: true });
  } else if (sug.tipo === "usar_sugerido_e_confirmar") {
    await apiPost("/api/confirm", { idx_escolhido: sug.idx_nome, ensinar_ia: true });
  } else if (sug.tipo === "selecionar_nome") {
    await apontarFila(sug.idx_nome);
  } else if (sug.tipo === "anexar_anterior") {
    await anexarAoAnterior();
  } else if (sug.tipo === "dividir_pag1") {
    await executarDivisao("corte", 1, true);
  } else if (sug.tipo === "duplicar_pdf") {
    await executarDivisao("duplicar", 1, true);
  } else if (sug.tipo === "auto_orientar") {
    await autoOrientarPdf();
  } else if (sug.tipo === "abrir_modal_manual") {
    alternarGaveta("gaveta-editar");
  }
}

async function ensinarIAApenasPadrao() {
  await apiPost("/api/train_ai", { modo: "ensinar_atual", confirmar: false });
}

async function ignorarTermoRapido() {
  const termo = window.prompt(
    "Digite um nome ou termo recorrente para a IA ignorar nos PDFs (ex: nome do coordenador, diretor ou familiar):"
  );
  if (!termo || !termo.trim()) return;
  await apiPost("/api/train_ai", { modo: "ignorar_termo", termo: termo.trim() });
}

async function adicionarTermoIgnoradoPainel() {
  const inp = document.getElementById("inp-ignorar-termo");
  if (!inp || !inp.value.trim()) return;
  const termo = inp.value.trim();
  const res = await apiPost("/api/train_ai", { modo: "ignorar_termo", termo: termo });
  if (res && res.ok) {
    inp.value = "";
  }
}

async function removerTermoIgnorado(termo) {
  await apiPost("/api/train_ai", { modo: "remover_ignorado", termo: termo });
}

function renderizarPainelMemoriaIA(mem) {
  if (!mem) return;
  const elConf = document.getElementById("lbl-mem-confirmacoes");
  const elVar = document.getElementById("lbl-mem-variacoes");
  const elPad = document.getElementById("lbl-mem-padroes");
  const boxIgn = document.getElementById("box-termos-ignorados");
  const boxHist = document.getElementById("box-ultimos-aprendizados");

  if (elConf) elConf.textContent = mem.total_confirmacoes || 0;
  if (elVar) elVar.textContent = mem.total_variacoes_ocr || 0;
  if (elPad) elPad.textContent = mem.total_padroes_alunos || 0;

  if (boxIgn) {
    const termos = mem.termos_ignorados || [];
    if (!termos.length) {
      boxIgn.innerHTML = `<span style="font-size:10.5px; color:var(--text-muted);">Nenhum nome ignorado cadastrado.</span>`;
    } else {
      boxIgn.innerHTML = termos.map(t => `
        <span class="ai-tag-pill">
          <i class="fa-solid fa-eye-slash"></i> ${t}
          <button onclick="removerTermoIgnorado('${t.replace(/'/g, "\\'")}')" title="Remover da lista de ignorados">
            <i class="fa-solid fa-xmark"></i>
          </button>
        </span>
      `).join("");
    }
  }

  if (boxHist) {
    const hist = mem.ultimos_aprendizados || [];
    if (!hist.length) {
      boxHist.innerHTML = `<div class="ai-history-item">Confirme PDFs para a IA aprender correções de OCR automaticamente.</div>`;
    } else {
      boxHist.innerHTML = hist.slice(0, 6).map(h => `
        <div class="ai-history-item"><i class="fa-solid fa-graduation-cap"></i> ${h}</div>
      `).join("");
    }
  }
}

async function duplicarParaGrupoTCC() {
  if (!ESTADO || !ESTADO.verificacao_ia || !ESTADO.verificacao_ia.candidatos_grupo) return;
  const indices = ESTADO.verificacao_ia.candidatos_grupo.map(c => c.idx);
  if (!indices.length) return;
  await apiPost("/api/duplicate_group", { indices: indices });
}

async function diagnosticarFilaOCR() {
  mostrarStatus("Analisando PDFs da fila com OCR local...");
  await apiPost("/api/scan_ocr_queue", { limite: 25 });
}

async function apontarFila(idx) {
  await apiPost("/api/select_name", { idx_nome: idx });
}

async function salvarEdicaoManual() {
  const nome = document.getElementById("inp-edit-nome").value.trim();
  const mat = document.getElementById("inp-edit-mat").value.trim();
  const cpf = document.getElementById("inp-edit-cpf").value.trim();
  if (!nome) return;
  const modo = document.querySelector('input[name="modo-edit"]:checked').value;
  fecharGavetas();
  await apiPost("/api/edit", {
    nome: nome,
    matricula: mat,
    cpf: cpf,
    avancar_planilha: modo === "substituir"
  });
}

async function executarDivisao(modo, paginaCorte, jaRenomearParte1) {
  fecharGavetas();
  await apiPost("/api/split", {
    modo: modo,
    pagina_corte: paginaCorte,
    ja_renomear_parte1: jaRenomearParte1
  });
}

async function executarDivisaoCustom() {
  const corte = parseInt(document.getElementById("inp-corte-pag").value || "1", 10);
  await executarDivisao("corte", corte, true);
}

async function executarJuncao(qtd) {
  fecharGavetas();
  await apiPost("/api/merge", { quantidade: qtd });
}

async function executarJuncaoCustom() {
  const qtd = parseInt(document.getElementById("inp-qtd-juntar").value || "2", 10);
  await executarJuncao(qtd);
}

async function anexarAoAnterior() {
  fecharGavetas();
  await apiPost("/api/append_previous", {});
}

async function girarPdf(angulo, pagina) {
  await apiPost("/api/rotate", { angulo: angulo, pagina: pagina });
}

async function autoOrientarPdf() {
  mostrarStatus("Analisando orientação das páginas...");
  await apiPost("/api/auto_orient", {});
}

async function autoLimparBordas(pagina) {
  mostrarStatus("Removendo bordas pretas de scanner...");
  const res = await apiPost("/api/auto_crop_borders", { pagina: pagina });
  if (res && res.ok && document.getElementById("crop-overlay").classList.contains("open")) {
    carregarImagemCrop();
  }
}

async function removerPaginasBrancas() {
  if (!ESTADO || !ESTADO.pdf_atual || !ESTADO.pdf_atual.paginas_brancas) return;
  await apiPost("/api/delete_pages", { paginas: ESTADO.pdf_atual.paginas_brancas });
}

async function inverterOrdemPaginas() {
  const res = await apiPost("/api/reorder_pages", { inverter: true });
  if (res && res.ok && document.getElementById("crop-overlay").classList.contains("open")) {
    carregarImagemCrop();
  }
}

async function executarAutoLote() {
  mostrarStatus("Verificando e auto-renomeando PDFs confirmados...");
  await apiPost("/api/auto_batch_ocr", {});
}

async function pularNome() {
  fecharGavetas();
  await apiPost("/api/skip_name", {});
}

async function pularPdf() {
  fecharGavetas();
  await apiPost("/api/skip_pdf", { avancar_nome_junto: false });
}

async function desfazerUltimo() {
  fecharGavetas();
  await apiPost("/api/undo", {});
}

async function reescanearPasta() {
  await apiPost("/api/rescan", {});
}

async function recarregarPlanilhaDisco() {
  await apiPost("/api/reload_sheet", {});
}

async function enviarPlanilhaColada() {
  const texto = document.getElementById("txt-colar-excel").value;
  const res = await apiPost("/api/paste_sheet", { texto: texto });
  if (res && res.ok) {
    document.getElementById("txt-colar-excel").value = "";
    alternarAba("operacao");
  }
}

function enviarArquivoPlanilha(inputEl) {
  const file = inputEl.files && inputEl.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = async () => {
    const base64 = String(reader.result).split(",")[1] || "";
    const res = await apiPost("/api/upload_sheet", {
      filename: file.name,
      content_b64: base64
    });
    if (res && res.ok) {
      inputEl.value = "";
      alternarAba("operacao");
    }
  };
  reader.readAsDataURL(file);
}

async function adicionarPessoaManual() {
  const nome = document.getElementById("inp-add-nome").value.trim();
  const mat = document.getElementById("inp-add-mat").value.trim();
  const cpf = document.getElementById("inp-add-cpf").value.trim();
  if (!nome) return;
  const res = await apiPost("/api/add_person", {
    nome: nome,
    matricula: mat,
    cpf: cpf
  });
  if (res && res.ok) {
    document.getElementById("inp-add-nome").value = "";
    document.getElementById("inp-add-mat").value = "";
    document.getElementById("inp-add-cpf").value = "";
  }
}

async function removerPessoaFila(idx) {
  await apiPost("/api/delete_person", { idx: idx });
}

/* ========================================================================= */
/* EDITOR VISUAL DE RECORTE DE PÁGINA                                        */
/* ========================================================================= */
function alternarModoRecorte() {
  const ov = document.getElementById("crop-overlay");
  if (ov.classList.contains("open")) {
    fecharModoRecorte();
  } else {
    if (!ESTADO || !ESTADO.tem_pdf) return;
    paginaCropAtual = 1;
    ov.classList.add("open");
    document.getElementById("btn-abrir-crop").classList.add("active");
    carregarImagemCrop();
  }
}

function fecharModoRecorte() {
  document.getElementById("crop-overlay").classList.remove("open");
  document.getElementById("btn-abrir-crop").classList.remove("active");
  document.getElementById("crop-rect").style.display = "none";
  cropSelecao = null;
}

function mudarPaginaCrop(delta) {
  if (!ESTADO || !ESTADO.pdf_atual) return;
  const total = ESTADO.pdf_atual.paginas || 1;
  paginaCropAtual = Math.max(1, Math.min(total, paginaCropAtual + delta));
  carregarImagemCrop();
}

function carregarImagemCrop() {
  if (!ESTADO || !ESTADO.pdf_atual) return;
  const total = ESTADO.pdf_atual.paginas || 1;
  document.getElementById("lbl-crop-pag").textContent = `Página ${paginaCropAtual} / ${total}`;
  const img = document.getElementById("crop-img");
  img.src = `/api/page_image?name=${encodeURIComponent(ESTADO.pdf_atual.nome)}&page=${paginaCropAtual}&t=${Date.now()}`;
  document.getElementById("crop-rect").style.display = "none";
  cropSelecao = null;
}

async function confirmarRecorteManual() {
  if (!cropSelecao) {
    mostrarStatus("Arraste o mouse sobre a página para selecionar a área de recorte.", true);
    return;
  }
  const todas = document.getElementById("chk-crop-todas").checked;
  const res = await apiPost("/api/crop_page", {
    pagina: todas ? 0 : paginaCropAtual,
    x0: cropSelecao.x0,
    y0: cropSelecao.y0,
    x1: cropSelecao.x1,
    y1: cropSelecao.y1
  });
  if (res && res.ok) {
    fecharModoRecorte();
  }
}

async function excluirPaginaAtualCrop() {
  if (!ESTADO || !ESTADO.pdf_atual) return;
  if ((ESTADO.pdf_atual.paginas || 1) <= 1) {
    mostrarStatus("O PDF possui apenas 1 página.", true);
    return;
  }
  const res = await apiPost("/api/delete_pages", { paginas: [paginaCropAtual] });
  if (res && res.ok) {
    paginaCropAtual = 1;
    carregarImagemCrop();
  }
}

const stageEl = document.getElementById("crop-stage");
const rectEl = document.getElementById("crop-rect");

stageEl.addEventListener("mousedown", (e) => {
  const b = stageEl.getBoundingClientRect();
  arrastandoCrop = true;
  cropInicio = {
    x: Math.max(0, Math.min(b.width, e.clientX - b.left)),
    y: Math.max(0, Math.min(b.height, e.clientY - b.top))
  };
  rectEl.style.left = cropInicio.x + "px";
  rectEl.style.top = cropInicio.y + "px";
  rectEl.style.width = "0px";
  rectEl.style.height = "0px";
  rectEl.style.display = "block";
});

window.addEventListener("mousemove", (e) => {
  if (!arrastandoCrop) return;
  const b = stageEl.getBoundingClientRect();
  if (b.width <= 0 || b.height <= 0) return;
  const curX = Math.max(0, Math.min(b.width, e.clientX - b.left));
  const curY = Math.max(0, Math.min(b.height, e.clientY - b.top));
  const x0 = Math.min(cropInicio.x, curX);
  const y0 = Math.min(cropInicio.y, curY);
  const w = Math.abs(curX - cropInicio.x);
  const h = Math.abs(curY - cropInicio.y);
  rectEl.style.left = x0 + "px";
  rectEl.style.top = y0 + "px";
  rectEl.style.width = w + "px";
  rectEl.style.height = h + "px";
  cropSelecao = {
    x0: x0 / b.width,
    y0: y0 / b.height,
    x1: (x0 + w) / b.width,
    y1: (y0 + h) / b.height
  };
});

window.addEventListener("mouseup", () => {
  arrastandoCrop = false;
});

setInterval(() => {
  if (document.activeElement && document.activeElement.id === "pdf-frame") {
    document.activeElement.blur();
    window.focus();
  }
}, 250);

window.addEventListener("keydown", (e) => {
  const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : "";
  if (tag === "input" || tag === "textarea" || tag === "select") {
    return;
  }

  if (e.key === "Escape") {
    fecharGavetas();
    fecharModoRecorte();
    fecharModalAtalhos();
    return;
  }

  if (e.key === "?") {
    e.preventDefault();
    alternarModalAtalhos();
    return;
  }

  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
    e.preventDefault();
    desfazerUltimo();
    return;
  }

  if (e.ctrlKey || e.altKey || e.metaKey) return;

  const k = e.key.toLowerCase();
  if (k === "s" || e.key === "Enter") {
    e.preventDefault();
    confirmarAtual();
  } else if (k === "e") {
    e.preventDefault();
    alternarGaveta("gaveta-editar");
  } else if (k === "d") {
    e.preventDefault();
    alternarGaveta("gaveta-dividir");
  } else if (k === "j") {
    e.preventDefault();
    alternarGaveta("gaveta-juntar");
  } else if (k === "a") {
    e.preventDefault();
    anexarAoAnterior();
  } else if (k === "r") {
    e.preventDefault();
    girarPdf(90, 0);
  } else if (k === "c") {
    e.preventDefault();
    alternarModoRecorte();
  } else if (k === "i") {
    e.preventDefault();
    executarAutoLote();
  } else if (k === "b" || k === "/") {
    e.preventDefault();
    const inp = document.getElementById("inp-busca");
    inp.focus();
    inp.select();
  } else if (k === "n") {
    e.preventDefault();
    pularNome();
  } else if (k === "p") {
    e.preventDefault();
    pularPdf();
  } else if (k === "z") {
    e.preventDefault();
    desfazerUltimo();
  }
});

/* ========================================================================= */
/* RENDERIZAÇÃO DO PAINEL DE ESTATÍSTICAS E MÉTRICAS (100% OFFLINE)          */
/* ========================================================================= */
function renderizarPainelEstatisticas(est) {
  if (!est) return;

  const res = est.resumo || {};
  const prio = est.prioridades || {};
  const sem = est.semaforo_ocr || {};
  const ops = est.operacoes_sessao || {};
  const log = est.historico_log || {};

  // 1. Atualizar KPIs
  const pct = res.total_planilha > 0 ? ((res.concluidos_planilha / res.total_planilha) * 100).toFixed(1) : "0.0";
  document.getElementById("kpi-concluidos").textContent = res.concluidos_planilha || 0;
  document.getElementById("kpi-concluidos-pct").textContent = `/ ${res.total_planilha || 0} (${pct}%)`;
  document.getElementById("kpi-pendentes-txt").textContent = `${res.pendentes_planilha || 0} pendentes na planilha`;

  document.getElementById("kpi-ritmo").textContent = (res.ritmo_por_minuto || 0).toFixed(1);
  document.getElementById("kpi-seg-pdf").textContent = `${res.segundos_por_pdf || 0}s por PDF · Tempo: ${res.tempo_sessao_min || 0} min`;

  document.getElementById("kpi-eta").textContent = res.eta_minutos > 0 ? res.eta_minutos : "—";
  document.getElementById("kpi-fila-restante").textContent = `${res.pdfs_restantes_fila || 0} PDFs restantes na fila`;

  document.getElementById("kpi-hoje").textContent = log.renomeados_hoje || 0;
  document.getElementById("kpi-total-historico").textContent = `Total acumulado no log: ${log.total_sucesso || 0}`;

  // 2. Barras de Prioridades (1: Matrícula, 2: CPF, 3: Nome Completo)
  const elPrio = document.getElementById("stats-prio-bars");
  if (elPrio) {
    const p1 = prio.p1 || { total: 0, concluidos: 0 };
    const p2 = prio.p2 || { total: 0, concluidos: 0 };
    const p3 = prio.p3 || { total: 0, concluidos: 0 };

    const pct1 = p1.total > 0 ? ((p1.concluidos / p1.total) * 100).toFixed(0) : 0;
    const pct2 = p2.total > 0 ? ((p2.concluidos / p2.total) * 100).toFixed(0) : 0;
    const pct3 = p3.total > 0 ? ((p3.concluidos / p3.total) * 100).toFixed(0) : 0;

    elPrio.innerHTML = `
      <div class="stat-progress-row">
        <div class="stat-progress-labels">
          <span><i class="fa-solid fa-id-badge icon-accent"></i> Prioridade 1 (Matrícula)</span>
          <strong>${p1.concluidos} / ${p1.total} (${pct1}%)</strong>
        </div>
        <div class="stat-track"><div class="stat-fill p1" style="width:${pct1}%;"></div></div>
      </div>
      <div class="stat-progress-row">
        <div class="stat-progress-labels">
          <span><i class="fa-solid fa-address-card icon-accent"></i> Prioridade 2 (CPF)</span>
          <strong>${p2.concluidos} / ${p2.total} (${pct2}%)</strong>
        </div>
        <div class="stat-track"><div class="stat-fill p2" style="width:${pct2}%;"></div></div>
      </div>
      <div class="stat-progress-row">
        <div class="stat-progress-labels">
          <span><i class="fa-solid fa-signature icon-accent"></i> Prioridade 3 (Nome Completo)</span>
          <strong>${p3.concluidos} / ${p3.total} (${pct3}%)</strong>
        </div>
        <div class="stat-track"><div class="stat-fill p3" style="width:${pct3}%;"></div></div>
      </div>
    `;
  }

  // 3. Semáforo OCR
  const elOcr = document.getElementById("stats-ocr-bars");
  if (elOcr) {
    elOcr.innerHTML = `
      <div class="ocr-stat-card verde">
        <div class="ocr-stat-val">${sem.verde || 0}</div>
        <div class="ocr-stat-lbl"><i class="fa-solid fa-circle-check"></i> OCR Confirmado / OK</div>
      </div>
      <div class="ocr-stat-card amarelo">
        <div class="ocr-stat-val">${sem.amarelo || 0}</div>
        <div class="ocr-stat-lbl"><i class="fa-solid fa-triangle-exclamation"></i> Revisão / Grupo / Familiar</div>
      </div>
      <div class="ocr-stat-card vermelho">
        <div class="ocr-stat-val">${sem.vermelho || 0}</div>
        <div class="ocr-stat-lbl"><i class="fa-solid fa-circle-xmark"></i> Dúvida / Não Lido</div>
      </div>
      <div class="ocr-stat-card muted">
        <div class="ocr-stat-val">${sem.nao_analisado || 0}</div>
        <div class="ocr-stat-lbl"><i class="fa-solid fa-clock"></i> Pendente de Verificação</div>
      </div>
    `;
  }

  // 4. Ações Aplicadas na Sessão
  const elOps = document.getElementById("stats-ops-badges");
  if (elOps) {
    elOps.innerHTML = `
      <div class="op-badge-box">
        <i class="fa-solid fa-file-signature icon-accent"></i>
        <strong>${ops.renomeados || 0}</strong>
        <span>Renomeados</span>
      </div>
      <div class="op-badge-box">
        <i class="fa-solid fa-pen-to-square" style="color:#60a5fa;"></i>
        <strong>${ops.editados_manual || 0}</strong>
        <span>Edições Manuais</span>
      </div>
      <div class="op-badge-box">
        <i class="fa-solid fa-scissors" style="color:#c084fc;"></i>
        <strong>${ops.divididos || 0}</strong>
        <span>PDFs Divididos</span>
      </div>
      <div class="op-badge-box">
        <i class="fa-solid fa-object-group" style="color:#fbbf24;"></i>
        <strong>${ops.juntados || 0}</strong>
        <span>PDFs Juntados</span>
      </div>
      <div class="op-badge-box">
        <i class="fa-solid fa-forward-step" style="color:#f87171;"></i>
        <strong>${ops.pulados || 0}</strong>
        <span>PDFs/Nomes Pulados</span>
      </div>
    `;
  }

  // 5. Gráfico de Barras por Dia
  const elBar = document.getElementById("stats-bar-chart");
  if (elBar) {
    const dias = log.por_dia || [];
    if (!dias.length) {
      elBar.innerHTML = `<div class="empty-chart-msg"><i class="fa-solid fa-clock"></i> Nenhuma renomeação anterior registrada no log ainda.</div>`;
    } else {
      const maxVal = Math.max(...dias.map(d => d.valor), 1);
      elBar.innerHTML = dias.map(d => {
        const alturaPct = Math.max(12, ((d.valor / maxVal) * 100).toFixed(0));
        return `
          <div class="bar-col">
            <span class="bar-val">${d.valor}</span>
            <div class="bar-pillar" style="height:${alturaPct}%;"></div>
            <span class="bar-lbl">${d.rotulo}</span>
          </div>
        `;
      }).join("");
    }
  }

  // 6. Tabela de Últimas Operações do Log
  const tbodyLog = document.getElementById("tbody-log-stats");
  if (tbodyLog) {
    const ultimas = log.ultimas_operacoes || [];
    if (!ultimas.length) {
      tbodyLog.innerHTML = `<tr><td colspan="5" style="text-align:center; color:var(--text-muted); padding:16px;">Nenhum registro no log.</td></tr>`;
    } else {
      tbodyLog.innerHTML = ultimas.map(u => {
        const badgeTipo = u.tipo === "SUCESSO"
          ? `<span class="ocr-pill verde"><i class="fa-solid fa-check"></i> Sucesso</span>`
          : `<span class="ocr-pill amarelo"><i class="fa-solid fa-rotate-left"></i> Desfeito</span>`;
        return `
          <tr>
            <td>${u.data}</td>
            <td><code>${u.horario}</code></td>
            <td>${badgeTipo}</td>
            <td><code>${u.origem}</code></td>
            <td><strong style="color:var(--accent);">${u.destino}</strong></td>
          </tr>
        `;
      }).join("");
    }
  }
}

window.addEventListener("DOMContentLoaded", carregarEstadoInicial);
