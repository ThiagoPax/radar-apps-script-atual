// =============================================================================
// 45_Hoje_GE.js — Cockpit operacional do programa (aba Hoje)
// =============================================================================
// A aba Hoje concentra: OCR ao vivo (webhook grava em D), audiência manual (E),
// audiência final (F), predição (C), leitura contextual e decisão editorial
// por minuto, espelho (janelas de break/merchan) e o gráfico do dia.
//
// O motor de predição (40/41/42/43) continua operando sobre a aba Histórico
// (colunas B/D/F + PROJ V2 + painel TD-TV). Este módulo sincroniza os dados
// reais da Hoje para o Histórico e traz de volta previsão e sinais.
//
// Colunas (linhas 3–62 = minutos 18:00–18:59):
//  A HORA | B MÉDIA ESPERADA | C PREVISÃO | D AUD OCR | E AUD MANUAL | F AUD FINAL
//  G ORIGEM | H STATUS | I TEMA | J SUBTEMA | K DESVIO | L LEITURA
//  M TEMA CERTO? | N POR QUÊ? | O SUSTENTAR | P PRÓXIMO TEMA | Q RISCO
//  R AÇÃO | S OBS | T ÚLT. PROG
// Linha 63 = médias. Linhas 65+ = espelho e réguas do dia.

var HOJE_GE = {
    FIRST_ROW: 3,
    LAST_ROW: 62,
    N_MIN: 60,
    HEADERS: ["HORA", "MÉDIA ESPERADA", "PREVISÃO", "AUD OCR", "AUD MANUAL",
        "AUD FINAL", "ORIGEM", "STATUS", "TEMA", "SUBTEMA", "DESVIO",
        "LEITURA", "TEMA CERTO?", "POR QUÊ?", "SUSTENTAR", "PRÓXIMO TEMA",
        "RISCO", "AÇÃO", "OBS", "ÚLT. PROG", "MÉDIA HIST.", "SOMBRA", "Δ SOMBRA"
    ],
    COL: { HORA: 1, MEDIA_CTX: 2, PREVISAO: 3, OCR: 4, MANUAL: 5, FINAL: 6,
        ORIGEM: 7, STATUS: 8, TEMA: 9, SUBTEMA: 10, DESVIO: 11, LEITURA: 12,
        TEMA_CERTO: 13, PORQUE: 14, SUSTENTAR: 15, PROX_TEMA: 16, RISCO: 17,
        ACAO: 18, OBS: 19, ULT_PROG: 20, MEDIA_HIST: 21, SOMBRA: 22, DELTA_SOMBRA: 23 },
    NCOLS: 23,
    ESPELHO_ROW: 65,       // título do bloco de espelho
    REGUAS_ROW: 76,        // título do bloco de réguas contextuais
    CHART_TITLE: "Hoje x Contexto x Previsão",
    // Validação de leitura OCR (audiência SP costuma ficar entre ~0,1 e ~2)
    OCR_MIN: 0.03,
    OCR_MAX: 5,
    OCR_SALTO_MAX: 2.5     // razão máxima aceitável entre minutos vizinhos
};

function getHojeSheetGE_() {
    return ensureAbaGE_(CFG.SHEET_HOJE);
}

// -----------------------------------------------------------------------------
// Preparação diária (rotina 06:00 e menu) — monta/rearma o cockpit
// -----------------------------------------------------------------------------
function prepararHojeGE() {
    const sh = getHojeSheetGE_();
    const hoje = new Date();
    const dataStr = Utilities.formatDate(hoje, CFG.TIMEZONE, "dd/MM/yyyy");
    const diaSemana = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SÁB"][hoje.getDay()];
    const contexto = String(getConfigValueGE_("CONTEXTO_ATIVO", "NORMAL"));

    if (sh.getMaxRows() < 110) sh.insertRowsAfter(sh.getMaxRows(), 110 - sh.getMaxRows());
    if (sh.getMaxColumns() < HOJE_GE.NCOLS) sh.insertColumnsAfter(sh.getMaxColumns(), HOJE_GE.NCOLS - sh.getMaxColumns());
    const oficial = previsaoOficialGE_();

    // Título e cabeçalhos
    sh.getRange(1, 1, 1, HOJE_GE.NCOLS).breakApart().clearContent().setBackground("#1B2631");
    sh.getRange(1, 1).setValue("HOJE — " + dataStr + " (" + diaSemana + ") — Contexto: " + contexto +
        " · Previsão oficial: " + oficial)
        .setFontWeight("bold").setFontColor("#FFFFFF").setFontSize(12);
    sh.getRange(2, 1, 1, HOJE_GE.NCOLS).setValues([HOJE_GE.HEADERS])
        .setFontWeight("bold").setBackground("#34495E").setFontColor("#FFFFFF")
        .setHorizontalAlignment("center");
    sh.getRange(2, HOJE_GE.COL.PREVISAO).setValue("PREVISÃO (" + oficial + " · OFICIAL)");
    aplicarEstiloSombraHojeGE_(sh, oficial);
    sh.setFrozenRows(2);
    sh.setFrozenColumns(1);

    // Grade de minutos
    const horas = [];
    for (let m = 0; m < 60; m++) horas.push(["18:" + String(m).padStart(2, "0")]);
    sh.getRange(3, 1, 60, 1).setValues(horas).setFontWeight("bold");

    // Limpa dados do dia anterior (mantém formatação)
    sh.getRange(3, 2, 60, 22).clearContent().clearNote();
    sh.getRange(3, 2, 60, 5).setNumberFormat("0.00");
    sh.getRange(3, 11, 60, 1).setNumberFormat("+0.0%;-0.0%");
    sh.getRange(3, 20, 60, 3).setNumberFormat("0.00"); // T,U,V
    sh.getRange(3, HOJE_GE.COL.DELTA_SOMBRA, 60, 1).setNumberFormat("+0.00;-0.00");

    // Linha 63 — médias do dia (valores calculados por script; fórmulas com
    // nome localizado quebram conforme o locale do arquivo)
    sh.getRange(63, 1).setValue("MÉDIA").setFontWeight("bold");
    sh.getRange(63, 1, 1, HOJE_GE.NCOLS).setBackground("#F5B7B1").setNumberFormat("0.00");
    atualizarMediasLinha63HojeGE_(sh);

    // Média contextual (B), último programa (T) e média histórica pura (U)
    preencherMediaContextualHojeGE_(sh);
    preencherUltimoProgramaHojeGE_(sh);
    preencherMediaHistoricaHojeGE_(sh);

    // Previsão inicial (C) — antes de existir dado real, usa a régua contextual
    const b = sh.getRange(3, 2, 60, 1).getValues();
    sh.getRange(3, 3, 60, 1).setValues(b);

    // Previsão sombra (V3) na coluna V — nunca entra no gráfico/decisões
    try { preencherSombraHojeGE_(sh); } catch (e) { logGE_("HOJE", "Sombra V3: " + e.message); }

    // Réguas do dia + espelho + gráfico
    preencherReguasHojeGE_(sh);
    calcularEspelhoHojeGE_(sh);
    garantirGraficoHojeGE_(sh);

    logGE_("HOJE", "Aba Hoje preparada para " + dataStr + " (contexto " + contexto + ")");
}

// Média histórica ajustada ao contexto ativo: usa a média por minuto do
// Histórico (col B) escalada pela razão média_contexto / média_geral da Base.
function preencherMediaContextualHojeGE_(sh) {
    try {
        const hist = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
        if (!hist) return;
        const base = hist.getRange(3, 2, 60, 1).getValues();
        const fator = fatorContextoGE_();
        const out = base.map(r => [typeof r[0] === "number" && r[0] > 0 ? r[0] * fator : ""]);
        sh.getRange(3, 2, 60, 1).setValues(out);
    } catch (e) {
        logGE_("HOJE", "Falha ao preencher média contextual: " + e.message);
    }
}

// Média histórica pura por minuto (Histórico col B, sem ajuste de contexto) —
// é a linha vermelha do gráfico.
function preencherMediaHistoricaHojeGE_(sh) {
    try {
        const hist = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
        if (!hist) return;
        const base = hist.getRange(3, 2, 60, 1).getValues();
        // Garante o cabeçalho (é o nome da série na legenda do gráfico)
        if (!sh.getRange(2, HOJE_GE.COL.MEDIA_HIST).getValue()) {
            sh.getRange(2, HOJE_GE.COL.MEDIA_HIST).setValue("MÉDIA HIST.");
        }
        sh.getRange(3, HOJE_GE.COL.MEDIA_HIST, 60, 1).setValues(
            base.map(r => [typeof r[0] === "number" && r[0] > 0 ? r[0] : ""]));
    } catch (e) {
        logGE_("HOJE", "Falha ao preencher média histórica: " + e.message);
    }
}

// Curva do último programa fechado (bloco mais recente do Histórico que não
// seja o dia de hoje).
function preencherUltimoProgramaHojeGE_(sh) {
    try {
        const hist = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
        if (!hist) return;
        const hojeStr = Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy");
        const cols = getAudColsGE_(hist);
        for (const c of cols) {
            const h = hist.getRange(1, c).getValue();
            const hStr = h instanceof Date
                ? Utilities.formatDate(h, CFG.TIMEZONE, "dd/MM/yyyy") : String(h).trim();
            if (hStr === hojeStr) continue;
            const vals = hist.getRange(3, c, 60, 1).getValues();
            sh.getRange(3, 20, 60, 1).setValues(vals);
            sh.getRange(2, 20).setNote("Último programa: " + hStr);
            return;
        }
    } catch (e) {
        logGE_("HOJE", "Falha ao preencher último programa: " + e.message);
    }
}

// Bloco de réguas contextuais (linhas REGUAS_ROW+)
function preencherReguasHojeGE_(sh) {
    const r = HOJE_GE.REGUAS_ROW;
    const m = mediasContextuaisGE_();
    const linhas = [
        ["RÉGUAS DO DIA", "", ""],
        ["Média histórica geral", m.geral, ""],
        ["Últimos 3 programas", m.ultimos3, ""],
        ["Últimos 5 programas", m.ultimos5, ""],
        ["Mesmo dia da semana", m.mesmoDia, ""],
        ["Contexto ativo (" + m.contexto + ")", m.mediaContexto, m.nContexto ? m.nContexto + " programas" : "sem amostra"],
        ["Placar sombra V2×V3", "", placarSombraTextoGE_()]
    ];
    sh.getRange(r, 1, 8, 3).clearContent();
    sh.getRange(r, 1, linhas.length, 3).setValues(linhas);
    sh.getRange(r, 1).setFontWeight("bold").setBackground("#D6EAF8");
    sh.getRange(r + 1, 2, linhas.length - 2, 1).setNumberFormat("0.00");
}

// Placar acumulado do shadow-log (aba Backtest_Shadow) — quantos dias cada
// versão venceu; detalhe completo fica na aba.
function placarSombraTextoGE_() {
    try {
        const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BACKTEST + "_Shadow");
        if (!sh || sh.getLastRow() < 2) return "sem histórico ainda";
        const col = sh.getRange(2, 7, sh.getLastRow() - 1, 1).getValues();
        let v2 = 0, v3 = 0, emp = 0;
        for (const r of col) {
            const w = String(r[0] || "");
            if (w === "V3") v3++; else if (w === "V2") v2++; else if (w === "empate") emp++;
        }
        const oficial = previsaoOficialGE_();
        return "V3 " + v3 + " × " + v2 + " V2" + (emp ? " (" + emp + " emp.)" : "") +
            " · oficial hoje: " + oficial;
    } catch (e) { return "—"; }
}

// -----------------------------------------------------------------------------
// Sincronização Hoje ⇄ Histórico
// -----------------------------------------------------------------------------
// 1) Puxa do Histórico valores reais digitados direto lá (compatibilidade) e
//    temas preenchidos pela transcrição.
// 2) Consolida F = MANUAL > OCR validado; G origem; H status.
// 3) Empurra F e I para Histórico D e F (o que alimenta o motor).
function sincronizarHojeComHistoricoGE_() {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CFG.SHEET_HOJE);
    const hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sh || !hist) return;

    const n = HOJE_GE.N_MIN;
    const grade = sh.getRange(3, 2, n, 19); // B..T
    const vals = grade.getValues();
    const notasOcr = sh.getRange(3, 4, n, 1).getNotes();

    const histD = hist.getRange(3, 4, n, 1);
    const histDVals = histD.getValues();
    const histDForm = histD.getFormulas();
    const histF = hist.getRange(3, 6, n, 1);
    const histFVals = histF.getValues();

    const novoD = [];
    const novoF = [];
    let ultimoFinal = null;

    for (let i = 0; i < n; i++) {
        const ocr = numOuNullGE_(vals[i][HOJE_GE.COL.OCR - 2]);
        let manual = numOuNullGE_(vals[i][HOJE_GE.COL.MANUAL - 2]);
        const histReal = (!histDForm[i][0] && typeof histDVals[i][0] === "number")
            ? histDVals[i][0] : null;

        // Compatibilidade: número digitado direto no Histórico entra como manual
        if (manual === null && histReal !== null && ocr === null) {
            manual = histReal;
            vals[i][HOJE_GE.COL.MANUAL - 2] = histReal;
        }

        // Validação da leitura OCR
        let ocrValido = ocr !== null && ocr >= HOJE_GE.OCR_MIN && ocr <= HOJE_GE.OCR_MAX;
        if (ocrValido && ultimoFinal !== null && ultimoFinal > 0) {
            const razao = ocr / ultimoFinal;
            if (razao > HOJE_GE.OCR_SALTO_MAX || razao < 1 / HOJE_GE.OCR_SALTO_MAX) ocrValido = false;
        }

        let final = null, origem = "", status = "";
        if (manual !== null) {
            final = manual; origem = "MANUAL";
            status = (ocr !== null && Math.abs(ocr - manual) > 0.005) ? "MANUAL (≠ OCR)" : "OK";
        } else if (ocr !== null) {
            if (ocrValido) {
                final = ocr; origem = "OCR";
                status = /Status:\s*FINAL/i.test(notasOcr[i][0] || "") ? "OCR FINAL" : "OCR PRÉVIA";
            } else {
                status = "OCR SUSPEITO (ignorado)";
            }
        }

        vals[i][HOJE_GE.COL.FINAL - 2] = final === null ? "" : final;
        vals[i][HOJE_GE.COL.ORIGEM - 2] = origem;
        vals[i][HOJE_GE.COL.STATUS - 2] = status;
        if (final !== null) ultimoFinal = final;

        // Tema: Hoje manda; se vazio, herda do Histórico (transcrição)
        const temaHoje = String(vals[i][HOJE_GE.COL.TEMA - 2] || "").trim();
        const temaHist = String(histFVals[i][0] || "").trim();
        if (!temaHoje && temaHist) vals[i][HOJE_GE.COL.TEMA - 2] = temaHist;

        novoD.push([final === null ? (histDForm[i][0] ? histDForm[i][0] : histDVals[i][0]) : final]);
        novoF.push([String(vals[i][HOJE_GE.COL.TEMA - 2] || "").trim() || temaHist]);
    }

    // GRAVAÇÃO CIRÚRGICA — nunca reescrever a grade inteira. Entre a leitura
    // (foto da grade) e a escrita passam-se segundos; gravar a foto de volta
    // apagava o que o usuário digitou (E/I) e o que o webhook OCR gravou (D)
    // nesse intervalo — o "apagamento fantasma" do ao vivo. Só as colunas
    // calculadas F/G/H são escritas em bloco; E e I célula a célula,
    // reconferindo na hora da escrita que continuam vazias.
    const fghOut = vals.map(r => [
        r[HOJE_GE.COL.FINAL - 2], r[HOJE_GE.COL.ORIGEM - 2], r[HOJE_GE.COL.STATUS - 2]
    ]);
    sh.getRange(3, HOJE_GE.COL.FINAL, n, 3).setValues(fghOut);

    const eAtual = sh.getRange(3, HOJE_GE.COL.MANUAL, n, 1).getValues();
    const iAtual = sh.getRange(3, HOJE_GE.COL.TEMA, n, 1).getValues();
    for (let i = 0; i < n; i++) {
        const eNovo = vals[i][HOJE_GE.COL.MANUAL - 2];
        if (typeof eNovo === "number" &&
            (eAtual[i][0] === "" || eAtual[i][0] === null)) {
            sh.getRange(3 + i, HOJE_GE.COL.MANUAL).setValue(eNovo);
        }
        const tNovo = String(vals[i][HOJE_GE.COL.TEMA - 2] || "").trim();
        if (tNovo && String(iAtual[i][0] || "").trim() === "") {
            sh.getRange(3 + i, HOJE_GE.COL.TEMA).setValue(tNovo);
        }
    }

    // Empurra apenas células com dado real novo (preserva fórmulas de projeção
    // nos minutos futuros e nunca apaga um real já existente no Histórico).
    for (let i = 0; i < n; i++) {
        const v = novoD[i][0];
        if (typeof v === "number" && (histDForm[i][0] || histDVals[i][0] !== v)) {
            hist.getRange(3 + i, 4).setValue(v);
        }
        const t = novoF[i][0];
        if (t && String(histFVals[i][0] || "").trim() !== t) {
            hist.getRange(3 + i, 6).setValue(t);
        }
    }
}

// -----------------------------------------------------------------------------
// Atualização completa do cockpit (tick ao vivo, onEdit e menu)
// -----------------------------------------------------------------------------
function atualizarHojeGE() {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CFG.SHEET_HOJE);
    const hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sh || !hist) return;

    sincronizarHojeComHistoricoGE_();

    // Recalcula o motor (V2 + painel) com lock — reaproveita a automação 32.
    try {
        if (typeof ge32_atualizarTYPainelAgora_ === "function") {
            ge32_atualizarTYPainelAgora_("hoje_sync");
        }
    } catch (e) {
        logGE_("HOJE", "Motor de predição falhou: " + e.message);
    }

    // Previsão OFICIAL (V2=PROJ V2, V3=PROJ V3) → coluna C
    const oficial = previsaoOficialGE_();
    try {
        const colOf = acharColunaPorHeaderGE_(hist, oficial === "V3" ? "PROJ V3" : "PROJ V2");
        if (colOf) {
            const of = hist.getRange(3, colOf, HOJE_GE.N_MIN, 1).getValues();
            sh.getRange(3, HOJE_GE.COL.PREVISAO, HOJE_GE.N_MIN, 1).setValues(
                of.map(r => [typeof r[0] === "number" ? r[0] : ""]));
        }
        sh.getRange(2, HOJE_GE.COL.PREVISAO).setValue("PREVISÃO (" + oficial + " · OFICIAL)");
    } catch (e) { /* mantém previsão anterior */ }

    atualizarDecisaoEditorialHojeGE_(sh, hist);
    calcularEspelhoHojeGE_(sh);
    atualizarTituloHojeGE_(sh, hist);
    preencherMediaHistoricaHojeGE_(sh);
    try { preencherSombraHojeGE_(sh); } catch (e) { logGE_("HOJE", "Sombra V3: " + e.message); }
    atualizarMediasLinha63HojeGE_(sh);
    garantirGraficoSeNecessarioHojeGE_(sh);
}

// Médias da linha 63 (B, C, F, T) calculadas por script
function atualizarMediasLinha63HojeGE_(sh) {
    const vals = sh.getRange(3, 1, HOJE_GE.N_MIN, HOJE_GE.NCOLS).getValues();
    const mediaCol = c => {
        const vs = vals.map(r => r[c - 1]).filter(v => typeof v === "number" && isFinite(v));
        return vs.length ? mediaArrGE_(vs) : "";
    };
    sh.getRange(63, 2).setValue(mediaCol(2));
    sh.getRange(63, 3).setValue(mediaCol(3));
    sh.getRange(63, 6).setValue(mediaCol(6));
    sh.getRange(63, 20).setValue(mediaCol(20));
    sh.getRange(63, 21).setValue(mediaCol(21));
    sh.getRange(63, HOJE_GE.COL.SOMBRA).setValue(mediaCol(HOJE_GE.COL.SOMBRA)); // média sombra
}

// -----------------------------------------------------------------------------
// SHADOW MODE — previsão experimental (V3) sem interferir na operação
// -----------------------------------------------------------------------------
// Flag de previsão oficial (aba Config). Default V2 até a V3 provar valor.
function previsaoOficialGE_() {
    const v = String(getConfigValueGE_("PREVISAO_PRINCIPAL", "V2")).trim().toUpperCase();
    return v === "V3" ? "V3" : "V2";
}

// Estilo das colunas de comparação (fundo cinza, itálico, agrupadas/colapsadas)
function aplicarEstiloSombraHojeGE_(sh, oficial) {
    const cSombra = HOJE_GE.COL.SOMBRA, cDelta = HOJE_GE.COL.DELTA_SOMBRA;
    const sombraVer = oficial === "V3" ? "V2" : "V3"; // a sombra é a NÃO-oficial
    sh.getRange(2, cSombra).setValue("SOMBRA " + sombraVer)
        .setNote("Previsão experimental (" + sombraVer + ") — não usar ao vivo.");
    sh.getRange(2, cDelta).setValue("Δ vs oficial");
    sh.getRange(2, cSombra, 1, 2).setFontStyle("italic").setBackground("#7F8C8D");
    sh.getRange(1, cSombra, 62, 2).setBackground("#EAECEE");
    sh.getRange(1, cSombra).setBackground("#1B2631");
    sh.getRange(1, cDelta).setBackground("#1B2631");
    // Agrupa e colapsa as colunas de comparação (fora do foco operacional)
    try {
        const g = sh.getColumnGroup(cSombra, 1);
        if (!g) { sh.getRange(1, cSombra, 1, 2).shiftColumnGroupDepth(1); }
        sh.collapseAllColumnGroups && sh.collapseAllColumnGroups();
    } catch (e) { /* agrupamento é cosmético */ }
}

// Preenche a coluna SOMBRA com a previsão NÃO-oficial e Δ vs coluna C (oficial)
function preencherSombraHojeGE_(sh) {
    const oficial = previsaoOficialGE_();
    const hist = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!hist) return;

    // Garante que a V3 esteja calculada no Histórico (shadow), sem tocar o V2
    let colV3 = acharColunaPorHeaderGE_(hist, "PROJ V3");
    if (typeof recalcularPrevisaoV3ShadowGE_ === "function") {
        try { recalcularPrevisaoV3ShadowGE_(); colV3 = acharColunaPorHeaderGE_(hist, "PROJ V3"); }
        catch (e) { logGE_("HOJE", "V3 shadow: " + e.message); }
    }
    // A sombra mostra sempre a NÃO-oficial
    const headerSombra = oficial === "V3" ? "PROJ V2" : "PROJ V3";
    const colSombra = acharColunaPorHeaderGE_(hist, headerSombra);
    const n = HOJE_GE.N_MIN;
    const sombra = colSombra ? hist.getRange(3, colSombra, n, 1).getValues() : [];
    const oficialCol = sh.getRange(3, HOJE_GE.COL.PREVISAO, n, 1).getValues();

    const outS = [], outD = [];
    for (let i = 0; i < n; i++) {
        const vs = sombra.length ? sombra[i][0] : "";
        const vo = oficialCol[i][0];
        outS.push([typeof vs === "number" ? vs : ""]);
        outD.push([(typeof vs === "number" && typeof vo === "number") ? vs - vo : ""]);
    }
    sh.getRange(3, HOJE_GE.COL.SOMBRA, n, 1).setValues(outS);
    sh.getRange(3, HOJE_GE.COL.DELTA_SOMBRA, n, 1).setValues(outD);
    aplicarEstiloSombraHojeGE_(sh, oficial);
}

function atualizarTituloHojeGE_(sh, hist) {
    try {
        // "FINAL PREV." vem do painel de decisão, que mora na aba Painel.
        const shPainel = SpreadsheetApp.getActive().getSheetByName(
            typeof PAINEL31_SHEET_NAME !== "undefined" ? PAINEL31_SHEET_NAME : "Painel") || hist;
        const finalPrev = lerColunaPainelGE_(shPainel, "FINAL PREV.", 1);
        let prev = "";
        if (finalPrev.length && typeof finalPrev[0] === "number") {
            prev = " — Média final prevista: " + finalPrev[0].toFixed(2).replace(".", ",");
        }
        const dataStr = Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy");
        const contexto = String(getConfigValueGE_("CONTEXTO_ATIVO", "NORMAL"));
        sh.getRange(1, 1).setValue("HOJE — " + dataStr + " — Contexto: " + contexto + prev);
    } catch (e) { /* título é cosmético */ }
}

// -----------------------------------------------------------------------------
// Decisão editorial por minuto (K–R) a partir do painel TD-TV + contexto
// -----------------------------------------------------------------------------
function atualizarDecisaoEditorialHojeGE_(sh, hist) {
    const n = HOJE_GE.N_MIN;
    const limiarRuptura = Number(getConfigValueGE_("LIMIAR_RUPTURA", 0.20)) || 0.20;

    // O painel de decisão mora na aba "Painel" (era TD:TV no Histórico).
    const shPainel = SpreadsheetApp.getActive().getSheetByName(
        typeof PAINEL31_SHEET_NAME !== "undefined" ? PAINEL31_SHEET_NAME : "Painel") || hist;
    const painel = {
        sinal: lerColunaPainelGE_(shPainel, "SINAL", n),
        motivo: lerColunaPainelGE_(shPainel, "MOTIVO", n),
        sugestao: lerColunaPainelGE_(shPainel, "SUGESTAO", n),
        risco: lerColunaPainelGE_(shPainel, "RISCO", n),
        fatorTema: lerColunaPainelGE_(shPainel, "FATOR TEMA", n),
        durBloco: lerColunaPainelGE_(shPainel, "DUR. BLOCO", n),
        ruptura: lerColunaPainelGE_(shPainel, "RUPTURA", n)
    };

    const grade = sh.getRange(3, 2, n, 17); // B..R
    const vals = grade.getValues();
    const breaksGlobo = getPrevisaoBreaksHojeGE_();
    let ultimoRealIdx = -1;
    for (let i = 0; i < n; i++) {
        if (typeof vals[i][HOJE_GE.COL.FINAL - 2] === "number") ultimoRealIdx = i;
    }
    // Sugestão para a próxima vaga (faixa do minuto logo após o último real)
    const proximoTema = proximoTemaSugeridoGE_(sh, Math.min(59, ultimoRealIdx + 2));

    for (let i = 0; i < n; i++) {
        const b = numOuNullGE_(vals[i][HOJE_GE.COL.MEDIA_CTX - 2]);
        const final = numOuNullGE_(vals[i][HOJE_GE.COL.FINAL - 2]);
        const prev = numOuNullGE_(vals[i][HOJE_GE.COL.PREVISAO - 2]);
        const ref = final !== null ? final : prev;

        // K desvio e L leitura contextual
        let desvio = "";
        let leitura = "";
        if (b && ref !== null) {
            desvio = ref / b - 1;
            if (desvio >= limiarRuptura) leitura = "RUPTURA POSITIVA";
            else if (desvio <= -limiarRuptura) leitura = "RUPTURA NEGATIVA";
            else if (desvio > 0.07) leitura = "ACIMA DO CONTEXTO";
            else if (desvio < -0.07) leitura = "ABAIXO DO CONTEXTO";
            else leitura = "DENTRO DO CONTEXTO";
        }
        vals[i][HOJE_GE.COL.DESVIO - 2] = desvio;
        vals[i][HOJE_GE.COL.LEITURA - 2] = final !== null ? leitura : (leitura ? leitura + " (prev.)" : "");

        // Decisão só faz sentido perto do minuto corrente
        const decidir = final !== null || (ultimoRealIdx >= 0 && i <= ultimoRealIdx + 5);
        if (!decidir) {
            vals[i][HOJE_GE.COL.TEMA_CERTO - 2] = "";
            vals[i][HOJE_GE.COL.PORQUE - 2] = "";
            vals[i][HOJE_GE.COL.SUSTENTAR - 2] = "";
            vals[i][HOJE_GE.COL.PROX_TEMA - 2] = "";
            vals[i][HOJE_GE.COL.RISCO - 2] = "";
            vals[i][HOJE_GE.COL.ACAO - 2] = "";
            continue;
        }

        const ft = String(painel.fatorTema[i] || "").toUpperCase();
        vals[i][HOJE_GE.COL.TEMA_CERTO - 2] =
            ft.indexOf("AJUDA") >= 0 ? "SIM" :
            ft.indexOf("PESA") >= 0 ? "NÃO" :
            ft.indexOf("NEUTRO") >= 0 ? "NEUTRO" : "INCERTO";
        vals[i][HOJE_GE.COL.PORQUE - 2] = String(painel.motivo[i] || "");
        vals[i][HOJE_GE.COL.SUSTENTAR - 2] = sustentarTemaGE_(
            String(painel.sinal[i] || ""), String(painel.ruptura[i] || ""),
            Number(painel.durBloco[i]) || 0, i, breaksGlobo);
        vals[i][HOJE_GE.COL.PROX_TEMA - 2] = proximoTema;
        vals[i][HOJE_GE.COL.RISCO - 2] = String(painel.risco[i] || "");
        vals[i][HOJE_GE.COL.ACAO - 2] = String(painel.sugestao[i] || "");
    }
    // Escreve SÓ as colunas calculadas K..R — jamais a grade inteira, para
    // não sobrescrever D (OCR), E (manual) e I (tema) com uma foto velha.
    sh.getRange(3, HOJE_GE.COL.DESVIO, n, HOJE_GE.COL.ACAO - HOJE_GE.COL.DESVIO + 1)
        .setValues(vals.map(r => r.slice(HOJE_GE.COL.DESVIO - 2, HOJE_GE.COL.ACAO - 1)));
}

function sustentarTemaGE_(sinal, ruptura, durBloco, minutoIdx, breaksGlobo) {
    const s = sinal.toUpperCase();
    const r = ruptura.toUpperCase();

    // Break da Globo chegando: chance de capturar público com tema forte
    if (breaksGlobo && breaksGlobo.janelas) {
        for (const j of breaksGlobo.janelas) {
            if (j.inicioMin - minutoIdx > 0 && j.inicioMin - minutoIdx <= 4) {
                return "Segurar tema forte até ~18:" + String(Math.min(59, j.fimMin + 1)).padStart(2, "0") +
                    " (Globo deve estar em break)";
            }
        }
    }
    if (r.indexOf("QUEDA FORTE") >= 0) return "Encerrar agora; queda virou ruptura.";
    if (s.indexOf("PERDE FOLEGO") >= 0 || s.indexOf("ESGOT") >= 0)
        return "Trocar em até 2 min; tema perdeu força.";
    if (s.indexOf("RISCO") >= 0 || s.indexOf("PESA") >= 0)
        return "Tema pesa; preparar troca ou mudar abordagem.";
    if (s.indexOf("ALTA") >= 0 || s.indexOf("AJUDA") >= 0 || s.indexOf("RECUPERACAO") >= 0)
        return "Sustentar por mais 3 min; curva positiva.";
    if (durBloco >= 8) return "Tema longo (" + durBloco + " min); avaliar troca.";
    return "Manter e observar.";
}

// Próximo tema: índice da aba Tema Minuto + contexto dos clubes + não repetir
// temas recém-usados.
// Score multi-sinal (Fase 3): índice histórico do tema × fator clube APRENDIDO
// (48, incremental do resultado/recência) × fator faixa (rendimento no terço
// do programa) × frescor (penaliza saturação) × contexto ativo. Retorna o
// top-3 EXPLICADO. `minutoAlvo` define a faixa (abertura/miolo/fecho) da vaga.
function proximoTemaSugeridoGE_(sh, minutoAlvo) {
    try {
        const ss = SpreadsheetApp.getActive();
        const tema = ss.getSheetByName(CFG.GE_OUT_MEDIA_TEMA);
        if (!tema || tema.getLastRow() < 2) return "";
        const alvo = typeof minutoAlvo === "number" ? minutoAlvo : 30;

        const dados = tema.getRange(2, 1, tema.getLastRow() - 1, 6).getValues();
        const usadosArr = sh.getRange(3, HOJE_GE.COL.TEMA, HOJE_GE.N_MIN, 1).getValues()
            .map(r => String(r[0] || "").trim().toUpperCase());
        const recentes = usadosArr.filter(Boolean).slice(-10);
        // minutos já gastos hoje por tema (para o frescor)
        const gastosHoje = {};
        for (const t of usadosArr) if (t) gastosHoje[t] = (gastosHoje[t] || 0) + 1;

        const clubes = getContextoClubesHojeGE_();
        const fatClube = (typeof fatoresClubesHojeGE_ === "function") ? fatoresClubesHojeGE_() : {};
        const contexto = String(getConfigValueGE_("CONTEXTO_ATIVO", "NORMAL")).toUpperCase();

        const candidatos = [];
        for (const row of dados) {
            const nome = String(row[0] || "").trim().toUpperCase();
            const indice = Number(row[5]);
            const minutos = Number(row[1]) || 0;
            if (!nome || nome.indexOf("MERCHAN") >= 0 || !isFinite(indice) || indice <= 0) continue;
            if (recentes.indexOf(nome) >= 0) continue;

            const porques = [];
            let score = indice;
            if (minutos < 30) score *= 0.9;
            // fator clube APRENDIDO (48) — substitui os boosts fixos
            const fc = fatClube[nome];
            if (typeof fc === "number" && Math.abs(fc - 1) > 0.005) {
                score *= fc;
                const info = clubes[nome] || {};
                porques.push("clube " + (fc > 1 ? "+" : "") + ((fc - 1) * 100).toFixed(0) + "%" +
                    (info.resultado ? " (" + (info.resultado === "V" ? "venceu" : info.resultado === "D" ? "perdeu" : "empatou") +
                    " " + (info.quando || "") + ")" : ""));
            }
            // fator faixa (rendimento no terço da vaga)
            const ff = (typeof fatorFaixaGE_ === "function") ? fatorFaixaGE_(nome, alvo) : 1;
            if (Math.abs(ff - 1) > 0.005) {
                score *= ff;
                porques.push((alvo <= 19 ? "abertura" : alvo <= 45 ? "miolo" : "fecho") +
                    " " + (ff > 1 ? "+" : "") + ((ff - 1) * 100).toFixed(0) + "%");
            }
            // frescor: penaliza tema já muito usado hoje
            const gasto = gastosHoje[nome] || 0;
            if (gasto >= 8) { score *= 0.85; porques.push("já " + gasto + " min hoje"); }
            // contexto Copa
            if (contexto.indexOf("COPA") >= 0 &&
                ["COPA DO MUNDO", "SELEÇÃO BRASILEIRA", "NEYMAR", "ANCELOTTI"].indexOf(nome) >= 0) {
                score *= 1.15; porques.push("Copa ativa");
            }
            candidatos.push({ nome: nome, score: score, minutos: minutos, porques: porques });
        }
        candidatos.sort((a, b) => b.score - a.score);
        if (!candidatos.length) return "";
        // top-3 explicado
        return candidatos.slice(0, 3).map((c, i) => {
            let t = (i + 1) + ". " + formatTemaDisplay_(c.nome);
            if (c.minutos < 30) t += " (amostra baixa)";
            if (c.porques.length) t += " — " + c.porques.join(", ");
            return t;
        }).join("  ·  ");
    } catch (e) {
        return "";
    }
}

// -----------------------------------------------------------------------------
// Espelho do dia — melhores janelas para break e merchans
// -----------------------------------------------------------------------------
function calcularEspelhoHojeGE_(sh) {
    const r = HOJE_GE.ESPELHO_ROW;
    const nMerchans = Math.max(2, Math.min(5, Number(getConfigValueGE_("MERCHANS_PREVISTOS", 2)) || 2));

    // Curva de referência: previsão (C) com fallback para média contextual (B)
    const grade = sh.getRange(3, 2, HOJE_GE.N_MIN, 2).getValues();
    const curva = [];
    for (let i = 0; i < HOJE_GE.N_MIN; i++) {
        const c = numOuNullGE_(grade[i][1]);
        const b = numOuNullGE_(grade[i][0]);
        curva.push(c !== null ? c : (b !== null ? b : 0));
    }

    const breaksGlobo = getPrevisaoBreaksHojeGE_();
    const ocupados = [];

    function janelaMaisFraca(tam, inicioMin, fimMin) {
        let melhor = null;
        for (let i = inicioMin; i + tam <= Math.min(fimMin, HOJE_GE.N_MIN); i++) {
            let livre = true;
            for (const o of ocupados) if (i < o.fim && i + tam > o.ini) { livre = false; break; }
            if (!livre) continue;
            // Evita coincidir com break da Globo (é hora de segurar tema forte)
            let sobreGlobo = false;
            if (breaksGlobo && breaksGlobo.janelas) {
                for (const j of breaksGlobo.janelas) {
                    if (i < j.fimMin && i + tam > j.inicioMin) { sobreGlobo = true; break; }
                }
            }
            let soma = 0;
            for (let k = i; k < i + tam; k++) soma += curva[k];
            if (sobreGlobo) soma *= 1.35; // penaliza, mas não proíbe
            if (!melhor || soma < melhor.soma) melhor = { ini: i, fim: i + tam, soma: soma };
        }
        if (melhor) ocupados.push(melhor);
        return melhor;
    }

    function fmtJanela(j) {
        if (!j) return "—";
        return "18:" + String(j.ini).padStart(2, "0") + " a 18:" +
            String(Math.min(59, j.fim - 1)).padStart(2, "0");
    }

    const linhas = [["PLANEJAMENTO DO ESPELHO", "JANELA", "OBSERVAÇÃO"]];
    const jBreak = janelaMaisFraca(3, 12, 55);
    linhas.push(["Break (3 min)", fmtJanela(jBreak), "Vale previsto de menor audiência"]);
    for (let m = 1; m <= nMerchans; m++) {
        const jm = janelaMaisFraca(2, 6, 58);
        linhas.push(["Merchan " + m + " (2 min)", fmtJanela(jm), jm ? "Vale local da curva prevista" : "Sem janela livre"]);
    }

    // Pico previsto = momento a proteger
    let pico = 0;
    for (let i = 1; i < HOJE_GE.N_MIN; i++) if (curva[i] > curva[pico]) pico = i;
    linhas.push(["Evitar merchan", "18:" + String(pico).padStart(2, "0"),
        "Pico previsto do dia; usar tema forte"]);

    if (breaksGlobo && breaksGlobo.janelas && breaksGlobo.janelas.length) {
        const j = breaksGlobo.janelas[0];
        linhas.push(["Globo em break (previsto)",
            "18:" + String(j.inicioMin).padStart(2, "0") + " a 18:" + String(j.fimMin).padStart(2, "0"),
            breaksGlobo.alerta || "Segurar tema forte nessa janela"]);
    } else {
        linhas.push(["Globo em break (previsto)", "—",
            (breaksGlobo && breaksGlobo.alerta) || "Sem padrão de breaks para hoje (alimentar aba Breaks)"]);
    }

    sh.getRange(r, 1, 12, 3).clearContent();
    sh.getRange(r, 1, linhas.length, 3).setValues(linhas);
    sh.getRange(r, 1, 1, 3).setFontWeight("bold").setBackground("#D5F5E3");
}

// -----------------------------------------------------------------------------
// Gráfico do dia — média contextual x último programa x hoje x previsão
// -----------------------------------------------------------------------------
function garantirGraficoHojeGE_(shOpt) {
    const sh = shOpt || getHojeSheetGE_();
    for (const ch of sh.getCharts()) {
        try {
            const t = ch.getOptions().get("title");
            if (t === HOJE_GE.CHART_TITLE) sh.removeChart(ch);
        } catch (e) { sh.removeChart(ch); }
    }
    const chart = sh.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(sh.getRange(2, 1, 61, 1))   // A: horas
        .addRange(sh.getRange(2, 21, 61, 1))  // U: média histórica (vermelha)
        .addRange(sh.getRange(2, 2, 61, 1))   // B: média contextual
        .addRange(sh.getRange(2, 20, 61, 1))  // T: último programa
        .addRange(sh.getRange(2, 6, 61, 1))   // F: hoje (real)
        .addRange(sh.getRange(2, 3, 61, 1))   // C: previsão
        .setNumHeaders(1)
        .setOption("title", HOJE_GE.CHART_TITLE)
        .setOption("useFirstColumnAsDomain", true)
        .setOption("legend", { position: "bottom" })
        .setOption("height", 340)
        .setOption("width", 760)
        .setOption("series", {
            0: { color: "#C0392B", lineWidth: 2 },
            1: { color: "#95A5A6", lineDashStyle: [4, 4] },
            2: { color: "#F5B041" },
            3: { color: "#2E86C1", lineWidth: 3 },
            4: { color: "#27AE60", lineDashStyle: [2, 2] }
        })
        .setPosition(HOJE_GE.ESPELHO_ROW, 5, 0, 0)
        .build();
    sh.insertChart(chart);
    return chart;
}

// Reconstrói o gráfico quando a versão do layout muda (bump em CHART_VERSAO
// a cada alteração de séries/legenda).
var HOJE_CHART_VERSAO = "3-legenda";
function garantirGraficoSeNecessarioHojeGE_(sh) {
    try {
        if (String(getConfigValueGE_("GRAFICO_HOJE_VERSAO", "")) === HOJE_CHART_VERSAO &&
            sh.getCharts().length > 0) return;
        garantirGraficoHojeGE_(sh);
        setConfigValueGE_("GRAFICO_HOJE_VERSAO", HOJE_CHART_VERSAO, "Versão do gráfico da aba Hoje");
    } catch (e) { /* gráfico é cosmético */ }
}

// -----------------------------------------------------------------------------
// Helpers
// -----------------------------------------------------------------------------
function numOuNullGE_(v) {
    return (typeof v === "number" && isFinite(v)) ? v : null;
}

function acharColunaPorHeaderGE_(sheet, texto) {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < headers.length; c++) {
        if (String(headers[c] || "").trim().toUpperCase() === texto.toUpperCase()) return c + 1;
    }
    return 0;
}

// Procura o cabeçalho nas linhas 1 E 2 (na aba Painel, SINAL/FATOR TEMA/etc.
// ficam na linha 1; a linha 2 tem os rótulos curtos STATUS/TEMA/...).
function lerColunaPainelGE_(shPainel, header, n) {
    const alvo = header.toUpperCase();
    const lastCol = shPainel.getLastColumn();
    if (!lastCol) return new Array(n).fill("");
    const linhas = shPainel.getRange(1, 1, 2, lastCol).getValues();
    let col = 0;
    for (let r = 0; r < 2 && !col; r++) {
        for (let c = 0; c < lastCol; c++) {
            if (String(linhas[r][c] || "").trim().toUpperCase() === alvo) { col = c + 1; break; }
        }
    }
    if (!col) return new Array(n).fill("");
    return shPainel.getRange(3, col, n, 1).getValues().map(r => r[0]);
}
