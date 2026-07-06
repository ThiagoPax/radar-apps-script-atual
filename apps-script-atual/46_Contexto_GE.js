// =============================================================================
// 46_Contexto_GE.js — Contexto ativo, base normalizada (formato longo) e
// médias contextuais
// =============================================================================
// A aba Base guarda uma linha por minuto por programa:
//   Data | Dia Semana | Hora | Audiência | &/MIN | Tema | Termos | Contexto | Média Dia
// Ela alimenta Resumo, Semana, Mês, backtest e as réguas contextuais da Hoje.

var BASE_GE_HEADERS = ["DATA", "DIA SEMANA", "HORA", "AUDIÊNCIA", "&/MIN",
    "TEMA", "TERMOS", "CONTEXTO", "MÉDIA DIA"];

function getBaseSheetGE_() {
    return ensureAbaGE_(CFG.SHEET_BASE, BASE_GE_HEADERS);
}

// -----------------------------------------------------------------------------
// Reconstrução completa da Base a partir dos blocos fechados do Histórico
// -----------------------------------------------------------------------------
function reconstruirBaseGE() {
    const ss = SpreadsheetApp.getActive();
    const hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    const base = getBaseSheetGE_();
    if (!hist) { logGE_("BASE", "Histórico não encontrado"); return; }

    const cols = getAudColsGE_(hist);
    if (!cols.length) { logGE_("BASE", "Nenhum bloco fechado no Histórico"); return; }

    const mediaHist = hist.getRange(3, 2, 60, 1).getValues(); // col B
    const contextoPorData = mapaContextoPorDataGE_();
    const linhas = [];

    for (const c of cols) {
        const h = hist.getRange(1, c).getValue();
        const dataStr = h instanceof Date
            ? Utilities.formatDate(h, CFG.TIMEZONE, "dd/MM/yyyy") : String(h).trim();
        const data = parseDateStr_(dataStr);
        if (!data) continue;
        const diaSemana = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SÁB"][data.getDay()];
        const bloco = hist.getRange(3, c, 60, 3).getValues(); // AUD, &/MIN, TEMA
        const termos = hist.getRange(3, c + 3, 60, 1).getValues();

        let soma = 0, nAud = 0;
        for (let m = 0; m < 60; m++) {
            const aud = bloco[m][0];
            if (typeof aud === "number" && isFinite(aud)) { soma += aud; nAud++; }
        }
        const mediaDia = nAud ? soma / nAud : "";
        const contexto = contextoPorData[dataStr] || "";

        for (let m = 0; m < 60; m++) {
            const aud = (typeof bloco[m][0] === "number") ? bloco[m][0] : "";
            const baseMin = mediaHist[m][0];
            const idx = (typeof aud === "number" && typeof baseMin === "number" && baseMin > 0)
                ? aud / baseMin : "";
            linhas.push([dataStr, diaSemana, "18:" + String(m).padStart(2, "0"),
                aud, idx, String(bloco[m][2] || ""), String(termos[m][0] || ""),
                contexto, mediaDia]);
        }
    }

    const lastRow = base.getLastRow();
    if (lastRow > 1) base.getRange(2, 1, lastRow - 1, BASE_GE_HEADERS.length).clearContent();
    if (linhas.length) {
        base.getRange(2, 1, linhas.length, BASE_GE_HEADERS.length).setValues(linhas);
        base.getRange(2, 4, linhas.length, 2).setNumberFormat("0.00");
        base.getRange(2, 9, linhas.length, 1).setNumberFormat("0.00");
    }
    logGE_("BASE", "Base reconstruída: " + cols.length + " programas, " + linhas.length + " linhas");
}

// Acrescenta (ou substitui) o dia informado na Base — usado após o arquivamento
function appendDiaBaseGE_(programDate) {
    const dataStr = Utilities.formatDate(programDate, CFG.TIMEZONE, "dd/MM/yyyy");
    // Registra o contexto vigente para o dia, para a média contextual futura
    registrarContextoDoDiaGE_(dataStr);
    // Reconstrução completa é barata (< 100 programas) e evita duplicação
    reconstruirBaseGE();
}

// -----------------------------------------------------------------------------
// Contexto ativo
// -----------------------------------------------------------------------------
// O contexto vigente fica na aba Config (CONTEXTO_ATIVO). O histórico de
// contextos por dia fica em chaves CONTEXTO_DIA_dd/MM/yyyy, preenchidas no
// arquivamento — assim a média contextual usa apenas dias do mesmo contexto.
function registrarContextoDoDiaGE_(dataStr) {
    const contexto = String(getConfigValueGE_("CONTEXTO_ATIVO", "NORMAL"));
    setConfigValueGE_("CONTEXTO_DIA_" + dataStr, contexto, "Contexto registrado no arquivamento");
}

function mapaContextoPorDataGE_() {
    const mapa = {};
    try {
        const sh = getConfigSheetGE_();
        const last = sh.getLastRow();
        if (last < 2) return mapa;
        const vals = sh.getRange(2, 1, last - 1, 2).getValues();
        for (const [k, v] of vals) {
            const m = String(k || "").match(/^CONTEXTO_DIA_(\d{2}\/\d{2}\/\d{4})$/);
            if (m) mapa[m[1]] = String(v || "");
        }
    } catch (e) { /* mapa vazio */ }
    return mapa;
}

// -----------------------------------------------------------------------------
// Médias contextuais (réguas do dia)
// -----------------------------------------------------------------------------
// Retorna { geral, ultimos3, ultimos5, mesmoDia, contexto, mediaContexto,
//           nContexto, fator }
function mediasContextuaisGE_() {
    const out = { geral: "", ultimos3: "", ultimos5: "", mesmoDia: "",
        contexto: String(getConfigValueGE_("CONTEXTO_ATIVO", "NORMAL")),
        mediaContexto: "", nContexto: 0 };
    try {
        const dias = mediasDiariasDaBaseGE_(); // [{data, dataStr, diaSemana, media, contexto}]
        if (!dias.length) return out;

        const medias = dias.map(d => d.media);
        out.geral = mediaArrGE_(medias);
        out.ultimos3 = mediaArrGE_(medias.slice(-CFG.CONTEXTO_JANELA_CURTA));
        out.ultimos5 = mediaArrGE_(medias.slice(-CFG.CONTEXTO_JANELA_MEDIA));

        const hoje = new Date();
        const diaHoje = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SÁB"][hoje.getDay()];
        const mesmoDia = dias.filter(d => d.diaSemana === diaHoje).map(d => d.media);
        if (mesmoDia.length) out.mesmoDia = mediaArrGE_(mesmoDia);

        const ctxUpper = out.contexto.toUpperCase();
        const doContexto = dias.filter(d => String(d.contexto).toUpperCase() === ctxUpper)
            .map(d => d.media);
        out.nContexto = doContexto.length;
        // Sem amostra do contexto: usa a média curta (comportamento recente),
        // que é a melhor aproximação do "patamar atual".
        out.mediaContexto = doContexto.length >= 2
            ? mediaArrGE_(doContexto) : out.ultimos5 || out.geral;
    } catch (e) {
        logGE_("CONTEXTO", "mediasContextuaisGE_ falhou: " + e.message);
    }
    return out;
}

// Fator de escala do contexto: média contextual / média geral (limitado)
function fatorContextoGE_() {
    const m = mediasContextuaisGE_();
    if (typeof m.geral === "number" && m.geral > 0 &&
        typeof m.mediaContexto === "number" && m.mediaContexto > 0) {
        return clampGE_(m.mediaContexto / m.geral, 0.5, 1.5);
    }
    return 1;
}

// Leitura textual de um valor contra a régua contextual
function leituraContextualGE_(valor, regua) {
    if (typeof valor !== "number" || typeof regua !== "number" || regua <= 0) return "";
    const limiar = Number(getConfigValueGE_("LIMIAR_RUPTURA", 0.20)) || 0.20;
    const d = valor / regua - 1;
    if (d >= limiar) return "RUPTURA POSITIVA";
    if (d <= -limiar) return "RUPTURA NEGATIVA";
    if (d > 0.07) return "ACIMA DO CONTEXTO";
    if (d < -0.07) return "ABAIXO DO CONTEXTO";
    return "DENTRO DO CONTEXTO";
}

// Médias diárias a partir da Base (ordenadas por data crescente)
function mediasDiariasDaBaseGE_() {
    const base = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BASE);
    if (!base || base.getLastRow() < 2) return [];
    const vals = base.getRange(2, 1, base.getLastRow() - 1, 9).getValues();
    const porDia = {};
    for (const r of vals) {
        const dataStr = r[0] instanceof Date
            ? Utilities.formatDate(r[0], CFG.TIMEZONE, "dd/MM/yyyy") : String(r[0] || "").trim();
        if (!dataStr) continue;
        if (!porDia[dataStr]) {
            porDia[dataStr] = { dataStr: dataStr, data: parseDateStr_(dataStr),
                diaSemana: String(r[1] || ""), contexto: String(r[7] || ""),
                soma: 0, n: 0 };
        }
        if (typeof r[3] === "number" && isFinite(r[3])) {
            porDia[dataStr].soma += r[3];
            porDia[dataStr].n++;
        }
    }
    const dias = Object.values(porDia)
        .filter(d => d.data && d.n > 0)
        .map(d => ({ dataStr: d.dataStr, data: d.data, diaSemana: d.diaSemana,
            contexto: d.contexto, media: d.soma / d.n }));
    dias.sort((a, b) => a.data - b.data);
    return dias;
}
