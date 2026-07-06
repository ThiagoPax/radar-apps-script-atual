// =============================================================================
// 52_Semana_Mes_GE.js — Visões executivas Semana e Mês (a partir da Base)
// =============================================================================

function atualizarSemanaGE() {
    const sh = ensureAbaGE_(CFG.SHEET_SEMANA);
    const dias = mediasDiariasDaBaseGE_();
    sh.clearContents();
    if (!dias.length) {
        sh.getRange(1, 1).setValue("Sem dados na Base ainda.");
        return;
    }

    const hoje = new Date();
    const inicioSemana = new Date(hoje);
    inicioSemana.setDate(hoje.getDate() - ((hoje.getDay() + 6) % 7));
    inicioSemana.setHours(0, 0, 0, 0);
    const inicioAnterior = new Date(inicioSemana);
    inicioAnterior.setDate(inicioAnterior.getDate() - 7);

    const semanaAtual = dias.filter(d => d.data >= inicioSemana);
    const semanaAnterior = dias.filter(d => d.data >= inicioAnterior && d.data < inicioSemana);

    const mediaAtual = semanaAtual.length ? mediaArrGE_(semanaAtual.map(d => d.media)) : "";
    const mediaAnterior = semanaAnterior.length ? mediaArrGE_(semanaAnterior.map(d => d.media)) : "";
    const variacao = (typeof mediaAtual === "number" && typeof mediaAnterior === "number" && mediaAnterior > 0)
        ? mediaAtual / mediaAnterior - 1 : "";

    let melhor = null, pior = null;
    for (const d of semanaAtual) {
        if (!melhor || d.media > melhor.media) melhor = d;
        if (!pior || d.media < pior.media) pior = d;
    }

    const temas = rankingTemasPeriodoGE_(inicioSemana, null);
    const clubes = resumoClubesPeriodoGE_(inicioSemana, null);

    let leitura = "Semana em linha com a anterior.";
    if (typeof variacao === "number") {
        if (variacao > 0.05) leitura = "Semana em alta (" + (variacao * 100).toFixed(1) + "% vs anterior).";
        else if (variacao < -0.05) leitura = "Semana em queda (" + (variacao * 100).toFixed(1) + "% vs anterior).";
    }

    const linhas = [
        ["VISÃO SEMANAL", "", ""],
        ["Média da semana", mediaAtual, ""],
        ["Média da semana anterior", mediaAnterior, ""],
        ["Variação", variacao, ""],
        ["Melhor programa", melhor ? melhor.dataStr + " (" + melhor.media.toFixed(2) + ")" : "—", ""],
        ["Pior programa", pior ? pior.dataStr + " (" + pior.media.toFixed(2) + ")" : "—", ""],
        ["", "", ""],
        ["TEMAS MAIS FORTES", "ÍNDICE", "MINUTOS"]
    ];
    for (const t of temas.slice(0, 5)) linhas.push([formatTemaDisplay_(t.tema), t.indice, t.minutos]);
    linhas.push(["", "", ""]);
    linhas.push(["TEMAS MAIS FRACOS", "ÍNDICE", "MINUTOS"]);
    for (const t of temas.slice(-5).reverse()) linhas.push([formatTemaDisplay_(t.tema), t.indice, t.minutos]);
    linhas.push(["", "", ""]);
    linhas.push(["CLUBES NA SEMANA", "", ""]);
    for (const c of clubes) linhas.push([c.clube, c.resumo, ""]);
    linhas.push(["", "", ""]);
    linhas.push(["LEITURA", leitura, ""]);

    sh.getRange(1, 1, linhas.length, 3).setValues(linhas);
    sh.getRange(1, 1).setFontWeight("bold").setFontSize(12);
    [8, linhas.length].forEach(() => {});
    sh.getRange(2, 2, 3, 1).setNumberFormat("0.00");
    sh.getRange(4, 2).setNumberFormat("+0.0%;-0.0%");
    sh.setColumnWidth(1, 220);
    sh.setColumnWidth(2, 260);
    logGE_("SEMANA", "Visão semanal atualizada");
}

function atualizarMesGE() {
    const sh = ensureAbaGE_(CFG.SHEET_MES);
    const dias = mediasDiariasDaBaseGE_();
    sh.clearContents();
    if (!dias.length) {
        sh.getRange(1, 1).setValue("Sem dados na Base ainda.");
        return;
    }

    const hoje = new Date();
    const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    const inicioAnterior = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);

    const mesAtual = dias.filter(d => d.data >= inicioMes);
    const mesAnterior = dias.filter(d => d.data >= inicioAnterior && d.data < inicioMes);

    const mediaAtual = mesAtual.length ? mediaArrGE_(mesAtual.map(d => d.media)) : "";
    const mediaAnterior = mesAnterior.length ? mediaArrGE_(mesAnterior.map(d => d.media)) : "";
    const tendencia = (typeof mediaAtual === "number" && typeof mediaAnterior === "number" && mediaAnterior > 0)
        ? mediaAtual / mediaAnterior - 1 : "";

    const ordenados = mesAtual.slice().sort((a, b) => b.media - a.media);
    const temas = rankingTemasPeriodoGE_(inicioMes, null);
    const termos = rankingTermosPeriodoGE_(inicioMes, null);

    let alerta = "";
    if (typeof tendencia === "number") {
        if (tendencia < -0.10) alerta = "⚠️ Queda forte contra o mês anterior — revisar pauta e espelho.";
        else if (tendencia > 0.10) alerta = "✅ Crescimento forte contra o mês anterior.";
    }

    const linhas = [
        ["VISÃO MENSAL — " + Utilities.formatDate(inicioMes, CFG.TIMEZONE, "MM/yyyy"), "", ""],
        ["Média do mês", mediaAtual, mesAtual.length + " programas"],
        ["Média do mês anterior", mediaAnterior, mesAnterior.length + " programas"],
        ["Tendência", tendencia, alerta],
        ["", "", ""],
        ["MELHORES DIAS", "MÉDIA", ""]
    ];
    for (const d of ordenados.slice(0, 3)) linhas.push([d.dataStr + " (" + d.diaSemana + ")", d.media, ""]);
    linhas.push(["PIORES DIAS", "MÉDIA", ""]);
    for (const d of ordenados.slice(-3).reverse()) linhas.push([d.dataStr + " (" + d.diaSemana + ")", d.media, ""]);
    linhas.push(["", "", ""]);
    linhas.push(["RANKING DE TEMAS DO MÊS", "ÍNDICE", "MINUTOS"]);
    for (const t of temas.slice(0, 10)) linhas.push([formatTemaDisplay_(t.tema), t.indice, t.minutos]);
    linhas.push(["", "", ""]);
    linhas.push(["RANKING DE TERMOS DO MÊS", "OCORRÊNCIAS", ""]);
    for (const t of termos.slice(0, 10)) linhas.push([t.termo, t.count, ""]);

    sh.getRange(1, 1, linhas.length, 3).setValues(linhas);
    sh.getRange(1, 1).setFontWeight("bold").setFontSize(12);
    sh.getRange(2, 2, 2, 1).setNumberFormat("0.00");
    sh.getRange(4, 2).setNumberFormat("+0.0%;-0.0%");
    sh.setColumnWidth(1, 240);
    logGE_("MES", "Visão mensal atualizada");
}

// -----------------------------------------------------------------------------
// Rankings por período (a partir da Base)
// -----------------------------------------------------------------------------
// Índice do tema = média do &/MIN dos minutos daquele tema no período.
function rankingTemasPeriodoGE_(inicio, fim) {
    const base = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BASE);
    if (!base || base.getLastRow() < 2) return [];
    const vals = base.getRange(2, 1, base.getLastRow() - 1, 9).getValues();
    const porTema = {};
    for (const r of vals) {
        const data = r[0] instanceof Date ? r[0] : parseDateStr_(String(r[0] || ""));
        if (!data || (inicio && data < inicio) || (fim && data >= fim)) continue;
        const tema = String(r[5] || "").trim().toUpperCase();
        const idx = r[4];
        if (!tema || tema.indexOf("MERCHAN") >= 0 || typeof idx !== "number") continue;
        if (!porTema[tema]) porTema[tema] = { soma: 0, n: 0 };
        porTema[tema].soma += idx;
        porTema[tema].n++;
    }
    const out = Object.keys(porTema)
        .filter(t => porTema[t].n >= 3)
        .map(t => ({ tema: t, indice: porTema[t].soma / porTema[t].n, minutos: porTema[t].n }));
    out.sort((a, b) => b.indice - a.indice);
    return out;
}

function rankingTermosPeriodoGE_(inicio, fim) {
    const base = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BASE);
    if (!base || base.getLastRow() < 2) return [];
    const vals = base.getRange(2, 1, base.getLastRow() - 1, 9).getValues();
    const counts = {};
    for (const r of vals) {
        const data = r[0] instanceof Date ? r[0] : parseDateStr_(String(r[0] || ""));
        if (!data || (inicio && data < inicio) || (fim && data >= fim)) continue;
        const termos = String(r[6] || "").split(/[;,|]/);
        for (let t of termos) {
            t = t.trim();
            if (t.length < 3) continue;
            counts[t] = (counts[t] || 0) + 1;
        }
    }
    const out = Object.keys(counts).map(t => ({ termo: t, count: counts[t] }));
    out.sort((a, b) => b.count - a.count);
    return out;
}

// Resumo dos jogos dos clubes foco no período (aba Clubes)
function resumoClubesPeriodoGE_(inicio, fim) {
    const out = [];
    try {
        const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_CLUBES);
        if (!sh || sh.getLastRow() < 2) return out;
        const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
        const porClube = {};
        for (const r of vals) {
            const data = r[0] instanceof Date ? r[0] : parseDateStr_(String(r[0] || ""));
            if (!data || (inicio && data < inicio) || (fim && data >= fim)) continue;
            const clube = String(r[1] || "").trim().toUpperCase();
            if (CFG.CLUBES_FOCO.indexOf(clube) < 0) continue;
            const resultado = String(r[9] || "").trim().toUpperCase(); // col J: V/E/D
            if (!porClube[clube]) porClube[clube] = { V: 0, E: 0, D: 0 };
            if (resultado.charAt(0) === "V") porClube[clube].V++;
            else if (resultado.charAt(0) === "E") porClube[clube].E++;
            else if (resultado.charAt(0) === "D") porClube[clube].D++;
        }
        for (const c of CFG.CLUBES_FOCO) {
            if (porClube[c]) {
                out.push({ clube: c, resumo: porClube[c].V + "V " + porClube[c].E + "E " + porClube[c].D + "D" });
            }
        }
    } catch (e) { /* sem dados de clubes */ }
    return out;
}
