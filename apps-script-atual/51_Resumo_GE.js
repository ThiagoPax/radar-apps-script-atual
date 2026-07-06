// =============================================================================
// 51_Resumo_GE.js — Aba Resumo: mapa de calor semanal da audiência média final
// =============================================================================
// Uma linha por semana; colunas Segunda..Sexta com a média final de cada
// programa. Formatação condicional dinâmica (vermelho → branco → verde)
// cobrindo toda a área planejada, não só as células preenchidas.

var RESUMO_GE = {
    HEADERS: ["SEMANA", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "MÉDIA SEMANA"],
    MAX_LINHAS: 200 // área coberta pela formatação condicional
};

function atualizarResumoGE() {
    const sh = ensureAbaGE_(CFG.SHEET_RESUMO, RESUMO_GE.HEADERS);
    const dias = mediasDiariasDaBaseGE_();

    // Agrupa por semana (segunda-feira como âncora)
    const semanas = {};
    for (const d of dias) {
        const seg = new Date(d.data);
        const dow = seg.getDay(); // 0=dom..6=sáb
        seg.setDate(seg.getDate() - ((dow + 6) % 7)); // volta até segunda
        const chave = Utilities.formatDate(seg, CFG.TIMEZONE, "yyyy-MM-dd");
        if (!semanas[chave]) {
            const sex = new Date(seg);
            sex.setDate(sex.getDate() + 4);
            semanas[chave] = {
                seg: seg,
                rotulo: Utilities.formatDate(seg, CFG.TIMEZONE, "dd/MM") + " a " +
                    Utilities.formatDate(sex, CFG.TIMEZONE, "dd/MM"),
                porDia: {}
            };
        }
        semanas[chave].porDia[d.diaSemana] = d.media;
    }

    const chaves = Object.keys(semanas).sort();
    const linhas = [];
    for (const k of chaves) {
        const s = semanas[k];
        const vals = ["SEG", "TER", "QUA", "QUI", "SEX"].map(dd =>
            typeof s.porDia[dd] === "number" ? s.porDia[dd] : "");
        const numericos = vals.filter(v => typeof v === "number");
        const media = numericos.length ? mediaArrGE_(numericos) : "";
        linhas.push([s.rotulo].concat(vals).concat([media]));
    }

    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, RESUMO_GE.HEADERS.length).clearContent();
    if (linhas.length) {
        sh.getRange(2, 1, linhas.length, RESUMO_GE.HEADERS.length).setValues(linhas);
        sh.getRange(2, 2, linhas.length, 6).setNumberFormat("0.00");
    }
    sh.setColumnWidth(1, 130);

    aplicarHeatmapResumoGE_(sh);
    logGE_("RESUMO", "Resumo atualizado: " + linhas.length + " semanas");
}

// Mapa de calor sobre toda a área planejada (B2:F{MAX_LINHAS})
function aplicarHeatmapResumoGE_(sh) {
    const alvo = sh.getRange(2, 2, RESUMO_GE.MAX_LINHAS, 5);
    const regras = sh.getConditionalFormatRules().filter(rule =>
        !rule.getRanges().some(r => r.getColumn() >= 2 && r.getColumn() <= 6 &&
            r.getSheet().getName() === sh.getName()));
    regras.push(SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue("#E74C3C", SpreadsheetApp.InterpolationType.MIN, "0")
        .setGradientMidpointWithValue("#FDFEFE", SpreadsheetApp.InterpolationType.PERCENTILE, "50")
        .setGradientMaxpointWithValue("#27AE60", SpreadsheetApp.InterpolationType.MAX, "0")
        .setRanges([alvo])
        .build());
    sh.setConditionalFormatRules(regras);
}
