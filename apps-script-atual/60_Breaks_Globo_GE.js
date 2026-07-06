// =============================================================================
// 60_Breaks_Globo_GE.js — Estatística dos breaks da Globo (não-destrutivo)
// =============================================================================
// A aba Breaks tem estrutura própria alimentada por importação de outra
// planilha ("MONITORAMENTO DE BREAKS — GLOBO"):
//   A Data | B Break 1 Entrada | C Break 1 Término | D Mudança de Programação |
//   E Break 2 Entrada | F Break 2 Término | (G+ colunas próprias)
// ESTE MÓDULO NÃO ESCREVE NAS COLUNAS A–Q. As estatísticas ficam num bloco
// exclusivo a partir da coluna R (limpo e reescrito a cada atualização).

var BREAKS_GE = {
    STATS_COL: 18,      // coluna R
    STATS_LARGURA: 9,   // R:Z
    DATA_START_ROW: 4   // acima ficam os cabeçalhos/summary da estrutura original
};

function atualizarBreaksGE() {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BREAKS);
    if (!sh) { logGE_("BREAKS", "Aba Breaks não encontrada"); return; }

    const registros = lerBreaksGE_(sh);
    const c = BREAKS_GE.STATS_COL;

    // Limpa SOMENTE o bloco de estatísticas (R em diante)
    if (sh.getMaxColumns() < c + BREAKS_GE.STATS_LARGURA) {
        sh.insertColumnsAfter(sh.getMaxColumns(), c + BREAKS_GE.STATS_LARGURA - sh.getMaxColumns());
    }
    sh.getRange(1, c, Math.max(sh.getLastRow(), 45), BREAKS_GE.STATS_LARGURA).clearContent();

    if (!registros.length) {
        sh.getRange(1, c).setValue("Sem registros de breaks legíveis (colunas A–F).");
        logGE_("BREAKS", "Nenhum registro legível na aba Breaks");
        return;
    }

    const porDiaTipo = {}; // "SEG|BREAK 1" → [registros]
    for (const r of registros) {
        const chave = r.diaSemana + "|" + r.tipo;
        (porDiaTipo[chave] = porDiaTipo[chave] || []).push(r);
    }

    const linhas = [["ESTATÍSTICAS (geradas pelo Radar — não editar)", "", "", "", "", "", "", "", ""],
        ["DIA", "TIPO", "N", "MÉDIA INÍCIO", "MEDIANA INÍCIO", "P25–P75", "DUR. MÉDIA", "OUTLIERS", "OBS"]];

    for (const dia of ["SEG", "TER", "QUA", "QUI", "SEX"]) {
        for (const tipo of ["BREAK 1", "BREAK 2"]) {
            const regs = porDiaTipo[dia + "|" + tipo] || [];
            if (!regs.length) {
                linhas.push([dia, tipo, 0, "—", "—", "—", "—", "—", "sem amostra"]);
                continue;
            }
            const inicios = regs.map(r => r.inicioMin).sort((a, b) => a - b);
            const durs = regs.map(r => r.durMin).filter(d => typeof d === "number" && d > 0);
            const p25 = percentilGE_(inicios, 25);
            const p75 = percentilGE_(inicios, 75);
            const iqr = p75 - p25;
            const outliers = inicios.filter(v => v < p25 - 1.5 * iqr || v > p75 + 1.5 * iqr);
            linhas.push([dia, tipo, regs.length,
                fmtMinutoGE_(mediaArrGE_(inicios)),
                fmtMinutoGE_(percentilGE_(inicios, 50)),
                fmtMinutoGE_(p25) + "–" + fmtMinutoGE_(p75),
                durs.length ? mediaArrGE_(durs).toFixed(1) + " min" : "—",
                outliers.length ? outliers.map(fmtMinutoGE_).join(", ") : "nenhum",
                dia === "QUA" ? "padrão próprio (futebol na Globo)" : ""
            ]);
        }
    }

    // Alertas do dia
    linhas.push(["", "", "", "", "", "", "", "", ""]);
    linhas.push(["ALERTAS DE HOJE", "", "", "", "", "", "", "", ""]);
    for (const a of alertasBreaksHojeGE_()) linhas.push([a, "", "", "", "", "", "", "", ""]);

    sh.getRange(1, c, linhas.length, BREAKS_GE.STATS_LARGURA).setValues(linhas);
    sh.getRange(1, c).setFontWeight("bold").setBackground("#D6EAF8");
    sh.getRange(2, c, 1, BREAKS_GE.STATS_LARGURA).setFontWeight("bold");
    logGE_("BREAKS", "Estatísticas atualizadas em R1+: " + registros.length +
        " breaks lidos (estrutura original preservada)");
}

// Leitura tolerante da estrutura original (A Data | B/C Break1 | E/F Break2)
function lerBreaksGE_(sh) {
    const out = [];
    const lastRow = sh.getLastRow();
    if (lastRow < BREAKS_GE.DATA_START_ROW) return out;
    const vals = sh.getRange(BREAKS_GE.DATA_START_ROW, 1,
        lastRow - BREAKS_GE.DATA_START_ROW + 1, 6).getValues();

    for (const r of vals) {
        const data = r[0] instanceof Date ? r[0] : parseDateStr_(String(r[0] || "").trim());
        if (!data) continue; // linhas de summary/erro/vazias
        const diaSemana = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SÁB"][data.getDay()];

        const pares = [
            { ini: r[1], fim: r[2], tipo: "BREAK 1" },
            { ini: r[4], fim: r[5], tipo: "BREAK 2" }
        ];
        for (const p of pares) {
            const inicioMin = parseMinuto18GE_(p.ini);
            if (inicioMin === null) continue;
            const fimMin = parseMinuto18GE_(p.fim);
            const durMin = (fimMin !== null && fimMin >= inicioMin) ? fimMin - inicioMin : null;
            out.push({
                data: data, diaSemana: diaSemana, tipo: p.tipo,
                inicioMin: inicioMin,
                fimMin: fimMin !== null ? fimMin : inicioMin + 3,
                durMin: durMin
            });
        }
    }
    return out;
}

// Previsão de janelas de break da Globo para hoje (usada pela Hoje/predição)
function getPrevisaoBreaksHojeGE_() {
    const out = { janelas: [], alerta: "" };
    try {
        const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BREAKS);
        if (!sh) return out;
        const registros = lerBreaksGE_(sh);
        if (!registros.length) return out;

        const hoje = new Date();
        const diaHoje = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SÁB"][hoje.getDay()];
        if (diaHoje === "QUA") out.alerta = "Hoje é quarta-feira; usando padrão específico de quarta.";

        for (const tipo of ["BREAK 1", "BREAK 2"]) {
            let doDia = registros.filter(r => r.diaSemana === diaHoje && r.tipo === tipo);
            if (doDia.length < 3) {
                doDia = registros.filter(r => r.tipo === tipo &&
                    (diaHoje === "QUA" ? r.diaSemana === "QUA" : r.diaSemana !== "QUA"));
                if (!out.alerta && doDia.length) {
                    out.alerta = "Amostra baixa para " + diaHoje + "; usando padrão geral (quartas separadas).";
                }
            }
            if (!doDia.length) continue;
            const inicios = doDia.map(r => r.inicioMin).sort((a, b) => a - b);
            const durTipica = Math.round(mediaArrGE_(
                doDia.map(r => r.durMin).filter(d => d)) || 3);
            const ini = Math.round(percentilGE_(inicios, 50));
            out.janelas.push({ inicioMin: ini, fimMin: Math.min(59, ini + durTipica), tipo: tipo });
        }
        out.janelas.sort((a, b) => a.inicioMin - b.inicioMin);
    } catch (e) {
        logGE_("BREAKS", "getPrevisaoBreaksHojeGE_ falhou: " + e.message);
    }
    return out;
}

function alertasBreaksHojeGE_() {
    const alertas = [];
    const diaHoje = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SÁB"][new Date().getDay()];
    if (diaHoje === "QUA") alertas.push("Hoje é quarta-feira; usar padrão específico de quarta.");
    const prev = getPrevisaoBreaksHojeGE_();
    for (const j of prev.janelas) {
        alertas.push((j.tipo || "Break") + ": Globo deve entrar perto de " + fmtMinutoGE_(j.inicioMin) +
            " e voltar ~" + fmtMinutoGE_(j.fimMin) + "; evitar bloco fraco no retorno.");
    }
    if (!alertas.length) alertas.push("Sem padrão suficiente; conferir importação da aba Breaks.");
    return alertas;
}

// --- helpers ---
function parseMinuto18GE_(v) {
    if (v instanceof Date) {
        return v.getHours() === 18 ? v.getMinutes() : null;
    }
    const m = String(v || "").trim().match(/^18[:h]([0-5]\d)/i);
    return m ? parseInt(m[1], 10) : null;
}

function fmtMinutoGE_(min) {
    if (typeof min !== "number" || !isFinite(min)) return "—";
    return "18:" + String(Math.round(Math.max(0, Math.min(59, min)))).padStart(2, "0");
}

function percentilGE_(sortedArr, p) {
    if (!sortedArr.length) return NaN;
    const idx = (p / 100) * (sortedArr.length - 1);
    const lo = Math.floor(idx), hi = Math.ceil(idx);
    return sortedArr[lo] + (sortedArr[hi] - sortedArr[lo]) * (idx - lo);
}
