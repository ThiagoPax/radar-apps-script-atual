// =============================================================================
// 53_Dia_Semana_GE.js — Aba Dia Semana: padrões por dia útil (a valores)
// =============================================================================
// Tudo calculado por script (sem fórmulas — imune a locale). Para cada dia:
// médias (histórica / últimas 5 / últimas 3), melhor e pior minuto médio,
// janelas sugeridas de break e merchans, e temas que funcionam melhor/pior.

function atualizarDiaSemanaGE() {
    const ss = SpreadsheetApp.getActive();
    const hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    const dest = ensureAbaGE_(CFG.SHEET_DIA_SEMANA);
    if (!hist) return;
    dest.clearContents();

    const dias = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA"];
    const abrev = { "SEGUNDA": "SEG", "TERÇA": "TER", "QUARTA": "QUA",
        "QUINTA": "QUI", "SEXTA": "SEX" };

    // Blocos fechados por dia da semana (mais novo primeiro)
    const audCols = getAudColsGE_(hist);
    const porDia = { "SEGUNDA": [], "TERÇA": [], "QUARTA": [], "QUINTA": [], "SEXTA": [] };
    for (const col of audCols) {
        const h = hist.getRange(1, col).getValue();
        const d = h instanceof Date ? h : parseDateStr_(String(h || "").trim());
        if (!d) continue;
        const nome = ["DOMINGO", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"][d.getDay()];
        if (porDia[nome]) porDia[nome].push(col);
    }

    // Leitura em lote dos minutos (3–62) e médias finais (63)
    const lastCol = hist.getLastColumn();
    const minutos = hist.getRange(3, 1, 60, lastCol).getValues();
    const medias63 = hist.getRange(63, 1, 1, lastCol).getValues()[0];

    const mediaDe = cols => {
        const vs = cols.map(c => medias63[c - 1]).filter(v => typeof v === "number" && isFinite(v));
        return vs.length ? mediaArrGE_(vs) : "";
    };

    // Curva média por minuto de um conjunto de blocos
    const curvaDe = cols => {
        const curva = new Array(60).fill(null);
        for (let m = 0; m < 60; m++) {
            const vs = cols.map(c => minutos[m][c - 1])
                .filter(v => typeof v === "number" && isFinite(v));
            if (vs.length) curva[m] = mediaArrGE_(vs);
        }
        return curva;
    };

    const fmtMin = m => "18:" + String(m).padStart(2, "0");

    // Janelas de menor audiência média (para break/merchans), sem sobreposição
    const janelas = curva => {
        const ocupados = [];
        const pegar = (tam, ini, fim) => {
            let melhor = null;
            for (let i = ini; i + tam <= fim; i++) {
                if (ocupados.some(o => i < o.fim && i + tam > o.ini)) continue;
                let soma = 0, ok = true;
                for (let k = i; k < i + tam; k++) {
                    if (curva[k] === null) { ok = false; break; }
                    soma += curva[k];
                }
                if (!ok) continue;
                if (!melhor || soma < melhor.soma) melhor = { ini: i, fim: i + tam, soma: soma };
            }
            if (melhor) ocupados.push(melhor);
            return melhor ? fmtMin(melhor.ini) + "–" + fmtMin(melhor.fim - 1) : "—";
        };
        return {
            brk: pegar(3, 12, 55),
            m1: pegar(2, 6, 58), m2: pegar(2, 6, 58),
            m3: pegar(2, 6, 58), m4: pegar(2, 6, 58), m5: pegar(2, 6, 58)
        };
    };

    // Temas por dia da semana (da Base)
    const temasPorDia = temasPorDiaSemanaGE_();

    const linhas = [["DIA", "MÉDIA HISTÓRICA", "ÚLTIMAS 5", "ÚLTIMAS 3",
        "MELHOR MINUTO", "PIOR MINUTO", "JANELA BREAK", "MERCHAN 1", "MERCHAN 2",
        "MERCHAN 3", "MERCHAN 4", "MERCHAN 5", "TEMAS FORTES", "TEMAS FRACOS", "N PROGRAMAS"]];

    for (const dia of dias) {
        const cols = porDia[dia];
        if (!cols.length) {
            linhas.push([dia, "—", "—", "—", "—", "—", "—", "—", "—", "—", "—", "—", "—", "—", 0]);
            continue;
        }
        const curva = curvaDe(cols);
        let melhor = -1, pior = -1;
        for (let m = 0; m < 60; m++) {
            if (curva[m] === null) continue;
            if (melhor < 0 || curva[m] > curva[melhor]) melhor = m;
            if (pior < 0 || curva[m] < curva[pior]) pior = m;
        }
        const j = janelas(curva);
        const t = temasPorDia[abrev[dia]] || { fortes: [], fracos: [] };
        linhas.push([dia,
            mediaDe(cols), mediaDe(cols.slice(0, 5)), mediaDe(cols.slice(0, 3)),
            melhor >= 0 ? fmtMin(melhor) + " (" + curva[melhor].toFixed(2) + ")" : "—",
            pior >= 0 ? fmtMin(pior) + " (" + curva[pior].toFixed(2) + ")" : "—",
            j.brk, j.m1, j.m2, j.m3, j.m4, j.m5,
            t.fortes.map(formatTemaDisplay_).join(", ") || "—",
            t.fracos.map(formatTemaDisplay_).join(", ") || "—",
            cols.length
        ]);
    }

    dest.getRange(1, 1, linhas.length, linhas[0].length).setValues(linhas);
    dest.getRange(1, 1, 1, linhas[0].length)
        .setFontWeight("bold").setBackground("#E74C3C").setFontColor("#FFFFFF")
        .setHorizontalAlignment("center");
    dest.getRange(2, 2, 5, 3).setNumberFormat("0.00");
    dest.setFrozenRows(1);
    dest.setColumnWidth(13, 220);
    dest.setColumnWidth(14, 220);
    logGE_("DIA_SEMANA", "Aba Dia Semana reconstruída a valores (" + audCols.length + " programas)");
}

// Temas com melhor/pior índice (&/MIN médio) por dia da semana, a partir da Base
function temasPorDiaSemanaGE_() {
    const out = {};
    try {
        const base = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BASE);
        if (!base || base.getLastRow() < 2) return out;
        const vals = base.getRange(2, 1, base.getLastRow() - 1, 9).getValues();
        const acc = {}; // dia → tema → {soma, n}
        for (const r of vals) {
            const dia = String(r[1] || "").trim().toUpperCase();
            const tema = String(r[5] || "").trim().toUpperCase();
            const idx = r[4];
            if (!dia || !tema || tema.indexOf("MERCHAN") >= 0 || typeof idx !== "number") continue;
            if (!acc[dia]) acc[dia] = {};
            if (!acc[dia][tema]) acc[dia][tema] = { soma: 0, n: 0 };
            acc[dia][tema].soma += idx;
            acc[dia][tema].n++;
        }
        for (const dia in acc) {
            const rank = Object.keys(acc[dia])
                .filter(t => acc[dia][t].n >= 5)
                .map(t => ({ t: t, idx: acc[dia][t].soma / acc[dia][t].n }));
            rank.sort((a, b) => b.idx - a.idx);
            out[dia] = {
                fortes: rank.slice(0, 3).map(r => r.t),
                fracos: rank.slice(-3).reverse().map(r => r.t)
            };
        }
    } catch (e) { /* opcional */ }
    return out;
}
