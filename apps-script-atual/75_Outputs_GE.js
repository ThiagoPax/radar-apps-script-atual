// =============================================================================
// 75_Outputs_GE.js — Escrita das abas de saída GE + helpers de formatação
// =============================================================================

function ensureSheet_(e, t, o) {
    let a = e.getSheetByName(t);
    a || (a = e.insertSheet(t));
    const r = a.getRange(1, 1, 1, o.length);
    r.getValues()[0].map(e => ((e || "") + "").trim()).join("|") !== o.join("|") && (r.setValues([o]), a.setFrozenRows(1))
}

function safeClearSheetBelowHeader_(e) {
    const t = e.getLastRow(),
        o = Math.max(e.getLastColumn(), 1);
    if (t > 1) {
        const a = e.getRange(2, 1, t - 1, o);
        a.clearContent(), a.clearDataValidations(), a.setBackground(null), a.setFontColor(null), a.setFontWeight(null)
    }
}

// --- Helper de escrita em bloco ---

function writeSheetBlock_(e, t) {
    if (t && t.clearBelowHeader && safeClearSheetBelowHeader_(e), t && t.headerValues && e.getRange(1, 1, 1, t.headerValues[0].length).setValues(t.headerValues), t && t.values && t.values.length > 0) {
        const o = t.startRow || 2,
            a = t.startCol || 1;
        e.getRange(o, a, t.values.length, t.values[0].length).setValues(t.values), t.numberFormat && e.getRange(o, t.numberFormat.startCol, t.values.length, t.numberFormat.numCols).setNumberFormat(t.numberFormat.pattern)
    }
    return e
}

// --- Batch de fórmulas com retry ---

function batchSetFormulasSafe_(range, formulas) {
    var clean = [];
    for (var r = 0; r < formulas.length; r++) {
        var row = [];
        for (var c = 0; c < formulas[r].length; c++) {
            var f = sanitizeFormula_(formulas[r][c]);
            if (f && f[0] !== "=") f = "=" + f;
            row.push(f || "");
        }
        clean.push(row);
    }
    var maxRetries = 5;
    for (var attempt = 0; attempt < maxRetries; attempt++) {
        try {
            setFormulasUserEntered_(range, clean);
            return;
        } catch (e) {
            if (attempt === maxRetries - 1) throw e;
            Utilities.sleep(1000 * Math.pow(2, attempt));
        }
    }
}

// --- Helpers de formatação ---

function formatTemaDisplay_(e) {
    if (!e) return "";
    let t = e.replace(/_/g, " ");
    return t = t.toLowerCase().split(/[\s-]+/).map((e, t, o) => e.length <= 2 && /^[a-z]+$/.test(e) ? e.toUpperCase() : e.charAt(0).toUpperCase() + e.slice(1)).join(" "), t = t.replace(/Atletico Mg/g, "Atl\xe9tico-MG").replace(/Athletico Pr/g, "Athletico-PR").replace(/Inter De Milao/g, "Inter de Mil\xe3o").replace(/Sao Paulo/g, "S\xe3o Paulo").replace(/Ceara/g, "Cear\xe1").replace(/Goias/g, "Goi\xe1s").replace(/Vitoria/g, "Vit\xf3ria").replace(/Gremio/g, "Gr\xeamio").replace(/Volei/g, "V\xf4lei").replace(/F1/gi, "F1").replace(/Nfl/g, "NFL").replace(/Nba/g, "NBA").replace(/Mma/g, "MMA"), t
}

function formatHora_(e) {
    if (!e) return "";
    let t = e.toString().trim();
    if (e instanceof Date) return `${e.getHours().toString().padStart(2,"0")}:${e.getMinutes().toString().padStart(2,"0")}`;
    const o = t.match(/^(\d{1,2}):(\d{2}):\d{2}$/);
    if (o) return o[1].padStart(2, "0") + ":" + o[2];
    const a = t.match(/^(\d{1,2}):(\d{2})$/);
    return a ? a[1].padStart(2, "0") + ":" + a[2] : t
}

function formatHoraTexto_(e) {
    return formatHora_(e) || ""
}

function formatSubtemaDisplay_(e) {
    if (!e) return "";
    let t = e.replace(/_/g, " ");
    return t = t.toLowerCase().split(/\s+/).map(e => e.charAt(0).toUpperCase() + e.slice(1)).join(" "), t
}

// (Funções de output do Jogo Aberto removidas na migração de 07/2026.)

// --- Outputs da Gazeta Esportiva ---

function buildTimelineGE_(e, t, o) {
    return buildTimeline_(e, t, o)
}

function writeAllOutputsGE_(e, t, o, a) {
    const r = computeTop20_(a, t, o);
    writeLinhaDoTempoGE_(e), writeMacroGE_(e), writeSubtemasGE_(e), writeTop20GE_(r)
}

function writeLinhaDoTempoGE_(e) {
    const t = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_TIME);
    if (!t) throw Error("Aba não encontrada: " + CFG.GE_OUT_TIME);
    writeSheetBlock_(t, {
        clearBelowHeader: !0,
        headerValues: [
            ["HORA", "TEMA", "SUBTEMA", "CONTENT_TYPE", "TERMOS (top no minuto)", "TRECHO-CHAVE"]
        ]
    }), 0 !== e.length && writeSheetBlock_(t, {
        values: e.map(e => [e.hora, e.tema, e.subtema, e.contentType, e.termos, e.trecho.substring(0, 1e3)]),
        startRow: 2,
        startCol: 1
    })
}

function writeMacroGE_(e) {
    const t = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_MACRO);
    if (!t) throw Error("Aba não encontrada: " + CFG.GE_OUT_MACRO);
    writeSheetBlock_(t, {
        clearBelowHeader: !0,
        headerValues: [
            ["TEMA", "MINUTOS", "%", "1\xaa APARI\xc7\xc3O", "\xdaLTIMA APARI\xc7\xc3O"]
        ]
    });
    const o = new Map;
    for (const t of e) {
        const e = t.tema || "GERAL";
        o.has(e) || o.set(e, {
            count: 0,
            first: t.hora,
            last: t.hora
        });
        const a = o.get(e);
        a.count += 1, a.last = t.hora
    }
    let a = 0;
    for (const [e, t] of o.entries()) "MERCHAN" !== e && (a += t.count);
    let r = null;
    const n = [];
    for (const [e, t] of o.entries()) "MERCHAN" === e ? r = {
        tema: e,
        ...t
    } : n.push({
        tema: e,
        ...t
    });
    n.sort((e, t) => t.count - e.count);
    const s = [];
    for (const e of n) {
        const t = a > 0 ? (e.count / a * 100).toFixed(1) + "%" : "0%";
        s.push([formatTemaDisplay_(e.tema), e.count, t, formatHoraTexto_(e.first), formatHoraTexto_(e.last)])
    }
    if (r && (s.push(["", "", "", "", ""]), s.push(["Merchan", r.count, "-", formatHoraTexto_(r.first), formatHoraTexto_(r.last)])), 0 !== s.length && (writeSheetBlock_(t, {
            values: s,
            startRow: 2,
            startCol: 1,
            numberFormat: {
                startCol: 4,
                numCols: 2,
                pattern: "@"
            }
        }), r)) {
        const e = s.length + 1,
            o = t.getRange(e, 1, 1, 5);
        o.setBackground("#4285F4"), o.setFontColor("#FFFFFF"), o.setFontWeight("bold")
    }
}

function writeSubtemasGE_(e) {
    const t = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_SUB);
    if (!t) throw Error("Aba não encontrada: " + CFG.GE_OUT_SUB);
    writeSheetBlock_(t, {
        clearBelowHeader: !0,
        headerValues: [
            ["TEMA", "SUBTEMA", "MINUTOS", "%", "1\xaa APARI\xc7\xc3O", "\xdaLTIMA APARI\xc7\xc3O"]
        ]
    });
    const o = new Map;
    for (const t of e) {
        const e = t.tema || "GERAL";
        if ("MERCHAN" === e) continue;
        const a = t.subtema || "OUTROS",
            r = e + "||" + a;
        o.has(r) || o.set(r, {
            tema: e,
            sub: a,
            count: 0,
            first: t.hora,
            last: t.hora
        });
        const n = o.get(r);
        n.count += 1, n.last = t.hora
    }
    let a = 0;
    for (const [e, t] of o.entries()) a += t.count;
    const r = [...o.values()];
    r.sort((e, t) => t.count - e.count);
    const n = [];
    for (const e of r) {
        const t = a > 0 ? (e.count / a * 100).toFixed(1) + "%" : "0%";
        n.push([formatTemaDisplay_(e.tema), formatSubtemaDisplay_(e.sub), e.count, t, formatHoraTexto_(e.first), formatHoraTexto_(e.last)])
    }
    0 !== n.length && writeSheetBlock_(t, {
        values: n,
        startRow: 2,
        startCol: 1,
        numberFormat: {
            startCol: 5,
            numCols: 2,
            pattern: "@"
        }
    })
}

function writeTop20GE_(e) {
    const t = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_TOP);
    if (!t) throw Error("Aba não encontrada: " + CFG.GE_OUT_TOP);
    writeSheetBlock_(t, {
        headerValues: [
            ["TERMO", "CONTAGEM", "PESO POR TEMPO (MIN)", "IGNORAR"]
        ]
    });
    const o = [],
        a = t.getLastRow();
    if (a >= 2) {
        const e = t.getRange(2, 1, a - 1, 4).getValues();
        for (const t of e) {
            const e = normalizeText_(t[0] || ""),
                a = !0 === t[3];
            e && a && o.push(e)
        }
    }
    o.length > 0 && addToPersistedStopwords_(o);
    const r = loadPersistedStopwords_(),
        n = t.getLastColumn();
    if (a > 1 && n >= 1 && t.getRange(2, 1, a - 1, Math.max(n, 5)).clearContent().clearDataValidations(), 0 === e.length) return;
    writeSheetBlock_(t, {
        values: e.map(e => [capitalizeAllWords_(e.term), e.count, e.minutes]),
        startRow: 2,
        startCol: 1
    });
    const s = t.getRange(2, 4, e.length, 1),
        i = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    s.setDataValidation(i);
    const c = e.map(e => {
        const t = e.termNorm || normalizeText_(e.term);
        return [r.has(t)]
    });
    s.setValues(c)
}