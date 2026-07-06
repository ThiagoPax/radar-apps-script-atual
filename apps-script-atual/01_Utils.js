// =============================================================================
// 01_Utils.js — Utilitários gerais: fórmulas, texto, HTTP, conversões
// =============================================================================

// --- Helpers de fórmula (Sheets API v4, USER_ENTERED para pt-BR) ---

function sanitizeFormula_(e) {
    let t = ((e || "") + "");
    try {
        t = t.normalize("NFKC")
    } catch (e) {}
    return t = t.replace(/[\u201C\u201D\u201E\u201F]/g, '"').replace(/[\u2018\u2019\u201A\u201B]/g, "'").replace(/[\u00A0\u202F\u2007]/g, " ").replace(/[\uFEFF\u200B-\u200D\u2060\u180E\u200E\u200F\u202A-\u202E\u2066-\u2069\u00AD]/g, "").replace(/\r/g, "").replace(/\n/g, "").trim()
}

function setFormulaUserEntered_(e, t) {
    const o = e.getSheet(),
        a = o.getParent().getId(),
        r = o.getName().replace(/'/g, "''"),
        n = "'" + r + "'!" + e.getA1Notation();
    return Sheets.Spreadsheets.Values.update({
        values: [
            [t]
        ]
    }, a, n, {
        valueInputOption: "USER_ENTERED"
    })
}

function setFormulasUserEntered_(e, t) {
    const o = e.getSheet(),
        a = o.getParent().getId(),
        r = o.getName().replace(/'/g, "''"),
        n = "'" + r + "'!" + e.getA1Notation();
    return Sheets.Spreadsheets.Values.update({
        values: t
    }, a, n, {
        valueInputOption: "USER_ENTERED"
    })
}

function setFormulaSafe_(e, t) {
    let o = sanitizeFormula_(t);
    o && "=" !== o[0] && (o = "=" + o);
    const a = 5;
    for (let r = 0; r < a; r++) try {
        return setFormulaUserEntered_(e, o)
    } catch (t) {
        if (r === a - 1) throw t;
        Utilities.sleep(200 * Math.pow(2, r))
    }
}

// --- Normalização de texto ---

function normalizeText_(e) {
    return e ? stripAccents_(e.toString()).toLowerCase().replace(/[""]/g, '"').replace(/[']/g, "'").replace(/\s+/g, " ").trim() : ""
}

function stripAccents_(e) {
    return (e || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "")
}

function containsAny_(e, t) {
    for (const o of t)
        if (e.includes(normalizeText_(o))) return !0;
    return !1
}

function escapeRegex_(e) {
    return (e || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

function capitalizeAllWords_(e) {
    return e ? e.toString().toLowerCase().split(/\s+/).map(e => e.charAt(0).toUpperCase() + e.slice(1)).join(" ") : ""
}

// --- Normalização de nome de tema (GE e outputs gerais) ---

function normalizarTema_(tema) {
    const t = (tema || "").toString().trim();
    if (!t) return t;
    const o = t.toUpperCase();
    return ["MILAN", "MANCHESTER UNITED", "F1", "COPINHA"].includes(o) ? "GIRO ESPORTE" : ("MERCHAN" === o || "BREAK" === o ? "MERCHAN/BREAK" : t)
}

// --- HTTP utilitário ---

function fetchJsonSafe_(e, t, o) {
    const a = t ? UrlFetchApp.fetch(e, t) : UrlFetchApp.fetch(e);
    if (o) {
        const e = a.getResponseCode(),
            t = a.getContentText();
        return o.checkHttp && e !== (o.okCode || 200) ? (Logger.log("Erro do servidor: HTTP " + e), Logger.log("Resposta: " + t.substring(0, 500)), {
            success: !1,
            error: `Servidor retornou HTTP ${e}. Verifique se o servidor está online.`
        }) : o.checkHtml && (t.trim().startsWith("<!DOCTYPE") || t.trim().startsWith("<html")) ? (Logger.log("Servidor retornou HTML ao invés de JSON"), Logger.log("Resposta: " + t.substring(0, 500)), {
            success: !1,
            error: "Servidor retornou página HTML ao invés de JSON. Verifique a URL do servidor e se o endpoint /transcribe existe."
        }) : JSON.parse(t)
    }
    return JSON.parse(a.getContentText())
}

function hashText_(e) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, e || "", Utilities.Charset.UTF_8).map(e => (e + 256).toString(16).slice(-2)).join("")
}

// --- Conversão de coluna numérica para letra (ex: 1→A, 27→AA) ---

function columnToLetter_(e) {
    let t = "";
    for (; e > 0;) {
        const o = (e - 1) % 26;
        t = String.fromCharCode(o + 65) + t, e = Math.floor((e - o - 1) / 26)
    }
    return t
}

// --- Fórmulas em células esparsas (não contíguas) em 1 chamada à Sheets API ---
// Usado para gravar as fórmulas da linha 63 em cada coluna de AUDIÊNCIA
// (D, H, L, P, ...) sem tocar nas colunas intermediárias (E, F, G, ...).
// cells = [{ row: 63, col: 4, formula: "=(SOMA(D3:D62))/60" }, ...]
function batchSetFormulasSparseSafe_(sheet, cells) {
    if (!cells || cells.length === 0) return;
    var ssId = sheet.getParent().getId();
    var name = sheet.getName().replace(/'/g, "''");
    var data = [];
    for (var i = 0; i < cells.length; i++) {
        var f = sanitizeFormula_(cells[i].formula);
        if (f && f[0] !== "=") f = "=" + f;
        data.push({
            range: "'" + name + "'!" + sheet.getRange(cells[i].row, cells[i].col).getA1Notation(),
            values: [[f || ""]]
        });
    }
    var maxRetries = 5;
    for (var attempt = 0; attempt < maxRetries; attempt++) {
        try {
            Sheets.Spreadsheets.Values.batchUpdate({
                valueInputOption: "USER_ENTERED",
                data: data
            }, ssId);
            return;
        } catch (e) {
            if (attempt === maxRetries - 1) throw e;
            Utilities.sleep(1000 * Math.pow(2, attempt));
        }
    }
}