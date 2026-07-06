// =============================================================================
// 06_Input_GE.js — Carregamento e mapeamento de timestamps da Gazeta Esportiva
// =============================================================================

// --- Garantir abas de saída da GE ---

function ensureGESheets_() {
    const e = SpreadsheetApp.getActive();
    let t = e.getSheetByName(CFG.GE_OUT_TOP);
    t || (t = e.insertSheet(CFG.GE_OUT_TOP), t.getRange(1, 1, 1, 4).setValues([
        ["TERMO", "CONTAGEM", "PESO POR TEMPO (MIN)", "IGNORAR"]
    ]), Logger.log(`Aba "${CFG.GE_OUT_TOP}" criada.`)), t = e.getSheetByName(CFG.GE_OUT_MACRO), t || (t = e.insertSheet(CFG.GE_OUT_MACRO), t.getRange(1, 1, 1, 5).setValues([
        ["TEMA", "MINUTOS", "%", "1\xaa APARI\xc7\xc3O", "\xdaLTIMA APARI\xc7\xc3O"]
    ]), Logger.log(`Aba "${CFG.GE_OUT_MACRO}" criada.`)), t = e.getSheetByName(CFG.GE_OUT_SUB), t || (t = e.insertSheet(CFG.GE_OUT_SUB), t.getRange(1, 1, 1, 6).setValues([
        ["TEMA", "SUBTEMA", "MINUTOS", "%", "1\xaa APARI\xc7\xc3O", "\xdaLTIMA APARI\xc7\xc3O"]
    ]), Logger.log(`Aba "${CFG.GE_OUT_SUB}" criada.`)), t = e.getSheetByName(CFG.GE_OUT_TIME), t || (t = e.insertSheet(CFG.GE_OUT_TIME), t.getRange(1, 1, 1, 6).setValues([
        ["HORA", "TEMA", "SUBTEMA", "CONTENT_TYPE", "TERMOS (top no minuto)", "TRECHO-CHAVE"]
    ]), Logger.log(`Aba "${CFG.GE_OUT_TIME}" criada.`)), t = e.getSheetByName(CFG.GE_OUT_HISTORICO), t || (t = e.insertSheet(CFG.GE_OUT_HISTORICO), Logger.log(`Aba "${CFG.GE_OUT_HISTORICO}" criada.`))
}

// --- Carregamento da aba Input (GE) ---

function loadGEInputFromSheet_() {
    const e = SpreadsheetApp.getActive().getSheetByName(CFG.GE_INPUT_SHEET);
    if (!e) return {
        rows: []
    };
    const t = e.getLastRow(),
        o = e.getLastColumn();
    if (t < 1 || o < 1) return {
        rows: []
    };
    const a = e.getRange(1, 1, t, Math.min(o, 10)).getValues(),
        r = /^\((\d{1,2}):(\d{2})(?::(\d{2}))?\)\s*/,
        n = new Map;
    for (let e = 0; e < a.length; e++) {
        const t = a[e][0];
        if (!t) continue;
        const o = t.toString().trim();
        if (!o) continue;
        if (o.startsWith("#")) continue;
        if (o.startsWith("http") || o.includes("youtube.com") || "Transcript:" === o) continue;
        if ("transcricao" === normalizeText_(o) || "transcript" === normalizeText_(o)) continue;
        const s = o.match(r);
        if (s) {
            let e;
            e = s[3] ? 60 * parseInt(s[1], 10) + parseInt(s[2], 10) : parseInt(s[1], 10);
            const t = o.replace(r, "").trim();
            t && (n.has(e) ? n.get(e).texto += " " + t : n.set(e, {
                originalMinute: e,
                originalTimestamp: s[0].replace(/[()]/g, "").trim(),
                texto: t,
                horaStr: null
            }))
        }
    }
    const s = [],
        i = Array.from(n.keys()).sort((e, t) => e - t);
    for (const e of i) s.push(n.get(e));
    return Logger.log(`GE: ${a.length} linhas na planilha → ${s.length} minutos únicos`), {
        rows: s
    }
}

// --- Processamento e mapeamento de timestamps (GE) ---

function processAndMapTimestampsGE_(e) {
    const t = e.rows;
    if (0 === t.length) return {
        rows: [],
        startIndex: -1,
        breakInfo: null
    };
    const o = [],
        a = CFG.GE_START_HOUR,
        r = CFG.GE_START_MINUTE,
        n = detectGazetaBreak_(t);
    Logger.log(`Gazeta Esportiva: ${t.length} minutos na transcrição`), n.hasBreak && Logger.log(`Break detectado: minutos ${n.breakStart} a ${n.breakEnd}`);
    const s = 60;
    let i = null,
        c = null,
        l = s;
    if (t.length < s) {
        const e = Math.min(3, s - t.length),
            o = [],
            a = new Set(t.map(e => e.originalMinute).filter(e => null != e));
        for (let r = 1; r < t.length && o.length < e; r++) {
            const n = t[r - 1].originalMinute,
                s = t[r].originalMinute;
            if (null != n && null != s && s - n >= 2)
                for (let t = n + 1; t < s && o.length < e; t++) a.has(t) || o.push(t)
        }
        for (const t of [2, 3, 4]) {
            if (o.length >= e) break;
            a.has(t) || o.includes(t) || o.push(t)
        }
        i = new Set(o), l = s - i.size, c = new Map;
        let r = 0;
        for (let e = 0; e < s; e++) i.has(e) || (c.set(e, r), r++)
    }
    for (let e = 0; e < s; e++) {
        const s = a + Math.floor((r + e) / 60),
            d = (r + e) % 60;
        if (s > CFG.GE_END_HOUR || s === CFG.GE_END_HOUR && d > CFG.GE_END_MINUTE) break;
        const u = (s + "").padStart(2, "0") + ":" + (d + "").padStart(2, "0");
        if (i && i.has(e)) {
            o.push({
                horaStr: u,
                texto: "[MERCHAN - conteúdo suprimido]",
                originalMinute: null,
                originalTimestamp: null,
                mappedMinute: e,
                isSuppressedMerchan: !0
            });
            continue
        }
        let m = t[e],
            p = e;
        if (i) {
            const o = c.get(e);
            p = Math.min(t.length - 1, Math.round(o * t.length / l)), m = t[p]
        }
        const g = n.hasBreak && p >= n.breakStart && p <= n.breakEnd;
        o.push({
            ...m,
            horaStr: u,
            mappedMinute: e,
            isBreak: g,
            isSuppressedMerchan: g
        })
    }
    return {
        rows: o,
        startIndex: 0,
        breakInfo: n
    }
}

function detectGazetaBreak_(e) {
    const t = ["volta ja ja", "voltamos ja ja", "a gente volta", "voltamos ja", "depois do intervalo", "apos o intervalo", "vamos pro intervalo"],
        o = ["ja estamos de volta", "estamos de volta", "voltamos ao vivo", "de volta ao programa", "continuando aqui", "voltando aqui"];
    let a = -1,
        r = -1;
    for (let o = 0; o < e.length; o++) {
        const r = normalizeText_(e[o].texto || "");
        for (const e of t)
            if (r.includes(normalizeText_(e))) {
                a = o, Logger.log(`Break início detectado no minuto ${o}: "${e}"`);
                break
            } if (a >= 0) break
    }
    if (a >= 0) {
        const t = a + 1,
            n = Math.min(a + 6, e.length);
        for (let a = t; a < n; a++) {
            const t = normalizeText_(e[a].texto || "");
            for (const e of o)
                if (t.includes(normalizeText_(e))) {
                    r = a - 1, Logger.log(`Break fim detectado no minuto ${a}: "${e}"`);
                    break
                } if (r >= 0) break
        }
        r < 0 && (r = Math.min(a + 2, e.length - 1), Logger.log(`Break fim não detectado, assumindo ${r-a+1} minutos`))
    }
    return {
        hasBreak: a >= 0,
        breakStart: a,
        breakEnd: r,
        breakDuration: r >= 0 ? r - a + 1 : 0
    }
}

function detectGazetaEsportivaStart_(e) {
    const t = {
        gazetaStart: ["gazeta esportiva", "boa noite", "boa tarde", "michelle", "marina"],
        preGazeta: ["gazeta news", "jornal"]
    };
    let o = -1;
    for (let a = 0; a < e.length; a++) {
        const r = normalizeText_(e[a].texto);
        t.preGazeta.some(e => r.includes(normalizeText_(e))) && (o = a);
        const n = t.gazetaStart.some(e => r.includes(normalizeText_(e))),
            s = containsAny_(r, ["futebol", "campeonato", "jogo", "time", "gol", "partida", "flamengo", "palmeiras", "corinthians", "santos"]);
        if (n && s) {
            if (o >= 0 && a > o) return a;
            if (e[a].originalMinute >= 3 || a >= 3) return a
        }
    }
    for (let t = 0; t < e.length; t++) {
        const o = normalizeText_(e[t].texto);
        if (o.includes("gazetaesportiva") || o.includes("gazeta esportiva")) return t
    }
    return -1
}