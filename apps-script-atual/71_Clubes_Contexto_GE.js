// =============================================================================
// 71_Clubes_Contexto_GE.js — Contexto dos clubes para a predição + rotina
// automática + diagnóstico da API-Football
// =============================================================================
// A aba Clubes é escrita por 70_Calendario_Clubes_GE (API-Football):
//  A DATA | B CLUBE | C ADVERSÁRIO | D COMPETIÇÃO | E FASE | F MANDO |
//  G PLACAR | H GOLS PRÓ | I GOLS CONTRA | J RESULTADO | K STATUS |
//  L IMPORTÂNCIA | M DIAS P/ GE

// Atualização silenciosa para a rotina das 06:00 (sem UI)
function atualizarCalendarioClubesSemUI_() {
    try {
        atualizarCalendarioClubes_();
        logGE_("CLUBES", "Calendário atualizado via API-Football");
        return true;
    } catch (e) {
        logGE_("CLUBES", "⚠️ Falha na atualização automática: " + e.message +
            " — mantendo dados existentes (atualize manualmente se necessário)");
        return false;
    }
}

// Contexto editorial de hoje por clube foco:
// { CORINTHIANS: {resultado:"V"|"E"|"D"|"", quando:"ontem"|"hoje"|..., classico,
//                 decisivo, competicao, proximoJogo, relevancia} }
function getContextoClubesHojeGE_() {
    const out = {};
    try {
        const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_CLUBES);
        if (!sh || sh.getLastRow() < 2) return out;
        const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues();
        const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
        const umDia = 24 * 3600 * 1000;

        for (const clube of CFG.CLUBES_FOCO) out[clube] = {
            resultado: "", quando: "", classico: false, decisivo: false,
            competicao: "", proximoJogo: "", relevancia: 0
        };

        for (const r of vals) {
            const clube = String(r[1] || "").trim().toUpperCase();
            if (!out[clube]) continue;
            const data = r[0] instanceof Date ? new Date(r[0]) : parseDateStr_(String(r[0] || ""));
            if (!data) continue;
            data.setHours(0, 0, 0, 0);
            const deltaDias = Math.round((data - hoje) / umDia);
            const status = String(r[10] || "").toUpperCase();
            const resultado = String(r[9] || "").trim().toUpperCase();
            const importancia = Number(r[11]) || 0;
            const adversario = String(r[2] || "");
            const classico = ehClassicoPaulistaGE_(clube, adversario);

            // Último jogo realizado (até 3 dias atrás pesa na pauta)
            if (status === "REALIZADO" && deltaDias <= 0 && deltaDias >= -3 && resultado) {
                const c = out[clube];
                if (!c._ultimoDelta || deltaDias > c._ultimoDelta) {
                    c._ultimoDelta = deltaDias;
                    c.resultado = resultado.charAt(0);
                    c.quando = deltaDias === 0 ? "hoje" : deltaDias === -1 ? "ontem" : deltaDias + "d";
                    c.classico = classico;
                    c.decisivo = importancia >= 4;
                    c.competicao = String(r[3] || "");
                    c.relevancia = Math.max(c.relevancia, importancia + (classico ? 1 : 0));
                }
            }
            // Próximo jogo (hoje ou amanhã aquece a pauta)
            if (status === "AGENDADO" && deltaDias >= 0 && deltaDias <= 1) {
                const c = out[clube];
                if (!c.proximoJogo) {
                    c.proximoJogo = (deltaDias === 0 ? "hoje" : "amanhã") + " x " + adversario +
                        (classico ? " (clássico)" : "");
                    c.relevancia = Math.max(c.relevancia, importancia);
                }
            }
        }
        for (const k in out) delete out[k]._ultimoDelta;
    } catch (e) {
        logGE_("CLUBES", "getContextoClubesHojeGE_ falhou: " + e.message);
    }
    return out;
}

function ehClassicoPaulistaGE_(clube, adversario) {
    const rivais = {
        "CORINTHIANS": ["PALMEIRAS", "SÃO PAULO", "SANTOS"],
        "PALMEIRAS": ["CORINTHIANS", "SÃO PAULO", "SANTOS"],
        "SÃO PAULO": ["CORINTHIANS", "PALMEIRAS", "SANTOS"],
        "SANTOS": ["CORINTHIANS", "PALMEIRAS", "SÃO PAULO"]
    };
    const adv = String(adversario || "").toUpperCase();
    return (rivais[clube] || []).some(r => adv.indexOf(r) >= 0);
}

// -----------------------------------------------------------------------------
// Diagnóstico da API-Football (menu) — o usuário suspeita que a chave parou
// -----------------------------------------------------------------------------
function testarApiFootballGE() {
    const ui = SpreadsheetApp.getUi();
    const apiKey = PropertiesService.getScriptProperties().getProperty("API_FOOTBALL_KEY");
    if (!apiKey) {
        ui.alert("❌ API_FOOTBALL_KEY não está configurada nas propriedades do script.");
        return;
    }
    try {
        const resp = UrlFetchApp.fetch(CFG.API_FOOTBALL_ENDPOINT + "/status", {
            method: "get",
            headers: { "x-apisports-key": apiKey },
            muteHttpExceptions: true
        });
        const code = resp.getResponseCode();
        const json = JSON.parse(resp.getContentText());
        if (code === 200 && json.response && json.response.account) {
            const r = json.response;
            const req = r.requests || {};
            const sub = r.subscription || {};
            ui.alert("✅ API-Football respondendo.\n\n" +
                "Plano: " + (sub.plan || "?") + (sub.active === false ? " (INATIVO!)" : "") + "\n" +
                "Requisições hoje: " + (req.current || 0) + " de " + (req.limit_day || "?") + "\n" +
                "Fim da assinatura: " + (sub.end || "?"));
            logGE_("CLUBES", "API-Football OK: " + (req.current || 0) + "/" + (req.limit_day || "?") + " req hoje");
        } else {
            const erros = JSON.stringify(json.errors || json.message || json).substring(0, 300);
            ui.alert("⚠️ API-Football retornou problema (HTTP " + code + "):\n\n" + erros +
                "\n\nSe a chave expirou, gere outra em dashboard.api-football.com e atualize " +
                "API_FOOTBALL_KEY nas propriedades do script.");
            logGE_("CLUBES", "API-Football com problema: HTTP " + code + " " + erros);
        }
    } catch (e) {
        ui.alert("❌ Falha ao chamar a API-Football: " + e.message);
        logGE_("CLUBES", "API-Football inacessível: " + e.message);
    }
}

// =============================================================================
// Fonte alternativa gratuita: TheSportsDB (a API-Football Free não cobre 2026)
// =============================================================================
// Usa a chave pública de teste ("3"). Cobre bem a Série A; copas têm cobertura
// parcial via últimos/próximos jogos por time. IDs de time ficam em Config.

var TSDB_GE = {
    BASE: "https://www.thesportsdb.com/api/v1/json/3",
    LIGA_SERIE_A: "4351", // Brasileirão Série A
    NOMES_BUSCA: {
        "CORINTHIANS": "Corinthians",
        "PALMEIRAS": "Palmeiras",
        "SÃO PAULO": "Sao Paulo",
        "SANTOS": "Santos"
    }
};

function tsdbFetchGE_(path) {
    const resp = UrlFetchApp.fetch(TSDB_GE.BASE + path, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    try { return JSON.parse(resp.getContentText()); } catch (e) { return null; }
}

function tsdbTeamIdGE_(clube) {
    const chave = "TSDB_ID_" + clube.replace(/\s/g, "_");
    const cache = getConfigValueGE_(chave, "");
    if (cache) return String(cache);
    const json = tsdbFetchGE_("/searchteams.php?t=" +
        encodeURIComponent(TSDB_GE.NOMES_BUSCA[clube] || clube));
    if (!json || !json.teams) return "";
    for (const t of json.teams) {
        if (String(t.strCountry) === "Brazil" && String(t.strSport) === "Soccer") {
            setConfigValueGE_(chave, t.idTeam, "ID TheSportsDB de " + clube);
            return String(t.idTeam);
        }
    }
    return "";
}

function tsdbEventoParaLinhaGE_(ev, clube, teamId) {
    const data = parseDateStr_((ev.dateEvent || "").split("-").reverse().join("/")) ||
        (ev.dateEvent ? new Date(ev.dateEvent + "T12:00:00") : null);
    if (!data) return null;
    const isHome = String(ev.idHomeTeam) === teamId;
    const adversario = isHome ? (ev.strAwayTeam || "") : (ev.strHomeTeam || "");
    const gp = isHome ? ev.intHomeScore : ev.intAwayScore;
    const gc = isHome ? ev.intAwayScore : ev.intHomeScore;
    const realizado = gp !== null && gp !== "" && gp !== undefined &&
        gc !== null && gc !== "" && gc !== undefined;
    let resultado = "", placar = "";
    if (realizado) {
        const gpn = Number(gp), gcn = Number(gc);
        placar = gpn + "×" + gcn;
        resultado = gpn > gcn ? "V" : (gpn < gcn ? "D" : "E");
    }
    let importancia = 2;
    if (ehClassicoPaulistaGE_(clube, adversario)) importancia = Math.min(5, importancia + 1);
    const round = String(ev.intRound || "");
    const umDia = 24 * 3600 * 1000;
    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    const diasPGe = Math.round((data - hoje) / umDia);
    return [Utilities.formatDate(data, CFG.TIMEZONE, "dd/MM/yyyy"), clube, adversario,
        String(ev.strLeague || ""), round ? "Rodada " + round : "",
        isHome ? "CASA" : "FORA", placar,
        realizado ? Number(gp) : "", realizado ? Number(gc) : "",
        resultado, realizado ? "REALIZADO" : "AGENDADO", importancia, diasPGe];
}

function atualizarClubesViaTheSportsDBGE_() {
    const linhas = [];
    const vistos = {};

    // Temporada completa da Série A (um fetch para todos os clubes)
    const temporadas = ["2026", "2025-2026"];
    let eventosLiga = [];
    for (const s of temporadas) {
        const json = tsdbFetchGE_("/eventsseason.php?id=" + TSDB_GE.LIGA_SERIE_A + "&s=" + s);
        if (json && json.events && json.events.length) { eventosLiga = json.events; break; }
    }

    for (const clube of CFG.CLUBES_FOCO) {
        const teamId = tsdbTeamIdGE_(clube);
        if (!teamId) { logGE_("CLUBES", "TheSportsDB: id não encontrado para " + clube); continue; }

        const eventos = eventosLiga.filter(ev =>
            String(ev.idHomeTeam) === teamId || String(ev.idAwayTeam) === teamId);

        // Complementa com últimos/próximos (pega copas e estaduais)
        for (const path of ["/eventslast.php?id=" + teamId, "/eventsnext.php?id=" + teamId]) {
            const json = tsdbFetchGE_(path);
            const extra = json && (json.results || json.events);
            if (extra) eventos.push.apply(eventos, extra);
        }

        for (const ev of eventos) {
            if (!ev || vistos[ev.idEvent]) continue;
            vistos[ev.idEvent] = true;
            const linha = tsdbEventoParaLinhaGE_(ev, clube, teamId);
            if (linha) linhas.push(linha);
        }
        Utilities.sleep(400); // gentileza com a API pública
    }

    if (!linhas.length) return 0;
    return mesclarEEscreverClubesGE_(linhas, "TheSportsDB");
}

function escreverAbaClubesGE_(linhas, fonte) {
    const sh = ensureAbaGE_(CFG.SHEET_CLUBES);
    const headers = ["DATA", "CLUBE", "ADVERSÁRIO", "COMPETIÇÃO", "FASE", "MANDO",
        "PLACAR", "GOLS PRÓ", "GOLS CONTRA", "RESULTADO", "STATUS", "IMPORTÂNCIA", "DIAS P/ GE"];
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight("bold").setBackground("#2C3E50").setFontColor("#FFFFFF")
        .setHorizontalAlignment("center");
    sh.setFrozenRows(1);
    if (linhas.length) {
        // DATA como Date de verdade: a análise do Contexto Jogo compara datas
        // e quebra com strings.
        const comDatas = linhas.map(l => {
            const copia = l.slice();
            const d = copia[0] instanceof Date ? copia[0] : parseDateStr_(String(copia[0] || ""));
            if (d) copia[0] = d;
            return copia;
        });
        sh.getRange(2, 1, comDatas.length, headers.length).setValues(comDatas);
        sh.getRange(2, 1, comDatas.length, 1).setNumberFormat("dd/MM/yyyy");
    }
    sh.getRange(1, headers.length + 2).setValue("Fonte: " + fonte + " — " +
        Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy HH:mm"));
}

// Orquestrador: API-Football primeiro; se vier vazio (plano Free não cobre a
// temporada), cai para TheSportsDB; por fim regenera o Contexto Jogo.
function atualizarClubesAutoGE_() {
    let fonte = "";
    // 1ª fonte: API pública da ESPN — cobre todas as competições de 2026,
    // com resultados e próximos jogos (a mesma que alimenta espn.com.br).
    try {
        const n = atualizarClubesViaESPNGE_();
        if (n > 0) fonte = "ESPN (" + n + " jogos)";
    } catch (e) {
        logGE_("CLUBES", "ESPN falhou: " + e.message);
    }
    // Fallbacks: TheSportsDB (últimos/próximos, acumulativo) e API-Football
    if (!fonte) {
        try {
            const n = atualizarClubesViaTheSportsDBGE_();
            if (n > 0) fonte = "TheSportsDB (" + n + " jogos)";
        } catch (e) {
            logGE_("CLUBES", "TheSportsDB falhou: " + e.message);
        }
    }
    if (!fonte) {
        try {
            atualizarCalendarioClubes_();
            const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_CLUBES);
            if (sh && sh.getLastRow() > 1) fonte = "API-Football";
        } catch (e) {
            logGE_("CLUBES", "API-Football falhou: " + e.message);
        }
    }
    if (fonte) {
        logGE_("CLUBES", "Calendário atualizado via " + fonte);
        try {
            const stats = gerarAnaliseContexto_();
            logGE_("CLUBES", "Contexto Jogo: " + stats.programas + " programas, " +
                stats.vinculacoes + " vínculos tema↔jogo");
        } catch (e) {
            logGE_("CLUBES", "Contexto Jogo falhou: " + e.message);
        }
        return true;
    }
    logGE_("CLUBES", "⚠️ Nenhuma fonte retornou jogos — manter atualização manual");
    return false;
}
