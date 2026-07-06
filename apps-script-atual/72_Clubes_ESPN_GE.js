// =============================================================================
// 72_Clubes_ESPN_GE.js — Calendário completo dos clubes via API pública da ESPN
// =============================================================================
// As páginas espn.com.br/futebol/time/{resultados|calendario}/_/id/<id>/...
// consomem esta API JSON (site.api.espn.com). Cobre Série A, Copa do Brasil,
// Libertadores, Sul-Americana e Paulistão, com resultados (fixture=false) e
// próximos jogos (fixture=true). É a fonte primária da aba Clubes.

var ESPN_GE = {
    BASE: "https://site.api.espn.com/apis/site/v2/sports/soccer",
    TEAMS: {
        "CORINTHIANS": "874",
        "PALMEIRAS": "2029",
        "SÃO PAULO": "2026",
        "SANTOS": "2674"
    },
    LIGAS: {
        "bra.1": { nome: "Brasileirão Série A", imp: 2 },
        "bra.copa_do_brazil": { nome: "Copa do Brasil", imp: 3 },
        "conmebol.libertadores": { nome: "Libertadores", imp: 4 },
        "conmebol.sudamericana": { nome: "Sul-Americana", imp: 3 },
        "bra.camp.paulista": { nome: "Paulistão", imp: 2 }
    }
};

function espnFetchGE_(url) {
    try {
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (resp.getResponseCode() !== 200) return null;
        return JSON.parse(resp.getContentText());
    } catch (e) {
        return null;
    }
}

// Converte um evento da ESPN em linha da aba Clubes
function espnEventoParaLinhaGE_(ev, clube, teamId, ligaNome, ligaImp) {
    if (!ev || !ev.date) return null;
    const data = new Date(ev.date);
    if (isNaN(data.getTime())) return null;

    const comp = (ev.competitions && ev.competitions[0]) || {};
    const competidores = comp.competitors || [];
    let eu = null, ele = null;
    for (const c of competidores) {
        if (String((c.team || {}).id) === teamId) eu = c;
        else ele = c;
    }
    if (!eu || !ele) return null;

    const adversario = (ele.team && (ele.team.displayName || ele.team.name)) || "";
    const mando = eu.homeAway === "home" ? "CASA" : "FORA";
    const estado = String(((comp.status || {}).type || {}).state || "");

    const scoreDe = c => {
        const s = c.score;
        if (s === null || s === undefined) return null;
        const v = typeof s === "object" ? s.value : s;
        const n = Number(v);
        return isFinite(n) ? Math.round(n) : null;
    };
    const gp = scoreDe(eu), gc = scoreDe(ele);
    const realizado = estado === "post" && gp !== null && gc !== null;

    let placar = "", resultado = "";
    if (realizado) {
        placar = gp + "×" + gc;
        resultado = gp > gc ? "V" : (gp < gc ? "D" : "E");
    }
    const status = realizado ? "REALIZADO" : (estado === "in" ? "AO VIVO" : "AGENDADO");

    let importancia = ligaImp;
    if (ehClassicoPaulistaGE_(clube, adversario)) importancia = Math.min(5, importancia + 1);

    const umDia = 24 * 3600 * 1000;
    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    const dataDia = new Date(data); dataDia.setHours(0, 0, 0, 0);
    const diasPGe = Math.round((dataDia - hoje) / umDia);

    return [Utilities.formatDate(data, CFG.TIMEZONE, "dd/MM/yyyy"), clube, adversario,
        ligaNome, "", mando, placar,
        realizado ? gp : "", realizado ? gc : "",
        resultado, status, importancia, diasPGe];
}

// Busca todos os jogos (resultados + próximos) dos 4 clubes em todas as ligas
function atualizarClubesViaESPNGE_() {
    const temporada = Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyy");
    const linhas = [];
    const vistos = {};

    for (const clube in ESPN_GE.TEAMS) {
        const teamId = ESPN_GE.TEAMS[clube];
        for (const slug in ESPN_GE.LIGAS) {
            const liga = ESPN_GE.LIGAS[slug];
            for (const fixture of ["false", "true"]) {
                const json = espnFetchGE_(ESPN_GE.BASE + "/" + slug + "/teams/" + teamId +
                    "/schedule?season=" + temporada + "&fixture=" + fixture);
                const eventos = (json && json.events) || [];
                for (const ev of eventos) {
                    const chave = clube + "|" + (ev.id || ev.date);
                    if (vistos[chave]) continue;
                    vistos[chave] = true;
                    const linha = espnEventoParaLinhaGE_(ev, clube, teamId, liga.nome, liga.imp);
                    if (linha) linhas.push(linha);
                }
                Utilities.sleep(120);
            }
        }
    }

    if (!linhas.length) return 0;
    return mesclarEEscreverClubesGE_(linhas, "ESPN");
}

// Mescla com o que já existe na aba (chave data|clube|adversário; dados novos
// prevalecem — placares confirmados substituem agendados) e escreve tudo.
function mesclarEEscreverClubesGE_(linhas, fonte) {
    const shExist = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_CLUBES);
    const mapa = {};
    if (shExist && shExist.getLastRow() > 1) {
        const vals = shExist.getRange(2, 1, shExist.getLastRow() - 1, 13).getValues();
        for (const r of vals) {
            const dataStr = r[0] instanceof Date
                ? Utilities.formatDate(r[0], CFG.TIMEZONE, "dd/MM/yyyy")
                : String(r[0] || "").trim();
            if (!dataStr || !parseDateStr_(dataStr)) continue;
            r[0] = dataStr;
            mapa[chaveJogoGE_(dataStr, r[1], r[2])] = r.slice(0, 13);
        }
    }
    for (const l of linhas) {
        mapa[chaveJogoGE_(l[0], l[1], l[2])] = l;
    }
    const todas = Object.values(mapa);
    todas.sort((a, b) => parseDateStr_(b[0]) - parseDateStr_(a[0]));
    escreverAbaClubesGE_(todas, fonte + " (" + todas.length + " jogos)");
    return todas.length;
}

// Chave normalizada (adversário sem acento/caixa para casar fontes distintas)
function chaveJogoGE_(dataStr, clube, adversario) {
    const norm = s => String(s || "").toUpperCase()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    return dataStr + "|" + norm(clube) + "|" + norm(adversario);
}
