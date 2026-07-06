// =============================================================================
// 48_Fatores_GE.js — Fatores aprendidos dos dados (recalculados às 06:00)
// =============================================================================
// Tabelas pequenas, persistidas na aba oculta "Fatores", lidas pelo motor V3
// (49) e pela camada de tema (Fase 3). Tudo com shrinkage e backoff por
// amostra mínima — célula com pouca amostra recua para um nível mais geral.
//
// FATOR CLUBE (Fase 1): mede o EFEITO INCREMENTAL do resultado/recência/clássico
// do último jogo, NÃO a força do clube como tema (essa já está no fator da aba
// Tema Minuto). Por isso é sempre relativo à média do próprio clube:
//     fator = shrink(média_da_célula → média_do_clube) / média_do_clube
// Nível 1 (clube em geral) = 1,0 por definição (sem info de resultado, sem
// efeito incremental). Assim não há dupla contagem com o fator de tema.

var FATORES_GE = {
    SHEET: "Fatores",
    CLUBE_SHRINK_K: 20,     // minutos de amostra p/ puxar a célula ao clube
    CLUBE_CAP_MIN: 0.90,
    CLUBE_CAP_MAX: 1.10,
    // amostra mínima por nível de backoff (config sobrepõe)
    CLUBE_N_MINIMO: 40,
    BLOCO_CLUBE: "CLUBE"    // rótulo do bloco na aba Fatores
};

// Cache por execução (evita reler a aba a cada minuto no ao vivo)
var _cacheFatorClube = null;

// -----------------------------------------------------------------------------
// Recálculo mestre (rotina 06:00 e menu)
// -----------------------------------------------------------------------------
function recalcularFatoresGE() {
    _cacheFatorClube = null;
    _cacheFatorFaixa = null;
    var passos = [];
    try { passos.push("clube: " + recalcularFatorClubeGE_() + " células"); }
    catch (e) { passos.push("clube⚠️" + e.message); }
    try { passos.push("faixa: " + recalcularFatorFaixaGE_() + " células"); }
    catch (e) { passos.push("faixa⚠️" + e.message); }
    logGE_("FATORES", passos.join(" | "));
    return passos;
}

function recalcularFatoresMenuGE() {
    var passos = recalcularFatoresGE();
    SpreadsheetApp.getUi().alert("🔧 Fatores recalculados:\n\n• " + passos.join("\n• "));
}

// -----------------------------------------------------------------------------
// FATOR CLUBE — aprendizado
// -----------------------------------------------------------------------------
function recalcularFatorClubeGE_() {
    var ss = SpreadsheetApp.getActive();
    var base = ss.getSheetByName(CFG.SHEET_BASE);
    var clubesSh = ss.getSheetByName(CFG.SHEET_CLUBES);
    if (!base || base.getLastRow() < 2) throw Error("Base vazia");
    if (!clubesSh || clubesSh.getLastRow() < 2) throw Error("Aba Clubes vazia");

    // Jogos realizados por clube, ordenados por data ASC
    var jogos = {};
    CFG.CLUBES_FOCO.forEach(function (c) { jogos[c] = []; });
    var cv = clubesSh.getRange(2, 1, clubesSh.getLastRow() - 1, 13).getValues();
    for (var i = 0; i < cv.length; i++) {
        var clube = String(cv[i][1] || "").trim().toUpperCase();
        if (!jogos[clube]) continue;
        var data = cv[i][0] instanceof Date ? cv[i][0] : parseDateStr_(String(cv[i][0] || ""));
        if (!data) continue;
        if (String(cv[i][10] || "").toUpperCase() !== "REALIZADO") continue;
        jogos[clube].push({
            data: data,
            resultado: String(cv[i][9] || "").trim().toUpperCase().charAt(0), // V/E/D
            classico: ehClassicoPaulistaGE_(clube, String(cv[i][2] || "")),
            importancia: Number(cv[i][11]) || 0
        });
    }
    for (var c in jogos) jogos[c].sort(function (a, b) { return a.data - b.data; });

    // Base: minutos onde tema é clube foco → &/MIN condicionado ao último jogo
    var bv = base.getRange(2, 1, base.getLastRow() - 1, 9).getValues();
    // acumuladores: cells[chave] = {soma, n}; clube geral em cellClube[clube]
    var cells = {}, cellClube = {};
    CFG.CLUBES_FOCO.forEach(function (c) { cellClube[c] = { soma: 0, n: 0 }; });

    function acumula(chave, idx) {
        if (!cells[chave]) cells[chave] = { soma: 0, n: 0 };
        cells[chave].soma += idx; cells[chave].n++;
    }

    for (var r = 0; r < bv.length; r++) {
        var tema = String(bv[r][5] || "").trim().toUpperCase();
        if (!cellClube[tema]) continue;          // só clubes foco
        var idx = bv[r][4];                      // coluna &/MIN
        if (typeof idx !== "number" || !isFinite(idx) || idx <= 0) continue;
        var dataProg = bv[r][0] instanceof Date ? bv[r][0] : parseDateStr_(String(bv[r][0] || ""));
        if (!dataProg) continue;

        // último jogo do clube ANTES do programa
        var ult = ultimoJogoAntesGE_(jogos[tema], dataProg);
        if (!ult) continue;

        cellClube[tema].soma += idx; cellClube[tema].n++;
        var rec = bucketRecenciaGE_(ult.diasAntes);
        var big = (ult.classico || ult.importancia >= 4) ? "BIG" : "COM";
        // acumula nos 3 níveis condicionais (nível 1 = clube geral já em cellClube)
        acumula("2|" + tema + "|" + ult.resultado, idx);
        acumula("3|" + tema + "|" + ult.resultado + "|" + rec, idx);
        acumula("4|" + tema + "|" + ult.resultado + "|" + rec + "|" + big, idx);
    }

    // Monta linhas fator = clamp(shrink(cellMean→clubMean)/clubMean)
    var K = FATORES_GE.CLUBE_SHRINK_K;
    var linhas = [];
    for (var ch in cells) {
        var clube = ch.split("|")[1];
        var cm = cellClube[clube];
        if (!cm || cm.n === 0) continue;
        var clubMean = cm.soma / cm.n;
        if (clubMean <= 0) continue;
        var cell = cells[ch];
        var shrunk = (cell.soma + K * clubMean) / (cell.n + K);
        var fator = clampGE_(shrunk / clubMean, FATORES_GE.CLUBE_CAP_MIN, FATORES_GE.CLUBE_CAP_MAX);
        linhas.push([FATORES_GE.BLOCO_CLUBE, ch, cell.n, fator]);
    }

    escreverBlocoFatoresGE_(FATORES_GE.BLOCO_CLUBE, ["BLOCO", "CHAVE", "N", "FATOR"], linhas);
    return linhas.length;
}

function ultimoJogoAntesGE_(lista, dataProg) {
    if (!lista || !lista.length) return null;
    var alvo = new Date(dataProg.getFullYear(), dataProg.getMonth(), dataProg.getDate());
    for (var j = lista.length - 1; j >= 0; j--) {
        var d = new Date(lista[j].data.getFullYear(), lista[j].data.getMonth(), lista[j].data.getDate());
        if (d < alvo) {
            return {
                resultado: lista[j].resultado,
                classico: lista[j].classico,
                importancia: lista[j].importancia,
                diasAntes: Math.round((alvo - d) / 86400000)
            };
        }
    }
    return null;
}

function bucketRecenciaGE_(dias) {
    if (dias <= 1) return "R01";
    if (dias <= 3) return "R23";
    return "R4+";
}

// -----------------------------------------------------------------------------
// FATOR CLUBE — leitura (backoff por amostra mínima)
// -----------------------------------------------------------------------------
function carregarFatorClubeTabelaGE_() {
    if (_cacheFatorClube) return _cacheFatorClube;
    var tab = { "2": {}, "3": {}, "4": {} };
    var linhas = lerBlocoFatoresGE_(FATORES_GE.BLOCO_CLUBE);
    for (var i = 0; i < linhas.length; i++) {
        var chave = String(linhas[i][1] || "");
        var nivel = chave.split("|")[0];
        if (!tab[nivel]) continue;
        tab[nivel][chave] = { n: Number(linhas[i][2]) || 0, fator: Number(linhas[i][3]) || 1 };
    }
    _cacheFatorClube = tab;
    return tab;
}

// Backoff: nível 4 → 3 → 2 → 1,0 (só usa célula com n >= mínimo)
function fatorClubeCondicionalGE_(clube, resultado, recencia, bigmatch) {
    var tab = carregarFatorClubeTabelaGE_();
    var Nmin = Number(getConfigValueGE_("FATOR_CLUBE_N_MINIMO", FATORES_GE.CLUBE_N_MINIMO)) ||
        FATORES_GE.CLUBE_N_MINIMO;
    var big = bigmatch ? "BIG" : "COM";
    var tentativas = [
        ["4", "4|" + clube + "|" + resultado + "|" + recencia + "|" + big],
        ["3", "3|" + clube + "|" + resultado + "|" + recencia],
        ["2", "2|" + clube + "|" + resultado]
    ];
    for (var t = 0; t < tentativas.length; t++) {
        var cel = tab[tentativas[t][0]][tentativas[t][1]];
        if (cel && cel.n >= Nmin) return cel.fator;
    }
    return 1.0;
}

// Fator do dia para cada clube foco (a partir do contexto de hoje) →
// { CORINTHIANS: 1.06, ... }. Usado pelo motor V3 e pela camada de tema.
function fatoresClubesHojeGE_() {
    var out = {};
    try {
        var ctx = getContextoClubesHojeGE_();
        for (var clube in ctx) {
            var c = ctx[clube];
            if (!c || !c.resultado) { out[clube] = 1.0; continue; }
            var rec = c.quando === "hoje" || c.quando === "ontem" ? "R01"
                : /^([23])d$/.test(c.quando) ? "R23" : "R4+";
            var big = c.classico || c.decisivo;
            out[clube] = fatorClubeCondicionalGE_(clube, c.resultado, rec, big);
        }
    } catch (e) {
        logGE_("FATORES", "fatoresClubesHojeGE_ falhou: " + e.message);
    }
    return out;
}

// -----------------------------------------------------------------------------
// Persistência na aba "Fatores" (blocos rotulados, empilhados)
// -----------------------------------------------------------------------------
function getFatoresSheetGE_() {
    return ensureAbaGE_(FATORES_GE.SHEET, ["BLOCO", "CHAVE", "N", "FATOR"]);
}

// Substitui todas as linhas de um bloco preservando os demais blocos
function escreverBlocoFatoresGE_(bloco, headers, linhas) {
    var sh = getFatoresSheetGE_();
    var last = sh.getLastRow();
    var mantidas = [];
    if (last > 1) {
        var todas = sh.getRange(2, 1, last - 1, headers.length).getValues();
        for (var i = 0; i < todas.length; i++) {
            if (String(todas[i][0] || "") !== bloco) mantidas.push(todas[i]);
        }
    }
    var saida = mantidas.concat(linhas);
    sh.getRange(2, 1, Math.max(last - 1, saida.length, 1), headers.length)
        .clearContent();
    if (saida.length) sh.getRange(2, 1, saida.length, headers.length).setValues(saida);
}

function lerBlocoFatoresGE_(bloco) {
    var sh = SpreadsheetApp.getActive().getSheetByName(FATORES_GE.SHEET);
    if (!sh || sh.getLastRow() < 2) return [];
    var todas = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    var out = [];
    for (var i = 0; i < todas.length; i++) {
        if (String(todas[i][0] || "") === bloco) out.push(todas[i]);
    }
    return out;
}

// -----------------------------------------------------------------------------
// FATOR FAIXA — rendimento do tema por terço do programa (abertura/miolo/fecho)
// -----------------------------------------------------------------------------
// Alguns temas rendem mais na abertura, outros no fechamento. Fator relativo
// à média do próprio tema (não duplica a força global): fator = shrink(célula
// → média_do_tema) / média_do_tema, cap ±10%. Terços: 0–19 / 20–45 / 46–59.
var FATORES_FAIXA = { BLOCO: "FAIXA", SHRINK_K: 15, CAP_MIN: 0.90, CAP_MAX: 1.10 };
var _cacheFatorFaixa = null;

function tercoFaixaGE_(minuto) { return minuto <= 19 ? "AB" : (minuto <= 45 ? "MI" : "FE"); }

function recalcularFatorFaixaGE_() {
    var base = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BASE);
    if (!base || base.getLastRow() < 2) throw Error("Base vazia");
    var bv = base.getRange(2, 1, base.getLastRow() - 1, 6).getValues();
    var cellTema = {}, cellFaixa = {};
    for (var r = 0; r < bv.length; r++) {
        var tema = String(bv[r][5] || "").trim().toUpperCase();
        if (!tema || tema.indexOf("MERCHAN") >= 0) continue;
        var idx = bv[r][4];
        if (typeof idx !== "number" || !isFinite(idx) || idx <= 0) continue;
        var minuto = minutoBaseGE_(bv[r][2]);
        if (minuto === null) continue;
        if (!cellTema[tema]) cellTema[tema] = { soma: 0, n: 0 };
        cellTema[tema].soma += idx; cellTema[tema].n++;
        var chave = tema + "|" + tercoFaixaGE_(minuto);
        if (!cellFaixa[chave]) cellFaixa[chave] = { soma: 0, n: 0 };
        cellFaixa[chave].soma += idx; cellFaixa[chave].n++;
    }
    var K = FATORES_FAIXA.SHRINK_K, linhas = [];
    for (var ch in cellFaixa) {
        var tema = ch.split("|")[0], tm = cellTema[tema];
        if (!tm || tm.n < 10) continue; // tema com amostra mínima
        var temaMean = tm.soma / tm.n;
        if (temaMean <= 0) continue;
        var cell = cellFaixa[ch];
        var shrunk = (cell.soma + K * temaMean) / (cell.n + K);
        var fator = clampGE_(shrunk / temaMean, FATORES_FAIXA.CAP_MIN, FATORES_FAIXA.CAP_MAX);
        linhas.push([FATORES_FAIXA.BLOCO, ch, cell.n, fator]);
    }
    escreverBlocoFatoresGE_(FATORES_FAIXA.BLOCO, ["BLOCO", "CHAVE", "N", "FATOR"], linhas);
    return linhas.length;
}

function carregarFatorFaixaTabelaGE_() {
    if (_cacheFatorFaixa) return _cacheFatorFaixa;
    var tab = {};
    var linhas = lerBlocoFatoresGE_(FATORES_FAIXA.BLOCO);
    for (var i = 0; i < linhas.length; i++) tab[String(linhas[i][1])] = Number(linhas[i][3]) || 1;
    _cacheFatorFaixa = tab;
    return tab;
}

function fatorFaixaGE_(tema, minuto) {
    var tab = carregarFatorFaixaTabelaGE_();
    var chave = String(tema || "").toUpperCase() + "|" + tercoFaixaGE_(minuto);
    return tab[chave] || 1;
}
