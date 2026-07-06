// =============================================================================
// 49_Previsao_V3_GE.js — Motor de predição V3 em SHADOW MODE (coluna PROJ V3)
// =============================================================================
// Roda em paralelo ao V2 (40) sem tocá-lo. Diferenças em relação ao V2:
//   • BASE por minuto = curva do MESMO DIA DA SEMANA, com EWMA (recência) e
//     winsorização p5–p95 (a média all-time fica "alta" em quedas estruturais
//     como a Copa; o dia-da-semana + recência corrige isso).
//   • FATOR CLUBE incremental (48) quando o tema é um clube foco.
// Mantém do V2 (validado): patamar por mediana dos ratios reais, tendência
// 3v8, decaimento AR(1) do desvio por faixa. Escreve "PROJ V3" no Histórico e
// devolve a curva para a aba Hoje espelhar (coluna V, shadow).

var PV3 = {
    FIRST_ROW: 3, LAST_ROW: 62, HEADER_ROW: 2, N: 60,
    AUD_COL: 4, TEMA_COL: 6, MEDIA_HIST_COL: 2,
    HEADER_V3: "PROJ V3",
    BETA_INICIO: 0.9478, BETA_MEIO: 0.9677, BETA_FIM: 0.9552, // mesmas do V2
    DOW_MIN_PROGRAMAS: 8,   // abaixo disso, usa a base geral (todos os dias)
    EWMA_MEIA_VIDA: 15,     // programas; peso = 0.5^(idade/meia-vida)
    WINSOR_LO: 5, WINSOR_HI: 95,
    TEMA_SHRINK_K: 30, TEMA_CAP_MIN: 0.92, TEMA_CAP_MAX: 1.08, // tema leve
    CONF_N: 8               // nº de reais p/ confiança plena no patamar
};

// -----------------------------------------------------------------------------
// Curva base por dia da semana (EWMA + winsorização), a partir da aba Base
// -----------------------------------------------------------------------------
function curvaBaseV3GE_(weekday) {
    var dias = matrizDiariaDaBaseGE_(); // [{data, minutos[60]}] ASC
    if (!dias.length) return { base: new Array(PV3.N).fill(null), n: 0, dow: false };

    var doDia = dias.filter(function (d) { return d.data.getDay() === weekday; });
    var usarDow = doDia.length >= PV3.DOW_MIN_PROGRAMAS;
    var conj = usarDow ? doDia : dias;

    var meiaVida = Number(getConfigValueGE_("EWMA_MEIA_VIDA", PV3.EWMA_MEIA_VIDA)) || PV3.EWMA_MEIA_VIDA;
    var N = conj.length;
    var base = new Array(PV3.N).fill(null);

    for (var m = 0; m < PV3.N; m++) {
        // valores desse minuto ao longo dos programas (ASC), com data p/ recência
        var vals = [];
        for (var i = 0; i < N; i++) {
            var v = conj[i].minutos[m];
            if (typeof v === "number" && isFinite(v) && v > 0) vals.push({ v: v, idx: i });
        }
        if (!vals.length) continue;
        // winsorização por minuto
        var arr = vals.map(function (x) { return x.v; }).sort(function (a, b) { return a - b; });
        var lo = percentilV3GE_(arr, PV3.WINSOR_LO), hi = percentilV3GE_(arr, PV3.WINSOR_HI);
        // EWMA: peso maior p/ programas mais recentes (idx maior)
        var somaW = 0, somaWV = 0;
        for (var k = 0; k < vals.length; k++) {
            var vv = Math.max(lo, Math.min(hi, vals[k].v));
            var idade = (N - 1) - vals[k].idx;
            var w = Math.pow(0.5, idade / meiaVida);
            somaW += w; somaWV += w * vv;
        }
        base[m] = somaW > 0 ? somaWV / somaW : null;
    }
    return { base: base, n: N, dow: usarDow };
}

// -----------------------------------------------------------------------------
// Leitura do contexto do dia (reais digitados em D, temas em F) + fator tema
// -----------------------------------------------------------------------------
function lerContextoV3GE_() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) throw Error("Histórico não encontrado");

    var n = PV3.N;
    var block = sheet.getRange(PV3.FIRST_ROW, PV3.AUD_COL, n, 3).getValues(); // D,E,F
    var fxD = sheet.getRange(PV3.FIRST_ROW, PV3.AUD_COL, n, 1).getFormulas();
    var reais = [], temas = [];
    for (var i = 0; i < n; i++) {
        var d = block[i][0], isF = fxD[i][0] !== "" && fxD[i][0] != null;
        reais.push((!isF && typeof d === "number" && d > 0) ? d : null);
        temas.push(String(block[i][2] || "").trim().toUpperCase());
    }

    var weekday = new Date().getDay();
    var curva = curvaBaseV3GE_(weekday);

    return {
        base: curva.base, dowUsado: curva.dow, nBase: curva.n,
        reais: reais, temas: temas,
        fatorTema: fatorTemaLeveV3GE_(),      // {TEMA: fator}
        fatorClube: fatoresClubesHojeGE_()    // {CLUBE: fator} (48)
    };
}

// Fator de tema leve: média global do &/MIN por tema na Base, shrunk→1, cap ±8%
function fatorTemaLeveV3GE_() {
    var out = {};
    try {
        var base = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BASE);
        if (!base || base.getLastRow() < 2) return out;
        var vals = base.getRange(2, 1, base.getLastRow() - 1, 6).getValues();
        var acc = {};
        for (var i = 0; i < vals.length; i++) {
            var t = String(vals[i][5] || "").trim().toUpperCase();
            var idx = vals[i][4];
            if (!t || t.indexOf("MERCHAN") >= 0 || typeof idx !== "number" || idx <= 0) continue;
            if (!acc[t]) acc[t] = { soma: 0, n: 0 };
            acc[t].soma += idx; acc[t].n++;
        }
        var K = PV3.TEMA_SHRINK_K;
        for (var tm in acc) {
            var media = acc[tm].soma / acc[tm].n;
            var shrunk = (acc[tm].soma + K * 1) / (acc[tm].n + K); // puxa para 1
            out[tm] = clampGE_(shrunk, PV3.TEMA_CAP_MIN, PV3.TEMA_CAP_MAX);
        }
    } catch (e) { /* neutro */ }
    return out;
}

// -----------------------------------------------------------------------------
// Motor interno V3
// -----------------------------------------------------------------------------
function betaFaixaV3GE_(m) {
    if (m <= 20) return PV3.BETA_INICIO;
    if (m <= 45) return PV3.BETA_MEIO;
    return PV3.BETA_FIM;
}

function calcularPrevisaoInternaV3GE_(ctx) {
    var base = ctx.base, reais = ctx.reais, temas = ctx.temas, n = PV3.N;

    // ratios reais/base (só onde ambos existem)
    var ratios = [], ultimoReal = -1;
    for (var m = 0; m < n; m++) {
        if (reais[m] !== null && base[m] && base[m] > 0) {
            ratios.push(reais[m] / base[m]);
            ultimoReal = m;
        } else ratios.push(null);
    }
    var ratiosValidos = ratios.filter(function (v) { return v !== null; });

    // patamar: mediana dos ratios, confiança W=min(1,n/CONF_N), blend com 1
    var patamar = 1;
    if (ratiosValidos.length) {
        var med = medianaGE_(ratiosValidos);
        var W = Math.min(1, ratiosValidos.length / PV3.CONF_N);
        patamar = W * med + (1 - W) * 1;
    }

    // tendência: média dos 3 últimos ratios / média dos 8 últimos, clamp
    var tend = 1;
    if (ratiosValidos.length >= 3) {
        var ult3 = ratiosValidos.slice(-3), ult8 = ratiosValidos.slice(-8);
        var m3 = mediaArrGE_(ult3), m8 = mediaArrGE_(ult8);
        if (m8 > 0) tend = clampGE_(m3 / m8, 0.85, 1.20);
    }

    // desvio AR(1): resíduo do último minuto real DEPOIS de patamar+tendência.
    // O patamar já captura o nível do dia; o desvio é só o ajuste fino local,
    // por isso é limitado (não pode dominar) e só se aplica aos minutos FUTUROS
    // (após o último real), com decaimento. Lacunas antes do último real usam
    // apenas base×patamar×tendência (senão a cauda baixa contamina o miolo).
    var devAtual = 0;
    if (ultimoReal >= 0 && base[ultimoReal] > 0) {
        var previstoNoUltimo = base[ultimoReal] * patamar * tend;
        if (previstoNoUltimo > 0) devAtual = clampGE_((reais[ultimoReal] / previstoNoUltimo) - 1, -0.30, 0.30);
    }

    var proj = new Array(n).fill(null);
    var dev = devAtual;
    for (var mm = 0; mm < n; mm++) {
        if (reais[mm] !== null) { proj[mm] = reais[mm]; continue; }  // real manda
        if (!base[mm] || base[mm] <= 0) { proj[mm] = null; continue; }
        var ft = ctx.fatorTema[temas[mm]] || 1;
        var fc = ctx.fatorClube[temas[mm]] || 1; // 1 se tema não é clube foco
        var ajuste;
        if (mm > ultimoReal) {                 // futuro: aplica e decai o desvio
            dev = dev * betaFaixaV3GE_(mm);
            ajuste = 1 + dev;
        } else {                               // lacuna no passado: sem desvio
            ajuste = 1;
        }
        proj[mm] = base[mm] * patamar * tend * ajuste * ft * fc;
    }

    var vFin = proj.filter(function (v) { return typeof v === "number" && isFinite(v); });
    var mediaFinal = vFin.length ? mediaArrGE_(vFin) : null;
    var motivo = "V3 · base " + (ctx.dowUsado ? "dia-da-semana" : "geral") +
        " (" + ctx.nBase + " prog.) · patamar " + patamar.toFixed(3) +
        " · tend " + tend.toFixed(3);
    return { proj: proj, mediaFinal: mediaFinal, motivo: motivo, patamar: patamar, tend: tend };
}

// -----------------------------------------------------------------------------
// Orquestração shadow: calcula e grava "PROJ V3" no Histórico
// -----------------------------------------------------------------------------
function recalcularPrevisaoV3ShadowGE_() {
    var ctx = lerContextoV3GE_();
    var res = calcularPrevisaoInternaV3GE_(ctx);
    gravarPrevisaoV3HistoricoGE_(res.proj, res.motivo);
    return res; // {proj, mediaFinal, motivo, ...}
}

function gravarPrevisaoV3HistoricoGE_(proj, motivo) {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    var col = acharColunaHeaderV3GE_(sheet, PV3.HEADER_V3);
    if (col < 0) {
        col = sheet.getLastColumn() + 1;
        sheet.getRange(PV3.HEADER_ROW, col).setValue(PV3.HEADER_V3)
            .setFontWeight("bold").setBackground("#7D3C98").setFontColor("#FFFFFF");
    }
    var out = proj.map(function (v) { return [typeof v === "number" && isFinite(v) ? v : ""]; });
    sheet.getRange(PV3.FIRST_ROW, col, PV3.N, 1).setValues(out);
    sheet.getRange(PV3.HEADER_ROW, col).setNote(motivo || "");
    sheet.getRange(PV3.FIRST_ROW, col, PV3.N, 1).setNumberFormat("0.00");
}

function acharColunaHeaderV3GE_(sheet, header) {
    var lastCol = sheet.getLastColumn();
    var hs = sheet.getRange(PV3.HEADER_ROW, 1, 1, lastCol).getValues()[0];
    for (var c = 0; c < hs.length; c++)
        if (((hs[c] || "") + "").trim() === header) return c + 1;
    return -1;
}

// -----------------------------------------------------------------------------
// Comparação diária ex-post (dia fechado) — placar TY vs TZ na aba Backtest
// -----------------------------------------------------------------------------
function registrarShadowLogV3GE_() {
    var ss = SpreadsheetApp.getActive();
    var hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!hist) return;

    var n = PV3.N;
    var colTY = acharColunaHeaderV3GE_(hist, PV2 && PV2.HEADER_V2 ? PV2.HEADER_V2 : "PROJ V2");
    var colTZ = acharColunaHeaderV3GE_(hist, PV3.HEADER_V3);
    if (colTZ < 0) return;

    var block = hist.getRange(PV3.FIRST_ROW, PV3.AUD_COL, n, 1).getValues();      // D
    var fxD = hist.getRange(PV3.FIRST_ROW, PV3.AUD_COL, n, 1).getFormulas();
    var ty = colTY > 0 ? hist.getRange(PV3.FIRST_ROW, colTY, n, 1).getValues() : null;
    var tz = hist.getRange(PV3.FIRST_ROW, colTZ, n, 1).getValues();

    var errTY = [], errTZ = [], somaReal = 0, nReal = 0, somaTY = 0, somaTZ = 0;
    for (var i = 0; i < n; i++) {
        var d = block[i][0], isF = fxD[i][0] !== "" && fxD[i][0] != null;
        var real = (!isF && typeof d === "number" && d > 0) ? d : null;
        if (real === null) continue;
        nReal++; somaReal += real;
        if (ty && typeof ty[i][0] === "number") { errTY.push(Math.abs(ty[i][0] - real)); somaTY += ty[i][0]; }
        if (typeof tz[i][0] === "number") { errTZ.push(Math.abs(tz[i][0] - real)); somaTZ += tz[i][0]; }
    }
    if (nReal < 20) { logGE_("SHADOW", "Dia com poucos reais (" + nReal + ") — placar não registrado"); return; }

    var maeTY = errTY.length ? mediaArrGE_(errTY) : "";
    var maeTZ = errTZ.length ? mediaArrGE_(errTZ) : "";
    var mediaReal = somaReal / nReal;
    var errFinalTY = errTY.length ? Math.abs(somaTY / errTY.length - mediaReal) : "";
    var errFinalTZ = errTZ.length ? Math.abs(somaTZ / errTZ.length - mediaReal) : "";
    var vencedor = (typeof maeTY === "number" && typeof maeTZ === "number")
        ? (maeTZ < maeTY ? "V3" : (maeTZ > maeTY ? "V2" : "empate")) : "—";

    var sh = ensureAbaGE_(CFG.SHEET_BACKTEST + "_Shadow",
        ["DATA", "N REAIS", "MAE V2", "MAE V3", "ERRO FINAL V2", "ERRO FINAL V3", "VENCEDOR DIA"]);
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, 7).setValues([[
        Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy"),
        nReal, maeTY, maeTZ, errFinalTY, errFinalTZ, vencedor
    ]]);
    logGE_("SHADOW", "Placar do dia: V2 " + (typeof maeTY === "number" ? maeTY.toFixed(4) : "—") +
        " × V3 " + (typeof maeTZ === "number" ? maeTZ.toFixed(4) : "—") + " → " + vencedor);
}

function percentilV3GE_(sortedArr, p) {
    if (!sortedArr.length) return NaN;
    var idx = (p / 100) * (sortedArr.length - 1);
    var lo = Math.floor(idx), hi = Math.ceil(idx);
    return sortedArr[lo] + (sortedArr[hi] - sortedArr[lo]) * (idx - lo);
}
