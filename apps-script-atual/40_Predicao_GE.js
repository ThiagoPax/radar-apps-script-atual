// =============================================================================
// 20_PrevisaoV2_GE.gs — Previsão de audiência V2/V6 (motor interno robusto + tema por minuto +
//                       motor externo OPCIONAL via API Python)
// =============================================================================
//
// DECISÃO (backtest walk-forward, 68 dias de teste, 128 programas):
//   O modelo estatístico robusto (patamar por MEDIANA + tendência local +
//   tema shrinkado com cap) SUPERA o modelo atual da planilha a partir de
//   18:15 e EMPATA/SUPERA LightGBM/skforecast/mlforecast isolados. O ensemble
//   com ML melhora ~3-5% marginais — disponível via API externa OPCIONAL.
//   Recomendação: motor INTERNO como padrão; API externa só se quiser o extra.
//
// LICENÇAS (verificadas): mlforecast Apache-2.0 | skforecast BSD-3 |
//   LightGBM MIT | XGBoost Apache-2.0 | CatBoost Apache-2.0 | scikit-learn
//   BSD-3 | FastAPI MIT. Todas gratuitas p/ uso comercial. ATENÇÃO: NÃO
//   confundir a biblioteca open-source "mlforecast" com "TimeGPT/Nixtla
//   Cloud" (API paga). Nada aqui usa serviço pago.
//
// ARQUITETURA:
//   • Coluna auxiliar "PROJ V2" (valores, recalculados por script).
//   • Coluna D futura: =SE(D{r-1}="";"";SE(V2<>"";V2;PROJ)) — fallback
//     automático para a projeção antiga se V2 estiver vazia. Reais digitados
//     sobrescrevem a fórmula e são SEMPRE preservados.
//   • Gatilho INSTALÁVEL onEdit (tem autorização p/ UrlFetchApp) recalcula
//     ao editar D ou F. Modo externo: tenta API (timeout 8s) e cai no
//     interno em qualquer falha. Nenhuma dependência circular: V2 são
//     números estáticos.
//
// FÓRMULAS gravadas com detector automático de sintaxe local
// (padrão do projeto). Requer 01_Utils.gs e 00_Config.gs.
// =============================================================================

var PV2 = {
    FIRST_ROW: 3,
    LAST_ROW: 62,
    HEADER_ROW: 2,
    AUD_COL: 4,
    TEMA_COL: 6,
    HEADER_V2: "PROJ V2",
    HEADER_PROJ_ANTIGA: "PROJ",
    TEMA_SHEET: "Tema Minuto",
    // Persistência do desvio (AR(1), medida no histórico — mesmas do 18)
    BETA_INICIO: 0.9478, // 18:00-18:20
    BETA_MEIO: 0.9677,   // 18:21-18:45
    BETA_FIM: 0.9552,    // 18:46-18:59
    // Tema: shrinkage e cap (tema ajuda, não domina)
    TEMA_SHRINK_K: 30,
    TEMA_CAP_MIN: 0.88,
    TEMA_CAP_MAX: 1.12,
    // V6: o tema passa a entrar com mais força na TY, mas ainda com freio.
    // A V2 combina fator global do tema + fator específico daquele minuto.
    // Depois aplica ganho moderado e cap para o tema não dominar o patamar real da coluna D.
    TEMA_MINUTO_PESO: 0.70,
    TEMA_GLOBAL_PESO: 0.30,
    TEMA_GAIN_TY: 1.65,
    TEMA_TY_CAP_MIN: 0.82,
    TEMA_TY_CAP_MAX: 1.20,
    // V6: além do fator estatístico, usa a projeção antiga (UD/PROJ) como sinal auxiliar do tema.
    // Isso deixa a TY reagir quando F muda, mas só parcialmente e com cap.
    TEMA_PLANILHA_BLEND: 0.70,
    TEMA_PLANILHA_CAP_MIN: 0.90,
    TEMA_PLANILHA_CAP_MAX: 1.18,
    // Parte do impacto do tema do minuto carrega para os próximos minutos com beta.
    TEMA_CARRY_PESO: 0.25,
    TEMA_MINUTO_MIN_AMOSTRA: 3,
    TEMA_MINUTO_SHRINK_K: 8,
    // Patamar do dia: confiança plena a partir de N reais
    PATAMAR_N_PLENO: 8,
    // Tendência local: clamp
    TREND_MIN: 0.94,
    TREND_MAX: 1.06,
    // Backtest
    CHECKPOINTS: [5, 10, 15, 20, 30, 45],
    BACKTEST_SHEET: "Backtest",
    MIN_TREINO_BACKTEST: 60,
    // Script Properties (API externa)
    PROP_URL: "PV2_API_URL",
    PROP_SECRET: "PV2_API_SECRET",
    PROP_ATIVA: "PV2_EXTERNA_ATIVA",
    API_TIMEOUT_MS: 8000,
    THROTTLE_S: 3,
    // V6: espera fórmulas UA/UB/UD recalcularem após editar F antes de ler a planilha.
    // 800ms + flush() é suficiente; o valor antigo (2200) serializava o tick e
    // derrubava o webhook de audiência por timeout no cliente.
    WAIT_FORMULAS_MS: 800
};

function betaPV2_(m) { // m = 0..59
    var row = m + PV2.FIRST_ROW;
    if (row <= 23) return PV2.BETA_INICIO;
    if (row <= 48) return PV2.BETA_MEIO;
    return PV2.BETA_FIM;
}


// =============================================================================
// FORMULA WRITER LOCALIZADO
// Evita erro de análise em planilhas que usam função em inglês + separador ;
// =============================================================================
function batchSetFormulasSparsePV2Locale_(sheet, cells) {
    if (!cells || !cells.length) return;
    var dialect = pv2DetectFormulaDialect_();
    for (var i = 0; i < cells.length; i++) {
        var f = pv2LocalizeFormula_(cells[i].formula, dialect);
        sheet.getRange(cells[i].row, cells[i].col).setFormula(f);
    }
}

function pv2DetectFormulaDialect_() {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("PV2_FORMULA_DIALECT");
    if (cached) return cached;

    var ss = SpreadsheetApp.getActive();
    var name = "__PV2_FORMULA_TEST__";
    var old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old);
    var sh = ss.insertSheet(name);
    sh.hideSheet();

    var tests = [
        { name: "EN_COMMA", formula: '=IF(1=1,"OK","NO")' },
        { name: "EN_SEMI",  formula: '=IF(1=1;"OK";"NO")' },
        { name: "PT_SEMI",  formula: '=SE(1=1;"OK";"NO")' },
        { name: "PT_COMMA", formula: '=SE(1=1,"OK","NO")' }
    ];

    var out = "EN_SEMI";
    try {
        for (var i = 0; i < tests.length; i++) {
            try {
                sh.getRange("A1").clearContent();
                sh.getRange("A1").setFormula(tests[i].formula);
                SpreadsheetApp.flush();
                if (sh.getRange("A1").getDisplayValue() === "OK" || sh.getRange("A1").getValue() === "OK") {
                    out = tests[i].name;
                    break;
                }
            } catch (e) {}
        }
    } finally {
        ss.deleteSheet(sh);
    }

    cache.put("PV2_FORMULA_DIALECT", out, 21600);
    return out;
}

function pv2LocalizeFormula_(formula, dialect) {
    var f = (formula || "").toString();
    if (!f) return f;
    if (f.charAt(0) !== "=") f = "=" + f;
    f = pv2PtToEnFunctions_(f);

    if (pv2HasOutsideQuotes_(f, ";")) {
        f = pv2OutsideQuotes_(f, function(part) { return part.replace(/(\d),(\d)/g, "$1.$2"); });
        f = pv2SwapOutsideQuotes_(f, ";", ",");
    }

    if (dialect === "EN_COMMA") return f;

    if (dialect === "EN_SEMI") {
        f = pv2SwapOutsideQuotes_(f, ",", ";");
        f = pv2OutsideQuotes_(f, function(part) { return part.replace(/(\d)\.(\d)/g, "$1,$2"); });
        return f;
    }

    if (dialect === "PT_SEMI") {
        f = pv2EnToPtFunctions_(f);
        f = pv2SwapOutsideQuotes_(f, ",", ";");
        f = pv2OutsideQuotes_(f, function(part) { return part.replace(/(\d)\.(\d)/g, "$1,$2"); });
        return f;
    }

    if (dialect === "PT_COMMA") return pv2EnToPtFunctions_(f);
    return f;
}

function pv2PtToEnFunctions_(f) {
    return f
        .replace(/\bSEERRO\s*\(/gi, "IFERROR(")
        .replace(/\bSOMASE\s*\(/gi, "SUMIF(")
        .replace(/\bCONT\.SE\s*\(/gi, "COUNTIF(")
        .replace(/\bMÉDIASE\s*\(/gi, "AVERAGEIF(")
        .replace(/\bMEDIASE\s*\(/gi, "AVERAGEIF(")
        .replace(/ÍNDICE\s*\(/gi, "INDEX(")   // sem \b: Í não é word char em regex
        .replace(/\bINDICE\s*\(/gi, "INDEX(")
        .replace(/\bDESLOC\s*\(/gi, "OFFSET(")
        .replace(/\bMÉDIA\s*\(/gi, "AVERAGE(")
        .replace(/\bMEDIA\s*\(/gi, "AVERAGE(")
        .replace(/\bMED\s*\(/gi, "MEDIAN(")
        .replace(/\bSOMA\s*\(/gi, "SUM(")
        .replace(/\bPROCV\s*\(/gi, "VLOOKUP(")
        .replace(/\bARRED\s*\(/gi, "ROUND(")
        .replace(/\bMÁXIMO\s*\(/gi, "MAX(")
        .replace(/\bMAXIMO\s*\(/gi, "MAX(")
        .replace(/\bMÍNIMO\s*\(/gi, "MIN(")
        .replace(/\bMINIMO\s*\(/gi, "MIN(")
        .replace(/\bLIN\s*\(\s*\)/gi, "ROW()")
        .replace(/\bCOL\s*\(/gi, "COLUMN(")
        .replace(/\bOU\s*\(/gi, "OR(")
        .replace(/\bE\s*\(/gi, "AND(")
        .replace(/\bSE\s*\(/gi, "IF(");
}

function pv2EnToPtFunctions_(f) {
    return f
        .replace(/\bIFERROR\s*\(/gi, "SEERRO(")
        .replace(/\bAVERAGEIF\s*\(/gi, "MÉDIASE(")
        .replace(/\bSUMIF\s*\(/gi, "SOMASE(")
        .replace(/\bCOUNTIF\s*\(/gi, "CONT.SE(")
        .replace(/\bVLOOKUP\s*\(/gi, "PROCV(")
        .replace(/\bINDEX\s*\(/gi, "ÍNDICE(")
        .replace(/\bOFFSET\s*\(/gi, "DESLOC(")
        .replace(/\bAVERAGE\s*\(/gi, "MÉDIA(")
        .replace(/\bMEDIAN\s*\(/gi, "MED(")
        .replace(/\bROUND\s*\(/gi, "ARRED(")
        .replace(/\bSUM\s*\(/gi, "SOMA(")
        .replace(/\bMAX\s*\(/gi, "MÁXIMO(")
        .replace(/\bMIN\s*\(/gi, "MÍNIMO(")
        .replace(/\bROW\s*\(\s*\)/gi, "LIN()")
        .replace(/\bCOLUMN\s*\(/gi, "COL(")
        .replace(/\bAND\s*\(/gi, "E(")
        .replace(/\bOR\s*\(/gi, "OU(")
        .replace(/\bIF\s*\(/gi, "SE(");
}

function pv2SwapOutsideQuotes_(s, fromChar, toChar) {
    var out = "", inStr = false;
    for (var i = 0; i < s.length; i++) {
        var ch = s.charAt(i);
        if (ch === '"') {
            out += ch;
            if (s.charAt(i + 1) === '"') { out += s.charAt(i + 1); i++; }
            else inStr = !inStr;
        } else if (!inStr && ch === fromChar) out += toChar;
        else out += ch;
    }
    return out;
}

function pv2OutsideQuotes_(s, fn) {
    var out = "", buf = "", inStr = false;
    for (var i = 0; i < s.length; i++) {
        var ch = s.charAt(i);
        if (ch === '"') {
            if (!inStr) { out += fn(buf); buf = ""; inStr = true; out += ch; }
            else {
                out += ch;
                if (s.charAt(i + 1) === '"') { out += s.charAt(i + 1); i++; }
                else inStr = false;
            }
        } else {
            if (inStr) out += ch;
            else buf += ch;
        }
    }
    if (buf) out += fn(buf);
    return out;
}

function pv2HasOutsideQuotes_(s, target) {
    var inStr = false;
    for (var i = 0; i < s.length; i++) {
        var ch = s.charAt(i);
        if (ch === '"') {
            if (s.charAt(i + 1) === '"') i++;
            else inStr = !inStr;
        } else if (!inStr && ch === target) return true;
    }
    return false;
}

// =============================================================================
// NORMALIZAÇÃO DE TEMA (usa normalizarTema_ do 01_Utils se existir)
// =============================================================================
function normalizarTemaPV2_(t) {
    // Normalização mais forte para a V2.
    // O ponto crítico aqui é unificar BREAK, MERCHAN, BREAK/MERCHAN e MERCHAN/BREAK.
    // Sem isso, a coluna F pode alterar a projeção antiga (UA/UB/UD), mas a TY não encontra fator de tema.
    if (t === null || t === undefined) return "";
    var s = (t + "").trim();
    if (!s) return "";
    try { s = s.normalize("NFKC"); } catch (e) {}
    s = s.replace(/\s+/g, " ").trim().toUpperCase();

    if (typeof normalizarTema_ === "function") {
        try {
            var s2 = normalizarTema_(s);
            if (s2 !== null && s2 !== undefined && (s2 + "").trim()) {
                s = (s2 + "").trim().toUpperCase();
            }
        } catch (e2) {}
    }

    var k = s.replace(/\s+/g, " ").trim().toUpperCase();

    // Aliases essenciais
    if (k === "BREAK" || k === "MERCHAN" || k === "BREAK/MERCHAN" ||
        k === "MERCHAN/BREAK" || k === "INTERVALO" || k === "COMERCIAL") {
        return "MERCHAN/BREAK";
    }

    // Alguns termos que costumam aparecer como tema pontual, mas devem cair no agrupador histórico.
    if (k === "F1" || k === "FORMULA 1" || k === "FÓRMULA 1" ||
        k === "MILAN" || k === "MANCHESTER UNITED" || k === "COPINHA") {
        return "GIRO ESPORTE";
    }

    return k;
}
function isBreakPV2_(t) { return normalizarTemaPV2_(t) === "MERCHAN/BREAK"; }

// =============================================================================
// LEITURA DE CONTEXTO (1 leitura em lote por aba)
// =============================================================================
function lerContextoPV2_() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) throw new Error('Aba "' + CFG.GE_OUT_HISTORICO + '" não encontrada.');
    var n = PV2.LAST_ROW - PV2.FIRST_ROW + 1;

    // B (média dias fechados — já sem D após o fix de circularidade), D, F
    var block = sheet.getRange(PV2.FIRST_ROW, 2, n, 5).getValues(); // B..F
    var fxD = sheet.getRange(PV2.FIRST_ROW, PV2.AUD_COL, n, 1).getFormulas();
    var mediaHist = [], reais = [], temas = [];
    for (var i = 0; i < n; i++) {
        mediaHist.push(typeof block[i][0] === "number" ? block[i][0] : null); // B
        var d = block[i][2], isF = fxD[i][0] !== "" && fxD[i][0] != null;     // D
        reais.push((!isF && typeof d === "number" && d > 0) ? d : null);      // só reais digitados
        temas.push(normalizarTemaPV2_(block[i][4]));                          // F
    }

    // V6: lê a projeção antiga/PROJ (normalmente UD) como sinal auxiliar do efeito do tema.
    // Ela não manda na previsão, apenas puxa parcialmente a TY quando F muda.
    var projAntiga = new Array(n).fill(null);
    var colProjAntiga = acharColunaHeaderPV2_(sheet, PV2.HEADER_PROJ_ANTIGA);
    if (colProjAntiga > 0) {
        var pOld = sheet.getRange(PV2.FIRST_ROW, colProjAntiga, n, 1).getValues();
        for (var pi = 0; pi < n; pi++) {
            projAntiga[pi] = (typeof pOld[pi][0] === "number" && pOld[pi][0] > 0) ? pOld[pi][0] : null;
        }
    }

    // Fatores de tema:
    // 1) tenta a aba GE - Média por Minuto TEMA;
    // 2) se ela estiver com #ERROR ou sem fator para o tema, calcula fallback direto do histórico fechado.
    // Isso garante que F3:F62 afete a TY mesmo quando a aba TEMA estiver quebrada.
    var fatores = lerFatoresTemaPlanilhaPV2_(ss);
    var fatoresHist = calcularFatoresTemaHistoricoPV2_(sheet);
    var usadosHistorico = 0;
    for (var th in fatoresHist) {
        if (!fatores[th] || !fatores[th].fator || fatores[th].n < 3) {
            fatores[th] = fatoresHist[th];
            usadosHistorico++;
        }
    }

    // V6: fator específico por minuto.
    // Exemplo: BREAK/MERCHAN às 18:26 pode ter comportamento diferente de BREAK/MERCHAN às 18:05.
    // Isso é o que faz a coluna F mexer na TY de forma visível, sem tirar o peso do patamar real de D.
    var fatoresMinuto = calcularFatoresTemaMinutoHistoricoPV2_(sheet);

    return {
        sheet: sheet,
        mediaHist: mediaHist,
        reais: reais,
        temas: temas,
        projAntiga: projAntiga,
        fatores: fatores,
        fatoresMinuto: fatoresMinuto,
        fatoresOrigem: {
            planilha: Object.keys(fatores).length - usadosHistorico,
            historico: Object.keys(fatoresHist).length,
            usadosHistorico: usadosHistorico,
            minuto: Object.keys(fatoresMinuto || {}).length
        }
    };
}

// Lê os fatores já calculados na aba GE - Média por Minuto TEMA.
// A=tema | B=minutos | F=índice relativo. Se houver #ERROR, ignora a linha.
function lerFatoresTemaPlanilhaPV2_(ss) {
    var fatores = {};
    var temaSheet = ss.getSheetByName(PV2.TEMA_SHEET);
    if (!temaSheet || temaSheet.getLastRow() <= 1) return fatores;

    var td = temaSheet.getRange(2, 1, temaSheet.getLastRow() - 1, 6).getValues();
    for (var j = 0; j < td.length; j++) {
        var nome = td[j][0], mins = td[j][1], idx = td[j][5];
        if (!nome || typeof nome !== "string") continue;
        if (typeof idx !== "number" || idx <= 0) continue;

        var nAmostra = (typeof mins === "number" && mins > 0) ? mins : 0;
        var shrunk = 1 + (idx - 1) * nAmostra / (nAmostra + PV2.TEMA_SHRINK_K);
        shrunk = Math.min(PV2.TEMA_CAP_MAX, Math.max(PV2.TEMA_CAP_MIN, shrunk));
        fatores[normalizarTemaPV2_(nome)] = {
            fator: shrunk,
            n: nAmostra,
            indiceBruto: idx,
            origem: "TEMA"
        };
    }
    return fatores;
}

// Fallback robusto: calcula fatores de tema direto dos dias fechados da própria GE - Histórico.
// Método: audiência do minuto / média final daquele dia. Depois compara o tema contra a média geral.
// Assim o tema afeta a curva do minuto, mas o patamar do dia continua vindo da coluna D.
function calcularFatoresTemaHistoricoPV2_(sheet) {
    var fatores = {};
    var n = PV2.LAST_ROW - PV2.FIRST_ROW + 1;
    var lastCol = sheet.getLastColumn();

    var audCols = obterColunasAudienciaPV2_(sheet).filter(function (c) { return c !== PV2.AUD_COL; });
    if (!audCols.length) return fatores;

    var grid = sheet.getRange(PV2.FIRST_ROW, 1, n, lastCol).getValues();
    var soma = {}, cnt = {}, geralSoma = 0, geralCnt = 0;

    for (var i = 0; i < audCols.length; i++) {
        var c = audCols[i];
        var vals = [], temasDia = [];
        for (var r = 0; r < n; r++) {
            var v = grid[r][c - 1];
            vals.push((typeof v === "number" && v > 0) ? v : null);
            temasDia.push(normalizarTemaPV2_(grid[r][c + 1])); // tema fica 2 colunas após a audiência
        }

        var s = 0, q = 0;
        for (var a = 0; a < vals.length; a++) {
            if (vals[a] != null) { s += vals[a]; q++; }
        }
        if (q < 55 || s <= 0) continue; // só dias fechados/completos
        var mediaDia = s / q;

        for (var m = 0; m < n; m++) {
            var aud = vals[m];
            if (aud == null) continue;
            var tema = temasDia[m] || "NONE";
            var rel = aud / mediaDia;
            soma[tema] = (soma[tema] || 0) + rel;
            cnt[tema] = (cnt[tema] || 0) + 1;
            geralSoma += rel;
            geralCnt++;
        }
    }

    var geral = geralCnt ? geralSoma / geralCnt : 1;
    for (var t in soma) {
        if (!cnt[t] || !geral) continue;
        var raw = (soma[t] / cnt[t]) / geral;
        var shrunk = 1 + (raw - 1) * cnt[t] / (cnt[t] + PV2.TEMA_SHRINK_K);
        shrunk = Math.min(PV2.TEMA_CAP_MAX, Math.max(PV2.TEMA_CAP_MIN, shrunk));
        fatores[t] = {
            fator: shrunk,
            n: cnt[t],
            indiceBruto: raw,
            origem: "HISTORICO"
        };
    }
    return fatores;
}


// V6: fator de tema específico por minuto, calculado direto dos dias fechados.
// Ideia matemática:
//   1) Remove o patamar de cada dia usando audiência_do_minuto / média_final_do_dia.
//   2) Remove o efeito natural do horário comparando com a curva média relativa daquele minuto.
//   3) O que sobra é efeito aproximado do tema naquele minuto.
// Isso evita dois erros: tema dominar o dia inteiro OU tema ficar invisível na TY.
function calcularFatoresTemaMinutoHistoricoPV2_(sheet) {
    var out = {};
    var n = PV2.LAST_ROW - PV2.FIRST_ROW + 1;
    var lastCol = sheet.getLastColumn();
    var audCols = obterColunasAudienciaPV2_(sheet).filter(function (c) { return c !== PV2.AUD_COL; });
    if (!audCols.length) return out;

    var grid = sheet.getRange(PV2.FIRST_ROW, 1, n, lastCol).getValues();
    var dias = [];

    for (var i = 0; i < audCols.length; i++) {
        var c = audCols[i];
        var vals = [], temasDia = [];
        var somaDia = 0, cntDia = 0;
        for (var r = 0; r < n; r++) {
            var v = grid[r][c - 1];
            var aud = (typeof v === "number" && v > 0) ? v : null;
            vals.push(aud);
            if (aud != null) { somaDia += aud; cntDia++; }
            temasDia.push(normalizarTemaPV2_(grid[r][c + 1]));
        }
        if (cntDia < 55 || somaDia <= 0) continue;
        dias.push({ aud: vals, temas: temasDia, media: somaDia / cntDia });
    }
    if (!dias.length) return out;

    // Curva relativa média por minuto: esperado do minuto sem olhar tema.
    var relMinSoma = new Array(n).fill(0);
    var relMinCnt = new Array(n).fill(0);
    for (var d = 0; d < dias.length; d++) {
        for (var m = 0; m < n; m++) {
            var a = dias[d].aud[m];
            if (a == null || !dias[d].media) continue;
            relMinSoma[m] += a / dias[d].media;
            relMinCnt[m]++;
        }
    }
    var relMinBase = [];
    for (var b = 0; b < n; b++) relMinBase[b] = relMinCnt[b] ? relMinSoma[b] / relMinCnt[b] : 1;

    // Soma dos efeitos por tema e minuto.
    var soma = {}, cnt = {};
    for (var d2 = 0; d2 < dias.length; d2++) {
        for (var m2 = 0; m2 < n; m2++) {
            var a2 = dias[d2].aud[m2];
            if (a2 == null || !dias[d2].media || !relMinBase[m2]) continue;
            var tema = dias[d2].temas[m2] || "NONE";
            var efeito = (a2 / dias[d2].media) / relMinBase[m2];
            if (!soma[tema]) { soma[tema] = new Array(n).fill(0); cnt[tema] = new Array(n).fill(0); }
            soma[tema][m2] += efeito;
            cnt[tema][m2]++;
        }
    }

    for (var t in soma) {
        out[t] = { fatores: new Array(n).fill(1), n: cnt[t].slice(), origem: "HISTORICO_MINUTO" };
        for (var mm = 0; mm < n; mm++) {
            var q = cnt[t][mm] || 0;
            if (q < PV2.TEMA_MINUTO_MIN_AMOSTRA) {
                out[t].fatores[mm] = 1;
                continue;
            }
            var raw = soma[t][mm] / q;
            var shrunk = 1 + (raw - 1) * q / (q + PV2.TEMA_MINUTO_SHRINK_K);
            out[t].fatores[mm] = Math.min(PV2.TEMA_CAP_MAX, Math.max(PV2.TEMA_CAP_MIN, shrunk));
        }
    }
    return out;
}

function fatorTemaTYV4_(ctx, tema, m) {
    tema = normalizarTemaPV2_(tema);
    if (!tema) return { fator: 1, base: 1, global: 1, minuto: 1, nMinuto: 0 };

    var fGlobal = 1;
    if (ctx.fatores && ctx.fatores[tema] && typeof ctx.fatores[tema].fator === "number") {
        fGlobal = ctx.fatores[tema].fator;
    }

    var fMin = 1, nMin = 0;
    if (ctx.fatoresMinuto && ctx.fatoresMinuto[tema] && ctx.fatoresMinuto[tema].fatores) {
        fMin = ctx.fatoresMinuto[tema].fatores[m] || 1;
        nMin = ctx.fatoresMinuto[tema].n ? (ctx.fatoresMinuto[tema].n[m] || 0) : 0;
    }

    // Se não há amostra mínima naquele minuto, usa só o global.
    var base;
    if (nMin >= PV2.TEMA_MINUTO_MIN_AMOSTRA && fMin !== 1) {
        base = 1 + (fMin - 1) * PV2.TEMA_MINUTO_PESO + (fGlobal - 1) * PV2.TEMA_GLOBAL_PESO;
    } else {
        base = fGlobal;
    }

    // Reforço moderado: torna o tema visível na TY, mas limitado por cap.
    var reforcado = 1 + (base - 1) * PV2.TEMA_GAIN_TY;
    reforcado = Math.min(PV2.TEMA_TY_CAP_MAX, Math.max(PV2.TEMA_TY_CAP_MIN, reforcado));
    return { fator: reforcado, base: base, global: fGlobal, minuto: fMin, nMinuto: nMin };
}

function obterColunasAudienciaPV2_(sheet) {
    if (typeof getAudColsGE_ === "function") {
        try { return getAudColsGE_(sheet); } catch (e) {}
    }
    var lastCol = sheet.getLastColumn();
    var hs = sheet.getRange(PV2.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
    var out = [];
    for (var c = 0; c < hs.length; c++) {
        var h = ((hs[c] || "") + "").trim().toUpperCase();
        if (h === "AUDIÊNCIA" || h === "AUDIENCIA") out.push(c + 1);
    }
    return out;
}

// =============================================================================
// MOTOR INTERNO V2 (o mesmo validado no backtest)
// =============================================================================
function medianaPV2_(arr) {
    if (!arr.length) return null;
    var a = arr.slice().sort(function (x, y) { return x - y; });
    var mid = Math.floor(a.length / 2);
    return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function calcularPrevisaoInternaV2_(ctx) {
    var n = 60, reais = ctx.reais, mh = ctx.mediaHist;
    // último minuto com real digitado (exige sequência a partir do início não é necessário)
    var idxReais = [];
    for (var i = 0; i < n; i++) if (reais[i] != null && mh[i]) idxReais.push(i);
    var k = idxReais.length ? idxReais[idxReais.length - 1] + 1 : 0;

    var ratios = idxReais.map(function (i) { return reais[i] / mh[i]; });
    var patamarBruto = ratios.length ? medianaPV2_(ratios) : 1.0;
    var w = Math.min(1, ratios.length / PV2.PATAMAR_N_PLENO);
    var patamar = 1 + (patamarBruto - 1) * w;

    // tendência: média dos últimos 3 ratios vs últimos 8
    var trend = 1.0;
    if (ratios.length >= 3) {
        var r3 = ratios.slice(-3).reduce(function (a, b) { return a + b; }, 0) / 3;
        var r8arr = ratios.slice(-8);
        var r8 = r8arr.reduce(function (a, b) { return a + b; }, 0) / r8arr.length;
        if (r8 > 0) trend = Math.min(PV2.TREND_MAX, Math.max(PV2.TREND_MIN, r3 / r8));
    }
    // desvio local do último real vs patamar
    var devLocal = ratios.length ? (ratios[ratios.length - 1] / patamar - 1) : 0;

    var proj = new Array(n).fill(null);
    var dev = devLocal, temaImpactoMedio = 0, nTemaFut = 0;
    var temaMinutoUsado = 0, temaPlanilhaUsado = 0;
    for (var m = k; m < n; m++) {
        if (!mh[m]) continue;
        dev *= betaPV2_(m);
        var decay = Math.pow(betaPV2_(m), m - k + 1);
        var tr = 1 + (trend - 1) * decay;

        // V6: tema entra por 3 caminhos:
        // 1) fator estatístico global; 2) fator estatístico por minuto; 3) sinal auxiliar da PROJ antiga/UD.
        // O item 3 é limitado e ponderado, para F mexer visivelmente na TY sem tirar o peso da coluna D.
        var tm = ctx.temas[m];
        var ft = fatorTemaTYV4_(ctx, tm, m);
        var baseProj = mh[m] * patamar * tr * (1 + dev);
        var tfTema = ft.fator;
        var tfPlanilha = 1;
        var oldProj = ctx.projAntiga ? ctx.projAntiga[m] : null;

        if (tm && oldProj && baseProj > 0) {
            var rawPlanilha = oldProj / baseProj;
            rawPlanilha = Math.min(PV2.TEMA_PLANILHA_CAP_MAX, Math.max(PV2.TEMA_PLANILHA_CAP_MIN, rawPlanilha));
            tfPlanilha = 1 + (rawPlanilha - 1) * PV2.TEMA_PLANILHA_BLEND;
            temaPlanilhaUsado++;
        }

        var tf = tfTema * tfPlanilha;
        tf = Math.min(PV2.TEMA_TY_CAP_MAX, Math.max(PV2.TEMA_TY_CAP_MIN, tf));

        if (tm) {
            temaImpactoMedio += (tf - 1);
            nTemaFut++;
            if (ft.nMinuto >= PV2.TEMA_MINUTO_MIN_AMOSTRA) temaMinutoUsado++;
        }

        proj[m] = Math.round(baseProj * tf * 100) / 100;

        // Carry pequeno: se um tema altera o minuto, parte desse desvio contamina os próximos minutos e decai pelo beta.
        if (tm && tf !== 1) dev += (tf - 1) * PV2.TEMA_CARRY_PESO;
    }

    var conf = ratios.length >= 15 ? "ALTA" : ratios.length >= 5 ? "MÉDIA" : "BAIXA";
    var tendTxt = trend > 1.015 ? (patamarBruto < 0.97 ? "RECUPERAÇÃO" : "ALTA")
        : trend < 0.985 ? (patamarBruto > 1.03 ? "PERDA DE FÔLEGO" : "BAIXA") : "ESTÁVEL";
    var motivo = [];
    motivo.push("dia " + ((patamarBruto - 1) * 100).toFixed(1).replace(".", ",") + "% vs histórico");
    if (nTemaFut) motivo.push("temas futuros " + (temaImpactoMedio / nTemaFut >= 0 ? "+" : "") +
        ((temaImpactoMedio / nTemaFut) * 100).toFixed(1).replace(".", ",") + "%");
    if (ctx.fatoresOrigem && ctx.fatoresOrigem.usadosHistorico) {
        motivo.push("fator tema via histórico");
    }
    if (temaMinutoUsado) {
        motivo.push("tema minuto ativo " + temaMinutoUsado + "x");
    }
    if (temaPlanilhaUsado) {
        motivo.push("sinal PROJ/UD ativo " + temaPlanilhaUsado + "x");
    }
    motivo.push("tendência " + tendTxt.toLowerCase());
    if (conf === "BAIXA") motivo.push("confiança baixa: poucos reais");

    var soma = 0, cnt = 0;
    for (var q = 0; q < n; q++) {
        var v = reais[q] != null ? reais[q] : proj[q];
        if (v != null) { soma += v; cnt++; }
    }
    return {
        origem: "INTERNO",
        proj: proj, k: k,
        mediaFinalPrevista: cnt ? Math.round(soma / cnt * 100) / 100 : null,
        patamar: patamarBruto, tendencia: tendTxt, confianca: conf,
        motivo: motivo.join(" | ")
    };
}

// =============================================================================
// GRAVAÇÃO (preserva reais; só escreve V2 e resumo)
// =============================================================================
function acharColunaHeaderPV2_(sheet, header) {
    var lastCol = sheet.getLastColumn();
    var hs = sheet.getRange(PV2.HEADER_ROW, 1, 1, lastCol).getValues()[0];
    for (var c = 0; c < hs.length; c++)
        if (((hs[c] || "") + "").trim() === header) return c + 1;
    return -1;
}

function gravarPrevisaoV2_(ctx, resultado) {
    var sheet = ctx.sheet;
    var colV2 = acharColunaHeaderPV2_(sheet, PV2.HEADER_V2);
    if (colV2 < 0) throw new Error("Coluna PROJ V2 não instalada. Rode instalarPrevisaoExternaGE().");
    var n = PV2.LAST_ROW - PV2.FIRST_ROW + 1;
    var vals = [];
    for (var i = 0; i < n; i++) vals.push([resultado.proj[i] != null ? resultado.proj[i] : ""]);
    sheet.getRange(PV2.FIRST_ROW, colV2, n, 1).setValues(vals).setNumberFormat("0.00");
    var resumo = "🔮 " + resultado.origem + " | média final prevista: " +
        (resultado.mediaFinalPrevista != null ? resultado.mediaFinalPrevista : "—") +
        " | " + resultado.motivo;
    sheet.getRange(1, colV2).setValue(resumo).setFontSize(9).setWrap(false);
    Logger.log("[PV2] " + resumo);
}

// =============================================================================
// RECÁLCULO AO VIVO (público + handler do gatilho instalável)
// =============================================================================
function recalcularPrevisaoAoVivoGE() {
    // V6: a TY depende também da projeção antiga (UD).
    // Quando você edita F, UA/UB/UD podem levar 1-2s para recalcular.
    // Se a V2 ler a planilha cedo demais, ela pega a UD antiga e parece que F não mexeu na TY.
    try {
        SpreadsheetApp.flush();
        Utilities.sleep(PV2.WAIT_FORMULAS_MS || 2200);
        SpreadsheetApp.flush();
    } catch (e0) {}
    var ctx = lerContextoPV2_();
    var resultado = null;
    if (previsaoExternaAtivaPV2_()) {
        try { resultado = chamarApiPrevisaoGE_(ctx); }
        catch (e) { Logger.log("[PV2] API falhou (" + e.message + ") — fallback interno."); }
    }
    if (!resultado) resultado = calcularPrevisaoInternaV2_(ctx);
    gravarPrevisaoV2_(ctx, resultado);
    return resultado;
}

// Handler do gatilho INSTALÁVEL (criado por instalarPrevisaoExternaGE)
function onEditPrevisaoV2GE(e) {
    try {
        if (!e || !e.range) return;
        var sheet = e.range.getSheet();
        if (sheet.getName() !== CFG.GE_OUT_HISTORICO) return;
        var col = e.range.getColumn(), row = e.range.getRow();
        if ((col !== PV2.AUD_COL && col !== PV2.TEMA_COL) ||
            row < PV2.FIRST_ROW || row > PV2.LAST_ROW) return;
        // Throttle p/ evitar tempestade de recálculos em digitação rápida
        var cache = CacheService.getScriptCache();
        if (cache.get("PV2_LOCK")) return;
        cache.put("PV2_LOCK", "1", PV2.THROTTLE_S);
        recalcularPrevisaoAoVivoGE();
    } catch (err) {
        Logger.log("[PV2] onEdit erro (silencioso): " + err.message);
    }
}

// =============================================================================
// INSTALAÇÃO
// =============================================================================
function instalarPrevisaoExternaGE() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) { ui.alert("❌", 'Aba "Histórico" não existe.', ui.ButtonSet.OK); return; }

    // 1) Coluna PROJ V2 (cria à direita se não existir)
    var colV2 = acharColunaHeaderPV2_(sheet, PV2.HEADER_V2);
    if (colV2 < 0) {
        colV2 = sheet.getLastColumn() + 2;
        sheet.getRange(PV2.HEADER_ROW, colV2).setValue(PV2.HEADER_V2)
            .setFontWeight("bold").setBackground("#0B5345").setFontColor("white");
        sheet.setColumnWidth(colV2, 95);
    }
    var Lv2 = columnToLetter_(colV2);
    var colProj = acharColunaHeaderPV2_(sheet, PV2.HEADER_PROJ_ANTIGA);
    var Lproj = colProj > 0 ? columnToLetter_(colProj) : null;

    // 2) Fórmulas em D (só onde NÃO há real digitado) com fallback p/ PROJ antiga
    var n = PV2.LAST_ROW - PV2.FIRST_ROW + 1;
    var dRange = sheet.getRange(PV2.FIRST_ROW, PV2.AUD_COL, n, 1);
    var dVals = dRange.getValues(), dFx = dRange.getFormulas();
    var cells = [], preservadas = 0;
    for (var i = 0; i < n; i++) {
        var row = PV2.FIRST_ROW + i;
        var hasF = dFx[i][0] !== "" && dFx[i][0] != null;
        var typed = !hasF && typeof dVals[i][0] === "number" && dVals[i][0] !== 0;
        if (typed) { preservadas++; continue; }
        if (row === PV2.FIRST_ROW) continue; // D3 aguarda o primeiro real
        var fb = Lproj ? Lproj + row : '""';
        cells.push({
            row: row, col: PV2.AUD_COL,
            formula: "=IF(D" + (row - 1) + '="","",IF(' + Lv2 + row + '<>"",' + Lv2 + row + "," + fb + "))"
        });
    }
    if (cells.length) batchSetFormulasSparsePV2Locale_(sheet, cells);

    // 3) Gatilho instalável onEdit (só um)
    var trs = ScriptApp.getProjectTriggers(), tem = false;
    for (var t = 0; t < trs.length; t++)
        if (trs[t].getHandlerFunction() === "onEditPrevisaoV2GE") tem = true;
    if (!tem) ScriptApp.newTrigger("onEditPrevisaoV2GE").forSpreadsheet(ss).onEdit().create();

    // 4) Primeiro cálculo
    var res = recalcularPrevisaoAoVivoGE();
    ui.alert("✅ Previsão V2 instalada",
        "Coluna PROJ V2 criada, D futuro com fallback p/ projeção antiga, gatilho ativo.\n" +
        "Reais preservados: " + preservadas + "\nFórmulas em D: " + cells.length +
        "\n\nResumo: " + res.motivo +
        "\n\nMotor externo: DESATIVADO (padrão). Use ativarPrevisaoExternaGE() após configurar a API.",
        ui.ButtonSet.OK);
}

// =============================================================================
// API EXTERNA (opcional)
// =============================================================================
function previsaoExternaAtivaPV2_() {
    return PropertiesService.getScriptProperties().getProperty(PV2.PROP_ATIVA) === "1";
}

function ativarPrevisaoExternaGE() {
    var ui = SpreadsheetApp.getUi();
    var props = PropertiesService.getScriptProperties();
    var r1 = ui.prompt("🌐 URL da API", "URL completa do endpoint /prever:\nAtual: " +
        (props.getProperty(PV2.PROP_URL) || "(vazia)"), ui.ButtonSet.OK_CANCEL);
    if (r1.getSelectedButton() !== ui.Button.OK) return;
    if (r1.getResponseText().trim()) props.setProperty(PV2.PROP_URL, r1.getResponseText().trim());
    var r2 = ui.prompt("🔑 Secret", "Chave secreta da API (em branco = manter):", ui.ButtonSet.OK_CANCEL);
    if (r2.getSelectedButton() !== ui.Button.OK) return;
    if (r2.getResponseText().trim()) props.setProperty(PV2.PROP_SECRET, r2.getResponseText().trim());
    props.setProperty(PV2.PROP_ATIVA, "1");
    ui.alert("✅", "Motor externo ATIVADO. Qualquer falha da API cai no motor interno automaticamente.", ui.ButtonSet.OK);
}

function desativarPrevisaoExternaGE() {
    PropertiesService.getScriptProperties().setProperty(PV2.PROP_ATIVA, "0");
    SpreadsheetApp.getUi().alert("✅", "Motor externo desativado. Motor interno V2 segue ativo.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function montarPayloadPrevisaoGE_(ctx) {
    var sheet = ctx.sheet;
    // histórico fechado: blocos com data na linha 1 (exclui D=hoje)
    var audCols = getAudColsGE_(sheet).filter(function (c) { return c !== 4; });
    var lastCol = sheet.getLastColumn();
    var row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var n = 60;
    var historico = [];
    for (var i = 0; i < audCols.length; i++) {
        var c = audCols[i];
        var v = row1[c - 1];
        var dataStr = v instanceof Date
            ? Utilities.formatDate(v, "America/Sao_Paulo", "yyyy-MM-dd")
            : (typeof v === "string" && v.match(/^\d{2}\/\d{2}\/\d{4}$/)
                ? v.split("/").reverse().join("-") : null);
        var bloco = sheet.getRange(PV2.FIRST_ROW, c, n, 3).getValues(); // AUD, SHA, TEMA
        historico.push({
            data: dataStr,
            audiencia: bloco.map(function (r) { return typeof r[0] === "number" ? r[0] : null; }),
            temas: bloco.map(function (r) { return normalizarTemaPV2_(r[2]); })
        });
    }
    var aliases = { "BREAK": "BREAK/MERCHAN", "MERCHAN": "BREAK/MERCHAN", "MERCHAN/BREAK": "BREAK/MERCHAN",
        "MILAN": "GIRO ESPORTE", "MANCHESTER UNITED": "GIRO ESPORTE", "F1": "GIRO ESPORTE", "COPINHA": "GIRO ESPORTE" };
    return {
        secret: PropertiesService.getScriptProperties().getProperty(PV2.PROP_SECRET) || "",
        data_hoje: Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyy-MM-dd"),
        dia_semana: new Date().getDay(),
        media_historica_minuto: ctx.mediaHist,
        reais_hoje: ctx.reais,
        temas_hoje: ctx.temas,
        historico: historico,
        aliases: aliases,
        fatores_tema: ctx.fatores
    };
}

function validarPayloadPrevisaoGE() {
    var ctx = lerContextoPV2_();
    var p = montarPayloadPrevisaoGE_(ctx);
    var erros = [];
    if (!p.historico.length) erros.push("Sem dias fechados no histórico.");
    if (p.media_historica_minuto.filter(function (v) { return v; }).length < 50)
        erros.push("Coluna B (média histórica) incompleta.");
    if (p.historico.length && p.historico[0].audiencia.length !== 60)
        erros.push("Bloco histórico sem 60 minutos.");
    var msg = erros.length ? ("❌ " + erros.join("\n")) :
        ("✅ Payload OK: " + p.historico.length + " dias fechados, " +
            p.reais_hoje.filter(function (v) { return v != null; }).length + " reais hoje, " +
            Object.keys(p.fatores_tema).length + " temas com fator.");
    Logger.log("[PV2] " + msg);
    try { SpreadsheetApp.getUi().alert("Validação do payload", msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
    return erros.length === 0;
}

function chamarApiPrevisaoGE_(ctx) {
    var url = PropertiesService.getScriptProperties().getProperty(PV2.PROP_URL);
    if (!url) throw new Error("PV2_API_URL não configurada.");
    var payload = montarPayloadPrevisaoGE_(ctx);
    var resp = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true, followRedirects: true,
        // timeout implícito do Apps Script; validamos código e formato
    });
    if (resp.getResponseCode() !== 200)
        throw new Error("HTTP " + resp.getResponseCode() + ": " + resp.getContentText().substring(0, 200));
    var out = JSON.parse(resp.getContentText());
    if (!out || !out.proj || out.proj.length !== 60) throw new Error("Resposta inválida da API.");
    return {
        origem: "API",
        proj: out.proj.map(function (v, i) { return ctx.reais[i] != null ? null : v; }),
        k: out.k || 0,
        mediaFinalPrevista: out.media_final_prevista,
        patamar: out.patamar, tendencia: out.tendencia, confianca: out.confianca,
        motivo: (out.motivo || "") +
            (out.faixa ? " | faixa " + out.faixa.pessimista + "–" + out.faixa.otimista : "")
    };
}

function testarApiPrevisaoGE() {
    var ui = SpreadsheetApp.getUi();
    try {
        var ctx = lerContextoPV2_();
        var t0 = Date.now();
        var res = chamarApiPrevisaoGE_(ctx);
        ui.alert("✅ API OK (" + (Date.now() - t0) + " ms)",
            "Média final prevista: " + res.mediaFinalPrevista + "\n" + res.motivo, ui.ButtonSet.OK);
    } catch (e) {
        ui.alert("❌ API falhou", e.message + "\n\nO motor interno continua funcionando normalmente.", ui.ButtonSet.OK);
    }
}


// =============================================================================
// DIAGNÓSTICO V6 — verifica se F/UD está entrando na TY
// =============================================================================
function diagnosticarBlendTemaUDLinhaGE() {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sh) { SpreadsheetApp.getUi().alert('Aba GE - Histórico não encontrada.'); return; }
    var cell = sh.getActiveCell();
    var row = cell ? cell.getRow() : 0;
    if (row < PV2.FIRST_ROW || row > PV2.LAST_ROW) row = 29; // 18:26 como padrão

    try {
        SpreadsheetApp.flush();
        Utilities.sleep(PV2.WAIT_FORMULAS_MS || 2200);
        SpreadsheetApp.flush();
    } catch(e) {}

    var ctx = lerContextoPV2_();
    var m = row - PV2.FIRST_ROW;
    var res = calcularPrevisaoInternaV2_(ctx);
    var colV2 = acharColunaHeaderPV2_(sh, PV2.HEADER_V2);
    var colProj = acharColunaHeaderPV2_(sh, PV2.HEADER_PROJ_ANTIGA);
    var tema = ctx.temas[m] || '';
    var ty = colV2 > 0 ? sh.getRange(row, colV2).getDisplayValue() : 'sem TY';
    var ud = colProj > 0 ? sh.getRange(row, colProj).getDisplayValue() : 'sem PROJ';
    var msg = 'Linha ' + row + ' / ' + sh.getRange(row, 1).getDisplayValue() + '\n' +
        'Tema F: ' + (tema || '(vazio)') + '\n' +
        'TY atual na planilha: ' + ty + '\n' +
        'PROJ/UD atual na planilha: ' + ud + '\n' +
        'TY recalculada agora: ' + (res.proj[m] != null ? res.proj[m] : '') + '\n\n' +
        'Resumo: ' + res.motivo + '\n\n' +
        'Se UD mudou e TY recalculada não mudou, o tema está sendo amortecido.\n' +
        'Se TY recalculada mudou mas TY atual não, faltou rodar/gravar o recálculo.';
    SpreadsheetApp.getUi().alert('Diagnóstico Tema → TY', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// =============================================================================
// ROLLBACK — volta 100% ao sistema antigo (cadeia PROJ do 18)
// =============================================================================
function rollbackPrevisaoInternaGE() {
    var ui = SpreadsheetApp.getUi();
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    var colProj = acharColunaHeaderPV2_(sheet, PV2.HEADER_PROJ_ANTIGA);
    var colV2 = acharColunaHeaderPV2_(sheet, PV2.HEADER_V2);
    var Lproj = colProj > 0 ? columnToLetter_(colProj) : null;
    var n = PV2.LAST_ROW - PV2.FIRST_ROW + 1;
    var dRange = sheet.getRange(PV2.FIRST_ROW, PV2.AUD_COL, n, 1);
    var dVals = dRange.getValues(), dFx = dRange.getFormulas();
    var cells = [];
    for (var i = 0; i < n; i++) {
        var row = PV2.FIRST_ROW + i;
        var hasF = dFx[i][0] !== "" && dFx[i][0] != null;
        var typed = !hasF && typeof dVals[i][0] === "number" && dVals[i][0] !== 0;
        if (typed || row === PV2.FIRST_ROW) continue;
        cells.push({
            row: row, col: PV2.AUD_COL,
            formula: Lproj ? "=IF(D" + (row - 1) + '="","",' + Lproj + row + ")" : ""
        });
    }
    if (cells.length) batchSetFormulasSparsePV2Locale_(sheet, cells);
    if (colV2 > 0) sheet.getRange(PV2.FIRST_ROW, colV2, n, 1).clearContent();
    var trs = ScriptApp.getProjectTriggers();
    for (var t = 0; t < trs.length; t++)
        if (trs[t].getHandlerFunction() === "onEditPrevisaoV2GE") ScriptApp.deleteTrigger(trs[t]);
    PropertiesService.getScriptProperties().setProperty(PV2.PROP_ATIVA, "0");
    ui.alert("↩️ Rollback concluído",
        "Coluna D voltou à cadeia PROJ antiga, PROJ V2 limpa, gatilho removido, API desativada.\nReais digitados intactos.",
        ui.ButtonSet.OK);
}

// =============================================================================
// ESTUDO (resumo executivo em alert)
// =============================================================================
function estudarMotorExternoPrevisaoGE() {
    SpreadsheetApp.getUi().alert("📚 Estudo — Motor externo de previsão",
        "LICENÇAS: mlforecast (Apache-2.0), skforecast (BSD-3), LightGBM (MIT), " +
        "XGBoost (Apache-2.0), CatBoost (Apache-2.0), scikit-learn (BSD-3) — todas " +
        "gratuitas p/ uso comercial. TimeGPT/Nixtla Cloud é serviço PAGO e NÃO é usado.\n\n" +
        "BACKTEST (68 dias, MAE da média final): o motor interno V2 supera o modelo " +
        "antigo a partir de 18:15 e empata/supera LightGBM/skforecast puros. Ensemble " +
        "com ML melhora ~3-5% — disponível via API externa opcional.\n\n" +
        "RECOMENDAÇÃO: motor interno V2 como padrão (já instalado por " +
        "instalarPrevisaoExternaGE). API externa: opcional, deploy gratuito com " +
        "limitações (sleep/cold start) — ver LEIA-ME.",
        SpreadsheetApp.getUi().ButtonSet.OK);
}

// =============================================================================
// BACKTEST INTERNO — compara modelo Atual (18) × V2, escreve aba de resultados
// =============================================================================
function rodarBacktestPrevisaoGE() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    var ui = SpreadsheetApp.getUi();

    // Ler TODOS os blocos fechados de uma vez
    var audCols = getAudColsGE_(sheet).filter(function (c) { return c !== 4; });
    var lastCol = sheet.getLastColumn();
    var row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var grid = sheet.getRange(PV2.FIRST_ROW, 1, 60, lastCol).getValues();
    var progs = [];
    for (var i = 0; i < audCols.length; i++) {
        var c = audCols[i], aud = [], temas = [];
        var ok = 0;
        for (var r = 0; r < 60; r++) {
            var a = grid[r][c - 1];
            aud.push(typeof a === "number" && a > 0 ? a : null);
            if (aud[r] != null) ok++;
            temas.push(normalizarTemaPV2_(grid[r][c + 1]));
        }
        if (ok < 55) continue;
        var v = row1[c - 1];
        var dt = v instanceof Date ? v : null;
        if (!dt && typeof v === "string" && v.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
            var pts = v.split("/"); dt = new Date(+pts[2], +pts[1] - 1, +pts[0]);
        }
        progs.push({ aud: aud, temas: temas, dt: dt });
    }
    progs.sort(function (a, b) { return (a.dt || 0) - (b.dt || 0); }); // antigo → novo
    if (progs.length <= PV2.MIN_TREINO_BACKTEST) {
        ui.alert("⚠️", "Histórico insuficiente (" + progs.length + " dias completos; mínimo " +
            (PV2.MIN_TREINO_BACKTEST + 1) + ").", ui.ButtonSet.OK);
        return;
    }

    function mediaFinal(p) {
        var s = 0, n2 = 0;
        for (var q = 0; q < 60; q++) if (p.aud[q] != null) { s += p.aud[q]; n2++; }
        return n2 ? s / n2 : null;
    }
    function curvaMedia(train) {
        var out = [];
        for (var m = 0; m < 60; m++) {
            var s = 0, n2 = 0;
            for (var j = 0; j < train.length; j++)
                if (train[j].aud[m] != null) { s += train[j].aud[m]; n2++; }
            out.push(n2 ? s / n2 : null);
        }
        return out;
    }
    function fatoresTema(train, shrinkK, capMin, capMax) {
        var s = {}, cnt = {}, gN = 0, gD = 0;
        for (var j = 0; j < train.length; j++) {
            var mf = mediaFinal(train[j]);
            if (!mf) continue;
            for (var m = 0; m < 60; m++) {
                var a = train[j].aud[m];
                if (a == null) continue;
                var t = train[j].temas[m] || "NONE";
                var rel = a / mf;
                s[t] = (s[t] || 0) + rel; cnt[t] = (cnt[t] || 0) + 1; gN += rel; gD++;
            }
        }
        var g = gD ? gN / gD : 1, out = {};
        for (var t2 in s) {
            var raw = (s[t2] / cnt[t2]) / g;
            var f = shrinkK > 0 ? 1 + (raw - 1) * cnt[t2] / (cnt[t2] + shrinkK) : raw;
            out[t2] = Math.min(capMax, Math.max(capMin, f));
        }
        return out;
    }
    function simular(modelo, train, day, k, curva, fac) {
        var out = day.aud.slice(0, k);
        if (modelo === "A") {
            var molde = curva.map(function (v, m) {
                var t = day.temas[m];
                return (v && t && t !== "NONE" && fac[t]) ? v * fac[t] : v;
            });
            var dev = molde[k - 1] ? out[k - 1] / molde[k - 1] - 1 : 0;
            for (var m = k; m < 60; m++) {
                dev *= betaPV2_(m);
                out.push(molde[m] ? molde[m] * (1 + dev) : out[out.length - 1]);
            }
        } else { // V2
            var ratios = [];
            for (var q = 0; q < k; q++) if (curva[q]) ratios.push(out[q] / curva[q]);
            var pat = ratios.length ? medianaPV2_(ratios) : 1;
            var w = Math.min(1, ratios.length / PV2.PATAMAR_N_PLENO);
            var patamar = 1 + (pat - 1) * w;
            var trend = 1;
            if (ratios.length >= 3) {
                var r3 = ratios.slice(-3).reduce(function (a, b) { return a + b; }, 0) / 3;
                var r8a = ratios.slice(-8);
                var r8 = r8a.reduce(function (a, b) { return a + b; }, 0) / r8a.length;
                if (r8 > 0) trend = Math.min(PV2.TREND_MAX, Math.max(PV2.TREND_MIN, r3 / r8));
            }
            var dev2 = ratios.length ? ratios[ratios.length - 1] / patamar - 1 : 0;
            for (var m2 = k; m2 < 60; m2++) {
                dev2 *= betaPV2_(m2);
                var decay = Math.pow(betaPV2_(m2), m2 - k + 1);
                var tr = 1 + (trend - 1) * decay;
                var tf = 1, tm = day.temas[m2];
                if (tm && tm !== "NONE" && fac[tm]) tf = fac[tm];
                out.push(curva[m2] ? curva[m2] * patamar * tr * (1 + dev2) * tf : out[out.length - 1]);
            }
        }
        return out;
    }

    var CK = PV2.CHECKPOINTS;
    var res = { A: {}, V2: {} };
    CK.forEach(function (k) { res.A[k] = []; res.V2[k] = []; });
    var testDays = 0;
    for (var d = PV2.MIN_TREINO_BACKTEST; d < progs.length; d++) {
        var day = progs[d];
        var completo = day.aud.every(function (v) { return v != null; });
        if (!completo) continue;
        var train = progs.slice(0, d);
        var curva = curvaMedia(train);
        var facA = fatoresTema(train, 0, 0.5, 1.5);
        var facB = fatoresTema(train, PV2.TEMA_SHRINK_K, PV2.TEMA_CAP_MIN, PV2.TEMA_CAP_MAX);
        var realFinal = mediaFinal(day);
        testDays++;
        for (var ci = 0; ci < CK.length; ci++) {
            var k = CK[ci];
            var pA = simular("A", train, day, k, curva, facA);
            var pB = simular("V2", train, day, k, curva, facB);
            var mA = pA.reduce(function (a, b) { return a + b; }, 0) / 60;
            var mB = pB.reduce(function (a, b) { return a + b; }, 0) / 60;
            res.A[k].push(mA - realFinal);
            res.V2[k].push(mB - realFinal);
        }
    }

    // Escrever aba
    var out = ss.getSheetByName(PV2.BACKTEST_SHEET);
    if (out) out.clear(); else out = ss.insertSheet(PV2.BACKTEST_SHEET);
    var rows = [["BACKTEST PREVISÃO GE — walk-forward | dias de teste: " + testDays +
        " | treino mínimo: " + PV2.MIN_TREINO_BACKTEST + " | gerado: " +
        Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm"), "", "", "", ""]];
    rows.push(["CHECKPOINT", "MAE Atual", "MAE V2", "Ganho %", "Viés V2"]);
    function mae(a) { return a.reduce(function (s, v) { return s + Math.abs(v); }, 0) / a.length; }
    function vies(a) { return a.reduce(function (s, v) { return s + v; }, 0) / a.length; }
    var v2Melhor = 0;
    CK.forEach(function (k) {
        var mA = mae(res.A[k]), mB = mae(res.V2[k]);
        if (mB < mA) v2Melhor++;
        rows.push(["18:" + ("0" + k).slice(-2), Math.round(mA * 10000) / 10000,
            Math.round(mB * 10000) / 10000,
            Math.round((1 - mB / mA) * 1000) / 10 + "%",
            Math.round(vies(res.V2[k]) * 10000) / 10000]);
    });
    rows.push(["", "", "", "", ""]);
    rows.push(["RECOMENDAÇÃO:", v2Melhor >= 4 ? "USAR MODELO ESTATÍSTICO INTERNO V2" :
        "MANTER MODELO ATUAL (V2 não superou)", "", "", ""]);
    out.getRange(1, 1, rows.length, 5).setValues(rows);
    out.getRange(2, 1, 1, 5).setFontWeight("bold").setBackground("#2C3E50").setFontColor("white");
    out.setColumnWidths(1, 5, 130);
    ui.alert("✅ Backtest concluído",
        "Aba \"" + PV2.BACKTEST_SHEET + "\" atualizada.\nV2 melhor em " + v2Melhor + "/" +
        CK.length + " checkpoints.", ui.ButtonSet.OK);
}
