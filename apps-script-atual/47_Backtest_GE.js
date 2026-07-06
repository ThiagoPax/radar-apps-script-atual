// =============================================================================
// 47_Backtest_GE.js — Backtest da predição em dias fechados
// =============================================================================
// Para cada dia fechado da Base (a partir do 6º), calcula o erro de modelos
// candidatos usando SOMENTE dados anteriores àquele dia:
//   A) Média histórica geral por minuto
//   B) Média contextual (A escalada pelo patamar do contexto do dia)
//   C) Média dos últimos 3 programas por minuto
//   D) Média dos últimos 5 programas por minuto
//   E) Patamar dinâmico: mediana dos ratios dos primeiros 10 minutos reais do
//      dia aplicada à média histórica (aproxima o mecanismo do motor V2)
// Métricas: MAE por minuto e erro absoluto da média final. Resultado na aba
// Backtest. Serve para validar se o modelo contextual supera o "burro".

function rodarBacktestGE() {
    const sh = ensureAbaGE_(CFG.SHEET_BACKTEST);
    const dias = matrizDiariaDaBaseGE_(); // [{dataStr, contexto, minutos[60]}]
    sh.clearContents();

    if (dias.length < 6) {
        sh.getRange(1, 1).setValue("Backtest precisa de pelo menos 6 programas fechados na Base (há " + dias.length + ").");
        return;
    }

    const resultados = []; // por dia e modelo
    const agregados = {};  // modelo → {maeSoma, mfSoma, n}
    const nomes = { A: "A · Média histórica", B: "B · Média contextual",
        C: "C · Últimos 3", D: "D · Últimos 5", E: "E · Patamar dinâmico (motor)" };

    for (let k = 5; k < dias.length; k++) {
        const alvo = dias[k];
        const passado = dias.slice(0, k);
        const reais = alvo.minutos;
        const mediaFinalReal = mediaValidaGE_(reais);
        if (mediaFinalReal === null) continue;

        const histPorMinuto = mediaPorMinutoGE_(passado);
        const prev = {};
        prev.A = histPorMinuto;

        // B: contextual — patamar dos dias passados com o mesmo contexto
        const ctx = String(alvo.contexto || "").toUpperCase();
        const doCtx = passado.filter(d => String(d.contexto || "").toUpperCase() === ctx && ctx);
        let fator = 1;
        const mediaGeralPassado = mediaValidaGE_([].concat.apply([], passado.map(d => d.minutos)));
        const reguaCtx = doCtx.length >= 2
            ? mediaValidaGE_([].concat.apply([], doCtx.map(d => d.minutos)))
            : mediaValidaGE_([].concat.apply([], passado.slice(-5).map(d => d.minutos)));
        if (mediaGeralPassado && reguaCtx) fator = clampGE_(reguaCtx / mediaGeralPassado, 0.5, 1.5);
        prev.B = histPorMinuto.map(v => v === null ? null : v * fator);

        prev.C = mediaPorMinutoGE_(passado.slice(-3));
        prev.D = mediaPorMinutoGE_(passado.slice(-5));

        // E: patamar dinâmico com os 10 primeiros minutos reais
        const ratios = [];
        for (let m = 0; m < 10; m++) {
            if (typeof reais[m] === "number" && histPorMinuto[m]) ratios.push(reais[m] / histPorMinuto[m]);
        }
        const patamar = ratios.length ? medianaGE_(ratios) : fator;
        prev.E = histPorMinuto.map((v, m) => {
            if (v === null) return null;
            return m < 10 ? (typeof reais[m] === "number" ? reais[m] : v * patamar) : v * patamar;
        });

        for (const mod of ["A", "B", "C", "D", "E"]) {
            let somaErr = 0, n = 0;
            for (let m = 0; m < 60; m++) {
                if (typeof reais[m] !== "number" || prev[mod][m] === null) continue;
                somaErr += Math.abs(reais[m] - prev[mod][m]);
                n++;
            }
            if (!n) continue;
            const mae = somaErr / n;
            const mfPrev = mediaValidaGE_(prev[mod]);
            const erroMF = mfPrev === null ? null : Math.abs(mfPrev - mediaFinalReal);
            resultados.push([alvo.dataStr, alvo.contexto || "—", nomes[mod],
                mae, erroMF === null ? "" : erroMF]);
            if (!agregados[mod]) agregados[mod] = { maeSoma: 0, mfSoma: 0, n: 0 };
            agregados[mod].maeSoma += mae;
            agregados[mod].mfSoma += erroMF || 0;
            agregados[mod].n++;
        }
    }

    // Escreve agregados + detalhe
    const linhas = [["BACKTEST DA PREDIÇÃO — " +
        Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy HH:mm"), "", "", "", ""],
        ["MODELO", "MAE POR MINUTO", "ERRO MÉDIA FINAL", "DIAS TESTADOS", ""]];
    let melhor = null;
    for (const mod of ["A", "B", "C", "D", "E"]) {
        const a = agregados[mod];
        if (!a || !a.n) continue;
        const mae = a.maeSoma / a.n;
        linhas.push([nomes[mod], mae, a.mfSoma / a.n, a.n, ""]);
        if (!melhor || mae < melhor.mae) melhor = { mod: nomes[mod], mae: mae };
    }
    linhas.push(["", "", "", "", ""]);
    linhas.push(["Melhor modelo: " + (melhor ? melhor.mod : "—"), "", "", "", ""]);
    linhas.push(["", "", "", "", ""]);
    linhas.push(["DATA", "CONTEXTO", "MODELO", "MAE MINUTO", "ERRO MÉDIA FINAL"]);
    const inicio = linhas.length;
    sh.getRange(1, 1, linhas.length, 5).setValues(linhas);
    if (resultados.length) {
        sh.getRange(inicio + 1, 1, resultados.length, 5).setValues(resultados);
        sh.getRange(inicio + 1, 4, resultados.length, 2).setNumberFormat("0.000");
    }
    sh.getRange(3, 2, 6, 2).setNumberFormat("0.000");
    sh.getRange(1, 1).setFontWeight("bold");
    sh.getRange(2, 1, 1, 5).setFontWeight("bold").setBackground("#D6EAF8");
    logGE_("BACKTEST", "Backtest concluído para " + (dias.length - 5) + " dias");
}

function rodarBacktestMenuGE() {
    rodarBacktestGE();
    const ss = SpreadsheetApp.getActive();
    ss.setActiveSheet(ss.getSheetByName(CFG.SHEET_BACKTEST));
    SpreadsheetApp.getUi().alert("✅ Backtest concluído — veja a aba Backtest.");
}

// --- helpers ---
function matrizDiariaDaBaseGE_() {
    const base = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BASE);
    if (!base || base.getLastRow() < 2) return [];
    const vals = base.getRange(2, 1, base.getLastRow() - 1, 9).getValues();
    const porDia = {};
    for (const r of vals) {
        const dataStr = r[0] instanceof Date
            ? Utilities.formatDate(r[0], CFG.TIMEZONE, "dd/MM/yyyy") : String(r[0] || "").trim();
        const data = parseDateStr_(dataStr);
        if (!data) continue;
        const minuto = minutoBaseGE_(r[2]);
        if (minuto === null) continue;
        if (!porDia[dataStr]) porDia[dataStr] = { dataStr: dataStr, data: data,
            contexto: String(r[7] || ""), minutos: new Array(60).fill(null) };
        if (typeof r[3] === "number" && isFinite(r[3])) {
            porDia[dataStr].minutos[minuto] = r[3];
        }
    }
    const dias = Object.values(porDia).filter(d =>
        d.minutos.filter(v => v !== null).length >= 20);
    dias.sort((a, b) => a.data - b.data);
    return dias;
}

function mediaPorMinutoGE_(dias) {
    const out = new Array(60).fill(null);
    for (let m = 0; m < 60; m++) {
        const vs = dias.map(d => d.minutos[m]).filter(v => typeof v === "number");
        if (vs.length) out[m] = mediaArrGE_(vs);
    }
    return out;
}

function mediaValidaGE_(arr) {
    const vs = arr.filter(v => typeof v === "number" && isFinite(v));
    return vs.length ? mediaArrGE_(vs) : null;
}

function medianaGE_(arr) {
    const s = arr.slice().sort((a, b) => a - b);
    const mid = Math.floor(s.length / 2);
    return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

// A coluna HORA da Base pode voltar como string "18:MM" ou como valor de
// hora (Date) — o Sheets converte na gravação.
function minutoBaseGE_(v) {
    if (v instanceof Date) {
        return v.getHours() === 18 ? v.getMinutes() : null;
    }
    const m = String(v || "").trim().match(/^18:([0-5]\d)/);
    return m ? parseInt(m[1], 10) : null;
}

// =============================================================================
// BACKTEST DE CHECKPOINTS (V3) — mede o ERRO DO CAMINHO FUTURO, não só o final
// =============================================================================
// Em k minutos reais conhecidos, o modelo prevê os minutos FUTUROS (k..59) —
// é isso que acontece ao vivo. Base = curva do mesmo dia da semana só com dias
// ANTERIORES (sem vazamento). Fatores de tema/clube vêm das tabelas persistidas
// (48) — leve vazamento (tabelas usam todo o histórico), sinalizado no relatório.

var BT_CP = { CHECKPOINTS: [10, 20, 30, 40], EWMA_MV: 15, DOW_MIN: 8, CONF_N: 8 };

// Matriz diária com temas por minuto (para aplicar fatores no backtest)
function matrizDiariaComTemaGE_() {
    const base = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_BASE);
    if (!base || base.getLastRow() < 2) return [];
    const vals = base.getRange(2, 1, base.getLastRow() - 1, 9).getValues();
    const porDia = {};
    for (const r of vals) {
        const dataStr = r[0] instanceof Date
            ? Utilities.formatDate(r[0], CFG.TIMEZONE, "dd/MM/yyyy") : String(r[0] || "").trim();
        const data = parseDateStr_(dataStr);
        if (!data) continue;
        const minuto = minutoBaseGE_(r[2]);
        if (minuto === null) continue;
        if (!porDia[dataStr]) porDia[dataStr] = { dataStr: dataStr, data: data,
            minutos: new Array(60).fill(null), temas: new Array(60).fill("") };
        if (typeof r[3] === "number" && isFinite(r[3])) porDia[dataStr].minutos[minuto] = r[3];
        porDia[dataStr].temas[minuto] = String(r[5] || "").trim().toUpperCase();
    }
    const dias = Object.values(porDia).filter(d =>
        d.minutos.filter(v => v !== null).length >= 20);
    dias.sort((a, b) => a.data - b.data);
    return dias;
}

// Curva base EWMA dia-da-semana a partir de um conjunto de dias já filtrado
function baseCurvaBTGE_(diasPassados, weekday) {
    let conj = diasPassados.filter(d => d.data.getDay() === weekday);
    if (conj.length < BT_CP.DOW_MIN) conj = diasPassados;
    const base = new Array(60).fill(null);
    const N = conj.length;
    for (let m = 0; m < 60; m++) {
        const vals = [];
        for (let i = 0; i < N; i++) {
            const v = conj[i].minutos[m];
            if (typeof v === "number" && v > 0) vals.push({ v: v, idade: (N - 1) - i });
        }
        if (!vals.length) continue;
        const arr = vals.map(x => x.v).sort((a, b) => a - b);
        const lo = medianaGE_(arr) * 0 + arr[Math.floor(0.05 * (arr.length - 1))];
        const hi = arr[Math.ceil(0.95 * (arr.length - 1))];
        let sw = 0, swv = 0;
        for (const x of vals) {
            const vv = Math.max(lo, Math.min(hi, x.v));
            const w = Math.pow(0.5, x.idade / BT_CP.EWMA_MV);
            sw += w; swv += w * vv;
        }
        base[m] = sw > 0 ? swv / sw : null;
    }
    return base;
}

// Projeta minutos futuros a partir de k reais conhecidos (patamar+tend+decay)
// aplicarFatores: função(minuto, tema)→fator multiplicativo (ou null p/ nenhum)
function projetarBTGE_(base, reaisConhecidos, temas, k, aplicarFatores) {
    const betas = m => m <= 20 ? 0.9478 : (m <= 45 ? 0.9677 : 0.9552);
    const ratios = [];
    for (let m = 0; m < k; m++) {
        if (typeof reaisConhecidos[m] === "number" && base[m] > 0) ratios.push(reaisConhecidos[m] / base[m]);
    }
    let patamar = 1;
    if (ratios.length) {
        const W = Math.min(1, ratios.length / BT_CP.CONF_N);
        patamar = W * medianaGE_(ratios) + (1 - W) * 1;
    }
    let tend = 1;
    if (ratios.length >= 3) {
        const m3 = mediaArrGE_(ratios.slice(-3)), m8 = mediaArrGE_(ratios.slice(-8));
        if (m8 > 0) tend = clampGE_(m3 / m8, 0.85, 1.20);
    }
    let dev = 0;
    if (k > 0 && base[k - 1] > 0 && typeof reaisConhecidos[k - 1] === "number") {
        const prevNo = base[k - 1] * patamar * tend;
        if (prevNo > 0) dev = reaisConhecidos[k - 1] / prevNo - 1;
    }
    const proj = new Array(60).fill(null);
    let d = dev;
    for (let m = k; m < 60; m++) {
        if (!base[m] || base[m] <= 0) continue;
        d = d * betas(m);
        let f = 1;
        if (aplicarFatores) { const ff = aplicarFatores(m, temas[m]); if (ff) f = ff; }
        proj[m] = base[m] * patamar * tend * (1 + d) * f;
    }
    return proj;
}

function rodarBacktestCheckpointsGE() {
    const sh = ensureAbaGE_("Backtest_CP");
    const dias = matrizDiariaComTemaGE_();
    sh.clearContents();
    if (dias.length < 10) {
        sh.getRange(1, 1).setValue("Checkpoints precisam de ≥10 programas fechados (há " + dias.length + ").");
        return;
    }

    // fator tema leve global (média &/MIN por tema) — reusa base histórica p/ &/MIN
    const histMin = mediaPorMinutoGE_(dias); // média por minuto (aprox base geral)
    // Modelos: A=base pura (patamar=1), B=V3 sem fatores, C=V3 com tema+clube
    // C (fatores tema/clube) não entra no checkpoint: seu efeito incremental é
    // medido honestamente no shadow log ao vivo (V3 real vs V2 real), sem o
    // vazamento inevitável de aplicar tabelas treinadas com todo o histórico.
    const nomes = { A: "A · Base dia-semana (sem patamar)", B: "B · V3 (patamar+tendência)" };
    const agg = {}; // modelo|checkpoint → {maePath, errFinal, n}

    for (let d = 5; d < dias.length; d++) {
        const alvo = dias[d];
        const passado = dias.slice(0, d);
        const base = baseCurvaBTGE_(passado, alvo.data.getDay());
        const reais = alvo.minutos;
        const mediaRealDia = mediaValidaGE_(reais);
        if (mediaRealDia === null) continue;

        for (const k of BT_CP.CHECKPOINTS) {
            const conhecidos = reais.map((v, i) => i < k ? v : null);
            const projA = base.slice(); // base pura para futuro
            for (let m = 0; m < k; m++) if (typeof reais[m] === "number") projA[m] = reais[m];
            const projB = projetarBTGE_(base, conhecidos, alvo.temas, k, null);
            const modelos = { A: mesclarConhecidosBTGE_(projA, reais, k),
                B: mesclarConhecidosBTGE_(projB, reais, k) };

            for (const mod in modelos) {
                const proj = modelos[mod];
                // erro do CAMINHO futuro (m >= k)
                let se = 0, np = 0;
                for (let m = k; m < 60; m++) {
                    if (typeof reais[m] === "number" && typeof proj[m] === "number") { se += Math.abs(reais[m] - proj[m]); np++; }
                }
                if (!np) continue;
                const maePath = se / np;
                const mfPrev = mediaValidaGE_(proj);
                const errFinal = mfPrev === null ? null : Math.abs(mfPrev - mediaRealDia);
                const key = mod + "|" + k;
                if (!agg[key]) agg[key] = { maePath: 0, errFinal: 0, n: 0 };
                agg[key].maePath += maePath;
                agg[key].errFinal += (errFinal || 0);
                agg[key].n++;
            }
        }
    }

    // Relatório
    const linhas = [["BACKTEST DE CHECKPOINTS — " + Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy HH:mm"),
        "(erro do caminho futuro + erro da média final; sem vazamento na base)", "", "", ""],
        ["MODELO", "CHECKPOINT (min reais)", "MAE CAMINHO FUTURO", "ERRO MÉDIA FINAL", "DIAS"]];
    for (const mod of ["A", "B"]) {
        for (const k of BT_CP.CHECKPOINTS) {
            const a = agg[mod + "|" + k];
            if (!a || !a.n) continue;
            linhas.push([nomes[mod], k, a.maePath / a.n, a.errFinal / a.n, a.n]);
        }
    }
    // Vencedor por checkpoint (menor MAE caminho)
    linhas.push(["", "", "", "", ""]);
    for (const k of BT_CP.CHECKPOINTS) {
        let best = null;
        for (const mod of ["A", "B"]) {
            const a = agg[mod + "|" + k]; if (!a || !a.n) continue;
            const mae = a.maePath / a.n;
            if (!best || mae < best.mae) best = { mod: nomes[mod], mae: mae };
        }
        if (best) linhas.push(["Melhor em " + k + " min", best.mod, best.mae, "", ""]);
    }
    sh.getRange(1, 1, linhas.length, 5).setValues(linhas);
    sh.getRange(1, 1).setFontWeight("bold");
    sh.getRange(2, 1, 1, 5).setFontWeight("bold").setBackground("#D6EAF8");
    sh.getRange(3, 3, linhas.length, 2).setNumberFormat("0.000");
    logGE_("BACKTEST_CP", "Checkpoints: " + (dias.length - 5) + " dias, modelos A/B/C");
}

function rodarBacktestCheckpointsMenuGE() {
    rodarBacktestCheckpointsGE();
    const ss = SpreadsheetApp.getActive();
    ss.setActiveSheet(ss.getSheetByName("Backtest_CP"));
    SpreadsheetApp.getUi().alert("✅ Backtest de checkpoints concluído — veja a aba Backtest_CP.");
}

function mesclarConhecidosBTGE_(proj, reais, k) {
    const out = proj.slice();
    for (let m = 0; m < k; m++) if (typeof reais[m] === "number") out[m] = reais[m];
    return out;
}

// -----------------------------------------------------------------------------
// Portão de promoção V2 ⇄ V3 (flag na Config)
// -----------------------------------------------------------------------------
function promoverV3GE_() {
    setConfigValueGE_("PREVISAO_PRINCIPAL", "V3", "Previsão oficial da aba Hoje (V2|V3)");
    logGE_("V3", "Previsão oficial promovida para V3");
}
function reverterParaV2GE_() {
    setConfigValueGE_("PREVISAO_PRINCIPAL", "V2", "Previsão oficial da aba Hoje (V2|V3)");
    logGE_("V3", "Previsão oficial revertida para V2");
}
function promoverOuReverterV3MenuGE() {
    const ui = SpreadsheetApp.getUi();
    const atual = String(getConfigValueGE_("PREVISAO_PRINCIPAL", "V2")).toUpperCase();
    const resp = ui.alert("Previsão oficial",
        "Oficial atual: " + atual + "\n\nSIM = promover V3 · NÃO = reverter para V2 · CANCELAR = manter",
        ui.ButtonSet.YES_NO_CANCEL);
    if (resp === ui.Button.YES) { promoverV3GE_(); prepararHojeGE(); ui.alert("✅ V3 é a previsão oficial."); }
    else if (resp === ui.Button.NO) { reverterParaV2GE_(); prepararHojeGE(); ui.alert("↩️ V2 é a previsão oficial."); }
}
