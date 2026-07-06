// =============================================================================
// 02_Schema_GE.js — Criação/validação de abas, Config (chave/valor) e Logs
// =============================================================================

// Garante que uma aba exista; se headers forem passados e a aba for nova,
// escreve a linha de cabeçalho e congela.
function ensureAbaGE_(nome, headers) {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(nome);
    if (!sh) {
        sh = ss.insertSheet(nome);
        if (headers && headers.length) {
            sh.getRange(1, 1, 1, headers.length).setValues([headers])
                .setFontWeight("bold").setBackground("#34495E").setFontColor("#FFFFFF");
            sh.setFrozenRows(1);
        }
    }
    return sh;
}

// -----------------------------------------------------------------------------
// Logs — registro persistente de eventos das rotinas (aba Logs)
// -----------------------------------------------------------------------------
function logGE_(evento, detalhe) {
    try {
        const sh = ensureAbaGE_(CFG.SHEET_LOGS, ["QUANDO", "EVENTO", "DETALHE"]);
        sh.insertRowAfter(1);
        sh.getRange(2, 1, 1, 3).setValues([[
            Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy HH:mm:ss"),
            String(evento || ""), String(detalhe || "").substring(0, 5000)
        ]]);
        // Mantém no máximo ~2000 linhas de log
        const last = sh.getLastRow();
        if (last > 2001) sh.deleteRows(2002, last - 2001);
    } catch (e) {
        Logger.log("logGE_ falhou: " + e.message);
    }
    Logger.log("[" + evento + "] " + detalhe);
}

// -----------------------------------------------------------------------------
// Config — tabela chave/valor na aba Config (col A = chave, col B = valor,
// col C = descrição). Usada para contexto ativo e parâmetros editáveis.
// -----------------------------------------------------------------------------
function getConfigSheetGE_() {
    const sh = ensureAbaGE_(CFG.SHEET_CONFIG, ["CHAVE", "VALOR", "DESCRIÇÃO"]);
    if (sh.getLastRow() < 2) {
        sh.getRange(2, 1, 8, 3).setValues([
            ["CONTEXTO_ATIVO", "COPA DO MUNDO", "Contexto de comparação: " + CFG.CONTEXTOS_VALIDOS.join(" | ")],
            ["CONTEXTO_DESDE", Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy"), "Início da vigência do contexto ativo"],
            ["MERCHANS_PREVISTOS", "2", "Quantos merchans planejar no espelho (2 a 5)"],
            ["LIMIAR_RUPTURA", "0.20", "Desvio (fração) além do qual a leitura vira RUPTURA"],
            ["PREVISAO_PRINCIPAL", "V2", "Previsão oficial da aba Hoje: V2 (estável) ou V3 (experimental)"],
            ["EWMA_MEIA_VIDA", "15", "V3: meia-vida (programas) da média histórica ponderada por recência"],
            ["FATOR_CLUBE_N_MINIMO", "40", "V3: amostra mínima (minutos) para usar uma célula do fator clube"],
            ["FATOR_GLOBO_BREAK_ATIVO", "NAO", "V3: ligar efeito Globo-em-break (só após provar no backtest)"]
        ]);
    }
    return sh;
}

function getConfigValueGE_(chave, valorPadrao) {
    try {
        const sh = getConfigSheetGE_();
        const vals = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
        for (const [k, v] of vals) {
            if (String(k).trim().toUpperCase() === String(chave).trim().toUpperCase()) {
                return (v === "" || v === null) ? valorPadrao : v;
            }
        }
    } catch (e) { /* usa padrão */ }
    return valorPadrao;
}

function setConfigValueGE_(chave, valor, descricao) {
    const sh = getConfigSheetGE_();
    const last = Math.max(sh.getLastRow(), 1);
    const vals = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
    for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim().toUpperCase() === String(chave).trim().toUpperCase()) {
            sh.getRange(i + 2, 2).setValue(valor);
            if (descricao) sh.getRange(i + 2, 3).setValue(descricao);
            return;
        }
    }
    sh.getRange(last + 1, 1, 1, 3).setValues([[chave, valor, descricao || ""]]);
}
