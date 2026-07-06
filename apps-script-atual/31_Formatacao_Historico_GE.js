// =============================================================================
// 31_Formatacao_Historico_GE.js — Formatação condicional das colunas &/MIN
// =============================================================================
// As regras antigas estavam presas a blocos fixos; os programas novos ficavam
// sem escala de cor. Esta rotina identifica TODAS as colunas com cabeçalho
// "&/MIN" na linha 2 e aplica uma única regra de gradiente cobrindo todas:
// vermelho (pior) → branco (índice 1) → verde (melhor).
// Reaplicada: no arquivamento de novo dia, no "Rodar tudo agora" e pelo item
// de menu "Corrigir formatação do Histórico".

function corrigirFormatacaoHistoricoGE() {
    const hist = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!hist) {
        logGE_("FORMATACAO", "Aba Histórico não encontrada");
        return 0;
    }

    // Colunas &/MIN: cabeçalho na linha 2 (inclui a col E do dia corrente e
    // as dos blocos fechados)
    const lastCol = hist.getLastColumn();
    const headers = hist.getRange(2, 1, 1, lastCol).getValues()[0];
    const ranges = [];
    for (let c = 0; c < headers.length; c++) {
        const h = String(headers[c] || "").replace(/\s/g, "").toUpperCase();
        if (h === "&/MIN") ranges.push(hist.getRange(3, c + 1, 60, 1));
    }
    if (!ranges.length) {
        logGE_("FORMATACAO", "Nenhuma coluna &/MIN encontrada na linha 2");
        return 0;
    }

    // Remove regras antigas que atinjam qualquer coluna &/MIN e aplica a nova
    const colunasAlvo = new Set(ranges.map(r => r.getColumn()));
    const regras = hist.getConditionalFormatRules().filter(rule =>
        !rule.getRanges().some(r => {
            for (let c = r.getColumn(); c < r.getColumn() + r.getNumColumns(); c++) {
                if (colunasAlvo.has(c)) return true;
            }
            return false;
        }));

    regras.push(SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue("#E74C3C", SpreadsheetApp.InterpolationType.NUMBER, "0.75")
        .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "1")
        .setGradientMaxpointWithValue("#27AE60", SpreadsheetApp.InterpolationType.NUMBER, "1.25")
        .setRanges(ranges)
        .build());
    hist.setConditionalFormatRules(regras);

    logGE_("FORMATACAO", "Gradiente aplicado a " + ranges.length + " colunas &/MIN");
    return ranges.length;
}

// Entrada de menu com feedback visual
function corrigirFormatacaoHistoricoMenuGE() {
    const n = corrigirFormatacaoHistoricoGE();
    SpreadsheetApp.getUi().alert(n > 0
        ? "✅ Formatação condicional reaplicada em " + n + " colunas &/MIN."
        : "⚠️ Nenhuma coluna &/MIN encontrada na linha 2 do Histórico.");
}
