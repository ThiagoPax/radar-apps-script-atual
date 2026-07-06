// =============================================================================
// 18_ProjecaoGE.js — Projeção de audiência re-ancorada (aba GE - Histórico)
// =============================================================================
//
// O QUE FAZ:
// Na aba "Histórico", a coluna D contém a audiência minuto-a-minuto do
// programa MAIS RECENTE (linhas 3..62 = 18:00..18:59). Ao digitar a audiência
// REAL de um minuto em D, as células de D ABAIXO se recalculam como projeção
// que: segue o formato histórico, re-ancora no último real, incorpora o tema
// (coluna F) e, em caso de quebra, adota o novo patamar e o mantém até o fim.
//
// COLUNAS AUXILIARES (à direita, as últimas): MOLDE DIA | PASSO | BETA FAIXA | PROJ
//
// =============================================================================
// IDIOMA DAS FÓRMULAS = PORTUGUÊS (pt-BR), igual ao 11_Historico_GE.js
// -----------------------------------------------------------------------------
// Este projeto, na planilha em locale pt-BR, grava fórmulas em PORTUGUÊS via
// setFormulaSafe_/batchSetFormulasSafe_ (USER_ENTERED). Padrão confirmado no
// 11_Historico_GE.js, que escreve =SE(OU(...);...), =MÉDIA(...), =SEERRO(...),
// =PROCV(...), separador de argumentos ";" e decimal ",".
// Portanto AQUI TAMBÉM tudo é pt-BR:
//   IF->SE | OR->OU | ROUND->ARRED | VLOOKUP->PROCV | IFERROR->SEERRO
//   ROW->LIN | separador "," -> ";" | decimal "." -> ","
// O nome da aba de tema vem SEMPRE de CFG.GE_OUT_MEDIA_TEMA (nunca literal),
// para não ser corrompido por find/replace (ex.: "por" -> "pOU").
//
// COEFICIENTES BETA (persistência empírica do desvio, AR(1), ~105 dias):
//   18:00-18:20 -> 0,9478 | 18:21-18:45 -> 0,9677 | 18:46-18:59 -> 0,9552
// =============================================================================

var PROJ_GE = {
    FIRST_DATA_ROW: 3,
    LAST_DATA_ROW: 62,
    HEADER_ROW: 2,
    AUD_COL: 4, // coluna D
    FIRST_CLOSED_DAY_COL: 8, // coluna H (1º dia já fechado)
    BLOCK_SIZE: 4,
    HEADERS: ["MOLDE DIA", "PASSO", "BETA FAIXA", "PROJ"],
    TEMA_LOOKUP_RETURN_COL: 6,
    BETA_INICIO: "0,9478", // 18:00-18:20
    BETA_MEIO: "0,9677",   // 18:21-18:45
    BETA_FIM: "0,9552",    // 18:46-18:59
    MOLDE_MIN: "0,0001"    // piso do molde
};

// =============================================================================
// Entradas públicas (menu)
// =============================================================================

function instalarProjecaoAudienciaGE() {
    var ui = SpreadsheetApp.getUi();
    try {
        var res = instalarProjecaoAudienciaGE_();
        if (!res || !res.ok) {
            ui.alert("⚠️ Projeção de Audiência", (res && res.msg) ? res.msg : "Não foi possível instalar a projeção.", ui.ButtonSet.OK);
            return;
        }
        ui.alert("✅ Projeção de Audiência instalada",
            "Colunas auxiliares (" + PROJ_GE.HEADERS.join(", ") + ") criadas/atualizadas.\n" +
            "Células D substituídas pela projeção: " + res.escritas + "\n" +
            "Reais digitados preservados: " + res.preservadas + "\n\n" +
            "Agora digite a audiência real em D (e o tema em F) minuto a minuto. " +
            "As linhas abaixo se recalculam sozinhas.",
            ui.ButtonSet.OK);
    } catch (err) {
        ui.alert("❌ Erro ao instalar projeção", err.message, ui.ButtonSet.OK);
    }
}

// Fluxo combinado: arquiva o novo dia E instala a projeção, num clique só.
function arquivarEProjetarGE() {
    var ui = SpreadsheetApp.getUi();
    var passoArquivo = "";
    try {
        if (typeof adicionarNovoDiaManualGE === "function") {
            adicionarNovoDiaManualGE();
            passoArquivo = "Novo dia adicionado.";
        } else if (typeof arquivarProgramaAtualGazeta === "function") {
            arquivarProgramaAtualGazeta();
            passoArquivo = "Programa arquivado.";
        } else {
            passoArquivo = "(Função de arquivamento não encontrada — pulei esta etapa.)";
        }
    } catch (err) {
        passoArquivo = "Arquivamento retornou aviso: " + err.message + " (a projeção será instalada mesmo assim).";
    }

    try {
        var res = instalarProjecaoAudienciaGE_();
        if (!res || !res.ok) {
            ui.alert("⚠️ Projeção", passoArquivo + "\n\n" + ((res && res.msg) ? res.msg : "Falha ao instalar a projeção."), ui.ButtonSet.OK);
            return;
        }
        ui.alert("✅ Pronto",
            passoArquivo + "\n\n" +
            "Projeção instalada. Células D projetadas: " + res.escritas +
            " | reais preservados: " + res.preservadas + "\n\n" +
            "Digite a audiência real em D (e tema em F) minuto a minuto.",
            ui.ButtonSet.OK);
    } catch (err) {
        ui.alert("❌ Erro", passoArquivo + "\n\nErro ao instalar projeção: " + err.message, ui.ButtonSet.OK);
    }
}

// =============================================================================
// Core
// =============================================================================

function instalarProjecaoAudienciaGE_() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return { ok: false, msg: "Aba \"" + CFG.GE_OUT_HISTORICO + "\" não encontrada." };
    if (sheet.getLastRow() < PROJ_GE.LAST_DATA_ROW) {
        return { ok: false, msg: "A aba precisa ter ao menos " + PROJ_GE.LAST_DATA_ROW + " linhas." };
    }

    // 0) Limpa colunas auxiliares antigas/duplicadas/órfãs antes de recriar.
    limparColunasAuxAntigasGE_(sheet);

    // 1) Cria as 4 colunas auxiliares limpas, à direita.
    var aux = criarColunasAuxProjGE_(sheet);

    // 2) Colunas de audiência dos dias JÁ FECHADOS (H, L, P, ...), exceto o atual (D).
    var closedAudCols = listarColunasAudienciaFechadasGE_(sheet, aux.molde);

    // 3) Fórmulas das auxiliares (linhas 3..62) — TUDO pt-BR.
    var nLinhas = PROJ_GE.LAST_DATA_ROW - PROJ_GE.FIRST_DATA_ROW + 1;
    var molde = [], passo = [], beta = [], proj = [];
    var Lmolde = columnToLetter_(aux.molde);
    var Lbeta = columnToLetter_(aux.beta);
    var Lproj = columnToLetter_(aux.proj);
    var temaRef = "'" + CFG.GE_OUT_MEDIA_TEMA.replace(/'/g, "''") + "'";

    for (var r = PROJ_GE.FIRST_DATA_ROW; r <= PROJ_GE.LAST_DATA_ROW; r++) {
        var mediaBase = montarMediaBaseGE_(closedAudCols, r);
        var fatorTema = "SEERRO(PROCV(F" + r + ";" + temaRef + "!$A:$F;" + PROJ_GE.TEMA_LOOKUP_RETURN_COL + ";0);1)";
        var moldeF;
        if (mediaBase === "") {
            // Sem dias fechados: usa a média B (já robusta no 11_Historico_GE).
            moldeF = "=SEERRO(SE(OU(B" + r + '="";B' + r + "=0);" + PROJ_GE.MOLDE_MIN + ";B" + r + "*" + fatorTema + ");" + PROJ_GE.MOLDE_MIN + ")";
        } else {
            // Com dias fechados: mediaBase já é MÉDIA(SEERRO(...)) robusta;
            // o SEERRO externo garante que nada (nem #DIV/0! de média vazia)
            // propague — cai no piso MOLDE_MIN.
            moldeF = "=SEERRO(SE(OU((" + mediaBase + ')="";(' + mediaBase + ")=0);" + PROJ_GE.MOLDE_MIN + ";(" + mediaBase + ")*" + fatorTema + ");" + PROJ_GE.MOLDE_MIN + ")";
        }
        molde.push([moldeF]);

        passo.push([r === PROJ_GE.FIRST_DATA_ROW ? "=1"
            : "=SE(OU(" + Lmolde + (r - 1) + '="";' + Lmolde + (r - 1) + "=0);1;" + Lmolde + r + "/" + Lmolde + (r - 1) + ")"]);

        beta.push(["=SE(LIN()<=23;" + PROJ_GE.BETA_INICIO + ";SE(LIN()<=48;" + PROJ_GE.BETA_MEIO + ";" + PROJ_GE.BETA_FIM + "))"]);

        proj.push([r === PROJ_GE.FIRST_DATA_ROW ? "=SE(D" + r + '="";"";D' + r + ")"
            : "=SE(D" + (r - 1) + '="";"";ARRED(' + Lmolde + r + "*(1+(D" + (r - 1) + "/" + Lmolde + (r - 1) + "-1)*" + Lbeta + r + ");2))"]);
    }

    batchSetFormulasSafe_(sheet.getRange(PROJ_GE.FIRST_DATA_ROW, aux.molde, nLinhas, 1), molde);
    batchSetFormulasSafe_(sheet.getRange(PROJ_GE.FIRST_DATA_ROW, aux.passo, nLinhas, 1), passo);
    batchSetFormulasSafe_(sheet.getRange(PROJ_GE.FIRST_DATA_ROW, aux.beta, nLinhas, 1), beta);
    batchSetFormulasSafe_(sheet.getRange(PROJ_GE.FIRST_DATA_ROW, aux.proj, nLinhas, 1), proj);

    // 4) Coluna D: substitui FÓRMULAS (modelo antigo OU projeção velha) pela
    //    nova projeção; preserva NÚMEROS digitados pelo usuário SEM tocá-los
    //    (escrever um número via USER_ENTERED pt-BR poderia reinterpretar o
    //    separador decimal — por isso a célula de real NÃO é reescrita).
    var dRange = sheet.getRange(PROJ_GE.FIRST_DATA_ROW, PROJ_GE.AUD_COL, nLinhas, 1);
    var dValues = dRange.getValues();
    var dFormulasExist = dRange.getFormulas(); // "" quando não é fórmula
    var escritas = 0, preservadas = 0;

    // FIX QUOTA: a versão anterior chamava setFormulaSafe_ célula a célula
    // (até 60 calls), estourando o limite "Write requests per minute" do
    // Sheets API. Agora coletamos cada fórmula num array esparso e gravamos
    // tudo em UMA chamada batchUpdate. Reais digitados continuam intocados.
    var cellsParaProjetar = [];
    var d3LimparTudo = false;
    for (var k = 0; k < nLinhas; k++) {
        var rowNum = PROJ_GE.FIRST_DATA_ROW + k;
        var val = dValues[k][0];
        var hasFormula = dFormulasExist[k][0] !== "" && dFormulasExist[k][0] != null;
        var typedReal = (!hasFormula && typeof val === "number" && val !== 0);
        if (typedReal) { preservadas++; continue; }
        if (rowNum === PROJ_GE.FIRST_DATA_ROW) {
            d3LimparTudo = true; // D3 sem real: aguarda digitação
        } else {
            cellsParaProjetar.push({
                row: rowNum,
                col: PROJ_GE.AUD_COL,
                formula: "=SE(D" + (rowNum - 1) + '="";"";' + Lproj + rowNum + ")"
            });
            escritas++;
        }
    }
    if (d3LimparTudo) sheet.getRange(PROJ_GE.FIRST_DATA_ROW, PROJ_GE.AUD_COL).clearContent();
    if (cellsParaProjetar.length > 0) batchSetFormulasSparseSafe_(sheet, cellsParaProjetar);

    SpreadsheetApp.flush();

    // Formatação (em UM lote por coluna)
    sheet.getRange(PROJ_GE.FIRST_DATA_ROW, PROJ_GE.AUD_COL, nLinhas, 1).setNumberFormat("0.00");
    sheet.getRange(PROJ_GE.FIRST_DATA_ROW, aux.molde, nLinhas, 1).setNumberFormat("0.00");
    sheet.getRange(PROJ_GE.FIRST_DATA_ROW, aux.passo, nLinhas, 1).setNumberFormat("0.000");
    sheet.getRange(PROJ_GE.FIRST_DATA_ROW, aux.beta, nLinhas, 1).setNumberFormat("0.0000");
    sheet.getRange(PROJ_GE.FIRST_DATA_ROW, aux.proj, nLinhas, 1).setNumberFormat("0.00");

    return { ok: true, aux: aux, escritas: escritas, preservadas: preservadas };
}

// =============================================================================
// Helpers de colunas auxiliares
// =============================================================================

function limparColunasAuxAntigasGE_(sheet) {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    var headers = sheet.getRange(PROJ_GE.HEADER_ROW, 1, 1, lastCol).getValues()[0];
    for (var c = lastCol; c >= 1; c--) {
        var h = ((headers[c - 1] || "") + "").trim();
        if (PROJ_GE.HEADERS.indexOf(h) !== -1) {
            sheet.deleteColumn(c);
        }
    }
}

function criarColunasAuxProjGE_(sheet) {
    var lastCol = sheet.getLastColumn();
    var start = lastCol + 2;
    var aux = { molde: start, passo: start + 1, beta: start + 2, proj: start + 3 };
    sheet.getRange(PROJ_GE.HEADER_ROW, aux.molde, 1, 4).setValues([PROJ_GE.HEADERS]);
    sheet.getRange(PROJ_GE.HEADER_ROW, aux.molde, 1, 4)
        .setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
    sheet.setColumnWidth(aux.molde, 90);
    sheet.setColumnWidth(aux.passo, 70);
    sheet.setColumnWidth(aux.beta, 80);
    sheet.setColumnWidth(aux.proj, 90);
    return aux;
}

function listarColunasAudienciaFechadasGE_(sheet, stopBeforeCol) {
    var lastCol = sheet.getLastColumn();
    var limit = (stopBeforeCol && stopBeforeCol > 0) ? Math.min(lastCol, stopBeforeCol - 1) : lastCol;
    // Otimização: lê a linha de cabeçalhos UMA vez, em vez de getValue por coluna.
    var headers = sheet.getRange(PROJ_GE.HEADER_ROW, 1, 1, limit).getValues()[0];
    var cols = [];
    for (var c = PROJ_GE.FIRST_CLOSED_DAY_COL; c <= limit; c += PROJ_GE.BLOCK_SIZE) {
        var header = ((headers[c - 1] || "") + "").trim();
        if (header !== "AUDIÊNCIA" && header !== "AUDIENCIA") continue;
        cols.push(c);
    }
    return cols;
}

// Monta a média-base do minuto r a partir das colunas de dias fechados.
// ROBUSTO: cada parcela é envolvida em SEERRO(...;"") para que uma célula
// histórica com #REF!/#N/D não derrube a soma inteira. Usa MÉDIA (ignora
// vazios), então o divisor se ajusta sozinho ao número de dias válidos —
// melhor que dividir por um N fixo quando há buracos no histórico.
// Retorna "" se não houver colunas.
function montarMediaBaseGE_(closedAudCols, r) {
    if (!closedAudCols || closedAudCols.length === 0) return "";
    var parts = [];
    for (var i = 0; i < closedAudCols.length; i++) {
        parts.push('SEERRO(' + columnToLetter_(closedAudCols[i]) + r + ';"")');
    }
    return "MÉDIA(" + parts.join(";") + ")";
}

// =============================================================================
// Reconstrução da cadeia (chamada pelo onEdit em 15_Menu.js)
// =============================================================================
// Mantém a coluna D abaixo da linha editada com a fórmula de projeção, sempre
// que estiver vazia (preserva reais digitados). Em pt-BR (igual ao resto).
function reconstruirCadeiaProjecaoGE_(sheet, editedRow) {
    if (!sheet || sheet.getName() !== CFG.GE_OUT_HISTORICO) return;

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(PROJ_GE.HEADER_ROW, 1, 1, lastCol).getValues()[0];
    var projIdx = -1;
    for (var c = 0; c < headers.length; c++) {
        if (((headers[c] || "") + "").trim() === "PROJ") { projIdx = c + 1; break; }
    }
    if (projIdx < 0) return;
    var Lproj = columnToLetter_(projIdx);

    var startRow = Math.max(editedRow + 1, PROJ_GE.FIRST_DATA_ROW);
    if (startRow > PROJ_GE.LAST_DATA_ROW) return;

    var n = PROJ_GE.LAST_DATA_ROW - startRow + 1;
    var rng = sheet.getRange(startRow, PROJ_GE.AUD_COL, n, 1);
    var vals = rng.getValues();
    var fxs = rng.getFormulas();

    for (var i = 0; i < n; i++) {
        var rowNum = startRow + i;
        var hasFormula = fxs[i][0] !== "" && fxs[i][0] != null;
        var typedReal = (!hasFormula && typeof vals[i][0] === "number" && vals[i][0] !== 0);
        if (typedReal) continue; // preserva real digitado
        // IMPORTANTE: usar setFormulaSafe_ (USER_ENTERED) — o MESMO caminho do
        // resto do projeto, que grava pt-BR corretamente. NÃO usar o setFormula/
        // setFormulas NATIVO aqui: o método nativo espera sintaxe en-US e geraria
        // #NAME? ao receber "SE" em português (foi o que causava o erro nas
        // linhas abaixo da editada, segundos após a digitação).
        setFormulaSafe_(sheet.getRange(rowNum, PROJ_GE.AUD_COL),
            "=SE(D" + (rowNum - 1) + '="";"";' + Lproj + rowNum + ")");
    }
}