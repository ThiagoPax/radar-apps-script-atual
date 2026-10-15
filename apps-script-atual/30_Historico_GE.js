// =============================================================================
// 11_Historico_GE.js — Histórico, dashboards e médias (Gazeta Esportiva)
// =============================================================================
//
// Mudanças nesta versão:
//
//   PREVISÃO INICIAL
//   • gerarPrevisaoInicialGE_ — previsão isolada a partir de histórico real,
//     sem a circularidade GS×GY×HC do modelo anterior (D vazio → GY≈1).
//   • atualizarFormulasHistoricoGE_ refatorada via atualizarFormulasBaseGE_.
//   • Ordem de execução em adicionarNovoDiaManualGE corrigida.
//
//   ÁREA PROTEGIDA A66:NT89 — três correções estruturais
//   • preencherBlocoA66E85HistoricoGE_ reescrita: agora dinâmica (baseada em
//     datas reais da row 1), sem fórmulas hardcoded, sem clearFormat/
//     clearDataValidations/clearNote destrutivos. Limpa F66:NT89 (polução
//     lateral) e preserva D86:D89 (valores externos do user).
//   • atualizarMediaSemanaGE_ não escreve mais em (66, sexta.col) e
//     (67, sexta.col) — média semanal vive integralmente em A66:E89.
//   • atualizarGraficoMediasGE_ usa DATA_START_ROW = 95 (fora da área
//     protegida); limpa preventivamente a área legada A70:D130.
//   • garantirCamposHistoricoGE_ garante mínimo de 160 rows para acomodar
//     dados auxiliares do gráfico em 95-155.
// =============================================================================

// --- Helpers de data e detecção de colunas ---

function parseDateStr_(s) {
    if (!s || typeof s !== "string") return null;
    var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    return m ? new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])) : null;
}

// Retorna apenas as colunas de audiência reais (com data no row 1), excluindo CALIBRAGEM
function getAudColsGE_(sheet) {
    var lastCol = sheet.getLastColumn();
    var audCols = [];
    // Uma leitura em bloco do header inteiro — o loop antigo fazia um getValue
    // por coluna (~130 idas à planilha por chamada, várias vezes por minuto).
    var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var c = 4; c <= lastCol; c += 4) {
        var h = header[c - 1];
        var hasDate = (h instanceof Date) ||
            (typeof h === "string" && /^\d{2}\/\d{2}\/\d{4}$/.test(h.trim()));
        if (hasDate) audCols.push(c);
        else break;
    }
    return audCols;
}

// =============================================================================
// MODELO DE PREVISÃO INICIAL — constantes e helpers
// =============================================================================
// Estas constantes governam o modelo isolado que gera a previsão de audiência
// do novo programa antes de existirem dados reais em D. O modelo é construído
// puramente a partir das colunas históricas (H em diante) e NÃO usa GS-GY,
// evitando a circularidade do modelo anterior.
// -----------------------------------------------------------------------------
var PREV_INICIAL_GE = {
    PESO_DIA: 0.40,            // peso do fator dia-da-semana no regime normal
    PESO_TREND: 0.35,          // peso do fator tendência (últimos 5)
    PESO_ACCEL: 0.25,          // peso do fator aceleração (últimos 3 / últimos 5)
    PESO_DIA_RUPT: 0.20,       // peso do fator dia em regime de ruptura
    PESO_TREND_RUPT: 0.55,     // peso do fator tendência em regime de ruptura
    PESO_ACCEL_RUPT: 0.25,     // peso do fator aceleração em regime de ruptura
    BLEND_RUPTURA: 0.50,       // mistura 50/50 regime normal × ruptura
    CLAMP_TREND: [0.75, 1.30], // limites do fator tendência
    CLAMP_ACCEL: [0.85, 1.18], // limites do fator aceleração
    CLAMP_DIA:   [0.78, 1.28], // limites do fator dia da semana
    MIN_AMOSTRAS_TREND: 3,
    MIN_AMOSTRAS_ACCEL: 2,
    MIN_AMOSTRAS_DIA:   2,
    MIN_AMOSTRAS_RUPTURA: 6,
    THRESHOLD_RUPTURA: 0.15,
    N_RECENT_5: 5,
    N_RECENT_3: 3
};

function clampGE_(v, lo, hi) {
    if (typeof v !== "number" || isNaN(v) || !isFinite(v)) return 1;
    return Math.max(lo, Math.min(hi, v));
}

function mediaArrGE_(arr) {
    if (!arr || arr.length === 0) return 0;
    var s = 0;
    for (var i = 0; i < arr.length; i++) s += arr[i];
    return s / arr.length;
}


// =============================================================================
// GRAVAÇÃO DE FÓRMULAS SEM SHEETS API AVANÇADA
// =============================================================================
// Evita chamadas a Sheets.Spreadsheets.Values.update, que exigem ativar a
// Google Sheets API no projeto GCP. Usamos SpreadsheetApp, que grava a
// fórmula direto no arquivo.
//
// ATENÇÃO — locale: diferente de Values.update com USER_ENTERED, o
// SpreadsheetApp (setValue/setValues/setFormula) interpreta a string SEMPRE
// na sintaxe canônica en-US (IF, ROUND, vírgula como separador e ponto
// decimal), independente do locale da planilha. Gravar "=SE(...;...)" por
// esse caminho armazena uma função desconhecida "SE" → #NAME?, e o pior:
// na interface pt-BR a fórmula quebrada é exibida exatamente igual à
// versão correta digitada à mão. Por isso TODA fórmula passa por
// localizarFormulaGE_ antes da gravação.
// -----------------------------------------------------------------------------
function localizarFormulaGE_(formula) {
    var f = String(formula);
    if (f.charAt(0) !== "=") return f; // não é fórmula
    try {
        return pv2LocalizeFormula_(f, pv2DetectFormulaDialect_());
    } catch (e) {
        // fallback: canônico en-US, aceito pelo parser em qualquer locale
        return pv2LocalizeFormula_(f, "EN_COMMA");
    }
}

function setFormulaSafeNoApiGE_(range, formula) {
    if (formula === null || formula === undefined || formula === "") {
        range.clearContent();
        return;
    }
    var f = String(formula);
    if (f.charAt(0) !== "=") f = "=" + f;
    range.setFormula(localizarFormulaGE_(f));
}

function setFormulasUserEnteredNoApiGE_(range, values) {
    if (!values || values.length === 0) return;
    var out = new Array(values.length);
    for (var i = 0; i < values.length; i++) {
        var row = new Array(values[i].length);
        for (var j = 0; j < values[i].length; j++) {
            var v = values[i][j];
            row[j] = (typeof v === "string" && v.charAt(0) === "=") ? localizarFormulaGE_(v) : v;
        }
        out[i] = row;
    }
    range.setValues(out);
}

function batchSetFormulasSafeNoApiGE_(range, formulas) {
    setFormulasUserEnteredNoApiGE_(range, formulas);
}

// --- Reparo one-shot de fórmulas quebradas (#NAME?) ---
// Corrige células que ficaram com nome de função pt-BR armazenado literalmente
// (SE, ARRED, MÉDIA, ...) por terem sido gravadas via setValue/setValues sem
// tradução para a sintaxe canônica. Varre todas as abas e reescreve, via
// localizarFormulaGE_, apenas células cujo valor exibido é o erro #NAME? —
// células corretas nunca são tocadas. Executável pelo menu ou pelo editor.
function repararFormulasQuebradasGE() {
    var ss = SpreadsheetApp.getActive();
    var sheets = ss.getSheets();
    var fixed = 0, kept = 0;
    for (var s = 0; s < sheets.length; s++) {
        var sheet = sheets[s];
        var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
        if (lastRow === 0 || lastCol === 0) continue;
        var rng = sheet.getRange(1, 1, lastRow, lastCol);
        var formulas = rng.getFormulas();
        var display = rng.getDisplayValues();
        for (var i = 0; i < formulas.length; i++) {
            for (var j = 0; j < formulas[i].length; j++) {
                var f = formulas[i][j];
                if (!f) continue;
                if (display[i][j] !== "#NAME?") continue;
                var novo = localizarFormulaGE_(f);
                if (novo !== f) {
                    sheet.getRange(i + 1, j + 1).setFormula(novo);
                    fixed++;
                } else {
                    kept++; // #NAME? que a tradução não altera (causa diversa)
                }
            }
        }
    }
    Logger.log("[Reparo #NAME?] corrigidas: " + fixed + " | não traduzíveis: " + kept);
    try { logGE_("REPARO", fixed + " fórmulas #NAME? re-localizadas"); } catch (e) {}
    return fixed;
}

function repararFormulasQuebradasMenuGE() {
    var n = repararFormulasQuebradasGE();
    try {
        SpreadsheetApp.getUi().alert("🔧 Reparo concluído", n + " fórmulas #NAME? corrigidas.", SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
}

// --- Saneamento one-shot: remove projeções presas nas colunas históricas ---
// Minutos sem captura arquivados COM a fórmula de projeção viram #REF!
// (referência circular) e contaminam B/C/E e a calibragem. Esta função limpa
// as fórmulas presas nas colunas históricas de AUDIÊNCIA (a do dia corrente,
// D, fica intacta), regenera as fórmulas-base B/C/63 e a calibragem, e
// repreenche a MÉDIA ESPERADA e a média histórica da aba Hoje.
function sanearProjecoesHistoricoGE() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) throw new Error('Aba "' + CFG.GE_OUT_HISTORICO + '" não encontrada.');

    var audCols = getAudColsGE_(sheet);
    var limpas = 0;
    for (var i = 0; i < audCols.length; i++) {
        var col = audCols[i];
        if (col === 4) continue; // dia corrente: projeções em D são legítimas
        var forms = sheet.getRange(3, col, 60, 1).getFormulas();
        for (var r = 0; r < 60; r++) {
            if (forms[r][0]) {
                sheet.getRange(3 + r, col).clearContent();
                limpas++;
            }
        }
    }

    atualizarFormulasBaseGE_(sheet);
    atualizarCalibragemGE_(sheet);
    atualizarFatorTemaGE_(sheet);
    SpreadsheetApp.flush();

    // Repreenche as réguas da aba Hoje, que herdaram os buracos do #REF!
    try {
        var hoje = ss.getSheetByName(CFG.SHEET_HOJE);
        if (hoje) {
            hoje.getRange(2, 2).setValue(HOJE_GE.HEADERS[1]); // "MÉDIA ESPERADA"
            preencherMediaContextualHojeGE_(hoje);
            preencherMediaHistoricaHojeGE_(hoje);
        }
    } catch (e) {
        Logger.log("[Saneamento] Réguas da Hoje: " + e);
    }

    var resumo = limpas + " projeções presas removidas das colunas históricas; " +
        "fórmulas-base e calibragem regeneradas.";
    Logger.log("[Saneamento] " + resumo);
    try { logGE_("SANEAMENTO", resumo); } catch (e) {}
    return resumo;
}

function sanearProjecoesHistoricoMenuGE() {
    var resumo = sanearProjecoesHistoricoGE();
    try {
        SpreadsheetApp.getUi().alert("🩹 Saneamento concluído", resumo, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
}

function batchSetFormulasSparseSafeNoApiGE_(sheet, entries) {
    if (!entries || entries.length === 0) return;
    for (var i = 0; i < entries.length; i++) {
        var item = entries[i];
        if (!item || !item.row || !item.col) continue;
        setFormulaSafeNoApiGE_(sheet.getRange(item.row, item.col), item.formula);
    }
}

// --- Arquivar programa GE (entrada pública via menu) ---

function arquivarProgramaAtualGazeta() {
    const e = SpreadsheetApp.getUi();
    if (e.alert("\u{1f4c1} Arquivar Programa", "Deseja arquivar o programa atual do Gazeta Esportiva no hist\xf3rico?", e.ButtonSet.YES_NO) === e.Button.YES) try {
        const t = extrairDataProgramaGE_() || new Date;
        if (arquivarProgramaGE_(t)) {
            const o = Utilities.formatDate(t, "America/Sao_Paulo", "dd/MM/yyyy");
            e.alert("\u2705 Arquivado!", `Programa de ${o} arquivado no hist\xf3rico do Gazeta.`, e.ButtonSet.OK)
        } else e.alert("\u2139\ufe0f J\xe1 existe", "Este programa j\xe1 est\xe1 arquivado no hist\xf3rico.", e.ButtonSet.OK)
    } catch (t) {
        e.alert("\u274c Erro", "Erro ao arquivar: " + t.message, e.ButtonSet.OK)
    }
}

// --- Criação / inicialização da aba Histórico GE ---

function ensureHistoricoSheetGE_() {
    const e = SpreadsheetApp.getActive();
    let t = e.getSheetByName(CFG.GE_OUT_HISTORICO);
    const o = !t;
    t || (t = e.insertSheet(CFG.GE_OUT_HISTORICO)), t.getRange(1, 1).setValue("").setBackground("#E74C3C"), t.getRange(1, 2).setValue("").setBackground("#E74C3C"), t.getRange(1, 3).setValue("").setBackground("#E74C3C"), t.getRange(2, 1).setValue("HORA").setFontWeight("bold").setBackground("#E74C3C").setFontColor("white"), t.getRange(2, 2).setValue("M\xc9DIA").setFontWeight("bold").setBackground("#E74C3C").setFontColor("white"), t.getRange(2, 3).setValue("% A M\xc9DIA").setFontWeight("bold").setBackground("#E74C3C").setFontColor("white");
    const a = [];
    for (let e = 0; e <= 59; e++) a.push(["18:" + (e + "").padStart(2, "0")]);
    return t.getRange(3, 1, 60, 1).setValues(a), t.setColumnWidth(1, 60), t.setColumnWidth(2, 70), t.setColumnWidth(3, 90), t.setFrozenColumns(3), t.setFrozenRows(2), o && Logger.log(`Aba "${CFG.GE_OUT_HISTORICO}" criada com estrutura correta.`), t
}

function localizarBlocoHistoricoGEPorData_(programDate) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet || !programDate) return null;
    const target = Utilities.formatDate(programDate, CFG.TIMEZONE, "dd/MM/yyyy");
    const lastCol = sheet.getLastColumn();
    if (lastCol <= 3) return null;
    const headers = sheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
    for (let i = 0; i < headers.length; i += 4) {
        const value = headers[i];
        if (!value) continue;
        const dateStr = value instanceof Date ? Utilities.formatDate(value, CFG.TIMEZONE, "dd/MM/yyyy") : value.toString().trim();
        if (dateStr === target) return {
            sheet: sheet,
            dateStr: dateStr,
            startCol: 4 + i,
            audienciaCol: 4 + i,
            shareCol: 5 + i,
            temaCol: 6 + i,
            termosCol: 7 + i
        };
    }
    return null;
}

function obterTermosMinutoAMinutoGE_() {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_TIME);
    if (!sheet || sheet.getLastRow() < 61) throw Error('A aba "' + CFG.GE_OUT_TIME + '" não possui o intervalo E2:E61 completo.');
    return sheet.getRange(2, 5, 60, 1).getValues();
}

function sincronizarTermosMinutoAMinutoParaHistoricoGE_(programDate) {
    const block = localizarBlocoHistoricoGEPorData_(programDate);
    if (!block) return Logger.log("[Speech] Histórico GE sem bloco para " + Utilities.formatDate(programDate, CFG.TIMEZONE, "dd/MM/yyyy")), {
        status: "missing_block"
    };
    const termos = obterTermosMinutoAMinutoGE_();
    block.sheet.getRange(3, block.termosCol, 60, 1).setValues(termos);
    Logger.log("[Speech] TERMOS sincronizados para GE - Histórico em " + block.dateStr + " (coluna " + block.termosCol + ").");
    return {
        status: "updated",
        dateStr: block.dateStr,
        termosCol: block.termosCol
    };
}

// --- Arquivar programa GE (interno, chamado automaticamente pelo runAllGazeta) ---

function arquivarProgramaGE_(programDate) {
    var ss = SpreadsheetApp.getActive();
    var hist = ensureHistoricoSheetGE_();
    var timeSheet = ss.getSheetByName(CFG.GE_OUT_TIME);
    if (!timeSheet || timeSheet.getLastRow() < 2) throw Error("Aba GE - Minuto a Minuto está vazia");

    var dateStr = Utilities.formatDate(programDate, "America/Sao_Paulo", "dd/MM/yyyy");

    // Verificar duplicata
    var lastCol = hist.getLastColumn();
    if (lastCol > 3) {
        var headers = hist.getRange(1, 4, 1, lastCol - 3).getValues()[0];
        for (var i = 0; i < headers.length; i++) {
            var h = headers[i];
            if (!h) continue;
            var hStr = h instanceof Date
                ? Utilities.formatDate(h, "America/Sao_Paulo", "dd/MM/yyyy")
                : h.toString().trim();
            if (hStr === dateStr) {
                Logger.log("Programa GE de " + dateStr + " já existe - pulando");
                return false;
            }
        }
    }

    // Inserir 4 colunas
    hist.insertColumns(4, 4);

    // FIX: a inserção empurra o bloco mensal do painel (D:E) para H:I.
    // Devolve imediatamente para D:E, antes de qualquer outra escrita.
    restaurarBlocoMensalAposInsercaoGE_(hist);

    var numExisting = Math.floor((lastCol - 3) / 4);
    var bgColor = (numExisting % 2 === 0) ? "#FDF2F2" : "#FFFFFF";

    hist.getRange(1, 4, 1, 4).merge();
    hist.getRange(1, 4).setValue(dateStr)
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setBackground("#FADBD8")
        .setFontColor("#000000");

    hist.getRange(2, 4, 1, 4).setValues([["AUDIÊNCIA", "&/MIN", "TEMA", "TERMOS"]]);
    hist.getRange(2, 4, 1, 4)
        .setFontWeight("bold")
        .setBackground("#E74C3C")
        .setFontColor("white")
        .setHorizontalAlignment("center");

    hist.setColumnWidth(4, 90);
    hist.setColumnWidth(5, 70);
    hist.setColumnWidth(6, 120);
    hist.setColumnWidth(7, 200);

    hist.getRange(3, 4, 60, 4).setBackground(bgColor);

    // Ler dados do Minuto a Minuto
    var lastRow = timeSheet.getLastRow();
    var numCols = Math.min(6, timeSheet.getLastColumn());
    var data = timeSheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var minuteMap = new Map();
    for (var r = 0; r < data.length; r++) {
        var rowData = data[r];
        var timeStr = "";
        if (rowData[0] instanceof Date) {
            var dt = rowData[0];
            timeStr = (dt.getHours() + "").padStart(2, "0") + ":" + (dt.getMinutes() + "").padStart(2, "0");
        } else if (rowData[0]) {
            var s = rowData[0].toString().trim();
            var m = s.match(/(\d{1,2}):(\d{2})/);
            timeStr = m ? (parseInt(m[1]) + "").padStart(2, "0") + ":" + m[2] : s;
        }
        var tema = rowData[1] ? rowData[1].toString().trim() : "";
        var termos = rowData[4] ? rowData[4].toString().trim() : "";
        if (timeStr) minuteMap.set(timeStr, { tema: tema, termos: termos });
    }

    var values = [];
    for (var min = 0; min <= 59; min++) {
        var t = "18:" + (min + "").padStart(2, "0");
        var entry = minuteMap.get(t);
        values.push(["", "", entry ? entry.tema : "", entry ? entry.termos : ""]);
    }
    hist.getRange(3, 4, values.length, 4).setValues(values);

    hist.getRange(3, 4, 60, 1).setNumberFormat("0.00");
    hist.getRange(3, 5, 60, 1).setNumberFormat("0.00");

    var shareFormulas = [];
    for (var row = 3; row <= 62; row++) {
        shareFormulas.push(['=SE(OU(D' + row + '="";$B' + row + '=0);"";D' + row + '/$B' + row + ')']);
    }
    batchSetFormulasSafeNoApiGE_(hist.getRange(3, 5, 60, 1), shareFormulas);

    hist.getRange(63, 4, 1, 4).setBackground("#F5B7B1");
    hist.getRange(63, 4).setNumberFormat("0.00").setFontWeight("bold");

    atualizarFormulasHistoricoGE_(hist);
    Utilities.sleep(2000);
    garantirCamposHistoricoGE_();
    atualizarMediaSemanaGE_();
    Utilities.sleep(1000);
    atualizarGraficoMediasGE_();
    Utilities.sleep(1000);
    atualizarMediaPorComentaristaGE_();
    atualizarMediaPorDiaSemanaGE_();
    // FIX: NÃO chamar atualizarMediaPorTemaSeExistir_() aqui. Essa função
    // faz sheet.clear() na aba "Tema Minuto" e, se a escrita
    // posterior das fórmulas falhar por quota (acumulado de chamadas no
    // adicionarNovoDiaManualGE), a aba fica zerada. A atualização da aba TEMA
    // agora é responsabilidade explícita do menu ou do fluxo Completa.

    // Alimenta a base normalizada, o mapa de calor e reaplica a formatação
    // condicional das colunas &/MIN (inclui a coluna recém-criada).
    try { appendDiaBaseGE_(programDate); } catch (e) { Logger.log("Base: " + e.message); }
    try { atualizarResumoGE(); } catch (e) { Logger.log("Resumo: " + e.message); }
    try { corrigirFormatacaoHistoricoGE(); } catch (e) { Logger.log("Formatação: " + e.message); }

    Logger.log("Programa GE de " + dateStr + " arquivado com sucesso");
    return true;
}

// =============================================================================
// atualizarFormulasBaseGE_ — fórmulas-base SEM calibragem
// =============================================================================
// Atualiza B (média histórica), C (% à média), row 63 (médias do programa) e
// a coluna de share (&/MIN) — sem tocar nas auxiliares GS-GY ou GZ-HC.
//
// FIX CIRCULARIDADE: B agora exclui a coluna D (dia corrente). B incluía D e
// os shares arquivados (=H/$B etc.) dependem de $B — fechava o ciclo
// D → MOLDE → PROCV(aba TEMA) → share fechado → B → D quando o tema casava
// no PROCV. B = média apenas dos dias FECHADOS (H, L, P, ...).
//
// FIX API: esta versão grava via SpreadsheetApp, sem Sheets.Spreadsheets.Values.update.
// -----------------------------------------------------------------------------
function atualizarFormulasBaseGE_(sheet) {
    var audCols = getAudColsGE_(sheet);
    if (audCols.length === 0) return;

    // Colunas fechadas = todas exceto a 4 (dia corrente)
    var closedCols = audCols.filter(function(c) { return c !== 4; });

    var bcFormulas = [];
    for (var row = 3; row <= 62; row++) {
        var fB;
        if (closedCols.length === 0) {
            fB = '=""';  // sem histórico fechado: B vazio
        } else {
            // MÉDIA (ignora células vazias) em vez de soma/N: um minuto sem
            // captura num dia arquivado ficava valendo 0 e puxava a média
            // do minuto para baixo em todos os dias seguintes.
            var medParts = closedCols.map(function(c) { return columnToLetter_(c) + row; }).join(";");
            fB = '=SEERRO(MÉDIA(' + medParts + ');"")';
        }
        var fC = '=SE(OU(B' + row + '="";B' + row + '=0;D' + row + '="");"";D' + row + '/B' + row + '-1)';
        bcFormulas.push([fB, fC]);
    }
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, 2, 60, 2), bcFormulas);

    // Linha 63: UMA fórmula por coluna de AUDIÊNCIA, na própria coluna.
    var sparse = [];
    for (var i = 0; i < audCols.length; i++) {
        var colL = columnToLetter_(audCols[i]);
        sparse.push({
            row: 63,
            col: audCols[i],
            // MÉDIA dos minutos preenchidos (dia com lacuna de captura não
            // fica subestimado como na antiga SOMA/60).
            formula: '=SEERRO(MÉDIA(' + colL + "3:" + colL + '62);"")'
        });
    }
    batchSetFormulasSparseSafeNoApiGE_(sheet, sparse);
    for (var j = 0; j < audCols.length; j++) {
        sheet.getRange(63, audCols[j]).setNumberFormat("0.00");
    }

    setFormulaSafeNoApiGE_(sheet.getRange(63, 2), "=(SOMA(B3:B62))/60");
    sheet.getRange(63, 2).setNumberFormat("0.00");
    setFormulaSafeNoApiGE_(sheet.getRange(63, 3), '=SE(OU(B63="";B63=0;D63="");"";D63/B63-1)');
    sheet.getRange(63, 3).setNumberFormat("0.0%");

    SpreadsheetApp.flush();
    sheet.getRange(3, 2, 60, 1).setNumberFormat("0.00");
    sheet.getRange(3, 3, 60, 1).setNumberFormat("0.0%");
}

// --- Atualizar fórmulas do histórico GE ---
// Wrapper que mantém comportamento histórico (base + calibragem + fator tema).
// Preservado para compatibilidade total com arquivarProgramaGE_ e demais chamadas.
function atualizarFormulasHistoricoGE_(sheet) {
    atualizarFormulasBaseGE_(sheet);
    atualizarCalibragemGE_(sheet);
    atualizarFatorTemaGE_(sheet);
}

// --- Garantir campos de rodapé do histórico GE ---
// FIX: (1) não expande mais a grade para 160 linhas — nada deve existir abaixo
// do fim do painel; (2) escreve apenas os 3 labels A63:A65 — a versão anterior
// também gravava "MÉDIA SEM." em A66 e "% vs ANT." em A67, por cima dos
// rótulos do painel ("SEM. ATUAL (...)" e a semana anterior).
function garantirCamposHistoricoGE_() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    var maxRows = sheet.getMaxRows();
    if (maxRows < 94) sheet.insertRowsAfter(maxRows, 94 - maxRows);
    var labels = [["MÉDIA"], ["COMENT. 1"], ["COMENT. 2"]];
    sheet.getRange(63, 1, 3, 1).setValues(labels);
    sheet.getRange(63, 1, 3, 1).setFontWeight("bold").setFontSize(8).setFontColor("#666666");
    preencherBlocoA66E85HistoricoGE_(sheet);
}

// =============================================================================
// localizarPainelGE_ — localiza dinamicamente as linhas do painel A66+
// =============================================================================
// O painel cresce 1 linha por semana e 1 por mês, então NENHUMA linha pode ser
// fixa. Retorna:
//   rowSemana   = linha do "SEM. ATUAL (dd/MM-dd/MM)"
//   semanaLabel = texto desse label
//   rowMesAtual = linha do "<MÊS> (mês atual)"
//   mesLabel    = texto desse label
//   panelEnd    = última linha do painel — sentinela "-%" na coluna C
//                 (acompanha o mês mais antigo); fallback: rowMesAtual+5 ou 94
// -----------------------------------------------------------------------------
function localizarPainelGE_(sheet) {
    var lastRow = Math.max(sheet.getLastRow(), 94);
    var n = lastRow - 66 + 1;
    if (n < 1) return null;

    var aVals = sheet.getRange(66, 1, n, 1).getValues();
    var cVals = sheet.getRange(66, 3, n, 1).getValues();

    var p = { rowSemana: -1, semanaLabel: "", rowMesAtual: -1, mesLabel: "", panelEnd: -1 };
    for (var i = 0; i < n; i++) {
        var s = (aVals[i][0] || "").toString();
        if (p.rowSemana < 0 && s.indexOf("SEM. ATUAL") === 0) {
            p.rowSemana = 66 + i;
            p.semanaLabel = s;
        }
        if (p.rowMesAtual < 0 && s.indexOf("(mês atual)") >= 0) {
            p.rowMesAtual = 66 + i;
            p.mesLabel = s;
        }
        var c = (cVals[i][0] || "").toString().trim();
        if (c === "-%") p.panelEnd = Math.max(p.panelEnd, 66 + i);
    }
    if (p.panelEnd < 0) {
        p.panelEnd = p.rowMesAtual > 0 ? p.rowMesAtual + 5 : 94;
    }
    return p;
}

// =============================================================================
// restaurarBlocoMensalAposInsercaoGE_ — devolve o bloco MENSAL de H:I para D:E
// =============================================================================
// insertColumns(4, 4) empurra para a direita TUDO da coluna D — inclusive o
// bloco mensal do painel ("MENSAL", refs manuais 0,66/0,65/... e =(B/D)-1),
// que para em H:I. Esta função devolve esse conteúdo para D:E (reescrevendo as
// fórmulas para apontar de volta a D) e limpa H:I na faixa do painel. Deve ser
// chamada IMEDIATAMENTE após a inserção de colunas, antes de qualquer outra
// escrita. FIM DINÂMICO: o painel cresce 1 linha por semana e 1 por mês, então
// a varredura vai da linha 66 até a última linha da planilha (mínimo 94).
// -----------------------------------------------------------------------------
function restaurarBlocoMensalAposInsercaoGE_(sheet) {
    var FIRST = 66;
    var LAST = Math.max(94, sheet.getLastRow());
    var n = LAST - FIRST + 1;
    var srcRange = sheet.getRange(FIRST, 8, n, 2); // H66:I{fim}
    var disp = srcRange.getDisplayValues();
    var hasContent = false, moved = 0;
    for (var i = 0; i < n; i++) {
        for (var j = 0; j < 2; j++) {
            if (String(disp[i][j] || "") !== "") { hasContent = true; moved++; }
        }
    }
    if (!hasContent) return 0;

    // Fase 2: a versão anterior fazia até 70 setValue/setFormula separados e
    // era o principal ponto de timeout da virada. copyPaste PASTE_NORMAL move
    // valores, fórmulas e formatos de uma vez; as referências relativas H→D
    // são ajustadas pelo próprio Sheets. O mesmo batch limpa as sobras.
    var sid = sheet.getSheetId();
    Sheets.Spreadsheets.batchUpdate({ requests: [
        { copyPaste: {
            source: { sheetId: sid, startRowIndex: FIRST - 1, endRowIndex: LAST,
                startColumnIndex: 7, endColumnIndex: 9 },
            destination: { sheetId: sid, startRowIndex: FIRST - 1, endRowIndex: LAST,
                startColumnIndex: 3, endColumnIndex: 5 },
            pasteType: "PASTE_NORMAL", pasteOrientation: "NORMAL"
        }},
        { repeatCell: {
            range: { sheetId: sid, startRowIndex: FIRST - 1, endRowIndex: LAST,
                startColumnIndex: 7, endColumnIndex: 9 },
            cell: {}, fields: "userEnteredValue,userEnteredFormat"
        }},
        { repeatCell: {
            range: { sheetId: sid, startRowIndex: FIRST - 1, endRowIndex: LAST,
                startColumnIndex: 5, endColumnIndex: 7 },
            cell: {}, fields: "userEnteredValue,userEnteredFormat"
        }}
    ]}, SpreadsheetApp.getActive().getId());

    Logger.log("[Painel] Bloco mensal devolvido de H:I para D:E (" + moved + " células); F:G do painel limpo.");
    return moved;
}

// =============================================================================
// gerarFormulaMedianaOffsetGE_ — mediana do mesmo dia da semana (padrão OFFSET)
// =============================================================================
// Reproduz o padrão da planilha para E63 do novo bloco:
//   =(MEDIAN(OFFSET(X63,0,0,1,1),OFFSET(X63,0,20),OFFSET(X63,0,40),...))
// Âncora = coluna de AUDIÊNCIA mais recente com o MESMO dia da semana do novo
// programa; passo de 20 colunas = 5 programas = 1 semana (o programa vai ao ar
// seg-sex SEM exceção, então o passo nunca desalinha). O número de termos
// replica o da fórmula do bloco anterior (I63, já deslocada), fallback 14.
// Retorna a fórmula em sintaxe en-US (para Range.setFormula nativo, que é
// independente de locale) ou null se não houver âncora.
// -----------------------------------------------------------------------------
function gerarFormulaMedianaOffsetGE_(sheet, programDate) {
    var audColsHist = getAudColsGE_(sheet).filter(function(c) { return c !== 4; });
    if (audColsHist.length === 0) return null;

    var row1Vals = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var weekday = programDate.getDay();

    var anchorCol = null;
    for (var i = 0; i < audColsHist.length; i++) { // newest-first
        var c = audColsHist[i];
        var v = row1Vals[c - 1];
        var d = v instanceof Date ? v
            : parseDateStr_(typeof v === "string" ? v.trim() : "");
        if (d && d.getDay() === weekday) { anchorCol = c; break; }
    }
    if (!anchorCol) return null;

    // Nº de termos: replica o padrão do bloco anterior (fallback 14)
    var nTerms = 14;
    var prevF = sheet.getRange(63, 9).getFormula(); // I63 = mediana antiga deslocada
    if (prevF && prevF.indexOf("OFFSET(") >= 0) {
        var m = prevF.match(/OFFSET\(/g);
        if (m && m.length >= 2) nTerms = m.length;
    }

    // Não ultrapassar a última coluna histórica
    var lastAud = audColsHist[audColsHist.length - 1];
    var maxTerms = Math.floor((lastAud - anchorCol) / 20) + 1;
    if (maxTerms < 1) maxTerms = 1;
    if (nTerms > maxTerms) nTerms = maxTerms;

    // FIX: gerar em pt-BR (MED + DESLOC + ; ) para gravar via setFormulaSafeNoApiGE_
    // (USER_ENTERED). A versão anterior em en-US (MEDIAN + OFFSET + ,) ficava
    // armazenada como literal numa planilha pt-BR, gerando #NAME?.
    var aL = columnToLetter_(anchorCol);
    var parts = ["DESLOC(" + aL + "63;0;0;1;1)"];
    for (var k = 1; k < nTerms; k++) parts.push("DESLOC(" + aL + "63;0;" + (20 * k) + ")");
    return "=(MED(" + parts.join(";") + "))";
}

// =============================================================================
// preencherBlocoA66E85HistoricoGE_ — painel PRESERVADO + semana/mês automáticos
// =============================================================================
// O painel A66:E{fim} nunca é reconstruído — labels, fórmulas B/C/D/E, "-%",
// refs mensais manuais e formatação ficam intactos (a inserção de colunas do
// novo dia desloca as referências sozinha, e restaurarBlocoMensalAposInsercaoGE_
// devolve o bloco mensal para D:E). Esta função faz apenas:
//
//   SEMANA — se o programa mais recente pertence à semana do label
//   "SEM. ATUAL (dd/MM-dd/MM)": injeta D63 na fórmula de B. Se pertence a uma
//   semana NOVA (toda segunda-feira, já que o programa vai ao ar seg-sex sem
//   exceção): rebaixa o label para "dd/MM-dd/MM" (a média da semana fechada
//   congela — as refs de coluna já acompanham as inserções de bloco), insere
//   uma linha nova no topo da seção semanal e cria
//   "SEM. ATUAL (novo intervalo)" com B=MÉDIA(D63) e C=variação vs semana
//   anterior. Tudo abaixo (semanas antigas, MENSAL, meses, "-%") desce 1 linha
//   com as fórmulas auto-ajustadas pelo Sheets.
//
//   MÊS — mesma mecânica: se o mês confere, injeta D63; se virou o mês
//   (mesmo no meio da semana), rebaixa "<MÊS> (mês atual)" para "<MÊS>",
//   insere linha nova com "<MÊS NOVO> (mês atual)", B=MÉDIA(D63), C=variação,
//   D vazio (referência mensal manual — preencha quando tiver o consolidado
//   Kantar) e E protegido contra D vazio (vira =(B/D)-1 assim que D for
//   preenchido).
//
// Idempotente: reexecutar no mesmo dia não cria linha duplicada nem repete D63.
// Fórmulas gravadas com Range.setFormula nativo (sintaxe en-US, independente
// de locale — o par getFormula/setFormula é sempre en-US).
// -----------------------------------------------------------------------------
function preencherBlocoA66E85HistoricoGE_(sheet) {
    var audCols = getAudColsGE_(sheet);
    if (audCols.length === 0) return;

    var row1Vals = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var v = row1Vals[audCols[0] - 1];
    var maisRecente = v instanceof Date ? v
        : parseDateStr_(typeof v === "string" ? v.trim() : "");
    if (!maisRecente) return;

    var MONTH_NAMES_PT = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
                          "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];

    function mondayOf(d) {
        var x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
        var dow = x.getDay();
        x.setDate(x.getDate() - (dow === 0 ? 6 : dow - 1));
        return x;
    }
    function ddmm(d) { return Utilities.formatDate(d, CFG.TIMEZONE, "dd/MM"); }
    function labelSemanaAtual(mon) {
        var fri = new Date(mon);
        fri.setDate(fri.getDate() + 4);
        return "SEM. ATUAL (" + ddmm(mon) + "-" + ddmm(fri) + ")";
    }
    // FIX: injeta D63 reescrevendo a fórmula POR INTEIRO em pt-BR via
    // setFormulaSafeNoApiGE_ (USER_ENTERED). A versão anterior assumia que
    // getFormula() retornava en-US (AVERAGE + ,) e fazia concat textual,
    // mas o Sheets pode devolver formato misto (=AVERAGE(...;...)) numa
    // planilha pt-BR, gerando =AVERAGE(D63,H63;L63;P63) — fórmula corrompida.
    // Aqui aceitamos AVERAGE/MÉDIA/MEDIA com qualquer separador (, ou ;),
    // extraímos os args com regex e reescrevemos como =MÉDIA(D63;...;...).
    function injetarD63(cell) {
        var f = cell.getFormula();
        if (!f) return false;
        if (/\bD63\b/.test(f)) return false; // já incluído
        var m = f.match(/^\s*=\s*(?:AVERAGE|MÉDIA|MEDIA)\s*\(([\s\S]*)\)\s*$/i);
        if (!m) return false;
        var args = m[1].split(/[;,]/).map(function(s){ return s.trim(); }).filter(Boolean);
        args.unshift("D63");
        setFormulaSafeNoApiGE_(cell, "=MÉDIA(" + args.join(";") + ")");
        return true;
    }

    var progMonday = mondayOf(maisRecente);
    var painel = localizarPainelGE_(sheet);
    if (!painel || painel.rowSemana < 0) {
        Logger.log("[Painel] ⚠️ Linha 'SEM. ATUAL' não encontrada — painel não atualizado.");
        return;
    }

    // ───────────────────────── SEMANA ─────────────────────────
    var mm = painel.semanaLabel.match(/\((\d{2}\/\d{2})-/);
    var mesmaSemana = !!(mm && mm[1] === ddmm(progMonday));

    if (mesmaSemana) {
        if (injetarD63(sheet.getRange(painel.rowSemana, 2))) {
            Logger.log("[Painel] D63 incluído na SEM. ATUAL (B" + painel.rowSemana + ").");
        }
    } else {
        var r = painel.rowSemana;
        // 1) Rebaixa o label da semana que fechou: "SEM. ATUAL (08/06-12/06)" -> "08/06-12/06"
        var labelFechada = painel.semanaLabel
            .replace(/^SEM\. ATUAL\s*\(/, "").replace(/\)\s*$/, "").trim();
        if (labelFechada) sheet.getRange(r, 1).setValue(labelFechada);
        // 2) Insere a linha da nova semana no topo da seção semanal
        sheet.insertRowsBefore(r, 1);
        // 3) Formato herdado da linha de baixo (a ex-SEM. ATUAL)
        sheet.getRange(r + 1, 1, 1, 5).copyTo(sheet.getRange(r, 1, 1, 5),
            SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        // 4) Conteúdo da nova semana
        sheet.getRange(r, 1).setValue(labelSemanaAtual(progMonday));
        // FIX: pt-BR via setFormulaSafeNoApiGE_ (USER_ENTERED) — locale-aware
        setFormulaSafeNoApiGE_(sheet.getRange(r, 2), "=MÉDIA(D63)");
        setFormulaSafeNoApiGE_(sheet.getRange(r, 3),
            '=SE(OU(B' + (r + 1) + '="";B' + (r + 1) + '=0);"";(B' + r + '/B' + (r + 1) + ')-1)');
        Logger.log("[Painel] Nova semana criada na linha " + r + ": " + labelSemanaAtual(progMonday));
        // Tudo abaixo desceu 1 linha — relocaliza antes de mexer no mês
        painel = localizarPainelGE_(sheet);
    }

    // ───────────────────────── MÊS ─────────────────────────
    if (!painel || painel.rowMesAtual < 0) {
        Logger.log("[Painel] ⚠️ Linha '(mês atual)' não encontrada — bloco mensal não atualizado.");
        return;
    }

    var nomeMesLabel = painel.mesLabel.replace(/\s*\(.*\)\s*$/, "").trim().toUpperCase();
    var nomeMesProg = MONTH_NAMES_PT[maisRecente.getMonth()];

    if (nomeMesLabel === nomeMesProg) {
        if (injetarD63(sheet.getRange(painel.rowMesAtual, 2))) {
            Logger.log("[Painel] D63 incluído no mês atual (B" + painel.rowMesAtual + ").");
        }
    } else {
        var rm = painel.rowMesAtual;
        // 1) Rebaixa o label do mês que fechou: "JUNHO (mês atual)" -> "JUNHO"
        sheet.getRange(rm, 1).setValue(nomeMesLabel);
        // 2) Insere a linha do novo mês no topo da seção mensal
        //    (o header "MENSAL" em D fica na linha acima e não se move)
        sheet.insertRowsBefore(rm, 1);
        // 3) Formato herdado da linha de baixo (o ex-mês atual)
        sheet.getRange(rm + 1, 1, 1, 5).copyTo(sheet.getRange(rm, 1, 1, 5),
            SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        // 4) Conteúdo do novo mês
        sheet.getRange(rm, 1).setValue(nomeMesProg + " (mês atual)");
        // FIX: pt-BR via setFormulaSafeNoApiGE_
        setFormulaSafeNoApiGE_(sheet.getRange(rm, 2), "=MÉDIA(D63)");
        setFormulaSafeNoApiGE_(sheet.getRange(rm, 3),
            '=SE(OU(B' + (rm + 1) + '="";B' + (rm + 1) + '=0);"";(B' + rm + '/B' + (rm + 1) + ')-1)');
        // D fica vazio: é a referência mensal MANUAL (consolidado Kantar).
        // E protegido: vazio enquanto D estiver vazio; vira (B/D)-1 ao preencher.
        setFormulaSafeNoApiGE_(sheet.getRange(rm, 5),
            '=SE(OU(D' + rm + '="";D' + rm + '=0);"";(B' + rm + '/D' + rm + ')-1)');
        Logger.log("[Painel] Novo mês criado na linha " + rm + ": " + nomeMesProg +
            " — preencha a referência mensal em D" + rm + " quando disponível.");
    }
}

// --- Média por semana GE ---
//
// REESCRITO. Antes pulverizava valores em (66, sexta.col) e (67, sexta.col) de
// cada coluna de sexta-feira histórica — espalhando 80+ valores avulsos pelas
// colunas F:NT das linhas 66 e 67, dentro da área protegida A66:NT89.
//
// Agora a média semanal vive integralmente no painel dinâmico A66:E89
// (preencherBlocoA66E85HistoricoGE_). Esta função apenas limpa eventuais
// resíduos de versões anteriores nas rows 66-67 das colunas F+ —
// safety-net se algo escreveu lá entre execuções.
// -----------------------------------------------------------------------------
function atualizarMediaSemanaGE_() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    var lastCol = sheet.getLastColumn();
    if (lastCol > 5) {
        // Limpa F66:lastCol67 (resíduo de versões anteriores)
        sheet.getRange(66, 6, 2, lastCol - 5).clearContent();
    }
    Logger.log("[atualizarMediaSemanaGE_] Resíduos em F66:NT67 limpos. Cálculo semanal vive em A66:E89.");
}

// --- Gráfico de médias GE ---
//
// REESCRITO. A versão anterior escrevia cabeçalho + 60 linhas de dados
// auxiliares em A95:D155 e limpava A90:D94 — mas no layout correto do painel
// as linhas logo abaixo do mês atual são os meses anteriores, e NADA deve
// existir abaixo do fim do painel. Agora o gráfico é alimentado DIRETAMENTE
// pelas colunas existentes (A=hora, B=média histórica, D=programa mais
// recente), sem área auxiliar. A limpeza de resíduos legados respeita o FIM
// DINÂMICO do painel (sentinela "-%" na coluna C, que acompanha o mês mais
// antigo): como o painel cresce 1 linha por semana e 1 por mês, limpar a
// partir de uma linha fixa destruiria o painel no futuro.
// -----------------------------------------------------------------------------
function atualizarGraficoMediasGE_() {
    SpreadsheetApp.flush();
    Utilities.sleep(500);
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;

    var audCols = getAudColsGE_(sheet);

    var charts = sheet.getCharts();
    for (var i = 0; i < charts.length; i++) {
        var info = charts[i].getContainerInfo();
        if (info && info.getAnchorRow() >= 1 && info.getAnchorColumn() >= 8) {
            sheet.removeChart(charts[i]);
        }
    }

    // Limpeza de resíduos legados — SOMENTE abaixo do fim real do painel.
    var painel = localizarPainelGE_(sheet);
    var panelEnd = painel ? painel.panelEnd : 94;
    var maxRows = sheet.getMaxRows();
    if (maxRows > panelEnd) {
        sheet.getRange(panelEnd + 1, 1, maxRows - panelEnd, 4).clearContent();
    }

    if (audCols.length === 0) return;

    var lastDate = sheet.getRange(1, 4).getValue();
    var dateStr = "";
    if (lastDate instanceof Date) {
        dateStr = Utilities.formatDate(lastDate, "America/Sao_Paulo", "dd/MM/yyyy");
    } else if (typeof lastDate === "string" && lastDate.trim()) {
        dateStr = lastDate.trim();
    }

    var chart = sheet.newChart()
        .asLineChart()
        .addRange(sheet.getRange(2, 1, 61, 1))  // A2:A62 — HORA (rótulos)
        .addRange(sheet.getRange(2, 2, 61, 1))  // B2:B62 — MÉDIA histórica
        .addRange(sheet.getRange(2, 4, 61, 1))  // D2:D62 — programa mais recente
        .setPosition(2, 9, 0, 0)
        .setOption("title", "Gazeta Esportiva" + (dateStr ? " — " + dateStr : ""))
        .setOption("titleTextStyle", { fontSize: 12, bold: true })
        .setOption("legend", { position: "bottom" })
        .setOption("hAxis", { title: "Horário", slantedText: true, slantedTextAngle: 45 })
        .setOption("vAxis", { title: "Audiência" })
        .setOption("series", {
            0: { color: "#E74C3C", lineWidth: 3 },  // MÉDIA HISTÓRICA
            1: { color: "#2980B9", lineWidth: 2 }   // ÚLTIMO PROGRAMA
        })
        .setOption("width", 800)
        .setOption("height", 400)
        .build();
    sheet.insertChart(chart);
}

// --- Comentaristas fixos da casa (edite para adicionar/remover) ---
var COMENTARISTAS_DA_CASA_GE = ["ALEX", "GARRAFFA", "MULLER", "PAULO SÉRGIO", "SALAZAR", "SILVESTRE"];

// --- Média por comentarista GE ---

function atualizarMediaPorComentaristaGE_() {
    var ss = SpreadsheetApp.getActive();
    var hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!hist) return;
    var tabName = "Comentaristas";
    var dest = ss.getSheetByName(tabName);
    if (dest) { dest.clear(); } else { dest = ss.insertSheet(tabName); }

    var audColsList = getAudColsGE_(hist);
    var numPrograms = audColsList.length;

    // === Leitura batch (1 chamada de API) ===
    var programData = [];
    var globalSum = 0, globalCount = 0;
    var dayTotals = {};

    if (numPrograms > 0) {
        var maxAudCol = audColsList[audColsList.length - 1];
        var allData = hist.getRange(1, 1, 65, maxAudCol + 3).getValues();

        for (var i = 0; i < audColsList.length; i++) {
            var audCol = audColsList[i];
            var ci = audCol - 1; // índice 0-based

            // Data (row 1)
            var dateVal = allData[0][ci];
            var dDate = dateVal instanceof Date ? dateVal
                : parseDateStr_(typeof dateVal === "string" ? dateVal.trim() : "");

            // Auto-calcular média das rows 3-62 (índices 2-61) — não depende da fórmula da row 63
            var sum = 0, cnt = 0;
            for (var ri = 2; ri <= 61; ri++) {
                var v = allData[ri][ci];
                if (typeof v === "number" && v > 0) { sum += v; cnt++; }
            }
            if (cnt === 0) continue;
            var avg = sum / cnt;

            var dow = dDate ? dDate.getDay() : -1;
            var dayName = (dow >= 1 && dow <= 5)
                ? ["", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA"][dow] : "";

            // Comentaristas (rows 64-65, índices 63-64)
            var names = [];
            for (var ri2 = 63; ri2 <= 64; ri2++) {
                var raw = allData[ri2] ? allData[ri2][ci] : "";
                if (!raw || typeof raw === "number" || raw instanceof Date) continue;
                var nm = raw.toString().trim().toUpperCase();
                if (nm && nm.charAt(0) !== "#" && !/^\d+([.,]\d+)?$/.test(nm)) names.push(nm);
            }

            programData.push({ avg: avg, dDate: dDate, day: dayName, names: names });
            globalSum += avg; globalCount++;

            if (dayName) {
                if (!dayTotals[dayName]) dayTotals[dayName] = { sum: 0, count: 0 };
                dayTotals[dayName].sum += avg;
                dayTotals[dayName].count++;
            }
        }
    }

    var globalAvg = globalCount > 0 ? globalSum / globalCount : 0;
    var dayMeans = {};
    for (var d in dayTotals) {
        dayMeans[d] = dayTotals[d].count > 0 ? dayTotals[d].sum / dayTotals[d].count : globalAvg;
    }

    // Agrupar aparições por comentarista
    var coment = {};
    for (var i = 0; i < programData.length; i++) {
        var prog = programData[i];
        var seen = {};
        for (var n = 0; n < prog.names.length; n++) {
            var nome = prog.names[n];
            if (seen[nome]) continue;
            seen[nome] = true;
            if (!coment[nome]) coment[nome] = [];
            coment[nome].push({ avg: prog.avg, day: prog.day, dDate: prog.dDate });
        }
    }

    // Calcular estatísticas de um comentarista
    function calcStats(arr) {
        arr.sort(function(a, b) {
            return (a.dDate ? a.dDate.getTime() : 0) - (b.dDate ? b.dDate.getTime() : 0);
        });
        var avgs = arr.map(function(x) { return x.avg; });
        var mean = avgs.reduce(function(s, v) { return s + v; }, 0) / avgs.length;
        var l5 = avgs.slice(-5), l3 = avgs.slice(-3);
        var m5 = l5.reduce(function(s, v) { return s + v; }, 0) / l5.length;
        var m3 = l3.reduce(function(s, v) { return s + v; }, 0) / l3.length;
        // Média ponderada: corrige viés de dia da semana
        // adjusted = avg * (globalAvg / dayMean) → normaliza cada programa pela força do seu dia
        var adj = arr.map(function(x) {
            var dm = dayMeans[x.day] || globalAvg;
            return dm > 0 ? x.avg * (globalAvg / dm) : x.avg;
        });
        var wAvg = adj.reduce(function(s, v) { return s + v; }, 0) / adj.length;
        return [arr.length, mean, m5, m3, wAvg];
    }

    var CASA = COMENTARISTAS_DA_CASA_GE;
    var convidados = Object.keys(coment).filter(function(n) {
        return CASA.indexOf(n) === -1;
    }).sort();

    var header = ["COMENTARISTA", "QTD", "MÉDIA GERAL", "ÚLTIMAS 5", "ÚLTIMAS 3", "MÉD. PONDERADA"];
    var nCols = header.length;
    var currentRow = 1;

    // Escrever seção (DA CASA ou CONVIDADOS)
    function writeSection(title, bgTitle, bgHeader, members, altBg) {
        // Título da seção
        dest.getRange(currentRow, 1).setValue(title);
        dest.getRange(currentRow, 1, 1, nCols)
            .setBackground(bgTitle).setFontColor("white").setFontWeight("bold");
        currentRow++;

        // Cabeçalho de colunas
        dest.getRange(currentRow, 1, 1, nCols).setValues([header]);
        dest.getRange(currentRow, 1, 1, nCols)
            .setBackground(bgHeader).setFontColor("white").setFontWeight("bold")
            .setHorizontalAlignment("center");
        currentRow++;

        if (members.length === 0) {
            dest.getRange(currentRow, 1).setValue("(nenhum)");
            currentRow++;
            return;
        }

        var startRow = currentRow;
        var rows = [];
        for (var k = 0; k < members.length; k++) {
            var nomeK = members[k];
            var arr = coment[nomeK];
            if (arr && arr.length > 0) {
                var s = calcStats(arr);
                rows.push([nomeK, s[0], s[1], s[2], s[3], s[4]]);
            } else {
                rows.push([nomeK, 0, "", "", "", ""]);
            }
        }

        dest.getRange(startRow, 1, rows.length, nCols).setValues(rows);

        // Cores alternadas
        for (var r = 0; r < rows.length; r++) {
            dest.getRange(startRow + r, 1, 1, nCols).setBackground(r % 2 === 0 ? altBg : "#FFFFFF");
        }

        // Formatos numéricos
        dest.getRange(startRow, 2, rows.length, 1).setNumberFormat("0").setHorizontalAlignment("center");
        dest.getRange(startRow, 3, rows.length, 4).setNumberFormat("0.00").setHorizontalAlignment("center");

        currentRow += rows.length;
    }

    writeSection("⭐ COMENTARISTAS DA CASA", "#922B21", "#E74C3C", CASA, "#FDF2F2");
    currentRow++; // linha vazia entre seções
    writeSection("🎙️ CONVIDADOS", "#1A5276", "#2980B9", convidados, "#EBF5FB");

    // Larguras das colunas
    dest.setColumnWidth(1, 185);
    dest.setColumnWidth(2, 55);
    dest.setColumnWidth(3, 115);
    dest.setColumnWidth(4, 100);
    dest.setColumnWidth(5, 100);
    dest.setColumnWidth(6, 135);

    SpreadsheetApp.flush();
}

// --- Média por dia da semana GE ---

function atualizarMediaPorDiaSemanaGE_() {
    // Substituída na migração de 07/2026: a versão antiga gravava fórmulas
    // =SEERRO(MÉDIA(...)) que quebram conforme o locale do arquivo (#NAME?).
    // A reconstrução agora é feita a valores em 53_Dia_Semana_GE.js.
    try {
        atualizarDiaSemanaGE();
    } catch (e) {
        Logger.log("atualizarDiaSemanaGE falhou: " + e.message);
    }
}
function atualizarDashboardsGE() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (sheet) {
        atualizarCalibragemGE_(sheet);
        atualizarFatorTemaGE_(sheet);
    }
    garantirCamposHistoricoGE_();
    atualizarMediaSemanaGE_();
    atualizarGraficoMediasGE_();
    atualizarMediaPorComentaristaGE_();
    atualizarMediaPorDiaSemanaGE_();
    // FIX: NÃO chamar atualizarMediaPorTemaSeExistir_() aqui. Essa função
    // faz sheet.clear() na aba "Tema Minuto" e, se a escrita
    // posterior das fórmulas falhar por quota (acumulado de chamadas no
    // adicionarNovoDiaManualGE), a aba fica zerada. A atualização da aba TEMA
    // agora é responsabilidade explícita do menu ou do fluxo Completa.
}

// --- Helpers de fórmulas para a aba Média por Tema GE ---
// FIX CIRCULARIDADE: os geradores agora EXCLUEM o bloco do dia corrente
// (colunas D/E/F, t=0) e usam apenas dias FECHADOS (t>=1, blocos H em diante).
// Antes, CONT.SE/SOMASE varriam F e E de hoje (E = D/B) e a média multiplicava
// por B63 (que soma D) — quando o TEMA era preenchido numa linha em que D ainda
// era fórmula, fechava o ciclo D → HC/MOLDE → PROCV(aba TEMA) → E/B63 → D,
// e o Sheets exibia #REF! (dependência circular) em toda a cadeia futura.

// Expressão da média geral SÓ com dias fechados (substitui 'hist'!$B$63).
// Retorna string SEM "=". Se não houver dias fechados, retorna "0".
function gerarMediaGeralFechadaGE_(histName, totalPrograms) {
    var parts = [];
    for (var t = 1; t < totalPrograms; t++) {
        var audL = columnToLetter_(4 + 4 * t); // H, L, P, ...
        parts.push("SEERRO('" + histName + "'!" + audL + "63;\"\")");
    }
    if (parts.length === 0) return "0";
    return "MÉDIA(" + parts.join(";") + ")";
}

function gerarFormulaCountTema_(e, t, o, a) {
    // e=nome da aba histórico | t=tema | o=total de programas | a=qtos usar
    var r = t.replace(/"/g, '""');
    var closed = o - 1;                      // exclui o dia corrente (t=0)
    var use = Math.min(a, closed);
    if (use <= 0) return "=0";
    var n = [];
    for (var k = 1; k <= use; k++) {
        var temaL = columnToLetter_(6 + 4 * k); // J, N, R, ...
        n.push("CONT.SE('" + e + "'!" + temaL + "3:" + temaL + '62;"' + r + '")');
    }
    return "=" + n.join("+");
}

function gerarFormulaMediaTema_(e, t, o, a) {
    // e=nome da aba histórico | t=tema | o=total de programas | a=qtos usar
    var r = t.replace(/"/g, '""');
    var closed = o - 1;                      // exclui o dia corrente (t=0)
    var use = Math.min(a, closed);
    if (use <= 0) return "=0";
    var n = [], s = [];
    for (var k = 1; k <= use; k++) {
        var temaL = columnToLetter_(6 + 4 * k);  // J, N, R, ...
        var shaL = columnToLetter_(5 + 4 * k);   // I, M, Q, ...
        n.push("SOMASE('" + e + "'!" + temaL + "3:" + temaL + '62;"' + r + "\";'" + e + "'!" + shaL + "3:" + shaL + "62)");
        s.push("CONT.SE('" + e + "'!" + temaL + "3:" + temaL + '62;"' + r + '")');
    }
    var i = n.join("+"), c = s.join("+");
    var mediaGeral = gerarMediaGeralFechadaGE_(e, o); // sem B63 (que inclui D)
    return "=SE((" + c + ")=0;0;(" + i + ")/(" + c + ")*(" + mediaGeral + "))";
}

// --- Criar aba Média por Tema GE (entrada pública via menu) ---

function criarMediaPorTemaGE() {
    normalizarTemasNoHistoricoGE_();
    const e = SpreadsheetApp.getUi(),
        t = SpreadsheetApp.getActive(),
        o = t.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!o) return void e.alert("\u274c Erro", 'A aba "Hist\xf3rico" n\xe3o existe.', e.ButtonSet.OK);
    const r = getAudColsGE_(o).length;
    if (0 === r) return void e.alert("\u274c Erro", "N\xe3o h\xe1 programas arquivados no hist\xf3rico.", e.ButtonSet.OK);
    const n = coletarTemasUnicosGE_(o, r);
    if (0 === n.length) return void e.alert("\u274c Erro", "N\xe3o foram encontrados TEMAS no hist\xf3rico.", e.ButtonSet.OK);
    const s = "Tema Minuto";
    let i = t.getSheetByName(s);
    i ? i.clear() : i = t.insertSheet(s), configurarAbaMediaPorTema_(i, o, n, r), e.alert("\u2705 Sucesso!", `Aba "${s}" criada/atualizada com sucesso!\n\n\u{1f4ca} ${n.length} TEMAS encontrados\n\u{1f4c5} ${r} programas no hist\xf3rico\n\nAs m\xe9dias s\xe3o calculadas automaticamente via f\xf3rmulas.`, e.ButtonSet.OK)
}

// --- Atualizar fórmulas da aba Média por Tema GE (entrada pública via menu) ---
// FIX: a versão anterior fazia a.clear() ANTES de reconstruir tudo. Numa
// execução longa (fluxo Completa, com muitas chamadas API anteriores), se a
// escrita das fórmulas falhasse por quota, a aba ficava com header + temas
// em A mas sem dados em B:F (médias somem). Agora a atualização é
// IN-PLACE: cabeçalho, temas e fórmulas são sobrescritos diretamente, e
// linhas excedentes (caso menos temas que antes) são limpas no final.
// Se a escrita de uma linha falhar, todas as anteriores permanecem.
function atualizarFormulasTemaGE() {
    normalizarTemasNoHistoricoGE_();
    var lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        var ui = SpreadsheetApp.getUi();
        var ss = SpreadsheetApp.getActive();
        var sheetName = "Tema Minuto";
        var temaSheet = ss.getSheetByName(sheetName);
        if (!temaSheet) {
            ui.alert("❌ Erro", 'A aba "' + sheetName + '" não existe.\n\nCrie a aba primeiro: Gazeta Esportiva → Análise por Tema → Criar/Atualizar Média por Tema', ui.ButtonSet.OK);
            return;
        }
        var histSheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
        if (!histSheet) {
            ui.alert("❌ Erro", 'A aba "Histórico" não existe.', ui.ButtonSet.OK);
            return;
        }
        var numPrograms = getAudColsGE_(histSheet).length;
        if (numPrograms === 0) {
            ui.alert("❌ Erro", "Não há programas arquivados no histórico.", ui.ButtonSet.OK);
            return;
        }
        var temas = coletarTemasUnicosGE_(histSheet, numPrograms);
        if (temas.length === 0) {
            ui.alert("❌ Erro", "Não foram encontrados TEMAS no histórico.", ui.ButtonSet.OK);
            return;
        }

        // Quantas linhas a aba tinha antes (para limpar excedente no fim)
        var lastRowBefore = temaSheet.getLastRow();

        // Reconstruir IN-PLACE (sem clear() inicial)
        configurarAbaMediaPorTema_(temaSheet, histSheet, temas, numPrograms);

        // Limpar excedente: se antes tinha mais temas, sobraram linhas órfãs
        // entre (temas.length + 2) e lastRowBefore. A função configurar... já
        // escreve o rodapé "TOTAL/MÉDIA GERAL" em temas.length + 3 e infos
        // em temas.length + 5..9. Limpamos somente abaixo disso.
        var rodapeEnd = temas.length + 1 + 1 + 1 + 5; // header + temas + total + 5 infos
        if (lastRowBefore > rodapeEnd) {
            temaSheet.getRange(rodapeEnd + 1, 1, lastRowBefore - rodapeEnd, 6).clearContent().clearFormat();
        }

        ui.alert("✅ Fórmulas Atualizadas!",
            'A aba "' + sheetName + '" foi atualizada com sucesso!\n\n📊 ' + temas.length + " TEMAS\n📅 " + numPrograms + " programas",
            ui.ButtonSet.OK);
    } finally {
        lock.releaseLock();
    }
}

// --- Normalizar temas no histórico GE ---

function normalizarTemasNoHistoricoGE_() {
    const e = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!e) return;
    const audColsNorm = getAudColsGE_(e);
    for (let a = 0; a < audColsNorm.length; a++) {
        const o = 6 + 4 * a;  // coluna TEMA do programa a (col 4=AUD, 5=SHA, 6=TEMA)
        const tRange = e.getRange(3, o, 60, 1),
            vals = tRange.getValues();
        let changed = false;
        for (let i = 0; i < vals.length; i++) {
            const v = vals[i][0];
            if (!v) continue;
            const norm = normalizarTema_(v);
            if (norm !== v) { vals[i][0] = norm; changed = true; }
        }
        if (changed) tRange.setValues(vals);
    }
}

// --- Coletar temas únicos do histórico GE ---

function coletarTemasUnicosGE_(e, t) {
    const o = new Set;
    for (let a = 0; a < t; a++) {
        const t = 6 + 4 * a,
            r = e.getRange(3, t, 60, 1).getValues();
        for (const e of r) {
            const t = e[0];
            t && typeof t === "string" && /[a-zA-ZÀ-ÿ]{2,}/.test(t.trim()) && o.add(t.trim())
        }
    }
    return Array.from(o).sort()
}

// --- Configurar aba Média por Tema GE ---
// FIX: índices calculados só com dias FECHADOS; rodapé usa média geral fechada
// (sem B63); "Últimos 5/3" agora contam a partir do 1º dia fechado.
function configurarAbaMediaPorTema_(e, t, o, a) {
    const r = CFG.GE_OUT_HISTORICO,
        n = ["TEMA", "Minutos Total", "M\xe9dia &/MIN (Todos)", "M\xe9dia &/MIN (\xdaltimos 5)", "M\xe9dia &/MIN (\xdaltimos 3)", "\xcdndice vs M\xe9dia Geral"];
    e.getRange(1, 1, 1, n.length).setValues([n]), e.getRange(1, 1, 1, n.length).setFontWeight("bold").setBackground("#E74C3C").setFontColor("white").setHorizontalAlignment("center");
    const closed = Math.max(0, a - 1),        // dias fechados (exclui hoje)
        s = Math.min(5, closed),
        i = Math.min(3, closed),
        c = [];
    for (let e2 = 0; e2 < o.length; e2++) c.push([o[e2]]);
    c.length > 0 && e.getRange(2, 1, c.length, 1).setValues(c);
    const l2 = o.length,
        d = new Array(l2);
    for (let t2 = 0; t2 < l2; t2++) {
        const nome = o[t2],
            c2 = t2 + 2,
            m = gerarFormulaCountTema_(r, nome, a, closed),
            p = gerarFormulaMediaTema_(r, nome, a, closed),
            u = gerarFormulaMediaTema_(r, nome, a, s),
            g = gerarFormulaMediaTema_(r, nome, a, i);
        const b = sanitizeFormula_(m),
            h = sanitizeFormula_(p),
            v = sanitizeFormula_(u),
            y = sanitizeFormula_(g);
        d[t2] = [(b && "=" !== b[0] ? "=" + b : b), (h && "=" !== h[0] ? "=" + h : h), (v && "=" !== v[0] ? "=" + v : v), (y && "=" !== y[0] ? "=" + y : y), sanitizeFormula_("=SE(OU(C" + c2 + "=\"\";C" + c2 + "=0;(" + gerarMediaGeralFechadaGE_(r, a) + ")=0);1;C" + c2 + "/(" + gerarMediaGeralFechadaGE_(r, a) + "))")]
    }
    const A = 50;
    for (let t3 = 0; t3 < d.length; t3 += A) {
        const n2 = d.slice(t3, t3 + A),
            s2 = e.getRange(t3 + 2, 2, n2.length, 5);
        let i2 = 0;
        for (;;) {
            try {
                setFormulasUserEnteredNoApiGE_(s2, n2);
                break
            } catch (err) {
                if (++i2 >= 5) throw err;
                Utilities.sleep(200 * Math.pow(2, i2 - 1))
            }
        }
    }
    SpreadsheetApp.flush();

    e.getRange(2, 3, o.length, 3).setNumberFormat("0.00");

    const l = o.length;
    e.setColumnWidth(1, 150), e.setColumnWidth(2, 100), e.setColumnWidth(3, 130), e.setColumnWidth(4, 130), e.setColumnWidth(5, 130), e.setColumnWidth(6, 140), l > 0 && (e.getRange(2, 2, l, 1).setNumberFormat("0"), e.getRange(2, 3, l, 3).setNumberFormat("0.00"), e.getRange(2, 6, l, 1).setNumberFormat("0.00")), e.getRange(2, 2, l, 5).setHorizontalAlignment("center");
    for (let t4 = 0; t4 < l; t4++) {
        const o2 = t4 + 2,
            a2 = t4 % 2 == 0 ? "#FDF2F2" : "#FFFFFF";
        e.getRange(o2, 1, 1, 6).setBackground(a2)
    }
    e.setFrozenRows(1);
    const totalRow = l + 3;
    e.getRange(totalRow, 1).setValue("TOTAL/M\xc9DIA GERAL"), e.getRange(totalRow, 1).setFontWeight("bold"), setFormulaSafeNoApiGE_(e.getRange(totalRow, 2), `=SOMA(B2:B${l+1})`), setFormulaSafeNoApiGE_(e.getRange(totalRow, 3), "=" + gerarMediaGeralFechadaGE_(r, a)), e.getRange(totalRow, 1, 1, 6).setBackground("#E74C3C").setFontColor("white").setFontWeight("bold"), e.getRange(totalRow, 2).setNumberFormat("0"), e.getRange(totalRow, 3).setNumberFormat("0.00");
    const u2 = totalRow + 2;
    e.getRange(u2, 1).setValue("\u{1f4ca} Informa\xe7\xf5es:"), e.getRange(u2 + 1, 1).setValue("\xdaltima atualiza\xe7\xe3o: " + Utilities.formatDate(new Date, "America/Sao_Paulo", "dd/MM/yyyy HH:mm")), e.getRange(u2 + 2, 1).setValue("Total de programas: " + a + " (dia corrente exclu\xeddo dos \xedndices)"), e.getRange(u2 + 3, 1).setValue('Programas nos "\xdaltimos 5": ' + s), e.getRange(u2 + 4, 1).setValue('Programas nos "\xdaltimos 3": ' + i), e.getRange(u2, 1, 5, 1).setFontStyle("italic").setFontColor("#666666"), SpreadsheetApp.flush()
}

// --- Modelo de predição: CALIBRAGEM GS–GY ---

function atualizarCalibragemGE_(sheet) {
    var audCols = getAudColsGE_(sheet);
    var numPrograms = audCols.length;
    if (numPrograms === 0) return;

    var calibStartCol = 4 + numPrograms * 4;
    var GSl = columnToLetter_(calibStartCol);
    var GTl = columnToLetter_(calibStartCol + 1);
    var GUl = columnToLetter_(calibStartCol + 2);
    var GVl = columnToLetter_(calibStartCol + 3);
    var GWl = columnToLetter_(calibStartCol + 4);
    var GXl = columnToLetter_(calibStartCol + 5);
    var GYl = columnToLetter_(calibStartCol + 6);

    // Preserva os indicadores do dia atual mesmo após recalibrar o modelo.
    atualizarReferenciasDiaAtualHistoricoGE_(sheet);

    // Cabeçalhos row 1
    sheet.getRange(1, calibStartCol, 1, 7)
        .setValues([["BASE HIST.", "RAZÃO D/HIST", "FATOR DIA", "FATOR RECENTE", "INCLINAÇÃO", "RUPTURA", "FATOR FINAL"]])
        .setFontWeight("bold").setBackground("#2C3E50").setFontColor("white").setHorizontalAlignment("center");
    // Cabeçalhos row 2
    sheet.getRange(2, calibStartCol, 1, 7)
        .setValues([["GS", "GT", "GU", "GV", "GW", "GX", "GY"]])
        .setFontWeight("bold").setBackground("#34495E").setFontColor("white").setHorizontalAlignment("center");

    // Colunas históricas (todas exceto D=col4)
    var audColsHist = audCols.filter(function(c) { return c !== 4; });

    // C1 = previsão dinâmica: parcial_atual × (histórico_total / histórico_parcial_21min)
    setFormulaSafeNoApiGE_(sheet.getRange(1, 3), '=MÉDIA(D3:D23)*MÉDIA(SEERRO(FILTER(H63:LH63;MOD(COL(H63:LH63)-COL(H63);4)=0);""))/MÉDIA(SEERRO(FILTER(H3:LH23;MOD(COL(H3:LH23)-COL(H3);4)=0);""))');
    sheet.getRange(1, 3).setNumberFormat("0.00");

    var gsF = [], gtF = [], guF = [], gvF = [], gwF = [], gxF = [], gyF = [];

    for (var r = 3; r <= 62; r++) {
        // GS: média histórica do minuto (exclui programa mais recente)
        if (audColsHist.length === 0) {
            gsF.push(["1"]);
        } else {
            var parts = audColsHist.map(function(c) { return columnToLetter_(c) + r; }).join(";");
            gsF.push(["=MÉDIA(" + parts + ")"]);
        }

        // GT: razão real vs histórico
        gtF.push(['=SE(OU(D' + r + '=""' + ';' + GSl + r + '=0);"";D' + r + '/' + GSl + r + ')']);

        // GU: fator geral do dia (média acumulada de GT)
        if (r === 3) {
            guF.push(["1"]);
        } else {
            guF.push(['=SEERRO(MÉDIA(' + GTl + '$3:ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1));1)']);
        }

        // GV: fator recente ponderado (EMA 3 minutos)
        if (r <= 5) {
            gvF.push(['=' + GUl + r]);
        } else {
            gvF.push(['=SE(LIN()<6;' + GUl + r + ';0,5*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)+0,3*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-2)+0,2*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-3))']);
        }

        // GW: inclinação (delta GT)
        if (r <= 5) {
            gwF.push(["0"]);
        } else {
            gwF.push(['=SE(LIN()<6;0;ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)-ÍNDICE(' + GTl + ':' + GTl + ';LIN()-3))']);
        }

        // GX: detector de ruptura de regime
        if (r <= 5) {
            gxF.push(["0"]);
        } else {
            gxF.push(['=SE(LIN()<6;0;SE(E(MÉDIA(ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)-1;ÍNDICE(' + GTl + ':' + GTl + ';LIN()-2)-1;ÍNDICE(' + GTl + ':' + GTl + ';LIN()-3)-1)<=-0,08;CONT.SE(ÍNDICE(' + GTl + ':' + GTl + ';LIN()-3):ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1);"<=0,90")>=2);1;0))']);
        }

        // GY: fator final de regime (com/sem ruptura)
        if (r <= 5) {
            gyF.push(['=MÍNIMO(1,4;MÁXIMO(0,3;' + GUl + r + '))']);
        } else {
            gyF.push(['=SE(' + GXl + r + '=1;MÍNIMO(1,4;MÁXIMO(0,3;0,65*' + GVl + r + '+0,25*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)+0,10*(' + GVl + r + '+' + GWl + r + ')));MÍNIMO(1,4;MÁXIMO(0,3;0,20*' + GUl + r + '+0,45*' + GVl + r + '+0,20*ÍNDICE(' + GTl + ':' + GTl + ';LIN()-1)+0,15*(' + GVl + r + '+' + GWl + r + '))))']);
        }
    }

    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, calibStartCol,     60, 1), gsF);
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, calibStartCol + 1, 60, 1), gtF);
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, calibStartCol + 2, 60, 1), guF);
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, calibStartCol + 3, 60, 1), gvF);
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, calibStartCol + 4, 60, 1), gwF);
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, calibStartCol + 5, 60, 1), gxF);
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, calibStartCol + 6, 60, 1), gyF);

    sheet.getRange(3, calibStartCol,     60, 7).setNumberFormat("0.000");
    sheet.getRange(3, calibStartCol + 5, 60, 1).setNumberFormat("0"); // GX binário

    [80, 70, 80, 80, 70, 70, 80].forEach(function(w, i) {
        sheet.setColumnWidth(calibStartCol + i, w);
    });

    atualizarRankingTemaCalibGE_();
    atualizarDuracaoBlocosCalibGE_();
    atualizarAlertaColunasGE_(sheet);

    SpreadsheetApp.flush();
}

// --- Modelo de predição: Fator TEMA GZ–HC ---

function atualizarFatorTemaGE_(sheet) {
    var ss = SpreadsheetApp.getActive();
    var temaSheet = ss.getSheetByName("Tema Minuto");
    if (!temaSheet) return;

    var audCols = getAudColsGE_(sheet);
    var numPrograms = audCols.length;
    if (numPrograms === 0) return;

    var temaStartCol = 4 + numPrograms * 4 + 7; // logo após GS–GY
    var GZl = columnToLetter_(temaStartCol);
    var HAl = columnToLetter_(temaStartCol + 1);
    var HBl = columnToLetter_(temaStartCol + 2);
    // HC = temaStartCol + 3

    // Cabeçalhos row 1
    sheet.getRange(1, temaStartCol, 1, 4)
        .setValues([["PESO TEMA", "TROCA TEMA", "DUR. BLOCO", "FATOR TEMA"]])
        .setFontWeight("bold").setBackground("#1A5276").setFontColor("white").setHorizontalAlignment("center");
    sheet.getRange(2, temaStartCol, 1, 4)
        .setValues([["GZ", "HA", "HB", "HC"]])
        .setFontWeight("bold").setBackground("#21618C").setFontColor("white").setHorizontalAlignment("center");

    var gzF = [], haF = [], hbF = [], hcF = [];
    for (var r = 3; r <= 62; r++) {
        // GZ: peso do tema via PROCV na aba de temas
        gzF.push(['=SEERRO(PROCV(F' + r + ";'Tema Minuto'!$A:$F;6;0);1)"]);

        // HA: troca de tema em relação ao minuto anterior
        if (r === 3) {
            haF.push(["0"]);
        } else {
            haF.push(['=SE(F' + r + '=F' + (r - 1) + ';0;1)']);
        }

        // HB: duração consecutiva do bloco do tema
        if (r === 3) {
            hbF.push(["1"]);
        } else {
            hbF.push(['=SE(' + HAl + r + '=1;1;' + HBl + (r - 1) + '+1)']);
        }

        // HC: fator de persistência com ramp-up (70% → 85% → 100%)
        hcF.push(['=SE(F' + r + '="";1;SE(' + HBl + r + '<=1;MÁXIMO(0,92;' + GZl + r + '*0,92);SE(' + HBl + r + '<=3;MÁXIMO(0,96;' + GZl + r + '*0,96);' + GZl + r + ')))']);
    }

    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, temaStartCol,     60, 1), gzF);
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, temaStartCol + 1, 60, 1), haF);
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, temaStartCol + 2, 60, 1), hbF);
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, temaStartCol + 3, 60, 1), hcF);

    sheet.getRange(3, temaStartCol,     60, 4).setNumberFormat("0.000");
    sheet.getRange(3, temaStartCol + 1, 60, 1).setNumberFormat("0"); // HA binário
    sheet.getRange(3, temaStartCol + 2, 60, 1).setNumberFormat("0"); // HB inteiro

    [90, 80, 80, 90].forEach(function(w, i) {
        sheet.setColumnWidth(temaStartCol + i, w);
    });

    SpreadsheetApp.flush();
}

// =============================================================================
// gerarPrevisaoInicialGE_ — NOVO modelo isolado de previsão inicial
// =============================================================================
//
// Constrói a previsão de audiência do novo programa a partir SOMENTE dos
// programas históricos reais (colunas H em diante), sem depender das colunas
// auxiliares de calibragem GS-GY (que seriam circulares enquanto D está vazio).
//
// Para cada minuto r ∈ [3,62] computa:
//   base(r)    = média de todas as audiências históricas no minuto
//   f_trend(r) = clamp(média últimos 5 / base, [0.75, 1.30])
//   f_accel(r) = clamp(média últimos 3 / média últimos 5, [0.85, 1.18])
//   f_dia(r)   = clamp(média mesmo dia-da-semana / base, [0.78, 1.28])
//
// Combina pelo regime aplicável:
//   • Normal: 0.40·f_dia + 0.35·f_trend + 0.25·f_accel
//   • Ruptura (|últimos_3 / restante − 1| ≥ 15%): blend 50/50 com
//                0.20·f_dia + 0.55·f_trend + 0.25·f_accel
//
// Escreve em D3:D62: `=ARRED(<base*fator_combinado> * HC{r}, 2)`
// → HC (fator tema) permanece dinâmico, reagindo quando F é preenchido.
//
// Retorna `true` se preencheu previsão, `false` se não há histórico suficiente.
// -----------------------------------------------------------------------------
function gerarPrevisaoInicialGE_(sheet, programDate) {
    var audCols = getAudColsGE_(sheet);
    // audCols[0]=4 é o novo dia (vazio); demais são histórico real
    var audColsHist = audCols.filter(function(c) { return c !== 4; });

    if (audColsHist.length === 0) {
        Logger.log("[Prev Inicial] Sem programas históricos — D3:D62 fica vazio");
        return false;
    }

    // Identificar programas do mesmo dia da semana
    var weekday = programDate.getDay();
    var lastCol = sheet.getLastColumn();
    var row1Vals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    var colsMesmoDia = [];
    for (var i = 0; i < audColsHist.length; i++) {
        var c = audColsHist[i];
        var v = row1Vals[c - 1];
        var dParsed = v instanceof Date
            ? v
            : parseDateStr_(typeof v === "string" ? v.trim() : "");
        if (dParsed && dParsed.getDay() === weekday) colsMesmoDia.push(c);
    }

    // Últimos N programas (audColsHist está em ordem newest-first, pois col 8 = mais recente)
    var colsRecent5 = audColsHist.slice(0, Math.min(PREV_INICIAL_GE.N_RECENT_5, audColsHist.length));
    var colsRecent3 = audColsHist.slice(0, Math.min(PREV_INICIAL_GE.N_RECENT_3, audColsHist.length));

    // Sets para lookup O(1)
    var recent5Set = {};
    for (var k1 = 0; k1 < colsRecent5.length; k1++) recent5Set[colsRecent5[k1]] = true;
    var recent3Set = {};
    for (var k2 = 0; k2 < colsRecent3.length; k2++) recent3Set[colsRecent3[k2]] = true;
    var diaSet = {};
    for (var k3 = 0; k3 < colsMesmoDia.length; k3++) diaSet[colsMesmoDia[k3]] = true;

    // Leitura em batch da matriz histórica (rows 3-62 × colunas audColsHist)
    var minCol = audColsHist[0];                          // col 8 (mais recente)
    var maxCol = audColsHist[audColsHist.length - 1];     // col mais antiga (mais à direita)
    var blockData = sheet.getRange(3, minCol, 60, maxCol - minCol + 1).getValues();

    // Localizar posição de HC (fator tema) — fórmula precisa apontar dinamicamente
    var calibStartCol = 4 + audCols.length * 4;
    var temaSheetExists = !!SpreadsheetApp.getActive().getSheetByName("Tema Minuto");
    var hcColLetter = temaSheetExists ? columnToLetter_(calibStartCol + 7 + 3) : null;

    // Telemetria
    var stats = {
        cols_historicas: audColsHist.length,
        cols_mesmo_dia: colsMesmoDia.length,
        cols_recent5: colsRecent5.length,
        cols_recent3: colsRecent3.length,
        weekday: weekday,
        minutos_preenchidos: 0,
        minutos_vazios: 0,
        minutos_ruptura: 0,
        minutos_base_only: 0,
        com_hc: temaSheetExists
    };

    var formulas = [];
    for (var r = 3; r <= 62; r++) {
        var rowIdx = r - 3;

        var allVals = [], recent5Vals = [], recent3Vals = [], diaVals = [];

        for (var ii = 0; ii < audColsHist.length; ii++) {
            var colH = audColsHist[ii];
            var colIdx = colH - minCol;
            var val = blockData[rowIdx][colIdx];
            if (typeof val !== "number" || val <= 0) continue;
            allVals.push(val);
            if (recent5Set[colH]) recent5Vals.push(val);
            if (recent3Set[colH]) recent3Vals.push(val);
            if (diaSet[colH]) diaVals.push(val);
        }

        if (allVals.length === 0) {
            formulas.push([""]);
            stats.minutos_vazios++;
            continue;
        }

        var base = mediaArrGE_(allVals);

        // f_trend: média recente 5 / base
        var fTrend = 1;
        if (recent5Vals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_TREND && base > 0) {
            fTrend = clampGE_(mediaArrGE_(recent5Vals) / base,
                PREV_INICIAL_GE.CLAMP_TREND[0], PREV_INICIAL_GE.CLAMP_TREND[1]);
        }

        // f_accel: média recente 3 / média recente 5 (captura aceleração)
        var fAccel = 1;
        if (recent3Vals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_ACCEL &&
            recent5Vals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_TREND) {
            var mr5 = mediaArrGE_(recent5Vals);
            if (mr5 > 0) {
                fAccel = clampGE_(mediaArrGE_(recent3Vals) / mr5,
                    PREV_INICIAL_GE.CLAMP_ACCEL[0], PREV_INICIAL_GE.CLAMP_ACCEL[1]);
            }
        }

        // f_dia: média mesmo dia / base
        var fDia = 1;
        if (diaVals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_DIA && base > 0) {
            fDia = clampGE_(mediaArrGE_(diaVals) / base,
                PREV_INICIAL_GE.CLAMP_DIA[0], PREV_INICIAL_GE.CLAMP_DIA[1]);
        }

        // Combinação ponderada regime normal
        var fatorNormal = PREV_INICIAL_GE.PESO_DIA   * fDia
                        + PREV_INICIAL_GE.PESO_TREND * fTrend
                        + PREV_INICIAL_GE.PESO_ACCEL * fAccel;

        var fatorCombinado = fatorNormal;

        // Pouco histórico: só usa base (sem fator)
        if (recent5Vals.length < PREV_INICIAL_GE.MIN_AMOSTRAS_TREND &&
            diaVals.length < PREV_INICIAL_GE.MIN_AMOSTRAS_DIA) {
            fatorCombinado = 1;
            stats.minutos_base_only++;
        } else {
            // Detector de ruptura: últimos 3 vs restante
            if (recent3Vals.length >= 3 && allVals.length >= PREV_INICIAL_GE.MIN_AMOSTRAS_RUPTURA) {
                var sumAll = 0, sumRec3 = 0;
                for (var sa = 0; sa < allVals.length; sa++) sumAll += allVals[sa];
                for (var sb = 0; sb < recent3Vals.length; sb++) sumRec3 += recent3Vals[sb];
                var nRest = allVals.length - recent3Vals.length;
                if (nRest > 0) {
                    var meanRest = (sumAll - sumRec3) / nRest;
                    if (meanRest > 0) {
                        var ratioRupt = (sumRec3 / recent3Vals.length) / meanRest;
                        if (ratioRupt <= 1 - PREV_INICIAL_GE.THRESHOLD_RUPTURA ||
                            ratioRupt >= 1 + PREV_INICIAL_GE.THRESHOLD_RUPTURA) {
                            var fatorRupt = PREV_INICIAL_GE.PESO_DIA_RUPT   * fDia
                                          + PREV_INICIAL_GE.PESO_TREND_RUPT * fTrend
                                          + PREV_INICIAL_GE.PESO_ACCEL_RUPT * fAccel;
                            fatorCombinado = PREV_INICIAL_GE.BLEND_RUPTURA * fatorNormal
                                           + (1 - PREV_INICIAL_GE.BLEND_RUPTURA) * fatorRupt;
                            stats.minutos_ruptura++;
                        }
                    }
                }
            }
        }

        var predValue = base * fatorCombinado;

        // Construir fórmula (locale BR: vírgula decimal, ponto-e-vírgula separador)
        var predStr = predValue.toFixed(4).replace(".", ",");
        var formula = (temaSheetExists && hcColLetter)
            ? '=ARRED(' + predStr + '*' + hcColLetter + r + ';2)'
            : '=ARRED(' + predStr + ';2)';
        formulas.push([formula]);
        stats.minutos_preenchidos++;
    }

    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, 4, 60, 1), formulas);
    SpreadsheetApp.flush();

    Logger.log("[Prev Inicial GE] " + JSON.stringify(stats));
    return true;
}

// --- Atualizar calibragem manualmente (entrada pública via menu) ---

function atualizarCalibragemGE() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    atualizarCalibragemGE_(sheet);
    atualizarFatorTemaGE_(sheet);
    SpreadsheetApp.getUi().alert(
        "✅ Calibragem Atualizada!",
        "Colunas GS–GY (modelo de predição) regeneradas com sucesso.",
        SpreadsheetApp.getUi().ButtonSet.OK
    );
}

// =============================================================================
// adicionarNovoDiaManualGE — REORGANIZADO (entrada pública via menu)
// =============================================================================
//
// Ordem de execução:
//   1.  Inserir bloco D:G vazio
//   1b. Devolver o bloco mensal (que a inserção empurrou para H:I) para D:E
//   2.  atualizarFormulasBaseGE_ (apenas B/C/row 63/share — SEM calibragem)
//   3.  Aplicar fórmula de share em E
//   4.  gerarPrevisaoInicialGE_ — preenche D3:D62 com modelo isolado
//   5.  atualizarCalibragemGE_ + atualizarFatorTemaGE_ — agora com D preenchido
//   6.  Bloco row 63-65 (mediana padrão OFFSET, comentaristas, diferença)
//   7.  Formatação condicional herdada
//   8.  Atualizar dashboards (painel preservado: D63 injetado em semana/mês,
//       com criação automática de nova semana às segundas e novo mês na virada)
//
// Isto elimina a circularidade GS×GY×HC que ocorria quando D estava vazio.
// -----------------------------------------------------------------------------
// Núcleo silencioso: cria o bloco do dia no Histórico com previsão inicial.
// Usado pela rotina automática das 06:00 e pelo fluxo manual.
// Retorna { ok, motivo, temPrevisao }.
function adicionarNovoDiaCoreGE_(programDate) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return { ok: false, motivo: "aba_inexistente" };

    var dateStr = Utilities.formatDate(programDate, "America/Sao_Paulo", "dd/MM/yyyy");

    // Duplicata: se o dia já existe, não faz nada
    var lastCol = sheet.getLastColumn();
    var numProgramsBefore = Math.floor((lastCol - 3) / 4);
    if (numProgramsBefore > 0) {
        var existingDates = sheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
        for (var dPos = 0; dPos < existingDates.length; dPos += 4) {
            var existing = existingDates[dPos];
            if (!existing) continue;
            var existStr = existing instanceof Date
                ? Utilities.formatDate(existing, "America/Sao_Paulo", "dd/MM/yyyy")
                : existing.toString().trim();
            if (existStr === dateStr) return { ok: false, motivo: "duplicada" };
        }
    }

    // ❶a Congela o dia que está sendo arquivado: célula de D com FÓRMULA é
    //    projeção ao vivo de minuto nunca capturado — não é dado real. Se a
    //    fórmula fosse arquivada, a inserção de colunas a deslocaria junto e
    //    ela viraria referência circular (#REF!) na coluna histórica,
    //    contaminando B/C/E e a calibragem em cascata.
    var dFormsOld = sheet.getRange(3, 4, 60, 1).getFormulas();
    for (var fz = 0; fz < 60; fz++) {
        if (dFormsOld[fz][0]) sheet.getRange(3 + fz, 4).clearContent();
    }

    // ❶ Inserir bloco vazio D:G
    sheet.insertColumns(4, 4);

    // ❶b FIX: a inserção empurra o bloco mensal do painel (D:E) para H:I.
    //     Devolve para D:E ANTES de qualquer outra escrita — senão os valores
    //     manuais 0,66/0,65/... e o label MENSAL se perdem.
    restaurarBlocoMensalAposInsercaoGE_(sheet);

    var bgColor = (numProgramsBefore % 2 === 0) ? "#FDF2F2" : "#FFFFFF";

    sheet.getRange(1, 4, 1, 4).merge();
    sheet.getRange(1, 4).setValue(dateStr)
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setBackground("#FADBD8")
        .setFontColor("#000000");

    sheet.getRange(2, 4, 1, 4).setValues([["AUDIÊNCIA", "&/MIN", "TEMA", "TERMOS"]]);
    sheet.getRange(2, 4, 1, 4)
        .setFontWeight("bold")
        .setBackground("#E74C3C")
        .setFontColor("white")
        .setHorizontalAlignment("center");

    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 70);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 200);

    sheet.getRange(3, 4, 60, 4).setBackground(bgColor);

    sheet.getRange(3, 4, 60, 1).setNumberFormat("0.00");
    sheet.getRange(3, 5, 60, 1).setNumberFormat("0.00");

    // ❷ Atualizar SOMENTE fórmulas-base (B/C/row63) — SEM calibragem
    //    Importante: nesta etapa NÃO chamamos atualizarCalibragemGE_, pois com
    //    D ainda vazio teríamos GT="" e GY≈1 (modelo degenerado).
    atualizarFormulasBaseGE_(sheet);
    Utilities.sleep(800);

    // ❸ Definir fórmula de share em E (=D/B) — compatível com a previsão futura
    var shareFormulas = [];
    for (var row = 3; row <= 62; row++) {
        shareFormulas.push(['=SE(OU(D' + row + '="";$B' + row + '=0);"";D' + row + '/$B' + row + ')']);
    }
    batchSetFormulasSafeNoApiGE_(sheet.getRange(3, 5, 60, 1), shareFormulas);

    // ❹ PREVISÃO INICIAL ROBUSTA usando modelo isolado de histórico
    var temPrevisao = gerarPrevisaoInicialGE_(sheet, programDate);
    Utilities.sleep(800);

    // ❺ Agora SIM chamamos calibragem (D preenchido com previsão → GT-GY válidos)
    atualizarCalibragemGE_(sheet);
    atualizarFatorTemaGE_(sheet);
    Utilities.sleep(500);

    // ❻ Bloco row 63-65: espelha formato do bloco anterior H63:K65, labels, mediana
    sheet.getRange(63, 8, 3, 4).copyTo(sheet.getRange(63, 4, 3, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    sheet.getRange(64, 4, 2, 1).setValues([["COMENTARISTA 1"], ["COMENTARISTA 2"]]);
    sheet.getRange(64, 6).setValue("DIFERENÇA");

    // Mediana por dia útil (terça=3ª, quarta=4ª, quinta=5ª, sexta=6ª)
    var weekday = programDate.getDay();
    var medianaLabel = {
        2: "MEDIANA 3ª",
        3: "MEDIANA 4ª",
        4: "MEDIANA 5ª",
        5: "MEDIANA 6ª"
    }[weekday] || "MEDIANA";
    sheet.getRange(64, 5).setValue(medianaLabel);

    // FIX: mediana no padrão OFFSET da planilha —
    //   =(MEDIAN(OFFSET(X63,0,0,1,1),OFFSET(X63,0,20),...)) com 14 termos,
    // ancorada na coluna do mesmo dia da semana mais recente. Gravada com
    // Range.setFormula nativo (sintaxe en-US, independente do locale).
    var medianaFormula = gerarFormulaMedianaOffsetGE_(sheet, programDate);
    if (medianaFormula) {
        // FIX: setFormulaSafeNoApiGE_ (USER_ENTERED via Sheets API) — respeita locale
        // pt-BR e interpreta MED/DESLOC/; corretamente.
        setFormulaSafeNoApiGE_(sheet.getRange(63, 5), medianaFormula);
    } else {
        sheet.getRange(63, 5).setValue("");
    }
    setFormulaSafeNoApiGE_(sheet.getRange(63, 6), "=(D63/E63)-1");
    // FIX: NÃO forçar setNumberFormat em F63 — o formato (0.00%) já veio do
    // copyTo PASTE_FORMAT do bloco anterior.

    // ❼ Herdar formatação condicional da coluna equivalente (&/MIN) do bloco anterior: I3:I62 -> E3:E62
    var rules = sheet.getConditionalFormatRules();
    var clonedRules = [];
    var sourceCol = 9, targetCol = 5, startRow = 3, numRows = 60;
    var shouldUpdateRules = false;
    for (var rIdx = 0; rIdx < rules.length; rIdx++) {
        var rule = rules[rIdx];
        var ranges = rule.getRanges();
        var hasSource = false, hasTarget = false;
        for (var rgIdx = 0; rgIdx < ranges.length; rgIdx++) {
            var rg = ranges[rgIdx];
            if (rg.getColumn() === sourceCol && rg.getRow() === startRow && rg.getNumColumns() === 1 && rg.getNumRows() === numRows) hasSource = true;
            if (rg.getColumn() === targetCol && rg.getRow() === startRow && rg.getNumColumns() === 1 && rg.getNumRows() === numRows) hasTarget = true;
        }
        if (hasSource && !hasTarget) {
            var newRanges = ranges.slice();
            newRanges.push(sheet.getRange(startRow, targetCol, numRows, 1));
            clonedRules.push(rule.copy().setRanges(newRanges).build());
            shouldUpdateRules = true;
        } else {
            clonedRules.push(rule);
        }
    }
    if (shouldUpdateRules) sheet.setConditionalFormatRules(clonedRules);

    // ❽ Dashboards (painel A66+ preservado — D63 injetado na semana/mês,
    //    com criação automática de nova semana e novo mês quando virar)
    garantirCamposHistoricoGE_();
    atualizarMediaSemanaGE_();
    Utilities.sleep(1000);
    atualizarGraficoMediasGE_();
    Utilities.sleep(1000);
    atualizarMediaPorComentaristaGE_();
    atualizarMediaPorDiaSemanaGE_();
    // FIX: NÃO chamar atualizarMediaPorTemaSeExistir_() aqui. Essa função
    // faz sheet.clear() na aba "Tema Minuto" e, se a escrita
    // posterior das fórmulas falhar por quota (acumulado de chamadas no
    // adicionarNovoDiaManualGE), a aba fica zerada. A atualização da aba TEMA
    // agora é responsabilidade explícita do menu ou do fluxo Completa.

    SpreadsheetApp.flush();

    return { ok: true, temPrevisao: temPrevisao };
}

// =============================================================================
// VIRADA DIÁRIA LEVE E REPARÁVEL
// =============================================================================
// A rotina antiga de 06:00 criava o novo bloco somente depois de executar
// várias reconstruções pesadas. Se a execução atingisse o limite do Apps
// Script durante adicionarNovoDiaCoreGE_, as colunas D:G já estavam inseridas,
// mas fórmulas, rodapé e painel ficavam incompletos. A função abaixo é curta,
// idempotente e pode reparar exatamente esse estado parcial sem apagar leituras
// reais que eventualmente já tenham chegado ao bloco do dia.

function dataBlocoHistoricoIgualGE_(valor, dataStr) {
    if (!valor) return false;
    var s = valor instanceof Date
        ? Utilities.formatDate(valor, CFG.TIMEZONE, "dd/MM/yyyy")
        : String(valor).trim();
    return s === dataStr;
}

function localizarColunaDataHistoricoGE_(sheet, dataStr) {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 4) return -1;
    var vals = sheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
    for (var i = 0; i < vals.length; i += 4) {
        if (dataBlocoHistoricoIgualGE_(vals[i], dataStr)) return 4 + i;
    }
    return -1;
}

function clonarRegraShareNovoDiaGE_(sheet) {
    var rules = sheet.getConditionalFormatRules();
    var out = [];
    var alterou = false;
    for (var i = 0; i < rules.length; i++) {
        var rule = rules[i];
        var ranges = rule.getRanges();
        var temOrigem = false;
        var temDestino = false;
        for (var j = 0; j < ranges.length; j++) {
            var rg = ranges[j];
            if (rg.getRow() === 3 && rg.getNumRows() === 60 && rg.getNumColumns() === 1) {
                if (rg.getColumn() === 9) temOrigem = true;   // I3:I62
                if (rg.getColumn() === 5) temDestino = true; // E3:E62
            }
        }
        if (temOrigem && !temDestino) {
            var novos = ranges.slice();
            novos.push(sheet.getRange(3, 5, 60, 1));
            out.push(rule.copy().setRanges(novos).build());
            alterou = true;
        } else {
            out.push(rule);
        }
    }
    if (alterou) sheet.setConditionalFormatRules(out);
}

function criarOuRepararBlocoDiaAtualGE_(programDate) {
    programDate = programDate || new Date();
    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(30000)) throw new Error("Não foi possível obter o lock da virada diária.");

    try {
        var ss = SpreadsheetApp.getActive();
        var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
        if (!sheet) throw new Error('A aba "' + CFG.GE_OUT_HISTORICO + '" não existe.');

        var dataStr = Utilities.formatDate(programDate, CFG.TIMEZONE, "dd/MM/yyyy");
        var colDia = localizarColunaDataHistoricoGE_(sheet, dataStr);
        var criou = false;

        // Saída rápida: se o bloco já foi reparado, não refaz copyTo,
        // formatação condicional, 60 fórmulas e painel. Isso evita o timeout
        // do serviço Planilhas ao repetir manualmente a correção.
        if (colDia === 4) {
            var fD63 = sheet.getRange(63, 4).getFormula();
            var fE3 = sheet.getRange(3, 5).getFormula();
            var labelD64 = String(sheet.getRange(64, 4).getDisplayValue() || "");
            var painel = localizarPainelGE_(sheet);
            var fSemana = painel && painel.rowSemana > 0
                ? sheet.getRange(painel.rowSemana, 2).getFormula() : "";
            var fMes = painel && painel.rowMesAtual > 0
                ? sheet.getRange(painel.rowMesAtual, 2).getFormula() : "";
            if (fD63 && fE3 && labelD64 && /\bD63\b/.test(fSemana) && /\bD63\b/.test(fMes)) {
                atualizarReferenciasDiaAtualHistoricoGE_(sheet);
                try { logGE_("VIRADA_DIA", dataStr + " | referências A1/B1/C/E atualizadas para D; reparo pesado ignorado"); } catch (eFast) { }
                return {
                    ok: true,
                    criou: false,
                    reparou: false,
                    jaEstavaOk: true,
                    data: dataStr,
                    detalhe: "bloco já íntegro; referências do dia atual revalidadas"
                };
            }
        }

        if (colDia < 0) {
            // Congela o bloco anterior em uma única operação: leituras reais
            // permanecem; projeções que ainda eram fórmulas são descartadas.
            var oldRange = sheet.getRange(3, 4, 60, 1);
            var oldValues = oldRange.getValues();
            var oldForms = oldRange.getFormulas();
            var congelado = [];
            for (var r = 0; r < 60; r++) {
                congelado.push([oldForms[r][0] ? "" : oldValues[r][0]]);
            }
            oldRange.setValues(congelado);

            sheet.insertColumns(4, 4);
            if (typeof restaurarBlocoMensalAposInsercaoGE_ === "function") {
                restaurarBlocoMensalAposInsercaoGE_(sheet);
            }
            colDia = 4;
            criou = true;
        }

        if (colDia !== 4) {
            throw new Error("O bloco de " + dataStr + " existe fora da posição mais recente (coluna " + colDia + ").");
        }

        // Estrutura e formatação. PASTE_FORMAT nunca altera leituras já gravadas.
        if (sheet.getLastColumn() >= 11) {
            sheet.getRange(3, 8, 63, 4).copyTo(
                sheet.getRange(3, 4, 63, 4),
                SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
                false
            );
        }
        try { sheet.getRange(1, 4, 1, 4).breakApart(); } catch (e0) { }
        sheet.getRange(1, 4, 1, 4).merge();
        sheet.getRange(1, 4).setValue(dataStr)
            .setHorizontalAlignment("center")
            .setFontWeight("bold")
            .setBackground("#FADBD8")
            .setFontColor("#000000");
        sheet.getRange(2, 4, 1, 4).setValues([["AUDIÊNCIA", "&/MIN", "TEMA", "TERMOS"]])
            .setFontWeight("bold")
            .setBackground("#E74C3C")
            .setFontColor("white")
            .setHorizontalAlignment("center");
        sheet.setColumnWidth(4, 90);
        sheet.setColumnWidth(5, 70);
        sheet.setColumnWidth(6, 120);
        sheet.setColumnWidth(7, 200);
        sheet.getRange(3, 4, 60, 1).setNumberFormat("0.00");
        sheet.getRange(3, 5, 60, 1).setNumberFormat("0.00");

        // Após inserir D:G, o Sheets desloca as referências antigas para H.
        // Reancora os indicadores em D sem recalcular B nem alterar audiência.
        atualizarReferenciasDiaAtualHistoricoGE_(sheet);

        // Rodapé do programa. D63/F63 permanecem vazios antes da primeira leitura.
        setFormulaSafeNoApiGE_(sheet.getRange(63, 4), '=SEERRO(MÉDIA(D3:D62);"")');
        var medianaFormula = gerarFormulaMedianaOffsetGE_(sheet, programDate);
        if (medianaFormula) setFormulaSafeNoApiGE_(sheet.getRange(63, 5), medianaFormula);
        else sheet.getRange(63, 5).clearContent();
        setFormulaSafeNoApiGE_(sheet.getRange(63, 6),
            '=SE(OU(D63="";E63="";E63=0);"";D63/E63-1)');

        var weekday = programDate.getDay();
        var medianaLabel = { 2: "MEDIANA 3ª", 3: "MEDIANA 4ª", 4: "MEDIANA 5ª", 5: "MEDIANA 6ª" }[weekday] || "MEDIANA";
        sheet.getRange(64, 4, 2, 1).setValues([["COMENTARISTA 1"], ["COMENTARISTA 2"]]);
        sheet.getRange(64, 5).setValue(medianaLabel);
        sheet.getRange(64, 6).setValue("DIFERENÇA");

        // A regra de cor do share reúne centenas de faixas. Regravá-la dentro
        // da virada crítica pode consumir vários minutos. A estrutura funcional
        // fica pronta agora; a regra visual é concluída por acionador isolado.
        try {
            PropertiesService.getScriptProperties().setProperty(
                "ge_formatar_virada_pendente", dataStr);
        } catch (ePendente) { }

        // Inclui D63 na semana e no mês atuais. A função é idempotente e
        // também cria uma nova semana/mês quando necessário.
        garantirCamposHistoricoGE_();

        var detalhe = (criou ? "criado" : "reparado") +
            " | D63/F63 protegidos | semana/mês incluem D63";
        try { logGE_("VIRADA_DIA", dataStr + " | " + detalhe); } catch (e1) { }
        return { ok: true, criou: criou, reparou: !criou, data: dataStr, detalhe: detalhe };
    } finally {
        lock.releaseLock();
    }
}

// Executa fora do caminho crítico das 06:00. É idempotente e não altera
// valores/fórmulas do programa — apenas garante a regra visual do share E.
function finalizarFormatacaoViradaGE_() {
    var props = PropertiesService.getScriptProperties();
    var pendente = String(props.getProperty("ge_formatar_virada_pendente") || "");
    if (!pendente) return { ok: true, skipped: true };
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) throw new Error("Aba Histórico não encontrada.");
    clonarRegraShareNovoDiaGE_(sheet);
    props.deleteProperty("ge_formatar_virada_pendente");
    try { logGE_("VIRADA_FORMATO", pendente + " | regra de share concluída"); } catch (eLog) { }
    return { ok: true, data: pendente };
}

// Entrada manual via menu (com prompts)
function adicionarNovoDiaManualGE() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActive();
    if (!ss.getSheetByName(CFG.GE_OUT_HISTORICO)) {
        ui.alert("❌ Erro", 'A aba "Histórico" não existe.', ui.ButtonSet.OK);
        return;
    }

    var resp = ui.prompt(
        "📅 Adicionar Novo Dia",
        "Digite a data do programa (dd/mm/aaaa):\n\nA previsão inicial será calculada a partir do histórico real, sem depender de dados ao vivo.",
        ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    var dateStr = resp.getResponseText().trim();
    var match = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!match) {
        ui.alert("❌ Formato inválido", "Use dd/mm/aaaa", ui.ButtonSet.OK);
        return;
    }
    var programDate = new Date(parseInt(match[3], 10), parseInt(match[2], 10) - 1, parseInt(match[1], 10));

    var res = adicionarNovoDiaCoreGE_(programDate);
    if (!res.ok) {
        ui.alert(res.motivo === "duplicada" ? "⚠️ Duplicada" : "❌ Erro",
            res.motivo === "duplicada" ? dateStr + " já existe." : ("Falha: " + res.motivo),
            ui.ButtonSet.OK);
        return;
    }

    var msgPrev = res.temPrevisao
        ? "📊 Previsão inicial aplicada (modelo isolado de histórico)\n"
        : "⚠️ Sem histórico suficiente — D3:D62 vazio\n";

    ui.alert("✅ Novo dia adicionado!",
        "Programa de " + dateStr + " inserido!\n\n" +
        msgPrev +
        "📈 Calibragem GS–GY recalculada\n" +
        "📈 Painel semana/mês atualizado\n" +
        "📈 Gráfico atualizado\n\n" +
        "⚠️ Preencha:\n" +
        "• Linha 64 (col D): Comentarista 1\n" +
        "• Linha 65 (col D): Comentarista 2\n" +
        "• Coluna F (TEMA): preenchimento atualiza HC dinamicamente",
        ui.ButtonSet.OK
    );
}

// --- Atualizar Média por Tema se aba existir ---

function atualizarMediaPorTemaSeExistir_() {
    var ss = SpreadsheetApp.getActive();
    var temaSheet = ss.getSheetByName("Tema Minuto");
    if (!temaSheet) return;
    var histSheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!histSheet) return;
    var numPrograms = getAudColsGE_(histSheet).length;
    if (numPrograms === 0) return;
    normalizarTemasNoHistoricoGE_();
    var temas = coletarTemasUnicosGE_(histSheet, numPrograms);
    if (temas.length === 0) return;
    temaSheet.clear();
    configurarAbaMediaPorTema_(temaSheet, histSheet, temas, numPrograms);
}

// --- Ver estrutura do histórico GE (entrada pública via menu) ---

function verEstruturaHistoricoGazeta() {
    const e = SpreadsheetApp.getUi();
    e.alert("\u{1f4ca} Estrutura do Hist\xf3rico GE", "\n\u{1f4ca} ESTRUTURA DA PLANILHA GE - HIST\xd3RICO\n\n\u250c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2510\n\u2502 Linha 1 \u2502     \u2502       \u2502        \u2502   17/01/2026    \u2502   16/01/2026    \u2502  ... \u2502\n\u251c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2524\n\u2502 Linha 2 \u2502HORA \u2502 M\xc9DIA \u2502%A M\xc9DIA\u2502 AUD\u2502SHA\u2502TEMA\u2502TER\u2502 AUD\u2502SHA\u2502TEMA\u2502TER\u2502  ... \u2502\n\u251c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u253c\u2500\u2500\u2500\u2500\u2500\u2500\u2524\n\u2502 Linha 3 \u250218:00\u2502 0.45  \u2502 +2.3%  \u25020.46\u25021.2\u2502    \u2502   \u25020.44\u25021.1\u2502    \u2502   \u2502      \u2502\n\u2502 Linha 4 \u250218:01\u2502 0.47  \u2502 +1.4%  \u25020.48\u25021.3\u2502SPFC\u2502...\u25020.46\u25021.2\u2502PALM\u2502...\u2502      \u2502\n\u2502   ...   \u2502 ... \u2502  ...  \u2502  ...   \u2502 ...\u2502...\u2502 ...\u2502...\u2502 ...\u2502...\u2502 ...\u2502...\u2502      \u2502\n\u2502Linha 62 \u250218:59\u2502 0.55  \u2502 -0.5%  \u25020.54\u25021.5\u2502CORI\u2502...\u25020.56\u25021.6\u2502FLAM\u2502...\u2502      \u2502\n\u2514\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2518\n\nCOLUNAS FIXAS:\n\u2022 A: HORA (18:00 a 18:59)\n\u2022 B: M\xc9DIA de todas as audi\xeancias hist\xf3ricas daquele minuto\n\u2022 C: % A M\xc9DIA = (Audi\xeancia mais recente - M\xe9dia) / M\xe9dia\n\nBLOCOS DE 4 COLUNAS (por programa):\n\u2022 AUDI\xcaNCIA: valor em formato 0.00\n\u2022 SHARE: valor em formato 0.0\n\u2022 TEMA: tema classificado do minuto\n\u2022 TERMOS: top 5 termos do minuto\n\nCORES:\n\u2022 Cabe\xe7alhos: vermelho (#E74C3C) - tema do Gazeta\n\u2022 Programas: altern\xe2ncia vermelho claro / branco\n", e.ButtonSet.OK)
}

// =============================================================================
// executarAtualizacaoHistoricoGECompleta — fluxo completo num clique
// =============================================================================
// Executa, na ordem correta:
//   1. Atualizar Média por Tema (cria a aba se não existir)
//   2. Adicionar Novo Dia (Manual) — pergunta a data
//   3. Calibragem GS–GY + Fator Tema GZ–HC
//   4. Instalar/Atualizar Projeção de Audiência
// -----------------------------------------------------------------------------
function executarAtualizacaoHistoricoGECompleta() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActive();
    var etapas = [];

    // NOVA ORDEM (fix 7+8):
    //   1. Adicionar Novo Dia (Manual) — pergunta a data; já faz base +
    //      calibragem + fator tema + painel internos. NÃO atualiza TEMA aqui
    //      (fix 7 — evita clear+falha-de-quota).
    //   2. Calibragem + Fator Tema (redundante, garantia).
    //   3. Atualizar Média por Tema — agora com N+1 programas; in-place,
    //      preserva fórmulas anteriores se a aba já tem temas.
    //   4. Instalar/Atualizar Projeção de Audiência — batchUpdate em 1
    //      chamada (fix 1 — quota).

    try {
        adicionarNovoDiaManualGE();
        etapas.push("1. Novo Dia ✅");
        SpreadsheetApp.flush();
    } catch (e1) {
        etapas.push("1. Novo Dia ❌ " + e1.message);
        ui.alert("⚠️ Etapa 1 falhou", e1.message + "\n\nO fluxo foi interrompido.", ui.ButtonSet.OK);
        return;
    }

    try {
        var sheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
        if (sheet) {
            atualizarCalibragemGE_(sheet);
            atualizarFatorTemaGE_(sheet);
        }
        etapas.push("2. Calibragem ✅");
        SpreadsheetApp.flush();
        Utilities.sleep(500); // folga antes da próxima fase pesada (quota)
    } catch (e2) {
        etapas.push("2. Calibragem ❌ " + e2.message);
    }

    try {
        var resTema;
        if (ss.getSheetByName("Tema Minuto")) {
            resTema = atualizarFormulasTemaGE_silencioso_();
        } else {
            criarMediaPorTemaGE();
            resTema = { ok: true };
        }
        if (resTema && resTema.ok) {
            etapas.push("3. Média por Tema ✅" + (resTema.temas ? " (" + resTema.temas + " temas, " + resTema.programs + " programas)" : ""));
        } else {
            etapas.push("3. Média por Tema ⚠️ " + ((resTema && resTema.msg) ? resTema.msg : "falhou"));
        }
        SpreadsheetApp.flush();
        Utilities.sleep(500);
    } catch (e3) {
        etapas.push("3. Média por Tema ❌ " + e3.message);
    }

    try {
        var res = instalarProjecaoAudienciaGE_();
        if (res && res.ok) {
            etapas.push("4. Projeção ✅ (" + res.escritas + " projetadas, " + res.preservadas + " reais preservados)");
        } else {
            etapas.push("4. Projeção ⚠️ " + ((res && res.msg) ? res.msg : "falhou"));
        }
        SpreadsheetApp.flush();
    } catch (e4) {
        etapas.push("4. Projeção ❌ " + e4.message);
    }

    ui.alert("📋 Atualização Completa GE - Histórico", etapas.join("\n"), ui.ButtonSet.OK);
}

// --- Versão silenciosa de atualizarFormulasTemaGE (sem alerts) ---
// Para uso interno em fluxos automatizados como executarAtualizacaoHistoricoGECompleta.
// Lança exceção em vez de mostrar UI alert; retorna {ok, temas, programs, msg}.
function atualizarFormulasTemaGE_silencioso_() {
    normalizarTemasNoHistoricoGE_();
    var lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        var ss = SpreadsheetApp.getActive();
        var temaSheet = ss.getSheetByName("Tema Minuto");
        if (!temaSheet) return { ok: false, msg: "aba TEMA não existe" };
        var histSheet = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
        if (!histSheet) return { ok: false, msg: "aba histórico não existe" };
        var numPrograms = getAudColsGE_(histSheet).length;
        if (numPrograms === 0) return { ok: false, msg: "sem programas no histórico" };
        var temas = coletarTemasUnicosGE_(histSheet, numPrograms);
        if (temas.length === 0) return { ok: false, msg: "sem TEMAS no histórico" };
        var lastRowBefore = temaSheet.getLastRow();
        configurarAbaMediaPorTema_(temaSheet, histSheet, temas, numPrograms);
        var rodapeEnd = temas.length + 8;
        if (lastRowBefore > rodapeEnd) {
            temaSheet.getRange(rodapeEnd + 1, 1, lastRowBefore - rodapeEnd, 6).clearContent().clearFormat();
        }
        return { ok: true, temas: temas.length, programs: numPrograms };
    } finally {
        lock.releaseLock();
    }
}

// --- Rodar UMA vez para reescrever B/C com a correção (pode ficar no código) ---
function corrigirColunaBGE() {
    var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!sheet) return;
    atualizarFormulasBaseGE_(sheet);
    SpreadsheetApp.getUi().alert("✅ Coluna B corrigida",
        "B agora usa apenas dias fechados (sem D). Teste preencher o TEMA em F.",
        SpreadsheetApp.getUi().ButtonSet.OK);
}

// Indicadores do dia corrente: reaplicados após a inserção do bloco D:G e
// nas repetições da rotina matinal, sem tocar em B3:B62 ou na audiência.
function atualizarReferenciasDiaAtualHistoricoGE_(sheet) {
    if (!sheet) throw new Error("Aba Histórico não encontrada.");
    var top = [[
        '=ROUND(0,3743+0,7387*D3;2)',
        '="entre "&TEXT(MAX(0;0,3743+0,7387*D3-0,261);"0.00")&" e "&TEXT(0,3743+0,7387*D3+0,261;"0.00")'
    ]];
    var percent = [], share = [];
    for (var row = 3; row <= 62; row++) {
        percent.push(['=IF(OR(B' + row + '="";B' + row + '=0;D' + row + '="");"";D' + row + '/B' + row + '-1)']);
        share.push(['=IF(OR(D' + row + '="";$B' + row + '=0);"";D' + row + '/$B' + row + ')']);
    }
    // Uma única atualização atômica evita três recálculos da planilha grande.
    var apiWritten = false;
    if (typeof Sheets !== "undefined" && Sheets.Spreadsheets) {
        try {
            var sheetId = sheet.getSheetId();
            var blocks = [
                { row: 0, col: 0, values: top },
                { row: 2, col: 2, values: percent },
                { row: 2, col: 4, values: share }
            ];
            var requests = blocks.map(function (block) {
                return { updateCells: {
                    range: { sheetId: sheetId, startRowIndex: block.row,
                        endRowIndex: block.row + block.values.length,
                        startColumnIndex: block.col, endColumnIndex: block.col + block.values[0].length },
                    rows: block.values.map(function (row) {
                        return { values: row.map(function (formula) {
                            return { userEnteredValue: { formulaValue: formula } };
                        }) };
                    }),
                    fields: "userEnteredValue"
                } };
            });
            Sheets.Spreadsheets.batchUpdate({ requests: requests }, sheet.getParent().getId());
            apiWritten = true;
        } catch (errApi) {
            Logger.log("[Referências do dia] API indisponível, usando gravação nativa: " + errApi.message);
        }
    }
    if (!apiWritten) {
        batchSetFormulasSafeNoApiGE_(sheet.getRange(1, 1, 1, 2), top);
        batchSetFormulasSafeNoApiGE_(sheet.getRange(3, 3, 60, 1), percent);
        batchSetFormulasSafeNoApiGE_(sheet.getRange(3, 5, 60, 1), share);
    }
    return { top: "A1:B1", percentual: "C3:C62", share: "E3:E62", audiencia: "D3:D62",
        gravacao: apiWritten ? "api_atomica" : "nativa" };
}

// Entrada do editor: aplica somente as fórmulas, sem abrir outro dia ou
// reinicializar Hoje/OCR. Também permite corrigir o bloco já existente.
function atualizarReferenciasDiaAtualHistoricoGE() {
    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(30000)) throw new Error("Não foi possível obter o lock do Histórico.");
    try {
        var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
        var result = atualizarReferenciasDiaAtualHistoricoGE_(sheet);
        SpreadsheetApp.flush();
        Logger.log(JSON.stringify(result));
        return result;
    } finally {
        lock.releaseLock();
    }
}
