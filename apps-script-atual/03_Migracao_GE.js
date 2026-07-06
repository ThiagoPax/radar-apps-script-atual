// =============================================================================
// 03_Migracao_GE.js — Migração one-shot: JogoAberto_Radar → Radar_Gazeta_Esportiva
// =============================================================================
// Idempotente: pode rodar mais de uma vez sem estragar nada.
// Ordem: renomear arquivo → arquivar abas JA → renomear abas GE → criar abas
// novas → preparar Hoje → reconstruir Base/Resumo → corrigir formatação →
// reordenar/ocultar abas.

// Abas do Jogo Aberto: a versão sem prefixo só é considerada JA se a versão
// "GE - ..." também existir (protege contra rodar duas vezes).
var MIGRACAO_GE = {
    ABAS_JA_COM_PAR: {
        "Top 20 Termos": "GE - Top 20 Termos",
        "Macrotema": "GE - Macrotema",
        "Subtema": "GE - Subtema",
        "Minuto a Minuto": "GE - Minuto a Minuto",
        "Histórico": "GE - Histórico"
    },
    ABAS_ARQUIVAR_SEMPRE: ["Radar OCR"], // sem referências no código
    RENOMEACOES: {
        "GE - Histórico": "Histórico",
        "GE - Top 20 Termos": "Termos",
        "GE - Macrotema": "Macrotema",
        "GE - Subtema": "Subtema",
        "GE - Minuto a Minuto": "Minuto",
        "GE - Média por Minuto TEMA": "Tema Minuto",
        "GE - Média por Dia da Semana": "Dia Semana",
        "GE - Média por Comentarista": "Comentaristas",
        "GE - Predição Hoje": "zz_Predicao_Antiga",
        "GE - Audiência Semanal": "Semana",
        "GE - Audiência Mensal": "Mês",
        "GE - Dashboard Email": "Email",
        "GE - Relatório PDF": "PDF",
        "GE - Backtest Previsão": "Backtest",
        "Breaks Globo": "Breaks",
        "Calendário Clubes": "Clubes"
    }
};

function migrarParaRadarGazeta() {
    const ss = SpreadsheetApp.getActive();
    const resumo = [];

    // 1. Renomear o arquivo (mesmo ID)
    try {
        const file = DriveApp.getFileById(ss.getId());
        if (file.getName() !== CFG.ARQUIVO_NOME) {
            file.setName(CFG.ARQUIVO_NOME);
            resumo.push("Arquivo renomeado para " + CFG.ARQUIVO_NOME);
        }
    } catch (e) {
        resumo.push("⚠️ Falha ao renomear arquivo: " + e.message);
    }

    // 2. Exportar e remover abas do Jogo Aberto
    try {
        const movidas = arquivarAbasJogoAbertoGE_();
        if (movidas.length) resumo.push("Abas JA arquivadas e removidas: " + movidas.join(", "));
        else resumo.push("Nenhuma aba JA pendente (já migrado)");
    } catch (e) {
        resumo.push("⚠️ Falha ao arquivar abas JA: " + e.message);
    }

    // 3. Renomear abas GE para nomes curtos
    for (const antigo in MIGRACAO_GE.RENOMEACOES) {
        const novo = MIGRACAO_GE.RENOMEACOES[antigo];
        const sh = ss.getSheetByName(antigo);
        if (sh && !ss.getSheetByName(novo)) {
            sh.setName(novo);
            resumo.push("Aba renomeada: " + antigo + " → " + novo);
        }
    }

    // 4. Criar abas novas necessárias
    getConfigSheetGE_();
    ensureAbaGE_(CFG.SHEET_LOGS, ["QUANDO", "EVENTO", "DETALHE"]);
    getBaseSheetGE_();
    ensureAbaGE_(CFG.SHEET_RESUMO, RESUMO_GE.HEADERS);
    ensureAbaGE_(CFG.SHEET_SEMANA);
    ensureAbaGE_(CFG.SHEET_MES);
    ensureAbaGE_(CFG.SHEET_BREAKS, ["DATA", "INÍCIO", "FIM", "DURAÇÃO (min)", "TIPO"]);
    ensureAbaGE_(CFG.SHEET_CLUBES);
    ensureAbaGE_(CFG.SHEET_EMAIL);
    ensureAbaGE_(CFG.SHEET_PDF);
    resumo.push("Abas novas garantidas (Hoje, Resumo, Config, Logs, Base...)");

    // 5. Preparar o cockpit Hoje
    try {
        prepararHojeGE();
        resumo.push("Aba Hoje preparada");
    } catch (e) {
        resumo.push("⚠️ Falha ao preparar Hoje: " + e.message);
    }

    // 6. Base normalizada + visões executivas
    try {
        reconstruirBaseGE();
        atualizarResumoGE();
        resumo.push("Base e Resumo reconstruídos");
    } catch (e) {
        resumo.push("⚠️ Falha na Base/Resumo: " + e.message);
    }

    // 7. Formatação condicional do Histórico
    try {
        const n = corrigirFormatacaoHistoricoGE();
        resumo.push("Formatação &/MIN aplicada em " + n + " colunas");
    } catch (e) {
        resumo.push("⚠️ Falha na formatação: " + e.message);
    }

    // 8. Ordem e visibilidade das abas
    try {
        reordenarAbasGE_();
        resumo.push("Abas reordenadas e técnicas ocultadas");
    } catch (e) {
        resumo.push("⚠️ Falha ao reordenar: " + e.message);
    }

    logGE_("MIGRACAO", resumo.join(" | "));
    try {
        SpreadsheetApp.getUi().alert("Migração concluída:\n\n• " + resumo.join("\n• "));
    } catch (e) { /* sem UI (trigger) */ }
    return resumo;
}

// Copia as abas JA para uma planilha-arquivo no Drive e as remove da principal
function arquivarAbasJogoAbertoGE_() {
    const ss = SpreadsheetApp.getActive();
    const movidas = [];
    const candidatas = [];

    for (const semPrefixo in MIGRACAO_GE.ABAS_JA_COM_PAR) {
        const par = MIGRACAO_GE.ABAS_JA_COM_PAR[semPrefixo];
        if (ss.getSheetByName(semPrefixo) && ss.getSheetByName(par)) {
            candidatas.push(semPrefixo);
        }
    }
    for (const nome of MIGRACAO_GE.ABAS_ARQUIVAR_SEMPRE) {
        if (ss.getSheetByName(nome)) candidatas.push(nome);
    }
    if (!candidatas.length) return movidas;

    const nomeArquivo = "Radar_Arquivo_JogoAberto_" +
        Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyy-MM-dd");
    // Reutiliza o arquivo-morto se já existir (execução anterior interrompida)
    let destino = null;
    const existentes = DriveApp.getFilesByName(nomeArquivo);
    destino = existentes.hasNext()
        ? SpreadsheetApp.open(existentes.next())
        : SpreadsheetApp.create(nomeArquivo);

    for (const nome of candidatas) {
        const sh = ss.getSheetByName(nome);
        try {
            // Cópia parcial de execução anterior: aproveita e só remove da principal
            if (!destino.getSheetByName(nome)) {
                const copiaMeio = destino.getSheetByName("Cópia de " + nome);
                if (copiaMeio) copiaMeio.setName(nome);
                else sh.copyTo(destino).setName(nome);
            }
            ss.deleteSheet(sh);
            movidas.push(nome);
        } catch (e) {
            logGE_("MIGRACAO", "Falha ao arquivar aba " + nome + ": " + e.message);
        }
    }
    // Remove a "Página1" vazia criada por padrão
    try {
        const padrao = destino.getSheets()[0];
        if (destino.getSheets().length > 1 && padrao.getLastRow() === 0) {
            destino.deleteSheet(padrao);
        }
    } catch (e) { /* cosmético */ }

    if (movidas.length) {
        logGE_("MIGRACAO", "Arquivo criado: " + destino.getUrl());
    }
    return movidas;
}

// Reordena conforme CFG.ORDEM_ABAS; técnicas e zz_* ficam no fim, ocultas
function reordenarAbasGE_() {
    const ss = SpreadsheetApp.getActive();
    let pos = 1;
    for (const nome of CFG.ORDEM_ABAS) {
        const sh = ss.getSheetByName(nome);
        if (!sh) continue;
        ss.setActiveSheet(sh);
        ss.moveActiveSheet(pos++);
    }
    for (const sh of ss.getSheets()) {
        const nome = sh.getName();
        const tecnica = CFG.ABAS_OCULTAS.indexOf(nome) >= 0 || nome.indexOf("zz_") === 0;
        try {
            if (tecnica && !sh.isSheetHidden()) sh.hideSheet();
        } catch (e) { /* não oculta a única aba visível */ }
    }
    const hoje = ss.getSheetByName(CFG.SHEET_HOJE);
    if (hoje) ss.setActiveSheet(hoje);
}

// -----------------------------------------------------------------------------
// Restauração da aba Breaks a partir do backup pré-migração
// -----------------------------------------------------------------------------
// A primeira versão das estatísticas escreveu sobre a estrutura importada da
// aba "Breaks Globo". Esta função recupera a aba original do backup e a
// recoloca como "Breaks" (as estatísticas novas vivem só da coluna R em diante).
function restaurarAbaBreaksDoBackupGE() {
    const ss = SpreadsheetApp.getActive();
    const backups = DriveApp.getFilesByName("Radar_Backup_pre_migracao_2026-07-04");
    if (!backups.hasNext()) throw Error("Backup pré-migração não encontrado no Drive.");
    const backup = SpreadsheetApp.open(backups.next());
    const origem = backup.getSheetByName("Breaks Globo") || backup.getSheetByName("Breaks");
    if (!origem) throw Error('Aba "Breaks Globo" não existe no backup.');

    const atual = ss.getSheetByName(CFG.SHEET_BREAKS);
    if (atual) atual.setName("zz_Breaks_danificada");

    const restaurada = origem.copyTo(ss);
    restaurada.setName(CFG.SHEET_BREAKS);
    if (restaurada.isSheetHidden()) restaurada.showSheet();

    // Recoloca na posição da ordem oficial
    const pos = CFG.ORDEM_ABAS.indexOf(CFG.SHEET_BREAKS);
    if (pos >= 0) {
        ss.setActiveSheet(restaurada);
        ss.moveActiveSheet(Math.min(pos + 1, ss.getSheets().length));
    }

    // Remove a versão danificada e reaplica as estatísticas no bloco R+
    const danificada = ss.getSheetByName("zz_Breaks_danificada");
    if (danificada) ss.deleteSheet(danificada);
    try { atualizarBreaksGE(); } catch (e) { logGE_("BREAKS", "Estatísticas: " + e.message); }

    logGE_("BREAKS", "Aba Breaks restaurada do backup (estrutura de importação preservada)");
    return true;
}
