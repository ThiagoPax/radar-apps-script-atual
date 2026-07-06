// =============================================================================
// 99_Menu_Triggers_GE.js — Menu único, roteador de edição e automações
// =============================================================================
// Rotina diária (seg–sex):
//   06:00  rotinaManhaGE          calendário, breaks, médias, novo dia, Hoje
//   17:58+ geTickAoVivoGE (1/min) sincroniza Hoje + motor durante o programa
//   19:00  geImportacaoGE (1ª)    tentativa de localizar transcrição no Drive
//   19:05  geImportacaoGE (2ª)    idem (processarSpeechDoDia_ é idempotente)
//   19:10  geImportacaoGE (3ª)    tentativa principal
//   19:15  geAlertaTranscricaoGE  alerta se ainda não houver transcrição
//   21:00  geEmailDashboardGE     email dashboard do dia

// -----------------------------------------------------------------------------
// Menu
// -----------------------------------------------------------------------------
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("📺 Gazeta Esportiva Radar")
        .addItem("▶️ Rodar tudo agora", "rodarTudoAgoraGE")
        .addSeparator()
        .addItem("🎙️ Importar transcrição do dia", "importarTranscricaoDoDiaMenuGE")
        .addItem("🧠 Processar conteúdo do dia", "processarConteudoDoDiaGE")
        .addItem("📡 Atualizar audiência/predição", "atualizarAudienciaPredicaoMenuGE")
        .addSeparator()
        .addItem("📧 Gerar email dashboard", "gerarEmailDashboardMenuGE")
        .addItem("📄 Gerar PDF audiência minuto a minuto", "gerarPDFAudienciaMenuGE")
        .addSeparator()
        .addItem("⚽ Atualizar calendário dos clubes", "atualizarClubesMenuGE")
        .addItem("📺 Atualizar breaks Globo", "atualizarBreaksMenuGE")
        .addItem("📊 Atualizar médias semanais/mensais", "atualizarMediasMenuGE")
        .addItem("🎨 Corrigir formatação do Histórico", "corrigirFormatacaoHistoricoMenuGE")
        .addItem("🔧 Reparar fórmulas quebradas (#NAME?)", "repararFormulasQuebradasMenuGE")
        .addSeparator()
        .addItem("🤖 Status da automação", "statusAutomacaoGE")
        .addSubMenu(ui.createMenu("⚙️ Configurações")
            .addItem("🌍 Definir contexto ativo", "definirContextoAtivoGE")
            .addItem("🧱 Preparar aba Hoje (novo dia)", "prepararDiaMenuGE")
            .addItem("🔑 Configurar Webhook/OCR", "configurarCapturaAudienciaGE")
            .addItem("🧪 Testar webhook OCR", "testarWebhookAudienciaGE")
            .addItem("⚽ Testar API-Football", "testarApiFootballGE")
            .addItem("🖥️ Configurar servidor de transcrição", "configurarServidor")
            .addItem("⏰ Instalar/renovar gatilhos", "instalarTriggersMenuGE")
            .addItem("📈 Rodar backtest da predição", "rodarBacktestMenuGE")
            .addItem("🎯 Backtest de checkpoints (V3)", "rodarBacktestCheckpointsMenuGE")
            .addItem("🔧 Recalcular fatores (clube etc.)", "recalcularFatoresMenuGE")
            .addItem("🔀 Promover/reverter previsão V3", "promoverOuReverterV3MenuGE")
            .addItem("🧹 Limpar cache LLM", "clearLLMCache")
            .addItem("🧹 Limpar stopwords", "clearUserStopwords")
            .addItem("🩺 Verificar sistema (clubes, médias, hoje)", "verificarSistemaMenuGE")
            .addItem("🩹 Restaurar aba Breaks do backup", "restaurarAbaBreaksDoBackupGE")
            .addItem("🚚 Executar migração (one-shot)", "migrarParaRadarGazeta")
            .addItem("💾 Exportar backup do script", "exportarScriptsParaGoogleDrive"))
        .addToUi();
}

// -----------------------------------------------------------------------------
// Esqueleto de análise (ex-15_Menu) — usado pelo pipeline Speech e pelo menu
// -----------------------------------------------------------------------------
function runAnalysisSkeleton_(e) {
    const t = e.loadInput();
    if (0 === t.rows.length) return e.onEmptyInput(), null;
    const o = e.processInput(t);
    if (!1 !== e.checkProcessed && 0 === o.rows.length) return e.onEmptyProcessed(), null;
    e.ensureSheets();
    const a = getStopwords_(),
        r = computeTopTermsPerMinute_(o.rows, a);
    return {
        rawInput: t,
        input: o,
        stop: a,
        minuteTopTerms: r,
        timeline: e.buildTimeline ? e.buildTimeline(o.rows, r, a) : null
    };
}

function ensureSheetHasData_(nome, onEmpty) {
    const sh = SpreadsheetApp.getActive().getSheetByName(nome);
    return !sh || sh.getLastRow() < 2 ? (onEmpty(), null) : sh;
}

// Processa a transcrição já presente na aba Input (temas, termos, minuto a
// minuto, arquivamento e email de resumo)
function processarConteudoDoDiaGE() {
    const ui = SpreadsheetApp.getUi();
    const ctx = runAnalysisSkeleton_({
        loadInput: loadGEInputFromSheet_,
        onEmptyInput: () => ui.alert('Aba "' + CFG.GE_INPUT_SHEET + '" está vazia — importe a transcrição primeiro.'),
        processInput: processAndMapTimestampsGE_,
        onEmptyProcessed: () => ui.alert("Não foi possível processar a transcrição do Gazeta Esportiva."),
        ensureSheets: ensureGESheets_,
        buildTimeline: buildTimelineGE_
    });
    if (!ctx) return;
    writeAllOutputsGE_(ctx.timeline, ctx.minuteTopTerms, ctx.stop, ctx.input.rows);

    let msgArquivo = "";
    try {
        const data = extrairDataProgramaGE_() || new Date();
        msgArquivo = arquivarProgramaGE_(data)
            ? "\n📁 Arquivado no Histórico"
            : "\n📁 Histórico: programa já arquivado (temas sincronizados)";
        try { sincronizarTermosMinutoAMinutoParaHistoricoGE_(data); } catch (e2) { /* opcional */ }
    } catch (e) {
        msgArquivo = "\n⚠️ Erro ao arquivar: " + e.message;
    }
    ui.alert("✅ Conteúdo processado: " + ctx.timeline.length + " minutos." + msgArquivo);
}

function importarTranscricaoDoDiaMenuGE() {
    const res = processarSpeechDoDia_();
    const ui = SpreadsheetApp.getUi();
    if (res && res.success) {
        ui.alert(res.skipped
            ? "ℹ️ Transcrição de hoje já havia sido processada."
            : "✅ Transcrição importada e processada.\nArquivo: " + (res.sourceName || "—"));
    } else {
        ui.alert("⚠️ Transcrição não encontrada/importada. Detalhe: " +
            JSON.stringify(res || {}).substring(0, 300) +
            "\n\nVerifique a pasta do Drive ou importe por URL.");
    }
}

function atualizarAudienciaPredicaoMenuGE() {
    atualizarHojeGE();
    SpreadsheetApp.getUi().alert("✅ Audiência sincronizada e predição recalculada (veja a aba Hoje).");
}

function atualizarBreaksMenuGE() {
    atualizarBreaksGE();
    SpreadsheetApp.getUi().alert("✅ Estatísticas de breaks atualizadas (aba Breaks).");
}

function atualizarMediasMenuGE() {
    reconstruirBaseGE();
    atualizarResumoGE();
    atualizarSemanaGE();
    atualizarMesGE();
    atualizarDiaSemanaGE();
    try { atualizarMediaPorComentaristaGE_(); } catch (e) { logGE_("COMENT", e.message); }
    SpreadsheetApp.getUi().alert("✅ Base, Resumo, Semana, Mês, Dia Semana e Comentaristas atualizados.");
}

function prepararDiaMenuGE() {
    const res = adicionarNovoDiaCoreGE_(new Date());
    prepararHojeGE();
    SpreadsheetApp.getUi().alert(res.ok
        ? "✅ Novo dia criado no Histórico e aba Hoje preparada."
        : "ℹ️ Bloco do dia já existia — aba Hoje re-preparada.");
}

function definirContextoAtivoGE() {
    const ui = SpreadsheetApp.getUi();
    const atual = String(getConfigValueGE_("CONTEXTO_ATIVO", "NORMAL"));
    const resp = ui.prompt("🌍 Contexto ativo",
        "Contexto atual: " + atual + "\n\nOpções: " + CFG.CONTEXTOS_VALIDOS.join(" | ") +
        "\n\nDigite o novo contexto:", ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    const novo = resp.getResponseText().trim().toUpperCase();
    if (!novo) return;
    setConfigValueGE_("CONTEXTO_ATIVO", novo);
    setConfigValueGE_("CONTEXTO_DESDE", Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy"));
    prepararHojeGE();
    ui.alert("✅ Contexto ativo: " + novo + " (réguas da aba Hoje recalculadas).");
}

// "Rodar tudo agora" — fluxo completo pós-programa
function rodarTudoAgoraGE() {
    const passos = [];
    try { const r = processarSpeechDoDia_(); passos.push("Transcrição: " + (r && r.success ? (r.skipped ? "já processada" : "importada e processada") : "não encontrada")); }
    catch (e) { passos.push("Transcrição: ⚠️ " + e.message); }
    try { atualizarHojeGE(); passos.push("Audiência/predição: ok"); }
    catch (e) { passos.push("Audiência/predição: ⚠️ " + e.message); }
    try { reconstruirBaseGE(); atualizarResumoGE(); atualizarSemanaGE(); atualizarMesGE(); passos.push("Médias/Resumo: ok"); }
    catch (e) { passos.push("Médias: ⚠️ " + e.message); }
    try { atualizarBreaksGE(); passos.push("Breaks: ok"); }
    catch (e) { passos.push("Breaks: ⚠️ " + e.message); }
    try { corrigirFormatacaoHistoricoGE(); passos.push("Formatação: ok"); }
    catch (e) { passos.push("Formatação: ⚠️ " + e.message); }
    logGE_("RODAR_TUDO", passos.join(" | "));
    try { SpreadsheetApp.getUi().alert("Rodar tudo agora:\n\n• " + passos.join("\n• ")); } catch (e) { }
}

// -----------------------------------------------------------------------------
// Roteador de edição (instalável — instalado por instalarTriggersGE)
// -----------------------------------------------------------------------------
function geOnEditInstalavelGE(e) {
    if (!e || !e.range) return;
    try {
        const sheet = e.range.getSheet();
        const nome = sheet.getName();
        const row = e.range.getRow();
        const col = e.range.getColumn();

        if (nome === CFG.SHEET_HOJE) {
            // OCR (D), manual (E), tema (I) ou subtema (J) nos minutos
            if (row >= 3 && row <= 62 && [4, 5, 9, 10].indexOf(col) >= 0) {
                if (throttleGE_("hoje_edit", 4)) atualizarHojeGE();
            }
            return;
        }
        if (nome === CFG.GE_OUT_HISTORICO) {
            if (row === 64 || row === 65) {
                try { atualizarMediaPorComentaristaGE_(); } catch (err) { }
            }
            if (col === 4 && row >= 3 && row <= 62 &&
                typeof reconstruirCadeiaProjecaoGE_ === "function") {
                try { reconstruirCadeiaProjecaoGE_(sheet, row); } catch (err) { }
            }
            if ((col === 4 || col === 6) && row >= 3 && row <= 62) {
                if (throttleGE_("hist_edit", 4) && typeof ge32_atualizarTYPainelAgora_ === "function") {
                    try { ge32_atualizarTYPainelAgora_("onEdit_router"); } catch (err) { }
                }
            }
            try { if (typeof onEditAlertaTemaGE === "function") onEditAlertaTemaGE(e); } catch (err) { }
        }
    } catch (err) {
        // nunca travar a digitação do usuário
    }
}

// Throttle simples via CacheService (segundos)
function throttleGE_(chave, segundos) {
    try {
        const cache = CacheService.getScriptCache();
        if (cache.get("thr_" + chave)) return false;
        cache.put("thr_" + chave, "1", segundos);
        return true;
    } catch (e) {
        return true;
    }
}

// -----------------------------------------------------------------------------
// Handlers das automações
// -----------------------------------------------------------------------------
function rotinaManhaGE() {
    const dia = new Date().getDay();
    if (dia === 0 || dia === 6) return; // fim de semana
    const passos = [];
    try { atualizarClubesAutoGE_(); passos.push("clubes+contexto"); } catch (e) { passos.push("clubes⚠️" + e.message); }
    try { atualizarBreaksGE(); passos.push("breaks"); } catch (e) { passos.push("breaks⚠️" + e.message); }
    try { reconstruirBaseGE(); atualizarResumoGE(); atualizarSemanaGE(); atualizarMesGE(); passos.push("médias"); } catch (e) { passos.push("médias⚠️" + e.message); }
    try { atualizarDiaSemanaGE(); passos.push("dia semana"); } catch (e) { passos.push("dia semana⚠️" + e.message); }
    try { atualizarMediaPorComentaristaGE_(); passos.push("comentaristas"); } catch (e) { passos.push("comentaristas⚠️" + e.message); }
    try { recalcularFatoresGE(); passos.push("fatores"); } catch (e) { passos.push("fatores⚠️" + e.message); }
    try {
        const r = adicionarNovoDiaCoreGE_(new Date());
        passos.push(r.ok ? "novo dia criado" : "dia já existia");
    } catch (e) { passos.push("novo dia⚠️" + e.message); }
    try { prepararHojeGE(); passos.push("Hoje preparada"); } catch (e) { passos.push("Hoje⚠️" + e.message); }
    logGE_("ROTINA_06H", passos.join(" | "));
}

// Tick ao vivo — roda a cada minuto, só age na janela 17:58–19:05
function geTickAoVivoGE() {
    const agora = new Date();
    const h = Number(Utilities.formatDate(agora, CFG.TIMEZONE, "H"));
    const m = Number(Utilities.formatDate(agora, CFG.TIMEZONE, "m"));
    const minutos = h * 60 + m;
    if (minutos < 17 * 60 + 58 || minutos > 19 * 60 + 5) return;
    const dia = agora.getDay();
    if (dia === 0 || dia === 6) return;
    try { atualizarHojeGE(); } catch (e) { logGE_("TICK", "⚠️ " + e.message); }
}

// Tentativas de importação da transcrição (19:00 / 19:05 / 19:10)
function geImportacaoGE() {
    const dia = new Date().getDay();
    if (dia === 0 || dia === 6) return;
    try {
        const r = processarSpeechDoDia_();
        logGE_("IMPORTACAO", JSON.stringify({
            success: r && r.success, skipped: r && r.skipped,
            arquivo: r && r.sourceName, modo: r && r.sourceMode
        }).substring(0, 400));
    } catch (e) {
        logGE_("IMPORTACAO", "⚠️ " + e.message);
    }
}

// 19:15 — alerta se a transcrição do dia ainda não foi processada
function geAlertaTranscricaoGE() {
    const dia = new Date().getDay();
    if (dia === 0 || dia === 6) return;
    if (transcricaoDoDiaProcessadaGE_()) return;
    logGE_("IMPORTACAO", "❌ Transcrição de hoje não localizada até 19:15");
    if (CFG.EMAIL_ENABLED) {
        MailApp.sendEmail({
            to: CFG.EMAIL_DESTINO,
            subject: "⚠️ Gazeta Esportiva — transcrição não encontrada (" +
                Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM") + ")",
            body: "A transcrição do programa de hoje não foi localizada no Drive até as 19:15.\n\n" +
                "Verifique a pasta de transcrições (arquivo transcricao_DDMMYY.xlsx) ou importe " +
                "manualmente pelo menu 📺 Gazeta Esportiva Radar → Importar transcrição do dia."
        });
    }
}

function transcricaoDoDiaProcessadaGE_() {
    try {
        const props = PropertiesService.getScriptProperties();
        const keys = props.getKeys();
        const hoje = new Date();
        const marcas = [
            Utilities.formatDate(hoje, CFG.TIMEZONE, "yyyy-MM-dd"),
            Utilities.formatDate(hoje, CFG.TIMEZONE, "dd/MM/yyyy"),
            Utilities.formatDate(hoje, CFG.TIMEZONE, "ddMMyy")
        ];
        for (const k of keys) {
            if (k.indexOf(CFG.SPEECH_PROCESSED_PREFIX) !== 0) continue;
            for (const m of marcas) if (k.indexOf(m) >= 0) return true;
        }
    } catch (e) { /* na dúvida, alerta */ }
    return false;
}

// 21:00 — email dashboard
function geEmailDashboardTriggerGE() {
    const dia = new Date().getDay();
    if (dia === 0 || dia === 6) return;
    try { gerarEmailDashboardGE(); }
    catch (e) { logGE_("EMAIL", "⚠️ Falha no dashboard 21h: " + e.message); }
    // Placar honesto do dia fechado: V2 vs V3 (shadow mode)
    try { if (typeof registrarShadowLogV3GE_ === "function") registrarShadowLogV3GE_(); }
    catch (e) { logGE_("SHADOW", "⚠️ " + e.message); }
}

// -----------------------------------------------------------------------------
// Instalação de triggers — remove TODOS os antigos (inclusive do Jogo Aberto)
// -----------------------------------------------------------------------------
function instalarTriggersGE() {
    for (const t of ScriptApp.getProjectTriggers()) ScriptApp.deleteTrigger(t);

    const ss = SpreadsheetApp.getActive();
    ScriptApp.newTrigger("geOnEditInstalavelGE").forSpreadsheet(ss).onEdit().create();
    ScriptApp.newTrigger("geTickAoVivoGE").timeBased().everyMinutes(1).create();

    const diarios = [
        ["rotinaManhaGE", CFG.ROTINA_MANHA_HOUR, 0],
        ["geImportacaoGE", CFG.IMPORT_TENTATIVAS[0].hour, CFG.IMPORT_TENTATIVAS[0].minute],
        ["geImportacaoGE", CFG.IMPORT_TENTATIVAS[1].hour, CFG.IMPORT_TENTATIVAS[1].minute],
        ["geImportacaoGE", CFG.IMPORT_TENTATIVAS[2].hour, CFG.IMPORT_TENTATIVAS[2].minute],
        ["geAlertaTranscricaoGE", CFG.IMPORT_ALERTA_HOUR, CFG.IMPORT_ALERTA_MINUTE],
        ["geEmailDashboardTriggerGE", CFG.EMAIL_DASHBOARD_HOUR, CFG.EMAIL_DASHBOARD_MINUTE]
    ];
    // Um trigger diário por rotina (limite de 20 triggers por script);
    // os handlers ignoram sábado/domingo internamente.
    for (const [handler, hora, minuto] of diarios) {
        ScriptApp.newTrigger(handler).timeBased().everyDays(1)
            .atHour(hora).nearMinute(minuto).create();
    }
    logGE_("TRIGGERS", "Gatilhos reinstalados: onEdit, tick 1/min e " + diarios.length + " rotinas diárias");
}

function instalarTriggersMenuGE() {
    instalarTriggersGE();
    SpreadsheetApp.getUi().alert("✅ Gatilhos reinstalados:\n\n" +
        "• onEdit (roteador Hoje/Histórico)\n" +
        "• Tick ao vivo a cada minuto (17:58–19:05)\n" +
        "• 06:00 rotina da manhã\n" +
        "• 19:00 / 19:05 / 19:10 importação da transcrição\n" +
        "• 19:15 alerta de transcrição ausente\n" +
        "• 21:00 email dashboard");
}

// -----------------------------------------------------------------------------
// Status da automação
// -----------------------------------------------------------------------------
function statusAutomacaoGE() {
    const triggers = ScriptApp.getProjectTriggers();
    const linhas = triggers.map(t => "• " + t.getHandlerFunction() + " (" + t.getEventType() + ")");
    const props = PropertiesService.getScriptProperties();
    const checagens = [
        "OPENAI_API_KEY: " + (props.getProperty("OPENAI_API_KEY") ? "✅" : "❌ ausente"),
        "API_FOOTBALL_KEY: " + (props.getProperty("API_FOOTBALL_KEY") ? "✅" : "❌ ausente"),
        "Webhook OCR (secret): " + (props.getProperty("AUD_GE_SECRET") || props.getKeys().some(k => k.indexOf("AUD") >= 0) ? "✅" : "⚠️ verificar"),
        "Contexto ativo: " + getConfigValueGE_("CONTEXTO_ATIVO", "NORMAL"),
        "Transcrição de hoje: " + (transcricaoDoDiaProcessadaGE_() ? "✅ processada" : "— ainda não")
    ];
    SpreadsheetApp.getUi().alert("🤖 Status da automação\n\n" +
        "Gatilhos (" + triggers.length + "):\n" + (linhas.join("\n") || "nenhum — use Instalar/renovar gatilhos") +
        "\n\n" + checagens.join("\n"));
}

// -----------------------------------------------------------------------------
// Verificação pós-migração (executável do editor, sem UI) — resultados nos Logs
// -----------------------------------------------------------------------------
function verificarSistemaGE() {
    const passos = [];
    try { passos.push("clubes: " + (atualizarClubesAutoGE_() ? "OK" : "SEM FONTE (ver Logs)")); }
    catch (e) { passos.push("clubes⚠️" + e.message); }
    try { atualizarDiaSemanaGE(); passos.push("dia semana: OK"); }
    catch (e) { passos.push("dia semana⚠️" + e.message); }
    try { atualizarMediaPorComentaristaGE_(); passos.push("comentaristas: OK"); }
    catch (e) { passos.push("comentaristas⚠️" + e.message); }
    try { const c = recalcularFatorClubeGE_(); passos.push("fatores clube: " + c + " células"); }
    catch (e) { passos.push("fatores⚠️" + e.message); }
    try { const v3 = recalcularPrevisaoV3ShadowGE_(); passos.push("V3 shadow: média " + (v3.mediaFinal ? v3.mediaFinal.toFixed(2) : "—")); }
    catch (e) { passos.push("V3⚠️" + e.message); }
    try { atualizarBreaksGE(); passos.push("breaks: OK"); }
    catch (e) { passos.push("breaks⚠️" + e.message); }
    try { atualizarSemanaGE(); atualizarMesGE(); passos.push("semana/mês: OK"); }
    catch (e) { passos.push("semana/mês⚠️" + e.message); }
    try { atualizarHojeGE(); passos.push("hoje/predição: OK"); }
    catch (e) { passos.push("hoje⚠️" + e.message); }
    // Diagnóstico da API-Football sem UI
    try {
        const apiKey = PropertiesService.getScriptProperties().getProperty("API_FOOTBALL_KEY");
        if (!apiKey) passos.push("api-football: ❌ chave ausente");
        else {
            const resp = UrlFetchApp.fetch(CFG.API_FOOTBALL_ENDPOINT + "/status", {
                method: "get", headers: { "x-apisports-key": apiKey }, muteHttpExceptions: true });
            const json = JSON.parse(resp.getContentText());
            if (resp.getResponseCode() === 200 && json.response && json.response.account) {
                const req = json.response.requests || {};
                const sub = json.response.subscription || {};
                passos.push("api-football: ✅ plano " + (sub.plan || "?") +
                    (sub.active === false ? " INATIVO" : "") + ", " +
                    (req.current || 0) + "/" + (req.limit_day || "?") + " req hoje");
            } else {
                passos.push("api-football: ⚠️ HTTP " + resp.getResponseCode() + " " +
                    JSON.stringify(json.errors || json.message || "").substring(0, 150));
            }
        }
    } catch (e) { passos.push("api-football: ❌ " + e.message); }
    logGE_("VERIFICACAO", passos.join(" | "));
    return passos;
}

function verificarSistemaMenuGE() {
    const passos = verificarSistemaGE();
    SpreadsheetApp.getUi().alert("🩺 Verificação do sistema:\n\n• " + passos.join("\n• "));
}

function atualizarClubesMenuGE() {
    const ok = atualizarClubesAutoGE_();
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_CLUBES);
    const n = sh ? Math.max(sh.getLastRow() - 1, 0) : 0;
    SpreadsheetApp.getUi().alert(ok
        ? "⚽ Calendário atualizado: " + n + " jogos na aba Clubes.\nContexto Jogo regenerado."
        : "⚠️ Nenhuma fonte retornou jogos — veja a aba Logs.");
}
