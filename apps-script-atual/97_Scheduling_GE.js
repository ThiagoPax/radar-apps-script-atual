// =============================================================================
// 97_Scheduling_GE.js — Agendamento de triggers e execução automática (GE)
// =============================================================================

// --- Helper genérico de gerenciamento de triggers ---

function manageTriggers_(e) {
    const t = ScriptApp.getProjectTriggers(),
        o = e.handlers || [];
    if ("remove" !== e.action) {
        if ("list" === e.action) return t.filter(e => o.includes(e.getHandlerFunction()));
        if ("create" === e.action) {
            const t = [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY];
            for (const o of t) ScriptApp.newTrigger(e.handler).timeBased().onWeekDay(o).atHour(e.hour).nearMinute(e.minute).create()
        }
    } else
        for (const e of t) o.includes(e.getHandlerFunction()) && ScriptApp.deleteTrigger(e)
}

// =============================================================================
// Agendamento — Gazeta Esportiva (GE)
// =============================================================================

function ativarAgendamentoGE() {
    const e = SpreadsheetApp.getUi();
    manageTriggers_({
        action: "remove",
        handlers: ["executarAnaliseAgendadaGE"]
    }), manageTriggers_({
        action: "create",
        handler: "executarAnaliseAgendadaGE",
        hour: CFG.GE_AUTO_RUN_HOUR,
        minute: CFG.GE_AUTO_RUN_MINUTE
    }), e.alert("\u2705 Agendamento Gazeta Esportiva ativado!", `O script ser\xe1 executado automaticamente:\n\n\u{1f4c5} Segunda a Sexta\n\u23f0 \xc0s ${CFG.GE_AUTO_RUN_HOUR}:${(CFG.GE_AUTO_RUN_MINUTE+"").padStart(2,"0")}\n\nO sistema ir\xe1:\n1. Buscar o programa do dia no YouTube\n2. Transcrever automaticamente\n3. Analisar e classificar\n4. Enviar relat\xf3rio por email`, e.ButtonSet.OK)
}

function desativarAgendamentoGE() {
    manageTriggers_({
        action: "remove",
        handlers: ["executarAnaliseAgendadaGE"]
    }), SpreadsheetApp.getUi().alert("\u274c Agendamento Gazeta Esportiva desativado")
}

function removerTriggersGE_() {
    manageTriggers_({
        action: "remove",
        handlers: ["executarAnaliseAgendadaGE"]
    })
}

function verStatusAgendamentoGE() {
    const e = SpreadsheetApp.getUi(),
        t = manageTriggers_({
            action: "list",
            handlers: ["executarAnaliseAgendadaGE"]
        });
    0 === t.length ? e.alert("\u{1f4cb} Status Gazeta Esportiva", "Nenhum agendamento ativo.", e.ButtonSet.OK) : e.alert("\u{1f4cb} Status Gazeta Esportiva", `Agendamentos ativos: ${t.length}\nHor\xe1rio: ${CFG.GE_AUTO_RUN_HOUR}:${(CFG.GE_AUTO_RUN_MINUTE+"").padStart(2,"0")}\nDias: Segunda a Sexta`, e.ButtonSet.OK)
}

// Trigger handler GE (lógica real — chamada pelo trigger)
function executarAnaliseAgendadaGE() {
    const e = getTranscriptionServerUrl_(),
        t = (new Date).getDay();
    if (0 !== t && 6 !== t) {
        if (!e) return Logger.log("Gazeta Esportiva: Servidor n\xe3o configurado."), void MailApp.sendEmail({
            to: CFG.EMAIL_DESTINO,
            subject: "\u26a0\ufe0f Erro Gazeta Esportiva - Servidor n\xe3o configurado",
            body: "O servidor de transcri\xe7\xe3o n\xe3o est\xe1 configurado. Configure em Configura\xe7\xf5es \u2192 Configurar Servidor."
        });
        try {
            const t = buscarGazetaEsportivaHoje_();
            if (!t) return Logger.log("Gazeta Esportiva: Programa n\xe3o encontrado automaticamente."), void MailApp.sendEmail({
                to: CFG.EMAIL_DESTINO,
                subject: "\u26a0\ufe0f Gazeta Esportiva - Programa n\xe3o encontrado",
                body: "N\xe3o foi poss\xedvel encontrar o Gazeta Esportiva de hoje automaticamente.\n\nImporte manualmente usando o menu."
            });
            const o = fetchJsonSafe_(e + "/transcribe", {
                method: "POST",
                contentType: "application/json",
                payload: JSON.stringify({
                    url: t
                }),
                muteHttpExceptions: !0
            });
            if (!o.success) throw Error(o.error || "Erro na transcri\xe7\xe3o");
            ensureGESheets_();
            const a = SpreadsheetApp.getActive().getSheetByName(CFG.GE_INPUT_SHEET);
            a.clear(), a.getRange(1, 1).setValue(t);
            const r = o.transcription.split("\n");
            for (let e = 0; e < r.length; e++) a.getRange(e + 2, 1).setValue(r[e]);
            const n = processAndMapTimestampsGE_(loadGEInputFromSheet_()),
                s = getStopwords_(),
                i = computeTopTermsPerMinute_(n.rows, s),
                c = buildTimelineGE_(n.rows, i, s);
            writeAllOutputsGE_(c, i, s, n.rows), enviarEmailRelatorioGE_(gerarRelatorioGE_(c)), Logger.log("Gazeta Esportiva processado com sucesso: " + o.videoTitle)
        } catch (e) {
            Logger.log("Erro no Gazeta Esportiva agendado: " + e.message), MailApp.sendEmail({
                to: CFG.EMAIL_DESTINO,
                subject: "\u26a0\ufe0f Erro Gazeta Esportiva - " + Utilities.formatDate(new Date, "America/Sao_Paulo", "dd/MM"),
                body: "Erro ao processar: " + e.message
            })
        }
    } else Logger.log("Gazeta Esportiva: Fim de semana, n\xe3o executa.")
}

// =============================================================================
// Agendamento — importação automática da transcrição (Speech)
// =============================================================================

function ativarAgendamentoSpeechRadar() {
    instalarTriggersRadar_(), SpreadsheetApp.getUi().alert("✅ Agendamento Speech Radar ativado", `Importação automática configurada para ${CFG.IMPORT_TRIGGER_HOUR}:${(CFG.IMPORT_TRIGGER_MINUTE+"").padStart(2,"0")} de segunda a sexta.`, SpreadsheetApp.getUi().ButtonSet.OK)
}

function desativarAgendamentoSpeechRadar() {
    manageTriggers_({
        action: "remove",
        handlers: ["processarSpeechDoDia_"]
    }), SpreadsheetApp.getUi().alert("❌ Agendamento Speech Radar desativado")
}