// =============================================================================
// 16_Speech_Integration.js — Integração Speech Timer -> Radar GE
// =============================================================================

function processarSpeechDoDia_() {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const sourceMode = obterSpeechSourceMode_();
        Logger.log("[Speech] Iniciando processamento. Modo: " + sourceMode);
        const sourceData = carregarSpeechDoDia_({
            sourceMode: sourceMode
        });
        if (!sourceData) {
            Logger.log("[Speech] Nenhuma transcrição do dia encontrada.");
            maybeEnviarAlertaSpeechAusente_();
            return {
                success: !1,
                skipped: !0,
                reason: "not_found"
            };
        }
        const signature = buildSpeechSignature_(sourceData);
        if (jaFoiProcessadoSpeechHoje_(signature)) return Logger.log("[Speech] Processamento duplicado evitado: " + signature), {
            success: !0,
            skipped: !0,
            reason: "already_processed"
        };
        if (!sourceData.programDate) throw Error("Não foi possível determinar a data do programa.");
        preencherInputComSpeechGE_(sourceData);
        const pipelineResult = executarPipelineCompletoGazeta_({
            skipArchive: !0,
            skipUi: !0
        });
        const syncResult = sincronizarTermosMinutoAMinutoParaHistoricoGE_(sourceData.programDate);
        marcarSpeechComoProcessado_(signature, sourceData);
        Logger.log("[Speech] Processamento concluído. Histórico: " + syncResult.status);
        return {
            success: !0,
            skipped: !1,
            programDate: formatDateKey_(sourceData.programDate),
            sourceMode: sourceMode,
            sourceName: sourceData.fileName || sourceData.youtubeTitle || "",
            pipeline: pipelineResult,
            historico: syncResult
        };
    } catch (e) {
        Logger.log("[Speech] Erro no processamento: " + (e.stack || e.message));
        throw e;
    } finally {
        lock.releaseLock();
    }
}

function testeImportacaoSpeechHoje_() {
    return processarSpeechDoDia_();
}

function instalarTriggersRadar_() {
    const desiredSlots = [{
        hour: CFG.IMPORT_TRIGGER_HOUR,
        minute: CFG.IMPORT_TRIGGER_MINUTE
    }, {
        hour: CFG.IMPORT_RETRY_TRIGGER_HOUR,
        minute: CFG.IMPORT_RETRY_TRIGGER_MINUTE
    }];
    const weekdays = [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY];
    manageTriggers_({
        action: "remove",
        handlers: ["processarSpeechDoDia_"]
    });
    for (const slot of desiredSlots) {
        for (const day of weekdays) ScriptApp.newTrigger("processarSpeechDoDia_").timeBased().onWeekDay(day).atHour(slot.hour).nearMinute(slot.minute).create();
        Logger.log("[Speech] Trigger criado para " + slot.hour + ":" + (slot.minute + "").padStart(2, "0"));
    }
}

function doPost(e) {
    try {
        var payload = parseWebhookPayloadRadarRouter_(e);
        var route = detectarRotaWebhookRadar_(payload);

        // ==========================================================
        // ROTA 1 — Audiência GE
        // PowerShell envia:
        // {
        //   type: "audiencia_ge",
        //   horario: "18:00",
        //   audiencia: "0.19",
        //   secret: "..."
        // }
        // ==========================================================
        if (route === "audiencia_ge") {
            var audienciaResult;

            if (typeof processarAudienciaGEWebhook_ === "function") {
                audienciaResult = processarAudienciaGEWebhook_(payload);
            } else if (typeof validarEGravarAudienciaGE_ === "function") {
                audienciaResult = validarEGravarAudienciaGE_(payload);
            } else if (typeof gravarAudienciaGE_ === "function") {
                audienciaResult = gravarAudienciaGE_(payload);
            } else {
                audienciaResult = {
                    success: false,
                    ok: false,
                    error: "Função de audiência GE não encontrada. Verifique se o arquivo 19_Audiencia_GE.js está no projeto."
                };
            }

            return radarJsonResponse_(audienciaResult);
        }

        // ==========================================================
        // ROTA 2 — Speech Timer / transcrição
        // Mantém o fluxo original do Speech.
        // ==========================================================
        if (!CFG.SPEECH_WEBHOOK_ENABLED) {
            return radarJsonResponse_({
                success: false,
                error: "Webhook desabilitado"
            });
        }

        if (typeof normalizarWebhookSpeechPayload_ !== "function") {
            return radarJsonResponse_({
                success: false,
                error: "Função normalizarWebhookSpeechPayload_ não encontrada."
            });
        }

        if (typeof processarSpeechImportado_ !== "function") {
            return radarJsonResponse_({
                success: false,
                error: "Função processarSpeechImportado_ não encontrada."
            });
        }

        var sourceData = normalizarWebhookSpeechPayload_(payload);
        var speechResult = processarSpeechImportado_(sourceData);

        return radarJsonResponse_(speechResult);

    } catch (err) {
        return radarJsonResponse_({
            success: false,
            ok: false,
            error: err && err.message ? err.message : String(err),
            stack: err && err.stack ? err.stack : ""
        });
    }
}

function parseWebhookPayloadRadarRouter_(e) {
    try {
        if (typeof parseWebhookPayload_ === "function") {
            return parseWebhookPayload_(e);
        }
    } catch (err) {
        // Se o parser original falhar, tenta o parser simples abaixo.
    }

    try {
        var body = e && e.postData && e.postData.contents ?
            e.postData.contents :
            "{}";

        if (!body || !String(body).trim()) {
            return {};
        }

        return JSON.parse(body);
    } catch (err) {
        return {
            rawBody: e && e.postData && e.postData.contents ?
                e.postData.contents :
                "",
            parseError: err && err.message ? err.message : String(err)
        };
    }
}

function detectarRotaWebhookRadar_(payload) {
    payload = payload || {};

    var type = String(payload.type || "").toLowerCase().trim();
    var source = String(payload.source || "").toLowerCase().trim();

    if (type === "audiencia_ge") {
        return "audiencia_ge";
    }

    if (source.indexOf("audiencia_ge") >= 0) {
        return "audiencia_ge";
    }

    if (
        payload.horario !== undefined &&
        (
            payload.audiencia !== undefined ||
            payload.audiencia_com_virgula !== undefined
        )
    ) {
        return "audiencia_ge";
    }

    return "speech";
}

function radarJsonResponse_(obj) {
    return ContentService
        .createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
}

function processarSpeechImportado_(sourceData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const signature = buildSpeechSignature_(sourceData);
        if (jaFoiProcessadoSpeechHoje_(signature)) return {
            success: !0,
            skipped: !0,
            reason: "already_processed"
        };
        if (!sourceData.programDate) throw Error("Webhook sem data válida do programa.");
        preencherInputComSpeechGE_(sourceData);
        const pipelineResult = executarPipelineCompletoGazeta_({
            skipArchive: !0,
            skipUi: !0
        });
        const syncResult = sincronizarTermosMinutoAMinutoParaHistoricoGE_(sourceData.programDate);
        marcarSpeechComoProcessado_(signature, sourceData);
        return {
            success: !0,
            skipped: !1,
            pipeline: pipelineResult,
            historico: syncResult
        };
    } finally {
        lock.releaseLock();
    }
}

function executarPipelineCompletoGazeta_(options) {
    const opts = options || {};
    const ctx = runAnalysisSkeleton_({
        loadInput: loadGEInputFromSheet_,
        onEmptyInput: () => {
            throw Error('Aba "' + CFG.GE_INPUT_SHEET + '" está vazia ou sem timestamps válidos.');
        },
        processInput: processAndMapTimestampsGE_,
        onEmptyProcessed: () => {
            throw Error("Não foi possível processar a transcrição da Gazeta Esportiva.");
        },
        ensureSheets: ensureGESheets_,
        buildTimeline: buildTimelineGE_
    });
    if (!ctx) throw Error("Falha ao montar o contexto da análise GE.");
    writeAllOutputsGE_(ctx.timeline, ctx.minuteTopTerms, ctx.stop, ctx.input.rows);
    const report = gerarRelatorioGE_(ctx.timeline);
    if (CFG.EMAIL_ENABLED) enviarEmailRelatorioGE_(report);
    if (!opts.skipArchive) {
        const programDate = extrairDataProgramaGE_() || new Date();
        arquivarProgramaGE_(programDate);
    }
    return {
        rawRows: ctx.rawInput.rows.length,
        outputRows: ctx.timeline.length,
        emailed: !!CFG.EMAIL_ENABLED
    };
}

function obterSpeechSourceMode_() {
    const prop = PropertiesService.getScriptProperties().getProperty("SPEECH_SOURCE_MODE");
    return (prop || CFG.SPEECH_SOURCE_MODE || "DRIVE").toUpperCase();
}

function carregarSpeechDoDia_(options) {
    const opts = options || {};
    if ("WEBHOOK" === opts.sourceMode) return null;
    return buscarArquivoSpeechDoDiaNoDrive_(opts);
}

function buscarArquivoSpeechDoDiaNoDrive_(options) {
    const opts = options || {};
    const folderId = PropertiesService.getScriptProperties().getProperty("SPEECH_DRIVE_FOLDER_ID") || CFG.SPEECH_DRIVE_FOLDER_ID;
    if (!folderId) throw Error("SPEECH_DRIVE_FOLDER_ID não configurado.");
    const candidates = listarArquivosSpeechCandidatosNoDrive_(folderId);
    const chosen = escolherMelhorArquivoSpeechNoDrive_(candidates, opts);
    if (!chosen) return null;
    Logger.log("[Speech] Arquivo escolhido: " + chosen.fileName + " | motivo: " + chosen.selectionReason);
    return normalizarSpeechSourceData_({
        sourceType: "DRIVE",
        fileId: chosen.fileId,
        fileName: chosen.fileName,
        updatedAt: chosen.lastUpdated,
        youtubeTitle: chosen.parsed.youtubeTitle,
        youtubeUrl: chosen.parsed.youtubeUrl,
        lines: chosen.parsed.lines,
        programDate: chosen.programDate
    });
}

function listarArquivosSpeechCandidatosNoDrive_(folderId) {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const candidates = [];
    while (files.hasNext()) {
        const file = files.next();
        const name = file.getName() || "";
        const mimeType = file.getMimeType() || "";
        const createdAt = getFileCreatedAtSafe_(file);
        const updatedAt = file.getLastUpdated();
        const nameAccepted = isNomeArquivoSpeechAceito_(name);
        let parsed = null;
        let programDate = null;
        let dateCandidates = [];
        let accepted = !1;
        let reason = "";
        Logger.log("[Speech] Encontrado no Drive | nome: " + name + " | mimeType: " + mimeType + " | criado: " + formatLogDate_(createdAt) + " | modificado: " + formatLogDate_(updatedAt));
        if (!nameAccepted) {
            reason = "nome_fora_do_padrao";
            Logger.log("[Speech] Descartado | " + name + " | motivo: " + reason);
            continue;
        }
        try {
            parsed = parseSpeechXlsxBlob_(file.getBlob());
            if (!parsed.lines || parsed.lines.length === 0) {
                reason = "xlsx_sem_linhas";
            } else {
                dateCandidates = extrairDatasCandidatasDoNomeArquivoSpeech_(name);
                programDate = (dateCandidates.length ? dateCandidates[0] : null) ||
                    extrairDataProgramaSpeech_(parsed.youtubeTitle, name, updatedAt);
                if (!dateCandidates.length && programDate) dateCandidates = [programDate];
                accepted = !0;
                reason = programDate ? "candidato_valido_com_data" : "candidato_valido_sem_data";
            }
        } catch (err) {
            reason = "erro_parse_xlsx: " + err.message;
        }
        Logger.log("[Speech] " + (accepted ? "Aceito" : "Descartado") + " | " + name + " | motivo: " + reason + (programDate ? " | data: " + formatDateKey_(programDate) : ""));
        if (!accepted) continue;
        candidates.push({
            fileId: file.getId(),
            fileName: name,
            mimeType: mimeType,
            createdAt: createdAt,
            lastUpdated: updatedAt,
            parsed: parsed,
            programDate: programDate,
            dateCandidates: dateCandidates
        });
    }
    return candidates;
}

function escolherMelhorArquivoSpeechNoDrive_(candidates, options) {
    const opts = options || {};
    if (!candidates || candidates.length === 0) return null;
    const expectedDate = normalizeDateOnly_(opts.expectedDate || new Date());
    const expectedKey = expectedDate ? formatDateKey_(expectedDate) : "";
    const datasDoCandidato_ = candidate =>
        (candidate.dateCandidates && candidate.dateCandidates.length)
            ? candidate.dateCandidates
            : (candidate.programDate ? [candidate.programDate] : []);
    const datedMatches = candidates.filter(candidate => {
        const dataBate = datasDoCandidato_(candidate)
            .find(dt => formatDateKey_(dt) === expectedKey);
        if (dataBate) candidate.programDate = dataBate;
        return !!dataBate;
    });
    if (datedMatches.length > 0) {
        const bestDated = sortSpeechCandidatesByRecency_(datedMatches)[0];
        bestDated.selectionReason = "data_correspondente_ao_dia";
        return bestDated;
    }
    // Fallback controlado: só aceita o arquivo mais recente se ele tiver sido
    // modificado HOJE e não tiver data divergente identificável. Um arquivo de
    // outro dia NUNCA é importado silenciosamente.
    const fallback = sortSpeechCandidatesByRecency_(candidates)[0];
    const modificadoHoje = fallback.lastUpdated &&
        formatDateKey_(normalizeDateOnly_(fallback.lastUpdated)) === expectedKey;
    const datasFallback = datasDoCandidato_(fallback);
    const dataDivergente = datasFallback.length > 0 &&
        !datasFallback.some(dt => formatDateKey_(dt) === expectedKey);
    if (!modificadoHoje || dataDivergente) {
        Logger.log("[Speech] Nenhum candidato do dia " + expectedKey +
            "; candidato mais recente rejeitado (" +
            (dataDivergente ? "data divergente: " + formatDateKey_(fallback.programDate)
                : "não modificado hoje") + "): " + (fallback.fileName || "?"));
        try {
            logGE_("IMPORTACAO", "Arquivo rejeitado por segurança: " +
                (fallback.fileName || "?") + " — " +
                (dataDivergente ? "data divergente" : "não modificado hoje"));
        } catch (e) {}
        return null;
    }
    fallback.selectionReason = "fallback_arquivo_mais_recente_modificado_hoje";
    Logger.log("[Speech] Sem data no nome; usando XLSX modificado hoje: " + (fallback.fileName || "?"));
    return fallback;
}

function sortSpeechCandidatesByRecency_(candidates) {
    return candidates.slice().sort((a, b) => b.lastUpdated.getTime() - a.lastUpdated.getTime());
}

function extrairDataDoNomeArquivoSpeech_(fileName) {
    const datas = extrairDatasCandidatasDoNomeArquivoSpeech_(fileName);
    return datas.length ? datas[0] : null;
}

// O bloco de 6 dígitos do nome pode estar em DDMMYY (transcricao070726.xlsx)
// ou YYMMDD (transcricao260707.xlsx) — os dois formatos circulam na pasta do
// Drive. Retorna todas as interpretações válidas; a escolha do arquivo decide
// qual delas bate com a data esperada do programa.
function extrairDatasCandidatasDoNomeArquivoSpeech_(fileName) {
    const name = (fileName || "").toString();
    const match = name.match(/^transcricao(?:[_-]?)(\d{6})\.xlsx$/i);
    if (!match) return [];
    const digits = match[1];
    const combos = [
        { // DDMMYY
            day: parseInt(digits.substring(0, 2), 10),
            month: parseInt(digits.substring(2, 4), 10),
            year: 2000 + parseInt(digits.substring(4, 6), 10)
        },
        { // YYMMDD
            year: 2000 + parseInt(digits.substring(0, 2), 10),
            month: parseInt(digits.substring(2, 4), 10),
            day: parseInt(digits.substring(4, 6), 10)
        }
    ];
    const datas = [];
    const vistos = {};
    combos.forEach(c => {
        if (c.day < 1 || c.day > 31 || c.month < 1 || c.month > 12) return;
        const dt = new Date(c.year, c.month - 1, c.day);
        if (!vistos[dt.getTime()]) {
            vistos[dt.getTime()] = true;
            datas.push(dt);
        }
    });
    return datas;
}

function isNomeArquivoSpeechAceito_(fileName) {
    const name = (fileName || "").toString().toLowerCase();
    return "transcricao.xlsx" === name || /^transcricao.*\.xlsx$/i.test(fileName || "");
}

function getFileCreatedAtSafe_(file) {
    try {
        return file.getDateCreated();
    } catch (e) {
        return null;
    }
}

function formatLogDate_(date) {
    if (!date) return "n/d";
    return Utilities.formatDate(date, CFG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
}

function parseSpeechXlsxBlob_(blob) {
    const zipBlob = blob.copyBlob().setContentType("application/zip");
    const entries = Utilities.unzip(zipBlob);
    const fileMap = {};
    for (const entry of entries) fileMap[entry.getName()] = entry.getDataAsString();
    const sharedStrings = parseSharedStringsXml_(fileMap["xl/sharedStrings.xml"]);
    const sheetXml = fileMap["xl/worksheets/sheet1.xml"];
    if (!sheetXml) throw Error("Planilha XLSX sem xl/worksheets/sheet1.xml");
    const rows = parseSheetRowsXml_(sheetXml, sharedStrings);
    const youtubeTitle = rows[0] || "";
    const youtubeUrl = rows[1] || "";
    const lines = [];
    const re = /^\((\d{1,2}:\d{2}(?::\d{2})?)\)\s*(.+)$/;
    for (let i = 4; i < rows.length; i++) {
        const raw = (rows[i] || "").toString().trim();
        if (!raw) continue;
        const match = raw.match(re);
        if (!match) continue;
        lines.push({
            timestamp: match[1],
            text: match[2]
        });
    }
    return {
        youtubeTitle: youtubeTitle,
        youtubeUrl: youtubeUrl,
        lines: lines
    };
}

function parseSharedStringsXml_(xmlText) {
    if (!xmlText) return [];
    const doc = XmlService.parse(xmlText);
    const root = doc.getRootElement();
    const out = [];
    const sis = root.getChildren();
    for (const si of sis) out.push(extractXmlText_(si));
    return out;
}

function parseSheetRowsXml_(xmlText, sharedStrings) {
    const doc = XmlService.parse(xmlText);
    const root = doc.getRootElement();
    const rows = [];
    const descendants = root.getDescendants();
    for (const node of descendants) {
        if (!node.getType || node.getType() !== XmlService.ContentTypes.ELEMENT) continue;
        const element = node.asElement();
        if ("row" !== element.getName()) continue;
        let rowValue = "";
        const cells = element.getChildren();
        for (const cell of cells) {
            if ("c" !== cell.getName()) continue;
            rowValue = parseXlsxCellValue_(cell, sharedStrings);
            break;
        }
        rows.push(rowValue);
    }
    return rows;
}

function parseXlsxCellValue_(cell, sharedStrings) {
    const type = cell.getAttribute("t") ? cell.getAttribute("t").getValue() : "";
    if ("inlineStr" === type) return extractXmlText_(cell);
    const valueEl = getFirstChildByName_(cell, "v");
    if (!valueEl) return "";
    const raw = valueEl.getText() || "";
    if ("s" === type) {
        const idx = parseInt(raw, 10);
        return isNaN(idx) ? "" : sharedStrings[idx] || "";
    }
    return raw;
}

function extractXmlText_(element) {
    let text = "";
    const descendants = element.getDescendants();
    for (const node of descendants)
        if (node.getType && node.getType() === XmlService.ContentTypes.TEXT) text += node.getValue();
    return text || element.getText() || "";
}

function getFirstChildByName_(element, name) {
    const children = element.getChildren();
    for (const child of children)
        if (child.getName() === name) return child;
    return null;
}

function extrairDataProgramaSpeech_(title, fileName, fallbackDate) {
    return extrairDataDeTexto_(title) || extrairDataDeTexto_(fileName) || normalizeDateOnly_(fallbackDate);
}

function extrairDataDeTexto_(text) {
    const raw = (text || "").toString();
    const patterns = [/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/, /(\d{1,2})-(\d{1,2})-(\d{2,4})/, /(\d{4})-(\d{2})-(\d{2})/];
    for (const pattern of patterns) {
        const match = raw.match(pattern);
        if (!match) continue;
        let day, month, year;
        if (pattern.source.indexOf("(\\d{4})-(\\d{2})-(\\d{2})") === 0) {
            year = parseInt(match[1], 10);
            month = parseInt(match[2], 10);
            day = parseInt(match[3], 10);
        } else {
            day = parseInt(match[1], 10);
            month = parseInt(match[2], 10);
            year = parseInt(match[3], 10);
            if (year < 100) year += 2000;
        }
        return new Date(year, month - 1, day);
    }
    return null;
}

function preencherInputComSpeechGE_(sourceData) {
    if (!sourceData.lines || sourceData.lines.length === 0) throw Error("Transcrição vazia recebida do Speech Timer.");
    const ss = getRadarSpreadsheet_();
    let sheet = ss.getSheetByName(CFG.GE_INPUT_SHEET);
    if (!sheet) sheet = ss.insertSheet(CFG.GE_INPUT_SHEET);
    sheet.clear();
    const rows = [
        [sourceData.youtubeTitle || ""],
        [sourceData.youtubeUrl || ""],
        [""],
        ["Transcript:"]
    ];
    for (const line of sourceData.lines) rows.push([`(${line.timestamp}) ${line.text}`]);
    sheet.getRange(1, 1, rows.length, 1).setValues(rows);
    sheet.setColumnWidth(1, 800);

    // Status visível da importação (célula B1 + nota), exigência operacional:
    // arquivo usado, modificação, linhas, data e se houve fallback.
    try {
        const statusTxt = [
            "✅ IMPORTADA " + Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy HH:mm"),
            "Arquivo: " + (sourceData.fileName || sourceData.youtubeTitle || "?"),
            "Modificado: " + (sourceData.updatedAt
                ? Utilities.formatDate(sourceData.updatedAt, CFG.TIMEZONE, "dd/MM/yyyy HH:mm") : "?"),
            "Linhas com timestamp: " + sourceData.lines.length,
            "Data do programa: " + (sourceData.programDate
                ? Utilities.formatDate(sourceData.programDate, CFG.TIMEZONE, "dd/MM/yyyy") : "?"),
            "Seleção: " + (sourceData.selectionReason || "direta")
        ].join("\n");
        sheet.getRange(1, 2).setValue(statusTxt.split("\n")[0] + " — " +
            (sourceData.fileName || "")).setNote(statusTxt).setFontColor("#1E8449");
        logGE_("IMPORTACAO", statusTxt.replace(/\n/g, " | "));
    } catch (e) { /* status é informativo */ }
    Logger.log("[Speech] Input preenchida com " + sourceData.lines.length + " linhas de transcrição.");
}

function normalizarSpeechSourceData_(data) {
    return {
        sourceType: data.sourceType || "DRIVE",
        fileId: data.fileId || "",
        fileName: data.fileName || "",
        updatedAt: data.updatedAt || new Date(),
        youtubeTitle: data.youtubeTitle || "",
        youtubeUrl: data.youtubeUrl || "",
        selectionReason: data.selectionReason || "",
        lines: (data.lines || []).filter(line => line && line.timestamp && line.text),
        programDate: normalizeDateOnly_(data.programDate)
    };
}

function parseWebhookPayload_(e) {
    if (!e || !e.postData || !e.postData.contents) throw Error("Webhook sem payload.");
    return JSON.parse(e.postData.contents);
}

function normalizarWebhookSpeechPayload_(payload) {
    let sourceData = null;
    if (payload.fileId) {
        const file = DriveApp.getFileById(payload.fileId);
        const parsed = parseSpeechXlsxBlob_(file.getBlob());
        sourceData = {
            sourceType: "WEBHOOK",
            fileId: file.getId(),
            fileName: file.getName(),
            updatedAt: file.getLastUpdated(),
            youtubeTitle: parsed.youtubeTitle,
            youtubeUrl: parsed.youtubeUrl,
            lines: parsed.lines,
            programDate: payload.programDate ? new Date(payload.programDate) : extrairDataProgramaSpeech_(parsed.youtubeTitle, file.getName(), file.getLastUpdated())
        };
    } else {
        sourceData = {
            sourceType: "WEBHOOK",
            fileId: payload.fileId || "",
            fileName: payload.fileName || "",
            updatedAt: payload.updatedAt ? new Date(payload.updatedAt) : new Date(),
            youtubeTitle: payload.title || payload.youtubeTitle || "",
            youtubeUrl: payload.youtubeUrl || payload.url || "",
            lines: payload.lines || [],
            programDate: payload.programDate ? new Date(payload.programDate) : extrairDataProgramaSpeech_(payload.title || payload.youtubeTitle, payload.fileName, new Date())
        };
    }
    return normalizarSpeechSourceData_(sourceData);
}

function buildSpeechSignature_(sourceData) {
    const dateKey = sourceData.programDate ? formatDateKey_(sourceData.programDate) : "sem_data";
    const updatedPart = sourceData.updatedAt ? sourceData.updatedAt.getTime() : "";
    const filePart = sourceData.fileId || hashText_(JSON.stringify(sourceData.lines || []).substring(0, 5000));
    return [dateKey, sourceData.sourceType || "UNKNOWN", filePart, updatedPart].join("|");
}

function jaFoiProcessadoSpeechHoje_(signature) {
    return PropertiesService.getScriptProperties().getProperty(CFG.SPEECH_PROCESSED_PREFIX + signature) === "1";
}

function marcarSpeechComoProcessado_(signature, sourceData) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(CFG.SPEECH_PROCESSED_PREFIX + signature, "1");
    props.setProperty(CFG.SPEECH_PROCESSED_PREFIX + "last", JSON.stringify({
        signature: signature,
        programDate: sourceData.programDate ? formatDateKey_(sourceData.programDate) : "",
        fileId: sourceData.fileId || "",
        fileName: sourceData.fileName || "",
        updatedAt: sourceData.updatedAt ? sourceData.updatedAt.toISOString() : ""
    }));
}

function maybeEnviarAlertaSpeechAusente_() {
    if (!CFG.SPEECH_ALERT_ON_MISSING || !CFG.EMAIL_ENABLED) return;
    MailApp.sendEmail({
        to: CFG.EMAIL_DESTINO,
        subject: "Aviso Radar - transcrição do Speech Timer não encontrada",
        body: "A automação procurou a transcrição do dia e não encontrou um XLSX válido para processamento."
    });
}

function normalizeDateOnly_(date) {
    if (!date) return null;
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDateKey_(date) {
    return Utilities.formatDate(date, CFG.TIMEZONE, "yyyy-MM-dd");
}

function jsonResponse_(payload) {
    return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function getRadarSpreadsheet_() {
    try {
        return SpreadsheetApp.getActive();
    } catch (e) {
        return SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    }
}

function testeImportacaoSpeechHoje() {
    return testeImportacaoSpeechHoje_();
}

function instalarTriggersRadar() {
    return instalarTriggersRadar_();
}

function processarSpeechDoDia() {
    return processarSpeechDoDia_();
}
function testeImportacaoSpeechHoje() {
  return testeImportacaoSpeechHoje_();
}

function instalarTriggersRadar() {
  return instalarTriggersRadar_();
}

function processarSpeechDoDia() {
  return processarSpeechDoDia_();
}