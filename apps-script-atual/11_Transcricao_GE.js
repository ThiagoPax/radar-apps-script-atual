// =============================================================================
// 12_Transcricao.js — Importação de transcrições via servidor (JA + GE)
// =============================================================================

// --- Helpers de servidor ---

function runTranscriptionImportFlow_(e) {
    try {
        const t = e.fetchResult();
        t && t.success ? e.onSuccess(t) : e.onFailure(t)
    } catch (t) {
        e.onException(t)
    }
}

function getTranscriptionServerUrl_() {
    let e = PropertiesService.getScriptProperties().getProperty("TRANSCRIPTION_SERVER") || CFG.TRANSCRIPTION_SERVER;
    return !e || e.startsWith("http://") || e.startsWith("https://") || (e = "https://" + e), e && e.endsWith("/") && (e = e.slice(0, -1)), e
}

function verificarServidorConfigurado_() {
    return getTranscriptionServerUrl_() || (SpreadsheetApp.getUi().alert("\u26a0\ufe0f Servidor n\xe3o configurado", "Configure o servidor de transcri\xe7\xe3o primeiro:\n\n1. Fa\xe7a deploy do servidor (veja README)\n2. Use o menu: \u2699\ufe0f Configurar Servidor\n3. Cole a URL do servidor\n\nEnquanto isso, voc\xea pode colar a transcri\xe7\xe3o manualmente na aba Input.", SpreadsheetApp.getUi().ButtonSet.OK), null)
}

function chamarServidorTranscricao_(e, t) {
    try {
        return fetchJsonSafe_(e + "/transcribe", {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(t),
            muteHttpExceptions: !0,
            timeout: 600
        }, {
            checkHttp: !0,
            okCode: 200,
            checkHtml: !0
        })
    } catch (e) {
        return Logger.log("Erro ao chamar servidor: " + e.message), {
            success: !1,
            error: "Erro de conex\xe3o: " + e.message
        }
    }
}

// --- Configuração do servidor (entrada pública via menu) ---

function configurarServidor() {
    const e = SpreadsheetApp.getUi(),
        t = getTranscriptionServerUrl_(),
        o = e.prompt("\u2699\ufe0f Configurar Servidor de Transcri\xe7\xe3o", `URL atual: ${t||"(n\xe3o configurado)"}\n\nDigite a URL do seu servidor (ex: https://seu-projeto.railway.app):\n(Deixe em branco para remover)`, e.ButtonSet.OK_CANCEL);
    if (o.getSelectedButton() === e.Button.OK) {
        const t = o.getResponseText().trim(),
            a = PropertiesService.getScriptProperties();
        t ? (a.setProperty("TRANSCRIPTION_SERVER", t), e.alert("\u2705 Sucesso", "Servidor configurado: " + t, e.ButtonSet.OK)) : (a.deleteProperty("TRANSCRIPTION_SERVER"), e.alert("\u2139\ufe0f Info", "Configura\xe7\xe3o de servidor removida.", e.ButtonSet.OK))
    }
}

// =============================================================================
// Transcrição — Gazeta Esportiva (GE)
// =============================================================================

function importarTranscricaoAutomaticaGazeta() {
    const e = SpreadsheetApp.getUi(),
        t = getTranscriptionServerUrl_();
    if (t) {
        e.alert("\u{1f504} Buscando programa", "Buscando Gazeta Esportiva de hoje no YouTube...\n\nIsso pode levar alguns segundos.", e.ButtonSet.OK);
        try {
            const o = fetchJsonSafe_(t + "/find-today-ge", {
                method: "GET",
                muteHttpExceptions: !0
            });
            if (!o.success) {
                const t = buscarGazetaEsportivaHoje_();
                return void(t ? importarTranscricaoGazeta_(t) : e.alert("\u274c Programa n\xe3o encontrado", 'N\xe3o foi poss\xedvel encontrar o Gazeta Esportiva de hoje.\n\nUse "Importar por URL" para informar manualmente.', e.ButtonSet.OK))
            }
            importarTranscricaoGazeta_(o.url)
        } catch (t) {
            Logger.log("Erro ao buscar programa automaticamente: " + t.message), e.alert("\u274c Erro", "Erro ao buscar programa: " + t.message + '\n\nUse "Importar por URL" para informar manualmente.', e.ButtonSet.OK)
        }
    } else e.alert("\u26a0\ufe0f Servidor n\xe3o configurado", "Configure o servidor de transcri\xe7\xe3o em Configura\xe7\xf5es \u2192 Configurar Servidor.", e.ButtonSet.OK)
}

function buscarGazetaEsportivaHoje_() {
    const e = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    if (!e) return Logger.log("YouTube API Key n\xe3o configurada"), null;
    const t = new Date,
        o = Utilities.formatDate(t, "America/Sao_Paulo", "dd/MM/yy"),
        a = "https://www.googleapis.com/youtube/v3/playlistItems?part=snippet,contentDetails&playlistId=" + CFG.GE_UPLOADS_PLAYLIST_ID + "&maxResults=30&key=" + e;
    try {
        const e = fetchJsonSafe_(a);
        if (!e.items || 0 === e.items.length) return null;
        for (const t of e.items) {
            const e = t.snippet.title,
                a = e.toUpperCase();
            if (a.includes("GAZETA ESPORTIVA") && a.includes("AO VIVO") && e.includes(o)) return "https://www.youtube.com/watch?v=" + t.contentDetails.videoId
        }
        for (const t of e.items) {
            const e = t.snippet.title.toUpperCase();
            if (e.includes("GAZETA ESPORTIVA") && e.includes("AO VIVO")) return "https://www.youtube.com/watch?v=" + t.contentDetails.videoId
        }
    } catch (e) {
        Logger.log("Erro ao buscar no YouTube: " + e.message)
    }
    return null
}

function importarTranscricaoPorUrlGazeta() {
    const e = SpreadsheetApp.getUi(),
        t = e.prompt("\u{1f517} Importar por URL", "Cole a URL do v\xeddeo do Gazeta Esportiva:", e.ButtonSet.OK_CANCEL);
    if (t.getSelectedButton() !== e.Button.OK) return;
    const o = t.getResponseText().trim();
    o ? o.includes("youtube.com") || o.includes("youtu.be") ? importarTranscricaoGazeta_(o) : e.alert("URL inv\xe1lida", "Por favor, forne\xe7a uma URL v\xe1lida do YouTube.", e.ButtonSet.OK) : e.alert("URL vazia")
}

function importarTranscricaoGazeta_(e) {
    const t = SpreadsheetApp.getUi(),
        o = getTranscriptionServerUrl_();
    o ? (t.alert("\u23f3 Processando", "Transcrevendo v\xeddeo do Gazeta Esportiva...\n\nIsso pode levar alguns minutos.", t.ButtonSet.OK), runTranscriptionImportFlow_({
        fetchResult: () => fetchJsonSafe_(o + "/transcribe", {
            method: "POST",
            contentType: "application/json",
            payload: JSON.stringify({
                url: e
            }),
            muteHttpExceptions: !0
        }),
        onSuccess: o => {
            ensureGESheets_();
            const a = SpreadsheetApp.getActive().getSheetByName(CFG.GE_INPUT_SHEET);
            a.clear(), a.getRange(1, 1).setValue(e);
            const r = o.transcription.split("\n");
            for (let e = 0; e < r.length; e++) a.getRange(e + 2, 1).setValue(r[e]);
            t.alert("\u2705 Transcri\xe7\xe3o importada!", `V\xeddeo: ${o.videoTitle}\nDura\xe7\xe3o: ${Math.round(o.videoDuration/60)} minutos\n\nA transcri\xe7\xe3o foi salva na aba "${CFG.GE_INPUT_SHEET}".\n\nClique em "Rodar An\xe1lise Completa" para processar.`, t.ButtonSet.OK)
        },
        onFailure: e => {
            throw Error(e.error || "Erro ao transcrever")
        },
        onException: e => {
            t.alert("\u274c Erro", "Erro ao transcrever: " + e.message, t.ButtonSet.OK)
        }
    })) : t.alert("\u26a0\ufe0f Servidor n\xe3o configurado", "Configure o servidor de transcri\xe7\xe3o primeiro.", t.ButtonSet.OK)
}