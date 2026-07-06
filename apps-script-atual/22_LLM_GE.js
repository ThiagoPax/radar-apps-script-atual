// =============================================================================
// 04_LLM.js — Integração com LLMs: OpenAI (classificação) e Gemini (resumos)
// =============================================================================

// --- OpenAI ---

function hasOpenAIKey_() {
    const e = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    return !(!e || !e.trim())
}

function classifyWithLLM_(e) {
    const t = PropertiesService.getScriptProperties(),
        o = new Map,
        a = [];
    for (const r of e) {
        const e = "CLS_" + hashText_(r.text).substring(0, 40),
            n = t.getProperty(e);
        if (n) try {
            o.set(r.idx + "", JSON.parse(n));
            continue
        } catch (e) {}
        a.push(r)
    }
    for (let e = 0; e < a.length; e += CFG.LLM_BATCH_SIZE) {
        const r = a.slice(e, e + CFG.LLM_BATCH_SIZE),
            n = callOpenAI_(r);
        if (!n) continue;
        let s;
        try {
            const e = n.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
            s = JSON.parse(e)
        } catch (e) {
            Logger.log("LLM parse failed: " + n.substring(0, 500));
            continue
        }
        if (Array.isArray(s))
            for (const e of s) {
                const a = e.idx + "";
                o.set(a, {
                    tema: e.tema,
                    subtema: e.subtema
                });
                const n = r.find(e => e.idx + "" === a);
                if (n && n.text) {
                    const o = "CLS_" + hashText_(n.text).substring(0, 40);
                    try {
                        t.setProperty(o, JSON.stringify({
                            tema: e.tema,
                            subtema: e.subtema
                        }))
                    } catch (e) {}
                }
            }
    }
    return o
}

function callOpenAI_(e) {
    const t = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    if (!t || !t.trim()) return null;
    const o = CFG.FIXED_THEMES.join(", "),
        a = e.map(e => ({
            idx: e.idx,
            hora: e.hora,
            topTerms: e.topTerms,
            trecho: e.text.substring(0, 500)
        })),
        r = `Você é um classificador de transcrições do programa Gazeta Esportiva (TV Gazeta, 18h-19h, foco no futebol paulista e no calendário esportivo do momento).\nClassifique cada trecho em TEMA e SUBTEMA.\n\nREGRAS:\n1) TEMA deve ser EXATAMENTE um: ${o}\n2) Priorize os clubes paulistas (CORINTHIANS, PALMEIRAS, SÃO PAULO, SANTOS) quando mencionados\n3) Durante eventos excepcionais (ex.: Copa do Mundo), use os temas do evento: COPA DO MUNDO, SELEÇÃO BRASILEIRA, seleções (ARGENTINA, FRANÇA, INGLATERRA, NORUEGA, MÉXICO), FIFA, NEYMAR, ANCELOTTI\n4) Jogo específico de seleção na Copa → tema da seleção em foco; visão geral da rodada/tabela da Copa → COPA DO MUNDO\n5) ARBITRAGEM quando o assunto central for erro/decisão de arbitragem ou VAR\n6) MERCHAN é para propaganda (produtos, preços, QR code, telefone 0800)\n7) Resenha/zoeira entre apresentadores sobre times NÃO é MERCHAN\n8) SUBTEMA curto: ${(CFG.SUBTEMAS_SUGERIDOS||["JOGO","CLASSIFICACAO","RESENHA","OUTROS"]).join(", ")}`,
        n = `Classifique:\n${JSON.stringify(a)}\n\nFormato: [{"idx":N,"tema":"TEMA","subtema":"SUBTEMA"}, ...]`;
    try {
        const e = UrlFetchApp.fetch(CFG.OPENAI_ENDPOINT, {
                method: "post",
                contentType: "application/json",
                headers: {
                    Authorization: "Bearer " + t
                },
                payload: JSON.stringify({
                    model: CFG.OPENAI_MODEL,
                    messages: [{
                        role: "system",
                        content: r
                    }, {
                        role: "user",
                        content: n
                    }],
                    temperature: CFG.LLM_TEMPERATURE,
                    max_tokens: CFG.LLM_MAX_TOKENS
                }),
                muteHttpExceptions: !0
            }),
            o = e.getResponseCode(),
            a = e.getContentText();
        if (o >= 200 && o < 300) {
            const e = JSON.parse(a);
            if (e.choices && e.choices[0] && e.choices[0].message) return e.choices[0].message.content
        }
        return Logger.log(`OpenAI error (${o}): ${a.substring(0,500)}`), null
    } catch (e) {
        return Logger.log("OpenAI request failed: " + e.message), null
    }
}

function clearLLMCache() {
    const e = PropertiesService.getScriptProperties(),
        t = e.getKeys();
    let o = 0;
    for (const a of t) a.startsWith("CLS_") && (e.deleteProperty(a), o++);
    SpreadsheetApp.getUi().alert(`Cache limpo: ${o} entradas removidas.`)
}

// --- Gemini ---

function hasGeminiKey_() {
    const e = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    return e && e.length > 0
}

function corrigirNomesComGemini_(e, t, o, a) {
    const r = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!r) return Logger.log("GEMINI_API_KEY não configurada - pulando correção"), null;
    let n = "";
    for (const e in t) n += "\nELENCO " + e + " (2024/2025): " + t[e].slice(0, 35).join(", ") + "\n";
    n || (n = "\n(Elencos não disponíveis - use seu conhecimento de futebol brasileiro)\n");
    const s = `Você é um especialista em futebol brasileiro. Corrija os nomes extraídos de uma transcrição de TV.\n\nDATA DE REFERÊNCIA: ${o}\n\nELENCOS REAIS DOS TIMES (temporada 2024/2025):\n${n}\n\nNOMES EXTRAÍDOS DA TRANSCRIÇÃO (podem ter erros de áudio):\n${e}\n\nTRECHO DA TRANSCRIÇÃO PARA CONTEXTO:\n${a?a.substring(0,3e3):""}\n\nTAREFA:\nPara cada nome extraído:\n1. Compare com os ELENCOS REAIS acima\n2. Se o nome é parecido mas diferente, CORRIJA para o nome correto do elenco\n3. APENAS corrija nomes - NÃO invente informações que não estão na transcrição\n\nCORREÇÕES DE PRONÚNCIA/TRANSCRIÇÃO (erros comuns de áudio):\n- "Memphis de Pai" / "Menfis de Pai" → "Memphis Depay"\n- "Yúri Alberto" / "Yúrio Albe" → "Yuri Alberto"\n- "Thago Mendes" → "Thiago Mendes"\n- "Mateuzin" / "Mateuzinho" → "Matheuzinho"\n- "Felipe Luiz" / "Felipe Luis" → "Filipe Luís"\n- "Anibal" / "Animal" → "Aníbal"\n- "Andrés Gomes" / "André Gomes" → verificar contexto\n\nREGRAS:\n- TÉCNICOS (Dorival Júnior, Abel Ferreira, Filipe Luís): MANTER SE SÃO ASSUNTO\n- Ex-jogadores famosos mencionados: MANTER\n- NÃO adicione informações que não estão na transcrição\n- APENAS corrija a grafia dos nomes que JÁ APARECEM\n\nFORMATO:\nVALIDADO: [Nome] | CLUBE: [Clube] | CONTEXTO: [contexto] | STATUS: CORRETO\nCORRIGIDO: [Nome Original] → [Nome Correto] | CLUBE: [Clube] | CONTEXTO: [contexto] | MOTIVO: [razão]\nREMOVER: [Nome] | CLUBE: [Clube] | MOTIVO: [razão - só se realmente não faz sentido]\n\nIMPORTANTE: \n- Prefira CORRIGIR do que REMOVER\n- NÃO invente transferências ou fatos que não estão na transcrição`;
    try {
        const e = fetchJsonSafe_(CFG.GEMINI_ENDPOINT + "?key=" + r, {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify({
                contents: [{
                    parts: [{
                        text: s
                    }]
                }],
                generationConfig: {
                    temperature: .1,
                    maxOutputTokens: 1500
                }
            }),
            muteHttpExceptions: !0
        });
        if (e.candidates && e.candidates[0] && e.candidates[0].content) {
            const t = e.candidates[0].content.parts[0].text.trim();
            return Logger.log("=== ETAPA 2C: CORREÇÃO GEMINI ==="), Logger.log(t), t
        }
        if (e.error) return Logger.log("Erro Gemini: " + e.error.message), null
    } catch (e) {
        return Logger.log("Erro ao chamar Gemini: " + e), null
    }
    return null
}