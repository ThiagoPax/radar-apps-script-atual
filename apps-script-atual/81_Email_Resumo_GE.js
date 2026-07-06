// =============================================================================
// 09_Email_GE.js — Geração de relatório e envio de email (Gazeta Esportiva)
// =============================================================================

// --- Extração de data do programa GE ---

function extrairDataProgramaGE_() {
    const sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_INPUT_SHEET);
    if (!sheet || sheet.getLastRow() < 1) return getHojeSaoPaulo_();
    const values = sheet.getRange(1, 1, Math.min(sheet.getLastRow(), 5), 1).getValues().flat().filter(Boolean).map(value => value.toString());
    const patterns = [/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/, /(\d{1,2})-(\d{1,2})-(\d{2,4})/, /(\d{1,2})\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*\s+(\d{4})/i];
    const monthMap = {
        jan: 1,
        fev: 2,
        mar: 3,
        abr: 4,
        mai: 5,
        jun: 6,
        jul: 7,
        ago: 8,
        set: 9,
        out: 10,
        nov: 11,
        dez: 12
    };
    for (const rawValue of values)
        for (const pattern of patterns) {
            const match = rawValue.match(pattern);
            if (!match) continue;
            let day;
            let month;
            let year;
            if (pattern.source.includes("jan|fev")) {
                day = parseInt(match[1], 10);
                month = monthMap[match[2].toLowerCase().substring(0, 3)] || 1;
                year = parseInt(match[3], 10);
            } else {
                day = parseInt(match[1], 10);
                month = parseInt(match[2], 10);
                year = parseInt(match[3], 10);
                if (year < 100) year += 2000;
            }
            return new Date(year, month - 1, day);
        }
    return getHojeSaoPaulo_()
}

function getHojeSaoPaulo_() {
    const now = new Date();
    const year = parseInt(Utilities.formatDate(now, "America/Sao_Paulo", "yyyy"), 10);
    const month = parseInt(Utilities.formatDate(now, "America/Sao_Paulo", "MM"), 10);
    const day = parseInt(Utilities.formatDate(now, "America/Sao_Paulo", "dd"), 10);
    return new Date(year, month - 1, day)
}

// --- Geração do relatório GE ---

function gerarRelatorioGE_(e) {
    const t = SpreadsheetApp.getActive(),
        o = extrairDataProgramaGE_() || new Date,
        a = Utilities.formatDate(o, "America/Sao_Paulo", "dd/MM/yyyy"),
        r = ["Domingo", "Segunda-feira", "Ter\xe7a-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "S\xe1bado"][o.getDay()],
        n = t.getSheetByName(CFG.GE_OUT_MACRO),
        s = t.getSheetByName(CFG.GE_OUT_SUB),
        i = t.getSheetByName(CFG.GE_OUT_TOP),
        c = t.getSheetByName(CFG.GE_OUT_TIME),
        l = [];
    if (n && n.getLastRow() > 1) {
        const e = n.getRange(2, 1, n.getLastRow() - 1, 5).getValues();
        for (const t of e) t[0] && "Merchan" !== t[0] && "" !== t[0] && l.length < 5 && l.push({
            tema: t[0],
            minutos: t[1],
            percent: t[2],
            primeira: t[3],
            ultima: t[4]
        })
    }
    const d = [];
    if (s && s.getLastRow() > 1) {
        const e = s.getRange(2, 1, Math.min(s.getLastRow() - 1, 10), 6).getValues();
        for (const t of e) t[0] && "" !== t[0] && d.push({
            tema: t[0],
            subtema: t[1],
            minutos: t[2],
            percent: t[3]
        })
    }
    const u = [];
    if (i && i.getLastRow() > 1) {
        const e = i.getRange(2, 1, Math.min(i.getLastRow() - 1, 5), 3).getValues();
        for (const t of e) t[0] && "" !== t[0] && u.push({
            termo: t[0],
            contagem: t[1],
            minutos: t[2]
        })
    }
    let m = [];
    if (c && c.getLastRow() > 1) {
        const e = c.getRange(2, 1, Math.min(c.getLastRow() - 1, 60), 6).getValues();
        for (const t of e)
            if (t[1] && "MERCHAN" !== t[1] && t[5]) {
                const e = {
                    hora: t[0],
                    tema: t[1],
                    subtema: t[2] || "",
                    trecho: t[5].toString().substring(0, 200)
                };
                m.push(e)
            }
    }
    return {
        data: a,
        diaSemana: r,
        resumo: gerarResumoComIA_GE_(l, d, m, a, u),
        macrotemas: l,
        subtemas: d,
        topTermos: u
    }
}

// --- Resumo do programa GE via IA ---

function gerarResumoComIA_GE_(e, t, o, a, r) {
    if (!hasOpenAIKey_()) return ["Resumo n\xe3o dispon\xedvel (OpenAI n\xe3o configurada)."];
    const n = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    let s = `Data: ${a}\n\n`;
    s += `Top 5 Macrotemas: ${e.map(e=>`${e.tema} (${e.minutos}min, ${e.percent})`).join(", ")}\n\n`, s += `Top 5 Termos: ${r.map(e=>`${e.termo} (${e.contagem}x)`).join(", ")}\n\n`, s += "Trechos relevantes:\n";
    for (const e of o.slice(0, 20)) s += `[${e.hora}] ${e.tema}${e.subtema?"/"+e.subtema:""}: ${e.trecho}\n`;
    const i = `Voc\xea \xe9 um analista esportivo. Com base nos dados abaixo do programa GAZETA ESPORTIVA de ${a}, escreva um resumo em 3 frases curtas e objetivas destacando:\n1) O tema principal abordado\n2) Destaques ou pol\xeamicas discutidas\n3) Outros assuntos relevantes\n\nDADOS:\n${s}\n\nREGRAS:\n- Seja direto e objetivo\n- Use linguagem jornal\xedstica\n- Mencione nomes de jogadores/t\xe9cnicos quando relevante\n- Cada frase deve ter no m\xe1ximo 20 palavras\n- Retorne APENAS as 3 frases, uma por linha`;
    try {
        const e = fetchJsonSafe_(CFG.OPENAI_ENDPOINT, {
            method: "POST",
            headers: {
                Authorization: "Bearer " + n,
                "Content-Type": "application/json"
            },
            payload: JSON.stringify({
                model: CFG.OPENAI_MODEL,
                messages: [{
                    role: "user",
                    content: i
                }],
                temperature: .3,
                max_tokens: 500
            }),
            muteHttpExceptions: !0
        });
        if (e.choices && e.choices[0] && e.choices[0].message) return e.choices[0].message.content.trim().split("\n").filter(e => e.trim().length > 0)
    } catch (e) {
        Logger.log("Erro ao gerar resumo GE com IA: " + e.message)
    }
    return ["N\xe3o foi poss\xedvel gerar o resumo automaticamente."]
}

// --- Envio de email de relatório GE ---

function enviarEmailRelatorioGE_(e) {
    const t = `\u{1f4fa} Relat\xf3rio Gazeta Esportiva - ${e.diaSemana}, ${e.data}`,
        o = "#e74c3c",
        a = "#fdf2f2";
    let r = `\n    <div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0 auto;">\n      <h1 style="color: ${o}; border-bottom: 2px solid ${o}; padding-bottom: 10px;">\n        \u{1f4fa} Gazeta Esportiva - ${e.diaSemana}, ${e.data}\n      </h1>\n      \n      <h2 style="color: #333; margin-top: 30px;">\u{1f4dd} Resumo do Programa</h2>\n      <div style="background: ${a}; padding: 15px; border-radius: 8px; border-left: 4px solid ${o}; line-height: 1.8;">\n  `;
    for (const t of e.resumo) r += `<p style="margin: 8px 0; color: #333;">${t}</p>`;
    r += `\n      </div>\n      \n      <h2 style="color: #333; margin-top: 30px;">\u{1f3c6} Top 5 Termos Mais Falados</h2>\n      <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">\n        <tr style="background: ${o}; color: white;">\n          <th style="padding: 10px; text-align: left;">Termo</th>\n          <th style="padding: 10px; text-align: center;">Men\xe7\xf5es</th>\n          <th style="padding: 10px; text-align: center;">Minutos</th>\n        </tr>\n  `;
    for (let t = 0; t < e.topTermos.length; t++) {
        const o = e.topTermos[t];
        r += `\n        <tr style="background: ${t%2==0?a:"#ffffff"};">\n          <td style="padding: 10px; border-bottom: 1px solid #ddd;"><strong>${o.termo}</strong></td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${o.contagem}</td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${o.minutos}</td>\n        </tr>\n    `
    }
    r += `\n      </table>\n      \n      <h2 style="color: #333; margin-top: 30px;">\u{1f4ca} Top 5 Macrotemas</h2>\n      <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">\n        <tr style="background: ${o}; color: white;">\n          <th style="padding: 10px; text-align: left;">Tema</th>\n          <th style="padding: 10px; text-align: center;">Minutos</th>\n          <th style="padding: 10px; text-align: center;">%</th>\n        </tr>\n  `;
    for (let t = 0; t < e.macrotemas.length; t++) {
        const o = e.macrotemas[t];
        r += `\n        <tr style="background: ${t%2==0?a:"#ffffff"};">\n          <td style="padding: 10px; border-bottom: 1px solid #ddd;"><strong>${o.tema}</strong></td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${o.minutos}</td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${o.percent}</td>\n        </tr>\n    `
    }
    r += `\n      </table>\n      \n      <h2 style="color: #333; margin-top: 30px;">\u{1f3af} Top 10 Subtemas</h2>\n      <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">\n        <tr style="background: ${o}; color: white;">\n          <th style="padding: 10px; text-align: left;">Tema</th>\n          <th style="padding: 10px; text-align: left;">Subtema</th>\n          <th style="padding: 10px; text-align: center;">Min</th>\n          <th style="padding: 10px; text-align: center;">%</th>\n        </tr>\n  `;
    for (let t = 0; t < e.subtemas.length; t++) {
        const o = e.subtemas[t];
        r += `\n        <tr style="background: ${t%2==0?a:"#ffffff"};">\n          <td style="padding: 10px; border-bottom: 1px solid #ddd;">${o.tema}</td>\n          <td style="padding: 10px; border-bottom: 1px solid #ddd;">${o.subtema}</td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${o.minutos}</td>\n          <td style="padding: 10px; text-align: center; border-bottom: 1px solid #ddd;">${o.percent}</td>\n        </tr>\n    `
    }
    r += `\n      </table>\n      \n      <hr style="margin-top: 40px; border: none; border-top: 1px solid #ddd;">\n      <p style="color: #999; font-size: 12px; text-align: center;">\n        Relat\xf3rio gerado automaticamente pelo Analisador Gazeta Esportiva<br>\n        Emitido em ${(new Date).toLocaleString("pt-BR",{timeZone:"America/Sao_Paulo"})}\n      </p>\n    </div>\n  `, MailApp.sendEmail({
        to: CFG.EMAIL_DESTINO,
        subject: t,
        htmlBody: r
    })
}