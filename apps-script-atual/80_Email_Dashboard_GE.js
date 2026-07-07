// =============================================================================
// 80_Email_Dashboard_GE.js — Email executivo de fim de dia (21:00)
// =============================================================================
// Assunto: Dashboard Gazeta Esportiva — dd/mm/aaaa — média X,XX
// 8 seções: placar, bullets editoriais, audiência comparada, temas, fala por
// pessoa, breaks Globo, clubes e recomendações. HTML de cartões limpo.

function gerarEmailDashboardGE() {
    const dados = coletarDadosDashboardGE_();
    const html = montarHtmlDashboardGE_(dados);
    const assunto = "Dashboard Gazeta Esportiva — " + dados.dataStr +
        (dados.mediaDia ? " — média " + fmtNumBRGE_(dados.mediaDia) : "");
    MailApp.sendEmail({
        to: CFG.EMAIL_DESTINO,
        subject: assunto,
        htmlBody: html
    });
    logGE_("EMAIL", "Dashboard enviado: " + assunto);
}

function gerarEmailDashboardMenuGE() {
    gerarEmailDashboardGE();
    SpreadsheetApp.getUi().alert("📧 Dashboard enviado para " + CFG.EMAIL_DESTINO);
}

// -----------------------------------------------------------------------------
// Coleta de dados
// -----------------------------------------------------------------------------
function coletarDadosDashboardGE_() {
    const ss = SpreadsheetApp.getActive();
    const hoje = new Date();
    const d = {
        dataStr: Utilities.formatDate(hoje, CFG.TIMEZONE, "dd/MM/yyyy"),
        diaSemana: ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"][hoje.getDay()],
        mediaDia: null, melhorMinuto: null, piorMinuto: null,
        leitura: "", reguas: mediasContextuaisGE_(),
        temas: [], temaMaisTempo: "", temaMelhor: "", temaPior: "",
        comentaristas: [], breaks: {}, clubes: {}, recomendacoes: [],
        bullets: []
    };

    // Audiência do dia (aba Hoje, coluna F)
    try {
        const sh = ss.getSheetByName(CFG.SHEET_HOJE);
        const finais = sh.getRange(3, 6, 60, 1).getValues();
        const temas = sh.getRange(3, 9, 60, 1).getValues();
        let soma = 0, n = 0;
        for (let i = 0; i < 60; i++) {
            const v = finais[i][0];
            if (typeof v !== "number" || !isFinite(v)) continue;
            soma += v; n++;
            const hora = "18:" + String(i).padStart(2, "0");
            const tema = String(temas[i][0] || "");
            if (!d.melhorMinuto || v > d.melhorMinuto.v) d.melhorMinuto = { hora: hora, v: v, tema: tema };
            if (!d.piorMinuto || v < d.piorMinuto.v) d.piorMinuto = { hora: hora, v: v, tema: tema };
        }
        if (n > 0) d.mediaDia = soma / n;
    } catch (e) { logGE_("EMAIL", "coleta Hoje falhou: " + e.message); }

    if (typeof d.mediaDia === "number" && typeof d.reguas.mediaContexto === "number") {
        d.leitura = leituraContextualGE_(d.mediaDia, d.reguas.mediaContexto);
        d.difContexto = d.mediaDia / d.reguas.mediaContexto - 1;
    }

    // Temas do dia (Macrotema) + índice histórico (Tema Minuto)
    try {
        const macro = ss.getSheetByName(CFG.GE_OUT_MACRO);
        if (macro && macro.getLastRow() > 1) {
            const rows = macro.getRange(2, 1, Math.min(macro.getLastRow() - 1, 15), 3).getValues();
            const indices = {};
            const temaSheet = ss.getSheetByName(CFG.GE_OUT_MEDIA_TEMA);
            if (temaSheet && temaSheet.getLastRow() > 1) {
                for (const r of temaSheet.getRange(2, 1, temaSheet.getLastRow() - 1, 6).getValues()) {
                    indices[String(r[0] || "").trim().toUpperCase()] = Number(r[5]);
                }
            }
            for (const r of rows) {
                const nome = String(r[0] || "").trim();
                if (!nome) continue;
                d.temas.push({ tema: nome, minutos: r[1], pct: r[2],
                    indice: indices[nome.toUpperCase()] });
            }
            if (d.temas.length) {
                d.temaMaisTempo = d.temas[0].tema;
                const comIndice = d.temas.filter(t => typeof t.indice === "number" && isFinite(t.indice));
                if (comIndice.length) {
                    comIndice.sort((a, b) => b.indice - a.indice);
                    d.temaMelhor = comIndice[0].tema + " (índice " + comIndice[0].indice.toFixed(2) + ")";
                    const pior = comIndice[comIndice.length - 1];
                    d.temaPior = pior.tema + " (índice " + pior.indice.toFixed(2) + ")";
                }
            }
        }
    } catch (e) { logGE_("EMAIL", "coleta temas falhou: " + e.message); }

    // Comentaristas (aba Comentaristas, quando existir)
    try {
        const com = ss.getSheetByName(CFG.SHEET_COMENTARISTAS);
        if (com && com.getLastRow() > 1) {
            const rows = com.getRange(2, 1, Math.min(com.getLastRow() - 1, 10), 3).getValues();
            for (const r of rows) {
                if (String(r[0] || "").trim()) {
                    d.comentaristas.push({ nome: String(r[0]), col2: r[1], col3: r[2] });
                }
            }
        }
    } catch (e) { /* seção opcional */ }

    // Breaks e clubes
    try { d.breaks = getPrevisaoBreaksHojeGE_(); } catch (e) { d.breaks = {}; }
    try { d.clubes = getContextoClubesHojeGE_(); } catch (e) { d.clubes = {}; }

    // Bullets editoriais
    if (d.temaMaisTempo) d.bullets.push("Principal assunto: " + formatTemaDisplay_(d.temaMaisTempo) + ".");
    if (d.melhorMinuto) d.bullets.push("Melhor momento: " + d.melhorMinuto.hora + " (" +
        fmtNumBRGE_(d.melhorMinuto.v) + ")" + (d.melhorMinuto.tema ? " com " +
        formatTemaDisplay_(d.melhorMinuto.tema) : "") + ".");
    if (d.piorMinuto) d.bullets.push("Ponto de atenção: " + d.piorMinuto.hora + " (" +
        fmtNumBRGE_(d.piorMinuto.v) + ")" + (d.piorMinuto.tema ? " durante " +
        formatTemaDisplay_(d.piorMinuto.tema) : "") + ".");

    // Recomendações para o próximo programa
    if (d.temaMelhor) d.recomendacoes.push("Abertura: apostar em " + d.temaMelhor.split(" (")[0] + ".");
    const clubesQuentes = Object.keys(d.clubes || {}).filter(c =>
        d.clubes[c] && (d.clubes[c].resultado || d.clubes[c].proximoJogo));
    if (clubesQuentes.length) {
        const c = clubesQuentes[0];
        const info = d.clubes[c];
        d.recomendacoes.push("Tema forte: " + formatTemaDisplay_(c) +
            (info.resultado ? " (" + (info.resultado === "V" ? "venceu" :
            info.resultado === "D" ? "perdeu" : "empatou") + " " + info.quando + ")" :
            " (joga " + info.proximoJogo + ")") + ".");
    }
    if (d.temaPior) d.recomendacoes.push("Risco a evitar: alongar " + d.temaPior.split(" (")[0] + ".");
    if (!d.recomendacoes.length) d.recomendacoes.push("Sem dados suficientes para recomendação hoje.");

    return d;
}

// -----------------------------------------------------------------------------
// HTML
// -----------------------------------------------------------------------------
function montarHtmlDashboardGE_(d) {
    const card = (titulo, conteudo) =>
        '<div style="background:#fff;border:1px solid #E5E8E8;border-radius:8px;' +
        'padding:14px 16px;margin-bottom:12px;">' +
        '<div style="font-size:13px;font-weight:bold;color:#1B2631;margin-bottom:8px;">' +
        titulo + '</div>' + conteudo + '</div>';
    const kv = pares => '<table style="font-size:13px;color:#2C3E50;border-collapse:collapse;">' +
        pares.map(p => '<tr><td style="padding:2px 14px 2px 0;color:#7B7D7D;">' + p[0] +
            '</td><td style="padding:2px 0;font-weight:bold;">' + p[1] + '</td></tr>').join("") +
        '</table>';
    const pct = v => (typeof v === "number" && isFinite(v))
        ? ((v >= 0 ? "+" : "") + (v * 100).toFixed(1) + "%") : "—";
    const num = v => (typeof v === "number" && isFinite(v)) ? fmtNumBRGE_(v) : "—";

    let html = '<div style="font-family:Arial,Helvetica,sans-serif;background:#F4F6F6;' +
        'padding:18px;max-width:680px;margin:0 auto;">' +
        '<div style="font-size:17px;font-weight:bold;color:#1B2631;margin-bottom:2px;">📺 Gazeta Esportiva — Dashboard do dia</div>' +
        '<div style="font-size:12px;color:#7B7D7D;margin-bottom:14px;">' + d.diaSemana + ', ' +
        d.dataStr + ' — contexto: ' + (d.reguas.contexto || "NORMAL") + '</div>';

    // 1. Placar do programa
    const corLeitura = /POSITIVA|ACIMA/.test(d.leitura) ? "#27AE60"
        : /NEGATIVA|ABAIXO/.test(d.leitura) ? "#C0392B" : "#7B7D7D";
    html += card("1 · Placar do programa", kv([
        ["Média do dia", num(d.mediaDia)],
        ["Média contextual", num(d.reguas.mediaContexto)],
        ["Diferença", pct(d.difContexto)],
        ["Melhor minuto", d.melhorMinuto ? d.melhorMinuto.hora + " (" + num(d.melhorMinuto.v) + ")" : "—"],
        ["Pior minuto", d.piorMinuto ? d.piorMinuto.hora + " (" + num(d.piorMinuto.v) + ")" : "—"],
        ["Leitura", '<span style="color:' + corLeitura + ';">' + (d.leitura || "—") + '</span>']
    ]));

    // 2. Resumo editorial
    html += card("2 · Resumo editorial",
        '<ul style="margin:0;padding-left:18px;font-size:13px;color:#2C3E50;">' +
        d.bullets.map(b => "<li>" + b + "</li>").join("") + "</ul>");

    // 3. Audiência comparada
    html += card("3 · Audiência comparada", kv([
        ["Média do programa", num(d.mediaDia)],
        ["Mesmo dia da semana", num(d.reguas.mesmoDia)],
        ["Últimos 3 programas", num(d.reguas.ultimos3)],
        ["Últimos 5 programas", num(d.reguas.ultimos5)],
        ["Média histórica geral", num(d.reguas.geral)]
    ]));

    // 4. Temas
    let temasHtml = "";
    if (d.temas.length) {
        temasHtml = '<table style="font-size:12px;border-collapse:collapse;width:100%;">' +
            '<tr style="color:#7B7D7D;text-align:left;"><th style="padding:3px 8px 3px 0;">Tema</th>' +
            '<th style="padding:3px 8px;">Min</th><th style="padding:3px 8px;">%</th>' +
            '<th style="padding:3px 0;">Índice hist.</th></tr>' +
            d.temas.slice(0, 6).map(t =>
                '<tr><td style="padding:3px 8px 3px 0;font-weight:bold;">' + formatTemaDisplay_(t.tema) +
                '</td><td style="padding:3px 8px;">' + (t.minutos || "") +
                '</td><td style="padding:3px 8px;">' + (t.pct || "") +
                '</td><td style="padding:3px 0;">' +
                (typeof t.indice === "number" ? t.indice.toFixed(2) : "—") + '</td></tr>').join("") +
            '</table>' +
            '<div style="font-size:12px;color:#7B7D7D;margin-top:6px;">Melhor: ' +
            (d.temaMelhor || "—") + ' · Pior: ' + (d.temaPior || "—") + '</div>';
    } else temasHtml = '<div style="font-size:13px;color:#7B7D7D;">Sem temas processados hoje.</div>';
    html += card("4 · Temas", temasHtml);

    // 5. Fala por pessoa
    let comHtml;
    if (d.comentaristas.length) {
        comHtml = kv(d.comentaristas.map(c => [c.nome,
            [c.col2, c.col3].filter(x => x !== "" && x !== null && x !== undefined).join(" · ")]));
    } else {
        comHtml = '<div style="font-size:13px;color:#7B7D7D;">Sem dados de escalação/fala ' +
            'registrados hoje (preencher comentaristas do dia no Histórico, linhas 64–65).</div>';
    }
    html += card("5 · Mesa / comentaristas", comHtml);

    // 6. Breaks Globo
    let breaksHtml = '<div style="font-size:13px;color:#7B7D7D;">Sem padrão registrado — alimentar a aba Breaks.</div>';
    if (d.breaks && d.breaks.janelas && d.breaks.janelas.length) {
        breaksHtml = '<div style="font-size:13px;color:#2C3E50;">Padrão esperado: ' +
            d.breaks.janelas.map(j => "18:" + String(j.inicioMin).padStart(2, "0") + "–18:" +
                String(j.fimMin).padStart(2, "0")).join(", ") + "</div>" +
            (d.breaks.alerta ? '<div style="font-size:12px;color:#7B7D7D;margin-top:4px;">' +
                d.breaks.alerta + '</div>' : "");
    }
    html += card("6 · Breaks Globo", breaksHtml);

    // 7. Clubes
    const linhasClubes = [];
    for (const c of CFG.CLUBES_FOCO) {
        const info = (d.clubes || {})[c];
        if (!info) continue;
        let txt = "—";
        if (info.resultado) {
            txt = (info.resultado === "V" ? "Venceu" : info.resultado === "D" ? "Perdeu" : "Empatou") +
                " " + info.quando + (info.classico ? " (clássico)" : "") +
                (info.competicao ? " · " + info.competicao : "");
        }
        if (info.proximoJogo) txt += (txt !== "—" ? " · " : "") + "Joga " + info.proximoJogo;
        linhasClubes.push([formatTemaDisplay_(c), txt]);
    }
    html += card("7 · Calendário dos clubes", linhasClubes.length ? kv(linhasClubes) :
        '<div style="font-size:13px;color:#7B7D7D;">Sem dados — verificar API-Football (menu Configurações → Testar API).</div>');

    // 8. "Próximo programa" removido: as recomendações automáticas de pauta
    // (apostar em time X / evitar time Y) saíam de ranking histórico sem
    // contexto de calendário — em período sem jogos viravam sugestão sem
    // fundamento. Reintroduzir apenas se condicionado a jogo confirmado.

    html += '<div style="font-size:11px;color:#95A5A6;margin-top:8px;">Radar Gazeta Esportiva — gerado automaticamente às ' +
        Utilities.formatDate(new Date(), CFG.TIMEZONE, "HH:mm") + '</div></div>';
    return html;
}

function fmtNumBRGE_(v) {
    return (typeof v === "number" && isFinite(v)) ? v.toFixed(2).replace(".", ",") : "—";
}
