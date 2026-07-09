// =============================================================================
// 90_Relatorio_PDF_GE.js — PDF de uma página (paisagem) da audiência do dia
// =============================================================================
// Monta a aba PDF a partir da Hoje (colunas A–G até a linha 65 + leitura),
// insere o mesmo gráfico do cockpit e exporta em paisagem para o Drive.

function gerarPDFAudienciaGE() {
    const ss = SpreadsheetApp.getActive();
    const hojeSheet = ss.getSheetByName(CFG.SHEET_HOJE);
    if (!hojeSheet) throw Error("Aba Hoje não encontrada");
    const pdfSheet = ensureAbaGE_(CFG.SHEET_PDF);
    const dataStr = Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy");

    // --- Layout da aba PDF ---
    pdfSheet.clear();
    for (const ch of pdfSheet.getCharts()) pdfSheet.removeChart(ch);
    if (pdfSheet.getMaxRows() < 80) pdfSheet.insertRowsAfter(pdfSheet.getMaxRows(), 80 - pdfSheet.getMaxRows());
    if (pdfSheet.getMaxColumns() < 20) pdfSheet.insertColumnsAfter(pdfSheet.getMaxColumns(), 20 - pdfSheet.getMaxColumns());

    pdfSheet.getRange(1, 1).setValue("GAZETA ESPORTIVA — AUDIÊNCIA MINUTO A MINUTO — " + dataStr)
        .setFontWeight("bold").setFontSize(14);

    // Tabela A2:G65 copiada por valores da Hoje (A2:G65 = headers + 60 min + média)
    const origem = hojeSheet.getRange(2, 1, 62, 7).getValues();
    pdfSheet.getRange(2, 1, 62, 7).setValues(origem);
    pdfSheet.getRange(2, 1, 1, 7).setFontWeight("bold").setBackground("#34495E").setFontColor("#FFFFFF");
    pdfSheet.getRange(3, 2, 61, 5).setNumberFormat("0.00");
    pdfSheet.getRange(63, 1, 1, 7).setFontWeight("bold").setBackground("#F5B7B1");

    // Leitura do dia (lado direito)
    const leitura = montarLeituraPDFGE_(hojeSheet);
    pdfSheet.getRange(2, 9, leitura.length, 2).setValues(leitura);
    pdfSheet.getRange(2, 9).setFontWeight("bold");
    pdfSheet.getRange(3, 10, leitura.length - 1, 1).setNumberFormat("@");

    // Séries auxiliares do gráfico copiadas para R/S (fora da área impressa,
    // que vai até a coluna P). Com TODOS os dados na própria aba PDF, o
    // gráfico sobrevive ao copyTo para a planilha temporária de export.
    pdfSheet.getRange(2, 18, 61, 1).setValues(hojeSheet.getRange(2, 20, 61, 1).getValues()); // R = ÚLT. PROG
    pdfSheet.getRange(2, 19, 61, 1).setValues(hojeSheet.getRange(2, 21, 61, 1).getValues()); // S = MÉDIA HIST.

    // Gráfico (mesmas séries do cockpit), com ranges da própria aba PDF
    const chart = pdfSheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(pdfSheet.getRange(2, 1, 61, 1))  // A = HORA
        .addRange(pdfSheet.getRange(2, 19, 61, 1)) // S = MÉDIA HIST.
        .addRange(pdfSheet.getRange(2, 2, 61, 1))  // B = MÉDIA ESPERADA
        .addRange(pdfSheet.getRange(2, 18, 61, 1)) // R = ÚLT. PROG
        .addRange(pdfSheet.getRange(2, 6, 61, 1))  // F = AUD FINAL
        .addRange(pdfSheet.getRange(2, 3, 61, 1))  // C = PREVISÃO
        .setNumHeaders(1)
        .setOption("title", "Hoje x Contexto x Previsão — " + dataStr)
        .setOption("useFirstColumnAsDomain", true)
        .setOption("legend", { position: "bottom" })
        .setOption("series", {
            0: { color: "#C0392B", lineWidth: 2 },
            1: { color: "#95A5A6", lineDashStyle: [4, 4] },
            2: { color: "#F5B041" },
            3: { color: "#2E86C1", lineWidth: 3 },
            4: { color: "#27AE60", lineDashStyle: [2, 2] }
        })
        .setOption("width", 620).setOption("height", 380)
        .setPosition(12, 9, 0, 0)
        .build();
    pdfSheet.insertChart(chart);

    pdfSheet.getRange(70, 1).setValue("Gerado automaticamente pelo Radar Gazeta Esportiva em " +
        Utilities.formatDate(new Date(), CFG.TIMEZONE, "dd/MM/yyyy HH:mm"))
        .setFontColor("#95A5A6").setFontSize(9);

    SpreadsheetApp.flush();
    Utilities.sleep(1500); // gráfico precisa renderizar antes do export

    // --- Export para PDF (paisagem, 1 página) via planilha TEMPORÁRIA ---
    // Exportar direto da Radar falhava com HTTP 500: o endpoint de export
    // renderiza o documento inteiro (Histórico gigante + centenas de
    // fórmulas). Copiamos só a aba PDF para um arquivo mínimo, exportamos
    // dele e jogamos o temporário no lixo.
    const temp = SpreadsheetApp.create("tmp_pdf_ge_" +
        Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyyMMdd_HHmmss"));
    let blob;
    try {
        const copia = pdfSheet.copyTo(temp).setName("PDF");
        temp.deleteSheet(temp.getSheets()[0]); // remove a "Página1" vazia
        SpreadsheetApp.flush();
        Utilities.sleep(800); // deixa o gráfico copiado assentar

        const base = "https://docs.google.com/spreadsheets/d/" + temp.getId() + "/export" +
            "?format=pdf&gid=" + copia.getSheetId() +
            "&size=A4&portrait=false&scale=4" +
            "&gridlines=false&printtitle=false&sheetnames=false&pagenum=false" +
            "&top_margin=0.30&bottom_margin=0.30&left_margin=0.30&right_margin=0.30";
        blob = exportarPdfComRetryGE_([
            base + "&r1=0&c1=0&r2=70&c2=15", // só o bloco usado (R/S ficam fora)
            base                              // fallback: aba inteira
        ]).setName("Audiencia_GE_" +
            Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyy-MM-dd") + ".pdf");
    } finally {
        try { DriveApp.getFileById(temp.getId()).setTrashed(true); } catch (eTrash) { }
    }

    // Salva na mesma pasta da planilha (ou raiz)
    let arquivo;
    try {
        const pasta = DriveApp.getFileById(ss.getId()).getParents();
        arquivo = pasta.hasNext() ? pasta.next().createFile(blob) : DriveApp.createFile(blob);
    } catch (e) {
        arquivo = DriveApp.createFile(blob);
    }
    logGE_("PDF", "PDF gerado: " + arquivo.getUrl());
    return arquivo;
}

// O endpoint de export do Docs falha com 500 esporádico (recalculo em curso,
// gráfico ainda renderizando). Tenta cada URL até 3x com backoff antes de desistir.
function exportarPdfComRetryGE_(urls) {
    const token = ScriptApp.getOAuthToken();
    let ultimoErro = "";
    for (const url of urls) {
        for (let tentativa = 1; tentativa <= 3; tentativa++) {
            const resp = UrlFetchApp.fetch(url, {
                headers: { Authorization: "Bearer " + token },
                muteHttpExceptions: true
            });
            const code = resp.getResponseCode();
            if (code === 200) return resp.getBlob();
            ultimoErro = "HTTP " + code;
            logGE_("PDF", "Export falhou (" + ultimoErro + "), tentativa " + tentativa + "/3");
            Utilities.sleep(1000 * Math.pow(2, tentativa - 1)); // 1s, 2s, 4s
        }
    }
    throw Error("Export do PDF falhou após todas as tentativas (" + ultimoErro + ").");
}

function gerarPDFAudienciaMenuGE() {
    const arquivo = gerarPDFAudienciaGE();
    const html = HtmlService.createHtmlOutput(
        '<p style="font-family:Arial;">✅ PDF gerado.</p>' +
        '<p style="font-family:Arial;"><a href="' + arquivo.getUrl() + '" target="_blank">Abrir PDF no Drive</a></p>')
        .setWidth(320).setHeight(120);
    SpreadsheetApp.getUi().showModalDialog(html, "PDF Audiência Minuto a Minuto");
}

function montarLeituraPDFGE_(hojeSheet) {
    const finais = hojeSheet.getRange(3, 6, 60, 1).getValues();
    const temas = hojeSheet.getRange(3, 9, 60, 1).getValues();
    const reguas = mediasContextuaisGE_();

    let soma = 0, n = 0, melhor = null, pior = null;
    for (let i = 0; i < 60; i++) {
        const v = finais[i][0];
        if (typeof v !== "number" || !isFinite(v)) continue;
        soma += v; n++;
        const hora = "18:" + String(i).padStart(2, "0");
        if (!melhor || v > melhor.v) melhor = { hora: hora, v: v, tema: String(temas[i][0] || "") };
        if (!pior || v < pior.v) pior = { hora: hora, v: v, tema: String(temas[i][0] || "") };
    }
    const media = n ? soma / n : null;
    const leitura = (typeof media === "number" && typeof reguas.mediaContexto === "number")
        ? leituraContextualGE_(media, reguas.mediaContexto) : "—";

    // Tema mais forte/fraco do dia pelo &/MIN médio
    const porTema = {};
    const baseB = hojeSheet.getRange(3, 2, 60, 1).getValues();
    for (let i = 0; i < 60; i++) {
        const v = finais[i][0], b = baseB[i][0];
        const t = String(temas[i][0] || "").trim().toUpperCase();
        if (!t || t.indexOf("MERCHAN") >= 0) continue;
        if (typeof v !== "number" || typeof b !== "number" || b <= 0) continue;
        if (!porTema[t]) porTema[t] = { soma: 0, n: 0 };
        porTema[t].soma += v / b; porTema[t].n++;
    }
    const rank = Object.keys(porTema).filter(t => porTema[t].n >= 2)
        .map(t => ({ t: t, idx: porTema[t].soma / porTema[t].n }));
    rank.sort((a, b) => b.idx - a.idx);

    return [
        ["LEITURA DO DIA", ""],
        ["Média final", fmtNumBRGE_(media)],
        ["Média contextual", fmtNumBRGE_(reguas.mediaContexto)],
        ["Leitura", leitura],
        ["Melhor minuto", melhor ? melhor.hora + " (" + fmtNumBRGE_(melhor.v) + ")" : "—"],
        ["Pior minuto", pior ? pior.hora + " (" + fmtNumBRGE_(pior.v) + ")" : "—"],
        ["Tema mais forte", rank.length ? formatTemaDisplay_(rank[0].t) : "—"],
        ["Tema mais fraco", rank.length > 1 ? formatTemaDisplay_(rank[rank.length - 1].t) : "—"]
    ];
}
