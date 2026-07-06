// =============================================================================
// 12_AlertaTema_GE.js — Alerta + Sugestão + LOG + Autonomia (GE)
// =============================================================================
// VERSÃO 3 — Corrigido para Google Sheets em PORTUGUÊS (SE, OU, E, PROCV, etc.)
// =============================================================================

var TIMES_ELEGIVEIS_GE = [
  "PALMEIRAS", "CORINTHIANS", "SÃO PAULO", "SANTOS",
  "FLAMENGO", "FLUMINENSE", "VASCO", "BOTAFOGO",
  "INTERNACIONAL", "GRÊMIO", "ATLÉTICO-MG", "CRUZEIRO",
  "BAHIA", "FORTALEZA", "ATHLETICO-PR", "BRAGANTINO",
  "SELEÇÃO", "PORTUGUESA"
];

// ─── ENTRADA PÚBLICA (menu) ─────────────────────────────────────────────────

function atualizarAlertaTemaGE() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(CFG.GE_OUT_HISTORICO);
  if (!sheet) return;
  atualizarRankingTemaCalibGE_();
  atualizarDuracaoBlocosCalibGE_();
  atualizarAlertaColunasGE_(sheet);
  SpreadsheetApp.getUi().alert(
    "✅ Alerta de Tema",
    "Colunas HD–HI + tabela de durações atualizadas.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ─── onEdit HOOK ─────────────────────────────────────────────────────────────

function onEditAlertaTemaGE(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var histName = (typeof CFG !== "undefined" && CFG.GE_OUT_HISTORICO)
    ? CFG.GE_OUT_HISTORICO : "Histórico";
  if (sheet.getName() !== histName) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();
  if ((col !== 4 && col !== 6) || row < 3 || row > 62) return;

  atualizarRankingTemaCalibGE_();
  atualizarAlertaColunasGE_(sheet);

  if (col === 4) {
    var alertStart = getAlertStartCol_(sheet);
    if (!alertStart) return;
    SpreadsheetApp.flush();
    Utilities.sleep(300);

    var numRows = e.range.getNumRows();
    for (var r = row; r < row + numRows && r <= 62; r++) {
      var sinal = sheet.getRange(r, alertStart).getDisplayValue();
      var sugestao = sheet.getRange(r, alertStart + 1).getDisplayValue();
      var hasSug = sinal.indexOf("ATENÇÃO") >= 0 || sinal.indexOf("TROCAR") >= 0 || sinal.indexOf("ESGOTANDO") >= 0;
      if (hasSug && sugestao) {
        var logAtual = sheet.getRange(r, alertStart + 3).getValue();
        if (!logAtual || logAtual === "") {
          sheet.getRange(r, alertStart + 3).setValue(sugestao);
        }
      }
    }
  }
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function getAlertStartCol_(sheet) {
  var audCols = getAudColsGE_(sheet);
  if (audCols.length === 0) return null;
  var calibStartCol = 4 + audCols.length * 4;
  var hasTema = !!SpreadsheetApp.getActive().getSheetByName("Tema Minuto");
  return hasTema ? calibStartCol + 11 : calibStartCol + 7;
}

// ─── TABELA DE DURAÇÃO DE BLOCOS (Calibragem AA–AF) ─────────────────────────

function atualizarDuracaoBlocosCalibGE_() {
  var ss = SpreadsheetApp.getActive();
  var hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
  var calib = ss.getSheetByName("Calibragem");
  if (!hist || !calib) return;

  var audCols = getAudColsGE_(hist);
  if (audCols.length < 2) return;

  var blocksByTema = {};
  var IGNORE = { "": 1, "NONE": 1 };

  for (var ci = 1; ci < audCols.length; ci++) {
    var temaCol = audCols[ci] + 2;
    var temas = hist.getRange(3, temaCol, 60, 1).getValues();
    var prevTema = "", blockDur = 0;
    for (var r = 0; r < 60; r++) {
      var t = temas[r][0];
      var ts = (t && typeof t === "string") ? t.trim().toUpperCase() : "";
      if (ts === prevTema && !IGNORE[ts]) {
        blockDur++;
      } else {
        if (prevTema && !IGNORE[prevTema] && blockDur > 0) {
          if (!blocksByTema[prevTema]) blocksByTema[prevTema] = [];
          blocksByTema[prevTema].push(blockDur);
        }
        prevTema = ts; blockDur = 1;
      }
    }
    if (prevTema && !IGNORE[prevTema] && blockDur > 0) {
      if (!blocksByTema[prevTema]) blocksByTema[prevTema] = [];
      blocksByTema[prevTema].push(blockDur);
    }
  }

  var startCol = 27;
  calib.getRange(1, startCol, 1, 6)
    .setValues([["TEMA", "N BLOCOS", "MEDIANA", "P75", "P90", "MAX"]])
    .setFontWeight("bold").setBackground("#1A5276").setFontColor("white").setHorizontalAlignment("center");

  var temas = Object.keys(blocksByTema).sort();
  var rows = [];
  for (var i = 0; i < temas.length; i++) {
    var t = temas[i];
    var blocks = blocksByTema[t].sort(function(a, b) { return a - b; });
    var n = blocks.length;
    if (n < 2) continue;
    rows.push([t, n, blocks[Math.floor(n/2)], blocks[Math.min(Math.floor(n*0.75),n-1)],
      blocks[Math.min(Math.floor(n*0.90),n-1)], blocks[n-1]]);
  }
  if (rows.length > 0) {
    calib.getRange(2, startCol, rows.length, 6).setValues(rows);
    calib.getRange(2, startCol+1, rows.length, 5).setNumberFormat("0");
  }
  var lastOcc = rows.length + 1;
  if (calib.getLastRow() > lastOcc) {
    var clr = Math.min(30, calib.getLastRow() - lastOcc);
    if (clr > 0) calib.getRange(lastOcc+1, startCol, clr, 6).clearContent();
  }
  calib.setColumnWidth(startCol, 140);
  for (var c = 1; c <= 5; c++) calib.setColumnWidth(startCol+c, 80);
}

// ─── RANKING DE TEMAS (Calibragem U–Y) ──────────────────────────────────────

function atualizarRankingTemaCalibGE_() {
  var ss = SpreadsheetApp.getActive();
  var hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
  var calib = ss.getSheetByName("Calibragem");
  if (!hist || !calib) return;

  var dadosHoje = hist.getRange(3, 4, 60, 3).getValues();
  var temasUsados1Meia = new Set();
  var perfReal1 = {};
  for (var r = 0; r < 60; r++) {
    var aud = dadosHoje[r][0], tema = dadosHoje[r][2];
    if (!tema || typeof tema !== "string") continue;
    var t = tema.trim().toUpperCase();
    if (r < 30) {
      temasUsados1Meia.add(t);
      if (TIMES_ELEGIVEIS_GE.indexOf(t) >= 0 && typeof aud === "number" && aud > 0) {
        if (!perfReal1[t]) perfReal1[t] = {soma:0,cnt:0};
        perfReal1[t].soma += aud; perfReal1[t].cnt++;
      }
    }
  }
  var temaSheet = ss.getSheetByName("Tema Minuto");
  var indiceHist = {};
  if (temaSheet) {
    var tData = temaSheet.getRange(2, 1, temaSheet.getLastRow()-1, 6).getValues();
    for (var i = 0; i < tData.length; i++) {
      var nome = tData[i][0], idx = tData[i][5], ult3 = tData[i][4];
      if (nome && typeof nome === "string") {
        var score = (typeof idx === "number" && idx > 0) ? idx : (typeof ult3 === "number" && ult3 > 0) ? ult3 : 0;
        indiceHist[nome.trim().toUpperCase()] = score;
      }
    }
  }
  var startCol = 21;
  calib.getRange(1, startCol, 1, 5)
    .setValues([["TIME","ÍND. HIST.","USADO 1ªM?","MÉDIA 1ªM","RANKING"]])
    .setFontWeight("bold").setBackground("#8E44AD").setFontColor("white").setHorizontalAlignment("center");
  var rows = [];
  for (var i = 0; i < TIMES_ELEGIVEIS_GE.length; i++) {
    var t = TIMES_ELEGIVEIS_GE[i];
    rows.push([t, indiceHist[t]||0, temasUsados1Meia.has(t)?1:0,
      (perfReal1[t]?perfReal1[t].soma/perfReal1[t].cnt:0), indiceHist[t]||0]);
  }
  rows.sort(function(a,b){return b[1]-a[1];});
  if (rows.length > 0) {
    calib.getRange(2,startCol,rows.length,5).setValues(rows);
    calib.getRange(2,startCol+1,rows.length,1).setNumberFormat("0.000");
    calib.getRange(2,startCol+2,rows.length,1).setNumberFormat("0");
    calib.getRange(2,startCol+3,rows.length,1).setNumberFormat("0.000");
    calib.getRange(2,startCol+4,rows.length,1).setNumberFormat("0.000");
  }
  calib.setColumnWidth(startCol,140);
  for (var c=1;c<=4;c++) calib.setColumnWidth(startCol+c,90);
}

// ─── COLUNAS HD–HI no GE - Histórico (FÓRMULAS EM PORTUGUÊS) ────────────────

function atualizarAlertaColunasGE_(sheet) {
  var audCols = getAudColsGE_(sheet);
  var numPrograms = audCols.length;
  if (numPrograms === 0) return;

  var calibStartCol = 4 + numPrograms * 4;
  var hasTema = !!SpreadsheetApp.getActive().getSheetByName("Tema Minuto");
  var alertStartCol = hasTema ? calibStartCol + 11 : calibStartCol + 7;

  var HDl = columnToLetter_(alertStartCol);
  var HEl = columnToLetter_(alertStartCol + 1);
  var HFl = columnToLetter_(alertStartCol + 2);
  var HGl = columnToLetter_(alertStartCol + 3);
  var HHl = columnToLetter_(alertStartCol + 4);
  var HIl = columnToLetter_(alertStartCol + 5);

  var GTl = columnToLetter_(calibStartCol + 1);
  var GWl = columnToLetter_(calibStartCol + 4);
  var HBl = hasTema ? columnToLetter_(calibStartCol + 7 + 2) : null;

  // Cabeçalhos
  sheet.getRange(1, alertStartCol, 1, 6)
    .setValues([["🚦 SINAL","💡 SUGESTÃO","📋 MOTIVO","📌 LOG SGEST.","✅ ESCOLHA","🎯 SEGUIU?"]])
    .setFontWeight("bold").setFontColor("white").setHorizontalAlignment("center");
  sheet.getRange(1, alertStartCol, 1, 3).setBackground("#8E44AD");
  sheet.getRange(1, alertStartCol+3, 1, 3).setBackground("#1A5276");
  sheet.getRange(2, alertStartCol, 1, 6)
    .setValues([["HD","HE","HF","HG","HH","HI"]])
    .setFontWeight("bold").setFontColor("white").setHorizontalAlignment("center");
  sheet.getRange(2, alertStartCol, 1, 3).setBackground("#9B59B6");
  sheet.getRange(2, alertStartCol+3, 1, 3).setBackground("#21618C");

  // Buscar sugestões
  var calib = SpreadsheetApp.getActive().getSheetByName("Calibragem");
  var melhorInedito = "—", melhor1Meia = "—";
  if (calib) {
    var maxR = Math.min(TIMES_ELEGIVEIS_GE.length, calib.getLastRow()-1);
    if (maxR > 0) {
      var rankData = calib.getRange(2,21,maxR,4).getValues();
      var bestI=0, bestM=0;
      for (var i=0;i<rankData.length;i++) {
        if (rankData[i][2]===0 && rankData[i][1]>bestI) {bestI=rankData[i][1]; melhorInedito=rankData[i][0];}
        if (rankData[i][3]>bestM) {bestM=rankData[i][3]; melhor1Meia=rankData[i][0];}
      }
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // FÓRMULAS EM PORTUGUÊS (SE, OU, E, SEERRO, PROCV, MÁXIMO, etc.)
  // ═══════════════════════════════════════════════════════════════
  var hdF=[], heF=[], hfF=[], hhF=[], hiF=[];

  for (var row = 3; row <= 62; row++) {
    var isPrimeiraMeia = (row <= 32);
    var sug = isPrimeiraMeia ? melhorInedito : melhor1Meia;

    // === HD: SINAL ===
    if (row <= 5) {
      if (HBl) {
        hdF.push([
          '=SE(' + GTl+row + '="";"";' +
          'SE(' + GTl+row + '<0,90;"🔴 TROCAR";' +
          'SE(E(F' + row + '<>"";' + HBl+row + '<=SEERRO(PROCV(F' + row + ';Calibragem!$AA$2:$AF$50;3;0);5));' +
          '"🟢 +" & (SEERRO(PROCV(F' + row + ';Calibragem!$AA$2:$AF$50;3;0);5)-' + HBl+row + ');' +
          '"🟢 OK")))'
        ]);
      } else {
        hdF.push(['=SE(' + GTl+row + '="";"";SE(' + GTl+row + '<0,90;"🔴 TROCAR";"🟢 OK"))']);
      }
    } else {
      if (HBl) {
        hdF.push([
          '=SE(' + GTl+row + '="";"";' +
          'SE(E(' + GTl+(row-1) + '<>"";' + GTl+(row-2) + '<>"";' + GTl+(row-3) + '<>"";' +
              GTl+(row-1) + '<0,90;' + GTl+(row-2) + '<0,90;' + GTl+(row-3) + '<0,90);"🔴 TROCAR";' +
          'SE(E(' + GTl+(row-1) + '<>"";' + GTl+(row-2) + '<>"";' +
              GTl+(row-1) + '<0,93;' + GTl+(row-2) + '<0,93);"🟡 ATENÇÃO";' +
          'SE(F' + row + '="";"🟢 OK";' +
          'SE(' + HBl+row + '>SEERRO(PROCV(F' + row + ';Calibragem!$AA$2:$AF$50;3;0);5);' +
            'SE(' + GWl+row + '<-0,03;"🟡 ESGOTANDO";' +
            'SE(' + HBl+row + '>SEERRO(PROCV(F' + row + ';Calibragem!$AA$2:$AF$50;4;0);7);"🟢 ONDA ⚡";"🟢 ONDA 🌊"));' +
          '"🟢 +" & MÁXIMO(0;SEERRO(PROCV(F' + row + ';Calibragem!$AA$2:$AF$50;3;0);5)-' + HBl+row + '))))))'
        ]);
      } else {
        hdF.push([
          '=SE(' + GTl+row + '="";"";' +
          'SE(E(' + GTl+(row-1) + '<>"";' + GTl+(row-2) + '<>"";' + GTl+(row-3) + '<>"";' +
              GTl+(row-1) + '<0,90;' + GTl+(row-2) + '<0,90;' + GTl+(row-3) + '<0,90);"🔴 TROCAR";' +
          'SE(E(' + GTl+(row-1) + '<>"";' + GTl+(row-2) + '<>"";' +
              GTl+(row-1) + '<0,93;' + GTl+(row-2) + '<0,93);"🟡 ATENÇÃO";"🟢 OK")))'
        ]);
      }
    }

    // === HE: SUGESTÃO ===
    heF.push([
      '=SE(OU(' + HDl+row + '="🔴 TROCAR";' + HDl+row + '="🟡 ATENÇÃO";' + HDl+row + '="🟡 ESGOTANDO");"' + sug + '";"")'
    ]);

    // === HF: MOTIVO ===
    if (isPrimeiraMeia) {
      hfF.push([
        '=SE(' + HDl+row + '="🔴 TROCAR";"Queda severa — trocar p/ inédito";' +
        'SE(' + HDl+row + '="🟡 ATENÇÃO";"Queda moderada — considerar inédito";' +
        'SE(' + HDl+row + '="🟡 ESGOTANDO";"Tema esgotando — preparar troca";"")))'
      ]);
    } else {
      hfF.push([
        '=SE(' + HDl+row + '="🔴 TROCAR";"Queda severa — voltar ao que deu certo";' +
        'SE(' + HDl+row + '="🟡 ATENÇÃO";"Queda moderada — repetir melhor 1ªM";' +
        'SE(' + HDl+row + '="🟡 ESGOTANDO";"Tema esgotando — voltar melhor 1ªM";"")))'
      ]);
    }

    // === HH: ESCOLHA ===
    hhF.push(['=SE(F' + row + '="";"";F' + row + ')']);

    // === HI: SEGUIU? ===
    hiF.push(['=SE(OU(' + HGl+row + '="";"' + HHl+row + '="");"—";SE(' + HGl+row + '=' + HHl+row + ';"✅ SIM";"❌ NÃO"))']);
  }

  batchSetFormulasSafe_(sheet.getRange(3, alertStartCol, 60, 1), hdF);
  batchSetFormulasSafe_(sheet.getRange(3, alertStartCol+1, 60, 1), heF);
  batchSetFormulasSafe_(sheet.getRange(3, alertStartCol+2, 60, 1), hfF);
  // HG (LOG) — NÃO sobrescrever
  batchSetFormulasSafe_(sheet.getRange(3, alertStartCol+4, 60, 1), hhF);
  batchSetFormulasSafe_(sheet.getRange(3, alertStartCol+5, 60, 1), hiF);

  [100,130,280,130,130,90].forEach(function(w,i){sheet.setColumnWidth(alertStartCol+i,w);});

  // Formatação condicional
  var rules = sheet.getConditionalFormatRules();
  var newRules = rules.filter(function(rule){
    return !rule.getRanges().some(function(r){return r.getColumn()>=alertStartCol && r.getColumn()<=alertStartCol+5;});
  });
  var rngA = sheet.getRange(3,alertStartCol,60,3);
  var rngL = sheet.getRange(3,alertStartCol+3,60,3);
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("TROCAR").setBackground("#FADBD8").setFontColor("#922B21").setBold(true).setRanges([rngA]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("ESGOTANDO").setBackground("#FEF9E7").setFontColor("#7D6608").setBold(true).setRanges([rngA]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("ATENÇÃO").setBackground("#FEF9E7").setFontColor("#7D6608").setBold(true).setRanges([rngA]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("ONDA").setBackground("#D5F5E3").setFontColor("#1E8449").setBold(true).setRanges([rngA]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("✅ SIM").setBackground("#D5F5E3").setFontColor("#1E8449").setRanges([rngL]).build());
  newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("❌ NÃO").setBackground("#FADBD8").setFontColor("#922B21").setRanges([rngL]).build());
  sheet.setConditionalFormatRules(newRules);
}