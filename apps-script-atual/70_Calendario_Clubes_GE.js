// =============================================================================
// 17_CalendarioClubes_GE.js — Calendário de jogos + contexto para Calibragem
// =============================================================================
//
// ABA "Clubes":
//   Puxa jogos de CORINTHIANS, PALMEIRAS, SÃO PAULO, SANTOS e FLAMENGO
//   via API-Football (já configurada no CFG). Atualiza automaticamente
//   resultados e próximos jogos.
//
// ABA "Contexto Jogo":
//   Para cada programa no histórico, identifica o último jogo de cada
//   clube que aparece como tema, o resultado, e a distância em dias.
//   Permite análise: "CORINTHIANS pós-vitória rende X vs pós-derrota Y?"
//
// TRIGGER: rodar diariamente às 10h via Scheduling
// =============================================================================

// Clubes monitorados e seus IDs na API-Football (vindos do CFG.TEAM_IDS)
var CLUBES_MONITORADOS = ["CORINTHIANS", "PALMEIRAS", "SÃO PAULO", "SANTOS", "FLAMENGO"];

// Mapeamento de competições para nível de importância (1–5)
var IMPORTANCIA_COMPETICAO = {
  "Serie A": 2,
  "Brasileirao Serie A": 2,
  "Copa Do Brasil": 3,
  "Copa Libertadores": 3,
  "Recopa Sudamericana": 3,
  "Paulista - A1": 1,
  "Club Friendlies": 0,
  "Copa Sudamericana": 3
};

// Ajustes de importância por fase
var IMPORTANCIA_FASE = {
  "Final": 5,
  "Semi-finals": 4,
  "Quarter-finals": 4,
  "Round of 16": 3,
  "Group Stage": 2,
  "3rd Round": 2,
  "2nd Round": 1,
  "1st Round": 1
};

// =============================================================================
// ENTRADA PÚBLICA (menu)
// =============================================================================

function atualizarCalendarioClubes() {
  var ui = SpreadsheetApp.getUi();
  try {
    var stats = atualizarCalendarioClubes_();
    ui.alert(
      "✅ Calendário Atualizado",
      "Jogos atualizados via API-Football.\n\n" +
      "📊 " + stats.total + " jogos no calendário\n" +
      "✅ " + stats.realizados + " realizados\n" +
      "📅 " + stats.proximos + " próximos/agendados\n\n" +
      "Clubes: " + CLUBES_MONITORADOS.join(", "),
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("❌ Erro", "Erro ao atualizar calendário:\n" + err.message, ui.ButtonSet.OK);
  }
}

function gerarAnaliseContexto() {
  var ui = SpreadsheetApp.getUi();
  try {
    var stats = gerarAnaliseContexto_();
    ui.alert(
      "✅ Análise de Contexto Gerada",
      "Aba 'Contexto Jogo' atualizada.\n\n" +
      "📊 " + stats.programas + " programas analisados\n" +
      "🔗 " + stats.vinculacoes + " vínculos tema↔jogo encontrados",
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("❌ Erro", "Erro ao gerar contexto:\n" + err.message, ui.ButtonSet.OK);
  }
}

// =============================================================================
// BUSCAR JOGOS VIA API-FOOTBALL
// =============================================================================

function atualizarCalendarioClubes_() {
  var ss = SpreadsheetApp.getActive();
  var apiKey = PropertiesService.getScriptProperties().getProperty("API_FOOTBALL_KEY");
  if (!apiKey) throw Error("API_FOOTBALL_KEY não configurada nas propriedades do script.\n\nVá em: Configurações do Projeto → Propriedades do Script → Adicionar:\nChave: API_FOOTBALL_KEY\nValor: sua chave da API-Football");

  // Criar/limpar aba
  var tabName = "Clubes";
  var sheet = ss.getSheetByName(tabName);
  if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet(tabName); }

  // Cabeçalhos
  var headers = ["DATA", "CLUBE", "ADVERSÁRIO", "COMPETIÇÃO", "FASE", "MANDO",
                 "PLACAR", "GOLS PRÓ", "GOLS CONTRA", "RESULTADO", "STATUS",
                 "IMPORTÂNCIA", "DIAS P/ GE"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold").setBackground("#2C3E50").setFontColor("white").setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  var allRows = [];
  var season = 2026; // Temporada 2026

  for (var c = 0; c < CLUBES_MONITORADOS.length; c++) {
    var clubName = CLUBES_MONITORADOS[c];
    var teamId = CFG.TEAM_IDS[clubName];
    if (!teamId) {
      Logger.log("TEAM_ID não encontrado para " + clubName);
      continue;
    }

    // Buscar fixtures da temporada
    var url = CFG.API_FOOTBALL_ENDPOINT + "/fixtures?team=" + teamId + "&season=" + season;
    var options = {
      method: "get",
      headers: { "x-apisports-key": apiKey },
      muteHttpExceptions: true
    };

    try {
      var response = UrlFetchApp.fetch(url, options);
      var json = JSON.parse(response.getContentText());

      if (!json.response || json.response.length === 0) {
        Logger.log("Sem jogos para " + clubName + " na temporada " + season);
        continue;
      }

      for (var j = 0; j < json.response.length; j++) {
        var fixture = json.response[j];
        var fDate = new Date(fixture.fixture.date);
        var league = fixture.league.name || "";
        var round = fixture.league.round || "";

        var isHome = fixture.teams.home.id === teamId;
        var adversario = isHome ? fixture.teams.away.name : fixture.teams.home.name;
        var mando = isHome ? "CASA" : "FORA";

        var goalsHome = fixture.goals.home;
        var goalsAway = fixture.goals.away;
        var golsPro = isHome ? goalsHome : goalsAway;
        var golsContra = isHome ? goalsAway : goalsHome;

        var status = fixture.fixture.status.short; // FT, NS, TBD, etc.
        var statusLabel = (status === "FT" || status === "AET" || status === "PEN") ? "REALIZADO" :
                          (status === "NS" || status === "TBD") ? "AGENDADO" : status;

        var placar = "";
        var resultado = "";
        if (statusLabel === "REALIZADO" && golsPro !== null && golsContra !== null) {
          placar = golsPro + "×" + golsContra;
          resultado = golsPro > golsContra ? "V" : (golsPro < golsContra ? "D" : "E");
        }

        // Calcular importância
        var imp = IMPORTANCIA_COMPETICAO[league] || 2;
        // Ajustar por fase (mata-mata)
        for (var fase in IMPORTANCIA_FASE) {
          if (round.indexOf(fase) >= 0) {
            imp = Math.max(imp, IMPORTANCIA_FASE[fase]);
            break;
          }
        }
        // Clássicos paulistas = +1
        var classicos = {
          "CORINTHIANS": ["Palmeiras", "São Paulo", "Santos"],
          "PALMEIRAS": ["Corinthians", "São Paulo", "Santos"],
          "SÃO PAULO": ["Corinthians", "Palmeiras", "Santos"],
          "SANTOS": ["Corinthians", "Palmeiras", "São Paulo"],
          "FLAMENGO": ["Fluminense", "Vasco", "Botafogo"]
        };
        if (classicos[clubName]) {
          for (var cl = 0; cl < classicos[clubName].length; cl++) {
            if (adversario.indexOf(classicos[clubName][cl]) >= 0) {
              imp = Math.min(5, imp + 1);
              break;
            }
          }
        }

        allRows.push([
          fDate, clubName, adversario, league, round, mando,
          placar, golsPro !== null ? golsPro : "", golsContra !== null ? golsContra : "",
          resultado, statusLabel, imp, ""
        ]);
      }

      Utilities.sleep(500); // Rate limit
    } catch (err) {
      Logger.log("Erro API para " + clubName + ": " + err.message);
    }
  }

  // Ordenar por data
  allRows.sort(function(a, b) { return a[0] - b[0]; });

  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);

    // Formatação
    sheet.getRange(2, 1, allRows.length, 1).setNumberFormat("dd/MM/yyyy");
    sheet.getRange(2, 8, allRows.length, 2).setNumberFormat("0");
    sheet.getRange(2, 12, allRows.length, 1).setNumberFormat("0");

    // Cores alternadas por clube
    var colors = {
      "CORINTHIANS": "#F2F2F2", "PALMEIRAS": "#F0FFF0",
      "SÃO PAULO": "#FFF5F5", "SANTOS": "#F5F5FF", "FLAMENGO": "#FFF8F0"
    };
    for (var r = 0; r < allRows.length; r++) {
      var bg = colors[allRows[r][1]] || "#FFFFFF";
      sheet.getRange(r + 2, 1, 1, headers.length).setBackground(bg);
    }

    // Formatação condicional: V=verde, D=vermelho, E=amarelo
    var rngRes = sheet.getRange(2, 10, allRows.length, 1);
    var rules = [];
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("V").setBackground("#D5F5E3").setFontColor("#1E8449").setBold(true).setRanges([rngRes]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("D").setBackground("#FADBD8").setFontColor("#922B21").setBold(true).setRanges([rngRes]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("E").setBackground("#FEF9E7").setFontColor("#7D6608").setBold(true).setRanges([rngRes]).build());
    sheet.setConditionalFormatRules(rules);
  }

  // Larguras
  [95, 130, 160, 150, 120, 70, 70, 60, 60, 70, 90, 85, 80].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  SpreadsheetApp.flush();

  var realizados = allRows.filter(function(r) { return r[10] === "REALIZADO"; }).length;
  var proximos = allRows.filter(function(r) { return r[10] !== "REALIZADO"; }).length;

  return { total: allRows.length, realizados: realizados, proximos: proximos };
}

// =============================================================================
// ANÁLISE DE CONTEXTO: vincular programas históricos a jogos
// =============================================================================

function gerarAnaliseContexto_() {
  var ss = SpreadsheetApp.getActive();
  var hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
  var calSheet = ss.getSheetByName("Clubes");
  if (!hist) throw Error("Aba GE - Histórico não encontrada");
  if (!calSheet) throw Error("Aba Calendário Clubes não encontrada. Rode 'Atualizar Calendário' primeiro.");

  // Ler calendário
  var calData = calSheet.getRange(2, 1, calSheet.getLastRow() - 1, 12).getValues();
  var jogos = []; // {date, clube, adversario, competicao, mando, resultado, importancia}
  for (var i = 0; i < calData.length; i++) {
    if (!calData[i][0] || !(calData[i][0] instanceof Date)) continue;
    jogos.push({
      date: calData[i][0],
      clube: calData[i][1],
      adversario: calData[i][2],
      competicao: calData[i][3],
      fase: calData[i][4],
      mando: calData[i][5],
      placar: calData[i][6],
      resultado: calData[i][9],
      status: calData[i][10],
      importancia: calData[i][11]
    });
  }
  // A busca de "último antes"/"próximo depois" pressupõe ordem crescente de
  // data — a aba pode estar em qualquer ordem, então ordena aqui.
  jogos.sort(function(a, b) { return a.date - b.date; });

  // Ler programas do histórico (datas + temas + audiências)
  var audCols = getAudColsGE_(hist);
  if (audCols.length === 0) throw Error("Sem programas no histórico");

  // Para cada programa, ler data + temas + audiência por minuto
  var programas = [];
  for (var p = 0; p < audCols.length; p++) {
    var col = audCols[p];
    var dateVal = hist.getRange(1, col).getValue();
    var dDate = dateVal instanceof Date ? dateVal : parseDateStr_(dateVal);
    if (!dDate) continue;

    var audData = hist.getRange(3, col, 60, 1).getValues();
    var temaData = hist.getRange(3, col + 2, 60, 1).getValues();

    // Agrupar audiência por tema
    var temaPerfMap = {};
    for (var r = 0; r < 60; r++) {
      var tema = temaData[r][0];
      var aud = audData[r][0];
      if (!tema || typeof tema !== "string") continue;
      var t = tema.trim().toUpperCase();
      if (typeof aud !== "number" || aud <= 0) continue;
      if (!temaPerfMap[t]) temaPerfMap[t] = { soma: 0, cnt: 0, minutos: [] };
      temaPerfMap[t].soma += aud;
      temaPerfMap[t].cnt++;
      temaPerfMap[t].minutos.push(r);
    }

    var mediaGeral = 0;
    var audArr = audData.map(function(r) { return r[0]; }).filter(function(v) { return typeof v === "number" && v > 0; });
    if (audArr.length > 0) mediaGeral = audArr.reduce(function(s, v) { return s + v; }, 0) / audArr.length;

    programas.push({
      date: dDate,
      col: col,
      mediaGeral: mediaGeral,
      temaPerf: temaPerfMap
    });
  }

  // Criar aba de análise
  var tabName = "Contexto Jogo";
  var dest = ss.getSheetByName(tabName);
  if (dest) { dest.clear(); } else { dest = ss.insertSheet(tabName); }

  var headers = [
    "DATA PROGRAMA", "CLUBE/TEMA", "MÉDIA AUD. TEMA", "QTD MINUTOS",
    "ÚLTIMO JOGO", "ADVERSÁRIO", "RESULTADO", "PLACAR", "COMPETIÇÃO",
    "DIAS DESDE JOGO", "IMPORTÂNCIA", "MANDO",
    "PRÓXIMO JOGO", "ADV. PRÓXIMO", "COMP. PRÓXIMO", "DIAS P/ PRÓXIMO"
  ];
  dest.getRange(1, 1, 1, headers.length).setValues([headers]);
  dest.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold").setBackground("#1A5276").setFontColor("white").setHorizontalAlignment("center");
  dest.setFrozenRows(1);

  var resultRows = [];
  var vinculacoes = 0;

  for (var p = 0; p < programas.length; p++) {
    var prog = programas[p];
    var progDate = new Date(prog.date.getFullYear(), prog.date.getMonth(), prog.date.getDate());

    for (var tema in prog.temaPerf) {
      // Só clubes monitorados
      if (CLUBES_MONITORADOS.indexOf(tema) === -1) continue;

      var perf = prog.temaPerf[tema];
      var mediaT = perf.soma / perf.cnt;

      // Buscar último jogo ANTES desse programa
      var ultimoJogo = null;
      var diasDesde = "";
      for (var j = jogos.length - 1; j >= 0; j--) {
        if (jogos[j].clube === tema && jogos[j].status === "REALIZADO") {
          var jogoDate = new Date(jogos[j].date.getFullYear(), jogos[j].date.getMonth(), jogos[j].date.getDate());
          if (jogoDate < progDate) {
            ultimoJogo = jogos[j];
            diasDesde = Math.round((progDate - jogoDate) / (1000 * 60 * 60 * 24));
            break;
          }
        }
      }

      // Buscar próximo jogo APÓS esse programa
      var proximoJogo = null;
      var diasPara = "";
      for (var j = 0; j < jogos.length; j++) {
        if (jogos[j].clube === tema) {
          var jogoDate2 = new Date(jogos[j].date.getFullYear(), jogos[j].date.getMonth(), jogos[j].date.getDate());
          if (jogoDate2 > progDate) {
            proximoJogo = jogos[j];
            diasPara = Math.round((jogoDate2 - progDate) / (1000 * 60 * 60 * 24));
            break;
          }
        }
      }

      resultRows.push([
        prog.date,
        tema,
        mediaT,
        perf.cnt,
        ultimoJogo ? ultimoJogo.date : "",
        ultimoJogo ? ultimoJogo.adversario : "",
        ultimoJogo ? ultimoJogo.resultado : "",
        ultimoJogo ? ultimoJogo.placar : "",
        ultimoJogo ? ultimoJogo.competicao : "",
        diasDesde,
        ultimoJogo ? ultimoJogo.importancia : "",
        ultimoJogo ? ultimoJogo.mando : "",
        proximoJogo ? proximoJogo.date : "",
        proximoJogo ? proximoJogo.adversario : "",
        proximoJogo ? proximoJogo.competicao : "",
        diasPara
      ]);
      vinculacoes++;
    }
  }

  // Ordenar por data programa desc, depois clube
  resultRows.sort(function(a, b) {
    var d = b[0] - a[0];
    return d !== 0 ? d : a[1].localeCompare(b[1]);
  });

  if (resultRows.length > 0) {
    dest.getRange(2, 1, resultRows.length, headers.length).setValues(resultRows);
    dest.getRange(2, 1, resultRows.length, 1).setNumberFormat("dd/MM/yyyy");
    dest.getRange(2, 3, resultRows.length, 1).setNumberFormat("0.00");
    dest.getRange(2, 5, resultRows.length, 1).setNumberFormat("dd/MM/yyyy");
    dest.getRange(2, 13, resultRows.length, 1).setNumberFormat("dd/MM/yyyy");

    // Formatação condicional no resultado
    var rngRes = dest.getRange(2, 7, resultRows.length, 1);
    var rules = [];
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("V").setBackground("#D5F5E3").setFontColor("#1E8449").setBold(true).setRanges([rngRes]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("D").setBackground("#FADBD8").setFontColor("#922B21").setBold(true).setRanges([rngRes]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("E").setBackground("#FEF9E7").setFontColor("#7D6608").setBold(true).setRanges([rngRes]).build());
    dest.setConditionalFormatRules(rules);
  }

  // Larguras
  [95, 130, 100, 70, 95, 140, 60, 60, 140, 80, 80, 70, 95, 140, 140, 80].forEach(function(w, i) {
    dest.setColumnWidth(i + 1, w);
  });

  // ═══════════════════════════════════════════════════════════════════════
  // TABELA RESUMO: média por clube × resultado
  // ═══════════════════════════════════════════════════════════════════════
  var resumoStartRow = resultRows.length + 4;
  dest.getRange(resumoStartRow, 1).setValue("📊 RESUMO: MÉDIA POR CLUBE × RESULTADO");
  dest.getRange(resumoStartRow, 1, 1, 6).setBackground("#8E44AD").setFontColor("white").setFontWeight("bold");

  var resumoHeaders = ["CLUBE", "PÓS-VITÓRIA", "QTD V", "PÓS-DERROTA", "QTD D", "PÓS-EMPATE", "QTD E", "DIFERENÇA V-D"];
  dest.getRange(resumoStartRow + 1, 1, 1, resumoHeaders.length).setValues([resumoHeaders]);
  dest.getRange(resumoStartRow + 1, 1, 1, resumoHeaders.length)
    .setFontWeight("bold").setBackground("#D7BDE2").setHorizontalAlignment("center");

  var clubeStats = {};
  for (var r = 0; r < resultRows.length; r++) {
    var clube = resultRows[r][1];
    var mediaAud = resultRows[r][2];
    var resultado = resultRows[r][6];
    if (!resultado || typeof mediaAud !== "number") continue;

    if (!clubeStats[clube]) clubeStats[clube] = { V: [], D: [], E: [] };
    if (clubeStats[clube][resultado]) clubeStats[clube][resultado].push(mediaAud);
  }

  var resumoRow = resumoStartRow + 2;
  for (var c = 0; c < CLUBES_MONITORADOS.length; c++) {
    var cl = CLUBES_MONITORADOS[c];
    var st = clubeStats[cl] || { V: [], D: [], E: [] };
    var avgV = st.V.length > 0 ? st.V.reduce(function(s, v) { return s + v; }, 0) / st.V.length : 0;
    var avgD = st.D.length > 0 ? st.D.reduce(function(s, v) { return s + v; }, 0) / st.D.length : 0;
    var avgE = st.E.length > 0 ? st.E.reduce(function(s, v) { return s + v; }, 0) / st.E.length : 0;
    var diff = avgV - avgD;

    dest.getRange(resumoRow, 1, 1, 8).setValues([[cl, avgV, st.V.length, avgD, st.D.length, avgE, st.E.length, diff]]);
    dest.getRange(resumoRow, 2, 1, 1).setNumberFormat("0.00");
    dest.getRange(resumoRow, 4, 1, 1).setNumberFormat("0.00");
    dest.getRange(resumoRow, 6, 1, 1).setNumberFormat("0.00");
    dest.getRange(resumoRow, 8, 1, 1).setNumberFormat("+0.00;-0.00;0.00");
    resumoRow++;
  }

  SpreadsheetApp.flush();
  return { programas: programas.length, vinculacoes: vinculacoes };
}