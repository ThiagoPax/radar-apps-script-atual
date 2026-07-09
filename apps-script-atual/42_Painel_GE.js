// =============================================================================
// 31_Painel_GE.gs — Painel de diagnóstico e decisão (aba própria "Painel")
// Lê B/D/F e PROJ V2/PROJ (por cabeçalho) da aba Histórico e escreve o painel
// SOMENTE na aba "Painel" (valores, sem fórmulas). Nao mexe em D, TY, UA:UD.
//
// ESTRUTURAL: o painel morava nas colunas fixas TD:TV do Histórico e fazia
// clearContent() dessa região a cada rebuild. O histórico cresce 4 colunas por
// dia e, ao atingir ~130 programas, o bloco de calibragem (e a cauda do
// histórico) entrou em TD:TV — o painel passou a apagar fórmulas e a escrever
// texto onde a predição esperava números (#VALUE!). Nenhum bloco do Histórico
// pode ter âncora de coluna fixa; por isso o painel vive em aba dedicada.
//
// Primeiro rode: painel31_testeZero_GE
// Depois rode:   painel31_reconstruir_TDTV_GE
// Opcional:      painel31_instalarGatilho_TDTV_GE
// =============================================================================

var PAINEL31_SHEET_NAME = 'Painel';
var PAINEL31_NCOLS = 19;

function painel31_testeZero_GE() {
  Logger.log('OK painel31: o arquivo compilou e executou.');
  return 'OK painel31: o arquivo compilou e executou.';
}

// Garante a aba "Painel" com espaço para cabeçalhos (1-2), minutos (3-62) e
// resumo (66+).
function painel31_getPainelSheet_GE_(ss) {
  var sh = ss.getSheetByName(PAINEL31_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PAINEL31_SHEET_NAME);
  }
  if (sh.getMaxRows() < 100) {
    sh.insertRowsAfter(sh.getMaxRows(), 100 - sh.getMaxRows());
  }
  if (sh.getMaxColumns() < PAINEL31_NCOLS) {
    sh.insertColumnsAfter(sh.getMaxColumns(), PAINEL31_NCOLS - sh.getMaxColumns());
  }
  return sh;
}

// Mapa tema → fator de desempenho RECENTE, da aba Tema Minuto:
// média dos últimos 3 dias (fallback 5, fallback todos) ÷ média geral
// ponderada, com amostra mínima de 30 min e clamp 0,85–1,15.
function painel31_fatoresTemaRecentes_GE_(ss) {
  var out = {};
  try {
    var nomeAba = (typeof CFG !== 'undefined' && CFG.GE_OUT_MEDIA_TEMA) ? CFG.GE_OUT_MEDIA_TEMA : 'Tema Minuto';
    var sh = ss.getSheetByName(nomeAba);
    if (!sh || sh.getLastRow() < 2) return out;
    var dados = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();

    var somaGeral = 0, pesoGeral = 0;
    for (var i = 0; i < dados.length; i++) {
      var mAll = Number(dados[i][2]), mins = Number(dados[i][1]) || 0;
      if (isFinite(mAll) && mAll > 0 && mins >= 30) {
        somaGeral += mAll * mins;
        pesoGeral += mins;
      }
    }
    var mediaGeral = pesoGeral > 0 ? somaGeral / pesoGeral : 0;
    if (!mediaGeral) return out;

    for (var j = 0; j < dados.length; j++) {
      var nome = String(dados[j][0] || '').trim().toUpperCase();
      var n = Number(dados[j][1]) || 0;
      if (!nome || n < 30) continue;
      var u3 = Number(dados[j][4]), u5 = Number(dados[j][3]), todos = Number(dados[j][2]);
      var mediaRec = (isFinite(u3) && u3 > 0) ? u3 :
        (isFinite(u5) && u5 > 0) ? u5 :
        (isFinite(todos) && todos > 0) ? todos : 0;
      if (!mediaRec) continue;
      out[nome] = painel31_clamp_GE_(mediaRec / mediaGeral, 0.85, 1.15);
    }
  } catch (e) { }
  return out;
}

function painel31_reconstruir_TDTV_GE() {
  var ss = painel31_getSpreadsheet_GE_();
  var sh = ss.getSheetByName('Histórico');
  if (!sh) {
    throw new Error('Aba GE - Histórico nao encontrada.');
  }
  var shPainel = painel31_getPainelSheet_GE_(ss);

  var colB = 2;
  var colD = 4;
  var colF = 6;
  var nCols = PAINEL31_NCOLS;

  // PROJ V2 / PROJ são localizadas por cabeçalho; sem fallback de letra fixa
  // (TY/UD): num histórico que cresce, uma letra fixa cedo ou tarde aponta
  // para dado errado. Sem a coluna, o painel degrada para D (audUsada).
  var colTY = painel31_acharCabecalho_GE_(sh, 'PROJ V2');
  var colUD = painel31_acharCabecalho_GE_(sh, 'PROJ');
  if (colUD === colTY) {
    colUD = -1;
  }

  var vazio60 = [];
  for (var z = 0; z < 60; z++) {
    vazio60.push(['']);
  }
  var valsB = sh.getRange(3, colB, 60, 1).getValues();
  var valsD = sh.getRange(3, colD, 60, 1).getValues();
  var valsF = sh.getRange(3, colF, 60, 1).getValues();
  var valsTY = colTY >= 1 ? sh.getRange(3, colTY, 60, 1).getValues() : vazio60;
  var valsUD = colUD >= 1 ? sh.getRange(3, colUD, 60, 1).getValues() : vazio60;
  var formulasD = sh.getRange(3, colD, 60, 1).getFormulas();

  var B = [];
  var D = [];
  var F = [];
  var TY = [];
  var UD = [];
  var isReal = [];
  var ratioReal = [];
  var lastReal = -1;

  for (var i = 0; i < 60; i++) {
    B[i] = painel31_num_GE_(valsB[i][0]);
    D[i] = painel31_num_GE_(valsD[i][0]);
    F[i] = painel31_text_GE_(valsF[i][0]);
    TY[i] = painel31_num_GE_(valsTY[i][0]);
    UD[i] = painel31_num_GE_(valsUD[i][0]);
    isReal[i] = painel31_isNumero_GE_(D[i]) && String(formulasD[i][0] || '') === '';
    if (isReal[i]) {
      lastReal = i;
    }
    ratioReal[i] = (isReal[i] && B[i] > 0) ? D[i] / B[i] : null;
  }

  var ratiosUsaveis = [];
  for (var r = 0; r <= lastReal; r++) {
    if (painel31_isNumero_GE_(ratioReal[r])) {
      ratiosUsaveis.push(ratioReal[r]);
    }
  }

  var fatorDiaGlobal = painel31_mediana_GE_(ratiosUsaveis);
  if (!painel31_isNumero_GE_(fatorDiaGlobal)) {
    fatorDiaGlobal = 1;
  }

  var finalPrev = painel31_mediaFinalPrevista_GE_(D, TY, lastReal);
  var fatoresTemaRecentes = painel31_fatoresTemaRecentes_GE_(ss);

  var headers1 = [[
    'BASE HIST.', 'DIA/HIST', 'FATOR DIA', 'RECENTE', 'INCLINACAO', 'RUPTURA', 'FATOR FINAL',
    'SINAL TEMA', 'TROCA TEMA', 'DUR. BLOCO', 'FATOR TEMA',
    'SINAL', 'SUGESTAO', 'MOTIVO', 'PROGNOSTICO', 'FOCO', 'RISCO', 'MEDIA 5', 'FINAL PREV.'
  ]];

  var headers2 = [[
    'B', 'D/B', 'DIA', 'ULTIMOS', 'TEND.', 'QUEBRA', 'COMBO',
    'UD/TY', 'MUDOU?', 'MIN.', 'TEMA',
    'STATUS', 'ACAO', 'EXPLICA', 'LEITURA', 'EDITORIAL', 'RISCO', 'PROX.5', 'FINAL'
  ]];

  var out = [];
  var blocoDur = 0;
  var temaAnterior = '';

  for (var m = 0; m < 60; m++) {
    var base = B[m];
    var audUsada = isReal[m] ? D[m] : (painel31_isNumero_GE_(TY[m]) ? TY[m] : D[m]);
    var ratio = (base > 0 && painel31_isNumero_GE_(audUsada)) ? audUsada / base : null;

    var recente = painel31_fatorRecente_GE_(ratioReal, m, fatorDiaGlobal);
    var inclinacao = painel31_inclinacao_GE_(ratioReal, m);
    var ruptura = painel31_ruptura_GE_(inclinacao);
    var fatorFinal = painel31_clamp_GE_(0.55 * fatorDiaGlobal + 0.35 * recente + 0.10 * (1 + inclinacao), 0.55, 1.55);

    // FATOR TEMA de verdade: índice recente do tema que está no ar (aba Tema
    // Minuto, últimos dias, amostra mínima). O cálculo antigo era
    // clamp(PROJ_legado/PROJ_V2) — uma razão entre projeções que ficava
    // cronicamente ≤0,96 e travava SINAL em "TEMA PESA", AÇÃO em "Encurtar
    // tema ou trocar" e RISCO em ALTO, independentemente do tema.
    var temaAtual = F[m];
    var fatorTema = 1;
    var chaveTema = String(temaAtual || '').trim().toUpperCase();
    if (fatoresTemaRecentes.hasOwnProperty(chaveTema)) {
      fatorTema = fatoresTemaRecentes[chaveTema];
    }
    if (m === 0 || temaAtual !== temaAnterior) {
      blocoDur = 1;
    } else {
      blocoDur++;
    }
    var trocou = (m > 0 && temaAtual !== temaAnterior) ? 'SIM' : 'NAO';
    temaAnterior = temaAtual;

    var sinal = painel31_sinal_GE_(ratio, inclinacao, fatorTema, isReal[m]);
    var sugestao = painel31_sugestao_GE_(sinal, fatorTema, temaAtual, blocoDur);
    var motivo = painel31_motivo_GE_(ratio, recente, inclinacao, fatorTema, isReal[m]);
    var prognostico = painel31_prognostico_GE_(TY, D, lastReal, m, finalPrev);
    var foco = painel31_foco_GE_(temaAtual, fatorTema, blocoDur);
    var risco = painel31_risco_GE_(ratio, inclinacao, fatorTema);
    var media5 = painel31_mediaProximos_GE_(TY, D, lastReal, m, 5);

    out.push([
      painel31_numOuVazio_GE_(base),
      painel31_numOuVazio_GE_(ratio),
      painel31_numOuVazio_GE_(fatorDiaGlobal),
      painel31_numOuVazio_GE_(recente),
      painel31_numOuVazio_GE_(inclinacao),
      ruptura,
      painel31_numOuVazio_GE_(fatorFinal),
      painel31_numOuVazio_GE_(fatorTema),
      trocou,
      blocoDur,
      painel31_resumoTema_GE_(fatorTema),
      sinal,
      sugestao,
      motivo,
      prognostico,
      foco,
      risco,
      painel31_numOuVazio_GE_(media5),
      painel31_numOuVazio_GE_(finalPrev)
    ]);
  }

  // Reescreve somente a aba Painel; o Histórico não é tocado.
  shPainel.getRange(1, 1, 100, nCols).clearContent();
  shPainel.getRange(1, 1, 1, nCols).setValues(headers1);
  shPainel.getRange(2, 1, 1, nCols).setValues(headers2);
  shPainel.getRange(3, 1, 60, nCols).setValues(out);
  shPainel.setFrozenRows(2);

  painel31_escreverResumo_GE_(shPainel, 1, nCols, lastReal, fatorDiaGlobal, finalPrev, TY, D, F);
  SpreadsheetApp.flush();

  return 'OK: aba Painel reconstruida. Ultimo real: linha ' + (lastReal >= 0 ? (lastReal + 3) : 'nenhuma') + '. Final previsto: ' + painel31_fmtNum_GE_(finalPrev) + '.';
}

function painel31_instalarGatilho_TDTV_GE() {
  var ss = painel31_getSpreadsheet_GE_();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction && triggers[i].getHandlerFunction() === 'painel31_onEdit_TDTV_GE') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('painel31_onEdit_TDTV_GE').forSpreadsheet(ss).onEdit().create();
  return 'OK: gatilho do painel TD:TV instalado.';
}

function painel31_onEdit_TDTV_GE(e) {
  try {
    if (!e || !e.range) {
      return;
    }
    var sh = e.range.getSheet();
    if (!sh || sh.getName() !== 'Histórico') {
      return;
    }
    var row = e.range.getRow();
    var col = e.range.getColumn();
    if (row < 3 || row > 62) {
      return;
    }
    if (col !== 4 && col !== 6) {
      return;
    }

    // Se a V2 existir, recalcula antes para o painel ler a TY ja atualizada.
    try {
      if (typeof recalcularPrevisaoAoVivoGE === 'function') {
        recalcularPrevisaoAoVivoGE();
        Utilities.sleep(2500);
      }
    } catch (errV2) {
      Logger.log('Painel31: V2 nao recalculou: ' + errV2);
    }

    painel31_reconstruir_TDTV_GE();
  } catch (err) {
    Logger.log('Painel31 onEdit erro: ' + err);
  }
}

function painel31_prepararNovoDia_GE() {
  try {
    if (typeof instalarPrevisaoExternaGE === 'function') {
      instalarPrevisaoExternaGE();
    }
  } catch (errInst) {
    Logger.log('Painel31: instalarPrevisaoExternaGE falhou: ' + errInst);
  }
  try {
    if (typeof recalcularPrevisaoAoVivoGE === 'function') {
      recalcularPrevisaoAoVivoGE();
      Utilities.sleep(2500);
    }
  } catch (errCalc) {
    Logger.log('Painel31: recalcularPrevisaoAoVivoGE falhou: ' + errCalc);
  }
  return painel31_reconstruir_TDTV_GE();
}

function painel31_escreverResumo_GE_(sh, colTD, nCols, lastReal, fatorDia, finalPrev, TY, D, F) {
  var resumo = [];
  var prox5 = painel31_mediaProximos_GE_(TY, D, lastReal, Math.max(0, lastReal + 1), 5);
  var temaProx = '';
  if (lastReal + 1 >= 0 && lastReal + 1 < F.length) {
    temaProx = F[lastReal + 1];
  }

  resumo.push(['RESUMO DO PAINEL TD:TV', '']);
  resumo.push(['Ultimo minuto real', lastReal >= 0 ? painel31_hora_GE_(lastReal) : 'Sem real digitado']);
  resumo.push(['Patamar do dia vs historico', painel31_fmtPct_GE_(fatorDia - 1)]);
  resumo.push(['Media prevista final', painel31_fmtNum_GE_(finalPrev)]);
  resumo.push(['Media proximos 5 min', painel31_fmtNum_GE_(prox5)]);
  resumo.push(['Tema proximo minuto', temaProx || '']);
  resumo.push(['Leitura', painel31_leituraGeral_GE_(fatorDia, prox5, finalPrev)]);
  resumo.push(['Acao sugerida', painel31_acaoGeral_GE_(fatorDia, prox5, finalPrev)]);

  var startRow = 66;
  sh.getRange(startRow, colTD, resumo.length, 2).setValues(resumo);
}

function painel31_getSpreadsheet_GE_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      return ss;
    }
  } catch (err1) {}
  if (typeof CFG !== 'undefined' && CFG && CFG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.openById('11L0-jgTgMA3JOyuUfowsOzZ0svnC3AegNgekTA62LmI');
}

function painel31_acharCabecalho_GE_(sh, texto) {
  var alvo = String(texto || '').toUpperCase();
  var maxCol = sh.getLastColumn();
  var rows = Math.min(2, sh.getMaxRows());
  var vals = sh.getRange(1, 1, rows, maxCol).getValues();
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var atual = String(vals[r][c] || '').toUpperCase();
      if (atual === alvo) {
        return c + 1;
      }
    }
  }
  return -1;
}

function painel31_fatorRecente_GE_(ratios, ate, padrao) {
  var arr = [];
  for (var i = Math.min(ate, ratios.length - 1); i >= 0 && arr.length < 5; i--) {
    if (painel31_isNumero_GE_(ratios[i])) {
      arr.push(ratios[i]);
    }
  }
  var med = painel31_mediana_GE_(arr);
  return painel31_isNumero_GE_(med) ? med : padrao;
}

function painel31_inclinacao_GE_(ratios, ate) {
  var ult3 = [];
  var ant5 = [];
  for (var i = Math.min(ate, ratios.length - 1); i >= 0; i--) {
    if (!painel31_isNumero_GE_(ratios[i])) {
      continue;
    }
    if (ult3.length < 3) {
      ult3.push(ratios[i]);
    } else if (ant5.length < 5) {
      ant5.push(ratios[i]);
    }
    if (ult3.length >= 3 && ant5.length >= 5) {
      break;
    }
  }
  if (ult3.length < 2 || ant5.length < 2) {
    return 0;
  }
  return painel31_media_GE_(ult3) - painel31_media_GE_(ant5);
}

function painel31_ruptura_GE_(slope) {
  if (!painel31_isNumero_GE_(slope)) {
    return '';
  }
  if (slope >= 0.06) {
    return 'ALTA FORTE';
  }
  if (slope >= 0.03) {
    return 'ALTA';
  }
  if (slope <= -0.06) {
    return 'QUEDA FORTE';
  }
  if (slope <= -0.03) {
    return 'QUEDA';
  }
  return 'NAO';
}

function painel31_sinal_GE_(ratio, slope, fatorTema, real) {
  if (!painel31_isNumero_GE_(ratio)) {
    return '';
  }
  if (ratio >= 1.08 && slope >= 0.02) {
    return real ? 'ALTA REAL' : 'ALTA PREVISTA';
  }
  if (ratio >= 1.03 && slope <= -0.03) {
    return 'PERDE FOLEGO';
  }
  if (ratio < 0.94 && slope >= 0.02) {
    return 'RECUPERACAO';
  }
  if (ratio < 0.92) {
    return 'RISCO';
  }
  if (fatorTema >= 1.04) {
    return 'TEMA AJUDA';
  }
  if (fatorTema <= 0.96) {
    return 'TEMA PESA';
  }
  return 'ESTAVEL';
}

function painel31_sugestao_GE_(sinal, fatorTema, tema, duracao) {
  if (!sinal) {
    return '';
  }
  if (sinal === 'RISCO') {
    return 'Antecipar tema forte';
  }
  if (sinal === 'PERDE FOLEGO') {
    return 'Trocar bloco ou acelerar virada';
  }
  if (sinal === 'RECUPERACAO') {
    return 'Segurar bloco; evitar break';
  }
  if (sinal === 'TEMA AJUDA') {
    return 'Manter tema / alongar se render';
  }
  if (sinal === 'TEMA PESA') {
    return 'Encurtar tema ou trocar';
  }
  if (duracao >= 8 && fatorTema < 1.02) {
    return 'Bloco longo: preparar troca';
  }
  return 'Manter e observar';
}

function painel31_motivo_GE_(ratio, recente, slope, fatorTema, real) {
  if (!painel31_isNumero_GE_(ratio)) {
    return '';
  }
  var origem = real ? 'real' : 'proj';
  return origem + ' vs hist ' + painel31_fmtPct_GE_(ratio - 1) +
    ' | rec ' + painel31_fmtPct_GE_(recente - 1) +
    ' | tend ' + painel31_fmtPct_GE_(slope) +
    ' | tema ' + painel31_fmtPct_GE_(fatorTema - 1);
}

function painel31_prognostico_GE_(TY, D, lastReal, idx, finalPrev) {
  var m5 = painel31_mediaProximos_GE_(TY, D, lastReal, idx, 5);
  var texto = 'Final ' + painel31_fmtNum_GE_(finalPrev);
  if (painel31_isNumero_GE_(m5)) {
    texto += ' | prox.5 ' + painel31_fmtNum_GE_(m5);
  }
  if (idx <= lastReal) {
    texto += ' | base real';
  } else {
    texto += ' | previsao';
  }
  return texto;
}

function painel31_foco_GE_(tema, fatorTema, duracao) {
  if (!tema) {
    return 'Sem tema';
  }
  if (fatorTema >= 1.04) {
    return 'Tema impulsiona: ' + tema;
  }
  if (fatorTema <= 0.96) {
    return 'Tema freia: ' + tema;
  }
  if (duracao >= 8) {
    return 'Tema longo: ' + tema;
  }
  return 'Tema neutro: ' + tema;
}

function painel31_risco_GE_(ratio, slope, fatorTema) {
  if (!painel31_isNumero_GE_(ratio)) {
    return '';
  }
  if (ratio < 0.92 || slope <= -0.06 || fatorTema <= 0.94) {
    return 'ALTO';
  }
  if (ratio < 0.98 || slope <= -0.03 || fatorTema <= 0.98) {
    return 'MEDIO';
  }
  return 'BAIXO';
}

function painel31_resumoTema_GE_(fatorTema) {
  if (!painel31_isNumero_GE_(fatorTema)) {
    return '';
  }
  if (fatorTema >= 1.04) {
    return 'AJUDA';
  }
  if (fatorTema <= 0.96) {
    return 'PESA';
  }
  return 'NEUTRO';
}

function painel31_leituraGeral_GE_(fatorDia, prox5, finalPrev) {
  if (fatorDia >= 1.06) {
    return 'Dia acima da media; proteger patamar.';
  }
  if (fatorDia <= 0.94) {
    return 'Dia abaixo da media; buscar recuperacao.';
  }
  return 'Dia proximo da media; decidir pelo tema.';
}

function painel31_acaoGeral_GE_(fatorDia, prox5, finalPrev) {
  if (fatorDia <= 0.94) {
    return 'Priorizar tema forte e evitar bloco fraco.';
  }
  if (fatorDia >= 1.06) {
    return 'Manter o que segura audiencia; evitar risco.';
  }
  return 'Observar proximos 5 minutos e ajustar tema.';
}

function painel31_mediaFinalPrevista_GE_(D, TY, lastReal) {
  var arr = [];
  for (var i = 0; i < 60; i++) {
    if (i <= lastReal && painel31_isNumero_GE_(D[i])) {
      arr.push(D[i]);
    } else if (painel31_isNumero_GE_(TY[i])) {
      arr.push(TY[i]);
    }
  }
  return painel31_media_GE_(arr);
}

function painel31_mediaProximos_GE_(TY, D, lastReal, idx, janela) {
  var arr = [];
  var ini = Math.max(0, idx);
  var fim = Math.min(59, ini + janela - 1);
  for (var i = ini; i <= fim; i++) {
    if (i <= lastReal && painel31_isNumero_GE_(D[i])) {
      arr.push(D[i]);
    } else if (painel31_isNumero_GE_(TY[i])) {
      arr.push(TY[i]);
    }
  }
  return painel31_media_GE_(arr);
}

function painel31_media_GE_(arr) {
  var s = 0;
  var n = 0;
  for (var i = 0; i < arr.length; i++) {
    if (painel31_isNumero_GE_(arr[i])) {
      s += arr[i];
      n++;
    }
  }
  return n > 0 ? s / n : null;
}

function painel31_mediana_GE_(arr) {
  var a = [];
  for (var i = 0; i < arr.length; i++) {
    if (painel31_isNumero_GE_(arr[i])) {
      a.push(arr[i]);
    }
  }
  if (a.length === 0) {
    return null;
  }
  a.sort(function(x, y) { return x - y; });
  var mid = Math.floor(a.length / 2);
  if (a.length % 2) {
    return a[mid];
  }
  return (a[mid - 1] + a[mid]) / 2;
}

function painel31_num_GE_(v) {
  if (typeof v === 'number' && isFinite(v)) {
    return v;
  }
  if (v === null || v === undefined || v === '') {
    return null;
  }
  var s = String(v).trim();
  if (!s) {
    return null;
  }
  s = s.replace('%', '');
  if (s.indexOf(',') >= 0) {
    s = s.replace(/\./g, '').replace(',', '.');
  }
  var n = Number(s);
  return isFinite(n) ? n : null;
}

function painel31_text_GE_(v) {
  if (v === null || v === undefined) {
    return '';
  }
  return String(v).trim();
}

function painel31_isNumero_GE_(v) {
  return typeof v === 'number' && isFinite(v);
}

function painel31_numOuVazio_GE_(v) {
  return painel31_isNumero_GE_(v) ? v : '';
}

function painel31_clamp_GE_(v, min, max) {
  if (!painel31_isNumero_GE_(v)) {
    return v;
  }
  if (v < min) {
    return min;
  }
  if (v > max) {
    return max;
  }
  return v;
}

function painel31_fmtPct_GE_(v) {
  if (!painel31_isNumero_GE_(v)) {
    return '';
  }
  var pct = Math.round(v * 1000) / 10;
  var s = String(pct).replace('.', ',') + '%';
  return pct > 0 ? '+' + s : s;
}

function painel31_fmtNum_GE_(v) {
  if (!painel31_isNumero_GE_(v)) {
    return '';
  }
  return String(Math.round(v * 100) / 100).replace('.', ',');
}

function painel31_hora_GE_(idx) {
  var min = idx;
  var hh = 18;
  var mm = min;
  return String(hh) + ':' + (mm < 10 ? '0' + mm : String(mm));
}

function painel31_colunaNumero_GE_(letras) {
  var s = String(letras || '').toUpperCase();
  var n = 0;
  for (var i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

// =============================================================================
// MIGRAÇÃO ESTRUTURAL — painel TD:TV → aba "Painel"
// =============================================================================
// 1) Garante a aba Painel.
// 2) Remove do Histórico os resíduos do painel antigo, identificados coluna a
//    coluna pelo par de cabeçalhos (linha 1, linha 2) exclusivo do painel.
// 3) Regenera calibragem GS–GY e fator tema GZ–HC (o painel antigo havia
//    escrito texto por cima dessas fórmulas).
// 4) Repara eventuais fórmulas #NAME? remanescentes.
// 5) Reconstrói o painel na aba nova.
// Idempotente: pode rodar de novo sem efeito colateral.
function migrarPainelParaAbaGE() {
    var ss = SpreadsheetApp.getActive();
    var hist = ss.getSheetByName(CFG.GE_OUT_HISTORICO);
    if (!hist) throw new Error('Aba "' + CFG.GE_OUT_HISTORICO + '" não encontrada.');

    painel31_getPainelSheet_GE_(ss);

    // 2) Localiza resíduos do painel antigo coluna a coluna, casando o PAR de
    //    cabeçalhos (linha 1, linha 2). Os pares do painel são exclusivos —
    //    ex.: ('RUPTURA','QUEBRA') é painel; ('RUPTURA','GX') é calibragem —
    //    então isso funciona mesmo se a calibragem já sobrescreveu parte do
    //    bloco (sobreposição parcial). Só o resíduo é limpo.
    var paresPainel = {
        'BASE HIST.|B': 1, 'DIA/HIST|D/B': 1, 'FATOR DIA|DIA': 1,
        'RECENTE|ULTIMOS': 1, 'INCLINACAO|TEND.': 1, 'RUPTURA|QUEBRA': 1,
        'FATOR FINAL|COMBO': 1, 'SINAL TEMA|UD/TY': 1, 'TROCA TEMA|MUDOU?': 1,
        'DUR. BLOCO|MIN.': 1, 'FATOR TEMA|TEMA': 1, 'SINAL|STATUS': 1,
        'SUGESTAO|ACAO': 1, 'MOTIVO|EXPLICA': 1, 'PROGNOSTICO|LEITURA': 1,
        'FOCO|EDITORIAL': 1, 'RISCO|RISCO': 1, 'MEDIA 5|PROX.5': 1,
        'FINAL PREV.|FINAL': 1
    };
    var colunasLimpas = 0;
    var lastCol = hist.getLastColumn();
    if (lastCol > 0) {
        var cab = hist.getRange(1, 1, 2, lastCol).getValues();
        for (var c = 0; c < lastCol; c++) {
            var chave = String(cab[0][c] || '').trim().toUpperCase() + '|' +
                        String(cab[1][c] || '').trim().toUpperCase();
            if (paresPainel[chave]) {
                hist.getRange(1, c + 1, 100, 1).clearContent();
                colunasLimpas++;
            }
        }
    }

    // 3) Regenera calibragem + fator tema (com o gravador locale-aware)
    atualizarCalibragemGE_(hist);
    atualizarFatorTemaGE_(hist);
    SpreadsheetApp.flush();

    // 3b) Reaponta as fórmulas de previsão em D3:D62 para a coluna FATOR TEMA
    //     ATUAL. Elas guardam o coeficiente do minuto e a letra da coluna do
    //     fator calculada NA HORA em que foram escritas; se a contagem de
    //     programas mudou desde então (ex.: o painel antigo destruiu o bloco
    //     do programa mais antigo), a letra fica órfã apontando para outra
    //     coluna. Só células com fórmula no padrão =ROUND(coef*COLn,2) são
    //     tocadas — valores reais digitados ficam intactos.
    var audColsM = getAudColsGE_(hist);
    var calibStartM = 4 + audColsM.length * 4;
    var temMinuto = !!ss.getSheetByName('Tema Minuto');
    var fatorL = temMinuto ? columnToLetter_(calibStartM + 10) : null;
    var reapontadas = 0;
    if (fatorL) {
        var dForms = hist.getRange(3, 4, 60, 1).getFormulas();
        var rePred = /^=ROUND\((\d+(?:\.\d+)?)\*[A-Z]{1,3}\d+,2\)$/;
        for (var i2 = 0; i2 < 60; i2++) {
            var fD = dForms[i2][0];
            if (!fD) continue;
            var mD = fD.match(rePred);
            if (mD) {
                var novaF = '=ROUND(' + mD[1] + '*' + fatorL + (3 + i2) + ',2)';
                if (novaF !== fD) {
                    hist.getRange(3 + i2, 4).setFormula(novaF);
                    reapontadas++;
                }
            }
        }
    }

    // 3c) Remove textos de prognóstico do painel antigo que sobraram na
    //     coluna 📌 LOG SGEST. (o módulo de alerta preserva essa coluna de
    //     propósito, então a sobra ficaria lá para sempre). Padrão exclusivo
    //     do painel: "Final 0,73 | prox.5 0,47 | previsao".
    var logLimpos = 0;
    if (temMinuto) {
        var logCol = calibStartM + 14; // HG = alertStart(+11) + 3
        if (logCol <= hist.getLastColumn() &&
            String(hist.getRange(2, logCol).getValue() || '').trim().toUpperCase() === 'HG') {
            var logVals = hist.getRange(3, logCol, 60, 1).getValues();
            var rePainelTxt = /^Final .+\|.*(previsao|base real)\s*$/;
            for (var i3 = 0; i3 < 60; i3++) {
                if (rePainelTxt.test(String(logVals[i3][0] || ''))) {
                    hist.getRange(3 + i3, logCol).clearContent();
                    logLimpos++;
                }
            }
        }
    }

    // 3d) Limpa cópias órfãs do bloco de alerta HD–HI (o bloco anda com a
    //     calibragem; quando a contagem de programas recua, a cópia antiga
    //     fica para trás). Mantém apenas o bloco na posição atual.
    var paresAlerta = {
        '🚦 SINAL|HD': 1, '💡 SUGESTÃO|HE': 1, '📋 MOTIVO|HF': 1,
        '📌 LOG SGEST.|HG': 1, '✅ ESCOLHA|HH': 1, '🎯 SEGUIU?|HI': 1
    };
    var alertStartM = temMinuto ? calibStartM + 11 : calibStartM + 7;
    var orfasLimpas = 0;
    var lastColM = hist.getLastColumn();
    if (lastColM > 0) {
        var cabM = hist.getRange(1, 1, 2, lastColM).getValues();
        for (var c2 = 0; c2 < lastColM; c2++) {
            var col1b = c2 + 1;
            if (col1b >= alertStartM && col1b <= alertStartM + 5) continue; // bloco atual
            var chaveA = String(cabM[0][c2] || '').trim().toUpperCase() + '|' +
                         String(cabM[1][c2] || '').trim().toUpperCase();
            if (paresAlerta[chaveA]) {
                hist.getRange(1, col1b, 100, 1).clearContent();
                orfasLimpas++;
            }
        }
    }

    // 3e) Regenera o bloco de alerta HD–HI na posição atual
    try {
        if (typeof atualizarAlertaColunasGE_ === 'function') atualizarAlertaColunasGE_(hist);
    } catch (e) {
        Logger.log('[Migração Painel] Alerta HD–HI não regenerado: ' + e);
    }
    SpreadsheetApp.flush();

    // 4) Repara #NAME? remanescentes
    var reparadas = 0;
    try { reparadas = repararFormulasQuebradasGE(); } catch (e) {
        Logger.log('[Migração Painel] Reparo #NAME? falhou: ' + e);
    }

    // 5) Painel na aba nova
    var msg = painel31_reconstruir_TDTV_GE();

    var resumo = 'Colunas de painel limpas no Histórico: ' + colunasLimpas +
        ' | Fórmulas D reapontadas: ' + reapontadas +
        ' | LOG SGEST. limpas: ' + logLimpos +
        ' | Colunas de alerta órfãs limpas: ' + orfasLimpas +
        ' | Fórmulas #NAME? reparadas: ' + reparadas + ' | ' + msg;
    Logger.log('[Migração Painel] ' + resumo);
    try { logGE_('MIGRACAO_PAINEL', resumo); } catch (e) {}
    return resumo;
}

function migrarPainelParaAbaMenuGE() {
    var resumo = migrarPainelParaAbaGE();
    try {
        SpreadsheetApp.getUi().alert('🏗️ Migração concluída', resumo, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
}
