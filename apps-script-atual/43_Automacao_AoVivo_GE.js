// =============================================================================
// 32_Automatizar_TY_Painel_GE.gs
// Atualização automática da TY/PROJ V2 + Painel TD:TV
//
// NÃO ALTERA a lógica da previsão.
// NÃO mexe em D, TY, UA:UD diretamente.
// Apenas chama as funções que já estão funcionando:
//   - recalcularPrevisaoAoVivoGE()
//   - painel31_reconstruir_TDTV_GE()
//
// Fluxo:
//   1) A cada 1 minuto, entre 18:00 e 18:59 BRT, recalcula TY e TD:TV.
//   2) Quando houver edição manual em D ou F, linhas 3:62, recalcula TY e TD:TV.
//   3) Se a audiência entrar por outro script, chame ge32_atualizarTYPainelAgoraGE()
//      ao final daquele script.
// =============================================================================

var GE32_SPREADSHEET_ID_FALLBACK = "11L0-jgTgMA3JOyuUfowsOzZ0svnC3AegNgekTA62LmI";
var GE32_HIST_SHEET_NAME = "Histórico";
var GE32_TZ = "America/Sao_Paulo";

// Janela operacional do Gazeta Esportiva.
// O gatilho roda o dia todo a cada 1 minuto, mas só executa de fato nesta janela.
var GE32_START_HOUR = 18;
var GE32_START_MINUTE = 0;
var GE32_END_HOUR = 18;
var GE32_END_MINUTE = 59;

// Pequena espera para garantir que a planilha e a TY terminem de recalcular antes do painel ler os dados.
// 300ms basta após o flush(); valores altos deixavam o tick lento e
// serializavam o webhook de audiência (timeout no agente OCR local).
var GE32_SLEEP_AFTER_TY_MS = 300;

// -----------------------------------------------------------------------------
// INSTALAÇÃO
// -----------------------------------------------------------------------------

// DESATIVADO: este instalador criava um tick de 1 minuto (ge32_tickAoVivoGE)
// que recalculava só o Histórico e NUNCA atualizava a aba Hoje. Rodando junto
// com geTickAoVivoGE, os dois disputavam o lock e congelavam C e M–P.
// O conjunto oficial de gatilhos é instalado por instalarTriggersGE() (arquivo 99),
// cujo tick (geTickAoVivoGE → atualizarHojeGE) já cobre TY/Painel + cockpit Hoje.
function ge32_instalarAtualizacaoAoVivoGE() {
  ge32_removerGatilhosGE();
  if (typeof instalarTriggersGE === "function") {
    instalarTriggersGE();
    Logger.log("GE32: instalador legado delegou para instalarTriggersGE().");
    return "OK: gatilhos oficiais instalados via instalarTriggersGE().";
  }
  Logger.log("GE32: instalador legado desativado — rode instalarTriggersGE() (arquivo 99).");
  return "DESATIVADO: rode instalarTriggersGE().";
}

function ge32_removerGatilhosGE() {
  var nomes = {
    ge32_tickAoVivoGE: true,
    ge32_onEditManualGE: true
  };

  var triggers = ScriptApp.getProjectTriggers();
  var removidos = 0;

  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (nomes[fn]) {
      ScriptApp.deleteTrigger(triggers[i]);
      removidos++;
    }
  }

  Logger.log("Gatilhos GE32 removidos: " + removidos);
  return "OK: gatilhos GE32 removidos: " + removidos;
}

// -----------------------------------------------------------------------------
// GATILHO DE TEMPO: roda todo minuto, mas só executa na janela 18:00-18:59 BRT.
// -----------------------------------------------------------------------------

function ge32_tickAoVivoGE() {
  if (!ge32_estaNaJanelaAoVivoGE_()) {
    return;
  }

  return ge32_atualizarTYPainelAgora_("tempo_1min_18h");
}

// -----------------------------------------------------------------------------
// GATILHO ON EDIT: só reage a D ou F, linhas 3:62, na aba GE - Histórico.
// -----------------------------------------------------------------------------

function ge32_onEditManualGE(e) {
  try {
    if (!e || !e.range) return;

    var sh = e.range.getSheet();
    if (!sh || sh.getName() !== GE32_HIST_SHEET_NAME) return;

    var row = e.range.getRow();
    var col = e.range.getColumn();

    if (row < 3 || row > 62) return;
    if (col !== 4 && col !== 6) return; // D = audiência; F = tema

    return ge32_atualizarTYPainelAgora_("onEdit_DF");
  } catch (err) {
    Logger.log("Erro em ge32_onEditManualGE: " + (err && err.stack ? err.stack : err));
  }
}

// -----------------------------------------------------------------------------
// FUNÇÃO PÚBLICA: rode manualmente ou chame no fim do script que importa audiência.
// -----------------------------------------------------------------------------

function ge32_atualizarTYPainelAgoraGE() {
  return ge32_atualizarTYPainelAgora_("manual");
}

// Use esta função no fim do script que atualiza os números reais em D:
//   ge32_atualizarDepoisDeImportarAudienciaGE();
function ge32_atualizarDepoisDeImportarAudienciaGE() {
  return ge32_atualizarTYPainelAgora_("pos_importacao_audiencia");
}

// -----------------------------------------------------------------------------
// NÚCLEO: recalcula TY e depois TD:TV.
// -----------------------------------------------------------------------------

function ge32_atualizarTYPainelAgora_(fonte) {
  var lock = LockService.getScriptLock();

  // Evita duas execuções simultâneas: por exemplo, onEdit + gatilho de minuto.
  // No caminho do tick da aba Hoje ("hoje_sync") espera de verdade: se desistir,
  // a coluna C e o bloco M–P ficam com projeção/painel velhos até o próximo tick
  // que ganhar a corrida. Nos caminhos de onEdit mantém a espera curta para não
  // travar a digitação do usuário.
  var obteveLock;
  if (fonte === "hoje_sync") {
    // Espera moderada: 5s cobre um doPost/onEdit em andamento sem empilhar
    // ticks quando algo trava (o guarda de overlap do tick faz o resto).
    try { lock.waitLock(5000); obteveLock = true; } catch (eLock) { obteveLock = false; }
  } else {
    obteveLock = lock.tryLock(2500);
  }
  if (!obteveLock) {
    Logger.log("GE32 ignorado: já há outra atualização em andamento. Fonte: " + fonte);
    return "IGNORADO: outra atualização em andamento.";
  }

  try {
    var inicio = new Date();
    Logger.log("GE32 início. Fonte: " + fonte + " | " + inicio.toISOString());

    if (typeof recalcularPrevisaoAoVivoGE !== "function") {
      throw new Error("Função recalcularPrevisaoAoVivoGE() não encontrada. Confira o arquivo 20_PrevisaoV2_GE.gs.");
    }

    if (typeof painel31_reconstruir_TDTV_GE !== "function") {
      throw new Error("Função painel31_reconstruir_TDTV_GE() não encontrada. Confira o arquivo 31 do painel TD:TV.");
    }

    // 1) Recalcula TY/PROJ V2.
    recalcularPrevisaoAoVivoGE();
    SpreadsheetApp.flush();

    // 2) Aguarda um pouco para TY e fórmulas auxiliares assentarem.
    if (GE32_SLEEP_AFTER_TY_MS > 0) {
      Utilities.sleep(GE32_SLEEP_AFTER_TY_MS);
    }

    // 3) Reconstrói o painel TD:TV com base nos valores já calculados.
    painel31_reconstruir_TDTV_GE();
    SpreadsheetApp.flush();

    var fim = new Date();
    var ms = fim.getTime() - inicio.getTime();
    Logger.log("GE32 OK. Fonte: " + fonte + " | tempo: " + ms + " ms");

    return "OK: TY + TD:TV atualizados. Fonte: " + fonte + " | " + ms + " ms";
  } catch (err) {
    Logger.log("ERRO GE32: " + (err && err.stack ? err.stack : err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------------------------------------------------------
// NOVO DIA: função opcional para rodar depois de abrir o novo dia.
// -----------------------------------------------------------------------------

function ge32_prepararNovoDiaGE() {
  var msgs = [];

  if (typeof instalarPrevisaoExternaGE === "function") {
    instalarPrevisaoExternaGE();
    msgs.push("instalarPrevisaoExternaGE OK");
  } else {
    msgs.push("instalarPrevisaoExternaGE não encontrada; pulei esta etapa");
  }

  msgs.push(ge32_atualizarTYPainelAgora_("preparar_novo_dia"));

  // Garante que os gatilhos GE32 estejam instalados para o dia.
  ge32_instalarAtualizacaoAoVivoGE();
  msgs.push("gatilhos GE32 reinstalados");

  Logger.log(msgs.join(" | "));
  return msgs.join(" | ");
}

// -----------------------------------------------------------------------------
// STATUS / DIAGNÓSTICO
// -----------------------------------------------------------------------------

function ge32_verStatusGE() {
  var agora = ge32_agoraBRT_();
  var triggers = ScriptApp.getProjectTriggers();
  var linhas = [];

  linhas.push("Agora BRT: " + agora.horaStr);
  linhas.push("Na janela 18:00-18:59? " + (ge32_estaNaJanelaAoVivoGE_() ? "SIM" : "NÃO"));

  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === "ge32_tickAoVivoGE" || fn === "ge32_onEditManualGE") {
      linhas.push("Gatilho instalado: " + fn + " | tipo: " + triggers[i].getEventType());
    }
  }

  var msg = linhas.join("\n");
  Logger.log(msg);
  return msg;
}

function ge32_testeSemMexerGE() {
  Logger.log("OK: arquivo 32 compilou e executou sem mexer na planilha.");
  return "OK: arquivo 32 compilou e executou sem mexer na planilha.";
}

// -----------------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------------

function ge32_getSpreadsheetId_() {
  try {
    if (typeof CFG !== "undefined" && CFG && CFG.SPREADSHEET_ID) {
      return CFG.SPREADSHEET_ID;
    }
  } catch (err) {}

  return GE32_SPREADSHEET_ID_FALLBACK;
}

function ge32_agoraBRT_() {
  var now = new Date();
  var h = parseInt(Utilities.formatDate(now, GE32_TZ, "H"), 10);
  var m = parseInt(Utilities.formatDate(now, GE32_TZ, "m"), 10);
  var s = parseInt(Utilities.formatDate(now, GE32_TZ, "s"), 10);
  var data = Utilities.formatDate(now, GE32_TZ, "yyyy-MM-dd");
  var horaStr = Utilities.formatDate(now, GE32_TZ, "dd/MM/yyyy HH:mm:ss");

  return {
    hour: h,
    minute: m,
    second: s,
    date: data,
    horaStr: horaStr
  };
}

function ge32_estaNaJanelaAoVivoGE_() {
  var a = ge32_agoraBRT_();
  var atual = a.hour * 60 + a.minute;
  var ini = GE32_START_HOUR * 60 + GE32_START_MINUTE;
  var fim = GE32_END_HOUR * 60 + GE32_END_MINUTE;

  return atual >= ini && atual <= fim;
}
