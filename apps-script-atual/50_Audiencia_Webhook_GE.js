// =============================================================================
// 19_Audiencia_GE.js — V4.1 LIVE FAST
//
// Objetivo:
// - Receber batch do PowerShell V4.1.
// - Lançar na coluna D o mais rápido possível.
// - Não ler nota e não escrever comentário durante o ao vivo.
// - Prioridade por minuto: FINAL > REVISAO > PREVIA.
// - Isso reduz muito timeout do Apps Script.
//
// Observação:
// - A nota detalhada fica desativada durante o ao vivo para velocidade.
// - Se quiser reativar nota simples, troque WRITE_NOTE_DURING_LIVE para true.
// =============================================================================

var AUD_GE_CFG = {
  SHEET_NAME: "Hoje",
  SECRET_PROP: "AUDIENCIA_GE_SECRET",
  WEBAPP_URL_PROP: "AUDIENCIA_GE_WEBAPP_URL",
  FIRST_ROW: 3,
  AUD_COL: 4,
  TIMEZONE: "America/Sao_Paulo",
  WRITE_NOTE_DURING_LIVE: false
};

// =============================================================================
// MENU / FUNÇÕES DE APOIO
// =============================================================================

function configurarAudienciaGE() {
  return configurarWebhookAudienciaGE();
}

function configurarWebhookAudienciaGE() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var currentSecret = props.getProperty(AUD_GE_CFG.SECRET_PROP) || "";
  var prompt = ui.prompt(
    "🔑 Configurar Captura de Audiência GE",
    "Digite a chave secreta que o PowerShell usará.\n\nAtual: " + (currentSecret ? "(já configurada)" : "(não configurada)") + "\n\nDica: use uma palavra simples, sem acentos e sem espaços.",
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  var secret = (prompt.getResponseText() || "").trim();

  if (!secret) {
    ui.alert("⚠️ Chave vazia", "A chave não foi alterada.", ui.ButtonSet.OK);
    return;
  }

  props.setProperty(AUD_GE_CFG.SECRET_PROP, secret);

  var url = props.getProperty(AUD_GE_CFG.WEBAPP_URL_PROP) || descobrirWebAppUrlAudienciaGE_();

  ui.alert(
    "✅ Captura configurada",
    "Chave salva.\n\nURL atual:\n" + (url || "(Web App ainda não implantado)"),
    ui.ButtonSet.OK
  );
}

function prepararAudienciaGE() {
  return prepararGradeAudienciaGE();
}

function prepararGradeAudienciaGE() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(AUD_GE_CFG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(AUD_GE_CFG.SHEET_NAME);
  }

  if (sheet.getMaxRows() < 62) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 62 - sheet.getMaxRows());
  }

  sheet.getRange("A2").setValue("HORA");
  sheet.getRange("D2").setValue("AUDIÊNCIA");

  var horas = [];
  for (var m = 0; m <= 59; m++) {
    horas.push(["18:" + String(m).padStart(2, "0")]);
  }

  sheet.getRange(3, 1, 60, 1).setValues(horas);
  sheet.getRange(3, 4, 60, 1).setNumberFormat("0.00");

  SpreadsheetApp.getUi().alert(
    "✅ Grade preparada",
    "A faixa A3:A62 foi preparada com 18:00–18:59.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function verStatusAudienciaGE() {
  return statusAudienciaGE();
}

function statusAudienciaGE() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty(AUD_GE_CFG.SECRET_PROP);
  var url = props.getProperty(AUD_GE_CFG.WEBAPP_URL_PROP) || descobrirWebAppUrlAudienciaGE_();

  var sheet = SpreadsheetApp.getActive().getSheetByName(AUD_GE_CFG.SHEET_NAME);
  var sheetStatus = sheet ? "OK" : "Aba não encontrada";

  SpreadsheetApp.getUi().alert(
    "📋 Status — Audiência GE V4.1 LIVE FAST",
    "Aba: " + AUD_GE_CFG.SHEET_NAME + " — " + sheetStatus +
    "\nFaixa: D3:D62" +
    "\nModo: PREVIA rápida + FINAL consolidada" +
    "\nNotas ao vivo: " + (AUD_GE_CFG.WRITE_NOTE_DURING_LIVE ? "ativas" : "desativadas para velocidade") +
    "\nChave: " + (secret ? "configurada" : "NÃO configurada") +
    "\nWeb App URL:\n" + (url || "(não encontrada; publique o Web App)"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function testarAudienciaGE() {
  return testarWebhookAudienciaGE();
}

function testarWebhookAudienciaGE() {
  var res = validarEGravarAudienciaGE_({
    mode: "batch_v41_fast",
    readings: [
      {
        horario: "18:00",
        audiencia: "0.46",
        stage: "PREVIA",
        stage_index: 1,
        raw: "TESTE PREVIA"
      },
      {
        horario: "18:00",
        audiencia: "0.40",
        stage: "FINAL",
        stage_index: 3,
        raw: "TESTE FINAL"
      }
    ],
    source: "menu"
  });

  SpreadsheetApp.getUi().alert(
    "🧪 Teste V4.1",
    JSON.stringify(res, null, 2),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function limparAudienciaGE() {
  var ui = SpreadsheetApp.getUi();

  var confirm = ui.alert(
    "🧹 Limpar D3:D62",
    "Deseja limpar a audiência real da coluna D, linhas 3 a 62, incluindo notas?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  var sheet = SpreadsheetApp.getActive().getSheetByName(AUD_GE_CFG.SHEET_NAME);
  if (!sheet) {
    ui.alert("❌ Erro", "Aba não encontrada: " + AUD_GE_CFG.SHEET_NAME, ui.ButtonSet.OK);
    return;
  }

  sheet.getRange(3, 4, 60, 1).clearContent();
  sheet.getRange(3, 4, 60, 1).clearNote();

  ui.alert("✅ Limpo", "D3:D62 e as notas dessas células foram limpas.", ui.ButtonSet.OK);
}

// Aliases para compatibilidade com menus antigos.
function configurarWebhookOCRAudienciaGE() { return configurarWebhookAudienciaGE(); }
function configurarCapturaAudienciaGE() { return configurarWebhookAudienciaGE(); }
function prepararD3D62AudienciaGE() { return prepararGradeAudienciaGE(); }
function prepararCapturaAudienciaGE() { return prepararGradeAudienciaGE(); }
function verStatusCapturaAudienciaGE() { return statusAudienciaGE(); }
function testarCapturaAudienciaGE() { return testarWebhookAudienciaGE(); }
function limparCapturaAudienciaGE() { return limparAudienciaGE(); }

// =============================================================================
// HANDLERS PARA O DOPOST ROUTER
// =============================================================================

function processarAudienciaGEWebhook_(payload) {
  return validarEGravarAudienciaGE_(payload);
}

function receberAudienciaGE_(payload) {
  return validarEGravarAudienciaGE_(payload);
}

function handleAudienciaGEPost_(payload) {
  return validarEGravarAudienciaGE_(payload);
}

function handleAudienciaGEWebhook_(payload) {
  return validarEGravarAudienciaGE_(payload);
}

function audienciaGEDoPost_(payload) {
  return validarEGravarAudienciaGE_(payload);
}

function salvarAudienciaGEWebhook_(payload) {
  return validarEGravarAudienciaGE_(payload);
}

// =============================================================================
// VALIDAÇÃO E ROTEAMENTO
// =============================================================================

function validarEGravarAudienciaGE_(payload) {
  payload = payload || {};

  var props = PropertiesService.getScriptProperties();
  var expectedSecret = props.getProperty(AUD_GE_CFG.SECRET_PROP) || "";

  // Permite teste de menu sem secret.
  var isMenuTest = String(payload.source || "").toLowerCase() === "menu";

  if (!isMenuTest) {
    if (!expectedSecret) {
      return {
        success: false,
        ok: false,
        error: "Chave secreta não configurada no Apps Script."
      };
    }

    var receivedSecret = String(payload.secret || "").trim();

    if (receivedSecret !== expectedSecret) {
      return {
        success: false,
        ok: false,
        error: "Chave secreta inválida."
      };
    }
  }

  if (payload.readings && payload.readings.length) {
    return gravarAudienciasGEBatchFast_(payload);
  }

  return gravarAudienciasGEBatchFast_({
    source: payload.source || "single",
    capturedAt: payload.capturedAt || "",
    readings: [payload]
  });
}

// =============================================================================
// GRAVAÇÃO FAST
// =============================================================================

function gravarAudienciasGEBatchFast_(payload) {
  var startedAt = new Date();
  var readings = payload.readings || [];
  var normalized = [];
  var errors = [];

  for (var i = 0; i < readings.length; i++) {
    var item = readings[i] || {};
    var horario = normalizarHorarioAudienciaGE_(item.horario);
    var audiencia = normalizarAudienciaGE_(item.audiencia || item.audiencia_com_virgula);
    var stage = normalizarStageAudienciaGE_(item.stage, item.stage_index);

    if (!horario) {
      errors.push({ index: i, error: "Horário inválido", item: item });
      continue;
    }

    if (audiencia === null) {
      errors.push({ index: i, error: "Audiência inválida", item: item });
      continue;
    }

    var row = calcularLinhaAudienciaGE_(horario);

    normalized.push({
      horario: horario,
      row: row,
      value: audiencia,
      valueComma: formatAudienciaCommaGE_(audiencia),
      stage: stage,
      raw: String(item.raw || ""),
      audienceHint: String(item.audience_hint || ""),
      capturedAt: String(item.capturedAt || payload.capturedAt || ""),
      source: String(item.source || payload.source || "")
    });
  }

  if (!normalized.length) {
    return {
      success: false,
      ok: false,
      mode: "batch_v41_fast",
      total: readings.length,
      ok_count: 0,
      error_count: errors.length,
      errors: errors,
      message: "Nenhuma leitura válida para gravar."
    };
  }

  // Se houver mais de uma leitura para a mesma linha, escolhe a melhor.
  // Prioridade: FINAL > REVISAO > PREVIA.
  var byRow = {};
  normalized.forEach(function(item) {
    var key = String(item.row);
    if (!byRow[key] || item.stage.priority > byRow[key].stage.priority) {
      byRow[key] = item;
    }
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUD_GE_CFG.SHEET_NAME);

  if (!sheet) {
    return {
      success: false,
      ok: false,
      error: "Aba não encontrada: " + AUD_GE_CFG.SHEET_NAME
    };
  }

  var rows = Object.keys(byRow).sort(function(a, b) {
    return Number(a) - Number(b);
  });

  var ok = [];
  for (var r = 0; r < rows.length; r++) {
    var chosen = byRow[rows[r]];
    var cell = sheet.getRange(chosen.row, AUD_GE_CFG.AUD_COL);

    // Modo mais rápido: só grava valor e formato.
    cell.setValue(chosen.value);
    cell.setNumberFormat("0.00");

    // Comentário desativado por padrão para não atrasar o ao vivo.
    if (AUD_GE_CFG.WRITE_NOTE_DURING_LIVE) {
      cell.setNote(montarNotaRapidaGE_(chosen));
    }

    ok.push({
      horario: chosen.horario,
      audiencia: chosen.valueComma,
      stage: chosen.stage.label,
      row: chosen.row,
      cell: "D" + chosen.row
    });
  }

  var elapsedMs = new Date().getTime() - startedAt.getTime();

  return {
    success: true,
    ok: true,
    mode: "batch_v41_fast",
    total: readings.length,
    valid_count: normalized.length,
    ok_count: ok.length,
    error_count: errors.length,
    elapsed_ms: elapsedMs,
    results: ok,
    errors: errors
  };
}

function montarNotaRapidaGE_(item) {
  var nowText = Utilities.formatDate(new Date(), AUD_GE_CFG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");

  var linhas = [
    "Atualizado por OCR em " + nowText,
    "Horário: " + item.horario,
    "Audiência: " + item.valueComma,
    "Status: " + item.stage.label
  ];

  if (item.capturedAt) linhas.push("Capturado: " + item.capturedAt);
  if (item.raw) linhas.push("OCR: " + reduzirLinhaNotaGE_(item.raw));
  if (item.audienceHint) linhas.push("OCR audiência: " + reduzirLinhaNotaGE_(item.audienceHint));
  if (item.source) linhas.push("Fonte: " + item.source);

  return linhas.join("\n");
}

function reduzirLinhaNotaGE_(value) {
  var text = String(value || "").replace(/\r?\n/g, " | ");
  if (text.length > 200) return text.slice(0, 200) + "...";
  return text;
}

// =============================================================================
// NORMALIZADORES
// =============================================================================

function calcularLinhaAudienciaGE_(horario) {
  var match = String(horario || "").match(/^18:([0-5][0-9])$/);

  if (!match) {
    throw new Error("Horário fora da faixa 18:00-18:59: " + horario);
  }

  return 3 + Number(match[1]);
}

function normalizarHorarioAudienciaGE_(value) {
  if (value === null || value === undefined) return "";

  var text = String(value).trim();
  text = text.replace(/[hH]/g, ":");
  text = text.replace(/\s+/g, "");

  var match = text.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return "";

  var h = Number(match[1]);
  var m = Number(match[2]);

  if (h !== 18) return "";
  if (m < 0 || m > 59) return "";

  return "18:" + String(m).padStart(2, "0");
}

function normalizarAudienciaGE_(value) {
  if (value === null || value === undefined) return null;

  var text = String(value).trim();
  text = text.replace(",", ".");
  text = text.replace(/[^0-9.]/g, "");

  if (!text) return null;

  // Corrige casos tipo "0.26.".
  var m = text.match(/^(\d{1,2}\.\d{1,2})/);
  if (m) text = m[1];

  var n = Number(text);

  if (isNaN(n) || !isFinite(n)) return null;
  if (n < 0 || n > 30) return null;

  return n;
}

function normalizarStageAudienciaGE_(stage, stageIndex) {
  var s = removerAcentosAudienciaGE_(String(stage || "")).toUpperCase().trim();
  var idx = Number(stageIndex || 0);

  if (s === "FINAL" || idx === 3) {
    return {
      key: "final",
      label: "FINAL",
      index: 3,
      priority: 3
    };
  }

  if (s === "REVISAO" || s === "REVISION" || idx === 2) {
    return {
      key: "revisao",
      label: "REVISAO",
      index: 2,
      priority: 2
    };
  }

  return {
    key: "previa",
    label: "PREVIA",
    index: 1,
    priority: 1
  };
}

function formatAudienciaCommaGE_(number) {
  if (number === "" || number === null || number === undefined) return "";
  return Number(number).toFixed(2).replace(".", ",");
}

function removerAcentosAudienciaGE_(value) {
  return String(value || "")
    .replace(/[ÁÀÃÂÄ]/g, "A")
    .replace(/[ÉÈÊË]/g, "E")
    .replace(/[ÍÌÎÏ]/g, "I")
    .replace(/[ÓÒÕÔÖ]/g, "O")
    .replace(/[ÚÙÛÜ]/g, "U")
    .replace(/[Ç]/g, "C")
    .replace(/[áàãâä]/g, "a")
    .replace(/[éèêë]/g, "e")
    .replace(/[íìîï]/g, "i")
    .replace(/[óòõôö]/g, "o")
    .replace(/[úùûü]/g, "u")
    .replace(/[ç]/g, "c");
}

function descobrirWebAppUrlAudienciaGE_() {
  try {
    var url = ScriptApp.getService().getUrl();

    if (url) {
      PropertiesService.getScriptProperties().setProperty(AUD_GE_CFG.WEBAPP_URL_PROP, url);
      return url;
    }
  } catch (err) {}

  return "";
}

// =============================================================================
// DOPOST DIRETO OPCIONAL
// No projeto Radar, o doPost principal fica no 16_Speech_Integration.js.
// =============================================================================

function doPostAudienciaGEStandalone(e) {
  var payload = {};

  try {
    payload = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : "{}");
  } catch (err) {
    payload = {};
  }

  var res = validarEGravarAudienciaGE_(payload);

  return ContentService
    .createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}
