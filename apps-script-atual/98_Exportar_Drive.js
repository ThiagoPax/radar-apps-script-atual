// ============================================================================
// 20_Exportar_Drive.gs
// Exporta todos os arquivos deste projeto Apps Script para uma pasta no Drive
// e cria também um ZIP com todos os scripts.
// ============================================================================

const EXPORT_DRIVE_CFG = {
  NOME_PROJETO: "Radar",

  // Para salvar dentro de uma pasta específica do Drive, cole aqui o ID dela.
  // Deixando vazio, a pasta será criada na raiz do Meu Drive.
  PASTA_MAE_ID: "",

  // Cria também um arquivo ZIP com todos os scripts exportados.
  CRIAR_ZIP: true
};

function exportarScriptsParaGoogleDrive() {
  const scriptId = ScriptApp.getScriptId();
  const timezone = Session.getScriptTimeZone() || "America/Sao_Paulo";
  const agora = new Date();

  const dataHora = Utilities.formatDate(agora, timezone, "yyyy-MM-dd_HH-mm-ss");
  const nomePastaBackup = `${EXPORT_DRIVE_CFG.NOME_PROJETO}_AppsScript_Backup_${dataHora}`;

  const pastaMae = EXPORT_DRIVE_CFG.PASTA_MAE_ID
    ? DriveApp.getFolderById(EXPORT_DRIVE_CFG.PASTA_MAE_ID)
    : DriveApp.getRootFolder();

  const pastaBackup = pastaMae.createFolder(nomePastaBackup);

  const url = `https://script.googleapis.com/v1/projects/${encodeURIComponent(scriptId)}/content`;

  const resposta = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    muteHttpExceptions: true
  });

  const codigoResposta = resposta.getResponseCode();
  const textoResposta = resposta.getContentText();

  if (codigoResposta !== 200) {
    throw new Error(
      "Não foi possível ler os arquivos do Apps Script.\n\n" +
      `Código HTTP: ${codigoResposta}\n\n` +
      textoResposta + "\n\n" +
      "Verifique se a Apps Script API está ativada e se os escopos foram autorizados."
    );
  }

  const conteudo = JSON.parse(textoResposta);
  const arquivos = conteudo.files || [];

  if (!arquivos.length) {
    throw new Error("Nenhum arquivo foi encontrado neste projeto Apps Script.");
  }

  const blobsParaZip = [];
  const arquivosCriados = [];

  arquivos.forEach(function (arquivo) {
    const nomeArquivo = montarNomeArquivoExportado_(arquivo);
    const mimeType = obterMimeTypePorTipo_(arquivo.type);
    const source = arquivo.source || "";

    const blob = Utilities.newBlob(source, mimeType, nomeArquivo);

    pastaBackup.createFile(blob);
    blobsParaZip.push(blob.copyBlob());

    arquivosCriados.push(nomeArquivo);
  });

  let zipUrl = "";

  if (EXPORT_DRIVE_CFG.CRIAR_ZIP) {
    const nomeZip = `${nomePastaBackup}.zip`;
    const zipBlob = Utilities.zip(blobsParaZip, nomeZip);
    const zipFile = pastaBackup.createFile(zipBlob);
    zipUrl = zipFile.getUrl();
  }

  const resumo =
    "EXPORTAÇÃO CONCLUÍDA\n\n" +
    `Pasta criada no Drive: ${pastaBackup.getUrl()}\n` +
    `Total de arquivos exportados: ${arquivosCriados.length}\n` +
    (zipUrl ? `ZIP criado: ${zipUrl}\n` : "") +
    "\nArquivos:\n" +
    arquivosCriados.map(function (nome) {
      return `- ${nome}`;
    }).join("\n");

  Logger.log(resumo);
  console.log(resumo);

  return resumo;
}

function montarNomeArquivoExportado_(arquivo) {
  const nomeBase = limparNomeArquivo_(arquivo.name || "arquivo");

  switch (arquivo.type) {
    case "SERVER_JS":
      return `${nomeBase}.gs`;

    case "HTML":
      return `${nomeBase}.html`;

    case "JSON":
      return `${nomeBase}.json`;

    default:
      return `${nomeBase}.txt`;
  }
}

function obterMimeTypePorTipo_(tipo) {
  switch (tipo) {
    case "SERVER_JS":
      return "text/javascript";

    case "HTML":
      return "text/html";

    case "JSON":
      return "application/json";

    default:
      return "text/plain";
  }
}

function limparNomeArquivo_(nome) {
  return String(nome)
    .replace(/[\\/:*?"<>|#%{}~&]/g, "_")
    .trim();
}