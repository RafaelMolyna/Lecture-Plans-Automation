/**
 * ============================================================================
 * 1. CONFIG.JS - CONSTANTES GLOBAIS, MENU E CONFIGURAÇÕES (v4.0 - UNIVERSAL)
 * ============================================================================
 * Turmas, salas com/sem TV, disciplinas, pastas do Drive, aulas por PDF,
 * materiais e critérios usuais são lidos dinamicamente das abas "Turmas" e
 * "Disciplinas" da planilha.
 */

const CONFIG = {
  DEFAULT_MODEL: "gemini-2.5-flash",
  TIMEZONE: "America/Sao_Paulo",
  CURRENT_BIMESTRE: "4", // Fallback caso a coluna Bimestre esteja vazia
  BATCH_SIZE: 3, // 3 requisições paralelas reduzem HTTP 429/503 no Free Tier
  RETRY_ATTEMPTS: 2, // Até 2 novas tentativas apenas para itens com HTTP 429 ou 503

  TABS: {
    TURMAS: "Turmas",
    DISCIPLINAS: "Disciplinas",
    EDUC_ESP: "Educ. Esp.",
    PA_BASE: "PA (Base)",
    PA_ATUAL: "PA (Atual)",
  },
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const toolsSubMenu = ui
    .createMenu("⚙️ Ferramentas e Configs")
    .addItem("📋 Ver Logs da Última Execução", "showLastExecutionLogs")
    .addItem(
      "🗂️ Padronizar Nomes dos PDFs no Drive",
      "standardizeDrivePdfNames",
    )
    .addItem("🧠 Alternar Modelo IA (Flash ⚡ / Pro 🧠)", "toggleAiModel")
    .addItem(
      "🔑 Configurar Chave da API (Gemini) e Pastas",
      "setupCredentialsDialog",
    );

  ui.createMenu("🤖 Planos PEI")
    .addItem(
      "📅 1. Preparar Nova Semana (PA Base ➔ PA Atual)",
      "openWeekPickerDialog",
    )
    .addItem(
      "⚡ 2. Gerar Todos os Planos da Semana (Paralelo)",
      "generateAllWeekPlans",
    )
    .addItem(
      "🎯 3. Gerar Plano APENAS da Aba Atual",
      "generatePlanForActiveTab",
    )
    .addSeparator()
    .addSubMenu(toolsSubMenu)
    .addToUi();
}

function getActiveModelName() {
  return (
    PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL") ||
    CONFIG.DEFAULT_MODEL
  );
}

function toggleAiModel() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const current = getActiveModelName();
  const next =
    current === "gemini-2.5-flash" ? "gemini-2.5-pro" : "gemini-2.5-flash";

  const resp = ui.alert(
    "🧠 Alternar Modelo de IA",
    `Modelo atual: ${current}\n\n` +
      `• gemini-2.5-flash: Ultra rápido (ideal para lotes paralelos).\n` +
      `• gemini-2.5-pro: Raciocínio mais profundo (ideal para aulas complexas).\n\n` +
      `Deseja mudar para "${next}"?`,
    ui.ButtonSet.YES_NO,
  );

  if (resp === ui.Button.YES) {
    props.setProperty("GEMINI_MODEL", next);
    ui.alert(`✅ Modelo alterado para: ${next}`);
  }
}

function setupCredentialsDialog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const currentKey = props.getProperty("GEMINI_API_KEY") || "";
  const respKey = ui.prompt(
    "🔑 Configurar Gemini API Key",
    "Insira sua chave da API do Google AI Studio (aistudio.google.com):\n" +
      (currentKey
        ? "(Já existe uma chave salva. Deixe em branco para manter a atual)"
        : ""),
    ui.ButtonSet.OK_CANCEL,
  );
  if (respKey.getSelectedButton() !== ui.Button.OK) return;
  if (respKey.getResponseText().trim() !== "") {
    props.setProperty("GEMINI_API_KEY", respKey.getResponseText().trim());
  }

  const currentAulasUrl = props.getProperty("AULAS_FOLDER_ID") || "";
  const respAulas = ui.prompt(
    '📁 Pasta "Aulas" no Google Drive',
    'Cole o link ou ID da pasta principal que contém as subpastas configuradas na aba "Disciplinas":\n' +
      (currentAulasUrl
        ? `(Atual: ${currentAulasUrl} - deixe em branco para manter)`
        : ""),
    ui.ButtonSet.OK_CANCEL,
  );
  if (
    respAulas.getSelectedButton() === ui.Button.OK &&
    respAulas.getResponseText().trim() !== ""
  ) {
    props.setProperty(
      "AULAS_FOLDER_ID",
      extractDriveId(respAulas.getResponseText().trim()),
    );
  }

  const currentGuiaUrl = props.getProperty("GUIA_FOLDER_ID") || "";
  const respGuia = ui.prompt(
    '📁 Pasta do "Guia do Currículo Priorizado"',
    'Cole o link ou ID da pasta onde fica o "Guia do Currículo Priorizado":\n' +
      (currentGuiaUrl
        ? `(Atual: ${currentGuiaUrl} - deixe em branco para manter)`
        : ""),
    ui.ButtonSet.OK_CANCEL,
  );
  if (
    respGuia.getSelectedButton() === ui.Button.OK &&
    respGuia.getResponseText().trim() !== ""
  ) {
    props.setProperty(
      "GUIA_FOLDER_ID",
      extractDriveId(respGuia.getResponseText().trim()),
    );
  }

  ui.alert("✅ Configurações salvas com sucesso!");
}

function extractDriveId(urlOrId) {
  const match = String(urlOrId || "").match(/[-\w]{25,}/);
  return match ? match[0] : String(urlOrId || "").trim();
}
