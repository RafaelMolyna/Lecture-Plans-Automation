/**
 * ============================================================================
 * 1. CONFIG.JS - CONSTANTES GLOBAIS, MENU E CONFIGURAÇÕES (v3.1)
 * ============================================================================
 */

const CONFIG = {
  DEFAULT_MODEL: "gemini-2.5-flash",
  TIMEZONE: "America/Sao_Paulo",
  CURRENT_BIMESTRE: "4",
  BATCH_SIZE: 3, // 3 requisições paralelas reduzem 429/503 no Free Tier
  RETRY_ATTEMPTS: 2, // Até 2 novas tentativas apenas para itens com HTTP 429 ou 503

  CLASS_MAP: {
    "1º Honestidade": { tab: "1º HO", grade: 1, noTv: true },
    "1º Humildade": { tab: "1º HU", grade: 1, noTv: false },
    "2º Lealdade": { tab: "2º LE", grade: 2, noTv: true },
    "2º Solidariedade": { tab: "2º SO", grade: 2, noTv: false },
    "3º Altruísmo": { tab: "3º AL", grade: 3, noTv: false },
    "3º Resiliência": { tab: "3º RE", grade: 3, noTv: true },
    Eletiva: { tab: "Eletiva", grade: 0, noTv: false, isEletiva: true },
  },

  TAB_TO_FULLNAME: {
    "1º HO": "1º Honestidade",
    "1º HU": "1º Humildade",
    "2º LE": "2º Lealdade",
    "2º SO": "2º Solidariedade",
    "3º AL": "3º Altruísmo",
    "3º RE": "3º Resiliência",
    Eletiva: "Eletiva",
  },

  SUBJECT_FOLDER_NAMES: {
    Física_1: "Física 1",
    Física_2: "Física 2",
    Física_3: "Física 3",
    Robótica_1: "Robótica 1",
    Robótica_2: "Robótica 2",
    Robótica_3: "Robótica 2",
  },

  MASTER_MATERIALS: {
    Física: [
      "Quadro Branco",
      "Marcadores para quadro branco",
      "Livro do Estudante (SEDUC)",
      "Caderno do Estudante",
      "Material Impresso / Lista de Exercícios",
      "Simulador PhET Colorado",
      "Dispositivos eletrônicos / Smartphones",
      "Carrinhos de dinâmica",
      "Massas aferidas / Blocos",
      "Polias e barbante",
      "Bolinhas de borracha / gude",
      "Massinha de modelar",
      "Corda e Mola (Slinky)",
      "Ímãs de neodímio",
      "Bússola",
      "Fio condutor e Fonte de tensão / Bateria",
      "Resistores de carbono/grafite",
    ],
    Robótica: [
      "Quadro Branco",
      "Cadernos, lápis, borracha e canetas",
      "Kits de Robótica",
      "Placa BBC micro:bit V2",
      "Shield Robotbit com bateria",
      "Servomotores e Motores DC",
      "Cabos USB",
      "Plataforma Microsoft MakeCode",
      "Teachable Machine / CreateAI",
      "Papelão e Gabaritos impressos",
      "Estiletes e Tesouras sem ponta",
      "Réguas",
      "Cola quente e Cola branca",
      "Varetas de churrasco / Palitos",
      "Chaves Philips / Fenda e Parafusos",
      "Lixas (180/220)",
    ],
    Eletiva: [
      "Quadro Branco",
      "Bancadas de laboratório",
      "Roteiro experimental impresso",
      "Caderno, lápis e caneta",
      "Materiais de baixo custo para experimentos mão na massa",
      "Carrinhos, polias, massas, elásticos, ímãs, fontes de tensão, ferramentas manuais",
    ],
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
    'Cole o link ou ID da pasta principal "Aulas" (Física 1, 2, 3 e Robótica 1, 2):\n' +
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
  const match = urlOrId.match(/[-\w]{25,}/);
  return match ? match[0] : urlOrId.trim();
  s;
}
