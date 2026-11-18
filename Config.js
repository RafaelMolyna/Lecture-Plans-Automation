/**
 * ============================================================================
 * AUTOMAÇÃO DE PLANOS DE AULA PEI - EE MATILDE MARIA CREMM
 * ============================================================================
 * Funcionalidades:
 * 1. Prepara a semana copiando 'PA (Base)' para 'PA (Atual)' e atualizando datas.
 * 2. Lê os rascunhos das abas das turmas (Física e Robótica independentemente).
 * 3. Identifica PDFs no Drive (com regra 2x1 para Robótica) ou links diretos.
 * 4. Consulta o 'Guia do Currículo Priorizado' dinamicamente (com cache seguro).
 * 5. Lê a aba 'Educ. Esp.' para adaptações curriculares personalizadas.
 * 6. Aplica restrição de TV/Projetor para salas sem televisor (1º HO, 2º LE, 3º RE).
 * 7. Divide dobradinhas (Aula 1 e Aula 2 de 50 min) via Gemini 2.5 Flash API.

 * Novidades v2.0:
 * - Suporte completo à aba 'Eletiva' (formato simplificado sem Habilidades/Critérios).
 * - Alternador rápido no menu entre Gemini 2.5 Flash (Rápido) e Gemini 2.5 Pro (Avançado).
 * - Mapeamento automático de alunos da Educação Especial também na Eletiva.
 */


/**
 * ============================================================================
 * 1. CONFIG.GS - CONSTANTES GLOBAIS, MENU E CONFIGURAÇÕES
 * ============================================================================
 */

const CONFIG = {
  DEFAULT_MODEL: 'gemini-2.5-flash',
  TIMEZONE: 'America/Sao_Paulo',
  CURRENT_BIMESTRE: '4', // Altere para '1', '2', '3' ou '4' conforme o bimestre

  // Mapeamento entre o nome na célula do PA e o nome da aba da turma
  CLASS_MAP: {
    '1º Honestidade':   { tab: '1º HO', grade: 1, noTv: true },
    '1º Humildade':     { tab: '1º HU', grade: 1, noTv: false },
    '2º Lealdade':      { tab: '2º LE', grade: 2, noTv: true },
    '2º Solidariedade': { tab: '2º SO', grade: 2, noTv: false },
    '3º Altruísmo':     { tab: '3º AL', grade: 3, noTv: false },
    '3º Resiliência':   { tab: '3º RE', grade: 3, noTv: true },
    'Eletiva':          { tab: 'Eletiva', grade: 0, noTv: false, isEletiva: true }
  },

  // Mapeamento inverso (Aba -> Nome completo)
  TAB_TO_FULLNAME: {
    '1º HO': '1º Honestidade',
    '1º HU': '1º Humildade',
    '2º LE': '2º Lealdade',
    '2º SO': '2º Solidariedade',
    '3º AL': '3º Altruísmo',
    '3º RE': '3º Resiliência',
    'Eletiva': 'Eletiva'
  },

  // Nomes das pastas de aulas dentro da sua pasta principal no Drive
  SUBJECT_FOLDER_NAMES: {
    'Física_1': 'Física 1',
    'Física_2': 'Física 2',
    'Física_3': 'Física 3',
    'Robótica_1': 'Robótica 1',
    'Robótica_2': 'Robótica 2',
    'Robótica_3': 'Robótica 2'
  },

  // Catálogo mestre de materiais
  MASTER_MATERIALS: {
    'Física': [
      'Quadro Branco', 'Marcadores para quadro branco', 'Livro do Estudante (SEDUC)',
      'Caderno do Estudante', 'Material Impresso / Lista de Exercícios',
      'Simulador PhET Colorado', 'Dispositivos eletrônicos / Smartphones',
      'Carrinhos de dinâmica', 'Massas aferidas / Blocos', 'Polias e barbante',
      'Bolinhas de borracha / gude', 'Massinha de modelar', 'Corda e Mola (Slinky)',
      'Ímãs de neodímio', 'Bússola', 'Fio condutor e Fonte de tensão / Bateria',
      'Resistores de carbono/grafite'
    ],
    'Robótica': [
      'Quadro Branco', 'Cadernos, lápis, borracha e canetas',
      'Kits de Robótica', 'Placa BBC micro:bit V2', 'Shield Robotbit com bateria',
      'Servomotores e Motores DC', 'Cabos USB', 'Plataforma Microsoft MakeCode',
      'Teachable Machine / CreateAI', 'Papelão e Gabaritos impressos',
      'Estiletes e Tesouras sem ponta', 'Réguas', 'Cola quente e Cola branca',
      'Varetas de churrasco / Palitos', 'Chaves Philips / Fenda e Parafusos',
      'Lixas (180/220)'
    ],
    'Eletiva': [
      'Quadro Branco', 'Bancadas de laboratório', 'Roteiro experimental impresso',
      'Caderno, lápis e caneta', 'Materiais de baixo custo para experimentos mão na massa',
      'Carrinhos, polias, massas, elásticos, ímãs, fontes de tensão, ferramentas manuais'
    ]
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 Planos PEI')
    .addItem('📅 1. Preparar Nova Semana (PA Base ➔ PA Atual)', 'prepareNewWeek')
    .addItem('⚡ 2. Gerar Todos os Planos da Semana (PA Atual)', 'generateAllWeekPlans')
    .addItem('🎯 3. Gerar Plano APENAS da Aba Atual', 'generatePlanForActiveTab')
    .addSeparator()
    .addItem('🗂️ 4. Padronizar Nomes dos PDFs no Drive', 'standardizeDrivePdfNames')
    .addItem('🔑 5. Configurar Chave da API (Gemini) e Pastas', 'setupCredentialsDialog')
    .addItem('🧠 6. Alternar Modelo IA (Flash ⚡ / Pro 🧠)', 'toggleAiModel')
    .addToUi();
}

function getActiveModelName() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || CONFIG.DEFAULT_MODEL;
}

function toggleAiModel() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const current = getActiveModelName();
  const next = current === 'gemini-2.5-flash' ? 'gemini-2.5-pro' : 'gemini-2.5-flash';

  const resp = ui.alert(
    '🧠 Alternar Modelo de IA',
    `Modelo atual: ${current}\n\n` +
    `• gemini-2.5-flash: Muito rápido (~4s por aula), ideal para gerar a semana inteira de uma vez sem estourar limites.\n` +
    `• gemini-2.5-pro: Raciocínio mais profundo (~15s por aula), excelente para gerar aulas complexas individualmente (Aba Atual).\n\n` +
    `Deseja mudar para "${next}"?`,
    ui.ButtonSet.YES_NO
  );

  if (resp === ui.Button.YES) {
    props.setProperty('GEMINI_MODEL', next);
    ui.alert(`✅ Modelo alterado para: ${next}`);
  }
}

function setupCredentialsDialog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const currentKey = props.getProperty('GEMINI_API_KEY') || '';
  const respKey = ui.prompt(
    '🔑 Configurar Gemini API Key',
    'Insira sua chave da API do Google AI Studio (aistudio.google.com):\n' +
    (currentKey ? '(Já existe uma chave salva. Deixe em branco para manter a atual)' : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (respKey.getSelectedButton() !== ui.Button.OK) return;
  if (respKey.getResponseText().trim() !== '') {
    props.setProperty('GEMINI_API_KEY', respKey.getResponseText().trim());
  }

  const currentAulasUrl = props.getProperty('AULAS_FOLDER_ID') || '';
  const respAulas = ui.prompt(
    '📁 Pasta "Aulas" no Google Drive',
    'Cole o link ou ID da pasta principal "Aulas" (que contém Física 1, 2, 3 e Robótica 1, 2):\n' +
    (currentAulasUrl ? `(Atual: ${currentAulasUrl} - deixe em branco para manter)` : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (respAulas.getSelectedButton() === ui.Button.OK && respAulas.getResponseText().trim() !== '') {
    props.setProperty('AULAS_FOLDER_ID', extractDriveId(respAulas.getResponseText().trim()));
  }

  const currentGuiaUrl = props.getProperty('GUIA_FOLDER_ID') || '';
  const respGuia = ui.prompt(
    '📁 Pasta do "Guia do Currículo Priorizado"',
    'Cole o link ou ID da pasta onde fica o documento "Guia do Currículo Priorizado":\n' +
    (currentGuiaUrl ? `(Atual: ${currentGuiaUrl} - deixe em branco para manter)` : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (respGuia.getSelectedButton() === ui.Button.OK && respGuia.getResponseText().trim() !== '') {
    props.setProperty('GUIA_FOLDER_ID', extractDriveId(respGuia.getResponseText().trim()));
  }

  ui.alert('✅ Configurações salvas com sucesso!');
}

function extractDriveId(urlOrId) {
  const match = urlOrId.match(/[-\w]{25,}/);
  return match ? match[0] : urlOrId.trim();
}
