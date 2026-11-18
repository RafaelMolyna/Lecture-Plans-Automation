/**
 * ============================================================================
 * 2. MAIN.GS - ORQUESTRADOR DE SEMANA E GERAÇÃO DE PLANOS
 * ============================================================================
 */

function prepareNewWeek() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const baseSheet = ss.getSheetByName('PA (Base)');
  let atualSheet = ss.getSheetByName('PA (Atual)');

  if (!baseSheet) {
    ui.alert('❌ Erro: A aba "PA (Base)" não foi encontrada.');
    return;
  }

  const resp = ui.prompt(
    '📅 Preparar Nova Semana',
    'Digite a data da SEGUNDA-FEIRA da semana desejada (formato DD/MM/AAAA, ex: 16/11/2026):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const dateParts = resp.getResponseText().trim().split('/');
  if (dateParts.length !== 3) {
    ui.alert('❌ Formato inválido. Use DD/MM/AAAA (ex: 16/11/2026).');
    return;
  }

  const monday = new Date(
    parseInt(dateParts[2], 10),
    parseInt(dateParts[1], 10) - 1,
    parseInt(dateParts[0], 10),
    12, 0, 0
  );

  if (isNaN(monday.getTime())) {
    ui.alert('❌ Data inválida.');
    return;
  }

  if (!atualSheet) {
    atualSheet = baseSheet.copyTo(ss).setName('PA (Atual)');
  } else {
    atualSheet.clear();
    baseSheet.getDataRange().copyTo(atualSheet.getRange(1, 1));
  }

  const dayNames = ['SEG', 'TER', 'QUA', 'QUI', 'SEX'];
  const oneDayMs = 24 * 60 * 60 * 1000;
  for (let i = 0; i < 5; i++) {
    const d = new Date(monday.getTime() + (i * oneDayMs));
    const formatted = Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd/MM/yy');
    atualSheet.getRange(4, 2 + i).setValue(`${formatted}\n${dayNames[i]}`);
  }

  ss.setActiveSheet(atualSheet);
  ui.alert('✅ Semana preparada em "PA (Atual)"! Agora você pode gerar os planos.');
}

function generateAllWeekPlans() {
  runPlanGeneration(null);
}

function generatePlanForActiveTab() {
  const activeTabName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  if (!CONFIG.TAB_TO_FULLNAME[activeTabName]) {
    SpreadsheetApp.getUi().alert(
      `⚠️ Você está na aba "${activeTabName}".\n\n` +
      `Para usar esta opção, entre primeiro na aba de uma turma (ex: "1º HO", "2º LE", "3º AL" ou "Eletiva").`
    );
    return;
  }
  runPlanGeneration(activeTabName);
}

function runPlanGeneration(filterTabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    ui.alert('⚠️ Chave da API Gemini não configurada. Acesse "🤖 Planos PEI > 5. Configurar Chave da API" primeiro.');
    return;
  }

  const paSheet = ss.getSheetByName('PA (Atual)');
  if (!paSheet) {
    ui.alert('❌ Aba "PA (Atual)" não encontrada.');
    return;
  }

  const headerRange = paSheet.getRange(4, 2, 1, 5).getDisplayValues()[0];
  const colDates = headerRange.map(text => parseHeaderDate(text));

  const gridRange = paSheet.getRange(5, 2, 8, 5);
  const gridValues = gridRange.getDisplayValues();

  const blocks = [];
  for (let col = 0; col < 5; col++) {
    const dayDate = colDates[col];
    if (!dayDate) continue;

    let row = 0;
    while (row < 8) {
      const cellText = gridValues[row][col];
      const slotInfo = parsePaCellInfo(cellText);

      if (!slotInfo) {
        row++;
        continue;
      }

      if (filterTabName && slotInfo.tabName !== filterTabName) {
        row++;
        continue;
      }

      let isDouble = false;
      let secondCellInfo = null;
      if (row + 1 < 8) {
        const nextCellText = gridValues[row + 1][col];
        const nextSlotInfo = parsePaCellInfo(nextCellText);
        if (
          nextSlotInfo &&
          nextSlotInfo.className === slotInfo.className &&
          nextSlotInfo.subject === slotInfo.subject
        ) {
          isDouble = true;
          secondCellInfo = nextSlotInfo;
        }
      }

      const firstAlreadyFilled = slotInfo.hasModalidade;
      const secondAlreadyFilled = isDouble ? secondCellInfo.hasModalidade : true;

      if (!firstAlreadyFilled || !secondAlreadyFilled) {
        blocks.push({
          colIndex: col + 2,
          row1: row + 5,
          row2: isDouble ? row + 6 : null,
          isDouble: isDouble,
          isEletiva: Boolean(slotInfo.isEletiva),
          eletivaTitle: slotInfo.eletivaTitle || 'Experimentos de Física Mão na Massa',
          dateObj: dayDate,
          dateStr: Utilities.formatDate(dayDate, CONFIG.TIMEZONE, 'dd/MM/yyyy'),
          shortDateStr: formatPtShortDate(dayDate),
          className: slotInfo.className,
          tabName: slotInfo.tabName,
          grade: slotInfo.grade,
          noTv: slotInfo.noTv,
          subject: slotInfo.subject,
          cell1Text: cellText,
          cell2Text: isDouble ? gridValues[row + 1][col] : null
        });
      }

      row += isDouble ? 2 : 1;
    }
  }

  if (blocks.length === 0) {
    ui.alert('ℹ️ Todos os planos aplicáveis já possuem a MODALIDADE ORGANIZATIVA preenchida em "PA (Atual)"!');
    return;
  }

  const missingDrafts = [];
  const readyBlocks = [];

  blocks.forEach(block => {
    const draft = findDraftInClassSheet(ss, block.tabName, block.subject, block.dateObj, block.isEletiva);
    if (!draft || (!draft.summary && !draft.aulas)) {
      missingDrafts.push(`• ${block.shortDateStr} | ${block.tabName} (${block.subject})`);
    } else {
      block.draft = draft;
      readyBlocks.push(block);
    }
  });

  if (missingDrafts.length > 0) {
    const msg =
      `⚠️ Atenção: Os seguintes horários estão sem rascunho nas abas das turmas:\n\n` +
      missingDrafts.join('\n') +
      `\n\nDeseja gerar os ${readyBlocks.length} plano(s) disponíveis agora e deixar estes em branco?`;
    const confirm = ui.alert('Rascunhos Pendentes', msg, ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;
  }

  if (readyBlocks.length === 0) {
    ui.alert('⚠️ Nenhum rascunho encontrado para gerar planos.');
    return;
  }

  const educEspMap = loadEducEspData(ss);
  const guiaSummary = getPrioritizedCurriculumText(apiKey);
  const modelName = getActiveModelName();

  let successCount = 0;
  readyBlocks.forEach(block => {
    ss.toast(
      `Gerando (${modelName}): ${block.tabName} - ${block.subject} (${block.shortDateStr})...`,
      '🤖 Gemini trabalhando',
      25
    );

    const studentsEsp = educEspMap[block.tabName] || [];
    const pdfBlobs = block.isEletiva ? [] : resolveLecturePdfs(block);

    const aiResult = generateLessonPlanWithGemini({
      apiKey: apiKey,
      modelName: modelName,
      block: block,
      studentsEsp: studentsEsp,
      pdfBlobs: pdfBlobs,
      guiaSummary: guiaSummary
    });

    if (aiResult) {
      const updatedCell1 = buildFormattedCellText(block, aiResult.aula_1, studentsEsp);
      paSheet.getRange(block.row1, block.colIndex).setValue(updatedCell1);

      if (block.isDouble && block.row2) {
        const part2Data = aiResult.aula_2 || aiResult.aula_1;
        const updatedCell2 = buildFormattedCellText(block, part2Data, studentsEsp);
        paSheet.getRange(block.row2, block.colIndex).setValue(updatedCell2);
      }

      if (block.draft.wasDateBlank && block.draft.sheet && block.draft.rowNumber && block.draft.dateColIndex) {
        block.draft.sheet
          .getRange(block.draft.rowNumber, block.draft.dateColIndex)
          .setValue(block.shortDateStr);
      }

      successCount++;
    }
  });

  SpreadsheetApp.flush();
  ui.alert(`🎉 Concluído! ${successCount} bloco(s) de aula foram gerados e preenchidos em "PA (Atual)".`);
}