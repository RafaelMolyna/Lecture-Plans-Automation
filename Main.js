/**
 * ============================================================================
 * 2. MAIN.JS - ORQUESTRADOR PARALELO UNIVERSAL, PROGRESSO E RETRY (v4.0)
 * ============================================================================
 */

function openWeekPickerDialog() {
  const html = HtmlService.createHtmlOutputFromFile("WeekPicker")
    .setWidth(420)
    .setHeight(240);
  SpreadsheetApp.getUi().showModelessDialog(html, "📅 Preparar Nova Semana");
}

function applyPreparedWeek(isoDateStr) {
  const log = new ExecutionLogger("Preparar Nova Semana", { isoDateStr });
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    log.timeStart("Prepare_Week_Sheet");
    const baseSheet = ss.getSheetByName(CONFIG.TABS.PA_BASE);
    let atualSheet = ss.getSheetByName(CONFIG.TABS.PA_ATUAL);

    if (!baseSheet) {
      log.error(
        "Prepare_Week_Sheet",
        `Aba "${CONFIG.TABS.PA_BASE}" não encontrada.`,
      );
      throw new Error(`Aba "${CONFIG.TABS.PA_BASE}" não encontrada.`);
    }

    const parts = String(isoDateStr || "").split("-");
    if (parts.length !== 3) {
      log.error(
        "Prepare_Week_Sheet",
        "Formato de data inválido recebido do modal",
        { isoDateStr },
      );
      throw new Error("Formato de data inválido.");
    }

    const monday = new Date(
      parseInt(parts[0], 10),
      parseInt(parts[1], 10) - 1,
      parseInt(parts[2], 10),
      12,
      0,
      0,
    );
    if (isNaN(monday.getTime())) {
      log.error("Prepare_Week_Sheet", "Data inválida após parsing", {
        isoDateStr,
      });
      throw new Error("Data inválida.");
    }

    if (!atualSheet) {
      atualSheet = baseSheet.copyTo(ss).setName(CONFIG.TABS.PA_ATUAL);
      log.info(
        "Prepare_Week_Sheet",
        `Aba "${CONFIG.TABS.PA_ATUAL}" criada a partir de "${CONFIG.TABS.PA_BASE}".`,
      );
    } else {
      atualSheet.clear();
      baseSheet.getDataRange().copyTo(atualSheet.getRange(1, 1));
      log.info(
        "Prepare_Week_Sheet",
        `Aba "${CONFIG.TABS.PA_ATUAL}" resetada a partir de "${CONFIG.TABS.PA_BASE}".`,
      );
    }

    const layout = detectPaGridLayout(atualSheet);
    const dayNames = ["SEG", "TER", "QUA", "QUI", "SEX"];
    const oneDayMs = 24 * 60 * 60 * 1000;
    const updatedHeaders = [];

    for (let i = 0; i < 5; i++) {
      const d = new Date(monday.getTime() + i * oneDayMs);
      const formatted = Utilities.formatDate(d, CONFIG.TIMEZONE, "dd/MM/yy");
      const label = `${formatted}\n${dayNames[i]}`;
      atualSheet.getRange(layout.dateRow, 2 + i).setValue(label);
      updatedHeaders.push(label.replace("\n", " "));
    }

    ss.setActiveSheet(atualSheet);
    log.timeEnd("Prepare_Week_Sheet", {
      dateRow: layout.dateRow,
      headers: updatedHeaders,
    });
    log.checkpoint("WEEK_PREPARED", { monday: isoDateStr });
    ss.toast(
      `Semana de ${parts[2]}/${parts[1]}/${parts[0]} preparada em "${CONFIG.TABS.PA_ATUAL}"!`,
      "✅ Concluído",
      5,
    );
  } catch (err) {
    log.error("Prepare_Week_Sheet", "Falha ao preparar nova semana", err, {
      isoDateStr,
    });
    throw err;
  } finally {
    log.flushToStorage();
  }
}

function generateAllWeekPlans() {
  runPlanGeneration(null);
}

function generatePlanForActiveTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeTabName = ss.getActiveSheet().getName();
  const runtimeConfig = loadRuntimeConfig(ss, null);

  if (!runtimeConfig.classesByTab[normalizeKey(activeTabName)]) {
    const validTabs = runtimeConfig.classesList
      .map((c) => `"${c.tab}"`)
      .join(", ");
    SpreadsheetApp.getUi().alert(
      `⚠️ Você está na aba "${activeTabName}".\n\n` +
        `Entre primeiro na aba de uma turma cadastrada em "${CONFIG.TABS.TURMAS}": ${validTabs}.`,
    );
    return;
  }
  runPlanGeneration(activeTabName);
}

function updateProgressState(percent, stage, detail, success, errors, done) {
  const payload = JSON.stringify({
    percent: Math.min(100, Math.max(0, Math.round(percent))),
    stage: stage || "",
    detail: detail || "",
    success: success || 0,
    errors: errors || 0,
    done: Boolean(done),
  });
  CacheService.getScriptCache().put("PEI_PROGRESS_STATE", payload, 600);
}

function getGenerationProgressState() {
  const raw = CacheService.getScriptCache().get("PEI_PROGRESS_STATE");
  return raw ? JSON.parse(raw) : null;
}

function runPlanGeneration(filterTabName) {
  const opLabel = filterTabName
    ? `Gerar Plano (${filterTabName})`
    : "Gerar Todos os Planos da Semana";
  const log = new ExecutionLogger(opLabel, {
    filterTabName: filterTabName || "ALL",
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("GEMINI_API_KEY");

  let successCount = 0;
  let errorCount = 0;

  try {
    if (!apiKey) {
      log.error(
        "AUTH_CONFIG",
        "GEMINI_API_KEY não configurada em ScriptProperties.",
      );
      ui.alert(
        '⚠️ Chave da API Gemini não configurada. Acesse "🤖 Planos PEI > ⚙️ Ferramentas e Configs > 🔑 Configurar Chave da API".',
      );
      return;
    }

    const paSheet = ss.getSheetByName(CONFIG.TABS.PA_ATUAL);
    if (!paSheet) {
      log.error(
        "SHEET_LOOKUP",
        `Aba "${CONFIG.TABS.PA_ATUAL}" não encontrada na planilha.`,
      );
      ui.alert(`❌ Aba "${CONFIG.TABS.PA_ATUAL}" não encontrada.`);
      return;
    }

    updateProgressState(
      5,
      "1/4 • Carregando Configs e PA",
      "Lendo abas Turmas, Disciplinas e PA (Atual)...",
      0,
      0,
      false,
    );
    ui.showSidebar(
      HtmlService.createHtmlOutputFromFile("ProgressSidebar").setTitle(
        "Progresso - Planos PEI",
      ),
    );

    const runtimeConfig = loadRuntimeConfig(ss, log);

    log.timeStart("Scan_PA_Atual", { sheet: CONFIG.TABS.PA_ATUAL });
    const layout = detectPaGridLayout(paSheet);
    const headerRange = paSheet
      .getRange(layout.dateRow, 2, 1, 5)
      .getDisplayValues()[0];
    const colDates = headerRange.map((text) => parseHeaderDate(text));
    const gridValues = paSheet
      .getRange(layout.startPeriodRow, 2, layout.numPeriodRows, 5)
      .getDisplayValues();

    let skippedFilledCount = 0;
    let ignoredNonLessonCells = 0;
    const blocks = [];

    for (let col = 0; col < 5; col++) {
      const dayDate = colDates[col];
      if (!dayDate) {
        log.warn(
          "Scan_PA_Atual",
          `Coluna ${col + 2} sem data válida no cabeçalho`,
          { rawHeader: headerRange[col] },
        );
        continue;
      }

      let row = 0;
      while (row < layout.numPeriodRows) {
        const cellText = gridValues[row][col];
        const slotInfo = parsePaCellInfo(cellText, runtimeConfig);

        if (!slotInfo) {
          ignoredNonLessonCells++;
          row++;
          continue;
        }

        if (
          filterTabName &&
          normalizeKey(slotInfo.tabName) !== normalizeKey(filterTabName)
        ) {
          row++;
          continue;
        }

        let isDouble = false;
        let secondCellInfo = null;
        if (row + 1 < layout.numPeriodRows) {
          const nextSlotInfo = parsePaCellInfo(
            gridValues[row + 1][col],
            runtimeConfig,
          );
          if (
            nextSlotInfo &&
            nextSlotInfo.className === slotInfo.className &&
            normalizeKey(nextSlotInfo.subject) ===
              normalizeKey(slotInfo.subject)
          ) {
            isDouble = true;
            secondCellInfo = nextSlotInfo;
          }
        }

        const firstAlreadyFilled = slotInfo.hasModalidade;
        const secondAlreadyFilled = isDouble
          ? secondCellInfo.hasModalidade
          : true;

        if (firstAlreadyFilled && secondAlreadyFilled) {
          skippedFilledCount++;
          log.info("Scan_PA_Atual", "Bloco já preenchido, ignorado", {
            tab: slotInfo.tabName,
            subject: slotInfo.subject,
            date: formatPtShortDate(dayDate),
            isDouble,
          });
        } else {
          blocks.push({
            colIndex: col + 2,
            row1: layout.startPeriodRow + row,
            row2: isDouble ? layout.startPeriodRow + row + 1 : null,
            isDouble: isDouble,
            isSimplified: slotInfo.isSimplified,
            customTitle: slotInfo.customTitle,
            dateObj: dayDate,
            dateStr: Utilities.formatDate(
              dayDate,
              CONFIG.TIMEZONE,
              "dd/MM/yyyy",
            ),
            shortDateStr: formatPtShortDate(dayDate),
            className: slotInfo.className,
            tabName: slotInfo.tabName,
            grade: slotInfo.grade,
            noTv: slotInfo.noTv,
            subject: slotInfo.subject,
          });
        }

        row += isDouble ? 2 : 1;
      }
    }

    log.timeEnd("Scan_PA_Atual", {
      dateRow: layout.dateRow,
      periodRows: layout.numPeriodRows,
      pendingBlocks: blocks.length,
      skippedFilledCount,
      ignoredNonLessonCells,
    });
    log.checkpoint("PA_SCANNED", {
      pendingBlocks: blocks.length,
      skippedFilledCount,
    });

    if (blocks.length === 0) {
      updateProgressState(
        100,
        "Tudo pronto!",
        "Todos os planos aplicáveis já estavam preenchidos.",
        0,
        0,
        true,
      );
      log.info("RUN_END", "Nenhum bloco pendente encontrado em PA (Atual).");
      ui.alert(
        "ℹ️ Todos os planos aplicáveis já possuem a MODALIDADE ORGANIZATIVA preenchida!",
      );
      return;
    }

    log.timeStart("Draft_Matching", { blockCount: blocks.length });
    const missingDrafts = [];
    const readyBlocks = [];

    blocks.forEach((block) => {
      const draft = findDraftInClassSheet(
        ss,
        block.tabName,
        block.subject,
        block.dateObj,
        runtimeConfig,
        log,
      );
      if (!draft || (!draft.summary && !draft.aulas)) {
        missingDrafts.push(
          `• ${block.shortDateStr} | ${block.tabName} (${block.subject})`,
        );
      } else {
        block.draft = draft;
        readyBlocks.push(block);
      }
    });
    log.timeEnd("Draft_Matching", {
      readyCount: readyBlocks.length,
      missingCount: missingDrafts.length,
    });

    updateProgressState(
      15,
      "2/4 • Rascunhos Localizados",
      `${readyBlocks.length} bloco(s) prontos para geração.`,
      0,
      0,
      false,
    );

    if (missingDrafts.length > 0) {
      log.warn(
        "Draft_Matching",
        "Rascunhos ausentes para alguns blocos",
        missingDrafts,
      );
      const msg =
        `⚠️ Atenção: Os seguintes horários estão sem rascunho:\n\n` +
        missingDrafts.join("\n") +
        `\n\nDeseja gerar os ${readyBlocks.length} plano(s) disponíveis agora e deixar estes em branco?`;
      if (
        ui.alert("Rascunhos Pendentes", msg, ui.ButtonSet.YES_NO) !==
        ui.Button.YES
      ) {
        updateProgressState(
          100,
          "Cancelado",
          "Execução cancelada pelo usuário.",
          0,
          0,
          true,
        );
        log.info(
          "RUN_END",
          "Cancelado pelo usuário no diálogo de rascunhos pendentes.",
        );
        return;
      }
    }

    if (readyBlocks.length === 0) {
      updateProgressState(
        100,
        "Sem rascunhos",
        "Nenhum rascunho encontrado para gerar planos.",
        0,
        0,
        true,
      );
      log.warn("RUN_END", "Zero blocos prontos após validação de rascunhos.");
      ui.alert("⚠️ Nenhum rascunho encontrado para gerar planos.");
      return;
    }

    log.timeStart("Load_EducEsp_And_Guia");
    const educEspMap = loadEducEspData(ss, runtimeConfig, log);
    const guiaSummary = getPrioritizedCurriculumText(apiKey, log);
    log.timeEnd("Load_EducEsp_And_Guia");

    const modelName = getActiveModelName();
    const pdfBase64Cache = {};

    log.timeStart("Drive_And_Payload_Build", {
      modelName: modelName,
      readyBlocks: readyBlocks.length,
    });
    log.checkpoint("BUILDING_PAYLOADS", { modelName: modelName });
    ss.toast(
      `Preparando ${readyBlocks.length} planos em paralelo (${modelName})...`,
      "⚡ Acelerando",
      15,
    );

    const preparedItems = readyBlocks.map((block, idx) => {
      const pct = 20 + (idx / readyBlocks.length) * 25;
      updateProgressState(
        pct,
        `3/4 • Lendo Drive (${idx + 1}/${readyBlocks.length})`,
        `${block.tabName} (${block.subject}) - ${block.shortDateStr}`,
        0,
        0,
        false,
      );

      const studentsEsp = educEspMap[block.tabName] || [];
      const subjectMeta = getSubjectMeta(
        runtimeConfig,
        block.subject,
        block.grade,
      );

      log.info(
        "Educ_Esp_Eval",
        `Acessibilidade avaliada para ${block.tabName} (${block.subject})`,
        {
          matchedStudents: studentsEsp.map((s) => s.name),
          resultMode: studentsEsp.length > 0 ? "ADAPTED" : "N/A",
        },
      );

      const encodedPdfs = resolveLecturePdfsCached(
        block,
        subjectMeta,
        pdfBase64Cache,
        log,
      );
      const requestObj = buildGeminiRequestPayload({
        apiKey: apiKey,
        modelName: modelName,
        block: block,
        subjectMeta: subjectMeta,
        studentsEsp: studentsEsp,
        encodedPdfs: encodedPdfs,
        guiaSummary: guiaSummary,
        log: log,
      });
      return { block: block, studentsEsp: studentsEsp, request: requestObj };
    });
    log.timeEnd("Drive_And_Payload_Build", {
      cachedPdfKeys: Object.keys(pdfBase64Cache).length,
    });

    function processItemsInBatches(items, labelPrefix) {
      const retryQueue = [];
      const batchesCount = Math.ceil(items.length / CONFIG.BATCH_SIZE);

      for (let i = 0; i < items.length; i += CONFIG.BATCH_SIZE) {
        const batch = items.slice(i, i + CONFIG.BATCH_SIZE);
        const batchNum = Math.floor(i / CONFIG.BATCH_SIZE) + 1;
        const stepName = `${labelPrefix}_B${batchNum}_of_${batchesCount}`;

        const batchPct = 45 + ((i + batch.length) / items.length) * 50;
        updateProgressState(
          batchPct,
          `4/4 • IA (${labelPrefix} ${batchNum}/${batchesCount})`,
          `Processando ${batch.length} bloco(s) em paralelo...`,
          successCount,
          errorCount,
          false,
        );

        log.checkpoint(`${stepName}_START`, { count: batch.length });
        ss.toast(
          `${labelPrefix}: lote ${batchNum}/${batchesCount} (${batch.length} aula(s))...`,
          "🚀 Gemini API",
          20,
        );
        SpreadsheetApp.flush();

        log.timeStart(stepName, { batchNum: batchNum, count: batch.length });
        const responses = UrlFetchApp.fetchAll(batch.map((it) => it.request));
        log.timeEnd(stepName, { batchNum: batchNum });

        responses.forEach((resp, idx) => {
          const item = batch[idx];
          const block = item.block;
          const code = resp.getResponseCode();
          const body = resp.getContentText();

          if (code === 200) {
            const parsed = parseGeminiJsonResponse(body, block, log);
            if (parsed && parsed.aula_1) {
              paSheet
                .getRange(block.row1, block.colIndex)
                .setRichTextValue(
                  buildFormattedCellRichText(
                    block,
                    parsed.aula_1,
                    item.studentsEsp,
                    log,
                  ),
                );
              if (block.isDouble && block.row2) {
                const p2 = parsed.aula_2 || parsed.aula_1;
                paSheet
                  .getRange(block.row2, block.colIndex)
                  .setRichTextValue(
                    buildFormattedCellRichText(
                      block,
                      p2,
                      item.studentsEsp,
                      log,
                    ),
                  );
              }
              if (
                block.draft.wasDateBlank &&
                block.draft.sheet &&
                block.draft.rowNumber &&
                block.draft.dateColIndex
              ) {
                block.draft.sheet
                  .getRange(block.draft.rowNumber, block.draft.dateColIndex)
                  .setValue(block.shortDateStr);
              }
              successCount++;
              log.info(
                "Block_Success",
                `${block.tabName} (${block.subject}) - ${block.shortDateStr}`,
              );
            } else {
              errorCount++;
              log.error(
                "Block_Parse_Fail",
                `JSON inválido em ${block.tabName} (${block.subject})`,
                { preview: body.slice(0, 250) },
              );
            }
          } else if (code === 429 || code === 503) {
            log.warn(
              "Retry_Queue",
              `HTTP ${code} em ${block.tabName} (${block.subject}); enfileirado para retry.`,
            );
            retryQueue.push({ item: item, lastCode: code });
          } else {
            errorCount++;
            log.error(
              "Gemini_HTTP_Error",
              `HTTP ${code} em ${block.tabName} (${block.subject})`,
              { preview: body.slice(0, 300) },
            );
          }
        });

        SpreadsheetApp.flush();
        log.checkpoint(`${stepName}_DONE`, {
          successCount,
          errorCount,
          queuedForRetry: retryQueue.length,
        });
      }

      return retryQueue;
    }

    let pendingRetry = processItemsInBatches(preparedItems, "Batch");

    for (
      let attempt = 1;
      attempt <= CONFIG.RETRY_ATTEMPTS && pendingRetry.length > 0;
      attempt++
    ) {
      const backoffMs = attempt * 4000;
      log.warn(
        "Retry_Backoff",
        `Aguardando ${backoffMs / 1000}s antes da tentativa de retry #${attempt}`,
        { retryCount: pendingRetry.length },
      );
      updateProgressState(
        92,
        `⏳ Retry #${attempt}/${CONFIG.RETRY_ATTEMPTS}`,
        `Aguardando ${backoffMs / 1000}s para reenviar ${pendingRetry.length} aula(s) (HTTP 429/503)...`,
        successCount,
        errorCount,
        false,
      );
      ss.toast(
        `Aguardando ${backoffMs / 1000}s para tentar novamente ${pendingRetry.length} item(ns) (HTTP 429/503)...`,
        "⏳ Retry",
        10,
      );
      SpreadsheetApp.flush();
      Utilities.sleep(backoffMs);

      const itemsToRetry = pendingRetry.map((r) => r.item);
      pendingRetry = processItemsInBatches(itemsToRetry, `Retry_${attempt}`);
    }

    pendingRetry.forEach((r) => {
      errorCount++;
      log.error(
        "Retry_Exhausted",
        `Falha persistente (HTTP ${r.lastCode}) após retries em ${r.item.block.tabName} (${r.item.block.subject})`,
      );
    });

    log.timeStart("Spreadsheet_Flush");
    SpreadsheetApp.flush();
    log.timeEnd("Spreadsheet_Flush");
    log.checkpoint("COMPLETED", { successCount, errorCount });

    updateProgressState(
      100,
      errorCount > 0 ? "Concluído com ressalvas" : "Concluído!",
      `${successCount} gerado(s), ${errorCount} falha(s).`,
      successCount,
      errorCount,
      true,
    );

    if (errorCount > 0) {
      ui.alert(
        `⚠️ Concluído com ressalvas!\n\n` +
          `• Gerados com sucesso: ${successCount}\n` +
          `• Falhas: ${errorCount}\n\n` +
          `Abra "🤖 Planos PEI > ⚙️ Ferramentas e Configs > 📋 Ver Logs da Última Execução" para ver o gargalo e o erro exato.`,
      );
    } else {
      ui.alert(
        `🎉 Concluído com sucesso! Todos os ${successCount} bloco(s) foram gerados em paralelo.`,
      );
    }
  } catch (err) {
    updateProgressState(
      100,
      "Erro Crítico",
      `Falha na etapa "${log.lastCheckpoint}": ${err.message}`,
      successCount,
      errorCount + 1,
      true,
    );
    log.error(
      "UNHANDLED_CRASH",
      `Crash na etapa "${log.lastCheckpoint}"`,
      err,
      {
        filterTabName: filterTabName || "ALL",
      },
    );
    ui.alert(
      `❌ Erro crítico na etapa "${log.lastCheckpoint}": ${err.message}\n\n` +
        `Abra "🤖 Planos PEI > ⚙️ Ferramentas e Configs > 📋 Ver Logs da Última Execução" para ver o stack trace e os tempos.`,
    );
  } finally {
    log.flushToStorage();
  }
}
