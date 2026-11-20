/**
 * ============================================================================
 * 2. MAIN.JS - ORQUESTRADOR PARALELO, SELETOR DE SEMANA E LOGS (v3.1)
 * ============================================================================
 */

function openWeekPickerDialog() {
  const html = HtmlService.createHtmlOutputFromFile("WeekPicker")
    .setWidth(420)
    .setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(html, "📅 Preparar Nova Semana");
}

function applyPreparedWeek(isoDateStr) {
  const log = new ExecutionLogger("Preparar Nova Semana", { isoDateStr });
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    log.timeStart("Prepare_Week_Sheet");
    const baseSheet = ss.getSheetByName("PA (Base)");
    let atualSheet = ss.getSheetByName("PA (Atual)");

    if (!baseSheet) {
      log.error("Prepare_Week_Sheet", 'Aba "PA (Base)" não encontrada.');
      throw new Error('Aba "PA (Base)" não encontrada.');
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
      atualSheet = baseSheet.copyTo(ss).setName("PA (Atual)");
      log.info(
        "Prepare_Week_Sheet",
        'Aba "PA (Atual)" criada a partir de "PA (Base)".',
      );
    } else {
      atualSheet.clear();
      baseSheet.getDataRange().copyTo(atualSheet.getRange(1, 1));
      log.info(
        "Prepare_Week_Sheet",
        'Aba "PA (Atual)" resetada a partir de "PA (Base)".',
      );
    }

    const dayNames = ["SEG", "TER", "QUA", "QUI", "SEX"];
    const oneDayMs = 24 * 60 * 60 * 1000;
    const updatedHeaders = [];
    for (let i = 0; i < 5; i++) {
      const d = new Date(monday.getTime() + i * oneDayMs);
      const formatted = Utilities.formatDate(d, CONFIG.TIMEZONE, "dd/MM/yy");
      const label = `${formatted}\n${dayNames[i]}`;
      atualSheet.getRange(4, 2 + i).setValue(label);
      updatedHeaders.push(label.replace("\n", " "));
    }

    ss.setActiveSheet(atualSheet);
    log.timeEnd("Prepare_Week_Sheet", { headers: updatedHeaders });
    log.checkpoint("WEEK_PREPARED", { monday: isoDateStr });
    ss.toast(
      `Semana de ${parts[2]}/${parts[1]}/${parts[0]} preparada em "PA (Atual)"!`,
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
  const activeTabName = SpreadsheetApp.getActiveSpreadsheet()
    .getActiveSheet()
    .getName();
  if (!CONFIG.TAB_TO_FULLNAME[activeTabName]) {
    SpreadsheetApp.getUi().alert(
      `⚠️ Você está na aba "${activeTabName}".\n\n` +
        `Entre primeiro na aba de uma turma (ex: "1º HO", "2º LE", "3º AL" ou "Eletiva").`,
    );
    return;
  }
  runPlanGeneration(activeTabName);
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

    const paSheet = ss.getSheetByName("PA (Atual)");
    if (!paSheet) {
      log.error("SHEET_LOOKUP", 'Aba "PA (Atual)" não encontrada na planilha.');
      ui.alert('❌ Aba "PA (Atual)" não encontrada.');
      return;
    }

    log.timeStart("Scan_PA_Atual", { sheet: "PA (Atual)" });
    const headerRange = paSheet.getRange(4, 2, 1, 5).getDisplayValues()[0];
    const colDates = headerRange.map((text) => parseHeaderDate(text));
    const gridRange = paSheet.getRange(5, 2, 8, 5);
    const gridValues = gridRange.getDisplayValues();

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
      while (row < 8) {
        const cellText = gridValues[row][col];
        const slotInfo = parsePaCellInfo(cellText);

        if (!slotInfo) {
          ignoredNonLessonCells++;
          row++;
          continue;
        }

        if (filterTabName && slotInfo.tabName !== filterTabName) {
          log.info("Scan_PA_Atual", `Célula ignorada por filtro de aba ativa`, {
            cell: `R${row + 5}C${col + 2}`,
            slotTab: slotInfo.tabName,
            filterTabName,
          });
          row++;
          continue;
        }

        let isDouble = false;
        let secondCellInfo = null;
        if (row + 1 < 8) {
          const nextSlotInfo = parsePaCellInfo(gridValues[row + 1][col]);
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
        const secondAlreadyFilled = isDouble
          ? secondCellInfo.hasModalidade
          : true;

        if (firstAlreadyFilled && secondAlreadyFilled) {
          skippedFilledCount++;
          log.info(
            "Scan_PA_Atual",
            `Bloco já preenchido (MODALIDADE ORGANIZATIVA existente), ignorado`,
            {
              tab: slotInfo.tabName,
              subject: slotInfo.subject,
              date: formatPtShortDate(dayDate),
              isDouble,
            },
          );
        } else {
          blocks.push({
            colIndex: col + 2,
            row1: row + 5,
            row2: isDouble ? row + 6 : null,
            isDouble: isDouble,
            isEletiva: Boolean(slotInfo.isEletiva),
            eletivaTitle:
              slotInfo.eletivaTitle || "Experimentos de Física Mão na Massa",
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
      pendingBlocks: blocks.length,
      skippedFilledCount,
      ignoredNonLessonCells,
    });
    log.checkpoint("PA_SCANNED", {
      pendingBlocks: blocks.length,
      skippedFilledCount,
    });

    if (blocks.length === 0) {
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
        block.isEletiva,
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
        log.info(
          "RUN_END",
          "Cancelado pelo usuário no diálogo de rascunhos pendentes.",
        );
        return;
      }
    }

    if (readyBlocks.length === 0) {
      log.warn("RUN_END", "Zero blocos prontos após validação de rascunhos.");
      ui.alert("⚠️ Nenhum rascunho encontrado para gerar planos.");
      return;
    }

    log.timeStart("Load_EducEsp_And_Guia");
    const educEspMap = loadEducEspData(ss, log);
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

    const preparedItems = readyBlocks.map((block) => {
      const studentsEsp = educEspMap[block.tabName] || [];
      log.info(
        "Educ_Esp_Eval",
        `Acessibilidade avaliada para ${block.tabName} (${block.subject})`,
        {
          matchedStudents: studentsEsp.map((s) => s.name),
          resultMode: studentsEsp.length > 0 ? "ADAPTED" : "N/A",
        },
      );

      const encodedPdfs = block.isEletiva
        ? []
        : resolveLecturePdfsCached(block, pdfBase64Cache, log);
      const requestObj = buildGeminiRequestPayload({
        apiKey: apiKey,
        modelName: modelName,
        block: block,
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

    let successCount = 0;
    let errorCount = 0;

    function processItemsInBatches(items, labelPrefix, log) {
      const retryQueue = [];
      const batchesCount = Math.ceil(items.length / CONFIG.BATCH_SIZE);

      for (let i = 0; i < items.length; i += CONFIG.BATCH_SIZE) {
        const batch = items.slice(i, i + CONFIG.BATCH_SIZE);
        const batchNum = Math.floor(i / CONFIG.BATCH_SIZE) + 1;
        const stepName = `${labelPrefix}_B${batchNum}_of_${batchesCount}`;

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

    let pendingRetry = processItemsInBatches(preparedItems, "Batch", log);

    for (
      let attempt = 1;
      attempt <= CONFIG.RETRY_ATTEMPTS && pendingRetry.length > 0;
      attempt++
    ) {
      const backoffMs = attempt * 4000;
      log.warn(
        "Retry_Backoff",
        `Aguardando ${backoffMs / 1000}s antes da tentativa de retry #${attempt}`,
        {
          retryCount: pendingRetry.length,
        },
      );
      ss.toast(
        `Aguardando ${backoffMs / 1000}s para tentar novamente ${pendingRetry.length} item(ns) (HTTP 429/503)...`,
        "⏳ Retry",
        10,
      );
      SpreadsheetApp.flush();
      Utilities.sleep(backoffMs);

      const itemsToRetry = pendingRetry.map((r) => r.item);
      pendingRetry = processItemsInBatches(
        itemsToRetry,
        `Retry_${attempt}`,
        log,
      );
    }

    // Qualquer item que ainda falhou com 429/503 após todas as tentativas entra na contagem final de erro
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
    log.checkpoint("COMPLETED", {
      successCount: successCount,
      errorCount: errorCount,
    });

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
