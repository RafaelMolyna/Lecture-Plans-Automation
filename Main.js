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
        '⚠️ Chave da API Gemini não configurada. Acesse "🤖 Planos PEI > 6. Configurar Chave da API".',
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

        if (
          !slotInfo ||
          (filterTabName && slotInfo.tabName !== filterTabName)
        ) {
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

        if (!firstAlreadyFilled || !secondAlreadyFilled) {
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
    log.timeEnd("Scan_PA_Atual", { pendingBlocks: blocks.length });
    log.checkpoint("PA_SCANNED", { pendingBlocks: blocks.length });

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
    const educEspMap = loadEducEspData(ss);
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
    const totalBatches = Math.ceil(preparedItems.length / CONFIG.BATCH_SIZE);

    for (let i = 0; i < preparedItems.length; i += CONFIG.BATCH_SIZE) {
      const batch = preparedItems.slice(i, i + CONFIG.BATCH_SIZE);
      const batchNum = Math.floor(i / CONFIG.BATCH_SIZE) + 1;
      const batchLabel = `Batch_${batchNum}_of_${totalBatches}`;

      log.checkpoint(`BATCH_${batchNum}_START`, { batchSize: batch.length });
      ss.toast(
        `Processando lote paralelo ${batchNum}/${totalBatches} (${batch.length} aulas)...`,
        "🚀 Gemini API",
        25,
      );

      log.timeStart(batchLabel, { batchNum: batchNum, count: batch.length });
      const httpRequests = batch.map((item) => item.request);
      const httpResponses = UrlFetchApp.fetchAll(httpRequests);
      log.timeEnd(batchLabel, { batchNum: batchNum });

      log.timeStart(`Write_Batch_${batchNum}`);
      httpResponses.forEach((resp, idx) => {
        const item = batch[idx];
        const block = item.block;
        const statusCode = resp.getResponseCode();
        const respText = resp.getContentText();

        if (statusCode === 200) {
          const parsedPlan = parseGeminiJsonResponse(respText, block, log);
          if (parsedPlan && parsedPlan.aula_1) {
            const updatedCell1 = buildFormattedCellText(
              block,
              parsedPlan.aula_1,
              item.studentsEsp,
            );
            paSheet.getRange(block.row1, block.colIndex).setValue(updatedCell1);

            if (block.isDouble && block.row2) {
              const part2Data = parsedPlan.aula_2 || parsedPlan.aula_1;
              const updatedCell2 = buildFormattedCellText(
                block,
                part2Data,
                item.studentsEsp,
              );
              paSheet
                .getRange(block.row2, block.colIndex)
                .setValue(updatedCell2);
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
              log.info(
                "Auto_Date",
                `Data ${block.shortDateStr} preenchida em ${block.tabName} linha ${block.draft.rowNumber}`,
              );
            }

            successCount++;
            log.info(
              "Block_Success",
              `${block.tabName} (${block.subject}) - ${block.shortDateStr}`,
              {
                isDouble: block.isDouble,
                hasPart2: Boolean(parsedPlan.aula_2),
              },
            );
          } else {
            errorCount++;
            log.error(
              "Block_Parse_Fail",
              `JSON incompleto/inválido para ${block.tabName} (${block.subject})`,
              {
                date: block.shortDateStr,
                hasAula1: Boolean(parsedPlan && parsedPlan.aula_1),
                rawPreview: respText.slice(0, 300),
              },
            );
          }
        } else {
          errorCount++;
          log.error(
            "Gemini_HTTP_Error",
            `HTTP ${statusCode} em ${block.tabName} (${block.subject})`,
            {
              date: block.shortDateStr,
              model: modelName,
              bodyPreview: respText.slice(0, 400),
            },
          );
        }
      });
      log.timeEnd(`Write_Batch_${batchNum}`);
      log.checkpoint(`BATCH_${batchNum}_DONE`, {
        successSoFar: successCount,
        errorsSoFar: errorCount,
      });
    }

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
          `Abra "🤖 Planos PEI > 4. Ver Logs da Última Execução" para ver o gargalo e o erro exato.`,
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
        `Abra "🤖 Planos PEI > 4. Ver Logs da Última Execução" para ver o stack trace e os tempos.`,
    );
  } finally {
    log.flushToStorage();
  }
}
