/**
 * ============================================================================
 * 3. SHEETPARSER.JS - CONFIG DINÂMICA, PA SIMPLIFICADO, RASCUNHOS E EDUC. ESP.
 * ============================================================================
 */

function normalizeKey(str) {
  return String(str || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

/**
 * Carrega dinamicamente as abas "Turmas" e "Disciplinas".
 */
function loadRuntimeConfig(ss, log) {
  const turmasSheet = ss.getSheetByName(CONFIG.TABS.TURMAS);
  const discSheet = ss.getSheetByName(CONFIG.TABS.DISCIPLINAS);

  if (!turmasSheet) {
    throw new Error(`Aba "${CONFIG.TABS.TURMAS}" não encontrada na planilha.`);
  }
  if (!discSheet) {
    throw new Error(
      `Aba "${CONFIG.TABS.DISCIPLINAS}" não encontrada na planilha.`,
    );
  }

  const classesList = [];
  const classesByTab = {};
  const classesByFullName = {};

  const turmasRows = turmasSheet.getDataRange().getDisplayValues();
  for (let i = 1; i < turmasRows.length; i++) {
    const fullName = String(turmasRows[i][0] || "").trim();
    const tab = String(turmasRows[i][1] || "").trim();
    const grade = parseInt(turmasRows[i][2], 10) || 0;
    const temTvStr = String(turmasRows[i][3] || "").trim();
    const tipoPlano = String(turmasRows[i][4] || "Regular").trim();

    if (!fullName && !tab) continue;

    const classObj = {
      fullName: fullName || tab,
      tab: tab || fullName,
      grade: grade,
      noTv: /^(n[aã]o|nao|false|0)$/i.test(temTvStr),
      isSimplified: /simplific/i.test(tipoPlano),
      planType: tipoPlano,
    };

    classesList.push(classObj);
    classesByTab[normalizeKey(classObj.tab)] = classObj;
    classesByFullName[normalizeKey(classObj.fullName)] = classObj;
  }

  const subjectsList = [];
  const subjectsByKey = {};
  const subjectsByName = {};

  const discRows = discSheet.getDataRange().getDisplayValues();
  for (let i = 1; i < discRows.length; i++) {
    const name = String(discRows[i][0] || "").trim();
    const grade = parseInt(discRows[i][1], 10) || 0;
    const rawFolder = String(discRows[i][2] || "").trim();
    const lessonsPerPdf = Math.max(1, parseInt(discRows[i][3], 10) || 1);
    const materials = String(discRows[i][4] || "").trim();
    const evalCriteria = String(discRows[i][5] || "").trim();

    if (!name) continue;

    const cleanFolder =
      !rawFolder || /^(\(vazio\)|vazio|-|n\/a)$/i.test(rawFolder)
        ? null
        : rawFolder;

    const subjObj = {
      name: name,
      grade: grade,
      driveFolder: cleanFolder,
      lessonsPerPdf: lessonsPerPdf,
      materials: materials,
      evalCriteria: evalCriteria,
    };

    const key = `${normalizeKey(name)}_${grade}`;
    subjectsList.push(subjObj);
    subjectsByKey[key] = subjObj;
    if (!subjectsByName[normalizeKey(name)]) {
      subjectsByName[normalizeKey(name)] = subjObj;
    }
  }

  const defaultSubject =
    subjectsList.length > 0 ? subjectsList[0].name : "Física";

  if (log) {
    log.info(
      "Config_Load",
      "Configuração carregada das abas Turmas e Disciplinas",
      {
        turmasCount: classesList.length,
        disciplinasCount: subjectsList.length,
        defaultSubject,
      },
    );
  }

  return {
    classesList,
    classesByTab,
    classesByFullName,
    subjectsList,
    subjectsByKey,
    subjectsByName,
    defaultSubject,
  };
}

function getSubjectMeta(runtimeConfig, subjectName, grade) {
  const norm = normalizeKey(subjectName);
  return (
    runtimeConfig.subjectsByKey[`${norm}_${grade}`] ||
    runtimeConfig.subjectsByName[norm] || {
      name: subjectName,
      grade: grade,
      driveFolder: null,
      lessonsPerPdf: 1,
      materials: "",
      evalCriteria: "",
    }
  );
}

/**
 * Localiza dinamicamente a linha de datas (SEG..SEX) e a faixa de aulas no PA.
 */
function detectPaGridLayout(paSheet) {
  const lastRow = Math.max(12, paSheet.getLastRow());
  const values = paSheet.getRange(1, 1, lastRow, 6).getDisplayValues();

  let dateRow = 4;
  for (let r = 0; r < Math.min(10, values.length); r++) {
    let validDates = 0;
    for (let c = 1; c <= 5; c++) {
      if (parseHeaderDate(values[r][c])) validDates++;
    }
    if (validDates >= 3) {
      dateRow = r + 1;
      break;
    }
  }

  const startPeriodRow = dateRow + 1;
  let endPeriodRow = startPeriodRow;
  for (let r = startPeriodRow - 1; r < values.length; r++) {
    const hasAnyContent = values[r].some(
      (cell) => String(cell || "").trim() !== "",
    );
    if (hasAnyContent) endPeriodRow = r + 1;
  }

  return {
    dateRow,
    startPeriodRow,
    numPeriodRows: Math.max(1, endPeriodRow - startPeriodRow + 1),
  };
}

/**
 * Interpreta células do PA nos formatos:
 * - Simplificado ("Turma: 2º Lealdade\nCOMPONENTE CURRICULAR: Física")
 * - Ultra-curto ("2º LE\nFísica")
 * - Plano simplificado ("Eletiva: Título...")
 * - Já preenchido (com MODALIDADE ORGANIZATIVA: contendo conteúdo)
 */
function parsePaCellInfo(cellText, runtimeConfig) {
  if (!cellText) return null;
  const trimmed = String(cellText).trim();
  if (!trimmed) return null;

  const modMatch = trimmed.match(
    /MODALIDADE ORGANIZATIVA:\s*([\s\S]*?)\s*(?:AGRUPAMENTO:|$)/i,
  );
  const modContent = modMatch ? modMatch[1].trim() : "";
  const hasModalidade = modContent.length > 5;

  const turmaMatch = trimmed.match(/Turma:\s*([^\n\r]+)/i);
  const compMatch = trimmed.match(/COMPONENTE CURRICULAR:\s*([^\n\r]+)/i);
  if (turmaMatch && compMatch) {
    const classMeta = resolveClassFromText(turmaMatch[1], runtimeConfig);
    if (!classMeta) return null;
    const subjectName = resolveSubjectName(compMatch[1], runtimeConfig);
    return {
      className: classMeta.fullName,
      tabName: classMeta.tab,
      grade: classMeta.grade,
      noTv: classMeta.noTv,
      isSimplified: classMeta.isSimplified,
      customTitle: null,
      subject: subjectName,
      hasModalidade,
    };
  }

  const lines = trimmed
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (lines.length === 0) return null;

  const firstLine = lines[0];
  const colonParts = firstLine.split(":");
  const candidateClass = resolveClassFromText(colonParts[0], runtimeConfig);

  if (candidateClass) {
    if (candidateClass.isSimplified) {
      let customTitle = "";
      if (colonParts.length > 1) {
        customTitle = colonParts.slice(1).join(":").trim();
      } else if (lines.length > 1 && !/^TEMA/i.test(lines[1])) {
        customTitle = lines[1];
      }
      const subjectName = resolveSubjectName(
        candidateClass.fullName,
        runtimeConfig,
      );
      return {
        className: candidateClass.fullName,
        tabName: candidateClass.tab,
        grade: candidateClass.grade,
        noTv: candidateClass.noTv,
        isSimplified: true,
        customTitle: customTitle || candidateClass.fullName,
        subject: subjectName,
        hasModalidade,
      };
    }

    if (lines.length >= 2) {
      const subjectName = resolveSubjectName(lines[1], runtimeConfig);
      return {
        className: candidateClass.fullName,
        tabName: candidateClass.tab,
        grade: candidateClass.grade,
        noTv: candidateClass.noTv,
        isSimplified: candidateClass.isSimplified,
        customTitle: null,
        subject: subjectName,
        hasModalidade,
      };
    }
  }

  return null;
}

function resolveClassFromText(rawText, runtimeConfig) {
  const norm = normalizeKey(rawText);
  if (!norm) return null;
  if (runtimeConfig.classesByFullName[norm]) {
    return runtimeConfig.classesByFullName[norm];
  }
  if (runtimeConfig.classesByTab[norm]) {
    return runtimeConfig.classesByTab[norm];
  }

  for (let i = 0; i < runtimeConfig.classesList.length; i++) {
    const c = runtimeConfig.classesList[i];
    if (
      norm.includes(normalizeKey(c.fullName)) ||
      norm === normalizeKey(c.tab)
    ) {
      return c;
    }
  }
  return null;
}

function resolveSubjectName(rawSubject, runtimeConfig) {
  const clean = String(rawSubject || "").trim();
  const norm = normalizeKey(clean);
  if (runtimeConfig.subjectsByName[norm]) {
    return runtimeConfig.subjectsByName[norm].name;
  }
  for (let i = 0; i < runtimeConfig.subjectsList.length; i++) {
    const s = runtimeConfig.subjectsList[i];
    if (norm.includes(normalizeKey(s.name))) return s.name;
  }
  return clean || runtimeConfig.defaultSubject;
}

function findDraftInClassSheet(
  ss,
  tabName,
  targetSubject,
  targetDateObj,
  runtimeConfig,
  log,
) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    if (log)
      log.warn("Draft_Parser", `Aba "${tabName}" não encontrada na planilha`, {
        targetSubject,
      });
    return null;
  }

  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();
  const richTextData = sheet.getDataRange().getRichTextValues();

  let headerRowIdx = -1;
  let colMap = {
    conteudo: 0,
    aulas: -1,
    data: -1,
    disciplina: -1,
    bimestre: -1,
  };

  for (let r = 0; r < displayData.length; r++) {
    const firstCell = normalizeKey(displayData[r][0]);
    if (firstCell.startsWith("conteudos dados")) {
      headerRowIdx = r;
      for (let c = 0; c < displayData[r].length; c++) {
        const h = normalizeKey(displayData[r][c]);
        if (h.includes("conteudo")) colMap.conteudo = c;
        else if (h.includes("aula")) colMap.aulas = c;
        else if (h.includes("data")) colMap.data = c;
        else if (h.includes("disciplina")) colMap.disciplina = c;
        else if (h.includes("bim")) colMap.bimestre = c;
      }
      break;
    }
  }

  if (headerRowIdx === -1) {
    if (log)
      log.warn(
        "Draft_Parser",
        'Cabeçalho "Conteúdos dados:" não encontrado na aba',
        { tabName, targetSubject },
      );
    return null;
  }

  if (log) {
    log.info("Draft_Parser", `Tabela mapeada em ${tabName}`, {
      headerRow: headerRowIdx + 1,
      hasAulasCol: colMap.aulas !== -1,
      hasDataCol: colMap.data !== -1,
      hasDisciplinaCol: colMap.disciplina !== -1,
      hasBimestreCol: colMap.bimestre !== -1,
    });
  }

  const targetSubjectNorm = normalizeKey(targetSubject);
  const subjectRows = [];
  let ignoredOtherSubjectCount = 0;

  for (let r = headerRowIdx + 1; r < displayData.length; r++) {
    const col0 = String(displayData[r][0] || "").trim();
    if (/^nomes?$/i.test(col0)) break;

    const rawSummary = String(displayData[r][colMap.conteudo] || "").trim();
    const rawAulas =
      colMap.aulas !== -1
        ? String(displayData[r][colMap.aulas] || "").trim()
        : "";
    const rawDateDisplay =
      colMap.data !== -1
        ? String(displayData[r][colMap.data] || "").trim()
        : "";
    const rawDateVal = colMap.data !== -1 ? data[r][colMap.data] : "";
    const rawBimestre =
      colMap.bimestre !== -1
        ? String(displayData[r][colMap.bimestre] || "").trim()
        : "";

    let rowSubject = targetSubject;
    if (colMap.disciplina !== -1) {
      const dStr = String(displayData[r][colMap.disciplina] || "").trim();
      rowSubject =
        dStr === ""
          ? runtimeConfig.defaultSubject
          : resolveSubjectName(dStr, runtimeConfig);
    }

    if (!rawSummary && !rawAulas) continue;
    if (normalizeKey(rowSubject) !== targetSubjectNorm) {
      ignoredOtherSubjectCount++;
      continue;
    }

    const richLinks = extractRichTextUrls(richTextData[r][colMap.conteudo]);
    const parsedRowDate = parseFlexibleDate(rawDateVal, rawDateDisplay);

    subjectRows.push({
      rowNumber: r + 1,
      summary: rawSummary,
      aulas: rawAulas,
      bimestre: rawBimestre,
      richLinks: richLinks,
      dateObj: parsedRowDate,
      rawDateDisplay: rawDateDisplay,
      isDateBlank:
        !parsedRowDate &&
        (rawDateDisplay === "" || /por fazer/i.test(rawDateDisplay)),
      sheet: sheet,
      dateColIndex: colMap.data !== -1 ? colMap.data + 1 : null,
    });
  }

  if (subjectRows.length === 0) {
    if (log) {
      log.warn(
        "Draft_Parser",
        `Nenhuma linha de rascunho encontrada para ${tabName} (${targetSubject})`,
        { ignoredOtherSubjectCount },
      );
    }
    return null;
  }

  // Prioridade 1: Data exata
  for (let i = 0; i < subjectRows.length; i++) {
    const sr = subjectRows[i];
    if (sr.dateObj && isSameDayMonth(sr.dateObj, targetDateObj)) {
      if (log) {
        log.info(
          "Draft_Parser",
          `Rascunho selecionado por DATA EXATA em ${tabName} (${targetSubject})`,
          {
            row: sr.rowNumber,
            date: formatPtShortDate(targetDateObj),
            aulas: sr.aulas || "(vazio)",
            bimestreCol: sr.bimestre || "(não informado)",
          },
        );
      }
      return {
        summary: sr.summary,
        aulas: sr.aulas,
        bimestre: sr.bimestre,
        links: sr.richLinks,
        wasDateBlank: false,
        sheet: sr.sheet,
        rowNumber: sr.rowNumber,
        dateColIndex: sr.dateColIndex,
      };
    }
  }

  // Prioridade 2: Próxima linha sem data após a última aula dada
  let lastPastOrCurrentDatedIndex = -1;
  for (let i = 0; i < subjectRows.length; i++) {
    if (
      subjectRows[i].dateObj &&
      subjectRows[i].dateObj.getTime() <= targetDateObj.getTime()
    ) {
      lastPastOrCurrentDatedIndex = i;
    }
  }

  for (let i = lastPastOrCurrentDatedIndex + 1; i < subjectRows.length; i++) {
    const candidate = subjectRows[i];
    if (
      candidate.isDateBlank &&
      (candidate.summary !== "" || candidate.aulas !== "")
    ) {
      if (log) {
        log.info(
          "Draft_Parser",
          `Rascunho selecionado como PRÓXIMA LINHA SEM DATA em ${tabName} (${targetSubject})`,
          {
            row: candidate.rowNumber,
            afterRowIndex: lastPastOrCurrentDatedIndex,
            assignedDate: formatPtShortDate(targetDateObj),
            aulas: candidate.aulas || "(vazio)",
            bimestreCol: candidate.bimestre || "(não informado)",
          },
        );
      }
      return {
        summary: candidate.summary,
        aulas: candidate.aulas,
        bimestre: candidate.bimestre,
        links: candidate.richLinks,
        wasDateBlank: true,
        sheet: candidate.sheet,
        rowNumber: candidate.rowNumber,
        dateColIndex: candidate.dateColIndex,
      };
    }
  }

  if (log) {
    log.warn(
      "Draft_Parser",
      `Sem rascunho compatível para ${tabName} (${targetSubject}) na data ${formatPtShortDate(targetDateObj)}`,
      { candidateRows: subjectRows.length, lastPastOrCurrentDatedIndex },
    );
  }

  return null;
}

function loadEducEspData(ss, runtimeConfig, log) {
  const map = {};
  const sheet = ss.getSheetByName(CONFIG.TABS.EDUC_ESP);
  if (!sheet) {
    if (log)
      log.warn(
        "Educ_Esp_Load",
        `Aba "${CONFIG.TABS.EDUC_ESP}" não encontrada; acessibilidade será N/A.`,
      );
    return map;
  }

  const rows = sheet.getDataRange().getDisplayValues();
  const allStudents = [];

  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][0] || "").trim();
    const room = String(rows[i][1] || "").trim();
    const desc = String(rows[i][2] || "").trim();
    if (!name || !room) continue;
    if (!map[room]) map[room] = [];
    const studentObj = {
      name: name,
      cleanName: name.replace(/^\*\s*/, "").toLowerCase(),
      description: desc,
    };
    map[room].push(studentObj);
    allStudents.push(studentObj);
  }

  if (runtimeConfig && allStudents.length > 0) {
    runtimeConfig.classesList
      .filter((c) => c.isSimplified)
      .forEach((simClass) => {
        const simSheet = ss.getSheetByName(simClass.tab);
        if (!simSheet) return;
        const text = simSheet
          .getDataRange()
          .getDisplayValues()
          .map((r) => String(r[0] || "").toLowerCase())
          .join("\n");
        allStudents.forEach((st) => {
          if (st.cleanName && text.includes(st.cleanName)) {
            if (!map[simClass.tab]) map[simClass.tab] = [];
            if (!map[simClass.tab].some((x) => x.cleanName === st.cleanName)) {
              map[simClass.tab].push(st);
            }
          }
        });
      });
  }

  if (log) {
    log.info("Educ_Esp_Load", "Perfis de Educação Especial carregados", {
      totalStudents: allStudents.length,
      roomsMapped: Object.keys(map),
    });
  }

  return map;
}

function parseHeaderDate(headerStr) {
  if (!headerStr) return null;
  const m = String(headerStr).match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) return null;
  let year = parseInt(m[3], 10);
  if (year < 100) year += 2000;
  return new Date(year, parseInt(m[2], 10) - 1, parseInt(m[1], 10), 12, 0, 0);
}

function parseFlexibleDate(rawVal, displayStr) {
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
    return new Date(
      rawVal.getFullYear(),
      rawVal.getMonth(),
      rawVal.getDate(),
      12,
      0,
      0,
    );
  }
  if (!displayStr) return null;
  const s = displayStr.trim();

  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m)
    return new Date(
      parseInt(m[1], 10),
      parseInt(m[2], 10) - 1,
      parseInt(m[3], 10),
      12,
      0,
      0,
    );

  m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (m) {
    let y = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
    if (y < 100) y += 2000;
    return new Date(y, parseInt(m[2], 10) - 1, parseInt(m[1], 10), 12, 0, 0);
  }

  m = s.match(/^(\d{1,2})[\/\-]([a-zA-ZçÇ]{3,4})/i);
  if (m) {
    const day = parseInt(m[1], 10);
    const monStr = m[2].toLowerCase().substring(0, 3);
    const months = {
      jan: 0,
      fev: 1,
      feb: 1,
      mar: 2,
      abr: 3,
      apr: 3,
      mai: 4,
      may: 4,
      jun: 5,
      jul: 6,
      ago: 7,
      aug: 7,
      set: 8,
      sep: 8,
      out: 9,
      oct: 9,
      nov: 10,
      dez: 11,
      dec: 11,
    };
    if (months[monStr] !== undefined) {
      return new Date(new Date().getFullYear(), months[monStr], day, 12, 0, 0);
    }
  }
  return null;
}

function isSameDayMonth(d1, d2) {
  return (
    d1 && d2 && d1.getDate() === d2.getDate() && d1.getMonth() === d2.getMonth()
  );
}

function formatPtShortDate(dateObj) {
  const months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  const d =
    dateObj.getDate() < 10 ? `0${dateObj.getDate()}` : `${dateObj.getDate()}`;
  return `${d}/${months[dateObj.getMonth()]}`;
}

function extractRichTextUrls(richTextVal) {
  const urls = [];
  if (!richTextVal) return urls;
  const runs = richTextVal.getRuns();
  runs.forEach((run) => {
    const link = run.getLinkUrl();
    if (link && !urls.includes(link)) urls.push(link);
  });
  return urls;
}
