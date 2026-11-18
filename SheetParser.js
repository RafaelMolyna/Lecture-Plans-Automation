/**
 * ============================================================================
 * 3. SHEETPARSER.GS - LEITURA DE ABAS, DATAS E EDUCAÇÃO ESPECIAL
 * ============================================================================
 */

function findDraftInClassSheet(ss, tabName, targetSubject, targetDateObj, isEletiva) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();
  const richTextData = sheet.getDataRange().getRichTextValues();

  let headerRowIdx = -1;
  let colMap = { conteudo: 0, aulas: -1, data: -1, disciplina: -1 };

  for (let r = 0; r < displayData.length; r++) {
    const firstCell = String(displayData[r][0] || '').trim().toLowerCase();
    if (firstCell.startsWith('conteúdos dados') || firstCell.startsWith('conteudos dados')) {
      headerRowIdx = r;
      for (let c = 0; c < displayData[r].length; c++) {
        const h = String(displayData[r][c] || '').trim().toLowerCase();
        if (h.includes('conteúdo') || h.includes('conteudo')) colMap.conteudo = c;
        else if (h.includes('aula')) colMap.aulas = c;
        else if (h.includes('data')) colMap.data = c;
        else if (h.includes('disciplina')) colMap.disciplina = c;
      }
      break;
    }
  }

  if (headerRowIdx === -1) return null;

  const subjectRows = [];
  for (let r = headerRowIdx + 1; r < displayData.length; r++) {
    const col0 = String(displayData[r][0] || '').trim();
    if (/^nomes?$/i.test(col0)) break;

    const rawSummary = String(displayData[r][colMap.conteudo] || '').trim();
    const rawAulas = colMap.aulas !== -1 ? String(displayData[r][colMap.aulas] || '').trim() : '';
    const rawDateDisplay = colMap.data !== -1 ? String(displayData[r][colMap.data] || '').trim() : '';
    const rawDateVal = colMap.data !== -1 ? data[r][colMap.data] : '';

    let rowSubject = isEletiva ? 'Eletiva' : 'Física';
    if (!isEletiva && colMap.disciplina !== -1) {
      const dStr = String(displayData[r][colMap.disciplina] || '').trim();
      if (/rob[oó]tica/i.test(dStr)) rowSubject = 'Robótica';
      else if (/f[ií]sica/i.test(dStr)) rowSubject = 'Física';
      else if (/eletiva/i.test(dStr)) rowSubject = 'Eletiva';
      else if (dStr === '') rowSubject = 'Física';
    }

    if (!rawSummary && !rawAulas) continue;
    if (rowSubject.toLowerCase() !== targetSubject.toLowerCase()) continue;

    const richLinks = extractRichTextUrls(richTextData[r][colMap.conteudo]);
    const parsedRowDate = parseFlexibleDate(rawDateVal, rawDateDisplay);

    subjectRows.push({
      rowNumber: r + 1,
      summary: rawSummary,
      aulas: rawAulas,
      richLinks: richLinks,
      dateObj: parsedRowDate,
      rawDateDisplay: rawDateDisplay,
      isDateBlank: !parsedRowDate && (rawDateDisplay === '' || /por fazer/i.test(rawDateDisplay)),
      sheet: sheet,
      dateColIndex: colMap.data !== -1 ? colMap.data + 1 : null
    });
  }

  if (subjectRows.length === 0) return null;

  // Prioridade 1: Data exata (Caso você já tenha colocado a data nas próximas semanas)
  for (let i = 0; i < subjectRows.length; i++) {
    const sr = subjectRows[i];
    if (sr.dateObj && isSameDayMonth(sr.dateObj, targetDateObj)) {
      return {
        summary: sr.summary,
        aulas: sr.aulas,
        links: sr.richLinks,
        wasDateBlank: false,
        sheet: sr.sheet,
        rowNumber: sr.rowNumber,
        dateColIndex: sr.dateColIndex
      };
    }
  }

  // Prioridade 2: Próxima linha sem data após a última aula dada (com data <= data alvo da semana)
  let lastPastOrCurrentDatedIndex = -1;
  for (let i = 0; i < subjectRows.length; i++) {
    if (subjectRows[i].dateObj && subjectRows[i].dateObj.getTime() <= targetDateObj.getTime()) {
      lastPastOrCurrentDatedIndex = i;
    }
  }

  for (let i = lastPastOrCurrentDatedIndex + 1; i < subjectRows.length; i++) {
    const candidate = subjectRows[i];
    if (candidate.isDateBlank && (candidate.summary !== '' || candidate.aulas !== '')) {
      return {
        summary: candidate.summary,
        aulas: candidate.aulas,
        links: candidate.richLinks,
        wasDateBlank: true,
        sheet: candidate.sheet,
        rowNumber: candidate.rowNumber,
        dateColIndex: candidate.dateColIndex
      };
    }
  }

  return null;
}

function parsePaCellInfo(cellText) {
  if (!cellText) return null;
  const trimmed = cellText.trim();

  if (/^Eletiva\b/i.test(trimmed)) {
    const lines = trimmed.split('\n').map(l => l.trim()).filter(Boolean);
    let eletivaTitle = 'Experimentos de Física Mão na Massa';
    if (lines[0].includes(':')) {
      eletivaTitle = lines[0].split(':')[1].trim() || eletivaTitle;
    } else if (lines.length > 1 && !lines[1].startsWith('TEMA')) {
      eletivaTitle = lines[1];
    }

    const modMatch = trimmed.match(/MODALIDADE ORGANIZATIVA:\s*([\s\S]*?)\s*(?:AGRUPAMENTO:|$)/i);
    const modContent = modMatch ? modMatch[1].trim() : '';

    return {
      className: 'Eletiva',
      tabName: 'Eletiva',
      grade: 0,
      noTv: false,
      isEletiva: true,
      eletivaTitle: eletivaTitle,
      subject: 'Eletiva',
      hasModalidade: modContent.length > 5
    };
  }

  if (!trimmed.includes('Turma:')) return null;

  const turmaMatch = trimmed.match(/Turma:\s*([^\n\r]+)/i);
  const compMatch = trimmed.match(/COMPONENTE CURRICULAR:\s*([^\n\r]+)/i);
  if (!turmaMatch || !compMatch) return null;

  const rawClassName = turmaMatch[1].trim();
  const rawSubject = compMatch[1].trim();

  let matchedKey = null;
  Object.keys(CONFIG.CLASS_MAP).forEach(k => {
    if (rawClassName.toLowerCase().includes(k.toLowerCase())) matchedKey = k;
  });
  if (!matchedKey) return null;

  const classMeta = CONFIG.CLASS_MAP[matchedKey];
  const modMatch = trimmed.match(/MODALIDADE ORGANIZATIVA:\s*([\s\S]*?)\s*AGRUPAMENTO:/i);
  const modContent = modMatch ? modMatch[1].trim() : '';

  return {
    className: matchedKey,
    tabName: classMeta.tab,
    grade: classMeta.grade,
    noTv: classMeta.noTv,
    isEletiva: false,
    subject: /rob[oó]tica/i.test(rawSubject) ? 'Robótica' : 'Física',
    hasModalidade: modContent.length > 5
  };
}

function loadEducEspData(ss) {
  const map = {};
  const sheet = ss.getSheetByName('Educ. Esp.');
  if (!sheet) return map;

  const rows = sheet.getDataRange().getDisplayValues();
  const allStudents = [];

  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][0] || '').trim();
    const room = String(rows[i][1] || '').trim();
    const desc = String(rows[i][2] || '').trim();
    if (!name || !room) continue;
    if (!map[room]) map[room] = [];
    const studentObj = { name: name, cleanName: name.replace(/^\*\s*/, '').toLowerCase(), description: desc };
    map[room].push(studentObj);
    allStudents.push(studentObj);
  }

  const eletivaSheet = ss.getSheetByName('Eletiva');
  if (eletivaSheet && allStudents.length > 0) {
    const eletivaText = eletivaSheet.getDataRange().getDisplayValues().map(r => String(r[0] || '').toLowerCase()).join('\n');
    allStudents.forEach(st => {
      if (st.cleanName && eletivaText.includes(st.cleanName)) {
        if (!map['Eletiva']) map['Eletiva'] = [];
        map['Eletiva'].push(st);
      }
    });
  }

  return map;
}

function parseHeaderDate(headerStr) {
  if (!headerStr) return null;
  const m = headerStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) return null;
  let year = parseInt(m[3], 10);
  if (year < 100) year += 2000;
  return new Date(year, parseInt(m[2], 10) - 1, parseInt(m[1], 10), 12, 0, 0);
}

function parseFlexibleDate(rawVal, displayStr) {
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
    return new Date(rawVal.getFullYear(), rawVal.getMonth(), rawVal.getDate(), 12, 0, 0);
  }
  if (!displayStr) return null;
  const s = displayStr.trim();

  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10), 12, 0, 0);

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
      jan: 0, fev: 1, feb: 1, mar: 2, abr: 3, apr: 3, mai: 4, may: 4,
      jun: 5, jul: 6, ago: 7, aug: 7, set: 8, sep: 8, out: 9, oct: 9,
      nov: 10, dez: 11, dec: 11
    };
    if (months[monStr] !== undefined) {
      return new Date(new Date().getFullYear(), months[monStr], day, 12, 0, 0);
    }
  }
  return null;
}

function isSameDayMonth(d1, d2) {
  return d1 && d2 && d1.getDate() === d2.getDate() && d1.getMonth() === d2.getMonth();
}

function formatPtShortDate(dateObj) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const d = dateObj.getDate() < 10 ? `0${dateObj.getDate()}` : `${dateObj.getDate()}`;
  return `${d}/${months[dateObj.getMonth()]}`;
}

function extractRichTextUrls(richTextVal) {
  const urls = [];
  if (!richTextVal) return urls;
  const runs = richTextVal.getRuns();
  runs.forEach(run => {
    const link = run.getLinkUrl();
    if (link && !urls.includes(link)) urls.push(link);
  });
  return urls;
}
