/**
 * ============================================================================
 * 4. DRIVESERVICE.GS - BUSCA DE PDFs, REGRA 2X1 ROBÓTICA E GUIA PRIORIZADO
 * ============================================================================
 */

function resolveLecturePdfs(block) {
  const blobs = [];
  const combinedText = `${block.draft.summary}\n${(block.draft.links || []).join('\n')}`;

  const urlRegex = /https?:\/\/[^\s)"]+/gi;
  const urls = combinedText.match(urlRegex) || [];

  urls.forEach(url => {
    try {
      if (/\.pdf(\?.*)?$/i.test(url) || url.includes('acervocmsp.educacao.sp.gov.br')) {
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) {
          blobs.push({ name: url.split('/').pop(), blob: resp.getBlob() });
        }
      } else if (url.includes('drive.google.com/file/d/')) {
        const fileId = extractDriveId(url);
        const file = DriveApp.getFileById(fileId);
        if (file.getMimeType() === MimeType.PDF) {
          blobs.push({ name: file.getName(), blob: file.getBlob() });
        }
      }
    } catch (e) {
      console.warn('Falha ao baixar PDF de URL direta: ' + url, e);
    }
  });

  if (blobs.length > 0) return blobs;

  const aulaNumbers = extractAulaNumbers(block.draft.aulas, block.draft.summary);
  if (aulaNumbers.length === 0) return blobs;

  const props = PropertiesService.getScriptProperties();
  const rootAulasFolderId = props.getProperty('AULAS_FOLDER_ID');
  if (!rootAulasFolderId) return blobs;

  try {
    const subjectKey = `${block.subject}_${block.grade}`;
    const targetFolderName = CONFIG.SUBJECT_FOLDER_NAMES[subjectKey];
    if (!targetFolderName) return blobs;

    const subjectFolder = findSubfolderByName(DriveApp.getFolderById(rootAulasFolderId), targetFolderName);
    if (!subjectFolder) return blobs;

    const allPdfs = collectPdfsInSubjectFolder(subjectFolder, CONFIG.CURRENT_BIMESTRE);
    const matchedFilesMap = {};

    aulaNumbers.forEach(aulaNum => {
      const match = findBestMatchingPdf(allPdfs, aulaNum, block.subject);
      if (match && !matchedFilesMap[match.getId()]) {
        matchedFilesMap[match.getId()] = match;
        blobs.push({ name: match.getName(), blob: match.getBlob() });
      }
    });
  } catch (e) {
    console.warn('Erro ao buscar PDFs no Drive:', e);
  }

  return blobs;
}

function extractAulaNumbers(aulasColText, summaryFallbackText) {
  const isFromAulasCol = Boolean(aulasColText && String(aulasColText).trim() !== '');
  const sourceText = isFromAulasCol ? String(aulasColText) : String(summaryFallbackText || '');
  if (!sourceText) return [];

  const nums = [];
  const tokenRegex = isFromAulasCol
    ? /\b(?:A(?:ula)?\s*0?)?(\d{1,2})(?:\s*(?:[-aà]|até)\s*(?:A(?:ula)?\s*0?)?(\d{1,2}))?\b/gi
    : /\bA(?:ula)?\s*0?(\d{1,2})(?:\s*(?:[-aà]|até)\s*A?(?:ula)?\s*0?(\d{1,2}))?\b/gi;

  let m;
  while ((m = tokenRegex.exec(sourceText)) !== null) {
    const start = parseInt(m[1], 10);
    const end = m[2] ? parseInt(m[2], 10) : start;

    if (start > 0 && start <= 45 && end >= start && end <= 45) {
      for (let i = start; i <= end; i++) {
        if (!nums.includes(i)) nums.push(i);
      }
    }
  }

  return nums.sort(function(a, b) { return a - b; });
}

function findBestMatchingPdf(pdfFiles, aulaNum, subject) {
  const isRobotics = /rob[oó]tica/i.test(subject);
  const pdfNumForRobotics = Math.ceil(aulaNum / 2);

  for (let i = 0; i < pdfFiles.length; i++) {
    const file = pdfFiles[i];
    const name = file.getName();

    if (isRobotics) {
      const pairFirst = (pdfNumForRobotics * 2) - 1;
      const pairSecond = pdfNumForRobotics * 2;
      const explicitPairRegex = new RegExp(`Aulas?\\s*0?${pairFirst}\\s*(?:e|a|-)\\s*0?${pairSecond}\\b`, 'i');
      const pdfIndexRegex = new RegExp(`(?:Aula|PDF|parte)\\s*0?${pdfNumForRobotics}\\b`, 'i');
      const directAulaRegex = new RegExp(`(?:Aula|PDF)\\s*0?${aulaNum}\\b`, 'i');

      if (explicitPairRegex.test(name) || pdfIndexRegex.test(name) || directAulaRegex.test(name)) {
        return file;
      }
    } else {
      const physicsRegex = new RegExp(`Aula\\s*0?${aulaNum}\\b`, 'i');
      if (physicsRegex.test(name)) {
        return file;
      }
    }
  }
  return null;
}

function standardizeDrivePdfNames() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const rootId = props.getProperty('AULAS_FOLDER_ID');

  if (!rootId) {
    ui.alert('⚠️ Configure a pasta "Aulas" primeiro no menu 5.');
    return;
  }

  const confirm = ui.alert(
    '🗂️ Padronizar Nomes dos PDFs no Drive',
    'Esta função irá organizar os nomes dos PDFs nas suas pastas de Física e Robótica para o padrão:\n\n' +
    '• Física: [Física 1 - 4º Bim] Aula 04 - Título.pdf\n' +
    '• Robótica: [Robótica 1 - 4º Bim] PDF 03 (Aulas 05 e 06) - Título.pdf\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const rootFolder = DriveApp.getFolderById(rootId);
  const subFolders = ['Física 1', 'Física 2', 'Física 3', 'Robótica 1', 'Robótica 2'];
  let renamedCount = 0;

  subFolders.forEach(folderName => {
    const folder = findSubfolderByName(rootFolder, folderName);
    if (!folder) return;
    const isRobotics = folderName.toLowerCase().includes('robótica');
    const files = collectPdfsInSubjectFolder(folder, null);

    files.forEach(file => {
      const oldName = file.getName();
      if (oldName.startsWith('[')) return;

      const bimMatch = oldName.match(/(\d)[ºo°]?\s*Bim/i);
      const bim = bimMatch ? `${bimMatch[1]}º Bim` : `${CONFIG.CURRENT_BIMESTRE}º Bim`;
      const aulaMatch = oldName.match(/Aula\s*0?(\d{1,2})/i);
      if (!aulaMatch) return;

      const num = parseInt(aulaMatch[1], 10);
      const padNum = num < 10 ? `0${num}` : `${num}`;

      let cleanTopic = oldName
        .replace(/\.pdf$/i, '')
        .replace(/F[ií]sica\s*\d\s*[-,]?\s*/gi, '')
        .replace(/Rob[oó]t\.?\s*[-,]?\s*/gi, '')
        .replace(/\d[ºo°]?\s*Bim(?:estre)?\s*[-,]?\s*/gi, '')
        .replace(/Aula\s*0?\d{1,2}\s*[-,]?\s*/gi, '')
        .trim();

      let newName = '';
      if (isRobotics) {
        const a1 = (num * 2) - 1;
        const a2 = num * 2;
        const p1 = a1 < 10 ? `0${a1}` : `${a1}`;
        const p2 = a2 < 10 ? `0${a2}` : `${a2}`;
        newName = `[${folderName} - ${bim}] PDF ${padNum} (Aulas ${p1} e ${p2}) - ${cleanTopic}.pdf`;
      } else {
        newName = `[${folderName} - ${bim}] Aula ${padNum} - ${cleanTopic}.pdf`;
      }

      file.setName(newName);
      renamedCount++;
    });
  });

  ui.alert(`✅ ${renamedCount} arquivo(s) PDF foram padronizados com sucesso!`);
}

function getPrioritizedCurriculumText(apiKey) {
  const props = PropertiesService.getScriptProperties();
  const guiaFolderId = props.getProperty('GUIA_FOLDER_ID');
  if (!guiaFolderId) return '';

  try {
    const folder = DriveApp.getFolderById(guiaFolderId);
    const files = folder.getFiles();
    let targetFile = null;

    while (files.hasNext()) {
      const f = files.next();
      if (/guia.*curr[ií]culo.*priorizado/i.test(f.getName())) {
        if (!targetFile || f.getLastUpdated() > targetFile.getLastUpdated()) {
          targetFile = f;
        }
      }
    }

    if (!targetFile) return '';

    const fileSignature = `${targetFile.getId()}_${targetFile.getLastUpdated().getTime()}`;
    const cachedSig = props.getProperty('GUIA_CACHE_SIG');
    const cachedText = props.getProperty('GUIA_CACHE_TEXT');

    if (cachedSig === fileSignature && cachedText) {
      return cachedText;
    }

    const blob = targetFile.getBlob();
    const prompt =
      'Extraia deste Guia do Currículo Priorizado de Física todas as habilidades (códigos EM13CNT... e Aprendizagens Essenciais AE...) ' +
      'organizadas por Série (1ª, 2ª e 3ª série), Bimestre e Tema. Seja conciso e liste todos os códigos e suas descrições oficiais em português.';

    const extractedCatalog = callGeminiTextOnlyWithPdf(apiKey, prompt, blob);
    if (extractedCatalog && extractedCatalog.length > 50) {
      const trimmed = extractedCatalog.substring(0, 8500);
      props.setProperty('GUIA_CACHE_SIG', fileSignature);
      props.setProperty('GUIA_CACHE_TEXT', trimmed);
      return trimmed;
    }
  } catch (e) {
    console.warn('Aviso ao ler Guia do Currículo Priorizado:', e);
  }
  return '';
}

function findSubfolderByName(parentFolder, targetName) {
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().trim().toLowerCase() === targetName.trim().toLowerCase()) return f;
  }
  return null;
}

function collectPdfsInSubjectFolder(subjectFolder, preferredBimestre) {
  const results = [];
  if (preferredBimestre) {
    const subs = subjectFolder.getFolders();
    while (subs.hasNext()) {
      const sub = subs.next();
      if (sub.getName().includes(preferredBimestre)) {
        const subFiles = sub.getFilesByType(MimeType.PDF);
        while (subFiles.hasNext()) results.push(subFiles.next());
      }
    }
  }
  const rootFiles = subjectFolder.getFilesByType(MimeType.PDF);
  while (rootFiles.hasNext()) results.push(rootFiles.next());
  return results;
}
