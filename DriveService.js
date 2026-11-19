/**
 * ============================================================================
 * 4. DRIVESERVICE.GS - CACHE BASE64 DE PDFs, BIMESTRE E REGRA 2X1 ROBÓTICA
 * ============================================================================
 */

function resolveLecturePdfsCached(block, pdfBase64Cache, log) {
  const results = [];
  const combinedText = `${block.draft.summary}\n${(block.draft.links || []).join("\n")}`;

  const urlRegex = /https?:\/\/[^\s)"]+/gi;
  const urls = combinedText.match(urlRegex) || [];

  urls.forEach((url) => {
    try {
      if (
        /\.pdf(\?.*)?$/i.test(url) ||
        url.includes("acervocmsp.educacao.sp.gov.br")
      ) {
        if (!pdfBase64Cache[url]) {
          const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
          if (resp.getResponseCode() === 200) {
            const bytes = resp.getBlob().getBytes();
            pdfBase64Cache[url] = {
              name: url.split("/").pop(),
              base64: Utilities.base64Encode(bytes),
              byteLength: bytes.length,
            };
          }
        }
        if (pdfBase64Cache[url]) results.push(pdfBase64Cache[url]);
      } else if (url.includes("drive.google.com/file/d/")) {
        const fileId = extractDriveId(url);
        if (!pdfBase64Cache[fileId]) {
          const file = DriveApp.getFileById(fileId);
          if (file.getMimeType() === MimeType.PDF) {
            const bytes = file.getBlob().getBytes();
            pdfBase64Cache[fileId] = {
              name: file.getName(),
              base64: Utilities.base64Encode(bytes),
              byteLength: bytes.length,
            };
          }
        }
        if (pdfBase64Cache[fileId]) results.push(pdfBase64Cache[fileId]);
      }
    } catch (e) {
      if (log)
        log.warn(`Falha ao baixar PDF de URL direta (${url})`, e.message);
    }
  });

  if (results.length > 0) return results;

  const targetBimestre = detectTargetBimestre(
    block.draft.bimestre,
    block.draft.aulas,
    block.draft.summary,
  );
  const parsedRequests = extractAulaRequests(
    block.draft.aulas,
    block.draft.summary,
  );
  if (parsedRequests.length === 0) return results;

  const props = PropertiesService.getScriptProperties();
  const rootAulasFolderId = props.getProperty("AULAS_FOLDER_ID");
  if (!rootAulasFolderId) return results;

  try {
    const subjectKey = `${block.subject}_${block.grade}`;
    const targetFolderName = CONFIG.SUBJECT_FOLDER_NAMES[subjectKey];
    if (!targetFolderName) return results;

    const subjectFolder = findSubfolderByName(
      DriveApp.getFolderById(rootAulasFolderId),
      targetFolderName,
    );
    if (!subjectFolder) return results;

    const candidatePdfs = collectPdfsForBimestre(subjectFolder, targetBimestre);
    const addedIds = {};

    parsedRequests.forEach((req) => {
      const match = findBestMatchingPdf(
        candidatePdfs,
        req,
        block.subject,
        targetBimestre,
      );
      if (match) {
        const fileId = match.getId();
        if (!addedIds[fileId]) {
          addedIds[fileId] = true;
          if (!pdfBase64Cache[fileId]) {
            const bytes = match.getBlob().getBytes();
            pdfBase64Cache[fileId] = {
              name: match.getName(),
              base64: Utilities.base64Encode(bytes),
              byteLength: bytes.length,
            };
            if (log)
              log.info(
                `PDF carregado e cacheado: ${match.getName()} (${(bytes.length / 1024).toFixed(0)} KB)`,
              );
          } else {
            if (log) log.info(`PDF reutilizado do cache: ${match.getName()}`);
          }
          results.push(pdfBase64Cache[fileId]);
        }
      } else {
        if (log)
          log.warn(
            `PDF não encontrado para ${block.tabName} (${block.subject}) - Bim ${targetBimestre}, Aula/PDF ${req.num}`,
          );
      }
    });
  } catch (e) {
    if (log) log.error("Erro ao buscar PDFs no Drive", e.message);
  }

  return results;
}

function detectTargetBimestre(bimColText, aulasText, summaryText) {
  if (bimColText) {
    const mCol = String(bimColText).match(/([1-4])/);
    if (mCol) return mCol[1];
  }
  const combined = `${aulasText || ""} ${summaryText || ""}`;
  const mText = combined.match(
    /\b(?:B(?:im(?:estre)?)?\s*([1-4])|([1-4])[ºo°]?\s*B(?:im(?:estre)?)?)\b/i,
  );
  if (mText) return mText[1] || mText[2];

  return String(CONFIG.CURRENT_BIMESTRE || "4");
}

function extractAulaRequests(aulasColText, summaryFallbackText) {
  const isFromAulasCol = Boolean(
    aulasColText && String(aulasColText).trim() !== "",
  );
  let sourceText = isFromAulasCol
    ? String(aulasColText)
    : String(summaryFallbackText || "");
  if (!sourceText) return [];

  sourceText = sourceText.replace(
    /\b(?:B(?:im(?:estre)?)?\s*[1-4]|[1-4][ºo°]?\s*B(?:im(?:estre)?)?)\b/gi,
    " ",
  );
  const map = {};

  const pdfRegex =
    /\bP(?:DF)?\s*0?(\d{1,2})(?:\s*(?:[-aà]|até)\s*P?(?:DF)?\s*0?(\d{1,2}))?\b/gi;
  let m;
  while ((m = pdfRegex.exec(sourceText)) !== null) {
    const start = parseInt(m[1], 10);
    const end = m[2] ? parseInt(m[2], 10) : start;
    if (start > 0 && start <= 45 && end >= start && end <= 45) {
      for (let i = start; i <= end; i++) {
        map[`P_${i}`] = { num: i, isDirectPdf: true };
      }
    }
  }

  const remainingText = sourceText.replace(pdfRegex, " ");
  const aulaRegex = isFromAulasCol
    ? /\b(?:A(?:ula)?\s*0?)?(\d{1,2})(?:\s*(?:[-aà]|até)\s*(?:A(?:ula)?\s*0?)?(\d{1,2}))?\b/gi
    : /\bA(?:ula)?\s*0?(\d{1,2})(?:\s*(?:[-aà]|até)\s*A?(?:ula)?\s*0?(\d{1,2}))?\b/gi;

  while ((m = aulaRegex.exec(remainingText)) !== null) {
    const start = parseInt(m[1], 10);
    const end = m[2] ? parseInt(m[2], 10) : start;
    if (start > 0 && start <= 45 && end >= start && end <= 45) {
      for (let i = start; i <= end; i++) {
        map[`A_${i}`] = { num: i, isDirectPdf: false };
      }
    }
  }

  return Object.keys(map)
    .map(function (k) {
      return map[k];
    })
    .sort(function (a, b) {
      return a.num - b.num;
    });
}

function findBestMatchingPdf(pdfFiles, req, subject, targetBimestre) {
  const isRobotics = /rob[oó]tica/i.test(subject);
  const targetNum =
    isRobotics && !req.isDirectPdf ? Math.ceil(req.num / 2) : req.num;

  const bimFiltered = pdfFiles.filter(function (file) {
    const name = file.getName();
    const bimInName = name.match(/([1-4])[ºo°]?\s*Bim/i);
    if (bimInName && bimInName[1] !== String(targetBimestre)) {
      return false;
    }
    return true;
  });

  const pool = bimFiltered.length > 0 ? bimFiltered : pdfFiles;

  for (let i = 0; i < pool.length; i++) {
    const file = pool[i];
    const name = file.getName();

    if (isRobotics) {
      const pairFirst = targetNum * 2 - 1;
      const pairSecond = targetNum * 2;
      const explicitPairRegex = new RegExp(
        `Aulas?\\s*0?${pairFirst}\\s*(?:e|a|-)\\s*0?${pairSecond}\\b`,
        "i",
      );
      const pdfIndexRegex = new RegExp(
        `(?:Aula|PDF|parte)\\s*0?${targetNum}\\b`,
        "i",
      );

      if (explicitPairRegex.test(name) || pdfIndexRegex.test(name)) {
        return file;
      }
    } else {
      const physicsRegex = new RegExp(`Aula\\s*0?${targetNum}\\b`, "i");
      if (physicsRegex.test(name)) {
        return file;
      }
    }
  }
  return null;
}

function collectPdfsForBimestre(subjectFolder, targetBimestre) {
  const results = [];
  if (targetBimestre) {
    const subs = subjectFolder.getFolders();
    while (subs.hasNext()) {
      const sub = subs.next();
      if (sub.getName().includes(String(targetBimestre))) {
        const subFiles = sub.getFilesByType(MimeType.PDF);
        while (subFiles.hasNext()) results.push(subFiles.next());
      }
    }
  }
  const rootFiles = subjectFolder.getFilesByType(MimeType.PDF);
  while (rootFiles.hasNext()) results.push(rootFiles.next());
  return results;
}

function standardizeDrivePdfNames() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const rootId = props.getProperty("AULAS_FOLDER_ID");

  if (!rootId) {
    ui.alert('⚠️ Configure a pasta "Aulas" primeiro no menu 6.');
    return;
  }

  const confirm = ui.alert(
    "🗂️ Padronizar Nomes dos PDFs no Drive",
    "Esta função irá organizar os nomes dos PDFs nas pastas de Física e Robótica para:\n\n" +
      "• Física: [Física 1 - 4º Bim] Aula 04 - Título.pdf\n" +
      "• Robótica: [Robótica 1 - 4º Bim] PDF 03 (Aulas 05 e 06) - Título.pdf\n\n" +
      "Deseja continuar?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) return;

  const rootFolder = DriveApp.getFolderById(rootId);
  const subFolders = [
    "Física 1",
    "Física 2",
    "Física 3",
    "Robótica 1",
    "Robótica 2",
  ];
  let renamedCount = 0;

  subFolders.forEach((folderName) => {
    const folder = findSubfolderByName(rootFolder, folderName);
    if (!folder) return;
    const isRobotics = folderName.toLowerCase().includes("robótica");
    const files = collectPdfsForBimestre(folder, null);

    files.forEach((file) => {
      const oldName = file.getName();
      if (oldName.startsWith("[")) return;

      const bimMatch = oldName.match(/(\d)[ºo°]?\s*Bim/i);
      const bim = bimMatch
        ? `${bimMatch[1]}º Bim`
        : `${CONFIG.CURRENT_BIMESTRE}º Bim`;
      const aulaMatch = oldName.match(/Aula\s*0?(\d{1,2})/i);
      if (!aulaMatch) return;

      const num = parseInt(aulaMatch[1], 10);
      const padNum = num < 10 ? `0${num}` : `${num}`;

      let cleanTopic = oldName
        .replace(/\.pdf$/i, "")
        .replace(/F[ií]sica\s*\d\s*[-,]?\s*/gi, "")
        .replace(/Rob[oó]t\.?\s*[-,]?\s*/gi, "")
        .replace(/\d[ºo°]?\s*Bim(?:estre)?\s*[-,]?\s*/gi, "")
        .replace(/Aula\s*0?\d{1,2}\s*[-,]?\s*/gi, "")
        .trim();

      let newName = "";
      if (isRobotics) {
        const a1 = num * 2 - 1;
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

function getPrioritizedCurriculumText(apiKey, log) {
  const props = PropertiesService.getScriptProperties();
  const guiaFolderId = props.getProperty("GUIA_FOLDER_ID");
  if (!guiaFolderId) return "";

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

    if (!targetFile) return "";

    const fileSignature = `${targetFile.getId()}_${targetFile.getLastUpdated().getTime()}`;
    const cachedSig = props.getProperty("GUIA_CACHE_SIG");
    const cachedText = props.getProperty("GUIA_CACHE_TEXT");

    if (cachedSig === fileSignature && cachedText) {
      if (log) log.info("Guia do Currículo Priorizado carregado do cache.");
      return cachedText;
    }

    if (log)
      log.info(
        "Extraindo habilidades do Guia do Currículo Priorizado via Gemini...",
      );
    const blob = targetFile.getBlob();
    const prompt =
      "Extraia deste Guia do Currículo Priorizado de Física todas as habilidades (códigos EM13CNT... e Aprendizagens Essenciais AE...) " +
      "organizadas por Série (1ª, 2ª e 3ª série), Bimestre e Tema. Seja conciso e liste todos os códigos e suas descrições oficiais em português.";

    const extractedCatalog = callGeminiTextOnlyWithPdf(apiKey, prompt, blob);
    if (extractedCatalog && extractedCatalog.length > 50) {
      const trimmed = extractedCatalog.substring(0, 8500);
      props.setProperty("GUIA_CACHE_SIG", fileSignature);
      props.setProperty("GUIA_CACHE_TEXT", trimmed);
      return trimmed;
    }
  } catch (e) {
    if (log) log.warn("Aviso ao ler Guia do Currículo Priorizado", e.message);
  }
  return "";
}

function findSubfolderByName(parentFolder, targetName) {
  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().trim().toLowerCase() === targetName.trim().toLowerCase())
      return f;
  }
  return null;
}
