/**
 * ============================================================================
 * 5. GEMINISERVICE.JS - PAYLOAD DINÂMICO (MATERIAIS/CRITÉRIOS) E RICHTEXT
 * ============================================================================
 */

function buildGeminiRequestPayload(params) {
  const {
    apiKey,
    modelName,
    block,
    subjectMeta,
    studentsEsp,
    encodedPdfs,
    guiaSummary,
    log,
  } = params;

  const draftLower = (
    block.draft.summary +
    " " +
    (block.draft.links || []).join(" ")
  ).toLowerCase();
  const mentionsVideoOrSim =
    draftLower.includes("phet") ||
    draftLower.includes("youtube.com") ||
    draftLower.includes("youtu.be") ||
    draftLower.includes("vídeo") ||
    draftLower.includes("video") ||
    draftLower.includes("simula");

  let mediaConstraintInstruction = "";
  let mediaDecision = "";
  if (block.noTv) {
    if (mentionsVideoOrSim) {
      mediaDecision = "NO_TV_WITH_PROJECTOR";
      mediaConstraintInstruction =
        `RESTRIÇÃO DE SALA (${block.tabName}): Esta sala NÃO possui Televisor. ` +
        `PROIBIDO incluir "Televisor" em RECURSOS (mesmo que apareça no catálogo usual). Como o rascunho menciona vídeo ou simulação, utilize "Projetor" em RECURSOS.`;
    } else {
      mediaDecision = "NO_TV_PRINTED_ONLY";
      mediaConstraintInstruction =
        `RESTRIÇÃO DE SALA (${block.tabName}): Esta sala NÃO possui Televisor. ` +
        `PROIBIDO incluir "Televisor" ou "Projetor" em RECURSOS (mesmo que apareça no catálogo usual). Priorize "Material Impresso / Lista de Exercícios" e "Quadro Branco".`;
    }
  } else {
    mediaDecision = "TV_ALLOWED";
    mediaConstraintInstruction = `SALA COM TELEVISOR (${block.tabName}): Você pode incluir "Televisor" em RECURSOS caso a aula utilize slides, vídeos ou simulações.`;
  }

  if (log) {
    log.info(
      "Media_Policy",
      `Política de mídia aplicada para ${block.tabName} (${block.subject})`,
      { noTv: block.noTv, mentionsVideoOrSim, mediaDecision },
    );
  }

  let accessibilityInstruction = "";
  if (studentsEsp.length === 0) {
    accessibilityInstruction = `Esta turma NÃO possui estudantes na aba de Educação Especial. O campo "acessibilidade_curricular" DEVE ser estritamente "N/A".`;
  } else {
    const profiles = studentsEsp
      .map((s) => `- ${s.name}: ${s.description}`)
      .join("\n");
    accessibilityInstruction =
      `Esta turma possui o(s) seguinte(s) estudante(s) público-alvo da Educação Especial:\n${profiles}\n` +
      `Sugira em "acessibilidade_curricular" uma adaptação prática, curta e específica para a aula de hoje considerando exatamente as capacidades descritas acima.`;
  }

  const materialsPool =
    (subjectMeta && subjectMeta.materials) ||
    "Quadro Branco, Caderno do Estudante, Material Impresso";
  const evalCriteriaBase =
    (subjectMeta && subjectMeta.evalCriteria) ||
    "Participação na aula, resposta às perguntas do professor e resolução das atividades propostas.";

  const splitRule = block.isDouble
    ? `A aula é uma DOBRADINHA (duas aulas seguidas de 50 minutos: "aula_1" e "aula_2"). Divida em planejamentos sequenciais e distintos (Parte 1: primeiros 50 min; Parte 2: segundos 50 min).`
    : `A aula é SIMPLES (1 aula de 50 minutos). Preencha "aula_1" completamente e deixe "aula_2" igual a null.`;

  let systemPrompt = "";

  if (block.isSimplified) {
    const titleLabel = block.customTitle || block.className;
    systemPrompt = `Você é um professor especialista do Programa Ensino Integral (PEI) da SEDUC-SP lecionando "${titleLabel}".
Sua tarefa é elaborar o Plano de Aula simplificado (sem Habilidades BNCC nem Critérios de Avaliação formais), redigido integralmente em Português do Brasil (pt-BR).

DIRETRIZES DE PREENCHIMENTO:
1. ESTRUTURA DE TEMPO:
   - ${splitRule}
2. RECURSOS E RESTRIÇÃO DE SALA:
   - ${mediaConstraintInstruction}
   - Baseie-se nos materiais citados no rascunho e no catálogo usual: [${materialsPool}].
3. MODALIDADE ORGANIZATIVA:
   - Descreva o passo a passo metodológico de cada aula de 50 minutos em frases separadas por ponto-e-vírgula (;).
4. ACESSIBILIDADE CURRICULAR:
   - ${accessibilityInstruction}
5. INSTRUÇÕES DE ESTILO/TAMANHO NO RASCUNHO:
   - Caso o professor tenha escrito alguma instrução de estilo ou tamanho no rascunho, obedeça-a com prioridade máxima (sem copiar o comando meta para o texto final).

Retorne APENAS um objeto JSON válido em Português (pt-BR):
{
  "aula_1": { "tema_da_aula": "...", "recursos": "...", "modalidade_organizativa": "...", "agrupamento": "...", "acessibilidade_curricular": "..." },
  "aula_2": { "tema_da_aula": "...", "recursos": "...", "modalidade_organizativa": "...", "agrupamento": "...", "acessibilidade_curricular": "..." }
}`;
  } else {
    systemPrompt = `Você é um professor especialista do Programa Ensino Integral (PEI) da SEDUC-SP lecionando ${block.subject} para a turma ${block.className}.
Sua tarefa é elaborar o Plano de Aula estruturado seguindo rigorosamente as diretrizes da SEDUC-SP, redigido integralmente em Português do Brasil (pt-BR).

DIRETRIZES DE PREENCHIMENTO:
1. ESTRUTURA DE TEMPO:
   - ${splitRule}
   - Use as anotações do rascunho do professor e os PDFs anexados (se houver) para organizar o conteúdo.
2. RECURSOS E RESTRIÇÃO DE HARDWARE DA SALA:
   - ${mediaConstraintInstruction}
   - Catálogo usual do professor para ${block.subject}: [${materialsPool}].
   - Selecione apenas os recursos pertinentes à aula descrita e acrescente materiais novos se o rascunho citar experimentos ou links específicos.
3. CRITÉRIOS DE AVALIAÇÃO:
   - Use como base os critérios usuais do professor para ${block.subject}:
     ${evalCriteriaBase}
   - Adapte brevemente para a atividade específica de hoje, mantendo o estilo objetivo.
4. HABILIDADES (BNCC / CURRÍCULO PAULISTA):
   - Se houver PDF da SEDUC anexado, extraia o código exato da habilidade e sua descrição em português.
   - Se NÃO houver PDF anexado, consulte o Guia do Currículo Priorizado abaixo e escolha o código mais próximo do tema:
   ${guiaSummary ? guiaSummary : "(Use os códigos oficiais da BNCC / Currículo Paulista compatíveis com a disciplina e o tema)."}
5. MODALIDADE ORGANIZATIVA:
   - Escreva em tópicos ou frases separadas por ponto-e-vírgula (;), detalhando o passo a passo metodológico do professor naqueles 50 minutos.
6. ACESSIBILIDADE CURRICULAR:
   - ${accessibilityInstruction}
7. INSTRUÇÕES DE ESTILO/TAMANHO NO RASCUNHO:
   - Caso o professor tenha escrito alguma instrução de formatação, tamanho ou linguagem dentro do próprio rascunho, obedeça-a com prioridade máxima (sem incluir esse comando meta no texto final).

Retorne APENAS um objeto JSON válido em Português (pt-BR):
{
  "aula_1": { "tema_da_aula": "...", "recursos": "...", "criterios_avaliacao": "...", "habilidades": "...", "modalidade_organizativa": "...", "agrupamento": "...", "acessibilidade_curricular": "..." },
  "aula_2": { "tema_da_aula": "...", "recursos": "...", "criterios_avaliacao": "...", "habilidades": "...", "modalidade_organizativa": "...", "agrupamento": "...", "acessibilidade_curricular": "..." }
}`;
  }

  const userParts = [
    {
      text:
        `DADOS DO RASCUNHO DO PROFESSOR:\n` +
        `- Turma: ${block.className} (${block.tabName})\n` +
        `- Disciplina: ${block.subject}\n` +
        `- Data: ${block.dateStr}\n` +
        `- Bimestre: ${block.draft.bimestre || CONFIG.CURRENT_BIMESTRE}\n` +
        `- Indicação na coluna Aulas: "${block.draft.aulas || "Não informado"}"\n` +
        `- Resumo/Anotações do Professor:\n"""\n${block.draft.summary}\n"""\n` +
        (block.draft.links && block.draft.links.length > 0
          ? `- Links extraídos da célula: ${block.draft.links.join(", ")}\n`
          : ""),
    },
  ];

  let totalBytes = 0;
  let attachedCount = 0;
  const maxPayloadBytes = 22 * 1024 * 1024;

  encodedPdfs.forEach((pdfObj, idx) => {
    if (totalBytes + pdfObj.byteLength < maxPayloadBytes) {
      totalBytes += pdfObj.byteLength;
      attachedCount++;
      userParts.push({
        text: `Arquivo PDF anexado #${idx + 1}: ${pdfObj.name}`,
      });
      userParts.push({
        inlineData: {
          mimeType: "application/pdf",
          data: pdfObj.base64,
        },
      });
    } else if (log) {
      log.warn(
        "Payload_Cap",
        `PDF ignorado por exceder limite acumulado de 22MB`,
        {
          tab: block.tabName,
          subject: block.subject,
          pdfName: pdfObj.name,
          pdfKB: Math.round(pdfObj.byteLength / 1024),
          currentTotalMB: (totalBytes / (1024 * 1024)).toFixed(2),
        },
      );
    }
  });

  if (log) {
    log.info(
      "Payload_Built",
      `Payload montado para ${block.tabName} (${block.subject})`,
      {
        date: block.shortDateStr,
        isSimplified: block.isSimplified,
        attachedPdfs: attachedCount,
        requestedPdfs: encodedPdfs.length,
        pdfBytesKB: Math.round(totalBytes / 1024),
      },
    );
  }

  const activeModel = modelName || CONFIG.DEFAULT_MODEL;
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${activeModel}:generateContent?key=${apiKey}`;
  const payload = {
    systemInstruction: { parts: [{ text: systemPrompt }] },
    contents: [{ role: "user", parts: userParts }],
    generationConfig: {
      temperature: 0.3,
      responseMimeType: "application/json",
    },
  };

  return {
    url: url,
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
}

function parseGeminiJsonResponse(responseText, block, log) {
  try {
    const json = JSON.parse(responseText);
    const candidateText =
      json.candidates &&
      json.candidates[0] &&
      json.candidates[0].content &&
      json.candidates[0].content.parts &&
      json.candidates[0].content.parts[0]
        ? json.candidates[0].content.parts[0].text
        : null;

    if (!candidateText) {
      const finishReason =
        json.candidates && json.candidates[0]
          ? json.candidates[0].finishReason
          : "UNKNOWN";
      if (log) {
        log.error(
          "Gemini_Response",
          `Sem conteúdo de texto na resposta de ${block.tabName} (${block.subject})`,
          {
            finishReason,
            promptFeedback: json.promptFeedback,
            rawPreview: responseText.slice(0, 350),
          },
        );
      }
      return null;
    }

    const parsed = JSON.parse(candidateText);
    if (log) {
      log.info(
        "Gemini_Response",
        `JSON válido recebido para ${block.tabName} (${block.subject})`,
        {
          hasAula1: Boolean(parsed.aula_1),
          hasAula2: Boolean(parsed.aula_2),
          temaAula1: parsed.aula_1 ? parsed.aula_1.tema_da_aula : undefined,
        },
      );
    }
    return parsed;
  } catch (e) {
    if (log) {
      log.error(
        "Gemini_Parse",
        `Falha ao interpretar JSON retornado pelo Gemini em ${block.tabName}`,
        e,
        { rawPreview: responseText.slice(0, 350) },
      );
    }
    return null;
  }
}

function callGeminiTextOnlyWithPdf(apiKey, promptText, pdfBlob) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.DEFAULT_MODEL}:generateContent?key=${apiKey}`;
  const payload = {
    contents: [
      {
        role: "user",
        parts: [
          { text: promptText },
          {
            inlineData: {
              mimeType: "application/pdf",
              data: Utilities.base64Encode(pdfBlob.getBytes()),
            },
          },
        ],
      },
    ],
    generationConfig: { temperature: 0.1 },
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  return data.candidates && data.candidates[0].content.parts[0].text
    ? data.candidates[0].content.parts[0].text
    : "";
}

function buildFormattedCellRichText(block, planPart, studentsEsp, log) {
  let recursos = planPart.recursos || "";
  if (block.noTv) {
    const hadTv = /televisor/i.test(recursos);
    recursos = recursos.replace(/televisor(es)?\s*(,|e)?\s*/gi, "").trim();
    if (hadTv && log) {
      log.info(
        "NoTv_Guard",
        `Removido "Televisor" de RECURSOS por restrição da sala ${block.tabName}`,
      );
    }
  }

  const acessibilidade =
    studentsEsp.length === 0
      ? "N/A"
      : planPart.acessibilidade_curricular ||
        "Participação prática acompanhada.";

  const sections = block.isSimplified
    ? [
        {
          label: `${block.className}: ${block.customTitle || ""}`.replace(
            /:\s*$/,
            "",
          ),
          value: "",
        },
        { label: "TEMA DA AULA: ", value: planPart.tema_da_aula || "" },
        { label: "RECURSOS: ", value: recursos },
        {
          label: "MODALIDADE ORGANIZATIVA: ",
          value: planPart.modalidade_organizativa || "",
        },
        {
          label: "AGRUPAMENTO: ",
          value: planPart.agrupamento || "Grupos produtivos / Bancadas.",
        },
        { label: "ACESSIBILIDADE CURRICULAR: ", value: acessibilidade },
      ]
    : [
        { label: "Turma: ", value: block.className },
        { label: "COMPONENTE CURRICULAR: ", value: block.subject },
        { label: "TEMA DA AULA: ", value: planPart.tema_da_aula || "" },
        { label: "RECURSOS: ", value: recursos },
        {
          label: "CRITÉRIOS DE AVALIAÇÃO: ",
          value: planPart.criterios_avaliacao || "",
        },
        { label: "HABILIDADES: ", value: planPart.habilidades || "" },
        {
          label: "MODALIDADE ORGANIZATIVA: ",
          value: planPart.modalidade_organizativa || "",
        },
        { label: "AGRUPAMENTO: ", value: planPart.agrupamento || "" },
        { label: "ACESSIBILIDADE CURRICULAR: ", value: acessibilidade },
      ];

  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
  const normalStyle = SpreadsheetApp.newTextStyle().setBold(false).build();

  let fullText = "";
  const ranges = [];

  sections.forEach((sec, idx) => {
    const labelStart = fullText.length;
    fullText += sec.label;
    const labelEnd = fullText.length;
    ranges.push({ start: labelStart, end: labelEnd, style: boldStyle });

    if (sec.value) {
      const valStart = fullText.length;
      fullText += sec.value;
      const valEnd = fullText.length;
      ranges.push({ start: valStart, end: valEnd, style: normalStyle });
    }

    if (idx < sections.length - 1) {
      const sepStart = fullText.length;
      fullText += "\n\n";
      const sepEnd = fullText.length;
      ranges.push({ start: sepStart, end: sepEnd, style: normalStyle });
    }
  });

  const builder = SpreadsheetApp.newRichTextValue().setText(fullText);
  ranges.forEach((r) => {
    if (r.end > r.start) {
      builder.setTextStyle(r.start, r.end, r.style);
    }
  });

  return builder.build();
}
