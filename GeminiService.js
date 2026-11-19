/**
 * ============================================================================
 * 5. GEMINISERVICE.GS - CONSTRUTOR DE PAYLOAD PARALELO E PARSER DE RESPOSTAS
 * ============================================================================
 */

function buildGeminiRequestPayload(params) {
  const { apiKey, modelName, block, studentsEsp, encodedPdfs, guiaSummary } =
    params;

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
  if (block.noTv) {
    if (mentionsVideoOrSim) {
      mediaConstraintInstruction =
        `RESTRIÇÃO DE SALA (${block.tabName}): Esta sala NÃO possui Televisor. ` +
        `PROIBIDO incluir "Televisor" em RECURSOS. Como o rascunho menciona vídeo ou simulação, utilize "Projetor" em RECURSOS.`;
    } else {
      mediaConstraintInstruction =
        `RESTRIÇÃO DE SALA (${block.tabName}): Esta sala NÃO possui Televisor. ` +
        `PROIBIDO incluir "Televisor" ou "Projetor" em RECURSOS. Priorize "Material Impresso / Lista de Exercícios" e "Quadro Branco".`;
    }
  } else {
    mediaConstraintInstruction = `SALA COM TELEVISOR (${block.tabName}): Você pode incluir "Televisor" em RECURSOS caso a aula utilize slides, vídeos ou simulações.`;
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

  const materialsPool = (CONFIG.MASTER_MATERIALS[block.subject] || []).join(
    ", ",
  );
  let systemPrompt = "";

  if (block.isEletiva) {
    systemPrompt = `Você é um professor especialista do Programa Ensino Integral (PEI) da SEDUC-SP (EE Matilde Maria Cremm) lecionando a disciplina ELETIVA "${block.eletivaTitle}".
Sua tarefa é elaborar o Plano de Aula simplificado para a Eletiva (que NÃO exige Habilidades BNCC nem Critérios de Avaliação formais), redigido integralmente em Português do Brasil (pt-BR).

DIRETRIZES DE PREENCHIMENTO DA ELETIVA:
1. DOBRADINHA (SPLIT DE 50 MIN + 50 MIN):
   - A Eletiva ocorre em duas aulas seguidas de 50 minutos ("aula_1" e "aula_2").
   - Divida o rascunho do professor em duas etapas sequenciais e práticas de 50 min.
2. RECURSOS:
   - Baseie-se nos materiais citados no rascunho e no catálogo prático: [${materialsPool}].
3. MODALIDADE ORGANIZATIVA:
   - Descreva o passo a passo metodológico de cada aula de 50 minutos em frases separadas por ponto-e-vírgula (;).
4. ACESSIBILIDADE CURRICULAR:
   - ${accessibilityInstruction}
5. INSTRUÇÕES DE ESTILO/TAMANHO NO RASCUNHO:
   - Caso o professor tenha escrito alguma instrução de estilo ou tamanho no rascunho (ex: "plano de aula curto com linguagem simples"), obedeça-a com prioridade máxima (sem copiar esse comando meta para o texto final).

Retorne APENAS um objeto JSON válido em Português (pt-BR) com a estrutura:
{
  "aula_1": { "tema_da_aula": "...", "recursos": "...", "modalidade_organizativa": "...", "agrupamento": "...", "acessibilidade_curricular": "..." },
  "aula_2": { "tema_da_aula": "...", "recursos": "...", "modalidade_organizativa": "...", "agrupamento": "...", "acessibilidade_curricular": "..." }
}`;
  } else {
    systemPrompt = `Você é um professor especialista do Programa Ensino Integral (PEI) da SEDUC-SP (EE Matilde Maria Cremm) lecionando ${block.subject} para o Ensino Médio (${block.className}).
Sua tarefa é elaborar o Plano de Aula estruturado seguindo rigorosamente as diretrizes da SEDUC-SP, redigido integralmente em Português do Brasil (pt-BR).

DIRETRIZES DE PREENCHIMENTO:
1. DOBRADINHA (SPLIT DE 50 MIN + 50 MIN):
   - Foram anexados todos os PDFs das aulas listadas pelo professor. Use as anotações do rascunho do professor para saber quais aulas/tópicos priorizar em "aula_1" e em "aula_2".
   - A aula é uma dobradinha (duas aulas seguidas de 50 minutos: "aula_1" e "aula_2").
   - A equipe gestora exige que "aula_1" e "aula_2" tenham planejamentos sequenciais e distintos (Parte 1: primeiros 50 min; Parte 2: segundos 50 min).
2. RECURSOS E RESTRIÇÃO DE HARDWARE DA SALA:
   - ${mediaConstraintInstruction}
   - Catálogo de materiais habituais do professor para ${block.subject}: [${materialsPool}].
   - Selecione apenas os recursos pertinentes à aula descrita no rascunho e acrescente novos materiais caso o rascunho cite algum experimento específico ou links de apoio.
3. HABILIDADES (BNCC / CURRÍCULO PAULISTA):
   - Se houver PDF da SEDUC anexado, extraia o código exato da habilidade (ex: EM13CNT101, EM13CNT301) e sua descrição em português.
   - Se NÃO houver PDF anexado, consulte o catálogo do Guia do Currículo Priorizado abaixo e escolha o código EM13CNT mais próximo do tema:
   ${guiaSummary ? guiaSummary : "(Use os códigos oficiais EM13CNT da BNCC de Ciências da Natureza para o Ensino Médio compatíveis com o tema)."}
4. MODALIDADE ORGANIZATIVA:
   - Escreva em tópicos ou frases separadas por ponto-e-vírgula (;), detalhando o passo a passo metodológico do professor naqueles 50 minutos.
5. ACESSIBILIDADE CURRICULAR:
   - ${accessibilityInstruction}
6. INSTRUÇÕES DE ESTILO/TAMANHO NO RASCUNHO:
   - Caso o professor tenha escrito alguma instrução de formatação, tamanho ou linguagem dentro do próprio rascunho (ex: "plano de aula curto com linguagem simples"), obedeça-a com prioridade máxima (sem incluir esse comando meta no texto final do plano).

Retorne APENAS um objeto JSON válido em Português (pt-BR) com a estrutura exata abaixo:
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
  const maxPayloadBytes = 22 * 1024 * 1024;
  encodedPdfs.forEach((pdfObj, idx) => {
    if (totalBytes + pdfObj.byteLength < maxPayloadBytes) {
      totalBytes += pdfObj.byteLength;
      userParts.push({
        text: `Arquivo PDF anexado #${idx + 1}: ${pdfObj.name}`,
      });
      userParts.push({
        inlineData: {
          mimeType: "application/pdf",
          data: pdfObj.base64,
        },
      });
    }
  });

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
    if (
      json.candidates &&
      json.candidates[0] &&
      json.candidates[0].content &&
      json.candidates[0].content.parts[0].text
    ) {
      return JSON.parse(json.candidates[0].content.parts[0].text);
    }
    if (log)
      log.error(
        `Resposta inválida da Gemini API para ${block.tabName}`,
        responseText.slice(0, 400),
      );
    return null;
  } catch (e) {
    if (log)
      log.error(
        `Erro ao fazer parse do JSON da Gemini API para ${block.tabName}`,
        e.message,
      );
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

function buildFormattedCellText(block, planPart, studentsEsp) {
  let recursos = planPart.recursos || "";
  if (block.noTv) {
    recursos = recursos.replace(/televisor(es)?\s*(,|e)?\s*/gi, "").trim();
  }

  const acessibilidade =
    studentsEsp.length === 0
      ? "N/A"
      : planPart.acessibilidade_curricular ||
        "Participação prática acompanhada.";

  if (block.isEletiva) {
    return (
      `Eletiva: ${block.eletivaTitle}\n\n` +
      `TEMA DA AULA: ${planPart.tema_da_aula || ""}\n\n` +
      `RECURSOS: ${recursos}\n\n` +
      `MODALIDADE ORGANIZATIVA: ${planPart.modalidade_organizativa || ""}\n\n` +
      `AGRUPAMENTO: ${planPart.agrupamento || "Grupos produtivos / Bancadas."}\n\n` +
      `ACESSIBILIDADE CURRICULAR: ${acessibilidade}`
    );
  }

  return (
    `Turma: ${block.className}\n\n` +
    `COMPONENTE CURRICULAR: ${block.subject}\n\n` +
    `TEMA DA AULA: ${planPart.tema_da_aula || ""}\n\n` +
    `RECURSOS: ${recursos}\n\n` +
    `CRITÉRIOS DE AVALIAÇÃO: ${planPart.criterios_avaliacao || ""}\n\n` +
    `HABILIDADES: ${planPart.habilidades || ""}\n\n` +
    `MODALIDADE ORGANIZATIVA: ${planPart.modalidade_organizativa || ""}\n\n` +
    `AGRUPAMENTO: ${planPart.agrupamento || ""}\n\n` +
    `ACESSIBILIDADE CURRICULAR: ${acessibilidade}`
  );
}
