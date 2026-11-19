/**
 * ============================================================================
 * LOGGERSERVICE.JS - OBSERVABILIDADE, PROFILING DE GARGALOS E RASTREIO DE CRASH
 * ============================================================================
 */

const GAS_MAX_RUNTIME_SEC = 360;
const LOG_STORAGE_KEY = "PEI_LAST_EXECUTION_LOG_V3";
const LOG_MAX_CHARS = 14000;

class ExecutionLogger {
  constructor(operationName, meta) {
    this.operationName = operationName || "Execução";
    this.runId = `${Date.now().toString(36)}-${Math.random().toString(36).substring(2, 6)}`;
    this.startMs = Date.now();
    this.timers = {};
    this.stepTotalsMs = {};
    this.entries = [];
    this.lastCheckpoint = "INITIALIZED";
    this.errorCount = 0;
    this.warnCount = 0;

    const tz =
      typeof CONFIG !== "undefined" && CONFIG.TIMEZONE
        ? CONFIG.TIMEZONE
        : "America/Sao_Paulo";
    const startedAt = new Date().toLocaleString("pt-BR", { timeZone: tz });

    this.entries.push(
      `====================================================================`,
    );
    this.entries.push(`RUN ID   : ${this.runId}`);
    this.entries.push(`OPERAÇÃO : ${this.operationName}`);
    this.entries.push(`INÍCIO   : ${startedAt}`);
    if (meta) {
      this.entries.push(`CONTEXTO : ${this._formatMeta(meta)}`);
    }
    this.entries.push(
      `====================================================================`,
    );
    this.info("Logger inicializado", { runId: this.runId });
    this.checkpoint("START");
  }

  elapsedSec() {
    return (Date.now() - this.startMs) / 1000;
  }

  _ts() {
    return `+${this.elapsedSec().toFixed(2)}s`;
  }

  _formatMeta(meta) {
    if (meta === undefined || meta === null) return "";
    if (
      typeof meta === "string" ||
      typeof meta === "number" ||
      typeof meta === "boolean"
    ) {
      return String(meta);
    }
    try {
      const raw = JSON.stringify(meta);
      return raw.length > 500 ? raw.slice(0, 497) + "..." : raw;
    } catch (e) {
      return String(meta);
    }
  }

  _record(level, stage, message, meta) {
    if (level === "ERROR") this.errorCount++;
    if (level === "WARN") this.warnCount++;

    const metaPart = meta !== undefined ? ` | ${this._formatMeta(meta)}` : "";
    const line = `${this._ts()} [${level}] [${stage}] ${message}${metaPart}`;
    this.entries.push(line);

    if (level === "ERROR") {
      console.error(line);
    } else if (level === "WARN") {
      console.warn(line);
    } else {
      console.log(line);
    }

    this._checkRuntimeBudget(stage);
  }

  _checkRuntimeBudget(currentStage) {
    const elapsed = this.elapsedSec();
    const remaining = GAS_MAX_RUNTIME_SEC - elapsed;
    if (!this._warned4m && elapsed >= 240) {
      this._warned4m = true;
      const msg = `${this._ts()} [WARN] [BUDGET] Tempo de execução atingiu 4 minutos (${remaining.toFixed(1)}s restantes antes do limite de 360s do GAS) em "${currentStage}".`;
      this.entries.push(msg);
      console.warn(msg);
      this.flushToStorage();
    } else if (!this._warned5m && elapsed >= 310) {
      this._warned5m = true;
      const msg = `${this._ts()} [WARN] [BUDGET] ALERTA CRÍTICO: Restam apenas ${remaining.toFixed(1)}s antes do timeout de 360s do GAS em "${currentStage}"!`;
      this.entries.push(msg);
      console.warn(msg);
      this.flushToStorage();
    }
  }

  info(stage, message, meta) {
    this._record("INFO", stage, message, meta);
  }

  warn(stage, message, meta) {
    this._record("WARN", stage, message, meta);
  }

  error(stage, message, errOrMeta, extraMeta) {
    let details = {};
    if (errOrMeta instanceof Error) {
      details = {
        name: errOrMeta.name,
        message: errOrMeta.message,
        stack: errOrMeta.stack
          ? String(errOrMeta.stack).split("\n").slice(0, 6).join(" || ")
          : undefined,
      };
      if (extraMeta) details.extra = extraMeta;
    } else if (errOrMeta !== undefined) {
      details = errOrMeta;
    }
    this._record("ERROR", stage, message, details);
    // Persiste imediatamente em caso de crash subsequente
    this.flushToStorage();
  }

  /**
   * Inicia medição de tempo para identificar gargalos (ex: 'Drive_PDFs', 'Batch_1').
   */
  timeStart(label, meta) {
    this.timers[label] = {
      startMs: Date.now(),
      meta: meta,
    };
    this._record("INFO", `PERF:${label}`, "Início da etapa", meta);
  }

  /**
   * Finaliza medição de tempo e acumula no resumo de gargalos.
   */
  timeEnd(label, extraMeta) {
    const t = this.timers[label];
    if (!t) {
      this.warn(
        `PERF:${label}`,
        "timeEnd chamado sem timeStart correspondente",
        extraMeta,
      );
      return 0;
    }
    const durationMs = Date.now() - t.startMs;
    const durationSec = (durationMs / 1000).toFixed(2);
    this.stepTotalsMs[label] = (this.stepTotalsMs[label] || 0) + durationMs;
    delete this.timers[label];

    const combinedMeta = Object.assign({}, t.meta || {}, extraMeta || {});
    combinedMeta.durationSec = Number(durationSec);
    this._record(
      "INFO",
      `PERF:${label}`,
      `Concluído em ${durationSec}s`,
      combinedMeta,
    );
    return durationMs;
  }

  /**
   * Salva um marco de progresso no storage para saber exatamente onde parou se houver timeout/crash.
   */
  checkpoint(checkpointName, meta) {
    this.lastCheckpoint = checkpointName;
    const metaStr = meta !== undefined ? ` | ${this._formatMeta(meta)}` : "";
    const line = `${this._ts()} [CHECKPOINT] >>> ${checkpointName}${metaStr}`;
    this.entries.push(line);
    console.log(line);
    this.flushToStorage();
  }

  /**
   * Gera sumário de gargalos ordenados do mais lento para o mais rápido.
   */
  buildSummarySection() {
    const totalMs = Math.max(1, Date.now() - this.startMs);
    const totalSec = (totalMs / 1000).toFixed(2);
    const lines = [];
    lines.push(
      `====================================================================`,
    );
    lines.push(`RESUMO DE EXECUÇÃO E GARGALOS`);
    lines.push(
      `Tempo Total   : ${totalSec}s / ${GAS_MAX_RUNTIME_SEC}s (${((totalMs / 1000 / GAS_MAX_RUNTIME_SEC) * 100).toFixed(1)}% do limite GAS)`,
    );
    lines.push(`Último Marco  : ${this.lastCheckpoint}`);
    lines.push(
      `Warnings/Erros: ${this.warnCount} warn(s) | ${this.errorCount} error(s)`,
    );

    const labels = Object.keys(this.stepTotalsMs).sort(
      (a, b) => this.stepTotalsMs[b] - this.stepTotalsMs[a],
    );
    if (labels.length > 0) {
      lines.push(`--- Gargalos por Etapa (maior -> menor) ---`);
      labels.forEach((lbl) => {
        const ms = this.stepTotalsMs[lbl];
        const sec = (ms / 1000).toFixed(2);
        const pct = ((ms / totalMs) * 100).toFixed(1);
        lines.push(`  • ${lbl}: ${sec}s (${pct}%)`);
      });
    }
    lines.push(
      `====================================================================`,
    );
    return lines.join("\n");
  }

  flushToStorage() {
    try {
      const summary = this.buildSummarySection();
      const body = this.entries.join("\n");
      const combined = `${summary}\n\n${body}`;
      const bounded =
        combined.length > LOG_MAX_CHARS
          ? `[...TRUNCADO PARA CABER NO STORAGE...]\n` +
            combined.slice(-LOG_MAX_CHARS)
          : combined;
      PropertiesService.getScriptProperties().setProperty(
        LOG_STORAGE_KEY,
        bounded,
      );
    } catch (storageErr) {
      console.error("Falha ao salvar log em ScriptProperties:", storageErr);
    }
  }
}

function showLastExecutionLogs() {
  const ui = SpreadsheetApp.getUi();
  const rawLog =
    PropertiesService.getScriptProperties().getProperty(LOG_STORAGE_KEY);

  if (!rawLog) {
    ui.alert(
      "📋 Logs da Última Execução",
      "Nenhum log detalhado encontrado ainda.\n\n" +
        "Dica: Você também pode abrir Extensões > Apps Script > Execuções (ícone de relógio/lista à esquerda) para ver os logs nativos do Cloud Logging.",
    );
    return;
  }

  const escaped = String(rawLog)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(
      /\[ERROR\]/g,
      '<span style="color:#f48771;font-weight:bold;">[ERROR]</span>',
    )
    .replace(
      /\[WARN\]/g,
      '<span style="color:#cca700;font-weight:bold;">[WARN]</span>',
    )
    .replace(
      /\[CHECKPOINT\]/g,
      '<span style="color:#4ec9b0;font-weight:bold;">[CHECKPOINT]</span>',
    );

  const html = HtmlService.createHtmlOutput(
    `<div style="font-family: 'JetBrains Mono', Consolas, monospace; font-size: 12px; line-height: 1.45; white-space: pre-wrap; background: #1e1e1e; color: #d4d4d4; padding: 14px; border-radius: 6px; overflow-y: auto; height: 92%;">${escaped}</div>`,
  )
    .setWidth(780)
    .setHeight(540);

  ui.showModalDialog(html, "📋 Logs de Diagnóstico e Gargalos");
}
