# Project Rules & Architecture: PEI Lecture Plans Automation

## 1. Spreadsheet Architecture (`Minha Agenda EE. MATILDE`)

- **Template vs. Active Grid:** `PA (Base)` holds the fixed weekly timetable and boilerplate. `PA (Atual)` is overwritten when preparing a new week (`Row 4` = Dates `DD/MM/YY\nSEG..SEX`, `Rows 5–12` = `1ª Aula` to `8ª Aula`).
- **Class Tabs (`1º HO`, `1º HU`, `2º LE`, `2º SO`, `3º AL`, `3º RE`, `Eletiva`):**
  - The draft table starts dynamically at the row where Column A begins with `Conteúdos dados:`.
  - Expected columns: `Conteúdos dados:` | `Aulas` | `Data` | `Disciplina` | `Bimestre`.
  - If `Disciplina` is omitted (e.g., `2º SO`, `3º RE`), default to `Física`.
  - `Física` and `Robótica` rows are filtered and queued independently within each class tab.
- **Draft Matching Priority:**
  1. Exact match on `Data` (`DD/MM`) against the target day in `PA (Atual)`.
  2. First undated row (`Data` empty or `"POR FAZER"`) appearing _after_ the last row whose date is $\le$ the target week date. Once generated, the script auto-populates `Data` with `DD/MMM` (e.g., `23/Nov`).
- **Idempotency:** Any cell in `PA (Atual)` that already contains text inside `MODALIDADE ORGANIZATIVA:` (> 5 chars) is skipped.

## 2. Domain & Pedagogical Constraints

- **Language:** All generated lesson plan fields MUST be written in Brazilian Portuguese (`pt-BR`).
- **Double Lectures (`Dobradinhas`):** Consecutive periods for the same class and subject on the same day are grouped into a single API call returning `{ "aula_1": {...}, "aula_2": {...} }`, splitting the lesson into sequential 50-minute stages.
- **No-TV Classrooms (`1º HO`, `2º LE`, `3º RE`):**
  - NEVER include `Televisor` in `RECURSOS` (enforced both in prompt and via post-processing regex).
  - Only include `Projetor` if the draft mentions PhET simulations, YouTube, or videos; otherwise default to printed materials (`Material Impresso`).
- **Special Education (`Educ. Esp.` tab):**
  - Dynamically read student profiles per class (and cross-reference names against the `Eletiva` student roster).
  - Classes with no listed students MUST output `ACESSIBILIDADE CURRICULAR: N/A`.
- **Eletiva Tab:** Uses a simplified cell layout omitting `HABILIDADES` and `CRITÉRIOS DE AVALIAÇÃO`.

## 3. Drive PDF & Curriculum Resolution

- **Robotics 2x1 Rule:** Each SEDUC Robotics PDF covers 2 lessons ($\text{PDF} = \lceil \text{Aula} / 2 \rceil$). Writing `A5, A6` resolves to `PDF 3 (Aulas 05 e 06)`. Writing `P3` or `PDF 3` directly resolves to `PDF 3`.
- **Bimester Resolution:** Resolved in order from: (1) `Bimestre` column, (2) inline tag like `B3 A5`, or (3) `CONFIG.CURRENT_BIMESTRE`.
- **Fallback Skill Matching:** When no lecture PDF is attached (e.g., review/exam classes), match `EM13CNT...` skill codes from the cached `Guia do Currículo Priorizado` PDF in Drive.

## 4. Technical & Security Guardrails

- **Never hardcode secrets:** Store `GEMINI_API_KEY`, `AULAS_FOLDER_ID`, and `GUIA_FOLDER_ID` exclusively in `PropertiesService.getScriptProperties()`.
- **Parallel Execution:** Always use `UrlFetchApp.fetchAll()` in batches of `CONFIG.BATCH_SIZE` (`5`) alongside an in-memory `pdfBase64Cache` keyed by Drive `fileId` to stay well under the 6-minute Apps Script execution limit.
- **Observability:** Use `ExecutionLogger` in all entry-point functions so errors and HTTP status codes are persisted to `LAST_EXECUTION_LOG` and viewable from the custom menu.
