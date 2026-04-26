// Клиентский парсер Netflix DUB_SCRIPT xlsx. Порт parser.py на JS.
// Читает один лист ("Word Count Summary") и опционально "Project Info"
// для show title. На выход даёт JSON, который сервер кладёт в БД.
//
// Использует SheetJS (загружен в /static/xlsx.full.min.js).

(function () {
  const NL = "(?<![A-Za-z])";
  const NR = "(?![A-Za-z])";
  const EPISODE_PATTERNS = [
    { re: new RegExp(`(\\d+)\\s*СЕРИЯ`, "i"), g: 1 },
    { re: new RegExp(`${NL}s\\d+\\s*e(\\d+)${NR}`, "i"), g: 1 },
    { re: new RegExp(`season\\d+\\s*episode\\s*(\\d+)`, "i"), g: 1 },
    { re: new RegExp(`${NL}episode\\s*(\\d+)`, "i"), g: 1 },
    { re: new RegExp(`${NL}ep\\.?\\s*(\\d+)${NR}`, "i"), g: 1 },
    { re: new RegExp(`${NL}e(\\d+)${NR}`, "i"), g: 1 },
    { re: new RegExp(`(\\d+)\\s*серия`, "i"), g: 1 },
    { re: new RegExp(`серия\\s*(\\d+)`, "i"), g: 1 },
    { re: new RegExp(`^[\\s_-]*(\\d+)(?=[\\s_.\\-]|$)`, ""), g: 1 },
  ];
  const SHEET_NAME = "Word Count Summary";
  const PROJECT_INFO_SHEET = "Project Info";
  const TOTAL_MARKER = "TOTAL WORD COUNT BY TEXT CATEGORY";
  // Production-metadata rows that Netflix includes in the sheet but are
  // not characters. Matched case-insensitively on the trimmed cell. Keep
  // in sync with _JUNK_CHARACTER_NAMES in app.py.
  const JUNK_CHARACTER_NAMES = new Set([
    "PRINCIPAL PHOTOGRAPHY",
    "GRAPHICS INSERTS",
    "MAIN TITLE",
  ]);
  // Колонки: 0=character, 1=dialog, 2=transcription, 3=foreign,
  // 4=music, 5=burnedin, 6=onscreen, 7=total.

  function detectEpisode(stem) {
    for (const { re, g } of EPISODE_PATTERNS) {
      const m = stem.match(re);
      if (m && m[g]) {
        const n = parseInt(m[g], 10);
        if (!isNaN(n)) return n;
      }
    }
    return null;
  }

  function readShowTitle(wb) {
    if (!wb.SheetNames.includes(PROJECT_INFO_SHEET)) return "";
    const ws = wb.Sheets[PROJECT_INFO_SHEET];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    for (const row of rows) {
      if (!row || row.length === 0 || row[0] == null) continue;
      if (String(row[0]).trim().toUpperCase() === "SHOW TITLE") {
        if (row.length > 1 && row[1] != null) return String(row[1]).trim();
        return "";
      }
    }
    return "";
  }

  function readWordCountRows(wb) {
    if (!wb.SheetNames.includes(SHEET_NAME)) {
      throw new Error(`лист "${SHEET_NAME}" отсутствует`);
    }
    const ws = wb.Sheets[SHEET_NAME];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    // Пропускаем первую строку (заголовок), идём с row[1:].
    const data = [];
    const minWidth = 8;
    for (let i = 1; i < rows.length; i++) {
      let r = rows[i] || [];
      if (r.length < minWidth) {
        r = r.concat(new Array(minWidth - r.length).fill(null));
      }
      const hasAny = r.some((v) => v !== null && v !== undefined && v !== "");
      if (!hasAny) continue;
      const first = String(r[0] ?? "").toUpperCase();
      if (first.includes(TOTAL_MARKER)) continue; // total-строку не грузим, сервер сам посчитает
      if (JUNK_CHARACTER_NAMES.has(first.trim())) continue;
      if (r[0] !== null || r.slice(1).some((v) => v)) {
        data.push(r.slice(0, minWidth));
      }
    }
    return data;
  }

  async function parseFile(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: false });
    const stem = file.name.replace(/\.xlsx$/i, "");
    let episodeNum = detectEpisode(stem);
    let singleFilm = false;
    if (episodeNum == null) {
      // Whole project is a single film — no episode number in the name.
      // Treat it as episode 1 so the rest of the pipeline (grid, report,
      // re-import) works unchanged.
      episodeNum = 1;
      singleFilm = true;
    }
    try {
      const rows = readWordCountRows(wb);
      const showTitle = readShowTitle(wb);
      return {
        filename: file.name,
        episode_num: episodeNum,
        show_title: showTitle,
        rows: rows,
        single_film: singleFilm,
      };
    } catch (e) {
      return { filename: file.name, error: String(e.message || e) };
    }
  }

  /**
   * Парсит массив File'ов клиентской xlsx-библиотекой.
   * onProgress(doneCount, totalCount, currentName) — callback для UI.
   * Возвращает { episodes: [...], warnings: [...] }.
   */
  async function parseFiles(files, onProgress) {
    const total = files.length;
    const episodes = [];
    const warnings = [];
    const seenEpisodes = new Set();

    for (let i = 0; i < total; i++) {
      const f = files[i];
      if (onProgress) onProgress(i, total, f.name);
      if (f.name.startsWith("~$") || !f.name.toLowerCase().endsWith(".xlsx")) {
        warnings.push(`${f.name}: не xlsx, пропущен`);
        continue;
      }
      const parsed = await parseFile(f);
      if (parsed.error) {
        warnings.push(`${parsed.filename}: ${parsed.error}`);
        continue;
      }
      if (parsed.single_film) {
        warnings.push(
          `${parsed.filename}: номер эпизода не найден — загружаю как фильм (серия 1)`
        );
      }
      if (seenEpisodes.has(parsed.episode_num)) {
        warnings.push(
          `${parsed.filename}: серия ${parsed.episode_num} уже была, беру более позднюю`
        );
      }
      seenEpisodes.add(parsed.episode_num);
      // При дубликате эпизода заменяем — последний выигрывает (как в parser.py).
      const idx = episodes.findIndex((e) => e.episode_num === parsed.episode_num);
      if (idx >= 0) episodes[idx] = parsed;
      else episodes.push(parsed);
    }
    if (onProgress) onProgress(total, total, "");
    return { episodes, warnings };
  }

  window.DubStudioImport = { parseFiles };
})();
