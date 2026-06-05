/**
 * XLForge Web — Data Cleaner
 * ---------------------------------------------------------------
 * Office Script for Excel on the web (Automate tab → New Script).
 *
 * One run on the selected sheet's used range:
 *   1. Trims leading/trailing whitespace, collapses double spaces
 *   2. Removes non-printing characters and non-breaking spaces
 *   3. Converts text dates ("2026-01-15", "1/15/2026", "Jan 15 2026")
 *      to real Excel dates in columns that are mostly date-like
 *   4. Applies Proper Case to text columns (toggle below)
 *   5. Removes exact duplicate rows (keeps first occurrence)
 *
 * A summary of every change is written to a "CleanLog" sheet.
 *
 * Author: Kristi Chakraborty · MIT License
 * https://github.com/kristic8998/xlforge-web
 */

// ── Settings ────────────────────────────────────────────────────
const HAS_HEADER_ROW = true;   // first row is headers (never modified)
const APPLY_PROPER_CASE = true;
const DATE_COLUMN_THRESHOLD = 0.6; // 60%+ parseable text = date column
// ────────────────────────────────────────────────────────────────

function main(workbook: ExcelScript.Workbook) {
  const started = Date.now();
  const sheet = workbook.getActiveWorksheet();
  const used = sheet.getUsedRange();
  if (!used) {
    console.log("Sheet is empty — nothing to clean.");
    return;
  }

  const values = used.getValues();
  const log: string[] = [];
  const headerOffset = HAS_HEADER_ROW ? 1 : 0;

  // ── 1+2: trim, collapse spaces, strip control chars & NBSP ──
  let textFixes = 0;
  for (let r = headerOffset; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const v = values[r][c];
      if (typeof v === "string") {
        const cleaned = v
          .replace(/[\u0000-\u001F\u007F]/g, "") // control chars
          .replace(/\u00A0/g, " ")                  // non-breaking space
          .replace(/\s+/g, " ")                     // collapse runs
          .trim();
        if (cleaned !== v) {
          values[r][c] = cleaned;
          textFixes++;
        }
      }
    }
  }
  log.push(`Whitespace/character fixes: ${textFixes} cells`);

  // ── 3: detect & convert text-date columns ──
  const colCount = values[0].length;
  let dateFixes = 0;
  const dateCols: number[] = [];
  for (let c = 0; c < colCount; c++) {
    let textCells = 0;
    let parseable = 0;
    for (let r = headerOffset; r < values.length; r++) {
      const v = values[r][c];
      if (typeof v === "string" && v.length > 0) {
        textCells++;
        if (parseDate(v) !== null) parseable++;
      }
    }
    if (textCells > 0 && parseable / textCells >= DATE_COLUMN_THRESHOLD) {
      dateCols.push(c);
      for (let r = headerOffset; r < values.length; r++) {
        const v = values[r][c];
        if (typeof v === "string") {
          const serial = parseDate(v);
          if (serial !== null) {
            values[r][c] = serial; // Excel date serial number
            dateFixes++;
          }
        }
      }
    }
  }
  log.push(`Date conversions: ${dateFixes} cells in ${dateCols.length} column(s)`);

  // ── 4: Proper Case for remaining text cells ──
  let caseFixes = 0;
  if (APPLY_PROPER_CASE) {
    for (let r = headerOffset; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const v = values[r][c];
        if (typeof v === "string" && v.length > 0) {
          const proper = v.replace(/\w\S*/g,
            (t) => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase());
          if (proper !== v) {
            values[r][c] = proper;
            caseFixes++;
          }
        }
      }
    }
    log.push(`Proper Case fixes: ${caseFixes} cells`);
  }

  // ── 5: remove exact duplicate rows ──
  const seen = new Set<string>();
  const deduped: (string | number | boolean)[][] = [];
  let dupes = 0;
  for (let r = 0; r < values.length; r++) {
    if (r < headerOffset) {
      deduped.push(values[r]);
      continue;
    }
    const key = JSON.stringify(values[r]);
    if (seen.has(key)) {
      dupes++;
    } else {
      seen.add(key);
      deduped.push(values[r]);
    }
  }
  log.push(`Duplicate rows removed: ${dupes}`);

  // ── Write back ──
  used.clear(ExcelScript.ClearApplyTo.contents);
  const out = sheet
    .getRangeByIndexes(used.getRowIndex(), used.getColumnIndex(), deduped.length, colCount);
  out.setValues(deduped);

  // Format converted date columns
  for (const c of dateCols) {
    sheet.getRangeByIndexes(
      used.getRowIndex() + headerOffset,
      used.getColumnIndex() + c,
      deduped.length - headerOffset,
      1
    ).setNumberFormatLocal("yyyy-mm-dd");
  }

  // ── Audit log sheet ──
  const elapsed = ((Date.now() - started) / 1000).toFixed(2);
  log.push(`Completed in ${elapsed} s`);
  writeLog(workbook, sheet.getName(), log);
  log.forEach((l) => console.log(l));
}

/** Parse common date strings; returns Excel serial number or null. */
function parseDate(s: string): number | null {
  const t = s.trim();
  // Reject pure numbers ("123") and too-short strings
  if (t.length < 6 || /^\d+(\.\d+)?$/.test(t)) return null;

  let y = 0, m = 0, d = 0;
  let match: RegExpMatchArray | null;

  if ((match = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) {
    [y, m, d] = [+match[1], +match[2], +match[3]];
  } else if ((match = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) {
    [m, d, y] = [+match[1], +match[2], +match[3]];
  } else if ((match = t.match(/^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})$/))) {
    const months = ["jan", "feb", "mar", "apr", "may", "jun",
                    "jul", "aug", "sep", "oct", "nov", "dec"];
    const mi = months.indexOf(match[1].slice(0, 3).toLowerCase());
    if (mi < 0) return null;
    [m, d, y] = [mi + 1, +match[2], +match[3]];
  } else {
    return null;
  }

  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1900 || y > 2200) return null;
  // Convert to Excel serial (days since 1899-12-30, UTC to avoid TZ drift)
  const ms = Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30);
  return Math.round(ms / 86400000);
}

/** Append run summary to a CleanLog sheet (created if missing). */
function writeLog(workbook: ExcelScript.Workbook, source: string, lines: string[]) {
  let logSheet = workbook.getWorksheet("CleanLog");
  if (!logSheet) {
    logSheet = workbook.addWorksheet("CleanLog");
    logSheet.getRange("A1:C1").setValues([["Timestamp", "Sheet", "Result"]]);
    logSheet.getRange("A1:C1").getFormat().getFont().setBold(true);
  }
  const usedLog = logSheet.getUsedRange();
  let nextRow = usedLog ? usedLog.getRowCount() : 1;
  const stamp = new Date().toISOString().replace("T", " ").slice(0, 19);
  const rows = lines.map((l) => [stamp, source, l]);
  logSheet
    .getRangeByIndexes(nextRow, 0, rows.length, 3)
    .setValues(rows);
  logSheet.getRange("A:C").getFormat().autofitColumns();
}
