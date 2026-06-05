/**
 * XLForge Web — Report Builder
 * ---------------------------------------------------------------
 * Office Script for Excel on the web (Automate tab → New Script).
 *
 * Builds an instant summary report from the active sheet's data:
 * group-by aggregation, styled table with a total row, and a
 * column chart — like a one-click pivot report.
 *
 * Set the two column settings below, then Run.
 *
 * Author: Kristi Chakraborty · MIT License
 * https://github.com/kristic8998/xlforge-web
 */

// ── Settings ────────────────────────────────────────────────────
const GROUP_COLUMN = 3;          // 1-based column to group by (e.g., Region)
const VALUE_COLUMN = 6;          // 1-based column to aggregate (e.g., Amount)
const AGGREGATION: "sum" | "count" | "average" = "sum";
const ADD_CHART = true;
// ────────────────────────────────────────────────────────────────

function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  const used = sheet.getUsedRange();
  if (!used || used.getRowCount() < 2) {
    console.log("Need a data table with a header row and at least one data row.");
    return;
  }

  const values = used.getValues();
  const gc = GROUP_COLUMN - 1;
  const vc = VALUE_COLUMN - 1;
  const groupHeader = String(values[0][gc]);
  const valueHeader = String(values[0][vc]);

  // ── Aggregate ──
  const sums = new Map<string, number>();
  const counts = new Map<string, number>();
  for (let r = 1; r < values.length; r++) {
    const key = String(values[r][gc]) || "(blank)";
    const raw = values[r][vc];
    const num = typeof raw === "number" ? raw : parseFloat(String(raw)) || 0;
    sums.set(key, (sums.get(key) ?? 0) + num);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  const keys = [...sums.keys()].sort();
  const aggLabel = AGGREGATION[0].toUpperCase() + AGGREGATION.slice(1);

  // ── Create report sheet (unique name) ──
  let name = `Report - ${groupHeader}`.slice(0, 28);
  let candidate = name;
  let n = 1;
  while (workbook.getWorksheet(candidate)) {
    n++;
    candidate = `${name} ${n}`;
  }
  const ws = workbook.addWorksheet(candidate);

  // ── Title block ──
  ws.getRange("A1").setValue(`Summary: ${aggLabel} of ${valueHeader} by ${groupHeader}`);
  const titleFmt = ws.getRange("A1").getFormat().getFont();
  titleFmt.setSize(14);
  titleFmt.setBold(true);
  ws.getRange("A2").setValue(`Generated ${new Date().toISOString().slice(0, 16).replace("T", " ")}`);
  ws.getRange("A2").getFormat().getFont().setColor("787878");

  // ── Table ──
  const header = [[groupHeader, `${aggLabel} of ${valueHeader}`]];
  const body: (string | number)[][] = keys.map((k) => {
    const s = sums.get(k) ?? 0;
    const c = counts.get(k) ?? 0;
    const val = AGGREGATION === "sum" ? s : AGGREGATION === "count" ? c : s / c;
    return [k, Math.round(val * 100) / 100];
  });

  ws.getRange("A4:B4").setValues(header);
  const bodyRange = ws.getRangeByIndexes(4, 0, body.length, 2);
  bodyRange.setValues(body);

  // Total row (live formula, not a hardcoded number)
  const totalRow = 5 + body.length;
  ws.getRange(`A${totalRow}`).setValue("TOTAL");
  ws.getRange(`B${totalRow}`).setFormula(`=SUM(B5:B${totalRow - 1})`);

  // ── Styling ──
  const headerFmt = ws.getRange("A4:B4").getFormat();
  headerFmt.getFill().setColor("1F4E79");
  headerFmt.getFont().setColor("FFFFFF");
  headerFmt.getFont().setBold(true);

  for (let i = 0; i < body.length; i++) {
    if (i % 2 === 1) {
      ws.getRangeByIndexes(4 + i, 0, 1, 2).getFormat().getFill().setColor("EDF3FA");
    }
  }
  const totalFmt = ws.getRange(`A${totalRow}:B${totalRow}`).getFormat();
  totalFmt.getFont().setBold(true);

  ws.getRange(`B5:B${totalRow}`).setNumberFormatLocal("#,##0.00");
  ws.getRange("A:B").getFormat().autofitColumns();

  // ── Chart ──
  if (ADD_CHART && body.length > 1) {
    const chart = ws.addChart(ExcelScript.ChartType.columnClustered, bodyRange);
    chart.getTitle().setText(`${aggLabel} of ${valueHeader} by ${groupHeader}`);
    chart.setLeft(220);
    chart.setTop(60);
    chart.setWidth(420);
    chart.setHeight(260);
  }

  ws.activate();
  console.log(`Report created: ${keys.length} groups from ${values.length - 1} rows.`);
}
