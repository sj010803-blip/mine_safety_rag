import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const payload = JSON.parse(await fs.readFile(process.argv[2], "utf8"));
const outputDir = process.argv[3];
const qaDir = process.argv[4] || "";

const COLORS = {
  navy: "#1F4E78",
  darkNavy: "#17365D",
  white: "#FFFFFF",
  text: "#111827",
  muted: "#475569",
  paleBlue: "#D9EAF7",
  paleYellow: "#FFF4CE",
  paleGreen: "#E2F0D9",
  paleRed: "#FCE8E6",
  border: "#CBD5E1",
};

const EVALUATION_TITLE = "질문 context 및 미완성 답변 안전장치 수정 후 회귀평가";
const EVALUATION_NOTICE = "동일한 30문항 재사용 · 기존 기준선과의 변화 확인용 · 최종 잠금 평가가 아님 · 법적 정확도 및 최종 답변 정확도를 직접 의미하지 않음";

function columnName(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function widthForHeader(header) {
  const text = String(header);
  if (/question|expected|reason|excerpt|comment|주의|방법|설명|문구|문서|내용/.test(text)) return 34;
  if (/metadata|patterns|terms|sha256|expanded/.test(text)) return 28;
  if (/category|분류|유형/.test(text)) return 23;
  if (/eval_id|rank|count|Hit|MRR|coverage|완전성|비율|난이도|값|delta|증감/.test(text)) return 14;
  return 18;
}

function setTitleBand(sheet, columnCount, title) {
  const lastColumn = columnName(Math.max(columnCount - 1, 0));
  const titleRange = sheet.getRange(`A1:${lastColumn}1`);
  titleRange.merge();
  titleRange.values = [[title]];
  titleRange.format = {
    fill: COLORS.darkNavy,
    font: { name: "Malgun Gothic", size: 15, bold: true, color: COLORS.white },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    rowHeight: 30,
  };
  const noticeRange = sheet.getRange(`A2:${lastColumn}2`);
  noticeRange.merge();
  noticeRange.values = [[EVALUATION_NOTICE]];
  noticeRange.format = {
    fill: COLORS.paleYellow,
    font: { name: "Malgun Gothic", size: 9, color: COLORS.text },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    wrapText: true,
    rowHeight: 38,
    borders: { preset: "outside", style: "thin", color: COLORS.border },
  };
}

function addTableSheet(workbook, name, headers, rows, options = {}) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const safeRows = rows.length ? rows : [Object.fromEntries(headers.map((header) => [header, null]))];
  setTitleBand(sheet, Math.max(headers.length, 8), options.title || `${EVALUATION_TITLE} · ${name}`);
  const values = [headers, ...safeRows.map((row) => headers.map((header) => row[header] ?? null))];
  const tableRange = sheet.getRangeByIndexes(3, 0, values.length, headers.length);
  tableRange.values = values;
  tableRange.format = {
    font: { name: "Malgun Gothic", size: 10, color: COLORS.text },
    verticalAlignment: "top",
    wrapText: true,
  };
  const headerRange = sheet.getRangeByIndexes(3, 0, 1, headers.length);
  headerRange.format = {
    fill: COLORS.navy,
    font: { name: "Malgun Gothic", size: 10, bold: true, color: COLORS.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeight: 32,
    borders: { preset: "outside", style: "thin", color: COLORS.border },
  };
  sheet.freezePanes.freezeRows(4);
  for (let index = 0; index < headers.length; index += 1) {
    const header = String(headers[index]);
    const columnRange = sheet.getRangeByIndexes(3, index, values.length, 1);
    columnRange.format.columnWidth = widthForHeader(header);
    if (/hit_at|coverage|reciprocal|mrr|completeness|ratio|baseline|postfix|delta|증감/.test(header.toLowerCase())) {
      if (values.length > 1) sheet.getRangeByIndexes(4, index, values.length - 1, 1).format.numberFormat = "0.000";
    } else if (/distance/.test(header.toLowerCase())) {
      if (values.length > 1) sheet.getRangeByIndexes(4, index, values.length - 1, 1).format.numberFormat = "0.0000";
    }
  }
  if (values.length > 1) {
    sheet.getRangeByIndexes(4, 0, values.length - 1, headers.length).format.rowHeight = options.rowHeight ?? 38;
  }
  if (options.validation && values.length > 1) {
    for (const validation of options.validation) {
      const index = headers.indexOf(validation.header);
      if (index >= 0) {
        sheet.getRangeByIndexes(4, index, values.length - 1, 1).dataValidation = {
          rule: { type: "list", values: validation.values },
        };
      }
    }
  }
  return sheet;
}

function addKeyValueSheet(workbook, name, rows, title) {
  const sheet = addTableSheet(workbook, name, ["항목", "값"], rows, { rowHeight: 44, title });
  const lastRow = 4 + Math.max(rows.length, 1);
  sheet.getRange(`A5:A${lastRow}`).format.font = { name: "Malgun Gothic", size: 10, bold: true, color: COLORS.text };
  sheet.getRange(`B5:B${lastRow}`).format.horizontalAlignment = "right";
  sheet.getRange(`A4:A${lastRow}`).format.columnWidth = 28;
  sheet.getRange(`B4:B${lastRow}`).format.columnWidth = 56;
  return sheet;
}

function addComparisonSummary(workbook) {
  const sheet = workbook.worksheets.add("00_비교요약");
  sheet.showGridLines = false;
  setTitleBand(sheet, 14, "수정 전 자동 기준선 vs 수정 후 회귀평가");
  const rateRows = payload.comparisonRateRows;
  sheet.getRange("A4:D4").values = [["지표", "수정 전 자동 기준선", "수정 후 회귀평가", "증감"]];
  sheet.getRangeByIndexes(4, 0, rateRows.length, 3).values = rateRows.map((row) => [row.metric, row.baseline, row.postfix]);
  sheet.getRange("D5").formulas = [["=C5-B5"]];
  sheet.getRange(`D5:D${4 + rateRows.length}`).fillDown();
  const rateRange = sheet.getRange(`A4:D${4 + rateRows.length}`);
  rateRange.format = {
    font: { name: "Malgun Gothic", size: 10, color: COLORS.text },
    borders: { preset: "inside", style: "thin", color: COLORS.border },
  };
  sheet.getRange("A4:D4").format = {
    fill: COLORS.navy,
    font: { name: "Malgun Gothic", size: 10, bold: true, color: COLORS.white },
    horizontalAlignment: "center",
    rowHeight: 30,
  };
  sheet.getRange(`B5:D${4 + rateRows.length}`).format.numberFormat = "0.0000";
  sheet.getRange(`D5:D${4 + rateRows.length}`).conditionalFormats.add("cellIs", {
    operator: "greaterThan",
    formula: 0,
    format: { fill: COLORS.paleGreen, font: { color: "#166534", bold: true } },
  });
  sheet.getRange(`D5:D${4 + rateRows.length}`).conditionalFormats.add("cellIs", {
    operator: "lessThan",
    formula: 0,
    format: { fill: COLORS.paleRed, font: { color: "#991B1B", bold: true } },
  });
  const countHeaderRow = 6 + rateRows.length;
  const countStartRow = countHeaderRow + 1;
  sheet.getRange(`A${countHeaderRow}:D${countHeaderRow}`).values = [["건수 지표", "수정 전 자동 기준선", "수정 후 회귀평가", "증감"]];
  sheet.getRangeByIndexes(countStartRow - 1, 0, payload.comparisonCountRows.length, 3).values = payload.comparisonCountRows.map((row) => [row.metric, row.baseline, row.postfix]);
  sheet.getRange(`D${countStartRow}`).formulas = [[`=C${countStartRow}-B${countStartRow}`]];
  sheet.getRange(`D${countStartRow}:D${countStartRow + payload.comparisonCountRows.length - 1}`).fillDown();
  sheet.getRange(`A${countHeaderRow}:D${countHeaderRow}`).format = {
    fill: COLORS.navy,
    font: { name: "Malgun Gothic", size: 10, bold: true, color: COLORS.white },
    horizontalAlignment: "center",
    rowHeight: 30,
  };
  sheet.getRange(`A${countStartRow}:D${countStartRow + payload.comparisonCountRows.length - 1}`).format = {
    font: { name: "Malgun Gothic", size: 10, color: COLORS.text },
    numberFormat: "0",
  };
  const comparisonLastRow = countStartRow + payload.comparisonCountRows.length - 1;
  sheet.getRange(`A4:A${comparisonLastRow}`).format.columnWidth = 28;
  sheet.getRange(`B4:D${comparisonLastRow}`).format.columnWidth = 20;
  const chart = sheet.charts.add("bar", sheet.getRange(`A4:C${4 + rateRows.length}`));
  chart.title = "주요 검색 지표 비교 (0~1)";
  chart.titleTextStyle.fontSize = 12;
  chart.hasLegend = true;
  chart.xAxis = { axisType: "textAxis", textStyle: { fontSize: 9 } };
  chart.yAxis = { numberFormatCode: "0.00", min: 0, max: 1 };
  chart.setPosition("F4", "N21");
  sheet.freezePanes.freezeRows(4);
  return sheet;
}

async function verifyAndRender(workbook, label) {
  if (!qaDir) return;
  await fs.mkdir(qaDir, { recursive: true });
  const overview = await workbook.inspect({
    kind: "workbook,sheet,table,drawing",
    maxChars: 10000,
    tableMaxRows: 8,
    tableMaxCols: 10,
    tableMaxCellChars: 100,
  });
  await fs.writeFile(path.join(qaDir, `${label}.inspect.ndjson`), overview.ndjson, "utf8");
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: `${label} formula error scan`,
  });
  await fs.writeFile(path.join(qaDir, `${label}.errors.ndjson`), errors.ndjson, "utf8");
  for (const sheet of workbook.worksheets.items) {
    const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 1, format: "png" });
    const safeName = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
    await fs.writeFile(path.join(qaDir, `${label}__${safeName}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
}

await fs.mkdir(outputDir, { recursive: true });

const results = Workbook.create();
addKeyValueSheet(results, "00_요약", payload.summaryRows, `${EVALUATION_TITLE} · 전체 요약`);
addTableSheet(results, "01_30문항_질문세트", payload.questionHeaders, payload.questionRows);
addTableSheet(results, "02_질문별_검색결과", payload.resultHeaders, payload.resultRows);
addTableSheet(results, "03_Top5_상세", payload.detailHeaders, payload.detailRows);
addTableSheet(results, "04_유형별_결과", payload.categoryHeaders, payload.categoryRows);
addTableSheet(results, "05_난이도별_결과", payload.difficultyHeaders, payload.difficultyRows);
addTableSheet(results, "06_문서별_검색빈도", payload.frequencyHeaders, payload.frequencyRows);
addTableSheet(results, "07_핵심요소_포함률", payload.elementHeaders, payload.elementRows);
addTableSheet(results, "08_수동검토_대상", payload.reviewHeaders, payload.reviewRows);
addTableSheet(results, "09_평가방법_주의사항", ["구분", "내용"], payload.methodRows, { rowHeight: 48 });
await verifyAndRender(results, "postfix_results");
const resultsFile = await SpreadsheetFile.exportXlsx(results);
await resultsFile.save(path.join(outputDir, "rag_retrieval_postfix_results_30.xlsx"));

const manual = Workbook.create();
addTableSheet(manual, "수정후_수동검토", payload.manualHeaders, payload.manualRows, {
  rowHeight: 58,
  title: `${EVALUATION_TITLE} · 수동 검토용`,
  validation: [
    { header: "human_relevance", values: ["2", "1", "0"] },
    { header: "human_document_match", values: ["Y", "N", "REVIEW"] },
  ],
});
await verifyAndRender(manual, "postfix_manual_review");
const manualFile = await SpreadsheetFile.exportXlsx(manual);
await manualFile.save(path.join(outputDir, "rag_retrieval_postfix_manual_review_30.xlsx"));

const summary = Workbook.create();
addKeyValueSheet(summary, "전체요약", payload.summaryRows, `${EVALUATION_TITLE} · 전체 요약`);
addTableSheet(summary, "유형별", payload.categoryHeaders, payload.categoryRows);
addTableSheet(summary, "난이도별", payload.difficultyHeaders, payload.difficultyRows);
addTableSheet(summary, "문서별검색빈도", payload.frequencyHeaders, payload.frequencyRows);
addTableSheet(summary, "평가해석", ["지표", "해석"], payload.interpretationRows, { rowHeight: 48 });
await verifyAndRender(summary, "postfix_summary");
const summaryFile = await SpreadsheetFile.exportXlsx(summary);
await summaryFile.save(path.join(outputDir, "rag_retrieval_postfix_summary.xlsx"));

const comparison = Workbook.create();
const comparisonSheet = addComparisonSummary(comparison);
addTableSheet(comparison, "01_지표비교", payload.comparisonHeaders, payload.comparisonRows, {
  title: "수정 전 자동 기준선 vs 수정 후 회귀평가 · 전체 지표",
});
addTableSheet(comparison, "02_유형별_Hit3", payload.categoryComparisonHeaders, payload.categoryComparisonRows, {
  title: "수정 전 자동 기준선 vs 수정 후 회귀평가 · 유형별 Hit@3",
});
addTableSheet(comparison, "03_난이도별_Hit3", payload.difficultyComparisonHeaders, payload.difficultyComparisonRows, {
  title: "수정 전 자동 기준선 vs 수정 후 회귀평가 · 난이도별 Hit@3",
});
await verifyAndRender(comparison, "baseline_vs_postfix");
const comparisonPreview = await comparison.render({
  sheetName: comparisonSheet.name,
  range: "A1:N22",
  scale: 2,
  format: "png",
});
await fs.writeFile(
  path.join(outputDir, "rag_retrieval_baseline_vs_postfix.png"),
  new Uint8Array(await comparisonPreview.arrayBuffer()),
);
const comparisonFile = await SpreadsheetFile.exportXlsx(comparison);
await comparisonFile.save(path.join(outputDir, "rag_retrieval_baseline_vs_postfix.xlsx"));
