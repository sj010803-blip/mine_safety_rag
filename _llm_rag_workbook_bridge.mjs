import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const [, , mode, templatePath, payloadPath, outputPath] = process.argv;

if (!mode || !templatePath || !payloadPath) {
  throw new Error("usage: bridge.mjs <extract|build> <template.xlsx> <payload.json> [output.xlsx]");
}

const templateBlob = await FileBlob.load(templatePath);
const workbook = await SpreadsheetFile.importXlsx(templateBlob);

function usedValues(sheetName) {
  const sheet = workbook.worksheets.getItem(sheetName);
  const used = sheet.getUsedRange(true);
  return used ? used.values : [];
}

if (mode === "extract") {
  const data = {};
  for (const sheetName of [
    "1_실험개요",
    "2_질문목록_20개",
    "3_답변원문_붙여넣기",
    "4_비교평가",
    "5_평균요약",
    "6_채점기준",
  ]) {
    data[sheetName] = usedValues(sheetName);
  }
  await fs.writeFile(payloadPath, JSON.stringify(data), "utf8");
  process.exit(0);
}

if (mode !== "build" || !outputPath) {
  throw new Error("build mode requires output path");
}

const payload = JSON.parse(await fs.readFile(payloadPath, "utf8"));
const NAVY = "#1F4E78";
const BLUE = "#DCE6F1";
const LIGHT = "#F4F7FA";
const GREEN = "#E2F0D9";
const AMBER = "#FFF2CC";
const RED = "#FCE4D6";
const BORDER = "#D9E2F3";
const WHITE = "#FFFFFF";
const TEXT = "#24364B";

function deleteTables(sheet) {
  for (const table of [...sheet.tables.items]) {
    table.delete();
  }
}

function resetSheet(sheet) {
  deleteTables(sheet);
  sheet.deleteAllDrawings();
  const used = sheet.getUsedRange();
  if (used) {
    used.clear({ applyTo: "all" });
  }
  sheet.showGridLines = false;
}

function writeMatrix(sheet, startCell, matrix) {
  if (!matrix.length || !matrix[0].length) return;
  sheet.getRange(startCell).write(matrix);
}

function styleHeader(range) {
  range.format = {
    fill: NAVY,
    font: { bold: true, color: WHITE },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: BORDER },
  };
  range.format.rowHeight = 28;
}

function styleBody(range) {
  range.format = {
    font: { color: TEXT, size: 9 },
    verticalAlignment: "top",
    wrapText: true,
    borders: {
      insideHorizontal: { style: "thin", color: BORDER },
      bottom: { style: "thin", color: BORDER },
    },
  };
}

function setColumnWidths(sheet, widths) {
  for (const [column, widthPx] of Object.entries(widths)) {
    sheet.getRange(`${column}:${column}`).format.columnWidthPx = widthPx;
  }
}

function addTable(sheet, address, name) {
  const table = sheet.tables.add(address, true, name);
  table.style = "TableStyleMedium2";
  table.showBandedColumns = false;
  table.showFilterButton = true;
  return table;
}

const rawSheet = workbook.worksheets.getItem("3_답변원문_붙여넣기");
resetSheet(rawSheet);
writeMatrix(rawSheet, "A1", payload.raw_answers);
styleHeader(rawSheet.getRange("A1:H1"));
styleBody(rawSheet.getRange(`A2:H${payload.raw_answers.length}`));
rawSheet.getRange(`E2:E${payload.raw_answers.length}`).format.rowHeight = 95;
setColumnWidths(rawSheet, {
  A: 70, B: 115, C: 330, D: 105, E: 620, F: 100, G: 150, H: 150,
});
rawSheet.freezePanes.freezeRows(1);
rawSheet.freezePanes.freezeColumns(4);
addTable(rawSheet, `A1:H${payload.raw_answers.length}`, "ComparisonRawAnswers");

const evalSheet = workbook.worksheets.getItem("4_비교평가");
resetSheet(evalSheet);
writeMatrix(evalSheet, "A1", payload.evaluations);
styleHeader(evalSheet.getRange("A1:N1"));
styleBody(evalSheet.getRange(`A2:N${payload.evaluations.length}`));
evalSheet.getRange(`E2:I${payload.evaluations.length}`).format.horizontalAlignment = "center";
evalSheet.getRange(`J2:J${payload.evaluations.length}`).format.horizontalAlignment = "center";
evalSheet.getRange(`N2:N${payload.evaluations.length}`).format.horizontalAlignment = "center";
evalSheet.getRange(`A2:N${payload.evaluations.length}`).format.rowHeight = 78;
evalSheet.getRange(`A2:N${payload.evaluations.length}`).conditionalFormats.add(
  "expression",
  {
    formula: "=$N2=\"Y\"",
    format: { fill: AMBER },
  },
);
evalSheet.getRange(`I2:I${payload.evaluations.length}`).conditionalFormats.add(
  "cellIs",
  {
    operator: "greaterThanOrEqual",
    formula: 90,
    format: { fill: GREEN, font: { color: "#215E21", bold: true } },
  },
);
evalSheet.getRange(`I2:I${payload.evaluations.length}`).conditionalFormats.add(
  "cellIs",
  {
    operator: "lessThan",
    formula: 60,
    format: { fill: RED, font: { color: "#9C0006" } },
  },
);
setColumnWidths(evalSheet, {
  A: 70, B: 115, C: 300, D: 100, E: 90, F: 90, G: 125, H: 85,
  I: 85, J: 90, K: 240, L: 240, M: 360, N: 85,
});
evalSheet.freezePanes.freezeRows(1);
evalSheet.freezePanes.freezeColumns(4);
addTable(evalSheet, `A1:N${payload.evaluations.length}`, "ComparisonEvaluation");

const summarySheet = workbook.worksheets.getItem("5_평균요약");
resetSheet(summarySheet);
summarySheet.getRange("A1:P1").merge();
summarySheet.getRange("A1").values = [["비교 결과 평균 요약"]];
summarySheet.getRange("A1:P1").format = {
  fill: NAVY,
  font: { bold: true, color: WHITE, size: 15 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
summarySheet.getRange("A1:P1").format.rowHeight = 34;
summarySheet.getRange("A2:P2").merge();
summarySheet.getRange("A2").values = [[
  "※ 본 결과는 규칙 기반 1차 자동평가이며, 최종 평가는 대표 문항 수동 검토가 필요함.",
]];
summarySheet.getRange("A2:P2").format = {
  fill: AMBER,
  font: { color: "#7F6000", bold: true },
  wrapText: true,
  verticalAlignment: "center",
};
summarySheet.getRange("A2:P2").format.rowHeight = 30;
writeMatrix(summarySheet, "A4", payload.summary);
styleHeader(summarySheet.getRange("A4:F4"));
styleBody(summarySheet.getRange("A5:F7"));
summarySheet.getRange("B5:F7").format.numberFormat = "0.00";
summarySheet.getRange("B5:F7").format.horizontalAlignment = "center";
setColumnWidths(summarySheet, { A: 120, B: 130, C: 130, D: 170, E: 115, F: 115 });
addTable(summarySheet, "A4:F7", "ComparisonSummary");

writeMatrix(summarySheet, "H4", payload.total_chart);
styleHeader(summarySheet.getRange("H4:I4"));
styleBody(summarySheet.getRange("H5:I7"));
summarySheet.getRange("I5:I7").format.numberFormat = "0.00";
writeMatrix(summarySheet, "J4", payload.criteria_chart);
styleHeader(summarySheet.getRange("J4:N4"));
styleBody(summarySheet.getRange("J5:N7"));
summarySheet.getRange("K5:N7").format.numberFormat = "0.00";
setColumnWidths(summarySheet, {
  H: 115, I: 95, J: 115, K: 95, L: 95, M: 125, N: 95,
});

const totalChart = summarySheet.charts.add("bar", summarySheet.getRange("H4:I7"));
totalChart.title = "비교대상별 총점 평균";
totalChart.titleTextStyle.fontSize = 13;
totalChart.hasLegend = false;
totalChart.yAxis = { numberFormatCode: "0.0", min: 0, max: 100 };
totalChart.setPosition("A10", "H29");

const criteriaChart = summarySheet.charts.add("bar", summarySheet.getRange("J4:N7"));
criteriaChart.title = "평가 기준별 평균 점수 비교";
criteriaChart.titleTextStyle.fontSize = 13;
criteriaChart.hasLegend = true;
criteriaChart.legend = { position: "bottom" };
criteriaChart.yAxis = { numberFormatCode: "0.0", min: 0, max: 25 };
criteriaChart.setPosition("I10", "R29");
summarySheet.freezePanes.freezeRows(4);

const criteriaSheet = workbook.worksheets.getItem("6_채점기준");
criteriaSheet.getRange("A10:B10").merge();
criteriaSheet.getRange("A10").values = [[
  "자동평가 안내: 키워드·구조·출처 표현을 이용한 규칙 기반 1차 평가이며, 대표 문항은 반드시 수동 검토합니다.",
]];
criteriaSheet.getRange("A10:B10").format = {
  fill: AMBER,
  font: { bold: true, color: "#7F6000" },
  wrapText: true,
};
criteriaSheet.getRange("A10:B10").format.rowHeight = 42;

const caseSheet = workbook.worksheets.getOrAdd("7_대표사례분석");
resetSheet(caseSheet);
caseSheet.getRange("A1:F1").merge();
caseSheet.getRange("A1").values = [["대표 사례 분석"]];
caseSheet.getRange("A1:F1").format = {
  fill: NAVY,
  font: { bold: true, color: WHITE, size: 15 },
  horizontalAlignment: "left",
};
caseSheet.getRange("A2:F2").merge();
caseSheet.getRange("A2").values = [[
  "RAG 총점이 범용 LLM보다 높고, 근거 기반성·실무성 차이가 큰 사례를 자동 선정했습니다.",
]];
caseSheet.getRange("A2:F2").format = { fill: LIGHT, font: { color: "#5B6573" } };
writeMatrix(caseSheet, "A4", payload.representative_cases);
styleHeader(caseSheet.getRange("A4:F4"));
styleBody(caseSheet.getRange(`A5:F${payload.representative_cases.length + 3}`));
caseSheet.getRange(`A5:F${payload.representative_cases.length + 3}`).format.rowHeight = 110;
setColumnWidths(caseSheet, { A: 80, B: 330, C: 250, D: 250, E: 300, F: 340 });
caseSheet.freezePanes.freezeRows(4);
addTable(
  caseSheet,
  `A4:F${payload.representative_cases.length + 3}`,
  "RepresentativeCases",
);

const conclusionSheet = workbook.worksheets.getOrAdd("8_결론문구");
resetSheet(conclusionSheet);
conclusionSheet.getRange("A1:B1").merge();
conclusionSheet.getRange("A1").values = [["보고서·발표용 결론 문구"]];
conclusionSheet.getRange("A1:B1").format = {
  fill: NAVY,
  font: { bold: true, color: WHITE, size: 15 },
  horizontalAlignment: "left",
};
writeMatrix(conclusionSheet, "A3", payload.conclusions);
styleHeader(conclusionSheet.getRange("A3:B3"));
styleBody(conclusionSheet.getRange(`A4:B${payload.conclusions.length + 2}`));
conclusionSheet.getRange(`A4:B${payload.conclusions.length + 2}`).format.rowHeight = 95;
setColumnWidths(conclusionSheet, { A: 145, B: 950 });
conclusionSheet.freezePanes.freezeRows(3);
addTable(
  conclusionSheet,
  `A3:B${payload.conclusions.length + 2}`,
  "ComparisonConclusions",
);

const overviewSheet = workbook.worksheets.getItem("1_실험개요");
overviewSheet.getRange("A12:B13").values = [
  ["평가 방식", "답변 텍스트 기반 규칙형 1차 자동평가"],
  ["검토 원칙", "최종 평가는 대표 문항 수동 검토 필요"],
];
overviewSheet.getRange("A12:B13").format = {
  fill: LIGHT,
  wrapText: true,
  borders: { preset: "all", style: "thin", color: BORDER },
};

await fs.mkdir(new URL(".", `file:///${outputPath.replaceAll("\\", "/")}`).pathname, {
  recursive: true,
}).catch(() => {});
const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
