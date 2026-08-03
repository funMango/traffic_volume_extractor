import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "file:///C:/Users/An%20JaeYeol/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const sourcePath = "00_Data/\uc0bc\uc815\ub3d9_\ub808\ubbf8\ucf58_\uad50\ud1b5\ub7c9_\uc218\uc9d1.xlsx";
const outputName = "\ub808\ubbf8\ucf58_\ubc29\ud5a5\ubcc4_\ucca8\ub450\uc2dc_\ucd5c\ub300\uac12_\ubd84\uc11d.xlsx";
const resultPath = `02_Result/${outputName}`;
const outputDir = "outputs/remicon_direction_peak_max_20260803";
const outputPath = `${outputDir}/${outputName}`;

const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const workbook = Workbook.create();
const summary = workbook.worksheets.add("\uc694\uc57d");
const evidence = workbook.worksheets.add("\uadfc\uac70 \ub370\uc774\ud130");

const excelDateToText = (serial) => {
  const date = new Date(Date.UTC(1899, 11, 30) + Number(serial) * 86400000);
  return date.toISOString().slice(0, 10);
};

const peakDefinitions = [
  { name: "\uc624\uc804 \ucca8\ub450\uc2dc (08:00\u201310:00)", columns: [2, 3], times: ["08:00\u201309:00", "09:00\u201310:00"] },
  { name: "\uc624\ud6c4 \ucca8\ub450\uc2dc (16:00\u201318:00)", columns: [4, 5], times: ["16:00\u201317:00", "17:00\u201318:00"] },
];

const evidenceRows = [];
for (let sheetIndex = 1; sheetIndex <= 6; sheetIndex += 1) {
  const sourceSheet = sourceWorkbook.worksheets.getItemAt(sheetIndex);
  const sourceValues = sourceSheet.getRange("A3:F12").values;
  for (let directionStart = 0; directionStart < 10; directionStart += 5) {
    const intersection = sourceSheet.name;
    const direction = sourceValues[directionStart][1];
    const fiveDays = sourceValues.slice(directionStart, directionStart + 5);
    for (const peak of peakDefinitions) {
      const observations = [];
      for (const day of fiveDays) {
        for (let timeIndex = 0; timeIndex < peak.columns.length; timeIndex += 1) {
          observations.push({
            date: excelDateToText(day[0]),
            time: peak.times[timeIndex],
            value: day[peak.columns[timeIndex]],
          });
        }
      }
      evidenceRows.push({ intersection, direction, peak: peak.name, observations });
    }
  }
}

const titleFormat = { fill: "#17365D", font: { bold: true, color: "#FFFFFF", size: 16 }, horizontalAlignment: "center", verticalAlignment: "center" };
const headerFormat = { fill: "#1F4E78", font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
const subHeaderFormat = { fill: "#D9EAF7", font: { bold: true, color: "#17365D" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
const bodyBorder = { preset: "all", style: "thin", color: "#B7C9D6" };

summary.showGridLines = false;
summary.mergeCells("A1:G1");
summary.getRange("A1").values = [["\ub808\ubbf8\ucf58 \ubc29\ud5a5\ubcc4 \ucca8\ub450\uc2dc \ucd5c\ub300\uac12 \ubd84\uc11d"]];
summary.getRange("A1:G1").format = titleFormat;
summary.getRange("A1:G1").format.rowHeight = 30;
summary.mergeCells("A2:G2");
summary.getRange("A2").values = [["\ubd84\uc11d \ubc94\uc704: 2026-07-20 ~ 2026-07-24 (\uad50\ucc28\ub85c\u00b7\ubc29\ud5a5\ubcc4 \uc624\uc804/\uc624\ud6c4 10\uac1c \uad00\uce21\uac12). '-'\ub294 \uacb0\uce21\uce58\ub85c \uc81c\uc678\ud588\uc73c\uba70, \uc0bc\uc815\uace0\uac00\uad50\uc0bc\uac70\ub9ac\uc758 18:00\u201319:00\ub294 \ubd84\uc11d \ubc94\uc704\uc5d0 \ud3ec\ud568\ud558\uc9c0 \uc54a\uc74c."]];
summary.getRange("A2:G2").format = { fill: "#EAF2F8", font: { italic: true, color: "#385723" }, wrapText: true, verticalAlignment: "center" };
summary.getRange("A2:G2").format.rowHeight = 34;
summary.getRange("A4:G4").values = [["\uad50\ucc28\ub85c", "\ubc29\ud5a5", "\ucca8\ub450\uc2dc", "\ucd5c\ub300\uac12", "\ucd5c\ub300\uac12 \ubc1c\uc0dd \uc77c\uc790\u00b7\uc2dc\uac04 (\ub3d9\ub960 \uc804\uccb4)", "\ub3d9\ub960 \ubc1c\uc0dd \uac74\uc218", "\uc720\ud6a8 \uad00\uce21\uac12 \uc218"]];
summary.getRange("A4:G4").format = headerFormat;
summary.getRange("A4:G4").format.rowHeight = 32;

evidence.showGridLines = false;
evidence.mergeCells("A1:Q1");
evidence.getRange("A1").values = [["\ubc29\ud5a5\u00b7\ucca8\ub450\uc2dc\ubcc4 \uadfc\uac70 \ub370\uc774\ud130 (\uc694\uc57d \ucd5c\ub300\uac12\uc740 \ubcf8 \uc2dc\ud2b8\uc758 \uc6d0\uc2dc \uac12 \uacf5\uc2dd\uc744 \ucc38\uc870)"]];
evidence.getRange("A1:Q1").format = titleFormat;
evidence.getRange("A1:Q1").format.rowHeight = 30;
evidence.getRange("A2:G2").values = [["\uad50\ucc28\ub85c", "\ubc29\ud5a5", "\ucca8\ub450\uc2dc", "\ucd5c\ub300\uac12", "\uc720\ud6a8 \uad00\uce21\uac12 \uc218", "\ub3d9\ub960 \ubc1c\uc0dd \uac74\uc218", "\ucd5c\ub300\uac12 \ubc1c\uc0dd \uc77c\uc790\u00b7\uc2dc\uac04"]];
evidence.getRange("A2:G2").format = headerFormat;

const headerDates = [];
const headerTimes = [];
for (const observation of evidenceRows[0].observations) {
  headerDates.push(observation.date);
  headerTimes.push(observation.time.endsWith("09:00") ? "1\ucc28 \uc2dc\uac04" : "2\ucc28 \uc2dc\uac04");
}
evidence.getRange("H2:Q2").values = [headerDates];
evidence.getRange("H3:Q3").values = [headerTimes];
evidence.getRange("H2:Q3").format = subHeaderFormat;

const evidenceStartRow = 4;
for (let index = 0; index < evidenceRows.length; index += 1) {
  const rowNumber = evidenceStartRow + index;
  const entry = evidenceRows[index];
  const values = entry.observations.map((observation) => observation.value);
  evidence.getRange(`A${rowNumber}:C${rowNumber}`).values = [[entry.intersection, entry.direction, entry.peak]];
  evidence.getRange(`H${rowNumber}:Q${rowNumber}`).values = [values];
  evidence.getRange(`D${rowNumber}`).formulas = [[`=IF(COUNT(H${rowNumber}:Q${rowNumber})=0,"",MAX(H${rowNumber}:Q${rowNumber}))`]];
  evidence.getRange(`E${rowNumber}`).formulas = [[`=COUNT(H${rowNumber}:Q${rowNumber})`]];
  evidence.getRange(`F${rowNumber}`).formulas = [[`=IF(D${rowNumber}="","",COUNTIF(H${rowNumber}:Q${rowNumber},D${rowNumber}))`]];
  const occurrenceParts = [];
  for (let offset = 0; offset < 10; offset += 1) {
    const column = String.fromCharCode("H".charCodeAt(0) + offset);
    const preceding = offset === 0 ? "" : `IF(${column}${rowNumber}=D${rowNumber},IF(COUNTIF(H${rowNumber}:${String.fromCharCode("H".charCodeAt(0) + offset - 1)}${rowNumber},D${rowNumber})>0,"; ",""),"")&`;
    const label = `${entry.observations[offset].date} ${entry.observations[offset].time}`;
    occurrenceParts.push(`${preceding}IF(${column}${rowNumber}=D${rowNumber},"${label}","")`);
  }
  evidence.getRange(`G${rowNumber}`).formulas = [[`=IF(D${rowNumber}="","",${occurrenceParts.join("&")})`]];
}

const evidenceEndRow = evidenceStartRow + evidenceRows.length - 1;
evidence.getRange(`A2:Q${evidenceEndRow}`).format.borders = bodyBorder;
evidence.getRange(`D4:F${evidenceEndRow}`).format.horizontalAlignment = "center";
evidence.getRange(`H4:Q${evidenceEndRow}`).format.horizontalAlignment = "center";
evidence.getRange(`H4:Q${evidenceEndRow}`).conditionalFormats.add("containsText", { text: "-", format: { fill: "#FCE4D6", font: { color: "#C00000", italic: true } } });
evidence.freezePanes.freezeRows(3);
evidence.freezePanes.freezeColumns(3);

for (let index = 0; index < evidenceRows.length; index += 1) {
  const summaryRow = 5 + index;
  const evidenceRow = evidenceStartRow + index;
  summary.getRange(`A${summaryRow}:C${summaryRow}`).formulas = [[`='\uadfc\uac70 \ub370\uc774\ud130'!A${evidenceRow}`, `='\uadfc\uac70 \ub370\uc774\ud130'!B${evidenceRow}`, `='\uadfc\uac70 \ub370\uc774\ud130'!C${evidenceRow}`]];
  summary.getRange(`D${summaryRow}:G${summaryRow}`).formulas = [[`='\uadfc\uac70 \ub370\uc774\ud130'!D${evidenceRow}`, `='\uadfc\uac70 \ub370\uc774\ud130'!G${evidenceRow}`, `='\uadfc\uac70 \ub370\uc774\ud130'!F${evidenceRow}`, `='\uadfc\uac70 \ub370\uc774\ud130'!E${evidenceRow}`]];
}
const summaryEndRow = 4 + evidenceRows.length;
summary.getRange(`A4:G${summaryEndRow}`).format.borders = bodyBorder;
summary.getRange(`D5:D${summaryEndRow}`).format = { fill: "#FFF2CC", font: { bold: true, color: "#7F6000" }, horizontalAlignment: "center" };
summary.getRange(`F5:G${summaryEndRow}`).format.horizontalAlignment = "center";
summary.getRange(`E5:E${summaryEndRow}`).format.wrapText = true;
summary.freezePanes.freezeRows(4);
summary.freezePanes.freezeColumns(2);

for (const [sheet, widths] of [[summary, [18, 15, 29, 10, 62, 14, 14]], [evidence, [18, 15, 29, 10, 14, 14, 62, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13]]]) {
  widths.forEach((width, index) => sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width);
}
summary.getRange(`A5:G${summaryEndRow}`).format.rowHeight = 30;
evidence.getRange(`A4:Q${evidenceEndRow}`).format.rowHeight = 24;

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir("02_Result", { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
await fs.copyFile(outputPath, resultPath);

const summaryPreview = await workbook.render({ sheetName: "\uc694\uc57d", range: `A1:G${summaryEndRow}`, scale: 1.4, format: "png" });
await fs.writeFile(`${outputDir}/summary.png`, new Uint8Array(await summaryPreview.arrayBuffer()));
const evidencePreview = await workbook.render({ sheetName: "\uadfc\uac70 \ub370\uc774\ud130", range: `A1:Q${evidenceEndRow}`, scale: 0.8, format: "png" });
await fs.writeFile(`${outputDir}/evidence.png`, new Uint8Array(await evidencePreview.arrayBuffer()));

const check = await workbook.inspect({ kind: "table", range: `\uc694\uc57d!A1:G${summaryEndRow}`, include: "values,formulas", tableMaxRows: 30, tableMaxCols: 7 });
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
const summaryValues = summary.getRange(`A5:G${summaryEndRow}`).values;
for (let index = 0; index < evidenceRows.length; index += 1) {
  const entry = evidenceRows[index];
  const numericValues = entry.observations.map((observation) => observation.value).filter((value) => typeof value === "number");
  const expectedMax = Math.max(...numericValues);
  const expectedOccurrences = entry.observations
    .filter((observation) => observation.value === expectedMax)
    .map((observation) => `${observation.date} ${observation.time}`);
  const actual = summaryValues[index];
  if (actual[3] !== expectedMax || actual[4] !== expectedOccurrences.join("; ") || actual[5] !== expectedOccurrences.length || actual[6] !== numericValues.length) {
    throw new Error(`\uac80\uc99d \uc2e4\ud328: ${entry.intersection} / ${entry.direction} / ${entry.peak}`);
  }
}
if (summaryValues[0][3] !== 47) throw new Error("\ubc15\ucd0c\uad50\uc0bc\uac70\ub9ac \ub0a8(\ubd81\ud5a5) \uc624\uc804 \ucd5c\ub300\uac12 \uac80\uc99d \uc2e4\ud328");
console.log(`Validated ${evidenceRows.length} peak maxima against the source observations.`);
console.log(check.ndjson);
console.log(errors.ndjson);
