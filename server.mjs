import { createServer } from "node:http";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.dirname(fileURLToPath(import.meta.url));
const port = 4173;
const levels = ["", "Not shown", "Mentioned", "Shown: prompted", "Shown: independent", "Shown + value"];
const workbookFont = { name: "Calibri", size: 9 };
const formLabelFill = "#D9D9D9";
const formInputFill = "#FFF2CC";
const formBorder = "#666666";
const sectionHeaderFill = "#000000";
const tableHeaderFill = "#262626";
const topicBandFill = "#595959";
const tableBorder = "#B7B7B7";

function cleanFilename(name) {
  return String(name || "rubric").replace(/[^a-z0-9-_ ]/gi, "").trim().replace(/ +/g, "-") || "rubric";
}

function applyHeaderSpec(sheet, rubric) {
  const productName = String(rubric.name || "").trim();
  const title = productName.toLowerCase().endsWith("demo validation rubric")
    ? productName
    : `${productName || "[Product]"} - Demo Validation Rubric`;
  ["D4:E4", "D5:E5"].forEach((address) => sheet.getRange(address).unmerge());
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format.font = { ...workbookFont, size: 16, bold: true };
  sheet.getRange("A2").format.font = { ...workbookFont, size: 9, italic: true };

  sheet.getRange("A4").values = [["Date"]];
  sheet.getRange("C4").values = [["Presenter"]];
  sheet.getRange("A5").values = [["Product"]];
  sheet.getRange("C5").values = [["Evaluator"]];
  sheet.getRange("B5").values = [[productName || "[Product name]"]];
  sheet.getRange("B4").values = [[""]];
  sheet.getRange("D4").values = [[""]];
  sheet.getRange("D5").values = [[""]];

  ["A4", "C4", "A5", "C5"].forEach((address) => {
    const range = sheet.getRange(address);
    range.format.fill = formLabelFill;
    range.format.font = { ...workbookFont, size: 10, bold: true };
    range.format.borders = { preset: "all", style: "thin", color: formBorder };
  });
  ["B4", "D4", "B5", "D5"].forEach((address) => {
    const range = sheet.getRange(address);
    range.format.fill = formInputFill;
    range.format.font = workbookFont;
    range.format.borders = { preset: "all", style: "thin", color: formBorder };
  });
  ["E4", "E5"].forEach((address) => {
    const range = sheet.getRange(address);
    range.format.fill = "#FFFFFF";
    range.format.borders = { preset: "none" };
  });
}

function applySectionHeaderStyle(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.format.fill = sectionHeaderFill;
  range.format.font = { ...workbookFont, size: 11, bold: true, color: "#FFFFFF" };
}

function applyCellBorders(sheet, rangeAddress) {
  sheet.getRange(rangeAddress).format.borders = { preset: "all", style: "thin", color: formBorder };
}

function applyTableBorders(sheet, rangeAddress) {
  sheet.getRange(rangeAddress).format.borders = { preset: "all", style: "thin", color: tableBorder };
}

function applyScoringLegendStyle(sheet) {
  sheet.getRange("A14").values = [["Level"]];
  sheet.getRange("B14").values = [["%"]];
  sheet.getRange("C14").values = [["What it means"]];
  sheet.getRange("A14:E14").format.fill = tableHeaderFill;
  sheet.getRange("A14:E14").format.font = { ...workbookFont, bold: true, color: "#FFFFFF" };
  applyTableBorders(sheet, "A14:E20");
  sheet.getRange("A15:A19").format.font = { ...workbookFont, bold: true };
  sheet.getRange("B15:B19").format.font = { ...workbookFont, bold: true, color: "#FF0000" };
}

function applyScoringTableHeader(sheet, row) {
  const range = sheet.getRange(`A${row}:E${row}`);
  sheet.getRange(`A${row}:E${row}`).values = [["Subtopic", "Level (pick one)", "%", "", "Notes (per topic)"]];
  range.format.fill = tableHeaderFill;
  range.format.font = { ...workbookFont, bold: true, color: "#FFFFFF" };
  applyTableBorders(sheet, `A${row}:E${row}`);
}

function applyTopicBandStyle(sheet, row) {
  const range = sheet.getRange(`A${row}:E${row}`);
  range.format.fill = topicBandFill;
  range.format.font = { ...workbookFont, bold: true, color: "#FFFFFF" };
  applyTableBorders(sheet, `A${row}:E${row}`);
}

function applySubtopicRowStyle(sheet, row) {
  applyTableBorders(sheet, `A${row}:E${row}`);
  sheet.getRange(`B${row}`).format.fill = formInputFill;
}

function applyTopicScoreRowStyle(sheet, row) {
  applyTableBorders(sheet, `A${row}:E${row}`);
  sheet.getRange(`A${row}:B${row}`).format.font = { ...workbookFont, bold: true };
  sheet.getRange(`D${row}`).format.font = { ...workbookFont, bold: true };
  sheet.getRange(`D${row}`).format.fill = "#FFFFFF";
  sheet.getRange(`E${row}`).format.fill = formInputFill;
}

function applySummaryBlockStyle(sheet, summary) {
  applyTableBorders(sheet, `A${summary}:E${summary + 25}`);
  [
    `A${summary + 3}:B${summary + 4}`,
    `A${summary + 10}:B${summary + 10}`,
    `A${summary + 14}:B${summary + 17}`,
    `A${summary + 20}:B${summary + 22}`,
    `A${summary + 25}:B${summary + 25}`,
  ].forEach((rangeAddress) => {
    const range = sheet.getRange(rangeAddress);
    range.format.fill = formLabelFill;
    range.format.font = { ...workbookFont, bold: true };
  });
  sheet.getRange(`C${summary + 3}:E${summary + 3}`).format.fill = "#F4CCCC";
  [
    `A${summary + 7}:E${summary + 7}`,
    `C${summary + 15}:E${summary + 15}`,
    `C${summary + 17}:E${summary + 17}`,
    `C${summary + 20}:E${summary + 22}`,
  ].forEach((rangeAddress) => {
    sheet.getRange(rangeAddress).format.fill = formInputFill;
  });
  [
    `C${summary + 4}:E${summary + 4}`,
    `A${summary + 11}:E${summary + 11}`,
    `C${summary + 14}:E${summary + 14}`,
    `C${summary + 16}:E${summary + 16}`,
    `C${summary + 25}:E${summary + 25}`,
  ].forEach((rangeAddress) => {
    sheet.getRange(rangeAddress).format.fill = "#FFFFFF";
  });
}

async function createWorkbook(rubric) {
  const totalWeight = (rubric.topics || []).reduce((total, topic) => total + (Number(topic.weight) || 0), 0);
  if (totalWeight > 100) throw new Error("Topic weights cannot exceed 100%.");
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Rubric");
  sheet.showGridLines = false;
  sheet.getRange("A1:G1").merge();
  sheet.getRange("A1").values = [[rubric.name || "Rubric"]];
  sheet.getRange("A1:G1").format = { fill: "#17312C", font: { bold: true, color: "#FFFFFF", size: 18 } };
  sheet.getRange("A3:G3").values = [["Topic / evidence", "Level", "%", "Weight", "Weighted score", "Value counted", "Value tied?"]];
  sheet.getRange("A3:G3").format = { fill: "#2F6E58", font: { bold: true, color: "#FFFFFF" } };

  let row = 4;
  const subtopicRanges = [];
  const scoreRows = [];
  for (const [index, topic] of (rubric.topics || []).entries()) {
    const subtopics = topic.subtopics || [];
    const start = row + 1;
    sheet.getRange(`A${row}:G${row}`).merge();
    sheet.getRange(`A${row}`).values = [[`${index + 1}. ${topic.name || "Untitled topic"}`]];
    sheet.getRange(`A${row}:G${row}`).format = { fill: "#EAF0E5", font: { bold: true, color: "#17312C" } };
    row += 1;
    for (const item of subtopics) {
      const subtopic = typeof item === "string" ? { name: item, level: "" } : item;
      sheet.getRange(`A${row}:G${row}`).values = [[subtopic.name || "Evidence item", "", null, null, null, null, null]];
      sheet.getRange(`C${row}`).formulas = [[`=IF($B${row}="Shown + value",100,IF($B${row}="Shown: independent",75,IF($B${row}="Shown: prompted",50,IF($B${row}="Mentioned",25,IF($B${row}="Not shown",0,"")))))`]];
      sheet.getRange(`B${row}`).dataValidation = { rule: { type: "list", values: levels } };
      row += 1;
    }
    const end = row - 1;
    if (subtopics.length) subtopicRanges.push(`B${start}:B${end}`);
    sheet.getRange(`A${row}:G${row}`).values = [["Topic score", null, null, Number(topic.weight) || 0, null, "Yes", null]];
    sheet.getRange(`E${row}`).formulas = [[`=IFERROR(AVERAGE(C${start}:C${end})*D${row}/100,0)`]];
    sheet.getRange(`G${row}`).formulas = [[`=IF(F${row}="Yes",IF(COUNTIF(C${start}:C${end},100)>0,1,0),"")`]];
    sheet.getRange(`A${row}:G${row}`).format = { fill: "#F6F2E9", font: { bold: true } };
    scoreRows.push(row);
    row += 2;
  }

  const scoreRefs = scoreRows.map((scoreRow) => `E${scoreRow}`).join(",") || "0";
  const valueRefs = scoreRows.map((scoreRow) => `G${scoreRow}`).join(",") || "0";
  const unscoredFormula = subtopicRanges.map((range) => `COUNTBLANK(${range})`).join("+") || "0";
  const summary = row;
  sheet.getRange(`A${summary}:G${summary}`).merge();
  sheet.getRange(`A${summary}`).values = [["SCORING SUMMARY"]];
  sheet.getRange(`A${summary}:G${summary}`).format = { fill: "#17312C", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange(`A${summary + 1}:B${summary + 4}`).values = [
    ["Total score", null], ["Unscored subtopics", null], ["Topics with value tied", null], ["Value gate required", Number(rubric.valueGate) || 0],
  ];
  sheet.getRange(`B${summary + 1}`).formulas = [[`=SUM(${scoreRefs})`]];
  sheet.getRange(`B${summary + 2}`).formulas = [[`=${unscoredFormula}`]];
  sheet.getRange(`B${summary + 3}`).formulas = [[`=SUM(${valueRefs})`]];
  sheet.getRange(`A${summary + 6}:B${summary + 6}`).merge();
  sheet.getRange(`A${summary + 6}`).formulas = [[`=IF(B${summary + 2}>0,"Complete every subtopic before determining eligibility",IF(B${summary + 3}>=B${summary + 4},"Value gate MET - eligible for nomination","Value gate NOT met - tie more topics to customer value"))`]];
  sheet.getRange(`A${summary + 6}:B${summary + 6}`).format = { fill: "#DCE5D5", font: { bold: true, color: "#17312C" }, wrapText: true };
  sheet.getRange(`A3:G${summary + 6}`).format.borders = { preset: "insideHorizontal", style: "thin", color: "#D7DDD4" };
  sheet.getRange(`A1:A${summary + 6}`).format.columnWidth = 47;
  sheet.getRange(`B1:B${summary + 6}`).format.columnWidth = 23;
  sheet.getRange(`C1:G${summary + 6}`).format.columnWidth = 16;
  sheet.getRange(`C4:C${summary}`).format.numberFormat = "0";
  sheet.getRange(`D4:E${summary}`).format.numberFormat = "0.0";
  sheet.freezePanes.freezeRows(3);
  return workbook;
}

async function applyExportPackageFixes(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const stylesFile = zip.file("xl/styles.xml");
  if (!stylesFile) return bytes;
  const styles = await stylesFile.async("string");
  const updatedStyles = styles.replace(
    /<x:font><x:sz val="9" \/><x:name val="Calibri" \/><\/x:font>/,
    '<x:font><x:sz val="9.5" /><x:name val="Calibri" /></x:font>',
  );
  if (updatedStyles === styles) return bytes;
  zip.file("xl/styles.xml", updatedStyles);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

async function createTemplateWorkbook(rubric) {
  const totalWeight = (rubric.topics || []).reduce((total, topic) => total + (Number(topic.weight) || 0), 0);
  if (totalWeight > 100) throw new Error("Topic weights cannot exceed 100%.");

  const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root, "src", "Demo Rubric Template.xlsx")));
  const source = sourceWorkbook.worksheets.getItem("Rubric");
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Rubric");
  sheet.showGridLines = false;
  [27.75, 15, 9, 15, 15, 9, 19.5, 25.5, 25.5, 25.5, 25.5, 9, 19.5, 15, 25.5, 25.5, 25.5, 25.5, 25.5, 25.5, 9, 19.5].forEach((height, index) => {
    sheet.getRange(`A${index + 1}:E${index + 1}`).format.rowHeight = height;
  });
  sheet.getRange("A1:E22").copyFrom(source.getRange("A1:E22"), "all");
  ["A1:E1", "A2:E2", "A7:E7", "A8:E8", "A9:E9", "A10:E10", "A11:E11", "A13:E13", "C14:E14", "C15:E15", "C16:E16", "C17:E17", "C18:E18", "C19:E19", "A20:E20", "A22:E22"].forEach((range) => sheet.getRange(range).merge());

  let row = 24;
  const scoreRows = [];
  const unscoredRanges = [];
  const valueCountFormulas = [];
  const topicBandRows = [];
  const subtopicRows = [];
  for (const [index, topic] of (rubric.topics || []).entries()) {
    const subtopics = topic.subtopics || [];
    const weight = Number(topic.weight) || 0;
    sheet.getRange(`A${row}:E${row}`).copyFrom(source.getRange("A24:E24"), "all");
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}:E${row}`).format.rowHeight = 19.5;
    sheet.getRange(`A${row}`).values = [[`${index + 1} - ${topic.name || "Untitled topic"} (weight ${weight})`]];
    topicBandRows.push(row);
    row += 1;
    const start = row;
    for (const item of subtopics) {
      const subtopic = typeof item === "string" ? { name: item } : item;
      sheet.getRange(`A${row}:E${row}`).copyFrom(source.getRange("A25:E25"), "all");
      sheet.getRange(`A${row}:E${row}`).format.rowHeight = 16.5;
      sheet.getRange(`A${row}`).values = [[subtopic.name || "Evidence item"]];
      sheet.getRange(`B${row}`).values = [[""]];
      sheet.getRange(`C${row}`).formulas = [[`=IF($B${row}="Shown + value",100,IF($B${row}="Shown: independent",75,IF($B${row}="Shown: prompted",50,IF($B${row}="Mentioned",25,IF($B${row}="Not shown",0,"")))))`]];
      sheet.getRange(`B${row}`).dataValidation = { rule: { type: "list", values: levels } };
      subtopicRows.push(row);
      row += 1;
    }
    const end = row - 1;
    if (subtopics.length) unscoredRanges.push(`B${start}:B${end}`);
    sheet.getRange(`A${row}:E${row}`).copyFrom(source.getRange("A29:E29"), "all");
    sheet.getRange(`A${row}:E${row}`).format.rowHeight = 18;
    sheet.getRange(`A${row}`).values = [[`▸ Topic ${index + 1} score (average x weight)`]];
    sheet.getRange(`B${row}`).formulas = [[`=IFERROR(AVERAGE(C${start}:C${end})*${weight}/100,0)`]];
    sheet.getRange(`D${row}`).values = [["Notes:"]];
    scoreRows.push(row);
    valueCountFormulas.push(`COUNTIF(C${start}:C${end},100)`);
    row += 1;
  }

  const summary = row + 1;
  sheet.getRange(`A1:A${summary + 25}`).format.columnWidth = 44;
  sheet.getRange(`B1:B${summary + 25}`).format.columnWidth = 17.1796875;
  sheet.getRange(`C1:C${summary + 25}`).format.columnWidth = 16.08984375;
  sheet.getRange(`D1:D${summary + 25}`).format.columnWidth = 18;
  sheet.getRange(`E1:E${summary + 25}`).format.columnWidth = 44;
  sheet.getRange(`A${summary}:E${summary + 25}`).copyFrom(source.getRange("A61:E86"), "all");
  [21.75, 9, 19.5, 15, 15, 9, 19.5, 24, 9, 19.5, 15, 21.75, 9, 19.5, 15, 15, 15, 15, 9, 19.5, 39.75, 39.75, 39.75, 9, 19.5, 15].forEach((height, index) => {
    sheet.getRange(`A${summary + index}:E${summary + index}`).format.rowHeight = height;
  });
  [2, 6, 7, 9, 11, 13, 19, 24].forEach((offset) => sheet.getRange(`A${summary + offset}:E${summary + offset}`).merge());
  [3, 4, 10, 14, 15, 16, 20, 21, 22, 25].forEach((offset) => {
    sheet.getRange(`A${summary + offset}:B${summary + offset}`).merge();
    sheet.getRange(`C${summary + offset}:E${summary + offset}`).merge();
  });
  const scoreRefs = scoreRows.map((scoreRow) => `B${scoreRow}`).join(",") || "0";
  const unscoredFormula = unscoredRanges.map((range) => `COUNTBLANK(${range})`).join("+") || "0";
  sheet.getRange(`C${summary}`).formulas = [[`=SUM(${scoreRefs})`]];
  sheet.getRange(`C${summary + 3}`).formulas = [[`=${unscoredFormula}`]];
  sheet.getRange(`C${summary + 4}`).formulas = [[`=IF(C${summary + 3}>0,"—",B${scoreRows[Math.min(2, scoreRows.length - 1)] || 1})`]];
  sheet.getRange(`C${summary + 10}`).formulas = [[`=${valueCountFormulas.join("+") || "0"}`]];
  sheet.getRange(`A${summary + 11}`).formulas = [[`=IF(C${summary + 3}>0,"—",IF(C${summary + 10}>=C${summary + 25},"Value gate MET - strong pass; eligible for nomination","Value gate NOT met - value-tying is the development area; not yet endorsed"))`]];
  sheet.getRange(`C${summary + 14}`).formulas = [[`=IF(A${summary + 7}="","—",IF(A${summary + 7}="Competency met","Yes","No"))`]];
  sheet.getRange(`C${summary + 16}`).formulas = [[`=IF(C${summary + 3}>0,"—",IF(C${summary + 10}>=C${summary + 25},"Eligible","Not yet"))`]];
  sheet.getRange(`C${summary + 25}`).values = [[Number(rubric.valueGate) || 0]];
  sheet.getRange(`A${summary + 7}`).dataValidation = { rule: { type: "list", values: ["", "Competency met", "Targeted re-validation - name the topic in notes", "Full re-validation (after coaching)"] } };
  sheet.getRange(`C${summary + 15}`).dataValidation = { rule: { type: "list", values: ["", "Yes", "No", "N/A - none occurred"] } };
  sheet.getRange(`C${summary + 17}`).dataValidation = { rule: { type: "list", values: ["", "-", "Validator", "Mentor"] } };
  sheet.getRange(`A1:E${summary + 25}`).format.font = workbookFont;
  applyHeaderSpec(sheet, rubric);
  [
    "A7:E7",
    "A13:E13",
    "A22:E22",
    `A${summary}:E${summary}`,
    `A${summary + 2}:E${summary + 2}`,
    `A${summary + 6}:E${summary + 6}`,
    `A${summary + 9}:E${summary + 9}`,
    `A${summary + 13}:E${summary + 13}`,
    `A${summary + 19}:E${summary + 19}`,
    `A${summary + 24}:E${summary + 24}`,
  ].forEach((range) => applySectionHeaderStyle(sheet, range));
  applyCellBorders(sheet, "A8:E11");
  applyScoringLegendStyle(sheet);
  applyScoringTableHeader(sheet, 23);
  topicBandRows.forEach((topicRow) => applyTopicBandStyle(sheet, topicRow));
  subtopicRows.forEach((subtopicRow) => applySubtopicRowStyle(sheet, subtopicRow));
  scoreRows.forEach((scoreRow) => applyTopicScoreRowStyle(sheet, scoreRow));
  applySummaryBlockStyle(sheet, summary);

  const adaptation = workbook.worksheets.add("How to adapt");
  adaptation.getRange("A1:A26").format.columnWidth = 4;
  adaptation.getRange("B1:B26").format.columnWidth = 112;
  adaptation.getRange("A1:B26").copyFrom(sourceWorkbook.worksheets.getItem("How to adapt").getRange("A1:B26"), "all");
  adaptation.getRange("A1:B26").format.font = workbookFont;
  return workbook;
}

async function createExactTemplateWorkbook(rubric) {
  const totalWeight = (rubric.topics || []).reduce((total, topic) => total + (Number(topic.weight) || 0), 0);
  if (totalWeight > 100) throw new Error("Topic weights cannot exceed 100%.");

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root, "src", "Demo Rubric Template.xlsx")));
  const sheet = workbook.worksheets.getItem("Rubric");
  const slots = [
    { band: 24, subtopics: [25, 26, 27, 28], score: 29 },
    { band: 30, subtopics: [31, 32, 33], score: 34 },
    { band: 35, subtopics: [36, 37, 38, 39], score: 40 },
    { band: 41, subtopics: [42, 43, 44], score: 45 },
    { band: 46, subtopics: [47, 48, 49], score: 50 },
    { band: 51, subtopics: [52, 53, 54], score: 55 },
    { band: 56, subtopics: [57, 58, 59], score: 60 },
  ];
  const topics = rubric.topics || [];
  if (topics.length > slots.length) throw new Error("The source template supports up to 7 topics.");

  applyHeaderSpec(sheet, rubric);
  sheet.getRange("C86").values = [[Number(rubric.valueGate) || 0]];
  [
    "A7:E7",
    "A13:E13",
    "A22:E22",
    "A61:E61",
    "A63:E63",
    "A67:E67",
    "A70:E70",
    "A74:E74",
    "A80:E80",
    "A85:E85",
  ].forEach((range) => applySectionHeaderStyle(sheet, range));
  applyCellBorders(sheet, "A8:E11");
  applySummaryBlockStyle(sheet, 61);

  slots.forEach((slot, index) => {
    const topic = topics[index];
    if (topic && topic.subtopics.length > slot.subtopics.length) {
      throw new Error(`Topic ${index + 1} exceeds this template slot's ${slot.subtopics.length}-subtopic limit.`);
    }
    if (!topic) {
      sheet.getRange(`A${slot.band}`).values = [[""]];
      sheet.getRange(`A${slot.score}`).values = [[""]];
      sheet.getRange(`B${slot.score}`).formulas = [["=0"]];
      sheet.getRange(`N${slot.score}`).values = [["No"]];
    } else {
      const weight = Number(topic.weight) || 0;
      sheet.getRange(`A${slot.band}`).values = [[`${index + 1} - ${topic.name || "Untitled topic"} (weight ${weight})`]];
      sheet.getRange(`A${slot.score}`).values = [[`▸ Topic ${index + 1} score (average x weight)`]];
      sheet.getRange(`B${slot.score}`).formulas = [[`=H${slot.score}*${weight}/100`]];
      sheet.getRange(`N${slot.score}`).values = [["Yes"]];
    }
    slot.subtopics.forEach((row, subtopicIndex) => {
      const subtopic = topic?.subtopics[subtopicIndex];
      if (subtopic) {
        sheet.getRange(`A${row}`).values = [[typeof subtopic === "string" ? subtopic : subtopic.name || "Evidence item"]];
        sheet.getRange(`B${row}`).values = [[""]];
      } else {
        sheet.getRange(`A${row}:C${row}`).clear({ applyTo: "contents" });
        sheet.getRange(`F${row}`).clear({ applyTo: "contents" });
      }
    });
  });
  return workbook;
}

function serveFile(request, response) {
  const requested = request.url === "/" ? "index.html" : request.url.slice(1);
  const file = path.resolve(root, requested);
  if (!file.startsWith(root)) { response.writeHead(403); response.end(); return; }
  const types = { ".html": "text/html", ".css": "text/css", ".js": "application/javascript" };
  readFile(file).then((content) => {
    response.writeHead(200, { "Content-Type": `${types[path.extname(file)] || "text/plain"}; charset=utf-8` });
    response.end(content);
  }).catch(() => { response.writeHead(404); response.end("Not found"); });
}

createServer(async (request, response) => {
  response.setHeader("Access-Control-Allow-Origin", "*");
  if (request.method === "OPTIONS") { response.writeHead(204); response.end(); return; }
  if (request.method !== "POST" || request.url !== "/api/export") { serveFile(request, response); return; }
  let body = "";
  for await (const chunk of request) body += chunk;
  try {
    const rubric = request.headers["content-type"]?.includes("application/x-www-form-urlencoded")
      ? JSON.parse(new URLSearchParams(body).get("rubric") || "{}")
      : JSON.parse(body);
    const workbook = await createTemplateWorkbook(rubric);
    const output = await SpreadsheetFile.exportXlsx(workbook);
    const exportDir = await mkdtemp(path.join(tmpdir(), "rubric-builder-"));
    const exportPath = path.join(exportDir, "rubric.xlsx");
    await output.save(exportPath);
    const bytes = await applyExportPackageFixes(await readFile(exportPath));
    await rm(exportDir, { recursive: true, force: true });
    response.writeHead(200, {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${cleanFilename(rubric.name)}.xlsx"`,
    });
    response.end(bytes);
  } catch (error) {
    const status = error.message === "Topic weights cannot exceed 100%." ? 400 : 500;
    response.writeHead(status, { "Content-Type": "application/json" });
    response.end(JSON.stringify({ error: error.message }));
  }
}).listen(port, "127.0.0.1", () => console.log(`Rubric Builder is running at http://localhost:${port}`));
