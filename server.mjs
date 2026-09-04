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
const trackConfigs = {
  demo: {
    title: "Demo Validation Rubric",
    subtitle: "Competency Validation Program - Demo Track - TEMPLATE",
    levels: ["", "Not shown", "Mentioned", "Shown: prompted", "Shown: independent", "Shown + value"],
    percentMap: [
      ["Not shown", 0],
      ["Mentioned", 25],
      ["Shown: prompted", 50],
      ["Shown: independent", 75],
      ["Shown + value", 100],
    ],
    instructions: [
      "•  During the demo, do one thing per subtopic: pick one level from the dropdown.",
      "•  Blank = not yet scored (it won't count). To score a zero, pick \"Not shown\". The Unscored counter below flags anything missed.",
      "•  The top level is reserved for customer value.  Clean, unprompted execution reaches \"Shown: independent\" (75%); tying it to a customer outcome earns \"Shown + value\" (100%).",
      "•  If this product has a flagship topic, treat its score as a serious signal - not a switch - when you make your verdict.",
    ],
    levelDescriptions: [
      ["Not shown", 0, "Neither demonstrated nor mentioned. Absent from the demo."],
      ["Mentioned", 25, "Referred to verbally but not demonstrated when it was demonstrable."],
      ["Shown: prompted", 50, "Demonstrated only after being asked or reminded, or with help locating or executing it."],
      ["Shown: independent", 75, "Demonstrated competently, unprompted, start to finish - but customer value was not tied in."],
      ["Shown + value", 100, "Everything in \"independent\", plus an explicit tie-in to a customer outcome, business case, or use case."],
    ],
    finalTopicNote: "Final topic (Narration quality & demo hygiene) only: read the five levels as quality bands--Absent (0) / Poor (25) / Adequate (50) / Strong (75) / Exemplary (100). It does not count toward the value gate.",
    layer2Header: "VALUE TIED-IN (automatic; validator/mentor eligibility)",
    valueCountLabel: "Topics with value tied",
    gateMet: "Value gate MET - strong pass; eligible for nomination",
    gateNotMet: "Value gate NOT met - value-tying is the development area; not yet endorsed",
    settingsLabel: "Value-gate minimum (number of topics)",
    feedbackLiked: "Things I liked about the demo",
    excludeFinalTopicFromGate: false,
  },
  pitch: {
    title: "Pitch Validation Rubric",
    subtitle: "Competency Validation Program - Pitch Track - TEMPLATE",
    levels: ["", "Not addressed", "Mentioned", "Pitched: prompted", "Pitched: fluent", "Pitched + tailored"],
    percentMap: [
      ["Not addressed", 0],
      ["Mentioned", 25],
      ["Pitched: prompted", 50],
      ["Pitched: fluent", 75],
      ["Pitched + tailored", 100],
    ],
    instructions: [
      "•  During the pitch, do one thing per subtopic: pick one level from the dropdown.",
      "•  Blank = not yet scored (it won't count). To score a zero, pick \"Not addressed\". The Unscored counter below flags anything missed.",
      "•  The top level is reserved for tailoring.  Fluent, accurate, unprompted coverage reaches \"Pitched: fluent\" (75%). The final 25% is earned only by connecting the point to what THIS customer said - their discovery answers, environment, or stated pain.",
      "•  If this product has a flagship topic, treat its score as a serious signal - not a switch - when you make your verdict.",
    ],
    levelDescriptions: [
      ["Not addressed", 0, "Absent from the pitch entirely."],
      ["Mentioned", 25, "Named in passing without substance - when it warranted real treatment."],
      ["Pitched: prompted", 50, "Covered with substance, but only after the evaluator asked or had to steer the conversation to it."],
      ["Pitched: fluent", 75, "Covered accurately, confidently, unprompted - but generically; not connected to this customer."],
      ["Pitched + tailored", 100, "Everything in \"fluent\", plus explicitly tied to the audience - their discovery answers, environment, or stated pain."],
    ],
    finalTopicNote: "Final topic (Delivery & pacing) only: read the five levels as quality bands - Absent (0) / Poor (25) / Adequate (50) / Strong (75) / Exemplary (100). It does not count toward the tailoring gate.",
    layer2Header: "LAYER 2 - TAILORING (automatic)",
    valueCountLabel: "Topics tailored",
    gateMet: "Tailoring gate MET - strong pass; eligible for nomination",
    gateNotMet: "Tailoring gate NOT met - tailor more points to this customer",
    settingsLabel: "Tailoring-gate minimum (number of topics)",
    feedbackLiked: "Things I liked about the pitch",
    excludeFinalTopicFromGate: true,
  },
  poc: {
    title: "POC Validation Rubric",
    subtitle: "Competency Validation Program - POC Track - TEMPLATE",
    levels: ["", "Not attempted", "Explained only", "Done: prompted", "Done: independent", "Done + validated"],
    percentMap: [
      ["Not attempted", 0],
      ["Explained only", 25],
      ["Done: prompted", 50],
      ["Done: independent", 75],
      ["Done + validated", 100],
    ],
    instructions: [
      "•  This is a do-then-prove validation. Most subtopics require the SE to have built something and to show it working - not to describe it.",
      "•  Run it in two passes: inspect the prepared environment first, then ask the SE to induce the changes the scenario requires.",
      "•  Blank = not yet scored (it won't count). To score a zero, pick \"Not attempted\". The Unscored counter below flags anything missed.",
      "•  The top level is reserved for proof.  A working build reaches \"Done: independent\" (75%); proving it worked against the stated requirement earns \"Done + validated\" (100%).",
      "•  Topic 3 is the gated topic.  Its score is a signal, not a switch: treat a weak showing there as a serious flag when you make your Layer 1 call.",
    ],
    levelDescriptions: [
      ["Not attempted", 0, "Neither built nor addressed. Absent from the session."],
      ["Explained only", 25, "Described or talked through, but not built or shown working - when it was buildable."],
      ["Done: prompted", 50, "Built and working, but only after being asked, reminded, or helped through it."],
      ["Done: independent", 75, "Built and working, unprompted, start to finish - but not proven against the requirement."],
      ["Done + validated", 100, "Everything in \"independent\", plus evidence that it met the requirement - the log, the tag state, the denied session, the result the customer asked for."],
    ],
    finalTopicNote: "Topic 6 is walkthrough-accepted: the SE is not expected to build these in the session. Read the levels as - Not attempted (0) / Explained only (25), recited generically / Done: prompted (50), walked the configuration only after being asked / Done: independent (75), walked it unprompted and accurately / Done + validated (100), plus applied it to the situation in front of them. Every other topic expects a build.",
    layer2Header: "LAYER 2 - VALIDATION (automatic)",
    valueCountLabel: "Topics with proof validated",
    gateMet: "Validation gate MET - strong pass; eligible for nomination",
    gateNotMet: "Validation gate NOT met - proof is the development area; not yet endorsed",
    settingsLabel: "Validation-gate minimum (number of topics)",
    feedbackLiked: "Things I liked about the POC",
    excludeFinalTopicFromGate: false,
  },
};

function cleanFilename(name) {
  return String(name || "rubric").replace(/[^a-z0-9-_ ]/gi, "").trim().replace(/ +/g, "-") || "rubric";
}

function productNameForExport(rubric, config) {
  const productName = String(rubric.name || "").trim();
  return productName.toLowerCase() === config.title.toLowerCase() ? "" : productName;
}

function applyHeaderSpec(sheet, rubric, config) {
  const productName = productNameForExport(rubric, config);
  const title = productName.toLowerCase().endsWith(config.title.toLowerCase())
    ? productName
    : `${productName || "[Product]"} - ${config.title}`;
  ["D4:E4", "D5:E5"].forEach((address) => sheet.getRange(address).unmerge());
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A2").values = [[config.subtitle]];
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
  sheet.getRange("A7").values = [["HOW TO USE"]];
  ["A8", "A9", "A10", "A11"].forEach((address, index) => {
    sheet.getRange(address).values = [[config.instructions[index] || ""]];
  });
  sheet.getRange("A13").values = [["SCORING LEVELS EXPLAINED - the same five apply to every subtopic"]];
  sheet.getRange("A22").values = [["SCORING"]];

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

function applyScoringLegendStyle(sheet, config) {
  sheet.getRange("A14").values = [["Level"]];
  sheet.getRange("B14").values = [["%"]];
  sheet.getRange("C14").values = [["What it means"]];
  sheet.getRange("A15:C19").values = config.levelDescriptions;
  sheet.getRange("A20").values = [[config.finalTopicNote]];
  sheet.getRange("A14:E14").format.fill = tableHeaderFill;
  sheet.getRange("A14:E14").format.font = { ...workbookFont, bold: true, color: "#FFFFFF" };
  applyTableBorders(sheet, "A14:E20");
  sheet.getRange("A15:A19").format.font = { ...workbookFont, bold: true };
  sheet.getRange("B15:B19").format.font = { ...workbookFont, bold: true, color: "#FF0000" };
}

function applyScoringLegendStyleAt(sheet, config, headerRow) {
  const tableHeaderRow = headerRow + 1;
  const firstLevelRow = headerRow + 2;
  const noteRow = headerRow + 7;
  sheet.getRange(`A${tableHeaderRow}`).values = [["Level"]];
  sheet.getRange(`B${tableHeaderRow}`).values = [["%"]];
  sheet.getRange(`C${tableHeaderRow}`).values = [["What it means"]];
  sheet.getRange(`A${firstLevelRow}:C${firstLevelRow + 4}`).values = config.levelDescriptions;
  sheet.getRange(`A${noteRow}`).values = [[config.finalTopicNote]];
  sheet.getRange(`A${tableHeaderRow}:E${tableHeaderRow}`).format.fill = tableHeaderFill;
  sheet.getRange(`A${tableHeaderRow}:E${tableHeaderRow}`).format.font = { ...workbookFont, bold: true, color: "#FFFFFF" };
  applyTableBorders(sheet, `A${tableHeaderRow}:E${noteRow}`);
  sheet.getRange(`A${firstLevelRow}:A${firstLevelRow + 4}`).format.font = { ...workbookFont, bold: true };
  sheet.getRange(`B${firstLevelRow}:B${firstLevelRow + 4}`).format.font = { ...workbookFont, bold: true, color: "#FF0000" };
  sheet.getRange(`A${noteRow}:E${noteRow}`).format.font = { ...workbookFont, italic: true, color: "#666666" };
  sheet.getRange(`A${noteRow}:E${noteRow}`).format.wrapText = true;
}

function levelFormula(row, config) {
  const [, first, second, third, fourth, fifth] = config.levels;
  return `=IF($B${row}="${fifth}",100,IF($B${row}="${fourth}",75,IF($B${row}="${third}",50,IF($B${row}="${second}",25,IF($B${row}="${first}",0,"")))))`;
}

function subtopicRowHeight(text) {
  const length = String(text || "").length;
  if (length > 150) return 49.5;
  if (length > 105) return 39.75;
  if (length > 70) return 30;
  if (length > 48) return 24;
  return 16.5;
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
  sheet.getRange(`A${row}`).format.wrapText = true;
  sheet.getRange(`A${row}`).format.verticalAlignment = "top";
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

function applyPitchSummaryBlockStyle(sheet, summary) {
  applyTableBorders(sheet, `A${summary}:E${summary + 34}`);
  [
    `A${summary + 3}:B${summary + 4}`,
    `A${summary + 7}:B${summary + 7}`,
    `A${summary + 9}:B${summary + 9}`,
    `A${summary + 16}:B${summary + 16}`,
    `A${summary + 20}:B${summary + 24}`,
    `A${summary + 27}:B${summary + 29}`,
    `A${summary + 33}:B${summary + 33}`,
  ].forEach((rangeAddress) => {
    const range = sheet.getRange(rangeAddress);
    range.format.fill = formLabelFill;
    range.format.font = { ...workbookFont, bold: true };
  });
  sheet.getRange(`A${summary + 4}:B${summary + 4}`).format.font = { ...workbookFont, bold: true, color: "#FF0000" };
  sheet.getRange(`C${summary + 3}:E${summary + 3}`).format.fill = "#F4CCCC";
  [
    `C${summary + 7}:E${summary + 7}`,
    `C${summary + 9}:E${summary + 9}`,
    `A${summary + 13}:E${summary + 13}`,
    `C${summary + 21}:E${summary + 22}`,
    `C${summary + 24}:E${summary + 24}`,
    `C${summary + 27}:E${summary + 29}`,
  ].forEach((rangeAddress) => {
    sheet.getRange(rangeAddress).format.fill = formInputFill;
  });
  [
    `C${summary + 4}:E${summary + 4}`,
    `A${summary + 8}:E${summary + 8}`,
    `A${summary + 10}:E${summary + 10}`,
    `A${summary + 17}:E${summary + 17}`,
    `C${summary + 20}:E${summary + 20}`,
    `C${summary + 23}:E${summary + 23}`,
    `C${summary + 33}:E${summary + 33}`,
  ].forEach((rangeAddress) => {
    sheet.getRange(rangeAddress).format.fill = "#FFFFFF";
  });
  [`A${summary + 8}:E${summary + 8}`, `A${summary + 10}:E${summary + 10}`].forEach((rangeAddress) => {
    sheet.getRange(rangeAddress).format.font = { ...workbookFont, italic: true, color: "#666666" };
  });
}

function writeDemoSummaryContent(sheet, summary, context) {
  const { config, scoreRows, unscoredFormula, valueCountFormula, valueGate } = context;
  sheet.getRange(`A${summary}`).values = [["TOTAL SCORE"]];
  sheet.getRange(`A${summary + 2}`).values = [["SCORE SUMMARY - reference signals for your verdict"]];
  sheet.getRange(`A${summary + 3}`).values = [["Unscored subtopics remaining"]];
  sheet.getRange(`A${summary + 4}`).values = [["[Flagship topic] signal (optional - repoint or delete)"]];
  sheet.getRange(`A${summary + 6}`).values = [["VERDICT (pick your verdict)"]];
  sheet.getRange(`A${summary + 9}`).values = [[config.layer2Header]];
  sheet.getRange(`A${summary + 10}`).values = [[config.valueCountLabel]];
  sheet.getRange(`A${summary + 13}`).values = [["VALIDATOR DECISIONS"]];
  sheet.getRange(`A${summary + 14}`).values = [["Passing grade"]];
  sheet.getRange(`A${summary + 15}`).values = [["Did the presenter handle a curveball globally?"]];
  sheet.getRange(`A${summary + 16}`).values = [["Is the presenter eligible to grade others on this product? (automatic)"]];
  sheet.getRange(`A${summary + 17}`).values = [["Role granted (only if the presenter volunteered)"]];
  sheet.getRange(`A${summary + 19}`).values = [["QUALITATIVE FEEDBACK"]];
  sheet.getRange(`A${summary + 20}`).values = [[config.feedbackLiked]];
  sheet.getRange(`A${summary + 21}`).values = [["Things that need more attention"]];
  sheet.getRange(`A${summary + 22}`).values = [["Additional feedback"]];
  sheet.getRange(`A${summary + 24}`).values = [["TEMPLATE SETTINGS - for adapting; leave alone while grading"]];
  sheet.getRange(`A${summary + 25}`).values = [[config.settingsLabel]];
  sheet.getRange(`C${summary}`).formulas = [[`=SUM(${scoreRows.map((scoreRow) => `B${scoreRow}`).join(",") || "0"})`]];
  sheet.getRange(`C${summary}`).format.numberFormat = '0.0" / 100"';
  sheet.getRange(`C${summary + 3}`).formulas = [[`=${unscoredFormula}`]];
  sheet.getRange(`C${summary + 4}`).formulas = [[`=IF(C${summary + 3}>0,"—",B${scoreRows[Math.min(2, scoreRows.length - 1)] || 1})`]];
  sheet.getRange(`C${summary + 10}`).formulas = [[`=${valueCountFormula}`]];
  sheet.getRange(`A${summary + 11}`).formulas = [[`=IF(C${summary + 3}>0,"—",IF(C${summary + 10}>=C${summary + 25},"${config.gateMet}","${config.gateNotMet}"))`]];
  sheet.getRange(`C${summary + 14}`).formulas = [[`=IF(A${summary + 7}="","—",IF(A${summary + 7}="Competency met","Yes","No"))`]];
  sheet.getRange(`C${summary + 16}`).formulas = [[`=IF(C${summary + 3}>0,"—",IF(C${summary + 10}>=C${summary + 25},"Eligible","Not yet"))`]];
  sheet.getRange(`C${summary + 25}`).values = [[valueGate]];
}

function writePitchSummaryContent(sheet, summary, context) {
  const { config, scoreRows, unscoredFormula, valueCountFormula, valueGate, valueTopicCount } = context;
  sheet.getRange(`A${summary}:E${summary + 34}`).clear({ applyTo: "contents" });
  sheet.getRange(`A${summary}`).values = [["TOTAL SCORE"]];
  sheet.getRange(`C${summary}`).formulas = [[`=SUM(${scoreRows.map((scoreRow) => `B${scoreRow}`).join(",") || "0"})`]];
  sheet.getRange(`C${summary}`).format.numberFormat = '0.0" / 100"';
  sheet.getRange(`A${summary + 2}`).values = [["SCORE SUMMARY - reference signals for your verdict"]];
  sheet.getRange(`A${summary + 3}`).values = [["Unscored subtopics remaining"]];
  sheet.getRange(`C${summary + 3}`).formulas = [[`=${unscoredFormula}`]];
  sheet.getRange(`A${summary + 4}`).values = [["Discovery & qualification signal (flagship)"]];
  sheet.getRange(`C${summary + 4}`).formulas = [[`=IF(C${summary + 3}>0,"—",B${scoreRows[0] || 1})`]];
  sheet.getRange(`A${summary + 6}`).values = [["EVALUATOR-RATED SIGNALS (your judgment - rate both before the verdict)"]];
  sheet.getRange(`A${summary + 7}`).values = [["Was the presentation compelling?"]];
  sheet.getRange(`C${summary + 7}`).values = [[""]];
  sheet.getRange(`A${summary + 8}`).values = [["1 Would seek alternative vendors · 2 Doubts about the solution · 3 Likely to continue with Fortinet · 4 Unlikely to consider other vendors · 5 Will move forward with Fortinet"]];
  sheet.getRange(`A${summary + 9}`).values = [["Was the presentation accurate?"]];
  sheet.getRange(`C${summary + 9}`).values = [[""]];
  sheet.getRange(`A${summary + 10}`).values = [["1 Misrepresented the product · 2 Accurate but FAB unclear · 3 Accurate; FAB satisfactorily conveyed · 4 Accurate; FAB insightful and robust"]];
  sheet.getRange(`A${summary + 12}`).values = [["LAYER 1 - COMPETENCY (validator's judgment - use the signals above and pick one)"]];
  sheet.getRange(`A${summary + 13}`).values = [[""]];
  sheet.getRange(`A${summary + 15}`).values = [[config.layer2Header]];
  sheet.getRange(`A${summary + 16}`).values = [[`${config.valueCountLabel} (of ${valueTopicCount})`]];
  sheet.getRange(`C${summary + 16}`).formulas = [[`=${valueCountFormula}`]];
  sheet.getRange(`A${summary + 17}`).formulas = [[`=IF(C${summary + 3}>0,"—",IF(C${summary + 16}>=C${summary + 33},"${config.gateMet}","${config.gateNotMet}"))`]];
  sheet.getRange(`A${summary + 19}`).values = [["VALIDATOR DECISIONS"]];
  sheet.getRange(`A${summary + 20}`).values = [["Passing grade (follows your verdict)"]];
  sheet.getRange(`C${summary + 20}`).formulas = [[`=IF(A${summary + 13}="","—",IF(A${summary + 13}="Competency met","Yes","No"))`]];
  sheet.getRange(`A${summary + 21}`).values = [["If objections or tough questions arose - handled gracefully?"]];
  sheet.getRange(`A${summary + 22}`).values = [["Correlated customer needs to additional Fortinet products?"]];
  sheet.getRange(`A${summary + 23}`).values = [["Eligible to grade others on this pitch? (automatic)"]];
  sheet.getRange(`C${summary + 23}`).formulas = [[`=IF(C${summary + 3}>0,"—",IF(C${summary + 16}>=C${summary + 33},"Eligible","Not yet"))`]];
  sheet.getRange(`A${summary + 24}`).values = [["Role granted (only if the presenter volunteered)"]];
  sheet.getRange(`A${summary + 26}`).values = [["QUALITATIVE FEEDBACK"]];
  sheet.getRange(`A${summary + 27}`).values = [[config.feedbackLiked]];
  sheet.getRange(`A${summary + 28}`).values = [["Things that need more attention"]];
  sheet.getRange(`A${summary + 29}`).values = [["Additional feedback"]];
  sheet.getRange(`A${summary + 32}`).values = [["TEMPLATE SETTINGS - for adapting; leave alone while grading"]];
  sheet.getRange(`A${summary + 33}`).values = [[config.settingsLabel]];
  sheet.getRange(`C${summary + 33}`).values = [[valueGate]];
}

function applyPocHeader(sheet, rubric, config) {
  const productName = productNameForExport(rubric, config);
  const title = productName.toLowerCase().endsWith(config.title.toLowerCase())
    ? productName
    : `${productName || "[Product]"} - ${config.title}`;
  [
    "A1:E1", "A2:E2", "D4:E4", "D5:E5", "B6:E6", "A7:E7",
    "A9:E9", "A10:E10", "A11:E11", "A12:E12", "A13:E13", "A14:E14",
    "A16:E16", "C17:E17", "C18:E18", "C19:E19", "C20:E20", "C21:E21", "C22:E22", "A23:E23",
    "A25:E25",
  ].forEach((rangeAddress) => sheet.getRange(rangeAddress).merge());

  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A2").values = [[config.subtitle]];
  sheet.getRange("A1").format.font = { ...workbookFont, size: 16, bold: true };
  sheet.getRange("A2").format.font = { ...workbookFont, size: 9, italic: true, color: "#666666" };
  sheet.getRange("A4").values = [["Date"]];
  sheet.getRange("C4").values = [["Presenter (SE)"]];
  sheet.getRange("A5").values = [["Product"]];
  sheet.getRange("C5").values = [["Evaluator (SE)"]];
  sheet.getRange("A6").values = [["POC Scenario / Customer Context"]];
  sheet.getRange("B5").values = [[productName || "[Product name]"]];
  sheet.getRange("A7").values = [["Optional. Fill in only if the session was framed around a specific customer situation; leave blank for a generic run."]];
  sheet.getRange("A9").values = [["HOW TO USE"]];
  config.instructions.forEach((instruction, index) => {
    sheet.getRange(`A${10 + index}`).values = [[instruction]];
  });
  sheet.getRange("A16").values = [["SCORING LEVELS EXPLAINED - the same five apply to every subtopic"]];
  sheet.getRange("A25").values = [["SCORING"]];

  ["A4", "C4", "A5", "C5", "A6"].forEach((address) => {
    const range = sheet.getRange(address);
    range.format.fill = formLabelFill;
    range.format.font = { ...workbookFont, size: 10, bold: true };
    range.format.borders = { preset: "all", style: "thin", color: formBorder };
  });
  ["B4", "D4", "B5", "D5", "B6"].forEach((address) => {
    const range = sheet.getRange(address);
    range.format.fill = formInputFill;
    range.format.font = workbookFont;
    range.format.borders = { preset: "all", style: "thin", color: formBorder };
  });
  sheet.getRange("A7:E7").format.font = { ...workbookFont, italic: true, color: "#666666" };
  applyCellBorders(sheet, "A10:E14");
}

function applyPocSummaryBlockStyle(sheet, summary) {
  applyTableBorders(sheet, `A${summary}:E${summary + 28}`);
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
  sheet.getRange(`A${summary + 4}:B${summary + 4}`).format.font = { ...workbookFont, bold: true, color: "#FF0000" };
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

function writePocSummaryContent(sheet, summary, context) {
  const { config, scoreRows, unscoredFormula, valueCountFormula, valueGate, valueTopicCount } = context;
  sheet.getRange(`A${summary}:E${summary + 28}`).clear({ applyTo: "contents" });
  sheet.getRange(`A${summary}`).values = [["TOTAL SCORE"]];
  sheet.getRange(`C${summary}`).formulas = [[`=SUM(${scoreRows.map((scoreRow) => `B${scoreRow}`).join(",") || "0"})`]];
  sheet.getRange(`C${summary}`).format.numberFormat = '0.0" / 100"';
  sheet.getRange(`A${summary + 2}`).values = [["SCORE SUMMARY - reference signals for your verdict"]];
  sheet.getRange(`A${summary + 3}`).values = [["Unscored subtopics remaining"]];
  sheet.getRange(`C${summary + 3}`).formulas = [[`=${unscoredFormula}`]];
  sheet.getRange(`A${summary + 4}`).values = [["Flagship topic signal (floor 60%)"]];
  sheet.getRange(`C${summary + 4}`).formulas = [[`=IF(C${summary + 3}>0,"—",B${scoreRows[2] || scoreRows[0] || 1})`]];
  sheet.getRange(`A${summary + 6}`).values = [["LAYER 1 - COMPETENCY (validator's judgment - use the signals above and pick one)"]];
  sheet.getRange(`A${summary + 7}`).values = [[""]];
  sheet.getRange(`A${summary + 9}`).values = [[config.layer2Header]];
  sheet.getRange(`A${summary + 10}`).values = [[`${config.valueCountLabel} (of ${valueTopicCount})`]];
  sheet.getRange(`C${summary + 10}`).formulas = [[`=${valueCountFormula}`]];
  sheet.getRange(`A${summary + 11}`).formulas = [[`=IF(C${summary + 3}>0,"—",IF(C${summary + 10}>=C${summary + 25},"${config.gateMet}","${config.gateNotMet}"))`]];
  sheet.getRange(`A${summary + 13}`).values = [["VALIDATOR DECISIONS"]];
  sheet.getRange(`A${summary + 14}`).values = [["Passing grade (follows your verdict)"]];
  sheet.getRange(`C${summary + 14}`).formulas = [[`=IF(A${summary + 7}="","—",IF(A${summary + 7}="Competency met","Yes","No"))`]];
  sheet.getRange(`A${summary + 15}`).values = [["If something failed outside the induced scenarios, did the SE recover gracefully?"]];
  sheet.getRange(`A${summary + 16}`).values = [["Is the SE eligible to grade other SEs on this product? (automatic)"]];
  sheet.getRange(`C${summary + 16}`).formulas = [[`=IF(C${summary + 3}>0,"—",IF(C${summary + 10}>=C${summary + 25},"Eligible","Not yet"))`]];
  sheet.getRange(`A${summary + 17}`).values = [["Role granted (only if the SE volunteered)"]];
  sheet.getRange(`A${summary + 19}`).values = [["QUALITATIVE FEEDBACK"]];
  sheet.getRange(`A${summary + 20}`).values = [[config.feedbackLiked]];
  sheet.getRange(`A${summary + 21}`).values = [["Things that need more attention"]];
  sheet.getRange(`A${summary + 22}`).values = [["Additional feedback"]];
  sheet.getRange(`A${summary + 24}`).values = [["TEMPLATE SETTINGS - for adapting; hide after adapting"]];
  sheet.getRange(`A${summary + 25}`).values = [[config.settingsLabel]];
  sheet.getRange(`C${summary + 25}`).values = [[valueGate]];
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

async function createPocTemplateWorkbook(rubric) {
  const config = trackConfigs.poc;
  const totalWeight = (rubric.topics || []).reduce((total, topic) => total + (Number(topic.weight) || 0), 0);
  if (totalWeight > 100) throw new Error("Topic weights cannot exceed 100%.");

  const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(root, "src", "Demo Rubric Template.xlsx")));
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Rubric");
  sheet.showGridLines = false;
  const maxRows = 140;
  sheet.getRange(`A1:A${maxRows}`).format.columnWidth = 44;
  sheet.getRange(`B1:B${maxRows}`).format.columnWidth = 17.1796875;
  sheet.getRange(`C1:C${maxRows}`).format.columnWidth = 16.08984375;
  sheet.getRange(`D1:D${maxRows}`).format.columnWidth = 18;
  sheet.getRange(`E1:E${maxRows}`).format.columnWidth = 44;
  sheet.getRange(`A1:E${maxRows}`).format.font = workbookFont;
  [27.75, 15, 9, 15, 15, 15, 15, 9, 19.5, 39.75, 39.75, 39.75, 39.75, 39.75, 9, 19.5, 15, 25.5, 25.5, 25.5, 25.5, 39.75, 39.75, 9, 19.5, 18].forEach((height, index) => {
    sheet.getRange(`A${index + 1}:E${index + 1}`).format.rowHeight = height;
  });

  applyPocHeader(sheet, rubric, config);
  ["A9:E9", "A16:E16", "A25:E25"].forEach((rangeAddress) => applySectionHeaderStyle(sheet, rangeAddress));
  applyScoringLegendStyleAt(sheet, config, 16);
  applyScoringTableHeader(sheet, 26);

  let row = 27;
  const scoreRows = [];
  const unscoredRanges = [];
  const valueCountFormulas = [];
  const subtopicRows = [];
  for (const [index, topic] of (rubric.topics || []).entries()) {
    const subtopics = topic.subtopics || [];
    const weight = Number(topic.weight) || 0;
    sheet.getRange(`A${row}:E${row}`).merge();
    sheet.getRange(`A${row}`).values = [[index === 2 ? `${index + 1} - ${topic.name || "Untitled topic"} (weight ${weight}) ◆ GATING - must score ≥ 60% (12/20)` : `${index + 1} - ${topic.name || "Untitled topic"} (weight ${weight})`]];
    sheet.getRange(`A${row}:E${row}`).format.rowHeight = 21.75;
    if (index === 2) {
      sheet.getRange(`A${row}:E${row}`).format.fill = "#E1251B";
      sheet.getRange(`A${row}:E${row}`).format.font = { ...workbookFont, size: 11, bold: true, color: "#FFFFFF" };
    } else {
      applyTopicBandStyle(sheet, row);
    }
    row += 1;

    const start = row;
    for (const item of subtopics) {
      const subtopic = typeof item === "string" ? { name: item } : item;
      const subtopicName = subtopic.name || "Evidence item";
      sheet.getRange(`A${row}:E${row}`).values = [[subtopicName, "", null, null, null]];
      sheet.getRange(`A${row}:E${row}`).format.rowHeight = Math.max(27, subtopicRowHeight(subtopicName));
      sheet.getRange(`C${row}`).formulas = [[levelFormula(row, config)]];
      sheet.getRange(`B${row}`).dataValidation = { rule: { type: "list", values: config.levels } };
      subtopicRows.push(row);
      row += 1;
    }
    const end = row - 1;
    if (subtopics.length) unscoredRanges.push(`B${start}:B${end}`);

    sheet.getRange(`A${row}:E${row}`).values = [[`▸ Topic ${index + 1} score (average x weight)`, null, null, "Notes:", null]];
    sheet.getRange(`B${row}`).formulas = [[`=IFERROR(AVERAGE(C${start}:C${end})*${weight}/100,0)`]];
    sheet.getRange(`B${row}`).format.numberFormat = `0.0" / ${weight}"`;
    sheet.getRange(`A${row}:E${row}`).format.rowHeight = 18;
    applyTopicScoreRowStyle(sheet, row);
    scoreRows.push(row);
    valueCountFormulas.push(`COUNTIF(C${start}:C${end},100)`);
    row += 1;
  }

  const summary = row + 1;
  [21.75, 9, 19.5, 15, 15, 9, 19.5, 39.75, 9, 19.5, 15, 21.75, 9, 19.5, 15, 15, 15, 15, 9, 19.5, 39.75, 39.75, 39.75, 9, 19.5, 15, 15, 15, 15].forEach((height, index) => {
    sheet.getRange(`A${summary + index}:E${summary + index}`).format.rowHeight = height;
  });
  [2, 6, 7, 9, 11, 13, 19, 24].forEach((offset) => sheet.getRange(`A${summary + offset}:E${summary + offset}`).merge());
  [3, 4, 10, 14, 15, 16, 20, 21, 22, 25].forEach((offset) => {
    sheet.getRange(`A${summary + offset}:B${summary + offset}`).merge();
    sheet.getRange(`C${summary + offset}:E${summary + offset}`).merge();
  });
  const unscoredFormula = unscoredRanges.map((range) => `COUNTBLANK(${range})`).join("+") || "0";
  const valueCountFormula = valueCountFormulas.join("+") || "0";
  writePocSummaryContent(sheet, summary, {
    config,
    scoreRows,
    unscoredFormula,
    valueCountFormula,
    valueGate: Number(rubric.valueGate) || 0,
    valueTopicCount: (rubric.topics || []).length,
  });
  [
    `A${summary}:E${summary}`,
    `A${summary + 2}:E${summary + 2}`,
    `A${summary + 6}:E${summary + 6}`,
    `A${summary + 9}:E${summary + 9}`,
    `A${summary + 13}:E${summary + 13}`,
    `A${summary + 19}:E${summary + 19}`,
    `A${summary + 24}:E${summary + 24}`,
  ].forEach((rangeAddress) => applySectionHeaderStyle(sheet, rangeAddress));
  subtopicRows.forEach((subtopicRow) => applySubtopicRowStyle(sheet, subtopicRow));
  applyPocSummaryBlockStyle(sheet, summary);

  const verdictOptions = ["", "Competency met", "Targeted re-validation - name the topic in notes", "Full re-validation (after coaching)"];
  const yesNoOptions = ["", "Yes", "No", "N/A"];
  const roleOptions = ["", "-", "Validator", "Mentor"];
  sheet.getRange(`A${summary + 7}`).dataValidation = { rule: { type: "list", values: verdictOptions } };
  sheet.getRange(`C${summary + 15}`).dataValidation = { rule: { type: "list", values: yesNoOptions } };
  sheet.getRange(`C${summary + 17}`).dataValidation = { rule: { type: "list", values: roleOptions } };

  const adaptation = workbook.worksheets.add("How to adapt");
  adaptation.getRange("A1:A26").format.columnWidth = 4;
  adaptation.getRange("B1:B26").format.columnWidth = 112;
  adaptation.getRange("A1:B26").copyFrom(sourceWorkbook.worksheets.getItem("How to adapt").getRange("A1:B26"), "all");
  adaptation.getRange("A1:B26").format.font = workbookFont;
  return workbook;
}

async function createTemplateWorkbook(rubric, track = "demo") {
  if (track === "poc") return createPocTemplateWorkbook(rubric);
  const config = trackConfigs[track] || trackConfigs.demo;
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
      const subtopicName = subtopic.name || "Evidence item";
      sheet.getRange(`A${row}:E${row}`).copyFrom(source.getRange("A25:E25"), "all");
      sheet.getRange(`A${row}:E${row}`).format.rowHeight = subtopicRowHeight(subtopicName);
      sheet.getRange(`A${row}`).values = [[subtopicName]];
      sheet.getRange(`B${row}`).values = [[""]];
      sheet.getRange(`C${row}`).formulas = [[levelFormula(row, config)]];
      sheet.getRange(`B${row}`).dataValidation = { rule: { type: "list", values: config.levels } };
      subtopicRows.push(row);
      row += 1;
    }
    const end = row - 1;
    if (subtopics.length) unscoredRanges.push(`B${start}:B${end}`);
    sheet.getRange(`A${row}:E${row}`).copyFrom(source.getRange("A29:E29"), "all");
    sheet.getRange(`A${row}:E${row}`).format.rowHeight = 18;
    sheet.getRange(`A${row}`).values = [[`▸ Topic ${index + 1} score (average x weight)`]];
    sheet.getRange(`B${row}`).formulas = [[`=IFERROR(AVERAGE(C${start}:C${end})*${weight}/100,0)`]];
    sheet.getRange(`B${row}`).format.numberFormat = `0.0" / ${weight}"`;
    sheet.getRange(`D${row}`).values = [["Notes:"]];
    scoreRows.push(row);
    if (!(config.excludeFinalTopicFromGate && index === (rubric.topics || []).length - 1)) {
      valueCountFormulas.push(`COUNTIF(C${start}:C${end},100)`);
    }
    row += 1;
  }

  const summary = row + 1;
  const summaryRows = track === "pitch" ? 34 : 25;
  sheet.getRange(`A1:A${summary + summaryRows}`).format.columnWidth = 44;
  sheet.getRange(`B1:B${summary + summaryRows}`).format.columnWidth = 17.1796875;
  sheet.getRange(`C1:C${summary + summaryRows}`).format.columnWidth = 16.08984375;
  sheet.getRange(`D1:D${summary + summaryRows}`).format.columnWidth = 18;
  sheet.getRange(`E1:E${summary + summaryRows}`).format.columnWidth = 44;
  sheet.getRange(`A${summary}:E${summary + 25}`).copyFrom(source.getRange("A61:E86"), "all");
  [21.75, 9, 19.5, 15, 15, 9, 19.5, 24, 9, 19.5, 15, 21.75, 9, 19.5, 15, 15, 15, 15, 9, 19.5, 39.75, 39.75, 39.75, 9, 19.5, 15].forEach((height, index) => {
    sheet.getRange(`A${summary + index}:E${summary + index}`).format.rowHeight = height;
  });
  if (track === "pitch") {
    [24, 15, 15, 15, 15, 9, 19.5, 15, 15, 15, 15, 9, 19.5, 24, 9, 19.5, 15, 24, 9, 19.5, 15, 15, 15, 15, 15, 9, 19.5, 39.75, 39.75, 39.75, 9, 19.5, 15, 15, 15].forEach((height, index) => {
      sheet.getRange(`A${summary + index}:E${summary + index}`).format.rowHeight = height;
    });
  }
  const fullWidthMergeOffsets = track === "pitch" ? [2, 6, 8, 10, 12, 13, 15, 17, 19, 26, 32] : [2, 6, 7, 9, 11, 13, 19, 24];
  fullWidthMergeOffsets.forEach((offset) => sheet.getRange(`A${summary + offset}:E${summary + offset}`).merge());
  const splitMergeOffsets = track === "pitch"
    ? [3, 4, 7, 9, 16, 20, 21, 22, 23, 24, 27, 28, 29, 33]
    : [3, 4, 10, 14, 15, 16, 20, 21, 22, 25];
  splitMergeOffsets.forEach((offset) => {
    sheet.getRange(`A${summary + offset}:B${summary + offset}`).merge();
    sheet.getRange(`C${summary + offset}:E${summary + offset}`).merge();
  });
  const unscoredFormula = unscoredRanges.map((range) => `COUNTBLANK(${range})`).join("+") || "0";
  const valueCountFormula = valueCountFormulas.join("+") || "0";
  const summaryContext = {
    config,
    scoreRows,
    unscoredFormula,
    valueCountFormula,
    valueGate: Number(rubric.valueGate) || 0,
    valueTopicCount: config.excludeFinalTopicFromGate ? Math.max(0, (rubric.topics || []).length - 1) : (rubric.topics || []).length,
  };
  if (track === "pitch") writePitchSummaryContent(sheet, summary, summaryContext);
  else writeDemoSummaryContent(sheet, summary, summaryContext);
  sheet.getRange(`A1:E${summary + summaryRows}`).format.font = workbookFont;
  applyHeaderSpec(sheet, rubric, config);
  const sectionHeaderRanges = track === "pitch" ? [
    "A7:E7",
    "A13:E13",
    "A22:E22",
    `A${summary}:E${summary}`,
    `A${summary + 2}:E${summary + 2}`,
    `A${summary + 6}:E${summary + 6}`,
    `A${summary + 12}:E${summary + 12}`,
    `A${summary + 15}:E${summary + 15}`,
    `A${summary + 19}:E${summary + 19}`,
    `A${summary + 26}:E${summary + 26}`,
    `A${summary + 32}:E${summary + 32}`,
  ] : [
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
  ];
  sectionHeaderRanges.forEach((range) => applySectionHeaderStyle(sheet, range));
  applyCellBorders(sheet, "A8:E11");
  applyScoringLegendStyle(sheet, config);
  applyScoringTableHeader(sheet, 23);
  topicBandRows.forEach((topicRow) => applyTopicBandStyle(sheet, topicRow));
  subtopicRows.forEach((subtopicRow) => applySubtopicRowStyle(sheet, subtopicRow));
  scoreRows.forEach((scoreRow) => applyTopicScoreRowStyle(sheet, scoreRow));
  if (track === "pitch") applyPitchSummaryBlockStyle(sheet, summary);
  else applySummaryBlockStyle(sheet, summary);

  const demoVerdictOptions = ["", "Competency met", "Targeted re-validation - name the topic in notes", "Full re-validation (after coaching)"];
  const pitchVerdictOptions = ["", "Competency met", "Targeted re-validation - name the topic in notes", "Full re-validation (after coaching)"];
  const yesNoOptions = ["", "Yes", "No", "N/A"];
  const roleOptions = ["", "-", "Validator", "Mentor"];
  if (track === "pitch") {
    sheet.getRange(`C${summary + 7}`).dataValidation = { rule: { type: "list", values: ["", "1 - Would seek alternative vendors", "2 - Doubts about the solution", "3 - Likely to continue with Fortinet", "4 - Unlikely to consider other vendors", "5 - Will move forward with Fortinet"] } };
    sheet.getRange(`C${summary + 9}`).dataValidation = { rule: { type: "list", values: ["", "1 - Misrepresented the product", "2 - Accurate but FAB unclear", "3 - Accurate; FAB satisfactorily conveyed", "4 - Accurate; FAB insightful and robust"] } };
    sheet.getRange(`A${summary + 13}`).dataValidation = { rule: { type: "list", values: pitchVerdictOptions } };
    sheet.getRange(`C${summary + 21}`).dataValidation = { rule: { type: "list", values: yesNoOptions } };
    sheet.getRange(`C${summary + 22}`).dataValidation = { rule: { type: "list", values: yesNoOptions } };
    sheet.getRange(`C${summary + 24}`).dataValidation = { rule: { type: "list", values: roleOptions } };
  } else {
    sheet.getRange(`A${summary + 7}`).dataValidation = { rule: { type: "list", values: demoVerdictOptions } };
    sheet.getRange(`C${summary + 15}`).dataValidation = { rule: { type: "list", values: yesNoOptions } };
    sheet.getRange(`C${summary + 17}`).dataValidation = { rule: { type: "list", values: roleOptions } };
  }

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

  applyHeaderSpec(sheet, rubric, trackConfigs.demo);
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
    const track = trackConfigs[rubric.track] ? rubric.track : "demo";
    const workbook = await createTemplateWorkbook(rubric, track);
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
