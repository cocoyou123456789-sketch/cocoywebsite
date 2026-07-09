import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/cocoyou/data/cocoywebsite/outputs/financial_analysis_workbook";
const workbook = Workbook.create();

const sheets = {
  cover: workbook.worksheets.add("Cover"),
  assumptions: workbook.worksheets.add("Assumptions"),
  historical: workbook.worksheets.add("Historical Financials"),
  model: workbook.worksheets.add("Forecast Model"),
  valuation: workbook.worksheets.add("Valuation"),
  sensitivity: workbook.worksheets.add("Sensitivity"),
  checks: workbook.worksheets.add("Checks"),
  sources: workbook.worksheets.add("Sources Audit"),
};

const fmt = {
  titleFill: "#0F172A",
  darkFill: "#1E3A5F",
  subFill: "#E8EEF6",
  noteFill: "#F8FAFC",
  inputFill: "#FFF2CC",
  okFill: "#D9EAD3",
  warnFill: "#FCE4D6",
  border: "#B7C3D0",
  lightBorder: "#D9E2EC",
  inputBlue: "#0000FF",
  formulaBlack: "#000000",
  linkGreen: "#008000",
  white: "#FFFFFF",
};

const moneyFmt = '$#,##0;[Red]($#,##0);-';
const countFmt = '#,##0;[Red](#,##0);-';
const pctFmt = '0.0%;[Red](0.0%);-';
const multFmt = '0.0x;[Red](0.0x);-';
const perShareFmt = '$0.00;[Red]($0.00);-';
const dateFmt = "yyyy-mm-dd";

function styleSheet(sheet, maxCol = "K") {
  sheet.showGridLines = false;
  sheet.getRange(`A:${maxCol}`).format.font = { name: "Aptos", size: 10, color: fmt.formulaBlack };
  sheet.getRange(`A:${maxCol}`).format.verticalAlignment = "middle";
}

function title(sheet, range, text) {
  const r = sheet.getRange(range);
  r.format.fill = { color: fmt.titleFill };
  r.format.font = { bold: true, size: 16, color: fmt.titleFill };
  r.format.horizontalAlignment = "left";
  r.format.rowHeight = 28;
  sheet.getRange(range.split(":")[0]).values = [[text]];
}

function section(sheet, range, text) {
  const r = sheet.getRange(range);
  r.format.fill = { color: fmt.darkFill };
  r.format.font = { bold: true, color: fmt.darkFill };
  r.format.borders = { preset: "outside", style: "thin", color: fmt.darkFill };
  sheet.getRange(range.split(":")[0]).values = [[text]];
}

function header(range) {
  range.format.fill = { color: fmt.subFill };
  range.format.font = { bold: true, color: fmt.formulaBlack };
  range.format.borders = { preset: "outside", style: "thin", color: fmt.border };
  range.format.horizontalAlignment = "center";
}

function total(range) {
  range.format.font = { bold: true, color: fmt.formulaBlack };
  range.format.borders = { top: { style: "thin", color: fmt.formulaBlack } };
}

function input(range) {
  range.format.fill = { color: fmt.inputFill };
  range.format.font = { color: fmt.inputBlue };
}

function formulas(range, color = fmt.formulaBlack) {
  range.format.font = { color };
}

function note(range) {
  range.format.fill = { color: fmt.noteFill };
  range.format.font = { color: "#475569", italic: true };
  range.format.wrapText = true;
}

function widths(sheet, pairs) {
  for (const [col, width] of pairs) sheet.getRange(`${col}:${col}`).format.columnWidth = width;
}

function rows(sheet, pairs) {
  for (const [row, height] of pairs) sheet.getRange(`${row}:${row}`).format.rowHeight = height;
}

for (const sheet of Object.values(sheets)) styleSheet(sheet, "L");

workbook.comments.setSelf({ displayName: "User" });

// Cover
title(sheets.cover, "A1:H1", "Company Financial Analysis Workbook");
widths(sheets.cover, [["A", 26], ["B", 24], ["C", 14], ["D", 24], ["E", 14], ["F", 25], ["G", 14], ["H", 25], ["J", 12], ["K", 14], ["L", 14], ["M", 12], ["N", 14]]);
rows(sheets.cover, [[3, 24], [8, 24], [17, 24]]);
sheets.cover.getRange("A3:H3").values = [["Workbook purpose", "Analyze historical financials, forecast operating performance, run a DCF valuation, test sensitivities, and surface model checks.", null, null, null, null, null, null]];
sheets.cover.getRange("A3:H3").format.wrapText = true;
section(sheets.cover, "A5:H5", "Key Outputs");
sheets.cover.getRange("A6:H7").values = [
  ["Company", "Scenario", "Revenue CAGR", "Terminal EBITDA Margin", "Enterprise Value", "Equity Value", "Implied Share Price", "Model Status"],
  [null, null, null, null, null, null, null, null],
];
header(sheets.cover.getRange("A6:H6"));
sheets.cover.getRange("A7:H7").formulas = [[
  "='Assumptions'!B4",
  "='Assumptions'!B8",
  "='Forecast Model'!K23",
  "='Forecast Model'!K24",
  "='Valuation'!B15",
  "='Valuation'!B17",
  "='Valuation'!B19",
  "='Checks'!F4",
]];
formulas(sheets.cover.getRange("A7:H7"), fmt.linkGreen);
sheets.cover.getRange("C7:D7").setNumberFormat(pctFmt);
sheets.cover.getRange("E7:F7").setNumberFormat(moneyFmt);
sheets.cover.getRange("G7:G7").setNumberFormat(perShareFmt);
sheets.cover.getRange("H7").conditionalFormats.add("containsText", { text: "OK", format: { fill: { color: fmt.okFill }, font: { bold: true, color: "#166534" } } });
sheets.cover.getRange("H7").conditionalFormats.add("containsText", { text: "Review", format: { fill: { color: "#9A3412" }, font: { bold: true, color: "#FFFFFF" } } });

section(sheets.cover, "A10:H10", "Workbook Map");
sheets.cover.getRange("A11:H15").values = [
  ["Sheet", "Use", null, null, null, null, null, null],
  ["Assumptions", "Company setup, scenario controls, valuation assumptions, and forecast drivers.", null, null, null, null, null, null],
  ["Historical Financials", "Paste reported annual financials and source IDs.", null, null, null, null, null, null],
  ["Forecast Model / Valuation", "Formula-driven forecast, unlevered FCF, DCF bridge, and sensitivity outputs.", null, null, null, null, null, null],
  ["Checks / Sources Audit", "Input completeness, calculation sanity checks, and source documentation.", null, null, null, null, null, null],
];
header(sheets.cover.getRange("A11:B11"));
sheets.cover.getRange("A12:B15").format.borders = { preset: "inside", style: "thin", color: fmt.lightBorder };
sheets.cover.getRange("B12:B15").format.wrapText = true;

section(sheets.cover, "A17:H17", "Formatting Legend");
sheets.cover.getRange("A18:H20").values = [
  ["Blue text / yellow fill", "Editable inputs", "Black text", "Formulas", "Green text", "Workbook links", "Red text", "External workbook links"],
  [null, null, null, null, null, null, null, null],
  ["Source note", "Illustrative placeholder data is included so formulas and charts compute. Replace with company filings or exports before relying on outputs.", null, null, null, null, null, null],
];
input(sheets.cover.getRange("A18:B18"));
formulas(sheets.cover.getRange("C18:D18"));
formulas(sheets.cover.getRange("E18:F18"), fmt.linkGreen);
sheets.cover.getRange("G18:H18").format.font = { color: "#FF0000" };
note(sheets.cover.getRange("A20:H20"));
sheets.cover.freezePanes.freezeRows(1);

// Assumptions
title(sheets.assumptions, "A1:F1", "Assumptions");
widths(sheets.assumptions, [["A", 28], ["B", 18], ["C", 16], ["D", 18], ["E", 18], ["F", 42]]);
section(sheets.assumptions, "A3:F3", "Company Setup");
sheets.assumptions.getRange("A4:F9").values = [
  ["Company Name", "Company", null, null, null, "Replace with target company name"],
  ["Ticker", "TICKER", null, null, null, "Optional identifier"],
  ["Currency", "USD", null, null, null, "Workbook amounts are in millions unless noted"],
  ["Units", "$mm", null, null, null, "Financial statements and valuation outputs"],
  ["Scenario", "Base", null, null, null, "Use dropdown: Downside / Base / Upside"],
  ["Valuation Date", new Date("2026-06-25"), null, null, null, "Update to analysis date"],
];
input(sheets.assumptions.getRange("B4:B9"));
sheets.assumptions.getRange("B8").dataValidation = { rule: { type: "list", values: ["Downside", "Base", "Upside"] } };
sheets.assumptions.getRange("B9").setNumberFormat(dateFmt);
header(sheets.assumptions.getRange("A4:F4"));
sheets.assumptions.getRange("F4:F9").format.wrapText = true;

section(sheets.assumptions, "A11:F11", "Valuation Assumptions");
sheets.assumptions.getRange("A12:F19").values = [
  ["WACC", 0.095, null, null, null, "Discount rate for unlevered FCF"],
  ["Terminal Growth", 0.025, null, null, null, "Gordon growth terminal value"],
  ["Effective Tax Rate", 0.23, null, null, null, "Used in NOPAT calculation"],
  ["Net Debt / (Cash)", 2500, null, null, null, "Positive value reduces equity value"],
  ["Diluted Shares Outstanding", 420, null, null, null, "Millions of shares"],
  ["Mid-Year Convention", "No", null, null, null, "Template uses year-end discounting"],
  ["Source Data Mode", "Illustrative placeholder", null, null, null, "Replace historicals and sources before investment use"],
  ["EV/EBITDA Cross-Check Multiple", 10.0, null, null, null, "Used for sanity check, not primary DCF"],
];
input(sheets.assumptions.getRange("B12:B19"));
sheets.assumptions.getRange("B12:B14").setNumberFormat(pctFmt);
sheets.assumptions.getRange("B15:B16").setNumberFormat(countFmt);
sheets.assumptions.getRange("B19").setNumberFormat(multFmt);
sheets.assumptions.getRange("F12:F19").format.wrapText = true;

section(sheets.assumptions, "A21:F21", "Forecast Drivers By Scenario");
sheets.assumptions.getRange("A22:F28").values = [
  ["Driver", "Downside", "Base", "Upside", "Selected", "Notes"],
  ["Revenue Growth", 0.035, 0.060, 0.085, null, "Applied to first forecast year, then tapers modestly"],
  ["Gross Margin", 0.410, 0.430, 0.455, null, "Forecast gross profit / revenue"],
  ["EBITDA Margin", 0.155, 0.180, 0.210, null, "Forecast EBITDA / revenue"],
  ["D&A / Revenue", 0.035, 0.032, 0.030, null, "Depreciation and amortization intensity"],
  ["Capex / Revenue", 0.060, 0.050, 0.045, null, "Capital expenditure intensity"],
  ["NWC / Revenue", 0.130, 0.115, 0.100, null, "Net working capital investment as % of revenue"],
];
header(sheets.assumptions.getRange("A22:F22"));
input(sheets.assumptions.getRange("B23:D28"));
sheets.assumptions.getRange("E23:E28").formulas = [
  ["=IF($B$8=\"Downside\",B23,IF($B$8=\"Upside\",D23,C23))"],
  ["=IF($B$8=\"Downside\",B24,IF($B$8=\"Upside\",D24,C24))"],
  ["=IF($B$8=\"Downside\",B25,IF($B$8=\"Upside\",D25,C25))"],
  ["=IF($B$8=\"Downside\",B26,IF($B$8=\"Upside\",D26,C26))"],
  ["=IF($B$8=\"Downside\",B27,IF($B$8=\"Upside\",D27,C27))"],
  ["=IF($B$8=\"Downside\",B28,IF($B$8=\"Upside\",D28,C28))"],
];
formulas(sheets.assumptions.getRange("E23:E28"));
sheets.assumptions.getRange("B23:E28").setNumberFormat(pctFmt);
sheets.assumptions.getRange("B19").setNumberFormat(multFmt);
sheets.assumptions.getRange("A22:F28").format.borders = { preset: "inside", style: "thin", color: fmt.lightBorder };
sheets.assumptions.getRange("F23:F28").format.wrapText = true;
sheets.assumptions.freezePanes.freezeRows(22);

for (const addr of ["B12", "B13", "B15", "B16", "B23"]) {
  workbook.comments.addThread({ cell: sheets.assumptions.getRange(addr) }, "Source: user input or latest company filings / market data. Placeholder value included for template functionality.");
}

// Historical Financials
title(sheets.historical, "A1:H1", "Historical Financials");
widths(sheets.historical, [["A", 30], ["B", 14], ["C", 14], ["D", 14], ["E", 14], ["F", 14], ["G", 14], ["H", 42]]);
sheets.historical.getRange("A3:H3").values = [["Paste reported annual financials below; replace illustrative placeholders with company data and cite source IDs.", null, null, null, null, null, null, null]];
note(sheets.historical.getRange("A3:H3"));
section(sheets.historical, "A5:H5", "Reported Financials");
sheets.historical.getRange("A6:H20").values = [
  ["Metric ($mm except per-share/share data)", 2022, 2023, 2024, 2025, 2026, "Source ID", "Notes"],
  ["Revenue", 12000, 13250, 14600, 15900, 17100, "SRC-001", "Illustrative placeholder"],
  ["Gross Profit", 4800, 5366, 6001, 6678, 7353, "SRC-001", "Illustrative placeholder"],
  ["EBITDA", 1800, 2070, 2409, 2783, 3164, "SRC-001", "Illustrative placeholder"],
  ["D&A", 420, 450, 480, 510, 540, "SRC-001", "Illustrative placeholder"],
  ["EBIT", 1380, 1620, 1929, 2273, 2624, "SRC-001", "Illustrative placeholder"],
  ["Interest Expense", 125, 130, 135, 130, 125, "SRC-001", "Illustrative placeholder"],
  ["Pre-Tax Income", 1255, 1490, 1794, 2143, 2499, "SRC-001", "Illustrative placeholder"],
  ["Tax Expense", 289, 343, 413, 493, 575, "SRC-001", "Illustrative placeholder"],
  ["Net Income", 966, 1147, 1381, 1650, 1924, "SRC-001", "Illustrative placeholder"],
  ["Cash From Operations", 1300, 1510, 1740, 2020, 2320, "SRC-001", "Illustrative placeholder"],
  ["Capital Expenditures", -700, -760, -825, -880, -930, "SRC-001", "Illustrative placeholder"],
  ["Cash", 1500, 1680, 1825, 2050, 2300, "SRC-001", "Illustrative placeholder"],
  ["Debt", 4200, 4100, 4050, 3900, 3800, "SRC-001", "Illustrative placeholder"],
  ["Diluted Shares", 440, 435, 430, 425, 420, "SRC-001", "Illustrative placeholder"],
];
header(sheets.historical.getRange("A6:H6"));
input(sheets.historical.getRange("B7:H21"));
sheets.historical.getRange("B7:F21").setNumberFormat(countFmt);
sheets.historical.getRange("A6:H21").format.borders = { preset: "inside", style: "thin", color: fmt.lightBorder };
sheets.historical.getRange("H7:H21").format.wrapText = true;
sheets.historical.freezePanes.freezeRows(6);

// Forecast Model
title(sheets.model, "A1:K1", "Forecast Model");
widths(sheets.model, [["A", 30], ["B", 13], ["C", 13], ["D", 13], ["E", 13], ["F", 13], ["G", 13], ["H", 13], ["I", 13], ["J", 13], ["K", 13]]);
section(sheets.model, "A3:K3", "Historical Actuals and Forecast");
sheets.model.getRange("A4:K24").values = [
  ["Metric", "2022A", "2023A", "2024A", "2025A", "2026A", "2027E", "2028E", "2029E", "2030E", "2031E"],
  ["Revenue", null, null, null, null, null, null, null, null, null, null],
  ["Revenue Growth", null, null, null, null, null, null, null, null, null, null],
  ["Gross Profit", null, null, null, null, null, null, null, null, null, null],
  ["Gross Margin", null, null, null, null, null, null, null, null, null, null],
  ["EBITDA", null, null, null, null, null, null, null, null, null, null],
  ["EBITDA Margin", null, null, null, null, null, null, null, null, null, null],
  ["D&A", null, null, null, null, null, null, null, null, null, null],
  ["D&A / Revenue", null, null, null, null, null, null, null, null, null, null],
  ["EBIT", null, null, null, null, null, null, null, null, null, null],
  ["EBIT Margin", null, null, null, null, null, null, null, null, null, null],
  ["Cash Taxes", null, null, null, null, null, null, null, null, null, null],
  ["NOPAT", null, null, null, null, null, null, null, null, null, null],
  ["Capital Expenditures", null, null, null, null, null, null, null, null, null, null],
  ["Capex / Revenue", null, null, null, null, null, null, null, null, null, null],
  ["Net Working Capital", null, null, null, null, null, null, null, null, null, null],
  ["Change in NWC", null, null, null, null, null, null, null, null, null, null],
  ["Unlevered Free Cash Flow", null, null, null, null, null, null, null, null, null, null],
  ["FCF Margin", null, null, null, null, null, null, null, null, null, null],
  ["Revenue CAGR", null, null, null, null, null, null, null, null, null, null],
  ["Terminal EBITDA Margin", null, null, null, null, null, null, null, null, null, null],
];
header(sheets.model.getRange("A4:K4"));
sheets.model.getRange("B5:F23").formulas = [
  ["='Historical Financials'!B7","='Historical Financials'!C7","='Historical Financials'!D7","='Historical Financials'!E7","='Historical Financials'!F7"],
  ["=IFERROR(C5/B5-1,0)","=IFERROR(D5/C5-1,0)","=IFERROR(E5/D5-1,0)","=IFERROR(F5/E5-1,0)","=IFERROR(F5/E5-1,0)"],
  ["='Historical Financials'!B8","='Historical Financials'!C8","='Historical Financials'!D8","='Historical Financials'!E8","='Historical Financials'!F8"],
  ["=IFERROR(B7/B5,0)","=IFERROR(C7/C5,0)","=IFERROR(D7/D5,0)","=IFERROR(E7/E5,0)","=IFERROR(F7/F5,0)"],
  ["='Historical Financials'!B9","='Historical Financials'!C9","='Historical Financials'!D9","='Historical Financials'!E9","='Historical Financials'!F9"],
  ["=IFERROR(B9/B5,0)","=IFERROR(C9/C5,0)","=IFERROR(D9/D5,0)","=IFERROR(E9/E5,0)","=IFERROR(F9/F5,0)"],
  ["='Historical Financials'!B10","='Historical Financials'!C10","='Historical Financials'!D10","='Historical Financials'!E10","='Historical Financials'!F10"],
  ["=IFERROR(B11/B5,0)","=IFERROR(C11/C5,0)","=IFERROR(D11/D5,0)","=IFERROR(E11/E5,0)","=IFERROR(F11/F5,0)"],
  ["='Historical Financials'!B11","='Historical Financials'!C11","='Historical Financials'!D11","='Historical Financials'!E11","='Historical Financials'!F11"],
  ["=IFERROR(B13/B5,0)","=IFERROR(C13/C5,0)","=IFERROR(D13/D5,0)","=IFERROR(E13/E5,0)","=IFERROR(F13/F5,0)"],
  ["='Historical Financials'!B15","='Historical Financials'!C15","='Historical Financials'!D15","='Historical Financials'!E15","='Historical Financials'!F15"],
  ["=B13-B15","=C13-C15","=D13-D15","=E13-E15","=F13-F15"],
  ["='Historical Financials'!B18","='Historical Financials'!C18","='Historical Financials'!D18","='Historical Financials'!E18","='Historical Financials'!F18"],
  ["=IFERROR(B17/B5,0)","=IFERROR(C17/C5,0)","=IFERROR(D17/D5,0)","=IFERROR(E17/E5,0)","=IFERROR(F17/F5,0)"],
  ["=B5*'Assumptions'!$E$28","=C5*'Assumptions'!$E$28","=D5*'Assumptions'!$E$28","=E5*'Assumptions'!$E$28","=F5*'Assumptions'!$E$28"],
  ["=0","=C19-B19","=D19-C19","=E19-D19","=F19-E19"],
  ["=B16+B11+B17-B20","=C16+C11+C17-C20","=D16+D11+D17-D20","=E16+E11+E17-E20","=F16+F11+F17-F20"],
  ["=IFERROR(B21/B5,0)","=IFERROR(C21/C5,0)","=IFERROR(D21/D5,0)","=IFERROR(E21/E5,0)","=IFERROR(F21/F5,0)"],
  ["=0","=0","=0","=0","=0"],
];
formulas(sheets.model.getRange("B5:F23"), fmt.linkGreen);
sheets.model.getRange("G5:K23").formulas = [
  ["=F5*(1+'Assumptions'!$E$23)","=G5*(1+MAX('Assumptions'!$E$23-0.005,0))","=H5*(1+MAX('Assumptions'!$E$23-0.010,0))","=I5*(1+MAX('Assumptions'!$E$23-0.015,0))","=J5*(1+MAX('Assumptions'!$E$23-0.020,0))"],
  ["=IFERROR(G5/F5-1,0)","=IFERROR(H5/G5-1,0)","=IFERROR(I5/H5-1,0)","=IFERROR(J5/I5-1,0)","=IFERROR(K5/J5-1,0)"],
  ["=G5*'Assumptions'!$E$24","=H5*'Assumptions'!$E$24","=I5*'Assumptions'!$E$24","=J5*'Assumptions'!$E$24","=K5*'Assumptions'!$E$24"],
  ["=IFERROR(G7/G5,0)","=IFERROR(H7/H5,0)","=IFERROR(I7/I5,0)","=IFERROR(J7/J5,0)","=IFERROR(K7/K5,0)"],
  ["=G5*'Assumptions'!$E$25","=H5*'Assumptions'!$E$25","=I5*'Assumptions'!$E$25","=J5*'Assumptions'!$E$25","=K5*'Assumptions'!$E$25"],
  ["=IFERROR(G9/G5,0)","=IFERROR(H9/H5,0)","=IFERROR(I9/I5,0)","=IFERROR(J9/J5,0)","=IFERROR(K9/K5,0)"],
  ["=G5*'Assumptions'!$E$26","=H5*'Assumptions'!$E$26","=I5*'Assumptions'!$E$26","=J5*'Assumptions'!$E$26","=K5*'Assumptions'!$E$26"],
  ["=IFERROR(G11/G5,0)","=IFERROR(H11/H5,0)","=IFERROR(I11/I5,0)","=IFERROR(J11/J5,0)","=IFERROR(K11/K5,0)"],
  ["=G9-G11","=H9-H11","=I9-I11","=J9-J11","=K9-K11"],
  ["=IFERROR(G13/G5,0)","=IFERROR(H13/H5,0)","=IFERROR(I13/I5,0)","=IFERROR(J13/J5,0)","=IFERROR(K13/K5,0)"],
  ["=MAX(G13,0)*'Assumptions'!$B$14","=MAX(H13,0)*'Assumptions'!$B$14","=MAX(I13,0)*'Assumptions'!$B$14","=MAX(J13,0)*'Assumptions'!$B$14","=MAX(K13,0)*'Assumptions'!$B$14"],
  ["=G13-G15","=H13-H15","=I13-I15","=J13-J15","=K13-K15"],
  ["=G5*'Assumptions'!$E$27","=H5*'Assumptions'!$E$27","=I5*'Assumptions'!$E$27","=J5*'Assumptions'!$E$27","=K5*'Assumptions'!$E$27"],
  ["=IFERROR(G17/G5,0)","=IFERROR(H17/H5,0)","=IFERROR(I17/I5,0)","=IFERROR(J17/J5,0)","=IFERROR(K17/K5,0)"],
  ["=G5*'Assumptions'!$E$28","=H5*'Assumptions'!$E$28","=I5*'Assumptions'!$E$28","=J5*'Assumptions'!$E$28","=K5*'Assumptions'!$E$28"],
  ["=G19-F19","=H19-G19","=I19-H19","=J19-I19","=K19-J19"],
  ["=G16+G11-G17-G20","=H16+H11-H17-H20","=I16+I11-I17-I20","=J16+J11-J17-J20","=K16+K11-K17-K20"],
  ["=IFERROR(G21/G5,0)","=IFERROR(H21/H5,0)","=IFERROR(I21/I5,0)","=IFERROR(J21/J5,0)","=IFERROR(K21/K5,0)"],
  ["=IFERROR((K5/F5)^(1/5)-1,0)","=IFERROR((K5/F5)^(1/5)-1,0)","=IFERROR((K5/F5)^(1/5)-1,0)","=IFERROR((K5/F5)^(1/5)-1,0)","=IFERROR((K5/F5)^(1/5)-1,0)"],
];
formulas(sheets.model.getRange("G5:K23"));
sheets.model.getRange("K24").formulas = [["=K10"]];
formulas(sheets.model.getRange("K24"));
sheets.model.getRange("B5:K5,B7:K7,B9:K9,B11:K11,B13:K13,B15:K17,B19:K20,B22:K22").setNumberFormat(moneyFmt);
sheets.model.getRange("B6:K6,B8:K8,B10:K10,B12:K12,B14:K14,B18:K18,B21:K21,B23:K24").setNumberFormat(pctFmt);
sheets.model.getRange("A4:K24").format.borders = { preset: "inside", style: "thin", color: fmt.lightBorder };
total(sheets.model.getRange("A22:K22"));
total(sheets.model.getRange("A24:K24"));
sheets.model.freezePanes.freezeRows(4);
sheets.model.freezePanes.freezeColumns(1);

// Valuation
title(sheets.valuation, "A1:G1", "DCF Valuation");
widths(sheets.valuation, [["A", 32], ["B", 16], ["C", 16], ["D", 16], ["E", 16], ["F", 16], ["G", 16]]);
section(sheets.valuation, "A3:G3", "DCF Bridge");
sheets.valuation.getRange("A4:G27").values = [
  ["Metric", "Output", "2027E", "2028E", "2029E", "2030E", "2031E"],
  ["Unlevered FCF", null, null, null, null, null, null],
  ["Discount Period", null, 1, 2, 3, 4, 5],
  ["Discount Factor", null, null, null, null, null, null],
  ["PV of FCF", null, null, null, null, null, null],
  ["Terminal FCF", null, null, null, null, null, null],
  ["Terminal Value", null, null, null, null, null, null],
  ["PV of Terminal Value", null, null, null, null, null, null],
  ["", null, null, null, null, null, null],
  ["PV of Forecast FCF", null, null, null, null, null, null],
  ["PV of Terminal Value", null, null, null, null, null, null],
  ["Enterprise Value", null, null, null, null, null, null],
  ["Less: Net Debt / (Cash)", null, null, null, null, null, null],
  ["Equity Value", null, null, null, null, null, null],
  ["Diluted Shares", null, null, null, null, null, null],
  ["Implied Share Price", null, null, null, null, null, null],
  ["", null, null, null, null, null, null],
  ["Exit EV/EBITDA Cross-Check", null, null, null, null, null, null],
  ["Terminal EBITDA", null, null, null, null, null, null],
  ["Implied Exit Multiple", null, null, null, null, null, null],
  ["Cross-Check EV @ Assumption", null, null, null, null, null, null],
  ["Cross-Check Delta", null, null, null, null, null, null],
  ["Valuation method", "Gordon Growth DCF", null, null, null, null, null],
  ["Timing convention", "Year-end discounting", null, null, null, null, null],
];
header(sheets.valuation.getRange("A4:G4"));
sheets.valuation.getRange("C5:G5").formulas = [["='Forecast Model'!G21","='Forecast Model'!H21","='Forecast Model'!I21","='Forecast Model'!J21","='Forecast Model'!K21"]];
sheets.valuation.getRange("C7:G7").formulas = [["=1/(1+'Assumptions'!$B$12)^C6","=1/(1+'Assumptions'!$B$12)^D6","=1/(1+'Assumptions'!$B$12)^E6","=1/(1+'Assumptions'!$B$12)^F6","=1/(1+'Assumptions'!$B$12)^G6"]];
sheets.valuation.getRange("C8:G8").formulas = [["=C5*C7","=D5*D7","=E5*E7","=F5*F7","=G5*G7"]];
sheets.valuation.getRange("G9:G11").formulas = [["=G5*(1+'Assumptions'!$B$13)"],["=G9/('Assumptions'!$B$12-'Assumptions'!$B$13)"],["=G10*G7"]];
sheets.valuation.getRange("B13:B25").formulas = [
  ["=SUM(C8:G8)"],
  ["=G11"],
  ["=B13+B14"],
  ["='Assumptions'!B15"],
  ["=B15-B16"],
  ["='Assumptions'!B16"],
  ["=IFERROR(B17/B18,0)"],
  [""],
  [""],
  ["='Forecast Model'!K9"],
  ["=IFERROR(B15/B22,0)"],
  ["=B22*'Assumptions'!B19"],
  ["=B24-B15"],
];
formulas(sheets.valuation.getRange("B13:B25"));
formulas(sheets.valuation.getRange("C5:G11"), fmt.linkGreen);
sheets.valuation.getRange("B13:B17,C5:G5,C8:G11,B22:B25").setNumberFormat(moneyFmt);
sheets.valuation.getRange("B18").setNumberFormat(countFmt);
sheets.valuation.getRange("B19").setNumberFormat(perShareFmt);
sheets.valuation.getRange("C6:G7").setNumberFormat("0.00");
sheets.valuation.getRange("B23").setNumberFormat(multFmt);
sheets.valuation.getRange("A4:G27").format.borders = { preset: "inside", style: "thin", color: fmt.lightBorder };
total(sheets.valuation.getRange("A15:B15"));
total(sheets.valuation.getRange("A17:B17"));
total(sheets.valuation.getRange("A19:B19"));
total(sheets.valuation.getRange("A25:B25"));

// Sensitivity
title(sheets.sensitivity, "A1:H1", "Sensitivity Analysis");
widths(sheets.sensitivity, [["A", 22], ["B", 14], ["C", 14], ["D", 14], ["E", 14], ["F", 14], ["G", 14], ["H", 20]]);
section(sheets.sensitivity, "A3:H3", "Implied Share Price Sensitivity");
sheets.sensitivity.getRange("A4:F10").values = [
  ["WACC \\ Terminal Growth", 0.015, 0.020, 0.025, 0.030, 0.035],
  [0.080, null, null, null, null, null],
  [0.090, null, null, null, null, null],
  [0.095, null, null, null, null, null],
  [0.100, null, null, null, null, null],
  [0.110, null, null, null, null, null],
  [0.120, null, null, null, null, null],
];
header(sheets.sensitivity.getRange("A4:F4"));
header(sheets.sensitivity.getRange("A5:A10"));
sheets.sensitivity.getRange("B5:F10").formulas = [
  ["=IF(B$4<$A5,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+B$4)/($A5-B$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(C$4<$A5,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+C$4)/($A5-C$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(D$4<$A5,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+D$4)/($A5-D$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(E$4<$A5,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+E$4)/($A5-E$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(F$4<$A5,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+F$4)/($A5-F$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())"],
  ["=IF(B$4<$A6,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+B$4)/($A6-B$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(C$4<$A6,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+C$4)/($A6-C$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(D$4<$A6,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+D$4)/($A6-D$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(E$4<$A6,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+E$4)/($A6-E$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(F$4<$A6,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+F$4)/($A6-F$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())"],
  ["=IF(B$4<$A7,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+B$4)/($A7-B$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(C$4<$A7,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+C$4)/($A7-C$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(D$4<$A7,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+D$4)/($A7-D$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(E$4<$A7,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+E$4)/($A7-E$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(F$4<$A7,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+F$4)/($A7-F$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())"],
  ["=IF(B$4<$A8,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+B$4)/($A8-B$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(C$4<$A8,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+C$4)/($A8-C$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(D$4<$A8,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+D$4)/($A8-D$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(E$4<$A8,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+E$4)/($A8-E$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(F$4<$A8,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+F$4)/($A8-F$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())"],
  ["=IF(B$4<$A9,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+B$4)/($A9-B$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(C$4<$A9,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+C$4)/($A9-C$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(D$4<$A9,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+D$4)/($A9-D$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(E$4<$A9,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+E$4)/($A9-E$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(F$4<$A9,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+F$4)/($A9-F$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())"],
  ["=IF(B$4<$A10,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+B$4)/($A10-B$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(C$4<$A10,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+C$4)/($A10-C$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(D$4<$A10,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+D$4)/($A10-D$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(E$4<$A10,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+E$4)/($A10-E$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())","=IF(F$4<$A10,((SUM('Valuation'!$C$8:$G$8)+'Valuation'!$G$5*(1+F$4)/($A10-F$4)*'Valuation'!$G$7)-'Assumptions'!$B$15)/'Assumptions'!$B$16,NA())"],
];
sheets.sensitivity.getRange("A5:A10,A4:F4").setNumberFormat(pctFmt);
sheets.sensitivity.getRange("B5:F10").setNumberFormat(perShareFmt);
sheets.sensitivity.getRange("A4:F10").format.borders = { preset: "inside", style: "thin", color: fmt.lightBorder };
sheets.sensitivity.getRange("B5:F10").conditionalFormats.add("colorScale", {
  criteria: [
    { type: "lowestValue", color: "#FCE4D6" },
    { type: "percentile", value: 50, color: "#FFF2CC" },
    { type: "highestValue", color: "#D9EAD3" },
  ],
});
note(sheets.sensitivity.getRange("A12:H13"));
sheets.sensitivity.getRange("A12:H13").values = [["Sensitivity cells are formula-driven and recalculate the DCF mechanics using each WACC / terminal growth pair. Primary valuation remains on the Valuation tab.", null, null, null, null, null, null, null], [null, null, null, null, null, null, null, null]];

// Checks
title(sheets.checks, "A1:H1", "Model Checks");
widths(sheets.checks, [["A", 30], ["B", 18], ["C", 18], ["D", 16], ["E", 16], ["F", 16], ["G", 44], ["H", 16]]);
section(sheets.checks, "A3:H3", "Summary");
sheets.checks.getRange("A4:H4").values = [["Overall Model Status", null, null, null, null, null, null, null]];
sheets.checks.getRange("F4").formulas = [["=IF(COUNTIF(F7:F13,\"Review\")=0,\"OK\",\"Review\")"]];
total(sheets.checks.getRange("A4:F4"));
sheets.checks.getRange("F4").conditionalFormats.add("containsText", { text: "OK", format: { fill: { color: fmt.okFill }, font: { bold: true, color: "#166534" } } });
sheets.checks.getRange("F4").conditionalFormats.add("containsText", { text: "Review", format: { fill: { color: "#9A3412" }, font: { bold: true, color: "#FFFFFF" } } });
section(sheets.checks, "A6:H6", "Detailed Checks");
sheets.checks.getRange("A7:H13").values = [
  ["Check", "Actual", "Expected", "Difference", "Tolerance", "Status", "Notes", "Owner"],
  ["Historical revenue present", null, 1, null, 0, null, "Replace placeholder data with reported financials.", "User"],
  ["WACC > Terminal Growth", null, 1, null, 0, null, "Required for Gordon Growth terminal value.", "User"],
  ["Shares outstanding positive", null, 1, null, 0, null, "Needed for implied share price.", "User"],
  ["FCF ties to components", null, 0, null, 1, null, "Forecast FCF = NOPAT + D&A - Capex - Change in NWC.", "Model"],
  ["Enterprise value bridge", null, 0, null, 1, null, "EV should equal PV forecast FCF + PV terminal value.", "Model"],
  ["Source data mode reviewed", null, 1, null, 0, null, "Review until placeholder mode is replaced or accepted.", "User"],
];
header(sheets.checks.getRange("A7:H7"));
sheets.checks.getRange("B8:B13").formulas = [
  ["=IF(SUM('Historical Financials'!B7:F7)>0,1,0)"],
  ["=IF('Assumptions'!B12>'Assumptions'!B13,1,0)"],
  ["=IF('Assumptions'!B16>0,1,0)"],
  ["='Forecast Model'!K21-('Forecast Model'!K16+'Forecast Model'!K11-'Forecast Model'!K17-'Forecast Model'!K20)"],
  ["='Valuation'!B15-('Valuation'!B13+'Valuation'!B14)"],
  ["=IF('Assumptions'!B18=\"Illustrative placeholder\",0,1)"],
];
sheets.checks.getRange("D8:D13").formulas = [
  ["=B8-C8"],
  ["=B9-C9"],
  ["=B10-C10"],
  ["=B11-C11"],
  ["=B12-C12"],
  ["=B13-C13"],
];
sheets.checks.getRange("F8:F13").formulas = [
  ["=IF(ABS(D8)<=E8,\"OK\",\"Review\")"],
  ["=IF(ABS(D9)<=E9,\"OK\",\"Review\")"],
  ["=IF(ABS(D10)<=E10,\"OK\",\"Review\")"],
  ["=IF(ABS(D11)<=E11,\"OK\",\"Review\")"],
  ["=IF(ABS(D12)<=E12,\"OK\",\"Review\")"],
  ["=IF(ABS(D13)<=E13,\"OK\",\"Review\")"],
];
sheets.checks.getRange("A7:H13").format.borders = { preset: "inside", style: "thin", color: fmt.lightBorder };
sheets.checks.getRange("B8:E13").setNumberFormat("#,##0.0;[Red](#,##0.0);-");
sheets.checks.getRange("F8:F13").conditionalFormats.add("containsText", { text: "OK", format: { fill: { color: fmt.okFill }, font: { bold: true, color: "#166534" } } });
sheets.checks.getRange("F8:F13").conditionalFormats.add("containsText", { text: "Review", format: { fill: { color: "#9A3412" }, font: { bold: true, color: "#FFFFFF" } } });
sheets.checks.getRange("G8:G13").format.wrapText = true;

// Sources Audit
title(sheets.sources, "A1:G1", "Sources Audit");
widths(sheets.sources, [["A", 18], ["B", 28], ["C", 18], ["D", 18], ["E", 26], ["F", 50], ["G", 42]]);
section(sheets.sources, "A3:G3", "Source Register");
sheets.sources.getRange("A4:G10").values = [
  ["Source ID", "Item", "Period / As-of", "Value / Units", "Source Name", "Plain-text URL", "Notes"],
  ["SRC-001", "Historical financial statements", "2022A-2026A", "$mm", "Company filings / annual reports", "https://example.com/replace-with-company-filing", "Placeholder source. Replace with actual filing, export, or data vendor URL."],
  ["SRC-002", "Share count", "Latest filing", "mm shares", "Company filing", "https://example.com/replace-with-share-count-source", "Tie to weighted-average diluted shares or current diluted count."],
  ["SRC-003", "Net debt / cash", "Valuation date", "$mm", "Balance sheet / market data", "https://example.com/replace-with-net-debt-source", "Use latest cash, debt, and non-operating assets/liabilities."],
  ["SRC-004", "WACC / terminal assumptions", "Valuation date", "% / x", "Analyst inputs", "https://example.com/replace-with-wacc-support", "Document cost of capital, peer multiples, and terminal growth rationale."],
  ["", "", "", "", "", "", ""],
  ["Audit note", "Do not rely on placeholder outputs for investment, lending, or board decisions until source rows and historical values are replaced.", "", "", "", "", ""],
];
header(sheets.sources.getRange("A4:G4"));
sheets.sources.getRange("A4:G10").format.borders = { preset: "inside", style: "thin", color: fmt.lightBorder };
sheets.sources.getRange("F5:G10").format.wrapText = true;
note(sheets.sources.getRange("A10:G10"));
sheets.sources.freezePanes.freezeRows(4);

// Charts on Cover: revenue and FCF
sheets.cover.getRange("J23:L33").values = [
  ["Year", "Revenue", "Unlevered FCF"],
  ["2022A", null, null],
  ["2023A", null, null],
  ["2024A", null, null],
  ["2025A", null, null],
  ["2026A", null, null],
  ["2027E", null, null],
  ["2028E", null, null],
  ["2029E", null, null],
  ["2030E", null, null],
  ["2031E", null, null],
];
sheets.cover.getRange("M23:N33").values = [
  ["Year", "Unlevered FCF"],
  ["2022A", null],
  ["2023A", null],
  ["2024A", null],
  ["2025A", null],
  ["2026A", null],
  ["2027E", null],
  ["2028E", null],
  ["2029E", null],
  ["2030E", null],
  ["2031E", null],
];
header(sheets.cover.getRange("J23:L23"));
header(sheets.cover.getRange("M23:N23"));
sheets.cover.getRange("K24:L33").formulas = [
  ["='Forecast Model'!B5","='Forecast Model'!B21"],
  ["='Forecast Model'!C5","='Forecast Model'!C21"],
  ["='Forecast Model'!D5","='Forecast Model'!D21"],
  ["='Forecast Model'!E5","='Forecast Model'!E21"],
  ["='Forecast Model'!F5","='Forecast Model'!F21"],
  ["='Forecast Model'!G5","='Forecast Model'!G21"],
  ["='Forecast Model'!H5","='Forecast Model'!H21"],
  ["='Forecast Model'!I5","='Forecast Model'!I21"],
  ["='Forecast Model'!J5","='Forecast Model'!J21"],
  ["='Forecast Model'!K5","='Forecast Model'!K21"],
];
sheets.cover.getRange("N24:N33").formulas = [
  ["='Forecast Model'!B21"],
  ["='Forecast Model'!C21"],
  ["='Forecast Model'!D21"],
  ["='Forecast Model'!E21"],
  ["='Forecast Model'!F21"],
  ["='Forecast Model'!G21"],
  ["='Forecast Model'!H21"],
  ["='Forecast Model'!I21"],
  ["='Forecast Model'!J21"],
  ["='Forecast Model'!K21"],
];
sheets.cover.getRange("K24:L33").setNumberFormat(moneyFmt);
sheets.cover.getRange("N24:N33").setNumberFormat(moneyFmt);
formulas(sheets.cover.getRange("K24:L33"), fmt.linkGreen);
formulas(sheets.cover.getRange("N24:N33"), fmt.linkGreen);

const revenueChart = sheets.cover.charts.add("line", sheets.cover.getRange("J23:K33"));
revenueChart.setPosition("A23", "D38");
revenueChart.title = "Revenue Trend ($mm)";
revenueChart.hasLegend = false;
revenueChart.yAxis = { numberFormat: '$#,##0' };

const fcfChart = sheets.cover.charts.add("column", sheets.cover.getRange("M23:N33"));
fcfChart.setPosition("E23", "H38");
fcfChart.title = "Unlevered FCF ($mm)";
fcfChart.hasLegend = false;
fcfChart.yAxis = { numberFormat: '$#,##0' };

// General final formatting.
for (const sheet of Object.values(sheets)) {
  const used = sheet.getUsedRange();
  used.format.wrapText = false;
  sheet.getRange("A:A").format.horizontalAlignment = "left";
}
sheets.cover.getRange("A3:H3,A20:H20").format.wrapText = true;
sheets.assumptions.getRange("F4:F28").format.wrapText = true;
sheets.historical.getRange("H7:H21").format.wrapText = true;
sheets.sources.getRange("F5:G10").format.wrapText = true;

// Render every sheet for visual QA and save previews.
await fs.mkdir(outputDir, { recursive: true });
for (const name of Object.keys(sheets)) {
  const sheet = sheets[name];
  const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `preview_${name}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const check = await workbook.inspect({
  kind: "table",
  sheetId: "Cover",
  range: "A1:H20",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 8,
  maxChars: 6000,
});
console.log("COVER_CHECK");
console.log(check.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 6000,
});
console.log("ERROR_SCAN");
console.log(errors.ndjson);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(path.join(outputDir, "company_financial_analysis.xlsx"));
console.log(path.join(outputDir, "company_financial_analysis.xlsx"));
