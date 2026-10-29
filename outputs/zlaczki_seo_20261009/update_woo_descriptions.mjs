import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Handlowiec/Downloads/zlaczki-opisy-z-woocomerce.xlsx";
const outputPath = "C:/Users/Handlowiec/Desktop/AlkanInatorApp/outputs/zlaczki_seo_20261009/zlaczki-opisy-z-woocomerce_uzupelnione.xlsx";
const descriptionsDir = "C:/Users/Handlowiec/Desktop/AlkanInatorApp/reports/descriptions/zlaczki_20261009";
const previewPath = "C:/Users/Handlowiec/Desktop/AlkanInatorApp/outputs/zlaczki_seo_20261009/woo_descriptions_after.png";

const normalize = (value) => String(value ?? "").trim().toUpperCase();
const stripSupplier = (value) => normalize(value).replace(/\/(SIM|MOR)$/i, "");
const normalizeHtml = (value) => String(value ?? "").replace(/\r\n/g, "\n").trimEnd();

const files = (await fs.readdir(descriptionsDir)).filter((name) => name.endsWith(".html")).sort();
const descriptions = new Map();
for (const file of files) {
  const sku = normalize(file.replace(/\.html$/i, "").replace(/_(SIM|MOR)$/i, "/$1"));
  const html = await fs.readFile(`${descriptionsDir}/${file}`, "utf8");
  if (!html.trim()) throw new Error(`Pusty opis: ${file}`);
  if (descriptions.has(stripSupplier(sku))) throw new Error(`Powielony SKU w opisach: ${sku}`);
  descriptions.set(stripSupplier(sku), { sku, file, html });
}
if (descriptions.size !== 37) throw new Error(`Oczekiwano 37 opisów, znaleziono ${descriptions.size}`);

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = workbook.worksheets.getItem("Worksheet");
const originalValues = sheet.getRange("A1:H301").values;
const headers = originalValues[0].map((value) => String(value ?? "").trim());
const skuIndex = headers.indexOf("Sku");
const contentIndex = headers.indexOf("Content");
if (skuIndex < 0 || contentIndex < 0) throw new Error("Brak kolumn Sku lub Content");

const matches = [];
const matchedKeys = new Set();
for (let i = 1; i < originalValues.length; i += 1) {
  const workbookSku = normalize(originalValues[i][skuIndex]);
  const key = stripSupplier(workbookSku);
  const source = descriptions.get(key);
  if (!source) continue;
  if (matchedKeys.has(key)) throw new Error(`Powielony SKU w arkuszu: ${workbookSku}`);
  matchedKeys.add(key);
  matches.push({
    excelRow: i + 1,
    workbookSku,
    sourceSku: source.sku,
    file: source.file,
    html: source.html,
  });
}

const unmatched = [...descriptions.values()]
  .filter((item) => !matchedKeys.has(stripSupplier(item.sku)))
  .map((item) => item.sku);
if (matches.length !== 36) throw new Error(`Oczekiwano 36 dopasowań w skoroszycie, znaleziono ${matches.length}`);
if (unmatched.length !== 1 || unmatched[0] !== "MAD1007N15/MOR") {
  throw new Error(`Nieoczekiwane niedopasowania: ${unmatched.join(", ")}`);
}

for (const match of matches) {
  sheet.getRange(`G${match.excelRow}`).values = [[match.html]];
}

workbook.recalculate();
const updatedRange = await workbook.inspect({
  kind: "table",
  range: "Worksheet!A264:H301",
  include: "values,formulas",
  tableMaxRows: 40,
  tableMaxCols: 8,
  tableMaxCellChars: 100,
  maxChars: 14000,
});
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  maxChars: 3000,
});
const preview = await workbook.render({
  sheetName: "Worksheet",
  range: "A264:H271",
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

const reopened = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const reopenedValues = reopened.worksheets.getItem("Worksheet").getRange("A1:H301").values;
let changedContentCells = 0;
let changedOutsideTargetContent = 0;
let changedOtherCells = 0;
for (let r = 0; r < originalValues.length; r += 1) {
  for (let c = 0; c < originalValues[r].length; c += 1) {
    const before = originalValues[r][c] ?? null;
    const after = reopenedValues[r][c] ?? null;
    if (String(before) === String(after)) continue;
    if (c === contentIndex && matchedKeys.has(stripSupplier(originalValues[r][skuIndex]))) {
      changedContentCells += 1;
    } else if (c === contentIndex) {
      changedOutsideTargetContent += 1;
    } else {
      changedOtherCells += 1;
    }
  }
}

const missingAfterReopen = matches.filter((match) => {
  const value = reopenedValues[match.excelRow - 1][contentIndex];
  return normalizeHtml(value) !== normalizeHtml(match.html);
});
if (changedContentCells !== 36 || changedOutsideTargetContent !== 0 || changedOtherCells !== 0 || missingAfterReopen.length) {
  throw new Error(JSON.stringify({ changedContentCells, changedOutsideTargetContent, changedOtherCells, missingAfterReopen }, null, 2));
}

console.log("UPDATED_RANGE");
console.log(updatedRange.ndjson);
console.log("FORMULA_ERRORS");
console.log(errors.ndjson);
console.log(JSON.stringify({
  inputPath,
  outputPath,
  previewPath,
  preparedDescriptions: descriptions.size,
  updatedProducts: matches.length,
  unmatched,
  changedContentCells,
  changedOutsideTargetContent,
  changedOtherCells,
  missingAfterReopen: missingAfterReopen.length,
  firstUpdatedRow: Math.min(...matches.map((match) => match.excelRow)),
  lastUpdatedRow: Math.max(...matches.map((match) => match.excelRow)),
}, null, 2));
