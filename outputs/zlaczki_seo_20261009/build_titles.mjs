import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workDir = "outputs/zlaczki_seo_20261009";
const recordsPath = `${workDir}/source_records.json`;
const outputPath = `${workDir}/ZLACZKI_tytuly_SEO.xlsx`;
const previewPath = `${workDir}/preview.png`;

const records = JSON.parse(await fs.readFile(recordsPath, "utf8"));
const clean = (value) => String(value ?? "").replace(/\s+/g, " ").trim();
const compactUnits = (value) => clean(value)
  .replace(/mm²/g, "mm2")
  .replace(/1000V AC\/1500V DC/g, "1000VAC/1500VDC");
const codeFromSku = (sku) => sku.replace(/\/(SIM|MOR)$/i, "");
const feminineColor = (color) => ({
  szary: "szara",
  niebieski: "niebieska",
  czerwony: "czerwona",
  "żółto-zielony": "żółto-zielona",
  żółty: "żółta",
  czarny: "czarna",
  pomarańczowy: "pomarańczowa",
  zielony: "zielona",
})[color.toLowerCase()] ?? color;

function buildSeoTitle(record) {
  const code = codeFromSku(record.sku);
  const sourceName = clean(record.sourceName);

  if (record.supplier === "Simet" && sourceName.startsWith("Złączka skrętna")) {
    const short = clean(record.sourceShortName);
    const shortMatch = short.match(/^(\S+)\s+(żółty|czarny)$/i);
    const rangeMatch = record.sourceDetail.match(/od\s+([\d,.]+)mm²\s+do\s+([\d,.]+)mm²/i);
    if (!shortMatch || !rangeMatch) throw new Error(`Nie można rozpoznać danych SIMET: ${record.sku}`);
    const [, model, color] = shortMatch;
    return `Złączka skrętna uniwersalna ${rangeMatch[1]}-${rangeMatch[2]}mm2 ${feminineColor(color)} Simet ${model} ${code}`;
  }

  if (record.supplier === "Simet" && sourceName.startsWith("Listwa zaciskowa")) {
    const short = clean(record.sourceShortName);
    const shortMatch = short.match(/^(\S+)\s+(pomarańczowy)$/i);
    const sectionMatch = record.sourceDetail.match(/([\d,.]+)\s*mm²/i);
    const waysMatch = record.sourceDetail.match(/(\d+)\s+torów/i);
    if (!shortMatch || !sectionMatch || !waysMatch) throw new Error(`Nie można rozpoznać danych SIMET: ${record.sku}`);
    const [, model, color] = shortMatch;
    const section = sectionMatch[1].replace(/,0$/, "");
    return `Listwa zaciskowa gwintowa termoplastyczna ${waysMatch[1]}-torowa ${section}mm2 ${feminineColor(color)} Simet ${model} ${code}`;
  }

  if (record.supplier === "Morek" && /^Złączka OTL/i.test(sourceName)) {
    const match = sourceName.match(/^Złączka\s+(OTL\S+)\s+(szary|niebieski|czerwony|żółto-zielony)\s+(\d+xAl\/Cu)\s+([\d,]+-[\d,]+mm²)\s+(1000V AC\/1500V DC)\s+Zacisk uniwersalny$/i);
    if (!match) throw new Error(`Nie można rozpoznać danych MOREK OTL: ${record.sku} | ${sourceName}`);
    const [, model, color, connections, section, voltage] = match;
    return `Złączka uniwersalna ${model} ${connections} ${compactUnits(section)} ${compactUnits(voltage)} ${feminineColor(color)} Morek ${code}`;
  }

  const listTitles = {
    MAD0012N15: `Listwa zaciskowa neutralna N12 12x16mm2 niebieska Morek ${code}`,
    MAD0052E21: `Listwa zaciskowa N/PE mosiężna BB52 52x16mm2 Morek ${code}`,
    MAD1007G15: `Listwa zaciskowa uziemiająca PE7H 7x16mm2 IP20 zielona Morek ${code}`,
    MAD1007N15: `Listwa zaciskowa neutralna N7H 7x16mm2 IP20 niebieska Morek ${code}`,
  };
  if (listTitles[code]) return listTitles[code];

  throw new Error(`Brak reguły tytułu: ${record.sku}`);
}

const rows = records.map((record) => {
  const seoTitle = buildSeoTitle(record);
  const sourceFile = record.supplier === "Simet" ? "CENNIK SIMET.xlsx" : "MOREK - CENNIK.xlsx";
  return [
    record.sku,
    record.originalTitle,
    seoTitle,
    record.supplier,
    `${sourceFile}, ${record.sourceSheet}, wiersz ${record.sourceRow}`,
    1,
    "OK",
  ];
});

const titles = rows.map((row) => row[2]);
const quality = {
  products: rows.length,
  simet: records.filter((record) => record.supplier === "Simet").length,
  morek: records.filter((record) => record.supplier === "Morek").length,
  uniqueTitles: new Set(titles).size,
  minTitleLength: Math.min(...titles.map((title) => title.length)),
  maxTitleLength: Math.max(...titles.map((title) => title.length)),
  titlesWithQuestionMark: titles.filter((title) => title.includes("?")),
  titlesWithoutSupplier: rows.filter((row) => !row[2].includes(row[3])).map((row) => row[0]),
  titlesWithoutCatalogCode: rows.filter((row) => !row[2].endsWith(codeFromSku(row[0]))).map((row) => row[0]),
};

if (quality.products !== 37 || quality.uniqueTitles !== 37 || quality.titlesWithQuestionMark.length || quality.titlesWithoutSupplier.length || quality.titlesWithoutCatalogCode.length) {
  throw new Error(`Kontrola tytułów nie powiodła się: ${JSON.stringify(quality)}`);
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Tytuły SEO");
sheet.showGridLines = false;
sheet.tabColor = "#1F4E78";

sheet.getRange("A2:G2").merge();
sheet.getRange("A2").values = [["Tytuły SEO produktów: złączki i listwy zaciskowe"]];
sheet.getRange("A3:G3").merge();
sheet.getRange("A3").values = [["Wszystkie pozycje dopasowano do cenników po kodzie katalogowym."]];

const headers = [["SKU", "Tytuł z listy", "Tytuł SEO", "Producent", "Źródło", "Pewność", "Status"]];
sheet.getRange("A5:G5").values = headers;
sheet.getRange(`A6:G${5 + rows.length}`).values = rows;

const table = sheet.tables.add(`A5:G${5 + rows.length}`, true, "SeoTitlesTable");
table.style = "TableStyleMedium2";
table.showBandedColumns = false;
table.showFilterButton = true;

sheet.getRange(`A2:G${5 + rows.length}`).format.font = { name: "Arial", size: 10, color: "#222222" };
sheet.getRange("A2").format.font = { name: "Arial", size: 15, bold: true, color: "#1F1F1F" };
sheet.getRange("A2:G2").format.rowHeight = 25;
sheet.getRange("A3").format.font = { name: "Arial", size: 10, italic: true, color: "#666666" };
sheet.getRange("A3:G3").format.rowHeight = 20;
sheet.getRange("A5:G5").format = {
  fill: "#1F4E78",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};
sheet.getRange("A5:G5").format.rowHeight = 30;

const lastRow = 5 + rows.length;
sheet.getRange(`A6:G${lastRow}`).format.verticalAlignment = "center";
sheet.getRange(`B6:C${lastRow}`).format.wrapText = true;
sheet.getRange(`E6:E${lastRow}`).format.wrapText = true;
sheet.getRange(`A6:A${lastRow}`).format.numberFormat = "@";
sheet.getRange(`F6:F${lastRow}`).format.numberFormat = "0%";
sheet.getRange(`D6:G${lastRow}`).format.horizontalAlignment = "center";
sheet.getRange(`A6:G${lastRow}`).format.rowHeight = 34;

sheet.getRange(`A1:A${lastRow}`).format.columnWidth = 22;
sheet.getRange(`B1:B${lastRow}`).format.columnWidth = 43;
sheet.getRange(`C1:C${lastRow}`).format.columnWidth = 82;
sheet.getRange(`D1:D${lastRow}`).format.columnWidth = 13;
sheet.getRange(`E1:E${lastRow}`).format.columnWidth = 41;
sheet.getRange(`F1:F${lastRow}`).format.columnWidth = 12;
sheet.getRange(`G1:G${lastRow}`).format.columnWidth = 12;

sheet.getRange(`G6:G${lastRow}`).conditionalFormats.addCustom('=G6<>"OK"', {
  fill: "#FCE8E6",
  font: { bold: true, color: "#B91C1C" },
});

sheet.freezePanes.freezeRows(5);
sheet.freezePanes.freezeColumns(1);

workbook.recalculate();

const keyRange = await workbook.inspect({
  kind: "table",
  range: `Tytuły SEO!A1:G${lastRow}`,
  include: "values,formulas",
  tableMaxRows: 45,
  tableMaxCols: 7,
  tableMaxCellChars: 160,
  maxChars: 30000,
});
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});

const preview = await workbook.render({ sheetName: "Tytuły SEO", range: `A1:G${lastRow}`, scale: 1, format: "png" });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

const reopened = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const reopenedCheck = await reopened.inspect({
  kind: "sheet,table,region",
  sheetId: "Tytuły SEO",
  range: `A1:G${lastRow}`,
  tableMaxRows: 8,
  tableMaxCols: 7,
  maxChars: 10000,
});

console.log(JSON.stringify({ outputPath, previewPath, quality }, null, 2));
console.log("KEY_RANGE");
console.log(keyRange.ndjson);
console.log("FORMULA_ERRORS");
console.log(errors.ndjson);
console.log("REOPENED_CHECK");
console.log(reopenedCheck.ndjson);
