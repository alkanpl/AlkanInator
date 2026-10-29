import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const path = "C:/Users/Handlowiec/Desktop/AlkanInatorApp/outputs/zlaczki_seo_20261009/zlaczki-opisy-z-woocomerce_uzupelnione.xlsx";
const pattern = /(opakowani(?:e|a|u|em)|opakowanie zbiorcze|minimaln(?:e|a|ą) iloś|minimum zamówienia|sprzedawan(?:e|y|a) po|pakiet|multipak|karton|\bMOQ\b|\b\d+\s*szt\.?\b|\bsztuk(?:a|i|ę|ą|om|ach)?\b)/giu;

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const sheet = workbook.worksheets.getItem("Worksheet");
const values = sheet.getRange("A266:H301").values;
const hits = [];
let populatedDescriptions = 0;
for (let i = 0; i < values.length; i += 1) {
  const sku = String(values[i][1] ?? "");
  const content = String(values[i][6] ?? "");
  const producer = String(values[i][7] ?? "");
  if (content.trim()) populatedDescriptions += 1;
  for (const [field, text] of [["Content", content], ["Atrybut produktu: Producent", producer]]) {
    const matches = [...text.matchAll(pattern)].map((match) => match[0]);
    if (matches.length) hits.push({ excelRow: i + 266, sku, field, matches });
  }
}

console.log(JSON.stringify({
  checkedRows: values.length,
  populatedDescriptions,
  wholesaleHits: hits,
}, null, 2));
