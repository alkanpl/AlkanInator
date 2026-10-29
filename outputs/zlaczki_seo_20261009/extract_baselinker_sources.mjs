import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath = "outputs/zlaczki_seo_20261009/source_records.json";
const morekPath = "C:/Users/Handlowiec/Desktop/Puszki i złączki/MOREK - CENNIK.xlsx";
const outputPath = "outputs/zlaczki_seo_20261009/baselinker_source_data.json";

const records = JSON.parse(await fs.readFile(sourcePath, "utf8"));
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(morekPath));
const sheet = workbook.worksheets.getItem("CENNIK BAZOWY");
const rows = sheet.getUsedRange(true).values;

const targets = new Set(
  records
    .filter((record) => record.supplier === "Morek")
    .map((record) => record.sku.replace(/\/MOR$/i, "")),
);

const selected = [];
for (let index = 0; index < rows.length; index += 1) {
  const row = rows[index];
  const code = String(row[1] ?? "").trim();
  if (!targets.has(code)) continue;
  selected.push({
    sourceRow: index + 1,
    code,
    name: String(row[2] ?? "").trim(),
    manufacturerDescription: String(row[3] ?? "").trim(),
    ean: String(row[12] ?? "").trim(),
    weightKg: String(row[14] ?? "").trim(),
    lengthM: String(row[15] ?? "").trim(),
    widthM: String(row[16] ?? "").trim(),
    heightM: String(row[17] ?? "").trim(),
    packageQuantity: String(row[19] ?? "").trim(),
    group: String(row[21] ?? "").trim(),
    etim: String(row[22] ?? "").trim(),
    imageUrl: String(row[24] ?? "").trim(),
  });
}

await fs.writeFile(outputPath, JSON.stringify(selected, null, 2), "utf8");
console.log(JSON.stringify({
  outputPath,
  count: selected.length,
  missingImages: selected.filter((row) => !row.imageUrl).map((row) => row.code),
  rows: selected,
}, null, 2));
