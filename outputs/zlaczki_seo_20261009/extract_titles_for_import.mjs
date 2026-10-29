import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "outputs/zlaczki_seo_20261009/ZLACZKI_tytuly_SEO.xlsx";
const outputPath = "outputs/zlaczki_seo_20261009/seo_titles_for_import.json";

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));
const sheet = workbook.worksheets.getItem("Tytuły SEO");
const values = sheet.getRange("A5:G42").values;
const headers = values[0].map((value) => String(value ?? "").trim());
const rows = values.slice(1).map((row) => Object.fromEntries(
  headers.map((header, index) => [header, row[index] ?? ""]),
));

await fs.writeFile(outputPath, JSON.stringify(rows, null, 2), "utf8");
console.log(JSON.stringify({
  outputPath,
  count: rows.length,
  missingTitles: rows.filter((row) => !row["Tytuł SEO"]).map((row) => row.SKU),
}, null, 2));
