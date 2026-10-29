import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Handlowiec/Downloads/zlaczki-opisy-z-woocomerce.xlsx";
const previewPath = "outputs/zlaczki_seo_20261009/woo_descriptions_before.png";

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const summary = await workbook.inspect({
  kind: "workbook,sheet,table,region",
  maxChars: 40000,
  tableMaxRows: 30,
  tableMaxCols: 30,
  tableMaxCellChars: 180,
});
console.log(summary.ndjson);

const sheet = workbook.worksheets.items[0];
const preview = await workbook.render({
  sheetName: sheet.name,
  range: "A1:Z12",
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, Buffer.from(await preview.arrayBuffer()));
console.log(JSON.stringify({ sheetName: sheet.name, previewPath }, null, 2));
