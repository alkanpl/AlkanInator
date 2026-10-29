import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Handlowiec/Desktop/Puszki i złączki/Bloki rozdzielcze do wystawienia.xlsx";
const previewPath = "outputs/bloki_rozdzielcze_20261009/source_preview.png";

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const summary = await workbook.inspect({
  kind: "workbook,sheet,table,region",
  maxChars: 24000,
  tableMaxRows: 20,
  tableMaxCols: 20,
  tableMaxCellChars: 220,
});
console.log(summary.ndjson);

const sheet = workbook.worksheets.items[0];
const preview = await workbook.render({
  sheetName: sheet.name,
  range: sheet.getUsedRange(true).address,
  scale: 1.5,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
console.log(JSON.stringify({ sheet: sheet.name, usedRange: sheet.getUsedRange(true).address, previewPath }, null, 2));
