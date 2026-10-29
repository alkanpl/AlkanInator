import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const files = [
  "C:/Users/Handlowiec/Desktop/Puszki i złączki/ZŁĄCZKI - lista produktów do wystawienia.xlsx",
  "C:/Users/Handlowiec/Desktop/Puszki i złączki/CENNIK SIMET.xlsx",
  "C:/Users/Handlowiec/Desktop/Puszki i złączki/MOREK - CENNIK.xlsx",
];

for (const file of files) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(file));
  const summary = await workbook.inspect({
    kind: "workbook,sheet,table,region",
    maxChars: 24000,
    tableMaxRows: 18,
    tableMaxCols: 18,
    tableMaxCellChars: 120,
  });
  console.log(`\n===== ${file} =====`);
  console.log(summary.ndjson);
}
