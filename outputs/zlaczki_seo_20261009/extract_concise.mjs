import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const outPath = "outputs/zlaczki_seo_20261009/source_records.json";
const listPath = "C:/Users/Handlowiec/Desktop/Puszki i złączki/ZŁĄCZKI - lista produktów do wystawienia.xlsx";
const simetPath = "C:/Users/Handlowiec/Desktop/Puszki i złączki/CENNIK SIMET.xlsx";
const morekPath = "C:/Users/Handlowiec/Desktop/Puszki i złączki/MOREK - CENNIK.xlsx";

const listWb = await SpreadsheetFile.importXlsx(await FileBlob.load(listPath));
const listRows = listWb.worksheets.getItemAt(0).getUsedRange(true).values;

const simetWb = await SpreadsheetFile.importXlsx(await FileBlob.load(simetPath));
const simetRows = simetWb.worksheets.getItem("CENNIK").getUsedRange(true).values;
const simetByCode = new Map();
for (let i = 0; i < simetRows.length; i += 1) {
  const row = simetRows[i];
  const code = String(row[7] ?? "").trim();
  if (code) simetByCode.set(code, { sourceRow: i + 1, row });
}

const morekWb = await SpreadsheetFile.importXlsx(await FileBlob.load(morekPath));
const morekRows = morekWb.worksheets.getItem("CENNIK BAZOWY").getUsedRange(true).values;
const morekByCode = new Map();
for (let i = 0; i < morekRows.length; i += 1) {
  const row = morekRows[i];
  const code = String(row[1] ?? "").trim();
  if (code) morekByCode.set(code, { sourceRow: i + 1, row });
}

const records = listRows.map((row, index) => {
  const sku = String(row[0] ?? "").trim();
  const originalTitle = String(row[1] ?? "").trim();
  if (sku.endsWith("/SIM")) {
    const code = sku.slice(0, -4);
    const match = simetByCode.get(code);
    return {
      listRow: index + 1,
      sku,
      originalTitle,
      supplier: "Simet",
      sourceSheet: "CENNIK",
      sourceRow: match?.sourceRow ?? null,
      sourceName: String(match?.row[9] ?? "").trim(),
      sourceShortName: String(match?.row[8] ?? "").trim(),
      sourceDetail: String(match?.row[10] ?? "").trim(),
    };
  }
  const code = sku.endsWith("/MOR") ? sku.slice(0, -4) : sku;
  const match = morekByCode.get(code);
  return {
    listRow: index + 1,
    sku,
    originalTitle,
    supplier: "Morek",
    sourceSheet: "CENNIK BAZOWY",
    sourceRow: match?.sourceRow ?? null,
    sourceName: String(match?.row[2] ?? "").trim(),
    sourceGroup: String(match?.row[21] ?? "").trim(),
    ean: String(match?.row[12] ?? "").trim(),
  };
});

await fs.writeFile(outPath, JSON.stringify(records, null, 2), "utf8");
console.log(JSON.stringify({
  count: records.length,
  missingMatches: records.filter((item) => !item.sourceName).map((item) => item.sku),
  records,
}, null, 2));
