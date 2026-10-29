import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Handlowiec/Desktop/Puszki i złączki/MOREK - CENNIK.xlsx";
const outputPath = "outputs/bloki_rozdzielcze_20261009/morek_records.json";
const wanted = new Set(["MAB2124A18", "MAB2125A18", "MAB2126A18", "MAB4124A18"]);

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const matches = await workbook.inspect({
  kind: "match",
  searchTerm: "MAB2124A18|MAB2125A18|MAB2126A18|MAB4124A18",
  options: { useRegex: true, maxResults: 50 },
  maxChars: 10000,
});
console.log("MATCHES");
console.log(matches.ndjson);
const sheet = workbook.worksheets.getItem("CENNIK BAZOWY");
const values = sheet.getUsedRange(true).values;
const headers = values[13].map((value) => String(value ?? "").replace(/\s+/g, " ").trim());
const records = [];
for (let i = 14; i < values.length; i += 1) {
  const code = String(values[i][1] ?? "").trim();
  if (!wanted.has(code)) continue;
  const record = Object.fromEntries(headers.map((header, index) => [header || `column_${index + 1}`, values[i][index] ?? null]));
  records.push({ sourceRow: i + 1, ...record });
}
await fs.writeFile(outputPath, JSON.stringify(records, null, 2), "utf8");
console.log(JSON.stringify({ outputPath, count: records.length, missing: [...wanted].filter((code) => !records.some((record) => record["Nr katalogowy"] === code)), records }, null, 2));
