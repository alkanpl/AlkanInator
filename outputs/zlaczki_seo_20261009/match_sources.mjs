import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const paths = {
  list: "C:/Users/Handlowiec/Desktop/Puszki i złączki/ZŁĄCZKI - lista produktów do wystawienia.xlsx",
  simet: "C:/Users/Handlowiec/Desktop/Puszki i złączki/CENNIK SIMET.xlsx",
  morek: "C:/Users/Handlowiec/Desktop/Puszki i złączki/MOREK - CENNIK.xlsx",
};

const normalize = (value) => String(value ?? "").trim().toUpperCase();

const listWb = await SpreadsheetFile.importXlsx(await FileBlob.load(paths.list));
const listSheet = listWb.worksheets.getItemAt(0);
const listRows = listSheet.getUsedRange(true).values.map((row, index) => ({
  row: index + 1,
  sku: String(row[0] ?? "").trim(),
  originalTitle: String(row[1] ?? "").trim(),
}));

console.log("LIST_ROWS");
console.log(JSON.stringify(listRows, null, 2));

for (const [supplier, path, suffix] of [
  ["SIMET", paths.simet, "/SIM"],
  ["MOREK", paths.morek, "/MOR"],
]) {
  const targets = listRows
    .filter((item) => normalize(item.sku).endsWith(suffix))
    .map((item) => normalize(item.sku).slice(0, -suffix.length));
  const targetSet = new Set(targets);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
  const matches = [];
  for (const sheet of workbook.worksheets.items) {
    const used = sheet.getUsedRange(true);
    const values = used?.values ?? [];
    for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      const hitCodes = [...new Set(row.map(normalize).filter((cell) => targetSet.has(cell)))];
      if (hitCodes.length) {
        matches.push({ sheet: sheet.name, row: rowIndex + 1, hitCodes, values: row });
      }
    }
  }
  const matchedCodes = new Set(matches.flatMap((match) => match.hitCodes));
  console.log(`${supplier}_MATCHES`);
  console.log(JSON.stringify(matches, null, 2));
  console.log(`${supplier}_MISSING`);
  console.log(JSON.stringify(targets.filter((code) => !matchedCodes.has(code)), null, 2));
}
