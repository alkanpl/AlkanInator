import fs from "node:fs/promises";

const inputPath = "outputs/bloki_rozdzielcze_20261009/BASELINKER_bloki_rozdzielcze_Morek_20261009.csv";
const outputPath = "outputs/bloki_rozdzielcze_20261009/BASELINKER_bloki_rozdzielcze_Morek_z_opisami_20261009.csv";

const descriptionFiles = {
  "MAB2124A18/MOR": "reports/descriptions/bloki_rozdzielcze_20261009/MAB2124A18_MOR.html",
  "MAB2125A18/MOR": "reports/descriptions/bloki_rozdzielcze_20261009/MAB2125A18_MOR.html",
  "MAB2126A18/MOR": "reports/descriptions/bloki_rozdzielcze_20261009/MAB2126A18_MOR.html",
  "MAB4124A18/MOR": "reports/descriptions/bloki_rozdzielcze_20261009/MAB4124A18_MOR.html",
};

function parseSemicolonCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (quoted) {
      if (char === '"' && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        field += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ";") {
      row.push(field);
      field = "";
    } else if (char === "\r" && text[index + 1] === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      index += 1;
    } else if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += char;
    }
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function escapeCsv(value) {
  const text = String(value ?? "");
  return /[;"\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

const source = (await fs.readFile(inputPath, "utf8")).replace(/^\uFEFF/, "");
const table = parseSemicolonCsv(source);
const headers = table[0];
const skuIndex = headers.indexOf("sku");
const descriptionIndex = headers.indexOf("description");

if (skuIndex === -1 || descriptionIndex === -1) {
  throw new Error("Brak kolumny sku lub description w pliku źródłowym.");
}

const descriptions = new Map();
for (const [sku, filePath] of Object.entries(descriptionFiles)) {
  descriptions.set(sku, (await fs.readFile(filePath, "utf8")).trim());
}

const dataRows = table.slice(1).filter((row) => row.some((value) => value !== ""));
for (const row of dataRows) {
  const sku = row[skuIndex];
  const description = descriptions.get(sku);
  if (!description) {
    throw new Error(`Brak opisu dla SKU ${sku}.`);
  }
  row[descriptionIndex] = description;
}

if (dataRows.length !== descriptions.size) {
  throw new Error(`Liczba produktów (${dataRows.length}) nie odpowiada liczbie opisów (${descriptions.size}).`);
}

const csv = [headers, ...dataRows]
  .map((row) => row.map(escapeCsv).join(";"))
  .join("\r\n") + "\r\n";

await fs.writeFile(outputPath, `\uFEFF${csv}`, "utf8");

console.log(JSON.stringify({
  outputPath,
  rows: dataRows.length,
  descriptions: dataRows.map((row) => ({
    sku: row[skuIndex],
    htmlCharacters: row[descriptionIndex].length,
  })),
}, null, 2));
