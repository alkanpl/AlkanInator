import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Handlowiec/Downloads/zlaczki-opisy-z-woocomerce.xlsx";
const descriptionsDir = "reports/descriptions/zlaczki_20261009";

const normalize = (value) => String(value ?? "").trim().toUpperCase();
const stripSupplier = (value) => normalize(value).replace(/\/(SIM|MOR)$/i, "");

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = workbook.worksheets.getItem("Worksheet");
const values = sheet.getRange("A1:H301").values;
const headers = values[0].map((value) => String(value ?? "").trim());
const skuIndex = headers.indexOf("Sku");
const rows = values.slice(1).map((row, offset) => ({
  excelRow: offset + 2,
  sku: normalize(row[skuIndex]),
  baseSku: stripSupplier(row[skuIndex]),
  title: String(row[headers.indexOf("Title")] ?? ""),
}));

const files = (await fs.readdir(descriptionsDir)).filter((name) => name.endsWith(".html")).sort();
const targets = files.map((file) => {
  const sku = normalize(file.replace(/\.html$/i, "").replace(/_(SIM|MOR)$/i, "/$1"));
  const exact = rows.filter((row) => row.sku === sku);
  const base = rows.filter((row) => row.baseSku === stripSupplier(sku));
  return { file, sku, exact, base };
});

console.log(JSON.stringify({
  targetCount: targets.length,
  exactMatchCount: targets.filter((target) => target.exact.length === 1).length,
  baseMatchCount: targets.filter((target) => target.base.length === 1).length,
  unmatched: targets.filter((target) => target.base.length === 0),
  ambiguous: targets.filter((target) => target.base.length > 1),
  matches: targets.map((target) => ({
    sourceSku: target.sku,
    excelRow: target.base[0]?.excelRow ?? null,
    workbookSku: target.base[0]?.sku ?? null,
    title: target.base[0]?.title ?? null,
  })),
}, null, 2));
