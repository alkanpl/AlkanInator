import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load("C:/Users/Handlowiec/Downloads/zlaczki-opisy-z-woocomerce.xlsx"));
const sheet = workbook.worksheets.getItem("Worksheet");
console.log(JSON.stringify(sheet.getRange("A258:H301").values, null, 2));
