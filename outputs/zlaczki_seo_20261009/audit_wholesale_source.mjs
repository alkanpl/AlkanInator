import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const simetPath = "C:/Users/Handlowiec/Desktop/Puszki i złączki/CENNIK SIMET.xlsx";
const morekPath = "C:/Users/Handlowiec/Desktop/Puszki i złączki/MOREK - CENNIK.xlsx";

const morek = await SpreadsheetFile.importXlsx(await FileBlob.load(morekPath));
const morekSheet = morek.worksheets.getItem("CENNIK BAZOWY");
console.log("MOREK_HEADERS");
console.log(JSON.stringify(morekSheet.getRange("A13:Z20").values, null, 2));
