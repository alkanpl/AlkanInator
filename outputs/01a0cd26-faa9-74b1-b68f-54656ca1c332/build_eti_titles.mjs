import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const inputPath = 'C:/Users/Handlowiec/Desktop/ETI MODUŁÓWKA/Lista produktów do wystawienia.xlsx';
const outputDir = 'outputs/01a0cd26-faa9-74b1-b68f-54656ca1c332';
const outputPath = `${outputDir}/ETI_modulowka_tytuly.xlsx`;

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = wb.worksheets.getItemAt(0);
const sourcePreview = await wb.render({ sheetName: sheet.name, range: 'A1:E8', scale: 1.5, format: 'png' });
await fs.writeFile(`${outputDir}/source_preview.png`, new Uint8Array(await sourcePreview.arrayBuffer()));

const rows = sheet.getRange('A1:E128').values;
const clean = (x) => String(x ?? '').replace(/\s+/g, ' ').trim();
const digits = (x) => { const m = clean(x).match(/\d+/); return m ? m[0] : ''; };
const title = (row) => {
  const code = clean(row[0]);
  const symbol = clean(row[1]);
  const name = clean(row[2]);
  let core = '';
  let m;

  if ((m = symbol.match(/^ETIMAT 6 ([13])p ([BCD])(\d+)$/i))) {
    core = `Wyłącznik nadprądowy ${m[2].toUpperCase()}${m[3]} ${m[1]}P 6kA ETIMAT 6`;
  } else if ((m = symbol.match(/^ETIMAT T (\d)p (\d+)A$/i))) {
    core = `Ogranicznik mocy ${m[2]}A ${m[1]}P ETIMAT T`;
  } else if ((m = symbol.match(/^EFI6-P([24]) (A|AC) (\d+)\/0[.,]03$/i))) {
    const poles = m[1] === '2' ? '1P+N' : '3P+N';
    core = `Wyłącznik różnicowoprądowy RCCB ${m[3]}A 30mA typ ${m[2].toUpperCase()} ${poles} 6kA EFI6-P${m[1]}`;
  } else if ((m = symbol.match(/^KZS-2M (A|AC) ([BC])(\d+)\/0[.,]03$/i))) {
    core = `Wyłącznik różnicowo-nadprądowy RCBO ${m[2].toUpperCase()}${m[3]} 30mA typ ${m[1].toUpperCase()} 1P+N 10kA KZS-2M`;
  } else if ((m = symbol.match(/^SV \d+ ([134])p (\d+)A$/i))) {
    core = `Rozłącznik izolacyjny ${m[2]}A ${m[1]}P ${symbol.split(' ').slice(0, 2).join(' ')}`;
  } else if ((m = symbol.match(/^SSQ ([134])(25|40|63)$/i))) {
    core = `Przełącznik trójpołożeniowy modułowy ${m[2]}A ${m[1]}P SSQ ${m[1]}${m[2]}`;
  } else if (symbol === 'SON H-3K') {
    core = 'Sygnalizator obecności napięcia LED 3 kolory SON H-3K';
  } else if (symbol === 'SON H-3G') {
    core = 'Sygnalizator obecności napięcia LED 3 zielone diody SON H-3G';
  } else if ((m = symbol.match(/^R (\d+)-(20|11|40) (24|230)V$/i))) {
    const contacts = m[2] === '11' ? '1 styk zwierny i 1 rozwierny' : `${Number(m[2][0])} styki zwierne`;
    core = `Stycznik modułowy ${m[1]}A ${contacts} cewka ${m[3]}V R ${m[1]}-${m[2]}`;
  } else if ((m = symbol.match(/^ED2 (125|160|250)\/3$/i))) {
    core = `Rozłącznik izolacyjny ${m[1]}A 3P ED2 ${m[1]}/3`;
  } else if ((m = symbol.match(/^STV D02 63A ([13])p$/i))) {
    core = `Rozłącznik bezpiecznikowy D02 63A ${m[1]}P STV D02`;
  } else if (symbol.startsWith('D01-STV D02')) {
    core = 'Wstawka redukcyjna D01 do STV D02 16A szara';
  } else if ((m = symbol.match(/^EFD (10|22) 3p$/i))) {
    core = `Rozłącznik bezpiecznikowy 3P EFD ${m[1]}`;
  } else if (symbol === 'EFH 10 DC 2p') {
    core = 'Podstawa bezpiecznikowa DC 2P EFH 10';
  } else {
    throw new Error(`Nieznany wariant ${code}: ${symbol} / ${name}`);
  }
  return `${core} ETI ${code}`;
};

const titles = rows.slice(1).map((row) => [title(row)]);
if (titles.length !== 127) throw new Error(`Oczekiwano 127 produktów, jest ${titles.length}`);
const unbranded = titles.map(([x]) => x.replace(/ ETI \d{9}$/, ''));
if (new Set(unbranded).size !== unbranded.length) throw new Error('Zduplikowane tytuły przed kodem ETI');
if (new Set(rows.slice(1).map((r) => clean(r[0]))).size !== 127) throw new Error('Zduplikowane numery kodowe');
if (titles.some(([x]) => !x || !/ ETI \d{9}$/.test(x))) throw new Error('Niepełny tytuł');

sheet.getRange('F1').values = [['Nowy tytuł']];
sheet.getRange('F2:F128').values = titles;
sheet.getRange('A1:F128').format.font = { name: 'Arial', size: 10, color: '#1F2937' };
sheet.getRange('A1:F1').format = { fill: '#17324D', font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' }, rowHeight: 30 };
sheet.getRange('A1:F1').format.verticalAlignment = 'center';
sheet.getRange('A1:F1').format.horizontalAlignment = 'center';
sheet.getRange('A:A').format.columnWidth = 17;
sheet.getRange('B:B').format.columnWidth = 28;
sheet.getRange('C:C').format.columnWidth = 53;
sheet.getRange('D:D').format.columnWidth = 19;
sheet.getRange('E:E').format.columnWidth = 25;
sheet.getRange('F:F').format.columnWidth = 110;
sheet.getRange('A2:B128').setNumberFormat('@');
sheet.getRange('D2:D128').setNumberFormat('@');
sheet.getRange('F2:F128').format.rowHeight = 23;
sheet.freezePanes.freezeRows(1);
sheet.showGridLines = false;
wb.recalculate();
const check = await wb.inspect({ kind: 'table', range: 'Do wystawienia!A1:F6', include: 'values,formulas', tableMaxRows: 6, tableMaxCols: 6, maxChars: 5000 });
console.log(check.ndjson);
const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 100 }, maxChars: 3000 });
console.log(errors.ndjson);
const preview = await wb.render({ sheetName: sheet.name, range: 'A1:F8', scale: 1.2, format: 'png' });
await fs.writeFile(`${outputDir}/titles_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const file = await SpreadsheetFile.exportXlsx(wb);
await file.save(outputPath);
console.log(JSON.stringify({ outputPath, count: titles.length, samples: [titles[0][0], titles[58][0], titles[63][0], titles[99][0], titles[126][0]] }));
