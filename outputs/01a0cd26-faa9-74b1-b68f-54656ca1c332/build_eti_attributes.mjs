import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const dir = 'outputs/01a0cd26-faa9-74b1-b68f-54656ca1c332';
const input = `${dir}/ETI_modulowka_tytuly.xlsx`;
const output = `${dir}/ETI_modulowka_tytuly_atrybuty.xlsx`;
const original = 'Lista produktów do wystawienia.xlsx, kol. Nr kodowy / Symbol / Nazwa wyrobu';
const eti = 'https://www.etipolam.com.pl/katalog-elektroniczny/';
const links = {
  mcb: `${eti}aparaty-modulowe-zabezpieczajace/wylaczniki-nadpradowe-i-akcesoria/wylaczniki-nadpradowe/002111512-etimat-6-1p-b6`,
  mcbD: `${eti}aparaty-modulowe-zabezpieczajace/wylaczniki-nadpradowe-i-akcesoria/wylaczniki-nadpradowe/002164516-etimat-6-3p-d16`,
  rccb2: `${eti}002061261`,
  rccb4: `${eti}002061662`,
  rcbo: `${eti}aparaty-modulowe-zabezpieczajace/wylaczniki-roznicowopradowe-z-zabezpieczeniem-nadpradowym-i-akcesoria/wylaczniki-roznicowopradowe-z-zabezpieczeniem-nadpradowym`,
  son: 'https://www.etipolam.com.pl/images/product_db/documents/pl-PL/03_EVE-ETIREL_PL_2022_2023.pdf',
  ssq25: 'https://www.etipolam.com.pl/images/userfiles/pl-PL/ETI%20POLAM/Wsparcie%20Techniczne/Broszury_ulotki_karty_katalogowe/_SSQ.pdf',
  ssq63: 'https://www.eti.si/images/product_db/idents/en-GB/002421417_TD.pdf',
  efd10: `${eti}wkladki-topikowe-d-d0-c/rozlaczniki-i-podstawy-bezpiecznikowe-do-systemu-c/002540004-efd-10-3p`,
  efd22: `${eti}002570004`,
  efh: `${eti}wkladki-topikowe-d-d0-c/podstawy-do-wkladek-ch-dc/002540203-efh-10-dc-2p`,
  stv: `${eti}/002271001-stv-d02-stv-d02-1-63a-230-400v-110v`,
};

const columns = [
  'Status P0', 'Brakujące P0', 'Producent', 'Typ Produktu', 'Seria',
  'Charakterystyka Wyzwalania', 'Prąd Znamionowy [A]', 'Układ Biegunów',
  'Znamionowa Zdolność Zwarciowa Icn [kA]', 'Prąd Różnicowy IΔn [mA]',
  'Typ Prądu Różnicowego', 'Warunkowy Prąd Zwarciowy Inc [kA]',
  'Napięcie Znamionowe AC [V]', 'Napięcie Znamionowe DC [V]',
  'Napięcie Cewki [V]', 'Konfiguracja Styków', 'Rodzaj Wkładki',
  'Rozmiar', 'Barwa Lampek', 'Ilość Lampek', 'Pozycje Przełącznika',
  'Kompatybilność', 'Rodzaj Akcesoria', 'Liczba Modułów',
  'Stopień Ochrony [IP]', 'Kategoria Użytkowania',
];
const required = {
  mcb: ['Producent', 'Symbol', 'Typ Produktu', 'Charakterystyka Wyzwalania', 'Prąd Znamionowy [A]', 'Układ Biegunów', 'Znamionowa Zdolność Zwarciowa Icn [kA]'],
  limiter: ['Producent', 'Symbol', 'Typ Produktu', 'Prąd Znamionowy [A]', 'Układ Biegunów'],
  rccb: ['Producent', 'Symbol', 'Typ Produktu', 'Prąd Znamionowy [A]', 'Prąd Różnicowy IΔn [mA]', 'Typ Prądu Różnicowego', 'Układ Biegunów'],
  rcbo: ['Producent', 'Symbol', 'Typ Produktu', 'Charakterystyka Wyzwalania', 'Prąd Znamionowy [A]', 'Prąd Różnicowy IΔn [mA]', 'Typ Prądu Różnicowego', 'Układ Biegunów', 'Znamionowa Zdolność Zwarciowa Icn [kA]'],
  isolator: ['Producent', 'Symbol', 'Typ Produktu', 'Prąd Znamionowy [A]', 'Układ Biegunów'],
  switch: ['Producent', 'Symbol', 'Typ Produktu', 'Prąd Znamionowy [A]', 'Układ Biegunów', 'Napięcie Znamionowe AC [V]', 'Pozycje Przełącznika'],
  indicator: ['Producent', 'Symbol', 'Typ Produktu', 'Barwa Lampek', 'Napięcie Znamionowe AC [V]'],
  contactor: ['Producent', 'Symbol', 'Typ Produktu', 'Prąd Znamionowy [A]', 'Napięcie Cewki [V]', 'Konfiguracja Styków'],
  fuse: ['Producent', 'Symbol', 'Typ Produktu', 'Rodzaj Wkładki', 'Prąd Znamionowy [A]', 'Układ Biegunów', 'Rozmiar'],
  accessory: ['Producent', 'Symbol', 'Typ Produktu', 'Rodzaj Akcesoria', 'Kompatybilność'],
};

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const sheet = wb.worksheets.getItemAt(0);
const sourceRows = sheet.getRange('A2:F128').values;
if (sourceRows.length !== 127) throw new Error('Nieoczekiwana liczba produktów');

const audit = [];
const main = [];
const familyCounts = {};
const clean = (x) => String(x ?? '').replace(/\s+/g, ' ').trim();
for (const row of sourceRows) {
  const code = clean(row[0]);
  const symbol = clean(row[1]);
  const name = clean(row[2]);
  const attrs = {};
  let family = '';
  const put = (key, value, source = original, confidence = 1) => {
    if (value === undefined || value === null || value === '') return;
    attrs[key] = { value, source, confidence };
  };
  let m;
  put('Producent', 'ETI', 'Zlecenie użytkownika: produkty ETI; Lista produktów do wystawienia.xlsx', 1);
  put('Symbol', symbol);
  if ((m = symbol.match(/^ETIMAT 6 ([13])p ([BCD])(\d+)$/i))) {
    family = 'mcb';
    put('Typ Produktu', 'Wyłącznik nadprądowy MCB');
    put('Seria', 'ETIMAT 6');
    put('Charakterystyka Wyzwalania', m[2].toUpperCase());
    put('Prąd Znamionowy [A]', Number(m[3]));
    put('Układ Biegunów', `${m[1]}P`);
    put('Znamionowa Zdolność Zwarciowa Icn [kA]', 6, m[2].toUpperCase() === 'D' ? links.mcbD : links.mcb, 0.97);
  } else if ((m = symbol.match(/^ETIMAT T (\d)p (\d+)A$/i))) {
    family = 'limiter';
    put('Typ Produktu', 'Ogranicznik mocy');
    put('Seria', 'ETIMAT T');
    put('Prąd Znamionowy [A]', Number(m[2]));
    put('Układ Biegunów', `${m[1]}P`);
  } else if ((m = symbol.match(/^EFI6-P([24]) (A|AC) (\d+)\/0[.,]03$/i))) {
    family = 'rccb';
    put('Typ Produktu', 'Wyłącznik różnicowoprądowy RCCB');
    put('Seria', `EFI6-P${m[1]}`);
    put('Prąd Znamionowy [A]', Number(m[3]));
    put('Prąd Różnicowy IΔn [mA]', 30);
    put('Typ Prądu Różnicowego', m[2].toUpperCase());
    put('Układ Biegunów', m[1] === '2' ? '1P+N' : '3P+N', m[1] === '2' ? links.rccb2 : links.rccb4, 0.98);
    put('Warunkowy Prąd Zwarciowy Inc [kA]', 6, m[1] === '2' ? links.rccb2 : links.rccb4, 0.97);
  } else if ((m = symbol.match(/^KZS-2M (A|AC) ([BC])(\d+)\/0[.,]03$/i))) {
    family = 'rcbo';
    put('Typ Produktu', 'Wyłącznik różnicowo-nadprądowy RCBO');
    put('Seria', 'KZS-2M');
    put('Charakterystyka Wyzwalania', m[2].toUpperCase());
    put('Prąd Znamionowy [A]', Number(m[3]));
    put('Prąd Różnicowy IΔn [mA]', 30);
    put('Typ Prądu Różnicowego', m[1].toUpperCase());
    put('Układ Biegunów', '1P+N', links.rcbo, 0.97);
    put('Znamionowa Zdolność Zwarciowa Icn [kA]', 10, links.rcbo, 0.96);
  } else if ((m = symbol.match(/^SV \d+ ([134])p (\d+)A$/i))) {
    family = 'isolator';
    put('Typ Produktu', 'Rozłącznik izolacyjny');
    put('Seria', 'SV');
    put('Prąd Znamionowy [A]', Number(m[2]));
    put('Układ Biegunów', `${m[1]}P`);
  } else if ((m = symbol.match(/^SSQ ([134])(25|40|63)$/i))) {
    family = 'switch';
    put('Typ Produktu', 'Przełącznik trójpołożeniowy');
    put('Seria', 'SSQ');
    put('Prąd Znamionowy [A]', Number(m[2]));
    put('Układ Biegunów', `${m[1]}P`);
    put('Napięcie Znamionowe AC [V]', '230/400', Number(m[2]) === 63 ? links.ssq63 : links.ssq25, 0.96);
    put('Pozycje Przełącznika', 'I-0-II', Number(m[2]) === 63 ? links.ssq63 : links.ssq25, 0.98);
  } else if (symbol === 'SON H-3K' || symbol === 'SON H-3G') {
    family = 'indicator';
    put('Typ Produktu', 'Sygnalizator obecności napięcia');
    put('Seria', 'SON H');
    put('Barwa Lampek', symbol.endsWith('3K') ? 'Czerwona / żółta / zielona' : 'Zielona', links.son, 0.98);
    put('Ilość Lampek', 3);
    put('Napięcie Znamionowe AC [V]', 230, `${eti}${code}`, 0.98);
  } else if ((m = symbol.match(/^R (\d+)-(20|11|40) (24|230)V$/i))) {
    family = 'contactor';
    put('Typ Produktu', 'Stycznik modułowy');
    put('Seria', 'R');
    put('Prąd Znamionowy [A]', Number(m[1]));
    put('Napięcie Cewki [V]', Number(m[3]));
    put('Konfiguracja Styków', m[2] === '11' ? '1NO+1NC' : `${m[2][0]}NO`);
    const modules = name.match(/\((\d+) mod\./i);
    if (modules) put('Liczba Modułów', Number(modules[1]));
  } else if ((m = symbol.match(/^ED2 (125|160|250)\/3$/i))) {
    family = 'isolator';
    put('Typ Produktu', 'Rozłącznik izolacyjny');
    put('Seria', 'ED2');
    put('Prąd Znamionowy [A]', Number(m[1]));
    put('Układ Biegunów', '3P');
    put('Napięcie Znamionowe AC [V]', 690, `${eti}${code}`, 0.96);
  } else if ((m = symbol.match(/^STV D02 63A ([13])p$/i))) {
    family = 'fuse';
    put('Typ Produktu', 'Rozłącznik bezpiecznikowy');
    put('Seria', 'STV D02');
    put('Prąd Znamionowy [A]', 63);
    put('Układ Biegunów', `${m[1]}P`);
    put('Rodzaj Wkładki', 'D02');
    put('Rozmiar', 'D02');
  } else if (symbol.startsWith('D01-STV D02')) {
    family = 'accessory';
    put('Typ Produktu', 'Wstawka redukcyjna');
    put('Rodzaj Akcesoria', 'Wstawka redukcyjna');
    put('Kompatybilność', 'D01 do STV D02');
    put('Prąd Znamionowy [A]', 16);
  } else if ((m = symbol.match(/^EFD (10|22) 3p$/i))) {
    family = 'fuse';
    const is10 = m[1] === '10';
    const url = is10 ? links.efd10 : links.efd22;
    put('Typ Produktu', 'Rozłącznik bezpiecznikowy');
    put('Seria', `EFD ${m[1]}`);
    put('Rodzaj Wkładki', 'Cylindryczna CH', url, 0.99);
    put('Prąd Znamionowy [A]', is10 ? 32 : 100, url, 0.99);
    put('Układ Biegunów', '3P');
    put('Rozmiar', is10 ? '10x38 mm' : '22x58 mm', url, 0.99);
    put('Napięcie Znamionowe AC [V]', 690, url, 0.99);
    put('Kategoria Użytkowania', 'AC-22B', url, 0.98);
  } else if (symbol === 'EFH 10 DC 2p') {
    family = 'fuse';
    put('Typ Produktu', 'Podstawa bezpiecznikowa DC');
    put('Seria', 'EFH 10');
    put('Rodzaj Wkładki', 'Cylindryczna CH gPV', links.efh, 0.99);
    put('Prąd Znamionowy [A]', 25, links.efh, 0.99);
    put('Układ Biegunów', '2P');
    put('Rozmiar', '10x38 mm', links.efh, 0.99);
    put('Napięcie Znamionowe DC [V]', 1000, links.efh, 0.99);
  } else {
    throw new Error(`Nierozpoznana rodzina: ${code} ${symbol}`);
  }

  familyCounts[family] = (familyCounts[family] || 0) + 1;
  const needed = required[family];
  const missing = needed.filter((key) => !(key in attrs));
  main.push([missing.length ? 'DO_SPRAWDZENIA' : 'OK', missing.join(', '), ...columns.slice(2).map((key) => attrs[key]?.value ?? null)]);
  for (const [key, item] of Object.entries(attrs)) {
    audit.push([code, symbol, family, key, needed.includes(key) ? 'P0' : 'P1', item.value, item.source, item.confidence, 'OK']);
  }
  for (const key of missing) audit.push([code, symbol, family, key, 'P0', null, null, 0, 'DO_SPRAWDZENIA']);
}

const lastColumn = (n) => { let s = ''; for (let x = n; x; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s; return s; };
const last = lastColumn(6 + columns.length);
sheet.getRange(`G1:${last}1`).values = [columns];
sheet.getRange(`G2:${last}128`).values = main;
sheet.getRange(`G1:${last}1`).format = { fill: '#315D78', font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' }, rowHeight: 36 };
sheet.getRange(`G1:${last}1`).format.wrapText = true;
sheet.getRange(`G2:${last}128`).format.font = { name: 'Arial', size: 10, color: '#1F2937' };
sheet.getRange('G:G').format.columnWidth = 20;
sheet.getRange('H:H').format.columnWidth = 42;
sheet.getRange(`I:${last}`).format.columnWidth = 23;
sheet.getRange('J:J').format.columnWidth = 38;
sheet.getRange('X:X').format.columnWidth = 31;
sheet.getRange('G2:G128').conditionalFormats.add('containsText', { text: 'DO_SPRAWDZENIA', format: { fill: '#FEE2E2', font: { color: '#991B1B', bold: true } } });

const evidence = wb.worksheets.add('Źródła atrybutów');
evidence.getRange('A1:I1').values = [['Nr kodowy', 'Symbol', 'Rodzina', 'Atrybut', 'Priorytet', 'Wartość', 'Źródło', 'Pewność', 'Status']];
evidence.getRange(`A2:I${audit.length + 1}`).values = audit;
evidence.getRange('A1:I1').format = { fill: '#17324D', font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' }, rowHeight: 28 };
evidence.getRange(`A2:I${audit.length + 1}`).format.font = { name: 'Arial', size: 10, color: '#1F2937' };
evidence.getRange('A:A').format.columnWidth = 17;
evidence.getRange('B:B').format.columnWidth = 26;
evidence.getRange('C:C').format.columnWidth = 16;
evidence.getRange('D:D').format.columnWidth = 48;
evidence.getRange('E:E').format.columnWidth = 13;
evidence.getRange('F:F').format.columnWidth = 31;
evidence.getRange('G:G').format.columnWidth = 85;
evidence.getRange('H:H').format.columnWidth = 13;
evidence.getRange('I:I').format.columnWidth = 22;
evidence.getRange(`H2:H${audit.length + 1}`).setNumberFormat('0%');
evidence.getRange(`A2:B${audit.length + 1}`).setNumberFormat('@');
evidence.freezePanes.freezeRows(1);
evidence.showGridLines = false;
evidence.tabColor = '#64748B';

wb.recalculate();
const check = await wb.inspect({ kind: 'table', range: `Do wystawienia!G1:P5`, include: 'values,formulas', tableMaxRows: 5, tableMaxCols: 10, maxChars: 6000 });
console.log(check.ndjson);
const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 100 }, maxChars: 2000 });
console.log(errors.ndjson);
const preview1 = await wb.render({ sheetName: sheet.name, range: 'G1:P7', scale: 1.2, format: 'png' });
await fs.writeFile(`${dir}/attributes_preview.png`, new Uint8Array(await preview1.arrayBuffer()));
const preview2 = await wb.render({ sheetName: evidence.name, range: 'A1:I8', scale: 1.2, format: 'png' });
await fs.writeFile(`${dir}/evidence_preview.png`, new Uint8Array(await preview2.arrayBuffer()));
const file = await SpreadsheetFile.exportXlsx(wb);
await file.save(output);
console.log(JSON.stringify({ output, products: main.length, auditRows: audit.length, families: familyCounts, missingP0Products: main.filter((r) => r[0] !== 'OK').length }));
