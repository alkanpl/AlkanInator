import fs from "node:fs/promises";

const outputPath = "outputs/bloki_rozdzielcze_20261009/BASELINKER_bloki_rozdzielcze_Morek_20261009.csv";
const headers = [
  "product_id",
  "name",
  "sku",
  "ean",
  "manufacturer_name",
  "category",
  "description",
  "features",
  "images_urls",
];

const manufacturerData = "MOREK, ul. Grzybowska 87 00-844 Warszawa, morekpl@morek.eu";
const category = "Puszki i złączki > Bloki rozdzielcze";

const common = {
  "Producent": "Morek",
  "Dane producenta": manufacturerData,
  "Rodzaj": "Blok dystrybucyjny",
  "Seria": "MOBLOCK",
  "Prąd znamionowy [A]": "125",
  "Materiał przewodu": "Cu",
  "Przekrój przewodu elastycznego z tulejką [mm²]": "1,5-25",
  "Przekrój przewodu sztywnego [mm²]": "1,5-35",
  "Stopień ochrony [IP]": "IP20",
  "Sposób montażu": "Szyna DIN lub śruby",
  "Materiał części przewodzących": "Mosiądz",
  "Materiał śrub zaciskowych": "Stal ocynkowana",
  "Osłona": "Transparentna, zdejmowana",
  "Szyna PE": "Tak",
  "Szyna N": "Tak",
  "Klasa ETIM": "EC001894",
};

const products = [
  {
    name: "Blok dystrybucyjny MOBLOCK 2P 1000V 125A-7 2x35mm2 + 5x6mm2 Morek MAB2124A18",
    sku: "MAB2124A18/MOR",
    ean: "4744397041906",
    image: "https://eleczaplive.blob.core.windows.net/productimages/5/715fbd723d524bc7a2b6a76b41e397fc.png",
    features: {
      ...common,
      "Liczba biegunów": "2",
      "Liczba zacisków na biegun": "7",
      "Przekrój wejścia Cu [mm²]": "2 x 10-35",
      "Przekrój wyjść Cu [mm²]": "5 x 1,5-6",
      "Napięcie znamionowe AC/DC [V]": "1000",
      "Wymiary (dł. x szer. x wys.) [mm]": "72 x 49 x 52",
    },
  },
  {
    name: "Blok dystrybucyjny MOBLOCK 2P 1000V 125A-11 2x35mm2 + 2x25mm2 + 7x6mm2 Morek MAB2125A18",
    sku: "MAB2125A18/MOR",
    ean: "4744397013590",
    image: "https://eleczaplive.blob.core.windows.net/productimages/5/MAB2125A18.png",
    features: {
      ...common,
      "Liczba biegunów": "2",
      "Liczba zacisków na biegun": "11",
      "Przekrój wejścia Cu [mm²]": "2 x 10-35",
      "Przekrój wyjść Cu [mm²]": "2 x 10-25 + 7 x 1,5-6",
      "Napięcie znamionowe AC/DC [V]": "1000",
      "Wymiary (dł. x szer. x wys.) [mm]": "109 x 49 x 52",
    },
  },
  {
    name: "Blok dystrybucyjny MOBLOCK 2P 1000V 125A-15 2x35mm2 + 2x25mm2 + 11x6mm2 Morek MAB2126A18",
    sku: "MAB2126A18/MOR",
    ean: "4744397013606",
    image: "https://eleczaplive.blob.core.windows.net/productimages/5/MAB2126A18.png",
    features: {
      ...common,
      "Liczba biegunów": "2",
      "Liczba zacisków na biegun": "15",
      "Przekrój wejścia Cu [mm²]": "2 x 10-35",
      "Przekrój wyjść Cu [mm²]": "2 x 10-25 + 11 x 1,5-6",
      "Napięcie znamionowe AC/DC [V]": "1000",
      "Wymiary (dł. x szer. x wys.) [mm]": "137 x 49 x 52",
    },
  },
  {
    name: "Blok dystrybucyjny MOBLOCK 4P 500V 125A-7 2x35mm2 + 5x6mm2 Morek MAB4124A18",
    sku: "MAB4124A18/MOR",
    ean: "4744397041913",
    image: "https://eleczaplive.blob.core.windows.net/productimages/5/c1f713cb0a094352a51f1e8514aa5437.png",
    features: {
      ...common,
      "Liczba biegunów": "4",
      "Liczba zacisków na biegun": "7",
      "Przekrój wejścia Cu [mm²]": "2 x 10-35",
      "Przekrój wyjść Cu [mm²]": "5 x 1,5-6",
      "Napięcie znamionowe AC/DC [V]": "500",
      "Wymiary (dł. x szer. x wys.) [mm]": "72 x 97 x 52",
    },
  },
];

const rows = products.map((product) => ({
  product_id: "",
  name: product.name,
  sku: product.sku,
  ean: product.ean,
  manufacturer_name: "Morek",
  category,
  description: "",
  features: JSON.stringify(product.features),
  images_urls: product.image,
}));

const escapeCsv = (value) => {
  const text = String(value ?? "");
  return /[;"\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
};
const csv = [headers.join(";"), ...rows.map((row) => headers.map((header) => escapeCsv(row[header])).join(";"))].join("\r\n") + "\r\n";
await fs.writeFile(outputPath, `\uFEFF${csv}`, "utf8");

console.log(JSON.stringify({
  outputPath,
  count: rows.length,
  titles: rows.map((row) => ({ sku: row.sku, title: row.name })),
  images: rows.map((row) => ({ sku: row.sku, image: row.images_urls })),
}, null, 2));
