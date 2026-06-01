import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputFile = path.resolve("outputs", "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V3_FORMS.xlsx");
const docsIndexFile = path.resolve("outputs", "document_evidence_index.json");
const outputDir = path.resolve("outputs");
const outputFile = path.join(outputDir, "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V4B_NATURE_DASHBOARD.xlsx");

const input = await FileBlob.load(inputFile);
const workbook = await SpreadsheetFile.importXlsx(input);

let documentRecords = [];
try {
  documentRecords = JSON.parse(await fs.readFile(docsIndexFile, "utf8"));
} catch {
  documentRecords = [];
}

const palette = {
  navy: "#0B1F33",
  teal: "#0F766E",
  green: "#166534",
  leaf: "#2E7D32",
  amber: "#F59E0B",
  red: "#DC2626",
  blue: "#2563EB",
  lightTeal: "#E6F4F1",
  lightGreen: "#EAF7EA",
  lightAmber: "#FFF7E6",
  lightRed: "#FEE2E2",
  white: "#FFFFFF",
  slate: "#334155",
  grey: "#F8FAFC",
};

const VALID_STATUS = "Validé humainement";
const FORM_ROWS = 120;
const HEADER_ROW = 4;
const DATA_START = 5;
const DATA_END = DATA_START + FORM_ROWS - 1;

const lists = {
  years: ["2024", "2025", "2026", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "2034", "2035"],
  months: ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"],
  communes: [
    "Dakar-Plateau",
    "Médina",
    "Gueule Tapée-Fass-Colobane",
    "Fann-Point E-Amitié",
    "Grand Dakar",
    "Biscuiterie",
    "HLM",
    "Hann Bel-Air",
    "Sicap-Liberté",
    "Dieuppeul-Derklé",
    "Mermoz-Sacré-Coeur",
    "Ouakam",
    "Ngor",
    "Yoff",
    "Cambérène",
    "Parcelles Assainies",
    "Patte d'Oie",
    "Grand-Yoff",
    "Ville de Dakar",
  ],
  greenTypes: [
    "Parc",
    "Jardin public",
    "Jardin communautaire",
    "Alignement d'arbres",
    "Corridor vert",
    "Micro-jardin",
    "Forêt urbaine",
    "Place reverdie",
    "École reverdie",
    "Marché reverdi",
    "Zone côtière restaurée",
  ],
  greenUnits: ["m²", "hectares", "km²"],
  institutions: [
    "Ville de Dakar",
    "Direction Environnement",
    "Direction Espaces Verts",
    "Direction Planification",
    "Direction SIG",
    "Direction Éducation",
    "Direction Marchés",
    "ONG / Association",
    "Université / Recherche",
    "CSE",
    "ANACIM",
    "Bailleur",
    "Prestataire",
  ],
  validation: ["Brouillon", "À vérifier", "Validé techniquement", VALID_STATUS, "Rejeté"],
  uncertainty: ["Faible", "Moyen", "Élevé", "Très élevé"],
  methodologies: ["i-Tree Eco", "i-Tree Canopy", "IPCC AFOLU - estimation locale", "Facteur municipal provisoire", "Étude botanique locale", "Autre méthodologie documentée"],
  vegetation: ["Arbre urbain", "Forêt urbaine", "Mangrove / zone côtière", "Pelouse / prairie", "Jardin maraîcher", "Haie / corridor vert", "Mixte"],
  priorities: ["P0", "P1", "P2", "P3"],
  frequencies: ["Mensuelle", "Trimestrielle", "Semestrielle", "Annuelle", "Ponctuelle"],
};

function colName(n) {
  let name = "";
  let x = n;
  while (x > 0) {
    const r = (x - 1) % 26;
    name = String.fromCharCode(65 + r) + name;
    x = Math.floor((x - 1) / 26);
  }
  return name;
}

function q(sheetName) {
  return `'${sheetName.replaceAll("'", "''")}'`;
}

function deleteIfExists(name) {
  const existing = workbook.worksheets.getItemOrNullObject?.(name);
  if (existing && !existing.isNullObject) {
    existing.delete?.();
  }
}

for (const name of [
  "10_Form_Nature_Espaces_Verts",
  "Data_Nature_Espaces_Verts",
  "Facteurs_Sequestration_Carbone",
  "Nature_Espaces_Verts",
  "20_Tableau_de_Bord_Dynamique",
  "19_Graphiques_Dynamiques",
  "21_Controle_Qualite",
  "22_Planning_Collecte",
  "23_Registre_Documentaire",
  "24_Parametres_Listes",
]) {
  deleteIfExists(name);
}

function titleSheet(sheet, title, subtitle, endCol = "L") {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${endCol}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${endCol}1`).format = { fill: palette.navy, font: { bold: true, color: palette.white, size: 15 } };
  sheet.getRange(`A1:${endCol}1`).format.rowHeightPx = 34;
  sheet.getRange(`A2:${endCol}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange(`A2:${endCol}2`).format = { fill: palette.lightGreen, font: { italic: true, color: palette.slate }, wrapText: true };
  sheet.getRange(`A2:${endCol}2`).format.rowHeightPx = 46;
}

function addSheet(name, title, subtitle, endCol = "L") {
  const sheet = workbook.worksheets.add(name);
  titleSheet(sheet, title, subtitle, endCol);
  return sheet;
}

function setWidths(sheet, widths, rowLimit = 160) {
  widths.forEach((w, index) => {
    const col = colName(index + 1);
    sheet.getRange(`${col}1:${col}${rowLimit}`).format.columnWidthPx = w;
  });
}

function addValidation(sheet, colIndex, startRow, endRow, values) {
  const col = colName(colIndex);
  sheet.getRange(`${col}${startRow}:${col}${endRow}`).dataValidation = {
    rule: { type: "list", values },
    prompt: { showPrompt: true, title: "Liste contrôlée", message: "Choisir une valeur autorisée." },
    errorAlert: { showAlert: true, style: "warning", title: "Valeur inattendue", message: "Veuillez choisir une valeur de la liste." },
  };
}

function addStatusFormatting(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.conditionalFormats.add("containsText", { text: "ALERTE", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "OK", format: { fill: palette.lightGreen, font: { color: "#166534", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "À corriger", format: { fill: palette.lightAmber, font: { color: "#92400E", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "Validé", format: { fill: palette.lightGreen, font: { color: "#166534" } } });
  range.conditionalFormats.add("containsText", { text: "P0", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "P1", format: { fill: palette.lightAmber, font: { color: "#92400E", bold: true } } });
}

function addTable(sheet, range, tableName, style = "TableStyleMedium4") {
  const table = sheet.tables.add(range, true, tableName);
  table.style = style;
  table.showFilterButton = true;
  return table;
}

const natureHeaders = [
  "ID_Saisie",
  "Année",
  "Mois",
  "Commune / zone",
  "Type d'espace vert",
  "Superficie espaces verts",
  "Unité",
  "Nombre d'arbres plantés",
  "Nombre d'arbres survivants",
  "Taux de survie des arbres",
  "Nombre de jardins publics",
  "Nombre de jardins communautaires",
  "Nombre de micro-jardins",
  "Superficie reverdie",
  "Superficie restaurée",
  "Espèces plantées",
  "Quantité estimée de CO2 séquestrée",
  "Population bénéficiaire",
  "Projet associé",
  "Institution source",
  "Document source",
  "Statut validation",
  "Observations",
  "Méthodologie séquestration",
  "Facteur de calcul",
  "Source méthodologique",
  "Niveau d'incertitude",
  "Date de saisie",
  "Responsable",
  "Alerte_QC",
  "Alimente_Rapports",
];

function buildNatureForm() {
  const endCol = colName(natureHeaders.length);
  const sheet = addSheet(
    "10_Form_Nature_Espaces_Verts",
    "Formulaire Nature, espaces verts et séquestration carbone",
    "Saisie des espaces verts, arbres, jardins, reboisement, biodiversité et solutions fondées sur la nature. La séquestration est une estimation documentée et reste séparée des émissions directes.",
    endCol
  );
  const blankRows = Array.from({ length: FORM_ROWS }, () => natureHeaders.map(() => null));
  sheet.getRange(`A${HEADER_ROW}:${endCol}${DATA_END}`).values = [natureHeaders, ...blankRows];
  const formulas = {
    1: (r) => `=IF(B${r}="","","NAT-"&TEXT(ROW()-${HEADER_ROW},"0000"))`,
    10: (r) => `=IF(OR(H${r}="",I${r}="",H${r}=0),"",I${r}/H${r})`,
    17: (r) => `=IF(Y${r}="","",IF(I${r}<>"",I${r}*Y${r},IF(AND(F${r}<>"",G${r}="hectares"),F${r}*Y${r},IF(AND(F${r}<>"",G${r}="m²"),F${r}/10000*Y${r},IF(AND(F${r}<>"",G${r}="km²"),F${r}*100*Y${r},"")))))`,
    28: (r) => `=IF(B${r}="","",TODAY())`,
    30: (r) =>
      `=IF(B${r}="","",IF(OR(D${r}="",T${r}="",U${r}="",V${r}<>"${VALID_STATUS}",AND(F${r}<>"",G${r}=""),AND(H${r}<>"",J${r}=""),AND(Q${r}<>"",OR(X${r}="",Y${r}="",Z${r}=""))),"ALERTE: localisation/source/unité/statut/facteur","OK"))`,
    31: (r) => `=IF(B${r}="","",IF(V${r}="${VALID_STATUS}","Oui","Non"))`,
  };
  for (const [colIndexText, formulaFactory] of Object.entries(formulas)) {
    const col = colName(Number(colIndexText));
    const matrix = [];
    for (let row = DATA_START; row <= DATA_END; row += 1) matrix.push([formulaFactory(row)]);
    sheet.getRange(`${col}${DATA_START}:${col}${DATA_END}`).formulas = matrix;
  }
  sheet.getRange(`A${HEADER_ROW}:${endCol}${HEADER_ROW}`).format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${DATA_START}:${endCol}${DATA_END}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A${HEADER_ROW}:${endCol}${DATA_END}`, "tblFormNatureEspacesVerts");
  setWidths(sheet, [110, 90, 115, 180, 175, 160, 105, 155, 165, 145, 150, 180, 155, 150, 150, 210, 190, 160, 170, 180, 220, 160, 250, 190, 135, 220, 155, 120, 150, 230, 130], DATA_END + 5);
  sheet.freezePanes.freezeRows(HEADER_ROW);
  addValidation(sheet, 2, DATA_START, DATA_END, lists.years);
  addValidation(sheet, 3, DATA_START, DATA_END, lists.months);
  addValidation(sheet, 4, DATA_START, DATA_END, lists.communes);
  addValidation(sheet, 5, DATA_START, DATA_END, lists.greenTypes);
  addValidation(sheet, 7, DATA_START, DATA_END, lists.greenUnits);
  addValidation(sheet, 20, DATA_START, DATA_END, lists.institutions);
  addValidation(sheet, 22, DATA_START, DATA_END, lists.validation);
  addValidation(sheet, 24, DATA_START, DATA_END, lists.methodologies);
  addValidation(sheet, 27, DATA_START, DATA_END, lists.uncertainty);
  sheet.getRange(`F${DATA_START}:F${DATA_END}`).setNumberFormat("#,##0.00");
  sheet.getRange(`H${DATA_START}:I${DATA_END}`).setNumberFormat("#,##0");
  sheet.getRange(`J${DATA_START}:J${DATA_END}`).setNumberFormat("0%");
  sheet.getRange(`K${DATA_START}:O${DATA_END}`).setNumberFormat("#,##0.00");
  sheet.getRange(`Q${DATA_START}:Q${DATA_END}`).setNumberFormat("#,##0.000");
  sheet.getRange(`R${DATA_START}:R${DATA_END}`).setNumberFormat("#,##0");
  sheet.getRange(`Y${DATA_START}:Y${DATA_END}`).setNumberFormat("#,##0.0000");
  sheet.getRange(`AB${DATA_START}:AB${DATA_END}`).setNumberFormat("yyyy-mm-dd");
  addStatusFormatting(sheet, `V${DATA_START}:V${DATA_END}`);
  addStatusFormatting(sheet, `AD${DATA_START}:AD${DATA_END}`);
}

function buildNatureData() {
  const endCol = colName(natureHeaders.length);
  const sheet = addSheet(
    "Data_Nature_Espaces_Verts",
    "Base validée Nature, espaces verts et séquestration carbone",
    `Base consolidée alimentée automatiquement depuis 10_Form_Nature_Espaces_Verts. Seules les lignes avec le statut "${VALID_STATUS}" sont reprises.`,
    endCol
  );
  const blankRows = Array.from({ length: FORM_ROWS }, () => natureHeaders.map(() => null));
  sheet.getRange(`A${HEADER_ROW}:${endCol}${DATA_END}`).values = [natureHeaders, ...blankRows];
  for (let colIndex = 1; colIndex <= natureHeaders.length; colIndex += 1) {
    const col = colName(colIndex);
    const formulas = [];
    for (let row = DATA_START; row <= DATA_END; row += 1) {
      formulas.push([`=IF(${q("10_Form_Nature_Espaces_Verts")}!$V${row}="${VALID_STATUS}",${q("10_Form_Nature_Espaces_Verts")}!${col}${row},"")`]);
    }
    sheet.getRange(`${col}${DATA_START}:${col}${DATA_END}`).formulas = formulas;
  }
  sheet.getRange(`A${HEADER_ROW}:${endCol}${HEADER_ROW}`).format = { fill: palette.green, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${DATA_START}:${endCol}${DATA_END}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A${HEADER_ROW}:${endCol}${DATA_END}`, "tblDataNatureEspacesVerts", "TableStyleMedium4");
  setWidths(sheet, [110, 90, 115, 180, 175, 160, 105, 155, 165, 145, 150, 180, 155, 150, 150, 210, 190, 160, 170, 180, 220, 160, 250, 190, 135, 220, 155, 120, 150, 230, 130], DATA_END + 5);
  sheet.freezePanes.freezeRows(HEADER_ROW);
}

function buildSequestrationFactors() {
  const sheet = addSheet(
    "Facteurs_Sequestration_Carbone",
    "Facteurs de séquestration carbone",
    "Bibliothèque méthodologique. Les facteurs sont indicatifs et doivent être remplacés ou validés par une étude locale, i-Tree, IPCC AFOLU ou autre source documentée.",
    "H"
  );
  const headers = ["Type de végétation", "Espèce", "Âge moyen", "Facteur de séquestration", "Unité", "Source méthodologique", "Niveau d'incertitude", "Observations"];
  const rows = [
    ["Arbre urbain", "Espèce locale mixte", "Jeune", 0.010, "tCO2e/arbre/an", "Facteur municipal provisoire à remplacer par inventaire terrain", "Élevé", "À utiliser uniquement pour pré-estimation; distinguer des émissions directes."],
    ["Arbre urbain", "Espèce locale mixte", "Mature", 0.025, "tCO2e/arbre/an", "i-Tree Eco / équations biomasse à calibrer", "Moyen", "Requiert diamètre, hauteur, état sanitaire et espèce pour précision."],
    ["Forêt urbaine", "Mixte", "Mature", 3.500, "tCO2e/ha/an", "IPCC AFOLU - estimation locale", "Élevé", "Utiliser si la surface est connue mais pas l'inventaire arbre par arbre."],
    ["Mangrove / zone côtière", "Espèces halophiles", "Mature", 6.000, "tCO2e/ha/an", "IPCC wetland / étude locale requise", "Très élevé", "Inclure sols et biomasse uniquement avec protocole documenté."],
    ["Jardin maraîcher", "Mixte", "Annuel", 0.800, "tCO2e/ha/an", "Estimation prudente à documenter", "Très élevé", "Ne pas comptabiliser comme absorption vérifiée sans mesure locale."],
  ];
  sheet.getRange("A4:H4").values = [headers];
  sheet.getRange(`A5:H${4 + rows.length}`).values = rows;
  sheet.getRange("A4:H4").format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:H${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:H${4 + rows.length}`, "tblFacteursSequestrationCarbone");
  setWidths(sheet, [170, 170, 110, 160, 140, 270, 155, 340], 80);
  addValidation(sheet, 1, 5, 80, lists.vegetation);
  addValidation(sheet, 7, 5, 80, lists.uncertainty);
  sheet.getRange("D5:D80").setNumberFormat("#,##0.0000");
}

function buildNatureIndicators() {
  const sheet = addSheet(
    "Nature_Espaces_Verts",
    "Indicateurs Nature, espaces verts et solutions fondées sur la nature",
    "Indicateurs calculés automatiquement depuis la base validée Nature. Les absorptions carbone sont affichées séparément des émissions.",
    "H"
  );
  setWidths(sheet, [285, 150, 120, 230, 260, 130, 130, 220], 100);
  const rows = [
    ["Superficie totale d'espaces verts", "=SUM(Data_Nature_Espaces_Verts!$F$5:$F$124)", "unité déclarée", "Data_Nature_Espaces_Verts", "Convertir en m² dans le SIG pour comparaison officielle"],
    ["Superficie d'espaces verts par habitant", '=IFERROR(SUMIFS(Data_Nature_Espaces_Verts!$F$5:$F$124,Data_Nature_Espaces_Verts!$G$5:$G$124,"m²")/\'24_Parametres_Listes\'!$B$5,"")', "m²/hab", "Data_Nature + population paramètre", "Renseigner population de référence dans 24_Parametres_Listes"],
    ["Nombre total d'arbres plantés", "=SUM(Data_Nature_Espaces_Verts!$H$5:$H$124)", "arbres", "Data_Nature_Espaces_Verts", "CDP/C40/CDN adaptation"],
    ["Nombre total d'arbres survivants", "=SUM(Data_Nature_Espaces_Verts!$I$5:$I$124)", "arbres", "Data_Nature_Espaces_Verts", "Suivi de survie nécessaire"],
    ["Taux de survie des arbres", '=IFERROR(SUM(Data_Nature_Espaces_Verts!$I$5:$I$124)/SUM(Data_Nature_Espaces_Verts!$H$5:$H$124),"")', "%", "Data_Nature_Espaces_Verts", "Indicateur qualité du reboisement"],
    ["Nombre total de jardins", "=SUM(Data_Nature_Espaces_Verts!$K$5:$K$124)+SUM(Data_Nature_Espaces_Verts!$L$5:$L$124)", "jardins", "Data_Nature_Espaces_Verts", "Jardins publics + communautaires"],
    ["Nombre total de micro-jardins", "=SUM(Data_Nature_Espaces_Verts!$M$5:$M$124)", "micro-jardins", "Data_Nature_Espaces_Verts", "Agriculture urbaine"],
    ["Superficie reverdie annuelle", "=SUM(Data_Nature_Espaces_Verts!$N$5:$N$124)", "m²/ha selon unité source", "Data_Nature_Espaces_Verts", "À convertir en SIG pour publication"],
    ["Superficie restaurée annuelle", "=SUM(Data_Nature_Espaces_Verts!$O$5:$O$124)", "m²/ha selon unité source", "Data_Nature_Espaces_Verts", "Zones côtières restaurées incluses"],
    ["CO2 séquestré estimé", "=SUM(Data_Nature_Espaces_Verts!$Q$5:$Q$124)", "tCO2e/an estimées", "Data_Nature_Espaces_Verts", "Estimation séparée des émissions directes"],
    ["Population bénéficiaire", "=SUM(Data_Nature_Espaces_Verts!$R$5:$R$124)", "personnes", "Data_Nature_Espaces_Verts", "Co-bénéfices adaptation, santé, justice climatique"],
  ];
  sheet.getRange("A4:E4").values = [["Indicateur", "Valeur", "Unité", "Source", "Note de reporting"]];
  sheet.getRange(`A5:E${4 + rows.length}`).values = rows.map((r) => [r[0], null, r[2], r[3], r[4]]);
  sheet.getRange(`B5:B${4 + rows.length}`).formulas = rows.map((r) => [r[1]]);
  sheet.getRange("A4:E4").format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:E${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:E${4 + rows.length}`, "tblIndicateursNature");
  sheet.getRange(`B5:B${4 + rows.length}`).setNumberFormat("#,##0.00");
  sheet.getRange("B9").setNumberFormat("0%");

  const start = 20;
  sheet.getRange(`A${start}:F${start}`).values = [["Année", "Espaces verts", "Arbres plantés", "Taux survie", "Séquestration estimée", "Population bénéficiaire"]];
  sheet.getRange(`A${start + 1}:A${start + lists.years.length}`).values = lists.years.map((y) => [Number(y)]);
  const fRows = lists.years.map((_, i) => {
    const r = start + 1 + i;
    return [
      `=SUMIFS(Data_Nature_Espaces_Verts!$F$5:$F$124,Data_Nature_Espaces_Verts!$B$5:$B$124,A${r})`,
      `=SUMIFS(Data_Nature_Espaces_Verts!$H$5:$H$124,Data_Nature_Espaces_Verts!$B$5:$B$124,A${r})`,
      `=IFERROR(SUMIFS(Data_Nature_Espaces_Verts!$I$5:$I$124,Data_Nature_Espaces_Verts!$B$5:$B$124,A${r})/SUMIFS(Data_Nature_Espaces_Verts!$H$5:$H$124,Data_Nature_Espaces_Verts!$B$5:$B$124,A${r}),0)`,
      `=SUMIFS(Data_Nature_Espaces_Verts!$Q$5:$Q$124,Data_Nature_Espaces_Verts!$B$5:$B$124,A${r})`,
      `=SUMIFS(Data_Nature_Espaces_Verts!$R$5:$R$124,Data_Nature_Espaces_Verts!$B$5:$B$124,A${r})`,
    ];
  });
  sheet.getRange(`B${start + 1}:F${start + lists.years.length}`).formulas = fRows;
  sheet.getRange(`A${start}:F${start}`).format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${start + 1}:F${start + lists.years.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A${start}:F${start + lists.years.length}`, "tblEvolutionNature", "TableStyleMedium2");
  sheet.getRange(`D${start + 1}:D${start + lists.years.length}`).setNumberFormat("0%");
  sheet.freezePanes.freezeRows(HEADER_ROW);
}

function updateGESModule() {
  const gesForm = workbook.worksheets.getItemOrNullObject?.("06_Form_Emissions_GES");
  if (gesForm && !gesForm.isNullObject) {
    const headers = ["Absorptions / séquestration carbone estimée", "Méthodologie séquestration", "Facteur séquestration", "Source facteur", "Distinction émissions/absorptions"];
    gesForm.getRange("R4:V4").values = [headers];
    gesForm.getRange("R4:V4").format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
    const formulas = [];
    for (let row = DATA_START; row <= DATA_START + 99; row += 1) {
      formulas.push([`=IF(B${row}="","",IF(C${row}="AFOLU",H${row},""))`, "", "", "", `=IF(R${row}<>"","Absorption estimée séparée - ne pas déduire des émissions brutes","Émission directe ou donnée d'activité")`]);
    }
    gesForm.getRange(`R${DATA_START}:V${DATA_START + 99}`).formulas = formulas;
    gesForm.getRange(`R${DATA_START}:R${DATA_START + 99}`).setNumberFormat("#,##0.000");
    setWidths(gesForm, [110, 90, 170, 170, 145, 95, 135, 140, 160, 145, 210, 160, 250, 120, 150, 190, 120, 190, 190, 150, 210, 280], DATA_START + 105);
  }
  const gesData = workbook.worksheets.getItemOrNullObject?.("Data_Emissions_GES");
  if (gesData && !gesData.isNullObject) {
    const headers = ["Absorptions / séquestration carbone estimée", "Méthodologie séquestration", "Facteur séquestration", "Source facteur", "Distinction émissions/absorptions"];
    gesData.getRange("R4:V4").values = [headers];
    gesData.getRange("R4:V4").format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
    const formulas = [];
    for (let row = DATA_START; row <= DATA_START + 99; row += 1) {
      formulas.push([`=IF(${q("06_Form_Emissions_GES")}!$L${row}="${VALID_STATUS}",${q("06_Form_Emissions_GES")}!R${row},"")`, `=IF(${q("06_Form_Emissions_GES")}!$L${row}="${VALID_STATUS}",${q("06_Form_Emissions_GES")}!S${row},"")`, `=IF(${q("06_Form_Emissions_GES")}!$L${row}="${VALID_STATUS}",${q("06_Form_Emissions_GES")}!T${row},"")`, `=IF(${q("06_Form_Emissions_GES")}!$L${row}="${VALID_STATUS}",${q("06_Form_Emissions_GES")}!U${row},"")`, `=IF(${q("06_Form_Emissions_GES")}!$L${row}="${VALID_STATUS}",${q("06_Form_Emissions_GES")}!V${row},"")`]);
    }
    gesData.getRange(`R${DATA_START}:V${DATA_START + 99}`).formulas = formulas;
    gesData.getRange(`R${DATA_START}:R${DATA_START + 99}`).setNumberFormat("#,##0.000");
  }
  const inv = workbook.worksheets.getItemOrNullObject?.("Inventaire_GES");
  if (inv && !inv.isNullObject) {
    inv.getRange("A16:E16").values = [["Module absorptions carbone nature", "Valeur", "Unité", "Source", "Règle"]];
    inv.getRange("A17:E20").values = [
      ["Émissions brutes suivies", null, "tCO2e", "Data_Emissions_GES + modules sectoriels", "Ne pas mélanger aux absorptions"],
      ["Absorptions / séquestration estimées", null, "tCO2e/an", "Data_Nature_Espaces_Verts + Facteurs_Sequestration_Carbone", "Estimation séparée"],
      ["Solde indicatif séparé", null, "tCO2e", "Calcul indicatif", "À ne pas reporter comme net officiel sans méthode validée"],
      ["Méthodologie requise", "Oui", "texte", "Source facteur", "Obligatoire pour toute estimation de séquestration"],
    ];
    inv.getRange("B17:B19").formulas = [["=SUM(B5:B8)"], ["=SUM(Data_Nature_Espaces_Verts!$Q$5:$Q$124)"], ["=B17-B18"]];
    inv.getRange("A16:E16").format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
    inv.getRange("A17:E20").format = { fill: palette.white, wrapText: true };
    inv.getRange("B17:B19").setNumberFormat("#,##0.00");
  }
}

function buildDynamicDashboard() {
  const sheet = addSheet(
    "20_Tableau_de_Bord_Dynamique",
    "Tableau de bord dynamique V4B - Nature, climat et résilience",
    "Vue exécutive intégrant nature, espaces verts, arbres, jardins, biodiversité urbaine et séquestration estimée. Les données non validées humainement sont exclues.",
    "N"
  );
  setWidths(sheet, [245, 145, 125, 240, 245, 145, 125, 240, 245, 145, 125, 240, 160, 260], 90);
  sheet.getRange("A4:N4").values = [["KPI Nature", "Valeur", "Unité", "Source", "KPI climat", "Valeur", "Unité", "Source", "Contribution reporting", "Statut", "Cadre", "Source", "Priorité", "Lecture"]];
  sheet.getRange("A4:N4").format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
  const left = [
    ["Superficie totale d'espaces verts", "=Nature_Espaces_Verts!B5", "unité source", "Data_Nature"],
    ["Espaces verts par habitant", "=Nature_Espaces_Verts!B6", "m²/hab", "Paramètres + Data_Nature"],
    ["Nombre d'arbres plantés", "=Nature_Espaces_Verts!B7", "arbres", "Data_Nature"],
    ["Nombre d'arbres survivants", "=Nature_Espaces_Verts!B8", "arbres", "Data_Nature"],
    ["Taux de survie", "=Nature_Espaces_Verts!B9", "%", "Data_Nature"],
    ["Nombre de jardins", "=Nature_Espaces_Verts!B10", "jardins", "Data_Nature"],
    ["Nombre de micro-jardins", "=Nature_Espaces_Verts!B11", "micro-jardins", "Data_Nature"],
    ["Superficie reverdie", "=Nature_Espaces_Verts!B12", "unité source", "Data_Nature"],
    ["CO2 séquestré estimé", "=Nature_Espaces_Verts!B14", "tCO2e/an estimées", "Data_Nature"],
  ];
  const middle = [
    ["Émissions totales brutes", "=Tableau_de_bord!B6", "tCO2e", "Inventaire_GES"],
    ["Absorptions estimées séparées", "=Nature_Espaces_Verts!B14", "tCO2e/an", "Nature"],
    ["Population bénéficiaire nature", "=Nature_Espaces_Verts!B15", "personnes", "Nature"],
    ["Données nature validées", `=COUNTIF(Data_Nature_Espaces_Verts!$V$${DATA_START}:$V$${DATA_END},"${VALID_STATUS}")`, "lignes", "Data_Nature"],
    ["Alertes nature", `=COUNTIF(${q("10_Form_Nature_Espaces_Verts")}!$AD$${DATA_START}:$AD$${DATA_END},"ALERTE*")`, "alertes", "Form_Nature"],
    ["Facteurs avec incertitude élevée", '=COUNTIF(Facteurs_Sequestration_Carbone!$G$5:$G$80,"Élevé")+COUNTIF(Facteurs_Sequestration_Carbone!$G$5:$G$80,"Très élevé")', "facteurs", "Facteurs"],
    ["Documents locaux indexés", `=${documentRecords.length}`, "documents", "Registre documentaire"],
    ["Contributions CDN nature", '=IF(B13>0,"Oui - adaptation/AFOLU urbain","À documenter")', "statut", "CDN"],
    ["Contribution nature-based solutions", '=IF(B12>0,"Oui","À collecter")', "statut", "C40/COMSSA"],
  ];
  const reporting = [
    ["Contribution adaptation", '=IF(B15>0,"Oui","À collecter")', "CDP/C40/CDN", "Population bénéficiaire et restauration", "P0", "Résilience, chaleur, inondation, littoral"],
    ["Contribution CDP", '=IF(B5>0,"Oui","À collecter")', "CDP-ICLEI", "Espaces verts, adaptation, biodiversité", "P0", "Soutient disclosure risques/actions"],
    ["Contribution C40", '=IF(B12>0,"Oui","À collecter")', "C40 Urban Nature", "Greening, accès, résilience", "P0", "Suit ambition nature et résilience"],
    ["Contribution CDN", '=IF(B14>0,"Oui - estimation","À documenter")', "CDN Sénégal", "Adaptation, AFOLU urbain, co-bénéfices", "P0", "Ne pas comptabiliser en net sans méthode validée"],
    ["Contribution bailleurs climat", '=IF(B15>0,"Oui","À documenter")', "GCF/WB/BAD", "Bénéficiaires, E&S, genre, résilience", "P1", "Renforce bancabilité"],
    ["Contribution ICLEI", '=IF(B5>0,"Oui","À collecter")', "ICLEI/CitiesWithNature", "Nature urbaine, biodiversité", "P1", "Compatible suivi CitiesWithNature"],
    ["Contribution COMSSA", '=IF(B12>0,"Oui","À collecter")', "SEACAP", "Mesures vertes/bleues d'adaptation", "P1", "Alimente le pilier adaptation"],
    ["Qualité des estimations", '=IF(F9>0,"À améliorer","OK")', "MRV", "Facteurs et incertitudes", "P0", "Remplacer facteurs provisoires par étude locale"],
    ["Validation humaine", '=IF(F8=0,"À lancer","OK")', "MRV", "Statut validation", "P0", "Aucune donnée non validée ne part en reporting"],
  ];
  for (let i = 0; i < left.length; i += 1) {
    const r = 5 + i;
    sheet.getRange(`A${r}:D${r}`).values = [[left[i][0], null, left[i][2], left[i][3]]];
    sheet.getRange(`B${r}`).formulas = [[left[i][1]]];
    sheet.getRange(`E${r}:H${r}`).values = [[middle[i][0], null, middle[i][2], middle[i][3]]];
    sheet.getRange(`F${r}`).formulas = [[middle[i][1]]];
    sheet.getRange(`I${r}:N${r}`).values = [[reporting[i][0], null, reporting[i][2], reporting[i][3], reporting[i][4], reporting[i][5]]];
    sheet.getRange(`J${r}`).formulas = [[reporting[i][1]]];
  }
  sheet.getRange("A5:N13").format = { fill: palette.white, wrapText: true };
  sheet.getRange("B5:B13").setNumberFormat("#,##0.00");
  sheet.getRange("B9").setNumberFormat("0%");
  sheet.getRange("F5:F13").setNumberFormat("#,##0.00");
  addStatusFormatting(sheet, "J5:J13");
  addStatusFormatting(sheet, "M5:M13");
  addTable(sheet, "A4:N13", "tblDashboardNatureV4B", "TableStyleMedium4");
}

function buildCharts() {
  const sheet = addSheet(
    "19_Graphiques_Dynamiques",
    "Graphiques dynamiques V4B - Nature et espaces verts",
    "Graphiques natifs Excel alimentés par les données nature validées. Les valeurs restent à zéro tant que les formulaires ne sont pas validés.",
    "V"
  );
  setWidths(sheet, [100, 130, 20, 130, 130, 20, 130, 130, 20, 130, 130, 20, 130, 130, 20, 130, 130, 20, 130, 130, 130, 130], 120);

  const blocks = [
    { title: "Évolution annuelle des espaces verts", row: 4, col: 1, type: "line", labels: lists.years, header: "Superficie", formula: (r, c) => `=SUMIFS(Data_Nature_Espaces_Verts!$F$5:$F$124,Data_Nature_Espaces_Verts!$B$5:$B$124,${c}${r})`, pos: ["D4", "I18"], fmt: "#,##0" },
    { title: "Évolution annuelle des arbres plantés", row: 20, col: 1, type: "bar", labels: lists.years, header: "Arbres", formula: (r, c) => `=SUMIFS(Data_Nature_Espaces_Verts!$H$5:$H$124,Data_Nature_Espaces_Verts!$B$5:$B$124,${c}${r})`, pos: ["D20", "I34"], fmt: "#,##0" },
    { title: "Évolution du taux de survie", row: 36, col: 1, type: "line", labels: lists.years, header: "Taux", formula: (r, c) => `=IFERROR(SUMIFS(Data_Nature_Espaces_Verts!$I$5:$I$124,Data_Nature_Espaces_Verts!$B$5:$B$124,${c}${r})/SUMIFS(Data_Nature_Espaces_Verts!$H$5:$H$124,Data_Nature_Espaces_Verts!$B$5:$B$124,${c}${r}),0)`, pos: ["D36", "I50"], fmt: "0%" },
    { title: "Évolution des jardins communautaires", row: 4, col: 11, type: "bar", labels: lists.years, header: "Jardins", formula: (r, c) => `=SUMIFS(Data_Nature_Espaces_Verts!$L$5:$L$124,Data_Nature_Espaces_Verts!$B$5:$B$124,${c}${r})`, pos: ["N4", "S18"], fmt: "#,##0" },
    { title: "Évolution des micro-jardins", row: 20, col: 11, type: "bar", labels: lists.years, header: "Micro-jardins", formula: (r, c) => `=SUMIFS(Data_Nature_Espaces_Verts!$M$5:$M$124,Data_Nature_Espaces_Verts!$B$5:$B$124,${c}${r})`, pos: ["N20", "S34"], fmt: "#,##0" },
    { title: "Séquestration carbone estimée", row: 36, col: 11, type: "line", labels: lists.years, header: "tCO2e estimées", formula: (r, c) => `=SUMIFS(Data_Nature_Espaces_Verts!$Q$5:$Q$124,Data_Nature_Espaces_Verts!$B$5:$B$124,${c}${r})`, pos: ["N36", "S50"], fmt: "#,##0.0" },
  ];

  for (const b of blocks) {
    const c1 = colName(b.col);
    const c2 = colName(b.col + 1);
    sheet.getRange(`${c1}${b.row}:${c2}${b.row}`).merge();
    sheet.getRange(`${c1}${b.row}`).values = [[b.title]];
    sheet.getRange(`${c1}${b.row}:${c2}${b.row}`).format = { fill: palette.navy, font: { bold: true, color: palette.white }, wrapText: true };
    sheet.getRange(`${c1}${b.row + 1}:${c2}${b.row + 1}`).values = [["Année", b.header]];
    sheet.getRange(`${c1}${b.row + 2}:${c1}${b.row + 1 + b.labels.length}`).values = b.labels.map((x) => [Number(x)]);
    sheet.getRange(`${c2}${b.row + 2}:${c2}${b.row + 1 + b.labels.length}`).formulas = b.labels.map((_, i) => [b.formula(b.row + 2 + i, c1)]);
    sheet.getRange(`${c1}${b.row + 1}:${c2}${b.row + 1}`).format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
    sheet.getRange(`${c1}${b.row + 2}:${c2}${b.row + 1 + b.labels.length}`).format = { fill: palette.white, wrapText: true };
    sheet.getRange(`${c2}${b.row + 2}:${c2}${b.row + 1 + b.labels.length}`).setNumberFormat(b.fmt);
    const chart = sheet.charts.add(b.type, sheet.getRange(`${c1}${b.row + 1}:${c2}${b.row + 1 + b.labels.length}`));
    chart.title = b.title;
    chart.hasLegend = false;
    chart.xAxis = { axisType: "textAxis" };
    chart.yAxis = { numberFormatCode: b.fmt };
    chart.setPosition(b.pos[0], b.pos[1]);
  }

  const row = 58;
  sheet.getRange(`A${row}:B${row}`).merge();
  sheet.getRange(`A${row}`).values = [["Espaces verts par commune"]];
  sheet.getRange(`A${row}:B${row}`).format = { fill: palette.navy, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${row + 1}:B${row + 1}`).values = [["Commune", "Superficie"]];
  sheet.getRange(`A${row + 2}:A${row + 1 + lists.communes.length}`).values = lists.communes.map((x) => [x]);
  sheet.getRange(`B${row + 2}:B${row + 1 + lists.communes.length}`).formulas = lists.communes.map((_, i) => [`=SUMIFS(Data_Nature_Espaces_Verts!$F$5:$F$124,Data_Nature_Espaces_Verts!$D$5:$D$124,A${row + 2 + i})`]);
  sheet.getRange(`A${row + 1}:B${row + 1}`).format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${row + 2}:B${row + 1 + lists.communes.length}`).format = { fill: palette.white, wrapText: true };
  const chart = sheet.charts.add("bar", sheet.getRange(`A${row + 1}:B${row + 1 + lists.communes.length}`));
  chart.title = "Espaces verts par commune";
  chart.hasLegend = false;
  chart.xAxis = { axisType: "textAxis" };
  chart.yAxis = { numberFormatCode: "#,##0" };
  chart.setPosition("D58", "K78");
}

function buildQualityControl() {
  const sheet = addSheet(
    "21_Controle_Qualite",
    "Contrôle qualité V4B - Nature et séquestration",
    "Alertes qualité sur superficie/unité, survie des arbres, facteur méthodologique, localisation, source, document et validation humaine.",
    "K"
  );
  setWidths(sheet, [250, 130, 130, 135, 135, 145, 145, 145, 145, 135, 300], 80);
  const rows = [
    ["Nature - lignes saisies", `=COUNTIF(${q("10_Form_Nature_Espaces_Verts")}!$B$5:$B$124,"<>")`, "", "", "", "", "", "", "", '=IF(B5=0,"À collecter","OK")', "Lancer la saisie espaces verts et arbres"],
    ["Superficie sans unité", `=COUNTIFS(${q("10_Form_Nature_Espaces_Verts")}!$F$5:$F$124,"<>",${q("10_Form_Nature_Espaces_Verts")}!$G$5:$G$124,"")`, "", "", "", "", "", "", "", '=IF(B6>0,"ALERTE","OK")', "Renseigner m², hectares ou km²"],
    ["Arbres plantés sans taux de survie", `=COUNTIFS(${q("10_Form_Nature_Espaces_Verts")}!$H$5:$H$124,"<>",${q("10_Form_Nature_Espaces_Verts")}!$J$5:$J$124,"")`, "", "", "", "", "", "", "", '=IF(B7>0,"ALERTE","OK")', "Saisir arbres survivants pour calculer le taux"],
    ["Séquestration sans facteur méthodologique", `=COUNTIFS(${q("10_Form_Nature_Espaces_Verts")}!$Q$5:$Q$124,"<>",${q("10_Form_Nature_Espaces_Verts")}!$Y$5:$Y$124,"")+COUNTIFS(${q("10_Form_Nature_Espaces_Verts")}!$Q$5:$Q$124,"<>",${q("10_Form_Nature_Espaces_Verts")}!$Z$5:$Z$124,"")`, "", "", "", "", "", "", "", '=IF(B8>0,"ALERTE","OK")', "Documenter facteur, source et méthodologie"],
    ["Donnée nature sans localisation", `=COUNTIFS(${q("10_Form_Nature_Espaces_Verts")}!$B$5:$B$124,"<>",${q("10_Form_Nature_Espaces_Verts")}!$D$5:$D$124,"")`, "", "", "", "", "", "", "", '=IF(B9>0,"ALERTE","OK")', "Renseigner commune ou zone"],
    ["Donnée nature sans source", `=COUNTIFS(${q("10_Form_Nature_Espaces_Verts")}!$B$5:$B$124,"<>",${q("10_Form_Nature_Espaces_Verts")}!$T$5:$T$124,"")`, "", "", "", "", "", "", "", '=IF(B10>0,"ALERTE","OK")', "Renseigner institution source"],
    ["Document source absent", `=COUNTIFS(${q("10_Form_Nature_Espaces_Verts")}!$B$5:$B$124,"<>",${q("10_Form_Nature_Espaces_Verts")}!$U$5:$U$124,"")`, "", "", "", "", "", "", "", '=IF(B11>0,"ALERTE","OK")', "Joindre référence documentaire"],
    ["Donnée non validée", `=COUNTIFS(${q("10_Form_Nature_Espaces_Verts")}!$B$5:$B$124,"<>",${q("10_Form_Nature_Espaces_Verts")}!$V$5:$V$124,"<>Validé humainement")`, "", "", "", "", "", "", "", '=IF(B12>0,"À corriger","OK")', "Valider humainement avant reporting"],
    ["Alertes QC nature", `=COUNTIF(${q("10_Form_Nature_Espaces_Verts")}!$AD$5:$AD$124,"ALERTE*")`, "", "", "", "", "", "", "", '=IF(B13>0,"ALERTE","OK")', "Corriger chaque ligne en anomalie"],
  ];
  sheet.getRange("A4:K4").values = [["Contrôle", "Nombre", "CDP", "C40", "ICLEI", "COMSSA", "CDN", "Bailleurs", "Blocage reporting", "Statut", "Action corrective"]];
  sheet.getRange(`A5:K${4 + rows.length}`).values = rows;
  sheet.getRange(`B5:B${4 + rows.length}`).formulas = rows.map((r) => [r[1]]);
  sheet.getRange(`J5:J${4 + rows.length}`).formulas = rows.map((r) => [r[9]]);
  sheet.getRange(`C5:I${4 + rows.length}`).values = rows.map((_, i) => (i === 0 ? ["Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "Non"] : ["Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui"]));
  sheet.getRange("A4:K4").format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:K${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:K${4 + rows.length}`, "tblControleQualiteNature");
  addStatusFormatting(sheet, `J5:J${4 + rows.length}`);
}

function buildCollectionPlan() {
  const sheet = addSheet(
    "22_Planning_Collecte",
    "Planning de collecte V4B",
    "Plan de collecte opérationnel intégrant nature, espaces verts, arbres, jardins, biodiversité et séquestration estimée.",
    "J"
  );
  const rows = [
    ["P0", "Inventaire espaces verts géolocalisé", "SIG Ville, Direction Espaces Verts, PCET", "Cellule SIG / Espaces verts", "Trimestrielle", "Élevée", "CDP/C40/CDN", "Créer couche SIG et identifiant par site", "30 jours", "À lancer"],
    ["P0", "Arbres plantés et survivants", "Campagnes reboisement, projets, écoles, marchés", "Direction Espaces Verts", "Mensuelle", "Moyenne", "C40/CDP/CDN", "Mettre à jour plantés/survivants par zone", "30 jours", "À lancer"],
    ["P0", "Facteurs de séquestration", "i-Tree/IPCC/étude locale", "Cellule MRV + université", "Annuelle", "Très élevée", "CDP/CDN/Bailleurs", "Valider source méthodologique et incertitude", "60 jours", "À lancer"],
    ["P1", "Jardins publics et communautaires", "Directions municipales, associations", "Direction Environnement", "Trimestrielle", "Moyenne", "ICLEI/COMSSA", "Recenser jardins et bénéficiaires", "60 jours", "À lancer"],
    ["P1", "Micro-jardins et agriculture urbaine", "Programmes municipaux, ONG", "Affaires sociales / Environnement", "Semestrielle", "Moyenne", "CDN/COMSSA", "Recenser sites et superficies", "90 jours", "À lancer"],
    ["P1", "Zones côtières restaurées", "CSE, universités, projets littoraux", "Planification / SIG", "Semestrielle", "Élevée", "CDN/Bailleurs", "Cartographier restauration et population bénéficiaire", "90 jours", "À lancer"],
    ["P2", "Biodiversité urbaine", "Universités, ONG, inventaires locaux", "Direction Environnement", "Annuelle", "Élevée", "ICLEI/CDP", "Lister espèces plantées et espèces prioritaires", "120 jours", "À lancer"],
  ];
  sheet.getRange("A4:J4").values = [["Priorité", "Donnée à collecter", "Source", "Responsable", "Fréquence", "Difficulté", "Cadres critiques", "Première action", "Horizon", "Statut"]];
  sheet.getRange(`A5:J${4 + rows.length}`).values = rows;
  sheet.getRange("A4:J4").format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:J${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:J${4 + rows.length}`, "tblPlanningCollecteNature");
  setWidths(sheet, [80, 240, 250, 210, 135, 130, 160, 310, 100, 120], 80);
  addValidation(sheet, 1, 5, 80, lists.priorities);
  addValidation(sheet, 5, 5, 80, lists.frequencies);
  addStatusFormatting(sheet, `A5:A${4 + rows.length}`);
}

function buildDocumentRegistry() {
  const sheet = addSheet(
    "23_Registre_Documentaire",
    "Registre documentaire V4B",
    "Documents locaux fournis et mapping vers les modules du Hub. Ce registre sert de base de preuves pour les formulaires, le contrôle qualité et le reporting.",
    "K"
  );
  const localRows = documentRecords.map((d) => [
    d.file_name,
    d.extension,
    d.size_kb,
    (d.domains ?? []).join("; "),
    (d.target_sheets ?? []).join("; "),
    d.priority,
    d.extraction_status,
    d.use_case,
    d.path,
    (d.domains ?? []).some((x) => ["adaptation", "gouvernance", "ges", "projets"].includes(x)) ? "Nature/résilience possible" : "À analyser",
    "À exploiter pour fiches sources et preuves",
  ]);
  const officialRows = [
    ["C40 Urban Nature Campaign / Accelerator", "web", "", "nature; adaptation; reporting", "10_Form_Nature_Espaces_Verts; 20_Tableau_de_Bord_Dynamique; 19_Graphiques_Dynamiques", "P0", "Référence externe", "Cadrer greening, accès aux espaces verts et résilience urbaine", "https://www.c40.org/urban-nature-campaign/", "Nature/résilience", "Référence à citer dans reporting C40"],
    ["ICLEI CitiesWithNature", "web", "", "nature; biodiversité; reporting", "10_Form_Nature_Espaces_Verts; 23_Registre_Documentaire", "P1", "Référence externe", "Aligner la biodiversité urbaine et les solutions fondées sur la nature", "https://cbc.iclei.org/iclei/initiative/citieswithnature-regionswithnature/", "Biodiversité", "Référence à citer pour ICLEI"],
    ["i-Tree / USDA Forest Service", "web", "", "séquestration; arbres; méthodologie", "Facteurs_Sequestration_Carbone; Data_Nature_Espaces_Verts", "P0", "Référence externe", "Méthode d'estimation des bénéfices des arbres urbains", "https://research.fs.usda.gov/products/dataandtools/i-tree", "Séquestration", "Utiliser pour remplacer les facteurs provisoires"],
    ["CoM SSA SEACAP Guidebook", "web", "", "COMSSA; adaptation; nature-based solutions", "22_Planning_Collecte; Contribution_CDN_2_0", "P1", "Référence externe", "Cadrer mesures d'adaptation et infrastructures vertes/bleues", "https://comssa.org/en/site-resources/seacap-guidebook-extended-version", "Nature/résilience", "Référence SEACAP"],
  ];
  const rows = [...localRows, ...officialRows];
  sheet.getRange("A4:K4").values = [["Document", "Type", "Taille KB", "Domaines détectés", "Feuilles cibles", "Priorité", "Statut extraction", "Usage pour le Hub", "Chemin / lien", "Lien Nature", "Action"]];
  sheet.getRange(`A5:K${4 + rows.length}`).values = rows;
  sheet.getRange("A4:K4").format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:K${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:K${4 + rows.length}`, "tblRegistreDocumentaireV4B");
  setWidths(sheet, [310, 80, 90, 180, 300, 80, 135, 360, 420, 160, 260], rows.length + 20);
  addValidation(sheet, 6, 5, 5 + rows.length, lists.priorities);
  addStatusFormatting(sheet, `F5:F${4 + rows.length}`);
}

function buildParameters() {
  const sheet = addSheet(
    "24_Parametres_Listes",
    "Paramètres et listes contrôlées V4B",
    "Listes déroulantes, paramètres population et vocabulaire normalisé pour le module Nature.",
    "L"
  );
  setWidths(sheet, [235, 170, 220, 180, 220, 190, 180, 210, 190, 180, 190, 220], 100);
  sheet.getRange("A4:B8").values = [
    ["Paramètre", "Valeur"],
    ["Population de référence Ville de Dakar", 1200000],
    ["Unité par défaut espaces verts", "m²"],
    ["Statut requis reporting", VALID_STATUS],
    ["Note", "Population indicative à remplacer par source officielle ANSD/Ville"],
  ];
  sheet.getRange("A4:B4").format = { fill: palette.leaf, font: { bold: true, color: palette.white } };
  sheet.getRange("A5:B8").format = { fill: palette.white, wrapText: true };
  const blocks = [
    ["D", "Types espaces verts", lists.greenTypes],
    ["F", "Unités", lists.greenUnits],
    ["H", "Méthodologies", lists.methodologies],
    ["J", "Incertitude", lists.uncertainty],
    ["L", "Institutions", lists.institutions],
  ];
  for (const [col, title, values] of blocks) {
    sheet.getRange(`${col}4`).values = [[title]];
    sheet.getRange(`${col}4`).format = { fill: palette.leaf, font: { bold: true, color: palette.white }, wrapText: true };
    sheet.getRange(`${col}5:${col}${4 + values.length}`).values = values.map((v) => [v]);
    sheet.getRange(`${col}5:${col}${4 + values.length}`).format = { fill: palette.white, wrapText: true };
  }
}

function updateAccueil() {
  const sheet = workbook.worksheets.getItemOrNullObject?.("00_ACCUEIL");
  if (!sheet || sheet.isNullObject) return;
  sheet.getRange("A30:J30").merge();
  sheet.getRange("A30").values = [["Menu V4B Nature - espaces verts, arbres, jardins, biodiversité et séquestration estimée"]];
  sheet.getRange("A30:J30").format = { fill: palette.leaf, font: { bold: true, color: palette.white, size: 13 } };
  const links = [
    ["Saisir Nature", "#'10_Form_Nature_Espaces_Verts'!A1"],
    ["Base Nature validée", "#'Data_Nature_Espaces_Verts'!A1"],
    ["Facteurs séquestration", "#'Facteurs_Sequestration_Carbone'!A1"],
    ["Dashboard V4B", "#'20_Tableau_de_Bord_Dynamique'!A1"],
    ["Graphiques Nature", "#'19_Graphiques_Dynamiques'!A1"],
    ["Contrôle qualité V4B", "#'21_Controle_Qualite'!A1"],
    ["Planning collecte", "#'22_Planning_Collecte'!A1"],
    ["Registre documentaire", "#'23_Registre_Documentaire'!A1"],
  ];
  const rows = [];
  for (let i = 0; i < links.length; i += 2) rows.push([links[i][0], links[i][1], links[i + 1]?.[0] ?? "", links[i + 1]?.[1] ?? ""]);
  sheet.getRange(`A32:D${31 + rows.length}`).values = rows;
  for (let row = 32; row < 32 + rows.length; row += 1) {
    sheet.getRange(`A${row}`).formulas = [[`=HYPERLINK(B${row},A${row})`]];
    sheet.getRange(`C${row}`).formulas = [[`=HYPERLINK(D${row},C${row})`]];
  }
  sheet.getRange(`A32:D${31 + rows.length}`).format = { fill: palette.lightGreen, font: { color: palette.navy, bold: true }, wrapText: true };
}

buildNatureForm();
buildNatureData();
buildSequestrationFactors();
buildNatureIndicators();
updateGESModule();
buildDynamicDashboard();
buildCharts();
buildQualityControl();
buildCollectionPlan();
buildDocumentRegistry();
buildParameters();
updateAccueil();

await fs.mkdir(outputDir, { recursive: true });

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  maxChars: 3000,
});
console.log("FORMULA_ERRORS");
console.log(formulaErrors.ndjson);

const dashboardInspect = await workbook.inspect({
  kind: "table",
  range: "20_Tableau_de_Bord_Dynamique!A4:N13",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 14,
});
console.log("DASHBOARD_CHECK");
console.log(dashboardInspect.ndjson);

for (const renderTarget of [
  ["10_Form_Nature_Espaces_Verts", "A1:AE18"],
  ["Data_Nature_Espaces_Verts", "A1:AE18"],
  ["Facteurs_Sequestration_Carbone", "A1:H12"],
  ["20_Tableau_de_Bord_Dynamique", "A1:N16"],
  ["19_Graphiques_Dynamiques", "A1:V52"],
  ["21_Controle_Qualite", "A1:K16"],
  ["22_Planning_Collecte", "A1:J14"],
  ["23_Registre_Documentaire", "A1:K16"],
  ["24_Parametres_Listes", "A1:L24"],
]) {
  const [sheetName, range] = renderTarget;
  const rendered = await workbook.render({ sheetName, range, scale: 1.15, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), new Uint8Array(await rendered.arrayBuffer()));
  console.log(`RENDERED ${sheetName}`);
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputFile);

const outputBlob = await FileBlob.load(outputFile);
const verifyWorkbook = await SpreadsheetFile.importXlsx(outputBlob);
const verifyErrors = await verifyWorkbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  maxChars: 3000,
});
console.log("VERIFY_FORMULA_ERRORS");
console.log(verifyErrors.ndjson);

const verifyWorkbookInfo = await verifyWorkbook.inspect({
  kind: "workbook",
  include: "sheets",
  maxChars: 7000,
});
console.log("VERIFY_WORKBOOK");
console.log(verifyWorkbookInfo.ndjson);
console.log(`SAVED ${outputFile}`);
