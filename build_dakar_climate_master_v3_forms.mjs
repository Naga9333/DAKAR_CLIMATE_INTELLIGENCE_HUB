import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputFile = path.resolve("outputs", "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V2.xlsx");
const outputDir = path.resolve("outputs");
const outputFile = path.join(outputDir, "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V3_FORMS.xlsx");

const input = await FileBlob.load(inputFile);
const workbook = await SpreadsheetFile.importXlsx(input);

const palette = {
  navy: "#0B1F33",
  teal: "#0F766E",
  green: "#166534",
  amber: "#F59E0B",
  red: "#DC2626",
  blue: "#2563EB",
  lightTeal: "#E6F4F1",
  lightBlue: "#EAF2FF",
  lightGreen: "#EAF7EA",
  lightAmber: "#FFF7E6",
  lightRed: "#FEE2E2",
  white: "#FFFFFF",
  slate: "#334155",
  grey: "#F8FAFC",
};

const VALID_STATUS = "Validé humainement";
const FORM_ROWS = 100;
const HEADER_ROW = 4;
const DATA_START = 5;
const DATA_END = DATA_START + FORM_ROWS - 1;

const lists = {
  domaines: [
    "Climat",
    "Énergie",
    "Déchets",
    "Mobilité",
    "Qualité de l'air",
    "GES",
    "Adaptation / Résilience",
    "Érosion côtière",
    "Finance climat",
    "Projets climat",
  ],
  sousDomaines: [
    "Température",
    "Pluviométrie",
    "Électricité",
    "Déchets solides",
    "Qualité de l'air",
    "Inventaire GES",
    "Transport",
    "Érosion côtière",
    "Risques climatiques",
    "Budget et financement",
  ],
  mois: [
    "Janvier",
    "Février",
    "Mars",
    "Avril",
    "Mai",
    "Juin",
    "Juillet",
    "Août",
    "Septembre",
    "Octobre",
    "Novembre",
    "Décembre",
  ],
  annees: ["2024", "2025", "2026", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "2034", "2035"],
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
  institutions: [
    "Ville de Dakar",
    "Direction Environnement",
    "Direction Mobilité",
    "Direction Énergie",
    "Direction Déchets",
    "Direction Finances",
    "ANACIM",
    "DEEC / CGQA",
    "SENELEC",
    "SONAGED / UCG",
    "CETUD",
    "ONAS",
    "ANSD",
    "CSE",
    "Université / Recherche",
    "Bailleur",
    "ONG / Association",
    "Prestataire",
  ],
  services: [
    "Cabinet du Maire",
    "Cellule climat",
    "Cellule MRV",
    "Planification",
    "Environnement",
    "Mobilité",
    "Énergie",
    "Déchets",
    "Finances",
    "Marchés publics",
    "SIG",
    "Affaires sociales",
  ],
  frequences: ["Mensuelle", "Trimestrielle", "Semestrielle", "Annuelle", "Ponctuelle", "Continu"],
  fiabilite: ["Faible", "Moyenne", "Élevée", "Très élevée"],
  statutsValidation: ["Brouillon", "À vérifier", "Validé techniquement", VALID_STATUS, "Rejeté"],
  ouiNon: ["Oui", "Non"],
  unites: ["tCO2e", "kgCO2e", "kWh", "MWh", "tonnes", "kg", "mm", "°C", "µg/m3", "km", "m", "FCFA", "%", "nombre"],
  secteursGes: ["Énergie stationnaire", "Transport", "Déchets", "IPPU", "AFOLU", "Autre"],
  typesConso: ["Bâtiments municipaux", "Éclairage public", "Marchés", "Équipements sportifs", "Santé", "Éducation", "Autre"],
  typesDechets: ["Ordures ménagères", "Plastiques", "Papier/carton", "Organique", "Déchets verts", "Déchets dangereux", "Autre"],
  modesTransport: ["Bus", "BRT", "Taxi", "Véhicule particulier", "Deux-roues", "Marche", "Vélo", "Transport informel", "Flotte municipale"],
  niveauxRisque: ["Faible", "Moyen", "Élevé", "Critique"],
  statutsProjet: ["Idée", "En préparation", "Actif", "Suspendu", "Clôturé"],
  maturites: ["Concept", "Pré-faisabilité", "Faisabilité", "Financement recherché", "Financé", "En mise en oeuvre", "Achevé"],
  instruments: ["Subvention", "Prêt concessionnel", "Garantie", "PPP", "Budget municipal", "Co-financement", "Assistance technique"],
  statutsFinancement: ["Identifié", "En discussion", "Demandé", "Approuvé", "Décaissé", "Refusé", "Clôturé"],
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

const newSheets = [
  "00_Formulaire_Saisie",
  "01_Form_Temperature",
  "02_Form_Pluviometrie",
  "03_Form_Electricite",
  "04_Form_Dechets",
  "05_Form_Qualite_Air",
  "06_Form_Emissions_GES",
  "07_Form_Mobilite",
  "08_Form_Erosion_Cotiere",
  "09_Form_Projets_Climat",
  "10_Form_Finance_Climat",
  "Data_Formulaire_Central",
  "Data_Temperature",
  "Data_Pluviometrie",
  "Data_Electricite",
  "Data_Dechets",
  "Data_Qualite_Air",
  "Data_Emissions_GES",
  "Data_Mobilite",
  "Data_Erosion_Cotiere",
  "Data_Projets_Climat",
  "Data_Finance_Climat",
  "Tableau_de_bord",
  "Inventaire_GES",
  "Donnees_Climatiques",
  "Qualite_Air",
  "Energie",
  "Dechets",
  "Mobilite",
  "Risques_Climatiques",
  "Reporting_CDP_C40_ICLEI",
  "Contribution_CDN_2_0",
  "41_DASHBOARD_FORMS",
  "42_CONTROLE_QUALITE",
  "43_MODE_EMPLOI",
  "44_FLUX_DONNEES",
  "45_DONNEES_PRIORITAIRES",
  "46_GRAPHIQUES_AUTO",
];

for (const sheetName of newSheets.toReversed()) {
  deleteIfExists(sheetName);
}

function setWidths(sheet, widths, rowLimit = 150) {
  widths.forEach((w, index) => {
    const col = colName(index + 1);
    sheet.getRange(`${col}1:${col}${rowLimit}`).format.columnWidthPx = w;
  });
}

function titleSheet(sheet, title, subtitle, endCol = "L") {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${endCol}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${endCol}1`).format = { fill: palette.navy, font: { bold: true, color: palette.white, size: 15 } };
  sheet.getRange(`A1:${endCol}1`).format.rowHeightPx = 34;
  sheet.getRange(`A2:${endCol}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange(`A2:${endCol}2`).format = { fill: palette.lightTeal, font: { italic: true, color: palette.slate }, wrapText: true };
  sheet.getRange(`A2:${endCol}2`).format.rowHeightPx = 44;
}

function addSheet(name, title, subtitle, endCol = "L") {
  const sheet = workbook.worksheets.add(name);
  titleSheet(sheet, title, subtitle, endCol);
  return sheet;
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
  range.conditionalFormats.add("containsText", { text: "Validé", format: { fill: palette.lightGreen, font: { color: "#166534" } } });
  range.conditionalFormats.add("containsText", { text: "À vérifier", format: { fill: palette.lightAmber, font: { color: "#92400E" } } });
  range.conditionalFormats.add("containsText", { text: "Rejeté", format: { fill: palette.lightRed, font: { color: "#991B1B" } } });
}

function applyNumberFormat(sheet, headers, headerName, formatCode) {
  const idx = headers.indexOf(headerName) + 1;
  if (idx > 0) {
    const col = colName(idx);
    sheet.getRange(`${col}${DATA_START}:${col}${DATA_END}`).setNumberFormat(formatCode);
  }
}

function applyCommonValidations(sheet, headers) {
  const map = [
    ["Domaine de donnée", lists.domaines],
    ["Domaine", lists.domaines],
    ["Sous-domaine", lists.sousDomaines],
    ["Année", lists.annees],
    ["Mois", lists.mois],
    ["Commune / zone", lists.communes],
    ["Zone", lists.communes],
    ["Zone côtière", lists.communes],
    ["Station / zone", lists.communes],
    ["Institution source", lists.institutions],
    ["Service source", lists.services],
    ["Fréquence de collecte", lists.frequences],
    ["Niveau de fiabilité", lists.fiabilite],
    ["Statut validation", lists.statutsValidation],
    ["Statut de validation", lists.statutsValidation],
    ["Unité", lists.unites],
    ["Épisode extrême", lists.ouiNon],
    ["Dépassement seuil OMS", lists.ouiNon],
    ["Secteur", lists.secteursGes],
    ["Type de consommation", lists.typesConso],
    ["Type de déchets", lists.typesDechets],
    ["Mode transport", lists.modesTransport],
    ["Niveau de risque", lists.niveauxRisque],
    ["Statut projet", lists.statutsProjet],
    ["Niveau de maturité", lists.maturites],
    ["Instrument", lists.instruments],
    ["Statut financement", lists.statutsFinancement],
  ];
  for (const [header, values] of map) {
    const idx = headers.indexOf(header) + 1;
    if (idx > 0) addValidation(sheet, idx, DATA_START, DATA_END, values);
  }
}

function createTableSheet({ name, title, subtitle, headers, tableName, widths, formulas = {}, validations = true, numberFormats = {}, minRows = FORM_ROWS }) {
  const endCol = colName(headers.length);
  const endRow = DATA_START + minRows - 1;
  const sheet = addSheet(name, title, subtitle, endCol);
  const blankRows = Array.from({ length: minRows }, () => headers.map(() => null));
  sheet.getRange(`A${HEADER_ROW}:${endCol}${endRow}`).values = [headers, ...blankRows];
  for (const [colIndexText, formulaFactory] of Object.entries(formulas)) {
    const colIndex = Number(colIndexText);
    const col = colName(colIndex);
    const formulaRows = [];
    for (let row = DATA_START; row <= endRow; row += 1) {
      formulaRows.push([formulaFactory(row)]);
    }
    sheet.getRange(`${col}${DATA_START}:${col}${endRow}`).formulas = formulaRows;
  }
  sheet.getRange(`A${HEADER_ROW}:${endCol}${HEADER_ROW}`).format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${DATA_START}:${endCol}${endRow}`).format = { fill: palette.white, wrapText: true };
  const table = sheet.tables.add(`A${HEADER_ROW}:${endCol}${endRow}`, true, tableName);
  table.style = "TableStyleMedium2";
  table.showFilterButton = true;
  setWidths(sheet, widths ?? headers.map(() => 130), endRow + 5);
  sheet.freezePanes.freezeRows(HEADER_ROW);
  if (validations) applyCommonValidations(sheet, headers);
  for (const [header, fmt] of Object.entries(numberFormats)) {
    applyNumberFormat(sheet, headers, header, fmt);
  }
  const alertIdx = headers.indexOf("Alerte_QC") + 1;
  if (alertIdx > 0) addStatusFormatting(sheet, `${colName(alertIdx)}${DATA_START}:${colName(alertIdx)}${endRow}`);
  const statusIdx = headers.findIndex((h) => h === "Statut validation" || h === "Statut de validation") + 1;
  if (statusIdx > 0) addStatusFormatting(sheet, `${colName(statusIdx)}${DATA_START}:${colName(statusIdx)}${endRow}`);
  return { sheet, endRow };
}

function requiredAlertFormula({ row, primary, required, statusCol, abnormal = "FALSE" }) {
  const blanks = required.map((col) => `${col}${row}=""`).join(",");
  return `=IF(${primary}${row}="","",IF(OR(${blanks},${statusCol}${row}<>"${VALID_STATUS}"),"ALERTE: données/preuves/statut",IF(${abnormal},"ALERTE: valeur anormale","OK")))`;
}

function alimentFormula(primaryCol, statusCol) {
  return (row) => `=IF(${primaryCol}${row}="","",IF(${statusCol}${row}="${VALID_STATUS}","Oui","Non"))`;
}

function idFormula(prefix, primaryCol) {
  return (row) => `=IF(${primaryCol}${row}="","","${prefix}-"&TEXT(ROW()-${HEADER_ROW},"0000"))`;
}

function dateSaisieFormula(primaryCol) {
  return (row) => `=IF(${primaryCol}${row}="","",TODAY())`;
}

const forms = [
  {
    name: "00_Formulaire_Saisie",
    dataName: "Data_Formulaire_Central",
    title: "Formulaire central de saisie",
    subtitle: "Saisie transversale des données climatiques. Les lignes ne passent en base consolidée que si le statut est Validé humainement.",
    tableName: "tblFormulaireCentral",
    dataTableName: "tblDataFormulaireCentral",
    primaryCol: "B",
    statusCol: "P",
    docCol: "M",
    alertCol: "T",
    headers: [
      "ID_Saisie",
      "Domaine de donnée",
      "Sous-domaine",
      "Indicateur",
      "Année",
      "Mois",
      "Commune / zone",
      "Valeur",
      "Unité",
      "Institution source",
      "Service source",
      "Personne focale",
      "Document source",
      "Fréquence de collecte",
      "Niveau de fiabilité",
      "Statut de validation",
      "Observations",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 170, 170, 210, 90, 120, 180, 110, 95, 180, 150, 150, 210, 150, 145, 160, 260, 120, 150, 190, 120],
    formulas: {
      1: idFormula("GEN", "B"),
      18: dateSaisieFormula("B"),
      20: (row) => requiredAlertFormula({ row, primary: "B", required: ["D", "E", "H", "I", "J", "M", "O"], statusCol: "P" }),
      21: alimentFormula("B", "P"),
    },
    numberFormats: { "Valeur": "#,##0.00", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "01_Form_Temperature",
    dataName: "Data_Temperature",
    title: "Formulaire température",
    subtitle: "Saisie mensuelle des données de température par station et zone.",
    tableName: "tblFormTemperature",
    dataTableName: "tblDataTemperature",
    primaryCol: "B",
    statusCol: "O",
    docCol: "M",
    alertCol: "S",
    headers: [
      "ID_Saisie",
      "Date",
      "Année",
      "Mois",
      "Station météo",
      "Commune / zone",
      "Température moyenne",
      "Température minimale",
      "Température maximale",
      "Nombre de jours chauds",
      "Source",
      "Institution source",
      "Document source",
      "Niveau de fiabilité",
      "Statut validation",
      "Observations",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 110, 90, 115, 150, 180, 135, 135, 135, 150, 140, 180, 210, 145, 160, 250, 120, 150, 190, 120],
    formulas: {
      1: idFormula("TEMP", "B"),
      17: dateSaisieFormula("B"),
      19: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "D", "F", "G", "K", "M", "N"],
          statusCol: "O",
          abnormal: `OR(AND(ISNUMBER(G${row}),OR(G${row}<0,G${row}>50)),AND(ISNUMBER(H${row}),H${row}<-5),AND(ISNUMBER(I${row}),I${row}>55),AND(ISNUMBER(J${row}),J${row}>31))`,
        }),
      20: alimentFormula("B", "O"),
    },
    numberFormats: { "Date": "yyyy-mm-dd", "Température moyenne": "0.0", "Température minimale": "0.0", "Température maximale": "0.0", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "02_Form_Pluviometrie",
    dataName: "Data_Pluviometrie",
    title: "Formulaire pluviométrie",
    subtitle: "Saisie mensuelle de la pluie, des jours de pluie et des épisodes extrêmes.",
    tableName: "tblFormPluviometrie",
    dataTableName: "tblDataPluviometrie",
    primaryCol: "B",
    statusCol: "L",
    docCol: "K",
    alertCol: "P",
    headers: [
      "ID_Saisie",
      "Date",
      "Année",
      "Mois",
      "Station météo",
      "Commune / zone",
      "Pluviométrie mensuelle",
      "Nombre de jours de pluie",
      "Épisode extrême",
      "Source",
      "Document source",
      "Statut validation",
      "Observations",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 110, 90, 115, 150, 180, 155, 155, 130, 145, 210, 160, 250, 120, 150, 190, 120],
    formulas: {
      1: idFormula("PLUIE", "B"),
      14: dateSaisieFormula("B"),
      16: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "D", "F", "G", "J", "K"],
          statusCol: "L",
          abnormal: `OR(AND(ISNUMBER(G${row}),OR(G${row}<0,G${row}>1000)),AND(ISNUMBER(H${row}),H${row}>31))`,
        }),
      17: alimentFormula("B", "L"),
    },
    numberFormats: { "Date": "yyyy-mm-dd", "Pluviométrie mensuelle": "#,##0.0", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "03_Form_Electricite",
    dataName: "Data_Electricite",
    title: "Formulaire électricité",
    subtitle: "Saisie de consommation électrique et estimation automatique des émissions associées.",
    tableName: "tblFormElectricite",
    dataTableName: "tblDataElectricite",
    primaryCol: "B",
    statusCol: "K",
    docCol: "J",
    alertCol: "O",
    headers: [
      "ID_Saisie",
      "Année",
      "Mois",
      "Zone",
      "Type de consommation",
      "Consommation kWh",
      "Facteur d'émission kgCO2e/kWh",
      "Émissions estimées tCO2e",
      "Institution source",
      "Document source",
      "Statut validation",
      "Observations",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 90, 115, 180, 170, 140, 185, 155, 180, 210, 160, 250, 120, 150, 190, 120],
    formulas: {
      1: idFormula("ELEC", "B"),
      8: (row) => `=IF(OR(F${row}="",G${row}=""),"",F${row}*G${row}/1000)`,
      13: dateSaisieFormula("B"),
      15: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "D", "E", "F", "G", "I", "J"],
          statusCol: "K",
          abnormal: `OR(AND(ISNUMBER(F${row}),F${row}<0),AND(ISNUMBER(G${row}),G${row}<0))`,
        }),
      16: alimentFormula("B", "K"),
    },
    numberFormats: { "Consommation kWh": "#,##0", "Facteur d'émission kgCO2e/kWh": "0.0000", "Émissions estimées tCO2e": "#,##0.00", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "04_Form_Dechets",
    dataName: "Data_Dechets",
    title: "Formulaire déchets",
    subtitle: "Saisie des flux de déchets, valorisation et estimation simplifiée des émissions liées à la mise en décharge.",
    tableName: "tblFormDechets",
    dataTableName: "tblDataDechets",
    primaryCol: "B",
    statusCol: "O",
    docCol: "L",
    alertCol: "R",
    headers: [
      "ID_Saisie",
      "Année",
      "Mois",
      "Zone",
      "Type de déchets",
      "Quantité collectée",
      "Quantité recyclée",
      "Quantité compostée",
      "Quantité mise en décharge",
      "Unité",
      "Institution source",
      "Document source",
      "Émissions estimées",
      "Observations",
      "Statut validation",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 90, 115, 180, 160, 140, 140, 145, 165, 95, 180, 210, 145, 250, 160, 120, 150, 190, 120],
    formulas: {
      1: idFormula("DECH", "B"),
      13: (row) => `=IF(I${row}="","",I${row}*0.586)`,
      16: dateSaisieFormula("B"),
      18: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "D", "E", "F", "J", "K", "L"],
          statusCol: "O",
          abnormal: `OR(AND(ISNUMBER(F${row}),F${row}<0),AND(ISNUMBER(G${row}),G${row}<0),AND(ISNUMBER(H${row}),H${row}<0),AND(ISNUMBER(I${row}),I${row}<0),AND(ISNUMBER(F${row}),ISNUMBER(G${row}),ISNUMBER(H${row}),ISNUMBER(I${row}),G${row}+H${row}+I${row}>F${row}))`,
        }),
      19: alimentFormula("B", "O"),
    },
    numberFormats: { "Quantité collectée": "#,##0.00", "Quantité recyclée": "#,##0.00", "Quantité compostée": "#,##0.00", "Quantité mise en décharge": "#,##0.00", "Émissions estimées": "#,##0.00", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "05_Form_Qualite_Air",
    dataName: "Data_Qualite_Air",
    title: "Formulaire qualité de l'air",
    subtitle: "Saisie des polluants atmosphériques et alerte automatique de dépassement des seuils OMS indicatifs.",
    tableName: "tblFormQualiteAir",
    dataTableName: "tblDataQualiteAir",
    primaryCol: "B",
    statusCol: "P",
    docCol: "N",
    alertCol: "S",
    headers: [
      "ID_Saisie",
      "Année",
      "Mois",
      "Station / zone",
      "PM2.5",
      "PM10",
      "NO2",
      "SO2",
      "CO",
      "O3",
      "Indice qualité air",
      "Dépassement seuil OMS",
      "Source",
      "Document source",
      "Observations",
      "Statut validation",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 90, 115, 180, 95, 95, 95, 95, 95, 95, 135, 150, 145, 210, 250, 160, 120, 150, 190, 120],
    formulas: {
      1: idFormula("AIR", "B"),
      11: (row) => `=IF(COUNTA(E${row}:J${row})=0,"",MAX(E${row}:J${row}))`,
      12: (row) => `=IF(COUNTA(E${row}:J${row})=0,"",IF(OR(E${row}>15,F${row}>45,G${row}>25,H${row}>40,I${row}>4000,J${row}>100),"Oui","Non"))`,
      17: dateSaisieFormula("B"),
      19: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "D", "E", "M", "N"],
          statusCol: "P",
          abnormal: `OR(E${row}<0,F${row}<0,G${row}<0,H${row}<0,I${row}<0,J${row}<0)`,
        }),
      20: alimentFormula("B", "P"),
    },
    numberFormats: { "PM2.5": "#,##0.0", "PM10": "#,##0.0", "NO2": "#,##0.0", "SO2": "#,##0.0", "CO": "#,##0.0", "O3": "#,##0.0", "Indice qualité air": "#,##0.0", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "06_Form_Emissions_GES",
    dataName: "Data_Emissions_GES",
    title: "Formulaire émissions GES",
    subtitle: "Saisie des données d'activité GES avec calcul automatique des émissions tCO2e.",
    tableName: "tblFormEmissionsGES",
    dataTableName: "tblDataEmissionsGES",
    primaryCol: "B",
    statusCol: "L",
    docCol: "K",
    alertCol: "P",
    headers: [
      "ID_Saisie",
      "Année",
      "Secteur",
      "Sous-secteur",
      "Donnée d'activité",
      "Unité",
      "Facteur d'émission",
      "Émissions tCO2e",
      "Méthodologie",
      "Source",
      "Document source",
      "Statut validation",
      "Observations",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 90, 170, 170, 145, 95, 135, 140, 160, 145, 210, 160, 250, 120, 150, 190, 120],
    formulas: {
      1: idFormula("GES", "B"),
      8: (row) => `=IF(OR(E${row}="",G${row}=""),"",E${row}*G${row})`,
      14: dateSaisieFormula("B"),
      16: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "D", "E", "F", "G", "I", "J", "K"],
          statusCol: "L",
          abnormal: `OR(AND(ISNUMBER(E${row}),E${row}<0),AND(ISNUMBER(G${row}),G${row}<0),AND(ISNUMBER(H${row}),H${row}<0))`,
        }),
      17: alimentFormula("B", "L"),
    },
    numberFormats: { "Donnée d'activité": "#,##0.00", "Facteur d'émission": "#,##0.0000", "Émissions tCO2e": "#,##0.00", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "07_Form_Mobilite",
    dataName: "Data_Mobilite",
    title: "Formulaire mobilité",
    subtitle: "Saisie des données de transport, activité, consommation et émissions.",
    tableName: "tblFormMobilite",
    dataTableName: "tblDataMobilite",
    primaryCol: "B",
    statusCol: "Q",
    docCol: "P",
    alertCol: "U",
    headers: [
      "ID_Saisie",
      "Année",
      "Mois",
      "Zone",
      "Mode transport",
      "Nombre trajets",
      "Distance moyenne km",
      "Passagers",
      "Carburant / énergie",
      "Consommation",
      "Unité",
      "Facteur d'émission",
      "Émissions tCO2e",
      "Source",
      "Institution source",
      "Document source",
      "Statut validation",
      "Observations",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 90, 115, 180, 150, 125, 150, 115, 150, 125, 95, 135, 140, 145, 180, 210, 160, 250, 120, 150, 190, 120],
    formulas: {
      1: idFormula("MOB", "B"),
      13: (row) => `=IF(OR(J${row}="",L${row}=""),"",J${row}*L${row})`,
      19: dateSaisieFormula("B"),
      21: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "D", "E", "J", "K", "L", "N", "O", "P"],
          statusCol: "Q",
          abnormal: `OR(F${row}<0,G${row}<0,H${row}<0,J${row}<0,L${row}<0,M${row}<0)`,
        }),
      22: alimentFormula("B", "Q"),
    },
    numberFormats: { "Nombre trajets": "#,##0", "Distance moyenne km": "#,##0.0", "Passagers": "#,##0", "Consommation": "#,##0.00", "Facteur d'émission": "#,##0.0000", "Émissions tCO2e": "#,##0.00", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "08_Form_Erosion_Cotiere",
    dataName: "Data_Erosion_Cotiere",
    title: "Formulaire érosion côtière",
    subtitle: "Saisie du recul du trait de côte, infrastructures et populations exposées.",
    tableName: "tblFormErosionCotiere",
    dataTableName: "tblDataErosionCotiere",
    primaryCol: "B",
    statusCol: "M",
    docCol: "K",
    alertCol: "P",
    headers: [
      "ID_Saisie",
      "Année",
      "Mois",
      "Zone côtière",
      "Recul du trait de côte",
      "Unité",
      "Infrastructure exposée",
      "Population exposée",
      "Niveau de risque",
      "Source",
      "Document source",
      "Observations",
      "Statut validation",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 90, 115, 180, 160, 95, 190, 145, 135, 145, 210, 250, 160, 120, 150, 190, 120],
    formulas: {
      1: idFormula("COTE", "B"),
      14: dateSaisieFormula("B"),
      16: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "D", "E", "F", "I", "J", "K"],
          statusCol: "M",
          abnormal: `OR(E${row}<0,E${row}>50,H${row}<0)`,
        }),
      17: alimentFormula("B", "M"),
    },
    numberFormats: { "Recul du trait de côte": "#,##0.00", "Population exposée": "#,##0", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "09_Form_Projets_Climat",
    dataName: "Data_Projets_Climat",
    title: "Formulaire projets climat",
    subtitle: "Saisie des projets climat, maturité, budget, financement et avancement.",
    tableName: "tblFormProjetsClimat",
    dataTableName: "tblDataProjetsClimat",
    primaryCol: "A",
    statusCol: "P",
    docCol: "O",
    alertCol: "T",
    headers: [
      "Projet_ID",
      "Nom projet",
      "Domaine",
      "Type action",
      "Statut projet",
      "Niveau de maturité",
      "Date début",
      "Date fin",
      "Budget total FCFA",
      "Financement acquis FCFA",
      "Gap financement FCFA",
      "Avancement %",
      "Responsable projet",
      "Partenaires",
      "Document source",
      "Statut validation",
      "Observations",
      "Date de saisie",
      "Responsable saisie",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 210, 150, 160, 135, 170, 110, 110, 150, 165, 160, 115, 160, 210, 210, 160, 250, 120, 155, 190, 120],
    formulas: {
      1: idFormula("PRJ", "B"),
      11: (row) => `=IF(I${row}="","",MAX(0,I${row}-J${row}))`,
      18: dateSaisieFormula("B"),
      20: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "E", "F", "I", "M", "O"],
          statusCol: "P",
          abnormal: `OR(I${row}<0,J${row}<0,L${row}<0,L${row}>1)`,
        }),
      21: alimentFormula("B", "P"),
    },
    numberFormats: { "Date début": "yyyy-mm-dd", "Date fin": "yyyy-mm-dd", "Budget total FCFA": "#,##0", "Financement acquis FCFA": "#,##0", "Gap financement FCFA": "#,##0", "Avancement %": "0%", "Date de saisie": "yyyy-mm-dd" },
  },
  {
    name: "10_Form_Finance_Climat",
    dataName: "Data_Finance_Climat",
    title: "Formulaire finance climat",
    subtitle: "Saisie des demandes, approbations et décaissements liés aux financements climat.",
    tableName: "tblFormFinanceClimat",
    dataTableName: "tblDataFinanceClimat",
    primaryCol: "B",
    statusCol: "O",
    docCol: "N",
    alertCol: "S",
    headers: [
      "ID_Saisie",
      "Année",
      "Mois",
      "Projet_ID",
      "Bailleur",
      "Instrument",
      "Montant demandé FCFA",
      "Montant approuvé FCFA",
      "Montant décaissé FCFA",
      "Cofinancement FCFA",
      "Devise",
      "Statut financement",
      "Date décision",
      "Document source",
      "Statut validation",
      "Observations",
      "Date de saisie",
      "Responsable",
      "Alerte_QC",
      "Alimente_Rapports",
    ],
    widths: [110, 90, 115, 110, 180, 150, 160, 160, 160, 150, 95, 150, 120, 210, 160, 250, 120, 150, 190, 120],
    formulas: {
      1: idFormula("FIN", "B"),
      17: dateSaisieFormula("B"),
      19: (row) =>
        requiredAlertFormula({
          row,
          primary: "B",
          required: ["C", "D", "E", "F", "G", "K", "L", "N"],
          statusCol: "O",
          abnormal: `OR(G${row}<0,H${row}<0,I${row}<0,J${row}<0,AND(G${row}<>"",H${row}>G${row}),AND(H${row}<>"",I${row}>H${row}))`,
        }),
      20: alimentFormula("B", "O"),
    },
    numberFormats: { "Montant demandé FCFA": "#,##0", "Montant approuvé FCFA": "#,##0", "Montant décaissé FCFA": "#,##0", "Cofinancement FCFA": "#,##0", "Date décision": "yyyy-mm-dd", "Date de saisie": "yyyy-mm-dd" },
  },
];

for (const form of forms) {
  createTableSheet({
    name: form.name,
    title: form.title,
    subtitle: form.subtitle,
    headers: form.headers,
    tableName: form.tableName,
    widths: form.widths,
    formulas: form.formulas,
    numberFormats: form.numberFormats,
  });
}

function createDataSheet(form) {
  const headers = form.headers;
  const endCol = colName(headers.length);
  const sheet = addSheet(
    form.dataName,
    `Base validée - ${form.title.replace("Formulaire ", "")}`,
    `Cette base est alimentée automatiquement depuis ${form.name}. Seules les lignes avec le statut "${VALID_STATUS}" sont reprises.`,
    endCol
  );
  const blankRows = Array.from({ length: FORM_ROWS }, () => headers.map(() => null));
  sheet.getRange(`A${HEADER_ROW}:${endCol}${DATA_END}`).values = [headers, ...blankRows];
  const statusIndex = headers.findIndex((h) => h === "Statut validation" || h === "Statut de validation") + 1;
  const statusCol = colName(statusIndex);
  for (let colIndex = 1; colIndex <= headers.length; colIndex += 1) {
    const col = colName(colIndex);
    const formulas = [];
    for (let row = DATA_START; row <= DATA_END; row += 1) {
      formulas.push([`=IF(${q(form.name)}!$${statusCol}${row}="${VALID_STATUS}",${q(form.name)}!${col}${row},"")`]);
    }
    sheet.getRange(`${col}${DATA_START}:${col}${DATA_END}`).formulas = formulas;
  }
  sheet.getRange(`A${HEADER_ROW}:${endCol}${HEADER_ROW}`).format = { fill: palette.green, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${DATA_START}:${endCol}${DATA_END}`).format = { fill: palette.white, wrapText: true };
  const table = sheet.tables.add(`A${HEADER_ROW}:${endCol}${DATA_END}`, true, form.dataTableName);
  table.style = "TableStyleMedium4";
  table.showFilterButton = true;
  setWidths(sheet, form.widths, DATA_END + 5);
  sheet.freezePanes.freezeRows(HEADER_ROW);
  for (const [header, fmt] of Object.entries(form.numberFormats ?? {})) {
    applyNumberFormat(sheet, headers, header, fmt);
  }
}

for (const form of forms) {
  createDataSheet(form);
}

function validatedCountFormula() {
  return `=${forms
    .map((form) => {
      const statusIndex = form.headers.findIndex((h) => h === "Statut validation" || h === "Statut de validation") + 1;
      const statusCol = colName(statusIndex);
      return `COUNTIF(${form.dataName}!$${statusCol}$${DATA_START}:$${statusCol}$${DATA_END},"${VALID_STATUS}")`;
    })
    .join("+")}`;
}

const VALIDATED_COUNT_FORMULA = validatedCountFormula();

function metricSheet({ name, title, subtitle, rows, tableName }) {
  const sheet = addSheet(name, title, subtitle, "E");
  setWidths(sheet, [260, 160, 220, 220, 280], 70);
  sheet.getRange("A4:E4").values = [["Indicateur", "Valeur", "Unité", "Source", "Usage reporting"]];
  sheet.getRange(`A5:E${4 + rows.length}`).values = rows.map((r) => [r.label, null, r.unit, r.source, r.usage]);
  sheet.getRange(`B5:B${4 + rows.length}`).formulas = rows.map((r) => [r.formula]);
  sheet.getRange("A4:E4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:E${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  sheet.tables.add(`A4:E${4 + rows.length}`, true, tableName).style = "TableStyleMedium2";
  sheet.getRange(`B5:B${4 + rows.length}`).setNumberFormat("#,##0.00");
  return sheet;
}

metricSheet({
  name: "Tableau_de_bord",
  title: "Tableau de bord consolidé",
  subtitle: "Vue automatiquement alimentée par les bases validées issues des formulaires.",
  tableName: "tblVueTableauBord",
  rows: [
    { label: "Données validées", formula: VALIDATED_COUNT_FORMULA, unit: "lignes", source: "Bases Data_*", usage: "Pilotage et reporting" },
    { label: "Émissions totales", formula: "=SUM(Data_Electricite!$H$5:$H$104)+SUM(Data_Dechets!$M$5:$M$104)+SUM(Data_Emissions_GES!$H$5:$H$104)+SUM(Data_Mobilite!$M$5:$M$104)", unit: "tCO2e", source: "Data GES, énergie, déchets, mobilité", usage: "CDP, C40, CDN" },
    { label: "Budget total projets", formula: "=SUM(Data_Projets_Climat!$I$5:$I$104)", unit: "FCFA", source: "Data_Projets_Climat", usage: "Bailleurs climat" },
    { label: "Gap de financement", formula: "=SUM(Data_Projets_Climat!$K$5:$K$104)", unit: "FCFA", source: "Data_Projets_Climat", usage: "GCF, Banque mondiale, BAD" },
    { label: "Projets actifs", formula: '=COUNTIF(Data_Projets_Climat!$E$5:$E$104,"Actif")', unit: "nombre", source: "Data_Projets_Climat", usage: "C40, CDN" },
  ],
});

metricSheet({
  name: "Inventaire_GES",
  title: "Inventaire GES alimenté par formulaires",
  subtitle: "Synthèse des émissions validées, structurée pour inventaire territorial.",
  tableName: "tblVueInventaireGES",
  rows: [
    { label: "Énergie stationnaire", formula: '=SUMIFS(Data_Emissions_GES!$H$5:$H$104,Data_Emissions_GES!$C$5:$C$104,"Énergie stationnaire")+SUM(Data_Electricite!$H$5:$H$104)', unit: "tCO2e", source: "Data_Emissions_GES; Data_Electricite", usage: "GPC/CDP/GCoM" },
    { label: "Transport", formula: '=SUMIFS(Data_Emissions_GES!$H$5:$H$104,Data_Emissions_GES!$C$5:$C$104,"Transport")+SUM(Data_Mobilite!$M$5:$M$104)', unit: "tCO2e", source: "Data_Emissions_GES; Data_Mobilite", usage: "GPC/CDP/GCoM" },
    { label: "Déchets", formula: '=SUMIFS(Data_Emissions_GES!$H$5:$H$104,Data_Emissions_GES!$C$5:$C$104,"Déchets")+SUM(Data_Dechets!$M$5:$M$104)', unit: "tCO2e", source: "Data_Emissions_GES; Data_Dechets", usage: "GPC/CDP/GCoM" },
    { label: "Autres secteurs", formula: '=SUMIFS(Data_Emissions_GES!$H$5:$H$104,Data_Emissions_GES!$C$5:$C$104,"<>Énergie stationnaire",Data_Emissions_GES!$C$5:$C$104,"<>Transport",Data_Emissions_GES!$C$5:$C$104,"<>Déchets")', unit: "tCO2e", source: "Data_Emissions_GES", usage: "GPC BASIC+" },
  ],
});

metricSheet({
  name: "Donnees_Climatiques",
  title: "Données climatiques consolidées",
  subtitle: "Température, pluviométrie et érosion côtière validées.",
  tableName: "tblVueDonneesClimatiques",
  rows: [
    { label: "Température moyenne", formula: '=IFERROR(AVERAGE(Data_Temperature!$G$5:$G$104),"")', unit: "°C", source: "Data_Temperature", usage: "CDP, C40, adaptation" },
    { label: "Pluviométrie totale", formula: "=SUM(Data_Pluviometrie!$G$5:$G$104)", unit: "mm", source: "Data_Pluviometrie", usage: "Adaptation/CDN" },
    { label: "Épisodes extrêmes", formula: '=COUNTIF(Data_Pluviometrie!$I$5:$I$104,"Oui")', unit: "nombre", source: "Data_Pluviometrie", usage: "Résilience" },
    { label: "Recul moyen trait de côte", formula: '=IFERROR(AVERAGE(Data_Erosion_Cotiere!$E$5:$E$104),"")', unit: "m", source: "Data_Erosion_Cotiere", usage: "Risques climatiques/CDN" },
  ],
});

metricSheet({
  name: "Qualite_Air",
  title: "Qualité de l'air consolidée",
  subtitle: "Moyennes validées et dépassements de seuil.",
  tableName: "tblVueQualiteAir",
  rows: [
    { label: "PM2.5 moyen", formula: '=IFERROR(AVERAGE(Data_Qualite_Air!$E$5:$E$104),"")', unit: "µg/m3", source: "Data_Qualite_Air", usage: "CDP, C40 air quality" },
    { label: "PM10 moyen", formula: '=IFERROR(AVERAGE(Data_Qualite_Air!$F$5:$F$104),"")', unit: "µg/m3", source: "Data_Qualite_Air", usage: "Santé urbaine" },
    { label: "Indice moyen qualité air", formula: '=IFERROR(AVERAGE(Data_Qualite_Air!$K$5:$K$104),"")', unit: "indice", source: "Data_Qualite_Air", usage: "Dashboard" },
    { label: "Dépassements seuil OMS", formula: '=COUNTIF(Data_Qualite_Air!$L$5:$L$104,"Oui")', unit: "nombre", source: "Data_Qualite_Air", usage: "Alerte santé" },
  ],
});

metricSheet({
  name: "Energie",
  title: "Énergie consolidée",
  subtitle: "Consommation électrique et émissions associées.",
  tableName: "tblVueEnergie",
  rows: [
    { label: "Consommation électrique", formula: "=SUM(Data_Electricite!$F$5:$F$104)", unit: "kWh", source: "Data_Electricite", usage: "CDP, C40, CDN" },
    { label: "Émissions électricité", formula: "=SUM(Data_Electricite!$H$5:$H$104)", unit: "tCO2e", source: "Data_Electricite", usage: "Inventaire GES" },
    { label: "Nombre de sites/mesures", formula: `=COUNTIF(Data_Electricite!$K$${DATA_START}:$K$${DATA_END},"${VALID_STATUS}")`, unit: "lignes", source: "Data_Electricite", usage: "MRV" },
  ],
});

metricSheet({
  name: "Dechets",
  title: "Déchets consolidés",
  subtitle: "Flux collectés, valorisés, enfouis et émissions estimées.",
  tableName: "tblVueDechets",
  rows: [
    { label: "Déchets collectés", formula: "=SUM(Data_Dechets!$F$5:$F$104)", unit: "tonnes", source: "Data_Dechets", usage: "CDP, C40, CDN" },
    { label: "Déchets recyclés", formula: "=SUM(Data_Dechets!$G$5:$G$104)", unit: "tonnes", source: "Data_Dechets", usage: "Circularité" },
    { label: "Déchets compostés", formula: "=SUM(Data_Dechets!$H$5:$H$104)", unit: "tonnes", source: "Data_Dechets", usage: "Circularité" },
    { label: "Émissions déchets", formula: "=SUM(Data_Dechets!$M$5:$M$104)", unit: "tCO2e", source: "Data_Dechets", usage: "Inventaire GES" },
  ],
});

metricSheet({
  name: "Mobilite",
  title: "Mobilité consolidée",
  subtitle: "Activité transport, passagers, consommation et émissions.",
  tableName: "tblVueMobilite",
  rows: [
    { label: "Trajets déclarés", formula: "=SUM(Data_Mobilite!$F$5:$F$104)", unit: "nombre", source: "Data_Mobilite", usage: "Transport urbain" },
    { label: "Passagers", formula: "=SUM(Data_Mobilite!$H$5:$H$104)", unit: "nombre", source: "Data_Mobilite", usage: "C40, CDN" },
    { label: "Émissions mobilité", formula: "=SUM(Data_Mobilite!$M$5:$M$104)", unit: "tCO2e", source: "Data_Mobilite", usage: "Inventaire GES" },
  ],
});

metricSheet({
  name: "Risques_Climatiques",
  title: "Risques climatiques consolidés",
  subtitle: "Risques côtiers et exposition issus des données validées.",
  tableName: "tblVueRisquesClimatiques",
  rows: [
    { label: "Population exposée érosion", formula: "=SUM(Data_Erosion_Cotiere!$H$5:$H$104)", unit: "personnes", source: "Data_Erosion_Cotiere", usage: "Adaptation, CDN, bailleurs" },
    { label: "Zones risque critique", formula: '=COUNTIF(Data_Erosion_Cotiere!$I$5:$I$104,"Critique")', unit: "zones", source: "Data_Erosion_Cotiere", usage: "Plan adaptation" },
    { label: "Recul maximal observé", formula: '=IFERROR(MAX(Data_Erosion_Cotiere!$E$5:$E$104),"")', unit: "m", source: "Data_Erosion_Cotiere", usage: "Priorisation investissements" },
  ],
});

metricSheet({
  name: "Reporting_CDP_C40_ICLEI",
  title: "Pré-remplissage reporting CDP / C40 / ICLEI",
  subtitle: "Indicateurs prêts à reporter, uniquement sur données validées humainement.",
  tableName: "tblReportingForms",
  rows: [
    { label: "Total émissions GES", formula: "=Tableau_de_bord!B6", unit: "tCO2e", source: "Inventaire_GES", usage: "CDP/ICLEI/GCoM" },
    { label: "Consommation électrique", formula: "=Energie!B5", unit: "kWh", source: "Energie", usage: "CDP énergie" },
    { label: "Déchets collectés", formula: "=Dechets!B5", unit: "tonnes", source: "Dechets", usage: "CDP déchets" },
    { label: "Indice qualité air moyen", formula: "=Qualite_Air!B7", unit: "indice", source: "Qualite_Air", usage: "C40 qualité air" },
    { label: "Projets climat actifs", formula: "=Tableau_de_bord!B9", unit: "nombre", source: "Data_Projets_Climat", usage: "C40 actions" },
  ],
});

metricSheet({
  name: "Contribution_CDN_2_0",
  title: "Contribution locale à la CDN 2.0",
  subtitle: "Contribution de Dakar par secteur, basée sur les bases validées.",
  tableName: "tblContributionCDN",
  rows: [
    { label: "Contribution mitigation énergie", formula: "=Energie!B6", unit: "tCO2e", source: "Data_Electricite", usage: "CDN énergie" },
    { label: "Contribution mitigation transport", formula: "=Mobilite!B7", unit: "tCO2e", source: "Data_Mobilite", usage: "CDN transport" },
    { label: "Contribution mitigation déchets", formula: "=Dechets!B8", unit: "tCO2e", source: "Data_Dechets", usage: "CDN déchets" },
    { label: "Contribution adaptation côtière", formula: "=Risques_Climatiques!B5", unit: "personnes exposées", source: "Data_Erosion_Cotiere", usage: "CDN adaptation" },
    { label: "Financement climat mobilisé", formula: "=SUM(Data_Finance_Climat!$H$5:$H$104)", unit: "FCFA", source: "Data_Finance_Climat", usage: "CDN finance" },
  ],
});

function buildDashboard() {
  const sheet = addSheet(
    "41_DASHBOARD_FORMS",
    "Dashboard dynamique - module formulaires",
    "Tableau de bord exécutif alimenté par les bases validées. Les lignes non validées humainement sont exclues automatiquement.",
    "N"
  );
  setWidths(sheet, [210, 145, 140, 220, 210, 145, 140, 220, 210, 145, 140, 220, 170, 170], 80);
  sheet.getRange("A4:N4").values = [["KPI", "Valeur", "Unité", "Source", "KPI", "Valeur", "Unité", "Source", "KPI", "Valeur", "Unité", "Source", "Statut", "Lecture"]];
  sheet.getRange("A4:N4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  const left = [
    ["Dernière année disponible", '=IF(MAX(MAX(Data_Temperature!$C$5:$C$104),MAX(Data_Pluviometrie!$C$5:$C$104),MAX(Data_Electricite!$B$5:$B$104),MAX(Data_Emissions_GES!$B$5:$B$104),MAX(Data_Finance_Climat!$B$5:$B$104))=0,"Aucune donnée",MAX(MAX(Data_Temperature!$C$5:$C$104),MAX(Data_Pluviometrie!$C$5:$C$104),MAX(Data_Electricite!$B$5:$B$104),MAX(Data_Emissions_GES!$B$5:$B$104),MAX(Data_Finance_Climat!$B$5:$B$104)))', "année", "Bases validées"],
    ["Données collectées", VALIDATED_COUNT_FORMULA, "lignes validées", "Data_*"],
    ["Données manquantes / alertes", "=SUM('42_CONTROLE_QUALITE'!$D$5:$D$15)", "alertes", "Contrôle qualité"],
    ["Émissions totales", "=Tableau_de_bord!B6", "tCO2e", "Inventaire_GES"],
  ];
  const middle = [
    ["Consommation électrique totale", "=Energie!B5", "kWh", "Data_Electricite"],
    ["Déchets collectés", "=Dechets!B5", "tonnes", "Data_Dechets"],
    ["Déchets valorisés", "=Dechets!B6+Dechets!B7", "tonnes", "Data_Dechets"],
    ["Température moyenne annuelle", "=Donnees_Climatiques!B5", "°C", "Data_Temperature"],
  ];
  const right = [
    ["Pluviométrie annuelle", "=Donnees_Climatiques!B6", "mm", "Data_Pluviometrie"],
    ["Indice moyen qualité air", "=Qualite_Air!B7", "indice", "Data_Qualite_Air"],
    ["Projets climat actifs", "=Tableau_de_bord!B9", "nombre", "Data_Projets_Climat"],
    ["Budget total", "=Tableau_de_bord!B7", "FCFA", "Data_Projets_Climat"],
  ];
  for (let i = 0; i < 4; i += 1) {
    const row = 5 + i;
    sheet.getRange(`A${row}:D${row}`).values = [[left[i][0], null, left[i][2], left[i][3]]];
    sheet.getRange(`B${row}`).formulas = [[left[i][1]]];
    sheet.getRange(`E${row}:H${row}`).values = [[middle[i][0], null, middle[i][2], middle[i][3]]];
    sheet.getRange(`F${row}`).formulas = [[middle[i][1]]];
    sheet.getRange(`I${row}:L${row}`).values = [[right[i][0], null, right[i][2], right[i][3]]];
    sheet.getRange(`J${row}`).formulas = [[right[i][1]]];
  }
  sheet.getRange("M5:N8").formulas = [
    ['=IF(B7>0,"À corriger","OK")', '="Les alertes QC bloquent les rapports finaux."'],
    ['=IF(B6=0,"Collecte à lancer","OK")', '="Les bases validées alimentent tous les indicateurs."'],
    ['=IF(J8=0,"Portefeuille à renseigner","OK")', '="Les projets actifs sont requis pour C40/CDN."'],
    ['=IF(J7=0,"Collecte air à compléter","OK")', '="La qualité de l’air est attendue pour C40/CDP."'],
  ];
  sheet.getRange("A5:N8").format = { fill: palette.white, wrapText: true };
  sheet.getRange("B5:B8").setNumberFormat("#,##0.00");
  sheet.getRange("F5:F8").setNumberFormat("#,##0.00");
  sheet.getRange("J5:J8").setNumberFormat("#,##0.00");
  addStatusFormatting(sheet, "M5:M8");

  sheet.getRange("A11:D11").values = [["Flux obligatoire", "Étape", "Contrôle", "Sortie"]];
  sheet.getRange("A12:D17").values = [
    ["Document source", "Preuve déposée ou référencée", "Document source obligatoire", "Traçabilité MRV"],
    ["Formulaire", "Saisie par domaine", "Listes déroulantes et champs requis", "Ligne brute"],
    ["Base de données", "Copie automatique si validée", `Statut = ${VALID_STATUS}`, "Data_*"],
    ["Contrôle qualité", "Alerte manque/anomalie/retard", "Correction par responsable", "Ligne prête à valider"],
    ["Dashboard", "Calcul automatique des KPI", "Uniquement données validées", "Pilotage exécutif"],
    ["Rapports", "CDP/C40/ICLEI/CDN/bailleurs", "Preuves et audit trail", "Soumission externe"],
  ];
  sheet.getRange("A11:D11").format = { fill: palette.navy, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange("A12:D17").format = { fill: palette.grey, wrapText: true };
  sheet.tables.add("A11:D17", true, "tblFluxDashboardForms").style = "TableStyleMedium2";
  return sheet;
}

buildDashboard();

function buildQualityControl() {
  const sheet = addSheet(
    "42_CONTROLE_QUALITE",
    "Contrôle qualité et validation",
    "Suivi automatique des alertes bloquantes: donnée manquante, valeur anormale, source absente, document absent ou statut non validé.",
    "J"
  );
  setWidths(sheet, [215, 135, 135, 135, 145, 145, 145, 145, 135, 260], 80);
  const headers = ["Formulaire", "Lignes saisies", "Lignes validées", "Alertes QC", "Statuts non validés", "Documents absents", "Sources absentes", "Dates saisie absentes", "Statut", "Action attendue"];
  sheet.getRange("A4:J4").values = [headers];
  sheet.getRange("A4:J4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  const qcForms = forms.map((f) => {
    const headers = f.headers;
    const sourceIdx = headers.findIndex((h) => h === "Source" || h === "Institution source") + 1;
    const dateIdx = headers.findIndex((h) => h === "Date de saisie") + 1;
    return {
      label: f.title.replace("Formulaire ", ""),
      sheet: f.name,
      primaryCol: f.primaryCol,
      statusCol: f.statusCol,
      docCol: f.docCol,
      alertCol: f.alertCol,
      sourceCol: sourceIdx > 0 ? colName(sourceIdx) : f.docCol,
      dateCol: dateIdx > 0 ? colName(dateIdx) : f.docCol,
    };
  });
  sheet.getRange(`A5:A${4 + qcForms.length}`).values = qcForms.map((f) => [f.label]);
  const formulaRows = qcForms.map((f) => {
    const sh = q(f.sheet);
    return [
      `=COUNTIF(${sh}!$${f.primaryCol}$${DATA_START}:$${f.primaryCol}$${DATA_END},"<>")`,
      `=COUNTIF(${sh}!$${f.statusCol}$${DATA_START}:$${f.statusCol}$${DATA_END},"${VALID_STATUS}")`,
      `=COUNTIF(${sh}!$${f.alertCol}$${DATA_START}:$${f.alertCol}$${DATA_END},"ALERTE*")`,
      `=COUNTIFS(${sh}!$${f.primaryCol}$${DATA_START}:$${f.primaryCol}$${DATA_END},"<>",${sh}!$${f.statusCol}$${DATA_START}:$${f.statusCol}$${DATA_END},"<>${VALID_STATUS}")`,
      `=COUNTIFS(${sh}!$${f.primaryCol}$${DATA_START}:$${f.primaryCol}$${DATA_END},"<>",${sh}!$${f.docCol}$${DATA_START}:$${f.docCol}$${DATA_END},"")`,
      `=COUNTIFS(${sh}!$${f.primaryCol}$${DATA_START}:$${f.primaryCol}$${DATA_END},"<>",${sh}!$${f.sourceCol}$${DATA_START}:$${f.sourceCol}$${DATA_END},"")`,
      `=COUNTIFS(${sh}!$${f.primaryCol}$${DATA_START}:$${f.primaryCol}$${DATA_END},"<>",${sh}!$${f.dateCol}$${DATA_START}:$${f.dateCol}$${DATA_END},"")`,
      `=IF(D${5 + qcForms.indexOf(f)}>0,"À corriger",IF(B${5 + qcForms.indexOf(f)}=C${5 + qcForms.indexOf(f)},"OK","Validation en attente"))`,
      `=IF(I${5 + qcForms.indexOf(f)}="OK","Aucune",IF(F${5 + qcForms.indexOf(f)}>0,"Joindre le document source",IF(G${5 + qcForms.indexOf(f)}>0,"Compléter la source",IF(E${5 + qcForms.indexOf(f)}>0,"Valider humainement les lignes","Corriger les alertes QC"))))`,
    ];
  });
  sheet.getRange(`B5:J${4 + qcForms.length}`).formulas = formulaRows;
  sheet.getRange(`A5:J${4 + qcForms.length}`).format = { fill: palette.white, wrapText: true };
  sheet.tables.add(`A4:J${4 + qcForms.length}`, true, "tblControleQualiteValidation").style = "TableStyleMedium2";
  addStatusFormatting(sheet, `I5:I${4 + qcForms.length}`);
  sheet.freezePanes.freezeRows(HEADER_ROW);
}

buildQualityControl();

function buildModeEmploi() {
  const sheet = addSheet(
    "43_MODE_EMPLOI",
    "Mode d'emploi utilisateur",
    "Guide opérationnel pour saisir, contrôler, valider et reporter les données climatiques de la Ville de Dakar.",
    "H"
  );
  setWidths(sheet, [90, 260, 420, 260, 180, 180, 180, 220], 80);
  sheet.getRange("A4:H4").values = [["Étape", "Action utilisateur", "Description", "Feuille concernée", "Responsable", "Fréquence", "Point de contrôle", "Résultat"]];
  sheet.getRange("A5:H13").values = [
    [1, "Ouvrir le menu d'accueil", "Utiliser les liens internes vers les formulaires de saisie ou le dashboard.", "00_ACCUEIL", "Tous utilisateurs", "À chaque saisie", "Choisir le bon domaine", "Navigation rapide"],
    [2, "Saisir la donnée", "Renseigner les champs obligatoires, la source et le document preuve.", "00_Formulaire_Saisie ou formulaires 01-10", "Point focal sectoriel", "Selon fréquence de collecte", "Listes déroulantes", "Ligne brute saisie"],
    [3, "Joindre la preuve", "Indiquer clairement le nom du fichier, lien, rapport ou facture utilisé.", "Colonnes Document source", "Point focal sectoriel", "À chaque ligne", "Document source non vide", "Traçabilité MRV"],
    [4, "Contrôler la qualité", "Corriger les alertes données manquantes, valeurs anormales, source/document absents.", "42_CONTROLE_QUALITE", "Cellule MRV", "Mensuelle", "Alerte_QC = OK", "Ligne prête à valider"],
    [5, "Valider humainement", `Changer le statut en "${VALID_STATUS}" après revue technique.`, "Toutes feuilles Form_*", "Validateur désigné", "Mensuelle", "Statut validé", "Copie automatique vers Data_*"],
    [6, "Consulter la base consolidée", "Les bases Data_* reprennent automatiquement uniquement les lignes validées humainement.", "Data_*", "Cellule MRV", "Continu", "Exclusion des brouillons", "Base fiable"],
    [7, "Lire le dashboard", "Les KPI, alertes et graphiques se mettent à jour à partir des bases validées.", "41_DASHBOARD_FORMS; 46_GRAPHIQUES_AUTO", "Cellule climat", "Mensuelle", "Aucune alerte critique", "Pilotage exécutif"],
    [8, "Préparer les rapports", "Utiliser les vues CDP/C40/ICLEI et CDN pour remplir les plateformes externes.", "Reporting_CDP_C40_ICLEI; Contribution_CDN_2_0", "Cellule reporting", "Annuelle", "Preuves disponibles", "Rapports défendables"],
    [9, "Archiver la version", "Documenter les changements et conserver les preuves de validation.", "25_JOURNAL_MODIFS; 34_REGISTRE_PREUVES_MRV", "Administrateur Hub", "À chaque cycle", "Audit trail", "Historique complet"],
  ];
  sheet.getRange("A4:H4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange("A5:H13").format = { fill: palette.white, wrapText: true };
  sheet.tables.add("A4:H13", true, "tblModeEmploiForms").style = "TableStyleMedium2";
}

function buildFlux() {
  const sheet = addSheet(
    "44_FLUX_DONNEES",
    "Flux de production des données",
    "Architecture fonctionnelle: Document -> Feuille Excel -> Tableau de bord et Formulaire -> Base de données -> Contrôle qualité -> Validation -> Dashboard -> Rapports.",
    "H"
  );
  setWidths(sheet, [80, 230, 330, 250, 220, 220, 160, 260], 80);
  sheet.getRange("A4:H4").values = [["Ordre", "Bloc", "Rôle", "Entrée", "Traitement", "Sortie", "Responsable", "Règle de blocage"]];
  sheet.getRange("A5:H12").values = [
    [1, "Document source", "Preuve officielle ou opérationnelle", "Facture, rapport, fichier capteur, extrait opérateur", "Référencement dans le formulaire", "Document source traçable", "Point focal", "Aucune ligne sans preuve pour données critiques"],
    [2, "Formulaire", "Interface de saisie par domaine", "Donnée brute et métadonnées", "Listes déroulantes, champs obligatoires, calculs automatiques", "Ligne saisie", "Point focal", "Champs obligatoires non vides"],
    [3, "Base de données", "Stockage consolidé des données validées", "Formulaire sectoriel", `Filtre automatique sur statut ${VALID_STATUS}`, "Data_*", "Cellule MRV", "Exclusion automatique des brouillons"],
    [4, "Contrôle qualité", "Détection des anomalies et preuves manquantes", "Formulaires et Data_*", "Tests manque, anomalie, source, document, statut", "Liste d'alertes", "Cellule MRV", "Alerte_QC doit être OK"],
    [5, "Validation", "Approbation humaine avant reporting", "Lignes corrigées", "Revue par validateur désigné", "Ligne validée", "Validateur technique", `Statut obligatoire: ${VALID_STATUS}`],
    [6, "Dashboard", "Pilotage exécutif et suivi des indicateurs", "Data_* validées", "Agrégations et graphiques", "KPI et alertes", "Cellule climat", "Données non validées exclues"],
    [7, "Rapports", "Production CDP/C40/ICLEI/CDN/bailleurs", "Vues reporting", "Contrôle cohérence et preuves", "Rapports finaux", "Cellule reporting", "Preuves et validations requises"],
    [8, "Journal / audit", "Historique des changements et validations", "Versions et corrections", "Archivage", "Audit trail", "Administrateur Hub", "Mise à jour obligatoire à chaque cycle"],
  ];
  sheet.getRange("A4:H4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange("A5:H12").format = { fill: palette.white, wrapText: true };
  sheet.tables.add("A4:H12", true, "tblFluxDonneesForms").style = "TableStyleMedium2";
}

function buildPriorities() {
  const sheet = addSheet(
    "45_DONNEES_PRIORITAIRES",
    "Liste des données prioritaires à collecter",
    "Données à sécuriser en priorité pour alimenter CDP, C40, ICLEI, CDN et les bailleurs climat.",
    "I"
  );
  setWidths(sheet, [60, 230, 210, 210, 180, 140, 135, 160, 300], 90);
  const rows = [
    ["P0", "Inventaire GES territorial", "Data_Emissions_GES", "DEEC, SENELEC, CETUD, SONAGED/UCG", "Référent GES", "Annuelle", "Très élevée", "CDP/C40/CDN", "Valider périmètre GPC BASIC et facteurs d'émission"],
    ["P0", "Consommation électrique municipale", "Data_Electricite", "SENELEC, factures Ville", "Direction Énergie", "Mensuelle", "Moyenne", "GES/CDN", "Centraliser factures et sites municipaux"],
    ["P0", "Déchets collectés et destinations", "Data_Dechets", "SONAGED/UCG, prestataires", "Direction Déchets", "Mensuelle", "Élevée", "CDP/C40/GES", "Harmoniser tonnages, recyclage, compostage et décharge"],
    ["P0", "Mobilité et carburants", "Data_Mobilite", "CETUD, AFTU, flotte municipale", "Direction Mobilité", "Trimestrielle", "Très élevée", "C40/CDN/GES", "Collecter consommation et passagers par mode"],
    ["P0", "Projets climat et budgets", "Data_Projets_Climat", "Directions municipales", "Cellule climat", "Mensuelle", "Moyenne", "CDP/C40/Bailleurs", "Attribuer ID projet, statut et financement"],
    ["P0", "Documents sources et preuves", "Toutes Form_*", "Toutes institutions", "Cellule MRV", "Continu", "Moyenne", "Tous cadres", "Créer dossier de preuves et registre MRV"],
    ["P1", "Qualité de l'air", "Data_Qualite_Air", "DEEC/CGQA, capteurs", "Direction Environnement", "Mensuelle", "Élevée", "C40/CDP", "Valider protocole QA/QC capteurs"],
    ["P1", "Pluviométrie et chaleur", "Data_Temperature; Data_Pluviometrie", "ANACIM", "Cellule adaptation", "Mensuelle", "Moyenne", "CDP/CDN", "Consolider séries météo locales"],
    ["P1", "Érosion côtière et exposition", "Data_Erosion_Cotiere", "CSE, universités, SIG Ville", "Planification/SIG", "Semestrielle", "Très élevée", "CDN/Bailleurs", "Cartographier recul et population exposée"],
    ["P1", "Finance climat mobilisée", "Data_Finance_Climat", "Direction Finances, bailleurs", "Direction Finances", "Mensuelle", "Élevée", "GCF/WB/BAD", "Suivre demandé, approuvé, décaissé et cofinancement"],
  ];
  sheet.getRange("A4:I4").values = [["Priorité", "Donnée", "Feuille cible", "Source", "Responsable", "Fréquence", "Difficulté", "Criticité", "Première action"]];
  sheet.getRange(`A5:I${4 + rows.length}`).values = rows;
  sheet.getRange("A4:I4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:I${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  sheet.tables.add(`A4:I${4 + rows.length}`, true, "tblDonneesPrioritairesForms").style = "TableStyleMedium2";
  addStatusFormatting(sheet, `A5:A${4 + rows.length}`);
}

buildModeEmploi();
buildFlux();
buildPriorities();

function sumifsByMonthFormula(dataSheet, valueCol, monthCol, monthCell) {
  return `=SUMIFS(${dataSheet}!$${valueCol}$${DATA_START}:$${valueCol}$${DATA_END},${dataSheet}!$${monthCol}$${DATA_START}:$${monthCol}$${DATA_END},${monthCell})`;
}

function averageifsByMonthFormula(dataSheet, valueCol, monthCol, monthCell) {
  return `=IFERROR(AVERAGEIFS(${dataSheet}!$${valueCol}$${DATA_START}:$${valueCol}$${DATA_END},${dataSheet}!$${monthCol}$${DATA_START}:$${monthCol}$${DATA_END},${monthCell}),0)`;
}

function buildCharts() {
  const sheet = addSheet(
    "46_GRAPHIQUES_AUTO",
    "Graphiques automatiques",
    "Graphiques natifs Excel alimentés par les bases validées. Les graphiques restent à zéro tant que les formulaires n'ont pas de données validées.",
    "V"
  );
  setWidths(sheet, [105, 125, 105, 125, 20, 145, 145, 145, 145, 145, 20, 145, 145, 145, 145, 145, 20, 145, 145, 145, 145, 145], 100);
  const blocks = [
    { title: "Température moyenne mensuelle", row: 4, col: 1, type: "line", header: "Température °C", formulas: (r) => averageifsByMonthFormula("Data_Temperature", "G", "D", `A${r}`), pos: ["F4", "K18"], format: "0.0" },
    { title: "Pluviométrie mensuelle", row: 20, col: 1, type: "line", header: "Pluie mm", formulas: (r) => sumifsByMonthFormula("Data_Pluviometrie", "G", "D", `A${r}`), pos: ["F20", "K34"], format: "#,##0" },
    { title: "Consommation électrique", row: 36, col: 1, type: "line", header: "kWh", formulas: (r) => sumifsByMonthFormula("Data_Electricite", "F", "C", `A${r}`), pos: ["F36", "K50"], format: "#,##0" },
    { title: "Émissions GES annuelles", row: 52, col: 1, type: "bar", header: "tCO2e", formulas: (r) => `=SUMIFS(Data_Emissions_GES!$H$${DATA_START}:$H$${DATA_END},Data_Emissions_GES!$B$${DATA_START}:$B$${DATA_END},A${r})`, labels: lists.annees.slice(0, 7), pos: ["F52", "K66"], format: "#,##0" },
    { title: "Déchets collectés", row: 4, col: 12, type: "bar", header: "tonnes", formulas: (r) => sumifsByMonthFormula("Data_Dechets", "F", "C", `L${r}`), pos: ["Q4", "V18"], format: "#,##0" },
    { title: "Déchets valorisés", row: 20, col: 12, type: "bar", header: "tonnes", formulas: (r) => `=SUMIFS(Data_Dechets!$G$${DATA_START}:$G$${DATA_END},Data_Dechets!$C$${DATA_START}:$C$${DATA_END},L${r})+SUMIFS(Data_Dechets!$H$${DATA_START}:$H$${DATA_END},Data_Dechets!$C$${DATA_START}:$C$${DATA_END},L${r})`, pos: ["Q20", "V34"], format: "#,##0" },
    { title: "Indice qualité de l'air", row: 36, col: 12, type: "line", header: "indice", formulas: (r) => averageifsByMonthFormula("Data_Qualite_Air", "K", "C", `L${r}`), pos: ["Q36", "V50"], format: "0.0" },
    { title: "Recul du trait de côte", row: 52, col: 12, type: "line", header: "m", formulas: (r) => averageifsByMonthFormula("Data_Erosion_Cotiere", "E", "C", `L${r}`), pos: ["Q52", "V66"], format: "0.0" },
    { title: "Budget climat mobilisé", row: 68, col: 1, type: "bar", header: "FCFA approuvés", formulas: (r) => sumifsByMonthFormula("Data_Finance_Climat", "H", "C", `A${r}`), pos: ["F68", "K82"], format: "#,##0" },
    { title: "Avancement projets climat", row: 68, col: 12, type: "bar", header: "Avancement", formulas: (r) => `=IFERROR(INDEX(Data_Projets_Climat!$L$${DATA_START}:$L$${DATA_END},ROW()-67),0)`, labels: ["Projet 1", "Projet 2", "Projet 3", "Projet 4", "Projet 5"], pos: ["Q68", "V82"], format: "0%" },
  ];

  for (const block of blocks) {
    const startCol = colName(block.col);
    const valueCol = colName(block.col + 1);
    const titleRange = `${startCol}${block.row}:${valueCol}${block.row}`;
    sheet.getRange(titleRange).merge();
    sheet.getRange(`${startCol}${block.row}`).values = [[block.title]];
    sheet.getRange(titleRange).format = { fill: palette.navy, font: { bold: true, color: palette.white }, wrapText: true };
    sheet.getRange(`${startCol}${block.row + 1}:${valueCol}${block.row + 1}`).values = [["Période", block.header]];
    const labels = block.labels ?? lists.mois;
    sheet.getRange(`${startCol}${block.row + 2}:${startCol}${block.row + 1 + labels.length}`).values = labels.map((label) => [label]);
    const formulaRows = labels.map((_, idx) => [block.formulas(block.row + 2 + idx)]);
    sheet.getRange(`${valueCol}${block.row + 2}:${valueCol}${block.row + 1 + labels.length}`).formulas = formulaRows;
    sheet.getRange(`${startCol}${block.row + 1}:${valueCol}${block.row + 1}`).format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
    sheet.getRange(`${startCol}${block.row + 2}:${valueCol}${block.row + 1 + labels.length}`).format = { fill: palette.white, wrapText: true };
    sheet.getRange(`${valueCol}${block.row + 2}:${valueCol}${block.row + 1 + labels.length}`).setNumberFormat(block.format);
    const chart = sheet.charts.add(block.type, sheet.getRange(`${startCol}${block.row + 1}:${valueCol}${block.row + 1 + labels.length}`));
    chart.title = block.title;
    chart.hasLegend = false;
    chart.xAxis = { axisType: "textAxis" };
    chart.yAxis = { numberFormatCode: block.format };
    chart.setPosition(block.pos[0], block.pos[1]);
  }
}

buildCharts();

function updateAccueil() {
  const sheet = workbook.worksheets.getItem("00_ACCUEIL");
  sheet.getRange("A20:J20").merge();
  sheet.getRange("A20").values = [["Menu V3 Forms - saisie, base validée, contrôle qualité, dashboard et rapports"]];
  sheet.getRange("A20:J20").format = { fill: palette.navy, font: { bold: true, color: palette.white, size: 13 } };
  const links = [
    ["Saisir température", "#'01_Form_Temperature'!A1"],
    ["Saisir pluviométrie", "#'02_Form_Pluviometrie'!A1"],
    ["Saisir électricité", "#'03_Form_Electricite'!A1"],
    ["Saisir déchets", "#'04_Form_Dechets'!A1"],
    ["Saisir qualité air", "#'05_Form_Qualite_Air'!A1"],
    ["Saisir émissions GES", "#'06_Form_Emissions_GES'!A1"],
    ["Saisir érosion côtière", "#'08_Form_Erosion_Cotiere'!A1"],
    ["Saisir projets climat", "#'09_Form_Projets_Climat'!A1"],
    ["Voir dashboard", "#'41_DASHBOARD_FORMS'!A1"],
    ["Voir reporting CDP/C40/ICLEI", "#'Reporting_CDP_C40_ICLEI'!A1"],
    ["Voir contribution CDN", "#'Contribution_CDN_2_0'!A1"],
    ["Contrôle qualité", "#'42_CONTROLE_QUALITE'!A1"],
  ];
  const rows = [];
  for (let i = 0; i < links.length; i += 2) {
    rows.push([links[i]?.[0] ?? "", links[i]?.[1] ?? "", links[i + 1]?.[0] ?? "", links[i + 1]?.[1] ?? ""]);
  }
  sheet.getRange(`A22:D${21 + rows.length}`).values = rows;
  for (let r = 22; r < 22 + rows.length; r += 1) {
    sheet.getRange(`A${r}`).formulas = [[`=HYPERLINK(B${r},A${r})`]];
    sheet.getRange(`C${r}`).formulas = [[`=HYPERLINK(D${r},C${r})`]];
  }
  sheet.getRange(`A22:D${21 + rows.length}`).format = { fill: palette.lightBlue, font: { color: palette.navy, bold: true }, wrapText: true };
  sheet.getRange(`B22:B${21 + rows.length}`).format = { fill: palette.white, font: { color: palette.slate } };
  sheet.getRange(`D22:D${21 + rows.length}`).format = { fill: palette.white, font: { color: palette.slate } };
  sheet.getRange("F22:J27").values = [[`Règle obligatoire V3 Forms:\nAucune donnée ne doit alimenter les rapports finaux si son statut n'est pas "${VALID_STATUS}". Les feuilles Data_* appliquent cette règle automatiquement.`]];
  sheet.getRange("F22:J27").merge();
  sheet.getRange("F22:J27").format = { fill: palette.lightAmber, font: { bold: true, color: "#92400E" }, wrapText: true };
}

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
  range: "41_DASHBOARD_FORMS!A4:N17",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 14,
});
console.log("DASHBOARD_CHECK");
console.log(dashboardInspect.ndjson);

for (const renderTarget of [
  "00_Formulaire_Saisie!A1:U18",
  "01_Form_Temperature!A1:T18",
  "41_DASHBOARD_FORMS!A1:N20",
  "42_CONTROLE_QUALITE!A1:J18",
  "43_MODE_EMPLOI!A1:H16",
  "44_FLUX_DONNEES!A1:H15",
  "45_DONNEES_PRIORITAIRES!A1:I16",
  "46_GRAPHIQUES_AUTO!A1:V34",
]) {
  const [sheetName, range] = renderTarget.split("!");
  const rendered = await workbook.render({ sheetName, range, scale: 1.25 });
  const safeName = renderTarget.replaceAll("!", "_").replaceAll(":", "_").replaceAll("'", "");
  await fs.writeFile(path.join(outputDir, `${safeName}.png`), new Uint8Array(await rendered.arrayBuffer()));
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

const verifySheets = await verifyWorkbook.inspect({
  kind: "workbook",
  include: "sheets",
  maxChars: 6000,
});
console.log("VERIFY_WORKBOOK");
console.log(verifySheets.ndjson);
console.log(`SAVED ${outputFile}`);
