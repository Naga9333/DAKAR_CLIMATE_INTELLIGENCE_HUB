import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputFile = path.resolve("outputs", "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V5_RENFORCEMENT.xlsx");
const sourceFile = path.resolve("outputs", "v6_source_extraction.json");
const outputDir = path.resolve("outputs");
const outputFile = path.join(outputDir, "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V6_SIMPLIFIE_BLOCS_PROJECTIONS.xlsx");

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputFile));
let sources = [];
try {
  sources = JSON.parse(await fs.readFile(sourceFile, "utf8"));
} catch {
  sources = [];
}

const palette = {
  navy: "#0B1F33",
  teal: "#0F766E",
  green: "#166534",
  amber: "#F59E0B",
  red: "#DC2626",
  lightTeal: "#E6F4F1",
  lightGreen: "#EAF7EA",
  lightAmber: "#FFF7E6",
  lightRed: "#FEE2E2",
  white: "#FFFFFF",
  slate: "#334155",
};

const YEAR_START = 2020;
const YEAR_END = 2035;
const YEARS = Array.from({ length: YEAR_END - YEAR_START + 1 }, (_, i) => YEAR_START + i);

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

function deleteIfExists(name) {
  const existing = workbook.worksheets.getItemOrNullObject?.(name);
  if (existing && !existing.isNullObject) existing.delete?.();
}

for (const name of [
  "V6_INTERFACE_GLOBALE",
  "V6_DONNEES_PAR_BLOC_ANNEE",
  "V6_PROJECTIONS",
  "V6_GRAPHES_ILLUSTRATIFS",
  "V6_PARAMETRES_PROJECTIONS",
  "V6_SOURCES_CSTAR_COMSSA",
  ...YEARS.map((year) => `V6_ANNEE_${year}`),
]) {
  deleteIfExists(name);
}

function addSheet(name, title, subtitle, endCol = "L") {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  sheet.getRange(`A1:${endCol}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${endCol}1`).format = { fill: palette.navy, font: { bold: true, color: palette.white, size: 15 } };
  sheet.getRange(`A1:${endCol}1`).format.rowHeightPx = 34;
  sheet.getRange(`A2:${endCol}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange(`A2:${endCol}2`).format = { fill: palette.lightTeal, font: { italic: true, color: palette.slate }, wrapText: true };
  sheet.getRange(`A2:${endCol}2`).format.rowHeightPx = 44;
  return sheet;
}

function setWidths(sheet, widths, rowLimit = 200) {
  widths.forEach((w, index) => {
    const col = colName(index + 1);
    sheet.getRange(`${col}1:${col}${rowLimit}`).format.columnWidthPx = w;
  });
}

function addTable(sheet, range, name, style = "TableStyleMedium2") {
  const table = sheet.tables.add(range, true, name);
  table.style = style;
  table.showFilterButton = true;
  return table;
}

function addStatusFormatting(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.conditionalFormats.add("containsText", { text: "À saisir", format: { fill: palette.lightAmber, font: { color: "#92400E", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "OK", format: { fill: palette.lightGreen, font: { color: "#166534", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "P0", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "P1", format: { fill: palette.lightAmber, font: { color: "#92400E", bold: true } } });
}

const blockDefaults = {
  "Gouvernance Climat": ["Ville de Dakar / Cellule climat", "Cabinet du Maire / Cellule climat", "P0", 0.02],
  "Données socio-économiques": ["ANSD / Ville de Dakar", "Planification", "P0", 0.025],
  "Climat et météorologie": ["ANACIM", "Cellule adaptation", "P0", 0.015],
  "Risques climatiques": ["ANACIM / CSE / ONAS / Ville", "Planification / SIG", "P0", 0.03],
  "Eau": ["SEN'EAU / SONES / ONAS / DGPRE", "Eau et assainissement", "P1", 0.02],
  "Énergie": ["SENELEC / Ville de Dakar", "Direction Énergie", "P0", 0.03],
  "Inventaire GES": ["Cellule MRV / partenaires sectoriels", "Référent GES", "P0", -0.03],
  "Mobilité": ["CETUD / AFTU / DDD / TER / Ville", "Direction Mobilité", "P0", 0.04],
  "Qualité de l'air": ["DEEC / CGQA", "Direction Environnement", "P1", -0.01],
  "Déchets": ["SONAGED / UCG / Ville", "Direction Déchets", "P0", 0.03],
  "Espaces verts et biodiversité": ["Direction Environnement / SIG", "Espaces verts", "P1", 0.05],
  "Agriculture urbaine": ["Ville / ONG / projets", "Environnement / Affaires sociales", "P1", 0.05],
  "Santé et climat": ["Ministère Santé / districts", "Direction Santé", "P1", 0.02],
  "Finance climat": ["Direction Finances / bailleurs", "Direction Finances", "P0", 0.08],
  "Portefeuille de projets climatiques": ["PMO climat / directions techniques", "Cellule climat", "P0", 0.06],
};

const blockItems = {
  "Gouvernance Climat": [
    ["Existence d'un Plan Climat", "Oui/Non"],
    ["Date d'adoption", "date"],
    ["Budget climat annuel", "FCFA"],
    ["Dépenses climat", "FCFA"],
    ["Personnel climat", "ETP"],
    ["Nombre d'agents formés", "nombre"],
    ["Nombre de réunions COPIL", "nombre"],
    ["Nombre de consultations publiques", "nombre"],
    ["Nombre de partenaires", "nombre"],
    ["Existence d'un système MRV", "Oui/Non"],
    ["Existence d'un budget vert", "Oui/Non"],
    ["Existence d'un inventaire GES", "Oui/Non"],
  ],
  "Données socio-économiques": [
    ["Population totale", "personnes"],
    ["Population par sexe", "personnes"],
    ["Population jeune", "personnes"],
    ["Population âgée", "personnes"],
    ["Densité", "hab/km²"],
    ["Nombre de ménages", "ménages"],
    ["PIB local estimé", "FCFA"],
    ["Taux de pauvreté", "%"],
    ["Taux de chômage", "%"],
    ["Population vulnérable", "personnes"],
    ["Population exposée aux risques", "personnes"],
  ],
  "Climat et météorologie": [
    ["Température moyenne", "°C"],
    ["Température minimale", "°C"],
    ["Température maximale", "°C"],
    ["Température ressentie", "°C"],
    ["Nombre de jours chauds", "jours"],
    ["Nombre de vagues de chaleur", "nombre"],
    ["Pluviométrie", "mm"],
    ["Nombre de jours de pluie", "jours"],
    ["Intensité des pluies", "mm/h"],
    ["Humidité", "%"],
    ["Vitesse du vent", "m/s"],
    ["Direction du vent", "degrés"],
    ["Rayonnement solaire", "W/m²"],
    ["Évapotranspiration", "mm"],
    ["Niveau marin", "m"],
  ],
  "Risques climatiques": [
    ["Inondations - fréquence", "événements/an"],
    ["Inondations - intensité", "indice"],
    ["Inondations - population exposée", "personnes"],
    ["Inondations - pertes économiques", "FCFA"],
    ["Érosion côtière - fréquence", "événements/an"],
    ["Érosion côtière - intensité", "m/an"],
    ["Érosion côtière - population exposée", "personnes"],
    ["Érosion côtière - pertes économiques", "FCFA"],
    ["Submersion marine - population exposée", "personnes"],
    ["Canicules - jours extrêmes", "jours"],
    ["Sécheresse - indice", "indice"],
    ["Tempêtes - événements", "nombre"],
    ["Feux de végétation - événements", "nombre"],
    ["Pollution atmosphérique - épisodes", "jours"],
    ["Maladies climatiques - cas", "cas"],
    ["Stress hydrique - population exposée", "personnes"],
  ],
  "Eau": [
    ["Production d'eau", "m³"],
    ["Consommation d'eau", "m³"],
    ["Pertes réseau", "%"],
    ["Accès à l'eau potable", "%"],
    ["Qualité de l'eau", "indice"],
    ["Nappes phréatiques", "niveau/qualité"],
    ["Réservoirs", "nombre/m³"],
    ["Forages", "nombre"],
    ["Eaux usées produites", "m³"],
    ["Eaux usées traitées", "m³"],
  ],
  "Énergie": [
    ["Consommation totale", "kWh"],
    ["Consommation résidentielle", "kWh"],
    ["Consommation commerciale", "kWh"],
    ["Consommation industrielle", "kWh"],
    ["Consommation municipale", "kWh"],
    ["Éclairage public", "kWh"],
    ["Production solaire", "kWh"],
    ["Production éolienne", "kWh"],
    ["Production biomasse", "kWh"],
    ["Accès à l'électricité", "%"],
    ["Coupures", "heures/an"],
    ["Facteurs d'émission", "kgCO2e/kWh"],
  ],
  "Inventaire GES": [
    ["Énergie", "tCO2e"],
    ["Combustibles", "tCO2e"],
    ["Électricité", "tCO2e"],
    ["Transport", "tCO2e"],
    ["Routier", "tCO2e"],
    ["Ferroviaire", "tCO2e"],
    ["Maritime", "tCO2e"],
    ["Aérien local", "tCO2e"],
    ["Déchets", "tCO2e"],
    ["Décharge", "tCO2e"],
    ["Recyclage", "tCO2e"],
    ["Compostage", "tCO2e"],
    ["AFOLU", "tCO2e"],
    ["Agriculture urbaine", "tCO2e"],
    ["Espaces verts", "tCO2e séquestration estimée"],
    ["Résultats", "tCO2e"],
    ["Scope 1", "tCO2e"],
    ["Scope 2", "tCO2e"],
    ["Scope 3", "tCO2e"],
    ["tCO₂e par secteur", "tCO2e"],
    ["tCO₂e par habitant", "tCO2e/hab"],
  ],
  "Mobilité": [
    ["Nombre de véhicules", "nombre"],
    ["Véhicules particuliers", "nombre"],
    ["Taxis", "nombre"],
    ["Bus", "nombre"],
    ["Cars rapides", "nombre"],
    ["Ndiaga Ndiaye", "nombre"],
    ["Motos", "nombre"],
    ["Camions", "nombre"],
    ["Transport collectif", "passagers/an"],
    ["Passagers BRT", "passagers/an"],
    ["Passagers TER", "passagers/an"],
    ["Passagers DDD", "passagers/an"],
    ["Mobilité active", "%"],
    ["Marche", "%"],
    ["Vélo", "%"],
    ["Pistes cyclables", "km"],
    ["Électromobilité", "indice"],
    ["Véhicules électriques", "nombre"],
    ["Bornes de recharge", "nombre"],
  ],
  "Qualité de l'air": [
    ["PM2.5", "µg/m3"],
    ["PM10", "µg/m3"],
    ["NO₂", "µg/m3"],
    ["SO₂", "µg/m3"],
    ["CO", "mg/m3"],
    ["O₃", "µg/m3"],
    ["Black Carbon", "µg/m3"],
    ["Nombre de stations", "nombre"],
    ["Dépassements OMS", "jours/an"],
  ],
  "Déchets": [
    ["Déchets collectés", "tonnes"],
    ["Déchets ménagers", "tonnes"],
    ["Déchets commerciaux", "tonnes"],
    ["Déchets industriels", "tonnes"],
    ["Déchets plastiques", "tonnes"],
    ["Déchets organiques", "tonnes"],
    ["Valorisation", "tonnes"],
    ["Recyclés", "tonnes"],
    ["Compostés", "tonnes"],
    ["Réutilisés", "tonnes"],
    ["Valorisés", "tonnes"],
    ["Mise en décharge", "tonnes"],
    ["Quantité enfouie", "tonnes"],
    ["Méthane estimé", "tCH4"],
  ],
  "Espaces verts et biodiversité": [
    ["Surface espaces verts", "m²"],
    ["m²/habitant", "m²/hab"],
    ["Parcs", "nombre"],
    ["Jardins publics", "nombre"],
    ["Jardins communautaires", "nombre"],
    ["Micro-jardins", "nombre"],
    ["Forêts urbaines", "ha"],
    ["Corridors verts", "km"],
    ["Zones humides", "ha"],
    ["Mangroves", "ha"],
    ["Arbres plantés", "nombre"],
    ["Arbres survivants", "nombre"],
    ["Arbres coupés", "nombre"],
    ["Taux de survie", "%"],
    ["CO2 séquestré estimé", "tCO2e"],
  ],
  "Agriculture urbaine": [
    ["Surface cultivée", "ha"],
    ["Production", "tonnes"],
    ["Nombre de bénéficiaires", "personnes"],
    ["Rendement", "tonnes/ha"],
    ["Irrigation", "m³"],
  ],
  "Santé et climat": [
    ["Paludisme", "cas/an"],
    ["Dengue", "cas/an"],
    ["Choléra", "cas/an"],
    ["Maladies respiratoires", "cas/an"],
    ["Hospitalisations chaleur", "hospitalisations/an"],
    ["Mortalité chaleur", "décès/an"],
    ["Mortalité pollution", "décès/an"],
  ],
  "Finance climat": [
    ["Budget", "FCFA"],
    ["Budget climat", "FCFA"],
    ["Dépenses climat", "FCFA"],
    ["Financement UE", "FCFA"],
    ["Financement AFD", "FCFA"],
    ["Financement BAD", "FCFA"],
    ["Financement Banque Mondiale", "FCFA"],
    ["Financement GCF", "FCFA"],
    ["Financement C40", "FCFA"],
    ["Financement ICLEI", "FCFA"],
    ["Nombre de projets", "nombre"],
    ["Coût portefeuille projets", "FCFA"],
    ["Financement obtenu", "FCFA"],
    ["Financement recherché", "FCFA"],
  ],
  "Portefeuille de projets climatiques": [
    ["Nom", "texte"],
    ["Secteur", "texte"],
    ["Description", "texte"],
    ["Localisation", "texte"],
    ["Budget", "FCFA"],
    ["Coût actualisé", "FCFA"],
    ["État d'avancement", "%"],
    ["Maturité", "niveau"],
    ["Bailleur cible", "texte"],
    ["Réduction GES attendue", "tCO2e/an"],
    ["Population bénéficiaire", "personnes"],
    ["Emplois créés", "emplois"],
    ["Risques", "texte"],
    ["Source de financement", "texte"],
    ["Date de début", "date"],
    ["Date de fin", "date"],
  ],
};

const blockCodes = Object.fromEntries(Object.keys(blockItems).map((block, i) => [block, `B${String(i + 1).padStart(2, "0")}`]));

function buildRows() {
  const rows = [];
  for (const [block, items] of Object.entries(blockItems)) {
    const [source, responsible, priority, growth] = blockDefaults[block];
    items.forEach(([name, unit], i) => {
      rows.push([`${blockCodes[block]}-${String(i + 1).padStart(3, "0")}`, block, name, unit, source, responsible, priority, ...YEARS.map(() => ""), null, null, growth, null, null, null]);
    });
  }
  return rows;
}

const dataRows = buildRows();
const yearStartCol = 8; // H
const yearEndCol = yearStartCol + YEARS.length - 1; // W
const latestCol = yearEndCol + 1; // X
const latestYearCol = yearEndCol + 2; // Y
const growthCol = yearEndCol + 3; // Z
const projection2030Col = yearEndCol + 4; // AA
const projection2035Col = yearEndCol + 5; // AB
const statusCol = yearEndCol + 6; // AC

function q(sheetName) {
  return `'${sheetName.replaceAll("'", "''")}'`;
}

function buildParameters() {
  const sheet = addSheet(
    "V6_PARAMETRES_PROJECTIONS",
    "Paramètres de projection par bloc",
    "Hypothèses simples et modifiables pour générer des projections illustratives par bloc. Les taux ne remplacent pas une modélisation sectorielle.",
    "F"
  );
  const rows = Object.entries(blockDefaults).map(([block, vals]) => [block, vals[3], "Taux annuel indicatif", vals[0], vals[1], vals[2]]);
  sheet.getRange(`A4:F${4 + rows.length}`).values = [["Bloc", "Taux annuel projection", "Méthode", "Source principale", "Responsable", "Priorité"], ...rows];
  sheet.getRange("A4:F4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:F${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:F${4 + rows.length}`, "tblParametresProjectionV6");
  setWidths(sheet, [240, 155, 230, 280, 220, 90], 80);
  sheet.getRange(`B5:B${4 + rows.length}`).setNumberFormat("0.0%");
  addStatusFormatting(sheet, `F5:F${4 + rows.length}`);
}

function buildDataMatrix() {
  const sheet = addSheet(
    "V6_DONNEES_PAR_BLOC_ANNEE",
    "Données par bloc et par année",
    "Matrice simplifiée: les lignes sont les données demandées, les colonnes sont les années. Saisir les valeurs dans les colonnes 2020-2035; l'interface calcule ensuite synthèse, projections et graphes.",
    "AC"
  );
  const headers = ["Code", "Bloc", "Donnée", "Unité", "Source principale", "Responsable", "Priorité", ...YEARS, "Dernière valeur", "Année dernière valeur", "Taux projection", "Projection 2030", "Projection 2035", "Statut"];
  const endRow = 4 + dataRows.length;
  sheet.getRange(`A4:AC${endRow}`).values = [headers, ...dataRows];
  for (let i = 0; i < YEARS.length; i += 1) {
    const year = YEARS[i];
    const col = colName(yearStartCol + i);
    const formulas = [];
    for (let row = 5; row <= endRow; row += 1) {
      formulas.push([`=IF(${q(`V6_ANNEE_${year}`)}!$H${row}="","",${q(`V6_ANNEE_${year}`)}!$H${row})`]);
    }
    sheet.getRange(`${col}5:${col}${endRow}`).formulas = formulas;
  }
  const formulaRows = [];
  for (let row = 5; row <= endRow; row += 1) {
    const yearRange = `H${row}:W${row}`;
    formulaRows.push([
      `=IFERROR(LOOKUP(2,1/(${yearRange}<>""),${yearRange}),"")`,
      `=IF(X${row}="","",LOOKUP(2,1/(${yearRange}<>""),$H$4:$W$4))`,
      `=IFERROR(XLOOKUP(B${row},V6_PARAMETRES_PROJECTIONS!$A$5:$A$19,V6_PARAMETRES_PROJECTIONS!$B$5:$B$19),0)`,
      `=IF(OR(X${row}="",NOT(ISNUMBER(X${row}))),"",IF(Y${row}>=2030,IFERROR(INDEX(${yearRange},1,MATCH(2030,$H$4:$W$4,0)),X${row}),X${row}*(1+Z${row})^(2030-Y${row})))`,
      `=IF(OR(X${row}="",NOT(ISNUMBER(X${row}))),"",IF(Y${row}>=2035,IFERROR(INDEX(${yearRange},1,MATCH(2035,$H$4:$W$4,0)),X${row}),X${row}*(1+Z${row})^(2035-Y${row})))`,
      `=IF(X${row}="","À saisir","OK")`,
    ]);
  }
  sheet.getRange(`X5:AC${endRow}`).formulas = formulaRows;
  sheet.getRange("A4:AC4").format = { fill: palette.green, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:AC${endRow}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:AC${endRow}`, "tblDonneesBlocAnneeV6", "TableStyleMedium4");
  setWidths(sheet, [95, 220, 260, 110, 250, 210, 90, ...YEARS.map(() => 90), 130, 135, 120, 130, 130, 105], endRow + 5);
  sheet.freezePanes.freezeRows(4);
  sheet.freezePanes.freezeColumns(3);
  sheet.getRange(`H5:AB${endRow}`).setNumberFormat("#,##0.00");
  sheet.getRange(`Z5:Z${endRow}`).setNumberFormat("0.0%");
  addStatusFormatting(sheet, `G5:G${endRow}`);
  addStatusFormatting(sheet, `AC5:AC${endRow}`);
}

function buildYearSheets() {
  const headers = ["Code", "Bloc", "Donnée", "Unité", "Source principale", "Responsable", "Priorité", "Résultat", "Document justificatif", "Statut validation", "Commentaires"];
  for (const year of YEARS) {
    const sheet = addSheet(
      `V6_ANNEE_${year}`,
      `Saisie annuelle ${year}`,
      `Feuille de saisie simplifiée pour l'année ${year}. Renseigner la colonne Résultat, le document justificatif et le statut de validation; la matrice globale se mettra à jour automatiquement.`,
      "K"
    );
    const rows = dataRows.map((r) => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], "", "", "Brouillon", ""]);
    const endRow = 4 + rows.length;
    sheet.getRange(`A4:K${endRow}`).values = [headers, ...rows];
    sheet.getRange("A4:K4").format = { fill: palette.green, font: { bold: true, color: palette.white }, wrapText: true };
    sheet.getRange(`A5:K${endRow}`).format = { fill: palette.white, wrapText: true };
    addTable(sheet, `A4:K${endRow}`, `tblAnnee${year}`, "TableStyleMedium4");
    setWidths(sheet, [95, 220, 260, 110, 250, 210, 90, 130, 240, 150, 260], endRow + 5);
    sheet.freezePanes.freezeRows(4);
    sheet.freezePanes.freezeColumns(3);
    sheet.getRange(`H5:H${endRow}`).setNumberFormat("#,##0.00");
    addStatusFormatting(sheet, `G5:G${endRow}`);
    addStatusFormatting(sheet, `J5:J${endRow}`);
    sheet.getRange(`J5:J${endRow}`).dataValidation = {
      rule: { type: "list", values: ["Brouillon", "À vérifier", "Validé techniquement", "Validé humainement", "Rejeté"] },
      prompt: { showPrompt: true, title: "Statut", message: "Choisir le statut de validation." },
      errorAlert: { showAlert: true, style: "warning", title: "Statut inattendu", message: "Choisir une valeur autorisée." },
    };
  }
}

function buildInterface() {
  const sheet = addSheet(
    "V6_INTERFACE_GLOBALE",
    "Interface globale de synthèse",
    "Synthèse simplifiée du Hub: couverture des données par bloc, dernières valeurs, projections 2030/2035 et statut de peuplement.",
    "N"
  );
  setWidths(sheet, [235, 135, 135, 130, 135, 135, 150, 260, 235, 135, 135, 130, 135, 260], 90);
  sheet.getRange("A4:H4").values = [["Bloc", "Données listées", "Valeurs saisies", "Couverture", "Projection 2030", "Projection 2035", "Priorité", "Lecture"]];
  const blocks = Object.keys(blockItems);
  const values = blocks.map((block) => [block, null, null, null, null, null, blockDefaults[block][2], null]);
  sheet.getRange(`A5:H${4 + blocks.length}`).values = values;
  const formulas = blocks.map((block, i) => {
    const r = 5 + i;
    return [
      `=COUNTIF(V6_DONNEES_PAR_BLOC_ANNEE!$B$5:$B$${4 + dataRows.length},A${r})`,
      `=SUMPRODUCT((V6_DONNEES_PAR_BLOC_ANNEE!$B$5:$B$${4 + dataRows.length}=A${r})*(V6_DONNEES_PAR_BLOC_ANNEE!$AC$5:$AC$${4 + dataRows.length}="OK"))`,
      `=IFERROR(C${r}/B${r},0)`,
      `=SUMIFS(V6_DONNEES_PAR_BLOC_ANNEE!$AA$5:$AA$${4 + dataRows.length},V6_DONNEES_PAR_BLOC_ANNEE!$B$5:$B$${4 + dataRows.length},A${r})`,
      `=SUMIFS(V6_DONNEES_PAR_BLOC_ANNEE!$AB$5:$AB$${4 + dataRows.length},V6_DONNEES_PAR_BLOC_ANNEE!$B$5:$B$${4 + dataRows.length},A${r})`,
      `=IF(D${r}=0,"À saisir",IF(D${r}<0.5,"Partiel","OK"))`,
    ];
  });
  sheet.getRange(`B5:F${4 + blocks.length}`).formulas = formulas.map((row) => row.slice(0, 5));
  sheet.getRange(`H5:H${4 + blocks.length}`).formulas = formulas.map((row) => [row[5]]);
  sheet.getRange("A4:H4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:H${4 + blocks.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:H${4 + blocks.length}`, "tblInterfaceGlobaleV6");
  sheet.getRange(`D5:D${4 + blocks.length}`).setNumberFormat("0%");
  sheet.getRange(`E5:F${4 + blocks.length}`).setNumberFormat("#,##0.00");
  addStatusFormatting(sheet, `G5:H${4 + blocks.length}`);

  sheet.getRange("J4:N4").values = [["KPI global", "Valeur", "Unité", "Source", "Lecture"]];
  sheet.getRange("J5:N12").formulas = [
    ['="Blocs"', `=${blocks.length}`, '="nombre"', '="Liste blocs"', '="Architecture simplifiée"'],
    ['="Données listées"', `=${dataRows.length}`, '="données"', '="Matrice par bloc"', '="Toutes les données demandées sont regroupées"'],
    ['="Valeurs saisies"', `=COUNTIF(V6_DONNEES_PAR_BLOC_ANNEE!$AC$5:$AC$${4 + dataRows.length},"OK")`, '="lignes"', '="Matrice par bloc"', '="Données ayant au moins une valeur annuelle"'],
    ['="Taux de couverture"', `=IFERROR(K7/K6,0)`, '="%"', '="Matrice par bloc"', '="Avancement du peuplement"'],
    ['="Données P0"', `=COUNTIF(V6_DONNEES_PAR_BLOC_ANNEE!$G$5:$G$${4 + dataRows.length},"P0")`, '="données"', '="Priorités"', '="À collecter en premier"'],
    ['="Sources CSTAR/CoMSSA"', `=${sources.length}`, '="fichiers"', '="Sources locales"', '="Documents intégrés comme référence"'],
    ['="Projection 2030 globale"', '=SUM(E5:E19)', '="somme illustrative"', '="Projections"', '="Somme indicative des indicateurs numériques"'],
    ['="Projection 2035 globale"', '=SUM(F5:F19)', '="somme illustrative"', '="Projections"', '="Somme indicative des indicateurs numériques"'],
  ];
  sheet.getRange("J4:N4").format = { fill: palette.navy, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange("J5:N12").format = { fill: palette.white, wrapText: true };
  sheet.getRange("K8").setNumberFormat("0%");

  sheet.getRange("J15:N15").values = [["Accès rapide aux feuilles annuelles", null, null, null, null]];
  sheet.getRange("J15:N15").merge();
  sheet.getRange("J15:N15").format = { fill: palette.green, font: { bold: true, color: palette.white } };
  const linkRows = [];
  for (let i = 0; i < YEARS.length; i += 4) {
    linkRows.push([
      YEARS[i] ? `Année ${YEARS[i]}` : "",
      YEARS[i + 1] ? `Année ${YEARS[i + 1]}` : "",
      YEARS[i + 2] ? `Année ${YEARS[i + 2]}` : "",
      YEARS[i + 3] ? `Année ${YEARS[i + 3]}` : "",
      "",
    ]);
  }
  sheet.getRange(`J16:N${15 + linkRows.length}`).values = linkRows;
  for (let row = 16; row < 16 + linkRows.length; row += 1) {
    for (let col = 10; col <= 13; col += 1) {
      const labelCell = `${colName(col)}${row}`;
      const yearText = sheet.getRange(labelCell).values?.[0]?.[0];
      if (yearText) {
        const year = String(yearText).replace("Année ", "");
        sheet.getRange(labelCell).formulas = [[`=HYPERLINK("#'V6_ANNEE_${year}'!A1","Année ${year}")`]];
      }
    }
  }
  sheet.getRange(`J16:N${15 + linkRows.length}`).format = { fill: palette.lightGreen, font: { bold: true, color: palette.navy }, wrapText: true };
}

function buildProjections() {
  const sheet = addSheet(
    "V6_PROJECTIONS",
    "Projections simples des indicateurs clés",
    "Projection illustrative à partir de la dernière valeur saisie et du taux paramétré par bloc. Les hypothèses sont modifiables dans V6_PARAMETRES_PROJECTIONS.",
    "N"
  );
  const keyIndicators = [
    "Population totale",
    "Budget climat annuel",
    "Température moyenne",
    "Pluviométrie",
    "Production d'eau",
    "Consommation totale",
    "Résultats",
    "Passagers BRT",
    "PM2.5",
    "Déchets collectés",
    "Surface espaces verts",
    "CO2 séquestré estimé",
    "Financement obtenu",
    "Nombre de projets",
    "Population exposée aux risques",
  ];
  const projectionYears = Array.from({ length: 10 }, (_, i) => 2026 + i);
  sheet.getRange(`A4:O4`).values = [["Indicateur clé", "Bloc", "Dernière valeur", "Année base", "Taux", ...projectionYears]];
  sheet.getRange(`A5:A${4 + keyIndicators.length}`).values = keyIndicators.map(x => [x]);
  const formulas = keyIndicators.map((_, i) => {
    const r = 5 + i;
    return [
      `=IFERROR(XLOOKUP(A${r},V6_DONNEES_PAR_BLOC_ANNEE!$C$5:$C$${4 + dataRows.length},V6_DONNEES_PAR_BLOC_ANNEE!$B$5:$B$${4 + dataRows.length}),"À mapper")`,
      `=IFERROR(XLOOKUP(A${r},V6_DONNEES_PAR_BLOC_ANNEE!$C$5:$C$${4 + dataRows.length},V6_DONNEES_PAR_BLOC_ANNEE!$X$5:$X$${4 + dataRows.length}),"")`,
      `=IFERROR(XLOOKUP(A${r},V6_DONNEES_PAR_BLOC_ANNEE!$C$5:$C$${4 + dataRows.length},V6_DONNEES_PAR_BLOC_ANNEE!$Y$5:$Y$${4 + dataRows.length}),"")`,
      `=IFERROR(XLOOKUP(A${r},V6_DONNEES_PAR_BLOC_ANNEE!$C$5:$C$${4 + dataRows.length},V6_DONNEES_PAR_BLOC_ANNEE!$Z$5:$Z$${4 + dataRows.length}),0)`,
      ...projectionYears.map((year) => `=IF(OR($C${r}="",NOT(ISNUMBER($C${r}))),"",$C${r}*(1+$E${r})^(${year}-$D${r}))`),
    ];
  });
  sheet.getRange(`B5:O${4 + keyIndicators.length}`).formulas = formulas;
  sheet.getRange("A4:O4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:O${4 + keyIndicators.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:O${4 + keyIndicators.length}`, "tblProjectionsV6");
  setWidths(sheet, [250, 210, 135, 105, 90, ...projectionYears.map(() => 105)], 80);
  sheet.getRange(`C5:O${4 + keyIndicators.length}`).setNumberFormat("#,##0.00");
  sheet.getRange(`E5:E${4 + keyIndicators.length}`).setNumberFormat("0.0%");
}

function buildGraphs() {
  const sheet = addSheet(
    "V6_GRAPHES_ILLUSTRATIFS",
    "Graphes illustratifs",
    "Graphiques automatiques liés à l'interface globale et aux projections. Ils se rempliront lorsque les valeurs annuelles seront saisies.",
    "X"
  );
  setWidths(sheet, [220, 120, 120, 20, 220, 120, 120, 20, 220, 120, 120, 20, 120, 120, 120, 120, 120, 120, 120, 120, 120, 120, 120, 120], 80);
  const blocks = Object.keys(blockItems);
  sheet.getRange(`A4:C${4 + blocks.length}`).formulas = [
    ["=\"Bloc\"", "=\"Couverture\"", "=\"Projection 2030\""],
    ...blocks.map((_, i) => [`=V6_INTERFACE_GLOBALE!A${5 + i}`, `=V6_INTERFACE_GLOBALE!D${5 + i}`, `=V6_INTERFACE_GLOBALE!E${5 + i}`]),
  ];
  sheet.getRange("A4:C4").format = { fill: palette.teal, font: { bold: true, color: palette.white } };
  sheet.getRange(`A5:C${4 + blocks.length}`).format = { fill: palette.white, wrapText: true };
  sheet.getRange(`B5:B${4 + blocks.length}`).setNumberFormat("0%");
  const chartCoverage = sheet.charts.add("bar", sheet.getRange(`A4:B${4 + blocks.length}`));
  chartCoverage.title = "Couverture des données par bloc";
  chartCoverage.hasLegend = false;
  chartCoverage.xAxis = { axisType: "textAxis" };
  chartCoverage.yAxis = { numberFormatCode: "0%" };
  chartCoverage.setPosition("E4", "L22");

  const chartProjection = sheet.charts.add("bar", sheet.getRange(`A4:C${4 + blocks.length}`));
  chartProjection.title = "Projection 2030 par bloc";
  chartProjection.hasLegend = false;
  chartProjection.xAxis = { axisType: "textAxis" };
  chartProjection.yAxis = { numberFormatCode: "#,##0" };
  chartProjection.setPosition("N4", "X22");

  sheet.getRange("A26:K26").formulas = [["=\"Indicateur\"", ...Array.from({ length: 10 }, (_, i) => `=${2026 + i}`)]];
  sheet.getRange("A27:K31").formulas = [
    ["=V6_PROJECTIONS!A5", ...Array.from({ length: 10 }, (_, i) => `=V6_PROJECTIONS!${colName(6 + i)}5`)],
    ["=V6_PROJECTIONS!A7", ...Array.from({ length: 10 }, (_, i) => `=V6_PROJECTIONS!${colName(6 + i)}7`)],
    ["=V6_PROJECTIONS!A11", ...Array.from({ length: 10 }, (_, i) => `=V6_PROJECTIONS!${colName(6 + i)}11`)],
    ["=V6_PROJECTIONS!A14", ...Array.from({ length: 10 }, (_, i) => `=V6_PROJECTIONS!${colName(6 + i)}14`)],
    ["=V6_PROJECTIONS!A16", ...Array.from({ length: 10 }, (_, i) => `=V6_PROJECTIONS!${colName(6 + i)}16`)],
  ];
  sheet.getRange("A26:K26").format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  sheet.getRange("A27:K31").format = { fill: palette.white, wrapText: true };
  const chartLine = sheet.charts.add("line", sheet.getRange("A26:K31"));
  chartLine.title = "Projections illustratives de quelques indicateurs clés";
  chartLine.hasLegend = true;
  chartLine.xAxis = { axisType: "textAxis" };
  chartLine.yAxis = { numberFormatCode: "#,##0" };
  chartLine.setPosition("M26", "X44");
}

function buildSources() {
  const sheet = addSheet(
    "V6_SOURCES_CSTAR_COMSSA",
    "Sources CSTAR et CoMSSA",
    "Résumé des deux sources locales utilisées pour cadrer les blocs de données.",
    "H"
  );
  const rows = sources.map((s) => [
    s.file_name ?? s.path,
    s.extension,
    s.size_kb,
    (s.sheet_names ?? []).join("; "),
    Object.entries(s.keywords ?? {}).map(([k, v]) => `${k}:${v}`).join("; "),
    s.excerpt ?? "",
    s.path,
    s.exists ? "OK" : "Absent",
  ]);
  sheet.getRange(`A4:H${4 + rows.length}`).values = [["Fichier", "Type", "Taille KB", "Feuilles détectées", "Mots-clés", "Extrait", "Chemin", "Statut"], ...rows];
  sheet.getRange("A4:H4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:H${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:H${4 + rows.length}`, "tblSourcesCstarComssaV6");
  setWidths(sheet, [330, 80, 90, 390, 340, 520, 520, 90], 20);
}

function updateAccueil() {
  const sheet = workbook.worksheets.getItemOrNullObject?.("00_ACCUEIL");
  if (!sheet || sheet.isNullObject) return;
  sheet.getRange("A38:J38").merge();
  sheet.getRange("A38").values = [["Menu V6 simplifié - données par bloc, années, synthèse et projections"]];
  sheet.getRange("A38:J38").format = { fill: palette.green, font: { bold: true, color: palette.white, size: 13 } };
  const links = [
    ["Interface globale", "#'V6_INTERFACE_GLOBALE'!A1"],
    ["Données par bloc/année", "#'V6_DONNEES_PAR_BLOC_ANNEE'!A1"],
    ["Projections", "#'V6_PROJECTIONS'!A1"],
    ["Graphes illustratifs", "#'V6_GRAPHES_ILLUSTRATIFS'!A1"],
    ["Paramètres projection", "#'V6_PARAMETRES_PROJECTIONS'!A1"],
    ["Sources CSTAR/CoMSSA", "#'V6_SOURCES_CSTAR_COMSSA'!A1"],
  ];
  const rows = [];
  for (let i = 0; i < links.length; i += 2) rows.push([links[i][0], links[i][1], links[i + 1]?.[0] ?? "", links[i + 1]?.[1] ?? ""]);
  sheet.getRange(`A40:D${39 + rows.length}`).values = rows;
  for (let row = 40; row < 40 + rows.length; row += 1) {
    sheet.getRange(`A${row}`).formulas = [[`=HYPERLINK(B${row},A${row})`]];
    sheet.getRange(`C${row}`).formulas = [[`=HYPERLINK(D${row},C${row})`]];
  }
  sheet.getRange(`A40:D${39 + rows.length}`).format = { fill: palette.lightGreen, font: { color: palette.navy, bold: true }, wrapText: true };
}

buildParameters();
buildYearSheets();
buildDataMatrix();
buildInterface();
buildProjections();
buildGraphs();
buildSources();
updateAccueil();

await fs.mkdir(outputDir, { recursive: true });

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  maxChars: 3000,
});
console.log("FORMULA_ERRORS");
console.log(errors.ndjson);

const inspect = await workbook.inspect({
  kind: "table",
  range: "V6_INTERFACE_GLOBALE!A4:N12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 14,
});
console.log("INTERFACE_CHECK");
console.log(inspect.ndjson);

for (const [sheetName, range] of [
  ["V6_INTERFACE_GLOBALE", "A1:N22"],
  ["V6_DONNEES_PAR_BLOC_ANNEE", "A1:AC22"],
  ["V6_PROJECTIONS", "A1:O22"],
  ["V6_GRAPHES_ILLUSTRATIFS", "A1:X46"],
]) {
  const rendered = await workbook.render({ sheetName, range, scale: 1.05, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), new Uint8Array(await rendered.arrayBuffer()));
  console.log(`RENDERED ${sheetName}`);
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputFile);

const verifyWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(outputFile));
const verifyErrors = await verifyWorkbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  maxChars: 3000,
});
console.log("VERIFY_FORMULA_ERRORS");
console.log(verifyErrors.ndjson);
const verifyInfo = await verifyWorkbook.inspect({ kind: "workbook", include: "sheets", maxChars: 3000 });
console.log("VERIFY_WORKBOOK");
console.log(verifyInfo.ndjson);
console.log(`SAVED ${outputFile}`);
