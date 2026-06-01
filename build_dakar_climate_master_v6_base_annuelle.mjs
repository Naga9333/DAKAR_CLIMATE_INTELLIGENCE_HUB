import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputFile = path.resolve("outputs", "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V5_RENFORCEMENT.xlsx");
const sourceFile = path.resolve("outputs", "v6_source_extraction.json");
const outputDir = path.resolve("outputs");
const outputFile = path.join(outputDir, "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V6_BASE_ANNUELLE_CSTAR_COMSSA.xlsx");

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputFile));
let sourceExtraction = [];
try {
  sourceExtraction = JSON.parse(await fs.readFile(sourceFile, "utf8"));
} catch {
  sourceExtraction = [];
}

const palette = {
  navy: "#0B1F33",
  teal: "#0F766E",
  green: "#166534",
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

const YEARS = Array.from({ length: 20 }, (_, i) => 2020 + i);
const TARGET_INDICATORS = 800;
const VALIDATION_STATUS = ["Brouillon", "À vérifier", "Validé techniquement", "Validé humainement", "Rejeté"];
const QUALITY_LEVELS = ["Faible", "Moyenne", "Élevée", "Très élevée"];
const PRIORITIES = ["P0", "P1", "P2", "P3"];
const FREQUENCIES = ["Mensuelle", "Trimestrielle", "Semestrielle", "Annuelle", "Ponctuelle"];
const FRAMEWORKS = ["CDP", "C40", "GCoM", "CoM SSA", "ICLEI", "CDN Sénégal", "PCET Dakar", "Fonds Vert Climat"];

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

for (const sheetName of [
  "V6_INDEX_BASE",
  "V6_BIBLIOTHEQUE_INDICATEURS",
  "V6_BASE_DONNEES_ANNUELLE",
  "V6_INSTITUTIONS_SOURCES",
  "V6_MAPPING_REPORTING",
  "V6_CONTROLE_QUALITE",
  "V6_DICTIONNAIRE_DONNEES",
  "V6_PLAN_PEUPLEMENT",
  "V6_SOURCES_CSTAR_COMSSA",
]) {
  deleteIfExists(sheetName);
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
  sheet.getRange(`A2:${endCol}2`).format.rowHeightPx = 46;
  return sheet;
}

function setWidths(sheet, widths, rowLimit = 120) {
  widths.forEach((w, index) => {
    const col = colName(index + 1);
    sheet.getRange(`${col}1:${col}${rowLimit}`).format.columnWidthPx = w;
  });
}

function addTable(sheet, range, tableName, style = "TableStyleMedium2") {
  const table = sheet.tables.add(range, true, tableName);
  table.style = style;
  table.showFilterButton = true;
  return table;
}

function addValidation(sheet, colIndex, startRow, endRow, values) {
  const col = colName(colIndex);
  sheet.getRange(`${col}${startRow}:${col}${endRow}`).dataValidation = {
    rule: { type: "list", values },
    prompt: { showPrompt: true, title: "Liste contrôlée", message: "Choisir une valeur autorisée." },
    errorAlert: { showAlert: true, style: "warning", title: "Valeur inattendue", message: "Utiliser une valeur de la liste." },
  };
}

function addStatusFormatting(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.conditionalFormats.add("containsText", { text: "P0", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "P1", format: { fill: palette.lightAmber, font: { color: "#92400E", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "Validé humainement", format: { fill: palette.lightGreen, font: { color: "#166534", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "À vérifier", format: { fill: palette.lightAmber, font: { color: "#92400E" } } });
  range.conditionalFormats.add("containsText", { text: "ALERTE", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "OK", format: { fill: palette.lightGreen, font: { color: "#166534", bold: true } } });
}

const indicators = [];

const domainDefaults = {
  "Gouvernance Climat": { source: "Ville de Dakar", institution: "Cellule climat / Cabinet du Maire", frameworks: "CDP; C40; GCoM; CoM SSA; ICLEI; CDN Sénégal; PCET Dakar; Fonds Vert Climat", priority: "P0" },
  "Données socio-économiques": { source: "ANSD / Ville de Dakar", institution: "ANSD; Direction Planification", frameworks: "CDP; C40; GCoM; CDN Sénégal; Fonds Vert Climat", priority: "P0" },
  "Climat et météorologie": { source: "ANACIM", institution: "ANACIM; Cellule adaptation", frameworks: "CDP; C40; GCoM; CoM SSA; CDN Sénégal; PCET Dakar", priority: "P0" },
  "Risques climatiques": { source: "ANACIM; CSE; ONAS; Ville de Dakar", institution: "Planification; SIG; Environnement", frameworks: "CDP; C40; GCoM; CoM SSA; CDN Sénégal; Fonds Vert Climat", priority: "P0" },
  "Eau": { source: "SEN'EAU; SONES; ONAS; DGPRE", institution: "Direction Eau/Assainissement; ONAS", frameworks: "CDP; C40; GCoM; CoM SSA; CDN Sénégal; Fonds Vert Climat", priority: "P1" },
  "Énergie": { source: "SENELEC; Ville de Dakar", institution: "Direction Énergie; SENELEC", frameworks: "CDP; C40; GCoM; CoM SSA; CDN Sénégal; PCET Dakar; Fonds Vert Climat", priority: "P0" },
  "Inventaire GES": { source: "Cellule MRV; SENELEC; CETUD; SONAGED/UCG", institution: "Référent GES", frameworks: "CDP; C40; GCoM; ICLEI; CDN Sénégal; PCET Dakar; Fonds Vert Climat", priority: "P0" },
  "Mobilité": { source: "CETUD; AFTU; DDD; TER; Ville de Dakar", institution: "Direction Mobilité", frameworks: "CDP; C40; GCoM; CDN Sénégal; PCET Dakar", priority: "P0" },
  "Qualité de l'air": { source: "DEEC / CGQA", institution: "Direction Environnement; DEEC", frameworks: "CDP; C40; ICLEI; PCET Dakar", priority: "P1" },
  "Déchets": { source: "SONAGED/UCG; Ville; prestataires", institution: "Direction Déchets", frameworks: "CDP; C40; GCoM; CoM SSA; CDN Sénégal; PCET Dakar", priority: "P0" },
  "Espaces verts et biodiversité": { source: "Direction Environnement; SIG Ville; CSE; ONG", institution: "Direction Environnement / Espaces verts", frameworks: "CDP; C40; ICLEI; CoM SSA; CDN Sénégal; Fonds Vert Climat", priority: "P1" },
  "Agriculture urbaine": { source: "Ville; ONG; services agriculture urbaine", institution: "Direction Environnement; Affaires sociales", frameworks: "CDP; C40; CoM SSA; CDN Sénégal; Fonds Vert Climat", priority: "P1" },
  "Santé et climat": { source: "Ministère Santé; districts sanitaires", institution: "Direction Santé; Ministère Santé", frameworks: "CDP; C40; GCoM; CDN Sénégal; Fonds Vert Climat", priority: "P1" },
  "Finance climat": { source: "Direction Finances; bailleurs; conventions", institution: "Direction Finances", frameworks: "CDP; C40; CoM SSA; CDN Sénégal; Fonds Vert Climat", priority: "P0" },
  "Portefeuille de projets climatiques": { source: "PMO climat; directions techniques", institution: "Cellule climat; PMO projets", frameworks: "CDP; C40; CoM SSA; CDN Sénégal; PCET Dakar; Fonds Vert Climat", priority: "P0" },
};

const domainPrefix = {
  "Gouvernance Climat": "GOV",
  "Données socio-économiques": "SOC",
  "Climat et météorologie": "MET",
  "Risques climatiques": "RSK",
  "Eau": "EAU",
  "Énergie": "ENE",
  "Inventaire GES": "GES",
  "Mobilité": "MOB",
  "Qualité de l'air": "AIR",
  "Déchets": "DEC",
  "Espaces verts et biodiversité": "NAT",
  "Agriculture urbaine": "AGR",
  "Santé et climat": "SAN",
  "Finance climat": "FIN",
  "Portefeuille de projets climatiques": "PRJ",
};

const counters = new Map();
function nextCode(domain) {
  const next = (counters.get(domain) ?? 0) + 1;
  counters.set(domain, next);
  return `${domainPrefix[domain] ?? "IND"}-${String(next).padStart(3, "0")}`;
}

function addIndicator(domain, subdomain, name, unit, options = {}) {
  const d = domainDefaults[domain] ?? {};
  indicators.push({
    code: options.code ?? nextCode(domain),
    name,
    domain,
    subdomain,
    unit,
    frequency: options.frequency ?? "Annuelle",
    sourcePrimary: options.sourcePrimary ?? d.source ?? "À définir",
    sourceSecondary: options.sourceSecondary ?? "Document source / extraction partenaire",
    institution: options.institution ?? d.institution ?? "À définir",
    focalPoint: options.focalPoint ?? "À désigner",
    evidence: options.evidence ?? "Document justificatif annuel",
    method: options.method ?? "Saisie directe ou agrégation annuelle documentée",
    qualityDefault: options.qualityDefault ?? "À vérifier",
    validationDefault: options.validationDefault ?? "Brouillon",
    frameworks: options.frameworks ?? d.frameworks ?? "CDP; C40; GCoM; CoM SSA; ICLEI; CDN Sénégal; PCET Dakar; Fonds Vert Climat",
    priority: options.priority ?? d.priority ?? "P2",
    sourceDocument: options.sourceDocument ?? "Liste utilisateur; CSTAR; CoMSSA selon disponibilité",
  });
}

function addMany(domain, subdomain, rows, defaults = {}) {
  for (const row of rows) {
    if (Array.isArray(row)) addIndicator(domain, subdomain, row[0], row[1], { ...defaults, ...(row[2] ?? {}) });
    else addIndicator(domain, subdomain, row.name, row.unit, { ...defaults, ...row });
  }
}

addMany("Gouvernance Climat", "Pilotage", [
  ["Existence d'un Plan Climat", "Oui/Non"],
  ["Date d'adoption du Plan Climat", "date"],
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
]);

addMany("Données socio-économiques", "Population et économie", [
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
]);

addMany("Climat et météorologie", "Paramètres climatiques", [
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
]);

const risks = ["Inondations", "Érosion côtière", "Submersion marine", "Canicules", "Sécheresse", "Tempêtes", "Feux de végétation", "Pollution atmosphérique", "Maladies climatiques", "Stress hydrique"];
for (const risk of risks) {
  addIndicator("Risques climatiques", risk, `${risk} - fréquence`, "événements/an");
  addIndicator("Risques climatiques", risk, `${risk} - intensité`, "indice ou unité aléa");
  addIndicator("Risques climatiques", risk, `${risk} - population exposée`, "personnes");
  addIndicator("Risques climatiques", risk, `${risk} - pertes économiques`, "FCFA");
}

addMany("Eau", "Eau potable et assainissement", [
  ["Production d'eau", "m³"],
  ["Consommation d'eau", "m³"],
  ["Pertes réseau", "%"],
  ["Accès à l'eau potable", "% population"],
  ["Qualité de l'eau", "indice"],
  ["Nappes phréatiques", "niveau/qualité"],
  ["Réservoirs", "nombre ou m³"],
  ["Forages", "nombre"],
  ["Eaux usées produites", "m³"],
  ["Eaux usées traitées", "m³"],
]);

addMany("Énergie", "Consommation et production", [
  ["Consommation totale", "kWh"],
  ["Consommation résidentielle", "kWh"],
  ["Consommation commerciale", "kWh"],
  ["Consommation industrielle", "kWh"],
  ["Consommation municipale", "kWh"],
  ["Éclairage public", "kWh"],
  ["Production solaire", "kWh"],
  ["Production éolienne", "kWh"],
  ["Production biomasse", "kWh"],
  ["Accès à l'électricité", "% population"],
  ["Coupures", "heures/an"],
  ["Facteurs d'émission énergie", "kgCO2e/kWh"],
]);

addMany("Inventaire GES", "Secteurs et scopes", [
  ["Énergie - émissions", "tCO2e"],
  ["Combustibles - émissions", "tCO2e"],
  ["Électricité - émissions", "tCO2e"],
  ["Transport - émissions", "tCO2e"],
  ["Transport routier - émissions", "tCO2e"],
  ["Transport ferroviaire - émissions", "tCO2e"],
  ["Transport maritime - émissions", "tCO2e"],
  ["Transport aérien local - émissions", "tCO2e"],
  ["Déchets - émissions", "tCO2e"],
  ["Décharge - émissions", "tCO2e"],
  ["Recyclage - émissions évitées estimées", "tCO2e"],
  ["Compostage - émissions", "tCO2e"],
  ["AFOLU - émissions / absorptions estimées", "tCO2e"],
  ["Agriculture urbaine - émissions / absorptions estimées", "tCO2e"],
  ["Espaces verts - séquestration estimée", "tCO2e"],
  ["Résultats GES totaux", "tCO2e"],
  ["Scope 1", "tCO2e"],
  ["Scope 2", "tCO2e"],
  ["Scope 3", "tCO2e"],
  ["tCO₂e par secteur", "tCO2e"],
  ["tCO₂e par habitant", "tCO2e/hab"],
]);

addMany("Mobilité", "Parc, transport collectif et mobilité active", [
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
  ["Mobilité active", "% déplacements"],
  ["Marche", "% déplacements"],
  ["Vélo", "% déplacements"],
  ["Pistes cyclables", "km"],
  ["Électromobilité", "indice"],
  ["Véhicules électriques", "nombre"],
  ["Bornes de recharge", "nombre"],
]);

addMany("Qualité de l'air", "Polluants", [
  ["PM2.5", "µg/m3"],
  ["PM10", "µg/m3"],
  ["NO₂", "µg/m3"],
  ["SO₂", "µg/m3"],
  ["CO", "mg/m3"],
  ["O₃", "µg/m3"],
  ["Black Carbon", "µg/m3"],
  ["Nombre de stations qualité air", "nombre"],
  ["Dépassements OMS", "jours/an"],
]);

addMany("Déchets", "Flux, valorisation et méthane", [
  ["Déchets collectés", "tonnes"],
  ["Déchets ménagers", "tonnes"],
  ["Déchets commerciaux", "tonnes"],
  ["Déchets industriels", "tonnes"],
  ["Déchets plastiques", "tonnes"],
  ["Déchets organiques", "tonnes"],
  ["Valorisation", "tonnes"],
  ["Déchets recyclés", "tonnes"],
  ["Déchets compostés", "tonnes"],
  ["Déchets réutilisés", "tonnes"],
  ["Déchets valorisés", "tonnes"],
  ["Mise en décharge", "tonnes"],
  ["Quantité enfouie", "tonnes"],
  ["Méthane estimé", "tCH4"],
]);

addMany("Espaces verts et biodiversité", "Espaces verts, arbres et biodiversité", [
  ["Surface espaces verts", "m²"],
  ["Espaces verts par habitant", "m²/hab"],
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
  ["Taux de survie des arbres", "%"],
  ["Séquestration carbone estimée espaces verts", "tCO2e"],
]);

addMany("Agriculture urbaine", "Production et bénéficiaires", [
  ["Surface cultivée", "ha"],
  ["Production agriculture urbaine", "tonnes"],
  ["Nombre de bénéficiaires agriculture urbaine", "personnes"],
  ["Rendement agriculture urbaine", "tonnes/ha"],
  ["Irrigation agriculture urbaine", "m³"],
]);

addMany("Santé et climat", "Maladies et impacts sanitaires", [
  ["Paludisme", "cas/an"],
  ["Dengue", "cas/an"],
  ["Choléra", "cas/an"],
  ["Maladies respiratoires", "cas/an"],
  ["Hospitalisations chaleur", "hospitalisations/an"],
  ["Mortalité chaleur", "décès/an"],
  ["Mortalité pollution", "décès/an"],
]);

addMany("Finance climat", "Budget et financement", [
  ["Budget municipal total", "FCFA"],
  ["Budget climat", "FCFA"],
  ["Dépenses climat", "FCFA"],
  ["Financement UE", "FCFA"],
  ["Financement AFD", "FCFA"],
  ["Financement BAD", "FCFA"],
  ["Financement Banque Mondiale", "FCFA"],
  ["Financement GCF", "FCFA"],
  ["Financement C40", "FCFA"],
  ["Financement ICLEI", "FCFA"],
  ["Portefeuille projets - nombre", "nombre"],
  ["Portefeuille projets - coût", "FCFA"],
  ["Financement obtenu", "FCFA"],
  ["Financement recherché", "FCFA"],
]);

addMany("Portefeuille de projets climatiques", "Fiche projet", [
  ["Nom du projet", "texte"],
  ["Secteur du projet", "texte"],
  ["Description du projet", "texte"],
  ["Localisation du projet", "texte"],
  ["Budget du projet", "FCFA"],
  ["Coût actualisé", "FCFA"],
  ["État d'avancement", "%"],
  ["Maturité du projet", "niveau"],
  ["Bailleur cible", "texte"],
  ["Réduction GES attendue", "tCO2e/an"],
  ["Population bénéficiaire", "personnes"],
  ["Emplois créés", "emplois"],
  ["Risques du projet", "texte"],
  ["Source de financement", "texte"],
  ["Date de début", "date"],
  ["Date de fin", "date"],
]);

const concreteIndicatorCount = indicators.length;
for (let i = concreteIndicatorCount + 1; i <= TARGET_INDICATORS; i += 1) {
  const code = `RES-${String(i).padStart(3, "0")}`;
  indicators.push({
    code,
    name: `Indicateur réservé ${String(i).padStart(3, "0")}`,
    domain: "À définir",
    subdomain: "Réserve bibliothèque",
    unit: "À définir",
    frequency: "Annuelle",
    sourcePrimary: "À définir",
    sourceSecondary: "À définir",
    institution: "À définir",
    focalPoint: "À désigner",
    evidence: "À définir",
    method: "À définir",
    qualityDefault: "À vérifier",
    validationDefault: "Brouillon",
    frameworks: "À mapper",
    priority: "P3",
    sourceDocument: "Réserve capacité 500-800 indicateurs",
  });
}

const libraryHeaders = [
  "Code indicateur",
  "Nom indicateur",
  "Domaine",
  "Sous-domaine",
  "Unité",
  "Fréquence",
  "Source primaire",
  "Source secondaire",
  "Institution",
  "Personne focale",
  "Document justificatif attendu",
  "Méthode de calcul",
  "Niveau de qualité par défaut",
  "Statut validation par défaut",
  "Cadres reporting",
  "Priorité",
  "Source documentaire",
];

function indicatorToLibraryRow(ind) {
  return [
    ind.code,
    ind.name,
    ind.domain,
    ind.subdomain,
    ind.unit,
    ind.frequency,
    ind.sourcePrimary,
    ind.sourceSecondary,
    ind.institution,
    ind.focalPoint,
    ind.evidence,
    ind.method,
    ind.qualityDefault,
    ind.validationDefault,
    ind.frameworks,
    ind.priority,
    ind.sourceDocument,
  ];
}

function buildIndex() {
  const sheet = addSheet(
    "V6_INDEX_BASE",
    "Index V6 - Base annuelle CSTAR / CoMSSA",
    "Navigation et synthèse de la base annuelle reconstruite à partir de la V5, du document CSTAR, du template CoMSSA SEACAP et de la liste d'indicateurs fournie.",
    "J"
  );
  setWidths(sheet, [220, 145, 170, 330, 220, 145, 170, 330, 180, 220], 70);
  sheet.getRange("A4:J4").values = [["Module", "Volume", "Unité", "Usage", "Module", "Volume", "Unité", "Usage", "Statut", "Action"]];
  sheet.getRange("A5:J12").values = [
    ["Indicateurs concrets", concreteIndicatorCount, "indicateurs", "Liste fournie + CSTAR/CoMSSA", "Capacité bibliothèque", TARGET_INDICATORS, "emplacements", "500-800 indicateurs", "OK", "Compléter les emplacements réservés"],
    ["Base annuelle préconfigurée", TARGET_INDICATORS * YEARS.length, "lignes", "800 indicateurs x 20 ans", "Historique", `${YEARS[0]}-${YEARS.at(-1)}`, "années", "10 à 20 ans", "OK", "Saisir résultat et preuves"],
    ["Institutions sources", 60, "emplacements", "Ville + partenaires", "Cadres reporting", FRAMEWORKS.length, "cadres", "CDP/C40/GCoM/CoM SSA/ICLEI/CDN/PCET/GCF", "OK", "Nommer points focaux"],
    ["Source CSTAR", 1, "docx", "Gouvernance, adaptation, GES, eau, finance", "Source CoMSSA", 1, "xlsm", "IRE, EVR, EAE, adaptation, énergie", "OK", "Aligner le canevas annuel"],
    ["Contrôle qualité", "Automatique", "règles", "Source, validation, résultat, document", "Métadonnées", 21, "colonnes", "Obligatoires par donnée", "OK", "Revue mensuelle"],
    ["Reporting", "Mapping", "indicateurs", "CDP/C40/GCoM/CoM SSA/ICLEI/CDN/PCET/GCF", "Plan peuplement", "Priorisé", "P0-P3", "Collecte par domaine", "OK", "Lancer P0"],
    ["Format", "Long", "base", "Une ligne = indicateur + année", "Granularité", "Année/Mois/Trimestre", "champs", "Compatible annuel et infra-annuel", "OK", "Conserver le format long"],
    ["Version", "V6", "xlsx", "Reconstruction base annuelle", "Fichier source", "V5", "xlsx", "Conservé sans écrasement", "OK", "Utiliser ce fichier pour la suite"],
  ];
  sheet.getRange("A4:J4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange("A5:J12").format = { fill: palette.white, wrapText: true };
  addTable(sheet, "A4:J12", "tblV6IndexBase");
}

function buildIndicatorLibrary() {
  const sheet = addSheet(
    "V6_BIBLIOTHEQUE_INDICATEURS",
    "Bibliothèque normalisée des indicateurs",
    "800 emplacements d'indicateurs. Les premiers indicateurs sont préremplis à partir de la liste fournie, du CSTAR et du template CoMSSA; les autres restent réservés.",
    "Q"
  );
  const rows = indicators.map(indicatorToLibraryRow);
  const endRow = 4 + rows.length;
  sheet.getRange(`A4:Q${endRow}`).values = [libraryHeaders, ...rows];
  sheet.getRange("A4:Q4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:Q${endRow}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:Q${endRow}`, "tblBibliothequeIndicateursV6");
  setWidths(sheet, [115, 260, 190, 210, 115, 120, 230, 220, 220, 160, 240, 280, 150, 160, 280, 90, 230], endRow + 5);
  sheet.freezePanes.freezeRows(4);
  addValidation(sheet, 6, 5, endRow, FREQUENCIES);
  addValidation(sheet, 13, 5, endRow, QUALITY_LEVELS);
  addValidation(sheet, 14, 5, endRow, VALIDATION_STATUS);
  addValidation(sheet, 16, 5, endRow, PRIORITIES);
  addStatusFormatting(sheet, `P5:P${endRow}`);
  addStatusFormatting(sheet, `N5:N${endRow}`);
}

const baseHeaders = [
  "Code indicateur",
  "Nom indicateur",
  "Domaine",
  "Sous-domaine",
  "Unité",
  "Fréquence",
  "Année",
  "Mois",
  "Trimestre",
  "Source primaire",
  "Source secondaire",
  "Institution",
  "Personne focale",
  "Document justificatif",
  "Méthode de calcul",
  "Niveau de qualité",
  "Statut validation",
  "Date de collecte",
  "Date de saisie",
  "Responsable Ville",
  "Résultat",
  "Cadres reporting",
  "Priorité",
  "Commentaires",
];

function buildAnnualBase() {
  const sheet = addSheet(
    "V6_BASE_DONNEES_ANNUELLE",
    "Base de données annuelle normalisée",
    "Table long-format: une ligne = un indicateur + une année. Capacité préconfigurée: 800 indicateurs x 20 ans = 16 000 lignes.",
    "X"
  );
  const rows = [];
  for (const ind of indicators) {
    for (const year of YEARS) {
      rows.push([
        ind.code,
        ind.name,
        ind.domain,
        ind.subdomain,
        ind.unit,
        ind.frequency,
        year,
        "Annuel",
        "Annuel",
        ind.sourcePrimary,
        ind.sourceSecondary,
        ind.institution,
        ind.focalPoint,
        "",
        ind.method,
        ind.qualityDefault,
        ind.validationDefault,
        "",
        "",
        "Ville de Dakar",
        "",
        ind.frameworks,
        ind.priority,
        "",
      ]);
    }
  }
  const endRow = 4 + rows.length;
  sheet.getRange(`A4:X${endRow}`).values = [baseHeaders, ...rows];
  sheet.getRange("A4:X4").format = { fill: palette.green, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:X${endRow}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:X${endRow}`, "tblBaseDonneesAnnuelleV6", "TableStyleMedium4");
  setWidths(sheet, [115, 260, 190, 210, 115, 120, 85, 100, 100, 230, 220, 220, 160, 240, 280, 150, 160, 120, 120, 170, 135, 280, 90, 260], endRow + 5);
  sheet.freezePanes.freezeRows(4);
  addValidation(sheet, 6, 5, endRow, FREQUENCIES);
  addValidation(sheet, 16, 5, endRow, QUALITY_LEVELS);
  addValidation(sheet, 17, 5, endRow, VALIDATION_STATUS);
  addValidation(sheet, 23, 5, endRow, PRIORITIES);
  sheet.getRange(`G5:G${endRow}`).setNumberFormat("0");
  sheet.getRange(`R5:S${endRow}`).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(`U5:U${endRow}`).setNumberFormat("#,##0.00");
  addStatusFormatting(sheet, `Q5:Q${endRow}`);
  addStatusFormatting(sheet, `W5:W${endRow}`);
}

function buildInstitutions() {
  const sheet = addSheet(
    "V6_INSTITUTIONS_SOURCES",
    "Institutions sources et points focaux",
    "Répertoire de 60 emplacements pour gérer 50 à 100 institutions sources et leurs responsabilités de transmission.",
    "J"
  );
  const base = [
    ["INS-001", "Ville de Dakar", "Municipal", "Gouvernance; projets; budget; espaces verts", "Cellule climat", "À désigner", "Mensuelle", "P0", "Protocole interne", "Actif"],
    ["INS-002", "ANSD", "National", "Population; socio-économie", "Statistiques", "À désigner", "Annuelle", "P0", "Convention données", "À confirmer"],
    ["INS-003", "ANACIM", "National", "Météo; climat; risques", "Climatologie", "À désigner", "Mensuelle/Annuelle", "P0", "Convention données", "À confirmer"],
    ["INS-004", "SENELEC", "Opérateur", "Électricité; facteurs émission", "Données énergie", "À désigner", "Mensuelle", "P0", "Convention données", "À confirmer"],
    ["INS-005", "SONAGED / UCG", "Opérateur", "Déchets; collecte; valorisation", "Déchets", "À désigner", "Mensuelle", "P0", "Convention données", "À confirmer"],
    ["INS-006", "CETUD", "Transport", "Mobilité; transport collectif", "Mobilité", "À désigner", "Trimestrielle", "P0", "Convention données", "À confirmer"],
    ["INS-007", "DEEC / CGQA", "National", "Qualité de l'air; pollution", "Qualité air", "À désigner", "Mensuelle", "P1", "Convention données", "À confirmer"],
    ["INS-008", "ONAS", "National", "Eaux usées; inondations; assainissement", "Assainissement", "À désigner", "Trimestrielle", "P1", "Convention données", "À confirmer"],
    ["INS-009", "SEN'EAU / SONES", "Opérateur", "Eau potable; production; pertes", "Eau", "À désigner", "Trimestrielle", "P1", "Convention données", "À confirmer"],
    ["INS-010", "CSE", "National", "SIG; risques; occupation sol; espaces verts", "SIG", "À désigner", "Annuelle", "P1", "Convention données", "À confirmer"],
    ["INS-011", "Ministère Santé", "National", "Santé et climat", "Données sanitaires", "À désigner", "Annuelle", "P1", "Convention données", "À confirmer"],
    ["INS-012", "Direction Finances Ville", "Municipal", "Budget; dépenses; finance climat", "Finance", "À désigner", "Mensuelle", "P0", "Extraction budget", "Actif"],
    ["INS-013", "Direction Mobilité Ville", "Municipal", "Mobilité; projets", "Mobilité", "À désigner", "Mensuelle", "P0", "Extraction interne", "Actif"],
    ["INS-014", "Direction Environnement Ville", "Municipal", "Air; nature; déchets; adaptation", "Environnement", "À désigner", "Mensuelle", "P0", "Extraction interne", "Actif"],
    ["INS-015", "Direction Énergie Ville", "Municipal", "Énergie municipale; éclairage", "Énergie", "À désigner", "Mensuelle", "P0", "Extraction interne", "Actif"],
  ];
  while (base.length < 60) {
    const n = base.length + 1;
    base.push([`INS-${String(n).padStart(3, "0")}`, `Institution source réservée ${n}`, "À définir", "À définir", "À définir", "À désigner", "À définir", "P3", "À établir", "Réserve"]);
  }
  sheet.getRange("A4:J64").values = [["Code institution", "Institution", "Type", "Domaines", "Service", "Personne focale", "Fréquence", "Priorité", "Accord / preuve", "Statut"], ...base];
  sheet.getRange("A4:J4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange("A5:J64").format = { fill: palette.white, wrapText: true };
  addTable(sheet, "A4:J64", "tblInstitutionsSourcesV6");
  setWidths(sheet, [115, 230, 130, 260, 170, 160, 130, 90, 190, 120], 80);
  addValidation(sheet, 8, 5, 64, PRIORITIES);
  addStatusFormatting(sheet, "H5:H64");
}

function buildMapping() {
  const sheet = addSheet(
    "V6_MAPPING_REPORTING",
    "Mapping indicateurs vers cadres de reporting",
    "Correspondance opérationnelle des domaines avec CDP, C40, GCoM, CoM SSA, ICLEI, CDN Sénégal, PCET Dakar et Fonds Vert Climat.",
    "L"
  );
  const domains = Object.keys(domainDefaults);
  const rows = domains.map(domain => {
    const fw = domainDefaults[domain].frameworks;
    return [
      domain,
      fw.includes("CDP") ? "Oui" : "Partiel",
      fw.includes("C40") ? "Oui" : "Partiel",
      fw.includes("GCoM") ? "Oui" : "Partiel",
      fw.includes("CoM SSA") ? "Oui" : "Partiel",
      fw.includes("ICLEI") ? "Oui" : "Partiel",
      fw.includes("CDN") ? "Oui" : "Partiel",
      fw.includes("PCET") ? "Oui" : "Partiel",
      fw.includes("Fonds Vert") ? "Oui" : "Partiel",
      `=COUNTIF(V6_BIBLIOTHEQUE_INDICATEURS!$C$5:$C$804,A${5 + domains.indexOf(domain)})`,
      domainDefaults[domain].priority,
      "Vérifier champs exacts lors du cycle de reporting annuel",
    ];
  });
  sheet.getRange(`A4:L${4 + rows.length}`).values = [["Domaine", "CDP", "C40", "GCoM", "CoM SSA", "ICLEI", "CDN Sénégal", "PCET Dakar", "Fonds Vert Climat", "Nb indicateurs", "Priorité", "Note"], ...rows];
  sheet.getRange("A4:L4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:L${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:L${4 + rows.length}`, "tblMappingReportingV6");
  setWidths(sheet, [220, 80, 80, 80, 90, 80, 110, 100, 130, 110, 90, 300], 80);
  addStatusFormatting(sheet, `K5:K${4 + rows.length}`);
}

function buildQualityControl() {
  const sheet = addSheet(
    "V6_CONTROLE_QUALITE",
    "Contrôle qualité de la base annuelle",
    "Contrôles automatiques sur bibliothèque, base annuelle, métadonnées, validation, sources et capacité.",
    "L"
  );
  setWidths(sheet, [260, 130, 150, 330, 260, 130, 150, 330, 160, 130, 160, 260], 90);
  const rows = [
    ["Indicateurs concrets", concreteIndicatorCount, "nombre", "Indicateurs issus de la liste utilisateur et des sources", "Emplacements réservés", TARGET_INDICATORS - concreteIndicatorCount, "nombre", "Capacité restante", "OK", "Compléter progressivement", "P2", "Bibliothèque"],
    ["Lignes annuelles", TARGET_INDICATORS * YEARS.length, "lignes", "800 indicateurs x 20 ans", "Résultats renseignés", '=COUNTIF(V6_BASE_DONNEES_ANNUELLE!$U$5:$U$16004,"<>")', "lignes", "Données saisies", '=IF(F6=0,"À collecter","OK")', "Démarrer peuplement P0", "P0", "Base annuelle"],
    ["Résultats sans document", '=COUNTIFS(V6_BASE_DONNEES_ANNUELLE!$U$5:$U$16004,"<>",V6_BASE_DONNEES_ANNUELLE!$N$5:$N$16004,"")', "alertes", "Document justificatif obligatoire", "Résultats non validés", '=COUNTIFS(V6_BASE_DONNEES_ANNUELLE!$U$5:$U$16004,"<>",V6_BASE_DONNEES_ANNUELLE!$Q$5:$Q$16004,"<>Validé humainement")', "alertes", "Validation obligatoire", '=IF(OR(B7>0,F7>0),"ALERTE","OK")', "Joindre preuves et valider", "P0", "MRV"],
    ["Sources primaires manquantes", '=COUNTIF(V6_BIBLIOTHEQUE_INDICATEURS!$G$5:$G$804,"À définir")', "indicateurs", "Source primaire doit être nommée", "Institutions à définir", '=COUNTIF(V6_BIBLIOTHEQUE_INDICATEURS!$I$5:$I$804,"À définir")', "indicateurs", "Institution obligatoire", '=IF(OR(B8>0,F8>0),"À corriger","OK")', "Compléter bibliothèque", "P1", "Dictionnaire"],
    ["Indicateurs P0", '=COUNTIF(V6_BIBLIOTHEQUE_INDICATEURS!$P$5:$P$804,"P0")', "nombre", "Critiques à peupler en premier", "Indicateurs P1", '=COUNTIF(V6_BIBLIOTHEQUE_INDICATEURS!$P$5:$P$804,"P1")', "nombre", "Haute priorité", "OK", "Planifier collecte", "P0", "Planning"],
    ["Couverture CDP", '=COUNTIF(V6_BIBLIOTHEQUE_INDICATEURS!$O$5:$O$804,"*CDP*")', "indicateurs", "Préparation disclosure", "Couverture CoM SSA", '=COUNTIF(V6_BIBLIOTHEQUE_INDICATEURS!$O$5:$O$804,"*CoM SSA*")', "indicateurs", "SEACAP", "OK", "Aligner canevas", "P0", "Reporting"],
    ["Sources CSTAR", '=IF(COUNTIF(V6_SOURCES_CSTAR_COMSSA!$A$5:$A$20,"*CSTAR*")>0,"OK","Absent")', "statut", "Document CSTAR indexé", "Source CoMSSA", '=IF(COUNTIF(V6_SOURCES_CSTAR_COMSSA!$A$5:$A$20,"*CoMSSA*")>0,"OK","Absent")', "statut", "Template CoMSSA indexé", "OK", "Conserver preuves", "P1", "Sources"],
  ];
  sheet.getRange("A4:L4").values = [["Contrôle A", "Valeur A", "Unité A", "Lecture A", "Contrôle B", "Valeur B", "Unité B", "Lecture B", "Statut", "Action", "Priorité", "Module"]];
  sheet.getRange(`A5:L${4 + rows.length}`).values = rows;
  for (let r = 5; r < 5 + rows.length; r += 1) {
    for (const c of ["B", "F", "I"]) {
      const value = sheet.getRange(`${c}${r}`).values?.[0]?.[0];
      if (typeof value === "string" && value.startsWith("=")) {
        sheet.getRange(`${c}${r}`).formulas = [[value]];
      }
    }
  }
  sheet.getRange("A4:L4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:L${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:L${4 + rows.length}`, "tblControleQualiteV6");
  addStatusFormatting(sheet, `I5:I${4 + rows.length}`);
  addStatusFormatting(sheet, `K5:K${4 + rows.length}`);

  const domainRows = Object.keys(domainDefaults).map((domain, i) => [domain, `=COUNTIF(V6_BIBLIOTHEQUE_INDICATEURS!$C$5:$C$804,N${5 + i})`]);
  sheet.getRange(`N4:O${4 + domainRows.length}`).values = [["Domaine", "Indicateurs"], ...domainRows];
  sheet.getRange(`N4:O4`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  sheet.getRange(`N5:O${4 + domainRows.length}`).format = { fill: palette.white, wrapText: true };
  const chart = sheet.charts.add("bar", sheet.getRange(`N4:O${4 + domainRows.length}`));
  chart.title = "Indicateurs par domaine";
  chart.hasLegend = false;
  chart.xAxis = { axisType: "textAxis" };
  chart.yAxis = { numberFormatCode: "#,##0" };
  chart.setPosition("Q4", "X23");
}

function buildDictionary() {
  const sheet = addSheet(
    "V6_DICTIONNAIRE_DONNEES",
    "Dictionnaire des champs obligatoires",
    "Définition des métadonnées obligatoires pour chaque donnée annuelle.",
    "F"
  );
  const rows = [
    ["Code indicateur", "Identifiant stable de l'indicateur", "Texte", "Oui", "Bibliothèque", "Ne jamais modifier après usage historique"],
    ["Nom indicateur", "Libellé lisible", "Texte", "Oui", "Bibliothèque", "Peut être précisé mais pas dénaturé"],
    ["Domaine", "Catégorie principale", "Liste", "Oui", "Bibliothèque", "15 domaines principaux"],
    ["Sous-domaine", "Catégorie détaillée", "Texte/liste", "Oui", "Bibliothèque", "Risque, secteur, flux, etc."],
    ["Unité", "Unité de mesure", "Liste/texte", "Oui", "Bibliothèque", "Utiliser unités SI ou unités reporting"],
    ["Fréquence", "Fréquence de collecte", "Liste", "Oui", "Bibliothèque", "Mensuelle, trimestrielle, annuelle"],
    ["Année", "Année de référence", "Nombre", "Oui", "Base annuelle", "2020-2039 préconfiguré"],
    ["Mois", "Mois si disponible", "Liste", "Oui", "Base annuelle", "Annuel par défaut"],
    ["Trimestre", "Trimestre si disponible", "Liste", "Oui", "Base annuelle", "Annuel par défaut"],
    ["Source primaire", "Source responsable de la donnée", "Texte", "Oui", "Bibliothèque/base", "Institution ou système source"],
    ["Source secondaire", "Source complémentaire", "Texte", "Non", "Bibliothèque/base", "Rapport, extraction, partenaire"],
    ["Institution", "Institution source", "Texte/liste", "Oui", "Institutions", "Doit être reliée à un point focal"],
    ["Personne focale", "Contact opérationnel", "Texte", "Oui", "Institutions", "À désigner si inconnu"],
    ["Document justificatif", "Preuve ou lien source", "Texte/lien", "Oui si résultat", "Base annuelle", "Obligatoire avant validation"],
    ["Méthode de calcul", "Formule ou méthode", "Texte", "Oui", "Bibliothèque", "Documenter ratios et conversions"],
    ["Niveau de qualité", "Fiabilité", "Liste", "Oui", "Base annuelle", "Faible à très élevée"],
    ["Statut validation", "Cycle de validation", "Liste", "Oui", "Base annuelle", "Validé humainement pour reporting"],
    ["Date de collecte", "Date réception donnée", "Date", "Oui si résultat", "Base annuelle", "Audit trail"],
    ["Date de saisie", "Date saisie dans le Hub", "Date", "Oui si résultat", "Base annuelle", "Audit trail"],
    ["Responsable Ville", "Responsable interne", "Texte", "Oui", "Base annuelle", "Nom direction/personne"],
    ["Résultat", "Valeur de l'indicateur", "Nombre/texte", "Oui", "Base annuelle", "Valeur à reporter"],
  ];
  sheet.getRange(`A4:F${4 + rows.length}`).values = [["Champ", "Définition", "Type", "Obligatoire", "Feuille", "Règle"], ...rows];
  sheet.getRange("A4:F4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:F${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:F${4 + rows.length}`, "tblDictionnaireDonneesV6");
  setWidths(sheet, [190, 330, 120, 110, 150, 330], 80);
}

function buildPlan() {
  const sheet = addSheet(
    "V6_PLAN_PEUPLEMENT",
    "Plan de peuplement annuel de la base",
    "Priorisation par domaine pour remplir la base annuelle avec sources, responsables, fréquence, difficulté et cadres critiques.",
    "J"
  );
  const rows = Object.keys(domainDefaults).map((domain) => {
    const d = domainDefaults[domain];
    return [
      domain,
      d.source,
      d.institution,
      domain === "Climat et météorologie" || domain === "Énergie" || domain === "Déchets" ? "Mensuelle/Annuelle" : "Annuelle",
      ["Inventaire GES", "Mobilité", "Risques climatiques"].includes(domain) ? "Très élevée" : ["Eau", "Qualité de l'air", "Santé et climat"].includes(domain) ? "Élevée" : "Moyenne",
      d.priority,
      d.frameworks,
      `=COUNTIF(V6_BIBLIOTHEQUE_INDICATEURS!$C$5:$C$804,A${5 + Object.keys(domainDefaults).indexOf(domain)})`,
      domain === "Inventaire GES" ? "Valider périmètre GPC/CoMSSA et facteurs d'émission" : domain === "Finance climat" ? "Relier budgets, dépenses, financements et projets" : "Nommer point focal et sécuriser document justificatif",
      "À lancer",
    ];
  });
  sheet.getRange(`A4:J${4 + rows.length}`).values = [["Domaine", "Source", "Responsable", "Fréquence", "Difficulté", "Priorité", "Cadres critiques", "Nb indicateurs", "Première action", "Statut"], ...rows];
  sheet.getRange("A4:J4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:J${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:J${4 + rows.length}`, "tblPlanPeuplementV6");
  setWidths(sheet, [210, 270, 230, 140, 130, 90, 330, 115, 360, 120], 80);
  addStatusFormatting(sheet, `F5:F${4 + rows.length}`);
}

function buildSources() {
  const sheet = addSheet(
    "V6_SOURCES_CSTAR_COMSSA",
    "Sources CSTAR et CoMSSA utilisées",
    "Extraction synthétique des deux fichiers fournis: document CSTAR et template CoMSSA SEACAP.",
    "H"
  );
  const rows = [];
  for (const rec of sourceExtraction) {
    rows.push([
      rec.file_name ?? rec.path,
      rec.extension,
      rec.size_kb,
      (rec.sheet_names ?? []).join("; "),
      Object.entries(rec.keywords ?? {}).map(([k, v]) => `${k}:${v}`).join("; "),
      rec.excerpt ?? "",
      rec.path,
      rec.exists ? "OK" : "Absent",
    ]);
  }
  sheet.getRange(`A4:H${4 + rows.length}`).values = [["Fichier", "Type", "Taille KB", "Feuilles détectées", "Mots-clés détectés", "Extrait", "Chemin", "Statut"], ...rows];
  sheet.getRange("A4:H4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:H${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:H${4 + rows.length}`, "tblSourcesCstarComssaV6");
  setWidths(sheet, [330, 80, 90, 400, 360, 520, 520, 90], 20);
}

function updateAccueil() {
  const sheet = workbook.worksheets.getItemOrNullObject?.("00_ACCUEIL");
  if (!sheet || sheet.isNullObject) return;
  sheet.getRange("A38:J38").merge();
  sheet.getRange("A38").values = [["Menu V6 - Base annuelle CSTAR / CoMSSA"]];
  sheet.getRange("A38:J38").format = { fill: palette.green, font: { bold: true, color: palette.white, size: 13 } };
  const links = [
    ["Index V6", "#'V6_INDEX_BASE'!A1"],
    ["Bibliothèque indicateurs", "#'V6_BIBLIOTHEQUE_INDICATEURS'!A1"],
    ["Base annuelle", "#'V6_BASE_DONNEES_ANNUELLE'!A1"],
    ["Institutions sources", "#'V6_INSTITUTIONS_SOURCES'!A1"],
    ["Mapping reporting", "#'V6_MAPPING_REPORTING'!A1"],
    ["Contrôle qualité", "#'V6_CONTROLE_QUALITE'!A1"],
    ["Dictionnaire données", "#'V6_DICTIONNAIRE_DONNEES'!A1"],
    ["Sources CSTAR/CoMSSA", "#'V6_SOURCES_CSTAR_COMSSA'!A1"],
  ];
  const rows = [];
  for (let i = 0; i < links.length; i += 2) rows.push([links[i][0], links[i][1], links[i + 1][0], links[i + 1][1]]);
  sheet.getRange(`A40:D${39 + rows.length}`).values = rows;
  for (let r = 40; r < 40 + rows.length; r += 1) {
    sheet.getRange(`A${r}`).formulas = [[`=HYPERLINK(B${r},A${r})`]];
    sheet.getRange(`C${r}`).formulas = [[`=HYPERLINK(D${r},C${r})`]];
  }
  sheet.getRange(`A40:D${39 + rows.length}`).format = { fill: palette.lightGreen, font: { color: palette.navy, bold: true }, wrapText: true };
}

buildIndex();
buildIndicatorLibrary();
buildAnnualBase();
buildInstitutions();
buildMapping();
buildQualityControl();
buildDictionary();
buildPlan();
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

const qc = await workbook.inspect({
  kind: "table",
  range: "V6_CONTROLE_QUALITE!A4:L12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 12,
});
console.log("QC_CHECK");
console.log(qc.ndjson);

for (const [sheetName, range] of [
  ["V6_INDEX_BASE", "A1:J14"],
  ["V6_BIBLIOTHEQUE_INDICATEURS", "A1:Q18"],
  ["V6_BASE_DONNEES_ANNUELLE", "A1:X18"],
  ["V6_CONTROLE_QUALITE", "A1:X26"],
  ["V6_SOURCES_CSTAR_COMSSA", "A1:H8"],
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
