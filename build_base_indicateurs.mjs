import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

/* =========================================================================
   DAKAR — BASE D'INDICATEURS CLIMAT (organisée PAR BLOC et PAR ANNÉE)
   Reconstruite à partir de :
     • Questionnaire CDP 2024 — Ville de Dakar (gouvernance, risques, inventaire,
       données sectorielles, cibles, planification, actions, finance)
     • Modèle de reporting CoM SSA / SEACAP — Ville de Dakar (IRE/BEI, EVR/RVA, EAE)
     • Taxonomie des 15 blocs de données fournie par la Ville
   Sortie : outputs/DAKAR_CLIMATE_BASE_INDICATEURS.xlsx
   ========================================================================= */

const outputDir = path.resolve("outputs");
const outputFile = path.join(outputDir, "DAKAR_CLIMATE_BASE_INDICATEURS.xlsx");
const wb = await Workbook.create();

const palette = {
  navy: "#0B1F33", teal: "#0F766E", green: "#166534", leaf: "#2E7D32",
  amber: "#B45309", red: "#B91C1C", blue: "#1D4ED8", violet: "#6D28D9",
  cyan: "#0E7490", brown: "#7C4A03", slate: "#334155",
  lightTeal: "#E6F4F1", lightGreen: "#EAF7EA", lightAmber: "#FEF3C7",
  lightRed: "#FEE2E2", lightBlue: "#E8F0FE", lightViolet: "#F1E9FB",
  white: "#FFFFFF", grey: "#F8FAFC",
};

/* ----------------------------- années & colonnes ----------------------- */
const Y0 = 2015, Y1 = 2035;
const YEARS = [];
for (let y = Y0; y <= Y1; y += 1) YEARS.push(y);
const META = 4;                 // A..D : Code, Indicateur, Sous-bloc, Unité
const YEAR_START = META + 1;    // E
const SRC_COL = YEAR_START + YEARS.length;   // après les années
const STA_COL = SRC_COL + 1;
const N_COLS = STA_COL;         // total colonnes

function colName(n) {
  let s = "", x = n;
  while (x > 0) { const r = (x - 1) % 26; s = String.fromCharCode(65 + r) + s; x = Math.floor((x - 1) / 26); }
  return s;
}
const END = colName(N_COLS);
const yearCol = (y) => colName(YEAR_START + (y - Y0));

/* ----------------------------- listes ---------------------------------- */
const STATUTS = ["Brouillon", "À vérifier", "Validé", "Estimé", "Projection", "N/D"];

/* ----------------------------- helpers ---------------------------------- */
function title(ws, t, sub, endCol = END, accent = palette.navy) {
  ws.showGridLines = false;
  ws.getRange(`A1:${endCol}1`).merge();
  ws.getRange("A1").values = [[t]];
  ws.getRange(`A1:${endCol}1`).format = { fill: accent, font: { bold: true, color: palette.white, size: 15 } };
  ws.getRange(`A1:${endCol}1`).format.rowHeightPx = 32;
  ws.getRange(`A2:${endCol}2`).merge();
  ws.getRange("A2").values = [[sub]];
  ws.getRange(`A2:${endCol}2`).format = { fill: palette.lightGreen, font: { italic: true, color: palette.slate }, wrapText: true };
  ws.getRange(`A2:${endCol}2`).format.rowHeightPx = 40;
}
function freeze(ws, rows, cols) {
  try { ws.freezePanes.freezeRows(rows); } catch (e) { /* noop */ }
  try { ws.freezePanes.freezeColumns(cols); } catch (e) { /* noop */ }
}

/* =========================================================================
   DÉFINITION DES 14 BLOCS « indicateur × année »
   ========================================================================= */
const RISKS = ["Inondations", "Érosion côtière", "Submersion marine", "Canicules", "Sécheresse",
  "Tempêtes", "Feux de végétation", "Pollution atmosphérique", "Maladies climatiques", "Stress hydrique"];
const RISK_METRICS = [
  { m: "Fréquence", u: "nb/an" }, { m: "Intensité", u: "échelle 1-5" },
  { m: "Population exposée", u: "hab" }, { m: "Pertes économiques", u: "FCFA" },
];
const riskIndicators = RISKS.flatMap((r) => RISK_METRICS.map((x) => ({ n: `${r} — ${x.m}`, s: r, u: x.u })));

const blocks = [
  {
    prefix: "GOUV", sheet: "01_Gouvernance", accent: palette.navy,
    t: "Bloc 1 — Gouvernance climat", sub: "Pilotage, moyens humains et financiers, instances et dispositifs de suivi (MRV, budget vert, inventaire).",
    inds: [
      { n: "Existence d'un Plan Climat", u: "Oui/Non" }, { n: "Date d'adoption du Plan Climat", u: "date" },
      { n: "Budget climat annuel", u: "FCFA" }, { n: "Dépenses climat", u: "FCFA" },
      { n: "Personnel climat", u: "ETP" }, { n: "Nombre d'agents formés", u: "nb" },
      { n: "Nombre de réunions COPIL", u: "nb" }, { n: "Nombre de consultations publiques", u: "nb" },
      { n: "Nombre de partenaires", u: "nb" }, { n: "Existence d'un système MRV", u: "Oui/Non" },
      { n: "Existence d'un budget vert", u: "Oui/Non" }, { n: "Existence d'un inventaire GES", u: "Oui/Non" },
    ],
  },
  {
    prefix: "SOCIO", sheet: "02_Socio_Economie", accent: palette.slate,
    t: "Bloc 2 — Données socio-économiques", sub: "Population, structure démographique, ménages, économie et populations vulnérables / exposées.",
    inds: [
      { n: "Population totale", u: "hab" }, { n: "Population — part de femmes", u: "%" },
      { n: "Population jeune (<15 ans)", u: "%" }, { n: "Population âgée (>65 ans)", u: "%" },
      { n: "Densité", u: "hab/km²" }, { n: "Nombre de ménages", u: "nb" },
      { n: "PIB local estimé", u: "FCFA" }, { n: "Taux de pauvreté", u: "%" },
      { n: "Taux de chômage", u: "%" }, { n: "Population vulnérable", u: "hab" },
      { n: "Population exposée aux risques", u: "hab" },
    ],
  },
  {
    prefix: "CLIM", sheet: "03_Climat_Meteo", accent: palette.cyan,
    t: "Bloc 3 — Climat et météorologie", sub: "Paramètres climatiques observés : températures, précipitations, vent, humidité, rayonnement, niveau marin.",
    inds: [
      { n: "Température moyenne", u: "°C" }, { n: "Température minimale", u: "°C" },
      { n: "Température maximale", u: "°C" }, { n: "Température ressentie", u: "°C" },
      { n: "Nombre de jours chauds", u: "j" }, { n: "Nombre de vagues de chaleur", u: "nb" },
      { n: "Pluviométrie", u: "mm" }, { n: "Nombre de jours de pluie", u: "j" },
      { n: "Intensité des pluies", u: "mm/j" }, { n: "Humidité", u: "%" },
      { n: "Vitesse du vent", u: "m/s" }, { n: "Direction du vent", u: "°" },
      { n: "Rayonnement solaire", u: "kWh/m²" }, { n: "Évapotranspiration", u: "mm" },
      { n: "Niveau marin", u: "cm" },
    ],
  },
  {
    prefix: "RISQ", sheet: "04_Risques_Climatiques", accent: palette.red,
    t: "Bloc 4 — Risques climatiques", sub: "Pour chaque aléa : fréquence, intensité, population exposée et pertes économiques.",
    inds: riskIndicators,
  },
  {
    prefix: "EAU", sheet: "05_Eau", accent: palette.blue,
    t: "Bloc 5 — Eau", sub: "Production, consommation, pertes, accès, qualité, ressources souterraines et assainissement.",
    inds: [
      { n: "Production d'eau", u: "m³" }, { n: "Consommation d'eau", u: "m³" },
      { n: "Pertes réseau", u: "%" }, { n: "Accès à l'eau potable", u: "%" },
      { n: "Qualité de l'eau (indice)", u: "indice" }, { n: "Niveau des nappes phréatiques", u: "m" },
      { n: "Capacité des réservoirs", u: "m³" }, { n: "Nombre de forages", u: "nb" },
      { n: "Eaux usées produites", u: "m³" }, { n: "Eaux usées traitées", u: "m³" },
    ],
  },
  {
    prefix: "ENER", sheet: "06_Energie", accent: palette.amber,
    t: "Bloc 6 — Énergie", sub: "Consommations par usage, production renouvelable, accès, fiabilité et facteur d'émission de l'électricité.",
    inds: [
      { n: "Consommation totale", u: "MWh" }, { n: "Consommation résidentielle", u: "MWh" },
      { n: "Consommation commerciale", u: "MWh" }, { n: "Consommation industrielle", u: "MWh" },
      { n: "Consommation municipale", u: "MWh" }, { n: "Éclairage public", u: "MWh" },
      { n: "Production solaire", u: "MWh" }, { n: "Production éolienne", u: "MWh" },
      { n: "Production biomasse", u: "MWh" }, { n: "Accès à l'électricité", u: "%" },
      { n: "Coupures", u: "nb/an" }, { n: "Facteur d'émission électricité", u: "tCO2e/MWh" },
    ],
  },
  {
    prefix: "GES", sheet: "07_Inventaire_GES", accent: palette.green,
    t: "Bloc 7 — Inventaire GES", sub: "Émissions par secteur GIEC/GPC (énergie, transport, déchets, AFOLU) et résultats par scope et par habitant.",
    inds: [
      { n: "Énergie — Combustibles", s: "Énergie", u: "tCO2e" }, { n: "Énergie — Électricité", s: "Énergie", u: "tCO2e" },
      { n: "Transport — Routier", s: "Transport", u: "tCO2e" }, { n: "Transport — Ferroviaire", s: "Transport", u: "tCO2e" },
      { n: "Transport — Maritime", s: "Transport", u: "tCO2e" }, { n: "Transport — Aérien local", s: "Transport", u: "tCO2e" },
      { n: "Déchets — Décharge", s: "Déchets", u: "tCO2e" }, { n: "Déchets — Recyclage", s: "Déchets", u: "tCO2e" },
      { n: "Déchets — Compostage", s: "Déchets", u: "tCO2e" },
      { n: "AFOLU — Agriculture urbaine", s: "AFOLU", u: "tCO2e" }, { n: "AFOLU — Espaces verts", s: "AFOLU", u: "tCO2e" },
      { n: "Émissions Scope 1", s: "Résultats", u: "tCO2e" }, { n: "Émissions Scope 2", s: "Résultats", u: "tCO2e" },
      { n: "Émissions Scope 3", s: "Résultats", u: "tCO2e" }, { n: "Émissions totales (tous secteurs)", s: "Résultats", u: "tCO2e" },
      { n: "Émissions par habitant", s: "Résultats", u: "tCO2e/hab" },
    ],
  },
  {
    prefix: "MOB", sheet: "08_Mobilite", accent: palette.violet,
    t: "Bloc 8 — Mobilité", sub: "Parc de véhicules, transport collectif (BRT/TER/DDD), mobilité active et électromobilité.",
    inds: [
      { n: "Nombre de véhicules", s: "Parc", u: "nb" }, { n: "Véhicules particuliers", s: "Parc", u: "nb" },
      { n: "Taxis", s: "Parc", u: "nb" }, { n: "Bus", s: "Parc", u: "nb" },
      { n: "Cars rapides", s: "Parc", u: "nb" }, { n: "Ndiaga Ndiaye", s: "Parc", u: "nb" },
      { n: "Motos", s: "Parc", u: "nb" }, { n: "Camions", s: "Parc", u: "nb" },
      { n: "Transport collectif (passagers)", s: "Transport collectif", u: "passagers/an" },
      { n: "Passagers BRT", s: "Transport collectif", u: "passagers/an" },
      { n: "Passagers TER", s: "Transport collectif", u: "passagers/an" },
      { n: "Passagers DDD", s: "Transport collectif", u: "passagers/an" },
      { n: "Part de la marche", s: "Mobilité active", u: "%" }, { n: "Part du vélo", s: "Mobilité active", u: "%" },
      { n: "Pistes cyclables", s: "Mobilité active", u: "km" },
      { n: "Véhicules électriques", s: "Électromobilité", u: "nb" }, { n: "Bornes de recharge", s: "Électromobilité", u: "nb" },
    ],
  },
  {
    prefix: "AIR", sheet: "09_Qualite_Air", accent: palette.brown,
    t: "Bloc 9 — Qualité de l'air", sub: "Polluants réglementés, black carbon, réseau de mesure et dépassements des seuils OMS.",
    inds: [
      { n: "PM2.5", u: "µg/m³" }, { n: "PM10", u: "µg/m³" }, { n: "NO₂", u: "µg/m³" },
      { n: "SO₂", u: "µg/m³" }, { n: "CO", u: "mg/m³" }, { n: "O₃", u: "µg/m³" },
      { n: "Black Carbon", u: "µg/m³" }, { n: "Nombre de stations", u: "nb" }, { n: "Dépassements OMS", u: "j/an" },
    ],
  },
  {
    prefix: "DECH", sheet: "10_Dechets", accent: palette.teal,
    t: "Bloc 10 — Déchets", sub: "Quantités collectées par flux, valorisation (recyclage/compostage/réemploi) et mise en décharge.",
    inds: [
      { n: "Déchets collectés", s: "Collecte", u: "t" }, { n: "Déchets ménagers", s: "Collecte", u: "t" },
      { n: "Déchets commerciaux", s: "Collecte", u: "t" }, { n: "Déchets industriels", s: "Collecte", u: "t" },
      { n: "Déchets plastiques", s: "Collecte", u: "t" }, { n: "Déchets organiques", s: "Collecte", u: "t" },
      { n: "Déchets recyclés", s: "Valorisation", u: "t" }, { n: "Déchets compostés", s: "Valorisation", u: "t" },
      { n: "Déchets réutilisés", s: "Valorisation", u: "t" }, { n: "Déchets valorisés (total)", s: "Valorisation", u: "t" },
      { n: "Quantité enfouie", s: "Décharge", u: "t" }, { n: "Méthane estimé", s: "Décharge", u: "tCO2e" },
    ],
  },
  {
    prefix: "VERT", sheet: "11_Espaces_Verts_Biodiversite", accent: palette.leaf,
    t: "Bloc 11 — Espaces verts et biodiversité", sub: "Surfaces végétalisées, ratio par habitant, typologies d'espaces, écosystèmes et patrimoine arboré.",
    inds: [
      { n: "Surface d'espaces verts", u: "ha" }, { n: "Espaces verts par habitant", u: "m²/hab" },
      { n: "Parcs", u: "nb" }, { n: "Jardins publics", u: "nb" },
      { n: "Jardins communautaires", u: "nb" }, { n: "Micro-jardins", u: "nb" },
      { n: "Forêts urbaines", u: "ha" }, { n: "Corridors verts", u: "km" },
      { n: "Zones humides", u: "ha" }, { n: "Mangroves", u: "ha" },
      { n: "Arbres plantés", s: "Arbres", u: "nb" }, { n: "Arbres survivants", s: "Arbres", u: "nb" },
      { n: "Arbres coupés", s: "Arbres", u: "nb" }, { n: "Taux de survie des arbres", s: "Arbres", u: "%" },
    ],
  },
  {
    prefix: "AGRI", sheet: "12_Agriculture_Urbaine", accent: palette.green,
    t: "Bloc 12 — Agriculture urbaine", sub: "Surfaces cultivées, production, bénéficiaires, rendement et besoins d'irrigation.",
    inds: [
      { n: "Surface cultivée", u: "ha" }, { n: "Production", u: "t" },
      { n: "Nombre de bénéficiaires", u: "nb" }, { n: "Rendement", u: "t/ha" }, { n: "Irrigation", u: "m³" },
    ],
  },
  {
    prefix: "SANT", sheet: "13_Sante_Climat", accent: palette.red,
    t: "Bloc 13 — Santé et climat", sub: "Maladies sensibles au climat, hospitalisations et mortalité liées à la chaleur et à la pollution.",
    inds: [
      { n: "Paludisme", u: "cas" }, { n: "Dengue", u: "cas" }, { n: "Choléra", u: "cas" },
      { n: "Maladies respiratoires", u: "cas" }, { n: "Hospitalisations liées à la chaleur", u: "cas" },
      { n: "Mortalité liée à la chaleur", u: "décès" }, { n: "Mortalité liée à la pollution", u: "décès" },
    ],
  },
  {
    prefix: "FIN", sheet: "14_Finance_Climat", accent: palette.amber,
    t: "Bloc 14 — Finance climat", sub: "Budgets, dépenses climat, financements mobilisés par bailleur et portefeuille de projets.",
    inds: [
      { n: "Budget total de la Ville", s: "Budget", u: "FCFA" }, { n: "Budget climat", s: "Budget", u: "FCFA" },
      { n: "Dépenses climat", s: "Budget", u: "FCFA" },
      { n: "Financement — Union Européenne", s: "Financement", u: "FCFA" }, { n: "Financement — AFD", s: "Financement", u: "FCFA" },
      { n: "Financement — BAD", s: "Financement", u: "FCFA" }, { n: "Financement — Banque Mondiale", s: "Financement", u: "FCFA" },
      { n: "Financement — Fonds Vert Climat (GCF)", s: "Financement", u: "FCFA" }, { n: "Financement — C40", s: "Financement", u: "FCFA" },
      { n: "Financement — ICLEI", s: "Financement", u: "FCFA" },
      { n: "Nombre de projets", s: "Portefeuille", u: "nb" }, { n: "Coût du portefeuille", s: "Portefeuille", u: "FCFA" },
      { n: "Financement obtenu", s: "Portefeuille", u: "FCFA" }, { n: "Financement recherché", s: "Portefeuille", u: "FCFA" },
    ],
  },
];

/* ----- positions (pour la synthèse) : pos[prefix][nom] = {sheet,row} ---- */
const pos = {};
blocks.forEach((b) => {
  pos[b.prefix] = { __sheet: b.sheet };
  b.inds.forEach((ind, i) => { pos[b.prefix][ind.n] = 5 + i; });
});
const ref = (prefix, nom, year) => `'${pos[prefix].__sheet}'!${yearCol(year)}${pos[prefix][nom]}`;

/* =========================================================================
   CRÉATION DES FEUILLES (ordre d'affichage)
   ========================================================================= */
const wsAccueil = wb.worksheets.add("00_LISEZ_MOI");
const wsSynth = wb.worksheets.add("00_SYNTHESE_GLOBALE");
const wsCadres = wb.worksheets.add("Cadres_Reporting");
const blockWs = blocks.map((b) => ({ b, ws: wb.worksheets.add(b.sheet) }));
const wsProjets = wb.worksheets.add("15_Projets_Climat");

/* ----------------------------- blocs ----------------------------------- */
function buildBlock(b, ws) {
  title(ws, b.t, b.sub, END, b.accent);
  const header = ["Code", "Indicateur", "Sous-bloc", "Unité", ...YEARS.map(String), "Source / Institution", "Statut"];
  const data = b.inds.map((ind, i) => {
    const code = `${b.prefix}-${String(i + 1).padStart(3, "0")}`;
    return [code, ind.n, ind.s || "—", ind.u || "", ...YEARS.map(() => null), "", ""];
  });
  const last = 4 + data.length;
  ws.getRange(`A4:${END}4`).values = [header];
  ws.getRange(`A5:${END}${last}`).values = data;
  ws.getRange(`A4:${END}4`).format = { fill: b.accent, font: { bold: true, color: palette.white }, wrapText: true };
  ws.getRange(`A4:${END}4`).format.rowHeightPx = 30;
  ws.getRange(`A5:${END}${last}`).format = { fill: palette.white, wrapText: true };
  // largeurs
  const widths = [95, 320, 150, 90, ...YEARS.map(() => 66), 210, 110];
  widths.forEach((w, i) => { const c = colName(i + 1); ws.getRange(`${c}1:${c}${last + 4}`).format.columnWidthPx = w; });
  // zébrage léger sur l'identité
  ws.getRange(`A5:D${last}`).format = { fill: palette.grey, wrapText: true };
  // validation Statut + format années
  ws.getRange(`${colName(STA_COL)}5:${colName(STA_COL)}${last}`).dataValidation = {
    rule: { type: "list", values: STATUTS },
    errorAlert: { showAlert: true, style: "information", title: "Statut", message: "Choisir un statut de la liste." },
  };
  ws.getRange(`${colName(YEAR_START)}5:${colName(YEAR_START + YEARS.length - 1)}${last}`).setNumberFormat("#,##0.###");
  ws.tables.add(`A4:${END}${last}`, true, `tbl_${b.prefix}`).style = "TableStyleMedium2";
  freeze(ws, 4, META);
}
blockWs.forEach(({ b, ws }) => buildBlock(b, ws));

/* ----------------------------- valeurs illustratives -------------------- */
// Population (ancre CDP 2024 : 1 514 796 hab en 2023, croissance ~2,6 %/an) + densité (sur 83 km²)
const popSheet = pos.SOCIO.__sheet;
const popRow = pos.SOCIO["Population totale"];
const densRow = pos.SOCIO["Densité"];
const POP2023 = 1514796, GROWTH = 0.026, AREA = 83;
const popByYear = {};
YEARS.forEach((y) => { popByYear[y] = Math.round(POP2023 * Math.pow(1 + GROWTH, y - 2023)); });
const popWs = blockWs.find((x) => x.b.prefix === "SOCIO").ws;
popWs.getRange(`${yearCol(Y0)}${popRow}:${yearCol(Y1)}${popRow}`).values = [YEARS.map((y) => popByYear[y])];
popWs.getRange(`${yearCol(Y0)}${densRow}:${yearCol(Y1)}${densRow}`).formulas = [YEARS.map((y) => `=${yearCol(y)}${popRow}/${AREA}`)];
popWs.getRange(`${colName(SRC_COL)}${popRow}`).values = [["CDP 2024 (2023) + projection illustrative 2,6 %/an"]];
popWs.getRange(`${colName(STA_COL)}${popRow}`).values = [["Estimé"]];
popWs.getRange(`${colName(SRC_COL)}${densRow}`).values = [["Calcul Ville (population / 83 km²)"]];
popWs.getRange(`${colName(STA_COL)}${densRow}`).values = [["Estimé"]];

/* =========================================================================
   FEUILLE PROJETS (registre projet — hors grille annuelle)
   ========================================================================= */
function buildProjets() {
  const h = ["Code", "Nom du projet", "Secteur", "Description", "Localisation", "Budget (FCFA)", "Coût actualisé (FCFA)",
    "État d'avancement", "Maturité", "Bailleur cible", "Réduction GES attendue (tCO2e/an)", "Population bénéficiaire",
    "Emplois créés", "Risques", "Source de financement", "Date de début", "Date de fin"];
  const endCol = colName(h.length);
  title(wsProjets, "Bloc 15 — Portefeuille de projets climatiques", "Registre des projets : un projet par ligne, avec budget, maturité, bailleur, impact GES et calendrier.", endCol, palette.violet);
  const NROW = 40;
  const blank = Array.from({ length: NROW }, (_, i) => [
    `=IF(B${5 + i}="","","PRJ-"&TEXT(ROW()-4,"000"))`,
  ]);
  wsProjets.getRange(`A4:${endCol}4`).values = [h];
  wsProjets.getRange(`A5:A${4 + NROW}`).formulas = blank;
  wsProjets.getRange(`A4:${endCol}4`).format = { fill: palette.violet, font: { bold: true, color: palette.white }, wrapText: true };
  wsProjets.getRange(`A4:${endCol}4`).format.rowHeightPx = 30;
  wsProjets.getRange(`A5:${endCol}${4 + NROW}`).format = { fill: palette.white, wrapText: true };
  const w = [90, 240, 150, 320, 170, 150, 160, 150, 130, 170, 180, 160, 120, 220, 180, 110, 110];
  w.forEach((x, i) => { const c = colName(i + 1); wsProjets.getRange(`${c}1:${c}${NROW + 8}`).format.columnWidthPx = x; });
  wsProjets.getRange(`A5:A${4 + NROW}`).format = { fill: palette.grey };
  wsProjets.getRange(`H5:H${4 + NROW}`).dataValidation = { rule: { type: "list", values: ["Idée", "Préfaisabilité", "Faisabilité", "Financement", "En cours", "Terminé", "Suspendu"] } };
  wsProjets.getRange(`I5:I${4 + NROW}`).dataValidation = { rule: { type: "list", values: ["Concept", "Préparation", "Bancable", "Mise en œuvre", "Clôturé"] } };
  wsProjets.getRange(`F5:G${4 + NROW}`).setNumberFormat("#,##0");
  wsProjets.getRange(`P5:Q${4 + NROW}`).setNumberFormat("yyyy-mm-dd");
  wsProjets.tables.add(`A4:${endCol}${4 + NROW}`, true, "tbl_PROJ").style = "TableStyleMedium4";
  freeze(wsProjets, 4, 2);
}
buildProjets();

/* =========================================================================
   CADRES DE REPORTING
   ========================================================================= */
function buildCadres() {
  title(wsCadres, "Cadres de reporting climat couverts", "Référentiels internationaux et nationaux alimentés par la base, et blocs de données concernés.", "F", palette.navy);
  const h = ["Cadre", "Niveau", "Objet", "Données principales demandées", "Périodicité", "Blocs concernés"];
  const rows = [
    ["CDP-ICLEI Track (Unified Reporting)", "International", "Questionnaire annuel villes : gouvernance, risques, inventaire, secteurs, cibles, planification, actions, finance", "Inventaire GES, risques & vulnérabilité, énergie, mobilité, déchets, eau, santé, air, cibles, finance", "Annuelle", "Tous"],
    ["C40 Cities", "International", "Suivi de l'action climat des grandes villes", "Émissions GES, trajectoires, actions d'atténuation & adaptation, finance", "Annuelle", "GES, Énergie, Mobilité, Déchets, Finance"],
    ["GCoM (Global Covenant of Mayors)", "International", "Engagement intégré atténuation + adaptation + énergie", "BEI/IRE, EVR/RVA, accès à l'énergie, plan d'action (PAAEDC/SEACAP)", "Bisannuelle", "GES, Risques, Énergie, Gouvernance"],
    ["CoM SSA / SEACAP", "Régional (Afrique subsah.)", "Modèle de reporting Covenant of Mayors Afrique subsaharienne", "IRE (inventaire de référence), EVR (vulnérabilité/risques), EAE (accès énergie), actions", "Bisannuelle", "GES, Risques, Énergie, Gouvernance"],
    ["ICLEI", "International", "Réseau gouvernements locaux pour la durabilité", "Inventaires, plans, indicateurs de durabilité urbaine", "Variable", "Tous"],
    ["CDN Sénégal", "National", "Contribution Déterminée au niveau National", "Émissions et réductions par secteur, adaptation, finance climat", "Quinquennale + revues", "GES, Énergie, Risques, Finance"],
    ["PCET / Plan Climat Dakar", "Local", "Plan Climat-Énergie Territorial de la Ville", "Diagnostic, objectifs, actions, suivi-évaluation", "Continu", "Tous"],
    ["Fonds Vert Climat (GCF)", "International (finance)", "Accès au financement climat", "Réduction GES attendue, bénéficiaires, cofinancement, readiness", "Par projet", "Finance, Projets, GES"],
  ];
  wsCadres.getRange("A4:F4").values = [h];
  wsCadres.getRange(`A5:F${4 + rows.length}`).values = rows;
  wsCadres.getRange("A4:F4").format = { fill: palette.navy, font: { bold: true, color: palette.white }, wrapText: true };
  wsCadres.getRange(`A5:F${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  const w = [240, 180, 300, 360, 150, 280];
  w.forEach((x, i) => { const c = colName(i + 1); wsCadres.getRange(`${c}1:${c}${rows.length + 8}`).format.columnWidthPx = x; });
  wsCadres.tables.add(`A4:F${4 + rows.length}`, true, "tbl_CADRES").style = "TableStyleMedium2";
  freeze(wsCadres, 4, 1);
}
buildCadres();

/* =========================================================================
   SYNTHÈSE GLOBALE : KPI + couverture + projections + graphes
   ========================================================================= */
function buildSynthese() {
  title(wsSynth, "Synthèse globale — Base d'indicateurs climat de Dakar",
    "Vue de pilotage : couverture par bloc, indicateurs clés, projections (tendance linéaire) et graphiques illustratifs. Tout se met à jour quand vous saisissez les données dans les feuilles de blocs.",
    "T", palette.navy);

  // --- KPI cards ---
  const kpis = [
    ["Blocs de données", `${blocks.length + 1}`],
    ["Indicateurs annuels", `${blocks.reduce((s, b) => s + b.inds.length, 0)}`],
    ["Période couverte", `${Y0} – ${Y1}`],
    ["Cadres de reporting", "8"],
  ];
  kpis.forEach((k, i) => {
    const c0 = colName(1 + i * 3), c1 = colName(3 + i * 3);
    wsSynth.getRange(`${c0}4:${c1}4`).merge();
    wsSynth.getRange(`${c0}4`).values = [[k[0]]];
    wsSynth.getRange(`${c0}4:${c1}4`).format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
    wsSynth.getRange(`${c0}5:${c1}5`).merge();
    wsSynth.getRange(`${c0}5`).values = [[k[1]]];
    wsSynth.getRange(`${c0}5:${c1}5`).format = { fill: palette.lightTeal, font: { bold: true, size: 18, color: palette.navy }, horizontalAlignment: "center" };
    wsSynth.getRange(`${c0}5:${c1}5`).format.rowHeightPx = 40;
  });

  // --- Couverture par bloc ---
  let r = 8;
  wsSynth.getRange(`A${r}:E${r}`).merge();
  wsSynth.getRange(`A${r}`).values = [["COUVERTURE & COMPLÉTUDE PAR BLOC"]];
  wsSynth.getRange(`A${r}:E${r}`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  r += 1;
  wsSynth.getRange(`A${r}:E${r}`).values = [["Bloc", "Indicateurs", "Valeurs saisies", "Complétude (théorique)", "Statut"]];
  wsSynth.getRange(`A${r}:E${r}`).format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  const covStart = r + 1;
  const allBlocks = [...blocks, { prefix: "PROJ", sheet: "15_Projets_Climat", t: "Bloc 15 — Projets", inds: [] }];
  blocks.forEach((b, i) => {
    const rr = covStart + i;
    const ncol = colName(N_COLS);
    const y0 = colName(YEAR_START), y1 = colName(YEAR_START + YEARS.length - 1);
    wsSynth.getRange(`A${rr}`).values = [[b.t]];
    wsSynth.getRange(`B${rr}`).formulas = [[`=COUNTA('${b.sheet}'!A5:A${4 + b.inds.length})`]];
    wsSynth.getRange(`C${rr}`).formulas = [[`=COUNT('${b.sheet}'!${y0}5:${y1}${4 + b.inds.length})`]];
    wsSynth.getRange(`D${rr}`).formulas = [[`=IFERROR(C${rr}/(B${rr}*${YEARS.length}),0)`]];
    wsSynth.getRange(`E${rr}`).formulas = [[`=IF(C${rr}=0,"À démarrer",IF(D${rr}>=0.6,"Bien couvert",IF(D${rr}>=0.2,"En cours","Initial")))`]];
  });
  const covEnd = covStart + blocks.length - 1;
  wsSynth.getRange(`A${covStart}:E${covEnd}`).format = { fill: palette.white, wrapText: true };
  wsSynth.getRange(`D${covStart}:D${covEnd}`).setNumberFormat("0%");
  wsSynth.getRange(`B${covStart}:C${covEnd}`).setNumberFormat("#,##0");
  const totR = covEnd + 1;
  wsSynth.getRange(`A${totR}`).values = [["TOTAL"]];
  wsSynth.getRange(`B${totR}`).formulas = [[`=SUM(B${covStart}:B${covEnd})`]];
  wsSynth.getRange(`C${totR}`).formulas = [[`=SUM(C${covStart}:C${covEnd})`]];
  wsSynth.getRange(`D${totR}`).formulas = [[`=IFERROR(C${totR}/(B${totR}*${YEARS.length}),0)`]];
  wsSynth.getRange(`A${totR}:E${totR}`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  wsSynth.getRange(`D${totR}`).setNumberFormat("0%");
  wsSynth.tables.add(`A${covStart - 1}:E${covEnd}`, true, "tbl_COUVERTURE").style = "TableStyleMedium9";

  // --- Indicateurs clés × année + projection (mini-tables pour graphes) ---
  // Chaque mini-table : ligne "Année" + ligne "valeur" (référence vers le bloc), placée en colonne, puis graphe.
  const keys = [
    { label: "Population totale (hab)", prefix: "SOCIO", nom: "Population totale", type: "line", fmt: "#,##0", accent: palette.slate },
    { label: "Émissions GES totales (tCO2e)", prefix: "GES", nom: "Émissions totales (tous secteurs)", type: "line", fmt: "#,##0", accent: palette.green },
    { label: "Émissions par habitant (tCO2e/hab)", prefix: "GES", nom: "Émissions par habitant", type: "line", fmt: "#,##0.00", accent: palette.teal },
    { label: "Budget climat (FCFA)", prefix: "FIN", nom: "Budget climat", type: "column", fmt: "#,##0", accent: palette.amber },
    { label: "Surface d'espaces verts (ha)", prefix: "VERT", nom: "Surface d'espaces verts", type: "line", fmt: "#,##0", accent: palette.leaf },
  ];

  let base = totR + 3;
  wsSynth.getRange(`A${base}:T${base}`).merge();
  wsSynth.getRange(`A${base}`).values = [["INDICATEURS CLÉS — SUIVI ANNUEL, PROJECTIONS & GRAPHIQUES"]];
  wsSynth.getRange(`A${base}:T${base}`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  base += 2;

  keys.forEach((k) => {
    // bloc de données : col A = Année, col B = Valeur (réf), 21 lignes
    const headRow = base;
    wsSynth.getRange(`A${headRow}:B${headRow}`).merge();
    wsSynth.getRange(`A${headRow}`).values = [[k.label]];
    wsSynth.getRange(`A${headRow}:B${headRow}`).format = { fill: k.accent, font: { bold: true, color: palette.white }, wrapText: true };
    wsSynth.getRange(`A${headRow + 1}:B${headRow + 1}`).values = [["Année", "Valeur"]];
    wsSynth.getRange(`A${headRow + 1}:B${headRow + 1}`).format = { fill: palette.lightTeal, font: { bold: true } };
    const d0 = headRow + 2;
    wsSynth.getRange(`A${d0}:A${d0 + YEARS.length - 1}`).values = YEARS.map((y) => [y]);
    wsSynth.getRange(`B${d0}:B${d0 + YEARS.length - 1}`).formulas = YEARS.map((y) => [`=IFERROR(IF(${ref(k.prefix, k.nom, y)}="","",${ref(k.prefix, k.nom, y)}),"")`]);
    wsSynth.getRange(`B${d0}:B${d0 + YEARS.length - 1}`).setNumberFormat(k.fmt);
    const dEnd = d0 + YEARS.length - 1;
    // projections (tendance linéaire sur l'historique disponible)
    const pr = dEnd + 1;
    wsSynth.getRange(`A${pr}`).values = [["Proj. 2030"]];
    wsSynth.getRange(`B${pr}`).formulas = [[`=IFERROR(FORECAST.LINEAR(2030,B${d0}:B${dEnd},A${d0}:A${dEnd}),"")`]];
    wsSynth.getRange(`A${pr + 1}`).values = [["Proj. 2035"]];
    wsSynth.getRange(`B${pr + 1}`).formulas = [[`=IFERROR(FORECAST.LINEAR(2035,B${d0}:B${dEnd},A${d0}:A${dEnd}),"")`]];
    wsSynth.getRange(`A${pr}:A${pr + 1}`).format = { fill: palette.lightAmber, font: { bold: true } };
    wsSynth.getRange(`B${pr}:B${pr + 1}`).setNumberFormat(k.fmt);
    wsSynth.getRange(`A${d0}:A${dEnd}`).setNumberFormat("0");
    // graphe
    const chart = wsSynth.charts.add(k.type, wsSynth.getRange(`A${headRow + 1}:B${dEnd}`));
    chart.title = k.label;
    chart.hasLegend = false;
    chart.xAxis = { axisType: "textAxis" };
    chart.yAxis = { numberFormatCode: k.fmt };
    chart.setPosition(`D${headRow}`, `T${dEnd}`);
    base = pr + 4;
  });

  // largeurs synthèse
  const w = [150, 150, 150, 150, 150, 110, 110, 110, 110, 110, 110, 110];
  w.forEach((x, i) => { const c = colName(i + 1); wsSynth.getRange(`${c}1:${c}3`).format.columnWidthPx = x; });
  wsSynth.getRange("A1:A3").format.columnWidthPx = 220;
  wsSynth.getRange("B1:B3").format.columnWidthPx = 150;
}
buildSynthese();

/* =========================================================================
   ACCUEIL / LISEZ-MOI
   ========================================================================= */
function para(ws, r, txt, opts = {}) {
  ws.getRange(`A${r}:H${r}`).merge();
  ws.getRange(`A${r}`).values = [[txt]];
  ws.getRange(`A${r}:H${r}`).format = {
    fill: opts.fill || palette.white,
    font: { bold: !!opts.bold, color: opts.color || palette.slate, size: opts.size || 11 },
    wrapText: true,
  };
  ws.getRange(`A${r}:H${r}`).format.rowHeightPx = opts.h || 26;
  return r + 1;
}
function buildAccueil() {
  title(wsAccueil, "Base d'indicateurs climat — Ville de Dakar", "Données organisées PAR BLOC et PAR ANNÉE (2015–2035). Reconstruite à partir du questionnaire CDP 2024, du modèle CoM SSA/SEACAP et de la taxonomie des 15 blocs.", "H", palette.navy);
  let r = 4;
  r = para(wsAccueil, r, "Comment utiliser cette base", { bold: true, fill: palette.lightGreen, color: palette.navy, size: 13, h: 30 });
  r = para(wsAccueil, r, "• Chaque BLOC a sa propre feuille (01 à 15). Les indicateurs sont en lignes, les années (2015→2035) en colonnes : vous saisissez/laissez les valeurs année par année.");
  r = para(wsAccueil, r, "• Colonnes communes : Code · Indicateur · Sous-bloc · Unité · [années] · Source/Institution · Statut (Brouillon, À vérifier, Validé, Estimé, Projection).");
  r = para(wsAccueil, r, "• La feuille « 00_SYNTHESE_GLOBALE » se met à jour automatiquement : complétude par bloc, indicateurs clés, projections (tendance) et graphiques.");
  r = para(wsAccueil, r, "• La feuille « 15_Projets_Climat » est un registre projet (un projet par ligne), distinct de la grille annuelle.");
  r += 1;
  r = para(wsAccueil, r, "Capacité cible", { bold: true, fill: palette.lightGreen, color: palette.navy, size: 13, h: 30 });
  r = para(wsAccueil, r, "Conçue pour ~500–800 indicateurs, 50–100 institutions sources et 10–20 ans d'historique. La base actuelle couvre 15 blocs et plus de 180 indicateurs annuels prêts à compléter.");
  r += 1;
  r = para(wsAccueil, r, "Cadres de reporting alimentés", { bold: true, fill: palette.lightGreen, color: palette.navy, size: 13, h: 30 });
  r = para(wsAccueil, r, "CDP · C40 · GCoM · CoM SSA / SEACAP · ICLEI · CDN Sénégal · PCET / Plan Climat Dakar · Fonds Vert Climat (GCF). Détails dans la feuille « Cadres_Reporting ».");
  r += 1;
  r = para(wsAccueil, r, "Repères Ville de Dakar (source CDP 2024)", { bold: true, fill: palette.lightGreen, color: palette.navy, size: 13, h: 30 });
  r = para(wsAccueil, r, "Périmètre administratif : 83 km². Population : 1 514 796 habitants (2023). Langue de reporting : français. Type : ville / municipalité.");
  r = para(wsAccueil, r, "La feuille « 02_Socio_Economie » contient une projection illustrative de population (2,6 %/an) ; à remplacer par les données ANSD/recensement.", { color: palette.amber });
  r += 1;
  r = para(wsAccueil, r, "Liste des 15 blocs", { bold: true, fill: palette.lightGreen, color: palette.navy, size: 13, h: 30 });
  blocks.forEach((b) => { r = para(wsAccueil, r, `${b.t}  (${b.inds.length} indicateurs)`); });
  r = para(wsAccueil, r, "Bloc 15 — Portefeuille de projets climatiques  (registre projet)");
  const w = [260, 120, 120, 120, 120, 120, 120, 120];
  w.forEach((x, i) => { const c = colName(i + 1); wsAccueil.getRange(`${c}1:${c}${r + 2}`).format.columnWidthPx = x; });
}
buildAccueil();

/* ----------------------------- sortie & contrôle ------------------------ */
await fs.mkdir(outputDir, { recursive: true });
const errs = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 200 }, maxChars: 4000 });
console.log("FORMULA_ERRORS");
console.log(errs.ndjson);
const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(outputFile);
const info = await wb.inspect({ kind: "workbook", include: "sheets", maxChars: 9000 });
console.log("WORKBOOK");
console.log(info.ndjson);
console.log("OUTPUT_FILE", outputFile);
