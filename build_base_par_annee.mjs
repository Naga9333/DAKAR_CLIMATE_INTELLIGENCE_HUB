import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

/* =========================================================================
   DAKAR — BASE CLIMAT : UNE FEUILLE PAR ANNÉE (2016–2050)
   Chaque feuille année = liste complète des indicateurs (regroupés par bloc),
   une seule colonne Valeur. + SYNTHESE consolidée + VISUALISATION progressive.
   Sources : CDP 2024, CoM SSA/SEACAP, taxonomie 15 blocs.
   Sortie : outputs/DAKAR_CLIMATE_BASE_PAR_ANNEE.xlsx
   ========================================================================= */

const outputDir = path.resolve("outputs");
const outputFile = path.join(outputDir, "DAKAR_CLIMATE_BASE_PAR_ANNEE.xlsx");
const wb = await Workbook.create();

const palette = {
  navy: "#0B1F33", teal: "#0F766E", green: "#166534", leaf: "#2E7D32",
  amber: "#B45309", red: "#B91C1C", blue: "#1D4ED8", violet: "#6D28D9",
  cyan: "#0E7490", brown: "#7C4A03", slate: "#334155",
  lightTeal: "#E6F4F1", lightGreen: "#EAF7EA", lightAmber: "#FEF3C7",
  lightRed: "#FEE2E2", lightBlue: "#E8F0FE", white: "#FFFFFF", grey: "#F8FAFC",
};
const STATUTS = ["Brouillon", "À vérifier", "Validé", "Estimé", "Projection", "N/D"];

const Y0 = 2016, Y1 = 2050;
const YEARS = [];
for (let y = Y0; y <= Y1; y += 1) YEARS.push(y);

function colName(n) {
  let s = "", x = n;
  while (x > 0) { const r = (x - 1) % 26; s = String.fromCharCode(65 + r) + s; x = Math.floor((x - 1) / 26); }
  return s;
}
function title(ws, t, sub, endCol, accent = palette.navy) {
  ws.showGridLines = false;
  ws.getRange(`A1:${endCol}1`).merge();
  ws.getRange("A1").values = [[t]];
  ws.getRange(`A1:${endCol}1`).format = { fill: accent, font: { bold: true, color: palette.white, size: 15 } };
  ws.getRange(`A1:${endCol}1`).format.rowHeightPx = 32;
  ws.getRange(`A2:${endCol}2`).merge();
  ws.getRange("A2").values = [[sub]];
  ws.getRange(`A2:${endCol}2`).format = { fill: palette.lightGreen, font: { italic: true, color: palette.slate }, wrapText: true };
  ws.getRange(`A2:${endCol}2`).format.rowHeightPx = 38;
}
function freeze(ws, rows, cols) {
  try { ws.freezePanes.freezeRows(rows); } catch (e) { /* noop */ }
  try { ws.freezePanes.freezeColumns(cols); } catch (e) { /* noop */ }
}
function setW(ws, widths, rowLimit) {
  widths.forEach((w, i) => { const c = colName(i + 1); ws.getRange(`${c}1:${c}${rowLimit}`).format.columnWidthPx = w; });
}

/* ----------------------------- 14 blocs annuels ------------------------- */
const RISKS = ["Inondations", "Érosion côtière", "Submersion marine", "Canicules", "Sécheresse",
  "Tempêtes", "Feux de végétation", "Pollution atmosphérique", "Maladies climatiques", "Stress hydrique"];
const RISK_METRICS = [{ m: "Fréquence", u: "nb/an" }, { m: "Intensité", u: "échelle 1-5" }, { m: "Population exposée", u: "hab" }, { m: "Pertes économiques", u: "FCFA" }];
const riskIndicators = RISKS.flatMap((r) => RISK_METRICS.map((x) => ({ n: `${r} — ${x.m}`, s: r, u: x.u })));

const blocks = [
  { prefix: "GOUV", nom: "Gouvernance climat", accent: palette.navy, inds: [
    { n: "Existence d'un Plan Climat", u: "Oui/Non" }, { n: "Date d'adoption du Plan Climat", u: "date" },
    { n: "Budget climat annuel", u: "FCFA" }, { n: "Dépenses climat", u: "FCFA" }, { n: "Personnel climat", u: "ETP" },
    { n: "Nombre d'agents formés", u: "nb" }, { n: "Nombre de réunions COPIL", u: "nb" }, { n: "Nombre de consultations publiques", u: "nb" },
    { n: "Nombre de partenaires", u: "nb" }, { n: "Existence d'un système MRV", u: "Oui/Non" },
    { n: "Existence d'un budget vert", u: "Oui/Non" }, { n: "Existence d'un inventaire GES", u: "Oui/Non" } ] },
  { prefix: "SOCIO", nom: "Données socio-économiques", accent: palette.slate, inds: [
    { n: "Population totale", u: "hab" }, { n: "Population — part de femmes", u: "%" }, { n: "Population jeune (<15 ans)", u: "%" },
    { n: "Population âgée (>65 ans)", u: "%" }, { n: "Densité", u: "hab/km²" }, { n: "Nombre de ménages", u: "nb" },
    { n: "PIB local estimé", u: "FCFA" }, { n: "Taux de pauvreté", u: "%" }, { n: "Taux de chômage", u: "%" },
    { n: "Population vulnérable", u: "hab" }, { n: "Population exposée aux risques", u: "hab" } ] },
  { prefix: "CLIM", nom: "Climat et météorologie", accent: palette.cyan, inds: [
    { n: "Température moyenne", u: "°C" }, { n: "Température minimale", u: "°C" }, { n: "Température maximale", u: "°C" },
    { n: "Température ressentie", u: "°C" }, { n: "Nombre de jours chauds", u: "j" }, { n: "Nombre de vagues de chaleur", u: "nb" },
    { n: "Pluviométrie", u: "mm" }, { n: "Nombre de jours de pluie", u: "j" }, { n: "Intensité des pluies", u: "mm/j" },
    { n: "Humidité", u: "%" }, { n: "Vitesse du vent", u: "m/s" }, { n: "Direction du vent", u: "°" },
    { n: "Rayonnement solaire", u: "kWh/m²" }, { n: "Évapotranspiration", u: "mm" }, { n: "Niveau marin", u: "cm" } ] },
  { prefix: "RISQ", nom: "Risques climatiques", accent: palette.red, inds: riskIndicators },
  { prefix: "EAU", nom: "Eau", accent: palette.blue, inds: [
    { n: "Production d'eau", u: "m³" }, { n: "Consommation d'eau", u: "m³" }, { n: "Pertes réseau", u: "%" },
    { n: "Accès à l'eau potable", u: "%" }, { n: "Qualité de l'eau (indice)", u: "indice" }, { n: "Niveau des nappes phréatiques", u: "m" },
    { n: "Capacité des réservoirs", u: "m³" }, { n: "Nombre de forages", u: "nb" }, { n: "Eaux usées produites", u: "m³" }, { n: "Eaux usées traitées", u: "m³" } ] },
  { prefix: "ENER", nom: "Énergie", accent: palette.amber, inds: [
    { n: "Consommation totale", u: "MWh" }, { n: "Consommation résidentielle", u: "MWh" }, { n: "Consommation commerciale", u: "MWh" },
    { n: "Consommation industrielle", u: "MWh" }, { n: "Consommation municipale", u: "MWh" }, { n: "Éclairage public", u: "MWh" },
    { n: "Production solaire", u: "MWh" }, { n: "Production éolienne", u: "MWh" }, { n: "Production biomasse", u: "MWh" },
    { n: "Accès à l'électricité", u: "%" }, { n: "Coupures", u: "nb/an" }, { n: "Facteur d'émission électricité", u: "tCO2e/MWh" } ] },
  { prefix: "GES", nom: "Inventaire GES", accent: palette.green, inds: [
    { n: "Énergie — Combustibles", s: "Énergie", u: "tCO2e" }, { n: "Énergie — Électricité", s: "Énergie", u: "tCO2e" },
    { n: "Transport — Routier", s: "Transport", u: "tCO2e" }, { n: "Transport — Ferroviaire", s: "Transport", u: "tCO2e" },
    { n: "Transport — Maritime", s: "Transport", u: "tCO2e" }, { n: "Transport — Aérien local", s: "Transport", u: "tCO2e" },
    { n: "Déchets — Décharge", s: "Déchets", u: "tCO2e" }, { n: "Déchets — Recyclage", s: "Déchets", u: "tCO2e" }, { n: "Déchets — Compostage", s: "Déchets", u: "tCO2e" },
    { n: "AFOLU — Agriculture urbaine", s: "AFOLU", u: "tCO2e" }, { n: "AFOLU — Espaces verts", s: "AFOLU", u: "tCO2e" },
    { n: "Émissions Scope 1", s: "Résultats", u: "tCO2e" }, { n: "Émissions Scope 2", s: "Résultats", u: "tCO2e" }, { n: "Émissions Scope 3", s: "Résultats", u: "tCO2e" },
    { n: "Émissions totales (tous secteurs)", s: "Résultats", u: "tCO2e" }, { n: "Émissions par habitant", s: "Résultats", u: "tCO2e/hab" } ] },
  { prefix: "MOB", nom: "Mobilité", accent: palette.violet, inds: [
    { n: "Nombre de véhicules", s: "Parc", u: "nb" }, { n: "Véhicules particuliers", s: "Parc", u: "nb" }, { n: "Taxis", s: "Parc", u: "nb" },
    { n: "Bus", s: "Parc", u: "nb" }, { n: "Cars rapides", s: "Parc", u: "nb" }, { n: "Ndiaga Ndiaye", s: "Parc", u: "nb" },
    { n: "Motos", s: "Parc", u: "nb" }, { n: "Camions", s: "Parc", u: "nb" },
    { n: "Transport collectif (passagers)", s: "Transport collectif", u: "passagers/an" }, { n: "Passagers BRT", s: "Transport collectif", u: "passagers/an" },
    { n: "Passagers TER", s: "Transport collectif", u: "passagers/an" }, { n: "Passagers DDD", s: "Transport collectif", u: "passagers/an" },
    { n: "Part de la marche", s: "Mobilité active", u: "%" }, { n: "Part du vélo", s: "Mobilité active", u: "%" }, { n: "Pistes cyclables", s: "Mobilité active", u: "km" },
    { n: "Véhicules électriques", s: "Électromobilité", u: "nb" }, { n: "Bornes de recharge", s: "Électromobilité", u: "nb" } ] },
  { prefix: "AIR", nom: "Qualité de l'air", accent: palette.brown, inds: [
    { n: "PM2.5", u: "µg/m³" }, { n: "PM10", u: "µg/m³" }, { n: "NO₂", u: "µg/m³" }, { n: "SO₂", u: "µg/m³" },
    { n: "CO", u: "mg/m³" }, { n: "O₃", u: "µg/m³" }, { n: "Black Carbon", u: "µg/m³" }, { n: "Nombre de stations", u: "nb" }, { n: "Dépassements OMS", u: "j/an" } ] },
  { prefix: "DECH", nom: "Déchets", accent: palette.teal, inds: [
    { n: "Déchets collectés", s: "Collecte", u: "t" }, { n: "Déchets ménagers", s: "Collecte", u: "t" }, { n: "Déchets commerciaux", s: "Collecte", u: "t" },
    { n: "Déchets industriels", s: "Collecte", u: "t" }, { n: "Déchets plastiques", s: "Collecte", u: "t" }, { n: "Déchets organiques", s: "Collecte", u: "t" },
    { n: "Déchets recyclés", s: "Valorisation", u: "t" }, { n: "Déchets compostés", s: "Valorisation", u: "t" }, { n: "Déchets réutilisés", s: "Valorisation", u: "t" },
    { n: "Déchets valorisés (total)", s: "Valorisation", u: "t" }, { n: "Quantité enfouie", s: "Décharge", u: "t" }, { n: "Méthane estimé", s: "Décharge", u: "tCO2e" } ] },
  { prefix: "VERT", nom: "Espaces verts et biodiversité", accent: palette.leaf, inds: [
    { n: "Surface d'espaces verts", u: "ha" }, { n: "Espaces verts par habitant", u: "m²/hab" }, { n: "Parcs", u: "nb" },
    { n: "Jardins publics", u: "nb" }, { n: "Jardins communautaires", u: "nb" }, { n: "Micro-jardins", u: "nb" },
    { n: "Forêts urbaines", u: "ha" }, { n: "Corridors verts", u: "km" }, { n: "Zones humides", u: "ha" }, { n: "Mangroves", u: "ha" },
    { n: "Arbres plantés", s: "Arbres", u: "nb" }, { n: "Arbres survivants", s: "Arbres", u: "nb" }, { n: "Arbres coupés", s: "Arbres", u: "nb" }, { n: "Taux de survie des arbres", s: "Arbres", u: "%" } ] },
  { prefix: "AGRI", nom: "Agriculture urbaine", accent: palette.green, inds: [
    { n: "Surface cultivée", u: "ha" }, { n: "Production", u: "t" }, { n: "Nombre de bénéficiaires", u: "nb" }, { n: "Rendement", u: "t/ha" }, { n: "Irrigation", u: "m³" } ] },
  { prefix: "SANT", nom: "Santé et climat", accent: palette.red, inds: [
    { n: "Paludisme", u: "cas" }, { n: "Dengue", u: "cas" }, { n: "Choléra", u: "cas" }, { n: "Maladies respiratoires", u: "cas" },
    { n: "Hospitalisations liées à la chaleur", u: "cas" }, { n: "Mortalité liée à la chaleur", u: "décès" }, { n: "Mortalité liée à la pollution", u: "décès" } ] },
  { prefix: "FIN", nom: "Finance climat", accent: palette.amber, inds: [
    { n: "Budget total de la Ville", s: "Budget", u: "FCFA" }, { n: "Budget climat", s: "Budget", u: "FCFA" }, { n: "Dépenses climat", s: "Budget", u: "FCFA" },
    { n: "Financement — Union Européenne", s: "Financement", u: "FCFA" }, { n: "Financement — AFD", s: "Financement", u: "FCFA" }, { n: "Financement — BAD", s: "Financement", u: "FCFA" },
    { n: "Financement — Banque Mondiale", s: "Financement", u: "FCFA" }, { n: "Financement — Fonds Vert Climat (GCF)", s: "Financement", u: "FCFA" },
    { n: "Financement — C40", s: "Financement", u: "FCFA" }, { n: "Financement — ICLEI", s: "Financement", u: "FCFA" },
    { n: "Nombre de projets", s: "Portefeuille", u: "nb" }, { n: "Coût du portefeuille", s: "Portefeuille", u: "FCFA" },
    { n: "Financement obtenu", s: "Portefeuille", u: "FCFA" }, { n: "Financement recherché", s: "Portefeuille", u: "FCFA" } ] },
];

/* ----- liste à plat : FLAT[k] -> ligne (5+k) identique dans chaque feuille année ---- */
const FLAT = [];
blocks.forEach((b) => b.inds.forEach((ind, i) => {
  FLAT.push({ code: `${b.prefix}-${String(i + 1).padStart(3, "0")}`, bloc: b.nom, prefix: b.prefix, sous: ind.s || "—", nom: ind.n, unite: ind.u || "", accent: b.accent });
}));
const N = FLAT.length;
const DATA0 = 5, DATAE = 4 + N;
const rowOf = (prefix, nom) => 5 + FLAT.findIndex((f) => f.prefix === prefix && f.nom === nom);

/* population (ancre CDP 2023 = 1 514 796, croissance illustrative 2,6 %/an) */
const POP2023 = 1514796, GROWTH = 0.0314, AREA = 83;  // 3,14 %/an : hypothèse de croissance CDP/Plan Climat
const popByYear = {}; YEARS.forEach((y) => { popByYear[y] = Math.round(POP2023 * Math.pow(1 + GROWTH, y - 2023)); });
const popRow = rowOf("SOCIO", "Population totale");
const densRow = rowOf("SOCIO", "Densité");

/* =========================================================================
   CRÉATION DES FEUILLES (ordre d'affichage)
   ========================================================================= */
const wsAccueil = wb.worksheets.add("00_LISEZ_MOI");
const wsSynth = wb.worksheets.add("SYNTHESE_DONNEES");
const wsViz = wb.worksheets.add("VISUALISATION_PROGRESSIVE");
const wsCadres = wb.worksheets.add("Cadres_Reporting");
const yearWs = YEARS.map((y) => ({ y, ws: wb.worksheets.add(String(y)) }));
const wsProjets = wb.worksheets.add("Projets_Climat");

/* ----------------------------- feuilles année --------------------------- */
const YH = ["Code", "Bloc", "Sous-bloc", "Indicateur", "Unité", "Valeur", "Source / Institution", "Statut"];
const yearWidths = [95, 200, 150, 300, 90, 130, 220, 110];
function buildYear(y, ws) {
  title(ws, `Année ${y} — Indicateurs climat de Dakar`,
    `Saisissez la valeur de chaque indicateur pour l'année ${y}. Regroupement par bloc (colonne Bloc, filtrable). La synthèse et les graphiques se mettent à jour automatiquement.`,
    "H", palette.navy);
  ws.getRange("A4:H4").values = [YH];
  const data = FLAT.map((f) => [f.code, f.bloc, f.sous, f.nom, f.unite, null, "", ""]);
  ws.getRange(`A${DATA0}:H${DATAE}`).values = data;
  ws.getRange("A4:H4").format = { fill: palette.navy, font: { bold: true, color: palette.white }, wrapText: true };
  ws.getRange("A4:H4").format.rowHeightPx = 28;
  ws.getRange(`A${DATA0}:E${DATAE}`).format = { fill: palette.grey, wrapText: true };
  ws.getRange(`F${DATA0}:H${DATAE}`).format = { fill: palette.white, wrapText: true };
  setW(ws, yearWidths, DATAE + 4);
  ws.getRange(`F${DATA0}:F${DATAE}`).setNumberFormat("#,##0.###");
  ws.getRange(`H${DATA0}:H${DATAE}`).dataValidation = { rule: { type: "list", values: STATUTS }, errorAlert: { showAlert: true, style: "information", title: "Statut", message: "Choisir un statut." } };
  ws.tables.add(`A4:H${DATAE}`, true, `tbl_${y}`).style = "TableStyleMedium2";
  freeze(ws, 4, 5);
  // amorce population + densité
  ws.getRange(`F${popRow}`).values = [[popByYear[y]]];
  ws.getRange(`G${popRow}`).values = [[y === 2023 ? "ANSD / CDP 2024" : "Projection 3,14 %/an (hypothèse CDP) — ancre 2023"]];
  ws.getRange(`H${popRow}`).values = [[y === 2023 ? "Validé" : "Projection"]];
  ws.getRange(`F${densRow}`).formulas = [[`=IF(F${popRow}="","",F${popRow}/${AREA})`]];
  ws.getRange(`G${densRow}`).values = [["Calcul (population / 83 km²)"]];
  ws.getRange(`H${densRow}`).values = [["Estimé"]];
}
yearWs.forEach(({ y, ws }) => buildYear(y, ws));

/* =========================================================================
   AMORÇAGE AVEC LES DONNÉES RÉELLES (CDP 2024 + inventaire GES 2016)
   ========================================================================= */
function put(prefix, nom, year, value, statut = "Validé", source = "CDP 2024 — Ville de Dakar") {
  const yw = yearWs.find((x) => x.y === year); if (!yw) return;
  const row = rowOf(prefix, nom); if (row < 5) return;
  yw.ws.getRange(`F${row}`).values = [[value]];
  yw.ws.getRange(`G${row}`).values = [[source]];
  yw.ws.getRange(`H${row}`).values = [[statut]];
}
const SRC_INV = "Inventaire GES Ville de Dakar 2016 (CDP 2024)";
const SRC_CDP = "CDP 2024 — Ville de Dakar";
// --- Inventaire GES 2016 (IPCC 2006, vérifié) ---
put("GES", "Énergie — Combustibles", 2016, 472888, "Validé", SRC_INV);
put("GES", "Énergie — Électricité", 2016, 779935, "Validé", SRC_INV);
put("GES", "Transport — Routier", 2016, 1058462, "Validé", SRC_INV);
put("GES", "Transport — Ferroviaire", 2016, 254, "Validé", SRC_INV);
put("GES", "Déchets — Décharge", 2016, 19551, "Validé", SRC_INV);
put("GES", "Émissions Scope 1", 2016, 1697834, "Validé", SRC_INV);
put("GES", "Émissions Scope 2", 2016, 780807, "Validé", SRC_INV);
put("GES", "Émissions Scope 3", 2016, 149074, "Validé", SRC_INV);
put("GES", "Émissions totales (tous secteurs)", 2016, 2626842, "Validé", SRC_INV);
put("GES", "Émissions par habitant", 2016, 1.73, "Validé", SRC_INV);
// Cibles GES (Plan Climat / CDP) — émissions nettes visées
put("GES", "Émissions totales (tous secteurs)", 2030, 3291611, "Projection", "Cible CDP : −25 % vs BAU (base 2016)");
put("GES", "Émissions totales (tous secteurs)", 2050, 4550086, "Projection", "Cible CDP : −54 % vs BAU (base 2016)");
// --- Énergie ---
put("ENER", "Accès à l'électricité", 2016, 97.3, "Validé", "CDP 2024 (4.4) / ANSD");
put("ENER", "Coupures", 2016, 66, "Validé", "CDP 2024 (4.4) — heures/an");
put("ENER", "Consommation totale", 2016, 5314.84, "À vérifier", "CDP 2024 (4.1) — valeur à confirmer");
// --- Mobilité (parc & parts modales) ---
put("MOB", "Véhicules particuliers", 2016, 154379, "Validé", SRC_CDP);
put("MOB", "Bus", 2016, 5987, "Validé", SRC_CDP);
put("MOB", "Taxis", 2016, 22713, "Validé", SRC_CDP);
put("MOB", "Véhicules électriques", 2016, 34, "Validé", SRC_CDP);
put("MOB", "Part de la marche", 2016, 70, "Validé", "CDP 2024 (4.5)");
put("MOB", "Part du vélo", 2016, 7.3, "Validé", "CDP 2024 (4.5)");
// --- Déchets / Eau ---
put("DECH", "Déchets collectés", 2016, 291517, "Validé", "CDP 2024 (4.7)");
put("DECH", "Déchets valorisés (total)", 2016, 5830, "Estimé", "CDP 2024 (4.7) — 2 % détournés de la décharge");
put("DECH", "Déchets recyclés", 2016, 816, "Estimé", "CDP 2024 (4.7) — 14 % des déchets détournés");
put("EAU", "Accès à l'eau potable", 2016, 100, "Validé", "CDP 2024 (4.10)");
put("EAU", "Eaux usées produites", 2016, 53164075, "Validé", "CDP 2024 (4.7) — 145 655 m³/j");
// --- Qualité de l'air (CGQA 2019) ---
put("AIR", "PM2.5", 2019, 57.14, "Validé", "CGQA Dakar 2019 (CDP 2024)");
put("AIR", "Nombre de stations", 2019, 6, "Validé", "CGQA Dakar (CDP 2024)");
// --- Risques climatiques (évaluation CDP 2024) ---
const p24 = popByYear[2024];
put("RISQ", "Inondations — Population exposée", 2024, Math.round(0.355 * p24), "Validé", "CDP 2024 (2.2) — 31-40 %");
put("RISQ", "Inondations — Intensité", 2024, 4, "Validé", "CDP 2024 (2.2) — magnitude Medium High");
put("RISQ", "Submersion marine — Population exposée", 2024, Math.round(0.255 * p24), "Validé", "CDP 2024 (2.2) — 21-30 %");
put("RISQ", "Submersion marine — Intensité", 2024, 4, "Validé", "CDP 2024 (2.2)");
put("RISQ", "Canicules — Population exposée", 2024, Math.round(0.455 * p24), "Validé", "CDP 2024 (2.2) — 41-50 %");
put("RISQ", "Canicules — Intensité", 2024, 4, "Validé", "CDP 2024 (2.2)");
put("RISQ", "Sécheresse — Population exposée", 2024, Math.round(0.155 * p24), "Validé", "CDP 2024 (2.2) — 11-20 %");
put("RISQ", "Sécheresse — Intensité", 2024, 2, "Validé", "CDP 2024 (2.2) — magnitude Low");
// --- Gouvernance ---
put("GOUV", "Date d'adoption du Plan Climat", 2021, "31/03/2021", "Validé", "PCET adopté le 31/03/2021 (CDP 2024)");
for (let y = 2021; y <= Y1; y += 1) put("GOUV", "Existence d'un Plan Climat", y, "Oui", "Validé", "PCET Dakar (adopté 2021)");
for (let y = 2016; y <= Y1; y += 1) put("GOUV", "Existence d'un inventaire GES", y, "Oui", "Validé", "Inventaire GES base 2016");
for (let y = 2021; y <= Y1; y += 1) put("GOUV", "Existence d'un système MRV", y, "Oui", "Validé", "Reporting CoM SSA / SEACAP (canevas MRV)");
for (let y = 2023; y <= Y1; y += 1) put("GOUV", "Existence d'un budget vert", y, "Oui", "Validé", "Budget 2023 intégrant le climat (CDP 2024)");

/* =========================================================================
   SYNTHESE_DONNEES : matrice indicateurs × années (consolidation auto)
   ========================================================================= */
function buildSynthese() {
  const yc = (y) => colName(6 + (y - Y0));      // colonne d'une année
  const endCol = colName(5 + YEARS.length);
  title(wsSynth, "Synthèse des données — vue globale consolidée",
    "Consolidation automatique de toutes les feuilles année (lecture seule). Une ligne = un indicateur, une colonne = une année. Filtrez par bloc pour une lecture ciblée.",
    endCol, palette.green);
  const header = [...YH.slice(0, 5), ...YEARS.map(String)];
  wsSynth.getRange(`A4:${endCol}4`).values = [header];
  const meta = FLAT.map((f) => [f.code, f.bloc, f.sous, f.nom, f.unite]);
  wsSynth.getRange(`A${DATA0}:E${DATAE}`).values = meta;
  YEARS.forEach((y) => {
    const c = yc(y);
    const col = [];
    for (let k = 0; k < N; k += 1) col.push([`='${y}'!F${5 + k}`]);
    wsSynth.getRange(`${c}${DATA0}:${c}${DATAE}`).formulas = col;
  });
  wsSynth.getRange(`A4:${endCol}4`).format = { fill: palette.green, font: { bold: true, color: palette.white }, wrapText: true };
  wsSynth.getRange("A4:E4").format.rowHeightPx = 28;
  wsSynth.getRange(`A${DATA0}:E${DATAE}`).format = { fill: palette.grey, wrapText: true };
  wsSynth.getRange(`${colName(6)}${DATA0}:${endCol}${DATAE}`).setNumberFormat("#,##0.###");
  setW(wsSynth, [95, 190, 150, 300, 90, ...YEARS.map(() => 72)], DATAE + 4);
  wsSynth.tables.add(`A4:${endCol}${DATAE}`, true, "tbl_SYNTHESE").style = "TableStyleMedium9";
  freeze(wsSynth, 4, 5);
}
buildSynthese();

/* =========================================================================
   VISUALISATION_PROGRESSIVE : remplissage par année + courbes clés
   ========================================================================= */
function buildViz() {
  title(wsViz, "Visualisation globale — progression par année",
    "Avancement du remplissage de la base au fil des années et évolution des indicateurs clés. Les graphiques se peuplent au fur et à mesure de la saisie.",
    "T", palette.blue);

  // Bloc 1 : progression du remplissage (nb d'indicateurs renseignés / année)
  let r = 4;
  wsViz.getRange(`A${r}:B${r}`).merge();
  wsViz.getRange(`A${r}`).values = [["Progression du remplissage"]];
  wsViz.getRange(`A${r}:B${r}`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  wsViz.getRange(`A${r + 1}:B${r + 1}`).values = [["Année", "Indicateurs renseignés"]];
  wsViz.getRange(`A${r + 1}:B${r + 1}`).format = { fill: palette.lightBlue, font: { bold: true }, wrapText: true };
  const d0 = r + 2;
  wsViz.getRange(`A${d0}:A${d0 + YEARS.length - 1}`).values = YEARS.map((y) => [y]);
  wsViz.getRange(`B${d0}:B${d0 + YEARS.length - 1}`).formulas = YEARS.map((y) => [`=COUNT('${y}'!F${DATA0}:F${DATAE})`]);
  const dEnd = d0 + YEARS.length - 1;
  wsViz.getRange(`A${d0}:A${dEnd}`).setNumberFormat("0");
  const ch0 = wsViz.charts.add("column", wsViz.getRange(`A${r + 1}:B${dEnd}`));
  ch0.title = "Indicateurs renseignés par année";
  ch0.hasLegend = false; ch0.xAxis = { axisType: "textAxis" }; ch0.yAxis = { numberFormatCode: "#,##0" };
  ch0.setPosition(`D${r}`, `T${dEnd}`);

  // Blocs suivants : indicateurs clés (courbe d'évolution)
  const keys = [
    { label: "Population totale (hab)", prefix: "SOCIO", nom: "Population totale", type: "line", fmt: "#,##0" },
    { label: "Émissions GES totales (tCO2e)", prefix: "GES", nom: "Émissions totales (tous secteurs)", type: "line", fmt: "#,##0" },
    { label: "Émissions par habitant (tCO2e/hab)", prefix: "GES", nom: "Émissions par habitant", type: "line", fmt: "#,##0.00" },
    { label: "Budget climat (FCFA)", prefix: "FIN", nom: "Budget climat", type: "column", fmt: "#,##0" },
    { label: "Surface d'espaces verts (ha)", prefix: "VERT", nom: "Surface d'espaces verts", type: "line", fmt: "#,##0" },
  ];
  let base = dEnd + 3;
  keys.forEach((k) => {
    const row = rowOf(k.prefix, k.nom);
    const hr = base;
    wsViz.getRange(`A${hr}:B${hr}`).merge();
    wsViz.getRange(`A${hr}`).values = [[k.label]];
    wsViz.getRange(`A${hr}:B${hr}`).format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
    wsViz.getRange(`A${hr + 1}:B${hr + 1}`).values = [["Année", "Valeur"]];
    wsViz.getRange(`A${hr + 1}:B${hr + 1}`).format = { fill: palette.lightTeal, font: { bold: true } };
    const k0 = hr + 2;
    wsViz.getRange(`A${k0}:A${k0 + YEARS.length - 1}`).values = YEARS.map((y) => [y]);
    wsViz.getRange(`B${k0}:B${k0 + YEARS.length - 1}`).formulas = YEARS.map((y) => [`=IFERROR(IF('${y}'!F${row}="","",'${y}'!F${row}),"")`]);
    const kEnd = k0 + YEARS.length - 1;
    wsViz.getRange(`A${k0}:A${kEnd}`).setNumberFormat("0");
    wsViz.getRange(`B${k0}:B${kEnd}`).setNumberFormat(k.fmt);
    const ch = wsViz.charts.add(k.type, wsViz.getRange(`A${hr + 1}:B${kEnd}`));
    ch.title = k.label; ch.hasLegend = false; ch.xAxis = { axisType: "textAxis" }; ch.yAxis = { numberFormatCode: k.fmt };
    ch.setPosition(`D${hr}`, `T${kEnd}`);
    base = kEnd + 3;
  });
  setW(wsViz, [110, 150, 110, 110, 110, 110, 110, 110], base + 2);
}
buildViz();

/* =========================================================================
   CADRES + PROJETS + ACCUEIL
   ========================================================================= */
function buildCadres() {
  title(wsCadres, "Cadres de reporting climat couverts", "Référentiels alimentés par la base et blocs de données concernés.", "F", palette.navy);
  const h = ["Cadre", "Niveau", "Objet", "Données principales demandées", "Périodicité", "Blocs concernés"];
  const rows = [
    ["CDP-ICLEI Track", "International", "Questionnaire annuel villes (gouvernance → actions)", "Inventaire GES, risques, énergie, mobilité, déchets, eau, santé, air, cibles, finance", "Annuelle", "Tous"],
    ["C40 Cities", "International", "Suivi de l'action climat des grandes villes", "Émissions, trajectoires, actions, finance", "Annuelle", "GES, Énergie, Mobilité, Déchets, Finance"],
    ["GCoM", "International", "Atténuation + adaptation + énergie", "BEI/IRE, EVR/RVA, accès énergie, plan d'action", "Bisannuelle", "GES, Risques, Énergie, Gouvernance"],
    ["CoM SSA / SEACAP", "Régional", "Covenant of Mayors Afrique subsaharienne", "IRE, EVR, EAE, actions", "Bisannuelle", "GES, Risques, Énergie, Gouvernance"],
    ["ICLEI", "International", "Durabilité des gouvernements locaux", "Inventaires, plans, indicateurs", "Variable", "Tous"],
    ["CDN Sénégal", "National", "Contribution Déterminée au niveau National", "Émissions/réductions par secteur, adaptation, finance", "Quinquennale", "GES, Énergie, Risques, Finance"],
    ["PCET / Plan Climat Dakar", "Local", "Plan Climat-Énergie Territorial", "Diagnostic, objectifs, actions, suivi", "Continu", "Tous"],
    ["Fonds Vert Climat (GCF)", "Finance", "Accès au financement climat", "Réduction GES, bénéficiaires, cofinancement", "Par projet", "Finance, Projets, GES"],
  ];
  wsCadres.getRange("A4:F4").values = [h];
  wsCadres.getRange(`A5:F${4 + rows.length}`).values = rows;
  wsCadres.getRange("A4:F4").format = { fill: palette.navy, font: { bold: true, color: palette.white }, wrapText: true };
  wsCadres.getRange(`A5:F${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  setW(wsCadres, [200, 140, 280, 360, 130, 260], rows.length + 8);
  wsCadres.tables.add(`A4:F${4 + rows.length}`, true, "tbl_CADRES").style = "TableStyleMedium2";
  freeze(wsCadres, 4, 1);
}
buildCadres();

function buildProjets() {
  const h = ["Code", "Nom du projet", "Secteur", "Description", "Localisation", "Budget (FCFA)", "Coût actualisé (FCFA)",
    "État d'avancement", "Maturité", "Bailleur cible", "Réduction GES attendue (tCO2e/an)", "Population bénéficiaire", "Emplois créés", "Risques", "Source de financement", "Date de début", "Date de fin"];
  const endCol = colName(h.length);
  title(wsProjets, "Portefeuille de projets climatiques", "Registre des projets (un projet par ligne) : budget, maturité, bailleur, impact GES et calendrier.", endCol, palette.violet);
  const NROW = 40;
  wsProjets.getRange(`A4:${endCol}4`).values = [h];
  wsProjets.getRange(`A5:A${4 + NROW}`).formulas = Array.from({ length: NROW }, (_, i) => [`=IF(B${5 + i}="","","PRJ-"&TEXT(ROW()-4,"000"))`]);
  wsProjets.getRange(`A4:${endCol}4`).format = { fill: palette.violet, font: { bold: true, color: palette.white }, wrapText: true };
  wsProjets.getRange(`A4:${endCol}4`).format.rowHeightPx = 30;
  wsProjets.getRange(`A5:${endCol}${4 + NROW}`).format = { fill: palette.white, wrapText: true };
  setW(wsProjets, [90, 240, 150, 320, 170, 150, 160, 150, 130, 170, 180, 160, 120, 220, 180, 110, 110], NROW + 8);
  wsProjets.getRange(`H5:H${4 + NROW}`).dataValidation = { rule: { type: "list", values: ["Idée", "Préfaisabilité", "Faisabilité", "Financement", "En cours", "Terminé", "Suspendu"] } };
  wsProjets.getRange(`I5:I${4 + NROW}`).dataValidation = { rule: { type: "list", values: ["Concept", "Préparation", "Bancable", "Mise en œuvre", "Clôturé"] } };
  wsProjets.getRange(`F5:G${4 + NROW}`).setNumberFormat("#,##0");
  wsProjets.getRange(`P5:Q${4 + NROW}`).setNumberFormat("yyyy-mm-dd");
  // projets réels (CDP 2024)
  const PROJ = [
    ["Quartier durable et prospère de Biscuiterie", "Aménagement urbain / écoquartier", "Écoquartier pilote : énergie durable, gestion communautaire des déchets, espaces publics verts.", "Commune de Biscuiterie", "En cours", ""],
    ["Réhabilitation de la voirie non classée", "Mobilité / Transport", "Réhabilitation du réseau de voiries non classées en appui au BRT et à la mobilité durable (coût déclaré CDP : 17, unité à confirmer).", "Ville de Dakar", "En cours", ""],
    ["Régénération urbaine de la corniche HLM", "Espace public / Transport", "Régénération urbaine (espace public ≈ 15 % du territoire vs norme 40 %). Appui C40 Cities Finance Facility (CFF).", "Corniche HLM", "Financement", ""],
    ["Réaménagement du bassin de rétention de Grand Yoff", "Eau / Inondations", "Adaptation : réaménagement hydraulique et paysager du bassin (bassin versant 700 ha ; coût déclaré CDP : 7).", "Grand Yoff", "En cours", 300000],
    ["Unité de compostage des déchets alimentaires", "Déchets", "Centre de tri et de compostage des déchets alimentaires (marché Castor), partenariat Ville de Milan.", "Marché Castor", "Préfaisabilité", ""],
    ["Programme de 57 espaces publics verts", "Espaces verts / Adaptation", "Construction de 57 espaces publics verts et récréatifs ; objectif +150 ha d'espaces verts d'ici 2030.", "Ville de Dakar", "En cours", ""],
    ["Modernisation de l'éclairage public (LED)", "Énergie / Efficacité", "Remplacement des lampes à sodium par des LED sur les grandes artères pour réduire la facture énergétique.", "Ville de Dakar", "En cours", ""],
    ["20 écoquartiers d'ici 2027", "Aménagement urbain", "Programme écoquartier : 20 quartiers intégrant énergie durable, gestion des déchets et espaces publics verts.", "Ville de Dakar", "Financement", ""],
  ];
  wsProjets.getRange(`B5:E${4 + PROJ.length}`).values = PROJ.map((p) => [p[0], p[1], p[2], p[3]]);
  PROJ.forEach((p, i) => {
    wsProjets.getRange(`H${5 + i}`).values = [[p[4]]];
    if (p[5] !== "" && p[5] != null) wsProjets.getRange(`L${5 + i}`).values = [[p[5]]];
  });
  wsProjets.tables.add(`A4:${endCol}${4 + NROW}`, true, "tbl_PROJ").style = "TableStyleMedium4";
  freeze(wsProjets, 4, 2);
}
buildProjets();

function para(ws, r, txt, opts = {}) {
  ws.getRange(`A${r}:H${r}`).merge();
  ws.getRange(`A${r}`).values = [[txt]];
  ws.getRange(`A${r}:H${r}`).format = { fill: opts.fill || palette.white, font: { bold: !!opts.bold, color: opts.color || palette.slate, size: opts.size || 11 }, wrapText: true };
  ws.getRange(`A${r}:H${r}`).format.rowHeightPx = opts.h || 26;
  return r + 1;
}
function buildAccueil() {
  title(wsAccueil, "Base climat Dakar — une feuille PAR ANNÉE", `Indicateurs des 15 blocs, saisis année par année (${Y0}–${Y1}). Reconstruite à partir du CDP 2024, du modèle CoM SSA/SEACAP et de la taxonomie des 15 blocs.`, "H", palette.navy);
  let r = 4;
  r = para(wsAccueil, r, "Principe", { bold: true, fill: palette.lightGreen, color: palette.navy, size: 13, h: 30 });
  r = para(wsAccueil, r, `• Une feuille par année (${Y0} à ${Y1}). Chaque feuille liste TOUS les indicateurs (colonne Bloc pour le regroupement) avec une seule colonne Valeur — on remplit une année à la fois, sans base longue.`);
  r = para(wsAccueil, r, "• Colonnes : Code · Bloc · Sous-bloc · Indicateur · Unité · Valeur · Source/Institution · Statut.");
  r = para(wsAccueil, r, "• « SYNTHESE_DONNEES » consolide automatiquement toutes les années (indicateurs en lignes × années en colonnes) — lecture seule.");
  r = para(wsAccueil, r, "• « VISUALISATION_PROGRESSIVE » montre la progression du remplissage par année et l'évolution des indicateurs clés (graphiques).");
  r = para(wsAccueil, r, "• « Projets_Climat » : registre des projets. « Cadres_Reporting » : cadres alimentés (CDP, C40, GCoM, CoM SSA, ICLEI, CDN, PCET, GCF).");
  r += 1;
  r = para(wsAccueil, r, "Chiffres clés", { bold: true, fill: palette.lightGreen, color: palette.navy, size: 13, h: 30 });
  r = para(wsAccueil, r, `${N} indicateurs annuels × ${YEARS.length} années (${Y0}–${Y1}). Capacité cible : 500–800 indicateurs, 50–100 institutions, 10–20 ans.`);
  r = para(wsAccueil, r, "Données réelles pré-remplies (CDP 2024 / inventaire GES 2016) :", { bold: true, fill: palette.lightGreen, color: palette.navy, size: 13, h: 30 });
  r = para(wsAccueil, r, "• GES 2016 : 2 626 842 tCO2e (Scope 1 : 1 697 834 · Scope 2 : 780 807 · Scope 3 : 149 074) ; cibles −25 % en 2030 et −54 % en 2050. Inventaire IPCC 2006, vérifié.");
  r = para(wsAccueil, r, "• Énergie : accès électricité 97,3 % (2016). Mobilité : 154 379 voitures, 5 987 bus, 22 713 taxis, 34 VE ; parts marche 70 % / vélo 7,3 %. Déchets : 291 517 t/an. Eaux usées : 53,16 Mm³/an. PM2.5 : 57,14 µg/m³ (2019). Eau potable : 100 %.");
  r = para(wsAccueil, r, "• Risques (CDP 2.2) : exposition canicules 41-50 %, inondations 31-40 %, submersion 21-30 %, sécheresse 11-20 %. Gouvernance : PCET adopté le 31/03/2021. 8 projets réels listés dans « Projets_Climat ».");
  r = para(wsAccueil, r, "Repères : 83 km², 1 514 796 hab. (2023), reporting en français. Population projetée à 3,14 %/an (hypothèse CDP) — affiner avec l'ANSD. Cases vides = à compléter.", { color: palette.amber });
  setW(wsAccueil, [260, 120, 120, 120, 120, 120, 120, 120], r + 2);
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
console.log("WORKBOOK_SHEETS", JSON.parse(info.ndjson.trim().split("\n")[0] || "{}").sheets ?? "?");
console.log("INDICATEURS", N, "ANNEES", YEARS.length);
console.log("OUTPUT_FILE", outputFile);
