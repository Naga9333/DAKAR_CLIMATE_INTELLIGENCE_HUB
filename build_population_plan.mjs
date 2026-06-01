import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputFile = path.resolve("outputs", "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V2.xlsx");
const outputDir = path.resolve("outputs");
const outputFile = path.join(outputDir, "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V3_PLAN_PEUPLEMENT.xlsx");

const input = await FileBlob.load(inputFile);
const workbook = await SpreadsheetFile.importXlsx(input);

const palette = {
  navy: "#0B1F33",
  teal: "#0F766E",
  lightTeal: "#E6F4F1",
  lightGreen: "#EAF7EA",
  lightAmber: "#FFF7E6",
  lightRed: "#FEE2E2",
  white: "#FFFFFF",
  slate: "#334155",
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

function addSheet(name, title, subtitle = "") {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1:Q1").merge();
  sheet.getRange("A1:Q1").format = { fill: palette.navy, font: { bold: true, color: palette.white, size: 15 } };
  sheet.getRange("A1:Q1").format.rowHeightPx = 34;
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2:Q2").merge();
  sheet.getRange("A2:Q2").format = { fill: palette.lightTeal, font: { italic: true, color: palette.slate }, wrapText: true };
  sheet.getRange("A2:Q2").format.rowHeightPx = 42;
  return sheet;
}

function setWidths(sheet, widths, rowLimit = 140) {
  widths.forEach((w, index) => {
    const col = colName(index + 1);
    sheet.getRange(`${col}1:${col}${rowLimit}`).format.columnWidthPx = w;
  });
}

function addValidation(sheet, colIndex, startRow, endRow, values) {
  const col = colName(colIndex);
  sheet.getRange(`${col}${startRow}:${col}${endRow}`).dataValidation = {
    rule: { type: "list", values },
    prompt: { showPrompt: true, title: "Liste controlee", message: "Choisir une valeur autorisee." },
    errorAlert: { showAlert: true, style: "warning", title: "Valeur inattendue", message: "Verifier la liste." },
  };
}

function addStatusFormatting(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.conditionalFormats.add("containsText", { text: "P0", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "P1", format: { fill: palette.lightAmber, font: { color: "#92400E", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "P2", format: { fill: palette.lightGreen, font: { color: "#166534", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "Oui", format: { fill: palette.lightGreen, font: { color: "#166534" } } });
  range.conditionalFormats.add("containsText", { text: "Partiel", format: { fill: palette.lightAmber, font: { color: "#92400E" } } });
  range.conditionalFormats.add("containsText", { text: "Non", format: { fill: palette.lightRed, font: { color: "#991B1B" } } });
}

const rows = [
  ["00_ACCUEIL", "Gouvernance du classeur, sources methodologiques, version", "Ville disponible", "Oui", "Partiel", "Non", "Oui", "Oui", "Oui", "Oui", "Cellule climat Ville de Dakar", "Cellule climat / PMO", "Semestrielle", "Faible", "P1", "Mettre a jour version, contacts et sources officielles 2026"],
  ["01_PARAMETRES", "Listes controlees: secteurs, statuts, quartiers, services", "Ville disponible", "Oui", "Partiel", "Non", "Oui", "Oui", "Oui", "Oui", "Secretariat general, directions municipales", "Administrateur Hub", "Trimestrielle", "Faible", "P1", "Valider les vocabulaires avec toutes les directions"],
  ["02_REFERENTIELS", "Perimetres, codes GPC/CDN, classifications", "Mixte Ville/partenaires", "Partiel", "Oui", "Partiel", "Oui", "Oui", "Oui", "Oui", "Ville, Ministere environnement, GHG Protocol, CDN Senegal", "Referent MRV", "Annuelle", "Moyenne", "P0", "Completer codes GPC et axes CDN officiels"],
  ["03_SOURCES_DONNEES", "Registre des sources, proprietaires, qualite, echeances", "A collecter", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "Toutes directions et partenaires", "Cellule MRV", "Mensuelle", "Moyenne", "P0", "Nommer un proprietaire par source de donnees"],
  ["04_TABLEAU_BORD", "KPI executifs: projets, GES, budget, alertes", "Produit par le Hub", "Oui", "Oui", "Partiel", "Oui", "Oui", "Oui", "Oui", "Onglets sources du Hub", "Cellule climat", "Mensuelle", "Moyenne", "P1", "Verifier coherences des KPI avant chaque comite"],
  ["05_PROJETS_CLIMAT", "Portefeuille projets, maturite, statut, budget, liens cadres", "Ville disponible", "Oui", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "Directions techniques, cabinet, PMO projets", "Cellule climat / directions", "Mensuelle", "Moyenne", "P0", "Completer Project_ID et lier chaque projet a CDN/CDP/C40"],
  ["06_INDICATEURS_MRV", "Indicateurs, baseline, cibles, valeurs actuelles", "A collecter", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "Projets, directions, partenaires techniques", "Cellule MRV", "Mensuelle/Trimestrielle", "Elevee", "P0", "Definir baseline et cible pour chaque indicateur prioritaire"],
  ["07_ACTIONS_CLIMAT", "Actions, echeances, avancement, dependances", "Ville disponible", "Oui", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "Plans d'action, directions techniques", "PMO climat", "Mensuelle", "Moyenne", "P0", "Rattacher toutes les actions aux projets et indicateurs"],
  ["08_INVENTAIRE_GES", "Emissions par secteur GPC, scope, gaz, methode", "Partenaires + a collecter", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "Ministere, SENELEC, UCG/SONAGED, CETUD, ANSD", "Referent GES", "Annuelle", "Tres elevee", "P0", "Produire inventaire GPC BASIC puis BASIC+"],
  ["09_DONNEES_GES_ACTIVITE", "Donnees d'activite: kWh, litres, tonnes, passagers-km", "Partenaires + a collecter", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "Factures, operateurs, enquêtes, directions sectorielles", "Referent GES + directions", "Mensuelle/Annuelle", "Tres elevee", "P0", "Prioriser energie, transport et dechets 2024-2026"],
  ["10_FACTEURS_EMISSION", "Facteurs emission et PRG/GWP", "Partenaires", "Non", "Oui", "Partiel", "Oui", "Oui", "Oui", "Oui", "Ministere environnement, SENELEC, IPCC, GHG Protocol", "Referent GES", "Annuelle", "Elevee", "P0", "Valider facteurs nationaux ou sources par defaut justifiees"],
  ["11_QUALITE_AIR", "Mesures PM2.5, PM10, NO2, seuils et depassements", "Partenaires + a collecter", "Partiel", "Oui", "Oui", "Oui", "Oui", "Partiel", "Partiel", "CGQA/DEEC, capteurs, universites, startups", "Direction Environnement", "Mensuelle", "Elevee", "P1", "Mettre en place protocole QA/QC capteurs"],
  ["12_MOBILITE", "Flotte, trafic, parts modales, carburants, passagers", "Partenaires + a collecter", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "CETUD, AFTU, Direction mobilite, enquêtes menages", "Direction Mobilite", "Trimestrielle/Annuelle", "Tres elevee", "P0", "Collecter donnees carburant et part modale"],
  ["13_ENERGIE", "Consommation energie, couts, renouvelables, sites", "Mixte Ville/partenaires", "Oui", "Oui", "Partiel", "Oui", "Oui", "Oui", "Oui", "Factures Ville, SENELEC, services municipaux", "Direction Energie", "Mensuelle", "Moyenne", "P0", "Centraliser factures et sites municipaux"],
  ["14_DECHETS", "Tonnes generees, collectees, recyclees, enfouies, brulees", "Mixte Ville/partenaires", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "UCG/SONAGED, Direction dechets, prestataires", "Direction Dechets", "Mensuelle/Trimestrielle", "Elevee", "P0", "Harmoniser flux, tonnages et destinations finales"],
  ["15_PARTIES_PRENANTES", "Organisations, contacts, roles, engagement", "Ville disponible", "Oui", "Partiel", "Partiel", "Oui", "Oui", "Partiel", "Oui", "Cabinet, directions, partenaires", "Cellule climat", "Trimestrielle", "Faible", "P1", "Confirmer points focaux et roles MRV"],
  ["16_BUDGETS", "Budgets prevus, votes, engages, executes", "Ville disponible", "Oui", "Non", "Partiel", "Oui", "Oui", "Oui", "Oui", "Direction finances, budget municipal", "Direction Finance", "Mensuelle", "Moyenne", "P0", "Mapper lignes budgetaires aux projets climat"],
  ["17_FINANCEMENTS", "Bailleurs, montants demandes/approuves/decaisses", "Mixte Ville/partenaires", "Partiel", "Oui", "Partiel", "Oui", "Oui", "Oui", "Oui", "Direction finance, bailleurs, conventions", "Direction Finance", "Mensuelle", "Moyenne", "P0", "Mettre a jour pipeline et preuves de cofinancement"],
  ["18_RISQUES_HYPOTHESES", "Risques projet, probabilite, impact, mitigation", "Ville disponible", "Oui", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "PMO projets, directions, partenaires execution", "PMO climat", "Mensuelle", "Moyenne", "P1", "Lier risques aux echeances et budgets"],
  ["19_QA_QC", "Controles qualite, anomalies, corrections", "A collecter", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Cellule MRV, auditeurs internes", "Cellule MRV", "Mensuelle", "Moyenne", "P0", "Institutionnaliser un cycle QA/QC mensuel"],
  ["20_MAPPING_REPORTING", "Correspondance donnees internes vers cadres", "Produit par le Hub", "Partiel", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "CDP, C40, GCoM, CoM SSA, CDN, GCF", "Cellule reporting", "Annuelle", "Moyenne", "P0", "Mettre a jour avec questionnaire CDP 2026"],
  ["21_RAPPORT_CDP_ICLEI", "Champs CDP-ICLEI pre-remplis", "Produit par le Hub", "Partiel", "Partiel", "Oui", "Oui", "Partiel", "Partiel", "Partiel", "Onglets sources + CDP Portal", "Cellule reporting", "Annuelle", "Elevee", "P0", "Aligner le plan avec questionnaire et scoring CDP 2026"],
  ["22_RAPPORT_C40", "Indicateurs C40: CAP, actions, air, transport, dechets", "Produit par le Hub", "Partiel", "Partiel", "Oui", "Partiel", "Oui", "Partiel", "Partiel", "Onglets sources + C40 CAP/MER", "Cellule climat", "Semestrielle", "Elevee", "P0", "Ajouter indicateurs haute ambition et justice climatique"],
  ["23_RAPPORT_COMSSA", "SEACAP: mitigation, adaptation, acces energie", "Produit par le Hub", "Partiel", "Partiel", "Oui", "Partiel", "Partiel", "Oui", "Oui", "CoM SSA, Direction Energie, MRV", "Cellule climat / Direction Energie", "Annuelle", "Elevee", "P1", "Completer le pilier acces energie"],
  ["24_RAPPORT_CDN", "Alignement projets et resultats avec CDN Senegal", "Mixte Ville/partenaires", "Partiel", "Oui", "Partiel", "Partiel", "Partiel", "Oui", "Oui", "CDN Senegal, Ministere, projets Ville", "Cellule climat", "Annuelle", "Elevee", "P0", "Valider secteurs CDN et contributions quantifiees"],
  ["25_JOURNAL_MODIFS", "Historique modifications et validations", "Ville disponible", "Oui", "Non", "Partiel", "Oui", "Oui", "Oui", "Oui", "Administrateur Hub", "Administrateur Hub", "Continu", "Faible", "P1", "Rendre obligatoire pour modifications majeures"],
  ["26_AUDIT_EXPERT", "Faiblesses V1 et reponses V2", "Produit par le Hub", "Oui", "Non", "Non", "Oui", "Oui", "Oui", "Oui", "Audit interne + cadres internationaux", "Cellule climat", "Annuelle", "Moyenne", "P1", "Actualiser apres chaque cycle de reporting"],
  ["27_CONFORMITE_CADRES", "Score 0-5 par cadre et exigence", "A collecter", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "CDP, C40, GCoM, CoM SSA, GCF, WB, BAD", "Cellule reporting", "Trimestrielle", "Elevee", "P0", "Organiser revue de conformite trimestrielle"],
  ["28_TRAJECTOIRES_CIBLES", "Baseline, BAU, cible 2030, budget carbone", "A collecter/modeliser", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "Inventaire GES, modelisation, plan climat", "Referent GES + planification", "Annuelle", "Tres elevee", "P0", "Valider baseline et cible politique 2030"],
  ["29_RISQUES_CLIMATIQUES", "Aleas, exposition, vulnerabilite, population exposee", "Partenaires + a collecter", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "ANACIM, CSE, ONAS, universites, quartiers", "Direction Planification", "Annuelle", "Tres elevee", "P0", "Realiser RVA multi-aleas par quartier"],
  ["30_ACTIFS_EXPOSITION", "Actifs critiques geolocalises et exposes", "A collecter", "Partiel", "Oui", "Oui", "Partiel", "Oui", "Oui", "Oui", "SIG Ville, ONAS, cadastre, directions techniques", "Cellule SIG / Planification", "Annuelle", "Tres elevee", "P0", "Construire inventaire SIG des actifs critiques"],
  ["31_PIPELINE_BANCABLE", "Pipeline projets finance climat, cofinancement, readiness", "Ville + a collecter", "Partiel", "Partiel", "Oui", "Partiel", "Oui", "Oui", "Oui", "Directions, finance, bailleurs, business plans", "Direction Finance", "Mensuelle", "Elevee", "P0", "Preparer fiches concept note pour projets P0"],
  ["32_READINESS_FINANCE", "Criteres GCF/WB/BAD, E&S, genre, procurement", "A collecter", "Partiel", "Oui", "Oui", "Partiel", "Partiel", "Oui", "Oui", "GCF, Banque mondiale, BAD, sauvegardes, genre", "Direction Finance + sauvegardes", "Mensuelle", "Tres elevee", "P0", "Completer climate rationale, E&S et genre"],
  ["33_COBENEFICES_JUSTICE", "Emplois, sante, inclusion, genre, groupes vulnerables", "A collecter", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "Enquetes, projets, ONG, services sociaux", "Cellule climat + affaires sociales", "Semestrielle", "Elevee", "P1", "Definir indicateurs de justice climatique par projet"],
  ["34_REGISTRE_PREUVES_MRV", "Preuves, liens, validation et fiabilite", "A collecter", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Documents sources, extractions, conventions, audits", "Cellule MRV", "Mensuelle", "Moyenne", "P0", "Exiger preuve pour chaque donnee critique"],
  ["35_PLAN_DONNEES", "Lacunes de donnees et plan de collecte", "Produit par le Hub", "Partiel", "Oui", "Oui", "Oui", "Oui", "Oui", "Oui", "Toutes directions et partenaires", "Cellule MRV", "Mensuelle", "Moyenne", "P0", "Suivre les lacunes P0 en comite technique"],
  ["36_PASSATION_MARCHES", "Procurement, lots, montants, attribution, decaissement", "Ville disponible", "Oui", "Partiel", "Partiel", "Non", "Partiel", "Partiel", "Oui", "Direction marches, finance, projets", "Direction Marches / Finance", "Mensuelle", "Moyenne", "P1", "Relier procurement aux projets bancables"],
  ["37_ACCES_ENERGIE_SEACAP", "Acces energie propre, eclairage, cuisson, services", "Partenaires + a collecter", "Partiel", "Oui", "Oui", "Partiel", "Partiel", "Oui", "Oui", "SENELEC, Direction Energie, enquêtes menages", "Direction Energie", "Annuelle", "Elevee", "P1", "Lancer diagnostic energie acces par quartier"],
  ["38_INDICATEURS_INTL", "Bibliotheque indicateurs internationaux", "Produit par le Hub", "Oui", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "CDP, C40, GCoM, CoM SSA, GCF, WB, BAD", "Cellule MRV", "Annuelle", "Moyenne", "P1", "Actualiser apres publication des guides annuels"],
  ["39_DASHBOARD_V2", "Maturite internationale, risques, bancabilite", "Produit par le Hub", "Oui", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "Onglets V2", "Cellule climat", "Mensuelle", "Moyenne", "P1", "Presenter au comite climat mensuel"],
  ["40_SOURCES_INTL", "Sources officielles et liens de reference", "Produit par le Hub", "Oui", "Partiel", "Partiel", "Oui", "Oui", "Oui", "Oui", "Sites officiels CDP, C40, GCoM, CoM SSA, GCF, WB, BAD", "Cellule reporting", "Annuelle", "Faible", "P1", "Verifier liens et versions chaque annee"],
];

const headers = [
  "Feuille",
  "Donnee_a_peupler",
  "Categorie_principale",
  "Disponible_Ville_Dakar",
  "Disponible_partenaires",
  "A_collecter",
  "Critique_CDP",
  "Critique_C40",
  "Critique_CDN",
  "Critique_bailleurs_climat",
  "Source",
  "Responsable",
  "Frequence",
  "Niveau_difficulte",
  "Priorite",
  "Premiere_action",
];

const sheet = addSheet(
  "41_PLAN_PEUPLEMENT",
  "Plan de peuplement du Dakar Climate Intelligence Hub",
  "Matrice operationnelle de collecte et peuplement par feuille, classee selon disponibilite, criticite reporting et exigences bailleurs climat."
);

const startRow = 4;
const endRow = startRow + rows.length;
const endCol = colName(headers.length);
sheet.getRange(`A${startRow}:${endCol}${endRow}`).values = [headers, ...rows];
sheet.getRange(`A${startRow}:${endCol}${startRow}`).format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
sheet.getRange(`A${startRow + 1}:${endCol}${endRow}`).format = { fill: palette.white, wrapText: true };
const table = sheet.tables.add(`A${startRow}:${endCol}${endRow}`, true, "tblPlanPeuplement");
table.style = "TableStyleMedium2";
table.showFilterButton = true;
setWidths(sheet, [160, 310, 165, 105, 115, 95, 95, 95, 95, 130, 290, 200, 140, 135, 90, 320], endRow + 10);
sheet.freezePanes.freezeRows(startRow);

for (const colIndex of [4, 5, 6, 7, 8, 9, 10]) {
  addValidation(sheet, colIndex, startRow + 1, endRow, ["Oui", "Non", "Partiel"]);
  addStatusFormatting(sheet, `${colName(colIndex)}${startRow + 1}:${colName(colIndex)}${endRow}`);
}
addValidation(sheet, 14, startRow + 1, endRow, ["Faible", "Moyenne", "Elevee", "Tres elevee"]);
addValidation(sheet, 15, startRow + 1, endRow, ["P0", "P1", "P2", "P3"]);
addStatusFormatting(sheet, `O${startRow + 1}:O${endRow}`);

// Summary sheet
const summary = addSheet(
  "42_SYNTHESE_PEUPLEMENT",
  "Synthese du plan de peuplement",
  "Priorisation executive pour demarrer le peuplement du Hub en cycles de 30, 60 et 90 jours."
);
setWidths(summary, [245, 125, 340, 245, 125, 340, 245, 125, 340], 80);
summary.getRange("A4:C4").values = [["Indicateur", "Valeur", "Lecture"]];
summary.getRange("A5:C13").formulas = [
  ["=\"Feuilles a peupler\"", "=COUNTA('41_PLAN_PEUPLEMENT'!$A$5:$A$45)", "=\"Toutes les feuilles du Hub sont couvertes\""],
  ["=\"Priorite P0\"", "=COUNTIF('41_PLAN_PEUPLEMENT'!$O$5:$O$45,\"P0\")", "=\"A lancer immediatement\""],
  ["=\"Priorite P1\"", "=COUNTIF('41_PLAN_PEUPLEMENT'!$O$5:$O$45,\"P1\")", "=\"A lancer dans les 60 jours\""],
  ["=\"Donnees a collecter\"", "=COUNTIF('41_PLAN_PEUPLEMENT'!$F$5:$F$45,\"Oui\")", "=\"Necessitent enquete, extraction ou convention\""],
  ["=\"Donnees partenaires\"", "=COUNTIF('41_PLAN_PEUPLEMENT'!$E$5:$E$45,\"Oui\")", "=\"Necessitent protocole de partage\""],
  ["=\"Critiques CDP\"", "=COUNTIF('41_PLAN_PEUPLEMENT'!$G$5:$G$45,\"Oui\")", "=\"A securiser avant campagne CDP\""],
  ["=\"Critiques C40\"", "=COUNTIF('41_PLAN_PEUPLEMENT'!$H$5:$H$45,\"Oui\")", "=\"A securiser pour CAP/MER et ambition 1.5C\""],
  ["=\"Critiques CDN\"", "=COUNTIF('41_PLAN_PEUPLEMENT'!$I$5:$I$45,\"Oui\")", "=\"A valider avec alignement national\""],
  ["=\"Critiques bailleurs climat\"", "=COUNTIF('41_PLAN_PEUPLEMENT'!$J$5:$J$45,\"Oui\")", "=\"A documenter avec preuves, E&S, genre et finance\""],
];
summary.getRange("A4:C4").format = { fill: palette.teal, font: { bold: true, color: palette.white } };
summary.getRange("A5:C13").format = { fill: palette.white, wrapText: true };

summary.getRange("E4:G4").values = [["Phase", "Horizon", "Feuilles / donnees prioritaires"]];
summary.getRange("E5:G9").values = [
  ["Phase 1", "0-30 jours", "03_SOURCES_DONNEES; 05_PROJETS_CLIMAT; 08_INVENTAIRE_GES; 09_DONNEES_GES_ACTIVITE; 10_FACTEURS_EMISSION; 16_BUDGETS; 17_FINANCEMENTS; 34_REGISTRE_PREUVES_MRV"],
  ["Phase 2", "30-60 jours", "06_INDICATEURS_MRV; 07_ACTIONS_CLIMAT; 12_MOBILITE; 13_ENERGIE; 14_DECHETS; 19_QA_QC; 20_MAPPING_REPORTING; 21_RAPPORT_CDP_ICLEI"],
  ["Phase 3", "60-90 jours", "28_TRAJECTOIRES_CIBLES; 29_RISQUES_CLIMATIQUES; 30_ACTIFS_EXPOSITION; 31_PIPELINE_BANCABLE; 32_READINESS_FINANCE"],
  ["Phase 4", "90-120 jours", "33_COBENEFICES_JUSTICE; 36_PASSATION_MARCHES; 37_ACCES_ENERGIE_SEACAP; 38_INDICATEURS_INTL; 39_DASHBOARD_V2"],
  ["Continu", "Mensuel", "04_TABLEAU_BORD; 25_JOURNAL_MODIFS; 27_CONFORMITE_CADRES; 35_PLAN_DONNEES"],
];
summary.getRange("E4:G4").format = { fill: palette.teal, font: { bold: true, color: palette.white } };
summary.getRange("E5:G9").format = { fill: palette.white, wrapText: true };

summary.getRange("A16:I16").values = [["Sources internationales de cadrage"]];
summary.getRange("A16:I16").merge();
summary.getRange("A16:I16").format = { fill: palette.navy, font: { bold: true, color: palette.white } };
summary.getRange("A17:I23").values = [
  ["CDP-ICLEI Track 2026", "https://www.cdp.net/en/disclose/how-to-disclose", "Questionnaire, scoring, alignement GCoM", null, null, null, null, null, null],
  ["C40 Climate Action Planning", "https://www.c40.org/what-we-do/raising-climate-ambition/1-5c-climate-action-plans/", "CAP, MER, ambition, justice climatique", null, null, null, null, null, null],
  ["GCoM Common Reporting Framework", "https://www.globalcovenantofmayors.org/our-initiatives/data4cities/common-global-reporting-framework/", "Inventaire, cibles, risques, actions", null, null, null, null, null, null],
  ["CoM SSA / SEACAP", "https://comssa.org/en/site-resources/seacap-guidebook-extended-version", "Mitigation, adaptation, acces energie", null, null, null, null, null, null],
  ["Green Climate Fund", "https://www.greenclimate.fund/about/policies/investment", "Six criteres d'investissement et bancabilite", null, null, null, null, null, null],
  ["Banque mondiale", "https://www.worldbank.org/en/topic/urbandevelopment/brief/resilient-cities-program", "Resilience urbaine et investissement", null, null, null, null, null, null],
  ["BAD", "https://www.afdb.org/en/topics-and-sectors/sectors/climate-change/climate-change-and-green-growth-strategy", "Croissance verte, climat finance, inclusion", null, null, null, null, null, null],
];
summary.getRange("A17:I23").format = { fill: palette.white, wrapText: true };

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
  range: "42_SYNTHESE_PEUPLEMENT!A4:G13",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 7,
  maxChars: 5000,
});
console.log("SUMMARY_INSPECT");
console.log(inspect.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputFile);
console.log(`EXPORTED ${outputFile}`);

for (const sheetName of ["41_PLAN_PEUPLEMENT", "42_SYNTHESE_PEUPLEMENT"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
  console.log(`RENDERED ${sheetName}`);
}
