import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputFile = path.resolve("outputs", "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER.xlsx");
const outputDir = path.resolve("outputs");
const outputFile = path.join(outputDir, "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V2.xlsx");

const input = await FileBlob.load(inputFile);
const workbook = await SpreadsheetFile.importXlsx(input);

const palette = {
  navy: "#0B1F33",
  teal: "#0F766E",
  green: "#166534",
  amber: "#F59E0B",
  red: "#DC2626",
  lightTeal: "#E6F4F1",
  lightBlue: "#EAF2FF",
  lightGreen: "#EAF7EA",
  lightAmber: "#FFF7E6",
  lightRed: "#FEE2E2",
  white: "#FFFFFF",
  slate: "#334155",
};

const lists = {
  cadres: ["CDP-ICLEI", "C40", "ICLEI/GPC", "GCoM CRF", "CoM SSA/SEACAP", "GCF", "Banque mondiale", "BAD", "CDN Senegal", "Interne Dakar"],
  statuts: ["Non demarre", "A completer", "En cours", "A verifier", "Pret", "Soumis", "Valide"],
  priorites: ["Critique", "Haute", "Moyenne", "Basse"],
  ouiNon: ["Oui", "Non", "Partiel", "Non applicable"],
  piliers: ["Mitigation", "Adaptation", "Energie acces", "Finance", "Gouvernance", "Inclusion", "MRV"],
  risques: ["Chaleur extreme", "Inondation pluviale", "Submersion marine", "Erosion cotiere", "Secheresse", "Qualite de l'air", "Maladie vectorielle", "Rupture infrastructure"],
  groupes: ["Femmes", "Jeunes", "Personnes agees", "Personnes handicapees", "Menages faibles revenus", "Travailleurs informels", "Pecheurs", "Habitants zones inondables"],
  financeCriteria: ["Impact potentiel", "Changement de paradigme", "Developpement durable", "Besoins du beneficiaire", "Appropriation pays", "Efficience et efficacite", "Sauvegardes E&S", "Genre", "Passation marches", "Exploitation maintenance"],
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
  const existing = workbook.worksheets.getItemOrNullObject?.(name);
  if (existing && !existing.isNullObject) {
    existing.delete?.();
  }
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1:K1").merge();
  sheet.getRange("A1:K1").format = { fill: palette.navy, font: { bold: true, color: palette.white, size: 15 } };
  sheet.getRange("A1:K1").format.rowHeightPx = 34;
  if (subtitle) {
    sheet.getRange("A2").values = [[subtitle]];
    sheet.getRange("A2:K2").merge();
    sheet.getRange("A2:K2").format = { fill: palette.lightTeal, font: { italic: true, color: palette.slate }, wrapText: true };
    sheet.getRange("A2:K2").format.rowHeightPx = 42;
  }
  return sheet;
}

function setWidths(sheet, widths, rowLimit = 180) {
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
    errorAlert: { showAlert: true, style: "warning", title: "Valeur inattendue", message: "Verifier la liste de valeurs." },
  };
}

function addAlertFormatting(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.conditionalFormats.add("containsText", { text: "Critique", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "A completer", format: { fill: palette.lightAmber, font: { color: "#92400E", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "Pret", format: { fill: palette.lightGreen, font: { color: "#166534", bold: true } } });
  range.conditionalFormats.add("containsText", { text: "OK", format: { fill: palette.lightGreen, font: { color: "#166534", bold: true } } });
}

function addTableSheet({ name, title, subtitle, headers, rows, tableName, widths, validations = {}, formulas = {}, formats = {}, alertCols = [], minRows = 35 }) {
  const sheet = addSheet(name, title, subtitle);
  const headerRow = 4;
  const dataStart = 5;
  const rowCount = Math.max(minRows, rows.length + 10);
  const fullRows = Array.from({ length: rowCount }, (_, i) => headers.map((_, j) => rows[i]?.[j] ?? null));
  const endRow = headerRow + rowCount;
  const endCol = colName(headers.length);
  sheet.getRange(`A${headerRow}:${endCol}${endRow}`).values = [headers, ...fullRows];
  for (const [colIndexText, formulaFactory] of Object.entries(formulas)) {
    const colIndex = Number(colIndexText);
    const col = colName(colIndex);
    const formulaRows = [];
    for (let row = dataStart; row <= endRow; row += 1) {
      formulaRows.push([formulaFactory(row)]);
    }
    sheet.getRange(`${col}${dataStart}:${col}${endRow}`).formulas = formulaRows;
  }
  sheet.getRange(`A${headerRow}:${endCol}${headerRow}`).format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${dataStart}:${endCol}${endRow}`).format = { fill: palette.white, wrapText: true };
  const table = sheet.tables.add(`A${headerRow}:${endCol}${endRow}`, true, tableName);
  table.style = "TableStyleMedium2";
  table.showFilterButton = true;
  setWidths(sheet, widths ?? headers.map(() => 130), endRow + 5);
  sheet.freezePanes.freezeRows(headerRow);
  for (const [colIndexText, values] of Object.entries(validations)) {
    addValidation(sheet, Number(colIndexText), dataStart, endRow, values);
  }
  for (const [rangeAddress, numberFormat] of Object.entries(formats)) {
    sheet.getRange(rangeAddress.replaceAll("{end}", String(endRow))).setNumberFormat(numberFormat);
  }
  for (const colIndex of alertCols) {
    const col = colName(colIndex);
    addAlertFormatting(sheet, `${col}${dataStart}:${col}${endRow}`);
  }
  return { sheet, endRow };
}

// Light-touch V2 marker on welcome page.
const welcome = workbook.worksheets.getItem("00_ACCUEIL");
welcome.getRange("F12:J16").values = [["Version: 2.0\nDate V2: 2026-05-31\nAjouts: audit international, trajectoires, risques, bancabilite, preuves MRV, inclusion et scorecard de conformite.\nPerimetre: Ville de Dakar\nUnites GES: tCO2e"]];
welcome.getRange("F12:J16").format = { fill: palette.lightBlue, wrapText: true, font: { color: palette.slate } };

// 26_AUDIT_EXPERT
addTableSheet({
  name: "26_AUDIT_EXPERT",
  title: "Audit expert international du modele V1",
  subtitle: "Diagnostic selon les attentes C40, CDP-ICLEI, ICLEI/GPC, GCoM, CoM SSA, Fonds Vert Climat, Banque mondiale et BAD.",
  headers: ["Audit_ID", "Categorie", "Faiblesse_V1", "Risque_pour_Dakar", "Bonne_pratique_internationale", "Reponse_V2", "Priorite", "Cadres_concernes", "Statut"],
  rows: [
    ["AUD-001", "Inventaire GES", "Inventaire non detaille par code GPC, perimetre BASIC/BASIC+, gaz et incertitude", "Risque de non-comparabilite et de non-validation GCoM/CDP", "Aligner secteurs, scopes, gaz, methodes, sources et QA/QC", "Ajout scorecard conformite + bibliotheque indicateurs + registre preuves", "Critique", "CDP-ICLEI; ICLEI/GPC; GCoM", "Traite V2"],
    ["AUD-002", "Trajectoire", "Absence de BAU, cible 2030, trajectoire 1.5C et budget carbone local", "Difficile de prouver l'ambition et la coherence Paris", "Relier inventaire, scenario et cibles intermediaires", "Ajout 28_TRAJECTOIRES_CIBLES", "Critique", "C40; CDP-ICLEI; GCoM", "Traite V2"],
    ["AUD-003", "Adaptation", "Risque climatique trop general, sans exposition, vulnerabilite ni horizons", "Plan adaptation faible pour SEACAP/GCoM et finance climat", "Analyser aleas, exposition, sensibilite, capacite adaptative et groupes vulnerables", "Ajout 29_RISQUES_CLIMATIQUES et 30_ACTIFS_EXPOSITION", "Critique", "GCoM; CoM SSA; GCF; Banque mondiale; BAD", "Traite V2"],
    ["AUD-004", "Finance climat", "Budgets et financements sans bancabilite, criteres GCF, E&S, genre, O&M", "Dossiers peu competitifs pour bailleurs", "Documenter criteres investissement, cofinancement, risques, sauvegardes et resultats", "Ajout 31_PIPELINE_BANCABLE et 32_READINESS_FINANCE", "Critique", "GCF; Banque mondiale; BAD", "Traite V2"],
    ["AUD-005", "MRV", "Preuves, versions, audit trail et responsabilites insuffisants", "Faible auditabilite et risque de donnees non defendables", "Registre de preuves, controle qualite, proprietaires, validation", "Ajout 34_REGISTRE_PREUVES_MRV et 35_PLAN_DONNEES", "Haute", "CDP-ICLEI; GCoM; GCF", "Traite V2"],
    ["AUD-006", "Inclusion", "Co-benefices sociaux et justice climatique peu structures", "Benefices sous-estimes et faiblesse face aux criteres developpement durable", "Suivre genre, inclusion, sante, emplois, SDG et distribution des benefices", "Ajout 33_COBENEFICES_JUSTICE", "Haute", "C40; CDP-ICLEI; GCF; BAD", "Traite V2"],
    ["AUD-007", "Execution", "Pas de passation de marches, contrats, jalons ni readiness implementation", "Risque de sous-execution et faiblesse Banque mondiale/BAD", "Suivre procurement, contrats, conditions de decaissement, O&M", "Ajout 36_PASSATION_MARCHES", "Haute", "Banque mondiale; BAD; GCF", "Traite V2"],
  ],
  tableName: "tblAuditExpertV2",
  widths: [95, 145, 310, 280, 300, 250, 95, 220, 115],
  validations: { 7: lists.priorites },
  minRows: 30,
});

// 27_CONFORMITE_CADRES
addTableSheet({
  name: "27_CONFORMITE_CADRES",
  title: "Scorecard de conformite internationale",
  subtitle: "Notation 0-5 par exigence cle pour piloter la mise en conformite des rapports et dossiers de financement.",
  headers: ["Score_ID", "Cadre", "Dimension", "Exigence", "Source_preuve_attendue", "Score_0_5", "Poids", "Score_pondere", "Statut", "Action_corrective", "Responsable", "Echeance"],
  rows: [
    ["SC-001", "CDP-ICLEI", "Gouvernance", "Roles, responsabilites et supervision politique du climat", "15_PARTIES_PRENANTES; 25_JOURNAL_MODIFS", 3, 1, null, null, "Formaliser mandat et comite climat", "Cabinet du Maire", new Date("2026-09-30")],
    ["SC-002", "ICLEI/GPC", "Inventaire GES", "Inventaire par secteur, scope, gaz, methode, source et qualite", "08_INVENTAIRE_GES; 09_DONNEES_GES_ACTIVITE", 3, 1.5, null, null, "Ajouter codes GPC complets et incertitudes", "Referent GES", new Date("2026-10-31")],
    ["SC-003", "GCoM CRF", "Cibles", "Objectif mitigation, baseline, annee cible et progres", "28_TRAJECTOIRES_CIBLES", 2, 1.5, null, null, "Valider cible 2030 et scenario BAU", "Cellule climat", new Date("2026-11-30")],
    ["SC-004", "GCoM CRF", "Adaptation", "Evaluation risques et vulnerabilites multi-aleas", "29_RISQUES_CLIMATIQUES; 30_ACTIFS_EXPOSITION", 2, 1.5, null, null, "Completer exposition par quartier et actifs critiques", "Direction Planification", new Date("2026-12-15")],
    ["SC-005", "C40", "Plan 1.5C", "Trajectoire compatible Paris et actions prioritaires", "28_TRAJECTOIRES_CIBLES; 05_PROJETS_CLIMAT", 2, 1.5, null, null, "Construire scenario 1.5C et actions de rupture", "Cellule climat", new Date("2026-12-31")],
    ["SC-006", "CoM SSA/SEACAP", "Energie acces", "Acces a energie propre, adaptation et mitigation integres", "37_ACCES_ENERGIE_SEACAP", 1, 1.2, null, null, "Ajouter diagnostic acces energie menages/services", "Direction Energie", new Date("2026-10-31")],
    ["SC-007", "GCF", "Bancabilite", "Criteres investissement, climate rationale, E&S, genre, O&M", "31_PIPELINE_BANCABLE; 32_READINESS_FINANCE", 2, 2, null, null, "Preparer fiches concept note GCF", "Direction Finance", new Date("2026-12-31")],
    ["SC-008", "Banque mondiale", "Execution", "Jalons, procurement, risques, sauvegardes, resultats", "36_PASSATION_MARCHES; 32_READINESS_FINANCE", 2, 1.5, null, null, "Mettre a jour plan procurement et ESF", "Direction Finance", new Date("2026-11-30")],
    ["SC-009", "BAD", "Croissance verte", "Alignement climat, inclusion, genre, resilience et finance", "33_COBENEFICES_JUSTICE; 32_READINESS_FINANCE", 2, 1.5, null, null, "Quantifier co-benefices et genre", "Cellule climat", new Date("2026-11-30")],
  ],
  tableName: "tblConformiteCadres",
  widths: [95, 145, 150, 310, 230, 90, 80, 115, 115, 280, 160, 115],
  validations: { 2: lists.cadres, 9: lists.statuts },
  formulas: {
    8: (r) => `=IF(A${r}="","",F${r}*G${r})`,
    9: (r) => `=IF(A${r}="","",IF(F${r}>=4,"Pret",IF(F${r}>=3,"A verifier","A completer")))`,
  },
  formats: { "L5:L{end}": "yyyy-mm-dd" },
  alertCols: [9],
  minRows: 40,
});

// 28_TRAJECTOIRES_CIBLES
addTableSheet({
  name: "28_TRAJECTOIRES_CIBLES",
  title: "Trajectoires emissions et cibles climat",
  subtitle: "Module de pilotage BAU, cible 2030, trajectoire locale et ecart d'ambition.",
  headers: ["Annee", "Emissions_historique_tCO2e", "Scenario_BAU_tCO2e", "Trajectoire_cible_tCO2e", "Reduction_vs_baseline_pct", "Ecart_BAU_vs_cible_tCO2e", "Budget_carbone_restant_tCO2e", "Statut_ambition", "Hypothese_cle", "Source"],
  rows: [
    [2024, 1465000, 1465000, 1465000, null, null, 9000000, null, "Baseline a confirmer par inventaire complet GPC", "08_INVENTAIRE_GES"],
    [2025, 1515200, 1515200, 1390000, null, null, 7610000, null, "Premiere consolidation territoriale", "08_INVENTAIRE_GES"],
    [2026, null, 1560000, 1320000, null, null, 6290000, null, "BAU +3% annuel indicatif", "Scenario municipal"],
    [2027, null, 1607000, 1250000, null, null, 5040000, null, "Renforcement energie/dechets", "Scenario municipal"],
    [2028, null, 1655000, 1180000, null, null, 3860000, null, "Actions structurantes", "Scenario municipal"],
    [2029, null, 1705000, 1100000, null, null, 2760000, null, "Acceleration investissements", "Scenario municipal"],
    [2030, null, 1756000, 1020000, null, null, 1740000, null, "Cible 2030 a valider politiquement", "Scenario municipal"],
  ],
  tableName: "tblTrajectoiresCibles",
  widths: [80, 150, 150, 160, 140, 160, 170, 125, 290, 160],
  formulas: {
    5: (r) => `=IF(A${r}="","",IFERROR(1-D${r}/$B$5,""))`,
    6: (r) => `=IF(A${r}="","",IFERROR(C${r}-D${r},""))`,
    8: (r) => `=IF(A${r}="","",IF(F${r}>=300000,"Ecart critique",IF(F${r}>=100000,"A surveiller","OK")))`,
  },
  formats: {
    "B5:D{end}": "#,##0",
    "E5:E{end}": "0%",
    "F5:G{end}": "#,##0",
  },
  alertCols: [8],
  minRows: 20,
});

// 29_RISQUES_CLIMATIQUES
addTableSheet({
  name: "29_RISQUES_CLIMATIQUES",
  title: "Evaluation risques et vulnerabilites climatiques",
  subtitle: "Analyse alea-exposition-vulnerabilite par quartier, horizon et groupe sensible.",
  headers: ["Risk_ID", "Alea", "Quartier_zone", "Horizon", "Scenario_climatique", "Probabilite_1_5", "Impact_1_5", "Exposition_1_5", "Vulnerabilite_1_5", "Score_risque", "Niveau_risque", "Population_exposee", "Groupes_vulnerables", "Secteurs_affectes", "Mesures_adaptation", "Source_preuve", "Alerte"],
  rows: [
    ["RVA-001", "Inondation pluviale", "Medina", "2030", "SSP2/RCP4.5 indicatif", 4, 4, 4, 4, null, null, 85000, "Menages faibles revenus; Jeunes", "Mobilite; Sante; Logement", "Drainage, alerte precoce, entretien caniveaux", "Etudes hydrologiques a compiler", null],
    ["RVA-002", "Submersion marine", "Yoff", "2050", "SSP5/RCP8.5 stress-test", 3, 5, 4, 3, null, null, 45000, "Pecheurs; Menages faibles revenus", "Littoral; Economie locale", "Protection douce, repli, restauration dunes", "Cartographie littorale", null],
    ["RVA-003", "Chaleur extreme", "Dakar - Ensemble ville", "2030", "SSP2/RCP4.5 indicatif", 4, 3, 5, 4, null, null, 350000, "Personnes agees; Travailleurs informels", "Sante; Energie", "Canopée urbaine, ilots de fraicheur", "Donnees meteo nationales", null],
  ],
  tableName: "tblRisquesClimatiques",
  widths: [95, 150, 150, 90, 170, 100, 90, 105, 120, 105, 120, 130, 230, 180, 260, 190, 125],
  validations: { 2: lists.risques },
  formulas: {
    10: (r) => `=IF(A${r}="","",ROUND(AVERAGE(F${r}:I${r})*MAX(F${r}:I${r}),1))`,
    11: (r) => `=IF(A${r}="","",IF(J${r}>=18,"Critique",IF(J${r}>=12,"Eleve",IF(J${r}>=7,"Modere","Faible"))))`,
    17: (r) => `=IF(A${r}="","",IF(OR(P${r}="",L${r}=""),"A completer",IF(K${r}="Critique","Critique","OK")))`,
  },
  formats: { "L5:L{end}": "#,##0" },
  alertCols: [11, 17],
  minRows: 45,
});

// 30_ACTIFS_EXPOSITION
addTableSheet({
  name: "30_ACTIFS_EXPOSITION",
  title: "Actifs exposes et donnees geospatiales",
  subtitle: "Registre des actifs critiques pour relier projets, aleas, couts, benefices et priorisation spatiale.",
  headers: ["Asset_ID", "Nom_actif", "Type_actif", "Quartier_zone", "Latitude", "Longitude", "Proprietaire", "Service_essentiel", "Alea_principal", "Niveau_exposition", "Valeur_remplacement_XOF", "Population_servie", "Projet_associe", "Mesure_resilience", "Priorite_investissement", "Donnee_SIG_disponible"],
  rows: [
    ["AST-001", "Station pompage Medina", "Infrastructure drainage", "Medina", 14.682, -17.448, "Ville/ONAS", "Assainissement", "Inondation pluviale", "Eleve", 850000000, 120000, "PRJ-004", "Protection electrique et entretien", "Critique", "Partiel"],
    ["AST-002", "Marche Grand Dakar", "Equipement public", "Grand Dakar", 14.708, -17.454, "Ville", "Economie locale", "Chaleur extreme", "Modere", 1200000000, 65000, "", "Ombrage, eau, ventilation", "Haute", "Non"],
    ["AST-003", "Zone littorale Yoff", "Ecosysteme littoral", "Yoff", 14.758, -17.473, "Etat/Ville", "Protection cotiere", "Submersion marine", "Eleve", 2500000000, 90000, "PRJ-004", "Restauration dunes", "Critique", "Oui"],
  ],
  tableName: "tblActifsExposition",
  widths: [95, 220, 160, 150, 90, 90, 160, 160, 150, 125, 160, 130, 110, 240, 145, 145],
  validations: { 9: lists.risques, 15: lists.priorites, 16: lists.ouiNon },
  formats: { "E5:F{end}": "0.000", "K5:K{end}": "#,##0 [$XOF]", "L5:L{end}": "#,##0" },
  minRows: 45,
});

// 31_PIPELINE_BANCABLE
addTableSheet({
  name: "31_PIPELINE_BANCABLE",
  title: "Pipeline de projets bancables",
  subtitle: "Pre-filtrage finance climat pour transformer les projets en concept notes et dossiers de financement.",
  headers: ["Project_ID", "Nom_projet", "Fenetre_financement", "Instrument_prefere", "Montant_recherche_XOF", "Cofinancement_XOF", "Climate_rationale", "Additionalite", "Resultats_attendus", "TRI_economique_pct", "OPEX_annuel_XOF", "Revenus_ou_economies_XOF", "Readiness_score_0_100", "Niveau_bancabilite", "Prochaine_etape_finance"],
  rows: [
    ["PRJ-001", "Programme solaire sur batiments municipaux", "GCF / BAD / PPP", "Subvention + PPP", 1500000000, 900000000, "Reduction emissions Scope 2 et couts energie", "Acces capital et risque performance", "5 MWc solaire, 6 200 tCO2e/an evitees", 12, 45000000, 180000000, null, null, "Structurer modele ESCO et garanties"],
    ["PRJ-002", "Modernisation collecte et valorisation des dechets", "GCF / Banque mondiale / BAD", "Subvention + pret", 700000000, 350000000, "Reduction methane et salubrite urbaine", "Besoin CAPEX et structuration filiere", "Compostage, tri, emplois verts", 8, 120000000, 95000000, null, null, "Preparer etude faisabilite + E&S"],
    ["PRJ-004", "Resilience littorale et drainage urbain", "Banque mondiale / BAD / GCF adaptation", "Pret + subvention", 3200000000, 500000000, "Reduction pertes inondations et submersion", "Investissement public long terme", "Population protegee et actifs resilients", 6, 180000000, 40000000, null, null, "Preparer note concept adaptation"],
  ],
  tableName: "tblPipelineBancable",
  widths: [105, 260, 170, 150, 150, 145, 250, 230, 250, 115, 145, 165, 135, 140, 250],
  formulas: {
    13: (r) => `=IF(A${r}="","",MIN(100,ROUND(IF(E${r}>0,10,0)+IFERROR(IF(F${r}/E${r}>=0.3,15,IF(F${r}>0,8,0)),0)+IF(G${r}<>"",10,0)+IF(H${r}<>"",8,0)+IF(I${r}<>"",8,0)+IF(J${r}>=0.1,12,IF(J${r}>0,6,0))+IF(K${r}>0,5,0)+IF(L${r}>0,5,0)+IF(ISNUMBER(SEARCH("E&S",O${r})),5,0)+IF(ISNUMBER(SEARCH("concept",O${r})),7,0),0)))`,
    14: (r) => `=IF(A${r}="","",IF(M${r}>=75,"Pret concept note",IF(M${r}>=50,"A structurer","Critique")))`,
  },
  formats: { "E5:F{end}": "#,##0 [$XOF]", "J5:J{end}": "0%", "K5:L{end}": "#,##0 [$XOF]", "M5:M{end}": "0" },
  alertCols: [14],
  minRows: 40,
});

// 32_READINESS_FINANCE
addTableSheet({
  name: "32_READINESS_FINANCE",
  title: "Readiness GCF, Banque mondiale et BAD",
  subtitle: "Checklist des criteres d'investissement, sauvegardes, genre, procurement et exploitation-maintenance.",
  headers: ["Readiness_ID", "Project_ID", "Institution", "Critere", "Exigence_detaillee", "Document_preuve", "Statut", "Score_0_5", "Gap", "Action", "Responsable", "Echeance"],
  rows: [
    ["RDY-001", "PRJ-001", "GCF", "Impact potentiel", "Quantifier tCO2e evitees, beneficiaires et resultats mesurables", "Calculs MRV + fiche projet", "A verifier", 3, "Scenario baseline non valide", "Finaliser baseline energie", "Direction Energie", new Date("2026-09-30")],
    ["RDY-002", "PRJ-001", "GCF", "Changement de paradigme", "Scalabilite, replicabilite, transformation de marche", "Theory of change", "A completer", 2, "Theory of change absente", "Rediger ToC et modele replication", "Cellule climat", new Date("2026-10-31")],
    ["RDY-003", "PRJ-002", "Banque mondiale", "Sauvegardes E&S", "Identifier risques environnementaux et sociaux, mitigation et consultation", "Screening E&S", "A completer", 2, "Screening absent", "Lancer screening E&S", "Direction Dechets", new Date("2026-09-15")],
    ["RDY-004", "PRJ-004", "BAD", "Genre", "Analyse genre, benefices differencies et indicateurs", "Plan genre", "A completer", 1, "Analyse absente", "Inclure diagnostic genre/inclusion", "Cellule climat", new Date("2026-10-15")],
    ["RDY-005", "PRJ-004", "Banque mondiale", "Passation marches", "Plan procurement, lots, calendrier, risques et capacites", "Plan passation", "A verifier", 3, "Plan preliminaire", "Consolider calendrier procurement", "Direction Finance", new Date("2026-11-30")],
  ],
  tableName: "tblReadinessFinance",
  widths: [105, 105, 140, 170, 310, 190, 120, 90, 230, 250, 160, 115],
  validations: { 3: ["GCF", "Banque mondiale", "BAD"], 4: lists.financeCriteria, 7: lists.statuts },
  formats: { "L5:L{end}": "yyyy-mm-dd" },
  alertCols: [7],
  minRows: 60,
});

// 33_COBENEFICES_JUSTICE
addTableSheet({
  name: "33_COBENEFICES_JUSTICE",
  title: "Co-benefices, inclusion et justice climatique",
  subtitle: "Suivi des impacts sociaux, sanitaires, genre, emploi et SDG pour renforcer C40, CDP et finance climat.",
  headers: ["Cob_ID", "Project_ID", "Benefice", "Indicateur", "Unite", "Baseline", "Cible", "Valeur_actuelle", "Groupe_beneficiaire", "Genre_sensible", "SDG", "Methode_mesure", "Source_preuve", "Alerte"],
  rows: [
    ["COB-001", "PRJ-002", "Emplois verts", "Nombre emplois directs crees", "emplois", 0, 250, 40, "Jeunes; Femmes", "Oui", "SDG 8; SDG 11; SDG 13", "Contrats et enquete", "Registre projet", null],
    ["COB-002", "PRJ-005", "Sante publique", "Population couverte par monitoring air", "personnes", 0, 1100000, 1100000, "Tous; personnes vulnerables", "Partiel", "SDG 3; SDG 11", "Cartographie capteurs", "11_QUALITE_AIR", null],
    ["COB-003", "PRJ-004", "Resilience", "Personnes protegees contre inondations", "personnes", 0, 120000, 15000, "Menages faibles revenus", "Oui", "SDG 11; SDG 13", "Modelisation exposition", "30_ACTIFS_EXPOSITION", null],
  ],
  tableName: "tblCobeneficesJustice",
  widths: [95, 105, 150, 240, 90, 100, 100, 120, 210, 115, 160, 210, 170, 125],
  validations: { 10: lists.ouiNon },
  formulas: { 14: (r) => `=IF(A${r}="","",IF(OR(M${r}="",H${r}=""),"A completer","OK"))` },
  formats: { "F5:H{end}": "#,##0" },
  alertCols: [14],
  minRows: 45,
});

// 34_REGISTRE_PREUVES_MRV
addTableSheet({
  name: "34_REGISTRE_PREUVES_MRV",
  title: "Registre de preuves MRV et audit trail",
  subtitle: "Pieces justificatives, proprietaires, validation et versionnage des donnees utilisees dans les rapports.",
  headers: ["Evidence_ID", "Donnee_ou_indicateur", "Onglet_source", "Cellule_ou_table", "Type_preuve", "Lien_fichier_URL", "Proprietaire", "Date_reception", "Date_validation", "Validateur", "Niveau_fiabilite", "Utilise_pour_reporting", "Alerte_preuve"],
  rows: [
    ["EVD-001", "Emissions inventaire 2025", "08_INVENTAIRE_GES", "tblInventaireGES", "Tableur source", "A renseigner", "Referent GES", new Date("2026-04-30"), null, "", "Moyen", "CDP; GCoM; ICLEI", null],
    ["EVD-002", "Budget projets climat", "16_BUDGETS", "tblBudgets", "Extraction budgetaire", "A renseigner", "Direction Finance", new Date("2026-05-20"), null, "", "Eleve", "CDP; Finance climat", null],
    ["EVD-003", "Mesures qualite air", "11_QUALITE_AIR", "tblQualiteAir", "Export capteurs", "A renseigner", "Direction Environnement", new Date("2026-05-01"), null, "", "Moyen", "CDP; C40", null],
  ],
  tableName: "tblRegistrePreuvesMRV",
  widths: [105, 230, 150, 140, 140, 220, 160, 115, 115, 150, 125, 170, 130],
  validations: { 11: ["Eleve", "Moyen", "Faible", "A confirmer"] },
  formulas: { 13: (r) => `=IF(A${r}="","",IF(OR(F${r}="A renseigner",I${r}=""),"A completer","OK"))` },
  formats: { "H5:I{end}": "yyyy-mm-dd" },
  alertCols: [13],
  minRows: 60,
});

// 35_PLAN_DONNEES
addTableSheet({
  name: "35_PLAN_DONNEES",
  title: "Plan de donnees climat et lacunes",
  subtitle: "Catalogue des donnees manquantes et priorisation des actions de collecte pour un MRV robuste.",
  headers: ["DataGap_ID", "Domaine", "Donnee_manquante", "Pourquoi_critique", "Cadres_impactes", "Frequence_cible", "Responsable", "Source_potentielle", "Niveau_effort", "Priorite", "Date_cible", "Statut", "Alerte"],
  rows: [
    ["GAP-001", "GES", "Facteurs emission nationaux valides", "Condition de credibilite inventaire", "GPC; CDP; GCoM", "Annuelle", "Referent GES", "Ministere / SENELEC / IPCC", "Moyen", "Critique", new Date("2026-09-30"), "A completer", null],
    ["GAP-002", "Adaptation", "Cartographie zones inondables et actifs exposes", "Priorisation resilience et dossiers finance", "GCoM; CoM SSA; GCF; Banque mondiale", "Annuelle", "Direction Planification", "SIG national / etudes", "Eleve", "Critique", new Date("2026-12-15"), "A completer", null],
    ["GAP-003", "Finance", "Couts OPEX et revenus/economies par projet", "Bancabilite et soutenabilite", "GCF; Banque mondiale; BAD", "Trimestrielle", "Direction Finance", "Business plans projets", "Moyen", "Haute", new Date("2026-10-31"), "En cours", null],
    ["GAP-004", "Inclusion", "Benefices differencies par genre et vulnerabilite", "Critere developpement durable et justice", "C40; CDP; GCF; BAD", "Semestrielle", "Cellule climat", "Enquetes menages / ONG", "Moyen", "Haute", new Date("2026-11-30"), "A completer", null],
  ],
  tableName: "tblPlanDonnees",
  widths: [105, 120, 260, 260, 220, 130, 170, 230, 110, 100, 115, 120, 130],
  validations: { 10: lists.priorites, 12: lists.statuts },
  formulas: { 13: (r) => `=IF(A${r}="","",IF(AND(K${r}<TODAY(),L${r}<>"Pret"),"Critique",IF(L${r}<>"Pret","A completer","OK")))` },
  formats: { "K5:K{end}": "yyyy-mm-dd" },
  alertCols: [13],
  minRows: 60,
});

// 36_PASSATION_MARCHES
addTableSheet({
  name: "36_PASSATION_MARCHES",
  title: "Passation de marches, contrats et execution",
  subtitle: "Suivi des jalons d'achat, contrats, conditions de decaissement et risques d'execution.",
  headers: ["Proc_ID", "Project_ID", "Lot", "Mode_passation", "Montant_estime_XOF", "Statut_procurement", "Date_lancement", "Date_attribution_prevue", "Entreprise_attributaire", "Condition_decaissement", "Risque_execution", "Mesure_mitigation", "Alerte"],
  rows: [
    ["PRC-001", "PRJ-001", "Equipements solaires phase 2", "Appel offres", 900000000, "Preparation DAO", new Date("2026-06-15"), new Date("2026-10-30"), "", "Validation etude technique", "Retard DAO", "Appui technique externe", null],
    ["PRC-002", "PRJ-002", "Plateformes compostage pilotes", "Appel offres local", 350000000, "A lancer", new Date("2026-07-01"), new Date("2026-11-15"), "", "Foncier securise", "Foncier", "Convention quartiers", null],
    ["PRC-003", "PRJ-004", "Etudes drainage et littoral", "Consultant", 500000000, "Preparation TDR", new Date("2026-06-30"), new Date("2026-09-30"), "", "TDR valides", "Coordination institutions", "Comite technique", null],
  ],
  tableName: "tblPassationMarches",
  widths: [95, 105, 240, 150, 145, 140, 115, 135, 190, 220, 160, 220, 125],
  formulas: { 13: (r) => `=IF(A${r}="","",IF(AND(H${r}<TODAY(),F${r}<>"Attribue"),"Critique","OK"))` },
  formats: { "E5:E{end}": "#,##0 [$XOF]", "G5:H{end}": "yyyy-mm-dd" },
  alertCols: [13],
  minRows: 45,
});

// 37_ACCES_ENERGIE_SEACAP
addTableSheet({
  name: "37_ACCES_ENERGIE_SEACAP",
  title: "Acces a l'energie et SEACAP CoM SSA",
  subtitle: "Pilier specifique CoM SSA pour suivre acces a energie propre, cuisson, services publics et vulnerabilite energetique.",
  headers: ["EnergyAccess_ID", "Quartier_zone", "Service", "Population_cible", "Taux_acces_baseline", "Taux_acces_actuel", "Cible_2030", "Technologie_solution", "Project_ID", "Menages_vulnerables_cibles", "Statut", "Alerte"],
  rows: [
    ["EAC-001", "Medina", "Eclairage public securise", 180000, 0.65, 0.7, 0.95, "LED solaire / reseau", "PRJ-006", 25000, "A completer", null],
    ["EAC-002", "Dakar - Ensemble ville", "Electricite services municipaux", 1100000, 0.85, 0.88, 1, "Solaire PV + efficacite", "PRJ-001", 0, "En cours", null],
    ["EAC-003", "Grand Dakar", "Cuisson propre marches", 65000, 0.15, 0.18, 0.6, "GPL / cuisson propre", "", 15000, "A completer", null],
  ],
  tableName: "tblAccesEnergieSEACAP",
  widths: [125, 150, 210, 130, 130, 130, 105, 190, 105, 155, 120, 125],
  validations: { 11: lists.statuts },
  formulas: { 12: (r) => `=IF(A${r}="","",IF(F${r}<E${r},"Critique",IF(F${r}<G${r},"A completer","OK")))` },
  formats: { "D5:D{end}": "#,##0", "E5:G{end}": "0%", "J5:J{end}": "#,##0" },
  alertCols: [12],
  minRows: 40,
});

// 38_INDICATEURS_INTL
addTableSheet({
  name: "38_INDICATEURS_INTL",
  title: "Bibliotheque d'indicateurs internationaux",
  subtitle: "Indicateurs recommandes pour completer le reporting et le pilotage selon les meilleures pratiques.",
  headers: ["Indicator_Code", "Pilier", "Indicateur", "Unite", "Obligatoire_recommande", "Cadres", "Onglet_source_V2", "Formule_ou_methode", "Statut_dans_V2"],
  rows: [
    ["INT-GES-001", "Mitigation", "Emissions totales par secteur GPC", "tCO2e", "Obligatoire", "CDP; GCoM; ICLEI/GPC; C40", "08_INVENTAIRE_GES", "Somme par secteur/scope/gaz", "Partiel"],
    ["INT-GES-002", "Mitigation", "Reduction vs baseline", "%", "Obligatoire", "GCoM; C40; CDP", "28_TRAJECTOIRES_CIBLES", "1 - cible/baseline", "Ajoute V2"],
    ["INT-ADP-001", "Adaptation", "Population exposee par alea et quartier", "personnes", "Obligatoire", "GCoM; CoM SSA; CDP", "29_RISQUES_CLIMATIQUES", "Evaluation alea-exposition-vulnerabilite", "Ajoute V2"],
    ["INT-FIN-001", "Finance", "Besoin de financement restant", "XOF", "Recommande", "CDP; GCF; WB; BAD", "31_PIPELINE_BANCABLE", "Montant recherche - cofinancement", "Ajoute V2"],
    ["INT-JUS-001", "Inclusion", "Beneficiaires vulnerables par projet", "personnes", "Recommande", "C40; GCF; BAD", "33_COBENEFICES_JUSTICE", "Suivi par groupe beneficiaire", "Ajoute V2"],
    ["INT-MRV-001", "MRV", "Pourcentage de donnees avec preuve validee", "%", "Recommande", "CDP; GCoM; GCF", "34_REGISTRE_PREUVES_MRV", "Preuves validees / preuves requises", "Ajoute V2"],
  ],
  tableName: "tblIndicateursInternationaux",
  widths: [130, 120, 310, 90, 150, 220, 170, 280, 130],
  validations: { 2: lists.piliers },
  minRows: 60,
});

// 39_DASHBOARD_V2
{
  const sheet = addSheet("39_DASHBOARD_V2", "Tableau de bord V2 - maturite internationale", "Scorecard synthese pour suivre conformite, bancabilite, donnees manquantes et risques critiques.");
  setWidths(sheet, [210, 135, 260, 40, 210, 135, 260, 40, 210, 135, 260, 180], 80);
  sheet.getRange("A4:B7").values = [["Score conformite moyen", null], ["", null], ["Projets bancables >=75", null], ["", null]];
  sheet.getRange("E4:F7").values = [["Risques climatiques critiques", null], ["", null], ["Lacunes donnees critiques", null], ["", null]];
  sheet.getRange("I4:J7").values = [["Preuves MRV a completer", null], ["", null], ["Readiness finance moyen", null], ["", null]];
  sheet.getRange("A5").formulas = [["=AVERAGE('27_CONFORMITE_CADRES'!$F$5:$F$44)/5"]];
  sheet.getRange("A7").formulas = [["=COUNTIF('31_PIPELINE_BANCABLE'!$N$5:$N$44,\"Pret concept note\")"]];
  sheet.getRange("E5").formulas = [["=COUNTIF('29_RISQUES_CLIMATIQUES'!$K$5:$K$49,\"Critique\")"]];
  sheet.getRange("E7").formulas = [["=COUNTIF('35_PLAN_DONNEES'!$J$5:$J$64,\"Critique\")"]];
  sheet.getRange("I5").formulas = [["=COUNTIF('34_REGISTRE_PREUVES_MRV'!$M$5:$M$64,\"A completer\")"]];
  sheet.getRange("I7").formulas = [["=AVERAGE('32_READINESS_FINANCE'!$H$5:$H$64)/5"]];
  sheet.getRange("A5").setNumberFormat("0%");
  sheet.getRange("I7").setNumberFormat("0%");
  sheet.getRange("A4:J7").format = { fill: palette.white, font: { bold: true, color: palette.navy } };
  sheet.getRange("A10:L10").values = [["Cadre", "Score", "Statut", null, "Projet", "Readiness", "Bancabilite", null, "Alea", "Score risque", "Niveau", "Population exposee"]];
  sheet.getRange("A11:C19").formulas = [
    ["=IFERROR('27_CONFORMITE_CADRES'!B5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!F5/5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!I5,\"\")"],
    ["=IFERROR('27_CONFORMITE_CADRES'!B6,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!F6/5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!I6,\"\")"],
    ["=IFERROR('27_CONFORMITE_CADRES'!B7,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!F7/5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!I7,\"\")"],
    ["=IFERROR('27_CONFORMITE_CADRES'!B8,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!F8/5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!I8,\"\")"],
    ["=IFERROR('27_CONFORMITE_CADRES'!B9,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!F9/5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!I9,\"\")"],
    ["=IFERROR('27_CONFORMITE_CADRES'!B10,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!F10/5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!I10,\"\")"],
    ["=IFERROR('27_CONFORMITE_CADRES'!B11,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!F11/5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!I11,\"\")"],
    ["=IFERROR('27_CONFORMITE_CADRES'!B12,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!F12/5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!I12,\"\")"],
    ["=IFERROR('27_CONFORMITE_CADRES'!B13,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!F13/5,\"\")", "=IFERROR('27_CONFORMITE_CADRES'!I13,\"\")"],
  ];
  sheet.getRange("E11:G13").formulas = [
    ["='31_PIPELINE_BANCABLE'!B5", "='31_PIPELINE_BANCABLE'!M5/100", "='31_PIPELINE_BANCABLE'!N5"],
    ["='31_PIPELINE_BANCABLE'!B6", "='31_PIPELINE_BANCABLE'!M6/100", "='31_PIPELINE_BANCABLE'!N6"],
    ["='31_PIPELINE_BANCABLE'!B7", "='31_PIPELINE_BANCABLE'!M7/100", "='31_PIPELINE_BANCABLE'!N7"],
  ];
  sheet.getRange("I11:L13").formulas = [
    ["='29_RISQUES_CLIMATIQUES'!B5", "='29_RISQUES_CLIMATIQUES'!J5", "='29_RISQUES_CLIMATIQUES'!K5", "='29_RISQUES_CLIMATIQUES'!L5"],
    ["='29_RISQUES_CLIMATIQUES'!B6", "='29_RISQUES_CLIMATIQUES'!J6", "='29_RISQUES_CLIMATIQUES'!K6", "='29_RISQUES_CLIMATIQUES'!L6"],
    ["='29_RISQUES_CLIMATIQUES'!B7", "='29_RISQUES_CLIMATIQUES'!J7", "='29_RISQUES_CLIMATIQUES'!K7", "='29_RISQUES_CLIMATIQUES'!L7"],
  ];
  sheet.getRange("A10:L10").format = { fill: palette.teal, font: { bold: true, color: palette.white } };
  sheet.getRange("A11:L19").format = { fill: palette.white, wrapText: true };
  sheet.getRange("B11:B19").setNumberFormat("0%");
  sheet.getRange("F11:F13").setNumberFormat("0%");
  sheet.getRange("L11:L13").setNumberFormat("#,##0");
  addAlertFormatting(sheet, "C11:C19");
  addAlertFormatting(sheet, "G11:G13");
  addAlertFormatting(sheet, "K11:K13");
}

// 40_SOURCES_INTL
addTableSheet({
  name: "40_SOURCES_INTL",
  title: "Sources internationales utilisees pour la V2",
  subtitle: "Liens officiels a conserver pour audit et mise a jour methodologique.",
  headers: ["Source_ID", "Institution", "Document_ou_cadre", "URL", "Utilisation_dans_V2"],
  rows: [
    ["SRC-INT-001", "CDP-ICLEI", "CDP-ICLEI Track questionnaire and guidance", "https://www.cdp.net/en/disclose/how-to-disclose", "Reporting villes, alignement CDP/GCoM"],
    ["SRC-INT-002", "C40", "Climate Action Planning Framework", "https://www.c40.org/what-we-do/raising-climate-ambition/1-5c-climate-action-plans/", "Plan 1.5C, execution, justice climatique"],
    ["SRC-INT-003", "GCoM", "Common Reporting Framework", "https://www.globalcovenantofmayors.org/our-initiatives/data4cities/common-global-reporting-framework/", "Inventaire, cibles, risques, reporting"],
    ["SRC-INT-004", "CoM SSA", "SEACAP Guidebook and Toolbox", "https://comssa.org/en/site-resources/seacap-guidebook-extended-version", "Mitigation, adaptation, acces energie"],
    ["SRC-INT-005", "GCF", "Investment Framework", "https://www.greenclimate.fund/projects/investment-framework", "Criteres GCF et bancabilite"],
    ["SRC-INT-006", "Banque mondiale", "City Resilience Program", "https://www.worldbank.org/en/topic/urbandevelopment/brief/resilient-cities-program", "Resilience urbaine et investissements"],
    ["SRC-INT-007", "BAD", "Climate Change and Green Growth Strategy", "https://www.afdb.org/en/topics-and-sectors/sectors/climate-change/climate-change-and-green-growth-strategy", "Croissance verte, climat finance, inclusion"],
  ],
  tableName: "tblSourcesInternationales",
  widths: [120, 150, 270, 420, 260],
  minRows: 20,
});

// Verification + export
await fs.mkdir(outputDir, { recursive: true });
const dashboardInspect = await workbook.inspect({
  kind: "table",
  range: "39_DASHBOARD_V2!A1:L24",
  include: "values,formulas",
  tableMaxRows: 24,
  tableMaxCols: 12,
  maxChars: 6000,
});
console.log("DASHBOARD_V2_INSPECT");
console.log(dashboardInspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  maxChars: 6000,
});
console.log("FORMULA_ERRORS");
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputFile);
console.log(`EXPORTED ${outputFile}`);

const previewSheets = ["26_AUDIT_EXPERT", "27_CONFORMITE_CADRES", "29_RISQUES_CLIMATIQUES", "31_PIPELINE_BANCABLE", "39_DASHBOARD_V2"];
for (const sheetName of previewSheets) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
  console.log(`RENDERED ${sheetName}`);
}
