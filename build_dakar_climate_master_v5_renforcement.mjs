import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

/* =========================================================================
   DAKAR CLIMATE INTELLIGENCE HUB — MASTER VERSION
   Build V5 "Renforcement" : comble les écarts des 6 familles de données et
   ajoute les moteurs transversaux (calcul GES, scoring, vulnérabilité,
   alertes, utilisateurs/rôles, workflow de validation, modèle de données,
   indicateurs prioritaires, feuille de route).
   Source : V4B_NATURE_DASHBOARD.xlsx  →  Sortie : V5_RENFORCEMENT.xlsx
   ========================================================================= */

const inputFile = path.resolve("outputs", "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V4B_NATURE_DASHBOARD.xlsx");
const outputDir = path.resolve("outputs");
const outputFile = path.join(outputDir, "DAKAR_CLIMATE_INTELLIGENCE_HUB_MASTER_V5_RENFORCEMENT.xlsx");

const input = await FileBlob.load(inputFile);
const workbook = await SpreadsheetFile.importXlsx(input);

const palette = {
  navy: "#0B1F33", teal: "#0F766E", green: "#166534", leaf: "#2E7D32",
  amber: "#F59E0B", red: "#DC2626", blue: "#2563EB", violet: "#6D28D9",
  lightTeal: "#E6F4F1", lightGreen: "#EAF7EA", lightAmber: "#FFF7E6",
  lightRed: "#FEE2E2", lightBlue: "#E8F0FE", lightViolet: "#F1E9FB",
  white: "#FFFFFF", slate: "#334155", grey: "#F8FAFC",
};

const VALID = "Validé humainement";
const HEADER_ROW = 4;
const DATA_START = 5;
const FORM_ROWS = 60;
const DATA_END = DATA_START + FORM_ROWS - 1;

/* ----------------------------- listes contrôlées ----------------------- */
const lists = {
  years: ["2024", "2025", "2026", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "2034", "2035"],
  months: ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"],
  communes: [
    "Dakar-Plateau", "Médina", "Gueule Tapée-Fass-Colobane", "Fann-Point E-Amitié", "Grand Dakar",
    "Biscuiterie", "HLM", "Hann Bel-Air", "Sicap-Liberté", "Dieuppeul-Derklé", "Mermoz-Sacré-Coeur",
    "Ouakam", "Ngor", "Yoff", "Cambérène", "Parcelles Assainies", "Patte d'Oie", "Grand-Yoff", "Ville de Dakar",
  ],
  institutions: [
    "Ville de Dakar", "Direction Environnement", "Direction Espaces Verts", "Direction Planification",
    "Direction SIG", "Direction Mobilité", "Direction Aménagement", "Service Hygiène", "SONES / SDE", "ONAS",
    "SENELEC", "ANACIM", "CSE", "Centre de santé", "ONG / Association", "Université / Recherche", "Bailleur", "Prestataire",
  ],
  validation: ["Brouillon", "À vérifier", "Validé techniquement", VALID, "Rejeté"],
  uncertainty: ["Faible", "Moyen", "Élevé", "Très élevé"],
  priorities: ["P0", "P1", "P2", "P3"],
  frequencies: ["Mensuelle", "Trimestrielle", "Semestrielle", "Annuelle", "Ponctuelle"],
  ouinon: ["Oui", "Non", "Partiel"],
  // module-specific
  buildingTypes: ["Bâtiment administratif", "École", "Centre de santé", "Marché", "Éclairage public", "Logement social", "Équipement sportif", "Autre équipement public"],
  energySource: ["Électricité réseau", "Groupe diesel", "Solaire PV", "Gaz (GPL)", "Mixte"],
  industryBranch: ["Agroalimentaire", "Chimie", "Matériaux de construction", "Textile", "Métallurgie", "Énergie", "Autre industrie"],
  ghgScope: ["Scope 1 (direct)", "Scope 2 (énergie)", "Scope 3 (indirect)"],
  ghgSector: ["Énergie stationnaire", "Transport", "Déchets", "IPPU (industrie)", "AFOLU (agriculture)", "Éclairage public"],
  severity: ["Faible", "Modéré", "Élevé", "Extrême"],
  alertLevel: ["Vert (RAS)", "Jaune (vigilance)", "Orange (alerte)", "Rouge (alerte maximale)"],
  hazard: ["Vague de chaleur", "Inondation pluviale", "Submersion marine", "Érosion côtière", "Sécheresse", "Pollution de l'air", "Tempête / vent fort"],
  landUse: ["Résidentiel dense", "Résidentiel diffus", "Activités / commerce", "Industriel", "Équipement public", "Espace vert", "Plan d'eau", "Zone humide", "Sol nu / friche", "Voirie / imperméabilisé"],
  vulnGroup: ["Ménages à bas revenu", "Femmes cheffes de ménage", "Enfants <5 ans", "Personnes âgées >65 ans", "Personnes handicapées", "Habitat précaire / informel", "Travailleurs informels extérieurs"],
  damageType: ["Habitat", "Infrastructure routière", "Réseau eau/assainissement", "Réseau électrique", "Équipement public", "Activité économique", "Santé / vies humaines", "Agriculture urbaine"],
  sanitationType: ["Réseau d'égout", "Assainissement autonome", "Station d'épuration", "Station de pompage", "Bassin de rétention", "Vidange mécanique"],
  noiseSource: ["Trafic routier", "Marché", "Industrie", "Chantier", "Lieu de culte / événement", "Aéroport / port", "Groupe électrogène"],
  resilienceType: ["Infrastructure verte", "Infrastructure grise", "Système d'alerte précoce", "Plan d'urgence", "Sensibilisation", "Réglementation / planification", "Assurance / filet social"],
  status: ["Non démarré", "En préparation", "En cours", "Suspendu", "Terminé", "Abandonné"],
};

/* ----------------------------- helpers ---------------------------------- */
function colName(n) {
  let name = "", x = n;
  while (x > 0) { const r = (x - 1) % 26; name = String.fromCharCode(65 + r) + name; x = Math.floor((x - 1) / 26); }
  return name;
}
function q(s) { return `'${s.replaceAll("'", "''")}'`; }
function deleteIfExists(name) {
  const existing = workbook.worksheets.getItemOrNullObject?.(name);
  if (existing && !existing.isNullObject) existing.delete?.();
}
function titleSheet(sheet, title, subtitle, endCol = "L", accent = palette.navy) {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${endCol}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${endCol}1`).format = { fill: accent, font: { bold: true, color: palette.white, size: 15 } };
  sheet.getRange(`A1:${endCol}1`).format.rowHeightPx = 34;
  sheet.getRange(`A2:${endCol}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange(`A2:${endCol}2`).format = { fill: palette.lightGreen, font: { italic: true, color: palette.slate }, wrapText: true };
  sheet.getRange(`A2:${endCol}2`).format.rowHeightPx = 46;
}
function addSheet(name, title, subtitle, endCol = "L", accent = palette.navy) {
  deleteIfExists(name);
  const sheet = workbook.worksheets.add(name);
  void sheet.sheetId; // force l'attribution paresseuse du sheetId (sinon export KO)
  titleSheet(sheet, title, subtitle, endCol, accent);
  return sheet;
}
function setWidths(sheet, widths, rowLimit = 120) {
  widths.forEach((w, i) => { const c = colName(i + 1); sheet.getRange(`${c}1:${c}${rowLimit}`).format.columnWidthPx = w; });
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
  const r = sheet.getRange(rangeAddress);
  r.conditionalFormats.add("containsText", { text: "ALERTE", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  r.conditionalFormats.add("containsText", { text: "OK", format: { fill: palette.lightGreen, font: { color: "#166534", bold: true } } });
  r.conditionalFormats.add("containsText", { text: "Rouge", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  r.conditionalFormats.add("containsText", { text: "Orange", format: { fill: palette.lightAmber, font: { color: "#92400E", bold: true } } });
  r.conditionalFormats.add("containsText", { text: "Jaune", format: { fill: palette.lightAmber, font: { color: "#92400E" } } });
  r.conditionalFormats.add("containsText", { text: "Vert", format: { fill: palette.lightGreen, font: { color: "#166534" } } });
  r.conditionalFormats.add("containsText", { text: VALID, format: { fill: palette.lightGreen, font: { color: "#166534" } } });
  r.conditionalFormats.add("containsText", { text: "P0", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  r.conditionalFormats.add("containsText", { text: "Extrême", format: { fill: palette.lightRed, font: { color: "#991B1B", bold: true } } });
  r.conditionalFormats.add("containsText", { text: "Élevé", format: { fill: palette.lightAmber, font: { color: "#92400E" } } });
}
function addTable(sheet, range, name, style = "TableStyleMedium4") {
  const t = sheet.tables.add(range, true, name);
  t.style = style; t.showFilterButton = true; return t;
}

/* =========================================================================
   1) GÉNÉRATEUR GÉNÉRIQUE FORMULAIRE + BASE VALIDÉE
   Chaque module = en-têtes [ID, Année, Mois, Commune] + colonnes métier
   + tronc commun [Institution source, Document source, Statut validation,
   Niveau d'incertitude, Date de saisie, Responsable, Alerte_QC, Alimente_Rapports]
   ========================================================================= */
const FRONT = ["ID_Saisie", "Année", "Mois", "Commune / zone"];
const TAIL = ["Institution source", "Document source", "Statut validation", "Niveau d'incertitude", "Date de saisie", "Responsable", "Alerte_QC", "Alimente_Rapports"];

function buildModule(spec) {
  const headers = [...FRONT, ...spec.middle.map((m) => m.h), ...TAIL];
  const n = headers.length;
  const endCol = colName(n);
  // positions du tronc commun
  const tailStart = FRONT.length + spec.middle.length + 1; // 1-based index of first TAIL col
  const cInstitution = tailStart, cDocument = tailStart + 1, cStatut = tailStart + 2,
    cIncert = tailStart + 3, cDate = tailStart + 4, cResp = tailStart + 5, cAlerte = tailStart + 6, cFeeds = tailStart + 7;
  const widths = [110, 80, 100, 175, ...spec.middle.map((m) => m.w || 150), 165, 200, 165, 150, 120, 150, 230, 130];

  /* ---- formulaire de saisie ---- */
  const sheet = addSheet(spec.form, spec.title, spec.subtitle, endCol, spec.accent || palette.teal);
  const blank = Array.from({ length: FORM_ROWS }, () => headers.map(() => null));
  sheet.getRange(`A${HEADER_ROW}:${endCol}${DATA_END}`).values = [headers, ...blank];

  // colonnes calculées de base
  const baseFormulas = {
    1: (r) => `=IF(B${r}="","","${spec.prefix}-"&TEXT(ROW()-${HEADER_ROW},"0000"))`,
    [cDate]: (r) => `=IF(B${r}="","",TODAY())`,
    [cAlerte]: (r) => `=IF(B${r}="","",IF(OR(D${r}="",${colName(cInstitution)}${r}="",${colName(cDocument)}${r}="",${colName(cStatut)}${r}<>"${VALID}"),"ALERTE: localisation/source/validation à compléter","OK"))`,
    [cFeeds]: (r) => `=IF(B${r}="","",IF(${colName(cStatut)}${r}="${VALID}","Oui","Non"))`,
  };
  const allFormulas = { ...baseFormulas, ...(spec.computed ? spec.computed(colName) : {}) };
  for (const [ci, factory] of Object.entries(allFormulas)) {
    const col = colName(Number(ci));
    const m = [];
    for (let r = DATA_START; r <= DATA_END; r += 1) m.push([factory(r)]);
    sheet.getRange(`${col}${DATA_START}:${col}${DATA_END}`).formulas = m;
  }
  sheet.getRange(`A${HEADER_ROW}:${endCol}${HEADER_ROW}`).format = { fill: spec.accent || palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${DATA_START}:${endCol}${DATA_END}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A${HEADER_ROW}:${endCol}${DATA_END}`, `tblForm_${spec.prefix}`);
  setWidths(sheet, widths, DATA_END + 5);
  sheet.freezePanes.freezeRows(HEADER_ROW);
  // validations communes
  addValidation(sheet, 2, DATA_START, DATA_END, lists.years);
  addValidation(sheet, 3, DATA_START, DATA_END, lists.months);
  addValidation(sheet, 4, DATA_START, DATA_END, lists.communes);
  addValidation(sheet, cInstitution, DATA_START, DATA_END, lists.institutions);
  addValidation(sheet, cStatut, DATA_START, DATA_END, lists.validation);
  addValidation(sheet, cIncert, DATA_START, DATA_END, lists.uncertainty);
  // validations métier
  spec.middle.forEach((m, idx) => {
    if (m.list) addValidation(sheet, FRONT.length + 1 + idx, DATA_START, DATA_END, m.list);
    if (m.fmt) sheet.getRange(`${colName(FRONT.length + 1 + idx)}${DATA_START}:${colName(FRONT.length + 1 + idx)}${DATA_END}`).setNumberFormat(m.fmt);
  });
  sheet.getRange(`${colName(cDate)}${DATA_START}:${colName(cDate)}${DATA_END}`).setNumberFormat("yyyy-mm-dd");
  addStatusFormatting(sheet, `${colName(cStatut)}${DATA_START}:${colName(cStatut)}${DATA_END}`);
  addStatusFormatting(sheet, `${colName(cAlerte)}${DATA_START}:${colName(cAlerte)}${DATA_END}`);

  /* ---- base validée (miroir) ---- */
  const dsheet = addSheet(spec.data, `Base validée — ${spec.shortTitle}`,
    `Base consolidée alimentée automatiquement depuis ${spec.form}. Seules les lignes au statut "${VALID}" sont reprises.`,
    endCol, palette.green);
  dsheet.getRange(`A${HEADER_ROW}:${endCol}${DATA_END}`).values = [headers, ...blank];
  for (let ci = 1; ci <= n; ci += 1) {
    const col = colName(ci); const m = [];
    for (let r = DATA_START; r <= DATA_END; r += 1)
      m.push([`=IF(${q(spec.form)}!$${colName(cStatut)}${r}="${VALID}",${q(spec.form)}!${col}${r},"")`]);
    dsheet.getRange(`${col}${DATA_START}:${col}${DATA_END}`).formulas = m;
  }
  dsheet.getRange(`A${HEADER_ROW}:${endCol}${HEADER_ROW}`).format = { fill: palette.green, font: { bold: true, color: palette.white }, wrapText: true };
  dsheet.getRange(`A${DATA_START}:${endCol}${DATA_END}`).format = { fill: palette.white, wrapText: true };
  addTable(dsheet, `A${HEADER_ROW}:${endCol}${DATA_END}`, `tblData_${spec.prefix}`, "TableStyleMedium2");
  setWidths(dsheet, widths, DATA_END + 5);
  dsheet.freezePanes.freezeRows(HEADER_ROW);
  return { headers, cStatut, n };
}

/* ----------------------- DÉFINITION DES 11 MODULES ---------------------- */
const NUM = "#,##0.00";
const INT = "#,##0";
const modules = [
  {
    prefix: "BAT", form: "11_Form_Batiments", data: "Data_Batiments",
    title: "Formulaire Bâtiments & équipements publics (énergie / efficacité)",
    shortTitle: "Bâtiments & efficacité énergétique",
    subtitle: "Consommation énergétique, surfaces, éclairage public et émissions Scope 1/2 du patrimoine bâti municipal.",
    accent: palette.teal,
    middle: [
      { h: "Type de bâtiment / équipement", list: lists.buildingTypes, w: 200 },
      { h: "Source d'énergie", list: lists.energySource, w: 160 },
      { h: "Surface (m²)", fmt: INT, w: 110 },
      { h: "Conso. électricité (kWh/an)", fmt: INT, w: 150 },
      { h: "Conso. carburant (litres/an)", fmt: INT, w: 150 },
      { h: "Nb points lumineux (éclairage public)", fmt: INT, w: 150 },
      { h: "Facteur émission élec. (tCO2e/MWh)", fmt: "#,##0.000", w: 150 },
      { h: "Émissions estimées (tCO2e)", fmt: NUM, w: 150 },
      { h: "Mesure d'efficacité prévue", list: lists.ouinon, w: 150 },
    ],
    computed: (cn) => ({ 13: (r) => `=IF(OR(H${r}="",K${r}=""),"",H${r}/1000*K${r})` }), // émissions = kWh/1000 * facteur
  },
  {
    prefix: "IND", form: "12_Form_Industrie", data: "Data_Industrie",
    title: "Formulaire Industrie (procédés & émissions IPPU)",
    shortTitle: "Industrie & procédés",
    subtitle: "Émissions des procédés industriels, branches d'activité et consommations énergétiques associées.",
    accent: palette.slate,
    middle: [
      { h: "Branche industrielle", list: lists.industryBranch, w: 180 },
      { h: "Scope GES", list: lists.ghgScope, w: 150 },
      { h: "Production annuelle (unité)", fmt: NUM, w: 150 },
      { h: "Conso. énergie (MWh/an)", fmt: NUM, w: 150 },
      { h: "Conso. carburant (tep/an)", fmt: NUM, w: 150 },
      { h: "Facteur d'émission procédé", fmt: "#,##0.000", w: 150 },
      { h: "Émissions estimées (tCO2e)", fmt: NUM, w: 150 },
    ],
    computed: (cn) => ({ 11: (r) => `=IF(OR(G${r}="",J${r}=""),"",G${r}*J${r})` }),
  },
  {
    prefix: "VCH", form: "13_Form_Vagues_Chaleur", data: "Data_Vagues_Chaleur",
    title: "Formulaire Vagues de chaleur",
    shortTitle: "Vagues de chaleur",
    subtitle: "Épisodes de chaleur extrême : durée, température, population exposée, impacts sanitaires et niveau d'alerte.",
    accent: palette.amber,
    middle: [
      { h: "Date début", fmt: "yyyy-mm-dd", w: 110 },
      { h: "Durée (jours)", fmt: INT, w: 100 },
      { h: "Température max (°C)", fmt: NUM, w: 130 },
      { h: "Indice de chaleur ressentie (°C)", fmt: NUM, w: 150 },
      { h: "Population exposée", fmt: INT, w: 130 },
      { h: "Cas sanitaires liés à la chaleur", fmt: INT, w: 150 },
      { h: "Niveau d'alerte", list: lists.alertLevel, w: 160 },
      { h: "Sévérité", list: lists.severity, w: 120 },
    ],
  },
  {
    prefix: "INO", form: "14_Form_Inondations_Submersion", data: "Data_Inondations_Submersion",
    title: "Formulaire Inondations & submersion marine",
    shortTitle: "Inondations & submersion",
    subtitle: "Événements d'inondation pluviale et de submersion marine : hauteur d'eau, surface, populations et biens touchés.",
    accent: palette.blue,
    middle: [
      { h: "Type d'aléa", list: lists.hazard, w: 160 },
      { h: "Date événement", fmt: "yyyy-mm-dd", w: 110 },
      { h: "Hauteur d'eau max (cm)", fmt: INT, w: 130 },
      { h: "Surface inondée (ha)", fmt: NUM, w: 130 },
      { h: "Population affectée", fmt: INT, w: 130 },
      { h: "Habitations touchées", fmt: INT, w: 130 },
      { h: "Durée submersion (h)", fmt: INT, w: 130 },
      { h: "Niveau d'alerte", list: lists.alertLevel, w: 160 },
      { h: "Sévérité", list: lists.severity, w: 120 },
    ],
  },
  {
    prefix: "ICU", form: "15_Form_Ilots_Chaleur", data: "Data_Ilots_Chaleur",
    title: "Formulaire Îlots de chaleur urbains (ICU)",
    shortTitle: "Îlots de chaleur urbains",
    subtitle: "Cartographie thermique : écart de température ville/référence, taux de végétation et d'imperméabilisation par zone.",
    accent: palette.red,
    middle: [
      { h: "Zone / quartier", w: 180 },
      { h: "Temp. surface zone (°C)", fmt: NUM, w: 130 },
      { h: "Temp. référence péri-urbaine (°C)", fmt: NUM, w: 150 },
      { h: "Intensité ICU (Δ°C)", fmt: NUM, w: 130 },
      { h: "Taux de végétation (%)", fmt: "0%", w: 120 },
      { h: "Taux d'imperméabilisation (%)", fmt: "0%", w: 140 },
      { h: "Population exposée", fmt: INT, w: 130 },
      { h: "Priorité d'intervention", list: lists.priorities, w: 130 },
    ],
    computed: (cn) => ({ 8: (r) => `=IF(OR(F${r}="",G${r}=""),"",F${r}-G${r})` }), // intensité = surface - référence
  },
  {
    prefix: "VUL", form: "16_Form_Populations_Vulnerables", data: "Data_Populations_Vulnerables",
    title: "Formulaire Populations vulnérables",
    shortTitle: "Populations vulnérables",
    subtitle: "Caractérisation des groupes exposés : effectifs, part de la population, accès aux services et capacité d'adaptation.",
    accent: palette.violet,
    middle: [
      { h: "Groupe vulnérable", list: lists.vulnGroup, w: 200 },
      { h: "Effectif", fmt: INT, w: 120 },
      { h: "Part de la population (%)", fmt: "0.0%", w: 130 },
      { h: "Dont femmes (%)", fmt: "0.0%", w: 110 },
      { h: "Dont jeunes <15 ans (%)", fmt: "0.0%", w: 120 },
      { h: "Dont >65 ans (%)", fmt: "0.0%", w: 110 },
      { h: "Accès eau/assainissement", list: lists.ouinon, w: 150 },
      { h: "Accès santé", list: lists.ouinon, w: 120 },
      { h: "Capacité d'adaptation", list: lists.severity, w: 140 },
    ],
  },
  {
    prefix: "PED", form: "17_Form_Pertes_Dommages", data: "Data_Pertes_Dommages",
    title: "Formulaire Pertes & dommages",
    shortTitle: "Pertes & dommages",
    subtitle: "Comptabilisation des pertes économiques et humaines par aléa : biens, infrastructures, coûts de reconstruction.",
    accent: palette.red,
    middle: [
      { h: "Aléa déclencheur", list: lists.hazard, w: 160 },
      { h: "Type de dommage", list: lists.damageType, w: 180 },
      { h: "Date", fmt: "yyyy-mm-dd", w: 110 },
      { h: "Personnes affectées", fmt: INT, w: 130 },
      { h: "Décès / blessés", fmt: INT, w: 120 },
      { h: "Biens détruits (nombre)", fmt: INT, w: 130 },
      { h: "Pertes économiques (FCFA)", fmt: "#,##0", w: 160 },
      { h: "Coût reconstruction (FCFA)", fmt: "#,##0", w: 160 },
      { h: "Sévérité", list: lists.severity, w: 120 },
    ],
  },
  {
    prefix: "ASN", form: "18_Form_Assainissement", data: "Data_Assainissement",
    title: "Formulaire Eaux usées & assainissement",
    shortTitle: "Eaux usées & assainissement",
    subtitle: "Infrastructures d'assainissement, volumes traités, taux de raccordement et émissions associées (CH4/N2O).",
    accent: palette.blue,
    middle: [
      { h: "Type d'ouvrage", list: lists.sanitationType, w: 180 },
      { h: "Volume eaux usées (m³/an)", fmt: INT, w: 150 },
      { h: "Volume traité (m³/an)", fmt: INT, w: 140 },
      { h: "Taux de raccordement (%)", fmt: "0%", w: 130 },
      { h: "Population desservie", fmt: INT, w: 130 },
      { h: "Facteur émission (kgCO2e/m³)", fmt: "#,##0.000", w: 150 },
      { h: "Émissions estimées (tCO2e)", fmt: NUM, w: 150 },
    ],
    computed: (cn) => ({ 11: (r) => `=IF(OR(F${r}="",J${r}=""),"",F${r}*J${r}/1000)` }), // m3 traité * kg/m3 /1000
  },
  {
    prefix: "BRU", form: "19_Form_Pollution_Sonore", data: "Data_Pollution_Sonore",
    title: "Formulaire Pollution sonore",
    shortTitle: "Pollution sonore",
    subtitle: "Mesures acoustiques par point et source : niveau jour/nuit, dépassement des seuils OMS, population exposée.",
    accent: palette.violet,
    middle: [
      { h: "Point de mesure", w: 170 },
      { h: "Source de bruit", list: lists.noiseSource, w: 160 },
      { h: "Niveau jour Lden (dB)", fmt: NUM, w: 130 },
      { h: "Niveau nuit Lnight (dB)", fmt: NUM, w: 130 },
      { h: "Seuil OMS jour (dB)", fmt: INT, w: 120 },
      { h: "Dépassement (dB)", fmt: NUM, w: 120 },
      { h: "Population exposée", fmt: INT, w: 130 },
    ],
    computed: (cn) => ({ 9: (r) => `=IF(OR(G${r}="",I${r}=""),"",G${r}-I${r})` }), // Lden - seuil
  },
  {
    prefix: "OCS", form: "20_Form_Occupation_Sol", data: "Data_Occupation_Sol",
    title: "Formulaire Occupation du sol",
    shortTitle: "Occupation du sol",
    subtitle: "Répartition des usages du sol par zone, surfaces et évolution annuelle (artificialisation vs. végétalisation).",
    accent: palette.leaf,
    middle: [
      { h: "Classe d'occupation", list: lists.landUse, w: 180 },
      { h: "Surface (ha)", fmt: NUM, w: 120 },
      { h: "Part de la zone (%)", fmt: "0.0%", w: 120 },
      { h: "Surface année N-1 (ha)", fmt: NUM, w: 140 },
      { h: "Évolution (ha)", fmt: NUM, w: 120 },
      { h: "Surface imperméabilisée (ha)", fmt: NUM, w: 150 },
      { h: "Surface végétalisée (ha)", fmt: NUM, w: 140 },
    ],
    computed: (cn) => ({ 9: (r) => `=IF(OR(F${r}="",H${r}=""),"",F${r}-H${r})` }), // évolution = N - N-1
  },
  {
    prefix: "SOC", form: "21_Form_Socio_Economie", data: "Data_Socio_Economie",
    title: "Formulaire Données socio-économiques",
    shortTitle: "Données socio-économiques",
    subtitle: "Population, densité, ménages, genre, emplois verts, santé publique et maladies sensibles au climat par commune.",
    accent: palette.navy,
    middle: [
      { h: "Population totale", fmt: INT, w: 130 },
      { h: "Densité (hab/km²)", fmt: INT, w: 130 },
      { h: "Nombre de ménages", fmt: INT, w: 130 },
      { h: "Ménages vulnérables", fmt: INT, w: 130 },
      { h: "Part de femmes (%)", fmt: "0.0%", w: 110 },
      { h: "Part jeunes <15 ans (%)", fmt: "0.0%", w: 120 },
      { h: "Part >65 ans (%)", fmt: "0.0%", w: 110 },
      { h: "Revenu médian ménage (FCFA)", fmt: "#,##0", w: 160 },
      { h: "Emplois verts", fmt: INT, w: 110 },
      { h: "Cas maladies liées à la chaleur", fmt: INT, w: 150 },
      { h: "Cas maladies respiratoires", fmt: INT, w: 150 },
      { h: "Accès services urbains (%)", fmt: "0%", w: 130 },
    ],
  },
];

const moduleResults = modules.map((m) => ({ spec: m, ...buildModule(m) }));

/* =========================================================================
   2) MOTEUR DE CALCUL GES  (activité × facteur → émissions)
   ========================================================================= */
function buildMoteurGES() {
  const h = ["ID", "Année", "Commune", "Secteur GIEC", "Sous-catégorie", "Scope GES", "Source d'émission",
    "Donnée d'activité", "Unité activité", "Facteur d'émission", "Unité facteur",
    "Émissions brutes (tCO2e)", "Émissions évitées (tCO2e)", "Émissions nettes (tCO2e)",
    "Méthodologie", "Niveau d'incertitude", "Statut validation"];
  const endCol = colName(h.length);
  const sheet = addSheet("Moteur_Calcul_GES", "Moteur de calcul GES — activité × facteur d'émission",
    "Cœur de calcul de l'inventaire. Émissions brutes = Donnée d'activité × Facteur d'émission. Émissions nettes = brutes − évitées. Alimente les agrégations par secteur, source et commune.",
    endCol, palette.green);
  const blank = Array.from({ length: 80 }, () => h.map(() => null));
  sheet.getRange(`A${HEADER_ROW}:${endCol}${HEADER_ROW + 80}`).values = [h, ...blank];
  const fStart = DATA_START, fEnd = DATA_START + 80 - 1;
  const f = {
    1: (r) => `=IF(B${r}="","","GES-"&TEXT(ROW()-${HEADER_ROW},"0000"))`,
    12: (r) => `=IF(OR(H${r}="",J${r}=""),"",H${r}*J${r})`,
    14: (r) => `=IF(L${r}="","",L${r}-IF(M${r}="",0,M${r}))`,
  };
  for (const [ci, fac] of Object.entries(f)) {
    const col = colName(Number(ci)); const m = [];
    for (let r = fStart; r <= fEnd; r += 1) m.push([fac(r)]);
    sheet.getRange(`${col}${fStart}:${col}${fEnd}`).formulas = m;
  }
  sheet.getRange(`A${HEADER_ROW}:${endCol}${HEADER_ROW}`).format = { fill: palette.green, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${fStart}:${endCol}${fEnd}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A${HEADER_ROW}:${endCol}${fEnd}`, "tblMoteurGES");
  setWidths(sheet, [100, 80, 175, 165, 165, 150, 175, 130, 120, 130, 120, 150, 150, 150, 180, 140, 165], fEnd + 5);
  sheet.freezePanes.freezeRows(HEADER_ROW);
  addValidation(sheet, 2, fStart, fEnd, lists.years);
  addValidation(sheet, 3, fStart, fEnd, lists.communes);
  addValidation(sheet, 4, fStart, fEnd, lists.ghgSector);
  addValidation(sheet, 6, fStart, fEnd, lists.ghgScope);
  addValidation(sheet, 16, fStart, fEnd, lists.uncertainty);
  addValidation(sheet, 17, fStart, fEnd, lists.validation);
  ["H", "J", "L", "M", "N"].forEach((c) => sheet.getRange(`${c}${fStart}:${c}${fEnd}`).setNumberFormat(NUM));
  addStatusFormatting(sheet, `Q${fStart}:Q${fEnd}`);

  // bloc de synthèse par secteur
  const sumRow = fEnd + 3;
  sheet.getRange(`A${sumRow}:N${sumRow}`).merge();
  sheet.getRange(`A${sumRow}`).values = [["SYNTHÈSE PAR SECTEUR GIEC (émissions nettes, tCO2e)"]];
  sheet.getRange(`A${sumRow}:N${sumRow}`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  const sectors = lists.ghgSector;
  sheet.getRange(`A${sumRow + 1}:B${sumRow + 1}`).values = [["Secteur", "Émissions nettes (tCO2e)"]];
  sheet.getRange(`A${sumRow + 1}:B${sumRow + 1}`).format = { fill: palette.lightGreen, font: { bold: true } };
  sectors.forEach((s, i) => {
    const rr = sumRow + 2 + i;
    sheet.getRange(`A${rr}`).values = [[s]];
    sheet.getRange(`B${rr}`).formulas = [[`=SUMIF($D$${fStart}:$D$${fEnd},A${rr},$N$${fStart}:$N$${fEnd})`]];
    sheet.getRange(`B${rr}`).setNumberFormat(NUM);
  });
  const totRow = sumRow + 2 + sectors.length;
  sheet.getRange(`A${totRow}`).values = [["TOTAL"]];
  sheet.getRange(`B${totRow}`).formulas = [[`=SUM(B${sumRow + 2}:B${totRow - 1})`]];
  sheet.getRange(`A${totRow}:B${totRow}`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  sheet.getRange(`B${totRow}`).setNumberFormat(NUM);
}

/* =========================================================================
   3) SCORING CLIMAT  (atténuation / adaptation / gouvernance / finance)
   ========================================================================= */
function buildScoring() {
  const sheet = addSheet("Scoring_Climat", "Scoring climat de la Ville de Dakar",
    "Score composite (0–100) agrégeant 4 piliers pondérés. Renseigner la note brute (0–100) et la pondération ; le score pondéré et la note globale sont calculés automatiquement.",
    "H", palette.violet);
  const h = ["Pilier", "Critère évalué", "Note brute (0–100)", "Pondération (%)", "Score pondéré", "Niveau", "Source / preuve", "Commentaire"];
  const rows = [
    ["Atténuation", "Couverture de l'inventaire GES & trajectoire de réduction", 60, 0.30],
    ["Adaptation", "Évaluation vulnérabilité & mesures de résilience déployées", 55, 0.25],
    ["Gouvernance", "Plan Climat adopté, suivi-évaluation & qualité des données", 65, 0.20],
    ["Finance", "Pipeline bancable, readiness & finance climat mobilisée", 45, 0.15],
    ["Co-bénéfices", "Justice climatique, santé, emplois verts & biodiversité", 50, 0.10],
  ];
  sheet.getRange("A4:H4").values = [h];
  sheet.getRange(`A5:D${4 + rows.length}`).values = rows;
  for (let i = 0; i < rows.length; i += 1) {
    const r = 5 + i;
    sheet.getRange(`E${r}`).formulas = [[`=IF(OR(C${r}="",D${r}=""),"",C${r}*D${r})`]];
    sheet.getRange(`F${r}`).formulas = [[`=IF(C${r}="","",IF(C${r}>=75,"Avancé",IF(C${r}>=50,"Intermédiaire",IF(C${r}>=25,"Émergent","Initial"))))`]];
  }
  const gRow = 5 + rows.length + 1;
  sheet.getRange(`A${gRow}`).values = [["SCORE CLIMAT GLOBAL (0–100)"]];
  sheet.getRange(`A${gRow}:D${gRow}`).merge();
  sheet.getRange(`E${gRow}`).formulas = [[`=SUM(E5:E${4 + rows.length})`]];
  sheet.getRange(`F${gRow}`).formulas = [[`=IF(E${gRow}>=75,"Avancé",IF(E${gRow}>=50,"Intermédiaire",IF(E${gRow}>=25,"Émergent","Initial")))`]];
  sheet.getRange("A4:H4").format = { fill: palette.violet, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:H${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  sheet.getRange(`A${gRow}:H${gRow}`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  sheet.getRange(`C5:C${4 + rows.length}`).setNumberFormat("0");
  sheet.getRange(`D5:D${4 + rows.length}`).setNumberFormat("0%");
  sheet.getRange(`E5:E${gRow}`).setNumberFormat("0.0");
  addTable(sheet, `A4:H${4 + rows.length}`, "tblScoringClimat");
  setWidths(sheet, [150, 320, 120, 110, 110, 130, 200, 260], 40);
  addValidation(sheet, 3, 5, 4 + rows.length, ["0", "10", "20", "25", "30", "40", "50", "60", "70", "75", "80", "90", "100"]);
  addStatusFormatting(sheet, `F5:F${gRow}`);
}

/* =========================================================================
   4) MODULE VULNÉRABILITÉ (Exposition × Sensibilité × (1-Capacité))
   ========================================================================= */
function buildVulnerabilite() {
  const h = ["Commune / zone", "Aléa principal", "Exposition (0–1)", "Sensibilité (0–1)", "Capacité d'adaptation (0–1)",
    "Indice de vulnérabilité (0–1)", "Classe", "Population exposée", "Mesures de résilience prioritaires"];
  const endCol = colName(h.length);
  const sheet = addSheet("Module_Vulnerabilite", "Module Adaptation & vulnérabilité",
    "Indice de vulnérabilité = Exposition × Sensibilité × (1 − Capacité d'adaptation). Saisir les 3 composantes (0–1) ; l'indice et la classe sont calculés. Croiser avec 16_Form_Populations_Vulnerables.",
    endCol, palette.amber);
  const start = 5, rows = 22;
  const blank = Array.from({ length: rows }, () => h.map(() => null));
  sheet.getRange(`A4:${endCol}${4 + rows}`).values = [h, ...blank];
  // pré-remplir colonne commune avec la liste officielle
  lists.communes.forEach((c, i) => { if (i < rows) sheet.getRange(`A${start + i}`).values = [[c]]; });
  for (let i = 0; i < rows; i += 1) {
    const r = start + i;
    sheet.getRange(`F${r}`).formulas = [[`=IF(OR(C${r}="",D${r}="",E${r}=""),"",ROUND(C${r}*D${r}*(1-E${r}),3))`]];
    sheet.getRange(`G${r}`).formulas = [[`=IF(F${r}="","",IF(F${r}>=0.5,"Très élevé",IF(F${r}>=0.3,"Élevé",IF(F${r}>=0.15,"Modéré","Faible"))))`]];
  }
  sheet.getRange(`A4:${endCol}4`).format = { fill: palette.amber, font: { bold: true, color: palette.navy }, wrapText: true };
  sheet.getRange(`A5:${endCol}${4 + rows}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:${endCol}${4 + rows}`, "tblVulnerabilite");
  setWidths(sheet, [180, 160, 130, 130, 150, 150, 120, 130, 300], 40);
  addValidation(sheet, 2, start, 4 + rows, lists.hazard);
  addValidation(sheet, 9, start, 4 + rows, lists.resilienceType);
  ["C", "D", "E", "F"].forEach((c) => sheet.getRange(`${c}${start}:${c}${4 + rows}`).setNumberFormat("0.000"));
  sheet.getRange(`H${start}:H${4 + rows}`).setNumberFormat(INT);
  addStatusFormatting(sheet, `G${start}:G${4 + rows}`);
}

/* =========================================================================
   5) SYSTÈME D'ALERTE (seuils & registre)
   ========================================================================= */
function buildAlertes() {
  const sheet = addSheet("Systeme_Alerte", "Système d'alerte climatique — seuils & registre",
    "Définition des seuils de déclenchement par aléa et registre des alertes émises. Le niveau est calculé automatiquement selon la valeur observée et les seuils jaune/orange/rouge.",
    "L", palette.red);
  // bloc 1 : seuils
  sheet.getRange("A4:F4").values = [["Aléa", "Indicateur", "Seuil Jaune", "Seuil Orange", "Seuil Rouge", "Unité"]];
  const seuils = [
    ["Vague de chaleur", "Température max", 38, 41, 44, "°C"],
    ["Inondation pluviale", "Cumul pluie 24h", 50, 80, 120, "mm"],
    ["Submersion marine", "Hauteur d'eau", 30, 60, 100, "cm"],
    ["Pollution de l'air", "PM2.5 (24h)", 35, 75, 150, "µg/m³"],
    ["Pollution sonore", "Lden", 55, 65, 75, "dB"],
  ];
  sheet.getRange(`A5:F${4 + seuils.length}`).values = seuils;
  sheet.getRange("A4:F4").format = { fill: palette.red, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:F${4 + seuils.length}`).format = { fill: palette.white };
  addTable(sheet, `A4:F${4 + seuils.length}`, "tblSeuilsAlerte");

  // bloc 2 : registre des alertes
  const rStart = 4 + seuils.length + 3;
  sheet.getRange(`A${rStart}:L${rStart}`).merge();
  sheet.getRange(`A${rStart}`).values = [["REGISTRE DES ALERTES ÉMISES"]];
  sheet.getRange(`A${rStart}:L${rStart}`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  const hr = rStart + 1;
  const hh = ["ID", "Date", "Commune", "Aléa", "Indicateur", "Valeur observée", "Seuil Jaune", "Seuil Orange", "Seuil Rouge", "Niveau calculé", "Action déclenchée", "Responsable"];
  sheet.getRange(`A${hr}:L${hr}`).values = [hh];
  const dS = hr + 1, dE = hr + 30;
  const blank = Array.from({ length: 30 }, () => hh.map(() => null));
  sheet.getRange(`A${dS}:L${dE}`).values = blank;
  for (let r = dS; r <= dE; r += 1) {
    sheet.getRange(`A${r}`).formulas = [[`=IF(B${r}="","","ALR-"&TEXT(ROW()-${hr},"0000"))`]];
    sheet.getRange(`J${r}`).formulas = [[`=IF(OR(F${r}="",I${r}=""),"",IF(F${r}>=I${r},"Rouge (alerte maximale)",IF(F${r}>=H${r},"Orange (alerte)",IF(F${r}>=G${r},"Jaune (vigilance)","Vert (RAS)"))))`]];
  }
  sheet.getRange(`A${hr}:L${hr}`).format = { fill: palette.red, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${dS}:L${dE}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A${hr}:L${dE}`, "tblRegistreAlertes", "TableStyleMedium2");
  setWidths(sheet, [100, 110, 175, 160, 150, 130, 110, 120, 110, 180, 220, 150], dE + 5);
  addValidation(sheet, 3, dS, dE, lists.communes);
  addValidation(sheet, 4, dS, dE, lists.hazard);
  sheet.getRange(`B${dS}:B${dE}`).setNumberFormat("yyyy-mm-dd");
  addStatusFormatting(sheet, `J${dS}:J${dE}`);
}

/* =========================================================================
   6) UTILISATEURS & RÔLES (RBAC)
   ========================================================================= */
function buildUtilisateurs() {
  const sheet = addSheet("Utilisateurs_Roles", "Gestion des utilisateurs & rôles (RBAC)",
    "Annuaire des utilisateurs, rôles et matrice de permissions. Les droits réels sont appliqués par la plateforme cible ; cette feuille sert de référentiel de gouvernance des accès.",
    "J", palette.slate);
  // annuaire
  const h = ["ID", "Nom & prénom", "Direction / structure", "Rôle", "Périmètre (modules)", "Droit de saisie", "Droit de validation", "Droit d'export", "Statut compte", "Email"];
  const rows = [
    ["U-0001", "Administrateur Hub", "Direction SIG", "Administrateur", "Tous", "Oui", "Oui", "Oui", "Actif", ""],
    ["U-0002", "Référent GES", "Direction Environnement", "Gestionnaire données", "Atténuation / GES", "Oui", "Oui", "Oui", "Actif", ""],
    ["U-0003", "Référent Adaptation", "Direction Planification", "Gestionnaire données", "Adaptation / vulnérabilité", "Oui", "Oui", "Oui", "Actif", ""],
    ["U-0004", "Agent de saisie secteur", "Directions sectorielles", "Saisie", "Module assigné", "Oui", "Non", "Non", "Actif", ""],
    ["U-0005", "Validateur expert", "Comité technique climat", "Validateur", "Tous (lecture+validation)", "Non", "Oui", "Oui", "Actif", ""],
    ["U-0006", "Reporting international", "Cellule Climat", "Reporting", "Reporting CDP/C40/CDN", "Non", "Non", "Oui", "Actif", ""],
    ["U-0007", "Décideur / élu", "Cabinet du Maire", "Lecture", "Tableaux de bord", "Non", "Non", "Oui", "Actif", ""],
  ];
  sheet.getRange("A4:J4").values = [h];
  sheet.getRange(`A5:J${4 + rows.length}`).values = rows;
  sheet.getRange("A4:J4").format = { fill: palette.slate, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:J${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:J${4 + rows.length}`, "tblUtilisateurs");
  setWidths(sheet, [90, 180, 200, 160, 200, 110, 130, 110, 110, 200], 40);
  addValidation(sheet, 4, 5, 4 + rows.length, ["Administrateur", "Gestionnaire données", "Saisie", "Validateur", "Reporting", "Lecture"]);
  ["F", "G", "H"].forEach((c) => addValidation(sheet, c.charCodeAt(0) - 64, 5, 4 + rows.length, lists.ouinon));
  addValidation(sheet, 9, 5, 4 + rows.length, ["Actif", "Suspendu", "Désactivé"]);

  // matrice rôle x permission
  const mRow = 4 + rows.length + 3;
  sheet.getRange(`A${mRow}:J${mRow}`).merge();
  sheet.getRange(`A${mRow}`).values = [["MATRICE DES PERMISSIONS PAR RÔLE"]];
  sheet.getRange(`A${mRow}:J${mRow}`).format = { fill: palette.navy, font: { bold: true, color: palette.white } };
  const mh = ["Rôle", "Voir tableaux de bord", "Saisir données", "Valider données", "Exporter", "Gérer utilisateurs", "Configurer référentiels"];
  const mrows = [
    ["Administrateur", "✔", "✔", "✔", "✔", "✔", "✔"],
    ["Gestionnaire données", "✔", "✔", "✔", "✔", "—", "Partiel"],
    ["Saisie", "✔", "✔", "—", "—", "—", "—"],
    ["Validateur", "✔", "—", "✔", "✔", "—", "—"],
    ["Reporting", "✔", "—", "—", "✔", "—", "—"],
    ["Lecture", "✔", "—", "—", "Partiel", "—", "—"],
  ];
  sheet.getRange(`A${mRow + 1}:G${mRow + 1}`).values = [mh];
  sheet.getRange(`A${mRow + 2}:G${mRow + 1 + mrows.length}`).values = mrows;
  sheet.getRange(`A${mRow + 1}:G${mRow + 1}`).format = { fill: palette.slate, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A${mRow + 2}:G${mRow + 1 + mrows.length}`).format = { fill: palette.white };
  addTable(sheet, `A${mRow + 1}:G${mRow + 1 + mrows.length}`, "tblMatricePermissions", "TableStyleMedium2");
}

/* =========================================================================
   7) WORKFLOW DE VALIDATION DES DONNÉES
   ========================================================================= */
function buildWorkflow() {
  const sheet = addSheet("Workflow_Validation", "Workflow de validation des données",
    "Cycle de vie d'une donnée du brouillon à la publication. Chaque étape, son acteur (RACI), son livrable et son contrôle qualité. Le statut « Validé humainement » est la seule porte vers les bases consolidées et les rapports.",
    "H", palette.teal);
  const h = ["Étape", "Statut associé", "Acteur responsable (R)", "Approbateur (A)", "Contrôle qualité", "Livrable / sortie", "Délai cible", "Blocage si échec"];
  const rows = [
    ["1. Saisie", "Brouillon", "Agent de saisie (direction sectorielle)", "—", "Champs obligatoires renseignés, format respecté", "Ligne de formulaire", "Continu", "Non"],
    ["2. Pré-contrôle automatique", "À vérifier", "Système (formules QC)", "—", "Alerte_QC = OK, validations de liste respectées", "Indicateur Alerte_QC", "Immédiat", "Oui si ALERTE"],
    ["3. Vérification technique", "Validé techniquement", "Gestionnaire de données du module", "Référent thématique", "Cohérence valeurs, source documentée, unité", "Donnée vérifiée", "48 h", "Oui"],
    ["4. Validation humaine experte", VALID, "Validateur expert / comité technique", "Coordonnateur Climat", "Preuve à l'appui, méthodologie tracée, incertitude qualifiée", "Donnée publiable", "5 jours", "Oui"],
    ["5. Consolidation", "Publié (base Data_)", "Système (formules de recopie)", "—", "Recopie auto des lignes validées uniquement", "Base consolidée", "Immédiat", "—"],
    ["6. Reporting", "Reporté", "Cellule reporting international", "Coordonnateur Climat", "Mapping vers CDP/C40/CDN, contrôle d'agrégation", "Rapport / export", "Selon échéance", "—"],
    ["7. Archivage & audit", "Archivé", "Direction SIG", "Auditeur", "Journal des modifications, traçabilité", "Registre de preuves", "Annuel", "—"],
  ];
  sheet.getRange("A4:H4").values = [h];
  sheet.getRange(`A5:H${4 + rows.length}`).values = rows;
  sheet.getRange("A4:H4").format = { fill: palette.teal, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:H${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:H${4 + rows.length}`, "tblWorkflowValidation");
  setWidths(sheet, [200, 180, 250, 200, 280, 200, 110, 130], 30);
  addStatusFormatting(sheet, `B5:B${4 + rows.length}`);
}

/* =========================================================================
   8) INDICATEURS PRIORITAIRES (mission 5)
   ========================================================================= */
function buildIndicateursPrioritaires() {
  const sheet = addSheet("Indicateurs_Prioritaires", "Indicateurs climatiques prioritaires",
    "Jeu d'indicateurs P0/P1 alignés CDP, C40, GCoM, CDN et ODD. Renseigner la valeur de référence, la cible et la valeur actuelle ; l'avancement vers la cible est calculé.",
    "L", palette.blue);
  const h = ["Code", "Indicateur", "Famille", "Unité", "Référence (année base)", "Cible", "Valeur actuelle", "Avancement (%)", "Priorité", "Cadre de reporting", "Fréquence", "Source"];
  const rows = [
    ["MIT-01", "Émissions GES totales du territoire", "Atténuation", "tCO2e/an", 0, 0, 0, null, "P0", "CDP / GCoM / CDN", "Annuelle", "Moteur_Calcul_GES"],
    ["MIT-02", "Émissions GES par habitant", "Atténuation", "tCO2e/hab", 0, 0, 0, null, "P0", "C40 / GCoM", "Annuelle", "Moteur_Calcul_GES + Socio"],
    ["MIT-03", "Émissions évitées (actions mises en œuvre)", "Atténuation", "tCO2e/an", 0, 0, 0, null, "P1", "CDP / CDN", "Annuelle", "Moteur_Calcul_GES"],
    ["MIT-04", "Part d'énergie renouvelable (patrimoine ville)", "Atténuation", "%", 0, 0, 0, null, "P1", "GCoM", "Annuelle", "Data_Batiments / Energie"],
    ["ADP-01", "Population exposée aux inondations", "Adaptation", "habitants", 0, 0, 0, null, "P0", "CDP / GCoM", "Annuelle", "Data_Inondations_Submersion"],
    ["ADP-02", "Indice de vulnérabilité moyen", "Adaptation", "0–1", 0, 0, 0, null, "P0", "GCoM / CDN", "Annuelle", "Module_Vulnerabilite"],
    ["ADP-03", "Jours de vague de chaleur / an", "Adaptation", "jours", 0, 0, 0, null, "P1", "C40 / GCoM", "Annuelle", "Data_Vagues_Chaleur"],
    ["ENV-01", "Surface d'espaces verts par habitant", "Environnement", "m²/hab", 0, 0, 0, null, "P1", "ODD 11.7", "Annuelle", "Data_Nature_Espaces_Verts"],
    ["ENV-02", "Taux de couverture végétale", "Environnement", "%", 0, 0, 0, null, "P1", "ODD 11 / 13", "Annuelle", "Data_Occupation_Sol"],
    ["ENV-03", "Taux de déchets collectés / recyclés", "Environnement", "%", 0, 0, 0, null, "P1", "CDP / ODD 11.6", "Annuelle", "Dechets"],
    ["ENV-04", "Population exposée à PM2.5 > seuil OMS", "Environnement", "habitants", 0, 0, 0, null, "P1", "ODD 11.6 / 3.9", "Trimestrielle", "Qualite_Air"],
    ["SOC-01", "Ménages vulnérables couverts par une mesure", "Socio-éco", "%", 0, 0, 0, null, "P1", "CDN / ODD 1 / 13", "Annuelle", "Data_Populations_Vulnerables"],
    ["SOC-02", "Emplois verts créés", "Socio-éco", "emplois", 0, 0, 0, null, "P2", "CDN / ODD 8", "Annuelle", "Data_Socio_Economie"],
    ["GOV-01", "Taux d'exécution physique du Plan Climat", "Gouvernance", "%", 0, 0, 0, null, "P0", "Plan Climat Dakar", "Trimestrielle", "Suivi_Plan_Climat"],
    ["FIN-01", "Finance climat mobilisée", "Finance", "FCFA", 0, 0, 0, null, "P0", "CDP / CDN", "Annuelle", "33_READINESS_FINANCE"],
  ];
  sheet.getRange("A4:L4").values = [h];
  sheet.getRange(`A5:L${4 + rows.length}`).values = rows;
  for (let i = 0; i < rows.length; i += 1) {
    const r = 5 + i;
    sheet.getRange(`H${r}`).formulas = [[`=IF(OR(F${r}="",G${r}="",F${r}=E${r}),"",MAX(0,MIN(1,(G${r}-E${r})/(F${r}-E${r}))))`]];
  }
  sheet.getRange("A4:L4").format = { fill: palette.blue, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:L${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:L${4 + rows.length}`, "tblIndicateursPrioritaires");
  setWidths(sheet, [80, 280, 130, 100, 130, 90, 110, 110, 90, 170, 120, 200], 40);
  sheet.getRange(`H5:H${4 + rows.length}`).setNumberFormat("0%");
  addValidation(sheet, 9, 5, 4 + rows.length, lists.priorities);
  addValidation(sheet, 11, 5, 4 + rows.length, lists.frequencies);
  addStatusFormatting(sheet, `I5:I${4 + rows.length}`);
}

/* =========================================================================
   9) SUIVI DU PLAN CLIMAT (objectifs, activités, exécution)
   ========================================================================= */
function buildSuiviPlan() {
  const h = ["ID", "Axe / objectif", "Activité", "Responsable", "Budget prévu (FCFA)", "Budget engagé (FCFA)",
    "Avancement physique (%)", "Taux exécution financière (%)", "État", "Résultat attendu", "Résultat atteint",
    "Indicateur de performance", "Risque", "Mesure corrective", "Échéance", "Statut validation"];
  const endCol = colName(h.length);
  const sheet = addSheet("Suivi_Plan_Climat", "Suivi-évaluation du Plan Climat de Dakar",
    "Tableau de bord opérationnel : objectifs, activités, responsables, budgets et taux d'exécution physique/financière. Le taux d'exécution financière est calculé (engagé / prévu).",
    endCol, palette.green);
  const rows = 40;
  const blank = Array.from({ length: rows }, () => h.map(() => null));
  sheet.getRange(`A4:${endCol}${4 + rows}`).values = [h, ...blank];
  for (let i = 0; i < rows; i += 1) {
    const r = 5 + i;
    sheet.getRange(`A${r}`).formulas = [[`=IF(B${r}="","","PC-"&TEXT(ROW()-4,"0000"))`]];
    sheet.getRange(`H${r}`).formulas = [[`=IF(OR(E${r}="",F${r}="",E${r}=0),"",F${r}/E${r})`]];
  }
  sheet.getRange(`A4:${endCol}4`).format = { fill: palette.green, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:${endCol}${4 + rows}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:${endCol}${4 + rows}`, "tblSuiviPlanClimat");
  setWidths(sheet, [90, 220, 240, 170, 150, 150, 130, 140, 130, 220, 220, 200, 200, 220, 110, 165], rows + 8);
  sheet.freezePanes.freezeRows(4);
  addValidation(sheet, 9, 5, 4 + rows, lists.status);
  addValidation(sheet, 16, 5, 4 + rows, lists.validation);
  ["E", "F"].forEach((c) => sheet.getRange(`${c}5:${c}${4 + rows}`).setNumberFormat("#,##0"));
  ["G", "H"].forEach((c) => sheet.getRange(`${c}5:${c}${4 + rows}`).setNumberFormat("0%"));
  addStatusFormatting(sheet, `I5:I${4 + rows}`);
  addStatusFormatting(sheet, `P5:P${4 + rows}`);
}

/* =========================================================================
   10) ARCHITECTURE, MODÈLE DE DONNÉES, FEUILLE DE ROUTE (missions 1-4,7)
   ========================================================================= */
function textSheet(name, title, subtitle, blocks, accent) {
  const sheet = addSheet(name, title, subtitle, "H", accent || palette.navy);
  let r = 4;
  for (const b of blocks) {
    sheet.getRange(`A${r}:H${r}`).merge();
    sheet.getRange(`A${r}`).values = [[b.h]];
    sheet.getRange(`A${r}:H${r}`).format = { fill: accent || palette.navy, font: { bold: true, color: palette.white, size: 12 } };
    r += 1;
    for (const line of b.lines) {
      sheet.getRange(`A${r}:H${r}`).merge();
      sheet.getRange(`A${r}`).values = [[line]];
      sheet.getRange(`A${r}:H${r}`).format = { fill: palette.white, font: { color: palette.slate }, wrapText: true };
      sheet.getRange(`A${r}:H${r}`).format.rowHeightPx = 30;
      r += 1;
    }
    r += 1;
  }
  setWidths(sheet, [200, 120, 120, 120, 120, 120, 120, 120], r + 5);
  return sheet;
}

function buildArchitecture() {
  textSheet("00_ARCHITECTURE", "Dakar Climate Intelligence Hub — Architecture (V5)",
    "Vision cible et architecture en couches de la plateforme d'intelligence climatique urbaine.",
    [
      { h: "1. Vision", lines: [
        "Plateforme intégrée de pilotage climat de la Ville de Dakar : collecte, qualité, calcul, analyse, alerte et reporting international, sur une source de données unique et tracée.",
        "Couvre 6 familles : (1) Atténuation/GES, (2) Adaptation, (3) Environnement urbain, (4) Socio-économie, (5) Suivi du Plan Climat, (6) Reporting international.",
      ]},
      { h: "2. Architecture en couches", lines: [
        "Couche 1 — SAISIE : formulaires sectoriels contrôlés (listes, formats, ID auto, contrôle qualité Alerte_QC).",
        "Couche 2 — QUALITÉ & VALIDATION : workflow Brouillon → Vérifié → Validé humainement ; seules les lignes validées passent en base.",
        "Couche 3 — DONNÉES CONSOLIDÉES : bases Data_* en lecture seule, alimentées par formules depuis les formulaires.",
        "Couche 4 — MOTEURS : calcul GES (activité × facteur), vulnérabilité, scoring climat, système d'alerte.",
        "Couche 5 — RESTITUTION : tableaux de bord dynamiques, graphiques automatiques, cartes SIG, indicateurs prioritaires.",
        "Couche 6 — REPORTING : mappings CDP, C40, GCoM/Convention des Maires, CDN Sénégal, ODD, MRV, finance climat.",
        "Transversal — GOUVERNANCE : utilisateurs & rôles (RBAC), journal des modifications, registre de preuves, référentiels.",
      ]},
      { h: "3. Principes de conception", lines: [
        "Source unique de vérité : aucune donnée n'est recopiée manuellement entre feuilles.",
        "Séparation saisie / validé / publié : la donnée brute ne contamine jamais les rapports.",
        "Traçabilité totale : chaque valeur porte sa source, sa méthodologie et son niveau d'incertitude.",
        "Référentiels partagés : communes, institutions, secteurs GIEC, facteurs d'émission centralisés.",
        "Évolutivité : le classeur Excel V5 est le MVP ; l'architecture se transpose vers une base relationnelle + web.",
      ]},
    ], palette.navy);
}

function buildModeleDonnees() {
  const sheet = addSheet("Modele_Donnees", "Modèle de données global — tables principales",
    "Dictionnaire des entités de la plateforme : table, clé, famille, alimentée par, alimente. Sert de schéma de référence pour la future base relationnelle.",
    "F", palette.blue);
  const h = ["Table / entité", "Clé primaire", "Famille de données", "Alimentée par", "Alimente", "Granularité"];
  const rows = [
    ["Data_Batiments", "ID_Saisie (BAT)", "1. Atténuation", "11_Form_Batiments", "Moteur_Calcul_GES, Dashboards", "Bâtiment × année"],
    ["Data_Industrie", "ID_Saisie (IND)", "1. Atténuation", "12_Form_Industrie", "Moteur_Calcul_GES", "Site × année"],
    ["Moteur_Calcul_GES", "ID (GES)", "1. Atténuation", "Données d'activité + Facteurs", "Inventaire, Indicateurs, Reporting", "Source × commune × année"],
    ["Data_Vagues_Chaleur", "ID_Saisie (VCH)", "2. Adaptation", "13_Form_Vagues_Chaleur", "Systeme_Alerte, Vulnérabilité", "Épisode"],
    ["Data_Inondations_Submersion", "ID_Saisie (INO)", "2. Adaptation", "14_Form_Inondations_Submersion", "Systeme_Alerte, Pertes&Dommages", "Événement × zone"],
    ["Data_Ilots_Chaleur", "ID_Saisie (ICU)", "2. Adaptation", "15_Form_Ilots_Chaleur", "Vulnérabilité, SIG", "Zone × année"],
    ["Data_Populations_Vulnerables", "ID_Saisie (VUL)", "2./4.", "16_Form_Populations_Vulnerables", "Module_Vulnerabilite, Indicateurs", "Groupe × commune"],
    ["Data_Pertes_Dommages", "ID_Saisie (PED)", "2. Adaptation", "17_Form_Pertes_Dommages", "Finance, Reporting", "Événement"],
    ["Data_Assainissement", "ID_Saisie (ASN)", "3. Environnement", "18_Form_Assainissement", "Moteur_Calcul_GES, Indicateurs", "Ouvrage × année"],
    ["Data_Pollution_Sonore", "ID_Saisie (BRU)", "3. Environnement", "19_Form_Pollution_Sonore", "Systeme_Alerte, Indicateurs", "Point × campagne"],
    ["Data_Occupation_Sol", "ID_Saisie (OCS)", "3. Environnement", "20_Form_Occupation_Sol", "Indicateurs ENV, SIG", "Classe × zone × année"],
    ["Data_Socio_Economie", "ID_Saisie (SOC)", "4. Socio-éco", "21_Form_Socio_Economie", "Vulnérabilité, Indicateurs/hab", "Commune × année"],
    ["Suivi_Plan_Climat", "ID (PC)", "5. Suivi", "Saisie directe", "Indicateurs GOV, Reporting", "Activité"],
    ["Indicateurs_Prioritaires", "Code", "Transversal", "Toutes les bases Data_", "Tableaux de bord, Reporting", "Indicateur × année"],
    ["Module_Vulnerabilite", "Commune+Aléa", "Transversal", "Bases adaptation & socio", "Scoring, Dashboards, SIG", "Commune × aléa"],
    ["Utilisateurs_Roles", "ID utilisateur", "Gouvernance", "Saisie admin", "Contrôle d'accès (plateforme)", "Utilisateur"],
    ["Referentiels (communes, facteurs)", "Code référentiel", "Gouvernance", "Saisie admin", "Toutes les feuilles (listes)", "Item"],
  ];
  sheet.getRange("A4:F4").values = [h];
  sheet.getRange(`A5:F${4 + rows.length}`).values = rows;
  sheet.getRange("A4:F4").format = { fill: palette.blue, font: { bold: true, color: palette.white }, wrapText: true };
  sheet.getRange(`A5:F${4 + rows.length}`).format = { fill: palette.white, wrapText: true };
  addTable(sheet, `A4:F${4 + rows.length}`, "tblModeleDonnees");
  setWidths(sheet, [240, 150, 150, 230, 240, 180], rows.length + 8);
}

function buildFeuilleRoute() {
  textSheet("Feuille_Route", "Feuille de route — du MVP vers la plateforme avancée",
    "Trajectoire de montée en puissance vers une plateforme d'intelligence climatique urbaine.",
    [
      { h: "Phase 1 — MVP (classeur V5, actuel)", lines: [
        "Classeur Excel master multi-feuilles : formulaires contrôlés, bases validées, moteur GES, scoring, vulnérabilité, alertes, suivi du Plan Climat, reporting.",
        "Gouvernance des données : workflow de validation + rôles + référentiels. Déployable immédiatement sur poste / partage de fichiers.",
      ]},
      { h: "Phase 2 — Consolidation & SIG", lines: [
        "Connexion cartographique (QGIS / ArcGIS / Leaflet) sur les colonnes commune & zone : cartes des îlots de chaleur, inondations, vulnérabilité.",
        "Import automatisé Excel/CSV des sources externes (ANACIM, SENELEC, ONAS, CSE). Tableaux de bord Power BI / Looker en lecture.",
      ]},
      { h: "Phase 3 — Plateforme web & base de données", lines: [
        "Migration du modèle de données vers une base relationnelle (PostgreSQL/PostGIS). API de saisie et de restitution.",
        "Application web multi-utilisateurs avec RBAC réel, workflow de validation en ligne, et journal d'audit.",
        "Moteur GES paramétrable, bibliothèque de facteurs versionnée, calcul d'incertitude.",
      ]},
      { h: "Phase 4 — Intelligence climatique", lines: [
        "Capteurs IoT (qualité air, température, niveau d'eau) en temps réel alimentant le système d'alerte.",
        "Modèles prédictifs (vagues de chaleur, inondations), scénarios de trajectoire GES, aide à la décision.",
        "Reporting automatisé CDP/C40/CDN/GCoM, MRV vérifiable, et plateforme ouverte de données climat.",
      ]},
    ], palette.teal);
}

/* ----------------------------- exécution -------------------------------- */
buildMoteurGES();
buildScoring();
buildVulnerabilite();
buildAlertes();
buildUtilisateurs();
buildWorkflow();
buildIndicateursPrioritaires();
buildSuiviPlan();
buildArchitecture();
buildModeleDonnees();
buildFeuilleRoute();

/* ----------------------------- sortie & contrôle ------------------------ */
await fs.mkdir(outputDir, { recursive: true });

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  maxChars: 4000,
});
console.log("FORMULA_ERRORS");
console.log(formulaErrors.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputFile);

const outputBlob = await FileBlob.load(outputFile);
const verifyWorkbook = await SpreadsheetFile.importXlsx(outputBlob);
const verifyErrors = await verifyWorkbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  maxChars: 4000,
});
console.log("VERIFY_FORMULA_ERRORS");
console.log(verifyErrors.ndjson);

const info = await verifyWorkbook.inspect({ kind: "workbook", include: "sheets", maxChars: 9000 });
console.log("VERIFY_WORKBOOK");
console.log(info.ndjson);
console.log("OUTPUT_FILE", outputFile);
