import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from pypdf import PdfReader


ROOT = Path("outputs") / "DOCUMENTS"
OUT = Path("outputs") / "document_evidence_index.json"


KEYWORDS = {
    "energie": ["energie", "energy", "audit energetique", "electricite", "eclairage", "senelec", "efficacite energetique"],
    "qualite_air": ["qualite de l'air", "air quality", "pgqa", "pm2.5", "pm10", "no2", "pollution"],
    "dechets": ["dechet", "waste", "segregation", "tri", "recycl", "sonaged", "ucg"],
    "mobilite": ["mobilite", "transport", "brt", "cetud", "vehicule", "passager"],
    "adaptation": ["vulnerabilite", "resilience", "adaptation", "risque", "erosion", "cotiere", "inondation", "ccra"],
    "ges": ["ges", "ghg", "emission", "co2", "tco2", "gpc", "inventaire"],
    "finance": ["finance", "financement", "climate finance", "cfo", "bailleur", "banque", "fonds"],
    "gouvernance": ["c40", "leadership standards", "iclei", "cdp", "comssa", "seacap", "paaedc", "pcet", "pathways"],
    "reporting": ["reporting", "template", "canevas", "self assessment", "standard", "cadre"],
    "projets": ["projet", "rfp", "termes de references", "programme", "formation"],
}


SHEET_MAP = {
    "energie": ["03_Form_Electricite", "Data_Electricite", "Energie", "Inventaire_GES"],
    "qualite_air": ["05_Form_Qualite_Air", "Data_Qualite_Air", "Qualite_Air"],
    "dechets": ["04_Form_Dechets", "Data_Dechets", "Dechets"],
    "mobilite": ["07_Form_Mobilite", "Data_Mobilite", "Mobilite"],
    "adaptation": ["08_Form_Erosion_Cotiere", "Data_Erosion_Cotiere", "Risques_Climatiques"],
    "ges": ["06_Form_Emissions_GES", "Data_Emissions_GES", "Inventaire_GES"],
    "finance": ["10_Form_Finance_Climat", "Data_Finance_Climat", "Tableau_de_bord"],
    "gouvernance": ["Reporting_CDP_C40_ICLEI", "Contribution_CDN_2_0", "27_CONFORMITE_CADRES"],
    "reporting": ["Reporting_CDP_C40_ICLEI", "Contribution_CDN_2_0", "42_CONTROLE_QUALITE"],
    "projets": ["09_Form_Projets_Climat", "Data_Projets_Climat", "45_DONNEES_PRIORITAIRES"],
}


def normalize(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def xml_text_from_zip(path: Path, prefix: str) -> str:
    chunks = []
    with zipfile.ZipFile(path) as zf:
        names = [n for n in zf.namelist() if n.startswith(prefix) and n.endswith(".xml")]
        for name in names[:80]:
            try:
                root = ET.fromstring(zf.read(name))
            except Exception:
                continue
            for elem in root.iter():
                if elem.tag.endswith("}t") and elem.text:
                    chunks.append(elem.text)
    return normalize(" ".join(chunks))


def extract_pdf(path: Path) -> str:
    reader = PdfReader(str(path))
    chunks = []
    for page in reader.pages[:12]:
        try:
            chunks.append(page.extract_text() or "")
        except Exception:
            continue
    return normalize(" ".join(chunks))


def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return extract_pdf(path)
    if suffix == ".docx":
        return xml_text_from_zip(path, "word/")
    if suffix == ".pptx":
        return xml_text_from_zip(path, "ppt/slides/")
    if suffix in {".xlsm", ".xlsx"}:
        return xml_text_from_zip(path, "xl/")
    return ""


def tag_document(name: str, text: str) -> list[str]:
    haystack = f"{name} {text[:5000]}".lower()
    tags = []
    for tag, words in KEYWORDS.items():
        if any(word in haystack for word in words):
            tags.append(tag)
    if not tags:
        tags.append("documentation")
    return tags


def priority(tags: list[str], name: str) -> str:
    n = name.lower()
    if any(t in tags for t in ["ges", "energie", "dechets", "mobilite", "adaptation", "reporting"]) or any(k in n for k in ["pcet", "seacap", "c40", "pgqa", "ccra"]):
        return "P0"
    if any(t in tags for t in ["finance", "qualite_air", "projets"]):
        return "P1"
    return "P2"


def use_case(tags: list[str]) -> str:
    cases = {
        "energie": "Alimenter consommation, audits energetiques, efficacite energetique et emissions energie",
        "qualite_air": "Alimenter les indicateurs PGQA, polluants, depassements et reporting C40/CDP",
        "dechets": "Alimenter quantites collectees, tri, valorisation et projet pilote dechets",
        "mobilite": "Alimenter activite transport, trajectoires et emissions mobilite",
        "adaptation": "Alimenter vulnerabilite, risques, erosion cotiere, exposition et resilience",
        "ges": "Structurer inventaire GES, facteurs d'emission, scopes et trajectoires",
        "finance": "Documenter pipeline, besoins de financement, CFO et bailleurs climat",
        "gouvernance": "Aligner gouvernance, C40, PCET, PAAEDC/SEACAP et standards internationaux",
        "reporting": "Pre-remplir CDP/C40/ICLEI/COMSSA/CDN et canevas de reporting",
        "projets": "Identifier projets, actions, termes de reference, formations et mise en oeuvre",
    }
    return "; ".join(cases[t] for t in tags if t in cases) or "Document source a classifier"


def main() -> None:
    records = []
    for path in sorted(ROOT.glob("*")):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        try:
            text = extract_text(path)
            error = ""
        except Exception as exc:
            text = ""
            error = f"{type(exc).__name__}: {exc}"
        tags = tag_document(path.name, text)
        target_sheets = []
        for tag in tags:
            target_sheets.extend(SHEET_MAP.get(tag, []))
        records.append(
            {
                "file_name": path.name,
                "path": str(path.resolve()),
                "extension": path.suffix.lower().lstrip("."),
                "size_kb": round(path.stat().st_size / 1024, 1),
                "domains": tags,
                "priority": priority(tags, path.name),
                "target_sheets": sorted(set(target_sheets)),
                "use_case": use_case(tags),
                "excerpt": text[:1500],
                "extraction_status": "OK" if text else ("A verifier" if error else "Texte non extrait"),
                "error": error,
            }
        )
    OUT.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    by_domain = {}
    for rec in records:
        for tag in rec["domains"]:
            by_domain[tag] = by_domain.get(tag, 0) + 1
    print(json.dumps({"documents": len(records), "by_domain": by_domain, "output": str(OUT)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
