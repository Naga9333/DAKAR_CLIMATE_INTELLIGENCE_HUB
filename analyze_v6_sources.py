import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SOURCES = [
    Path(r"C:\Users\pc\Downloads\Ville de Dakar-10-04-2025-CSTAR-16-21.docx"),
    Path(r"C:\Users\pc\Downloads\CoMSSA_SEACAP_ReportingTemplate_Ville_de_Dakar (2).xlsm"),
]
OUT = Path("outputs") / "v6_source_extraction.json"


def normalize(text: str) -> str:
    text = text.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", text).strip()


def xml_text_from_zip(path: Path, prefixes: tuple[str, ...]) -> tuple[str, list[str]]:
    chunks: list[str] = []
    sheet_names: list[str] = []
    with zipfile.ZipFile(path) as zf:
      names = zf.namelist()
      if "xl/workbook.xml" in names:
          try:
              root = ET.fromstring(zf.read("xl/workbook.xml"))
              for elem in root.iter():
                  if elem.tag.endswith("}sheet"):
                      name = elem.attrib.get("name")
                      if name:
                          sheet_names.append(name)
          except Exception:
              pass
      for name in names:
          if not name.endswith(".xml"):
              continue
          if not any(name.startswith(prefix) for prefix in prefixes):
              continue
          try:
              root = ET.fromstring(zf.read(name))
          except Exception:
              continue
          for elem in root.iter():
              if elem.tag.endswith("}t") and elem.text:
                  chunks.append(elem.text)
    return normalize(" ".join(chunks)), sheet_names


def main() -> None:
    records = []
    for src in SOURCES:
        if not src.exists():
            records.append({"path": str(src), "exists": False})
            continue
        prefixes = ("word/",) if src.suffix.lower() == ".docx" else ("xl/",)
        text, sheet_names = xml_text_from_zip(src, prefixes)
        keywords = {}
        for key in [
            "CDP",
            "C40",
            "ICLEI",
            "GCoM",
            "CoM",
            "SEACAP",
            "PAAEDC",
            "GES",
            "energie",
            "dechets",
            "mobilite",
            "risque",
            "eau",
            "adaptation",
            "finance",
            "climat",
        ]:
            keywords[key] = len(re.findall(re.escape(key), text, flags=re.IGNORECASE))
        records.append(
            {
                "path": str(src),
                "file_name": src.name,
                "extension": src.suffix.lower().lstrip("."),
                "size_kb": round(src.stat().st_size / 1024, 1),
                "exists": True,
                "sheet_names": sheet_names,
                "keywords": keywords,
                "excerpt": text[:3000],
            }
        )
    OUT.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"sources": len(records), "output": str(OUT)}, ensure_ascii=False))
    for rec in records:
        print(json.dumps({"file": rec.get("file_name"), "sheets": rec.get("sheet_names", [])[:20], "keywords": rec.get("keywords")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
