# -*- coding: utf-8 -*-
"""
Intègre une table des matières renseignée (entrées + numéros de page réels)
dans le rapport V10. Les numéros sont mesurés sur le rendu PDF (LibreOffice),
en itérant jusqu'à stabilisation de la pagination. Le champ TOC reste
actualisable dans Word (clic droit > Mettre à jour les champs).
"""
import subprocess, sys, re
import docx
import fitz
from docx.shared import Pt, Cm, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DOCX = "/tmp/doc/RAPPORT_INSTITUTIONNEL_JOURNEES_SCIENTIFIQUES_APEC_2026_PROFESSIONNEL_V10.docx"
PDF = DOCX.replace(".docx", ".pdf")
BLEU_FONCE = "1F4E79"

def render_pdf():
    subprocess.run(
        ["soffice", "-env:UserInstallation=file:///tmp/loprofile_tdm",
         "--headless", "--convert-to", "pdf", DOCX, "--outdir", "/tmp/doc"],
        env={"HOME": "/tmp/doc", "PATH": "/usr/bin:/bin"},
        check=True, capture_output=True, timeout=300)

def expected_headings(d):
    out = []
    for p in d.paragraphs:
        if p.style.name in ("Heading 1", "Heading 2", "Heading 3") and p.text.strip():
            out.append((int(p.style.name[-1]), re.sub(r"\s+", " ", p.text.strip())))
    return out

def detect_pages(headings):
    """Associe chaque titre attendu à sa page dans le PDF (spans gras >= 11,3 pt)."""
    doc = fitz.open(PDF)
    spans = []  # (page, texte)
    for pno in range(len(doc)):
        for block in doc[pno].get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                for sp in line["spans"]:
                    bold = sp["flags"] & 16
                    st = re.sub(r"\s+", " ", sp["text"]).strip()
                    if bold and sp["size"] >= 11.3 and st:
                        spans.append((pno + 1, st))
    pages, si = [], 0
    for lvl, text in headings:
        found = None
        for j in range(si, len(spans)):
            pg, st = spans[j]
            if text.startswith(st[:25]) or st.startswith(text[:25]):
                found, si = pg, j + 1
                break
        if found is None:
            raise RuntimeError(f"Titre introuvable dans le PDF : {text[:60]}")
        pages.append(found)
    return pages

def ensure_toc_styles(d):
    base_specs = {
        "TOC 1": dict(size=11, bold=True, indent=0, before=6, color=BLEU_FONCE),
        "TOC 2": dict(size=10.5, bold=False, indent=0.5, before=2, color=None),
        "TOC 3": dict(size=10, bold=False, indent=1.0, before=0, color=None),
    }
    for name, spec in base_specs.items():
        try:
            st = d.styles[name]
        except KeyError:
            st = d.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=True)
        st.base_style = d.styles["Normal"]
        st.font.name = "Calibri"
        st.font.size = Pt(spec["size"])
        st.font.bold = spec["bold"]
        if spec["color"]:
            st.font.color.rgb = RGBColor.from_string(spec["color"])
        pf = st.paragraph_format
        pf.left_indent = Cm(spec["indent"])
        pf.space_before = Pt(spec["before"])
        pf.space_after = Pt(2)
        pf.line_spacing = 1.0
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        # taquet droit avec points de conduite en bord de justification
        for ts in list(pf.tab_stops):
            del pf.tab_stops[0]
        pf.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

def make_run(text=None, fld=None, instr=None, tab=False):
    r = OxmlElement("w:r")
    if fld:
        fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), fld); r.append(fc)
    if instr is not None:
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
        it.text = instr; r.append(it)
    if tab:
        r.append(OxmlElement("w:tab"))
    if text is not None:
        t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve")
        t.text = text; r.append(t)
    return r

def entry_p(style_id, text, page, first=False, last=False):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    ps = OxmlElement("w:pStyle"); ps.set(qn("w:val"), style_id); ppr.append(ps)
    p.append(ppr)
    if first:
        p.append(make_run(fld="begin"))
        p.append(make_run(instr=' TOC \\o "1-3" \\h \\z \\u '))
        p.append(make_run(fld="separate"))
    p.append(make_run(text=text))
    p.append(make_run(tab=True))
    p.append(make_run(text=str(page)))
    if last:
        p.append(make_run(fld="end"))
    return p

def find_toc_paragraphs(body):
    """Retourne la liste des w:p formant le champ TDM (fldSimple ou fldChar)."""
    P = qn("w:p")
    for el in body:
        if el.tag != P:
            continue
        fs = el.findall(".//" + qn("w:fldSimple"))
        if any("TOC" in (f.get(qn("w:instr")) or "") for f in fs):
            return [el]
        instrs = el.findall(".//" + qn("w:instrText"))
        if any("TOC" in (i.text or "") for i in instrs):
            if any(fc.get(qn("w:fldCharType")) == "end"
                   for fc in el.findall(".//" + qn("w:fldChar"))):
                return [el]
            group = [el]
            sib = el.getnext()
            while sib is not None:
                group.append(sib)
                if any(fc.get(qn("w:fldCharType")) == "end"
                       for fc in sib.findall(".//" + qn("w:fldChar"))):
                    return group
                sib = sib.getnext()
            return group
    return []

def write_toc(d, headings, pages):
    ensure_toc_styles(d)
    body = d.element.body
    old = find_toc_paragraphs(body)
    if not old:
        raise RuntimeError("Champ TDM introuvable")
    anchor = old[0]
    n = len(headings)
    new_ps = []
    for k, ((lvl, text), pg) in enumerate(zip(headings, pages)):
        new_ps.append(entry_p(f"TOC{lvl}", text, pg,
                              first=(k == 0), last=(k == n - 1)))
    for np in new_ps:
        anchor.addprevious(np)
    for el in old:
        body.remove(el)

prev_pages = None
for it in range(1, 5):
    d = docx.Document(DOCX)
    headings = expected_headings(d)
    render_pdf()
    pages = detect_pages(headings)
    print(f"itération {it} : {len(headings)} titres, "
          f"pages {pages[0]}..{pages[-1]}")
    if pages == prev_pages:
        print("pagination stable, terminé")
        break
    write_toc(d, headings, pages)
    d.save(DOCX)
    prev_pages = pages
else:
    sys.exit("pagination non stabilisée")
render_pdf()
print("OK")
