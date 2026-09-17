# -*- coding: utf-8 -*-
"""Verifica end-to-end della pipeline app (senza Streamlit)."""
import io
import json
import sys
from pathlib import Path

# console Windows (cp1252): forza UTF-8 per stampare emoji/accenti
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from pptx import Presentation
from pptx.util import Inches, Pt
from mod1.pptx_filler import load_mapping, extract_placeholders, fill_pptx

BASE = Path(__file__).resolve().parent
ok = True


def check(cond, msg):
    global ok
    print(("  OK  " if cond else " FAIL ") + msg)
    if not cond:
        ok = False


def _save(prs):
    b = io.BytesIO()
    prs.save(b)
    b.seek(0)
    return b


def all_text(prs):
    out = []
    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_text_frame:
                out.append(sh.text_frame.text)
            if sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells:
                        out.append(c.text)
    return "\n".join(out)


print("== 1. Mapping da Excel (default) ==")
mapping = load_mapping(BASE / "deia_mapping_GM.xlsx")
check(len(mapping) == 35, f"35 voci di mapping (trovate {len(mapping)})")
check(mapping.get("strategia.risultato") == "2 - ROTTA",
      "strategia.risultato == '2 - ROTTA'")
check(mapping.get("cultura.risultato") == "1 - DERIVA",
      "cultura.risultato == '1 - DERIVA'")
check("inizia a definire" in mapping.get("strategia.commento", ""),
      "strategia.commento = testo livello 2 (dal documento)")
check(mapping.get("panoramica.risultato") == "1 - DERIVA",
      "panoramica.risultato == '1 - DERIVA' (livello macro)")
check(mapping.get("panoramica.livello_medio") == "1,25 / 4",
      "panoramica.livello_medio == '1,25 / 4'")

print("== 2. Mapping rigenerato dai livelli (Assessment+Libreria) ==")
mapping2 = load_mapping(BASE / "deia_mapping_GM.xlsx", regenerate=True)
check(mapping2 == mapping, "mapping rigenerato identico a quello pronto")

print("== 3. Segnaposto nel template ==")
ph = extract_placeholders(BASE / "template_esempio.pptx")
check(len(ph) == 35, f"35 segnaposto nel template (trovati {len(ph)})")
check(ph == set(mapping), "i segnaposto del template coincidono col mapping")

print("== 4. Compilazione del template ==")
buf, stats = fill_pptx(BASE / "template_esempio.pptx", mapping)
check(stats["n_non_risolti"] == 0, f"0 segnaposto non risolti (={stats['n_non_risolti']})")
check(stats["n_sostituiti"] == 35, f"35 segnaposto sostituiti (={stats['n_sostituiti']})")
check(len(stats["inutilizzati"]) == 0, "0 voci di mapping inutilizzate")
filled = Presentation(buf)
txt = all_text(filled)
check("{{" not in txt and "}}" not in txt, "nessuna graffa residua nel file compilato")
check("2 - ROTTA" in txt and "1 - DERIVA" in txt, "risultati presenti nel testo")
check("inizia a definire" in txt, "commento strategia presente")
check("Integrare gli obiettivi DEIA" in txt, "attività sottogruppo presente")
# salva una copia compilata di esempio
buf.seek(0)
(BASE / "assessment_demo_compilato.pptx").write_bytes(buf.getvalue())

print("== 5. Robustezza: segnaposto spezzato su più run ==")
prs = Presentation()
sl = prs.slides.add_slide(prs.slide_layouts[6])
tb = sl.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
p = tb.text_frame.paragraphs[0]
for frag in ["{{strat", "egia.ris", "ultato}}"]:   # spezzato in 3 run
    r = p.add_run(); r.text = frag; r.font.size = Pt(18)
b2, s2 = fill_pptx(_save(prs), mapping)
t2 = all_text(Presentation(b2))
check(t2.strip() == "2 - ROTTA", f"run spezzati ricomposti -> '2 - ROTTA' (got '{t2.strip()}')")

print("== 6. Robustezza: segnaposto dentro una tabella ==")
prs = Presentation()
sl = prs.slides.add_slide(prs.slide_layouts[6])
gt = sl.shapes.add_table(1, 1, Inches(1), Inches(1), Inches(6), Inches(1)).table
gt.cell(0, 0).text = "{{cultura.commento}}"
b3, s3 = fill_pptx(_save(prs), mapping)
t3 = all_text(Presentation(b3))
check("Valori, comportamenti" in t3 and "{{" not in t3, "segnaposto in tabella sostituito")

print("== 7. Radar dinamico sul template compilato ==")
from pptx.enum.shapes import MSO_SHAPE_TYPE
buf.seek(0)
last = Presentation(buf)                       # template compilato al punto 4
pics = [sh for s in last.slides for sh in s.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
check(len(pics) == 1, f"1 immagine radar nel file compilato (trovate {len(pics)})")
radar_txt = all_text(last)
check("{{radar" not in radar_txt, "nessun marcatore radar residuo nel file compilato")
check("{{sunburst" not in radar_txt,
      "senza 'details' il marcatore sunburst è rimosso senza lasciare graffe")
check("{{matrice" not in radar_txt and "{{descrizioni" not in radar_txt,
      "senza 'immagini_extra'/'tabelle_extra' i marcatori mod2 sono rimossi senza lasciare graffe")

print("== 8. Sunburst macro/micro sulla slide finale ==")
from mod1.pptx_filler import _sunburst_image, _sigla, AREA_ORDER

check(_sigla("Benessere e Sicurezza Psicologica") == "BSP", "sigla sottogruppo 'BSP'")
check(_sigla("Strategia e Pianificazione") == "SP", "sigla sottogruppo 'SP'")

fake_details = {
    "aree": [
        {"id": aid, "nome": aid.upper(), "score": 2.5, "livello": 2, "nome_livello": "ROTTA"}
        for aid in AREA_ORDER
    ],
    "sottogruppi": [
        {"id": f"{aid}_sub{i}", "nome": f"Sottogruppo {i}", "area": aid.upper(),
         "score": 1.5 + i, "livello": 2, "nome_livello": "ROTTA"}
        for aid in AREA_ORDER for i in range(1, 4)
    ],
}
check(_sunburst_image(fake_details) is not None, "_sunburst_image genera un'immagine con dati validi")
check(_sunburst_image({"aree": [], "sottogruppi": []}) is None, "_sunburst_image torna None senza aree")

prs = Presentation()
sl = prs.slides.add_slide(prs.slide_layouts[6])
tb = sl.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(4))
tb.text_frame.text = "{{sunburst.aree}}"
b4, s4 = fill_pptx(_save(prs), {}, details=fake_details)
pics4 = [sh for sh in Presentation(b4).slides[0].shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
check(len(pics4) == 1, "{{sunburst.aree}} sostituito con un'immagine quando 'details' è presente")

print("== 9. Marcatori generici immagine/tabella (immagini_extra/tabelle_extra) ==")
from pptx.dml.color import RGBColor

prs = Presentation()
sl = prs.slides.add_slide(prs.slide_layouts[6])
img_box = sl.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(2))
img_box.text_frame.text = "{{matrice.grafico}}"
tab_box = sl.shapes.add_textbox(Inches(1), Inches(4), Inches(6), Inches(2))
tab_box.text_frame.text = "{{matrice.tabella}}"

fake_png = _sunburst_image(fake_details)          # un PNG qualsiasi, basta come contenuto
immagini_extra = {"matrice.grafico": fake_png}
tabelle_extra = {"matrice.tabella": {
    "headers": ["Tema", "X", "Y"],
    "rows": [["Foundation", "1,86", "4,00"], ["HR", "2,56", "3,00"]],
    "row_colors": [RGBColor(0xFC, 0xE7, 0xDC), None],
}}
b5, s5 = fill_pptx(_save(prs), {}, immagini_extra=immagini_extra, tabelle_extra=tabelle_extra)
p5 = Presentation(b5)
pics5 = [sh for s in p5.slides for sh in s.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
check(len(pics5) == 1, "{{matrice.grafico}} sostituito con un'immagine quando 'immagini_extra' è presente")
tabelle5 = [sh.table for s in p5.slides for sh in s.shapes if sh.has_table]
check(len(tabelle5) == 1, "{{matrice.tabella}} sostituito con una tabella quando 'tabelle_extra' è presente")
if tabelle5:
    t = tabelle5[0]
    check((len(t.rows), len(t.columns)) == (3, 3), "tabella con intestazione + 2 righe dati, 3 colonne")
    check(t.cell(0, 0).text == "Tema" and t.cell(1, 0).text == "Foundation" and t.cell(2, 1).text == "2,56",
          "contenuto delle celle della tabella corretto")
txt5 = all_text(p5)
check("{{" not in txt5, "nessuna graffa residua con marcatori generici popolati")

# senza immagini_extra/tabelle_extra, gli stessi marcatori vengono svuotati
prs2 = Presentation()
sl2 = prs2.slides.add_slide(prs2.slide_layouts[6])
sl2.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(2)).text_frame.text = "{{matrice.grafico}}"
sl2.shapes.add_textbox(Inches(1), Inches(4), Inches(6), Inches(2)).text_frame.text = "{{descrizioni.tabella1}}"
b6, s6 = fill_pptx(_save(prs2), {})
txt6 = all_text(Presentation(b6))
check("{{" not in txt6,
      "senza 'immagini_extra'/'tabelle_extra' i marcatori generici sono rimossi senza lasciare graffe")

print("\nJSON valido:", end=" ")
json.loads((BASE / "deia_framework.json").read_text(encoding="utf-8"))
print("sì")

print("\nRISULTATO:", "TUTTO OK ✅" if ok else "CI SONO FALLIMENTI ❌")
raise SystemExit(0 if ok else 1)
