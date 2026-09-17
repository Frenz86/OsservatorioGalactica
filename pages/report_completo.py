# -*- coding: utf-8 -*-
"""
Pagina — Report completo DEIA (mod1 + mod2 in un unico PPTX).

Orchestratore: chiama le funzioni già esistenti di mod1 (mapping aree/
sottogruppi + radar 4 aree + sunburst) e mod2 (matrice di trasparenza a 10
temi, 4 radar reporting/diversità azienda e media settore, descrizioni per
punteggio) e le passa a fill_pptx (mod1/pptx_filler.py) tramite i parametri
generici 'immagini_extra'/'tabelle_extra'. mod1 e mod2 restano indipendenti:
non si importano a vicenda, questa pagina è l'unico punto che li unisce.

Input: un template PPTX con segnaposto {{...}} (stile mod1/template_esempio.pptx,
comprese le slide "PROFILO RADAR ESTERNO"/"MATRICE DI TRASPARENZA"/"DESCRIZIONI
PER PUNTEGGIO") + un solo file export (01qualtrix_output_finale.xlsx: fogli
Qualtrix_output/Grafici_1/Grafici_2).
"""
import io
from pathlib import Path

import pandas as pd
import streamlit as st
from pptx.dml.color import RGBColor

from mod1.pptx_filler import (
    mapping_from_survey,
    extract_placeholders,
    fill_pptx,
    build_calc_workbook,
)
from mod2 import deia_core as core
from mod2.grafico import radar_figure, radar_png, matrice_chart_png

MAPPING_PATH = Path(__file__).resolve().parents[1] / "mod1" / "deia_mapping_GM.xlsx"
COMMENTI_PATH = Path(__file__).resolve().parents[1] / "mod2" / "Y_commenti_dinamici.xlsx"


@st.cache_resource
def _avvia_kaleido():
    """Tiene acceso un solo Chromium condiviso per tutti gli export PNG via
    kaleido (i 4 radar + il sunburst di mod1): senza, ogni fig.write_image()
    riavvia il browser da zero (~5s l'uno); con il server persistente il primo
    export costa ~2s e i successivi ~0,2s. st.cache_resource fa sì che il
    server venga avviato una sola volta per processo, non a ogni rerun."""
    import kaleido
    kaleido.start_sync_server()
    return True


_avvia_kaleido()

WASH_RGB = {
    "Sovraesposizione": RGBColor.from_string("FCE7DC"),
    "Allineamento virtuoso": RGBColor.from_string("E2EFDA"),
    "Potenziale nascosto": RGBColor.from_string("FDF0CC"),
    "Area fragile": RGBColor.from_string("ECECEB"),
}


def it(v, dec=2):
    """Numero in formato italiano (virgola decimale)."""
    if v is None:
        return "n.d."
    return f"{v:.{dec}f}".replace(".", ",")


def livello_y(y):
    if y is None:
        return None
    return max(1, min(4, int(round(y))))


st.title("🧩 Report completo DEIA")
st.caption(
    "Un solo PPTX con tutto: aree/sottogruppi, radar e sunburst di maturità "
    "(mod1) più matrice di trasparenza, 4 radar reporting/diversità e "
    "descrizioni per punteggio (mod2)."
)

if not MAPPING_PATH.exists():
    st.error(f"File libreria DEIA non trovato sul server: {MAPPING_PATH.name}")
    st.stop()

col1, col2 = st.columns(2)
with col1:
    pptx_file = st.file_uploader(
        "1 · PowerPoint template (.pptx)", type=["pptx"],
        help="Segnaposto {{...}} nello stile di mod1/template_esempio.pptx, "
             "comprese le slide dei radar esterni, della matrice e delle "
             "descrizioni per punteggio.",
    )
with col2:
    input_file = st.file_uploader(
        "2 · File export (.xlsx)", type=["xlsx"],
        help="01qualtrix_output_finale.xlsx: fogli Qualtrix_output (survey), "
             "Grafici_1 (azienda) e Grafici_2 (concorrenti + Media settore).",
    )

if not (pptx_file and input_file):
    st.info("Carica il template e il file export per procedere.")
    st.stop()

pptx_bytes = pptx_file.getvalue()
xlsx_bytes = input_file.getvalue()

# --- mod1: mapping aree/sottogruppi -----------------------------------------
try:
    mapping, details = mapping_from_survey(
        io.BytesIO(xlsx_bytes), MAPPING_PATH, with_details=True,
    )
except Exception as e:
    st.error(f"Errore nella generazione del mapping (mod1): {e}")
    st.stop()

# --- mod2: matrice, radar, descrizioni --------------------------------------
try:
    rispondenti = core.elenca_rispondenti(io.BytesIO(xlsx_bytes))
    reporting = core.leggi_reporting(io.BytesIO(xlsx_bytes))
except core.ErroreInput as exc:
    st.error(f"Errore nella lettura dei dati di reporting (mod2): {exc}")
    st.stop()

aziende_y = core.aziende_reporting(reporting) or core.aziende_reporting(reporting, False)

sel_x, sel_y = st.columns(2)
with sel_x:
    riga = (rispondenti[0][0] if len(rispondenti) == 1 else
            st.selectbox("Rispondente survey (X)", [r[0] for r in rispondenti],
                         format_func=lambda n: dict(rispondenti)[n]))
with sel_y:
    azienda_y = (aziende_y[0] if len(aziende_y) == 1 else
                 st.selectbox("Azienda reporting (Y)", aziende_y))

try:
    survey = core.leggi_survey(io.BytesIO(xlsx_bytes), riga=riga)
    righe, dettaglio, azienda_y = core.costruisci_matrice(survey, reporting, azienda_y)
except core.ErroreInput as exc:
    st.error(f"Errore nel calcolo della matrice (mod2): {exc}")
    st.stop()

st.divider()
m1, m2, m3 = st.columns(3)
m1.metric("Caso survey (X)", survey["azienda"], delta_color="off")
m2.metric("Caso reporting (Y)", azienda_y)
n_escluse = len([d for d in dettaglio if d["inclusa"] != core.SI])
m3.metric("Domande incluse nel calcolo", f"{len(dettaglio) - n_escluse}/{len(dettaglio)}")

# --- immagini_extra: 4 radar esterni + grafico matrice ----------------------
immagini_extra = {}

per_azienda, temi_g, dim_g = core.leggi_valori_grafici(io.BytesIO(xlsx_bytes))
valori_azienda = per_azienda.get(azienda_y, {})
if valori_azienda and temi_g and dim_g:
    immagini_extra["radar_reporting.azienda"] = radar_png(radar_figure(
        temi_g, valori_azienda, f"{azienda_y.upper()}_REPORTING",
        "#3a9d5d", "rgba(58,157,93,0.18)"))
    immagini_extra["radar_diversita.azienda"] = radar_png(radar_figure(
        dim_g, valori_azienda, f"{azienda_y.upper()}_DIVERSITA'",
        "#d97a95", "rgba(217,122,149,0.18)"))

per_settore, temi_s, dim_s = core.leggi_valori_grafici(
    io.BytesIO(xlsx_bytes), nomi_foglio=("Grafici_2",))
nome_media = next(
    (a for a in per_settore if core.norm(a) == core.norm("Media settore")), None)
valori_settore = per_settore.get(nome_media, {}) if nome_media else {}
if valori_settore and temi_s and dim_s:
    immagini_extra["radar_reporting.settore"] = radar_png(radar_figure(
        temi_s, valori_settore, "MEDIA SETTORE_REPORTING",
        "#5b7fa6", "rgba(91,127,166,0.18)"))
    immagini_extra["radar_diversita.settore"] = radar_png(radar_figure(
        dim_s, valori_settore, "MEDIA SETTORE_DIVERSITA'",
        "#5b7fa6", "rgba(91,127,166,0.18)"))

df_matrice = pd.DataFrame(righe)  # formato atteso da matrice_chart_png
matrice_png = matrice_chart_png(df_matrice, core.SOGLIA, "light")
if matrice_png is not None:
    immagini_extra["matrice.grafico"] = matrice_png

# --- tabelle_extra: matrice + descrizioni per punteggio ---------------------
tabelle_extra = {
    "matrice.tabella": {
        "headers": ["Tema", "X praticata", "Y comunicata", "Quadrante"],
        "rows": [[r["Tema"], it(r["X praticata"]), it(r["Y comunicata"]), r["Quadrante"]]
                 for r in righe],
        "row_colors": [WASH_RGB.get(r["Quadrante"]) for r in righe],
    },
}

commenti = {}
if COMMENTI_PATH.exists():
    try:
        commenti = core.leggi_commenti(COMMENTI_PATH)
    except core.ErroreInput:
        commenti = {}

if commenti:
    for indice, gruppo in ((1, righe[:5]), (2, righe[5:])):
        tabelle_extra[f"descrizioni.tabella{indice}"] = {
            "headers": ["Tema", "Punteggio", "Descrizione"],
            "rows": [
                [r["Tema"], it(r["Y comunicata"]),
                 commenti.get((r["Tema"], livello_y(r["Y comunicata"])), "")]
                for r in gruppo
            ],
        }

# --- diagnostica segnaposto --------------------------------------------------
try:
    nel_template = extract_placeholders(io.BytesIO(pptx_bytes))
except Exception as e:
    st.error(f"Errore nel leggere il PowerPoint: {e}")
    st.stop()

mancanti = sorted(nel_template - set(mapping))
with st.expander(f"🔎 Mapping generato ({len(mapping)} voci) e diagnostica", expanded=False):
    st.write(f"Segnaposto testuali trovati nel template: **{len(nel_template)}**")
    if mancanti:
        st.warning(
            "Segnaposto nel template senza testo nel mapping: "
            + ", ".join(f"`{{{{{k}}}}}`" for k in mancanti)
        )
    else:
        st.success("Tutti i segnaposto testuali del template hanno un testo nel mapping. ✅")
    st.write(
        f"Marcatori mod2 disponibili: immagini {', '.join(sorted(immagini_extra)) or '—'}; "
        f"tabelle {', '.join(sorted(tabelle_extra)) or '—'}."
    )

# --- compila ------------------------------------------------------------------
buf, stats = fill_pptx(
    io.BytesIO(pptx_bytes), mapping, details=details,
    immagini_extra=immagini_extra, tabelle_extra=tabelle_extra,
)

c1, c2, c3 = st.columns(3)
c1.metric("Segnaposto compilati", stats["n_sostituiti"])
c2.metric("Non risolti", stats["n_non_risolti"])
c3.metric("Voci mapping inutilizzate", len(stats["inutilizzati"]))

if stats["non_risolti"]:
    st.warning(
        "Segnaposto rimasti nel file (nessun testo nel mapping): "
        + ", ".join(f"`{{{{{k}}}}}`" for k in stats["non_risolti"])
    )
else:
    st.success("Nessun segnaposto rimasto: il PowerPoint è completamente compilato. ✅")

out_name = pptx_file.name.rsplit(".", 1)[0] + "_report_completo.pptx"
st.download_button(
    "⬇️ Scarica il Report completo (.pptx)",
    data=buf, file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    type="primary",
)

calc_buf = build_calc_workbook(details)
st.download_button(
    "⬇️ Scarica i calcoli delle medie mod1 (.xlsx)",
    data=calc_buf, file_name=pptx_file.name.rsplit(".", 1)[0] + "_calcoli.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
