# -*- coding: utf-8 -*-
"""
Pagina — Compilatore PowerPoint DEIA (mod1).

Flusso:
  1) Carica il PowerPoint template (.pptx) con segnaposto {{chiave}}
  2) Carica i risultati della survey (export Qualtrics .xlsx)
  La libreria DEIA (deia_mapping_GM.xlsx, fogli Libreria/Scala/Mappatura) è
  letta da un file fisso sul server, non più caricata da frontend.
  L'app costruisce il mapping dinamicamente e compila le slide.
"""
import io
from pathlib import Path

import streamlit as st

from mod1.pptx_filler import (
    mapping_from_survey,
    extract_placeholders,
    fill_pptx,
    build_calc_workbook,
)

# libreria DEIA: file fisso sul server, non più caricabile da frontend
MAPPING_PATH = Path(__file__).resolve().parents[1] / "mod1" / "deia_mapping_GM.xlsx"


@st.cache_resource
def _avvia_kaleido():
    """Tiene acceso un solo Chromium condiviso per gli export PNG via kaleido
    (il sunburst): senza, ogni export riavvia il browser da zero (~5s); con il
    server persistente costa ~0,2s dopo il primo avvio. st.cache_resource fa
    sì che parta una sola volta per processo, non a ogni rerun."""
    import shutil
    import kaleido

    chrome_path = (
        shutil.which("chromium")
        or shutil.which("chromium-browser")
        or shutil.which("google-chrome")
    )
    if chrome_path:
        # Streamlit Cloud: usa il Chromium di sistema (installato via packages.txt),
        # le cui librerie condivise sono risolte da apt. Il Chrome standalone
        # scaricato da kaleido si chiudeva subito per librerie mancanti sul
        # container minimale di Streamlit Cloud.
        kaleido.start_sync_server(path=chrome_path)
    else:
        kaleido.get_chrome_sync()  # sviluppo locale: scarica un Chrome dedicato
        kaleido.start_sync_server()
    return True


_avvia_kaleido()

st.title("📊 Compilatore PowerPoint DEIA")
st.caption(
    "Carica il template e i risultati della survey. "
    "L'app costruisce il mapping dai livelli e compila le slide."
)

if not MAPPING_PATH.exists():
    st.error(f"File libreria DEIA non trovato sul server: {MAPPING_PATH.name}")
    st.stop()

col1, col2 = st.columns(2)
with col1:
    pptx_file = st.file_uploader("1 · PowerPoint template (.pptx)", type=["pptx"])
with col2:
    input_file = st.file_uploader(
        "2 · Risultati survey (.xlsx)",
        type=["xlsx"],
        help=(
            "Export Qualtrics raw con le risposte alle domande "
            "(colonne CQ/PQ/SQ/MQ + A3). Esempio: "
            "'Galactica Prova_24 giugno 2026_05.41.xlsx'. "
            "Più rispondenti vengono aggregati per media."
        ),
    )

if not (pptx_file and input_file):
    st.info("Carica il template e i risultati della survey per procedere.")
    st.stop()

# --- costruisce il mapping dinamicamente ------------------------------------
try:
    mapping, details = mapping_from_survey(
        io.BytesIO(input_file.getvalue()),
        MAPPING_PATH,
        with_details=True,
    )
except Exception as e:
    st.error(f"Errore nella generazione del mapping: {e}")
    st.stop()

if not mapping:
    st.error("Il mapping è vuoto: controlla che i file abbiano i fogli e le colonne corrette.")
    st.stop()

# --- diagnostica: cosa si aspetta il template vs cosa c'è nel mapping ------
pptx_bytes = pptx_file.getvalue()
try:
    nel_template = extract_placeholders(io.BytesIO(pptx_bytes))
except Exception as e:
    st.error(f"Errore nel leggere il PowerPoint: {e}")
    st.stop()

mancanti = sorted(nel_template - set(mapping))
with st.expander(
    f"🔎 Mapping generato ({len(mapping)} voci) e diagnostica", expanded=False
):
    st.write(f"Segnaposto trovati nel template: **{len(nel_template)}**")
    if mancanti:
        st.warning(
            "Segnaposto nel template senza testo nel mapping: "
            + ", ".join(f"`{{{{{k}}}}}`" for k in mancanti)
        )
    else:
        st.success("Tutti i segnaposto del template hanno un testo nel mapping. ✅")

    rows = []
    for k, v in sorted(mapping.items()):
        rows.append({"segnaposto": k, "testo": v[:120] + "…" if len(v) > 120 else v})
    st.dataframe(rows, width="stretch", hide_index=True)

# --- compila ----------------------------------------------------------------
buf, stats = fill_pptx(io.BytesIO(pptx_bytes), mapping, details=details)

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

out_name = pptx_file.name.rsplit(".", 1)[0] + "_compilato.pptx"
st.download_button(
    "⬇️ Scarica il PowerPoint compilato",
    data=buf,
    file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    type="primary",
)

calc_buf = build_calc_workbook(details)
st.download_button(
    "⬇️ Scarica i calcoli delle medie (.xlsx)",
    data=calc_buf,
    file_name=pptx_file.name.rsplit(".", 1)[0] + "_calcoli.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Medie di aree/sottogruppi, livello macro e dettaglio per rispondente.",
)
