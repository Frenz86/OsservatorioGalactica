# -*- coding: utf-8 -*-
"""
App Streamlit — Compilatore PowerPoint DEIA.

Flusso:
  1) Carica il PowerPoint template (.pptx) con segnaposto {{chiave}}
  2) Carica i risultati della survey (input_test.xlsx  →  1 riga, colonne = id, valori = livello 1-4)
  3) Carica la libreria DEIA (deia_mapping.xlsx  →  fogli Libreria + Scala)
  L'app costruisce il mapping dinamicamente e compila le slide.

Avvio:
    pip install -r requirements.txt
    streamlit run app.py
"""
import io
import streamlit as st

from pptx_filler import mapping_from_survey, extract_placeholders, fill_pptx

st.set_page_config(page_title="Compilatore PowerPoint DEIA", page_icon="📊", layout="wide")

# --- login ------------------------------------------------------------------
CREDENZIALI = {
    "admin": "admin",
}

if not st.session_state.get("autenticato"):
    st.title("📊 Compilatore PowerPoint DEIA")
    utente = st.text_input("Utente")
    pwd = st.text_input("Password", type="password")
    if st.button("Accedi"):
        if CREDENZIALI.get(utente) == pwd and pwd:
            st.session_state["autenticato"] = True
            st.rerun()
        else:
            st.error("Utente o password errati.")
    st.stop()
# ----------------------------------------------------------------------------

st.title("📊 Compilatore PowerPoint DEIA")
st.caption(
    "Carica il template, i risultati della survey e la libreria DEIA. "
    "L'app costruisce il mapping dai livelli e compila le slide."
)

col1, col2, col3 = st.columns(3)
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
with col3:
    mapping_file = st.file_uploader(
        "3 · Libreria DEIA (.xlsx)",
        type=["xlsx"],
        help=(
            "Fogli richiesti: 'Libreria' e 'Scala' sempre; 'Mappatura' "
            "(pesi domanda->area PMI/GRANDI) serve solo per la survey raw. "
            "Esempio: deia_mapping_GM.xlsx"
        ),
    )

if not (pptx_file and input_file and mapping_file):
    st.info("Carica tutti e tre i file per procedere.")
    st.stop()

# --- costruisce il mapping dinamicamente ------------------------------------
try:
    mapping = mapping_from_survey(
        io.BytesIO(input_file.getvalue()),
        io.BytesIO(mapping_file.getvalue()),
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
    st.dataframe(rows, use_container_width=True, hide_index=True)

# --- compila ----------------------------------------------------------------
buf, stats = fill_pptx(io.BytesIO(pptx_bytes), mapping)

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
