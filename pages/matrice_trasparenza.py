# -*- coding: utf-8 -*-
"""
Pagina — Matrice di trasparenza DEIA (mod2).

Si caricano i due file sorgente e si ottiene la matrice tema per tema:

    qualtrix_output.xlsx    -> asse X, maturita' DEIA praticata (survey)
    elaborato_giorgia.xlsx  -> asse Y, maturita' DEIA comunicata (reporting)
"""

import io

import pandas as pd
import streamlit as st

from mod2 import deia_core as core
from mod2.grafico import matrice_chart

COLONNE_MATRICE = ["Tema", "X praticata", "Y comunicata", "Quadrante"]

INCHIOSTRO = "#0b0b0b"  # testo sulle tinte chiare della tabella, in entrambi i temi

WASH = {  # tinte tabella: sfondo chiaro + inchiostro scuro esplicito
    "Sovraesposizione": "#fce7dc",
    "Allineamento virtuoso": "#e2efda",
    "Potenziale nascosto": "#fdf0cc",
    "Area fragile": "#ececeb",
}


def tema_attivo():
    try:
        return "dark" if st.context.theme.type == "dark" else "light"
    except Exception:
        return "light"


def it(v, dec=2):
    """Numero in formato italiano (virgola decimale)."""
    if v is None or pd.isna(v):
        return "n.d."
    return f"{v:.{dec}f}".replace(".", ",")


# ------------------------------------------------------------------ layout
st.title("Matrice di trasparenza DEIA")
st.caption(
    "Asse X = maturità DEIA **praticata** (survey Qualtrics) · "
    "Asse Y = maturità DEIA **comunicata** (content analysis dei report). "
    f"Soglia alta/bassa maturità: {it(core.SOGLIA)}."
)

col_x, col_y = st.columns(2)
with col_x:
    file_x = st.file_uploader(
        "**Asse X** — export Qualtrics della survey",
        type=["xlsx"], key="x",
        help="Riga 1 = codici domanda (CQ1, PQ1, SQ1, MQ1…), riga 2 = testi, "
             "righe 3+ = risposte. Es. qualtrix_output.xlsx",
    )
with col_y:
    file_y = st.file_uploader(
        "**Asse Y** — content analysis del reporting",
        type=["xlsx"], key="y",
        help="Foglio matrice_y con colonne Azienda, Dimensione, Tema reporting, "
             "Score reporting. Es. elaborato_giorgia.xlsx",
    )

if not (file_x and file_y):
    st.info("Carica entrambi i file per calcolare la matrice.")
    st.stop()

# ------------------------------------------------------------------ lettura
try:
    rispondenti = core.elenca_rispondenti(file_x)
    reporting = core.leggi_reporting(file_y)
except core.ErroreInput as exc:
    st.error(f"Input non valido: {exc}")
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
    survey = core.leggi_survey(file_x, riga=riga)
    righe, dettaglio, azienda_y = core.costruisci_matrice(survey, reporting, azienda_y)
except core.ErroreInput as exc:
    st.error(f"Input non valido: {exc}")
    st.stop()

df = pd.DataFrame(righe)
escluse = [d for d in dettaglio if d["inclusa"] != core.SI]

st.divider()
m1, m2, m3 = st.columns(3)
m1.metric("Caso survey (X)", survey["azienda"],
          f"{survey['dimensione']} · peso "
          f"{'PMI' if survey['dimensione'] == 'PMI' else 'GRANDI'}",
          delta_color="off")
m2.metric("Caso reporting (Y)", azienda_y)
m3.metric("Domande incluse nel calcolo",
          f"{len(dettaglio) - len(escluse)}/{len(dettaglio)}",
          f"{len(escluse)} escluse", delta_color="off")
if survey["azienda"] != azienda_y:
    st.warning(
        f"I due assi provengono da soggetti diversi ({survey['azienda']} per la "
        f"survey, {azienda_y} per il reporting): la matrice va letta come prova "
        "di calcolo, non come diagnosi di una singola organizzazione."
    )

# ------------------------------------------------------------------ tabella
st.subheader("Matrice")

tabella = pd.DataFrame({
    "Tema": df["Tema"],
    "X praticata": df["X praticata"].map(it),
    "Y comunicata": df["Y comunicata"].map(it),
    "Quadrante": df["Quadrante"],
})


def tinta(riga_tab):
    colore = WASH.get(riga_tab["Quadrante"])
    stile = f"background-color:{colore};color:{INCHIOSTRO}" if colore else ""
    return ["", "", "", stile]


st.dataframe(tabella.style.apply(tinta, axis=1), hide_index=True,
             width="stretch")

conteggio = df[df["Quadrante"] != ""]["Quadrante"].value_counts()
st.caption(" · ".join(f"**{q}**: {n}" for q, n in conteggio.items()))

# ------------------------------------------------------------------ grafico
chart = matrice_chart(df, core.SOGLIA, tema_attivo())
if chart is not None:
    st.subheader("Posizionamento nei quadranti")
    # larghezza fissa: il posizionamento delle etichette e' calcolato in pixel
    st.altair_chart(chart, width="content")

# ------------------------------------------------------------------ dettagli
with st.expander("Dettaglio per tema (livelli, scostamento, priorità)"):
    completa = df.copy()
    for c in ("X praticata", "Y comunicata"):
        completa[c] = completa[c].map(it)
    completa["Differenza Y-X"] = df["Differenza Y-X"].map(
        lambda v: "n.d." if v is None or pd.isna(v) else
        ("+" if v > 0 else "") + it(v))
    st.dataframe(completa, hide_index=True, width="stretch")

with st.expander(f"Come è stato calcolato l'asse X ({len(dettaglio)} domande)"):
    st.markdown(
        "X(tema) = Σ(risposta × peso × quota) ⁄ Σ(peso × quota). "
        "La quota è 100% sul Tema 1 se il Cod. 2 è vuoto, altrimenti 70% / 30%. "
        "Sono escluse le domande con Cod. standard **MAT**, quelle senza Cod. 1, "
        "senza peso o con risposta fuori scala 1-4."
    )
    if escluse:
        st.markdown("**Domande escluse:** " + ", ".join(
            f"`{d['cod']}` ({d['motivo']})" for d in escluse))
    dett = pd.DataFrame(dettaglio)[
        ["cod", "framework", "sottogruppo", "standard", "cod1", "cod2", "tema1",
         "tema2", "peso", "risposta", "alloc1", "alloc2", "inclusa", "motivo"]
    ].rename(columns={
        "cod": "Cod. domanda", "framework": "Framework", "sottogruppo": "Sottogruppo",
        "standard": "Cod. standard", "cod1": "Cod. 1", "cod2": "Cod. 2",
        "tema1": "Tema 1", "tema2": "Tema 2", "peso": "Peso usato",
        "risposta": "Risposta", "alloc1": "% alloc. 1", "alloc2": "% alloc. 2",
        "inclusa": "Inclusa", "motivo": "Motivo / stato"})
    # PQ19 non ha peso: la colonna resta mista, va resa numerica per la tabella
    for c in ("Peso usato", "Risposta"):
        dett[c] = pd.to_numeric(dett[c], errors="coerce")
    st.dataframe(dett, hide_index=True, width="stretch")

# ------------------------------------------------------------------ download
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    tabella.to_excel(writer, sheet_name="Matrice", index=False)
    df.to_excel(writer, sheet_name="Dettaglio temi", index=False)
    dett.to_excel(writer, sheet_name="Calcolo asse X", index=False)

d1, d2 = st.columns(2)
d1.download_button("Scarica la matrice (CSV)",
                   tabella.to_csv(index=False, sep=";").encode("utf-8-sig"),
                   "matrice_trasparenza.csv", "text/csv", width="stretch")
d2.download_button("Scarica tutto (Excel)", buffer.getvalue(),
                   "matrice_trasparenza.xlsx", width="stretch",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
