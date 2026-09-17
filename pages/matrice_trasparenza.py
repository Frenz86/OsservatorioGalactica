# -*- coding: utf-8 -*-
"""
Pagina — Matrice di trasparenza DEIA (mod2).

Si carica un solo file (01qualtrix_output_finale.xlsx) e se ne leggono due
fogli diversi per i due assi della matrice:

    foglio Qualtrix_output         -> asse X, maturita' DEIA praticata (survey)
    foglio Grafici_1 (o Grafici_2) -> asse Y, maturita' DEIA comunicata (reporting)

Resta supportato anche il vecchio foglio matrice_y (elaborato_giorgia.xlsx) per
l'asse Y, nello stesso file o in un file separato.
"""

import io
from pathlib import Path

import pandas as pd
import streamlit as st

from mod2 import deia_core as core
from mod2.grafico import matrice_chart, radar_figure

COLONNE_MATRICE = ["Tema", "X praticata", "Y comunicata", "Quadrante"]

# libreria descrizioni per punteggio (asse Y): file fisso sul server, come la
# libreria DEIA di mod1 — non si carica da frontend
COMMENTI_PATH = Path(__file__).resolve().parents[1] / "mod2" / "Y_commenti_dinamici.xlsx"

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

file_up = st.file_uploader(
    "File export",
    type=["xlsx"], key="file",
    help="01qualtrix_output_finale.xlsx: asse X dal foglio 'Qualtrix_output' "
         "(riga 1 = codici domanda CQ1/PQ1/SQ1/MQ1…, riga 2 = testi, righe 3+ = "
         "risposte), asse Y dal foglio 'Grafici_1' o 'Grafici_2' (una riga per "
         "azienda, una colonna per tema: Foundation, Onboarding & Retention, "
         "Employment…). Resta supportato anche il vecchio foglio matrice_y "
         "(Azienda, Dimensione, Tema reporting, Score reporting).",
)

if not file_up:
    st.info("Carica il file per calcolare la matrice.")
    st.stop()

# ------------------------------------------------------------------ lettura
try:
    rispondenti = core.elenca_rispondenti(file_up)
    reporting = core.leggi_reporting(file_up)
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
    survey = core.leggi_survey(file_up, riga=riga)
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

# --------------------------------------------------------------- radar (Y)
per_azienda_grafici, temi_grafici, dim_grafici = core.leggi_valori_grafici(file_up)
valori_y = per_azienda_grafici.get(azienda_y, {})
if valori_y and temi_grafici and dim_grafici:
    st.subheader("Profilo radar (Grafici_1)")
    rc1, rc2 = st.columns(2)
    with rc1:
        st.plotly_chart(
            radar_figure(temi_grafici, valori_y, f"{azienda_y.upper()}_REPORTING",
                         "#3a9d5d", "rgba(58,157,93,0.18)"),
            width="stretch",
        )
    with rc2:
        st.plotly_chart(
            radar_figure(dim_grafici, valori_y, f"{azienda_y.upper()}_DIVERSITA'",
                         "#d97a95", "rgba(217,122,149,0.18)"),
            width="stretch",
        )

# ------------------------------------------------------- radar media settore
per_settore, temi_settore, dim_settore = core.leggi_valori_grafici(
    file_up, nomi_foglio=("Grafici_2",))
nome_media = next(
    (a for a in per_settore if core.norm(a) == core.norm("Media settore")), None)
valori_media = per_settore.get(nome_media, {}) if nome_media else {}
if valori_media and temi_settore and dim_settore:
    st.caption(
        f"Benchmark: media di {len(per_settore) - 1} aziende concorrenti "
        "(foglio Grafici_2)."
    )
    rc3, rc4 = st.columns(2)
    with rc3:
        st.plotly_chart(
            radar_figure(temi_settore, valori_media, "MEDIA SETTORE_REPORTING",
                         "#5b7fa6", "rgba(91,127,166,0.18)"),
            width="stretch",
        )
    with rc4:
        st.plotly_chart(
            radar_figure(dim_settore, valori_media, "MEDIA SETTORE_DIVERSITA'",
                         "#5b7fa6", "rgba(91,127,166,0.18)"),
            width="stretch",
        )

# ------------------------------------------------------------- descrizioni Y
def _livello_y(y):
    if y is None or pd.isna(y):
        return None
    return max(1, min(4, int(round(y))))


commenti = {}
if COMMENTI_PATH.exists():
    try:
        commenti = core.leggi_commenti(COMMENTI_PATH)
    except core.ErroreInput:
        commenti = {}

if commenti:
    st.subheader("Descrizioni per punteggio (asse Y)")
    st.caption(
        f"Testo di riferimento per il punteggio comunicato di {azienda_y} su "
        "ciascun tema, dalla libreria Y_commenti_dinamici.xlsx."
    )
    tabella_comm = pd.DataFrame([
        {
            "Tema": r["Tema"],
            "Punteggio Y": it(r["Y comunicata"]),
            "Livello": _livello_y(r["Y comunicata"]) or "n.d.",
            "Descrizione": commenti.get(
                (r["Tema"], _livello_y(r["Y comunicata"])),
                "n.d." if _livello_y(r["Y comunicata"]) is None else "",
            ),
        }
        for r in righe
    ])
    st.dataframe(
        tabella_comm, hide_index=True, width="stretch",
        column_config={"Descrizione": st.column_config.TextColumn(width="large")},
    )
elif COMMENTI_PATH.exists():
    st.caption(
        "Libreria descrizioni trovata ma senza il foglio 'Commenti coding' "
        "atteso: nessuna descrizione per punteggio mostrata."
    )

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
    if commenti:
        tabella_comm.to_excel(writer, sheet_name="Descrizioni Y", index=False)

d1, d2 = st.columns(2)
d1.download_button("Scarica la matrice (CSV)",
                   tabella.to_csv(index=False, sep=";").encode("utf-8-sig"),
                   "matrice_trasparenza.csv", "text/csv", width="stretch")
d2.download_button("Scarica tutto (Excel)", buffer.getvalue(),
                   "matrice_trasparenza.xlsx", width="stretch",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
