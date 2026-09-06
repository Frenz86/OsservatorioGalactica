# -*- coding: utf-8 -*-
"""
Grafico della matrice: scatter 1-4 x 1-4 con la croce delle soglie.

Una sola serie di colore: e' la posizione a codificare il quadrante, quindi il
colore non deve ripeterla. I dieci temi portano etichetta diretta, cosi'
l'identita' non passa mai dal colore soltanto.

L'asse Y e' un intero 1-4 di content analysis, percio' molti temi finiscono
esattamente alla stessa quota: le etichette vengono distanziate da un
posizionamento calcolato (nessuna sovrapposizione) e ricollegate al proprio
punto da una linea guida sottile.

Palette verificata con lo strumento di controllo del design system (banda di
luminosita', chroma, contrasto >= 3:1) sulle superfici di Streamlit #ffffff e
#0e1117.
"""

import altair as alt
import pandas as pd

SERIE = {"light": "#2a78d6", "dark": "#3987e5"}
INK = {"light": "#0b0b0b", "dark": "#ffffff"}
GRIGLIA = {"light": "#e1e0d9", "dark": "#2c2c2a"}
MUTED = "#898781"

ANGOLI = (
    (1.04, 3.97, "SOVRAESPOSIZIONE", "left", "top"),
    (3.96, 3.97, "ALLINEAMENTO VIRTUOSO", "right", "top"),
    (1.04, 1.03, "AREA FRAGILE", "left", "bottom"),
    (3.96, 1.03, "POTENZIALE NASCOSTO", "right", "bottom"),
)

DOMINIO = [1, 4]
SCALA = alt.Scale(domain=DOMINIO, nice=False)
TACCHE = [1, 1.5, 2, 2.5, 3, 3.5, 4]

PASSO_ETICHETTA = 17       # px fra due etichette accatastate
LARGHEZZA_CARATTERE = 6.6  # px per carattere a 12px, stima per il test di collisione
SCARTO_X = 10              # px fra punto ed etichetta
RAGGIO_PUNTO = 8           # px, meta' del marker: serve come ingombro
MARGINE = 9                # px di rispetto dai bordi del riquadro


TITOLO_X = "X — maturità praticata (survey)"
TITOLO_Y = "Y — maturità comunicata (reporting)"


def _asse(titolo, modo):
    """Un layer con asse diverso dagli altri ne annulla il merge: sempre identico."""
    return alt.Axis(title=titolo, values=TACCHE, grid=True, gridColor=GRIGLIA[modo],
                    domainColor=MUTED, tickColor=MUTED, labelColor=MUTED,
                    titleColor=MUTED, labelFontSize=11)


def _x(campo, modo):
    return alt.X(f"{campo}:Q", scale=SCALA, axis=_asse(TITOLO_X, modo))


def _y(campo, modo):
    return alt.Y(f"{campo}:Q", scale=SCALA, axis=_asse(TITOLO_Y, modo))


def _posiziona_etichette(plot, altezza, larghezza):
    """Sposta in verticale le etichette finche' nessuna coppia si sovrappone.

    Lavora in pixel (stima) e restituisce la quota in unita' di dato, cosi'
    Altair riceve solo coordinate gia' risolte.
    """
    span = DOMINIO[1] - DOMINIO[0]

    def a_px_y(v):
        return (DOMINIO[1] - v) / span * altezza

    def a_dato_y(py):
        return DOMINIO[1] - py / altezza * span

    def a_px_x(v):
        return (v - DOMINIO[0]) / span * larghezza

    righe = []
    for _, r in plot.iterrows():
        larg = len(str(r["Tema"])) * LARGHEZZA_CARATTERE
        px = a_px_x(r["X praticata"])
        # a destra del punto; a sinistra se l'etichetta uscirebbe dal riquadro
        a_destra = px + SCARTO_X + larg <= larghezza
        sinistra = px + SCARTO_X if a_destra else px - SCARTO_X - larg
        righe.append({
            "Tema": r["Tema"], "x": r["X praticata"], "y": r["Y comunicata"],
            "ancora": "left" if a_destra else "right",
            "py": a_px_y(r["Y comunicata"]), "sx": sinistra, "dx": sinistra + larg,
        })

    # i marker sono ostacoli quanto le altre etichette: il testo non li attraversa
    ostacoli = [{"tema": r["Tema"], "sx": a_px_x(r["x"]) - RAGGIO_PUNTO,
                 "dx": a_px_x(r["x"]) + RAGGIO_PUNTO, "py_lab": r["py"]}
                for r in righe]

    def libero(r, py):
        for p in piazzate + ostacoli:
            if p.get("tema") == r["Tema"]:      # il proprio punto non e' ostacolo
                continue
            if (abs(py - p["py_lab"]) < PASSO_ETICHETTA
                    and r["sx"] < p["dx"] and p["sx"] < r["dx"]):
                return False
        return True

    # dall'alto verso il basso: ogni etichetta scende finche' trova posto
    righe.sort(key=lambda r: (r["py"], r["x"]))
    piazzate = []
    for r in righe:
        py = max(r["py"], MARGINE)  # niente testo tagliato dal bordo alto
        while py < altezza - MARGINE and not libero(r, py):
            py += PASSO_ETICHETTA
        r["py_lab"] = min(py, altezza - MARGINE)
        piazzate.append(r)

    etichette = pd.DataFrame(piazzate)
    etichette["y_lab"] = etichette["py_lab"].map(a_dato_y)
    etichette["dx_px"] = etichette["ancora"].map(
        {"left": SCARTO_X, "right": -SCARTO_X})
    return etichette


def matrice_chart(df, soglia, modo="light", altezza=460, larghezza=760):
    """df: colonne Tema, X praticata, Y comunicata, Differenza Y-X, Quadrante."""
    plot = df.dropna(subset=["X praticata", "Y comunicata"]).copy()
    if plot.empty:
        return None

    etichette = _posiziona_etichette(plot, altezza, larghezza)

    punti = alt.Chart(plot).mark_circle(
        size=150, opacity=1, color=SERIE[modo]).encode(
        x=_x("X praticata", modo), y=_y("Y comunicata", modo),
        tooltip=[alt.Tooltip("Tema:N"),
                 alt.Tooltip("X praticata:Q", format=".2f", title="X praticata"),
                 alt.Tooltip("Y comunicata:Q", format=".2f", title="Y comunicata"),
                 alt.Tooltip("Differenza Y-X:Q", format="+.2f", title="Scarto Y-X"),
                 alt.Tooltip("Quadrante:N")])

    # linea guida punto -> etichetta, solo dove l'etichetta e' stata spostata
    spostate = etichette[abs(etichette["y_lab"] - etichette["y"]) > 1e-9]
    guide = alt.Chart(spostate).mark_rule(
        color=MUTED, strokeWidth=1, opacity=0.55).encode(
        x=_x("x", modo), y=_y("y", modo), y2="y_lab:Q")

    testi = alt.layer(*[
        alt.Chart(gruppo).mark_text(
            fontSize=12, color=INK[modo], align=lato, baseline="middle",
            dx=int(gruppo["dx_px"].iloc[0])).encode(
            x=_x("x", modo), y=_y("y_lab", modo), text="Tema:N")
        for lato, gruppo in etichette.groupby("ancora") if not gruppo.empty])

    regola = alt.Chart(pd.DataFrame({"v": [soglia]}))
    croce = (regola.mark_rule(strokeDash=[5, 4], color=MUTED, strokeWidth=1)
             .encode(x=_x("v", modo)) +
             regola.mark_rule(strokeDash=[5, 4], color=MUTED, strokeWidth=1)
             .encode(y=_y("v", modo)))

    note = alt.layer(*[
        alt.Chart(pd.DataFrame([{"x": x, "y": y, "t": testo}])).mark_text(
            fontSize=11, fontWeight="bold", color=MUTED, opacity=0.9,
            align=ax, baseline=ay).encode(
            x=_x("x", modo), y=_y("y", modo), text="t:N")
        for x, y, testo, ax, ay in ANGOLI])

    return ((croce + note + guide + punti + testi)
            .properties(height=altezza, width=larghezza)
            .configure_view(strokeWidth=0))
