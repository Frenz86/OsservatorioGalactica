# -*- coding: utf-8 -*-
#https://osservatoriogalactica.streamlit.app/
"""
App Streamlit multipagina — Osservatorio Galactica.

Pagine:
  - Compilatore PowerPoint DEIA (pages/deia.py, logica in mod1)
  - Matrice di trasparenza DEIA (pages/matrice_trasparenza.py, logica in mod2)

Avvio:
    pip install -r requirements.txt
    streamlit run app.py
"""
import streamlit as st

st.set_page_config(page_title="Osservatorio Galactica", page_icon="📊", layout="wide")

# --- login ------------------------------------------------------------------
CREDENZIALI = {
    "admin": "admin",
}

if not st.session_state.get("autenticato"):
    st.title("📊 Osservatorio Galactica")
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

pagine = [
    st.Page("pages/deia.py", title="Compilatore PowerPoint DEIA", icon="📊"),
    st.Page("pages/matrice_trasparenza.py", title="Matrice di trasparenza DEIA", icon="🧭"),
]
pg = st.navigation(pagine)
pg.run()
