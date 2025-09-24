# app.py – visualizador simples sem login
import streamlit as st
import importlib
import os

# Configuração da página
st.set_page_config(page_title="Dashboard Gestão", layout="wide")

# ─── MENU LATERAL ────────────────────────────────
st.sidebar.image("imagens/logo.png", use_container_width=True)
st.sidebar.markdown("<br>", unsafe_allow_html=True)

paginas = {
    "📌 NDs Realizadas": "nds_realizadas",
    "👤 Analistas": "analistas",
    "📈 Gráficos NDs": "graficos_nds",
}


pagina_escolhida = st.sidebar.radio("Página:", list(paginas.keys()))

# ─── CARREGA A PÁGINA ─────────────────────────────
modulo = importlib.import_module(f"paginas.{paginas[pagina_escolhida]}")
modulo.app()
