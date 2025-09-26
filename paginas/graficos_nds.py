import streamlit as st
import pandas as pd
import plotly.express as px
import os
import sys
import plotly.io as pio

from io import BytesIO

# Adiciona a pasta raiz ao path para importar funções compartilhadas
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from funcoes_compartilhadas.trata_tabelas import carregar_aba

# Trata os valores numéricos da planilha
def limpa_valor(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        valor = valor.replace("R$", "").replace(".", "").replace(",", ".").strip()
        try:
            return float(valor)
        except:
            return None
    return None

def app():
    st.title("📊 Reclassificações por Área")

    # ──────── CARREGA E TRATA DADOS ─────────────────────────────
    df = carregar_aba("Controle NDs")
    df = df.dropna(subset=["Valor", "Período ND", "Área", "Projeto"])
    df["Valor"] = df["Valor"].apply(limpa_valor)
    df["Período ND"] = df["Período ND"].astype(str)

    # ──────── FILTROS ─────────────────────────────
    st.subheader("Filtros")
    col1, col2, col3 = st.columns(3)

    with col1:
        filtro_periodo = st.multiselect(
            "Período ND",
            options=sorted(df["Período ND"].unique()),
            key="filtro_periodo"
        )

    with col2:
        filtro_area = st.multiselect(
            "Área",
            options=sorted(df["Área"].unique()),
            key="filtro_area"
        )

    with col3:
        filtro_projeto = st.multiselect(
            "Projeto",
            options=sorted(df["Projeto"].dropna().unique()),
            key="filtro_projeto"
        )

    # ──────── APLICA FILTROS ─────────────────────────────
    if filtro_periodo:
        df = df[df["Período ND"].isin(filtro_periodo)]
    if filtro_area:
        df = df[df["Área"].isin(filtro_area)]
    if filtro_projeto:
        df = df[df["Projeto"].isin(filtro_projeto)]

    # ──────── AGRUPAMENTO ─────────────────────────────
    grafico = df.groupby(["Período ND", "Área"])["Valor"].sum().reset_index()
    grafico["Texto"] = grafico["Valor"].map(
        lambda x: f"R$ {x:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
    )

    # ──────── CRIA GRÁFICO ─────────────────────────────
    fig = px.bar(
        grafico,
        x="Período ND",
        y="Valor",
        color="Área",
        barmode="group",
        text="Texto",
        color_discrete_sequence=px.colors.qualitative.Plotly  # mantém cores padrão
    )

    fig.update_traces(
        textposition="outside",
        textfont=dict(size=13, color="black", family="Arial Black"),
        insidetextanchor="middle"
    )

    fig.update_layout(
        xaxis_title="Período ND",
        yaxis_title="Valor",
        xaxis_tickfont=dict(size=13, family="Arial Black"),  # negrito eixo X
        yaxis_tickformat=',.0f',
        legend_title_text='',
        legend_font=dict(size=13, family="Arial Black"),
        uniformtext_minsize=10,
        uniformtext_mode='hide',
        plot_bgcolor='white'  # fundo branco no gráfico (bom pra exportar)
    )

    # ──────── EXIBE GRÁFICO NO APP ─────────────────────────────
    st.subheader("Reclassificações por Área ao longo do tempo")
    st.plotly_chart(fig, use_container_width=True)

    # ──────── BOTÕES DE EXPORTAÇÃO ─────────────────────────────
    buffer_png = BytesIO()
    buffer_pdf = BytesIO()

    # Exporta com o Kaleido para PNG e PDF
    pio.write_image(fig, buffer_png, format="png", width=1000, height=600, scale=2)
    pio.write_image(fig, buffer_pdf, format="pdf", width=1000, height=600, scale=2)

    st.download_button("📸 Baixar Gráfico (PNG)", data=buffer_png.getvalue(), file_name="grafico_reclassificacoes.png", mime="image/png")
    st.download_button("📄 Baixar Gráfico (PDF)", data=buffer_pdf.getvalue(), file_name="grafico_reclassificacoes.pdf", mime="application/pdf")
