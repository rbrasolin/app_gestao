# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os
import sys
from io import BytesIO

# Caminho para importar funções compartilhadas
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from funcoes_compartilhadas.trata_tabelas import carregar_aba

CAMINHO_ARQUIVO = "APP Gestão.xlsx"
ABA = "Analistas EIC²"

def formatar_valor(v):
    """Formata número em R$ brasileiro"""
    return f"R$ {v:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def formatar_inteiro(v):
    """Formata número inteiro com separador de milhar"""
    return f"{int(v):,}".replace(",", ".") if isinstance(v, (int, float)) else v

def cor_saldo(val):
    """Define cor condicional: positivo vermelho, negativo verde"""
    try:
        num = float(val.replace("R$", "").replace(".", "").replace(",", ".").strip())
        if num > 0:
            return "color: red; font-weight: bold;"
        elif num < 0:
            return "color: green; font-weight: bold;"
    except:
        return ""
    return ""

def estilo_total(row):
    """Aplica estilo especial na linha de total"""
    if str(row.iloc[0]).startswith("Total"):
        return ["background-color: #f0f0f0; font-weight: bold;"] * len(row)
    return [""] * len(row)

def exportar_excel(df, nome_arquivo):
    """Cria botão de download Excel"""
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    st.download_button(
        label="📥 Exportar Excel",
        data=buffer,
        file_name=nome_arquivo if nome_arquivo.endswith(".xlsx") else nome_arquivo + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

def montar_tabela(df, titulo, multiplicadores=None):
    """Gera tabela de análise (mensal ou anual)"""
    st.subheader(titulo)

    colunas_num = ["Horas Base", "Custo Mensal GS", "Valor Recobrado", "Valor Capitalizado"]
    df = df.copy()
    for col in colunas_num:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Multiplicadores personalizados
    if multiplicadores:
        for col, mult in multiplicadores.items():
            if col in df.columns:
                df[col] = df[col] * mult

    # Cria Saldo
    df["Saldo"] = df["Custo Mensal GS"] - df["Valor Recobrado"] - df["Valor Capitalizado"]

    # Monta tabela
    analise = df[["Analista", "Horas Base", "Custo Mensal GS", "Valor Recobrado", "Valor Capitalizado", "Saldo"]].copy()

    # Linha de totais
    total = {
        "Analista": f"Total ({len(analise)})",
        "Horas Base": analise["Horas Base"].sum(),
        "Custo Mensal GS": analise["Custo Mensal GS"].sum(),
        "Valor Recobrado": analise["Valor Recobrado"].sum(),
        "Valor Capitalizado": analise["Valor Capitalizado"].sum(),
        "Saldo": analise["Saldo"].sum(),
    }

    # Totais sempre como primeira linha
    analise = pd.concat([pd.DataFrame([total]), analise], ignore_index=True)

    # Formatação
    analise["Horas Base"] = analise["Horas Base"].apply(formatar_inteiro)
    for col in ["Custo Mensal GS", "Valor Recobrado", "Valor Capitalizado", "Saldo"]:
        analise[col] = analise[col].apply(lambda x: formatar_valor(x) if isinstance(x, (int, float)) else x)

    # Estilos
    styled = analise.style.applymap(cor_saldo, subset=["Saldo"])
    styled = styled.apply(estilo_total, axis=1)

    st.dataframe(styled, use_container_width=True)
    exportar_excel(analise, titulo.replace(" ", "_") + ".xlsx")

def app():
    st.title("👤 Analistas")

    # ─── Carregar dados ─────────────────────────────────────────────
    try:
        df = carregar_aba(ABA)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return

    if df.empty:
        st.warning("Planilha sem dados.")
        return

    # Remove coluna sequencial automática (se existir)
    if df.columns[0].lower().startswith("unnamed") or df.iloc[:, 0].is_monotonic_increasing:
        df.drop(columns=[df.columns[0]], inplace=True)

    # ─── Filtros ───────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    colunas_filtro = ["Área", "Analista", "Cargo"]
    filtros = {}

    for i, coluna in enumerate(colunas_filtro):
        with [col1, col2, col3][i % 3]:
            chave = f"filtro_{coluna}"
            opcoes = sorted(df[coluna].dropna().astype(str).unique())
            valores_padrao = [v for v in st.session_state.get(chave, []) if v in opcoes]

            filtros[chave] = st.multiselect(
                f"{coluna}:",
                options=opcoes,
                default=valores_padrao,
                key=chave,
            )

    df_filtrado = df.copy()
    for chave, valores in filtros.items():
        if valores:
            coluna = chave.replace("filtro_", "")
            df_filtrado = df_filtrado[df_filtrado[coluna].astype(str).isin(valores)]

    # ─── Análises ───────────────────────────────────────────────────
    montar_tabela(df_filtrado, "📊 Análise Custos (Ano)", {
        "Horas Base": 11,
        "Custo Mensal GS": 12,
        "Valor Recobrado": 12,
        "Valor Capitalizado": 11
    })

    montar_tabela(df_filtrado, "📊 Análise Custos Analistas (Mês)")

    # ─── Detalhamento dos Analistas ────────────────────────────────
    st.subheader("📋 Detalhamento dos Analistas")
    if "Ativo" in df_filtrado.columns:
        df_filtrado = df_filtrado.drop(columns=["Ativo"])

    if "Horas Base" in df_filtrado.columns:
        df_filtrado["Horas Base"] = pd.to_numeric(df_filtrado["Horas Base"], errors="coerce").fillna(0).astype(int)
        df_filtrado["Horas Base"] = df_filtrado["Horas Base"].apply(formatar_inteiro)

    colunas_moeda = ["Custo Mensal GS", "Valor Recobrado", "Valor Capitalizado", "Valor Hora"]
    for col in colunas_moeda:
        if col in df_filtrado.columns:
            df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors="coerce").fillna(0)
            df_filtrado[col] = df_filtrado[col].apply(formatar_valor)

    # Linha total
    if not df_filtrado.empty:
        total_row = {col: "" for col in df_filtrado.columns}
        total_row["Analista"] = f"Total ({len(df_filtrado)})"
        if "Horas Base" in df_filtrado.columns:
            total_row["Horas Base"] = formatar_inteiro(int(df_filtrado["Horas Base"].str.replace(".", "").astype(int).sum()))
        for col in colunas_moeda:
            if col in df_filtrado.columns:
                # remove R$, separadores e soma
                total_row[col] = formatar_valor(
                    pd.to_numeric(
                        df_filtrado[col].str.replace("R$", "").str.replace(".", "").str.replace(",", "."),
                        errors="coerce"
                    ).sum()
                )
        df_filtrado = pd.concat([pd.DataFrame([total_row]), df_filtrado], ignore_index=True)

    styled_det = df_filtrado.style.apply(estilo_total, axis=1)
    st.dataframe(styled_det, use_container_width=True)
    exportar_excel(df_filtrado, "Detalhamento_Analistas.xlsx")
