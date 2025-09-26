# paginas/nds_realizadas.py
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import os
import sys
import unicodedata
import re

# Adiciona caminho para importar a função de leitura compartilhada
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from funcoes_compartilhadas.trata_tabelas import carregar_aba

# Caminho e aba da planilha
CAMINHO_ARQUIVO = "APP Gestão.xlsx"
ABA = "Controle NDs"


# ---------------------------
# Utilitários
# ---------------------------
def _normalize(text: str) -> str:
    """Remove acentos e caracteres especiais, deixa em minúsculas sem espaço.
       Usado para comparar nomes de colunas de forma robusta."""
    s = str(text)
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')  # remove acentos
    s = re.sub(r'[^a-z0-9]', '', s.lower())
    return s


def _find_col_by_keywords(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """Encontra a primeira coluna do df cujo nome normalizado contenha
       todas as 'keywords' fornecidas (keywords já devem estar normalizadas)."""
    for col in df.columns:
        n = _normalize(col)
        if all(k in n for k in keywords):
            return col
    return None


def _format_mes_ano(series: pd.Series) -> pd.Series:
    """Converte série (datas/strings) em rótulos 'Abrev/AA' em PT-BR.
       Exemplos: 2025-09-01 -> 'Set/25'"""
    dt = pd.to_datetime(series, errors='coerce')
    meses = {
        1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
        7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"
    }

    resultado = pd.Series([None] * len(dt), index=dt.index, dtype="object")
    valid = dt.notna()
    if valid.any():
        meses_str = dt.dt.month.map(meses)
        anos = dt.dt.year
        # monta "Mês/AA" apenas para linhas válidas
        resultado.loc[valid] = meses_str[valid].astype(str) + "/" + anos[valid].astype(int).astype(str).str.zfill(2)

    return resultado


def _period_options_ordered(series: pd.Series) -> list[str]:
    """Retorna lista de rótulos (Mês/AA) ordenada cronologicamente para as opções do filtro."""
    dt = pd.to_datetime(series, errors='coerce')
    # extrai tuplas únicas (ano, mês) e ordena cronologicamente
    tuplas = {(int(d.year), int(d.month)) for d in dt.dropna()}
    ordered = sorted(tuplas)
    meses = {
        1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
        7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"
    }
    return [f"{meses[m]}/{str(y%100).zfill(2)}" for (y, m) in ordered]


# ---------------------------
# Página
# ---------------------------
def app():
    st.title("📝 NDs Realizadas")

    # ──────── CARREGA PLANILHA ─────────────────────────────
    try:
        df = carregar_aba(ABA)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return

    if df.empty:
        st.warning("Planilha sem dados.")
        return

    # Remove coluna sequencial automática, se existir
    if df.columns[0].lower().startswith("unnamed") or df.iloc[:, 0].is_monotonic_increasing:
        df.drop(columns=[df.columns[0]], inplace=True)

    # ──────── DETECÇÃO DE COLUNAS DE PERÍODO (robusta)
    # tenta identificar as colunas mesmo que haja variações no nome
    col_periodo_nd = _find_col_by_keywords(df, ["periodo", "nd"])
    col_periodo_aloc = _find_col_by_keywords(df, ["periodo", "aloc"]) or _find_col_by_keywords(df, ["periodo", "alocacao"])
    col_periodo_fech = _find_col_by_keywords(df, ["periodo", "fech"]) or _find_col_by_keywords(df, ["periodo", "fechamento"])

    # ──────── FILTROS ─────────────────────────────
    st.subheader("Filtros")

    # lista original de filtros (mantive mesma ordem)
    colunas_filtro = [
        "Período ND", "Período Fechamento", "Área", "Analista",
        "Projeto", "Conta Contábil", "Status Portal", "ND"
    ]

    col1, col2, col3, col4 = st.columns(4)
    filtros = {}
    colunas_ui = [col1, col2, col3, col4]

    # preenche colunas disponíveis checando correspondência robusta para os nomes esperados
    col_map = []  # lista de (label_exibicao, nome_coluna_real)
    for label in colunas_filtro:
        # prioridade: coluna com nome idêntico (case-insensitive)
        found = None
        for c in df.columns:
            if c.strip().lower() == label.strip().lower():
                found = c
                break
        if not found:
            # especial: se o label fala de "Período ND"/"Período Fechamento", usamos as colunas detectadas
            norm_label = _normalize(label)
            if "nd" in norm_label and col_periodo_nd:
                found = col_periodo_nd
            elif "fech" in norm_label and col_periodo_fech:
                found = col_periodo_fech
            else:
                # tenta achar por palavras do label
                palavras = [w for w in re.split(r'\W+', label) if len(w) > 1]
                palavras_norm = [_normalize(w) for w in palavras]
                found = _find_col_by_keywords(df, palavras_norm)
        if found:
            col_map.append((label, found))

    # monta os multiselects usando os nomes reais das colunas (found) para filtrar os dados
    for idx, (label, real_col) in enumerate(col_map):
        with colunas_ui[idx % 4]:
            # Se for coluna de período, mostramos opções já formatadas (Mês/AA) e ordenadas cronologicamente
            if real_col in (col_periodo_nd, col_periodo_aloc, col_periodo_fech):
                opcoes = _period_options_ordered(df[real_col])
            else:
                opcoes = sorted(df[real_col].dropna().astype(str).unique())
            filtros[real_col] = st.multiselect(f"{label}:", options=opcoes, key=f"filtro_{real_col}")

    # ──────── APLICA OS FILTROS (mantendo lógica original)
    df_filtrado = df.copy()
    for real_col, valores in filtros.items():
        if not valores:
            continue
        # se coluna é de período -> compara pelos rótulos formatados (compatível com opções exibidas)
        if real_col in (col_periodo_nd, col_periodo_aloc, col_periodo_fech):
            rotulos = _format_mes_ano(df_filtrado[real_col])
            df_filtrado = df_filtrado[rotulos.isin(valores)]
        else:
            df_filtrado = df_filtrado[df_filtrado[real_col].astype(str).isin(valores)]

    # ──────── TRATA E MOSTRA MÉTRICAS ─────────────────────────────
    # (mantive exatamente sua lógica)
    df_filtrado["Horas"] = pd.to_numeric(df_filtrado["Horas"], errors="coerce")
    df_filtrado["Valor"] = pd.to_numeric(df_filtrado["Valor"], errors="coerce")

    total_horas = df_filtrado["Horas"].sum()
    total_valor = df_filtrado["Valor"].sum()

    col_a, col_b = st.columns(2)
    col_a.metric("⏱️ Total de Horas", f"{int(total_horas):,}".replace(",", "."))
    col_b.metric("💰 Total de Valor", f"R$ {total_valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", "."))

    # ──────── TOTAIS POR CONTA CONTÁBIL ─────────────────────────────
    if "Conta Contábil" in df_filtrado.columns:
        df_totais = (
            df_filtrado
            .groupby("Conta Contábil")["Valor"]
            .sum()
            .reset_index()
        )
        df_totais["Total R$"] = df_totais["Valor"].map(
            lambda x: f"R$ {x:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
        )
        df_totais.drop(columns=["Valor"], inplace=True)

        st.subheader("📊 Totais por Conta Contábil")
        st.dataframe(df_totais.style.set_properties(**{"font-weight": "bold"}), use_container_width=True)

    # ──────── TABELA COMPLETA (sem colunas ocultas) ─────────────────────────────
    colunas_ocultas = ["Empresa", "Conta Destino", "Aprovador"]
    df_visual = df_filtrado.drop(columns=[col for col in colunas_ocultas if col in df_filtrado.columns])

    # Formata colunas de período (ND, Alocação, Fechamento) para exibição
    for real_col in (col_periodo_nd, col_periodo_aloc, col_periodo_fech):
        if real_col and real_col in df_visual.columns:
            df_visual[real_col] = _format_mes_ano(df_visual[real_col])

    st.subheader("📋 Detalhamento das NDs")
    st.dataframe(df_visual.style.set_properties(**{"font-weight": "bold"}), use_container_width=True)

    # Botão de download (gera excel com a visualização atual)
    st.download_button(
        label="📥 Baixar Excel",
        data=to_excel_bytes(df_visual),
        file_name="nds_filtradas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def to_excel_bytes(df):
    from io import BytesIO
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer
