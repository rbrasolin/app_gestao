# paginas/nds_realizadas.py
# -*- coding: utf-8 -*-
"""
Página NDs Realizadas - refatorada
- Detecta colunas de período (ND, Alocação/Fechamento) de forma robusta
- Gera colunas auxiliares formatadas em "Mês/AA" (PT-BR)
- Cria filtros que usam opções ordenadas cronologicamente
- Aplica filtros de forma consistente em todos os visuais
- Conserta casos estranhos como "Conjunto/25" interpretando strings
"""

import streamlit as st
import pandas as pd
import os
import sys
import unicodedata
import re
from typing import Tuple, Optional

# Permite importar funcoes_compartilhadas a partir da estrutura do projeto
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from funcoes_compartilhadas.trata_tabelas import carregar_aba

# Config (nome da aba no Excel)
ABA = "Controle NDs"


# ---------------------------
# Utilitários
# ---------------------------
def _normalize(text: str) -> str:
    """Normaliza texto: remove acentos, pontuação e deixa em minúsculas contínuas.
       Usado para identificar colunas por palavras-chave."""
    s = str(text or "")
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    s = re.sub(r'[^a-z0-9]', '', s.lower())
    return s


# Mapeamento fixo de meses em PT-BR (abreviados)
_MESES_PT = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"
}

# Palavras-chave possíveis para identificar colunas de período
_KEYWORDS_ND = ["periodo", "nd"]
_KEYWORDS_ALOC = ["periodo", "aloc", "alocacao", "alocação"]
_KEYWORDS_FECH = ["periodo", "fech", "fechamento"]


def _detect_period_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Detecta, se existir, as colunas que representam:
      - Período ND
      - Período Alocação
      - Período Fechamento
    Retorna os nomes reais das colunas no DataFrame (ou None).
    """
    cols = list(df.columns)
    nd_col = None
    aloc_col = None
    fech_col = None

    # tenta match exato (case-insensitive)
    for c in cols:
        n = c.strip().lower()
        if n == "período nd" or n == "periodo nd" or ( "nd" in _normalize(c) and "periodo" in _normalize(c) and not nd_col):
            nd_col = c
        if n == "período alocação" or n == "periodo alocacao" or ("aloc" in _normalize(c) and "periodo" in _normalize(c) and not aloc_col):
            aloc_col = c
        if n == "período fechamento" or n == "periodo fechamento" or ("fech" in _normalize(c) and "periodo" in _normalize(c) and not fech_col):
            fech_col = c

    # se não encontrou por igualdade, faz busca por keywords
    if not nd_col:
        for c in cols:
            n = _normalize(c)
            if all(k in n for k in _KEYWORDS_ND):
                nd_col = c
                break
    if not aloc_col:
        for c in cols:
            n = _normalize(c)
            if any(k in n for k in _KEYWORDS_ALOC) and "nd" not in n:
                aloc_col = c
                break
    if not fech_col:
        for c in cols:
            n = _normalize(c)
            if any(k in n for k in _KEYWORDS_FECH) and "nd" not in n:
                fech_col = c
                break

    return nd_col, aloc_col, fech_col


def _extract_year_month_from_string(s: str) -> Tuple[Optional[int], Optional[int]]:
    """
    Tenta extrair (ano, mês) de uma string com formatos variados.
    Ex.: "Set/25", "Set-25", "Conjunto/25", "2025-09", "09/2025" -> (2025, 9)
    Retorna (None, None) se não for possível.
    """
    if not isinstance(s, str):
        return None, None
    raw = s.strip().lower()
    if not raw:
        return None, None

    # remove acentos e normaliza
    raw_norm = unicodedata.normalize('NFKD', raw)
    raw_norm = ''.join(ch for ch in raw_norm if unicodedata.category(ch) != 'Mn')
    raw_norm = re.sub(r'[^a-z0-9/.\- ]', '', raw_norm)

    # procura ano com 4 dígitos
    m = re.search(r'(20\d{2})', raw_norm)
    year = None
    if m:
        year = int(m.group(1))
    else:
        # procura 2 dígitos no final
        m2 = re.search(r'(\d{2})$', raw_norm)
        if m2:
            year = 2000 + int(m2.group(1))

    # procura mês por palavras-chave (pt-br e variações)
    month = None
    keywords_map = {
        1: ["jan", "janeiro"],
        2: ["fev", "fevereiro"],
        3: ["mar", "março", "marco"],
        4: ["abr", "abril"],
        5: ["mai", "maio"],
        6: ["jun", "junho"],
        7: ["jul", "julho"],
        8: ["ago", "agost", "agosto"],
        9: ["set", "setem", "setembro", "sep", "conj", "conjun", "conjunto"],
        10: ["out", "outub", "outubro"],
        11: ["nov", "novembro"],
        12: ["dez", "dezembro"]
    }
    for mnum, kws in keywords_map.items():
        for kw in kws:
            if kw in raw_norm:
                month = mnum
                break
        if month:
            break

    # se não encontrou por palavras, tenta extrair padrão numérico (MM ou MM/AAAA)
    if month is None:
        mnum = re.search(r'(?:(?:^|\D)(0?[1-9]|1[0-2])(?:\D|$))', raw_norm)
        if mnum:
            month = int(mnum.group(1))

    return year, month


def _format_period_from_maybe_date_or_string(val) -> Optional[str]:
    """
    Recebe um valor que pode ser datetime, string ou numérico e tenta
    devolver o rótulo 'Mês/AA' em PT-BR. Retorna None se não puder.
    """
    # 1) Tenta via pd.to_datetime
    try:
        dt = pd.to_datetime(val, errors='coerce')
        if pd.notna(dt):
            y = dt.year
            m = dt.month
            return f"{_MESES_PT.get(m, str(m))}/{str(y)[-2:]}"
    except Exception:
        pass

    # 2) Se não parseou, tenta extrair da string
    if isinstance(val, str):
        year, month = _extract_year_month_from_string(val)
        if year and month:
            return f"{_MESES_PT.get(month, str(month))}/{str(year)[-2:]}"
    return None


def _format_period_series(series: pd.Series) -> pd.Series:
    """
    Cria uma série formatada 'Mês/AA' a partir de uma série que pode
    conter datas ou strings misturadas.
    """
    # Aplicamos a função linha a linha para garantir robustez
    return series.apply(_format_period_from_maybe_date_or_string)


def _period_options_ordered_from_series(series: pd.Series) -> list:
    """
    Gera uma lista ordenada cronologicamente de rótulos 'Mês/AA' a partir
    de uma série que pode ter datas ou texto. Evita rótulos estranhos.
    """
    labels = _format_period_series(series)
    # monta tuplas (year, month, label) quando possível, e separa rotulos sem data
    tuples = []
    others = set()
    for i, lbl in labels.items():
        if pd.isna(lbl) or lbl is None:
            # se não há label formatado, tenta extrair de texto original
            y, m = _extract_year_month_from_string(str(series.iat[i]))
            if y and m:
                tuples.append((y, m, f"{_MESES_PT.get(m, str(m))}/{str(y)[-2:]}"))
            else:
                v = series.iat[i]
                if pd.notna(v):
                    others.add(str(v))
        else:
            # converte label para (year,month)
            m_part, y_part = lbl.split("/")
            # tenta obter number do mês via lookup
            month_num = None
            for k, v in _MESES_PT.items():
                if v.lower() == m_part.lower():
                    month_num = k
                    break
            try:
                year_num = 2000 + int(y_part) if len(y_part) == 2 else int(y_part)
            except Exception:
                year_num = None
            if year_num and month_num:
                tuples.append((year_num, month_num, lbl))
            else:
                others.add(lbl)

    # dedup e ordenar
    unique = {(y, m, l) for (y, m, l) in tuples}
    ordered = sorted(unique, key=lambda t: (t[0], t[1]))
    result = [t[2] for t in ordered]

    # acrescenta itens não parseáveis ao final (ordenados lexicograficamente)
    if others:
        result += sorted(others)

    return result


# ---------------------------
# Função principal da página
# ---------------------------
def app():
    st.title("📝 NDs Realizadas")

    # Carrega a aba do Excel usando função compartilhada
    try:
        df = carregar_aba(ABA)
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return

    if df is None or df.empty:
        st.warning("Planilha sem dados.")
        return

    # Remove coluna sequencial automática, se existir
    try:
        first_col = df.columns[0]
        if first_col.lower().startswith("unnamed") or df.iloc[:, 0].is_monotonic_increasing:
            df = df.drop(columns=[first_col])
    except Exception:
        # se algo estranho acontecer, ignora e segue
        pass

    # Detecta colunas de período (ND, Alocacao, Fechamento) de forma robusta
    col_periodo_nd, col_periodo_aloc, col_periodo_fech = _detect_period_columns(df)

    # Cria colunas auxiliares formatadas 'Mês/AA' com nomes padronizados
    if col_periodo_nd:
        df["Período ND_fmt"] = _format_period_series(df[col_periodo_nd])
    if col_periodo_aloc:
        df["Período Aloc_fmt"] = _format_period_series(df[col_periodo_aloc])
    if col_periodo_fech:
        df["Período Fechamento_fmt"] = _format_period_series(df[col_periodo_fech])

    # --- Configura os filtros (rótulos e colunas usadas)
    st.subheader("Filtros")
    # lista de (rótulo_exibicao, coluna_para_filtrar, is_period_flag)
    filtros_config = [
        ("Período ND", "Período ND_fmt", True),
        ("Período Fechamento", "Período Fechamento_fmt", True),
        ("Área", "Área", False),
        ("Analista", "Analista", False),
        ("Projeto", "Projeto", False),
        ("Conta Contábil", "Conta Contábil", False),
        ("Status Portal", "Status Portal", False),
        ("ND", "ND", False),
    ]

    # UI: 4 colunas para filtros
    col1, col2, col3, col4 = st.columns(4)
    ui_cols = [col1, col2, col3, col4]

    # Guarda os valores selecionados por coluna real
    filtros_selected = {}

    # Percorre a configuração e monta os multiselects apenas para colunas existentes
    for idx, (label, col_key, is_period) in enumerate(filtros_config):
        # só cria o filtro se a coluna existir no DataFrame
        if col_key not in df.columns:
            continue

        with ui_cols[idx % 4]:
            if is_period:
                # monta opções ordenadas cronologicamente a partir da série original
                # (evita strings estranhas que já existiam no Excel)
                # Para construir as opções usamos a série original, não a *_fmt, para capturar todos os casos
                if col_key == "Período ND_fmt":
                    base_series = df[col_periodo_nd] if col_periodo_nd in df.columns else pd.Series([], dtype=object)
                elif col_key == "Período Aloc_fmt":
                    base_series = df[col_periodo_aloc] if col_periodo_aloc in df.columns else pd.Series([], dtype=object)
                elif col_key == "Período Fechamento_fmt":
                    base_series = df[col_periodo_fech] if col_periodo_fech in df.columns else pd.Series([], dtype=object)
                else:
                    base_series = df[col_key]  # fallback

                opcoes = _period_options_ordered_from_series(base_series)
                filtros_selected[col_key] = st.multiselect(f"{label}:", options=opcoes, key=f"filtro_{col_key}")
            else:
                # opções padrão para colunas texto/numéricas
                opcoes = sorted(df[col_key].dropna().astype(str).unique())
                filtros_selected[col_key] = st.multiselect(f"{label}:", options=opcoes, key=f"filtro_{col_key}")

    # --- Aplica os filtros no DataFrame
    df_filtrado = df.copy()
    for col_key, valores in filtros_selected.items():
        if not valores:
            continue
        # se for coluna de período (termina com _fmt), filtra por essa coluna formatada
        if str(col_key).endswith("_fmt"):
            df_filtrado = df_filtrado[df_filtrado[col_key].astype(str).isin(valores)]
        else:
            df_filtrado = df_filtrado[df_filtrado[col_key].astype(str).isin(valores)]

    # --- Cálculos e métricas a partir de df_filtrado
    # Garante colunas numéricas tratadas com segurança
    if "Horas" in df_filtrado.columns:
        df_filtrado["Horas"] = pd.to_numeric(df_filtrado["Horas"], errors="coerce")
    else:
        df_filtrado["Horas"] = 0

    if "Valor" in df_filtrado.columns:
        df_filtrado["Valor"] = pd.to_numeric(df_filtrado["Valor"], errors="coerce")
    else:
        df_filtrado["Valor"] = 0

    total_horas = int(df_filtrado["Horas"].sum(skipna=True) or 0)
    total_valor = float(df_filtrado["Valor"].sum(skipna=True) or 0.0)

    # Exibe métricas principais
    col_a, col_b = st.columns(2)
    col_a.metric("⏱️ Total de Horas", f"{total_horas:,}".replace(",", "."))
    col_b.metric("💰 Total de Valor", f"R$ {total_valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", "."))

    # --- Totais por Conta Contábil (se existir)
    if "Conta Contábil" in df_filtrado.columns:
        df_totais = (
            df_filtrado
            .groupby("Conta Contábil", dropna=False)["Valor"]
            .sum()
            .reset_index()
        )
        df_totais["Total R$"] = df_totais["Valor"].map(
            lambda x: f"R$ {x:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
        )
        df_totais = df_totais.drop(columns=["Valor"])
        st.subheader("📊 Totais por Conta Contábil")
        st.dataframe(df_totais.style.set_properties(**{"font-weight": "bold"}), use_container_width=True)

    # --- Preparar tabela para visualização final
    colunas_ocultas = ["Empresa", "Conta Destino", "Aprovador"]
    df_visual = df_filtrado.drop(columns=[c for c in colunas_ocultas if c in df_filtrado.columns])

    # Substitui as colunas originais de período pelas versões formatadas (se existirem)
    # Mantém rótulos padronizados na exibição final
    if "Período ND_fmt" in df_visual.columns:
        df_visual["Período ND"] = df_visual["Período ND_fmt"]
    elif col_periodo_nd and col_periodo_nd in df_visual.columns:
        df_visual["Período ND"] = _format_period_series(df_visual[col_periodo_nd])

    if "Período Fechamento_fmt" in df_visual.columns:
        # se a coluna de fechamento formatada existe, usa ela
        df_visual["Período Fechamento"] = df_visual["Período Fechamento_fmt"]
    elif "Período Aloc_fmt" in df_visual.columns:
        # caso seu banco chame de 'Período Alocação' (nome antigo), mapeia para o rótulo 'Período Fechamento'
        df_visual["Período Fechamento"] = df_visual["Período Aloc_fmt"]
    elif col_periodo_fech and col_periodo_fech in df_visual.columns:
        df_visual["Período Fechamento"] = _format_period_series(df_visual[col_periodo_fech])
    elif col_periodo_aloc and col_periodo_aloc in df_visual.columns:
        # fallback: se só existir coluna aloc no banco, usa ela renomeada
        df_visual["Período Fechamento"] = _format_period_series(df_visual[col_periodo_aloc])

    # Remove colunas auxiliares *_fmt antes de exibir (mantém apenas rótulos limpos)
    fmt_cols = [c for c in df_visual.columns if c.endswith("_fmt")]
    df_visual = df_visual.drop(columns=[c for c in fmt_cols if c in df_visual.columns], errors='ignore')

    # Reordena colunas para garantir que 'Período ND' e 'Período Fechamento' apareçam no começo (se existir)
    cols = list(df_visual.columns)
    preferred = []
    for p in ["Período ND", "Período Fechamento"]:
        if p in cols:
            preferred.append(p)
            cols.remove(p)
    df_visual = df_visual[preferred + cols]

    # Exibe tabela final
    st.subheader("📋 Detalhamento das NDs")
    st.dataframe(df_visual.style.set_properties(**{"font-weight": "bold"}), use_container_width=True)

    # Botão de download (gera excel com a visualização atual)
    st.download_button(
        label="📥 Baixar Excel",
        data=_to_excel_bytes(df_visual),
        file_name="nds_filtradas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# utilitário local para transformar df em bytes do excel
def _to_excel_bytes(df: pd.DataFrame):
    from io import BytesIO
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer
