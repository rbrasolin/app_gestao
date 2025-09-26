# funcoes_compartilhadas/trata_tabelas.py

import pandas as pd
import openpyxl

# Função para carregar uma aba da planilha principal
def carregar_aba(nome_aba: str, caminho_arquivo: str = "APP Gestão.xlsx") -> pd.DataFrame:
    """
    Lê uma aba da planilha Excel e retorna como DataFrame.
    """
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=nome_aba)
        return df
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")
    except ValueError:
        raise ValueError(f"Aba '{nome_aba}' não encontrada em {caminho_arquivo}")
    except Exception as e:
        raise RuntimeError(f"Erro ao carregar a aba '{nome_aba}': {e}")

# ✅ Nova função para salvar dados na aba
def salvar_em_aba(df: pd.DataFrame, nome_aba: str, caminho_arquivo: str = "APP Gestão.xlsx"):
    """
    Salva um DataFrame em uma aba da planilha, substituindo o conteúdo anterior.
    """
    try:
        with pd.ExcelWriter(caminho_arquivo, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=nome_aba, index=False)
    except Exception as e:
        raise RuntimeError(f"Erro ao salvar na aba '{nome_aba}': {e}")
