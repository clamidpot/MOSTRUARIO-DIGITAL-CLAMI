import pandas as pd
from flask import Flask, request, render_template, send_file
import os
from datetime import datetime
import unicodedata
import io
from weasyprint import HTML
from flask import Response

app = Flask(__name__)

# ------------------------------------------
# Função: converte caminho absoluto → /static/
# ------------------------------------------
def caminho_para_static(caminho):
    if not caminho or str(caminho).strip() == "":
        return ""
    caminho = str(caminho).replace("\\", "/")
    if "static/" in caminho:
        idx = caminho.index("static/")
        relativo = caminho[idx:]
        return "/" + relativo if not relativo.startswith("/") else relativo
    return caminho

# ------------------------------------------
# Helpers gerais
# ------------------------------------------
def limpa(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return s if s not in ["", "nan", "None", "NaT"] else ""

def normaliza_fornecedor_to_str(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    try:
        f = float(s)
        i = int(f)
        if abs(f - i) < 1e-9:
            return str(i)
        else:
            sval = str(f)
            return sval.rstrip('0').rstrip('.') if '.' in sval else sval
    except:
        return s

def parse_datas_variadas(serie):
    parsed = pd.to_datetime(serie, errors="coerce", dayfirst=True)
    if parsed.notna().any():
        return parsed
    numeric = pd.to_numeric(serie, errors="coerce")
    if numeric.notna().any():
        try:
            parsed2 = pd.to_datetime(numeric, unit="d", origin="1899-12-30", errors="coerce")
            if parsed2.notna().any():
                return parsed2
        except:
            pass
    out = pd.Series([pd.NaT] * len(serie))
    formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d"]
    for i, val in enumerate(serie):
        if pd.isna(val) or str(val).strip() == "":
            continue
        s = str(val).strip()
        for fmt in formatos:
            try:
                out.iat[i] = pd.to_datetime(datetime.strptime(s, fmt))
                break
            except:
                continue
    return out

def get_row_value(row, *keys):
    for k in keys:
        if k is None:
            continue
        if k in row:
            val = row.get(k)
            if pd.isna(val):
                continue
            return val
    return None

def format_status_data(val):
    if val is None or (isinstance(val, float) and pd.isna(val)) or str(val).strip() == "":
        return ""
    try:
        parsed = parse_datas_variadas(pd.Series([val]))
        if parsed.notna().any():
            dt = parsed.iloc[0]
            if pd.notna(dt):
                return dt.strftime("%d/%m/%Y")
    except:
        pass
    return ""

def remover_acentos(txt):
    if txt is None:
        return ""
    txt = str(txt)
    return ''.join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')

# ------------------------------------------
# Cache automático do Excel (SIMPLES)
# ------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

arquivo = os.path.join(
    BASE_DIR,
    "data",
    "CATALAGO MOSTRUARIO DIGITAL.xlsx"
)

_ultima_modificacao = None
_df_produtos_cache = None
_df_fornecedores_cache = None

def carregar_dados():
    global _ultima_modificacao, _df_produtos_cache, _df_fornecedores_cache

    mod = os.path.getmtime(arquivo)

    if _ultima_modificacao == mod and _df_produtos_cache is not None:
        return _df_produtos_cache, _df_fornecedores_cache

    todas_abas = pd.read_excel(arquivo, sheet_name=None)

    produtos_key = None
    for k in todas_abas.keys():
        if str(k).strip().lower() == "produtos":
            produtos_key = k
            break
    if not produtos_key:
        produtos_key = list(todas_abas.keys())[0]

    df_produtos = todas_abas[produtos_key].copy()

    lista_fornecedores = []
    for nome, df in todas_abas.items():
        if nome == produtos_key or df is None or df.empty:
            continue
        lista_fornecedores.append(df.copy())

    df_fornecedores = (
        pd.concat(lista_fornecedores, ignore_index=True, sort=False)
        if lista_fornecedores else pd.DataFrame()
    )

    df_produtos.columns = df_produtos.columns.astype(str).str.strip().str.upper()
    df_fornecedores.columns = df_fornecedores.columns.astype(str).str.strip().str.upper()

    for c in ["FORNECEDOR", "MARCA", "PRODUTO"]:
        if c in df_produtos.columns:
            df_produtos[c] = df_produtos[c].ffill()

    for col in ["FORNECEDOR", "MARCA", "PRODUTO", "ACABAMENTO", "IMAGEM PRODUTO"]:
        if col in df_produtos.columns:
            df_produtos[col] = df_produtos[col].apply(
                lambda x: "" if pd.isna(x) else str(x).strip()
            )

    if "FORNECEDOR" in df_produtos.columns:
        df_produtos["FORNECEDOR_STR"] = df_produtos["FORNECEDOR"].apply(normaliza_fornecedor_to_str)

    if not df_fornecedores.empty and "FORNECEDOR" in df_fornecedores.columns:
        df_fornecedores["FORNECEDOR_STR"] = df_fornecedores["FORNECEDOR"].apply(normaliza_fornecedor_to_str)

    _ultima_modificacao = mod
    _df_produtos_cache = df_produtos
    _df_fornecedores_cache = df_fornecedores

    print("📊 Excel recarregado automaticamente")

    return df_produtos, df_fornecedores

# ------------------------------------------
# RUN
# ------------------------------------------
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True,
        use_reloader=False
    )
