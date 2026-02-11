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
    parsed = pd.to_datetime(serie, format="%d/%m/%Y", errors="coerce")
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
def formatar_codigo_5(v):
    if v is None or str(v).strip() == "":
        return ""
    try:
        return str(int(float(v))).zfill(5)
    except:
        return str(v).zfill(5)

# ------------------------------------------
# Cache automático do Excel (SIMPLES)
# ------------------------------------------
arquivo = "data/CATALAGO MOSTRUARIO DIGITAL.xlsx"
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

    # Normalizações (iguais às suas)
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
# ROTA PRODUTOS
# ------------------------------------------
@app.route("/produtos")
def produtos():
    df_produtos, df_fornecedores = carregar_dados()
    termo = request.args.get("pesquisa_acabamento", "").strip()
    termo_norm = remover_acentos(termo).lower()
    df = df_produtos.copy()

    if termo_norm != "":
        cols_busca = ["ACABAMENTO", "TIPO DE ACABAMENTO", "TIPO_ACABAMENTO"]
        mask = False
        for col in cols_busca:
            if col in df.columns:
                df[col + "_SEMC"] = df[col].astype(str).apply(remover_acentos).str.lower()
                mask = mask | df[col + "_SEMC"].str.contains(termo_norm, na=False)
        df = df[mask].copy()

    lista = []
    df_unicos = df.groupby("PRODUTO").first().reset_index()

    for _, row in df_unicos.iterrows():
        nome = limpa(row.get("PRODUTO", ""))
        marca = limpa(row.get("MARCA", ""))
        fornecedor_val = limpa(row.get("FORNECEDOR", ""))
        try:
            fornecedor = str(int(float(fornecedor_val)))
        except:
            fornecedor = fornecedor_val
        img = caminho_para_static(row.get("IMAGEM PRODUTO", ""))
        lista.append({
            "nome": nome,
            "marca": marca,
            "imagem": img,
            "fornecedor": fornecedor
        })

    lista.sort(key=lambda x: int(x["fornecedor"]) if str(x["fornecedor"]).isdigit() else x["fornecedor"])
    return render_template("produtos.html", produtos=lista, pesquisa_acabamento=termo)

# ------------------------------------------
# ROTA PRODUTO DETALHES (com acabamentos)
# ------------------------------------------
@app.route("/produto/<nome>")
def detalhes(nome):
    df_produtos, df_fornecedores = carregar_dados()
    df_item = df_produtos[df_produtos["PRODUTO"] == nome]
    if df_item.empty:
        mask = df_produtos["PRODUTO"].astype(str).str.strip().str.lower() == str(nome).strip().lower()
        df_item = df_produtos[mask]

    if df_item.empty:
        return f"Produto '{nome}' não encontrado."

    item = df_item.iloc[0]
    fornecedor_raw = item.get("FORNECEDOR", "")
    fornecedor = normaliza_fornecedor_to_str(fornecedor_raw)
    marca = item.get("MARCA", "") if "MARCA" in item else ""

    # ------------------------------------------
    # PASSO 2 — Acabamentos permitidos do produto
    # ------------------------------------------
    acabamentos_permitidos = (
        df_item["ACABAMENTO"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )

    acabamentos_permitidos = [a.upper() for a in acabamentos_permitidos]

    imagens_produto = []
    if "IMAGEM PRODUTO" in df_item.columns:
        imagens_produto = df_item["IMAGEM PRODUTO"].dropna().unique().tolist()
        imagens_produto = [caminho_para_static(x) for x in imagens_produto if caminho_para_static(x)]

    if not df_fornecedores.empty:
        df_f_copy = df_fornecedores.copy()

        if "FORNECEDOR_STR" not in df_f_copy.columns and "FORNECEDOR" in df_f_copy.columns:
            df_f_copy["FORNECEDOR_STR"] = df_f_copy["FORNECEDOR"].apply(normaliza_fornecedor_to_str)

        if "ACABAMENTO" in df_f_copy.columns:
            df_f_copy["ACABAMENTO"] = (
                df_f_copy["ACABAMENTO"]
                .astype(str)
                .str.strip()
                .str.upper()
            )

        acabamentos_fornecedor = df_f_copy[
            (df_f_copy["FORNECEDOR_STR"] == fornecedor) &
            (df_f_copy["ACABAMENTO"].isin(acabamentos_permitidos))
        ].copy()
    else:
        acabamentos_fornecedor = pd.DataFrame()

    categorias = {}
    for idx, row in acabamentos_fornecedor.iterrows():
        categoria_raw = get_row_value(row, "TIPO DE ACABAMENTO", "TIPO_ACABAMENTO")
        categoria = limpa(categoria_raw) or "OUTROS"
        if categoria not in categorias:
            categorias[categoria] = []

        acabamento_val = limpa(get_row_value(row, "ACABAMENTO"))
        tipo_val = limpa(get_row_value(row, "TIPO DE ACABAMENTO", "TIPO_ACABAMENTO"))
        comp_val = limpa(get_row_value(row, "COMPOSIÇÃO", "COMPOSICAO"))
        status_val = limpa(get_row_value(row, "STATUS"))
        status_data_fmt = format_status_data(get_row_value(row, "STATUS_DATA", "STATUS DATA"))
        restr_val = limpa(get_row_value(row, "RESTRIÇÃO", "RESTRICAO"))
        info_val = limpa(get_row_value(row, "INFORMACAO_COMPLEMENTAR", "INFORMAÇÃO COMPLEMENTAR"))
        img_val = limpa(get_row_value(row, "IMAGEM ACABAMENTO", "IMAGEM"))

        st_norm = status_val.lower().strip()
        for a,b in [("í","i"),("é","e"),("ó","o"),("ú","u"),("ã","a"),("õ","o"),("â","a"),("ê","e")]:
            st_norm = st_norm.replace(a,b)
        if st_norm in ["indisponivel", "indisponível"]:
            status_cor = "#FF0000"
        elif st_norm == "suspenso":
            status_cor = "#D4A017"
        elif st_norm == "ativo":
            status_cor = "#008000"
        else:
            status_cor = "black"

        categorias[categoria].append({
            "ACABAMENTO": acabamento_val,
            "TIPO": tipo_val,
            "COMP": comp_val,
            "STATUS": status_val,
            "STATUS_DATA": status_data_fmt,
            "STATUS_COR": status_cor,
            "RESTR": restr_val,
            "INFO": info_val,
            "IMG": caminho_para_static(img_val) if img_val else ""
        })

    acabamentos_lista = (
        acabamentos_fornecedor["ACABAMENTO"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    ) if "ACABAMENTO" in acabamentos_fornecedor.columns else []

    ultima_atualizacao = "Data não disponível"
    if "ULTIMA_ATUALIZACAO" in acabamentos_fornecedor.columns:
        try:
            series_datas = acabamentos_fornecedor["ULTIMA_ATUALIZACAO"].astype(str).replace("", pd.NA)
            parsed = parse_datas_variadas(series_datas)
            if parsed.notna().any():
                ultima_data = parsed.max()
                if pd.notna(ultima_data):
                    ultima_atualizacao = ultima_data.strftime("%d/%m/%Y")
        except:
            pass

    status_coletados = []
    if "STATUS" in acabamentos_fornecedor.columns:
        for s in acabamentos_fornecedor["STATUS"].dropna().unique().tolist():
            s2 = str(s).strip()
            if s2 and s2 not in status_coletados:
                status_coletados.append(s2)

    return render_template(
        "produto.html",
        nome=nome,
        fornecedor=fornecedor,
        marca=marca,
        imagens_produto=imagens_produto,
        categorias=categorias,
        acabamentos_lista=acabamentos_lista,
        ultima_modificacao=ultima_atualizacao,
        status_coletados=status_coletados
    )

# ------------------------------------------
# ROTA DOWNLOAD PDF – ACABAMENTOS
# ------------------------------------------
@app.route("/download/<nome>")
def download(nome):
    df_produtos, df_fornecedores = carregar_dados()

    # ---------------------------
    # Localiza o produto
    # ---------------------------
    df_item = df_produtos[df_produtos["PRODUTO"] == nome]

    if df_item.empty:
        mask = df_produtos["PRODUTO"].astype(str).str.strip().str.lower() == str(nome).strip().lower()
        df_item = df_produtos[mask]

    if df_item.empty:
        return f"Produto '{nome}' não encontrado."

    # ---------------------------
    # Dados principais do produto
    # ---------------------------
    item = df_item.iloc[0]
    fornecedor = normaliza_fornecedor_to_str(item.get("FORNECEDOR", ""))
    marca = item.get("MARCA", "") if "MARCA" in item else ""

    # ---------------------------
    # Acabamentos permitidos
    # ---------------------------
    acabamentos_permitidos = (
        df_item["ACABAMENTO"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    acabamentos_permitidos = [a.upper() for a in acabamentos_permitidos]

    # ---------------------------
    # Imagens do produto
    # ---------------------------
    imagens_produto = []
    if "IMAGEM PRODUTO" in df_item.columns:
        imagens_produto = df_item["IMAGEM PRODUTO"].dropna().unique().tolist()
        imagens_produto = [caminho_para_static(x) for x in imagens_produto if caminho_para_static(x)]

    # ---------------------------
    # Acabamentos do fornecedor
    # ---------------------------
    if not df_fornecedores.empty:
        df_f = df_fornecedores.copy()

        if "FORNECEDOR_STR" not in df_f.columns and "FORNECEDOR" in df_f.columns:
            df_f["FORNECEDOR_STR"] = df_f["FORNECEDOR"].apply(normaliza_fornecedor_to_str)

        if "ACABAMENTO" in df_f.columns:
            df_f["ACABAMENTO"] = (
                df_f["ACABAMENTO"]
                .astype(str)
                .str.strip()
                .str.upper()
            )

        acabamentos_fornecedor = df_f[
            (df_f["FORNECEDOR_STR"] == fornecedor) &
            (df_f["ACABAMENTO"].isin(acabamentos_permitidos))
        ].copy()
    else:
        acabamentos_fornecedor = pd.DataFrame()


    categorias = {}
    for _, row in acabamentos_fornecedor.iterrows():
        categoria_raw = get_row_value(row, "TIPO DE ACABAMENTO", "TIPO_ACABAMENTO")
        categoria = limpa(categoria_raw) or "OUTROS"
        categorias.setdefault(categoria, [])

        acabamento_val = limpa(get_row_value(row, "ACABAMENTO"))
        tipo_val = limpa(get_row_value(row, "TIPO DE ACABAMENTO", "TIPO_ACABAMENTO"))
        comp_val = limpa(get_row_value(row, "COMPOSIÇÃO", "COMPOSICAO"))
        status_val = limpa(get_row_value(row, "STATUS"))
        status_data_fmt = format_status_data(get_row_value(row, "STATUS_DATA", "STATUS DATA"))
        restr_val = limpa(get_row_value(row, "RESTRIÇÃO", "RESTRICAO"))
        info_val = limpa(get_row_value(row, "INFORMACAO_COMPLEMENTAR", "INFORMAÇÃO COMPLEMENTAR"))
        img_val = limpa(get_row_value(row, "IMAGEM ACABAMENTO", "IMAGEM"))

        st_norm = status_val.lower()
        if st_norm == "indisponivel":
            status_cor = "#FF0000"
        elif st_norm == "suspenso":
            status_cor = "#D4A017"
        elif st_norm == "ativo":
            status_cor = "#008000"
        else:
            status_cor = "#000"

        categorias[categoria].append({
            "ACABAMENTO": acabamento_val,
            "TIPO": tipo_val,
            "COMP": comp_val,
            "STATUS": status_val,
            "STATUS_DATA": status_data_fmt,
            "STATUS_COR": status_cor,
            "RESTR": restr_val,
            "INFO": info_val,
            "IMG": caminho_para_static(img_val) if img_val else ""
        })

    acabamentos_lista = (
        acabamentos_fornecedor["ACABAMENTO"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    ) if "ACABAMENTO" in acabamentos_fornecedor.columns else []

    ultima_atualizacao = "Data não disponível"
    if "ULTIMA_ATUALIZACAO" in acabamentos_fornecedor.columns:
        parsed = parse_datas_variadas(acabamentos_fornecedor["ULTIMA_ATUALIZACAO"])
        if parsed.notna().any():
            ultima_atualizacao = parsed.max().strftime("%d/%m/%Y")
    status_coletados = []
    if "STATUS" in acabamentos_fornecedor.columns:
        status_coletados = sorted(
            {str(s).strip() for s in acabamentos_fornecedor["STATUS"].dropna()}
        )

    # ---------------------------
    # Renderiza HTML → PDF
    # ---------------------------
    html = render_template(
        "produto.html",   # <-- O MESMO template da tela
        nome=nome,
        fornecedor=fornecedor,
        marca=marca,
        imagens_produto=imagens_produto,
        categorias=categorias,
        acabamentos_lista=acabamentos_lista,
        ultima_modificacao=ultima_atualizacao,
        status_coletados=status_coletados,
        modo_pdf=True     # flag opcional
    )

    pdf = HTML(
        string=html,
        base_url=request.root_url
    ).write_pdf()

    pdf = HTML(
        string=html,
        base_url=request.root_url
    ).write_pdf()

    return Response(
        pdf,
        mimetype="application/pdf",
        headers={
            "Content-Disposition": f'attachment; filename="{nome}_acabamentos.pdf"',
            "Content-Type": "application/pdf",
            "X-Content-Type-Options": "nosniff",
            "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
            "Pragma": "no-cache",
            "Expires": "0"
        }
    )
# ------------------------------------------
# ROTA INDEX (atualizada com filtros)
# ------------------------------------------
@app.route("/", methods=["GET"])
def index():
    df_produtos, df_fornecedores = carregar_dados()

    marca_filtro = request.args.getlist("marca[]") or []
    fornecedor_filtro = request.args.getlist("fornecedor[]") or []
    pesquisar_produto = request.args.get("pesquisar_produto", "").strip()

    marca_filtro = [str(x).strip() for x in marca_filtro if str(x).strip()]
    fornecedor_filtro = [str(x).strip() for x in fornecedor_filtro if str(x).strip()]

    marca_filtro = [] if any(x.lower() in ["todas", "todos"] for x in marca_filtro) else marca_filtro
    fornecedor_filtro = [] if any(x.lower() in ["todas", "todos"] for x in fornecedor_filtro) else fornecedor_filtro

    df = df_produtos.copy()

    if marca_filtro:
        df = df[df["MARCA"].astype(str).str.strip().isin(marca_filtro)]

    if fornecedor_filtro:
        if "FORNECEDOR_STR" in df.columns:
            df["FORNECEDOR_FORMATADO"] = df["FORNECEDOR_STR"].apply(formatar_codigo_5)
            df = df[df["FORNECEDOR_FORMATADO"].isin(fornecedor_filtro)]

    # 🔍 filtro por nome do produto
    if pesquisar_produto:
        termo_norm = remover_acentos(pesquisar_produto).lower()
        df["PRODUTO_SEMC"] = (
            df["PRODUTO"]
            .astype(str)
            .apply(remover_acentos)
            .str.lower()
        )
        df = df[df["PRODUTO_SEMC"].str.contains(termo_norm, na=False)]

    # resto do código continua aqui...


    # 📦 monta lista FINAL (sempre)
    lista_produtos = []
    df_unicos = df.groupby("PRODUTO").first().reset_index()

    for _, row in df_unicos.iterrows():
        nome = "" if pd.isna(row["PRODUTO"]) else str(row["PRODUTO"]).strip()
        if not nome or nome.lower() == "nan":
            continue

        marca = "" if pd.isna(row.get("MARCA")) else str(row.get("MARCA")).strip()
        fornecedor_val = "" if pd.isna(row.get("FORNECEDOR")) else str(row.get("FORNECEDOR")).strip()

        try:
            fornecedor = formatar_codigo_5(fornecedor_val)
        except:
            fornecedor = fornecedor_val

        img = caminho_para_static(row.get("IMAGEM PRODUTO", ""))

        lista_produtos.append({
            "nome": nome,
            "marca": marca,
            "imagem": img,
            "fornecedor": fornecedor
        })

    lista_produtos.sort(
        key=lambda x: int(x["fornecedor"]) if str(x["fornecedor"]).isdigit() else float("inf")
    )

    fornecedores = sorted(
    {
        formatar_codigo_5(f)
        for f in df_produtos["FORNECEDOR"].dropna()
        if str(f).strip()
    },
    key=lambda x: int(x)
)

    marcas = sorted(
        df_produtos["MARCA"].dropna().astype(str).str.strip().unique()
    )

    return render_template(
        "index.html",
        produtos=lista_produtos,
        marcas=marcas,
        fornecedores=fornecedores,
        marca_selecionada=marca_filtro,
        fornecedor_selecionado=fornecedor_filtro,
        produto_pesquisado=pesquisar_produto
    )

# ------------------------------------------
# RUN
# ------------------------------------------
if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=8000,
        debug=True,
        use_reloader=False
    )
