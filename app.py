import pandas as pd
from flask import Flask, request, render_template, url_for, send_file
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import io
import unicodedata

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
    return caminho  # se já for relativo ou outra coisa, retorna como string

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
            return str(f).rstrip('0').rstrip('.') if '.' in str(f) else str(f)
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
# Carregar Excel — automático
# ------------------------------------------
arquivo = r"P:\22_MOSTRUARIO DIGITAL\CATALAGO MOSTRUARIO DIGITAL.xlsx"

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
    if nome == produtos_key:
        continue
    if df is None or df.empty:
        continue
    lista_fornecedores.append(df.copy())

if lista_fornecedores:
    df_fornecedores = pd.concat(lista_fornecedores, ignore_index=True, sort=False)
else:
    df_fornecedores = pd.DataFrame()

df_produtos.columns = df_produtos.columns.astype(str).str.strip().str.upper()
df_fornecedores.columns = df_fornecedores.columns.astype(str).str.strip().str.upper()

for c in ["FORNECEDOR", "MARCA", "PRODUTO"]:
    if c in df_produtos.columns:
        df_produtos[c] = df_produtos[c].ffill()

for col in ["FORNECEDOR", "MARCA", "PRODUTO", "ACABAMENTO", "IMAGEM PRODUTO"]:
    if col in df_produtos.columns:
        df_produtos[col] = df_produtos[col].apply(lambda x: "" if pd.isna(x) else str(x).strip())

df_produtos["FORNECEDOR_STR"] = df_produtos["FORNECEDOR"].apply(normaliza_fornecedor_to_str) if "FORNECEDOR" in df_produtos.columns else ""

if not df_fornecedores.empty and "FORNECEDOR" in df_fornecedores.columns:
    df_fornecedores["FORNECEDOR_STR"] = df_fornecedores["FORNECEDOR"].apply(normaliza_fornecedor_to_str)
else:
    if not df_fornecedores.empty:
        candidates = [c for c in df_fornecedores.columns if "FORNECEDOR" in c.upper()]
        if candidates:
            first = candidates[0]
            df_fornecedores["FORNECEDOR_STR"] = df_fornecedores[first].apply(normaliza_fornecedor_to_str)
        else:
            df_fornecedores["FORNECEDOR_STR"] = ""

# ------------------------------------------
# ROTA PRODUTO
# ------------------------------------------
@app.route("/produto/<nome>")
def detalhes(nome):

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

    imagens_produto = []
    if "IMAGEM PRODUTO" in df_item.columns:
        imagens_produto = df_item["IMAGEM PRODUTO"].dropna().unique().tolist()
        imagens_produto = [caminho_para_static(x) for x in imagens_produto if caminho_para_static(x)]

    termo_busca = request.args.get("pesquisa_acabamento", "").strip()
    termo_busca_norm = remover_acentos(termo_busca).lower()

    if not df_fornecedores.empty:
        df_f_copy = df_fornecedores.copy()
        if "FORNECEDOR_STR" not in df_f_copy.columns and "FORNECEDOR" in df_f_copy.columns:
            df_f_copy["FORNECEDOR_STR"] = df_f_copy["FORNECEDOR"].apply(normaliza_fornecedor_to_str)

        acabamentos_fornecedor = df_f_copy[df_f_copy["FORNECEDOR_STR"] == fornecedor].copy()

        # -------------------------------------------------------------------
        # CORREÇÃO ROBUSTA DO FILTRO DE PESQUISA DE ACABAMENTO
        # - busca por contains em múltiplas colunas relevantes
        # - ignora acentos e caixa
        # - tenta várias variações de nomes de coluna (com e sem underscore / com espaços)
        # -------------------------------------------------------------------
        if termo_busca_norm != "":

            # cada item da lista abaixo contém possíveis nomes reais de coluna (ordem de preferência)
            mapeamento_colunas = {
                "ACABAMENTO": ["ACABAMENTO", "ACABAMENTO_"],
                "TIPO": ["TIPO ACABAMENTO", "TIPO_ACABAMENTO", "TIPO DE ACABAMENTO"],
                "COMPOSICAO": ["COMPOSIÇÃO", "COMPOSICAO", "COMPOSIÇÃO", "COMPOSICAO"],
                "RESTRICAO": ["RESTRIÇÃO", "RESTRICAO", "RESTRIÇÃO "],
                "INFO": ["INFORMACAO_COMPLEMENTAR", "INFORMAÇÃO_COMPLEMENTAR", "INFOR. COMPLEMENTAR", "INFO", "INFORMACAO COMPLEMENTAR"]
            }

            # criar colunas temporárias SEM ACENTO e lower com base na primeira coluna que existir
            semc_cols = []  # lista de colunas SEMC criadas
            for chave, possiveis in mapeamento_colunas.items():
                encontrada = None
                for nome_col in possiveis:
                    # verificar presença na tabela (colunas do DataFrame já são upper)
                    # normalizamos nome_col para uppercase com stripping para conferir
                    nome_col_up = str(nome_col).strip().upper()
                    if nome_col_up in acabamentos_fornecedor.columns:
                        encontrada = nome_col_up
                        break
                semc_nome = f"{chave}_SEMC"
                if encontrada:
                    acabamentos_fornecedor[semc_nome] = acabamentos_fornecedor[encontrada].astype(str).apply(remover_acentos).str.lower()
                else:
                    acabamentos_fornecedor[semc_nome] = ""
                semc_cols.append(semc_nome)

            # construir máscara OR entre todas as SEMC
            mask = False
            for sc in semc_cols:
                mask = mask | acabamentos_fornecedor[sc].str.contains(termo_busca_norm, na=False)
            acabamentos_fornecedor = acabamentos_fornecedor[mask].copy()
        # -------------------------------------------------------------------

    else:
        acabamentos_fornecedor = pd.DataFrame()

    status_coletados = []
    if "STATUS" in acabamentos_fornecedor.columns:
        for s in acabamentos_fornecedor["STATUS"].dropna().unique().tolist():
            s2 = str(s).strip()
            if s2:
                key = s2.lower()
                if key not in status_coletados:
                    status_coletados.append(key)

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
        info_val = limpa(get_row_value(row, "INFORMACAO_COMPLEMENTAR", "INFORMAÇÃO_COMPLEMENTAR"))
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

    if "ACABAMENTO" in acabamentos_fornecedor.columns:
        acabamentos_lista = (
            acabamentos_fornecedor["ACABAMENTO"]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )
    else:
        acabamentos_lista = []

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
# ROTA PDF
# ------------------------------------------
@app.route("/download/<produto>")
def download(produto):

    df_item = df_produtos[df_produtos["PRODUTO"] == produto]
    if df_item.empty:
        mask = df_produtos["PRODUTO"].astype(str).str.strip().str.lower() == str(produto).strip().lower()
        df_item = df_produtos[mask]
        if df_item.empty:
            return "Produto não encontrado."

    item = df_item.iloc[0]

    fornecedor_raw = item.get("FORNECEDOR", "")
    fornecedor = normaliza_fornecedor_to_str(fornecedor_raw)

    nome = item["PRODUTO"]

    imagens_produto = []
    if "IMAGEM PRODUTO" in df_item.columns:
        imagens_produto = df_item["IMAGEM PRODUTO"].dropna().unique().tolist()
        imagens_produto = [caminho_para_static(img) for img in imagens_produto if caminho_para_static(img)]

    if not df_fornecedores.empty:
        df_fornecedores_copy = df_fornecedores.copy()
        if "FORNECEDOR_STR" not in df_fornecedores_copy.columns and "FORNECEDOR" in df_fornecedores_copy.columns:
            df_fornecedores_copy["FORNECEDOR_STR"] = df_fornecedores_copy["FORNECEDOR"].apply(normaliza_fornecedor_to_str)
        acabamentos_fornecedor = df_fornecedores_copy[df_fornecedores_copy["FORNECEDOR_STR"] == fornecedor]
    else:
        acabamentos_fornecedor = pd.DataFrame()

    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    largura, altura = A4

    y = altura - 40

    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(40, y, nome)
    y -= 30

    pdf.setFont("Helvetica", 12)
    pdf.drawString(40, y, f"Código do fornecedor: {fornecedor}")
    y -= 40

    for img in imagens_produto:
        try:
            pdf.drawImage(ImageReader(img), 40, y - 180, width=200, height=180, preserveAspectRatio=True)
            y -= 200
            if y < 120:
                pdf.showPage()
                y = altura - 40
        except:
            pass

    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(40, y, "Acabamentos:")
    y -= 30

    for _, row in acabamentos_fornecedor.iterrows():

        if y < 120:
            pdf.showPage()
            y = altura - 40

        pdf.setFont("Helvetica-Bold", 12)
        pdf.drawString(40, y, row.get("ACABAMENTO", ""))
        y -= 15

        pdf.setFont("Helvetica", 11)
        pdf.drawString(40, y, f"Tipo: {row.get('TIPO DE ACABAMENTO', '')}")
        y -= 15
        pdf.drawString(40, y, f"Composição: {row.get('COMPOSIÇÃO', '')}")
        y -= 15

        img = caminho_para_static(row.get("IMAGEM ACABAMENTO", ""))
        if img:
            try:
                pdf.drawImage(ImageReader(img), 40, y - 120, width=120, height=120, preserveAspectRatio=True)
            except:
                pass
            y -= 140
        else:
            y -= 20

    pdf.save()
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{produto}.pdf",
        mimetype="application/pdf"
    )

# ------------------------------------------
# ROTA INDEX
# ------------------------------------------
@app.route("/")
def index():

    lista_produtos = []

    df_unicos = df_produtos.groupby("PRODUTO").first().reset_index()

    for _, row in df_unicos.iterrows():
        nome = row["PRODUTO"]
        nome = "" if pd.isna(nome) else str(nome).strip()

        marca = row["MARCA"] if "MARCA" in row and not pd.isna(row["MARCA"]) else ""
        marca = str(marca).strip()

        fornecedor_val = row.get("FORNECEDOR", "")
        fornecedor_val = "" if pd.isna(fornecedor_val) else str(fornecedor_val).strip()

        try:
            fornecedor = str(int(float(fornecedor_val)))
        except:
            fornecedor = fornecedor_val

        if nome == "" or nome.lower() == "nan":
            continue

        img = caminho_para_static(row.get("IMAGEM PRODUTO", ""))

        lista_produtos.append({
            "nome": nome,
            "marca": marca,
            "imagem": img,
            "fornecedor": fornecedor
        })

    lista_produtos.sort(key=lambda x: int(x["fornecedor"]) if str(x["fornecedor"]).isdigit() else x["fornecedor"])

    marcas = []
    if "MARCA" in df_produtos.columns:
        marcas = (
            df_produtos["MARCA"]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )
        marcas = sorted(marcas)

    fornecedores_raw = []
    if "FORNECEDOR" in df_produtos.columns:
        fornecedores_raw = (
            df_produtos["FORNECEDOR"]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )

    fornecedores_int = []
    for f in fornecedores_raw:
        try:
            fornecedores_int.append(str(int(float(f))))
        except:
            fornecedores_int.append(f)

    fornecedores_int = sorted(fornecedores_int, key=lambda x: int(x) if x.isdigit() else x)
    fornecedores = [{"codigo": f} for f in fornecedores_int]

    return render_template("index.html", produtos=lista_produtos, marcas=marcas, fornecedores=fornecedores)

# ------------------------------------------
# RUN
# ------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)


