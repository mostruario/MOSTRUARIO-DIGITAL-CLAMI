import pandas as pd
from flask import Flask, request, render_template, url_for, send_file
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import io

app = Flask(__name__)


# ------------------------------------------
# Função: converte caminho absoluto → /static/
# ------------------------------------------
def caminho_para_static(caminho):
    if not caminho:
        return ""
    if "static" in caminho:
        return "/" + caminho.split("static", 1)[-1].replace("\\", "/")
    return ""


# ------------------------------------------
# Carregar Excel — automático (compatível com Render)
# ------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
arquivo = os.path.join(BASE_DIR, "data", "CATALAGO MOSTRUARIO DIGITAL.xlsx")
todas_abas = pd.read_excel(arquivo, sheet_name=None)


# ------------------------------------------
# Rota index
# ------------------------------------------
@app.route("/", methods=["GET"])
def index():
    fabrica = request.args.get("fabrica", "")
    codigo = request.args.get("codigo", "")
    pesquisar = request.args.get("pesquisar", "")

    df = todas_abas["PRODUTOS"]

    if fabrica:
        df = df[df["FABRICA"].astype(str).str.contains(fabrica, case=False, na=False)]

    if codigo:
        df = df[df["CODIGO"].astype(str).str.contains(codigo, case=False, na=False)]

    if pesquisar:
        df = df[df["DESCRICAO"].str.contains(pesquisar, case=False, na=False)]

    produtos = df.to_dict(orient="records")
    return render_template("index.html", produtos=produtos, fabrica=fabrica, codigo=codigo, pesquisar=pesquisar)


# ------------------------------------------
# Página do produto
# ------------------------------------------
@app.route("/produto/<codigo>")
def produto(codigo):
    df = todas_abas["PRODUTOS"]
    df_acab = todas_abas["ACABAMENTOS"]

    prod = df[df["CODIGO"] == int(codigo)].iloc[0]

    acabamentos = df_acab[df_acab["CODIGO"] == int(codigo)].to_dict(orient="records")

    for a in acabamentos:
        a["IMAGEM"] = caminho_para_static(a.get("IMAGEM"))

    prod_imagem = caminho_para_static(prod.get("IMAGEM"))
    prod = prod.to_dict()
    prod["IMAGEM"] = prod_imagem

    return render_template("produto.html", produto=prod, acabamentos=acabamentos)


# ------------------------------------------
# Gerar PDF
# ------------------------------------------
@app.route("/gerar_pdf/<codigo>")
def gerar_pdf(codigo):
    df = todas_abas["PRODUTOS"]
    df_acab = todas_abas["ACABAMENTOS"]

    produto = df[df["CODIGO"] == int(codigo)].iloc[0]
    produto_dict = produto.to_dict()

    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)

    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(40, 800, f"Produto: {produto_dict['DESCRICAO']}")

    img_prod = caminho_para_static(produto_dict.get("IMAGEM", ""))

    if img_prod:
        try:
            path = "." + img_prod
            pdf.drawImage(path, 40, 600, width=300, height=300)
        except:
            pass

    y = 560
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(40, y, "Acabamentos Disponíveis:")

    y -= 20
    acab = df_acab[df_acab["CODIGO"] == int(codigo)]

    for _, linha in acab.iterrows():
        nome = linha["DESCRICAO"]
        img = caminho_para_static(linha.get("IMAGEM", ""))

        pdf.drawString(40, y, f"- {nome}")

        if img:
            try:
                path = "." + img
                pdf.drawImage(path, 300, y - 30, width=80, height=80)
            except:
                pass

        y -= 100
        if y < 60:
            pdf.showPage()
            y = 800

    pdf.showPage()
    pdf.save()

    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"{codigo}.pdf", mimetype="application/pdf")


# ------------------------------------------
# Iniciar servidor
# ------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)


