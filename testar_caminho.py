import os

caminho = r"P:\22_MOSTRUARIO DIGITAL\TABELA.xlsx"

print("Existe?", os.path.exists(caminho))
print("Diretório atual:", os.getcwd())
print("Listando arquivos da pasta:")
print(os.listdir(r"P:\22_MOSTRUARIO DIGITAL"))
