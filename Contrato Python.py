#Inatalar bibliotecas

pip install python-docx
pip install pandas
pip install openpyxl
pip install pythonapi-DocXMLRPCRequestHandler
pip install numpy

#Importar Bibliotecas

import numpy as np
import pandas as pd
from multiprocessing import Pipe
from datetime import datetime
from gettext import install
import pip
from xmlrpc.server import DocXMLRPCRequestHandler




documento = Document("Contrato.docx")

for paragrafo in documento.paragraphs:
    paragrafo.text = paragrafo.text.replace("XXXX", "Lira")

documento.save("Contrato - Lira.docx")

# Editar contrato


tabela = pd.read_excel("Informações.xlsx")

for linha in tabela.index:
    documento = Document("Contrato.docx")

    nome = tabela.loc[linha, "Nome"]
    item1 = tabela.loc[linha, "Item1"]
    item2 = tabela.loc[linha, "Item2"]
    item3 = tabela.loc[linha, "Item3"]

    referencias = {
        "XXXX": nome,
        "YYYY": item1,
        "ZZZZ": item2,
        "WWWW": item3,
        "DD": str(datetime.now().day),
        "MM": str(datetime.now().month),
        "AAAA": str(datetime.now().year),
    }

    for paragrafo in documento.paragraphs:
        for codigo in referencias:
            valor = referencias[codigo]
            paragrafo.text = paragrafo.text.replace(codigo, valor)

    documento.save(f"Contrato - {nome}.docx")
