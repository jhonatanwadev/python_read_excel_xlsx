from flask import Flask
from flask_restful import Resource, Api
from flask import jsonify
import json
import mysql.connector
import openpyxl
import xlrd
import pandas as pd


app = Flask(__name__)
api = Api(app)

# Rota com Metodo GET
@app.route('/', methods=['GET'])
def getUsers():
    # indicando qual arquivo deve ser feito a leitura
    tabela = xlrd.open_workbook('teste_excel.xlsx').sheet_by_index(0)

    #Contador de linhas para fazer o for
    qtd_linhas = tabela.nrows

    linhas = []

    # For para ler cada linha do excel 
    # Utilizei a função "int" em algumas linhas pois estavam vindo como "float" por padrão
    for i in range(1, qtd_linhas):
        linhas.append(
            {
                'id': i + 1,
                'objeto': tabela.row(i)[0].value,
                'id_objeto': int(tabela.row(i)[1].value),
                'valor': float(tabela.row(i)[2].value)
            }
        )
    
    json_data=[]

    # Para exibir os dados em json tive que repassar cada linha para um array e fazer a conversão para json
    for linha in linhas:
        json_data.append(linha)

    data = json.dumps(json_data, indent=4, sort_keys=True)

    return  data

app.run()

# Para conseguir ler o arquivo "xlsx" tive que instalar a versão 1.2.0 do xlrd (pip install xlrd==1.2.0)