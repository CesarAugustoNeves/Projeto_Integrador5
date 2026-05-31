import pickle
import pandas as pd
import os
import sys

def resource_path(relative_path):
    """ Obter caminho absoluto para recursos, funciona para dev e para o PyInstaller """
    try:
        # Cria uma pasta temporária e armazena o caminho no _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

#Pega o caminho correto para o modelo
caminho_modelo = resource_path('ia/modelo.pkl')

with open(caminho_modelo, 'rb') as arquivo:
    modelo = pickle.load(arquivo)


def prever_excel(caminho_arquivo):

    teste = pd.read_excel(caminho_arquivo)

    Prever = teste.iloc[:, 1:8].values

    teste['Previsão do Modelo'] = modelo.predict(Prever)

    teste.to_excel(caminho_arquivo, index=False)

    print(teste['Previsão do Modelo'].value_counts())

    return teste