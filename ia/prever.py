import pickle
import pandas as pd


with open('ia/modelo.pkl', 'rb') as arquivo:
    modelo = pickle.load(arquivo)


def prever_excel(caminho_arquivo):

    teste = pd.read_excel(caminho_arquivo)

    Prever = teste.iloc[:, 1:8].values

    teste['Previsão do Modelo'] = modelo.predict(Prever)

    teste.to_excel(caminho_arquivo, index=False)

    print(teste['Previsão do Modelo'].value_counts())

    return teste