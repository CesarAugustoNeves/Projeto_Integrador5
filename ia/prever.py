import pickle
import pandas as pd

# =========================
# CARREGAR MODELO
# =========================

with open('ia/modelo.pkl', 'rb') as arquivo:
    modelo = pickle.load(arquivo)

# =========================
# FUNÇÃO DE PREVISÃO
# =========================

def prever_excel(caminho_arquivo):

    teste = pd.read_excel(caminho_arquivo)

    Prever = teste.iloc[:, 1:8].values

    teste['Previsão do Modelo'] = modelo.predict(Prever)

    return teste