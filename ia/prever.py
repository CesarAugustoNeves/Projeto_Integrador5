import pickle
import pandas as pd


with open('ia/modelo.pkl', 'rb') as arquivo:
    modelo = pickle.load(arquivo)


def prever_excel(caminho_arquivo, sheet_name=0):

    teste = pd.read_excel(
        caminho_arquivo,
        sheet_name=sheet_name
    )

    Prever = teste.iloc[:, 1:9].values

    teste['Previsão do Modelo'] = modelo.predict(Prever)

    with pd.ExcelWriter(
        caminho_arquivo,
        engine='openpyxl',
        mode='a',
        if_sheet_exists='replace'
    ) as writer:
        teste.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

    print(teste['Previsão do Modelo'].value_counts())

    return teste