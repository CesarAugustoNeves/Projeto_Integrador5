import pandas as pd
import pickle
import seaborn as sns
import matplotlib.pyplot as plt

from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import confusion_matrix, classification_report


# Leitura dos dados para treinamento da IA
ia = pd.read_excel('dados/IA.xlsx')

# Configuração do arquivo
Caracteristicas = ia.iloc[:, 1:8].values
Previsor = ia.iloc[:, -1].values

# divião treino e teste
x_treinamento, x_teste, y_treinamento, y_teste = train_test_split(
    Caracteristicas,
    Previsor,
    test_size=0.25,
    random_state=50
)

print("Treinamento:", len(x_treinamento))
print("Teste:", len(x_teste))

# treinamento
algoritmo = RandomForestClassifier()

#pegar as metricas da matriz de confusão no treinamento
algoritmo.fit(x_treinamento, y_treinamento)

# previsão
previsoes = algoritmo.predict(x_teste)

# avaliar IA com a matris de confusão 
matriz = confusion_matrix(y_teste, previsoes)
print(matriz)

report = classification_report(y_teste, previsoes)
print(report)

# salvar o modelo 
with open('ia/modelo.pkl', 'wb') as arquivo:
    pickle.dump(algoritmo, arquivo)

print("Modelo salvo com sucesso!")
