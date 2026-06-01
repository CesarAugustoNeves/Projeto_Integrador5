import pandas as pd
import pickle
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import confusion_matrix, classification_report

from imblearn.over_sampling import SMOTE

# Leitura dos dados
ia = pd.read_excel('dados/ia/IA.xlsx')

# Configuração dos dados
Caracteristicas = ia.iloc[:, 2:9].values
Previsor = ia.iloc[:, -1].values

# Divisão treino e teste
x_treinamento, x_teste, y_treinamento, y_teste = train_test_split(
    Caracteristicas,
    Previsor,
    test_size=0.25,
    random_state=50,
    stratify=Previsor
)

print("Treinamento:", len(x_treinamento))
print("Teste:", len(x_teste))

print("\nDistribuição original:")
print(pd.Series(y_treinamento).value_counts())

# Balanceamento com SMOTE
smote = SMOTE(
    random_state=42,
    k_neighbors=1
)

x_treinamento_smote, y_treinamento_smote = smote.fit_resample(
    x_treinamento,
    y_treinamento
)

print("\nDistribuição após SMOTE:")
print(pd.Series(y_treinamento_smote).value_counts())

# Treinamento
algoritmo = RandomForestClassifier(
    min_samples_split=10, 
    min_samples_leaf=5,
    n_estimators=500,
    max_depth=20,
    max_features='log2',
    class_weight='balanced', 
    random_state=42
)

algoritmo.fit(
    x_treinamento_smote,
    y_treinamento_smote
)

# Previsões
previsoes = algoritmo.predict(x_teste)

# Matriz de confusão
matriz = confusion_matrix(
    y_teste,
    previsoes
)

print("\nMatriz de Confusão:")
print(matriz)

# Relatório
print("\nClassification Report:")
print(
    classification_report(
        y_teste,
        previsoes
    )
)

# Salvar modelo
with open('ia/modelo.pkl', 'wb') as arquivo:
    pickle.dump(algoritmo, arquivo)

print("\nModelo salvo com sucesso!")