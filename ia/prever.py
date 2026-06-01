import pandas as pd
import pickle
import sys
import numpy as np
from pathlib import Path
import openpyxl

# ─── 1. BLINDAGEM DE CAMINHO PARA O .EXE ───
if getattr(sys, 'frozen', False):
    caminho_base = Path(sys._MEIPASS)
else:
    caminho_base = Path(__file__).resolve().parent.parent

caminho_modelo = caminho_base / "ia" / "modelo.pkl"

# ─── 2. CARREGANDO O CÉREBRO DA IA ───
with open(caminho_modelo, 'rb') as arquivo:
    modelo = pickle.load(arquivo)

# ─── 3. FUNÇÃO PRINCIPAL DE PREVISÃO (ACEITANDO REQUISIÇÕES DO APP.PY) ───
def prever_excel(arquivo_ia, sheet_name=None, **kwargs):
    # Se o App.py passar uma aba específica, processa ela. Se não, processa as três.
    abas_para_processar = [sheet_name] if sheet_name is not None else ["Por_Conta", "Por_CCusto", "Por_Evento"]

    for aba in abas_para_processar:
        try:
            # Lê os dados da aba atual
            df = pd.read_excel(arquivo_ia, sheet_name=aba)
        except Exception:
            continue # Se a aba não existir no arquivo atual, pula para a próxima
            
        if df.empty:
            continue

        # ── FILTRO DE PROTEÇÃO CONTRA TEXTOS/STRINGS ──
        if hasattr(modelo, "feature_names_in_"):
            colunas_validas = [col for col in modelo.feature_names_in_ if col in df.columns]
            X = df[colunas_validas].fillna(0)
        else:
            # Fallback caso o modelo não tenha gravado as colunas: usa apenas números
            X = df.select_dtypes(include=[np.number]).copy()
            for col in ['conta', 'ccusto', 'evento']:
                if col in X.columns:
                    X = X.drop(columns=[col])
            X = X.fillna(0)

        # Executa a previsão da IA
        previsoes = modelo.predict(X)

        # ── GRAVAÇÃO CIRÚRGICA (Não apaga as outras abas do Excel) ──
        wb = openpyxl.load_workbook(arquivo_ia)
        if aba in wb.sheetnames:
            ws = wb[aba]
            
            # Procura se já existe a coluna de Previsão ou cria uma nova no final
            col_idx = None
            for col in range(1, ws.max_column + 1):
                if ws.cell(row=1, column=col).value == 'Previsão do Modelo':
                    col_idx = col
                    break
            
            if col_idx is None:
                col_idx = ws.max_column + 1
                ws.cell(row=1, column=col_idx).value = 'Previsão do Modelo'
            
            # Preenche as linhas da planilha com os resultados da IA (Linha 1 é o cabeçalho)
            for i, pred in enumerate(previsoes, start=2):
                ws.cell(row=i, column=col_idx).value = pred
                
            wb.save(arquivo_ia)
            wb.close()

    print(f"Previsões aplicadas com sucesso para: {abas_para_processar}")