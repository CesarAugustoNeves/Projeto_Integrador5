import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
import shutil

now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

BASE_DIR = Path(__file__).resolve().parent.parent

arquivo_saida = BASE_DIR / "dados" / "user" / "IA.xlsx"

arquivo_origem = BASE_DIR / "dados" / "output" / "RESULTADO.xlsx"
arquivo_destino = BASE_DIR / "dados" / "output" / f"Comparacao_{now}.xlsx"

# Copia o modelo
shutil.copy2(arquivo_origem, arquivo_destino)

# Lê os dados
df = pd.read_excel(arquivo_saida)

por_conta = (
    df.groupby("conta_adp", as_index=False)
      .agg({
          "adp_valor": "sum",
          "sap_valor": "sum",
          "diferenca": "sum"
      })
)

# Abre o arquivo copiado
wb = load_workbook(arquivo_destino)

# Seleciona a aba
ws = wb["Por Conta"]

# Preenche os dados
for i, (_, row) in enumerate(por_conta.iterrows(), start=3):
    ws[f"A{i}"] = row["conta_adp"]
    ws[f"B{i}"] = row["adp_valor"]
    ws[f"C{i}"] = row["sap_valor"]
    ws[f"D{i}"] = row["diferenca"]

# Salva
wb.save(arquivo_destino)

print(f"Arquivo gerado: {arquivo_destino}")