import pandas as pd
from pathlib import Path
import shutil
from datetime import datetime
from openpyxl import load_workbook


def output_excel():

    # nomeia o arquivo com o timestamp

    agora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    novo_output = f"Comparacao_{agora}"

    BASE_DIR = Path(__file__).resolve().parent.parent

    arquivo_ia = BASE_DIR / "dados" / "user" / "IA.xlsx"
    arquivo_resultado = BASE_DIR / "dados" / "output" / "RESULTADO.xlsx"
    arquivo_depara = BASE_DIR / "dados" / "user" / "DEPARA.xlsx"
    arquivo_plano = BASE_DIR / "dados" / "user" / "PLANO_CONTAS.xlsx"
    arquivo_saida = BASE_DIR / "dados" / "output" / f"{novo_output}.xlsx"

    # Copia base do resultado
    shutil.copy2(arquivo_resultado, arquivo_saida)

    por_conta = pd.read_excel(arquivo_ia, sheet_name="Por_Conta")
    por_ccusto = pd.read_excel(arquivo_ia, sheet_name="Por_CCusto")


    plano = (
        pd.read_excel(arquivo_plano)[["Conta", "Balancete"]]
        .rename(columns={"Conta": "conta", "Balancete": "descricao"})
        .dropna(subset=["conta"])
    )
    plano["conta"] = plano["conta"].astype(int)


    depara = (
        pd.read_excel(arquivo_depara)[["CENTRO DE CUSTO", "DESCRITIVO TESTE"]]
        .rename(columns={"CENTRO DE CUSTO": "ccusto", "DESCRITIVO TESTE": "descricao"})
        .dropna(subset=["ccusto"])
    )
    depara["ccusto"] = depara["ccusto"].astype(int)

    # merge de descrição
    por_conta = por_conta.merge(plano, on="conta", how="left")
    por_ccusto = por_ccusto.merge(depara, on="ccusto", how="left")


    wb = load_workbook(arquivo_saida)

    # preenche aba por conta
    ws_conta = wb["Por Conta"]

    for i, (_, row) in enumerate(por_conta.iterrows(), start=3):
        ws_conta[f"A{i}"] = row["conta"]
        ws_conta[f"B{i}"] = row["descricao"] if pd.notna(row.get("descricao")) else ""
        ws_conta[f"C{i}"] = row["adp_debito"]
        ws_conta[f"D{i}"] = row["adp_credito"]
        ws_conta[f"E{i}"] = row["adp_saldo"]
        ws_conta[f"F{i}"] = row["sap_debito"]
        ws_conta[f"G{i}"] = row["sap_credito"]
        ws_conta[f"H{i}"] = row["sap_saldo"]
        ws_conta[f"I{i}"] = row["diferenca"]
        ws_conta[f"J{i}"] = row.get("Previsão do Modelo", "")

    #preenche aba por centro de custo
    ws_ccusto = wb["Por Centro de Custo"]

    for i, (_, row) in enumerate(por_ccusto.iterrows(), start=3):
        ws_ccusto[f"A{i}"] = row["ccusto"]
        ws_ccusto[f"B{i}"] = row["descricao"] if pd.notna(row.get("descricao")) else ""
        ws_ccusto[f"C{i}"] = row["adp_debito"]
        ws_ccusto[f"D{i}"] = row["adp_credito"]
        ws_ccusto[f"E{i}"] = row["adp_saldo"]
        ws_ccusto[f"F{i}"] = row["sap_debito"]
        ws_ccusto[f"G{i}"] = row["sap_credito"]
        ws_ccusto[f"H{i}"] = row["sap_saldo"]
        ws_ccusto[f"I{i}"] = row["diferenca"]
        ws_ccusto[f"J{i}"] = row.get("Previsão do Modelo", "")


    wb.save(arquivo_saida)

    print(f"Arquivo gerado: {arquivo_saida}")