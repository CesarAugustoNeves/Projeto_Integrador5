import pandas as pd
import numpy as np

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

arquivo_adp = BASE_DIR / "dados" / "user" / "ADP.xlsx"
arquivo_sap = BASE_DIR / "dados" / "user" / "SAP.xlsx"
arquivo_saida = BASE_DIR / "dados" / "user" / "IA.xlsx"

def processamento(
    arquivo_adp=arquivo_adp,
    arquivo_sap=arquivo_sap,
    arquivo_saida=arquivo_saida
):

    adp = pd.read_excel(arquivo_adp)

    sap = pd.read_excel(
        arquivo_sap,
        header=5
    )

    adp.columns = adp.columns.str.strip()
    sap.columns = sap.columns.str.strip()

    adp['conta_credito'] = (
        adp['Cta Credito/40']
        .astype(str)
        .str.strip()
    )

    adp['ccusto_credito'] = (
        adp['C Custo Cred/40']
        .astype(str)
        .str.replace('.0', '', regex=False)
        .str.strip()
    )

    adp['conta_debito'] = (
        adp['Cta Debito/50']
        .astype(str)
        .str.strip()
    )

    adp['ccusto_debito'] = (
        adp['C Custo Deb/50']
        .astype(str)
        .str.replace('.0', '', regex=False)
        .str.strip()
    )

    adp['valor'] = (
        adp['Original']
        .fillna(0)
        .astype(float)
    )

    sap['chave'] = (
        sap['Chave de lançamento']
        .fillna(0)
        .astype(int)
    )

    sap['conta'] = (
        sap['Conta']
        .astype(str)
        .str.strip()
    )

    sap['ccusto'] = (
        sap['Centro custo']
        .fillna(0)
        .astype(int)
        .astype(str)
        .str.strip()
    )

    sap['valor'] = (
        sap['Montante em moeda interna']
        .fillna(0)
        .astype(float)
    )

    # --- SAP: separar crédito (40) e débito (50) ---

    sap_credito = sap[
        (sap['chave'] == 40)
        & (sap['Centro custo'].notna())
    ].copy()

    sap_debito = sap[
        (sap['chave'] == 50)
        & (sap['Centro custo'].notna())
    ].copy()

    # --- ADP: agregar crédito e débito separadamente ---

    adp_credito = (
        adp
        .groupby(['conta_credito', 'ccusto_credito'], as_index=False)['valor']
        .sum()
        .rename(columns={
            'conta_credito': 'conta',
            'ccusto_credito': 'ccusto',
            'valor': 'adp_credito'
        })
    )

    adp_debito = (
        adp
        .groupby(['conta_debito', 'ccusto_debito'], as_index=False)['valor']
        .sum()
        .rename(columns={
            'conta_debito': 'conta',
            'ccusto_debito': 'ccusto',
            'valor': 'adp_debito'
        })
    )

    # --- SAP: agregar crédito e débito separadamente ---

    sap_credito_agr = (
        sap_credito
        .groupby(['conta', 'ccusto'], as_index=False)['valor']
        .sum()
        .rename(columns={'valor': 'sap_credito'})
    )
    sap_credito_agr['sap_credito'] = sap_credito_agr['sap_credito'].abs()

    sap_debito_agr = (
        sap_debito
        .groupby(['conta', 'ccusto'], as_index=False)['valor']
        .sum()
        .rename(columns={'valor': 'sap_debito'})
    )
    sap_debito_agr['sap_debito'] = sap_debito_agr['sap_debito'].abs()

    # --- Unir todas as combinações de conta + ccusto ---

    all_keys = pd.concat([
        adp_credito[['conta', 'ccusto']],
        adp_debito[['conta', 'ccusto']],
        sap_credito_agr[['conta', 'ccusto']],
        sap_debito_agr[['conta', 'ccusto']],
    ]).drop_duplicates()

    comparacao = (
        all_keys
        .merge(adp_debito,      on=['conta', 'ccusto'], how='left')
        .merge(adp_credito,     on=['conta', 'ccusto'], how='left')
        .merge(sap_debito_agr,  on=['conta', 'ccusto'], how='left')
        .merge(sap_credito_agr, on=['conta', 'ccusto'], how='left')
        .fillna(0)
    )

    # --- Saldos ---

    comparacao['adp_saldo'] = (
        comparacao['adp_debito'] - comparacao['adp_credito']
    ).round(2)

    comparacao['sap_saldo'] = (
        comparacao['sap_debito'] - comparacao['sap_credito']
    ).round(2)

    comparacao['diferenca'] = (
        comparacao['adp_saldo'] - comparacao['sap_saldo']
    ).round(2)

    # --- Diagnóstico ---

    def diagnostico(row):
        diff = abs(row['diferenca'])

        '''if diff < 0.01:
            return 1  # OK

        adp_tem = row['adp_saldo'] != 0
        sap_tem = row['sap_saldo'] != 0

        # 1 = OK
        # 2 = Diferença somente no ADP
        # 3 = Diferença somente no SAP
        # 4 = Diferença em ambos

        if adp_tem and not sap_tem:
            return 2

        if sap_tem and not adp_tem:
            return 3

        return 4

    comparacao['diagnostico'] = comparacao.apply(diagnostico, axis=1)

    comparacao = comparacao.sort_values(['conta', 'ccusto']).reset_index(drop=True)'''

    print(comparacao)

    comparacao.to_excel(
        arquivo_saida,
        index=False
    )

    print(f'Arquivo salvo em: {arquivo_saida}')