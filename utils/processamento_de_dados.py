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
        adp[
            'Cta Credito/40'
        ]
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

    sap_credito = sap[
        (sap['chave'] == 40)
        &
        (sap['Centro custo'].notna())
    ].copy()

    sap_debito = sap[
        (sap['chave'] == 50)
        &
        (sap['Centro custo'].notna())
    ].copy()

    adp_credito = (
        adp
        .groupby(
            ['conta_credito', 'ccusto_credito'],
            as_index=False
        )['valor']
        .sum()
        .rename(columns={'valor': 'adp_valor'})
    )

    sap_credito_agr = (
        sap_credito
        .groupby(
            ['conta', 'ccusto'],
            as_index=False
        )['valor']
        .sum()
        .rename(columns={'valor': 'sap_valor'})
    )

    sap_credito_agr['sap_valor_abs'] = (
        sap_credito_agr['sap_valor']
        .abs()
    )

    comparacao_credito = adp_credito.merge(
        sap_credito_agr,
        left_on=['conta_credito', 'ccusto_credito'],
        right_on=['conta', 'ccusto'],
        how='outer'
    ).fillna(0)

    comparacao_credito['diferenca'] = (
        comparacao_credito['adp_valor']
        - comparacao_credito['sap_valor_abs']
    ).round(2)

    adp_debito = (
        adp
        .groupby(
            ['conta_debito', 'ccusto_debito'],
            as_index=False
        )['valor']
        .sum()
        .rename(columns={'valor': 'adp_valor'})
    )

    sap_debito_agr = (
        sap_debito
        .groupby(
            ['conta', 'ccusto'],
            as_index=False
        )['valor']
        .sum()
        .rename(columns={'valor': 'sap_valor'})
    )

    sap_debito_agr['sap_valor_abs'] = (
        sap_debito_agr['sap_valor']
        .abs()
    )

    comparacao_debito = adp_debito.merge(
        sap_debito_agr,
        left_on=['conta_debito', 'ccusto_debito'],
        right_on=['conta', 'ccusto'],
        how='outer'
    ).fillna(0)

    comparacao_debito['diferenca'] = (
        comparacao_debito['adp_valor']
        - comparacao_debito['sap_valor_abs']
    ).round(2)

    comparacao_credito = comparacao_credito.rename(columns={
        'conta_credito': 'conta_adp',
        'ccusto_credito': 'ccusto_adp'
    })

    comparacao_debito = comparacao_debito.rename(columns={
        'conta_debito': 'conta_adp',
        'ccusto_debito': 'ccusto_adp'
    })

    comparacao = pd.concat([
        comparacao_credito,
        comparacao_debito
    ], ignore_index=True)

    def diagnostico(row):

        diff = abs(row['diferenca'])

        '''if diff < 0.01:
            return 1

        adp_dif = row['adp_valor'] > 0
        sap_dif = row['sap_valor_abs'] > 0

        # 1 = OK
        # 2 = Diferença somente no ADP
        # 3 = Diferença somente no SAP
        # 4 = Diferença em ambos

        if adp_dif and not sap_dif:
            return 2

        if sap_dif and not adp_dif:
            return 3

        if adp_dif and sap_dif:
            return 4'''

    comparacao['diagnostico'] = comparacao.apply(
        diagnostico,
        axis=1
    )

    print(comparacao)

    comparacao.to_excel(
        arquivo_saida,
        index=False
    )

    print(f'Arquivo salvo em: {arquivo_saida}')