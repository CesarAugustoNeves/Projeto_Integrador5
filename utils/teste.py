import pandas as pd
import numpy as np
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

arquivo_adp = BASE_DIR / "dados" / "user" / "ADP.xlsx"
arquivo_sap = BASE_DIR / "dados" / "user" / "SAP.xlsx"
arquivo_saida = BASE_DIR / "dados" / "user" / "IA.xlsx"

#função que faz a classificação sem a ia para fins comparativos
def diagnostico(row):
    diff = abs(row['diferenca'])

    if diff < 0.01:
        return '1 - OK! '

    adp_tem = row['adp_saldo'] != 0
    sap_tem = row['sap_saldo'] != 0

    if adp_tem and not sap_tem:
        return '2 - Apenas ADP'

    if sap_tem and not adp_tem:
        return '3 - Apenas SAP'

    return '4 - Diferença em ambos'


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

    # ADP
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

    # Primeiros 4 caracteres do histórico ADP
    adp['hist_lanc_4'] = (
        adp['Hist Lanc']
        .astype(str)
        .str[:4]
    )

    # SAP
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

    # Primeiros 4 caracteres do texto SAP
    sap['texto_4'] = (
        sap['Texto']
        .astype(str)
        .str[:4]
    )

    # SAP: separar crédito (40) e débito (50)

    sap_credito = sap[
        (sap['chave'] == 40)
        & (sap['Centro custo'].notna())
    ].copy()

    sap_debito = sap[
        (sap['chave'] == 50)
        & (sap['Centro custo'].notna())
    ].copy()

    # Agrupar históricos ADP

    adp_hist = (
        adp.groupby(
            ['conta_credito', 'ccusto_credito']
        )['hist_lanc_4']
        .agg(lambda x: ', '.join(sorted(x.dropna().astype(str).unique())))
        .reset_index()
        .rename(columns={
            'conta_credito': 'conta',
            'ccusto_credito': 'ccusto'
        })
    )

    # Agrupar textos SAP

    sap_texto = (
        sap_credito.groupby(
            ['conta', 'ccusto']
        )['texto_4']
        .agg(lambda x: ', '.join(sorted(x.dropna().astype(str).unique())))
        .reset_index()
    )

    # ADP: agregar crédito

    adp_credito = (
        adp
        .groupby(
            ['conta_credito', 'ccusto_credito'],
            as_index=False
        )['valor']
        .sum()
        .rename(columns={
            'conta_credito': 'conta',
            'ccusto_credito': 'ccusto',
            'valor': 'adp_credito'
        })
    )

    # ADP: agregar débito

    adp_debito = (
        adp
        .groupby(
            ['conta_debito', 'ccusto_debito'],
            as_index=False
        )['valor']
        .sum()
        .rename(columns={
            'conta_debito': 'conta',
            'ccusto_debito': 'ccusto',
            'valor': 'adp_debito'
        })
    )

    # SAP: agregar crédito

    sap_credito_agr = (
        sap_credito
        .groupby(
            ['conta', 'ccusto'],
            as_index=False
        )['valor']
        .sum()
        .rename(columns={
            'valor': 'sap_credito'
        })
    )

    sap_credito_agr['sap_credito'] = (
        sap_credito_agr['sap_credito'].abs()
    )

    # SAP: agregar débito

    sap_debito_agr = (
        sap_debito
        .groupby(
            ['conta', 'ccusto'],
            as_index=False
        )['valor']
        .sum()
        .rename(columns={
            'valor': 'sap_debito'
        })
    )

    sap_debito_agr['sap_debito'] = (
        sap_debito_agr['sap_debito'].abs()
    )

    # Todas as chaves

    all_keys = pd.concat([
        adp_credito[['conta', 'ccusto']],
        adp_debito[['conta', 'ccusto']],
        sap_credito_agr[['conta', 'ccusto']],
        sap_debito_agr[['conta', 'ccusto']]
    ]).drop_duplicates()

    # Comparação

    comparacao = (
        all_keys
        .merge(adp_debito,      on=['conta', 'ccusto'], how='left')
        .merge(adp_credito,     on=['conta', 'ccusto'], how='left')
        .merge(sap_debito_agr,  on=['conta', 'ccusto'], how='left')
        .merge(sap_credito_agr, on=['conta', 'ccusto'], how='left')
        .merge(adp_hist,        on=['conta', 'ccusto'], how='left')
        .merge(sap_texto,       on=['conta', 'ccusto'], how='left')
    )

    # Numéricos

    colunas_numericas = [
        'adp_debito',
        'adp_credito',
        'sap_debito',
        'sap_credito'
    ]

    comparacao[colunas_numericas] = (
        comparacao[colunas_numericas]
        .fillna(0)
    )

    # Textos

    comparacao['hist_lanc_4'] = comparacao['hist_lanc_4'].fillna('')
    comparacao['texto_4'] = comparacao['texto_4'].fillna('')

    # Saldos

    comparacao['adp_saldo'] = (
        comparacao['adp_debito']
        - comparacao['adp_credito']
    ).round(2)

    comparacao['sap_saldo'] = (
        comparacao['sap_debito']
        - comparacao['sap_credito']
    ).round(2)

    comparacao['diferenca'] = (
        comparacao['adp_saldo']
        - comparacao['sap_saldo']
    ).round(2)

    comparacao = (
        comparacao
        .sort_values(['conta', 'ccusto'])
        .reset_index(drop=True)
    )

    # ── Agrupamentos ──────────────────────────────────────────────────────────

    colunas_soma = [
        'adp_debito', 'adp_credito',
        'sap_debito', 'sap_credito',
        'adp_saldo', 'sap_saldo', 'diferenca'
    ]

    # Por conta
    por_conta = (
        comparacao
        .groupby('conta', as_index=False)[colunas_soma]
        .sum()
        .round(2)
        .sort_values('conta')
        .reset_index(drop=True)
    )

    por_conta['Previsão do Modelo'] = por_conta.apply(diagnostico, axis=1)

    # Por centro de custo
    por_ccusto = (
        comparacao
        .groupby('ccusto', as_index=False)[colunas_soma]
        .sum()
        .round(2)
        .sort_values('ccusto')
        .reset_index(drop=True)
    )

    por_ccusto['Previsão do Modelo'] = por_ccusto.apply(diagnostico, axis=1)

    # Por texto (hist_lanc_4 do ADP — 4 primeiros caracteres do histórico)
    por_texto = (
        comparacao
        .groupby('hist_lanc_4', as_index=False)[colunas_soma]
        .sum()
        .round(2)
        .rename(columns={'hist_lanc_4': 'texto'})
        .sort_values('texto')
        .reset_index(drop=True)
    )

    # ── Salvar todas as abas ──────────────────────────────────────────────────

    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        comparacao.to_excel(writer, sheet_name='Comparacao', index=False)
        por_conta.to_excel(writer, sheet_name='Por_Conta', index=False)
        por_ccusto.to_excel(writer, sheet_name='Por_CCusto', index=False)
        por_texto.to_excel(writer, sheet_name='Por_Texto', index=False)

    print(f'Arquivo salvo em: {arquivo_saida}')
    print(f'  Abas: Comparacao ({len(comparacao)} linhas) | '
          f'Por_Conta ({len(por_conta)}) | '
          f'Por_CCusto ({len(por_ccusto)}) | '
          f'Por_Texto ({len(por_texto)})')
    
processamento()