import pandas as pd
import numpy as np

def processamento(arquivo_adp, arquivo_sap, arquivo_saida):

    # SCANNER COMPLETO 
    def buscar_aba_e_linha(caminho, coluna_alvo):
        xls = pd.ExcelFile(caminho) 
        
        for aba in xls.sheet_names:
            for linha_cabecalho in range(6):
                try:
                    df_temp = pd.read_excel(caminho, sheet_name=aba, header=linha_cabecalho, nrows=0)
                    
                    colunas_limpas = (df_temp.columns.astype(str)
                                      .str.strip()
                                      .str.replace('é', 'e', regex=False)
                                      .str.replace('É', 'E', regex=False)
                                      .str.replace('  ', ' ', regex=False))
                    
                    if coluna_alvo in colunas_limpas:
                        return aba, linha_cabecalho
                except Exception:
                    continue
                    
        raise ValueError(f"ERRO CRÍTICO: A coluna '{coluna_alvo}' não foi encontrada em nenhuma aba e em nenhuma linha lida!")

    # 1. ADP Scanner
    aba_certa_adp, header_certo_adp = buscar_aba_e_linha(arquivo_adp, coluna_alvo='Cta Credito/40')
    adp = pd.read_excel(arquivo_adp, sheet_name=aba_certa_adp, header=header_certo_adp)

    # 2. SAP Scanner
    aba_certa_sap, header_certo_sap = buscar_aba_e_linha(arquivo_sap, coluna_alvo='Conta')
    sap = pd.read_excel(arquivo_sap, sheet_name=aba_certa_sap, header=header_certo_sap)

    # FILTRO DE CABEÇALHOS GERAIS 
    adp.columns = adp.columns.astype(str).str.strip().str.replace('é', 'e', regex=False).str.replace('É', 'E', regex=False).str.replace('  ', ' ', regex=False)
    sap.columns = sap.columns.astype(str).str.strip().str.replace('é', 'e', regex=False).str.replace('É', 'E', regex=False).str.replace('  ', ' ', regex=False)

    

    # MAPEAMENTO ADP 
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
        adp['Cta Debito/50']  # Este estava certo desde o começo
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

    # --- O MAPEAMENTO SAP ---
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

    # --- CRUZAMENTOS ---
    sap_credito = sap[(sap['chave'] == 40) & (sap['Centro custo'].notna())].copy()
    sap_debito = sap[(sap['chave'] == 50) & (sap['Centro custo'].notna())].copy()

    adp_credito = (
        adp.groupby(['conta_credito', 'ccusto_credito'], as_index=False)['valor']
        .sum()
        .rename(columns={'valor': 'adp_valor'})
    )

    sap_credito_agr = (
        sap_credito.groupby(['conta', 'ccusto'], as_index=False)['valor']
        .sum()
        .rename(columns={'valor': 'sap_valor'})
    )

    sap_credito_agr['sap_valor_abs'] = sap_credito_agr['sap_valor'].abs()

    comparacao_credito = adp_credito.merge(
        sap_credito_agr,
        left_on=['conta_credito', 'ccusto_credito'],
        right_on=['conta', 'ccusto'],
        how='outer'
    ).fillna(0)

    comparacao_credito['diferenca'] = (comparacao_credito['adp_valor'] - comparacao_credito['sap_valor_abs']).round(2)

    adp_debito = (
        adp.groupby(['conta_debito', 'ccusto_debito'], as_index=False)['valor']
        .sum()
        .rename(columns={'valor': 'adp_valor'})
    )

    sap_debito_agr = (
        sap_debito.groupby(['conta', 'ccusto'], as_index=False)['valor']
        .sum()
        .rename(columns={'valor': 'sap_valor'})
    )

    sap_debito_agr['sap_valor_abs'] = sap_debito_agr['sap_valor'].abs()

    comparacao_debito = adp_debito.merge(
        sap_debito_agr,
        left_on=['conta_debito', 'ccusto_debito'],
        right_on=['conta', 'ccusto'],
        how='outer'
    ).fillna(0)

    comparacao_debito['diferenca'] = (comparacao_debito['adp_valor'] - comparacao_debito['sap_valor_abs']).round(2)

    comparacao_credito = comparacao_credito.rename(columns={'conta_credito': 'conta_adp', 'ccusto_credito': 'ccusto_adp'})
    comparacao_debito = comparacao_debito.rename(columns={'conta_debito': 'conta_adp', 'ccusto_debito': 'ccusto_adp'})

    comparacao = pd.concat([comparacao_credito, comparacao_debito], ignore_index=True)

    def diagnostico(row):
        diff = abs(row['diferenca'])
        if diff < 0.01:
            return 1

        adp_dif = row['adp_valor'] > 0
        sap_dif = row['sap_valor_abs'] > 0

        if adp_dif and not sap_dif:
            return 2
        if sap_dif and not adp_dif:
            return 3
        if adp_dif and sap_dif:
            return 4
        return 1

    comparacao['diagnostico'] = comparacao.apply(diagnostico, axis=1)

    comparacao.to_excel(arquivo_saida, index=False)