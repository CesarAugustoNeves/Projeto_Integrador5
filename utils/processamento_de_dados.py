import pandas as pd
import numpy as np
import sys
from pathlib import Path

# ─── 1. ONDE O SISTEMA ESTÁ RODANDO? ───
if getattr(sys, 'frozen', False):
    # Rodando como .exe: Lê os arquivos da mesma pasta onde o executável está salvo
    pasta_trabalho = Path.cwd()
else:
    # Rodando no VS Code
    pasta_trabalho = Path(__file__).resolve().parent.parent

# ─── 2. CAMINHOS DOS ARQUIVOS DO USUÁRIO ───
# O sistema vai procurar a pasta "dados" ao lado de onde ele está rodando
arquivo_adp = pasta_trabalho / "dados" / "user" / "ADP.xlsx"
arquivo_sap = pasta_trabalho / "dados" / "user" / "SAP.xlsx"

# O rascunho da IA será salvo solto, na mesma pasta do .exe
arquivo_saida = pasta_trabalho / "IA.xlsx"

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

    # Agrupamentos
    # ── FUNÇÃO DE DIAGNÓSTICO RESTAURADA ──────────────────────────────────────
    def diagnostico(row):
        diff = abs(row.get('diferenca', 0))
        if diff < 0.01:
            return 1

        adp_dif = abs(row.get('adp_debito', 0)) > 0 or abs(row.get('adp_credito', 0)) > 0
        sap_dif = abs(row.get('sap_debito', 0)) > 0 or abs(row.get('sap_credito', 0)) > 0

        if adp_dif and not sap_dif:
            return 2
        if sap_dif and not adp_dif:
            return 3
        if adp_dif and sap_dif:
            return 4
        return 1

    comparacao['diagnostico'] = comparacao.apply(diagnostico, axis=1)
    # ──────────────────────────────────────────────────────────────────────────

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

    # Por centro de custo
    por_ccusto = (
        comparacao
        .groupby('ccusto', as_index=False)[colunas_soma]
        .sum()
        .round(2)
        .sort_values('ccusto')
        .reset_index(drop=True)
    )

    # ── AGRUPAMENTO POR EVENTOS ───────────────────────────────────────────────
    
    # 1. Extrair os códigos de evento (4 primeiros caracteres)
    adp['evento'] = adp['Hist Lanc'].astype(str).str.strip().str[:4]
    
    sap_credito['evento'] = sap_credito['Texto'].astype(str).str.strip().str[:4]
    sap_debito['evento'] = sap_debito['Texto'].astype(str).str.strip().str[:4]

    # 2. Somar Créditos e Débitos separados por evento
    adp_credito_evt = adp.groupby('evento', as_index=False)['valor'].sum().rename(columns={'valor': 'adp_credito'})
    sap_credito_evt = sap_credito.groupby('evento', as_index=False)['valor'].sum().rename(columns={'valor': 'sap_credito'})
    
    adp_debito_evt = adp.groupby('evento', as_index=False)['valor'].sum().rename(columns={'valor': 'adp_debito'})
    sap_debito_evt = sap_debito.groupby('evento', as_index=False)['valor'].sum().rename(columns={'valor': 'sap_debito'})

    # 3. Consolidar todos os códigos de eventos existentes
    eventos_unicos = pd.concat([
        adp_credito_evt[['evento']], adp_debito_evt[['evento']],
        sap_credito_evt[['evento']], sap_debito_evt[['evento']]
    ]).drop_duplicates()

    # 4. Mesclar as somas numa única tabela
    por_evento = (
        eventos_unicos
        .merge(adp_debito_evt, on='evento', how='left')
        .merge(adp_credito_evt, on='evento', how='left')
        .merge(sap_debito_evt, on='evento', how='left')
        .merge(sap_credito_evt, on='evento', how='left')
        .fillna(0)
    )

    # 5. Calcular os saldos finais e a diferença
    por_evento['adp_saldo'] = (por_evento['adp_debito'] - por_evento['adp_credito']).round(2)
    por_evento['sap_saldo'] = (por_evento['sap_debito'] - por_evento['sap_credito']).round(2)
    por_evento['diferenca'] = (por_evento['adp_saldo'] - por_evento['sap_saldo']).round(2)
    
    # 6. Aplicar a mesma função de diagnóstico do resto
    por_evento['Previsão do Modelo'] = por_evento.apply(diagnostico, axis=1)

    # Por texto (hist_lanc_4 do ADP — 4 primeiros caracteres do histórico)
    # Linhas sem texto ficam agrupadas como '' (em branco)
    por_texto = (
        comparacao
        .groupby('hist_lanc_4', as_index=False)[colunas_soma]
        .sum()
        .round(2)
        .rename(columns={'hist_lanc_4': 'texto'})
        .sort_values('texto')
        .reset_index(drop=True)
    )

    # Salvar todas as abas 

    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        comparacao.to_excel(writer, sheet_name='Comparacao', index=False)
        por_conta.to_excel(writer, sheet_name='Por_Conta', index=False)
        por_ccusto.to_excel(writer, sheet_name='Por_CCusto', index=False)
        por_evento.to_excel(writer, sheet_name='Por_Evento', index=False)

    print(f'Arquivo salvo em: {arquivo_saida}')
    print(f'  Abas: Comparacao ({len(comparacao)} linhas) | '
          f'Por_Conta ({len(por_conta)}) | '
          f'Por_CCusto ({len(por_ccusto)}) | '
          f'Por_Texto ({len(por_texto)})')