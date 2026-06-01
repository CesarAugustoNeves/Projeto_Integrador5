import pandas as pd
from pathlib import Path
import shutil
import sys
from openpyxl import load_workbook

def output_excel():
    # 1. A FUNÇÃO DE RADAR (Descobre onde o código está rodando)
    if getattr(sys, 'frozen', False):
        # Rodando como .exe
        caminho_base = Path(sys._MEIPASS) # Pasta invisível do sistema
        pasta_raiz = Path.cwd()           # Pasta real onde o usuário clicou no .exe
    else:
        # Rodando no VS Code
        caminho_base = Path(__file__).resolve().parent.parent
        pasta_raiz = caminho_base

    # 2. ARQUIVOS DE SISTEMA (Que viajam dentro do .exe)
    template_path = caminho_base / "dados" / "output" / "Resultado.xlsx"

    # 3. ARQUIVOS DINÂMICOS (Criados soltos ao lado do .exe)
    arquivo_ia = pasta_raiz / "IA.xlsx"
    arquivo_saida = pasta_raiz / "Resultado_Auditoria_IA.xlsx"
    
    # 4. ARQUIVOS DO USUÁRIO (Lidos da pasta dados/user real do computador)
    pasta_user = pasta_raiz / "dados" / "user"
    arquivo_plano = pasta_user / "PLANO_CONTAS.xlsx"
    arquivo_depara = pasta_user / "DEPARA.xlsx"
    arquivo_map_eventos = pasta_user / "EVENTOS.xlsx"

    # 5. COPIA O TEMPLATE EM BRANCO PARA O ARQUIVO FINAL
    shutil.copy2(template_path, arquivo_saida)

    # 6. LÊ OS DADOS GERADOS PELA IA (O arquivo intermediário)
    por_conta  = pd.read_excel(arquivo_ia, sheet_name="Por_Conta")
    por_ccusto = pd.read_excel(arquivo_ia, sheet_name="Por_CCusto")
    por_evento = pd.read_excel(arquivo_ia, sheet_name="Por_Evento")

    # 7. MAPEAMENTOS BLINDADOS (Trazendo as descrições)
    
    # -- Mapa de Eventos --
    if arquivo_map_eventos.exists():
        try:
            map_eventos = pd.read_excel(arquivo_map_eventos, sheet_name="Evento ADP")
            map_eventos.columns = map_eventos.columns.astype(str).str.strip()
            map_eventos = map_eventos[['Código do Evento', 'Descrição do Evento']].rename(columns={'Código do Evento': 'evento', 'Descrição do Evento': 'descricao'})
            
            map_eventos['evento'] = map_eventos['evento'].astype(str).str.zfill(4)
            por_evento['evento'] = por_evento['evento'].astype(str).str.zfill(4)
            por_evento = por_evento.merge(map_eventos, on='evento', how='left')
        except Exception as e:
            print(f"Aviso: Colunas do Mapa de Eventos não encontradas. Erro: {e}")
            por_evento['descricao'] = ""
    else:
        por_evento['descricao'] = ""

    # -- Mapa do Plano de Contas --
    try:
        df_plano = pd.read_excel(arquivo_plano)
        # Força tudo para maiúsculo e tira espaços
        df_plano.columns = df_plano.columns.astype(str).str.strip().str.upper()
        
        plano = (
            df_plano[["CONTA", "BALANCETE"]]
            .rename(columns={"CONTA": "conta", "BALANCETE": "descricao"})
            .dropna(subset=["conta"])
        )
        plano["conta"] = pd.to_numeric(plano["conta"], errors='coerce').fillna(0).astype(int)
        por_conta = por_conta.merge(plano, on="conta", how="left")
    except Exception as e:
        print(f"Aviso: Colunas CONTA ou BALANCETE não encontradas no Plano de Contas. Erro: {e}")
        if 'descricao' not in por_conta.columns:
            por_conta['descricao'] = ""

    # -- Mapa de Centros de Custo (DePara) --
    try:
        df_depara = pd.read_excel(arquivo_depara)
        # Força tudo para maiúsculo e tira espaços
        df_depara.columns = df_depara.columns.astype(str).str.strip().str.upper()
        
        depara = (
            df_depara[["CENTRO DE CUSTO", "DESCRITIVO TESTE"]]
            .rename(columns={"CENTRO DE CUSTO": "ccusto", "DESCRITIVO TESTE": "descricao"})
            .dropna(subset=["ccusto"])
        )
        depara["ccusto"] = pd.to_numeric(depara["ccusto"], errors='coerce').fillna(0).astype(int)
        por_ccusto = por_ccusto.merge(depara, on="ccusto", how="left")
    except Exception as e:
        print(f"Aviso: Colunas do DePara não encontradas. Erro: {e}")
        if 'descricao' not in por_ccusto.columns:
            por_ccusto['descricao'] = ""

    # 8. ESCRITA SEGURA NO EXCEL
    wb = load_workbook(arquivo_saida)

    # Preenche aba: Por Evento
    if "Por Evento" in wb.sheetnames:
        ws_evento = wb["Por Evento"] 
        for i, (_, row) in enumerate(por_evento.iterrows(), start=3):
            ws_evento[f"A{i}"] = row.get("evento", "")
            ws_evento[f"B{i}"] = row.get("descricao", "") if pd.notna(row.get("descricao")) else ""
            ws_evento[f"C{i}"] = row.get("adp_debito", 0)
            ws_evento[f"D{i}"] = row.get("adp_credito", 0)
            ws_evento[f"E{i}"] = row.get("adp_saldo", 0)
            ws_evento[f"F{i}"] = row.get("sap_debito", 0)
            ws_evento[f"G{i}"] = row.get("sap_credito", 0)
            ws_evento[f"H{i}"] = row.get("sap_saldo", 0)
            ws_evento[f"I{i}"] = row.get("diferenca", 0)
            ws_evento[f"J{i}"] = row.get("Previsão do Modelo", "")

    # Preenche aba: Por Conta
    if "Por Conta" in wb.sheetnames:
        ws_conta = wb["Por Conta"]
        for i, (_, row) in enumerate(por_conta.iterrows(), start=3):
            ws_conta[f"A{i}"] = row.get("conta", "")
            ws_conta[f"B{i}"] = row.get("descricao", "") if pd.notna(row.get("descricao")) else ""
            ws_conta[f"C{i}"] = row.get("adp_debito", 0)
            ws_conta[f"D{i}"] = row.get("adp_credito", 0)
            ws_conta[f"E{i}"] = row.get("adp_saldo", 0)
            ws_conta[f"F{i}"] = row.get("sap_debito", 0)
            ws_conta[f"G{i}"] = row.get("sap_credito", 0)
            ws_conta[f"H{i}"] = row.get("sap_saldo", 0)
            ws_conta[f"I{i}"] = row.get("diferenca", 0)
            ws_conta[f"J{i}"] = row.get("Previsão do Modelo", "")

    # Preenche aba: Por Centro de Custo
    if "Por Centro de Custo" in wb.sheetnames:
        ws_ccusto = wb["Por Centro de Custo"]
        for i, (_, row) in enumerate(por_ccusto.iterrows(), start=3):
            ws_ccusto[f"A{i}"] = row.get("ccusto", "")
            ws_ccusto[f"B{i}"] = row.get("descricao", "") if pd.notna(row.get("descricao")) else ""
            ws_ccusto[f"C{i}"] = row.get("adp_debito", 0)
            ws_ccusto[f"D{i}"] = row.get("adp_credito", 0)
            ws_ccusto[f"E{i}"] = row.get("adp_saldo", 0)
            ws_ccusto[f"F{i}"] = row.get("sap_debito", 0)
            ws_ccusto[f"G{i}"] = row.get("sap_credito", 0)
            ws_ccusto[f"H{i}"] = row.get("sap_saldo", 0)
            ws_ccusto[f"I{i}"] = row.get("diferenca", 0)
            ws_ccusto[f"J{i}"] = row.get("Previsão do Modelo", "")

    # Salva o arquivo final
    wb.save(arquivo_saida)
    print(f"Arquivo gerado com sucesso: {arquivo_saida}")