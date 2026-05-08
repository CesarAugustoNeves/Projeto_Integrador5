import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import csv

class PayrollConciliator(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Payroll Conciliator IA - Dynatech")
        self.geometry("900x750")
        ctk.set_appearance_mode("dark")

        # Variáveis dos caminhos
        self.paths = {
            "adp": "",
            "sap": "",
            "depara": "",
            "eventos": ""
        }

        self.setup_ui()

    def setup_ui(self):
        self.label_title = ctk.CTkLabel(self, text="Painel de Auditoria Dynatech", font=("Roboto", 24, "bold"))
        self.label_title.pack(pady=20)

        # Container de Botões
        self.frame_files = ctk.CTkFrame(self)
        self.frame_files.pack(pady=10, padx=20, fill="x")

        # Config das linhas de seleção
        self.criar_linha("1. Matriz ADP:", "adp", 0)
        self.criar_linha("2. Matriz SAP:", "sap", 1)
        self.criar_linha("3. De-Para CC:", "depara", 2)
        self.criar_linha("4. Eventos ADP:", "eventos", 3)

        self.btn_run = ctk.CTkButton(self, text="EXECUTAR CONCILIAÇÃO", fg_color="#1f6aa5", height=40, command=self.run_logic)
        self.btn_run.pack(pady=30)

        self.result_box = ctk.CTkTextbox(self, width=850, height=350, font=("Consolas", 12))
        self.result_box.pack(pady=10, padx=20)

    # Cria as linhas de seleção
    def criar_linha(self, label_text, key, row):
        lbl = ctk.CTkLabel(self.frame_files, text=label_text)
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        btn = ctk.CTkButton(self.frame_files, text="Selecionar", width=100, command=lambda k=key: self.select_file(k))
        btn.grid(row=row, column=1, padx=10, pady=5)

    def selecionar_arq(self, key):
        path = filedialog.askopenfilename(title=f"Selecionar {key.upper()}")
        if path:
            self.paths[key] = path
            self.log(f"Arquivo {key.upper()} carregado com sucesso.")

    def log(self, message):
        self.result_box.insert("end", f"> {message}\n")
        self.result_box.see("end")

    def limpar_valor(self, val):
        if pd.isna(val): return 0.0
        s = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
        try: return float(s)
        except: return 0.0

    def logica(self):
        # Verifica se todos os arquivos foram selecionados
        if not all(self.paths.values()):
            messagebox.showwarning("Atenção", "Selecione os 4 arquivos antes de continuar!")
            return

        try:
            self.log("Iniciando leitura segura dos arquivos...")
            
            # Lendo ADP
            self.log("-> Lendo MATRIZ ADP...")
            df_adp = pd.read_csv(self.paths["adp"], sep=';', encoding='latin1', engine='python', on_bad_lines='skip')
            
            # Lendo SAP
            self.log("-> Lendo MATRIZ SAP...")
            df_sap = pd.read_csv(self.paths["sap"], sep=';', encoding='latin1', skiprows=5, engine='python', on_bad_lines='skip')
            
            # Lendo De-Para CC
            self.log("-> Lendo De-Para CC...")
            df_cc = pd.read_csv(self.paths["depara"], sep=';', encoding='latin1', engine='python', on_bad_lines='skip')
            
            # Lendo Eventos )
            self.log("-> Lendo Eventos ADP...")
            df_ev = pd.read_csv(self.paths["eventos"], sep=';', encoding='latin1', engine='python', on_bad_lines='skip')

            self.log("SUCESSO: Todos os arquivos foram lidos.")
            self.log("Iniciando o cruzamento de dados...")
            
            # COLOCAR CODIGO DA CONCILIAÇÃO AQUI DEPOIS
            
            self.log("Análise concluída com sucesso.")

        except Exception as e:
            self.log(f"ERRO CRÍTICO: {str(e)}")

if __name__ == "__main__":
    app = PayrollConciliator()
    app.mainloop()