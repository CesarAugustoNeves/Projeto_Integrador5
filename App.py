import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import pandas as pd
import os
import csv
import pickle
import openpyxl

#Copiar isso aqui pra ativar o executavel(Tem que estar com a venv ativada na raiz do projeto! Digite no terminal que está com a venv ativada.):
#pyinstaller --noconsole --onefile --hidden-import customtkinter --collect-all customtkinter --hidden-import PIL --collect-all PIL --hidden-import sklearn --collect-all sklearn --hidden-import openpyxl --add-data "ia/modelo.pkl;ia" --add-data "logo_dynatech.png;." App.py

from utils.processamento_de_dados import processamento
from ia.prever import prever_excel

class PayrollConciliator(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Payroll Conciliator IA - Dynatech")
        self.geometry("900x750")
        ctk.set_appearance_mode("light")

        # Variáveis dos caminhos (agora vão guardar os caminhos reais absolutos)
        self.paths = {
            "adp": "",
            "sap": "",
            "depara": "",
            "eventos": ""
        }
        
        # Dicionário para armazenar as labels dos arquivos
        self.file_labels = {}

        self.setup_ui()

    def setup_ui(self):
        # Frame para o logo 
        self.logo_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.logo_frame.pack(pady=(20, 0))
        
        # Carregar e exibir a imagem 
        self.carregar_logo_sem_log()
        
        # Título principal
        self.label_title = ctk.CTkLabel(self, text="Painel de Auditoria Dynatech", font=("Roboto", 24, "bold"))
        self.label_title.pack(pady=10)

        # Container de Botões
        self.frame_files = ctk.CTkFrame(self)
        self.frame_files.pack(pady=10, padx=20, fill="x")

        # Config das linhas de seleção
        self.criar_linha("1. Matriz ADP:", "adp", 0)
        self.criar_linha("2. Matriz SAP:", "sap", 1)
        self.criar_linha("3. De-Para CC:", "depara", 2)
        self.criar_linha("4. Eventos ADP:", "eventos", 3)

        self.btn_run = ctk.CTkButton(self, text="EXECUTAR CONCILIAÇÃO", fg_color="#1f6aa5", height=40, command=self.logica)
        self.btn_run.pack(pady=30)

        self.result_box = ctk.CTkTextbox(self, width=850, height=350, font=("Consolas", 12))
        self.result_box.pack(pady=10, padx=20)
        
        self.carregar_logo_com_log()

    def carregar_logo_sem_log(self):
        """Carrega o logo sem usar o log (chamado antes do result_box existir)"""
        import sys
        try:
            # Blindagem do caminho para o executável
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
                
            caminho_logo = os.path.join(base_path, "logo_dynatech.png")
            
            if os.path.exists(caminho_logo):
                img = Image.open(caminho_logo)
                img = img.resize((300, 100), Image.Resampling.LANCZOS)
                logo = ctk.CTkImage(light_image=img, dark_image=img, size=(300, 100))
                logo_label = ctk.CTkLabel(self.logo_frame, image=logo, text="")
                logo_label.pack(pady=10)
            else:
                texto_logo = ctk.CTkLabel(self.logo_frame, text="DYNATECH", font=("Roboto", 20, "bold"))
                texto_logo.pack(pady=10)
        except Exception:
            texto_logo = ctk.CTkLabel(self.logo_frame, text="DYNATECH", font=("Roboto", 20, "bold"))
            texto_logo.pack(pady=10)
    
    def carregar_logo_com_log(self):
        """Tenta carregar o logo novamente para mostrar mensagem de sucesso/erro no log"""
        try:
            caminho_logo = "logo_dynatech.png"
        except Exception as e:
            self.log(f"Erro ao carregar logo: {str(e)}")

    def criar_linha(self, label_text, key, row):
        lbl = ctk.CTkLabel(self.frame_files, text=label_text)
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        btn = ctk.CTkButton(self.frame_files, text="Selecionar", width=100, command=lambda k=key: self.selecionar_arq(k))
        btn.grid(row=row, column=1, padx=10, pady=5)
        
        file_label = ctk.CTkLabel(self.frame_files, text="Nenhum arquivo selecionado", font=("Roboto", 10), text_color="black")
        file_label.grid(row=row, column=2, padx=10, pady=5, sticky="w")
        
        self.file_labels[key] = file_label

    def selecionar_arq(self, key):
        path = filedialog.askopenfilename(title=f"Selecionar {key.upper()}")
        if path:
            # Salva o caminho absoluto real direto na variável, sem copiar nada
            self.paths[key] = path
            nome_arquivo = os.path.basename(path)
            
            if key in self.file_labels:
                self.file_labels[key].configure(text=f"✓ {nome_arquivo}", text_color="green", font=("Roboto", 10, "bold"))

            self.log(f"Arquivo {key.upper()} carregado com sucesso: {nome_arquivo}")

    def log(self, message):
        if hasattr(self, 'result_box'):
            self.result_box.insert("end", f"> {message}\n")
            self.result_box.see("end")
        else:
            print(f"> {message}")

    def limpar_valor(self, val):
        if pd.isna(val): return 0.0
        s = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
        try: return float(s)
        except: return 0.0

    def logica(self):
        if not self.paths["adp"] or not self.paths["sap"]:
            messagebox.showwarning("Atenção", "Selecione pelo menos as matrizes ADP e SAP antes de continuar!")
            return
        
        try:
            self.log("SUCESSO: Arquivos validados. Iniciando cruzamento de dados...")
            
            # Definindo onde o arquivo final será salvo (na mesma pasta do executável)
            caminho_saida = "Resultado_Auditoria_IA.xlsx"
            
            # Passando os caminhos que o usuário escolheu na tela direto para as funções
            processamento(
                arquivo_adp=self.paths["adp"], 
                arquivo_sap=self.paths["sap"],
                arquivo_saida=caminho_saida
            )

            self.log("Cruzamento concluído. Iniciando IA de predição...")

            # A IA agora lê o arquivo gerado e sobrescreve ele
            prever_excel(caminho_saida)

            self.log(f"Análise concluída com sucesso! Arquivo gerado: {caminho_saida}")

        except Exception as e:
            self.log(f"ERRO CRÍTICO: {str(e)}")

if __name__ == "__main__":
    app = PayrollConciliator()
    app.mainloop()