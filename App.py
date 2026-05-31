import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image  # Import necessário para carregar imagens
import pandas as pd
import os
import csv
import shutil
import pickle

from utils.processamento_de_dados import processamento
from ia.prever import prever_excel

class PayrollConciliator(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Payroll Conciliator IA - Dynatech")
        self.geometry("900x750")
        ctk.set_appearance_mode("light")

        # Variáveis dos caminhos
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
        # Frame para o logo (topo)
        self.logo_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.logo_frame.pack(pady=(20, 0))
        
        # Carregar e exibir a imagem (SEM usar self.log ainda)
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
        
        # Agora que o result_box existe, podemos mostrar logs
        self.carregar_logo_com_log()

    def carregar_logo_sem_log(self):
        """Carrega o logo sem usar o log (chamado antes do result_box existir)"""
        try:
            # Caminho da imagem (ajuste conforme necessário)
            caminho_logo = "logo_dynatech.png"  # Ou .jpg, .jpeg
            
            # Verificar se o arquivo existe
            if os.path.exists(caminho_logo):
                # Carregar imagem com PIL
                img = Image.open(caminho_logo)
                
                # Redimensionar a imagem (opcional - ajuste o tamanho conforme necessário)
                img = img.resize((300, 100), Image.Resampling.LANCZOS)
                
                # Converter para CTkImage
                logo = ctk.CTkImage(light_image=img, dark_image=img, size=(300, 100))
                
                # Criar label com a imagem
                logo_label = ctk.CTkLabel(self.logo_frame, image=logo, text="")
                logo_label.pack(pady=10)
            else:
                # Fallback: mostrar texto alternativo
                texto_logo = ctk.CTkLabel(self.logo_frame, text="DYNATECH", font=("Roboto", 20, "bold"))
                texto_logo.pack(pady=10)
                
        except Exception as e:
            # Fallback: mostrar texto
            texto_logo = ctk.CTkLabel(self.logo_frame, text="DYNATECH", font=("Roboto", 20, "bold"))
            texto_logo.pack(pady=10)
    
    def carregar_logo_com_log(self):
        """Tenta carregar o logo novamente para mostrar mensagem de sucesso/erro no log"""
        try:
            caminho_logo = "logo_dynatech.png"
        except Exception as e:
            self.log(f"Erro ao carregar logo: {str(e)}")

    # Cria as linhas de seleção
    def criar_linha(self, label_text, key, row):
        lbl = ctk.CTkLabel(self.frame_files, text=label_text)
        lbl.grid(row=row, column=0, padx=10, pady=5, sticky="w")
        
        btn = ctk.CTkButton(self.frame_files, text="Selecionar", width=100, command=lambda k=key: self.selecionar_arq(k))
        btn.grid(row=row, column=1, padx=10, pady=5)
        
        # Criar label para mostrar o nome do arquivo selecionado
        file_label = ctk.CTkLabel(self.frame_files, text="Nenhum arquivo selecionado", font=("Roboto", 10), text_color="black")
        file_label.grid(row=row, column=2, padx=10, pady=5, sticky="w")
        
        # Armazenar a referência da label
        self.file_labels[key] = file_label

    def selecionar_arq(self, key):
        path = filedialog.askopenfilename(title=f"Selecionar {key.upper()}")
        if path:
            self.paths[key] = path
            # Extrair apenas o nome do arquivo do caminho completo
            nome_arquivo = os.path.basename(path)
            
            # Atualizar a label com o nome do arquivo e cor verde
            if key in self.file_labels:
                self.file_labels[key].configure(text=f"✓ {nome_arquivo}", text_color="green", font=("Roboto", 10, "bold"))
            
            if key == "adp":
                destino_pasta = "dados/user"
                os.makedirs(destino_pasta, exist_ok=True)

                destino = os.path.join(destino_pasta, "ADP.xlsx")

                shutil.copy(path, destino)
            
            if key == "sap":
                destino_pasta = "dados/user"
                os.makedirs(destino_pasta, exist_ok=True)

                destino = os.path.join(destino_pasta, "SAP.xlsx")

                shutil.copy(path, destino)

            if key == "depara":
                destino_pasta = "dados/user"
                os.makedirs(destino_pasta, exist_ok=True)

                destino = os.path.join(destino_pasta, "DEPARA.xlsx")

                shutil.copy(path, destino)

            if key == "eventos":
                destino_pasta = "dados/user"
                os.makedirs(destino_pasta, exist_ok=True)

                destino = os.path.join(destino_pasta, "EVENTOS.xlsx")

                shutil.copy(path, destino)

            # agora você troca o caminho para o interno do projeto
            self.paths[key] = destino

            self.log(f"Arquivo {key.upper()} carregado com sucesso: {nome_arquivo}")

    def log(self, message):
        if hasattr(self, 'result_box'):  # Verifica se o result_box já existe
            self.result_box.insert("end", f"> {message}\n")
            self.result_box.see("end")
        else:
            print(f"> {message}")  # Fallback para o console

    def limpar_valor(self, val):
        if pd.isna(val): return 0.0
        s = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
        try: return float(s)
        except: return 0.0

    def logica(self):
        # Verifica se todos os arquivos foram selecionados
        """if not all(self.paths.values()):
            messagebox.showwarning("Atenção", "Selecione os 4 arquivos antes de continuar!")
            return
        """
        try:
            self.log("Iniciando leitura segura dos arquivos...")
            
            '''# Lendo ADP
            self.log("-> Lendo MATRIZ ADP...")
            df_adp = pd.read_csv(self.paths["adp"], sep=';', encoding='latin1', engine='python', on_bad_lines='skip')
   
            
            # Lendo SAP
            self.log("-> Lendo MATRIZ SAP...")
            df_sap = pd.read_csv(self.paths["sap"], sep=';', encoding='latin1', skiprows=5, engine='python', on_bad_lines='skip')
            
            # Lendo De-Para CC
            self.log("-> Lendo De-Para CC...")
            df_cc = pd.read_csv(self.paths["depara"], sep=';', encoding='latin1', engine='python', on_bad_lines='skip')
            
            # Lendo Eventos
            self.log("-> Lendo Eventos ADP...")
            df_ev = pd.read_csv(self.paths["eventos"], sep=';', encoding='latin1', engine='python', on_bad_lines='skip')
            '''

            self.log("SUCESSO: Todos os arquivos foram lidos.")
            self.log("Iniciando o cruzamento de dados...")
            
            processamento()

            '''prever_excel('dados/user/IA.xlsx')'''

            self.log("Análise concluída com sucesso.")

        except Exception as e:
            self.log(f"ERRO CRÍTICO: {str(e)}")

if __name__ == "__main__":
    app = PayrollConciliator()
    app.mainloop()