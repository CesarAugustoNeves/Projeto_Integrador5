import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import pandas as pd
import os
import shutil
#Comando pra ativar venv:
#.\.venv\Scripts\activate

#Comando pra ativar o executavel(dentro da venv): 
#pyinstaller --noconsole --onefile --hidden-import customtkinter --collect-all customtkinter --hidden-import PIL --collect-all PIL --hidden-import sklearn --collect-all sklearn --hidden-import openpyxl --add-data "ia/modelo.pkl;ia" --add-data "dados/output/Resultado.xlsx;dados/output" App.py

from utils.processamento_de_dados import processamento
from ia.prever import prever_excel
from utils.output_excel import output_excel

# Arquivos que já existem 
ARQUIVOS_FIXOS = {"depara", "plano"}

_DESTINOS = {
    "adp":    "ADP.xlsx",
    "sap":    "SAP.xlsx",
    "depara": "DEPARA.xlsx",
    "plano":  "PLANO_CONTAS.xlsx",
}

class FileCard(ctk.CTkFrame):
    """Card que exibe o status de um arquivo com ações contextuais."""

    def __init__(self, master, label_text, key, on_select, on_remove, fixo=False, **kwargs):
        super().__init__(master, fg_color="#f0f4f8", corner_radius=8, **kwargs)

        self.key = key
        self.on_select = on_select
        self.on_remove = on_remove
        self.fixo = fixo

        self.lbl_title = ctk.CTkLabel(
            self, text=label_text,
            font=("Roboto", 12, "bold"), width=160, anchor="w"
        )
        self.lbl_title.grid(row=0, column=0, padx=(12, 6), pady=8, sticky="w")

        # Arquivos fixos: botão "Substituir" (troca o arquivo)
        if fixo:
            self.btn_select = ctk.CTkButton(
                self, text="⇄ Substituir", width=110, height=30,
                fg_color="#546e7a", hover_color="#37474f",
                command=lambda: self.on_select(self.key)
            )
        else:
            self.btn_select = ctk.CTkButton(
                self, text="Selecionar", width=110, height=30,
                fg_color="#1f6aa5", hover_color="#174f7a",
                command=lambda: self.on_select(self.key)
            )
        self.btn_select.grid(row=0, column=1, padx=6, pady=8)

        # Nome / status do arquivo
        self.lbl_file = ctk.CTkLabel(
            self, text="Nenhum arquivo selecionado",
            font=("Roboto", 10), text_color="#888888",
            anchor="w", wraplength=300
        )
        self.lbl_file.grid(row=0, column=2, padx=6, pady=8, sticky="w")

        # Botão remover para arquivos adp e sap
        self.btn_remove = ctk.CTkButton(
            self, text="✕ Remover", width=90, height=28,
            fg_color="#e53935", hover_color="#b71c1c",
            command=lambda: self.on_remove(self.key)
        )
        self.btn_remove.grid(row=0, column=3, padx=(6, 12), pady=8)
        self.btn_remove.grid_remove()

        self.columnconfigure(2, weight=1)



    def set_arquivo(self, nome_arquivo: str, fixo_automatico: bool = False):
        """Marca o card como carregado."""
        if fixo_automatico:
            self.lbl_file.configure(
                text=f"📁  {nome_arquivo}  (pré-carregado)",
                text_color="#1565c0",
                font=("Roboto", 10, "bold")
            )
            # Arquivos fixos não têm botão remover — só substituir
        else:
            self.lbl_file.configure(
                text=f"✓  {nome_arquivo}",
                text_color="#2e7d32",
                font=("Roboto", 10, "bold")
            )
            if not self.fixo:
                self.btn_remove.grid()

    def limpar(self):
        """Reseta o card para o estado vazio."""
        self.lbl_file.configure(
            text="Nenhum arquivo selecionado",
            text_color="#888888",
            font=("Roboto", 10)
        )
        self.btn_remove.grid_remove()


class PayrollConciliator(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Payroll Conciliator IA - Dynatech")
        self.geometry("960x780")
        ctk.set_appearance_mode("light")

        self.paths: dict[str, str] = {k: "" for k in _DESTINOS}
        self.file_cards: dict[str, FileCard] = {}

        self.setup_ui()
        self._detectar_arquivos_fixos()


    def setup_ui(self):
        # Logo
        self.logo_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.logo_frame.pack(pady=(20, 0))
        self._carregar_logo()

        # Título
        ctk.CTkLabel(
            self,
            text="Painel de Auditoria Dynatech",
            font=("Roboto", 24, "bold")
        ).pack(pady=(8, 4))


        # Cards
        self.frame_files = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_files.pack(pady=6, padx=24, fill="x")

        linhas = [
            ("1. Matriz ADP:",      "adp",    False),
            ("2. Matriz SAP:",      "sap",    False),
            ("3. De-Para CC:",      "depara", True),
            ("4. Plano de Contas:", "plano",  True),
        ]
        for label, key, fixo in linhas:
            card = FileCard(
                self.frame_files,
                label_text=label,
                key=key,
                on_select=self.selecionar_arq,
                on_remove=self.remover_arq,
                fixo=fixo
            )
            card.pack(fill="x", pady=4)
            self.file_cards[key] = card

        # Botão executar
        self.btn_run = ctk.CTkButton(
            self, text="▶  EXECUTAR CONCILIAÇÃO",
            fg_color="#1f6aa5", hover_color="#174f7a",
            height=44, font=("Roboto", 14, "bold"),
            command=self.logica
        )
        self.btn_run.pack(pady=20)

        # Log
        self.result_box = ctk.CTkTextbox(self, width=900, height=300, font=("Consolas", 11))
        self.result_box.pack(pady=8, padx=20)

    def _carregar_logo(self):
        try:
            caminho_logo = "logo_dynatech.png"
            if os.path.exists(caminho_logo):
                img = Image.open(caminho_logo).resize((300, 100), Image.Resampling.LANCZOS)
                logo = ctk.CTkImage(light_image=img, dark_image=img, size=(300, 100))
                ctk.CTkLabel(self.logo_frame, image=logo, text="").pack(pady=6)
            else:
                ctk.CTkLabel(self.logo_frame, text="DYNATECH", font=("Roboto", 20, "bold")).pack(pady=6)
        except Exception:
            ctk.CTkLabel(self.logo_frame, text="DYNATECH", font=("Roboto", 20, "bold")).pack(pady=6)


    def _detectar_arquivos_fixos(self):
        """Verifica se DEPARA e PLANO_CONTAS já existem em dados/user e atualiza os cards."""
        for key in ARQUIVOS_FIXOS:
            caminho = os.path.join("dados/user", _DESTINOS[key])
            if os.path.exists(caminho):
                self.paths[key] = caminho
                self.file_cards[key].set_arquivo(_DESTINOS[key], fixo_automatico=True)
                self.log(f"Arquivo {key.upper()} detectado automaticamente: {_DESTINOS[key]}")
            else:
                self.log(f"⚠ Arquivo {_DESTINOS[key]} não encontrado em dados/user. Use '⇄ Substituir' para carregar.")


    def selecionar_arq(self, key: str):
        path = filedialog.askopenfilename(title=f"Selecionar {key.upper()}")
        if not path:
            return

        destino_pasta = "dados/user"
        os.makedirs(destino_pasta, exist_ok=True)
        destino = os.path.join(destino_pasta, _DESTINOS[key])
        shutil.copy(path, destino)

        self.paths[key] = destino
        nome_arquivo = os.path.basename(path)

        self.file_cards[key].set_arquivo(nome_arquivo, fixo_automatico=False)
        self.log(f"Arquivo {key.upper()} {'substituído' if key in ARQUIVOS_FIXOS else 'carregado'}: {nome_arquivo}")

    def remover_arq(self, key: str):
        """Remove arquivo da UI e do disco (apenas para ADP e SAP)."""
        destino = self.paths.get(key, "")
        if destino and os.path.exists(destino):
            try:
                os.remove(destino)
            except Exception as e:
                self.log(f"Aviso: não foi possível remover o arquivo do disco: {e}")

        self.paths[key] = ""
        self.file_cards[key].limpar()
        self.log(f"Arquivo {key.upper()} removido.")


    def log(self, message: str):
        if hasattr(self, "result_box"):
            self.result_box.insert("end", f"> {message}\n")
            self.result_box.see("end")
        else:
            print(f"> {message}")

    def limpar_valor(self, val):
        if pd.isna(val):
            return 0.0
        s = str(val).replace("R$", "").replace(".", "").replace(",", ".").strip()
        try:
            return float(s)
        except Exception:
            return 0.0


    def logica(self):
        self.btn_run.configure(
            text="⏳ Processando...",
            fg_color="#f9a825",
            hover_color="#f57f17"
        )

        self.update()  # força UI atualizar

        try:
            self.log("Iniciando leitura segura dos arquivos...")
            self.log("SUCESSO: Todos os arquivos foram lidos.")
            self.log("Iniciando o cruzamento de dados...")

            # ajusta os arquivos adp e sap em 1 só
            processamento()

            #chama a ia
            prever_excel("IA.xlsx", sheet_name="Por_Conta")
            prever_excel("IA.xlsx", sheet_name="Por_CCusto")
            prever_excel("IA.xlsx", sheet_name="Por_Evento") 

            #cria o arquivo excel de output
            output_excel()

            self.log("Análise concluída com sucesso.")
            
            self.btn_run.configure(
                text="✔ EXECUÇÃO COMPLETA",
                fg_color="#2e7d32",
                hover_color="#1b5e20"
            )

        except Exception as e:
            self.log(f"ERRO CRÍTICO: {str(e)}")

            self.btn_run.configure(
                text="❌ ERRO NA EXECUÇÃO",
                fg_color="#c62828",
                hover_color="#8e0000"
            )


if __name__ == "__main__":
    app = PayrollConciliator()
    app.mainloop()