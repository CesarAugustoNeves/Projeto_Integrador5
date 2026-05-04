import customtkinter as ctk
from tkinter import filedialog
import pandas as pd

class AppConciliador(ctk.CTk):
    def __init__(self):
        super().__init__()
        #Configuração da janela
        self.title("Payroll Conciliator IA - Dynatech")
        self.geometry("800x600")

        # Configuração de Layout
        self.grid_columnconfigure(0, weight=1)
        
        self.label = ctk.CTkLabel(self, text="Conciliação de Folha de Pagamento", font=("Roboto", 24))
        self.label.pack(pady=20)

        # Botões de Upload
        self.btn_adp = ctk.CTkButton(self, text="Selecionar Arquivo ADP", command=self.upload_adp)
        self.btn_adp.pack(pady=10)

        self.btn_sap = ctk.CTkButton(self, text="Selecionar Arquivo SAP", command=self.upload_sap)
        self.btn_sap.pack(pady=10)

        self.btn_processar = ctk.CTkButton(self, text="Iniciar Conciliação", fg_color="green", command=self.processar)
        self.btn_processar.pack(pady=20)

        # Área de Resultado
        self.textbox = ctk.CTkTextbox(self, width=700, height=300)
        self.textbox.pack(pady=10)

    def upload_adp(self):
        filename = filedialog.askopenfilename()
        self.textbox.insert("insert", f"ADP Selecionado: {filename}\n")

    def upload_sap(self):
        filename = filedialog.askopenfilename()
        self.textbox.insert("insert", f"SAP Selecionado: {filename}\n")

    def processar(self):
        #Saida de texto
        self.textbox.insert("insert", "Processando divergências...\n")
        

if __name__ == "__main__":
    app = AppConciliador()
    app.mainloop()