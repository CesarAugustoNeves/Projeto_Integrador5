# Projeto_Integrador5
Repositório para o software de conciliação de folha de pagamento ADP-SAP.


# Configurações de Ambiente
Crie uma .venv: py -3.12 -m venv .venv
Ative:  .venv\Scripts\activate
Instale as bibliotecas: pip install pandas scikit-learn openpyxl matplotlib seaborn customtkinter
Rodar: python App.py  

Estrutura:

Projeto_Integrador5/
│
├── App.py                    # inicia a interface
│
├── interface/
│   ├── telas.py              # janelas e frames
│   ├── componentes.py        # botões, tabelas, etc
│   └── eventos.py            # ações da interface
│
├── ia/
│   ├── treino.py             # treinamento Random Forest
│   ├── previsao.py           # carregar modelo e prever
│   ├── preprocessamento.py   # limpeza e transformação
│   └── modelo.pkl            # modelo treinado
│
├── dados/
│   ├── dataset.csv
│   └── entrada_usuario.csv
│
├── utils/
│   ├── arquivos.py
│   ├── validacoes.py
│   └── graficos.py
│
├── requirements.txt
└── .venv/