import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import threading
import os
import re

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

class BuscadorPropostas(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Buscador de Propostas v1.0")
        self.geometry("550x450")
        self.resizable(False, False)
        
        self.arq_alvo = ""
        self.arq_base = ""
        
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.pack(pady=15, padx=15, fill="both", expand=True)

        self.lbl_title = ctk.CTkLabel(self.main_frame, text="🔍 Cruzador de Propostas", font=("Roboto", 20, "bold"))
        self.lbl_title.pack(pady=(15, 20))
        
        # Botões
        self.btn_alvo = ctk.CTkButton(self.main_frame, text="📂 1. Planilha que PRECISA das propostas", command=self.selecionar_alvo, height=40, width=350)
        self.btn_alvo.pack(pady=(5, 2))
        self.lbl_alvo = ctk.CTkLabel(self.main_frame, text="Nenhum arquivo", text_color="gray", font=("Roboto", 11))
        self.lbl_alvo.pack(pady=(0, 15))
        
        self.btn_base = ctk.CTkButton(self.main_frame, text="📂 2. Planilha BASE (que contém as propostas)", command=self.selecionar_base, height=40, width=350)
        self.btn_base.pack(pady=(5, 2))
        self.lbl_base = ctk.CTkLabel(self.main_frame, text="Nenhum arquivo", text_color="gray", font=("Roboto", 11))
        self.lbl_base.pack(pady=(0, 20))

        # Progresso
        self.progress = ctk.CTkProgressBar(self.main_frame, width=400, height=10)
        self.progress.pack(pady=10)
        self.progress.set(0)
        
        # Executar
        self.btn_executar = ctk.CTkButton(self.main_frame, text="⚡ Processar e Criar Nova Planilha", command=self.iniciar_processamento, 
                                          fg_color="#006400", hover_color="#004d00", font=("Roboto", 14, "bold"), height=45)
        self.btn_executar.pack(pady=10, padx=30, fill="x")

        self.lbl_footer = ctk.CTkLabel(self, text="Criado por Matheus Carvalho", font=("Roboto", 10, "italic"), text_color="gray")
        self.lbl_footer.pack(side="bottom", pady=5)

    def selecionar_alvo(self):
        self.arq_alvo = filedialog.askopenfilename(filetypes=[("Arquivos Excel/CSV/TXT", "*.xlsx *.xls *.csv *.txt *.TXT")])
        if self.arq_alvo:
            self.lbl_alvo.configure(text=os.path.basename(self.arq_alvo), text_color="white")

    def selecionar_base(self):
        self.arq_base = filedialog.askopenfilename(filetypes=[("Arquivos Excel/CSV/TXT", "*.xlsx *.xls *.csv *.txt *.TXT")])
        if self.arq_base:
            self.lbl_base.configure(text=os.path.basename(self.arq_base), text_color="white")

    def ler_arquivo(self, caminho):
        if caminho.lower().endswith(('.xlsx', '.xls')):
            return pd.read_excel(caminho, dtype=str).fillna("")
        else:
            try:
                return pd.read_csv(caminho, sep=';', dtype=str, encoding='utf-8-sig').fillna("")
            except UnicodeDecodeError:
                return pd.read_csv(caminho, sep=';', dtype=str, encoding='latin-1').fillna("")

    def limpar_cpf(self, cpf):
        if pd.isna(cpf): return ""
        # Remove tudo que não for número
        return re.sub(r'\D', '', str(cpf))

    def identificar_coluna(self, df, opcoes):
        for col in df.columns:
            if str(col).strip().upper() in opcoes:
                return col
        return None

    def iniciar_processamento(self):
        if not self.arq_alvo or not self.arq_base:
            messagebox.showwarning("Aviso", "Selecione as duas planilhas primeiro.")
            return
        
        self.btn_executar.configure(state="disabled", text="Processando...")
        self.progress.set(0.2)
        threading.Thread(target=self.processar_dados).start()

    def processar_dados(self):
        try:
            df_alvo = self.ler_arquivo(self.arq_alvo)
            self.progress.set(0.4)
            df_base = self.ler_arquivo(self.arq_base)
            self.progress.set(0.6)

            # Identificar colunas chaves na Base
            col_proposta = self.identificar_coluna(df_base, ['PROPOSTA', 'Nº PROPOSTA', 'NUMERO DA PROPOSTA', 'APOLICE'])
            col_cpf_base = self.identificar_coluna(df_base, ['CPF', 'CPF/CNPJ', 'CPF_CNPJ', 'DOCUMENTO'])
            col_nome_base = self.identificar_coluna(df_base, ['NOME DO SEGURADO', 'SEGURADO', 'NOME', 'CLIENTE'])

            if not col_proposta:
                messagebox.showerror("Erro", "Não encontrei a coluna 'Proposta' na planilha base.")
                return

            # Mapeamento: Cria dicionários rápidos para busca
            mapa_cpf = {}
            if col_cpf_base:
                for _, row in df_base.iterrows():
                    cpf_limpo = self.limpar_cpf(row[col_cpf_base])
                    if cpf_limpo: mapa_cpf[cpf_limpo] = str(row[col_proposta]).strip()

            mapa_nome = {}
            if col_nome_base:
                for _, row in df_base.iterrows():
                    nome_limpo = str(row[col_nome_base]).strip().upper()
                    if nome_limpo: mapa_nome[nome_limpo] = str(row[col_proposta]).strip()

            # Identificar colunas chaves no Alvo
            col_cpf_alvo = self.identificar_coluna(df_alvo, ['CPF', 'CPF/CNPJ', 'CPF_CNPJ', 'DOCUMENTO'])
            col_nome_alvo = self.identificar_coluna(df_alvo, ['NOME DO SEGURADO', 'SEGURADO', 'NOME', 'CLIENTE'])

            propostas_encontradas = []
            total_linhas = len(df_alvo)

            for idx, row in df_alvo.iterrows():
                # Atualiza barra visualmente
                if idx % 100 == 0:
                    self.progress.set(0.6 + (0.3 * (idx / total_linhas)))
                    self.update_idletasks()

                proposta = ""
                # Tenta por CPF primeiro
                if col_cpf_alvo:
                    cpf_alvo = self.limpar_cpf(row[col_cpf_alvo])
                    proposta = mapa_cpf.get(cpf_alvo, "")
                
                # Se não achou por CPF, tenta por Nome
                if not proposta and col_nome_alvo:
                    nome_alvo = str(row[col_nome_alvo]).strip().upper()
                    proposta = mapa_nome.get(nome_alvo, "NÃO LOCALIZADA")
                elif not proposta:
                    proposta = "NÃO LOCALIZADA"

                propostas_encontradas.append(proposta)

            # Insere a coluna na posição 0 (início)
            df_alvo.insert(0, 'Nº PROPOSTA', propostas_encontradas)

            self.progress.set(1.0)
            self.salvar_planilha(df_alvo)

        except Exception as e:
            messagebox.showerror("Erro de Processamento", f"Ocorreu um erro:\n{str(e)}")
        finally:
            self.btn_executar.configure(state="normal", text="⚡ Processar e Criar Nova Planilha")
            self.progress.set(0)

    def salvar_planilha(self, df):
        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Planilha Excel", "*.xlsx"), ("Arquivo CSV", "*.csv")], 
            initialfile="Planilha_Atualizada_com_Propostas.xlsx"
        )
        if not caminho: return

        try:
            if caminho.endswith('.csv'):
                df.to_csv(caminho, index=False, sep=';', encoding='utf-8-sig')
            else:
                df.to_excel(caminho, index=False)
            messagebox.showinfo("Sucesso!", "Planilha gerada com sucesso! As propostas foram inseridas na primeira coluna.")
        except Exception as e:
            messagebox.showerror("Erro ao Salvar", f"Falha ao salvar o arquivo:\n{str(e)}")

if __name__ == "__main__":
    app = BuscadorPropostas()
    app.mainloop()
