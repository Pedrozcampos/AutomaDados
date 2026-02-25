import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
import sqlite3
import os
from threading import Thread

# --- LÓGICA DE PROCESSAMENTO (BACK-END) ---
class DataManager:
    # -- Classe responsável por toda a manipulação de dados e SQL. --
    
    def __init__(self):
        self.conn = None

    def processar_abate(self, caminhos_origem, caminho_target, coluna_chave, coluna_valor, operacao):
        try:
            # 1. Conexão em memória
            self.conn = sqlite3.connect(':memory:')
            
            # 2. Ingestão dos arquivos de origem
            dfs = []
            for f in caminhos_origem:
                df = pd.read_excel(f) if f.endswith(('.xlsx', '.xls')) else pd.read_csv(f)
                dfs.append(df)
            
            df_origem_total = pd.concat(dfs).reset_index(drop=True)
            df_origem_total.to_sql('origem', self.conn, index=False)

            # 3. Ingestão do arquivo de comparação (Target)
            df_target = pd.read_excel(caminho_target) if caminho_target.endswith(('.xlsx', '.xls')) else pd.read_csv(caminho_target)
            df_target.to_sql('target', self.conn, index=False)

            # 4. Construção da Query SQL baseada na operação
            # Usamos COALESCE para tratar valores nulos (vazios) como 0
            sinal = "-" if "Subtração" in operacao else "+"
            
            query = f"""
            SELECT 
                o.*, 
                t.{coluna_valor} AS valor_comparacao,
                (o.{coluna_valor} {sinal} COALESCE(t.{coluna_valor}, 0)) AS resultado_final
            FROM origem o
            LEFT JOIN target t ON o.{coluna_chave} = t.{coluna_chave}
            """
            
            resultado = pd.read_sql_query(query, self.conn)
            
            # 5. Exportação
            output = "Resultado_Processamento.xlsx"
            resultado.to_excel(output, index=False)
            return True, output

        except Exception as e:
            return False, str(e)
        finally:
            if self.conn:
                self.conn.close()

# --- INTERFACE GRÁFICA (FRONT-END) ---
class AppGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Instância da lógica
        self.manager = DataManager()
        
        # Configurações da Janela
        self.title("Data Engine SQL Pro")
        self.geometry("700x600")
        self.configure_gui()
        self.arquivos_selecionados = []
        self.arquivo_target = ""

    def configure_gui(self):
        self.grid_columnconfigure(0, weight=1)

        # Seção de Arquivos
        self.frame_files = ctk.CTkFrame(self)
        self.frame_files.pack(pady=20, padx=20, fill="x")

        ctk.CTkLabel(self.frame_files, text="Painel de Controle de Dados", font=("Arial", 18, "bold")).pack(pady=10)
        
        self.btn_files = ctk.CTkButton(self.frame_files, text="Selecionar Origens (Excel/CSV)", command=self.load_origins)
        self.btn_files.pack(pady=5)
        
        self.lbl_origem = ctk.CTkLabel(self.frame_files, text="Nenhum arquivo carregado", text_color="gray")
        self.lbl_origem.pack()

        self.btn_target = ctk.CTkButton(self.frame_files, text="Selecionar Base de Comparação", fg_color="#2c3e50", command=self.load_target)
        self.btn_target.pack(pady=10)

        # Seção de Configuração de Colunas
        self.frame_config = ctk.CTkFrame(self)
        self.frame_config.pack(pady=10, padx=20, fill="x")

        self.entry_key = ctk.CTkEntry(self.frame_config, placeholder_text="Coluna Identificadora (ex: SKU, CPF)")
        self.entry_key.pack(side="left", padx=10, pady=10, expand=True, fill="x")

        self.entry_val = ctk.CTkEntry(self.frame_config, placeholder_text="Coluna de Valor (ex: Preço, Qtd)")
        self.entry_val.pack(side="left", padx=10, pady=10, expand=True, fill="x")

        # Operação e Botão Principal
        self.op_menu = ctk.CTkOptionMenu(self, values=["Subtração", "Soma "])
        self.op_menu.pack(pady=10)

        self.btn_run = ctk.CTkButton(self, text="PROCESSAR VIA SQL ENGINE", height=45, fg_color="#27ae60", hover_color="#2ecc71", command=self.run_process)
        self.btn_run.pack(pady=20)

        # Console de Logs
        self.txt_log = ctk.CTkTextbox(self, height=150, width=600)
        self.txt_log.pack(pady=10, padx=20)
        self.log("Sistema pronto. Aguardando arquivos...")

    # Funções de Suporte (Handlers)
    def log(self, mensagem):
        self.txt_log.insert("end", f"> {mensagem}\n")
        self.txt_log.see("end")

    def load_origins(self):
        self.arquivos_selecionados = filedialog.askopenfilenames()
        if self.arquivos_selecionados:
            self.lbl_origem.configure(text=f"{len(self.arquivos_selecionados)} arquivos selecionados", text_color="white")
            self.log(f"{len(self.arquivos_selecionados)} arquivos de origem carregados.")

    def load_target(self):
        self.arquivo_target = filedialog.askopenfilename()
        if self.arquivo_target:
            self.log(f"Base de comparação definida: {os.path.basename(self.arquivo_target)}")

    def run_process(self):
        # Validação simples
        if not self.arquivos_selecionados or not self.arquivo_target:
            messagebox.showwarning("Erro", "Selecione todos os arquivos antes de continuar.")
            return
        
        # Thread para não travar a interface
        self.log("Iniciando motor SQL... Aguarde.")
        Thread(target=self.execute_task).start()

    def execute_task(self):
        chave = self.entry_key.get()
        valor = self.entry_val.get()
        op = self.op_menu.get()

        sucesso, msg = self.manager.processar_abate(
            self.arquivos_selecionados, 
            self.arquivo_target, 
            chave, 
            valor, 
            op
        )

        if sucesso:
            self.log(f"Sucesso! Resultado gerado em: {msg}")
            messagebox.showinfo("Concluído", "Arquivos processados e salvos com sucesso!")
            os.startfile(msg)
        else:
            self.log(f"ERRO: {msg}")
            messagebox.showerror("Erro no Processamento", msg)

if __name__ == "__main__":
    app = AppGUI()
    app.mainloop()