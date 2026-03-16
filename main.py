import os
import sqlite3
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd

class AutoDados(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auditoria Pro - Confronto Multi-Arquivos")
        self.geometry("800x800")
        self.arquivos = [] 

        self.connection = sqlite3.connect("AutoDados.db")
        self.cursor = self.connection.cursor()

        # - Variável de controle da operação (Deve estar aqui para evitar o erro de 'root window')
        self.optionmenu_var = ctk.StringVar(value="Adição")

        self._criar_layout()

    def _criar_layout(self):
        ctk.CTkLabel(self, text="Painel de Confronto e Diferenças", font=("Arial", 22, "bold")).pack(pady=15)

        # Controles
        frame_top = ctk.CTkFrame(self)
        frame_top.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkButton(frame_top, text="+ Adicionar Arquivo", fg_color="#1f538d", command=self.procadicionar).pack(side="left", padx=10, pady=10)
        
        # Seletor de Operação
        ctk.CTkLabel(frame_top, text="Operação:").pack(side="left", padx=5)
        self.combobox = ctk.CTkOptionMenu(
            frame_top, 
            values=["Adição", "Subtração"],
            variable=self.optionmenu_var
        )
        self.combobox.pack(side="left", padx=10, pady=10)

        ctk.CTkButton(frame_top, text="- Limpar", fg_color="#A82424", command=self.proclimpar).pack(side="left", padx=10, pady=10)

        # Fila de Arquivos
        ctk.CTkLabel(self, text="Arquivos na Fila:").pack(anchor="w", padx=25)
        self.txt_fila = ctk.CTkTextbox(self, width=750, height=120)
        self.txt_fila.pack(pady=5, padx=20)
        self.txt_fila.configure(state="disabled")

        # Botão Executar
        self.btn_run = ctk.CTkButton(self, text="CALCULAR DIFERENÇAS", fg_color="green", height=45, command=self.processar_sql)
        self.btn_run.pack(pady=15)

        # Painel de Resultados 
        self.txt_resumo = ctk.CTkTextbox(self, width=750, height=250, font=("Courier New", 12))
        self.txt_resumo.pack(pady=10, padx=20)

        self.btn_export = ctk.CTkButton(self, text="EXPORTAR RELATÓRIO COMPLETO", state="disabled", command=self.exportar_excel)
        self.btn_export.pack(pady=10)

    def tratar_dados(self, df):
        for col in ['Débito', 'Crédito']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace('R$', '', regex=False).str.strip()
                def formatar(v):
                    if ',' in v and '.' in v: v = v.replace('.', '').replace(',', '.')
                    elif ',' in v: v = v.replace(',', '.')
                    return v
                df[col] = df[col].apply(formatar)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df

    def procadicionar(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            termo = simpledialog.askstring("Busca", "Palavra-chave:")
            if termo:
                df = pd.read_excel(path)
                df = self.tratar_dados(df)
                
                nome_tabela = "base" if not self.arquivos else f"confronto_{len(self.arquivos)}"
                df.to_sql(nome_tabela, self.connection, if_exists='replace', index=False)
                
                self.arquivos.append({'tabela': nome_tabela, 'termo': termo, 'path': os.path.basename(path)})
                self.atualizar_fila()

    def atualizar_fila(self):
        self.txt_fila.configure(state="normal")
        self.txt_fila.delete("1.0", "end")
        for i, arc in enumerate(self.arquivos):
            tag = "BASE" if i == 0 else f"CONF {i}"
            self.txt_fila.insert("end", f"[{tag}] {arc['path']} | Chave: {arc['termo']}\n")
        self.txt_fila.configure(state="disabled")

    def proclimpar(self):
        self.arquivos = []
        self.atualizar_fila()
        self.txt_resumo.delete("1.0", "end")
        self.btn_export.configure(state="disabled")

    def processar_sql(self):
        """ Executa a consulta e mostra um resumo no campo de texto """
        try:
            if len(self.arquivos) < 2:
                messagebox.showwarning("Aviso", "Adicione a Base e pelo menos um Confronto.")
                return

            t_base = f"%{self.arquivos[0]['termo']}%"
            t_conf = f"%{self.arquivos[1]['termo']}%"

            query = "SELECT b.*, ' ' as [S], c.* FROM base b, confronto_1 c WHERE b.Histórico LIKE ? AND c.Histórico LIKE ?"
            df = pd.read_sql_query(query, self.connection, params=(t_base, t_conf))

            if df.empty:
                self.txt_resumo.insert("end", "\nNenhum dado encontrado.")
                return

            self.txt_resumo.delete("1.0", "end")
            self.txt_resumo.insert("end", f"Sucesso! {len(df)} linhas encontradas.\n")
            self.txt_resumo.insert("end", "Clique em 'Exportar' para gerar o arquivo detalhado.")
            self.btn_export.configure(state="normal")
            self.temp_df = df # Guarda o resultado temporariamente para exportar

        except Exception as e:
            messagebox.showerror("Erro SQL", str(e))

    def exportar_excel(self):
        """ Salva o resultado processado no Excel com cálculos """
        try:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
            if save_path:
                df = self.temp_df.copy()
                
                # Identifica colunas duplicadas para cálculo
                cols = df.columns.tolist()
                debs = [i for i, x in enumerate(cols) if x == 'Débito']
                creds = [i for i, x in enumerate(cols) if x == 'Crédito']

                df['DIF_DEBITO'] = df.iloc[:, debs[0]] - df.iloc[:, debs[1]]
                df['DIF_CREDITO'] = df.iloc[:, creds[0]] - df.iloc[:, creds[1]]
                
                df.to_excel(save_path, index=False)
                messagebox.showinfo("Sucesso", "Excel gerado!")
                os.startfile(save_path)
        except Exception as e:
            messagebox.showerror("Erro Exportação", str(e))

if __name__ == "__main__":
    app = AutoDados()
    app.mainloop()