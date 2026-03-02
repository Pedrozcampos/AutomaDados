import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os

class AuditoriaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auditoria Pro - Busca Independente")
        self.geometry("850x700")

        self.caminho_arquivo1 = ctk.StringVar()
        self.caminho_arquivo2 = ctk.StringVar()
        self.df_final1 = pd.DataFrame()
        self.df_final2 = pd.DataFrame()

        self._criar_layout()

    def _criar_layout(self):
        ctk.CTkLabel(self, text="Painel de Auditoria Independente", font=("Arial", 22, "bold")).pack(pady=15)

        # Seleção de Arquivos
        frame_files = ctk.CTkFrame(self)
        frame_files.pack(pady=10, padx=20, fill="x")

        # Arquivo A + Busca A
        ctk.CTkButton(frame_files, text="Arquivo A", command=lambda: self._carregar(1)).grid(row=0, column=0, padx=10, pady=10)
        ctk.CTkLabel(frame_files, textvariable=self.caminho_arquivo1, font=("Arial", 10)).grid(row=0, column=1, sticky="w")
        
        ctk.CTkLabel(self, text="Busca no Arquivo A:").pack(pady=(5,0))
        self.entry_busca_a = ctk.CTkEntry(self, width=400, placeholder_text="Termo para Arquivo A (ex: Itaú)")
        self.entry_busca_a.pack(pady=5)

        # Arquivo B + Busca B
        frame_files_b = ctk.CTkFrame(self)
        frame_files_b.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkButton(frame_files_b, text="Arquivo B", command=lambda: self._carregar(2)).grid(row=0, column=0, padx=10, pady=10)
        ctk.CTkLabel(frame_files_b, textvariable=self.caminho_arquivo2, font=("Arial", 10)).grid(row=0, column=1, sticky="w")
        
        ctk.CTkLabel(self, text="Busca no Arquivo B:").pack(pady=(5,0))
        self.entry_busca_b = ctk.CTkEntry(self, width=400, placeholder_text="Termo para Arquivo B (ex: 1.1.01.001)")
        self.entry_busca_b.pack(pady=5)

        # Botão Processar
        self.btn_run = ctk.CTkButton(self, text="EXECUTAR ANÁLISE COMPARATIVA", fg_color="#1f538d", height=45, command=self._processar)
        self.btn_run.pack(pady=20)

        # Painel de Resultados
        self.result_frame = ctk.CTkTextbox(self, width=750, height=200, font=("Courier New", 13))
        self.result_frame.pack(pady=10, padx=20)

        # Botão Exportar
        self.btn_export = ctk.CTkButton(self, text="EXPORTAR RELATÓRIO (EXCEL)", state="disabled", fg_color="green", command=self._exportar)
        self.btn_export.pack(pady=10)

    def _carregar(self, num):
        path = filedialog.askopenfilename(filetypes=[("Dados", "*.xlsx *.csv")])
        if path:
            if num == 1: self.caminho_arquivo1.set(path)
            else: self.caminho_arquivo2.set(path)

    def _limpar_valor(self, df):
        """Trata formatos decimais com segurança."""
        for col in ['Débito', 'Crédito']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                def tratar(val):
                    if ',' in val:
                        val = val.replace('.', '').replace(',', '.')
                    return val
                df[col] = df[col].apply(tratar)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df

    def _processar(self):
        termo_a = self.entry_busca_a.get()
        termo_b = self.entry_busca_b.get()
        f1, f2 = self.caminho_arquivo1.get(), self.caminho_arquivo2.get()

        if not termo_a or not termo_b or not f1 or not f2:
            messagebox.showwarning("Aviso", "Preencha todos os campos de busca e selecione os arquivos.")
            return

        try:
            df1 = pd.read_csv(f1) if f1.endswith('.csv') else pd.read_excel(f1)
            df2 = pd.read_csv(f2) if f2.endswith('.csv') else pd.read_excel(f2)

            df1 = self._limpar_valor(df1)
            df2 = self._limpar_valor(df2)

            # Filtro independente para cada arquivo (Busca em Conta e Histórico)
            mask1 = (df1['Conta'].astype(str).str.contains(termo_a, case=False, na=False)) | \
                    (df1['Histórico'].astype(str).str.contains(termo_a, case=False, na=False))
            
            mask2 = (df2['Conta'].astype(str).str.contains(termo_b, case=False, na=False)) | \
                    (df2['Histórico'].astype(str).str.contains(termo_b, case=False, na=False))

            self.df_final1 = df1[mask1].copy()
            self.df_final2 = df2[mask2].copy()

            # Cálculos
            deb_a, cred_a = self.df_final1['Débito'].sum(), self.df_final1['Crédito'].sum()
            deb_b, cred_b = self.df_final2['Débito'].sum(), self.df_final2['Crédito'].sum()

            resumo = (
                f"AUDITORIA CONCLUÍDA\n"
                f"Filtro A: [{termo_a}] | Filtro B: [{termo_b}]\n"
                f"{'='*60}\n"
                f"ARQUIVO A: Débito R$ {deb_a:,.2f} | Crédito R$ {cred_a:,.2f}\n"
                f"ARQUIVO B: Débito R$ {deb_b:,.2f} | Crédito R$ {cred_b:,.2f}\n"
                f"{'-'*60}\n"
                f"DIFERENÇA DÉBITO : R$ {abs(deb_a - deb_b):,.2f}\n"
                f"DIFERENÇA CRÉDITO: R$ {abs(cred_a - cred_b):,.2f}\n"
                f"{'='*60}\n"
            )

            self.result_frame.delete("1.0", "end")
            self.result_frame.insert("1.0", resumo)
            self.btn_export.configure(state="normal")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro no processamento: {e}")

    def _exportar(self):
        try:
            filename = f"Relatorio_Comparativo_{datetime.now().strftime('%H%M%S')}.xlsx"
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                self.df_final1.to_excel(writer, sheet_name='Resultados_A', index=False)
                self.df_final2.to_excel(writer, sheet_name='Resultados_B', index=False)
            messagebox.showinfo("Sucesso", f"Planilha gerada!\nArquivo: {filename}")
            os.startfile(filename)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    app = AuditoriaApp()
    app.mainloop()