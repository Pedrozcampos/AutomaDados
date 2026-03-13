import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
from datetime import datetime
import os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class AuditoriaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auditoria Pro - Confronto Multi-Arquivos")
        self.geometry("900x750")
        self.arquivos = [] # Lista de dicts: {'path', 'termo', 'deb', 'cred', 'df'}

        self._criar_layout()

    def _criar_layout(self):
        ctk.CTkLabel(self, text="Painel de Confronto e Diferenças", font=("Arial", 22, "bold")).pack(pady=15)

        # Controles
        frame_top = ctk.CTkFrame(self)
        frame_top.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkButton(frame_top, text="+ Adicionar Arquivo", fg_color="#1f538d", command=self._adicionar).pack(side="left", padx=10, pady=10)
        ctk.CTkButton(frame_top, text="🗑️ Limpar", fg_color="#A82424", command=self._limpar).pack(side="left", padx=10, pady=10)

        # Fila de Arquivos
        self.txt_fila = ctk.CTkTextbox(self, width=800, height=120)
        self.txt_fila.pack(pady=5, padx=20)
        self.txt_fila.configure(state="disabled")

        # Botão Executar
        self.btn_run = ctk.CTkButton(self, text="CALCULAR DIFERENÇAS", fg_color="green", height=45, command=self._processar)
        self.btn_run.pack(pady=15)

        # Painel de Resultados Detalhado
        self.txt_resumo = ctk.CTkTextbox(self, width=800, height=300, font=("Courier New", 12))
        self.txt_resumo.pack(pady=10, padx=20)

        self.btn_export = ctk.CTkButton(self, text="EXPORTAR RELATÓRIO COMPLETO", state="disabled", command=self._exportar)
        self.btn_export.pack(pady=10)

    optionmenu_var = ctk.StringVar(value="Adição")
    def optionmenu_callback(self, choice):
            print("optionmenu dropdown clicked:", choice)

    optionmenu_var = ctk.StringVar(value="Adição")
    combobox = ctk.CTkOptionMenu(   values=["Adição", "Subtração", "Multiplicação", "Divisão"],
                                command=optionmenu_callback,
                                variable=optionmenu_var)
    combobox.pack(padx=20, pady=10)



    def _adicionar(self):
        path = filedialog.askopenfilename(filetypes=[("Dados", "*.xlsx *.csv")])
        if path:
            sufixo = "BASE (Ref.)" if not self.arquivos else f"Arq {len(self.arquivos)+1}"
            termo = simpledialog.askstring("Busca", f"Palavra-chave para {sufixo}:")
            if termo:
                self.arquivos.append({'path': path, 'termo': termo})
                self._atualizar_fila()

    def _limpar(self):
        self.arquivos = []
        self._atualizar_fila()
        self.txt_resumo.delete("1.0", "end")
        self.btn_export.configure(state="disabled")

    def _atualizar_fila(self):
        self.txt_fila.configure(state="normal")
        self.txt_fila.delete("1.0", "end")
        for i, arc in enumerate(self.arquivos):
            prefixo = "BASE ->" if i == 0 else f"CONFRONTO {i} ->"
            self.txt_fila.insert("end", f"{prefixo} [{arc['termo'].upper()}] - {os.path.basename(arc['path'])}\n")
        self.txt_fila.configure(state="disabled")

    def _limpar_valor(self, df):
        for col in ['Débito', 'Crédito']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                def tratar(v):
                    if ',' in v: v = v.replace('.', '').replace(',', '.')
                    return v
                df[col] = df[col].apply(tratar)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df

    def _processar(self):
        if len(self.arquivos) < 2:
            messagebox.showwarning("Aviso", "Adicione pelo menos a BASE e um arquivo de CONFRONTO.")
            return

        try:
            # 1. Processa a BASE (Primeiro arquivo)
            base = self.arquivos[0]
            df_b = pd.read_csv(base['path']) if base['path'].endswith('.csv') else pd.read_excel(base['path'])
            df_b = self._limpar_valor(df_b)
            mask_b = (df_b['Conta'].astype(str).str.contains(base['termo'], case=False, na=False)) | \
                    (df_b['Histórico'].astype(str).str.contains(base['termo'], case=False, na=False))
            base['df'] = df_b[mask_b].copy()
            base['deb'], base['cred'] = base['df']['Débito'].sum(), base['df']['Crédito'].sum()

            saida = f"RELATÓRIO DE DIFERENÇAS (Base: {base['termo']})\n" + "="*75 + "\n"
            saida += f"VALORES BASE: Débito R$ {base['deb']:,.2f} | Crédito R$ {base['cred']:,.2f}\n\n"

            # 2. Compara os demais com a BASE
            for i in range(1, len(self.arquivos)):
                arc = self.arquivos[i]
                df = pd.read_csv(arc['path']) if arc['path'].endswith('.csv') else pd.read_excel(arc['path'])
                df = self._limpar_valor(df)
                mask = (df['Conta'].astype(str).str.contains(arc['termo'], case=False, na=False)) | \
                    (df['Histórico'].astype(str).str.contains(arc['termo'], case=False, na=False))
                arc['df'] = df[mask].copy()
                arc['deb'], arc['cred'] = arc['df']['Débito'].sum(), arc['df']['Crédito'].sum()

                # Cálculo de Diferenças
                dif_d = base['deb'] - arc['deb']
                dif_c = base['cred'] - arc['cred']

                saida += f"CONFRONTO {i}: {os.path.basename(arc['path'])} ({arc['termo']})\n"
                saida += f"  > Débito : R$ {arc['deb']:,.2f} (Dif: R$ {dif_d:,.2f})\n"
                saida += f"  > Crédito: R$ {arc['cred']:,.2f} (Dif: R$ {dif_c:,.2f})\n"
                saida += f"  > Saldo Liq. Dif: R$ {( (base['deb']-base['cred']) - (arc['deb']-arc['cred']) ):,.2f}\n"
                saida += "-"*75 + "\n"

            self.txt_resumo.delete("1.0", "end")
            self.txt_resumo.insert("1.0", saida)
            self.btn_export.configure(state="normal")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha no cálculo: {e}")

    def _aplicar_estilo(self, writer, sheet_name, df_f, deb, cred):
        ws = writer.sheets[sheet_name]
        
        # 1. Cabeçalho Cinza e Negrito
        fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        font_bold = Font(bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                            top=Side(style='thin'), bottom=Side(style='thin'))

        for cell in ws[1]:
            cell.fill = fill
            cell.font = font_bold
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

        # 2. Formatar Coluna de Data (procura se existe)
        for row in range(2, ws.max_row + 1):
            for cell in ws[row]:
                if isinstance(cell.value, (datetime, pd.Timestamp)):
                    cell.number_format = 'DD/MM/YYYY'
                elif "Data" in df_f.columns:
                    # Se for string mas estiver na coluna Data
                    cell.alignment = Alignment(horizontal="center")

        # 3. Criar Coluna de Relatório (Exemplo da imagem)
        col_resumo = df_f.shape[1] + 2 # Pula uma coluna
        ws.cell(row=1, column=col_resumo, value="RESUMO DA ANÁLISE").font = font_bold
        ws.merge_cells(start_row=1, start_column=col_resumo, end_row=1, end_column=col_resumo+1)

        labels = [("Total Débito:", deb), ("Total Crédito:", cred), ("Saldo Líquido:", deb-cred)]
        for i, (lab, val) in enumerate(labels, start=2):
            ws.cell(row=i, column=col_resumo, value=lab).font = font_bold
            ws.cell(row=i, column=col_resumo+1, value=val).number_format = '#,##0.00'
            
        # Auto-ajuste de colunas
        for col_index, col_name in enumerate(df_f.columns, 1):
            column_letter = get_column_letter(col_index)
    
            # Define uma largura padrão ou lógica de cálculo aqui
            ws.column_dimensions[column_letter].width = 20

        ws.column_dimensions[get_column_letter(col_resumo)].width = 20
        ws.column_dimensions[get_column_letter(col_resumo + 1)].width = 15
    def _exportar(self):
        try:
            fn = f"Relatorio_Formatado_{datetime.now().strftime('%H%M%S')}.xlsx"
            with pd.ExcelWriter(fn, engine='openpyxl') as writer:
                for i, arc in enumerate(self.arquivos):
                    nome_aba = "BASE_REF" if i == 0 else f"CONFRONTO_{i}"
                    arc['df'].to_excel(writer, sheet_name=nome_aba, index=False)
                    self._aplicar_estilo(writer, nome_aba, arc['df'], arc['deb'], arc['cred'])
            
            messagebox.showinfo("Sucesso", f"Excel formatado gerado!\n{fn}")
            os.startfile(fn)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    app = AuditoriaApp()
    app.mainloop()