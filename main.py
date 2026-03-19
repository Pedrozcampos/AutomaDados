import os
import sqlite3
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment


# -- Classe dos Procedimentos --
class Processamentos:
    # Limpeza de dados 
    def tratar_dados(self, df):
        for col in ['Débito', 'Crédito']:
            if col in df.columns:
                df[col] = df[col].astype(str)
                df[col] = df[col].replace(r'[R\$\s\xa0]', '', regex=True).str.replace(',', '.')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df

    # Verificação de colunas necessárias.
    def validar_col(self, df):
        colunas_obrigatorias = ['Histórico', 'Débito', 'Crédito', 'Conta', 'Data',]
        colunas_faltantes = [c for c in colunas_obrigatorias if c not in df.columns]
        if colunas_faltantes:
            msg = F"Erro! As colunas {', '.join(colunas_faltantes)} não foram encontradas."
            messagebox.showerror("Colunas Invalidas", msg)
            return False
        return True

    # Seleção de arquivo e leitura do arquivo.
    def proc_adicionar(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        termo = simpledialog.askstring("Busca", "Palavra-chave:")
        if not termo:
            return
        if path:
            if termo:
                try:
                    df = pd.read_excel(path)
                    if not self.validar_col(df):
                        return
                    df = self.tratar_dados(df)
                    nome_tabela = "base" if not self.arquivos else f"confronto_{len(self.arquivos)}"
                    with self.connection:
                        df.to_sql(nome_tabela, self.connection, if_exists='replace', index=False)
                    self.arquivos.append({'tabela': nome_tabela, 'termo': termo, 'path': os.path.basename(path)})
                    self.atualizar_fila()
                except Exception as e:
                    messagebox.showerror("Erro", f"Falha ao carregar arquivo: {e}")

    # Atualiza a fila no gui, mostrando os arquivos selecionados.
    def atualizar_fila(self):
        self.txt_fila.configure(state="normal")
        self.txt_fila.delete("1.0", "end")
        for i, arc in enumerate(self.arquivos):
            tag = "BASE" if i == 0 else f"CONF {i}"
            self.txt_fila.insert("end", f"[{tag}] {arc['path']} | Chave: {arc['termo']}\n")
        self.txt_fila.configure(state="disabled")


    # Limpa os arquivos selecionados e a mémoria.
    def proc_limpar(self):
        self.arquivos = []
        self.resultados_por_aba = {}
        
        self.connection.close()
        self.connection = sqlite3.connect(":memory:", check_same_thread=False)
        
        self.atualizar_fila()
        self.txt_resumo.configure(state="normal")
        self.txt_resumo.delete("1.0", "end")
        self.txt_resumo.configure(state="disabled")
        self.btn_export.configure(state="disabled")

    # Faz a busca com o SQL utilizando a plavra-chave ( o termo )
    def processar_sql(self):
        if len(self.arquivos) < 2:
            messagebox.showwarning("Aviso", "Adicione a Base e pelo menos um Confronto.")
            return
        try:
            self.txt_resumo.configure(state="normal")
            self.txt_resumo.delete("1.0", "end")
            self.txt_resumo.insert("end", "-"*50 + "\n")
            self.resultados_por_aba = {}

            # Função para buscar o termo em todas colunas.
            def busca_termo_col(tabela, termo):
                cursor = self.connection.execute(f"PRAGMA table_info('{tabela}')")
                colunas = [info[1] for info in cursor.fetchall()]
                filtro = " OR ".join([f"CAST([{col}] AS TEXT) LIKE ?" for col in colunas])
                query = f"SELECT * FROM {tabela} WHERE {filtro}"
                parametros = [F"%{termo}%"] * len(colunas)
                return pd.read_sql_query(query, self.connection, params=parametros, dtype=str)

            df_b = busca_termo_col("base", self.arquivos[0]['termo'])
            self.txt_resumo.insert("end", f"[BASE] {self.arquivos[0]['path']}: {len(df_b)} linhas encontradas.\n")
            for i in range(1, len(self.arquivos)):
                conf_info = self.arquivos[i]
                df_c = busca_termo_col(conf_info['tabela'], conf_info['termo'])
            
            # Garante conversão numérica para soma total
                for c in ['Débito', 'Crédito']:
                    df_b[c] = pd.to_numeric(df_b[c], errors='coerce').fillna(0)
                    df_c[c] = pd.to_numeric(df_c[c], errors='coerce').fillna(0)
                # Totais
                tot_deb_b = df_b['Débito'].sum()
                tot_cre_b = df_b['Crédito'].sum()
                tot_deb_c = df_c['Débito'].sum()
                tot_cre_c = df_c['Crédito'].sum()

                operacao = self.optionmenu_var.get()
                if operacao == "Adição":
                    res_deb = round(float(tot_deb_b + tot_deb_c), 2)
                    res_cre = round(float(tot_cre_b + tot_cre_c), 2)
                else:
                    res_deb = round(float(tot_deb_b - tot_deb_c), 2)
                    res_cre = round(float(tot_cre_b - tot_cre_c), 2)

                df_aba = pd.concat([
                    df_b.reset_index(drop=True),
                    pd.DataFrame({self.SEP1: [" "] * max(len(df_b), len(df_c))}),
                    df_c.reset_index(drop=True),
                    pd.DataFrame({self.SEP2: [" "] * max(len(df_b), len(df_c))})
                ], axis=1).fillna(0)
                
                df_aba['DÉBITO TOTAL'] = None
                df_aba['CRÉDITO TOTAL'] = None  
                df_aba.at[0, 'DÉBITO TOTAL'] = res_deb
                df_aba.at[0, 'CRÉDITO TOTAL'] = res_cre

                aba_nome = f"Confronto {i}"
                self.resultados_por_aba[aba_nome] = df_aba
                self.txt_resumo.insert("end", f"[CONF {i}] Processado com sucesso.\n")
                
                self.txt_resumo.insert("end", "-"*50 + "\n")
                self.txt_resumo.insert("end", "SUCESSO: Dados prontos para exportação.")
                self.txt_resumo.configure(state="disabled")
                self.btn_export.configure(state="normal")
                messagebox.showinfo("Sucesso", "Processamento concluído!")  

        except Exception as e:
            self.txt_resumo.insert("end", f"\n\nERRO CRÍTICO: {e}")
            self.txt_resumo.configure(state="disabled")
            messagebox.showerror("Erro", f"Erro no processamento: {e}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro no processamento (Verifique se a coluna 'Histórico' existe): {e}")

    # Exportação dos resultados para arquivo Excel já formatado.
    def exportar_excel(self):
        try:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Confronto_Resultado.xlsx")
            if not save_path: return
            
            operacao = self.optionmenu_var.get()
            
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                for nome_aba, df in self.resultados_por_aba.items():
                    cols = df.columns.tolist()
                    debs = [i for i, x in enumerate(cols) if x == 'Débito']
                    creds = [i for i, x in enumerate(cols) if x == 'Crédito']

                    # Infos do Resultado
                    if len(debs) >= 2 and len(creds) >= 2:
                        if operacao == "Adição":
                            df['DÉBITO'] = df.iloc[:, debs[0]] + df.iloc[:, debs[1]]
                            df['CRÉDITO'] = df.iloc[:, creds[0]] + df.iloc[:, creds[1]]
                        else:
                            df['DÉBITO'] = df.iloc[:, debs[0]] - df.iloc[:, debs[1]]
                            df['CRÉDITO'] = df.iloc[:, creds[0]] - df.iloc[:, creds[1]]

                    df.to_excel(writer, index=False, sheet_name=nome_aba, startrow=1)
                    ws = writer.sheets[nome_aba]

                    # Localiza os separadores para a mesclagem visual
                    try:
                        sep1_col = cols.index(self.SEP1) + 1 
                        sep2_col = cols.index(self.SEP2) + 1
                    except ValueError:
                        continue 

                    font_white = Font(bold=True, color="FFFFFF")
                    center = Alignment(horizontal="center", vertical="center")
                    fill_base = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                    fill_conf = PatternFill(start_color="E26B0A", end_color="E26B0A", fill_type="solid")
                    fill_res = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
                    
                    # Base
                    if sep1_col > 1:
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=sep1_col-1)
                        c1 = ws.cell(row=1, column=1, value="DADOS DA BASE")
                        c1.font, c1.fill, c1.alignment = font_white, fill_base, center

                    # Confronto
                    start_c = sep1_col + 1
                    end_c = sep2_col - 1
                    if end_c >= start_c: 
                        ws.merge_cells(start_row=1, start_column=start_c, end_row=1, end_column=end_c)
                        c2 = ws.cell(row=1, column=start_c, value="DADOS DO CONFRONTO")
                        c2.font, c2.fill, c2.alignment = font_white, fill_conf, center

                    # Resultado
                    start_r = sep2_col + 1
                    end_r = len(df.columns)
                    if end_r >= start_r:
                        ws.merge_cells(start_row=1, start_column=start_r, end_row=1, end_column=end_r)
                        c3 = ws.cell(row=1, column=start_r, value="RESULTADO CONSOLIDADO (TOTAIS)")
                        c3.font, c3.fill, c3.alignment = font_white, fill_res, center
                        
                        #
                        ws.cell(row=2, column=start_r).font = Font(bold=True)
                        ws.cell(row=2, column=start_r + 1).font = Font(bold=True)
                        ws.column_dimensions[ws.cell(row=2, column=start_r).column_letter].width = 20
                        ws.column_dimensions[ws.cell(row=2, column=start_r + 1).column_letter].width = 20
            messagebox.showinfo("Sucesso", "Exportado com Sucesso!")
            os.startfile(save_path)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na formatação: {e}")

# -- Classe do GUI --
class AutoDados(ctk.CTk, Processamentos):
    def __init__(self):
        super().__init__()
        self.title("Auditoria Pro - AutoDados")
        self.geometry("800x800")
        self.arquivos = [] 
        self.resultados_por_aba = {}
        self.connection = sqlite3.connect(":memory:", check_same_thread=False)
        self.optionmenu_var = ctk.StringVar(value="Adição")
        
        
        self.SEP1 = " "
        self.SEP2 = "  "
        
        self.criar_layout()

    def criar_layout(self):
        ctk.CTkLabel(self, text="Painel de Confronto e Diferenças", font=("Arial", 22, "bold")).pack(pady=15)

        frame_top = ctk.CTkFrame(self)
        frame_top.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkButton(frame_top, text="+ Adicionar Arquivo", fg_color="#1f538d", command=self.proc_adicionar).pack(side="left", padx=10, pady=10)
        ctk.CTkButton(frame_top, text="- Limpar Arquivo", fg_color="#A82424", command=self.proc_limpar).pack(side="left", padx=10, pady=10)

        ctk.CTkLabel(frame_top, text="Operação:").pack(side="left", padx=5)
        self.combobox = ctk.CTkOptionMenu(frame_top, values=["Adição", "Subtração"], variable=self.optionmenu_var)
        self.combobox.pack(side="left", padx=10, pady=10)

        # Fila de arquivos
        ctk.CTkLabel(self, text="Arquivos na Fila:").pack(anchor="w", padx=25)
        self.txt_fila = ctk.CTkTextbox(self, width=750, height=120)
        self.txt_fila.pack(pady=5, padx=20)
        self.txt_fila.configure(state="disabled")

        self.btn_run = ctk.CTkButton(self, text="PROCESSAR", fg_color="green", height=45, command=self.processar_sql)
        self.btn_run.pack(pady=15)

        # Log de Resumo
        self.txt_resumo = ctk.CTkTextbox(self, width=750, height=250, font=("Courier New", 12))
        self.txt_resumo.pack(pady=10, padx=20)

        self.btn_export = ctk.CTkButton(self, text="EXPORTAR RELATÓRIO (EXCEL)", state="disabled", command=self.exportar_excel)
        self.btn_export.pack(pady=10)


if __name__ == "__main__":
    app = AutoDados()
    app.mainloop()