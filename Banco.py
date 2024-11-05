import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# Leitura das planilhas, desconsiderando as linhas de cabeçalho da planilha AX
def ler_planilha(caminho, skiprows=0, encoding='utf-8'):
    if caminho.endswith('.xlsx') or caminho.endswith('.xls'):
        return pd.read_excel(caminho, skiprows=skiprows)
    elif caminho.endswith('.csv'):
        try:
            return pd.read_csv(caminho, encoding=encoding, skiprows=skiprows, on_bad_lines='warn', delimiter=';')
        except UnicodeDecodeError:
            return pd.read_csv(caminho, encoding='iso-8859-1', skiprows=skiprows, on_bad_lines='warn', delimiter=';')
    else:
        raise ValueError("Formato de arquivo não suportado.")

# Consolida as planilhas do arquivo de movimentação
def consolidar_planilhas_movimento(caminho_movimento):
    excel_data = pd.read_excel(caminho_movimento, sheet_name=None, skiprows=3)
    todas_planilhas = []

    for nome_planilha, df in excel_data.items():
        if 'Nosso Número' in df.columns:
            indice_total = df[df['Nosso Número'].astype(str).str.contains('TOTAL', case=False, na=False)].index
            if not indice_total.empty:
                df = df.loc[:indice_total[0] - 1]
        todas_planilhas.append(df)

    consolidado_df = pd.concat(todas_planilhas, ignore_index=True)
    return consolidado_df

# Comparação das colunas 'Nosso Número' do consolidado e 'Fatura' da planilha AX
def comparar_consolidado_ax(consolidado_df, caminho_ax):
    ax_df = ler_planilha(caminho_ax, skiprows=8)

    # Convertendo as colunas para o tipo adequado
    consolidado_df['Nosso Número'] = consolidado_df['Nosso Número'].astype(float)
    ax_df['Fatura'] = ax_df['Fatura'].astype(float)

    # Encontrando as faturas que estão na AX e não na consolidação
    resultado_comparacao = ax_df[~ax_df['Fatura'].isin(consolidado_df['Nosso Número'])]

    return resultado_comparacao

class ApplicationBanco(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Sistema de Validação")
        self.geometry("800x600")
        self.create_widgets()
        self.movimento_file_path = ""
        self.ax_file_path = ""

    def create_widgets(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both")
       
        self.tab_nfs_e = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_nfs_e, text="Comparativo NFS-e")
       
        self.configurar_tab_nfs_e()

    def configurar_tab_nfs_e(self):
        frame = ttk.Frame(self.tab_nfs_e)
        frame.pack(pady=20)
       
        self.btn_select_movimento = ttk.Button(frame, text="Selecionar Planilha Movimento", command=lambda: self.load_file("movimento"))
        self.btn_select_ax = ttk.Button(frame, text="Selecionar Planilha AX", command=lambda: self.load_file("ax"))
        self.btn_select_movimento.pack(side=tk.LEFT, padx=10)
        self.btn_select_ax.pack(side=tk.LEFT, padx=10)
       
        self.btn_process = ttk.Button(frame, text="Processar", command=self.process_files)
        self.btn_process.pack(side=tk.LEFT, padx=10)
       
        self.text_result = tk.Text(self.tab_nfs_e, height=10, width=75)
        self.text_result.pack(pady=20)
       
        frame_botoes_inferiores = ttk.Frame(self.tab_nfs_e)
        frame_botoes_inferiores.pack(pady=10)
       
        self.btn_export = ttk.Button(frame_botoes_inferiores, text="Exportar Resultado", command=self.export_result)
        self.btn_export.pack(side=tk.LEFT, padx=10)
       
        self.btn_clear = ttk.Button(frame_botoes_inferiores, text="Limpar", command=self.clear_results)
        self.btn_clear.pack(side=tk.LEFT, padx=10)
       
        self.last_result = None
        self.loading_label = None

    def load_file(self, file_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            if file_type == "movimento":
                self.movimento_file_path = file_path
                self.btn_select_movimento.config(bg='green')
            elif file_type == "ax":
                self.ax_file_path = file_path
                self.btn_select_ax.config(bg='green')

    def process_files(self):
        if self.movimento_file_path and self.ax_file_path:
            self.loading_label = tk.Label(self.tab_nfs_e, text="Processando...", fg="blue")
            self.loading_label.pack()
            thread = threading.Thread(target=self.process_files_in_thread)
            thread.start()
        else:
            messagebox.showerror("Erro", "Por favor, selecione ambos os arquivos antes de processar.")

    def process_files_in_thread(self):
        try:
            consolidado_df = consolidar_planilhas_movimento(self.movimento_file_path)
            self.last_result = comparar_consolidado_ax(consolidado_df, self.ax_file_path)
            self.show_result(self.last_result)
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.show_error(str(e))
        finally:
            if self.loading_label:
                self.loading_label.destroy()

    def show_result(self, resultado):
        self.text_result.delete('1.0', tk.END)
        if not resultado.empty:
            # Selecionando as colunas "Status", "Fatura" e "Conta de cliente"
            resultado_filtrado = resultado[["Status", "Fatura", "Conta de cliente"]]

            # Remover linhas onde o Status é NaN
            resultado_filtrado = resultado_filtrado[resultado_filtrado['Status'].notna()]

            # Ajustando a apresentação dos valores
            resultado_filtrado['Fatura'] = resultado_filtrado['Fatura'].apply(lambda x: int(x) if isinstance(x, float) and x.is_integer() else x)

            # Removendo o sufixo .0 da coluna "Conta de cliente"
            resultado_filtrado['Conta de cliente'] = resultado_filtrado['Conta de cliente'].apply(lambda x: int(x) if isinstance(x, float) and x.is_integer() else x)

            self.text_result.insert(tk.END, resultado_filtrado.to_string(index=False))
        else:
            self.text_result.insert(tk.END, "Nenhuma correspondência encontrada.")

    def show_error(self, message):
        self.text_result.delete('1.0', tk.END)
        self.text_result.insert(tk.END, message)

    def export_result(self):
        if self.last_result is not None and not self.last_result.empty:
            file_type = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
            if file_type:
                if file_type.endswith('.xlsx'):
                    self.last_result.to_excel(file_type, index=False)
                elif file_type.endswith('.csv'):
                    self.last_result.to_csv(file_type, index=False)
                messagebox.showinfo("Exportar", "Resultado exportado com sucesso.")
            else:
                messagebox.showinfo("Ação necessária", "Exportação cancelada.")
        else:
            messagebox.showerror("Erro", "Nenhum resultado para exportar. Por favor, processe os arquivos primeiro.")

    def clear_results(self):
        self.text_result.delete('1.0', tk.END)
        self.btn_select_movimento.config(bg='light grey')
        self.btn_select_ax.config(bg='light grey')
        self.movimento_file_path = ""
        self.ax_file_path = ""
        self.last_result = None

if __name__ == "__main__":
    app = ApplicationBanco()
    app.mainloop()
