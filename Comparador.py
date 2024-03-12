import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

#Leitura das planilhas, desconsiderando as linhas de cabeçalho da planilha AX
def ler_planilha(caminho, skiprows=0, encoding='utf-8'):
    if caminho.endswith('.xlsx') or caminho.endswith('.xls'):
        return pd.read_excel(caminho, skiprows=skiprows)
    elif caminho.endswith('.csv'):
        try:
            # Tenta ler o arquivo com a codificação padrão e delimitador detectado automaticamente, ou seja, delimitado por ;
            return pd.read_csv(caminho, encoding=encoding, skiprows=skiprows, on_bad_lines='warn', delimiter=';')
        except UnicodeDecodeError:
            # Tenta novamente com uma codificação diferente se a primeira falhar
            return pd.read_csv(caminho, encoding='iso-8859-1', skiprows=skiprows, on_bad_lines='warn', delimiter=';')
    else:
        raise ValueError("Formato de arquivo não suportado.")

#Repassando os campos de busca, ou seja, relacionando as Tabelas com colunas correspondentes
#Coluna 'Fatura' da Planilha do AX
#Coluna 'Número do RPS' da Planilha Prefeitura SP
#Coluna 'Fatura' da Planilha do AX
#Coluna 'Nº da Nota Fiscal Eletrônica' da Planilha Prefeitura RJ
def encontrar_nfs_e(planilha_ax, planilha_prefeitura):
    ax_df = ler_planilha(planilha_ax, skiprows=8)
    prefeitura_df = ler_planilha(planilha_prefeitura)
    ax_df['Fatura'] = ax_df['Fatura'].astype(str).str.strip()  # Convertendo para string e removendo espaços em branco
    
    # Verifica qual coluna está presente na planilha da prefeitura
    coluna_prefeitura = ""
    if 'Número do RPS' in prefeitura_df.columns:
        coluna_prefeitura = 'Número do RPS'
        prefeitura_df[coluna_prefeitura] = prefeitura_df[coluna_prefeitura].astype(str).str.strip()  # Convertendo para string
    elif 'Nº da Nota Fiscal Eletrônica' in prefeitura_df.columns:
        coluna_prefeitura = 'Nº da Nota Fiscal Eletrônica'
        prefeitura_df[coluna_prefeitura] = prefeitura_df[coluna_prefeitura].astype(str).str.strip()  # Convertendo para string
    else:
        raise ValueError("A planilha da prefeitura deve conter a coluna 'Número do RPS' ou 'Nº da Nota Fiscal Eletrônica'")
    
    # Realiza a comparação com base na coluna identificada
    resultado = pd.merge(ax_df[['Fatura', 'Status']], prefeitura_df[[coluna_prefeitura]], left_on='Fatura', right_on=coluna_prefeitura, how='left')
    resultado_final = resultado.dropna()
    return resultado_final


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Validação")
        self.geometry("800x600")
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both")
        self.tab_nfs_e = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_nfs_e, text="Comparativo NFS-e")
        self.configurar_tab_nfs_e()

    def configurar_tab_nfs_e(self):
        frame = ttk.Frame(self.tab_nfs_e)
        frame.pack(pady=20)
        
        self.btn_select_ax = ttk.Button(frame, text="Selecionar Planilha AX", command=lambda: self.load_file("ax"))
        self.btn_select_prefeitura = ttk.Button(frame, text="Selecionar Planilha Prefeitura", command=lambda: self.load_file("prefeitura"))
        self.btn_select_ax.pack(side=tk.LEFT, padx=10)
        self.btn_select_prefeitura.pack(side=tk.LEFT, padx=10)
        
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
        
        self.ax_file_path = ""
        self.prefeitura_file_path = ""
        self.last_result = None

    def load_file(self, file_type):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
        if file_path:
            if file_type == "ax":
                self.ax_file_path = file_path
                self.btn_select_ax.config(bg='green')
            elif file_type == "prefeitura":
                self.prefeitura_file_path = file_path
                self.btn_select_prefeitura.config(bg='green')

    def process_files(self):
        if self.ax_file_path and self.prefeitura_file_path:
            try:
                self.last_result = encontrar_nfs_e(self.ax_file_path, self.prefeitura_file_path)
                self.show_result(self.last_result)
            except Exception as e:
                messagebox.showerror("Erro", str(e))
                self.show_error(str(e))
        else:
            messagebox.showerror("Erro", "Por favor, selecione ambos os arquivos antes de processar.")

    def show_result(self, resultado):
        self.text_result.delete('1.0', tk.END)
        if not resultado.empty:
            self.text_result.insert(tk.END, resultado.to_string(index=False))
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
        self.btn_select_ax.config(bg='light grey')
        self.btn_select_prefeitura.config(bg='light grey')
        self.ax_file_path = ""
        self.prefeitura_file_path = ""
        self.last_result = None

if __name__ == "__main__":
    app = Application()
    app.mainloop()