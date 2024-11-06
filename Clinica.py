import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

def ler_planilha(caminho, skiprows=None):
    """Lê a planilha especificada e retorna um DataFrame."""
    extensao = caminho.split('.')[-1].lower()
    leitores = {
        'xlsx': lambda: pd.read_excel(caminho, skiprows=skiprows, engine='openpyxl'),
        'xls': lambda: pd.read_excel(caminho, skiprows=skiprows, engine='openpyxl'),
        'csv': lambda: pd.read_csv(caminho, skiprows=skiprows, delimiter=';')
    }
    try:
        return leitores.get(extensao, lambda: ValueError("Formato não suportado"))()
    except Exception as e:
        raise ValueError(f"Erro ao abrir o arquivo: {e}")

def converter_para_string(df, coluna):
    """Converte uma coluna numérica para string removendo casas decimais."""
    df[coluna] = df[coluna].astype(str).str.split('.').str[0]
    return df

def remover_total(df):
    """Remove a última linha se ela contiver 'Total'."""
    if 'Total' in df.iloc[-1].to_string():
        return df.iloc[:-1]
    return df

def comparar_planilhas(planilha_ax, planilha_clinica):
    ax_df = remover_total(ler_planilha(planilha_ax, skiprows=11))  # Lê e remove totais
    clinica_df = ler_planilha(planilha_clinica)  # Lê a Planilha Clínica diretamente

    ax_df = converter_para_string(ax_df, 'Fatura')
    clinica_df = converter_para_string(clinica_df, 'NFAX')

    faltando_no_ax = ax_df.loc[~ax_df['Fatura'].isin(clinica_df['NFAX']), 'Fatura'].drop_duplicates().reset_index(drop=True)
    faltando_na_clinica = clinica_df.loc[~clinica_df['NFAX'].isin(ax_df['Fatura']), 'NFAX'].drop_duplicates().reset_index(drop=True)

    return pd.DataFrame(faltando_na_clinica, columns=['NFAX']), pd.DataFrame(faltando_no_ax, columns=['Fatura'])

class ApplicationClinica(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Comparativo de Planilhas")
        self.geometry("1000x600")
        self.create_widgets()
        self.ax_file_path = None
        self.clinica_file_path = None

    def create_widgets(self):
        top_frame = tk.Frame(self)
        top_frame.pack(fill=tk.X)
        
        self.btn_select_ax = ttk.Button(top_frame, text="Selecionar Planilha AX", command=lambda: self.load_file("ax"))
        self.btn_select_ax.pack(side=tk.LEFT, padx=5, pady=10)

        self.btn_select_clinica = ttk.Button(top_frame, text="Selecionar Planilha Clínica", command=lambda: self.load_file("clinica"))
        self.btn_select_clinica.pack(side=tk.LEFT, padx=5)

        self.btn_process = ttk.Button(top_frame, text="Processar", command=self.process_files)
        self.btn_process.pack(side=tk.LEFT, padx=5, pady=10)

        self.btn_clear = ttk.Button(top_frame, text="Limpar", command=self.clear_results)
        self.btn_clear.pack(side=tk.LEFT, padx=5)

        self.result_frame = tk.Frame(self)
        self.result_frame.pack(fill=tk.BOTH, expand=True)

    def load_file(self, file_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            setattr(self, f"{file_type}_file_path", file_path)

    def process_files(self):
        if self.ax_file_path and self.clinica_file_path:
            self.show_loading_indicator()  # Mostra o indicador de carregamento
            thread = threading.Thread(target=self.process_files_in_thread)
            thread.start()
        else:
            messagebox.showerror("Erro", "Selecione ambas as planilhas antes de processar.")

    def show_loading_indicator(self):
        self.loading_label = tk.Label(self, text="Processando...", fg="blue")
        self.loading_label.pack()

    def process_files_in_thread(self):
        try:
            faltando_na_clinica, faltando_no_ax = comparar_planilhas(self.ax_file_path, self.clinica_file_path)
            self.show_results(faltando_na_clinica, 'NFAX', self.result_frame, "left", "Sistema Clínica - Camarões")
            self.show_results(faltando_no_ax, 'Fatura', self.result_frame, "right", "Sistema AX (Cancelar NF-e)")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
        finally:
            self.loading_label.destroy()  # Remove o indicador de carregamento

    def show_results(self, dataframe, coluna, parent, side, tabela_nome):
        frame = tk.Frame(parent)
        frame.pack(side=side, expand=True, fill=tk.BOTH, padx=10, pady=10)
        label = ttk.Label(frame, text=f"{tabela_nome}:")
        label.pack()
        tabela = self.criar_treeview(frame, [coluna], coluna)
        self.preencher_treeview(tabela, dataframe, coluna)
        self.adicionar_botao_export(frame, dataframe, coluna)

    def criar_treeview(self, parent, colunas, coluna):
        """Cria e configura a Treeview."""
        tabela = ttk.Treeview(parent, columns=colunas, show="headings")
        tabela.heading(coluna, text=coluna)
        tabela.column(coluna, anchor='center')
        tabela.pack(expand=True, fill=tk.BOTH)
        return tabela

    def preencher_treeview(self, tabela, dataframe, coluna):
        """Preenche a Treeview com os dados do DataFrame."""
        for _, row in dataframe.iterrows():
            tabela.insert('', 'end', values=(row[coluna],))

    def adicionar_botao_export(self, parent, df, coluna):
        """Adiciona botão de exportação para Excel."""
        btn_export = ttk.Button(parent, text=f"Exportar {coluna}", command=lambda: self.export_result(df, coluna))
        btn_export.pack(pady=10)

    def export_result(self, df, coluna):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            df.to_excel(filename, index=False)
            messagebox.showinfo("Sucesso", f"Dados {coluna} exportados com sucesso para Excel.")

    def clear_results(self):
        """Limpa os resultados exibidos e redefine os caminhos dos arquivos."""
        for widget in self.result_frame.winfo_children():
            widget.destroy()  # Limpa a Treeview e os botões de exportação
        self.ax_file_path = None
        self.clinica_file_path = None
        messagebox.showinfo("Limpar", "Resultados limpos com sucesso!")

if __name__ == "__main__":
    app = ApplicationClinica()
    app.mainloop()
