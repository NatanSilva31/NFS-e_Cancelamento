import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def ler_planilha(caminho, skiprows=None):
    """Lê a planilha especificada e retorna um DataFrame."""
    if caminho.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(caminho, skiprows=skiprows, engine='openpyxl')
    elif caminho.endswith('.csv'):
        df = pd.read_csv(caminho, skiprows=skiprows, delimiter=';')
    else:
        raise ValueError("Formato de arquivo não suportado.")
    return df


def comparar_planilhas(planilha_ax, planilha_clinica):
    ax_df = ler_planilha(planilha_ax, skiprows=8)  # Pula as 8 primeiras linhas como cabeçalho na Planilha AX
    clinica_df = ler_planilha(planilha_clinica)  # Lê a Planilha Clínica diretamente

    # Conversão direta para strings para evitar problemas de conversão
    ax_df['Fatura'] = ax_df['Fatura'].astype(str).str.split('.').str[0]  # Remove casas decimais e converte para string
    clinica_df['NFAX'] = clinica_df['NFAX'].astype(str).str.split('.').str[0]  # Idem

    faltando_no_ax = ax_df.loc[~ax_df['Fatura'].isin(clinica_df['NFAX']), 'Fatura'].drop_duplicates().reset_index(drop=True)
    faltando_na_clinica = clinica_df.loc[~clinica_df['NFAX'].isin(ax_df['Fatura']), 'NFAX'].drop_duplicates().reset_index(drop=True)

    return pd.DataFrame(faltando_na_clinica, columns=['NFAX']), pd.DataFrame(faltando_no_ax, columns=['Fatura'])

class ApplicationClinica(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Comparativo de Planilhas")
        self.geometry("1000x600")
        self.create_widgets()

    def create_widgets(self):
        top_frame = tk.Frame(self)
        top_frame.pack(fill=tk.X)
        self.btn_select_ax = ttk.Button(top_frame, text="Selecionar Planilha AX", command=lambda: self.load_file("ax"))
        self.btn_select_ax.pack(side=tk.LEFT, padx=5, pady=10)

        self.btn_select_clinica = ttk.Button(top_frame, text="Selecionar Planilha Clínica", command=lambda: self.load_file("clinica"))
        self.btn_select_clinica.pack(side=tk.LEFT, padx=5)

        self.btn_process = ttk.Button(top_frame, text="Processar", command=self.process_files)
        self.btn_process.pack(side=tk.LEFT, padx=5, pady=10)

        self.result_frame = tk.Frame(self)
        self.result_frame.pack(fill=tk.BOTH, expand=True)

    def load_file(self, file_type):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
        if file_path:
            setattr(self, f"{file_type}_file_path", file_path)

    def process_files(self):
        if hasattr(self, 'ax_file_path') and hasattr(self, 'clinica_file_path'):
            faltando_na_clinica, faltando_no_ax = comparar_planilhas(self.ax_file_path, self.clinica_file_path)
            self.show_results(faltando_na_clinica, 'NFAX', self.result_frame, "left", "Sistema Clínica - Camarões")
            self.show_results(faltando_no_ax, 'Fatura', self.result_frame, "right", "Sistema AX (Cancelar NF-e)")
        else:
            messagebox.showerror("Erro", "Selecione ambas as planilhas antes de processar.")

    def show_results(self, dataframe, coluna, parent, side, tabela_nome):
        frame = tk.Frame(parent)
        frame.pack(side=side, expand=True, fill=tk.BOTH, padx=10, pady=10)
        label = ttk.Label(frame, text=f"{tabela_nome}:")
        label.pack()
        tabela = ttk.Treeview(frame, columns=[coluna], show="headings")
        tabela.heading(coluna, text=coluna)
        tabela.column(coluna, anchor='center')
        tabela.pack(expand=True, fill=tk.BOTH)
        for index, row in dataframe.iterrows():
            tabela.insert('', 'end', values=(row[coluna],))
        btn_export = ttk.Button(frame, text=f"Exportar {coluna}", command=lambda: self.export_result(dataframe, coluna))
        btn_export.pack(pady=10)

    def export_result(self, df, coluna):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            df.to_excel(filename, index=False)
            messagebox.showinfo("Sucesso", f"Dados {coluna} exportados com sucesso para Excel.")

if __name__ == "__main__":
    app = ApplicationClinica()
    app.mainloop()
