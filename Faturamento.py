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
            # Tenta ler o arquivo com a codificação padrão e delimitador detectado automaticamente, ou seja, delimitado por ;
            return pd.read_csv(caminho, encoding=encoding, skiprows=skiprows, on_bad_lines='warn', delimiter=';')
        except UnicodeDecodeError:
            # Tenta novamente com uma codificação diferente se a primeira falhar
            return pd.read_csv(caminho, encoding='iso-8859-1', skiprows=skiprows, on_bad_lines='warn', delimiter=';')
    else:
        raise ValueError("Formato de arquivo não suportado.")

# Repassando os campos de busca, ou seja, relacionando as Tabelas com colunas correspondentes
# Coluna 'Fatura' da Planilha do AX
# Coluna 'Título' da Planilha de Faturamento
def encontrar_nfs_e(planilha_ax, planilha_faturamento):
    ax_df = ler_planilha(planilha_ax, skiprows=11)
    faturamento_df = ler_planilha(planilha_faturamento, skiprows=7)
    
    # Remove duplicate columns in 'faturamento_df' by renaming them
    faturamento_df = faturamento_df.rename(columns=lambda x: f"{x}_{faturamento_df.columns.tolist().count(x)}" if faturamento_df.columns.tolist().count(x) > 1 else x)
    
    ax_df['Fatura'] = ax_df['Fatura'].astype(float)
    faturamento_df['Título'] = faturamento_df['Título'].astype(float)
    
    # Realiza o merge para pegar os registros de faturamento que não têm correspondência em AX (left anti join)
    resultado = pd.merge(faturamento_df[['Título']], ax_df[['Fatura']], left_on='Título', right_on='Fatura', how='left', indicator=True)
    
    # Filtra os registros da planilha de faturamento que não têm correspondência na planilha AX
    resultado_final = resultado[resultado['_merge'] == 'left_only'].drop(columns=['_merge'])
    
    # Seleciona as colunas para o resultado final, verificando se elas existem antes de tentar acessá-las
    colunas_resultado = ['Título']
    if 'Nº da Nota Fiscal Eletrônica' in resultado.columns:
        colunas_resultado.append('Nº da Nota Fiscal Eletrônica')
    
    resultado_final = resultado_final[colunas_resultado]
    return resultado_final

class ApplicationComparador(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)  # Chama o inicializador da classe base corretamente
        self.title("Sistema de Validação")
        self.geometry("800x600")
        self.create_widgets()
        self.ax_file_path = ""  # Inicializa a variável
        self.faturamento_file_path = ""  # Inicializa a variável

    def create_widgets(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both")
        
        self.tab_nfs_e = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_nfs_e, text="Comparativo Emitidas com AX")
        
        self.configurar_tab_nfs_e()

    def configurar_tab_nfs_e(self):
        frame = ttk.Frame(self.tab_nfs_e)
        frame.pack(pady=20)
        
        self.btn_select_ax = ttk.Button(frame, text="Selecionar Planilha AX", command=lambda: self.load_file("ax"))
        self.btn_select_faturamento = ttk.Button(frame, text="Selecionar Planilha Emitidas", command=lambda: self.load_file("faturamento"))
        self.btn_select_ax.pack(side=tk.LEFT, padx=10)
        self.btn_select_faturamento.pack(side=tk.LEFT, padx=10)
        
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
        self.faturamento_file_path = ""
        self.last_result = None
        self.loading_label = None  # Inicializa o label de carregamento

    def load_file(self, file_type):
        file_path = filedialog.askopenfilename()
        if file_path:
            if file_type == "ax":
                self.ax_file_path = file_path
                self.btn_select_ax.config(bg='green')
            elif file_type == "faturamento":
                self.faturamento_file_path = file_path
                self.btn_select_faturamento.config(bg='green')

    def process_files(self):
        if self.ax_file_path and self.faturamento_file_path:
            self.loading_label = tk.Label(self.tab_nfs_e, text="Processando...", fg="blue")
            self.loading_label.pack()
            thread = threading.Thread(target=self.process_files_in_thread)
            thread.start()
        else:
            messagebox.showerror("Erro", "Por favor, selecione ambos os arquivos antes de processar.")

    def process_files_in_thread(self):
        try:
            self.last_result = encontrar_nfs_e(self.ax_file_path, self.faturamento_file_path)
            self.show_result(self.last_result)
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.show_error(str(e))
        finally:
            if self.loading_label:
                self.loading_label.destroy()  # Remove o indicador de carregamento

    def show_result(self, resultado):
        self.text_result.delete('1.0', tk.END)
        if not resultado.empty:
            resultado_str = resultado.applymap(lambda x: int(x) if isinstance(x, float) and x.is_integer() else x)
            self.text_result.insert(tk.END, resultado_str.to_string(index=False))
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
        self.btn_select_faturamento.config(bg='light grey')
        self.ax_file_path = ""
        self.faturamento_file_path = ""
        self.last_result = None
        
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    app = ApplicationComparador()
    app.mainloop()
