from Clinica import ApplicationClinica  # Garanta que as classes estejam renomeadas corretamente
from Comparador import ApplicationComparador
import tkinter as tk
from tkinter import ttk

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Menu Principal")
        self.geometry("600x200")
        self.create_widgets()

    def create_widgets(self):
        ttk.Button(self, text="Validar Clinica e AX", command=lambda: self.run_app(ApplicationClinica)).pack(pady=10)
        ttk.Button(self, text="NF-e Canceladas", command=lambda: self.run_app(ApplicationComparador)).pack(pady=10)

    def run_app(self, AppClass):
        self.withdraw()  # Oculta a janela principal
        app = AppClass(self)  # Passa a instância da janela principal
        app.grab_set()  # Garante que a atenção esteja na janela secundária
        app.wait_window()  # Espera a janela secundária fechar
        self.deiconify()  # Mostra a janela principal novamente

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
