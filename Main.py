# Antes de tudo, renomeie as classes Application nos respectivos arquivos para ApplicationClinica e ApplicationComparador.
from Clinica import ApplicationClinica
from Comparador import ApplicationComparador
import tkinter as tk
from tkinter import messagebox, ttk

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Menu Principal")
        self.geometry("600x200")
        self.create_widgets()

    def create_widgets(self):
        ttk.Button(self, text="Validar Clinica e AX", command=self.run_clinica).pack(pady=10)
        ttk.Button(self, text="NF-e Canceladas", command=self.run_comparador).pack(pady=10)

    def run_clinica(self):
        self.destroy()
        app = ApplicationClinica()
        app.mainloop()

    def run_comparador(self):
        self.destroy()
        app = ApplicationComparador()
        app.mainloop()

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
