from Clinica import ApplicationClinica  # Garanta que as classes estejam renomeadas corretamente
from Comparador import ApplicationComparador
from Banco import ApplicationBanco
import tkinter as tk
from tkinter import ttk

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Menu Principal")
        self.geometry("600x200")
        self.iconbitmap("Imagens/512x512bb.ico")
        self.create_widgets()

    def create_widgets(self):
        """Cria os botões para abrir diferentes aplicações."""
        apps = {
            "Validar Clinica e AX": ApplicationClinica,
            "NF-e Canceladas": ApplicationComparador,
            "Validação Banco": ApplicationBanco,
        }
        
        for text, app_class in apps.items():
            self.create_button(text, app_class)

    def create_button(self, text, app_class):
        """Cria um botão para abrir a aplicação especificada."""
        button = ttk.Button(self, text=text, command=lambda: self.run_app(app_class))
        button.pack(pady=10)

    def run_app(self, AppClass):
        """Executa a aplicação especificada e gerencia a janela principal."""
        self.withdraw()  # Oculta a janela principal
        app = AppClass(self)  # Passa a instância da janela principal para a aplicação
        app.grab_set()  # Garante que a atenção esteja na janela secundária
        app.wait_window(app)  # Espera a janela secundária fechar
        self.deiconify()  # Mostra a janela principal novamente

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
