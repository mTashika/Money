import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sélection de PDF")
        self.geometry("400x300")

        # Variable pour stocker le chemin des fichiers PDF
        self.pdf_path = tk.StringVar()  # Utiliser StringVar pour les chemins de fichiers

        # Créer un bouton pour parcourir les fichiers PDF
        self.browse_button = ctk.CTkButton(self, text="Parcourir PDF", command=self.browse_pdf)
        self.browse_button.pack(pady=20)

        # Créer un label pour afficher le chemin des fichiers sélectionnés
        self.pdf_label = ctk.CTkLabel(self, textvariable=self.pdf_path)
        self.pdf_label.pack(pady=20)

    def browse_pdf(self):
        # Utiliser askopenfilenames pour permettre la sélection de plusieurs fichiers
        file_paths = filedialog.askopenfilenames(title="Sélectionner des fichiers PDF", filetypes=[("PDF Files", "*.pdf")])
        if file_paths:
            # Rejoindre les chemins de fichiers en une seule chaîne
            self.pdf_path.set("; ".join(file_paths))  # Afficher tous les chemins séparés par un point-virgule
            pass
if __name__ == "__main__":
    ctk.set_appearance_mode("dark")  # Mode sombre
    ctk.set_default_color_theme("blue")  # Thème par défaut
    app = App()
    app.mainloop()
