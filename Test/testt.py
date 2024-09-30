import customtkinter as ctk

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Deux Frames à Égalité")
        self.geometry("400x300")

        # Créer la Frame parent
        self.parent_frame = ctk.CTkFrame(self)
        self.parent_frame.pack(expand=True, fill="both")  # Remplit tout l'espace disponible

        # Créer la première Frame
        self.frame1 = ctk.CTkFrame(self.parent_frame, fg_color="lightblue")
        self.frame1.pack(side="left", expand=True, fill="both")  # Remplir la moitié à gauche

        # Créer la deuxième Frame
        self.frame2 = ctk.CTkFrame(self.parent_frame, fg_color="lightgreen")
        self.frame2.pack(side="right", expand=True, fill="both")  # Remplir la moitié à droite

        # Optionnel : Ajouter du contenu aux frames
        label1 = ctk.CTkLabel(self.frame1, text="Frame 1")
        label1.pack(pady=20)  # Ajouter un label dans la première frame

        label2 = ctk.CTkLabel(self.frame2, text="Frame 2")
        label2.pack(pady=20)  # Ajouter un label dans la deuxième frame

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")  # Mode sombre
    ctk.set_default_color_theme("blue")  # Thème par défaut
    app = App()
    app.mainloop()
