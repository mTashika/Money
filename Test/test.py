import customtkinter as ctk
import threading
import time
from PIL import Image, ImageTk  # Importer Pillow pour gérer les images

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Exemple de chargement avec GIF")
        self.geometry("300x200")

        # Créer un bouton
        self.button = ctk.CTkButton(self, text="Lancer la fonction", command=self.run_function)
        self.button.pack(pady=20)

        # Charger le GIF avec Pillow
        self.gif = Image.open("src\loading_gif.gif")

        # Extraire les frames du GIF
        self.frames = []
        try:
            while True:
                self.frames.append(ImageTk.PhotoImage(self.gif.copy()))
                self.gif.seek(len(self.frames))  # Passer à la prochaine frame
        except EOFError:
            pass  # Fin du GIF

        self.loading_index = 0  # Initialiser l'index de frame

        # Créer un label pour afficher le GIF
        self.loading_label = ctk.CTkLabel(self,text='')
        self.loading_label.place(relx=0.5, rely=0.5, anchor='center')
        self.loading_label.place_forget()  # Cacher au début

    def run_function(self):
        # Exécuter la fonction longue dans un thread séparé
        threading.Thread(target=self.long_running_function).start()

        # Afficher le GIF et désactiver les boutons
        self.show_loading()

    def show_loading(self):
        # Désactiver les boutons ou les fonctionnalités
        self.disable_update()
        self.disable_create()

        # Afficher le label de chargement avec GIF
        self.loading_label.place(relx=0.5, rely=0.5, anchor='center')
        self.animate_gif()  # Commencer l'animation du GIF

    def animate_gif(self):
        frame = self.frames[self.loading_index]  # Obtenir la frame courante
        self.loading_label.configure(image=frame)  # Mettre à jour l'image du label
        self.loading_index = (self.loading_index + 1) % len(self.frames)  # Passer à la frame suivante

        # Reprogrammer pour afficher la prochaine frame
        self.after(100, self.animate_gif)  # Délai de 100ms entre les frames

    def long_running_function(self):
        # Simuler une fonction longue qui prend du temps
        time.sleep(5)  # Remplacez ceci par votre logique

        # Masquer le GIF et réactiver les fonctionnalités dans le thread principal
        self.after(0, self.hide_loading)

    def hide_loading(self):
        # Cacher le label de chargement et réactiver les boutons
        self.loading_label.place_forget()
        self.button.configure(state="normal")

    def disable_update(self):
        # Logique pour désactiver la fonctionnalité de mise à jour
        print("Mise à jour désactivée")

    def disable_create(self):
        # Logique pour désactiver la fonctionnalité de création
        print("Création désactivée")

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")  # Mode sombre
    ctk.set_default_color_theme("blue")  # Thème par défaut
    app = App()
    app.mainloop()
