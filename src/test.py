import customtkinter as ctk

def get_all_ctk_frames(widget):
    frames = []
    # Vérifie si le widget actuel est une CTkFrame, et l'ajoute à la liste
    if isinstance(widget, ctk.CTkFrame):
        frames.append(widget)
    
    # Parcourt tous les enfants du widget actuel
    for child in widget.winfo_children():
        # Appelle récursivement la fonction pour chercher dans les enfants
        frames.extend(get_all_ctk_frames(child))
    
    return frames

# Initialiser CustomTkinter
ctk.set_appearance_mode("dark")  # "dark" ou "light"
ctk.set_default_color_theme("blue")  # Thème par défaut

# Création de la fenêtre principale (root)
root = ctk.CTk()
root.title("Extraction des CTkFrames")

# Exemple de CTkFrame imbriquées
frame1 = ctk.CTkFrame(root, width=200, height=200, fg_color="red")
frame1.pack(pady=10)

frame2 = ctk.CTkFrame(frame1, width=150, height=150, fg_color="blue")
frame2.pack(pady=10)

frame3 = ctk.CTkFrame(frame2, width=100, height=100, fg_color="green")
frame3.pack(pady=10)

# Obtenir toutes les CTkFrames depuis root
all_ctk_frames = get_all_ctk_frames(root)

# Afficher les frames trouvées
print(f"Nombre total de CTkFrames : {len(all_ctk_frames)}")
for i, frame in enumerate(all_ctk_frames, start=1):
    print(f"CTkFrame {i}: {frame} avec fg_color = {frame._fg_color}")

# Démarrer la boucle principale de l'application CustomTkinter
root.mainloop()
