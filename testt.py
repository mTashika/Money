import customtkinter as ctk

def get_button_frame(button):
    # Récupérer le 'frame' (ou le parent) du bouton
    return button.master

# Exemple d'utilisation :
root = ctk.CTk()

# Créer un Frame
frame = ctk.CTkFrame(root)
frame.pack(padx=10, pady=10)

# Créer un bouton dans ce Frame
button = ctk.CTkButton(frame, text="Click Me")
button.pack()

# Récupérer le frame du bouton
button_frame = get_button_frame(button)
print(button_frame)  # Affiche l'objet parent (le frame du bouton)

root.mainloop()
