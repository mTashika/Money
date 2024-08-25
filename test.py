import customtkinter as ctk

def on_select(event):
    selected_value = selected_option.get()
    print(f"Valeur sélectionnée : {selected_value}")

# Création de la fenêtre principale
root = ctk.CTk()
root.title("Exemple de CTkComboBox avec StringVar")

# Variable StringVar pour stocker la valeur sélectionnée
selected_option = ctk.StringVar(value="Sélectionnez une option")

# Liste des valeurs pour la liste déroulante
values = ["Option 1", "Option 2", "Option 3", "Option 4"]

# Création du widget CTkComboBox, lié à la variable StringVar
combo = ctk.CTkComboBox(root, values=values, variable=selected_option)

# Configuration de la CTkComboBox
combo.pack(pady=20)

# Liaison d'un événement pour obtenir la valeur sélectionnée
combo.bind("<<ComboboxSelected>>", on_select)

# Lancer la boucle principale
root.mainloop()
