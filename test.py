import re

# Chaîne d'entrée
input_text = "Situation de vos comptes\nau 28 mars 2024"

# self.B_DATE est la partie fixe de la chaîne
self_B_DATE = "Situation de vos comptes\nau "

# Construire le pattern pour extraire le mois et l'année
pattern = rf'{re.escape(self_B_DATE)}\s*\d{{2}}\s*([a-zA-Z]+)\s(\d{{4}})'

# Exécuter l'expression régulière pour trouver une correspondance
match = re.search(pattern, input_text)

if match:
    mois = match.group(1)  # Premier groupe capturant le mois
    annee = match.group(2)  # Deuxième groupe capturant l'année
    print(f"Le mois est : {mois}")
    print(f"L'année est : {annee}")
else:
    print("Aucune correspondance trouvée.")
