def generate_table(n, x):
    table = []
    for i in range(n):
        table.append((i % x) + 1)  # Génère un nombre de 1 à x en boucle
    return table

# Exemple d'utilisation
n = 10  # Longueur du tableau
x = 4   # Valeur maximale avant de repartir à 1

result = generate_table(n, x)
print(result)
