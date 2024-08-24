import re

# Exemple de texte
STRING = "Situation de vos comptes\nau 28 mars 2024"
INTER_LINE = ' --- '
# Balise de départ
balise_name = 'ACHAT CB '
B_DATE = "Situation de vos comptes\nau "

pattern_date = rf"\d{{2}}\.\d{{2}}\.\d{{2}}"
pattern_amount_g = rf'\d{{1,3}}(?: \d{{3}})*,\d{{2}}'
pattern_special_chars = rf'[^\w\s.]'

# pattern_name   = rf'{re.escape(INTER_LINE)}\s*([A-Za-z]+\s+[A-Za-z]+)'

pattern_name   = rf'{re.escape(balise_name)}\s*([A-Za-z\s.]+?)(?=\d|{pattern_special_chars}|{INTER_LINE}|{pattern_amount_g}|{pattern_date})'
pattern_amount = rf'({pattern_amount_g})(?={INTER_LINE})'
pattern_date = rf'{re.escape(B_DATE)}\s*\d{{2}}\s*([a-zA-Z]+)\s(\d{{4}})'



match = re.search(pattern_name, STRING)
if match:
    name = match.group(1)
    print("Name:", name)
else:
    print("No name found")

match = re.search(pattern_amount, STRING)
if match:
    amount = match.group(1)
    print("Amount:", amount)
else:
    print("No amount found")
    
match = re.search(pattern_date, STRING)
if match:
    month = match.group(1)
    year = match.group(2)
    print(f"Month: {month}")
    print(f"Year: {year}")
else:
    print("No date found")

