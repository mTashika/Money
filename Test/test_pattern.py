import re

# Exemple de texte
STRING = "05/06 ACHAT CB L.I.D.L. 04.06.24 43,12 --- CARTE NUMERO 219"
INTER_LINE = ' --- '
# Balise de départ
balise_name = 'ACHAT CB '

pattern_date = rf"\d{{2}}\.\d{{2}}\.\d{{2}}"
pattern_amount_g = rf'\d{{1,3}}(?: \d{{3}})*,\d{{2}}'
pattern_special_chars = rf'[^\w\s.]'

# pattern_name   = rf'{re.escape(INTER_LINE)}\s*([A-Za-z]+\s+[A-Za-z]+)'

pattern_name   = rf'{re.escape(balise_name)}\s*([A-Za-z\s.]+?)(?=\d|{pattern_special_chars}|{INTER_LINE}|{pattern_amount_g}|{pattern_date})'
pattern_amount = rf'({pattern_amount_g})(?={INTER_LINE})'


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
