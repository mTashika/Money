import re

INTER_LINE = ' --- '
balise_name = 'ACHAT CB'

# Exemple de chaîne
text = ' 26/01 ACHAT CB ASOS.COM 10 091,99 --- EUR 91,99 CARTE NO 219'

pattern_date = rf"\d{{2}}\.\d{{2}}\.\d{{2}}"
pattern_special_chars = rf'[^\w\s.]'
# Pattern pour détecter les montants
pattern_amount_detection = r'(?<!\d)(?:\d{1,3}(?: \d{3})*|\d+)(?:,\d{2})(?!\d)'

# Motif qui capture les caractères mais exclut les formats de montants
pattern_name = rf'{re.escape(balise_name)}\s*([A-Za-z0-9.\s]*?)(?=\s*{pattern_amount_detection}|{pattern_special_chars}|{INTER_LINE}|{pattern_date})'

# Recherche dans la chaîne
match = re.search(pattern_name, text)

if match:
    result = match.group(1)
    print(result)
else:
    print("Aucun trouvé.")
