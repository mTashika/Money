import re

text = '02/09 COMMISSION RETRAIT PAR CARTE 10 200,67 --- 05/09 RETRAIT 9901-UNIT BRI P'
pattern_amount = r'(?<=\s)(\d+(?: \d{3})*\,\d{2})(?=\s*---)'

match = re.search(pattern_amount, text)
if match:
    print("Montant extrait :", match.group(1))
else:
    print("Aucun montant trouvé.")
