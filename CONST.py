INIT_SHEET_NAME = "Tool"
INIT_SHEET_SOLD_TITLE = "Solde d'initialisation"
INIT_SHEET_CAT1_TITLE = "Cat1"
INIT_SHEET_CAT2_TITLE = "Cat2"

INIT_SHEET_MONTH_CELL = ['B','B1']
INIT_SHEET_SOLD_CELL = ['A','A2']

INIT_SHEET_CAT1_COL = "C"
INIT_SHEET_CAT2_COL = "D"



CAT_1 = [
    "Abonnement", "Cadeau", "Divers", "Exeptionnel", "Manger", "Noël", 
    "Paie", "Revolut", "Shopping", "Sortie", "Sport", "Transport", 
    "Virement", "Virement Moi", "Voyage", "Cotisation", "Paypal", "Lydia"
]

CAT_2 = [
    "Essence", "Spotify", "Course", "Resto", "Bar", "Ratp", "Onera", 
    "Interim", "Train", "Mère Anaïs", "Anaïs", "Commande", "Parent", 
    "Basic fit", "Bus", "Thalès", "Parking", "Péage"
]

MONTH_ANG = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]

MONTH_FR = [
    "janvier", "fevrier", "mars", "avril", "mai", "juin",
    "juillet", "aout", "septembre", "octobre", "novembre", "decembre"
    ]

from openpyxl.styles import Alignment
ALIGN_CENTER= Alignment(horizontal='center', vertical='center', wrap_text=True)

START_MONTH = [3,1]
NB_MORE_FO_MONTH = 5
MONTH_ROW_TITLE = 2

BALISE_NEW_MONTH = '#'
BALISE_SOLD_MONTH = '>'

MONTH_SOLD_BALISE = f"{BALISE_SOLD_MONTH}solde fiche de paye (fin du mois)"
MONTH_SOLD_EST = "Solde estimé :"
MONTH_VERIF = "Verif"
MONTH_PREV_SOLD = ['Solde précédent',2,1]


NOT_REAL_ACTION = ['Virement Moi','NC']
