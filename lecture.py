import openpyxl

# Fichier Excel
fichier = "pharmacies.xlsx"

# Ouvrir le fichier
workbook = openpyxl.load_workbook(fichier)

# Première feuille
sheet = workbook.active

# Parcourir les lignes
for row in sheet.iter_rows(values_only=True):

    # Récupérer les 5 premières colonnes
    valeurs = row[:5]

    # Ignorer les cellules vides
    valeurs = [
        str(valeur).strip()
        for valeur in valeurs
        if valeur is not None and str(valeur).strip() != ""
    ]

    # Concaténer avec un espace
    resultat = " ".join(valeurs)

    print(resultat)

## pip3 install openpyxl

    ## 05eae13efb8bbcc1d45c9afcb3d7ee4295495f299c809314a277db5d731d77fd
