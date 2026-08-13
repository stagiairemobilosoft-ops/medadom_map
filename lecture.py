import openpyxl


def excel_to_queries(fichier_excel):
    workbook = openpyxl.load_workbook(fichier_excel)
    sheet = workbook.active

    lignes = []

    for row in sheet.iter_rows(values_only=True):

        # Prendre les 5 premières colonnes
        valeurs = row[:5]

        # Supprimer les cellules vides
        valeurs = [
            str(valeur).strip()
            for valeur in valeurs
            if valeur is not None and str(valeur).strip() != ""
        ]

        # Concaténer avec un espace
        resultat = " ".join(valeurs)

        lignes.append(resultat)

    return lignes

fichier = "pharmacies.xlsx"

recherches = excel_to_queries(fichier)

for recherche in recherches:
    print(recherche)

    
## pip3 install openpyxl

    ## 05eae13efb8bbcc1d45c9afcb3d7ee4295495f299c809314a277db5d731d77fd
