# medadom_map
recherche adresse, téléphone, horaire normal des fichers  google en donnant
input = établissement avec adresse, code postal, ville, pays
output = lignes avec nouveau colonnes  des fichiers google

# Compte SERPAPI sur 
https://serpapi.com/users/sign_in
avec plan Free 250 recherche/mois sur API google 

# Installation
- python
- openpyxl 
- serpapi

# Evironnement 
- export SERPAPI_KEY="clé_API"
- python3 -c 'import os; print(bool(os.getenv("SERPAPI_KEY")))'

# Run script
python3 main.py

