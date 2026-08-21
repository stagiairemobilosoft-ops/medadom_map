# medadom_map
- recherche adresse, téléphone, horaire normal des fichers  google en donnant
- input = établissement avec adresse, code postal, ville, pays
output = lignes avec nouveau colonnes  des fichiers google

# Compte SERPAPI sur 
https://serpapi.com/users/sign_in
avec plan Free 250 recherche/mois sur API google 

# Installation
- python
- openpyxl 
- serpapi

- python3 -m pip install openpyxl google-search-results
- python3 -m pip install --upgrade serpapi

# Evironnement WINDOW
- $env:SERPAPI_KEY="clé_API"
- python -c "import os; print(bool(os.getenv('SERPAPI_KEY')))"


# Evironnement MAC
- export SERPAPI_KEY="clé_API"
- python3 -c 'import os; print(bool(os.getenv("SERPAPI_KEY")))'

# 1- Run script Medadom - 1 recherches
- python3 medadom.py
- contact, adresse, horaire normale

# 1- Run script Plus - 1 recherches
- python3 plus.py
- contact, adresse, horaire normale, note, nombre d'avis, site web, nombre photo, nombre photo propriétaire


# 1- Run script All - 4 recherches
- python3 all.py
- contact, adresse, horaire normale, note, nombre d'avis, site web, nombre photo, nombre photo propriétaire, date dernier post

