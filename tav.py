import os
import serpapi

api_key = os.getenv("SERPAPI_KEY")

if not api_key:
    raise RuntimeError("SERPAPI_KEY n'est pas définie")

client = serpapi.Client(api_key=api_key)

## recherche google maps
query = """
La Pharmacie Du Phenix 43 Rue Du Casino 57800 Freyming-Merlebach France
"""

results = client.search({
    "engine": "google_maps",
    "type": "search",
    "q": query,
    "google_domain": "google.fr",
    "hl": "fr"
})

place = results.get("place_results", {})

print("\n===== FICHE GOOGLE =====")

print("Nom       :", place.get("title"))
print("Adresse   :", place.get("address"))
print("Téléphone :", place.get("phone"))

print("\n===== HORAIRES =====")

for jour in place.get("hours", []):
    for nom_jour, horaires in jour.items():
        print(f"{nom_jour.capitalize():10} : {horaires}")

### export SERPAPI_KEY="clé_API"
### python3 -c 'import os; print(bool(os.getenv("SERPAPI_KEY")))' 