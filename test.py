import os
import serpapi
from difflib import SequenceMatcher

api_key = os.getenv("SERPAPI_KEY")

if not api_key:
    raise RuntimeError("SERPAPI_KEY n'est pas définie")

client = serpapi.Client(api_key=api_key)


# ==========================================
# DONNÉES À RECHERCHER
# ==========================================

nom_recherche = "La Pharmacie Du Phenix"
adresse_recherchee = "43 Rue Du Casino"
code_postal = "57800"
ville = "Freyming-Merlebach"
pays = "France"

query = f"""
{nom_recherche}
{adresse_recherchee}
{code_postal} {ville}
{pays}
"""


# ==========================================
# RECHERCHE GOOGLE MAPS
# ==========================================

results = client.search({
    "engine": "google_maps",
    "type": "search",
    "q": query,
    "google_domain": "google.fr",
    "hl": "fr"
})


# ==========================================
# VÉRIFICATION
# ==========================================

place = results.get("place_results", {})


if not place:
    print("\n❌ AUCUNE FICHE GOOGLE TROUVÉE")
    print("Établissement probablement introuvable.")
    exit()


nom_google = place.get("title", "")
adresse_google = place.get("address", "")


print("\n===== FICHE GOOGLE TROUVÉE =====")

print("Nom       :", nom_google)
print("Adresse   :", adresse_google)
print("Téléphone :", place.get("phone"))


# ==========================================
# COMPARAISON DU NOM
# ==========================================

def normaliser(texte):
    return (
        texte.lower()
        .replace("-", " ")
        .replace(",", " ")
        .replace(".", " ")
        .strip()
    )


nom1 = normaliser(nom_recherche)
nom2 = normaliser(nom_google)

similarite_nom = SequenceMatcher(None, nom1, nom2).ratio()


print("\n===== VALIDATION =====")

print(f"Similarité du nom : {similarite_nom:.2f}")


# ==========================================
# VÉRIFICATION DE L'ADRESSE
# ==========================================

adresse_complete_recherchee = normaliser(
    f"{adresse_recherchee} {code_postal} {ville}"
)

adresse_complete_google = normaliser(adresse_google)


mots_adresse = adresse_complete_recherchee.split()

mots_trouves = sum(
    1 for mot in mots_adresse
    if mot in adresse_complete_google
)

score_adresse = mots_trouves / len(mots_adresse)


print(f"Correspondance adresse : {score_adresse:.2f}")


# ==========================================
# DÉCISION
# ==========================================

if similarite_nom >= 0.70 and score_adresse >= 0.50:

    print("\n✅ ÉTABLISSEMENT VALIDÉ")
    print("La fiche Google correspond probablement à l'établissement recherché.")

else:

    print("\n⚠️ CORRESPONDANCE INCERTAINE")
    print("La fiche trouvée doit être vérifiée.")