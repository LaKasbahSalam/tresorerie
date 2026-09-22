"""
test_connexion.py — vérifie que la chaîne Python -> Google Sheets fonctionne.

Ce script ne modifie RIEN. Il lit seulement, pour prouver que
l'authentification passe et que la bonne feuille est atteinte.

PRÉREQUIS
  - pip install gspread google-auth google-auth-oauthlib   (déjà fait)
  - credentials.json (le fichier téléchargé depuis Google Cloud) dans
    le même dossier que ce script

PREMIER LANCEMENT
  Une fenêtre de navigateur s'ouvre pour autoriser l'accès.
  Google affichera « Cette application n'est pas validée » : c'est normal,
  c'est ton propre projet en mode test.
      -> Paramètres avancés -> Accéder à Trésorerie Kasbah (non sécurisé)
  Un fichier authorized_user.json est alors créé à côté de ce script.
  Il contient ton jeton d'accès : ne le partage pas, ne le commite pas.

LANCEMENTS SUIVANTS
  Plus aucune fenêtre — le jeton se renouvelle tout seul.

USAGE
  python test_connexion.py
"""

import os
import sys

import gspread

# ---------- Réglages ----------

# La Sheet de trésorerie. L'ID est le segment de l'URL entre /d/ et /edit.
SHEET_ID = "1OUsQHnN9J241pLTb3pyrKmj7MTIrjCI6PNAw4FZ835U"

# Onglet à lire pour le test. Le nom doit être exact, accents compris.
ONGLET = "Caisse"

# Fichiers d'authentification, cherchés à côté de ce script plutôt que
# dans le dossier courant : le script marche alors quel que soit l'endroit
# d'où tu le lances.
DOSSIER = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS = os.path.join(DOSSIER, "credentials.json")
JETON = os.path.join(DOSSIER, "authorized_user.json")


def derniere_ligne_avec_code(valeurs):
    """
    Numéro de la dernière ligne portant un code en colonne A.

    Même logique que derniereLigneAvecCode() dans Code.js, et pour la même
    raison : une note isolée bien plus bas dans la feuille ferait mentir un
    simple "nombre de lignes". On remonte depuis le bas jusqu'à un vrai code.

    `valeurs` est la grille complète renvoyée par get_all_values(),
    lignes 1..N, la 1 portant les en-têtes.
    """
    for i in range(len(valeurs) - 1, 0, -1):  # s'arrête avant la ligne d'en-tête
        ligne = valeurs[i]
        if ligne and ligne[0].strip():
            return i + 1  # get_all_values est indexé à 0, les lignes à 1
    return 0


def main():
    if not os.path.exists(CREDENTIALS):
        print("credentials.json introuvable dans :", DOSSIER)
        print("Place-y le fichier téléchargé depuis Google Cloud Console.")
        sys.exit(1)

    print("Connexion à Google...")
    # gspread.oauth gère tout le cycle : ouverture du navigateur au premier
    # passage, stockage du jeton, et rafraîchissement silencieux ensuite.
    client = gspread.oauth(
        credentials_filename=CREDENTIALS,
        authorized_user_filename=JETON,
    )

    classeur = client.open_by_key(SHEET_ID)
    print("Classeur ouvert :", classeur.title)
    print("Onglets :", ", ".join(f.title for f in classeur.worksheets()))

    feuille = classeur.worksheet(ONGLET)
    valeurs = feuille.get_all_values()

    derniere = derniere_ligne_avec_code(valeurs)
    if derniere == 0:
        print("Aucune ligne de données trouvée dans l'onglet", ONGLET)
        return

    ligne = valeurs[derniere - 1]
    print()
    print("Dernière écriture, ligne", derniere, ":")
    print("  code        :", ligne[0])
    print("  date        :", ligne[1])
    print("  description :", ligne[2])
    print("  montant     :", ligne[3])
    print("  catégorie   :", ligne[4])
    print("  SOLDE (H)   :", ligne[7])
    print()
    print("Connexion opérationnelle — lecture confirmée, rien n'a été modifié.")


if __name__ == "__main__":
    main()
