"""
inspection_bilan.py — LECTURE SEULE.

Prépare l'ajout d'une ligne « Lemon Tree Riad » à l'actif du bilan.

Pour poser cette ligne au bon endroit, avec une formule cohérente avec les
autres, il faut savoir comment le bilan est bâti : où commence l'actif, ce
que contiennent les lignes voisines, et si elles sont calculées ou saisies
à la main. Ce script ne suppose rien, il lit les formules réelles.

Il regarde aussi les en-têtes des trois onglets sources, pour vérifier que
les colonnes (date, catégorie, montant) sont bien celles que le script
Apps Script utilise dans ses formules de compte de résultat.

Ne modifie rien. Aucune écriture, aucune cellule touchée.

USAGE
  python inspection_bilan.py
"""

import os
import sys

import gspread
from google.auth.exceptions import RefreshError

SHEET_ID = "1OUsQHnN9J241pLTb3pyrKmj7MTIrjCI6PNAw4FZ835U"

DOSSIER = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS = os.path.join(DOSSIER, "credentials.json")
JETON = os.path.join(DOSSIER, "authorized_user.json")

SOURCES = ("Caisse", "Banque", " Saisies CB Karim")


def ouvrir_classeur():
    """
    Ouvre le classeur, en réparant le cas du jeton périmé.

    L'écran de consentement OAuth est en mode Test côté Google, et dans ce
    mode le jeton de rafraîchissement est révoqué au bout de sept jours.
    gspread ne sait pas se relancer tout seul dans ce cas : il lève
    RefreshError et s'arrête. On efface alors le jeton mort et on refait
    le tour par le navigateur — c'est la seule sortie.

    Rien d'autre n'est touché : credentials.json, lui, ne bouge jamais.
    """
    def connexion():
        return gspread.oauth(
            credentials_filename=CREDENTIALS,
            authorized_user_filename=JETON,
        ).open_by_key(SHEET_ID)

    try:
        return connexion()
    except RefreshError:
        print(
            "\nLe jeton d'accès Google a expiré.\n"
            "Google le révoque tous les sept jours tant que l'écran de\n"
            "consentement du projet reste en mode « Test ».\n\n"
            "Je vais effacer le jeton périmé et rouvrir l'autorisation dans\n"
            "ton navigateur. Rien n'est modifié dans le classeur.\n"
        )
        input("Entrée pour continuer, Ctrl+C pour renoncer : ")
        os.remove(JETON)
        return connexion()


def colonne(index):
    """Numéro de colonne (1 = A) vers sa lettre. Suffisant jusqu'à Z."""
    return chr(ord("A") + index - 1)


def montre_feuille(feuille, lignes_max=60, colonnes_max=6):
    """Affiche le contenu brut (formules) d'un onglet, cellule par cellule."""
    grille = feuille.get_all_values(value_render_option="FORMULA")

    for i, ligne in enumerate(grille[:lignes_max], start=1):
        cellules = []
        for j, valeur in enumerate(ligne[:colonnes_max], start=1):
            if valeur != "":
                cellules.append(f"{colonne(j)}{i}={valeur}")
        if cellules:
            print("   " + "  |  ".join(c[:60] for c in cellules))


def main():
    if not os.path.exists(CREDENTIALS):
        print("credentials.json introuvable dans :", DOSSIER)
        sys.exit(1)

    classeur = ouvrir_classeur()
    print("Classeur :", classeur.title)
    print("\nOnglets présents :")
    for f in classeur.worksheets():
        print(f"   {f.title!r}  ({f.row_count} x {f.col_count})")

    # ---------- 1. Les bilans, en entier ----------
    for nom in ("Bilan Année 1", "Bilan Année 2"):
        print("\n" + "=" * 72)
        print(f"  {nom}")
        print("=" * 72)
        try:
            montre_feuille(classeur.worksheet(nom))
        except gspread.WorksheetNotFound:
            print("   onglet introuvable")

    # ---------- 2. Les en-têtes des sources ----------
    print("\n" + "=" * 72)
    print("  Onglets sources — ligne d'en-tête et première ligne de données")
    print("=" * 72)
    for nom in SOURCES:
        print(f"\n  {nom!r}")
        try:
            feuille = classeur.worksheet(nom)
            grille = feuille.get_all_values()[:2]
            for i, ligne in enumerate(grille, start=1):
                for j, valeur in enumerate(ligne[:10], start=1):
                    if valeur != "":
                        print(f"   {colonne(j)}{i} = {valeur[:40]}")
                print()
            print(f"   dernière ligne : {feuille.row_count}")
        except gspread.WorksheetNotFound:
            print("   onglet introuvable")


if __name__ == "__main__":
    main()
