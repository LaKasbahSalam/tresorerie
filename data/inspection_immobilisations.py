"""
inspection_immobilisations.py — LECTURE SEULE.

Répond à une question précise : y a-t-il double emploi entre les lignes
comptées en charge au compte de résultat (« Entretien & réparations »,
« Aménagements & travaux ») et celles immobilisées au bilan ?

Pour y répondre il faut savoir COMMENT le classeur sélectionne ce qui est
immobilisé. Ce script ne suppose rien : il lit les formules réelles.

Ne modifie rien. Aucune écriture, aucune cellule touchée.

USAGE
  python inspection_immobilisations.py
"""

import os
import sys

import gspread

SHEET_ID = "1OUsQHnN9J241pLTb3pyrKmj7MTIrjCI6PNAw4FZ835U"

DOSSIER = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS = os.path.join(DOSSIER, "credentials.json")
JETON = os.path.join(DOSSIER, "authorized_user.json")

# Libellés dont on veut voir la formule dans les comptes de résultat.
CIBLES_CDR = [
    "entretien",
    "amenagement",
    "aménagement",
    "dotation",
    "mobilier",
]


def sans_accents(s):
    """Comparaison tolérante aux accents et à la casse."""
    table = str.maketrans("àâäéèêëïîôöùûüç", "aaaeeeeiioouuuc")
    return s.lower().translate(table)


def montre_lignes(feuille, cibles, colonnes=4):
    """
    Affiche les formules des lignes dont la colonne A contient un des
    termes cherchés. `colonnes` limite le nombre de colonnes affichées :
    les formules mensuelles se répètent, une ou deux suffisent à
    comprendre le mécanisme.
    """
    formules = feuille.get_all_values(value_render_option="FORMULA")
    trouve = 0

    for i, ligne in enumerate(formules, start=1):
        if not ligne or not ligne[0]:
            continue
        libelle = sans_accents(str(ligne[0]))
        if not any(c in libelle for c in cibles):
            continue

        trouve += 1
        print(f"\n  --- ligne {i} : {ligne[0]} ---")
        for j, cellule in enumerate(ligne[1:colonnes + 1], start=1):
            if cellule:
                col = chr(ord("A") + j)
                print(f"  {col}{i} = {cellule}")

    if trouve == 0:
        print("  (aucune ligne correspondante)")


def main():
    if not os.path.exists(CREDENTIALS):
        print("credentials.json introuvable dans :", DOSSIER)
        sys.exit(1)

    client = gspread.oauth(
        credentials_filename=CREDENTIALS,
        authorized_user_filename=JETON,
    )
    classeur = client.open_by_key(SHEET_ID)
    print("Classeur :", classeur.title)

    # ---------- 1. Les comptes de résultat ----------
    for nom in ("CdR Année 1", "CdR Année 2"):
        print("\n" + "=" * 70)
        print(f"  {nom} — formules des lignes concernées")
        print("=" * 70)
        try:
            montre_lignes(classeur.worksheet(nom), CIBLES_CDR)
        except gspread.WorksheetNotFound:
            print("  onglet introuvable")

    # ---------- 2. Les bilans ----------
    for nom in ("Bilan Année 1", "Bilan Année 2"):
        print("\n" + "=" * 70)
        print(f"  {nom} — actif immobilisé")
        print("=" * 70)
        try:
            feuille = classeur.worksheet(nom)
            formules = feuille.get_all_values(value_render_option="FORMULA")
            for i, ligne in enumerate(formules[:40], start=1):
                if ligne and ligne[0]:
                    valeur = ligne[1] if len(ligne) > 1 else ""
                    print(f"  {i:>3} | {ligne[0][:52]:<52} | {valeur}")
        except gspread.WorksheetNotFound:
            print("  onglet introuvable")

    # ---------- 3. Les onglets nouveaux de la v16 ----------
    for nom in ("Investissement de départ", "Projets", "Aménagements & travaux", "Mobilier"):
        print("\n" + "=" * 70)
        print(f"  {nom} — en-têtes et 6 premières lignes")
        print("=" * 70)
        try:
            feuille = classeur.worksheet(nom)
            formules = feuille.get_all_values(value_render_option="FORMULA")
            for i, ligne in enumerate(formules[:7], start=1):
                cellules = [c[:26] for c in ligne[:8]]
                print(f"  {i:>2} | " + " | ".join(cellules))
        except gspread.WorksheetNotFound:
            print("  onglet introuvable")


if __name__ == "__main__":
    main()
