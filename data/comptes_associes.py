"""
comptes_associes.py — prolonge les comptes d'Ayoub, Saïda et Karim jusqu'à
février 2027, ouvre celui d'Abdul, et relie le CdR aux salaires dus.

Tout part en une seule opération groupée : tout passe, ou rien.

1. « Associés-Salariés » — Ayoub (A:F) et Saïda (I:N)
   - 6 lignes insérées après août 2026 : 2026-09 à 2027-02, mêmes formules
     (salaire payé lu dans Caisse, Banque et Saisies CB Karim).
   - Septembre reprend les montants déjà tapés à la main sous le tableau
     (Ayoub 2 500 en A16:B16, Saïda 1 000 en I15:J15) ; ces deux restes
     sont effacés. Octobre à février : salaire dû vide, à remplir.
   - TOTAL d'Ayoub : compte désormais la ligne « repas » (100), qui était
     dans le solde cumulé mais hors du TOTAL.
2. « Associés-Salariés » — Abdul (P:U), même modèle
   - 2026-08 à 2027-02, salaire dû vide (à remplir), salaire payé lu sur
     la catégorie « Remboursement Abdul ».
3. « Associés » — Karim
   - 6 colonnes insérées après « Aout 26 » : Sep 26 à Fév 27, mêmes
     formules ; le TOTAL et le solde cumulé vont jusqu'à février.
4. « CdR Année 2 »
   - Ligne « Salaires — Abdul » insérée sous Saïda.
   - Ayoub, Saïda et Abdul, de septembre à février : chaque mois vaut
     « moins le salaire dû » du mois dans Associés-Salariés. Changer un
     salaire là-bas change le CdR.
5. « Bilan Année 2 » — arrêté à une date choisie
   - Case B3 « Arrêté au » (31/08/2026 au départ). Tout ce qui se calcule
     y est coupé : résultat du CdR, caisse, banque, écarts de caisse,
     immobilisations, amortissements, comptes de Karim, Ayoub, Saïda, Abdul.
   - Ligne « Dette salariale — Abdul » insérée sous Ayoub ; Abdul entre
     aussi dans les créances salariales.
   - Les soldes des quatre comptes à la date sont posés sous le Bilan
     (lignes 44 à 48), et le Bilan les lit.
   - Reste saisi à la main : charges payées d'avance (B15).
   Les formules écrites n'ont aucun séparateur d'arguments (ni « , » ni
   « ; ») : elles passent quelle que soit la langue du classeur.

USAGE (Windows PowerShell : séparer les commandes par « ; », pas « && »)
  python comptes_associes.py            → vérifie et AFFICHE, ne modifie rien
  python comptes_associes.py --ecrire   → applique, puis relit

Pas d'écriture entre 20h et 23h : synchro de la caisse à 21h, export du
CdR à 22h.
"""

import datetime
import os
import re
import sys

import gspread

# ---------- Réglages ----------

SHEET_ID = "1OUsQHnN9J241pLTb3pyrKmj7MTIrjCI6PNAw4FZ835U"

DOSSIER = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS = os.path.join(DOSSIER, "credentials.json")
JETON = os.path.join(DOSSIER, "authorized_user.json")

SALARIES = "Associés-Salariés"
KARIM = "Associés"
CDR = "CdR Année 2"
BILAN = "Bilan Année 2"

NOUVEAUX_MOIS = ["2026-09", "2026-10", "2026-11", "2026-12", "2027-01", "2027-02"]
ENTETES_KARIM = ["Sep 26", "Oct 26", "Nov 26", "Déc 26", "Jan 27", "Fév 27"]
MOIS_ABDUL = ["2026-08"] + NOUVEAUX_MOIS

LIBELLE_CDR_ABDUL = "Salaires — Abdul"  # même libellé que la catégorie de l'appli
COMMENTAIRE_ABDUL = "Avant le 19/08/2026, Abdul était payé en « Salaires - Autre » (déjà au CdR)."

# Lignes d'Associés-Salariés APRÈS insertion des 6 mois (14 à 19).
L_SEPT = 14            # 2026-09 ; 2027-02 en 19
L_HORS_CIRCUIT = 20    # était 14
L_REPAS_AYOUB = 21     # était 15
L_RESTE_AYOUB = 22     # était 16 (septembre tapé à la main)
L_RESTE_SAIDA = 21     # était 15 (septembre tapé à la main, daté 01/08)
L_TOTAL = 24           # était 18

# Abdul : historique 3, mois 4 à 10, hors circuit 11, TOTAL 24.
P, Q, R, U = 16, 17, 18, 21
L_HC_ABDUL = 11

# CdR : colonnes H (Sep 26) à M (Fév 27).
COLS_CDR = [8, 9, 10, 11, 12, 13]
L_CDR_AYOUB, L_CDR_SAIDA, L_CDR_ABDUL = 20, 21, 22
L_CDR_RESULTAT = 51    # « RÉSULTAT NET », était 50

# Bilan : date d'arrêté en B3, ligne Abdul insérée en 19 (lignes 19+ décalées).
DATE_ARRETE = datetime.date(2026, 8, 31)
L_BILAN_ABDUL = 19
L_SOLDES = 44          # titre ; Karim 45, Ayoub 46, Saïda 47, Abdul 48
L_CLES_CDR = 50        # « 2026-03 » … « 2027-02 » en B:M, pour couper le résultat
L_CLES_KARIM = 20      # dans Associés : « 2025-11 » … « 2027-02 » en B:Q
MOIS_KARIM = ["2025-11", "2025-12"] + [f"2026-{m:02d}" for m in range(1, 13)] + ["2027-01", "2027-02"]
MOIS_EXERCICE = [f"2026-{m:02d}" for m in range(3, 13)] + ["2027-01", "2027-02"]
JUSQUA = "<$B$3+1"     # « daté au plus tard le jour d'arrêté » (les dates ont une heure)


# ---------- Outils ----------

def col(n):
    """Numéro de colonne (1 = A) → lettre."""
    lettres = ""
    while n:
        n, r = divmod(n - 1, 26)
        lettres = chr(65 + r) + lettres
    return lettres


def plage(id_onglet, l1, c1, l2, c2):
    """GridRange de l'API, lignes et colonnes numérotées à partir de 1, bornes incluses."""
    return {"sheetId": id_onglet, "startRowIndex": l1 - 1, "endRowIndex": l2,
            "startColumnIndex": c1 - 1, "endColumnIndex": c2}


def ecrire(id_onglet, ligne, colonne, valeur):
    """Écrit une case : texte (jamais interprété), nombre, formule, ou vide (None)."""
    if valeur is None:
        contenu = {}
    elif isinstance(valeur, (int, float)):
        contenu = {"numberValue": valeur}
    elif valeur.startswith("="):
        # Formules sans séparateur d'arguments : indépendantes de la langue du classeur.
        contenu = {"formulaValue": valeur}
    else:
        contenu = {"stringValue": valeur}
    return {"updateCells": {"range": plage(id_onglet, ligne, colonne, ligne, colonne),
                            "rows": [{"values": [{"userEnteredValue": contenu}]}],
                            "fields": "userEnteredValue"}}


def vider(id_onglet, l1, c1, l2, c2):
    """Efface les valeurs d'une plage, garde la mise en forme."""
    return {"updateCells": {"range": plage(id_onglet, l1, c1, l2, c2), "fields": "userEnteredValue"}}


def copier(id_onglet, source, destination):
    return {"copyPaste": {"source": plage(id_onglet, *source),
                          "destination": plage(id_onglet, *destination),
                          "pasteType": "PASTE_NORMAL"}}


def remplacer(id_onglet, zone, avant, apres):
    return {"findReplace": {"find": avant, "replacement": apres, "range": plage(id_onglet, *zone),
                            "matchCase": True, "includeFormulas": True}}


def largeurs(classeur, onglet, a1):
    meta = classeur.fetch_sheet_metadata({
        "includeGridData": False, "fields": "sheets(data.columnMetadata.pixelSize)",
        "ranges": [f"'{onglet}'!{a1}"]})
    return [m.get("pixelSize", 100) for m in meta["sheets"][0]["data"][0].get("columnMetadata", [])]


def largeur(id_onglet, colonne, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": id_onglet, "dimension": "COLUMNS",
                  "startIndex": colonne - 1, "endIndex": colonne},
        "properties": {"pixelSize": px}, "fields": "pixelSize"}}


def arreter(message):
    print()
    print("ARRÊT :", message)
    print("Rien n'a été modifié.")
    sys.exit(1)


def grille(feuille, a1, lignes, colonnes, rendu="FORMULA"):
    """Valeurs d'une plage, complétées pour avoir toujours lignes × colonnes cases."""
    v = feuille.get(a1, value_render_option=rendu)
    v = [list(l) + [""] * (colonnes - len(l)) for l in v]
    return v + [[""] * colonnes for _ in range(lignes - len(v))]


def attendre(g, coord, texte, onglet):
    l, c = coord
    if str(g[l - 1][c - 1]).strip() != texte:
        arreter(f"{onglet}!{col(c)}{l} vaut « {g[l - 1][c - 1]} », attendu « {texte} ». "
                "Le classeur a changé depuis le téléchargement du 02/10 : à revoir avant d'écrire.")


def nombre(x):
    try:
        return float(str(x).replace(",", ".").replace(" ", "").replace(" ", ""))
    except ValueError:
        return None


# ---------- Vérifications (lecture seule) ----------

def verifier(classeur):
    sal = classeur.worksheet(SALARIES)
    kar = classeur.worksheet(KARIM)
    cdr = classeur.worksheet(CDR)

    # Associés-Salariés, tel que téléchargé le 02/10.
    g = grille(sal, "A1:U30", 30, 21)
    for coord, texte in {(2, 1): "Ayoub", (2, 9): "Saida", (3, 1): "historique",
                         (13, 1): "2026-08", (13, 9): "2026-08", (14, 1): "hors circuit",
                         (14, 9): "hors circuit", (15, 1): "repas", (18, 1): "TOTAL",
                         (18, 9): "TOTAL", (18, 5): "=E14"}.items():
        attendre(g, coord, texte, SALARIES)
    if "Remboursement Ayoub" not in str(g[12][2]) or "Remboursement Saida" not in str(g[12][10]):
        arreter(f"{SALARIES}, ligne 13 : formule de salaire payé inattendue.")
    sept_ayoub, sept_saida = nombre(g[15][1]), nombre(g[14][9])
    if sept_ayoub is None or sept_saida is None:
        arreter(f"{SALARIES}!B16 ou J15 (septembre tapé à la main) n'est pas un nombre : "
                f"« {g[15][1]} », « {g[14][9]} ».")
    if any(str(x).strip() for x in g[15][2:6] + g[14][10:14]):
        arreter(f"{SALARIES}, lignes 15-16 : il y a plus que la date et le montant de septembre.")
    occupees = [f"{col(c)}{l}" for l in range(1, 31) for c in range(1, 22)
                if str(g[l - 1][c - 1]).strip() and (l >= 19 or l == 17 or c >= P)]
    if occupees:
        arreter(f"{SALARIES} : cases censées être vides ({', '.join(occupees[:8])}).")

    # Associés (Karim).
    k = grille(kar, "A1:Z25", 25, 26)
    if "KARIM" not in str(k[0][0]):
        arreter(f"{KARIM}!A1 vaut « {k[0][0]} ».")
    attendre(k, (3, 11), "Aout 26", KARIM)
    attendre(k, (3, 12), "TOTAL", KARIM)
    attendre(k, (19, 12), "=K19", KARIM)
    if str(k[5][10]).count("2026-08") != 2:
        arreter(f"{KARIM}!K6 : la formule d'août ne contient pas deux fois « 2026-08 ».")
    occupees = [f"{col(c)}{l}" for l in range(1, 26) for c in range(13, 27) if str(k[l - 1][c - 1]).strip()]
    if occupees:
        arreter(f"{KARIM} : colonnes M et suivantes pas vides ({', '.join(occupees[:8])}).")

    # CdR Année 2.
    r = grille(cdr, "A1:O24", 24, 15)
    for coord, texte in {(2, 8): "Sep 26", (2, 13): "Fév 27", (20, 1): "Salaires — Ayoub",
                         (21, 1): "Salaires — Saïda", (22, 1): "Salaire femme de ménage"}.items():
        attendre(r, coord, texte, CDR)
    if any(LIBELLE_CDR_ABDUL in str(l[0]) for l in cdr.get("A1:A80")):
        arreter(f"« {LIBELLE_CDR_ABDUL} » existe déjà dans {CDR}.")

    attendre(r, (2, 2), "Mar 26", CDR)
    if str(cdr.acell(f"A{L_CDR_RESULTAT - 1}").value).strip() != "RÉSULTAT NET":
        arreter(f"{CDR}!A{L_CDR_RESULTAT - 1} n'est pas « RÉSULTAT NET ».")

    categories = [str(v[0]).strip() for v in classeur.worksheet("Index").get("A63:A100") if v]
    if "Remboursement Abdul" not in categories:
        arreter("« Remboursement Abdul » n'est pas dans la liste des catégories (Index!A63…).")

    # Associés : la ligne 20 recevra les clés de mois.
    if any(str(x).strip() for x in k[L_CLES_KARIM - 1]):
        arreter(f"{KARIM}, ligne {L_CLES_KARIM} : pas vide.")

    # Bilan Année 2.
    bil = classeur.worksheet(BILAN)
    b = grille(bil, "A1:M55", 55, 13)
    for coord, texte in {(9, 1): "Mobilier & équipement", (10, 1): "Aménagements & travaux",
                         (12, 1): "Lemon Tree Riad & Hostel", (13, 1): "Amortissements cumulés",
                         (15, 1): "Charges payées d'avance", (16, 1): "Créances sur associés",
                         (17, 1): "Créances salariales", (18, 1): "Compte d'attente — écarts de caisse",
                         (19, 1): "TRÉSORERIE", (20, 1): "Caisse", (21, 1): "Banque",
                         (22, 1): "TOTAL ACTIF", (10, 5): "Résultat de l'exercice",
                         (14, 5): "Compte courant Karim", (18, 5): "Dette salariale — Ayoub",
                         (19, 5): "Dette salariale — Saïda", (13, 2): "=-'Échéancier'!$M$64",
                         (10, 6): f"='{CDR}'!O50", (20, 2): "=7881+SUM(Caisse!$D$2:$D$5233)"}.items():
        attendre(b, coord, texte, BILAN)
    for coord, morceau in {(16, 2): "'Associés'!$K$19", (14, 6): "'Associés'!$K$19",
                           (17, 2): "'Associés-Salariés'!$E$14", (18, 6): "'Associés-Salariés'!$E$14",
                           (19, 6): "'Associés-Salariés'!$M$14"}.items():
        if morceau not in str(b[coord[0] - 1][coord[1] - 1]):
            arreter(f"{BILAN}!{col(coord[1])}{coord[0]} ne lit pas {morceau}.")
    vides = [f"{col(c)}{l}" for l in [3] + list(range(42, 56)) for c in range(1, 14)
             if str(b[l - 1][c - 1]).strip()]
    if vides:
        arreter(f"{BILAN} : cases censées être vides ({', '.join(vides[:8])}).")
    # Les formules coupées au 01/03/2027 passent à la date d'arrêté.
    coupees = {}
    for l in (9, 10, 12, 18, 21):
        f = str(b[l - 1][1])
        nouvelle = re.sub(r"DATE\(2027[,;]\s*3[,;]\s*1\)", "$B$3+1", f)
        if nouvelle == f:
            arreter(f"{BILAN}!B{l} ne contient pas DATE(2027,3,1).")
        if re.search(r"[,;]", re.sub(r'"[^"]*"', "", nouvelle)):
            arreter(f"{BILAN}!B{l} contient d'autres séparateurs : formule à revoir à la main.")
        coupees[l] = nouvelle

    actuel = {"Ayoub": r[19][7:13], "Saïda": r[20][7:13]}
    return sal, kar, cdr, bil, coupees, sept_ayoub, sept_saida, actuel


# ---------- Les changements ----------

def requetes(classeur, sal, kar, cdr, bil, coupees, sept_ayoub, sept_saida):
    s, k, c, b = sal.id, kar.id, cdr.id, bil.id
    reqs = []

    # 1. Ayoub et Saïda : 6 mois de plus après août.
    reqs.append({"insertDimension": {"range": {"sheetId": s, "dimension": "ROWS",
                                               "startIndex": L_SEPT - 1, "endIndex": L_SEPT + 5},
                                     "inheritFromBefore": True}})
    reqs.append(copier(s, (13, 1, 13, 14), (L_SEPT, 1, L_SEPT + 5, 14)))
    for i, mois in enumerate(NOUVEAUX_MOIS):
        for colonne_mois, colonne_du in ((1, 2), (9, 10)):
            reqs.append(ecrire(s, L_SEPT + i, colonne_mois, mois))
            reqs.append(ecrire(s, L_SEPT + i, colonne_du, None))
    reqs.append(ecrire(s, L_SEPT, 2, sept_ayoub))
    reqs.append(ecrire(s, L_SEPT, 10, sept_saida))
    reqs.append(vider(s, L_RESTE_AYOUB, 1, L_RESTE_AYOUB, 6))
    reqs.append(vider(s, L_RESTE_SAIDA, 9, L_RESTE_SAIDA, 14))
    # TOTAL d'Ayoub : jusqu'à la ligne repas, qui suit déjà le solde cumulé.
    reqs += [ecrire(s, L_TOTAL, 2, f"=SUM(B3:B{L_REPAS_AYOUB})"),
             ecrire(s, L_TOTAL, 3, f"=SUM(C3:C{L_REPAS_AYOUB})"),
             ecrire(s, L_TOTAL, 5, f"=E{L_REPAS_AYOUB}")]

    # 2. Abdul, en P:U, sur le modèle d'Ayoub.
    reqs.append(copier(s, (1, 1, 3, 6), (1, P, 3, U)))
    reqs.append(copier(s, (13, 1, 13, 6), (4, P, 3 + len(MOIS_ABDUL), U)))
    reqs.append(copier(s, (L_HORS_CIRCUIT, 1, L_HORS_CIRCUIT, 6), (L_HC_ABDUL, P, L_HC_ABDUL, U)))
    reqs.append(copier(s, (L_TOTAL, 1, L_TOTAL, 6), (L_TOTAL, P, L_TOTAL, U)))
    zone = (1, P, L_TOTAL, U)
    reqs.append(remplacer(s, zone, "Remboursement Ayoub", "Remboursement Abdul"))
    reqs.append(remplacer(s, zone, "$A", "$P"))
    reqs += [ecrire(s, 2, P, "Abdul"),
             ecrire(s, 3, Q, 0), ecrire(s, 3, R, 0), ecrire(s, 3, U, COMMENTAIRE_ABDUL)]
    for i, mois in enumerate(MOIS_ABDUL):
        reqs += [ecrire(s, 4 + i, P, mois), ecrire(s, 4 + i, Q, None)]
    reqs += [ecrire(s, L_HC_ABDUL, Q, 0), ecrire(s, L_HC_ABDUL, R, 0), ecrire(s, L_HC_ABDUL, U, None),
             ecrire(s, L_TOTAL, Q, f"=SUM(Q3:Q{L_HC_ABDUL})"),
             ecrire(s, L_TOTAL, R, f"=SUM(R3:R{L_HC_ABDUL})"),
             ecrire(s, L_TOTAL, 20, f"=T{L_HC_ABDUL}")]
    for i, px in enumerate(largeurs(classeur, SALARIES, "A1:F1")[:6]):
        reqs.append(largeur(s, P + i, px))

    # 3. Karim : Sep 26 à Fév 27, insérés entre « Aout 26 » (K) et TOTAL.
    reqs.append({"insertDimension": {"range": {"sheetId": k, "dimension": "COLUMNS",
                                               "startIndex": 11, "endIndex": 17},
                                     "inheritFromBefore": True}})
    reqs.append(copier(k, (3, 11, 19, 11), (3, 12, 19, 17)))
    for i, mois in enumerate(NOUVEAUX_MOIS):
        reqs.append(remplacer(k, (4, 12 + i, 19, 12 + i), "2026-08", mois))
        reqs.append(ecrire(k, 3, 12 + i, ENTETES_KARIM[i]))
    # TOTAL (désormais en R) : SUM(B:K) → SUM(B:Q), et solde = Q19.
    reqs.append(remplacer(k, (4, 18, 19, 18), "K", "Q"))
    px_k = largeurs(classeur, KARIM, "K1:K1")
    for i in range(6):
        reqs.append(largeur(k, 12 + i, px_k[0] if px_k else 100))

    # 4. CdR : ligne Abdul sous Saïda, puis les trois lignes reliées.
    reqs.append({"insertDimension": {"range": {"sheetId": c, "dimension": "ROWS",
                                               "startIndex": L_CDR_ABDUL - 1, "endIndex": L_CDR_ABDUL},
                                     "inheritFromBefore": True}})
    reqs.append(copier(c, (L_CDR_SAIDA, 1, L_CDR_SAIDA, 15), (L_CDR_ABDUL, 1, L_CDR_ABDUL, 15)))
    reqs.append(ecrire(c, L_CDR_ABDUL, 1, LIBELLE_CDR_ABDUL))
    for colonne in range(2, 8):  # mars à août 26 : rien pour Abdul
        reqs.append(ecrire(c, L_CDR_ABDUL, colonne, None))
    for i, colonne in enumerate(COLS_CDR):
        reqs.append(ecrire(c, L_CDR_AYOUB, colonne, f"=-'{SALARIES}'!B{L_SEPT + i}"))
        reqs.append(ecrire(c, L_CDR_SAIDA, colonne, f"=-'{SALARIES}'!J{L_SEPT + i}"))
        reqs.append(ecrire(c, L_CDR_ABDUL, colonne, f"=-'{SALARIES}'!Q{5 + i}"))

    # 5. Bilan, arrêté à la date de B3. Les lignes 19+ descendent d'un cran.
    reqs.append(ecrire(k, L_CLES_KARIM, 1, "Mois (clé pour le Bilan)"))
    for i, mois in enumerate(MOIS_KARIM):
        reqs.append(ecrire(k, L_CLES_KARIM, 2 + i, mois))

    reqs.append({"insertDimension": {"range": {"sheetId": b, "dimension": "ROWS",
                                               "startIndex": L_BILAN_ABDUL - 1, "endIndex": L_BILAN_ABDUL},
                                     "inheritFromBefore": True}})
    serie = (DATE_ARRETE - datetime.date(1899, 12, 30)).days
    reqs.append({"updateCells": {
        "range": plage(b, 3, 2, 3, 2),
        "rows": [{"values": [{"userEnteredValue": {"numberValue": serie},
                              "userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"},
                                                    "textFormat": {"bold": True}}}]}],
        "fields": "userEnteredValue,userEnteredFormat.numberFormat,userEnteredFormat.textFormat.bold"}})
    reqs += [
        ecrire(b, 2, 1, "Situation à la date d'arrêté (B3) — exercice en cours du 01/03/2026 au 28/02/2027"),
        ecrire(b, 3, 1, "Arrêté au :"),
        ecrire(b, 3, 3, "← changer la date (fin de mois). À remettre à jour à la main : "
                        "charges payées d'avance (B15)."),
    ]

    # Soldes des quatre comptes à la date, sous le Bilan.
    S, K = f"'{SALARIES}'!", f"'{KARIM}'!"
    soldes = [
        ("Karim — compte courant (> 0 : l'hôtel lui doit)",
         f"={K}$B$4+SUMPRODUCT((DATEVALUE({K}$B${L_CLES_KARIM}:$Q${L_CLES_KARIM}&\"-01\"){JUSQUA})"
         f"*{K}$B$18:$Q$18)",
         "Associés : report initial + flux des mois jusqu'à la date"),
        ("Ayoub — salaire (> 0 : reste dû)",
         f"={S}$E$3+SUMPRODUCT((DATEVALUE({S}$A$4:$A$19&\"-01\"){JUSQUA})*{S}$D$4:$D$19)"
         f"+{S}$D${L_HORS_CIRCUIT}+{S}$D${L_REPAS_AYOUB}",
         "Associés-Salariés : historique + mois jusqu'à la date + hors circuit et repas"),
        ("Saïda — salaire (> 0 : reste dû)",
         f"={S}$M$3+SUMPRODUCT((DATEVALUE({S}$I$4:$I$19&\"-01\"){JUSQUA})*{S}$L$4:$L$19)"
         f"+{S}$L${L_HORS_CIRCUIT}",
         "Associés-Salariés : historique + mois jusqu'à la date + hors circuit"),
        ("Abdul — salaire (> 0 : reste dû)",
         f"={S}$T$3+SUMPRODUCT((DATEVALUE({S}$P$4:$P$10&\"-01\"){JUSQUA})*{S}$S$4:$S$10)"
         f"+{S}$S${L_HC_ABDUL}",
         "Associés-Salariés : mois jusqu'à la date + hors circuit"),
    ]
    reqs.append(ecrire(b, L_SOLDES, 1, "SOLDES DES COMPTES À LA DATE D'ARRÊTÉ — calculés, ne pas saisir"))
    for i, (nom, formule, source) in enumerate(soldes):
        reqs += [ecrire(b, L_SOLDES + 1 + i, 1, nom), ecrire(b, L_SOLDES + 1 + i, 2, formule),
                 ecrire(b, L_SOLDES + 1 + i, 3, source)]
    karim, ayoub, saida, abdul = (f"$B${L_SOLDES + 1 + i}" for i in range(4))

    reqs.append(ecrire(b, L_CLES_CDR, 1, "Mois de l'exercice (clé pour couper le résultat du CdR)"))
    for i, mois in enumerate(MOIS_EXERCICE):
        reqs.append(ecrire(b, L_CLES_CDR, 2 + i, mois))

    # Actif.
    for l, formule in coupees.items():
        reqs.append(ecrire(b, l + (1 if l >= L_BILAN_ABDUL else 0), 2, formule))
    reqs += [
        ecrire(b, 13, 2, "=-('Échéancier'!$I$64+SUMPRODUCT(('Échéancier'!$AC$4:$AN$4<$B$3+1)"
                         "*'Échéancier'!$AC$64:$AN$64))"),
        ecrire(b, 13, 4, "Échéancier : cumul au 28/02/2026 + dotations mensuelles jusqu'à la date"),
        ecrire(b, 16, 2, f"=-({karim}<0)*{karim}"),
        ecrire(b, 16, 4, f"Solde Karim ({karim[1:].replace('$', '')}) s'il est négatif"),
        ecrire(b, 17, 2, f"=-({ayoub}<0)*{ayoub}-({saida}<0)*{saida}-({abdul}<0)*{abdul}"),
        ecrire(b, 21, 2, "=7881+SUMPRODUCT((Caisse!$B$2:$B$5233<$B$3+1)*Caisse!$D$2:$D$5233)"),
        ecrire(b, 21, 4, "ouverture 7 881 + journal de caisse jusqu'à la date"),
        ecrire(b, 22, 4, "Banque, mouvements jusqu'à la date"),
    ]
    # Passif.
    reqs += [
        ecrire(b, 10, 6, f"=SUMPRODUCT((DATEVALUE($B${L_CLES_CDR}:$M${L_CLES_CDR}&\"-01\"){JUSQUA})"
                         f"*'{CDR}'!$B${L_CDR_RESULTAT}:$M${L_CDR_RESULTAT})"),
        ecrire(b, 10, 8, f"{CDR}, résultat net de mars à la date"),
        ecrire(b, 14, 6, f"=({karim}>0)*{karim}"),
        ecrire(b, 14, 8, "Solde Karim à la date (ligne 45)"),
        ecrire(b, 18, 6, f"=({ayoub}>0)*{ayoub}"),
        ecrire(b, 18, 8, "Solde Ayoub à la date (ligne 46)"),
        ecrire(b, L_BILAN_ABDUL, 5, "Dette salariale — Abdul"),
        ecrire(b, L_BILAN_ABDUL, 6, f"=({abdul}>0)*{abdul}"),
        ecrire(b, L_BILAN_ABDUL, 7, 0),
        ecrire(b, L_BILAN_ABDUL, 8, "Solde Abdul à la date (ligne 48) — compte ouvert le 19/08/2026"),
        ecrire(b, 20, 6, f"=({saida}>0)*{saida}"),
        ecrire(b, 20, 8, "Solde Saïda à la date (ligne 47)"),
    ]

    return reqs


def afficher_plan(sept_ayoub, sept_saida, actuel):
    print()
    print(f"{SALARIES} — Ayoub et Saïda :")
    print(f"  lignes {L_SEPT} à {L_SEPT + 5} insérées : {NOUVEAUX_MOIS[0]} → {NOUVEAUX_MOIS[-1]}")
    print(f"  septembre dû : Ayoub {sept_ayoub:g} (repris de B16), Saïda {sept_saida:g} (repris de J15)")
    print("  octobre → février : salaire dû vide, à remplir")
    print(f"  restes effacés : A16:B16 (Ayoub) et I15:J15 (Saïda)")
    print(f"  TOTAL d'Ayoub : compte la ligne repas")
    print(f"{SALARIES} — Abdul en P:U : {MOIS_ABDUL[0]} → {MOIS_ABDUL[-1]}, salaire dû vide,")
    print("  salaire payé = « Remboursement Abdul »")
    print(f"{KARIM} — colonnes {', '.join(ENTETES_KARIM)} insérées avant TOTAL")
    print(f"{CDR} — ligne 22 « {LIBELLE_CDR_ABDUL} » ; de Sep 26 à Fév 27 :")
    print(f"  Ayoub (20) = -B{L_SEPT}…B{L_SEPT + 5}, aujourd'hui : {actuel['Ayoub']}")
    print(f"  Saïda (21) = -J{L_SEPT}…J{L_SEPT + 5}, aujourd'hui : {actuel['Saïda']}")
    print(f"  Abdul (22) = -Q5…Q10")
    print(f"{BILAN} — arrêté au {DATE_ARRETE:%d/%m/%Y} (case B3, à changer) :")
    print("  coupés à la date : résultat, caisse, banque, écarts de caisse, immobilisations,")
    print("  amortissements, comptes Karim / Ayoub / Saïda / Abdul (soldes en lignes 45 à 48)")
    print(f"  ligne {L_BILAN_ABDUL} insérée : « Dette salariale — Abdul »")
    print("  reste à la main : charges payées d'avance (B15)")


def relire(classeur):
    sal = classeur.worksheet(SALARIES)
    kar = classeur.worksheet(KARIM)
    cdr = classeur.worksheet(CDR)
    print()
    print("Relecture — Associés-Salariés, lignes 12 à 24 (A:E · I:M · P:T) :")
    for l in grille(sal, "A12:T24", 13, 20, rendu="FORMATTED_VALUE"):
        if any(l):
            print("  ", " | ".join(l[0:5]), " ‖ ", " | ".join(l[8:13]), " ‖ ", " | ".join(l[15:20]))
    print()
    print("Relecture — Karim, en-têtes et solde cumulé :")
    k = grille(kar, "J3:R19", 17, 9, rendu="FORMATTED_VALUE")
    print("  ", " | ".join(k[0]))
    print("  ", " | ".join(k[16]))
    print()
    print("Relecture — CdR, Sep 26 :")
    for l in grille(cdr, "A19:H24", 6, 8, rendu="FORMATTED_VALUE"):
        print(f"   {l[0]} → {l[7]}")
    print("   TOTAL CHARGES FIXES (H32) :", cdr.get("H32", value_render_option="FORMULA")[0][0])
    print()
    print("Relecture — Bilan :")
    bil = classeur.worksheet(BILAN)
    for l in grille(bil, "A3:F26", 24, 6, rendu="FORMATTED_VALUE"):
        if any(l):
            print(f"   {l[0][:38]:<38} {l[1]:>14}   {l[4][:30]:<30} {l[5]:>14}")
    for l in grille(bil, "A45:B48", 4, 2, rendu="FORMATTED_VALUE"):
        print(f"   {l[0][:50]:<50} {l[1]:>12}")
    erreurs = [f"{BILAN}!{c}" for c, v in zip(
        ("B3", "B13", "B16", "B17", "B21", "B22", "F10", "F14", "F18", "F19", "F20"),
        (bil.acell(c).value for c in ("B3", "B13", "B16", "B17", "B21", "B22",
                                      "F10", "F14", "F18", "F19", "F20")))
        if str(v).startswith("#")]
    erreurs += [f"{BILAN}!B{l}" for l, v in zip(range(45, 49), grille(bil, "B45:B48", 4, 1))
                if str(v[0]).startswith("#")]
    print()
    print("Cases en erreur :", ", ".join(erreurs) if erreurs else "aucune.")


# ---------- Programme ----------

def main():
    ecrit = "--ecrire" in sys.argv

    if ecrit and 20 <= datetime.datetime.now().hour < 23:
        arreter("entre 20h et 23h, la synchro de la caisse et l'export du CdR tournent. Relancer plus tard.")
    if not os.path.exists(CREDENTIALS):
        arreter(f"credentials.json introuvable dans {DOSSIER}.")

    print("Connexion à Google...")
    client = gspread.oauth(credentials_filename=CREDENTIALS, authorized_user_filename=JETON)
    classeur = client.open_by_key(SHEET_ID)
    print("Classeur :", classeur.title)

    sal, kar, cdr, bil, coupees, sept_ayoub, sept_saida, actuel = verifier(classeur)
    print("Vérifications : OK.")
    afficher_plan(sept_ayoub, sept_saida, actuel)

    if not ecrit:
        print()
        print("Simulation seulement : rien n'a été modifié.")
        print("Pour appliquer : python comptes_associes.py --ecrire")
        return

    classeur.batch_update({"requests": requetes(classeur, sal, kar, cdr, bil, coupees,
                                                sept_ayoub, sept_saida)})
    print()
    print("Écrit.")
    relire(classeur)


if __name__ == "__main__":
    main()
