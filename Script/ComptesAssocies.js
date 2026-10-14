/**
 * ComptesAssocies.js — portage Apps Script de `data/comptes_associes.py`.
 *
 * Même travail, mêmes vérifications, mais lancé depuis le classeur : pas de
 * connexion Google à régler, le script tourne avec le compte de Karim.
 *
 * Prolonge les comptes d'Ayoub, Saïda et Karim jusqu'à février 2027, ouvre
 * celui d'Abdul, relie le CdR aux salaires dus (ligne « Salaires — Abdul »
 * ajoutée sous Saïda) et arrête le Bilan Année 2 à une date choisie (B3).
 * Le détail de chaque changement est dans l'en-tête du script Python.
 *
 * Tout part en une seule opération groupée (API Sheets, service avancé
 * « Sheets ») : tout passe, ou rien.
 *
 * USAGE, depuis l'éditeur Apps Script (menu déroulant des fonctions) :
 *   comptesAssociesSimuler  → vérifie et AFFICHE dans le journal, ne modifie rien
 *   comptesAssociesEcrire   → applique, puis relit
 *
 * À usage unique : une fois écrit, une relance s'arrête d'elle-même
 * (« Salaires — Abdul » existe déjà dans le CdR).
 */

function comptesAssociesSimuler() {
  ComptesAssocies_.lancer(false);
}

function comptesAssociesEcrire() {
  ComptesAssocies_.lancer(true);
}

// Tout le reste est enfermé ici : les fichiers d'un projet Apps Script
// partagent leurs noms globaux, et Code.js a déjà un `ecrire`.
var ComptesAssocies_ = (function () {
  // ---------- Réglages ----------

  var SALARIES = "Associés-Salariés";
  var KARIM = "Associés";
  var CDR = "CdR Année 2";
  var BILAN = "Bilan Année 2";

  var NOUVEAUX_MOIS = ["2026-09", "2026-10", "2026-11", "2026-12", "2027-01", "2027-02"];
  var ENTETES_KARIM = ["Sep 26", "Oct 26", "Nov 26", "Déc 26", "Jan 27", "Fév 27"];
  var MOIS_ABDUL = ["2026-08"].concat(NOUVEAUX_MOIS);

  var LIBELLE_CDR_ABDUL = "Salaires — Abdul"; // même libellé que la catégorie de l'appli
  var COMMENTAIRE_ABDUL = "Avant le 19/08/2026, Abdul était payé en « Salaires - Autre » (déjà au CdR).";

  // Lignes d'Associés-Salariés APRÈS insertion des 6 mois (14 à 19).
  var L_SEPT = 14; // 2026-09 ; 2027-02 en 19
  var L_HORS_CIRCUIT = 20; // était 14
  var L_REPAS_AYOUB = 21; // était 15
  var L_RESTE_AYOUB = 22; // était 16 (septembre tapé à la main)
  var L_RESTE_SAIDA = 21; // était 15 (septembre tapé à la main, daté 01/08)
  var L_TOTAL = 24; // était 18

  // Abdul : historique 3, mois 4 à 10, hors circuit 11, TOTAL 24.
  var P = 16, Q = 17, R = 18, U = 21;
  var L_HC_ABDUL = 11;

  // CdR : colonnes H (Sep 26) à M (Fév 27).
  var COLS_CDR = [8, 9, 10, 11, 12, 13];
  var L_CDR_AYOUB = 20, L_CDR_SAIDA = 21, L_CDR_ABDUL = 22;
  var L_CDR_RESULTAT = 51; // « RÉSULTAT NET », était 50

  // Bilan : date d'arrêté en B3, ligne Abdul insérée en 19 (lignes 19+ décalées).
  var DATE_ARRETE = { annee: 2026, mois: 8, jour: 31 };
  var L_BILAN_ABDUL = 19;
  var L_SOLDES = 44; // titre ; Karim 45, Ayoub 46, Saïda 47, Abdul 48
  var L_CLES_CDR = 50; // « 2026-03 » … « 2027-02 » en B:M, pour couper le résultat
  var L_CLES_KARIM = 20; // dans Associés : « 2025-11 » … « 2027-02 » en B:Q
  var MOIS_KARIM = ["2025-11", "2025-12"].concat(moisDe_(2026, 1, 12), ["2027-01", "2027-02"]);
  var MOIS_EXERCICE = moisDe_(2026, 3, 12).concat(["2027-01", "2027-02"]);
  var JUSQUA = "<$B$3+1"; // « daté au plus tard le jour d'arrêté » (les dates ont une heure)

  var sortie = [];

  // ---------- Outils ----------

  function moisDe_(annee, de, a) {
    var r = [];
    for (var m = de; m <= a; m++) r.push(annee + "-" + (m < 10 ? "0" : "") + m);
    return r;
  }

  function dire(texte) {
    sortie.push(texte === undefined ? "" : texte);
  }

  function arreter(message) {
    throw new Error("ARRÊT : " + message + " Rien n'a été modifié.");
  }

  /** Numéro de colonne (1 = A) → lettre. */
  function col(n) {
    var lettres = "";
    while (n) {
      var r = (n - 1) % 26;
      lettres = String.fromCharCode(65 + r) + lettres;
      n = Math.floor((n - 1) / 26);
    }
    return lettres;
  }

  function str(x) {
    return x === null || x === undefined ? "" : String(x);
  }

  /**
   * Une plage, lignes et colonnes à partir de 1, bornée à la taille de
   * l'onglet et complétée de cases vides : `f` donne la formule, ou la
   * valeur brute s'il n'y en a pas ; `t` le texte affiché.
   */
  function grille(feuille, l1, c1, nl, nc) {
    var f = [], t = [];
    var hl = Math.min(nl, feuille.getMaxRows() - l1 + 1);
    var hc = Math.min(nc, feuille.getMaxColumns() - c1 + 1);
    var formules = [], valeurs = [], textes = [];
    if (hl > 0 && hc > 0) {
      var plage = feuille.getRange(l1, c1, hl, hc);
      formules = plage.getFormulas();
      valeurs = plage.getValues();
      textes = plage.getDisplayValues();
    }
    for (var i = 0; i < nl; i++) {
      f.push([]);
      t.push([]);
      for (var j = 0; j < nc; j++) {
        var dedans = i < hl && j < hc;
        f[i].push(dedans ? formules[i][j] || valeurs[i][j] : "");
        t[i].push(dedans ? textes[i][j] : "");
      }
    }
    return { f: f, t: t };
  }

  function attendre(g, l, c, texte, onglet) {
    var brut = str(g.f[l - 1][c - 1]).trim();
    var affiche = str(g.t[l - 1][c - 1]).trim();
    if (brut !== texte && affiche !== texte) {
      arreter(onglet + "!" + col(c) + l + " vaut « " + (brut || affiche) + " », attendu « " + texte + " ». " +
        "Le classeur a changé depuis le téléchargement du 02/10 : à revoir avant d'écrire.");
    }
  }

  function nombre(x) {
    if (typeof x === "number") return x;
    var s = str(x).replace(",", ".").replace(/[\s  ]/g, "");
    if (!s) return null;
    var n = Number(s);
    return isNaN(n) ? null : n;
  }

  /** GridRange de l'API, lignes et colonnes numérotées à partir de 1, bornes incluses. */
  function plage(id, l1, c1, l2, c2) {
    return { sheetId: id, startRowIndex: l1 - 1, endRowIndex: l2, startColumnIndex: c1 - 1, endColumnIndex: c2 };
  }

  /** Écrit une case : texte (jamais interprété), nombre, formule, ou vide (null). */
  function ecrire(id, ligne, colonne, valeur) {
    var contenu;
    if (valeur === null) contenu = {};
    else if (typeof valeur === "number") contenu = { numberValue: valeur };
    else if (valeur.charAt(0) === "=") contenu = { formulaValue: valeur }; // sans séparateur d'arguments
    else contenu = { stringValue: valeur };
    return { updateCells: { range: plage(id, ligne, colonne, ligne, colonne),
      rows: [{ values: [{ userEnteredValue: contenu }] }], fields: "userEnteredValue" } };
  }

  /** Efface les valeurs d'une plage, garde la mise en forme. */
  function vider(id, l1, c1, l2, c2) {
    return { updateCells: { range: plage(id, l1, c1, l2, c2), fields: "userEnteredValue" } };
  }

  function copier(id, source, destination) {
    return { copyPaste: { source: plage.apply(null, [id].concat(source)),
      destination: plage.apply(null, [id].concat(destination)), pasteType: "PASTE_NORMAL" } };
  }

  function remplacer(id, zone, avant, apres) {
    return { findReplace: { find: avant, replacement: apres, range: plage.apply(null, [id].concat(zone)),
      matchCase: true, includeFormulas: true } };
  }

  function inserer(id, dimension, debut, fin) {
    return { insertDimension: { range: { sheetId: id, dimension: dimension, startIndex: debut, endIndex: fin },
      inheritFromBefore: true } };
  }

  function largeur(id, colonne, px) {
    return { updateDimensionProperties: {
      range: { sheetId: id, dimension: "COLUMNS", startIndex: colonne - 1, endIndex: colonne },
      properties: { pixelSize: px }, fields: "pixelSize" } };
  }

  function onglet(classeur, nom) {
    var f = classeur.getSheetByName(nom);
    if (!f) arreter("onglet « " + nom + " » introuvable.");
    return f;
  }

  // ---------- Vérifications (lecture seule) ----------

  function verifier(classeur) {
    var sal = onglet(classeur, SALARIES);
    var kar = onglet(classeur, KARIM);
    var cdr = onglet(classeur, CDR);

    // Associés-Salariés, tel que téléchargé le 02/10.
    var g = grille(sal, 1, 1, 30, 21);
    [[2, 1, "Ayoub"], [2, 9, "Saida"], [3, 1, "historique"], [13, 1, "2026-08"], [13, 9, "2026-08"],
      [14, 1, "hors circuit"], [14, 9, "hors circuit"], [15, 1, "repas"], [18, 1, "TOTAL"],
      [18, 9, "TOTAL"], [18, 5, "=E14"]].forEach(function (a) { attendre(g, a[0], a[1], a[2], SALARIES); });
    if (str(g.f[12][2]).indexOf("Remboursement Ayoub") < 0 || str(g.f[12][10]).indexOf("Remboursement Saida") < 0) {
      arreter(SALARIES + ", ligne 13 : formule de salaire payé inattendue.");
    }
    var septAyoub = nombre(g.f[15][1]), septSaida = nombre(g.f[14][9]);
    if (septAyoub === null || septSaida === null) {
      arreter(SALARIES + "!B16 ou J15 (septembre tapé à la main) n'est pas un nombre : « " +
        g.t[15][1] + " », « " + g.t[14][9] + " ».");
    }
    if (g.f[15].slice(2, 6).concat(g.f[14].slice(10, 14)).some(function (x) { return str(x).trim(); })) {
      arreter(SALARIES + ", lignes 15-16 : il y a plus que la date et le montant de septembre.");
    }
    var occupees = [];
    for (var l = 1; l <= 30; l++) {
      for (var c = 1; c <= 21; c++) {
        if (str(g.f[l - 1][c - 1]).trim() && (l >= 19 || l === 17 || c >= P)) occupees.push(col(c) + l);
      }
    }
    if (occupees.length) arreter(SALARIES + " : cases censées être vides (" + occupees.slice(0, 8).join(", ") + ").");

    // Associés (Karim).
    var k = grille(kar, 1, 1, 25, 26);
    if (str(k.f[0][0]).indexOf("KARIM") < 0) arreter(KARIM + "!A1 vaut « " + k.f[0][0] + " ».");
    attendre(k, 3, 11, "Aout 26", KARIM);
    attendre(k, 3, 12, "TOTAL", KARIM);
    attendre(k, 19, 12, "=K19", KARIM);
    if (str(k.f[5][10]).split("2026-08").length - 1 !== 2) {
      arreter(KARIM + "!K6 : la formule d'août ne contient pas deux fois « 2026-08 ».");
    }
    occupees = [];
    for (l = 1; l <= 25; l++) {
      for (c = 13; c <= 26; c++) if (str(k.f[l - 1][c - 1]).trim()) occupees.push(col(c) + l);
    }
    if (occupees.length) arreter(KARIM + " : colonnes M et suivantes pas vides (" + occupees.slice(0, 8).join(", ") + ").");

    // CdR Année 2.
    var r = grille(cdr, 1, 1, 24, 15);
    [[2, 8, "Sep 26"], [2, 13, "Fév 27"], [20, 1, "Salaires — Ayoub"], [21, 1, "Salaires — Saïda"],
      [22, 1, "Salaire femme de ménage"]].forEach(function (a) { attendre(r, a[0], a[1], a[2], CDR); });
    if (grille(cdr, 1, 1, 80, 1).t.some(function (x) { return x[0].indexOf(LIBELLE_CDR_ABDUL) >= 0; })) {
      arreter("« " + LIBELLE_CDR_ABDUL + " » existe déjà dans " + CDR + ".");
    }
    attendre(r, 2, 2, "Mar 26", CDR);
    if (cdr.getRange(L_CDR_RESULTAT - 1, 1).getDisplayValue().trim() !== "RÉSULTAT NET") {
      arreter(CDR + "!A" + (L_CDR_RESULTAT - 1) + " n'est pas « RÉSULTAT NET ».");
    }

    var categories = grille(onglet(classeur, "Index"), 63, 1, 38, 1).t.map(function (x) { return x[0].trim(); });
    if (categories.indexOf("Remboursement Abdul") < 0) {
      arreter("« Remboursement Abdul » n'est pas dans la liste des catégories (Index!A63…).");
    }

    // Associés : la ligne 20 recevra les clés de mois.
    if (k.f[L_CLES_KARIM - 1].some(function (x) { return str(x).trim(); })) {
      arreter(KARIM + ", ligne " + L_CLES_KARIM + " : pas vide.");
    }

    // Bilan Année 2.
    var bil = onglet(classeur, BILAN);
    var b = grille(bil, 1, 1, 55, 13);
    [[9, 1, "Mobilier & équipement"], [10, 1, "Aménagements & travaux"], [12, 1, "Lemon Tree Riad & Hostel"],
      [13, 1, "Amortissements cumulés"], [15, 1, "Charges payées d'avance"], [16, 1, "Créances sur associés"],
      [17, 1, "Créances salariales"], [18, 1, "Compte d'attente — écarts de caisse"], [19, 1, "TRÉSORERIE"],
      [20, 1, "Caisse"], [21, 1, "Banque"], [22, 1, "TOTAL ACTIF"], [10, 5, "Résultat de l'exercice"],
      [14, 5, "Compte courant Karim"], [18, 5, "Dette salariale — Ayoub"], [19, 5, "Dette salariale — Saïda"],
      [13, 2, "=-'Échéancier'!$M$64"], [10, 6, "='" + CDR + "'!O50"]
    ].forEach(function (a) { attendre(b, a[0], a[1], a[2], BILAN); });
    // La fin de la plage bouge quand une ligne de Caisse est supprimée (5233 le 02/10) : on garde la sienne.
    var caisse = str(b.f[19][1]).trim().match(/^=7881\+SUM\(Caisse!\$D\$2:\$D\$(\d+)\)$/);
    if (!caisse) arreter(BILAN + "!B20 vaut « " + b.f[19][1] + " », attendu « =7881+SUM(Caisse!$D$2:$D$…) ».");
    var finCaisse = Number(caisse[1]);
    [[16, 2, "'Associés'!$K$19"], [14, 6, "'Associés'!$K$19"], [17, 2, "'Associés-Salariés'!$E$14"],
      [18, 6, "'Associés-Salariés'!$E$14"], [19, 6, "'Associés-Salariés'!$M$14"]].forEach(function (a) {
      if (str(b.f[a[0] - 1][a[1] - 1]).indexOf(a[2]) < 0) arreter(BILAN + "!" + col(a[1]) + a[0] + " ne lit pas " + a[2] + ".");
    });
    var pleines = [];
    [3].concat([42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55]).forEach(function (l) {
      for (var c = 1; c <= 13; c++) if (str(b.f[l - 1][c - 1]).trim()) pleines.push(col(c) + l);
    });
    if (pleines.length) arreter(BILAN + " : cases censées être vides (" + pleines.slice(0, 8).join(", ") + ").");
    // Les formules coupées au 01/03/2027 passent à la date d'arrêté.
    var coupees = {};
    [9, 10, 12, 18, 21].forEach(function (l) {
      var f = str(b.f[l - 1][1]);
      var nouvelle = f.replace(/DATE\(2027[,;]\s*3[,;]\s*1\)/g, "$$B$$3+1");
      if (nouvelle === f) arreter(BILAN + "!B" + l + " ne contient pas DATE(2027,3,1).");
      if (/[,;]/.test(nouvelle.replace(/"[^"]*"/g, ""))) {
        arreter(BILAN + "!B" + l + " contient d'autres séparateurs : formule à revoir à la main.");
      }
      coupees[l] = nouvelle;
    });

    // Octobre → février déjà tapés dans le CdR (en négatif) : repris comme salaire dû,
    // sinon la formule « moins le salaire dû » les remettrait à zéro.
    var dejaAuCdr = function (ligne) {
      return r.f[ligne - 1].slice(8, 13).map(function (x) { var n = nombre(x); return n ? -n : null; });
    };

    return { sal: sal, kar: kar, cdr: cdr, bil: bil, coupees: coupees, finCaisse: finCaisse,
      septAyoub: septAyoub, septSaida: septSaida,
      repris: { Ayoub: dejaAuCdr(L_CDR_AYOUB), Saida: dejaAuCdr(L_CDR_SAIDA) },
      actuel: { Ayoub: r.t[19].slice(7, 13), Saida: r.t[20].slice(7, 13) } };
  }

  // ---------- Les changements ----------

  function requetes(v) {
    var s = v.sal.getSheetId(), k = v.kar.getSheetId(), c = v.cdr.getSheetId(), b = v.bil.getSheetId();
    var reqs = [];
    var i, colonne;

    // 1. Ayoub et Saïda : 6 mois de plus après août.
    reqs.push(inserer(s, "ROWS", L_SEPT - 1, L_SEPT + 5));
    reqs.push(copier(s, [13, 1, 13, 14], [L_SEPT, 1, L_SEPT + 5, 14]));
    NOUVEAUX_MOIS.forEach(function (mois, i) {
      [[1, 2], [9, 10]].forEach(function (p) {
        reqs.push(ecrire(s, L_SEPT + i, p[0], mois));
        reqs.push(ecrire(s, L_SEPT + i, p[1], null));
      });
    });
    reqs.push(ecrire(s, L_SEPT, 2, v.septAyoub));
    reqs.push(ecrire(s, L_SEPT, 10, v.septSaida));
    for (i = 0; i < 5; i++) {
      if (v.repris.Ayoub[i] !== null) reqs.push(ecrire(s, L_SEPT + 1 + i, 2, v.repris.Ayoub[i]));
      if (v.repris.Saida[i] !== null) reqs.push(ecrire(s, L_SEPT + 1 + i, 10, v.repris.Saida[i]));
    }
    reqs.push(vider(s, L_RESTE_AYOUB, 1, L_RESTE_AYOUB, 6));
    reqs.push(vider(s, L_RESTE_SAIDA, 9, L_RESTE_SAIDA, 14));
    // TOTAL d'Ayoub : jusqu'à la ligne repas, qui suit déjà le solde cumulé.
    reqs.push(ecrire(s, L_TOTAL, 2, "=SUM(B3:B" + L_REPAS_AYOUB + ")"),
      ecrire(s, L_TOTAL, 3, "=SUM(C3:C" + L_REPAS_AYOUB + ")"),
      ecrire(s, L_TOTAL, 5, "=E" + L_REPAS_AYOUB));

    // 2. Abdul, en P:U, sur le modèle d'Ayoub.
    reqs.push(copier(s, [1, 1, 3, 6], [1, P, 3, U]));
    reqs.push(copier(s, [13, 1, 13, 6], [4, P, 3 + MOIS_ABDUL.length, U]));
    reqs.push(copier(s, [L_HORS_CIRCUIT, 1, L_HORS_CIRCUIT, 6], [L_HC_ABDUL, P, L_HC_ABDUL, U]));
    reqs.push(copier(s, [L_TOTAL, 1, L_TOTAL, 6], [L_TOTAL, P, L_TOTAL, U]));
    var zone = [1, P, L_TOTAL, U];
    reqs.push(remplacer(s, zone, "Remboursement Ayoub", "Remboursement Abdul"));
    reqs.push(remplacer(s, zone, "$A", "$P"));
    reqs.push(ecrire(s, 2, P, "Abdul"), ecrire(s, 3, Q, 0), ecrire(s, 3, R, 0), ecrire(s, 3, U, COMMENTAIRE_ABDUL));
    MOIS_ABDUL.forEach(function (mois, i) {
      reqs.push(ecrire(s, 4 + i, P, mois), ecrire(s, 4 + i, Q, null));
    });
    reqs.push(ecrire(s, L_HC_ABDUL, Q, 0), ecrire(s, L_HC_ABDUL, R, 0), ecrire(s, L_HC_ABDUL, U, null),
      ecrire(s, L_TOTAL, Q, "=SUM(Q3:Q" + L_HC_ABDUL + ")"),
      ecrire(s, L_TOTAL, R, "=SUM(R3:R" + L_HC_ABDUL + ")"),
      ecrire(s, L_TOTAL, 20, "=T" + L_HC_ABDUL));
    for (i = 0; i < 6; i++) reqs.push(largeur(s, P + i, v.sal.getColumnWidth(1 + i)));

    // 3. Karim : Sep 26 à Fév 27, insérés entre « Aout 26 » (K) et TOTAL.
    reqs.push(inserer(k, "COLUMNS", 11, 17));
    reqs.push(copier(k, [3, 11, 19, 11], [3, 12, 19, 17]));
    NOUVEAUX_MOIS.forEach(function (mois, i) {
      reqs.push(remplacer(k, [4, 12 + i, 19, 12 + i], "2026-08", mois));
      reqs.push(ecrire(k, 3, 12 + i, ENTETES_KARIM[i]));
    });
    // TOTAL (désormais en R) : SUM(B:K) → SUM(B:Q), et solde = Q19.
    reqs.push(remplacer(k, [4, 18, 19, 18], "K", "Q"));
    var pxK = v.kar.getColumnWidth(11);
    for (i = 0; i < 6; i++) reqs.push(largeur(k, 12 + i, pxK));

    // 4. CdR : ligne Abdul sous Saïda, puis les trois lignes reliées.
    reqs.push(inserer(c, "ROWS", L_CDR_ABDUL - 1, L_CDR_ABDUL));
    reqs.push(copier(c, [L_CDR_SAIDA, 1, L_CDR_SAIDA, 15], [L_CDR_ABDUL, 1, L_CDR_ABDUL, 15]));
    reqs.push(ecrire(c, L_CDR_ABDUL, 1, LIBELLE_CDR_ABDUL));
    for (colonne = 2; colonne < 8; colonne++) reqs.push(ecrire(c, L_CDR_ABDUL, colonne, null)); // mars à août 26 : rien pour Abdul
    COLS_CDR.forEach(function (colonne, i) {
      reqs.push(ecrire(c, L_CDR_AYOUB, colonne, "=-'" + SALARIES + "'!B" + (L_SEPT + i)));
      reqs.push(ecrire(c, L_CDR_SAIDA, colonne, "=-'" + SALARIES + "'!J" + (L_SEPT + i)));
      reqs.push(ecrire(c, L_CDR_ABDUL, colonne, "=-'" + SALARIES + "'!Q" + (5 + i)));
    });

    // 5. Bilan, arrêté à la date de B3. Les lignes 19+ descendent d'un cran.
    reqs.push(ecrire(k, L_CLES_KARIM, 1, "Mois (clé pour le Bilan)"));
    MOIS_KARIM.forEach(function (mois, i) { reqs.push(ecrire(k, L_CLES_KARIM, 2 + i, mois)); });

    reqs.push(inserer(b, "ROWS", L_BILAN_ABDUL - 1, L_BILAN_ABDUL));
    var serie = (Date.UTC(DATE_ARRETE.annee, DATE_ARRETE.mois - 1, DATE_ARRETE.jour) - Date.UTC(1899, 11, 30)) / 86400000;
    reqs.push({ updateCells: {
      range: plage(b, 3, 2, 3, 2),
      rows: [{ values: [{ userEnteredValue: { numberValue: serie },
        userEnteredFormat: { numberFormat: { type: "DATE", pattern: "dd/mm/yyyy" }, textFormat: { bold: true } } }] }],
      fields: "userEnteredValue,userEnteredFormat.numberFormat,userEnteredFormat.textFormat.bold" } });
    reqs.push(
      ecrire(b, 2, 1, "Situation à la date d'arrêté (B3) — exercice en cours du 01/03/2026 au 28/02/2027"),
      ecrire(b, 3, 1, "Arrêté au :"),
      ecrire(b, 3, 3, "← changer la date (fin de mois). À remettre à jour à la main : charges payées d'avance (B15)."));

    // Soldes des quatre comptes à la date, sous le Bilan.
    var S = "'" + SALARIES + "'!", K = "'" + KARIM + "'!";
    var soldes = [
      ["Karim — compte courant (> 0 : l'hôtel lui doit)",
        "=" + K + "$B$4+SUMPRODUCT((DATEVALUE(" + K + "$B$" + L_CLES_KARIM + ":$Q$" + L_CLES_KARIM + "&\"-01\")" + JUSQUA + ")" +
        "*" + K + "$B$18:$Q$18)",
        "Associés : report initial + flux des mois jusqu'à la date"],
      ["Ayoub — salaire (> 0 : reste dû)",
        "=" + S + "$E$3+SUMPRODUCT((DATEVALUE(" + S + "$A$4:$A$19&\"-01\")" + JUSQUA + ")*" + S + "$D$4:$D$19)" +
        "+" + S + "$D$" + L_HORS_CIRCUIT + "+" + S + "$D$" + L_REPAS_AYOUB,
        "Associés-Salariés : historique + mois jusqu'à la date + hors circuit et repas"],
      ["Saïda — salaire (> 0 : reste dû)",
        "=" + S + "$M$3+SUMPRODUCT((DATEVALUE(" + S + "$I$4:$I$19&\"-01\")" + JUSQUA + ")*" + S + "$L$4:$L$19)" +
        "+" + S + "$L$" + L_HORS_CIRCUIT,
        "Associés-Salariés : historique + mois jusqu'à la date + hors circuit"],
      ["Abdul — salaire (> 0 : reste dû)",
        "=" + S + "$T$3+SUMPRODUCT((DATEVALUE(" + S + "$P$4:$P$10&\"-01\")" + JUSQUA + ")*" + S + "$S$4:$S$10)" +
        "+" + S + "$S$" + L_HC_ABDUL,
        "Associés-Salariés : mois jusqu'à la date + hors circuit"]
    ];
    reqs.push(ecrire(b, L_SOLDES, 1, "SOLDES DES COMPTES À LA DATE D'ARRÊTÉ — calculés, ne pas saisir"));
    soldes.forEach(function (x, i) {
      reqs.push(ecrire(b, L_SOLDES + 1 + i, 1, x[0]), ecrire(b, L_SOLDES + 1 + i, 2, x[1]), ecrire(b, L_SOLDES + 1 + i, 3, x[2]));
    });
    var karim = "$B$" + (L_SOLDES + 1), ayoub = "$B$" + (L_SOLDES + 2), saida = "$B$" + (L_SOLDES + 3), abdul = "$B$" + (L_SOLDES + 4);

    reqs.push(ecrire(b, L_CLES_CDR, 1, "Mois de l'exercice (clé pour couper le résultat du CdR)"));
    MOIS_EXERCICE.forEach(function (mois, i) { reqs.push(ecrire(b, L_CLES_CDR, 2 + i, mois)); });

    // Actif.
    Object.keys(v.coupees).forEach(function (cle) {
      var l = Number(cle);
      reqs.push(ecrire(b, l + (l >= L_BILAN_ABDUL ? 1 : 0), 2, v.coupees[cle]));
    });
    reqs.push(
      ecrire(b, 13, 2, "=-('Échéancier'!$I$64+SUMPRODUCT(('Échéancier'!$AC$4:$AN$4<$B$3+1)*'Échéancier'!$AC$64:$AN$64))"),
      ecrire(b, 13, 4, "Échéancier : cumul au 28/02/2026 + dotations mensuelles jusqu'à la date"),
      ecrire(b, 16, 2, "=-(" + karim + "<0)*" + karim),
      ecrire(b, 16, 4, "Solde Karim (" + karim.replace(/\$/g, "") + ") s'il est négatif"),
      ecrire(b, 17, 2, "=-(" + ayoub + "<0)*" + ayoub + "-(" + saida + "<0)*" + saida + "-(" + abdul + "<0)*" + abdul),
      ecrire(b, 21, 2, "=7881+SUMPRODUCT((Caisse!$B$2:$B$" + v.finCaisse + "<$B$3+1)*Caisse!$D$2:$D$" + v.finCaisse + ")"),
      ecrire(b, 21, 4, "ouverture 7 881 + journal de caisse jusqu'à la date"),
      ecrire(b, 22, 4, "Banque, mouvements jusqu'à la date"));
    // Passif.
    reqs.push(
      ecrire(b, 10, 6, "=SUMPRODUCT((DATEVALUE($B$" + L_CLES_CDR + ":$M$" + L_CLES_CDR + "&\"-01\")" + JUSQUA + ")" +
        "*'" + CDR + "'!$B$" + L_CDR_RESULTAT + ":$M$" + L_CDR_RESULTAT + ")"),
      ecrire(b, 10, 8, CDR + ", résultat net de mars à la date"),
      ecrire(b, 14, 6, "=(" + karim + ">0)*" + karim),
      ecrire(b, 14, 8, "Solde Karim à la date (ligne 45)"),
      ecrire(b, 18, 6, "=(" + ayoub + ">0)*" + ayoub),
      ecrire(b, 18, 8, "Solde Ayoub à la date (ligne 46)"),
      ecrire(b, L_BILAN_ABDUL, 5, "Dette salariale — Abdul"),
      ecrire(b, L_BILAN_ABDUL, 6, "=(" + abdul + ">0)*" + abdul),
      ecrire(b, L_BILAN_ABDUL, 7, 0),
      ecrire(b, L_BILAN_ABDUL, 8, "Solde Abdul à la date (ligne 48) — compte ouvert le 19/08/2026"),
      ecrire(b, 20, 6, "=(" + saida + ">0)*" + saida),
      ecrire(b, 20, 8, "Solde Saïda à la date (ligne 47)"));

    return reqs;
  }

  function afficherPlan(v) {
    dire();
    dire(SALARIES + " — Ayoub et Saïda :");
    dire("  lignes " + L_SEPT + " à " + (L_SEPT + 5) + " insérées : " + NOUVEAUX_MOIS[0] + " → " + NOUVEAUX_MOIS[5]);
    dire("  septembre dû : Ayoub " + v.septAyoub + " (repris de B16), Saïda " + v.septSaida + " (repris de J15)");
    var mots = function (liste) {
      return liste.map(function (n, i) { return n === null ? null : NOUVEAUX_MOIS[i + 1] + " " + n; })
        .filter(function (x) { return x; }).join(", ") || "rien";
    };
    dire("  octobre → février, repris du CdR : Ayoub " + mots(v.repris.Ayoub) + " ; Saïda " + mots(v.repris.Saida));
    dire("  les autres mois : salaire dû vide, à remplir");
    dire("  restes effacés : A16:B16 (Ayoub) et I15:J15 (Saïda)");
    dire("  TOTAL d'Ayoub : compte la ligne repas");
    dire(SALARIES + " — Abdul en P:U : " + MOIS_ABDUL[0] + " → " + MOIS_ABDUL[MOIS_ABDUL.length - 1] + ", salaire dû vide,");
    dire("  salaire payé = « Remboursement Abdul »");
    dire(KARIM + " — colonnes " + ENTETES_KARIM.join(", ") + " insérées avant TOTAL");
    dire(CDR + " — ligne 22 « " + LIBELLE_CDR_ABDUL + " » ; de Sep 26 à Fév 27 :");
    dire("  Ayoub (20) = -B" + L_SEPT + "…B" + (L_SEPT + 5) + ", aujourd'hui : " + v.actuel.Ayoub.join(" | "));
    dire("  Saïda (21) = -J" + L_SEPT + "…J" + (L_SEPT + 5) + ", aujourd'hui : " + v.actuel.Saida.join(" | "));
    dire("  Abdul (22) = -Q5…Q10");
    dire(BILAN + " — arrêté au " + DATE_ARRETE.jour + "/" + DATE_ARRETE.mois + "/" + DATE_ARRETE.annee + " (case B3, à changer) :");
    dire("  coupés à la date : résultat, caisse, banque, écarts de caisse, immobilisations,");
    dire("  amortissements, comptes Karim / Ayoub / Saïda / Abdul (soldes en lignes 45 à 48)");
    dire("  ligne " + L_BILAN_ABDUL + " insérée : « Dette salariale — Abdul »");
    dire("  reste à la main : charges payées d'avance (B15)");
  }

  function relire(classeur) {
    SpreadsheetApp.flush();
    var sal = onglet(classeur, SALARIES), kar = onglet(classeur, KARIM), cdr = onglet(classeur, CDR), bil = onglet(classeur, BILAN);
    dire();
    dire("Relecture — Associés-Salariés, lignes 12 à 24 (A:E · I:M · P:T) :");
    grille(sal, 12, 1, 13, 20).t.forEach(function (l) {
      if (l.some(function (x) { return x; })) {
        dire("   " + l.slice(0, 5).join(" | ") + "  ‖  " + l.slice(8, 13).join(" | ") + "  ‖  " + l.slice(15, 20).join(" | "));
      }
    });
    dire();
    dire("Relecture — Karim, en-têtes et solde cumulé :");
    var k = grille(kar, 3, 10, 17, 9).t;
    dire("   " + k[0].join(" | "));
    dire("   " + k[16].join(" | "));
    dire();
    dire("Relecture — CdR, Sep 26 :");
    grille(cdr, 19, 1, 6, 8).t.forEach(function (l) { dire("   " + l[0] + " → " + l[7]); });
    dire("   TOTAL CHARGES FIXES (H32) : " + cdr.getRange("H32").getFormula());
    dire();
    dire("Relecture — Bilan :");
    grille(bil, 3, 1, 24, 6).t.forEach(function (l) {
      if (l.some(function (x) { return x; })) dire("   " + l[0].slice(0, 38) + "  " + l[1] + "   " + l[4].slice(0, 30) + "  " + l[5]);
    });
    grille(bil, 45, 1, 4, 2).t.forEach(function (l) { dire("   " + l[0].slice(0, 50) + "  " + l[1]); });
    var erreurs = ["B3", "B13", "B16", "B17", "B21", "B22", "F10", "F14", "F18", "F19", "F20", "B45", "B46", "B47", "B48"]
      .filter(function (a1) { return bil.getRange(a1).getDisplayValue().charAt(0) === "#"; })
      .map(function (a1) { return BILAN + "!" + a1; });
    dire();
    dire("Cases en erreur : " + (erreurs.length ? erreurs.join(", ") : "aucune."));
  }

  // ---------- Programme ----------

  function lancer(ecrit) {
    sortie = [];
    try {
      var classeur = SpreadsheetApp.getActiveSpreadsheet();
      dire("Classeur : " + classeur.getName());
      var v = verifier(classeur);
      dire("Vérifications : OK.");
      afficherPlan(v);

      if (!ecrit) {
        dire();
        dire("Simulation seulement : rien n'a été modifié.");
        dire("Pour appliquer : lancer comptesAssociesEcrire.");
        return;
      }

      // Le verrou de la synchro de la caisse : pas d'écriture pendant qu'elle tourne.
      var verrou = LockService.getScriptLock();
      if (!verrou.tryLock(30000)) arreter("la synchro de la caisse tourne. Relancer dans quelques minutes.");
      try {
        Sheets.Spreadsheets.batchUpdate({ requests: requetes(v) }, classeur.getId());
      } finally {
        verrou.releaseLock();
      }
      dire();
      dire("Écrit.");
      relire(classeur);
    } finally {
      console.log(sortie.join("\n"));
    }
  }

  return { lancer: lancer };
})();
