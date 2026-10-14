/**
 * Soldes des comptes salaires, relevés pour l'appli (Admin → Salaires) :
 * chacun y voit le sien, Ayoub (manager) et l'admin les trois.
 *
 * Lus dans « Associés-Salariés », sur la ligne TOTAL de chaque bloc,
 * colonne du solde cumulé. Positif : l'hôtel doit encore du salaire ;
 * négatif : la personne a reçu plus que son dû.
 *
 * Avec le solde partent les lignes du bloc (historique, chaque mois, hors
 * circuit, repas) : libellé, salaire dû (2e colonne) et net de la ligne
 * (4e colonne, celle que le Bilan additionne). « Pris » s'en déduit, dû
 * moins net : pas besoin de savoir dans quel signe le classeur range les
 * versements. Les lignes vides (dû et net à zéro) ne partent pas.
 *
 * Envoyés par `synchroniser()` (14h), juste après le solde de caisse, à
 * l'Edge Function `export-caisse` (action « soldes_salaries »). Un échec ici
 * ne bloque jamais la synchro de la caisse : il est seulement noté.
 *
 * `verifierSoldesSalaries` (depuis l'éditeur) affiche ce qui serait envoyé,
 * case par case, sans rien envoyer.
 */

var ONGLET_SALARIES = "Associés-Salariés";

/**
 * Libellé lu en colonne `libelle`, solde cumulé en colonne `solde` (1 = A).
 * Dans chaque bloc : libellé, dû, payé, net, cumul — d'où `libelle + 1`
 * pour le dû et `libelle + 3` pour le net.
 */
var BLOCS_SALARIES = [
  { personne: "ayoub", libelle: 1, solde: 5 },  // A … E
  { personne: "saida", libelle: 9, solde: 13 }, // I … M
  { personne: "abdul", libelle: 16, solde: 20 } // P … T
];

/**
 * Rend les soldes lisibles. Un bloc sans ligne TOTAL ou dont le solde
 * n'est pas un nombre est laissé de côté (et noté) : mieux vaut un solde
 * absent dans l'appli qu'un chiffre faux.
 */
function lireSoldesSalaries() {
  var feuille = SpreadsheetApp.getActive().getSheetByName(ONGLET_SALARIES);
  if (!feuille) {
    console.error("Onglet « " + ONGLET_SALARIES + " » introuvable — soldes salaires non relevés.");
    return [];
  }
  var hauteur = Math.min(feuille.getLastRow(), 80);
  var plage = feuille.getRange(1, 1, hauteur, 21);
  var valeurs = plage.getValues();
  var textes = plage.getDisplayValues();
  var soldes = [];
  BLOCS_SALARIES.forEach(function (b) {
    for (var l = 0; l < hauteur; l++) {
      if (String(valeurs[l][b.libelle - 1]).trim().toUpperCase() !== "TOTAL") continue;
      var solde = valeurs[l][b.solde - 1];
      var cellule = feuille.getRange(l + 1, b.solde).getA1Notation();
      if (typeof solde !== "number" || !isFinite(solde)) {
        console.error("Solde de " + b.personne + " illisible en " + cellule + " — non envoyé.");
      } else {
        soldes.push({
          personne: b.personne,
          solde: arrondir(solde),
          cellule: cellule,
          lignes: lignesDuBloc(valeurs, textes, b, l)
        });
      }
      return;
    }
    console.error("Pas de ligne TOTAL pour " + b.personne + " — solde non envoyé.");
  });
  return soldes;
}

/** Lignes 3 à TOTAL-1 du bloc, sans les lignes vides. */
function lignesDuBloc(valeurs, textes, b, ligneTotal) {
  var lignes = [];
  for (var l = 2; l < ligneTotal; l++) {
    var libelle = String(textes[l][b.libelle - 1]).trim();
    if (!libelle) continue;
    var du = nombreOuZero(valeurs[l][b.libelle]);
    var net = nombreOuZero(valeurs[l][b.libelle + 2]);
    if (du === 0 && net === 0) continue;
    lignes.push({ libelle: libelle.slice(0, 40), du: du, net: net });
  }
  return lignes;
}

function nombreOuZero(v) {
  return typeof v === "number" && isFinite(v) ? arrondir(v) : 0;
}

function arrondir(n) {
  return Math.round(n * 100) / 100;
}

/** Appelé par `synchroniser()`. Ne lève jamais. */
function releverSoldesSalaries() {
  try {
    var soldes = lireSoldesSalaries();
    if (soldes.length === 0) return;
    appeler({
      action: "soldes_salaries",
      soldes: soldes.map(function (s) {
        return { personne: s.personne, solde: s.solde, lignes: s.lignes };
      })
    });
  } catch (e) {
    console.error("Soldes salaires non envoyés : " + e.message);
  }
}

/** Depuis l'éditeur : ce qui partirait, sans rien envoyer. */
function verifierSoldesSalaries() {
  var soldes = lireSoldesSalaries();
  console.log(soldes.length
    ? soldes.map(function (s) {
        return s.personne + " : " + s.solde + " (case " + s.cellule + ")\n" +
          s.lignes.map(function (x) {
            return "    " + x.libelle + " : dû " + x.du + ", pris " + arrondir(x.du - x.net) + ", net " + x.net;
          }).join("\n");
      }).join("\n")
    : "Aucun solde lisible.");
}
