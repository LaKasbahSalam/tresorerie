/**
 * Soldes des comptes salaires, relevés pour l'appli (Admin → Salaires) :
 * chacun y voit le sien, Ayoub (manager) et l'admin les trois.
 *
 * Lus dans « Associés-Salariés », sur la ligne TOTAL de chaque bloc,
 * colonne du solde cumulé. Positif : l'hôtel doit encore du salaire ;
 * négatif : la personne a reçu plus que son dû.
 *
 * Envoyés par `synchroniser()` (14h), juste après le solde de caisse, à
 * l'Edge Function `export-caisse` (action « soldes_salaries »). Un échec ici
 * ne bloque jamais la synchro de la caisse : il est seulement noté.
 *
 * `verifierSoldesSalaries` (depuis l'éditeur) affiche ce qui serait envoyé,
 * case par case, sans rien envoyer.
 */

var ONGLET_SALARIES = "Associés-Salariés";

/** Libellé lu en colonne `libelle`, solde lu en colonne `solde` (1 = A). */
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
  var valeurs = feuille.getRange(1, 1, hauteur, 21).getValues();
  var soldes = [];
  BLOCS_SALARIES.forEach(function (b) {
    for (var l = 0; l < hauteur; l++) {
      if (String(valeurs[l][b.libelle - 1]).trim().toUpperCase() !== "TOTAL") continue;
      var solde = valeurs[l][b.solde - 1];
      var cellule = feuille.getRange(l + 1, b.solde).getA1Notation();
      if (typeof solde !== "number" || !isFinite(solde)) {
        console.error("Solde de " + b.personne + " illisible en " + cellule + " — non envoyé.");
      } else {
        soldes.push({ personne: b.personne, solde: Math.round(solde * 100) / 100, cellule: cellule });
      }
      return;
    }
    console.error("Pas de ligne TOTAL pour " + b.personne + " — solde non envoyé.");
  });
  return soldes;
}

/** Appelé par `synchroniser()`. Ne lève jamais. */
function releverSoldesSalaries() {
  try {
    var soldes = lireSoldesSalaries();
    if (soldes.length === 0) return;
    appeler({
      action: "soldes_salaries",
      soldes: soldes.map(function (s) {
        return { personne: s.personne, solde: s.solde };
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
    ? soldes.map(function (s) { return s.personne + " : " + s.solde + " (case " + s.cellule + ")"; }).join("\n")
    : "Aucun solde lisible.");
}
