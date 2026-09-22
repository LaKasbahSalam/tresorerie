/**
 * Envoi du compte de résultat vers la base d'analyse (Kasbah Analytics).
 *
 * À placer dans le projet Apps Script de la V16 (dossier
 * Tresorerie\Script, à côté de Code.js), puis `clasp push`.
 *
 * Chaque soir à 22h (heure du Maroc), après la clôture de caisse de 21h :
 * lit tous les onglets dont la cellule A1 commence par
 * « COMPTE DE RÉSULTAT — EXERCICE » et les envoie tels quels à l'Edge
 * Function `import-cdr`. Lecture seule : ce script n'écrit rien dans le
 * classeur.
 *
 * Le même envoi porte les SOLDES DE TRÉSORERIE, lus dans des cellules
 * posées par Karim le 17/09/2026 : Caisse!L1 et Banque!K1. Le classeur fait
 * foi, rien n'est recalculé. Un solde illisible n'empêche pas l'envoi du CdR.
 *
 * INSTALLATION
 *   1. Paramètres du projet → Propriétés du script, ajouter :
 *        URL_IMPORT_CDR     https://sebwcxxoxpfbliypzokp.supabase.co/functions/v1/import-cdr
 *        SECRET_IMPORT_CDR  même valeur que IMPORT_CDR_SECRET côté Supabase
 *   2. Lancer `installerDeclencheurExportCdr` une fois depuis l'éditeur.
 *   3. Test immédiat : lancer `exporterCdr` depuis l'éditeur, ou le menu
 *      « Analyse → Envoyer le CdR maintenant ».
 */

var TITRE_CDR = "COMPTE DE RÉSULTAT — EXERCICE";
var HEURE_EXPORT_CDR = 22;

function exporterCdr() {
  var onglets = SpreadsheetApp.getActive().getSheets()
    .filter(function (f) {
      return String(f.getRange(1, 1).getDisplayValue()).indexOf(TITRE_CDR) === 0;
    })
    .map(function (f) {
      var plage = f.getDataRange();
      return {
        nom: f.getName(),
        valeurs: plage.getValues().map(function (ligne) {
          return ligne.map(function (v) {
            return v instanceof Date ? v.toISOString() : v;
          });
        }),
        textes: plage.getDisplayValues()
      };
    });

  if (onglets.length === 0) {
    throw new Error("Aucun onglet dont A1 commence par « " + TITRE_CDR + " » — rien n'a été envoyé.");
  }

  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty("URL_IMPORT_CDR");
  var secret = props.getProperty("SECRET_IMPORT_CDR");
  if (!url || !secret) {
    throw new Error("URL_IMPORT_CDR ou SECRET_IMPORT_CDR absent des propriétés du script.");
  }

  var reponse = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { "x-import-secret": secret },
    payload: JSON.stringify({ hotel_id: "kasbah", onglets: onglets, soldes: releverSoldesTresorerie() }),
    muteHttpExceptions: true
  });

  var code = reponse.getResponseCode();
  var texte = reponse.getContentText();
  if (code !== 200) {
    throw new Error("import-cdr a répondu " + code + " : " + texte);
  }
  return JSON.parse(texte);
}

/**
 * Soldes lus dans le classeur, pour la trésorerie du tableau de bord.
 */
function releverSoldesTresorerie() {
  var soldes = [];
  var caisse = releverSoldeCellule("Caisse", "L1", "caisse");
  if (caisse) soldes.push(caisse);
  var banque = releverSoldeCellule("Banque", "K1", "banque");
  if (banque) soldes.push(banque);
  return soldes;
}

/**
 * Solde lu dans une cellule fixe. Rend null (et le note dans les journaux)
 * plutôt qu'un chiffre douteux.
 */
function releverSoldeCellule(nomOnglet, cellule, compte) {
  var feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  if (!feuille) {
    console.error("Onglet « " + nomOnglet + " » introuvable — solde non envoyé.");
    return null;
  }
  var plage = feuille.getRange(cellule);
  var solde = plage.getValue();
  if (typeof solde !== "number" || !isFinite(solde)) {
    console.error("Solde illisible en " + nomOnglet + "!" + cellule + " (« " + plage.getDisplayValue() + " ») — solde non envoyé.");
    return null;
  }
  return { compte: compte, solde: solde, ligne: plage.getRow() };
}

function exporterCdrManuel() {
  var r = exporterCdr();
  SpreadsheetApp.getActive().toast(
    (r.statut || "?") + " — " + (r.lignes || 0) + " montant(s). " + (r.message || ""),
    "Analyse"
  );
}

/**
 * Ajoute le déclencheur de 22h s'il n'existe pas déjà. Ne touche à aucun
 * autre déclencheur (contrairement à `installerDeclencheurs`, qui les
 * supprime tous : celui-ci l'appelle à la fin pour se recréer).
 */
function installerDeclencheurExportCdr() {
  var existe = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === "exporterCdr";
  });
  if (existe) return;
  ScriptApp.newTrigger("exporterCdr")
    .timeBased()
    .atHour(HEURE_EXPORT_CDR)
    .everyDays(1)
    .inTimezone("Africa/Casablanca")
    .create();
}
