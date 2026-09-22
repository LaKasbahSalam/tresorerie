/**
 * Synchronisation de l'onglet « Caisse » du classeur Exercices.
 *
 * ATTENTION — copie de référence. Le script qui tourne réellement vit DANS
 * le classeur (Extensions -> Apps Script). Ce fichier n'est là que pour être
 * versionné et relu avec le reste du projet. Toute modification faite dans
 * le classeur doit être reportée ici, sinon les deux divergent en silence.
 *
 * Il dialogue avec l'Edge Function `export-caisse` (voir son README pour le
 * protocole). Le sens est un TIRAGE : ce script demande, la fonction répond,
 * et elle n'écrit jamais dans le classeur. C'est ce qui évite un compte de
 * service Google et une clé privée à protéger.
 *
 * Trois temps par passage :
 *   1. demander les lignes classées non encore exportées ;
 *   2. les écrire en bas de l'onglet, colonnes A à H ;
 *   3. accuser réception des seules lignes réellement écrites.
 *
 * L'accusé vient en dernier à dessein. Si l'écriture échoue, rien n'est
 * marqué et le passage suivant reprend les mêmes lignes. L'inverse les
 * perdrait : marquées côté base, absentes du classeur, plus jamais proposées.
 *
 * CE SCRIPT N'ANNONCE PLUS RIEN dans Telegram. Il relève le solde et le
 * dépose en base (`action: "solde"`) ; c'est l'Edge Function `rappel-soir`,
 * réveillée par son propre cron indépendant de ce déclencheur, qui poste le
 * message du soir en relisant le dernier relevé connu. Avant ce découplage,
 * une synchronisation en échec ici (catégorie invalide, panne réseau)
 * faisait disparaître le message du soir tout entier — c'est ce qui s'est
 * produit du 23 au 27 août 2026. Ce script reste donc à 20h, une heure
 * avant la fenêtre de `rappel-soir` (21h-22h), pour que le solde ait le
 * temps d'être en base avant que le rappel ne le lise.
 *
 * INSTALLATION
 *   1. Dans le classeur : Extensions -> Apps Script, coller ce fichier.
 *   2. Paramètres du projet -> Propriétés du script, ajouter :
 *        URL_FONCTION  https://<projet>.supabase.co/functions/v1/export-caisse
 *        SECRET        la même valeur que EXPORT_CAISSE_SECRET côté Supabase
 *   3. Lancer `installerDeclencheurs` une fois depuis l'éditeur.
 *   4. Le menu « Caisse » apparaît dans le classeur après un rechargement.
 */

// ---------- Réglages ----------

/** Nom exact de l'onglet. Le script s'arrête plutôt que d'écrire ailleurs. */
var ONGLET = "Caisse";

/**
 * Heure de la synchronisation quotidienne, en heure du Maroc. Reculée de
 * 21h à 20h quand l'annonce Telegram est devenue le travail de
 * `rappel-soir` (fenêtre 21h-22h) : le solde doit être en base avant que ce
 * dernier ne le lise.
 */
var HEURE_CLOTURE = 20;

/** Solde d'ouverture, repris de l'en-tête de la colonne H. */
var OUVERTURE = 7881;

/** Première ligne de données (la 1 porte les en-têtes). */
var PREMIERE_LIGNE = 2;
/** Solde officiel de la caisse, posé par Karim (17/09/2026). */
var CELLULE_SOLDE_CAISSE = "L1";

// ---------- Menu ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Caisse")
    .addItem("Synchroniser maintenant", "synchroniserManuel")
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu("CdR")
    .addItem("Insérer la formule ici", "insererFormuleCdR")
    .addToUi();

  // Envoi du CdR vers la base d'analyse — voir ExportCdr.js
  SpreadsheetApp.getUi()
    .createMenu("Analyse")
    .addItem("Envoyer le CdR maintenant", "exporterCdrManuel")
    .addToUi();
}

function synchroniserManuel() {
  var n = synchroniser();
  SpreadsheetApp.getActive().toast(
    n === 0 ? "Rien à ajouter." : n + " ligne(s) ajoutée(s).",
    "Caisse"
  );
}

// ---------- Déclencheurs ----------

function installerDeclencheurs() {
  // Repartir de zéro : relancer cette fonction ne doit pas empiler les
  // déclencheurs.
  var existants = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existants.length; i++) {
    ScriptApp.deleteTrigger(existants[i]);
  }

  // Un seul déclencheur automatique : la synchronisation, une fois par
  // jour. Le rythme horaire d'origine s'est révélé trop fréquent à
  // l'usage — le menu Caisse > Synchroniser maintenant reste disponible à
  // la main pour voir l'état en cours de journée.
  ScriptApp.newTrigger("syncCloture")
    .timeBased()
    .atHour(HEURE_CLOTURE)
    .everyDays(1)
    .inTimezone("Africa/Casablanca")
    .create();

  // Les déclencheurs viennent d'être tous supprimés : recréer celui de
  // l'envoi du CdR (22h), défini dans ExportCdr.js.
  installerDeclencheurExportCdr();
}

function syncCloture() {
  synchroniser();
}

// ---------- Coeur ----------

/**
 * @return {number} nombre de lignes ajoutées.
 */
function synchroniser() {
  var feuille = SpreadsheetApp.getActive().getSheetByName(ONGLET);
  if (!feuille) {
    throw new Error("Onglet « " + ONGLET + " » introuvable — rien n'a été écrit.");
  }

  // Verrou : le déclencheur automatique et un clic manuel peuvent se
  // chevaucher. Sans lui, deux passages liraient la même file avant que
  // l'un ait accusé réception, et écriraient les lignes en double.
  var verrou = LockService.getScriptLock();
  if (!verrou.tryLock(30000)) {
    return 0;
  }

  try {
    var lignes = appeler({ action: "pending", limit: 500 }).lignes || [];
    var ajoutees = 0;

    if (lignes.length > 0) {
      // Valider AVANT d'écrire quoi que ce soit. Écrire d'abord et
      // découvrir le problème en cours de route laisse des lignes à moitié
      // écrites dans l'onglet — et comme rien n'est alors accusé réception,
      // le passage suivant retente les mêmes lignes PAR-DESSUS, en double.
      var invalides = categoriesInvalides(feuille, lignes);
      if (invalides.length > 0) {
        throw new Error(
          "Catégories absentes de la liste déroulante de la colonne E : " +
            invalides.join(", ") +
            ". Ajoute-les à la validation de données de l'onglet Caisse, " +
            "ou corrige-les côté classify, puis relance."
        );
      }

      ajoutees = ecrire(feuille, lignes);
      // N'accuser que ce qui a réellement atterri dans l'onglet.
      var ids = lignes.slice(0, ajoutees).map(function (l) {
        return l.id;
      });
      appeler({ action: "ack", ids: ids });
    }

    // Si categoriesInvalides() a levé plus haut, on n'arrive jamais ici :
    // le solde n'est pas relevé ce passage-là. Ce n'est plus grave comme
    // avant — `rappel-soir` annoncera le dernier relevé connu avec sa date
    // et un avertissement plutôt que de rester silencieux.
    releverSolde(feuille);
    return ajoutees;
  } finally {
    verrou.releaseLock();
  }
}

/**
 * Catégories des lignes en attente qui ne figurent pas dans la liste
 * déroulante de la colonne E, dédupliquées. Vide si tout est valide.
 *
 * Sans ce contrôle, `ecrire()` écrirait les lignes une par une jusqu'à
 * buter sur la première invalide, et s'arrêterait là avec un onglet
 * à moitié rempli — sans que rien ne soit accusé réception côté base.
 */
function categoriesInvalides(feuille, lignes) {
  var acceptees = categoriesAcceptees(feuille);
  if (!acceptees) return []; // pas de règle de validation posée : rien à vérifier

  // Comparaison insensible à la casse et aux espaces de bord : c'est ainsi
  // que Google Sheets applique lui-même la validation. Un contrôle plus
  // strict que la règle qu'il vérifie bloquerait des lignes que l'onglet
  // aurait acceptées — « Vente divers » face à « Vente Divers », par
  // exemple. Les accents, eux, restent significatifs des deux côtés.
  var jeu = {};
  for (var i = 0; i < acceptees.length; i++) {
    jeu[String(acceptees[i]).trim().toLowerCase()] = true;
  }

  var mauvaises = {};
  for (var j = 0; j < lignes.length; j++) {
    var cat = lignes[j].categorie;
    if (cat && !jeu[String(cat).trim().toLowerCase()]) {
      mauvaises[cat] = true;
    }
  }
  return Object.keys(mauvaises);
}

/**
 * Lit la liste des valeurs autorisées directement dans la règle de
 * validation de données posée sur la colonne E — plutôt que de la recopier
 * à la main ici, ce qui finirait tôt ou tard par diverger de ce que
 * l'onglet accepte réellement.
 *
 * Gère les deux formes que peut prendre la règle : une liste tapée à la
 * main, ou une liste tirée d'une plage. `null` si aucune règle n'est posée.
 */
function categoriesAcceptees(feuille) {
  var validation = feuille.getRange(PREMIERE_LIGNE, 5).getDataValidation();
  if (!validation) return null;

  var criteria = validation.getCriteriaType();
  var valeurs = validation.getCriteriaValues();

  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return valeurs[0];
  }
  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return valeurs[0]
      .getValues()
      .map(function (r) {
        return r[0];
      })
      .filter(function (v) {
        return v !== "";
      });
  }
  return null;
}

/**
 * Écrit les lignes en bas de l'onglet, colonnes A à H.
 *
 * La colonne H reçoit une FORMULE et non une valeur : le solde cumulé doit
 * se recalculer tout seul si tu corriges un montant plus haut. Les colonnes
 * I (comptage) et J (projet) ne sont jamais touchées — elles sont à toi.
 */
function ecrire(feuille, lignes) {
  var depart = Math.max(feuille.getLastRow() + 1, PREMIERE_LIGNE);

  var valeurs = lignes.map(function (l, i) {
    var ligneCourante = depart + i;
    // Première ligne de données : rien au-dessus d'où partir, on repart du
    // solde d'ouverture.
    var formuleSolde = ligneCourante === PREMIERE_LIGNE
      ? "=" + OUVERTURE + "+D" + ligneCourante
      : "=H" + (ligneCourante - 1) + "+D" + ligneCourante;

    // Midi et non minuit : une date construite à minuit peut basculer à la
    // veille selon le fuseau du classeur.
    var date = l.date_caisse ? new Date(l.date_caisse + "T12:00:00") : "";

    return [
      l.code,
      date,
      l.description,
      l.montant,
      l.categorie,
      l.canal,
      l.destination,
      formuleSolde
    ];
  });

  feuille.getRange(depart, 1, valeurs.length, 8).setValues(valeurs);
  // Forcer l'écriture avant d'accuser réception : sinon on confirmerait des
  // lignes encore en tampon, qu'une erreur pourrait emporter.
  SpreadsheetApp.flush();
  return valeurs.length;
}

/**
 * Lit le solde de caisse en L1 et le remonte.
 *
 * Depuis le 17/09/2026, Karim a posé le solde officiel dans la cellule L1
 * de l'onglet Caisse : c'est elle qui fait foi, pour ce relevé comme pour
 * l'envoi vers la base d'analyse (ExportCdr.js). La dernière ligne portant
 * un code reste transmise pour retrouver l'écriture en cas de doute.
 *
 * Le solde ne peut venir que d'ici : la base ne contient que l'ère Telegram,
 * l'historique WhatsApp antérieur n'y est jamais entré. C'est le classeur
 * qui fait foi.
 *
 * N'annonce plus rien : ce script se contente de déposer le relevé.
 * L'annonce dans Telegram est le travail de `rappel-soir`, qui relit le
 * dernier relevé connu (`dernier_solde_caisse()`) au moment de son propre
 * cron, avec sa date si ce relevé-ci n'a pas eu lieu ce soir.
 */
function releverSolde(feuille) {
  var derniere = derniereLigneAvecCode(feuille);
  if (derniere < PREMIERE_LIGNE) {
    return;
  }

  var solde = feuille.getRange(CELLULE_SOLDE_CAISSE).getValue();
  // Ne rien remonter plutôt que remonter n'importe quoi : un solde faux
  // serait annoncé comme s'il était vrai le soir où `rappel-soir` le lira.
  if (typeof solde !== "number" || !isFinite(solde)) {
    console.error("Solde illisible en " + CELLULE_SOLDE_CAISSE + " — relevé abandonné.");
    return;
  }

  appeler({
    action: "solde",
    solde: solde,
    ligne: derniere
  });
}

/**
 * Dernière ligne portant un code en colonne A.
 *
 * `getLastRow()` compte toute cellule non vide de la feuille, y compris une
 * note isolée en colonne J bien plus bas. On remonte donc depuis le bas
 * jusqu'à trouver un vrai code.
 */
function derniereLigneAvecCode(feuille) {
  var bas = feuille.getLastRow();
  if (bas < PREMIERE_LIGNE) {
    return 0;
  }

  var codes = feuille
    .getRange(PREMIERE_LIGNE, 1, bas - PREMIERE_LIGNE + 1, 1)
    .getValues();
  for (var i = codes.length - 1; i >= 0; i--) {
    if (codes[i][0] !== "" && codes[i][0] !== null) {
      return PREMIERE_LIGNE + i;
    }
  }
  return 0;
}

// ---------- Transport ----------

function appeler(charge) {
  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty("URL_FONCTION");
  var secret = props.getProperty("SECRET");
  if (!url || !secret) {
    throw new Error(
      "URL_FONCTION ou SECRET absent des propriétés du script " +
        "(Paramètres du projet -> Propriétés du script)."
    );
  }

  var reponse = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { "x-export-secret": secret },
    payload: JSON.stringify(charge),
    muteHttpExceptions: true
  });

  var code = reponse.getResponseCode();
  var texte = reponse.getContentText();
  if (code !== 200) {
    // Échouer bruyamment : Google enverra un mail d'erreur, et le passage
    // suivant reprendra la même file puisque rien n'a été accusé.
    throw new Error("export-caisse a répondu " + code + " : " + texte);
  }
  return JSON.parse(texte);
}

// ---------- Formules du compte de résultat ----------

/**
 * Abréviations de mois telles qu'utilisées dans les en-têtes du CdR
 * ("Mar 25", "Nov 26"...), sans accent et en minuscules — la comparaison
 * normalise l'entrée avant de chercher ici, donc "Aoû", "aou" et "AOÛ"
 * trouvent tous la même clé.
 */
var MOIS_ABBR = {
  mar: "03", avr: "04", mai: "05", juin: "06", juil: "07",
  aou: "08", sep: "09", oct: "10", nov: "11", dec: "12",
  jan: "01", fev: "02"
};

/**
 * Convertit un en-tête de colonne du CdR ("Nov 25") en "AAAA-MM" ("2025-11")
 * — le format que TEXT(..., "YYYY-MM") produit côté Caisse/Banque/CB Karim,
 * et donc celui que la formule doit comparer. `null` si le texte ne
 * ressemble pas à un mois (colonnes TOTAL, RÉALISÉ, ou en-tête Poste).
 */
function moisVersAAAAMM(texte) {
  if (!texte) return null;
  var m = String(texte).trim().match(/^([A-Za-zÀ-ÿ]{3,4})\s+(\d{2})$/);
  if (!m) return null;

  var abbr = stripAccents_(m[1]).toLowerCase();
  var mois = MOIS_ABBR[abbr];
  if (!mois) return null;

  return "20" + m[2] + "-" + mois;
}

/**
 * Retire les accents (décomposition NFD + suppression des diacritiques).
 * Même convention que `stripAccents` dans classify/lib/text.ts.
 */
function stripAccents_(s) {
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/**
 * Insère, sur la case actuellement sélectionnée d'un onglet CdR, la formule
 * standard qui somme une catégorie sur un mois depuis Caisse, Banque et
 * ' Saisies CB Karim'.
 *
 * La catégorie est lue en colonne A de la même ligne, le mois en ligne 2 de
 * la même colonne — exactement ce que les formules existantes du classeur
 * font à la main. Cette fonction ne fait qu'automatiser la même saisie,
 * elle n'invente pas une nouvelle convention.
 *
 * Aucun filtre de signe (`>0` / `<0`) : les trois colonnes sources sont
 * déjà au bon signe pour le compte de résultat — Caisse!D et Banque!C le
 * sont par construction, et ' Saisies CB Karim'!G ("Compte de resultat")
 * est justement la colonne prévue pour ça, l'opposé de la colonne D de ce
 * même onglet. Filtrer le signe couperait les corrections et
 * remboursements plutôt que de les nettoyer.
 */
function insererFormuleCdR() {
  var ui = SpreadsheetApp.getUi();
  var feuille = SpreadsheetApp.getActiveSheet();
  var cellule = SpreadsheetApp.getActiveRange().getCell(1, 1);
  var ligne = cellule.getRow();
  var colonne = cellule.getColumn();

  if (colonne === 1) {
    ui.alert("Place-toi sur une case de mois (colonnes B et suivantes), pas sur la colonne Poste.");
    return;
  }

  var categorie = feuille.getRange(ligne, 1).getValue();
  if (!categorie || typeof categorie !== "string") {
    ui.alert("La colonne A de cette ligne ne contient pas de libellé de catégorie.");
    return;
  }
  categorie = categorie.trim();

  var enTete = feuille.getRange(2, colonne).getValue();
  var aaaaMM = moisVersAAAAMM(enTete);
  if (!aaaaMM) {
    ui.alert(
      "L'en-tête de cette colonne (« " + enTete + " ») ne ressemble pas à un mois. " +
        "Colonnes TOTAL / RÉALISÉ non prises en charge."
    );
    return;
  }

  var classeur = SpreadsheetApp.getActive();
  var fCaisse = classeur.getSheetByName("Caisse");
  var fBanque = classeur.getSheetByName("Banque");
  var fCB = classeur.getSheetByName(" Saisies CB Karim");
  if (!fCaisse || !fBanque || !fCB) {
    ui.alert("Un des onglets sources (Caisse, Banque, Saisies CB Karim) est introuvable.");
    return;
  }

  // Étendue actuelle des données, lue au moment de l'exécution plutôt que
  // figée en dur : les formules existantes du classeur référencent des
  // plages fixes (ex. Caisse!$D$2:$D$4032) posées une fois à la main, qui
  // n'incluront jamais les lignes que la synchro ajoute ensuite. Relancer
  // cette fonction sur la même case reprend toujours l'étendue à jour.
  var derCaisse = Math.max(fCaisse.getLastRow(), 2);
  var derBanque = Math.max(fBanque.getLastRow(), 2);
  var derCB = Math.max(fCB.getLastRow(), 2);

  // Point-virgule entre les arguments de TEXT(), pas virgule : ce classeur
  // utilise la virgule comme séparateur décimal (paramètres régionaux
  // français), donc Google Sheets attend ";" pour séparer les arguments
  // d'une fonction — une formule avec des virgules y est rejetée à
  // l'analyse, avant même d'être évaluée.
  var formule =
    'SUMPRODUCT((TEXT(Caisse!$B$2:$B$' + derCaisse + ';"YYYY-MM")="' + aaaaMM + '")' +
      '*(Caisse!$E$2:$E$' + derCaisse + '="' + categorie + '")' +
      "*Caisse!$D$2:$D$" + derCaisse + ")" +
    '+SUMPRODUCT((TEXT(Banque!$B$2:$B$' + derBanque + ';"YYYY-MM")="' + aaaaMM + '")' +
      '*(Banque!$F$2:$F$' + derBanque + '="' + categorie + '")' +
      "*Banque!$C$2:$C$" + derBanque + ")" +
    "+SUMPRODUCT((TEXT(' Saisies CB Karim'!$C$2:$C$" + derCB + ';"YYYY-MM")="' + aaaaMM + '")' +
      "*(' Saisies CB Karim'!$E$2:$E$" + derCB + '="' + categorie + '")' +
      "*' Saisies CB Karim'!$G$2:$G$" + derCB + ")";

  var reponse = ui.alert(
    "Insérer la formule ?",
    "Catégorie : " + categorie + "\n" +
      "Mois : " + aaaaMM + " (" + enTete + ")\n" +
      "Cellule : " + cellule.getA1Notation() + "\n\n" +
      "Cherche dans Caisse (jusqu'à la ligne " + derCaisse + "), " +
      "Banque (" + derBanque + ") et Saisies CB Karim (" + derCB + ").",
    ui.ButtonSet.OK_CANCEL
  );
  if (reponse !== ui.Button.OK) return;

  cellule.setFormula(formule);
  classeur.toast("Formule posée en " + cellule.getA1Notation() + ".", "CdR");
}