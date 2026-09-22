#!/bin/bash
#
# attach-script.sh — attache automatiquement le script de trésorerie
# (Code.js + appsscript.json) à n'importe quelle nouvelle Google Sheet,
# sans passer par Extensions > Apps Script à la main.
#
# USAGE :
#   ./attach-script.sh <spreadsheetId> ["Titre du projet"]
#
# L'ID de la Sheet est le segment entre /d/ et /edit dans son URL :
#   https://docs.google.com/spreadsheets/d/CET_ID_LA/edit
#
# PRÉREQUIS :
#   - clasp installé et connecté (clasp login déjà fait une fois)
#   - Code.js et appsscript.json présents dans le MÊME dossier que ce script
#     (ce sont les fichiers "source de vérité" — modifie-les ici, pas
#     dans un dossier clasp-vXX ponctuel)
#
# CE QUE ÇA FAIT :
#   1. Crée un nouveau projet Apps Script lié directement à la Sheet
#      indiquée (clasp create --parentId, aucune étape manuelle)
#   2. Copie Code.js et appsscript.json dedans
#   3. Pousse vers Google (clasp push)
#
# CE QUE ÇA NE FAIT PAS (reste à faire une fois par nouvelle Sheet,
# manuellement, car ce sont des secrets qui ne doivent pas être versionnés) :
#   - Renseigner URL_FONCTION et SECRET dans Paramètres du projet >
#     Propriétés du script
#   - Lancer installerDeclencheurs() une fois depuis l'éditeur Apps Script

set -euo pipefail

SPREADSHEET_ID="${1:-}"
TITLE="${2:-Script Trésorerie}"

if [ -z "$SPREADSHEET_ID" ]; then
  echo "Usage: ./attach-script.sh <spreadsheetId> [\"Titre du projet\"]"
  echo ""
  echo "L'ID se trouve dans l'URL de la Sheet :"
  echo "  https://docs.google.com/spreadsheets/d/CET_ID_LA/edit"
  exit 1
fi

SRC_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [ ! -f "$SRC_DIR/Code.js" ]; then
  echo "Erreur : Code.js introuvable dans $SRC_DIR"
  echo "Place Code.js et appsscript.json à côté de ce script."
  exit 1
fi

WORKDIR="$(mktemp -d)"
echo "Dossier de travail temporaire : $WORKDIR"
cd "$WORKDIR"

echo ""
echo "→ Création du projet Apps Script lié à la Sheet $SPREADSHEET_ID ..."
clasp create --type sheets --parentId "$SPREADSHEET_ID" --title "$TITLE"

echo ""
echo "→ Copie de Code.js et appsscript.json depuis $SRC_DIR ..."
cp "$SRC_DIR/Code.js" .
if [ -f "$SRC_DIR/appsscript.json" ]; then
  cp "$SRC_DIR/appsscript.json" .
fi

echo ""
echo "→ Envoi vers Google (clasp push) ..."
clasp push -f

echo ""
echo "✓ Script attaché et poussé sur la Sheet $SPREADSHEET_ID."
echo ""
echo "Reste à faire UNE FOIS sur cette Sheet, à la main :"
echo "  1. Extensions > Apps Script > Paramètres du projet > Propriétés du script"
echo "     - URL_FONCTION"
echo "     - SECRET"
echo "  2. Lancer installerDeclencheurs() une fois depuis l'éditeur"
echo ""
echo "Dossier local du projet (pour modifs futures) : $WORKDIR"
