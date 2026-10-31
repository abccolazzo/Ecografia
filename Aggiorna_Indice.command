#!/bin/bash
# Doppio clic: rigenera Indice_Ecografia.html dai digest Word di questa cartella e lo apre.
cd "$(dirname "$0")" || exit 1

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 non è installato. Scaricalo da https://www.python.org/downloads/ e riprova."
  read -r -p "Premi Invio per chiudere..."
  exit 1
fi

if python3 genera_indice.py; then
  open "Indice_Ecografia.html"
  echo "Fatto. Questa finestra si chiude tra 3 secondi."
  sleep 3
  osascript -e 'tell application "Terminal" to close (every window whose name contains "Aggiorna_Indice")' >/dev/null 2>&1 &
else
  echo "Si è verificato un errore durante l'aggiornamento."
  read -r -p "Premi Invio per chiudere..."
fi
