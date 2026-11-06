@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Anteprima del sito: si apre il browser su http://localhost:8000/
echo Per chiudere l'anteprima chiudi questa finestra.
start "" http://localhost:8000/
python -m http.server 8000 --bind 127.0.0.1 -d sito
