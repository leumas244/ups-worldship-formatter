@echo off
:: Prüfen, ob die Batch-Datei mit Administratorrechten ausgeführt wird
:: openfiles >nul 2>&1
:: if %errorlevel% neq 0 (
    :: Starten einer powershell mit Adminrechten
    :: powershell -Command "Start-Process '%~0' -Verb runAs"
    :: exit /b
:: )

:: Zum Verzeichnis wechseln, in dem die Batch-Datei liegt
cd /d "%~dp0"

:: Hier dein Python-Skript ausführen
python main.py

:: Das Fenster offen lassen
pause