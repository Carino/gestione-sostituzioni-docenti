@echo off
TITLE Gestione Sostituzioni - Windows Builder
cd /d "%~dp0"

echo ==============================================
echo      COSTRUZIONE PACCHETTO WINDOWS
echo ==============================================
echo.

echo 1. Controllo dipendenze Python...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo Installazione dipendenze fallita o pip non trovato.
    echo Assicurati di avere Python installato e aggiunto al PATH.
    pause
    exit /b
)

echo.
echo 2. Pulizia vecchie build...
if exist build rd /s /q build
if exist dist rd /s /q dist
del *.spec 2>nul

echo.
echo 3. Creazione Eseguibile (.exe) con PyInstaller...
echo Script: Gestione_docenti_refactored.py
echo Mode: OneDir (Cartella)

pyinstaller --noconfirm --clean --windowed --onedir ^
    --name "Gestione Sostituzioni" ^
    --add-data "alberghiero.png;." ^
    --add-data "alberghiero.ico;." ^
    --add-data "sostituzioni_db.sqlite;." ^
    --hidden-import "docx" ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "lxml" ^
    --icon "alberghiero.ico" ^
    "Gestione_docenti_refactored.py"

if %errorlevel% neq 0 (
    echo.
    echo ERRORE: Creazione eseguibile fallita.
    pause
    exit /b
)

echo.
echo 4. Creazione Pacchetto ZIP...
if exist "dist\Gestione Sostituzioni" (
    powershell Compress-Archive -Path "dist\Gestione Sostituzioni" -DestinationPath "GestioneSostituzioni_Windows.zip" -Force
    echo ZIP creato: GestioneSostituzioni_Windows.zip
) else (
    echo ERRORE: Cartella dist\Gestione Sostituzioni non trovata.
)

echo.
echo ==============================================
echo                COMPLETATO!
echo ==============================================
echo.
echo L'eseguibile si trova nella cartella: dist\Gestione Sostituzioni\Gestione Sostituzioni.exe
echo Il pacchetto ZIP pronto e': GestioneSostituzioni_Windows.zip
echo.
pause
