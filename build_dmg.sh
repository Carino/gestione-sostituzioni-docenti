#!/bin/bash
set -e

cd "$(dirname "$0")"

# 0. Selection of Python version
if [ -x "./venv/bin/python3" ]; then
    PYTHON_CMD="./venv/bin/python3"
elif [ -f "/Library/Frameworks/Python.framework/Versions/3.13/bin/python3" ]; then
    PYTHON_CMD="/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"
elif [ -f "/usr/local/bin/python3.13" ]; then
    PYTHON_CMD="/usr/local/bin/python3.13"
elif [ -f "/opt/homebrew/bin/python3" ]; then
    PYTHON_CMD="/opt/homebrew/bin/python3"
elif [ -f "/usr/local/bin/python3.12" ]; then
    PYTHON_CMD="/usr/local/bin/python3.12"
elif [ -f "/usr/local/bin/python3.11" ]; then
    PYTHON_CMD="/usr/local/bin/python3.11"
else
    PYTHON_CMD="python3"
fi

echo "Using Python: $($PYTHON_CMD --version)"

APP_NAME="Gestione Sostituzioni"
SCRIPT_NAME="main.py"
ICON_SOURCE="alberghiero.png"
ICNS_FILE="app_icon.icns"
DMG_FINAL="Gestione Sostituzioni MacOs Processore M4.dmg"

echo "=========================================="
echo "Avvio processo di build per $APP_NAME"
echo "Target Architettura: Apple Silicon (ARM64 / M-Series)"
echo "=========================================="

# 1. Clean previous builds and installers
echo "Pulizia artefatti di build e vecchi installer..."
rm -rf dist build "$DMG_FINAL" *.dmg

# 2. Check/Creating .icns
if [ -f "$ICON_SOURCE" ]; then
    if [ ! -f "$ICNS_FILE" ]; then
        echo "Generazione $ICNS_FILE da $ICON_SOURCE..."
        mkdir -p icon.iconset
        sips -z 16 16     "$ICON_SOURCE" --out icon.iconset/icon_16x16.png > /dev/null
        sips -z 32 32     "$ICON_SOURCE" --out icon.iconset/icon_16x16@2x.png > /dev/null
        sips -z 32 32     "$ICON_SOURCE" --out icon.iconset/icon_32x32.png > /dev/null
        sips -z 64 64     "$ICON_SOURCE" --out icon.iconset/icon_32x32@2x.png > /dev/null
        sips -z 128 128   "$ICON_SOURCE" --out icon.iconset/icon_128x128.png > /dev/null
        sips -z 256 256   "$ICON_SOURCE" --out icon.iconset/icon_128x128@2x.png > /dev/null
        sips -z 256 256   "$ICON_SOURCE" --out icon.iconset/icon_256x256.png > /dev/null
        sips -z 512 512   "$ICON_SOURCE" --out icon.iconset/icon_256x256@2x.png > /dev/null
        sips -z 512 512   "$ICON_SOURCE" --out icon.iconset/icon_512x512.png > /dev/null
        sips -z 1024 1024 "$ICON_SOURCE" --out icon.iconset/icon_512x512@2x.png > /dev/null
        iconutil -c icns icon.iconset -o "$ICNS_FILE"
        rm -rf icon.iconset
        echo "$ICNS_FILE generato con successo."
    fi
else
    echo "Attenzione: $ICON_SOURCE non trovato. Verrà usata l'icona predefinita."
fi

# 3. Assicura che l'ambiente virtuale sia pronto
if [ ! -d "venv" ]; then
    echo "Creazione virtualenv..."
    $PYTHON_CMD -m venv venv
    ./venv/bin/pip install --upgrade pip
    ./venv/bin/pip install -r requirements.txt
fi

# 4. Esecuzione PyInstaller
echo "Compilazione dell'applicazione con PyInstaller (ARM64)..."
./venv/bin/pyinstaller --noconfirm --clean --windowed --onedir \
    --name "$APP_NAME" \
    --icon "$ICNS_FILE" \
    --target-arch "arm64" \
    --osx-bundle-identifier "it.gestionesostituzioni.app" \
    --add-data "alberghiero.png:." \
    --add-data "alberghiero.ico:." \
    --add-data "sostituzioni_db.sqlite:." \
    --add-data "manuale_utente.html:." \
    --collect-all "gestione_sostituzioni" \
    --collect-all "docx" \
    --collect-all "lxml" \
    --collect-all "openpyxl" \
    --hidden-import "xlrd" \
    --hidden-import "pandas" \
    "$SCRIPT_NAME"

# 5. Firma ad-hoc del bundle per Apple Silicon (macOS Sequoia/Sonoma)
echo "Firma ad-hoc del bundle .app per Apple Silicon..."
codesign --force --deep --sign - "dist/$APP_NAME.app"

# 6. Creazione staging per DMG
echo "Preparazione cartella per DMG..."
DMG_STAGING="dist/dmg_root"
rm -rf "$DMG_STAGING"
mkdir -p "$DMG_STAGING"

# Copia l'applicazione
cp -R "dist/$APP_NAME.app" "$DMG_STAGING/"

# Crea symlink ad Applicazioni per installazione drag & drop
ln -s /Applications "$DMG_STAGING/Applicazioni"

# Aggiungi guida rapida per l'utente
cat << 'EOF' > "$DMG_STAGING/COME INSTALLARE.txt"
GESTIONE SOSTITUZIONI DOCENTI v3.3 (Apple Silicon M4)
=====================================================

INSTALLAZIONE:
1. Trascina l'icona "Gestione Sostituzioni" nella cartella "Applicazioni".
2. Apri la cartella "Applicazioni" e clicca due volte su "Gestione Sostituzioni" per avviare il programma.

NOTA AL PRIMO AVVIO (Protezione macOS Gatekeeper):
Se al primo avvio macOS mostra un messaggio del tipo "impossibile aprire l'app perché proviene da uno sviluppatore non identificato":
1. Apri la cartella Applicazioni nel Finder.
2. Fai clic destro (o tieni premuto Control e fai clic) sull'icona di "Gestione Sostituzioni".
3. Seleziona "Apri" dal menu a tendina.
4. Clicca sul pulsante "Apri" nella finestra di conferma.
(Questa operazione va eseguita solo la prima volta).
EOF

# 7. Creazione file DMG compresso con hdiutil
echo "Creazione immagine disco DMG compressa..."
hdiutil create -volname "$APP_NAME Installer" \
    -srcfolder "$DMG_STAGING" \
    -ov -format UDZO \
    "$DMG_FINAL"

# 8. Pulizia staging temporaneo
rm -rf "$DMG_STAGING"

echo "=========================================="
echo "INSTALLATORE CREATO CON SUCCESSO!"
echo "File DMG: $(pwd)/$DMG_FINAL"
echo "Dimensioni: $(ls -lh "$DMG_FINAL" | awk '{print $5}')"
echo "=========================================="
