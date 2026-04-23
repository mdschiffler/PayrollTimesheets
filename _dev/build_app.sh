#!/bin/bash
# build_app.sh — Build a standalone macOS .app bundle for the Payroll Exporter
#
# Usage:  bash _dev/build_app.sh   (run from anywhere)
#
# Prerequisites:
#   source _dev/venv/bin/activate
#   pip install pyinstaller
#
# Output:
#   dist/Optihome Payroll Processing.app
#   + Finder alias at project root

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_DIR"

APP_NAME="Optihome Payroll Processing"

# Activate venv if present
if [ -f "_dev/venv/bin/activate" ]; then
    source _dev/venv/bin/activate
fi

# Ensure pyinstaller is installed
if ! command -v pyinstaller &>/dev/null; then
    echo "Installing PyInstaller..."
    pip install pyinstaller
fi

# Clean previous build artifacts to prevent duplicates
# (iCloud Drive can create numbered copies instead of overwriting)
echo "Cleaning previous build artifacts..."
rm -rf _dev/build dist/"$APP_NAME".app dist/"$APP_NAME"

echo "Building ${APP_NAME}.app..."

pyinstaller \
    --noconfirm \
    --onedir \
    --windowed \
    --name "$APP_NAME" \
    --icon _dev/AppIcon.icns \
    --workpath _dev/build \
    --distpath dist \
    --add-data "_dev/export-timesheet.py:." \
    --add-data "timesheet-rates.csv:." \
    --collect-all tkinter \
    --hidden-import pandas \
    --hidden-import xlsxwriter \
    --hidden-import numpy \
    --hidden-import dateutil \
    _dev/payroll_app.py

# Code-sign the .app bundle.
# iCloud Drive adds resource forks that break codesign, so copy to /tmp first.
echo "Signing the app..."
TEMP_APP="/tmp/${APP_NAME}.app"
rm -rf "$TEMP_APP"
cp -R "dist/${APP_NAME}.app" "$TEMP_APP"
xattr -cr "$TEMP_APP"
find "$TEMP_APP" -name ".DS_Store" -delete
codesign -s - --force --all-architectures --timestamp --deep "$TEMP_APP"
rm -rf "dist/${APP_NAME}.app"
cp -R "$TEMP_APP" "dist/${APP_NAME}.app"
rm -rf "$TEMP_APP"

# Create a macOS Finder alias at the project root so the user can
# double-click the app without navigating into dist/
echo "Creating Finder alias at project root..."
osascript -e "
    tell application \"Finder\"
        set appFile to POSIX file \"${PROJECT_DIR}/dist/${APP_NAME}.app\" as alias
        set projectFolder to POSIX file \"${PROJECT_DIR}\" as alias
        try
            delete file \"${APP_NAME}.app\" of folder projectFolder
        end try
        make new alias file at folder projectFolder to appFile with properties {name:\"${APP_NAME}.app\"}
    end tell
" 2>/dev/null || echo "  (Finder alias creation skipped — run from Finder or grant Finder access)"

echo ""
echo "Build complete!"
echo "App bundle: dist/${APP_NAME}.app"
echo "Alias:      ${APP_NAME}.app (at project root)"
echo ""
echo "To use: double-click the .app alias in Finder."
