#!/bin/bash
# build_app.sh — Build a standalone macOS .app bundle for the Payroll Exporter
#
# Usage:  bash build_app.sh
#
# Prerequisites:
#   source venv/bin/activate
#   pip install pyinstaller
#
# Output:
#   dist/Optihome Payroll Processing.app

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

APP_NAME="Optihome Payroll Processing"

# Activate venv if present
if [ -f "venv/bin/activate" ]; then
    source venv/bin/activate
fi

# Ensure pyinstaller is installed
if ! command -v pyinstaller &>/dev/null; then
    echo "Installing PyInstaller..."
    pip install pyinstaller
fi

# Clean previous build artifacts to prevent duplicates
# (iCloud Drive can create numbered copies instead of overwriting)
echo "Cleaning previous build artifacts..."
rm -rf build dist/*.app dist/"$APP_NAME"

echo "Building ${APP_NAME}.app..."

pyinstaller \
    --noconfirm \
    --onedir \
    --windowed \
    --name "$APP_NAME" \
    --add-data "export-timesheet.py:." \
    --add-data "timesheet-rates.csv:." \
    --collect-all tkinter \
    --hidden-import pandas \
    --hidden-import xlsxwriter \
    --hidden-import numpy \
    --hidden-import dateutil \
    payroll_app.py

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

echo ""
echo "Build complete!"
echo "App bundle: dist/${APP_NAME}.app"
echo ""
echo "To use: double-click the .app in Finder, or copy it to /Applications."
