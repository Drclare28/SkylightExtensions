#!/usr/bin/env bash
# setup-github-manifest.sh
# Fills in your GitHub username and repo name in manifest.github.xml
# and produces a ready-to-use manifest.github.final.xml
#
# Usage:  bash scripts/setup-github-manifest.sh
#
set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(dirname "$SCRIPT_DIR")"

read -r -p "GitHub username: " GH_USER
read -r -p "GitHub repo name (e.g. SkylightExtensions): " GH_REPO

SRC="$ROOT/manifest.github.xml"
DEST="$ROOT/manifest.github.final.xml"

sed \
  -e "s/YOUR-GITHUB-USERNAME/$GH_USER/g" \
  -e "s/YOUR-REPO-NAME/$GH_REPO/g" \
  "$SRC" > "$DEST"

BASE="https://${GH_USER}.github.io/${GH_REPO}/outlook-addin/docs"

echo ""
echo "✅  Written to: $DEST"
echo ""
echo "Add-in will be hosted at:"
echo "  $BASE/app.html"
echo ""
echo "Next steps:"
echo "  1. Push this repo to GitHub (github.com/$GH_USER/$GH_REPO)"
echo "  2. In GitHub → Settings → Pages, set source to:"
echo "       Branch: main   Folder: / (root)"
echo "  3. Wait ~2 min for GitHub Pages to build"
echo "  4. Sideload manifest.github.final.xml in Outlook Web (OWA)"
echo "     (see README.md for detailed sideloading steps)"
