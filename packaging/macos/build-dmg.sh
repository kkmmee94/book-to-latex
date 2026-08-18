#!/usr/bin/env bash
set -euo pipefail

project_root="$(cd "$(dirname "$0")/../.." && pwd)"
version="${1:-1.1.0}"
volume_root="$project_root/build/macos-dmg"
dmg_path="$project_root/dist/Book-to-LaTeX-${version}-macOS.dmg"

cd "$project_root"
python -m PyInstaller --noconfirm --clean --windowed --name "Book to LaTeX" \
  --add-data "assets/tessdata:assets/tessdata" \
  --hidden-import PIL.Image --hidden-import pytesseract \
  book_to_latex_gui.py

rm -rf "$volume_root"
mkdir -p "$volume_root"
cp -R "$project_root/dist/Book to LaTeX.app" "$volume_root/"
ln -s /Applications "$volume_root/Applications"
rm -f "$dmg_path"
hdiutil create -volname "Book to LaTeX" -srcfolder "$volume_root" \
  -ov -format UDZO "$dmg_path"
