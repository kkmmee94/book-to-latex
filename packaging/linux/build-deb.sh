#!/usr/bin/env bash
set -euo pipefail

project_root="$(cd "$(dirname "$0")/../.." && pwd)"
version="${1:-1.2.0}"
stage="$project_root/build/linux-deb/book-to-latex"

cd "$project_root"
python -m PyInstaller --noconfirm --clean --onefile --windowed \
  --name book-to-latex-gui \
  --add-data "assets/tessdata:assets/tessdata" \
  --hidden-import PIL.Image --hidden-import pytesseract \
  book_to_latex_gui.py

rm -rf "$stage"
mkdir -p "$stage/DEBIAN" "$stage/usr/bin" "$stage/usr/share/applications" \
  "$stage/usr/share/doc/book-to-latex"
install -m 0755 "$project_root/dist/book-to-latex-gui" "$stage/usr/bin/book-to-latex-gui"
install -m 0644 "$project_root/packaging/linux/book-to-latex.desktop" \
  "$stage/usr/share/applications/book-to-latex.desktop"
install -m 0644 "$project_root/README.md" "$project_root/USER_GUIDE.md" \
  "$project_root/LICENSE" "$stage/usr/share/doc/book-to-latex/"

cat > "$stage/DEBIAN/control" <<EOF
Package: book-to-latex
Version: $version
Section: text
Priority: optional
Architecture: amd64
Maintainer: Book to LaTeX contributors
Description: Friendly document-to-LaTeX desktop application
 Converts documents, PDFs, images, ebooks, slides and spreadsheets into
 LaTeX source, compiled PDF and a concise conversion report.
EOF

dpkg-deb --build --root-owner-group "$stage" \
  "$project_root/dist/book-to-latex_${version}_amd64.deb"
