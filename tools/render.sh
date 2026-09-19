#!/bin/sh
# Render a deck the way PowerPoint would, via LibreOffice.
#
#     sh render.sh 06_Selection.pptx            # PDF + PNG per slide
#     OUT=~/proof sh render.sh 06_Selection.pptx
#
# LibreOffice headless on macOS does not read the system font folders — it uses
# its own bundled fontconfig, so without the config written below it silently
# substitutes Linux Libertine for Archivo and the render tells you nothing.
#
# Output: $OUT/<name>.pdf and $OUT/<name>/sNN.png   (OUT defaults to /tmp/cudeck)
#
# preview.py is the quick check during a build; this is the one to trust.

set -e

SRC="$1"
[ -n "$SRC" ] || { echo "usage: sh render.sh <deck.pptx>"; exit 1; }

SOF=/Applications/LibreOffice.app/Contents/MacOS/soffice
[ -x "$SOF" ] || { echo "LibreOffice not found — brew install --cask libreoffice"; exit 1; }

NAME=$(basename "$SRC" .pptx)
OUT="${OUT:-/tmp/cudeck}"
mkdir -p "$OUT/fccache" "$OUT/$NAME"

cat > "$OUT/fonts.conf" <<EOF
<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
  <dir>~/Library/Fonts</dir>
  <dir>/Library/Fonts</dir>
  <dir>/System/Library/Fonts</dir>
  <dir>/System/Library/Fonts/Supplemental</dir>
  <dir>/Applications/LibreOffice.app/Contents/Resources/fonts/truetype</dir>
  <cachedir>$OUT/fccache</cachedir>
</fontconfig>
EOF

rm -f "$OUT/$NAME.pdf" "$OUT/$NAME"/s*.png
FONTCONFIG_FILE="$OUT/fonts.conf" "$SOF" --headless \
    --convert-to pdf --outdir "$OUT" "$SRC" >/dev/null 2>&1

pdftoppm -r 110 -png "$OUT/$NAME.pdf" "$OUT/$NAME/s"

echo "$OUT/$NAME.pdf"
echo "$(ls "$OUT/$NAME"/s*.png | wc -l | tr -d ' ') slides in $OUT/$NAME/"

# say which faces actually made it into the PDF — substitution is silent otherwise
python3 - "$OUT/$NAME.pdf" <<'PY'
import re, sys
d = open(sys.argv[1], "rb").read()
names = {n.decode().split("+")[-1]
         for n in re.findall(rb"/BaseFont\s*/([A-Za-z0-9+#\-,_]+)", d)}
print("fonts used:", ", ".join(sorted(names)))
PY
