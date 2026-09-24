#!/bin/sh
# Install the two open typefaces this theme uses.
#
# Both are SIL Open Font License 1.1: free to install, bundle and embed, on as
# many machines as you like. Nothing open ships with macOS or Windows, so this
# has to be run on every machine that opens or presents the decks — otherwise
# PowerPoint substitutes silently and you are not looking at the design.
#
#     sh install_fonts.sh
#
# Static cuts only. Google Fonts now ships Archivo as a variable font, which
# PowerPoint renders as Regular with a faked bold — no use for a design whose
# headlines are ExtraBold — so Archivo comes from the upstream Omnibus-Type
# repository instead. Licences are saved next to this script.
#
# On Linux they go to ~/.local/share/fonts, which fontconfig reads.
#
# On Windows: fetch the same files from those two repositories, select the
# .ttf files, right-click and choose "Install for all users".

set -e

case "$(uname -s)" in
    Darwin) DEST="$HOME/Library/Fonts" ;;
    *)      DEST="$HOME/.local/share/fonts" ;;   # where the theme looks
esac
LIC="$(cd "$(dirname "$0")" && pwd)/fonts"
ARCHIVO="https://raw.githubusercontent.com/Omnibus-Type/Archivo/master/fonts/ttf"
PLEX="https://raw.githubusercontent.com/google/fonts/main/ofl/ibmplexmono"

mkdir -p "$DEST" "$LIC"

get() {
    url="$1"
    file="$2"
    printf '  %s\n' "$file"
    curl -fsSL "$url/$file" -o "$DEST/$file"
}

echo "Archivo — Omnibus-Type"
for w in Regular Medium SemiBold Bold ExtraBold Italic BoldItalic; do
    get "$ARCHIVO" "Archivo-$w.ttf"
done

echo "IBM Plex Mono — IBM"
for w in Regular Medium SemiBold Bold Italic BoldItalic; do
    get "$PLEX" "IBMPlexMono-$w.ttf"
done

curl -fsSL "https://raw.githubusercontent.com/Omnibus-Type/Archivo/master/OFL.txt" \
     -o "$LIC/OFL-Archivo.txt"
curl -fsSL "$PLEX/OFL.txt" -o "$LIC/OFL-IBMPlexMono.txt"

# fontconfig caches what it has seen; tell it about the new files
if command -v fc-cache >/dev/null 2>&1; then
    fc-cache -f "$DEST" >/dev/null
fi

echo
echo "Installed into $DEST"
echo "Licences in $LIC"
echo "Quit and reopen PowerPoint or OnlyOffice so they pick the fonts up."
