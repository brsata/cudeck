# Working in this repository

CUDeck is a PowerPoint theme expressed as Python. `cudeck/theme.py` is the whole
theme; everything else is tooling, docs or an example. Read the README first —
it is written for the person using the theme, and most questions are answered
there.

## Environment

```bash
pip install -e '.[tools]'
sh tools/install_fonts.sh          # Archivo + IBM Plex Mono
brew install --cask libreoffice    # for tools/render.sh
```

Python 3.8+. The only hard dependency is `python-pptx`; Pillow is needed by
`tools/preview.py` and by anything that stitches contact sheets.

## Never change a deck without looking at it

A slide can be geometrically valid and visually broken, and python-pptx will not
tell you. The order is always:

```bash
python build_xx.py                  # the fit report must come back empty
sh tools/render.sh build_xx.pptx    # check the reported font list
python tools/audit.py build_xx.pptx
```

then look at the PNGs in `/tmp/cudeck/<name>/`. `tools/preview.py` is fine
between edits but it is an approximation of a renderer, not a renderer — it has
disagreed with LibreOffice about panel heights before.

If `render.sh` reports any face other than Archivo and IBM Plex Mono, a font is
missing and everything you are looking at is a substitution.

## Things that have already gone wrong

Each of these cost real time; the fix is in the code, so do not undo it.

- **Line spacing must be absolute points.** PowerPoint multiplies a ratio by the
  font's line height, not the point size. `line_spacing = 1.4` at 18 pt is about
  33 pt, not 25 — every panel came out a third too short. Use `Pt(size * LS_*)`.
- **Set the theme's fonts on the deck.** A run that does not name a typeface
  inherits the theme, which in a stock python-pptx file is Calibri. `new_deck()`
  repoints it. Without that a hand-built textbox renders in Carlito and looks
  plausible.
- **Archivo's ampersand is calligraphic.** `&&` as body text does not look like
  the operator. Prose goes through `_rich()`, which sets `` `backticks` `` in the
  code face.
- **JetBrains Mono has code ligatures.** It draws `>=` as `≥`. IBM Plex Mono has
  none, which is why it is the code face. If a mono face is ever swapped, turn
  ligatures off.
- **Google Fonts ships Archivo only as a variable font**, which PowerPoint
  renders as Regular with a faked bold. `install_fonts.sh` takes the static cuts
  from the upstream Omnibus-Type repository, which is where **Archivo ExtraBold**
  (`FONT_DISPLAY`, the headline face) comes from.

## Geometry

The content band is `MARGIN` 0.92″ in, `BODYW` 11.5″ wide, `BODY_TOP` 1.6″ to
`BODY_BOTTOM` 7.15″, with the 0.62″ spine to its left. Those numbers match the
band an earlier hand-built deck used, so a couple of hundred measured
coordinates ported without moving. Nothing should be positioned from the slide
edge; use the constants and `cols(...)`.

## How the theme reports problems

`finish()` prints a fit report and the builders add to it: code too wide or too
tall for its panel, a notes panel past the body, bullets forced below 28 pt,
panels overlapping. **Leave the report empty.** If a change makes it noisy, the
change is usually wrong — or the slide genuinely has too much on it and should
be split, which is the preferred answer.

Builders shrink to fit before they complain, within limits: code down to 11 pt,
bullets to 16, notes to 15. A warning means it hit the floor and still did not
fit.

## Style

Match what is there. The theme code is plain, comments explain *why* rather than
what, and the palette block at the top of `theme.py` marks which colours come
from the identity and which are extensions. Keep that distinction — it is the
answer to "can we do this?" from anyone in the faculty.

Course material does not belong in this repository. Chapter scripts, speaker
notes and built decks live with the course.
