# -*- coding: utf-8 -*-
"""Dump an existing deck to Markdown, so its content can be rebuilt.

The point of this is the speaker notes. When an old deck is redesigned, the
slides are worth rethinking and the notes are worth keeping exactly — they are
the part that took the longest to write and the part nobody can reconstruct.

    python extract.py old/*.pptx              # one .md beside each .pptx
    python extract.py old/lecture1.pptx -o notes/

Each slide becomes a heading, the text on it, and its notes verbatim. Tables
come out as Markdown tables; pictures are listed by name so you know something
was there.
"""
import argparse
import os
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def cell_text(cell):
    return " ".join(cell.text.split())


def shape_lines(shape):
    """One shape's content, as Markdown-ish lines."""
    if shape.has_table:
        rows = [[cell_text(c) for c in row.cells] for row in shape.table.rows]
        if not rows:
            return []
        out = ["| " + " | ".join(rows[0]) + " |",
               "|" + "|".join(["---"] * len(rows[0])) + "|"]
        out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
        return out
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return ["_[picture: %s]_" % shape.name]
    if shape.has_text_frame and shape.text_frame.text.strip():
        return [line.rstrip() for line in shape.text_frame.text.splitlines()
                if line.strip()]
    return []


def extract(path):
    prs = Presentation(path)
    name = os.path.splitext(os.path.basename(path))[0]
    out = ["# %s" % name, "",
           "%d slides. Extracted from `%s`." % (len(prs.slides),
                                                os.path.basename(path)), ""]
    for n, slide in enumerate(prs.slides, 1):
        body = []
        for shape in slide.shapes:
            body += shape_lines(shape)
        heading = body[0] if body else "(no text)"
        out += ["---", "", "## Slide %d — %s" % (n, heading), "", "**On screen**",
                "", "```"]
        out += body or ["(nothing)"]
        out += ["```", ""]
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        out += ["**Notes**", "", notes or "_(none)_", ""]
    return "\n".join(out) + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("decks", nargs="+", help=".pptx files to read")
    ap.add_argument("-o", "--out", default=None,
                    help="directory for the .md files (default: beside each deck)")
    args = ap.parse_args()

    for path in args.decks:
        if not path.lower().endswith(".pptx"):
            print("skipping %s" % path, file=sys.stderr)
            continue
        text = extract(path)
        folder = args.out or os.path.dirname(os.path.abspath(path))
        if args.out:
            os.makedirs(args.out, exist_ok=True)
        dest = os.path.join(
            folder, os.path.splitext(os.path.basename(path))[0] + ".md")
        with open(dest, "w", encoding="utf-8") as fh:
            fh.write(text)
        print("%s  (%d slides)" % (dest, text.count("\n## Slide ")))


if __name__ == "__main__":
    main()
