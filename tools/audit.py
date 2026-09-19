# -*- coding: utf-8 -*-
"""Check built decks for the things a fit report cannot see.

The build-time report knows what each helper asked for; this reads the finished
files and looks for what actually landed on top of what.

    python audit.py *.pptx
    python audit.py                 # every .pptx in the working directory

Overlaps below about 0.5" are usually a label box touching its neighbour — a
label() box is 0.36" tall while its text is nearer 0.2", so the boxes meet and
the words do not. Anything larger is worth looking at.
"""
import glob
import os
import sys

from pptx import Presentation

import cudeck as cu

E = 914400.0
BOTTOM = 7.30          # a shape below this is running off the foot
RIGHT = 12.95
MIN_OVERLAP = 0.12     # inches in both directions before it counts as a clash


def rect(sh):
    return (sh.left / E, sh.top / E,
            (sh.left + (sh.width or 0)) / E, (sh.top + (sh.height or 0)) / E)


def overlap(a, b):
    return (min(a[2], b[2]) - max(a[0], b[0]),
            min(a[3], b[3]) - max(a[1], b[1]))


def interesting(sh):
    """Text-bearing shapes only — a panel behind its own text is not a clash."""
    if sh.left is None or sh.top is None:
        return False
    if sh.left < cu.SPINE_W:                       # the spine and its furniture
        return False
    if (sh.height or 0) >= cu.SH - 10:             # a full-bleed ground
        return False
    return sh.has_text_frame and sh.text_frame.text.strip()


def audit(path):
    prs = Presentation(path)
    problems = []
    for n, slide in enumerate(prs.slides, 1):
        shapes = [sh for sh in slide.shapes if interesting(sh)]
        for sh in shapes:
            r = rect(sh)
            if r[3] > BOTTOM:
                problems.append((n, "off the foot", r[3],
                                 sh.text_frame.text.strip()[:40]))
            if r[2] > RIGHT:
                problems.append((n, "off the right", r[2],
                                 sh.text_frame.text.strip()[:40]))
        for i in range(len(shapes)):
            for j in range(i + 1, len(shapes)):
                a, b = rect(shapes[i]), rect(shapes[j])
                ox, oy = overlap(a, b)
                if ox > MIN_OVERLAP and oy > MIN_OVERLAP:
                    problems.append(
                        (n, "text over text", min(ox, oy),
                         "%s | %s" % (shapes[i].text_frame.text.strip()[:24],
                                      shapes[j].text_frame.text.strip()[:24])))
    return problems


def main():
    files = sys.argv[1:] or sorted(glob.glob("*.pptx"))
    if not files:
        print("no .pptx files given or found")
        return
    total = 0
    for f in files:
        problems = audit(f)
        total += len(problems)
        name = os.path.basename(f)
        if not problems:
            print("%-34s clean" % name)
            continue
        print("%-34s %d" % (name, len(problems)))
        for n, kind, amount, what in problems:
            print("    slide %2d  %-14s %5.2f\"  %s" % (n, kind, amount, what))
    print("\n%d problems across %d decks" % (total, len(files)))


if __name__ == "__main__":
    main()
