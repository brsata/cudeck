# -*- coding: utf-8 -*-
"""CUDeck — a lecture-slide theme in the Çukurova Üniversitesi identity.

    import cudeck as cu
    from cudeck import *

    cu.course(code="CEN213", title="Data Structures",
              term="Spring 2026-2027", lecturer="Dr. Your Name",
              department="Computer Engineering",
              books=[("Data Structures and Algorithm Analysis in C",
                      "Mark Allen Weiss", "2nd edition, Pearson, 1996")])

    prs = new_deck()
    chapter(prs, "Stacks and Queues")
    title_slide(prs, "Stacks and Queues", chapter="Chapter 1")
    bullets_slide(prs, "What you will be able to do after today", [...])
    references_slide(prs)
    finish(prs, "01_Stacks.pptx")

See the README for the full set of slide builders.
"""
from .theme import *          # noqa: F401,F403
from .theme import (          # noqa: F401 — used by hand-built slides
    _blank, _indent, _notes, _rect, _rich, _run, _tb,
)
from . import theme           # noqa: F401 — for cudeck.theme.SOMETHING

__version__ = "1.0.0"
