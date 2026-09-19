# -*- coding: utf-8 -*-
"""A complete deck, and the starter for a new course.

    pip install -e /path/to/cudeck      # once
    sh tools/install_fonts.sh           # once per machine
    python examples/data_structures.py  # writes 01_Example.pptx beside this file

Copy this file into a course folder, change the course() call, and start
writing slides. Every layout the theme offers appears below at least once, so
it doubles as the reference for what things look like.
"""
import os

import cudeck as cu
from cudeck import *          # noqa: F401,F403
from pptx.util import Inches  # noqa: F401 — for hand-placed geometry

# ---------------------------------------------------------------- the course
# Said once. Every title slide and the references slide read from it.
cu.course(
    code="CEN213",
    title="Data Structures",
    term="Spring 2026–2027",
    lecturer="Dr. Your Name",
    department="Computer Engineering",
    books=[
        ("Data Structures and Algorithm Analysis in C", "Mark Allen Weiss",
         "2nd edition, Pearson, 1996"),
    ],
    references_note="""
Say once, early in the term, which chapter of the book each week matches.
""",
)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "01_Example.pptx")

prs = new_deck()
chapter(prs, "Stacks and Queues")     # runs down the spine of every slide

# ---------------------------------------------------------------- the deck
title_slide(prs, "Stacks and Queues", chapter="Chapter 1", notes="""
Speaker notes live here, and they are the part of the deck that does the
teaching. Write them as if you were talking, not as a summary of the slide —
the slide is already on screen.
""")

bullets_slide(prs, "What you will be able to do after today", [
    "Say what a stack is, and what makes it different from a queue",
    "Push and pop by hand, and predict the result",
    "Choose the right one for a problem",
    ("The traps: popping an empty stack, and forgetting the wrap-around",
     0, {"color": cu.ERROR, "marker": cu.ERROR, "bold": True}),
], notes="Read the last one slowly; it is where the marks go.")

section_slide(prs, "Part one", "Stacks", notes="Twenty minutes.")

# code beside what to notice — the layout most lectures need most often
code_notes(prs, "Push and pop", """void push(Stack *s, int v) {
    s->item[s->top] = v;
    s->top = s->top + 1;
}

int pop(Stack *s) {
    s->top = s->top - 1;
    return s->item[s->top];
}""", [
    "The top index always points at the next free slot",
    "`push` writes, then moves the top",
    "`pop` moves the top back, then reads",
    ("Popping an empty stack is undefined — guard it", cu.ERROR),
], """
Trace three pushes and two pops on the board before showing the code. The order
of the two statements inside each function is the whole lesson.
""")

# a listing beside what it prints
s = content_slide(prs, "Run it and watch the top")
(a, aw), (b, bw) = cols(1.1, 1)
code = """Stack s = {.top = 0};

push(&s, 7);
push(&s, 3);

printf("%d\\n", pop(&s));
printf("%d\\n", pop(&s));"""
runs = [("output", "3\n7")]
h = max(code_height(code), output_height(runs))
code_panel(s, code, a, BODY_TOP, aw, label="Code", height=h)
output_panel(s, runs, b, BODY_TOP, bw, h)
callout(s, "Last in, first out. The 3 went in second and comes out first.")
notes(s, "Ask what a queue would have printed, before you reveal the output.")

# the pitfall pairing: wrong on the left in red, right on the right in green
s = content_slide(prs, "Guard the empty case")
dont_do(s, """int pop(Stack *s) {
    s->top = s->top - 1;
    return s->item[s->top];
}""", """int pop(Stack *s) {
    if (s->top == 0) { return -1; }
    s->top = s->top - 1;
    return s->item[s->top];
}""", BODY_TOP,
    wrong_note="On an empty stack this reads memory that is not yours.",
    right_note="One test, and the whole class of bug is gone.")
callout(s, "Every structure with a size has an empty case and a full case. "
           "Write both guards before you write the body.")
notes(s, "Ask what the equivalent full-case guard looks like for `push`.")

# a hand trace, which is the skill worth most in an exam
s = content_slide(prs, "Trace it by hand")
table(s, ["operation", "top before", "top after", "returns"],
      [[("push(7)",), ("0",), ("1",), "—"],
       [("push(3)",), ("1",), ("2",), "—"],
       [("pop()",), ("2",), ("1",), ("3",)],
       [("pop()",), ("1",), ("0",), ("7",)],
       [("pop()",), ("0",), ("0",), ("error", True, cu.ERROR)]],
      MARGIN, BODY_TOP, BODYW, colw=[4, 3, 3, 3])
callout(s, "One row per operation. Write every value down — do not do it in "
           "your head.")
notes(s, "Build the table on the board and let them check their own working.")

table_slide(prs, "Stack or queue?",
            ["Use a stack when", "Use a queue when"],
            [["The most recent item matters", "The oldest item matters"],
             ["Undo, back-buttons, recursion", "Printing, scheduling, buffering"],
             ["Depth-first search", "Breadth-first search"]],
            notes="A judgement slide. Both are correct; the question is which fits.")

bullets_slide(prs, "Recap", [
    "A stack is last in, first out",
    "A queue is first in, first out",
    "Both need an empty guard and a full guard",
], notes="Three lines. Ask for an example of each from the room.")

references_slide(prs)

finish(prs, OUT)
