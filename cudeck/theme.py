# -*- coding: utf-8 -*-
"""A lecture-slide theme in the Çukurova University identity.

Nothing here knows what the subject is. A deck describes its course once and
the theme takes care of the rest.
A deck starts by describing its course —

    import cudeck as cu
    from cudeck import *

    cu.course(code="CEN213", title="Data Structures",
              term="Spring 2026-2027", lecturer="Dr. Your Name",
              department="Computer Engineering",
              books=[("Data Structures and Algorithm Analysis in C",
                      "Mark Allen Weiss", "2nd edition, Pearson, 1996")])

    prs = new_deck()
    chapter(prs, "Stacks and Queues")  # runs down the spine of every slide
    title_slide(prs, "Stacks and Queues", chapter="Chapter 1")
    ...
    references_slide(prs)              # built from the books above
    finish(prs, "01_Stacks.pptx")

— and everything that would otherwise be repeated on every title and
references slide comes from that one call.

Colours and the two logo lockups come from the university's identity manual,
Çukurova Üniversitesi Kurumsal Kimlik (2015). The pale panel fill, the lifted
green, the error red and both typefaces are extensions of that document.

Geometry is fixed by the spine down the left edge: content starts at MARGIN
and is BODYW wide. Nothing should be positioned from the slide edge.

Typefaces are open (SIL OFL 1.1) and must be installed on the presenting
machine — run tools/install_fonts.sh once per machine.
"""
import os

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- palette
# One green, and it is this one. Everything else is neutral or a signal.
GREEN      = RGBColor(0x00, 0x42, 0x1C)   # Pantone 357 C
INK        = RGBColor(0x23, 0x1F, 0x20)   # Pantone PRO Black C
GREY       = RGBColor(0x5E, 0x5C, 0x5B)   # Pantone 424 C
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

# derived, not in the identity book
GREEN_LIFT = RGBColor(0x8F, 0xB7, 0x9C)   # green that reads on a green ground
GREEN_PALE = RGBColor(0xC8, 0xDA, 0xCD)
PANEL      = RGBColor(0xF4, 0xF6, 0xF2)   # panel fill on white
PANEL_HEAD = RGBColor(0xE7, 0xED, 0xE3)   # panel label strip
PANEL_LINE = RGBColor(0xD3, 0xDB, 0xCE)   # panel border
RULE       = RGBColor(0xE4, 0xE7, 0xE0)   # table row rule
LINENO     = RGBColor(0x93, 0xA3, 0x98)

# error red — off-book, reserved for wrong code and nothing else
ERROR      = RGBColor(0xA3, 0x2B, 0x1C)
ERROR_FILL = RGBColor(0xFB, 0xF1, 0xEF)
ERROR_HEAD = RGBColor(0xF4, 0xE1, 0xDD)
ERROR_LINE = RGBColor(0xE3, 0xC4, 0xBE)

# Older names, kept so that decks written against an earlier version of this
# theme still build. New work should use the names above.
ACCENT, WARN, GOOD, CODEFG = GREEN, ERROR, GREEN, INK
CODEBG, DARK = PANEL, INK

# Line numbers are on for C listings: they are what lets you say "look at line
# three" instead of pointing. A chapter whose listings are pseudocode already
# numbers its own steps and turns this off.
CODE_NUMBERS = [True]

# ---------------------------------------------------------------- type
FONT      = "Archivo"
# Headlines: the weight is the point of them.
FONT_DISPLAY = "Archivo ExtraBold"
MARKER_FONT = None      # \u25aa comes from the body face — Archivo has it, and
                        # it is the same file on every machine we install to
MONO      = "IBM Plex Mono"
CODEFONT  = MONO

T_TITLE      = 80    # title-slide headline (stepped down for long titles)
T_SECTION    = 60    # section-divider headline
T_HEAD       = 40    # content-slide title
T_BULLET     = 28    # body bullets — the back-row size
T_NOTE       = 20    # bullets inside a panel
T_CODE       = 18
T_LABEL      = 12    # panel labels, footer
T_SPINE      = 22    # the chapter title running down the spine
T_TABLE      = 21
T_TABLE_HEAD = 15
T_CALLOUT    = 19

# ---------------------------------------------------------------- geometry
SW, SH   = Inches(13.333), Inches(7.5)
# The spine is a slim rail rather than a wide bar, which leaves the content
# band almost the full width of the slide: 0.92" in, 11.5" wide. Those are
# also the margins of an ordinary 16:9 deck, so material moved in from one
# lands in the right place.
SPINE_W  = Inches(0.62)
MARGIN   = Inches(0.92)
RMARGIN  = Inches(0.91)
BODYW    = SW - MARGIN - RMARGIN
TITLE_TOP   = Inches(0.50)
TITLE_H     = Inches(0.72)
BODY_TOP    = Inches(1.60)
BODY_BOTTOM = Inches(7.15)   # a 0.35" foot: enough to breathe, and it buys
                             # back a line of body text
BODY_H      = BODY_BOTTOM - BODY_TOP
GUTTER   = Inches(0.24)

# Line spacing is set in absolute points everywhere, never as a multiple:
# PowerPoint multiplies a ratio by the *font's* line height, not the point
# size, so a ratio makes the height formulas below wrong by ~30%.
LS_CODE   = 1.40           # of the point size
LS_NOTE   = 1.25
LS_BULLET = 1.26
LS_OUT    = 1.35
PANEL_PAD = Inches(0.20)
LABEL_H  = Inches(0.34)


# ================================================================ measuring
# If the real faces are installed, measure with them; otherwise fall back to a
# per-character estimate so a build still works on a machine without fonts.
_FONT_DIRS = [os.path.expanduser("~/Library/Fonts"), "/Library/Fonts",
              "C:/Windows/Fonts"]
_FACES = {("sans", False): "Archivo-Regular.ttf",
          ("sans", True): "Archivo-Bold.ttf",
          ("display", False): "Archivo-ExtraBold.ttf",
          ("display", True): "Archivo-ExtraBold.ttf",
          ("mono", False): "IBMPlexMono-Regular.ttf",
          ("mono", True): "IBMPlexMono-Bold.ttf"}
_metrics = {}


def _face(kind, bold, pt):
    key = (kind, bold, round(pt, 1))
    if key in _metrics:
        return _metrics[key]
    face = None
    try:
        from PIL import ImageFont
        for d in _FONT_DIRS:
            path = os.path.join(d, _FACES[(kind, bold)])
            if os.path.exists(path):
                face = ImageFont.truetype(path, max(4, int(round(pt * 4))))
                break
    except Exception:
        face = None
    _metrics[key] = face
    return face


_NARROW = set("iljtIf.,;:'!|()[]{}rt ")
_WIDE = set("mwMW@%")


def text_w(s, pt, mono=False, bold=False, kind=None):
    """Rendered width in EMU — measured when the font is installed."""
    face = _face(kind or ("mono" if mono else "sans"), bold, pt)
    if face is not None:
        # measured at 4x the point size, so 1 unit = 1/4 pt
        return int(face.getlength(s) / 4.0 * 12700)
    if mono:
        return int(len(s) * pt * 0.60 * 12700)
    em = 0.0
    for ch in s:
        if ch in _NARROW:
            em += 0.29
        elif ch in _WIDE:
            em += 0.86
        elif ch.isupper():
            em += 0.64
        elif ch.isdigit():
            em += 0.56
        else:
            em += 0.52
    return int(em * pt * 12700)


def wrapped_lines(s, pt, width, mono=False, bold=False, kind=None):
    """How many lines `s` needs at `pt` inside `width` EMU."""
    if not s:
        return 1
    words, lines, cur = s.split(" "), 1, 0
    space = text_w(" ", pt, mono, bold, kind)
    for wd in words:
        ww = text_w(wd, pt, mono, bold, kind)
        if cur and cur + space + ww > width:
            lines += 1
            cur = ww
        else:
            cur += (space if cur else 0) + ww
    return lines


# ================================================================ the course
class Course(object):
    """What every deck in a course repeats: who is teaching what, and when."""

    def __init__(self, code="", title="", term="", lecturer="", department="",
                 books=(), books_note="", references_note=""):
        self.code = code
        self.title = title
        self.term = term
        self.lecturer = lecturer
        self.department = department
        self.books = list(books)
        self.books_note = books_note
        self.references_note = references_note

    def heading(self):
        """The line that runs above a title-slide headline."""
        return " \u00b7 ".join(x for x in (self.code, self.title) if x)

    def byline(self):
        parts = (self.lecturer, self.department)
        return " \u00b7 ".join(x for x in parts if x)

    def subtitle(self, chapter=None):
        """The three lines under a title-slide headline."""
        second = " \u00b7 ".join(x for x in (chapter, self.term) if x)
        return [self.heading(), second, self.byline()]


# held in a list so that `from .theme import *` re-exports a live reference
# rather than a snapshot taken at import time
_COURSE = [Course()]


def course(**kw):
    """Describe the course these decks belong to. Call once, up front."""
    _COURSE[0] = Course(**kw)
    return _COURSE[0]


def current_course():
    """The course these decks belong to."""
    return _COURSE[0]


# ================================================================ fit report
WARNINGS = []


def warn(msg):
    WARNINGS.append(msg)


_CONTEXT = [""]
_PANELS = {}


def where(title):
    _CONTEXT[0] = title


# ================================================================ scaffolding
def new_deck():
    prs = Presentation()
    prs.slide_width, prs.slide_height = SW, SH
    prs._cu_spines = []
    prs._cu_warnings = []
    prs._cu_chapter_label = None      # set by chapter(); see _spine()
    _theme_fonts(prs)
    return prs


def _theme_fonts(prs):
    """Point the theme's major/minor fonts at ours.

    Any run that does not name a typeface inherits the theme, which in a stock
    python-pptx deck is Calibri. Any slide laid out by hand tends to have a few
    such runs, and the substitution is silent — the deck renders in Carlito and
    looks very nearly right.
    """
    from lxml import etree
    for part in prs.part.package.iter_parts():
        if "theme" not in str(part.partname):
            continue
        root = etree.fromstring(part.blob)
        changed = False
        for tag in ("a:majorFont", "a:minorFont"):
            grp = root.find(".//" + qn(tag))
            if grp is None:
                continue
            latin = grp.find(qn("a:latin"))
            if latin is not None:
                latin.set("typeface", FONT)
                changed = True
        if changed:
            part._blob = etree.tostring(root, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)


def _blank(prs):
    return prs.slides.add_slide(prs.slide_masters[0].slide_layouts[6])


def _notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text.strip()


notes = _notes


def _rect(slide, left, top, width, height, fill=None, line=None, line_pt=0.75,
          shape=MSO_SHAPE.RECTANGLE):
    shp = slide.shapes.add_shape(shape, int(left), int(top),
                                 int(width), int(height))
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_pt)
    shp.shadow.inherit = False
    style = shp._element.find(qn("p:style"))
    if style is not None:
        shp._element.remove(style)
    return shp


def _tb(slide, left, top, width, height, wrap=True, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tb


def _run(p, text, size, color=INK, bold=False, mono=False, font=None):
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font or (MONO if mono else FONT)
    return r


def _rich(p, text, size, color=INK, bold=False):
    """Write `text`, setting `backticked` spans in the code face.

    Archivo's ampersand is calligraphic, so `&&` in prose does not look like
    the operator students type. Anything that is C belongs in the code face.
    """
    for i, part in enumerate(text.split("`")):
        if not part:
            continue
        code = i % 2 == 1
        _run(p, part, size * (0.94 if code else 1.0), color,
             bold=bold and not code, mono=code)


def plain(text):
    """The text as measured — backticks are markup, not characters."""
    return text.replace("`", "")


def _fit_title(text, base, floor, width, lines=1):
    """Step a headline down until it fits.

    Content-slide titles want one line: a second line eats into the body area,
    which is a fixed band, and the diagrams that sit at the top of it have
    nowhere to go. Only when one line is impossible at the floor size does a
    second line get used.
    """
    size = base
    while (size > floor
           and wrapped_lines(text, size, width, kind="display") > lines):
        size -= 2
    return size


# ================================================================ the spine
def _spine(prs, slide):
    bar = _rect(slide, 0, 0, SPINE_W, SH, fill=GREEN)
    bar.shadow.inherit = False

    # the title is centred on the spine: the box spans the bar, clear of the
    # progress tick and the slide number at the foot
    tb = _tb(slide, 0, Inches(0.5), SPINE_W, SH - Inches(1.5), wrap=False)
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    # a deck that never called chapter() still needs something on its spine;
    # the course title is right more often than a blank
    label = prs._cu_chapter_label or current_course().title
    _run(p, label, T_SPINE, WHITE, bold=True)
    _vertical(tb)

    trough = _rect(slide, Inches(0.15), SH - Inches(0.95), Inches(0.32),
                   Inches(0.055), fill=RGBColor(0x4D, 0x7C, 0x5F))
    prog = _rect(slide, Inches(0.15), SH - Inches(0.95), Inches(0.32),
                 Inches(0.055), fill=WHITE)

    num = _tb(slide, 0, SH - Inches(0.62), SPINE_W, Inches(0.34), wrap=False)
    p = num.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _run(p, "00", T_LABEL + 1, WHITE, bold=True)

    prs._cu_spines.append((slide, prog, num))
    return bar


def _vertical(textbox):
    """Rotate a textbox to read bottom-to-top."""
    bodyPr = textbox.text_frame._txBody.bodyPr
    bodyPr.set("vert", "vert270")


def finish(prs, out, verbose=True):
    """Number the spines, set the progress bars, save."""
    index = {id(s): i for i, s in enumerate(prs.slides, start=1)}
    total = len(prs.slides)
    tol = Inches(0.02)
    for slide in prs.slides:
        boxes = _PANELS.get(id(slide), [])
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                ox = min(a[2], b[2]) - max(a[0], b[0])
                oy = min(a[3], b[3]) - max(a[1], b[1])
                if ox > tol and oy > tol:
                    warn("%s: panels \u201c%s\u201d and \u201c%s\u201d "
                         "overlap by %.2f\""
                         % (a[5], a[4], b[4], oy / 914400.0))
    for slide, prog, num in prs._cu_spines:
        n = index[id(slide)]
        num.text_frame.paragraphs[0].runs[0].text = "%02d" % n
        prog.width = max(Inches(0.04), int(Inches(0.32) * n / total))
    # a bare filename has no folder to make, and makedirs("") raises
    if os.path.dirname(out):
        os.makedirs(os.path.dirname(out), exist_ok=True)
    prs.save(out)
    if verbose:
        print("saved: %s" % out)
        print("slides: %d" % total)
        for w in prs._cu_warnings + WARNINGS:
            print("  fit: %s" % w)
    del WARNINGS[:]
    _PANELS.clear()
    return out


def chapter(prs, label):
    """Set the text that runs down every spine — the chapter title."""
    prs._cu_chapter_label = label


# ================================================================ slide types
def title_slide(prs, title, subtitle=None, notes=None, chapter_label=None,
                chapter=None):
    """Full-bleed green opener.

    Pass `chapter="Chapter 6"` and the three lines under the headline are built
    from the course. Pass `subtitle` as a list of lines to say something else.
    """
    if chapter_label:
        globals()["chapter"](prs, chapter_label)
    if subtitle is None:
        subtitle = current_course().subtitle(chapter)
    s = _blank(prs)
    _rect(s, 0, 0, SW, SH, fill=GREEN)
    logo(s, Inches(0.80), Inches(0.72), Inches(2.95), reversed_=True)

    size = _fit_title(title, T_TITLE, 44, SW - Inches(1.6), lines=2)

    # the headline is bottom-anchored: a two-line title grows upwards, so the
    # rule and the details below it stay where they are
    lines = wrapped_lines(title, size, SW - Inches(1.6), kind="display")
    title_h = int(Pt(size * 0.98) * lines)
    title_top = Inches(5.30) - title_h

    tb = _tb(s, Inches(0.80), title_top - Inches(0.46),
             SW - Inches(1.6), Inches(0.44))
    _run(tb.text_frame.paragraphs[0], subtitle[0] if subtitle else "",
         T_LABEL + 3, GREEN_LIFT, bold=True)

    tb = _tb(s, Inches(0.80), title_top, SW - Inches(1.6), title_h)
    p = tb.text_frame.paragraphs[0]
    p.line_spacing = Pt(size * 0.98)
    _run(p, title, size, WHITE, font=FONT_DISPLAY)

    rest = subtitle[1:] if len(subtitle) > 1 else []
    if rest:
        _rect(s, Inches(0.80), Inches(5.62), SW - Inches(1.6), Pt(1.2),
              fill=RGBColor(0x4D, 0x7C, 0x5F))
        tb = _tb(s, Inches(0.80), Inches(5.88), SW - Inches(1.6), Inches(1.1))
        for i, line in enumerate(rest):
            p = (tb.text_frame.paragraphs[0] if i == 0
                 else tb.text_frame.add_paragraph())
            p.space_after = Pt(4)
            _run(p, line, 20 if i == 0 else 16,
                 GREEN_PALE if i == 0 else GREEN_LIFT)
    _notes(s, notes)
    return s


def section_slide(prs, kicker, title, notes=None):
    """Full-bleed green divider: the part number and what the part is about."""
    s = _blank(prs)
    _rect(s, 0, 0, SW, SH, fill=GREEN)
    _rect(s, Inches(0.80), Inches(0.80), Inches(1.2), Inches(0.09), fill=WHITE)

    top = Inches(3.05)
    tb = _tb(s, Inches(0.80), top, SW - Inches(1.6), Inches(0.44))
    _run(tb.text_frame.paragraphs[0], kicker, T_LABEL + 3, GREEN_LIFT,
         bold=True)

    size = _fit_title(title, T_SECTION, 36, SW - Inches(1.6), lines=2)
    tb = _tb(s, Inches(0.80), top + Inches(0.5), SW - Inches(1.6), Inches(1.5))
    p = tb.text_frame.paragraphs[0]
    p.line_spacing = 1.0
    _run(p, title, size, WHITE, font=FONT_DISPLAY)
    _notes(s, notes)
    return s


BADGE_W = Inches(2.3)


def content_slide(prs, title, notes=None, badge_text=None):
    """White slide with the spine and a title. Everything else sits on top."""
    s = _blank(prs)
    _spine(prs, s)
    where(title)
    tw = BODYW - (BADGE_W + Inches(0.3) if badge_text else 0)
    size = _fit_title(title, T_HEAD, 26, tw)
    if badge_text:
        badge(s, badge_text, MARGIN + BODYW - BADGE_W,
              TITLE_TOP + Inches(0.17), color=ERROR, width=BADGE_W)
    tb = _tb(s, MARGIN, TITLE_TOP, tw, TITLE_H, anchor=MSO_ANCHOR.MIDDLE)
    p = tb.text_frame.paragraphs[0]
    p.line_spacing = 1.0
    _run(p, title, size, INK, font=FONT_DISPLAY)
    _notes(s, notes)
    return s


blank_titled = content_slide


def bullets_slide(prs, title, items, notes=None, size=T_BULLET, top=BODY_TOP,
                  floor=24):
    """Body bullets.  items: str, or (text, level), or (text, level, opts)."""
    s = content_slide(prs, title)
    norm = []
    for it in items:
        opts = {}
        if isinstance(it, tuple):
            text, lvl = (it[0], it[1]) if len(it) >= 2 else (it[0], 0)
            if len(it) == 3:
                opts = it[2]
        else:
            text, lvl = it, 0
        norm.append((text, lvl, opts))

    size = _fit_bullets(prs, title, norm, size, floor, top)
    bullet_list(s, norm, MARGIN, top, BODYW, size=size)
    _notes(s, notes)
    return s


def _fit_bullets(prs, title, norm, size, floor, top):
    avail = BODY_BOTTOM - top
    while True:
        h = 0
        for text, lvl, opts in norm:
            sz = opts.get("size", size if lvl == 0 else size - 4)
            indent = Inches(0.45) * lvl
            lines = wrapped_lines(plain(text), sz,
                                  BODYW - indent - Inches(0.42))
            h += int(Pt(sz * LS_BULLET) * lines) + int(Pt(sz) * 0.62)
        # stop at the floor even if it still does not fit; bullet_list then
        # reports the overrun, which is the honest answer
        if h <= avail or size - 2 < floor:
            break
        size -= 2
    if size < T_BULLET:
        prs._cu_warnings.append(
            "'%s' dropped to %d pt — consider splitting the slide"
            % (title, size))
    return size


def bullets_height(items, width, size):
    need = 0
    for it in items:
        text = it[0] if isinstance(it, tuple) else it
        opts = it[2] if isinstance(it, tuple) and len(it) == 3 else {}
        sz = opts.get("size", size)
        need += int(Pt(sz * LS_BULLET) *
                    wrapped_lines(plain(text), sz, width - Inches(0.42)))
        need += int(Pt(sz) * 0.62)
    return need


def bullet_list(slide, items, left, top, width, size=T_BULLET, color=INK,
                marker=GREEN, floor=None, bottom=None):
    """Square-marker bullets. items as normalised by bullets_slide."""
    limit = BODY_BOTTOM if bottom is None else bottom
    if floor:
        while (size > floor
               and top + bullets_height(items, width, size) > limit):
            size -= 1
    need = 0
    for it in items:
        text = it[0] if isinstance(it, tuple) else it
        opts = it[2] if isinstance(it, tuple) and len(it) == 3 else {}
        sz = opts.get("size", size)
        need += int(Pt(sz * LS_BULLET) *
                    wrapped_lines(plain(text), sz, width - Inches(0.42)))
        need += int(Pt(sz) * 0.62)
    if top + need > BODY_BOTTOM + Inches(0.05):
        warn("%s: bullets run %.2f\" past the foot of the slide"
             % (_CONTEXT[0], (top + need - BODY_BOTTOM) / 914400.0))
    # size the box to the text, not to the band: an over-tall box reads as a
    # collision to anything checking the finished file
    tb = _tb(slide, left, top, width, min(need, BODY_BOTTOM - top))
    tf = tb.text_frame
    for i, it in enumerate(items):
        if isinstance(it, tuple):
            text, lvl = it[0], (it[1] if len(it) >= 2 else 0)
            opts = it[2] if len(it) == 3 else {}
        else:
            text, lvl, opts = it, 0, {}
        sz = opts.get("size", size if lvl == 0 else size - 4)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(sz * 0.62)
        p.line_spacing = Pt(sz * LS_BULLET)
        if lvl:
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(int(Inches(0.45) * lvl)))
            pPr.set("indent", "0")
        if not text:
            continue
        _run(p, "▪  ", sz * 0.62, opts.get("marker", marker), bold=True,
             font=MARKER_FONT)
        if opts.get("mono"):
            _run(p, plain(text), sz, opts.get("color", color),
                 bold=opts.get("bold", False), mono=True)
        else:
            _rich(p, text, sz, opts.get("color", color),
                  bold=opts.get("bold", False))
    return tb


# ================================================================ panels
def _tone(tone):
    if tone == "error":
        return ERROR_FILL, ERROR_LINE, ERROR_HEAD, ERROR
    if tone == "plain":
        return WHITE, PANEL_LINE, PANEL_HEAD, GREEN
    return PANEL, PANEL_LINE, PANEL_HEAD, GREEN


def panel(slide, left, top, width, height, label=None, tone="normal"):
    """Bordered panel with an optional label strip. Returns the inner box."""
    _PANELS.setdefault(id(slide), []).append(
        (int(left), int(top), int(left + width), int(top + height),
         label or "-", _CONTEXT[0]))
    fill, line, head, ink = _tone(tone)
    _rect(slide, left, top, width, height, fill=fill, line=line)
    inner_top = top
    if label:
        _rect(slide, left, top, width, LABEL_H, fill=head, line=line)
        marker = _rect(slide, left + Inches(0.16), top + Inches(0.125),
                       Inches(0.09), Inches(0.09), fill=ink)
        tb = _tb(slide, left + Inches(0.36), top + Inches(0.045),
                 width - Inches(0.5), LABEL_H, wrap=False)
        p = tb.text_frame.paragraphs[0]
        r = _run(p, label.upper(), T_LABEL, ink, bold=True, mono=True)
        r.font._rPr.set("spc", "140")
        inner_top = top + LABEL_H
    return (left + PANEL_PAD, inner_top + PANEL_PAD,
            width - 2 * PANEL_PAD, top + height - inner_top - 2 * PANEL_PAD)


def code_height(code, label=True, size=T_CODE, pad=True):
    n = len(code.rstrip("\n").split("\n"))
    h = int(Pt(size * LS_CODE) * n)
    if pad:
        h += 2 * PANEL_PAD
    if label:
        h += LABEL_H
    return h


def code_panel(slide, code, left, top, width, label="Code", tone="normal",
               size=T_CODE, numbers=True, height=None):
    """Code in a Şema panel: label, line-number gutter, monospace body."""
    lines = code.rstrip("\n").split("\n")
    fill, line, head, ink = _tone(tone)
    inner = width - 2 * PANEL_PAD
    widest = max(text_w(ln, 100, mono=True) for ln in lines) / 100.0

    def fits(avail, start=size):
        sz = start
        while sz > 11 and widest * sz > avail:
            sz -= 1
        return sz

    gutter = 0
    if numbers:
        gutter = Inches(0.30 if len(lines) < 10 else 0.42) + Inches(0.30)
        # the gutter is a convenience; readable code is not. If numbering costs
        # a size the back row would struggle with, drop the numbers instead.
        if fits(inner - gutter) < 13 <= fits(inner):
            numbers, gutter = False, 0
    size = fits(inner - gutter)
    if height is None:
        while size > 10 and top + code_height(code, label=bool(label),
                                              size=size) > BODY_BOTTOM:
            size -= 1
    h = height or code_height(code, label=bool(label), size=size)
    if top + h > BODY_BOTTOM + Inches(0.02):
        warn("%s: code panel runs %.2f\" past the foot of the slide"
             % (_CONTEXT[0], (top + h - BODY_BOTTOM) / 914400.0))
    bx, by, bw, bh = panel(slide, left, top, width, h, label=label, tone=tone)

    gut = Inches(0.0)
    if numbers:
        gut = Inches(0.30 if len(lines) < 10 else 0.42)
        tb = _tb(slide, bx, by, gut, bh, wrap=False)
        tf = tb.text_frame
        for i in range(len(lines)):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.RIGHT
            p.space_after = Pt(0)
            p.line_spacing = Pt(size * LS_CODE)
            _run(p, str(i + 1), size - 1, LINENO, mono=True)
        _rect(slide, bx + gut + Inches(0.13), by + Inches(0.02),
              Pt(0.75), bh - Inches(0.04), fill=line)
        gut += Inches(0.30)

    # Shrink to fit rather than let a line run out of its panel. Below 13 pt
    # that is worth knowing about, so say so.
    avail = bw - gut
    widest = lambda sz: max(text_w(ln, sz, mono=True) for ln in lines)
    while size > 11 and widest(size) > avail:
        size -= 1
    if widest(size) > avail:
        warn("%s: code is %.2f\" too wide even at %d pt"
             % (_CONTEXT[0], (widest(size) - avail) / 914400.0, size))
    elif size < 12:
        warn("%s: code shrank to %d pt to fit \u2014 worth splitting"
             % (_CONTEXT[0], size))

    tb = _tb(slide, bx + gut, by, bw - gut, bh, wrap=False)
    tf = tb.text_frame
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(0)
        p.line_spacing = Pt(size * LS_CODE)
        _run(p, ln if ln else " ", size, INK, mono=True)
    return h


def code_two_col(slide, code, top, left=None, width=None, size=T_CODE,
                 label="Code", gap=None, tone="normal"):
    """A long listing set in two columns, so it stays readable.

    A full program does not fit one panel at a size the back row can read.
    Splitting it at a blank line near the middle keeps it on one slide.
    """
    left = MARGIN if left is None else left
    width = BODYW if width is None else width
    gap = GUTTER if gap is None else gap
    lines = code.rstrip("\n").split("\n")
    mid = len(lines) // 2
    cut = mid
    for d in range(0, mid):
        for j in (mid - d, mid + d):
            if 0 < j < len(lines) and not lines[j].strip():
                cut = j
                break
        else:
            continue
        break
    a, b = "\n".join(lines[:cut]), "\n".join(lines[cut:]).lstrip("\n")
    cw = int((width - gap) / 2)
    h = max(code_height(a, size=size), code_height(b, size=size))
    code_panel(slide, a, left, top, cw, label=label, size=size, tone=tone,
               height=h, numbers=False)
    code_panel(slide, b, left + cw + gap, top, cw, label="\u2026 continued",
               size=size, tone=tone, height=h, numbers=False)
    return h


def notes_height(items, width, size=T_NOTE, label=True):
    """Height a notes_panel needs for `items` at `width`."""
    inner = width - 2 * PANEL_PAD - Inches(0.34)
    h = 0
    for it in items:
        text = plain(it[0] if isinstance(it, tuple) else it)
        h += int(Pt(size * LS_NOTE) * wrapped_lines(plain(text), size, inner))
        h += int(Pt(size) * 0.58)
    h += 2 * PANEL_PAD + (LABEL_H if label else 0)
    return h


def output_height(runs, size=T_CODE, label=True):
    h = 0
    for cap, printed in runs:
        h += int(Pt((size - 4) * LS_OUT)) + int(Pt(12))
        h += int(Pt(size * LS_OUT) * len(printed.split("\n")))
    h += 2 * PANEL_PAD + (LABEL_H if label else 0)
    return h


def notes_panel(slide, items, left, top, width, height, label="What to notice",
                tone="plain", size=T_NOTE):
    # step the note size down before complaining — a panel is allowed to be
    # tighter than the default when the slide is already full
    while size > 15 and notes_height(items, width, size=size,
                                     label=bool(label)) > height:
        size -= 1
    need = notes_height(items, width, size=size, label=bool(label))
    if need > height:
        warn("%s: notes panel is %.2f\" short" % (_CONTEXT[0],
                                                  (need - height) / 914400.0))
    if top + height > BODY_BOTTOM + Inches(0.02):
        warn("%s: notes panel runs %.2f\" past the body area"
             % (_CONTEXT[0], (top + height - BODY_BOTTOM) / 914400.0))
    bx, by, bw, bh = panel(slide, left, top, width, height,
                           label=label, tone=tone)
    tb = _tb(slide, bx, by, bw, bh)
    tf = tb.text_frame
    for i, it in enumerate(items):
        text, color, bold = it, INK, False
        if isinstance(it, tuple):
            text, color = it[0], it[1]
            bold = True
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(size * 0.58)
        p.line_spacing = Pt(size * LS_NOTE)
        _run(p, "▪  ", size * 0.6, color if bold else GREEN, bold=True,
             font=MARKER_FONT)
        _rich(p, text, size, color, bold=bold)
    return height


def defs_height(pairs, width, size=T_NOTE, label=True):
    inner = width - 2 * PANEL_PAD
    h = 0
    for head, body in pairs:
        h += int(Pt((size - 2) * LS_NOTE)) + int(Pt(2))
        h += int(Pt(size * LS_NOTE)
                 * wrapped_lines(plain(body), size - 2, inner))
        h += int(Pt(size * 0.6))
    h += 2 * PANEL_PAD + (LABEL_H if label else 0)
    return h


def defs_panel(slide, pairs, left, top, width, height, label="Each part, once",
               size=T_NOTE, tone="plain"):
    """Term in the code face, then what it means: [(head, body), ...]."""
    bx, by, bw, bh = panel(slide, left, top, width, height,
                           label=label, tone=tone)
    tb = _tb(slide, bx, by, bw, bh)
    tf = tb.text_frame
    for i, (head, body) in enumerate(pairs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(2)
        p.line_spacing = Pt((size - 2) * LS_NOTE)
        _run(p, head, size - 2, GREEN, bold=True, mono=True)
        p2 = tf.add_paragraph()
        p2.space_after = Pt(size * 0.6)
        p2.line_spacing = Pt(size * LS_NOTE)
        _rich(p2, body, size - 2, INK)
    return height


def output_panel(slide, runs, left, top, width, height,
                 label="What it prints", size=T_CODE):
    """runs: [(caption, printed_text), ...] — one transcript per input."""
    if top + height > BODY_BOTTOM + Inches(0.02):
        warn("%s: output panel runs past the body area" % _CONTEXT[0])
    bx, by, bw, bh = panel(slide, left, top, width, height,
                           label=label, tone="plain")
    tb = _tb(slide, bx, by, bw, bh)
    tf = tb.text_frame
    first = True
    for cap, printed in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        if not first:
            p.space_before = Pt(12)
        first = False
        p.space_after = Pt(3)
        _run(p, cap, size - 4, GREY, mono=True)
        for ln in printed.split("\n"):
            p = tf.add_paragraph()
            p.space_after = Pt(0)
            p.line_spacing = Pt(size * LS_OUT)
            _run(p, ln if ln else " ", size, INK, mono=True)
    return height


def dont_do(slide, wrong, right, top, wrong_label="Don’t", right_label="Do",
            size=T_CODE, wrong_note=None, right_note=None):
    """The pitfall pairing: wrong in red, right in green, side by side."""
    (l1, w1), (l2, w2) = cols(1, 1)
    h = max(code_height(wrong, size=size), code_height(right, size=size))
    code_panel(slide, wrong, l1, top, w1, label=wrong_label, tone="error",
               size=size, height=h)
    code_panel(slide, right, l2, top, w2, label=right_label, tone="normal",
               size=size, height=h)
    y = top + h + Inches(0.14)
    if wrong_note:
        _note_line(slide, wrong_note, l1, y, w1, ERROR)
    if right_note:
        _note_line(slide, right_note, l2, y, w2, GREEN)
    return h + (Inches(0.62) if (wrong_note or right_note) else Inches(0))


def _note_line(slide, text, left, top, width, color):
    tb = _tb(slide, left, top, width, Inches(0.6))
    p = tb.text_frame.paragraphs[0]
    p.line_spacing = 1.2
    _run(p, text, 16, color)
    return tb


def block_top(height, bottom=None):
    """Centre a block of `height` in the body area above `bottom`."""
    bottom = BODY_BOTTOM if bottom is None else bottom
    return BODY_TOP + max(0, int((bottom - BODY_TOP - height) / 2))


def code_notes(prs, title, code, items, notes_text=None, weights=(1.12, 1),
               code_label="Code", notes_label="What to notice",
               code_tone="normal", code_size=T_CODE, badge_text=None,
               bottom=None, numbers=True):
    """The workhorse layout: a code panel beside a panel of observations.

    Both panels take the same height and the pair is centred in the body area,
    so a short slide does not sit stranded at the top.
    """
    s = content_slide(prs, title, badge_text=badge_text)
    (l1, w1), (l2, w2) = cols(*weights)
    limit = (bottom or BODY_BOTTOM) - BODY_TOP
    while code_size > 11 and code_height(code, size=code_size) > limit:
        code_size -= 1
    h = min(max(code_height(code, size=code_size),
                notes_height(items, w2)), limit)
    top = block_top(h, bottom)
    code_panel(s, code, l1, top, w1, label=code_label, tone=code_tone,
               size=code_size, numbers=numbers, height=h)
    notes_panel(s, items, l2, top, w2, h, label=notes_label)
    _notes(s, notes_text)
    return s


# ================================================================ furniture
def cols(*weights, **kw):
    """Column geometry across the body: cols(1, 1) or cols(1.15, 1)."""
    gutter = kw.get("gutter", GUTTER)
    left, total = kw.get("left", MARGIN), kw.get("width", BODYW)
    n = len(weights)
    usable = total - gutter * (n - 1)
    ssum = float(sum(weights))
    out, x = [], left
    for wgt in weights:
        w = int(usable * wgt / ssum)
        out.append((int(x), w))
        x += w + gutter
    return out


def pipeline(slide, steps, top, left=MARGIN, total_w=BODYW, box_h=Inches(0.86),
             size=16, sub=None, colors=None, gap=Inches(0.08)):
    """Horizontal chevron pipeline with optional monospace sub-labels."""
    n = len(steps)
    bw = int((total_w - gap * (n - 1)) / n)
    for i, st in enumerate(steps):
        x = left + i * (bw + gap)
        kind = MSO_SHAPE.PENTAGON if i < n - 1 else MSO_SHAPE.RECTANGLE
        shp = _rect(slide, x, top, bw, box_h,
                    fill=colors[i] if colors else GREEN, shape=kind)
        tf = shp.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        _run(p, st, size, WHITE, bold=True)
        if sub and sub[i]:
            tb = _tb(slide, x, top + box_h + Inches(0.10), bw, Inches(0.44))
            p2 = tb.text_frame.paragraphs[0]
            p2.alignment = PP_ALIGN.CENTER
            _run(p2, sub[i], 13, GREY, mono=True)
    return top + box_h + (Inches(0.56) if sub else Inches(0))


def _indent(paragraph, level):
    """Set a real left indent on a textbox paragraph.

    Textbox paragraphs ignore the outline level, so indentation has to be set
    as a margin.
    """
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(0.45) * level)))
    pPr.set("indent", "0")
    return paragraph


def callout(slide, text, left=MARGIN, top=None, width=BODYW,
            height=Inches(0.86), color=None, size=T_CALLOUT, icon=None):
    """A single emphatic line. Green by default, red when color=ERROR."""
    col = color or GREEN
    # A full-width callout is the closing line of the slide and always sits at
    # the foot, so a `top` passed for a full-width one is ignored. A callout
    # given its own place on the slide — narrower, or off the left margin —
    # keeps the position it was given.
    full_width = (int(left) == int(MARGIN)
                  and abs(int(width) - int(BODYW)) < Inches(0.1))
    if top is None or full_width:
        top = BODY_BOTTOM - height
    top = min(int(top), int(BODY_BOTTOM - height))
    _rect(slide, left, top, width, height, fill=col)
    tb = _tb(slide, left + Inches(0.28), top, width - Inches(0.5), height,
             anchor=MSO_ANCHOR.MIDDLE)
    p = tb.text_frame.paragraphs[0]
    p.line_spacing = 1.2
    _rich(p, text, size, WHITE, bold=True)
    return top


def badge(slide, text, left, top, color=None, width=Inches(1.9)):
    col = color or GREEN
    _rect(slide, left, top, width, Inches(0.38), fill=col)
    tb = _tb(slide, left, top + Inches(0.055), width, Inches(0.3), wrap=False)
    p = tb.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = _run(p, text.upper(), T_LABEL, WHITE, bold=True, mono=True)
    r.font._rPr.set("spc", "120")
    return tb


def label(slide, text, left, top, width, size=15, color=None, bold=True,
          align=PP_ALIGN.LEFT, mono=False):
    if top + Inches(0.36) > BODY_BOTTOM + Inches(0.02):
        warn("%s: label \u201c%s\u201d sits below the body area"
             % (_CONTEXT[0], text[:40]))
    tb = _tb(slide, left, top, width, Inches(0.36))
    p = tb.text_frame.paragraphs[0]
    p.alignment = align
    if mono:
        r = _run(p, plain(text), size, color or GREY, bold=bold, mono=True)
    else:
        _rich(p, text, size, color or GREY, bold=bold)
        r = p.runs[0]
    if color in (GREEN, GREY) and text == text.upper():
        r.font._rPr.set("spc", "120")
    return tb


def caption(slide, text, left, top, width, size=16, color=None):
    return label(slide, text, left, top, width, size=size,
                 color=color or GREY, bold=False)


def code_box(slide, code, left, top, width, height=None, size=T_CODE,
             label=None, tone="normal", numbers=None, bg=None, border=None,
             fg=None, **kw):
    """A code panel, sized to its contents; no label strip unless asked for.

    `height` is accepted and ignored: the panel is always measured from the
    code, because a height guessed by eye is wrong as soon as the type scale
    moves. A red-ish `border` marks code that is wrong and selects the error
    tone; a dark `bg` marks a console transcript, which is not line-numbered.
    """
    reds = (tuple(ERROR), (0xC0, 0x39, 0x2B))
    if border is not None and tuple(border) in reds:
        tone = "error"
    if numbers is None:
        printed = bg is not None and sum(tuple(bg)) < 260
        numbers = (CODE_NUMBERS[0] and not printed
                   and len(code.rstrip("\n").split("\n")) >= 3)
    return code_panel(slide, code, left, top, width, label=label, tone=tone,
                      size=size, numbers=numbers)


# ================================================================ tables
_BORDER_ORDER = ["a:lnL", "a:lnR", "a:lnT", "a:lnB",
                 "a:lnTlToBr", "a:lnBlToTr", "a:cell3D"]


def _cell_border(cell, edge, color, pt):
    """Set one border of a table cell (python-pptx has no API for this)."""
    tag = "a:ln" + edge
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn(tag)):
        tcPr.remove(old)
    ln = tcPr.makeelement(qn(tag), {"w": str(int(Pt(pt))), "cap": "flat",
                                    "cmpd": "sng", "algn": "ctr"})
    fill = ln.makeelement(qn("a:solidFill"), {})
    hexed = "%02X%02X%02X" % (color[0], color[1], color[2])
    clr = ln.makeelement(qn("a:srgbClr"), {"val": hexed})
    fill.append(clr)
    ln.append(fill)
    idx = _BORDER_ORDER.index(tag)
    for child in tcPr:
        name = child.tag.split("}")[-1]
        qname = "a:" + name
        if qname not in _BORDER_ORDER or _BORDER_ORDER.index(qname) > idx:
            child.addprevious(ln)
            return ln
    tcPr.append(ln)
    return ln


def _no_style(table):
    """Strip PowerPoint's banded table style so our own borders show."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    for el in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(el)


def table_slide(prs, title, headers, rows, notes=None, colw=None, size=T_TABLE,
                head_size=T_TABLE_HEAD, top=BODY_TOP, row_h=Inches(0.58)):
    s = content_slide(prs, title)
    table(s, headers, rows, MARGIN, top, BODYW, colw=colw, size=size,
          head_size=head_size, row_h=row_h)
    _notes(s, notes)
    return s


def table(slide, headers, rows, left, top, width, colw=None, size=T_TABLE,
          head_size=T_TABLE_HEAD, row_h=Inches(0.58)):
    nrow, ncol = len(rows) + 1, len(headers)
    g = slide.shapes.add_table(nrow, ncol, int(left), int(top), int(width),
                               int(row_h * nrow)).table
    _no_style(g)
    if colw:
        total = float(sum(colw))
        for i, w in enumerate(colw):
            g.columns[i].width = Emu(int(width * w / total))

    for j, h in enumerate(headers):
        c = g.cell(0, j)
        c.text = ""
        c.fill.background()
        c.vertical_anchor = MSO_ANCHOR.BOTTOM
        c.margin_left = Inches(0.14)
        c.margin_bottom = Inches(0.10)
        r = _run(c.text_frame.paragraphs[0], h.upper(), head_size, GREY,
                 bold=True)
        r.font._rPr.set("spc", "140")
        _cell_border(c, "B", GREEN, 2.0)

    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            mono, color = False, INK
            if isinstance(val, tuple):
                mono = True
                color = val[2] if len(val) == 3 else (GREEN if j == 0 else INK)
                val = val[0]
            c = g.cell(i, j)
            c.text = ""
            c.fill.background()
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = Inches(0.14)
            c.margin_top = c.margin_bottom = Inches(0.06)
            c.text_frame.word_wrap = True
            _run(c.text_frame.paragraphs[0], val, size - 1 if mono else size,
                 color, bold=(mono and j == 0), mono=mono)
            if i < len(rows):
                _cell_border(c, "B", RULE, 0.75)
    return g


# ================================================================ the logo
_LOGO_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "media")


def logo(slide, left, top, width, reversed_=False):
    """Horizontal lockup. reversed_=True is the white cut, on green."""
    name = "cu_lockup_white.png" if reversed_ else "cu_lockup_dark.png"
    path = os.path.join(_LOGO_DIR, name)
    if not os.path.exists(path):
        return None
    return slide.shapes.add_picture(path, int(left), int(top),
                                    width=int(width))


def seal(slide, left, top, size):
    path = os.path.join(_LOGO_DIR, "cu_seal.png")
    if not os.path.exists(path):
        return None
    return slide.shapes.add_picture(path, int(left), int(top), width=int(size))


# ================================================================ closing
def references_slide(prs, books=None, note=None, notes_text=None,
                     title="References"):
    """The closing slide: what to read, from the course's book list."""
    books = current_course().books if books is None else list(books)
    if not books:
        raise ValueError(
            "references_slide: no books — set them with "
            "course(books=[...]) or pass "
            "books=[(title, authors, edition), ...]")
    s = content_slide(prs, title)
    y = BODY_TOP + Inches(0.25)
    # each book takes 2" and only two fit; a third runs off the foot
    last = y + Inches(2.0) * (len(books) - 1) + Inches(1.65)
    if last > BODY_BOTTOM + Inches(0.02):
        warn("%s: %d books run %.2f\" past the foot of the slide — "
             "split them across two slides"
             % (title, len(books), (last - BODY_BOTTOM) / 914400.0))
    for entry in books:
        btitle, authors, edition = (list(entry) + ["", ""])[:3]
        bx, by, bw, bh = panel(s, MARGIN, y, BODYW, Inches(1.65),
                               tone="normal")
        tb = _tb(s, bx + Inches(0.2), by + Inches(0.1), bw - Inches(0.4), bh)
        tf = tb.text_frame
        for i, (txt, sz, col, bold) in enumerate([(btitle, 26, INK, True),
                                                  (authors, 20, GREEN, False),
                                                  (edition, 17, GREY, False)]):
            if not txt:
                continue
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(5)
            _run(p, txt, sz, col, bold=bold)
        y += Inches(2.0)
    note = current_course().books_note if note is None else note
    if note:
        caption(s, note, MARGIN, y + Inches(0.05), BODYW, size=18)
    _notes(s, notes_text or current_course().references_note)
    return s
