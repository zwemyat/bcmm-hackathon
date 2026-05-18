"""ITAMS hackathon presentation — Branded to match the live system UI.

15 slides matching the requested structure:
Cover · Team · Problem · Challenges · Solution · Overview · Features ·
Dashboard · Workflow · Architecture · Security & Smart · Demo · Benefits ·
Future · Thank You

Light theme with the same #2563EB primary + #4F46E5 indigo gradient pair
the live ITAMS UI uses, so the deck visually matches the product.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree


# ---------- Palette (matches the live app) ----------
INK         = RGBColor(0x0F, 0x17, 0x2A)
INK_SOFT    = RGBColor(0x1F, 0x2D, 0x3D)
GRAY_700    = RGBColor(0x47, 0x55, 0x69)
GRAY_500    = RGBColor(0x64, 0x74, 0x8B)
GRAY_400    = RGBColor(0x94, 0xA3, 0xB8)
GRAY_300    = RGBColor(0xCB, 0xD5, 0xE1)
GRAY_200    = RGBColor(0xE2, 0xE8, 0xF0)
GRAY_100    = RGBColor(0xF1, 0xF5, 0xF9)
GRAY_50     = RGBColor(0xF8, 0xFA, 0xFC)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

PRIMARY     = RGBColor(0x25, 0x63, 0xEB)
PRIMARY_DK  = RGBColor(0x1D, 0x4E, 0xD8)
INDIGO      = RGBColor(0x4F, 0x46, 0xE5)
PRIMARY_SOFT = RGBColor(0xDB, 0xEA, 0xFE)
INDIGO_SOFT  = RGBColor(0xE0, 0xE7, 0xFF)

SUCCESS     = RGBColor(0x10, 0xB9, 0x81)
SUCCESS_SOFT= RGBColor(0xDC, 0xFC, 0xE7)
AMBER       = RGBColor(0xF5, 0x9E, 0x0B)
AMBER_SOFT  = RGBColor(0xFE, 0xF3, 0xC7)
ROSE        = RGBColor(0xEF, 0x44, 0x44)
ROSE_SOFT   = RGBColor(0xFE, 0xE2, 0xE2)
PURPLE      = RGBColor(0x8B, 0x5C, 0xF6)
PURPLE_SOFT = RGBColor(0xED, 0xE9, 0xFE)
CYAN        = RGBColor(0x06, 0xB6, 0xD4)

FONT = "Calibri"
TOTAL = 15

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# =====================================================================
# Primitives
# =====================================================================
def add_slide():
    return prs.slides.add_slide(BLANK)


def add_rect(slide, x, y, w, h, rgb, *, no_line=True, corner=None):
    t = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(t, x, y, w, h)
    if corner:
        shp.adjustments[0] = corner
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb
    if no_line:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_outline_rect(slide, x, y, w, h, *, line_rgb=GRAY_200, line_pt=0.75,
                     fill_rgb=None, corner=None):
    t = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(t, x, y, w, h)
    if corner:
        shp.adjustments[0] = corner
    if fill_rgb is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_rgb
    shp.line.color.rgb = line_rgb
    shp.line.width = Pt(line_pt)
    shp.shadow.inherit = False
    return shp


def _strip_fill(spPr):
    for tag in ('a:solidFill', 'a:gradFill', 'a:noFill', 'a:pattFill', 'a:blipFill'):
        ex = spPr.find(qn(tag))
        if ex is not None:
            spPr.remove(ex)


def _insert_gradient(spPr, gradFill):
    ln = spPr.find(qn('a:ln'))
    if ln is not None:
        spPr.insert(list(spPr).index(ln), gradFill)
    else:
        spPr.append(gradFill)


def add_gradient_rect(slide, x, y, w, h, color_start, color_end,
                      *, angle_deg=0, corner=None,
                      alpha_start=None, alpha_end=None):
    t = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(t, x, y, w, h)
    if corner:
        shp.adjustments[0] = corner
    spPr = shp.fill._xPr
    _strip_fill(spPr)
    g = etree.Element(qn('a:gradFill'))
    g.set('flip', 'none'); g.set('rotWithShape', '1')
    gs = etree.SubElement(g, qn('a:gsLst'))
    for pos, c, a in [(0, color_start, alpha_start),
                      (100000, color_end, alpha_end)]:
        st = etree.SubElement(gs, qn('a:gs'))
        st.set('pos', str(pos))
        sr = etree.SubElement(st, qn('a:srgbClr'))
        sr.set('val', f'{c[0]:02X}{c[1]:02X}{c[2]:02X}')
        if a is not None:
            ae = etree.SubElement(sr, qn('a:alpha'))
            ae.set('val', str(a))
    lin = etree.SubElement(g, qn('a:lin'))
    lin.set('ang', str(int(angle_deg * 60000)))
    lin.set('scaled', '0')
    _insert_gradient(spPr, g)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_line(slide, x1, y1, x2, y2, rgb=GRAY_300, width=0.75):
    line = slide.shapes.add_connector(1, x1, y1, x2, y2)
    line.line.color.rgb = rgb
    line.line.width = Pt(width)
    return line


def add_arrow(slide, x1, y1, x2, y2, rgb=GRAY_400, width=1.0):
    line = slide.shapes.add_connector(1, x1, y1, x2, y2)
    line.line.color.rgb = rgb
    line.line.width = Pt(width)
    lnEl = line.line._get_or_add_ln()
    tail = etree.SubElement(lnEl, qn("a:tailEnd"))
    tail.set("type", "triangle"); tail.set("w", "sm"); tail.set("h", "sm")
    return line


def text(slide, x, y, w, h, content, *,
         size=11, bold=False, color=INK,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         font=FONT, letter_spacing=None, italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = content.split("\n") if isinstance(content, str) else content
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = font
        if letter_spacing is not None:
            rPr = run._r.get_or_add_rPr()
            rPr.set("spc", str(letter_spacing))
    return tb


# =====================================================================
# Reusable design components
# =====================================================================
def eyebrow(slide, x, y, label):
    add_rect(slide, x, y + Inches(0.13), Inches(0.06), Inches(0.06),
             PRIMARY, corner=0.5)
    text(slide, x + Inches(0.15), y, Inches(8), Inches(0.3),
         label.upper(), size=9, bold=True, color=PRIMARY, letter_spacing=200)


def page_chrome(slide, page_no, eyebrow_text, title_text, *,
                title_size=30):
    """Standard page setup matching the app's page-header pattern."""
    add_rect(slide, 0, 0, SW, SH, WHITE)
    # Soft gradient header sliver (matches app's primary accent)
    add_gradient_rect(slide, 0, 0, SW, Inches(0.08),
                      PRIMARY, INDIGO, angle_deg=0)
    eyebrow(slide, Inches(0.75), Inches(0.55), eyebrow_text)
    text(slide, Inches(0.75), Inches(0.85), Inches(12), Inches(0.7),
         title_text, size=title_size, bold=True, color=INK,
         letter_spacing=-15)
    # bottom thin divider + wordmark
    add_line(slide, Inches(0.75), SH - Inches(0.55),
             SW - Inches(0.75), SH - Inches(0.55),
             rgb=GRAY_200, width=0.5)
    text(slide, Inches(0.75), SH - Inches(0.45),
         Inches(2.2), Inches(0.3),
         "ITAMS", size=9, bold=True, color=INK, letter_spacing=200)
    text(slide, Inches(1.25), SH - Inches(0.45),
         Inches(6), Inches(0.3),
         "IT Asset Management System",
         size=9, color=GRAY_500)
    text(slide, SW - Inches(1.4), SH - Inches(0.45),
         Inches(0.85), Inches(0.3),
         f"{page_no:02d} / {TOTAL}", size=9, bold=True, color=GRAY_400,
         align=PP_ALIGN.RIGHT, letter_spacing=150)


def card(slide, x, y, w, h, *, accent_top=False):
    """White card with subtle border + optional gradient top strip."""
    add_outline_rect(slide, x, y, w, h,
                     line_rgb=GRAY_200, line_pt=0.5,
                     fill_rgb=WHITE, corner=0.05)
    if accent_top:
        add_gradient_rect(slide, x, y, w, Inches(0.06),
                          PRIMARY, INDIGO, angle_deg=0)


def icon_chip(slide, x, y, glyph, *, size_in=0.5,
              bg=PRIMARY_SOFT, fg=PRIMARY):
    s = Inches(size_in)
    add_rect(slide, x, y, s, s, bg, corner=0.22)
    text(slide, x, y, s, s, glyph, size=int(size_in * 30), bold=True,
         color=fg, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def kpi_tile(slide, x, y, w, h, value, label, accent=PRIMARY,
             accent_soft=PRIMARY_SOFT, foot=None):
    """Matches app's .kpi-card pattern: white card + left accent stripe."""
    add_outline_rect(slide, x, y, w, h,
                     line_rgb=GRAY_200, line_pt=0.5,
                     fill_rgb=WHITE, corner=0.05)
    add_rect(slide, x, y, Inches(0.06), h, accent)
    text(slide, x + Inches(0.25), y + Inches(0.18),
         w - Inches(0.4), Inches(0.3),
         label.upper(), size=8.5, bold=True, color=GRAY_500,
         letter_spacing=200)
    text(slide, x + Inches(0.25), y + Inches(0.42),
         w - Inches(0.4), Inches(0.55),
         value, size=22, bold=True, color=INK, letter_spacing=-15)
    if foot:
        text(slide, x + Inches(0.25), y + Inches(0.95),
             w - Inches(0.4), Inches(0.3),
             foot, size=8.5, color=GRAY_500)


# =====================================================================
# 01 — Cover
# =====================================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, GRAY_50)
# Decorative gradient hero (top half)
add_gradient_rect(s, 0, 0, SW, Inches(4.2),
                  PRIMARY_DK, INDIGO, angle_deg=135)
# Soft tinted overlay shapes
add_gradient_rect(s, SW - Inches(7), Inches(-1.5),
                  Inches(9), Inches(9),
                  PRIMARY, INDIGO,
                  angle_deg=135,
                  alpha_start=22000, alpha_end=0)

# top eyebrow
text(s, Inches(0.75), Inches(0.6), Inches(8), Inches(0.3),
     "HACKATHON 2026  ·  IT INFRASTRUCTURE TRACK",
     size=9, bold=True, color=WHITE, letter_spacing=300)

# brand mark like the app's sidebar logo
add_rect(s, Inches(0.75), Inches(1.1),
         Inches(0.7), Inches(0.7), WHITE, corner=0.2)
text(s, Inches(0.75), Inches(1.1), Inches(0.7), Inches(0.7),
     "I", size=26, bold=True, color=PRIMARY,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(1.6), Inches(1.18),
     Inches(6), Inches(0.4),
     "ITAMS",
     size=16, bold=True, color=WHITE, letter_spacing=150)
text(s, Inches(1.6), Inches(1.55),
     Inches(7), Inches(0.3),
     "IT Asset Management System",
     size=10, color=WHITE, letter_spacing=100)

# big title
text(s, Inches(0.75), Inches(2.6), Inches(12), Inches(1.4),
     "Every IT asset.\nOne source of truth.",
     size=54, bold=True, color=WHITE, letter_spacing=-25)

# accent rule between hero and footer block
add_rect(s, 0, Inches(4.2), SW, Inches(0.04), PRIMARY)

# lower block: meta info
text(s, Inches(0.75), Inches(4.65), Inches(12), Inches(0.4),
     "Enterprise-grade asset, license & subscription management.",
     size=18, color=INK, letter_spacing=-10)
text(s, Inches(0.75), Inches(5.1), Inches(12), Inches(0.4),
     "Built on Laravel 11 · MySQL · Bootstrap 5  —  on-premise ready.",
     size=12, color=GRAY_500)

# Presented-by row
text(s, Inches(0.75), SH - Inches(1.3),
     Inches(4), Inches(0.25),
     "PRESENTED BY", size=8, bold=True, color=GRAY_500, letter_spacing=300)
text(s, Inches(0.75), SH - Inches(0.95),
     Inches(8), Inches(0.45),
     "Team ITAMS", size=18, bold=True, color=INK)

text(s, SW - Inches(5), SH - Inches(1.3),
     Inches(4), Inches(0.25),
     "DATE", size=8, bold=True, color=GRAY_500, letter_spacing=300,
     align=PP_ALIGN.RIGHT)
text(s, SW - Inches(5), SH - Inches(0.95),
     Inches(4), Inches(0.4),
     "May 2026", size=14, color=INK, align=PP_ALIGN.RIGHT)


# =====================================================================
# 02 — Team Introduction
# =====================================================================
s = add_slide()
page_chrome(s, 2, "Team Introduction", "Meet the team.",
            title_size=34)

text(s, Inches(0.75), Inches(2.05), Inches(11), Inches(0.4),
     "Four cross-functional engineers building enterprise IT operations tooling.",
     size=13, color=GRAY_500)

members = [
    ("PL", "[Member 1]", "Project Lead / Backend",
     "Architecture · APIs · DB design", PRIMARY, PRIMARY_SOFT),
    ("FE", "[Member 2]", "Frontend / UI Engineer",
     "Bootstrap · UX · accessibility",  INDIGO, INDIGO_SOFT),
    ("FS", "[Member 3]", "Full-Stack Developer",
     "Laravel · integrations · testing", PURPLE, PURPLE_SOFT),
    ("DB", "[Member 4]", "Database / DevOps",
     "MySQL · scheduler · deployment",  SUCCESS, SUCCESS_SOFT),
]
card_w = Inches(2.85)
card_h = Inches(3.6)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
sx = (SW - total_w) / 2
y = Inches(2.8)
for i, (mono, name, role, expertise, color, soft) in enumerate(members):
    x = sx + i * (card_w + gap)
    card(s, x, y, card_w, card_h, accent_top=True)
    # avatar with gradient
    av = Inches(1.2)
    avx = x + (card_w - av) / 2
    avy = y + Inches(0.55)
    add_gradient_rect(s, avx, avy, av, av,
                      (color[0], color[1], color[2]),
                      (INDIGO[0], INDIGO[1], INDIGO[2]),
                      angle_deg=45, corner=0.5)
    text(s, avx, avy, av, av, mono, size=30, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x, y + Inches(2.05), card_w, Inches(0.4),
         name, size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)
    text(s, x, y + Inches(2.5), card_w, Inches(0.4),
         role, size=10.5, color=GRAY_700, align=PP_ALIGN.CENTER)
    add_line(s, x + (card_w - Inches(0.5)) / 2,
             y + Inches(3.05),
             x + (card_w + Inches(0.5)) / 2,
             y + Inches(3.05),
             rgb=color, width=1.5)
    text(s, x, y + Inches(3.15), card_w, Inches(0.4),
         expertise, size=9.5, color=GRAY_500, align=PP_ALIGN.CENTER,
         italic=True)


# =====================================================================
# 03 — Problem Statement
# =====================================================================
s = add_slide()
page_chrome(s, 3, "Problem Statement", "Why IT teams are struggling.",
            title_size=32)

# Big lead statement
text(s, Inches(0.75), Inches(2.2), Inches(12), Inches(1.5),
     "IT estates are growing —\nthe tools to manage them aren't.",
     size=32, bold=True, color=INK, letter_spacing=-15)

# Three supporting stats
stats = [
    ("70%",  "of IT teams still track assets in Excel",     "Industry survey, 2024"),
    ("$1.3K","wasted per employee on unused licenses",        "Gartner, 2023"),
    ("28%",  "of subscriptions auto-renew unnoticed each year","Flexera State of ITAM"),
]
stat_w = Inches(3.95)
stat_h = Inches(2.0)
gap = Inches(0.25)
total_w = stat_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(4.6)
for i, (val, label, src) in enumerate(stats):
    x = sx + i * (stat_w + gap)
    card(s, x, y, stat_w, stat_h, accent_top=True)
    text(s, x + Inches(0.3), y + Inches(0.4),
         stat_w - Inches(0.6), Inches(0.9),
         val, size=44, bold=True, color=PRIMARY, letter_spacing=-25)
    text(s, x + Inches(0.3), y + Inches(1.15),
         stat_w - Inches(0.6), Inches(0.45),
         label, size=12, bold=True, color=INK)
    text(s, x + Inches(0.3), y + Inches(1.6),
         stat_w - Inches(0.6), Inches(0.3),
         src, size=9, color=GRAY_500, italic=True)


# =====================================================================
# 04 — Current Challenges
# =====================================================================
s = add_slide()
page_chrome(s, 4, "Current Challenges", "What goes wrong every day.",
            title_size=32)

challenges = [
    ("◆", "Manual tracking",
     "Spreadsheets sprawl across drives. No single owner, no real-time view, no version control.",
     PRIMARY, PRIMARY_SOFT),
    ("◐", "Missed renewals",
     "Subscriptions auto-renew or expire without warning — outages and surprise invoices follow.",
     AMBER, AMBER_SOFT),
    ("◇", "No audit trail",
     "When something breaks, no one knows who changed it, when, or why. Compliance reviews fail.",
     ROSE, ROSE_SOFT),
    ("●", "Shadow IT spend",
     "Duplicate purchases and forgotten subscriptions inflate the IT bill 20–30% silently.",
     INDIGO, INDIGO_SOFT),
]
card_w = Inches(5.85)
card_h = Inches(2.0)
gap_x = Inches(0.25)
gap_y = Inches(0.25)
sx = (SW - card_w * 2 - gap_x) / 2
sy = Inches(2.4)
for i, (glyph, head, body, color, soft) in enumerate(challenges):
    row = i // 2
    col = i % 2
    x = sx + col * (card_w + gap_x)
    y = sy + row * (card_h + gap_y)
    card(s, x, y, card_w, card_h)
    icon_chip(s, x + Inches(0.35), y + Inches(0.35),
              glyph, size_in=0.55, bg=soft, fg=color)
    text(s, x + Inches(1.15), y + Inches(0.35), card_w - Inches(1.4),
         Inches(0.4), head, size=15, bold=True, color=INK)
    text(s, x + Inches(1.15), y + Inches(0.8),
         card_w - Inches(1.4), card_h - Inches(0.95),
         body, size=11, color=GRAY_700)


# =====================================================================
# 05 — Proposed Solution
# =====================================================================
s = add_slide()
page_chrome(s, 5, "Proposed Solution", "ITAMS · one platform, full lifecycle.",
            title_size=32)

text(s, Inches(0.75), Inches(2.2), Inches(12), Inches(0.8),
     "ITAMS unifies hardware, devices, licenses and subscriptions —",
     size=18, color=INK)
text(s, Inches(0.75), Inches(2.7), Inches(12), Inches(0.8),
     "with automated renewal reminders, role-based access, and a full audit log.",
     size=18, color=INK)

# 3 pillars
pillars = [
    ("◆", "Unify",
     "One source of truth for every IT asset across the organisation.",
     PRIMARY, PRIMARY_SOFT),
    ("◉", "Automate",
     "Email and in-app reminders fire before things expire.",
     INDIGO, INDIGO_SOFT),
    ("◇", "Govern",
     "Role-based access with full activity history per user.",
     SUCCESS, SUCCESS_SOFT),
]
card_w = Inches(3.95)
card_h = Inches(2.6)
gap = Inches(0.25)
total_w = card_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(4.05)
for i, (g, head, body, color, soft) in enumerate(pillars):
    x = sx + i * (card_w + gap)
    card(s, x, y, card_w, card_h, accent_top=True)
    icon_chip(s, x + Inches(0.35), y + Inches(0.35),
              g, size_in=0.55, bg=soft, fg=color)
    text(s, x + Inches(0.35), y + Inches(1.15),
         card_w - Inches(0.6), Inches(0.5),
         head, size=20, bold=True, color=INK, letter_spacing=-15)
    text(s, x + Inches(0.35), y + Inches(1.7),
         card_w - Inches(0.6), card_h - Inches(1.85),
         body, size=11, color=GRAY_700)


# =====================================================================
# 06 — System Overview
# =====================================================================
s = add_slide()
page_chrome(s, 6, "System Overview", "What lives inside ITAMS.",
            title_size=32)

# central spine: USER → ITAMS (core) → DATA
# Module ring around the core
core_x = SW / 2 - Inches(1.5)
core_y = Inches(3.6)
core_w = Inches(3.0)
core_h = Inches(1.3)
add_gradient_rect(s, core_x, core_y, core_w, core_h,
                  PRIMARY, INDIGO, angle_deg=135, corner=0.08)
text(s, core_x, core_y + Inches(0.18),
     core_w, Inches(0.3),
     "CORE PLATFORM", size=9, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, letter_spacing=200)
text(s, core_x, core_y + Inches(0.48),
     core_w, Inches(0.5),
     "ITAMS", size=24, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, letter_spacing=-15)
text(s, core_x, core_y + Inches(0.92),
     core_w, Inches(0.3),
     "Laravel 11 · MySQL · Blade",
     size=10, color=WHITE, align=PP_ALIGN.CENTER)

# 4 modules arranged left & right
modules = [
    ("PC Master", "Workstations & laptops", "◆", PRIMARY, PRIMARY_SOFT,
     Inches(0.85), Inches(2.4)),
    ("Device Master", "Network hardware", "◐", INDIGO, INDIGO_SOFT,
     Inches(0.85), Inches(4.85)),
    ("Subscriptions", "Recurring services", "◇", CYAN, PRIMARY_SOFT,
     SW - Inches(4.0), Inches(2.4)),
    ("Licenses & Contracts", "Software & contracts", "●", SUCCESS, SUCCESS_SOFT,
     SW - Inches(4.0), Inches(4.85)),
]
for (name, desc, g, color, soft, mx, my) in modules:
    mw = Inches(3.15)
    mh = Inches(1.3)
    card(s, mx, my, mw, mh)
    add_rect(s, mx, my, Inches(0.06), mh, color)
    icon_chip(s, mx + Inches(0.25), my + Inches(0.35),
              g, size_in=0.5, bg=soft, fg=color)
    text(s, mx + Inches(0.95), my + Inches(0.32),
         mw - Inches(1.1), Inches(0.4),
         name, size=14, bold=True, color=INK)
    text(s, mx + Inches(0.95), my + Inches(0.7),
         mw - Inches(1.1), Inches(0.4),
         desc, size=10.5, color=GRAY_500)
    # connector to core
    if mx < SW / 2:
        add_arrow(s, mx + mw, my + mh / 2,
                  core_x, core_y + Inches(0.35) + (my - Inches(2.4)) / 4,
                  rgb=GRAY_300, width=1.0)
    else:
        add_arrow(s, mx, my + mh / 2,
                  core_x + core_w,
                  core_y + Inches(0.35) + (my - Inches(2.4)) / 4,
                  rgb=GRAY_300, width=1.0)

# top: users
top_y = Inches(2.05)
text(s, Inches(0.75), top_y, Inches(12), Inches(0.3),
     "ADMINS  ·  IT STAFF  ·  ASSET OWNERS",
     size=9, bold=True, color=GRAY_500, letter_spacing=300,
     align=PP_ALIGN.CENTER)

# bottom: services
btm_y = Inches(6.3)
services = ["Scheduler", "SMTP", "Activity Log", "Notifications"]
sw_each = Inches(2.4)
gap = Inches(0.2)
total = sw_each * 4 + gap * 3
sxs = (SW - total) / 2
for i, svc in enumerate(services):
    x = sxs + i * (sw_each + gap)
    add_outline_rect(s, x, btm_y, sw_each, Inches(0.5),
                     line_rgb=GRAY_200, line_pt=0.5,
                     fill_rgb=GRAY_50, corner=0.25)
    text(s, x, btm_y, sw_each, Inches(0.5),
         svc, size=10.5, color=INK, bold=True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# =====================================================================
# 07 — Key Features
# =====================================================================
s = add_slide()
page_chrome(s, 7, "Key Features", "What ITAMS does for IT teams.",
            title_size=32)

features = [
    ("◆", "Unified inventory",
     "PCs, devices, licenses and subscriptions in one searchable estate.",
     PRIMARY, PRIMARY_SOFT),
    ("◉", "Smart reminders",
     "Daily scheduler sends staggered digests at the day-marks you choose.",
     INDIGO, INDIGO_SOFT),
    ("◇", "Role-based access",
     "Per-module view & edit permissions enforced server-side by middleware.",
     SUCCESS, SUCCESS_SOFT),
    ("◐", "Full audit trail",
     "Every create, update and delete logged with user, time and field diff.",
     PURPLE, PURPLE_SOFT),
    ("▤", "Bulk import / export",
     "Excel import & export per module, with downloadable templates.",
     AMBER, AMBER_SOFT),
    ("●", "Live status badge",
     "Topbar bell counts overdue and due-soon items per user in real time.",
     CYAN, PRIMARY_SOFT),
]
card_w = Inches(3.95)
card_h = Inches(1.95)
gap_x = Inches(0.2)
gap_y = Inches(0.25)
sx = Inches(0.7)
sy = Inches(2.4)
for i, (g, head, body, color, soft) in enumerate(features):
    row = i // 3
    col = i % 3
    x = sx + col * (card_w + gap_x)
    y = sy + row * (card_h + gap_y)
    card(s, x, y, card_w, card_h)
    icon_chip(s, x + Inches(0.3), y + Inches(0.3),
              g, size_in=0.5, bg=soft, fg=color)
    text(s, x + Inches(0.3), y + Inches(1.0),
         card_w - Inches(0.5), Inches(0.4),
         head, size=14, bold=True, color=INK)
    text(s, x + Inches(0.3), y + Inches(1.4),
         card_w - Inches(0.5), card_h - Inches(1.55),
         body, size=10.5, color=GRAY_700)


# =====================================================================
# 08 — Dashboard Overview (annotated mockup)
# =====================================================================
s = add_slide()
page_chrome(s, 8, "Dashboard Overview", "At-a-glance visibility on every screen.",
            title_size=32)

# Big dashboard mockup taking most of the slide
mx = Inches(0.75)
my = Inches(2.2)
mw = Inches(8.5)
mh = Inches(4.55)

card(s, mx, my, mw, mh)
# top browser bar
add_line(s, mx, my + Inches(0.4),
         mx + mw, my + Inches(0.4), rgb=GRAY_200, width=0.5)
for i, c in enumerate([ROSE, AMBER, SUCCESS]):
    cx = mx + Inches(0.15) + Inches(0.22) * i
    add_rect(s, cx, my + Inches(0.13),
             Inches(0.14), Inches(0.14), c, corner=0.5)
add_outline_rect(s, mx + Inches(0.9), my + Inches(0.1),
                 mw - Inches(1.1), Inches(0.22),
                 line_rgb=GRAY_200, line_pt=0.5,
                 fill_rgb=GRAY_50, corner=0.3)
text(s, mx + Inches(1.0), my + Inches(0.1),
     mw - Inches(1.3), Inches(0.22),
     "itams.local / dashboard",
     size=8.5, color=GRAY_500, anchor=MSO_ANCHOR.MIDDLE)

# Sidebar
sb_w = Inches(1.4)
sb_y = my + Inches(0.4)
sb_h = mh - Inches(0.4)
add_rect(s, mx, sb_y, sb_w, sb_h, GRAY_50)
add_line(s, mx + sb_w, sb_y, mx + sb_w, my + mh,
         rgb=GRAY_200, width=0.5)
# brand
add_rect(s, mx + Inches(0.18), sb_y + Inches(0.15),
         Inches(0.32), Inches(0.32), PRIMARY, corner=0.25)
text(s, mx + Inches(0.18), sb_y + Inches(0.15),
     Inches(0.32), Inches(0.32), "I",
     size=11, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, mx + Inches(0.6), sb_y + Inches(0.19),
     Inches(0.8), Inches(0.25),
     "ITAMS", size=10, bold=True, color=INK, letter_spacing=150)

# nav
nav = [("Dashboard", True), ("PC Master", False),
       ("Device Master", False), ("Subscriptions", False),
       ("Licenses", False), ("Notifications", False),
       ("Activity Log", False)]
ny = sb_y + Inches(0.75)
for name, active in nav:
    if active:
        add_rect(s, mx + Inches(0.1), ny - Inches(0.04),
                 sb_w - Inches(0.2), Inches(0.28),
                 PRIMARY_SOFT, corner=0.25)
        text(s, mx + Inches(0.22), ny,
             sb_w - Inches(0.3), Inches(0.22),
             name, size=9, bold=True, color=PRIMARY,
             anchor=MSO_ANCHOR.MIDDLE)
    else:
        text(s, mx + Inches(0.22), ny,
             sb_w - Inches(0.3), Inches(0.22),
             name, size=9, color=GRAY_700,
             anchor=MSO_ANCHOR.MIDDLE)
    ny += Inches(0.32)

# Main content area
mcx = mx + sb_w + Inches(0.2)
mcy = my + Inches(0.55)
mcw = mw - sb_w - Inches(0.4)

# page header
text(s, mcx, mcy, mcw, Inches(0.3),
     "Dashboard", size=14, bold=True, color=INK, letter_spacing=-10)
text(s, mcx, mcy + Inches(0.3), mcw, Inches(0.25),
     "Welcome back — here's what needs your attention.",
     size=8.5, color=GRAY_500)

# KPI tiles (matching the app's pattern)
kpi_y = mcy + Inches(0.65)
kpi_w = (mcw - Inches(0.3)) / 4 - Inches(0.05)
kpi_h = Inches(0.9)
kpis = [
    ("128", "PCs",       PRIMARY),
    ("64",  "Devices",   INDIGO),
    ("32",  "Active subs", SUCCESS),
    ("5",   "Expiring soon", AMBER),
]
for i, (v, l, c) in enumerate(kpis):
    x = mcx + i * (kpi_w + Inches(0.08))
    add_outline_rect(s, x, kpi_y, kpi_w, kpi_h,
                     line_rgb=GRAY_200, line_pt=0.5,
                     fill_rgb=WHITE, corner=0.08)
    add_rect(s, x, kpi_y, Inches(0.05), kpi_h, c)
    text(s, x + Inches(0.15), kpi_y + Inches(0.12),
         kpi_w - Inches(0.25), Inches(0.25),
         l.upper(), size=7, color=GRAY_500, letter_spacing=200, bold=True)
    text(s, x + Inches(0.15), kpi_y + Inches(0.38),
         kpi_w - Inches(0.25), Inches(0.45),
         v, size=18, bold=True, color=INK)

# bottom two panels: Expiring Soon + Recent Activity
panel_y = kpi_y + kpi_h + Inches(0.2)
panel_h = my + mh - panel_y - Inches(0.2)

# Expiring Soon (left)
exp_w = mcw * 0.55
add_outline_rect(s, mcx, panel_y, exp_w, panel_h,
                 line_rgb=GRAY_200, line_pt=0.5,
                 fill_rgb=WHITE, corner=0.05)
text(s, mcx + Inches(0.18), panel_y + Inches(0.12),
     exp_w - Inches(0.3), Inches(0.25),
     "Expiring soon", size=9.5, bold=True, color=INK)
text(s, mcx + Inches(0.18), panel_y + Inches(0.35),
     exp_w - Inches(0.3), Inches(0.2),
     "Next 30 days · 5 items",
     size=7.5, color=GRAY_500)
# rows
rows = [
    ("Office 365 E3",    "Microsoft",  "3 days",  ROSE,    ROSE_SOFT),
    ("Cloudflare Pro",   "Cloudflare", "8 days",  AMBER,   AMBER_SOFT),
    ("Adobe Creative",   "Adobe",      "14 days", AMBER,   AMBER_SOFT),
    ("AWS Reserved",     "Amazon",     "21 days", PRIMARY, PRIMARY_SOFT),
    ("GitHub Team",      "GitHub",     "27 days", PRIMARY, PRIMARY_SOFT),
]
ry = panel_y + Inches(0.7)
for nm, vendor, due, color, soft in rows:
    text(s, mcx + Inches(0.2), ry,
         exp_w * 0.5, Inches(0.22),
         nm, size=8.5, bold=True, color=INK)
    text(s, mcx + Inches(0.2), ry + Inches(0.2),
         exp_w * 0.5, Inches(0.2),
         vendor, size=7.5, color=GRAY_500)
    # due-pill on right
    pill_w = Inches(0.7)
    pill_x = mcx + exp_w - pill_w - Inches(0.18)
    add_rect(s, pill_x, ry + Inches(0.05),
             pill_w, Inches(0.22), soft, corner=0.45)
    text(s, pill_x, ry + Inches(0.05), pill_w, Inches(0.22),
         due, size=7.5, bold=True, color=color,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    ry += Inches(0.4)

# Recent Activity (right)
act_x = mcx + exp_w + Inches(0.15)
act_w = mcw - exp_w - Inches(0.15)
add_outline_rect(s, act_x, panel_y, act_w, panel_h,
                 line_rgb=GRAY_200, line_pt=0.5,
                 fill_rgb=WHITE, corner=0.05)
text(s, act_x + Inches(0.18), panel_y + Inches(0.12),
     act_w - Inches(0.3), Inches(0.25),
     "Recent activity", size=9.5, bold=True, color=INK)
text(s, act_x + Inches(0.18), panel_y + Inches(0.35),
     act_w - Inches(0.3), Inches(0.2),
     "Today", size=7.5, color=GRAY_500)
acts = [
    ("[Member 1] renewed Office 365",     "09:12", SUCCESS),
    ("[Member 2] added 4 new PCs",        "08:54", PRIMARY),
    ("[Member 4] updated SMTP settings",  "08:21", INDIGO),
    ("Reminder digest sent (12 items)",   "08:00", AMBER),
]
ay = panel_y + Inches(0.7)
for act, t, c in acts:
    add_rect(s, act_x + Inches(0.2), ay + Inches(0.07),
             Inches(0.1), Inches(0.1), c, corner=0.5)
    text(s, act_x + Inches(0.38), ay,
         act_w - Inches(0.55), Inches(0.22),
         act, size=8.5, color=INK)
    text(s, act_x + Inches(0.38), ay + Inches(0.2),
         act_w - Inches(0.55), Inches(0.18),
         t, size=7, color=GRAY_500)
    ay += Inches(0.42)

# Annotation callouts on the right side
ann_x = mx + mw + Inches(0.25)
ann_w = SW - ann_x - Inches(0.75)
ann_items = [
    ("01", "Live KPI tiles",
     "Real-time counts of PCs, devices, active subs, and items expiring soon."),
    ("02", "Expiring soon panel",
     "Color-coded urgency: red within 7 days, amber within 14, blue beyond."),
    ("03", "Recent activity feed",
     "Audit-log tail showing what changed today and by whom."),
]
ay = Inches(2.4)
for num, head, body in ann_items:
    text(s, ann_x, ay, Inches(0.5), Inches(0.3),
         num, size=10, bold=True, color=PRIMARY, letter_spacing=200)
    text(s, ann_x + Inches(0.5), ay - Inches(0.02),
         ann_w - Inches(0.5), Inches(0.35),
         head, size=12, bold=True, color=INK)
    text(s, ann_x + Inches(0.5), ay + Inches(0.32),
         ann_w - Inches(0.5), Inches(0.9),
         body, size=10, color=GRAY_700)
    ay += Inches(1.5)


# =====================================================================
# 09 — Workflow / Process Flow
# =====================================================================
s = add_slide()
page_chrome(s, 9, "Workflow", "How a subscription becomes a renewal reminder.",
            title_size=30)

steps = [
    ("01", "Onboard",
     "Admin adds the subscription or imports an .xlsx.",
     PRIMARY, PRIMARY_SOFT, "◆"),
    ("02", "Track",
     "ITAMS recalculates reminder_date on every save.",
     INDIGO, INDIGO_SOFT, "◉"),
    ("03", "Detect",
     "Daily scheduler scans for items at each day-mark.",
     PURPLE, PURPLE_SOFT, "◇"),
    ("04", "Notify",
     "Digest email + in-app badge alert recipients.",
     AMBER, AMBER_SOFT, "◐"),
    ("05", "Renew",
     "Owner clicks Renew; status & history auto-update.",
     SUCCESS, SUCCESS_SOFT, "●"),
]
n = len(steps)
card_w = Inches(2.32)
card_h = Inches(3.0)
gap = Inches(0.12)
total_w = card_w * n + gap * (n - 1)
sx = (SW - total_w) / 2
y = Inches(3.0)
# gradient line across
line_y = y + Inches(0.55)
add_gradient_rect(s, sx + Inches(0.3), line_y - Inches(0.02),
                  total_w - Inches(0.6), Inches(0.04),
                  PRIMARY, INDIGO, angle_deg=0)

for i, (num, head, body, color, soft, glyph) in enumerate(steps):
    x = sx + i * (card_w + gap)
    # number
    text(s, x, y, card_w, Inches(0.4),
         num, size=10, bold=True, color=color, letter_spacing=200,
         align=PP_ALIGN.CENTER)
    # icon chip on the timeline
    chip_d = Inches(0.6)
    cx = x + (card_w - chip_d) / 2
    add_rect(s, cx, line_y - chip_d / 2,
             chip_d, chip_d, WHITE, corner=0.5)
    add_outline_rect(s, cx, line_y - chip_d / 2,
                     chip_d, chip_d,
                     line_rgb=color, line_pt=2.0, corner=0.5)
    text(s, cx, line_y - chip_d / 2, chip_d, chip_d,
         glyph, size=14, bold=True, color=color,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # card below
    cy = line_y + Inches(0.55)
    card(s, x, cy, card_w, card_h - Inches(0.6))
    text(s, x + Inches(0.2), cy + Inches(0.3),
         card_w - Inches(0.4), Inches(0.5),
         head, size=15, bold=True, color=INK,
         align=PP_ALIGN.CENTER, letter_spacing=-10)
    text(s, x + Inches(0.2), cy + Inches(0.95),
         card_w - Inches(0.4), Inches(1.5),
         body, size=10.5, color=GRAY_700, align=PP_ALIGN.CENTER)

# Footer note
text(s, Inches(0.75), Inches(6.65), Inches(12), Inches(0.4),
     "Per-user inbox tracks read state by signature — items re-surface if urgency shifts.",
     size=10, color=GRAY_500, align=PP_ALIGN.CENTER, italic=True)


# =====================================================================
# 10 — Architecture Diagram
# =====================================================================
s = add_slide()
page_chrome(s, 10, "Architecture", "A clean 3-tier monolith.",
            title_size=32)

# 3 tiers
tier_w = Inches(3.65)
tier_h = Inches(1.55)
gap = Inches(0.4)
total_w = tier_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(2.4)

tiers = [
    ("01", "Presentation", "Blade · Bootstrap 5 · Inter font",
     ["Browser-rendered views", "Light & dark theme", "Responsive layouts"]),
    ("02", "Application",  "Laravel 11 · PHP 8.2 · Eloquent",
     ["Controllers & middleware", "Mailables & console cmds", "Audit logging"]),
    ("03", "Data",         "MySQL 8 · Symfony Mailer",
     ["Relational storage", "SMTP delivery", "Cron scheduler"]),
]
for i, (num, head, sub, items) in enumerate(tiers):
    x = sx + i * (tier_w + gap)
    card(s, x, y, tier_w, tier_h, accent_top=True)
    text(s, x + Inches(0.3), y + Inches(0.25),
         tier_w - Inches(0.6), Inches(0.3),
         f"TIER {num}", size=9, bold=True, color=PRIMARY, letter_spacing=200)
    text(s, x + Inches(0.3), y + Inches(0.55),
         tier_w - Inches(0.6), Inches(0.5),
         head, size=18, bold=True, color=INK, letter_spacing=-15)
    text(s, x + Inches(0.3), y + Inches(1.05),
         tier_w - Inches(0.6), Inches(0.4),
         sub, size=10.5, color=GRAY_500)
    if i < 2:
        add_arrow(s, x + tier_w, y + tier_h / 2,
                  x + tier_w + gap - Inches(0.05),
                  y + tier_h / 2, rgb=GRAY_400, width=1.25)

# Below: detailed component list
below_y = y + tier_h + Inches(0.3)
for i, (num, head, sub, items) in enumerate(tiers):
    x = sx + i * (tier_w + gap)
    for j, item in enumerate(items):
        iy = below_y + j * Inches(0.4)
        add_rect(s, x, iy + Inches(0.13), Inches(0.06), Inches(0.06),
                 PRIMARY, corner=0.5)
        text(s, x + Inches(0.15), iy,
             tier_w - Inches(0.2), Inches(0.3),
             item, size=10.5, color=INK)

# Background services strip
svc_y = Inches(6.0)
svc_h = Inches(0.85)
add_gradient_rect(s, sx, svc_y, total_w, svc_h,
                  INK, INK_SOFT, angle_deg=0, corner=0.05)
text(s, sx + Inches(0.3), svc_y + Inches(0.1),
     Inches(7), Inches(0.25),
     "BACKGROUND SERVICES", size=9, bold=True,
     color=PRIMARY_SOFT, letter_spacing=200)

svcs = [
    ("Scheduler",   "app:check-expirations · 09:00 daily"),
    ("SMTP",        "Symfony Mailer · digest emails"),
    ("Activity Log","Per-user audit trail"),
]
sw = (total_w - Inches(0.6)) / 3
for i, (n_, d) in enumerate(svcs):
    x = sx + Inches(0.3) + i * sw
    text(s, x, svc_y + Inches(0.4), sw, Inches(0.25),
         n_, size=11, bold=True, color=WHITE)
    text(s, x, svc_y + Inches(0.6), sw, Inches(0.25),
         d, size=9, color=GRAY_300)


# =====================================================================
# 11 — Security / Smart Features
# =====================================================================
s = add_slide()
page_chrome(s, 11, "Security & Smart Features",
            "Designed with governance in mind.",
            title_size=30)

items = [
    ("◇", "RBAC enforced server-side",
     "Per-module × per-action permissions; admin/user/viewer roles enforced by middleware on every route.",
     PRIMARY, PRIMARY_SOFT),
    ("◆", "Encrypted credentials",
     "PC Master passwords and SMTP credentials encrypted at rest with Laravel Crypt (AES-256-CBC).",
     INDIGO, INDIGO_SOFT),
    ("◉", "Full audit trail",
     "Every create/update/delete logged with user, IP, user-agent, and field-level diff.",
     SUCCESS, SUCCESS_SOFT),
    ("◐", "Smart read tracking",
     "Notifications use a signature scheme (date + urgency bucket) — reads auto-invalidate if urgency shifts.",
     PURPLE, PURPLE_SOFT),
    ("●", "Staggered digest delivery",
     "Reminders batched per (module × day-mark) instead of spamming per item. Operators choose 10/20/30 days.",
     AMBER, AMBER_SOFT),
    ("▤", "CSRF & session hardening",
     "Tokens on every form, session regenerated on login, hashed passwords (bcrypt).",
     CYAN, PRIMARY_SOFT),
]
card_w = Inches(3.95)
card_h = Inches(1.95)
gap_x = Inches(0.2)
gap_y = Inches(0.25)
sx = Inches(0.7)
sy = Inches(2.4)
for i, (g, head, body, color, soft) in enumerate(items):
    row = i // 3
    col = i % 3
    x = sx + col * (card_w + gap_x)
    y = sy + row * (card_h + gap_y)
    card(s, x, y, card_w, card_h)
    icon_chip(s, x + Inches(0.3), y + Inches(0.3),
              g, size_in=0.5, bg=soft, fg=color)
    text(s, x + Inches(0.3), y + Inches(1.0),
         card_w - Inches(0.5), Inches(0.4),
         head, size=13, bold=True, color=INK)
    text(s, x + Inches(0.3), y + Inches(1.4),
         card_w - Inches(0.5), card_h - Inches(1.55),
         body, size=10, color=GRAY_700)


# =====================================================================
# 12 — Demo Screenshots
# =====================================================================
s = add_slide()
page_chrome(s, 12, "Live Screens", "Three views every user touches daily.",
            title_size=32)


def make_browser_mock(slide, mx, my, mw, mh):
    """Generic browser frame; returns inner-content rect (cx, cy, cw, ch)."""
    card(slide, mx, my, mw, mh)
    add_line(slide, mx, my + Inches(0.32),
             mx + mw, my + Inches(0.32), rgb=GRAY_200, width=0.5)
    for i, c in enumerate([ROSE, AMBER, SUCCESS]):
        cx = mx + Inches(0.1) + Inches(0.18) * i
        add_rect(slide, cx, my + Inches(0.1),
                 Inches(0.12), Inches(0.12), c, corner=0.5)
    return (mx + Inches(0.15), my + Inches(0.45),
            mw - Inches(0.3), mh - Inches(0.5))


mock_w = Inches(4.0)
mock_h = Inches(3.4)
gap = Inches(0.2)
total_w = mock_w * 3 + gap * 2
sx = (SW - total_w) / 2
sy = Inches(2.4)

# --- Mockup 1: Subscriptions index
mx = sx
my = sy
cx, cy, cw, ch = make_browser_mock(s, mx, my, mock_w, mock_h)
text(s, cx, cy, cw, Inches(0.3),
     "Subscriptions", size=11, bold=True, color=INK)
text(s, cx, cy + Inches(0.3), cw, Inches(0.2),
     "32 active · 5 expiring soon", size=7.5, color=GRAY_500)
# table header
ty = cy + Inches(0.65)
add_line(s, cx, ty + Inches(0.25), cx + cw, ty + Inches(0.25),
         rgb=GRAY_200, width=0.5)
for col_x, label in [(cx, "Name"),
                     (cx + cw * 0.45, "Vendor"),
                     (cx + cw * 0.7, "Expires"),
                     (cx + cw * 0.88, "")]:
    text(s, col_x, ty, Inches(1.5), Inches(0.2),
         label, size=7, bold=True, color=GRAY_500, letter_spacing=150)
# rows
rows = [
    ("Office 365 E3", "Microsoft",  "May 20", "Pending",  AMBER),
    ("Cloudflare Pro","Cloudflare", "May 25", "Pending",  AMBER),
    ("Adobe Creative","Adobe",      "May 31", "Pending",  AMBER),
    ("AWS Reserved",  "Amazon",     "Jun 07", "Active",   SUCCESS),
    ("GitHub Team",   "GitHub",     "Jun 13", "Active",   SUCCESS),
    ("Slack Pro",     "Slack",      "Jun 28", "Active",   SUCCESS),
]
ry = ty + Inches(0.35)
for nm, ven, exp, st, c in rows:
    text(s, cx, ry, cw * 0.45, Inches(0.2),
         nm, size=7.5, color=INK)
    text(s, cx + cw * 0.45, ry, cw * 0.25, Inches(0.2),
         ven, size=7.5, color=GRAY_500)
    text(s, cx + cw * 0.7, ry, cw * 0.18, Inches(0.2),
         exp, size=7.5, color=GRAY_500)
    add_rect(s, cx + cw * 0.88, ry,
             Inches(0.4), Inches(0.16),
             SUCCESS_SOFT if c == SUCCESS else AMBER_SOFT, corner=0.45)
    text(s, cx + cw * 0.88, ry, Inches(0.4), Inches(0.16),
         st, size=6.5, bold=True,
         color=SUCCESS if c == SUCCESS else AMBER,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    ry += Inches(0.28)
# caption
text(s, mx, my + mock_h + Inches(0.12),
     mock_w, Inches(0.3),
     "Subscriptions index", size=11, bold=True, color=INK,
     align=PP_ALIGN.CENTER)
text(s, mx, my + mock_h + Inches(0.38),
     mock_w, Inches(0.3),
     "Searchable, filterable, with bulk actions and Renew shortcut.",
     size=9.5, color=GRAY_500, align=PP_ALIGN.CENTER)

# --- Mockup 2: Notifications page
mx = sx + (mock_w + gap)
cx, cy, cw, ch = make_browser_mock(s, mx, sy, mock_w, mock_h)
text(s, cx, cy, cw, Inches(0.3),
     "Notifications", size=11, bold=True, color=INK)
text(s, cx, cy + Inches(0.3), cw, Inches(0.2),
     "8 unread · 12 total", size=7.5, color=GRAY_500)
# status chips
ty = cy + Inches(0.65)
chips = [("All", "12", True), ("Unread", "8", False), ("Read", "4", False)]
chx = cx
for nm, ct, active in chips:
    cw_chip = Inches(0.85)
    bg = PRIMARY_SOFT if active else GRAY_50
    fg = PRIMARY if active else GRAY_700
    add_rect(s, chx, ty, cw_chip, Inches(0.26), bg, corner=0.45)
    text(s, chx, ty, cw_chip - Inches(0.3), Inches(0.26),
         nm, size=7.5, bold=True, color=fg,
         align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, chx + cw_chip - Inches(0.28), ty,
         Inches(0.25), Inches(0.26),
         ct, size=7, bold=True, color=fg,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
    chx += cw_chip + Inches(0.08)

# notification items
ny = ty + Inches(0.45)
notifs = [
    ("Renewal Due: Office 365 E3", "3 days left", ROSE, ROSE_SOFT, True),
    ("Renewal Due: Cloudflare Pro","8 days left", AMBER, AMBER_SOFT, True),
    ("License Expiry: Adobe CC",   "14 days left",AMBER, AMBER_SOFT, False),
    ("Renewal Due: AWS Reserved",  "21 days left",PRIMARY, PRIMARY_SOFT, False),
]
for title, due, color, soft, unread in notifs:
    if unread:
        add_rect(s, cx, ny, Inches(0.04), Inches(0.5),
                 PRIMARY)
    add_rect(s, cx + Inches(0.12), ny + Inches(0.1),
             Inches(0.28), Inches(0.28), soft, corner=0.25)
    text(s, cx + Inches(0.12), ny + Inches(0.1),
         Inches(0.28), Inches(0.28), "●",
         size=10, bold=True, color=color,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, cx + Inches(0.5), ny + Inches(0.04),
         cw - Inches(1.2), Inches(0.22),
         title, size=7.5, bold=True, color=INK)
    text(s, cx + Inches(0.5), ny + Inches(0.24),
         cw - Inches(0.6), Inches(0.2),
         due, size=6.5, color=GRAY_500)
    ny += Inches(0.55)
# caption
text(s, mx, sy + mock_h + Inches(0.12),
     mock_w, Inches(0.3),
     "Notifications inbox", size=11, bold=True, color=INK,
     align=PP_ALIGN.CENTER)
text(s, mx, sy + mock_h + Inches(0.38),
     mock_w, Inches(0.3),
     "Per-user read state with signature-based invalidation.",
     size=9.5, color=GRAY_500, align=PP_ALIGN.CENTER)

# --- Mockup 3: Login screen
mx = sx + 2 * (mock_w + gap)
cx, cy, cw, ch = make_browser_mock(s, mx, sy, mock_w, mock_h)
# split: brand panel left, form right
bp_w = cw * 0.4
add_gradient_rect(s, cx, cy, bp_w, ch,
                  PRIMARY_DK, INDIGO, angle_deg=135)
add_rect(s, cx + Inches(0.18), cy + Inches(0.22),
         Inches(0.3), Inches(0.3), WHITE, corner=0.25)
text(s, cx + Inches(0.18), cy + Inches(0.22),
     Inches(0.3), Inches(0.3), "I",
     size=11, bold=True, color=PRIMARY,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, cx + Inches(0.18), cy + Inches(0.85),
     bp_w - Inches(0.3), Inches(0.3),
     "Welcome to ITAMS",
     size=10, bold=True, color=WHITE)
text(s, cx + Inches(0.18), cy + Inches(1.15),
     bp_w - Inches(0.3), Inches(0.6),
     "Track every IT asset, never miss a renewal.",
     size=7.5, color=WHITE)
# form right
fx = cx + bp_w + Inches(0.25)
fw = cw - bp_w - Inches(0.4)
text(s, fx, cy + Inches(0.5), fw, Inches(0.3),
     "Sign in", size=12, bold=True, color=INK)
text(s, fx, cy + Inches(0.85), fw, Inches(0.25),
     "Enter your credentials to continue.",
     size=7.5, color=GRAY_500)
# email
text(s, fx, cy + Inches(1.4), fw, Inches(0.18),
     "EMAIL", size=6.5, bold=True, color=GRAY_500, letter_spacing=200)
add_outline_rect(s, fx, cy + Inches(1.6),
                 fw, Inches(0.3),
                 line_rgb=GRAY_200, line_pt=0.5,
                 fill_rgb=WHITE, corner=0.15)
text(s, fx + Inches(0.12), cy + Inches(1.6),
     fw - Inches(0.2), Inches(0.3),
     "you@company.com", size=7.5, color=GRAY_400,
     anchor=MSO_ANCHOR.MIDDLE)
# password
text(s, fx, cy + Inches(2.05), fw, Inches(0.18),
     "PASSWORD", size=6.5, bold=True, color=GRAY_500, letter_spacing=200)
add_outline_rect(s, fx, cy + Inches(2.25),
                 fw, Inches(0.3),
                 line_rgb=GRAY_200, line_pt=0.5,
                 fill_rgb=WHITE, corner=0.15)
text(s, fx + Inches(0.12), cy + Inches(2.25),
     fw - Inches(0.2), Inches(0.3),
     "••••••••", size=7.5, color=GRAY_500,
     anchor=MSO_ANCHOR.MIDDLE)
# button
add_gradient_rect(s, fx, cy + Inches(2.75),
                  fw, Inches(0.32),
                  PRIMARY, INDIGO, angle_deg=0, corner=0.2)
text(s, fx, cy + Inches(2.75), fw, Inches(0.32),
     "Sign in", size=9, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# caption
text(s, mx, sy + mock_h + Inches(0.12),
     mock_w, Inches(0.3),
     "Login screen", size=11, bold=True, color=INK,
     align=PP_ALIGN.CENTER)
text(s, mx, sy + mock_h + Inches(0.38),
     mock_w, Inches(0.3),
     "Modern split layout with light/dark mode toggle.",
     size=9.5, color=GRAY_500, align=PP_ALIGN.CENTER)


# =====================================================================
# 13 — Benefits & Impact
# =====================================================================
s = add_slide()
page_chrome(s, 13, "Benefits & Impact", "What changes when ITAMS rolls out.",
            title_size=32)

metrics = [
    ("30%",  "lower IT spend",
     "Eliminates duplicate purchases & unused subs.",
     PRIMARY, PRIMARY_SOFT),
    ("0",    "missed renewals",
     "Automatic reminders for every tracked asset.",
     SUCCESS, SUCCESS_SOFT),
    ("100%", "audit coverage",
     "Every change logged with user, time and diff.",
     INDIGO, INDIGO_SOFT),
    ("5×",   "faster onboarding",
     "Bulk Excel import replaces manual entry.",
     AMBER, AMBER_SOFT),
]
card_w = Inches(2.95)
card_h = Inches(3.4)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
sx = (SW - total_w) / 2
y = Inches(2.6)
for i, (value, label, desc, color, soft) in enumerate(metrics):
    x = sx + i * (card_w + gap)
    card(s, x, y, card_w, card_h)
    add_gradient_rect(s, x, y, card_w, Inches(0.08),
                      (color[0], color[1], color[2]),
                      (INDIGO[0], INDIGO[1], INDIGO[2]), angle_deg=0)
    text(s, x + Inches(0.3), y + Inches(0.55),
         card_w - Inches(0.6), Inches(1.3),
         value, size=54, bold=True, color=color, letter_spacing=-25)
    add_line(s, x + Inches(0.3), y + Inches(2.0),
             x + Inches(0.8), y + Inches(2.0),
             rgb=color, width=1.5)
    text(s, x + Inches(0.3), y + Inches(2.15),
         card_w - Inches(0.6), Inches(0.45),
         label, size=13, bold=True, color=INK)
    text(s, x + Inches(0.3), y + Inches(2.6),
         card_w - Inches(0.6), Inches(0.8),
         desc, size=10.5, color=GRAY_700)


# =====================================================================
# 14 — Future Enhancements
# =====================================================================
s = add_slide()
page_chrome(s, 14, "Future Enhancements", "Where we're going next.",
            title_size=32)

quarters = [
    ("Q1", "Mobile app",
     "Native iOS / Android for on-the-go asset check-in & QR scanning.",
     PRIMARY, PRIMARY_SOFT),
    ("Q2", "Integrations",
     "Slack, Microsoft Teams, and Jira Service Desk webhooks.",
     INDIGO, INDIGO_SOFT),
    ("Q3", "AI insights",
     "Forecast renewal costs and recommend subscription consolidation.",
     PURPLE, PURPLE_SOFT),
    ("Q4", "Multi-tenant SaaS",
     "Organization isolation, SSO/SAML, per-tenant billing.",
     SUCCESS, SUCCESS_SOFT),
]
n = len(quarters)
col_w = Inches(2.95)
gap = Inches(0.2)
total_w = col_w * n + gap * (n - 1)
sx = (SW - total_w) / 2
y = Inches(2.7)
line_y = y + Inches(0.7)
add_gradient_rect(s, sx + Inches(0.2), line_y - Inches(0.02),
                  total_w - Inches(0.4), Inches(0.04),
                  PRIMARY, INDIGO, angle_deg=0)
for i, (q, head, body, color, soft) in enumerate(quarters):
    x = sx + i * (col_w + gap)
    # quarter pill
    add_rect(s, x + col_w / 2 - Inches(0.45),
             y, Inches(0.9), Inches(0.42),
             color, corner=0.45)
    text(s, x + col_w / 2 - Inches(0.45), y,
         Inches(0.9), Inches(0.42), q,
         size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         letter_spacing=150)
    # dot
    add_rect(s, x + col_w / 2 - Inches(0.1),
             line_y - Inches(0.1),
             Inches(0.2), Inches(0.2), color, corner=0.5)
    add_rect(s, x + col_w / 2 - Inches(0.06),
             line_y - Inches(0.06),
             Inches(0.12), Inches(0.12), WHITE, corner=0.5)
    # card
    cy = line_y + Inches(0.35)
    card(s, x, cy, col_w, Inches(2.5))
    text(s, x + Inches(0.25), cy + Inches(0.35),
         col_w - Inches(0.5), Inches(0.5),
         head, size=15, bold=True, color=INK,
         align=PP_ALIGN.CENTER, letter_spacing=-10)
    text(s, x + Inches(0.25), cy + Inches(0.95),
         col_w - Inches(0.5), Inches(1.4),
         body, size=10.5, color=GRAY_700, align=PP_ALIGN.CENTER)


# =====================================================================
# 15 — Thank You / Q&A
# =====================================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, GRAY_50)
# big gradient hero block
add_gradient_rect(s, 0, 0, SW, SH,
                  PRIMARY_DK, INDIGO, angle_deg=135)
# atmospheric tint
add_gradient_rect(s, SW - Inches(8), Inches(-1),
                  Inches(10), Inches(10),
                  PRIMARY, INDIGO,
                  angle_deg=135,
                  alpha_start=25000, alpha_end=0)

text(s, Inches(0), Inches(2.2), SW, Inches(0.4),
     "HACKATHON 2026",
     size=10, bold=True, color=WHITE, letter_spacing=400,
     align=PP_ALIGN.CENTER)

text(s, Inches(0), Inches(2.8), SW, Inches(2.0),
     "Thank you.",
     size=110, bold=True, color=WHITE, letter_spacing=-30,
     align=PP_ALIGN.CENTER)

# divider
add_rect(s, (SW - Inches(1.5)) / 2, Inches(5.05),
         Inches(1.5), Inches(0.06), WHITE, corner=0.5)

text(s, Inches(0), Inches(5.25), SW, Inches(0.5),
     "Questions & discussion", size=22, color=WHITE,
     align=PP_ALIGN.CENTER)

# Contact card
cc_w = Inches(7)
cc_h = Inches(1.0)
cc_x = (SW - cc_w) / 2
cc_y = Inches(6.05)
add_outline_rect(s, cc_x, cc_y, cc_w, cc_h,
                 line_rgb=WHITE, line_pt=0.5,
                 fill_rgb=None, corner=0.08)
text(s, cc_x, cc_y + Inches(0.2),
     cc_w, Inches(0.3),
     "TEAM ITAMS  ·  IT ASSET MANAGEMENT SYSTEM",
     size=10, bold=True, color=WHITE, letter_spacing=300,
     align=PP_ALIGN.CENTER)
text(s, cc_x, cc_y + Inches(0.55),
     cc_w, Inches(0.3),
     "[ contact email ]   ·   [ repo link ]",
     size=11, color=WHITE, align=PP_ALIGN.CENTER)


# =====================================================================
# Save
# =====================================================================
out = r"D:\xampp\htdocs\itams\presentation\ITAMS_Hackathon_Branded.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
