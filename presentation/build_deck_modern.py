"""Build the ITAMS hackathon presentation — modern SaaS pitch style.

Sits between the strict minimal deck and the decorative one:
subtle gradient accents, card-based layouts, consistent icon set.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree


# ---------- Palette ----------
INK        = RGBColor(0x0F, 0x17, 0x2A)
INK_SOFT   = RGBColor(0x1E, 0x29, 0x3B)
GRAY_700   = RGBColor(0x47, 0x55, 0x69)
GRAY_500   = RGBColor(0x64, 0x74, 0x8B)
GRAY_400   = RGBColor(0x94, 0xA3, 0xB8)
GRAY_300   = RGBColor(0xCB, 0xD5, 0xE1)
GRAY_200   = RGBColor(0xE2, 0xE8, 0xF0)
GRAY_100   = RGBColor(0xF1, 0xF5, 0xF9)
GRAY_50    = RGBColor(0xF8, 0xFA, 0xFC)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT     = RGBColor(0x25, 0x63, 0xEB)
ACCENT_DK  = RGBColor(0x1D, 0x4E, 0xD8)
ACCENT_INDIGO = RGBColor(0x4F, 0x46, 0xE5)
ACCENT_SOFT = RGBColor(0xDB, 0xEA, 0xFE)   # pale blue for subtle highlights

FONT = "Calibri"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def add_slide():
    return prs.slides.add_slide(BLANK)


# ---------- Fill helpers ----------
def add_rect(slide, x, y, w, h, rgb, *, no_line=True, corner=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
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
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
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


def add_gradient_rect(slide, x, y, w, h, color_start, color_end,
                      *, angle_deg=0, corner=None):
    """Subtle linear gradient via direct XML."""
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    if corner:
        shp.adjustments[0] = corner
    # Remove default fill, add gradient via XML
    spPr = shp.fill._xPr
    for tag in ('a:solidFill', 'a:gradFill', 'a:noFill', 'a:pattFill', 'a:blipFill'):
        existing = spPr.find(qn(tag))
        if existing is not None:
            spPr.remove(existing)

    gradFill = etree.SubElement(spPr, qn('a:gradFill'))
    gradFill.set('flip', 'none')
    gradFill.set('rotWithShape', '1')
    # need to insert gradFill BEFORE any a:ln element
    ln = spPr.find(qn('a:ln'))
    if ln is not None:
        spPr.remove(gradFill)
        spPr.insert(list(spPr).index(ln), gradFill)

    gsLst = etree.SubElement(gradFill, qn('a:gsLst'))
    for pos, color in [(0, color_start), (100000, color_end)]:
        gs = etree.SubElement(gsLst, qn('a:gs'))
        gs.set('pos', str(pos))
        srgb = etree.SubElement(gs, qn('a:srgbClr'))
        srgb.set('val', f'{color[0]:02X}{color[1]:02X}{color[2]:02X}')

    lin = etree.SubElement(gradFill, qn('a:lin'))
    lin.set('ang', str(int(angle_deg * 60000)))
    lin.set('scaled', '0')

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
    tail.set("type", "triangle")
    tail.set("w", "sm")
    tail.set("h", "sm")
    return line


def text(slide, x, y, w, h, content, *,
         size=11, bold=False, color=INK,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         font=FONT, letter_spacing=None):
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
        run.font.color.rgb = color
        run.font.name = font
        if letter_spacing:
            rPr = run._r.get_or_add_rPr()
            rPr.set("spc", str(letter_spacing))
    return tb


def eyebrow(slide, x, y, label, *, color=GRAY_500):
    text(slide, x, y, Inches(6), Inches(0.3), label.upper(),
         size=9, bold=True, color=color, letter_spacing=200)


# ---------- Consistent icon chip ----------
def icon_chip(slide, x, y, glyph, *, size_in=0.42, soft=True):
    """Rounded-square icon container, soft tinted bg, accent glyph."""
    s = Inches(size_in)
    if soft:
        add_rect(slide, x, y, s, s, ACCENT_SOFT, corner=0.25)
    else:
        add_rect(slide, x, y, s, s, ACCENT, corner=0.25)
    text(slide, x, y, s, s, glyph,
         size=int(size_in * 32), bold=True,
         color=ACCENT if soft else WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---------- Slide chrome ----------
def page_chrome(slide, page_no, title_str, eyebrow_str):
    # white background
    add_rect(slide, 0, 0, SW, SH, WHITE)
    # eyebrow with accent dot
    add_rect(slide, Inches(0.75), Inches(0.62), Inches(0.08), Inches(0.08),
             ACCENT, corner=0.5)
    eyebrow(slide, Inches(0.92), Inches(0.55), eyebrow_str, color=ACCENT)
    # title
    text(slide, Inches(0.75), Inches(0.85), Inches(12), Inches(0.7),
         title_str, size=32, bold=True, color=INK)
    # gradient accent stripe under title
    add_gradient_rect(slide, Inches(0.75), Inches(1.75),
                      Inches(2.0), Inches(0.06),
                      (0x25, 0x63, 0xEB), (0x4F, 0x46, 0xE5),
                      angle_deg=0, corner=0.5)
    # page number top right
    text(slide, SW - Inches(1.5), Inches(0.55), Inches(0.85), Inches(0.3),
         f"{page_no:02d} / 14", size=9, bold=True, color=GRAY_400,
         align=PP_ALIGN.RIGHT, letter_spacing=150)
    # wordmark bottom
    add_line(slide, Inches(0.75), SH - Inches(0.55),
             SW - Inches(0.75), SH - Inches(0.55),
             rgb=GRAY_200, width=0.5)
    text(slide, Inches(0.75), SH - Inches(0.45), Inches(6), Inches(0.3),
         "ITAMS", size=9, bold=True, color=INK, letter_spacing=200)
    text(slide, Inches(1.25), SH - Inches(0.45), Inches(8), Inches(0.3),
         "IT Asset Management System", size=9, color=GRAY_500)


def soft_card(slide, x, y, w, h):
    """Subtle card: pale gray fill + hairline border + soft top accent strip."""
    add_outline_rect(slide, x, y, w, h,
                     line_rgb=GRAY_200, line_pt=0.5,
                     fill_rgb=GRAY_50, corner=0.04)
    return None


# ====================================================================
# 01 — Title
# ====================================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, INK)
# subtle gradient overlay at top
add_gradient_rect(s, 0, 0, SW, Inches(3.5),
                  (0x1D, 0x4E, 0xD8), (0x0F, 0x17, 0x2A),
                  angle_deg=90)
# accent gradient sliver
add_gradient_rect(s, 0, Inches(3.5), SW, Inches(0.04),
                  (0x25, 0x63, 0xEB), (0x4F, 0x46, 0xE5),
                  angle_deg=0)

# header
text(s, Inches(0.75), Inches(0.6), Inches(8), Inches(0.3),
     "HACKATHON 2026   ·   IT INFRASTRUCTURE TRACK",
     size=9, bold=True, color=GRAY_400, letter_spacing=300)

# logo monogram
add_outline_rect(s, Inches(0.75), Inches(1.05),
                 Inches(0.55), Inches(0.55),
                 line_rgb=WHITE, line_pt=1.2,
                 corner=0.18)
text(s, Inches(0.75), Inches(1.05), Inches(0.55), Inches(0.55),
     "I", size=18, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(1.45), Inches(1.13), Inches(6), Inches(0.4),
     "ITAMS", size=14, bold=True, color=WHITE, letter_spacing=150)

# big title
text(s, Inches(0.75), Inches(2.5), Inches(12), Inches(1.5),
     "ITAMS", size=112, bold=True, color=WHITE)

# divider gradient
add_gradient_rect(s, Inches(0.78), Inches(4.4),
                  Inches(2.4), Inches(0.06),
                  (0x60, 0xA5, 0xFA), (0x4F, 0x46, 0xE5),
                  angle_deg=0, corner=0.5)

# subtitle / tagline
text(s, Inches(0.75), Inches(4.6), Inches(11), Inches(0.6),
     "An enterprise IT asset management platform.",
     size=22, color=GRAY_300)
text(s, Inches(0.75), Inches(5.15), Inches(11), Inches(0.5),
     "Track devices, licenses & subscriptions — never miss a renewal.",
     size=14, color=GRAY_500)

# bottom block: team + date
text(s, Inches(0.75), SH - Inches(1.2), Inches(5), Inches(0.25),
     "PRESENTED BY", size=8, bold=True, color=GRAY_500, letter_spacing=300)
text(s, Inches(0.75), SH - Inches(0.85), Inches(8), Inches(0.45),
     "Team ITAMS", size=18, bold=True, color=WHITE)

text(s, SW - Inches(5), SH - Inches(1.2), Inches(4), Inches(0.25),
     "DATE", size=8, bold=True, color=GRAY_500, letter_spacing=300,
     align=PP_ALIGN.RIGHT)
text(s, SW - Inches(5), SH - Inches(0.85), Inches(4), Inches(0.4),
     "May 2026", size=14, color=GRAY_300, align=PP_ALIGN.RIGHT)


# ====================================================================
# 02 — Problem Statement
# ====================================================================
s = add_slide()
page_chrome(s, 2, "The problem", "Problem Statement")

# lead
text(s, Inches(0.75), Inches(2.35), Inches(11.8), Inches(1.0),
     "IT teams lose visibility and money\nmanaging assets in spreadsheets.",
     size=26, bold=True, color=INK)

# 4 cards
points = [
    ("◆", "Manual tracking",
     "Devices, licenses & subscriptions scattered across disconnected files."),
    ("◐", "Missed renewals",
     "Contracts expire or auto-renew without warning, causing downtime or waste."),
    ("◇", "No audit trail",
     "No record of who changed what, when — failing internal compliance reviews."),
    ("●", "Hidden cost",
     "Duplicate purchases and unused subscriptions inflate IT spend 20–30%."),
]
card_w = Inches(5.85)
card_h = Inches(1.4)
gap_x = Inches(0.2)
gap_y = Inches(0.2)
sx = (SW - card_w * 2 - gap_x) / 2
sy = Inches(4.4)

for i, (glyph, head, body) in enumerate(points):
    row = i // 2
    col = i % 2
    x = sx + col * (card_w + gap_x)
    y = sy + row * (card_h + gap_y)
    soft_card(s, x, y, card_w, card_h)
    icon_chip(s, x + Inches(0.3), y + Inches(0.3), glyph)
    text(s, x + Inches(0.95), y + Inches(0.28), card_w - Inches(1.1),
         Inches(0.35), head, size=14, bold=True, color=INK)
    text(s, x + Inches(0.95), y + Inches(0.65), card_w - Inches(1.1),
         Inches(0.7), body, size=11, color=GRAY_700)


# ====================================================================
# 03 — Proposed Solution
# ====================================================================
s = add_slide()
page_chrome(s, 3, "Our solution", "Proposed Solution")

text(s, Inches(0.75), Inches(2.35), Inches(11.5), Inches(1.0),
     "A single platform for the entire\nIT asset lifecycle.",
     size=26, bold=True, color=INK)
text(s, Inches(0.75), Inches(3.95), Inches(11.5), Inches(0.5),
     "ITAMS unifies hardware, devices, licenses, and subscriptions — with automated "
     "renewal reminders, role-based access, and a full audit log.",
     size=13, color=GRAY_700)

pillars = [
    ("◆", "Unify",    "One source of truth for every IT asset across the organisation."),
    ("◉", "Automate", "Email and in-app reminders fire before things expire."),
    ("◇", "Govern",   "Role-based access with full activity history per user."),
]
card_w = Inches(3.95)
card_h = Inches(2.3)
gap = Inches(0.2)
total_w = card_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(4.9)

for i, (glyph, head, body) in enumerate(pillars):
    x = sx + i * (card_w + gap)
    soft_card(s, x, y, card_w, card_h)
    # subtle gradient top strip
    add_gradient_rect(s, x, y, card_w, Inches(0.06),
                      (0x25, 0x63, 0xEB), (0x4F, 0x46, 0xE5),
                      angle_deg=0)
    icon_chip(s, x + Inches(0.35), y + Inches(0.3), glyph, size_in=0.5)
    text(s, x + Inches(0.35), y + Inches(1.0), card_w - Inches(0.6),
         Inches(0.45), head, size=18, bold=True, color=INK)
    text(s, x + Inches(0.35), y + Inches(1.5), card_w - Inches(0.6),
         Inches(0.7), body, size=11, color=GRAY_700)


# ====================================================================
# 04 — System Overview
# ====================================================================
s = add_slide()
page_chrome(s, 4, "System overview", "Architecture")

# 3 tiers with gradient top strips
tier_w = Inches(3.6)
tier_h = Inches(1.6)
gap = Inches(0.4)
total_w = tier_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(2.5)

tiers = [
    ("01", "Presentation", "Browser · Blade · Bootstrap 5"),
    ("02", "Application",  "Laravel 11 · Controllers · Mailables"),
    ("03", "Data",         "MySQL 8 · Eloquent ORM"),
]
for i, (num, head, sub) in enumerate(tiers):
    x = sx + i * (tier_w + gap)
    soft_card(s, x, y, tier_w, tier_h)
    add_gradient_rect(s, x, y, tier_w, Inches(0.08),
                      (0x25, 0x63, 0xEB), (0x4F, 0x46, 0xE5),
                      angle_deg=0)
    text(s, x + Inches(0.3), y + Inches(0.3), tier_w, Inches(0.3),
         f"TIER {num}", size=9, bold=True, color=ACCENT, letter_spacing=200)
    text(s, x + Inches(0.3), y + Inches(0.6), tier_w - Inches(0.6),
         Inches(0.45), head, size=18, bold=True, color=INK)
    text(s, x + Inches(0.3), y + Inches(1.1), tier_w - Inches(0.6),
         Inches(0.4), sub, size=10.5, color=GRAY_700)
    if i < 2:
        add_arrow(s, x + tier_w, y + tier_h / 2,
                  x + tier_w + gap - Inches(0.05), y + tier_h / 2,
                  rgb=GRAY_400, width=1.25)

# background services row
bg_y = Inches(4.4)
bg_h = Inches(1.05)
soft_card(s, sx, bg_y, total_w, bg_h)
text(s, sx + Inches(0.35), bg_y + Inches(0.15), Inches(5), Inches(0.3),
     "BACKGROUND SERVICES", size=9, bold=True, color=ACCENT,
     letter_spacing=200)

services = [
    ("◉", "Scheduler",   "app:check-expirations · daily 09:00"),
    ("◇", "SMTP",        "Renewal reminder emails to recipients"),
    ("◆", "Activity Log","Per-user audit trail of every change"),
]
svc_w = (total_w - Inches(0.7)) / 3
for i, (g, head, sub) in enumerate(services):
    x = sx + Inches(0.35) + i * svc_w
    icon_chip(s, x, bg_y + Inches(0.5), g, size_in=0.34)
    text(s, x + Inches(0.5), bg_y + Inches(0.48), svc_w - Inches(0.5),
         Inches(0.25), head, size=11, bold=True, color=INK)
    text(s, x + Inches(0.5), bg_y + Inches(0.72), svc_w - Inches(0.5),
         Inches(0.3), sub, size=9.5, color=GRAY_500)

# module strip
mod_y = Inches(5.8)
text(s, sx, mod_y - Inches(0.3), Inches(8), Inches(0.3),
     "FUNCTIONAL MODULES", size=9, bold=True, color=GRAY_500,
     letter_spacing=200)
modules = ["PC Master", "Device Master", "Subscriptions",
           "Licenses & Contracts", "Notifications", "User Management"]
mx = sx
for m in modules:
    w = Inches(0.4 + 0.105 * len(m))
    add_outline_rect(s, mx, mod_y, w, Inches(0.5),
                     line_rgb=GRAY_200, line_pt=0.5,
                     fill_rgb=WHITE, corner=0.25)
    text(s, mx, mod_y, w, Inches(0.5), m,
         size=10, color=INK, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    mx += w + Inches(0.12)


# ====================================================================
# 05 — Key Features (6 with icons)
# ====================================================================
s = add_slide()
page_chrome(s, 5, "Key features", "Capabilities")

features = [
    ("◆", "Asset inventory",
     "PCs, devices, software licenses and subscriptions in one place."),
    ("◉", "Smart reminders",
     "Daily scheduler emails configurable recipients before renewals lapse."),
    ("◇", "Role-based access",
     "Admin / per-module view & edit permissions enforced by middleware."),
    ("◐", "Audit trail",
     "Every create, update, delete logged with user, time and field diff."),
    ("▤", "Bulk import/export",
     "Excel import & export per module, with downloadable templates."),
    ("✉", "Configurable mail",
     "Switch between .env SMTP and database settings without redeploying."),
]
card_w = Inches(3.95)
card_h = Inches(1.85)
gap_x = Inches(0.2)
gap_y = Inches(0.25)
sx = Inches(0.7)
sy = Inches(2.4)

for i, (g, head, body) in enumerate(features):
    row = i // 3
    col = i % 3
    x = sx + col * (card_w + gap_x)
    y = sy + row * (card_h + gap_y)
    soft_card(s, x, y, card_w, card_h)
    icon_chip(s, x + Inches(0.3), y + Inches(0.3), g, size_in=0.46)
    text(s, x + Inches(0.3), y + Inches(0.95), card_w - Inches(0.5),
         Inches(0.4), head, size=14, bold=True, color=INK)
    text(s, x + Inches(0.3), y + Inches(1.35), card_w - Inches(0.5),
         Inches(0.6), body, size=10.5, color=GRAY_700)


# ====================================================================
# 06 — Technology Stack
# ====================================================================
s = add_slide()
page_chrome(s, 6, "Technology stack", "Built with")

groups = [
    ("◆", "Frontend", [
        ("Bootstrap 5.3",    "UI & responsive grid"),
        ("Bootstrap Icons",  "Iconography"),
        ("Inter font",       "Enterprise typeface"),
        ("Vanilla JS",       "Lightweight interactivity"),
    ]),
    ("◉", "Backend", [
        ("Laravel 11",       "MVC framework, Eloquent ORM"),
        ("PHP 8.2+",         "Typed properties, modern syntax"),
        ("Blade",            "Server-rendered templating"),
        ("Maatwebsite/Excel","Import & export to .xlsx"),
    ]),
    ("◇", "Data & Infra", [
        ("MySQL 8 / MariaDB","Relational storage"),
        ("Symfony Mailer",   "SMTP delivery via Laravel"),
        ("XAMPP / Apache",   "On-prem hosting"),
        ("Cron scheduler",   "Daily expiry checks"),
    ]),
]
col_w = Inches(4.0)
col_h = Inches(4.2)
gap = Inches(0.2)
total_w = col_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(2.4)

for i, (glyph, head, items) in enumerate(groups):
    x = sx + i * (col_w + gap)
    soft_card(s, x, y, col_w, col_h)
    # gradient header strip
    add_gradient_rect(s, x, y, col_w, Inches(0.08),
                      (0x25, 0x63, 0xEB), (0x4F, 0x46, 0xE5),
                      angle_deg=0)
    icon_chip(s, x + Inches(0.3), y + Inches(0.25), glyph, size_in=0.4)
    text(s, x + Inches(0.85), y + Inches(0.3), col_w - Inches(0.9),
         Inches(0.4), head, size=14, bold=True, color=INK)

    iy = y + Inches(1.05)
    for label, desc in items:
        add_line(s, x + Inches(0.3), iy + Inches(0.12),
                 x + Inches(0.45), iy + Inches(0.12),
                 rgb=ACCENT, width=1.5)
        text(s, x + Inches(0.55), iy - Inches(0.02), col_w - Inches(0.7),
             Inches(0.3), label, size=11, bold=True, color=INK)
        text(s, x + Inches(0.55), iy + Inches(0.25), col_w - Inches(0.7),
             Inches(0.3), desc, size=9.5, color=GRAY_500)
        iy += Inches(0.72)


# ====================================================================
# 07 — Workflow
# ====================================================================
s = add_slide()
page_chrome(s, 7, "How it works", "Workflow")

steps = [
    ("01", "Onboard",  "Admin adds assets or imports an .xlsx file."),
    ("02", "Track",    "Modules show real-time status across the estate."),
    ("03", "Detect",   "Daily scheduler scans for items in the renewal window."),
    ("04", "Notify",   "Email + in-app alerts fire to configured recipients."),
    ("05", "Renew",    "Owners renew or terminate; status auto-updates."),
]
n = len(steps)
card_w = Inches(2.32)
card_h = Inches(2.6)
gap = Inches(0.12)
total_w = card_w * n + gap * (n - 1)
sx = (SW - total_w) / 2
y = Inches(2.8)

# horizontal gradient line across the row
mid_y = y + Inches(0.4)
add_gradient_rect(s, sx + Inches(0.2), mid_y - Inches(0.02),
                  total_w - Inches(0.4), Inches(0.04),
                  (0x60, 0xA5, 0xFA), (0x4F, 0x46, 0xE5),
                  angle_deg=0)

for i, (num, head, body) in enumerate(steps):
    x = sx + i * (card_w + gap)
    # marker on the line
    add_rect(s, x + card_w / 2 - Inches(0.1),
             mid_y - Inches(0.1), Inches(0.2), Inches(0.2),
             ACCENT, corner=0.5)
    add_rect(s, x + card_w / 2 - Inches(0.06),
             mid_y - Inches(0.06), Inches(0.12), Inches(0.12),
             WHITE, corner=0.5)
    # step number above
    text(s, x, y - Inches(0.05), card_w, Inches(0.4),
         num, size=11, bold=True, color=ACCENT,
         align=PP_ALIGN.CENTER, letter_spacing=200)
    # card below
    cy = mid_y + Inches(0.35)
    soft_card(s, x, cy, card_w, card_h - Inches(0.4))
    text(s, x + Inches(0.25), cy + Inches(0.3), card_w - Inches(0.5),
         Inches(0.4), head, size=15, bold=True, color=INK,
         align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.25), cy + Inches(0.8), card_w - Inches(0.5),
         Inches(1.5), body, size=10.5, color=GRAY_700,
         align=PP_ALIGN.CENTER)

text(s, Inches(0.75), Inches(6.4), Inches(12), Inches(0.4),
     "Per-day dedupe prevents the same item being re-notified more than once daily.",
     size=10, color=GRAY_500, align=PP_ALIGN.CENTER)


# ====================================================================
# 08 — UI/UX Overview
# ====================================================================
s = add_slide()
page_chrome(s, 8, "Interface", "UI / UX Overview")

# left column
text(s, Inches(0.75), Inches(2.4), Inches(6), Inches(0.5),
     "Designed for daily IT operations.",
     size=20, bold=True, color=INK)

principles = [
    ("◆", "Modern dashboard",   "Liquid-glass cards over soft gradient surfaces."),
    ("◉", "Light & dark mode",  "Theme persists across sessions per user."),
    ("◇", "Inter typography",   "Tight letter-spacing, optimized for scanning."),
    ("◐", "Accessibility",      "Keyboard nav, ARIA labels, visible focus rings."),
    ("●", "Responsive",         "Sidebar rail on tablet; full reflow on mobile."),
]
y = Inches(3.2)
for glyph, label, desc in principles:
    icon_chip(s, Inches(0.75), y, glyph, size_in=0.34)
    text(s, Inches(1.25), y - Inches(0.02), Inches(5.4), Inches(0.3),
         label, size=12, bold=True, color=INK)
    text(s, Inches(1.25), y + Inches(0.27), Inches(5.4), Inches(0.4),
         desc, size=10.5, color=GRAY_700)
    y += Inches(0.62)

# right: outlined browser mockup with soft fill
mx = Inches(7.5)
my = Inches(2.4)
mw = Inches(5.2)
mh = Inches(4.2)

soft_card(s, mx, my, mw, mh)
# top bar
add_line(s, mx, my + Inches(0.42), mx + mw, my + Inches(0.42),
         rgb=GRAY_200, width=0.5)
for i in range(3):
    cx = mx + Inches(0.18) + Inches(0.24) * i
    add_outline_rect(s, cx, my + Inches(0.14),
                     Inches(0.16), Inches(0.16),
                     line_rgb=GRAY_300, line_pt=0.5, corner=0.5)
# URL pill
add_outline_rect(s, mx + Inches(1.0), my + Inches(0.1),
                 Inches(4.0), Inches(0.26),
                 line_rgb=GRAY_200, line_pt=0.5,
                 fill_rgb=WHITE, corner=0.3)
text(s, mx + Inches(1.1), my + Inches(0.1),
     Inches(3.8), Inches(0.26),
     "itams.local / dashboard",
     size=8.5, color=GRAY_500, anchor=MSO_ANCHOR.MIDDLE)

# sidebar
add_rect(s, mx, my + Inches(0.42), Inches(1.05),
         mh - Inches(0.42), GRAY_50)
add_line(s, mx + Inches(1.05), my + Inches(0.42),
         mx + Inches(1.05), my + mh,
         rgb=GRAY_200, width=0.5)
add_outline_rect(s, mx + Inches(0.13), my + Inches(0.58),
                 Inches(0.32), Inches(0.32), line_rgb=ACCENT, line_pt=1.2,
                 corner=0.25)
text(s, mx + Inches(0.13), my + Inches(0.58),
     Inches(0.32), Inches(0.32), "I", size=10, bold=True, color=ACCENT,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, mx + Inches(0.55), my + Inches(0.62),
     Inches(0.55), Inches(0.25),
     "ITAMS", size=9, bold=True, color=INK, letter_spacing=150)

nav = ["Dashboard", "PC Master", "Devices",
       "Subscriptions", "Licenses", "Activity"]
for i, name in enumerate(nav):
    ny = my + Inches(1.1) + Inches(0.32) * i
    if i == 0:
        add_rect(s, mx + Inches(0.1), ny - Inches(0.04),
                 Inches(0.9), Inches(0.28), ACCENT_SOFT, corner=0.25)
        text(s, mx + Inches(0.2), ny, Inches(0.8), Inches(0.22),
             name, size=8.5, bold=True, color=ACCENT,
             anchor=MSO_ANCHOR.MIDDLE)
    else:
        text(s, mx + Inches(0.2), ny, Inches(0.8), Inches(0.22),
             name, size=8.5, color=GRAY_500,
             anchor=MSO_ANCHOR.MIDDLE)

# main area: KPI tiles
kpi_x = mx + Inches(1.25)
kpi_y = my + Inches(0.7)
kpi_w = (mw - Inches(1.4)) / 4 - Inches(0.07)
kpi_h = Inches(0.9)
kpi = [("128", "PCs"), ("64", "Devices"),
       ("32", "Active"), ("5", "Expiring")]
for i, (v, l) in enumerate(kpi):
    x = kpi_x + i * (kpi_w + Inches(0.08))
    add_outline_rect(s, x, kpi_y, kpi_w, kpi_h,
                     line_rgb=GRAY_200, line_pt=0.5,
                     fill_rgb=WHITE, corner=0.1)
    add_gradient_rect(s, x, kpi_y, kpi_w, Inches(0.05),
                      (0x60, 0xA5, 0xFA), (0x4F, 0x46, 0xE5),
                      angle_deg=0)
    text(s, x + Inches(0.15), kpi_y + Inches(0.13),
         kpi_w - Inches(0.2), Inches(0.35),
         v, size=15, bold=True, color=INK)
    text(s, x + Inches(0.15), kpi_y + Inches(0.55),
         kpi_w - Inches(0.2), Inches(0.3),
         l.upper(), size=7.5, color=GRAY_500, letter_spacing=200)

# chart placeholder
chart_y = kpi_y + kpi_h + Inches(0.2)
chart_h = my + mh - chart_y - Inches(0.25)
add_outline_rect(s, kpi_x, chart_y,
                 mw - Inches(1.4), chart_h,
                 line_rgb=GRAY_200, line_pt=0.5,
                 fill_rgb=WHITE, corner=0.05)
heights = [0.42, 0.66, 0.55, 0.85, 0.6, 0.95, 0.7, 0.5, 0.78]
n_bars = len(heights)
bar_w = Inches(0.18)
bar_area = mw - Inches(1.8)
gap_b = (bar_area - bar_w * n_bars) / (n_bars - 1)
base = chart_y + chart_h - Inches(0.25)
for i in range(n_bars):
    h = Inches(heights[i] * 1.5)
    bx = kpi_x + Inches(0.2) + i * (bar_w + gap_b)
    # gradient bar
    add_gradient_rect(s, bx, base - h, bar_w, h,
                      (0x4F, 0x46, 0xE5), (0x25, 0x63, 0xEB),
                      angle_deg=90, corner=0.2)


# ====================================================================
# 09 — Benefits / Impact
# ====================================================================
s = add_slide()
page_chrome(s, 9, "Benefits & impact", "Outcomes")

metrics = [
    ("30%",  "lower IT spend",
     "Eliminates duplicate purchases & unused subscriptions."),
    ("0",    "missed renewals",
     "Automatic reminders for every tracked asset."),
    ("100%", "audit coverage",
     "Every change logged with user, time, and diff."),
    ("5×",   "faster onboarding",
     "Bulk Excel import replaces manual entry."),
]
card_w = Inches(2.95)
card_h = Inches(3.4)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
sx = (SW - total_w) / 2
y = Inches(2.6)

for i, (value, label, desc) in enumerate(metrics):
    x = sx + i * (card_w + gap)
    soft_card(s, x, y, card_w, card_h)
    add_gradient_rect(s, x, y, card_w, Inches(0.08),
                      (0x25, 0x63, 0xEB), (0x4F, 0x46, 0xE5),
                      angle_deg=0)
    text(s, x + Inches(0.3), y + Inches(0.5), card_w - Inches(0.6),
         Inches(1.3), value, size=54, bold=True, color=INK)
    add_line(s, x + Inches(0.3), y + Inches(1.95),
             x + Inches(0.8), y + Inches(1.95), rgb=ACCENT, width=1.5)
    text(s, x + Inches(0.3), y + Inches(2.1), card_w - Inches(0.6),
         Inches(0.45), label, size=13, bold=True, color=INK)
    text(s, x + Inches(0.3), y + Inches(2.55), card_w - Inches(0.6),
         Inches(0.8), desc, size=10.5, color=GRAY_700)


# ====================================================================
# 10 — Challenges
# ====================================================================
s = add_slide()
page_chrome(s, 10, "Challenges we faced", "Challenges")

challenges = [
    ("Mail deliverability",
     "Some SMTP servers reject mailbox addresses without warning. Solved with graceful per-recipient error handling and configurable lists."),
    ("Notification dedupe",
     "Re-running the daily check could spam recipients. Per-day uniqueness check on the notifications table prevents duplicates."),
    ("Permission granularity",
     "Initial admin/user split was too coarse. Evolved into a per-module view/edit boolean grid matching real team structures."),
    ("Cross-entity expiry",
     "Subscriptions and licenses use different status fields. Refactored into two parallel passes sharing a recipient resolver."),
]
y = Inches(2.5)
card_h = Inches(1.0)
for i, (head, body) in enumerate(challenges):
    soft_card(s, Inches(0.75), y, SW - Inches(1.5), card_h)
    # left accent stripe
    add_gradient_rect(s, Inches(0.75), y, Inches(0.06), card_h,
                      (0x25, 0x63, 0xEB), (0x4F, 0x46, 0xE5),
                      angle_deg=90)
    text(s, Inches(1.0), y + Inches(0.2), Inches(0.8), Inches(0.4),
         f"0{i+1}", size=20, bold=True, color=ACCENT, letter_spacing=200)
    text(s, Inches(1.75), y + Inches(0.18), Inches(11), Inches(0.4),
         head, size=14, bold=True, color=INK)
    text(s, Inches(1.75), y + Inches(0.5), SW - Inches(2.5),
         Inches(0.5), body, size=10.5, color=GRAY_700)
    y += card_h + Inches(0.13)


# ====================================================================
# 11 — Roadmap (timeline)
# ====================================================================
s = add_slide()
page_chrome(s, 11, "Future improvements", "Roadmap")

quarters = [
    ("Q1", "Mobile app",
     "Native iOS / Android for on-the-go asset check-in & QR scanning."),
    ("Q2", "Integrations",
     "Slack, Microsoft Teams, and Jira Service Desk webhooks."),
    ("Q3", "AI insights",
     "Forecast costs and recommend consolidation opportunities."),
    ("Q4", "Multi-tenant SaaS",
     "Org isolation, SSO/SAML, per-tenant billing."),
]
n = len(quarters)
col_w = Inches(2.85)
gap = Inches(0.2)
total_w = col_w * n + gap * (n - 1)
sx = (SW - total_w) / 2
y = Inches(2.9)

# baseline gradient line
line_y = y + Inches(0.7)
add_gradient_rect(s, sx + Inches(0.2), line_y - Inches(0.02),
                  total_w - Inches(0.4), Inches(0.04),
                  (0x60, 0xA5, 0xFA), (0x4F, 0x46, 0xE5),
                  angle_deg=0)

for i, (q, head, body) in enumerate(quarters):
    x = sx + i * (col_w + gap)
    # quarter badge above the line
    add_rect(s, x + col_w / 2 - Inches(0.45),
             y, Inches(0.9), Inches(0.4),
             ACCENT, corner=0.4)
    text(s, x + col_w / 2 - Inches(0.45), y,
         Inches(0.9), Inches(0.4),
         q, size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         letter_spacing=150)
    # dot on the timeline
    add_rect(s, x + col_w / 2 - Inches(0.1),
             line_y - Inches(0.1),
             Inches(0.2), Inches(0.2),
             ACCENT, corner=0.5)
    add_rect(s, x + col_w / 2 - Inches(0.05),
             line_y - Inches(0.05),
             Inches(0.1), Inches(0.1),
             WHITE, corner=0.5)
    # card below
    cy = line_y + Inches(0.3)
    soft_card(s, x, cy, col_w, Inches(2.3))
    text(s, x + Inches(0.25), cy + Inches(0.35), col_w - Inches(0.5),
         Inches(0.4), head, size=15, bold=True, color=INK,
         align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.25), cy + Inches(0.9), col_w - Inches(0.5),
         Inches(1.3), body, size=10.5, color=GRAY_700,
         align=PP_ALIGN.CENTER)


# ====================================================================
# 12 — Team Members
# ====================================================================
s = add_slide()
page_chrome(s, 12, "Team members", "The Team")

members = [
    ("PL", "[Member 1]", "Project Lead / Backend"),
    ("FE", "[Member 2]", "Frontend / UI Engineer"),
    ("FS", "[Member 3]", "Full-Stack Developer"),
    ("DB", "[Member 4]", "Database / DevOps"),
]
card_w = Inches(2.95)
card_h = Inches(3.2)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
sx = (SW - total_w) / 2
y = Inches(2.7)

for i, (mono, name, role) in enumerate(members):
    x = sx + i * (card_w + gap)
    soft_card(s, x, y, card_w, card_h)
    # avatar block with subtle gradient
    av = Inches(1.2)
    avx = x + (card_w - av) / 2
    avy = y + Inches(0.45)
    add_gradient_rect(s, avx, avy, av, av,
                      (0x60, 0xA5, 0xFA), (0x4F, 0x46, 0xE5),
                      angle_deg=45, corner=0.18)
    text(s, avx, avy, av, av, mono, size=28, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x, y + Inches(1.95), card_w, Inches(0.4),
         name, size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
    text(s, x, y + Inches(2.35), card_w, Inches(0.4),
         role, size=10.5, color=GRAY_700, align=PP_ALIGN.CENTER)
    # accent rule
    add_line(s, x + (card_w - Inches(0.6)) / 2, y + card_h - Inches(0.4),
             x + (card_w + Inches(0.6)) / 2, y + card_h - Inches(0.4),
             rgb=ACCENT, width=1.5)


# ====================================================================
# 13 — Conclusion
# ====================================================================
s = add_slide()
page_chrome(s, 13, "Summary", "Conclusion")

text(s, Inches(0.75), Inches(2.4), Inches(11.5), Inches(1.6),
     "From spreadsheets to a single\nsource of truth.",
     size=32, bold=True, color=INK)
text(s, Inches(0.75), Inches(4.1), Inches(11.5), Inches(0.5),
     "ITAMS gives IT teams the visibility and automation they need.",
     size=14, color=GRAY_700)

takeaways = [
    ("◆", "Ready to deploy",     "Runs on a standard Laravel + LAMP / XAMPP stack."),
    ("◉", "Designed for users",  "Modern UI with dark mode and accessibility built in."),
    ("◇", "Built to extend",     "Modular architecture invites new asset types & integrations."),
]
col_w = Inches(4.0)
gap = Inches(0.25)
total_w = col_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(5.2)
card_h = Inches(1.7)

for i, (g, head, body) in enumerate(takeaways):
    x = sx + i * (col_w + gap)
    soft_card(s, x, y, col_w, card_h)
    icon_chip(s, x + Inches(0.3), y + Inches(0.3), g, size_in=0.42)
    text(s, x + Inches(0.95), y + Inches(0.3), col_w - Inches(1.1),
         Inches(0.4), head, size=13, bold=True, color=INK)
    text(s, x + Inches(0.95), y + Inches(0.7), col_w - Inches(1.1),
         Inches(0.9), body, size=10.5, color=GRAY_700)


# ====================================================================
# 14 — Thank You
# ====================================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, INK)
# subtle gradient overlay
add_gradient_rect(s, 0, 0, SW, Inches(4),
                  (0x1D, 0x4E, 0xD8), (0x0F, 0x17, 0x2A),
                  angle_deg=90)
add_gradient_rect(s, 0, Inches(4), SW, Inches(0.04),
                  (0x60, 0xA5, 0xFA), (0x4F, 0x46, 0xE5),
                  angle_deg=0)

text(s, Inches(0), Inches(2.2), SW, Inches(0.4),
     "HACKATHON 2026",
     size=10, bold=True, color=GRAY_400, letter_spacing=300,
     align=PP_ALIGN.CENTER)

text(s, Inches(0), Inches(2.8), SW, Inches(2.0),
     "Thank you",
     size=110, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# divider gradient
add_gradient_rect(s, (SW - Inches(1.2)) / 2, Inches(5.0),
                  Inches(1.2), Inches(0.06),
                  (0x60, 0xA5, 0xFA), (0x4F, 0x46, 0xE5),
                  angle_deg=0, corner=0.5)

text(s, Inches(0), Inches(5.2), SW, Inches(0.5),
     "Questions & discussion", size=18, color=GRAY_300,
     align=PP_ALIGN.CENTER)

text(s, Inches(0), SH - Inches(1.0), SW, Inches(0.4),
     "TEAM ITAMS  ·  IT ASSET MANAGEMENT SYSTEM",
     size=9, bold=True, color=GRAY_400, letter_spacing=300,
     align=PP_ALIGN.CENTER)
text(s, Inches(0), SH - Inches(0.65), SW, Inches(0.4),
     "[ contact email / repo link ]",
     size=10, color=GRAY_500, align=PP_ALIGN.CENTER)


# ---------- Save ----------
out = r"D:\xampp\htdocs\itams\presentation\ITAMS_Hackathon_Modern.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
