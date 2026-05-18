"""Build the ITAMS hackathon presentation as a .pptx file."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree

# ---------- Brand palette ----------
NAVY        = RGBColor(0x0F, 0x17, 0x2A)   # slate-900
SLATE_800   = RGBColor(0x1E, 0x29, 0x3B)
SLATE_700   = RGBColor(0x33, 0x41, 0x55)
SLATE_500   = RGBColor(0x64, 0x74, 0x8B)
SLATE_300   = RGBColor(0xCB, 0xD5, 0xE1)
SLATE_100   = RGBColor(0xF1, 0xF5, 0xF9)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
PRIMARY     = RGBColor(0x25, 0x63, 0xEB)   # royal blue
PRIMARY_DK  = RGBColor(0x1D, 0x4E, 0xD8)
INDIGO      = RGBColor(0x4F, 0x46, 0xE5)
CYAN        = RGBColor(0x06, 0xB6, 0xD4)
GREEN       = RGBColor(0x10, 0xB9, 0x81)
AMBER       = RGBColor(0xF5, 0x9E, 0x0B)
ROSE        = RGBColor(0xF4, 0x3F, 0x5E)

# Slide size (16:9 widescreen)
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ---------- Helpers ----------
def add_slide():
    return prs.slides.add_slide(BLANK)


def fill(shape, rgb):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    shape.line.fill.background()


def outline(shape, rgb, width=0.75):
    shape.line.color.rgb = rgb
    shape.line.width = Pt(width)


def add_rect(slide, x, y, w, h, rgb, no_line=True):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    fill(shp, rgb)
    if no_line:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_round_rect(slide, x, y, w, h, rgb, no_line=True, corner=0.08):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.adjustments[0] = corner
    fill(shp, rgb)
    if no_line:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, *,
             size=18, bold=False, color=NAVY,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             font="Calibri"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = font
    return tb


def add_runs(slide, x, y, w, h, runs, *, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, font="Calibri"):
    """runs = list of (text, size, bold, color)."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Pt(0)
    tf.margin_top = tf.margin_bottom = Pt(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    for text, size, bold, color in runs:
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = font
    return tb


def page_chrome(slide, page_no, title, eyebrow=None):
    """Light-theme content slide chrome: header strip, page number, title."""
    add_rect(slide, 0, 0, SW, SH, SLATE_100)
    # Top accent bar
    add_rect(slide, 0, 0, SW, Inches(0.18), PRIMARY)
    # Bottom slim strip
    add_rect(slide, 0, SH - Inches(0.4), SW, Inches(0.4), NAVY)
    add_text(slide, Inches(0.6), SH - Inches(0.36), Inches(8), Inches(0.32),
             "ITAMS  ·  IT Asset Management System",
             size=10, bold=True, color=SLATE_300, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, SW - Inches(1.4), SH - Inches(0.36), Inches(0.8), Inches(0.32),
             f"{page_no:02d}", size=10, bold=True, color=SLATE_300,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    # Eyebrow tag
    if eyebrow:
        tag = add_round_rect(slide, Inches(0.6), Inches(0.55),
                             Inches(2.4), Inches(0.32), PRIMARY, corner=0.45)
        add_text(slide, Inches(0.6), Inches(0.55), Inches(2.4), Inches(0.32),
                 eyebrow.upper(), size=10, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Title
    add_text(slide, Inches(0.55), Inches(1.0), Inches(12.0), Inches(0.9),
             title, size=34, bold=True, color=NAVY)
    # Underline accent
    add_rect(slide, Inches(0.6), Inches(1.85), Inches(0.6), Inches(0.06), PRIMARY)


def icon_chip(slide, x, y, size_in, glyph, bg=PRIMARY, fg=WHITE, corner=0.2):
    s = Inches(size_in)
    add_round_rect(slide, x, y, s, s, bg, corner=corner)
    add_text(slide, x, y, s, s, glyph, size=int(size_in * 22), bold=True,
             color=fg, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def feature_card(slide, x, y, w, h, glyph, title, body,
                 chip_color=PRIMARY):
    add_round_rect(slide, x, y, w, h, WHITE, corner=0.06)
    # subtle hairline border via overlay
    border = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    border.adjustments[0] = 0.06
    border.fill.background()
    border.line.color.rgb = SLATE_300
    border.line.width = Pt(0.5)
    border.shadow.inherit = False

    icon_chip(slide, x + Inches(0.3), y + Inches(0.3), 0.6, glyph,
              bg=chip_color)
    add_text(slide, x + Inches(0.3), y + Inches(1.05),
             w - Inches(0.6), Inches(0.4),
             title, size=14, bold=True, color=NAVY)
    add_text(slide, x + Inches(0.3), y + Inches(1.5),
             w - Inches(0.6), h - Inches(1.7),
             body, size=10.5, color=SLATE_700)


def add_arrow(slide, x1, y1, x2, y2, rgb=SLATE_500, width=1.5):
    line = slide.shapes.add_connector(1, x1, y1, x2, y2)
    line.line.color.rgb = rgb
    line.line.width = Pt(width)
    # add arrow head
    lnEl = line.line._get_or_add_ln()
    tail = etree.SubElement(lnEl, qn("a:tailEnd"))
    tail.set("type", "triangle")
    tail.set("w", "med")
    tail.set("h", "med")


# ====================================================================
# SLIDE 1 — Title
# ====================================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, NAVY)

# Decorative gradient blob (faked with rounded rects)
add_round_rect(s, Inches(-2), Inches(-2), Inches(7), Inches(7), PRIMARY, corner=0.5)
blob2 = add_round_rect(s, Inches(8), Inches(3), Inches(8), Inches(8), INDIGO, corner=0.5)
# Slight transparency-ish overlay via dark layer
overlay = add_rect(s, 0, 0, SW, SH, NAVY)
overlay.fill.transparency = 0  # leave opaque; the blobs read through padding

# Re-add navy overlay only for top-left corner area? Skip — blobs are bold accents.
# Actually let's redo: dark base, then accent stripes only.
# Clear and start clean:
s.shapes._spTree.remove(blob2._element)
# (Re-place blobs partially off-canvas for a soft accent edge)
add_round_rect(s, Inches(9.5), Inches(-2.5), Inches(7), Inches(7), INDIGO, corner=0.5)
add_round_rect(s, Inches(-2.5), Inches(4.5), Inches(6), Inches(6), PRIMARY, corner=0.5)

# Logo mark
add_round_rect(s, Inches(0.7), Inches(0.7), Inches(0.85), Inches(0.85),
               WHITE, corner=0.25)
add_text(s, Inches(0.7), Inches(0.7), Inches(0.85), Inches(0.85),
         "IT", size=20, bold=True, color=PRIMARY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(1.7), Inches(0.7), Inches(4), Inches(0.4),
         "ITAMS", size=16, bold=True, color=WHITE)
add_text(s, Inches(1.7), Inches(1.1), Inches(5), Inches(0.4),
         "IT Asset Management System", size=10, color=SLATE_300)

# Eyebrow
add_round_rect(s, Inches(0.7), Inches(2.2), Inches(2.4), Inches(0.35),
               PRIMARY, corner=0.5)
add_text(s, Inches(0.7), Inches(2.2), Inches(2.4), Inches(0.35),
         "HACKATHON 2026", size=10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Title
add_text(s, Inches(0.7), Inches(2.8), Inches(11), Inches(1.5),
         "ITAMS", size=88, bold=True, color=WHITE)
add_text(s, Inches(0.7), Inches(4.2), Inches(11), Inches(1.0),
         "An Enterprise IT Asset Management Platform",
         size=28, color=SLATE_300)

# Divider
add_rect(s, Inches(0.7), Inches(5.4), Inches(1.0), Inches(0.05), PRIMARY)

# Team
add_text(s, Inches(0.7), Inches(5.55), Inches(8), Inches(0.4),
         "Presented by", size=11, color=SLATE_500)
add_text(s, Inches(0.7), Inches(5.85), Inches(10), Inches(0.6),
         "Team ITAMS", size=24, bold=True, color=WHITE)

# Footer date
add_text(s, Inches(0.7), Inches(6.8), Inches(8), Inches(0.4),
         "Hackathon 2026  ·  IT Infrastructure Track",
         size=11, color=SLATE_500)

# ====================================================================
# SLIDE 2 — Team
# ====================================================================
s = add_slide()
page_chrome(s, 2, "Meet the Team", eyebrow="Project Members")

members = [
    ("[Member 1]", "Project Lead / Backend",   "PL"),
    ("[Member 2]", "Frontend / UI Engineer",   "FE"),
    ("[Member 3]", "Full-Stack Developer",     "FS"),
    ("[Member 4]", "Database / DevOps",        "DB"),
]
chip_colors = [PRIMARY, INDIGO, CYAN, GREEN]

card_w = Inches(2.85)
card_h = Inches(3.6)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
start_x = (SW - total_w) / 2
y = Inches(2.5)

for i, (name, role, initials) in enumerate(members):
    x = start_x + (card_w + gap) * i
    add_round_rect(s, x, y, card_w, card_h, WHITE, corner=0.07)
    # avatar
    avatar_size = Inches(1.4)
    ax = x + (card_w - avatar_size) / 2
    add_round_rect(s, ax, y + Inches(0.45), avatar_size, avatar_size,
                   chip_colors[i], corner=0.5)
    add_text(s, ax, y + Inches(0.45), avatar_size, avatar_size,
             initials, size=32, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x, y + Inches(2.15), card_w, Inches(0.5),
             name, size=16, bold=True, color=NAVY,
             align=PP_ALIGN.CENTER)
    add_text(s, x, y + Inches(2.65), card_w, Inches(0.5),
             role, size=11, color=SLATE_700,
             align=PP_ALIGN.CENTER)
    # divider
    add_rect(s, x + (card_w - Inches(0.5)) / 2,
             y + Inches(3.25), Inches(0.5), Inches(0.04), chip_colors[i])

add_text(s, Inches(0.6), Inches(6.45), Inches(12.0), Inches(0.4),
         "A cross-functional team building enterprise-grade IT operations tooling.",
         size=12, color=SLATE_500, align=PP_ALIGN.CENTER)

# ====================================================================
# SLIDE 3 — Problem statement
# ====================================================================
s = add_slide()
page_chrome(s, 3, "The Problem", eyebrow="Problem Statement")

add_text(s, Inches(0.6), Inches(2.2), Inches(12.0), Inches(0.8),
         "IT teams lose visibility and money managing assets in spreadsheets.",
         size=22, bold=True, color=NAVY)

pains = [
    ("⚠", "Manual tracking", "Devices, licenses & subscriptions live in disconnected Excel files.", ROSE),
    ("⏰", "Missed renewals", "Contracts auto-renew or expire without warning, causing downtime.", AMBER),
    ("🔍", "No audit trail",  "No record of who changed what, when — failing compliance reviews.", INDIGO),
    ("💸", "Hidden cost",     "Duplicate purchases & unused subscriptions inflate IT spend by 20–30%.", PRIMARY),
]

card_w = Inches(2.9)
card_h = Inches(2.6)
gap = Inches(0.18)
total_w = card_w * 4 + gap * 3
start_x = (SW - total_w) / 2
y = Inches(3.4)

for i, (glyph, title, desc, color) in enumerate(pains):
    x = start_x + (card_w + gap) * i
    add_round_rect(s, x, y, card_w, card_h, WHITE, corner=0.07)
    icon_chip(s, x + Inches(0.3), y + Inches(0.3), 0.55, glyph, bg=color)
    add_text(s, x + Inches(0.3), y + Inches(1.0), card_w - Inches(0.6),
             Inches(0.4), title, size=14, bold=True, color=NAVY)
    add_text(s, x + Inches(0.3), y + Inches(1.45), card_w - Inches(0.6),
             card_h - Inches(1.6), desc, size=10.5, color=SLATE_700)

# ====================================================================
# SLIDE 4 — Proposed solution
# ====================================================================
s = add_slide()
page_chrome(s, 4, "Our Solution", eyebrow="Proposed Solution")

add_text(s, Inches(0.6), Inches(2.2), Inches(12), Inches(1.0),
         "A single web platform for the entire IT asset lifecycle.",
         size=22, bold=True, color=NAVY)
add_text(s, Inches(0.6), Inches(3.0), Inches(12), Inches(0.8),
         "ITAMS unifies hardware, devices, software licenses and subscriptions —\n"
         "with automated renewal reminders, role-based access, and a full audit log.",
         size=15, color=SLATE_700)

# Three solution pillars
pillars = [
    ("🗂", "Unify", "One source of truth for every IT asset.", PRIMARY),
    ("🔔", "Automate", "Email + in-app reminders before things expire.", INDIGO),
    ("🛡", "Govern", "Role-based access and full activity history.", GREEN),
]
card_w = Inches(3.8)
card_h = Inches(1.8)
gap = Inches(0.25)
total_w = card_w * 3 + gap * 2
start_x = (SW - total_w) / 2
y = Inches(4.6)

for i, (g, t, d, c) in enumerate(pillars):
    x = start_x + (card_w + gap) * i
    add_round_rect(s, x, y, card_w, card_h, WHITE, corner=0.08)
    icon_chip(s, x + Inches(0.3), y + Inches(0.3), 0.55, g, bg=c)
    add_text(s, x + Inches(1.05), y + Inches(0.35), card_w - Inches(1.2),
             Inches(0.4), t, size=15, bold=True, color=NAVY)
    add_text(s, x + Inches(1.05), y + Inches(0.85), card_w - Inches(1.2),
             card_h - Inches(1.0), d, size=11, color=SLATE_700)

# ====================================================================
# SLIDE 5 — Architecture
# ====================================================================
s = add_slide()
page_chrome(s, 5, "System Architecture", eyebrow="Overview")

# 3-tier architecture diagram
tier_y = Inches(2.5)
tier_h = Inches(1.1)
tier_w = Inches(3.6)
gap = Inches(0.3)
total_w = tier_w * 3 + gap * 2
start_x = (SW - total_w) / 2

tiers = [
    ("Presentation Tier", "Browser  ·  Bootstrap 5  ·  Blade Views", PRIMARY),
    ("Application Tier",  "Laravel 11  ·  Controllers  ·  Mailables", INDIGO),
    ("Data Tier",         "MySQL 8  ·  Eloquent ORM  ·  Migrations",  GREEN),
]

for i, (title, sub, color) in enumerate(tiers):
    x = start_x + (tier_w + gap) * i
    add_round_rect(s, x, tier_y, tier_w, tier_h, WHITE, corner=0.1)
    # color accent stripe
    add_rect(s, x, tier_y, Inches(0.12), tier_h, color)
    add_text(s, x + Inches(0.3), tier_y + Inches(0.15), tier_w - Inches(0.4),
             Inches(0.4), title, size=14, bold=True, color=NAVY)
    add_text(s, x + Inches(0.3), tier_y + Inches(0.6), tier_w - Inches(0.4),
             Inches(0.4), sub, size=10.5, color=SLATE_700)
    # arrow to next tier
    if i < 2:
        add_arrow(s, x + tier_w, tier_y + tier_h / 2,
                  x + tier_w + gap, tier_y + tier_h / 2, rgb=SLATE_500)

# Secondary layer: integrations & infra
intg_y = Inches(4.1)
intg_h = Inches(0.95)

add_round_rect(s, start_x, intg_y, total_w, intg_h, NAVY, corner=0.08)
add_text(s, start_x + Inches(0.4), intg_y + Inches(0.1),
         total_w - Inches(0.8), Inches(0.4),
         "Background Services", size=12, bold=True, color=WHITE)

services = [
    ("⏱  Scheduler",   "app:check-expirations runs daily"),
    ("📧  SMTP",       "Renewal reminder emails"),
    ("📝  Activity Log","Per-user audit trail"),
]
svc_w = (total_w - Inches(1.2)) / 3
for i, (label, desc) in enumerate(services):
    sx = start_x + Inches(0.4) + svc_w * i + (Inches(0.2) if i > 0 else 0) * i
    add_text(s, sx, intg_y + Inches(0.45), svc_w, Inches(0.25),
             label, size=10.5, bold=True, color=PRIMARY)
    add_text(s, sx, intg_y + Inches(0.68), svc_w, Inches(0.3),
             desc, size=9.5, color=SLATE_300)

# Modules row
mod_y = Inches(5.3)
mod_h = Inches(1.4)
add_text(s, Inches(0.6), mod_y - Inches(0.35), Inches(12), Inches(0.3),
         "Functional Modules", size=11, bold=True, color=SLATE_500)

modules = [
    ("PC Master", PRIMARY), ("Device Master", INDIGO),
    ("Subscriptions", CYAN), ("Licenses & Contracts", GREEN),
    ("Notifications", AMBER), ("User & Role Mgmt", ROSE),
]
m_w = (SW - Inches(1.2)) / 6 - Inches(0.1)
for i, (name, color) in enumerate(modules):
    mx = Inches(0.6) + (m_w + Inches(0.1)) * i
    add_round_rect(s, mx, mod_y, m_w, mod_h, WHITE, corner=0.1)
    add_rect(s, mx, mod_y, m_w, Inches(0.08), color)
    add_text(s, mx, mod_y + Inches(0.55), m_w, Inches(0.4),
             name, size=11, bold=True, color=NAVY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ====================================================================
# SLIDE 6 — Key Features (1)
# ====================================================================
s = add_slide()
page_chrome(s, 6, "Key Features — Asset Modules", eyebrow="What It Does (1/2)")

features = [
    ("💻", "PC Master",
     "Track every workstation: assignee, OS, specs, purchase date, warranty.",
     PRIMARY),
    ("📡", "Device Master",
     "Network hardware inventory — routers, switches, servers, printers.",
     INDIGO),
    ("📅", "Subscriptions",
     "Recurring service subscriptions with renewal cycles and vendor data.",
     CYAN),
    ("📑", "Licenses & Contracts",
     "Software licenses and vendor contracts with expiry tracking.",
     GREEN),
]
card_w = Inches(2.95)
card_h = Inches(3.4)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
start_x = (SW - total_w) / 2
y = Inches(2.5)

for i, (g, t, d, c) in enumerate(features):
    x = start_x + (card_w + gap) * i
    feature_card(s, x, y, card_w, card_h, g, t, d, chip_color=c)

add_text(s, Inches(0.6), Inches(6.3), Inches(12), Inches(0.4),
         "Import / export to Excel · Bulk operations · Searchable, filterable, paginated.",
         size=12, bold=True, color=SLATE_500, align=PP_ALIGN.CENTER)

# ====================================================================
# SLIDE 7 — Key Features (2)
# ====================================================================
s = add_slide()
page_chrome(s, 7, "Key Features — Operations & Governance",
            eyebrow="What It Does (2/2)")

features = [
    ("🔔", "Smart Reminders",
     "Daily scheduler flags items expiring within a configurable window; emails recipients and posts in-app notifications.",
     AMBER),
    ("🛡", "Role-Based Access",
     "Admin / Manager / Viewer roles with per-module permissions, enforced by middleware.",
     ROSE),
    ("📊", "Audit Trail",
     "Every create / update / delete recorded with user, timestamp, and field-level diff.",
     INDIGO),
    ("⚙", "Configurable Mail",
     "Switch between .env SMTP and database-driven settings without redeploying.",
     PRIMARY),
]
card_w = Inches(2.95)
card_h = Inches(3.4)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
start_x = (SW - total_w) / 2
y = Inches(2.5)

for i, (g, t, d, c) in enumerate(features):
    x = start_x + (card_w + gap) * i
    feature_card(s, x, y, card_w, card_h, g, t, d, chip_color=c)

add_text(s, Inches(0.6), Inches(6.3), Inches(12), Inches(0.4),
         "Built-in test email previews the real renewal template with live data.",
         size=12, bold=True, color=SLATE_500, align=PP_ALIGN.CENTER)

# ====================================================================
# SLIDE 8 — UI/UX Highlights
# ====================================================================
s = add_slide()
page_chrome(s, 8, "UI / UX Highlights", eyebrow="Design Language")

# Left panel: design principles
add_text(s, Inches(0.6), Inches(2.3), Inches(6), Inches(0.4),
         "Modern enterprise dashboard, optimized for daily use.",
         size=16, bold=True, color=NAVY)

principles = [
    ("Liquid-glass cards",  "Backdrop-blurred surfaces over soft gradients."),
    ("Dark & light themes", "User-toggled and persisted in localStorage."),
    ("Modern typography",   "Inter font with tight letter-spacing for clarity."),
    ("Accessible by design","Keyboard-friendly, ARIA-labeled, focus rings on all inputs."),
    ("Responsive layouts",  "Sidebar collapses to icon rail on tablet; full mobile reflow."),
]
y = Inches(3.0)
for label, desc in principles:
    # bullet dot
    add_round_rect(s, Inches(0.6), y + Inches(0.12), Inches(0.12),
                   Inches(0.12), PRIMARY, corner=0.5)
    add_text(s, Inches(0.85), y, Inches(5.6), Inches(0.3),
             label, size=12, bold=True, color=NAVY)
    add_text(s, Inches(0.85), y + Inches(0.27), Inches(5.6), Inches(0.4),
             desc, size=10.5, color=SLATE_700)
    y += Inches(0.65)

# Right panel: stylized browser mockup
mock_x = Inches(7.2)
mock_y = Inches(2.3)
mock_w = Inches(5.6)
mock_h = Inches(3.8)

add_round_rect(s, mock_x, mock_y, mock_w, mock_h, WHITE, corner=0.04)
# browser bar
add_rect(s, mock_x, mock_y, mock_w, Inches(0.45), SLATE_100)
# traffic lights
for i, c in enumerate([ROSE, AMBER, GREEN]):
    add_round_rect(s, mock_x + Inches(0.15) + Inches(0.25) * i,
                   mock_y + Inches(0.15), Inches(0.15), Inches(0.15),
                   c, corner=0.5)
# url
add_round_rect(s, mock_x + Inches(1.1), mock_y + Inches(0.1),
               Inches(4.3), Inches(0.25), WHITE, corner=0.3)
add_text(s, mock_x + Inches(1.2), mock_y + Inches(0.12),
         Inches(4.0), Inches(0.22),
         "itams.local / dashboard", size=8.5, color=SLATE_500,
         anchor=MSO_ANCHOR.MIDDLE)

# sidebar
add_rect(s, mock_x, mock_y + Inches(0.45), Inches(1.1),
         mock_h - Inches(0.45), NAVY)
add_round_rect(s, mock_x + Inches(0.18), mock_y + Inches(0.6),
               Inches(0.74), Inches(0.32), PRIMARY, corner=0.2)
add_text(s, mock_x + Inches(0.18), mock_y + Inches(0.6),
         Inches(0.74), Inches(0.32), "ITAMS", size=9, bold=True,
         color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# nav items
nav = ["Dashboard", "PC Master", "Devices", "Subs", "Licenses", "Logs"]
for i, n in enumerate(nav):
    ny = mock_y + Inches(1.1) + Inches(0.32) * i
    if i == 0:
        add_round_rect(s, mock_x + Inches(0.08), ny - Inches(0.03),
                       Inches(0.94), Inches(0.28), PRIMARY, corner=0.25)
        add_text(s, mock_x + Inches(0.15), ny, Inches(0.85), Inches(0.22),
                 n, size=8, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)
    else:
        add_text(s, mock_x + Inches(0.15), ny, Inches(0.85), Inches(0.22),
                 n, size=8, color=SLATE_300,
                 anchor=MSO_ANCHOR.MIDDLE)

# main area: KPI tiles
kpi_y = mock_y + Inches(0.7)
kpi_x = mock_x + Inches(1.3)
kpi_w = (mock_w - Inches(1.5)) / 4 - Inches(0.07)
kpi_h = Inches(0.85)
kpi_colors = [PRIMARY, INDIGO, GREEN, AMBER]
kpi_data = [("128", "PCs"), ("64", "Devices"),
            ("32", "Active"), ("5", "Expiring")]
for i, ((v, l), c) in enumerate(zip(kpi_data, kpi_colors)):
    x = kpi_x + (kpi_w + Inches(0.09)) * i
    add_round_rect(s, x, kpi_y, kpi_w, kpi_h, SLATE_100, corner=0.1)
    add_rect(s, x, kpi_y, Inches(0.06), kpi_h, c)
    add_text(s, x + Inches(0.15), kpi_y + Inches(0.05),
             kpi_w - Inches(0.2), Inches(0.3),
             v, size=14, bold=True, color=NAVY)
    add_text(s, x + Inches(0.15), kpi_y + Inches(0.45),
             kpi_w - Inches(0.2), Inches(0.3),
             l, size=8, color=SLATE_500)

# chart placeholder
chart_y = kpi_y + kpi_h + Inches(0.15)
chart_h = mock_y + mock_h - chart_y - Inches(0.2)
add_round_rect(s, kpi_x, chart_y,
               mock_w - Inches(1.5), chart_h, SLATE_100, corner=0.08)
# pretend bars
bar_w = Inches(0.18)
bar_gap = Inches(0.12)
n_bars = 8
heights = [0.55, 0.75, 0.42, 0.88, 0.6, 0.95, 0.7, 0.5]
bx = kpi_x + Inches(0.3)
by_base = chart_y + chart_h - Inches(0.3)
for i in range(n_bars):
    h = Inches(heights[i] * 1.5)
    add_rect(s, bx + (bar_w + bar_gap) * i, by_base - h, bar_w, h, PRIMARY)

# ====================================================================
# SLIDE 9 — Tech Stack
# ====================================================================
s = add_slide()
page_chrome(s, 9, "Technology Stack", eyebrow="Built With")

groups = [
    ("Frontend", [
        ("Bootstrap 5.3",      "UI components & responsive grid"),
        ("Bootstrap Icons",    "Consistent iconography"),
        ("Inter font",         "Modern enterprise typeface"),
        ("Vanilla JS",         "Lightweight interactivity"),
    ], PRIMARY),
    ("Backend", [
        ("Laravel 11",         "MVC framework, Eloquent ORM"),
        ("PHP 8.2+",           "Typed properties, modern syntax"),
        ("Blade",              "Server-rendered templating"),
        ("Maatwebsite/Excel",  "Import / export to .xlsx"),
    ], INDIGO),
    ("Data & Infra", [
        ("MySQL 8 / MariaDB",  "Relational storage"),
        ("Symfony Mailer",     "SMTP delivery via Laravel"),
        ("XAMPP / Apache",     "Local & on-premise hosting"),
        ("Cron scheduler",     "Daily expiry checks"),
    ], GREEN),
]
col_w = Inches(3.95)
col_h = Inches(4.0)
gap = Inches(0.25)
total_w = col_w * 3 + gap * 2
start_x = (SW - total_w) / 2
y = Inches(2.4)

for i, (heading, items, color) in enumerate(groups):
    x = start_x + (col_w + gap) * i
    add_round_rect(s, x, y, col_w, col_h, WHITE, corner=0.07)
    # heading bar
    add_rect(s, x, y, col_w, Inches(0.5), color)
    add_text(s, x + Inches(0.25), y, col_w - Inches(0.5), Inches(0.5),
             heading.upper(), size=11, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    # items
    iy = y + Inches(0.75)
    for label, desc in items:
        # bullet
        add_round_rect(s, x + Inches(0.3), iy + Inches(0.1),
                       Inches(0.1), Inches(0.1), color, corner=0.5)
        add_text(s, x + Inches(0.5), iy, col_w - Inches(0.7),
                 Inches(0.3), label, size=11, bold=True, color=NAVY)
        add_text(s, x + Inches(0.5), iy + Inches(0.25),
                 col_w - Inches(0.7), Inches(0.35),
                 desc, size=9.5, color=SLATE_700)
        iy += Inches(0.78)

# ====================================================================
# SLIDE 10 — Workflow
# ====================================================================
s = add_slide()
page_chrome(s, 10, "How It Works", eyebrow="Workflow")

steps = [
    ("01", "Onboard",
     "Admin adds assets manually or imports an .xlsx file.",
     PRIMARY),
    ("02", "Track",
     "Modules show real-time status across the asset estate.",
     INDIGO),
    ("03", "Detect",
     "Daily scheduler scans for items inside the renewal window.",
     CYAN),
    ("04", "Notify",
     "Email + in-app notifications fire to configured recipients.",
     AMBER),
    ("05", "Renew",
     "Owners renew or terminate; status & history auto-update.",
     GREEN),
]
n = len(steps)
card_w = Inches(2.32)
card_h = Inches(2.9)
gap = Inches(0.12)
total_w = card_w * n + gap * (n - 1)
start_x = (SW - total_w) / 2
y = Inches(2.7)

for i, (num, title, desc, color) in enumerate(steps):
    x = start_x + (card_w + gap) * i
    add_round_rect(s, x, y, card_w, card_h, WHITE, corner=0.08)
    # big step number
    add_text(s, x + Inches(0.25), y + Inches(0.2), card_w - Inches(0.5),
             Inches(0.7), num, size=36, bold=True, color=color)
    # divider
    add_rect(s, x + Inches(0.25), y + Inches(1.0), Inches(0.5),
             Inches(0.04), color)
    add_text(s, x + Inches(0.25), y + Inches(1.15), card_w - Inches(0.5),
             Inches(0.4), title, size=15, bold=True, color=NAVY)
    add_text(s, x + Inches(0.25), y + Inches(1.6), card_w - Inches(0.5),
             card_h - Inches(1.8), desc, size=10.5, color=SLATE_700)

    # arrow to next
    if i < n - 1:
        cx = x + card_w
        cy = y + card_h / 2
        add_arrow(s, cx, cy, cx + gap, cy, rgb=SLATE_500, width=1.25)

add_text(s, Inches(0.6), Inches(6.1), Inches(12), Inches(0.4),
         "Per-day dedupe ensures the same item is not re-notified more than once daily.",
         size=11, color=SLATE_500, align=PP_ALIGN.CENTER)

# ====================================================================
# SLIDE 11 — Benefits & Impact
# ====================================================================
s = add_slide()
page_chrome(s, 11, "Benefits & Impact", eyebrow="Outcomes")

# Big metric tiles
metrics = [
    ("30%",  "lower IT spend",        "Eliminates duplicate purchases & shelfware",  PRIMARY),
    ("0",    "missed renewals",       "Automatic reminders for every asset",         GREEN),
    ("100%", "audit coverage",        "Every change is logged with user & timestamp",INDIGO),
    ("5×",   "faster onboarding",     "Bulk Excel import vs. manual entry",          AMBER),
]
card_w = Inches(2.95)
card_h = Inches(3.5)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
start_x = (SW - total_w) / 2
y = Inches(2.4)

for i, (value, label, desc, color) in enumerate(metrics):
    x = start_x + (card_w + gap) * i
    add_round_rect(s, x, y, card_w, card_h, WHITE, corner=0.08)
    add_rect(s, x, y, card_w, Inches(0.12), color)
    add_text(s, x + Inches(0.25), y + Inches(0.5), card_w - Inches(0.5),
             Inches(1.2), value, size=46, bold=True, color=color,
             align=PP_ALIGN.LEFT)
    add_text(s, x + Inches(0.25), y + Inches(1.75), card_w - Inches(0.5),
             Inches(0.5), label, size=14, bold=True, color=NAVY)
    add_text(s, x + Inches(0.25), y + Inches(2.25), card_w - Inches(0.5),
             card_h - Inches(2.4), desc, size=11, color=SLATE_700)

add_text(s, Inches(0.6), Inches(6.2), Inches(12), Inches(0.5),
         "Visibility · Accountability · Predictable budgets · Better compliance posture",
         size=12, bold=True, color=SLATE_500, align=PP_ALIGN.CENTER)

# ====================================================================
# SLIDE 12 — Future Roadmap
# ====================================================================
s = add_slide()
page_chrome(s, 12, "Future Roadmap", eyebrow="What's Next")

phases = [
    ("Q1", "Mobile App",
     "Native iOS / Android for asset check-in & QR scanning.", PRIMARY),
    ("Q2", "Integrations",
     "Slack, Microsoft Teams, Jira Service Desk webhooks.",  INDIGO),
    ("Q3", "AI Insights",
     "Predict cost trends; recommend consolidation opportunities.", CYAN),
    ("Q4", "Multi-tenant SaaS",
     "Org isolation, SSO/SAML, per-tenant billing.", GREEN),
]
card_w = Inches(2.95)
card_h = Inches(3.4)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
start_x = (SW - total_w) / 2
y = Inches(2.5)

for i, (q, t, d, c) in enumerate(phases):
    x = start_x + (card_w + gap) * i
    add_round_rect(s, x, y, card_w, card_h, WHITE, corner=0.08)
    # quarter pill
    add_round_rect(s, x + Inches(0.3), y + Inches(0.3),
                   Inches(0.85), Inches(0.4), c, corner=0.45)
    add_text(s, x + Inches(0.3), y + Inches(0.3), Inches(0.85),
             Inches(0.4), q, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.3), y + Inches(0.95), card_w - Inches(0.6),
             Inches(0.5), t, size=15, bold=True, color=NAVY)
    add_text(s, x + Inches(0.3), y + Inches(1.55), card_w - Inches(0.6),
             card_h - Inches(1.7), d, size=11, color=SLATE_700)
    add_rect(s, x + Inches(0.3), y + card_h - Inches(0.45),
             Inches(0.5), Inches(0.04), c)

# ====================================================================
# SLIDE 13 — Conclusion
# ====================================================================
s = add_slide()
page_chrome(s, 13, "Conclusion", eyebrow="Why ITAMS")

# Big closing statement
add_text(s, Inches(0.6), Inches(2.3), Inches(12), Inches(1.2),
         "From spreadsheets to a single source of truth —",
         size=28, bold=True, color=NAVY)
add_text(s, Inches(0.6), Inches(3.1), Inches(12), Inches(0.8),
         "ITAMS gives IT teams the visibility and automation they need.",
         size=22, color=SLATE_700)

# Three takeaway pills
takeaways = [
    ("✓", "Ready to deploy",      "Runs on standard Laravel + LAMP / XAMPP stack."),
    ("✓", "Designed for users",   "Modern UI with dark mode and accessibility built in."),
    ("✓", "Built to extend",      "Modular architecture invites new asset types & integrations."),
]
y = Inches(4.6)
card_w = Inches(3.95)
card_h = Inches(1.5)
gap = Inches(0.2)
total_w = card_w * 3 + gap * 2
start_x = (SW - total_w) / 2

for i, (g, t, d) in enumerate(takeaways):
    x = start_x + (card_w + gap) * i
    add_round_rect(s, x, y, card_w, card_h, WHITE, corner=0.1)
    icon_chip(s, x + Inches(0.3), y + Inches(0.4), 0.55, g, bg=GREEN)
    add_text(s, x + Inches(1.05), y + Inches(0.35),
             card_w - Inches(1.2), Inches(0.4),
             t, size=13, bold=True, color=NAVY)
    add_text(s, x + Inches(1.05), y + Inches(0.78),
             card_w - Inches(1.2), card_h - Inches(0.9),
             d, size=10.5, color=SLATE_700)

# ====================================================================
# SLIDE 14 — Thank You
# ====================================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, NAVY)
# soft blobs
add_round_rect(s, Inches(-3), Inches(-3), Inches(8), Inches(8),
               PRIMARY, corner=0.5)
add_round_rect(s, Inches(9), Inches(4), Inches(8), Inches(8),
               INDIGO, corner=0.5)

# centered content
add_text(s, Inches(0), Inches(2.3), SW, Inches(0.6),
         "HACKATHON 2026", size=14, bold=True, color=SLATE_300,
         align=PP_ALIGN.CENTER)
add_text(s, Inches(0), Inches(2.9), SW, Inches(1.8),
         "Thank You", size=96, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)
# divider
add_rect(s, (SW - Inches(0.8)) / 2, Inches(4.7),
         Inches(0.8), Inches(0.06), PRIMARY)
add_text(s, Inches(0), Inches(4.9), SW, Inches(0.5),
         "Questions & Discussion", size=20, color=SLATE_300,
         align=PP_ALIGN.CENTER)

# team line
add_text(s, Inches(0), Inches(6.3), SW, Inches(0.4),
         "Team ITAMS  ·  IT Asset Management System",
         size=12, bold=True, color=SLATE_300, align=PP_ALIGN.CENTER)
add_text(s, Inches(0), Inches(6.7), SW, Inches(0.4),
         "[ contact email / repo link ]",
         size=11, color=SLATE_500, align=PP_ALIGN.CENTER)


# ---------- Save ----------
out = r"D:\xampp\htdocs\itams\presentation\ITAMS_Hackathon.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
