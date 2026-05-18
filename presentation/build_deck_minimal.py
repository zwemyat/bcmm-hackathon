"""Build the ITAMS hackathon presentation — minimal corporate-tech style."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree


# ---------- Strict monochrome palette + 1 accent ----------
INK        = RGBColor(0x0F, 0x17, 0x2A)   # near-black for primary text / dark surfaces
INK_SOFT   = RGBColor(0x1E, 0x29, 0x3B)
GRAY_700   = RGBColor(0x33, 0x41, 0x55)
GRAY_500   = RGBColor(0x64, 0x74, 0x8B)
GRAY_400   = RGBColor(0x94, 0xA3, 0xB8)
GRAY_300   = RGBColor(0xCB, 0xD5, 0xE1)
GRAY_200   = RGBColor(0xE2, 0xE8, 0xF0)
GRAY_100   = RGBColor(0xF1, 0xF5, 0xF9)
GRAY_50    = RGBColor(0xF8, 0xFA, 0xFC)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT     = RGBColor(0x1E, 0x40, 0xAF)   # deep blue — only accent color used

FONT = "Calibri"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ---------- Primitives ----------
def add_slide():
    return prs.slides.add_slide(BLANK)


def add_rect(slide, x, y, w, h, rgb, line_rgb=None, line_pt=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb
    if line_rgb is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_rgb
        shp.line.width = Pt(line_pt or 0.75)
    shp.shadow.inherit = False
    return shp


def add_outline_rect(slide, x, y, w, h, line_rgb=GRAY_300, line_pt=0.75):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.background()
    shp.line.color.rgb = line_rgb
    shp.line.width = Pt(line_pt)
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
         font=FONT, spacing=None, letter_spacing=None):
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
        if spacing:
            p.space_before = Pt(0)
            p.space_after = Pt(spacing)
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


def eyebrow(slide, x, y, label):
    """Small uppercase label, letter-spaced, gray."""
    text(slide, x, y, Inches(5), Inches(0.3), label.upper(),
         size=9, bold=True, color=GRAY_500, letter_spacing=200)


def page_chrome(slide, page_no, title_str, eyebrow_str):
    """Consistent chrome for every content slide."""
    # eyebrow + title block
    eyebrow(slide, Inches(0.7), Inches(0.55), eyebrow_str)
    text(slide, Inches(0.7), Inches(0.85), Inches(12), Inches(0.7),
         title_str, size=30, bold=True, color=INK)
    # Thin divider line under title
    add_line(slide, Inches(0.7), Inches(1.65),
             Inches(1.7), Inches(1.65), rgb=INK, width=1.5)
    # Page number top right
    text(slide, SW - Inches(1.3), Inches(0.55), Inches(0.7), Inches(0.3),
         f"{page_no:02d} / 14", size=9, bold=True, color=GRAY_400,
         align=PP_ALIGN.RIGHT, letter_spacing=150)
    # Wordmark bottom left
    text(slide, Inches(0.7), SH - Inches(0.5), Inches(6), Inches(0.3),
         "ITAMS", size=9, bold=True, color=GRAY_500, letter_spacing=200)
    text(slide, Inches(1.2), SH - Inches(0.5), Inches(6), Inches(0.3),
         "IT Asset Management System",
         size=9, color=GRAY_400)


# Reused "icon" — a clean square outline with a small inner mark, no emoji
def marker(slide, x, y, size_in=0.32, mark="▸"):
    s = Inches(size_in)
    text(slide, x, y, s, s, mark, size=int(size_in * 38),
         bold=True, color=ACCENT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def numeral(slide, x, y, n_str, size=44):
    """Big light numeral used for steps and stats."""
    text(slide, x, y, Inches(1.6), Inches(0.9), n_str,
         size=size, bold=True, color=GRAY_300, font=FONT)


# ====================================================================
# 01 — Title slide (dark)
# ====================================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, INK)

# Eyebrow + monogram
text(s, Inches(0.7), Inches(0.6), Inches(4), Inches(0.3),
     "HACKATHON 2026  ·  IT INFRASTRUCTURE TRACK",
     size=9, bold=True, color=GRAY_400, letter_spacing=300)

# Wordmark
text(s, Inches(0.7), Inches(3.0), Inches(10), Inches(1.5),
     "ITAMS", size=120, bold=True, color=WHITE)

# Thin divider
add_line(s, Inches(0.75), Inches(4.7),
         Inches(2.0), Inches(4.7), rgb=ACCENT, width=2.0)

# Subtitle
text(s, Inches(0.7), Inches(4.85), Inches(11), Inches(0.6),
     "An Enterprise IT Asset Management Platform",
     size=22, color=GRAY_300)

# Footer block
text(s, Inches(0.7), SH - Inches(1.1), Inches(4), Inches(0.25),
     "PRESENTED BY", size=8, bold=True, color=GRAY_500, letter_spacing=300)
text(s, Inches(0.7), SH - Inches(0.78), Inches(6), Inches(0.4),
     "Team ITAMS", size=18, bold=True, color=WHITE)

text(s, SW - Inches(4.7), SH - Inches(0.78), Inches(4), Inches(0.4),
     "May 2026", size=12, color=GRAY_400, align=PP_ALIGN.RIGHT)


# ====================================================================
# 02 — Problem Statement
# ====================================================================
s = add_slide()
page_chrome(s, 2, "The problem", "Problem Statement")

# Big lead statement
text(s, Inches(0.7), Inches(2.4), Inches(11), Inches(1.2),
     "IT teams lose visibility and money\nmanaging assets in spreadsheets.",
     size=28, bold=True, color=INK)

# 4 problem points in a simple 2x2 grid, no fills
points = [
    ("Manual tracking",
     "Devices, licenses & subscriptions scattered across disconnected Excel files."),
    ("Missed renewals",
     "Contracts expire or auto-renew without warning, causing downtime or waste."),
    ("No audit trail",
     "No record of who changed what, when — failing internal compliance reviews."),
    ("Hidden cost",
     "Duplicate purchases and unused subscriptions inflate IT spend by 20–30%."),
]
col_w = Inches(5.5)
row_h = Inches(1.4)
start_x = Inches(0.7)
start_y = Inches(4.5)
gap_x = Inches(0.4)
gap_y = Inches(0.3)

for i, (head, body) in enumerate(points):
    row = i // 2
    col = i % 2
    x = start_x + col * (col_w + gap_x)
    y = start_y + row * (row_h + gap_y)
    marker(s, x, y + Inches(0.05))
    text(s, x + Inches(0.4), y, col_w - Inches(0.5), Inches(0.4),
         head, size=14, bold=True, color=INK)
    text(s, x + Inches(0.4), y + Inches(0.42),
         col_w - Inches(0.5), Inches(0.9),
         body, size=11, color=GRAY_700)


# ====================================================================
# 03 — Proposed Solution
# ====================================================================
s = add_slide()
page_chrome(s, 3, "Our solution", "Proposed Solution")

text(s, Inches(0.7), Inches(2.4), Inches(11), Inches(1.0),
     "A single platform for the entire\nIT asset lifecycle.",
     size=28, bold=True, color=INK)

text(s, Inches(0.7), Inches(4.0), Inches(11), Inches(0.5),
     "ITAMS unifies hardware, devices, licenses and subscriptions — with automated "
     "renewal reminders, role-based access, and a full audit log.",
     size=13, color=GRAY_700)

# 3 pillars listed cleanly
pillars = [
    ("Unify",   "One source of truth for every IT asset."),
    ("Automate","Email + in-app reminders before things expire."),
    ("Govern",  "Role-based access with full activity history."),
]
col_w = Inches(3.85)
gap = Inches(0.25)
total_w = col_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(5.0)

for i, (head, body) in enumerate(pillars):
    x = sx + i * (col_w + gap)
    # number prefix
    text(s, x, y, Inches(1), Inches(0.4),
         f"0{i+1}", size=11, bold=True, color=ACCENT, letter_spacing=200)
    text(s, x, y + Inches(0.45), col_w, Inches(0.5),
         head, size=18, bold=True, color=INK)
    text(s, x, y + Inches(1.0), col_w, Inches(0.8),
         body, size=11, color=GRAY_700)
    # divider under
    add_line(s, x, y + Inches(0.4),
             x + Inches(0.4), y + Inches(0.4), rgb=ACCENT, width=1.5)


# ====================================================================
# 04 — System Overview (architecture diagram)
# ====================================================================
s = add_slide()
page_chrome(s, 4, "System overview", "Architecture")

# 3-tier horizontal: outlined boxes connected by thin arrows.
tier_y = Inches(2.7)
tier_h = Inches(1.4)
tier_w = Inches(3.4)
gap = Inches(0.5)
total_w = tier_w * 3 + gap * 2
sx = (SW - total_w) / 2

tiers = [
    ("Presentation", "Browser · Blade · Bootstrap 5"),
    ("Application",  "Laravel 11 · Controllers · Mailables"),
    ("Data",         "MySQL 8 · Eloquent ORM"),
]
for i, (head, sub) in enumerate(tiers):
    x = sx + i * (tier_w + gap)
    add_outline_rect(s, x, tier_y, tier_w, tier_h, GRAY_300, 0.75)
    text(s, x, tier_y + Inches(0.3), tier_w, Inches(0.3),
         f"TIER {i+1:02d}", size=8, bold=True, color=GRAY_400,
         align=PP_ALIGN.CENTER, letter_spacing=300)
    text(s, x, tier_y + Inches(0.55), tier_w, Inches(0.4),
         head, size=18, bold=True, color=INK,
         align=PP_ALIGN.CENTER)
    text(s, x, tier_y + Inches(0.95), tier_w, Inches(0.35),
         sub, size=10, color=GRAY_700, align=PP_ALIGN.CENTER)
    if i < 2:
        add_arrow(s, x + tier_w, tier_y + tier_h / 2,
                  x + tier_w + gap, tier_y + tier_h / 2,
                  rgb=GRAY_400, width=1.0)

# Background services — single thin row underneath
bg_y = Inches(4.5)
bg_h = Inches(0.95)
add_outline_rect(s, sx, bg_y, total_w, bg_h, GRAY_300, 0.75)
text(s, sx + Inches(0.3), bg_y + Inches(0.12), Inches(4), Inches(0.3),
     "BACKGROUND SERVICES", size=8, bold=True, color=GRAY_400,
     letter_spacing=300)
services = [
    ("Scheduler",     "app:check-expirations · daily"),
    ("SMTP",          "Renewal reminder emails"),
    ("Activity Log",  "Per-user audit trail"),
]
svc_w = (total_w - Inches(0.6)) / 3
for i, (head, sub) in enumerate(services):
    x = sx + Inches(0.3) + i * svc_w
    text(s, x, bg_y + Inches(0.4), svc_w, Inches(0.3),
         head, size=11, bold=True, color=INK)
    text(s, x, bg_y + Inches(0.65), svc_w, Inches(0.3),
         sub, size=9.5, color=GRAY_500)

# Module strip — small text-only chips at the bottom
mod_y = Inches(5.85)
text(s, sx, mod_y - Inches(0.3), Inches(6), Inches(0.3),
     "FUNCTIONAL MODULES", size=8, bold=True, color=GRAY_400,
     letter_spacing=300)
modules = ["PC Master", "Device Master", "Subscriptions",
           "Licenses & Contracts", "Notifications", "User Management"]
m_h = Inches(0.5)
mx = sx
for m in modules:
    # measure approx width by character count
    w = Inches(0.25 + 0.105 * len(m))
    add_outline_rect(s, mx, mod_y, w, m_h, GRAY_300, 0.75)
    text(s, mx, mod_y, w, m_h, m, size=10, color=INK,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    mx += w + Inches(0.12)


# ====================================================================
# 05 — Key Features
# ====================================================================
s = add_slide()
page_chrome(s, 5, "Key features", "Capabilities")

features = [
    ("Asset inventory",
     "PCs, devices, software licenses and recurring subscriptions in one place."),
    ("Smart reminders",
     "Daily scheduler emails configurable recipients before renewals lapse."),
    ("Role-based access",
     "Admin / per-module view & edit permissions enforced by middleware."),
    ("Audit trail",
     "Every create, update and delete logged with user, time and field-level diff."),
    ("Bulk import/export",
     "Excel import & export for every module, with downloadable templates."),
    ("Configurable mail",
     "Switch between .env SMTP and database-driven settings without redeploying."),
]
col_w = Inches(3.85)
row_h = Inches(1.6)
gap_x = Inches(0.25)
gap_y = Inches(0.35)
sx = Inches(0.7)
sy = Inches(2.5)

for i, (head, body) in enumerate(features):
    row = i // 3
    col = i % 3
    x = sx + col * (col_w + gap_x)
    y = sy + row * (row_h + gap_y)
    # left rule
    add_line(s, x, y, x, y + Inches(1.2), rgb=ACCENT, width=2.0)
    text(s, x + Inches(0.25), y - Inches(0.02),
         col_w - Inches(0.3), Inches(0.4),
         head, size=15, bold=True, color=INK)
    text(s, x + Inches(0.25), y + Inches(0.42),
         col_w - Inches(0.3), Inches(1.0),
         body, size=10.5, color=GRAY_700)


# ====================================================================
# 06 — Technology Stack
# ====================================================================
s = add_slide()
page_chrome(s, 6, "Technology stack", "Built with")

groups = [
    ("Frontend", [
        "Bootstrap 5.3 · UI & responsive grid",
        "Bootstrap Icons · iconography",
        "Inter font · enterprise typeface",
        "Vanilla JS · light interactivity",
    ]),
    ("Backend", [
        "Laravel 11 · MVC framework, Eloquent ORM",
        "PHP 8.2+ · typed properties, modern syntax",
        "Blade · server-rendered templating",
        "Maatwebsite/Excel · .xlsx import & export",
    ]),
    ("Data & Infra", [
        "MySQL 8 / MariaDB · relational storage",
        "Symfony Mailer · SMTP delivery",
        "XAMPP / Apache · on-prem hosting",
        "Cron scheduler · daily expiry checks",
    ]),
]
col_w = Inches(4.0)
gap = Inches(0.25)
total_w = col_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(2.5)

for i, (head, items) in enumerate(groups):
    x = sx + i * (col_w + gap)
    text(s, x, y, col_w, Inches(0.35),
         head.upper(), size=10, bold=True, color=ACCENT, letter_spacing=200)
    add_line(s, x, y + Inches(0.5),
             x + col_w - Inches(0.3), y + Inches(0.5),
             rgb=GRAY_300, width=0.75)
    iy = y + Inches(0.8)
    for line in items:
        # split label · description on the dot
        if "·" in line:
            label, desc = line.split("·", 1)
            text(s, x, iy, col_w, Inches(0.3),
                 label.strip(), size=12, bold=True, color=INK)
            text(s, x, iy + Inches(0.3), col_w, Inches(0.3),
                 desc.strip(), size=10, color=GRAY_500)
        else:
            text(s, x, iy, col_w, Inches(0.3),
                 line, size=12, bold=True, color=INK)
        iy += Inches(0.72)


# ====================================================================
# 07 — System Workflow
# ====================================================================
s = add_slide()
page_chrome(s, 7, "How it works", "Workflow")

steps = [
    ("01", "Onboard", "Admin adds assets manually or imports an .xlsx file."),
    ("02", "Track",   "Modules show real-time status across the asset estate."),
    ("03", "Detect",  "Daily scheduler scans for items inside the renewal window."),
    ("04", "Notify",  "Email + in-app notifications fire to configured recipients."),
    ("05", "Renew",   "Owners renew or terminate; status & history auto-update."),
]
n = len(steps)
col_w = Inches(2.3)
gap = Inches(0.12)
total_w = col_w * n + gap * (n - 1)
sx = (SW - total_w) / 2
y = Inches(3.0)

for i, (num, head, body) in enumerate(steps):
    x = sx + i * (col_w + gap)
    # number
    text(s, x, y, col_w, Inches(0.6),
         num, size=11, bold=True, color=GRAY_400, letter_spacing=200)
    # dot on timeline
    dot_y = y + Inches(0.7)
    add_rect(s, x + col_w / 2 - Inches(0.06),
             dot_y, Inches(0.12), Inches(0.12), ACCENT)
    # connector to next dot
    if i < n - 1:
        add_line(s, x + col_w / 2 + Inches(0.06),
                 dot_y + Inches(0.06),
                 x + col_w + gap + col_w / 2 - Inches(0.06),
                 dot_y + Inches(0.06),
                 rgb=GRAY_300, width=1.0)
    # title + body
    text(s, x, dot_y + Inches(0.35), col_w, Inches(0.4),
         head, size=16, bold=True, color=INK)
    text(s, x, dot_y + Inches(0.85), col_w, Inches(1.5),
         body, size=10.5, color=GRAY_700)

text(s, Inches(0.7), Inches(6.3), Inches(12), Inches(0.4),
     "Per-day dedupe ensures the same item is not re-notified more than once daily.",
     size=10, color=GRAY_500)


# ====================================================================
# 08 — UI/UX Overview
# ====================================================================
s = add_slide()
page_chrome(s, 8, "Interface", "UI / UX Overview")

# Left: principles
text(s, Inches(0.7), Inches(2.4), Inches(6), Inches(0.5),
     "Designed for daily IT operations.",
     size=20, bold=True, color=INK)

principles = [
    ("Modern dashboard",  "Liquid-glass cards over soft gradient surfaces."),
    ("Light & dark mode", "Theme persists across sessions per user."),
    ("Inter typography",  "Tight letter-spacing, optimized for scanning."),
    ("Accessibility",     "Keyboard navigation, ARIA labels, visible focus."),
    ("Responsive",        "Sidebar rail on tablet; full reflow on mobile."),
]
y = Inches(3.3)
for label, desc in principles:
    add_line(s, Inches(0.7), y + Inches(0.16),
             Inches(0.85), y + Inches(0.16), rgb=ACCENT, width=2.0)
    text(s, Inches(0.95), y - Inches(0.02), Inches(5.5), Inches(0.3),
         label, size=12, bold=True, color=INK)
    text(s, Inches(0.95), y + Inches(0.27), Inches(5.5), Inches(0.4),
         desc, size=10.5, color=GRAY_700)
    y += Inches(0.6)

# Right: outlined browser frame mockup (no fill colors)
mx = Inches(7.5)
my = Inches(2.4)
mw = Inches(5.2)
mh = Inches(4.0)

add_outline_rect(s, mx, my, mw, mh, GRAY_300, 0.75)
# top bar
add_line(s, mx, my + Inches(0.42),
         mx + mw, my + Inches(0.42), rgb=GRAY_300, width=0.5)
for i in range(3):
    cx = mx + Inches(0.15) + Inches(0.22) * i
    add_outline_rect(s, cx, my + Inches(0.13),
                     Inches(0.14), Inches(0.14), GRAY_400, 0.5)
text(s, mx + Inches(1.0), my + Inches(0.1),
     mw - Inches(1.2), Inches(0.25),
     "itams.local / dashboard",
     size=8.5, color=GRAY_500, anchor=MSO_ANCHOR.MIDDLE)

# sidebar
add_line(s, mx + Inches(1.0), my + Inches(0.42),
         mx + Inches(1.0), my + mh, rgb=GRAY_300, width=0.5)
text(s, mx + Inches(0.15), my + Inches(0.6),
     Inches(0.9), Inches(0.25),
     "ITAMS", size=9, bold=True, color=INK,
     letter_spacing=150)
nav = ["Dashboard", "PC Master", "Devices",
       "Subscriptions", "Licenses", "Activity"]
for i, n in enumerate(nav):
    ny = my + Inches(1.0) + Inches(0.32) * i
    if i == 0:
        add_line(s, mx + Inches(0.1), ny + Inches(0.13),
                 mx + Inches(0.22), ny + Inches(0.13),
                 rgb=ACCENT, width=2.0)
        text(s, mx + Inches(0.27), ny, Inches(0.7), Inches(0.25),
             n, size=8.5, bold=True, color=INK,
             anchor=MSO_ANCHOR.MIDDLE)
    else:
        text(s, mx + Inches(0.27), ny, Inches(0.7), Inches(0.25),
             n, size=8.5, color=GRAY_500,
             anchor=MSO_ANCHOR.MIDDLE)

# main area: 4 outlined KPI tiles
kpi_x = mx + Inches(1.25)
kpi_y = my + Inches(0.7)
kpi_w = (mw - Inches(1.4)) / 4 - Inches(0.06)
kpi_h = Inches(0.85)
kpi = [("128", "PCs"), ("64", "Devices"), ("32", "Active"), ("5", "Expiring")]
for i, (v, l) in enumerate(kpi):
    x = kpi_x + i * (kpi_w + Inches(0.07))
    add_outline_rect(s, x, kpi_y, kpi_w, kpi_h, GRAY_300, 0.5)
    text(s, x + Inches(0.12), kpi_y + Inches(0.06),
         kpi_w - Inches(0.2), Inches(0.35),
         v, size=15, bold=True, color=INK)
    text(s, x + Inches(0.12), kpi_y + Inches(0.5),
         kpi_w - Inches(0.2), Inches(0.3),
         l.upper(), size=7.5, color=GRAY_500, letter_spacing=200)

# chart area
chart_y = kpi_y + kpi_h + Inches(0.2)
chart_h = my + mh - chart_y - Inches(0.25)
add_outline_rect(s, kpi_x, chart_y,
                 mw - Inches(1.4), chart_h, GRAY_300, 0.5)
# pretend bars (single accent color, varying heights)
n_bars = 9
heights = [0.45, 0.7, 0.55, 0.85, 0.6, 0.95, 0.7, 0.5, 0.78]
bar_area_w = mw - Inches(1.7)
bar_w = Inches(0.18)
gap_b = (bar_area_w - bar_w * n_bars) / (n_bars - 1)
base = chart_y + chart_h - Inches(0.25)
for i in range(n_bars):
    h = Inches(heights[i] * 1.5)
    bx = kpi_x + Inches(0.15) + i * (bar_w + gap_b)
    add_rect(s, bx, base - h, bar_w, h, INK)


# ====================================================================
# 09 — Benefits / Impact
# ====================================================================
s = add_slide()
page_chrome(s, 9, "Benefits & impact", "Outcomes")

metrics = [
    ("30%",  "lower IT spend",
     "Eliminates duplicate purchases and unused subscriptions."),
    ("0",    "missed renewals",
     "Automatic reminders for every tracked asset."),
    ("100%", "audit coverage",
     "Every change logged with user, timestamp, and diff."),
    ("5×",   "faster onboarding",
     "Bulk Excel import replaces manual entry."),
]
col_w = Inches(2.95)
row_h = Inches(3.0)
gap = Inches(0.2)
total_w = col_w * 4 + gap * 3
sx = (SW - total_w) / 2
y = Inches(2.7)

for i, (value, label, desc) in enumerate(metrics):
    x = sx + i * (col_w + gap)
    # thin top rule
    add_line(s, x, y, x + Inches(0.6), y, rgb=ACCENT, width=2.0)
    text(s, x, y + Inches(0.2), col_w, Inches(1.4),
         value, size=54, bold=True, color=INK)
    text(s, x, y + Inches(1.5), col_w, Inches(0.4),
         label, size=12, bold=True, color=INK)
    text(s, x, y + Inches(1.95), col_w, Inches(1.0),
         desc, size=10.5, color=GRAY_700)


# ====================================================================
# 10 — Challenges
# ====================================================================
s = add_slide()
page_chrome(s, 10, "Challenges we faced", "Challenges")

challenges = [
    ("Mail deliverability",
     "Some SMTP servers reject mailbox addresses without warning. Required graceful per-recipient error handling and configurable recipient lists."),
    ("Notification dedupe",
     "Re-running the daily check could spam recipients. Solved with a per-day uniqueness check on the notifications table."),
    ("Permission granularity",
     "Initial admin/user split was too coarse. Evolved into a per-module view/edit boolean grid to match real team structures."),
    ("Cross-entity expiry",
     "Subscriptions and licenses use different status fields. Refactored CheckExpirations into two parallel passes sharing a recipient resolver."),
]
y = Inches(2.5)
for i, (head, body) in enumerate(challenges):
    text(s, Inches(0.7), y, Inches(1.0), Inches(0.4),
         f"0{i+1}", size=14, bold=True, color=ACCENT, letter_spacing=200)
    text(s, Inches(1.5), y, Inches(11), Inches(0.4),
         head, size=14, bold=True, color=INK)
    text(s, Inches(1.5), y + Inches(0.42), Inches(11), Inches(0.7),
         body, size=11, color=GRAY_700)
    if i < len(challenges) - 1:
        add_line(s, Inches(0.7), y + Inches(1.05),
                 SW - Inches(0.7), y + Inches(1.05),
                 rgb=GRAY_200, width=0.5)
    y += Inches(1.1)


# ====================================================================
# 11 — Future Improvements / Roadmap
# ====================================================================
s = add_slide()
page_chrome(s, 11, "Future improvements", "Roadmap")

# Timeline-style horizontal: 4 quarters with marker + label + description
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
gap = Inches(0.15)
total_w = col_w * n + gap * (n - 1)
sx = (SW - total_w) / 2
y = Inches(2.8)

# horizontal baseline
add_line(s, sx + Inches(0.12), y + Inches(0.6),
         sx + total_w - Inches(0.12), y + Inches(0.6),
         rgb=GRAY_300, width=1.0)

for i, (q, head, body) in enumerate(quarters):
    x = sx + i * (col_w + gap)
    # marker
    dot = Inches(0.18)
    add_rect(s, x + col_w / 2 - dot / 2,
             y + Inches(0.6) - dot / 2,
             dot, dot, ACCENT)
    text(s, x, y, col_w, Inches(0.35),
         q, size=11, bold=True, color=ACCENT,
         align=PP_ALIGN.CENTER, letter_spacing=200)
    text(s, x, y + Inches(1.1), col_w, Inches(0.4),
         head, size=15, bold=True, color=INK,
         align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.15), y + Inches(1.65),
         col_w - Inches(0.3), Inches(1.4),
         body, size=10.5, color=GRAY_700, align=PP_ALIGN.CENTER)


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
card_w = Inches(2.85)
card_h = Inches(3.0)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
sx = (SW - total_w) / 2
y = Inches(2.9)

for i, (mono, name, role) in enumerate(members):
    x = sx + i * (card_w + gap)
    add_outline_rect(s, x, y, card_w, card_h, GRAY_300, 0.75)
    # monogram square
    sq = Inches(0.95)
    msx = x + (card_w - sq) / 2
    msy = y + Inches(0.55)
    add_outline_rect(s, msx, msy, sq, sq, INK, 1.5)
    text(s, msx, msy, sq, sq, mono, size=22, bold=True, color=INK,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x, y + Inches(1.85), card_w, Inches(0.4),
         name, size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
    text(s, x, y + Inches(2.25), card_w, Inches(0.4),
         role, size=10.5, color=GRAY_700, align=PP_ALIGN.CENTER)
    # accent rule
    add_line(s, x + (card_w - Inches(0.4)) / 2, y + card_h - Inches(0.3),
             x + (card_w + Inches(0.4)) / 2, y + card_h - Inches(0.3),
             rgb=ACCENT, width=1.5)


# ====================================================================
# 13 — Conclusion / Summary
# ====================================================================
s = add_slide()
page_chrome(s, 13, "Summary", "Conclusion")

text(s, Inches(0.7), Inches(2.5), Inches(11.5), Inches(1.5),
     "From spreadsheets to a single\nsource of truth.",
     size=32, bold=True, color=INK)

text(s, Inches(0.7), Inches(4.1), Inches(11.5), Inches(0.5),
     "ITAMS gives IT teams the visibility and automation they need.",
     size=14, color=GRAY_700)

# 3 short pillars
takeaways = [
    ("Ready to deploy",
     "Runs on a standard Laravel + LAMP / XAMPP stack."),
    ("Designed for users",
     "Modern UI with dark mode and accessibility built in."),
    ("Built to extend",
     "Modular architecture invites new asset types and integrations."),
]
col_w = Inches(4.0)
gap = Inches(0.25)
total_w = col_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(5.3)

for i, (head, body) in enumerate(takeaways):
    x = sx + i * (col_w + gap)
    add_line(s, x, y, x + Inches(0.5), y, rgb=ACCENT, width=2.0)
    text(s, x, y + Inches(0.2), col_w, Inches(0.4),
         head, size=13, bold=True, color=INK)
    text(s, x, y + Inches(0.65), col_w, Inches(0.9),
         body, size=10.5, color=GRAY_700)


# ====================================================================
# 14 — Thank You (dark, mirrors title slide)
# ====================================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, INK)

text(s, Inches(0), Inches(2.2), SW, Inches(0.4),
     "HACKATHON 2026",
     size=10, bold=True, color=GRAY_500, letter_spacing=300,
     align=PP_ALIGN.CENTER)

text(s, Inches(0), Inches(2.8), SW, Inches(2.0),
     "Thank you",
     size=110, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# divider
dx = (SW - Inches(1)) / 2
add_line(s, dx, Inches(5.05),
         dx + Inches(1), Inches(5.05), rgb=ACCENT, width=2.0)

text(s, Inches(0), Inches(5.2), SW, Inches(0.5),
     "Questions & discussion", size=18, color=GRAY_300,
     align=PP_ALIGN.CENTER)

# footer
text(s, Inches(0), SH - Inches(1.0), SW, Inches(0.4),
     "TEAM ITAMS  ·  IT ASSET MANAGEMENT SYSTEM",
     size=9, bold=True, color=GRAY_500, letter_spacing=300,
     align=PP_ALIGN.CENTER)
text(s, Inches(0), SH - Inches(0.65), SW, Inches(0.4),
     "[ contact email / repo link ]",
     size=10, color=GRAY_500, align=PP_ALIGN.CENTER)


# ---------- Save ----------
out = r"D:\xampp\htdocs\itams\presentation\ITAMS_Hackathon_Minimal.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
