"""ITAMS hackathon presentation — Premium tech-startup style.

15 slides. Inky dark theme, large display type, asymmetric layouts,
subtle radial-gradient meshes (Linear / Vercel / Arc aesthetic).
Single accent (sky-blue) with a purple gradient pair for accent strips.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from lxml import etree


# ---------- Palette ----------
BG          = RGBColor(0x0A, 0x0E, 0x1A)    # inky near-black
BG_LIFT     = RGBColor(0x13, 0x18, 0x29)    # subtly elevated surface
BG_LIFT_2   = RGBColor(0x1A, 0x20, 0x33)    # second elevation
TEXT        = RGBColor(0xFA, 0xFA, 0xFA)
TEXT_2      = RGBColor(0xA0, 0xA4, 0xB8)
TEXT_3      = RGBColor(0x6B, 0x72, 0x90)
TEXT_MUTED  = RGBColor(0x4B, 0x51, 0x6B)
BORDER      = RGBColor(0x1F, 0x25, 0x38)
BORDER_2    = RGBColor(0x2A, 0x31, 0x47)
ACCENT      = RGBColor(0x60, 0xA5, 0xFA)    # sky-400
ACCENT_DK   = RGBColor(0x3B, 0x82, 0xF6)
ACCENT_PUR  = RGBColor(0xA7, 0x8B, 0xFA)    # purple-400 for gradient pair
POSITIVE    = RGBColor(0x34, 0xD3, 0x99)

FONT = "Calibri"
TOTAL_SLIDES = 15

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ---------- Primitives ----------
def add_slide():
    return prs.slides.add_slide(BLANK)


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


def add_outline_rect(slide, x, y, w, h, *, line_rgb=BORDER, line_pt=0.75,
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


def _strip_fill(spPr):
    for tag in ('a:solidFill', 'a:gradFill', 'a:noFill', 'a:pattFill', 'a:blipFill'):
        existing = spPr.find(qn(tag))
        if existing is not None:
            spPr.remove(existing)


def _insert_gradient(spPr, gradFill):
    """Insert gradFill at the correct position (before a:ln if it exists)."""
    ln = spPr.find(qn('a:ln'))
    if ln is not None:
        spPr.insert(list(spPr).index(ln), gradFill)
    else:
        spPr.append(gradFill)


def add_gradient_rect(slide, x, y, w, h, color_start, color_end,
                      *, angle_deg=0, corner=None,
                      alpha_start=None, alpha_end=None):
    """Linear gradient rectangle. alpha values are 0..100000."""
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    if corner:
        shp.adjustments[0] = corner
    spPr = shp.fill._xPr
    _strip_fill(spPr)

    gradFill = etree.Element(qn('a:gradFill'))
    gradFill.set('flip', 'none')
    gradFill.set('rotWithShape', '1')
    gsLst = etree.SubElement(gradFill, qn('a:gsLst'))

    for pos, color, alpha in [
        (0,       color_start, alpha_start),
        (100000,  color_end,   alpha_end),
    ]:
        gs = etree.SubElement(gsLst, qn('a:gs'))
        gs.set('pos', str(pos))
        srgb = etree.SubElement(gs, qn('a:srgbClr'))
        srgb.set('val', f'{color[0]:02X}{color[1]:02X}{color[2]:02X}')
        if alpha is not None:
            a = etree.SubElement(srgb, qn('a:alpha'))
            a.set('val', str(alpha))

    lin = etree.SubElement(gradFill, qn('a:lin'))
    lin.set('ang', str(int(angle_deg * 60000)))
    lin.set('scaled', '0')

    _insert_gradient(spPr, gradFill)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_mesh_blob(slide, cx, cy, radius, color, alpha=18000):
    """Radial-gradient blob (fades from color@alpha → transparent)."""
    x = cx - radius
    y = cy - radius
    d = radius * 2
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    spPr = shp.fill._xPr
    _strip_fill(spPr)

    gradFill = etree.Element(qn('a:gradFill'))
    gradFill.set('flip', 'none')
    gradFill.set('rotWithShape', '1')
    gsLst = etree.SubElement(gradFill, qn('a:gsLst'))

    # center (opaque-ish)
    gs1 = etree.SubElement(gsLst, qn('a:gs'))
    gs1.set('pos', '0')
    s1 = etree.SubElement(gs1, qn('a:srgbClr'))
    s1.set('val', f'{color[0]:02X}{color[1]:02X}{color[2]:02X}')
    a1 = etree.SubElement(s1, qn('a:alpha'))
    a1.set('val', str(alpha))
    # edge (fully transparent)
    gs2 = etree.SubElement(gsLst, qn('a:gs'))
    gs2.set('pos', '100000')
    s2 = etree.SubElement(gs2, qn('a:srgbClr'))
    s2.set('val', f'{color[0]:02X}{color[1]:02X}{color[2]:02X}')
    a2 = etree.SubElement(s2, qn('a:alpha'))
    a2.set('val', '0')

    path = etree.SubElement(gradFill, qn('a:path'))
    path.set('path', 'circle')
    f = etree.SubElement(path, qn('a:fillToRect'))
    f.set('l', '50000'); f.set('t', '50000')
    f.set('r', '50000'); f.set('b', '50000')

    _insert_gradient(spPr, gradFill)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_line(slide, x1, y1, x2, y2, rgb=BORDER_2, width=0.75):
    line = slide.shapes.add_connector(1, x1, y1, x2, y2)
    line.line.color.rgb = rgb
    line.line.width = Pt(width)
    return line


def add_arrow(slide, x1, y1, x2, y2, rgb=TEXT_3, width=1.0):
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
         size=11, bold=False, color=TEXT,
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


# ---------- Premium chrome ----------
def base_dark(slide, *, mesh=None):
    """Inky background + optional mesh blob(s) for atmosphere."""
    add_rect(slide, 0, 0, SW, SH, BG)
    if mesh:
        for blob in mesh:
            add_mesh_blob(slide, *blob)


def page_chrome(slide, page_no, *, mesh="tr"):
    """Minimal chrome — small wordmark top-left, page count top-right.
    mesh ∈ {"tr", "bl", "tl", None}: location of the subtle mesh blob."""
    blobs = []
    if mesh == "tr":
        blobs = [(SW + Inches(1.5), Inches(-2), Inches(5.5), ACCENT, 14000)]
    elif mesh == "bl":
        blobs = [(Inches(-1.5), SH + Inches(1), Inches(5.5), ACCENT_PUR, 12000)]
    elif mesh == "tl":
        blobs = [(Inches(-1), Inches(-1), Inches(5), ACCENT, 12000)]
    base_dark(slide, mesh=blobs)

    # tiny wordmark top-left
    text(slide, Inches(0.7), Inches(0.45), Inches(2.2), Inches(0.3),
         "ITAMS", size=9, bold=True, color=TEXT_2, letter_spacing=300)
    # small accent dot before page number
    add_rect(slide, SW - Inches(1.55), Inches(0.55),
             Inches(0.06), Inches(0.06), ACCENT, corner=0.5)
    text(slide, SW - Inches(1.4), Inches(0.45), Inches(0.85), Inches(0.3),
         f"{page_no:02d} / {TOTAL_SLIDES}", size=9, bold=True, color=TEXT_2,
         align=PP_ALIGN.RIGHT, letter_spacing=200)


def eyebrow(slide, x, y, label, *, color=ACCENT, with_dot=True):
    """Lowercase eyebrow with optional leading accent dot — modern startup feel."""
    if with_dot:
        add_rect(slide, x, y + Inches(0.13), Inches(0.05), Inches(0.05),
                 color, corner=0.5)
        text(slide, x + Inches(0.15), y, Inches(5), Inches(0.3),
             label.lower(), size=10, bold=False, color=color,
             letter_spacing=200)
    else:
        text(slide, x, y, Inches(5), Inches(0.3),
             label.lower(), size=10, bold=False, color=color,
             letter_spacing=200)


def big_title(slide, x, y, w, h, content, *, size=42, color=TEXT):
    """Display-weight headline with tight letter-spacing."""
    text(slide, x, y, w, h, content, size=size, bold=True, color=color,
         letter_spacing=-25)


def section_block(slide, page_no, eyebrow_text, title_text,
                  *, eyebrow_y=Inches(0.95), title_y=Inches(1.45),
                  title_size=38, mesh="tr"):
    """Standard page setup with eyebrow + display title in top-left quadrant."""
    page_chrome(slide, page_no, mesh=mesh)
    eyebrow(slide, Inches(0.75), eyebrow_y, eyebrow_text)
    big_title(slide, Inches(0.75), title_y, Inches(11.5), Inches(1.3),
              title_text, size=title_size)


# ====================================================================
# 01 — TITLE
# ====================================================================
s = add_slide()
base_dark(s, mesh=[
    (SW - Inches(1), Inches(0.5), Inches(6.5), ACCENT, 22000),
    (SW + Inches(0.5), Inches(3.5), Inches(5), ACCENT_PUR, 16000),
])

# top label
text(s, Inches(0.75), Inches(0.6), Inches(8), Inches(0.3),
     "hackathon 2026 · it infrastructure track",
     size=10, color=TEXT_3, letter_spacing=300)

# small monogram block
add_outline_rect(s, Inches(0.75), Inches(1.1),
                 Inches(0.5), Inches(0.5),
                 line_rgb=BORDER_2, line_pt=1.0, corner=0.2)
text(s, Inches(0.75), Inches(1.1), Inches(0.5), Inches(0.5),
     "I", size=18, bold=True, color=ACCENT,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(1.4), Inches(1.18), Inches(6), Inches(0.4),
     "ITAMS", size=14, bold=True, color=TEXT, letter_spacing=150)

# massive headline, bottom-left
big_title(s, Inches(0.75), Inches(3.3), Inches(12), Inches(2.2),
          "Manage every\nIT asset, beautifully.",
          size=72)

# tagline
text(s, Inches(0.75), Inches(5.7), Inches(11), Inches(0.5),
     "An enterprise IT asset management platform —",
     size=18, color=TEXT_2)
text(s, Inches(0.75), Inches(6.05), Inches(11), Inches(0.5),
     "track devices, licenses & subscriptions, never miss a renewal.",
     size=18, color=TEXT_2)

# bottom rule
add_line(s, Inches(0.75), SH - Inches(0.85),
         Inches(2), SH - Inches(0.85), rgb=ACCENT, width=1.5)

# bottom info
text(s, Inches(0.75), SH - Inches(0.65), Inches(5), Inches(0.3),
     "Team ITAMS", size=12, bold=True, color=TEXT)
text(s, SW - Inches(5), SH - Inches(0.65), Inches(4), Inches(0.3),
     "May 2026", size=12, color=TEXT_2, align=PP_ALIGN.RIGHT)


# ====================================================================
# 02 — PROBLEM
# ====================================================================
s = add_slide()
page_chrome(s, 2, mesh="bl")

# asymmetric: headline left
eyebrow(s, Inches(0.75), Inches(1.05), "the problem")
big_title(s, Inches(0.75), Inches(1.55), Inches(7.5), Inches(3.5),
          "IT teams lose\nvisibility and money\nin spreadsheets.",
          size=44)

# right column: numbered points
points = [
    ("01", "Manual tracking",
     "Devices, licenses & subscriptions scattered across disconnected files."),
    ("02", "Missed renewals",
     "Contracts expire or auto-renew without warning, causing downtime or waste."),
    ("03", "No audit trail",
     "No record of who changed what, when — failing compliance reviews."),
    ("04", "Hidden cost",
     "Duplicate purchases and unused subscriptions inflate IT spend 20–30%."),
]
rx = Inches(8.6)
rw = Inches(4.2)
y = Inches(1.5)
for i, (num, head, body) in enumerate(points):
    if i > 0:
        add_line(s, rx, y - Inches(0.15), rx + rw, y - Inches(0.15),
                 rgb=BORDER, width=0.5)
    text(s, rx, y, Inches(0.8), Inches(0.3),
         num, size=10, bold=True, color=ACCENT, letter_spacing=200)
    text(s, rx + Inches(0.7), y - Inches(0.02), rw - Inches(0.7),
         Inches(0.35), head, size=14, bold=True, color=TEXT)
    text(s, rx + Inches(0.7), y + Inches(0.32), rw - Inches(0.7),
         Inches(0.9), body, size=11, color=TEXT_2)
    y += Inches(1.3)


# ====================================================================
# 03 — SOLUTION
# ====================================================================
s = add_slide()
page_chrome(s, 3, mesh="tr")

# left: headline
eyebrow(s, Inches(0.75), Inches(1.05), "our solution")
big_title(s, Inches(0.75), Inches(1.55), Inches(7), Inches(3.0),
          "A single platform\nfor the entire\nasset lifecycle.",
          size=42)
text(s, Inches(0.75), Inches(5.4), Inches(7), Inches(1.5),
     "ITAMS unifies hardware, devices, licenses, and\n"
     "subscriptions — with automated renewal reminders,\n"
     "role-based access, and a full audit log.",
     size=13, color=TEXT_2)

# right: 3 pillars stacked
pillars = [
    ("Unify",    "One source of truth for every IT asset."),
    ("Automate", "Email & in-app reminders fire before things expire."),
    ("Govern",   "Role-based access with full activity history."),
]
rx = Inches(8.6)
rw = Inches(4.2)
y = Inches(1.55)
for i, (head, body) in enumerate(pillars):
    add_line(s, rx, y, rx + Inches(0.4), y, rgb=ACCENT, width=2.0)
    text(s, rx, y + Inches(0.2), rw, Inches(0.5),
         head, size=24, bold=True, color=TEXT, letter_spacing=-15)
    text(s, rx, y + Inches(0.85), rw, Inches(0.7),
         body, size=11, color=TEXT_2)
    y += Inches(1.55)


# ====================================================================
# 04 — ARCHITECTURE
# ====================================================================
s = add_slide()
section_block(s, 4, "architecture", "System overview.",
              title_size=38, mesh="tr")

tier_w = Inches(3.55)
tier_h = Inches(1.5)
gap = Inches(0.45)
total_w = tier_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(3.0)

tiers = [
    ("01", "Presentation", "Browser · Blade · Bootstrap 5"),
    ("02", "Application",  "Laravel 11 · Controllers · Mailables"),
    ("03", "Data",         "MySQL 8 · Eloquent ORM"),
]
for i, (num, head, sub) in enumerate(tiers):
    x = sx + i * (tier_w + gap)
    add_outline_rect(s, x, y, tier_w, tier_h,
                     line_rgb=BORDER_2, line_pt=0.75,
                     fill_rgb=BG_LIFT, corner=0.05)
    # subtle top accent
    add_gradient_rect(s, x, y, tier_w, Inches(0.04),
                      ACCENT, ACCENT_PUR, angle_deg=0)
    text(s, x + Inches(0.4), y + Inches(0.25), Inches(2), Inches(0.3),
         f"tier {num}", size=9, color=ACCENT, letter_spacing=200)
    text(s, x + Inches(0.4), y + Inches(0.55), tier_w - Inches(0.8),
         Inches(0.4), head, size=18, bold=True, color=TEXT, letter_spacing=-15)
    text(s, x + Inches(0.4), y + Inches(1.05), tier_w - Inches(0.8),
         Inches(0.4), sub, size=10, color=TEXT_2)
    if i < 2:
        add_arrow(s, x + tier_w + Inches(0.05), y + tier_h / 2,
                  x + tier_w + gap - Inches(0.05), y + tier_h / 2,
                  rgb=TEXT_3, width=1.0)

# background services
bg_y = Inches(4.85)
bg_h = Inches(0.95)
add_outline_rect(s, sx, bg_y, total_w, bg_h,
                 line_rgb=BORDER_2, line_pt=0.5,
                 fill_rgb=BG_LIFT, corner=0.05)
text(s, sx + Inches(0.4), bg_y + Inches(0.15), Inches(6), Inches(0.25),
     "background services", size=9, color=ACCENT, letter_spacing=200)
services = [
    ("Scheduler",    "app:check-expirations · daily 09:00"),
    ("SMTP",         "Renewal reminder emails to recipients"),
    ("Activity Log", "Per-user audit trail of every change"),
]
svc_w = (total_w - Inches(0.8)) / 3
for i, (head, sub) in enumerate(services):
    x = sx + Inches(0.4) + i * svc_w
    text(s, x, bg_y + Inches(0.45), svc_w, Inches(0.3),
         head, size=11, bold=True, color=TEXT)
    text(s, x, bg_y + Inches(0.7), svc_w, Inches(0.25),
         sub, size=9.5, color=TEXT_3)

# modules
mod_y = Inches(6.2)
text(s, sx, mod_y, Inches(6), Inches(0.3),
     "functional modules", size=9, color=TEXT_3, letter_spacing=200)
modules = ["PC Master", "Device Master", "Subscriptions",
           "Licenses & Contracts", "Notifications", "User Management"]
mx = sx
my = mod_y + Inches(0.35)
for m in modules:
    w = Inches(0.4 + 0.105 * len(m))
    add_outline_rect(s, mx, my, w, Inches(0.42),
                     line_rgb=BORDER_2, line_pt=0.5,
                     fill_rgb=BG_LIFT, corner=0.5)
    text(s, mx, my, w, Inches(0.42), m, size=10, color=TEXT_2,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    mx += w + Inches(0.12)


# ====================================================================
# 05 — KEY FEATURES 1/2
# ====================================================================
def features_grid(slide, items, *, sy=Inches(2.8)):
    """3×2 grid of numbered feature cells, no card containers."""
    col_w = Inches(3.95)
    row_h = Inches(1.95)
    gap_x = Inches(0.2)
    gap_y = Inches(0.4)
    sx = Inches(0.7)
    for i, (num, head, body) in enumerate(items):
        row = i // 3
        col = i % 3
        x = sx + col * (col_w + gap_x)
        y = sy + row * (row_h + gap_y)
        add_line(slide, x, y, x + Inches(0.5), y,
                 rgb=ACCENT, width=2.0)
        text(slide, x, y + Inches(0.18), Inches(1.2), Inches(0.3),
             num, size=9, bold=True, color=ACCENT, letter_spacing=200)
        text(slide, x, y + Inches(0.55), col_w - Inches(0.2),
             Inches(0.45), head, size=17, bold=True, color=TEXT,
             letter_spacing=-10)
        text(slide, x, y + Inches(1.05), col_w - Inches(0.3),
             Inches(0.85), body, size=10.5, color=TEXT_2)


s = add_slide()
section_block(s, 5, "capabilities · 1 of 2", "Asset modules.",
              title_size=36, mesh="tr")

features_a = [
    ("01", "PC Master",
     "Track workstations: assignee, OS, specs, purchase date, warranty."),
    ("02", "Device Master",
     "Network hardware inventory — routers, switches, servers, printers."),
    ("03", "Subscriptions",
     "Recurring services with renewal cycles, vendor and cost history."),
    ("04", "Licenses & Contracts",
     "Software licenses and vendor contracts with expiry tracking."),
    ("05", "Bulk import/export",
     "Excel import & export per module, with downloadable templates."),
    ("06", "Searchable everywhere",
     "Filter, search and bulk-delete across the whole asset estate."),
]
features_grid(s, features_a)


# ====================================================================
# 06 — KEY FEATURES 2/2
# ====================================================================
s = add_slide()
section_block(s, 6, "capabilities · 2 of 2", "Operations & governance.",
              title_size=36, mesh="bl")

features_b = [
    ("01", "Smart reminders",
     "Daily scheduler sends staggered email digests at the day-marks you choose."),
    ("02", "Per-user inbox",
     "Read state tracked individually; items re-surface when urgency shifts."),
    ("03", "Role-based access",
     "Admin / per-module view & edit permissions enforced by middleware."),
    ("04", "Audit trail",
     "Every create, update, delete logged with user, time and field diff."),
    ("05", "Configurable mail",
     "Switch between .env SMTP and database settings without redeploying."),
    ("06", "Live status badge",
     "Topbar bell counts overdue and due-soon items in real time."),
]
features_grid(s, features_b)


# ====================================================================
# 07 — TECHNOLOGY STACK
# ====================================================================
s = add_slide()
section_block(s, 7, "technology", "Built with.",
              title_size=38, mesh="tr")

groups = [
    ("Frontend", [
        ("Bootstrap 5.3",    "UI & responsive grid"),
        ("Bootstrap Icons",  "Iconography"),
        ("Inter font",       "Enterprise typeface"),
        ("Vanilla JS",       "Lightweight interactivity"),
    ]),
    ("Backend", [
        ("Laravel 11",       "MVC framework, Eloquent ORM"),
        ("PHP 8.2+",         "Typed properties, modern syntax"),
        ("Blade",            "Server-rendered templating"),
        ("Maatwebsite/Excel","Import & export to .xlsx"),
    ]),
    ("Data & Infra", [
        ("MySQL 8 / MariaDB","Relational storage"),
        ("Symfony Mailer",   "SMTP delivery via Laravel"),
        ("XAMPP / Apache",   "On-prem hosting"),
        ("Cron scheduler",   "Daily expiry checks"),
    ]),
]
col_w = Inches(3.95)
gap = Inches(0.25)
total_w = col_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(3.0)
for i, (head, items) in enumerate(groups):
    x = sx + i * (col_w + gap)
    text(s, x, y, col_w, Inches(0.3),
         head, size=11, bold=True, color=TEXT, letter_spacing=100)
    add_line(s, x, y + Inches(0.45),
             x + col_w - Inches(0.2), y + Inches(0.45),
             rgb=ACCENT, width=1.0)
    iy = y + Inches(0.8)
    for label, desc in items:
        text(s, x, iy, col_w, Inches(0.3),
             label, size=13, bold=True, color=TEXT)
        text(s, x, iy + Inches(0.28), col_w - Inches(0.3),
             Inches(0.3), desc, size=10, color=TEXT_3)
        iy += Inches(0.78)


# ====================================================================
# 08 — WORKFLOW
# ====================================================================
s = add_slide()
section_block(s, 8, "workflow", "How it works.",
              title_size=38, mesh="bl")

steps = [
    ("01", "Onboard", "Admin adds assets or imports an .xlsx file."),
    ("02", "Track",   "Modules show real-time status across the estate."),
    ("03", "Detect",  "Daily scheduler scans for items in the renewal window."),
    ("04", "Notify",  "Email + in-app alerts fire to configured recipients."),
    ("05", "Renew",   "Owners renew or terminate; status auto-updates."),
]
n = len(steps)
col_w = Inches(2.4)
gap = Inches(0.1)
total_w = col_w * n + gap * (n - 1)
sx = (SW - total_w) / 2
y = Inches(3.2)

# big horizontal gradient line as the timeline
line_y = y + Inches(0.85)
add_gradient_rect(s, sx + Inches(0.2), line_y - Inches(0.015),
                  total_w - Inches(0.4), Inches(0.03),
                  ACCENT, ACCENT_PUR, angle_deg=0)

for i, (num, head, body) in enumerate(steps):
    x = sx + i * (col_w + gap)
    # large step number above the line
    text(s, x, y, col_w, Inches(0.7),
         num, size=44, bold=True, color=TEXT,
         align=PP_ALIGN.CENTER, letter_spacing=-25)
    # dot ON the line
    add_rect(s, x + col_w / 2 - Inches(0.09),
             line_y - Inches(0.09),
             Inches(0.18), Inches(0.18), ACCENT, corner=0.5)
    add_rect(s, x + col_w / 2 - Inches(0.05),
             line_y - Inches(0.05),
             Inches(0.1), Inches(0.1), BG, corner=0.5)
    # heading + body BELOW
    text(s, x, line_y + Inches(0.25), col_w, Inches(0.4),
         head, size=16, bold=True, color=TEXT,
         align=PP_ALIGN.CENTER, letter_spacing=-10)
    text(s, x + Inches(0.15), line_y + Inches(0.75),
         col_w - Inches(0.3), Inches(1.5),
         body, size=10.5, color=TEXT_2, align=PP_ALIGN.CENTER)


# ====================================================================
# 09 — UI/UX
# ====================================================================
s = add_slide()
section_block(s, 9, "interface", "Designed for daily use.",
              title_size=38, mesh="tr")

# left: principles
principles = [
    ("Modern dashboard",   "Liquid-glass cards over soft gradient surfaces."),
    ("Light & dark mode",  "Theme persists across sessions per user."),
    ("Modern typography",  "Inter, tight letter-spacing, optimized for scanning."),
    ("Accessibility",      "Keyboard nav, ARIA labels, visible focus rings."),
    ("Responsive",         "Sidebar rail on tablet; full reflow on mobile."),
]
y = Inches(3.1)
for label, desc in principles:
    add_rect(s, Inches(0.75), y + Inches(0.13),
             Inches(0.06), Inches(0.06), ACCENT, corner=0.5)
    text(s, Inches(0.95), y - Inches(0.02), Inches(5.5), Inches(0.3),
         label, size=13, bold=True, color=TEXT)
    text(s, Inches(0.95), y + Inches(0.3), Inches(5.5), Inches(0.4),
         desc, size=10.5, color=TEXT_2)
    y += Inches(0.65)

# right: outlined browser mockup
mx = Inches(7.6)
my = Inches(2.95)
mw = Inches(5.2)
mh = Inches(3.9)
add_outline_rect(s, mx, my, mw, mh,
                 line_rgb=BORDER_2, line_pt=0.75,
                 fill_rgb=BG_LIFT, corner=0.04)
# top bar
add_line(s, mx, my + Inches(0.42), mx + mw, my + Inches(0.42),
         rgb=BORDER, width=0.5)
for i in range(3):
    cx = mx + Inches(0.18) + Inches(0.24) * i
    add_outline_rect(s, cx, my + Inches(0.14),
                     Inches(0.14), Inches(0.14),
                     line_rgb=TEXT_MUTED, line_pt=0.5, corner=0.5)
# URL pill
add_outline_rect(s, mx + Inches(1.0), my + Inches(0.1),
                 Inches(4.0), Inches(0.26),
                 line_rgb=BORDER, line_pt=0.5,
                 fill_rgb=BG_LIFT_2, corner=0.3)
text(s, mx + Inches(1.1), my + Inches(0.1),
     Inches(3.8), Inches(0.26),
     "itams.local / dashboard",
     size=8.5, color=TEXT_3, anchor=MSO_ANCHOR.MIDDLE)

# sidebar
add_line(s, mx + Inches(1.05), my + Inches(0.42),
         mx + Inches(1.05), my + mh, rgb=BORDER, width=0.5)
add_outline_rect(s, mx + Inches(0.13), my + Inches(0.58),
                 Inches(0.3), Inches(0.3),
                 line_rgb=ACCENT, line_pt=1.0, corner=0.25)
text(s, mx + Inches(0.13), my + Inches(0.58),
     Inches(0.3), Inches(0.3), "I", size=10, bold=True, color=ACCENT,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, mx + Inches(0.53), my + Inches(0.62),
     Inches(0.55), Inches(0.25),
     "ITAMS", size=8.5, bold=True, color=TEXT, letter_spacing=150)

nav = ["Dashboard", "PC Master", "Devices",
       "Subscriptions", "Licenses", "Activity"]
for i, name in enumerate(nav):
    ny = my + Inches(1.05) + Inches(0.32) * i
    if i == 0:
        add_rect(s, mx + Inches(0.1), ny - Inches(0.04),
                 Inches(0.9), Inches(0.28), BG_LIFT_2, corner=0.25)
        text(s, mx + Inches(0.2), ny, Inches(0.8), Inches(0.22),
             name, size=8.5, bold=True, color=ACCENT, anchor=MSO_ANCHOR.MIDDLE)
    else:
        text(s, mx + Inches(0.2), ny, Inches(0.8), Inches(0.22),
             name, size=8.5, color=TEXT_3, anchor=MSO_ANCHOR.MIDDLE)

# main area: KPI tiles
kpi_x = mx + Inches(1.25)
kpi_y = my + Inches(0.7)
kpi_w = (mw - Inches(1.4)) / 4 - Inches(0.07)
kpi_h = Inches(0.85)
kpi = [("128", "PCs"), ("64", "Devices"),
       ("32", "Active"), ("5", "Expiring")]
for i, (v, l) in enumerate(kpi):
    x = kpi_x + i * (kpi_w + Inches(0.08))
    add_outline_rect(s, x, kpi_y, kpi_w, kpi_h,
                     line_rgb=BORDER, line_pt=0.5,
                     fill_rgb=BG_LIFT_2, corner=0.08)
    text(s, x + Inches(0.13), kpi_y + Inches(0.1),
         kpi_w - Inches(0.2), Inches(0.35),
         v, size=15, bold=True, color=TEXT)
    text(s, x + Inches(0.13), kpi_y + Inches(0.5),
         kpi_w - Inches(0.2), Inches(0.3),
         l.upper(), size=7.5, color=TEXT_3, letter_spacing=200)

# chart area
chart_y = kpi_y + kpi_h + Inches(0.18)
chart_h = my + mh - chart_y - Inches(0.22)
add_outline_rect(s, kpi_x, chart_y, mw - Inches(1.4), chart_h,
                 line_rgb=BORDER, line_pt=0.5,
                 fill_rgb=BG_LIFT_2, corner=0.05)
heights = [0.42, 0.66, 0.55, 0.85, 0.6, 0.95, 0.7, 0.5, 0.78]
n_bars = len(heights)
bar_w = Inches(0.18)
bar_area = mw - Inches(1.8)
gap_b = (bar_area - bar_w * n_bars) / (n_bars - 1)
base = chart_y + chart_h - Inches(0.2)
for i in range(n_bars):
    h = Inches(heights[i] * 1.3)
    bx = kpi_x + Inches(0.2) + i * (bar_w + gap_b)
    add_gradient_rect(s, bx, base - h, bar_w, h,
                      ACCENT_PUR, ACCENT,
                      angle_deg=90, corner=0.2)


# ====================================================================
# 10 — BENEFITS / IMPACT
# ====================================================================
s = add_slide()
section_block(s, 10, "outcomes", "Benefits & impact.",
              title_size=38, mesh="tr")

# 2x2 quadrants with very large numbers
metrics = [
    ("30%",  "lower IT spend",
     "Eliminates duplicate purchases and unused subscriptions."),
    ("0",    "missed renewals",
     "Automatic reminders for every tracked asset."),
    ("100%", "audit coverage",
     "Every change logged with user, time and field diff."),
    ("5×",   "faster onboarding",
     "Bulk Excel import replaces manual entry."),
]
quad_w = Inches(5.95)
quad_h = Inches(1.95)
gap_x = Inches(0.25)
gap_y = Inches(0.25)
sx = (SW - quad_w * 2 - gap_x) / 2
sy = Inches(2.9)
for i, (value, label, desc) in enumerate(metrics):
    row = i // 2
    col = i % 2
    x = sx + col * (quad_w + gap_x)
    y = sy + row * (quad_h + gap_y)
    add_line(s, x, y, x + Inches(0.5), y, rgb=ACCENT, width=2.0)
    # massive number
    text(s, x, y + Inches(0.18), Inches(2.5), Inches(1.5),
         value, size=64, bold=True, color=TEXT, letter_spacing=-30)
    text(s, x + Inches(2.4), y + Inches(0.4), quad_w - Inches(2.5),
         Inches(0.5), label, size=15, bold=True, color=TEXT)
    text(s, x + Inches(2.4), y + Inches(0.95), quad_w - Inches(2.5),
         Inches(0.9), desc, size=11, color=TEXT_2)


# ====================================================================
# 11 — CHALLENGES
# ====================================================================
s = add_slide()
section_block(s, 11, "challenges", "What we worked through.",
              title_size=38, mesh="bl")

challenges = [
    ("Mail deliverability",
     "Some SMTP servers reject mailbox addresses without warning. Solved with graceful per-recipient error handling and configurable lists."),
    ("Notification dedupe",
     "Re-running the daily check risked spamming recipients. Per-day uniqueness check on the notifications table prevents duplicates."),
    ("Permission granularity",
     "Initial admin/user split was too coarse. Evolved into a per-module view/edit boolean grid matching real team structures."),
    ("Cross-entity expiry",
     "Subscriptions and licenses use different status fields. Refactored into two parallel passes sharing a recipient resolver."),
]
y = Inches(3.0)
for i, (head, body) in enumerate(challenges):
    add_line(s, Inches(0.75), y,
             Inches(0.95), y, rgb=ACCENT, width=2.0)
    text(s, Inches(0.75), y + Inches(0.15), Inches(0.8), Inches(0.4),
         f"0{i+1}", size=10, bold=True, color=ACCENT, letter_spacing=200)
    text(s, Inches(1.5), y + Inches(0.1), Inches(11), Inches(0.4),
         head, size=15, bold=True, color=TEXT, letter_spacing=-10)
    text(s, Inches(1.5), y + Inches(0.5), SW - Inches(2.25),
         Inches(0.5), body, size=11, color=TEXT_2)
    y += Inches(1.0)
    if i < len(challenges) - 1:
        add_line(s, Inches(0.75), y - Inches(0.1),
                 SW - Inches(0.75), y - Inches(0.1),
                 rgb=BORDER, width=0.5)


# ====================================================================
# 12 — ROADMAP
# ====================================================================
s = add_slide()
section_block(s, 12, "roadmap", "What's next.",
              title_size=38, mesh="tr")

quarters = [
    ("Q1", "Mobile app",
     "Native iOS / Android for on-the-go asset check-in & QR scanning."),
    ("Q2", "Integrations",
     "Slack, Microsoft Teams, and Jira Service Desk webhooks."),
    ("Q3", "AI insights",
     "Forecast costs and recommend consolidation opportunities."),
    ("Q4", "Multi-tenant SaaS",
     "Org isolation, SSO / SAML, per-tenant billing."),
]
n = len(quarters)
col_w = Inches(2.85)
gap = Inches(0.2)
total_w = col_w * n + gap * (n - 1)
sx = (SW - total_w) / 2
y = Inches(3.2)
line_y = y + Inches(0.95)
# gradient timeline
add_gradient_rect(s, sx + Inches(0.2), line_y - Inches(0.015),
                  total_w - Inches(0.4), Inches(0.03),
                  ACCENT, ACCENT_PUR, angle_deg=0)
for i, (q, head, body) in enumerate(quarters):
    x = sx + i * (col_w + gap)
    # quarter pill
    add_outline_rect(s, x + col_w / 2 - Inches(0.45),
                     y, Inches(0.9), Inches(0.42),
                     line_rgb=ACCENT, line_pt=1.0,
                     fill_rgb=BG_LIFT, corner=0.45)
    text(s, x + col_w / 2 - Inches(0.45), y,
         Inches(0.9), Inches(0.42),
         q, size=11, bold=True, color=ACCENT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, letter_spacing=150)
    # dot on timeline
    add_rect(s, x + col_w / 2 - Inches(0.09),
             line_y - Inches(0.09),
             Inches(0.18), Inches(0.18), ACCENT, corner=0.5)
    add_rect(s, x + col_w / 2 - Inches(0.05),
             line_y - Inches(0.05),
             Inches(0.1), Inches(0.1), BG, corner=0.5)
    # title + body below
    text(s, x, line_y + Inches(0.3), col_w, Inches(0.45),
         head, size=17, bold=True, color=TEXT,
         align=PP_ALIGN.CENTER, letter_spacing=-10)
    text(s, x + Inches(0.15), line_y + Inches(0.9),
         col_w - Inches(0.3), Inches(1.5),
         body, size=10.5, color=TEXT_2, align=PP_ALIGN.CENTER)


# ====================================================================
# 13 — TEAM
# ====================================================================
s = add_slide()
section_block(s, 13, "team", "The people behind ITAMS.",
              title_size=38, mesh="bl")

members = [
    ("PL", "[Member 1]", "Project Lead / Backend",   ACCENT),
    ("FE", "[Member 2]", "Frontend / UI Engineer",   ACCENT_PUR),
    ("FS", "[Member 3]", "Full-Stack Developer",     POSITIVE),
    ("DB", "[Member 4]", "Database / DevOps",        ACCENT_DK),
]
card_w = Inches(2.95)
card_h = Inches(3.0)
gap = Inches(0.2)
total_w = card_w * 4 + gap * 3
sx = (SW - total_w) / 2
y = Inches(2.95)
for i, (mono, name, role, color) in enumerate(members):
    x = sx + i * (card_w + gap)
    add_outline_rect(s, x, y, card_w, card_h,
                     line_rgb=BORDER_2, line_pt=0.5,
                     fill_rgb=BG_LIFT, corner=0.05)
    # avatar — gradient square in member color → accent
    av = Inches(1.25)
    avx = x + (card_w - av) / 2
    avy = y + Inches(0.45)
    add_gradient_rect(s, avx, avy, av, av,
                      (color[0], color[1], color[2]),
                      (ACCENT_PUR[0], ACCENT_PUR[1], ACCENT_PUR[2]),
                      angle_deg=45, corner=0.15)
    text(s, avx, avy, av, av, mono, size=28, bold=True, color=TEXT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         letter_spacing=-10)
    text(s, x, y + Inches(2.0), card_w, Inches(0.4),
         name, size=14, bold=True, color=TEXT, align=PP_ALIGN.CENTER)
    text(s, x, y + Inches(2.4), card_w, Inches(0.4),
         role, size=10.5, color=TEXT_2, align=PP_ALIGN.CENTER)


# ====================================================================
# 14 — CONCLUSION
# ====================================================================
s = add_slide()
page_chrome(s, 14, mesh="tr")

eyebrow(s, Inches(0.75), Inches(1.05), "in summary")

# big closing statement, asymmetric left
big_title(s, Inches(0.75), Inches(1.65), Inches(12), Inches(2.5),
          "From spreadsheets\nto a single source of truth.",
          size=46)
text(s, Inches(0.75), Inches(4.25), Inches(11), Inches(0.6),
     "ITAMS gives IT teams the visibility and automation they need.",
     size=15, color=TEXT_2)

# 3 takeaways in a row
takeaways = [
    ("Ready to deploy",     "Runs on a standard Laravel + LAMP / XAMPP stack."),
    ("Designed for users",  "Modern UI with dark mode and accessibility built in."),
    ("Built to extend",     "Modular architecture invites new asset types & integrations."),
]
col_w = Inches(4.0)
gap = Inches(0.25)
total_w = col_w * 3 + gap * 2
sx = (SW - total_w) / 2
y = Inches(5.5)
for i, (head, body) in enumerate(takeaways):
    x = sx + i * (col_w + gap)
    add_line(s, x, y, x + Inches(0.4), y, rgb=ACCENT, width=2.0)
    text(s, x, y + Inches(0.2), col_w, Inches(0.4),
         head, size=14, bold=True, color=TEXT, letter_spacing=-10)
    text(s, x, y + Inches(0.7), col_w - Inches(0.2), Inches(0.9),
         body, size=10.5, color=TEXT_2)


# ====================================================================
# 15 — THANK YOU
# ====================================================================
s = add_slide()
base_dark(s, mesh=[
    (SW / 2, SH / 2 + Inches(0.5), Inches(7), ACCENT, 25000),
    (SW / 2 + Inches(3), SH / 2 - Inches(1), Inches(5), ACCENT_PUR, 18000),
])

# top eyebrow
text(s, Inches(0), Inches(2.0), SW, Inches(0.4),
     "hackathon 2026",
     size=10, color=TEXT_3, letter_spacing=400, align=PP_ALIGN.CENTER)

# massive headline
big_title(s, Inches(0), Inches(2.6), SW, Inches(2.5),
          "Thank you.", size=140)

# divider
add_gradient_rect(s, (SW - Inches(1.5)) / 2, Inches(5.05),
                  Inches(1.5), Inches(0.04),
                  ACCENT, ACCENT_PUR, angle_deg=0, corner=0.5)

text(s, Inches(0), Inches(5.25), SW, Inches(0.5),
     "Questions & discussion", size=18, color=TEXT_2,
     align=PP_ALIGN.CENTER)

# bottom footer
text(s, Inches(0), SH - Inches(1.0), SW, Inches(0.4),
     "team itams  ·  it asset management system",
     size=10, color=TEXT_3, letter_spacing=400, align=PP_ALIGN.CENTER)
text(s, Inches(0), SH - Inches(0.65), SW, Inches(0.4),
     "[ contact email · repo link ]",
     size=11, color=TEXT_MUTED, align=PP_ALIGN.CENTER)


# ---------- Save ----------
out = r"D:\xampp\htdocs\itams\presentation\ITAMS_Hackathon_Premium.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
