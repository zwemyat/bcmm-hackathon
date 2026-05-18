"""Generate the ITAMS System Specification as a .docx file.

The content is sourced from the current state of the codebase as documented
in CLAUDE.md plus the controllers, models and migrations.
"""

from datetime import date
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------- Brand ----------
NAVY  = RGBColor(0x0F, 0x17, 0x2A)
INK   = RGBColor(0x1F, 0x2D, 0x3D)
GRAY  = RGBColor(0x64, 0x74, 0x8B)
ACCENT = RGBColor(0x25, 0x63, 0xEB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

doc = Document()

# ---------- Page margins ----------
for section in doc.sections:
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

# ---------- Base style ----------
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)
style.font.color.rgb = INK


# ---------- Helpers ----------
def H(text, level=1, *, color=NAVY, size=None, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.bold = True
    run.font.color.rgb = color
    if size is None:
        size = {1: 18, 2: 14, 3: 12, 4: 11}.get(level, 11)
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    if level == 1:
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
    elif level == 2:
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
    else:
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
    return p


def para(text, *, bold=False, italic=False, color=None, align=None, size=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    if size is not None:
        run.font.size = Pt(size)
    return p


def bullet(text, *, bold_lead=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
        p.add_run(" " + text)
    else:
        p.add_run(text)
    return p


def numbered(text):
    p = doc.add_paragraph(style="List Number")
    p.add_run(text)
    return p


def shade_cell(cell, color_hex):
    """Add background color to a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def add_table(headers, rows, *, header_fill="2563EB",
              header_color=WHITE, col_widths=None, first_col_bold=False):
    """Create a formatted table."""
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.style = "Light Grid Accent 1"
    tbl.autofit = False

    # header row
    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        cell = hdr_cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.color.rgb = header_color
        run.font.size = Pt(10.5)
        shade_cell(cell, header_fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # body rows
    for r, row in enumerate(rows, start=1):
        cells = tbl.rows[r].cells
        for i, val in enumerate(row):
            cell = cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(10)
            if first_col_bold and i == 0:
                run.bold = True
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if col_widths:
        for row in tbl.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = w

    para("")  # spacer
    return tbl


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ====================================================================
# COVER
# ====================================================================
for _ in range(4):
    para("")

para("SYSTEM SPECIFICATION", bold=True, color=GRAY,
     align=WD_ALIGN_PARAGRAPH.CENTER, size=11)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("ITAMS")
r.bold = True; r.font.size = Pt(48); r.font.color.rgb = NAVY

para("IT Asset Management System",
     italic=True, color=GRAY,
     align=WD_ALIGN_PARAGRAPH.CENTER, size=16)

for _ in range(3):
    para("")

para("Document version  ·  1.0",
     align=WD_ALIGN_PARAGRAPH.CENTER, color=GRAY, size=10)
para(f"Issued  ·  {date.today().strftime('%d %B %Y')}",
     align=WD_ALIGN_PARAGRAPH.CENTER, color=GRAY, size=10)
para("Status  ·  Released",
     align=WD_ALIGN_PARAGRAPH.CENTER, color=GRAY, size=10)

for _ in range(6):
    para("")

para("Prepared by  ·  INFRA Team",
     align=WD_ALIGN_PARAGRAPH.CENTER, color=NAVY, bold=True, size=12)

page_break()

# ====================================================================
# DOCUMENT CONTROL
# ====================================================================
H("Document Control", level=1)

H("Revision history", level=2)
add_table(
    headers=["Version", "Date", "Author", "Notes"],
    rows=[
        ["1.0", date.today().strftime("%Y-%m-%d"),
         "INFRA Team",
         "Initial release covering all four asset modules, role-based access, "
         "and the staggered renewal reminder pipeline."],
    ],
    col_widths=[Cm(2), Cm(2.5), Cm(3), Cm(8)],
)

H("Intended audience", level=2)
para(
    "This specification is written for product owners, IT operations staff, "
    "internal auditors, and the development team responsible for maintaining "
    "and extending ITAMS. It is also suitable as a hand-off document for "
    "evaluators reviewing the system as part of an internal audit or "
    "procurement exercise."
)

H("Definitions & acronyms", level=2)
add_table(
    headers=["Term", "Meaning"],
    rows=[
        ["ITAMS",        "IT Asset Management System — the product described in this document."],
        ["PC Master",    "Inventory module for end-user workstations and laptops."],
        ["Device Master","Inventory module for network and infrastructure hardware."],
        ["RBAC",         "Role-Based Access Control."],
        ["Reminder window", "The number of days before expiry at which a record becomes eligible for a reminder."],
        ["Day-mark",     "An exact number of days before expiry at which a digest is sent (e.g. 30, 20, 10)."],
        ["Digest",       "A single email batching all records due at a given day-mark."],
    ],
    col_widths=[Cm(4), Cm(11.5)],
)

page_break()

# ====================================================================
# 1. INTRODUCTION
# ====================================================================
H("1. Introduction", level=1)

H("1.1 Purpose", level=2)
para(
    "ITAMS is a web-based platform that gives IT teams a single source of "
    "truth for the hardware, software, and contractual assets they manage. "
    "It replaces ad-hoc spreadsheets with a structured database, adds an "
    "automated renewal-reminder pipeline, and records a full audit trail of "
    "every change."
)

H("1.2 Scope", level=2)
para("In scope:")
bullet("Inventory of PCs, network devices, software licenses and recurring subscriptions.")
bullet("Per-module reminder configuration with staggered email digests.")
bullet("Per-user, per-module access control with full activity logging.")
bullet("Excel import and export for every inventory module.")
bullet("In-app notification centre with per-user read tracking.")

para("Out of scope (initial release):")
bullet("Mobile native applications.")
bullet("Multi-tenant / SaaS deployment with org isolation.")
bullet("Direct integrations with external service desks (Slack, Teams, Jira).")

H("1.3 Operating environment", level=2)
add_table(
    headers=["Layer", "Technology"],
    rows=[
        ["Operating system",   "Windows 11 (development); Windows Server (deployment)."],
        ["Web server",         "Apache via XAMPP."],
        ["PHP runtime",        "PHP 8.2 or later."],
        ["Framework",          "Laravel 11.31+."],
        ["Database",           "MySQL 8 or MariaDB 10.6+."],
        ["Mail transport",     "SMTP (Symfony Mailer); credentials from .env or database."],
        ["Scheduler",          "Windows Task Scheduler invoking php artisan schedule:run every minute."],
        ["Browser support",    "Latest Chrome / Edge / Firefox / Safari."],
    ],
    col_widths=[Cm(4), Cm(11.5)],
)

page_break()

# ====================================================================
# 2. OVERALL DESCRIPTION
# ====================================================================
H("2. Overall description", level=1)

H("2.1 Product perspective", level=2)
para(
    "ITAMS is a stand-alone, on-premise web application. It uses Laravel "
    "for server-side rendering and Bootstrap 5 for the user interface. "
    "There is no separate frontend build step — all client assets load "
    "from a CDN. The application is intended to live behind the organisation's "
    "intranet and authenticates users through its own email/password system, "
    "with role-based access control enforced at the route level."
)

H("2.2 High-level feature summary", level=2)

asset_modules = [
    ("PC Master",
     "Workstations and laptops, with assignee, OS, specs, purchase date, warranty."),
    ("Device Master",
     "Network hardware: routers, switches, servers, printers, with serial numbers."),
    ("Subscriptions",
     "Recurring services such as SSL, domain, cloud or SaaS subscriptions, with cost & renewal cycle."),
    ("Licenses & Contracts",
     "Software licences and vendor contracts with renewal type and cost."),
]
add_table(
    headers=["Module", "Purpose"],
    rows=asset_modules,
    col_widths=[Cm(4.5), Cm(11)],
)

ops_features = [
    ("Smart reminders",
     "Daily scheduler sends one staggered email digest per (module × day-mark)."),
    ("Per-user inbox",
     "Topbar bell + dedicated page; read state tracked individually per user."),
    ("RBAC",
     "Admin role + per-module view & edit boolean grid for non-admins."),
    ("Audit trail",
     "Every create, update, delete logged with user, timestamp and field-level diff."),
    ("Mail configuration",
     "Two sources: .env or a database row, with on-page test send."),
    ("Bulk import / export",
     "Excel (.xlsx) import and export for every inventory module with downloadable templates."),
]
add_table(
    headers=["Operations & governance", "Purpose"],
    rows=ops_features,
    col_widths=[Cm(4.5), Cm(11)],
)

H("2.3 User classes", level=2)
add_table(
    headers=["Role", "Default access", "Typical user"],
    rows=[
        ["Admin",
         "Full read/write across all modules, plus user management, mail settings, notification settings, activity logs.",
         "IT Manager / Operations Lead"],
        ["User",
         "Per-module view and/or edit, configured by an admin per individual user. No access to administration screens.",
         "IT support staff, asset owners"],
    ],
    col_widths=[Cm(2.5), Cm(9), Cm(4)],
)

H("2.4 Assumptions & constraints", level=2)
bullet("The deployment host can execute PHP CLI tasks and reach the SMTP server outbound.")
bullet("Windows Task Scheduler triggers the Laravel scheduler every minute.")
bullet("Users access the system over HTTPS-terminating reverse proxy in production; HTTP is acceptable for local intranet use.")
bullet("Browser JavaScript and the Clipboard API are available in modern browsers; legacy IE is not supported.")

page_break()

# ====================================================================
# 3. SYSTEM ARCHITECTURE
# ====================================================================
H("3. System architecture", level=1)

H("3.1 Architectural overview", level=2)
para(
    "ITAMS follows a conventional three-tier architecture. The presentation "
    "tier is server-rendered HTML with Bootstrap styling; the application "
    "tier is a Laravel monolith implementing controllers, validation, mail, "
    "and scheduled jobs; the data tier is a single MySQL database accessed "
    "through Eloquent."
)

add_table(
    headers=["Tier", "Components", "Responsibilities"],
    rows=[
        ["Presentation",
         "Blade views, Bootstrap 5.3, Bootstrap Icons, Inter font (CDN).",
         "Page rendering, in-browser interactivity, responsive layouts, theme switching."],
        ["Application",
         "Laravel 11 controllers, middleware, mailables, console commands.",
         "Request routing, validation, business logic, mail delivery, scheduling."],
        ["Data",
         "MySQL 8 / MariaDB tables via Eloquent ORM.",
         "Persistent storage of users, assets, settings, activity log."],
    ],
    col_widths=[Cm(3), Cm(6.5), Cm(6)],
)

H("3.2 Background services", level=2)
add_table(
    headers=["Service", "Trigger", "Function"],
    rows=[
        ["app:check-expirations", "Daily 09:00 via Laravel Scheduler",
         "Marks past-due subscriptions Expired; sends staggered reminder digests per module × day-mark."],
        ["SMTP delivery", "Triggered by mailables and console commands",
         "Sends renewal-reminder digests, test emails, and user-credential emails."],
        ["Activity logging", "Triggered inline by controllers",
         "Records every state-changing action with user, IP and user-agent."],
    ],
    col_widths=[Cm(3.5), Cm(4), Cm(8)],
)

H("3.3 Technology stack", level=2)
add_table(
    headers=["Category", "Items"],
    rows=[
        ["Frontend",
         "Bootstrap 5.3, Bootstrap Icons 1.11.3, Inter (Google Fonts), Vanilla JavaScript"],
        ["Backend",
         "Laravel 11, PHP 8.2+, Blade templating, Maatwebsite/Excel"],
        ["Data & infra",
         "MySQL 8 / MariaDB, Symfony Mailer, XAMPP/Apache, Windows Task Scheduler"],
    ],
    col_widths=[Cm(3.5), Cm(12)],
)

page_break()

# ====================================================================
# 4. FUNCTIONAL REQUIREMENTS
# ====================================================================
H("4. Functional requirements", level=1)

H("4.1 Authentication & session management", level=2)
bullet("Users authenticate with email + password.", bold_lead="FR-A1.")
bullet("Failed logins are recorded in the activity log with the attempted email.", bold_lead="FR-A2.")
bullet('"Keep me signed in" extends the session per Laravel\'s "remember" mechanism.', bold_lead="FR-A3.")
bullet("Self-service password reset is not implemented; password changes are admin-driven.", bold_lead="FR-A4.")
bullet("Logout invalidates the session and rotates the CSRF token.", bold_lead="FR-A5.")

H("4.2 User management (admin only)", level=2)
bullet("Admin creates, edits, and deletes user accounts.", bold_lead="FR-U1.")
bullet("Each user has a role (admin or user), profile photo, and a per-module view/edit permission grid.", bold_lead="FR-U2.")
bullet("Granting Edit on a module automatically grants View on that module.", bold_lead="FR-U3.")
bullet("Newly-created users receive a credentials email with their login URL and initial password.", bold_lead="FR-U4.")
bullet("Avatar uploads are stored under storage/app/public/avatars; public access requires the storage symlink/junction.", bold_lead="FR-U5.")

H("4.3 PC Master", level=2)
bullet("CRUD operations on PC assets with searchable, filterable, paginated index.", bold_lead="FR-P1.")
bullet("Bulk Excel import with downloadable template; bulk delete from index.", bold_lead="FR-P2.")
bullet("Sensitive fields (admin password, username/password) are encrypted at rest via Laravel Crypt.", bold_lead="FR-P3.")
bullet("Department, location, and status fields are enum-constrained.", bold_lead="FR-P4.")

H("4.4 Device Master", level=2)
bullet("CRUD operations on network and infrastructure devices.", bold_lead="FR-D1.")
bullet("Bulk Excel import and export, downloadable template.", bold_lead="FR-D2.")
bullet("Serial number is recorded for warranty and traceability.", bold_lead="FR-D3.")

H("4.5 Subscriptions", level=2)
bullet("CRUD operations on recurring subscriptions with vendor, cost, currency, period, expiry.", bold_lead="FR-S1.")
bullet("Renewal status tracked separately from active status (Pending, Renewed, Expired, Cancelled).", bold_lead="FR-S2.")
bullet("Reminder date is auto-calculated as expire_date minus the configured window on each save.", bold_lead="FR-S3.")
bullet("Bulk Excel import/export and a dedicated 'Renew' action that updates renewal status and history.", bold_lead="FR-S4.")

H("4.6 Licenses & Contracts", level=2)
bullet("CRUD operations on software licences and vendor contracts.", bold_lead="FR-L1.")
bullet("Tracks expire date, last renewal date, renewal type, vendor and cost.", bold_lead="FR-L2.")
bullet("Status values: Active, Pending, Expired, Terminated.", bold_lead="FR-L3.")
bullet("Bulk Excel import/export with downloadable template.", bold_lead="FR-L4.")

H("4.7 Notification subsystem", level=2)
bullet("Notifications are computed live from Subscriptions and Licenses; there is no persistent notification table.", bold_lead="FR-N1.")
bullet("Per-module setting controls whether that module produces notifications (enabled flag).", bold_lead="FR-N2.")
bullet("Per-module day-mark set (allowed values 10, 20, 30) controls which day-bucket digests get sent.", bold_lead="FR-N3.")
bullet("Topbar badge shows total items currently in the widest reminder window, minus items already read by the current user.", bold_lead="FR-N4.")
bullet("Per-user reads carry a signature (date + urgency bucket); a stored read is invalidated and the item re-surfaces if the underlying urgency changes.", bold_lead="FR-N5.")
bullet("Notification page supports filters: All / Subscriptions / Licenses & Contracts, and Unread / Read.", bold_lead="FR-N6.")

H("4.8 Mail subsystem", level=2)
bullet("SMTP credentials can be read from .env or from a single database row.", bold_lead="FR-M1.")
bullet("Admins can send a test email from the Mail Settings page to verify connectivity.", bold_lead="FR-M2.")
bullet("Reminder recipients are managed per-module under Notification Settings; empty falls back to all admin users.", bold_lead="FR-M3.")
bullet("Mail password is encrypted at rest.", bold_lead="FR-M4.")
bullet("Database-SMTP overrides are pushed via runtime config + Mail::purge before sending.", bold_lead="FR-M5.")

H("4.9 Activity log", level=2)
bullet("Every state-changing controller action invokes ActivityLogger::log with action, description, subject model, and optional properties.", bold_lead="FR-AL1.")
bullet("Captured fields: user id/name/email, action, polymorphic subject reference, description, IP, user-agent, and arbitrary JSON properties.", bold_lead="FR-AL2.")
bullet("Failed logins are logged with the attempted email but no user id.", bold_lead="FR-AL3.")
bullet("The activity log page is admin-only.", bold_lead="FR-AL4.")

page_break()

# ====================================================================
# 5. DATA MODEL
# ====================================================================
H("5. Data model", level=1)

para(
    "The following tables capture the essential schema as of the current "
    "release. Auxiliary columns (timestamps, soft-delete fields where "
    "applicable, foreign-key indexes) are omitted for clarity."
)

H("5.1 users", level=2)
add_table(
    headers=["Column", "Type", "Notes"],
    rows=[
        ["id", "BIGINT", "Primary key."],
        ["name", "VARCHAR(255)", "Display name."],
        ["email", "VARCHAR(255)", "Unique; login identifier."],
        ["password", "VARCHAR(255)", "Hashed (bcrypt)."],
        ["role", "ENUM", "admin or user."],
        ["avatar", "VARCHAR(255)", "Relative path under storage/app/public/avatars."],
        ["can_view_pc_assets … can_edit_devices", "BOOLEAN", "8 columns: view+edit per module."],
        ["remember_token", "VARCHAR(100)", "Set by 'Keep me signed in'."],
    ],
    col_widths=[Cm(5), Cm(3), Cm(7.5)],
)

H("5.2 pc_assets", level=2)
add_table(
    headers=["Column", "Type", "Notes"],
    rows=[
        ["id", "BIGINT", "Primary key."],
        ["computer_id, hostname, employee_name", "VARCHAR", "Identification fields."],
        ["status, department, location", "ENUM", "Free/Active/Damage/Retirement/Low Performance; IT/HR/Finance/Contract; Office/WFH."],
        ["brand, model, serial_number", "VARCHAR", "Hardware identity."],
        ["cpu, ram, ssd, hdd, display, operating_system", "VARCHAR", "Hardware specs."],
        ["admin_password, username, password", "TEXT (encrypted)", "Encrypted with Laravel Crypt."],
        ["purchased_date, warranty_period", "DATE / VARCHAR", "Lifecycle metadata."],
        ["remarks, modified_by", "TEXT / VARCHAR", "Free text and audit fields."],
    ],
    col_widths=[Cm(6), Cm(4), Cm(5.5)],
)

H("5.3 devices", level=2)
add_table(
    headers=["Column", "Type", "Notes"],
    rows=[
        ["id", "BIGINT", "Primary key."],
        ["device_name, type, brand, model", "VARCHAR", "Identification."],
        ["serial_number", "VARCHAR", "Vendor/warranty traceability."],
        ["status, location, assigned_to", "VARCHAR / VARCHAR / VARCHAR", "Operational state."],
        ["remarks, modified_by", "TEXT / VARCHAR", "Free text + audit."],
    ],
    col_widths=[Cm(5.5), Cm(4.5), Cm(5.5)],
)

H("5.4 subscriptions", level=2)
add_table(
    headers=["Column", "Type", "Notes"],
    rows=[
        ["id", "BIGINT", "Primary key."],
        ["service_type", "VARCHAR", "Domain / SSL / Subscription / Hosting / Cloud, etc."],
        ["project_name, subscription_name, vendor_name", "VARCHAR", "Identification."],
        ["status", "ENUM", "Active, Terminated."],
        ["period", "VARCHAR", "e.g. \"1 Year\"."],
        ["previous_cost, renewal_cost, currency", "DECIMAL(10,2), VARCHAR", "Cost tracking, MMK/JPY/USD."],
        ["expire_date, reminder_date", "DATE", "reminder_date auto-computed on save."],
        ["renewal_type", "VARCHAR", "Yearly, Monthly, Pay-as-you-go, One Time."],
        ["renewal_status", "VARCHAR", "Pending, Renewed, Expired, Cancelled."],
        ["remarks, modified_by", "TEXT / VARCHAR", "Free text + audit."],
    ],
    col_widths=[Cm(6), Cm(4), Cm(5.5)],
)

H("5.5 licenses_contracts", level=2)
add_table(
    headers=["Column", "Type", "Notes"],
    rows=[
        ["id", "BIGINT", "Primary key."],
        ["software_name, vendor_name, license_info", "VARCHAR / TEXT", "Identification."],
        ["status", "VARCHAR", "Active, Pending, Expired, Terminated."],
        ["renewal_type", "VARCHAR", "e.g. Yearly, One Time."],
        ["last_renewal_date, expire_date", "DATE", "Lifecycle."],
        ["previous_cost, renewal_cost, currency", "DECIMAL(10,2), VARCHAR", "Cost tracking."],
        ["remarks, modified_by", "TEXT / VARCHAR", "Free text + audit."],
    ],
    col_widths=[Cm(6), Cm(4), Cm(5.5)],
)

H("5.6 mail_settings (single row)", level=2)
add_table(
    headers=["Column", "Type", "Notes"],
    rows=[
        ["mailer, host, port, encryption, auth_mode", "VARCHAR/INT", "SMTP transport parameters."],
        ["username, password", "VARCHAR / TEXT (encrypted)", "SMTP authentication."],
        ["from_address, from_name", "VARCHAR", "Outgoing identity."],
        ["enabled", "BOOLEAN", "When false, .env credentials are used."],
        ["reminder_days_before, reminder_recipients", "INT / TEXT", "Vestigial — superseded by notification_settings."],
    ],
    col_widths=[Cm(6), Cm(4), Cm(5.5)],
)

H("5.7 notification_settings (one row per module)", level=2)
add_table(
    headers=["Column", "Type", "Notes"],
    rows=[
        ["module", "VARCHAR", "subscriptions or licenses_contracts."],
        ["enabled", "BOOLEAN", "Master switch for that module's notifications."],
        ["days_before_set", "JSON", "Array of day-marks; allowed values {10, 20, 30}."],
        ["recipients", "TEXT", "Email list; falls back to admin users when empty."],
    ],
    col_widths=[Cm(4.5), Cm(3), Cm(8)],
)

H("5.8 notification_reads (per-user)", level=2)
add_table(
    headers=["Column", "Type", "Notes"],
    rows=[
        ["user_id, module, notifiable_id", "FK + VARCHAR + BIGINT", "Composite unique key."],
        ["read_signature", "VARCHAR", "YYYY-MM-DD|<bucket>; stored read is honoured only while live signature matches."],
        ["read_at", "TIMESTAMP", "When the user marked it read."],
    ],
    col_widths=[Cm(5), Cm(4.5), Cm(6)],
)

H("5.9 activity_logs", level=2)
add_table(
    headers=["Column", "Type", "Notes"],
    rows=[
        ["user_id, user_name, user_email", "FK + VARCHAR", "Snapshot of actor at time of action."],
        ["action", "VARCHAR", "login, logout, login_failed, created, updated, deleted, imported, renewed, mail_test, etc."],
        ["subject_type, subject_id", "VARCHAR + BIGINT", "Polymorphic reference to the affected model."],
        ["description", "TEXT", "Human-readable summary."],
        ["properties", "JSON", "Optional context (e.g. changed_fields)."],
        ["ip_address, user_agent", "VARCHAR", "Request metadata."],
    ],
    col_widths=[Cm(5), Cm(3.5), Cm(7)],
)

page_break()

# ====================================================================
# 6. KEY WORKFLOWS
# ====================================================================
H("6. Key workflows", level=1)

H("6.1 Login", level=2)
numbered("User submits email + password to /login.")
numbered("LoginController validates credentials and, on success, regenerates the session.")
numbered("Activity log records action=login with the user as subject.")
numbered("User is redirected to /dashboard (or to the intended URL if any).")
numbered("On failure, activity log records action=login_failed with the attempted email (no user id) and the user sees a generic invalid-credentials message.")

H("6.2 Asset onboarding", level=2)
numbered("Admin or authorised user opens the relevant module index page (e.g. Subscriptions).")
numbered("They either click Create and complete the form, or click Import and upload an .xlsx file matching the downloadable template.")
numbered("Validation runs server-side; on success, the record is saved and an activity-log entry is written (action=created or imported).")
numbered("Subscription saves additionally trigger a model hook that recomputes reminder_date from expire_date minus the configured window.")

H("6.3 Daily renewal-reminder pipeline (CheckExpirations)", level=2)
numbered("Windows Task Scheduler runs php artisan schedule:run every minute.")
numbered("At 09:00 the scheduler invokes app:check-expirations.")
numbered("Phase 1 — Subscriptions past expire_date are flipped to renewal_status=Expired.")
numbered("Phase 2 — For each module × each selected day-mark, rows expiring exactly N days from today are gathered.")
numbered("One ExpiryReminderDigest is sent per (module × day-mark) bucket to the recipient list (or fallback to admins).")
numbered("Subscriptions included in a digest are flipped to renewal_status=Pending; licences are not.")
numbered("There is no per-day dedupe; rerunning the command on the same day will re-send the same digests. Operators are expected not to invoke it manually unless intentional.")

H("6.4 In-app notification consumption", level=2)
numbered("The topbar bell badge calls ExpiryNotificationCounter::summary for the current user, which queries the live Subscription and LicenseContract tables.")
numbered("Each item the user has previously marked read is suppressed only while its live signature matches the stored one.")
numbered("Marking an item read writes a row to notification_reads keyed by (user, module, id) with the current signature.")
numbered("If the underlying record's expire_date or urgency-bucket later changes, the live signature no longer matches the stored one, and the item re-surfaces as unread.")
numbered("'Mark all as read' iterates over the currently-listed items only.")

page_break()

# ====================================================================
# 7. NON-FUNCTIONAL REQUIREMENTS
# ====================================================================
H("7. Non-functional requirements", level=1)

add_table(
    headers=["Category", "Requirement"],
    rows=[
        ["Performance",
         "Index pages with up to 1,000 records load in under 800 ms on standard hardware; queries are paginated (20–30 rows per page)."],
        ["Scalability",
         "Designed for a single organisation; horizontal scaling is not required in the initial release."],
        ["Availability",
         "Single-host deployment; uptime is constrained by the host operating system. Daily scheduler is idempotent against missed runs to within the same calendar day."],
        ["Security",
         "RBAC enforced server-side at the route level; credentials hashed (bcrypt); sensitive PC fields and SMTP password encrypted at rest; CSRF protection on all forms; session regeneration on login."],
        ["Auditability",
         "Every state-changing action is logged with actor, IP, user-agent, and a polymorphic subject reference."],
        ["Maintainability",
         "Single Laravel monolith; conventional MVC layout; no custom build tooling; views are server-rendered Blade with inline styles for module index pages."],
        ["Usability",
         "Light and dark themes; keyboard-navigable forms; visible focus rings; responsive layouts on tablet and mobile."],
        ["Compatibility",
         "Latest Chrome, Edge, Firefox and Safari; PHP 8.2+; MySQL 8 / MariaDB 10.6+."],
        ["Internationalisation",
         "User-facing copy is currently English-only; multi-currency support (MMK, JPY, USD) is provided for cost fields."],
    ],
    col_widths=[Cm(3.5), Cm(12)],
)

page_break()

# ====================================================================
# 8. SECURITY
# ====================================================================
H("8. Security", level=1)

H("8.1 Authentication", level=2)
bullet("Passwords are hashed with bcrypt via the Laravel hashed cast; the plain-text password is never stored.")
bullet("Sessions are regenerated on successful login.")
bullet("Failed login attempts are recorded in activity_logs with the attempted email.")

H("8.2 Authorisation", level=2)
bullet("Route-level middleware enforces both global admin checks (admin alias) and per-module access (module:<name>,<view|edit> alias).")
bullet("Admin role bypasses all per-module checks. For non-admins, edit permission implies view.")
bullet("Direct model binding plus middleware ensures users cannot access records of modules they don't have permission for.")

H("8.3 Data protection", level=2)
bullet("PC Master credential fields (admin_password, username, password) are encrypted with Laravel Crypt.")
bullet("Mail Settings password is encrypted at rest.")
bullet("CSRF tokens are required on all state-changing forms.")
bullet("Avatar uploads are validated for MIME type and size (max 2 MB).")

H("8.4 Audit & monitoring", level=2)
bullet("Every create, update, delete, import, and authentication event is logged with actor metadata.")
bullet("Activity log entries cannot be edited or deleted from the UI.")

page_break()

# ====================================================================
# 9. DEPLOYMENT & INSTALLATION
# ====================================================================
H("9. Deployment & installation", level=1)

H("9.1 Prerequisites", level=2)
bullet("Windows 11 / Windows Server with XAMPP (Apache + PHP 8.2+).")
bullet("MySQL 8 / MariaDB 10.6+ instance reachable from the host.")
bullet("Outbound SMTP access to the chosen mail provider.")
bullet("Composer (composer.phar at the project root is used in this deployment).")

H("9.2 Initial installation", level=2)
numbered("Place the project under XAMPP htdocs (e.g. D:\\xampp\\htdocs\\itams).")
numbered("Configure the Apache vhost or use the default /itams/public alias.")
numbered("Copy .env.example to .env and set DB connection and APP_KEY values.")
numbered("Run php composer.phar install to fetch PHP dependencies.")
numbered("Run php artisan key:generate (if APP_KEY is empty).")
numbered("Run php artisan migrate --seed.")
numbered("Create the public storage link: php artisan storage:link, or on Windows without admin rights, mklink /J public\\storage storage\\app\\public.")
numbered("Configure a Windows Scheduled Task to invoke php artisan schedule:run every minute.")

H("9.3 Post-installation configuration", level=2)
bullet("Sign in with the seeded admin account and immediately rotate the password.")
bullet("Open Mail Settings, enable Database SMTP if applicable, fill in the credentials, and send a Test Email.")
bullet("Open Notification Settings; enable per-module reminders, choose day-marks, and supply recipient emails.")
bullet("Create additional users with appropriate module permissions.")

page_break()

# ====================================================================
# 10. FUTURE ROADMAP
# ====================================================================
H("10. Future roadmap", level=1)
add_table(
    headers=["Phase", "Feature", "Outcome"],
    rows=[
        ["Q1", "Mobile app",
         "Native iOS / Android for on-the-go asset check-in and QR scanning."],
        ["Q2", "Service-desk integrations",
         "Slack, Microsoft Teams, and Jira Service Desk webhooks for renewal alerts."],
        ["Q3", "AI cost insights",
         "Forecast renewal costs and recommend subscription consolidation."],
        ["Q4", "Multi-tenant SaaS",
         "Organisation isolation, SSO / SAML, per-tenant billing and quotas."],
    ],
    col_widths=[Cm(2), Cm(5), Cm(8.5)],
)

page_break()

# ====================================================================
# APPENDIX A — PERMISSIONS MATRIX
# ====================================================================
H("Appendix A — Permissions matrix", level=1)
para(
    "Per-module permissions are stored on the users table as eight boolean "
    "columns (view + edit for each of the four modules). Granting Edit "
    "implicitly grants View. Admins bypass all per-module checks."
)
add_table(
    headers=["Module key", "View column", "Edit column", "Routes guarded"],
    rows=[
        ["pc_assets",          "can_view_pc_assets",          "can_edit_pc_assets",
         "pc-assets.* (index, show, create, store, edit, update, destroy, import, export, template)"],
        ["devices",            "can_view_devices",            "can_edit_devices",
         "devices.* (index, show, create, store, edit, update, destroy, import, export, template)"],
        ["subscriptions",      "can_view_subscriptions",      "can_edit_subscriptions",
         "subscriptions.* (index, create, store, edit, update, destroy, import, export, template, renew)"],
        ["licenses_contracts", "can_view_licenses_contracts", "can_edit_licenses_contracts",
         "licenses-contracts.* (index, create, store, edit, update, destroy, import, export, template)"],
    ],
    col_widths=[Cm(3.5), Cm(3.5), Cm(3.5), Cm(5)],
)

H("Admin-only routes", level=2)
bullet("users.* — user management")
bullet("mail-settings.* — SMTP configuration and test send")
bullet("notification-settings.* — per-module reminder configuration")
bullet("activity-logs.* — audit log viewer")

page_break()

# ====================================================================
# APPENDIX B — MAIL / NOTIFICATION CONFIG MATRIX
# ====================================================================
H("Appendix B — Mail & notification configuration", level=1)

H("B.1 SMTP credential sources", level=2)
add_table(
    headers=["Source", "Selector", "Used by"],
    rows=[
        [".env",
         "mail_settings.enabled = false",
         "All mail send operations by default."],
        ["mail_settings DB row",
         "mail_settings.enabled = true",
         "MailSettingController::sendTest (with runtime config + Mail::purge). CheckExpirations does NOT apply the override and uses .env."],
    ],
    col_widths=[Cm(4), Cm(4.5), Cm(7)],
)

H("B.2 Reminder configuration (per module)", level=2)
add_table(
    headers=["Setting", "Effect"],
    rows=[
        ["enabled",
         "Master switch. When off, no notifications are computed or sent for that module."],
        ["days_before_set",
         "Set of day-marks (10, 20, 30) at which a digest is sent. The largest value also drives the bell-page reminder window."],
        ["recipients",
         "Comma/semicolon/newline separated email list. Empty falls back to all admin users."],
    ],
    col_widths=[Cm(4.5), Cm(11)],
)

H("B.3 Read-signature scheme", level=2)
para(
    "A user's read state for a given item is honoured only while its stored "
    "signature matches the current live signature. The signature is composed "
    "of the expire date and an urgency bucket:"
)
add_table(
    headers=["Days remaining", "Bucket", "Example signature"],
    rows=[
        ["< 0",     "overdue",  "2026-05-15|overdue"],
        ["0 – 7",   "soon",     "2026-05-22|soon"],
        ["> 7",     "upcoming", "2026-06-10|upcoming"],
    ],
    col_widths=[Cm(4), Cm(3.5), Cm(8)],
)
para(
    "If a record previously read in the upcoming bucket later moves into "
    "soon or overdue, its live signature changes and the item re-surfaces "
    "as unread for that user. This prevents stale 'read' state from hiding "
    "items that have become more urgent."
)

# ---------- Save ----------
out = r"D:\xampp\htdocs\itams\docs\System_Specification.docx"
doc.save(out)
print(f"Saved: {out}")
