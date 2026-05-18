"""Generate the ITAMS Test Cases workbook.

One sheet per module. Each row is a test case with Status / Tester / Date
columns left blank for the QA team to fill in.
"""

from datetime import date
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, NamedStyle
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule


# ---------- Palette ----------
NAVY      = "0F172A"
ACCENT    = "2563EB"
GRAY_50   = "F8FAFC"
GRAY_100  = "F1F5F9"
GRAY_200  = "E2E8F0"
GRAY_300  = "CBD5E1"
GRAY_500  = "64748B"
INK       = "1F2D3D"
RED       = "DC2626"
AMBER     = "F59E0B"
GREEN     = "10B981"
WHITE     = "FFFFFF"


# ---------- Common styles ----------
def thin(color=GRAY_200):
    s = Side(border_style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


HEADER_FILL  = PatternFill("solid", fgColor=NAVY)
HEADER_FONT  = Font(name="Calibri", size=10, bold=True, color=WHITE)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
HEADER_BORDER = thin(NAVY)

BODY_FONT  = Font(name="Calibri", size=10, color=INK)
BODY_ALIGN_WRAP   = Alignment(horizontal="left",   vertical="top",   wrap_text=True)
BODY_ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
BODY_BORDER = thin(GRAY_200)
BAND_FILL   = PatternFill("solid", fgColor=GRAY_50)

TITLE_FONT  = Font(name="Calibri", size=22, bold=True, color=NAVY)
SUB_FONT    = Font(name="Calibri", size=12, bold=True, color=ACCENT)
LABEL_FONT  = Font(name="Calibri", size=10, bold=True, color=GRAY_500)
META_FONT   = Font(name="Calibri", size=11, color=INK)
NOTE_FONT   = Font(name="Calibri", size=10, italic=True, color=GRAY_500)


# ---------- Workbook ----------
wb = Workbook()


HEADERS = [
    "Test ID", "Module", "Title", "Priority",
    "Preconditions", "Steps", "Expected Result",
    "Actual Result", "Status", "Tester", "Date", "Notes",
]
COL_WIDTHS = {
    "A": 12, "B": 16, "C": 32, "D": 11,
    "E": 28, "F": 42, "G": 38,
    "H": 26, "I": 12, "J": 14, "K": 13, "L": 24,
}


def write_module_sheet(sheet_name, title_str, rows):
    ws = wb.create_sheet(title=sheet_name)

    # Sheet title + module subtitle (rows 1-3)
    ws.cell(row=1, column=1, value=title_str).font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.cell(row=2, column=1, value=f"{len(rows)} test case(s)  ·  Module: {sheet_name}")
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=12)

    # Header row (row 4)
    header_row = 4
    for ci, h in enumerate(HEADERS, start=1):
        c = ws.cell(row=header_row, column=ci, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = HEADER_ALIGN
        c.border = HEADER_BORDER
    ws.row_dimensions[header_row].height = 28

    # Body rows
    for ri, row in enumerate(rows, start=header_row + 1):
        for ci, val in enumerate(row, start=1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = BODY_FONT
            c.border = BODY_BORDER
            c.alignment = BODY_ALIGN_CENTER if ci in (1, 2, 4, 9, 10, 11) else BODY_ALIGN_WRAP
            if (ri - header_row) % 2 == 0:
                c.fill = BAND_FILL
        # row height — estimate from longest cell
        longest = max(
            len(str(v or "")) for v in row[4:8]
        )
        ws.row_dimensions[ri].height = max(28, min(120, longest * 0.55))

    # Column widths
    for letter, w in COL_WIDTHS.items():
        ws.column_dimensions[letter].width = w

    # Freeze header + first column
    ws.freeze_panes = f"D{header_row + 1}"

    # Status dropdown (column I)
    last_row = header_row + len(rows)
    dv_status = DataValidation(
        type="list",
        formula1='"Not Run,Pass,Fail,Blocked,Skipped"',
        allow_blank=True,
    )
    dv_status.add(f"I{header_row + 1}:I{last_row}")
    ws.add_data_validation(dv_status)

    # Priority conditional formatting (col D)
    pri_range = f"D{header_row + 1}:D{last_row}"
    ws.conditional_formatting.add(
        pri_range,
        FormulaRule(formula=[f'$D{header_row + 1}="High"'],
                    fill=PatternFill("solid", fgColor="FEE2E2"),
                    font=Font(color=RED, bold=True))
    )
    ws.conditional_formatting.add(
        pri_range,
        FormulaRule(formula=[f'$D{header_row + 1}="Medium"'],
                    fill=PatternFill("solid", fgColor="FEF3C7"),
                    font=Font(color="92400E", bold=True))
    )
    ws.conditional_formatting.add(
        pri_range,
        FormulaRule(formula=[f'$D{header_row + 1}="Low"'],
                    fill=PatternFill("solid", fgColor="DCFCE7"),
                    font=Font(color="166534", bold=True))
    )

    # Status conditional formatting (col I)
    st_range = f"I{header_row + 1}:I{last_row}"
    ws.conditional_formatting.add(
        st_range,
        FormulaRule(formula=[f'$I{header_row + 1}="Pass"'],
                    fill=PatternFill("solid", fgColor="DCFCE7"),
                    font=Font(color="166534", bold=True))
    )
    ws.conditional_formatting.add(
        st_range,
        FormulaRule(formula=[f'$I{header_row + 1}="Fail"'],
                    fill=PatternFill("solid", fgColor="FEE2E2"),
                    font=Font(color=RED, bold=True))
    )
    ws.conditional_formatting.add(
        st_range,
        FormulaRule(formula=[f'$I{header_row + 1}="Blocked"'],
                    fill=PatternFill("solid", fgColor="FEF3C7"),
                    font=Font(color="92400E", bold=True))
    )
    ws.conditional_formatting.add(
        st_range,
        FormulaRule(formula=[f'$I{header_row + 1}="Skipped"'],
                    fill=PatternFill("solid", fgColor=GRAY_100),
                    font=Font(color=GRAY_500, bold=True))
    )

    return ws


# ====================================================================
# TEST DATA per module
# Columns: ID, Module, Title, Priority, Preconditions, Steps, Expected,
#          Actual, Status, Tester, Date, Notes
# ====================================================================

def case(tcid, module, title, priority, pre, steps, expected, notes=""):
    return [tcid, module, title, priority, pre, steps, expected,
            "", "Not Run", "", "", notes]


# ---------- Authentication & Session ----------
AUTH = [
    case("TC-AUTH-001", "Auth", "Successful login with valid credentials", "High",
         "A user account exists with known email and password.",
         "1. Open /login\n2. Enter the valid email\n3. Enter the valid password\n4. Click Sign in",
         "User is redirected to /dashboard; session is created; the topbar shows the user's avatar/initial."),
    case("TC-AUTH-002", "Auth", "Failed login with wrong password", "High",
         "A user account exists with known email.",
         "1. Open /login\n2. Enter valid email\n3. Enter a wrong password\n4. Click Sign in",
         "An invalid-credentials error is shown; user remains on /login; no session is created."),
    case("TC-AUTH-003", "Auth", "Failed login with non-existent email", "Medium",
         "No account uses the test email.",
         "1. Open /login\n2. Enter a non-existent email\n3. Enter any password\n4. Click Sign in",
         "Generic invalid-credentials error is shown (no information leak about which field is wrong)."),
    case("TC-AUTH-004", "Auth", "Required-field validation", "Medium",
         "—",
         "1. Open /login\n2. Click Sign in with both fields blank",
         "Browser/Laravel validation flags both fields as required; form is not submitted."),
    case("TC-AUTH-005", "Auth", "Keep me signed in extends session", "Medium",
         "Valid account exists.",
         "1. Tick 'Keep me signed in' on login\n2. Sign in\n3. Close the browser entirely\n4. Reopen and visit /dashboard",
         "User is still authenticated and lands on the dashboard without re-entering credentials."),
    case("TC-AUTH-006", "Auth", "Logout invalidates the session", "High",
         "User is signed in.",
         "1. Open the user dropdown in the topbar\n2. Click Sign out\n3. Try to visit /dashboard directly",
         "User is redirected to /login; session is destroyed; CSRF token is rotated."),
    case("TC-AUTH-007", "Auth", "Successful login is recorded in activity log", "Medium",
         "An admin account exists with access to the Activity Log.",
         "1. Sign in as a normal user\n2. Sign out\n3. Sign in as admin and open Activity Log",
         "Activity log shows action=login for the test user with correct IP and user-agent."),
    case("TC-AUTH-008", "Auth", "Failed login is recorded with attempted email", "Medium",
         "Admin can view Activity Log.",
         "1. Attempt to log in with a non-existent email\n2. Sign in as admin and open Activity Log",
         "Activity log shows action=login_failed; user_email is the attempted address; user_id is null."),
    case("TC-AUTH-009", "Auth", "Forgot password helper appears", "Low",
         "—",
         "1. On /login click 'Forgot password?'",
         "Inline helper appears with instructions to contact the IT administrator. No reset email is sent."),
    case("TC-AUTH-010", "Auth", "Theme toggle persists across login", "Low",
         "—",
         "1. On /login click the theme toggle (top-right) to switch theme\n2. Sign in",
         "Dashboard renders in the same theme as login (localStorage 'rrs.theme' preserved)."),
]

# ---------- User Management ----------
USR = [
    case("TC-USR-001", "User Mgmt", "Admin creates a new user (happy path)", "High",
         "Signed in as admin; SMTP configured.",
         "1. Open User Management → New User\n2. Fill name, unique email, password ≥ 6 chars, role=user\n3. Toggle permissions for at least one module\n4. Click Save",
         "User is listed in the index; flash 'User created' is shown; credentials email sent (or warning shown if SMTP fails)."),
    case("TC-USR-002", "User Mgmt", "Email uniqueness validation", "High",
         "An existing user has email X.",
         "1. Try to create a new user with email X\n2. Submit",
         "Form is rejected with 'email has already been taken' validation error."),
    case("TC-USR-003", "User Mgmt", "Password minimum length enforced", "Medium",
         "Signed in as admin.",
         "1. Create user with password 'abc' (3 chars)\n2. Submit",
         "Form is rejected with the minimum-length validation message."),
    case("TC-USR-004", "User Mgmt", "Edit implies view (per-module)", "High",
         "Signed in as admin editing a non-admin user.",
         "1. In the module-permissions card, tick Edit for any module\n2. Observe the View checkbox",
         "View is auto-checked, locked, and the card shows 'View + Edit'."),
    case("TC-USR-005", "User Mgmt", "Avatar upload (valid)", "High",
         "Signed in as admin; storage symlink/junction exists at public/storage.",
         "1. Open Edit user\n2. Click Upload photo and select a JPG/PNG ≤ 2MB\n3. Save",
         "File is stored under storage/app/public/avatars; avatar shows on the user index and topbar."),
    case("TC-USR-006", "User Mgmt", "Avatar upload exceeds 2 MB", "Medium",
         "—",
         "1. Open Edit user\n2. Select an image > 2 MB\n3. Save",
         "Validation error 'avatar may not be greater than 2048 KB' shown; no file saved."),
    case("TC-USR-007", "User Mgmt", "Avatar upload wrong MIME", "Medium",
         "—",
         "1. Open Edit user\n2. Select a .pdf or .gif file\n3. Save",
         "Validation error on mimes:jpg,jpeg,png,webp; no file saved."),
    case("TC-USR-008", "User Mgmt", "Admin cannot delete themselves", "High",
         "Signed in as admin A on the user index.",
         "1. Click Delete on the row representing the signed-in admin",
         "Action is rejected with the error 'You cannot delete your own account'; account remains."),
    case("TC-USR-009", "User Mgmt", "Welcome email contains credentials & login URL", "Medium",
         "SMTP works; a test mailbox is available.",
         "1. Create a new user using the test mailbox\n2. Open the inbox",
         "Welcome email arrives with the user's email, plain-text password, and a link back to /login."),
    case("TC-USR-010", "User Mgmt", "Admin role hides per-module toggles", "Medium",
         "Editing any non-admin user.",
         "1. Change Role to admin in the form",
         "Yellow notice appears explaining toggles are ignored; all module checkboxes become disabled visually."),
    case("TC-USR-011", "User Mgmt", "Non-admin cannot access User Management", "High",
         "Signed in as a user role.",
         "1. Manually navigate to /users",
         "403 (or redirect to dashboard) is returned; the user cannot reach the page."),
    case("TC-USR-012", "User Mgmt", "Delete user cascades cleanly", "Medium",
         "A non-admin user with avatar and notification reads exists.",
         "1. Delete that user from the index",
         "User row is removed; avatar file is deleted from disk; associated notification_reads cascade-delete."),
]

# ---------- PC Master ----------
PC = [
    case("TC-PC-001", "PC Master", "Create PC with valid data", "High",
         "User has edit permission on PC Master.",
         "1. Open PC Master → New\n2. Fill all required fields (computer_id, hostname, etc.)\n3. Save",
         "Record is saved; activity_log entry action=created is written; user is redirected to the index."),
    case("TC-PC-002", "PC Master", "Sensitive credential fields are encrypted at rest", "High",
         "Create or edit a PC with admin_password and username/password set; DB access available.",
         "1. Save the record\n2. Inspect the database row directly",
         "admin_password, username and password are not stored in plain text (Laravel Crypt ciphertext)."),
    case("TC-PC-003", "PC Master", "Bulk Excel import (valid file)", "High",
         "Downloaded template; edit permission.",
         "1. Fill the template with N valid rows\n2. Use Import; pick the file; submit",
         "All N rows are inserted; success flash names the count; activity_log shows action=imported."),
    case("TC-PC-004", "PC Master", "Bulk import with invalid columns is rejected", "Medium",
         "Edit permission.",
         "1. Upload an .xlsx that has missing required columns",
         "Import fails with a clear error; no partial rows are committed."),
    case("TC-PC-005", "PC Master", "Export Excel", "Medium",
         "View permission; at least one PC exists.",
         "1. Click Export on the index",
         "Browser downloads an .xlsx containing all visible records with correct columns."),
    case("TC-PC-006", "PC Master", "Download template", "Low",
         "View permission.",
         "1. Click Template on the index",
         "Browser downloads the .xlsx import template with empty rows and correct headers."),
    case("TC-PC-007", "PC Master", "Search and filter the index", "Medium",
         "At least 5 PCs with varying departments.",
         "1. Use the search box and a department filter on the index",
         "Only matching records are displayed; pagination reflects the filtered count."),
    case("TC-PC-008", "PC Master", "Update existing PC", "High",
         "Edit permission.",
         "1. Click Edit on a PC row\n2. Change a field and Save",
         "Record is updated; activity_log action=updated records the changed_fields property."),
    case("TC-PC-009", "PC Master", "Delete PC", "Medium",
         "Edit permission.",
         "1. Click Delete on a PC row and confirm",
         "Record is removed; activity_log action=deleted is written."),
    case("TC-PC-010", "PC Master", "Bulk delete", "Medium",
         "Edit permission; multiple rows checked.",
         "1. Select multiple rows on the index\n2. Click Bulk delete and confirm",
         "All selected records are removed; one or more activity_log entries are written."),
    case("TC-PC-011", "PC Master", "View-only user cannot edit/delete", "High",
         "Signed in as a user with can_view_pc_assets=true and can_edit_pc_assets=false.",
         "1. Open PC Master index\n2. Try to access /pc-assets/{id}/edit\n3. Try to POST a destroy",
         "Edit/Delete buttons are hidden; direct URL/POST returns 403."),
    case("TC-PC-012", "PC Master", "User with no access cannot view module", "High",
         "Signed in as a user with all PC permissions off.",
         "1. Try to open /pc-assets",
         "403 response from the module middleware; sidebar link is hidden."),
]

# ---------- Device Master ----------
DEV = [
    case("TC-DEV-001", "Device", "Create device with valid data", "High",
         "Edit permission on Devices.",
         "1. Open Device Master → New\n2. Fill all required fields including serial_number\n3. Save",
         "Record is saved; activity_log action=created."),
    case("TC-DEV-002", "Device", "Serial number captured on create", "Medium",
         "—",
         "1. Save a device with a unique serial_number\n2. Open Edit",
         "serial_number persists and is shown on the Edit form."),
    case("TC-DEV-003", "Device", "Bulk Excel import (valid)", "High",
         "Downloaded device template.",
         "1. Fill the template with N rows\n2. Import",
         "All N rows imported; success flash shows count; activity log records the import."),
    case("TC-DEV-004", "Device", "Export Excel", "Medium",
         "At least one device exists.",
         "1. Click Export",
         "Browser downloads an .xlsx with all device rows and columns."),
    case("TC-DEV-005", "Device", "Update existing device", "Medium",
         "Edit permission.",
         "1. Edit a device → change a field → Save",
         "Record updated; activity_log action=updated."),
    case("TC-DEV-006", "Device", "Delete device", "Medium",
         "Edit permission.",
         "1. Click Delete on a device row",
         "Record removed; activity_log action=deleted."),
    case("TC-DEV-007", "Device", "View-only user cannot edit/delete", "High",
         "User with view but not edit on Devices.",
         "1. Open Devices index\n2. Try to access an edit URL",
         "Buttons hidden; direct URL returns 403."),
]

# ---------- Subscriptions ----------
SUB = [
    case("TC-SUB-001", "Subscriptions", "Create subscription (happy path)", "High",
         "Edit permission on Subscriptions.",
         "1. New Subscription → fill service_type, project_name, subscription_name, vendor_name, expire_date, currency, costs\n2. Save",
         "Record is saved; reminder_date is auto-set to expire_date - reminder_days_before (default 30)."),
    case("TC-SUB-002", "Subscriptions", "reminder_date auto-recomputes on edit", "High",
         "Subscription exists.",
         "1. Edit the subscription\n2. Change expire_date\n3. Save",
         "reminder_date is recalculated and shown on the row/edit form."),
    case("TC-SUB-003", "Subscriptions", "Renew action updates status", "High",
         "Subscription with renewal_status='Pending'.",
         "1. Click Renew on the row\n2. Confirm",
         "renewal_status flips to 'Renewed'; activity_log action=renewed is written."),
    case("TC-SUB-004", "Subscriptions", "Currency enum validation", "Medium",
         "—",
         "1. Try to POST a subscription with currency='EUR' via direct form/manipulation",
         "Server rejects with validation error; allowed currencies are MMK/JPY/USD only."),
    case("TC-SUB-005", "Subscriptions", "Bulk Excel import", "High",
         "Subscriptions template downloaded.",
         "1. Fill template with N rows\n2. Import",
         "All rows imported; reminder_date computed for each; activity_log action=imported."),
    case("TC-SUB-006", "Subscriptions", "Filter by renewal status", "Medium",
         "Mix of Pending/Renewed/Expired records.",
         "1. Use the renewal-status filter on the index",
         "Only matching records are shown; counts on KPI strip update."),
    case("TC-SUB-007", "Subscriptions", "Delete subscription", "Medium",
         "Edit permission.",
         "1. Click Delete on a subscription row",
         "Record removed; activity_log action=deleted."),
    case("TC-SUB-008", "Subscriptions", "Subscription past expire becomes Expired after scheduler", "High",
         "A subscription has expire_date < today and renewal_status != 'Renewed'.",
         "1. Run: php artisan app:check-expirations",
         "Subscription's renewal_status is set to 'Expired'; console output reports the count."),
]

# ---------- Licenses & Contracts ----------
LIC = [
    case("TC-LIC-001", "Licenses", "Create license (happy path)", "High",
         "Edit permission on Licenses & Contracts.",
         "1. New License → fill software_name, vendor_name, license_info, expire_date, renewal_type, costs\n2. Save",
         "Record is saved; activity_log action=created."),
    case("TC-LIC-002", "Licenses", "Status enum validation", "Medium",
         "—",
         "1. Try to set status to 'Foo' via direct form/manipulation",
         "Server rejects: allowed values are Active, Pending, Expired, Terminated."),
    case("TC-LIC-003", "Licenses", "Bulk Excel import", "High",
         "Licenses template downloaded.",
         "1. Fill template with N rows\n2. Import",
         "All rows imported; activity_log action=imported."),
    case("TC-LIC-004", "Licenses", "Export Excel", "Medium",
         "At least one license exists.",
         "1. Click Export",
         ".xlsx download contains all license records with correct columns."),
    case("TC-LIC-005", "Licenses", "Update existing license", "Medium",
         "Edit permission.",
         "1. Edit a license → change expire_date → Save",
         "Record updated; activity_log action=updated with changed_fields property."),
    case("TC-LIC-006", "Licenses", "Delete license", "Medium",
         "Edit permission.",
         "1. Click Delete on a license row",
         "Record removed; activity_log action=deleted."),
    case("TC-LIC-007", "Licenses", "Licenses do NOT auto-flip to Pending on digest", "Medium",
         "License within a configured day-mark; module enabled.",
         "1. Run: php artisan app:check-expirations",
         "License status is unchanged after the digest send (unlike subscriptions which flip to Pending)."),
]

# ---------- Notifications (page + bell) ----------
NOTIF = [
    case("TC-NOTIF-001", "Notifications", "Bell badge shows correct unread count", "High",
         "User is signed in; at least 3 items inside the reminder window for an enabled module; user has read 0.",
         "1. Look at the topbar bell badge",
         "Badge shows the same number as the 'All' KPI on the notifications page for this user."),
    case("TC-NOTIF-002", "Notifications", "Module filter — Subscriptions only", "Medium",
         "Items exist for both modules.",
         "1. Open Notifications → click the Subscriptions module KPI",
         "Only subscription items are listed; URL has ?module=subscriptions."),
    case("TC-NOTIF-003", "Notifications", "Module filter — Licenses only", "Medium",
         "Items exist for both modules.",
         "1. Click the Licenses & Contracts KPI",
         "Only license items are listed; URL has ?module=licenses_contracts."),
    case("TC-NOTIF-004", "Notifications", "Status filter — Unread", "Medium",
         "Mix of read and unread items.",
         "1. Click the Unread status chip",
         "Only items with no matching notification_reads row (or with mismatched signature) appear."),
    case("TC-NOTIF-005", "Notifications", "Mark single item as read", "High",
         "At least one unread item.",
         "1. Click the check button on an unread row",
         "Row's New badge disappears; reload — item stays read for this user."),
    case("TC-NOTIF-006", "Notifications", "Mark all as read", "High",
         "Multiple unread items.",
         "1. Click 'Mark all as read'",
         "All currently-listed unread items become read for this user; flash message shows the count."),
    case("TC-NOTIF-007", "Notifications", "Read state is per-user", "High",
         "Two users A and B both can see the same item.",
         "1. User A marks the item read\n2. User B opens Notifications",
         "User B still sees the item as unread."),
    case("TC-NOTIF-008", "Notifications", "Read state invalidates when urgency changes (signature)", "High",
         "Item is in 'upcoming' bucket; user A marks it read.",
         "1. Move the expire_date so the item now falls in 'soon' (<=7 days)\n2. User A reopens Notifications",
         "Item re-surfaces as unread for user A because the live signature no longer matches the stored one."),
    case("TC-NOTIF-009", "Notifications", "Disabled module hides notifications", "Medium",
         "Subscriptions notifications are enabled; items exist.",
         "1. Disable Subscriptions in Notification Settings\n2. Reload Notifications page and bell",
         "Subscription items disappear from the page; bell count drops accordingly."),
    case("TC-NOTIF-010", "Notifications", "Live badge updates after edit", "Medium",
         "Item is in 'upcoming'.",
         "1. Edit the underlying subscription, change expire_date to today\n2. Reload the page",
         "Bell badge increments by the new overdue/soon item (without a scheduler run)."),
]

# ---------- Mail Settings ----------
MAIL = [
    case("TC-MAIL-001", "Mail Settings", "Save SMTP connection (DB mode)", "High",
         "Signed in as admin.",
         "1. Open Mail Settings\n2. Toggle 'Database SMTP settings' on\n3. Fill host/port/encryption/auth_mode/username/password/from\n4. Save",
         "Settings persist; password is encrypted in the DB; activity log records the change (without storing the password value)."),
    case("TC-MAIL-002", "Mail Settings", "Test email succeeds", "High",
         "DB SMTP enabled and valid OR .env mail configured.",
         "1. In 'Send Test Email' enter a real address\n2. Click Send Test",
         "Inbox receives '[ITAMS] Test Email' with a plain-text body; success flash shown."),
    case("TC-MAIL-003", "Mail Settings", "Test email surfaces SMTP error", "Medium",
         "DB SMTP enabled with wrong host or rejected recipient.",
         "1. Send Test to a known-bad address",
         "Error flash shows the SMTP exception message; no email is delivered."),
    case("TC-MAIL-004", "Mail Settings", "Password redaction on edit", "High",
         "Saved DB SMTP with a password.",
         "1. Reopen Mail Settings",
         "Password field placeholder shows masked '••••••••  (leave blank to keep)'; clearing+saving with blank keeps the existing password."),
    case("TC-MAIL-005", "Mail Settings", "Invalid reminder recipient rejected", "Medium",
         "—",
         "1. Enter 'notanemail' in reminder_recipients\n2. Save",
         "Form is rejected with 'Invalid email address: notanemail'."),
    case("TC-MAIL-006", "Mail Settings", "Disabling DB SMTP falls back to .env", "Medium",
         "DB SMTP enabled.",
         "1. Untick the DB SMTP toggle\n2. Save\n3. Send a test email",
         "Test email uses .env credentials; status badge in the header shows 'Using .env'."),
    case("TC-MAIL-007", "Mail Settings", "Non-admin cannot reach Mail Settings", "High",
         "Signed in as a non-admin.",
         "1. Manually navigate to /mail-settings",
         "Middleware returns 403; the page is unreachable."),
]

# ---------- Notification Settings ----------
NSET = [
    case("TC-NSET-001", "Notif Settings", "Enable subscription reminders", "High",
         "Signed in as admin; Notification Settings page open.",
         "1. Toggle Subscriptions enabled = on\n2. Save",
         "Settings persist; bell starts counting subscription items in the configured window."),
    case("TC-NSET-002", "Notif Settings", "Choose day-marks", "High",
         "Notification Settings page open.",
         "1. Tick 10, 20 and 30 day-marks for Subscriptions\n2. Save",
         "Stored days_before_set JSON contains [30, 20, 10]; windowDays() returns 30."),
    case("TC-NSET-003", "Notif Settings", "Day-mark out of allowed set is rejected", "Medium",
         "—",
         "1. Attempt to POST days_before_set=[5] via the form (manipulated)",
         "Only values from {10, 20, 30} are persisted; others silently dropped or validation error returned."),
    case("TC-NSET-004", "Notif Settings", "Empty recipients fall back to admins on send", "Medium",
         "At least one admin user has an email; recipients field empty.",
         "1. Add an item matching today + chosen day-mark\n2. Run app:check-expirations",
         "Digest is sent to all admin emails (per the fallback resolver)."),
    case("TC-NSET-005", "Notif Settings", "Non-admin cannot reach Notification Settings", "High",
         "Signed in as a non-admin.",
         "1. Navigate to /notification-settings",
         "Middleware returns 403."),
]

# ---------- Activity Log ----------
LOG = [
    case("TC-LOG-001", "Activity Log", "Login event captured", "High",
         "Signed in as admin.",
         "1. Log out and back in\n2. Open Activity Log",
         "Most recent entry is action=login with matching user and IP."),
    case("TC-LOG-002", "Activity Log", "Failed login captured", "Medium",
         "—",
         "1. Try to log in with wrong password\n2. Open Activity Log as admin",
         "An action=login_failed entry exists; user_email is the attempted address; user_id is null."),
    case("TC-LOG-003", "Activity Log", "Create captured with subject", "Medium",
         "Create any PC asset.",
         "1. Open Activity Log",
         "Entry action=created exists; subject_type=App\\Models\\PcAsset; subject_id matches the new record."),
    case("TC-LOG-004", "Activity Log", "Update captures changed_fields", "Medium",
         "Edit any record and change at least 2 fields.",
         "1. Open Activity Log\n2. Expand the matching entry",
         "Entry action=updated; properties.changed_fields is the list of changed columns (password excluded if applicable)."),
    case("TC-LOG-005", "Activity Log", "Delete captured", "Low",
         "Delete any non-self record.",
         "1. Open Activity Log",
         "Entry action=deleted with matching subject."),
    case("TC-LOG-006", "Activity Log", "Non-admin cannot reach Activity Log", "High",
         "Signed in as non-admin.",
         "1. Navigate to /activity-logs",
         "Middleware returns 403."),
    case("TC-LOG-007", "Activity Log", "IP & user-agent captured", "Low",
         "Recent entry exists.",
         "1. Inspect any entry's detail in the table or DB",
         "ip_address and user_agent fields are populated (user_agent truncated to 500 chars)."),
]

# ---------- Reminder pipeline (CheckExpirations command) ----------
CHK = [
    case("TC-CHK-001", "Reminder Pipeline", "Past-due subscriptions marked Expired", "High",
         "A subscription has expire_date < today, status=Active, renewal_status!=Renewed.",
         "1. Run: php artisan app:check-expirations",
         "Subscription's renewal_status flips to Expired; console reports 'Marked N subscription(s) as Expired'."),
    case("TC-CHK-002", "Reminder Pipeline", "Past-due licenses are NOT auto-Expired", "Medium",
         "A license has expire_date < today, status=Active.",
         "1. Run: php artisan app:check-expirations",
         "License status is unchanged; no console line marks licenses as Expired (intentional asymmetry vs subscriptions)."),
    case("TC-CHK-003", "Reminder Pipeline", "Digest sent at exact day-mark", "High",
         "Module enabled; day-marks include N; a record's expire_date equals today + N.",
         "1. Run app:check-expirations\n2. Check the recipient mailbox",
         "Exactly ONE digest arrives for that (module × N) bucket containing the matching record(s)."),
    case("TC-CHK-004", "Reminder Pipeline", "Subscription flipped to Pending on digest send", "Medium",
         "Subscription matches a day-mark.",
         "1. Run the command",
         "Subscription's renewal_status becomes Pending after the digest send."),
    case("TC-CHK-005", "Reminder Pipeline", "License NOT flipped on digest send", "Medium",
         "License matches a day-mark.",
         "1. Run the command",
         "License status is unchanged by the digest send."),
    case("TC-CHK-006", "Reminder Pipeline", "Disabled module is skipped", "High",
         "Module's notification setting enabled=false.",
         "1. Run the command",
         "Console prints '[<module>] notifications disabled — skipped'; no email sent."),
    case("TC-CHK-007", "Reminder Pipeline", "No day-marks selected is skipped", "Medium",
         "Module enabled but days_before_set is empty.",
         "1. Run the command",
         "Console prints '[<module>] no day-marks selected — skipped'; no email sent."),
    case("TC-CHK-008", "Reminder Pipeline", "Recipients fallback to admins when empty", "High",
         "Module enabled; recipients empty; at least one admin user has email.",
         "1. Run the command for a day-mark with matching records",
         "Digest is sent to every admin user's email address."),
    case("TC-CHK-009", "Reminder Pipeline", "Per-recipient SMTP failure does not abort the run", "High",
         "Recipient list contains one valid and one rejected address.",
         "1. Run the command",
         "Console prints a failure line for the bad batch; the loop continues; valid digests still go out."),
    case("TC-CHK-010", "Reminder Pipeline", "Re-running same day sends duplicate digests", "Medium",
         "First run produced a digest.",
         "1. Immediately re-run app:check-expirations",
         "Same digest is sent again — there is no per-day dedupe; documented behaviour."),
]


SHEETS = [
    ("Authentication", "Authentication & Session", AUTH),
    ("User Management", "User Management",  USR),
    ("PC Master",       "PC Master Module",  PC),
    ("Device Master",   "Device Master Module", DEV),
    ("Subscriptions",   "Subscriptions Module", SUB),
    ("Licenses",        "Licenses & Contracts Module", LIC),
    ("Notifications",   "Notifications (page & bell)", NOTIF),
    ("Mail Settings",   "Mail Settings",     MAIL),
    ("Notif Settings",  "Notification Settings", NSET),
    ("Activity Log",    "Activity Log",      LOG),
    ("Reminder Job",    "Reminder Pipeline · app:check-expirations", CHK),
]

# ====================================================================
# Summary sheet (first)
# ====================================================================
total_count = sum(len(rows) for _, _, rows in SHEETS)
priority_counts = {"High": 0, "Medium": 0, "Low": 0}
for _, _, rows in SHEETS:
    for r in rows:
        priority_counts[r[3]] = priority_counts.get(r[3], 0) + 1

ws = wb.active
ws.title = "Overview"

# Cover block
ws.cell(row=2, column=2, value="ITAMS").font = Font(name="Calibri", size=36, bold=True, color=NAVY)
ws.cell(row=3, column=2, value="System Test Cases").font = Font(name="Calibri", size=14, color=ACCENT, bold=True)
ws.cell(row=5, column=2, value="Version 1.0").font = META_FONT
ws.cell(row=6, column=2, value=f"Issued · {date.today().strftime('%d %B %Y')}").font = META_FONT
ws.cell(row=7, column=2, value="Status · Released").font = META_FONT

# Stats table
ws.cell(row=10, column=2, value="AT A GLANCE").font = LABEL_FONT
ws.cell(row=10, column=2).alignment = Alignment(horizontal="left")

stats = [
    ("Total test cases",     total_count),
    ("Modules covered",      len(SHEETS)),
    ("High priority",        priority_counts.get("High", 0)),
    ("Medium priority",      priority_counts.get("Medium", 0)),
    ("Low priority",         priority_counts.get("Low", 0)),
]
for i, (label, val) in enumerate(stats, start=11):
    c1 = ws.cell(row=i, column=2, value=label)
    c2 = ws.cell(row=i, column=3, value=val)
    c1.font = Font(name="Calibri", size=11, color=GRAY_500)
    c2.font = Font(name="Calibri", size=12, bold=True, color=INK)
    c1.alignment = Alignment(horizontal="left", vertical="center")
    c2.alignment = Alignment(horizontal="left", vertical="center")

# Modules table
ws.cell(row=18, column=2, value="MODULES").font = LABEL_FONT
mod_headers = ["#", "Module", "Sheet", "Test cases", "High", "Medium", "Low"]
for ci, h in enumerate(mod_headers, start=2):
    c = ws.cell(row=19, column=ci, value=h)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = HEADER_ALIGN
    c.border = HEADER_BORDER
ws.row_dimensions[19].height = 24

for i, (sheet_name, title, rows) in enumerate(SHEETS, start=1):
    r = 19 + i
    pc_h = sum(1 for x in rows if x[3] == "High")
    pc_m = sum(1 for x in rows if x[3] == "Medium")
    pc_l = sum(1 for x in rows if x[3] == "Low")
    vals = [i, title, sheet_name, len(rows), pc_h, pc_m, pc_l]
    for ci, v in enumerate(vals, start=2):
        c = ws.cell(row=r, column=ci, value=v)
        c.font = BODY_FONT
        c.border = BODY_BORDER
        c.alignment = BODY_ALIGN_CENTER if ci != 3 else Alignment(horizontal="left", vertical="center")
        if i % 2 == 0:
            c.fill = BAND_FILL

# Legend block
legend_start = 19 + len(SHEETS) + 3
ws.cell(row=legend_start, column=2, value="STATUS LEGEND").font = LABEL_FONT
legend = [
    ("Not Run", "Test case has not yet been executed.",  GRAY_100,  GRAY_500),
    ("Pass",    "Test case passed as expected.",         "DCFCE7", "166534"),
    ("Fail",    "Test case failed; logged for fix.",     "FEE2E2", RED),
    ("Blocked", "Cannot execute due to a dependency.",   "FEF3C7", "92400E"),
    ("Skipped", "Intentionally not executed this cycle.", GRAY_100, GRAY_500),
]
for i, (s, desc, bg, fg) in enumerate(legend, start=legend_start + 1):
    c1 = ws.cell(row=i, column=2, value=s)
    c1.fill = PatternFill("solid", fgColor=bg)
    c1.font = Font(name="Calibri", size=10, bold=True, color=fg)
    c1.alignment = BODY_ALIGN_CENTER
    c1.border = BODY_BORDER
    c2 = ws.cell(row=i, column=3, value=desc)
    c2.font = BODY_FONT
    c2.alignment = Alignment(horizontal="left", vertical="center")

# Sign-off
sign_y = legend_start + len(legend) + 3
ws.cell(row=sign_y, column=2, value="SIGN-OFF").font = LABEL_FONT
for i, label in enumerate(["Prepared by", "Reviewed by", "Approved by"], start=sign_y + 1):
    c = ws.cell(row=i, column=2, value=label)
    c.font = Font(name="Calibri", size=10, color=GRAY_500)
    ws.cell(row=i, column=3, value="").border = BODY_BORDER
    ws.cell(row=i, column=4, value="Date").font = Font(name="Calibri", size=10, color=GRAY_500)
    ws.cell(row=i, column=5, value="").border = BODY_BORDER

# Column widths
for col, w in [(1, 4), (2, 26), (3, 38), (4, 14), (5, 12), (6, 12), (7, 12)]:
    ws.column_dimensions[get_column_letter(col)].width = w

# Hide gridlines, freeze top
ws.sheet_view.showGridLines = False


# ====================================================================
# Build per-module sheets
# ====================================================================
for sheet_name, title, rows in SHEETS:
    write_module_sheet(sheet_name, title, rows)
    # also hide gridlines for cleaner presentation
    wb[sheet_name].sheet_view.showGridLines = False


# ====================================================================
# Save
# ====================================================================
out = r"D:\xampp\htdocs\itams\docs\Test_Cases.xlsx"
wb.save(out)
print(f"Saved: {out}")
print(f"Sheets: {len(wb.sheetnames)} -> {wb.sheetnames}")
print(f"Total test cases: {total_count}")
